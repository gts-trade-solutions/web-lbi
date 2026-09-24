import { randomUUID } from "crypto";
import { GetObjectCommand, PutObjectCommand, DeleteObjectCommand, S3Client } from "@aws-sdk/client-s3";
import { getSignedUrl } from "@aws-sdk/s3-request-presigner";

const region = process.env.AWS_S3_REGION;
const accessKeyId = process.env.AWS_S3_ACCESS_KEY_ID;
const secretAccessKey = process.env.AWS_S3_SECRET_ACCESS_KEY;
const bucketName = process.env.AWS_S3_BUCKET_NAME;
const publicBucketUrl = (process.env.NEXT_PUBLIC_S3_BUCKET_URL || "").replace(/\/+$/, "");

if (!region || !accessKeyId || !secretAccessKey || !bucketName) {
  // Runtime checks are handled in route handlers as well.
}

export const s3Client = new S3Client({
  // Fallback only so this module can load where S3 isn't configured (local
  // dev): the SDK throws "Region is missing" at construction otherwise, which
  // would take down every route importing this file, not just S3 ones.
  region: region || "us-east-1",
  credentials: {
    accessKeyId: accessKeyId || "",
    secretAccessKey: secretAccessKey || "",
  },
});

export const S3_BUCKET_NAME = bucketName || "";

export function getBucketName() {
  if (!bucketName) throw new Error("AWS_S3_BUCKET_NAME is not configured");
  return bucketName;
}

export function safeObjectKey(rawPath?: string, originalFileName?: string) {
  if (rawPath && rawPath.trim()) {
    return rawPath
      .trim()
      .replace(/^\/+/, "")
      .replace(/\\/g, "/")
      .replace(/\.\./g, "")
      .replace(/\/{2,}/g, "/");
  }

  const base = (originalFileName || "file")
    .replace(/[^\w.\-]+/g, "_")
    .replace(/_{2,}/g, "_")
    .slice(0, 140);
  return `uploads/${new Date().toISOString().slice(0, 10)}/${randomUUID()}_${base}`;
}

export async function uploadBufferToS3(params: {
  key: string;
  body: Buffer | Uint8Array;
  contentType?: string;
}) {
  const bucket = getBucketName();
  await s3Client.send(
    new PutObjectCommand({
      Bucket: bucket,
      Key: params.key,
      Body: params.body,
      ContentType: params.contentType,
    })
  );

  return {
    bucket,
    key: params.key,
    url: getPublicS3Url(params.key),
  };
}

export function getPublicS3Url(key: string) {
  const normalizedKey = key.replace(/^\/+/, "");
  if (publicBucketUrl) return `${publicBucketUrl}/${normalizedKey}`;
  const bucket = getBucketName();
  if (!region) throw new Error("AWS_S3_REGION is not configured");
  return `https://${bucket}.s3.${region}.amazonaws.com/${normalizedKey}`;
}

/**
 * Extracts the S3 object key from a URL pointing at OUR bucket
 * (NEXT_PUBLIC_S3_BUCKET_URL host, or the virtual-host / path-style AWS forms).
 * Returns null for anything else, so callers can refuse foreign URLs.
 */
export function s3KeyFromUrl(url: string): string | null {
  if (!url) return null;
  try {
    const u = new URL(url);
    if (publicBucketUrl) {
      const baseUrl = new URL(publicBucketUrl);
      if (u.host === baseUrl.host) {
        const basePath = baseUrl.pathname.replace(/\/+$/, "");
        let p = u.pathname;
        if (basePath && p.startsWith(basePath)) p = p.slice(basePath.length);
        return decodeURIComponent(p.replace(/^\/+/, "")) || null;
      }
    }
    if (
      bucketName &&
      (u.host === `${bucketName}.s3.${region}.amazonaws.com` ||
        u.host === `${bucketName}.s3.amazonaws.com` ||
        u.host === `s3.${region}.amazonaws.com` ||
        u.host === `s3.amazonaws.com`)
    ) {
      let p = u.pathname.replace(/^\/+/, "");
      if (u.host.startsWith("s3.")) {
        // s3.region.amazonaws.com/bucket/key form
        const prefix = `${bucketName}/`;
        if (!p.startsWith(prefix)) return null;
        p = p.slice(prefix.length);
      }
      return decodeURIComponent(p) || null;
    }
  } catch {
    return null;
  }
  return null;
}

// ---- Private-bucket photo URLs -------------------------------------------
// The DB keeps the permanent public URL of every photo. The bucket is private,
// so whatever we hand to a client is swapped for a presigned GET URL, and any
// presigned URL a client sends back is reduced to the permanent URL again
// before it is stored.
//
// Signing is pinned to the start of a 6-hour window, so the same photo gets
// the SAME URL for 6 hours (browser + phone image caches keep working) and
// every URL stays valid for at least 12 hours after it is handed out.
const SIGN_WINDOW_SEC = 6 * 60 * 60;
const SIGN_MIN_VALID_SEC = 12 * 60 * 60;
const signedCache = new Map<string, string>();
let signedCacheWindow = 0;

export async function signS3Url<T>(url: T): Promise<T | string> {
  if (typeof url !== "string" || !url) return url;
  const key = s3KeyFromUrl(url);
  if (!key) return url; // relative /uploads/..., other hosts: leave alone
  const windowStart = Math.floor(Date.now() / 1000 / SIGN_WINDOW_SEC) * SIGN_WINDOW_SEC;
  if (windowStart !== signedCacheWindow) {
    signedCache.clear();
    signedCacheWindow = windowStart;
  }
  const hit = signedCache.get(key);
  if (hit) return hit;
  try {
    const signed = await getSignedUrl(
      s3Client,
      new GetObjectCommand({ Bucket: getBucketName(), Key: key }),
      { expiresIn: SIGN_WINDOW_SEC + SIGN_MIN_VALID_SEC, signingDate: new Date(windowStart * 1000) }
    );
    if (signedCache.size > 20000) signedCache.clear();
    signedCache.set(key, signed);
    return signed;
  } catch (err) {
    console.error("[s3] presign failed, returning stored URL:", key, err);
    return url;
  }
}

/** Presigned URL (…?X-Amz-…) for our bucket → the permanent URL we store. */
export function canonicalS3Url<T>(url: T): T | string {
  if (typeof url !== "string" || !url.includes("X-Amz-")) return url;
  if (!s3KeyFromUrl(url)) return url;
  return url.split("?")[0];
}

// Columns that hold photo/file URLs, per table.
export const URL_COLUMNS: Record<string, string[]> = {
  report_photos: ["url", "anno_base_url"],
  project_route_pages: ["map_file_url"],
  project_route_page_images: ["file_url"],
  project_ga_drawings: ["image_url"],
  profiles: ["avatar_url"],
};

/** Sign the URL columns of rows read from `table` (mutates + returns rows). */
export async function signRowUrls<R>(table: string, rows: R): Promise<R> {
  const cols = URL_COLUMNS[table];
  if (!cols) return rows;
  const list = (Array.isArray(rows) ? rows : [rows]) as Array<Record<string, unknown> | null>;
  await Promise.all(
    list.map(async (row) => {
      if (!row || typeof row !== "object") return;
      for (const c of cols) {
        if (typeof row[c] === "string") row[c] = await signS3Url(row[c]);
      }
    })
  );
  return rows;
}

/** Canonicalize the URL columns of rows about to be written to `table`. */
export function canonicalRowUrls<R>(table: string, rows: R): R {
  const cols = URL_COLUMNS[table];
  if (!cols) return rows;
  const list = (Array.isArray(rows) ? rows : [rows]) as Array<Record<string, unknown> | null>;
  for (const row of list) {
    if (!row || typeof row !== "object") continue;
    for (const c of cols) {
      if (typeof row[c] === "string") row[c] = canonicalS3Url(row[c]);
    }
  }
  return rows;
}

export async function deleteS3Object(key: string) {
  const bucket = getBucketName();
  await s3Client.send(
    new DeleteObjectCommand({
      Bucket: bucket,
      Key: key,
    })
  );
}

export async function getReadSignedUrl(key: string, expiresInSeconds = 600) {
  const bucket = getBucketName();
  const cmd = new GetObjectCommand({
    Bucket: bucket,
    Key: key,
  });
  return getSignedUrl(s3Client, cmd, { expiresIn: expiresInSeconds });
}
