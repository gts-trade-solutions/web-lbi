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
  region,
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
