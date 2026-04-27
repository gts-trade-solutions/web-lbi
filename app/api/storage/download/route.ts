import { GetObjectCommand } from "@aws-sdk/client-s3";
import { requireAuth } from "../../../../lib/auth";
import { getBucketName, safeObjectKey, s3Client } from "../../../../lib/s3";

export const runtime = "nodejs";

export async function GET(request: Request) {
  try {
    requireAuth(request);
    const { searchParams } = new URL(request.url);
    const path = safeObjectKey(searchParams.get("path") || "");
    if (!path) return Response.json({ error: "path is required" }, { status: 400 });

    const object = await s3Client.send(
      new GetObjectCommand({
        Bucket: getBucketName(),
        Key: path,
      })
    );

    if (!object.Body) return Response.json({ error: "Not found" }, { status: 404 });

    const arrayBuffer = await object.Body.transformToByteArray();
    return new Response(arrayBuffer, {
      status: 200,
      headers: {
        "Content-Type": object.ContentType || "application/octet-stream",
      },
    });
  } catch (e: unknown) {
    const err = e as { message?: string };
    if (err?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    return Response.json({ error: "Failed to download file" }, { status: 500 });
  }
}
