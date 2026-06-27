import { requireAuth } from "../../../../lib/auth";
import { getBucketName } from "../../../../lib/s3";

export const runtime = "nodejs";

export async function GET(request: Request) {
  try {
    requireAuth(request);
    const bucket = getBucketName();
    return Response.json({
      buckets: [{ name: bucket }],
    });
  } catch (e: unknown) {
    const err = e as { message?: string };
    if (err?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    return Response.json({ error: "Failed to list buckets" }, { status: 500 });
  }
}
