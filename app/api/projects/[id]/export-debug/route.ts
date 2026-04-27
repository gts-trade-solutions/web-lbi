import { NextResponse } from "next/server";
import { requireAuth } from "../../../../../lib/auth";
import { buildExportDebug } from "../../../../../lib/reenaTemplateExport";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type Ctx = { params: { id: string } };

export async function GET(request: Request, context: Ctx) {
  try {
    requireAuth(request);
    const projectId = String(context.params?.id || "").trim();
    if (!projectId) {
      return NextResponse.json({ error: "Project id is required" }, { status: 400 });
    }
    const report = await buildExportDebug(projectId);
    return NextResponse.json(report);
  } catch (error) {
    if ((error as { message?: string })?.message === "Unauthorized") {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }
    console.error("[api/projects/[id]/export-debug] error:", error);
    return NextResponse.json({ error: "Failed to build debug report" }, { status: 500 });
  }
}
