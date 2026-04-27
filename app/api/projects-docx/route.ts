/* eslint-disable @typescript-eslint/no-explicit-any */
import { AlignmentType, Document, Packer, Paragraph, TextRun } from "docx";
import pool from "../../../lib/db";
import { requireAuth } from "../../../lib/auth";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export async function GET(request: Request) {
  try {
    requireAuth(request);
    return new Response("OK: /api/projects-docx route is working", {
      status: 200,
      headers: { "Content-Type": "text/plain" },
    });
  } catch {
    return new Response("Unauthorized", { status: 401 });
  }
}

export async function POST(request: Request) {
  try {
    requireAuth(request);
    const body = await request.json().catch(() => ({}));
    const projectIds = Array.isArray(body?.projectIds) ? body.projectIds : [];
    if (!projectIds.length) return new Response("projectIds[] required", { status: 400 });

    const placeholders = projectIds.map(() => "?").join(", ");
    const [projectRows] = await pool.query(
      `SELECT id, name, title, project_name, description, created_at FROM projects WHERE id IN (${placeholders})`,
      projectIds
    );
    const projects = Array.isArray(projectRows) ? (projectRows as any[]) : [];

    const [reportRows] = await pool.query(
      `SELECT id, project_id, category, description, created_at FROM reports WHERE project_id IN (${placeholders}) ORDER BY created_at ASC`,
      projectIds
    );
    const reports = Array.isArray(reportRows) ? (reportRows as any[]) : [];

    const lines: Paragraph[] = [];
    lines.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Projects Summary", bold: true, size: 36 })],
      }),
      new Paragraph({ text: "" })
    );

    for (const p of projects) {
      const name = p.name || p.title || p.project_name || "Untitled Project";
      const projectReports = reports.filter((r) => r.project_id === p.id);
      lines.push(
        new Paragraph({
          children: [new TextRun({ text: `${name} (${p.id})`, bold: true, size: 28 })],
        }),
        new Paragraph({
          children: [new TextRun({ text: `Reports: ${projectReports.length}` })],
        })
      );
      for (const r of projectReports) {
        lines.push(
          new Paragraph({
            children: [new TextRun({ text: `- ${r.category || "Report"} (${r.id})` })],
          })
        );
      }
      lines.push(new Paragraph({ text: "" }));
    }

    const doc = new Document({
      sections: [{ children: lines }],
    });
    const buf = await Packer.toBuffer(doc);

    return new Response(buf, {
      status: 200,
      headers: {
        "Content-Type": DOCX_MIME,
        "Content-Disposition": `attachment; filename="projects-route-report.docx"`,
      },
    });
  } catch (e: any) {
    if (e?.message === "Unauthorized") return new Response("Unauthorized", { status: 401 });
    return new Response("Failed to generate projects docx", { status: 500 });
  }
}
