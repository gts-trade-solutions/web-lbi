/* eslint-disable @typescript-eslint/no-explicit-any */
import pool from "../../../../../lib/db";
import { requireAuth } from "../../../../../lib/auth";
import {
  AlignmentType,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";

export const runtime = "nodejs";

function headerCell(text: string) {
  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, color: "FFFFFF", fill: "2F5E8F" },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text, color: "FFFFFF", bold: true })],
      }),
    ],
  });
}

function normalCell(text: string) {
  return new TableCell({
    children: [new Paragraph({ children: [new TextRun(text || "")] })],
  });
}

async function fetchImageBytes(url: string): Promise<Uint8Array | null> {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const ab = await res.arrayBuffer();
    return new Uint8Array(ab);
  } catch {
    return null;
  }
}

export async function GET(request: Request, { params }: { params: { id: string } }) {
  try {
    requireAuth(request);

    const reportId = params.id;

    const [reportRows] = await pool.query("SELECT * FROM reports WHERE id = ? LIMIT 1", [reportId]);
    const report = Array.isArray(reportRows) ? (reportRows[0] as any) : null;
    if (!report) return Response.json({ error: "Report not found" }, { status: 404 });

    let projectName = "";
    if (report.project_id) {
      const [projRows] = await pool.query("SELECT * FROM projects WHERE id = ? LIMIT 1", [report.project_id]);
      const proj = Array.isArray(projRows) ? (projRows[0] as any) : null;
      projectName = proj?.name || proj?.title || proj?.project_name || "";
    }

    const [photoRows] = await pool.query(
      "SELECT * FROM report_photos WHERE report_id = ? ORDER BY created_at ASC",
      [reportId]
    );
    const photos = Array.isArray(photoRows) ? (photoRows as any[]) : [];

    const [pointRows] = await pool.query(
      "SELECT * FROM report_path_points WHERE report_id = ? ORDER BY seq ASC",
      [reportId]
    );
    const points = Array.isArray(pointRows) ? (pointRows as any[]) : [];

    const photoUrls: string[] = photos
      .map((p: any) => p.url || p.public_url)
      .filter((u: any) => typeof u === "string" && /^https?:\/\//i.test(u));

    const rows: TableRow[] = [
      new TableRow({
        children: [
          headerCell("GPS NO"),
          headerCell("KMS"),
          headerCell("NE COORDINATE"),
          headerCell("DETAILS"),
          headerCell("LOCATION"),
          headerCell("PHOTO"),
          headerCell("VEHICLE MOVEMENT"),
        ],
      }),
    ];

    for (const pt of points) {
      const lat = pt.latitude ?? "";
      const lng = pt.longitude ?? "";
      const ne = lat && lng ? `N${lat} E${lng}` : "";

      const img1 = photoUrls[0] ? await fetchImageBytes(photoUrls[0]) : null;
      const img2 = photoUrls[1] ? await fetchImageBytes(photoUrls[1]) : null;
      const photoCellChildren: Paragraph[] = [];

      if (img1) {
        photoCellChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new ImageRun({ data: img1, transformation: { width: 220, height: 140 } })],
          })
        );
      }
      if (img2) {
        photoCellChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new ImageRun({ data: img2, transformation: { width: 220, height: 140 } })],
          })
        );
      }
      if (!photoCellChildren.length) photoCellChildren.push(new Paragraph(""));

      rows.push(
        new TableRow({
          children: [
            normalCell(String(pt.seq ?? "")),
            normalCell(String(pt.km ?? pt.kms ?? "")),
            normalCell(ne),
            normalCell(pt.details ?? report.description ?? ""),
            normalCell(pt.location_text ?? report.location ?? ""),
            new TableCell({ children: photoCellChildren }),
            normalCell(pt.vehicle_movement ?? report.vehicle_movement ?? ""),
          ],
        })
      );
    }

    const table = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows,
    });

    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              text: "ROUTE SURVEY REPORT",
              heading: HeadingLevel.TITLE,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: projectName ? `Project: ${projectName}` : "Project",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: `Report: ${report.category || "Report"}`,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: report.created_at ? new Date(report.created_at).toLocaleString() : "",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({ text: "" }),
            new Paragraph({
              text: "Stage Details",
              heading: HeadingLevel.HEADING_1,
            }),
            table,
          ],
        },
      ],
    });

    const buf = await Packer.toBuffer(doc);
    const base64 = Buffer.from(buf).toString("base64");
    const filename = `${(projectName || "project").replaceAll(" ", "_")}_${String(report.category || "report").replaceAll(
      " ",
      "_"
    )}.docx`;

    return Response.json({ base64, filename });
  } catch (e: any) {
    if (e?.message === "Unauthorized") {
      return Response.json({ error: "Unauthorized" }, { status: 401 });
    }
    return Response.json({ error: "Failed to generate DOCX" }, { status: 500 });
  }
}
