import { downloadBase64AsFile } from "@/utils/downloadBase64";

const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export async function exportReportDocx(reportId: string) {
  const res = await fetch(`/api/reports/${encodeURIComponent(reportId)}/docx`, {
    method: "GET",
    credentials: "include",
  });
  const data = await res.json().catch(() => ({} as any));
  if (!res.ok) throw new Error(data?.error || "Failed to generate DOCX");

  if (!data?.base64) {
    throw new Error("DOCX API did not return { base64 }");
  }

  downloadBase64AsFile(
    data.base64,
    data.filename || `report-${reportId}.docx`,
    DOCX_MIME
  );
}
