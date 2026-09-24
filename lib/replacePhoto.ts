// Save an edited photo over the original — used by the draw and crop tools.
//
// Both tools used to upload their result as a NEW photo on the report, so
// every save added one more picture (and crop then deleted the original,
// which moved the photo to the end of the list and reset its settings).
//
// Instead: upload the edited image on its own, then point the existing photo
// at it. Same photo, same position, same Word-export tick — new picture.

function authHeaders(): Record<string, string> {
  const t = typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  return t ? { Authorization: `Bearer ${t}` } : {};
}

export async function replacePhotoImage(opts: {
  reportId: string;
  photoId: string;
  blob: Blob;
  fileName: string;
  width: number;
  height: number;
  // Re-editable drawings: the draw tool passes the vector strokes and the
  // untouched base image so the annotation can be reopened and edited later.
  // Omitted by other callers (e.g. crop), which clears any stored strokes.
  annoJson?: string | null;
  annoBaseUrl?: string | null;
}): Promise<string> {
  // 1. Upload the image by itself. No reportId: with one, /api/upload also
  //    creates a new photo row, which is exactly the duplicate we avoid.
  const fd = new FormData();
  fd.append("file", opts.blob, opts.fileName);
  fd.append("folder", "uploads");
  const up = await fetch("/api/upload", {
    method: "POST",
    credentials: "include",
    headers: authHeaders(),
    body: fd,
  });
  const upData = (await up.json().catch(() => ({}))) as { url?: string; error?: string };
  if (!up.ok || !upData.url) throw new Error(upData.error || "Upload failed");

  // 2. Swap it into the existing photo.
  const res = await fetch(`/api/reports/${encodeURIComponent(opts.reportId)}/photos`, {
    method: "PATCH",
    credentials: "include",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({
      photoId: opts.photoId,
      url: upData.url,
      file_name: opts.fileName,
      width: opts.width,
      height: opts.height,
      ...(opts.annoJson !== undefined ? { anno_json: opts.annoJson } : {}),
      ...(opts.annoBaseUrl !== undefined ? { anno_base_url: opts.annoBaseUrl } : {}),
    }),
  });
  const data = (await res.json().catch(() => ({}))) as { error?: string };
  if (!res.ok) throw new Error(data.error || "Could not replace the photo");

  return upData.url;
}
