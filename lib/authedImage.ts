"use client";

// Load a report photo so it can be drawn on a <canvas> and exported.
//
// External (S3) photos go through /api/image-proxy, which needs the login
// token. <img src> / new Image() can't send an Authorization header (and the
// Secure auth cookie isn't stored over plain HTTP), so we fetch() the bytes and
// hand back a same-origin blob: URL — which also keeps the canvas untainted.
// The caller must call revoke() when done.

function authHeaders(): Record<string, string> {
  const t = typeof window !== "undefined" ? localStorage.getItem("auth_token") : null;
  return t ? { Authorization: `Bearer ${t}` } : {};
}

export async function loadCanvasSafeImage(
  photoUrl: string,
  signal?: AbortSignal
): Promise<{ src: string; revoke: () => void }> {
  const origin = typeof window !== "undefined" ? window.location.origin : "";
  const isExternal = /^https?:\/\//i.test(photoUrl) && !photoUrl.startsWith(origin);
  if (!isExternal) return { src: photoUrl, revoke: () => {} };

  const res = await fetch(`/api/image-proxy?url=${encodeURIComponent(photoUrl)}`, {
    headers: authHeaders(),
    credentials: "include",
    signal,
  });
  if (!res.ok) throw new Error(`Image load failed (${res.status})`);
  const src = URL.createObjectURL(await res.blob());
  return { src, revoke: () => URL.revokeObjectURL(src) };
}
