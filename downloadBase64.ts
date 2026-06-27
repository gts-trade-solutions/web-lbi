// utils/downloadBase64.ts

export function downloadBase64AsFile(
  base64: string,
  filename: string,
  mime: string
) {
  if (!base64 || typeof base64 !== "string") {
    throw new Error(`Invalid base64 value. Got: ${typeof base64}`);
  }

  const clean = normalizeBase64Strict(base64);

  // Decode -> bytes (use modern API if available)
  const bytes: Uint8Array =
    // @ts-ignore - some TS libs don't know this yet
    typeof Uint8Array.fromBase64 === "function"
      // @ts-ignore
      ? Uint8Array.fromBase64(clean)
      : uint8FromAtob(clean);

  const blob = new Blob([bytes], { type: mime });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename || "download";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function uint8FromAtob(b64: string) {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function normalizeBase64Strict(input: string) {
  let s = (input ?? "").trim();

  // Remove possible data-url prefix
  s = s.replace(/^data:.*;base64,/, "");

  // If it looks percent-encoded, decode it (base64 never contains % normally)
  if (s.includes("%")) {
    try {
      s = decodeURIComponent(s);
    } catch {
      // ignore
    }
  }

  // IMPORTANT: if + got converted to spaces, restore it
  s = s.replace(/ /g, "+");

  // Remove newlines/tabs etc (keep +)
  s = s.replace(/[\r\n\t]/g, "");

  // base64url -> base64
  s = s.replace(/-/g, "+").replace(/_/g, "/");

  // Padding
  while (s.length % 4 !== 0) s += "=";

  // Validate characters (after normalization)
  const bad = s.match(/[^A-Za-z0-9+/=]/g);
  if (bad) {
    const uniqueBad = Array.from(new Set(bad)).join("");
    throw new Error(
      `Returned "base64" is not valid base64. Bad chars: "${uniqueBad}". Head: "${s.slice(
        0,
        60
      )}"`
    );
  }

  return s;
}
