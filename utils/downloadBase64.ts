export function downloadBase64AsFile(
  base64: string,
  filename: string,
  mime: string
) {
  const clean = normalizeBase64Strict(base64);

  const bin = atob(clean);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);

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

function normalizeBase64Strict(input: string) {
  let s = (input ?? "").trim();

  // remove data url prefix if present
  s = s.replace(/^data:.*;base64,/, "");

  // decode percent-encoding if any
  if (s.includes("%")) {
    try {
      s = decodeURIComponent(s);
    } catch {}
  }

  // restore + that may become spaces
  s = s.replace(/ /g, "+");

  // remove newlines/tabs
  s = s.replace(/[\r\n\t]/g, "");

  // base64url -> base64
  s = s.replace(/-/g, "+").replace(/_/g, "/");

  // pad
  while (s.length % 4 !== 0) s += "=";

  // validate
  const bad = s.match(/[^A-Za-z0-9+/=]/g);
  if (bad) {
    const uniq = Array.from(new Set(bad)).join("");
    throw new Error(`Invalid base64 from server. Bad chars: "${uniq}"`);
  }

  return s;
}
