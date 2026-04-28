/**
 * Statement PDFs often embed invoice links as URI strings in the file stream.
 * Text extraction alone does not expose annotations; a lightweight scan finds many URLs.
 */

/**
 * Unescape common PDF literal-string escapes inside a URI capture.
 * @param {string} s
 */
function decodePdfLiteralUriFragment(s) {
  return String(s || "")
    .replace(/\\\)/g, ")")
    .replace(/\\\(/g, "(")
    .replace(/\\\\/g, "\\")
    .trim();
}

/**
 * INV tokens anywhere in decoded PDF bytes (annotations / uncompressed streams often still visible).
 * @param {Buffer} buf
 * @returns {string[]} uppercased INV- numbers
 */
export function extractInvNumbersFromPdfBinary(buf) {
  if (!buf?.length) return [];
  const raw = buf.toString("latin1");
  const refs = new Set();
  const re = /\bINV-\d{4,8}\b/gi;
  let m;
  while ((m = re.exec(raw))) {
    refs.add(m[0].toUpperCase());
    if (refs.size >= 80) break;
  }
  return [...refs];
}

/**
 * Link annotations store targets as `/URI (https://...)` even when visible text is missing from text extraction.
 * @param {Buffer} buf
 * @returns {string[]} unique https? URLs
 */
function extractUrlsFromPdfUriAnnotations(buf) {
  if (!buf?.length) return [];
  const raw = buf.toString("latin1");
  const found = new Set();
  /** PDF: /URI (https://example.com/path) — stop at first unescaped ) */
  const uriRe = /\/URI\s*\(\s*(https?:\/\/(?:\\.|[^\)\x00])+?)\)/gi;
  let m;
  while ((m = uriRe.exec(raw))) {
    let frag = decodePdfLiteralUriFragment(m[1]).replace(/\\\r?\n/g, "");
    frag = frag.replace(/\\\)/g, ")");
    try {
      const parsed = new URL(frag.split(/\s/)[0]);
      if (parsed.protocol === "http:" || parsed.protocol === "https:") {
        found.add(parsed.href);
      }
    } catch {
      continue;
    }
    if (found.size >= 40) break;
  }
  return [...found];
}

/**
 * @param {Buffer} buf
 * @returns {string[]} unique https? URLs (capped)
 */
export function extractHttpsUrlsFromPdfBuffer(buf) {
  if (!buf?.length) return [];
  const raw = buf.toString("latin1");
  const found = new Set();
  for (const u of extractUrlsFromPdfUriAnnotations(buf)) found.add(u);
  const re =
    /https?:\/\/[A-Za-z0-9._~:/?#[\]@!$&'()*+,;=%\-]{12,900}/gi;
  let m;
  while ((m = re.exec(raw))) {
    let u = m[0];
    u = u.replace(/[\]\)\}\>'"]+$/g, "");
    try {
      const parsed = new URL(u);
      if (parsed.protocol === "http:" || parsed.protocol === "https:") {
        found.add(parsed.href);
      }
    } catch {
      continue;
    }
    if (found.size >= 40) break;
  }
  return [...found];
}
