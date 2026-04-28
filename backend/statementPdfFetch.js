/**
 * Statement PDFs often link to viewer pages (e.g. Xero `in.xero.com/m/...`) rather than raw
 * `%PDF` bytes. This module follows one level of indirection when the first response is HTML.
 */

const BROWSER_UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36";

function isPdfMagic(buf) {
  return buf?.length >= 4 && buf.slice(0, 4).toString("latin1") === "%PDF";
}

/** Trust these hosts for follow-up fetches after a statement link (avoids open redirect abuse). */
function hostAllowedForStatementFollow(hostname) {
  const h = String(hostname || "").toLowerCase();
  return (
    h.endsWith("xero.com") ||
    h.endsWith("xero.net") ||
    h.endsWith("xeroservices.com")
  );
}

function safeJoinUrl(base, candidate) {
  try {
    return new URL(candidate, base).href;
  } catch {
    return null;
  }
}

/**
 * @param {string} html
 * @param {string} pageUrl resolved page URL (for relative links)
 * @returns {string[]}
 */
export function extractCandidatePdfUrlsFromHtml(html, pageUrl) {
  const seen = new Set();
  const add = (u) => {
    if (!u || seen.size >= 24) return;
    const abs = u.startsWith("http") ? u : safeJoinUrl(pageUrl, u);
    if (!abs) return;
    try {
      const p = new URL(abs);
      if (!/^https?:$/i.test(p.protocol)) return;
      if (
        /\.pdf(?:$|[?#])/i.test(p.pathname + p.search) ||
        hostAllowedForStatementFollow(p.hostname)
      ) {
        seen.add(p.href.split("#")[0]);
      }
    } catch {
      return;
    }
  };

  const meta = html.match(
    /http-equiv\s*=\s*["']refresh["'][^>]*content\s*=\s*["'][^"']*url\s*=\s*([^"']+)/i,
  );
  if (meta) add(meta[1].trim());

  const patterns = [
    /https?:\/\/[^\s"'<>()]{6,2400}(?:[\w/%._-]+\.pdf(?:\?[^\s"'<>]{0,800})?)/gi,
    /"((?:https?:)?\/\/[^"]+\.pdf[^"]*)"/gi,
    /'((?:https?:)?\/\/[^']+\.pdf[^']*)'/gi,
    /"(https?:\/\/[^"]*(?:download|attachment|attachmentId|invoice|document)[^"]{0,400})"/gi,
    /href\s*=\s*["']([^"']*(?:pdf|PDF|download|attachment|invoice)[^"']*)["']/gi,
    /"(?:pdfUrl|downloadUrl|documentUrl|fileUrl|contentUrl)"\s*:\s*"([^"]+)"/gi,
    /'(?:pdfUrl|downloadUrl|documentUrl|fileUrl|contentUrl)'\s*:\s*'([^']+)'/gi,
  ];
  for (const re of patterns) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(html)) && seen.size < 24) {
      let raw = (m[1] != null && m[1] !== "" ? m[1] : m[0]).trim();
      raw = raw.replace(/^\/\//, "https://");
      add(raw);
    }
  }

  /** SPA shells often embed Xero API-ish paths without `.pdf` in the first HTML chunk. */
  const xeroLoose =
    /https?:\/\/[a-z0-9.-]*xero\.com\/[^\s"'<]{8,2200}/gi;
  let xm;
  let looseN = 0;
  while ((xm = xeroLoose.exec(html)) && looseN < 10 && seen.size < 24) {
    const s = xm[0].replace(/[,;)\]}>'"]+$/g, "");
    if (
      /pdf|download|attachment|document|files?|Organisation|network|invoice/i.test(s) &&
      !/\.(?:png|svg|jpe?g|gif|woff2?|css|ico)(?:\?|$)/i.test(s)
    ) {
      add(s);
      looseN++;
    }
  }

  return [...seen];
}

/**
 * Fetch URL; return PDF buffer, or parse HTML and try embedded PDF / Xero URLs (limited depth).
 *
 * @param {string} url
 * @param {number} maxBytes
 * @param {number} timeoutMs
 * @returns {Promise<Buffer|null>}
 */
export async function fetchPdfBufferResolvingViewerPages(url, maxBytes, timeoutMs = 25000) {
  const visited = new Set();

  async function step(href, depth, maxDepth) {
    if (depth > maxDepth) return null;
    let normalized;
    try {
      normalized = new URL(href).href.split("#")[0];
    } catch {
      return null;
    }
    if (visited.has(normalized)) return null;
    visited.add(normalized);

    const res = await fetch(normalized, {
      redirect: "follow",
      headers: {
        Accept:
          "application/pdf,application/octet-stream,text/html;q=0.9,application/xhtml+xml;q=0.8,*/*;q=0.7",
        "User-Agent": BROWSER_UA,
      },
      signal: AbortSignal.timeout(timeoutMs),
    });
    if (!res.ok) return null;
    const buf = Buffer.from(await res.arrayBuffer());
    if (!buf.length || buf.length > maxBytes) return null;
    if (isPdfMagic(buf)) return buf;

    const ct = (res.headers.get("content-type") || "").toLowerCase();
    const sniff = buf.slice(0, Math.min(1200, buf.length)).toString("utf8");
    const looksHtml =
      ct.includes("html") ||
      sniff.trimStart().startsWith("<") ||
      /<html[\s>]/i.test(sniff);

    if (!looksHtml || depth >= maxDepth) return null;

    const html = buf.toString("utf8").slice(0, Math.min(buf.length, 4_000_000));
    const candidates = extractCandidatePdfUrlsFromHtml(html, normalized);
    for (const next of candidates) {
      try {
        const nu = new URL(next);
        if (
          !hostAllowedForStatementFollow(nu.hostname) &&
          !/\.pdf(?:$|[?#])/i.test(nu.pathname + nu.search)
        ) {
          continue;
        }
      } catch {
        continue;
      }
      const pdf = await step(next, depth + 1, maxDepth);
      if (pdf) return pdf;
    }
    return null;
  }

  /** Xero viewer: first hop often HTML; allow 4 steps for redirects + embedded asset. */
  return step(url, 0, 4);
}
