/**
 * Heuristic extraction of invoice / campaign document fields from plain text (PDF extract).
 * Handles Bizcommunity-style, Volt.africa / marketing IO layouts, and generic patterns.
 */

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

export function formatDocDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const d = `${date.getDate()}`.padStart(2, "0");
  const mon = MONTHS[date.getMonth()] || "";
  const yy = `${date.getFullYear()}`.slice(-2);
  return `${d}-${mon}-${yy}`;
}

function parseAmount(s) {
  if (!s) return "";
  const cleaned = String(s).replace(/[^\d,.-]/g, "").replace(/,/g, "");
  const n = parseFloat(cleaned);
  if (Number.isNaN(n)) return "";
  return n.toLocaleString("en-ZA", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

/** Numeric parse for reconciliation (no locale formatting). */
function parseMoneyFloat(s) {
  if (!s) return NaN;
  const n = parseFloat(String(s).replace(/,/g, ""));
  return Number.isNaN(n) ? NaN : n;
}

/**
 * Classic invoice footer: SUBTOTAL + TAX + TOTAL, optional tax-summary row "TAX … NET …".
 * Avoids mistaking payment terms ("Net 30") or stray words for monetary NET.
 */
/** Line-item row: ISO-like date + description ending in BO No / booking ref */
function extractCampaignFromBoLineItem(flatOneLine) {
  const one = String(flatOneLine || "").replace(/\s+/g, " ");
  const re =
    /\d{1,2}\/\d{1,2}\/\d{4}\s+(.{12,420}?\bBO\s*(?:No\.?|number|nr\.?)?\s*[:\s]*[\d\s]{7,14})/gi;
  let best = "";
  let m;
  while ((m = re.exec(one)) !== null) {
    const body = m[1];
    if (/^\s*terms\b/i.test(body) || /^\s*bill\s*to\b/i.test(body)) continue;
    if (
      /\b(advertisement|editorial|digitorial|A4\s+ad)\b/i.test(body) ||
      /\bFBR\b/i.test(body)
    ) {
      best = trimCampaignProductCapture(body);
      break;
    }
    if (!best && body.length > 20) best = trimCampaignProductCapture(body);
  }
  return best && best.length >= 12 ? best : "";
}

function extractClassicInvoiceFooterAmounts(flat) {
  const out = {
    subtotal: "",
    tax: "",
    total: "",
    netLabel: "",
  };
  out.subtotal = matchFirst(flat, [
    /sub[-\s]?total\s*\(\s*net\s*\)\s*[:\s]*(?:ZAR\s*)?([\d\s,]+\.\d{2})/i,
    /sub[-\s]?total\s*[:\s]*(?:ZAR\s*)?([\d\s,]+\.\d{2})/i,
  ]);
  const taxNetPair = flat.match(
    /\bTAX\b\s+([\d\s,]+\.\d{2})\s+\bNET\b\s+([\d\s,]+\.\d{2})/i,
  );
  if (taxNetPair) {
    out.tax = taxNetPair[1].trim();
    out.netLabel = taxNetPair[2].trim();
  }
  if (!out.tax) {
    out.tax = matchFirst(flat, [
      /(?:^|[^\w])\bTAX\b\s*[:\s]*(?:ZAR\s*)?([\d\s,]+\.\d{2})(?=\s+(?:TOTAL|NET|SUBTOTAL|BALANCE|$))/i,
    ]);
  }
  out.total = matchFirst(flat, [
    /\bTOTAL\s+ZAR\s*[:\s]*([\d\s,]+\.\d{2})/i,
    /\bTOTAL\b\s*[:\s]*(?:ZAR\s*)?([\d\s,]+\.\d{2})(?=\s+(?:BALANCE|ZAR|$))/i,
    /balance\s*due\s*[:\s]*(?:ZAR\s*)?([\d\s,]+\.\d{2})/i,
  ]);
  return out;
}

/** Pull first capturing group across multiple patterns. */
function matchFirst(text, patterns) {
  const t = text;
  for (const re of patterns) {
    const m = t.match(re);
    if (m && m[1]) return m[1].trim();
  }
  return "";
}

/** Money tokens with or without thousand separators */
const MONEY_RE = /\d{1,3}(?:,\d{3})*\.\d{2}|\d+\.\d{2}/g;

const MONEY_CAP =
  "(?:(?:\\d{1,3}(?:,\\d{3})*|\\d+)\\.\\d{2})";

/** Last monetary amount on a line (handles 30,000.00 and 34500.00). */
function lastMoneyOnLine(line) {
  if (!line) return "";
  const s = String(line).replace(/\s+/g, " ");
  const nums = s.match(MONEY_RE);
  if (nums && nums.length) return nums[nums.length - 1];
  return "";
}

/** First money on a line — use for TOTAL VAT EXCLUDED when the row lists all three figures on one line. */
function firstMoneyOnLine(line) {
  if (!line) return "";
  const s = String(line).replace(/\s+/g, " ");
  const nums = s.match(MONEY_RE);
  if (nums && nums.length) return nums[0];
  return "";
}

/** First money after a label, or before it, in a flattened fragment (PDFs often have one long line). */
function moneyNearLabel(flatFragment, labelPattern) {
  const f = flatFragment.replace(/\s+/g, " ");
  const after = f.match(
    new RegExp(labelPattern + `\\D{0,240}?(${MONEY_CAP})`, "i"),
  );
  if (after?.[1]) return after[1];
  const before = f.match(
    new RegExp(`(${MONEY_CAP})\\D{0,240}?` + labelPattern, "i"),
  );
  return before?.[1] || "";
}

/**
 * Some Volt PDFs list column amounts in an order that breaks "label + number" pairing.
 * Find three money values in the footer region where tax ≈15% of base and total = base + tax.
 */
function inferVoltFooterTriplet(flat) {
  const low = flat.toLowerCase();
  const footerHint = low.lastIndexOf("total vat excluded");
  const footStart =
    footerHint >= 0
      ? Math.max(0, footerHint - 600)
      : Math.max(0, flat.length - 6000);
  const slice = flat.slice(footStart);
  const nums = [...slice.matchAll(/\d{1,3}(?:,\d{3})*\.\d{2}/g)].map((m) => ({
    raw: m[0],
    val: parseFloat(m[0].replace(/,/g, "")),
  }));
  /** Prefer last triplet in footer (summary is usually after line items). */
  for (let i = nums.length - 3; i >= 0; i--) {
    const a = nums[i].val;
    const b = nums[i + 1].val;
    const c = nums[i + 2].val;
    if (a < 500 || a > 1e8) continue;
    const ratio = b / a;
    /** SA standard 15% — exclude mistaken triplets where VAT was computed on VAT-inclusive base (~13%). */
    if (
      ratio >= 0.135 &&
      ratio <= 0.165 &&
      Math.abs(c - a - b) < Math.max(2, a * 0.001)
    ) {
      return {
        excluded: nums[i].raw,
        vatTax: nums[i + 1].raw,
        included: nums[i + 2].raw,
      };
    }
  }
  return { excluded: "", vatTax: "", included: "" };
}

/**
 * Normalise PDF quirks: NBSP/ZWSP, glued "TotalVATExcluded", then split pseudo-lines so
 * regexes see the same tokens whether or not the PDF emitted real newlines.
 */
function normalizePdfPlainText(raw) {
  let t = String(raw || "")
    .replace(/\r\n/g, "\n")
    .replace(/[\u00a0\u200b\ufeff]/g, " ");
  t = t.replace(/(total)(vat)(excluded|included)/gi, "$1 $2 $3");
  t = t.replace(/(amount)(vat)/gi, "$1 $2 ");
  /** DStv exports: labels glued to values (no space / no word boundary before digits) */
  t = t.replace(/\bINVOICENO(?=\d)/gi, "INVOICE NO ");
  t = t.replace(/\bINVOICENO\b/gi, "INVOICE NO");
  t = t.replace(/\bINVOICEDATE(?=\d)/gi, "INVOICE DATE ");
  t = t.replace(/\bINVOICEDATE\b/gi, "INVOICE DATE");
  /** Keep newlines — collapsing here breaks line-based VAT row detection */
  return t;
}

function buildExtractionLines(text) {
  return text
    .split(/\n/)
    .flatMap((line) => {
      const L = line.replace(/\s+/g, " ").trim();
      if (!L) return [];
      return L.split(
        /\s+(?=\b(?:Total\s+VAT|Amount\s+VAT|Sub[\s-]?total|Grand\s+total|Invoice\s+total|VAT\s+Exclusive|VAT\s+Inclusive)\b)/i,
      );
    })
    .map((l) => l.replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

/**
 * Volt.africa summary table (exact labels):
 * TOTAL VAT EXCLUDED → net / gross base · TOTAL VAT (alone) → tax · TOTAL VAT INCLUDED → payable total.
 * Must use (? !EXCLUDED|INCLUDED) with /i — JS (?!excluded|included) is case-sensitive and breaks on "INCLUDED".
 */
function extractVoltSummaryTriple(f) {
  const flat = f.replace(/\s+/g, " ");
  let excluded = moneyNearLabel(flat, "\\bTOTAL\\s+VAT\\s+EXCLUDED\\b");
  let included = moneyNearLabel(flat, "\\bTOTAL\\s+VAT\\s+INCLUDED\\b");
  let vatTax = matchFirst(flat, [
    /\bTOTAL\s+VAT\s+(?!EXCLUDED\b)(?!INCLUDED\b)\s*:?\s*R?\s*([\d\s,]+\.\d{2})\b/i,
    /\bTOTAL\s+VAT\s+(?!EXCLUDED\b)(?!INCLUDED\b)\s*:?\s*R?\s*([\d\s,]+\.?\d*)\b/i,
    /([\d\s,]+\.\d{2})\s*(?:R\s*)?\bTOTAL\s+VAT\b(?!\s*(?:EXCLUDED|INCLUDED)\b)/i,
  ]);

  if (!excluded)
    excluded =
      moneyNearLabel(flat, "total\\s*v\\.?a\\.?t\\.?\\s*excluded") ||
      moneyNearLabel(flat, "amount\\s*v\\.?a\\.?t\\.?\\s*exclud(?:ed|e)?");
  if (!included)
    included =
      moneyNearLabel(flat, "total\\s*v\\.?a\\.?t\\.?\\s*included") ||
      moneyNearLabel(flat, "amount\\s*v\\.?a\\.?t\\.?\\s*includ");

  /** Three-value sequence (single-line PDF) */
  if (!excluded || !included || !vatTax) {
    const seq = flat.match(
      new RegExp(
        `TOTAL\\s+VAT\\s+EXCLUDED\\D{0,140}?(${MONEY_CAP})\\D{0,220}?TOTAL\\s+VAT\\s+(?!EXCLUDED\\b)(?!INCLUDED\\b)\\D{0,120}?(${MONEY_CAP})\\D{0,220}?TOTAL\\s+VAT\\s+INCLUDED\\D{0,140}?(${MONEY_CAP})`,
        "i",
      ),
    );
    if (seq) {
      if (!excluded) excluded = seq[1];
      if (!vatTax) vatTax = seq[2];
      if (!included) included = seq[3];
    }
  }

  return { excluded, vatTax, included };
}

function extractVatSummaryFromFlat(flatOneLine) {
  const volt = extractVoltSummaryTriple(flatOneLine);
  const f = flatOneLine.replace(/\s+/g, " ");
  let excluded =
    volt.excluded ||
    moneyNearLabel(f, "total\\s*v\\.?a\\.?t\\.?\\s*excluded") ||
    moneyNearLabel(f, "amount\\s*v\\.?a\\.?t\\.?\\s*exclud(?:ed|e)?") ||
    moneyNearLabel(f, "vat\\s*exclusive") ||
    moneyNearLabel(f, "exclusive\\s*v\\.?a\\.?t\\.?");
  let included =
    volt.included ||
    moneyNearLabel(f, "total\\s*v\\.?a\\.?t\\.?\\s*included") ||
    moneyNearLabel(f, "amount\\s*v\\.?a\\.?t\\.?\\s*includ") ||
    moneyNearLabel(f, "vat\\s*inclusive") ||
    moneyNearLabel(f, "inclusive\\s*v\\.?a\\.?t\\.?");
  let vatTax =
    volt.vatTax ||
    matchFirst(f, [
      /\bTOTAL\s+VAT\s+(?!EXCLUDED\b)(?!INCLUDED\b)\s*:?\s*R?\s*([\d\s,]+\.\d{2})\b/i,
      /\btotal\s+v\.?a\.?t\.?\s+(?!EXCLUDED\b)(?!INCLUDED\b)\s*R?\s*([\d\s,]+\.?\d*)\b/i,
    ]) ||
    "";

  /** Three-value sequence — generic wording */
  if (!excluded || !included) {
    const seq = f.match(
      new RegExp(
        `total\\s*v\\.?a\\.?t\\.?\\s*excluded\\D{0,120}?(${MONEY_CAP})\\D{0,200}?total\\s*v\\.?a\\.?t\\.?\\s+(?!EXCLUDED\\b)(?!INCLUDED\\b)\\D{0,100}?(${MONEY_CAP})\\D{0,200}?total\\s*v\\.?a\\.?t\\.?\\s*included\\D{0,120}?(${MONEY_CAP})`,
        "i",
      ),
    );
    if (seq) {
      if (!excluded) excluded = seq[1];
      if (!vatTax) vatTax = seq[2];
      if (!included) included = seq[3];
    }
  }

  return { excluded, vatTax, included };
}

/**
 * SA IO templates (Volt, etc.): three-line VAT summary where flat-string regexes often fail.
 */
function extractVatSummaryFromLines(lines) {
  const norm = lines.map((l) => l.replace(/\s+/g, " ").trim()).filter(Boolean);
  let excluded = "";
  let vatTax = "";
  let included = "";

  for (let i = 0; i < norm.length; i++) {
    const L = norm[i];
    const LNext = norm[i + 1] || "";
    const low = L.toLowerCase();

    if (
      /total\s*v\.?a\.?t\.?\s*excluded|amount\s*v\.?a\.?t\.?\s*exclud|^v\.?a\.?t\.?\s*exclusive|exclusive\s*v\.?a\.?t\.?\b|^vat\s*excluded\b/.test(
        low,
      )
    ) {
      excluded =
        firstMoneyOnLine(L) ||
        firstMoneyOnLine(LNext) ||
        lastMoneyOnLine(L) ||
        lastMoneyOnLine(LNext);
      continue;
    }
    if (
      /total\s*v\.?a\.?t\.?\s*included|amount\s*v\.?a\.?t\.?\s*includ|^v\.?a\.?t\.?\s*inclusive|inclusive\s*v\.?a\.?t\.?\b|^vat\s*included\b/.test(
        low,
      )
    ) {
      included =
        lastMoneyOnLine(L) ||
        lastMoneyOnLine(LNext) ||
        firstMoneyOnLine(L) ||
        firstMoneyOnLine(LNext);
      continue;
    }
    /** Middle row: "TOTAL VAT" only (not EXCLUDED / INCLUDED) — case-insensitive word boundaries */
    if (/\bTOTAL\s+VAT\b/i.test(L) && !/\bTOTAL\s+VAT\s+(?:EXCLUDED|INCLUDED)\b/i.test(L)) {
      const m = L.match(
        /\bTOTAL\s+VAT\s+(?!EXCLUDED\b)(?!INCLUDED\b)\s*:?\s*R?\s*([\d,\s]+\.\d{2})\b/i,
      );
      if (m) {
        vatTax = m[1].trim();
      } else {
        vatTax = lastMoneyOnLine(L) || lastMoneyOnLine(LNext);
      }
    }
  }

  return { excluded, vatTax, included };
}

/** Prefer longer / cleaner capture for multi-line PDF dumps */
function cleanLabelValue(s) {
  if (!s) return "";
  return String(s)
    .replace(/\s+/g, " ")
    .replace(/\s*,\s*$/g, "")
    .trim()
    .slice(0, 180);
}

/** Flat PDF strings often glue the product slug to the next label — cut before line-item / footer tokens */
function trimCampaignProductCapture(raw) {
  let t = String(raw || "").trim();
  t = t.split(
    /\s+(?:START DATE|END DATE|TOTAL\s+VAT|DESCRIPTION\s+QTY|DESCRIPTION\b|UNIT\s+PRICE|QTY\s*\||IMPRESSIONS|CAMPAIGN\s+(?:START|END|NO)|Banking|Swift|INVOICE\s+DETAILS)/i,
  )[0];
  return cleanLabelValue(t);
}

/** DD/MM/YYYY only — not a campaign description */
function isStandaloneSlashDateLine(s) {
  const t = String(s || "").trim();
  return /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(t);
}

function isCampaignStartDateField(s) {
  const t = String(s || "").replace(/\s+/g, " ").trim();
  if (!t) return false;
  if (/^\s*campaign\s*start\s*date\b/i.test(t)) return true;
  if (/^\s*campaign\s*start\s*date\s*[:\s]*[\d/\s-]+\s*$/i.test(t)) return true;
  return isStandaloneSlashDateLine(t);
}

/** Table headers / banking / supplier lines mistaken for campaign description */
function isInvoiceTableHeaderOrJunkCampaign(s) {
  const t = String(s || "").replace(/\s+/g, " ").trim();
  if (!t) return true;
  if (/\bDESCRIPTION\s+QTY\b/i.test(t)) return true;
  if (/\bUNIT\s+PRICE\b/i.test(t) && /\b(?:QTY|IMPRESSIONS)\b/i.test(t)) return true;
  if (/^\s*VOLTAfrica\b/i.test(t)) return true;
  if (/^\s*volt\.?\s*africa\b/i.test(t) && /\bpty\b/i.test(t)) return true;
  if (/\bBanking\s+Details\b|\bSwift\s+Code\b|\bBranch\s+Code\b/i.test(t)) return true;
  if (/\bQuantity\b/i.test(t) && /\bUnit\s+Price\b/i.test(t) && /\bAmount\s+ZAR\b/i.test(t))
    return true;
  return false;
}

/** Statement headers like "From Date … To Date …" mistaken for supplier/client */
function isStatementPeriodGarbageField(s) {
  const t = String(s || "").replace(/\s+/g, " ").trim();
  if (!t) return false;
  if (/\bFrom\s+Date\b/i.test(t) || /\bTo\s+Date\b/i.test(t)) return true;
  if (/^\s*Date\s+\d/i.test(t) && /\bTo\s+Date\b/i.test(t)) return true;
  return false;
}

function isIpointDocument(flat) {
  return /\bipoint\b/i.test(flat) && /\bmedia\b/i.test(flat);
}

/** DStv Media Sales tax invoices — grid labels INVOICE NO., INVOICE DATE (M/D/Y), footer Gross/Net/VAT/Total Inv. Amount */
function isDstvMediaSalesDocument(flat) {
  if (/\bDStv\s+Media\s+Sales\b/i.test(flat)) return true;
  if (/\bDSTV\s+Media\s+Sales\b/i.test(flat)) return true;
  if (/\bMultiChoice\b/i.test(flat) && /\bMedia\s+Sales\b/i.test(flat)) return true;
  /** Header/logo may be non-text — match the distinctive metadata grid + totals row */
  if (
    /\bADVERTISER\b/i.test(flat) &&
    /\bPRODUCT\b/i.test(flat) &&
    /\bINVOICE\s+NO\.?\b/i.test(flat) &&
    /\bINVOICE\s+DATE\b/i.test(flat) &&
    (/\bTotal\s+Inv\.?\s*Amount\b/i.test(flat) ||
      /\bREF\.?\s*NO\.?\s*[:\s]*0000\d+/i.test(flat) ||
      /\bCAMPAIGN\s+NO\.?\s*[:\s]*\d+/i.test(flat))
  )
    return true;
  /** “DStv” missing from text layer but full tax-invoice + advertiser grid + Gross/Net/VAT block */
  if (
    /\bTAX\s+INVOICE\b/i.test(flat) &&
    /\bADVERTISER\b/i.test(flat) &&
    /\bPRODUCT\b/i.test(flat) &&
    /\bINVOICE\s+DATE\b/i.test(flat) &&
    /\bGross\b/i.test(flat) &&
    /\bNet\b/i.test(flat) &&
    /\bVAT\b/i.test(flat) &&
    (/\bTotal\s+Inv/i.test(flat) || /\bVAT\s+REG/i.test(flat))
  )
    return true;
  return false;
}

/** Booking confirmation PDF (not tax invoice) — PRINTED DATE + GROSS TOTAL footer */
function isDstvOrderConfirmationDocument(flat) {
  const f = String(flat || "");
  if (!isDstvMediaSalesDocument(f)) return false;
  if (!/\bORDER\s+CONFIRMATION\b/i.test(f)) return false;
  if (!/\bPRINTED\s+DATE\b/i.test(f)) return false;
  if (!/\bGROSS\s+TOTAL\b/i.test(f)) return false;
  return true;
}

/** ZA-style money on order confirmations: spaces as thousands group, last comma = decimals (e.g. 13 950,00) */
function normalizeDstvOrderConfirmationMoney(s) {
  let t = String(s || "")
    .trim()
    .replace(/^R\s*/i, "")
    .replace(/\s+/g, "");
  const li = t.lastIndexOf(",");
  if (li >= 0 && li < t.length - 1) {
    const intPart = t.slice(0, li).replace(/\./g, "");
    const decPart = t.slice(li + 1);
    if (/^\d{2}$/.test(decPart)) return `${intPart}.${decPart}`;
  }
  return t.replace(/,/g, ".");
}

function parseDstvOrderConfirmationFooterTotals(flat) {
  const one = String(flat || "").replace(/\s+/g, " ");
  /** One ZA amount: thousands with spaces, comma decimals (e.g. 93 000,00 or 0,00) — avoid greedy [\d\s,]+ swallowing neighbours */
  const zaAmt = String.raw`\d{1,3}(?:\s\d{3})*,\d{2}`;
  const re = new RegExp(
    `GROSS\\s+TOTAL\\s+DISCOUNT\\s+SUBTOTAL\\s+VAT\\s+CONTRACT\\s+TOTAL\\s+(${zaAmt})\\s+(${zaAmt})\\s+(${zaAmt})\\s+(${zaAmt})\\s+R?\\s*(${zaAmt})`,
    "gi",
  );
  let last = null;
  let m;
  while ((m = re.exec(one)) !== null) last = m;
  if (!last) return { gross: "", net: "", vat: "", total: "" };
  const fmt = (tok) => parseAmount(normalizeDstvOrderConfirmationMoney(tok));
  return {
    gross: fmt(last[1]),
    net: fmt(last[3]),
    vat: fmt(last[4]),
    total: fmt(last[5]),
  };
}

function parseDstvOrderConfirmationPrintedDate(flat) {
  const one = String(flat || "").replace(/\s+/g, " ");
  const m = one.match(
    /\bPRINTED\s+DATE\s+(\d{4})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{1,2})\b/i,
  );
  if (!m) return "";
  const y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10);
  const d = parseInt(m[3], 10);
  const dt = new Date(y, mo - 1, d);
  return Number.isNaN(dt.getTime()) ? "" : formatDocDate(dt);
}

function extractDstvOrderConfirmationHoldingCompany(flat) {
  const one = String(flat || "").replace(/\s+/g, " ");
  const m = one.match(
    /\bHOLDING\s+COMPANY\s+(.+?)(?=\s+YAH\s|\s+Channel\s+Date\b)/i,
  );
  return m ? cleanLabelValue(m[1]).slice(0, 80) : "";
}

function extractDstvOrderConfirmationBrand(flat) {
  const one = String(flat || "").replace(/\s+/g, " ");
  const m = one.match(
    /\b([A-Z][A-Za-z0-9&]{1,22})\s+PRINTED\s+DATE\s+\d{4}\/\d{1,2}\/\d{1,2}\b/,
  );
  if (!m) return "";
  const w = cleanLabelValue(m[1]);
  if (/^(delia|faith|mzikazi|quintee)\b/i.test(w)) return "";
  if (w.length < 2 || w.length > 40) return "";
  return w;
}

function extractDstvOrderConfirmationCampaign(flat) {
  const one = String(flat || "").replace(/\s+/g, " ");
  const parts = [];
  const portal = one.match(
    /\bPORTAL\s+NUMBER\s+(.+?)\s+\d{4}\/\d{2}\/\d{2}\s+\d{4}\/\d{2}\/\d{2}\b/i,
  );
  if (portal) {
    let seg = cleanLabelValue(portal[1]).slice(0, 72);
    seg = seg.split(/\s+(?=CELLULAR|CAMPAIGN\s+ATTN|:)/i)[0].trim();
    if (seg.length >= 4) parts.push(seg);
  }
  const yah = one.match(/\b(YAH\s+\d+WK\s+PACKAGE)\b/i);
  if (yah) parts.push(yah[1]);
  return [...new Set(parts)].filter(Boolean).join(" — ").slice(0, 96);
}

/** Capture invoice number only from explicit INVOICE NO (never “TAX INVOICE …”) */
const DSTV_INVOICE_NO_PATTERNS = [
  /\bINVOICE\s*N[O0]\.?\s*[:\s#.]*(\d{5,14})\b/i,
  /\bInvoice\s+No\.?\s*[:\s#.]*(\d{5,14})\b/i,
  /\bINVOICENO\.?\s*[:\s#.]*(\d{5,14})\b/i,
  /\bINVO[0O1lI]CE\s*N[O0]\.?\s*[:\s#.]*(\d{5,14})\b/i,
  /** Label split across glued spans — “INV OICE NO” etc. */
  /\bIN\s*VO\s*ICE\s*N\s*[O0]\.?\s*[:\s#.]*(\d{5,14})\b/i,
  /\bIN\s*VOICE\s*N\s*[O0]\.?\s*[:\s#.]*(\d{5,14})\b/i,
];

/** REF. NO. cell — numeric leading-zero ref or text (e.g. campaign ref line). */
function extractDstvRefNoField(flat) {
  const one = String(flat || "").replace(/\s+/g, " ").trim();
  let m = one.match(/\bREF\.?\s*NO\.?\s*[:\s]*(0000\d{4,14})\b/i);
  if (m) return cleanLabelValue(m[1]);
  m = one.match(/\bREF\.?\s*NO\.?\s*[:\s]*([\d]{7,14})\b/);
  if (m) return cleanLabelValue(m[1]);
  m = one.match(
    /\bREF\.?\s*NO\.?\s*[:\s]*([A-Za-z0-9][A-Za-z0-9\s&./-]{1,72}?)(?=\s+(?:INVOICE\s+DATE|INVOICE\s+NO\.?|PRODUCT|START\s+DATE|VAT\s+REG|CAMPAIGN)\b)/i,
  );
  return m ? cleanLabelValue(m[1]) : "";
}

/** PO NUMBER cell (often BO reference). */
function extractDstvPoNumberField(flat) {
  const one = String(flat || "").replace(/\s+/g, " ").trim();
  let m = one.match(
    /\bPO\s*(?:NUMBER|NO\.?)\s*[:\s]*([A-Za-z0-9./\s-]{2,72}?)(?=\s+(?:PORTAL\s+NUMBER|PROGRAM|DATE\s+DAY|ORD\b|LEN\s+COPY|CATEGORY|REF\.?\s*NO)\b|\s+MMS\b|\s+HOLDING\s+COMPANY\b)/i,
  );
  if (!m)
    m = one.match(
      /\bPO\s*(?:NUMBER|NO\.?)\s*[:\s]*([^\s]{2,46}(?:\s+[^\s]+){0,4}?)(?=\s+PORTAL\b)/i,
    );
  return m ? cleanLabelValue(m[1]) : "";
}

/**
 * Booking order vs PO from REF + PO grid cells:
 * — If exactly one field looks like a BO booking ref (digits + BO token) → that value is booking; PO gets the other when it has digits.
 * — Else if neither does but exactly one field has digits → duplicate that value into booking + PO.
 * — Else if both have digits (no BO-style ref) → booking = REF, PO = PO.
 * — Else empty.
 */
/** REF/PO cell looks like a booking reference (digits + explicit BO token, not e.g. “BOOKING”). */
function dstvCellHasBookingOrderMarker(s) {
  const t = String(s || "").trim();
  if (!/\d/.test(t)) return false;
  return (
    /\bBO[\s_/:-]*[\d]/i.test(t) ||
    /\bBO\d/i.test(t) ||
    /\d[\d\s_-]*\s*BO\b/i.test(t)
  );
}

function mergeDstvBookingOrderAndPurchaseOrder(refRaw, poRaw) {
  const ref = cleanLabelValue(String(refRaw || "")).slice(0, 96).trim();
  const po = cleanLabelValue(String(poRaw || "")).slice(0, 96).trim();

  const refHasDig = /\d/.test(ref);
  const poHasDig = /\d/.test(po);
  const refBo = dstvCellHasBookingOrderMarker(ref);
  const poBo = dstvCellHasBookingOrderMarker(po);

  const clip = (s) => s.replace(/\s+/g, " ").trim().slice(0, 46);

  if (refBo && !poBo) {
    return {
      bookingOrderNo: clip(ref),
      purchaseOrderNumber: poHasDig ? clip(po) : "",
    };
  }
  if (poBo && !refBo) {
    return {
      bookingOrderNo: clip(po),
      purchaseOrderNumber: refHasDig ? clip(ref) : "",
    };
  }
  if (refBo && poBo) {
    return {
      bookingOrderNo: clip(ref),
      purchaseOrderNumber: clip(po),
    };
  }

  if (refHasDig && !poHasDig) {
    const v = clip(ref);
    return { bookingOrderNo: v, purchaseOrderNumber: v };
  }
  if (poHasDig && !refHasDig) {
    const v = clip(po);
    return { bookingOrderNo: v, purchaseOrderNumber: v };
  }
  if (refHasDig && poHasDig) {
    return { bookingOrderNo: clip(ref), purchaseOrderNumber: clip(po) };
  }

  return { bookingOrderNo: "", purchaseOrderNumber: "" };
}

/**
 * Some DStv exports list every grid label first, then every value (“column-major”).
 * Flat text looks like: ADVERTISER VAT REG … END DATE AUTOMARK TAX INVOICE … 23053632 3/31/2026 …
 */
function parseDstvColumnMajorMetadata(flat) {
  const one = String(flat || "").replace(/\s+/g, " ").trim();
  if (!/\bADVERTISER\s+VAT\s+REG\b/i.test(one)) return null;

  const seg =
    one.match(/\bEND DATE\s+(.+?)\s+MMS\s+Communications\s+South\b/i) ||
    one.match(/\bEND DATE\s+(.+?)\s+MMS\s+COMMUNICATIONS\b/i);
  if (!seg) return null;

  let block = seg[1].trim();
  /** VAT ref digits + optional REF text + 8-digit invoice # + invoice date — ref may sit flush against invoice with no spacer text */
  const hdr = block.match(
    /^(.+?)\s+TAX\s+INVOICE\s+(\d+)\s+(?:(.+?)\s+)?(\d{8})\s+(\d{1,2}\/\d{1,2}\/\d{4})\s+/i,
  );
  if (!hdr) return null;

  const brandRaw = hdr[1].trim();
  const invoiceNo = hdr[4];
  const dateSlash = hdr[5];
  const refChunk = String(hdr[3] || "").trim();

  const dc = dateSlash.split("/");
  const issued =
    dc.length >= 3
      ? dateFromSlashParts(dc[1], dc[0], dc[2])
      : null;
  const dateDocumentIssued =
    issued && !Number.isNaN(issued.getTime()) ? formatDocDate(issued) : "";

  const tail = block.slice(hdr[0].length).trim();
  let product = "";
  const pmLit = tail.match(/^(.+?)\s+LITHA\s+/i);
  if (pmLit) product = cleanLabelValue(pmLit[1]);
  else {
    const px = tail.match(/^(.+?)\s+(?=\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}\/\d{1,2}\/\d{4})/);
    if (px) product = cleanLabelValue(px[1]);
  }

  return {
    brand: brandRaw.length >= 2 ? cleanLabelValue(brandRaw) : "",
    product,
    invoiceNo,
    dateDocumentIssued,
    bookingRefText:
      refChunk.length >= 3 && refChunk.length < 140 ? cleanLabelValue(refChunk) : "",
  };
}

/** Grid label mis-read as client / product when PDF column order breaks */
function isDstvGridJunkCapture(s) {
  const t = cleanLabelValue(s);
  if (!t || t.length < 2) return true;
  if (/^vat(?:\s*reg)?(?:istration)?(?:\.?\s*no\.?)?$/i.test(t)) return true;
  if (/^vat\s*reg/i.test(t)) return true;
  if (/vat\s*reg\.?\s*no\.?/i.test(t)) return true;
  if (/^acc\.?\s*exec/i.test(t)) return true;
  if (/^portal\s*number$/i.test(t)) return true;
  if (/^po\s*number$/i.test(t)) return true;
  if (/^holding\s+company$/i.test(t)) return true;
  if (/^category$/i.test(t)) return true;
  return false;
}

/** Next grid / footer label after a metadata cell — PDF flatten often concatenates columns */
const DSTV_AFTER_ADVERTISER_STOP =
  /\s+(?=PRODUCT\b|CATEGORY\b|HOLDING\s+COMPANY\b|VAT\s+REG|ACC\.?\s*EXEC|REF\.?\s*NO|CAMPAIGN\s+NO|START\s+DATE|END\s+DATE|INVOICE\s+NO\.?|INVOICE\s+DATE|TAX\s+INVOICE|ADVERTISER\b|Gross\b|Discount\b|Net\b|Total\s+Inv)/i;

const DSTV_AFTER_PRODUCT_STOP =
  /\s+(?=CATEGORY\b|START\s+DATE|END\s+DATE|INVOICE\s+NO\.?|INVOICE\s+DATE|VAT\s+REG|ACC\.?\s*EXEC|REF\.?\s*NO|CAMPAIGN\s+NO|PORTAL|PO\s+NUMBER|TAX\s+INVOICE|ADVERTISER\b|HOLDING\s+COMPANY\b|Gross\b|Discount\b|Net\b|Total\s+Inv|Building\b|Silver\b|Office\s+Park|DESCRIPTION\b|\d{1,2}\/\d{1,2}\/\d{4})/i;

function truncateDstvGridBleed(seg) {
  let s = cleanLabelValue(seg);
  if (!s) return "";
  const cuts = [
    /\s+\d{1,2}\/\d{1,2}\/\d{4}\b/,
    /\s+\d{8,}\b/,
    /\s+START\s+DATE\b/i,
    /\s+END\s+DATE\b/i,
    /\s+TAX\s+INVOICE\b/i,
    /\s+VAT\s+REG/i,
    /\s+ACC\.?\s*EXEC/i,
    /\s+INVOICE\s+NO/i,
    /\s+INVOICE\s+DATE/i,
    /\s+Building\s+\d/i,
    /\s+Silver\s+Stream/i,
    /\s+Office\s+Park/i,
  ];
  for (const r of cuts) {
    const i = s.search(r);
    if (i >= 12) s = s.slice(0, i).trim();
    else if (i >= 0 && i < 12) s = s.slice(0, i).trim();
  }
  let words = s.split(/\s+/).filter(Boolean);
  while (
    words.length &&
    /^(START|END|DATE|ACC\.?|EXEC\.?|VAT|REG\.?|NO\.?)$/i.test(words[0])
  )
    words.shift();

  const out = [];
  for (let wi = 0; wi < words.length; wi++) {
    const w = words[wi];
    if (/^\d{1,2}\/\d/.test(w)) break;
    if (/^\d{6,}$/.test(w)) break;
    if (
      /^(START|END|DATE|INVOICE|VAT|TAX|ACC|REF|CAMPAIGN|BUILDING|SILVER|OFFICE|CATEGORY|HOLDING|PRODUCT|PORTAL|PO)$/i.test(
        w,
      )
    )
      break;
    if (/^tax$/i.test(w) && /^invoice$/i.test(words[wi + 1] || "")) break;
    out.push(w);
    if (out.length >= 8) break;
    if (out.join(" ").length > 88) break;
  }
  s = out.join(" ").trim();
  return s.slice(0, 96).trim();
}

function isPlausibleDstvAdvertiser(s) {
  const t = truncateDstvGridBleed(s);
  if (!t || t.length < 2 || t.length > 72) return false;
  if (/^(START|END|INVOICE|VAT|TAX|ACC|REF|CAMPAIGN|PRODUCT)\b/i.test(t)) return false;
  if (isDstvGridJunkCapture(t)) return false;
  return true;
}

function isPlausibleDstvProduct(s) {
  const t = truncateDstvGridBleed(s);
  if (!t || t.length < 3 || t.length > 96) return false;
  if (/^(START|END|INVOICE|VAT|TAX|ACC|REF|CAMPAIGN|ADVERTISER)\b/i.test(t)) return false;
  if (isDstvGridJunkCapture(t)) return false;
  const digits = (t.match(/\d/g) || []).length;
  if (digits > t.length * 0.35) return false;
  return true;
}

/** Holding company grid cell → Client column (Advertiser maps to Brand). */
function extractDstvHoldingCompanyName(flat, lines) {
  const STOP =
    /\s+(?=ADVERTISER\b|PRODUCT\b|CATEGORY\b|VAT\s+REG|START\s+DATE|END\s+DATE|TAX\s+INVOICE|INVOICE\s+NO|INVOICE\s+DATE|HOLDING\s+COMPANY\b)/i;
  const reHold = /\bHOLDING\s+COMPANY\s*[:\s]*/gi;
  let hm;
  while ((hm = reHold.exec(flat)) !== null) {
    const start = hm.index + hm[0].length;
    const tailRaw = flat.slice(start);
    const cut = tailRaw.search(STOP);
    let seg = (cut >= 0 ? tailRaw.slice(0, cut) : tailRaw).trim();
    seg = seg.replace(/\s+/g, " ").slice(0, 120).trim();
    if (
      seg.length >= 2 &&
      !/^vat\s*reg/i.test(seg) &&
      !/^advertiser\b/i.test(seg)
    )
      return cleanLabelValue(seg);
  }
  const row = flat.match(
    /\bHOLDING\s+COMPANY\s*[:\s]+\s*(.+?)\s+(?=ADVERTISER\b)/i,
  );
  if (row && row[1])
    return cleanLabelValue(row[1].replace(/\s+/g, " ").trim().slice(0, 120));
  const joined = lines.map((l) => l.replace(/\s+/g, " ").trim()).filter(Boolean);
  const idx = joined.findIndex((l) => /\bHOLDING\s+COMPANY\b/i.test(l));
  if (idx >= 0) {
    let seg = joined[idx].replace(/^.*?HOLDING\s+COMPANY\s*[:\s]*/i, "").trim();
    if (!seg) seg = joined[idx + 1] || "";
    seg = seg.split(/\s+(?=ADVERTISER\b|PRODUCT\b)/i)[0];
    if (seg.length >= 2 && !/^vat\s*reg/i.test(seg))
      return cleanLabelValue(seg);
  }
  return "";
}

/**
 * Advertiser — slice after each ADVERTISER label until next grid token (handles column scramble).
 * @param {string} flat
 * @param {string[]} lines
 */
function extractDstvAdvertiserName(flat, lines) {
  /** Do not use “ADVERTISER … PRODUCT” across the whole flat string — the first PRODUCT may be far below and swallows CATEGORY / PO / PORTAL rows. */

  const reAdv = /\bADVERTISER\s*[:\s]*/gi;
  let mm;
  while ((mm = reAdv.exec(flat)) !== null) {
    const start = mm.index + mm[0].length;
    const tail = flat.slice(start);
    const cut = tail.search(DSTV_AFTER_ADVERTISER_STOP);
    let seg = (cut >= 0 ? tail.slice(0, cut) : tail).trim();
    seg = truncateDstvGridBleed(seg);
    if (isPlausibleDstvAdvertiser(seg)) return cleanLabelValue(seg);
  }

  const narrow = flat.match(
    /\bADVERTISER\s*[:\s]+\s*([A-Za-z0-9][A-Za-z0-9\s&.()-]{0,96}?)(?=\s+\bPRODUCT\b)/i,
  );
  if (narrow && narrow[1] && !isDstvGridJunkCapture(narrow[1]))
    return cleanLabelValue(truncateDstvGridBleed(narrow[1]));

  const joined = lines.map((l) => l.replace(/\s+/g, " ").trim()).filter(Boolean);
  const idx = joined.findIndex((l) => /\bADVERTISER\b/i.test(l));
  if (idx >= 0) {
    let seg = joined[idx].replace(/^.*?ADVERTISER\s*[:\s]*/i, "").trim();
    if (!seg || isDstvGridJunkCapture(seg)) seg = joined[idx + 1] || "";
    seg = seg.split(/\s+(?=PRODUCT\b|CATEGORY\b)/i)[0];
    seg = truncateDstvGridBleed(seg);
    if (seg && !isDstvGridJunkCapture(seg)) return cleanLabelValue(seg);
  }

  return "";
}

/**
 * Campaign / product — slice after PRODUCT until next label (never greedy-to-EOF).
 * @param {string} flat
 * @param {string[]} lines
 */
function extractDstvProductName(flat, lines) {
  const catRow = flat.match(/\bPRODUCT\s*[:\s]+\s*(.+?)\s+\bCATEGORY\b/i);
  if (catRow && catRow[1] && !isDstvGridJunkCapture(catRow[1]))
    return cleanLabelValue(truncateDstvGridBleed(catRow[1]));

  const rePr = /\bPRODUCT\s*[:\s]*/gi;
  let pm;
  while ((pm = rePr.exec(flat)) !== null) {
    const start = pm.index + pm[0].length;
    const tailRaw = flat.slice(start);
    const trimmed = tailRaw.trimStart();
    /** Cell empty in export — next tokens are another column’s labels */
    if (
      /^(START\s+DATE|END\s+DATE|INVOICE\s+NO\.?|INVOICE\s+DATE|VAT\s+REG|TAX\s+INVOICE)\b/i.test(
        trimmed,
      )
    )
      continue;
    const cut = tailRaw.search(DSTV_AFTER_PRODUCT_STOP);
    let seg = (cut >= 0 ? tailRaw.slice(0, cut) : tailRaw).trim();
    seg = truncateDstvGridBleed(seg);
    if (isPlausibleDstvProduct(seg)) return cleanLabelValue(seg);
  }

  const joined = lines.map((l) => l.replace(/\s+/g, " ").trim()).filter(Boolean);
  const pidx = joined.findIndex((l) => /\bPRODUCT\b/i.test(l));
  if (pidx >= 0) {
    let seg = joined[pidx].replace(/^.*?PRODUCT\s*[:\s]*/i, "").trim();
    if (!seg || isDstvGridJunkCapture(seg)) seg = joined[pidx + 1] || "";
    seg = seg.split(/\s+(?=CATEGORY\b|VAT\s+REG\b|START\s+DATE\b)/i)[0];
    seg = truncateDstvGridBleed(seg);
    if (seg && isPlausibleDstvProduct(seg)) return cleanLabelValue(seg);
  }
  return "";
}

/** When regex misses (split spans), digits following INVOICE NO label within a short window */
function extractDstvInvoiceNoLoose(flat, lines) {
  let hit = matchFirst(flat, DSTV_INVOICE_NO_PATTERNS);
  if (hit) return hit.trim();
  const joined = lines.map((l) => l.replace(/\s+/g, " ").trim()).join(" ");
  hit = matchFirst(joined, DSTV_INVOICE_NO_PATTERNS);
  if (hit) return hit.trim();

  let idx = -1;
  while ((idx = flat.toUpperCase().indexOf("INVOICE", idx + 1)) >= 0) {
    const slice = flat.slice(idx, idx + 72);
    if (!/\bINVOICE\s*N[O0]/i.test(slice)) continue;
    const m = slice.match(/(\d{5,14})\b/);
    if (m && m[1]) return m[1];
  }

  /** Last resort: DStv invoice numbers often 23xxxxxx (8 digits); use last match if unique */
  let last23 = "";
  let m23;
  const re23 = /\b(23\d{6})\b/g;
  while ((m23 = re23.exec(flat)) !== null) last23 = m23[1];
  const all23 = [...flat.matchAll(/\b(23\d{6})\b/g)].map((x) => x[1]);
  const uniq = [...new Set(all23)];
  if (uniq.length === 1) return uniq[0];
  if (last23 && uniq.length <= 3) return last23;
  return "";
}

/** Last monetary capture for a label (footer summary is usually the last occurrence). */
function lastMoneyForLabel(flatOneLine, labelRegex) {
  const flat = String(flatOneLine || "");
  const re = new RegExp(labelRegex.source + String.raw`\s*[:\s]*R?\s*([\d,]+\.\d{2})`, "gi");
  let last = "";
  let m;
  while ((m = re.exec(flat)) !== null) last = m[1];
  return last;
}

/** Last Gross…Total Inv. Amount row — `.match()` would take an earlier duplicate block and skew Gross/Net. */
function matchLastDstvFooterSummaryRow(flat) {
  const one = String(flat || "").replace(/\s+/g, " ");
  const re =
    /Gross\s*[:\s]*R?\s*([\d,]+\.\d{2})\s+Discount\s*[:\s]*R?\s*([\d,]+\.\d{2})\s+Net\s*[:\s]*R?\s*([\d,]+\.\d{2})\s+VAT\s*[:\s]*R?\s*([\d,]+\.\d{2})\s+Total\s*Inv\.?\s*Amount\s*[:\s]*R?\s*([\d,]+\.\d{2})/gi;
  let last = null;
  let m;
  while ((m = re.exec(one)) !== null) last = m;
  return last;
}

/** INVOICE DATE as US M/D/Y (e.g. 3/31/2026). */
function parseDstvInvoiceDateUs(flat) {
  let m = flat.match(
    /\bINVOICE\s*DATE\s*[:\s]+(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\b/i,
  );
  if (!m)
    m = flat.match(/\bINVOICE\s*DATE\s*[:\s]+(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})\b/i);
  /** OCR: INVO1CE / tight spacing */
  if (!m)
    m = flat.match(
      /\bINVO[0O1lI]CE\s*DATE\s*[:\s]+(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\b/i,
    );
  /** Label glued or spacing odd — slash date shortly after INVOICE DATE substring */
  if (!m) {
    const ix = flat.search(/\bINVO[0O1lI]?CE\s*DATE\b/i);
    if (ix >= 0) {
      const slice = flat.slice(ix, ix + 72);
      m = slice.match(
        /(\d{1,2})\s*[\/\-\.]\s*(\d{1,2})\s*[\/\-\.]\s*(\d{2,4})\b/,
      );
    }
  }
  if (!m) return "";
  const mo = parseInt(m[1], 10);
  const d = parseInt(m[2], 10);
  let y = parseInt(m[3], 10);
  if (m[3].length === 2) y += y >= 70 ? 1900 : 2000;
  if (mo < 1 || mo > 12 || d < 1 || d > 31 || y < 2000 || y > 2100) return "";
  const dt = new Date(y, mo - 1, d);
  return Number.isNaN(dt.getTime()) ? "" : formatDocDate(dt);
}

/**
 * iPoint invoices: the programme name sits on its own line/span immediately **above**
 * "Booking Order No." (after the last cost line). PDF text often glues the last amount
 * (e.g. 7,260.00) directly before the campaign phrase.
 */
function extractIpointCampaignBeforeBookingOrder(flatOneLine) {
  const flat = String(flatOneLine || "").replace(/\s+/g, " ");
  const bm = flat.match(/\bBooking\s+Order\s*No\.?\b/i);
  if (!bm || bm.index == null) return "";
  const before = flat.slice(0, bm.index).trim();
  let win = before.slice(Math.max(0, before.length - 360)).trim();

  /** Drop trailing money from the last table row if glued to the campaign line */
  win = win.replace(/(?:[\d,]+\.\d{2}\s*)+$/g, "").trim();

  /** If the last mall line sticks to the end before campaign text, strip it */
  win = win.replace(
    /\s+Premium\s+Digital\s+Package\s+-\s+[A-Za-z0-9\s,'-]{3,85}$/i,
    "",
  ).trim();

  const exact = win.match(
    /\b(Absa\s+PBB\s+Transactional\s+Summer\s+\d{4})\s*$/i,
  );
  if (exact) return cleanLabelValue(exact[1]);

  const absaTail = win.match(
    /\b(Absa\s+[A-Za-z0-9][A-Za-z0-9\s,'&.-]{6,85})\s*$/i,
  );
  if (absaTail && !/Premium\s+Digital/i.test(absaTail[1])) {
    return cleanLabelValue(absaTail[1]);
  }

  const tailWords = win.match(
    /([A-Z][a-z]+(?:\s+[A-Z][a-z]+|\s+\d{4}){2,8})\s*$/i,
  );
  if (
    tailWords &&
    tailWords[1].length >= 12 &&
    !/Mall|Centre|digital\s+package/i.test(tailWords[1])
  ) {
    return cleanLabelValue(tailWords[1]);
  }

  return "";
}

/** iPoint: campaign above BO (preferred), then footer "Campaign:"; booking from BO / MPO reference */
function extractIpointCampaignBooking(flatOneLine) {
  const flat = String(flatOneLine || "").replace(/\s+/g, " ");
  let campaign = extractIpointCampaignBeforeBookingOrder(flat);
  if (!campaign || campaign.length < 12) {
    campaign = cleanLabelValue(
      matchFirst(flat, [
        /\bCampaign\s*[:\s]+(.+?)(?=\s+Booking\s+Order|\s+Subtotal\s*\(?Net\)?|\s+TOTAL\s+VAT|\s+TOTAL\s+ZAR|BALANCE\s+DUE|\bCompany\s+Registration\b|$)/i,
      ]) || "",
    );
  }
  let booking =
    matchFirst(flat, [
      /\bBooking\s+Order\s*No\.?\s*[:\s]*([^\n]{4,55}?)(?=\s+Subtotal|\s+TOTAL\s+VAT|\s+TOTAL\s+ZAR|\s+BALANCE|$)/i,
      /\bReference\s*[:\s]*(MPO\d{4,}\s*[\/\.]\s*[\d]+(?:\s*\([^)]*\))?)/i,
    ]) || "";
  booking = cleanLabelValue(booking);
  return { campaign, booking };
}

/**
 * Flattened PDF text: find PRODUCT NAME(S): … (may appear without CAMPAIGN DETAILS newline).
 * When several matches exist (line-item table vs campaign box), prefer slug-like / longest.
 */
function extractProductNameFromFlatVolt(flatOneLine) {
  const flat = String(flatOneLine || "").replace(/\s+/g, " ");
  const re =
    /\bPRODUCT\s+NAME\s*(?:\(S\)|S)?\s*:\s*([A-Za-z0-9][A-Za-z0-9_\s\-./]{5,400})/gi;
  const candidates = [];
  let m;
  while ((m = re.exec(flat)) !== null) {
    const c = trimCampaignProductCapture(m[1]);
    if (c.length < 8) continue;
    if (isInvoiceTableHeaderOrJunkCampaign(c)) continue;
    candidates.push(c);
  }
  if (!candidates.length) return "";
  candidates.sort((a, b) => {
    const au = /_/.test(a) ? 1 : 0;
    const bu = /_/.test(b) ? 1 : 0;
    if (bu !== au) return bu - au;
    return b.length - a.length;
  });
  return candidates[0];
}

/**
 * Volt / IO layouts: boxed "CAMPAIGN DETAILS" with CAMPAIGN NO + PRODUCT NAME(S).
 * Not all invoices include this — only populate when the section exists.
 */
function extractCampaignDetailsBlock(rawText) {
  const text = rawText.replace(/\r\n/g, "\n");
  const lower = text.toLowerCase();
  const labelIdx = lower.search(/\bcampaign\s*details\b/i);
  if (labelIdx < 0) return { productName: "", campaignNo: "" };

  const fromLabel = text.slice(labelIdx);
  /** Scan ahead of truncated block — PDF order often places line-item headers before PRODUCT NAME(S): */
  const scanWin = fromLabel.slice(0, Math.min(fromLabel.length, 14000));
  let productFromWideScan = trimCampaignProductCapture(
    matchFirst(scanWin, [
      /product\s*name\s*(?:\(s\)|s)?\s*:\s*\s*(?:\r?\n\s*)?([A-Za-z0-9][A-Za-z0-9_\s\-./]{5,400})/im,
    ]) || "",
  );

  let endRel = fromLabel.length;
  /**
   * Do NOT cut at bare "START DATE" — pdf text often emits the line-items header
   * (START DATE | PRODUCT CODE | …) before "CAMPAIGN START DATE", which truncated
   * the block and dropped CAMPAIGN NO / PRODUCT NAME(S).
   *
   * Do NOT cut at "AMOUNT VAT EXCLUDED" — that table header often appears *before*
   * PRODUCT NAME(S): in extraction order and would drop the campaign description.
   */
  const cutters = [
    /\n\s*PRODUCT\s+CODE\b/i,
    /\bSTART\s+DATE\s+PRODUCT\s+CODE\b/i,
    /\n\s*Banking\s+Details\b/i,
    /\bBanking\s+Details\b/i,
    /\bTOTAL\s+VAT\s+EXCLUDED\b/i,
    /\bSwift\s+Code\b/i,
  ];
  for (const re of cutters) {
    const m = fromLabel.search(re);
    if (m > 40 && m < endRel) endRel = m;
  }
  const block = fromLabel.slice(0, endRel);
  const flatBlock = block.replace(/\s+/g, " ");
  const lines = block.split(/\n/).map((l) => l.replace(/\s+/g, " ").trim()).filter(Boolean);

  let campaignNo = "";
  for (const line of lines) {
    const cm =
      line.match(/^\s*campaign\s*no\.?\s*[:\s]*(\d{4,12})\s*$/i) ||
      line.match(/\bcampaign\s*no\.?\s*[:\s]*(\d{4,12})\b/i) ||
      line.match(/\bcampaign\s*number\s*[:\s]*(\d{4,12})\b/i);
    if (cm?.[1]) {
      campaignNo = cm[1];
      break;
    }
  }
  if (!campaignNo) {
    campaignNo = matchFirst(flatBlock, [
      /\bcampaign\s*no\.?\s*[:\s]*(\d{4,12})\b/i,
      /\bcampaign\s*number\s*[:\s]*(\d{4,12})\b/i,
    ]);
  }

  /**
   * PDF row order sometimes drops labels: date line → bare campaign id → product text.
   * Accept a line that is only digits (campaign ref), between a start-date row and prose.
   */
  if (!campaignNo) {
    for (let i = 0; i < lines.length; i++) {
      const cur = lines[i].trim();
      if (!/^\d{4,12}$/.test(cur)) continue;
      if (/^20\d{2}$/.test(cur)) continue;
      const prev = lines[i - 1] || "";
      const next = lines[i + 1] || "";
      const prevLooksDate =
        /\d{1,2}\/\d{1,2}\/\d{4}/.test(prev) ||
        /campaign\s*start\s*date/i.test(prev);
      const nextLooksText = /[A-Za-z]{3,}/.test(next);
      if (prevLooksDate && nextLooksText) {
        campaignNo = cur;
        break;
      }
    }
  }

  let productName = "";
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (/^\s*campaign\s*start\s*date\b/i.test(line)) continue;
    const pm = line.match(/product\s*name\s*(?:\(s\)|s)?\s*[:\s]*(.*)$/i);
    if (!pm) continue;
    let val = (pm[1] || "").trim();
    /** PDFs often emit the label line with no text; value is on the following line */
    if (!val || val.length < 2) {
      for (let j = i + 1; j < lines.length; j++) {
        const cand = lines[j].trim();
        if (!cand) continue;
        if (/^\s*campaign\s*(?:start|end)\s*date\b/i.test(cand)) continue;
        if (/^\s*campaign\s*no\b/i.test(cand)) continue;
        if (/^\s*product\s*name/i.test(cand)) continue;
        if (isStandaloneSlashDateLine(cand)) continue;
        if (/^\d{4,12}$/.test(cand)) continue;
        val = cand;
        break;
      }
    }
    if (val.length >= 2) {
      productName = trimCampaignProductCapture(val);
      break;
    }
  }
  if (!productName) {
    productName = matchFirst(block, [
      /product\s*name\s*(?:\(s\)|s)?\s*[:\s]*\s*(?:\r?\n\s*)?([A-Za-z0-9][A-Za-z0-9_\s\-./]{5,400})/im,
      /product\s*name\s*(?:\(s\)|s)?\s*[:\s]*([^\n\r]{6,240})/i,
      /product\s*name\s*(?:\(s\)|s)?\s*[:\s]*(.+?)(?=\n\s*CAMPAIGN\s|\n\s*START\s+DATE|\n\s*$|$)/ims,
    ]);
    productName = trimCampaignProductCapture(productName);
  }
  /**
   * Prefer explicit PRODUCT NAME(S): from the wide window *before* unlabeled table-row
   * fallbacks — otherwise a line-item row can fill Campaign and hide the real slug.
   */
  if (!productName && productFromWideScan) {
    productName = productFromWideScan;
  }
  if (productName) {
    productName = cleanLabelValue(
      productName.split(
        /\d{1,3}(?:,\d{3})*\.\d{2}\s+|Banking\s+Details|Bidvest|Swift\s+Code|Branch\s+Code|VOLTAfrica|\bBranch\b/i,
      )[0],
    );
  }

  if (productName && /^\s*campaign\s*start\s*date\b/i.test(productName)) {
    productName = "";
  }

  /** Unlabeled product line: first substantive line that is not a date or bare id */
  if (!productName) {
    for (const line of lines) {
      const L = line.trim();
      if (!L || /^campaign\s*details\b/i.test(L)) continue;
      if (/^\s*campaign\s*(?:start|end)\s*date\b/i.test(L)) continue;
      if (/^\s*campaign\s*no\b/i.test(L)) continue;
      if (/^\s*product\s*name\s*(?:\(s\)|s)?\s*:?\s*$/i.test(L)) continue;
      if (isStandaloneSlashDateLine(L)) continue;
      if (/^\d{4,12}$/.test(L)) continue;
      if (/[A-Za-z]{4,}/.test(L) && L.length >= 8) {
        productName = cleanLabelValue(L);
        break;
      }
    }
  }

  /** Slug-style names (underscores), often alone on a line after labels */
  if (
    !productName &&
    lines.some((l) => /^[A-Za-z0-9][A-Za-z0-9_\-]{8,120}$/.test(l.trim()))
  ) {
    const slug = lines.find((l) => {
      const t = l.trim();
      return (
        /^[A-Za-z0-9][A-Za-z0-9_\-]{8,120}$/.test(t) &&
        t.includes("_") &&
        !/^\d+$/.test(t.replace(/_/g, ""))
      );
    });
    if (slug) productName = cleanLabelValue(slug.trim());
  }

  if (productName && isInvoiceTableHeaderOrJunkCampaign(productName)) {
    productName = "";
  }
  if (!productName) {
    const tailFlat = fromLabel
      .slice(0, Math.min(fromLabel.length, 14000))
      .replace(/\s+/g, " ");
    productName = extractProductNameFromFlatVolt(tailFlat);
  }

  return {
    productName,
    campaignNo,
  };
}

/** DD/MM/YYYY (South African style) */
function dateFromSlashParts(dayStr, monthStr, yearStr) {
  const d = parseInt(dayStr, 10);
  const mo = parseInt(monthStr, 10);
  let y = parseInt(yearStr, 10);
  if (y < 100) y += 2000;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  const dt = new Date(y, mo - 1, d);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

/**
 * Statement PDFs (e.g. iPoint activity) list activity lines like "Invoice # INV-5615".
 * Used to correlate with other attachments or with hyperlinked PDF targets.
 *
 * @param {string} normalizedPlainText output of {@link normalizePdfPlainText}
 * @returns {string[]} unique invoice document numbers (e.g. INV-5615), uppercased
 */
export function extractStatementInvoiceReferences(normalizedPlainText) {
  const flat = String(normalizedPlainText || "").replace(/\s+/g, " ");
  const refs = new Set();
  const patterns = [
    /\bInvoice\s*#\s*(INV-\d{3,8})\b/gi,
    /\bPayment\s+on\s+Invoice\s*#\s*(INV-\d{3,8})\b/gi,
  ];
  for (const re of patterns) {
    let m;
    re.lastIndex = 0;
    while ((m = re.exec(flat))) {
      refs.add(String(m[1]).toUpperCase());
    }
  }
  /** Activity rows: all INV tokens (covers text glued oddly or missing "Invoice #" prefix) */
  const loose = /\bINV-\d{4,8}\b/gi;
  let lm;
  while ((lm = loose.exec(flat))) {
    refs.add(lm[0].toUpperCase());
  }
  return [...refs];
}

/**
 * @param {string} rawText
 * @returns {Record<string, string>}
 */
export function extractInvoiceFields(rawText) {
  const text = normalizePdfPlainText(rawText);
  const flat = text.replace(/\s+/g, " ");
  const lines = buildExtractionLines(text);

  let documentType = "";
  if (/statement\s*[-–]\s*activity/i.test(flat) || /\bSTATEMENT\b.*\bActivity\b/i.test(flat)) {
    documentType = "Statement";
  } else if (isDstvOrderConfirmationDocument(flat)) {
    documentType = "Order Confirmation";
  } else if (/tax\s*invoice/i.test(flat)) documentType = "Tax Invoice";
  else if (/credit\s*note/i.test(flat)) documentType = "Credit Note";
  else if (/pro[- ]?forma/i.test(flat)) documentType = "Pro forma invoice";
  else if (/invoice/i.test(flat)) documentType = "Invoice";

  /**
   * DStv / grid layouts: doc no is always INVOICE NO. (digits). Must run before generic patterns —
   * especially avoid matching `INVOICE` inside `TAX INVOICE` + glued VAT reg digits (479…).
   */
  let documentNo = "";
  if (isDstvOrderConfirmationDocument(flat)) {
    documentNo = "N/A";
  } else if (isDstvMediaSalesDocument(flat)) {
    documentNo = extractDstvInvoiceNoLoose(flat, lines) || "";
  }

  /** Invoice ID# IN490561, iPoint INV-, numeric-only inv # — before loose patterns */
  if (!documentNo) {
    documentNo = matchFirst(flat, [
      /\b(INV-\d{3,8})\b/i,
      /\bInvoice\s*#\s*(INV-\d{3,8})\b/i,
      /\bPayment\s+on\s+Invoice\s*#\s*(INV-\d{3,8})\b/i,
      /** Not `TAX INVOICE` / `tax invoice` + digits (title glued to VAT reg in PDF order) */
      /(?<![Tt][Aa][Xx]\s)\bINVOICE\s+(?!DATE\b)(?:[#:]?\s*)?(\d{3,12})\b/i,
      /invoice\s*ID\s*[#:]?\s*([A-Z]{1,4}\d[\w\-]*)/i,
      /invoice\s*(?:number|nr\.?|no\.?)\s*[:\s#]*([A-Z]{1,4}\d[\w\-]*)/i,
      /invoice\s+number\s*[:\s]*([A-Z0-9][A-Z0-9\-/_]*)/i,
      /inv\.?\s*#\s*([A-Z0-9\-/_]+)/i,
      /(?:doc|document)\s*(?:number|no\.?)\s*[:\s]*([A-Z0-9][A-Z0-9\-/_]*)/i,
    ]);
  }

  let dateDocumentIssued = "";

  /** Prefer grid INVOICE DATE before generic slash-date scan (avoids START DATE winning in flat text order) */
  if (isDstvMediaSalesDocument(flat)) {
    const dstvEarlyDate = parseDstvInvoiceDateUs(flat);
    if (dstvEarlyDate) dateDocumentIssued = dstvEarlyDate;
  }

  if (!dateDocumentIssued) {
    const slashInv = flat.match(
      /invoice\s*date\s*[:\s]*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/i,
    );
    if (slashInv) {
      const dt = dateFromSlashParts(slashInv[1], slashInv[2], slashInv[3]);
      if (dt) dateDocumentIssued = formatDocDate(dt);
    }
  }

  if (!dateDocumentIssued) {
    const slashAny = flat.match(/\b(\d{1,2})[\/\-](\d{1,2})[\/\-](20\d{2})\b/);
    if (slashAny) {
      const dt = dateFromSlashParts(slashAny[1], slashAny[2], slashAny[3]);
      if (dt) dateDocumentIssued = formatDocDate(dt);
    }
  }

  if (!dateDocumentIssued) {
    const datePatterns = [
      /(?:invoice\s*)?date\s*[:\s]*(\d{1,2})[\s\-/]([A-Za-z]{3,})[\s\-/](\d{2,4})/i,
      /\b(\d{1,2})[\s\-/]([A-Za-z]{3,})[\s\-/](\d{4})\b/,
      /\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})\b/i,
    ];
    for (const re of datePatterns) {
      const m = flat.match(re);
      if (m) {
        const d = parseInt(m[1], 10);
        const monthStr = m[2];
        let y = parseInt(m[3], 10);
        if (y < 100) y += 2000;
        const mi = MONTHS.findIndex((x) =>
          monthStr.toLowerCase().startsWith(x.toLowerCase().slice(0, 3)),
        );
        if (mi >= 0 && d >= 1 && d <= 31) {
          const dt = new Date(y, mi, d);
          dateDocumentIssued = formatDocDate(dt);
          break;
        }
      }
    }
  }

  /** Prefer line-based VAT summary; Volt triple-label scan; then generic flat scan. */
  const lineVat = extractVatSummaryFromLines(lines);
  const voltTriple = extractVoltSummaryTriple(flat);
  const flatVat = extractVatSummaryFromFlat(flat);
  let vatPick = {
    excluded: lineVat.excluded || voltTriple.excluded || flatVat.excluded,
    vatTax: lineVat.vatTax || voltTriple.vatTax || flatVat.vatTax,
    included: lineVat.included || voltTriple.included || flatVat.included,
  };

  if (
    (!vatPick.vatTax || !vatPick.included || !vatPick.excluded) &&
    (/volt\.?africa|voltafrica|total\s*v\.?a\.?t\.?\s*excluded/i.test(flat) ||
      /tax\s*invoice/i.test(flat))
  ) {
    const inferred = inferVoltFooterTriplet(flat);
    if (!vatPick.excluded && inferred.excluded) vatPick.excluded = inferred.excluded;
    if (!vatPick.vatTax && inferred.vatTax) vatPick.vatTax = inferred.vatTax;
    if (!vatPick.included && inferred.included) vatPick.included = inferred.included;
  }

  /** Volt / SA marketing — flat-string fallbacks */
  let grossNetFromExcluded =
    vatPick.excluded ||
    matchFirst(flat, [
      /total\s*v\.?a\.?t\.?\s*excluded\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
      /total\s*v\.?a\.?t\.?\s*excl(?:uded)?\.?\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
      /amount\s*v\.?a\.?t\.?\s*excluded\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
      /amount\s*v\.?a\.?t\.?\s*excl(?:\.|uded)?\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
    ]);

  let vatOnlyTotal =
    vatPick.vatTax ||
    matchFirst(flat, [
      /\bTOTAL\s+VAT\s+(?!EXCLUDED\b)(?!INCLUDED\b)\s*:?\s*R?\s*([\d\s,]+\.\d{2})\b/i,
      /\btotal\s+v\.?a\.?t\.?\s+(?!EXCLUDED\b)(?!INCLUDED\b)\s*R?\s*([\d\s,]+\.?\d*)\b/i,
    ]);

  /** TOTAL VAT INCLUDED = payable total. Never prefer Balance/Amount due when this label exists. */
  let totalAmtInclusive =
    vatPick.included ||
    matchFirst(flat, [
      /\bTOTAL\s+VAT\s+INCLUDED\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
      /total\s*v\.?a\.?t\.?\s*included\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
      /total\s+incl(?:uding)?\.?\s*v\.?a\.?t\.?\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
    ]);

  const hasVoltIncludedLabel = /\bTOTAL\s+VAT\s+INCLUDED\b/i.test(flat);
  if (!totalAmtInclusive && !hasVoltIncludedLabel) {
    totalAmtInclusive = matchFirst(flat, [
      /balance\s*due\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
      /amount\s*due\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
    ]);
  }

  const totalAmtLoose = matchFirst(flat, [
    /(?:^|[^\w])(?:grand\s*total|invoice\s*total)\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
  ]);

  const subTotal = matchFirst(flat, [
    /sub[-\s]?total\s*\(\s*net\s*\)\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
    /sub[-\s]?total\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
  ]);

  const classicFooter = extractClassicInvoiceFooterAmounts(flat);

  /** Avoid "TERMS Net 30" — require currency-like NET or use tax-summary row */
  const netLine =
    classicFooter.netLabel ||
    matchFirst(flat, [
      /(?:amount\s*after\s*discount)\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
      /\bNET\b\s+(?:ZAR\s*)?([\d\s,]+\.\d{2})\b/i,
    ]);

  const grossLine = matchFirst(flat, [
    /(?:gross|rate\s*card)\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
  ]);

  let grossAmount = parseAmount(grossNetFromExcluded || grossLine || subTotal);
  let netAmount = parseAmount(netLine || grossNetFromExcluded || subTotal || grossLine);

  let vatAmount = parseAmount(vatOnlyTotal);
  if (!vatAmount) {
    const vatAlt = matchFirst(flat, [
      /v\.?a\.?t\.?\s*@\s*[\d.]+%\s*[:\s]*R?\s*([\d\s,]+\.?\d*)/i,
    ]);
    vatAmount = parseAmount(vatAlt);
  }

  /** Final total: prefer VAT inclusive; fallback to grand/invoice total. */
  let totalAmount = parseAmount(totalAmtInclusive);
  if (!totalAmount) totalAmount = parseAmount(totalAmtLoose);

  const parseMoneyNum = (s) => {
    if (!s) return NaN;
    const n = parseFloat(String(s).replace(/,/g, ""));
    return Number.isNaN(n) ? NaN : n;
  };

  let gNum = parseMoneyNum(grossAmount);
  let vNum = parseMoneyNum(vatAmount);
  let tNum = parseMoneyNum(totalAmount);

  /** Total wrongly equals VAT-excluded base while VAT exists — use gross + VAT. */
  if (!Number.isNaN(gNum) && !Number.isNaN(vNum) && vNum > 0.001) {
    if (
      !totalAmount ||
      (Math.abs(tNum - gNum) < 0.02 && Math.abs(tNum - (gNum + vNum)) > 0.05)
    ) {
      totalAmount = parseAmount(String(gNum + vNum));
      tNum = parseMoneyNum(totalAmount);
    }
  }

  /** Infer VAT from inclusive total − excluded when VAT line missing. */
  if (
    !Number.isNaN(gNum) &&
    !Number.isNaN(tNum) &&
    tNum > gNum + 0.02 &&
    (!vatAmount || Math.abs(vNum) < 0.001 || Math.abs(gNum + vNum - tNum) > 1)
  ) {
    const implied = tNum - gNum;
    if (implied > 0.01) vatAmount = parseAmount(String(implied));
  }

  if (!grossAmount && netAmount) grossAmount = netAmount;
  if (!netAmount && grossAmount) netAmount = grossAmount;

  /** Classic SA invoice (e.g. Food & Beverage Reporter): tax table "TAX … NET …", SUBTOTAL = ex-VAT */
  if (!isDstvOrderConfirmationDocument(flat)) {
    const cf = classicFooter;
    const cfNet = cf.netLabel ? parseMoneyFloat(parseAmount(cf.netLabel)) : NaN;
    const cfSub = cf.subtotal ? parseMoneyFloat(parseAmount(cf.subtotal)) : NaN;
    const cfTax = cf.tax ? parseMoneyFloat(parseAmount(cf.tax)) : NaN;
    const cfTot = cf.total ? parseMoneyFloat(parseAmount(cf.total)) : NaN;
    const nCur = netAmount ? parseMoneyFloat(parseAmount(netAmount)) : NaN;

    if (!Number.isNaN(cfNet) && cfNet > 50) {
      netAmount = parseAmount(cf.netLabel);
      grossAmount = parseAmount(cf.netLabel);
    } else if (!Number.isNaN(cfSub) && cfSub > 100) {
      if (
        Number.isNaN(nCur) ||
        nCur < cfSub * 0.35 ||
        (nCur < 500 && cfSub > 5000)
      ) {
        netAmount = parseAmount(cf.subtotal);
        grossAmount = parseAmount(cf.subtotal);
      }
    }
    if (!Number.isNaN(cfTax) && cfTax > 1 && !/total\s*v\.?a\.?t\.?\s*excluded/i.test(flat)) {
      vatAmount = parseAmount(cf.tax);
    }
    if (!Number.isNaN(cfTot) && cfTot > 1 && !hasVoltIncludedLabel) {
      totalAmount = parseAmount(cf.total);
    }
  }

  /** Supplier — branded marketing vendors */
  let supplierName = "";
  if (/volt\.africa/i.test(flat)) supplierName = "Volt.Africa";
  else if (/voltafrica/i.test(flat)) supplierName = "VOLTAfrica";
  else if (/food\s*(?:&|and)\s*beverage\s*reporter/i.test(flat)) {
    supplierName = "Food & Beverage Reporter (PTY) LTD";
  } else if (isDstvMediaSalesDocument(flat)) {
    supplierName = "DStv Media Sales (PTY) Ltd";
  } else if (isIpointDocument(flat)) {
    supplierName = "iPoint Media CC";
  } else {
    supplierName = matchFirst(flat, [
      /(?:from|supplier|vendor)\s*[:\s]*([^\n]{3,80}?)(?:\s+(?:vat|reg|tel|phone|address))/i,
    ]);
  }
  if (!supplierName) {
    supplierName = cleanLabelValue(
      matchFirst(text.slice(0, 900), [
        /^([A-Z][A-Z0-9\s&,.'-]{6,72}\(PTY\)\s+LTD)/m,
        /^(FOOD\s+(?:AND|&)\s+BEVERAGE\s+REPORTER\s*\(PTY\)\s+LTD)/im,
      ]),
    );
  }
  if (!supplierName) {
    const headerCompany = matchFirst(flat, [
      /([A-Z][a-z]+(?:community|media|digital|publish|group|pty|ltd)[A-Za-z0-9\s&,.'()-]{0,60})/i,
    ]);
    if (headerCompany) supplierName = headerCompany.trim();
  }

  /** Client — Advertiser Name (campaign invoice) preferred over Invoice To agency */
  let clientName = "";
  if (!isDstvMediaSalesDocument(flat)) {
    clientName = cleanLabelValue(
      matchFirst(flat, [
        /bill\s*to\s*[:\s]*(.+?)(?=\s+(?:VAT|Tax\s+Registration|VAT\s+Reg|invoice|DATE|TERMS)|$)/i,
        /advertiser\s*name\s*[:\s]+([A-Za-z0-9\s&(),.'-]{5,140}?)(?=\s+brand\s*name)/i,
        /advertiser\s*name\s*[:\s]+([A-Za-z0-9\s&(),.'-]{5,140}?)(?=\s+company\s+(?:VAT|Reg))/i,
      ]),
    );
    if (!clientName) {
      clientName = cleanLabelValue(
        matchFirst(text.replace(/\r/g, "\n"), [
          /advertiser\s*name\s*[:\s]*(.+?)(?=\s*(?:brand\s*name|company\s*VAT|PO\s*number|payment\s*due|invoice\s*to|campaign\s*details|$))/ims,
          /advertiser\s*name\s*[:\s]*([^\n]{3,160})/i,
        ]),
      );
    }

    if (!clientName) {
      clientName = cleanLabelValue(
        matchFirst(flat, [
          /to\s*[:\s]+([A-Za-z0-9\s&,.'()-]{3,80}?)(?:\s+Vat\b|\s+Silver|\s+Tel|\s+Phone|\s+Email|\s+www\.)/i,
        ]),
      );
    }

    if (!clientName) {
      const toIdx = lines.findIndex((l) => /^invoice\s*to\s*[:\s]?/i.test(l) || /^to\s*[:\s]/i.test(l));
      if (toIdx >= 0) {
        const next = lines[toIdx + 1];
        const same = lines[toIdx].replace(/^invoice\s*to\s*[:\s]*/i, "").replace(/^to\s*[:\s]+/i, "").trim();
        clientName = cleanLabelValue(same.length > 2 ? same : next || "");
      }
    }
    if (!clientName) {
      clientName = cleanLabelValue(
        matchFirst(flat, [
          /(?:bill\s*to|sold\s*to|customer)\s*[:\s]*([A-Za-z0-9\s&,.'()-]{3,80}?)(?:\s*\n|\s{2,}|$)/i,
        ]),
      );
    }
  }

  if (isStatementPeriodGarbageField(supplierName)) supplierName = "";
  if (isStatementPeriodGarbageField(clientName)) clientName = "";

  /** Brand — stop before advertiser/PO bleed from flattened PDF text */
  let brandName = cleanLabelValue(
    matchFirst(flat, [/brand\s*name\s*[:\s]*([^\n]{2,80})/i]),
  );
  if (!brandName) {
    brandName = cleanLabelValue(matchFirst(flat, [/^brand\s*[:\s]*([^\n]{2,60})/im]));
  }
  if (brandName) {
    brandName = cleanLabelValue(
      brandName.split(
        /\bADVERTISER\b|\bPO\s*NUMBER\b|\bCompany\s+VAT\b|\(\s*s\s*\)\s*\d+\s*days|\bPayment\s+Terms\b|\bLess\s+\d+|\bVAT\s+Amount\b|\bBalance\s+Owing\b/i,
      )[0],
    );
  }

  /** Campaign — CAMPAIGN DETAILS section first (product name + campaign no), then fallbacks */
  const campaignDetails = extractCampaignDetailsBlock(text);

  let campaignName = cleanLabelValue(campaignDetails.productName || "");

  if (!campaignName || campaignName.length < 8) {
    const flatPn = extractProductNameFromFlatVolt(flat);
    if (flatPn) campaignName = flatPn;
  }

  if (!campaignName || campaignName.length < 12) {
    const boFromLine = extractCampaignFromBoLineItem(flat);
    if (
      boFromLine.length >= 15 &&
      !isInvoiceTableHeaderOrJunkCampaign(boFromLine)
    ) {
      campaignName = boFromLine;
    }
    if (!campaignName || campaignName.length < 12) {
      const boLine = matchFirst(flat, [
        /\b(A4\s+[^\n]{10,260}?\bBO\s*(?:No\.?|number|nr\.?)?\s*[:\s]*[\d\s]{7,14})/i,
        /\b((?:advertisement|editorial|digitorial)[^\n]{8,260}?\bBO\s*(?:No\.?|number|nr\.?)?\s*[:\s]*[\d\s]{7,14})/i,
      ]);
      const tidied = trimCampaignProductCapture(boLine);
      if (tidied.length >= 15 && !isInvoiceTableHeaderOrJunkCampaign(tidied)) {
        campaignName = tidied;
      }
    }
  }

  if (!campaignName || campaignName.length < 8) {
    campaignName = cleanLabelValue(
      matchFirst(text.replace(/\r/g, "\n"), [
        /product\s*name\s*(?:\(s\)|s)?\s*[:\s]*(.+?)(?=\s*campaign\s*start|\s*campaign\s*no|$)/ims,
        /product\s*name\s*(?:\(s\)|s)?\s*[:\s]*([^\n]{10,200})/i,
      ]),
    );
    if (campaignName && isInvoiceTableHeaderOrJunkCampaign(campaignName)) {
      campaignName = "";
    }
  }

  if (!campaignName || campaignName.length < 8) {
    const descGuess = cleanLabelValue(
      matchFirst(flat, [
        /description\s*[:\s]*([^\n]{15,200})/i,
      ]),
    );
    if (descGuess && !isInvoiceTableHeaderOrJunkCampaign(descGuess)) {
      campaignName = descGuess;
    }
  }

  if (!campaignName || campaignName.length < 10) {
    const descIdx = lines.findIndex((l) => {
      const L = l.trim();
      if (/^campaign\s*details\b/i.test(L)) return false;
      if (/^\s*campaign\s*start\s*date\b/i.test(L)) return false;
      return /description|(?:^|\s)details\b|advertising|digitorial|product\s*code/i.test(l);
    });
    if (descIdx >= 0 && lines[descIdx + 1]) {
      const cand = cleanLabelValue(lines[descIdx + 1]);
      if (
        !isCampaignStartDateField(cand) &&
        !isStandaloneSlashDateLine(cand) &&
        !isInvoiceTableHeaderOrJunkCampaign(cand)
      ) {
        campaignName = cand;
      }
    }
  }
  if (!campaignName || campaignName.length < 10) {
    const longLine = lines.find(
      (l) =>
        l.length > 35 &&
        /advert|campaign|website|digital|month|digitorial|phase|basalt|partnered/i.test(l) &&
        !/^\s*campaign\s*start\s*date\b/i.test(l) &&
        !isInvoiceTableHeaderOrJunkCampaign(l),
    );
    if (longLine) campaignName = cleanLabelValue(longLine);
  }

  if (
    campaignName &&
    /AMOUNT\s*VAT|Bidvest|Swift\s*:|Bank\s*Nr|Banking\s+Details/i.test(campaignName)
  ) {
    const idx = campaignName.search(
      /\bCitroen\b|\bPartnered\b|\bDigitorial\b|\bPhase\b|\bBasalt\b|\bCampaign\b|\bApril\b|\b\d{4}\b/i,
    );
    let cut = idx >= 0 ? campaignName.slice(idx) : campaignName;
    cut = cut.split(/\bBidvest\b|\bSwift\b|\bBank\s*Nr\b|\bBanking\b|\d{1,3}(?:,\d{3})*\.\d{2}/i)[0];
    if (/amount\s*v\.?a\.?t\.?\s*exclud/i.test(cut)) {
      cut = cut.split(/\bamount\s*v\.?a\.?t\.?\s*exclud/i)[0];
    }
    campaignName = cleanLabelValue(cut);
  }

  /** Reject start-date-only or bare invoice dates mistaken for campaign description */
  if (campaignName && isCampaignStartDateField(campaignName)) {
    campaignName = "";
  }
  if (campaignName && isStandaloneSlashDateLine(campaignName)) {
    campaignName = "";
  }
  if (campaignName && isInvoiceTableHeaderOrJunkCampaign(campaignName)) {
    campaignName = "";
  }

  /** PO — explicit PO Number label */
  let purchaseOrderNumber = cleanLabelValue(
    matchFirst(flat, [
      /PO\s*(?:number|no\.?)\s*[:\s]*(\d[\d\s]*)/i,
      /purchase\s*order\s*(?:number|no\.?|#)?\s*[:\s]*([A-Z0-9/\-_]+)/i,
      /\bPO\s*[:\s#]*(\d[\dA-Z/\-]*)/i,
    ]),
  );

  /** Booking order — explicit line, or "BO No" embedded in description */
  let bookingOrderNo = cleanLabelValue(
    matchFirst(flat, [
      /booking\s*order(?:\s*number|\s*no\.?)?\s*[:\s]*([A-Z0-9][A-Z0-9\/\-_]{2,40})/i,
      /\bBO\s*(?:number|no\.?|nr\.?|#)?\s*[:\s]*([\d\s]{7,14})/i,
    ]),
  );
  if (bookingOrderNo) {
    bookingOrderNo = bookingOrderNo.replace(/\s+/g, "").slice(0, 24);
  }

  /** Contract — only explicit contract fields (not account number) */
  let contractNumber = cleanLabelValue(
    matchFirst(flat, [
      /contract\s*(?:number|no\.?|ref)?\s*[:\s]*([A-Z0-9][A-Z0-9\-/_]{2,30})/i,
      /contract\s*ref(?:erence)?\s*[:\s]*([A-Z0-9][A-Z0-9\-/_]{2,30})/i,
    ]),
  );

  /** Campaign number / CAP — CAMPAIGN DETAILS first, then generic labels (not PO) */
  let campCampaignNo = cleanLabelValue(campaignDetails.campaignNo || "");
  if (!campCampaignNo) {
    campCampaignNo = cleanLabelValue(
      matchFirst(flat, [
        /campaign\s*no\.?\s*[:\s]*(\d{4,12})/i,
        /campaign\s*number\s*[:\s]*(\d{4,12})/i,
      ]),
    );
  }

  /** iPoint Media CC: footer labels; Absa programme client; statements may omit Campaign line */
  if (isIpointDocument(flat)) {
    supplierName = "iPoint Media CC";
    const ip = extractIpointCampaignBooking(flat);
    if (
      ip.campaign &&
      ip.campaign.length >= 12 &&
      !isInvoiceTableHeaderOrJunkCampaign(ip.campaign)
    ) {
      campaignName = ip.campaign;
    }
    if (ip.booking && ip.booking.replace(/\s/g, "").length >= 6) {
      bookingOrderNo = ip.booking
        .replace(/\s+/g, " ")
        .trim()
        .replace(/\s*\([^)]*\)\s*$/, "")
        .trim()
        .slice(0, 46);
    }
    if (/absa/i.test(flat) || /absa/i.test(campaignName || "")) {
      clientName = "Absa PPB";
    }
    if (
      (!campaignName || campaignName.length < 15) &&
      /\bAbsa\s+PBB\s+Transactional\s+Summer\s+2025\b/i.test(flat)
    ) {
      campaignName = "Absa PBB Transactional Summer 2025";
    }
    if (
      (!campaignName || campaignName.length < 15) &&
      /\bstatement\b/i.test(flat) &&
      /\bMPO\d+/i.test(flat) &&
      (/absa/i.test(flat) || /transactional/i.test(flat))
    ) {
      campaignName = "Absa PBB Transactional Summer 2025";
    }
    if (isStatementPeriodGarbageField(supplierName)) supplierName = "iPoint Media CC";
    if (isStatementPeriodGarbageField(clientName) && /absa/i.test(flat)) {
      clientName = "Absa PPB";
    }
  }

  /** DStv Media Sales (PTY) Ltd — tax invoice grid + footer, or order confirmation layout */
  if (isDstvMediaSalesDocument(flat)) {
    supplierName = "DStv Media Sales (PTY) Ltd";

    if (isDstvOrderConfirmationDocument(flat)) {
      documentType = "Order Confirmation";
      documentNo = "N/A";
      contractNumber = "";
      const printed = parseDstvOrderConfirmationPrintedDate(flat);
      if (printed) dateDocumentIssued = printed;
      const hold = extractDstvOrderConfirmationHoldingCompany(flat);
      if (hold) clientName = hold;
      const br = extractDstvOrderConfirmationBrand(flat);
      if (br) brandName = br;
      const camp = extractDstvOrderConfirmationCampaign(flat);
      if (camp) campaignName = camp;
      const attn = matchFirst(flat, [/\bCAMPAIGN\s+ATTN\s+(\d{4,8})\b/i]);
      if (attn) campCampaignNo = attn.trim();
      const ocTot = parseDstvOrderConfirmationFooterTotals(flat);
      if (ocTot.gross) grossAmount = ocTot.gross;
      if (ocTot.net) netAmount = ocTot.net;
      if (ocTot.vat) vatAmount = ocTot.vat;
      if (ocTot.total) totalAmount = ocTot.total;
    } else {
      /** PDF sometimes exports labels in one block, values after END DATE — see parseDstvColumnMajorMetadata */
      const cmMeta = parseDstvColumnMajorMetadata(flat);

      if (cmMeta?.invoiceNo) documentNo = String(cmMeta.invoiceNo).trim();
      else {
        const invNo = extractDstvInvoiceNoLoose(flat, lines);
        if (invNo) documentNo = invNo.trim();
      }

      const dstvIssued = parseDstvInvoiceDateUs(flat);
      if (dstvIssued) dateDocumentIssued = dstvIssued;
      else if (cmMeta?.dateDocumentIssued)
        dateDocumentIssued = cmMeta.dateDocumentIssued;

      const dstvHold = extractDstvHoldingCompanyName(flat, lines);
      clientName = dstvHold || "";

      if (cmMeta?.brand) brandName = cmMeta.brand;
      else {
        const dstvAdv = extractDstvAdvertiserName(flat, lines);
        brandName = dstvAdv || "";
      }

      if (cmMeta?.product) campaignName = cmMeta.product;
      else {
        const prod = extractDstvProductName(flat, lines);
        campaignName = prod || "";
      }

      const refRaw =
        extractDstvRefNoField(flat) ||
        (cmMeta?.bookingRefText ? String(cmMeta.bookingRefText).trim() : "");
      const poRaw = extractDstvPoNumberField(flat);
      const mergedBoPo = mergeDstvBookingOrderAndPurchaseOrder(refRaw, poRaw);
      bookingOrderNo = mergedBoPo.bookingOrderNo;
      purchaseOrderNumber = mergedBoPo.purchaseOrderNumber;

      const campNo = matchFirst(flat, [/\bCAMPAIGN\s*NO\.?\s*[:\s]*(\d{4,12})\b/i]);
      if (campNo) campCampaignNo = campNo.trim();

      /** Same layout can appear twice — first hit is often a mid-document row; footer summary is usually last. */
      const footRow = matchLastDstvFooterSummaryRow(flat);
      if (footRow) {
        grossAmount = parseAmount(footRow[1]);
        netAmount = parseAmount(footRow[3]);
        vatAmount = parseAmount(footRow[4]);
        totalAmount = parseAmount(footRow[5]);
      } else {
        const g = lastMoneyForLabel(flat, /\bGross\b/);
        const n = lastMoneyForLabel(flat, /\bNet\b/);
        const v = lastMoneyForLabel(flat, /\bVAT\b/);
        const t = lastMoneyForLabel(flat, /\bTotal\s*Inv\.?\s*Amount\b/);
        if (g) grossAmount = parseAmount(g);
        if (n) netAmount = parseAmount(n);
        if (v) vatAmount = parseAmount(v);
        if (t) totalAmount = parseAmount(t);
      }
    }
  }

  if (!documentNo) {
    const alt = flat.match(/\b([A-Z]{2}\d{5,9})\b/);
    if (alt) documentNo = alt[1];
  }

  const statementInvoiceRefs =
    documentType === "Statement" ? extractStatementInvoiceReferences(text) : [];

  return {
    dateDocumentIssued,
    documentType,
    documentNo,
    grossAmount: grossAmount || "",
    netAmount: netAmount || "",
    vatAmount: vatAmount || "",
    totalAmount: totalAmount || "",
    supplierName: supplierName || "",
    clientName: clientName || "",
    brandName: brandName || "",
    campaignName: campaignName || "",
    campCampaignNo: campCampaignNo || "",
    bookingOrderNo: bookingOrderNo || "",
    contractNumber: contractNumber || "",
    purchaseOrderNumber: purchaseOrderNumber || "",
    statementInvoiceRefs,
  };
}

/**
 * Parse iPoint-style activity rows: "Invoice # INV-xxxx" with line date and invoice amount
 * (first money after the INV token). Skips "Payment on Invoice # …" lines.
 *
 * @param {string} normalizedPlainText
 * @returns {{ invoiceNo: string, dateRaw: string, dateDocumentIssued: string, inclusiveTotal: number }[]}
 */
export function parseStatementTaxInvoiceActivityRows(normalizedPlainText) {
  const flat = String(normalizedPlainText || "").replace(/\s+/g, " ");
  const re = /Invoice\s*#\s*(INV-\d{4,8})/gi;
  const out = [];
  let m;
  while ((m = re.exec(flat)) !== null) {
    const idx = m.index;
    const beforeSlice = flat.slice(Math.max(0, idx - 72), idx);
    if (/payment\s+on\s+$/i.test(beforeSlice)) continue;

    const invoiceNo = m[1].toUpperCase();
    const dateSearch = flat.slice(Math.max(0, idx - 160), idx);
    const dm = dateSearch.match(
      /(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})\s*$/i,
    );
    const dateRaw = dm ? dm[1].trim() : "";

    const tail = flat.slice(idx + m[0].length, idx + m[0].length + 520);
    const monies = tail.match(MONEY_RE) || [];
    let amountNum = NaN;
    for (let k = 0; k < monies.length; k++) {
      const n = parseFloat(String(monies[k]).replace(/,/g, ""));
      if (!Number.isNaN(n) && n >= 1) {
        amountNum = n;
        break;
      }
    }
    if (Number.isNaN(amountNum)) continue;

    let dateDocumentIssued = "";
    if (dateRaw) {
      const dt = parseStatementActivityDateToDate(dateRaw);
      if (dt) dateDocumentIssued = formatDocDate(dt);
    }

    out.push({
      invoiceNo,
      dateRaw,
      dateDocumentIssued,
      inclusiveTotal: amountNum,
    });
  }
  return out;
}

function parseStatementActivityDateToDate(dateRaw) {
  const m = dateRaw
    .trim()
    .match(/^(\d{1,2})\s+([A-Za-z]{3})[a-z]*\s+(\d{4})$/i);
  if (!m) return null;
  const d = parseInt(m[1], 10);
  const mi = MONTHS.findIndex((x) =>
    m[2].toLowerCase().startsWith(x.toLowerCase().slice(0, 3)),
  );
  const y = parseInt(m[3], 10);
  if (mi < 0) return null;
  const dt = new Date(y, mi, d);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

/**
 * Pair activity rows with PDF hyperlink URLs by INV token in URL or aligned order.
 *
 * @param {{ invoiceNo: string }[]} activityRows
 * @param {string[]} urls
 * @returns {{ invoiceNo: string, externalUrl: string }[]}
 */
export function pairStatementLinesWithHyperlinkUrls(activityRows, urls) {
  const httpUrls = [...new Set(urls || [])].filter((u) =>
    /^https?:\/\//i.test(String(u)),
  );
  const used = new Set();
  return activityRows.map((row, i) => {
    let externalUrl = "";
    const bare = String(row.invoiceNo || "").replace(/^INV-/i, "");
    const byInv = httpUrls.find(
      (u) =>
        !used.has(u) &&
        bare.length >= 4 &&
        u.toUpperCase().includes(bare),
    );
    if (byInv) {
      used.add(byInv);
      externalUrl = byInv;
    } else if (httpUrls[i] && !used.has(httpUrls[i])) {
      externalUrl = httpUrls[i];
      used.add(httpUrls[i]);
    }
    return { ...row, externalUrl };
  });
}

/**
 * South Africa 15% VAT included in total: derive net/VAT; gross column = net (per statement synthesis).
 *
 * @param {number} totalInclusive
 */
export function zaInclusiveTotalToNetVatGrossStrings(totalInclusive) {
  const T = totalInclusive;
  const vat = Math.round(((T * 15) / 115) * 100) / 100;
  const net = Math.round((T - vat) * 100) / 100;
  const fmt = (n) =>
    n.toLocaleString("en-ZA", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  return {
    totalAmount: fmt(T),
    vatAmount: fmt(vat),
    netAmount: fmt(net),
    grossAmount: fmt(net),
  };
}
