/**
 * FlightBox Mail API
 * Node backend: gets inbox emails from proofofflightZA mailbox via Microsoft Graph.
 * Serves the dashboard and exposes /api/mail/inbox.
 */
import dotenv from "dotenv";
import express from "express";
import path from "path";
import crypto from "node:crypto";
import { Readable } from "node:stream";
import { fileURLToPath } from "url";
import {
  mergeRowsByAttachmentKey,
  readFolderRowsFromCache,
  writeFolderRowsToCache,
  maxEmailReceivedIso,
} from "./documentDetailsCache.js";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { PDFParse } from "pdf-parse";
import {
  extractInvoiceFields,
  isSabcAccountStatement,
  parseSabcEmbeddedTaxInvoiceDetails,
  parseSabcStatementOutstandingInvoiceLines,
  parseStatementTaxInvoiceActivityRows,
  pairStatementLinesWithHyperlinkUrls,
  zaInclusiveTotalToNetVatGrossStrings,
} from "./invoiceExtract.js";
import { writeGoogleExtractDebuggerFile } from "./googleExtractDebug.js";
import {
  extractHttpsUrlsFromPdfBuffer,
  extractInvNumbersFromPdfBinary,
} from "./pdfHyperlinks.js";
import { fetchPdfBufferResolvingViewerPages } from "./statementPdfFetch.js";
import multer from "multer";
import {
  readQueriesImport,
  writeQueriesImport,
  parseQueriesExcelBuffer,
  detectQueryMergeColumns,
} from "./queriesImport.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
// Load .env from project root (same as Python app) so one file works for both
dotenv.config({ path: path.join(__dirname, "..", ".env") });
const app = express();
const PORT = process.env.PORT || 3000;
/** Bind address (use 0.0.0.0 in Docker / cloud so the port is reachable). */
const LISTEN_HOST = process.env.LISTEN_HOST || "0.0.0.0";

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const TENANT_ID = process.env.TENANT_ID;
const MAILBOX_EMAIL = process.env.MAILBOX_EMAIL || "proofofflightZA@publicis.co.za";
const GRAPH_BASE = process.env.GRAPH_API_ENDPOINT || "https://graph.microsoft.com/v1.0";
const MAX_DOCUMENT_PDF_BYTES =
  parseInt(process.env.MAX_DOCUMENT_PDF_MB || "25", 10) * 1024 * 1024;

/** How many messages per folder to consider when building document-details (newest first). */
const DOCUMENT_DETAILS_MESSAGE_LIMIT_DEFAULT = Math.max(
  200,
  Math.min(
    10000,
    parseInt(process.env.DOCUMENT_DETAILS_FOLDER_MESSAGE_LIMIT || "2500", 10) || 2500,
  ),
);
const DOCUMENT_DETAILS_MESSAGE_LIMIT_MAX = 10000;

function effectiveDocumentDetailsLimit(raw) {
  const n = parseInt(String(raw || "").trim(), 10);
  if (!Number.isFinite(n) || n < 1) return DOCUMENT_DETAILS_MESSAGE_LIMIT_DEFAULT;
  return Math.min(Math.max(n, 50), DOCUMENT_DETAILS_MESSAGE_LIMIT_MAX);
}
/** Max PDF file attachments parsed per email (Graph + parse cost). Was hard-coded 5 — bulk packs need more. */
const MAX_PDFS_PER_MESSAGE = Math.max(
  1,
  Math.min(200, parseInt(process.env.MAX_PDFS_PER_MESSAGE || "40", 10)),
);
/** When a received-date range is active, allow up to this many PDFs per message (import “all” attachments in range). */
const MAX_PDFS_PER_MESSAGE_DATE_RANGE = Math.min(
  200,
  Math.max(1, parseInt(process.env.MAX_PDFS_PER_MESSAGE_DATE_RANGE || "200", 10) || 200),
);
/** Safety cap on messages loaded per folder when filtering by received date (pagination continues until nextLink ends or cap). */
const DATE_RANGE_MESSAGE_CAP = Math.min(
  50000,
  Math.max(
    1000,
    parseInt(process.env.DOCUMENT_DETAILS_DATE_RANGE_MESSAGE_CAP || "10000", 10) || 10000,
  ),
);

function documentDetailsCacheStorageKey(folderId, receivedFrom, receivedTo) {
  const f = String(receivedFrom || "").trim();
  const t = String(receivedTo || "").trim();
  if (!f && !t) return folderId;
  const fs = f.replace(/[^\d-]/g, "") || "any";
  const ts = t.replace(/[^\d-]/g, "") || "any";
  return `${folderId}__rf_${fs}__rt_${ts}`;
}

/**
 * @returns {{ range: { fromIso: string|null, toIso: string|null }|null, from: string, to: string }}
 */
function parseReceivedDateQuery(fromStr, toStr) {
  const from = String(fromStr || "").trim();
  const to = String(toStr || "").trim();
  if (!from && !to) {
    return { range: null, from: "", to: "" };
  }
  if (from && !/^\d{4}-\d{2}-\d{2}$/.test(from)) {
    throw new Error("Invalid Import from date (use YYYY-MM-DD).");
  }
  if (to && !/^\d{4}-\d{2}-\d{2}$/.test(to)) {
    throw new Error("Invalid Import to date (use YYYY-MM-DD).");
  }
  const fromIso = from ? `${from}T00:00:00.000Z` : null;
  const toIso = to ? `${to}T23:59:59.999Z` : null;
  if (fromIso && toIso && fromIso > toIso) {
    throw new Error("Import from must be on or before Import to.");
  }
  return {
    range: fromIso || toIso ? { fromIso, toIso } : null,
    from,
    to,
  };
}

const QUERIES_UPLOAD_MB = Math.max(
  2,
  Math.min(32, parseInt(process.env.MAX_QUERIES_XLSX_MB || "12", 10)),
);
const queriesUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: QUERIES_UPLOAD_MB * 1024 * 1024 },
});

/**
 * Parallel Graph calls listing attachments per message.
 * Default 4 — Microsoft Graph enforces a per-mailbox MailboxConcurrency limit; higher values → 429 ApplicationThrottled.
 */
const GRAPH_ATTACHMENT_LIST_CONCURRENCY = Math.max(
  1,
  Math.min(12, parseInt(process.env.GRAPH_ATTACHMENT_LIST_CONCURRENCY || "6", 10)),
);
/** Parallel PDF download ($value) + parse. Each download hits the same mailbox concurrency pool. */
const PDF_PARSE_CONCURRENCY = Math.max(
  1,
  Math.min(12, parseInt(process.env.PDF_PARSE_CONCURRENCY || "5", 10)),
);

const GRAPH_HTTP_MAX_ATTEMPTS = Math.max(
  3,
  Math.min(12, parseInt(process.env.GRAPH_HTTP_MAX_ATTEMPTS || "8", 10)),
);

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** Retry-After header: delta-seconds or HTTP-date */
function graphThrottleWaitMs(res, attemptIndex) {
  const raw = res.headers?.get?.("retry-after");
  if (raw) {
    const sec = parseInt(raw, 10);
    if (!Number.isNaN(sec)) return Math.min(Math.max(sec * 1000, 500), 120000);
    const when = Date.parse(raw);
    if (!Number.isNaN(when)) return Math.min(Math.max(0, when - Date.now()), 120000);
  }
  const backoff = Math.min(2500 * Math.pow(2, attemptIndex), 60000);
  const jitter = Math.floor(Math.random() * 600);
  return backoff + jitter;
}

/**
 * Map with bounded concurrency (no dependencies between items).
 * @template T, R
 * @param {number} concurrency
 * @param {T[]} items
 * @param {(item: T, index: number) => Promise<R>} fn
 * @returns {Promise<R[]>}
 */
async function asyncPool(concurrency, items, fn) {
  if (!items.length) return [];
  const c = Math.max(1, Math.min(concurrency, items.length));
  const results = new Array(items.length);
  let next = 0;
  async function worker() {
    for (;;) {
      const i = next++;
      if (i >= items.length) return;
      results[i] = await fn(items[i], i);
    }
  }
  await Promise.all(Array.from({ length: c }, () => worker()));
  return results;
}

/**
 * Like asyncPool but invokes onChunk(done, total) every `chunkEvery` completions (and at end).
 */
async function asyncPoolWithProgress(concurrency, items, fn, onChunk, chunkEvery = 25) {
  if (!items.length) return [];
  const c = Math.max(1, Math.min(concurrency, items.length));
  const results = new Array(items.length);
  let next = 0;
  let completed = 0;
  const total = items.length;
  const every = Math.max(1, chunkEvery);
  async function worker() {
    for (;;) {
      const i = next++;
      if (i >= items.length) return;
      results[i] = await fn(items[i], i);
      completed++;
      if (onChunk && (completed % every === 0 || completed === total)) {
        onChunk(completed, total);
      }
    }
  }
  await Promise.all(Array.from({ length: c }, () => worker()));
  return results;
}

function getMsalClient() {
  if (!CLIENT_ID || !CLIENT_SECRET || !TENANT_ID) {
    throw new Error("Missing CLIENT_ID, CLIENT_SECRET, or TENANT_ID in .env");
  }
  return new ConfidentialClientApplication({
    auth: {
      clientId: CLIENT_ID,
      clientSecret: CLIENT_SECRET,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    },
  });
}

async function getAccessToken() {
  const client = getMsalClient();
  const result = await client.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  if (!result?.accessToken) {
    throw new Error(result?.errorDescription || "Failed to get access token");
  }
  return result.accessToken;
}

const userPath = () => `/users/${encodeURIComponent(MAILBOX_EMAIL)}`;

async function graphGet(token, pathname) {
  const url = pathname.startsWith("http") ? pathname : `${GRAPH_BASE}${pathname}`;
  let lastErrText = "";
  let lastStatus = 0;
  for (let attempt = 0; attempt < GRAPH_HTTP_MAX_ATTEMPTS; attempt++) {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (res.status === 429 || res.status === 503) {
      lastStatus = res.status;
      lastErrText = await res.text();
      await delay(graphThrottleWaitMs(res, attempt));
      continue;
    }
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Graph API ${res.status}: ${text}`);
    }
    return res.json();
  }
  throw new Error(
    `Graph API ${lastStatus}: ${lastErrText || "too many retries (MailboxConcurrency / throttling)"}`,
  );
}

/** GET binary stream ($value); caller reads body. Retries on 429/503. */
async function graphFetchOkResponse(token, url) {
  let lastErrText = "";
  let lastStatus = 0;
  for (let attempt = 0; attempt < GRAPH_HTTP_MAX_ATTEMPTS; attempt++) {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (res.status === 429 || res.status === 503) {
      lastStatus = res.status;
      lastErrText = await res.text();
      await delay(graphThrottleWaitMs(res, attempt));
      continue;
    }
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Graph API ${res.status}: ${text}`);
    }
    return res;
  }
  throw new Error(
    `Graph API ${lastStatus}: ${lastErrText || "too many retries (MailboxConcurrency / throttling)"}`,
  );
}

function graphErrorStatus(err) {
  const m = err.message || "";
  const match = m.match(/Graph API (\d{3})/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * List attachments. Try the message-global path first so links keep working after the sorter
 * moves mail out of the folder that was used when document details were cached. Fall back to
 * folder-scoped URL for edge cases.
 */
async function fetchAttachmentCollection(token, messageId, folderId) {
  const encMsg = encodeURIComponent(messageId);
  const paths = [];
  paths.push(`${userPath()}/messages/${encMsg}/attachments`);
  if (folderId) {
    paths.push(
      `${userPath()}/mailFolders/${encodeURIComponent(folderId)}/messages/${encMsg}/attachments`,
    );
  }

  let lastErr;
  for (const basePath of paths) {
    try {
      let data = await graphGet(token, basePath);
      const list = [];
      let nextLink = data["@odata.nextLink"];
      list.push(...(data.value || []));
      while (nextLink) {
        data = await graphGet(token, nextLink);
        list.push(...(data.value || []));
        nextLink = data["@odata.nextLink"];
      }
      return list;
    } catch (err) {
      lastErr = err;
      if (graphErrorStatus(err) === 404) continue;
      throw err;
    }
  }
  throw lastErr;
}

/** Resolve attachment metadata: prefer global message path (survives folder moves), then folder-scoped. */
async function resolveAttachmentMeta(token, messageId, attachmentId, folderId) {
  const encMsg = encodeURIComponent(messageId);
  const encAtt = encodeURIComponent(attachmentId);
  const paths = [];
  paths.push(`${userPath()}/messages/${encMsg}/attachments/${encAtt}`);
  if (folderId) {
    paths.push(
      `${userPath()}/mailFolders/${encodeURIComponent(folderId)}/messages/${encMsg}/attachments/${encAtt}`,
    );
  }

  let lastErr;
  for (const metaPath of paths) {
    try {
      const meta = await graphGet(token, metaPath);
      return { meta, metaPath };
    } catch (err) {
      lastErr = err;
      if (graphErrorStatus(err) === 404) continue;
      throw err;
    }
  }
  throw lastErr;
}

function isPdfAttachment(meta, fileName) {
  const ct = (meta.contentType || "").toLowerCase();
  if (ct.includes("pdf")) return true;
  const n = (fileName || "").toLowerCase();
  return n.endsWith(".pdf");
}

async function fetchAttachmentBytes(token, messageId, attachmentId, folderId) {
  const { meta, metaPath } = await resolveAttachmentMeta(
    token,
    messageId,
    attachmentId,
    folderId,
  );
  const odataType = meta["@odata.type"] || "";
  if (!odataType.includes("fileAttachment")) {
    throw new Error("Not a downloadable file attachment");
  }
  const valueUrl = `${GRAPH_BASE}${metaPath}/$value`;
  const streamRes = await graphFetchOkResponse(token, valueUrl);
  return Buffer.from(await streamRes.arrayBuffer());
}

async function pdfBufferToText(buffer) {
  const parser = new PDFParse({ data: buffer });
  try {
    const result = await parser.getText();
    return result?.text ? String(result.text) : "";
  } finally {
    await parser.destroy();
  }
}

function hashUrlForSyntheticAttachment(url) {
  return crypto.createHash("sha1").update(url).digest("hex").slice(0, 14);
}

/**
 * Skip synthesizing a Tax Invoice row only when **this PDF's main extract** is already a Tax
 * Invoice / Invoice with the same doc no. Do not look at other messages/PDFs — otherwise the
 * second statement pack never gets a row for the same doc no. (dedupe across PDFs is separate).
 */
function taxInvoiceMainRowBlocksSynth(mainRow, invUpper) {
  if (!mainRow) return false;
  const u = String(invUpper || "").trim().toUpperCase();
  const dn = String(mainRow.documentNo || "").trim().toUpperCase();
  if (dn !== u) return false;
  const t = String(mainRow.documentType || "").trim().toLowerCase();
  return t === "tax invoice" || t === "invoice";
}

/**
 * Synthetic Tax Invoice rows from statement activity table + embedded PDF hyperlinks (no HTTP fetch).
 */
function appendSyntheticTaxInvoicesFromStatement(rows, ctx) {
  const { text, extracted, pdfHyperlinkUrls, msg, att, folderId, mainRow } = ctx;
  if (!text || extracted.documentType !== "Statement") return;

  const flatProbe = text.replace(/\s+/g, " ");
  if (isSabcAccountStatement(flatProbe)) {
    const sabcLines = parseSabcStatementOutstandingInvoiceLines(text);
    if (sabcLines.length) {
      const embeddedTi = parseSabcEmbeddedTaxInvoiceDetails(text);
      for (const line of sabcLines) {
        const invUpper = String(line.invoiceNo || "").trim();
        if (!invUpper) continue;
        if (taxInvoiceMainRowBlocksSynth(mainRow, invUpper)) continue;

        const amounts = zaInclusiveTotalToNetVatGrossStrings(line.balance);
        const synId = `${att.id}:stmt-synth:${invUpper}`;
        const emb = embeddedTi.get(invUpper);

        rows.push({
          dateDocumentIssued:
            (emb && emb.dateDocumentIssued) || extracted.dateDocumentIssued || "",
          documentType: "Tax Invoice",
          documentNo: invUpper,
          grossAmount: amounts.grossAmount,
          netAmount: amounts.netAmount,
          vatAmount: amounts.vatAmount,
          totalAmount: amounts.totalAmount,
          supplierName: extracted.supplierName || "SABC",
          clientName: (emb && emb.clientName) || extracted.clientName || "",
          brandName: (emb && emb.brandName) || "",
          campaignName: (emb && emb.campaignName) || "",
          campCampaignNo:
            (emb && emb.campCampaignNo) || extracted.campCampaignNo || "",
          bookingOrderNo: (emb && emb.bookingOrderNo) || extracted.bookingOrderNo || "",
          contractNumber: extracted.contractNumber || "",
          purchaseOrderNumber: "",
          statementInvoiceRefs: [],
          messageId: msg.id,
          attachmentId: synId,
          hyperlink: hyperlinkForAttachment(msg.id, att.id, folderId),
          sourceFileName: `${invUpper}-from-statement.pdf`,
          emailSubject: msg.subject || "",
          emailReceivedDateTime: msg.receivedDateTime || "",
          folderId,
          parseFromStatementSynthesis: true,
          statementHyperlinkSourceUrl: "",
        });
      }
      return;
    }
  }

  const activity = parseStatementTaxInvoiceActivityRows(text);
  if (!activity.length) return;

  const paired = pairStatementLinesWithHyperlinkUrls(activity, pdfHyperlinkUrls || []);

  for (const line of paired) {
    const invUpper = line.invoiceNo.toUpperCase();
    if (taxInvoiceMainRowBlocksSynth(mainRow, invUpper)) continue;

    const amounts = zaInclusiveTotalToNetVatGrossStrings(line.inclusiveTotal);
    const synId = `${att.id}:stmt-synth:${invUpper}`;

    rows.push({
      dateDocumentIssued: line.dateDocumentIssued || "",
      documentType: "Tax Invoice",
      documentNo: line.invoiceNo,
      grossAmount: amounts.grossAmount,
      netAmount: amounts.netAmount,
      vatAmount: amounts.vatAmount,
      totalAmount: amounts.totalAmount,
      supplierName: extracted.supplierName || "",
      clientName: extracted.clientName || "",
      brandName: extracted.brandName || "",
      campaignName: extracted.campaignName || "",
      campCampaignNo: extracted.campCampaignNo || "",
      bookingOrderNo: extracted.bookingOrderNo || "",
      contractNumber: extracted.contractNumber || "",
      purchaseOrderNumber: extracted.purchaseOrderNumber || "",
      messageId: msg.id,
      attachmentId: synId,
      hyperlink: line.externalUrl || "#",
      sourceFileName: `${line.invoiceNo}-from-statement.pdf`,
      emailSubject: msg.subject || "",
      emailReceivedDateTime: msg.receivedDateTime || "",
      folderId,
      parseFromStatementSynthesis: true,
      statementHyperlinkSourceUrl: line.externalUrl || "",
    });
  }
}

function invoiceDocMentionedInStatementUrl(url, docUpper) {
  if (!url || !docUpper) return false;
  let u = url;
  try {
    u = decodeURIComponent(url);
  } catch {
    u = url;
  }
  const U = u.toUpperCase();
  if (U.includes(docUpper)) return true;
  const bare = docUpper.replace(/^INV-/i, "");
  if (bare.length < 4) return false;
  return new RegExp(`(?:INV[-_./]?|\\b)${bare}\\b`, "i").test(U);
}

/**
 * Activity statements embed https links to tax invoices. When those PDFs are not separate
 * attachments, fetch each distinct URL once and add rows for INV numbers listed on the statement.
 */
async function augmentRowsFromStatementHyperlinks(rows, folderId) {
  const existingInv = new Set(
    rows
      .filter((r) => r.documentNo && /^INV-/i.test(String(r.documentNo)))
      .map((r) => String(r.documentNo).toUpperCase()),
  );

  const triedUrls = new Set();
  const added = [];

  for (const baseRow of rows) {
    if (baseRow.documentType !== "Statement") continue;
    const refsRaw = baseRow.statementInvoiceRefs || [];
    const urls = baseRow.pdfHyperlinkUrls || [];
    if (!urls.length) continue;

    const refSet = new Set(refsRaw.map((x) => String(x).toUpperCase()));
    const needsFetch =
      refSet.size === 0 ||
      [...refSet].some((inv) => !existingInv.has(inv));

    if (!needsFetch) continue;

    for (const url of urls) {
      const stillNeeded =
        refSet.size === 0 ||
        [...refSet].some((inv) => !existingInv.has(inv));
      if (!stillNeeded) break;
      if (triedUrls.has(url)) continue;
      triedUrls.add(url);

      let buf;
      try {
        buf = await fetchPdfBufferResolvingViewerPages(url, MAX_DOCUMENT_PDF_BYTES);
      } catch {
        continue;
      }
      if (!buf) {
        try {
          const host = new URL(url).hostname;
          if (/xero\.com$/i.test(host)) {
            console.warn(
              "[document-details] Xero invoice links open a browser app (Vue/JS), not a raw PDF. Server-side fetch cannot see the PDF—save/download the Tax Invoice PDF from Xero into the mailbox folder, or attach it to the email:",
              url.slice(0, 160),
            );
          }
        } catch (_) {}
        continue;
      }

      let text = "";
      try {
        text = await pdfBufferToText(buf);
      } catch {
        continue;
      }

      const extracted = extractInvoiceFields(text);
      const docUpper = extracted.documentNo
        ? String(extracted.documentNo).trim().toUpperCase()
        : "";
      if (!docUpper || !/^INV-\d+/i.test(docUpper)) continue;

      const inRefSet = refSet.has(docUpper);
      const urlHint = invoiceDocMentionedInStatementUrl(url, docUpper);
      if (refSet.size > 0 && !inRefSet && !urlHint) continue;
      if (refSet.size === 0 && !urlHint) continue;
      if (existingInv.has(docUpper)) continue;

      existingInv.add(docUpper);
      const synId = `${baseRow.attachmentId}:stmt-href:${hashUrlForSyntheticAttachment(url)}`;
      const fid = baseRow.folderId || folderId;

      added.push({
        ...extracted,
        messageId: baseRow.messageId,
        attachmentId: synId,
        hyperlink: hyperlinkForAttachment(baseRow.messageId, synId, fid),
        sourceFileName: `${docUpper}-from-statement-link.pdf`,
        emailSubject: baseRow.emailSubject || "",
        emailReceivedDateTime: baseRow.emailReceivedDateTime || "",
        folderId: fid,
        parseFromStatementHyperlink: true,
        statementHyperlinkSourceUrl: url,
      });
    }
  }

  return added;
}

function enrichRowIdsFromHyperlink(row) {
  if (row.messageId && row.attachmentId) return row;
  const h = row.hyperlink || "";
  const qi = h.indexOf("?");
  if (qi < 0) return row;
  try {
    const sp = new URLSearchParams(h.slice(qi));
    const mid = sp.get("messageId");
    const aid = sp.get("attachmentId");
    if (mid && aid) return { ...row, messageId: mid, attachmentId: aid };
  } catch (_) {}
  return row;
}

function hyperlinkForAttachment(messageId, attachmentId, folderId) {
  let q =
    "?messageId=" +
    encodeURIComponent(messageId) +
    "&attachmentId=" +
    encodeURIComponent(attachmentId);
  if (folderId) q += "&folderId=" + encodeURIComponent(folderId);
  return `/api/mail/attachment-content${q}`;
}

/**
 * Prefer duplicates with more cells filled from Client through PO (dashboard columns).
 * Each non-empty field adds a large base score plus a small length tie-breaker.
 */
function documentRowDetailRichness(r) {
  const fields = [
    "clientName",
    "brandName",
    "campaignName",
    "campCampaignNo",
    "bookingOrderNo",
    "contractNumber",
    "purchaseOrderNumber",
  ];
  let score = 0;
  for (const f of fields) {
    const t = String(r[f] || "").trim();
    if (t) score += 10000 + Math.min(8000, t.length);
  }
  return score;
}

function compareRowsForDuplicateKeep(a, b) {
  const ra = documentRowDetailRichness(a);
  const rb = documentRowDetailRichness(b);
  if (rb !== ra) return rb - ra;
  return String(b.emailReceivedDateTime || "").localeCompare(
    String(a.emailReceivedDateTime || ""),
  );
}

let warnedDocumentDedupDisabled = false;

/**
 * Identity key for merging: same Doc no. + type. For Tax Invoice / Invoice, **date is ignored** so
 * two PDFs (statement pack vs invoice pack) that set different "Date issued" still collapse to one row.
 */
function documentIdentityDedupKey(r, norm) {
  const doc = norm(r.documentNo);
  const dt = norm(r.documentType);
  if (dt === "tax invoice" || dt === "invoice") {
    return `${doc}\t${dt}`;
  }
  return `${doc}\t${norm(r.dateDocumentIssued)}\t${dt}`;
}

/**
 * Drop duplicate logical documents (same Doc no. + Doc type; date included except for Tax Invoice/Invoice).
 * Keeps the row with the richest Client → PO columns; ties → newest email.
 * Order confirmations: keep every row. Empty Doc no.: never deduped.
 *
 * Set env `DISABLE_DOCUMENT_DETAILS_DEDUP=1` to skip merging (shows every row for debugging).
 */
function deduplicateRowsByDocumentIdentity(rows) {
  const skipDedup =
    process.env.DISABLE_DOCUMENT_DETAILS_DEDUP === "1" ||
    process.env.DISABLE_DOCUMENT_DETAILS_DEDUP === "true";
  if (skipDedup) {
    if (!warnedDocumentDedupDisabled) {
      warnedDocumentDedupDisabled = true;
      console.warn(
        "[document-details] DISABLE_DOCUMENT_DETAILS_DEDUP is set — duplicate doc rows are not merged.",
      );
    }
    return [...(rows || [])];
  }

  const norm = (s) => String(s || "").trim().toLowerCase();
  const list = [...(rows || [])];
  const groups = new Map();

  for (const r of list) {
    if (norm(r.documentType) === "order confirmation") continue;
    const doc = norm(r.documentNo);
    if (!doc || doc === "n/a") continue;
    const k = documentIdentityDedupKey(r, norm);
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(r);
  }

  const winnerByKey = new Map();
  for (const [k, group] of groups) {
    if (group.length === 1) winnerByKey.set(k, group[0]);
    else {
      const sorted = [...group].sort(compareRowsForDuplicateKeep);
      winnerByKey.set(k, sorted[0]);
    }
  }

  const emitted = new Set();
  const out = [];
  for (const r of list) {
    if (norm(r.documentType) === "order confirmation") {
      out.push(r);
      continue;
    }
    const doc = norm(r.documentNo);
    if (!doc || doc === "n/a") {
      out.push(r);
      continue;
    }
    const k = documentIdentityDedupKey(r, norm);
    if (emitted.has(k)) continue;
    const w = winnerByKey.get(k);
    if (w === r) {
      out.push(r);
      emitted.add(k);
    }
  }
  return out;
}

async function buildPdfAttachmentQueue(token, scanFolderId, limit, report, queueOpts = {}) {
  const rep = typeof report === "function" ? report : () => {};
  const receivedDateRange = queueOpts.receivedDateRange || null;
  const dateActive = !!(
    receivedDateRange &&
    (receivedDateRange.fromIso || receivedDateRange.toIso)
  );
  const messageCap = dateActive ? DATE_RANGE_MESSAGE_CAP : limit;
  const effectivePdfCap = dateActive ? MAX_PDFS_PER_MESSAGE_DATE_RANGE : MAX_PDFS_PER_MESSAGE;
  rep({
    phase: "queue",
    percent: 1,
    done: 0,
    total: 0,
    label: dateActive
      ? `Loading messages in date range (up to ${messageCap})…`
      : `Loading message list (up to ${limit})…`,
  });
  const messages = await fetchFolderMessages(
    token,
    scanFolderId,
    messageCap,
    (loaded) => {
      const cap = Math.max(messageCap, 1);
      const p = Math.min(10, 1 + Math.floor((loaded / cap) * 9));
      rep({
        phase: "queue",
        percent: p,
        done: 0,
        total: 0,
        label: `Fetching messages… ${loaded} / ${messageCap}`,
      });
    },
    {
      receivedDateRange,
      paginateAllInRange: dateActive,
      maxMessages: messageCap,
    },
  );
  rep({
    phase: "queue",
    percent: 11,
    done: 0,
    total: 0,
    label: `Found ${messages.length} messages · listing attachments…`,
  });
  const messagesWithAttachments = messages.filter((m) => m.hasAttachments);
  const attachmentLists = await asyncPoolWithProgress(
    GRAPH_ATTACHMENT_LIST_CONCURRENCY,
    messagesWithAttachments,
    async (msg) => {
      try {
        const attachments = await fetchAttachmentCollection(token, msg.id, scanFolderId);
        return { msg, attachments };
      } catch {
        return { msg, attachments: [] };
      }
    },
    (done, total) => {
      const p = total > 0 ? 11 + Math.floor((done / total) * 5) : 14;
      rep({
        phase: "queue",
        percent: Math.min(15, p),
        done,
        total,
        label: `Listing attachments ${done} / ${total}…`,
      });
    },
    35,
  );
  const queue = [];
  const perMsgPdf = new Map();
  for (const { msg, attachments } of attachmentLists) {
    for (const att of attachments) {
      const odataType = att["@odata.type"] || "";
      if (!odataType.includes("fileAttachment")) continue;
      if (!isPdfAttachment(att, att.name)) continue;
      const size = att.size || 0;
      if (size > MAX_DOCUMENT_PDF_BYTES) continue;
      const n = perMsgPdf.get(msg.id) || 0;
      if (n >= effectivePdfCap) continue;
      perMsgPdf.set(msg.id, n + 1);
      queue.push({ msg, att, folderId: scanFolderId });
    }
  }
  return { queue, messagesWithAttachmentsCount: messagesWithAttachments.length };
}

/** Inbox + every child folder under Inbox; dedupe messages by id; PDFs parsed with correct folder scope. */
async function buildPdfAttachmentQueueAllFolders(token, limit, report, queueOpts = {}) {
  const rep = typeof report === "function" ? report : () => {};
  rep({
    phase: "queue",
    percent: 1,
    done: 0,
    total: 0,
    label: "Listing folders under Inbox…",
  });
  const folderIds = ["inbox"];
  const childList = await fetchInboxChildFolders(token);
  folderIds.push(...childList.map((f) => f.id));
  const receivedDateRange = queueOpts.receivedDateRange || null;
  const dateActive = !!(
    receivedDateRange &&
    (receivedDateRange.fromIso || receivedDateRange.toIso)
  );
  const effectivePdfCap = dateActive ? MAX_PDFS_PER_MESSAGE_DATE_RANGE : MAX_PDFS_PER_MESSAGE;
  const perFolder = dateActive
    ? DATE_RANGE_MESSAGE_CAP
    : Math.max(150, Math.ceil(limit / Math.max(1, folderIds.length)));
  const seenMsg = new Set();
  const merged = [];
  let fi = 0;
  for (const fid of folderIds) {
    fi++;
    rep({
      phase: "queue",
      percent: Math.min(8, 1 + Math.floor((fi / folderIds.length) * 7)),
      done: 0,
      total: 0,
      label: `Reading folder ${fi} / ${folderIds.length}…`,
    });
    const messages = await fetchFolderMessages(
      token,
      fid,
      perFolder,
      (loaded) => {
        rep({
          phase: "queue",
          percent: Math.min(11, 2 + Math.floor((loaded / Math.max(perFolder, 1)) * 9)),
          done: 0,
          total: 0,
          label: `Folder ${fi}/${folderIds.length}: ${loaded} messages loaded…`,
        });
      },
      {
        receivedDateRange,
        paginateAllInRange: dateActive,
        maxMessages: perFolder,
      },
    );
    for (const msg of messages) {
      if (seenMsg.has(msg.id)) continue;
      seenMsg.add(msg.id);
      merged.push({ msg, fid });
    }
  }
  merged.sort((a, b) =>
    String(b.msg.receivedDateTime || "").localeCompare(String(a.msg.receivedDateTime || "")),
  );
  rep({
    phase: "queue",
    percent: 12,
    done: 0,
    total: 0,
    label: `Merged ${merged.length} messages · listing attachments…`,
  });
  const withAtt = merged.filter((x) => x.msg.hasAttachments);
  const attachmentLists = await asyncPoolWithProgress(
    GRAPH_ATTACHMENT_LIST_CONCURRENCY,
    withAtt,
    async ({ msg, fid }) => {
      try {
        const attachments = await fetchAttachmentCollection(token, msg.id, fid);
        return { msg, fid, attachments };
      } catch {
        return { msg, fid, attachments: [] };
      }
    },
    (done, total) => {
      const p = total > 0 ? 12 + Math.floor((done / total) * 4) : 14;
      rep({
        phase: "queue",
        percent: Math.min(15, p),
        done,
        total,
        label: `Listing attachments ${done} / ${total}…`,
      });
    },
    35,
  );
  const queue = [];
  const perMsgPdf = new Map();
  for (const { msg, fid, attachments } of attachmentLists) {
    for (const att of attachments) {
      const odataType = att["@odata.type"] || "";
      if (!odataType.includes("fileAttachment")) continue;
      if (!isPdfAttachment(att, att.name)) continue;
      const size = att.size || 0;
      if (size > MAX_DOCUMENT_PDF_BYTES) continue;
      const n = perMsgPdf.get(msg.id) || 0;
      if (n >= effectivePdfCap) continue;
      perMsgPdf.set(msg.id, n + 1);
      queue.push({ msg, att, folderId: fid });
    }
  }
  const messagesWithAttachmentsCount = merged.filter((x) => x.msg.hasAttachments).length;
  return { queue, messagesWithAttachmentsCount };
}

function sseWrite(res, eventName, payload) {
  if (res.writableEnded) return;
  res.write(`event: ${eventName}\ndata: ${JSON.stringify(payload)}\n\n`);
}

/**
 * Queue all PDF attachments, then parse each. onProgress({ phase, percent, done, total, label }).
 * opts: { forceFull?: boolean } — forceFull skips cache and fetches all messages up to limit.
 * folderId "all" scans Inbox plus every folder under Inbox (merged view).
 * @param {number} limit — max messages per folder (newest first); use env default ~2500.
 */
async function computeDocumentDetailRows(token, folderId, limit, onProgress, opts = {}) {
  const noop = () => {};
  const report = typeof onProgress === "function" ? onProgress : noop;
  const forceFull = !!opts.forceFull;
  const scanAllFolders = String(folderId || "").toLowerCase() === "all";

  const parsedDates = parseReceivedDateQuery(opts.receivedFrom, opts.receivedTo);
  const receivedDateRange = parsedDates.range;
  const rawFrom = parsedDates.from;
  const rawTo = parsedDates.to;
  const dateRangeActive = !!(
    receivedDateRange &&
    (receivedDateRange.fromIso || receivedDateRange.toIso)
  );
  const logicalFolderId = scanAllFolders ? "all" : folderId;
  const cacheStorageKey = documentDetailsCacheStorageKey(logicalFolderId, rawFrom, rawTo);
  const queueOpts = { receivedDateRange };
  const effectivePdfCapForIncremental = dateRangeActive
    ? MAX_PDFS_PER_MESSAGE_DATE_RANGE
    : MAX_PDFS_PER_MESSAGE;

  let cachedRows = [];
  let lastCachedIso = null;
  if (!forceFull) {
    try {
      const raw = await readFolderRowsFromCache(cacheStorageKey);
      cachedRows = deduplicateRowsByDocumentIdentity(raw.map(enrichRowIdsFromHyperlink));
      lastCachedIso = maxEmailReceivedIso(cachedRows);
    } catch {
      cachedRows = [];
    }
  }

  if (!forceFull && scanAllFolders && cachedRows.length > 0) {
    report({
      phase: "done",
      percent: 100,
      done: 0,
      total: 0,
      label: "Loaded merged folders from store…",
    });
    return {
      rows: cachedRows,
      scannedMessages: 0,
      pdfAttachmentsParsed: 0,
      incremental: true,
      cachedRowsUsed: cachedRows.length,
      newPdfAttachmentsParsed: 0,
      forceFull: false,
      maxPdfsPerMessage: dateRangeActive
        ? MAX_PDFS_PER_MESSAGE_DATE_RANGE
        : MAX_PDFS_PER_MESSAGE,
      receivedFrom: rawFrom,
      receivedTo: rawTo,
      dateRangeActive,
    };
  }

  let incremental =
    !forceFull && !scanAllFolders && cachedRows.length > 0 && !!lastCachedIso;

  let messagesWithAttachmentsCount = 0;
  let queue = [];

  report({
    phase: "queue",
    percent: 0,
    done: 0,
    total: 0,
    label: forceFull
      ? "Starting full sync…"
      : scanAllFolders
        ? "Preparing multi-folder scan…"
        : "Connecting to mailbox…",
  });

  if (scanAllFolders) {
    const built = await buildPdfAttachmentQueueAllFolders(token, limit, report, queueOpts);
    queue = built.queue;
    messagesWithAttachmentsCount = built.messagesWithAttachmentsCount;
  } else if (incremental) {
    report({
      phase: "queue",
      percent: 1,
      done: 0,
      total: 0,
      label: "Checking for new mail since last sync…",
    });
    const messages = await fetchFolderMessagesSince(
      token,
      folderId,
      lastCachedIso,
      limit,
      (loaded) => {
        report({
          phase: "queue",
          percent: Math.min(9, 1 + Math.floor((loaded / Math.max(limit, 1)) * 8)),
          done: 0,
          total: 0,
          label: `Loading new messages… ${loaded} / ${limit}`,
        });
      },
      { receivedDateRange },
    );
    report({
      phase: "queue",
      percent: 10,
      done: 0,
      total: 0,
      label:
        messages.length === 0
          ? "No new messages since last sync — loading from store…"
          : `Found ${messages.length} new message(s) · listing attachments…`,
    });
    const messagesWithAttachments = messages.filter((m) => m.hasAttachments);
    messagesWithAttachmentsCount = messagesWithAttachments.length;
    const attachmentLists = await asyncPoolWithProgress(
      GRAPH_ATTACHMENT_LIST_CONCURRENCY,
      messagesWithAttachments,
      async (msg) => {
        try {
          const attachments = await fetchAttachmentCollection(token, msg.id, folderId);
          return { msg, attachments };
        } catch {
          return { msg, attachments: [] };
        }
      },
      (done, total) => {
        const p = total > 0 ? 10 + Math.floor((done / total) * 5) : 13;
        report({
          phase: "queue",
          percent: Math.min(15, p),
          done,
          total,
          label: `Listing attachments ${done} / ${total}…`,
        });
      },
      35,
    );
    const perMsgPdf = new Map();
    for (const { msg, attachments } of attachmentLists) {
      for (const att of attachments) {
        const odataType = att["@odata.type"] || "";
        if (!odataType.includes("fileAttachment")) continue;
        if (!isPdfAttachment(att, att.name)) continue;
        const size = att.size || 0;
        if (size > MAX_DOCUMENT_PDF_BYTES) continue;
        const n = perMsgPdf.get(msg.id) || 0;
        if (n >= effectivePdfCapForIncremental) continue;
        perMsgPdf.set(msg.id, n + 1);
        queue.push({ msg, att, folderId });
      }
    }
  } else {
    const built = await buildPdfAttachmentQueue(token, folderId, limit, report, queueOpts);
    queue = built.queue;
    messagesWithAttachmentsCount = built.messagesWithAttachmentsCount;
    incremental = false;
  }

  const totalPdfs = queue.length;

  if (incremental && totalPdfs === 0 && cachedRows.length > 0) {
    report({
      phase: "done",
      percent: 100,
      done: 0,
      total: 0,
      label: "Done",
    });
    return {
      rows: cachedRows,
      scannedMessages: 0,
      pdfAttachmentsParsed: 0,
      incremental: true,
      cachedRowsUsed: cachedRows.length,
      newPdfAttachmentsParsed: 0,
      forceFull: false,
      maxPdfsPerMessage: dateRangeActive
        ? MAX_PDFS_PER_MESSAGE_DATE_RANGE
        : MAX_PDFS_PER_MESSAGE,
      receivedFrom: rawFrom,
      receivedTo: rawTo,
      dateRangeActive,
    };
  }

  const rows = [];
  let pdfAttempted = 0;

  let parseDone = 0;
  const parseResults = await asyncPool(PDF_PARSE_CONCURRENCY, queue, async (item) => {
    const { msg, att, folderId: scanFid } = item;
    let text = "";
    let parseError = "";
    let pdfBuf = null;
    try {
      pdfBuf = await fetchAttachmentBytes(token, msg.id, att.id, scanFid);
      text = await pdfBufferToText(pdfBuf);
    } catch (e) {
      parseError = e.message || String(e);
    }

    const extracted = extractInvoiceFields(text);
    if (
      extracted.supplierName === "Google" &&
      (extracted.documentType === "Credit Memo" ||
        extracted.documentType === "Invoice")
    ) {
      void writeGoogleExtractDebuggerFile(text, {
        messageId: msg.id,
        attachmentId: att.id,
        sourceFileName: att.name || "",
        folderId: scanFid,
      });
    }
    let statementInvoiceRefs = extracted.statementInvoiceRefs || [];
    let pdfHyperlinkUrls = [];
    if (extracted.documentType === "Statement" && pdfBuf) {
      try {
        pdfHyperlinkUrls = extractHttpsUrlsFromPdfBuffer(pdfBuf);
        statementInvoiceRefs = [
          ...new Set([
            ...statementInvoiceRefs,
            ...extractInvNumbersFromPdfBinary(pdfBuf),
          ]),
        ];
      } catch {
        pdfHyperlinkUrls = [];
      }
    }

    parseDone++;
    const pct =
      totalPdfs > 0
        ? Math.round(15 + (parseDone / totalPdfs) * 84)
        : 100;
    report({
      phase: "parse",
      percent: Math.min(99, pct),
      done: parseDone,
      total: totalPdfs,
      label: att.name || "attachment.pdf",
    });

    const mainRow = {
      ...extracted,
      statementInvoiceRefs,
      pdfHyperlinkUrls,
      messageId: msg.id,
      attachmentId: att.id,
      hyperlink: hyperlinkForAttachment(msg.id, att.id, scanFid),
      sourceFileName: att.name || "",
      emailSubject: msg.subject || "",
      emailReceivedDateTime: msg.receivedDateTime || "",
      folderId: scanFid,
      parseError: parseError || undefined,
    };

    return {
      mainRow,
      synthCtx: {
        text,
        extracted,
        pdfHyperlinkUrls,
        msg,
        att,
        folderId: scanFid,
        mainRow,
      },
    };
  });

  pdfAttempted = parseResults.length;
  for (let i = 0; i < parseResults.length; i++) {
    rows.push(parseResults[i].mainRow);
    appendSyntheticTaxInvoicesFromStatement(rows, parseResults[i].synthCtx);
  }

  try {
    const linked = await augmentRowsFromStatementHyperlinks(rows, cacheStorageKey);
    if (linked.length) rows.push(...linked);
  } catch (e) {
    console.warn("Statement hyperlink expansion failed:", e.message || e);
  }

  let mergedRows = rows;
  if (incremental && cachedRows.length > 0 && !scanAllFolders) {
    mergedRows = mergeRowsByAttachmentKey(cacheStorageKey, cachedRows, rows);
  }

  mergedRows = deduplicateRowsByDocumentIdentity(mergedRows);

  try {
    await writeFolderRowsToCache(cacheStorageKey, mergedRows);
  } catch (e) {
    console.warn("Could not save document-details cache:", e.message);
  }

  report({
    phase: "done",
    percent: 100,
    done: totalPdfs,
    total: totalPdfs,
    label: "Done",
  });

  return {
    rows: mergedRows,
    scannedMessages: messagesWithAttachmentsCount,
    pdfAttachmentsParsed: pdfAttempted,
    maxPdfsPerMessage: dateRangeActive
      ? MAX_PDFS_PER_MESSAGE_DATE_RANGE
      : MAX_PDFS_PER_MESSAGE,
    incremental,
    cachedRowsUsed: incremental ? cachedRows.length : 0,
    newPdfAttachmentsParsed: pdfAttempted,
    forceFull,
    receivedFrom: rawFrom,
    receivedTo: rawTo,
    dateRangeActive,
  };
}

function preferInlineDisposition(contentType) {
  if (!contentType || typeof contentType !== "string") return false;
  const base = contentType.split(";")[0].trim().toLowerCase();
  return (
    base.startsWith("image/") ||
    base === "application/pdf" ||
    base.startsWith("text/")
  );
}

async function graphPost(token, pathname, body) {
  const url = pathname.startsWith("http") ? pathname : `${GRAPH_BASE}${pathname}`;
  const payload = JSON.stringify(body);
  let lastErrText = "";
  let lastStatus = 0;
  for (let attempt = 0; attempt < GRAPH_HTTP_MAX_ATTEMPTS; attempt++) {
    const res = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: payload,
    });
    if (res.status === 429 || res.status === 503) {
      lastStatus = res.status;
      lastErrText = await res.text();
      await delay(graphThrottleWaitMs(res, attempt));
      continue;
    }
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Graph API ${res.status}: ${text}`);
    }
    return res.json();
  }
  throw new Error(
    `Graph API ${lastStatus}: ${lastErrText || "too many retries (MailboxConcurrency / throttling)"}`,
  );
}

function getSenderDomain(msg) {
  const addr = msg.from?.emailAddress?.address;
  if (!addr || typeof addr !== "string") return null;
  const i = addr.indexOf("@");
  return i === -1 ? null : addr.slice(i + 1).toLowerCase();
}

// Fetch all child folders of Inbox (paginated – Graph returns ~10 per page by default)
async function fetchInboxChildFolders(token) {
  const path = `${userPath()}/mailFolders/inbox/childFolders`;
  let data = await graphGet(token, path);
  let list = data.value || [];
  let nextLink = data["@odata.nextLink"];
  while (nextLink) {
    data = await graphGet(token, nextLink);
    list = list.concat(data.value || []);
    nextLink = data["@odata.nextLink"];
  }
  return list;
}

// Fetch all messages in a folder (paginated). Optional onPageProgress(loadedCount) after each page.
// Optional arg5 fetchOpts: { receivedDateRange?, paginateAllInRange?, maxMessages?, onPageProgress? }
async function fetchFolderMessages(token, folderId, limit = 500, arg4, arg5) {
  let onPageProgress;
  let fetchOpts = {};
  if (typeof arg4 === "function") {
    onPageProgress = arg4;
    if (arg5 && typeof arg5 === "object") fetchOpts = arg5;
  } else if (arg4 && typeof arg4 === "object") {
    fetchOpts = arg4;
    onPageProgress = fetchOpts.onPageProgress;
  }
  const receivedDateRange = fetchOpts.receivedDateRange || null;
  const paginateAllInRange = !!fetchOpts.paginateAllInRange;
  const hasDate = !!(
    receivedDateRange &&
    (receivedDateRange.fromIso || receivedDateRange.toIso)
  );
  let hardCap = limit;
  if (hasDate && paginateAllInRange) {
    const cap =
      fetchOpts.maxMessages != null ? fetchOpts.maxMessages : DATE_RANGE_MESSAGE_CAP;
    hardCap = Math.min(DATE_RANGE_MESSAGE_CAP, Math.max(limit, cap));
  }
  const pageTop = Math.min(hardCap, 999);
  const filterParts = [];
  if (hasDate) {
    if (receivedDateRange.fromIso) {
      filterParts.push(`receivedDateTime ge ${receivedDateRange.fromIso}`);
    }
    if (receivedDateRange.toIso) {
      filterParts.push(`receivedDateTime le ${receivedDateRange.toIso}`);
    }
  }
  const dateFilter = filterParts.join(" and ");
  const path = dateFilter
    ? `${userPath()}/mailFolders/${encodeURIComponent(folderId)}/messages?$filter=${encodeURIComponent(dateFilter)}&$orderby=receivedDateTime desc&$top=${pageTop}&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments`
    : `${userPath()}/mailFolders/${encodeURIComponent(folderId)}/messages?$top=${pageTop}&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments`;
  const data = await graphGet(token, path);
  let list = data.value || [];
  let nextLink = data["@odata.nextLink"];
  if (typeof onPageProgress === "function") {
    onPageProgress(Math.min(list.length, hardCap));
  }
  while (nextLink && list.length < hardCap) {
    const next = await graphGet(token, nextLink);
    list = list.concat(next.value || []);
    nextLink = next["@odata.nextLink"];
    if (typeof onPageProgress === "function") {
      onPageProgress(Math.min(list.length, hardCap));
    }
  }
  return list.slice(0, hardCap);
}

/** Messages received strictly after `sinceIso` (ISO 8601). Falls back to full folder fetch on filter errors. */
async function fetchFolderMessagesSince(
  token,
  folderId,
  sinceIso,
  limit = 500,
  onPageProgress,
  fetchOpts = {},
) {
  const receivedDateRange = fetchOpts.receivedDateRange || null;
  const hasRange = !!(
    receivedDateRange &&
    (receivedDateRange.fromIso || receivedDateRange.toIso)
  );
  if (!sinceIso) {
    return fetchFolderMessages(token, folderId, limit, onPageProgress, fetchOpts);
  }
  const parts = [`receivedDateTime gt ${sinceIso}`];
  if (hasRange) {
    if (receivedDateRange.fromIso) {
      parts.push(`receivedDateTime ge ${receivedDateRange.fromIso}`);
    }
    if (receivedDateRange.toIso) {
      parts.push(`receivedDateTime le ${receivedDateRange.toIso}`);
    }
  }
  const filter = parts.join(" and ");
  const pageTop = Math.min(limit, 999);
  const base = `${userPath()}/mailFolders/${encodeURIComponent(folderId)}/messages?$filter=${encodeURIComponent(filter)}&$orderby=receivedDateTime desc&$top=${pageTop}&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments`;
  try {
    const data = await graphGet(token, base);
    let list = data.value || [];
    let nextLink = data["@odata.nextLink"];
    if (typeof onPageProgress === "function") {
      onPageProgress(Math.min(list.length, limit));
    }
    while (nextLink && list.length < limit) {
      const next = await graphGet(token, nextLink);
      list = list.concat(next.value || []);
      nextLink = next["@odata.nextLink"];
      if (typeof onPageProgress === "function") {
        onPageProgress(Math.min(list.length, limit));
      }
    }
    return list.slice(0, limit);
  } catch (err) {
    console.warn("Incremental message filter failed, using full folder fetch:", err.message);
    return fetchFolderMessages(token, folderId, limit, onPageProgress, fetchOpts);
  }
}

// Build domain → folder map: each domain maps to the folder that has the most senders with that domain
async function buildDomainToFolderMap(token, inboxChildFolders, samplePerFolder = 150) {
  const domainFolderCounts = new Map(); // domain -> Map(folderId -> { folderName, count })
  for (const folder of inboxChildFolders) {
    const messages = await fetchFolderMessages(token, folder.id, samplePerFolder);
    for (const msg of messages) {
      const domain = getSenderDomain(msg);
      if (!domain) continue;
      if (!domainFolderCounts.has(domain)) {
        domainFolderCounts.set(domain, new Map());
      }
      const byFolder = domainFolderCounts.get(domain);
      const prev = byFolder.get(folder.id) || { folderName: folder.displayName, count: 0 };
      byFolder.set(folder.id, {
        folderName: folder.displayName,
        count: prev.count + 1,
      });
    }
  }
  const domainToFolder = new Map();
  for (const [domain, byFolder] of domainFolderCounts) {
    let best = null;
    for (const [folderId, info] of byFolder) {
      if (!best || info.count > best.count) {
        best = { folderId, folderName: info.folderName, count: info.count };
      }
    }
    if (best) {
      domainToFolder.set(domain, { folderId: best.folderId, folderName: best.folderName });
    }
  }
  return domainToFolder;
}

/**
 * @param {string} token
 * @param {(p: { phase: string, percent: number, label?: string }) => void} [onProgress]
 */
async function runMailSort(token, onProgress) {
  const emit =
    typeof onProgress === "function"
      ? onProgress
      : () => {};
  const base = userPath();

  const report = {
    domainMap: [],
    allocated: [],
    corrected: [],
    leftInInbox: 0,
    errors: [],
  };

  emit({ phase: "setup", percent: 3, label: "Loading Inbox folders…" });
  await graphGet(token, `${base}/mailFolders/inbox`);
  const childList = await fetchInboxChildFolders(token);
  const childFolders = childList.map((f) => ({
    id: f.id,
    displayName: f.displayName,
  }));

  if (childFolders.length === 0) {
    return {
      report,
      message:
        "No child folders under Inbox. Create folders (e.g. 12 Star, 365 Digital) first.",
    };
  }

  emit({
    phase: "domain-map",
    percent: 8,
    label: "Mapping sender domains to folders…",
  });
  const domainToFolder = await buildDomainToFolderMap(token, childFolders, 150);
  for (const [domain, info] of domainToFolder) {
    report.domainMap.push({
      domain,
      folderId: info.folderId,
      folderName: info.folderName,
    });
  }

  emit({ phase: "allocate", percent: 12, label: "Reading Inbox…" });
  const inboxMessages = await fetchFolderMessages(token, "inbox", 500);
  const nInbox = inboxMessages.length;

  for (let i = 0; i < inboxMessages.length; i++) {
    const msg = inboxMessages[i];
    const domain = getSenderDomain(msg);
    const target = domain ? domainToFolder.get(domain) : null;
    if (!target) {
      report.leftInInbox++;
    } else {
      try {
        await graphPost(token, `${base}/messages/${msg.id}/move`, {
          destinationId: target.folderId,
        });
        report.allocated.push({
          messageId: msg.id,
          subject: msg.subject,
          from: msg.from?.emailAddress?.address,
          fromDomain: domain,
          toFolder: target.folderName,
        });
      } catch (e) {
        report.errors.push({
          action: "allocate",
          subject: msg.subject,
          error: e.message,
        });
      }
    }
    const pct = 12 + Math.floor(((i + 1) / Math.max(nInbox, 1)) * 38);
    emit({
      phase: "allocate",
      percent: Math.min(pct, 50),
      label:
        nInbox === 0
          ? "Inbox is empty."
          : "Inbox: " + (i + 1) + " / " + nInbox + " messages…",
    });
  }

  const numFolders = childFolders.length;
  for (let fi = 0; fi < childFolders.length; fi++) {
    const folder = childFolders[fi];
    const messages = await fetchFolderMessages(token, folder.id, 300);
    const n = messages.length;
    emit({
      phase: "correct",
      percent: 50 + Math.floor((fi / Math.max(numFolders, 1)) * 4),
      label: "Checking folder “" + folder.displayName + "”…",
    });
    for (let mi = 0; mi < messages.length; mi++) {
      const msg = messages[mi];
      const domain = getSenderDomain(msg);
      const target = domain ? domainToFolder.get(domain) : null;
      if (target && target.folderId !== folder.id) {
        try {
          await graphPost(token, `${base}/messages/${msg.id}/move`, {
            destinationId: target.folderId,
          });
          report.corrected.push({
            messageId: msg.id,
            subject: msg.subject,
            from: msg.from?.emailAddress?.address,
            fromDomain: domain,
            fromFolder: folder.displayName,
            toFolder: target.folderName,
          });
        } catch (e) {
          report.errors.push({
            action: "correct",
            subject: msg.subject,
            error: e.message,
          });
        }
      }
      const sub = (mi + 1) / Math.max(n, 1);
      const pct = 50 + Math.floor(((fi + sub) / Math.max(numFolders, 1)) * 48);
      emit({
        phase: "correct",
        percent: Math.min(Math.round(pct), 99),
        label:
          folder.displayName +
          ": " +
          (mi + 1) +
          " / " +
          n +
          " messages…",
      });
    }
  }

  emit({ phase: "done", percent: 100, label: "Finishing…" });

  const message = `Allocated ${report.allocated.length} to folders, corrected ${report.corrected.length} mis-filed. ${report.leftInInbox} left in Inbox (no matching domain).`;
  return { report, message };
}

// Inbox endpoint: returns messages for the configured mailbox
app.get("/api/mail/inbox", async (req, res) => {
  try {
    const limit = Math.min(parseInt(req.query.limit, 10) || 100, 500);
    const token = await getAccessToken();
    const inboxPath = `${userPath()}/mailFolders/inbox/messages?$top=${limit}&$orderby=receivedDateTime desc&$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments`;
    const data = await graphGet(token, inboxPath);
    res.json({ mailbox: MAILBOX_EMAIL, messages: data.value || [] });
  } catch (err) {
    console.error("Inbox error:", err.message);
    res.status(500).json({
      error: "Failed to load inbox",
      message: err.message,
    });
  }
});

// List folders: Inbox + its child folders (12 Star, 365 Digital, etc.)
app.get("/api/mail/folders", async (req, res) => {
  try {
    const token = await getAccessToken();
    const base = userPath();
    const inboxRes = await graphGet(token, `${base}/mailFolders/inbox`);
    const childList = await fetchInboxChildFolders(token);
    const childFolders = childList.map((f) => ({
      id: f.id,
      displayName: f.displayName,
      totalItemCount: f.totalItemCount ?? 0,
      unreadItemCount: f.unreadItemCount ?? 0,
    }));
    const folders = [
      {
        id: "inbox",
        displayName: inboxRes.displayName || "Inbox",
        totalItemCount: inboxRes.totalItemCount ?? 0,
        unreadItemCount: inboxRes.unreadItemCount ?? 0,
      },
      ...childFolders,
    ];
    res.json({ mailbox: MAILBOX_EMAIL, folders });
  } catch (err) {
    console.error("Folders error:", err.message);
    res.status(500).json({
      error: "Failed to load folders",
      message: err.message,
    });
  }
});

// Last-saved rows only — no Microsoft Graph call (instant first paint for Document details).
// GET /api/mail/document-details/cached?folderId=inbox&receivedFrom=&receivedTo=
app.get("/api/mail/document-details/cached", async (req, res) => {
  try {
    const folderId =
      typeof req.query.folderId === "string" && req.query.folderId.length > 0
        ? req.query.folderId
        : "inbox";
    const receivedFrom =
      typeof req.query.receivedFrom === "string" ? req.query.receivedFrom : "";
    const receivedTo =
      typeof req.query.receivedTo === "string" ? req.query.receivedTo : "";
    parseReceivedDateQuery(receivedFrom, receivedTo);
    const cacheKey = documentDetailsCacheStorageKey(folderId, receivedFrom, receivedTo);
    const rows = deduplicateRowsByDocumentIdentity(
      (await readFolderRowsFromCache(cacheKey)).map(enrichRowIdsFromHyperlink),
    );
    res.json({
      mailbox: MAILBOX_EMAIL,
      folderId,
      receivedFrom,
      receivedTo,
      rows,
      rowCount: rows.length,
      fromDisk: true,
    });
  } catch (err) {
    console.error("Document details cache read error:", err.message);
    res.status(500).json({
      error: "Failed to read cached document details",
      message: err.message,
    });
  }
});

// Parsed invoice rows — JSON snapshot (same logic as stream endpoint)
// GET /api/mail/document-details?folderId=inbox&limit=120&full=1 (full rescan, ignore incremental cache)
app.get("/api/mail/document-details", async (req, res) => {
  try {
    const folderId =
      typeof req.query.folderId === "string" && req.query.folderId.length > 0
        ? req.query.folderId
        : "inbox";
    const limit = effectiveDocumentDetailsLimit(req.query.limit);
    const forceFull =
      req.query.full === "1" || req.query.full === "true" || req.query.mode === "full";
    const receivedFrom =
      typeof req.query.receivedFrom === "string" ? req.query.receivedFrom : "";
    const receivedTo =
      typeof req.query.receivedTo === "string" ? req.query.receivedTo : "";
    const token = await getAccessToken();
    const result = await computeDocumentDetailRows(token, folderId, limit, null, {
      forceFull,
      receivedFrom,
      receivedTo,
    });
    res.json({
      mailbox: MAILBOX_EMAIL,
      folderId,
      receivedFrom: result.receivedFrom ?? receivedFrom,
      receivedTo: result.receivedTo ?? receivedTo,
      dateRangeActive: !!result.dateRangeActive,
      messageScanLimit: limit,
      scannedMessages: result.scannedMessages,
      pdfAttachmentsParsed: result.pdfAttachmentsParsed,
      maxPdfsPerMessage: result.maxPdfsPerMessage ?? MAX_PDFS_PER_MESSAGE,
      rows: result.rows,
      incremental: !!result.incremental,
      cachedRowsUsed: result.cachedRowsUsed ?? 0,
      newPdfAttachmentsParsed: result.newPdfAttachmentsParsed ?? result.pdfAttachmentsParsed,
      forceFull: !!result.forceFull,
    });
  } catch (err) {
    console.error("Document details error:", err.message);
    res.status(500).json({
      error: "Failed to build document details",
      message: err.message,
    });
  }
});

// SSE: progress events while scanning — use with EventSource on the dashboard
// GET /api/mail/document-details/stream?folderId=inbox&limit=120&full=1
app.get("/api/mail/document-details/stream", async (req, res) => {
  const folderId =
    typeof req.query.folderId === "string" && req.query.folderId.length > 0
      ? req.query.folderId
      : "inbox";
  const limit = effectiveDocumentDetailsLimit(req.query.limit);
  const forceFull =
    req.query.full === "1" || req.query.full === "true" || req.query.mode === "full";
  const receivedFrom =
    typeof req.query.receivedFrom === "string" ? req.query.receivedFrom : "";
  const receivedTo =
    typeof req.query.receivedTo === "string" ? req.query.receivedTo : "";

  res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no");
  if (typeof res.flushHeaders === "function") res.flushHeaders();

  try {
    const token = await getAccessToken();
    const result = await computeDocumentDetailRows(
      token,
      folderId,
      limit,
      (p) => {
        sseWrite(res, "progress", p);
        if (typeof res.flush === "function") res.flush();
      },
      { forceFull, receivedFrom, receivedTo },
    );
    sseWrite(res, "complete", {
      mailbox: MAILBOX_EMAIL,
      folderId,
      receivedFrom: result.receivedFrom ?? receivedFrom,
      receivedTo: result.receivedTo ?? receivedTo,
      dateRangeActive: !!result.dateRangeActive,
      messageScanLimit: limit,
      scannedMessages: result.scannedMessages,
      pdfAttachmentsParsed: result.pdfAttachmentsParsed,
      maxPdfsPerMessage: result.maxPdfsPerMessage ?? MAX_PDFS_PER_MESSAGE,
      rows: result.rows,
      incremental: !!result.incremental,
      cachedRowsUsed: result.cachedRowsUsed ?? 0,
      newPdfAttachmentsParsed: result.newPdfAttachmentsParsed ?? result.pdfAttachmentsParsed,
      forceFull: !!result.forceFull,
    });
    res.end();
  } catch (err) {
    console.error("Document details stream error:", err.message);
    sseWrite(res, "fatal", { message: err.message });
    res.end();
  }
});

// GET /api/queries — last imported Excel (for Query details tab + merge keys).
app.get("/api/queries", async (req, res) => {
  try {
    const data = await readQueriesImport();
    res.json({
      columns: data.columns,
      rows: data.rows,
      rowCount: data.rows.length,
      queryNumberHeader: data.queryNumberHeader,
      documentNoHeader: data.documentNoHeader,
      importedAt: data.importedAt,
      fileName: data.fileName,
    });
  } catch (err) {
    console.error("Queries read error:", err.message);
    res.status(500).json({
      error: "Failed to read queries import",
      message: err.message,
    });
  }
});

// POST /api/queries/import — multipart field "file" (.xlsx / .xls, first sheet).
app.post(
  "/api/queries/import",
  (req, res, next) => {
    queriesUpload.single("file")(req, res, (err) => {
      if (err) {
        if (err.code === "LIMIT_FILE_SIZE") {
          return res.status(400).json({
            error: "File too large",
            message: `Maximum upload size is ${QUERIES_UPLOAD_MB} MB.`,
          });
        }
        return res.status(400).json({
          error: "Upload failed",
          message: err.message || String(err),
        });
      }
      next();
    });
  },
  async (req, res) => {
    try {
      if (!req.file || !req.file.buffer) {
        return res.status(400).json({
          error: "Missing file",
          message: 'Choose an Excel file (.xlsx or .xls). The form field name must be "file".',
        });
      }
      const lower = (req.file.originalname || "").toLowerCase();
      if (!lower.endsWith(".xlsx") && !lower.endsWith(".xls")) {
        return res.status(400).json({
          error: "Unsupported format",
          message: "Upload a .xlsx or .xls workbook.",
        });
      }
      const { columns, rows } = parseQueriesExcelBuffer(req.file.buffer);
      if (!columns.length) {
        return res.status(400).json({
          error: "Empty sheet",
          message: "The first sheet has no column headers.",
        });
      }
      const { queryNumberHeader, documentNoHeader } = detectQueryMergeColumns(columns);
      if (!documentNoHeader || !queryNumberHeader) {
        return res.status(400).json({
          error: "Could not find merge columns",
          message:
            'The sheet needs a "Query Number" column and a "Document no" (or Document number) column to link rows to document details.',
          columns,
        });
      }
      const payload = {
        columns,
        rows,
        queryNumberHeader,
        documentNoHeader,
        importedAt: new Date().toISOString(),
        fileName: req.file.originalname || "",
      };
      await writeQueriesImport(payload);
      res.json({
        ok: true,
        rowCount: rows.length,
        columns,
        queryNumberHeader,
        documentNoHeader,
        importedAt: payload.importedAt,
        fileName: payload.fileName,
      });
    } catch (err) {
      console.error("Queries import error:", err.message);
      res.status(500).json({
        error: "Failed to import spreadsheet",
        message: err.message || String(err),
      });
    }
  },
);

// Attachments on a message (IDs in query — Graph message ids may contain "/" etc., which breaks path routes).
// GET /api/mail/attachments?messageId=...&folderId=inbox
app.get("/api/mail/attachments", async (req, res) => {
  try {
    const messageId = req.query.messageId;
    if (!messageId || typeof messageId !== "string") {
      return res.status(400).json({
        error: "Bad request",
        message: "Missing or invalid messageId query parameter.",
      });
    }
    const folderId = typeof req.query.folderId === "string" ? req.query.folderId : "";
    const token = await getAccessToken();
    const list = await fetchAttachmentCollection(token, messageId, folderId);
    const attachments = list.map((a) => {
      const odataType = a["@odata.type"] || "";
      return {
        id: a.id,
        name: a.name,
        contentType: a.contentType,
        size: a.size,
        isInline: a.isInline,
        downloadable: odataType.includes("fileAttachment"),
      };
    });
    res.json({ mailbox: MAILBOX_EMAIL, messageId, attachments });
  } catch (err) {
    console.error("Attachments list error:", err.message);
    res.status(500).json({
      error: "Failed to load attachments",
      message: err.message,
    });
  }
});

// Binary content for a file attachment (proxied for same-origin viewing in the browser)
// GET /api/mail/attachment-content?messageId=...&attachmentId=...&folderId=...
app.get("/api/mail/attachment-content", async (req, res) => {
  try {
    const messageId = req.query.messageId;
    const attachmentId = req.query.attachmentId;
    if (
      !messageId ||
      typeof messageId !== "string" ||
      !attachmentId ||
      typeof attachmentId !== "string"
    ) {
      return res.status(400).json({
        error: "Bad request",
        message:
          "Missing messageId or attachmentId query parameter.",
      });
    }
    const folderId = typeof req.query.folderId === "string" ? req.query.folderId : "";
    if (
      String(attachmentId).includes(":stmt-href:") ||
      String(attachmentId).includes(":stmt-synth:")
    ) {
      return res.status(400).json({
        error: "Not a mailbox file attachment",
        message:
          "This row is generated from a statement link, not a PDF in the email. Open the tax-invoice link from the source email or add the PDF as an attachment.",
      });
    }
    const token = await getAccessToken();
    const { meta, metaPath } = await resolveAttachmentMeta(
      token,
      messageId,
      attachmentId,
      folderId,
    );
    const odataType = meta["@odata.type"] || "";
    if (!odataType.includes("fileAttachment")) {
      return res.status(415).json({
        error: "Unsupported attachment type",
        message:
          "Only file attachments can be opened here. Open the message in Outlook for embedded items.",
      });
    }

    const valueUrl = `${GRAPH_BASE}${metaPath}/$value`;
    const streamRes = await graphFetchOkResponse(token, valueUrl);

    const contentType =
      meta.contentType ||
      streamRes.headers.get("content-type") ||
      "application/octet-stream";
    const filename = meta.name || "attachment";
    const ct = contentType.split(";")[0].trim();
    res.setHeader("Content-Type", ct);
    const disp = preferInlineDisposition(ct) ? "inline" : "attachment";
    res.setHeader(
      "Content-Disposition",
      `${disp}; filename*=UTF-8''${encodeURIComponent(filename)}`,
    );

    if (streamRes.body && typeof Readable.fromWeb === "function") {
      Readable.fromWeb(streamRes.body).pipe(res);
    } else {
      const buf = Buffer.from(await streamRes.arrayBuffer());
      res.send(buf);
    }
  } catch (err) {
    console.error("Attachment content error:", err.message);
    if (!res.headersSent) {
      res.status(500).json({
        error: "Failed to load attachment",
        message: err.message,
      });
    }
  }
});

// Messages in a specific folder (folderId can be "inbox" or a child folder id)
app.get("/api/mail/folders/:folderId/messages", async (req, res) => {
  try {
    const folderId = req.params.folderId;
    const limit = Math.min(parseInt(req.query.limit, 10) || 200, 500);
    const token = await getAccessToken();
    const messages = await fetchFolderMessages(token, folderId, limit);
    res.json({ folderId, mailbox: MAILBOX_EMAIL, messages });
  } catch (err) {
    console.error("Folder messages error:", err.message);
    res.status(500).json({
      error: "Failed to load folder messages",
      message: err.message,
    });
  }
});

// SSE: progress while running domain sorter — used by dashboard Run sorter button
// GET /api/mail/run-sort/stream
app.get("/api/mail/run-sort/stream", async (req, res) => {
  res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no");
  if (typeof res.flushHeaders === "function") res.flushHeaders();

  try {
    const token = await getAccessToken();
    const result = await runMailSort(token, (p) => sseWrite(res, "progress", p));
    sseWrite(res, "complete", {
      mailbox: MAILBOX_EMAIL,
      report: result.report,
      message: result.message,
    });
    res.end();
  } catch (err) {
    console.error("Run-sort stream error:", err.message);
    sseWrite(res, "fatal", { message: err.message });
    res.end();
  }
});

// Run sorter: allocate inbox by domain, correct mis-filed, return report (no progress events)
app.post("/api/mail/run-sort", async (req, res) => {
  try {
    const token = await getAccessToken();
    const { report, message } = await runMailSort(token);
    res.json({
      mailbox: MAILBOX_EMAIL,
      report,
      message,
    });
  } catch (err) {
    console.error("Run-sort error:", err.message);
    res.status(500).json({
      error: "Failed to run sorter",
      message: err.message,
    });
  }
});

// Health / config check (no secret exposed)
app.get("/api/health", (req, res) => {
  res.json({
    ok: !!(CLIENT_ID && CLIENT_SECRET && TENANT_ID),
    mailbox: MAILBOX_EMAIL,
  });
});

// Serve dashboard
app.use(express.static(path.join(__dirname, "public")));
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.listen(PORT, LISTEN_HOST, () => {
  const hostLabel = LISTEN_HOST === "0.0.0.0" ? "localhost (all interfaces)" : LISTEN_HOST;
  console.log(`FlightBox Mail API listening on ${LISTEN_HOST}:${PORT} (${hostLabel})`);
  console.log(`Dashboard: http://localhost:${PORT}/`);
  console.log(`Document details (JSON): GET http://localhost:${PORT}/api/mail/document-details?folderId=inbox`);
  console.log(`Document details cache: GET http://localhost:${PORT}/api/mail/document-details/cached?folderId=inbox`);
  console.log(`Document details (SSE):  GET http://localhost:${PORT}/api/mail/document-details/stream?folderId=inbox`);
  console.log(`Run sorter (SSE):       GET http://localhost:${PORT}/api/mail/run-sort/stream`);
  console.log(
    `Document details scan depth: ${DOCUMENT_DETAILS_MESSAGE_LIMIT_DEFAULT} msgs/folder (cap ${DOCUMENT_DETAILS_MESSAGE_LIMIT_MAX}, env DOCUMENT_DETAILS_FOLDER_MESSAGE_LIMIT)`,
  );
});
