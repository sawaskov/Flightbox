/**
 * Persist parsed document-detail rows per folder so repeat visits avoid re-parsing every PDF.
 * Incremental sync uses Graph receivedDateTime gt last stored message time.
 */
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export const DOCUMENT_DETAILS_CACHE_VERSION = 1;

/** One file per extraction revision so bumping logic does not wipe other folders mid-migrate. */
export const DOCUMENT_DETAILS_EXTRACTOR_REVISION = 42;

const CACHE_PATH = path.join(
  __dirname,
  "data",
  `document-details-r${DOCUMENT_DETAILS_EXTRACTOR_REVISION}.json`,
);

export function attachmentRowKey(folderId, messageId, attachmentId) {
  return `${folderId}|${messageId}|${attachmentId}`;
}

export async function loadDocumentDetailsCacheFile() {
  try {
    const raw = await fs.readFile(CACHE_PATH, "utf8");
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

export async function saveDocumentDetailsCacheFile(data) {
  await fs.mkdir(path.dirname(CACHE_PATH), { recursive: true });
  await fs.writeFile(CACHE_PATH, JSON.stringify(data), "utf8");
}

export function mergeRowsByAttachmentKey(folderId, existingRows, incomingRows) {
  const map = new Map();
  for (const r of existingRows || []) {
    const mid = r.messageId;
    const aid = r.attachmentId;
    if (mid && aid) map.set(attachmentRowKey(folderId, mid, aid), r);
  }
  for (const r of incomingRows || []) {
    const mid = r.messageId;
    const aid = r.attachmentId;
    if (!mid || !aid) continue;
    map.set(attachmentRowKey(folderId, mid, aid), r);
  }
  return [...map.values()].sort((a, b) =>
    String(b.emailReceivedDateTime || "").localeCompare(String(a.emailReceivedDateTime || "")),
  );
}

export function maxEmailReceivedIso(rows) {
  let max = "";
  for (const r of rows || []) {
    const t = r.emailReceivedDateTime;
    if (typeof t === "string" && t && (!max || t > max)) max = t;
  }
  return max || null;
}

export async function readFolderRowsFromCache(folderId) {
  const data = await loadDocumentDetailsCacheFile();
  if (!data?.folders?.[folderId]?.rows) return [];
  return data.folders[folderId].rows;
}

export async function writeFolderRowsToCache(folderId, rows) {
  let data = await loadDocumentDetailsCacheFile();
  if (!data?.folders) {
    data = {
      version: DOCUMENT_DETAILS_CACHE_VERSION,
      extractorRevision: DOCUMENT_DETAILS_EXTRACTOR_REVISION,
      folders: {},
    };
  }
  data.folders[folderId] = {
    rows,
    savedAt: new Date().toISOString(),
    lastReceivedDateTime: maxEmailReceivedIso(rows),
  };
  await saveDocumentDetailsCacheFile(data);
}
