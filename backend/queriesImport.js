import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import * as XLSX from "xlsx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const QUERIES_IMPORT_PATH = path.join(__dirname, "data", "queries-import.json");

export function normalizeLookupKey(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

/**
 * Find Excel headers for merge: Query Number ↔ Document no (matches Document details "Doc no.").
 * @param {string[]} headers
 */
/** Strip trailing punctuation so "Document No." still matches token checks. */
function headerTokens(s) {
  return normalizeLookupKey(String(s || "").replace(/[.:;,\s]+$/g, ""));
}

export function detectQueryMergeColumns(headers) {
  const list = (headers || []).map((h) => String(h || "").trim());
  let queryNumberHeader = "";
  let documentNoHeader = "";

  for (const raw of list) {
    if (!raw) continue;
    const n = headerTokens(raw);
    if (!documentNoHeader) {
      if (/^document\s*(?:no\.?|#|number)\.?$/i.test(raw.trim())) {
        documentNoHeader = raw;
        continue;
      }
      if (
        n.includes("document") &&
        (n.includes("no") || n.includes("number") || n.includes("#")) &&
        !n.includes("query")
      ) {
        documentNoHeader = raw;
        continue;
      }
    }
  }

  for (const raw of list) {
    if (!raw) continue;
    const n = headerTokens(raw);
    if (!queryNumberHeader) {
      if (/^query\s*(?:no\.?|#|number)\.?$/i.test(raw.trim())) {
        queryNumberHeader = raw;
        continue;
      }
      if (
        (n.includes("query") || n.includes("enquiry")) &&
        (n.includes("no") || n.includes("number") || n.includes("#"))
      ) {
        queryNumberHeader = raw;
        continue;
      }
    }
  }

  return { queryNumberHeader, documentNoHeader };
}

/**
 * Reports often have title rows; real headers may start on row 5+.
 * Picks the first row (within scan limit) where both merge columns are detectable.
 * @param {string[][]} matrix
 * @param {number} maxScanRows
 */
export function findQueriesHeaderRowIndex(matrix, maxScanRows = 50) {
  const limit = Math.min(matrix.length, maxScanRows);
  for (let r = 0; r < limit; r++) {
    const headers = matrix[r].map((c) => String(c ?? "").trim());
    const { queryNumberHeader, documentNoHeader } = detectQueryMergeColumns(headers);
    if (queryNumberHeader && documentNoHeader) return r;
  }
  return -1;
}

function uniqueColumnKeys(headerCells) {
  const keys = [];
  const counts = new Map();
  for (let i = 0; i < headerCells.length; i++) {
    let base = String(headerCells[i] ?? "").trim();
    if (!base) base = `__col_${i + 1}`;
    const n = (counts.get(base) || 0) + 1;
    counts.set(base, n);
    keys.push(n === 1 ? base : `${base}__${n}`);
  }
  return keys;
}

/**
 * First sheet → array of row objects (keys from detected header row).
 * Skips title rows when headers are not on row 1 (e.g. SAP-style reports).
 * @param {Buffer} buffer
 */
export function parseQueriesExcelBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) return { columns: [], rows: [] };
  const sheet = wb.Sheets[sheetName];

  let numCols = 0;
  if (sheet["!ref"]) {
    const ref = XLSX.utils.decode_range(sheet["!ref"]);
    numCols = Math.max(0, ref.e.c - ref.s.c + 1);
  }

  const rawMatrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
  });
  if (!rawMatrix.length) return { columns: [], rows: [] };

  let width = numCols;
  for (const row of rawMatrix) width = Math.max(width, row.length);
  const matrix = rawMatrix.map((row) => {
    const out = row.slice();
    while (out.length < width) out.push("");
    return out;
  });

  let headerRowIndex = findQueriesHeaderRowIndex(matrix);
  if (headerRowIndex < 0) {
    headerRowIndex = 0;
  }

  const headerCells = matrix[headerRowIndex].map((c) => String(c ?? "").trim());
  const columns = uniqueColumnKeys(headerCells);

  const normalizedRows = [];
  for (let r = headerRowIndex + 1; r < matrix.length; r++) {
    const row = matrix[r];
    const o = {};
    let any = false;
    for (let c = 0; c < columns.length; c++) {
      const v = row[c];
      const s = v == null ? "" : String(v).trim();
      if (s) any = true;
      o[columns[c]] = s;
    }
    if (any) normalizedRows.push(o);
  }

  return { columns, rows: normalizedRows };
}

export async function readQueriesImport() {
  try {
    const raw = await fs.readFile(QUERIES_IMPORT_PATH, "utf8");
    const data = JSON.parse(raw);
    return {
      columns: Array.isArray(data.columns) ? data.columns : [],
      rows: Array.isArray(data.rows) ? data.rows : [],
      queryNumberHeader: data.queryNumberHeader || "",
      documentNoHeader: data.documentNoHeader || "",
      importedAt: data.importedAt || null,
      fileName: data.fileName || "",
    };
  } catch {
    return {
      columns: [],
      rows: [],
      queryNumberHeader: "",
      documentNoHeader: "",
      importedAt: null,
      fileName: "",
    };
  }
}

export async function writeQueriesImport(payload) {
  await fs.mkdir(path.dirname(QUERIES_IMPORT_PATH), { recursive: true });
  await fs.writeFile(QUERIES_IMPORT_PATH, JSON.stringify(payload), "utf8");
}
