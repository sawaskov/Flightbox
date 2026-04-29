/**
 * Writes diagnostic snapshots for Google Ireland DV360 PDFs (Credit Memo + Invoice).
 * Output: data/google-debugger.txt (overwritten on each parse).
 */
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { getGoogleCreditMemoDebugSnapshot, formatGoogleDebuggerTxt } from "./invoiceExtract.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DEBUG_PATH = path.join(__dirname, "data", "google-debugger.txt");

/**
 * @param {string} rawPdfText
 * @param {{ messageId?: string, attachmentId?: string, sourceFileName?: string, folderId?: string }} meta
 */
export async function writeGoogleExtractDebuggerFile(rawPdfText, meta = {}) {
  try {
    const snap = getGoogleCreditMemoDebugSnapshot(rawPdfText);
    const body = formatGoogleDebuggerTxt(snap, meta);
    await fs.mkdir(path.dirname(DEBUG_PATH), { recursive: true });
    await fs.writeFile(DEBUG_PATH, body, "utf8");
  } catch (err) {
    console.warn("[google-debugger] write failed:", err.message || err);
  }
}
