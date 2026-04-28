/**
 * Dump plain text from a PDF using the same pdf-parse stack as the API, then print extractInvoiceFields JSON.
 *
 * Usage (from the backend folder, use YOUR real filename — include .pdf):
 *
 *   node scripts/dump-pdf-text.mjs "DSTV Invoice no 23053632.pdf"
 *
 * Full path also works:
 *
 *   node scripts/dump-pdf-text.mjs "C:\Users\...\Desktop\FlightBox\backend\DSTV Invoice no 23053632.pdf"
 *
 * Redirect output:
 *   node scripts/dump-pdf-text.mjs "DSTV Invoice no 23053632.pdf" > dstv-debug.txt
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PDFParse } from "pdf-parse";
import { extractInvoiceFields } from "../invoiceExtract.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const pdfPath = process.argv[2];

if (!pdfPath) {
  console.error(
    'Usage: node scripts/dump-pdf-text.mjs "<your-file.pdf>"',
  );
  console.error(
    'Example: node scripts/dump-pdf-text.mjs "DSTV Invoice no 23053632.pdf"',
  );
  console.error("(The placeholder that-file.pdf was only an example — use the real PDF name.)");
  process.exit(1);
}

const abs = path.resolve(pdfPath);
if (!fs.existsSync(abs)) {
  console.error("File not found:", abs);
  process.exit(1);
}

const buf = fs.readFileSync(abs);
const parser = new PDFParse({ data: buf });
let text = "";
try {
  const result = await parser.getText();
  text = result?.text ? String(result.text) : "";
} finally {
  await parser.destroy();
}

const flat = text.replace(/\s+/g, " ");
const extracted = extractInvoiceFields(text);

console.log("=== FILE ===");
console.log(abs);
console.log("\n=== RAW TEXT (as returned by pdf-parse) ===");
console.log(text);
console.log("\n=== SINGLE-LINE FLAT ===");
console.log(flat);
console.log("\n=== extractInvoiceFields (same as Document details API) ===");
console.log(JSON.stringify(extracted, null, 2));
