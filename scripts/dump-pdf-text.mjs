/**
 * Wrapper: runs the backend PDF dump script so you can invoke from the FlightBox folder.
 *
 * Usage (from Desktop\FlightBox):
 *
 *   node scripts/dump-pdf-text.mjs googlecreditmemo.pdf > google-debugger.txt
 *
 * Paths may be relative to the current working directory or absolute.
 */
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const backendScript = path.join(__dirname, "..", "backend", "scripts", "dump-pdf-text.mjs");
const arg = process.argv[2];

if (!arg) {
  console.error("Usage: node scripts/dump-pdf-text.mjs \"<file.pdf>\"");
  console.error("Example: node scripts/dump-pdf-text.mjs googlecreditmemo.pdf > google-debugger.txt");
  process.exit(1);
}

if (!fs.existsSync(backendScript)) {
  console.error("Expected backend script at:", backendScript);
  process.exit(1);
}

const absPdf = path.isAbsolute(arg) ? arg : path.resolve(process.cwd(), arg);
if (!fs.existsSync(absPdf)) {
  console.error("File not found:", absPdf);
  process.exit(1);
}

const r = spawnSync(process.execPath, [backendScript, absPdf], { stdio: "inherit" });
process.exit(r.status === null ? 1 : r.status);
