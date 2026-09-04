import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = process.argv[2];
const outputDir = process.argv[3];
if (!inputPath || !outputDir) throw new Error("Usage: verify-workbook.mjs <xlsx> <output-dir>");

await fs.mkdir(outputDir, { recursive: true });
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheets = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 12000 });
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 6000,
});

const sheetRecords = String(sheets.ndjson || "")
  .split(/\r?\n/)
  .filter(Boolean)
  .map((line) => JSON.parse(line))
  .filter((item) => item.name);
const renderResults = [];
for (let index = 0; index < sheetRecords.length; index += 1) {
  const sheetName = sheetRecords[index].name;
  try {
    const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 0.7, format: "png" });
    const filename = `${String(index + 1).padStart(2, "0")}_${sheetName.replace(/[\\/:*?"<>|]/g, "_")}.png`;
    const target = path.join(outputDir, filename);
    await fs.writeFile(target, new Uint8Array(await preview.arrayBuffer()));
    renderResults.push({ sheetName, status: "ok", target });
  } catch (error) {
    renderResults.push({ sheetName, status: "failed", error: String(error) });
  }
}

const report = {
  inputPath,
  sheetCount: sheetRecords.length,
  sheets: sheetRecords.map((item) => item.name),
  formulaErrors: String(errors.ndjson || ""),
  renderResults,
};
await fs.writeFile(path.join(outputDir, "verification.json"), JSON.stringify(report, null, 2), "utf8");
console.log(JSON.stringify(report, null, 2));
