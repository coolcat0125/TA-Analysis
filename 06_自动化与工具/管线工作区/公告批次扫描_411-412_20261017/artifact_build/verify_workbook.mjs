import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const root = path.resolve(process.cwd(), "..");
const file = path.join(
  root,
  "outputs",
  "01a147a8-938a-7ff3-a15c-94b7fcac7a5e",
  "411-412公告NEV车型扫描底表.xlsx",
);
const blob = await FileBlob.load(file);
const workbook = await SpreadsheetFile.importXlsx(blob);
workbook.recalculate();
const sheetOverview = await workbook.inspect({
  kind: "sheet",
  include: "id,name",
});
console.log(sheetOverview.ndjson);
const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 200 },
});
console.log(errorScan.ndjson || "no formula errors");
const summary = await workbook.inspect({
  kind: "table",
  sheetId: "扫描总览",
  range: "A5:K20",
  include: "values",
  tableMaxRows: 16,
  tableMaxCols: 11,
});
console.log(summary.ndjson);
const out = await SpreadsheetFile.exportXlsx(workbook);
await out.save(path.join(path.dirname(file), "verify.tmp.xlsx"));
console.log(`verified ${file}`);
