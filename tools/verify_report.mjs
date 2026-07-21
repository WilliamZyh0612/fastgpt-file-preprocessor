import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const [input, preview] = process.argv.slice(2);
if (!input || !preview) throw new Error("用法：node verify_report.mjs <xlsx> <preview.png>");
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
const inspection = await workbook.inspect({
  kind: "table",
  range: "知识库总目录!A1:P6",
  include: "values,formulas",
  tableMaxRows: 6,
  tableMaxCols: 16,
});
console.log(inspection.ndjson);
const image = await workbook.render({ sheetName: "知识库总目录", range: "A1:H3", scale: 1, format: "png" });
await fs.writeFile(preview, new Uint8Array(await image.arrayBuffer()));
