import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
import fs from "node:fs/promises";
import path from "node:path";

const outputDir = process.argv[2];
const names = ["FastGPT知识库总目录.xlsx", "产品选型矩阵.xlsx", "故障案例库.xlsx", "配件适配表.xlsx", "CRM_ERP业务规则.xlsx"];
for (const name of names) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(outputDir, name)));
  const check = await workbook.inspect({kind: "table", range: "数据!A1:L6", include: "values,formulas", tableMaxRows: 6, tableMaxCols: 12, tableMaxCellChars: 80});
  const image = await workbook.render({sheetName: "数据", range: "A1:F3", scale: 1, format: "png"});
  await fs.writeFile(path.join(outputDir, `${name}.png`), new Uint8Array(await image.arrayBuffer()));
  console.log(`${name}\n${check.ndjson}`);
}
