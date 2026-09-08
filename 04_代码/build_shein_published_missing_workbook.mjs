import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [, , jsonPath, outputDirArg] = process.argv;
if (!jsonPath || !outputDirArg) {
  console.error("Usage: node build_shein_published_missing_workbook.mjs <stats.json> <output_dir>");
  process.exit(1);
}

const raw = await fs.readFile(jsonPath, "utf8");
const data = JSON.parse(raw);
const outputDir = path.resolve(outputDirArg);
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();

function safeValue(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "number" || typeof value === "boolean") return value;
  let text = String(value);
  if (text.length > 32760) text = `${text.slice(0, 32757)}...`;
  if (/^[=+\-@]/.test(text)) return `'${text}`;
  return text;
}

function safeDetailValue(header, value) {
  if (header === "en" && value) {
    return safeValue(String(value).slice(0, 300));
  }
  return safeValue(value);
}

function colName(index1) {
  let n = index1;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function writeMatrix(sheet, startRow0, startCol0, rows, chunkSize = 5000) {
  for (let offset = 0; offset < rows.length; offset += chunkSize) {
    const chunk = rows.slice(offset, offset + chunkSize);
    const rowCount = chunk.length;
    const colCount = chunk[0]?.length ?? 0;
    if (!rowCount || !colCount) continue;
    sheet.getRangeByIndexes(startRow0 + offset, startCol0, rowCount, colCount).values = chunk;
  }
}

function applyHeaderStyle(range) {
  range.format = {
    fill: "#1F4E79",
    font: { bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
}

function addTableIfPossible(sheet, rowCount, colCount, tableName) {
  if (rowCount < 1 || colCount < 1) return;
  const address = `A1:${colName(colCount)}${rowCount}`;
  try {
    const table = sheet.tables.add(address, true, tableName);
    table.showFilterButton = true;
    table.showBandedRows = true;
  } catch (error) {
    console.warn(`table skipped for ${tableName}: ${error.message}`);
  }
}

function finishSheet(sheet, rows, cols, widthMap = {}, options = {}) {
  sheet.showGridLines = false;
  if (rows > 1) sheet.freezePanes.freezeRows(1);
  if (options.fullBorders !== false) {
    const used = sheet.getRangeByIndexes(0, 0, Math.max(rows, 1), cols);
    used.format.borders = { preset: "inside", style: "thin", color: "#E5E7EB" };
  }
  for (const [colIndex0, width] of Object.entries(widthMap)) {
    sheet.getRangeByIndexes(0, Number(colIndex0), 1, 1).format.columnWidth = width;
  }
}

function writeTableSheet(name, headers, objects, tableName, widthMap = {}, options = {}) {
  const sheet = workbook.worksheets.add(name);
  const rows = [
    headers,
    ...objects.map((item) => headers.map((header) => safeValue(item[header]))),
  ];
  writeMatrix(sheet, 0, 0, rows);
  applyHeaderStyle(sheet.getRangeByIndexes(0, 0, 1, headers.length));
  finishSheet(sheet, rows.length, headers.length, widthMap, options);
  if (options.table !== false) {
    addTableIfPossible(sheet, rows.length, headers.length, tableName);
  }
  return sheet;
}

function writeLargeDetailSheet(name, headers, objects, widthMap = {}) {
  const sheet = workbook.worksheets.add(name);
  sheet.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
  applyHeaderStyle(sheet.getRangeByIndexes(0, 0, 1, headers.length));
  const chunkSize = 2500;
  for (let offset = 0; offset < objects.length; offset += chunkSize) {
    const chunk = objects
      .slice(offset, offset + chunkSize)
      .map((item) => headers.map((header) => safeDetailValue(header, item[header])));
    writeMatrix(sheet, 1 + offset, 0, chunk, chunkSize);
    if ((offset / chunkSize) % 10 === 0) {
      console.log(`detail rows written: ${Math.min(offset + chunkSize, objects.length)}/${objects.length}`);
    }
  }
  finishSheet(sheet, objects.length + 1, headers.length, widthMap, { fullBorders: false });
  return sheet;
}

const summarySheet = workbook.worksheets.add("汇总");
summarySheet.showGridLines = false;
summarySheet.getRange("A1:F1").merge();
summarySheet.getRange("A1").values = [["SHEIN 已发布类 Key 空翻译对比表"]];
summarySheet.getRange("A1").format = {
  fill: "#17324D",
  font: { bold: true, color: "#FFFFFF", size: 16 },
  horizontalAlignment: "left",
};

const summaryRows = [
  ["生成时间", data.generated_at],
  ["数据来源", "文案系统 Headless API 本次重新导出"],
  ["筛选范围", "SHEIN；前端key/提示语/错误码 3 个 tab"],
  ["发布状态", `${data.release_statuses.join(", ")}（${data.release_status_note}）`],
  ["统计排除", "en 为源文案不计入；zh 语言列不计入；zh-cn/zh-tw/zh-hk 按独立语言统计"],
  ["导出 key 总数", data.summary.exported_keys],
  ["存在空翻译 key 数", data.summary.missing_keys],
  ["空翻译单元格数", data.summary.missing_cells],
  ["参与统计语言数", data.language_count_excluding_en_zh],
  ["原始导出目录", data.raw_dir],
];
summarySheet.getRange("A3:B12").values = summaryRows.map((row) => row.map(safeValue));
summarySheet.getRange("A3:A12").format = { fill: "#EAF2F8", font: { bold: true } };
summarySheet.getRange("B8:B10").format.numberFormat = "#,##0";
summarySheet.getRange("A3:B12").format.borders = { preset: "all", style: "thin", color: "#D6DEE8" };
summarySheet.getRange("A3:A12").format.columnWidth = 20;
summarySheet.getRange("B3:B12").format.columnWidth = 88;
summarySheet.getRange("B3:B12").format.wrapText = true;

const tabSheet = writeTableSheet(
  "Tab对比",
  ["tab", "exported_keys", "missing_keys", "missing_key_rate", "missing_cells", "language_count_excluding_en_zh"],
  data.tab_summary,
  "TabSummaryTable",
  { 0: 16, 1: 16, 2: 18, 3: 18, 4: 18, 5: 28 },
);
tabSheet.getRange(`B2:C${data.tab_summary.length + 1}`).format.numberFormat = "#,##0";
tabSheet.getRange(`D2:D${data.tab_summary.length + 1}`).format.numberFormat = "0.00%";
tabSheet.getRange(`E2:F${data.tab_summary.length + 1}`).format.numberFormat = "#,##0";

const platformSheet = writeTableSheet(
  "平台导出明细",
  ["tab", "platform_label", "api_count", "status_1_count", "status_2_count", "rows", "missing_keys", "missing_cells"],
  data.platform_summary,
  "PlatformSummaryTable",
  { 0: 14, 1: 16, 2: 16, 3: 16, 4: 18, 5: 14, 6: 16, 7: 18 },
);
platformSheet.getRange(`C2:H${data.platform_summary.length + 1}`).format.numberFormat = "#,##0";

const langHeaders = ["language", "missing_key_count", "前端key_missing", "提示语_missing", "错误码_missing"];
const langSheet = writeTableSheet(
  "语种对比",
  langHeaders,
  data.language_summary,
  "LanguageSummaryTable",
  { 0: 14, 1: 20, 2: 18, 3: 18, 4: 18 },
);
langSheet.getRange(`B2:E${data.language_summary.length + 1}`).format.numberFormat = "#,##0";

const detailHeaders = [
  "tab",
  "platform",
  "keyId",
  "missing_lang_count",
  "missing_langs",
  "en",
];
const detailSheet = writeLargeDetailSheet(
  "缺失Key明细",
  detailHeaders,
  data.missing_details,
  { 0: 12, 1: 14, 2: 30, 3: 18, 4: 50, 5: 70 },
);
detailSheet.getRange(`D2:D${data.missing_details.length + 1}`).format.numberFormat = "#,##0";
detailSheet.getRange("E2:F200").format.wrapText = true;

const exportRows = data.exports.map((item) => ({
  label: item.label,
  tab: item.tab,
  platform_label: item.platform_label,
  type: item.type,
  brand: item.brand,
  platform: item.platform.join("|"),
  release_statuses: item.release_statuses.join("|"),
  api_count: item.api_count,
  status_1_count: item.status_1_count,
  status_2_count: item.status_2_count,
  rows: item.rows,
  file: item.file,
  task_id: item.task_id ?? "",
}));
const exportSheet = writeTableSheet(
  "导出清单",
  ["label", "tab", "platform_label", "type", "brand", "platform", "release_statuses", "api_count", "status_1_count", "status_2_count", "rows", "file", "task_id"],
  exportRows,
  "ExportManifestTable",
  { 0: 24, 1: 12, 2: 14, 3: 10, 4: 10, 5: 14, 6: 18, 7: 14, 8: 16, 9: 16, 10: 14, 11: 72, 12: 28 },
);
exportSheet.getRange(`H2:K${exportRows.length + 1}`).format.numberFormat = "#,##0";
exportSheet.getRange(`L2:L${exportRows.length + 1}`).format.wrapText = true;

for (const sheetName of ["汇总", "Tab对比", "平台导出明细", "语种对比", "导出清单"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(outputDir, `${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
}
const detailPreview = await workbook.render({ sheetName: "缺失Key明细", range: "A1:F25", scale: 1, format: "png" });
await fs.writeFile(path.join(outputDir, "缺失Key明细.png"), new Uint8Array(await detailPreview.arrayBuffer()));

const inspectSummary = await workbook.inspect({
  kind: "table",
  range: "汇总!A1:B12",
  include: "values,formulas",
  tableMaxRows: 20,
  tableMaxCols: 4,
});
console.log(inspectSummary.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const outputPath = path.join(outputDir, `shein已发布key空翻译对比表_${data.run_id}.xlsx`);
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(JSON.stringify({ outputPath, sheetCount: 6, missingKeys: data.summary.missing_keys }, null, 2));
