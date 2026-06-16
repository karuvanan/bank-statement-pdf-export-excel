import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = process.cwd();
const inputPath = path.join(root, "tmp", "pdfs", "extracted_transactions.json");
const outputDir = path.join(root, "outputs", "pdf_export_excel");
const outputPath = path.join(outputDir, "MBB_PDF_Export_Optimized.xlsx");

const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));
const rows = payload.rows;
const summaries = payload.statement_summaries;

function sheetName(name) {
  return name.replace(/\.pdf$/i, "").replace(/[\[\]:*?/\\]/g, " ").slice(0, 31);
}

function moneyOrBlank(value) {
  return value === null || value === undefined ? null : Number(value);
}

function excelDateOrBlank(value) {
  if (!value) return null;
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) return value;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function splitDescription(value) {
  const parts = String(value || "")
    .split(/\r?\n/)
    .map((part) => part.trim())
    .filter(Boolean);
  const desc = parts[0] || "";
  return {
    desc,
    details: parts.slice(1).join(" ") || desc,
  };
}

function txMatrix(txRows) {
  return txRows.map((row, index) => {
    const description = splitDescription(row.description);
    return [
      index + 1,
      row.source_pdf,
      row.page,
      excelDateOrBlank(row.entry_date_full),
      description.desc,
      description.details,
      moneyOrBlank(row.amount),
      moneyOrBlank(row.debit),
      moneyOrBlank(row.credit),
      moneyOrBlank(row.balance),
      row.amount_text,
      row.balance_text,
      row.notes,
    ];
  });
}

function setWidths(sheet) {
  const widths = [46, 110, 48, 105, 180, 350, 105, 105, 105, 115, 110, 110, 190];
  widths.forEach((width, idx) => {
    sheet.getRangeByIndexes(0, idx, 1, 1).format.columnWidthPx = width;
  });
}

function styleTransactionSheet(sheet, rowCount) {
  const cols = 13;
  sheet.freezePanes.freezeRows(1);
  sheet.getRangeByIndexes(0, 0, 1, cols).format = {
    fill: "#174A7C",
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
  };
  if (rowCount > 0) {
    const used = sheet.getRangeByIndexes(0, 0, rowCount + 1, cols);
    used.format.borders = { preset: "all", style: "thin", color: "#D9E2EC" };
    sheet.getRangeByIndexes(1, 3, rowCount, 1).format.numberFormat = "dd/mm/yyyy";
    sheet.getRangeByIndexes(1, 6, rowCount, 4).format.numberFormat = "#,##0.00";
    sheet.getRangeByIndexes(1, 4, rowCount, 2).format.wrapText = false;
    sheet.getRangeByIndexes(1, 12, rowCount, 1).format.wrapText = true;
    for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
      const rowNumber = rowIndex + 2;
      const desc = sheet.getCell(rowIndex + 1, 4).values?.[0]?.[0];
      const details = sheet.getCell(rowIndex + 1, 5).values?.[0]?.[0];
      const amount = sheet.getCell(rowIndex + 1, 6).values?.[0]?.[0];
      const isBeginningBalance = String(desc || "").trim().toUpperCase() === "BEGINNING BALANCE";
      if (!isBeginningBalance && desc && details && (amount === null || amount === undefined || amount === "")) {
        sheet.getRange(`G${rowNumber}:I${rowNumber}`).format.fill = "#FFF2CC";
      }
    }
  }
  setWidths(sheet);
}

function addTransactionSheet(workbook, name, txRows) {
  const sheet = workbook.worksheets.add(name);
  const headers = [
    "#",
    "Source PDF",
    "Page",
    "Date",
    "Desc",
    "Details Payee Reference",
    "Amount",
    "Debit",
    "Credit",
    "Statement Balance",
    "OCR Amount",
    "OCR Balance",
    "Review Notes",
  ];
  sheet.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
  const matrix = txMatrix(txRows);
  if (matrix.length) {
    sheet.getRangeByIndexes(1, 0, matrix.length, headers.length).values = matrix;
    sheet.tables.add(`A1:M${matrix.length + 1}`, true, `${name.replace(/[^A-Za-z0-9]/g, "") || "Transactions"}Table`);
  }
  styleTransactionSheet(sheet, matrix.length);
  return sheet;
}

const workbook = Workbook.create();

const summarySheet = workbook.worksheets.add("Summary");
summarySheet.getRange("A1:J1").values = [[
  "Source PDF",
  "Period",
  "Beginning Balance",
  "Ending Balance",
  "Transactions",
  "Calculated Debit",
  "Statement Total Debit",
  "Calculated Credit",
  "Statement Total Credit",
  "Rows With Notes",
]];
summarySheet.getRangeByIndexes(1, 0, summaries.length, 10).values = summaries.map((item) => [
  item.source_pdf,
  item.statement_period,
  moneyOrBlank(item.beginning_balance),
  moneyOrBlank(item.ending_balance),
  item.transaction_count,
  moneyOrBlank(item.calculated_debit),
  moneyOrBlank(item.statement_total_debit),
  moneyOrBlank(item.calculated_credit),
  moneyOrBlank(item.statement_total_credit),
  item.rows_with_notes,
]);
summarySheet.tables.add(`A1:J${summaries.length + 1}`, true, "StatementSummaryTable");
summarySheet.freezePanes.freezeRows(1);
summarySheet.getRange("A1:J1").format = {
  fill: "#174A7C",
  font: { bold: true, color: "#FFFFFF" },
  wrapText: true,
};
summarySheet.getRangeByIndexes(1, 2, summaries.length, 2).format.numberFormat = "#,##0.00";
summarySheet.getRangeByIndexes(1, 5, summaries.length, 4).format.numberFormat = "#,##0.00";
[140, 88, 132, 120, 95, 125, 138, 130, 142, 110].forEach((width, idx) => {
  summarySheet.getRangeByIndexes(0, idx, 1, 1).format.columnWidthPx = width;
});
summarySheet.getRangeByIndexes(0, 0, summaries.length + 1, 10).format.borders = {
  preset: "all",
  style: "thin",
  color: "#D9E2EC",
};

addTransactionSheet(workbook, "All Transactions", rows);

const byPdf = new Map();
for (const row of rows) {
  if (!byPdf.has(row.source_pdf)) byPdf.set(row.source_pdf, []);
  byPdf.get(row.source_pdf).push(row);
}
for (const [pdf, pdfRows] of byPdf.entries()) {
  addTransactionSheet(workbook, sheetName(pdf), pdfRows);
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const previewDir = path.join(outputDir, "previews");
await fs.mkdir(previewDir, { recursive: true });

for (const sheet of workbook.worksheets.items) {
  const preview = await workbook.render({
    sheetName: sheet.name,
    range: sheet.name === "Summary" ? "A1:J8" : "A1:M18",
    scale: 1,
    format: "png",
  });
  const previewName = `${sheet.name.replace(/[^A-Za-z0-9._-]+/g, "_")}.png`;
  await fs.writeFile(path.join(previewDir, previewName), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
