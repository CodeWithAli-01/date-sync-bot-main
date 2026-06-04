import ExcelJS from "exceljs";

export interface DoctorCoverageResult {
  totalEmployees: number;
  matchedEmployees: number;
  unmatchedEmployees: string[];
  warnings: string[];
  debug: {
    sourceRows: number;
    templateRows: number;
  };
  blob: Blob;
  fileName: string;
  sheetName: string;
  preview: { name: string; total: number }[];
}

interface CoverageRow {
  code: string;
  name: string;
  targetDoctors: number;
  coveredDoctors: number;
  coveragePercent: string;
}

interface TemplateColumns {
  codeCol: number;
  nameCol: number;
  headerRow: number;
  targetDoctorsCol: number;
  coveredDoctorsCol: number;
  coveragePercentCol: number;
}

const SOURCE_HINTS: Record<string, string[]> = {
  code: ["employee code", "emp code", "employee id", "emp id"],
  name: ["employee name", "name"],
  targetDoctors: ["target doctors"],
  coveredDoctors: ["covered doctors"],
  coveragePercent: ["coverage %", "coverage percent", "coverage"],
};

const TEMPLATE_HINTS: Record<string, string[]> = {
  code: ["employee code", "emp code", "employee id", "emp id"],
  name: ["name", "employee name"],
};

export async function processDoctorCoverageReport(
  sourceFile: File,
  templateFile: File,
): Promise<DoctorCoverageResult> {
  const sourceRows = await readCoverageRows(sourceFile);
  const sourceByCode = new Map(sourceRows.map((row) => [row.code, row]));
  const sourceByName = new Map(sourceRows.map((row) => [personNameKey(row.name), row]));
  const unmatched = new Set(sourceRows.map((row) => `${row.code} ${row.name}`));

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await templateFile.arrayBuffer());
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("Sample workbook does not contain a sheet.");

  const columns = ensureTemplateColumns(sheet);
  const preview: { name: string; total: number }[] = [];
  let matchedEmployees = 0;
  let templateRows = 0;

  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const code = normalizeEmployeeCode(row.getCell(columns.codeCol).text);
    const name = row.getCell(columns.nameCol).text;
    if (!code && !personNameKey(name)) continue;
    templateRows++;

    const match = (code && sourceByCode.get(code)) || sourceByName.get(personNameKey(name));
    if (!match) {
      clearCoverageCells(row, columns);
      continue;
    }

    matchedEmployees++;
    unmatched.delete(`${match.code} ${match.name}`);
    row.getCell(columns.targetDoctorsCol).value = match.targetDoctors;
    row.getCell(columns.coveredDoctorsCol).value = match.coveredDoctors;
    row.getCell(columns.coveragePercentCol).value = match.coveragePercent;
    styleCoverageCells(row, columns);
    preview.push({ name: name || match.name, total: match.coveredDoctors });
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  return {
    totalEmployees: templateRows,
    matchedEmployees,
    unmatchedEmployees: [...unmatched],
    warnings: unmatched.size
      ? [`${unmatched.size} coverage row(s) were not found in the sample file.`]
      : [],
    debug: {
      sourceRows: sourceRows.length,
      templateRows,
    },
    blob,
    fileName: templateFile.name.replace(/\.xlsx$/i, "") + " - Doctor Coverage Report.xlsx",
    sheetName: sheet.name,
    preview: preview.sort((a, b) => b.total - a.total).slice(0, 10),
  };
}

async function readCoverageRows(file: File): Promise<CoverageRow[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("Doctor coverage workbook does not contain a sheet.");

  const header = sheet.getRow(1);
  const labels = new Map<string, number>();
  for (let c = 1; c <= sheet.columnCount; c++) labels.set(normalize(header.getCell(c).text), c);

  const codeCol = findHeader(labels, SOURCE_HINTS.code);
  const nameCol = findHeader(labels, SOURCE_HINTS.name);
  const targetCol = findHeader(labels, SOURCE_HINTS.targetDoctors);
  const coveredCol = findHeader(labels, SOURCE_HINTS.coveredDoctors);
  const coverageCol = findHeader(labels, SOURCE_HINTS.coveragePercent);
  const rows: CoverageRow[] = [];

  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const code = normalizeEmployeeCode(row.getCell(codeCol).text);
    if (!code) continue;
    rows.push({
      code,
      name: row.getCell(nameCol).text.trim(),
      targetDoctors: numericCell(row.getCell(targetCol)),
      coveredDoctors: numericCell(row.getCell(coveredCol)),
      coveragePercent: coverageText(row.getCell(coverageCol)),
    });
  }

  return rows;
}

function ensureTemplateColumns(sheet: ExcelJS.Worksheet): TemplateColumns {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const labels = new Map<string, number>();
    for (let c = 1; c <= sheet.columnCount; c++) labels.set(normalize(row.getCell(c).text), c);

    const codeCol = findHeader(labels, TEMPLATE_HINTS.code, false);
    const nameCol = findHeader(labels, TEMPLATE_HINTS.name, false);
    if (!codeCol || !nameCol) continue;

    const lastUsedCol = Math.max(findLastUsedColumn(sheet), nameCol);
    const targetDoctorsCol =
      findHeader(labels, SOURCE_HINTS.targetDoctors, false) || lastUsedCol + 1;
    const coveredDoctorsCol =
      findHeader(labels, SOURCE_HINTS.coveredDoctors, false) ||
      Math.max(lastUsedCol + 1, targetDoctorsCol + 1);
    const coveragePercentCol =
      findHeader(labels, SOURCE_HINTS.coveragePercent, false) ||
      Math.max(lastUsedCol + 1, targetDoctorsCol, coveredDoctorsCol) + 1;

    const titleRow = sheet.getRow(Math.max(1, rowNumber - 1));
    applyCoverageGroupTitle(sheet, titleRow, [
      targetDoctorsCol,
      coveredDoctorsCol,
      coveragePercentCol,
    ]);

    const headers: Array<[number, string]> = [
      [targetDoctorsCol, "Target Doctors"],
      [coveredDoctorsCol, "Covered Doctors"],
      [coveragePercentCol, "Coverage %"],
    ];
    for (const [col, label] of headers) {
      const cell = row.getCell(col);
      cell.value = label;
      styleCoverageHeaderCell(cell);
      sheet.getColumn(col).width = label.length + 4;
    }
    titleRow.commit();
    row.commit();

    return {
      codeCol,
      nameCol,
      headerRow: rowNumber,
      targetDoctorsCol,
      coveredDoctorsCol,
      coveragePercentCol,
    };
  }

  throw new Error("Sample columns were not found. Required: Employee Code and Name.");
}

function applyCoverageGroupTitle(sheet: ExcelJS.Worksheet, titleRow: ExcelJS.Row, columns: number[]) {
  const labels = ["Target Doctors", "Covered Doctors", "Coverage %"];
  for (const [index, col] of columns.entries()) {
    const cell = titleRow.getCell(col);
    cell.value = labels[index] ?? "Doctor Coverage";
    styleCoverageHeaderCell(cell);
  }
  titleRow.height = Math.max(Number(titleRow.height) || 0, 22);
}

function findLastUsedColumn(sheet: ExcelJS.Worksheet): number {
  let lastUsedCol = 0;
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (cell.text.trim()) lastUsedCol = Math.max(lastUsedCol, colNumber);
    });
  }
  return lastUsedCol || sheet.columnCount;
}

function styleCoverageHeaderCell(cell: ExcelJS.Cell) {
  cell.font = { ...(cell.font ?? {}), bold: true, color: { argb: "FF000000" } };
  cell.alignment = { ...(cell.alignment ?? {}), horizontal: "center", vertical: "middle" };
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF00" },
  };
  cell.border = thinBorder();
}

function clearCoverageCells(row: ExcelJS.Row, columns: TemplateColumns) {
  row.getCell(columns.targetDoctorsCol).value = null;
  row.getCell(columns.coveredDoctorsCol).value = null;
  row.getCell(columns.coveragePercentCol).value = null;
}

function styleCoverageCells(row: ExcelJS.Row, columns: TemplateColumns) {
  for (const col of [columns.targetDoctorsCol, columns.coveredDoctorsCol, columns.coveragePercentCol]) {
    const cell = row.getCell(col);
    cell.alignment = { ...(cell.alignment ?? {}), horizontal: "center", vertical: "middle" };
    cell.border = thinBorder();
  }
}

function thinBorder(): Partial<ExcelJS.Borders> {
  return {
    top: { style: "thin", color: { argb: "FF111827" } },
    left: { style: "thin", color: { argb: "FF111827" } },
    bottom: { style: "thin", color: { argb: "FF111827" } },
    right: { style: "thin", color: { argb: "FF111827" } },
  };
}

function findHeader(headers: Map<string, number>, hints: string[], required = true): number {
  for (const hint of hints) {
    const exact = headers.get(normalize(hint));
    if (exact) return exact;
  }
  for (const [label, col] of headers) {
    if (hints.some((hint) => label.includes(normalize(hint)))) return col;
  }
  if (!required) return 0;
  throw new Error(`Required column not found: ${hints[0]}`);
}

function normalize(value: string): string {
  return String(value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
}

function normalizeEmployeeCode(value: string): string {
  const digits = String(value ?? "").match(/\d+/g)?.join("") ?? "";
  return digits.length >= 4 ? digits : "";
}

function personNameKey(value: string): string {
  return normalize(value)
    .replace(/\([^)]*\)/g, " ")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\b(mr|mrs|ms|dr|mio|asm|hos|dbu|smio|sm)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function numericCell(cell: ExcelJS.Cell): number {
  const value = cell.value;
  if (typeof value === "number") return value;
  const parsed = Number(String(cell.text ?? "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function coverageText(cell: ExcelJS.Cell): string {
  if (cell.text) return cell.text.trim();
  const value = cell.value;
  if (typeof value === "number") return `${Math.round(value * 100)}%`;
  return String(value ?? "").trim();
}
