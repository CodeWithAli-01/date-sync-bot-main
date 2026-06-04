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

interface CoverageSourceColumns {
  headerRow: number;
  codeCol: number;
  nameCol: number;
  targetDoctorsCol: number;
  coveredDoctorsCol: number;
  coveragePercentCol: number;
}

const SOURCE_HINTS: Record<string, string[]> = {
  code: ["employee code", "emp code", "employee id", "emp id", "emp. id", "code"],
  name: ["employee name", "employee", "name"],
  targetDoctors: ["target doctors", "target doctor", "target drs", "targets"],
  coveredDoctors: ["covered doctors", "covered doctor", "covered drs", "coverage doctors"],
  coveragePercent: [
    "coverage %",
    "coverage%",
    "coverage percent",
    "coverage percentage",
    "coverage",
  ],
};

const TEMPLATE_HINTS: Record<string, string[]> = {
  code: ["employee code", "emp code", "employee id", "emp id", "emp. id", "code"],
  name: ["name", "employee name"],
};

const COVERAGE_OUTPUT_HEADERS = {
  targetDoctors: "Target Doctors",
  coveredDoctors: "Covered Doctors",
  coveragePercent: "Coverage %",
} as const;

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

  for (let rowNumber = columns.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber++) {
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
  const sheet =
    workbook.worksheets.find((worksheet) => findCoverageSourceColumns(worksheet)) ??
    workbook.worksheets[0];
  if (!sheet) throw new Error("Doctor coverage workbook does not contain a sheet.");

  const columns = findCoverageSourceColumns(sheet);
  if (!columns) {
    throw new Error(
      "Doctor coverage columns were not found. Required: Employee Code/Name, Target Doctors, and Covered Doctors.",
    );
  }
  const rows: CoverageRow[] = [];

  for (let r = columns.headerRow + 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const code = normalizeEmployeeCode(rowText(row, columns.codeCol));
    const name = rowText(row, columns.nameCol).trim();
    if (!code && !personNameKey(name)) continue;
    const targetDoctors = numericCell(row.getCell(columns.targetDoctorsCol));
    const coveredDoctors = numericCell(row.getCell(columns.coveredDoctorsCol));
    const coveragePercent = columns.coveragePercentCol
      ? coverageText(row.getCell(columns.coveragePercentCol), targetDoctors, coveredDoctors)
      : formatCoveragePercent(targetDoctors, coveredDoctors);

    rows.push({
      code,
      name,
      targetDoctors,
      coveredDoctors,
      coveragePercent,
    });
  }

  return rows;
}

function rowText(row: ExcelJS.Row, columnNumber: number): string {
  return columnNumber > 0 ? cellText(row.getCell(columnNumber)) : "";
}

function findCoverageSourceColumns(sheet: ExcelJS.Worksheet): CoverageSourceColumns | null {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 20); rowNumber++) {
    const labels = headerLabels(sheet.getRow(rowNumber), sheet.columnCount);
    const codeCol = findHeader(labels, SOURCE_HINTS.code, false);
    const nameCol = findHeader(labels, SOURCE_HINTS.name, false);
    const targetDoctorsCol = findHeader(labels, SOURCE_HINTS.targetDoctors, false);
    const coveredDoctorsCol = findHeader(labels, SOURCE_HINTS.coveredDoctors, false);
    const coveragePercentCol = findHeader(labels, SOURCE_HINTS.coveragePercent, false);

    if ((codeCol || nameCol) && targetDoctorsCol && coveredDoctorsCol) {
      return {
        headerRow: rowNumber,
        codeCol,
        nameCol,
        targetDoctorsCol,
        coveredDoctorsCol,
        coveragePercentCol,
      };
    }
  }

  return null;
}

function ensureTemplateColumns(sheet: ExcelJS.Worksheet): TemplateColumns {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const labels = headerLabels(row, sheet.columnCount);

    const codeCol = findHeader(labels, TEMPLATE_HINTS.code, false);
    const nameCol = findHeader(labels, TEMPLATE_HINTS.name, false);
    if (!codeCol || !nameCol) continue;

    const lastUsedCol = Math.max(findLastUsedColumn(sheet), nameCol);
    const usedColumns = new Set<number>();
    const nextColumn = () => {
      let col = Math.max(lastUsedCol, ...usedColumns) + 1;
      while (usedColumns.has(col)) col++;
      usedColumns.add(col);
      return col;
    };

    const targetDoctorsCol = reserveOutputColumn(
      labels,
      SOURCE_HINTS.targetDoctors,
      usedColumns,
      nextColumn,
    );
    const coveredDoctorsCol = reserveOutputColumn(
      labels,
      SOURCE_HINTS.coveredDoctors,
      usedColumns,
      nextColumn,
    );
    const coveragePercentCol = reserveOutputColumn(
      labels,
      SOURCE_HINTS.coveragePercent,
      usedColumns,
      nextColumn,
    );

    const headers: Array<[number, string]> = [
      [targetDoctorsCol, COVERAGE_OUTPUT_HEADERS.targetDoctors],
      [coveredDoctorsCol, COVERAGE_OUTPUT_HEADERS.coveredDoctors],
      [coveragePercentCol, COVERAGE_OUTPUT_HEADERS.coveragePercent],
    ];
    for (const [col, label] of headers) {
      const cell = row.getCell(col);
      cell.value = label;
      styleCoverageHeaderCell(cell);
      sheet.getColumn(col).width = Math.max(sheet.getColumn(col).width ?? 0, label.length + 4);
    }
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

function reserveOutputColumn(
  labels: Map<string, number>,
  hints: string[],
  usedColumns: Set<number>,
  nextColumn: () => number,
): number {
  const existing = findHeader(labels, hints, false);
  if (existing && !usedColumns.has(existing)) {
    usedColumns.add(existing);
    return existing;
  }
  return nextColumn();
}

function headerLabels(row: ExcelJS.Row, columnCount: number): Map<string, number> {
  const labels = new Map<string, number>();
  for (let c = 1; c <= Math.max(columnCount, row.cellCount); c++) {
    const normalized = normalizeHeader(cellText(row.getCell(c)));
    if (normalized && !labels.has(normalized)) labels.set(normalized, c);
  }
  return labels;
}

function findLastUsedColumn(sheet: ExcelJS.Worksheet): number {
  let lastUsedCol = 0;
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (cellText(cell).trim()) lastUsedCol = Math.max(lastUsedCol, colNumber);
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
  for (const col of [
    columns.targetDoctorsCol,
    columns.coveredDoctorsCol,
    columns.coveragePercentCol,
  ]) {
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
  const normalizedHints = hints.map(normalizeHeader);
  for (const hint of hints) {
    const exact = headers.get(normalizeHeader(hint));
    if (exact) return exact;
  }
  for (const [label, col] of headers) {
    if (normalizedHints.some((hint) => label.includes(hint) || hint.includes(label))) return col;
  }
  if (!required) return 0;
  throw new Error(`Required column not found: ${hints[0]}`);
}

function normalize(value: string): string {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function normalizeHeader(value: string): string {
  return normalize(value)
    .replace(/%/g, " percent ")
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\b(dr|drs|doctor)\b/g, "doctors")
    .replace(/\s+/g, " ")
    .trim();
}

function cellText(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (value == null) return cell.text ?? "";
  if (typeof value === "object") {
    if ("text" in value) return String((value as { text?: unknown }).text ?? "");
    if ("result" in value) return String((value as { result?: unknown }).result ?? "");
    if ("richText" in value && Array.isArray((value as { richText?: unknown }).richText)) {
      return (value as { richText: Array<{ text?: string }> }).richText
        .map((part) => part.text ?? "")
        .join("");
    }
  }
  return cell.text || String(value);
}

function normalizeEmployeeCode(value: string): string {
  const digits =
    String(value ?? "")
      .match(/\d+/g)
      ?.join("") ?? "";
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
  const parsed = Number(cellText(cell).replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function coverageText(cell: ExcelJS.Cell, targetDoctors: number, coveredDoctors: number): string {
  const text = cellText(cell).trim();
  if (text) return text;
  const value = cell.value;
  if (typeof value === "number") return `${Math.round(value * 100)}%`;
  return formatCoveragePercent(targetDoctors, coveredDoctors);
}

function formatCoveragePercent(targetDoctors: number, coveredDoctors: number): string {
  if (!targetDoctors) return "";
  return `${Math.round((coveredDoctors / targetDoctors) * 100)}%`;
}
