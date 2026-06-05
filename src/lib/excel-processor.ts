// Read the Active Members workbook and update only the pharma report cells needed.
// Employee rows and source fields such as City/Designation are preserved.
import ExcelJS from "exceljs";
import type { PdfParseResult } from "./pdf-extractor";

export interface ProcessOptions {
  pdfResults: PdfParseResult[];
}

export interface ProcessReport {
  totalEmployees: number;
  matchedEmployees: number;
  unmatchedNames: string[];
  warnings: string[];
  debug: {
    totalPdfsUploaded: number;
    totalEmployeesDetected: number;
    totalMatched: number;
    totalSkipped: number;
    totalRecordsInsertedUpdated: number;
    parsingErrors: number;
  };
  dates: string[];
  days: number[];
  blob: Blob;
  fileName: string;
  sheetName: string;
  preview: { name: string; total: number }[];
}

const NAME_HINTS = ["employee name", "employee", "name", "emp name", "staff name", "member name"];
const CODE_HINTS = ["employee code", "emp code", "code", "emp id", "employee id", "id"];
const REGION_HINTS = ["region", "zone", "area"];
const CITY_HINTS = ["city", "town", "territory", "headquarter", "hq"];
const DESIGNATION_HINTS = ["designation", "desig", "position", "title"];
const SPECIAL_TEXT = ["meeting", "new joining", "resigned"];

function normalize(s: string): string {
  return (s ?? "").toString().replace(/\s+/g, " ").trim().toLowerCase();
}

function fuzzyKey(s: string): string {
  return normalize(s)
    .replace(/[^\p{L}\p{N}\s]/gu, "")
    .replace(/\s+/g, " ");
}

function normalizeEmployeeCode(value: string | null): string | null {
  const digits =
    String(value ?? "")
      .match(/\d+/g)
      ?.join("") ?? "";
  return digits.length >= 3 ? digits : null;
}

function codeSuffixes(code: string): string[] {
  const out: string[] = [];
  for (const len of [4, 5]) {
    if (code.length > len) out.push(code.slice(-len));
  }
  return out;
}

function personNameKey(s: string): string {
  const noise = new Set([
    "mr",
    "mrs",
    "ms",
    "dr",
    "psv",
    "mio",
    "asm",
    "rsm",
    "sm",
    "am",
    "fm",
    "tm",
    "swt",
    "hq",
    "territory",
    "location",
    "locations",
    "new",
    "joining",
    "resigned",
    "meeting",
    "grp",
    "with",
    "selfie",
    "selfies",
    "total",
  ]);
  return fuzzyKey(s.replace(/\b(new joining|resigned|meeting)\b.*$/i, " "))
    .split(" ")
    .filter((token) => token && !/^\d+$/.test(token) && !noise.has(token))
    .join(" ");
}

function levenshteinSimilarity(a: string, b: string): number {
  if (a === b) return 1;
  if (!a.length || !b.length) return 0;
  const previous = Array.from({ length: b.length + 1 }, (_, i) => i);
  const current = new Array<number>(b.length + 1);

  for (let i = 1; i <= a.length; i++) {
    current[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      current[j] = Math.min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + cost);
    }
    previous.splice(0, previous.length, ...current);
  }

  return 1 - previous[b.length] / Math.max(a.length, b.length);
}

function tokenSimilarity(a: string, b: string): number {
  const aTokens = a.split(" ").filter(Boolean);
  const bTokens = b.split(" ").filter(Boolean);
  if (!aTokens.length || !bTokens.length) return 0;
  let overlap = 0;
  for (const token of aTokens) {
    if (bTokens.some((candidate) => tokensMatch(token, candidate))) overlap++;
  }
  return overlap / Math.max(aTokens.length, bTokens.length);
}

function tokensMatch(a: string, b: string): boolean {
  if (a === b) return true;
  if (a.length === 1 && b.startsWith(a)) return true;
  if (b.length === 1 && a.startsWith(b)) return true;
  return false;
}

function orderedTokenCoverage(shorter: string, longer: string): number {
  const shortTokens = shorter.split(" ").filter(Boolean);
  const longTokens = longer.split(" ").filter(Boolean);
  if (!shortTokens.length || !longTokens.length) return 0;

  let longIndex = 0;
  let matched = 0;
  for (const token of shortTokens) {
    while (longIndex < longTokens.length && !tokensMatch(token, longTokens[longIndex])) {
      longIndex++;
    }
    if (longIndex >= longTokens.length) continue;
    matched++;
    longIndex++;
  }

  return matched / shortTokens.length;
}

function nameMatchScore(a: string, b: string): number {
  if (a === b) return 1;
  const shorter = a.length <= b.length ? a : b;
  const longer = a.length <= b.length ? b : a;
  const coverage = orderedTokenCoverage(shorter, longer);
  const tokenScore = tokenSimilarity(a, b);
  const editScore = levenshteinSimilarity(a, b);
  return Math.max(tokenScore, editScore, coverage >= 1 ? 0.94 : coverage * 0.9);
}

interface DetectedSheet {
  ws: ExcelJS.Worksheet;
  headerRow: number;
  nameCol: number;
  codeCol: number | null;
  regionCol: number | null;
  cityCol: number | null;
  designationCol: number | null;
  dataStartRow: number;
  dataEndRow: number;
}

interface Emp {
  row: number;
  name: string;
  code: string | null;
  nameKey: string;
  cleanNameKey: string;
}

interface DatePairCols {
  selfiesCol: number;
  totalCol: number;
  plannedCol?: number;
  unplannedCol?: number;
  callsCol?: number;
}

interface MonthlyTotalCols {
  callsCol: number;
  callsAvgCol: number;
  selfiesCol: number;
  selfiesAvgCol: number;
}

interface MatchValue {
  planned: number;
  unplanned: number;
  selfies: number;
  selfieText: string;
  total: number;
}

interface UnmatchedPdfRow {
  date: string;
  fileName: string;
  row: PdfParseResult["rows"][number];
}

function cellText(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (value == null) return "";
  if (value instanceof Date) return String(value.getDate());
  if (typeof value === "object") {
    if ("text" in value) return String(value.text ?? "");
    if ("result" in value) return String(value.result ?? "");
    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text ?? "").join("");
    }
    return "";
  }
  return String(value);
}

function cellDateKey(cell: ExcelJS.Cell): string | null {
  const value = cell.value;
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  const text = cellText(cell).trim();
  const iso = text.match(/\b(\d{4}-\d{2}-\d{2})\b/);
  if (iso) return iso[1];
  const short = text.match(
    /\b(\d{1,2})\s*[- ]\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b/i,
  );
  if (!short) return null;
  const months: Record<string, string> = {
    jan: "01",
    feb: "02",
    mar: "03",
    apr: "04",
    may: "05",
    jun: "06",
    jul: "07",
    aug: "08",
    sep: "09",
    oct: "10",
    nov: "11",
    dec: "12",
  };
  const month = months[short[2].slice(0, 3).toLowerCase()];
  const year = new Date().getFullYear();
  return month ? `${year}-${month}-${short[1].padStart(2, "0")}` : null;
}

function detectActiveSheet(wb: ExcelJS.Workbook): DetectedSheet {
  for (const ws of wb.worksheets) {
    const maxScan = Math.min(10, ws.rowCount || 10);
    for (let r = 1; r <= maxScan; r++) {
      const row = ws.getRow(r);
      let nameCol = 0;
      let codeCol = 0;
      let regionCol = 0;
      let cityCol = 0;
      let designationCol = 0;
      let hasPlanned = false;
      let hasUnplanned = false;
      let hasCalls = false;
      let hasSelfies = false;

      const maxC = Math.max(row.cellCount || 0, 80);
      for (let c = 1; c <= maxC; c++) {
        const v = normalize(cellText(row.getCell(c)));
        if (!v) continue;
        if (!regionCol && REGION_HINTS.includes(v)) regionCol = c;
        if (!codeCol && (CODE_HINTS.includes(v) || /\bcode\b/.test(v) || /\bemp.*id\b/.test(v))) {
          codeCol = c;
        }
        if (!nameCol && (NAME_HINTS.includes(v) || (v.includes("name") && !v.includes("file")))) {
          nameCol = c;
        }
        if (!cityCol && CITY_HINTS.includes(v)) cityCol = c;
        if (!designationCol && DESIGNATION_HINTS.includes(v)) designationCol = c;
        if (v === "planned") hasPlanned = true;
        if (v === "unplanned") hasUnplanned = true;
        if (v === "calls") hasCalls = true;
        if (v === "selfies") hasSelfies = true;
      }

      if (nameCol && hasPlanned && hasUnplanned && hasCalls && hasSelfies) {
        return {
          ws,
          headerRow: r,
          nameCol,
          codeCol: codeCol || null,
          regionCol: regionCol || null,
          cityCol: cityCol || null,
          designationCol: designationCol || null,
          dataStartRow: r + 1,
          dataEndRow: ws.rowCount || r + 1,
        };
      }
    }
  }

  for (const ws of wb.worksheets) {
    const maxScan = Math.min(10, ws.rowCount || 10);
    for (let r = 1; r <= maxScan; r++) {
      const row = ws.getRow(r);
      if (!row.cellCount) continue;

      let nameCol = 0;
      let codeCol = 0;
      let regionCol = 0;
      let cityCol = 0;
      let designationCol = 0;

      const maxC = Math.max(row.cellCount || 0, 40);
      for (let c = 1; c <= maxC; c++) {
        const v = normalize(cellText(row.getCell(c)));
        if (!v) continue;
        if (!regionCol && REGION_HINTS.includes(v)) regionCol = c;
        if (!codeCol && (CODE_HINTS.includes(v) || /\bcode\b/.test(v) || /\bemp.*id\b/.test(v))) {
          codeCol = c;
        }
        if (!nameCol && (NAME_HINTS.includes(v) || (v.includes("name") && !v.includes("file")))) {
          nameCol = c;
        }
        if (!cityCol && CITY_HINTS.includes(v)) cityCol = c;
        if (!designationCol && DESIGNATION_HINTS.includes(v)) designationCol = c;
      }

      if (!nameCol) continue;

      let hasValues = 0;
      for (let rr = r + 1; rr <= Math.min(r + 50, ws.rowCount || r + 50); rr++) {
        if (cellText(ws.getRow(rr).getCell(nameCol)).trim()) hasValues++;
      }
      if (hasValues >= 1) {
        return {
          ws,
          headerRow: r,
          nameCol,
          codeCol: codeCol || null,
          regionCol: regionCol || null,
          cityCol: cityCol || null,
          designationCol: designationCol || null,
          dataStartRow: r + 1,
          dataEndRow: ws.rowCount || r + 1,
        };
      }
    }
  }

  const ws = wb.worksheets[0];
  return {
    ws,
    headerRow: 1,
    nameCol: 3,
    codeCol: 2,
    regionCol: 1,
    cityCol: 4,
    designationCol: 5,
    dataStartRow: 2,
    dataEndRow: ws.rowCount || 2,
  };
}

function detectWorksheet(
  ws: ExcelJS.Worksheet,
  requireReportHeaders: boolean,
): DetectedSheet | null {
  const maxScan = Math.min(10, ws.rowCount || 10);
  for (let r = 1; r <= maxScan; r++) {
    const row = ws.getRow(r);
    if (!row.cellCount) continue;

    let nameCol = 0;
    let codeCol = 0;
    let regionCol = 0;
    let cityCol = 0;
    let designationCol = 0;
    let hasPlanned = false;
    let hasUnplanned = false;
    let hasCalls = false;
    let hasSelfies = false;

    const maxC = Math.max(row.cellCount || 0, requireReportHeaders ? 80 : 40);
    for (let c = 1; c <= maxC; c++) {
      const v = normalize(cellText(row.getCell(c)));
      if (!v) continue;
      if (!regionCol && REGION_HINTS.includes(v)) regionCol = c;
      if (!codeCol && (CODE_HINTS.includes(v) || /\bcode\b/.test(v) || /\bemp.*id\b/.test(v))) {
        codeCol = c;
      }
      if (!nameCol && (NAME_HINTS.includes(v) || (v.includes("name") && !v.includes("file")))) {
        nameCol = c;
      }
      if (!cityCol && CITY_HINTS.includes(v)) cityCol = c;
      if (!designationCol && DESIGNATION_HINTS.includes(v)) designationCol = c;
      if (v === "planned") hasPlanned = true;
      if (v === "unplanned") hasUnplanned = true;
      if (v === "calls") hasCalls = true;
      if (v === "selfies") hasSelfies = true;
    }

    if (!nameCol) continue;
    if (requireReportHeaders && !(hasPlanned && hasUnplanned && hasCalls && hasSelfies)) continue;

    let hasValues = 0;
    for (let rr = r + 1; rr <= Math.min(r + 50, ws.rowCount || r + 50); rr++) {
      if (cellText(ws.getRow(rr).getCell(nameCol)).trim()) hasValues++;
    }
    if (hasValues < 1) continue;

    return {
      ws,
      headerRow: r,
      nameCol,
      codeCol: codeCol || null,
      regionCol: regionCol || null,
      cityCol: cityCol || null,
      designationCol: designationCol || null,
      dataStartRow: r + 1,
      dataEndRow: ws.rowCount || r + 1,
    };
  }

  return null;
}

function detectReportSheets(wb: ExcelJS.Workbook): DetectedSheet[] {
  const sheets = wb.worksheets.filter((ws) => normalize(ws.name) !== "pdf extracted data");
  const reportSheets = sheets
    .map((ws) => detectWorksheet(ws, true))
    .filter((sheet): sheet is DetectedSheet => Boolean(sheet));
  if (reportSheets.length > 1) return reportSheets;

  return sheets
    .map((ws) => detectWorksheet(ws, false))
    .filter((sheet): sheet is DetectedSheet => Boolean(sheet));
}

function readEmployees(ws: ExcelJS.Worksheet, det: DetectedSheet): Emp[] {
  const employees: Emp[] = [];
  let lastDataRow = det.dataStartRow - 1;

  for (let r = det.dataStartRow; r <= Math.max(det.dataEndRow, det.dataStartRow + 1000); r++) {
    const row = ws.getRow(r);
    const nameVal = cellText(row.getCell(det.nameCol)).trim();
    const regionVal = det.regionCol ? normalize(cellText(row.getCell(det.regionCol))) : "";
    const designationVal = det.designationCol
      ? normalize(cellText(row.getCell(det.designationCol)))
      : "";
    if (regionVal === "pdf needs review" || designationVal === "unmatched pdf row") {
      lastDataRow = r;
      continue;
    }
    if (!nameVal) {
      if (r - lastDataRow > 5) break;
      continue;
    }
    const codeVal = det.codeCol ? normalizeEmployeeCode(cellText(row.getCell(det.codeCol))) : null;
    employees.push({
      row: r,
      name: nameVal,
      code: codeVal,
      nameKey: normalize(nameVal),
      cleanNameKey: personNameKey(nameVal),
    });
    lastDataRow = r;
  }

  return employees;
}

function findLastUsedColumn(ws: ExcelJS.Worksheet): number {
  let last = 0;
  ws.eachRow((row) => {
    row.eachCell({ includeEmpty: false }, (cell, col) => {
      if (cell.value != null && cellText(cell).trim() !== "") last = Math.max(last, col);
    });
  });
  return Math.max(last, ws.actualColumnCount || 0, 1);
}

function dateLabel(date: string): string {
  const parsed = new Date(`${date}T00:00:00`);
  const month = parsed.toLocaleString("en-US", { month: "short" });
  return `${parsed.getDate()}-${month}`;
}

function columnLetter(col: number): string {
  let letter = "";
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

function findDatePairColumns(
  ws: ExcelJS.Worksheet,
  headerRow: number,
  dates: string[],
): Map<string, DatePairCols> {
  const header = ws.getRow(headerRow);
  const dateHeader = ws.getRow(Math.max(1, headerRow - 1));
  const pairs = new Map<string, DatePairCols>();
  const lastUsed = findLastUsedColumn(ws);

  for (const date of dates) {
    for (let c = 1; c <= Math.max(lastUsed, header.cellCount || 0); c++) {
      if (cellDateKey(dateHeader.getCell(c)) !== date) continue;
      const labels = new Map<string, number>();
      for (let offset = 0; offset < 5; offset++) {
        const label = normalize(cellText(header.getCell(c + offset)));
        if (label) labels.set(label, c + offset);
      }
      const plannedCol = labels.get("planned");
      const unplannedCol = labels.get("unplanned");
      const callsCol = labels.get("calls");
      const selfiesCol = labels.get("selfies");
      if (plannedCol && unplannedCol && callsCol && selfiesCol) {
        pairs.set(date, {
          plannedCol,
          unplannedCol,
          callsCol,
          selfiesCol,
          totalCol: callsCol,
        });
        break;
      }
    }
  }

  for (const date of dates) {
    if (pairs.has(date)) continue;
    const selfieHeader = `${dateLabel(date)} Selfies`.toLowerCase();
    for (let c = 1; c <= Math.max(lastUsed, header.cellCount || 0); c++) {
      if (normalize(cellText(header.getCell(c))) !== selfieHeader) continue;
      const totalCol = c + 1;
      header.getCell(totalCol).value = "Total";
      pairs.set(date, { selfiesCol: c, totalCol });
      break;
    }
  }

  let nextCol = lastUsed + 1;
  for (const date of dates) {
    if (pairs.has(date)) continue;
    const plannedCol = nextCol;
    const unplannedCol = nextCol + 1;
    const totalCol = nextCol + 2;
    const selfiesCol = nextCol + 3;
    const dateCell = dateHeader.getCell(plannedCol);
    dateCell.value = new Date(`${date}T00:00:00`);
    dateCell.numFmt = "d-mmm";
    styleHeader(dateCell);
    header.getCell(plannedCol).value = "Planned";
    header.getCell(unplannedCol).value = "Unplanned";
    header.getCell(totalCol).value = "Calls";
    header.getCell(selfiesCol).value = "Selfies";
    for (const col of [plannedCol, unplannedCol, totalCol, selfiesCol]) {
      styleHeader(header.getCell(col));
      ws.getColumn(col).width = col === selfiesCol ? 10 : 11;
    }
    pairs.set(date, { plannedCol, unplannedCol, callsCol: totalCol, selfiesCol, totalCol });
    nextCol += 4;
  }

  header.commit?.();
  return pairs;
}

function findMonthlyTotalColumns(ws: ExcelJS.Worksheet, headerRow: number): MonthlyTotalCols {
  const dateHeader = ws.getRow(Math.max(1, headerRow - 1));
  const header = ws.getRow(headerRow);
  const lastUsed = findLastUsedColumn(ws);

  let callsCol = 0;
  let callsAvgCol = 0;
  let selfiesCol = 0;
  let selfiesAvgCol = 0;
  for (let c = 1; c <= Math.max(lastUsed + 2, header.cellCount || 0); c++) {
    const top = normalize(cellText(dateHeader.getCell(c)));
    const label = normalize(cellText(header.getCell(c)));
    if ((top === "total monthly" || label === "total calls") && label === "total calls") {
      callsCol = c;
    }
    if ((top === "total monthly" || label === "total calls avg") && label === "total calls avg") {
      callsAvgCol = c;
    }
    if ((top === "total monthly" || label === "total selfies") && label === "total selfies") {
      selfiesCol = c;
    }
    if (
      (top === "total monthly" || label === "total selfies avg") &&
      label === "total selfies avg"
    ) {
      selfiesAvgCol = c;
    }
  }

  if (callsCol && selfiesCol && !callsAvgCol && selfiesCol === callsCol + 1) {
    ws.spliceColumns(selfiesCol, 0, []);
    callsAvgCol = selfiesCol;
    selfiesCol += 1;
    if (selfiesAvgCol >= callsAvgCol) selfiesAvgCol += 1;
  }

  const existingCols = [callsCol, callsAvgCol, selfiesCol, selfiesAvgCol].filter(Boolean);
  let nextCol = existingCols.length ? Math.max(...existingCols) + 1 : lastUsed + 1;
  if (!callsCol) callsCol = nextCol++;
  if (!callsAvgCol) callsAvgCol = nextCol++;
  if (!selfiesCol) selfiesCol = nextCol++;
  if (!selfiesAvgCol) selfiesAvgCol = nextCol++;

  for (const col of [callsCol, callsAvgCol, selfiesCol, selfiesAvgCol]) {
    const topCell = dateHeader.getCell(col);
    topCell.value = "Total Monthly";
    styleHeader(topCell);
  }

  const labels: Array<[number, string]> = [
    [callsCol, "Total Calls"],
    [callsAvgCol, "Total Calls Avg"],
    [selfiesCol, "Total Selfies"],
    [selfiesAvgCol, "Total Selfies Avg"],
  ];

  for (const [col, label] of labels) {
    const cell = header.getCell(col);
    cell.value = label;
    styleHeader(cell);
    ws.getColumn(col).width = label.includes("Avg") ? 16 : 14;
  }

  dateHeader.commit?.();
  header.commit?.();
  return { callsCol, callsAvgCol, selfiesCol, selfiesAvgCol };
}

function sumFormula(cols: number[], rowNumber: number): string | null {
  if (!cols.length) return null;
  return `SUM(${cols.map((col) => `${columnLetter(col)}${rowNumber}`).join(",")})`;
}

function numericCellValue(cell: ExcelJS.Cell): number {
  const value = cell.value;
  if (typeof value === "number") return value;
  if (typeof value === "object" && value && "result" in value) {
    const result = Number(value.result ?? 0);
    return Number.isFinite(result) ? result : 0;
  }
  const parsed = Number(cellText(cell));
  return Number.isFinite(parsed) ? parsed : 0;
}

function rowSum(ws: ExcelJS.Worksheet, cols: number[], rowNumber: number): number {
  return cols.reduce((sum, col) => sum + numericCellValue(ws.getRow(rowNumber).getCell(col)), 0);
}

function averageFormula(cols: number[], rowNumber: number): string | null {
  if (!cols.length) return null;
  return `ROUND(${sumFormula(cols, rowNumber)}/${cols.length},0)`;
}

function rowAverage(ws: ExcelJS.Worksheet, cols: number[], rowNumber: number): number {
  if (!cols.length) return 0;
  return Math.round(rowSum(ws, cols, rowNumber) / cols.length);
}

function fillMonthlyTotals(
  ws: ExcelJS.Worksheet,
  dataStartRow: number,
  dataEndRow: number,
  monthlyTotals: MonthlyTotalCols,
  datePairs: Map<string, DatePairCols>,
) {
  const orderedPairs = [...datePairs.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([, pair]) => pair);
  const callCols = orderedPairs
    .map((pair) => pair.callsCol ?? pair.totalCol)
    .filter(
      (col) =>
        col !== monthlyTotals.callsCol &&
        col !== monthlyTotals.callsAvgCol &&
        col !== monthlyTotals.selfiesCol &&
        col !== monthlyTotals.selfiesAvgCol,
    );
  const selfieCols = orderedPairs
    .map((pair) => pair.selfiesCol)
    .filter(
      (col) =>
        col !== monthlyTotals.callsCol &&
        col !== monthlyTotals.callsAvgCol &&
        col !== monthlyTotals.selfiesCol &&
        col !== monthlyTotals.selfiesAvgCol,
    );

  for (let r = dataStartRow; r <= dataEndRow; r++) {
    const row = ws.getRow(r);
    const callsCell = row.getCell(monthlyTotals.callsCol);
    const callsFormula = sumFormula(callCols, r);
    const callsResult = rowSum(ws, callCols, r);
    callsCell.value = callsFormula ? { formula: callsFormula, result: callsResult } : null;
    styleBodyCell(callsCell);

    const callsAvgCell = row.getCell(monthlyTotals.callsAvgCol);
    const callsAvgFormula = averageFormula(callCols, r);
    const callsAvgResult = rowAverage(ws, callCols, r);
    callsAvgCell.value = callsAvgFormula
      ? { formula: callsAvgFormula, result: callsAvgResult }
      : null;
    styleBodyCell(callsAvgCell);

    const selfiesCell = row.getCell(monthlyTotals.selfiesCol);
    const selfiesFormula = sumFormula(selfieCols, r);
    const selfiesResult = rowSum(ws, selfieCols, r);
    selfiesCell.value = selfiesFormula ? { formula: selfiesFormula, result: selfiesResult } : null;
    styleBodyCell(selfiesCell);

    const selfiesAvgCell = row.getCell(monthlyTotals.selfiesAvgCol);
    const selfiesAvgFormula = averageFormula(selfieCols, r);
    const selfiesAvgResult = rowAverage(ws, selfieCols, r);
    selfiesAvgCell.value = selfiesAvgFormula
      ? { formula: selfiesAvgFormula, result: selfiesAvgResult }
      : null;
    styleBodyCell(selfiesAvgCell);
    row.commit?.();
  }
}

function ensureMasterHeaders(ws: ExcelJS.Worksheet, det: DetectedSheet) {
  const row = ws.getRow(det.headerRow);
  const headers: Array<[number | null, string]> = [
    [det.regionCol, "Region"],
    [det.codeCol, "Employee Code"],
    [det.nameCol, "Name"],
    [det.cityCol, "City"],
    [det.designationCol, "Designation"],
  ];

  for (const [col, label] of headers) {
    if (!col) continue;
    const cell = row.getCell(col);
    if (!cellText(cell).trim()) {
      cell.value = label;
      styleHeader(cell);
    } else {
      keepTemplateHeaderStyle(cell);
    }
  }
}

function createEmployeeMatcher(employees: Emp[]) {
  const empByCode = new Map<string, number>();
  const empByCodeSuffix = new Map<string, number>();
  const empByName = new Map<string, number>();
  const empByCleanName = new Map<string, number>();
  const duplicateCodes = new Set<string>();
  const duplicateCodeSuffixes = new Set<string>();
  const duplicateNames = new Set<string>();
  const duplicateCleanNames = new Set<string>();
  let codeFallbackMatches = 0;
  let codeSuffixMatches = 0;

  employees.forEach((e, i) => {
    if (e.code) {
      if (empByCode.has(e.code)) duplicateCodes.add(e.code);
      else empByCode.set(e.code, i);
      for (const suffix of [e.code, ...codeSuffixes(e.code)]) {
        if (empByCodeSuffix.has(suffix)) duplicateCodeSuffixes.add(suffix);
        else empByCodeSuffix.set(suffix, i);
      }
    }
    if (empByName.has(e.nameKey)) duplicateNames.add(e.nameKey);
    empByName.set(e.nameKey, i);
    const cleanName = e.cleanNameKey;
    if (empByCleanName.has(cleanName)) duplicateCleanNames.add(cleanName);
    empByCleanName.set(cleanName, i);
  });

  function findEmployeeByUniqueName(name: string): number {
    const nameKey = normalize(name);
    const byName = empByName.get(nameKey);
    if (byName !== undefined && !duplicateNames.has(nameKey)) return byName;

    const cleanName = personNameKey(name);
    const byCleanName = empByCleanName.get(cleanName);
    if (byCleanName !== undefined && !duplicateCleanNames.has(cleanName)) return byCleanName;

    const prefixMatches = employees
      .map((employee, idx) => ({ idx, cleanName: employee.cleanNameKey }))
      .filter(
        (employee) =>
          employee.cleanName.length >= 8 &&
          !duplicateCleanNames.has(employee.cleanName) &&
          (cleanName.startsWith(`${employee.cleanName} `) ||
            employee.cleanName.startsWith(`${cleanName} `)),
      );

    if (prefixMatches.length === 1) return prefixMatches[0].idx;

    const scored = employees
      .map((employee, idx) => ({ idx, score: nameMatchScore(cleanName, employee.cleanNameKey) }))
      .filter((candidate) => candidate.score >= 0.86)
      .sort((a, b) => b.score - a.score);

    if (!scored.length) return -1;
    const [best, second] = scored;
    return !second || best.score - second.score >= 0.05 ? best.idx : -1;
  }

  function findEmployee(row: { code: string | null; name: string }, countFallback = true): number {
    const code = normalizeEmployeeCode(row.code);
    if (code) {
      const byCode = empByCode.get(code);
      if (byCode !== undefined) return byCode;
      for (const suffix of [code, ...codeSuffixes(code)]) {
        const bySuffix = empByCodeSuffix.get(suffix);
        if (bySuffix !== undefined && !duplicateCodeSuffixes.has(suffix)) {
          if (countFallback) codeSuffixMatches++;
          return bySuffix;
        }
      }
      const byName = findEmployeeByUniqueName(row.name);
      if (byName >= 0 && countFallback) codeFallbackMatches++;
      return byName;
    }

    return findEmployeeByUniqueName(row.name);
  }

  return {
    findEmployee,
    metrics: () => ({
      duplicateCodes,
      codeFallbackMatches,
      codeSuffixMatches,
    }),
  };
}

interface ExcelSheetState {
  det: DetectedSheet;
  employees: Emp[];
  matcher: ReturnType<typeof createEmployeeMatcher>;
}

function normalizeTeamName(value: string): string {
  return normalize(value).replace(/[^a-z0-9]+/g, "");
}

function teamNamesMatch(sourceValue: string, sheetName: string): boolean {
  const sourceKey = normalizeTeamName(sourceValue);
  const sheetKey = normalizeTeamName(sheetName);
  if (!sourceKey || !sheetKey) return false;
  return sourceKey.includes(sheetKey) || sheetKey.includes(sourceKey);
}

function assignPdfsToReportSheets(
  states: ExcelSheetState[],
  pdfResults: PdfParseResult[],
): Map<ExcelSheetState, PdfParseResult[]> {
  const assigned = new Map(states.map((state) => [state, [] as PdfParseResult[]]));

  for (const pdf of pdfResults) {
    const directMatches = states.filter((state) => teamNamesMatch(pdf.fileName, state.det.ws.name));
    if (directMatches.length) {
      directMatches.forEach((state) => assigned.get(state)?.push(pdf));
      continue;
    }

    const scored = states
      .map((state) => ({
        state,
        score: pdf.rows.reduce(
          (count, row) => count + (state.matcher.findEmployee(row, false) >= 0 ? 1 : 0),
          0,
        ),
      }))
      .filter((item) => item.score > 0);

    scored.forEach((item) => assigned.get(item.state)?.push(pdf));
  }

  return assigned;
}

function fillMonthlyReportSheet(
  state: ExcelSheetState,
  pdfResults: PdfParseResult[],
  dates: string[],
) {
  const { det, employees, matcher } = state;
  const ws = det.ws;
  const warnings: string[] = [];
  const matchByEmp = new Map<number, Map<string, MatchValue>>();
  const unmatched = new Set<string>();
  const unmatchedRows: UnmatchedPdfRow[] = [];

  for (const pdf of pdfResults) {
    if (pdf.rows.length < employees.length) {
      warnings.push(
        `PDF parsing incomplete, please review. ${pdf.fileName}: ${pdf.rows.length}/${employees.length} rows extracted for ${ws.name}.`,
      );
    }

    for (const row of pdf.rows) {
      const idx = matcher.findEmployee(row);
      if (idx < 0) {
        unmatched.add(row.code ? `${row.code} - ${row.name}` : row.name);
        unmatchedRows.push({ date: pdf.date, fileName: pdf.fileName, row });
        continue;
      }

      const dateMap = matchByEmp.get(idx) ?? new Map<string, MatchValue>();
      const existing = dateMap.get(pdf.date);
      if (!existing || row.count > existing.selfies) {
        dateMap.set(pdf.date, {
          planned: row.planned,
          unplanned: row.unplanned,
          selfies: row.count,
          selfieText: row.selfieText,
          total: row.total,
        });
      } else if (existing.total === 0 && row.total > 0) {
        dateMap.set(pdf.date, {
          ...existing,
          planned: row.planned,
          unplanned: row.unplanned,
          total: row.total,
        });
      }
      matchByEmp.set(idx, dateMap);
    }
  }

  ensureMasterHeaders(ws, det);
  const datePairs = findDatePairColumns(ws, det.headerRow, dates);

  const empTotals = employees.map((employee, idx) => {
    const matches = matchByEmp.get(idx);
    let total = 0;
    for (const date of dates) {
      const match = matches?.get(date);
      if (match) total += match.total;
    }
    return { ...employee, idx, total };
  });

  for (const emp of empTotals) {
    const row = ws.getRow(emp.row);
    const matches = matchByEmp.get(emp.idx);

    for (const date of dates) {
      const match = matches?.get(date);
      const pair = datePairs.get(date);
      if (!pair) continue;
      const value = match ?? null;

      if (pair.plannedCol) {
        const plannedCell = row.getCell(pair.plannedCol);
        plannedCell.value = value ? value.planned : null;
        styleBodyCell(plannedCell);
      }

      if (pair.unplannedCol) {
        const unplannedCell = row.getCell(pair.unplannedCol);
        unplannedCell.value = value ? value.unplanned : null;
        styleBodyCell(unplannedCell);
      }

      if (pair.callsCol) {
        const callsCell = row.getCell(pair.callsCol);
        callsCell.value = value ? value.total : null;
        styleBodyCell(callsCell);
      }

      const selfieCell = row.getCell(pair.selfiesCol);
      selfieCell.value = value ? value.selfies : null;
      styleBodyCell(selfieCell);

      if (!pair.callsCol) {
        const totalCell = row.getCell(pair.totalCol);
        totalCell.value = value ? value.total : null;
        styleBodyCell(totalCell);
      }
    }

    row.commit?.();
  }

  const finalDataEndRow = appendUnmatchedRows(ws, det, unmatchedRows, datePairs);
  const monthlyTotals = findMonthlyTotalColumns(ws, det.headerRow);
  fillMonthlyTotals(ws, det.dataStartRow, finalDataEndRow, monthlyTotals, datePairs);
  applyBordersAndWidths(
    ws,
    det.headerRow,
    det.dataStartRow,
    Math.max(finalDataEndRow, det.dataEndRow, ws.rowCount),
  );

  const metrics = matcher.metrics();
  if (unmatched.size > 0) {
    warnings.push(
      `${unmatched.size} PDF row(s) still need manual review on ${ws.name}. They were added at the bottom of that sheet.`,
    );
  }
  if (metrics.codeFallbackMatches > 0) {
    warnings.push(
      `${metrics.codeFallbackMatches} PDF row(s) on ${ws.name} had an unknown employee code but were safely matched by name.`,
    );
  }
  if (metrics.codeSuffixMatches > 0) {
    warnings.push(
      `${metrics.codeSuffixMatches} PDF row(s) on ${ws.name} had extra digits in the code and were safely matched by a unique code suffix.`,
    );
  }
  if (metrics.duplicateCodes.size > 0) {
    warnings.push(
      `Employee mapping mismatch detected on ${ws.name}: duplicate employee code(s): ${[
        ...metrics.duplicateCodes,
      ]
        .slice(0, 5)
        .join(", ")}.`,
    );
  }

  const matched = [...matchByEmp.values()].filter((m) =>
    [...m.values()].some((v) => v.selfies > 0 || v.total > 0 || v.planned > 0 || v.unplanned > 0),
  ).length;

  return {
    totalEmployees: employees.length,
    matchedEmployees: matched,
    unmatchedNames: [...unmatched],
    warnings,
    totalMatchedRows: matchByEmp.size,
    skippedRows: unmatched.size,
    preview: empTotals
      .slice(0, 10)
      .map((s) => ({ name: `${ws.name} - ${s.name}`, total: s.total })),
  };
}

async function processExcelMultiSheet(
  excelFile: File,
  options: ProcessOptions,
  wb: ExcelJS.Workbook,
  detections: DetectedSheet[],
): Promise<ProcessReport> {
  const states = detections
    .map((det) => {
      const employees = readEmployees(det.ws, det);
      return { det, employees, matcher: createEmployeeMatcher(employees) };
    })
    .filter((state) => state.employees.length > 0);

  if (!states.length) {
    throw new Error(
      "No employees found in the Excel sheets. Make sure the file has a 'Name' column.",
    );
  }

  const dates = [...new Set(options.pdfResults.map((r) => r.date))].sort();
  const days = dates.map((date) => Number(date.slice(8, 10)));
  const assigned = assignPdfsToReportSheets(states, options.pdfResults);
  const warnings: string[] = [];
  const unmatchedNames = new Set<string>();
  const preview: { name: string; total: number }[] = [];
  let totalEmployees = 0;
  let matchedEmployees = 0;
  let totalMatchedRows = 0;
  let totalSkipped = 0;

  for (const state of states) {
    const pdfs = assigned.get(state) ?? [];
    totalEmployees += state.employees.length;
    if (!pdfs.length) {
      warnings.push(
        `No matching PDF data was found for ${state.det.ws.name}; that sheet was kept unchanged.`,
      );
      continue;
    }

    const result = fillMonthlyReportSheet(state, pdfs, dates);
    matchedEmployees += result.matchedEmployees;
    totalMatchedRows += result.totalMatchedRows;
    totalSkipped += result.skippedRows;
    result.unmatchedNames.forEach((name) => unmatchedNames.add(name));
    warnings.push(...result.warnings);
    preview.push(...result.preview);
  }

  addPdfAuditSheet(wb, options.pdfResults, (row) => {
    for (const state of states) {
      const idx = state.matcher.findEmployee(row, false);
      if (idx >= 0) return state.employees[idx];
    }
    return null;
  });

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  return {
    totalEmployees,
    matchedEmployees,
    unmatchedNames: [...unmatchedNames],
    warnings,
    debug: {
      totalPdfsUploaded: options.pdfResults.length,
      totalEmployeesDetected: options.pdfResults.reduce((sum, pdf) => sum + pdf.rows.length, 0),
      totalMatched: totalMatchedRows,
      totalSkipped,
      totalRecordsInsertedUpdated: 0,
      parsingErrors: 0,
    },
    dates,
    days,
    blob,
    fileName: excelFile.name,
    sheetName: states[0].det.ws.name,
    preview: preview.sort((a, b) => b.total - a.total).slice(0, 10),
  };
}

export async function processExcel(
  excelFile: File,
  options: ProcessOptions,
): Promise<ProcessReport> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await excelFile.arrayBuffer());

  const reportSheets = detectReportSheets(wb);
  if (reportSheets.length > 1) {
    return processExcelMultiSheet(excelFile, options, wb, reportSheets);
  }

  const det = detectActiveSheet(wb);
  const ws = det.ws;
  const employees = readEmployees(ws, det);

  if (employees.length === 0) {
    throw new Error(
      "No employees found in the Excel sheet. Make sure the file has a 'Name' column.",
    );
  }

  const dates = [...new Set(options.pdfResults.map((r) => r.date))].sort();
  const days = dates.map((date) => Number(date.slice(8, 10)));
  const warnings: string[] = [];

  const matchByEmp = new Map<number, Map<string, MatchValue>>();
  const unmatched = new Set<string>();
  const unmatchedRows: UnmatchedPdfRow[] = [];
  const empByCode = new Map<string, number>();
  const empByCodeSuffix = new Map<string, number>();
  const empByName = new Map<string, number>();
  const empByCleanName = new Map<string, number>();
  const duplicateCodes = new Set<string>();
  const duplicateCodeSuffixes = new Set<string>();
  const duplicateNames = new Set<string>();
  const duplicateCleanNames = new Set<string>();
  let codeFallbackMatches = 0;
  let codeSuffixMatches = 0;

  employees.forEach((e, i) => {
    if (e.code) {
      if (empByCode.has(e.code)) duplicateCodes.add(e.code);
      else empByCode.set(e.code, i);
      for (const suffix of [e.code, ...codeSuffixes(e.code)]) {
        if (empByCodeSuffix.has(suffix)) duplicateCodeSuffixes.add(suffix);
        else empByCodeSuffix.set(suffix, i);
      }
    }
    if (empByName.has(e.nameKey)) duplicateNames.add(e.nameKey);
    empByName.set(e.nameKey, i);
    const cleanName = e.cleanNameKey;
    if (empByCleanName.has(cleanName)) duplicateCleanNames.add(cleanName);
    empByCleanName.set(cleanName, i);
  });

  function findEmployeeByUniqueName(name: string): number {
    const nameKey = normalize(name);
    const byName = empByName.get(nameKey);
    if (byName !== undefined && !duplicateNames.has(nameKey)) return byName;

    const cleanName = personNameKey(name);
    const byCleanName = empByCleanName.get(cleanName);
    if (byCleanName !== undefined && !duplicateCleanNames.has(cleanName)) return byCleanName;

    const prefixMatches = employees
      .map((employee, idx) => ({ idx, cleanName: employee.cleanNameKey }))
      .filter(
        (employee) =>
          employee.cleanName.length >= 8 &&
          !duplicateCleanNames.has(employee.cleanName) &&
          (cleanName.startsWith(`${employee.cleanName} `) ||
            employee.cleanName.startsWith(`${cleanName} `)),
      );

    if (prefixMatches.length === 1) return prefixMatches[0].idx;

    const scored = employees
      .map((employee, idx) => {
        return { idx, score: nameMatchScore(cleanName, employee.cleanNameKey) };
      })
      .filter((candidate) => candidate.score >= 0.86)
      .sort((a, b) => b.score - a.score);

    if (!scored.length) return -1;
    const [best, second] = scored;
    return !second || best.score - second.score >= 0.05 ? best.idx : -1;
  }

  function findEmployee(row: { code: string | null; name: string }, countFallback = true): number {
    const code = normalizeEmployeeCode(row.code);
    if (code) {
      const byCode = empByCode.get(code);
      if (byCode !== undefined) return byCode;
      for (const suffix of [code, ...codeSuffixes(code)]) {
        const bySuffix = empByCodeSuffix.get(suffix);
        if (bySuffix !== undefined && !duplicateCodeSuffixes.has(suffix)) {
          if (countFallback) codeSuffixMatches++;
          return bySuffix;
        }
      }
      const byName = findEmployeeByUniqueName(row.name);
      if (byName >= 0 && countFallback) codeFallbackMatches++;
      return byName;
    }

    return findEmployeeByUniqueName(row.name);
  }

  for (const pdf of options.pdfResults) {
    if (pdf.rows.length < employees.length) {
      warnings.push(
        `PDF parsing incomplete, please review. ${pdf.fileName}: ${pdf.rows.length}/${employees.length} rows extracted.`,
      );
    }

    for (const row of pdf.rows) {
      const idx = findEmployee(row);
      if (idx < 0) {
        unmatched.add(row.code ? `${row.code} - ${row.name}` : row.name);
        unmatchedRows.push({ date: pdf.date, fileName: pdf.fileName, row });
        continue;
      }

      const dateMap = matchByEmp.get(idx) ?? new Map<string, MatchValue>();
      const existing = dateMap.get(pdf.date);
      if (!existing || row.count > existing.selfies) {
        dateMap.set(pdf.date, {
          planned: row.planned,
          unplanned: row.unplanned,
          selfies: row.count,
          selfieText: row.selfieText,
          total: row.total,
        });
      } else if (existing.total === 0 && row.total > 0) {
        dateMap.set(pdf.date, {
          ...existing,
          planned: row.planned,
          unplanned: row.unplanned,
          total: row.total,
        });
      }
      matchByEmp.set(idx, dateMap);
    }
  }

  ensureMasterHeaders(ws, det);
  const datePairs = findDatePairColumns(ws, det.headerRow, dates);

  const empTotals = employees.map((employee, idx) => {
    const matches = matchByEmp.get(idx);
    let total = 0;
    for (const date of dates) {
      const match = matches?.get(date);
      if (match) total += match.total;
    }
    return { ...employee, idx, total };
  });

  for (const emp of empTotals) {
    const row = ws.getRow(emp.row);
    const matches = matchByEmp.get(emp.idx);

    for (const date of dates) {
      const match = matches?.get(date);
      const pair = datePairs.get(date);
      if (!pair) continue;
      const value = match ?? null;

      if (pair.plannedCol) {
        const plannedCell = row.getCell(pair.plannedCol);
        plannedCell.value = value ? value.planned : null;
        styleBodyCell(plannedCell);
      }

      if (pair.unplannedCol) {
        const unplannedCell = row.getCell(pair.unplannedCol);
        unplannedCell.value = value ? value.unplanned : null;
        styleBodyCell(unplannedCell);
      }

      if (pair.callsCol) {
        const callsCell = row.getCell(pair.callsCol);
        callsCell.value = value ? value.total : null;
        styleBodyCell(callsCell);
      }

      const selfieCell = row.getCell(pair.selfiesCol);
      selfieCell.value = value ? value.selfies : null;
      styleBodyCell(selfieCell);

      if (!pair.callsCol) {
        const totalCell = row.getCell(pair.totalCol);
        totalCell.value = value ? value.total : null;
        styleBodyCell(totalCell);
      }
    }

    row.commit?.();
  }

  const finalDataEndRow = appendUnmatchedRows(ws, det, unmatchedRows, datePairs);
  const monthlyTotals = findMonthlyTotalColumns(ws, det.headerRow);
  fillMonthlyTotals(ws, det.dataStartRow, finalDataEndRow, monthlyTotals, datePairs);

  applyBordersAndWidths(
    ws,
    det.headerRow,
    det.dataStartRow,
    Math.max(finalDataEndRow, det.dataEndRow, ws.rowCount),
  );
  addPdfAuditSheet(wb, options.pdfResults, (row) => {
    const idx = findEmployee(row, false);
    return idx >= 0 ? employees[idx] : null;
  });

  if (unmatched.size > 0) {
    warnings.push(
      `${unmatched.size} PDF row(s) still need manual review. They were added at the bottom of the Excel report and listed in the PDF Extracted Data sheet.`,
    );
  }
  if (codeFallbackMatches > 0) {
    warnings.push(
      `${codeFallbackMatches} PDF row(s) had an unknown employee code but were safely matched by name. Please review the PDF Extracted Data sheet.`,
    );
  }
  if (codeSuffixMatches > 0) {
    warnings.push(
      `${codeSuffixMatches} PDF row(s) had extra digits in the code and were safely matched by a unique code suffix.`,
    );
  }
  if (duplicateCodes.size > 0) {
    warnings.push(
      `Employee mapping mismatch detected: duplicate employee code(s) in Excel: ${[
        ...duplicateCodes,
      ]
        .slice(0, 5)
        .join(", ")}.`,
    );
  }

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const matched = [...matchByEmp.values()].filter((m) =>
    [...m.values()].some((v) => v.selfies > 0 || v.total > 0 || v.planned > 0 || v.unplanned > 0),
  ).length;

  console.info("[Excel matcher]", {
    totalEmployees: employees.length,
    extractedRows: options.pdfResults.reduce((sum, pdf) => sum + pdf.rows.length, 0),
    matchedEmployees: matchByEmp.size,
    skippedRows: unmatched.size,
    unmatchedRows: [...unmatched],
    warnings,
  });

  return {
    totalEmployees: employees.length,
    matchedEmployees: matched,
    unmatchedNames: [...unmatched],
    warnings,
    debug: {
      totalPdfsUploaded: options.pdfResults.length,
      totalEmployeesDetected: options.pdfResults.reduce((sum, pdf) => sum + pdf.rows.length, 0),
      totalMatched: matchByEmp.size,
      totalSkipped: unmatched.size,
      totalRecordsInsertedUpdated: 0,
      parsingErrors: 0,
    },
    dates,
    days,
    blob,
    fileName: excelFile.name,
    sheetName: ws.name,
    preview: empTotals.slice(0, 10).map((s) => ({ name: s.name, total: s.total })),
  };
}

function addPdfAuditSheet(
  wb: ExcelJS.Workbook,
  pdfResults: PdfParseResult[],
  resolveEmployee: (row: PdfParseResult["rows"][number]) => Emp | null,
) {
  const sheetName = "PDF Extracted Data";
  const existing = wb.getWorksheet(sheetName);
  if (existing) wb.removeWorksheet(existing.id);

  const ws = wb.addWorksheet(sheetName);
  ws.columns = [
    { header: "Source PDF", key: "sourcePdf", width: 28 },
    { header: "Date", key: "date", width: 14 },
    { header: "PDF Code", key: "pdfCode", width: 14 },
    { header: "PDF Name", key: "pdfName", width: 28 },
    { header: "Matched Excel Code", key: "matchedCode", width: 18 },
    { header: "Matched Excel Name", key: "matchedName", width: 28 },
    { header: "Planned", key: "planned", width: 10 },
    { header: "Unplanned", key: "unplanned", width: 12 },
    { header: "Calls", key: "calls", width: 10 },
    { header: "Selfie Text", key: "selfieText", width: 32 },
    { header: "Selfie Count", key: "selfieCount", width: 14 },
    { header: "Validation", key: "validation", width: 18 },
  ];

  const header = ws.getRow(1);
  header.eachCell((cell) => styleHeader(cell));

  for (const pdf of pdfResults) {
    for (const row of pdf.rows) {
      const employee = resolveEmployee(row);
      const validation = employee ? "Matched" : "Needs review";
      const added = ws.addRow({
        sourcePdf: pdf.fileName,
        date: pdf.date,
        pdfCode: row.code,
        pdfName: row.name,
        matchedCode: employee?.code,
        matchedName: employee?.name,
        planned: row.planned,
        unplanned: row.unplanned,
        calls: row.total,
        selfieText: row.selfieText,
        selfieCount: row.count,
        validation,
      });
      added.eachCell((cell) => {
        styleBodyCell(cell);
      });
    }
  }

  ws.views = [{ state: "frozen", ySplit: 1 }];
}

function appendUnmatchedRows(
  ws: ExcelJS.Worksheet,
  det: DetectedSheet,
  unmatchedRows: UnmatchedPdfRow[],
  datePairs: Map<string, DatePairCols>,
): number {
  if (!unmatchedRows.length) return Math.max(det.dataEndRow, ws.rowCount);

  const grouped = new Map<string, UnmatchedPdfRow>();
  for (const item of unmatchedRows) {
    const key = `${item.date}|${item.row.code ?? ""}|${personNameKey(item.row.name)}|${item.fileName}`;
    const existing = grouped.get(key);
    if (!existing || item.row.count > existing.row.count || item.row.total > existing.row.total) {
      grouped.set(key, item);
    }
  }

  let nextRow = Math.max(det.dataEndRow, ws.rowCount) + 1;
  for (const item of grouped.values()) {
    const row = ws.getRow(nextRow);
    if (det.regionCol) row.getCell(det.regionCol).value = "PDF Needs Review";
    if (det.codeCol) row.getCell(det.codeCol).value = item.row.code;
    row.getCell(det.nameCol).value = item.row.name || item.row.code || "Unmatched PDF row";
    if (det.cityCol) row.getCell(det.cityCol).value = item.fileName;
    if (det.designationCol) row.getCell(det.designationCol).value = "Unmatched PDF row";

    const pair = datePairs.get(item.date);
    if (pair) {
      if (pair.plannedCol) {
        const plannedCell = row.getCell(pair.plannedCol);
        plannedCell.value = item.row.planned || null;
        styleBodyCell(plannedCell);
      }
      if (pair.unplannedCol) {
        const unplannedCell = row.getCell(pair.unplannedCol);
        unplannedCell.value = item.row.unplanned || null;
        styleBodyCell(unplannedCell);
      }
      if (pair.callsCol) {
        const callsCell = row.getCell(pair.callsCol);
        callsCell.value = item.row.total || null;
        styleBodyCell(callsCell);
      }
      const selfieCell = row.getCell(pair.selfiesCol);
      selfieCell.value = item.row.count || null;
      styleBodyCell(selfieCell);

      if (!pair.callsCol) {
        const totalCell = row.getCell(pair.totalCol);
        totalCell.value = item.row.total || null;
        styleBodyCell(totalCell);
      }
    }

    row.eachCell((cell) => {
      cell.border = thinBorder();
      cell.alignment = { vertical: "middle", wrapText: true };
      if (!cell.fill) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
      }
    });
    row.commit?.();
    nextRow++;
  }

  return nextRow - 1;
}

function thinBorder(): Partial<ExcelJS.Borders> {
  const c = { style: "thin" as const, color: { argb: "FF000000" } };
  return { top: c, left: c, right: c, bottom: c };
}

function styleHeader(cell: ExcelJS.Cell) {
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  cell.font = { ...(cell.font ?? {}), bold: true, color: { argb: "FF000000" } };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
  cell.border = thinBorder();
}

function keepTemplateHeaderStyle(cell: ExcelJS.Cell) {
  cell.alignment = { ...(cell.alignment ?? {}), vertical: "middle", wrapText: true };
  cell.border = thinBorder();
}

function styleBodyCell(cell: ExcelJS.Cell) {
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  cell.border = thinBorder();
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
}

function applyBordersAndWidths(
  ws: ExcelJS.Worksheet,
  headerRow: number,
  dataStartRow: number,
  dataEndRow: number,
) {
  const lastCol = findLastUsedColumn(ws);
  for (let r = headerRow; r <= dataEndRow; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= lastCol; c++) {
      const cell = row.getCell(c);
      cell.border = thinBorder();
      if (r === headerRow) {
        keepTemplateHeaderStyle(cell);
      } else if (SPECIAL_TEXT.includes(normalize(cellText(cell)))) {
        cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
      }
    }
    row.commit?.();
  }
}
