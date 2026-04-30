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
  dates: string[];
  days: number[];
  blob: Blob;
  fileName: string;
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

function similarity(a: string, b: string): number {
  if (a === b) return 1;
  if (!a.length || !b.length) return 0;
  const m = a.length;
  const n = b.length;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const c = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + c);
    }
  }
  return 1 - dp[m][n] / Math.max(m, n);
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
  fuzzy: string;
}

interface DatePairCols {
  selfiesCol: number;
  callsCol: number;
}

interface MatchValue {
  selfies: number;
  calls: number;
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
  }
  return String(value);
}

function detectActiveSheet(wb: ExcelJS.Workbook): DetectedSheet {
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

function readEmployees(ws: ExcelJS.Worksheet, det: DetectedSheet): Emp[] {
  const employees: Emp[] = [];
  let lastDataRow = det.dataStartRow - 1;

  for (let r = det.dataStartRow; r <= Math.max(det.dataEndRow, det.dataStartRow + 1000); r++) {
    const row = ws.getRow(r);
    const nameVal = cellText(row.getCell(det.nameCol)).trim();
    if (!nameVal) {
      if (r - lastDataRow > 5) break;
      continue;
    }
    const codeVal = det.codeCol
      ? cellText(row.getCell(det.codeCol)).trim().toUpperCase() || null
      : null;
    employees.push({
      row: r,
      name: nameVal,
      code: codeVal,
      nameKey: normalize(nameVal),
      fuzzy: fuzzyKey(nameVal),
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

function findDatePairColumns(
  ws: ExcelJS.Worksheet,
  headerRow: number,
  dates: string[],
): Map<string, DatePairCols> {
  const header = ws.getRow(headerRow);
  const pairs = new Map<string, DatePairCols>();
  const lastUsed = findLastUsedColumn(ws);

  for (const date of dates) {
    const selfieHeader = `${dateLabel(date)} Selfies`.toLowerCase();
    for (let c = 1; c <= Math.max(lastUsed, header.cellCount || 0); c++) {
      if (normalize(cellText(header.getCell(c))) !== selfieHeader) continue;
      const callsCol = normalize(cellText(header.getCell(c + 1))) === "calls" ? c + 1 : c + 1;
      pairs.set(date, { selfiesCol: c, callsCol });
      break;
    }
  }

  let nextCol = lastUsed + 1;
  for (const date of dates) {
    if (pairs.has(date)) continue;
    const selfiesCol = nextCol;
    const callsCol = nextCol + 1;
    const selfieHeader = header.getCell(selfiesCol);
    const callsHeader = header.getCell(callsCol);

    selfieHeader.value = `${dateLabel(date)} Selfies`;
    callsHeader.value = "Calls";
    styleHeader(selfieHeader);
    styleHeader(callsHeader);
    selfieHeader.note = date;
    callsHeader.note = date;
    ws.getColumn(selfiesCol).width = 30;
    ws.getColumn(callsCol).width = 10;
    pairs.set(date, { selfiesCol, callsCol });
    nextCol += 2;
  }

  header.height = Math.max(header.height || 0, 24);
  header.commit?.();
  return pairs;
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
    if (!cellText(cell).trim()) cell.value = label;
    styleHeader(cell);
  }
}

export async function processExcel(
  excelFile: File,
  options: ProcessOptions,
): Promise<ProcessReport> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await excelFile.arrayBuffer());

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
  const empByCode = new Map<string, number>();
  const empByName = new Map<string, number>();

  employees.forEach((e, i) => {
    if (e.code) empByCode.set(e.code, i);
    empByName.set(e.nameKey, i);
  });

  function findEmployee(row: { code: string | null; name: string }): number {
    if (row.code) {
      const byCode = empByCode.get(row.code.toUpperCase());
      if (byCode !== undefined) return byCode;
    }

    const byName = empByName.get(normalize(row.name));
    if (byName !== undefined) return byName;

    const target = fuzzyKey(row.name);
    let best = -1;
    let bestScore = 0;
    for (let i = 0; i < employees.length; i++) {
      const score = similarity(target, employees[i].fuzzy);
      if (score > bestScore) {
        bestScore = score;
        best = i;
      }
    }
    return bestScore >= 0.85 ? best : -1;
  }

  for (const pdf of options.pdfResults) {
    if (pdf.rows.length < employees.length) {
      warnings.push(
        `Parsing incomplete - some employees missing in ${pdf.fileName} (${pdf.rows.length}/${employees.length} rows extracted).`,
      );
    }

    for (const row of pdf.rows) {
      const idx = findEmployee(row);
      if (idx < 0) {
        unmatched.add(row.name);
        continue;
      }

      const dateMap = matchByEmp.get(idx) ?? new Map<string, MatchValue>();
      const existing = dateMap.get(pdf.date);
      if (!existing || row.count > existing.selfies) {
        dateMap.set(pdf.date, { selfies: row.count, calls: row.calls ?? existing?.calls ?? 0 });
      } else if (existing.calls === 0 && row.calls > 0) {
        dateMap.set(pdf.date, { ...existing, calls: row.calls });
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
      if (match) total += match.selfies;
    }
    return { ...employee, idx, total };
  });

  for (const emp of empTotals) {
    const row = ws.getRow(emp.row);
    const matches = matchByEmp.get(emp.idx);

    for (const date of dates) {
      const match = matches?.get(date);
      if (!match) continue;
      const pair = datePairs.get(date);
      if (!pair) continue;

      const selfieCell = row.getCell(pair.selfiesCol);
      selfieCell.value = `${match.selfies} selfies with locations in grp`;
      styleBodyCell(selfieCell, match.selfies === 0 ? "zero" : "normal");

      const callsCell = row.getCell(pair.callsCol);
      callsCell.value = match.calls;
      styleBodyCell(callsCell, match.calls === 0 ? "zero" : "normal");
    }

    row.commit?.();
  }

  applyBordersAndWidths(ws, det.headerRow, det.dataStartRow, Math.max(det.dataEndRow, ws.rowCount));

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const matched = [...matchByEmp.values()].filter((m) =>
    [...m.values()].some((v) => v.selfies > 0),
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
    dates,
    days,
    blob,
    fileName: excelFile.name,
    preview: empTotals.slice(0, 10).map((s) => ({ name: s.name, total: s.total })),
  };
}

function thinBorder(): Partial<ExcelJS.Borders> {
  const c = { style: "thin" as const, color: { argb: "FF000000" } };
  return { top: c, left: c, right: c, bottom: c };
}

function styleHeader(cell: ExcelJS.Cell) {
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  cell.font = { bold: true, color: { argb: "FF000000" } };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
  cell.border = thinBorder();
}

function styleBodyCell(cell: ExcelJS.Cell, tone: "normal" | "zero") {
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  cell.border = thinBorder();
  cell.fill =
    tone === "zero"
      ? { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } }
      : { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
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
        styleHeader(cell);
      } else if (SPECIAL_TEXT.includes(normalize(cellText(cell)))) {
        cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
      }
    }
    row.height = r === headerRow ? Math.max(row.height || 0, 24) : row.height;
    row.commit?.();
  }

  for (let c = 1; c <= lastCol; c++) {
    const column = ws.getColumn(c);
    let maxLength = 8;
    for (let r = headerRow; r <= Math.max(dataEndRow, dataStartRow); r++) {
      maxLength = Math.max(maxLength, cellText(ws.getRow(r).getCell(c)).length);
    }
    column.width = Math.min(Math.max(maxLength + 2, column.width || 8), 35);
  }
}
