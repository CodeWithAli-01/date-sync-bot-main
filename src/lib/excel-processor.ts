// Read the Active Members workbook and update only the date/total cells needed.
// Employee rows and existing columns are preserved exactly as supplied.
import ExcelJS from "exceljs";
import type { PdfParseResult } from "./pdf-extractor";

export interface ProcessOptions {
  pdfResults: PdfParseResult[];
}

export interface ProcessReport {
  totalEmployees: number;
  matchedEmployees: number;
  unmatchedNames: string[];
  dates: string[];
  days: number[];
  blob: Blob;
  fileName: string;
  preview: { name: string; total: number }[];
}

const NAME_HINTS = ["employee name", "employee", "name", "emp name", "staff name", "member name"];
const CODE_HINTS = ["employee code", "emp code", "code", "emp id", "employee id", "id"];

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

function detectActiveSheet(wb: ExcelJS.Workbook): DetectedSheet {
  for (const ws of wb.worksheets) {
    const maxScan = Math.min(10, ws.rowCount || 10);
    for (let r = 1; r <= maxScan; r++) {
      const row = ws.getRow(r);
      if (!row.cellCount) continue;
      let nameCol = 0;
      let codeCol = 0;
      const maxC = Math.max(row.cellCount || 0, 30);
      for (let c = 1; c <= maxC; c++) {
        const v = normalize(cellText(row.getCell(c)));
        if (!v) continue;
        if (!nameCol && (NAME_HINTS.includes(v) || (v.includes("name") && !v.includes("file"))))
          nameCol = c;
        if (!codeCol && (CODE_HINTS.includes(v) || /\bcode\b/.test(v) || /\bemp.*id\b/.test(v)))
          codeCol = c;
      }
      if (nameCol) {
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
            dataStartRow: r + 1,
            dataEndRow: ws.rowCount || r + 1,
          };
        }
      }
    }
  }
  const ws = wb.worksheets[0];
  return {
    ws,
    headerRow: 1,
    nameCol: 1,
    codeCol: null,
    dataStartRow: 2,
    dataEndRow: ws.rowCount || 2,
  };
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

function headerDay(cell: ExcelJS.Cell): number | null {
  const value = cell.value;
  if (typeof value === "number" && value >= 1 && value <= 31) return value;
  if (value instanceof Date) return value.getDate();
  const text = cellText(cell).trim();
  const n = Number(text);
  if (Number.isInteger(n) && n >= 1 && n <= 31) return n;
  const iso = text.match(/^\d{4}-\d{2}-(\d{2})$/);
  return iso ? Number(iso[1]) : null;
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

function findOrCreateDateColumns(
  ws: ExcelJS.Worksheet,
  headerRow: number,
  days: number[],
  dayMap: Map<number, string>,
): Map<number, number> {
  const header = ws.getRow(headerRow);
  const dayCols = new Map<number, number>();
  const lastUsed = findLastUsedColumn(ws);

  for (let c = 1; c <= Math.max(lastUsed, header.cellCount || 0); c++) {
    const day = headerDay(header.getCell(c));
    if (day != null && !dayCols.has(day)) dayCols.set(day, c);
  }

  let nextCol = lastUsed + 1;
  for (const day of days) {
    if (dayCols.has(day)) continue;
    const cell = header.getCell(nextCol);
    cell.value = day;
    cell.note = dayMap.get(day) ?? "";
    styleDateHeader(cell);
    dayCols.set(day, nextCol);
    ws.getColumn(nextCol).width = 6;
    nextCol++;
  }

  header.height = Math.max(header.height || 0, 24);
  header.commit?.();
  return dayCols;
}

function findOrCreateTotalColumn(
  ws: ExcelJS.Worksheet,
  headerRow: number,
  afterCol: number,
): number {
  const header = ws.getRow(headerRow);
  const lastUsed = findLastUsedColumn(ws);
  for (let c = 1; c <= Math.max(lastUsed, header.cellCount || 0); c++) {
    if (normalize(cellText(header.getCell(c))) === "total") return c;
  }

  const totalCol = Math.max(afterCol + 1, lastUsed + 1);
  const cell = header.getCell(totalCol);
  cell.value = "Total";
  styleTotalHeader(cell);
  ws.getColumn(totalCol).width = 10;
  header.commit?.();
  return totalCol;
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

  const dayMap = new Map<number, string>();
  for (const r of options.pdfResults) dayMap.set(r.day, r.date);
  const days = [...dayMap.keys()].sort((a, b) => a - b);
  const dates = days.map((d) => dayMap.get(d)!);

  const matchByEmp = new Map<number, Map<number, number>>();
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
    for (const row of pdf.rows) {
      const idx = findEmployee(row);
      if (idx < 0) {
        unmatched.add(row.name);
        continue;
      }
      const dayCounts = matchByEmp.get(idx) ?? new Map<number, number>();
      const existing = dayCounts.get(pdf.day);
      dayCounts.set(pdf.day, existing == null ? row.count : Math.max(existing, row.count));
      matchByEmp.set(idx, dayCounts);
    }
  }

  const dayCols = findOrCreateDateColumns(ws, det.headerRow, days, dayMap);
  const lastDayCol = Math.max(...[...dayCols.values()], det.nameCol);
  const totalCol = findOrCreateTotalColumn(ws, det.headerRow, lastDayCol);

  const empTotals = employees.map((employee, idx) => {
    const matches = matchByEmp.get(idx);
    let total = 0;
    for (const day of days) {
      const matchedValue = matches?.get(day);
      if (matchedValue != null) {
        total += matchedValue;
        continue;
      }

      const col = dayCols.get(day);
      const existingValue = col ? Number(ws.getRow(employee.row).getCell(col).value ?? 0) : 0;
      if (Number.isFinite(existingValue)) total += existingValue;
    }
    return { ...employee, idx, total };
  });

  const totals = empTotals.map((e) => e.total);
  const maxT = Math.max(...totals, 0);
  const avgT = totals.length ? totals.reduce((a, b) => a + b, 0) / totals.length : 0;
  const highThr = Math.max(avgT, maxT * 0.7);
  const lowThr = avgT * 0.4;

  for (const emp of empTotals) {
    const row = ws.getRow(emp.row);
    const matches = matchByEmp.get(emp.idx);

    for (const day of days) {
      const value = matches?.get(day);
      if (value == null) continue;

      const col = dayCols.get(day);
      if (!col) continue;
      const cell = row.getCell(col);
      cell.value = value;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = thinBorder();
      if (value === 0) cell.font = { color: { argb: "FFB0B0B0" } };
    }

    const totalRefs = days
      .map((day) => dayCols.get(day))
      .filter((col): col is number => col != null)
      .map((col) => `${colLetter(col)}${emp.row}`);
    const totalCell = row.getCell(totalCol);
    totalCell.value = totalRefs.length
      ? { formula: `SUM(${totalRefs.join(",")})`, result: emp.total }
      : emp.total;
    totalCell.font = { bold: true };
    totalCell.alignment = { horizontal: "center", vertical: "middle" };
    totalCell.border = thinBorder();

    let bg = "FFFFF4CC";
    if (emp.total >= highThr && emp.total > 0) bg = "FFD4F5DD";
    else if (emp.total <= lowThr) bg = "FFFCD9D9";
    totalCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
    row.commit?.();
  }

  ws.getColumn(totalCol).width = Math.max(ws.getColumn(totalCol).width || 0, 10);

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const matched = [...matchByEmp.values()].filter((m) => [...m.values()].some((v) => v > 0)).length;
  return {
    totalEmployees: employees.length,
    matchedEmployees: matched,
    unmatchedNames: [...unmatched],
    dates,
    days,
    blob,
    fileName: excelFile.name,
    preview: empTotals.slice(0, 10).map((s) => ({ name: s.name, total: s.total })),
  };
}

function thinBorder(): Partial<ExcelJS.Borders> {
  const c = { style: "thin" as const, color: { argb: "FFE2E6EE" } };
  return { top: c, left: c, right: c, bottom: c };
}

function styleDateHeader(cell: ExcelJS.Cell) {
  cell.alignment = { horizontal: "center", vertical: "middle" };
  cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3B5BDB" } };
  cell.border = thinBorder();
}

function styleTotalHeader(cell: ExcelJS.Cell) {
  cell.alignment = { horizontal: "center", vertical: "middle" };
  cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F2A4D" } };
  cell.border = thinBorder();
}

function colLetter(col: number): string {
  let s = "";
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}
