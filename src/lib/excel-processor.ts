// Read the Active Members workbook, fill day-of-month columns (1..N) with selfie
// counts (hybrid match: code -> exact name -> fuzzy name), add SUM Total column,
// format, sort, write back into the SAME workbook.
import ExcelJS from "exceljs";
import type { PdfParseResult } from "./pdf-extractor";

export interface ProcessOptions {
  pdfResults: PdfParseResult[];
}

export interface ProcessReport {
  totalEmployees: number;
  matchedEmployees: number;
  unmatchedNames: string[];
  dates: string[]; // YYYY-MM-DD covered
  days: number[]; // day numbers covered (1..31)
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
  // strip punctuation, collapse, lowercase — for fuzzy matching
  return normalize(s).replace(/[^\p{L}\p{N}\s]/gu, "").replace(/\s+/g, " ");
}

// Levenshtein similarity 0..1
function similarity(a: string, b: string): number {
  if (a === b) return 1;
  if (!a.length || !b.length) return 0;
  const m = a.length, n = b.length;
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

function detectActiveSheet(wb: ExcelJS.Workbook): DetectedSheet {
  for (const ws of wb.worksheets) {
    const maxScan = Math.min(10, ws.rowCount || 10);
    for (let r = 1; r <= maxScan; r++) {
      const row = ws.getRow(r);
      if (!row.cellCount) continue;
      let nameCol = 0, codeCol = 0;
      const maxC = Math.max(row.cellCount || 0, 30);
      for (let c = 1; c <= maxC; c++) {
        const v = normalize(String(row.getCell(c).value ?? ""));
        if (!v) continue;
        if (!nameCol && (NAME_HINTS.includes(v) || (v.includes("name") && !v.includes("file")))) nameCol = c;
        if (!codeCol && (CODE_HINTS.includes(v) || /\bcode\b/.test(v) || /\bemp.*id\b/.test(v))) codeCol = c;
      }
      if (nameCol) {
        let hasValues = 0;
        for (let rr = r + 1; rr <= Math.min(r + 50, ws.rowCount || r + 50); rr++) {
          if (String(ws.getRow(rr).getCell(nameCol).value ?? "").trim()) hasValues++;
        }
        if (hasValues >= 1) {
          return {
            ws, headerRow: r, nameCol,
            codeCol: codeCol || null,
            dataStartRow: r + 1,
            dataEndRow: ws.rowCount || r + 1,
          };
        }
      }
    }
  }
  const ws = wb.worksheets[0];
  return { ws, headerRow: 1, nameCol: 1, codeCol: null, dataStartRow: 2, dataEndRow: ws.rowCount || 2 };
}

export async function processExcel(
  excelFile: File,
  options: ProcessOptions
): Promise<ProcessReport> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await excelFile.arrayBuffer());

  const det = detectActiveSheet(wb);
  const ws = det.ws;

  // 1. Collect employees
  interface Emp {
    row: number;
    name: string;
    code: string | null;
    nameKey: string;
    fuzzy: string;
  }
  const employees: Emp[] = [];
  let lastDataRow = det.dataStartRow - 1;
  for (let r = det.dataStartRow; r <= Math.max(det.dataEndRow, det.dataStartRow + 1000); r++) {
    const nameVal = String(ws.getRow(r).getCell(det.nameCol).value ?? "").trim();
    if (!nameVal) {
      if (r - lastDataRow > 5) break;
      continue;
    }
    const codeVal = det.codeCol
      ? String(ws.getRow(r).getCell(det.codeCol).value ?? "").trim().toUpperCase() || null
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

  if (employees.length === 0) {
    throw new Error("No employees found in the Excel sheet. Make sure the file has a 'Name' column.");
  }

  // 2. Days covered (use day-of-month numbers)
  const dayMap = new Map<number, string>(); // day -> YYYY-MM-DD (last seen)
  for (const r of options.pdfResults) dayMap.set(r.day, r.date);
  const days = [...dayMap.keys()].sort((a, b) => a - b);
  const dates = days.map((d) => dayMap.get(d)!);

  // 3. Build per-day match: day -> empRow -> count (hybrid match)
  const matchByEmp = new Map<number, Map<number, number>>(); // empIndex -> day -> count
  const unmatched = new Set<string>();

  const empByCode = new Map<string, number>();
  employees.forEach((e, i) => { if (e.code) empByCode.set(e.code, i); });
  const empByName = new Map<string, number>();
  employees.forEach((e, i) => empByName.set(e.nameKey, i));

  function findEmployee(row: { code: string | null; name: string }): number {
    if (row.code) {
      const ci = empByCode.get(row.code.toUpperCase());
      if (ci !== undefined) return ci;
    }
    const k = normalize(row.name);
    const ni = empByName.get(k);
    if (ni !== undefined) return ni;
    // fuzzy: best similarity >= 0.85
    const target = fuzzyKey(row.name);
    let best = -1, bestScore = 0;
    for (let i = 0; i < employees.length; i++) {
      const s = similarity(target, employees[i].fuzzy);
      if (s > bestScore) { bestScore = s; best = i; }
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
      const m = matchByEmp.get(idx) ?? new Map<number, number>();
      m.set(pdf.day, row.count);
      matchByEmp.set(idx, m);
    }
  }

  // 4. Layout: keep codeCol (if present) + nameCol; write day cols + Total
  const firstWriteCol = det.nameCol; // we keep name col where it is
  const codeColFinal = det.codeCol; // may be null
  const firstDayCol = det.nameCol + 1;
  const totalCol = firstDayCol + days.length;

  // Compute totals (we'll write SUM formulas, but we also need values for sorting & coloring)
  const empTotals = employees.map((e, i) => {
    const m = matchByEmp.get(i);
    let total = 0;
    for (const d of days) total += m?.get(d) ?? 0;
    return { ...e, idx: i, total };
  });

  // 5. Sort: highest total first
  const sorted = [...empTotals].sort((a, b) => b.total - a.total);

  // 6. Headers
  const headerRowObj = ws.getRow(det.headerRow);
  for (let c = firstDayCol; c <= firstDayCol + 80; c++) headerRowObj.getCell(c).value = null;
  days.forEach((d, idx) => {
    const cell = headerRowObj.getCell(firstDayCol + idx);
    cell.value = d; // day number 1..31
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3B5BDB" } };
    cell.border = thinBorder();
    // tooltip with full date
    const fullDate = dayMap.get(d);
    if (fullDate) cell.note = fullDate;
  });
  const totalCell = headerRowObj.getCell(totalCol);
  totalCell.value = "Total";
  totalCell.alignment = { horizontal: "center", vertical: "middle" };
  totalCell.font = { bold: true, color: { argb: "FFFFFFFF" } };
  totalCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F2A4D" } };
  totalCell.border = thinBorder();

  // Style code + name headers
  if (codeColFinal) styleHeader(headerRowObj.getCell(codeColFinal), "Employee Code");
  styleHeader(headerRowObj.getCell(det.nameCol), "Employee Name");
  headerRowObj.height = 24;
  headerRowObj.commit?.();

  // 7. Clear data area then re-write sorted
  const writeStart = det.dataStartRow;
  const colStart = codeColFinal ? Math.min(codeColFinal, det.nameCol) : det.nameCol;
  for (let r = writeStart; r <= Math.max(lastDataRow, writeStart + sorted.length) + 5; r++) {
    const row = ws.getRow(r);
    for (let c = colStart; c <= totalCol; c++) {
      row.getCell(c).value = null;
      row.getCell(c).fill = { type: "pattern", pattern: "none" } as any;
    }
  }

  const totals = sorted.map((s) => s.total);
  const maxT = Math.max(...totals, 0);
  const avgT = totals.length ? totals.reduce((a, b) => a + b, 0) / totals.length : 0;
  const highThr = Math.max(avgT, maxT * 0.7);
  const lowThr = avgT * 0.4;

  sorted.forEach((emp, idx) => {
    const r = writeStart + idx;
    const row = ws.getRow(r);

    if (codeColFinal) {
      const cc = row.getCell(codeColFinal);
      cc.value = emp.code ?? "";
      cc.alignment = { horizontal: "left", vertical: "middle" };
      cc.border = thinBorder();
    }

    const nameCell = row.getCell(det.nameCol);
    nameCell.value = emp.name;
    nameCell.font = { bold: true };
    nameCell.alignment = { horizontal: "left", vertical: "middle" };
    nameCell.border = thinBorder();

    days.forEach((d, di) => {
      const cell = row.getCell(firstDayCol + di);
      const v = matchByEmp.get(emp.idx)?.get(d) ?? 0;
      cell.value = v;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = thinBorder();
      if (v === 0) cell.font = { color: { argb: "FFB0B0B0" } };
    });

    // SUM formula for Total
    const firstColLetter = colLetter(firstDayCol);
    const lastColLetter = colLetter(firstDayCol + days.length - 1);
    const tCell = row.getCell(totalCol);
    tCell.value = days.length
      ? { formula: `SUM(${firstColLetter}${r}:${lastColLetter}${r})`, result: emp.total }
      : 0;
    tCell.font = { bold: true };
    tCell.alignment = { horizontal: "center", vertical: "middle" };
    tCell.border = thinBorder();

    let bg = "FFFFF4CC"; // yellow
    if (emp.total >= highThr && emp.total > 0) bg = "FFD4F5DD"; // green
    else if (emp.total <= lowThr) bg = "FFFCD9D9"; // red
    tCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };

    if (idx % 2 === 1) {
      for (let c = colStart; c < totalCol; c++) {
        const cc = row.getCell(c);
        if (!cc.fill || (cc.fill as any).pattern === "none") {
          cc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF7F8FB" } } as any;
        }
      }
    }
    row.height = 20;
    row.commit?.();
  });

  // 8. Column widths
  if (codeColFinal) ws.getColumn(codeColFinal).width = 14;
  ws.getColumn(det.nameCol).width = Math.max(
    22,
    Math.min(40, Math.max(...sorted.map((s) => s.name.length + 2)))
  );
  for (let i = 0; i < days.length; i++) ws.getColumn(firstDayCol + i).width = 6;
  ws.getColumn(totalCol).width = 10;

  // 9. Freeze header row + name col
  ws.views = [{ state: "frozen", xSplit: det.nameCol, ySplit: det.headerRow, activeCell: "A1" }];

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const matched = sorted.filter((s) => s.total > 0).length;
  return {
    totalEmployees: employees.length,
    matchedEmployees: matched,
    unmatchedNames: [...unmatched].sort(),
    dates,
    days,
    blob,
    fileName: excelFile.name,
    preview: sorted.slice(0, 10).map((s) => ({ name: s.name, total: s.total })),
  };
}

function thinBorder(): Partial<ExcelJS.Borders> {
  const c = { style: "thin" as const, color: { argb: "FFE2E6EE" } };
  return { top: c, left: c, right: c, bottom: c };
}

function styleHeader(cell: ExcelJS.Cell, fallback: string) {
  cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
  cell.alignment = { horizontal: "left", vertical: "middle" };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F2A4D" } };
  cell.border = thinBorder();
  if (!String(cell.value ?? "").trim()) cell.value = fallback;
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
