import ExcelJS from "exceljs";

export interface MonthlyPlannedResult {
  totalEmployees: number;
  matchedEmployees: number;
  unmatchedEmployees: string[];
  warnings: string[];
  debug: {
    sourceRows: number;
    faceToFaceRows: number;
    contactPointRows: number;
    templateRows: number;
  };
  blob: Blob;
  fileName: string;
  sheetName: string;
  preview: { name: string; total: number }[];
}

interface MonthlySummary {
  code: string;
  name: string;
  days: Set<string>;
  cpTimes: number[];
  planned: number;
  unplanned: number;
  totalCalls: number;
}

interface MonthlyColumns {
  codeCol: number;
  nameCol: number;
  plannedCol: number;
  plannedAvgCol: number;
  unplannedCol: number;
  unplannedAvgCol: number;
  totalCallsCol: number;
  totalCallsAvgCol: number;
  cpAvgTimeCol: number;
}

const CALL_HEADER_HINTS: Record<string, string[]> = {
  code: ["emp. id", "emp id", "employee id", "employee code"],
  name: ["employee name", "name"],
  date: ["date"],
  startTime: ["start time"],
  eventType: ["event type"],
  meetingType: ["meeting type"],
};

const TEMPLATE_HEADER_HINTS: Record<string, string[]> = {
  code: ["employee code", "emp code", "employee id", "code"],
  name: ["name", "employee name"],
};

export async function processMonthlyPlannedReport(
  callLogFile: File,
  templateFile: File,
): Promise<MonthlyPlannedResult> {
  const callLog = await readMonthlyCallLog(callLogFile);
  const summariesByCode = new Map(callLog.summaries.map((item) => [item.code, item]));
  const unmatched = new Set(callLog.summaries.map((item) => `${item.code} ${item.name}`));

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await templateFile.arrayBuffer());
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("Sample workbook does not contain a sheet.");

  const columns = ensureMonthlyColumns(sheet);
  const preview: { name: string; total: number }[] = [];
  let matchedEmployees = 0;
  let templateRows = 0;

  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const code = normalizeEmployeeCode(row.getCell(columns.codeCol).text);
    const name = row.getCell(columns.nameCol).text;
    if (!code) continue;
    templateRows++;

    const match = summariesByCode.get(code);
    if (!match) {
      clearMonthlyCells(row, columns);
      continue;
    }

    const days = match.days.size;
    matchedEmployees++;
    unmatched.delete(`${match.code} ${match.name}`);

    row.getCell(columns.plannedCol).value = match.planned;
    row.getCell(columns.plannedAvgCol).value = average(match.planned, days);
    row.getCell(columns.unplannedCol).value = match.unplanned;
    row.getCell(columns.unplannedAvgCol).value = average(match.unplanned, days);
    row.getCell(columns.totalCallsCol).value = match.totalCalls;
    row.getCell(columns.totalCallsAvgCol).value = average(match.totalCalls, days);
    row.getCell(columns.cpAvgTimeCol).value = averageTime(match.cpTimes);
    styleMonthlyCells(row, columns);

    if (match.totalCalls > 0) preview.push({ name: name || match.name, total: match.totalCalls });
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
      ? [`${unmatched.size} monthly call log employee(s) were not found in the sample file.`]
      : [],
    debug: {
      sourceRows: callLog.sourceRows,
      faceToFaceRows: callLog.faceToFaceRows,
      contactPointRows: callLog.contactPointRows,
      templateRows,
    },
    blob,
    fileName: templateFile.name.replace(/\.xlsx$/i, "") + " - Monthly Planned Unplanned.xlsx",
    sheetName: sheet.name,
    preview: preview.sort((a, b) => b.total - a.total).slice(0, 10),
  };
}

async function readMonthlyCallLog(file: File): Promise<{
  summaries: MonthlySummary[];
  sourceRows: number;
  faceToFaceRows: number;
  contactPointRows: number;
}> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("Call log workbook does not contain a sheet.");

  const headerRow = sheet.getRow(1);
  const headers = new Map<string, number>();
  for (let c = 1; c <= sheet.columnCount; c++) headers.set(normalize(headerRow.getCell(c).text), c);

  const codeCol = findHeader(headers, CALL_HEADER_HINTS.code);
  const nameCol = findHeader(headers, CALL_HEADER_HINTS.name);
  const dateCol = findHeader(headers, CALL_HEADER_HINTS.date);
  const startTimeCol = findHeader(headers, CALL_HEADER_HINTS.startTime);
  const eventTypeCol = findHeader(headers, CALL_HEADER_HINTS.eventType);
  const meetingTypeCol = findHeader(headers, CALL_HEADER_HINTS.meetingType);

  const byCode = new Map<string, MonthlySummary>();
  let sourceRows = 0;
  let faceToFaceRows = 0;
  let contactPointRows = 0;

  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const code = normalizeEmployeeCode(row.getCell(codeCol).text);
    if (!code) continue;

    const meetingType = normalize(row.getCell(meetingTypeCol).text);
    if (meetingType !== "face to face call" && meetingType !== "contact point") continue;

    const name = row.getCell(nameCol).text.trim();
    const eventType = normalize(row.getCell(eventTypeCol).text);
    const dateKey = normalizeDateKey(row.getCell(dateCol).value, row.getCell(dateCol).text);
    let summary = byCode.get(code);
    if (!summary) {
      summary = {
        code,
        name,
        days: new Set<string>(),
        cpTimes: [],
        planned: 0,
        unplanned: 0,
        totalCalls: 0,
      };
      byCode.set(code, summary);
    }

    sourceRows++;
    if (dateKey) summary.days.add(dateKey);

    if (meetingType === "contact point") {
      contactPointRows++;
      const cpTime = timeToMinutes(row.getCell(startTimeCol).value, row.getCell(startTimeCol).text);
      if (cpTime > 0) summary.cpTimes.push(cpTime);
      continue;
    }

    faceToFaceRows++;
    if (eventType === "planned") summary.planned++;
    if (eventType === "unplanned") summary.unplanned++;
    summary.totalCalls = summary.planned + summary.unplanned;
  }

  return { summaries: [...byCode.values()], sourceRows, faceToFaceRows, contactPointRows };
}

function ensureMonthlyColumns(sheet: ExcelJS.Worksheet): MonthlyColumns {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const labels = new Map<string, number>();
    for (let c = 1; c <= sheet.columnCount; c++) labels.set(normalize(row.getCell(c).text), c);

    const codeCol = findHeader(labels, TEMPLATE_HEADER_HINTS.code, false);
    const nameCol = findHeader(labels, TEMPLATE_HEADER_HINTS.name, false);
    if (!codeCol || !nameCol) continue;

    const lastUsedCol = Math.max(findLastUsedColumn(sheet), nameCol);
    const columns: Array<[keyof MonthlyColumns, string]> = [
      ["plannedCol", "Total Planned"],
      ["plannedAvgCol", "Planned Avg"],
      ["unplannedCol", "Total Unplanned"],
      ["unplannedAvgCol", "Unplanned Avg"],
      ["totalCallsCol", "Total Calls"],
      ["totalCallsAvgCol", "Total Calls Avg"],
      ["cpAvgTimeCol", "CP Avg Time"],
    ];

    const result: MonthlyColumns = {
      codeCol,
      nameCol,
      plannedCol: 0,
      plannedAvgCol: 0,
      unplannedCol: 0,
      unplannedAvgCol: 0,
      totalCallsCol: 0,
      totalCallsAvgCol: 0,
      cpAvgTimeCol: 0,
    };

    const titleRow = sheet.getRow(Math.max(1, rowNumber - 1));
    for (let index = 0; index < columns.length; index++) {
      const [key, label] = columns[index];
      const col = lastUsedCol + index + 1;
      result[key] = col;
      const titleCell = titleRow.getCell(col);
      titleCell.value = "Monthly Planned Unplanned";
      styleHeaderCell(titleCell);
      const cell = row.getCell(col);
      cell.value = label;
      styleHeaderCell(cell);
      sheet.getColumn(col).width = Math.max(label.length + 4, 13);
    }
    titleRow.commit();
    row.commit();
    return result;
  }

  throw new Error("Sample columns were not found. Required: Employee Code and Name.");
}

function clearMonthlyCells(row: ExcelJS.Row, columns: MonthlyColumns) {
  for (const col of monthlyValueColumns(columns)) row.getCell(col).value = null;
}

function styleMonthlyCells(row: ExcelJS.Row, columns: MonthlyColumns) {
  for (const col of monthlyValueColumns(columns)) {
    const cell = row.getCell(col);
    cell.alignment = { ...(cell.alignment ?? {}), horizontal: "center", vertical: "middle" };
    cell.border = thinBorder();
    if (col !== columns.cpAvgTimeCol) cell.numFmt = "0";
  }
}

function monthlyValueColumns(columns: MonthlyColumns): number[] {
  return [
    columns.plannedCol,
    columns.plannedAvgCol,
    columns.unplannedCol,
    columns.unplannedAvgCol,
    columns.totalCallsCol,
    columns.totalCallsAvgCol,
    columns.cpAvgTimeCol,
  ];
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

function styleHeaderCell(cell: ExcelJS.Cell) {
  cell.font = { ...(cell.font ?? {}), bold: true, color: { argb: "FF000000" } };
  cell.alignment = { ...(cell.alignment ?? {}), horizontal: "center", vertical: "middle" };
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF00" },
  };
  cell.border = thinBorder();
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

function average(total: number, days: number): number {
  return days > 0 ? Math.round(total / days) : 0;
}

function averageTime(minutes: number[]): string {
  if (!minutes.length) return "";
  const averageMinutes = Math.round(
    minutes.reduce((sum, minute) => sum + minute, 0) / minutes.length,
  );
  return minutesToTime(averageMinutes);
}

function timeToMinutes(value: unknown, text: string): number {
  if (value instanceof Date) return value.getHours() * 60 + value.getMinutes();
  const match = String(text ?? "").match(/\b(\d{1,2}):(\d{2})(?::\d{2})?\b/);
  if (!match) return -1;
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (!Number.isFinite(hour) || !Number.isFinite(minute)) return -1;
  return hour * 60 + minute;
}

function minutesToTime(totalMinutes: number): string {
  const normalized = ((totalMinutes % 1440) + 1440) % 1440;
  let hour = Math.floor(normalized / 60);
  const minute = normalized % 60;
  const suffix = hour >= 12 ? "PM" : "AM";
  hour = hour % 12 || 12;
  return `${hour}:${String(minute).padStart(2, "0")} ${suffix}`;
}

function normalize(value: string): string {
  return String(value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
}

function normalizeEmployeeCode(value: string): string {
  const digits = String(value ?? "").match(/\d+/g)?.join("") ?? "";
  return digits.length >= 4 ? digits : "";
}

function normalizeDateKey(value: unknown, text: string): string {
  const date = value instanceof Date ? value : new Date(text);
  if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  return normalize(text);
}
