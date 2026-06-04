import ExcelJS from "exceljs";

export interface DailyReportResult {
  totalEmployees: number;
  matchedEmployees: number;
  unmatchedEmployees: string[];
  warnings: string[];
  debug: {
    callRows: number;
    faceToFaceRows: number;
    contactPointRows: number;
    templateRows: number;
  };
  blob: Blob;
  fileName: string;
  sheetName: string;
  preview: { name: string; total: number }[];
}

interface CallSummary {
  code: string;
  name: string;
  planned: number;
  unplanned: number;
  morning: number;
  evening: number;
  total: number;
  cpTime: string;
}

interface DailyColumns {
  codeCol: number;
  nameCol: number;
  plannedCol: number;
  unplannedCol: number;
  morningCol: number;
  eveningCol: number;
  totalCol: number;
  cpCol: number;
}

const CALL_HEADER_HINTS: Record<string, string[]> = {
  code: ["emp. id", "emp id", "employee id", "employee code"],
  name: ["employee name", "name"],
  startTime: ["start time"],
  eventType: ["event type"],
  meetingType: ["meeting type"],
  shift: ["shift"],
};

const TEMPLATE_HEADER_HINTS: Record<string, string[]> = {
  code: ["employee code", "emp code", "code"],
  name: ["name", "employee name"],
  planned: ["planned"],
  unplanned: ["unplanned"],
  morning: ["mor", "morning"],
  evening: ["eve", "evening"],
  total: ["total"],
  cp: ["cp"],
};

export async function processDailyReport(
  callLogFile: File,
  templateFile: File,
): Promise<DailyReportResult> {
  const callLog = await readCallLog(callLogFile);
  const templateWorkbook = new ExcelJS.Workbook();
  await templateWorkbook.xlsx.load(await templateFile.arrayBuffer());
  const sheet = templateWorkbook.worksheets[0];
  if (!sheet) throw new Error("Template workbook does not contain a sheet.");

  const columns = findTemplateColumns(sheet);
  const summariesByCode = new Map(callLog.summaries.map((item) => [item.code, item]));
  const unmatched = new Set(callLog.summaries.map((item) => `${item.code} ${item.name}`));
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
      clearDailyCells(row, columns);
      continue;
    }

    matchedEmployees++;
    unmatched.delete(`${match.code} ${match.name}`);
    row.getCell(columns.plannedCol).value = match.planned || 0;
    row.getCell(columns.unplannedCol).value = match.unplanned || 0;
    row.getCell(columns.morningCol).value = match.morning || 0;
    row.getCell(columns.eveningCol).value = match.evening || 0;
    row.getCell(columns.totalCol).value = match.total || 0;
    row.getCell(columns.cpCol).value = match.cpTime || null;

    styleDailyCells(row, columns);
    if (match.total > 0) preview.push({ name: name || match.name, total: match.total });
  }

  const buffer = await templateWorkbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  return {
    totalEmployees: templateRows,
    matchedEmployees,
    unmatchedEmployees: [...unmatched],
    warnings: unmatched.size ? [`${unmatched.size} call log employee(s) were not found in the template.`] : [],
    debug: {
      callRows: callLog.callRows,
      faceToFaceRows: callLog.faceToFaceRows,
      contactPointRows: callLog.contactPointRows,
      templateRows,
    },
    blob,
    fileName: templateFile.name.replace(/\.xlsx$/i, "") + " - Daily Report.xlsx",
    sheetName: sheet.name,
    preview: preview.sort((a, b) => b.total - a.total).slice(0, 10),
  };
}

async function readCallLog(file: File): Promise<{
  summaries: CallSummary[];
  callRows: number;
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
  const startTimeCol = findHeader(headers, CALL_HEADER_HINTS.startTime);
  const eventTypeCol = findHeader(headers, CALL_HEADER_HINTS.eventType);
  const meetingTypeCol = findHeader(headers, CALL_HEADER_HINTS.meetingType);
  const shiftCol = findHeader(headers, CALL_HEADER_HINTS.shift);

  const byCode = new Map<string, CallSummary>();
  let callRows = 0;
  let faceToFaceRows = 0;
  let contactPointRows = 0;

  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const code = normalizeEmployeeCode(row.getCell(codeCol).text);
    if (!code) continue;

    const name = row.getCell(nameCol).text.trim();
    const meetingType = normalize(row.getCell(meetingTypeCol).text);
    const eventType = normalize(row.getCell(eventTypeCol).text);
    const shift = normalize(row.getCell(shiftCol).text);
    const startTime = formatTime(row.getCell(startTimeCol).value, row.getCell(startTimeCol).text);

    let summary = byCode.get(code);
    if (!summary) {
      summary = {
        code,
        name,
        planned: 0,
        unplanned: 0,
        morning: 0,
        evening: 0,
        total: 0,
        cpTime: "",
      };
      byCode.set(code, summary);
    }

    callRows++;
    if (meetingType === "contact point") {
      contactPointRows++;
      if (!summary.cpTime && startTime) summary.cpTime = startTime;
      continue;
    }

    if (meetingType && meetingType !== "face to face call") continue;
    faceToFaceRows++;
    if (eventType === "planned") summary.planned++;
    if (eventType === "unplanned") summary.unplanned++;
    if (shift === "morning") summary.morning++;
    if (shift === "evening") summary.evening++;
    summary.total = summary.morning + summary.evening;
  }

  return { summaries: [...byCode.values()], callRows, faceToFaceRows, contactPointRows };
}

function findTemplateColumns(sheet: ExcelJS.Worksheet): DailyColumns {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const labels = new Map<string, number>();
    for (let c = 1; c <= sheet.columnCount; c++) labels.set(normalize(row.getCell(c).text), c);

    const codeCol = findHeader(labels, TEMPLATE_HEADER_HINTS.code, false);
    const nameCol = findHeader(labels, TEMPLATE_HEADER_HINTS.name, false);
    const plannedCol = findHeader(labels, TEMPLATE_HEADER_HINTS.planned, false);
    const unplannedCol = findHeader(labels, TEMPLATE_HEADER_HINTS.unplanned, false);
    const morningCol = findHeader(labels, TEMPLATE_HEADER_HINTS.morning, false);
    const eveningCol = findHeader(labels, TEMPLATE_HEADER_HINTS.evening, false);
    const totalCol = findHeader(labels, TEMPLATE_HEADER_HINTS.total, false);
    const cpCol = findHeader(labels, TEMPLATE_HEADER_HINTS.cp, false);

    if (
      codeCol &&
      nameCol &&
      plannedCol &&
      unplannedCol &&
      morningCol &&
      eveningCol &&
      totalCol &&
      cpCol
    ) {
      return { codeCol, nameCol, plannedCol, unplannedCol, morningCol, eveningCol, totalCol, cpCol };
    }
  }

  throw new Error("Template columns were not found. Required: Employee Code, Name, Planned, Unplanned, Mor, Eve, Total, Cp.");
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

function clearDailyCells(row: ExcelJS.Row, columns: DailyColumns) {
  for (const col of [
    columns.plannedCol,
    columns.unplannedCol,
    columns.morningCol,
    columns.eveningCol,
    columns.totalCol,
    columns.cpCol,
  ]) {
    row.getCell(col).value = null;
  }
}

function styleDailyCells(row: ExcelJS.Row, columns: DailyColumns) {
  for (const col of [
    columns.plannedCol,
    columns.unplannedCol,
    columns.morningCol,
    columns.eveningCol,
    columns.totalCol,
    columns.cpCol,
  ]) {
    const cell = row.getCell(col);
    cell.alignment = { ...(cell.alignment ?? {}), horizontal: "center", vertical: "middle" };
  }
}

function normalize(value: string): string {
  return String(value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
}

function normalizeEmployeeCode(value: string): string {
  const digits = String(value ?? "").match(/\d+/g)?.join("") ?? "";
  return digits.length >= 4 ? digits : "";
}

function formatTime(value: unknown, text: string): string {
  if (value instanceof Date) {
    return value.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
  }
  const match = String(text ?? "").match(/\b(\d{1,2}):(\d{2})(?::\d{2})?\b/);
  if (!match) return "";
  let hour = Number(match[1]);
  const minute = match[2];
  const suffix = hour >= 12 ? "PM" : "AM";
  hour = hour % 12 || 12;
  return `${hour}:${minute} ${suffix}`;
}
