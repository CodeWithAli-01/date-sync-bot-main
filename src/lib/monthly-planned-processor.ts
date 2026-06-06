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
  performanceRows: MonthlyPlannedPerformanceRow[];
}

export interface MonthlyPlannedPerformanceRow {
  teamName: string;
  employeeCode: string;
  name: string;
  designation: string;
  planned: number;
  unplanned: number;
  totalCalls: number;
  plannedPercent: number;
  cpAvgTime: string;
  topQualified: boolean;
}

interface MonthlySummary {
  code: string;
  name: string;
  teamName: string;
  sourceName: string;
  days: Set<string>;
  cpTimes: number[];
  planned: number;
  unplanned: number;
  totalCalls: number;
}

interface MonthlyColumns {
  codeCol: number;
  nameCol: number;
  designationCol?: number;
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
  team: ["team name", "team", "team id", "team code"],
};

const TEMPLATE_HEADER_HINTS: Record<string, string[]> = {
  code: ["employee code", "emp code", "employee id", "code"],
  name: ["name", "employee name"],
  designation: ["designation", "desig", "position", "title"],
};

export async function processMonthlyPlannedReport(
  callLogFile: File | File[],
  templateFile: File,
): Promise<MonthlyPlannedResult> {
  const callLog = await readMonthlyCallLog(callLogFile);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await templateFile.arrayBuffer());
  const candidateSheets = workbook.worksheets.filter((sheet) => findTemplateIdentityColumns(sheet));
  const sheets =
    candidateSheets.length > 1 ? candidateSheets : [workbook.worksheets[0]].filter(Boolean);
  if (!sheets.length) throw new Error("Sample workbook does not contain a sheet.");

  const unmatched = new Set(callLog.summaries.map(monthlySummaryKey));
  const preview: { name: string; total: number }[] = [];
  const performanceRows: MonthlyPlannedPerformanceRow[] = [];
  let matchedEmployees = 0;
  let templateRows = 0;
  let firstSheetName = sheets[0].name;

  for (const sheet of sheets) {
    const summaries = filterMonthlySummariesForSheet(callLog.summaries, sheet, sheets.length > 1);
    const result = fillMonthlyPlannedSheet(sheet, summaries);
    firstSheetName = firstSheetName || sheet.name;
    matchedEmployees += result.matchedEmployees;
    templateRows += result.templateRows;
    result.matchedKeys.forEach((key) => {
      unmatched.delete(key);
    });
    preview.push(...result.preview);
    performanceRows.push(...result.performanceRows);
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
    sheetName: firstSheetName,
    preview: preview.sort((a, b) => b.total - a.total).slice(0, 10),
    performanceRows,
  };
}

function fillMonthlyPlannedSheet(sheet: ExcelJS.Worksheet, summaries: MonthlySummary[]) {
  const summariesByCode = new Map(summaries.map((item) => [item.code, item]));
  const columns = ensureMonthlyColumns(sheet);
  const matchedKeys = new Set<string>();
  const preview: { name: string; total: number }[] = [];
  const performanceRows: MonthlyPlannedPerformanceRow[] = [];
  let matchedEmployees = 0;
  let templateRows = 0;

  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const code = normalizeEmployeeCode(cellText(row.getCell(columns.codeCol)));
    const name = cellText(row.getCell(columns.nameCol));
    if (!code) continue;
    templateRows++;

    const match = summariesByCode.get(code);
    if (!match) {
      clearMonthlyCells(row, columns);
      continue;
    }

    const days = match.days.size;
    matchedEmployees++;
    matchedKeys.add(monthlySummaryKey(match));

    row.getCell(columns.plannedCol).value = match.planned;
    row.getCell(columns.plannedAvgCol).value = average(match.planned, days);
    row.getCell(columns.unplannedCol).value = match.unplanned;
    row.getCell(columns.unplannedAvgCol).value = average(match.unplanned, days);
    row.getCell(columns.totalCallsCol).value = match.totalCalls;
    row.getCell(columns.totalCallsAvgCol).value = average(match.totalCalls, days);
    row.getCell(columns.cpAvgTimeCol).value = averageTime(match.cpTimes);
    styleMonthlyCells(row, columns);

    if (match.totalCalls > 0) preview.push({ name: name || match.name, total: match.totalCalls });
    performanceRows.push({
      teamName: match.teamName || sheet.name,
      employeeCode: code,
      name: name || match.name,
      designation: columns.designationCol ? cellText(row.getCell(columns.designationCol)) : "",
      planned: match.planned,
      unplanned: match.unplanned,
      totalCalls: match.totalCalls,
      plannedPercent: percent(match.planned, match.totalCalls),
      cpAvgTime: averageTime(match.cpTimes),
      topQualified: percent(match.planned, match.totalCalls) >= 70,
    });
  }

  return { matchedEmployees, templateRows, matchedKeys, preview, performanceRows };
}

async function readMonthlyCallLog(input: File | File[]): Promise<{
  summaries: MonthlySummary[];
  sourceRows: number;
  faceToFaceRows: number;
  contactPointRows: number;
}> {
  const files = Array.isArray(input) ? input : [input];
  const byCode = new Map<string, MonthlySummary>();
  let sourceRows = 0;
  let faceToFaceRows = 0;
  let contactPointRows = 0;

  for (const file of files) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
    const sheets = workbook.worksheets.filter((sheet) => findMonthlyCallLogColumns(sheet));
    const sourceSheets = sheets.length ? sheets : workbook.worksheets.slice(0, 1);
    if (!sourceSheets.length) throw new Error("Call log workbook does not contain a sheet.");

    for (const sheet of sourceSheets) {
      const columns = findMonthlyCallLogColumns(sheet);
      if (!columns) {
        throw new Error(
          "Call log columns were not found. Required: Emp ID, Name, Date, Start Time, Event Type, and Meeting Type.",
        );
      }

      for (let rowNumber = columns.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber++) {
        const row = sheet.getRow(rowNumber);
        const code = normalizeEmployeeCode(cellText(row.getCell(columns.codeCol)));
        if (!code) continue;

        const meetingType = normalize(cellText(row.getCell(columns.meetingTypeCol)));
        if (meetingType !== "face to face call" && meetingType !== "contact point") continue;

        const name = cellText(row.getCell(columns.nameCol)).trim();
        const eventType = normalize(cellText(row.getCell(columns.eventTypeCol)));
        const dateKey = normalizeDateKey(
          row.getCell(columns.dateCol).value,
          cellText(row.getCell(columns.dateCol)),
        );
        const teamName = columns.teamCol ? cellText(row.getCell(columns.teamCol)).trim() : "";
        const sourceName = `${file.name} ${sheet.name}`;
        const summaryKey = `${normalizeTeamName(teamName || sourceName)}|${code}`;
        let summary = byCode.get(summaryKey);
        if (!summary) {
          summary = {
            code,
            name,
            teamName: teamName || sheet.name,
            sourceName,
            days: new Set<string>(),
            cpTimes: [],
            planned: 0,
            unplanned: 0,
            totalCalls: 0,
          };
          byCode.set(summaryKey, summary);
        }

        sourceRows++;
        if (dateKey) summary.days.add(dateKey);

        if (meetingType === "contact point") {
          contactPointRows++;
          const cpTime = timeToMinutes(
            row.getCell(columns.startTimeCol).value,
            cellText(row.getCell(columns.startTimeCol)),
          );
          if (cpTime > 0) summary.cpTimes.push(cpTime);
          continue;
        }

        faceToFaceRows++;
        if (eventType === "planned") summary.planned++;
        if (eventType === "unplanned") summary.unplanned++;
        summary.totalCalls = summary.planned + summary.unplanned;
      }
    }
  }

  return { summaries: [...byCode.values()], sourceRows, faceToFaceRows, contactPointRows };
}

function filterMonthlySummariesForSheet(
  summaries: MonthlySummary[],
  sheet: ExcelJS.Worksheet,
  isBulkTemplate: boolean,
): MonthlySummary[] {
  if (!isBulkTemplate) return summaries;
  const teamSummaries = summaries.filter((summary) =>
    [summary.teamName, summary.sourceName].some((value) => teamNamesMatch(value, sheet.name)),
  );
  return teamSummaries.length ? teamSummaries : summaries;
}

function monthlySummaryKey(summary: MonthlySummary): string {
  return `${summary.sourceName}|${summary.teamName}|${summary.code}`;
}

function findMonthlyCallLogColumns(sheet: ExcelJS.Worksheet) {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 10); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const headers = new Map<string, number>();
    for (let c = 1; c <= sheet.columnCount; c++)
      headers.set(normalize(cellText(row.getCell(c))), c);
    const codeCol = findHeader(headers, CALL_HEADER_HINTS.code, false);
    const nameCol = findHeader(headers, CALL_HEADER_HINTS.name, false);
    const dateCol = findHeader(headers, CALL_HEADER_HINTS.date, false);
    const startTimeCol = findHeader(headers, CALL_HEADER_HINTS.startTime, false);
    const eventTypeCol = findHeader(headers, CALL_HEADER_HINTS.eventType, false);
    const meetingTypeCol = findHeader(headers, CALL_HEADER_HINTS.meetingType, false);
    const teamCol = findHeader(headers, CALL_HEADER_HINTS.team, false);
    if (codeCol && nameCol && dateCol && startTimeCol && eventTypeCol && meetingTypeCol) {
      return {
        headerRow: rowNumber,
        codeCol,
        nameCol,
        dateCol,
        startTimeCol,
        eventTypeCol,
        meetingTypeCol,
        teamCol,
      };
    }
  }
  return null;
}

function findTemplateIdentityColumns(sheet: ExcelJS.Worksheet): boolean {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const labels = new Map<string, number>();
    for (let c = 1; c <= sheet.columnCount; c++) labels.set(normalize(cellText(row.getCell(c))), c);
    if (
      findHeader(labels, TEMPLATE_HEADER_HINTS.code, false) &&
      findHeader(labels, TEMPLATE_HEADER_HINTS.name, false)
    ) {
      return true;
    }
  }
  return false;
}

function ensureMonthlyColumns(sheet: ExcelJS.Worksheet): MonthlyColumns {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const labels = new Map<string, number>();
    for (let c = 1; c <= sheet.columnCount; c++) labels.set(normalize(cellText(row.getCell(c))), c);

    const codeCol = findHeader(labels, TEMPLATE_HEADER_HINTS.code, false);
    const nameCol = findHeader(labels, TEMPLATE_HEADER_HINTS.name, false);
    const designationCol = findHeader(labels, TEMPLATE_HEADER_HINTS.designation, false);
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
      designationCol: designationCol || undefined,
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
      if (cellText(cell).trim()) lastUsedCol = Math.max(lastUsedCol, colNumber);
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

function percent(part: number, total: number): number {
  return total > 0 ? Math.round((part / total) * 100) : 0;
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
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function cellText(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (value == null) return "";
  if (typeof value === "object") {
    if ("text" in value) return String((value as { text?: unknown }).text ?? "");
    if ("result" in value) return String((value as { result?: unknown }).result ?? "");
    if ("richText" in value && Array.isArray((value as { richText?: unknown }).richText)) {
      return (value as { richText: Array<{ text?: string }> }).richText
        .map((part) => part.text ?? "")
        .join("");
    }
    return "";
  }
  return String(value);
}

function normalizeEmployeeCode(value: string): string {
  const digits =
    String(value ?? "")
      .match(/\d+/g)
      ?.join("") ?? "";
  return digits.length >= 3 ? digits : "";
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

function normalizeDateKey(value: unknown, text: string): string {
  const date = value instanceof Date ? value : new Date(text);
  if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  return normalize(text);
}
