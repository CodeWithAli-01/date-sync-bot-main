import ExcelJS from "exceljs";
import JSZip from "jszip";

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

export interface BulkDailyReportSummaryItem {
  teamName: string;
  fileName: string;
  status: "success" | "failed";
  totalEmployees: number;
  matchedEmployees: number;
  error?: string;
}

export interface BulkDailyReportResult {
  totalTeams: number;
  reportsGenerated: number;
  failedReports: number;
  blob: Blob;
  fileName: string;
  summary: BulkDailyReportSummaryItem[];
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
  headerRow: number;
  codeCol: number;
  nameCol: number;
  plannedCol: number;
  unplannedCol: number;
  morningCol: number;
  eveningCol: number;
  totalCol: number;
  cpCol: number;
  teamCol?: number;
}

type DailyTeamSource =
  | {
      kind: "sheet";
      teamName: string;
      sheetName: string;
    }
  | {
      kind: "column";
      teamName: string;
      sheetName: string;
      headerRow: number;
      teamCol: number;
    };

const CALL_HEADER_HINTS: Record<string, string[]> = {
  code: ["emp. id", "emp id", "employee id", "employee code"],
  name: ["employee name", "name"],
  startTime: ["start time"],
  eventType: ["event type"],
  meetingType: ["meeting type"],
  shift: ["shift"],
};

const TEMPLATE_HEADER_HINTS: Record<string, string[]> = {
  team: ["team name", "team id", "team", "team code", "group", "division"],
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

  return await processDailyWorkbook(callLog, templateWorkbook, sheet, templateFile.name);
}

export async function processBulkDailyReports(
  callLogFile: File,
  teamsWorkbookFile: File,
  onProgress?: (progress: { current: number; total: number; teamName: string }) => void,
): Promise<BulkDailyReportResult> {
  const callLog = await readCallLog(callLogFile);
  const teamsWorkbook = new ExcelJS.Workbook();
  await teamsWorkbook.xlsx.load(await teamsWorkbookFile.arrayBuffer());
  const teamSources = findDailyTeamSources(teamsWorkbook);

  if (!teamSources.length) {
    throw new Error(
      "No teams were found. Add a Team Name/Team ID column, or use one worksheet per team.",
    );
  }

  const zip = new JSZip();
  const summary: BulkDailyReportSummaryItem[] = [];
  const usedReportFileNames = new Set<string>();

  for (let index = 0; index < teamSources.length; index++) {
    const source = teamSources[index];
    onProgress?.({ current: index + 1, total: teamSources.length, teamName: source.teamName });

    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await teamsWorkbookFile.arrayBuffer());
      const sheet = prepareTeamWorksheet(workbook, source);
      const result = await processDailyWorkbook(
        callLog,
        workbook,
        sheet,
        teamsWorkbookFile.name,
        source.teamName,
      );
      const reportFileName = uniqueFileName(result.fileName, usedReportFileNames);
      zip.file(reportFileName, await result.blob.arrayBuffer());
      summary.push({
        teamName: source.teamName,
        fileName: reportFileName,
        status: "success",
        totalEmployees: result.totalEmployees,
        matchedEmployees: result.matchedEmployees,
      });
    } catch (error) {
      summary.push({
        teamName: source.teamName,
        fileName: "",
        status: "failed",
        totalEmployees: 0,
        matchedEmployees: 0,
        error: error instanceof Error ? error.message : "Unknown error",
      });
    }
  }

  const summaryWorkbook = buildBulkSummaryWorkbook(summary);
  zip.file("Bulk_Daily_Report_Summary.xlsx", await summaryWorkbook.xlsx.writeBuffer());

  const blob = await zip.generateAsync({
    type: "blob",
    mimeType: "application/zip",
  });

  const reportsGenerated = summary.filter((item) => item.status === "success").length;
  return {
    totalTeams: teamSources.length,
    reportsGenerated,
    failedReports: summary.length - reportsGenerated,
    blob,
    fileName: "Daily Reports Output.zip",
    summary,
  };
}

async function processDailyWorkbook(
  callLog: Awaited<ReturnType<typeof readCallLog>>,
  templateWorkbook: ExcelJS.Workbook,
  sheet: ExcelJS.Worksheet,
  templateFileName: string,
  teamName?: string,
): Promise<DailyReportResult> {
  const columns = findTemplateColumns(sheet);
  if (!columns) throw new Error("Template columns were not found.");
  const summariesByCode = new Map(callLog.summaries.map((item) => [item.code, item]));
  const unmatched = new Set(callLog.summaries.map((item) => `${item.code} ${item.name}`));
  const preview: { name: string; total: number }[] = [];
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
    warnings: unmatched.size
      ? [`${unmatched.size} call log employee(s) were not found in the template.`]
      : [],
    debug: {
      callRows: callLog.callRows,
      faceToFaceRows: callLog.faceToFaceRows,
      contactPointRows: callLog.contactPointRows,
      templateRows,
    },
    blob,
    fileName: teamName
      ? `${safeFileName(teamName)}_Report.xlsx`
      : templateFileName.replace(/\.xlsx$/i, "") + " - Daily Report.xlsx",
    sheetName: sheet.name,
    preview: preview.sort((a, b) => b.total - a.total).slice(0, 10),
  };
}

function findDailyTeamSources(workbook: ExcelJS.Workbook): DailyTeamSource[] {
  const sheetSources = workbook.worksheets
    .filter((sheet) => hasTemplateColumns(sheet))
    .map((sheet) => ({
      kind: "sheet" as const,
      teamName: sheet.name.trim(),
      sheetName: sheet.name,
    }))
    .filter((source) => source.teamName);

  if (sheetSources.length > 1) return sheetSources;

  for (const sheet of workbook.worksheets) {
    const columns = findTemplateColumns(sheet, false);
    if (!columns?.teamCol) continue;

    const teams = new Map<string, DailyTeamSource>();
    for (let rowNumber = columns.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber++) {
      const row = sheet.getRow(rowNumber);
      const teamName = cellText(row.getCell(columns.teamCol)).trim();
      const code = normalizeEmployeeCode(cellText(row.getCell(columns.codeCol)));
      if (!teamName || !code) continue;
      teams.set(normalize(teamName), {
        kind: "column",
        teamName,
        sheetName: sheet.name,
        headerRow: columns.headerRow,
        teamCol: columns.teamCol,
      });
    }

    if (teams.size) return [...teams.values()].sort((a, b) => a.teamName.localeCompare(b.teamName));
  }

  return sheetSources;
}

function prepareTeamWorksheet(
  workbook: ExcelJS.Workbook,
  source: DailyTeamSource,
): ExcelJS.Worksheet {
  const sheet = workbook.getWorksheet(source.sheetName);
  if (!sheet) throw new Error(`Team sheet not found: ${source.sheetName}`);

  for (const otherSheet of [...workbook.worksheets]) {
    if (otherSheet.id !== sheet.id) workbook.removeWorksheet(otherSheet.id);
  }

  sheet.name = safeSheetName(source.teamName);
  if (source.kind === "column") filterWorksheetToTeam(sheet, source);
  return sheet;
}

function filterWorksheetToTeam(
  sheet: ExcelJS.Worksheet,
  source: Extract<DailyTeamSource, { kind: "column" }>,
) {
  for (let rowNumber = sheet.rowCount; rowNumber > source.headerRow; rowNumber--) {
    const row = sheet.getRow(rowNumber);
    const teamName = cellText(row.getCell(source.teamCol)).trim();
    if (normalize(teamName) !== normalize(source.teamName)) {
      sheet.spliceRows(rowNumber, 1);
    }
  }
}

function hasTemplateColumns(sheet: ExcelJS.Worksheet): boolean {
  return Boolean(findTemplateColumns(sheet, false));
}

function buildBulkSummaryWorkbook(summary: BulkDailyReportSummaryItem[]): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Summary");
  const successful = summary.filter((item) => item.status === "success").length;
  const failed = summary.length - successful;

  sheet.addRow(["Total teams processed", summary.length]);
  sheet.addRow(["Total reports generated", successful]);
  sheet.addRow(["Failed reports", failed]);
  sheet.addRow([]);
  sheet.addRow([
    "Team",
    "Status",
    "File Name",
    "Total Employees",
    "Matched Employees",
    "Error Reason",
  ]);

  for (const item of summary) {
    sheet.addRow([
      item.teamName,
      item.status,
      item.fileName,
      item.totalEmployees,
      item.matchedEmployees,
      item.error ?? "",
    ]);
  }

  sheet.getColumn(1).width = 26;
  sheet.getColumn(2).width = 14;
  sheet.getColumn(3).width = 34;
  sheet.getColumn(4).width = 18;
  sheet.getColumn(5).width = 18;
  sheet.getColumn(6).width = 50;

  for (const rowNumber of [1, 2, 3, 5]) {
    sheet.getRow(rowNumber).font = { bold: true };
  }

  return workbook;
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
  for (let c = 1; c <= sheet.columnCount; c++) {
    headers.set(normalize(cellText(headerRow.getCell(c))), c);
  }

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
    const code = normalizeEmployeeCode(cellText(row.getCell(codeCol)));
    if (!code) continue;

    const name = cellText(row.getCell(nameCol)).trim();
    const meetingType = normalize(cellText(row.getCell(meetingTypeCol)));
    const eventType = normalize(cellText(row.getCell(eventTypeCol)));
    const shift = normalize(cellText(row.getCell(shiftCol)));
    const startTime = formatTime(
      row.getCell(startTimeCol).value,
      cellText(row.getCell(startTimeCol)),
    );

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

function findTemplateColumns(sheet: ExcelJS.Worksheet, required = true): DailyColumns | null {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const labels = new Map<string, number>();
    for (let c = 1; c <= sheet.columnCount; c++) {
      labels.set(normalize(cellText(row.getCell(c))), c);
    }

    const teamCol = findHeader(labels, TEMPLATE_HEADER_HINTS.team, false);
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
      return {
        headerRow: rowNumber,
        codeCol,
        nameCol,
        plannedCol,
        unplannedCol,
        morningCol,
        eveningCol,
        totalCol,
        cpCol,
        teamCol: teamCol || undefined,
      };
    }
  }

  if (!required) return null;
  throw new Error(
    "Template columns were not found. Required: Employee Code, Name, Planned, Unplanned, Mor, Eve, Total, Cp.",
  );
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

function cellText(cell: ExcelJS.Cell): string {
  try {
    return cell.text ?? "";
  } catch {
    const value = cell.value;
    if (value == null) return "";
    if (value instanceof Date) return value.toISOString();
    if (typeof value !== "object") return String(value);

    if ("text" in value && typeof value.text === "string") return value.text;
    if ("result" in value && value.result != null) return String(value.result);
    if ("hyperlink" in value && typeof value.hyperlink === "string") return value.hyperlink;
    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => String(part.text ?? "")).join("");
    }

    return "";
  }
}

function normalize(value: string): string {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function normalizeEmployeeCode(value: string): string {
  const digits =
    String(value ?? "")
      .match(/\d+/g)
      ?.join("") ?? "";
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

function safeFileName(value: string): string {
  return (
    value
      .trim()
      .replace(/[\\/:*?"<>|]+/g, "_")
      .replace(/\s+/g, "_")
      .replace(/_+/g, "_")
      .slice(0, 90) || "Team"
  );
}

function safeSheetName(value: string): string {
  return (
    value
      .trim()
      .replace(/[\\/?*[\]:]+/g, " ")
      .replace(/\s+/g, " ")
      .slice(0, 31) || "Team"
  );
}

function uniqueFileName(fileName: string, usedFileNames: Set<string>): string {
  if (!usedFileNames.has(fileName)) {
    usedFileNames.add(fileName);
    return fileName;
  }

  const extensionMatch = fileName.match(/(\.[^.]+)$/);
  const extension = extensionMatch?.[1] ?? "";
  const baseName = extension ? fileName.slice(0, -extension.length) : fileName;
  let counter = 2;
  let candidate = `${baseName}_${counter}${extension}`;

  while (usedFileNames.has(candidate)) {
    counter++;
    candidate = `${baseName}_${counter}${extension}`;
  }

  usedFileNames.add(candidate);
  return candidate;
}
