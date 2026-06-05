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
    selfieFiles: number;
    selfieRows: number;
    templateRows: number;
  };
  blob: Blob;
  fileName: string;
  sheetName: string;
  preview: { name: string; total: number }[];
}

export interface BulkDailyReportSummaryItem {
  teamName: string;
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
  dates: Set<string>;
  planned: number;
  unplanned: number;
  morning: number;
  evening: number;
  total: number;
  cpTime: string;
  selfies: number;
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
  selfieCol?: number;
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

interface DailySheetProcessResult {
  totalEmployees: number;
  matchedEmployees: number;
  unmatchedEmployees: string[];
  preview: { name: string; total: number }[];
}

interface SelfieEmployeeSummary {
  code: string;
  name: string;
  nameKey: string;
  dateCounts: Map<string, number>;
}

interface SelfieColumns {
  headerRow: number;
  messageTimeCol: number;
  messageTypeCol: number;
  messageBodyCol: number;
}

const CALL_HEADER_HINTS: Record<string, string[]> = {
  code: ["emp. id", "emp id", "employee id", "employee code"],
  name: ["employee name", "name"],
  date: ["date", "call date", "start date", "activity date"],
  startTime: ["start time"],
  eventType: ["event type"],
  meetingType: ["meeting type"],
  shift: ["shift"],
  team: ["team name", "team id", "team"],
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
  selfies: ["selfies", "selfie", "images", "image"],
};

const SELFIE_HEADER_HINTS: Record<string, string[]> = {
  messageTime: ["message time", "time", "date time", "date"],
  messageType: ["message type", "type"],
  messageBody: ["message body", "body", "message", "chat"],
};

export async function processDailyReport(
  callLogFile: File | File[],
  templateFile: File,
  selfieFile?: File | File[],
): Promise<DailyReportResult> {
  const callLog = await readCallLogs(callLogFile);
  await applySelfiesToCallLog(callLog, selfieFile);
  const templateWorkbook = new ExcelJS.Workbook();
  await templateWorkbook.xlsx.load(await templateFile.arrayBuffer());
  const sheet = selectDailyTemplateSheet(templateWorkbook, callLog);
  if (!sheet) throw new Error("Template workbook does not contain a sheet.");

  return await processDailyWorkbook(callLog, templateWorkbook, sheet, templateFile.name);
}

export async function processBulkDailyReports(
  callLogFile: File | File[],
  teamsWorkbookFile: File,
  selfieFileOrProgress?:
    | File
    | File[]
    | ((progress: { current: number; total: number; teamName: string }) => void),
  onProgress?: (progress: { current: number; total: number; teamName: string }) => void,
): Promise<BulkDailyReportResult> {
  const callLog = await readCallLogs(callLogFile);
  const selfieFile = typeof selfieFileOrProgress === "function" ? undefined : selfieFileOrProgress;
  const progressCallback =
    typeof selfieFileOrProgress === "function" ? selfieFileOrProgress : onProgress;
  await applySelfiesToCallLog(callLog, selfieFile);
  const teamsWorkbook = new ExcelJS.Workbook();
  await teamsWorkbook.xlsx.load(await teamsWorkbookFile.arrayBuffer());
  const allTeamSources = findDailyTeamSources(teamsWorkbook);
  const teamSources = filterTeamSourcesByCallLogTeams(allTeamSources, callLog.teamNames);

  if (!teamSources.length) {
    throw new Error(
      "No teams were found. Add a Team Name/Team ID column, or use one worksheet per team.",
    );
  }

  const summary: BulkDailyReportSummaryItem[] = [];

  for (let index = 0; index < teamSources.length; index++) {
    const source = teamSources[index];
    progressCallback?.({
      current: index + 1,
      total: teamSources.length,
      teamName: source.teamName,
    });

    try {
      const sheet = teamsWorkbook.getWorksheet(source.sheetName);
      if (!sheet) throw new Error(`Team sheet not found: ${source.sheetName}`);
      const shouldProcessRow =
        source.kind === "column"
          ? (row: ExcelJS.Row) =>
              teamKey(cellText(row.getCell(source.teamCol))) === teamKey(source.teamName)
          : undefined;
      const result = processDailySheet(callLog, sheet, shouldProcessRow);
      summary.push({
        teamName: source.teamName,
        status: "success",
        totalEmployees: result.totalEmployees,
        matchedEmployees: result.matchedEmployees,
      });
    } catch (error) {
      summary.push({
        teamName: source.teamName,
        status: "failed",
        totalEmployees: 0,
        matchedEmployees: 0,
        error: error instanceof Error ? error.message : "Unknown error",
      });
    }
  }

  const buffer = await teamsWorkbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const reportsGenerated = summary.filter((item) => item.status === "success").length;
  return {
    totalTeams: teamSources.length,
    reportsGenerated,
    failedReports: summary.length - reportsGenerated,
    blob,
    fileName: teamsWorkbookFile.name.replace(/\.xlsx$/i, "") + " - Daily Report.xlsx",
    summary,
  };
}

async function processDailyWorkbook(
  callLog: Awaited<ReturnType<typeof readCallLogs>>,
  templateWorkbook: ExcelJS.Workbook,
  sheet: ExcelJS.Worksheet,
  templateFileName: string,
  teamName?: string,
): Promise<DailyReportResult> {
  const result = processDailySheet(callLog, sheet);

  const buffer = await templateWorkbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  return {
    totalEmployees: result.totalEmployees,
    matchedEmployees: result.matchedEmployees,
    unmatchedEmployees: result.unmatchedEmployees,
    warnings: result.unmatchedEmployees.length
      ? [`${result.unmatchedEmployees.length} call log employee(s) were not found in the template.`]
      : [],
    debug: {
      callRows: callLog.callRows,
      faceToFaceRows: callLog.faceToFaceRows,
      contactPointRows: callLog.contactPointRows,
      selfieFiles: callLog.selfieFileCount,
      selfieRows: callLog.selfieRows,
      templateRows: result.totalEmployees,
    },
    blob,
    fileName: teamName
      ? `${safeFileName(teamName)}_Report.xlsx`
      : templateFileName.replace(/\.xlsx$/i, "") + " - Daily Report.xlsx",
    sheetName: sheet.name,
    preview: result.preview,
  };
}

function processDailySheet(
  callLog: Awaited<ReturnType<typeof readCallLogs>>,
  sheet: ExcelJS.Worksheet,
  shouldProcessRow: (row: ExcelJS.Row) => boolean = () => true,
): DailySheetProcessResult {
  const columns = findTemplateColumns(sheet);
  if (!columns) throw new Error("Template columns were not found.");
  const outputColumns = callLog.selfieFileCount > 0 ? ensureSelfieColumn(sheet, columns) : columns;

  const summariesByCode = new Map(callLog.summaries.map((item) => [item.code, item]));
  const unmatched = new Set(callLog.summaries.map((item) => `${item.code} ${item.name}`));
  const preview: { name: string; total: number }[] = [];
  let matchedEmployees = 0;
  let totalEmployees = 0;

  for (let rowNumber = columns.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    if (!shouldProcessRow(row)) continue;

    const code = normalizeEmployeeCode(cellText(row.getCell(outputColumns.codeCol)));
    const name = cellText(row.getCell(outputColumns.nameCol));
    if (!code) continue;
    totalEmployees++;

    const match = summariesByCode.get(code);
    if (!match) {
      clearDailyCells(row, outputColumns);
      continue;
    }

    matchedEmployees++;
    unmatched.delete(`${match.code} ${match.name}`);
    row.getCell(outputColumns.plannedCol).value = match.planned || 0;
    row.getCell(outputColumns.unplannedCol).value = match.unplanned || 0;
    row.getCell(outputColumns.morningCol).value = match.morning || 0;
    row.getCell(outputColumns.eveningCol).value = match.evening || 0;
    row.getCell(outputColumns.totalCol).value = match.total || 0;
    row.getCell(outputColumns.cpCol).value = match.cpTime || null;
    if (outputColumns.selfieCol) row.getCell(outputColumns.selfieCol).value = match.selfies || 0;

    styleDailyCells(row, outputColumns);
    if (match.total > 0) preview.push({ name: name || match.name, total: match.total });
  }

  return {
    totalEmployees,
    matchedEmployees,
    unmatchedEmployees: [...unmatched],
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

function filterTeamSourcesByCallLogTeams(
  teamSources: DailyTeamSource[],
  callLogTeamNames: string[],
): DailyTeamSource[] {
  if (!callLogTeamNames.length) return teamSources;

  const callLogTeamKeys = new Set(callLogTeamNames.map(teamKey).filter(Boolean));
  const matchingSources = teamSources.filter((source) =>
    callLogTeamKeys.has(teamKey(source.teamName)),
  );
  return matchingSources.length ? matchingSources : teamSources;
}

function selectDailyTemplateSheet(
  workbook: ExcelJS.Workbook,
  callLog: Awaited<ReturnType<typeof readCallLogs>>,
): ExcelJS.Worksheet | undefined {
  const templateSheets = workbook.worksheets.filter((sheet) => hasTemplateColumns(sheet));
  if (!templateSheets.length) return workbook.worksheets[0];

  const callLogTeamKeys = new Set(callLog.teamNames.map(teamKey).filter(Boolean));
  if (callLogTeamKeys.size) {
    const teamSheet = templateSheets.find((sheet) => callLogTeamKeys.has(teamKey(sheet.name)));
    if (teamSheet) return teamSheet;
  }

  const callLogCodes = new Set(callLog.summaries.map((summary) => summary.code));
  let bestSheet = templateSheets[0];
  let bestMatches = -1;

  for (const sheet of templateSheets) {
    const matches = countTemplateEmployeeCodeMatches(sheet, callLogCodes);
    if (matches > bestMatches) {
      bestMatches = matches;
      bestSheet = sheet;
    }
  }

  return bestSheet;
}

function countTemplateEmployeeCodeMatches(
  sheet: ExcelJS.Worksheet,
  callLogCodes: Set<string>,
): number {
  const columns = findTemplateColumns(sheet, false);
  if (!columns) return 0;

  let matches = 0;
  for (let rowNumber = columns.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber++) {
    const code = normalizeEmployeeCode(cellText(sheet.getRow(rowNumber).getCell(columns.codeCol)));
    if (code && callLogCodes.has(code)) matches++;
  }
  return matches;
}

function hasTemplateColumns(sheet: ExcelJS.Worksheet): boolean {
  return Boolean(findTemplateColumns(sheet, false));
}

async function readCallLog(file: File): Promise<{
  summaries: CallSummary[];
  callRows: number;
  faceToFaceRows: number;
  contactPointRows: number;
  teamNames: string[];
  selfieFileCount: number;
  selfieRows: number;
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
  const dateCol = findHeader(headers, CALL_HEADER_HINTS.date, false);
  const startTimeCol = findHeader(headers, CALL_HEADER_HINTS.startTime);
  const eventTypeCol = findHeader(headers, CALL_HEADER_HINTS.eventType);
  const meetingTypeCol = findHeader(headers, CALL_HEADER_HINTS.meetingType);
  const shiftCol = findHeader(headers, CALL_HEADER_HINTS.shift);
  const teamCol = findHeader(headers, CALL_HEADER_HINTS.team, false);

  const byCode = new Map<string, CallSummary>();
  const teamNames = new Map<string, string>();
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
    const teamName = teamCol ? cellText(row.getCell(teamCol)).trim() : "";
    const startTime = formatTime(
      row.getCell(startTimeCol).value,
      cellText(row.getCell(startTimeCol)),
    );

    if (teamName) teamNames.set(teamKey(teamName), teamName);

    let summary = byCode.get(code);
    if (!summary) {
      summary = {
        code,
        name,
        dates: new Set<string>(),
        planned: 0,
        unplanned: 0,
        morning: 0,
        evening: 0,
        total: 0,
        cpTime: "",
        selfies: 0,
      };
      byCode.set(code, summary);
    }

    callRows++;
    const dateKey = dateCol
      ? normalizeDateKey(row.getCell(dateCol).value, cellText(row.getCell(dateCol)))
      : "";
    if (dateKey) summary.dates.add(dateKey);
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

  return {
    summaries: [...byCode.values()],
    callRows,
    faceToFaceRows,
    contactPointRows,
    teamNames: [...teamNames.values()],
    selfieFileCount: 0,
    selfieRows: 0,
  };
}

async function readCallLogs(files: File | File[]): Promise<{
  summaries: CallSummary[];
  callRows: number;
  faceToFaceRows: number;
  contactPointRows: number;
  teamNames: string[];
  selfieFileCount: number;
  selfieRows: number;
}> {
  const fileList = Array.isArray(files) ? files : [files];
  if (!fileList.length) throw new Error("Please upload at least one call log Excel file.");

  const mergedByCode = new Map<string, CallSummary>();
  const teamNames = new Map<string, string>();
  let callRows = 0;
  let faceToFaceRows = 0;
  let contactPointRows = 0;

  for (const file of fileList) {
    const callLog = await readCallLog(file);
    callRows += callLog.callRows;
    faceToFaceRows += callLog.faceToFaceRows;
    contactPointRows += callLog.contactPointRows;

    for (const teamName of callLog.teamNames) {
      teamNames.set(teamKey(teamName), teamName);
    }

    for (const summary of callLog.summaries) {
      const existing = mergedByCode.get(summary.code);
      if (!existing) {
        mergedByCode.set(summary.code, {
          ...summary,
          dates: new Set(summary.dates),
        });
        continue;
      }

      summary.dates.forEach((date) => existing.dates.add(date));
      existing.planned += summary.planned;
      existing.unplanned += summary.unplanned;
      existing.morning += summary.morning;
      existing.evening += summary.evening;
      existing.total += summary.total;
      if (!existing.cpTime && summary.cpTime) existing.cpTime = summary.cpTime;
    }
  }

  return {
    summaries: [...mergedByCode.values()],
    callRows,
    faceToFaceRows,
    contactPointRows,
    teamNames: [...teamNames.values()],
    selfieFileCount: 0,
    selfieRows: 0,
  };
}

async function applySelfiesToCallLog(
  callLog: Awaited<ReturnType<typeof readCallLogs>>,
  selfieFiles?: File | File[],
) {
  const fileList = selfieFiles ? (Array.isArray(selfieFiles) ? selfieFiles : [selfieFiles]) : [];
  callLog.selfieFileCount = fileList.length;
  callLog.selfieRows = 0;
  for (const summary of callLog.summaries) summary.selfies = 0;
  if (!fileList.length) return;

  const selfieData = await readSelfieFiles(fileList);
  callLog.selfieRows = selfieData.imageRows;

  for (const summary of callLog.summaries) {
    const source =
      selfieData.byCode.get(summary.code) || selfieData.byName.get(personNameKey(summary.name));
    if (!source) continue;

    let selfies = 0;
    for (const date of summary.dates) {
      selfies += source.dateCounts.get(date) ?? 0;
    }
    summary.selfies = selfies;
  }
}

async function readSelfieFiles(files: File[]): Promise<{
  byCode: Map<string, SelfieEmployeeSummary>;
  byName: Map<string, SelfieEmployeeSummary>;
  imageRows: number;
}> {
  const byCode = new Map<string, SelfieEmployeeSummary>();
  const byName = new Map<string, SelfieEmployeeSummary>();
  let imageRows = 0;

  for (const file of files) {
    const employee = parseSelfieEmployeeFromFileName(file.name);
    if (!employee.code && !employee.nameKey) continue;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());

    for (const sheet of workbook.worksheets) {
      const columns = findSelfieColumns(sheet);
      if (!columns) continue;

      for (let rowNumber = columns.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber++) {
        const row = sheet.getRow(rowNumber);
        const body = cellText(row.getCell(columns.messageBodyCol));
        const type = columns.messageTypeCol ? cellText(row.getCell(columns.messageTypeCol)) : "";
        if (!isImageMessage(body, type)) continue;

        const dateKey = normalizeDateKey(
          row.getCell(columns.messageTimeCol).value,
          cellText(row.getCell(columns.messageTimeCol)),
        );
        if (!dateKey) continue;

        imageRows++;
        const summary = getSelfieSummary(byCode, byName, employee);
        summary.dateCounts.set(dateKey, (summary.dateCounts.get(dateKey) ?? 0) + 1);
      }
    }
  }

  return { byCode, byName, imageRows };
}

function getSelfieSummary(
  byCode: Map<string, SelfieEmployeeSummary>,
  byName: Map<string, SelfieEmployeeSummary>,
  employee: { code: string; name: string; nameKey: string },
): SelfieEmployeeSummary {
  const existing =
    (employee.code && byCode.get(employee.code)) ||
    (employee.nameKey && byName.get(employee.nameKey));
  if (existing) return existing;

  const summary: SelfieEmployeeSummary = {
    code: employee.code,
    name: employee.name,
    nameKey: employee.nameKey,
    dateCounts: new Map<string, number>(),
  };
  if (employee.code) byCode.set(employee.code, summary);
  if (employee.nameKey) byName.set(employee.nameKey, summary);
  return summary;
}

function findSelfieColumns(sheet: ExcelJS.Worksheet): SelfieColumns | null {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 20); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const headers = new Map<string, number>();
    for (let c = 1; c <= sheet.columnCount; c++) {
      headers.set(normalize(cellText(row.getCell(c))), c);
    }

    const messageTimeCol = findHeader(headers, SELFIE_HEADER_HINTS.messageTime, false);
    const messageTypeCol = findHeader(headers, SELFIE_HEADER_HINTS.messageType, false);
    const messageBodyCol = findHeader(headers, SELFIE_HEADER_HINTS.messageBody, false);

    if (messageTimeCol && messageBodyCol) {
      return {
        headerRow: rowNumber,
        messageTimeCol,
        messageTypeCol,
        messageBodyCol,
      };
    }
  }

  return null;
}

function parseSelfieEmployeeFromFileName(fileName: string): {
  code: string;
  name: string;
  nameKey: string;
} {
  const withoutExtension = fileName.replace(/\.[^.]+$/i, "");
  const bracketValue = withoutExtension.match(/\[\s*([^\]]+)\s*\]/)?.[1] ?? withoutExtension;
  const code = normalizeEmployeeCode(bracketValue);
  const name = bracketValue
    .replace(/\d+/g, " ")
    .replace(/[-_()[\]]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return {
    code,
    name,
    nameKey: personNameKey(name),
  };
}

function isImageMessage(body: string, type: string): boolean {
  const normalizedBody = normalize(body);
  const normalizedType = normalize(type);
  return normalizedBody.includes("image") || normalizedType === "image";
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
    const selfieCol = findHeader(labels, TEMPLATE_HEADER_HINTS.selfies, false);

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
        selfieCol: selfieCol || undefined,
        teamCol: teamCol || undefined,
      };
    }
  }

  if (!required) return null;
  throw new Error(
    "Template columns were not found. Required: Employee Code, Name, Planned, Unplanned, Mor, Eve, Total, Cp.",
  );
}

function ensureSelfieColumn(sheet: ExcelJS.Worksheet, columns: DailyColumns): DailyColumns {
  if (columns.selfieCol) return columns;

  const header = sheet.getRow(columns.headerRow);
  const col = Math.max(findLastUsedColumn(sheet), columns.cpCol, columns.totalCol) + 1;
  const headerCell = header.getCell(col);
  const sourceHeader = header.getCell(columns.cpCol);
  headerCell.value = "Selfies";
  headerCell.style = { ...sourceHeader.style };
  sheet.getColumn(col).width = Math.max(sheet.getColumn(col).width ?? 0, 12);
  header.commit();

  return { ...columns, selfieCol: col };
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
    columns.selfieCol,
  ].filter((col): col is number => Boolean(col))) {
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
    columns.selfieCol,
  ].filter((col): col is number => Boolean(col))) {
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

function teamKey(value: string): string {
  return normalize(value).replace(/[^a-z0-9]+/g, "");
}

function normalizeEmployeeCode(value: string): string {
  const digits =
    String(value ?? "")
      .match(/\d+/g)
      ?.join("") ?? "";
  return digits.length >= 3 ? digits : "";
}

function personNameKey(value: string): string {
  return normalize(value)
    .replace(/\([^)]*\)/g, " ")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\b(mr|mrs|ms|dr|mio|asm|hos|dbu|smio|sm)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeDateKey(value: unknown, text: string): string {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  const raw = String(text ?? "").trim();
  const iso = raw.match(/\b(\d{4})-(\d{1,2})-(\d{1,2})\b/);
  if (iso) return `${iso[1]}-${iso[2].padStart(2, "0")}-${iso[3].padStart(2, "0")}`;

  const slash = raw.match(/\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b/);
  if (slash) {
    const year = slash[3].length === 2 ? `20${slash[3]}` : slash[3];
    return `${year}-${slash[2].padStart(2, "0")}-${slash[1].padStart(2, "0")}`;
  }

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return "";
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
