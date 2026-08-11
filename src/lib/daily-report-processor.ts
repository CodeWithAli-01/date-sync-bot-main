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
    selfieFallbackMatches: number;
    selfieMatchedEmployees: number;
    selfieMissingSources: number;
    selfieAmbiguousNameMatches: number;
    selfieDateMisses: number;
    selfieEmptyFiles: number;
    selfieColumnMissFiles: number;
    selfieUnreadableFiles: number;
  };
  blob: Blob;
  fileName: string;
  sheetName: string;
  preview: { name: string; total: number }[];
  performanceRows: DailyPerformanceRow[];
}

export interface DailyPerformanceRow {
  teamName: string;
  sortOrder: number;
  employeeCode: string;
  name: string;
  region: string;
  city: string;
  designation: string;
  planned: number;
  unplanned: number;
  morningCalls: number;
  morningHours: number;
  morningMinutes: number;
  morningFirstCall: string;
  morningLastCall: string;
  eveningCalls: number;
  eveningHours: number;
  eveningMinutes: number;
  eveningFirstCall: string;
  eveningLastCall: string;
  totalWorkingHours: number;
  totalWorkingMinutes: number;
  totalCalls: number;
  selfies: number;
  cpTime: string;
  remarks: string;
  plannedPercent: number;
  topQualified: boolean;
  lowQualified: boolean;
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
  performanceRows: DailyPerformanceRow[];
  warnings: string[];
  debug: DailyReportResult["debug"];
}

interface CallSummary {
  code: string;
  name: string;
  teamName: string;
  dates: Set<string>;
  planned: number;
  unplanned: number;
  morning: number;
  evening: number;
  total: number;
  cpTime: string;
  cpMinutes: number | null;
  morningFirstMinutes: number | null;
  morningLastMinutes: number | null;
  eveningFirstMinutes: number | null;
  eveningLastMinutes: number | null;
  selfies: number | null;
  remarks: Set<string>;
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
  remarksCol?: number;
  selfieCol?: number;
  teamCol?: number;
  regionCol?: number;
  cityCol?: number;
  designationCol?: number;
}

interface CallLogColumns {
  headerRow: number;
  codeCol: number;
  nameCol: number;
  dateCol: number;
  startTimeCol: number;
  eventTypeCol: number;
  meetingTypeCol: number;
  shiftCol: number;
  teamCol?: number;
}

const SELFIE_WARNING_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFFFF00" },
};

const SELFIE_ZERO_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFFBF00" },
};

const SELFIE_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: "thin", color: { argb: "FF000000" } },
  left: { style: "thin", color: { argb: "FF000000" } },
  bottom: { style: "thin", color: { argb: "FF000000" } },
  right: { style: "thin", color: { argb: "FF000000" } },
};

const DAILY_REMARK_KEYWORDS = [
  "SICK LEAVE",
  "VISIT",
  "TOUR",
  "TRAVELLING",
  "MEETING",
  "TRAINING",
  "PRODUCT LAUNCH",
  "CASUAL LEAVE",
] as const;

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
  performanceRows: DailyPerformanceRow[];
}

interface SelfieEmployeeSummary {
  code: string;
  name: string;
  nameKey: string;
  fileName: string;
  dateCounts: Map<string, number>;
}

interface SelfieMatchResult {
  source?: SelfieEmployeeSummary;
  reason: "code" | "exact-name" | "similar-name" | "ambiguous-name" | "missing-source";
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
  meetingType: ["meeting type", "meetingtype"],
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
  remarks: ["remarks", "remark"],
  selfies: ["selfies", "selfie", "images", "image"],
  region: ["region", "zone", "area"],
  city: ["city", "town", "territory", "headquarter", "hq"],
  designation: ["designation", "desig", "position", "title"],
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
  removeRemarksColumnsFromWorkbook(templateWorkbook);
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
  removeRemarksColumnsFromWorkbook(teamsWorkbook);
  const allTeamSources = findDailyTeamSources(teamsWorkbook);
  const teamSources = filterTeamSourcesByCallLogTeams(allTeamSources, callLog.teamNames);

  if (!teamSources.length) {
    throw new Error(
      "No teams were found. Add a Team Name/Team ID column, or use one worksheet per team.",
    );
  }

  const summary: BulkDailyReportSummaryItem[] = [];
  const performanceRows: DailyPerformanceRow[] = [];

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
      performanceRows.push(
        ...result.performanceRows.map((row) => ({
          ...row,
          teamName: source.teamName,
          sortOrder: index * 100000 + row.sortOrder,
        })),
      );
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
    performanceRows,
    warnings: buildDailySelfieWarnings(callLog, []),
    debug: buildDailyDebug(
      callLog,
      summary.reduce((sum, item) => sum + item.totalEmployees, 0),
    ),
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
    warnings: buildDailySelfieWarnings(callLog, result.unmatchedEmployees),
    debug: buildDailyDebug(callLog, result.totalEmployees),
    blob,
    fileName: teamName
      ? `${safeFileName(teamName)}_Report.xlsx`
      : templateFileName.replace(/\.xlsx$/i, "") + " - Daily Report.xlsx",
    sheetName: sheet.name,
    preview: result.preview,
    performanceRows: result.performanceRows.map((row) => ({
      ...row,
      teamName: teamName || row.teamName || sheet.name,
    })),
  };
}

function buildDailyDebug(
  callLog: Awaited<ReturnType<typeof readCallLogs>>,
  templateRows: number,
): DailyReportResult["debug"] {
  return {
    callRows: callLog.callRows,
    faceToFaceRows: callLog.faceToFaceRows,
    contactPointRows: callLog.contactPointRows,
    selfieFiles: callLog.selfieFileCount,
    selfieRows: callLog.selfieRows,
    templateRows,
    selfieFallbackMatches: callLog.selfieFallbackMatches,
    selfieMatchedEmployees: callLog.selfieMatchedEmployees,
    selfieMissingSources: callLog.selfieMissingSources,
    selfieAmbiguousNameMatches: callLog.selfieAmbiguousNameMatches,
    selfieDateMisses: callLog.selfieDateMisses,
    selfieEmptyFiles: callLog.selfieEmptyFiles,
    selfieColumnMissFiles: callLog.selfieColumnMissFiles,
    selfieUnreadableFiles: callLog.selfieUnreadableFiles,
  };
}

function buildDailySelfieWarnings(
  callLog: Awaited<ReturnType<typeof readCallLogs>>,
  unmatchedEmployees: string[],
): string[] {
  return [
    ...(unmatchedEmployees.length
      ? [`${unmatchedEmployees.length} call log employee(s) were not found in the template.`]
      : []),
    ...(callLog.selfieFallbackMatches
      ? [`${callLog.selfieFallbackMatches} employee selfie count(s) used fallback matching.`]
      : []),
    ...(callLog.selfieEmptyFiles
      ? [`${callLog.selfieEmptyFiles} selfie workbook(s) were empty and could not be counted.`]
      : []),
    ...(callLog.selfieColumnMissFiles
      ? [
          `${callLog.selfieColumnMissFiles} selfie workbook(s) did not have readable Message Time / Message Body columns.`,
        ]
      : []),
    ...(callLog.selfieUnreadableFiles
      ? [`${callLog.selfieUnreadableFiles} selfie workbook(s) could not be opened.`]
      : []),
    ...(callLog.selfieMissingSources
      ? [
          `${callLog.selfieMissingSources} matched employee row(s) had no selfie workbook matched by employee code or unique name.`,
        ]
      : []),
    ...(callLog.selfieAmbiguousNameMatches
      ? [
          `${callLog.selfieAmbiguousNameMatches} employee selfie match(es) were skipped because the name matched more than one source file.`,
        ]
      : []),
    ...(callLog.selfieDateMisses
      ? [
          `${callLog.selfieDateMisses} employee selfie workbook(s) had images, but not on the exact call-log date. These were left as 0 to keep the daily report date-accurate.`,
        ]
      : []),
  ];
}

function processDailySheet(
  callLog: Awaited<ReturnType<typeof readCallLogs>>,
  sheet: ExcelJS.Worksheet,
  shouldProcessRow: (row: ExcelJS.Row) => boolean = () => true,
): DailySheetProcessResult {
  removeRemarksColumns(sheet);
  const columns = findTemplateColumns(sheet);
  if (!columns) throw new Error("Template columns were not found.");
  const withSelfieColumn =
    callLog.selfieFileCount > 0 ? ensureSelfieColumn(sheet, columns) : columns;
  const outputColumns = ensureRemarksColumn(sheet, withSelfieColumn);

  const summariesByTeamCode = new Map(
    callLog.summaries.map((item) => [callSummaryKey(item.teamName, item.code), item]),
  );
  const summariesByCode = new Map<string, CallSummary[]>();
  for (const summary of callLog.summaries) {
    const summaries = summariesByCode.get(summary.code) ?? [];
    summaries.push(summary);
    summariesByCode.set(summary.code, summaries);
  }
  const unmatched = new Set(callLog.summaries.map((item) => unmatchedSummaryKey(item)));
  const preview: { name: string; total: number }[] = [];
  const performanceRows: DailyPerformanceRow[] = [];
  let matchedEmployees = 0;
  let totalEmployees = 0;
  let sortOrder = 0;

  for (let rowNumber = columns.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    if (!shouldProcessRow(row)) continue;

    const code = normalizeEmployeeCode(cellText(row.getCell(outputColumns.codeCol)));
    const name = cellText(row.getCell(outputColumns.nameCol));
    if (!code) continue;
    totalEmployees++;
    sortOrder++;

    const rowTeamName = outputColumns.teamCol
      ? cellText(row.getCell(outputColumns.teamCol))
      : sheet.name;
    const match = findDailySummaryForRow(code, rowTeamName, summariesByTeamCode, summariesByCode);
    if (!match) {
      writeUnmatchedDailyDefaults(row, outputColumns);
      performanceRows.push({
        teamName: rowTeamName || sheet.name,
        sortOrder,
        employeeCode: code,
        name,
        region: outputColumns.regionCol ? cellText(row.getCell(outputColumns.regionCol)) : "",
        city: outputColumns.cityCol ? cellText(row.getCell(outputColumns.cityCol)) : "",
        designation: outputColumns.designationCol
          ? cellText(row.getCell(outputColumns.designationCol))
          : "",
        planned: 0,
        unplanned: 0,
        morningCalls: 0,
        morningHours: 0,
        morningMinutes: 0,
        morningFirstCall: "",
        morningLastCall: "",
        eveningCalls: 0,
        eveningHours: 0,
        eveningMinutes: 0,
        eveningFirstCall: "",
        eveningLastCall: "",
        totalWorkingHours: 0,
        totalWorkingMinutes: 0,
        totalCalls: 0,
        selfies: outputColumns.selfieCol
          ? parseSelfieCountFromCell(cellText(row.getCell(outputColumns.selfieCol)))
          : 0,
        cpTime: "",
        remarks: "",
        plannedPercent: 0,
        topQualified: false,
        lowQualified: true,
      });
      continue;
    }

    matchedEmployees++;
    unmatched.delete(unmatchedSummaryKey(match));
    row.getCell(outputColumns.plannedCol).value = match.planned || 0;
    row.getCell(outputColumns.unplannedCol).value = match.unplanned || 0;
    row.getCell(outputColumns.morningCol).value = match.morning || 0;
    row.getCell(outputColumns.eveningCol).value = match.evening || 0;
    row.getCell(outputColumns.totalCol).value = match.total || 0;
    row.getCell(outputColumns.cpCol).value = match.cpTime || null;
    if (outputColumns.remarksCol) {
      const remarksCell = row.getCell(outputColumns.remarksCol);
      remarksCell.value = formatRemarks(match.remarks);
      styleRemarksCell(remarksCell);
    }
    styleDailyCells(row, outputColumns);
    if (outputColumns.selfieCol) {
      const selfieCell = row.getCell(outputColumns.selfieCol);
      if (match.selfies === null) {
        selfieCell.value = null;
        applyIsolatedSelfieStyle(selfieCell, undefined);
      } else {
        selfieCell.value = formatSelfieText(match.selfies);
        styleSelfieCell(selfieCell, match.selfies);
      }
    }
    const performanceSelfies =
      match.selfies ??
      (outputColumns.selfieCol
        ? parseSelfieCountFromCell(cellText(row.getCell(outputColumns.selfieCol)))
        : 0);
    const morningMinutes = dailyShiftMinutes(
      match.cpMinutes,
      match.morningFirstMinutes,
      match.morningLastMinutes,
      match.morning,
    );
    const eveningMinutes = dailyShiftMinutes(
      null,
      match.eveningFirstMinutes,
      match.eveningLastMinutes,
      match.evening,
    );
    const morningHours = roundHours(morningMinutes / 60);
    const eveningHours = roundHours(eveningMinutes / 60);
    const totalWorkingMinutes = morningMinutes + eveningMinutes;
    const totalWorkingHours = roundHours(totalWorkingMinutes / 60);
    if (match.total > 0) preview.push({ name: name || match.name, total: match.total });
    performanceRows.push({
      teamName: rowTeamName || match.teamName || sheet.name,
      sortOrder,
      employeeCode: code,
      name: name || match.name,
      region: outputColumns.regionCol ? cellText(row.getCell(outputColumns.regionCol)) : "",
      city: outputColumns.cityCol ? cellText(row.getCell(outputColumns.cityCol)) : "",
      designation: outputColumns.designationCol
        ? cellText(row.getCell(outputColumns.designationCol))
        : "",
      planned: match.planned,
      unplanned: match.unplanned,
      morningCalls: match.morning,
      morningHours,
      morningMinutes,
      morningFirstCall: formatMinutes(match.morningFirstMinutes),
      morningLastCall: formatMinutes(match.morningLastMinutes),
      eveningCalls: match.evening,
      eveningHours,
      eveningMinutes,
      eveningFirstCall: formatMinutes(match.eveningFirstMinutes),
      eveningLastCall: formatMinutes(match.eveningLastMinutes),
      totalWorkingHours,
      totalWorkingMinutes,
      totalCalls: match.total,
      selfies: performanceSelfies,
      cpTime: match.cpTime,
      remarks: formatRemarks(match.remarks),
      plannedPercent: percent(match.planned, match.total),
      topQualified: match.total >= 12 && morningHours >= 4 && eveningHours >= 3,
      lowQualified: match.total <= 5 || match.morning === 0 || match.evening === 0,
    });
  }

  applySelfieColumnStyles(sheet, outputColumns);
  applyRemarksColumnStyles(sheet, outputColumns);

  return {
    totalEmployees,
    matchedEmployees,
    unmatchedEmployees: [...unmatched],
    preview: preview.sort((a, b) => b.total - a.total).slice(0, 10),
    performanceRows,
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

function callSummaryKey(teamName: string, code: string): string {
  const team = teamKey(teamName);
  return team ? `${team}|${code}` : code;
}

function unmatchedSummaryKey(summary: CallSummary): string {
  return `${callSummaryKey(summary.teamName, summary.code)}|${summary.name}`;
}

function findDailySummaryForRow(
  code: string,
  teamName: string,
  summariesByTeamCode: Map<string, CallSummary>,
  summariesByCode: Map<string, CallSummary[]>,
): CallSummary | undefined {
  const teamKeyValue = teamKey(teamName);
  if (teamKeyValue) {
    const direct = summariesByTeamCode.get(callSummaryKey(teamName, code));
    if (direct) return direct;

    const teamMatch = (summariesByCode.get(code) ?? []).find((summary) =>
      teamNamesMatch(summary.teamName, teamName),
    );
    if (teamMatch) return teamMatch;
  }

  const codeMatches = summariesByCode.get(code) ?? [];
  return codeMatches.length === 1 ? codeMatches[0] : undefined;
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
  selfieFallbackMatches: number;
  selfieMatchedEmployees: number;
  selfieMissingSources: number;
  selfieAmbiguousNameMatches: number;
  selfieDateMisses: number;
  selfieEmptyFiles: number;
  selfieColumnMissFiles: number;
  selfieUnreadableFiles: number;
}> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  if (!workbook.worksheets.length) throw new Error("Call log workbook does not contain a sheet.");

  const byKey = new Map<string, CallSummary>();
  const teamNames = new Map<string, string>();
  let callRows = 0;
  let faceToFaceRows = 0;
  let contactPointRows = 0;
  let readableSheets = 0;

  for (const sheet of workbook.worksheets) {
    const columns = findCallLogColumns(sheet);
    if (!columns) continue;
    readableSheets++;

    const result = readCallLogSheet(sheet, columns, byKey, teamNames);
    callRows += result.callRows;
    faceToFaceRows += result.faceToFaceRows;
    contactPointRows += result.contactPointRows;
  }

  if (!readableSheets) {
    throw new Error("Call log workbook does not contain readable call-log columns.");
  }

  return {
    summaries: [...byKey.values()],
    callRows,
    faceToFaceRows,
    contactPointRows,
    teamNames: [...teamNames.values()],
    selfieFileCount: 0,
    selfieRows: 0,
    selfieFallbackMatches: 0,
    selfieMatchedEmployees: 0,
    selfieMissingSources: 0,
    selfieAmbiguousNameMatches: 0,
    selfieDateMisses: 0,
    selfieEmptyFiles: 0,
    selfieColumnMissFiles: 0,
    selfieUnreadableFiles: 0,
  };
}

function readCallLogSheet(
  sheet: ExcelJS.Worksheet,
  columns: CallLogColumns,
  byKey: Map<string, CallSummary>,
  teamNames: Map<string, string>,
): { callRows: number; faceToFaceRows: number; contactPointRows: number } {
  let callRows = 0;
  let faceToFaceRows = 0;
  let contactPointRows = 0;

  for (let rowNumber = columns.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const code = normalizeEmployeeCode(cellText(row.getCell(columns.codeCol)));
    if (!code) continue;

    const name = cellText(row.getCell(columns.nameCol)).trim();
    const meetingType = normalize(cellText(row.getCell(columns.meetingTypeCol)));
    const eventType = normalize(cellText(row.getCell(columns.eventTypeCol)));
    const shift = normalize(cellText(row.getCell(columns.shiftCol)));
    const teamName = columns.teamCol ? cellText(row.getCell(columns.teamCol)).trim() : sheet.name;
    const startTimeInfo = parseTimeInfo(
      row.getCell(columns.startTimeCol).value,
      cellText(row.getCell(columns.startTimeCol)),
    );
    const startTime = startTimeInfo.label;

    if (teamName) teamNames.set(teamKey(teamName), teamName);

    const summaryKey = callSummaryKey(teamName, code);
    let summary = byKey.get(summaryKey);
    if (!summary) {
      summary = {
        code,
        name,
        teamName,
        dates: new Set<string>(),
        planned: 0,
        unplanned: 0,
        morning: 0,
        evening: 0,
        total: 0,
        cpTime: "",
        cpMinutes: null,
        morningFirstMinutes: null,
        morningLastMinutes: null,
        eveningFirstMinutes: null,
        eveningLastMinutes: null,
        selfies: null,
        remarks: new Set<string>(),
      };
      byKey.set(summaryKey, summary);
    }

    callRows++;
    extractDailyRemarks(cellText(row.getCell(columns.meetingTypeCol))).forEach((remark) =>
      summary.remarks.add(remark),
    );
    const dateKey = columns.dateCol
      ? normalizeDateKey(row.getCell(columns.dateCol).value, cellText(row.getCell(columns.dateCol)))
      : "";
    if (dateKey) summary.dates.add(dateKey);
    if (isContactPointMeeting(meetingType)) {
      contactPointRows++;
      if (startTimeInfo.minutes !== null) {
        summary.cpMinutes =
          summary.cpMinutes === null
            ? startTimeInfo.minutes
            : Math.min(summary.cpMinutes, startTimeInfo.minutes);
        summary.cpTime = formatMinutes(summary.cpMinutes);
      } else if (!summary.cpTime && startTime) {
        summary.cpTime = startTime;
      }
      continue;
    }

    if (meetingType && meetingType !== "face to face call") continue;
    faceToFaceRows++;
    if (eventType === "planned") summary.planned++;
    if (eventType === "unplanned") summary.unplanned++;
    if (shift === "morning") {
      summary.morning++;
      updateShiftWindow(summary, "morning", startTimeInfo.minutes);
    }
    if (shift === "evening") {
      summary.evening++;
      updateShiftWindow(summary, "evening", startTimeInfo.minutes);
    }
    summary.total = summary.morning + summary.evening;
  }

  return {
    callRows,
    faceToFaceRows,
    contactPointRows,
  };
}

function findCallLogColumns(sheet: ExcelJS.Worksheet): CallLogColumns | null {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const headers = new Map<string, number>();
    for (let c = 1; c <= sheet.columnCount; c++) {
      headers.set(normalize(cellText(row.getCell(c))), c);
    }

    const codeCol = findHeader(headers, CALL_HEADER_HINTS.code, false);
    const nameCol = findHeader(headers, CALL_HEADER_HINTS.name, false);
    const dateCol = findHeader(headers, CALL_HEADER_HINTS.date, false);
    const startTimeCol = findHeader(headers, CALL_HEADER_HINTS.startTime, false);
    const eventTypeCol = findHeader(headers, CALL_HEADER_HINTS.eventType, false);
    const meetingTypeCol = findHeader(headers, CALL_HEADER_HINTS.meetingType, false);
    const shiftCol = findHeader(headers, CALL_HEADER_HINTS.shift, false);
    const teamCol = findHeader(headers, CALL_HEADER_HINTS.team, false);

    if (codeCol && nameCol && startTimeCol && eventTypeCol && meetingTypeCol && shiftCol) {
      return {
        headerRow: rowNumber,
        codeCol,
        nameCol,
        dateCol,
        startTimeCol,
        eventTypeCol,
        meetingTypeCol,
        shiftCol,
        teamCol: teamCol || undefined,
      };
    }
  }

  return null;
}

async function readCallLogs(files: File | File[]): Promise<{
  summaries: CallSummary[];
  callRows: number;
  faceToFaceRows: number;
  contactPointRows: number;
  teamNames: string[];
  selfieFileCount: number;
  selfieRows: number;
  selfieFallbackMatches: number;
  selfieMatchedEmployees: number;
  selfieMissingSources: number;
  selfieAmbiguousNameMatches: number;
  selfieDateMisses: number;
  selfieEmptyFiles: number;
  selfieColumnMissFiles: number;
  selfieUnreadableFiles: number;
}> {
  const fileList = Array.isArray(files) ? files : [files];
  if (!fileList.length) throw new Error("Please upload at least one call log Excel file.");

  const mergedByKey = new Map<string, CallSummary>();
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
      const key = callSummaryKey(summary.teamName, summary.code);
      const existing = mergedByKey.get(key);
      if (!existing) {
        mergedByKey.set(key, {
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
      summary.remarks.forEach((remark) => existing.remarks.add(remark));
      existing.cpMinutes = minNullable(existing.cpMinutes, summary.cpMinutes);
      existing.cpTime = formatMinutes(existing.cpMinutes) || existing.cpTime || summary.cpTime;
      existing.morningFirstMinutes = minNullable(
        existing.morningFirstMinutes,
        summary.morningFirstMinutes,
      );
      existing.morningLastMinutes = maxNullable(
        existing.morningLastMinutes,
        summary.morningLastMinutes,
      );
      existing.eveningFirstMinutes = minNullable(
        existing.eveningFirstMinutes,
        summary.eveningFirstMinutes,
      );
      existing.eveningLastMinutes = maxNullable(
        existing.eveningLastMinutes,
        summary.eveningLastMinutes,
      );
    }
  }

  return {
    summaries: [...mergedByKey.values()],
    callRows,
    faceToFaceRows,
    contactPointRows,
    teamNames: [...teamNames.values()],
    selfieFileCount: 0,
    selfieRows: 0,
    selfieFallbackMatches: 0,
    selfieMatchedEmployees: 0,
    selfieMissingSources: 0,
    selfieAmbiguousNameMatches: 0,
    selfieDateMisses: 0,
    selfieEmptyFiles: 0,
    selfieColumnMissFiles: 0,
    selfieUnreadableFiles: 0,
  };
}

async function applySelfiesToCallLog(
  callLog: Awaited<ReturnType<typeof readCallLogs>>,
  selfieFiles?: File | File[],
) {
  const fileList = selfieFiles ? (Array.isArray(selfieFiles) ? selfieFiles : [selfieFiles]) : [];
  callLog.selfieFileCount = fileList.length;
  callLog.selfieRows = 0;
  callLog.selfieFallbackMatches = 0;
  callLog.selfieMatchedEmployees = 0;
  callLog.selfieMissingSources = 0;
  callLog.selfieAmbiguousNameMatches = 0;
  callLog.selfieDateMisses = 0;
  callLog.selfieEmptyFiles = 0;
  callLog.selfieColumnMissFiles = 0;
  callLog.selfieUnreadableFiles = 0;
  for (const summary of callLog.summaries) summary.selfies = fileList.length ? null : 0;
  if (!fileList.length) return;

  const selfieData = await readSelfieFiles(fileList);
  callLog.selfieRows = selfieData.imageRows;
  callLog.selfieEmptyFiles = selfieData.emptyFiles.length;
  callLog.selfieColumnMissFiles = selfieData.columnMissFiles.length;
  callLog.selfieUnreadableFiles = selfieData.unreadableFiles.length;

  for (const summary of callLog.summaries) {
    const match = findSelfieSourceForSummary(summary, selfieData);
    if (!match.source) {
      if (match.reason === "ambiguous-name") callLog.selfieAmbiguousNameMatches++;
      else callLog.selfieMissingSources++;
      continue;
    }

    callLog.selfieMatchedEmployees++;
    const resolved = resolveSelfieCountForDates(summary.dates, match.source.dateCounts);
    summary.selfies = resolved.count;
    if (resolved.usedFallback) callLog.selfieFallbackMatches++;
    if (resolved.dateMissed) callLog.selfieDateMisses++;
  }
}

async function readSelfieFiles(files: File[]): Promise<{
  byCode: Map<string, SelfieEmployeeSummary>;
  byName: Map<string, SelfieEmployeeSummary[]>;
  sources: SelfieEmployeeSummary[];
  imageRows: number;
  emptyFiles: string[];
  columnMissFiles: string[];
  unreadableFiles: string[];
}> {
  const byCode = new Map<string, SelfieEmployeeSummary>();
  const byName = new Map<string, SelfieEmployeeSummary[]>();
  const sources: SelfieEmployeeSummary[] = [];
  const emptyFiles: string[] = [];
  const columnMissFiles: string[] = [];
  const unreadableFiles: string[] = [];
  let imageRows = 0;

  for (const file of files) {
    const employee = parseSelfieEmployeeFromFileName(file.name);
    if (!employee.code && !employee.nameKey) continue;

    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.load(await file.arrayBuffer());
    } catch {
      unreadableFiles.push(file.name);
      continue;
    }

    if (workbook.worksheets.every((sheet) => sheet.rowCount === 0 && sheet.columnCount === 0)) {
      emptyFiles.push(file.name);
      continue;
    }

    let foundColumns = false;

    for (const sheet of workbook.worksheets) {
      const columns = findSelfieColumns(sheet);
      if (!columns) continue;
      foundColumns = true;

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
        const summary = getSelfieSummary(byCode, byName, sources, employee, file.name);
        summary.dateCounts.set(dateKey, (summary.dateCounts.get(dateKey) ?? 0) + 1);
      }
    }

    if (!foundColumns) columnMissFiles.push(file.name);
  }

  return {
    byCode,
    byName,
    sources,
    imageRows,
    emptyFiles,
    columnMissFiles,
    unreadableFiles,
  };
}

function findSelfieSourceForSummary(
  summary: CallSummary,
  selfieData: {
    byCode: Map<string, SelfieEmployeeSummary>;
    byName: Map<string, SelfieEmployeeSummary[]>;
    sources: SelfieEmployeeSummary[];
  },
): SelfieMatchResult {
  const byCode = selfieData.byCode.get(summary.code);
  if (byCode) return { source: byCode, reason: "code" };

  const summaryNameKey = personNameKey(summary.name);
  const byName = selfieData.byName.get(summaryNameKey) ?? [];
  if (byName.length === 1) return { source: byName[0], reason: "exact-name" };
  if (byName.length > 1) return { reason: "ambiguous-name" };

  const candidates = uniqueSelfieSources(
    selfieData.sources.filter((source) => selfieNamesMatch(summaryNameKey, source.nameKey)),
  );
  if (candidates.length === 1) return { source: candidates[0], reason: "similar-name" };
  if (candidates.length > 1) return { reason: "ambiguous-name" };

  return { reason: "missing-source" };
}

function resolveSelfieCountForDates(
  callDates: Set<string>,
  selfieDateCounts: Map<string, number>,
): { count: number; usedFallback: boolean; dateMissed: boolean } {
  const exactCount = [...callDates].reduce(
    (sum, date) => sum + (selfieDateCounts.get(date) ?? 0),
    0,
  );
  if (exactCount > 0) return { count: exactCount, usedFallback: false, dateMissed: false };

  const sourceHasImages = [...selfieDateCounts.values()].some((count) => count > 0);
  return {
    count: 0,
    usedFallback: false,
    dateMissed: callDates.size > 0 && sourceHasImages,
  };
}

function getSelfieSummary(
  byCode: Map<string, SelfieEmployeeSummary>,
  byName: Map<string, SelfieEmployeeSummary[]>,
  sources: SelfieEmployeeSummary[],
  employee: { code: string; name: string; nameKey: string },
  fileName: string,
): SelfieEmployeeSummary {
  const existing = employee.code ? byCode.get(employee.code) : undefined;
  if (existing) return existing;

  const summary: SelfieEmployeeSummary = {
    code: employee.code,
    name: employee.name,
    nameKey: employee.nameKey,
    fileName,
    dateCounts: new Map<string, number>(),
  };
  if (employee.code) byCode.set(employee.code, summary);
  if (employee.nameKey) {
    const summaries = byName.get(employee.nameKey) ?? [];
    summaries.push(summary);
    byName.set(employee.nameKey, summaries);
  }
  sources.push(summary);
  return summary;
}

function uniqueSelfieSources(sources: SelfieEmployeeSummary[]): SelfieEmployeeSummary[] {
  return [...new Set(sources)];
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
  const code = extractEmployeeCodeFromText(bracketValue);
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

function extractEmployeeCodeFromText(value: string): string {
  const tokens = String(value ?? "").match(/\d{3,8}/g) ?? [];
  return tokens.find((token) => token.length >= 3 && token.length <= 6) ?? "";
}

function isImageMessage(body: string, type: string): boolean {
  const normalizedBody = normalize(body);
  const normalizedType = normalize(type);
  return normalizedBody.includes("image") || normalizedType === "image";
}

function extractDailyRemarks(meetingTypeText: string): string[] {
  const normalizedMeetingType = normalizeRemarkText(meetingTypeText);
  if (!normalizedMeetingType) return [];

  const keywordRemarks = DAILY_REMARK_KEYWORDS.filter((keyword) =>
    new RegExp(`(^|\\s)${escapeRegExp(normalizeRemarkText(keyword))}(?=\\s|$)`).test(
      normalizedMeetingType,
    ),
  );
  if (keywordRemarks.length) return keywordRemarks;

  if (
    normalizedMeetingType === "face to face call" ||
    isContactPointMeeting(normalizedMeetingType)
  ) {
    return [];
  }

  return [toTitleCase(normalizedMeetingType)];
}

function formatRemarks(remarks: Set<string>): string {
  const knownRemarks = new Set<string>(DAILY_REMARK_KEYWORDS);
  return [
    ...DAILY_REMARK_KEYWORDS.filter((keyword) => remarks.has(keyword)),
    ...[...remarks].filter((remark) => !knownRemarks.has(remark)),
  ].join(", ");
}

function normalizeRemarkText(value: string): string {
  return normalize(value)
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function toTitleCase(value: string): string {
  return value.replace(/\b[a-z]/g, (letter) => letter.toUpperCase());
}

function percent(part: number, total: number): number {
  return total > 0 ? Math.round((part / total) * 100) : 0;
}

function isTimeOnOrBefore(value: string, maxMinutes: number): boolean {
  const minutes = parseTimeToMinutes(value);
  return minutes !== null && minutes <= maxMinutes;
}

function dailyShiftMinutes(
  preferredStartMinutes: number | null,
  firstCallMinutes: number | null,
  lastCallMinutes: number | null,
  callCount: number,
): number {
  const end = lastCallMinutes;
  if (end === null) return 0;

  const start = preferredStartMinutes ?? firstCallMinutes;
  if (start === null) return 0;
  if (end <= start) return callCount > 0 ? 1 : 0;
  return end - start;
}

function isContactPointMeeting(meetingType: string): boolean {
  return (
    meetingType === "contact point" ||
    meetingType === "cp" ||
    meetingType.includes("contact point") ||
    meetingType.includes("cp punch")
  );
}

function updateShiftWindow(
  summary: CallSummary,
  shift: "morning" | "evening",
  minutes: number | null,
) {
  if (minutes === null) return;

  if (shift === "morning") {
    summary.morningFirstMinutes = minNullable(summary.morningFirstMinutes, minutes);
    summary.morningLastMinutes = maxNullable(summary.morningLastMinutes, minutes);
    return;
  }

  summary.eveningFirstMinutes = minNullable(summary.eveningFirstMinutes, minutes);
  summary.eveningLastMinutes = maxNullable(summary.eveningLastMinutes, minutes);
}

function minNullable(a: number | null, b: number | null): number | null {
  if (a === null) return b;
  if (b === null) return a;
  return Math.min(a, b);
}

function maxNullable(a: number | null, b: number | null): number | null {
  if (a === null) return b;
  if (b === null) return a;
  return Math.max(a, b);
}

function roundHours(value: number): number {
  return Math.round(value * 100) / 100;
}

function parseTimeInfo(value: unknown, text: string): { label: string; minutes: number | null } {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const minutes = value.getHours() * 60 + value.getMinutes();
    return { label: formatMinutes(minutes), minutes };
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const dayFraction = value >= 0 && value < 1 ? value : value % 1;
    if (dayFraction >= 0) {
      const minutes = Math.round(dayFraction * 24 * 60);
      return { label: formatMinutes(minutes), minutes };
    }
  }

  const minutes = parseTimeToMinutes(text);
  return { label: minutes === null ? "" : formatMinutes(minutes), minutes };
}

function formatMinutes(value: number | null): string {
  if (value === null) return "";
  const normalized = ((value % (24 * 60)) + 24 * 60) % (24 * 60);
  let hours = Math.floor(normalized / 60);
  const minutes = normalized % 60;
  const suffix = hours >= 12 ? "PM" : "AM";
  hours = hours % 12 || 12;
  return `${hours}:${String(minutes).padStart(2, "0")} ${suffix}`;
}

function parseTimeToMinutes(value: string): number | null {
  const text = String(value ?? "").trim();
  if (!text) return null;
  const match = text.match(/\b(\d{1,2})(?::(\d{2}))?\s*(am|pm)?\b/i);
  if (!match) return null;
  let hours = Number(match[1]);
  const minutes = Number(match[2] ?? 0);
  const meridiem = match[3]?.toLowerCase();
  if (!Number.isFinite(hours) || !Number.isFinite(minutes)) return null;
  if (meridiem === "pm" && hours < 12) hours += 12;
  if (meridiem === "am" && hours === 12) hours = 0;
  return hours * 60 + minutes;
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
    const remarksCol = findHeader(labels, TEMPLATE_HEADER_HINTS.remarks, false);
    const selfieCol = findHeader(labels, TEMPLATE_HEADER_HINTS.selfies, false);
    const regionCol = findHeader(labels, TEMPLATE_HEADER_HINTS.region, false);
    const cityCol = findHeader(labels, TEMPLATE_HEADER_HINTS.city, false);
    const designationCol = findHeader(labels, TEMPLATE_HEADER_HINTS.designation, false);

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
        remarksCol: remarksCol || undefined,
        selfieCol: selfieCol || undefined,
        teamCol: teamCol || undefined,
        regionCol: regionCol || undefined,
        cityCol: cityCol || undefined,
        designationCol: designationCol || undefined,
      };
    }
  }

  if (!required) return null;
  throw new Error(
    "Template columns were not found. Required: Employee Code, Name, Planned, Unplanned, Mor, Eve, Total, Cp.",
  );
}

function removeRemarksColumns(sheet: ExcelJS.Worksheet) {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 12); rowNumber++) {
    const row = sheet.getRow(rowNumber);
    const remarksColumns: number[] = [];

    for (let c = 1; c <= sheet.columnCount; c++) {
      const label = normalize(cellText(row.getCell(c)));
      if (label === "remarks" || label === "remark") remarksColumns.push(c);
    }

    if (!remarksColumns.length) continue;
    const titleValue = findSheetTitleValue(sheet, rowNumber);
    for (const col of remarksColumns.sort((a, b) => b - a)) {
      sheet.spliceColumns(col, 1);
    }
    restoreSheetTitleValue(sheet, rowNumber, titleValue);
    return;
  }
}

function removeRemarksColumnsFromWorkbook(workbook: ExcelJS.Workbook) {
  for (const sheet of workbook.worksheets) removeRemarksColumns(sheet);
}

function findSheetTitleValue(sheet: ExcelJS.Worksheet, headerRowNumber: number): string {
  for (let rowNumber = 1; rowNumber < headerRowNumber; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    for (let colNumber = 1; colNumber <= sheet.columnCount; colNumber++) {
      const value = cellText(row.getCell(colNumber)).trim();
      if (value) return value;
    }
  }
  return sheet.name;
}

function restoreSheetTitleValue(
  sheet: ExcelJS.Worksheet,
  headerRowNumber: number,
  titleValue: string,
) {
  if (!titleValue) return;
  for (let rowNumber = 1; rowNumber < headerRowNumber; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    for (let colNumber = 1; colNumber <= sheet.columnCount; colNumber++) {
      const cell = row.getCell(colNumber);
      const isMergedChild = cell.isMerged && cell.master?.address !== cell.address;
      if (isMergedChild) continue;

      const currentValue = cellText(cell).trim();
      const looksLikeTitleCell =
        cell.isMerged || colNumber === 1 || normalize(currentValue) === normalize(sheet.name);
      if (looksLikeTitleCell && !currentValue) {
        cell.value = titleValue;
        row.commit();
        return;
      }
    }
  }
}

function repairSheetTitleAfterColumnInsert(
  sheet: ExcelJS.Worksheet,
  headerRowNumber: number,
  titleValue: string,
) {
  const normalizedTitle = normalize(titleValue);
  if (!normalizedTitle) return;

  for (let rowNumber = 1; rowNumber < headerRowNumber; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    let rowHasTitle = false;
    for (let colNumber = 1; colNumber <= sheet.columnCount; colNumber++) {
      if (normalize(cellText(row.getCell(colNumber)).trim()) === normalizedTitle) {
        rowHasTitle = true;
        break;
      }
    }
    if (!rowHasTitle) continue;

    const targetCell = firstWritableTitleCell(row, sheet.columnCount);
    targetCell.value = titleValue;

    for (let colNumber = 1; colNumber <= sheet.columnCount; colNumber++) {
      const cell = row.getCell(colNumber);
      const writableCell = cell.isMerged ? cell.master : cell;
      if (writableCell.address === targetCell.address) continue;
      if (normalize(cellText(cell).trim()) === normalizedTitle) writableCell.value = null;
    }

    row.commit();
    return;
  }
}

function firstWritableTitleCell(row: ExcelJS.Row, columnCount: number): ExcelJS.Cell {
  for (let colNumber = 1; colNumber <= columnCount; colNumber++) {
    const cell = row.getCell(colNumber);
    if (cell.isMerged && cell.master?.address !== cell.address) continue;
    if (cell.isMerged || colNumber === 1) return cell.isMerged ? cell.master : cell;
  }
  return row.getCell(1);
}

function ensureRemarksColumn(sheet: ExcelJS.Worksheet, columns: DailyColumns): DailyColumns {
  if (columns.remarksCol) {
    sheet.getColumn(columns.remarksCol).width = Math.max(
      sheet.getColumn(columns.remarksCol).width ?? 0,
      20,
    );
    sheet.getRow(columns.headerRow).getCell(columns.remarksCol).border = {
      ...(sheet.getRow(columns.headerRow).getCell(columns.remarksCol).border ?? {}),
      ...SELFIE_BORDER,
    };
    return columns;
  }

  const header = sheet.getRow(columns.headerRow);
  const insertAt = columns.selfieCol ? columns.selfieCol + 1 : columns.cpCol + 1;
  const titleValue = findSheetTitleValue(sheet, columns.headerRow);
  sheet.spliceColumns(insertAt, 0, []);
  const headerCell = header.getCell(insertAt);
  const sourceHeader = header.getCell(columns.selfieCol ?? columns.cpCol);
  headerCell.value = "Remarks";
  headerCell.style = cloneStyle(sourceHeader.style);
  headerCell.border = { ...(headerCell.border ?? {}), ...SELFIE_BORDER };
  sheet.getColumn(insertAt).width = Math.max(sheet.getColumn(insertAt).width ?? 0, 20);
  header.commit();
  repairSheetTitleAfterColumnInsert(sheet, columns.headerRow, titleValue);

  return shiftDailyColumnsAfterInsert({ ...columns, remarksCol: insertAt }, insertAt);
}

function shiftDailyColumnsAfterInsert(columns: DailyColumns, insertAt: number): DailyColumns {
  const shift = (col: number | undefined) => (col && col >= insertAt ? col + 1 : col);
  return {
    ...columns,
    codeCol: shift(columns.codeCol)!,
    nameCol: shift(columns.nameCol)!,
    plannedCol: shift(columns.plannedCol)!,
    unplannedCol: shift(columns.unplannedCol)!,
    morningCol: shift(columns.morningCol)!,
    eveningCol: shift(columns.eveningCol)!,
    totalCol: shift(columns.totalCol)!,
    cpCol: shift(columns.cpCol)!,
    selfieCol: shift(columns.selfieCol),
    teamCol: shift(columns.teamCol),
    regionCol: shift(columns.regionCol),
    cityCol: shift(columns.cityCol),
    designationCol: shift(columns.designationCol),
    remarksCol: insertAt,
  };
}

function ensureSelfieColumn(sheet: ExcelJS.Worksheet, columns: DailyColumns): DailyColumns {
  if (columns.selfieCol) {
    sheet.getColumn(columns.selfieCol).width = Math.max(
      sheet.getColumn(columns.selfieCol).width ?? 0,
      24,
    );
    sheet.getRow(columns.headerRow).getCell(columns.selfieCol).border = {
      ...(sheet.getRow(columns.headerRow).getCell(columns.selfieCol).border ?? {}),
      ...SELFIE_BORDER,
    };
    return columns;
  }

  const header = sheet.getRow(columns.headerRow);
  const col =
    Math.max(findLastUsedColumn(sheet, columns.headerRow), columns.cpCol, columns.totalCol) + 1;
  const headerCell = header.getCell(col);
  const sourceHeader = header.getCell(columns.cpCol);
  headerCell.value = "Selfies";
  headerCell.style = { ...sourceHeader.style };
  headerCell.border = { ...(headerCell.border ?? {}), ...SELFIE_BORDER };
  sheet.getColumn(col).width = Math.max(sheet.getColumn(col).width ?? 0, 24);
  header.commit();

  return { ...columns, selfieCol: col };
}

function findLastUsedColumn(sheet: ExcelJS.Worksheet, startRow = 1): number {
  let lastUsedCol = 0;
  for (
    let rowNumber = startRow;
    rowNumber <= Math.min(sheet.rowCount, startRow + 11);
    rowNumber++
  ) {
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

function writeUnmatchedDailyDefaults(row: ExcelJS.Row, columns: DailyColumns) {
  for (const col of [
    columns.plannedCol,
    columns.unplannedCol,
    columns.morningCol,
    columns.eveningCol,
    columns.totalCol,
  ]) {
    row.getCell(col).value = 0;
  }

  row.getCell(columns.cpCol).value = null;
  if (columns.remarksCol) {
    const remarksCell = row.getCell(columns.remarksCol);
    remarksCell.value = null;
    styleRemarksCell(remarksCell);
  }
  if (columns.selfieCol) {
    const selfieCell = row.getCell(columns.selfieCol);
    selfieCell.value = null;
    applyIsolatedSelfieStyle(selfieCell, undefined);
  }
  styleDailyCells(row, columns);
}

function clearDailyCells(row: ExcelJS.Row, columns: DailyColumns) {
  for (const col of [
    columns.plannedCol,
    columns.unplannedCol,
    columns.morningCol,
    columns.eveningCol,
    columns.totalCol,
    columns.cpCol,
    columns.remarksCol,
    columns.selfieCol,
  ].filter((col): col is number => Boolean(col))) {
    const cell = row.getCell(col);
    cell.value = null;
    if (col === columns.selfieCol) {
      applyIsolatedSelfieStyle(cell, undefined);
    } else if (col === columns.remarksCol) {
      styleRemarksCell(cell);
    }
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
    columns.remarksCol,
    columns.selfieCol,
  ].filter((col): col is number => Boolean(col))) {
    const cell = row.getCell(col);
    cell.alignment = { ...(cell.alignment ?? {}), horizontal: "center", vertical: "middle" };
  }
}

function applySelfieColumnStyles(sheet: ExcelJS.Worksheet, columns: DailyColumns) {
  if (!columns.selfieCol) return;

  const headerCell = sheet.getRow(columns.headerRow).getCell(columns.selfieCol);
  headerCell.border = { ...(headerCell.border ?? {}), ...SELFIE_BORDER };
  sheet.getColumn(columns.selfieCol).width = Math.max(
    sheet.getColumn(columns.selfieCol).width ?? 0,
    24,
  );

  for (let rowNumber = columns.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    row.eachCell((cell, colNumber) => {
      if (colNumber !== columns.selfieCol) clearKnownSelfieFill(cell);
    });
    const cell = row.getCell(columns.selfieCol);
    const text = cellText(cell);
    if (!text.trim()) continue;
    const count = parseSelfieCountFromCell(text);
    styleSelfieCell(cell, count);
  }
}

function applyRemarksColumnStyles(sheet: ExcelJS.Worksheet, columns: DailyColumns) {
  if (!columns.remarksCol) return;

  const headerCell = sheet.getRow(columns.headerRow).getCell(columns.remarksCol);
  headerCell.border = { ...(headerCell.border ?? {}), ...SELFIE_BORDER };
  sheet.getColumn(columns.remarksCol).width = Math.max(
    sheet.getColumn(columns.remarksCol).width ?? 0,
    20,
  );

  for (let rowNumber = columns.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber++) {
    styleRemarksCell(sheet.getRow(rowNumber).getCell(columns.remarksCol));
  }
}

function formatSelfieText(count: number): string {
  if (count <= 0) return "0 selfies 0 locations";
  if (count === 1) return "1 selfie with location";
  return `${count} selfies with locations`;
}

function parseSelfieCountFromCell(value: string): number {
  const match = value.match(/\b(\d+)\s*selfies?\b/i);
  return match ? Number(match[1]) : 0;
}

function styleSelfieCell(cell: ExcelJS.Cell, count: number) {
  if (count <= 0) {
    applyIsolatedSelfieStyle(cell, SELFIE_ZERO_FILL);
    return;
  }
  if (count <= 11) {
    applyIsolatedSelfieStyle(cell, SELFIE_WARNING_FILL);
    return;
  }

  applyIsolatedSelfieStyle(cell, undefined);
}

function styleRemarksCell(cell: ExcelJS.Cell) {
  cell.alignment = { ...(cell.alignment ?? {}), horizontal: "center", vertical: "middle" };
  cell.border = { ...(cell.border ?? {}), ...SELFIE_BORDER };
}

function applyIsolatedSelfieStyle(cell: ExcelJS.Cell, fill: ExcelJS.Fill | undefined) {
  const {
    fill: _sharedFill,
    border: sharedBorder,
    alignment: sharedAlignment,
    ...rest
  } = cell.style;
  cell.style = {
    ...rest,
    alignment: { ...(sharedAlignment ?? {}), horizontal: "center", vertical: "middle" },
    border: { ...(sharedBorder ?? {}), ...SELFIE_BORDER },
    ...(fill ? { fill: cloneFill(fill) } : {}),
  };
}

function cloneStyle(style: Partial<ExcelJS.Style>): Partial<ExcelJS.Style> {
  return JSON.parse(JSON.stringify(style ?? {})) as Partial<ExcelJS.Style>;
}

function cloneFill(fill: ExcelJS.Fill): ExcelJS.Fill {
  return JSON.parse(JSON.stringify(fill)) as ExcelJS.Fill;
}

function clearKnownSelfieFill(cell: ExcelJS.Cell) {
  const argb =
    cell.fill && "fgColor" in cell.fill ? cell.fill.fgColor?.argb?.toUpperCase() : undefined;
  if (argb !== SELFIE_WARNING_FILL.fgColor.argb && argb !== SELFIE_ZERO_FILL.fgColor.argb) return;

  const { fill: _sharedFill, ...rest } = cell.style;
  cell.style = { ...rest };
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

function teamNamesMatch(sourceValue: string, sheetName: string): boolean {
  const sourceKey = teamKey(sourceValue);
  const sheetKey = teamKey(sheetName);
  if (!sourceKey || !sheetKey) return false;
  return sourceKey.includes(sheetKey) || sheetKey.includes(sourceKey);
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

function selfieNamesMatch(a: string, b: string): boolean {
  if (!a || !b) return false;
  if (a === b) return true;
  const aTokens = a.split(" ").filter((token) => token.length > 1);
  const bTokens = b.split(" ").filter((token) => token.length > 1);
  if (!aTokens.length || !bTokens.length) return false;

  const overlap = aTokens.filter((token) => bTokens.includes(token)).length;
  return overlap >= Math.min(2, aTokens.length, bTokens.length);
}

function normalizeDateKey(value: unknown, text: string): string {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return localDateKey(value);
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
  if (!Number.isNaN(parsed.getTime())) return localDateKey(parsed);
  return "";
}

function localDateKey(date: Date): string {
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
  ].join("-");
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
