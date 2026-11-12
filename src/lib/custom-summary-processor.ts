import ExcelJS from "exceljs";

export interface CustomSummaryResult {
  fileName: string;
  sheetName: string;
  totalEmployees: number;
  matchedEmployees: number;
  warnings: string[];
  preview: { name: string; total: number }[];
  blob: Blob;
}

export interface ProcessCustomSummaryInput {
  sampleFile: File;
  callReportFile: File;
  whatsappFiles: File[];
}

const REPORT_FIELD_ALIASES: Record<string, string[]> = {
  totalPlanCallDone: [
    "total plan call done individual",
    "total plan call done",
    "plan call done individual",
    "plan call done",
    "planned calls",
    "total planned",
  ],
  totalUnplannedCall: [
    "total unplanned call individual",
    "total unplanned call",
    "unplanned calls",
    "total unplanned",
  ],
  totalMorningCalls: [
    "total morning calls individual",
    "total morning calls",
    "morning calls",
    "morning total",
  ],
  totalEveningCalls: [
    "total evening calls individual",
    "total evening calls",
    "evening calls",
    "evening total",
  ],
  totalReportedCalls: [
    "total reported calls individual",
    "total reported calls",
    "reported calls",
    "total calls",
  ],
  avgCallPerDay: ["avg call per day", "average call per day", "avg calls per day"],
  targetDoctor: ["target doctor", "target doctors", "doctor target"],
  totalVisitedDoctors: ["total visited doctors", "visited doctors", "total doctors visited"],
  coveragePercent: ["coverage %", "coverage percentage", "coverage"],
  meetingDays: ["meeting days includes all meetings", "meeting days", "days meetings"],
  leaveDays: ["leave days", "days leave", "total leave days"],
};

export async function processCustomSummaryReport({
  sampleFile,
  callReportFile,
  whatsappFiles,
}: ProcessCustomSummaryInput): Promise<CustomSummaryResult> {
  const outputWorkbook = new ExcelJS.Workbook();
  await outputWorkbook.xlsx.load(await sampleFile.arrayBuffer());
  const callMap = await readEmployeeCallReport(callReportFile);
  const selfieMap = await readWhatsappSelfieCounts(whatsappFiles);
  const preview: { name: string; total: number }[] = [];
  let totalEmployees = 0;
  let matchedEmployees = 0;
  let missingCallEmployees = 0;
  const fields: { key: keyof EmployeeMetricRecord | "selfieCount" | "selfieAvg"; label: string; aliases: string[] }[] = [
    { key: "totalPlanCallDone", label: "Total Plan Call Done Individual", aliases: [...REPORT_FIELD_ALIASES.totalPlanCallDone, "planned"] },
    { key: "totalUnplannedCall", label: "Total Unplanned Call (Individual)", aliases: [...REPORT_FIELD_ALIASES.totalUnplannedCall, "unplanned"] },
    { key: "totalMorningCalls", label: "Total Morning Calls (Individual)", aliases: [...REPORT_FIELD_ALIASES.totalMorningCalls, "mor", "morning"] },
    { key: "totalEveningCalls", label: "Total Evening Calls (Individual)", aliases: [...REPORT_FIELD_ALIASES.totalEveningCalls, "eve", "evening"] },
    { key: "totalReportedCalls", label: "Total Reported Calls (Individual)", aliases: [...REPORT_FIELD_ALIASES.totalReportedCalls, "total"] },
    ...(["avgCallPerDay", "targetDoctor", "totalVisitedDoctors", "coveragePercent", "meetingDays", "leaveDays"] as const).map((key) => ({ key, label: REPORT_FIELD_ALIASES[key][0], aliases: REPORT_FIELD_ALIASES[key] })),
    { key: "selfieCount", label: "Selfies Count", aliases: ["selfies count", "selfies", "selfie"] },
    { key: "selfieAvg", label: "Selfies Avg/Day", aliases: ["selfies avg day", "selfies average day"] },
  ];

  for (const sheet of outputWorkbook.worksheets) {
    const headerRowIndex = findIdentityHeaderRow(sheet);
    if (!headerRowIndex) continue;
    const header = sheet.getRow(headerRowIndex);
    const codeCol = findExactHeaderColumn(header, ["employee code", "emp code", "employee id", "emp id", "t code", "code"]);
    const nameCol = findExactHeaderColumn(header, ["employee name", "emp name", "name"]);
    if (!codeCol && !nameCol) continue;
    const columns = new Map<string, number>();
    for (const field of fields) {
      let col = findExactHeaderColumn(header, field.aliases);
      if (!col) {
        col = sheet.columnCount + 1;
        const cell = header.getCell(col);
        cell.style = structuredClone(header.getCell(Math.max(1, col - 1)).style);
        cell.value = field.label;
        sheet.getColumn(col).width = sheet.getColumn(Math.max(1, col - 1)).width;
        for (let rowNumber = headerRowIndex + 1; rowNumber <= sheet.rowCount; rowNumber++) {
          const row = sheet.getRow(rowNumber);
          row.getCell(col).style = structuredClone(row.getCell(Math.max(1, col - 1)).style);
        }
      }
      columns.set(field.key, col);
    }
    for (let rowNumber = headerRowIndex + 1; rowNumber <= sheet.rowCount; rowNumber++) {
      const row = sheet.getRow(rowNumber);
      const code = codeCol ? normalizeEmployeeCode(readCellValue(row.getCell(codeCol))) : "";
      const name = nameCol ? readCellValue(row.getCell(nameCol)).trim() : "";
      if (!code && !name) continue;
      if (code && !/\d/.test(code)) continue;
      totalEmployees++;
      const metrics = code ? callMap.get(code) : undefined;
      if (!metrics) missingCallEmployees++;
      const selfieCount = code ? (selfieMap.get(code) ?? 0) : 0;
      if (metrics || selfieCount > 0) matchedEmployees++;
      for (const field of fields) {
        const value = field.key === "selfieCount" ? selfieCount : field.key === "selfieAvg" ? Math.round(selfieCount / 26 * 10) / 10 : metrics?.[field.key] ?? 0;
        row.getCell(columns.get(field.key)!).value = value;
      }
      preview.push({ name: name || code, total: metrics?.totalReportedCalls ?? 0 });
    }
  }
  const buffer = await outputWorkbook.xlsx.writeBuffer();
  return {
    fileName: sampleFile.name.replace(/\.[^.]+$/i, "") + "-kpi.xlsx",
    sheetName: outputWorkbook.worksheets[0]?.name ?? "MREP Monthly KPI",
    totalEmployees,
    matchedEmployees,
    warnings: totalEmployees === 0
      ? ["No employee rows were detected in the sample workbook."]
      : missingCallEmployees
        ? [`${missingCallEmployees} sample employee row(s) have no matching employee code in the call report; their call metrics are 0.`]
        : [],
    preview: preview.sort((a, b) => b.total - a.total).slice(0, 10),
    blob: new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
  };
}

function findExactHeaderColumn(row: ExcelJS.Row, hints: string[]): number {
  for (const hint of hints) for (let col = 1; col <= row.cellCount; col++) {
    if (normalizeText(readCellValue(row.getCell(col))) === normalizeText(hint)) return col;
  }
  return 0;
}

function findIdentityHeaderRow(sheet: ExcelJS.Worksheet): number {
  for (let i = 1; i <= Math.min(sheet.rowCount, 30); i++) {
    const row = sheet.getRow(i);
    if (findExactHeaderColumn(row, ["employee code", "emp code", "employee id", "emp id", "t code", "code"]) || findExactHeaderColumn(row, ["employee name", "emp name", "name"])) return i;
  }
  return 0;
}

function normalizeEmployeeCode(value: string): string {
  return value.trim().toUpperCase().replace(/\s+/g, "").replace(/\.0+$/, "");
}

interface EmployeeRecord {
  employeeCode: string;
  name: string;
  region?: string;
  designation?: string;
  teamName?: string;
}

interface EmployeeMetricRecord {
  totalPlanCallDone?: number;
  totalUnplannedCall?: number;
  totalMorningCalls?: number;
  totalEveningCalls?: number;
  totalReportedCalls?: number;
  avgCallPerDay?: number;
  targetDoctor?: number;
  totalVisitedDoctors?: number;
  coveragePercent?: number;
  meetingDays?: number;
  leaveDays?: number;
}

async function readSampleEmployeeRows(file: File): Promise<EmployeeRecord[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const rows: EmployeeRecord[] = [];
  const isMultiSheetWorkbook = workbook.worksheets.length > 1;

  for (const sheet of workbook.worksheets) {
    const headerMap = detectHeaderMap(sheet);
    const headerRowIndex = findHeaderRowIndex(sheet, [
      "employee code",
      "t code",
      "employee name",
      "employee id",
      "emp id",
      "name",
      "region",
      "designation",
    ]);

    if (!headerRowIndex) continue;

    const teamTitleAboveHeader =
      headerRowIndex > 1 && !!readCellValue(sheet.getRow(1).getCell(1)).trim();
    const useSheetNameAsTeam = isMultiSheetWorkbook || teamTitleAboveHeader;

    for (let rowNumber = headerRowIndex + 1; rowNumber <= sheet.rowCount; rowNumber++) {
      const row = sheet.getRow(rowNumber);
      const codeCell = safeCell(row, headerMap.codeCol);
      const nameCell = safeCell(row, headerMap.nameCol);
      const employeeCode = codeCell ? readCellValue(codeCell).trim() : "";
      const name = nameCell ? readCellValue(nameCell).trim() : "";
      if (!employeeCode && !name) continue;

      rows.push({
        employeeCode: employeeCode || "",
        name: normalizeEmployeeName(name),
        region: headerMap.regionCol
          ? safeCell(row, headerMap.regionCol)
            ? readCellValue(safeCell(row, headerMap.regionCol)!).trim()
            : ""
          : "",
        designation: headerMap.designationCol
          ? safeCell(row, headerMap.designationCol)
            ? readCellValue(safeCell(row, headerMap.designationCol)!).trim()
            : ""
          : "",
        teamName: useSheetNameAsTeam ? sheet.name : undefined,
      });
    }
  }

  return rows.filter((row) => row.employeeCode || row.name);
}

async function readEmployeeCallReport(file: File): Promise<Map<string, EmployeeMetricRecord>> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const result = new Map<string, EmployeeMetricRecord>();
  const nameBuckets = new Map<string, EmployeeMetricRecord[]>();
  const nameToCodes = new Map<string, Set<string>>();

  const upsertMetric = (key: string, metrics: EmployeeMetricRecord) => {
    const existing = result.get(key);
    if (!existing || shouldReplaceMetric(existing, metrics)) {
      result.set(key, metrics);
    }
  };

  for (const sheet of workbook.worksheets) {
    const headerMap = detectCallHeaderMap(sheet);
    const headerRowIndex = findHeaderRowIndex(sheet, [
      "employee code",
      "employee name",
      "employee id",
      "name",
      "total reported calls",
      "coverage",
      "meeting days",
    ]);

    if (!headerRowIndex) continue;

    for (let rowNumber = headerRowIndex + 1; rowNumber <= sheet.rowCount; rowNumber++) {
      const row = sheet.getRow(rowNumber);
      if (!hasMeaningfulRow(row)) continue;

      const codeCell = safeCell(row, headerMap.codeCol);
      const nameCell = safeCell(row, headerMap.nameCol);
      if (!codeCell && !nameCell) continue;

      const employeeCode = codeCell ? normalizeEmployeeCode(readCellValue(codeCell)) : "";
      const rawName = nameCell ? readCellValue(nameCell).trim() : "";
      const key = employeeKey(employeeCode, rawName);
      const metrics: EmployeeMetricRecord = {
        totalPlanCallDone: readNumericCell(row, headerMap.columns.totalPlanCallDone),
        totalUnplannedCall: readNumericCell(row, headerMap.columns.totalUnplannedCall),
        totalMorningCalls: readNumericCell(row, headerMap.columns.totalMorningCalls),
        totalEveningCalls: readNumericCell(row, headerMap.columns.totalEveningCalls),
        totalReportedCalls: readNumericCell(row, headerMap.columns.totalReportedCalls),
        avgCallPerDay: readNumericCell(row, headerMap.columns.avgCallPerDay),
        targetDoctor: readNumericCell(row, headerMap.columns.targetDoctor),
        totalVisitedDoctors: readNumericCell(row, headerMap.columns.totalVisitedDoctors),
        coveragePercent: readNumericCell(row, headerMap.columns.coveragePercent),
        meetingDays: readNumericCell(row, headerMap.columns.meetingDays),
        leaveDays: readNumericCell(row, headerMap.columns.leaveDays),
      };

      if (employeeCode || rawName) {
        upsertMetric(key, metrics);
        if (employeeCode) upsertMetric(employeeCode, metrics);

        if (rawName) {
          const nameKey = employeeKeyFromName(rawName);
          const bucket = nameBuckets.get(nameKey) ?? [];
          bucket.push(metrics);
          nameBuckets.set(nameKey, bucket);

          if (employeeCode) {
            const codes = nameToCodes.get(nameKey) ?? new Set<string>();
            codes.add(employeeCode);
            nameToCodes.set(nameKey, codes);
          }
        }
      }
    }
  }

  for (const [nameKey, metricsList] of nameBuckets) {
    const codes = nameToCodes.get(nameKey) ?? new Set<string>();
    if (codes.size !== 1) continue;

    const bestMetric = metricsList.reduce((best, current) =>
      shouldReplaceMetric(best, current) ? current : best,
    );
    result.set(nameKey, bestMetric);
  }

  return result;
}

function shouldReplaceMetric(current: EmployeeMetricRecord, candidate: EmployeeMetricRecord): boolean {
  const currentScore = metricStrength(current);
  const candidateScore = metricStrength(candidate);
  if (candidateScore !== currentScore) return candidateScore > currentScore;
  return Number(candidate.totalReportedCalls ?? 0) > Number(current.totalReportedCalls ?? 0);
}

function metricStrength(metrics: EmployeeMetricRecord): number {
  let score = 0;
  for (const value of [
    metrics.totalPlanCallDone,
    metrics.totalUnplannedCall,
    metrics.totalMorningCalls,
    metrics.totalEveningCalls,
    metrics.totalReportedCalls,
    metrics.avgCallPerDay,
    metrics.targetDoctor,
    metrics.totalVisitedDoctors,
    metrics.coveragePercent,
    metrics.meetingDays,
    metrics.leaveDays,
  ]) {
    if (value != null && value !== 0) score += 1;
  }
  return score;
}

async function readWhatsappSelfieCounts(files: File[]): Promise<Map<string, number>> {
  const result = new Map<string, number>();
  const duplicateNameKeys = new Set<string>();

  for (const file of files) {
    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.load(await file.arrayBuffer());
    } catch {
      continue;
    }

    const parsedName = parseEmployeeNameFromWhatsappFile(file.name);
    const employeeCode = normalizeEmployeeCode(parsedName.code);
    const employeeName = parsedName.name;

    let total = 0;
    for (const worksheet of workbook.worksheets) {
      for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        for (let colNumber = 1; colNumber <= worksheet.columnCount; colNumber++) {
          const value = readCellValue(row.getCell(colNumber));
          if (isImageString(value)) total += 1;
        }
      }
    }

    if (total > 0) {
      if (employeeCode) result.set(employeeCode, total);

      if (employeeName) {
        const nameKey = employeeKeyFromName(employeeName);
        if (duplicateNameKeys.has(nameKey)) {
          result.delete(nameKey);
        } else if (result.has(nameKey)) {
          duplicateNameKeys.add(nameKey);
          result.delete(nameKey);
        } else {
          result.set(nameKey, total);
        }
      }

      result.set(employeeKey(employeeCode, employeeName), total);
    }
  }
  return result;
}

function detectHeaderMap(sheet: ExcelJS.Worksheet): {
  codeCol: number;
  nameCol: number;
  regionCol?: number;
  designationCol?: number;
} {
  const firstRow = sheet.getRow(1);
  const codeCol =
    findHeaderColumn(firstRow, ["employee code", "emp code", "employee id", "emp id", "t code", "code"]) ||
    inferHeaderColumn(sheet, ["employee code", "emp code", "employee id", "emp id", "t code", "code"]);
  const nameCol =
    findHeaderColumn(firstRow, ["name", "employee name", "emp name"]) ||
    inferHeaderColumn(sheet, ["name", "employee name", "emp name"]);
  const regionCol =
    findHeaderColumn(firstRow, ["region", "zone", "area"]) ||
    inferHeaderColumn(sheet, ["region", "zone", "area"]);
  const designationCol =
    findHeaderColumn(firstRow, ["designation", "desig", "title", "position"]) ||
    inferHeaderColumn(sheet, ["designation", "desig", "title", "position"]);

  return {
    codeCol: codeCol || 1,
    nameCol: nameCol || 2,
    regionCol: regionCol || undefined,
    designationCol: designationCol || undefined,
  };
}

function detectCallHeaderMap(sheet: ExcelJS.Worksheet): {
  codeCol: number;
  nameCol: number;
  columns: {
    totalPlanCallDone: number;
    totalUnplannedCall: number;
    totalMorningCalls: number;
    totalEveningCalls: number;
    totalReportedCalls: number;
    avgCallPerDay: number;
    targetDoctor: number;
    totalVisitedDoctors: number;
    coveragePercent: number;
    meetingDays: number;
    leaveDays: number;
  };
} {
  const firstRow = sheet.getRow(1);
  const columns = {
    totalPlanCallDone:
      findHeaderColumn(firstRow, REPORT_FIELD_ALIASES.totalPlanCallDone) ||
      inferHeaderColumn(sheet, REPORT_FIELD_ALIASES.totalPlanCallDone),
    totalUnplannedCall:
      findHeaderColumn(firstRow, REPORT_FIELD_ALIASES.totalUnplannedCall) ||
      inferHeaderColumn(sheet, REPORT_FIELD_ALIASES.totalUnplannedCall),
    totalMorningCalls:
      findHeaderColumn(firstRow, REPORT_FIELD_ALIASES.totalMorningCalls) ||
      inferHeaderColumn(sheet, REPORT_FIELD_ALIASES.totalMorningCalls),
    totalEveningCalls:
      findHeaderColumn(firstRow, REPORT_FIELD_ALIASES.totalEveningCalls) ||
      inferHeaderColumn(sheet, REPORT_FIELD_ALIASES.totalEveningCalls),
    totalReportedCalls:
      findHeaderColumn(firstRow, REPORT_FIELD_ALIASES.totalReportedCalls) ||
      inferHeaderColumn(sheet, REPORT_FIELD_ALIASES.totalReportedCalls),
    avgCallPerDay:
      findHeaderColumn(firstRow, REPORT_FIELD_ALIASES.avgCallPerDay) ||
      inferHeaderColumn(sheet, REPORT_FIELD_ALIASES.avgCallPerDay),
    targetDoctor:
      findHeaderColumn(firstRow, REPORT_FIELD_ALIASES.targetDoctor) ||
      inferHeaderColumn(sheet, REPORT_FIELD_ALIASES.targetDoctor),
    totalVisitedDoctors:
      findHeaderColumn(firstRow, REPORT_FIELD_ALIASES.totalVisitedDoctors) ||
      inferHeaderColumn(sheet, REPORT_FIELD_ALIASES.totalVisitedDoctors),
    coveragePercent:
      findHeaderColumn(firstRow, REPORT_FIELD_ALIASES.coveragePercent) ||
      inferHeaderColumn(sheet, REPORT_FIELD_ALIASES.coveragePercent),
    meetingDays:
      findHeaderColumn(firstRow, REPORT_FIELD_ALIASES.meetingDays) ||
      inferHeaderColumn(sheet, REPORT_FIELD_ALIASES.meetingDays),
    leaveDays:
      findHeaderColumn(firstRow, REPORT_FIELD_ALIASES.leaveDays) ||
      inferHeaderColumn(sheet, REPORT_FIELD_ALIASES.leaveDays),
  };

  const identityHeader = sheet.getRow(findIdentityHeaderRow(sheet) || 1);
  const codeCol = findExactHeaderColumn(identityHeader, ["employee code", "emp code", "employee id", "emp id", "t code", "code"]);
  const nameCol =
    findHeaderColumn(firstRow, ["employee name", "emp name", "name"]) ||
    inferHeaderColumn(sheet, ["employee name", "emp name", "name"]);

  return {
    codeCol,
    nameCol: nameCol || 2,
    columns,
  };
}

function findHeaderRowIndex(sheet: ExcelJS.Worksheet, hints: string[]): number {
  const maxRows = Math.min(sheet.rowCount, 30);
  for (let rowNumber = 1; rowNumber <= maxRows; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    for (let col = 1; col <= Math.max(row.cellCount, 30); col++) {
      const value = normalizeText(readCellValue(row.getCell(col)));
      if (!value) continue;
      if (hints.some((hint) => {
        const normalizedHint = normalizeText(hint);
        return value === normalizedHint || value.includes(normalizedHint);
      })) {
        return rowNumber;
      }
    }
  }
  return 0;
}

function findHeaderColumn(row: ExcelJS.Row, hints: string[]): number {
  for (let col = 1; col <= row.cellCount; col++) {
    const value = normalizeText(readCellValue(row.getCell(col)));
    if (!value) continue;
    for (const hint of hints) {
      if (value.includes(normalizeText(hint))) return col;
    }
  }
  return 0;
}

function inferHeaderColumn(sheet: ExcelJS.Worksheet, hints: string[]): number {
  const maxRows = Math.min(sheet.rowCount, 10);
  for (let rowNumber = 1; rowNumber <= maxRows; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    for (let col = 1; col <= Math.max(row.cellCount, 10); col++) {
      const value = normalizeText(readCellValue(row.getCell(col)));
      if (!value) continue;
      if (hints.some((hint) => value.includes(normalizeText(hint)))) {
        return col;
      }
    }
  }
  return 0;
}

function styleTitleRow(sheet: ExcelJS.Worksheet, columnCount: number, title: string): void {
  const titleCell = sheet.getCell(1, 1);
  titleCell.value = title;
  titleCell.font = { bold: true, size: 28, color: { argb: "FF000000" } };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  titleCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF00" },
  };
  titleCell.border = {
    top: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
  };
  sheet.getRow(1).height = 28;

  sheet.columns.forEach((column, index) => {
    if (index < columnCount) {
      column.width = 18;
    }
  });
}

function styleReportSheet(sheet: ExcelJS.Worksheet, columnCount: number): void {
  for (let i = 1; i <= columnCount; i++) {
    const cell = sheet.getCell(2, i);
    cell.font = { bold: true, size: 11, color: { argb: "FF000000" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF00" },
    };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };
  }

  for (let rowNumber = 3; rowNumber <= sheet.rowCount; rowNumber++) {
    const row = sheet.getRow(rowNumber);
    for (let col = 1; col <= columnCount; col++) {
      const cell = row.getCell(col);
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFDAF7A6" },
      };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
    }
  }

  sheet.columns.forEach((column) => {
    const values = column.values ?? [];
    const maxLength = values.reduce<number>((max, value) => {
      if (value == null) return max;
      const length = String(value).length;
      return Math.max(max, length + 2);
    }, 12);
    column.width = Number(Math.min(maxLength, 24));
  });
}

function readNumericCell(row: ExcelJS.Row, col: number): number | undefined {
  if (!col) return undefined;
  const cell = safeCell(row, col);
  if (!cell) return undefined;
  const raw = readCellValue(cell).trim();
  if (!raw) return undefined;

  const normalized = raw
    .replace(/[$€£¥]/g, "")
    .replace(/[%\s]/g, "")
    .replace(/,/g, "")
    .replace(/\((.*)\)/, "-$1")
    .trim();

  const value = Number(normalized);
  return Number.isFinite(value) ? value : undefined;
}

function safeCell(row: ExcelJS.Row, column: number): ExcelJS.Cell | null {
  if (!column || column < 1) return null;
  return row.getCell(column);
}

function normalizeEmployeeName(value: string): string {
  return value.replace(/\s+/g, " ").trim().toLowerCase();
}

function normalizeText(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function employeeKey(code: string, name: string): string {
  return `${normalizeText(code || "")}::${normalizeEmployeeName(name)}`;
}

function employeeKeyFromName(name: string): string {
  return `::${normalizeEmployeeName(name)}`;
}

function parseEmployeeNameFromWhatsappFile(name: string): { code: string; name: string } {
  const match = name.match(/(\d+)\s*[-_]\s*([^\.]+)/i);
  const code = match ? match[1].trim() : "";
  const label = match ? match[2].replace(/\s+/g, " ").trim() : name.replace(/\.[^.]+$/i, "").trim();
  return { code, name: normalizeEmployeeName(label) };
}

function hasMeaningfulRow(row: ExcelJS.Row): boolean {
  for (let col = 1; col <= row.cellCount; col++) {
    const value = readCellValue(row.getCell(col));
    if (value && value.toString().trim()) return true;
  }
  return false;
}

function readCellValue(cell: ExcelJS.Cell | null | undefined): string {
  if (cell == null) return "";
  const value = cell.value;
  if (value == null) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "true" : "false";
  if (typeof value === "object") {
    if ("text" in value && typeof value.text === "string") return value.text;
    if ("result" in value && value.result != null) return String(value.result);
    if ("richText" in value && Array.isArray((value as { richText?: Array<{ text?: string }> }).richText)) {
      return (value as { richText: Array<{ text?: string }> }).richText
        .map((part) => part.text ?? "")
        .join("");
    }
  }
  return String(value);
}

function isImageString(value: string): boolean {
  const text = value.toLowerCase();
  return /\.(jpg|jpeg|png|gif|webp|bmp|heic|heif)/.test(text) || text.includes("image") || text.includes("photo") || text.includes("selfie");
}
