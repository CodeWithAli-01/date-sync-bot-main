import ExcelJS from "exceljs";

export interface AvgPerDayResult {
  blob: Blob;
  fileName: string;
  totalEmployees: number;
  matchedEmployees: number;
  warnings: string[];
}

const normalize = (value: string) => value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
const ID_HEADERS = ["employee code", "emp code", "employee id", "emp id", "id", "t code"];
const REPORT_FIELDS = [
  { label: "Morning Calls", aliases: ["morning calls", "total morning calls", "total morning calls individual", "mor"] },
  { label: "Evening Calls", aliases: ["evening calls", "total evening calls", "total evening calls individual", "eve"] },
  { label: "Total Calls", aliases: ["total calls", "total reported calls individual", "total reported calls", "total"] },
  { label: "Avg. Call Per Day.", aliases: ["avg call per day", "avg calls", "average call per day", "avg calls per day"] },
  { label: "Field Days", aliases: ["field days"] },
];
const codeKey = (value: string) => value.trim().toUpperCase().replace(/\s+/g, "").replace(/\.0+$/, "");

function text(cell: ExcelJS.Cell): string {
  return cell.text.trim();
}

function header(sheet: ExcelJS.Worksheet): { row: number; codeCol: number; columns: Map<string, number> } | null {
  for (let r = 1; r <= Math.min(sheet.rowCount, 30); r++) {
    const columns = new Map<string, number>();
    sheet.getRow(r).eachCell((cell, col) => {
      const label = normalize(text(cell));
      if (label && !columns.has(label)) columns.set(label, col);
    });
    const codeCol = ID_HEADERS.map((label) => columns.get(label)).find(Boolean);
    if (codeCol) return { row: r, codeCol, columns };
  }
  return null;
}

function sourceValue(cell: ExcelJS.Cell): ExcelJS.CellValue {
  const value = cell.value;
  // Copy cached results, not formulas whose references belong to the source workbook.
  if (value && typeof value === "object" && "formula" in value) return value.result ?? null;
  if (value && typeof value === "object" && "sharedFormula" in value) return value.result ?? null;
  return structuredClone(value);
}

export async function processAvgPerDayReport(sampleFile: File, sourceFile: File): Promise<AvgPerDayResult> {
  const sample = new ExcelJS.Workbook();
  const source = new ExcelJS.Workbook();
  await sample.xlsx.load(await sampleFile.arrayBuffer());
  await source.xlsx.load(await sourceFile.arrayBuffer());
  const metrics = new Map(REPORT_FIELDS.map((field) => [normalize(field.label), field.label]));
  const byCode = new Map<string, { team: string; values: Map<string, ExcelJS.CellValue> }[]>();
  let sourceSheets = 0;
  for (const sheet of source.worksheets) {
    const identity = header(sheet);
    if (!identity) continue;
    // Some exports have identity headers on row 1 and the metric headings on row 2.
    const next = sheet.getRow(identity.row + 1);
    const metricHeader = ID_HEADERS.includes(normalize(text(next.getCell(identity.codeCol)))) ? next : sheet.getRow(identity.row);
    const fields: [string, number][] = [];
    const sourceColumns = new Map<string, number>();
    metricHeader.eachCell((cell, col) => {
      sourceColumns.set(normalize(text(cell)), col);
    });
    for (const field of REPORT_FIELDS) {
      const col = field.aliases.map((label) => sourceColumns.get(label)).find(Boolean);
      if (col) fields.push([normalize(field.label), col]);
    }
    if (fields.length !== REPORT_FIELDS.length) continue;
    sourceSheets++;
    const teamCol = sourceColumns.get("team name") ?? sourceColumns.get("team");
    for (let r = metricHeader.number + 1; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      const code = codeKey(text(row.getCell(identity.codeCol)));
      if (!code || ID_HEADERS.includes(normalize(code))) continue;
      const values = new Map(fields.map(([key, col]) => [key, sourceValue(row.getCell(col))]));
      const candidates = byCode.get(code) ?? [];
      candidates.push({ team: teamCol ? normalize(text(row.getCell(teamCol))) : "", values });
      byCode.set(code, candidates);
    }
  }
  if (!sourceSheets) throw new Error("The avg report must contain Employee Code, Morning Calls, Evening Calls, Total Calls, Avg. Call Per Day., and Field Days.");
  let totalEmployees = 0;
  let matchedEmployees = 0;
  let missingEmployees = 0;
  let ambiguousEmployees = 0;
  for (const sheet of sample.worksheets) {
    const identity = header(sheet);
    if (!identity) continue;
    const columns = new Map<string, number>();
    for (const [key, label] of metrics) {
      const field = REPORT_FIELDS.find((item) => normalize(item.label) === key)!;
      let col = field.aliases.map((alias) => identity.columns.get(alias)).find(Boolean);
      if (!col) {
        col = sheet.columnCount + 1;
        const cell = sheet.getRow(identity.row).getCell(col);
        cell.style = structuredClone(sheet.getRow(identity.row).getCell(col - 1).style);
        cell.value = label;
        sheet.getColumn(col).width = sheet.getColumn(col - 1).width ?? 18;
        for (let r = identity.row + 1; r <= sheet.rowCount; r++) {
          sheet.getRow(r).getCell(col).style = structuredClone(sheet.getRow(r).getCell(col - 1).style);
        }
      }
      sheet.getRow(identity.row).getCell(col).value = label;
      columns.set(key, col);
    }
    for (let r = identity.row + 1; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      const codeCell = row.getCell(identity.codeCol);
      const code = codeKey(text(codeCell));
      if (!code || !/\d/.test(code) || codeCell.isMerged) continue;
      totalEmployees++;
      let candidates = byCode.get(code) ?? [];
      const teamMatches = candidates.filter((candidate) => candidate.team === normalize(sheet.name));
      if (teamMatches.length) candidates = teamMatches;
      // Old assignments often have an all-zero row alongside the current assignment.
      const withData = candidates.filter((candidate) => [...candidate.values.values()].some((value) => value !== null && value !== "" && Number(value) !== 0));
      if (withData.length) candidates = withData;
      const unique = new Map(candidates.map((candidate) => [JSON.stringify([...candidate.values]), candidate.values]));
      const values = unique.size === 1 ? [...unique.values()][0] : undefined;
      if (unique.size > 1) ambiguousEmployees++;
      else if (!values) missingEmployees++;
      else matchedEmployees++;
      for (const [key, col] of columns) {
        row.getCell(col).value = values ? values.get(key) ?? null : null;
      }
    }
  }
  if (!totalEmployees) throw new Error("No employee IDs were found in the sample Excel.");
  const buffer = await sample.xlsx.writeBuffer();
  return {
    blob: new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    fileName: sampleFile.name.replace(/\.[^.]+$/, "") + " - Avg. Per Day.xlsx",
    totalEmployees,
    matchedEmployees,
    warnings: [
      ...(missingEmployees ? [`${missingEmployees} employee ID(s) were not found in the source report; their report values are blank.`] : []),
      ...(ambiguousEmployees ? [`${ambiguousEmployees} employee row(s) have conflicting source values for the same ID; their report values are blank.`] : []),
    ],
  };
}
