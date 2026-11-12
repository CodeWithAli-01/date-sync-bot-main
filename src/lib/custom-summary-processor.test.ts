import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { processCustomSummaryReport } from "./custom-summary-processor.ts";

async function createWorkbook(rows: unknown[][]): Promise<Blob> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Test Sheet");
  rows.forEach((row) => sheet.addRow(row));
  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

test("custom summary report does not crash when a sheet is missing expected columns", async () => {
  const sampleFile = new File(
    [await createWorkbook([["Employee Name", "Region"], ["Ali", "SGD"]])],
    "sample.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const callFile = new File(
    [await createWorkbook([["Name", "Total Calls"], ["Ali", 10]])],
    "call-report.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const result = await processCustomSummaryReport({
    sampleFile,
    callReportFile: callFile,
    whatsappFiles: [],
  });

  assert.equal(result.totalEmployees, 1);
  assert.equal(result.preview.length, 1);
  assert.equal(result.sheetName, "Test Sheet");
});

test("custom summary report supports team sheets with a title row above the employee headers", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("ALPHA 3");
  sheet.addRow(["ALPHA 3"]);
  sheet.addRow(["SR #", "Region", "Employee Code", "Name", "City", "Designation"]);
  sheet.addRow([1, "SGD", 53081, "Malik Farooq", "SGD", "HOS"]);
  sheet.getCell("D3").font = { name: "Arial", size: 15, bold: true };
  sheet.getColumn(4).width = 32;

  const sampleFile = new File(
    [await workbook.xlsx.writeBuffer()],
    "sample-team.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const callFile = new File(
    [await createWorkbook([["Employee Code", "Employee Name", "Total Reported Calls"], [53081, "Malik Farooq", 20]])],
    "call-report.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const result = await processCustomSummaryReport({
    sampleFile,
    callReportFile: callFile,
    whatsappFiles: [],
  });

  assert.equal(result.totalEmployees, 1);
  assert.equal(result.matchedEmployees, 1);
  assert.ok(result.preview.some((item) => item.name.toLowerCase() === "malik farooq"));

  const generatedWorkbook = new ExcelJS.Workbook();
  await generatedWorkbook.xlsx.load(await new Blob([new Uint8Array(await result.blob.arrayBuffer())], { type: result.blob.type }).arrayBuffer());
  const createdSheet = generatedWorkbook.getWorksheet("ALPHA 3");
  assert.equal(createdSheet?.getRow(1).getCell(1).value, "ALPHA 3");
  assert.equal(createdSheet?.getRow(2).getCell(1).value, "SR #");
  assert.equal(createdSheet?.getCell("D3").value, "Malik Farooq");
  assert.equal(createdSheet?.getCell("D3").font.size, 15);
  assert.equal(createdSheet?.getColumn(4).width, 32);
});

test("custom summary report recognizes real KPI header labels with parenthetical headings", async () => {
  const sampleFile = new File(
    [await createWorkbook([["Employee Code", "Employee Name", "Region", "Designation"], [53081, "Malik Farooq", "LHR", "ASM"]])],
    "sample.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const callFile = new File(
    [await createWorkbook([
      ["Employee Code", "Employee Name", "Total Plan Call Done Individual", "Total Unplanned Call (Individual)", "Total Morning Calls (Individual)", "Total Evening Calls (Individual)", "Total Reported Calls (Individual)", "Avg Call Per Day", "Target Doctor", "Total Visited Doctors", "Coverage %", "Meeting Days (Includes ALL Meetings)", "Leave Days"],
      [53081, "Malik Farooq", 3, 2, 5, 6, 16, 1.5, 20, 18, 80, 14, 1],
    ])],
    "call-report.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const result = await processCustomSummaryReport({
    sampleFile,
    callReportFile: callFile,
    whatsappFiles: [],
  });

  assert.equal(result.totalEmployees, 1);
  assert.equal(result.matchedEmployees, 1);
  assert.equal(result.preview[0]?.total, 16);
});

test("custom summary report keeps totals when KPI values are comma-formatted", async () => {
  const sampleFile = new File(
    [await createWorkbook([["Employee Code", "Employee Name", "Region", "Designation"], [53081, "Malik Farooq", "LHR", "ASM"]])],
    "sample.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const callFile = new File(
    [await createWorkbook([
      ["Employee Code", "Employee Name", "Total Plan Call Done Individual", "Total Unplanned Call (Individual)", "Total Morning Calls (Individual)", "Total Evening Calls (Individual)", "Total Reported Calls (Individual)", "Avg Call Per Day", "Target Doctor", "Total Visited Doctors", "Coverage %", "Meeting Days (Includes ALL Meetings)", "Leave Days"],
      [53081, "Malik Farooq", "3,200", "1,500", "2,200", "1,800", "8,700", "335", "20", "18", "80%", "14", "1"],
    ])],
    "call-report.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const result = await processCustomSummaryReport({
    sampleFile,
    callReportFile: callFile,
    whatsappFiles: [],
  });

  assert.equal(result.totalEmployees, 1);
  assert.equal(result.matchedEmployees, 1);
  assert.equal(result.preview[0]?.total, 8700);
});

test("custom summary report matches employees regardless of name casing differences", async () => {
  const sampleFile = new File(
    [await createWorkbook([["Employee Code", "Employee Name", "Region", "Designation"], [53081, "ANSAR ALI", "LHR", "ASM"]])],
    "sample.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const callFile = new File(
    [await createWorkbook([
      ["Employee Code", "Employee Name", "Total Plan Call Done Individual", "Total Unplanned Call (Individual)", "Total Morning Calls (Individual)", "Total Evening Calls (Individual)", "Total Reported Calls (Individual)", "Avg Call Per Day", "Target Doctor", "Total Visited Doctors", "Coverage %", "Meeting Days (Includes ALL Meetings)", "Leave Days"],
      [53081, "Ansar Ali", 12, 8, 7, 9, 36, 1.4, 20, 18, 80, 14, 1],
    ])],
    "call-report.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const result = await processCustomSummaryReport({
    sampleFile,
    callReportFile: callFile,
    whatsappFiles: [],
  });

  assert.equal(result.matchedEmployees, 1);
  assert.equal(result.preview[0]?.total, 36);
});

test("custom summary report does not guess metrics when duplicate names are ambiguous and no code is available", async () => {
  const sampleFile = new File(
    [await createWorkbook([
      ["Employee Name", "Region", "Designation"],
      ["Ali Khan", "LHR", "ASM"],
      ["Ali Khan", "RWP", "ASM"],
    ])],
    "sample-duplicate-name.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const callFile = new File(
    [await createWorkbook([
      ["Employee Code", "Employee Name", "Total Plan Call Done Individual", "Total Unplanned Call (Individual)", "Total Morning Calls (Individual)", "Total Evening Calls (Individual)", "Total Reported Calls (Individual)", "Avg Call Per Day", "Target Doctor", "Total Visited Doctors", "Coverage %", "Meeting Days (Includes ALL Meetings)", "Leave Days"],
      [101, "Ali Khan", 5, 2, 7, 1, 15, 1.2, 20, 18, 75, 10, 0],
      [102, "Ali Khan", 9, 3, 10, 2, 24, 1.8, 22, 19, 80, 12, 1],
    ])],
    "call-report-duplicate-name.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const result = await processCustomSummaryReport({
    sampleFile,
    callReportFile: callFile,
    whatsappFiles: [],
  });

  assert.equal(result.totalEmployees, 2);
  assert.equal(result.matchedEmployees, 0);
  assert.deepEqual(result.preview.map((item) => item.total), [0, 0]);
});

test("custom summary report supports T. Code / Emp. ID employee headers from the MREP file", async () => {
  const sampleFile = new File(
    [await createWorkbook([
      ["Team", "Zone", "Region", "T. Code", "Designation", "Emp. ID", "Employee Name"],
      ["ALPHA 3", "CENTER-1", "LHR", 53081, "HOS", 53081, "Malik Farooq"],
      ["ALPHA 3", "CENTER-1", "LHR", 23781, "ASM", 23781, "Nadeem Hussain"],
    ])],
    "sample-mrep.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const callFile = new File(
    [await createWorkbook([
      ["Employee Code", "Employee Name", "Total Plan Call Done Individual", "Total Unplanned Call (Individual)", "Total Morning Calls (Individual)", "Total Evening Calls (Individual)", "Total Reported Calls (Individual)", "Avg Call Per Day", "Target Doctor", "Total Visited Doctors", "Coverage %", "Meeting Days (Includes ALL Meetings)", "Leave Days"],
      [53081, "Malik Farooq", 7, 3, 9, 7, 26, 1.0, 20, 18, 80, 14, 1],
      [23781, "Nadeem Hussain", 11, 4, 15, 6, 36, 1.3, 22, 20, 85, 10, 0],
    ])],
    "call-report-mrep.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const result = await processCustomSummaryReport({
    sampleFile,
    callReportFile: callFile,
    whatsappFiles: [],
  });

  assert.equal(result.totalEmployees, 2);
  assert.equal(result.matchedEmployees, 2);
  assert.deepEqual(
    result.preview.slice().sort((a, b) => b.total - a.total).map((item) => item.total),
    [36, 26],
  );
});

test("monthly KPI source matches Emp. ID instead of the different T. Code", async () => {
  const sampleFile = new File([await createWorkbook([
    ["Employee Code", "Name", "Total"],
    [53081, "Malik Farooq", 0],
    [98605, "Zohaib Hassan", 0],
  ])], "sample.xlsx");
  const callFile = new File([await createWorkbook([
    ["Team", "T. Code", "Emp. ID", "Employee Name", "Monthly Summary", "Monthly Summary"],
    ["Team", "T. Code", "Emp. ID", "Employee Name", "Total Plan Call Done Individual", "Total Reported Calls (Individual)"],
    ["ALPHA", "AP3-HOS", "53081", "MALIK FAROOQ", 0, 151],
    ["ALPHA", "AP3-MIO", "98605", "ZOHAIB HASSAN", 156, 159],
  ])], "source.xlsx");
  const result = await processCustomSummaryReport({ sampleFile, callReportFile: callFile, whatsappFiles: [] });
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await result.blob.arrayBuffer());
  assert.equal(result.matchedEmployees, 2);
  assert.equal(workbook.worksheets[0].getCell("C2").value, 151);
  assert.equal(workbook.worksheets[0].getCell("C3").value, 159);
});

test("custom summary report prefers the stronger duplicate-name row and exact employee code before name fallback", async () => {
  const sampleFile = new File(
    [await createWorkbook([
      ["Employee Code", "Employee Name", "Region", "Designation"],
      [101, "Ali Khan", "LHR", "ASM"],
      [102, "Ali Khan", "RWP", "ASM"],
    ])],
    "sample-duplicate-stronger-row.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const callFile = new File(
    [await createWorkbook([
      ["Employee Code", "Employee Name", "Total Plan Call Done Individual", "Total Unplanned Call (Individual)", "Total Morning Calls (Individual)", "Total Evening Calls (Individual)", "Total Reported Calls (Individual)", "Avg Call Per Day", "Target Doctor", "Total Visited Doctors", "Coverage %", "Meeting Days (Includes ALL Meetings)", "Leave Days"],
      [101, "Ali Khan", 5, 2, 7, 1, 15, 1.2, 20, 18, 75, 10, 0],
      [101, "Ali Khan", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
      [102, "Ali Khan", 9, 3, 10, 2, 24, 1.8, 22, 19, 80, 12, 1],
    ])],
    "call-report-duplicate-stronger-row.xlsx",
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  );

  const result = await processCustomSummaryReport({
    sampleFile,
    callReportFile: callFile,
    whatsappFiles: [],
  });

  assert.equal(result.totalEmployees, 2);
  assert.equal(result.matchedEmployees, 2);
  assert.deepEqual(
    result.preview.slice().sort((a, b) => b.total - a.total).map((item) => item.total),
    [24, 15],
  );
});
