import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";

import { processMonthlyPlannedReport } from "./monthly-planned-processor.ts";

async function makeWorkbook(rows: unknown[][], fileName: string) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Call Log");
  rows.forEach((row) => sheet.addRow(row));
  const buffer = await workbook.xlsx.writeBuffer();
  return new File([buffer], fileName, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

async function makeTemplateWorkbook(fileName: string) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Template");
  sheet.addRow(["Employee Code", "Name"]);
  sheet.addRow(["123", "Ali"]);
  const buffer = await workbook.xlsx.writeBuffer();
  return new File([buffer], fileName, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

test("monthly report counts planned call variants and AM/PM shifts correctly", async () => {
  const callLogFile = await makeWorkbook(
    [
      ["Emp ID", "Employee Name", "Date", "Start Time", "Event Type", "Meeting Type", "Shift"],
      ["123", "Ali", "2026-10-01", "08:00", "Planned Call", "Face-to-Face", "AM"],
      ["123", "Ali", "2026-10-01", "12:00", "Planned Call", "Face-to-Face", "AM"],
      ["123", "Ali", "2026-10-01", "18:00", "Unplanned Call", "Face-to-Face", "PM"],
      ["123", "Ali", "2026-10-01", "07:30", "Planned", "Contact Point", ""],
      ["123", "Ali", "2026-10-01", "10:00", "Planned", "Meeting", "Morning"],
      ["123", "Ali", "2026-10-01", "11:00", "Unplanned", "Leave", "Evening"],
    ],
    "monthly-call-log.xlsx",
  );

  const templateFile = await makeTemplateWorkbook("monthly-template.xlsx");
  const result = await processMonthlyPlannedReport(callLogFile, templateFile);

  assert.equal(result.performanceRows[0].planned, 2);
  assert.equal(result.performanceRows[0].unplanned, 1);
  assert.equal(result.performanceRows[0].totalCalls, 3);
  assert.equal(result.performanceRows[0].morningCalls, 2);
  assert.equal(result.performanceRows[0].eveningCalls, 1);
  assert.equal(result.performanceRows[0].dailyDetails[0].morningMinutes, 270);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await result.blob.arrayBuffer());
  const sheet = workbook.getWorksheet("Template")!;
  assert.equal(sheet.getCell("M1").value, "Total Morning Calls");
  assert.equal(sheet.getCell("N1").value, "Total Evening Calls");
  assert.equal(sheet.getCell("M2").value, 2);
  assert.equal(sheet.getCell("N2").value, 1);
});

test("monthly report ignores contact-point-only days when calculating average work metrics", async () => {
  const callLogFile = await makeWorkbook(
    [
      ["Emp ID", "Employee Name", "Date", "Start Time", "Event Type", "Meeting Type", "Shift"],
      ["123", "Ali", "2026-10-01", "08:00", "Planned Call", "Face-to-Face", "Morning"],
      ["123", "Ali", "2026-10-01", "07:30", "Planned Call", "Contact Point", ""],
      ["123", "Ali", "2026-10-01", "12:00", "Planned Call", "Face-to-Face", "Morning"],
      ["123", "Ali", "2026-10-01", "18:00", "Planned Call", "Face-to-Face", "Evening"],
      ["123", "Ali", "2026-10-02", "11:00", "Planned Call", "Contact Point", ""],
    ],
    "monthly-call-log-avg.xlsx",
  );

  const templateFile = await makeTemplateWorkbook("monthly-template-avg.xlsx");
  const result = await processMonthlyPlannedReport(callLogFile, templateFile);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await result.blob.arrayBuffer());
  const row = workbook.getWorksheet("Template")?.getRow(2);

  assert.equal(row?.getCell(3).value, 3);
  assert.equal(row?.getCell(9).value, "9:15 AM");
  assert.equal(row?.getCell(10).value, "4 hours 30 min");
  assert.equal(row?.getCell(11).value, "0 hours 0 min");
  assert.equal(row?.getCell(12).value, "4 hours 30 min");
});

test("monthly report aggregates employees across files and counts every face-to-face call", async () => {
  const header = ["Emp ID", "Employee Name", "Date", "Start Time", "Event Type", "Meeting Type", "Shift"];
  const first = await makeWorkbook([header,
    ["123", "Ali", "2026-10-01", "8:30 AM", "Planned", "Contact Point", ""],
    ["123", "Ali", "2026-10-01", "9:30 AM", "Planned", "Face to Face Call", "Morning"],
    ["123", "Ali", "2026-10-01", "12:30 PM", "Other", "Face to Face Call", "Morning"],
    ["123", "Ali", "2026-10-01", "2:00 PM", "Planned", "Meeting", "Evening"],
  ], "first.xlsx");
  const second = await makeWorkbook([header,
    ["123", "Ali", "2026-10-02", "9:30 AM", "Planned", "Contact Point", ""],
    ["123", "Ali", "2026-10-02", "10:00 AM", "Unplanned", "Face to Face Call", "Morning"],
    ["123", "Ali", "2026-10-02", "12:00 PM", "Planned", "Face to Face Call", "Morning"],
  ], "second.xlsx");
  const result = await processMonthlyPlannedReport([first, second], await makeTemplateWorkbook("template.xlsx"));
  const employee = result.performanceRows[0];
  assert.equal(employee.planned, 2);
  assert.equal(employee.unplanned, 1);
  assert.equal(employee.totalCalls, 4);
  assert.equal(employee.cpAvgTime, "9:00 AM");
  assert.deepEqual(employee.dailyDetails.map((day) => day.morningMinutes), [240, 150]);
});

test("working averages use CP to last morning call and first to last evening call", async () => {
  const file = await makeWorkbook([
    ["Emp ID", "Employee Name", "Date", "Start Time", "Event Type", "Meeting Type", "Shift"],
    ["123", "Ali", "2026-10-01", "8:00 AM", "Planned", "Contact Point", ""],
    ["123", "Ali", "2026-10-01", "9:00 AM", "Planned", "Face to Face Call", "Morning"],
    ["123", "Ali", "2026-10-01", "12:00 PM", "Planned", "Face to Face Call", "Morning"],
    ["123", "Ali", "2026-10-01", "4:00 PM", "Planned", "Face to Face Call", "Evening"],
    ["123", "Ali", "2026-10-01", "7:30 PM", "Unplanned", "Face to Face Call", "Evening"],
    ["123", "Ali", "2026-10-02", "9:00 AM", "Planned", "Face to Face Call", "Morning"],
  ], "calls.xlsx");
  const result = await processMonthlyPlannedReport(file, await makeTemplateWorkbook("template.xlsx"));
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await result.blob.arrayBuffer());
  const row = workbook.worksheets[0].getRow(2);
  assert.equal(row.getCell(10).value, "4 hours 0 min");
  assert.equal(row.getCell(11).value, "3 hours 30 min");
  assert.equal(row.getCell(12).value, "7 hours 30 min");
});
