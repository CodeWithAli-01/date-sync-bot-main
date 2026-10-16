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
      ["123", "Ali", "2026-10-01", "09:00", "Planned", "Contact Point", ""],
    ],
    "monthly-call-log.xlsx",
  );

  const templateFile = await makeTemplateWorkbook("monthly-template.xlsx");
  const result = await processMonthlyPlannedReport(callLogFile, templateFile);

  assert.equal(result.performanceRows[0].planned, 2);
  assert.equal(result.performanceRows[0].unplanned, 1);
  assert.equal(result.performanceRows[0].totalCalls, 3);
  assert.equal(result.performanceRows[0].dailyDetails[0].morningMinutes, 240);
});
