import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";

import { processDailyReport } from "./daily-report-processor.ts";

async function makeCallLogWorkbook(rows: unknown[][], fileName: string) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Call Log");
  rows.forEach((row) => sheet.addRow(row));
  const buffer = await workbook.xlsx.writeBuffer();
  return new File([buffer], fileName, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

async function makeTemplateWorkbook(fileName: string) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Team A");
  sheet.addRow(["Employee Code", "Name", "Planned", "Unplanned", "Morning", "Evening", "Total", "CP"]);
  sheet.addRow(["123", "Ali", 0, 0, 0, 0, 0, null]);
  const buffer = await workbook.xlsx.writeBuffer();
  return new File([buffer], fileName, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

async function makeSelfieWorkbook(fileName: string) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Chat");
  sheet.addRow(["Message Time", "Message Type", "Message Body"]);
  sheet.addRow(["2026-10-01", "Image", "image attachment"]);
  const buffer = await workbook.xlsx.writeBuffer();
  return new File([buffer], fileName, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

test("daily report matches selfie files by employee name when code is missing", async () => {
  const callLogFile = await makeCallLogWorkbook(
    [
      ["Emp ID", "Employee Name", "Date", "Start Time", "Event Type", "Meeting Type", "Shift"],
      ["123", "Ali", "2026-10-01", "08:00", "Planned Call", "Face-to-Face", "Morning"],
    ],
    "call-log.xlsx",
  );
  const templateFile = await makeTemplateWorkbook("template.xlsx");
  const selfieFile = await makeSelfieWorkbook("Ali.xlsx");

  const report = await processDailyReport(callLogFile, templateFile, selfieFile);

  assert.equal(report.debug.selfieMatchedEmployees, 1);
  assert.equal(report.debug.selfieFallbackMatches, 1);
  assert.equal(report.performanceRows[0].selfies, 1);
});

test("daily report ignores similar but different employee names when matching selfies", async () => {
  const callLogFile = await makeCallLogWorkbook(
    [
      ["Emp ID", "Employee Name", "Date", "Start Time", "Event Type", "Meeting Type", "Shift"],
      ["123", "Ali", "2026-10-01", "08:00", "Planned Call", "Face-to-Face", "Morning"],
    ],
    "call-log.xlsx",
  );
  const templateFile = await makeTemplateWorkbook("template.xlsx");
  const selfieFile = await makeSelfieWorkbook("Ali Raza.xlsx");

  const report = await processDailyReport(callLogFile, templateFile, selfieFile);

  assert.equal(report.debug.selfieMatchedEmployees, 0);
  assert.equal(report.debug.selfieMissingSources, 1);
  assert.equal(report.performanceRows[0].selfies, 0);
});
