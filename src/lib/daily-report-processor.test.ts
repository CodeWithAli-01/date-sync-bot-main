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

test("daily report does not match a selfie file without the employee code", async () => {
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

  assert.equal(report.debug.selfieMatchedEmployees, 0);
  assert.equal(report.performanceRows[0].selfies, 0);
});

test("daily report does not count selfies from other dates", async () => {
  const callLogFile = await makeCallLogWorkbook(
    [
      ["Emp ID", "Employee Name", "Date", "Start Time", "Event Type", "Meeting Type", "Shift"],
      ["123", "Ali", "2026-10-01", "08:00", "Planned Call", "Face-to-Face", "Morning"],
    ],
    "call-log.xlsx",
  );
  const templateFile = await makeTemplateWorkbook("template.xlsx");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Chat");
  sheet.addRow(["Message Time", "Message Type", "Message Body"]);
  sheet.addRow(["2026-10-02", "Image", "image attachment"]);
  const buffer = await workbook.xlsx.writeBuffer();
  const selfieFile = new File([buffer], "Ali [123].xlsx", {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const report = await processDailyReport(callLogFile, templateFile, selfieFile);

  assert.equal(report.debug.selfieMatchedEmployees, 1);
  assert.equal(report.debug.selfieFallbackMatches, 0);
  assert.equal(report.debug.selfieDateMisses, 1);
  assert.equal(report.performanceRows[0].selfies, 0);
});

test("daily report counts only image rows for the matching employee and call date", async () => {
  const callLogFile = await makeCallLogWorkbook(
    [
      ["Emp ID", "Employee Name", "Date", "Start Time", "Event Type", "Meeting Type", "Shift"],
      ["123", "Ali", "2026-10-01", "08:00", "Planned Call", "Face-to-Face", "Morning"],
    ],
    "call-log.xlsx",
  );
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Chat");
  sheet.addRow(["Message Time", "Message Type", "Message Body"]);
  sheet.addRow(["2026-10-01 08:00", "Image", "image attachment"]);
  sheet.addRow(["2026-10-01 09:00", "Image", "image attachment"]);
  sheet.addRow(["2026-10-02", "Image", "image attachment"]);
  sheet.addRow(["2026-10-01", "Text", "hello"]);
  const buffer = await workbook.xlsx.writeBuffer();
  const matching = new File([buffer], "Ali [123].xlsx");
  const otherEmployee = new File([buffer], "Ali [456].xlsx");
  const report = await processDailyReport(callLogFile, await makeTemplateWorkbook("template.xlsx"), [matching, otherEmployee]);
  assert.equal(report.performanceRows[0].selfies, 2);
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
