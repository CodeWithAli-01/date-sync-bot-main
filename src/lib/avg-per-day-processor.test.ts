import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { processAvgPerDayReport } from "./avg-per-day-processor.ts";

async function file(rows: unknown[][], name = "report.xlsx") {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Team A");
  rows.forEach((row) => sheet.addRow(row));
  sheet.getCell("B2").font = { name: "Arial", size: 16 };
  sheet.getColumn(2).width = 31;
  return new File([await workbook.xlsx.writeBuffer()], name);
}

test("avg per day matches IDs, preserves sample names and styles, and copies source values", async () => {
  const sample = await file([["Employee Code", "Name", "Avg Calls"], [123, "Ali", 0], [456, "Same Name", 0], [789, "Missing", 99]]);
  const source = await file([["T. Code", "Emp. ID", "Employee Name", "Avg. Call Per Day.", "Morning Calls", "Evening Calls", "Total Calls", "Field Days", "Working Avg"],
    ["T-OTHER", "456", "Different Name", 6.5, 20, 19, 39, 6, "4 hours"],
    ["T-OTHER", "123", "Ali Khan", 9.25, 21, 16, 37, 4, "5 hours"]]);
  const result = await processAvgPerDayReport(sample, source);
  const output = new ExcelJS.Workbook();
  await output.xlsx.load(await result.blob.arrayBuffer());
  const sheet = output.worksheets[0];
  assert.equal(result.matchedEmployees, 2);
  assert.equal(sheet.getCell("B2").value, "Ali");
  assert.equal(sheet.getCell("B2").font.size, 16);
  assert.equal(sheet.getColumn(2).width, 31);
  assert.equal(sheet.getCell("C2").value, 9.25);
  assert.equal(sheet.getCell("C3").value, 6.5);
  assert.equal(sheet.getCell("D2").value, 21);
  assert.equal(sheet.getCell("E2").value, 16);
  assert.equal(sheet.getCell("F2").value, 37);
  assert.equal(sheet.getCell("G2").value, 4);
  assert.equal(sheet.columnCount, 7);
  assert.equal(sheet.getCell("C4").value, null);
  assert.equal(result.warnings.length, 1);
});

test("avg per day leaves conflicting IDs blank instead of picking another employee row", async () => {
  const result = await processAvgPerDayReport(
    await file([["Employee ID", "Name"], [123, "Ali"]]),
    await file([["Employee ID", "Morning Calls", "Evening Calls", "Total Calls", "Avg. Call Per Day.", "Field Days"], [123, 2, 1, 3, 3, 1], [123, 8, 1, 9, 9, 1]]),
  );
  assert.equal(result.matchedEmployees, 0);
  assert.match(result.warnings[0], /conflicting/);
});

test("avg per day selects the sample team for a duplicate ID and ignores old zero assignments", async () => {
  const sample = await file([["Employee ID", "Name"], [123, "Ali"], [456, "Ahmed"]]);
  const source = await file([
    ["Employee ID", "Team Name", "Morning Calls", "Evening Calls", "Total Calls", "Avg. Call Per Day.", "Field Days"],
    [123, "Team B", 50, 50, 100, 10, 10],
    [123, "Team A", 12, 6, 18, 9, 2],
    [456, "Old Team", 0, 0, 0, 0, 0],
    [456, "Current Team", 10, 5, 15, 5, 3],
  ]);
  const result = await processAvgPerDayReport(sample, source);
  const output = new ExcelJS.Workbook();
  await output.xlsx.load(await result.blob.arrayBuffer());
  assert.equal(result.matchedEmployees, 2);
  assert.equal(output.worksheets[0].getCell("E2").value, 18);
  assert.equal(output.worksheets[0].getCell("E3").value, 15);
});

test("avg per day reads cached formula results under two-row export headers", async () => {
  const result = await processAvgPerDayReport(
    await file([["Employee ID", "Name"], [123, "Ali"]]),
    await file([["Emp. ID", "Monthly Summary"], ["Emp. ID", "Morning Calls", "Evening Calls", "Total Calls", "Avg. Call Per Day.", "Field Days"], [123, 12, 6, 18, { formula: "18/2", result: 9 }, 2]]),
  );
  const output = new ExcelJS.Workbook();
  await output.xlsx.load(await result.blob.arrayBuffer());
  assert.equal(output.worksheets[0].getCell("F2").value, 9);
});
