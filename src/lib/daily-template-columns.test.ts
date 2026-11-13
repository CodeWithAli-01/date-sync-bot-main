import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { processDailyReport, processBulkDailyReports } from "./daily-report-processor.ts";

async function file(workbook: ExcelJS.Workbook, name: string) {
  return new File([await workbook.xlsx.writeBuffer()], name);
}

async function callLog() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Call Log");
  sheet.addRow(["Emp ID", "Employee Name", "Date", "Start Time", "Event Type", "Meeting Type", "Shift"]);
  sheet.addRow(["123", "Ali", "2026-11-10", "09:00", "Planned", "Face to Face Call", "Morning"]);
  sheet.addRow(["123", "Ali", "2026-11-10", "17:00", "Unplanned", "Face to Face Call", "Evening"]);
  return file(workbook, "call-log.xlsx");
}

test("Daily Report adds missing metrics while preserving employee columns, styles and merged headings", async () => {
  for (const partial of [false, true]) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Team A");
    sheet.mergeCells("A1:F1");
    sheet.getCell("A1").value = "Team A";
    sheet.addRow(["SR #", "Region", "Employee Code ", "Name", "City", "Designation", ...(partial ? ["Morning"] : [])]);
    sheet.addRow([1, "LHR", "123", "Ali", "Lahore", "MIO", ...(partial ? [99] : [])]);
    sheet.getCell("F2").font = { name: "Arial", size: 14, bold: true };
    sheet.getCell("F3").font = { name: "Arial", size: 12 };
    const result = await processDailyReport(await callLog(), await file(workbook, "sample.xlsx"));
    assert.equal(result.matchedEmployees, 1);
    const output = new ExcelJS.Workbook();
    await output.xlsx.load(await result.blob.arrayBuffer());
    const out = output.worksheets[0];
    assert.deepEqual(out.getRow(3).values.slice(1, 7), [1, "LHR", "123", "Ali", "Lahore", "MIO"]);
    assert.equal(out.getCell("A1").value, "Team A");
    assert.equal(out.getCell("B1").isMerged, true);
    for (const [aliases, value] of [[ ["Planned"], 1 ], [ ["Unplanned"], 1 ], [ ["Mor", "Morning"], 1 ], [ ["Eve"], 1 ], [ ["Total"], 2 ]] as const) {
      const cols: number[] = [];
      out.getRow(2).eachCell((cell, col) => { if ((aliases as readonly string[]).includes(cell.text)) cols.push(col); });
      assert.equal(cols.length, 1, aliases.join("/"));
      assert.equal(out.getRow(3).getCell(cols[0]).value, value);
    }
    if (!partial) {
      assert.equal(out.getCell("G2").font?.name, "Arial");
      assert.equal(out.getCell("G2").font?.size, 14);
    }
  }
});

test("bulk Daily Report accepts identity-only samples in multiple team sheets", async () => {
  const workbook = new ExcelJS.Workbook();
  for (const name of ["Team A", "Team B"]) {
    const sheet = workbook.addWorksheet(name);
    sheet.addRow(["Employee Code", "Name"]);
    sheet.addRow(["123", "Ali"]);
  }
  const result = await processBulkDailyReports(await callLog(), await file(workbook, "teams.xlsx"));
  assert.equal(result.summary.length, 2);
  assert.ok(result.summary.every((item) => item.status === "success" && item.matchedEmployees === 1));
});
