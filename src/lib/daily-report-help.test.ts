import test from "node:test";
import assert from "node:assert/strict";
import { DAILY_REPORT_HELP, getDailyReportHelp } from "./daily-report-help.ts";

test("Daily Report quick question returns the approved English guide", () => {
  assert.equal(getDailyReportHelp("How can I Genrate the Daily Report?"), DAILY_REPORT_HELP.english);
  assert.match(DAILY_REPORT_HELP.english, /Reports → Call Reports → Call Log Live/);
  assert.match(DAILY_REPORT_HELP.english, /Happened filter to Happened/);
  assert.match(DAILY_REPORT_HELP.english, /Employee Code and Name/);
  assert.match(DAILY_REPORT_HELP.english, /Preview/);
});

test("Daily Report guide follows Urdu and Roman Urdu questions", () => {
  assert.equal(getDailyReportHelp("ڈیلی رپورٹ کیسے بنانی ہے؟"), DAILY_REPORT_HELP.urdu);
  assert.equal(getDailyReportHelp("daily report kasa bnao aur kia files chahiye?"), DAILY_REPORT_HELP.romanUrdu);
});

test("Daily Report knowledge does not answer another report's quick question", () => {
  for (const report of ["Monthly Report", "Doctor Coverage", "Planned Summary", "MREP Monthly Report", "Avg. Per Day"]) {
    assert.equal(getDailyReportHelp(`How can I generate the ${report}?`), null);
  }
});
