import test from "node:test";
import assert from "node:assert/strict";
import { MONTHLY_REPORT_HELP, isMonthlyReportQuestion } from "./monthly-report-help.ts";
import { planSupportReply } from "./support-feedback.ts";

test("Monthly Report answers contain the approved sample, Daily Report PDFs and generation steps", () => {
  const plan = planSupportReply("How can I generate the Monthly Report?", null);
  assert.equal(plan.replies[0].text, MONTHLY_REPORT_HELP.english);
  assert.match(plan.replies[0].text, /Employee Code/);
  assert.match(plan.replies[0].text, /Name/);
  assert.match(plan.replies[0].text, /PDF files of the Daily Reports/);
  assert.match(plan.replies[0].text, /Generate Report/);
  assert.match(plan.replies[0].text, /Preview/);
  assert.match(plan.replies[0].text, /Download/);
  assert.doesNotMatch(plan.replies[0].text, /has not yet provided/);
  assert.doesNotMatch(plan.replies[0].text, /Call Log Live/);
  assert.equal(plan.monthlyReport, true);
  assert.equal(plan.dailyReport, false);
});

test("Monthly Report sample help follows the customer's language", () => {
  assert.equal(planSupportReply("monthly report kasa bnti ha", null).replies[0].text, MONTHLY_REPORT_HELP.romanUrdu);
  assert.equal(planSupportReply("ماہانہ رپورٹ کا نمونہ کیسے دیں؟", null).replies[0].text, MONTHLY_REPORT_HELP.urdu);
});

test("Monthly Report context never returns the Daily Report sample guide", () => {
  const plan = planSupportReply("sample mein kya hona chahiye", "romanUrdu", { language: "romanUrdu", dailyReport: false, monthlyReport: true });
  assert.equal(plan.replies[0].text, MONTHLY_REPORT_HELP.romanUrdu);
  assert.equal(plan.dailyReport, false);
  assert.equal(isMonthlyReportQuestion("MREP Monthly Report ka sample"), false);
});
