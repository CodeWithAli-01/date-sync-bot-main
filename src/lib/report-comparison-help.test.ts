import test from "node:test";
import assert from "node:assert/strict";
import { getReportComparison, mentionedComparisonReports, isReportComparisonRequest } from "./report-comparison-help.ts";
import { planSupportReply } from "./support-feedback.ts";

const reports = ["Daily Report", "Monthly Report", "Doctor Coverage", "Planned Summary", "Avg. Per Day", "MREP Monthly Report"];

test("every pair of the six reports can be compared in English, Urdu and Roman Urdu", () => {
  for (let i = 0; i < reports.length; i++) for (let j = i + 1; j < reports.length; j++) {
    const question = `${reports[i]} vs ${reports[j]}`;
    assert.equal(mentionedComparisonReports(question).length, 2, question);
    for (const language of ["english", "romanUrdu", "urdu"] as const) {
      const comparison = getReportComparison(question, language);
      assert.ok(comparison, `${language}: ${question}`);
      assert.doesNotMatch(comparison, /Step 1|مرحلہ 1/);
    }
  }
});

test("all report comparison describes the six distinct purposes and source files", () => {
  const plan = planSupportReply("sari reports ka diifernce ak sath batao", null);
  assert.equal(plan.language, "romanUrdu");
  const text = plan.replies[0].text;
  for (const report of reports) assert.ok(text.includes(report), report);
  assert.equal(text.match(/^\d\. /gm)?.length, 6);
  assert.match(text, /PDF/);
  assert.match(text, /CP/);
  assert.match(text, /Field Days/);
  assert.match(text, /Coverage %/);
  assert.match(text, /Employee-wise Call Status Report/);
});

test("Monthly Report and MREP Monthly Report stay separate", () => {
  const text = getReportComparison("MREP Monthly Report aur Monthly Report mein farq?", "romanUrdu")!;
  assert.equal(text.match(/^\d\. /gm)?.length, 2);
  assert.match(text, /PDF/);
  assert.match(text, /Employee-wise Call Status Report Excel/);
  assert.deepEqual(mentionedComparisonReports("MREP Monthly Report vs Avg. Per Day"), ["avg", "mrep"]);
  assert.deepEqual(mentionedComparisonReports("Monthly Planned Unplanned vs Doctor Coverage"), ["doctor", "planned"]);
});

test("comparison replies follow the customer's language and tolerate difference typos", () => {
  assert.match(planSupportReply("Planned Summary or Avg. Per Day mn kia diifrence ha?", null).replies[0].text, /farq/);
  assert.match(planSupportReply("Explain the difference between Doctor Coverage and Avg. Per Day", null).replies[0].text, /different source files/);
  const urdu = planSupportReply("تمام رپورٹوں میں کیا فرق ہے؟", null);
  assert.match(urdu.replies[0].text, /فرق/);
  assert.equal(urdu.replies[0].text.match(/^\d\. /gm)?.length, 6);
});

test("an incomplete comparison asks for report names rather than choosing a guide", () => {
  const plan = planSupportReply("daily report ka difference batao", null);
  assert.match(plan.replies[0].text, /Do ya zyada reports/);
  assert.equal(plan.unresolvedCount, 1);
  assert.doesNotMatch(plan.replies[0].text, /Call Log Live|1\./);
});

test("a different month or a date range is not automatically a report comparison", () => {
  assert.equal(isReportComparisonRequest("select a different month"), false);
  assert.equal(isReportComparisonRequest("select dates between November 1 and November 10"), false);
  assert.equal(isReportComparisonRequest("how are all reports different?"), true);
});
