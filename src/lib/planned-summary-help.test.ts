import test from "node:test";
import assert from "node:assert/strict";
import { PLANNED_SUMMARY_HELP } from "./planned-summary-help.ts";
import { planSupportReply, SUPPORT_COPY } from "./support-feedback.ts";

test("Planned Summary answers use the approved six steps", () => {
  const plan = planSupportReply("How can I Genrate the Planned Summary rport?", null);
  assert.equal(plan.replies[0].text, PLANNED_SUMMARY_HELP.english);
  assert.match(plan.replies[0].text, /Employee Code and Name/);
  assert.match(plan.replies[0].text, /Reports → Call Reports → Call Log/);
  assert.match(plan.replies[0].text, /Select all the dates/);
  assert.match(plan.replies[0].text, /Happened filter to Happened/);
  assert.match(plan.replies[0].text, /Monthly Call Log Excel/);
  assert.match(plan.replies[0].text, /Preview/);
  assert.match(plan.replies[0].text, /Download/);
  assert.doesNotMatch(plan.replies[0].text, /PDF|Territory/);
  assert.equal(plan.plannedSummary, true);
  assert.equal(plan.dailyReport, false);
  assert.equal(plan.monthlyReport, false);
  assert.equal(plan.doctorCoverage, false);
});

test("Planned Summary help follows Urdu and Roman Urdu questions", () => {
  assert.equal(planSupportReply("planned summary report kasa bnti ha", null).replies[0].text, PLANNED_SUMMARY_HELP.romanUrdu);
  assert.equal(planSupportReply("پلانڈ سمری رپورٹ کیسے بنائیں؟", null).replies[0].text, PLANNED_SUMMARY_HELP.urdu);
});

test("Planned Summary follow-ups retain the right report without interrupting ordinary chat", () => {
  const context = { language: "romanUrdu" as const, dailyReport: false, plannedSummary: true };
  for (const question of ["happened filter kahan hai", "monthly call log excel mein kya upload krna hai"]) {
    const plan = planSupportReply(question, null, context);
    assert.equal(plan.replies[0].text, PLANNED_SUMMARY_HELP.romanUrdu);
    assert.equal(plan.monthlyReport, false);
  }
  assert.equal(planSupportReply("thnx", null, context).replies[0].text, SUPPORT_COPY.romanUrdu.thanks);
  assert.match(planSupportReply("mn bi thk", null, context).replies[0].text, /sun kar acha laga/);
  const monthly = planSupportReply("monthly report kasa bnti ha", null, context);
  assert.equal(monthly.plannedSummary, false);
  assert.equal(monthly.monthlyReport, true);
});
