import test from "node:test";
import assert from "node:assert/strict";
import { AVG_PER_DAY_HELP } from "./avg-per-day-help.ts";
import { planSupportReply, SUPPORT_COPY } from "./support-feedback.ts";

test("Avg. Per Day quick question returns the approved three-step guide", () => {
  const plan = planSupportReply("How can I Genrate the Avg. Per Day rport?", null);
  assert.equal(plan.replies[0].text, AVG_PER_DAY_HELP.english);
  assert.match(plan.replies[0].text, /Employee Code \/ Emp\. ID and Name/);
  assert.match(plan.replies[0].text, /Reports → Call Reports → Avg Per Day/);
  assert.match(plan.replies[0].text, /Select the month/);
  assert.match(plan.replies[0].text, /Filter/);
  assert.match(plan.replies[0].text, /Avg Report Excel/);
  assert.match(plan.replies[0].text, /Morning Calls, Evening Calls, Total Calls/);
  assert.match(plan.replies[0].text, /Preview/);
  assert.match(plan.replies[0].text, /Download/);
  assert.doesNotMatch(plan.replies[0].text, /Call Log Live|Happened|Territory|PDF/);
  assert.equal(plan.avgPerDay, true);
  assert.equal(plan.dailyReport, false);
  assert.equal(plan.monthlyReport, false);
});

test("Avg. Per Day guidance follows Urdu and Roman Urdu questions", () => {
  assert.equal(planSupportReply("avg per day report kasa bnti ha", null).replies[0].text, AVG_PER_DAY_HELP.romanUrdu);
  assert.equal(planSupportReply("ایوریج پر ڈے رپورٹ کیسے بنائیں؟", null).replies[0].text, AVG_PER_DAY_HELP.urdu);
});

test("Avg. Per Day follow-ups keep the right report and still answer small talk", () => {
  const context = { language: "romanUrdu" as const, dailyReport: false, avgPerDay: true };
  assert.equal(planSupportReply("month select karke file kahan upload krun", null, context).replies[0].text, AVG_PER_DAY_HELP.romanUrdu);
  assert.equal(planSupportReply("thnx", null, context).replies[0].text, SUPPORT_COPY.romanUrdu.thanks);
  assert.match(planSupportReply("mn bi thk", null, context).replies[0].text, /sun kar acha laga/);
  const monthly = planSupportReply("monthly report kasa bnti ha", null, context);
  assert.equal(monthly.avgPerDay, false);
  assert.equal(monthly.monthlyReport, true);
});
