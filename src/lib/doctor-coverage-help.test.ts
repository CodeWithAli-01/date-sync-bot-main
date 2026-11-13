import test from "node:test";
import assert from "node:assert/strict";
import { DOCTOR_COVERAGE_HELP } from "./doctor-coverage-help.ts";
import { planSupportReply, SUPPORT_COPY } from "./support-feedback.ts";

test("Doctor Coverage quick question uses all four approved steps", () => {
  const plan = planSupportReply("How can I Genrate the Doctor Coverage rport?", null);
  assert.equal(plan.replies[0].text, DOCTOR_COVERAGE_HELP.english);
  assert.match(plan.replies[0].text, /Employee Code and Name/);
  assert.match(plan.replies[0].text, /Reports → Call Reports → Doctor Coverage/);
  assert.match(plan.replies[0].text, /month, then select the team/);
  assert.match(plan.replies[0].text, /Territory, choose Select All/);
  assert.match(plan.replies[0].text, /Filter/);
  assert.match(plan.replies[0].text, /upload that downloaded file/);
  assert.match(plan.replies[0].text, /Generate Report/);
  assert.match(plan.replies[0].text, /Preview/);
  assert.doesNotMatch(plan.replies[0].text, /Call Log Live|Happened|PDF/);
  assert.equal(plan.doctorCoverage, true);
  assert.equal(plan.dailyReport, false);
  assert.equal(plan.monthlyReport, false);
});

test("Doctor Coverage guidance supports Urdu and Roman Urdu including the admin's spelling", () => {
  assert.equal(planSupportReply("doctor covrage report kasa banani ha", null).replies[0].text, DOCTOR_COVERAGE_HELP.romanUrdu);
  assert.equal(planSupportReply("ڈاکٹر کوریج رپورٹ کیسے بنائیں؟", null).replies[0].text, DOCTOR_COVERAGE_HELP.urdu);
});

test("Doctor Coverage follow-ups retain the correct guide without interrupting small talk", () => {
  const context = { language: "romanUrdu" as const, dailyReport: false, monthlyReport: false, doctorCoverage: true };
  assert.equal(planSupportReply("territory mein kya select karna hai", null, context).replies[0].text, DOCTOR_COVERAGE_HELP.romanUrdu);
  assert.equal(planSupportReply("mrep se file download kaise krun", null, context).replies[0].text, DOCTOR_COVERAGE_HELP.romanUrdu);
  assert.match(planSupportReply("mn bi thk", null, context).replies[0].text, /sun kar acha laga/);
  assert.equal(planSupportReply("thnx", null, context).replies[0].text, SUPPORT_COPY.romanUrdu.thanks);
});
