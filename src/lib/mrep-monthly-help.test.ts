import test from "node:test";
import assert from "node:assert/strict";
import { MREP_MONTHLY_HELP } from "./mrep-monthly-help.ts";
import { planSupportReply, SUPPORT_COPY } from "./support-feedback.ts";

test("MREP Monthly help uses the four approved steps and optional export date guidance", () => {
  const plan = planSupportReply("How can I Genrate the MREP Monthly Report?", null);
  assert.equal(plan.replies[0].text, MREP_MONTHLY_HELP.english);
  assert.match(plan.replies[0].text, /Reports → Call Reports → Employee-wise Call Status Report/);
  assert.match(plan.replies[0].text, /Select the required month/);
  assert.match(plan.replies[0].text, /Employee-wise Call Report/);
  assert.match(plan.replies[0].text, /WhatsApp Selfie Export Files/);
  assert.match(plan.replies[0].text, /dates and data present in the uploaded exports/);
  assert.match(plan.replies[0].text, /optional/);
  assert.match(plan.replies[0].text, /Preview/);
  assert.doesNotMatch(plan.replies[0].text, /PDF|Happened|Territory/);
  assert.equal(plan.mrepMonthly, true);
  assert.equal(plan.monthlyReport, false);
});

test("MREP Monthly guide responds in the customer's language", () => {
  assert.equal(planSupportReply("mrep monthly report kasa bnti ha", null).replies[0].text, MREP_MONTHLY_HELP.romanUrdu);
  assert.equal(planSupportReply("ایم ریپ ماہانہ رپورٹ کیسے بنائیں؟", null).replies[0].text, MREP_MONTHLY_HELP.urdu);
});

test("MREP follow-ups do not switch to the PDF-based Monthly Report", () => {
  const context = { language: "romanUrdu" as const, dailyReport: false, mrepMonthly: true };
  assert.equal(planSupportReply("whatsapp selfie file kaise upload krun", null, context).replies[0].text, MREP_MONTHLY_HELP.romanUrdu);
  assert.equal(planSupportReply("thnx", null, context).replies[0].text, SUPPORT_COPY.romanUrdu.thanks);
  const monthly = planSupportReply("monthly report kaise generate karun", null, context);
  assert.equal(monthly.monthlyReport, true);
  assert.equal(monthly.mrepMonthly, false);
});
