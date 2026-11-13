import test from "node:test";
import assert from "node:assert/strict";
import { planSupportReply, SUPPORT_COPY } from "./support-feedback.ts";
import { ADMIN_CONTACT } from "./support-feedback.ts";

test("a short overview covers all six reports without changing the approved methods", () => {
  const plan = planSupportReply("mere pass time kam ha sari report ki short summary bata do", null);
  const text = plan.replies[0].text;
  assert.equal(plan.language, "romanUrdu");
  for (const name of ["Daily Report", "Monthly Report", "Doctor Coverage", "Planned Summary", "Avg. Per Day", "MREP Monthly Report"]) assert.ok(text.includes(name), name);
  assert.match(text, /Daily Report PDF/);
  assert.match(text, /Employee-wise Call Status Report/);
  assert.match(text, /Territory mein Select All/);
  assert.match(text, /Happened/);
  assert.match(text, /Generate Report/);
  assert.equal(plan.unresolvedCount, 0);
});

test("all-report overview follows English and Urdu too", () => {
  assert.match(planSupportReply("Give me a short summary of all reports", null).replies[0].text, /selected date/);
  const urdu = planSupportReply("تمام رپورٹوں کا مختصر خلاصہ بتائیں", null);
  assert.equal(urdu.language, "urdu");
  assert.match(urdu.replies[0].text, /ڈیلی رپورٹ/);
});

test("single-report summary stays concise and does not include the other guides", () => {
  const plan = planSupportReply("Avg. Per Day ki short summary batao", null);
  assert.match(plan.replies[0].text, /Avg Report Excel/);
  assert.doesNotMatch(plan.replies[0].text, /WhatsApp|Call Log Live|Territory/);
  assert.ok(plan.replies[0].text.length < 750);
});

test("MREP login is outside the approved help scope even after a report conversation", () => {
  for (const context of [undefined, { language: "romanUrdu" as const, dailyReport: true }, { language: "romanUrdu" as const, dailyReport: false, mrepMonthly: true }]) {
    const plan = planSupportReply("mrep kasa login ho ga", null, context);
    assert.equal(plan.replies[0].text, SUPPORT_COPY.romanUrdu.outsideScope);
    assert.equal(plan.replies[0].contact, true);
    assert.equal(plan.forceSavedReply, true);
    assert.doesNotMatch(plan.replies[0].text, /Call Status Report|Call Log Live/);
  }
});

test("two unresolved questions escalate politely with the configured admin contact", () => {
  const context = { language: "romanUrdu" as const, dailyReport: false, unresolvedCount: 0 };
  const first = planSupportReply("mujhe is wali cheez ka batao", null, context);
  assert.equal(first.replies[0].text, SUPPORT_COPY.romanUrdu.clarify);
  assert.equal(first.unresolvedCount, 1);
  const second = planSupportReply("mujhe is wali cheez ka batao", null, { ...context, unresolvedCount: first.unresolvedCount });
  assert.equal(second.replies[0].text, SUPPORT_COPY.romanUrdu.unresolved);
  assert.equal(second.replies[0].contact, true);
  assert.equal(second.unresolvedCount, 2);
  assert.equal(second.forceSavedReply, true);
  assert.equal(ADMIN_CONTACT.phone, "03702220009");
  assert.equal(ADMIN_CONTACT.email, "neutrophamra5@gmail.com");
  assert.equal(planSupportReply("thnx", null, { ...context, unresolvedCount: 1 }).unresolvedCount, 0);
});

test("isolated reporting words do not select and paste an unrelated guide", () => {
  for (const message of ["mrep", "download", "sample", "I just said the word report"]) {
    const plan = planSupportReply(message, null);
    assert.equal(plan.unresolvedCount, 1, message);
    assert.doesNotMatch(plan.replies[0].text, /1\.|Step 1|Call Log Live/, message);
  }
});

test("the screenshot's i fine reply is ordinary conversation even in monthly context", () => {
  const plan = planSupportReply("i fine", null, { language: "english", dailyReport: false, monthlyReport: true });
  assert.match(plan.replies[0].text, /doing well/);
  assert.equal(plan.unresolvedCount, 0);
  assert.doesNotMatch(plan.replies[0].text, /Which part|sample|Step 1/);
});

test("Roman Urdu thank-you reply is exactly the requested short message", () => {
  const plan = planSupportReply("shukria", null);
  assert.equal(plan.replies[0].text, "Khushi hui ke main madad kar saka.😊");
  assert.equal(plan.forceSavedReply, true);
});
