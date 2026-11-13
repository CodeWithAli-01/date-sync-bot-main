import test from "node:test";
import assert from "node:assert/strict";
import { ADMIN_CONTACT, SUPPORT_COPY, planSupportReply } from "./support-feedback.ts";

test("Daily help is followed by a satisfaction question", () => {
  const plan = planSupportReply("How can I Genrate the Daily Report?", null);
  assert.equal(plan.replies.length, 2);
  assert.match(plan.replies[0].text, /Call Log Live/);
  assert.equal(plan.replies[1].text, SUPPORT_COPY.english.question);
  assert.equal(plan.replies[1].feedback, "english");
});

test("negative feedback receives admin contacts, even if it mentions Daily Report", () => {
  const plan = planSupportReply("No, I don't understand the Daily Report", "english");
  assert.equal(plan.replies.length, 1);
  assert.equal(plan.replies[0].contact, true);
  assert.equal(plan.nextFeedback, null);
  assert.equal(ADMIN_CONTACT.phone, "03702220009");
  assert.equal(ADMIN_CONTACT.email, "neutrophamra5@gmail.com");
});

test("satisfied customers receive a thank-you response", () => {
  assert.equal(planSupportReply("Yes, I'm satisfied", "english").replies[0].text, SUPPORT_COPY.english.thanks);
  assert.equal(planSupportReply("Ji, main satisfied hoon", "romanUrdu").replies[0].text, SUPPORT_COPY.romanUrdu.thanks);
  assert.equal(planSupportReply("جی ہاں، میں مطمئن ہوں", "urdu").replies[0].text, SUPPORT_COPY.urdu.thanks);
});

test("Urdu and Roman Urdu feedback gets the corresponding contact response", () => {
  assert.equal(planSupportReply("نہیں، مجھے سمجھ نہیں آئی", "urdu").replies[0].text, SUPPORT_COPY.urdu.contact);
  assert.equal(planSupportReply("samaj nahi ai", "romanUrdu").replies[0].text, SUPPORT_COPY.romanUrdu.contact);
});

test("requesting admin contact works without previous feedback", () => {
  assert.equal(planSupportReply("How can I contact admin?", null).replies[0].contact, true);
  assert.equal(planSupportReply("How do I mark contact point in the daily report?", null).replies[0].contact, undefined);
});

test("thanks spelling variants receive replies even after feedback has finished", () => {
  for (const message of ["thankx", "thnx", "thanx", "thnks", "thanks", "thank u"]) {
    const plan = planSupportReply(message, null, { language: "romanUrdu", dailyReport: true });
    assert.equal(plan.replies[0].text, SUPPORT_COPY.romanUrdu.thanks, message);
  }
  assert.equal(planSupportReply("shukria", null).replies[0].text, SUPPORT_COPY.romanUrdu.thanks);
  assert.equal(planSupportReply("شکریہ", null).replies[0].text, SUPPORT_COPY.urdu.thanks);
});

test("Daily Report follow-up questions get the relevant step without repeating the whole guide", () => {
  const context = { language: "romanUrdu" as const, dailyReport: true };
  const log = planSupportReply("call log kahan se download krni hai", null, context);
  assert.match(log.replies[0].text, /Call Log Live/);
  assert.doesNotMatch(log.replies[0].text, /1\. Daily Report/);
  const sample = planSupportReply("sample mn kon se columns chahiye", null, context);
  assert.match(sample.replies[0].text, /Employee Code/);
  assert.match(sample.replies[0].text, /Planned/);
  assert.doesNotMatch(sample.replies[0].text, /Call Log Live/);
  assert.match(planSupportReply("selfie wali file kaise deni hai", null, context).replies[0].text, /WhatsApp extension/);
});

test("conversation language survives short messages and completed satisfaction replies", () => {
  const context = { language: "urdu" as const, dailyReport: true };
  assert.equal(planSupportReply("thankx", null, context).language, "urdu");
  assert.match(planSupportReply("call log", null, context).replies[0].text, /لاگ/);
  assert.equal(planSupportReply("yes satisfied with Daily Report", "english").replies[0].text, SUPPORT_COPY.english.thanks);
});

test("simple help requests receive a direct support reply rather than a generic clarification", () => {
  assert.equal(planSupportReply("I need help", null).replies[0].text, SUPPORT_COPY.english.simpleHelp);
  assert.equal(planSupportReply("mujhe madad chahiye", null).replies[0].text, SUPPORT_COPY.romanUrdu.simpleHelp);
  assert.equal(planSupportReply("مجھے مدد چاہیے", null).replies[0].text, SUPPORT_COPY.urdu.simpleHelp);
});

test("unknown questions ask for clarification instead of silently returning no reply", () => {
  assert.equal(planSupportReply("mujhe thori help chahiye", null).replies[0].text, SUPPORT_COPY.romanUrdu.clarify);
  assert.equal(planSupportReply("How can I generate the Avg. Per Day report?", null).replies[0].text, SUPPORT_COPY.english.otherReport);
});
