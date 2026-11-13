import test from "node:test";
import assert from "node:assert/strict";
import { ADMIN_CONTACT, SUPPORT_COPY, planSupportReply } from "./support-feedback.ts";
import { MREP_MONTHLY_HELP } from "./mrep-monthly-help.ts";

test("Daily help is followed by a satisfaction question", () => {
  const plan = planSupportReply("How can I Genrate the Daily Report?", null);
  assert.equal(plan.replies.length, 2);
  assert.match(plan.replies[0].text, /Call Log Live/);
  assert.equal(plan.replies[1].text, SUPPORT_COPY.english.question);
  assert.equal(plan.replies[1].feedback, "english");
});

test("negative feedback first asks what needs clarification instead of sharing contacts", () => {
  const plan = planSupportReply("No, I don't understand the Daily Report", "english");
  assert.equal(plan.replies.length, 1);
  assert.equal(plan.replies[0].contact, undefined);
  assert.equal(plan.replies[0].text, SUPPORT_COPY.english.moreHelp);
  assert.equal(plan.unresolvedCount, 1);
  assert.equal(plan.nextFeedback, null);
  assert.equal(ADMIN_CONTACT.phone, "03702220009");
  assert.equal(ADMIN_CONTACT.email, "neutrophamra5@gmail.com");
});

test("satisfied customers receive a thank-you response", () => {
  assert.equal(planSupportReply("Yes, I'm satisfied", "english").replies[0].text, SUPPORT_COPY.english.thanks);
  assert.equal(planSupportReply("Ji, main satisfied hoon", "romanUrdu").replies[0].text, SUPPORT_COPY.romanUrdu.thanks);
  assert.equal(planSupportReply("جی ہاں، میں مطمئن ہوں", "urdu").replies[0].text, SUPPORT_COPY.urdu.thanks);
});

test("Urdu and Roman Urdu negative feedback asks what further help is needed", () => {
  assert.equal(planSupportReply("نہیں، مجھے سمجھ نہیں آئی", "urdu").replies[0].text, SUPPORT_COPY.urdu.moreHelp);
  assert.equal(planSupportReply("samaj nahi ai", "romanUrdu").replies[0].text, SUPPORT_COPY.romanUrdu.moreHelp);
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

test("standalone gratitude, punctuation and typing variants never fall through to report clarification", () => {
  for (const message of ["thnx", "THNX!", "thnxx", "thanksss", "t h n x", "thn\u200bx", "thankx 😊"]) {
    const plan = planSupportReply(message, null);
    assert.equal(plan.replies[0].text, SUPPORT_COPY.english.thanks, message);
    assert.equal(plan.replies.length, 1);
    assert.equal(plan.nextFeedback, null);
  }
});

test("normal Roman Urdu conversation answers the user's question without asking for a report step", () => {
  for (const message of ["kia hal hn", "kasa hn ap", "aap kaise hain?", "kese ho?"]) {
    const plan = planSupportReply(message, null);
    assert.equal(plan.language, "romanUrdu", message);
    assert.match(plan.replies[0].text, /Aap ki kis tarah madad/);
    assert.doesNotMatch(plan.replies[0].text, /kis step/);
  }
  assert.match(planSupportReply("ap kia kr rha hn", null).replies[0].text, /aap se baat/);
  assert.match(planSupportReply("ap kon ho", null).replies[0].text, /virtual reporting assistant/);
  assert.match(planSupportReply("main theek hoon", null).replies[0].text, /sun kar acha laga/);
});

test("small talk follows English and Urdu and does not swallow mixed reporting questions", () => {
  assert.match(planSupportReply("How are you?", null).replies[0].text, /How can I help you\?/);
  assert.match(planSupportReply("What are you doing?", null).replies[0].text, /chatting with you/);
  assert.match(planSupportReply("کیا حال ہے؟", null).replies[0].text, /کس طرح مدد/);
  assert.match(planSupportReply("آپ کیا کر رہے ہیں؟", null).replies[0].text, /آپ سے بات/);
  assert.match(planSupportReply("hi, how can I generate the Daily Report?", null).replies[0].text, /Call Log Live/);
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
  assert.equal(planSupportReply("How can I generate the MREP Monthly Report?", null).replies[0].text, MREP_MONTHLY_HELP.english);
});

test("screenshot comparison question explains the difference instead of showing the monthly guide", () => {
  const plan = planSupportReply("daily report or monthly report mn kia deference ha?", null);
  assert.equal(plan.language, "romanUrdu");
  assert.match(plan.replies[0].text, /farq/);
  assert.match(plan.replies[0].text, /Call Log Excel/);
  assert.match(plan.replies[0].text, /PDF files/);
  assert.doesNotMatch(plan.replies[0].text, /Step 1/);
  assert.equal(plan.dailyReport, false);
  assert.equal(plan.monthlyReport, false);
});

test("wellbeing follow-up in monthly context never restarts the report guide", () => {
  const monthly = { language: "romanUrdu" as const, dailyReport: false, monthlyReport: true };
  assert.match(planSupportReply("kia hal ha", null, monthly).replies[0].text, /Aap ki kis tarah madad/);
  for (const message of ["mn bi thk", "mein bhi theek hun", "main bhi thik hoon"]) {
    const plan = planSupportReply(message, null, monthly);
    assert.equal(plan.language, "romanUrdu");
    assert.match(plan.replies[0].text, /sun kar acha laga/);
    assert.equal(plan.replies.length, 1);
    assert.doesNotMatch(plan.replies[0].text, /Step 1/);
  }
});

test("monthly context alone cannot trigger the full guide for unrelated text", () => {
  const plan = planSupportReply("hmm acha suno", null, { language: "romanUrdu", dailyReport: false, monthlyReport: true });
  assert.equal(plan.replies[0].text, SUPPORT_COPY.romanUrdu.clarify);
  assert.doesNotMatch(plan.replies[0].text, /Step 1/);
});

test("comparison replies follow English and Urdu questions", () => {
  assert.match(planSupportReply("What is the difference between Daily and Monthly Report?", null).replies[0].text, /different source files/);
  assert.match(planSupportReply("ڈیلی اور ماہانہ رپورٹ میں کیا فرق ہے؟", null).replies[0].text, /فرق/);
});
