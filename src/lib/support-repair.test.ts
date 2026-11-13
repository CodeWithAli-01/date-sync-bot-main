import test from "node:test";
import assert from "node:assert/strict";
import { planSupportReply, SUPPORT_COPY } from "./support-feedback.ts";
import { normalizeSupportText } from "./support-smalltalk.ts";
import { DAILY_REPORT_HELP } from "./daily-report-help.ts";
import { MONTHLY_REPORT_HELP } from "./monthly-report-help.ts";
import { DOCTOR_COVERAGE_HELP } from "./doctor-coverage-help.ts";
import { PLANNED_SUMMARY_HELP } from "./planned-summary-help.ts";
import { AVG_PER_DAY_HELP } from "./avg-per-day-help.ts";
import { MREP_MONTHLY_HELP } from "./mrep-monthly-help.ts";

test("unclear messages ask how to help without guessing a report step", () => {
  for (const language of ["english", "romanUrdu", "urdu"] as const) {
    for (const flag of ["dailyReport", "monthlyReport", "doctorCoverage", "plannedSummary", "avgPerDay", "mrepMonthly"]) {
      for (const message of ["g", "q", "???", "asdfgh", "hmm acha suno"]) {
        const plan = planSupportReply(message, null, { language, dailyReport: false, [flag]: true });
        assert.equal(plan.replies[0].text, SUPPORT_COPY[language].clarify, `${message}, ${flag}, ${language}`);
        assert.equal(plan.replies.length, 1);
        assert.equal(plan.nextFeedback, null);
        assert.equal(plan.replies[0].contact, undefined);
      }
    }
  }
  assert.equal(planSupportReply("g", null).replies[0].text, "Sorry, how can I help you?");
});

test("short greeting spellings stay greetings in every report context", () => {
  const greetings = ["hi", "hlo", "Hello!", "hii", "hiii 😊", "hellooo", "helo", "hey", "heyy", "h\u200blo"];
  const contexts = ["dailyReport", "monthlyReport", "doctorCoverage", "plannedSummary", "avgPerDay", "mrepMonthly"];
  const copies = { english: "Hello! How can I help you?", romanUrdu: "Assalam-o-Alaikum! Aap ki kis tarah madad kar sakta hoon?", urdu: "السلام علیکم! میں آپ کی کس طرح مدد کر سکتا ہوں؟" } as const;
  for (const language of ["english", "romanUrdu", "urdu"] as const) {
    for (const flag of contexts) {
      for (const message of greetings) {
        for (const unresolvedCount of [0, 1, 2]) {
          const plan = planSupportReply(message, language, { language, dailyReport: false, [flag]: true, unresolvedCount });
          assert.equal(plan.replies[0].text, copies[language], `${message}, ${flag}, ${language}, ${unresolvedCount}`);
          assert.equal(plan.replies.length, 1);
          assert.equal(plan.nextFeedback, null);
          assert.equal(plan.unresolvedCount, 0);
        }
      }
    }
  }
  assert.match(planSupportReply("hlo, how can I generate the Avg. Per Day report?", null).replies[0].text, /Avg Report Excel/);
});

test("every report quick question returns its full guide after an unresolved conversation", () => {
  const reports = [
    { name: "Daily Report", flag: "dailyReport", guide: DAILY_REPORT_HELP },
    { name: "Monthly Report", flag: "monthlyReport", guide: MONTHLY_REPORT_HELP },
    { name: "Doctor Coverage", flag: "doctorCoverage", guide: DOCTOR_COVERAGE_HELP },
    { name: "Planned Summary", flag: "plannedSummary", guide: PLANNED_SUMMARY_HELP },
    { name: "MREP Monthly Report", flag: "mrepMonthly", guide: MREP_MONTHLY_HELP },
    { name: "Avg. Per Day", flag: "avgPerDay", guide: AVG_PER_DAY_HELP },
  ];
  for (const previous of reports) {
    for (const report of reports) {
      const message = `How can I Genrate the ${report.name}${/report$/i.test(report.name) ? "" : " rport"}?`;
      for (const unresolvedCount of [0, 1, 2]) {
        const plan = planSupportReply(message, "english", {
          language: "english", dailyReport: false, [previous.flag]: true, unresolvedCount,
        });
        const label = `${previous.name} -> ${report.name}, unresolved=${unresolvedCount}`;
        assert.equal(plan.replies[0].text, report.guide.english, label);
        assert.equal(plan.unresolvedCount, 0, label);
        assert.equal(plan.replies[0].contact, undefined, label);
        assert.equal(plan.nextFeedback, "english", label);
      }
    }
  }
});

test("No feedback leads to a clarification and then a focused Daily Report explanation", () => {
  const context = { language: "romanUrdu" as const, dailyReport: true };
  const no = planSupportReply("Nahi, mujhe mazeed madad chahiye", "romanUrdu", context);
  assert.equal(no.replies[0].text, SUPPORT_COPY.romanUrdu.moreHelp);
  assert.equal(no.replies[0].contact, undefined);
  const detail = planSupportReply("smaple file ke columns samajh nahi aye", null, { ...context, unresolvedCount: no.unresolvedCount });
  assert.match(detail.replies[0].text, /Employee Code/);
  assert.match(detail.replies[0].text, /Planned/);
  assert.doesNotMatch(detail.replies[0].text, /Call Log Live/);
  assert.equal(detail.replies[0].contact, undefined);
});

test("doctor and MREP clarification reuses the relevant approved steps", () => {
  const doctor = planSupportReply("territory select nahi samajh aya", null, { language: "romanUrdu", dailyReport: false, doctorCoverage: true, unresolvedCount: 1 });
  assert.match(doctor.replies[0].text, /Select All/);
  assert.equal(doctor.replies[0].contact, undefined);
  const mrep = planSupportReply("selfie extension ki file kaise deni hai", null, { language: "romanUrdu", dailyReport: false, mrepMonthly: true, unresolvedCount: 1 });
  assert.match(mrep.replies[0].text, /WhatsApp Selfie Export Files/);
  assert.doesNotMatch(mrep.replies[0].text, /Employee-wise Call Status Report/);
});

test("unresolved feedback still escalates after the second attempt and explicit contact remains available", () => {
  const context = { language: "english" as const, dailyReport: false, monthlyReport: true };
  const first = planSupportReply("No, I need more help", "english", context);
  const second = planSupportReply("No, I still don't understand", null, { ...context, unresolvedCount: first.unresolvedCount });
  assert.equal(first.replies[0].contact, undefined);
  assert.equal(second.replies[0].contact, true);
  assert.equal(second.replies[0].text, SUPPORT_COPY.english.unresolved);
  assert.equal(planSupportReply("contact admin please", null).replies[0].contact, true);
});

test("common summary spellings are understood without changing displayed customer text", () => {
  for (const word of ["suumery", "sumary", "summery", "summry", "suummary"]) {
    const plan = planSupportReply(`daily report ki ${word} batao`, null);
    assert.match(plan.replies[0].text, /Call Log Excel/);
    assert.doesNotMatch(plan.replies[0].text, /1\. Daily Report ki sample Excel upload/);
    assert.equal(plan.unresolvedCount, 0, word);
  }
  assert.equal(normalizeSupportText("genrate smaple exel uplaod preveiw"), "generate sample excel upload preview");
  assert.equal(normalizeSupportText("summer holidays"), "summer holidays");
});
