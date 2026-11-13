import { DAILY_REPORT_HELP, getDailyReportHelp, getDailyReportStepHelp, getSupportLanguage, type SupportLanguage } from "./daily-report-help.ts";
import { getSupportSmallTalk, isSupportGratitude, normalizeSupportText } from "./support-smalltalk.ts";
import { MONTHLY_REPORT_HELP, isMonthlyReportQuestion } from "./monthly-report-help.ts";
import { getReportComparison, isReportComparisonRequest, getComparisonClarification } from "./report-comparison-help.ts";
import { DOCTOR_COVERAGE_HELP, isDoctorCoverageQuestion } from "./doctor-coverage-help.ts";
import { PLANNED_SUMMARY_HELP, isPlannedSummaryQuestion } from "./planned-summary-help.ts";
import { AVG_PER_DAY_HELP, isAvgPerDayQuestion } from "./avg-per-day-help.ts";
import { MREP_MONTHLY_HELP, isMrepMonthlyQuestion } from "./mrep-monthly-help.ts";
import { getSupportReportSummary, wantsSupportSummary, wantsAllReportSummaries } from "./support-report-summary.ts";
import { getFocusedSupportHelp } from "./support-focused-help.ts";

export const ADMIN_CONTACT = { phone: "03702220009", phoneHref: "tel:+923702220009", email: "neutrophamra5@gmail.com" };

export const SUPPORT_COPY = {
  english: {
    question: "Did this answer your question? Are you satisfied with the guidance?",
    yes: "Yes, I'm satisfied", no: "No, I need more help",
    thanks: "You're welcome! Happy to help.",
    contact: "If the report has an issue or the steps are still unclear, please contact the admin using the phone number or email below. Thank you!",
    typing: "Support assistant is typing...",
    greeting: "Hello! How can I help you with your report?",
    simpleHelp: "I’m here to help with your reporting question. Please tell me which report or step you need help with: Daily Report, Monthly Report, Avg. Per Day, Doctor Coverage, or Planned Summary.",
    clarify: "Sorry, how can I help you?",
    outsideScope: "Sorry, I can guide you through report preparation, but I don't have approved instructions for account access or login. Please contact the admin using the details below.",
    unresolved: "Sorry, please contact the admin for further help.",
    moreHelp: "I'm sorry the answer wasn't clear enough. What do you need help with, and which part was confusing? Tell me the step or issue so I can explain it more clearly.",
    otherReport: "I don't have the approved help guide for that report yet. I can help with Daily Report, Monthly Report and Doctor Coverage, or share the admin's contact details.",
    matching: "For the Daily Report, check that Employee Code in the sample matches Emp. ID in the Call Log. Selfie filenames must include the employee code, and only image messages on the matching call-log date are counted. If the result still looks wrong, please contact the admin.",
    monthlyClarify: "Which part of the Monthly Report would you like help with: sample Excel, Daily Report PDF files, or generating, previewing and downloading the report?",
    coverageClarify: "Which Doctor Coverage step do you need help with: sample Excel, downloading from MREP, month/team/Territory filters, or generating and downloading the report?",
    plannedClarify: "Which Planned Summary step do you need help with: sample Excel, Call Log dates and Happened filter, uploading the Excel, or Generate/Preview/Download?",
    avgClarify: "Which Avg. Per Day step do you need help with: sample Excel, selecting the month in MREP, downloading/uploading Avg Report Excel, or Generate/Preview/Download?",
    mrepClarify: "Which MREP Monthly Report step do you need help with: sample Excel, Employee-wise Call Status Report, optional WhatsApp selfie exports, or Generate/Preview/Download?",
  },
  romanUrdu: {
    question: "Kya aap ko samajh aa gaya? Kya aap is rehnumai se satisfied hain?",
    yes: "Ji, main satisfied hoon", no: "Nahi, mujhe mazeed madad chahiye",
    thanks: "Khushi hui ke main madad kar saka.😊",
    contact: "Agar report mein koi issue aa raha hai ya steps samajh nahi aaye, to neeche diye gaye number ya email par admin se rabta karein. Shukriya!",
    typing: "Support assistant type kar raha hai...",
    greeting: "Assalam-o-Alaikum! Aap ko report ke kis step mein madad chahiye?",
    simpleHelp: "Main aap ki madad ke liye yahan hoon. Please bataein ke aap kis report ya step mein help chahte hain: Daily Report, Monthly Report, Avg. Per Day, Doctor Coverage, ya Planned Summary.",
    clarify: "Sorry, main aap ki kis tarah madad kar sakta hoon?",
    outsideScope: "Sorry, main report banane ke tareeqe mein madad kar sakta hoon, lekin account ya login ki approved instructions mere paas nahi hain. Neeche diye gaye number ya email par admin se rabta karein.",
    unresolved: "Sorry, mazeed madad ke liye admin se contact kar lein.",
    moreHelp: "Sorry, jawab poori tarah clear nahi hua. Aap ko kis cheez mein help chahiye aur kaunsa hissa samajh nahi aaya? Step ya masla bata dein, main usay aasaan tareeqe se samjhane ki koshish karta hoon.",
    otherReport: "Is report ki approved help guide abhi add nahi hui. Main Daily Report, Monthly Report aur Doctor Coverage mein madad kar sakta hoon, ya admin ka contact de sakta hoon.",
    matching: "Daily Report mein sample ka Employee Code, Call Log ki Emp. ID se match hona chahiye. Selfie file ke naam mein bhi employee code ho aur sirf call-log ki matching date ke image messages count hote hain. Result phir bhi ghalat ho to admin se rabta karein.",
    monthlyClarify: "Monthly Report mein kis cheez ke bare mein poochna chahte hain: sample Excel, Daily Report ki PDF files, ya generate, preview aur download karna?",
    coverageClarify: "Doctor Coverage ke kis step mein madad chahiye: sample Excel, MREP se download, month/team/Territory ke filters, ya report generate aur download karna?",
    plannedClarify: "Planned Summary ke kis step mein madad chahiye: sample Excel, Call Log ki dates aur Happened filter, Excel upload, ya Generate/Preview/Download?",
    avgClarify: "Avg. Per Day ke kis step mein madad chahiye: sample Excel, MREP mein month select karna, Avg Report Excel download/upload, ya Generate/Preview/Download?",
    mrepClarify: "MREP Monthly Report ke kis step mein madad chahiye: sample Excel, Employee-wise Call Status Report, optional WhatsApp selfie exports, ya Generate/Preview/Download?",
  },
  urdu: {
    question: "کیا آپ کو سمجھ آ گیا؟ کیا آپ اس رہنمائی سے مطمئن ہیں؟",
    yes: "جی ہاں، میں مطمئن ہوں", no: "نہیں، مجھے مزید مدد چاہیے",
    thanks: "آپ کا خیر مقدم ہے! مدد کرکے خوشی ہوئی۔",
    contact: "اگر رپورٹ میں کوئی مسئلہ آ رہا ہے یا طریقہ سمجھ نہیں آیا تو نیچے دیے گئے فون نمبر یا ای میل پر ایڈمن سے رابطہ کریں۔ شکریہ!",
    typing: "سپورٹ اسسٹنٹ ٹائپ کر رہا ہے…",
    greeting: "السلام علیکم! آپ کو رپورٹ کے کس مرحلے میں مدد چاہیے؟",
    simpleHelp: "میں آپ کی مدد کے لیے یہاں ہوں۔ براہِ کرم بتائیں کہ آپ کس رپورٹ یا مرحلے میں مدد چاہتے ہیں: ڈیلی رپورٹ، ماہانہ رپورٹ، Avg. Per Day، ڈاکٹر کورج، یا Planned Summary۔",
    clarify: "معذرت، میں آپ کی کس طرح مدد کر سکتا ہوں؟",
    outsideScope: "معذرت، میں رپورٹ بنانے کے طریقے میں مدد کر سکتا ہوں، لیکن اکاؤنٹ یا لاگ اِن کی منظور شدہ ہدایات میرے پاس نہیں ہیں۔ نیچے دیے گئے نمبر یا ای میل پر ایڈمن سے رابطہ کریں۔",
    unresolved: "معذرت، مزید مدد کے لیے ایڈمن سے رابطہ کر لیجیے۔",
    moreHelp: "معذرت، جواب پوری طرح واضح نہیں ہوا۔ آپ کو کس چیز میں مدد چاہیے اور کون سا حصہ سمجھ نہیں آیا؟ مرحلہ یا مسئلہ بتائیں تاکہ میں اسے آسان طریقے سے سمجھانے کی کوشش کروں۔",
    otherReport: "اس رپورٹ کی منظور شدہ مدد کی گائیڈ ابھی شامل نہیں ہوئی۔ میں ڈیلی رپورٹ، ماہانہ رپورٹ اور ڈاکٹر کوریج میں مدد کر سکتا ہوں یا ایڈمن کی رابطہ تفصیل دے سکتا ہوں۔",
    matching: "ڈیلی رپورٹ میں نمونہ فائل کا Employee Code، کال لاگ کی Emp. ID سے ملنا چاہیے۔ سیلفی فائل کے نام میں بھی ملازم کا کوڈ ہو۔ صرف کال لاگ کی متعلقہ تاریخ کے image پیغامات گنے جاتے ہیں۔ نتیجہ پھر بھی غلط ہو تو ایڈمن سے رابطہ کریں۔",
    monthlyClarify: "ماہانہ رپورٹ کے کس حصے میں مدد چاہیے: نمونہ Excel، ڈیلی رپورٹ کی PDF فائلیں، یا رپورٹ بنانا، Preview اور Download کرنا؟",
    coverageClarify: "ڈاکٹر کوریج کے کس مرحلے میں مدد چاہیے: نمونہ Excel، MREP سے ڈاؤن لوڈ، مہینہ، ٹیم اور Territory کے فلٹر، یا رپورٹ بنانا اور ڈاؤن لوڈ کرنا؟",
    plannedClarify: "پلانڈ سمری کے کس مرحلے میں مدد چاہیے: نمونہ Excel، کال لاگ کی تاریخیں اور Happened فلٹر، Excel اپ لوڈ، یا Generate، Preview اور Download؟",
    avgClarify: "Avg. Per Day کے کس مرحلے میں مدد چاہیے: نمونہ Excel، MREP میں مہینہ منتخب کرنا، Avg Report Excel ڈاؤن لوڈ اور اپ لوڈ، یا Generate، Preview اور Download؟",
    mrepClarify: "MREP Monthly Report کے کس مرحلے میں مدد چاہیے: نمونہ Excel، Employee-wise Call Status Report، اختیاری WhatsApp سیلفی ایکسپورٹس، یا Generate، Preview اور Download؟",
  },
};

export interface SupportReply {
  text: string;
  feedback?: SupportLanguage;
  contact?: boolean;
}

export function planSupportReply(message: string, feedbackLanguage: SupportLanguage | null, context?: { language: SupportLanguage; dailyReport: boolean; monthlyReport?: boolean; doctorCoverage?: boolean; plannedSummary?: boolean; avgPerDay?: boolean; mrepMonthly?: boolean; unresolvedCount?: number }): {
  replies: SupportReply[]; language: SupportLanguage; nextFeedback: SupportLanguage | null; dailyReport: boolean; monthlyReport: boolean; doctorCoverage: boolean; plannedSummary: boolean; avgPerDay: boolean; mrepMonthly: boolean; unresolvedCount: number; forceSavedReply: boolean;
} {
  message = normalizeSupportText(message);
  const language = getSupportLanguage(message, context?.language ?? feedbackLanguage ?? "english");
  const copy = SUPPORT_COPY[language];
  const normalized = normalizeSupportText(message);
  const fullHelp = getDailyReportHelp(message);
  const otherReport = /\b(monthly|avg|average per day|planned summary|doctor coverage)\b|منتھلی|ماہانہ|ڈاکٹر کوریج/i.test(message);
  const monthlyQuestion = isMonthlyReportQuestion(message);
  const coverageQuestion = isDoctorCoverageQuestion(message);
  const plannedQuestion = isPlannedSummaryQuestion(message)
    || Boolean(context?.plannedSummary && /\bmonthly\s+call\s*log\b/i.test(message));
  const avgQuestion = isAvgPerDayQuestion(message);
  const mrepQuestion = isMrepMonthlyQuestion(message);
  const mrepMonthly = mrepQuestion || (!otherReport && !monthlyQuestion && !coverageQuestion && !plannedQuestion && !avgQuestion && !fullHelp && Boolean(context?.mrepMonthly));
  const avgPerDay = !mrepMonthly && (avgQuestion || (!otherReport && !monthlyQuestion && !coverageQuestion && !plannedQuestion && !fullHelp && Boolean(context?.avgPerDay)));
  const plannedSummary = !mrepMonthly && !avgPerDay && (plannedQuestion || (!otherReport && !monthlyQuestion && !coverageQuestion && !fullHelp && Boolean(context?.plannedSummary)));
  const doctorCoverage = !mrepMonthly && !avgPerDay && !plannedSummary && (coverageQuestion || (!otherReport && !monthlyQuestion && !fullHelp && Boolean(context?.doctorCoverage)));
  const monthlyReport = !mrepMonthly && !avgPerDay && !plannedSummary && !doctorCoverage && (monthlyQuestion || (!otherReport && !fullHelp && Boolean(context?.monthlyReport)));
  const dailyReport = !mrepMonthly && !avgPerDay && !otherReport && !monthlyReport && !doctorCoverage && !plannedSummary && (Boolean(fullHelp) || Boolean(context?.dailyReport) || Boolean(feedbackLanguage && !context));
  const reply = (text: string, askFeedback = false, contact = false) => ({
    language, dailyReport, monthlyReport, doctorCoverage, plannedSummary, avgPerDay, mrepMonthly, unresolvedCount: 0, forceSavedReply: false, nextFeedback: askFeedback ? language : null,
    replies: askFeedback ? [{ text }, { text: copy.question, feedback: language }] : [{ text, ...(contact ? { contact: true } : {}) }],
  });
  const unresolved = (question = copy.clarify) => {
    const count = Math.min(2, (context?.unresolvedCount ?? 0) + 1);
    return { ...reply(count >= 2 ? copy.unresolved : question, false, count >= 2), unresolvedCount: count, forceSavedReply: count >= 2 };
  };
  const loginQuestion = /\b(log\s*in|login|sign\s*in|password|otp|account|subscription|payment|billing)\b|لاگ\s*(اِن|ان)|پاس\s*ورڈ|اکاؤنٹ/i.test(normalized);
  if (loginQuestion) return { ...reply(copy.outsideScope, false, true), forceSavedReply: true };
  const contactRequest = /\b(admin|administrator|phone|email)\b|\bcontact (you|admin|support)\b|ایڈمن|رابطہ|فون|ای میل/i.test(normalized);
  if (contactRequest) return reply(copy.contact, false, true);
  const comparison = getReportComparison(message, language);
  if (comparison) return { ...reply(comparison, true), forceSavedReply: true, dailyReport: false, monthlyReport: false, doctorCoverage: false, plannedSummary: false, avgPerDay: false, mrepMonthly: false };
  if (isReportComparisonRequest(message)) return unresolved(getComparisonClarification(language));
  if (wantsSupportSummary(message)) {
    const report = wantsAllReportSummaries(message) ? undefined
      : mrepMonthly ? "mrep" : avgPerDay ? "avg" : plannedSummary ? "planned" : doctorCoverage ? "doctor" : monthlyReport ? "monthly" : dailyReport ? "daily" : undefined;
    return { ...reply(getSupportReportSummary(language, report), true), forceSavedReply: true };
  }
  if (otherReport && !monthlyQuestion && !coverageQuestion && !plannedQuestion && !avgQuestion && !mrepQuestion) return unresolved();

  const simpleHelpRequest = /(i need help|need help|help me|can you help|can you support|i need support|need support|please help|help please|mujhe madad chahiye|mujhe help chahiye|madad chahiye|meri madad karo|mujhe madad karni hai|mujhe support chahiye|مجھے مدد چاہیے|مدد چاہیے|میری مدد کرو|کمک چاہیے|مساعدة)/i.test(message);

  const smallTalk = getSupportSmallTalk(normalized, language);
  if (smallTalk && !simpleHelpRequest) return { ...reply(isSupportGratitude(normalized) ? copy.thanks : smallTalk), forceSavedReply: isSupportGratitude(normalized) };

  const gratitude = isSupportGratitude(normalized);
  const negative = /\b(no|unsatisfied|unsatisfy|dissatisfied|dissatisfy|nahi|nhi|nai|nahin|confused|issue|problem|error)\b|un\s+satisf\w*|not\s+(satisf\w*|clear|helpful|underst\w*)|didn['’]?t (understand|get)|don['’]?t (understand|get)|نہیں|نہيں|مسئلہ|سمجھ نہیں/i.test(normalized);
  const positive = /\b(yes|satisfied|satisfy|understood|han|haan|ji|ok|okay|good)\b|samajh? (aa|a) (gaya|gya|gai)|جی ہاں|مطمئن|سمجھ آ گیا/i.test(normalized);
  const activeReport = mrepMonthly ? "mrep" : avgPerDay ? "avg" : plannedSummary ? "planned" : doctorCoverage ? "doctor" : monthlyReport ? "monthly" : dailyReport ? "daily" : undefined;
  const namesReport = Boolean(fullHelp) || monthlyQuestion || coverageQuestion || plannedQuestion || avgQuestion || mrepQuestion;
  // A new report question starts a full guide, even after an unresolved follow-up.
  if (activeReport && (negative || (!namesReport && (context?.unresolvedCount ?? 0) > 0))) {
    const focused = getFocusedSupportHelp(message, language, activeReport);
    if (focused) return reply(focused, true);
  }
  if (negative && (feedbackLanguage || activeReport || /\breport\b|رپورٹ/i.test(message))) return unresolved(copy.moreHelp);
  const asksForGuide = /\b(how|where|what|which|help|guide|steps?|create|generate|genrate|upload|download|preview|sample|template|files?|excel|exel|columns?|format|filter|select|kaise|kasa|kese|kasy|kia|kya|kis|kahan|batao|btao|bna|banani|banaye)\b|کیسے|کیا|کہاں|سیمپل|نمونہ|فائل|فلٹر|کالم/i.test(message)
    || /^(daily report|monthly report|doctor coverage|planned summary|avg\.? per day|mrep monthly report)[.!?\s]*$/i.test(normalized);
  if (mrepMonthly && (gratitude || (feedbackLanguage && positive))) return { ...reply(copy.thanks), forceSavedReply: true };
  const mrepWorkQuestion = /\b(sample|template|columns?|format|files?|excel|exel|report|mrep|month|employee|call|status|selfies?|whatsapp|extension|export|dates?|filter|generate|genrate|preview|download|upload|steps?|next|agla)\b|سیمپل|نمونہ|فائل|رپورٹ|مہینہ|سیلفی|واٹس|فلٹر|اگلا/i.test(message);
  if (asksForGuide && (mrepQuestion || (mrepMonthly && mrepWorkQuestion))) return reply(MREP_MONTHLY_HELP[language], true);
  if (mrepMonthly) return unresolved();
  if (avgPerDay && (gratitude || (feedbackLanguage && positive))) return { ...reply(copy.thanks), forceSavedReply: true };
  const avgWorkQuestion = /\b(sample|template|columns?|format|files?|excel|exel|report|mrep|month|filter|generate|genrate|preview|download|upload|steps?|next|agla)\b|سیمپل|نمونہ|فائل|رپورٹ|مہینہ|فلٹر|اگلا/i.test(message);
  if (asksForGuide && (avgQuestion || (avgPerDay && avgWorkQuestion))) return reply(AVG_PER_DAY_HELP[language], true);
  if (avgPerDay) return unresolved();
  if (plannedSummary && (gratitude || (feedbackLanguage && positive))) return { ...reply(copy.thanks), forceSavedReply: true };
  const plannedWorkQuestion = /\b(sample|template|columns?|format|files?|excel|exel|report|mrep|dates?|days?|month|filter|happened|generate|genrate|preview|download|steps?|next|agla|call\s*log)\b|سیمپل|نمونہ|فائل|رپورٹ|تاریخ|دن|فلٹر|کال لاگ|اگلا/i.test(message);
  if (asksForGuide && (plannedQuestion || (plannedSummary && plannedWorkQuestion))) return reply(PLANNED_SUMMARY_HELP[language], true);
  if (plannedSummary) return unresolved();
  if (doctorCoverage && (gratitude || (feedbackLanguage && positive))) return { ...reply(copy.thanks), forceSavedReply: true };
  const coverageWorkQuestion = /\b(sample|template|columns?|format|files?|excel|exel|report|mrep|month|team|territor(?:y|ies)|filter|generate|genrate|preview|download|steps?|next|agla)\b|سیمپل|نمونہ|فائل|رپورٹ|مہینہ|ٹیم|فلٹر|اگلا/i.test(message);
  if (asksForGuide && (coverageQuestion || (doctorCoverage && coverageWorkQuestion))) return reply(DOCTOR_COVERAGE_HELP[language], true);
  if (doctorCoverage) return unresolved();
  if (monthlyReport && (gratitude || (feedbackLanguage && positive))) return { ...reply(copy.thanks), forceSavedReply: true };
  const monthlyWorkQuestion = /\b(sample|template|columns?|format|files?|pdfs?|excel|exel|report|generate|genrate|preview|download|steps?|next|agla)\b|سیمپل|نمونہ|فائل|رپورٹ|پی ڈی ایف|کالم|اگلا/i.test(message);
  if (asksForGuide && (monthlyQuestion || (monthlyReport && monthlyWorkQuestion))) return reply(MONTHLY_REPORT_HELP[language], true);
  if (monthlyReport) return unresolved();
  const specificQuestion = /call\s*log|calllog|happen(?:ed|d)?|filter|selfi|slefi|whatsapp|extension|sample|sampl|template|column|format|preview|download|کال\s*لاگ|فلٹر|سیلفی|واٹس|سیمپل|نمونہ|پریویو|ڈاؤن/i.test(message);
  const step = getDailyReportStepHelp(message, language);
  if (step && dailyReport && (!fullHelp || specificQuestion)) return reply(step, true);
  if (dailyReport && /\b(files?|excel|exel|xlsx)\b|فائل|ایکسل/i.test(message) && !negative) return reply(DAILY_REPORT_HELP[language], true);
  if (negative && (feedbackLanguage || dailyReport || /\breport\b|رپورٹ/i.test(message))) return reply(copy.contact, false, true);
  const askingQuestion = /[?؟]|\b(how|which|what|where|kaise|kasa|kese|kasy|kis|kia|kya|batao|btao)\b|کیسے|کیا|کہاں/i.test(message);
  if (!askingQuestion && (gratitude || (feedbackLanguage && positive))) return { ...reply(copy.thanks), forceSavedReply: true };
  if (fullHelp && asksForGuide) return reply(DAILY_REPORT_HELP[language], true);
  if (gratitude || (feedbackLanguage && positive)) return { ...reply(copy.thanks), forceSavedReply: true };
  if (simpleHelpRequest) return unresolved(copy.simpleHelp);
  if (/\b(hello|hi|hey|salam|slaam|assalam|assalamualaikum|assalamalaikum)\b|السلام|سلام/i.test(normalized)) return reply(copy.greeting);
  if (dailyReport && /\b(code|id|match|missing|zero|count|wrong|galat|ghalat)\b|کوڈ|غلط|صفر/i.test(message)) return reply(copy.matching, true);
  return unresolved();
}
