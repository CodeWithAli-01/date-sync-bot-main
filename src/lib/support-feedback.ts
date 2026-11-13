import { DAILY_REPORT_HELP, getDailyReportHelp, getDailyReportStepHelp, getSupportLanguage, type SupportLanguage } from "./daily-report-help.ts";

export const ADMIN_CONTACT = { phone: "03702220009", phoneHref: "tel:+923702220009", email: "neutrophamra5@gmail.com" };

export const SUPPORT_COPY = {
  english: {
    question: "Did this answer your question? Are you satisfied with the guidance?",
    yes: "Yes, I'm satisfied", no: "No, I need more help",
    thanks: "You're welcome! I'm glad the guidance helped. Thank you for contacting Help & Support.",
    contact: "If the report has an issue or the steps are still unclear, please contact the admin using the phone number or email below. Thank you!",
    typing: "Support assistant is typing...",
    greeting: "Hello! How can I help you with your report?",
    simpleHelp: "I’m here to help with your reporting question. Please tell me which report or step you need help with: Daily Report, Monthly Report, Avg. Per Day, Doctor Coverage, or Planned Summary.",
    clarify: "Which part of the Daily Report do you need help with: the sample Excel, Call Log download, selfie Excel, or generating and downloading the report? Please tell me a little more.",
    otherReport: "I don't have the approved help guide for that report yet. You can ask me about the Daily Report or request the admin's contact details.",
    matching: "For the Daily Report, check that Employee Code in the sample matches Emp. ID in the Call Log. Selfie filenames must include the employee code, and only image messages on the matching call-log date are counted. If the result still looks wrong, please contact the admin.",
  },
  romanUrdu: {
    question: "Kya aap ko samajh aa gaya? Kya aap is rehnumai se satisfied hain?",
    yes: "Ji, main satisfied hoon", no: "Nahi, mujhe mazeed madad chahiye",
    thanks: "Shukriya! Khushi hui ke aap ko samajh aa gaya. Help & Support se rabta karne ka shukriya!",
    contact: "Agar report mein koi issue aa raha hai ya steps samajh nahi aaye, to neeche diye gaye number ya email par admin se rabta karein. Shukriya!",
    typing: "Support assistant type kar raha hai...",
    greeting: "Assalam-o-Alaikum! Aap ko report ke kis step mein madad chahiye?",
    simpleHelp: "Main aap ki madad ke liye yahan hoon. Please bataein ke aap kis report ya step mein help chahte hain: Daily Report, Monthly Report, Avg. Per Day, Doctor Coverage, ya Planned Summary.",
    clarify: "Aap ko Daily Report ke kis step mein madad chahiye: sample Excel, Call Log download, selfie Excel, ya report generate/download karna? Thora sa mazeed bata dein.",
    otherReport: "Is report ki approved help guide abhi add nahi hui. Aap Daily Report ka sawal pooch sakte hain ya admin ka contact mang sakte hain.",
    matching: "Daily Report mein sample ka Employee Code, Call Log ki Emp. ID se match hona chahiye. Selfie file ke naam mein bhi employee code ho aur sirf call-log ki matching date ke image messages count hote hain. Result phir bhi ghalat ho to admin se rabta karein.",
  },
  urdu: {
    question: "کیا آپ کو سمجھ آ گیا؟ کیا آپ اس رہنمائی سے مطمئن ہیں؟",
    yes: "جی ہاں، میں مطمئن ہوں", no: "نہیں، مجھے مزید مدد چاہیے",
    thanks: "شکریہ! خوشی ہوئی کہ آپ کو سمجھ آ گیا۔ ہیلپ اینڈ سپورٹ سے رابطہ کرنے کا شکریہ!",
    contact: "اگر رپورٹ میں کوئی مسئلہ آ رہا ہے یا طریقہ سمجھ نہیں آیا تو نیچے دیے گئے فون نمبر یا ای میل پر ایڈمن سے رابطہ کریں۔ شکریہ!",
    typing: "سپورٹ اسسٹنٹ ٹائپ کر رہا ہے…",
    greeting: "السلام علیکم! آپ کو رپورٹ کے کس مرحلے میں مدد چاہیے؟",
    simpleHelp: "میں آپ کی مدد کے لیے یہاں ہوں۔ براہِ کرم بتائیں کہ آپ کس رپورٹ یا مرحلے میں مدد چاہتے ہیں: ڈیلی رپورٹ، ماہانہ رپورٹ، Avg. Per Day، ڈاکٹر کورج، یا Planned Summary۔",
    clarify: "آپ کو ڈیلی رپورٹ کے کس مرحلے میں مدد چاہیے: نمونہ Excel، کال لاگ ڈاؤن لوڈ، سیلفی Excel، یا رپورٹ بنانا اور ڈاؤن لوڈ کرنا؟ براہِ کرم تھوڑی مزید تفصیل بتائیں۔",
    otherReport: "اس رپورٹ کی منظور شدہ مدد کی گائیڈ ابھی شامل نہیں ہوئی۔ آپ ڈیلی رپورٹ کے بارے میں پوچھ سکتے ہیں یا ایڈمن کی رابطہ تفصیل مانگ سکتے ہیں۔",
    matching: "ڈیلی رپورٹ میں نمونہ فائل کا Employee Code، کال لاگ کی Emp. ID سے ملنا چاہیے۔ سیلفی فائل کے نام میں بھی ملازم کا کوڈ ہو۔ صرف کال لاگ کی متعلقہ تاریخ کے image پیغامات گنے جاتے ہیں۔ نتیجہ پھر بھی غلط ہو تو ایڈمن سے رابطہ کریں۔",
  },
};

export interface SupportReply {
  text: string;
  feedback?: SupportLanguage;
  contact?: boolean;
}

export function planSupportReply(message: string, feedbackLanguage: SupportLanguage | null, context?: { language: SupportLanguage; dailyReport: boolean }): {
  replies: SupportReply[]; language: SupportLanguage; nextFeedback: SupportLanguage | null; dailyReport: boolean;
} {
  const language = getSupportLanguage(message, context?.language ?? feedbackLanguage ?? "english");
  const copy = SUPPORT_COPY[language];
  const normalized = message.toLowerCase().trim();
  const fullHelp = getDailyReportHelp(message);
  const otherReport = /\b(monthly|avg|average per day|planned summary|doctor coverage)\b|منتھلی|ماہانہ|ڈاکٹر کوریج/i.test(message);
  const dailyReport = !otherReport && (Boolean(fullHelp) || Boolean(context?.dailyReport) || Boolean(feedbackLanguage));
  const reply = (text: string, askFeedback = false, contact = false) => ({
    language, dailyReport, nextFeedback: askFeedback ? language : null,
    replies: askFeedback ? [{ text }, { text: copy.question, feedback: language }] : [{ text, ...(contact ? { contact: true } : {}) }],
  });
  const contactRequest = /\b(admin|administrator|phone|email)\b|\bcontact (you|admin|support)\b|ایڈمن|رابطہ|فون|ای میل/i.test(normalized);
  if (contactRequest) return reply(copy.contact, false, true);
  if (otherReport) return reply(copy.otherReport);

  const simpleHelpRequest = /(i need help|need help|help me|can you help|can you support|i need support|need support|please help|help please|mujhe madad chahiye|mujhe help chahiye|madad chahiye|meri madad karo|mujhe madad karni hai|mujhe support chahiye|مجھے مدد چاہیے|مدد چاہیے|میری مدد کرو|کمک چاہیے|مساعدة)/i.test(message);

  const gratitude = /\b(thanks*|thankx|thanx|thnx|thnks|thnku|thank\s*(you|u)|thankyou|shukriya+|shukria+|shukrya)\b|شکریہ|مہربانی/i.test(normalized);
  const negative = /\b(no|unsatisfied|dissatisfied|nahi|nhi|nai|nahin|confused|issue|problem|error)\b|not\s+(satisf\w*|clear|helpful|underst\w*)|didn['’]?t understand|don['’]?t understand|نہیں|نہيں|مسئلہ|سمجھ نہیں/i.test(normalized);
  const positive = /\b(yes|satisfied|satisfy|understood|han|haan|ji|ok|okay|good)\b|samajh? (aa|a) (gaya|gya|gai)|جی ہاں|مطمئن|سمجھ آ گیا/i.test(normalized);
  const specificQuestion = /call\s*log|calllog|happen(?:ed|d)?|filter|selfi|slefi|whatsapp|extension|sample|sampl|template|column|format|preview|download|کال\s*لاگ|فلٹر|سیلفی|واٹس|سیمپل|نمونہ|پریویو|ڈاؤن/i.test(message);
  const step = getDailyReportStepHelp(message, language);
  if (step && (!fullHelp || specificQuestion)) return reply(step, true);
  if (dailyReport && /\b(files?|excel|exel|xlsx)\b|فائل|ایکسل/i.test(message) && !negative) return reply(DAILY_REPORT_HELP[language], true);
  if (negative && (dailyReport || /\breport\b|رپورٹ/i.test(message))) return reply(copy.contact, false, true);
  const askingQuestion = /[?؟]|\b(how|which|what|where|kaise|kasa|kese|kasy|kis|kia|kya|batao|btao)\b|کیسے|کیا|کہاں/i.test(message);
  if (!askingQuestion && (gratitude || (feedbackLanguage && positive))) return reply(copy.thanks);
  if (fullHelp) return reply(DAILY_REPORT_HELP[language], true);
  if (gratitude || (feedbackLanguage && positive)) return reply(copy.thanks);
  if (simpleHelpRequest) return reply(copy.simpleHelp);
  if (/\b(hello|hi|hey|salam|slaam|assalam|assalamualaikum|assalamalaikum)\b|السلام|سلام/i.test(normalized)) return reply(copy.greeting);
  if (dailyReport && /\b(code|id|match|missing|zero|count|wrong|galat|ghalat)\b|کوڈ|غلط|صفر/i.test(message)) return reply(copy.matching, true);
  return reply(copy.clarify);
}
