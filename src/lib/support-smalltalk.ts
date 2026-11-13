import type { SupportLanguage } from "./daily-report-help.ts";

const CHAT_COPY = {
  english: {
    thanks: "You're welcome! Happy to help.",
    greeting: "Hello! How can I help you?",
    wellbeing: "I'm here and ready to help! How can I help you?",
    activity: "I'm here chatting with you and helping with reporting questions.",
    identity: "I'm your virtual reporting assistant. You can talk to me here for help.",
    capabilities: "I can explain the Daily Report steps, required Excel files and Call Log download, or share the admin's contact details when you need more help.",
    farewell: "Take care! You can come back whenever you need help.",
    acknowledgement: "Sure, I'm here whenever you need me.",
    glad: "Glad to hear you're doing well! How can I help you?",
  },
  romanUrdu: {
    thanks: "Khushi hui ke main madad kar saka.😊",
    greeting: "Assalam-o-Alaikum! Aap ki kis tarah madad kar sakta hoon?",
    wellbeing: "Main yahan hoon aur aap ki madad ke liye tayyar hoon! Aap ki kis tarah madad kar sakta hoon?",
    activity: "Main aap se baat kar raha hoon aur report ke sawalon mein madad ke liye yahan hoon.",
    identity: "Main aap ka virtual reporting assistant hoon. Aap yahan mujh se madad le sakte hain.",
    capabilities: "Main Daily Report ke steps, zaroori Excel files aur Call Log download karne ka tareeqa bata sakta hoon. Zaroorat ho to admin ka contact bhi de sakta hoon.",
    farewell: "Allah Hafiz! Jab bhi madad chahiye ho, wapas aa sakte hain.",
    acknowledgement: "Theek hai, main yahin hoon. Jab chahiye pooch lein.",
    glad: "Yeh sun kar acha laga ke aap theek hain! Aap ki kis tarah madad kar sakta hoon?",
  },
  urdu: {
    thanks: "آپ کا خیر مقدم ہے! مدد کرکے خوشی ہوئی۔",
    greeting: "السلام علیکم! میں آپ کی کس طرح مدد کر سکتا ہوں؟",
    wellbeing: "میں یہاں آپ کی مدد کے لیے تیار ہوں! میں آپ کی کس طرح مدد کر سکتا ہوں؟",
    activity: "میں آپ سے بات کر رہا ہوں اور رپورٹ سے متعلق سوالات میں مدد کے لیے یہاں ہوں۔",
    identity: "میں آپ کا ورچوئل رپورٹنگ اسسٹنٹ ہوں۔ آپ یہاں مجھ سے مدد لے سکتے ہیں۔",
    capabilities: "میں ڈیلی رپورٹ کے مراحل، ضروری Excel فائلوں اور کال لاگ ڈاؤن لوڈ کرنے کا طریقہ بتا سکتا ہوں۔ ضرورت ہو تو ایڈمن کی رابطہ تفصیل بھی دے سکتا ہوں۔",
    farewell: "اللہ حافظ! جب بھی مدد چاہیے ہو، دوبارہ آ سکتے ہیں۔",
    acknowledgement: "ٹھیک ہے، میں یہیں ہوں۔ جب ضرورت ہو پوچھ لیجیے۔",
    glad: "یہ سن کر اچھا لگا کہ آپ خیریت سے ہیں! میں آپ کی کس طرح مدد کر سکتا ہوں؟",
  },
};

export function normalizeSupportText(message: string): string {
  const corrections: Record<string, string> = {
    genrate: "generate", genarate: "generate", generat: "generate", genrette: "generate",
    uplod: "upload", uplaod: "upload", downlod: "download", donwload: "download", dowload: "download",
    preveiw: "preview", smaple: "sample", sampl: "sample", samplye: "sample",
    rpeort: "report", reprot: "report", rport: "report", exel: "excel",
  };
  return message.normalize("NFKC").replace(/[\u200b-\u200d\ufeff]/g, "").toLowerCase()
    .replace(/([a-z])\1{2,}/g, "$1$1")
    .replace(/\bsu+m+(?:a|e)?r+y\b/g, "summary")
    .replace(/\b[a-z]+\b/g, (word) => corrections[word] ?? word)
    .replace(/\s+/g, " ").trim();
}

export function isSupportGratitude(message: string): boolean {
  const text = normalizeSupportText(message);
  return /\b(thank+(?:s+|x+)?|thanx+|thnx+|thnks+|thnku|thank\s*(you|u)|thankyou|shukriya+|shukria+|shukrya)\b|شکریہ|مہربانی/i.test(text)
    || /^(thnx+|thnks+|thankx+|thanks+|thankyou|thanku)$/.test(text.replace(/[^a-z]/g, ""));
}

export function getSupportSmallTalk(message: string, language: SupportLanguage): string | null {
  const text = normalizeSupportText(message);
  // A greeting must not hide a reporting question in the same message.
  if (/\b(report|rport|daily|monthly|avg|sample|files?|excel|exel|call|log|selfie|download|preview|generate|genrate|admin|email|phone|problem|error|issue)\b|رپورٹ|فائل|کال|سیلفی|ایڈمن|مسئلہ/i.test(text)) return null;
  const copy = CHAT_COPY[language];
  if (isSupportGratitude(text)) return copy.thanks;
  if (/what can (you|u) do|what do (you|u) do|(?:kya|kia).*(?:kar|kr).*sakt|کیا.*کر.*سکت/i.test(text)) return copy.capabilities;
  if (/what (?:are|r) (?:you|u) (?:doing|up to)|(?:kya|kia)\s+(?:kar|kr)\s+(?:rahe|rahi|rha|rhi|rhe|rhy|rahy)|کیا کر/i.test(text)) return copy.activity;
  if (/who (?:are|r) (?:you|u)|(?:ap|aap|tum)\s+(?:kon|kaun|koun)|(?:ap|aap)\s+ka\s+naam|آپ کون|آپ کا نام/i.test(text)) return copy.identity;
  if (/how (?:are|r) (?:you|u)|how(?:['’]s| is) it going|(?:kya|kia)\s+(?:haal|hal)|(?:kesa|kaisa|kese|kaise|kesay|kaisay|kasy|kaisy|kasa|ksa)\s+(?:(?:ap|aap|tum)\s+)?(?:ho|hain|hn|hen|hai)|(?:ap|aap)\s+(?:kesa|kaisa|kese|kaise|ka\s+(?:hal|haal))|کیا حال|آپ کیسے|کیسے ہیں/i.test(text)) return copy.wellbeing;
  if (/\b(?:main|mn|mein)\s+(?:(?:bhi|bi|bilkul)\s+)*(?:theek|thik|thk|acha)\b|\bi(?:['’]m|\s+am)?\s+(?:(?:also|too)\s+)?(?:fine|good|well)\b|میں (?:بھی )?ٹھیک|میں خیریت/i.test(text)) return copy.glad;
  if (/\b(bye|goodbye|allah hafiz|allahhafiz|khuda hafiz)\b|اللہ حافظ|خدا حافظ/i.test(text)) return copy.farewell;
  if (/^(?:ok(?:ay)?|acha|achha|theek|thik)[.!\s]*$/i.test(text)) return copy.acknowledgement;
  if (/\b(he*l+o+|hi+|hey+|salam|slaam|salaam|assalam|assalamualaikum|assalamalaikum|assalamu alaikum)\b|السلام|سلام/i.test(text)) return copy.greeting;
  return null;
}
