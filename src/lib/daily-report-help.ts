/** Customer-help content only. This does not change report generation. */
export const DAILY_REPORT_HELP = {
  english: `To generate your Daily Report:

1. Upload your Daily Report sample Excel. Include the employees you want in the report, with Employee Code and Name filled in. Keep these report columns: Planned, Unplanned, Mor (Morning), Eve (Evening), Total and Cp. You may also include Selfies, Region, City and Designation.

2. Download and upload the Call Log Excel:
• Sign in to MREP → Reports → Call Reports → Call Log Live.
• Select the date you want to report on, then select Filter.
• Once the results load, scroll to the right and set the Happened filter to Happened.
• When the filtered report is ready, download it as Excel and upload it under Call Log Excel.

3. Upload the WhatsApp selfie Excel file(s) you downloaded using your WhatsApp extension. The filename should include the employee code so the selfies can be matched to the correct employee and call-log date.

4. Select Generate Report and wait for processing to finish. You can open Preview to check the report before selecting Download.`,
  romanUrdu: `Daily Report banane ke liye:

1. Daily Report ki sample Excel upload karein. Jin employees ki report chahiye, unke Employee Code aur Name bhare hue hon. Sample mein Planned, Unplanned, Mor (Morning), Eve (Evening), Total aur Cp ke columns rakhein. Selfies, Region, City aur Designation bhi rakh sakte hain.

2. Call Log Excel download karke upload karein:
• MREP mein login karein → Reports → Call Reports → Call Log Live kholein.
• Jis date ki report banani hai, woh date select karke Filter karein.
• Results aa jayen to right taraf scroll karein aur Happened ke filter mein Happened select karein.
• Filtered report tayyar ho jaye to Excel download karein aur Call Log Excel mein upload karein.

3. WhatsApp extension se download ki hui selfie Excel files upload karein. File ke naam mein employee code hona chahiye, taake sahi employee aur call-log date ki selfies match ho sakein.

4. Generate Report dabayein aur processing mukammal hone ka intezar karein. Download se pehle Preview mein report check bhi kar sakte hain.`,
  urdu: `ڈیلی رپورٹ بنانے کے لیے:

1. ڈیلی رپورٹ کی نمونہ Excel فائل اپ لوڈ کریں۔ جن ملازمین کی رپورٹ چاہیے، ان کے Employee Code اور Name درج ہوں۔ فائل میں Planned، Unplanned، Mor (Morning)، Eve (Evening)، Total اور Cp کے کالم رکھیں۔ Selfies، Region، City اور Designation بھی شامل کر سکتے ہیں۔

2. Call Log Excel ڈاؤن لوڈ کرکے اپ لوڈ کریں:
• MREP میں لاگ اِن کریں، پھر Reports → Call Reports → Call Log Live کھولیں۔
• جس تاریخ کی رپورٹ بنانی ہے، وہ تاریخ منتخب کرکے Filter کریں۔
• نتائج آنے کے بعد دائیں طرف اسکرول کریں اور Happened کے فلٹر میں Happened منتخب کریں۔
• فلٹر شدہ رپورٹ تیار ہونے پر Excel ڈاؤن لوڈ کریں اور Call Log Excel میں اپ لوڈ کریں۔

3. WhatsApp ایکسٹینشن سے ڈاؤن لوڈ کی ہوئی سیلفی Excel فائلیں اپ لوڈ کریں۔ فائل کے نام میں ملازم کا کوڈ ہونا چاہیے تاکہ درست ملازم اور کال لاگ کی تاریخ کی سیلفیاں ملائی جا سکیں۔

4. Generate Report دبائیں اور پراسیسنگ مکمل ہونے کا انتظار کریں۔ Download کرنے سے پہلے Preview میں رپورٹ دیکھ بھی سکتے ہیں۔`,
};

/** Serves the approved Daily Report guide; general AI replies are not connected yet. */
export function getDailyReportHelp(message: string): string | null {
  if (!/\b(daily|daliy|dayli|rozana)\b|ڈیلی|روزانہ/i.test(message)) return null;
  return DAILY_REPORT_HELP[getSupportLanguage(message)];
}

export type SupportLanguage = keyof typeof DAILY_REPORT_HELP;

export function getSupportLanguage(message: string, fallback: SupportLanguage = "english"): SupportLanguage {
  if (/[\u0600-\u06ff]/u.test(message)) return "urdu";
  if (/\b(kaise|kaisa|kase|kasa|kese|kesay|kaisay|kesy|kaisy|kro|karna|karein|banani|banane|bana|banau|banaye|bna|bnani|bnau|kya|kia|mujhe|muje|mjy|mujhy|chahiye|chaiye|chahye|kahan|kahn|kidhar|kidr|batao|btao|bta|btaein|samjhao|ki|ka|ke|hai|ha|hain|hn|nahi|nhi|nai|nahin|han|haan|ji|shukriya|shukria|shukrya|samajh|samaj|ap|aap|mera|meri|mra|mri|krna|krni|krun|kryn|karo|kron|kasy|kesy|hal|haal|ho|hun|hoon|rha|rhi|rhy|rahe|theek|thik|salam|salaam|alaikum|kon|kaun|koun|mn|mein|main|bhi|bi|thk|sari|saari|sb|tamam|samjha|waqt|jaldi|mukhtasar|thora|tora)\b/i.test(message)) return "romanUrdu";
  // Short acknowledgements such as "thanks", "thankx", "yes" and "ok"
  // keep the conversation language rather than switching it to English.
  if (/\b(i|you|your|the|please|how|where|which|what|can|for|need|understand|understood)\b/i.test(message)) return "english";
  return fallback;
}

export function getDailyReportStepHelp(message: string, language: SupportLanguage): string | null {
  const topics: [RegExp, number][] = [
    [/call\s*log|calllog|happen(?:ed|d)?|filter|کال\s*لاگ|فلٹر|ہیپنڈ/i, 2],
    [/selfi(?:e|es)?|slefi|slefie|whats\s*app|whatsapp|extension|سیلفی|سیلفیاں|واٹس|ایکسٹینشن/i, 3],
    [/sample|sampl[ey]?|template|columns?|format|نمونہ|سیمپل|کالم|فارمیٹ/i, 1],
    [/preview|download|ڈاؤن\s*لوڈ|پریویو/i, 4],
    [/gener?ate|genrate|genarate|process|ready|wait|تیار|جنریٹ/i, 4],
  ];
  const topic = topics.find(([pattern]) => pattern.test(message));
  return topic ? DAILY_REPORT_HELP[language].split(/\n\n/)[topic[1]] : null;
}
