import type { SupportLanguage } from "./daily-report-help.ts";

/** Admin-approved Help & Support guide, separate from report generation. */
export const MREP_MONTHLY_HELP: Record<SupportLanguage, string> = {
  english: `To generate your MREP Monthly Report:

1. Upload the sample employee Excel under Sample Employee File. Include the employees you need, with their correct Employee Code / Emp. ID and Name. Keep your team sheets, sample format and any existing report columns. Employee codes should match the source report.

2. In MREP, open Reports → Call Reports → Employee-wise Call Status Report. Select the required month, then select Filter. When the report is ready, download it as Excel and upload it under Employee-wise Call Report.

3. If you also need selfie counts, upload the Excel files downloaded using your WhatsApp extension under WhatsApp Selfie Export Files. Include employee codes in the filenames. Selfie counts depend on the dates and data present in the uploaded exports, so use exports containing the dates you want included. This upload is optional when selfie counts are not needed.

4. Select Generate Report and wait for processing to finish. When the report is ready, you can open Preview to check it before selecting Download.`,
  romanUrdu: `MREP Monthly Report banane ke liye:

1. Sample Employee File mein sample employee Excel upload karein. Jin employees ki report chahiye, unke durust Employee Code / Emp. ID aur Name dein. Apni team sheets, sample ka format aur pehle se mojood report columns rehne dein. Employee codes source report se match hone chahiyein.

2. MREP mein Reports → Call Reports → Employee-wise Call Status Report kholein. Jis month ki report chahiye woh month select karke Filter karein. Report tayyar ho jaye to Excel download karein aur Employee-wise Call Report mein upload karein.

3. Agar selfie counts bhi chahiyein to WhatsApp extension se download ki hui Excel files, WhatsApp Selfie Export Files mein upload karein. File ke naam mein employee code ho. Selfies ki ginti uploaded exports mein mojood dates aur data ke mutabiq hogi, is liye unhi dates ka export dein jo report mein chahiyein. Selfies nahi chahiyein to yeh upload optional hai.

4. Generate Report dabayein aur processing mukammal hone ka intezar karein. Report tayyar ho jaye to Download se pehle Preview mein check kar sakte hain.`,
  urdu: `MREP Monthly Report بنانے کے لیے:

1. Sample Employee File میں ملازمین کی نمونہ Excel فائل اپ لوڈ کریں۔ جن ملازمین کی رپورٹ چاہیے، ان کے درست Employee Code / Emp. ID اور Name درج کریں۔ اپنی ٹیم شیٹس، نمونہ فائل کا فارمیٹ اور موجودہ رپورٹ کالم برقرار رکھیں۔ ملازمین کے کوڈ اصل رپورٹ سے ملنے چاہئیں۔

2. MREP میں Reports → Call Reports → Employee-wise Call Status Report کھولیں۔ مطلوبہ مہینہ منتخب کرکے Filter کریں۔ رپورٹ تیار ہونے پر Excel ڈاؤن لوڈ کریں اور Employee-wise Call Report میں اپ لوڈ کریں۔

3. اگر سیلفیوں کی گنتی بھی چاہیے تو WhatsApp ایکسٹینشن سے ڈاؤن لوڈ کی ہوئی Excel فائلیں، WhatsApp Selfie Export Files میں اپ لوڈ کریں۔ فائل کے نام میں ملازم کا کوڈ ہو۔ سیلفیوں کی گنتی اپ لوڈ شدہ ایکسپورٹس میں موجود تاریخوں اور ڈیٹا کے مطابق ہوگی، اس لیے صرف متعلقہ تاریخوں کے ایکسپورٹس دیں۔ سیلفیوں کی گنتی نہیں چاہیے تو یہ اپ لوڈ اختیاری ہے۔

4. Generate Report دبائیں اور پراسیسنگ مکمل ہونے کا انتظار کریں۔ رپورٹ تیار ہونے پر Download سے پہلے Preview میں دیکھ سکتے ہیں۔`,
};

export function isMrepMonthlyQuestion(message: string): boolean {
  return /\bmrep\b|ایم\s*ریپ/i.test(message)
    && /\b(monthly|mentholy|monthl[yie])\b|منتھلی|ماہانہ/i.test(message);
}
