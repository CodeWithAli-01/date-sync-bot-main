import type { SupportLanguage } from "./daily-report-help.ts";

/** Admin-approved guidance for Help & Support only. */
export const AVG_PER_DAY_HELP: Record<SupportLanguage, string> = {
  english: `To generate your Avg. Per Day report:

1. Upload your sample employee Excel. Include the employees whose reports you need, with their correct Employee Code / Emp. ID and Name. Employee codes must match the source report. Keep your usual team sheets and sample format.

2. In MREP, open Reports → Call Reports → Avg Per Day. Select the month you need, then select Filter. When the filtered report is ready, download it as Excel. Upload this downloaded file under Avg Report Excel in the Avg. Per Day report.

The source Excel should contain Employee Code and these report fields: Morning Calls, Evening Calls, Total Calls, Avg. Call Per Day. and Field Days. Values are matched to each employee's code.

3. Select Generate Report and wait for processing to finish. When the report is ready, you can open Preview to check it before selecting Download.`,
  romanUrdu: `Avg. Per Day report banane ke liye:

1. Sample employee Excel upload karein. Jin employees ki report chahiye, unke durust Employee Code / Emp. ID aur Name dein. Employee codes source report se match hone chahiyein. Apni team sheets aur sample ka format rehne dein.

2. MREP mein Reports → Call Reports → Avg Per Day kholein. Jis month ki report chahiye woh month select karke Filter karein. Filtered report tayyar ho jaye to Excel download karein. Yeh downloaded file Avg. Per Day report ke Avg Report Excel mein upload karein.

Source Excel mein Employee Code ke saath Morning Calls, Evening Calls, Total Calls, Avg. Call Per Day. aur Field Days hone chahiyein. Data har employee ke code se match hota hai.

3. Generate Report dabayein aur processing mukammal hone ka intezar karein. Report tayyar ho jaye to Download se pehle Preview mein check kar sakte hain.`,
  urdu: `Avg. Per Day رپورٹ بنانے کے لیے:

1. ملازمین کی نمونہ Excel فائل اپ لوڈ کریں۔ جن ملازمین کی رپورٹ چاہیے، ان کے درست Employee Code / Emp. ID اور Name درج کریں۔ ملازم کا کوڈ اصل رپورٹ سے ملنا چاہیے۔ اپنی ٹیم شیٹس اور نمونہ فائل کا فارمیٹ برقرار رکھیں۔

2. MREP میں Reports → Call Reports → Avg Per Day کھولیں۔ جس مہینے کی رپورٹ چاہیے وہ مہینہ منتخب کرکے Filter کریں۔ فلٹر شدہ رپورٹ تیار ہونے پر Excel ڈاؤن لوڈ کریں۔ یہ فائل Avg. Per Day رپورٹ کے Avg Report Excel میں اپ لوڈ کریں۔

اصل Excel میں Employee Code کے ساتھ Morning Calls، Evening Calls، Total Calls، Avg. Call Per Day. اور Field Days ہونے چاہئیں۔ ہر ملازم کا ڈیٹا اس کے کوڈ سے ملایا جاتا ہے۔

3. Generate Report دبائیں اور پراسیسنگ مکمل ہونے کا انتظار کریں۔ رپورٹ تیار ہونے پر Download سے پہلے Preview میں دیکھ سکتے ہیں۔`,
};

export function isAvgPerDayQuestion(message: string): boolean {
  return /\b(?:avg\.?|average|avgerage)\s*(?:per\s*)?day\b|ایوریج\s*(?:پر\s*)?ڈے|اوسط\s*فی\s*دن/i.test(message);
}
