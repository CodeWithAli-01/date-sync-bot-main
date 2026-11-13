import type { SupportLanguage } from "./daily-report-help.ts";

/** Admin-approved customer guidance; independent of coverage report processing. */
export const DOCTOR_COVERAGE_HELP: Record<SupportLanguage, string> = {
  english: `To generate your Doctor Coverage report:

1. Upload your sample employee Excel. Include the employees whose reports you need, with Employee Code and Name filled in. Keep your existing team sheets and sample format. You may also keep Region, City and Designation columns.

2. Download the source report from MREP: open Reports → Call Reports → Doctor Coverage.

3. Select the required month, then select the team. Under Territory, choose Select All. Select Filter and wait for the filtered report to load. Download the report as Excel, then upload that downloaded file in the Doctor Coverage report's source Excel upload card.

4. Select Generate Report and wait for processing to finish. When the report is ready, you can open Preview to check it before selecting Download.`,
  romanUrdu: `Doctor Coverage report banane ke liye:

1. Sample employee Excel upload karein. Jin employees ki report chahiye, unke Employee Code aur Name bhare hue hon. Apni team sheets aur sample ka format rehne dein. Region, City aur Designation ke columns bhi rakh sakte hain.

2. Source report download karne ke liye MREP kholein → Reports → Call Reports → Doctor Coverage par jayein.

3. Required month select karein, phir team select karein. Territory mein Select All karein. Filter dabayein aur filtered report load hone ka intezar karein. Report Excel mein download karein, phir woh downloaded file Doctor Coverage report ke source Excel upload card mein dein.

4. Generate Report dabayein aur processing mukammal hone ka intezar karein. Report tayyar ho jaye to Download se pehle Preview mein check kar sakte hain.`,
  urdu: `ڈاکٹر کوریج رپورٹ بنانے کے لیے:

1. ملازمین کی نمونہ Excel فائل اپ لوڈ کریں۔ جن ملازمین کی رپورٹ چاہیے، ان کے Employee Code اور Name درج ہوں۔ اپنی ٹیم شیٹس اور نمونہ فائل کا فارمیٹ برقرار رکھیں۔ Region، City اور Designation کے کالم بھی رکھ سکتے ہیں۔

2. اصل رپورٹ ڈاؤن لوڈ کرنے کے لیے MREP کھولیں، پھر Reports → Call Reports → Doctor Coverage میں جائیں۔

3. مطلوبہ مہینہ منتخب کریں، پھر ٹیم منتخب کریں۔ Territory میں Select All کریں۔ Filter دبائیں اور فلٹر شدہ رپورٹ لوڈ ہونے کا انتظار کریں۔ رپورٹ Excel میں ڈاؤن لوڈ کریں، پھر وہ فائل Doctor Coverage رپورٹ کے سورس Excel اپ لوڈ کارڈ میں دیں۔

4. Generate Report دبائیں اور پراسیسنگ مکمل ہونے کا انتظار کریں۔ رپورٹ تیار ہونے پر Download سے پہلے Preview میں دیکھ سکتے ہیں۔`,
};

export function isDoctorCoverageQuestion(message: string): boolean {
  return /\bdoctor\s*(coverage|covrage|coverge|coverag|covarage)\b|ڈاکٹر\s*(کوریج|کورج)/i.test(message);
}
