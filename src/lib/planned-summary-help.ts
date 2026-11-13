import type { SupportLanguage } from "./daily-report-help.ts";

/** Admin-approved customer help only; does not change report calculations. */
export const PLANNED_SUMMARY_HELP: Record<SupportLanguage, string> = {
  english: `To generate your Planned Summary report:

1. Upload your sample employee Excel. Include the employees whose reports you need, with Employee Code and Name filled in. Keep your team sheets and sample format. You may also keep Region and Designation columns.

2. In MREP, open Reports → Call Reports → Call Log.

3. Select all the dates you need for this summary, then select Filter and wait for the results to load.

4. Scroll to the right and set the Happened filter to Happened. Once the filtered report is ready, download it as Excel.

5. Upload the downloaded file under Monthly Call Log Excel in Planned Summary. You can attach multiple call-log Excel files when needed for your reporting period.

6. Select Generate Report and wait for processing to finish. When the report is ready, you can select Preview to check it before selecting Download.`,
  romanUrdu: `Planned Summary report banane ke liye:

1. Sample employee Excel upload karein. Jin employees ki report chahiye, unke Employee Code aur Name bhare hue hon. Apni team sheets aur sample ka format rehne dein. Region aur Designation ke columns bhi rakh sakte hain.

2. MREP mein Reports → Call Reports → Call Log kholein.

3. Jitni dates ki summary chahiye, woh dates select karein. Phir Filter karein aur results load hone ka intezar karein.

4. Right taraf scroll karein aur Happened ke filter mein Happened select karein. Filtered report tayyar ho jaye to Excel download karein.

5. Download ki hui file Planned Summary ke Monthly Call Log Excel mein upload karein. Reporting period ke liye zaroorat ho to multiple call-log Excel files bhi de sakte hain.

6. Generate Report dabayein aur processing mukammal hone ka intezar karein. Report tayyar ho jaye to Download se pehle Preview mein check kar sakte hain.`,
  urdu: `پلانڈ سمری رپورٹ بنانے کے لیے:

1. ملازمین کی نمونہ Excel فائل اپ لوڈ کریں۔ جن ملازمین کی رپورٹ چاہیے، ان کے Employee Code اور Name درج ہوں۔ اپنی ٹیم شیٹس اور نمونہ فائل کا فارمیٹ برقرار رکھیں۔ Region اور Designation کے کالم بھی رکھ سکتے ہیں۔

2. MREP میں Reports → Call Reports → Call Log کھولیں۔

3. جن تاریخوں کی سمری چاہیے وہ تاریخیں منتخب کریں۔ پھر Filter کریں اور نتائج لوڈ ہونے کا انتظار کریں۔

4. دائیں طرف اسکرول کریں اور Happened کے فلٹر میں Happened منتخب کریں۔ فلٹر شدہ رپورٹ تیار ہونے پر Excel ڈاؤن لوڈ کریں۔

5. ڈاؤن لوڈ کی ہوئی فائل Planned Summary کے Monthly Call Log Excel میں اپ لوڈ کریں۔ متعلقہ مدت کے لیے ضرورت ہو تو کئی کال لاگ Excel فائلیں بھی دے سکتے ہیں۔

6. Generate Report دبائیں اور پراسیسنگ مکمل ہونے کا انتظار کریں۔ رپورٹ تیار ہونے پر Download سے پہلے Preview میں دیکھ سکتے ہیں۔`,
};

export function isPlannedSummaryQuestion(message: string): boolean {
  return /\bplanned\s+summary\b|\bmonthly\s+planned\s+unplanned\b|پلانڈ\s*سمری/i.test(message);
}
