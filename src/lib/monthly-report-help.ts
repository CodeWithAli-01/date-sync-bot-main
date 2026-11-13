import type { SupportLanguage } from "./daily-report-help.ts";

/** Admin-approved Monthly Report guidance for customer support only. */
export const MONTHLY_REPORT_HELP: Record<SupportLanguage, string> = {
  english: `Monthly Report — Step 1: Sample Excel

Upload the sample employee Excel workbook in the Monthly Report's sample upload card. Include the employees whose reports you want to generate.

• Keep a Name / Employee Name column with each employee's name filled in.
• Include Employee Code / Emp. ID for accurate matching. Each employee's code should be correct and consistent with the source report.
• You can also include Region, City and Designation. Keep your usual team sheets and sample formatting.
• If your sample already contains Planned, Unplanned, Calls and Selfies columns, keep them in place.

Step 2: Daily Report PDF files
Upload the PDF files of the Daily Reports for the dates you want included in the Monthly Report. You can attach multiple PDF files together.

Step 3: Generate Report
Select Generate Report and wait for processing to finish. When the report is ready, you can select Preview to check it before selecting Download. Processing time depends on the number and size of the PDF files.`,
  romanUrdu: `Monthly Report — Step 1: Sample Excel

Monthly Report ke sample upload card mein apni employee sample Excel file dein. Jin employees ki report chahiye, unhein sample mein shamil karein.

• Name / Employee Name ka column ho aur har employee ka naam bhara hua ho.
• Sahi matching ke liye Employee Code / Emp. ID bhi dein. Har employee ka code durust ho aur source report se match kare.
• Region, City aur Designation bhi rakh sakte hain. Apni team sheets aur sample ka mamool ka format rehne dein.
• Agar sample mein Planned, Unplanned, Calls aur Selfies ke columns pehle se hain to unhein rehne dein.

Step 2: Daily Report ki PDF files
Jin dates ki Monthly Report banani hai, un dates ki Daily Reports ki PDF files upload karein. Aap ek saath multiple PDF files bhi de sakte hain.

Step 3: Generate Report
Generate Report dabayein aur processing mukammal hone ka intezar karein. Report tayyar ho jaye to Download se pehle Preview mein check kar sakte hain. Processing ka waqt PDF files ki tadaad aur size ke mutabiq ho sakta hai.`,
  urdu: `ماہانہ رپورٹ — مرحلہ 1: نمونہ Excel

Monthly Report کے سیمپل اپ لوڈ کارڈ میں ملازمین کی نمونہ Excel فائل دیں۔ جن ملازمین کی رپورٹ بنانی ہے، انہیں نمونہ فائل میں شامل کریں۔

• Name / Employee Name کا کالم ہو اور ہر ملازم کا نام درج ہو۔
• درست میچنگ کے لیے Employee Code / Emp. ID بھی دیں۔ ہر ملازم کا کوڈ درست ہو اور اصل رپورٹ سے ملتا ہو۔
• Region، City اور Designation بھی شامل کر سکتے ہیں۔ اپنی ٹیم شیٹس اور نمونہ فائل کا معمول کا فارمیٹ برقرار رکھیں۔
• اگر Planned، Unplanned، Calls اور Selfies کے کالم پہلے سے موجود ہیں تو انہیں برقرار رکھیں۔

مرحلہ 2: ڈیلی رپورٹ کی PDF فائلیں
جن تاریخوں کی ماہانہ رپورٹ بنانی ہے، ان تاریخوں کی ڈیلی رپورٹس کی PDF فائلیں اپ لوڈ کریں۔ ایک ساتھ کئی PDF فائلیں بھی دے سکتے ہیں۔

مرحلہ 3: Generate Report
Generate Report دبائیں اور پراسیسنگ مکمل ہونے کا انتظار کریں۔ رپورٹ تیار ہونے پر Download سے پہلے Preview میں دیکھ سکتے ہیں۔ پراسیسنگ کا وقت PDF فائلوں کی تعداد اور سائز کے مطابق ہو سکتا ہے۔`,
};

export function isMonthlyReportQuestion(message: string): boolean {
  return /\b(monthly|monthl[yie]|mentholy)\b|ماہانہ|منتھلی/i.test(message)
    && !/\bmrep\b|ایم ریپ/i.test(message);
}
