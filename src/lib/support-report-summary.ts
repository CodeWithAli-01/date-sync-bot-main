import type { SupportLanguage } from "./daily-report-help.ts";
import { normalizeSupportText } from "./support-smalltalk.ts";

export type SupportReportKey = "daily" | "monthly" | "doctor" | "planned" | "avg" | "mrep";

const SUMMARIES: Record<SupportLanguage, Record<SupportReportKey, string>> = {
  english: {
    daily: "Daily Report: Upload the employee sample Excel and the selected date's Call Log Excel from MREP → Reports → Call Reports → Call Log Live. Filter by date and Happened before downloading. Add WhatsApp selfie Excel files if needed.",
    monthly: "Monthly Report: Upload the employee sample Excel and the Daily Report PDF files for the dates you want included in the monthly summary.",
    doctor: "Doctor Coverage: Upload the employee sample Excel. In MREP → Reports → Call Reports → Doctor Coverage, select month and team, set Territory to Select All, then Filter, download Excel and upload it as the source file.",
    planned: "Planned Summary: Upload the employee sample Excel. Download MREP Call Log Excel for all required dates after applying Filter and Happened. Upload it under Monthly Call Log Excel.",
    avg: "Avg. Per Day: Upload the employee sample Excel. In MREP → Reports → Call Reports → Avg Per Day, select month, Filter and download Excel. Upload it under Avg Report Excel.",
    mrep: "MREP Monthly Report: Upload the employee sample Excel and the month's Employee-wise Call Status Report Excel from MREP. Optionally add WhatsApp selfie Excel exports containing the dates you need.",
  },
  romanUrdu: {
    daily: "Daily Report: Employee sample Excel aur selected date ki Call Log Excel dein. MREP → Reports → Call Reports → Call Log Live mein date aur Happened filter laga kar Excel download karein. Selfies chahiyein to WhatsApp selfie Excel bhi dein.",
    monthly: "Monthly Report: Employee sample Excel aur jin dates ki monthly summary chahiye unki Daily Report PDF files upload karein.",
    doctor: "Doctor Coverage: Employee sample Excel dein. MREP → Reports → Call Reports → Doctor Coverage mein month aur team select karein, Territory mein Select All karein, Filter ke baad Excel download karke source file mein upload karein.",
    planned: "Planned Summary: Employee sample Excel dein. MREP Call Log mein required dates select karke Filter aur Happened filter lagayein. Excel download karke Monthly Call Log Excel mein upload karein.",
    avg: "Avg. Per Day: Employee sample Excel dein. MREP → Reports → Call Reports → Avg Per Day mein month select karke Filter aur Excel download karein. File Avg Report Excel mein upload karein.",
    mrep: "MREP Monthly Report: Employee sample Excel aur MREP se selected month ki Employee-wise Call Status Report Excel upload karein. Selfies chahiyein to required dates ke WhatsApp selfie Excel exports bhi dein.",
  },
  urdu: {
    daily: "ڈیلی رپورٹ: ملازمین کی نمونہ Excel اور منتخب تاریخ کی Call Log Excel دیں۔ MREP → Reports → Call Reports → Call Log Live میں تاریخ اور Happened فلٹر لگا کر Excel ڈاؤن لوڈ کریں۔ سیلفیاں چاہیے ہوں تو WhatsApp سیلفی Excel بھی دیں۔",
    monthly: "ماہانہ رپورٹ: ملازمین کی نمونہ Excel اور جن تاریخوں کا ماہانہ خلاصہ چاہیے ان کی ڈیلی رپورٹ PDF فائلیں اپ لوڈ کریں۔",
    doctor: "ڈاکٹر کوریج: ملازمین کی نمونہ Excel دیں۔ MREP → Reports → Call Reports → Doctor Coverage میں مہینہ، ٹیم اور Territory میں Select All منتخب کریں۔ Filter کے بعد Excel ڈاؤن لوڈ کرکے اصل فائل کے طور پر اپ لوڈ کریں۔",
    planned: "پلانڈ سمری: ملازمین کی نمونہ Excel دیں۔ MREP Call Log میں مطلوبہ تاریخیں منتخب کرکے Filter اور Happened فلٹر لگائیں۔ Excel ڈاؤن لوڈ کرکے Monthly Call Log Excel میں اپ لوڈ کریں۔",
    avg: "Avg. Per Day: ملازمین کی نمونہ Excel دیں۔ MREP → Reports → Call Reports → Avg Per Day میں مہینہ منتخب کرکے Filter اور Excel ڈاؤن لوڈ کریں۔ فائل Avg Report Excel میں اپ لوڈ کریں۔",
    mrep: "MREP Monthly Report: ملازمین کی نمونہ Excel اور MREP سے منتخب مہینے کی Employee-wise Call Status Report Excel اپ لوڈ کریں۔ سیلفیاں چاہیے ہوں تو متعلقہ تاریخوں کے WhatsApp سیلفی Excel ایکسپورٹس بھی دیں۔",
  },
};

export function getSupportReportSummary(language: SupportLanguage, report?: SupportReportKey): string {
  const entries = report ? [SUMMARIES[language][report]] : Object.values(SUMMARIES[language]);
  const ending = {
    english: "Then select Generate Report, wait for processing, check Preview if needed, and Download. Keep employee codes correct in the sample and source files.",
    romanUrdu: "Phir Generate Report dabayein, processing ka intezar karein, zaroorat ho to Preview dekhein aur Download karein. Sample aur source mein employee codes durust rakhein.",
    urdu: "پھر Generate Report دبائیں، پراسیسنگ مکمل ہونے دیں، ضرورت ہو تو Preview دیکھیں اور Download کریں۔ نمونہ اور اصل فائل میں ملازمین کے کوڈ درست رکھیں۔",
  }[language];
  return entries.map((entry, index) => `${index + 1}. ${entry}`).join("\n\n") + `\n\n${ending}`;
}

export function wantsSupportSummary(message: string): boolean {
  const request = normalizeSupportText(message).replace(/planned\s+summary|پلانڈ\s*سمری/gi, "");
  const explicitSummary = /\b(summary|summry|summery|summarize|summarise|overview|brief|mukhtasar)\b|خلاصہ|مختصر/i.test(request);
  const conciseExplanation = /\b(short|quick|jaldi)\b|جلدی/i.test(request)
    && /\breports?\b|\b(explain|batao|btao|bata|bta|samjhao|samjha)\b|رپورٹ|سمجھا|بتا/i.test(request);
  const allReports = /\b(all|sari|saari|sab|sb|tamam)\b.{0,30}\breports?\b|(?:ساری|تمام|سب).*رپورٹ/i.test(request);
  return explicitSummary || conciseExplanation || allReports;
}

export function wantsAllReportSummaries(message: string): boolean {
  return /\b(all|sari|saari|sab|sb|tamam|every)\b.{0,25}\breports?\b|(?:ساری|تمام|سب).*رپورٹ/i.test(message);
}
