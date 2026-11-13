import type { SupportLanguage } from "./daily-report-help.ts";
import { normalizeSupportText } from "./support-smalltalk.ts";
import type { SupportReportKey } from "./support-report-summary.ts";

const COMPARISON_DETAILS: Record<SupportLanguage, Record<SupportReportKey, string>> = {
  english: {
    daily: "Daily Report: Shows employee activity for a selected date using Call Log Excel. WhatsApp selfie Excel is optional when selfie counts are needed.",
    monthly: "Monthly Report: Combines the required dates' Daily Report PDFs into a period/monthly summary. Its source files are PDFs, rather than the employee-wise Excel used by MREP Monthly Report.",
    doctor: "Doctor Coverage: Uses Doctor Coverage Excel for the selected month and team. It focuses on Target Doctors, Covered Doctors and Coverage %, rather than call totals or working time.",
    planned: "Planned Summary: Uses Call Log Excel for the required dates to calculate face-to-face planned/unplanned calls, morning/evening call totals, CP average time and working averages.",
    avg: "Avg. Per Day: Copies Morning Calls, Evening Calls, Total Calls, Avg. Call Per Day. and Field Days from MREP's Avg Per Day Excel, matching employee codes. It uses the supplied source figures instead of recalculating working times from Call Log rows.",
    mrep: "MREP Monthly Report: Uses the selected month's Employee-wise Call Status Report Excel to fill employee-wise monthly KPI figures. Optional WhatsApp exports provide selfie counts for the dates present in those exports.",
  },
  romanUrdu: {
    daily: "Daily Report: Selected date par employee ka kaam Call Log Excel se dikhati hai. Selfies ki ginti chahiye ho to WhatsApp selfie Excel optional upload hoti hai.",
    monthly: "Monthly Report: Required dates ki Daily Report PDF files ko mila kar monthly summary banati hai. Is ka source PDF hota hai; MREP Monthly Report ki tarah employee-wise Excel nahi.",
    doctor: "Doctor Coverage: Selected month aur team ki Doctor Coverage Excel use karti hai. Is ka focus Target Doctors, Covered Doctors aur Coverage % par hai, calls ya working time par nahi.",
    planned: "Planned Summary: Required dates ki Call Log Excel se face-to-face planned/unplanned calls, morning/evening total calls, CP ka average time aur working averages nikalti hai.",
    avg: "Avg. Per Day: MREP ki Avg Per Day Excel se Morning Calls, Evening Calls, Total Calls, Avg. Call Per Day. aur Field Days employee code match karke leti hai. Yeh source ke diye hue figures use karti hai; Call Log ki rows se working time dobara calculate nahi karti.",
    mrep: "MREP Monthly Report: Selected month ki Employee-wise Call Status Report Excel se employee-wise monthly KPI data leti hai. Optional WhatsApp exports se un exports mein mojood dates ki selfie counts milti hain.",
  },
  urdu: {
    daily: "ڈیلی رپورٹ: منتخب تاریخ کے ملازم کے کام کی تفصیل Call Log Excel سے دکھاتی ہے۔ سیلفیوں کی گنتی چاہیے تو WhatsApp سیلفی Excel اختیاری طور پر اپ لوڈ ہوتی ہے۔",
    monthly: "ماہانہ رپورٹ: متعلقہ تاریخوں کی ڈیلی رپورٹ PDF فائلیں ملا کر ماہانہ خلاصہ بناتی ہے۔ اس کی اصل فائل PDF ہوتی ہے، MREP Monthly Report کی طرح ملازم وار Excel نہیں۔",
    doctor: "ڈاکٹر کوریج: منتخب مہینے اور ٹیم کی Doctor Coverage Excel استعمال کرتی ہے۔ توجہ Target Doctors، Covered Doctors اور Coverage % پر ہوتی ہے، کالوں یا کام کے وقت پر نہیں۔",
    planned: "پلانڈ سمری: متعلقہ تاریخوں کی Call Log Excel سے فیس ٹو فیس پلانڈ اور ان پلانڈ کالیں، صبح و شام کی کل کالیں، اوسط CP وقت اور کام کے اوسط اوقات نکالتی ہے۔",
    avg: "Avg. Per Day: MREP کی Avg Per Day Excel سے ملازم کا کوڈ ملا کر Morning Calls، Evening Calls، Total Calls، Avg. Call Per Day. اور Field Days لیتی ہے۔ اصل فائل کے اعداد استعمال ہوتے ہیں؛ کال لاگ کی قطاروں سے کام کے اوقات دوبارہ نہیں نکالے جاتے۔",
    mrep: "MREP Monthly Report: منتخب مہینے کی Employee-wise Call Status Report Excel سے ملازم وار ماہانہ KPI ڈیٹا لیتی ہے۔ اختیاری WhatsApp ایکسپورٹس سے ان میں موجود تاریخوں کی سیلفیوں کی گنتی ملتی ہے۔",
  },
};

export function mentionedComparisonReports(message: string): SupportReportKey[] {
  let text = normalizeSupportText(message);
  const found = new Set<SupportReportKey>();
  text = text.replace(/\bmrep\s+(?:monthly|mentholy|monthlie)\s*(?:report)?\b|ایم\s*ریپ\s*(?:ماہانہ|منتھلی)\s*(?:رپورٹ)?/gi, () => { found.add("mrep"); return " "; });
  text = text.replace(/\bmonthly\s+planned\s+unplanned\b/gi, () => { found.add("planned"); return " "; });
  if (/\b(daily|daliy|dayli|rozana)\b|ڈیلی|روزانہ/i.test(text)) found.add("daily");
  if (/\b(monthly|mentholy|monthlie)\b|ماہانہ|منتھلی/i.test(text)) found.add("monthly");
  if (/\bdoctor\s*(coverage|covrage|coverge)\b|ڈاکٹر\s*(کوریج|کورج)/i.test(text)) found.add("doctor");
  if (/\bplanned\s+summary\b|پلانڈ\s*سمری/i.test(text)) found.add("planned");
  if (/\b(avg\.?|average|avgerage)\s*(per\s*)?day\b|ایوریج\s*(پر\s*)?ڈے|اوسط\s*فی\s*دن/i.test(text)) found.add("avg");
  return (["daily", "monthly", "doctor", "planned", "avg", "mrep"] as const).filter((key) => found.has(key));
}

function typoDistance(a: string, b: string): number {
  const grid = Array.from({ length: a.length + 1 }, (_, i) => Array.from({ length: b.length + 1 }, (_, j) => i === 0 ? j : j === 0 ? i : 0));
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) {
    grid[i][j] = Math.min(grid[i - 1][j] + 1, grid[i][j - 1] + 1, grid[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) grid[i][j] = Math.min(grid[i][j], grid[i - 2][j - 2] + 1);
  }
  return grid[a.length][b.length];
}

export function isReportComparisonRequest(message: string): boolean {
  const text = normalizeSupportText(message);
  if (/\b(difference|diffrence|differance|deference|diff|farq|farak|faraq|vs|versus|compare|comparison)\b|فرق|موازنہ/i.test(text)) return true;
  const hasReports = mentionedComparisonReports(text).length >= 2 || /\breports\b|رپورٹوں/i.test(text);
  if (/\bbetween\b/.test(text) && hasReports) return true;
  return (text.match(/\b[a-z]+\b/g) ?? []).some((word) => {
    const compact = word.replace(/([a-z])\1+/g, "$1");
    if (compact === "diferent") return hasReports;
    return /^dif/.test(compact) && typoDistance(compact, "diference") <= 2;
  });
}

export function getComparisonClarification(language: SupportLanguage): string {
  return {
    english: "Which reports would you like to compare? Name two or more reports, or say 'all reports'.",
    romanUrdu: "Aap kin reports ka farq samajhna chahte hain? Do ya zyada reports ke naam bata dein, ya 'sari reports' keh dein.",
    urdu: "آپ کن رپورٹس کا فرق سمجھنا چاہتے ہیں؟ دو یا زیادہ رپورٹس کے نام بتائیں یا 'تمام رپورٹس' کہہ دیں۔",
  }[language];
}

export function getReportComparison(message: string, language: SupportLanguage): string | null {
  if (!isReportComparisonRequest(message)) return null;
  const named = mentionedComparisonReports(message);
  const all = /\b(all|sari|saari|sab|sb|tamam|every)\b.{0,25}\breports?\b|\b(in|inn|un)\s+(sab|sb)\b|(?:ساری|تمام|سب).*رپورٹ|ان سب/i.test(message)
    || (named.length < 2 && /\b(all|sari|saari|sab|sb|tamam)\b|ساری|تمام|سب/i.test(message));
  const reports = all ? ["daily", "monthly", "doctor", "planned", "avg", "mrep"] as const : named;
  if (reports.length < 2) return null;
  if (reports.length === 2 && reports.includes("daily") && reports.includes("monthly")) return DAILY_MONTHLY_DIFFERENCE[language];
  const heading = {
    english: "These reports use different source files and serve different purposes:",
    romanUrdu: "In reports ka farq source files aur unke kaam mein hai:",
    urdu: "ان رپورٹس میں فرق اصل فائلوں اور ان کے مقصد کا ہے:",
  }[language];
  return heading + "\n\n" + reports.map((key, index) => `${index + 1}. ${COMPARISON_DETAILS[language][key]}`).join("\n\n");
}

const DAILY_MONTHLY_DIFFERENCE: Record<SupportLanguage, string> = {
  english: `Daily Report and Monthly Report cover different reporting periods and use different source files:

• Daily Report: For a selected date. Upload the employee sample Excel and Call Log Excel for that date. Add WhatsApp selfie Excel files if selfie counts are needed.
• Monthly Report: Combines the required dates' Daily Report PDFs into a monthly summary. Upload the employee sample Excel and those PDF files.

Both reports can be generated, previewed and downloaded.`,
  romanUrdu: `Daily Report aur Monthly Report ka farq dates aur source files ka hai:

• Daily Report: Ek selected date ki report hoti hai. Employee sample Excel aur us date ki Call Log Excel dein. Selfie counts chahiye hon to WhatsApp selfie Excel files bhi dein.
• Monthly Report: Required dates ki Daily Reports ki PDF files ko mila kar monthly summary banti hai. Employee sample Excel aur woh PDF files dein.

Dono reports ko generate karke Preview mein dekh sakte hain aur Download kar sakte hain.`,
  urdu: `ڈیلی رپورٹ اور ماہانہ رپورٹ میں فرق تاریخوں اور اصل فائلوں کا ہے:

• ڈیلی رپورٹ: ایک منتخب تاریخ کی رپورٹ ہوتی ہے۔ ملازمین کی نمونہ Excel اور اس تاریخ کی Call Log Excel دیں۔ سیلفیوں کی گنتی چاہیے تو WhatsApp سیلفی Excel فائلیں بھی دیں۔
• ماہانہ رپورٹ: متعلقہ تاریخوں کی ڈیلی رپورٹس کی PDF فائلیں ملا کر ماہانہ خلاصہ بنتا ہے۔ ملازمین کی نمونہ Excel اور وہ PDF فائلیں دیں۔

دونوں رپورٹس بنا کر Preview میں دیکھ سکتے ہیں اور Download کر سکتے ہیں۔`,
};

export function getDailyMonthlyComparison(message: string, language: SupportLanguage): string | null {
  return getReportComparison(message, language);
}
