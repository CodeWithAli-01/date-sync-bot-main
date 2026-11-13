import { DAILY_REPORT_HELP, type SupportLanguage } from "./daily-report-help.ts";
import { MONTHLY_REPORT_HELP } from "./monthly-report-help.ts";
import { DOCTOR_COVERAGE_HELP } from "./doctor-coverage-help.ts";
import { PLANNED_SUMMARY_HELP } from "./planned-summary-help.ts";
import { AVG_PER_DAY_HELP } from "./avg-per-day-help.ts";
import { MREP_MONTHLY_HELP } from "./mrep-monthly-help.ts";
import type { SupportReportKey } from "./support-report-summary.ts";

/** Re-explain the relevant approved step without changing the report method. */
export function getFocusedSupportHelp(message: string, language: SupportLanguage, report: SupportReportKey): string | null {
  const guides = { daily: DAILY_REPORT_HELP, monthly: MONTHLY_REPORT_HELP, doctor: DOCTOR_COVERAGE_HELP, planned: PLANNED_SUMMARY_HELP, avg: AVG_PER_DAY_HELP, mrep: MREP_MONTHLY_HELP };
  const guide = guides[report][language];
  const sections = report === "monthly"
    ? guide.split(/\n\n(?=Step\s+\d|مرحلہ\s+\d)/)
    : guide.split(/(?=^\d+\.\s)/m).filter((part) => /^\d+\.\s/.test(part));
  let indexes: number[] | undefined;
  if (/sample|template|columns?|format|employee\s*(code|id|name)|سیمپل|نمونہ|کالم|فارمیٹ/i.test(message)) indexes = [0];
  else if (/selfi|slefi|whatsapp|extension|export|سیلفی|واٹس|ایکسٹینشن/i.test(message)) {
    if (report === "daily" || report === "mrep") indexes = [2];
  } else if (/call\s*log|calllog|mrep|happened|filter|territory|team|month|source|\bpdf\b|کال لاگ|فلٹر|ٹیم|مہینہ/i.test(message)) {
    indexes = report === "planned" ? [1, 2, 3, 4] : report === "doctor" ? [1, 2] : [1];
  } else if (/generate|genrate|processing|ready|wait|preview|download|تیار|پریویو|ڈاؤن/i.test(message)) indexes = [sections.length - 1];
  return indexes ? indexes.map((index) => sections[index]?.trim()).filter(Boolean).join("\n\n") || null : null;
}
