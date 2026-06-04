// UPSERT-only PostgreSQL sync for the Pharma Selfie Reporting System.
import { supabase } from "@/integrations/supabase/client";
import type { PdfParseResult } from "./pdf-extractor";

export interface EmployeeInput {
  name: string;
  code: string | null;
  region: string | null;
  city: string | null;
  designation: string | null;
  originalOrder: number;
}

export interface SyncInput {
  employees: EmployeeInput[];
  pdfResults: PdfParseResult[];
  failedFiles?: { fileName: string; error: string }[];
}

export interface SyncResult {
  syncedRecords: number;
  skippedRows: number;
  parsingErrors: number;
}

export interface GeneratedReportSyncInput {
  id: string;
  fileName: string;
  reportType?: string;
  dates: string[];
  pdfCount: number;
  totalEmployees: number;
  matchedEmployees: number;
  size: number;
  blob: Blob;
}

interface EmployeeRow extends EmployeeInput {
  nameKey: string;
  cleanNameKey: string;
}

interface DailyRecordInput {
  employee_code: string;
  report_date: string;
  selfie_text: string;
  selfie_count: number;
  total_count: number;
  source_file_id: string | null;
  updated_at: string;
}

type LooseQuery = {
  select: (columns?: string) => LooseQuery;
  insert: (values: unknown) => LooseQuery;
  upsert: (values: unknown, options?: unknown) => LooseQuery;
  eq: (column: string, value: unknown) => LooseQuery;
  in: (column: string, values: unknown[]) => LooseQuery;
  gt: (column: string, value: unknown) => LooseQuery;
  maybeSingle: () => Promise<{ data: Record<string, unknown> | null; error: unknown }>;
  then: Promise<{ data: Record<string, unknown>[] | null; error: unknown }>["then"];
};

function table(name: string): LooseQuery {
  return supabase.from(name as "employees") as unknown as LooseQuery;
}

function normalize(s: string): string {
  return (s ?? "").toString().replace(/\s+/g, " ").trim().toLowerCase();
}

function fuzzyKey(s: string): string {
  return normalize(s)
    .replace(/[^\p{L}\p{N}\s]/gu, "")
    .replace(/\s+/g, " ");
}

function normalizeEmployeeCode(value: string | null): string | null {
  const digits =
    String(value ?? "")
      .match(/\d+/g)
      ?.join("") ?? "";
  return digits.length >= 4 ? digits : null;
}

function codeSuffixes(code: string): string[] {
  const out: string[] = [];
  for (const len of [4, 5]) {
    if (code.length > len) out.push(code.slice(-len));
  }
  return out;
}

function personNameKey(s: string): string {
  const noise = new Set([
    "mr",
    "mrs",
    "ms",
    "dr",
    "psv",
    "mio",
    "asm",
    "rsm",
    "sm",
    "am",
    "fm",
    "tm",
    "swt",
    "hq",
    "territory",
    "location",
    "locations",
    "new",
    "joining",
    "resigned",
    "meeting",
    "grp",
    "with",
    "selfie",
    "selfies",
    "total",
  ]);
  return fuzzyKey(s.replace(/\b(new joining|resigned|meeting)\b.*$/i, " "))
    .split(" ")
    .filter((token) => token && !/^\d+$/.test(token) && !noise.has(token))
    .join(" ");
}

function levenshteinSimilarity(a: string, b: string): number {
  if (a === b) return 1;
  if (!a.length || !b.length) return 0;
  const previous = Array.from({ length: b.length + 1 }, (_, i) => i);
  const current = new Array<number>(b.length + 1);
  for (let i = 1; i <= a.length; i++) {
    current[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      current[j] = Math.min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + cost);
    }
    previous.splice(0, previous.length, ...current);
  }
  return 1 - previous[b.length] / Math.max(a.length, b.length);
}

function tokenSimilarity(a: string, b: string): number {
  const aTokens = a.split(" ").filter(Boolean);
  const bTokens = b.split(" ").filter(Boolean);
  if (!aTokens.length || !bTokens.length) return 0;
  let overlap = 0;
  for (const token of aTokens) {
    if (bTokens.some((candidate) => tokensMatch(token, candidate))) overlap++;
  }
  return overlap / Math.max(aTokens.length, bTokens.length);
}

function tokensMatch(a: string, b: string): boolean {
  if (a === b) return true;
  if (a.length === 1 && b.startsWith(a)) return true;
  if (b.length === 1 && a.startsWith(b)) return true;
  return false;
}

function orderedTokenCoverage(shorter: string, longer: string): number {
  const shortTokens = shorter.split(" ").filter(Boolean);
  const longTokens = longer.split(" ").filter(Boolean);
  if (!shortTokens.length || !longTokens.length) return 0;

  let longIndex = 0;
  let matched = 0;
  for (const token of shortTokens) {
    while (longIndex < longTokens.length && !tokensMatch(token, longTokens[longIndex])) {
      longIndex++;
    }
    if (longIndex >= longTokens.length) continue;
    matched++;
    longIndex++;
  }

  return matched / shortTokens.length;
}

function nameMatchScore(a: string, b: string): number {
  if (a === b) return 1;
  const shorter = a.length <= b.length ? a : b;
  const longer = a.length <= b.length ? b : a;
  const coverage = orderedTokenCoverage(shorter, longer);
  const tokenScore = tokenSimilarity(a, b);
  const editScore = levenshteinSimilarity(a, b);
  return Math.max(tokenScore, editScore, coverage >= 1 ? 0.94 : coverage * 0.9);
}

function buildEmployeeIndexes(employees: EmployeeInput[]) {
  const rows: EmployeeRow[] = employees.map((e) => ({
    ...e,
    code: normalizeEmployeeCode(e.code),
    nameKey: normalize(e.name),
    cleanNameKey: personNameKey(e.name),
  }));
  const byCode = new Map<string, EmployeeRow>();
  const byCodeSuffix = new Map<string, EmployeeRow>();
  const byName = new Map<string, EmployeeRow>();
  const byCleanName = new Map<string, EmployeeRow>();
  const duplicateCodes = new Set<string>();
  const duplicateCodeSuffixes = new Set<string>();
  const duplicateNames = new Set<string>();
  const duplicateCleanNames = new Set<string>();
  for (const row of rows) {
    if (row.code) {
      if (byCode.has(row.code)) duplicateCodes.add(row.code);
      else byCode.set(row.code, row);
      for (const suffix of [row.code, ...codeSuffixes(row.code)]) {
        if (byCodeSuffix.has(suffix)) duplicateCodeSuffixes.add(suffix);
        else byCodeSuffix.set(suffix, row);
      }
    }
    if (byName.has(row.nameKey)) duplicateNames.add(row.nameKey);
    byName.set(row.nameKey, row);
    if (byCleanName.has(row.cleanNameKey)) duplicateCleanNames.add(row.cleanNameKey);
    byCleanName.set(row.cleanNameKey, row);
  }
  return {
    rows,
    byCode,
    byCodeSuffix,
    byName,
    byCleanName,
    duplicateCodes,
    duplicateCodeSuffixes,
    duplicateNames,
    duplicateCleanNames,
  };
}

function findEmployee(
  row: { code: string | null; name: string },
  indexes: ReturnType<typeof buildEmployeeIndexes>,
): EmployeeRow | null {
  const nameKey = normalize(row.name);
  const cleanName = personNameKey(row.name);
  const findByUniqueName = () => {
    const byName = indexes.byName.get(nameKey);
    if (byName && !indexes.duplicateNames.has(nameKey)) return byName;

    const byCleanName = indexes.byCleanName.get(cleanName);
    if (byCleanName && !indexes.duplicateCleanNames.has(cleanName)) return byCleanName;

    const prefixMatches = indexes.rows.filter(
      (employee) =>
        employee.cleanNameKey.length >= 8 &&
        !indexes.duplicateCleanNames.has(employee.cleanNameKey) &&
        (cleanName.startsWith(`${employee.cleanNameKey} `) ||
          employee.cleanNameKey.startsWith(`${cleanName} `)),
    );

    if (prefixMatches.length === 1) return prefixMatches[0];

    const scored = indexes.rows
      .map((employee) => ({
        employee,
        score: nameMatchScore(cleanName, employee.cleanNameKey),
      }))
      .filter((candidate) => candidate.score >= 0.86)
      .sort((a, b) => b.score - a.score);

    if (!scored.length) return null;
    const [best, second] = scored;
    return !second || best.score - second.score >= 0.05 ? best.employee : null;
  };

  const code = normalizeEmployeeCode(row.code);
  if (code) {
    const byCode = indexes.byCode.get(code);
    if (byCode) return byCode;
    for (const suffix of [code, ...codeSuffixes(code)]) {
      const bySuffix = indexes.byCodeSuffix.get(suffix);
      if (bySuffix && !indexes.duplicateCodeSuffixes.has(suffix)) return bySuffix;
    }
    return findByUniqueName();
  }

  return findByUniqueName();
}

export async function syncToDatabase(opts: SyncInput): Promise<SyncResult> {
  const { employees, pdfResults, failedFiles = [] } = opts;
  const indexes = buildEmployeeIndexes(employees);
  let skippedRows = 0;

  const employeesWithCodes = indexes.rows.filter((e) => e.code);
  if (employeesWithCodes.length) {
    await table("employees").upsert(
      employeesWithCodes.map((e) => ({
        employee_code: e.code,
        name: e.name,
        region: e.region,
        city: e.city,
        designation: e.designation,
        status: "active",
        original_order: e.originalOrder,
        updated_at: new Date().toISOString(),
      })),
      { onConflict: "employee_code" },
    );
  }

  const fileIdByName = new Map<string, string>();
  for (const r of pdfResults) {
    const { data } = await table("report_files")
      .upsert(
        {
          file_name: r.fileName,
          file_hash: r.fileHash,
          report_date: r.date,
          status: "done",
          processed_status: "done",
        },
        { onConflict: "file_hash" },
      )
      .select("id")
      .maybeSingle();

    if (typeof data?.id === "string") {
      fileIdByName.set(r.fileName, data.id);
    }
  }

  const recordByKey = new Map<string, DailyRecordInput>();
  for (const pdf of pdfResults) {
    for (const row of pdf.rows) {
      const employee = findEmployee(row, indexes);
      if (!employee?.code) {
        skippedRows++;
        continue;
      }

      const key = `${employee.code}|${pdf.date}`;
      const existing = recordByKey.get(key);
      const selfieCount = existing ? Math.max(existing.selfie_count, row.count) : row.count;
      const totalCount = existing ? Math.max(existing.total_count, row.total) : row.total;
      recordByKey.set(key, {
        employee_code: employee.code,
        report_date: pdf.date,
        selfie_count: selfieCount,
        selfie_text: row.selfieText,
        total_count: totalCount,
        source_file_id: fileIdByName.get(pdf.fileName) ?? existing?.source_file_id ?? null,
        updated_at: new Date().toISOString(),
      });
    }
  }

  const records = await preserveExistingPositiveValues([...recordByKey.values()]);
  for (let i = 0; i < records.length; i += 500) {
    await table("daily_records").upsert(records.slice(i, i + 500), {
      onConflict: "employee_code,report_date",
    });
  }

  if (failedFiles.length) {
    await table("error_logs").insert(
      failedFiles.map((f) => ({ file_name: f.fileName, error_message: f.error })),
    );
  }

  return {
    syncedRecords: records.length,
    skippedRows,
    parsingErrors: failedFiles.length,
  };
}

async function preserveExistingPositiveValues(
  records: DailyRecordInput[],
): Promise<DailyRecordInput[]> {
  if (!records.length) return records;

  const codes = [...new Set(records.map((r) => r.employee_code))];
  const dates = [...new Set(records.map((r) => r.report_date))];
  const { data } = await table("daily_records")
    .select("employee_code,report_date,selfie_count,total_count")
    .in("employee_code", codes)
    .in("report_date", dates);

  const existingByKey = new Map<string, Record<string, unknown>>();
  for (const row of data ?? []) {
    existingByKey.set(`${row.employee_code}|${row.report_date}`, row);
  }

  return records.map((record) => {
    const existing = existingByKey.get(`${record.employee_code}|${record.report_date}`);
    const existingSelfies = Number(existing?.selfie_count ?? 0);
    const existingTotal = Number(existing?.total_count ?? 0);
    const selfieCount =
      record.selfie_count === 0 && existingSelfies > 0 ? existingSelfies : record.selfie_count;
    const totalCount =
      record.total_count === 0 && existingTotal > 0 ? existingTotal : record.total_count;

    return {
      ...record,
      selfie_count: selfieCount,
      selfie_text:
        record.selfie_count === 0 && existingSelfies > 0
          ? `${selfieCount} selfies with locations in grp`
          : record.selfie_text,
      total_count: totalCount,
    };
  });
}

export async function findProcessedHashes(hashes: string[]): Promise<Set<string>> {
  if (!hashes.length) return new Set();
  const { data } = await table("report_files").select("file_hash").in("file_hash", hashes);
  return new Set((data ?? []).map((d) => String(d.file_hash)));
}

export async function syncGeneratedReportToDatabase(
  report: GeneratedReportSyncInput,
): Promise<void> {
  const fileHash = await hashBlob(report.blob);
  const createdAt = new Date().toISOString();
  const firstDate = firstValidDate(report.dates);

  try {
    const generatedResult = await table("generated_reports").upsert(
      {
        report_key: fileHash,
        local_history_id: report.id,
        file_name: report.fileName,
        report_type: report.reportType ?? "Report",
        dates: report.dates,
        pdf_count: report.pdfCount,
        total_employees: report.totalEmployees,
        matched_employees: report.matchedEmployees,
        file_size: report.size,
        file_hash: fileHash,
        created_at: createdAt,
        updated_at: createdAt,
      },
      { onConflict: "report_key" },
    );
    if (generatedResult.error) throw generatedResult.error;
    return;
  } catch (error) {
    console.warn("generated_reports sync failed, falling back to report_files", error);
  }

  const fileResult = await table("report_files").upsert(
    {
      file_name: report.fileName,
      file_hash: fileHash,
      report_date: firstDate,
      status: "done",
      processed_status: "done",
    },
    { onConflict: "file_hash" },
  );
  if (fileResult.error) throw fileResult.error;

  if (firstDate) {
    const reportResult = await table("reports").upsert(
      {
        date: firstDate,
        file_name: report.fileName,
      },
      { onConflict: "date,file_name" },
    );
    if (reportResult.error) throw reportResult.error;
  }
}

async function hashBlob(blob: Blob): Promise<string> {
  const buffer = await blob.arrayBuffer();
  const digest = await crypto.subtle.digest("SHA-256", buffer);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function firstValidDate(dates: string[]): string | null {
  return dates.find((date) => /^\d{4}-\d{2}-\d{2}$/.test(date)) ?? null;
}
