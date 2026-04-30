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

interface EmployeeRow extends EmployeeInput {
  nameKey: string;
  fuzzy: string;
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

function similarity(a: string, b: string): number {
  if (a === b) return 1;
  if (!a.length || !b.length) return 0;
  const m = a.length;
  const n = b.length;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const c = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + c);
    }
  }
  return 1 - dp[m][n] / Math.max(m, n);
}

function buildEmployeeIndexes(employees: EmployeeInput[]) {
  const rows: EmployeeRow[] = employees.map((e) => ({
    ...e,
    code: e.code?.trim().toUpperCase() || null,
    nameKey: normalize(e.name),
    fuzzy: fuzzyKey(e.name),
  }));
  const byCode = new Map<string, EmployeeRow>();
  const byName = new Map<string, EmployeeRow>();
  for (const row of rows) {
    if (row.code) byCode.set(row.code, row);
    byName.set(row.nameKey, row);
  }
  return { rows, byCode, byName };
}

function findEmployee(
  row: { code: string | null; name: string },
  indexes: ReturnType<typeof buildEmployeeIndexes>,
): EmployeeRow | null {
  if (row.code) {
    const byCode = indexes.byCode.get(row.code.toUpperCase());
    if (byCode) return byCode;
  }

  const byName = indexes.byName.get(normalize(row.name));
  if (byName) return byName;

  const target = fuzzyKey(row.name);
  let best: EmployeeRow | null = null;
  let bestScore = 0;
  for (const employee of indexes.rows) {
    const score = similarity(target, employee.fuzzy);
    if (score > bestScore) {
      bestScore = score;
      best = employee;
    }
  }
  return bestScore >= 0.85 ? best : null;
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
        selfie_text: `${selfieCount} selfies with locations in grp`,
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
      selfie_text: `${selfieCount} selfies with locations in grp`,
      total_count: totalCount,
    };
  });
}

export async function findProcessedHashes(hashes: string[]): Promise<Set<string>> {
  if (!hashes.length) return new Set();
  const { data } = await table("report_files").select("file_hash").in("file_hash", hashes);
  return new Set((data ?? []).map((d) => String(d.file_hash)));
}
