// Sync processed results to Lovable Cloud without destructive rewrites.
// Employee/date records are matched by code first, then exact/fuzzy name.
import { supabase } from "@/integrations/supabase/client";
import type { PdfParseResult } from "./pdf-extractor";

export interface SyncInput {
  employees: { name: string; code?: string | null }[];
  pdfResults: PdfParseResult[];
  failedFiles?: { fileName: string; error: string }[];
}

interface EmployeeRow {
  name: string;
  code: string | null;
  nameKey: string;
  fuzzy: string;
}

interface SelfieRecordInput {
  employee_name: string;
  employee_code: string | null;
  date: string;
  count: number;
  source_file_id: string | null;
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

function buildEmployeeIndexes(employees: { name: string; code?: string | null }[]) {
  const rows: EmployeeRow[] = employees.map((e) => ({
    name: e.name,
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

export async function syncToDatabase(opts: SyncInput) {
  const { employees, pdfResults, failedFiles = [] } = opts;
  const indexes = buildEmployeeIndexes(employees);

  // 1. Upsert employees. This is additive/safe: no delete + insert cycle.
  if (indexes.rows.length) {
    await supabase.from("employees").upsert(
      indexes.rows.map((e) => ({
        name: e.name,
        employee_code: e.code,
        status: "active",
      })),
      { onConflict: "name" },
    );
  }

  // Mark missing employees inactive, but do not delete rows or historical records.
  const { data: existing } = await supabase.from("employees").select("name");
  const sheetSet = new Set(indexes.rows.map((e) => e.nameKey));
  const toDeactivate = (existing ?? [])
    .map((e) => e.name)
    .filter((name) => !sheetSet.has(normalize(name)));
  if (toDeactivate.length) {
    await supabase.from("employees").update({ status: "inactive" }).in("name", toDeactivate);
  }

  // 2. Insert report_files with hash dedup. Map fileName -> id for FK use.
  const fileIdByName = new Map<string, string>();
  for (const r of pdfResults) {
    const { data: ins, error: insErr } = await supabase
      .from("report_files")
      .insert({
        file_name: r.fileName,
        file_hash: r.fileHash,
        report_date: r.date,
        processed_status: "done",
      })
      .select("id")
      .maybeSingle();

    if (ins?.id) {
      fileIdByName.set(r.fileName, ins.id);
    } else {
      const { data: existingFile } = await supabase
        .from("report_files")
        .select("id")
        .eq("file_hash", r.fileHash)
        .maybeSingle();
      if (existingFile?.id) fileIdByName.set(r.fileName, existingFile.id);
      if (insErr && !existingFile) console.warn("report_files insert failed", insErr);
    }

    await supabase
      .from("reports")
      .upsert(
        { date: r.date, file_name: r.fileName },
        { onConflict: "date,file_name", ignoreDuplicates: true },
      );
  }

  // 3. Build unique selfie_records from PDFs. Keep the strongest count if the
  // same employee/day appears more than once, so a parsed 0 cannot win over 11.
  const recordByKey = new Map<string, SelfieRecordInput>();
  for (const pdf of pdfResults) {
    for (const row of pdf.rows) {
      const employee = findEmployee(row, indexes);
      if (!employee) continue;

      const key = `${employee.code ?? employee.nameKey}|${pdf.date}`;
      const next: SelfieRecordInput = {
        employee_name: employee.name,
        employee_code: employee.code ?? row.code ?? null,
        date: pdf.date,
        count: row.count,
        source_file_id: fileIdByName.get(pdf.fileName) ?? null,
      };
      const existingRecord = recordByKey.get(key);
      if (!existingRecord || next.count > existingRecord.count) {
        recordByKey.set(key, next);
      }
    }
  }

  const candidateRecords = [...recordByKey.values()];
  const records = await removeUnsafeZeroOverwrites(candidateRecords);
  const codeRecords = records.filter((r) => r.employee_code);
  const nameOnlyRecords = records.filter((r) => !r.employee_code);

  for (let i = 0; i < codeRecords.length; i += 500) {
    await supabase
      .from("selfie_records")
      .upsert(codeRecords.slice(i, i + 500), { onConflict: "employee_code,date" });
  }

  for (let i = 0; i < nameOnlyRecords.length; i += 500) {
    await supabase
      .from("selfie_records")
      .upsert(nameOnlyRecords.slice(i, i + 500), { onConflict: "employee_name,date" });
  }

  // 4. Error logs are append-only.
  if (failedFiles.length) {
    await supabase
      .from("error_logs")
      .insert(failedFiles.map((f) => ({ file_name: f.fileName, error_message: f.error })));
  }

  return { syncedRecords: records.length, deactivated: toDeactivate.length };
}

async function removeUnsafeZeroOverwrites(
  records: SelfieRecordInput[],
): Promise<SelfieRecordInput[]> {
  const zeroRecords = records.filter((r) => r.count === 0);
  if (!zeroRecords.length) return records;

  const dates = [...new Set(zeroRecords.map((r) => r.date))];
  const codes = [
    ...new Set(zeroRecords.map((r) => r.employee_code).filter((code): code is string => !!code)),
  ];
  const names = [...new Set(zeroRecords.map((r) => r.employee_name))];

  const existingByCode = new Set<string>();
  const existingByName = new Set<string>();

  if (dates.length && codes.length) {
    const { data } = await supabase
      .from("selfie_records")
      .select("employee_code,date,count")
      .in("date", dates)
      .in("employee_code", codes)
      .gt("count", 0);
    for (const row of data ?? []) {
      if (row.employee_code) existingByCode.add(`${row.employee_code}|${row.date}`);
    }
  }

  if (dates.length && names.length) {
    const { data } = await supabase
      .from("selfie_records")
      .select("employee_name,date,count")
      .in("date", dates)
      .in("employee_name", names)
      .gt("count", 0);
    for (const row of data ?? []) {
      existingByName.add(`${row.employee_name}|${row.date}`);
    }
  }

  return records.filter((record) => {
    if (record.count !== 0) return true;
    if (record.employee_code && existingByCode.has(`${record.employee_code}|${record.date}`))
      return false;
    return !existingByName.has(`${record.employee_name}|${record.date}`);
  });
}

// Check which file hashes are already processed; used to skip duplicate uploads.
export async function findProcessedHashes(hashes: string[]): Promise<Set<string>> {
  if (!hashes.length) return new Set();
  const { data } = await supabase.from("report_files").select("file_hash").in("file_hash", hashes);
  return new Set((data ?? []).map((d) => d.file_hash));
}
