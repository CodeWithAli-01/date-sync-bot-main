// Sync processed results to Lovable Cloud:
//  - employees (active/inactive based on sheet)
//  - report_files (file_hash dedup; returns map of fileName -> id)
//  - selfie_records (one per employee+date; updated on conflict)
//  - error_logs
import { supabase } from "@/integrations/supabase/client";
import type { PdfParseResult } from "./pdf-extractor";

export interface SyncInput {
  employees: { name: string; code?: string | null }[];
  pdfResults: PdfParseResult[];
  failedFiles?: { fileName: string; error: string }[];
}

export async function syncToDatabase(opts: SyncInput) {
  const { employees, pdfResults, failedFiles = [] } = opts;

  // 1. Upsert employees
  if (employees.length) {
    await supabase
      .from("employees")
      .upsert(
        employees.map((e) => ({
          name: e.name,
          employee_code: e.code ?? null,
          status: "active",
        })),
        { onConflict: "name" }
      );
  }
  // mark missing as inactive
  const { data: existing } = await supabase.from("employees").select("name");
  const sheetSet = new Set(employees.map((e) => e.name.toLowerCase()));
  const toDeactivate = (existing ?? [])
    .map((e) => e.name)
    .filter((n) => !sheetSet.has(n.toLowerCase()));
  if (toDeactivate.length) {
    await supabase.from("employees").update({ status: "inactive" }).in("name", toDeactivate);
  }

  // 2. Insert report_files (dedup by hash). Map fileName -> id for FK use.
  const fileIdByName = new Map<string, string>();
  for (const r of pdfResults) {
    // Try insert; if hash conflict, fetch existing
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
      // Already existed — fetch by hash
      const { data: existing2 } = await supabase
        .from("report_files")
        .select("id")
        .eq("file_hash", r.fileHash)
        .maybeSingle();
      if (existing2?.id) fileIdByName.set(r.fileName, existing2.id);
      if (insErr && !existing2) console.warn("report_files insert failed", insErr);
    }
    // also keep legacy reports table populated
    await supabase
      .from("reports")
      .upsert({ date: r.date, file_name: r.fileName }, { onConflict: "date,file_name", ignoreDuplicates: true });
  }

  // 3. Upsert selfie_records (only employees in sheet)
  const records: {
    employee_name: string;
    employee_code: string | null;
    date: string;
    count: number;
    source_file_id: string | null;
  }[] = [];
  const empMap = new Map(employees.map((e) => [e.name.toLowerCase(), e]));
  const seen = new Set<string>();
  for (const r of pdfResults) {
    for (const row of r.rows) {
      const emp = empMap.get(row.name.toLowerCase());
      if (!emp) continue;
      const key = `${emp.name}|${r.date}`;
      if (seen.has(key)) continue;
      seen.add(key);
      records.push({
        employee_name: emp.name,
        employee_code: emp.code ?? row.code ?? null,
        date: r.date,
        count: row.count,
        source_file_id: fileIdByName.get(r.fileName) ?? null,
      });
    }
  }
  if (records.length) {
    for (let i = 0; i < records.length; i += 500) {
      await supabase
        .from("selfie_records")
        .upsert(records.slice(i, i + 500), { onConflict: "employee_name,date" });
    }
  }

  // 4. Error logs
  if (failedFiles.length) {
    await supabase
      .from("error_logs")
      .insert(failedFiles.map((f) => ({ file_name: f.fileName, error_message: f.error })));
  }

  return { syncedRecords: records.length, deactivated: toDeactivate.length };
}

// Check which file hashes are already processed — used to skip duplicate uploads.
export async function findProcessedHashes(hashes: string[]): Promise<Set<string>> {
  if (!hashes.length) return new Set();
  const { data } = await supabase
    .from("report_files")
    .select("file_hash")
    .in("file_hash", hashes);
  return new Set((data ?? []).map((d: any) => d.file_hash));
}
