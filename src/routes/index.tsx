import { createFileRoute } from "@tanstack/react-router";
import { useCallback, useMemo, useRef, useState } from "react";
import {
  FileSpreadsheet,
  FileText,
  Sparkles,
  Upload,
  Download,
  CheckCircle2,
  AlertTriangle,
  Loader2,
  Camera,
  X,
  Database,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Progress } from "@/components/ui/progress";
import { toast } from "sonner";
import { Toaster } from "@/components/ui/sonner";
import { parsePdf, type PdfParseResult } from "@/lib/pdf-extractor";
import { processExcel, type ProcessReport } from "@/lib/excel-processor";
import { syncToDatabase, findProcessedHashes, type EmployeeInput } from "@/lib/db-sync";

export const Route = createFileRoute("/")({
  component: HomePage,
  head: () => ({
    meta: [
      { title: "neutropharmacallsummry.com" },
      {
        name: "description",
        content:
          "neutropharmacallsummry.com builds date-wise pharma selfie and total reports from PDFs into your Active Members Excel.",
      },
    ],
  }),
});

function HomePage() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [pdfFiles, setPdfFiles] = useState<File[]>([]);
  const [processing, setProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressLabel, setProgressLabel] = useState("");
  const [report, setReport] = useState<ProcessReport | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  const excelInputRef = useRef<HTMLInputElement>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);

  const onExcelChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) {
      setExcelFile(f);
      setReport(null);
      setDownloadUrl(null);
    }
  };

  const onPdfChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const list = e.target.files;
    if (!list) return;
    const incoming = Array.from(list).filter((f) => f.name.toLowerCase().endsWith(".pdf"));
    setPdfFiles((prev) => {
      const map = new Map(prev.map((p) => [p.name, p]));
      for (const f of incoming) map.set(f.name, f);
      return [...map.values()].sort((a, b) => a.name.localeCompare(b.name));
    });
    setReport(null);
    setDownloadUrl(null);
  };

  const removePdf = (name: string) => setPdfFiles((prev) => prev.filter((p) => p.name !== name));

  const sortedDates = useMemo(() => {
    const ds = pdfFiles
      .map((f) => f.name.match(/(\d{4}-\d{2}-\d{2})/)?.[1])
      .filter((x): x is string => !!x);
    return [...new Set(ds)].sort();
  }, [pdfFiles]);

  const onProcess = useCallback(async () => {
    if (!excelFile) {
      toast.error("Please upload your Active Members Excel file.");
      return;
    }
    if (pdfFiles.length === 0) {
      toast.error("Please upload at least one daily PDF report.");
      return;
    }

    setProcessing(true);
    setProgress(0);
    setReport(null);
    setDownloadUrl(null);

    try {
      // 1. Parse PDFs (with file-hash dedup against DB)
      const results: PdfParseResult[] = [];
      const skipped: string[] = [];
      const failed: { fileName: string; error: string }[] = [];

      // Pre-hash all files to check duplicates in one query
      setProgressLabel("Checking for duplicate uploads…");
      const fileHashes: { file: File; hash: string }[] = [];
      for (const f of pdfFiles) {
        try {
          const buf = await f.arrayBuffer();
          const digest = await crypto.subtle.digest("SHA-256", buf);
          const hex = [...new Uint8Array(digest)]
            .map((b) => b.toString(16).padStart(2, "0"))
            .join("");
          fileHashes.push({ file: f, hash: hex });
        } catch {
          failed.push({ fileName: f.name, error: "Unable to read file" });
        }
      }
      const processedSet = await findProcessedHashes(fileHashes.map((x) => x.hash));

      for (let i = 0; i < fileHashes.length; i++) {
        const { file: f, hash } = fileHashes[i];
        setProgressLabel(`Reading PDF ${i + 1}/${fileHashes.length}: ${f.name}`);
        setProgress(Math.round((i / fileHashes.length) * 70));
        if (processedSet.has(hash)) {
          skipped.push(f.name);
          // still parse so user gets results in this run, but we know it's a re-run
        }
        try {
          const r = await parsePdf(f);
          results.push(r);
        } catch (err) {
          console.error("PDF parse failed", f.name, err);
          failed.push({
            fileName: f.name,
            error: err instanceof Error ? err.message : "parse failed",
          });
        }
      }
      if (skipped.length) {
        toast.info(`${skipped.length} file(s) were already processed before — re-running.`);
      }
      if (failed.length) {
        toast.warning(`Skipped ${failed.length} unreadable PDF(s)`);
      }
      if (results.length === 0) throw new Error("No PDFs could be read.");

      // 2. Update Excel
      setProgressLabel("Updating your Excel file…");
      setProgress(80);
      const rep = await processExcel(excelFile, { pdfResults: results });
      if (rep.warnings.length) {
        toast.warning(rep.warnings[0]);
      }

      // 3. Sync to database (best-effort)
      setProgressLabel("Saving history to database…");
      setProgress(92);
      try {
        const employeeRows = await readEmployees(excelFile);
        const syncResult = await syncToDatabase({
          employees: employeeRows,
          pdfResults: results,
          failedFiles: failed,
        });
        rep.debug.totalRecordsInsertedUpdated = syncResult.syncedRecords;
        rep.debug.totalSkipped += syncResult.skippedRows;
        rep.debug.parsingErrors = failed.length + syncResult.parsingErrors;
      } catch (e) {
        console.warn("DB sync failed (non-fatal)", e);
        rep.debug.parsingErrors += 1;
      }

      // 4. Prepare download
      setProgress(100);
      setProgressLabel("Done!");
      const url = URL.createObjectURL(rep.blob);
      setDownloadUrl(url);
      setReport(rep);

      toast.success(
        `Report ready! ${rep.matchedEmployees}/${rep.totalEmployees} employees with selfies across ${rep.dates.length} day(s).`,
      );
    } catch (err) {
      console.error(err);
      toast.error(err instanceof Error ? err.message : "Something went wrong while processing.");
    } finally {
      setProcessing(false);
    }
  }, [excelFile, pdfFiles]);

  const triggerDownload = () => {
    if (!downloadUrl || !report) return;
    const a = document.createElement("a");
    a.href = downloadUrl;
    a.download = report.fileName.replace(/\.xlsx$/i, "") + " - Updated.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div className="min-h-screen bg-background">
      <Toaster richColors position="top-center" />

      {/* Hero */}
      <header className="relative overflow-hidden border-b border-border">
        <div
          className="absolute inset-0 opacity-90"
          style={{ background: "var(--gradient-hero)" }}
        />
        <div className="relative mx-auto max-w-6xl px-6 py-14 text-primary-foreground">
          <div className="flex items-center gap-2 text-sm font-medium opacity-90">
            <Sparkles className="h-4 w-4" />
            One-click automated reporting
          </div>
          <h1 className="mt-3 text-4xl font-bold tracking-tight md:text-5xl">
            neutropharmacallsummry.com
          </h1>
          <p className="mt-4 max-w-2xl text-base opacity-90 md:text-lg">
            Upload your Active Members Excel and the daily PDF reports. We fill the same Excel with
            date-wise selfie counts, totals and color-coded performance — in seconds.
          </p>
          <div className="mt-6 flex flex-wrap gap-2 text-xs">
            <Pill icon={<Camera className="h-3 w-3" />} label="Smart selfie extraction" />
            <Pill icon={<FileSpreadsheet className="h-3 w-3" />} label="Updates same Excel" />
            <Pill icon={<Database className="h-3 w-3" />} label="Saved history" />
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-6 py-10">
        {/* Step 1 + 2 */}
        <div className="grid gap-6 md:grid-cols-2">
          <UploadCard
            step={1}
            title="Active Members Excel"
            subtitle="Your master file. Will be updated in place."
            icon={<FileSpreadsheet className="h-5 w-5" />}
            accept=".xlsx,.xls"
            onClick={() => excelInputRef.current?.click()}
          >
            <input
              ref={excelInputRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={onExcelChange}
            />
            {excelFile ? (
              <FilePill name={excelFile.name} onRemove={() => setExcelFile(null)} color="primary" />
            ) : (
              <EmptyHint label="Choose .xlsx file" />
            )}
          </UploadCard>

          <UploadCard
            step={2}
            title="Daily PDF Reports"
            subtitle="File names like 2026-04-24.pdf"
            icon={<FileText className="h-5 w-5" />}
            accept=".pdf"
            multiple
            onClick={() => pdfInputRef.current?.click()}
          >
            <input
              ref={pdfInputRef}
              type="file"
              accept=".pdf"
              multiple
              className="hidden"
              onChange={onPdfChange}
            />
            {pdfFiles.length === 0 ? (
              <EmptyHint label="Choose one or many PDFs" />
            ) : (
              <div className="space-y-2">
                <div className="text-xs text-muted-foreground">
                  {pdfFiles.length} file(s) · {sortedDates.length} unique date(s)
                </div>
                <div className="max-h-40 space-y-1.5 overflow-y-auto pr-1">
                  {pdfFiles.map((f) => (
                    <FilePill
                      key={f.name}
                      name={f.name}
                      onRemove={() => removePdf(f.name)}
                      color="accent"
                    />
                  ))}
                </div>
              </div>
            )}
          </UploadCard>
        </div>

        {/* Process button */}
        <Card className="mt-6 border-border bg-card p-6 shadow-[var(--shadow-soft)]">
          <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-center">
            <div>
              <h2 className="text-lg font-semibold text-foreground">
                Step 3 — Generate the report
              </h2>
              <p className="text-sm text-muted-foreground">
                We&apos;ll read every PDF, match exact names to your sheet, fill date columns,
                calculate totals and color-code performers without changing row order.
              </p>
            </div>
            <Button
              size="lg"
              onClick={onProcess}
              disabled={processing || !excelFile || pdfFiles.length === 0}
              className="min-w-48 shadow-[var(--shadow-elegant)]"
            >
              {processing ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Processing…
                </>
              ) : (
                <>
                  <Sparkles className="mr-2 h-4 w-4" />
                  Process & Update Excel
                </>
              )}
            </Button>
          </div>

          {processing && (
            <div className="mt-5 space-y-2">
              <Progress value={progress} />
              <div className="text-xs text-muted-foreground">{progressLabel}</div>
            </div>
          )}
        </Card>

        {/* Report */}
        {report && (
          <Card className="mt-6 border-border bg-card p-6 shadow-[var(--shadow-soft)]">
            <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-center">
              <div className="flex items-start gap-3">
                <div className="flex h-10 w-10 items-center justify-center rounded-full bg-success/15 text-success">
                  <CheckCircle2 className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="text-lg font-semibold text-foreground">Report ready</h3>
                  <p className="text-sm text-muted-foreground">
                    Updated <span className="font-medium">{report.fileName}</span> —{" "}
                    {report.totalEmployees} employees · {report.dates.length} dates ·{" "}
                    {report.matchedEmployees} active performers
                  </p>
                </div>
              </div>
              <Button onClick={triggerDownload} size="lg" variant="default">
                <Download className="mr-2 h-4 w-4" />
                Download Updated Excel
              </Button>
            </div>

            <div className="mt-6 grid gap-4 md:grid-cols-3">
              <Stat label="Employees" value={report.totalEmployees} />
              <Stat label="Days covered" value={report.dates.length} />
              <Stat
                label="Unmatched PDF names"
                value={report.unmatchedNames.length}
                tone={report.unmatchedNames.length ? "warning" : "default"}
              />
            </div>

            <div className="mt-6 grid gap-3 rounded-md border border-border bg-muted/20 p-4 md:grid-cols-3">
              <Stat label="PDFs uploaded" value={report.debug.totalPdfsUploaded} />
              <Stat label="Detected rows" value={report.debug.totalEmployeesDetected} />
              <Stat label="Matched rows" value={report.debug.totalMatched} />
              <Stat
                label="Skipped rows"
                value={report.debug.totalSkipped}
                tone={report.debug.totalSkipped ? "warning" : "default"}
              />
              <Stat label="DB upserts" value={report.debug.totalRecordsInsertedUpdated} />
              <Stat
                label="Parsing errors"
                value={report.debug.parsingErrors}
                tone={report.debug.parsingErrors ? "warning" : "default"}
              />
            </div>

            {report.warnings.length > 0 && (
              <div className="mt-6 rounded-md border border-warning/40 bg-warning/10 p-4">
                <div className="flex items-center gap-2 text-sm font-semibold text-warning-foreground">
                  <AlertTriangle className="h-4 w-4" />
                  PDF parsing incomplete, please review.
                </div>
                <p className="mt-1 text-xs text-muted-foreground">{report.warnings[0]}</p>
              </div>
            )}

            {report.preview.length > 0 && (
              <div className="mt-6">
                <div className="mb-2 text-sm font-semibold text-foreground">Top performers</div>
                <div className="space-y-1.5">
                  {report.preview.map((p, i) => (
                    <div
                      key={p.name}
                      className="flex items-center justify-between rounded-md border border-border bg-muted/40 px-3 py-2 text-sm"
                    >
                      <span className="flex items-center gap-3">
                        <span className="inline-flex h-6 w-6 items-center justify-center rounded-full bg-primary/10 text-xs font-bold text-primary">
                          {i + 1}
                        </span>
                        {p.name}
                      </span>
                      <span className="font-semibold text-foreground">{p.total}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {report.unmatchedNames.length > 0 && (
              <div className="mt-6 rounded-md border border-warning/40 bg-warning/10 p-4">
                <div className="mb-1 flex items-center gap-2 text-sm font-semibold text-warning-foreground">
                  <AlertTriangle className="h-4 w-4" />
                  {report.unmatchedNames.length} name(s) in PDFs didn&apos;t match your Active
                  Members sheet
                </div>
                <p className="text-xs text-muted-foreground">
                  These were ignored (resigned / not active). Examples:{" "}
                  {report.unmatchedNames.slice(0, 6).join(", ")}
                  {report.unmatchedNames.length > 6 ? "…" : ""}
                </p>
              </div>
            )}
          </Card>
        )}
      </main>
    </div>
  );
}

/* ---------- helpers (read employee list once) ---------- */

async function readEmployees(excelFile: File): Promise<EmployeeInput[]> {
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await excelFile.arrayBuffer());
  for (const ws of wb.worksheets) {
    const maxScan = Math.min(10, ws.rowCount || 10);
    for (let r = 1; r <= maxScan; r++) {
      const row = ws.getRow(r);
      let nameCol = 0,
        codeCol = 0,
        regionCol = 0,
        cityCol = 0,
        designationCol = 0;
      for (let c = 1; c <= 40; c++) {
        const v = String(row.getCell(c).value ?? "")
          .trim()
          .toLowerCase();
        if (!v) continue;
        if (!regionCol && ["region", "zone", "area"].includes(v)) regionCol = c;
        if (!nameCol && v.includes("name") && !v.includes("file")) nameCol = c;
        if (!codeCol && (/\bcode\b/.test(v) || /\bemp.*id\b/.test(v) || v === "id")) codeCol = c;
        if (!cityCol && ["city", "town", "territory", "headquarter", "hq"].includes(v)) cityCol = c;
        if (!designationCol && ["designation", "desig", "position", "title"].includes(v)) {
          designationCol = c;
        }
      }
      if (nameCol) {
        const out: EmployeeInput[] = [];
        for (let rr = r + 1; rr <= (ws.rowCount || r + 500); rr++) {
          const name = String(ws.getRow(rr).getCell(nameCol).value ?? "").trim();
          if (!name) continue;
          const sheetRow = ws.getRow(rr);
          const code = codeCol
            ? String(sheetRow.getCell(codeCol).value ?? "")
                .trim()
                .toUpperCase() || null
            : null;
          out.push({
            name,
            code,
            region: regionCol
              ? String(sheetRow.getCell(regionCol).value ?? "").trim() || null
              : null,
            city: cityCol ? String(sheetRow.getCell(cityCol).value ?? "").trim() || null : null,
            designation: designationCol
              ? String(sheetRow.getCell(designationCol).value ?? "").trim() || null
              : null,
            originalOrder: out.length + 1,
          });
        }
        if (out.length) return out;
      }
    }
  }
  return [];
}

/* ---------- presentational ---------- */

function UploadCard({
  step,
  title,
  subtitle,
  icon,
  onClick,
  children,
}: {
  step: number;
  title: string;
  subtitle: string;
  icon: React.ReactNode;
  accept: string;
  multiple?: boolean;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <Card className="border-border bg-card p-6 shadow-[var(--shadow-soft)]">
      <div className="mb-4 flex items-start justify-between">
        <div>
          <div className="text-xs font-bold uppercase tracking-wider text-primary">Step {step}</div>
          <h2 className="mt-1 flex items-center gap-2 text-lg font-semibold text-foreground">
            <span className="flex h-7 w-7 items-center justify-center rounded-md bg-primary/10 text-primary">
              {icon}
            </span>
            {title}
          </h2>
          <p className="text-sm text-muted-foreground">{subtitle}</p>
        </div>
        <Button variant="outline" size="sm" onClick={onClick}>
          <Upload className="mr-1.5 h-3.5 w-3.5" />
          Upload
        </Button>
      </div>
      <div className="rounded-md border border-dashed border-border bg-muted/30 p-3">
        {children}
      </div>
    </Card>
  );
}

function EmptyHint({ label }: { label: string }) {
  return <div className="py-4 text-center text-sm text-muted-foreground">{label}</div>;
}

function FilePill({
  name,
  onRemove,
  color,
}: {
  name: string;
  onRemove: () => void;
  color: "primary" | "accent";
}) {
  return (
    <div
      className={`flex items-center justify-between gap-2 rounded-md border px-3 py-2 text-sm ${
        color === "primary"
          ? "border-primary/30 bg-primary/5 text-foreground"
          : "border-accent bg-accent/40 text-accent-foreground"
      }`}
    >
      <span className="truncate">{name}</span>
      <button
        onClick={onRemove}
        className="rounded p-0.5 text-muted-foreground transition-colors hover:bg-background hover:text-destructive"
        aria-label={`Remove ${name}`}
      >
        <X className="h-3.5 w-3.5" />
      </button>
    </div>
  );
}

function Pill({ icon, label }: { icon: React.ReactNode; label: string }) {
  return (
    <span className="inline-flex items-center gap-1.5 rounded-full bg-white/15 px-3 py-1 backdrop-blur-sm">
      {icon}
      {label}
    </span>
  );
}

function Stat({
  label,
  value,
  tone = "default",
}: {
  label: string;
  value: number;
  tone?: "default" | "warning";
}) {
  return (
    <div
      className={`rounded-md border p-4 ${
        tone === "warning" ? "border-warning/40 bg-warning/10" : "border-border bg-muted/30"
      }`}
    >
      <div className="text-xs uppercase tracking-wider text-muted-foreground">{label}</div>
      <div className="mt-1 text-2xl font-bold text-foreground">{value}</div>
    </div>
  );
}
