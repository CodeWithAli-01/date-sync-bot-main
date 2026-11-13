import { useRef, useState, type ReactNode } from "react";
import { Download, Eye, Loader2, FileText, FileSpreadsheet, Sparkles, ShieldCheck, Check } from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { UploadCard, EmptyHint, FilePill } from "@/components/report-upload-card";
import { processAvgPerDayReport, type AvgPerDayResult } from "@/lib/avg-per-day-processor";

export function AvgPerDayReport({ buildPreview, onGenerated }: {
  buildPreview: (blob: Blob) => Promise<ReactNode>;
  onGenerated: (result: AvgPerDayResult) => Promise<void>;
}) {
  const [sample, setSample] = useState<File | null>(null);
  const [source, setSource] = useState<File | null>(null);
  const [result, setResult] = useState<AvgPerDayResult | null>(null);
  const [processing, setProcessing] = useState(false);
  const [preview, setPreview] = useState<ReactNode>(null);
  const [previewLoading, setPreviewLoading] = useState(false);
  const previewRef = useRef<HTMLDivElement>(null);

  const resetResult = () => { setResult(null); setPreview(null); };
  const generate = async () => {
    if (!sample || !source) return;
    setProcessing(true);
    resetResult();
    try {
      const report = await processAvgPerDayReport(sample, source);
      setResult(report);
      toast.success(`Avg. Per Day report ready: ${report.matchedEmployees}/${report.totalEmployees} employees matched.`);
      try { await onGenerated(report); }
      catch { toast.warning("Report ready, but it could not be saved to History."); }
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "Unable to generate Avg. Per Day report.");
    } finally { setProcessing(false); }
  };
  const openPreview = async () => {
    if (!result) return;
    setPreviewLoading(true);
    try {
      setPreview(await buildPreview(result.blob));
      setTimeout(() => previewRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }), 0);
    } catch { toast.error("Unable to open the report preview."); }
    finally { setPreviewLoading(false); }
  };
  const download = () => {
    if (!result) return;
    const url = URL.createObjectURL(result.blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = result.fileName;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
  return (
    <main className="grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_260px]">
      <div className="min-w-0 space-y-5">
      <div className="grid items-stretch gap-4 md:grid-cols-2">
        {([
          { id: "avg-per-day-sample", title: "Sample Employee Excel", file: sample, setFile: setSample },
          { id: "avg-per-day-source", title: "Avg Report Excel", file: source, setFile: setSource },
        ]).map(({ id, title, file, setFile }, index) => (
          <UploadCard key={id} step={index + 1} title={title}
            subtitle={file ? "File selected" : index === 0 ? "Upload your employee sample and report format" : "Upload the employee avg report Excel"}
            icon={index === 0 ? <FileText className="h-5 w-5" /> : <FileSpreadsheet className="h-5 w-5" />}
            accept=".xlsx,.xlsm" ready={Boolean(file)}
            onClick={() => { if (!processing) document.getElementById(id)?.click(); }}>
            <input id={id} type="file" accept=".xlsx,.xlsm" disabled={processing}
              className="hidden"
              onChange={(event) => { setFile(event.target.files?.[0] ?? null); resetResult(); }} />
            {file ? <FilePill name={file.name} color="primary" onRemove={() => {
              if (processing) return;
              setFile(null); resetResult();
              const input = document.getElementById(id) as HTMLInputElement | null;
              if (input) input.value = "";
            }} /> : <EmptyHint icon={<FileSpreadsheet className="h-5 w-5" />} label={index === 0 ? "Choose sample file" : "Choose avg report file"} />}
          </UploadCard>
        ))}
      </div>
      <Card className="flex flex-col gap-4 border-border bg-card p-6 shadow-[var(--shadow-soft)] md:flex-row md:items-center md:justify-between">
        <div className="min-w-0">
          <div className="text-xs font-bold uppercase text-primary">Step 3</div>
          <h2 className="mt-1 text-lg font-semibold text-foreground">Generate Avg. Per Day report</h2>
          <p className="text-sm text-muted-foreground">Match employee IDs and copy Morning Calls, Evening Calls, Total Calls, Avg. Call Per Day., and Field Days into your sample.</p>
        </div>
        <Button size="lg" className="w-full shrink-0 md:w-auto" disabled={!sample || !source || processing} onClick={generate}>
          {processing ? <Loader2 className="animate-spin" /> : <Sparkles />}
          {processing ? "Processing..." : "Generate Avg. Per Day Report"}
        </Button>
      </Card>
      {result && <Card className="space-y-4 p-6">
        <h2 className="font-semibold">Avg. Per Day report ready</h2>
        <p className="text-sm">{result.matchedEmployees}/{result.totalEmployees} employees matched by ID.</p>
        {result.warnings.map((warning) => <p key={warning} className="text-sm text-muted-foreground">{warning}</p>)}
        <div className="flex flex-wrap gap-3">
          <Button variant="outline" onClick={openPreview} disabled={previewLoading}>{previewLoading ? <Loader2 className="animate-spin" /> : <Eye />}Preview</Button>
          <Button onClick={download}><Download />Download</Button>
        </div>
      </Card>}
      {(preview || previewLoading) && <Card ref={previewRef} className="space-y-3 p-5">
        <h2 className="font-semibold">Report preview</h2>
        <div className="max-h-[70vh] overflow-auto">{previewLoading ? "Opening preview..." : preview}</div>
      </Card>}
      </div>
      <aside className="space-y-4">
        <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
          <div className="mb-4 flex items-center justify-between gap-3">
            <div>
              <div className="text-xs font-bold uppercase text-primary">Run status</div>
              <h2 className="mt-1 text-base font-semibold text-foreground">Avg. Per Day workspace</h2>
            </div>
            <ShieldCheck className="h-5 w-5 shrink-0 text-primary" />
          </div>
          <div className="space-y-3" aria-live="polite">
            {[
              { done: Boolean(sample), label: "Sample employee file attached" },
              { done: Boolean(source), label: "Avg report file attached" },
              { done: Boolean(result), label: "Report generated" },
              { done: Boolean(preview), label: "Preview opened" },
            ].map(({ done, label }) => (
              <div key={label} className="flex min-w-0 items-center gap-2 text-sm">
                <span className={`flex h-5 w-5 shrink-0 items-center justify-center rounded-full border ${done
                  ? "border-success bg-success text-success-foreground"
                  : "border-white/40 bg-muted text-muted-foreground shadow-[var(--shadow-inset)]"}`}>
                  {done && <Check className="h-3 w-3" />}
                </span>
                <span className={done ? "min-w-0 text-foreground" : "min-w-0 text-muted-foreground"}>{label}</span>
              </div>
            ))}
            {processing && <div className="flex items-center gap-2 text-sm text-primary"><Loader2 className="h-4 w-4 animate-spin" />Generating report...</div>}
          </div>
        </Card>
      </aside>
    </main>
  );
}
