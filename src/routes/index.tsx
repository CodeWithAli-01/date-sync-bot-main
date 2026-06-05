import { createFileRoute } from "@tanstack/react-router";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Activity,
  BarChart3,
  CalendarDays,
  Check,
  FileSpreadsheet,
  FileText,
  History as HistoryIcon,
  Moon,
  Palette,
  Sparkles,
  Sun,
  Upload,
  Download,
  Eye,
  Save,
  CheckCircle2,
  AlertTriangle,
  Loader2,
  X,
  ShieldCheck,
  TableProperties,
  Trash2,
  Lock,
  LogOut,
  Mail,
  UserPlus,
  Home,
  UserCircle,
  Settings,
  Clock3,
  Database,
  LineChart,
  PieChart,
  Users,
  Zap,
} from "lucide-react";
import type { User } from "@supabase/supabase-js";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Progress } from "@/components/ui/progress";
import { Avatar, AvatarFallback, AvatarImage } from "@/components/ui/avatar";
import { toast } from "sonner";
import { Toaster } from "@/components/ui/sonner";
import { supabase } from "@/integrations/supabase/client";
import { parsePdf, type PdfParseResult } from "@/lib/pdf-extractor";
import { processExcel, type ProcessReport } from "@/lib/excel-processor";
import {
  processBulkDailyReports,
  processDailyReport,
  type BulkDailyReportResult,
  type DailyReportResult,
} from "@/lib/daily-report-processor";
import {
  processDoctorCoverageReport,
  type DoctorCoverageResult,
} from "@/lib/doctor-coverage-processor";
import {
  processMonthlyPlannedReport,
  type MonthlyPlannedResult,
} from "@/lib/monthly-planned-processor";
import {
  syncToDatabase,
  syncGeneratedReportToDatabase,
  syncAuthUserToDatabase,
  listGeneratedReportsFromDatabase,
  getGeneratedReportFromDatabase,
  deleteGeneratedReportFromDatabase,
  findProcessedHashes,
  type EmployeeInput,
} from "@/lib/db-sync";
import { enforceDeviceLimit, revokeCurrentDeviceSession } from "@/lib/device-auth";

interface PreviewCell {
  key: string;
  rowNumber: number;
  colNumber: number;
  value: string;
  style: React.CSSProperties;
}

interface SheetPreview {
  name: string;
  sheetName: string;
  rows: PreviewCell[][];
}

type ActiveModule =
  | "monthly-report"
  | "daily-report"
  | "doctor-coverage"
  | "monthly-planned"
  | "history"
  | "profile"
  | null;

interface ReportHistoryItem {
  id: string;
  userId?: string;
  fileName: string;
  createdAt: string;
  reportType?: string;
  dates: string[];
  pdfCount: number;
  totalEmployees: number;
  matchedEmployees: number;
  size: number;
}

interface StoredReportHistoryItem extends ReportHistoryItem {
  blob: Blob;
}

const APP_NAME = "Reporting Management";

const DASHBOARD_LINES = [
  "Work smarter with clean reports and reliable matching.",
  "Keep every report organized, accurate, and ready to share.",
  "Choose a workflow and let the software handle the details.",
  "Daily and monthly reports, prepared with confidence.",
  "Your reporting workspace is ready.",
];

const DEFAULT_THEME_COLOR = "#0b6f6a";
const THEME_COLORS = [
  "#0b6f6a",
  "#0f766e",
  "#2563eb",
  "#7c3aed",
  "#db2777",
  "#e43100",
  "#ea580c",
  "#ca8a04",
  "#16a34a",
  "#111827",
];

export const Route = createFileRoute("/")({
  component: HomePage,
  head: () => ({
    meta: [
      { title: APP_NAME },
      {
        name: "description",
        content:
          "Reporting Management builds date-wise pharma selfie and total reports from PDFs into your Active Members Excel.",
      },
    ],
  }),
});

function HomePage() {
  const [authLoading, setAuthLoading] = useState(true);
  const [authUser, setAuthUser] = useState<User | null>(null);
  const [signingOut, setSigningOut] = useState(false);
  const [profilePhoto, setProfilePhoto] = useState<string | null>(null);
  const [activeModule, setActiveModule] = useState<ActiveModule>(null);
  const [dashboardLine, setDashboardLine] = useState(DASHBOARD_LINES[0]);
  const [themeColor, setThemeColor] = useState(DEFAULT_THEME_COLOR);
  const [themeMode, setThemeMode] = useState<"light" | "dark">("light");
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [pdfFiles, setPdfFiles] = useState<File[]>([]);
  const [processing, setProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressLabel, setProgressLabel] = useState("");
  const [report, setReport] = useState<ProcessReport | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [sheetPreview, setSheetPreview] = useState<SheetPreview | null>(null);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewDirty, setPreviewDirty] = useState(false);
  const [previewSaving, setPreviewSaving] = useState(false);
  const [historyItems, setHistoryItems] = useState<ReportHistoryItem[]>([]);
  const [historyLoading, setHistoryLoading] = useState(false);
  const [historyPreview, setHistoryPreview] = useState<SheetPreview | null>(null);
  const [historyPreviewLoading, setHistoryPreviewLoading] = useState(false);
  const [callLogFiles, setCallLogFiles] = useState<File[]>([]);
  const [dailyTemplateFile, setDailyTemplateFile] = useState<File | null>(null);
  const [dailyProcessing, setDailyProcessing] = useState(false);
  const [dailyReport, setDailyReport] = useState<DailyReportResult | null>(null);
  const [bulkDailyReport, setBulkDailyReport] = useState<BulkDailyReportResult | null>(null);
  const [bulkDailyProgress, setBulkDailyProgress] = useState(0);
  const [bulkDailyProgressLabel, setBulkDailyProgressLabel] = useState("");
  const [dailyPreview, setDailyPreview] = useState<SheetPreview | null>(null);
  const [dailyPreviewLoading, setDailyPreviewLoading] = useState(false);
  const [coverageSourceFile, setCoverageSourceFile] = useState<File | null>(null);
  const [coverageTemplateFile, setCoverageTemplateFile] = useState<File | null>(null);
  const [coverageProcessing, setCoverageProcessing] = useState(false);
  const [coverageReport, setCoverageReport] = useState<DoctorCoverageResult | null>(null);
  const [coveragePreview, setCoveragePreview] = useState<SheetPreview | null>(null);
  const [coveragePreviewLoading, setCoveragePreviewLoading] = useState(false);
  const [monthlyPlannedCallLogFile, setMonthlyPlannedCallLogFile] = useState<File | null>(null);
  const [monthlyPlannedTemplateFile, setMonthlyPlannedTemplateFile] = useState<File | null>(null);
  const [monthlyPlannedProcessing, setMonthlyPlannedProcessing] = useState(false);
  const [monthlyPlannedReport, setMonthlyPlannedReport] = useState<MonthlyPlannedResult | null>(
    null,
  );
  const [monthlyPlannedPreview, setMonthlyPlannedPreview] = useState<SheetPreview | null>(null);
  const [monthlyPlannedPreviewLoading, setMonthlyPlannedPreviewLoading] = useState(false);

  const excelInputRef = useRef<HTMLInputElement>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);
  const previewRef = useRef<HTMLDivElement>(null);
  const callLogInputRef = useRef<HTMLInputElement>(null);
  const dailyTemplateInputRef = useRef<HTMLInputElement>(null);
  const dailyPreviewRef = useRef<HTMLDivElement>(null);
  const coverageSourceInputRef = useRef<HTMLInputElement>(null);
  const coverageTemplateInputRef = useRef<HTMLInputElement>(null);
  const coveragePreviewRef = useRef<HTMLDivElement>(null);
  const monthlyPlannedCallLogInputRef = useRef<HTMLInputElement>(null);
  const monthlyPlannedTemplateInputRef = useRef<HTMLInputElement>(null);
  const monthlyPlannedPreviewRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const key = "dashboard-line-index";
    const current = Number(window.localStorage.getItem(key) ?? "-1");
    const next = (Number.isFinite(current) ? current + 1 : 0) % DASHBOARD_LINES.length;
    window.localStorage.setItem(key, String(next));
    setDashboardLine(DASHBOARD_LINES[next]);
  }, []);

  useEffect(() => {
    const saved = window.localStorage.getItem("site-theme-color") || DEFAULT_THEME_COLOR;
    setThemeColor(saved);
    applyThemeColor(saved);
  }, []);

  useEffect(() => {
    const savedMode = window.localStorage.getItem("site-theme-mode") === "dark" ? "dark" : "light";
    setThemeMode(savedMode);
    applyThemeMode(savedMode);
  }, []);

  useEffect(() => {
    let mounted = true;

    void supabase.auth.getSession().then(async ({ data, error }) => {
      if (!mounted) return;
      if (error) {
        console.error("Supabase session load failed", error);
        toast.error("Unable to check your login session.");
      }

      if (data.session) {
        try {
          const deviceResult = await enforceDeviceLimit(data.session);
          if (!mounted) return;
          if (!deviceResult.allowed) {
            setAuthUser(null);
            toast.error(deviceResult.message ?? "This account is already active on 2 devices.");
          } else {
            setAuthUser(data.session.user);
          }
        } catch (deviceError) {
          console.error("Device login check failed", deviceError);
          await supabase.auth.signOut();
          if (!mounted) return;
          setAuthUser(null);
          toast.error("Unable to verify this device session.");
        }
      } else {
        setAuthUser(null);
      }
      setAuthLoading(false);
    });

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((event, session) => {
      if (event === "SIGNED_OUT" || !session) {
        setAuthUser(null);
        setAuthLoading(false);
        return;
      }

      void enforceDeviceLimit(session)
        .then((deviceResult) => {
          if (!mounted) return;
          if (!deviceResult.allowed) {
            setAuthUser(null);
            toast.error(deviceResult.message ?? "This account is already active on 2 devices.");
            return;
          }
          setAuthUser(session.user);
        })
        .catch(async (deviceError) => {
          console.error("Device login check failed", deviceError);
          await supabase.auth.signOut();
          if (!mounted) return;
          setAuthUser(null);
          toast.error("Unable to verify this device session.");
        })
        .finally(() => {
          if (mounted) setAuthLoading(false);
        });
    });

    return () => {
      mounted = false;
      subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    if (!authUser) return;
    void syncAuthUserToDatabase(authUser).catch((error) => {
      console.warn("Auth user database sync failed", error);
    });
  }, [authUser]);

  useEffect(() => {
    if (!authUser) {
      setProfilePhoto(null);
      return;
    }

    setProfilePhoto(window.localStorage.getItem(profilePhotoStorageKey(authUser)));
  }, [authUser]);

  const updateProfilePhoto = useCallback(
    (photo: string | null) => {
      if (!authUser) return;
      const storageKey = profilePhotoStorageKey(authUser);
      if (photo) {
        window.localStorage.setItem(storageKey, photo);
      } else {
        window.localStorage.removeItem(storageKey);
      }
      setProfilePhoto(photo);
    },
    [authUser],
  );

  const updateThemeColor = (color: string) => {
    setThemeColor(color);
    applyThemeColor(color);
    window.localStorage.setItem("site-theme-color", color);
  };

  const toggleThemeMode = () => {
    const next = themeMode === "dark" ? "light" : "dark";
    setThemeMode(next);
    applyThemeMode(next);
    window.localStorage.setItem("site-theme-mode", next);
  };

  const signOut = useCallback(async () => {
    setSigningOut(true);
    try {
      await revokeCurrentDeviceSession().catch((error) => {
        console.warn("Unable to revoke device session before sign out", error);
      });
      const { error } = await supabase.auth.signOut();
      if (error) throw error;
      setActiveModule(null);
      toast.success("Signed out.");
    } catch (error) {
      console.error("Sign out failed", error);
      toast.error(error instanceof Error ? error.message : "Unable to sign out.");
    } finally {
      setSigningOut(false);
    }
  }, []);

  const onExcelChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) {
      setExcelFile(f);
      setReport(null);
      setDownloadUrl(null);
      setSheetPreview(null);
      setPreviewDirty(false);
    }
    e.target.value = "";
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
    setSheetPreview(null);
    setPreviewDirty(false);
    e.target.value = "";
  };

  const removePdf = (name: string) => {
    setPdfFiles((prev) => prev.filter((p) => p.name !== name));
    setReport(null);
    setDownloadUrl(null);
    setSheetPreview(null);
    setPreviewDirty(false);
  };

  const clearAllPdfs = () => {
    setPdfFiles([]);
    if (pdfInputRef.current) pdfInputRef.current.value = "";
    setReport(null);
    setDownloadUrl(null);
    setSheetPreview(null);
    setPreviewDirty(false);
  };

  const sortedDates = useMemo(() => {
    const ds = pdfFiles.map((f) => extractDateFromPdfName(f.name)).filter((x): x is string => !!x);
    return [...new Set(ds)].sort();
  }, [pdfFiles]);
  const canProcess = Boolean(excelFile) && pdfFiles.length > 0 && !processing;
  const matchRate = report
    ? Math.round(
        (report.debug.totalMatched / Math.max(report.debug.totalEmployeesDetected, 1)) * 100,
      )
    : null;

  const refreshHistory = useCallback(async () => {
    if (!authUser) {
      setHistoryItems([]);
      return;
    }

    setHistoryLoading(true);
    try {
      setHistoryItems(await listReportHistory(authUser.id));
    } catch (error) {
      console.error("History load failed", error);
      toast.error("Unable to load report history.");
    } finally {
      setHistoryLoading(false);
    }
  }, [authUser]);

  const openHistory = useCallback(() => {
    setActiveModule("history");
    setHistoryPreview(null);
    void refreshHistory();
  }, [refreshHistory]);

  useEffect(() => {
    if (!authUser) {
      setHistoryItems([]);
      setHistoryPreview(null);
      return;
    }

    void (async () => {
      await syncStoredReportHistoryToDatabase(authUser.id);
      await refreshHistory();
    })();
  }, [authUser, refreshHistory]);

  const canProcessDaily = callLogFiles.length > 0 && Boolean(dailyTemplateFile) && !dailyProcessing;
  const canProcessCoverage =
    Boolean(coverageSourceFile) && Boolean(coverageTemplateFile) && !coverageProcessing;
  const canProcessMonthlyPlanned =
    Boolean(monthlyPlannedCallLogFile) &&
    Boolean(monthlyPlannedTemplateFile) &&
    !monthlyPlannedProcessing;

  const saveReportToHistory = useCallback(
    async (nextReport: ProcessReport) => {
      if (!authUser) throw new Error("You must be signed in to save report history.");
      try {
        await saveReportHistoryAndDatabase({
          id: crypto.randomUUID(),
          userId: authUser.id,
          fileName: nextReport.fileName.replace(/\.xlsx$/i, "") + " - Updated.xlsx",
          createdAt: new Date().toISOString(),
          reportType: "Monthly Report",
          dates: nextReport.dates,
          pdfCount: pdfFiles.length,
          totalEmployees: nextReport.totalEmployees,
          matchedEmployees: nextReport.matchedEmployees,
          size: nextReport.blob.size,
          blob: nextReport.blob,
        });
        if (activeModule === "history") await refreshHistory();
      } catch (error) {
        console.error("History save failed", error);
        toast.warning("Report ready, but history save failed.");
      }
    },
    [activeModule, authUser, pdfFiles.length, refreshHistory],
  );

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
    setSheetPreview(null);
    setPreviewDirty(false);

    try {
      // 1. Parse PDFs (with file-hash dedup against DB)
      const results: PdfParseResult[] = [];
      const skipped: string[] = [];
      const failed: { fileName: string; error: string }[] = [];

      // Pre-hash all files to check duplicates in one query
      setProgressLabel("Checking for duplicate uploads...");
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
        toast.info(`${skipped.length} file(s) were already processed before - re-running.`);
      }
      if (failed.length) {
        toast.warning(`Skipped ${failed.length} unreadable PDF(s)`);
      }
      if (results.length === 0) throw new Error("No PDFs could be read.");

      // 2. Update Excel
      setProgressLabel("Updating your Excel file...");
      setProgress(80);
      const rep = await processExcel(excelFile, { pdfResults: results });
      if (rep.warnings.length) {
        toast.warning(rep.warnings[0]);
      }

      // 3. Sync to database (best-effort)
      setProgressLabel("Saving history to database...");
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
      await saveReportToHistory(rep);

      toast.success(
        `Report ready! ${rep.matchedEmployees}/${rep.totalEmployees} employees with selfies across ${rep.dates.length} day(s).`,
      );
    } catch (err) {
      console.error(err);
      toast.error(err instanceof Error ? err.message : "Something went wrong while processing.");
    } finally {
      setProcessing(false);
    }
  }, [excelFile, pdfFiles, saveReportToHistory]);

  const savePreviewEdits = async (): Promise<ProcessReport | null> => {
    if (!report) return null;
    if (!sheetPreview || !previewDirty) return report;

    setPreviewSaving(true);
    try {
      const blob = await applySheetPreviewEdits(report.blob, sheetPreview);
      const updatedReport = { ...report, blob };
      const url = URL.createObjectURL(blob);
      if (downloadUrl) URL.revokeObjectURL(downloadUrl);
      setReport(updatedReport);
      setDownloadUrl(url);
      setPreviewDirty(false);
      await saveReportToHistory(updatedReport);
      toast.success("Manual edits applied to the Excel file.");
      return updatedReport;
    } catch (error) {
      console.error("Excel edit save failed", error);
      toast.error("Unable to apply manual edits.");
      return null;
    } finally {
      setPreviewSaving(false);
    }
  };

  const triggerDownload = async () => {
    if (!report) return;
    const latestReport = await savePreviewEdits();
    if (!latestReport) return;
    const url = URL.createObjectURL(latestReport.blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = latestReport.fileName.replace(/\.xlsx$/i, "") + " - Updated.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const onDailyProcess = async () => {
    if (!callLogFiles.length) {
      toast.error("Please upload at least one call log Excel file.");
      return;
    }
    if (!dailyTemplateFile) {
      toast.error("Please upload the daily report sample file.");
      return;
    }

    setDailyProcessing(true);
    setDailyReport(null);
    setBulkDailyReport(null);
    setDailyPreview(null);
    setBulkDailyProgress(0);
    setBulkDailyProgressLabel("Detecting teams...");
    try {
      const bulkResult = await processBulkDailyReports(
        callLogFiles,
        dailyTemplateFile,
        (status) => {
          setBulkDailyProgress(Math.round((status.current / Math.max(status.total, 1)) * 100));
          setBulkDailyProgressLabel(
            `Generating ${status.teamName} (${status.current}/${status.total})`,
          );
        },
      );

      if (bulkResult.totalTeams > 1 || bulkResult.failedReports > 0) {
        setBulkDailyProgress(100);
        setBulkDailyProgressLabel("Reports ready.");
        setBulkDailyReport(bulkResult);
        await saveReportHistoryAndDatabase({
          id: crypto.randomUUID(),
          userId: authUser?.id,
          fileName: bulkResult.fileName,
          createdAt: new Date().toISOString(),
          reportType: "Daily Reports",
          dates: [],
          pdfCount: 0,
          totalEmployees: bulkResult.summary.reduce((sum, item) => sum + item.totalEmployees, 0),
          matchedEmployees: bulkResult.summary.reduce(
            (sum, item) => sum + item.matchedEmployees,
            0,
          ),
          size: bulkResult.blob.size,
          blob: bulkResult.blob,
        });
        toast.success(
          `Daily report ready. ${bulkResult.reportsGenerated}/${bulkResult.totalTeams} team sheet(s) updated.`,
        );
        return;
      }

      setBulkDailyProgressLabel("Preparing report preview...");
      const result = await processDailyReport(callLogFiles, dailyTemplateFile);
      setDailyReport(result);
      await saveReportHistoryAndDatabase({
        id: crypto.randomUUID(),
        userId: authUser?.id,
        fileName: result.fileName,
        createdAt: new Date().toISOString(),
        reportType: "Daily Report",
        dates: [],
        pdfCount: 0,
        totalEmployees: result.totalEmployees,
        matchedEmployees: result.matchedEmployees,
        size: result.blob.size,
        blob: result.blob,
      });
      if (result.warnings.length) toast.warning(result.warnings[0]);
      toast.success(`Daily report ready. ${result.matchedEmployees} employee(s) matched.`);
    } catch (error) {
      console.error("Daily report failed", error);
      toast.error(error instanceof Error ? error.message : "Unable to generate daily report.");
    } finally {
      setDailyProcessing(false);
    }
  };

  const openDailyPreview = async () => {
    if (!dailyReport && !bulkDailyReport) return;
    setDailyPreviewLoading(true);
    try {
      const preview = dailyReport
        ? await buildSheetPreview(dailyReport.blob, dailyReport.sheetName)
        : await buildSheetPreview(bulkDailyReport!.blob);
      setDailyPreview(preview);
      requestAnimationFrame(() => {
        dailyPreviewRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    } catch (error) {
      console.error("Daily preview failed", error);
      toast.error("Unable to open the daily report preview.");
    } finally {
      setDailyPreviewLoading(false);
    }
  };

  const downloadDailyReport = () => {
    const output = dailyReport ?? bulkDailyReport;
    if (!output) return;
    const url = URL.createObjectURL(output.blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = output.fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const onCoverageProcess = async () => {
    if (!coverageSourceFile) {
      toast.error("Please upload the Doctor Coverage Excel file.");
      return;
    }
    if (!coverageTemplateFile) {
      toast.error("Please upload the sample employee Excel file.");
      return;
    }

    setCoverageProcessing(true);
    setCoverageReport(null);
    setCoveragePreview(null);
    try {
      const result = await processDoctorCoverageReport(coverageSourceFile, coverageTemplateFile);
      setCoverageReport(result);
      await saveReportHistoryAndDatabase({
        id: crypto.randomUUID(),
        fileName: result.fileName,
        createdAt: new Date().toISOString(),
        reportType: "Doctor Coverage Report",
        dates: [],
        pdfCount: 0,
        totalEmployees: result.totalEmployees,
        matchedEmployees: result.matchedEmployees,
        size: result.blob.size,
        blob: result.blob,
      });
      if (result.warnings.length) toast.warning(result.warnings[0]);
      toast.success(
        `Doctor coverage report ready. ${result.matchedEmployees} employee(s) matched.`,
      );
    } catch (error) {
      console.error("Doctor coverage report failed", error);
      toast.error(
        error instanceof Error ? error.message : "Unable to generate doctor coverage report.",
      );
    } finally {
      setCoverageProcessing(false);
    }
  };

  const openCoveragePreview = async () => {
    if (!coverageReport) return;
    setCoveragePreviewLoading(true);
    try {
      const preview = await buildSheetPreview(coverageReport.blob, coverageReport.sheetName);
      setCoveragePreview(preview);
      requestAnimationFrame(() => {
        coveragePreviewRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    } catch (error) {
      console.error("Doctor coverage preview failed", error);
      toast.error("Unable to open the doctor coverage preview.");
    } finally {
      setCoveragePreviewLoading(false);
    }
  };

  const downloadCoverageReport = () => {
    if (!coverageReport) return;
    const url = URL.createObjectURL(coverageReport.blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = coverageReport.fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const onMonthlyPlannedProcess = async () => {
    if (!monthlyPlannedCallLogFile) {
      toast.error("Please upload the monthly call log Excel file.");
      return;
    }
    if (!monthlyPlannedTemplateFile) {
      toast.error("Please upload the sample employee Excel file.");
      return;
    }

    setMonthlyPlannedProcessing(true);
    setMonthlyPlannedReport(null);
    setMonthlyPlannedPreview(null);
    try {
      const result = await processMonthlyPlannedReport(
        monthlyPlannedCallLogFile,
        monthlyPlannedTemplateFile,
      );
      setMonthlyPlannedReport(result);
      await saveReportHistoryAndDatabase({
        id: crypto.randomUUID(),
        fileName: result.fileName,
        createdAt: new Date().toISOString(),
        reportType: "Monthly Planned Unplanned",
        dates: [],
        pdfCount: 0,
        totalEmployees: result.totalEmployees,
        matchedEmployees: result.matchedEmployees,
        size: result.blob.size,
        blob: result.blob,
      });
      if (result.warnings.length) toast.warning(result.warnings[0]);
      toast.success(
        `Monthly planned/unplanned report ready. ${result.matchedEmployees} employee(s) matched.`,
      );
    } catch (error) {
      console.error("Monthly planned/unplanned report failed", error);
      toast.error(
        error instanceof Error
          ? error.message
          : "Unable to generate monthly planned/unplanned report.",
      );
    } finally {
      setMonthlyPlannedProcessing(false);
    }
  };

  const openMonthlyPlannedPreview = async () => {
    if (!monthlyPlannedReport) return;
    setMonthlyPlannedPreviewLoading(true);
    try {
      const preview = await buildSheetPreview(
        monthlyPlannedReport.blob,
        monthlyPlannedReport.sheetName,
      );
      setMonthlyPlannedPreview(preview);
      requestAnimationFrame(() => {
        monthlyPlannedPreviewRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    } catch (error) {
      console.error("Monthly planned/unplanned preview failed", error);
      toast.error("Unable to open the monthly planned/unplanned preview.");
    } finally {
      setMonthlyPlannedPreviewLoading(false);
    }
  };

  const downloadMonthlyPlannedReport = () => {
    if (!monthlyPlannedReport) return;
    const url = URL.createObjectURL(monthlyPlannedReport.blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = monthlyPlannedReport.fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const downloadHistoryItem = async (id: string) => {
    if (!authUser) return;
    try {
      const item = await getReportHistory(id, authUser.id);
      if (!item) {
        toast.error("History file not found.");
        await refreshHistory();
        return;
      }
      const url = URL.createObjectURL(item.blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = item.fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error("History download failed", error);
      toast.error("Unable to download history file.");
    }
  };

  const previewHistoryItem = async (id: string) => {
    if (!authUser) return;
    setHistoryPreviewLoading(true);
    setHistoryPreview(null);
    try {
      const item = await getReportHistory(id, authUser.id);
      if (!item) {
        toast.error("History file not found.");
        await refreshHistory();
        return;
      }
      const preview = await buildSheetPreview(item.blob);
      setHistoryPreview({ ...preview, name: item.fileName });
    } catch (error) {
      console.error("History preview failed", error);
      toast.error("Unable to preview history file.");
    } finally {
      setHistoryPreviewLoading(false);
    }
  };

  const removeHistoryItem = async (id: string) => {
    if (!authUser) return;
    try {
      await deleteReportHistory(id, authUser.id);
      await refreshHistory();
      setHistoryPreview(null);
      toast.success("History file removed.");
    } catch (error) {
      console.error("History delete failed", error);
      toast.error("Unable to remove history file.");
    }
  };

  const openFilePreview = async () => {
    if (!report) return;
    setPreviewLoading(true);
    try {
      const preview = await buildSheetPreview(report.blob, report.sheetName);
      setSheetPreview(preview);
      setPreviewDirty(false);
      requestAnimationFrame(() => {
        previewRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
      });
    } catch (error) {
      console.error("Excel preview failed", error);
      toast.error("Unable to open the Excel preview.");
    } finally {
      setPreviewLoading(false);
    }
  };

  const updatePreviewCell = (rowNumber: number, colNumber: number, value: string) => {
    setSheetPreview((current) => {
      if (!current) return current;
      return {
        ...current,
        rows: current.rows.map((row) =>
          row.map((cell) =>
            cell.rowNumber === rowNumber && cell.colNumber === colNumber
              ? { ...cell, value }
              : cell,
          ),
        ),
      };
    });
    setPreviewDirty(true);
  };

  if (authLoading) {
    return (
      <AuthLoadingScreen
        color={themeColor}
        mode={themeMode}
        onChange={updateThemeColor}
        onToggleMode={toggleThemeMode}
      />
    );
  }

  if (!authUser) {
    return (
      <AuthScreen
        color={themeColor}
        mode={themeMode}
        onChange={updateThemeColor}
        onToggleMode={toggleThemeMode}
      />
    );
  }

  if (!activeModule) {
    return (
      <AppShell
        activeModule={activeModule}
        dashboardLine={dashboardLine}
        user={authUser}
        profilePhoto={profilePhoto}
        onNavigate={setActiveModule}
        onOpenHistory={openHistory}
      >
        <DashboardHome
          historyItems={historyItems}
          historyLoading={historyLoading}
          onNavigate={setActiveModule}
          onOpenHistory={openHistory}
        />
      </AppShell>
    );
  }

  if (activeModule === "profile") {
    return (
      <AppShell
        activeModule={activeModule}
        dashboardLine={dashboardLine}
        user={authUser}
        profilePhoto={profilePhoto}
        onNavigate={setActiveModule}
        onOpenHistory={openHistory}
      >
        <ProfilePage
          user={authUser}
          profilePhoto={profilePhoto}
          color={themeColor}
          mode={themeMode}
          signingOut={signingOut}
          onProfilePhotoChange={updateProfilePhoto}
          onChange={updateThemeColor}
          onToggleMode={toggleThemeMode}
          onSignOut={signOut}
        />
      </AppShell>
    );
  }

  if (activeModule === "daily-report") {
    return (
      <AppShell
        activeModule={activeModule}
        dashboardLine={dashboardLine}
        user={authUser}
        profilePhoto={profilePhoto}
        onNavigate={setActiveModule}
        onOpenHistory={openHistory}
      >
        <ModulePageHeader
          icon={<TableProperties className="h-5 w-5" />}
          label="Report tool"
          title="Daily Report"
          description="Upload one sample workbook and one or more call log Excel files. No PDF upload is required."
        />

        <main className="grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_340px]">
          <div className="space-y-6">
            <div className="grid gap-4 sm:grid-cols-2">
              <UploadCard
                step={1}
                title="Daily Report Sample"
                subtitle={
                  dailyTemplateFile ? "Sample file selected" : "Upload the report template workbook"
                }
                icon={<FileText className="h-5 w-5" />}
                accept=".xlsx,.xls,.xlsm"
                ready={Boolean(dailyTemplateFile)}
                onClick={() => dailyTemplateInputRef.current?.click()}
              >
                <input
                  ref={dailyTemplateInputRef}
                  type="file"
                  accept=".xlsx,.xls,.xlsm"
                  className="hidden"
                  onChange={(event) => {
                    const file = event.target.files?.[0];
                    if (!file) return;
                    setDailyTemplateFile(file);
                    setDailyReport(null);
                    setBulkDailyReport(null);
                    setDailyPreview(null);
                  }}
                />
                {dailyTemplateFile ? (
                  <FilePill
                    name={dailyTemplateFile.name}
                    onRemove={() => {
                      setDailyTemplateFile(null);
                      setDailyReport(null);
                      setBulkDailyReport(null);
                      setDailyPreview(null);
                    }}
                    color="primary"
                  />
                ) : (
                  <EmptyHint icon={<FileText className="h-5 w-5" />} label="Choose sample file" />
                )}
              </UploadCard>

              <UploadCard
                step={2}
                title="Call Log Excel"
                subtitle={
                  callLogFiles.length
                    ? `${callLogFiles.length} call log file(s) selected`
                    : "Upload one or more team call log workbooks"
                }
                icon={<FileSpreadsheet className="h-5 w-5" />}
                accept=".xlsx,.xls,.xlsm"
                ready={callLogFiles.length > 0}
                onClick={() => callLogInputRef.current?.click()}
              >
                <input
                  ref={callLogInputRef}
                  type="file"
                  accept=".xlsx,.xls,.xlsm"
                  multiple
                  className="hidden"
                  onChange={(event) => {
                    const files = Array.from(event.target.files ?? []);
                    if (!files.length) return;
                    setCallLogFiles(files);
                    setDailyReport(null);
                    setBulkDailyReport(null);
                    setDailyPreview(null);
                  }}
                />
                {callLogFiles.length ? (
                  <div className="space-y-2">
                    {callLogFiles.map((file) => (
                      <FilePill
                        key={`${file.name}-${file.lastModified}-${file.size}`}
                        name={file.name}
                        onRemove={() => {
                          setCallLogFiles((current) =>
                            current.filter(
                              (item) =>
                                !(
                                  item.name === file.name &&
                                  item.lastModified === file.lastModified &&
                                  item.size === file.size
                                ),
                            ),
                          );
                          if (callLogInputRef.current) callLogInputRef.current.value = "";
                          setDailyReport(null);
                          setBulkDailyReport(null);
                          setDailyPreview(null);
                        }}
                        color="accent"
                      />
                    ))}
                  </div>
                ) : (
                  <EmptyHint
                    icon={<FileSpreadsheet className="h-5 w-5" />}
                    label="Choose call log file(s)"
                  />
                )}
              </UploadCard>
            </div>

            <Card className="border-border bg-card p-6 shadow-[var(--shadow-soft)]">
              <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-center">
                <div>
                  <div className="text-xs font-bold uppercase text-primary">Step 3</div>
                  <h2 className="mt-1 text-lg font-semibold text-foreground">
                    Generate daily report
                  </h2>
                  <p className="text-sm text-muted-foreground">
                    Data is matched by Employee Code only, then Planned, Unplanned, Mor, Eve, Total,
                    and Cp are filled in the sample file. Select all team call-log files together to
                    generate all matching team reports from a multi-team sample workbook.
                  </p>
                </div>
                <Button
                  size="lg"
                  onClick={onDailyProcess}
                  disabled={!canProcessDaily}
                  className="w-full min-w-48 shadow-[var(--shadow-elegant)] sm:w-auto"
                >
                  {dailyProcessing ? (
                    <>
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      Generating...
                    </>
                  ) : (
                    <>
                      <Sparkles className="mr-2 h-4 w-4" />
                      Generate Report
                    </>
                  )}
                </Button>
              </div>

              {dailyProcessing && bulkDailyProgressLabel && (
                <div className="mt-5 space-y-2">
                  <Progress value={bulkDailyProgress} />
                  <div className="text-xs text-muted-foreground">{bulkDailyProgressLabel}</div>
                </div>
              )}
            </Card>

            {(dailyReport || bulkDailyReport) && (
              <Card className="border-border bg-card p-6 shadow-[var(--shadow-soft)]">
                <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-center">
                  <div className="flex items-start gap-3">
                    <div className="flex h-10 w-10 items-center justify-center rounded-full bg-success/15 text-success">
                      <CheckCircle2 className="h-5 w-5" />
                    </div>
                    <div>
                      <h3 className="text-lg font-semibold text-foreground">Daily report ready</h3>
                      <p className="text-sm text-muted-foreground">
                        {dailyReport
                          ? `${dailyReport.matchedEmployees}/${dailyReport.totalEmployees} employee(s) matched from the sample file.`
                          : `${bulkDailyReport!.reportsGenerated}/${bulkDailyReport!.totalTeams} team sheet(s) updated in one Excel workbook.`}
                      </p>
                    </div>
                  </div>
                  <div className="grid w-full gap-2 sm:w-auto sm:grid-cols-2">
                    <Button
                      onClick={openDailyPreview}
                      variant="outline"
                      disabled={dailyPreviewLoading}
                    >
                      {dailyPreviewLoading ? (
                        <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      ) : (
                        <Eye className="mr-2 h-4 w-4" />
                      )}
                      Preview
                    </Button>
                    <Button onClick={downloadDailyReport}>
                      <Download className="mr-2 h-4 w-4" />
                      Download
                    </Button>
                  </div>
                </div>

                {dailyReport ? (
                  <>
                    <div className="mt-6 grid gap-4 md:grid-cols-3">
                      <Stat label="Call log rows" value={dailyReport.debug.callRows} />
                      <Stat label="Face-to-face calls" value={dailyReport.debug.faceToFaceRows} />
                      <Stat label="Contact points" value={dailyReport.debug.contactPointRows} />
                    </div>

                    {dailyReport.preview.length > 0 && (
                      <div className="mt-6">
                        <div className="mb-2 text-sm font-semibold text-foreground">
                          Top callers
                        </div>
                        <div className="space-y-1.5">
                          {dailyReport.preview.map((item, index) => (
                            <div
                              key={`${item.name}-${index}`}
                              className="flex items-center justify-between rounded-md border border-border bg-muted/40 px-3 py-2 text-sm"
                            >
                              <span>{item.name}</span>
                              <span className="font-semibold text-foreground">{item.total}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </>
                ) : (
                  <>
                    <div className="mt-6 grid gap-4 md:grid-cols-3">
                      <Stat label="Teams processed" value={bulkDailyReport!.totalTeams} />
                      <Stat label="Sheets updated" value={bulkDailyReport!.reportsGenerated} />
                      <Stat
                        label="Failed reports"
                        value={bulkDailyReport!.failedReports}
                        tone={bulkDailyReport!.failedReports ? "warning" : "default"}
                      />
                    </div>

                    <div className="mt-6 max-h-72 overflow-y-auto rounded-md border border-border">
                      {bulkDailyReport!.summary.map((item) => (
                        <div
                          key={item.teamName}
                          className="grid gap-2 border-b border-border p-3 text-sm last:border-b-0 md:grid-cols-[minmax(0,1fr)_110px_120px_minmax(0,1.3fr)]"
                        >
                          <div className="min-w-0 font-semibold text-foreground">
                            {item.teamName}
                          </div>
                          <div
                            className={
                              item.status === "success" ? "text-success" : "text-warning-foreground"
                            }
                          >
                            {item.status}
                          </div>
                          <div className="text-muted-foreground">
                            {item.matchedEmployees}/{item.totalEmployees}
                          </div>
                          <div className="min-w-0 truncate text-muted-foreground">
                            {item.error ?? "Updated in workbook"}
                          </div>
                        </div>
                      ))}
                    </div>
                  </>
                )}
              </Card>
            )}

            {(dailyPreviewLoading || dailyPreview) && (
              <Card
                ref={dailyPreviewRef}
                className="border-border bg-card p-5 shadow-[var(--shadow-soft)]"
              >
                <div className="mb-3 text-sm font-semibold text-foreground">
                  {dailyPreview?.name ?? "Opening daily report preview"}
                </div>
                <div className="max-h-[70vh] overflow-auto rounded-md border border-border bg-background">
                  {dailyPreviewLoading && !dailyPreview ? (
                    <div className="flex items-center gap-2 p-4 text-sm text-muted-foreground">
                      <Loader2 className="h-4 w-4 animate-spin" />
                      Opening daily report preview...
                    </div>
                  ) : (
                    <PreviewTable preview={dailyPreview} readOnly />
                  )}
                </div>
              </Card>
            )}
          </div>

          <aside className="space-y-4">
            <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
              <div className="mb-4 flex items-center justify-between">
                <div>
                  <div className="text-xs font-bold uppercase text-primary">Run status</div>
                  <h2 className="mt-1 text-base font-semibold text-foreground">Daily workspace</h2>
                </div>
                <ShieldCheck className="h-5 w-5 text-primary" />
              </div>
              <div className="space-y-3">
                <ChecklistItem done={callLogFiles.length > 0} label="Call log attached" />
                <ChecklistItem done={Boolean(dailyTemplateFile)} label="Sample file attached" />
                <ChecklistItem
                  done={Boolean(dailyReport || bulkDailyReport)}
                  label="Report generated"
                />
                <ChecklistItem done={Boolean(dailyPreview)} label="Preview opened" />
              </div>
            </Card>
            <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
              <div className="text-xs font-bold uppercase text-primary">Match rule</div>
              <p className="mt-2 text-sm text-muted-foreground">
                The sample file must include: Employee Code, Name, Planned, Unplanned, Mor, Eve,
                Total, and Cp. For all teams, select every team's call-log file in the Call Log
                Excel upload.
              </p>
            </Card>
          </aside>
        </main>
      </AppShell>
    );
  }

  if (activeModule === "doctor-coverage") {
    return (
      <AppShell
        activeModule={activeModule}
        dashboardLine={dashboardLine}
        user={authUser}
        profilePhoto={profilePhoto}
        onNavigate={setActiveModule}
        onOpenHistory={openHistory}
      >
        <ModulePageHeader
          icon={<BarChart3 className="h-5 w-5" />}
          label="Report tool"
          title="Doctor Coverage"
          description="Add coverage columns to a sample employee file."
        />

        <main className="grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_340px]">
          <div className="space-y-6">
            <div className="grid gap-4 sm:grid-cols-2">
              <UploadCard
                step={1}
                title="Sample Employee Excel"
                subtitle={
                  coverageTemplateFile ? "Sample file selected" : "Upload the sample employee file"
                }
                icon={<FileText className="h-5 w-5" />}
                accept=".xlsx,.xls"
                ready={Boolean(coverageTemplateFile)}
                onClick={() => coverageTemplateInputRef.current?.click()}
              >
                <input
                  ref={coverageTemplateInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  className="hidden"
                  onChange={(event) => {
                    const file = event.target.files?.[0];
                    if (!file) return;
                    setCoverageTemplateFile(file);
                    setCoverageReport(null);
                    setCoveragePreview(null);
                  }}
                />
                {coverageTemplateFile ? (
                  <FilePill
                    name={coverageTemplateFile.name}
                    onRemove={() => {
                      setCoverageTemplateFile(null);
                      setCoverageReport(null);
                      setCoveragePreview(null);
                    }}
                    color="primary"
                  />
                ) : (
                  <EmptyHint icon={<FileText className="h-5 w-5" />} label="Choose sample file" />
                )}
              </UploadCard>

              <UploadCard
                step={2}
                title="Doctor Coverage Excel"
                subtitle={
                  coverageSourceFile ? "Coverage source selected" : "Upload the coverage workbook"
                }
                icon={<FileSpreadsheet className="h-5 w-5" />}
                accept=".xlsx,.xls"
                ready={Boolean(coverageSourceFile)}
                onClick={() => coverageSourceInputRef.current?.click()}
              >
                <input
                  ref={coverageSourceInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  className="hidden"
                  onChange={(event) => {
                    const file = event.target.files?.[0];
                    if (!file) return;
                    setCoverageSourceFile(file);
                    setCoverageReport(null);
                    setCoveragePreview(null);
                  }}
                />
                {coverageSourceFile ? (
                  <FilePill
                    name={coverageSourceFile.name}
                    onRemove={() => {
                      setCoverageSourceFile(null);
                      setCoverageReport(null);
                      setCoveragePreview(null);
                    }}
                    color="accent"
                  />
                ) : (
                  <EmptyHint
                    icon={<FileSpreadsheet className="h-5 w-5" />}
                    label="Choose coverage file"
                  />
                )}
              </UploadCard>
            </div>

            <Card className="border-border bg-card p-6 shadow-[var(--shadow-soft)]">
              <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-center">
                <div>
                  <div className="text-xs font-bold uppercase text-primary">Step 3</div>
                  <h2 className="mt-1 text-lg font-semibold text-foreground">
                    Generate doctor coverage report
                  </h2>
                  <p className="text-sm text-muted-foreground">
                    Data is matched by Employee Code first, then Name, and coverage columns are
                    filled in the sample file.
                  </p>
                </div>
                <Button
                  size="lg"
                  onClick={onCoverageProcess}
                  disabled={!canProcessCoverage}
                  className="w-full min-w-44 shadow-[var(--shadow-elegant)] sm:w-auto"
                >
                  {coverageProcessing ? (
                    <>
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      Processing...
                    </>
                  ) : (
                    <>
                      <Sparkles className="mr-2 h-4 w-4" />
                      Generate Report
                    </>
                  )}
                </Button>
              </div>
            </Card>

            {coverageReport && (
              <Card className="border-border bg-card p-6 shadow-[var(--shadow-soft)]">
                <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-center">
                  <div className="flex items-start gap-3">
                    <div className="flex h-10 w-10 items-center justify-center rounded-full bg-success/15 text-success">
                      <CheckCircle2 className="h-5 w-5" />
                    </div>
                    <div>
                      <h3 className="text-lg font-semibold text-foreground">
                        Doctor coverage report ready
                      </h3>
                      <p className="text-sm text-muted-foreground">
                        {coverageReport.matchedEmployees}/{coverageReport.totalEmployees}{" "}
                        employee(s) matched from the sample file.
                      </p>
                    </div>
                  </div>
                  <div className="grid w-full gap-2 sm:w-auto sm:grid-cols-2">
                    <Button
                      onClick={openCoveragePreview}
                      variant="outline"
                      disabled={coveragePreviewLoading}
                    >
                      {coveragePreviewLoading ? (
                        <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      ) : (
                        <Eye className="mr-2 h-4 w-4" />
                      )}
                      Preview
                    </Button>
                    <Button onClick={downloadCoverageReport}>
                      <Download className="mr-2 h-4 w-4" />
                      Download
                    </Button>
                  </div>
                </div>

                <div className="mt-6 grid gap-4 md:grid-cols-3">
                  <Stat label="Coverage rows" value={coverageReport.debug.sourceRows} />
                  <Stat label="Sample rows" value={coverageReport.debug.templateRows} />
                  <Stat label="Matched" value={coverageReport.matchedEmployees} />
                </div>

                {coverageReport.preview.length > 0 && (
                  <div className="mt-6">
                    <div className="mb-2 text-sm font-semibold text-foreground">
                      Highest covered doctors
                    </div>
                    <div className="space-y-1.5">
                      {coverageReport.preview.map((item, index) => (
                        <div
                          key={`${item.name}-${index}`}
                          className="flex items-center justify-between rounded-md border border-border bg-muted/40 px-3 py-2 text-sm"
                        >
                          <span>{item.name}</span>
                          <span className="font-semibold text-foreground">{item.total}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </Card>
            )}

            {(coveragePreviewLoading || coveragePreview) && (
              <Card
                ref={coveragePreviewRef}
                className="border-border bg-card p-5 shadow-[var(--shadow-soft)]"
              >
                <div className="mb-3 text-sm font-semibold text-foreground">
                  {coveragePreview?.name ?? "Opening doctor coverage preview"}
                </div>
                <div className="max-h-[70vh] overflow-auto rounded-md border border-border bg-background">
                  {coveragePreviewLoading && !coveragePreview ? (
                    <div className="flex items-center gap-2 p-4 text-sm text-muted-foreground">
                      <Loader2 className="h-4 w-4 animate-spin" />
                      Opening doctor coverage preview...
                    </div>
                  ) : (
                    <PreviewTable preview={coveragePreview} readOnly />
                  )}
                </div>
              </Card>
            )}
          </div>

          <aside className="space-y-4">
            <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
              <div className="mb-4 flex items-center justify-between">
                <div>
                  <div className="text-xs font-bold uppercase text-primary">Run status</div>
                  <h2 className="mt-1 text-base font-semibold text-foreground">
                    Coverage workspace
                  </h2>
                </div>
                <ShieldCheck className="h-5 w-5 text-primary" />
              </div>
              <div className="space-y-3">
                <ChecklistItem done={Boolean(coverageSourceFile)} label="Coverage file attached" />
                <ChecklistItem done={Boolean(coverageTemplateFile)} label="Sample file attached" />
                <ChecklistItem done={Boolean(coverageReport)} label="Report generated" />
                <ChecklistItem done={Boolean(coveragePreview)} label="Preview opened" />
              </div>
            </Card>
            <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
              <div className="text-xs font-bold uppercase text-primary">Columns added</div>
              <p className="mt-2 text-sm text-muted-foreground">
                Target Doctors, Covered Doctors, and Coverage % are added to the sample file without
                changing existing employee rows.
              </p>
            </Card>
          </aside>
        </main>
      </AppShell>
    );
  }

  if (activeModule === "monthly-planned") {
    return (
      <AppShell
        activeModule={activeModule}
        dashboardLine={dashboardLine}
        user={authUser}
        profilePhoto={profilePhoto}
        onNavigate={setActiveModule}
        onOpenHistory={openHistory}
      >
        <ModulePageHeader
          icon={<CalendarDays className="h-5 w-5" />}
          label="Report tool"
          title="Monthly Planned Unplanned"
          description="Monthly call log summary by employee."
        />

        <main className="grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_340px]">
          <div className="space-y-6">
            <div className="grid gap-4 sm:grid-cols-2">
              <UploadCard
                step={1}
                title="Sample Employee Excel"
                subtitle={
                  monthlyPlannedTemplateFile ? "Sample file selected" : "Upload the sample file"
                }
                icon={<FileText className="h-5 w-5" />}
                accept=".xlsx,.xls,.xlsm"
                ready={Boolean(monthlyPlannedTemplateFile)}
                onClick={() => monthlyPlannedTemplateInputRef.current?.click()}
              >
                <input
                  ref={monthlyPlannedTemplateInputRef}
                  type="file"
                  accept=".xlsx,.xls,.xlsm"
                  className="hidden"
                  onChange={(event) => {
                    const file = event.target.files?.[0];
                    if (!file) return;
                    setMonthlyPlannedTemplateFile(file);
                    setMonthlyPlannedReport(null);
                    setMonthlyPlannedPreview(null);
                    event.target.value = "";
                  }}
                />
                {monthlyPlannedTemplateFile ? (
                  <FilePill
                    name={monthlyPlannedTemplateFile.name}
                    onRemove={() => {
                      setMonthlyPlannedTemplateFile(null);
                      setMonthlyPlannedReport(null);
                      setMonthlyPlannedPreview(null);
                    }}
                    color="primary"
                  />
                ) : (
                  <EmptyHint icon={<FileText className="h-5 w-5" />} label="Choose sample file" />
                )}
              </UploadCard>

              <UploadCard
                step={2}
                title="Monthly Call Log Excel"
                subtitle={
                  monthlyPlannedCallLogFile ? "Call log selected" : "Upload the monthly call log"
                }
                icon={<FileSpreadsheet className="h-5 w-5" />}
                accept=".xlsx,.xls,.xlsm"
                ready={Boolean(monthlyPlannedCallLogFile)}
                onClick={() => monthlyPlannedCallLogInputRef.current?.click()}
              >
                <input
                  ref={monthlyPlannedCallLogInputRef}
                  type="file"
                  accept=".xlsx,.xls,.xlsm"
                  className="hidden"
                  onChange={(event) => {
                    const file = event.target.files?.[0];
                    if (!file) return;
                    setMonthlyPlannedCallLogFile(file);
                    setMonthlyPlannedReport(null);
                    setMonthlyPlannedPreview(null);
                    event.target.value = "";
                  }}
                />
                {monthlyPlannedCallLogFile ? (
                  <FilePill
                    name={monthlyPlannedCallLogFile.name}
                    onRemove={() => {
                      setMonthlyPlannedCallLogFile(null);
                      setMonthlyPlannedReport(null);
                      setMonthlyPlannedPreview(null);
                    }}
                    color="accent"
                  />
                ) : (
                  <EmptyHint
                    icon={<FileSpreadsheet className="h-5 w-5" />}
                    label="Choose call log"
                  />
                )}
              </UploadCard>
            </div>

            <Card className="border-border bg-card p-6 shadow-[var(--shadow-soft)]">
              <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-center">
                <div>
                  <div className="text-xs font-bold uppercase text-primary">Step 3</div>
                  <h2 className="mt-1 text-lg font-semibold text-foreground">
                    Generate monthly planned/unplanned report
                  </h2>
                  <p className="text-sm text-muted-foreground">
                    Employee Code matching is used first. Planned/unplanned totals, call averages,
                    and CP average time are added to the sample file.
                  </p>
                </div>
                <Button
                  size="lg"
                  onClick={onMonthlyPlannedProcess}
                  disabled={!canProcessMonthlyPlanned}
                  className="w-full min-w-44 shadow-[var(--shadow-elegant)] sm:w-auto"
                >
                  {monthlyPlannedProcessing ? (
                    <>
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      Processing...
                    </>
                  ) : (
                    <>
                      <Sparkles className="mr-2 h-4 w-4" />
                      Generate Report
                    </>
                  )}
                </Button>
              </div>
            </Card>

            {monthlyPlannedReport && (
              <Card className="border-border bg-card p-6 shadow-[var(--shadow-soft)]">
                <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-center">
                  <div className="flex items-start gap-3">
                    <div className="flex h-10 w-10 items-center justify-center rounded-full bg-success/15 text-success">
                      <CheckCircle2 className="h-5 w-5" />
                    </div>
                    <div>
                      <h3 className="text-lg font-semibold text-foreground">
                        Monthly planned/unplanned report ready
                      </h3>
                      <p className="text-sm text-muted-foreground">
                        {monthlyPlannedReport.matchedEmployees}/
                        {monthlyPlannedReport.totalEmployees} employee(s) matched from the sample
                        file.
                      </p>
                    </div>
                  </div>
                  <div className="grid w-full gap-2 sm:w-auto sm:grid-cols-2">
                    <Button
                      onClick={openMonthlyPlannedPreview}
                      variant="outline"
                      disabled={monthlyPlannedPreviewLoading}
                    >
                      {monthlyPlannedPreviewLoading ? (
                        <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                      ) : (
                        <Eye className="mr-2 h-4 w-4" />
                      )}
                      Preview
                    </Button>
                    <Button onClick={downloadMonthlyPlannedReport}>
                      <Download className="mr-2 h-4 w-4" />
                      Download
                    </Button>
                  </div>
                </div>

                <div className="mt-6 grid gap-4 md:grid-cols-3">
                  <Stat label="Call rows" value={monthlyPlannedReport.debug.sourceRows} />
                  <Stat label="Sample rows" value={monthlyPlannedReport.debug.templateRows} />
                  <Stat label="Matched" value={monthlyPlannedReport.matchedEmployees} />
                </div>

                {monthlyPlannedReport.preview.length > 0 && (
                  <div className="mt-6">
                    <div className="mb-2 text-sm font-semibold text-foreground">
                      Highest total calls
                    </div>
                    <div className="space-y-1.5">
                      {monthlyPlannedReport.preview.map((item, index) => (
                        <div
                          key={`${item.name}-${index}`}
                          className="flex items-center justify-between rounded-md border border-border bg-muted/40 px-3 py-2 text-sm"
                        >
                          <span>{item.name}</span>
                          <span className="font-semibold text-foreground">{item.total}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </Card>
            )}

            {(monthlyPlannedPreviewLoading || monthlyPlannedPreview) && (
              <Card
                ref={monthlyPlannedPreviewRef}
                className="border-border bg-card p-5 shadow-[var(--shadow-soft)]"
              >
                <div className="mb-3 text-sm font-semibold text-foreground">
                  {monthlyPlannedPreview?.name ?? "Opening monthly planned/unplanned preview"}
                </div>
                <div className="max-h-[70vh] overflow-auto rounded-md border border-border bg-background">
                  {monthlyPlannedPreviewLoading && !monthlyPlannedPreview ? (
                    <div className="flex items-center gap-2 p-4 text-sm text-muted-foreground">
                      <Loader2 className="h-4 w-4 animate-spin" />
                      Opening monthly planned/unplanned preview...
                    </div>
                  ) : (
                    <PreviewTable preview={monthlyPlannedPreview} readOnly />
                  )}
                </div>
              </Card>
            )}
          </div>

          <aside className="space-y-4">
            <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
              <div className="mb-4 flex items-center justify-between">
                <div>
                  <div className="text-xs font-bold uppercase text-primary">Run status</div>
                  <h2 className="mt-1 text-base font-semibold text-foreground">
                    Monthly summary workspace
                  </h2>
                </div>
                <ShieldCheck className="h-5 w-5 text-primary" />
              </div>
              <div className="space-y-3">
                <ChecklistItem
                  done={Boolean(monthlyPlannedCallLogFile)}
                  label="Call log attached"
                />
                <ChecklistItem
                  done={Boolean(monthlyPlannedTemplateFile)}
                  label="Sample file attached"
                />
                <ChecklistItem done={Boolean(monthlyPlannedReport)} label="Report generated" />
                <ChecklistItem done={Boolean(monthlyPlannedPreview)} label="Preview opened" />
              </div>
            </Card>
            <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
              <div className="text-xs font-bold uppercase text-primary">Columns added</div>
              <p className="mt-2 text-sm text-muted-foreground">
                Total Planned, Planned Avg, Total Unplanned, Unplanned Avg, Total Calls, Total Calls
                Avg, and CP Avg Time.
              </p>
            </Card>
          </aside>
        </main>
      </AppShell>
    );
  }

  if (activeModule === "history") {
    return (
      <AppShell
        activeModule={activeModule}
        dashboardLine={dashboardLine}
        user={authUser}
        profilePhoto={profilePhoto}
        onNavigate={setActiveModule}
        onOpenHistory={openHistory}
      >
        <ModulePageHeader
          icon={<HistoryIcon className="h-5 w-5" />}
          label="Archive"
          title="History"
          description="Saved report files, previews, and downloads."
        />

        <main>
          <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
            <div className="mb-4 flex flex-col items-start justify-between gap-3 sm:flex-row sm:items-center">
              <div>
                <div className="text-xs font-bold uppercase text-primary">Saved reports</div>
                <h2 className="mt-1 text-lg font-semibold text-foreground">
                  {historyItems.length} file(s)
                </h2>
              </div>
              <Button
                variant="outline"
                onClick={refreshHistory}
                disabled={historyLoading}
                className="w-full sm:w-auto"
              >
                {historyLoading ? (
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                ) : (
                  <HistoryIcon className="mr-2 h-4 w-4" />
                )}
                Refresh
              </Button>
            </div>

            {historyLoading ? (
              <div className="flex items-center gap-2 rounded-md border border-border bg-muted/30 p-4 text-sm text-muted-foreground">
                <Loader2 className="h-4 w-4 animate-spin" />
                Loading history...
              </div>
            ) : historyItems.length === 0 ? (
              <div className="rounded-md border border-dashed border-border bg-muted/20 p-6 text-sm text-muted-foreground">
                No reports have been saved yet. Generated Monthly Report files will appear here.
              </div>
            ) : (
              <div className="space-y-3">
                {historyItems.map((item) => (
                  <div
                    key={item.id}
                    className="flex flex-col gap-3 rounded-md border border-border bg-muted/20 p-4 md:flex-row md:items-center md:justify-between"
                  >
                    <div className="min-w-0">
                      <div className="truncate text-sm font-semibold text-foreground">
                        {item.fileName}
                      </div>
                      <div className="mt-1 flex flex-wrap gap-x-4 gap-y-1 text-xs text-muted-foreground">
                        <span>{formatHistoryDate(item.createdAt)}</span>
                        <span>{item.reportType ?? "Report"}</span>
                        <span>{item.dates.length} date(s)</span>
                        {item.pdfCount > 0 && <span>{item.pdfCount} PDF(s)</span>}
                        <span>
                          {item.matchedEmployees}/{item.totalEmployees} matched
                        </span>
                        <span>{formatBytes(item.size)}</span>
                      </div>
                    </div>
                    <div className="grid shrink-0 gap-2 sm:grid-cols-3 md:flex">
                      <Button
                        size="sm"
                        variant="outline"
                        onClick={() => previewHistoryItem(item.id)}
                        disabled={historyPreviewLoading}
                      >
                        {historyPreviewLoading ? (
                          <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                        ) : (
                          <Eye className="mr-2 h-4 w-4" />
                        )}
                        Preview
                      </Button>
                      <Button size="sm" onClick={() => downloadHistoryItem(item.id)}>
                        <Download className="mr-2 h-4 w-4" />
                        Download
                      </Button>
                      <Button
                        size="sm"
                        variant="outline"
                        onClick={() => removeHistoryItem(item.id)}
                      >
                        <Trash2 className="mr-2 h-4 w-4" />
                        Delete
                      </Button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </Card>

          {(historyPreviewLoading || historyPreview) && (
            <Card className="mt-6 border-border bg-card p-5 shadow-[var(--shadow-soft)]">
              <div className="mb-3 flex items-center justify-between gap-3">
                <div>
                  <div className="text-xs font-bold uppercase text-primary">Preview</div>
                  <h2 className="mt-1 text-base font-semibold text-foreground">
                    {historyPreview?.name ?? "Opening saved file"}
                  </h2>
                </div>
              </div>
              <div className="max-h-[70vh] overflow-auto rounded-md border border-border bg-background">
                {historyPreviewLoading && !historyPreview ? (
                  <div className="flex items-center gap-2 p-4 text-sm text-muted-foreground">
                    <Loader2 className="h-4 w-4 animate-spin" />
                    Opening history preview...
                  </div>
                ) : (
                  <PreviewTable preview={historyPreview} readOnly />
                )}
              </div>
            </Card>
          )}
        </main>
      </AppShell>
    );
  }

  return (
    <AppShell
      activeModule="monthly-report"
      dashboardLine={dashboardLine}
      user={authUser}
      profilePhoto={profilePhoto}
      onNavigate={setActiveModule}
      onOpenHistory={openHistory}
    >
      <ModulePageHeader
        icon={<FileSpreadsheet className="h-5 w-5" />}
        label="Report tool"
        title="Monthly Report"
        description="PDF to Excel report workbench."
      />

      <main className="grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_340px]">
        <div className="space-y-6">
          <div className="grid gap-4 sm:grid-cols-2">
            <UploadCard
              step={1}
              title="Monthly Report Sample"
              subtitle={excelFile ? "Sample file selected" : "Upload the monthly report sample"}
              icon={<FileSpreadsheet className="h-5 w-5" />}
              accept=".xlsx,.xls"
              ready={Boolean(excelFile)}
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
                <FilePill
                  name={excelFile.name}
                  onRemove={() => {
                    setExcelFile(null);
                    setReport(null);
                    setDownloadUrl(null);
                    setSheetPreview(null);
                    setPreviewDirty(false);
                  }}
                  color="primary"
                />
              ) : (
                <EmptyHint
                  icon={<FileSpreadsheet className="h-5 w-5" />}
                  label="Choose .xlsx file"
                />
              )}
            </UploadCard>

            <UploadCard
              step={2}
              title="Daily PDF Reports"
              subtitle={
                pdfFiles.length
                  ? `${pdfFiles.length} file(s), ${sortedDates.length} date(s)`
                  : "Upload one or many dated PDF reports"
              }
              icon={<FileText className="h-5 w-5" />}
              accept=".pdf"
              multiple
              ready={pdfFiles.length > 0}
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
                <EmptyHint icon={<FileText className="h-5 w-5" />} label="Choose PDF reports" />
              ) : (
                <div className="space-y-2">
                  <div className="flex items-center justify-between gap-3 text-xs font-medium text-muted-foreground">
                    <div className="flex items-center gap-2">
                      <CalendarDays className="h-3.5 w-3.5" />
                      {sortedDates.length} unique date(s)
                    </div>
                    <Button
                      type="button"
                      size="sm"
                      variant="outline"
                      onClick={clearAllPdfs}
                      className="h-7 px-2 text-xs"
                    >
                      Clear All
                    </Button>
                  </div>
                  <div className="max-h-44 space-y-1.5 overflow-y-auto pr-1">
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
          <Card className="border-border bg-card p-6 shadow-[var(--shadow-soft)]">
            <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-center">
              <div>
                <div className="text-xs font-bold uppercase text-primary">Step 3</div>
                <h2 className="mt-1 text-lg font-semibold text-foreground">
                  Generate monthly report
                </h2>
                <p className="text-sm text-muted-foreground">
                  Extraction, validation sheet, database history, and editable Excel export.
                </p>
              </div>
              <Button
                size="lg"
                onClick={onProcess}
                disabled={!canProcess}
                className="w-full min-w-48 shadow-[var(--shadow-elegant)] sm:w-auto"
              >
                {processing ? (
                  <>
                    <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                    Processing...
                  </>
                ) : (
                  <>
                    <Sparkles className="mr-2 h-4 w-4" />
                    Generate report
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
        </div>

        <aside className="space-y-4">
          <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
            <div className="mb-4 flex items-center justify-between">
              <div>
                <div className="text-xs font-bold uppercase text-primary">Run status</div>
                <h2 className="mt-1 text-base font-semibold text-foreground">Workspace</h2>
              </div>
              <ShieldCheck className="h-5 w-5 text-primary" />
            </div>
            <div className="space-y-3">
              <ChecklistItem done={Boolean(excelFile)} label="Master Excel attached" />
              <ChecklistItem done={pdfFiles.length > 0} label="Daily PDFs attached" />
              <ChecklistItem done={sortedDates.length > 0} label="Dates detected from files" />
              <ChecklistItem done={Boolean(report)} label="Report generated" />
              <ChecklistItem done={Boolean(sheetPreview)} label="Manual review opened" />
            </div>
          </Card>

          <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
            <div className="grid grid-cols-2 gap-3">
              <CompactMetric
                icon={<FileText className="h-4 w-4" />}
                label="PDFs"
                value={pdfFiles.length}
              />
              <CompactMetric
                icon={<CalendarDays className="h-4 w-4" />}
                label="Dates"
                value={sortedDates.length}
              />
              <CompactMetric
                icon={<BarChart3 className="h-4 w-4" />}
                label="Match"
                value={matchRate == null ? "-" : `${matchRate}%`}
              />
              <CompactMetric
                icon={<TableProperties className="h-4 w-4" />}
                label="Rows"
                value={report?.debug.totalEmployeesDetected ?? "-"}
              />
            </div>
          </Card>
        </aside>

        {/* Report */}
        {report && (
          <Card className="border-border bg-card p-4 shadow-[var(--shadow-soft)] sm:p-6 xl:col-span-2">
            <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-center">
              <div className="flex items-start gap-3">
                <div className="flex h-10 w-10 items-center justify-center rounded-full bg-success/15 text-success">
                  <CheckCircle2 className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="text-lg font-semibold text-foreground">Report ready</h3>
                  <p className="text-sm text-muted-foreground">
                    Updated <span className="font-medium">{report.fileName}</span> -{" "}
                    {report.totalEmployees} employees / {report.dates.length} dates /{" "}
                    {report.matchedEmployees} active performers
                  </p>
                </div>
              </div>
              <div className="grid w-full gap-2 sm:w-auto sm:grid-cols-2">
                <Button
                  onClick={openFilePreview}
                  size="lg"
                  variant="outline"
                  disabled={previewLoading || previewSaving}
                >
                  {previewLoading ? (
                    <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  ) : (
                    <Eye className="mr-2 h-4 w-4" />
                  )}
                  Review & Edit
                </Button>
                <Button
                  onClick={triggerDownload}
                  size="lg"
                  variant="default"
                  disabled={previewSaving}
                >
                  {previewSaving ? (
                    <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  ) : (
                    <Download className="mr-2 h-4 w-4" />
                  )}
                  Download File
                </Button>
              </div>
            </div>

            <div className="mt-6 grid gap-4 md:grid-cols-3">
              <Stat label="Employees" value={report.totalEmployees} />
              <Stat label="Days covered" value={report.dates.length} />
              <Stat
                label="Needs review"
                value={report.unmatchedNames.length}
                tone={report.unmatchedNames.length ? "warning" : "default"}
              />
            </div>

            <div className="mt-6 grid gap-3 rounded-md border border-border bg-muted/20 p-4 md:grid-cols-3">
              <Stat label="PDFs uploaded" value={report.debug.totalPdfsUploaded} />
              <Stat label="Detected rows" value={report.debug.totalEmployeesDetected} />
              <Stat label="Matched rows" value={report.debug.totalMatched} />
              <Stat
                label="Review rows"
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

            {(previewLoading || sheetPreview) && (
              <div ref={previewRef} className="mt-6">
                <div className="mb-2 flex flex-col items-start justify-between gap-3 sm:flex-row sm:items-center">
                  <div className="text-sm font-semibold text-foreground">
                    {sheetPreview?.name ?? "Opening file preview"}
                  </div>
                  <div className="flex w-full flex-col gap-2 sm:w-auto sm:flex-row sm:items-center">
                    <div className="text-xs text-muted-foreground">
                      Edit any cell inline before downloading.
                    </div>
                    {sheetPreview && (
                      <Button
                        size="sm"
                        variant="outline"
                        onClick={savePreviewEdits}
                        disabled={!previewDirty || previewSaving}
                      >
                        {previewSaving ? (
                          <Loader2 className="mr-1.5 h-3.5 w-3.5 animate-spin" />
                        ) : (
                          <Save className="mr-1.5 h-3.5 w-3.5" />
                        )}
                        Apply Edits
                      </Button>
                    )}
                  </div>
                </div>
                <div className="max-h-[70vh] overflow-auto rounded-md border border-border bg-background">
                  {previewLoading && !sheetPreview ? (
                    <div className="flex items-center gap-2 p-4 text-sm text-muted-foreground">
                      <Loader2 className="h-4 w-4 animate-spin" />
                      Opening Excel preview...
                    </div>
                  ) : (
                    <table className="min-w-max border-collapse text-xs">
                      <tbody>
                        {sheetPreview?.rows.map((row, rowIndex) => (
                          <tr key={`row-${rowIndex}`}>
                            {row.map((cell) => (
                              <td
                                key={cell.key}
                                className="min-w-24 p-0 align-middle"
                                style={cell.style}
                              >
                                <input
                                  className="h-full min-h-8 w-full min-w-24 bg-transparent px-2 py-1 text-inherit outline-none focus:bg-primary/10 focus:ring-1 focus:ring-primary"
                                  value={cell.value}
                                  onChange={(event) =>
                                    updatePreviewCell(
                                      cell.rowNumber,
                                      cell.colNumber,
                                      event.target.value,
                                    )
                                  }
                                  aria-label={`Cell ${cell.rowNumber}, ${cell.colNumber}`}
                                />
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            )}

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
                  {report.unmatchedNames.length} row(s) in PDFs need review against your Active
                  Members sheet
                </div>
                <p className="text-xs text-muted-foreground">
                  These were kept in the Excel output as review rows. Examples:{" "}
                  {report.unmatchedNames.slice(0, 6).join(", ")}
                  {report.unmatchedNames.length > 6 ? "..." : ""}
                </p>
              </div>
            )}
          </Card>
        )}
      </main>
    </AppShell>
  );
}

function AppShell({
  activeModule,
  dashboardLine,
  user,
  profilePhoto,
  onNavigate,
  onOpenHistory,
  children,
}: {
  activeModule: ActiveModule;
  dashboardLine: string;
  user: User;
  profilePhoto: string | null;
  onNavigate: (module: ActiveModule) => void;
  onOpenHistory: () => void;
  children: React.ReactNode;
}) {
  const displayName = getUserDisplayName(user);
  const tools = [
    {
      id: "daily-report" as const,
      label: "Daily Report",
      icon: <TableProperties className="h-4 w-4" />,
      description: "Call log",
    },
    {
      id: "monthly-report" as const,
      label: "Monthly Report",
      icon: <FileSpreadsheet className="h-4 w-4" />,
      description: "PDF sync",
    },
    {
      id: "doctor-coverage" as const,
      label: "Doctor Coverage",
      icon: <BarChart3 className="h-4 w-4" />,
      description: "Coverage %",
    },
    {
      id: "monthly-planned" as const,
      label: "Planned Summary",
      icon: <CalendarDays className="h-4 w-4" />,
      description: "Averages",
    },
  ];

  return (
    <div className="min-h-screen bg-background text-foreground">
      <Toaster richColors position="top-center" />
      <aside className="neuro-sidebar app-sidebar fixed z-40 flex h-auto flex-col rounded-2xl p-2 sm:p-3 lg:inset-y-4 lg:left-4 lg:right-auto lg:w-72">
        <button
          type="button"
          onClick={() => onNavigate(null)}
          className="neuro-brand flex w-full items-center gap-3 rounded-xl p-2.5 text-left transition sm:p-3"
        >
          <BrandMark icon={<Activity className="h-5 w-5" />} />
          <div className="min-w-0">
            <div className="truncate text-base font-bold text-foreground">{APP_NAME}</div>
            <div className="truncate text-xs text-muted-foreground">{dashboardLine}</div>
          </div>
        </button>

        <nav className="app-sidebar-nav mt-2 flex gap-2 overflow-x-auto pb-1 lg:mt-3 lg:flex-1 lg:flex-col lg:overflow-visible lg:pb-0">
          <SidebarButton
            active={!activeModule}
            icon={<Home className="h-4 w-4" />}
            label="Dashboard"
            description="Overview"
            onClick={() => onNavigate(null)}
          />
          {tools.map((tool) => (
            <SidebarButton
              key={tool.id}
              active={activeModule === tool.id}
              icon={tool.icon}
              label={tool.label}
              description={tool.description}
              onClick={() => onNavigate(tool.id)}
            />
          ))}
          <SidebarButton
            active={activeModule === "history"}
            icon={<HistoryIcon className="h-4 w-4" />}
            label="History"
            description="Files"
            onClick={onOpenHistory}
          />
          <SidebarButton
            active={activeModule === "profile"}
            icon={<UserCircle className="h-4 w-4" />}
            label="Profile"
            description="Theme"
            onClick={() => onNavigate("profile")}
            className="lg:hidden"
          />
        </nav>

        <button
          type="button"
          onClick={() => onNavigate("profile")}
          className={`neuro-user mt-3 hidden items-center gap-3 rounded-xl p-3 text-left transition lg:flex ${
            activeModule === "profile" ? "is-active" : ""
          }`}
        >
          <UserAvatar user={user} photo={profilePhoto} className="h-10 w-10 rounded-xl" />
          <span className="hidden min-w-0 lg:block">
            <span className="block truncate text-sm font-semibold text-foreground">
              {displayName}
            </span>
            <span className="block text-xs text-muted-foreground">Profile & theme</span>
          </span>
        </button>
      </aside>

      <div className="px-3 pb-8 pt-40 sm:px-5 md:pt-36 lg:ml-80 lg:px-8 lg:pt-8">
        <div className="mx-auto max-w-[1540px] animate-in fade-in slide-in-from-bottom-2 duration-500">
          {children}
        </div>
      </div>
    </div>
  );
}

function SidebarButton({
  active,
  icon,
  label,
  description,
  onClick,
  className = "",
}: {
  active: boolean;
  icon: React.ReactNode;
  label: string;
  description: string;
  onClick: () => void;
  className?: string;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`neuro-nav-item group flex min-w-[4.9rem] flex-col items-center justify-center gap-1 rounded-xl px-2 py-2 text-center transition sm:min-w-[5.6rem] lg:min-h-14 lg:w-full lg:min-w-0 lg:flex-row lg:justify-start lg:gap-3 lg:px-3 lg:text-left ${
        active ? "is-active" : ""
      } ${className}`}
    >
      <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg text-primary transition group-hover:scale-105 lg:h-9 lg:w-9">
        {icon}
      </span>
      <span className="block min-w-0">
        <span className="block max-w-[4.6rem] truncate text-[0.7rem] font-semibold leading-tight sm:max-w-[5.1rem] lg:max-w-none lg:text-sm">
          {label}
        </span>
        <span className="hidden truncate text-xs text-muted-foreground lg:block">
          {description}
        </span>
      </span>
    </button>
  );
}

function ModulePageHeader({
  icon,
  label,
  title,
  description,
}: {
  icon: React.ReactNode;
  label: string;
  title: string;
  description: string;
}) {
  return (
    <section className="neuro-panel mb-5 p-4 sm:mb-6 sm:p-5">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <div className="flex min-w-0 items-center gap-3 sm:gap-4">
          <BrandMark icon={icon} />
          <div className="min-w-0">
            <div className="text-xs font-bold uppercase tracking-wide text-primary">{label}</div>
            <h1 className="mt-1 break-words text-xl font-bold text-foreground sm:text-2xl">
              {title}
            </h1>
            <p className="text-sm text-muted-foreground">{description}</p>
          </div>
        </div>
      </div>
    </section>
  );
}

function DashboardHome({
  historyItems,
  historyLoading,
  onNavigate,
  onOpenHistory,
}: {
  historyItems: ReportHistoryItem[];
  historyLoading: boolean;
  onNavigate: (module: ActiveModule) => void;
  onOpenHistory: () => void;
}) {
  const latestItems = historyItems.slice(0, 5);
  const totalReports = historyItems.length;
  const totalFiles = historyItems.reduce((sum, item) => sum + Math.max(item.pdfCount, 1), 0);
  const totalEmployees = historyItems.reduce((sum, item) => sum + item.totalEmployees, 0);
  const totalMatched = historyItems.reduce((sum, item) => sum + item.matchedEmployees, 0);
  const hasMatchData = totalEmployees > 0;
  const matchRate = hasMatchData ? Math.round((totalMatched / totalEmployees) * 100) : 0;
  const reviewQueue = historyItems.reduce(
    (sum, item) => sum + Math.max(item.totalEmployees - item.matchedEmployees, 0),
    0,
  );
  const totalSize = historyItems.reduce((sum, item) => sum + item.size, 0);
  const reportMix = buildReportMix(historyItems);

  return (
    <main className="space-y-6">
      <section className="neuro-hero p-4 sm:p-6">
        <div className="grid min-w-0 gap-5 lg:grid-cols-[minmax(0,1fr)_320px] xl:grid-cols-[minmax(0,1fr)_360px]">
          <div>
            <div className="text-xs font-bold uppercase tracking-wide text-primary">
              Admin dashboard
            </div>
            <h1 className="mt-2 text-2xl font-bold text-foreground sm:text-4xl">{APP_NAME}</h1>
            <p className="mt-3 max-w-2xl text-sm leading-6 text-muted-foreground">
              Monitor report activity, open tools from the sidebar, and keep generated files moving
              through a cleaner reporting workflow.
            </p>
            <div className="mt-5 grid gap-3 sm:flex sm:flex-wrap">
              <Button onClick={() => onNavigate("daily-report")} className="neuro-button">
                <Zap className="mr-2 h-4 w-4" />
                Start daily report
              </Button>
              <Button variant="outline" onClick={onOpenHistory} className="neuro-button-muted">
                <HistoryIcon className="mr-2 h-4 w-4" />
                View history
              </Button>
            </div>
          </div>
          <div className="neuro-inset flex min-h-44 items-center justify-center p-4 sm:min-h-48 sm:p-5">
            <CircularProgress
              value={matchRate}
              label={hasMatchData ? "Match rate" : "No data"}
              displayValue={hasMatchData ? `${matchRate}%` : "-"}
            />
          </div>
        </div>
      </section>

      <section className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
        <DashboardMetric
          icon={<FileText className="h-5 w-5" />}
          label="Saved reports"
          value={totalReports}
          hint={historyLoading ? "Refreshing..." : "History entries"}
        />
        <DashboardMetric
          icon={<Users className="h-5 w-5" />}
          label="Matched employees"
          value={totalMatched}
          hint={`${totalEmployees || 0} total rows`}
        />
        <DashboardMetric
          icon={<Database className="h-5 w-5" />}
          label="Files handled"
          value={totalFiles}
          hint={formatBytes(totalSize)}
        />
        <DashboardMetric
          icon={<AlertTriangle className="h-5 w-5" />}
          label="Needs review"
          value={hasMatchData ? reviewQueue : "-"}
          hint={
            hasMatchData
              ? reviewQueue
                ? "Unmatched employees"
                : "All matched cleanly"
              : "No saved report rows yet"
          }
        />
      </section>

      <section className="grid gap-6 xl:grid-cols-[minmax(0,1.2fr)_minmax(320px,0.8fr)]">
        <div className="neuro-panel p-5">
          <div className="mb-5 flex items-center justify-between gap-3">
            <div>
              <div className="text-xs font-bold uppercase tracking-wide text-primary">Activity</div>
              <h2 className="text-lg font-semibold text-foreground">Recent report flow</h2>
            </div>
            <LineChart className="h-5 w-5 text-primary" />
          </div>
          <div className="grid gap-3 sm:grid-cols-2">
            {[
              { label: "Daily Report", value: 72 },
              { label: "Monthly Report", value: 88 },
              { label: "Doctor Coverage", value: 64 },
              { label: "Planned Summary", value: 78 },
            ].map((item) => (
              <div key={item.label} className="neuro-inset p-4">
                <div className="mb-3 flex items-center justify-between text-sm">
                  <span className="font-semibold text-foreground">{item.label}</span>
                  <span className="text-muted-foreground">{item.value}%</span>
                </div>
                <div className="h-3 overflow-hidden rounded-full bg-background shadow-[var(--shadow-inset-sm)]">
                  <div
                    className="h-full rounded-full bg-primary transition-all duration-700"
                    style={{ width: `${item.value}%` }}
                  />
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="neuro-panel p-5">
          <div className="mb-5 flex items-center justify-between gap-3">
            <div>
              <div className="text-xs font-bold uppercase tracking-wide text-primary">Mix</div>
              <h2 className="text-lg font-semibold text-foreground">Report categories</h2>
            </div>
            <PieChart className="h-5 w-5 text-primary" />
          </div>
          <div className="space-y-3">
            {reportMix.map((item) => (
              <div key={item.label} className="flex items-center gap-3">
                <span className="h-3 w-3 rounded-full" style={{ backgroundColor: item.color }} />
                <span className="min-w-0 flex-1 text-sm font-medium text-foreground">
                  {item.label}
                </span>
                <span className="text-sm text-muted-foreground">{item.value}</span>
              </div>
            ))}
          </div>
        </div>
      </section>

      <section className="neuro-panel p-5">
        <div className="mb-5 flex items-center justify-between gap-3">
          <div>
            <div className="text-xs font-bold uppercase tracking-wide text-primary">Logs</div>
            <h2 className="text-lg font-semibold text-foreground">Latest saved reports</h2>
          </div>
          <Clock3 className="h-5 w-5 text-primary" />
        </div>
        {latestItems.length ? (
          <div className="dashboard-logs-scroll grid max-h-[19rem] gap-3 overflow-y-auto pr-1">
            {latestItems.map((item) => (
              <button
                key={item.id}
                type="button"
                onClick={onOpenHistory}
                className="neuro-list-row grid gap-2 rounded-xl p-4 text-left transition md:grid-cols-[minmax(0,1fr)_140px_120px] md:gap-3"
              >
                <div className="min-w-0">
                  <div className="truncate text-sm font-semibold text-foreground">
                    {item.fileName}
                  </div>
                  <div className="mt-1 text-xs text-muted-foreground">
                    {formatHistoryDate(item.createdAt)}
                  </div>
                </div>
                <div className="text-sm text-muted-foreground">{item.reportType ?? "Report"}</div>
                <div className="text-sm font-semibold text-primary">
                  {item.matchedEmployees}/{item.totalEmployees}
                </div>
              </button>
            ))}
          </div>
        ) : (
          <div className="neuro-inset p-5 text-sm text-muted-foreground">
            No generated reports yet. Start with Monthly Report or Daily Report from the sidebar.
          </div>
        )}
      </section>
    </main>
  );
}

function UserAvatar({
  user,
  photo,
  className = "",
}: {
  user: User;
  photo: string | null;
  className?: string;
}) {
  const displayName = getUserDisplayName(user);
  return (
    <Avatar
      className={`bg-primary text-primary-foreground shadow-[var(--shadow-soft)] ${className}`}
    >
      {photo && <AvatarImage src={photo} alt={displayName} className="object-cover" />}
      <AvatarFallback className="bg-primary text-sm font-bold text-primary-foreground">
        {getUserInitials(displayName)}
      </AvatarFallback>
    </Avatar>
  );
}

function getUserDisplayName(user: User): string {
  const metadata = user.user_metadata ?? {};
  const metadataName =
    stringValue(metadata.full_name) ||
    stringValue(metadata.name) ||
    stringValue(metadata.display_name);
  if (metadataName) return toTitleCase(metadataName);

  const localPart = (user.email ?? "Profile").split("@")[0] || "Profile";
  const words = localPart
    .replace(/[._-]+/g, " ")
    .replace(/\d+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return toTitleCase(splitCompactDisplayName(words || localPart || "Profile"));
}

function getUserInitials(displayName: string): string {
  const words = displayName.split(/\s+/).filter(Boolean);
  return (words[0]?.[0] ?? "U").toUpperCase() + (words[1]?.[0] ?? "").toUpperCase();
}

function toTitleCase(value: string): string {
  return value
    .trim()
    .split(/\s+/)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(" ");
}

function splitCompactDisplayName(value: string): string {
  if (value.includes(" ")) return value;
  const compact = value.toLowerCase();
  const commonSuffixes = ["awan", "khan", "ahmad", "ahmed", "ali", "butt"];
  const suffix = commonSuffixes.find(
    (item) => compact.endsWith(item) && compact.length > item.length + 1,
  );
  return suffix ? `${compact.slice(0, -suffix.length)} ${suffix}` : value;
}

function stringValue(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

function profilePhotoStorageKey(user: User): string {
  return `profile-photo:${user.id}`;
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function ProfilePage({
  user,
  profilePhoto,
  color,
  mode,
  signingOut,
  onProfilePhotoChange,
  onChange,
  onToggleMode,
  onSignOut,
}: {
  user: User;
  profilePhoto: string | null;
  color: string;
  mode: "light" | "dark";
  signingOut: boolean;
  onProfilePhotoChange: (photo: string | null) => void;
  onChange: (color: string) => void;
  onToggleMode: () => void;
  onSignOut: () => void;
}) {
  const photoInputRef = useRef<HTMLInputElement>(null);
  const displayName = getUserDisplayName(user);

  const onProfilePhotoSelected = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      toast.error("Please choose an image file.");
      return;
    }
    if (file.size > 2 * 1024 * 1024) {
      toast.error("Profile photo must be 2 MB or smaller.");
      return;
    }

    try {
      onProfilePhotoChange(await readFileAsDataUrl(file));
      toast.success("Profile photo updated.");
    } catch (error) {
      console.error("Profile photo read failed", error);
      toast.error("Unable to use this profile photo.");
    }
  };

  return (
    <main className="grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_340px]">
      <section className="neuro-panel p-4 sm:p-6">
        <div className="flex min-w-0 items-center gap-3 sm:gap-4">
          <UserAvatar user={user} photo={profilePhoto} className="h-14 w-14 rounded-2xl" />
          <div className="min-w-0">
            <div className="text-xs font-bold uppercase tracking-wide text-primary">
              User profile
            </div>
            <h1 className="mt-1 truncate text-xl font-bold text-foreground sm:text-2xl">
              {displayName}
            </h1>
            <p className="text-sm text-muted-foreground">
              Account controls and interface settings.
            </p>
          </div>
        </div>

        <div className="mt-8 grid gap-4 sm:grid-cols-2">
          <div className="neuro-inset p-4">
            <div className="text-xs font-bold uppercase tracking-wide text-primary">Name</div>
            <div className="mt-2 truncate text-sm font-semibold text-foreground">{displayName}</div>
          </div>
          <div className="neuro-inset p-4">
            <div className="text-xs font-bold uppercase tracking-wide text-primary">Theme</div>
            <div className="mt-2 text-sm font-semibold text-foreground">
              {mode === "dark" ? "Dark mode" : "Light mode"}
            </div>
          </div>
        </div>
      </section>

      <aside className="space-y-6">
        <section className="neuro-panel p-5">
          <div className="mb-4 flex items-center gap-2 text-sm font-semibold text-foreground">
            <UserCircle className="h-4 w-4 text-primary" />
            Profile photo
          </div>
          <div className="mb-4 flex items-center gap-4">
            <UserAvatar user={user} photo={profilePhoto} className="h-16 w-16 rounded-2xl" />
            <div className="min-w-0 text-sm text-muted-foreground">
              This photo appears in the profile shortcut.
            </div>
          </div>
          <input
            ref={photoInputRef}
            type="file"
            accept="image/*"
            className="hidden"
            onChange={onProfilePhotoSelected}
          />
          <div className="grid gap-2 sm:grid-cols-2 xl:grid-cols-1">
            <Button
              type="button"
              onClick={() => photoInputRef.current?.click()}
              className="neuro-button w-full"
            >
              <Upload className="mr-2 h-4 w-4" />
              Upload photo
            </Button>
            {profilePhoto && (
              <Button
                type="button"
                variant="outline"
                onClick={() => onProfilePhotoChange(null)}
                className="neuro-button-muted w-full"
              >
                <X className="mr-2 h-4 w-4" />
                Remove photo
              </Button>
            )}
          </div>
        </section>

        <section className="neuro-panel p-5">
          <div className="mb-4 flex items-center gap-2 text-sm font-semibold text-foreground">
            <Settings className="h-4 w-4 text-primary" />
            Interface
          </div>
          <Button type="button" onClick={onToggleMode} className="neuro-button mb-4 w-full">
            {mode === "dark" ? <Sun className="mr-2 h-4 w-4" /> : <Moon className="mr-2 h-4 w-4" />}
            {mode === "dark" ? "Use light theme" : "Use dark theme"}
          </Button>
          <div className="grid grid-cols-5 gap-2">
            {THEME_COLORS.map((item) => (
              <button
                key={item}
                type="button"
                className={`h-9 rounded-lg transition hover:scale-105 ${
                  color.toLowerCase() === item.toLowerCase() ? "ring-2 ring-primary" : ""
                }`}
                style={{ backgroundColor: item }}
                onClick={() => onChange(item)}
                aria-label={`Use color ${item}`}
              />
            ))}
          </div>
          <input
            type="color"
            value={color}
            onChange={(event) => onChange(event.target.value)}
            className="mt-4 h-10 w-full cursor-pointer rounded-xl border border-border bg-background p-1 shadow-[var(--shadow-inset)]"
            aria-label="Choose custom theme color"
          />
        </section>

        <section className="neuro-panel p-5">
          <Button
            type="button"
            onClick={onSignOut}
            disabled={signingOut}
            variant="outline"
            className="neuro-button-muted w-full"
          >
            {signingOut ? (
              <Loader2 className="mr-2 h-4 w-4 animate-spin" />
            ) : (
              <LogOut className="mr-2 h-4 w-4" />
            )}
            Sign out
          </Button>
        </section>
      </aside>
    </main>
  );
}

function DashboardMetric({
  icon,
  label,
  value,
  hint,
}: {
  icon: React.ReactNode;
  label: string;
  value: number | string;
  hint: string;
}) {
  return (
    <div className="neuro-panel p-4 transition hover:-translate-y-0.5 sm:p-5">
      <div className="mb-4 flex items-center justify-between">
        <span className="flex h-11 w-11 items-center justify-center rounded-xl bg-primary text-primary-foreground shadow-[var(--shadow-soft)]">
          {icon}
        </span>
        <Activity className="h-4 w-4 text-primary/70" />
      </div>
      <div className="text-2xl font-bold text-foreground">{value}</div>
      <div className="mt-1 text-sm font-semibold text-foreground">{label}</div>
      <div className="mt-1 text-xs text-muted-foreground">{hint}</div>
    </div>
  );
}

function CircularProgress({
  value,
  label,
  displayValue,
}: {
  value: number;
  label: string;
  displayValue?: string;
}) {
  const clamped = Math.max(0, Math.min(100, value));
  return (
    <div className="relative flex h-36 w-36 items-center justify-center rounded-full shadow-[var(--shadow-elegant)] sm:h-40 sm:w-40">
      <div
        className="absolute inset-0 rounded-full"
        style={{
          background: `conic-gradient(var(--primary) ${clamped}%, color-mix(in srgb, var(--muted) 75%, transparent) 0)`,
        }}
      />
      <div className="absolute inset-4 rounded-full bg-card shadow-[var(--shadow-inset)]" />
      <div className="relative text-center">
        <div className="text-2xl font-bold text-foreground sm:text-3xl">
          {displayValue ?? `${clamped}%`}
        </div>
        <div className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">
          {label}
        </div>
      </div>
    </div>
  );
}

function buildReportMix(items: ReportHistoryItem[]) {
  const colors = ["#0b6f6a", "#2563eb", "#db2777", "#ca8a04"];
  const counts = new Map<string, number>();
  for (const item of items) {
    const key = item.reportType ?? "Report";
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  const entries = [...counts.entries()];
  if (!entries.length) {
    return [
      { label: "Monthly Report", value: 0, color: colors[0] },
      { label: "Daily Report", value: 0, color: colors[1] },
      { label: "Doctor Coverage", value: 0, color: colors[2] },
      { label: "Planned Summary", value: 0, color: colors[3] },
    ];
  }
  return entries.map(([label, value], index) => ({
    label,
    value,
    color: colors[index % colors.length],
  }));
}

function AuthLoadingScreen({
  color: _color,
  mode: _mode,
  onChange: _onChange,
  onToggleMode: _onToggleMode,
}: {
  color: string;
  mode: "light" | "dark";
  onChange: (color: string) => void;
  onToggleMode: () => void;
}) {
  return (
    <div className="min-h-screen bg-background">
      <Toaster richColors position="top-center" />
      <main className="flex min-h-screen items-center justify-center px-5">
        <div className="neuro-panel flex items-center gap-3 px-5 py-4 text-sm font-medium text-foreground">
          <Loader2 className="h-4 w-4 animate-spin text-primary" />
          Checking login session...
        </div>
      </main>
    </div>
  );
}

function AuthScreen({
  color: _color,
  mode: _mode,
  onChange: _onChange,
  onToggleMode: _onToggleMode,
}: {
  color: string;
  mode: "light" | "dark";
  onChange: (color: string) => void;
  onToggleMode: () => void;
}) {
  const [authMode, setAuthMode] = useState<"login" | "signup">("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [submitting, setSubmitting] = useState(false);

  const isSignup = authMode === "signup";

  const submitAuth = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setSubmitting(true);

    try {
      const normalizedEmail = email.trim();
      if (!normalizedEmail || !password) throw new Error("Enter your email and password.");

      const result = isSignup
        ? await supabase.auth.signUp({
            email: normalizedEmail,
            password,
            options: { emailRedirectTo: window.location.origin },
          })
        : await supabase.auth.signInWithPassword({
            email: normalizedEmail,
            password,
          });

      if (result.error) throw result.error;

      if (isSignup && !result.data.session) {
        toast.success("Account created. Check your email to confirm your login.");
      } else {
        if (result.data.session) {
          const deviceResult = await enforceDeviceLimit(result.data.session);
          if (!deviceResult.allowed) {
            toast.error(deviceResult.message ?? "This account is already active on 2 devices.");
            return;
          }
        }
        toast.success("Signed in.");
      }
    } catch (error) {
      console.error("Authentication failed", error);
      toast.error(error instanceof Error ? error.message : "Authentication failed.");
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="min-h-screen bg-background">
      <Toaster richColors position="top-center" />
      <main className="mx-auto flex min-h-screen max-w-[1720px] items-center justify-center px-3 py-8 sm:px-8 sm:py-14">
        <Card className="neuro-panel w-full max-w-md border-border bg-card p-4 sm:p-6">
          <div className="mb-6 flex items-center gap-3">
            <BrandMark icon={<ShieldCheck className="h-5 w-5" />} />
            <div>
              <h1 className="text-lg font-bold text-foreground sm:text-xl">{APP_NAME}</h1>
              <p className="text-sm text-muted-foreground">Sign in to continue</p>
            </div>
          </div>

          <div className="neuro-inset mb-5 grid grid-cols-2 p-1">
            <button
              type="button"
              onClick={() => setAuthMode("login")}
              className={`h-9 rounded-md text-sm font-semibold transition ${
                !isSignup ? "bg-card text-foreground shadow-sm" : "text-muted-foreground"
              }`}
            >
              Login
            </button>
            <button
              type="button"
              onClick={() => setAuthMode("signup")}
              className={`h-9 rounded-md text-sm font-semibold transition ${
                isSignup ? "bg-card text-foreground shadow-sm" : "text-muted-foreground"
              }`}
            >
              Create
            </button>
          </div>

          <form className="space-y-4" onSubmit={submitAuth}>
            <label className="block text-sm font-medium text-foreground">
              Email
              <span className="neuro-inset mt-1 flex items-center gap-2 px-3 py-2">
                <Mail className="h-4 w-4 shrink-0 text-primary" />
                <input
                  value={email}
                  onChange={(event) => setEmail(event.target.value)}
                  type="email"
                  autoComplete="email"
                  className="h-8 min-w-0 flex-1 bg-transparent text-sm outline-none"
                  placeholder="you@example.com"
                  required
                />
              </span>
            </label>

            <label className="block text-sm font-medium text-foreground">
              Password
              <span className="neuro-inset mt-1 flex items-center gap-2 px-3 py-2">
                <Lock className="h-4 w-4 shrink-0 text-primary" />
                <input
                  value={password}
                  onChange={(event) => setPassword(event.target.value)}
                  type="password"
                  autoComplete={isSignup ? "new-password" : "current-password"}
                  minLength={6}
                  className="h-8 min-w-0 flex-1 bg-transparent text-sm outline-none"
                  placeholder="At least 6 characters"
                  required
                />
              </span>
            </label>

            <Button type="submit" className="neuro-button h-11 w-full" disabled={submitting}>
              {submitting ? (
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
              ) : isSignup ? (
                <UserPlus className="mr-2 h-4 w-4" />
              ) : (
                <ShieldCheck className="mr-2 h-4 w-4" />
              )}
              {isSignup ? "Create account" : "Login"}
            </Button>
          </form>
        </Card>
      </main>
    </div>
  );
}

function AccountControls({
  user,
  signingOut,
  onSignOut,
}: {
  user: User;
  signingOut: boolean;
  onSignOut: () => void;
}) {
  const displayName = getUserDisplayName(user);
  return (
    <div className="flex flex-col gap-2 sm:flex-row sm:items-center">
      <div className="max-w-full truncate rounded-md border border-white/40 bg-white/10 px-3 py-2 text-sm font-medium text-[var(--header-foreground)]">
        {displayName}
      </div>
      <Button
        size="sm"
        onClick={onSignOut}
        disabled={signingOut}
        className="self-start border border-white/70 bg-white text-primary shadow-sm hover:bg-white/90"
      >
        {signingOut ? (
          <Loader2 className="mr-2 h-4 w-4 animate-spin" />
        ) : (
          <LogOut className="mr-2 h-4 w-4" />
        )}
        Sign out
      </Button>
    </div>
  );
}

function applyThemeColor(color: string) {
  const root = document.documentElement;
  const foreground = readableTextColor(color);
  root.style.setProperty("--primary", color);
  root.style.setProperty("--primary-foreground", foreground);
  root.style.setProperty("--primary-glow", `color-mix(in srgb, ${color} 65%, white)`);
  root.style.setProperty("--accent", color);
  root.style.setProperty("--accent-foreground", foreground);
  root.style.setProperty("--header-surface", color);
  root.style.setProperty("--header-foreground", foreground);
  root.style.setProperty(
    "--header-muted",
    foreground === "#ffffff"
      ? `color-mix(in srgb, ${color} 18%, white)`
      : `color-mix(in srgb, ${color} 52%, black)`,
  );
  root.style.setProperty("--ring", color);
  root.style.setProperty(
    "--gradient-hero",
    `linear-gradient(135deg, ${color}, color-mix(in srgb, ${color} 68%, white))`,
  );
  root.style.setProperty(
    "--shadow-elegant",
    `16px 16px 36px color-mix(in srgb, ${color} 24%, black 20%), -14px -14px 34px color-mix(in srgb, white 92%, ${color})`,
  );
}

function applyThemeMode(mode: "light" | "dark") {
  document.documentElement.classList.toggle("dark", mode === "dark");
}

function readableTextColor(hex: string): string {
  const rgb = hexToRgb(hex);
  if (!rgb) return "#ffffff";
  const luminance = (0.2126 * rgb.r + 0.7152 * rgb.g + 0.0722 * rgb.b) / 255;
  return luminance > 0.62 ? "#111827" : "#ffffff";
}

function hexToRgb(hex: string): { r: number; g: number; b: number } | null {
  const normalized = hex.replace("#", "");
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return null;
  return {
    r: parseInt(normalized.slice(0, 2), 16),
    g: parseInt(normalized.slice(2, 4), 16),
    b: parseInt(normalized.slice(4, 6), 16),
  };
}

/* ---------- history storage ---------- */

const REPORT_HISTORY_DB = "report-history-db";
const REPORT_HISTORY_STORE = "reports";

function openReportHistoryDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(REPORT_HISTORY_DB, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(REPORT_HISTORY_STORE)) {
        db.createObjectStore(REPORT_HISTORY_STORE, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function addReportHistory(item: StoredReportHistoryItem): Promise<void> {
  const db = await openReportHistoryDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(REPORT_HISTORY_STORE, "readwrite");
    tx.objectStore(REPORT_HISTORY_STORE).put(item);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

async function saveReportHistoryAndDatabase(item: StoredReportHistoryItem): Promise<void> {
  await addReportHistory(item);
  try {
    await syncGeneratedReportToDatabase(item);
  } catch (error) {
    console.warn("Generated report database sync failed", error);
  }
}

async function listStoredReportHistory(userId: string): Promise<StoredReportHistoryItem[]> {
  const db = await openReportHistoryDb();
  const items = await new Promise<StoredReportHistoryItem[]>((resolve, reject) => {
    const tx = db.transaction(REPORT_HISTORY_STORE, "readonly");
    const request = tx.objectStore(REPORT_HISTORY_STORE).getAll();
    request.onsuccess = () => resolve(request.result ?? []);
    request.onerror = () => reject(request.error);
  });
  db.close();
  return items
    .filter((item) => item.userId === userId)
    .sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
}

async function syncStoredReportHistoryToDatabase(userId: string): Promise<void> {
  try {
    const items = await listStoredReportHistory(userId);
    for (const item of items) {
      try {
        await syncGeneratedReportToDatabase(item);
      } catch (error) {
        console.warn("Stored report database backfill failed", item.fileName, error);
      }
    }
  } catch (error) {
    console.warn("Unable to read stored report history for database backfill", error);
  }
}

async function listReportHistory(userId: string): Promise<ReportHistoryItem[]> {
  const [storedItems, databaseItems] = await Promise.all([
    listStoredReportHistory(userId),
    listGeneratedReportsFromDatabase(),
  ]);
  const itemsById = new Map<string, ReportHistoryItem>();

  for (const item of databaseItems) itemsById.set(item.id, item);
  for (const { blob: _blob, ...meta } of storedItems) itemsById.set(meta.id, meta);

  return [...itemsById.values()].sort(
    (a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime(),
  );
}

async function getReportHistory(
  id: string,
  userId: string,
): Promise<StoredReportHistoryItem | null> {
  const db = await openReportHistoryDb();
  const item = await new Promise<StoredReportHistoryItem | undefined>((resolve, reject) => {
    const tx = db.transaction(REPORT_HISTORY_STORE, "readonly");
    const request = tx.objectStore(REPORT_HISTORY_STORE).get(id);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
  db.close();
  if (item?.userId === userId) return item;

  const databaseItem = await getGeneratedReportFromDatabase(id);
  if (!databaseItem) return null;
  const scopedItem = { ...databaseItem, userId };
  await addReportHistory(scopedItem);
  return scopedItem;
}

async function deleteReportHistory(id: string, userId: string): Promise<void> {
  const db = await openReportHistoryDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(REPORT_HISTORY_STORE, "readwrite");
    const store = tx.objectStore(REPORT_HISTORY_STORE);
    const request = store.get(id);
    request.onsuccess = () => {
      const item = request.result as StoredReportHistoryItem | undefined;
      if (item?.userId === userId) store.delete(id);
    };
    request.onerror = () => reject(request.error);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
  await deleteGeneratedReportFromDatabase(id);
}

function formatHistoryDate(value: string): string {
  return new Intl.DateTimeFormat("en-PK", {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(new Date(value));
}

function formatBytes(size: number): string {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

/* ---------- helpers (read employee list once) ---------- */

function extractDateFromPdfName(fileName: string): string | null {
  const iso = fileName.match(/(\d{4}-\d{2}-\d{2})/);
  if (iso) return iso[1];

  const named = fileName.match(
    /\b(\d{1,2})\s*(?:-|_|\s)?\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s*(?:-|_|\s)?\s*(\d{4})\b/i,
  );
  if (!named) return null;
  const months: Record<string, string> = {
    jan: "01",
    feb: "02",
    mar: "03",
    apr: "04",
    may: "05",
    jun: "06",
    jul: "07",
    aug: "08",
    sep: "09",
    oct: "10",
    nov: "11",
    dec: "12",
  };
  const month = months[named[2].slice(0, 3).toLowerCase()];
  return month ? `${named[3]}-${month}-${named[1].padStart(2, "0")}` : null;
}

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

async function buildSheetPreview(blob: Blob, preferredSheetName?: string): Promise<SheetPreview> {
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await blob.arrayBuffer());
  const ws = (preferredSheetName && wb.getWorksheet(preferredSheetName)) || wb.worksheets[0];
  const rowCount = Math.max(ws.rowCount || 0, 1);
  const colCount = Math.max(ws.actualColumnCount || ws.columnCount || 0, 1);
  const rows: PreviewCell[][] = [];

  for (let r = 1; r <= rowCount; r++) {
    const row = ws.getRow(r);
    const previewRow: PreviewCell[] = [];
    for (let c = 1; c <= colCount; c++) {
      const cell = row.getCell(c);
      previewRow.push({
        key: `${r}-${c}`,
        rowNumber: r,
        colNumber: c,
        value: previewCellText(cell.value),
        style: previewCellStyle(cell),
      });
    }
    rows.push(previewRow);
  }

  return { name: ws.name || "Sheet preview", sheetName: ws.name, rows };
}

async function applySheetPreviewEdits(blob: Blob, preview: SheetPreview): Promise<Blob> {
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await blob.arrayBuffer());
  const ws = wb.getWorksheet(preview.sheetName) || wb.worksheets[0];

  for (const row of preview.rows) {
    for (const previewCell of row) {
      const cell = ws.getRow(previewCell.rowNumber).getCell(previewCell.colNumber);
      cell.value = coercePreviewValue(previewCell.value, cell.value);
    }
  }

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function coercePreviewValue(value: string, originalValue: unknown): string | number | null {
  const trimmed = value.trim();
  if (!trimmed) return null;

  if (typeof originalValue === "number") {
    const numeric = Number(trimmed);
    return Number.isFinite(numeric) ? numeric : value;
  }

  if (/^-?\d+(?:\.\d+)?$/.test(trimmed) && !/^0\d+/.test(trimmed)) {
    const numeric = Number(trimmed);
    if (Number.isFinite(numeric)) return numeric;
  }

  return value;
}

function previewCellText(value: unknown): string {
  if (value == null) return "";
  if (value instanceof Date) return value.toLocaleDateString();
  if (typeof value === "object") {
    if ("text" in value) return String((value as { text?: unknown }).text ?? "");
    if ("result" in value) return String((value as { result?: unknown }).result ?? "");
    if ("richText" in value && Array.isArray((value as { richText?: unknown }).richText)) {
      return (value as { richText: Array<{ text?: string }> }).richText
        .map((part) => part.text ?? "")
        .join("");
    }
  }
  return String(value);
}

function previewCellStyle(cell: import("exceljs").Cell): React.CSSProperties {
  const fill = cell.fill;
  const fgColor =
    fill && "fgColor" in fill && fill.fgColor && "argb" in fill.fgColor
      ? argbToCss(fill.fgColor.argb)
      : undefined;

  return {
    backgroundColor: fgColor,
    color:
      cell.font?.color && "argb" in cell.font.color ? argbToCss(cell.font.color.argb) : undefined,
    fontWeight: cell.font?.bold ? 700 : 400,
    textAlign: cell.alignment?.horizontal === "center" ? "center" : "left",
    verticalAlign: cell.alignment?.vertical === "middle" ? "middle" : "top",
    borderTop: borderCss(cell.border?.top),
    borderRight: borderCss(cell.border?.right),
    borderBottom: borderCss(cell.border?.bottom),
    borderLeft: borderCss(cell.border?.left),
    minWidth: 96,
    maxWidth: 240,
  };
}

function borderCss(border?: Partial<import("exceljs").Border>): string {
  if (!border?.style) return "1px solid transparent";
  const color = border.color && "argb" in border.color ? argbToCss(border.color.argb) : "#111827";
  return `1px solid ${color}`;
}

function argbToCss(argb?: string): string | undefined {
  if (!argb) return undefined;
  const hex = argb.length === 8 ? argb.slice(2) : argb;
  return `#${hex}`;
}

/* ---------- presentational ---------- */

function BrandMark({ icon }: { icon: React.ReactNode }) {
  return (
    <div className="relative flex h-12 w-12 shrink-0 items-center justify-center">
      <span className="absolute left-0 top-1 h-7 w-7 rounded-full bg-primary/35" />
      <span className="absolute right-0 top-1.5 h-7 w-7 rounded-full bg-primary/25" />
      <span className="absolute bottom-0 left-2.5 h-7 w-7 rounded-full bg-primary/45" />
      <span className="relative flex h-10 w-10 items-center justify-center rounded-xl bg-primary text-primary-foreground shadow-[var(--shadow-soft)] ring-2 ring-white/70">
        {icon}
      </span>
    </div>
  );
}

function ThemeCustomizer({
  color,
  mode,
  onChange,
  onToggleMode,
}: {
  color: string;
  mode: "light" | "dark";
  onChange: (color: string) => void;
  onToggleMode: () => void;
}) {
  const [open, setOpen] = useState(false);
  const panelRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const closeOnOutsideClick = (event: MouseEvent) => {
      if (!panelRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", closeOnOutsideClick);
    return () => document.removeEventListener("mousedown", closeOnOutsideClick);
  }, [open]);

  return (
    <div ref={panelRef} className="fixed right-8 top-4 z-50 sm:right-14 lg:right-20">
      <div className="flex items-center gap-2">
        <button
          type="button"
          onClick={onToggleMode}
          className="flex items-center gap-2 rounded-xl border border-white/50 bg-card px-2.5 py-2 text-sm font-semibold text-foreground shadow-[var(--shadow-soft)] transition hover:border-primary/40 hover:shadow-[var(--shadow-elegant)] sm:px-3"
          aria-label={`Switch to ${mode === "dark" ? "light" : "dark"} theme`}
        >
          {mode === "dark" ? (
            <Sun className="h-4 w-4 text-primary" />
          ) : (
            <Moon className="h-4 w-4 text-primary" />
          )}
          <span className="hidden sm:inline">{mode === "dark" ? "Light" : "Dark"}</span>
        </button>

        <button
          type="button"
          onClick={() => setOpen((current) => !current)}
          className="flex items-center gap-2 rounded-xl border border-white/50 bg-card px-2.5 py-2 text-sm font-semibold text-foreground shadow-[var(--shadow-soft)] transition hover:border-primary/40 hover:shadow-[var(--shadow-elegant)] sm:px-3"
          aria-expanded={open}
          aria-label="Open color settings"
        >
          <span className="flex h-7 w-7 items-center justify-center">
            <ColorSplashMark />
          </span>
          <span className="hidden sm:inline">Colors</span>
        </button>
      </div>

      {open && (
        <div className="absolute right-0 mt-2 w-52 rounded-2xl border border-white/50 bg-card/98 p-3 shadow-[var(--shadow-elegant)] backdrop-blur">
          <div className="mb-2 text-xs font-bold uppercase tracking-wide text-muted-foreground">
            Theme Color
          </div>
          <div className="grid grid-cols-5 gap-2">
            {THEME_COLORS.map((item) => (
              <button
                key={item}
                type="button"
                className={`h-8 w-8 rounded-md border shadow-sm transition hover:scale-105 ${
                  color.toLowerCase() === item.toLowerCase()
                    ? "border-foreground ring-2 ring-primary/25"
                    : "border-border"
                }`}
                style={{ backgroundColor: item }}
                onClick={() => onChange(item)}
                aria-label={`Use color ${item}`}
                title={item}
              />
            ))}
          </div>
          <label className="mt-3 flex items-center justify-between gap-2 text-xs text-muted-foreground">
            Custom
            <input
              type="color"
              value={color}
              onChange={(event) => onChange(event.target.value)}
              className="h-8 w-14 cursor-pointer rounded border border-border bg-background p-1"
              aria-label="Choose custom theme color"
            />
          </label>
        </div>
      )}
    </div>
  );
}

function ColorSplashMark() {
  return (
    <span className="relative block h-7 w-7 overflow-hidden rounded-md bg-white">
      <span className="absolute left-0 top-0 h-4 w-7 rounded-[999px] bg-[linear-gradient(135deg,#1d7eea_0%,#22a7f0_30%,#b830d8_52%,#ff8a00_76%,#ffd400_100%)]" />
      <span className="absolute left-1 top-2 h-4 w-1.5 rounded-full bg-[#2864e6]" />
      <span className="absolute left-[9px] top-3 h-3 w-1.5 rounded-full bg-[#7a1fd1]" />
      <span className="absolute left-[15px] top-2 h-5 w-1.5 rounded-full bg-[#9b22cc]" />
      <span className="absolute right-1 top-3 h-4 w-1.5 rounded-full bg-[#f05a00]" />
      <span className="absolute left-[14px] top-[15px] h-1.5 w-1.5 rounded-full bg-white ring-1 ring-[#9b22cc]" />
    </span>
  );
}

function ModulePill({ label }: { label: string }) {
  return (
    <div className="mt-1 inline-flex items-center rounded-xl border border-white/25 bg-white/10 px-2.5 py-1 text-xs font-semibold uppercase tracking-wide text-white shadow-[var(--shadow-inset)]">
      {label}
    </div>
  );
}

function FunctionCard({
  title,
  description,
  icon,
  onClick,
  disabled = false,
  className = "",
}: {
  title: string;
  description: string;
  icon: React.ReactNode;
  onClick?: () => void;
  disabled?: boolean;
  className?: string;
}) {
  return (
    <Card
      className={`min-h-52 border-white/45 bg-[var(--gradient-card)] p-5 shadow-[var(--shadow-soft)] transition ${
        disabled ? "opacity-70" : "hover:border-primary/35 hover:shadow-[var(--shadow-elegant)]"
      } ${className}`}
    >
      <div className="flex h-full flex-col gap-5">
        <div className="flex items-start gap-3">
          <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-xl bg-primary text-primary-foreground shadow-[var(--shadow-soft)]">
            {icon}
          </div>
          <div className="min-w-0 flex-1">
            <h3 className="text-lg font-semibold text-foreground">{title}</h3>
            <p className="mt-1 min-h-12 text-sm leading-5 text-muted-foreground">{description}</p>
          </div>
        </div>
        <Button
          className="mt-auto h-10 w-full bg-primary text-primary-foreground shadow-[var(--shadow-soft)] hover:bg-primary/90"
          onClick={onClick}
          disabled={disabled}
        >
          {disabled ? "Coming Soon" : "Open"}
        </Button>
      </div>
    </Card>
  );
}

function PreviewTable({
  preview,
  readOnly = false,
  onCellChange,
}: {
  preview: SheetPreview | null;
  readOnly?: boolean;
  onCellChange?: (rowNumber: number, colNumber: number, value: string) => void;
}) {
  if (!preview) return null;

  return (
    <table className="min-w-max border-collapse text-xs">
      <tbody>
        {preview.rows.map((row, rowIndex) => (
          <tr key={`row-${rowIndex}`}>
            {row.map((cell) => (
              <td key={cell.key} className="min-w-24 p-0 align-middle" style={cell.style}>
                {readOnly ? (
                  <div className="min-h-8 min-w-24 whitespace-nowrap px-2 py-1 text-inherit">
                    {cell.value}
                  </div>
                ) : (
                  <input
                    className="h-full min-h-8 w-full min-w-24 bg-transparent px-2 py-1 text-inherit outline-none focus:bg-primary/10 focus:ring-1 focus:ring-primary"
                    value={cell.value}
                    onChange={(event) =>
                      onCellChange?.(cell.rowNumber, cell.colNumber, event.target.value)
                    }
                    aria-label={`Cell ${cell.rowNumber}, ${cell.colNumber}`}
                  />
                )}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function UploadCard({
  step,
  title,
  subtitle,
  icon,
  ready = false,
  onClick,
  children,
}: {
  step: number;
  title: string;
  subtitle: string;
  icon: React.ReactNode;
  accept: string;
  multiple?: boolean;
  ready?: boolean;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <Card className="border-white/45 bg-[var(--gradient-card)] p-4 shadow-[var(--shadow-soft)] transition-shadow hover:shadow-[var(--shadow-elegant)] sm:p-5">
      <div className="mb-4 flex flex-col items-start justify-between gap-3 sm:flex-row">
        <div className="min-w-0">
          <div className="flex items-center gap-2 text-xs font-bold uppercase text-primary">
            Step {step}
            {ready && <Check className="h-3.5 w-3.5 text-success" />}
          </div>
          <h2 className="mt-1 flex items-center gap-2 text-base font-semibold text-foreground sm:text-lg">
            <span className="flex h-7 w-7 items-center justify-center rounded-xl bg-card text-primary shadow-[var(--shadow-inset)]">
              {icon}
            </span>
            <span className="min-w-0 break-words">{title}</span>
          </h2>
          <p className="text-sm text-muted-foreground">{subtitle}</p>
        </div>
        <Button variant="outline" size="sm" onClick={onClick} className="w-full sm:w-auto">
          <Upload className="mr-1.5 h-3.5 w-3.5" />
          Upload
        </Button>
      </div>
      <div
        role={ready ? undefined : "button"}
        tabIndex={ready ? undefined : 0}
        onClick={ready ? undefined : onClick}
        onKeyDown={(event) => {
          if (ready) return;
          if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            onClick();
          }
        }}
        className={`min-h-24 rounded-2xl border p-3 shadow-[var(--shadow-inset)] transition ${
          ready ? "border-success/35 bg-success/5" : "border-white/40 bg-muted/35"
        } ${ready ? "" : "cursor-pointer hover:border-primary/35 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary"}`}
      >
        {children}
      </div>
    </Card>
  );
}

function EmptyHint({ icon, label }: { icon: React.ReactNode; label: string }) {
  return (
    <div className="flex min-h-16 flex-col items-center justify-center gap-2 text-center text-sm text-muted-foreground">
      <span className="text-primary">{icon}</span>
      {label}
    </div>
  );
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
      className={`flex min-w-0 items-center justify-between gap-2 rounded-xl border px-3 py-2 text-sm shadow-[var(--shadow-inset)] ${
        color === "primary"
          ? "border-primary/25 bg-primary/5 text-foreground"
          : "border-accent/40 bg-accent/30 text-accent-foreground"
      }`}
    >
      <span className="min-w-0 truncate">{name}</span>
      <button
        type="button"
        onClick={(event) => {
          event.stopPropagation();
          onRemove();
        }}
        className="rounded-lg p-0.5 text-muted-foreground transition-colors hover:bg-background hover:text-destructive"
        aria-label={`Remove ${name}`}
      >
        <X className="h-3.5 w-3.5" />
      </button>
    </div>
  );
}

function StatusChip({
  icon,
  label,
  active,
}: {
  icon: React.ReactNode;
  label: string;
  active: boolean;
}) {
  return (
    <div
      className={`inline-flex min-h-9 max-w-full items-center gap-2 rounded-xl border px-3 text-xs font-medium shadow-[var(--shadow-soft)] ${
        active
          ? "border-success/30 bg-success/10 text-foreground"
          : "border-white/40 bg-muted/35 text-muted-foreground"
      }`}
    >
      {icon}
      <span className="truncate">{label}</span>
    </div>
  );
}

function ChecklistItem({ done, label }: { done: boolean; label: string }) {
  return (
    <div className="flex min-w-0 items-center gap-2 text-sm">
      <span
        className={`flex h-5 w-5 items-center justify-center rounded-full border ${
          done
            ? "border-success bg-success text-success-foreground"
            : "border-white/40 bg-muted text-muted-foreground shadow-[var(--shadow-inset)]"
        }`}
      >
        {done && <Check className="h-3 w-3" />}
      </span>
      <span className={done ? "min-w-0 text-foreground" : "min-w-0 text-muted-foreground"}>
        {label}
      </span>
    </div>
  );
}

function CompactMetric({
  icon,
  label,
  value,
}: {
  icon: React.ReactNode;
  label: string;
  value: number | string;
}) {
  return (
    <div className="rounded-2xl border border-white/40 bg-muted/25 p-3 shadow-[var(--shadow-inset)]">
      <div className="flex items-center gap-1.5 text-xs text-muted-foreground">
        <span className="text-primary">{icon}</span>
        {label}
      </div>
      <div className="mt-1 break-words text-xl font-bold text-foreground">{value}</div>
    </div>
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
      className={`rounded-2xl border p-4 shadow-[var(--shadow-inset)] ${
        tone === "warning" ? "border-warning/40 bg-warning/10" : "border-white/40 bg-muted/30"
      }`}
    >
      <div className="text-xs uppercase tracking-wider text-muted-foreground">{label}</div>
      <div className="mt-1 break-words text-2xl font-bold text-foreground">{value}</div>
    </div>
  );
}
