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
  Info,
  Brush,
  MonitorSmartphone,
  Search,
} from "lucide-react";
import type { User } from "@supabase/supabase-js";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Progress } from "@/components/ui/progress";
import { Avatar, AvatarFallback, AvatarImage } from "@/components/ui/avatar";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Slider } from "@/components/ui/slider";
import { toast } from "sonner";
import { Toaster } from "@/components/ui/sonner";
import { supabase } from "@/integrations/supabase/client";
import type { Tables } from "@/integrations/supabase/types";
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
  type MonthlyPlannedDayDetail,
  type MonthlyPlannedResult,
} from "@/lib/monthly-planned-processor";
import {
  BUILT_IN_DISTRIBUTOR_PROFILES,
  DISTRIBUTOR_SALES_COLUMNS,
  DISTRIBUTOR_NUMERIC_COLUMNS,
  exportDistributorSalesExcel,
  processDistributorSalesPdfs,
  type DistributorFormatProfile,
  type DistributorNumericColumnKey,
  type DistributorSalesResult,
  type DistributorSalesRow,
} from "@/lib/distributor-sales-processor";
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
import {
  enforceDeviceLimit,
  getCurrentDeviceId,
  MAX_AUTH_DEVICES,
  revokeCurrentDeviceSession,
} from "@/lib/device-auth";

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
  | "distributor-sales"
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

interface ActiveDeviceSession {
  id: string;
  deviceId: string;
  userAgent: string | null;
  firstSeenAt: string;
  lastSeenAt: string;
  isCurrent: boolean;
}

interface PerformanceReport {
  variant?: "daily" | "monthly-planned";
  title: string;
  description: string;
  fileName: string;
  sourceBlob?: Blob;
  allRows?: PerformanceRow[];
  monthlyPlannedDates?: string[];
  topRows: PerformanceRow[];
  lowRows: PerformanceRow[];
  summaryRows: Record<string, string | number>[];
  teamReports?: PerformanceTeamReport[];
}

interface PerformanceRow {
  teamName?: string;
  employeeCode?: string;
  name: string;
  region?: string;
  city?: string;
  designation?: string;
  value: number | string;
  note?: string;
  details?: Record<string, string | number>;
  monthlyPlannedDailyDetails?: MonthlyPlannedDayDetail[];
  lowReasons?: string[];
}

interface PerformanceTeamReport {
  teamName: string;
  allRows: PerformanceRow[];
  topRows: PerformanceRow[];
  lowRows: PerformanceRow[];
  summaryRows: Record<string, string | number>[];
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
const DISTRIBUTOR_PROFILE_STORAGE_KEY = "distributor-sales-format-profiles";
const DISTRIBUTOR_DELETED_PROFILE_STORAGE_KEY = "distributor-sales-deleted-format-profiles";
const DISTRIBUTOR_FORMAT_DB = "distributor-sales-format-db";
const DISTRIBUTOR_FORMAT_STORE = "distributor_formats";
const DISTRIBUTOR_SAMPLE_BUCKET = "distributor-format-samples";
const DISTRIBUTOR_PROFILE_STORAGE_FOLDER = "profiles";
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

const MANUAL_DISTRIBUTOR_MAPPING_FIELDS = [
  "Distributor Name",
  "Product Code",
  "Product Name",
  "Trade Price",
  "Open Stock",
  "Receipt Qty",
  "Receipt Bns",
  "Total Stock",
  "Sales Qty",
  "Sales Bns",
  "Return Qty",
  "Return Bns",
  "Net Sale Qty",
  "Net Sale Bns",
  "Sale Value",
  "Transfer In",
  "Transfer Out",
  "Closing Stock",
  "Stock Value",
  "Today Sale",
  "Previous Month",
  "Variance Qty",
  "Variance %",
  "Group Name",
  "Group Total",
  "From Date",
  "To Date",
];

const MANUAL_DISTRIBUTOR_HEADER_OPTIONS = [
  "",
  "Code",
  "Product Code",
  "Item Code",
  "Description",
  "Product Name",
  "Item Name",
  "Trade Price",
  "T.P.",
  "Rate",
  "Open Stock",
  "Opening Balance",
  "Receipt Qty",
  "Purchase",
  "Bns",
  "Total Stock",
  "Sales Qty",
  "Return Qty",
  "Net Sale",
  "Sale Value",
  "Closing Stock",
  "Stock Value",
  "Today Sale",
  "Previous Month",
  "Variance",
  "Group",
  "Group Total",
  "From Date",
  "To Date",
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
  const [passwordRecoveryMode, setPasswordRecoveryMode] = useState(() => isPasswordRecoveryUrl());
  const passwordRecoveryModeRef = useRef(passwordRecoveryMode);
  const [signingOut, setSigningOut] = useState(false);
  const [profilePhoto, setProfilePhoto] = useState<string | null>(null);
  const [activeDeviceCount, setActiveDeviceCount] = useState<number | null>(null);
  const [activeDevices, setActiveDevices] = useState<ActiveDeviceSession[]>([]);
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
  const [dailySelfieFiles, setDailySelfieFiles] = useState<File[]>([]);
  const [dailyTemplateFile, setDailyTemplateFile] = useState<File | null>(null);
  const [dailyProcessing, setDailyProcessing] = useState(false);
  const [dailyReport, setDailyReport] = useState<DailyReportResult | null>(null);
  const [bulkDailyReport, setBulkDailyReport] = useState<BulkDailyReportResult | null>(null);
  const [bulkDailyProgress, setBulkDailyProgress] = useState(0);
  const [bulkDailyProgressLabel, setBulkDailyProgressLabel] = useState("");
  const [dailyPreview, setDailyPreview] = useState<SheetPreview | null>(null);
  const [dailyPreviewLoading, setDailyPreviewLoading] = useState(false);
  const [coverageSourceFiles, setCoverageSourceFiles] = useState<File[]>([]);
  const [coverageTemplateFile, setCoverageTemplateFile] = useState<File | null>(null);
  const [coverageProcessing, setCoverageProcessing] = useState(false);
  const [coverageReport, setCoverageReport] = useState<DoctorCoverageResult | null>(null);
  const [coveragePreview, setCoveragePreview] = useState<SheetPreview | null>(null);
  const [coveragePreviewLoading, setCoveragePreviewLoading] = useState(false);
  const [monthlyPlannedCallLogFiles, setMonthlyPlannedCallLogFiles] = useState<File[]>([]);
  const [monthlyPlannedTemplateFile, setMonthlyPlannedTemplateFile] = useState<File | null>(null);
  const [monthlyPlannedProcessing, setMonthlyPlannedProcessing] = useState(false);
  const [monthlyPlannedReport, setMonthlyPlannedReport] = useState<MonthlyPlannedResult | null>(
    null,
  );
  const [monthlyPlannedPreview, setMonthlyPlannedPreview] = useState<SheetPreview | null>(null);
  const [monthlyPlannedPreviewLoading, setMonthlyPlannedPreviewLoading] = useState(false);
  const [distributorFiles, setDistributorFiles] = useState<File[]>([]);
  const [distributorProcessing, setDistributorProcessing] = useState(false);
  const [distributorResult, setDistributorResult] = useState<DistributorSalesResult | null>(null);
  const [distributorExporting, setDistributorExporting] = useState(false);
  const [distributorGeneratedFile, setDistributorGeneratedFile] = useState<{
    blob: Blob;
    url: string;
    fileName: string;
  } | null>(null);
  const [distributorReportPreviewOpen, setDistributorReportPreviewOpen] = useState(false);
  const [distributorSampleMappingOpen, setDistributorSampleMappingOpen] = useState(false);
  const [distributorProfiles, setDistributorProfiles] = useState<DistributorFormatProfile[]>([]);
  const [selectedDistributorFormatId, setSelectedDistributorFormatId] = useState("");
  const [distributorProfileSearch, setDistributorProfileSearch] = useState("");
  const [selectedDistributorProfile, setSelectedDistributorProfile] =
    useState<DistributorFormatProfile | null>(null);
  const [distributorProfileDialogMode, setDistributorProfileDialogMode] = useState<
    "view" | "edit" | "add" | null
  >(null);
  const [distributorFormatPanel, setDistributorFormatPanel] = useState<"add" | "search" | null>(
    null,
  );
  const [distributorLargePanelMode, setDistributorLargePanelMode] = useState<
    "add" | "search" | "view" | "edit" | null
  >(null);
  const [distributorAddSourceType, setDistributorAddSourceType] = useState<
    "PDF" | "Excel" | "Screenshot" | null
  >(null);
  const [distributorSampleFormatName, setDistributorSampleFormatName] = useState("");
  const [distributorSampleFormatFile, setDistributorSampleFormatFile] = useState<File | null>(null);
  const [selectedDistributorSampleFile, setSelectedDistributorSampleFile] = useState<File | null>(
    null,
  );
  const [distributorSampleInputKey, setDistributorSampleInputKey] = useState(0);
  const [distributorManualDraft, setDistributorManualDraft] =
    useState<DistributorFormatProfile | null>(null);
  const [performanceDialogOpen, setPerformanceDialogOpen] = useState(false);
  const [performanceDialogMode, setPerformanceDialogMode] = useState<"options" | "view">("options");

  const excelInputRef = useRef<HTMLInputElement>(null);
  const pdfInputRef = useRef<HTMLInputElement>(null);
  const previewRef = useRef<HTMLDivElement>(null);
  const callLogInputRef = useRef<HTMLInputElement>(null);
  const dailySelfieInputRef = useRef<HTMLInputElement>(null);
  const dailyTemplateInputRef = useRef<HTMLInputElement>(null);
  const dailyPreviewRef = useRef<HTMLDivElement>(null);
  const coverageSourceInputRef = useRef<HTMLInputElement>(null);
  const coverageTemplateInputRef = useRef<HTMLInputElement>(null);
  const coveragePreviewRef = useRef<HTMLDivElement>(null);
  const monthlyPlannedCallLogInputRef = useRef<HTMLInputElement>(null);
  const monthlyPlannedTemplateInputRef = useRef<HTMLInputElement>(null);
  const monthlyPlannedPreviewRef = useRef<HTMLDivElement>(null);
  const distributorInputRef = useRef<HTMLInputElement>(null);

  const savedDistributorProfiles = useMemo(
    () => distributorProfiles.filter((profile) => profile.active !== false),
    [distributorProfiles],
  );
  const visibleDistributorProfiles = useMemo(() => {
    const query = distributorProfileSearch.trim().toLowerCase();
    return savedDistributorProfiles.filter((profile) => {
      const source = profile.sourceSampleType ?? "Manual";
      const haystack = `${profile.distributorName} ${profile.profileName} ${source}`.toLowerCase();
      return query ? haystack.includes(query) : true;
    });
  }, [savedDistributorProfiles, distributorProfileSearch]);
  const selectedDistributorFormat = useMemo(
    () => savedDistributorProfiles.find((profile) => profile.id === selectedDistributorFormatId),
    [savedDistributorProfiles, selectedDistributorFormatId],
  );
  const clearDistributorGeneratedFile = useCallback(() => {
    setDistributorGeneratedFile((current) => {
      if (current) URL.revokeObjectURL(current.url);
      return null;
    });
  }, []);
  const replaceDistributorGeneratedFile = useCallback((blob: Blob) => {
    const nextFile = {
      blob,
      url: URL.createObjectURL(blob),
      fileName: "Distributor_Sales_Report.xlsx",
    };
    setDistributorGeneratedFile((current) => {
      if (current) URL.revokeObjectURL(current.url);
      return nextFile;
    });
  }, []);

  useEffect(() => {
    const key = "dashboard-line-index";
    const current = Number(window.localStorage.getItem(key) ?? "-1");
    const next = (Number.isFinite(current) ? current + 1 : 0) % DASHBOARD_LINES.length;
    window.localStorage.setItem(key, String(next));
    setDashboardLine(DASHBOARD_LINES[next]);
  }, []);

  useEffect(
    () => () => {
      if (distributorGeneratedFile) URL.revokeObjectURL(distributorGeneratedFile.url);
    },
    [distributorGeneratedFile],
  );

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

      const isRecoveryUrl = isPasswordRecoveryUrl();
      if (isRecoveryUrl || passwordRecoveryModeRef.current) {
        passwordRecoveryModeRef.current = true;
        setPasswordRecoveryMode(true);
        setAuthUser(null);
      } else if (data.session) {
        try {
          const deviceResult = await enforceDeviceLimit(data.session);
          if (!mounted) return;
          if (!deviceResult.allowed) {
            setAuthUser(null);
            toast.error(
              deviceResult.message ??
                `This account is already active on ${MAX_AUTH_DEVICES} devices.`,
            );
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
      const isRecoveryUrl = isPasswordRecoveryUrl();
      if (event === "PASSWORD_RECOVERY" || isRecoveryUrl || passwordRecoveryModeRef.current) {
        passwordRecoveryModeRef.current = true;
        setPasswordRecoveryMode(true);
        setAuthUser(null);
        setAuthLoading(false);
        return;
      }

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
            toast.error(
              deviceResult.message ??
                `This account is already active on ${MAX_AUTH_DEVICES} devices.`,
            );
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
    let cancelled = false;
    void getDistributorFormats()
      .then((profiles) => {
        if (cancelled) return;
        setDistributorProfiles(profiles);
      })
      .catch((error) => {
        console.warn("Distributor format profile load failed", error);
        if (!cancelled) {
          toast.error("Saved distributor formats could not be loaded.");
        }
      });

    return () => {
      cancelled = true;
    };
  }, [authUser]);

  useEffect(() => {
    if (
      !distributorFiles.length ||
      !savedDistributorProfiles.length ||
      selectedDistributorFormatId
    ) {
      return;
    }

    const fileText = distributorFiles.map((file) => file.name).join(" ");
    const normalizedFileText = normalizeDistributorFormatName(fileText);
    const exactFileMatch = savedDistributorProfiles.find((profile) => {
      const normalizedName = normalizeDistributorFormatName(profile.distributorName);
      const normalizedProfile = normalizeDistributorFormatName(profile.profileName);
      return (
        normalizedFileText.includes(normalizedName) ||
        normalizedFileText.includes(normalizedProfile)
      );
    });

    if (exactFileMatch?.id) {
      setSelectedDistributorFormatId(exactFileMatch.id);
    }
  }, [distributorFiles, savedDistributorProfiles, selectedDistributorFormatId]);

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
  const performanceReport = useMemo(
    () =>
      buildPerformanceReport({
        activeModule,
        report,
        dailyReport,
        bulkDailyReport,
        coverageReport,
        monthlyPlannedReport,
      }),
    [activeModule, bulkDailyReport, coverageReport, dailyReport, monthlyPlannedReport, report],
  );

  const openPerformanceActions = useCallback(() => {
    if (!performanceReport) {
      toast.info("Generate this report first, then performance details will be available.");
      return;
    }
    setPerformanceDialogMode("options");
    setPerformanceDialogOpen(true);
  }, [performanceReport]);

  const downloadPerformanceReport = useCallback(async () => {
    if (!performanceReport) {
      toast.info("Generate this report first, then performance details will be available.");
      return;
    }

    try {
      const blob = await buildPerformanceWorkbook(performanceReport);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = performanceReport.fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Performance download failed", error);
      toast.error("Unable to download performance Excel file.");
    }
  }, [performanceReport]);

  const refreshHistory = useCallback(
    async (force = false) => {
      if (!authUser) {
        setHistoryItems([]);
        return;
      }

      const cachedItems = force ? null : readReportHistoryCache(authUser.id);
      if (cachedItems) {
        setHistoryItems(cachedItems);
        setHistoryLoading(false);
        return;
      }

      setHistoryLoading(true);
      try {
        const nextItems = await listReportHistory(authUser.id);
        setHistoryItems(nextItems);
        writeReportHistoryCache(authUser.id, nextItems);
      } catch (error) {
        console.error("History load failed", error);
        toast.error("Unable to load report history.");
      } finally {
        setHistoryLoading(false);
      }
    },
    [authUser],
  );

  const openHistory = useCallback(() => {
    setActiveModule("history");
    setHistoryPreview(null);
    void refreshHistory();
  }, [refreshHistory]);

  useEffect(() => {
    if (!authUser) {
      setHistoryItems([]);
      setHistoryPreview(null);
      setActiveDeviceCount(null);
      return;
    }

    const cachedItems = readReportHistoryCache(authUser.id);
    if (cachedItems) {
      setHistoryItems(cachedItems);
      setHistoryLoading(false);
      return;
    }

    void (async () => {
      await syncStoredReportHistoryToDatabase(authUser.id);
      await refreshHistory(true);
    })();
  }, [authUser, refreshHistory]);

  useEffect(() => {
    if (!authUser) return;

    const channel = supabase
      .channel(`generated-reports:${authUser.id}`)
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "generated_reports",
          filter: `user_id=eq.${authUser.id}`,
        },
        () => {
          void refreshHistory(true);
        },
      )
      .subscribe();

    return () => {
      void supabase.removeChannel(channel);
    };
  }, [authUser, refreshHistory]);

  useEffect(() => {
    if (!authUser) {
      setActiveDeviceCount(null);
      setActiveDevices([]);
      return;
    }

    void (async () => {
      const activeSince = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
      const currentDeviceId = getCurrentDeviceId();
      const { data, error } = await supabase
        .from("auth_device_sessions")
        .select("id, device_id, user_agent, first_seen_at, last_seen_at")
        .eq("user_id", authUser.id)
        .is("revoked_at", null)
        .gt("last_seen_at", activeSince)
        .order("last_seen_at", { ascending: false });

      if (error) {
        console.warn("Unable to load active devices", error);
        setActiveDeviceCount(null);
        setActiveDevices([]);
        return;
      }

      const devices = (data ?? []).map((item) => ({
        id: item.id,
        deviceId: item.device_id,
        userAgent: item.user_agent,
        firstSeenAt: item.first_seen_at,
        lastSeenAt: item.last_seen_at,
        isCurrent: Boolean(currentDeviceId && item.device_id === currentDeviceId),
      }));

      setActiveDevices(devices);
      setActiveDeviceCount(devices.length);
    })();
  }, [authUser]);

  const canProcessDaily = callLogFiles.length > 0 && Boolean(dailyTemplateFile) && !dailyProcessing;
  const dailySelfiesAttached = dailySelfieFiles.length > 0;
  const canProcessCoverage =
    coverageSourceFiles.length > 0 && Boolean(coverageTemplateFile) && !coverageProcessing;
  const canProcessMonthlyPlanned =
    monthlyPlannedCallLogFiles.length > 0 &&
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
        await refreshHistory(true);
      } catch (error) {
        console.error("History save failed", error);
        toast.warning("Report ready, but history save failed.");
      }
    },
    [authUser, pdfFiles.length, refreshHistory],
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
        dailySelfieFiles.length ? dailySelfieFiles : undefined,
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
        await refreshHistory(true);
        if (bulkResult.warnings.length) toast.warning(bulkResult.warnings[0]);
        toast.success(
          `Daily report ready. ${bulkResult.reportsGenerated}/${bulkResult.totalTeams} team sheet(s) updated.`,
        );
        return;
      }

      setBulkDailyProgressLabel("Preparing report preview...");
      const result = await processDailyReport(
        callLogFiles,
        dailyTemplateFile,
        dailySelfieFiles.length ? dailySelfieFiles : undefined,
      );
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
      await refreshHistory(true);
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
    if (coverageSourceFiles.length === 0) {
      toast.error("Please upload the Doctor Coverage Excel file(s).");
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
      const result = await processDoctorCoverageReport(coverageSourceFiles, coverageTemplateFile);
      setCoverageReport(result);
      await saveReportHistoryAndDatabase({
        id: crypto.randomUUID(),
        userId: authUser?.id,
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
      await refreshHistory(true);
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
    if (monthlyPlannedCallLogFiles.length === 0) {
      toast.error("Please upload the monthly call log Excel file(s).");
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
        monthlyPlannedCallLogFiles,
        monthlyPlannedTemplateFile,
      );
      setMonthlyPlannedReport(result);
      await saveReportHistoryAndDatabase({
        id: crypto.randomUUID(),
        userId: authUser?.id,
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
      await refreshHistory(true);
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

  const processDistributorSales = async () => {
    if (!distributorFiles.length) {
      toast.error("Please upload distributor sales PDF file(s).");
      return;
    }
    if (savedDistributorProfiles.length > 0 && !selectedDistributorFormatId) {
      toast.error("Please select the correct saved distributor format before generating.");
      return;
    }
    setDistributorProcessing(true);
    setDistributorResult(null);
    clearDistributorGeneratedFile();
    setDistributorReportPreviewOpen(false);
    setDistributorSampleMappingOpen(false);
    try {
      const result = await processDistributorSalesPdfs(distributorFiles, {
        profiles: distributorProfiles,
        selectedProfileId: selectedDistributorFormatId || undefined,
      });
      if (!result.rows.length) {
        setDistributorResult(result);
        if (result.warnings.length) toast.warning(result.warnings[0]);
        toast.error("No distributor rows were generated. Please review the PDF or mapping.");
        return;
      }
      const blob = await exportDistributorSalesExcel(result.rows);
      replaceDistributorGeneratedFile(blob);
      setDistributorResult(result);
      if (result.warnings.length) toast.warning(result.warnings[0]);
      toast.success(
        `Distributor report ready. ${result.rows.filter((row) => row.rowType === "product").length} product row(s) extracted.`,
      );
    } catch (error) {
      console.error("Distributor sales processing failed", error);
      clearDistributorGeneratedFile();
      toast.error(error instanceof Error ? error.message : "Unable to process distributor PDFs.");
    } finally {
      setDistributorProcessing(false);
    }
  };

  const saveDistributorMappings = async () => {
    if (!distributorResult?.summaries.length) {
      toast.error("Generate a distributor preview before saving a mapping.");
      return;
    }

    const now = new Date().toISOString();
    let savedCount = 0;

    for (const summary of distributorResult.summaries) {
      if (!summary.distributorName) continue;
      const numericOrder = sanitizeDistributorNumericOrder(summary.suggestedNumericOrder);
      const existingProfile = distributorProfiles.find(
        (item) =>
          item.distributorName.trim().toLowerCase() ===
          summary.distributorName.trim().toLowerCase(),
      );
      const profile: DistributorFormatProfile = {
        id: existingProfile?.id,
        distributorName: summary.distributorName,
        profileName: `${summary.distributorName} mapping`,
        sourceSampleType: "PDF",
        sourceSampleName: summary.fileName,
        numericOrder,
        sourceHeaders: summary.sourceHeaders,
        sourceColumnPositions: {},
        headerRowRule: "Auto-detected from PDF header aliases and saved for this distributor.",
        productRowRule:
          "Product code is separated first; product name remains complete until numeric values start.",
        groupHeadingRule: "Detect group headings from group labels or uppercase section headings.",
        groupTotalRule: "Read PDF Group Total rows, otherwise calculate from extracted products.",
        dateExtractionRule:
          "Read From Date and To Date from the PDF header when confidently found.",
        distributorNameExtractionRule: "Detected from PDF text or source file name.",
        columnMappingRules: numericOrder.map((key) => distributorNumericLabel(key)).join(" -> "),
        productCodeExtractionRule: "First product/item code at the start of the product row.",
        productNameExtractionRule: "Product name before the first mapped numeric value.",
        multilineProductNameRule: "Merge text continuation lines into previous product name.",
        pageContinuationRule: "Read all pages in order and continue current group across pages.",
        createdAt: existingProfile?.createdAt ?? now,
        lastUpdated: now,
        active: true,
      };

      try {
        await saveDistributorFormat(profile);
        savedCount += 1;
      } catch (error) {
        console.warn("Distributor mapping profile save failed", error);
      }
    }

    setDistributorProfiles(await getDistributorFormats());
    toast.success(
      savedCount
        ? `${savedCount} distributor mapping profile(s) saved.`
        : "No distributor names were available to save.",
    );
  };

  const saveSelectedDistributorProfile = async () => {
    if (!selectedDistributorProfile?.distributorName.trim()) {
      toast.error("Distributor name is required.");
      return;
    }
    const now = new Date().toISOString();
    const profile: DistributorFormatProfile = {
      ...selectedDistributorProfile,
      profileName:
        selectedDistributorProfile.profileName ||
        `${selectedDistributorProfile.distributorName} mapping`,
      numericOrder: sanitizeDistributorNumericOrder(selectedDistributorProfile.numericOrder),
      sourceHeaders: selectedDistributorProfile.sourceHeaders ?? [],
      sourceColumnPositions: selectedDistributorProfile.sourceColumnPositions ?? {},
      createdAt: selectedDistributorProfile.createdAt ?? now,
      lastUpdated: now,
      active: selectedDistributorProfile.active !== false,
    };
    if (await saveDistributorProfile(profile, selectedDistributorSampleFile)) {
      setSelectedDistributorProfile(profile);
      setSelectedDistributorSampleFile(null);
      setDistributorLargePanelMode("view");
    }
  };

  const createDistributorFormatDraft = (
    sourceSampleType: DistributorFormatProfile["sourceSampleType"],
  ): DistributorFormatProfile => {
    const now = new Date().toISOString();
    return {
      distributorName: "",
      profileName: "",
      sourceSampleType,
      sourceSampleName: "",
      numericOrder: DISTRIBUTOR_NUMERIC_COLUMNS.map((column) => column.key),
      sourceHeaders: [],
      sourceColumnPositions: {},
      headerRowRule: "Define the source header row or confirm detected headers.",
      productRowRule: "Product code, full product name, then mapped numeric values.",
      groupHeadingRule: "Detect group heading rows and keep products under the active group.",
      groupTotalRule: "Detect Group Total rows or calculate totals from product rows.",
      dateExtractionRule: "Detect From Date and To Date from report header.",
      distributorNameExtractionRule: "Detect distributor name from report header or file name.",
      columnMappingRules: "Map source columns to the standard distributor sales columns.",
      productCodeExtractionRule: "Extract product code from the start of the product description.",
      productNameExtractionRule: "Keep product name complete and separate from product code.",
      multilineProductNameRule: "Merge continuation lines into previous product name.",
      pageContinuationRule: "Read all pages in order and keep current group across pages.",
      createdAt: now,
      lastUpdated: now,
      active: true,
    };
  };

  const addDistributorFormat = (sourceSampleType: DistributorFormatProfile["sourceSampleType"]) => {
    setSelectedDistributorProfile(createDistributorFormatDraft(sourceSampleType));
    setDistributorLargePanelMode("edit");
  };

  const resetDistributorAddForm = () => {
    setDistributorAddSourceType(null);
    setDistributorSampleFormatFile(null);
    setSelectedDistributorSampleFile(null);
    setDistributorSampleFormatName("");
    setDistributorManualDraft(null);
    setDistributorSampleInputKey((key) => key + 1);
  };

  const closeDistributorLargePanel = () => {
    resetDistributorAddForm();
    setDistributorLargePanelMode(null);
    setSelectedDistributorProfile(null);
  };

  const openDistributorLargePanel = (panel: "add" | "search") => {
    if (distributorLargePanelMode === panel) {
      closeDistributorLargePanel();
      return;
    }
    resetDistributorAddForm();
    setSelectedDistributorProfile(null);
    setDistributorLargePanelMode(panel);
  };

  const toggleDistributorFormatPanel = (panel: "add" | "search") => {
    if (distributorLargePanelMode === panel || distributorFormatPanel === panel) {
      setDistributorFormatPanel(null);
      closeDistributorLargePanel();
      return;
    }
    setDistributorFormatPanel(null);
    openDistributorLargePanel(panel);
  };

  const beginDistributorSampleFormat = (type: "PDF" | "Excel" | "Screenshot") => {
    setDistributorAddSourceType(type);
    setDistributorSampleFormatFile(null);
    setSelectedDistributorSampleFile(null);
    setDistributorSampleFormatName("");
    setDistributorManualDraft(null);
    setDistributorSampleInputKey((key) => key + 1);
  };

  const beginManualDistributorFormat = () => {
    const draft = createDistributorFormatDraft("Manual");
    setDistributorManualDraft({
      ...draft,
      manualColumnMappings: {},
    });
    setDistributorAddSourceType(null);
    setDistributorSampleFormatFile(null);
    setSelectedDistributorSampleFile(null);
    setDistributorSampleFormatName("");
    setDistributorSampleInputKey((key) => key + 1);
  };

  const uploadDistributorSampleFile = async (
    profileName: string,
    file: File,
  ): Promise<
    Pick<
      DistributorFormatProfile,
      "sourceStoragePath" | "sourceMimeType" | "sourceFileSize" | "uploadedAt"
    >
  > => {
    if (!authUser) throw new Error("Please sign in before saving distributor formats.");
    const safeName = safeDistributorProfileName(profileName);
    const path = `${authUser.id}/${safeName || "distributor-format"}/${Date.now()}-${file.name}`;
    const { error } = await supabase.storage.from(DISTRIBUTOR_SAMPLE_BUCKET).upload(path, file, {
      cacheControl: "3600",
      upsert: true,
      contentType: file.type || undefined,
    });
    if (error) throw error;
    return {
      sourceStoragePath: path,
      sourceMimeType: file.type || undefined,
      sourceFileSize: file.size,
      uploadedAt: new Date().toISOString(),
    };
  };

  const saveDistributorProfile = async (
    profile: DistributorFormatProfile,
    sampleFile?: File | null,
  ) => {
    const name = profile.distributorName.trim();
    if (!name) {
      toast.error("Distributor name is required.");
      return false;
    }
    forgetDeletedDistributorProfile(name);
    const existingProfile = distributorProfiles.find(
      (item) => item.distributorName.trim().toLowerCase() === name.toLowerCase(),
    );
    const existingBuiltIn = BUILT_IN_DISTRIBUTOR_PROFILES.find(
      (item) => item.distributorName.trim().toLowerCase() === name.toLowerCase(),
    );
    if (
      (existingProfile || existingBuiltIn) &&
      !window.confirm("A format with this name already exists. Do you want to replace/update it?")
    ) {
      return false;
    }

    const now = new Date().toISOString();
    const profileId = existingProfile?.id ?? profile.id ?? crypto.randomUUID();
    let uploadedSample: Partial<DistributorFormatProfile> = {};
    if (sampleFile) {
      uploadedSample = {
        sourceSampleName: sampleFile.name,
        sourceStoragePath: buildDistributorSampleStorageKey(profileId, name, sampleFile),
        sourceMimeType: sampleFile.type || undefined,
        sourceFileSize: sampleFile.size,
        uploadedAt: now,
      };
    }
    if (sampleFile && authUser) {
      try {
        uploadedSample = {
          ...uploadedSample,
          ...(await uploadDistributorSampleFile(name, sampleFile)),
          sourceSampleName: sampleFile.name,
        };
      } catch (uploadError) {
        console.warn("Distributor sample upload failed; saving metadata locally.", uploadError);
      }
    }

    const nextProfile: DistributorFormatProfile = {
      ...profile,
      ...uploadedSample,
      id: profileId,
      distributorName: name,
      profileName: profile.profileName.trim() || name,
      numericOrder: sanitizeDistributorNumericOrder(profile.numericOrder),
      sourceHeaders: profile.sourceHeaders ?? [],
      sourceColumnPositions: profile.sourceColumnPositions ?? {},
      createdAt: existingProfile?.createdAt ?? profile.createdAt ?? now,
      lastUpdated: now,
      active: true,
      deletedAt: null,
    };

    try {
      await saveDistributorFormat(nextProfile, sampleFile ?? null);
      const nextProfiles = await getDistributorFormats();
      setDistributorProfiles(nextProfiles);
      setDistributorFormatPanel(null);
      setDistributorLargePanelMode("search");
      toast.success("Distributor format saved successfully.", {
        description: nextProfile.sourceSampleName
          ? `${nextProfile.distributorName} - ${nextProfile.sourceSampleName} (${formatBytes(
              nextProfile.sourceFileSize,
            )})`
          : nextProfile.distributorName,
      });

      if (authUser) {
        void saveDistributorProfileToDatabase(authUser.id, nextProfile).catch((error) => {
          console.warn("Distributor format Supabase sync failed after IndexedDB save", error);
        });
      }

      return true;
    } catch (error) {
      console.error("Distributor format IndexedDB save failed", error);
      toast.error("Distributor format could not be saved. Please try again.");
      return false;
    }
  };

  const saveDistributorSampleFormat = async () => {
    if (!distributorAddSourceType) return;
    if (!distributorSampleFormatFile) {
      toast.error(`Please upload a ${distributorAddSourceType.toLowerCase()} sample file.`);
      return;
    }
    const name = distributorSampleFormatName.trim();
    if (!name) {
      toast.error("Please enter distributor name / format name.");
      return;
    }

    const now = new Date().toISOString();
    const profile: DistributorFormatProfile = {
      distributorName: name,
      profileName: name,
      sourceSampleType: distributorAddSourceType,
      sourceSampleName: distributorSampleFormatFile.name,
      numericOrder: DISTRIBUTOR_NUMERIC_COLUMNS.map((column) => column.key),
      sourceHeaders: [`${distributorAddSourceType} sample: ${distributorSampleFormatFile.name}`],
      sourceColumnPositions: {},
      manualColumnMappings: {},
      headerRowRule: `Saved from ${distributorAddSourceType} sample.`,
      productRowRule: `Use ${distributorAddSourceType} sample as distributor format reference.`,
      groupHeadingRule: "Detect group headings from the saved sample format.",
      groupTotalRule: "Detect group total rows from the saved sample format.",
      dateExtractionRule: "Detect From Date and To Date from the saved sample format.",
      distributorNameExtractionRule: "Use saved distributor format name.",
      columnMappingRules: `Saved ${distributorAddSourceType} sample reference.`,
      productCodeExtractionRule: "Use saved sample mapping reference.",
      productNameExtractionRule: "Use saved sample mapping reference.",
      multilineProductNameRule: "Use saved sample mapping reference.",
      pageContinuationRule: "Use saved sample mapping reference.",
      createdAt: now,
      lastUpdated: now,
      active: true,
    };

    if (await saveDistributorProfile(profile, distributorSampleFormatFile))
      resetDistributorAddForm();
  };

  const saveDistributorManualFormat = async () => {
    if (!distributorManualDraft) return;
    const name = distributorManualDraft.distributorName.trim();
    if (!name) {
      toast.error("Distributor name is required.");
      return;
    }

    const manualColumnMappings = distributorManualDraft.manualColumnMappings ?? {};
    if (!Object.values(manualColumnMappings).some((mapping) => mapping.trim())) {
      toast.error("Please add at least one manual mapping value.");
      return;
    }
    const profile: DistributorFormatProfile = {
      ...distributorManualDraft,
      distributorName: name,
      profileName: distributorManualDraft.profileName.trim() || name,
      sourceSampleType: "Manual",
      sourceSampleName: "Manual mapping",
      manualColumnMappings,
      sourceHeaders: Object.values(manualColumnMappings).filter(Boolean),
      columnMappingRules: Object.entries(manualColumnMappings)
        .filter(([, mapping]) => mapping.trim())
        .map(([column, mapping]) => `${column}: ${mapping}`)
        .join("; "),
    };

    if (await saveDistributorProfile(profile)) resetDistributorAddForm();
  };

  const updateSelectedDistributorSample = (file: File | null) => {
    if (!selectedDistributorProfile || !file) return;
    const profileId = selectedDistributorProfile.id ?? crypto.randomUUID();
    setSelectedDistributorSampleFile(file);
    setSelectedDistributorProfile({
      ...selectedDistributorProfile,
      id: profileId,
      sourceSampleName: file.name,
      sourceStoragePath: buildDistributorSampleStorageKey(
        profileId,
        selectedDistributorProfile.distributorName || selectedDistributorProfile.profileName,
        file,
      ),
      sourceMimeType: file.type || undefined,
      sourceFileSize: file.size,
      uploadedAt: new Date().toISOString(),
      sourceHeaders:
        selectedDistributorProfile.sourceSampleType === "PDF"
          ? ["Detected from PDF sample after upload"]
          : selectedDistributorProfile.sourceSampleType === "Excel"
            ? ["Detected from Excel sample after upload"]
            : selectedDistributorProfile.sourceSampleType === "Screenshot"
              ? ["Detected from screenshot/image after OCR review"]
              : selectedDistributorProfile.sourceHeaders,
      columnMappingRules:
        selectedDistributorProfile.sourceSampleType === "Screenshot"
          ? "OCR/table detection requires user verification before saving."
          : selectedDistributorProfile.columnMappingRules,
    });
  };

  const deleteDistributorProfile = async (profile: DistributorFormatProfile) => {
    try {
      if (!window.confirm("Are you sure you want to delete this distributor format?")) return;
      const name = profile.distributorName.trim().toLowerCase();

      await deleteDistributorFormat(profile.id ?? profile.distributorName);
      rememberDeletedDistributorProfile(profile.distributorName);

      if (authUser) {
        void deleteDistributorProfileFromDatabase(profile).catch((error) => {
          console.warn(
            "Distributor format Supabase delete sync failed after IndexedDB delete",
            error,
          );
        });
      }

      const nextProfiles = (await getDistributorFormats()).filter(
        (item) => item.distributorName.trim().toLowerCase() !== name,
      );
      setDistributorProfiles(nextProfiles);

      if (selectedDistributorProfile?.distributorName.trim().toLowerCase() === name) {
        setSelectedDistributorProfile(null);
        setDistributorProfileDialogMode(null);
        setDistributorLargePanelMode("search");
      }

      toast.success("Distributor format deleted successfully.");
    } catch (error) {
      console.error("Distributor format delete failed", error);
      toast.error("Distributor format could not be deleted. Please try again.");
    }
  };

  const updateDistributorCell = (
    rowIndex: number,
    key: keyof DistributorSalesRow,
    value: string,
  ) => {
    setDistributorResult((current) => {
      if (!current) return current;
      const column = DISTRIBUTOR_SALES_COLUMNS.find((item) => item.key === key);
      const rows = current.rows.map((row, index) => {
        if (index !== rowIndex) return row;
        return {
          ...row,
          [key]: column?.numeric && value.trim() !== "" ? Number(value) || 0 : value,
        };
      });
      return { ...current, rows };
    });
  };

  const previewDistributorSalesFile = () => {
    if (!distributorResult?.rows.length) {
      toast.error("Preview data not available. Please generate report again.");
      return;
    }
    setDistributorReportPreviewOpen((open) => !open);
  };

  const downloadDistributorSalesFile = () => {
    if (!distributorGeneratedFile) {
      toast.error("Download failed. Please generate report again.");
      return;
    }
    setDistributorExporting(true);
    try {
      const a = document.createElement("a");
      a.href = distributorGeneratedFile.url;
      a.download = distributorGeneratedFile.fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      toast.success("Distributor sales Excel exported.");
    } catch (error) {
      console.error("Distributor export failed", error);
      toast.error("Download failed. Please generate report again.");
    } finally {
      setDistributorExporting(false);
    }
  };

  const downloadHistoryItem = async (id: string) => {
    if (!authUser) return;
    try {
      const item = await getReportHistory(id, authUser.id);
      if (!item) {
        toast.error("History file not found.");
        await refreshHistory(true);
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
        await refreshHistory(true);
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
      clearReportHistoryCache(authUser.id);
      await refreshHistory(true);
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

  const performanceActions = (
    <PerformanceHeaderActions
      available={Boolean(performanceReport)}
      onOpen={openPerformanceActions}
    />
  );
  const performanceDialog = (
    <PerformanceActionsDialog
      open={performanceDialogOpen}
      mode={performanceDialogMode}
      performanceReport={performanceReport}
      onOpenChange={setPerformanceDialogOpen}
      onModeChange={setPerformanceDialogMode}
      onDownload={downloadPerformanceReport}
    />
  );

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

  if (!authUser || passwordRecoveryMode) {
    return (
      <AuthScreen
        color={themeColor}
        mode={themeMode}
        passwordRecoveryMode={passwordRecoveryMode}
        onRecoveryComplete={() => {
          passwordRecoveryModeRef.current = false;
          setPasswordRecoveryMode(false);
          setAuthUser(null);
        }}
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
          activeDeviceCount={activeDeviceCount}
          activeDevices={activeDevices}
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
          actions={performanceActions}
        />
        {performanceDialog}

        <main className="grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_260px] 2xl:grid-cols-[minmax(0,1fr)_280px]">
          <div className="space-y-6">
            <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-3">
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
                  <MultiFileList
                    files={callLogFiles}
                    onClear={() => {
                      setCallLogFiles([]);
                      if (callLogInputRef.current) callLogInputRef.current.value = "";
                      setDailyReport(null);
                      setBulkDailyReport(null);
                      setDailyPreview(null);
                    }}
                    onRemove={(file) => {
                      setCallLogFiles((current) => current.filter((item) => !sameFile(item, file)));
                      if (callLogInputRef.current) callLogInputRef.current.value = "";
                      setDailyReport(null);
                      setBulkDailyReport(null);
                      setDailyPreview(null);
                    }}
                  />
                ) : (
                  <EmptyHint
                    icon={<FileSpreadsheet className="h-5 w-5" />}
                    label="Choose call log file(s)"
                  />
                )}
              </UploadCard>

              <UploadCard
                step={3}
                title="Selfies Excel (Optional)"
                subtitle={
                  dailySelfiesAttached
                    ? `${dailySelfieFiles.length} selfies file(s) selected`
                    : "Upload only when selfie counts are needed"
                }
                icon={<FileSpreadsheet className="h-5 w-5" />}
                accept=".xlsx,.xls,.xlsm"
                ready={dailySelfiesAttached}
                onClick={() => dailySelfieInputRef.current?.click()}
              >
                <input
                  ref={dailySelfieInputRef}
                  type="file"
                  accept=".xlsx,.xls,.xlsm"
                  multiple
                  className="hidden"
                  onChange={(event) => {
                    const files = Array.from(event.target.files ?? []);
                    if (!files.length) return;
                    setDailySelfieFiles(files);
                    setDailyReport(null);
                    setBulkDailyReport(null);
                    setDailyPreview(null);
                  }}
                />
                {dailySelfieFiles.length ? (
                  <MultiFileList
                    files={dailySelfieFiles}
                    onClear={() => {
                      setDailySelfieFiles([]);
                      if (dailySelfieInputRef.current) dailySelfieInputRef.current.value = "";
                      setDailyReport(null);
                      setBulkDailyReport(null);
                      setDailyPreview(null);
                    }}
                    onRemove={(file) => {
                      setDailySelfieFiles((current) =>
                        current.filter((item) => !sameFile(item, file)),
                      );
                      if (dailySelfieInputRef.current) dailySelfieInputRef.current.value = "";
                      setDailyReport(null);
                      setBulkDailyReport(null);
                      setDailyPreview(null);
                    }}
                  />
                ) : (
                  <EmptyHint
                    icon={<FileSpreadsheet className="h-5 w-5" />}
                    label="Choose selfies file(s)"
                  />
                )}
              </UploadCard>
            </div>

            <Card className="border-border bg-card p-6 shadow-[var(--shadow-soft)]">
              <div className="flex flex-col items-start justify-between gap-4 md:flex-row md:items-center">
                <div>
                  <div className="text-xs font-bold uppercase text-primary">Step 4</div>
                  <h2 className="mt-1 text-lg font-semibold text-foreground">
                    Generate daily report
                  </h2>
                  <p className="text-sm text-muted-foreground">
                    Data is matched by Employee Code only, then Planned, Unplanned, Mor, Eve, Total,
                    Cp, and Selfies are filled in the sample file. Selfies are optional; without
                    selfie workbooks the report still generates with selfie counts left at 0.
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

                {(dailyReport?.warnings.length || bulkDailyReport?.warnings.length) && (
                  <div className="mt-5 space-y-2 rounded-md border border-warning/40 bg-warning/10 p-3 text-sm text-warning-foreground">
                    {(dailyReport?.warnings ?? bulkDailyReport?.warnings ?? []).map(
                      (warning, index) => (
                        <div key={`${warning}-${index}`}>{warning}</div>
                      ),
                    )}
                  </div>
                )}

                {dailyReport ? (
                  <>
                    <div className="mt-6 grid gap-4 md:grid-cols-3">
                      <Stat label="Call log rows" value={dailyReport.debug.callRows} />
                      <Stat label="Face-to-face calls" value={dailyReport.debug.faceToFaceRows} />
                      <Stat label="Contact points" value={dailyReport.debug.contactPointRows} />
                      <Stat label="Selfie files" value={dailyReport.debug.selfieFiles} />
                      <Stat label="Selfie images" value={dailyReport.debug.selfieRows} />
                      <Stat
                        label="Selfie matches"
                        value={dailyReport.debug.selfieMatchedEmployees}
                      />
                      <Stat
                        label="No selfie source"
                        value={dailyReport.debug.selfieMissingSources}
                      />
                      <Stat label="Wrong-date selfies" value={dailyReport.debug.selfieDateMisses} />
                      <Stat
                        label="Empty/unreadable files"
                        value={
                          dailyReport.debug.selfieEmptyFiles +
                          dailyReport.debug.selfieColumnMissFiles +
                          dailyReport.debug.selfieUnreadableFiles
                        }
                      />
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
                      <Stat
                        label="Selfie matches"
                        value={bulkDailyReport!.debug.selfieMatchedEmployees}
                      />
                      <Stat
                        label="No selfie source"
                        value={bulkDailyReport!.debug.selfieMissingSources}
                      />
                      <Stat
                        label="Wrong-date selfies"
                        value={bulkDailyReport!.debug.selfieDateMisses}
                      />
                      <Stat
                        label="Empty/unreadable files"
                        value={
                          bulkDailyReport!.debug.selfieEmptyFiles +
                          bulkDailyReport!.debug.selfieColumnMissFiles +
                          bulkDailyReport!.debug.selfieUnreadableFiles
                        }
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
                <ChecklistItem
                  done
                  label={dailySelfiesAttached ? "Selfies Excel attached" : "Selfies Excel optional"}
                />
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
                Total, and Cp. Selfies Excel can be skipped; add WhatsApp export workbooks only when
                image rows should be counted by employee and call-log date.
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
          actions={performanceActions}
        />
        {performanceDialog}

        <main className="grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_260px] 2xl:grid-cols-[minmax(0,1fr)_280px]">
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
                  coverageSourceFiles.length
                    ? `${coverageSourceFiles.length} coverage file(s) selected`
                    : "Upload all coverage workbook(s)"
                }
                icon={<FileSpreadsheet className="h-5 w-5" />}
                accept=".xlsx,.xls"
                ready={coverageSourceFiles.length > 0}
                onClick={() => coverageSourceInputRef.current?.click()}
              >
                <input
                  ref={coverageSourceInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  multiple
                  className="hidden"
                  onChange={(event) => {
                    const files = Array.from(event.target.files ?? []);
                    if (!files.length) return;
                    setCoverageSourceFiles(files);
                    setCoverageReport(null);
                    setCoveragePreview(null);
                    event.target.value = "";
                  }}
                />
                {coverageSourceFiles.length ? (
                  <MultiFileList
                    files={coverageSourceFiles}
                    onClear={() => {
                      setCoverageSourceFiles([]);
                      if (coverageSourceInputRef.current) coverageSourceInputRef.current.value = "";
                      setCoverageReport(null);
                      setCoveragePreview(null);
                    }}
                    onRemove={(file) => {
                      setCoverageSourceFiles((files) =>
                        files.filter((item) => !sameFile(item, file)),
                      );
                      if (coverageSourceInputRef.current) coverageSourceInputRef.current.value = "";
                      setCoverageReport(null);
                      setCoveragePreview(null);
                    }}
                  />
                ) : (
                  <EmptyHint
                    icon={<FileSpreadsheet className="h-5 w-5" />}
                    label="Choose coverage file(s)"
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
                <ChecklistItem
                  done={coverageSourceFiles.length > 0}
                  label="Coverage file(s) attached"
                />
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
          actions={performanceActions}
        />
        {performanceDialog}

        <main className="grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_260px] 2xl:grid-cols-[minmax(0,1fr)_280px]">
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
                  monthlyPlannedCallLogFiles.length
                    ? `${monthlyPlannedCallLogFiles.length} call log file(s) selected`
                    : "Upload all monthly call logs"
                }
                icon={<FileSpreadsheet className="h-5 w-5" />}
                accept=".xlsx,.xls,.xlsm"
                ready={monthlyPlannedCallLogFiles.length > 0}
                onClick={() => monthlyPlannedCallLogInputRef.current?.click()}
              >
                <input
                  ref={monthlyPlannedCallLogInputRef}
                  type="file"
                  accept=".xlsx,.xls,.xlsm"
                  multiple
                  className="hidden"
                  onChange={(event) => {
                    const files = Array.from(event.target.files ?? []);
                    if (!files.length) return;
                    setMonthlyPlannedCallLogFiles(files);
                    setMonthlyPlannedReport(null);
                    setMonthlyPlannedPreview(null);
                    event.target.value = "";
                  }}
                />
                {monthlyPlannedCallLogFiles.length ? (
                  <MultiFileList
                    files={monthlyPlannedCallLogFiles}
                    onClear={() => {
                      setMonthlyPlannedCallLogFiles([]);
                      if (monthlyPlannedCallLogInputRef.current) {
                        monthlyPlannedCallLogInputRef.current.value = "";
                      }
                      setMonthlyPlannedReport(null);
                      setMonthlyPlannedPreview(null);
                    }}
                    onRemove={(file) => {
                      setMonthlyPlannedCallLogFiles((files) =>
                        files.filter((item) => !sameFile(item, file)),
                      );
                      if (monthlyPlannedCallLogInputRef.current) {
                        monthlyPlannedCallLogInputRef.current.value = "";
                      }
                      setMonthlyPlannedReport(null);
                      setMonthlyPlannedPreview(null);
                    }}
                  />
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
                  done={monthlyPlannedCallLogFiles.length > 0}
                  label="Call log file(s) attached"
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

  if (activeModule === "distributor-sales") {
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
          icon={<FileText className="h-5 w-5" />}
          label="Converter"
          title="Distributor Sales"
          description="Convert distributor sales and stock PDFs into one standard Excel workbook."
        />

        <Dialog
          open={Boolean(distributorProfileDialogMode)}
          onOpenChange={(open) => {
            if (!open) {
              setDistributorProfileDialogMode(null);
              setSelectedDistributorProfile(null);
            }
          }}
        >
          <DialogContent className="max-h-[90vh] overflow-auto sm:max-w-3xl">
            <DialogHeader>
              <DialogTitle>
                {distributorProfileDialogMode === "add"
                  ? "Add Distributor Format"
                  : distributorProfileDialogMode === "edit"
                    ? "Edit Distributor Format"
                    : "View Distributor Format"}
              </DialogTitle>
            </DialogHeader>

            {distributorProfileDialogMode === "add" && (
              <div className="grid gap-3 sm:grid-cols-2">
                {(["PDF", "Excel", "Screenshot", "Manual"] as const).map((type) => (
                  <Button
                    key={type}
                    variant="secondary"
                    className="h-auto justify-start p-4 text-left"
                    onClick={() => addDistributorFormat(type)}
                  >
                    <FileText className="mr-3 h-5 w-5" />
                    <span>
                      <span className="block font-semibold">
                        {type === "PDF"
                          ? "Add format from PDF sample"
                          : type === "Excel"
                            ? "Add format from Excel sample"
                            : type === "Screenshot"
                              ? "Add format from screenshot/image"
                              : "Add format manually"}
                      </span>
                      <span className="mt-1 block text-xs text-muted-foreground">
                        Save mapping rules for future PDFs from this distributor.
                      </span>
                    </span>
                  </Button>
                ))}
              </div>
            )}

            {selectedDistributorProfile && distributorProfileDialogMode !== "add" && (
              <div className="space-y-4">
                {distributorProfileDialogMode === "edit" ? (
                  <>
                    <DistributorFormatExcelLayout
                      profile={selectedDistributorProfile}
                      editable
                      onNumericOrderChange={(columnIndex, key) => {
                        const numericOrder = [...selectedDistributorProfile.numericOrder];
                        numericOrder[columnIndex] = key;
                        setSelectedDistributorProfile({
                          ...selectedDistributorProfile,
                          numericOrder: sanitizeDistributorNumericOrder(numericOrder),
                          columnMappingRules: sanitizeDistributorNumericOrder(numericOrder)
                            .map((item) => distributorNumericLabel(item))
                            .join(" -> "),
                        });
                      }}
                    />
                    {selectedDistributorProfile.sourceSampleType !== "Manual" && (
                      <div className="rounded-md border border-border bg-background/60 p-3">
                        <div className="text-xs font-bold uppercase text-primary">
                          {selectedDistributorProfile.sourceSampleType} sample
                        </div>
                        <p className="mt-1 text-sm text-muted-foreground">
                          {selectedDistributorProfile.sourceSampleType === "PDF"
                            ? "Upload a distributor PDF sample to review detected headers, groups, dates, and mapping."
                            : selectedDistributorProfile.sourceSampleType === "Excel"
                              ? "Upload an Excel sample to review the expected output/layout mapping."
                              : "Upload a screenshot/image. OCR confidence should be verified before saving."}
                        </p>
                        <label className="mt-3 flex min-h-24 cursor-pointer flex-col items-center justify-center rounded-md border border-dashed border-border bg-card p-4 text-center text-sm text-muted-foreground hover:bg-muted/40">
                          <Upload className="mb-2 h-5 w-5 text-primary" />
                          {selectedDistributorProfile.sourceSampleName ||
                            `Choose ${selectedDistributorProfile.sourceSampleType?.toLowerCase()} sample`}
                          <input
                            key={`edit-${selectedDistributorProfile.id ?? selectedDistributorProfile.distributorName}-${distributorSampleInputKey}`}
                            type="file"
                            className="hidden"
                            accept={
                              selectedDistributorProfile.sourceSampleType === "PDF"
                                ? ".pdf"
                                : selectedDistributorProfile.sourceSampleType === "Excel"
                                  ? ".xlsx,.xls"
                                  : "image/*"
                            }
                            onChange={(event) =>
                              updateSelectedDistributorSample(event.target.files?.[0] ?? null)
                            }
                          />
                        </label>
                        <div className="mt-3 rounded-md border border-border bg-card p-3 text-xs text-muted-foreground">
                          Mapping preview:{" "}
                          {selectedDistributorProfile.sourceHeaders.length
                            ? selectedDistributorProfile.sourceHeaders.join(", ")
                            : "Upload a sample to attach the source reference, then confirm mapping below."}
                        </div>
                      </div>
                    )}
                    <DistributorProfileField
                      label="Distributor Name"
                      value={selectedDistributorProfile.distributorName}
                      onChange={(value) =>
                        setSelectedDistributorProfile({
                          ...selectedDistributorProfile,
                          distributorName: value,
                        })
                      }
                    />
                    <DistributorProfileField
                      label="Format/Profile Name"
                      value={selectedDistributorProfile.profileName}
                      onChange={(value) =>
                        setSelectedDistributorProfile({
                          ...selectedDistributorProfile,
                          profileName: value,
                        })
                      }
                    />
                    <DistributorProfileField
                      label="Sample Source Name"
                      value={selectedDistributorProfile.sourceSampleName ?? ""}
                      onChange={(value) =>
                        setSelectedDistributorProfile({
                          ...selectedDistributorProfile,
                          sourceSampleName: value,
                        })
                      }
                    />
                    <DistributorProfileField
                      label="Column Mapping Rules"
                      value={selectedDistributorProfile.columnMappingRules ?? ""}
                      onChange={(value) =>
                        setSelectedDistributorProfile({
                          ...selectedDistributorProfile,
                          columnMappingRules: value,
                        })
                      }
                      multiline
                    />
                    <DistributorProfileField
                      label="Group Detection Rules"
                      value={selectedDistributorProfile.groupHeadingRule}
                      onChange={(value) =>
                        setSelectedDistributorProfile({
                          ...selectedDistributorProfile,
                          groupHeadingRule: value,
                        })
                      }
                      multiline
                    />
                    <DistributorProfileField
                      label="Date Detection Rules"
                      value={selectedDistributorProfile.dateExtractionRule}
                      onChange={(value) =>
                        setSelectedDistributorProfile({
                          ...selectedDistributorProfile,
                          dateExtractionRule: value,
                        })
                      }
                      multiline
                    />
                    <DistributorProfileField
                      label="Product Code Rule"
                      value={selectedDistributorProfile.productCodeExtractionRule ?? ""}
                      onChange={(value) =>
                        setSelectedDistributorProfile({
                          ...selectedDistributorProfile,
                          productCodeExtractionRule: value,
                        })
                      }
                      multiline
                    />
                    <DistributorProfileField
                      label="Product Name Rule"
                      value={selectedDistributorProfile.productNameExtractionRule ?? ""}
                      onChange={(value) =>
                        setSelectedDistributorProfile({
                          ...selectedDistributorProfile,
                          productNameExtractionRule: value,
                        })
                      }
                      multiline
                    />
                    {selectedDistributorProfile.sourceSampleType === "Manual" && (
                      <ManualDistributorMappings
                        mappings={selectedDistributorProfile.manualColumnMappings ?? {}}
                        onChange={(key, value) => {
                          const manualColumnMappings = {
                            ...(selectedDistributorProfile.manualColumnMappings ?? {}),
                            [key]: value,
                          };
                          setSelectedDistributorProfile({
                            ...selectedDistributorProfile,
                            manualColumnMappings,
                            columnMappingRules: Object.entries(manualColumnMappings)
                              .map(([column, mapping]) => `${column}: ${mapping}`)
                              .join("; "),
                          });
                        }}
                      />
                    )}
                    <Button onClick={saveSelectedDistributorProfile} className="w-full">
                      <Save className="mr-2 h-4 w-4" />
                      Save Changes
                    </Button>
                  </>
                ) : (
                  <div className="grid gap-3 text-sm">
                    <DistributorFormatExcelLayout profile={selectedDistributorProfile} />
                    <ProfileDetail
                      label="Distributor Name"
                      value={selectedDistributorProfile.distributorName}
                    />
                    <ProfileDetail
                      label="Format Name"
                      value={selectedDistributorProfile.profileName}
                    />
                    <ProfileDetail
                      label="Source"
                      value={`${selectedDistributorProfile.sourceSampleType ?? "Manual"}${
                        selectedDistributorProfile.sourceSampleName
                          ? ` - ${selectedDistributorProfile.sourceSampleName}`
                          : ""
                      }`}
                    />
                    <ProfileDetail
                      label="Saved Column Mapping"
                      value={
                        selectedDistributorProfile.columnMappingRules ||
                        selectedDistributorProfile.numericOrder
                          .map((key) => distributorNumericLabel(key))
                          .join(" -> ")
                      }
                    />
                    <ProfileDetail
                      label="Group Rules"
                      value={selectedDistributorProfile.groupHeadingRule}
                    />
                    <ProfileDetail
                      label="Date Rules"
                      value={selectedDistributorProfile.dateExtractionRule}
                    />
                    <ProfileDetail
                      label="Product Code Rule"
                      value={selectedDistributorProfile.productCodeExtractionRule}
                    />
                    <ProfileDetail
                      label="Product Name Rule"
                      value={selectedDistributorProfile.productNameExtractionRule}
                    />
                    <ProfileDetail
                      label="Last Updated"
                      value={formatShortDate(selectedDistributorProfile.lastUpdated)}
                    />
                  </div>
                )}
              </div>
            )}
          </DialogContent>
        </Dialog>

        <main className="grid min-w-0 gap-4 xl:grid-cols-[minmax(0,1fr)_280px]">
          <div className="space-y-4">
            {distributorLargePanelMode && (
              <Card className="border-border bg-card p-4 shadow-[var(--shadow-soft)] sm:p-6">
                <div className="mb-4 flex items-start justify-between gap-4">
                  <div>
                    <div className="text-xs font-bold uppercase text-primary">
                      Distributor Format
                    </div>
                    <h2 className="mt-1 text-xl font-semibold text-foreground">
                      {distributorLargePanelMode === "add"
                        ? distributorAddSourceType
                          ? distributorAddSourceType === "PDF"
                            ? "Add from PDF sample"
                            : distributorAddSourceType === "Excel"
                              ? "Add from Excel sample"
                              : "Add from screenshot/image"
                          : distributorManualDraft
                            ? "Add manually"
                            : "Add Distributor Format"
                        : distributorLargePanelMode === "search"
                          ? "Search Saved Distributor Formats"
                          : distributorLargePanelMode === "edit"
                            ? "Edit Distributor Format"
                            : "View Distributor Format"}
                    </h2>
                    <p className="mt-1 text-sm text-muted-foreground">
                      {distributorLargePanelMode === "search"
                        ? `Showing ${visibleDistributorProfiles.length} of ${savedDistributorProfiles.length} saved format(s).`
                        : "Manage distributor format settings in a wider, readable workspace."}
                    </p>
                  </div>
                  <Button variant="secondary" size="sm" onClick={closeDistributorLargePanel}>
                    <X className="mr-2 h-4 w-4" />
                    Close
                  </Button>
                </div>

                {distributorLargePanelMode === "add" &&
                  !distributorAddSourceType &&
                  !distributorManualDraft && (
                    <div className="grid gap-3 sm:grid-cols-2 2xl:grid-cols-4">
                      {(["PDF", "Excel", "Screenshot"] as const).map((type) => (
                        <Button
                          key={type}
                          variant="secondary"
                          className="h-auto min-w-0 items-start justify-start p-4 text-left"
                          onClick={() => beginDistributorSampleFormat(type)}
                        >
                          <FileText className="mr-3 mt-0.5 h-5 w-5 shrink-0" />
                          <span className="min-w-0 flex-1 whitespace-normal">
                            <span className="block whitespace-normal break-words text-sm font-semibold leading-snug">
                              {type === "PDF"
                                ? "Add from PDF sample"
                                : type === "Excel"
                                  ? "Add from Excel sample"
                                  : "Add from screenshot/image"}
                            </span>
                            <span className="mt-1 block whitespace-normal break-words text-xs leading-snug text-muted-foreground">
                              Save a reusable sample reference.
                            </span>
                          </span>
                        </Button>
                      ))}
                      <Button
                        variant="secondary"
                        className="h-auto min-w-0 items-start justify-start p-4 text-left"
                        onClick={beginManualDistributorFormat}
                      >
                        <FileText className="mr-3 mt-0.5 h-5 w-5 shrink-0" />
                        <span className="min-w-0 flex-1 whitespace-normal">
                          <span className="block whitespace-normal break-words text-sm font-semibold leading-snug">
                            Add manually
                          </span>
                          <span className="mt-1 block whitespace-normal break-words text-xs leading-snug text-muted-foreground">
                            Map each column with dropdowns or custom text.
                          </span>
                        </span>
                      </Button>
                    </div>
                  )}

                {distributorLargePanelMode === "add" && distributorAddSourceType && (
                  <div className="grid gap-4 lg:grid-cols-[minmax(0,1fr)_minmax(260px,360px)]">
                    <div className="rounded-md border border-border bg-background/60 p-4">
                      <label className="block text-sm font-semibold text-foreground">
                        Distributor/Format Name
                        <input
                          value={distributorSampleFormatName}
                          onChange={(event) => setDistributorSampleFormatName(event.target.value)}
                          placeholder="Enter distributor name / format name"
                          className="mt-2 h-11 w-full rounded-md border border-border bg-background px-3 text-sm font-normal outline-none focus:ring-2 focus:ring-primary/30"
                        />
                      </label>
                      <label className="mt-4 flex min-h-44 cursor-pointer flex-col items-center justify-center rounded-md border border-dashed border-border bg-card/70 p-5 text-center text-sm text-muted-foreground hover:bg-muted/40">
                        <Upload className="mb-3 h-7 w-7 text-primary" />
                        <span className="text-base font-semibold text-foreground">
                          {distributorSampleFormatFile?.name ||
                            `Choose ${distributorAddSourceType.toLowerCase()} sample`}
                        </span>
                        <span className="mt-1 text-xs text-muted-foreground">
                          {distributorAddSourceType === "PDF"
                            ? "PDF files only"
                            : distributorAddSourceType === "Excel"
                              ? "Excel .xlsx or .xls files only"
                              : "Image files only"}
                        </span>
                        <input
                          key={distributorSampleInputKey}
                          type="file"
                          className="hidden"
                          accept={
                            distributorAddSourceType === "PDF"
                              ? ".pdf"
                              : distributorAddSourceType === "Excel"
                                ? ".xlsx,.xls"
                                : "image/*"
                          }
                          onChange={(event) =>
                            setDistributorSampleFormatFile(event.target.files?.[0] ?? null)
                          }
                        />
                      </label>
                    </div>
                    <div className="rounded-md border border-border bg-background/60 p-4">
                      <div className="text-sm font-semibold text-foreground">
                        Save sample format
                      </div>
                      <p className="mt-2 text-sm text-muted-foreground">
                        The selected sample will be saved with this format name and shown in Saved
                        Distributor Formats.
                      </p>
                      <div className="mt-5 grid gap-2">
                        <Button onClick={saveDistributorSampleFormat}>
                          <Save className="mr-2 h-4 w-4" />
                          Save Format
                        </Button>
                        <Button variant="secondary" onClick={resetDistributorAddForm}>
                          Cancel
                        </Button>
                      </div>
                    </div>
                  </div>
                )}

                {distributorLargePanelMode === "add" && distributorManualDraft && (
                  <div className="space-y-4">
                    <label className="block text-sm font-semibold text-foreground">
                      Distributor/Format Name
                      <input
                        value={distributorManualDraft.distributorName}
                        onChange={(event) =>
                          setDistributorManualDraft({
                            ...distributorManualDraft,
                            distributorName: event.target.value,
                            profileName: distributorManualDraft.profileName || event.target.value,
                          })
                        }
                        placeholder="Enter distributor name / format name"
                        className="mt-2 h-11 w-full rounded-md border border-border bg-background px-3 text-sm font-normal outline-none focus:ring-2 focus:ring-primary/30"
                      />
                    </label>
                    <ManualDistributorMappings
                      mappings={distributorManualDraft.manualColumnMappings ?? {}}
                      onChange={(key, value) => {
                        const manualColumnMappings = {
                          ...(distributorManualDraft.manualColumnMappings ?? {}),
                          [key]: value,
                        };
                        setDistributorManualDraft({
                          ...distributorManualDraft,
                          manualColumnMappings,
                          columnMappingRules: Object.entries(manualColumnMappings)
                            .filter(([, mapping]) => mapping.trim())
                            .map(([column, mapping]) => `${column}: ${mapping}`)
                            .join("; "),
                        });
                      }}
                    />
                    <div className="flex flex-col gap-2 sm:flex-row sm:justify-end">
                      <Button variant="secondary" onClick={resetDistributorAddForm}>
                        Cancel
                      </Button>
                      <Button onClick={saveDistributorManualFormat}>
                        <Save className="mr-2 h-4 w-4" />
                        Save Format
                      </Button>
                    </div>
                  </div>
                )}

                {distributorLargePanelMode === "search" && (
                  <div className="space-y-4">
                    <input
                      value={distributorProfileSearch}
                      onChange={(event) => setDistributorProfileSearch(event.target.value)}
                      placeholder="Search by distributor name or format name..."
                      className="h-11 w-full rounded-md border border-border bg-background px-3 text-sm outline-none focus:ring-2 focus:ring-primary/30"
                    />
                    <div className="grid max-h-[58vh] gap-3 overflow-y-auto pr-1 md:grid-cols-2">
                      {visibleDistributorProfiles.map((profile) => (
                        <div
                          key={`${profile.distributorName}-${profile.profileName}`}
                          className="rounded-md border border-border bg-background/50 p-4"
                        >
                          <div className="min-w-0 truncate text-base font-semibold text-foreground">
                            {profile.distributorName}
                          </div>
                          <div className="mt-1 space-y-0.5 text-xs text-muted-foreground">
                            <div className="truncate">
                              File: {profile.sourceSampleName || "No sample file"}
                            </div>
                            <div className="truncate">
                              Ref: {profile.sourceStoragePath || "No file reference"}
                            </div>
                          </div>
                          <div className="mt-3 flex flex-wrap gap-2">
                            <Button
                              variant="secondary"
                              size="sm"
                              onClick={() => {
                                setSelectedDistributorProfile(profile);
                                setDistributorLargePanelMode("view");
                              }}
                            >
                              View
                            </Button>
                            <Button
                              variant="secondary"
                              size="sm"
                              onClick={() => {
                                setSelectedDistributorProfile(profile);
                                setDistributorLargePanelMode("edit");
                              }}
                            >
                              Edit
                            </Button>
                            <Button
                              variant="secondary"
                              size="sm"
                              onClick={() => deleteDistributorProfile(profile)}
                            >
                              Delete
                            </Button>
                          </div>
                        </div>
                      ))}
                      {!visibleDistributorProfiles.length && (
                        <div className="rounded-md border border-border bg-background/50 p-4 text-sm text-muted-foreground">
                          No saved distributor formats match your search.
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {selectedDistributorProfile && distributorLargePanelMode === "view" && (
                  <div className="grid gap-4">
                    <DistributorFormatExcelLayout profile={selectedDistributorProfile} />
                    <div className="grid gap-3 md:grid-cols-2">
                      <ProfileDetail
                        label="Distributor Name"
                        value={selectedDistributorProfile.distributorName}
                      />
                      <ProfileDetail
                        label="Format Name"
                        value={selectedDistributorProfile.profileName}
                      />
                      <ProfileDetail
                        label="Source"
                        value={`${selectedDistributorProfile.sourceSampleType ?? "Manual"}${
                          selectedDistributorProfile.sourceSampleName
                            ? ` - ${selectedDistributorProfile.sourceSampleName}`
                            : ""
                        }`}
                      />
                      <ProfileDetail
                        label="Original File Name"
                        value={selectedDistributorProfile.sourceSampleName}
                      />
                      <ProfileDetail
                        label="File Size"
                        value={formatBytes(selectedDistributorProfile.sourceFileSize)}
                      />
                      <ProfileDetail
                        label="Saved File Reference"
                        value={selectedDistributorProfile.sourceStoragePath}
                      />
                      <ProfileDetail
                        label="Last Updated"
                        value={formatShortDate(selectedDistributorProfile.lastUpdated)}
                      />
                    </div>
                  </div>
                )}

                {selectedDistributorProfile && distributorLargePanelMode === "edit" && (
                  <div className="space-y-4">
                    <DistributorFormatExcelLayout
                      profile={selectedDistributorProfile}
                      editable
                      onNumericOrderChange={(columnIndex, key) => {
                        const numericOrder = [...selectedDistributorProfile.numericOrder];
                        numericOrder[columnIndex] = key;
                        setSelectedDistributorProfile({
                          ...selectedDistributorProfile,
                          numericOrder: sanitizeDistributorNumericOrder(numericOrder),
                          columnMappingRules: sanitizeDistributorNumericOrder(numericOrder)
                            .map((item) => distributorNumericLabel(item))
                            .join(" -> "),
                        });
                      }}
                    />
                    <div className="grid gap-3 md:grid-cols-2">
                      <DistributorProfileField
                        label="Distributor Name"
                        value={selectedDistributorProfile.distributorName}
                        onChange={(value) =>
                          setSelectedDistributorProfile({
                            ...selectedDistributorProfile,
                            distributorName: value,
                          })
                        }
                      />
                      <DistributorProfileField
                        label="Format/Profile Name"
                        value={selectedDistributorProfile.profileName}
                        onChange={(value) =>
                          setSelectedDistributorProfile({
                            ...selectedDistributorProfile,
                            profileName: value,
                          })
                        }
                      />
                    </div>
                    <DistributorProfileField
                      label="Column Mapping Rules"
                      value={selectedDistributorProfile.columnMappingRules ?? ""}
                      onChange={(value) =>
                        setSelectedDistributorProfile({
                          ...selectedDistributorProfile,
                          columnMappingRules: value,
                        })
                      }
                      multiline
                    />
                    {selectedDistributorProfile.sourceSampleType === "Manual" && (
                      <ManualDistributorMappings
                        mappings={selectedDistributorProfile.manualColumnMappings ?? {}}
                        onChange={(key, value) => {
                          const manualColumnMappings = {
                            ...(selectedDistributorProfile.manualColumnMappings ?? {}),
                            [key]: value,
                          };
                          setSelectedDistributorProfile({
                            ...selectedDistributorProfile,
                            manualColumnMappings,
                            columnMappingRules: Object.entries(manualColumnMappings)
                              .map(([column, mapping]) => `${column}: ${mapping}`)
                              .join("; "),
                          });
                        }}
                      />
                    )}
                    <div className="flex flex-col gap-2 sm:flex-row sm:justify-end">
                      <Button variant="secondary" onClick={closeDistributorLargePanel}>
                        Cancel
                      </Button>
                      <Button onClick={saveSelectedDistributorProfile}>
                        <Save className="mr-2 h-4 w-4" />
                        Save Changes
                      </Button>
                    </div>
                  </div>
                )}
              </Card>
            )}

            <div className="grid gap-3 lg:grid-cols-2">
              <UploadCard
                step={1}
                title="Distributor PDF Files"
                subtitle={
                  distributorFiles.length
                    ? `${distributorFiles.length} distributor PDF file(s) selected`
                    : "Upload one or many distributor sales PDF reports"
                }
                icon={<FileText className="h-5 w-5" />}
                accept=".pdf"
                multiple
                ready={distributorFiles.length > 0}
                onClick={() => distributorInputRef.current?.click()}
                compact
              >
                <input
                  ref={distributorInputRef}
                  type="file"
                  accept=".pdf"
                  multiple
                  className="hidden"
                  onChange={(event) => {
                    const files = Array.from(event.target.files ?? []).filter((file) =>
                      file.name.toLowerCase().endsWith(".pdf"),
                    );
                    setDistributorFiles(files);
                    setSelectedDistributorFormatId("");
                    setDistributorResult(null);
                    clearDistributorGeneratedFile();
                    setDistributorReportPreviewOpen(false);
                    setDistributorSampleMappingOpen(false);
                  }}
                />
                {distributorFiles.length ? (
                  <MultiFileList
                    files={distributorFiles}
                    onClear={() => {
                      setDistributorFiles([]);
                      setSelectedDistributorFormatId("");
                      setDistributorResult(null);
                      clearDistributorGeneratedFile();
                      setDistributorReportPreviewOpen(false);
                      setDistributorSampleMappingOpen(false);
                      if (distributorInputRef.current) distributorInputRef.current.value = "";
                    }}
                    onRemove={(file) => {
                      setDistributorFiles((files) => files.filter((item) => item !== file));
                      setSelectedDistributorFormatId("");
                      setDistributorResult(null);
                      clearDistributorGeneratedFile();
                      setDistributorReportPreviewOpen(false);
                      setDistributorSampleMappingOpen(false);
                      if (distributorInputRef.current) distributorInputRef.current.value = "";
                    }}
                  />
                ) : (
                  <EmptyHint icon={<FileText className="h-7 w-7" />} label="Choose PDF files" />
                )}
              </UploadCard>

              <Card className="flex h-[22rem] flex-col overflow-hidden border-border bg-card p-4 shadow-[var(--shadow-soft)]">
                <div>
                  <div className="text-xs font-bold uppercase text-primary">Step 2</div>
                  <h2 className="mt-1 text-lg font-semibold text-foreground">Distributor Format</h2>
                  <p className="mt-1 text-sm leading-relaxed text-muted-foreground">
                    Add new distributor format or search saved distributor formats.
                  </p>
                  <div className="mt-2 rounded-md border border-border bg-background/60 px-3 py-1.5 text-sm font-semibold text-foreground">
                    Saved Formats: {savedDistributorProfiles.length}
                  </div>
                </div>
                <div className="mt-3 grid gap-2">
                  <Button
                    className="w-full shadow-[var(--shadow-elegant)]"
                    onClick={() => toggleDistributorFormatPanel("add")}
                  >
                    <FileText className="mr-2 h-4 w-4" />
                    Add Distributor Format
                  </Button>
                  <Button
                    variant="secondary"
                    className="w-full"
                    onClick={() => toggleDistributorFormatPanel("search")}
                  >
                    <Search className="mr-2 h-4 w-4" />
                    Search Saved Distributor Formats
                  </Button>
                </div>
              </Card>
            </div>

            <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
              <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
                <div>
                  <div className="text-xs font-bold uppercase text-primary">Step 3</div>
                  <h2 className="mt-1 text-lg font-semibold text-foreground">Generate Report</h2>
                  <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
                    Generate the distributor sales Excel after selecting PDF files.
                  </p>
                  <label className="mt-4 block max-w-md text-sm font-semibold text-foreground">
                    Select Distributor Format
                    <select
                      value={selectedDistributorFormatId}
                      onChange={(event) => {
                        setSelectedDistributorFormatId(event.target.value);
                        setDistributorResult(null);
                        clearDistributorGeneratedFile();
                        setDistributorReportPreviewOpen(false);
                        setDistributorSampleMappingOpen(false);
                      }}
                      className="mt-2 h-11 w-full rounded-md border border-border bg-background px-3 text-sm outline-none focus:ring-2 focus:ring-primary/30"
                    >
                      <option value="">Select saved distributor format</option>
                      {savedDistributorProfiles.map((profile) => (
                        <option
                          key={profile.id ?? profile.distributorName}
                          value={profile.id ?? ""}
                        >
                          {profile.distributorName}
                        </option>
                      ))}
                    </select>
                  </label>
                  {selectedDistributorFormat && (
                    <div className="mt-2 max-w-md rounded-md border border-border bg-background/60 px-3 py-2 text-xs text-muted-foreground">
                      Using{" "}
                      {selectedDistributorFormat.profileName ||
                        selectedDistributorFormat.distributorName}
                      {selectedDistributorFormat.sourceSampleName
                        ? ` - ${selectedDistributorFormat.sourceSampleName}`
                        : ""}
                    </div>
                  )}
                </div>
                <Button
                  size="lg"
                  className="w-full shadow-[var(--shadow-elegant)] sm:w-auto"
                  onClick={processDistributorSales}
                  disabled={!distributorFiles.length || distributorProcessing}
                >
                  {distributorProcessing ? (
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
            </Card>

            {distributorResult && (
              <Card className="border-border bg-card p-4 shadow-[var(--shadow-soft)] sm:p-6">
                <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
                  <div>
                    <div className="text-xs font-bold uppercase text-primary">Step 4</div>
                    <h2 className="mt-1 text-lg font-semibold text-foreground">
                      Generated Report Options
                    </h2>
                    <p className="text-sm text-muted-foreground">
                      {distributorResult.rows.filter((row) => row.rowType === "product").length}{" "}
                      product row(s) and{" "}
                      {distributorResult.summaries.reduce(
                        (sum, summary) => sum + summary.groups.length,
                        0,
                      )}{" "}
                      group(s) detected from {distributorResult.summaries.length} PDF file(s).
                    </p>
                  </div>
                  <div className="flex flex-col gap-2 sm:flex-row">
                    <Button
                      variant="secondary"
                      onClick={() => setDistributorSampleMappingOpen((open) => !open)}
                      disabled={!distributorResult.summaries.length}
                      className="w-full sm:w-auto"
                    >
                      <TableProperties className="mr-2 h-4 w-4" />
                      Sample Mapping
                    </Button>
                    <Button
                      variant="secondary"
                      onClick={previewDistributorSalesFile}
                      disabled={!distributorGeneratedFile || !distributorResult.rows.length}
                      className="w-full sm:w-auto"
                    >
                      <Eye className="mr-2 h-4 w-4" />
                      Preview File
                    </Button>
                    <Button
                      onClick={downloadDistributorSalesFile}
                      disabled={!distributorGeneratedFile || distributorExporting}
                      className="w-full shadow-[var(--shadow-elegant)] sm:w-auto"
                    >
                      {distributorExporting ? (
                        <>
                          <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                          Preparing...
                        </>
                      ) : (
                        <>
                          <Download className="mr-2 h-4 w-4" />
                          Download File
                        </>
                      )}
                    </Button>
                  </div>
                </div>

                {distributorSampleMappingOpen && (
                  <>
                    <div className="mt-5 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                      {distributorResult.summaries.map((summary) => (
                        <div
                          key={summary.fileName}
                          className="rounded-md border border-border bg-background/50 p-3"
                        >
                          <div className="truncate text-sm font-semibold text-foreground">
                            {summary.distributorName || summary.fileName}
                          </div>
                          <div className="mt-1 text-xs text-muted-foreground">
                            {summary.rowsExtracted} product rows - {summary.groups.length} group(s)
                            - {summary.mappedColumns.length} mapped column(s)
                          </div>
                          <div className="mt-2 text-xs text-muted-foreground">
                            {summary.fromDate || "From date unknown"} -{" "}
                            {summary.toDate || "To date unknown"}
                          </div>
                          <div className="mt-3 rounded-md border border-border bg-card/70 p-2 text-xs">
                            <div className="flex items-center justify-between gap-2">
                              <span className="font-semibold text-foreground">
                                {summary.profileName}
                              </span>
                              <span
                                className={
                                  summary.profileStatus === "saved"
                                    ? "text-emerald-600"
                                    : summary.profileStatus === "auto-detected"
                                      ? "text-primary"
                                      : "text-warning-foreground"
                                }
                              >
                                {summary.profileStatus === "saved"
                                  ? "Saved profile"
                                  : summary.profileStatus === "auto-detected"
                                    ? "Auto-detected"
                                    : "Verify mapping"}
                              </span>
                            </div>
                            <div className="mt-2 line-clamp-3 text-muted-foreground">
                              {summary.suggestedNumericOrder
                                .map((key) => distributorNumericLabel(key))
                                .join(" -> ")}
                            </div>
                          </div>
                          {summary.warnings.length > 0 && (
                            <div className="mt-2 text-xs text-warning-foreground">
                              {summary.warnings[0]}
                            </div>
                          )}
                          {summary.groups.length > 0 && (
                            <div className="mt-2 space-y-1 text-xs text-muted-foreground">
                              {summary.groups.slice(0, 4).map((group) => (
                                <div key={group.name} className="truncate">
                                  {group.name}: {group.productCount} products, total{" "}
                                  {group.totalSource}
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      ))}
                    </div>

                    {distributorResult.warnings.length > 0 && (
                      <div className="mt-5 rounded-md border border-warning/40 bg-warning/10 p-4 text-sm text-warning-foreground">
                        <div className="mb-1 flex items-center gap-2 font-semibold">
                          <AlertTriangle className="h-4 w-4" />
                          Review warnings
                        </div>
                        <ul className="list-inside list-disc space-y-1 text-xs">
                          {distributorResult.warnings.slice(0, 8).map((warning) => (
                            <li key={warning}>{warning}</li>
                          ))}
                        </ul>
                      </div>
                    )}
                  </>
                )}

                {distributorReportPreviewOpen && (
                  <div className="mt-5 max-h-[70vh] overflow-auto rounded-md border border-border bg-background">
                    <DistributorSalesPreviewTable
                      rows={distributorResult.rows}
                      onCellChange={updateDistributorCell}
                    />
                  </div>
                )}
              </Card>
            )}
          </div>

          <aside className="space-y-4">
            <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
              <div className="mb-4 flex items-center justify-between">
                <div>
                  <div className="text-xs font-bold uppercase text-primary">Run status</div>
                  <h2 className="mt-1 text-base font-semibold text-foreground">
                    Distributor converter
                  </h2>
                </div>
                <ShieldCheck className="h-5 w-5 text-primary" />
              </div>
              <div className="space-y-3">
                <ChecklistItem done={distributorFiles.length > 0} label="PDF files attached" />
                <ChecklistItem done={Boolean(distributorResult)} label="Preview generated" />
                <ChecklistItem
                  done={Boolean(distributorResult?.rows.length)}
                  label="Product rows extracted"
                />
              </div>
            </Card>

            <Card className="border-border bg-card p-5 shadow-[var(--shadow-soft)]">
              <div className="grid grid-cols-2 gap-3">
                <CompactMetric
                  icon={<FileText className="h-4 w-4" />}
                  label="PDFs"
                  value={distributorFiles.length}
                />
                <CompactMetric
                  icon={<TableProperties className="h-4 w-4" />}
                  label="Rows"
                  value={
                    distributorResult
                      ? distributorResult.rows.filter((row) => row.rowType === "product").length
                      : "-"
                  }
                />
                <CompactMetric
                  icon={<AlertTriangle className="h-4 w-4" />}
                  label="Warnings"
                  value={distributorResult?.warnings.length ?? "-"}
                />
                <CompactMetric
                  icon={<CheckCircle2 className="h-4 w-4" />}
                  label="Mapped"
                  value={
                    distributorResult
                      ? Math.max(
                          ...distributorResult.summaries.map((item) => item.mappedColumns.length),
                          0,
                        )
                      : "-"
                  }
                />
              </div>
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
                onClick={() => refreshHistory(true)}
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
        actions={performanceActions}
      />
      {performanceDialog}

      <main className="grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_260px] 2xl:grid-cols-[minmax(0,1fr)_280px]">
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
                  <div className="report-card-scroll max-h-44 space-y-1.5 overflow-y-auto pr-1">
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
    {
      id: "distributor-sales" as const,
      label: "Distributor Sales",
      icon: <FileText className="h-4 w-4" />,
      description: "PDF to Excel",
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
  actions,
}: {
  icon: React.ReactNode;
  label: string;
  title: string;
  description: string;
  actions?: React.ReactNode;
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
        {actions && <div className="flex shrink-0 justify-end sm:min-w-48">{actions}</div>}
      </div>
    </section>
  );
}

function PerformanceHeaderActions({
  available,
  onOpen,
}: {
  available: boolean;
  onOpen: () => void;
}) {
  return (
    <Button
      type="button"
      variant="outline"
      onClick={onOpen}
      className={`neuro-button-muted w-full sm:w-auto ${available ? "" : "opacity-75"}`}
    >
      <BarChart3 className="mr-2 h-4 w-4" />
      Top / Low
    </Button>
  );
}

function PerformanceActionsDialog({
  open,
  mode,
  performanceReport,
  onOpenChange,
  onModeChange,
  onDownload,
}: {
  open: boolean;
  mode: "options" | "view";
  performanceReport: PerformanceReport | null;
  onOpenChange: (open: boolean) => void;
  onModeChange: (mode: "options" | "view") => void;
  onDownload: () => void;
}) {
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-h-[90vh] max-w-4xl overflow-y-auto">
        <DialogHeader>
          <DialogTitle>{performanceReport?.title ?? "Performance"}</DialogTitle>
        </DialogHeader>

        {!performanceReport ? (
          <div className="rounded-md border border-dashed border-border p-5 text-sm text-muted-foreground">
            Generate this report first, then performance details will be available.
          </div>
        ) : mode === "options" ? (
          <div className="grid gap-3 sm:grid-cols-2">
            <Button type="button" size="lg" onClick={() => onModeChange("view")}>
              <Eye className="mr-2 h-4 w-4" />
              View Top / Low
            </Button>
            <Button type="button" size="lg" variant="outline" onClick={onDownload}>
              <Download className="mr-2 h-4 w-4" />
              Download Excel
            </Button>
          </div>
        ) : (
          <div>
            <div className="mb-4 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
              <p className="text-sm text-muted-foreground">{performanceReport.description}</p>
              <div className="grid gap-2 sm:grid-cols-2">
                <Button type="button" variant="outline" onClick={() => onModeChange("options")}>
                  Back
                </Button>
                <Button type="button" onClick={onDownload}>
                  <Download className="mr-2 h-4 w-4" />
                  Download
                </Button>
              </div>
            </div>
            <PerformanceReportView report={performanceReport} />
          </div>
        )}
      </DialogContent>
    </Dialog>
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
  const reportFlow = buildReportFlow(historyItems);

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
            {reportFlow.map((item) => (
              <div key={item.label} className="neuro-inset p-4">
                <div className="mb-3 flex items-center justify-between text-sm">
                  <span className="font-semibold text-foreground">{item.label}</span>
                  <span className="text-muted-foreground">
                    {item.hasRows ? `${item.value}%` : "-"}
                  </span>
                </div>
                <div className="h-3 overflow-hidden rounded-full bg-background shadow-[var(--shadow-inset-sm)]">
                  <div
                    className="h-full rounded-full bg-primary transition-all duration-700"
                    style={{ width: `${item.value}%` }}
                  />
                </div>
                <div className="mt-3 flex items-center justify-between gap-3 text-xs text-muted-foreground">
                  <span>{item.reportCount ? `${item.reportCount} saved` : "No reports"}</span>
                  <span>{item.reviewRows ? `${item.reviewRows} review` : "Clean"}</span>
                </div>
                <div className="mt-1 truncate text-xs text-muted-foreground">
                  {item.latestAt
                    ? `Latest ${formatHistoryDate(item.latestAt)}`
                    : "Waiting for data"}
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
      {photo && <AvatarImage src={photo} alt={displayName} className="object-contain" />}
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

function describeDevice(userAgent: string | null): string {
  const text = userAgent || "";
  const browser = text.includes("Edg/")
    ? "Microsoft Edge"
    : text.includes("OPR/") || text.includes("Opera")
      ? "Opera"
      : text.includes("Firefox/")
        ? "Firefox"
        : text.includes("Chrome/")
          ? "Chrome"
          : text.includes("Safari/")
            ? "Safari"
            : "Unknown browser";
  const platform = text.includes("Android")
    ? "Android"
    : text.includes("iPhone") || text.includes("iPad")
      ? "iOS"
      : text.includes("Windows")
        ? "Windows"
        : text.includes("Mac OS")
          ? "macOS"
          : text.includes("Linux")
            ? "Linux"
            : "Unknown device";
  return `${browser} on ${platform}`;
}

function formatDeviceDate(value: string): string {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Unknown";
  return date.toLocaleString(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  });
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = src;
  });
}

async function cropProfileImage(
  src: string,
  crop: { zoom: number; x: number; y: number },
): Promise<string> {
  const image = await loadImage(src);
  const previewSize = 240;
  const outputSize = 512;
  const canvas = document.createElement("canvas");
  canvas.width = outputSize;
  canvas.height = outputSize;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Canvas is not available.");

  const containScale = Math.min(
    previewSize / image.naturalWidth,
    previewSize / image.naturalHeight,
  );
  const drawScale = containScale * crop.zoom * (outputSize / previewSize);
  const width = image.naturalWidth * drawScale;
  const height = image.naturalHeight * drawScale;
  const offsetX = crop.x * (outputSize / previewSize);
  const offsetY = crop.y * (outputSize / previewSize);

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, outputSize, outputSize);
  context.drawImage(
    image,
    (outputSize - width) / 2 + offsetX,
    (outputSize - height) / 2 + offsetY,
    width,
    height,
  );

  return canvas.toDataURL("image/jpeg", 0.9);
}

function CropSlider({
  label,
  value,
  min,
  max,
  step,
  onChange,
}: {
  label: string;
  value: number;
  min: number;
  max: number;
  step: number;
  onChange: (value: number) => void;
}) {
  return (
    <label className="block space-y-2">
      <div className="flex items-center justify-between gap-3 text-sm font-semibold text-foreground">
        <span>{label}</span>
        <span className="text-xs text-muted-foreground">{value.toFixed(step < 1 ? 2 : 0)}</span>
      </div>
      <Slider
        value={[value]}
        min={min}
        max={max}
        step={step}
        onValueChange={(next) => onChange(next[0] ?? value)}
      />
    </label>
  );
}

function ProfilePage({
  user,
  profilePhoto,
  color,
  mode,
  activeDeviceCount,
  activeDevices,
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
  activeDeviceCount: number | null;
  activeDevices: ActiveDeviceSession[];
  signingOut: boolean;
  onProfilePhotoChange: (photo: string | null) => void;
  onChange: (color: string) => void;
  onToggleMode: () => void;
  onSignOut: () => void;
}) {
  const photoInputRef = useRef<HTMLInputElement>(null);
  const [cropImage, setCropImage] = useState<string | null>(null);
  const [cropZoom, setCropZoom] = useState(1);
  const [cropX, setCropX] = useState(0);
  const [cropY, setCropY] = useState(0);
  const [cropSaving, setCropSaving] = useState(false);
  const [showEditProfile, setShowEditProfile] = useState(false);
  const [showColorPanel, setShowColorPanel] = useState(false);
  const [showAbout, setShowAbout] = useState(false);
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
      openCropEditor(await readFileAsDataUrl(file));
    } catch (error) {
      console.error("Profile photo read failed", error);
      toast.error("Unable to use this profile photo.");
    }
  };

  const openCropEditor = (image: string) => {
    setCropImage(image);
    setCropZoom(1);
    setCropX(0);
    setCropY(0);
  };

  const saveCroppedProfilePhoto = async () => {
    if (!cropImage) return;
    setCropSaving(true);
    try {
      onProfilePhotoChange(
        await cropProfileImage(cropImage, { zoom: cropZoom, x: cropX, y: cropY }),
      );
      setCropImage(null);
      toast.success("Profile photo updated.");
    } catch (error) {
      console.error("Profile crop failed", error);
      toast.error("Unable to crop this profile photo.");
    } finally {
      setCropSaving(false);
    }
  };

  return (
    <main className="grid min-w-0 gap-5 xl:grid-cols-[minmax(0,1fr)_340px]">
      <div className="space-y-5">
        <section className="neuro-panel p-4 sm:p-6">
          <div className="flex min-w-0 items-center gap-3 sm:gap-4">
            <UserAvatar user={user} photo={profilePhoto} className="h-14 w-14 rounded-2xl" />
            <div className="min-w-0">
              <div className="text-xs font-bold uppercase tracking-wide text-primary">Profile</div>
              <h1 className="mt-1 truncate text-xl font-bold text-foreground sm:text-2xl">
                {displayName}
              </h1>
              <p className="text-sm text-muted-foreground">Account, theme, and login details.</p>
            </div>
          </div>

          <div className="mt-6 grid gap-4 sm:grid-cols-3">
            <div className="neuro-inset p-4">
              <div className="text-xs font-bold uppercase tracking-wide text-primary">Name</div>
              <div className="mt-2 truncate text-sm font-semibold text-foreground">
                {displayName}
              </div>
            </div>
            <div className="neuro-inset p-4">
              <div className="text-xs font-bold uppercase tracking-wide text-primary">Theme</div>
              <div className="mt-2 text-sm font-semibold text-foreground">
                {mode === "dark" ? "Dark mode" : "Light mode"}
              </div>
            </div>
            <div className="neuro-inset p-4">
              <div className="text-xs font-bold uppercase tracking-wide text-primary">Devices</div>
              <div className="mt-2 text-sm font-semibold text-foreground">
                {activeDeviceCount === null
                  ? "Not available"
                  : `${activeDeviceCount}/${MAX_AUTH_DEVICES}`}
              </div>
            </div>
          </div>
        </section>

        <section className="neuro-panel p-5 sm:p-6">
          <div className="mb-5 flex items-center justify-between gap-3">
            <div>
              <div className="text-xs font-bold uppercase tracking-wide text-primary">
                Interface
              </div>
              <h2 className="text-lg font-semibold text-foreground">Theme controls</h2>
            </div>
            <Settings className="h-5 w-5 text-primary" />
          </div>

          <div className="grid gap-3 sm:grid-cols-2">
            <Button type="button" onClick={onToggleMode} className="neuro-button min-h-12 w-full">
              {mode === "dark" ? (
                <Sun className="mr-2 h-4 w-4" />
              ) : (
                <Moon className="mr-2 h-4 w-4" />
              )}
              {mode === "dark" ? "Light theme" : "Dark theme"}
            </Button>
            <Button
              type="button"
              variant="outline"
              onClick={() => setShowColorPanel((value) => !value)}
              className="neuro-button-muted min-h-12 w-full"
            >
              <Brush className="mr-2 h-4 w-4" />
              Change color
            </Button>
          </div>

          {showColorPanel && (
            <div className="mt-5 rounded-2xl border border-border bg-card/45 p-4 shadow-[var(--shadow-inset)]">
              <div className="grid grid-cols-5 gap-2 sm:grid-cols-10">
                {THEME_COLORS.map((item) => (
                  <button
                    key={item}
                    type="button"
                    className={`h-10 rounded-xl transition hover:scale-105 ${
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
                className="mt-4 h-11 w-full cursor-pointer rounded-xl border border-border bg-background p-1 shadow-[var(--shadow-inset)]"
                aria-label="Choose custom theme color"
              />
            </div>
          )}
        </section>

        <section className="neuro-panel p-5 sm:p-6">
          <div className="mb-4 flex items-center justify-between gap-3">
            <div>
              <div className="text-xs font-bold uppercase tracking-wide text-primary">About</div>
              <h2 className="text-lg font-semibold text-foreground">Account information</h2>
            </div>
            <Info className="h-5 w-5 text-primary" />
          </div>

          <Button
            type="button"
            variant="outline"
            onClick={() => setShowAbout((value) => !value)}
            className="neuro-button-muted w-full sm:w-auto"
          >
            <Info className="mr-2 h-4 w-4" />
            About
          </Button>

          {showAbout && (
            <div className="mt-4 grid gap-3 sm:grid-cols-2">
              <div className="neuro-inset p-4">
                <div className="mb-2 flex items-center gap-2 text-xs font-bold uppercase tracking-wide text-primary">
                  <Mail className="h-4 w-4" />
                  Email
                </div>
                <div className="break-all text-sm font-semibold text-foreground">
                  {user.email ?? "No email"}
                </div>
              </div>
              <div className="neuro-inset p-4">
                <div className="mb-2 flex items-center gap-2 text-xs font-bold uppercase tracking-wide text-primary">
                  <MonitorSmartphone className="h-4 w-4" />
                  Active devices
                </div>
                <div className="text-sm font-semibold text-foreground">
                  {activeDeviceCount === null
                    ? "Not available"
                    : `${activeDeviceCount}/${MAX_AUTH_DEVICES} device(s)`}
                </div>
              </div>
              <div className="neuro-inset p-4 sm:col-span-2">
                <div className="mb-3 flex items-center gap-2 text-xs font-bold uppercase tracking-wide text-primary">
                  <MonitorSmartphone className="h-4 w-4" />
                  Device activity
                </div>
                {activeDevices.length ? (
                  <div className="max-h-72 space-y-2 overflow-y-auto pr-1 [scrollbar-width:none] [&::-webkit-scrollbar]:hidden">
                    {activeDevices.map((device) => (
                      <div
                        key={device.id}
                        className="rounded-xl border border-border bg-card/45 p-3 text-sm shadow-[var(--shadow-inset)]"
                      >
                        <div className="flex items-start justify-between gap-3">
                          <div className="min-w-0">
                            <div className="truncate font-semibold text-foreground">
                              {describeDevice(device.userAgent)}
                            </div>
                            <div className="mt-1 truncate text-xs text-muted-foreground">
                              {device.userAgent || "Unknown device"}
                            </div>
                          </div>
                          {device.isCurrent && (
                            <span className="shrink-0 rounded-full bg-primary/10 px-2 py-1 text-xs font-semibold text-primary">
                              This device
                            </span>
                          )}
                        </div>
                        <div className="mt-3 grid gap-2 text-xs text-muted-foreground sm:grid-cols-2">
                          <div>
                            <span className="font-semibold text-foreground">Last used:</span>{" "}
                            {formatDeviceDate(device.lastSeenAt)}
                          </div>
                          <div>
                            <span className="font-semibold text-foreground">First seen:</span>{" "}
                            {formatDeviceDate(device.firstSeenAt)}
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="text-sm text-muted-foreground">
                    Active device details are not available yet.
                  </div>
                )}
              </div>
            </div>
          )}
        </section>
      </div>

      <aside className="space-y-6">
        <section className="neuro-panel p-5">
          <div className="mb-4 flex items-center gap-2 text-sm font-semibold text-foreground">
            <UserCircle className="h-4 w-4 text-primary" />
            Profile
          </div>
          <div className="mb-4 flex items-center gap-4">
            <UserAvatar user={user} photo={profilePhoto} className="h-16 w-16 rounded-2xl" />
            <div className="min-w-0 text-sm text-muted-foreground">{displayName}</div>
          </div>
          <input
            ref={photoInputRef}
            type="file"
            accept="image/*"
            className="hidden"
            onChange={onProfilePhotoSelected}
          />
          <Button
            type="button"
            onClick={() => setShowEditProfile((value) => !value)}
            className="neuro-button w-full"
          >
            <UserCircle className="mr-2 h-4 w-4" />
            Edit profile
          </Button>

          {showEditProfile && (
            <div className="mt-4 grid gap-2">
              <Button
                type="button"
                onClick={() => photoInputRef.current?.click()}
                className="neuro-button w-full"
              >
                <Upload className="mr-2 h-4 w-4" />
                Upload photo
              </Button>
              {profilePhoto && (
                <>
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => openCropEditor(profilePhoto)}
                    className="neuro-button-muted w-full"
                  >
                    <Palette className="mr-2 h-4 w-4" />
                    Edit crop
                  </Button>
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => onProfilePhotoChange(null)}
                    className="neuro-button-muted w-full"
                  >
                    <X className="mr-2 h-4 w-4" />
                    Remove photo
                  </Button>
                </>
              )}
            </div>
          )}
        </section>

        <Dialog open={Boolean(cropImage)} onOpenChange={(open) => !open && setCropImage(null)}>
          <DialogContent className="max-w-md border-border bg-card">
            <DialogHeader>
              <DialogTitle>Edit profile photo</DialogTitle>
            </DialogHeader>
            {cropImage && (
              <div className="space-y-5">
                <div className="relative mx-auto h-60 w-60 overflow-hidden rounded-2xl bg-muted shadow-[var(--shadow-inset)]">
                  <img
                    src={cropImage}
                    alt="Profile crop preview"
                    className="h-full w-full object-contain"
                    style={{
                      transform: `translate(${cropX}px, ${cropY}px) scale(${cropZoom})`,
                    }}
                  />
                  <div className="pointer-events-none absolute inset-4 rounded-full border border-white/80 shadow-[0_0_0_999px_rgba(0,0,0,0.08)]" />
                </div>

                <CropSlider
                  label="Zoom"
                  value={cropZoom}
                  min={1}
                  max={3}
                  step={0.05}
                  onChange={setCropZoom}
                />
                <CropSlider
                  label="Horizontal"
                  value={cropX}
                  min={-100}
                  max={100}
                  step={1}
                  onChange={setCropX}
                />
                <CropSlider
                  label="Vertical"
                  value={cropY}
                  min={-100}
                  max={100}
                  step={1}
                  onChange={setCropY}
                />

                <div className="grid gap-2 sm:grid-cols-4">
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => {
                      setCropZoom(1);
                      setCropX(0);
                      setCropY(0);
                    }}
                  >
                    Fit
                  </Button>
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => {
                      setCropZoom(1.35);
                      setCropX(0);
                      setCropY(0);
                    }}
                  >
                    Fill
                  </Button>
                  <Button type="button" variant="outline" onClick={() => setCropImage(null)}>
                    Cancel
                  </Button>
                  <Button type="button" onClick={saveCroppedProfilePhoto} disabled={cropSaving}>
                    {cropSaving ? (
                      <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                    ) : (
                      <Save className="mr-2 h-4 w-4" />
                    )}
                    Save
                  </Button>
                </div>
              </div>
            )}
          </DialogContent>
        </Dialog>

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

function buildPerformanceReport({
  activeModule,
  report,
  dailyReport,
  bulkDailyReport,
  coverageReport,
  monthlyPlannedReport,
}: {
  activeModule: ActiveModule;
  report: ProcessReport | null;
  dailyReport: DailyReportResult | null;
  bulkDailyReport: BulkDailyReportResult | null;
  coverageReport: DoctorCoverageResult | null;
  monthlyPlannedReport: MonthlyPlannedResult | null;
}): PerformanceReport | null {
  if (activeModule === "monthly-report" && report) return buildMonthlyPerformanceReport(report);
  if (activeModule === "monthly-planned" && monthlyPlannedReport) {
    return buildMonthlyPlannedPerformanceReport(monthlyPlannedReport);
  }
  if (activeModule === "daily-report") {
    if (bulkDailyReport) return buildBulkDailyPerformanceReport(bulkDailyReport);
    if (dailyReport) return buildDailyPerformanceReport(dailyReport);
  }
  if (activeModule === "doctor-coverage" && coverageReport) {
    return buildCoveragePerformanceReport(coverageReport);
  }

  return null;
}

function buildMonthlyPerformanceReport(report: ProcessReport): PerformanceReport {
  const rows = report.performanceRows.map((row) => ({
    teamName: row.teamName,
    employeeCode: row.employeeCode,
    name: row.name,
    designation: row.designation,
    value: row.totalCalls,
    note: `${row.totalCalls} calls, ${row.totalSelfies} selfies`,
    details: {
      team: row.teamName,
      employeeId: row.employeeCode,
      name: row.name,
      designation: row.designation,
      totalCalls: row.totalCalls,
      totalSelfies: row.totalSelfies,
      callsPerDay: row.callsPerDay,
      selfiesPerDay: row.selfiesPerDay,
      plannedCalls: row.planned,
      unplannedCalls: row.unplanned,
      plannedPercent: row.plannedPercent,
      cpTime: row.cpTime || "Not available",
    },
    topQualified: row.topQualified,
    lowScore: row.callsPerDay + row.selfiesPerDay + row.plannedPercent,
  }));
  return {
    title: "Monthly Report Performance",
    description:
      "Top: 10 calls/day, 10 selfies/day, and 70% planned calls. CP time is not available from monthly PDF data.",
    fileName: report.fileName.replace(/\.xlsx$/i, "") + " - Performance.xlsx",
    topRows: rows
      .filter((row) => row.topQualified)
      .sort((a, b) => Number(b.value) - Number(a.value))
      .slice(0, 20),
    lowRows: [
      ...buildReviewRows(report.unmatchedNames),
      ...rows
        .filter((row) => !row.topQualified)
        .sort((a, b) => a.lowScore - b.lowScore)
        .slice(0, 20),
    ],
    summaryRows: [
      {
        totalEmployees: report.totalEmployees,
        matchedEmployees: report.matchedEmployees,
        needsReview: report.unmatchedNames.length,
        daysCovered: report.dates.length,
      },
    ],
  };
}

function buildMonthlyPlannedPerformanceReport(report: MonthlyPlannedResult): PerformanceReport {
  const rows = report.performanceRows.map((row) => ({
    teamName: row.teamName,
    employeeCode: row.employeeCode,
    name: row.name,
    region: row.region,
    designation: row.designation,
    value: row.lateCpDays + row.lowCallDays + row.missingShiftDays + row.lowWorkingHourDays,
    note: row.lowReasons.length
      ? row.lowReasons.join("; ")
      : `${row.plannedPercent}% planned, ${row.totalCalls} calls`,
    details: {
      team: row.teamName,
      employeeCode: row.employeeCode,
      name: row.name,
      region: row.region,
      designation: row.designation,
      plannedCalls: row.planned,
      unplannedCalls: row.unplanned,
      totalCalls: row.totalCalls,
      plannedPercent: row.plannedPercent,
      cpAvgTime: row.cpAvgTime,
      lateCpDays: row.lateCpDays,
      lowCallDays: row.lowCallDays,
      missingShiftDays: row.missingShiftDays,
      lowWorkingHourDays: row.lowWorkingHourDays,
    },
    monthlyPlannedDailyDetails: row.dailyDetails,
    lowReasons: row.lowReasons,
    topQualified: row.topQualified,
    lowQualified: row.lowQualified,
  }));
  return {
    variant: "monthly-planned",
    title: "Monthly Planned Unplanned Performance",
    description:
      "Shows every employee from the sample file with month-wise planned/unplanned performance details.",
    fileName: report.fileName.replace(/\.xlsx$/i, "") + " - Performance.xlsx",
    monthlyPlannedDates: report.dates,
    allRows: rows,
    topRows: rows
      .filter((row) => row.topQualified)
      .sort((a, b) => Number(b.details.plannedPercent) - Number(a.details.plannedPercent)),
    lowRows: [
      ...buildReviewRows(report.unmatchedEmployees),
      ...rows.filter((row) => row.lowQualified).sort((a, b) => Number(b.value) - Number(a.value)),
    ],
    summaryRows: [
      {
        totalEmployees: report.totalEmployees,
        matchedEmployees: report.matchedEmployees,
        needsReview: report.unmatchedEmployees.length,
      },
    ],
  };
}

function buildDailyPerformanceReport(report: DailyReportResult): PerformanceReport {
  const unmatched = Math.max(report.totalEmployees - report.matchedEmployees, 0);
  const rows = dailyPerformanceRows(report.performanceRows);
  return {
    variant: "daily",
    title: "Daily Report Performance",
    description:
      "Top: 12 total calls with 4+ morning hours and 3+ evening hours. Low: 5 or fewer total calls, or missing morning/evening shift.",
    fileName: report.fileName.replace(/\.xlsx$/i, "") + " - Performance.xlsx",
    sourceBlob: report.blob,
    allRows: rows,
    topRows: rows
      .filter((row) => row.topQualified)
      .sort((a, b) => Number(b.value) - Number(a.value)),
    lowRows: [
      ...buildReviewRows(report.unmatchedEmployees),
      ...rows.filter((row) => row.lowQualified).sort((a, b) => Number(a.value) - Number(b.value)),
    ],
    summaryRows: [
      {
        totalEmployees: report.totalEmployees,
        matchedEmployees: report.matchedEmployees,
        reviewEmployees: unmatched,
        callRows: report.debug.callRows,
        faceToFaceRows: report.debug.faceToFaceRows,
        contactPointRows: report.debug.contactPointRows,
        selfieFiles: report.debug.selfieFiles,
        selfieImages: report.debug.selfieRows,
      },
    ],
  };
}

function buildBulkDailyPerformanceReport(report: BulkDailyReportResult): PerformanceReport {
  const rows = dailyPerformanceRows(report.performanceRows);
  const failedTeamRows: PerformanceRow[] = report.summary
    .filter((item) => item.status === "failed")
    .map((item) => ({
      teamName: item.teamName,
      name: item.teamName,
      value: "Failed",
      note: item.error ?? "Team sheet could not be processed.",
      details: {
        team: item.teamName,
        status: item.status,
        totalEmployees: item.totalEmployees,
        matchedEmployees: item.matchedEmployees,
        error: item.error ?? "Team sheet could not be processed.",
      },
    }));
  const teamNames = [
    ...new Set([
      ...rows
        .map((row) => row.teamName)
        .filter((teamName): teamName is string => Boolean(teamName)),
      ...report.summary.map((item) => item.teamName).filter(Boolean),
    ]),
  ].sort((a, b) => a.localeCompare(b));
  const teamReports = teamNames.map((teamName) => {
    const teamRows = rows.filter((row) => row.teamName === teamName);
    const teamFailedRows = failedTeamRows.filter((row) => row.teamName === teamName);
    const teamSummary = report.summary.find((item) => item.teamName === teamName);
    return {
      teamName,
      allRows: teamRows,
      topRows: teamRows
        .filter((row) => row.topQualified)
        .sort((a, b) => Number(b.value) - Number(a.value)),
      lowRows: teamRows
        .filter((row) => row.lowQualified)
        .sort((a, b) => Number(a.value) - Number(b.value))
        .concat(teamFailedRows),
      summaryRows: teamSummary
        ? [
            {
              team: teamSummary.teamName,
              status: teamSummary.status,
              totalEmployees: teamSummary.totalEmployees,
              matchedEmployees: teamSummary.matchedEmployees,
              error: teamSummary.error ?? "",
            },
          ]
        : [],
    };
  });
  return {
    variant: "daily",
    title: "Daily Report Performance",
    description:
      "Top: 12 total calls with 4+ morning hours and 3+ evening hours. Low: 5 or fewer total calls, or missing morning/evening shift.",
    fileName: report.fileName.replace(/\.xlsx$/i, "") + " - Performance.xlsx",
    sourceBlob: report.blob,
    allRows: rows,
    topRows: rows
      .filter((row) => row.topQualified)
      .sort((a, b) => Number(b.value) - Number(a.value)),
    lowRows: rows
      .filter((row) => row.lowQualified)
      .sort((a, b) => Number(a.value) - Number(b.value))
      .concat(failedTeamRows),
    summaryRows: [
      {
        totalTeams: report.totalTeams,
        reportsGenerated: report.reportsGenerated,
        failedReports: report.failedReports,
      },
      ...report.summary.map((item) => ({
        team: item.teamName,
        status: item.status,
        totalEmployees: item.totalEmployees,
        matchedEmployees: item.matchedEmployees,
        error: item.error ?? "",
      })),
    ],
    teamReports,
  };
}

function buildCoveragePerformanceReport(report: DoctorCoverageResult): PerformanceReport {
  const unmatched = Math.max(report.totalEmployees - report.matchedEmployees, 0);
  const rows = report.performanceRows.map((row) => ({
    teamName: row.teamName,
    employeeCode: row.employeeCode,
    name: row.name,
    designation: row.designation,
    value: row.coveragePercent,
    note: `${row.coveredDoctors}/${row.targetDoctors} doctors`,
    details: {
      team: row.teamName,
      employeeId: row.employeeCode,
      name: row.name,
      designation: row.designation,
      targetDoctors: row.targetDoctors,
      coveredDoctors: row.coveredDoctors,
      coveragePercent: row.coveragePercent,
    },
    topQualified: row.topQualified,
  }));
  return {
    title: "Doctor Coverage Performance",
    description: "Top: employees with 75% or higher doctor coverage.",
    fileName: report.fileName.replace(/\.xlsx$/i, "") + " - Performance.xlsx",
    topRows: rows
      .filter((row) => row.topQualified)
      .sort((a, b) => Number(b.value) - Number(a.value))
      .slice(0, 20),
    lowRows: [
      ...buildReviewRows(report.unmatchedEmployees),
      ...rows
        .filter((row) => !row.topQualified)
        .sort((a, b) => Number(a.value) - Number(b.value))
        .slice(0, 20),
    ],
    summaryRows: [
      {
        totalEmployees: report.totalEmployees,
        matchedEmployees: report.matchedEmployees,
        reviewEmployees: unmatched,
        coverageRows: report.debug.sourceRows,
        sampleRows: report.debug.templateRows,
      },
    ],
  };
}

function dailyPerformanceRows(
  rows: DailyReportResult["performanceRows"],
): Array<PerformanceRow & { topQualified: boolean; lowQualified: boolean }> {
  return rows.map((row) => ({
    teamName: row.teamName,
    employeeCode: row.employeeCode,
    name: row.name,
    region: row.region,
    city: row.city,
    designation: row.designation,
    value: row.totalCalls,
    note: `${row.totalCalls} calls, ${formatPerformanceHours(row.morningHours)} morning, ${formatPerformanceHours(row.eveningHours)} evening`,
    details: {
      performanceStatus: row.topQualified ? "Top" : row.lowQualified ? "Low" : "Normal",
      employeeCode: row.employeeCode,
      name: row.name,
      region: row.region,
      city: row.city,
      designation: row.designation,
      cpTime: row.cpTime,
      morningCalls: row.morningCalls,
      morningHours: row.morningHours,
      morningLastCall: row.morningLastCall,
      eveningCalls: row.eveningCalls,
      eveningHours: row.eveningHours,
      eveningFirstCall: row.eveningFirstCall,
      eveningLastCall: row.eveningLastCall,
      totalWorkingHours: row.totalWorkingHours,
      totalCalls: row.totalCalls,
      plannedCalls: row.planned,
      unplannedCalls: row.unplanned,
      team: row.teamName,
      selfies: row.selfies,
      plannedPercent: row.plannedPercent,
    },
    topQualified: row.topQualified,
    lowQualified: row.lowQualified,
  }));
}

function buildReviewRows(reviewNames: string[]): PerformanceRow[] {
  return reviewNames.map((name) => ({
    name,
    value: "Review",
    note: "Needs review",
  }));
}

function PerformanceReportView({ report }: { report: PerformanceReport }) {
  return (
    <section className="mt-6">
      <div className="mb-3">
        <div className="text-xs font-bold uppercase tracking-wide text-primary">Performance</div>
        <h3 className="text-lg font-semibold text-foreground">{report.title}</h3>
        <p className="text-sm text-muted-foreground">{report.description}</p>
      </div>

      <div className="grid gap-4 lg:grid-cols-2">
        <PerformancePanel
          title="Top performance"
          rows={report.topRows}
          empty="No top data found."
        />
        <PerformancePanel
          title="Low / needs review"
          rows={report.lowRows}
          empty="No low or review rows found."
        />
      </div>
    </section>
  );
}

function PerformancePanel({
  title,
  rows,
  empty,
}: {
  title: string;
  rows: PerformanceRow[];
  empty: string;
}) {
  return (
    <div className="rounded-md border border-border bg-muted/20 p-4">
      <div className="mb-3 text-sm font-semibold text-foreground">{title}</div>
      {rows.length ? (
        <div className="space-y-2">
          {rows.slice(0, 12).map((row, index) => (
            <div
              key={`${row.name}-${index}`}
              className="flex items-center justify-between gap-3 rounded-md border border-border bg-background/60 px-3 py-2 text-sm"
            >
              <div className="min-w-0">
                <div className="truncate font-semibold text-foreground">{row.name}</div>
                {(row.employeeCode || row.city || row.designation || row.teamName) && (
                  <div className="truncate text-xs text-muted-foreground">
                    {[row.employeeCode, row.city, row.designation, row.teamName]
                      .filter(Boolean)
                      .join(" - ")}
                  </div>
                )}
                {row.note && (
                  <div className="truncate text-xs text-muted-foreground">{row.note}</div>
                )}
              </div>
              <div className="shrink-0 font-bold text-primary">{row.value}</div>
            </div>
          ))}
        </div>
      ) : (
        <div className="flex min-h-36 items-center justify-center rounded-md border border-dashed border-border text-sm text-muted-foreground">
          {empty}
        </div>
      )}
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

const REPORT_FLOW_GROUPS = [
  {
    label: "Daily Report",
    aliases: ["daily report", "daily reports"],
  },
  {
    label: "Monthly Report",
    aliases: ["monthly report"],
  },
  {
    label: "Doctor Coverage",
    aliases: ["doctor coverage", "doctor coverage report"],
  },
  {
    label: "Planned Summary",
    aliases: ["monthly planned unplanned", "monthly planned", "planned summary"],
  },
];

function buildReportFlow(items: ReportHistoryItem[]) {
  return REPORT_FLOW_GROUPS.map((group) => {
    const groupItems = items.filter((item) => reportTypeMatches(item.reportType, group.aliases));
    const totalEmployees = groupItems.reduce((sum, item) => sum + item.totalEmployees, 0);
    const matchedEmployees = groupItems.reduce((sum, item) => sum + item.matchedEmployees, 0);
    const reviewRows = groupItems.reduce(
      (sum, item) => sum + Math.max(item.totalEmployees - item.matchedEmployees, 0),
      0,
    );
    const latestAt = groupItems.reduce<string | null>((latest, item) => {
      if (!latest) return item.createdAt;
      return new Date(item.createdAt).getTime() > new Date(latest).getTime()
        ? item.createdAt
        : latest;
    }, null);
    const hasRows = totalEmployees > 0;

    return {
      label: group.label,
      value: hasRows ? Math.round((matchedEmployees / totalEmployees) * 100) : 0,
      reportCount: groupItems.length,
      totalEmployees,
      matchedEmployees,
      reviewRows,
      latestAt,
      hasRows,
    };
  });
}

function reportTypeMatches(reportType: string | undefined, aliases: string[]) {
  const normalized = normalizeReportType(reportType);
  return aliases.some((alias) => normalized === normalizeReportType(alias));
}

function normalizeReportType(value: string | undefined): string {
  return String(value ?? "Report")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
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

function isPasswordRecoveryUrl(): boolean {
  if (typeof window === "undefined") return false;
  const url = new URL(window.location.href);
  return url.searchParams.get("type") === "recovery" || url.hash.includes("type=recovery");
}

function cleanPasswordRecoveryUrl() {
  if (typeof window === "undefined") return;
  window.history.replaceState(
    {},
    document.title,
    window.location.origin + window.location.pathname,
  );
}

function AuthScreen({
  color: _color,
  mode: _mode,
  passwordRecoveryMode,
  onRecoveryComplete,
  onChange: _onChange,
  onToggleMode: _onToggleMode,
}: {
  color: string;
  mode: "light" | "dark";
  passwordRecoveryMode: boolean;
  onRecoveryComplete: () => void;
  onChange: (color: string) => void;
  onToggleMode: () => void;
}) {
  const [authMode, setAuthMode] = useState<"login" | "signup" | "forgot">("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [submitting, setSubmitting] = useState(false);

  const isSignup = authMode === "signup";
  const isForgot = authMode === "forgot";
  const isReset = passwordRecoveryMode;

  const submitAuth = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setSubmitting(true);

    try {
      const normalizedEmail = email.trim();
      if (isReset) {
        if (!password || password.length < 6) {
          throw new Error("Enter a new password with at least 6 characters.");
        }
        if (password !== confirmPassword) throw new Error("Passwords do not match.");

        const { error } = await supabase.auth.updateUser({ password });
        if (error) throw error;

        toast.success("Password updated. Please login with your new password.");
        setPassword("");
        setConfirmPassword("");
        onRecoveryComplete();
        cleanPasswordRecoveryUrl();
        await supabase.auth.signOut();
        setAuthMode("login");
        return;
      }

      if (isForgot) {
        if (!normalizedEmail) throw new Error("Enter your email address.");
        const { error } = await supabase.auth.resetPasswordForEmail(normalizedEmail, {
          redirectTo: window.location.origin,
        });
        if (error) throw error;

        toast.success("Password reset email sent. Check your inbox.");
        setAuthMode("login");
        return;
      }

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
            toast.error(
              deviceResult.message ??
                `This account is already active on ${MAX_AUTH_DEVICES} devices.`,
            );
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
              <p className="text-sm text-muted-foreground">
                {isReset
                  ? "Set a new password"
                  : isForgot
                    ? "Recover your account"
                    : "Sign in to continue"}
              </p>
            </div>
          </div>

          {!isReset && (
            <div className="neuro-inset mb-5 grid grid-cols-2 p-1">
              <button
                type="button"
                onClick={() => setAuthMode("login")}
                className={`h-9 rounded-md text-sm font-semibold transition ${
                  !isSignup && !isForgot
                    ? "bg-card text-foreground shadow-sm"
                    : "text-muted-foreground"
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
          )}

          <form className="space-y-4" onSubmit={submitAuth}>
            {!isReset && (
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
            )}

            {!isForgot && (
              <label className="block text-sm font-medium text-foreground">
                {isReset ? "New password" : "Password"}
                <span className="neuro-inset mt-1 flex items-center gap-2 px-3 py-2">
                  <Lock className="h-4 w-4 shrink-0 text-primary" />
                  <input
                    value={password}
                    onChange={(event) => setPassword(event.target.value)}
                    type="password"
                    autoComplete={isSignup || isReset ? "new-password" : "current-password"}
                    minLength={6}
                    className="h-8 min-w-0 flex-1 bg-transparent text-sm outline-none"
                    placeholder="At least 6 characters"
                    required
                  />
                </span>
              </label>
            )}

            {isReset && (
              <label className="block text-sm font-medium text-foreground">
                Confirm password
                <span className="neuro-inset mt-1 flex items-center gap-2 px-3 py-2">
                  <Lock className="h-4 w-4 shrink-0 text-primary" />
                  <input
                    value={confirmPassword}
                    onChange={(event) => setConfirmPassword(event.target.value)}
                    type="password"
                    autoComplete="new-password"
                    minLength={6}
                    className="h-8 min-w-0 flex-1 bg-transparent text-sm outline-none"
                    placeholder="Repeat new password"
                    required
                  />
                </span>
              </label>
            )}

            <Button type="submit" className="neuro-button h-11 w-full" disabled={submitting}>
              {submitting ? (
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
              ) : isForgot ? (
                <Mail className="mr-2 h-4 w-4" />
              ) : isReset ? (
                <Lock className="mr-2 h-4 w-4" />
              ) : isSignup ? (
                <UserPlus className="mr-2 h-4 w-4" />
              ) : (
                <ShieldCheck className="mr-2 h-4 w-4" />
              )}
              {isReset
                ? "Update password"
                : isForgot
                  ? "Send reset email"
                  : isSignup
                    ? "Create account"
                    : "Login"}
            </Button>

            {!isReset && (
              <div className="flex flex-col items-center gap-2 text-sm sm:flex-row sm:justify-between">
                {!isSignup && (
                  <button
                    type="button"
                    onClick={() => setAuthMode(isForgot ? "login" : "forgot")}
                    className="font-semibold text-primary hover:underline"
                  >
                    {isForgot ? "Back to login" : "Forgot password?"}
                  </button>
                )}
                {isForgot && (
                  <span className="text-center text-xs text-muted-foreground sm:text-right">
                    We will send a secure reset link to your email.
                  </span>
                )}
              </div>
            )}
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
const REPORT_HISTORY_CACHE_PREFIX = "report-history-cache";
const REPORT_HISTORY_CACHE_TTL_MS = 10 * 60 * 1000;

interface ReportHistoryCachePayload {
  cachedAt: number;
  items: ReportHistoryItem[];
}

function reportHistoryCacheKey(userId: string): string {
  return `${REPORT_HISTORY_CACHE_PREFIX}:${userId}`;
}

function readReportHistoryCache(userId: string): ReportHistoryItem[] | null {
  if (typeof window === "undefined") return null;

  try {
    const raw = window.localStorage.getItem(reportHistoryCacheKey(userId));
    if (!raw) return null;

    const payload = JSON.parse(raw) as Partial<ReportHistoryCachePayload>;
    if (!payload.cachedAt || !Array.isArray(payload.items)) return null;
    if (Date.now() - payload.cachedAt > REPORT_HISTORY_CACHE_TTL_MS) {
      window.localStorage.removeItem(reportHistoryCacheKey(userId));
      return null;
    }

    return payload.items;
  } catch (error) {
    console.warn("Report history cache read failed", error);
    window.localStorage.removeItem(reportHistoryCacheKey(userId));
    return null;
  }
}

function writeReportHistoryCache(userId: string, items: ReportHistoryItem[]) {
  if (typeof window === "undefined") return;

  try {
    const payload: ReportHistoryCachePayload = {
      cachedAt: Date.now(),
      items,
    };
    window.localStorage.setItem(reportHistoryCacheKey(userId), JSON.stringify(payload));
  } catch (error) {
    console.warn("Report history cache write failed", error);
  }
}

function clearReportHistoryCache(userId: string) {
  if (typeof window === "undefined") return;
  window.localStorage.removeItem(reportHistoryCacheKey(userId));
}

interface DistributorFormatIndexedRecord extends DistributorFormatProfile {
  id: string;
  isActive: boolean;
  deletedAt: string | null;
  sourceFileBlob?: Blob | null;
}

function openDistributorFormatDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DISTRIBUTOR_FORMAT_DB, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(DISTRIBUTOR_FORMAT_STORE)) {
        const store = db.createObjectStore(DISTRIBUTOR_FORMAT_STORE, { keyPath: "id" });
        store.createIndex("distributorName", "distributorName", { unique: false });
        store.createIndex("isActive", "isActive", { unique: false });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("Unable to open distributor DB."));
  });
}

function distributorStoreTransaction(
  mode: IDBTransactionMode,
): Promise<{ db: IDBDatabase; transaction: IDBTransaction; store: IDBObjectStore }> {
  return openDistributorFormatDb().then((db) => {
    const transaction = db.transaction(DISTRIBUTOR_FORMAT_STORE, mode);
    return {
      db,
      transaction,
      store: transaction.objectStore(DISTRIBUTOR_FORMAT_STORE),
    };
  });
}

async function migrateLegacyDistributorFormats() {
  if (typeof window === "undefined") return;
  const migrationKey = `${DISTRIBUTOR_PROFILE_STORAGE_KEY}:indexeddb-migrated`;
  if (window.localStorage.getItem(migrationKey) === "true") return;
  const legacyProfiles = readDistributorProfiles();
  for (const profile of legacyProfiles) {
    await saveDistributorFormat({
      ...profile,
      id: profile.id ?? crypto.randomUUID(),
      active: profile.active !== false,
      deletedAt: profile.deletedAt ?? null,
    });
  }
  window.localStorage.removeItem(DISTRIBUTOR_PROFILE_STORAGE_KEY);
  window.localStorage.setItem(migrationKey, "true");
}

async function saveDistributorFormat(
  profile: DistributorFormatProfile,
  sampleFile?: File | Blob | null,
): Promise<DistributorFormatProfile> {
  const now = new Date().toISOString();
  const id = profile.id ?? crypto.randomUUID();

  const { db, transaction, store } = await distributorStoreTransaction("readwrite");
  await new Promise<void>((resolve, reject) => {
    const existingRequest = store.get(id);
    existingRequest.onsuccess = () => {
      const existing = existingRequest.result as DistributorFormatIndexedRecord | undefined;
      const record: DistributorFormatIndexedRecord = {
        ...profile,
        id,
        profileName: profile.profileName || profile.distributorName,
        active: true,
        isActive: true,
        deletedAt: null,
        createdAt: profile.createdAt ?? existing?.createdAt ?? now,
        lastUpdated: now,
        sourceFileBlob: sampleFile ?? existing?.sourceFileBlob ?? undefined,
      };
      const putRequest = store.put(record);
      putRequest.onerror = () =>
        reject(putRequest.error ?? new Error("Distributor format save failed."));
    };
    existingRequest.onerror = () =>
      reject(existingRequest.error ?? new Error("Distributor format save failed."));
    transaction.oncomplete = () => {
      db.close();
      resolve();
    };
    transaction.onerror = () => {
      db.close();
      reject(transaction.error ?? new Error("Distributor format save failed."));
    };
    transaction.onabort = () => {
      db.close();
      reject(transaction.error ?? new Error("Distributor format save failed."));
    };
  });
  return {
    ...profile,
    id,
    profileName: profile.profileName || profile.distributorName,
    active: true,
    deletedAt: null,
    createdAt: profile.createdAt ?? now,
    lastUpdated: now,
  };
}

async function getDistributorFormats(): Promise<DistributorFormatProfile[]> {
  await migrateLegacyDistributorFormats();
  const deletedNames = readDeletedDistributorProfileNames();
  const { db, transaction, store } = await distributorStoreTransaction("readonly");
  const records = await new Promise<DistributorFormatIndexedRecord[]>((resolve, reject) => {
    let result: DistributorFormatIndexedRecord[] = [];
    const request = store.getAll();
    request.onsuccess = () => {
      result = (request.result ?? []) as DistributorFormatIndexedRecord[];
    };
    request.onerror = () => reject(request.error ?? new Error("Distributor format load failed."));
    transaction.oncomplete = () => {
      db.close();
      resolve(result);
    };
    transaction.onerror = () => {
      db.close();
      reject(transaction.error ?? new Error("Distributor format load failed."));
    };
    transaction.onabort = () => {
      db.close();
      reject(transaction.error ?? new Error("Distributor format load failed."));
    };
  });

  return mergeDistributorProfiles(
    records
      .filter(
        (record) =>
          record.isActive !== false &&
          record.active !== false &&
          !record.deletedAt &&
          !deletedNames.has(record.distributorName.trim().toLowerCase()),
      )
      .map(({ sourceFileBlob, isActive, ...profile }) => profile),
  );
}

async function getDistributorFormatById(id: string): Promise<DistributorFormatProfile | null> {
  const { db, transaction, store } = await distributorStoreTransaction("readonly");
  const record = await new Promise<DistributorFormatIndexedRecord | undefined>(
    (resolve, reject) => {
      let result: DistributorFormatIndexedRecord | undefined;
      const request = store.get(id);
      request.onsuccess = () => {
        result = request.result as DistributorFormatIndexedRecord | undefined;
      };
      request.onerror = () =>
        reject(request.error ?? new Error("Distributor format lookup failed."));
      transaction.oncomplete = () => {
        db.close();
        resolve(result);
      };
      transaction.onerror = () => {
        db.close();
        reject(transaction.error ?? new Error("Distributor format lookup failed."));
      };
      transaction.onabort = () => {
        db.close();
        reject(transaction.error ?? new Error("Distributor format lookup failed."));
      };
    },
  );
  if (!record || record.isActive === false || record.deletedAt) return null;
  const { sourceFileBlob, isActive, ...profile } = record;
  return profile;
}

async function updateDistributorFormat(
  id: string,
  updates: Partial<DistributorFormatProfile>,
): Promise<DistributorFormatProfile> {
  const existing = await getDistributorFormatById(id);
  if (!existing) throw new Error("Distributor format not found.");
  return saveDistributorFormat({ ...existing, ...updates, id });
}

async function deleteDistributorFormat(idOrName: string): Promise<void> {
  const formats = await getDistributorFormats();
  const target = formats.find(
    (format) =>
      format.id === idOrName ||
      format.distributorName.trim().toLowerCase() === idOrName.trim().toLowerCase(),
  );
  if (!target?.id) throw new Error("Distributor format not found.");

  const { db, transaction, store } = await distributorStoreTransaction("readwrite");
  await new Promise<void>((resolve, reject) => {
    const request = store.get(target.id);
    request.onsuccess = () => {
      const record = request.result as DistributorFormatIndexedRecord | undefined;
      if (!record) {
        reject(new Error("Distributor format not found."));
        return;
      }
      const now = new Date().toISOString();
      const updateRequest = store.put({
        ...record,
        active: false,
        isActive: false,
        deletedAt: now,
        lastUpdated: now,
      } satisfies DistributorFormatIndexedRecord);
      updateRequest.onsuccess = () => resolve();
      updateRequest.onerror = () =>
        reject(updateRequest.error ?? new Error("Distributor format delete failed."));
    };
    request.onerror = () => reject(request.error ?? new Error("Distributor format delete failed."));
    transaction.oncomplete = () => {
      db.close();
      resolve();
    };
    transaction.onerror = () => {
      db.close();
      reject(transaction.error ?? new Error("Distributor format delete failed."));
    };
    transaction.onabort = () => {
      db.close();
      reject(transaction.error ?? new Error("Distributor format delete failed."));
    };
  });
}

function readDistributorProfiles(): DistributorFormatProfile[] {
  if (typeof window === "undefined") return [];

  try {
    const raw = window.localStorage.getItem(DISTRIBUTOR_PROFILE_STORAGE_KEY);
    if (!raw) return [];
    const deletedNames = readDeletedDistributorProfileNames();
    const profiles = JSON.parse(raw) as Partial<DistributorFormatProfile>[];
    if (!Array.isArray(profiles)) return [];
    return profiles
      .filter((profile): profile is DistributorFormatProfile =>
        Boolean(profile.distributorName && Array.isArray(profile.numericOrder)),
      )
      .filter((profile) => !deletedNames.has(profile.distributorName.trim().toLowerCase()))
      .map((profile) => ({
        ...profile,
        numericOrder: sanitizeDistributorNumericOrder(profile.numericOrder),
        sourceHeaders: Array.isArray(profile.sourceHeaders) ? profile.sourceHeaders : [],
        sourceColumnPositions: profile.sourceColumnPositions ?? {},
        manualColumnMappings: profile.manualColumnMappings ?? {},
        headerRowRule: profile.headerRowRule || "Auto-detected from PDF header aliases",
        productRowRule:
          profile.productRowRule ||
          "Product code, complete product name, then mapped numeric values",
        groupHeadingRule: profile.groupHeadingRule || "Auto-detect uppercase/group heading lines",
        groupTotalRule: profile.groupTotalRule || "Use PDF Group Total or calculated fallback",
        dateExtractionRule: profile.dateExtractionRule || "Detect From Date and To Date from PDF",
        profileName: profile.profileName || `${profile.distributorName} profile`,
        lastUpdated: profile.lastUpdated || new Date().toISOString(),
        active: profile.active !== false,
      }));
  } catch (error) {
    console.warn("Distributor profile read failed", error);
    return [];
  }
}

function writeDistributorProfiles(profiles: DistributorFormatProfile[]) {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(
    DISTRIBUTOR_PROFILE_STORAGE_KEY,
    JSON.stringify(mergeDistributorProfiles(profiles)),
  );
}

function mergeDistributorProfiles(
  profiles: DistributorFormatProfile[],
): DistributorFormatProfile[] {
  const deletedNames = readDeletedDistributorProfileNames();
  const byName = new Map<string, DistributorFormatProfile>();
  for (const profile of profiles) {
    const key = profile.distributorName?.trim().toLowerCase();
    if (!key || deletedNames.has(key) || profile.active === false) continue;
    const existing = byName.get(key);
    if (!existing) {
      byName.set(key, profile);
      continue;
    }
    const existingTime = Date.parse(existing.lastUpdated || existing.createdAt || "");
    const nextTime = Date.parse(profile.lastUpdated || profile.createdAt || "");
    if (!Number.isFinite(existingTime) || (Number.isFinite(nextTime) && nextTime >= existingTime)) {
      byName.set(key, profile);
    }
  }
  return [...byName.values()].sort(
    (a, b) =>
      Date.parse(b.lastUpdated || b.createdAt || "") -
      Date.parse(a.lastUpdated || a.createdAt || ""),
  );
}

function readDeletedDistributorProfileNames(): Set<string> {
  if (typeof window === "undefined") return new Set();
  try {
    const raw = window.localStorage.getItem(DISTRIBUTOR_DELETED_PROFILE_STORAGE_KEY);
    const names = raw ? (JSON.parse(raw) as unknown) : [];
    return new Set(
      Array.isArray(names) ? names.filter((name): name is string => typeof name === "string") : [],
    );
  } catch (error) {
    console.warn("Distributor deleted profile cache read failed", error);
    return new Set();
  }
}

function rememberDeletedDistributorProfile(distributorName: string) {
  if (typeof window === "undefined") return;
  const names = readDeletedDistributorProfileNames();
  names.add(distributorName.trim().toLowerCase());
  window.localStorage.setItem(DISTRIBUTOR_DELETED_PROFILE_STORAGE_KEY, JSON.stringify([...names]));
}

function forgetDeletedDistributorProfile(distributorName: string) {
  if (typeof window === "undefined") return;
  const names = readDeletedDistributorProfileNames();
  names.delete(distributorName.trim().toLowerCase());
  window.localStorage.setItem(DISTRIBUTOR_DELETED_PROFILE_STORAGE_KEY, JSON.stringify([...names]));
}

type DistributorFormatProfileRow = Tables<"distributor_format_profiles">;

function safeDistributorProfileName(value: string): string {
  return (
    value
      .trim()
      .replace(/[^a-z0-9._-]+/gi, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 80) || "distributor-format"
  );
}

function buildDistributorSampleStorageKey(
  profileId: string,
  profileName: string,
  file: File,
): string {
  const safeName = safeDistributorProfileName(profileName);
  const safeFileName = safeDistributorProfileName(file.name.replace(/\.[^.]+$/, ""));
  const extension = file.name.includes(".") ? file.name.split(".").pop() : "";
  return [
    "indexeddb",
    "distributor-format-samples",
    profileId,
    `${Date.now()}-${safeName || "format"}-${safeFileName || "sample"}${
      extension ? `.${extension}` : ""
    }`,
  ].join("/");
}

function normalizeDistributorFormatName(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function stringArrayFromJson(value: unknown): string[] {
  return Array.isArray(value)
    ? value.filter((item): item is string => typeof item === "string")
    : [];
}

function recordFromJson(value: unknown): Record<string, number> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const result: Record<string, number> = {};
  for (const [key, raw] of Object.entries(value)) {
    if (typeof raw === "number") result[key] = raw;
  }
  return result;
}

function stringRecordFromJson(value: unknown): Record<string, string> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const result: Record<string, string> = {};
  for (const [key, raw] of Object.entries(value)) {
    if (typeof raw === "string") result[key] = raw;
  }
  return result;
}

function distributorProfileFromRow(row: DistributorFormatProfileRow): DistributorFormatProfile {
  return {
    id: row.id,
    distributorName: row.distributor_name,
    profileName: row.profile_name,
    sourceSampleType: row.source_sample_type as DistributorFormatProfile["sourceSampleType"],
    sourceSampleName: row.source_sample_name ?? undefined,
    sourceStoragePath: row.source_storage_path ?? undefined,
    sourceMimeType: row.source_mime_type ?? undefined,
    sourceFileSize: row.source_file_size ?? undefined,
    uploadedAt: row.uploaded_at ?? undefined,
    numericOrder: sanitizeDistributorNumericOrder(row.numeric_order),
    sourceHeaders: stringArrayFromJson(row.source_headers),
    sourceColumnPositions: recordFromJson(row.source_column_positions),
    manualColumnMappings: stringRecordFromJson(row.manual_column_mappings),
    headerRowRule: row.header_row_rule,
    productRowRule: row.product_row_rule,
    groupHeadingRule: row.group_heading_rule,
    groupTotalRule: row.group_total_rule,
    dateExtractionRule: row.date_extraction_rule,
    distributorNameExtractionRule: row.distributor_name_extraction_rule ?? undefined,
    columnMappingRules: row.column_mapping_rules ?? undefined,
    productCodeExtractionRule: row.product_code_extraction_rule ?? undefined,
    productNameExtractionRule: row.product_name_extraction_rule ?? undefined,
    multilineProductNameRule: row.multiline_product_name_rule ?? undefined,
    pageContinuationRule: row.page_continuation_rule ?? undefined,
    createdAt: row.created_at,
    lastUpdated: row.updated_at,
    active: row.active,
  };
}

async function loadDistributorProfilesFromDatabase(): Promise<DistributorFormatProfile[]> {
  const profiles: DistributorFormatProfile[] = [...readDistributorProfiles()];

  const { data, error } = await supabase
    .from("distributor_format_profiles")
    .select("*")
    .eq("active", true)
    .order("updated_at", { ascending: false });
  if (error) {
    console.warn("Distributor format database load failed", error);
  } else {
    profiles.push(...(data ?? []).map(distributorProfileFromRow));
  }

  const { data: userResult } = await supabase.auth.getUser();
  const userId = userResult.user?.id;
  if (userId) {
    try {
      profiles.push(...(await loadDistributorProfilesFromStorage(userId)));
    } catch (storageError) {
      console.warn("Distributor format storage load failed", storageError);
    }
  }

  return mergeDistributorProfiles(profiles);
}

async function loadDistributorProfilesFromStorage(
  userId: string,
): Promise<DistributorFormatProfile[]> {
  const folder = `${userId}/${DISTRIBUTOR_PROFILE_STORAGE_FOLDER}`;
  const deletedNames = readDeletedDistributorProfileNames();
  const { data: files, error } = await supabase.storage
    .from(DISTRIBUTOR_SAMPLE_BUCKET)
    .list(folder, { limit: 1000, sortBy: { column: "updated_at", order: "desc" } });
  if (error) throw error;

  const profiles: DistributorFormatProfile[] = [];
  for (const file of files ?? []) {
    if (!file.name.endsWith(".json")) continue;
    const path = `${folder}/${file.name}`;
    const { data, error: downloadError } = await supabase.storage
      .from(DISTRIBUTOR_SAMPLE_BUCKET)
      .download(path);
    if (downloadError) {
      console.warn("Distributor format profile download failed", downloadError);
      continue;
    }
    try {
      const parsed = JSON.parse(await data.text()) as DistributorFormatProfile;
      if (
        parsed.active !== false &&
        parsed.distributorName &&
        !deletedNames.has(parsed.distributorName.trim().toLowerCase())
      ) {
        profiles.push(parsed);
      }
    } catch (parseError) {
      console.warn("Distributor format profile JSON parse failed", parseError);
    }
  }
  return profiles;
}

async function findDistributorProfileRowByName(
  userId: string,
  distributorName: string,
): Promise<DistributorFormatProfileRow | null> {
  const { data, error } = await supabase
    .from("distributor_format_profiles")
    .select("*")
    .eq("user_id", userId)
    .ilike("distributor_name", distributorName)
    .maybeSingle();
  if (error) throw error;
  return data;
}

async function saveDistributorProfileToDatabase(
  userId: string,
  profile: DistributorFormatProfile,
): Promise<void> {
  const existing = await findDistributorProfileRowByName(userId, profile.distributorName);
  const payload = {
    user_id: userId,
    distributor_name: profile.distributorName,
    profile_name: profile.profileName || profile.distributorName,
    source_sample_type: profile.sourceSampleType ?? "Manual",
    source_sample_name: profile.sourceSampleName ?? null,
    source_storage_path: profile.sourceStoragePath ?? null,
    source_mime_type: profile.sourceMimeType ?? null,
    source_file_size: profile.sourceFileSize ?? null,
    uploaded_at: profile.uploadedAt ?? null,
    numeric_order: profile.numericOrder,
    source_headers: profile.sourceHeaders,
    source_column_positions: profile.sourceColumnPositions,
    manual_column_mappings: profile.manualColumnMappings ?? {},
    header_row_rule: profile.headerRowRule,
    product_row_rule: profile.productRowRule,
    group_heading_rule: profile.groupHeadingRule,
    group_total_rule: profile.groupTotalRule,
    date_extraction_rule: profile.dateExtractionRule,
    distributor_name_extraction_rule: profile.distributorNameExtractionRule ?? null,
    column_mapping_rules: profile.columnMappingRules ?? null,
    product_code_extraction_rule: profile.productCodeExtractionRule ?? null,
    product_name_extraction_rule: profile.productNameExtractionRule ?? null,
    multiline_product_name_rule: profile.multilineProductNameRule ?? null,
    page_continuation_rule: profile.pageContinuationRule ?? null,
    active: profile.active !== false,
    updated_at: profile.lastUpdated,
  };

  if (existing) {
    const { error } = await supabase
      .from("distributor_format_profiles")
      .update(payload)
      .eq("id", existing.id);
    if (error) throw error;
    return;
  }

  const { error } = await supabase.from("distributor_format_profiles").insert({
    ...payload,
    created_at: profile.createdAt ?? profile.lastUpdated,
  });
  if (error) throw error;
}

async function saveDistributorProfileToStorage(
  userId: string,
  profile: DistributorFormatProfile,
): Promise<void> {
  const path = `${userId}/${DISTRIBUTOR_PROFILE_STORAGE_FOLDER}/${safeDistributorProfileName(
    profile.distributorName,
  )}.json`;
  const blob = new Blob([JSON.stringify(profile, null, 2)], {
    type: "application/json",
  });
  const { error } = await supabase.storage.from(DISTRIBUTOR_SAMPLE_BUCKET).upload(path, blob, {
    cacheControl: "60",
    upsert: true,
    contentType: "application/json",
  });
  if (error) throw error;
}

async function deleteDistributorProfileFromDatabase(
  profile: DistributorFormatProfile,
): Promise<void> {
  const { data: userResult, error: userError } = await supabase.auth.getUser();
  if (userError) throw userError;
  const userId = userResult.user?.id;
  if (!userId) throw new Error("Not authenticated");

  let row: DistributorFormatProfileRow | null = null;
  try {
    row = await findDistributorProfileRowByName(userId, profile.distributorName);
  } catch (error) {
    console.warn("Distributor format database lookup failed during delete", error);
  }
  let databaseRowDeleted = false;
  if (row) {
    const { error } = await supabase.from("distributor_format_profiles").delete().eq("id", row.id);
    if (error) throw error;
    databaseRowDeleted = true;
  }

  const profileJsonPath = `${userId}/${DISTRIBUTOR_PROFILE_STORAGE_FOLDER}/${safeDistributorProfileName(
    profile.distributorName,
  )}.json`;
  const pathsToRemove = [profileJsonPath];
  if (row?.source_storage_path) pathsToRemove.push(row.source_storage_path);
  else if (profile.sourceStoragePath) pathsToRemove.push(profile.sourceStoragePath);

  if (pathsToRemove.length) {
    const { error: removeError } = await supabase.storage
      .from(DISTRIBUTOR_SAMPLE_BUCKET)
      .remove(pathsToRemove);
    if (removeError) {
      if (!databaseRowDeleted) throw removeError;
      console.warn("Distributor sample cleanup failed after profile delete", removeError);
    }
  }
}

function distributorNumericLabel(key: DistributorNumericColumnKey): string {
  return DISTRIBUTOR_NUMERIC_COLUMNS.find((column) => column.key === key)?.header ?? key;
}

function formatShortDate(value?: string): string {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "-";
  return date.toLocaleDateString();
}

function sanitizeDistributorNumericOrder(value: unknown): DistributorNumericColumnKey[] {
  const allowed = new Set(DISTRIBUTOR_NUMERIC_COLUMNS.map((column) => column.key));
  const raw = Array.isArray(value) ? value : [];
  const seen = new Set<DistributorNumericColumnKey>();
  const keys = raw.filter((key): key is DistributorNumericColumnKey => {
    if (typeof key !== "string" || !allowed.has(key as DistributorNumericColumnKey)) return false;
    if (seen.has(key as DistributorNumericColumnKey)) return false;
    seen.add(key as DistributorNumericColumnKey);
    return true;
  });
  return keys.length ? keys : DISTRIBUTOR_NUMERIC_COLUMNS.map((column) => column.key);
}

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

function formatBytes(size?: number): string {
  if (!size || !Number.isFinite(size)) return "0 B";
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

async function buildPerformanceWorkbook(report: PerformanceReport): Promise<Blob> {
  const ExcelJS = (await import("exceljs")).default;
  if (report.variant === "monthly-planned") {
    return buildMonthlyPlannedPerformanceWorkbook(ExcelJS, report);
  }
  if (report.variant === "daily") return buildDailyPerformanceWorkbook(ExcelJS, report);

  const wb = new ExcelJS.Workbook();
  if (report.sourceBlob) {
    await wb.xlsx.load(await report.sourceBlob.arrayBuffer());
  } else {
    wb.creator = APP_NAME;
    wb.created = new Date();
  }

  const sheets = [
    { name: "Perf Summary", rows: report.summaryRows },
    ...(report.allRows?.length
      ? [{ name: "All Performance", rows: performanceRowsToSheetRows(report.allRows) }]
      : []),
    { name: "Top Performance", rows: performanceRowsToSheetRows(report.topRows, "Top") },
    { name: "Low Performance", rows: performanceRowsToSheetRows(report.lowRows, "Low") },
  ];

  for (const sheet of sheets) {
    addPerformanceWorksheet(wb, report, sheet.name, sheet.rows);
  }

  for (const teamReport of report.teamReports ?? []) {
    addPerformanceWorksheet(
      wb,
      report,
      `${teamReport.teamName} Perf`,
      performanceRowsToTeamSheetRows(teamReport),
      teamReport.teamName,
    );
  }

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

async function buildDailyPerformanceWorkbook(
  ExcelJS: typeof import("exceljs").default,
  report: PerformanceReport,
): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  wb.creator = APP_NAME;
  wb.created = new Date();

  addDailyFlatPerformanceWorksheet(
    wb,
    "All Performance",
    dailyRowsToPerformanceSheetRows(report.allRows ?? []),
  );
  if (report.topRows.length) {
    addDailyFlatPerformanceWorksheet(
      wb,
      "Top Performance",
      dailyRowsToPerformanceSheetRows(report.topRows, "Top"),
    );
  }
  addDailyFlatPerformanceWorksheet(
    wb,
    "Low Performance",
    dailyRowsToPerformanceSheetRows(report.lowRows, "Low"),
  );

  const teamReports = report.teamReports?.length
    ? report.teamReports
    : buildDailyTeamReports(report.allRows ?? []);
  for (const teamReport of teamReports) {
    addDailyTeamPerformanceWorksheet(
      wb,
      teamReport.teamName,
      dailyRowsToTeamPerformanceSheetRows(teamReport.lowRows),
    );
  }

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

async function buildMonthlyPlannedPerformanceWorkbook(
  ExcelJS: typeof import("exceljs").default,
  report: PerformanceReport,
): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  wb.creator = APP_NAME;
  wb.created = new Date();

  const performanceRows = (report.allRows ?? report.lowRows).filter(
    (row) => row.employeeCode || row.details?.employeeCode,
  );
  const reportDates = report.monthlyPlannedDates ?? monthlyPlannedDates(performanceRows);
  const rowsByTeam = new Map<string, PerformanceRow[]>();
  for (const row of performanceRows) {
    const teamName = row.teamName || String(row.details?.team ?? "").trim() || "Team";
    const teamRows = rowsByTeam.get(teamName) ?? [];
    teamRows.push(row);
    rowsByTeam.set(teamName, teamRows);
  }

  if (!rowsByTeam.size) {
    addMonthlyPlannedTeamPerformanceWorksheet(wb, "All Performance", [], reportDates);
  } else {
    for (const [teamName, teamRows] of [...rowsByTeam.entries()].sort(([a], [b]) =>
      a.localeCompare(b),
    )) {
      addMonthlyPlannedTeamPerformanceWorksheet(
        wb,
        teamName,
        teamRows.sort((a, b) => String(a.name).localeCompare(String(b.name))),
        reportDates,
      );
    }
  }

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

const MONTHLY_PLANNED_BASE_HEADERS = ["Employee Code", "Name", "Designation", "Region"] as const;

const MONTHLY_PLANNED_BASE_WIDTHS = [14, 22, 14, 12];

function addMonthlyPlannedTeamPerformanceWorksheet(
  wb: import("exceljs").Workbook,
  teamName: string,
  rows: PerformanceRow[],
  dateColumns?: string[],
) {
  const dates = dateColumns?.length ? dateColumns : monthlyPlannedDates(rows);
  const headers = [...MONTHLY_PLANNED_BASE_HEADERS, ...dates];
  const ws = wb.addWorksheet(uniqueWorksheetName(wb, `${teamName} Perf`));
  headers.forEach((header, index) => {
    ws.getColumn(index + 1).width =
      index < MONTHLY_PLANNED_BASE_WIDTHS.length ? MONTHLY_PLANNED_BASE_WIDTHS[index] : 34;
  });

  ws.mergeCells(1, 1, 1, headers.length);
  ws.mergeCells(2, 1, 2, headers.length);
  ws.getCell(1, 1).value = teamName;
  ws.getCell(2, 1).value = "All performance";
  ws.getRow(1).height = 28.5;
  ws.getRow(2).height = 21;
  styleDailyPerformanceTitleRow(ws.getRow(1), 16);
  styleDailyPerformanceTitleRow(ws.getRow(2), 11);

  setPlainWorksheetRow(ws, 3, headers);
  styleDailyPerformanceHeader(ws.getRow(3));

  if (!rows.length) {
    setPlainWorksheetRow(ws, 4, ["No performance data available"]);
    styleMonthlyPlannedDataRow(ws.getRow(4), headers.length);
  } else {
    rows.forEach((row, index) => {
      const detailsByDate = new Map(
        (row.monthlyPlannedDailyDetails ?? []).map((detail) => [detail.date, detail]),
      );
      setPlainWorksheetRow(ws, index + 4, [
        row.employeeCode ?? detailValue(row.details ?? {}, "employeeCode"),
        row.name,
        row.designation ?? detailValue(row.details ?? {}, "designation"),
        row.region ?? detailValue(row.details ?? {}, "region"),
        ...dates.map((date) => monthlyPlannedDateCell(detailsByDate.get(date))),
      ]);
      styleMonthlyPlannedDataRow(ws.getRow(index + 4), headers.length);
    });
  }

  ws.views = [{ state: "frozen", ySplit: 3 }];
  ws.autoFilter = { from: "A3", to: `${columnLetter(headers.length)}3` };
}

function monthlyPlannedDates(rows: PerformanceRow[]): string[] {
  return [
    ...new Set(
      rows.flatMap((row) => (row.monthlyPlannedDailyDetails ?? []).map((detail) => detail.date)),
    ),
  ].sort((a, b) => a.localeCompare(b));
}

function monthlyPlannedDateCell(detail?: MonthlyPlannedDayDetail): string {
  if (!detail) return "Absent";
  return [
    `CP: ${detail.cpTime || "N/A"}`,
    `Total working: ${formatPerformanceMinutes(detail.totalWorkingMinutes)}`,
    detail.morningCalls
      ? `Morning: ${formatPerformanceMinutes(detail.morningMinutes)}, first ${detail.morningFirstCall || "N/A"}, last ${detail.morningLastCall || "N/A"}`
      : "Morning: Absent",
    detail.eveningCalls
      ? `Evening: ${formatPerformanceMinutes(detail.eveningMinutes)}, first ${detail.eveningFirstCall || "N/A"}, last ${detail.eveningLastCall || "N/A"}`
      : "Evening: Absent",
    detail.totalCalls
      ? `Calls: ${detail.totalCalls}, planned ${detail.planned}, unplanned ${detail.unplanned}`
      : "Calls: 0 calls",
    detail.activities.length ? `Activities: ${detail.activities.join("; ")}` : "",
  ]
    .filter(Boolean)
    .join("\n");
}

function styleMonthlyPlannedDataRow(row: import("exceljs").Row, columnCount: number) {
  row.height = 92;
  for (let column = 1; column <= columnCount; column++) {
    const cell = row.getCell(column);
    cell.font = { name: "Calibri", size: 10, color: { argb: "FF000000" } };
    cell.alignment = {
      horizontal: column === 2 || column > MONTHLY_PLANNED_BASE_HEADERS.length ? "left" : "center",
      vertical: "middle",
      wrapText: true,
    };
    cell.border = DAILY_PERFORMANCE_BORDER;
  }
}

const DAILY_PERFORMANCE_HEADERS = [
  "Performance Status",
  "Team",
  "Employee Code",
  "Name",
  "Region",
  "Designation",
  "Note",
  "Cp Time",
  "Morning Calls",
  "Morning Hours",
  "Morning Last Call",
  "Evening Calls",
  "Evening Hours",
  "Evening First Call",
  "Evening Last Call",
  "Total Working Hours",
  "Total Calls",
  "Planned Calls",
  "Unplanned Calls",
  "Selfies",
] as const;

const DAILY_TEAM_PERFORMANCE_HEADERS = DAILY_PERFORMANCE_HEADERS.slice(2);

const DAILY_PERFORMANCE_WIDTHS = [
  23, 14, 18, 22, 14, 17, 30, 14, 18, 18, 21, 18, 18, 22, 21, 23, 16, 18, 20, 14,
];

const DAILY_TEAM_PERFORMANCE_WIDTHS = [
  14.5, 21.5, 10, 12, 29, 14, 18, 18, 21, 16, 16, 22, 21, 18, 16, 18, 20, 14,
];

function buildDailyTeamReports(rows: PerformanceRow[]): PerformanceTeamReport[] {
  const byTeam = new Map<string, PerformanceRow[]>();
  for (const row of rows) {
    const teamName = row.teamName || String(row.details?.team ?? "").trim() || "Team";
    const teamRows = byTeam.get(teamName) ?? [];
    teamRows.push(row);
    byTeam.set(teamName, teamRows);
  }

  return [...byTeam.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([teamName, teamRows]) => ({
      teamName,
      allRows: teamRows,
      topRows: teamRows.filter((row) => row.details?.performanceStatus === "Top"),
      lowRows: teamRows
        .filter((row) => row.details?.performanceStatus === "Low")
        .sort((a, b) => Number(a.value || 0) - Number(b.value || 0)),
      summaryRows: [],
    }));
}

function addDailyFlatPerformanceWorksheet(
  wb: import("exceljs").Workbook,
  name: string,
  rows: Record<string, string | number>[],
) {
  const ws = wb.addWorksheet(uniqueWorksheetName(wb, name));
  DAILY_PERFORMANCE_HEADERS.forEach((header, index) => {
    const column = ws.getColumn(index + 1);
    column.width = DAILY_PERFORMANCE_WIDTHS[index];
    column.key = dailyHeaderKey(header);
  });

  setPlainWorksheetRow(ws, 1, DAILY_PERFORMANCE_HEADERS);
  styleDailyPerformanceHeader(ws.getRow(1));
  rows.forEach((row, index) => {
    setPlainWorksheetRow(
      ws,
      index + 2,
      DAILY_PERFORMANCE_HEADERS.map((header) => row[dailyHeaderKey(header)] ?? ""),
    );
    styleDailyPerformanceDataRow(ws.getRow(index + 2));
  });
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.autoFilter = { from: "A1", to: `${columnLetter(DAILY_PERFORMANCE_HEADERS.length)}1` };
}

function addDailyTeamPerformanceWorksheet(
  wb: import("exceljs").Workbook,
  teamName: string,
  rows: Record<string, string | number>[],
) {
  const ws = wb.addWorksheet(uniqueWorksheetName(wb, `${teamName} Perf`));
  DAILY_TEAM_PERFORMANCE_HEADERS.forEach((header, index) => {
    const column = ws.getColumn(index + 1);
    column.width = DAILY_TEAM_PERFORMANCE_WIDTHS[index];
    column.key = dailyHeaderKey(header);
  });

  ws.mergeCells(1, 1, 1, DAILY_TEAM_PERFORMANCE_HEADERS.length);
  ws.mergeCells(2, 1, 2, DAILY_TEAM_PERFORMANCE_HEADERS.length);
  ws.getCell(1, 1).value = teamName;
  ws.getCell(2, 1).value = "Low performance";
  ws.getRow(1).height = 28.5;
  ws.getRow(2).height = 21;
  styleDailyPerformanceTitleRow(ws.getRow(1), 16);
  styleDailyPerformanceTitleRow(ws.getRow(2), 11);

  setPlainWorksheetRow(ws, 3, DAILY_TEAM_PERFORMANCE_HEADERS);
  styleDailyPerformanceHeader(ws.getRow(3));
  rows.forEach((row, index) => {
    setPlainWorksheetRow(
      ws,
      index + 4,
      DAILY_TEAM_PERFORMANCE_HEADERS.map((header) => row[dailyHeaderKey(header)] ?? ""),
    );
    styleDailyPerformanceDataRow(ws.getRow(index + 4));
  });
  ws.views = [{ state: "frozen", ySplit: 3 }];
  ws.autoFilter = { from: "A3", to: `${columnLetter(DAILY_TEAM_PERFORMANCE_HEADERS.length)}3` };
}

function dailyRowsToPerformanceSheetRows(
  rows: PerformanceRow[],
  forcedStatus?: string,
): Record<string, string | number>[] {
  return rows
    .filter((row) => row.details?.employeeCode || row.employeeCode)
    .map((row) => dailyPerformanceSheetRow(row, forcedStatus));
}

function dailyRowsToTeamPerformanceSheetRows(
  rows: PerformanceRow[],
): Record<string, string | number>[] {
  return dailyRowsToPerformanceSheetRows(rows, "Low").map((row) => {
    const { performanceStatus: _performanceStatus, team: _team, ...teamRow } = row;
    return teamRow;
  });
}

function dailyPerformanceSheetRow(
  row: PerformanceRow,
  forcedStatus?: string,
): Record<string, string | number> {
  const details = row.details ?? {};
  const morningHours = detailValue(details, "morningHours");
  const eveningHours = detailValue(details, "eveningHours");
  const totalWorkingHours = detailValue(details, "totalWorkingHours");
  const totalCalls = detailValue(details, "totalCalls") || row.value || 0;
  return {
    performanceStatus: forcedStatus ?? detailValue(details, "performanceStatus"),
    team: row.teamName ?? detailValue(details, "team"),
    employeeCode:
      row.employeeCode ??
      detailValue(details, "employeeCode") ??
      detailValue(details, "employeeId"),
    name: row.name,
    region: row.region ?? detailValue(details, "region"),
    designation: row.designation ?? detailValue(details, "designation"),
    note:
      row.note && !String(row.note).includes("h ")
        ? row.note
        : `${totalCalls} calls, ${formatPerformanceHours(morningHours)} morning, ${formatPerformanceHours(eveningHours)} evening`,
    cpTime: detailValue(details, "cpTime"),
    morningCalls: detailValue(details, "morningCalls"),
    morningHours: formatPerformanceHours(morningHours),
    morningLastCall: detailValue(details, "morningLastCall"),
    eveningCalls: detailValue(details, "eveningCalls"),
    eveningHours: formatPerformanceHours(eveningHours),
    eveningFirstCall: detailValue(details, "eveningFirstCall"),
    eveningLastCall: detailValue(details, "eveningLastCall"),
    totalWorkingHours: formatPerformanceHours(totalWorkingHours),
    totalCalls,
    plannedCalls: detailValue(details, "plannedCalls"),
    unplannedCalls: detailValue(details, "unplannedCalls"),
    selfies: detailValue(details, "selfies"),
  };
}

function detailValue(details: Record<string, string | number>, key: string): string | number {
  return details[key] ?? "";
}

function formatPerformanceHours(value: string | number): string {
  if (value === "") return "";
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric <= 0) return "0 min";
  const hours = Math.floor(numeric);
  const decimal = String(value).split(".")[1]?.replace(/0+$/g, "") ?? "";
  if (decimal) {
    const directMinutes = Number(decimal);
    if (Number.isFinite(directMinutes) && directMinutes < 60) {
      return formatHourMinuteParts(hours, directMinutes);
    }
  }

  const totalMinutes = Math.round(numeric * 60);
  return formatHourMinuteParts(Math.floor(totalMinutes / 60), totalMinutes % 60);
}

function formatPerformanceMinutes(totalMinutes: number): string {
  if (!Number.isFinite(totalMinutes) || totalMinutes <= 0) return "0 min";
  return formatHourMinuteParts(Math.floor(totalMinutes / 60), Math.round(totalMinutes % 60));
}

function formatHourMinuteParts(hours: number, minutes: number): string {
  const parts: string[] = [];
  if (hours) parts.push(`${hours} ${hours === 1 ? "hour" : "hours"}`);
  if (minutes) parts.push(`${minutes} min`);
  return parts.length ? parts.join(" ") : "0 min";
}

function setPlainWorksheetRow(
  ws: import("exceljs").Worksheet,
  rowNumber: number,
  values: readonly (string | number)[],
) {
  values.forEach((value, index) => {
    ws.getCell(rowNumber, index + 1).value = value;
  });
}

function styleDailyPerformanceTitleRow(row: import("exceljs").Row, size: number) {
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE6F3F1" } };
    cell.font = { name: "Calibri", size, bold: true, color: { argb: "FF000000" } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = DAILY_PERFORMANCE_BORDER;
  });
}

function styleDailyPerformanceHeader(row: import("exceljs").Row) {
  row.height = Math.max(row.height ?? 0, 18);
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEFEA" } };
    cell.font = { name: "Calibri", size: 10, bold: true, color: { argb: "FF000000" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = DAILY_PERFORMANCE_BORDER;
  });
}

function styleDailyPerformanceDataRow(row: import("exceljs").Row) {
  row.height = Math.max(row.height ?? 0, 18);
  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    cell.font = { name: "Calibri", size: 10, color: { argb: "FF000000" } };
    cell.alignment = {
      horizontal: [4, 7].includes(colNumber) ? "left" : "center",
      vertical: "middle",
      wrapText: colNumber === 7,
    };
    cell.border = DAILY_PERFORMANCE_BORDER;
  });
}

const DAILY_PERFORMANCE_BORDER: Partial<import("exceljs").Borders> = {
  top: { style: "thin", color: { argb: "FF000000" } },
  left: { style: "thin", color: { argb: "FF000000" } },
  bottom: { style: "thin", color: { argb: "FF000000" } },
  right: { style: "thin", color: { argb: "FF000000" } },
};

function dailyHeaderKey(header: string): string {
  const [first, ...rest] = header.split(/\s+/);
  return `${first.toLowerCase()}${rest.join("")}`;
}

function columnLetter(columnNumber: number): string {
  let value = "";
  let current = columnNumber;
  while (current > 0) {
    const mod = (current - 1) % 26;
    value = String.fromCharCode(65 + mod) + value;
    current = Math.floor((current - mod) / 26);
  }
  return value;
}

function addPerformanceWorksheet(
  wb: import("exceljs").Workbook,
  report: PerformanceReport,
  name: string,
  rows: Record<string, string | number>[],
  teamName?: string,
) {
  const outputRows = rows.length ? rows : [{ status: "No performance data available" }];
  const ws = wb.addWorksheet(uniqueWorksheetName(wb, name));
  const keys = Array.from(
    outputRows.reduce((set, row) => {
      Object.keys(row).forEach((key) => set.add(key));
      return set;
    }, new Set<string>()),
  );

  ws.columns = keys.map((key) => ({
    header: toTitleCase(key.replace(/([A-Z])/g, " $1")),
    key,
    width: Math.max(14, Math.min(32, key.length + 6)),
  }));
  outputRows.forEach((row) => ws.addRow(row));
  ws.spliceRows(
    1,
    0,
    [`${report.title} - ${teamName ? `${teamName} ` : ""}${name}`],
    [report.description],
  );
  ws.mergeCells(1, 1, 1, Math.max(keys.length, 1));
  ws.mergeCells(2, 1, 2, Math.max(keys.length, 1));
  ws.getRow(1).font = { bold: true, size: 14 };
  ws.getRow(2).font = { italic: true, color: { argb: "FF64748B" } };
  ws.getRow(1).alignment = { horizontal: "center" };
  ws.getRow(2).alignment = { horizontal: "center" };
  ws.views = [{ state: "frozen", ySplit: 3 }];

  const header = ws.getRow(3);
  header.font = { bold: true };
  header.alignment = { horizontal: "center" };
  header.eachCell((cell) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE6F4F1" },
    };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });

  ws.eachRow((row) => {
    row.eachCell((cell) => {
      cell.alignment = { vertical: "middle" };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });
  });
}

function performanceRowsToSheetRows(
  rows: PerformanceRow[],
  forcedStatus?: string,
): Record<string, string | number>[] {
  return rows.map((row) => ({
    performanceStatus: forcedStatus ?? row.details?.performanceStatus ?? "",
    team: row.teamName ?? row.details?.team ?? "",
    employeeCode: row.employeeCode ?? row.details?.employeeCode ?? row.details?.employeeId ?? "",
    name: row.name,
    region: row.region ?? row.details?.region ?? "",
    city: row.city ?? row.details?.city ?? "",
    designation: row.designation ?? row.details?.designation ?? "",
    value: row.value,
    note: row.note ?? "",
    ...(row.details ?? {}),
  }));
}

function performanceRowsToTeamSheetRows(
  report: PerformanceTeamReport,
): Record<string, string | number>[] {
  const rowsByCode = new Map<string, PerformanceRow>();
  for (const row of report.allRows) rowsByCode.set(row.employeeCode || row.name, row);
  for (const row of report.topRows) rowsByCode.set(row.employeeCode || row.name, row);
  for (const row of report.lowRows) rowsByCode.set(row.employeeCode || row.name, row);

  const rows = [...rowsByCode.values()].sort((a, b) => {
    const statusRank = (value: PerformanceRow) =>
      value.details?.performanceStatus === "Top"
        ? 0
        : value.details?.performanceStatus === "Low" || value.value === "Failed"
          ? 1
          : 2;
    return statusRank(a) - statusRank(b) || Number(b.value || 0) - Number(a.value || 0);
  });

  return [
    ...report.summaryRows,
    ...performanceRowsToSheetRows(rows.length ? rows : report.lowRows),
  ];
}

function safeWorksheetName(value: string): string {
  return (value || "Performance")
    .replace(/[:\\/?*[\]]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 31);
}

function uniqueWorksheetName(wb: import("exceljs").Workbook, value: string): string {
  const base = safeWorksheetName(value);
  if (!wb.getWorksheet(base)) return base;

  for (let index = 2; index < 100; index++) {
    const suffix = ` ${index}`;
    const candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    if (!wb.getWorksheet(candidate)) return candidate;
  }

  return safeWorksheetName(`${base} ${Date.now()}`);
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

function DistributorSalesPreviewTable({
  rows,
  onCellChange,
}: {
  rows: DistributorSalesRow[];
  onCellChange: (rowIndex: number, key: keyof DistributorSalesRow, value: string) => void;
}) {
  if (!rows.length) {
    return (
      <div className="p-4 text-sm text-muted-foreground">
        No product rows were extracted. Review the warnings above or try a clearer distributor PDF.
      </div>
    );
  }

  return (
    <table className="min-w-max border-collapse text-xs">
      <thead className="sticky top-0 z-10 bg-primary text-primary-foreground">
        <tr>
          {DISTRIBUTOR_SALES_COLUMNS.map((column) => (
            <th
              key={column.key}
              className="min-w-32 border border-primary-foreground/30 px-2 py-2 text-left font-semibold"
            >
              {column.header}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.map((row, rowIndex) => (
          <tr key={`${row.sourcePdfFileName}-${row.productCode}-${rowIndex}`}>
            {DISTRIBUTOR_SALES_COLUMNS.map((column) => {
              const value = row[column.key as keyof DistributorSalesRow];
              return (
                <td
                  key={column.key}
                  className="min-w-32 border border-border bg-card p-0 align-middle"
                >
                  <input
                    className="min-h-8 w-full bg-transparent px-2 py-1 outline-none focus:bg-primary/10 focus:ring-1 focus:ring-primary"
                    value={value ?? ""}
                    onChange={(event) =>
                      onCellChange(
                        rowIndex,
                        column.key as keyof DistributorSalesRow,
                        event.target.value,
                      )
                    }
                    aria-label={`${column.header} row ${rowIndex + 1}`}
                  />
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function DistributorFormatExcelLayout({
  profile,
  editable = false,
  onNumericOrderChange,
}: {
  profile: DistributorFormatProfile;
  editable?: boolean;
  onNumericOrderChange?: (columnIndex: number, key: DistributorNumericColumnKey) => void;
}) {
  const order = sanitizeDistributorNumericOrder(profile.numericOrder);
  const sampleValues: Record<string, string> = {
    productCode: "Code rule",
    productName: "Product name rule",
    tradePrice: "Rate",
    openingQty: "Opening",
    purchaseQty: "Receipt",
    purchaseBonus: "Bns",
    totalStock: "Total Stock",
    salesQty: "Sales",
    salesBonus: "Bns",
    returnQty: "Return",
    returnBonus: "Bns",
    netSaleQty: "Net Qty",
    netSaleBonus: "Net Bns",
    netSaleValue: "Sale Value",
    transferIn: "Transfer In",
    transferOut: "Transfer Out",
    closingQty: "Closing",
    closingValue: "Stock Value",
    todaySales: "Today",
    previousMonthSalesQty: "Previous",
    varianceQty: "Variance",
    variancePercent: "Per%",
  };

  const numericHeaders = DISTRIBUTOR_NUMERIC_COLUMNS.map((column) => column.key);

  return (
    <div className="overflow-auto rounded-md border border-border bg-background">
      <div className="min-w-[980px]">
        <div className="border-b border-border bg-card px-3 py-2 text-center text-sm font-bold uppercase text-foreground">
          {profile.distributorName || "Distributor Name"}
        </div>
        <div className="grid grid-cols-2 border-b border-border text-xs font-semibold text-foreground">
          <div className="border-r border-border px-3 py-2">
            FROM DATE: {profile.dateExtractionRule}
          </div>
          <div className="px-3 py-2">TO DATE: {profile.dateExtractionRule}</div>
        </div>
        <div className="border-b border-border bg-yellow-100 px-3 py-2 text-xs font-bold text-slate-900">
          Group: {profile.groupHeadingRule || "Group heading rule"}
        </div>
        <table className="w-full border-collapse text-xs">
          <thead>
            <tr className="bg-primary text-primary-foreground">
              {[
                "Sr#",
                "Product Code",
                "Product Name",
                ...DISTRIBUTOR_NUMERIC_COLUMNS.map((c) => c.header),
              ].map((header) => (
                <th
                  key={header}
                  className="border border-primary-foreground/30 px-2 py-2 text-left"
                >
                  {header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            <tr>
              <td className="border border-border px-2 py-2">1</td>
              <td className="border border-border px-2 py-2">
                {profile.productCodeExtractionRule || "Product code"}
              </td>
              <td className="border border-border px-2 py-2">
                {profile.productNameExtractionRule || "Product name"}
              </td>
              {numericHeaders.map((key, index) => (
                <td key={key} className="border border-border px-2 py-2 align-top">
                  {editable ? (
                    <select
                      value={order[index] ?? key}
                      onChange={(event) =>
                        onNumericOrderChange?.(
                          index,
                          event.target.value as DistributorNumericColumnKey,
                        )
                      }
                      className="w-36 rounded border border-border bg-background px-2 py-1"
                    >
                      {DISTRIBUTOR_NUMERIC_COLUMNS.map((column) => (
                        <option key={column.key} value={column.key}>
                          {column.header}
                        </option>
                      ))}
                    </select>
                  ) : (
                    sampleValues[order[index] ?? key] ||
                    distributorNumericLabel(order[index] ?? key)
                  )}
                </td>
              ))}
            </tr>
            <tr className="bg-muted/60 font-semibold">
              <td className="border border-border px-2 py-2" />
              <td className="border border-border px-2 py-2" />
              <td className="border border-border px-2 py-2">Group Total</td>
              {numericHeaders.map((key) => (
                <td key={key} className="border border-border px-2 py-2">
                  {profile.groupTotalRule || "Total mapping"}
                </td>
              ))}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

function DistributorProfileField({
  label,
  value,
  onChange,
  multiline = false,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  multiline?: boolean;
}) {
  return (
    <label className="block text-sm">
      <span className="mb-1 block font-semibold text-foreground">{label}</span>
      {multiline ? (
        <textarea
          value={value}
          onChange={(event) => onChange(event.target.value)}
          className="min-h-20 w-full rounded-md border border-border bg-background px-3 py-2 outline-none focus:ring-2 focus:ring-primary/30"
        />
      ) : (
        <input
          value={value}
          onChange={(event) => onChange(event.target.value)}
          className="h-10 w-full rounded-md border border-border bg-background px-3 outline-none focus:ring-2 focus:ring-primary/30"
        />
      )}
    </label>
  );
}

function ManualDistributorMappings({
  mappings,
  onChange,
}: {
  mappings: Record<string, string>;
  onChange: (key: string, value: string) => void;
}) {
  return (
    <div className="rounded-md border border-border bg-background/60 p-3">
      <div className="text-xs font-bold uppercase text-primary">Manual column mapping</div>
      <p className="mt-1 text-sm text-muted-foreground">
        Select a known PDF header or type a custom header for each standard Excel column.
      </p>
      <div className="mt-3 grid max-h-80 gap-3 overflow-y-auto pr-1 sm:grid-cols-2">
        {MANUAL_DISTRIBUTOR_MAPPING_FIELDS.map((field) => (
          <div key={field} className="rounded-md border border-border bg-card p-2">
            <div className="text-xs font-semibold text-foreground">{field}</div>
            <select
              value={
                MANUAL_DISTRIBUTOR_HEADER_OPTIONS.includes(mappings[field] ?? "")
                  ? (mappings[field] ?? "")
                  : ""
              }
              onChange={(event) => onChange(field, event.target.value)}
              className="mt-2 h-9 w-full rounded-md border border-border bg-background px-2 text-sm outline-none focus:ring-2 focus:ring-primary/30"
            >
              <option value="">Select PDF header</option>
              {MANUAL_DISTRIBUTOR_HEADER_OPTIONS.filter(Boolean).map((option) => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
            </select>
            <input
              value={mappings[field] ?? ""}
              onChange={(event) => onChange(field, event.target.value)}
              placeholder="Or type custom header"
              className="mt-2 h-9 w-full rounded-md border border-border bg-background px-2 text-sm outline-none focus:ring-2 focus:ring-primary/30"
            />
          </div>
        ))}
      </div>
    </div>
  );
}

function ProfileDetail({ label, value }: { label: string; value?: string }) {
  return (
    <div className="rounded-md border border-border bg-background/60 p-3">
      <div className="text-xs font-bold uppercase text-primary">{label}</div>
      <div className="mt-1 whitespace-pre-wrap text-sm text-foreground">{value || "-"}</div>
    </div>
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
  compact = false,
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
  compact?: boolean;
}) {
  return (
    <Card
      className={`flex flex-col border-white/45 bg-[var(--gradient-card)] shadow-[var(--shadow-soft)] transition-shadow hover:shadow-[var(--shadow-elegant)] ${
        compact ? "h-[22rem] p-4" : "min-h-[16rem] p-4 sm:p-5"
      }`}
    >
      <div className={compact ? "mb-3 space-y-3" : "mb-4 space-y-4"}>
        <div className="flex items-start justify-between gap-3">
          <div className="flex min-h-9 items-center gap-2 text-xs font-bold uppercase text-primary">
            Step {step}
            {ready && <Check className="h-3.5 w-3.5 text-success" />}
          </div>
          <Button
            variant="outline"
            size="sm"
            onClick={onClick}
            className="min-h-9 shrink-0 whitespace-nowrap px-4"
          >
            <Upload className="mr-1.5 h-3.5 w-3.5" />
            Upload
          </Button>
        </div>
        <div className="min-w-0">
          <h2
            className={`flex min-w-0 items-start gap-3 font-semibold leading-tight text-foreground ${
              compact ? "text-base sm:text-lg" : "text-base sm:text-lg 2xl:text-xl"
            }`}
          >
            <span
              className={`mt-0.5 flex shrink-0 items-center justify-center rounded-xl bg-card text-primary shadow-[var(--shadow-inset)] ${
                compact ? "h-8 w-8" : "h-9 w-9"
              }`}
            >
              {icon}
            </span>
            <span className="min-w-0 whitespace-normal break-words">{title}</span>
          </h2>
          <p
            className={`mt-2 max-w-[28rem] text-sm leading-relaxed text-muted-foreground ${
              compact ? "" : "sm:text-base"
            }`}
          >
            {subtitle}
          </p>
        </div>
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
        className={`report-card-scroll mt-auto overflow-y-auto rounded-2xl border shadow-[var(--shadow-inset)] transition ${
          compact ? "h-[115px] max-h-[115px] p-5" : "min-h-28 max-h-52 p-3 sm:min-h-32"
        } ${
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

function MultiFileList({
  files,
  onClear,
  onRemove,
}: {
  files: File[];
  onClear: () => void;
  onRemove: (file: File) => void;
}) {
  return (
    <div className="space-y-2">
      <div className="flex items-center justify-between gap-3 text-xs font-medium text-muted-foreground">
        <div className="flex min-w-0 items-center gap-2">
          <FileSpreadsheet className="h-3.5 w-3.5 shrink-0" />
          <span className="truncate">{files.length} file(s) selected</span>
        </div>
        <Button
          type="button"
          size="sm"
          variant="outline"
          onClick={(event) => {
            event.stopPropagation();
            onClear();
          }}
          className="h-7 shrink-0 px-2 text-xs"
        >
          Clear All
        </Button>
      </div>
      <div className="report-card-scroll max-h-44 space-y-1.5 overflow-y-auto pr-1">
        {files.map((file) => (
          <FilePill
            key={`${file.name}-${file.lastModified}-${file.size}`}
            name={file.name}
            onRemove={() => onRemove(file)}
            color="accent"
          />
        ))}
      </div>
    </div>
  );
}

function sameFile(a: File, b: File): boolean {
  return a.name === b.name && a.lastModified === b.lastModified && a.size === b.size;
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
