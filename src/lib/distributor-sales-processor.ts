import ExcelJS from "exceljs";

export interface DistributorSalesRow {
  rowType: "product" | "groupTotal";
  groupName: string;
  distributorName: string;
  fromDate: string;
  toDate: string;
  productCode: string;
  productName: string;
  tradePrice: number | "";
  openingQty: number | "";
  openingValue: number | "";
  purchaseQty: number | "";
  purchaseBonus: number | "";
  purchaseValue: number | "";
  totalStock: number | "";
  salesQty: number | "";
  salesBonus: number | "";
  returnQty: number | "";
  returnBonus: number | "";
  netSaleQty: number | "";
  netSaleBonus: number | "";
  netSaleValue: number | "";
  transferIn: number | "";
  transferOut: number | "";
  closingQty: number | "";
  closingBonus: number | "";
  closingValue: number | "";
  todaySales: number | "";
  previousMonthSalesQty: number | "";
  previousMonthSalesValue: number | "";
  varianceQty: number | "";
  variancePercent: number | "";
  sourcePdfFileName: string;
  remarksWarnings: string;
}

export interface DistributorFileSummary {
  fileName: string;
  distributorName: string;
  fromDate: string;
  toDate: string;
  rowsExtracted: number;
  mappedColumns: string[];
  profileName: string;
  profileStatus: DistributorProfileStatus;
  suggestedNumericOrder: DistributorNumericColumnKey[];
  sourceHeaders: string[];
  groups: DistributorGroupSummary[];
  warnings: string[];
}

export interface DistributorGroupSummary {
  name: string;
  productCount: number;
  totalSource: "pdf" | "calculated" | "missing";
}

export interface DistributorSalesResult {
  rows: DistributorSalesRow[];
  summaries: DistributorFileSummary[];
  warnings: string[];
}

export type DistributorProfileStatus = "saved" | "auto-detected" | "unrecognized";

export type DistributorNumericColumnKey = Extract<
  keyof DistributorSalesRow,
  | "tradePrice"
  | "openingQty"
  | "openingValue"
  | "purchaseQty"
  | "purchaseBonus"
  | "purchaseValue"
  | "totalStock"
  | "salesQty"
  | "salesBonus"
  | "returnQty"
  | "returnBonus"
  | "netSaleQty"
  | "netSaleBonus"
  | "netSaleValue"
  | "transferIn"
  | "transferOut"
  | "closingQty"
  | "closingBonus"
  | "closingValue"
  | "todaySales"
  | "previousMonthSalesQty"
  | "previousMonthSalesValue"
  | "varianceQty"
  | "variancePercent"
>;

export interface DistributorFormatProfile {
  distributorName: string;
  profileName: string;
  sourceSampleType?: "PDF" | "Excel" | "Screenshot" | "Manual";
  sourceSampleName?: string;
  sourceStoragePath?: string;
  sourceMimeType?: string;
  sourceFileSize?: number;
  uploadedAt?: string;
  numericOrder: DistributorNumericColumnKey[];
  sourceHeaders: string[];
  sourceColumnPositions: Record<string, number>;
  manualColumnMappings?: Record<string, string>;
  headerRowRule: string;
  productRowRule: string;
  groupHeadingRule: string;
  groupTotalRule: string;
  dateExtractionRule: string;
  distributorNameExtractionRule?: string;
  columnMappingRules?: string;
  productCodeExtractionRule?: string;
  productNameExtractionRule?: string;
  multilineProductNameRule?: string;
  pageContinuationRule?: string;
  createdAt?: string;
  lastUpdated: string;
  active: boolean;
}

export interface ProcessDistributorSalesOptions {
  profiles?: DistributorFormatProfile[];
}

type PdfJs = typeof import("pdfjs-dist");

type PdfTextItem = {
  str?: string;
  transform?: number[];
};

type PositionedText = {
  x: number;
  y: number;
  text: string;
};

type PdfLine = {
  pageNumber: number;
  order: number;
  y: number;
  text: string;
  items: PositionedText[];
};

type StandardColumn = {
  key: keyof DistributorSalesRow;
  header: string;
  width: number;
  aliases?: string[];
  numeric?: boolean;
};

type HeaderDetection = {
  y: number;
  order: number;
  columnMap: Map<keyof DistributorSalesRow, number>;
  anchors: Array<{ key: keyof DistributorSalesRow; x: number }>;
  sourceHeaders: string[];
};

type ParsedProductLine = {
  productCode: string;
  productName: string;
  numbers: number[];
};

type AminParsedLine = {
  productCode: string;
  productName: string;
  values: Array<number | "">;
  ignoredWarnings: string[];
};

const STANDARD_COLUMNS: StandardColumn[] = [
  { key: "distributorName", header: "Distributor Name", width: 24 },
  { key: "fromDate", header: "From Date", width: 14 },
  { key: "toDate", header: "To Date", width: 14 },
  {
    key: "productCode",
    header: "Product Code",
    width: 16,
    aliases: ["code", "product code", "item code"],
  },
  {
    key: "productName",
    header: "Product Name",
    width: 34,
    aliases: ["description", "product name", "item name", "product"],
  },
  {
    key: "tradePrice",
    header: "T.P / Trade Price",
    width: 15,
    aliases: ["trade price", "t.p.", "tp", "rate"],
    numeric: true,
  },
  {
    key: "openingQty",
    header: "Opening Qty",
    width: 14,
    aliases: ["open stock", "opening stock", "opening qty", "opening"],
    numeric: true,
  },
  {
    key: "openingValue",
    header: "Opening Value",
    width: 15,
    aliases: ["opening value", "open stock value"],
    numeric: true,
  },
  {
    key: "purchaseQty",
    header: "Purchase Qty / Receipt Qty",
    width: 18,
    aliases: ["receipt qty", "purchase qty", "purchase", "rec qty"],
    numeric: true,
  },
  {
    key: "purchaseBonus",
    header: "Purchase Bonus / Receipt Bonus",
    width: 20,
    aliases: ["receipt bonus", "purchase bonus", "bon", "bns"],
    numeric: true,
  },
  {
    key: "purchaseValue",
    header: "Purchase Value / Receipt Value",
    width: 20,
    aliases: ["receipt value", "purchase value"],
    numeric: true,
  },
  {
    key: "totalStock",
    header: "Total Stock",
    width: 14,
    aliases: ["total stock", "total stk", "stock total"],
    numeric: true,
  },
  {
    key: "salesQty",
    header: "Sales Qty",
    width: 13,
    aliases: ["sales qty", "sale qty", "discounted sale", "sale"],
    numeric: true,
  },
  {
    key: "salesBonus",
    header: "Sales Bonus",
    width: 13,
    aliases: ["sales bns", "sale bonus", "sale bns"],
    numeric: true,
  },
  {
    key: "returnQty",
    header: "Return Qty",
    width: 13,
    aliases: ["return qty", "ret qty", "sales return", "return"],
    numeric: true,
  },
  {
    key: "returnBonus",
    header: "Return Bonus",
    width: 14,
    aliases: ["return bns", "return bonus"],
    numeric: true,
  },
  {
    key: "netSaleQty",
    header: "Net Sale Qty",
    width: 14,
    aliases: ["net sale qty", "net sale", "net qty"],
    numeric: true,
  },
  {
    key: "netSaleBonus",
    header: "Net Sale Bonus",
    width: 15,
    aliases: ["net sale bns", "net bonus"],
    numeric: true,
  },
  {
    key: "netSaleValue",
    header: "Net Sale Value",
    width: 15,
    aliases: ["sale value", "net sale value", "net value", "value"],
    numeric: true,
  },
  {
    key: "transferIn",
    header: "Transfer In",
    width: 13,
    aliases: ["transfer in", "t in", "in"],
    numeric: true,
  },
  {
    key: "transferOut",
    header: "Transfer Out",
    width: 13,
    aliases: ["transfer out", "t out", "out"],
    numeric: true,
  },
  {
    key: "closingQty",
    header: "Closing Qty",
    width: 14,
    aliases: [
      "closing stock qty",
      "closing balance qty",
      "closing qty",
      "closing stock",
      "closing balance",
    ],
    numeric: true,
  },
  {
    key: "closingBonus",
    header: "Closing Bonus",
    width: 15,
    aliases: ["closing bns", "closing bonus"],
    numeric: true,
  },
  {
    key: "closingValue",
    header: "Closing Value",
    width: 15,
    aliases: ["stock value", "closing value", "closing balance value"],
    numeric: true,
  },
  { key: "todaySales", header: "Today Sales", width: 14, aliases: ["today sales"], numeric: true },
  {
    key: "previousMonthSalesQty",
    header: "Previous Month Sales Qty",
    width: 22,
    aliases: ["p. month sales", "previous month sales qty"],
    numeric: true,
  },
  {
    key: "previousMonthSalesValue",
    header: "Previous Month Sales Value",
    width: 24,
    aliases: ["previous month sales value"],
    numeric: true,
  },
  {
    key: "varianceQty",
    header: "Variance Qty",
    width: 14,
    aliases: ["variance qty", "var qty"],
    numeric: true,
  },
  {
    key: "variancePercent",
    header: "Variance Per%",
    width: 14,
    aliases: ["variance per", "variance %", "var %", "per%"],
    numeric: true,
  },
  { key: "sourcePdfFileName", header: "Source PDF File Name", width: 30 },
  { key: "remarksWarnings", header: "Remarks / Warnings", width: 38 },
];

const TEMPLATE_COLUMNS: StandardColumn[] = [
  { key: "productCode", header: "Product code", width: 14 },
  { key: "productName", header: "Product name", width: 28 },
  { key: "tradePrice", header: "Trade Price", width: 13, numeric: true },
  { key: "openingQty", header: "Open Stock", width: 13, numeric: true },
  { key: "purchaseQty", header: "Qty", width: 11, numeric: true },
  { key: "purchaseBonus", header: "Bns", width: 10, numeric: true },
  { key: "totalStock", header: "Total Stock", width: 13, numeric: true },
  { key: "salesQty", header: "Qty", width: 11, numeric: true },
  { key: "salesBonus", header: "Bns", width: 10, numeric: true },
  { key: "returnQty", header: "Qty", width: 11, numeric: true },
  { key: "returnBonus", header: "Bns", width: 10, numeric: true },
  { key: "netSaleQty", header: "Qty", width: 11, numeric: true },
  { key: "netSaleBonus", header: "Bns", width: 10, numeric: true },
  { key: "netSaleValue", header: "Sale Value", width: 13, numeric: true },
  { key: "transferIn", header: "Transfer In", width: 13, numeric: true },
  { key: "transferOut", header: "Transfer Out", width: 13, numeric: true },
  { key: "closingQty", header: "Closing Stock", width: 14, numeric: true },
  { key: "closingValue", header: "Stock Value", width: 13, numeric: true },
  { key: "todaySales", header: "Today Sale", width: 13, numeric: true },
  { key: "previousMonthSalesQty", header: "Previous Month", width: 17, numeric: true },
  { key: "varianceQty", header: "Qty", width: 11, numeric: true },
  { key: "variancePercent", header: "Per%", width: 11, numeric: true },
];

export const DISTRIBUTOR_SALES_COLUMNS = [
  { key: "rowType", header: "Row Type", numeric: false },
  { key: "groupName", header: "Group Name", numeric: false },
  { key: "distributorName", header: "Distributor Name", numeric: false },
  { key: "fromDate", header: "From Date", numeric: false },
  { key: "toDate", header: "To Date", numeric: false },
  ...TEMPLATE_COLUMNS.map(({ key, header, numeric }) => ({
    key,
    header,
    numeric: Boolean(numeric),
  })),
  { key: "sourcePdfFileName", header: "Source PDF File Name", numeric: false },
  { key: "remarksWarnings", header: "Remarks / Warnings", numeric: false },
];

export const DISTRIBUTOR_NUMERIC_COLUMNS = TEMPLATE_COLUMNS.filter((column) => column.numeric).map(
  ({ key, header }) => ({ key: key as DistributorNumericColumnKey, header }),
);

const TEMPLATE_NUMERIC_KEYS = new Set<DistributorNumericColumnKey>(
  DISTRIBUTOR_NUMERIC_COLUMNS.map((column) => column.key),
);

const FALLBACK_NUMERIC_ORDER: DistributorNumericColumnKey[] = [
  "tradePrice",
  "openingQty",
  "purchaseQty",
  "purchaseBonus",
  "totalStock",
  "salesQty",
  "salesBonus",
  "returnQty",
  "returnBonus",
  "netSaleQty",
  "netSaleBonus",
  "netSaleValue",
  "transferIn",
  "transferOut",
  "closingQty",
  "closingValue",
  "todaySales",
  "previousMonthSalesQty",
  "varianceQty",
  "variancePercent",
];

const AMIN_TRADERS_NUMERIC_ORDER: DistributorNumericColumnKey[] = [
  "tradePrice",
  "openingQty",
  "purchaseQty",
  "totalStock",
  "netSaleQty",
  "netSaleBonus",
  "netSaleValue",
  "closingQty",
  "closingValue",
  "todaySales",
];

export const BUILT_IN_DISTRIBUTOR_PROFILES: DistributorFormatProfile[] = [
  {
    distributorName: "AMIN TRADERS - CHARSADDA",
    profileName: "AMIN TRADERS - CHARSADDA built-in mapping",
    sourceSampleType: "Manual",
    sourceSampleName: "Built-in AMIN TRADERS mapping",
    numericOrder: AMIN_TRADERS_NUMERIC_ORDER,
    sourceHeaders: [
      "Item Description",
      "Rate",
      "Pack",
      "Opening Balance",
      "Purchase",
      "Purchase Return",
      "Purchase Total",
      "Net Sales",
      "Bonus",
      "Value",
      "Adjustment",
      "Closing Balance",
      "Closing Value",
      "Today Sale",
      "Today Return",
    ],
    sourceColumnPositions: {},
    headerRowRule:
      "Sale And Stock Report header with Item Description, Rate, Pack, Opening Balance, Purchase, Sale, Closing.",
    productRowRule:
      "Product code and complete product name from Item Description; ignore Pack so values cannot shift.",
    groupHeadingRule: "Group lines such as 58001 / ALPHA-I (NEUTRO), carried across pages.",
    groupTotalRule: "Total For Group rows are detected as group totals.",
    dateExtractionRule: "Read From Date and To Date from report header.",
    distributorNameExtractionRule: "Detect AMIN TRADERS and CHARSADDA in report header.",
    columnMappingRules:
      "Ignore Pack, Purchase Return, Adjustment, and Today Return. Map only confirmed AMIN source fields.",
    productCodeExtractionRule: "First 4-8 digits at the start of Item Description.",
    productNameExtractionRule: "Text after product code before Rate.",
    multilineProductNameRule: "Merge text-only continuation lines into the previous product name.",
    pageContinuationRule: "Keep current group across page breaks until next group or group total.",
    createdAt: "2026-06-16T00:00:00.000Z",
    lastUpdated: "2026-06-16T00:00:00.000Z",
    active: true,
  },
];

const DATA_COLUMNS = STANDARD_COLUMNS.filter((column) => column.aliases);
const TOTAL_LINE = /\b(group|page|grand)?\s*total\b/i;
const FOOTER_LINE = /\b(page\s*\d+|printed|generated|software|powered by)\b/i;
const UNRECOGNIZED_FORMAT_WARNING =
  "PDF format not recognized. Please verify column mapping before export.";

async function loadPdfJs(): Promise<PdfJs> {
  const [pdfjsLib, worker] = await Promise.all([
    import("pdfjs-dist"),
    import("pdfjs-dist/build/pdf.worker.min.mjs?url"),
  ]);
  pdfjsLib.GlobalWorkerOptions.workerSrc = worker.default;
  return pdfjsLib;
}

export async function processDistributorSalesPdfs(
  files: File[],
  options: ProcessDistributorSalesOptions = {},
): Promise<DistributorSalesResult> {
  const rows: DistributorSalesRow[] = [];
  const summaries: DistributorFileSummary[] = [];
  const warnings: string[] = [];

  for (const file of files) {
    try {
      const parsed = await parseDistributorPdf(file, options.profiles ?? []);
      rows.push(...parsed.rows);
      summaries.push(parsed.summary);
      warnings.push(...parsed.summary.warnings.map((warning) => `${file.name}: ${warning}`));
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unable to read this PDF.";
      warnings.push(`${file.name}: ${message}`);
      summaries.push({
        fileName: file.name,
        distributorName: distributorNameFromFile(file.name),
        fromDate: "",
        toDate: "",
        rowsExtracted: 0,
        mappedColumns: [],
        profileName: "No profile",
        profileStatus: "unrecognized",
        suggestedNumericOrder: [...FALLBACK_NUMERIC_ORDER],
        sourceHeaders: [],
        groups: [],
        warnings: [message],
      });
    }
  }

  return { rows, summaries, warnings };
}

async function parseDistributorPdf(
  file: File,
  profiles: DistributorFormatProfile[],
): Promise<{
  rows: DistributorSalesRow[];
  summary: DistributorFileSummary;
}> {
  const pdfjs = await loadPdfJs();
  const pdf = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
  const lines: PdfLine[] = [];
  let rawText = "";

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const items: PositionedText[] = [];
    for (const item of content.items as PdfTextItem[]) {
      if (!item.str?.trim() || !item.transform) continue;
      items.push({ text: item.str.trim(), x: item.transform[4], y: item.transform[5] });
    }
    const pageLines = groupItemsIntoLines(items, pageNumber, lines.length);
    lines.push(...pageLines);
    rawText += `${pageLines.map((line) => line.text).join("\n")}\n`;
  }

  const distributorName = detectDistributorName(rawText, file.name);
  const { fromDate, toDate } = detectDateRange(rawText);
  const header = detectHeader(lines);
  const inactiveProfileNames = new Set(
    profiles
      .filter((profile) => profile.active === false)
      .map((profile) => normalizeProfileName(profile.distributorName)),
  );
  const availableBuiltInProfiles = BUILT_IN_DISTRIBUTOR_PROFILES.filter(
    (profile) => !inactiveProfileNames.has(normalizeProfileName(profile.distributorName)),
  );
  const savedProfile = findDistributorProfile(
    [...profiles, ...availableBuiltInProfiles],
    distributorName,
  );
  const detectedNumericOrder = header ? detectNumericOrder(header) : [];
  const activeNumericOrder = sanitizeNumericOrder(
    savedProfile?.numericOrder?.length
      ? savedProfile.numericOrder
      : detectedNumericOrder.length >= 8
        ? detectedNumericOrder
        : FALLBACK_NUMERIC_ORDER,
  );
  const profileStatus: DistributorProfileStatus = savedProfile
    ? "saved"
    : detectedNumericOrder.length >= 8
      ? "auto-detected"
      : "unrecognized";
  const profileName =
    savedProfile?.profileName ??
    (profileStatus === "auto-detected" ? "Auto-detected mapping" : "Unrecognized format");
  const warnings: string[] = [];
  if (!fromDate) warnings.push("From Date was not detected confidently and was left blank.");
  if (!toDate) warnings.push("To Date was not detected confidently and was left blank.");
  if (!savedProfile) {
    warnings.push("No saved format found; using auto-detected distributor sales mapping.");
    if (detectedNumericOrder.length < 8 && !header) warnings.push(UNRECOGNIZED_FORMAT_WARNING);
  }

  if (isAminTradersFormat(rawText, distributorName, savedProfile)) {
    return parseAminTradersPdf({
      file,
      lines,
      distributorName: normalizeAminDistributorName(distributorName),
      fromDate,
      toDate,
      profileName,
      profileStatus,
      numericOrder: activeNumericOrder,
      warnings,
    });
  }

  if (!header) {
    warnings.push(
      "Could not confidently detect the product table header; using text-row fallback.",
    );
    const fallbackRows = rowsFromTextFallback(lines, {
      distributorName,
      fromDate,
      toDate,
      sourcePdfFileName: file.name,
      groupName: "Ungrouped",
    });
    const fallbackGroups = summarizeGroups(fallbackRows);
    return {
      rows: fallbackRows,
      summary: {
        fileName: file.name,
        distributorName,
        fromDate,
        toDate,
        rowsExtracted: fallbackRows.filter((row) => row.rowType === "product").length,
        mappedColumns: FALLBACK_NUMERIC_ORDER.map((key) => columnHeader(key)),
        profileName,
        profileStatus: fallbackRows.length ? "auto-detected" : profileStatus,
        suggestedNumericOrder: activeNumericOrder,
        sourceHeaders: [],
        groups: fallbackGroups,
        warnings,
      },
    };
  }

  const mappedColumns = [...header.columnMap.keys()];
  for (const required of ["productCode", "productName"]) {
    if (!header.columnMap.has(required as keyof DistributorSalesRow)) {
      warnings.push(`Missing or uncertain column mapping: ${columnHeader(required)}.`);
    }
  }

  const rows: DistributorSalesRow[] = [];
  const tableLines = lines
    .filter((line) => line.order > header.order)
    .sort((a, b) => a.order - b.order);
  let currentGroup = "Ungrouped";

  for (const line of tableLines) {
    if (isGroupTotalLine(line.text)) {
      const totalRow = rowFromTotalLine(
        line.text,
        {
          distributorName,
          fromDate,
          toDate,
          sourcePdfFileName: file.name,
          groupName: currentGroup,
        },
        activeNumericOrder,
      );
      if (totalRow) rows.push(totalRow);
      continue;
    }
    const groupName = detectGroupHeading(line.text);
    if (groupName) {
      currentGroup = groupName;
      continue;
    }
    if (!isLikelyProductLine(line.text)) continue;
    const row = rowFromLine(
      line,
      header,
      {
        distributorName,
        fromDate,
        toDate,
        sourcePdfFileName: file.name,
        groupName: currentGroup,
      },
      activeNumericOrder,
    );
    if (!row) continue;
    rows.push(row);
  }

  const groupSummaries = summarizeGroups(rows);
  if (!rows.some((row) => row.rowType === "product")) {
    warnings.push("No product rows were extracted from the detected table.");
  }
  for (const group of groupSummaries) {
    if (group.totalSource === "calculated" && group.productCount > 0) {
      warnings.push(
        `Group total was not clearly detected for ${group.name}; it will be calculated from extracted rows.`,
      );
    }
  }

  return {
    rows,
    summary: {
      fileName: file.name,
      distributorName,
      fromDate,
      toDate,
      rowsExtracted: rows.filter((row) => row.rowType === "product").length,
      mappedColumns: mappedColumns.map((key) => columnHeader(key)),
      profileName,
      profileStatus,
      suggestedNumericOrder: activeNumericOrder,
      sourceHeaders: header.sourceHeaders,
      groups: groupSummaries,
      warnings,
    },
  };
}

function groupItemsIntoLines(
  items: PositionedText[],
  pageNumber: number,
  startingOrder: number,
): PdfLine[] {
  const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
  const lines: PdfLine[] = [];
  for (const item of sorted) {
    let line = lines.find((candidate) => Math.abs(candidate.y - item.y) <= 3);
    if (!line) {
      line = { pageNumber, order: startingOrder + lines.length, y: item.y, text: "", items: [] };
      lines.push(line);
    }
    line.items.push(item);
    line.y = (line.y + item.y) / 2;
  }

  for (const line of lines) {
    line.items.sort((a, b) => a.x - b.x);
    line.text = line.items
      .map((item) => item.text)
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();
  }

  return lines;
}

function detectHeader(lines: PdfLine[]): HeaderDetection | null {
  let best: HeaderDetection | null = null;

  for (let index = 0; index < lines.length; index++) {
    const headerLines = [lines[index], lines[index + 1], lines[index + 2]].filter(Boolean);
    const headerText = normalizeHeader(headerLines.map((line) => line.text).join(" "));
    const columnMap = new Map<keyof DistributorSalesRow, number>();
    const anchors: Array<{ key: keyof DistributorSalesRow; x: number }> = [];

    for (const column of DATA_COLUMNS) {
      const aliases = column.aliases ?? [];
      const matched = aliases.some((alias) => headerText.includes(normalizeHeader(alias)));
      if (!matched) continue;
      const x = findAliasX(headerLines, aliases);
      columnMap.set(column.key, x);
      anchors.push({ key: column.key, x });
    }

    const score =
      Number(columnMap.has("productCode")) * 2 +
      Number(columnMap.has("productName")) * 2 +
      [...columnMap.keys()].filter((key) => key !== "productCode" && key !== "productName").length;

    if (score >= 4 && (!best || columnMap.size > best.columnMap.size)) {
      best = {
        y: headerLines[headerLines.length - 1].y,
        order: headerLines[headerLines.length - 1].order,
        columnMap,
        anchors: anchors.sort((a, b) => a.x - b.x),
        sourceHeaders: headerLines.map((line) => line.text),
      };
    }
  }

  return best;
}

function detectNumericOrder(header: HeaderDetection): DistributorNumericColumnKey[] {
  const seen = new Set<DistributorNumericColumnKey>();
  const order: DistributorNumericColumnKey[] = [];
  for (const anchor of header.anchors) {
    const key = anchor.key as DistributorNumericColumnKey;
    if (!TEMPLATE_NUMERIC_KEYS.has(key) || seen.has(key)) continue;
    seen.add(key);
    order.push(key);
  }
  return order;
}

function sanitizeNumericOrder(order: DistributorNumericColumnKey[]): DistributorNumericColumnKey[] {
  const seen = new Set<DistributorNumericColumnKey>();
  const safeOrder = order.filter((key) => {
    if (!TEMPLATE_NUMERIC_KEYS.has(key) || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  return safeOrder.length ? safeOrder : [...FALLBACK_NUMERIC_ORDER];
}

function findDistributorProfile(
  profiles: DistributorFormatProfile[],
  distributorName: string,
): DistributorFormatProfile | undefined {
  const normalizedName = normalizeProfileName(distributorName);
  return profiles.find(
    (profile) => profile.active && normalizeProfileName(profile.distributorName) === normalizedName,
  );
}

function normalizeProfileName(value: string): string {
  return cleanText(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function findAliasX(lines: PdfLine[], aliases: string[]): number {
  const normalizedAliases = aliases.map(normalizeHeader);
  for (const line of lines) {
    for (const item of line.items) {
      const text = normalizeHeader(item.text);
      if (normalizedAliases.some((alias) => alias === text || text.includes(alias))) return item.x;
    }
  }
  return 0;
}

function isAminTradersFormat(
  rawText: string,
  distributorName: string,
  profile?: DistributorFormatProfile,
): boolean {
  const normalized = normalizeProfileName(`${distributorName} ${rawText}`);
  return (
    normalizeProfileName(profile?.distributorName ?? "").includes("amin traders charsadda") ||
    (normalized.includes("amin traders") &&
      normalized.includes("charsadda") &&
      normalized.includes("sale and stock report"))
  );
}

function normalizeAminDistributorName(distributorName: string): string {
  return /amin\s+traders/i.test(distributorName) ? "AMIN TRADERS - CHARSADDA" : distributorName;
}

function parseAminTradersPdf({
  file,
  lines,
  distributorName,
  fromDate,
  toDate,
  profileName,
  profileStatus,
  numericOrder,
  warnings,
}: {
  file: File;
  lines: PdfLine[];
  distributorName: string;
  fromDate: string;
  toDate: string;
  profileName: string;
  profileStatus: DistributorProfileStatus;
  numericOrder: DistributorNumericColumnKey[];
  warnings: string[];
}): {
  rows: DistributorSalesRow[];
  summary: DistributorFileSummary;
} {
  const rows: DistributorSalesRow[] = [];
  const sortedLines = [...lines].sort((a, b) => a.order - b.order);
  let currentGroup = "Ungrouped";
  let lastProductRow: DistributorSalesRow | null = null;
  let pendingProductLine = "";

  warnings.push(
    "AMIN TRADERS format detected: Pack, Purchase Return, Adjustment, and Today Return are not mapped to standard columns.",
  );

  for (const line of sortedLines) {
    const text = cleanText(line.text);
    if (!text || FOOTER_LINE.test(text)) continue;
    if (isAminHeaderOrReportLine(text)) continue;

    const groupName = detectAminGroupHeading(text);
    if (groupName) {
      currentGroup = groupName;
      lastProductRow = null;
      pendingProductLine = "";
      continue;
    }

    if (isGroupTotalLine(text)) {
      const totalRow = rowFromAminTotalLine(text, {
        distributorName,
        fromDate,
        toDate,
        sourcePdfFileName: file.name,
        groupName: currentGroup,
      });
      if (totalRow) rows.push(totalRow);
      lastProductRow = null;
      pendingProductLine = "";
      continue;
    }

    if (startsWithAminProductCode(text)) {
      const combinedText = pendingProductLine ? `${pendingProductLine} ${text}` : text;
      const parsed = parseAminProductLine(combinedText);
      if (!parsed) {
        pendingProductLine = combinedText;
        lastProductRow = null;
        continue;
      }
      const row = rowFromAminLine(parsed, {
        distributorName,
        fromDate,
        toDate,
        sourcePdfFileName: file.name,
        groupName: currentGroup,
      });
      rows.push(row);
      lastProductRow = row;
      pendingProductLine = "";
      continue;
    }

    if (pendingProductLine) {
      const combinedText = `${pendingProductLine} ${text}`;
      const parsed = parseAminProductLine(combinedText);
      if (parsed) {
        const row = rowFromAminLine(parsed, {
          distributorName,
          fromDate,
          toDate,
          sourcePdfFileName: file.name,
          groupName: currentGroup,
        });
        rows.push(row);
        lastProductRow = row;
        pendingProductLine = "";
      } else {
        pendingProductLine = combinedText;
      }
      continue;
    }

    if (lastProductRow && isAminProductNameContinuation(text)) {
      lastProductRow.productName = cleanText(`${lastProductRow.productName} ${text}`);
    }
  }

  if (pendingProductLine) {
    warnings.push(`Unclear AMIN product row was skipped: ${pendingProductLine}`);
  }

  const groupSummaries = summarizeGroups(rows);
  if (!rows.some((row) => row.rowType === "product")) {
    warnings.push("No AMIN product rows were extracted from the detected table.");
  }
  for (const group of groupSummaries) {
    if (group.totalSource === "calculated" && group.productCount > 0) {
      warnings.push(
        `Group total was not clearly detected for ${group.name}; it will be calculated from extracted rows.`,
      );
    }
  }

  const mappedColumns = [
    "Product Code",
    "Product Name",
    "Trade Price",
    "Open Stock",
    "Receipt Qty",
    "Total Stock",
    "Net Sale Qty",
    "Net Sale Bns",
    "Sale Value",
    "Closing Stock",
    "Stock Value",
    "Today Sale",
  ];

  return {
    rows,
    summary: {
      fileName: file.name,
      distributorName,
      fromDate,
      toDate,
      rowsExtracted: rows.filter((row) => row.rowType === "product").length,
      mappedColumns,
      profileName,
      profileStatus,
      suggestedNumericOrder: numericOrder,
      sourceHeaders: BUILT_IN_DISTRIBUTOR_PROFILES[0].sourceHeaders,
      groups: groupSummaries,
      warnings,
    },
  };
}

function isAminHeaderOrReportLine(text: string): boolean {
  return /\b(item description|opening balance|purchase|closing|company|from date|to date|sale and stock report|rate|pack)\b/i.test(
    text,
  );
}

function detectAminGroupHeading(text: string): string | null {
  const clean = cleanText(text);
  const labelled = clean.match(/(?:^|\b)group\s*[:.-]?\s*(.+)$/i);
  if (labelled) return cleanText(labelled[1]);
  if (/^\d{4,6}\s*\/\s*[A-Z0-9][A-Z0-9\s()./-]+$/i.test(clean)) return clean;
  return detectGroupHeading(clean);
}

function startsWithAminProductCode(text: string): boolean {
  return /^\s*\d{4,8}\s+/.test(text);
}

function isAminProductNameContinuation(text: string): boolean {
  const clean = cleanText(text);
  if (!clean || startsWithAminProductCode(clean)) return false;
  if (isGroupTotalLine(clean) || detectAminGroupHeading(clean) || isAminHeaderOrReportLine(clean)) {
    return false;
  }
  if (/-?\d[\d,]*(?:\.\d+)?/.test(clean)) return false;
  return /[A-Za-z]/.test(clean) && clean.length <= 80;
}

function parseAminProductLine(text: string): AminParsedLine | null {
  const clean = cleanText(text);
  const codeMatch = clean.match(/^(\d{4,8})\s+(.+)$/);
  if (!codeMatch) return null;

  const productCode = codeMatch[1];
  const rest = codeMatch[2];
  const rateMatch = rest.match(/(?:^|\s)(-?\d[\d,]*\.\d+)(?=\s|$)/);
  if (!rateMatch || rateMatch.index === undefined) return null;

  const productName = cleanText(rest.slice(0, rateMatch.index));
  const valueText = rest.slice(rateMatch.index).trim();
  const sourceValues = tokenizeAminValues(valueText);
  if (sourceValues.length < 3) return null;

  const ignoredWarnings: string[] = [];
  const pack = sourceValues[1];
  if (pack !== "") ignoredWarnings.push(`Pack ignored: ${pack}`);
  if (sourceValues[4] !== "")
    ignoredWarnings.push(`Purchase Return not mapped: ${sourceValues[4]}`);
  if (sourceValues[9] !== "") ignoredWarnings.push(`Adjustment not mapped: ${sourceValues[9]}`);
  if (sourceValues[13] !== "") ignoredWarnings.push(`Today Return not mapped: ${sourceValues[13]}`);

  return {
    productCode,
    productName,
    values: [
      valueToNumber(sourceValues[0]),
      valueToNumber(sourceValues[2]),
      valueToNumber(sourceValues[3]),
      valueToNumber(sourceValues[5]),
      valueToNumber(sourceValues[6]),
      valueToNumber(sourceValues[7]),
      valueToNumber(sourceValues[8]),
      valueToNumber(sourceValues[10]),
      valueToNumber(sourceValues[11]),
      valueToNumber(sourceValues[12]),
    ],
    ignoredWarnings,
  };
}

function tokenizeAminValues(text: string): string[] {
  return [...text.matchAll(/-?\d[\d,]*(?:\.\d+)?|[-–—]|(?:\d+)\s*`?s/gi)].map((match) =>
    cleanText(match[0]),
  );
}

function valueToNumber(value: string | undefined): number | "" {
  if (!value || /^[-–—]$/.test(value)) return "";
  const clean = value.replace(/,/g, "").replace(/`?s$/i, "");
  const parsed = Number(clean);
  return Number.isFinite(parsed) ? parsed : "";
}

function rowFromAminLine(
  parsed: AminParsedLine,
  meta: Pick<
    DistributorSalesRow,
    "distributorName" | "fromDate" | "toDate" | "sourcePdfFileName" | "groupName"
  >,
): DistributorSalesRow {
  const row = blankRow(meta);
  row.rowType = "product";
  row.productCode = parsed.productCode;
  row.productName = parsed.productName;
  applyParsedValues(row, parsed.values, AMIN_TRADERS_NUMERIC_ORDER);
  row.remarksWarnings = parsed.ignoredWarnings.join("; ");
  return row;
}

function rowFromAminTotalLine(
  text: string,
  meta: Pick<
    DistributorSalesRow,
    "distributorName" | "fromDate" | "toDate" | "sourcePdfFileName" | "groupName"
  >,
): DistributorSalesRow | null {
  const values = tokenizeAminValues(text).map(valueToNumber);
  if (!values.some((value) => typeof value === "number")) return null;
  const row = blankRow(meta);
  row.rowType = "groupTotal";
  row.productName = "Group Total";
  applyParsedValues(row, values, AMIN_TRADERS_NUMERIC_ORDER);
  row.remarksWarnings = "Group total detected from PDF.";
  return row;
}

function rowFromLine(
  line: PdfLine,
  header: HeaderDetection,
  meta: Pick<
    DistributorSalesRow,
    "distributorName" | "fromDate" | "toDate" | "sourcePdfFileName" | "groupName"
  >,
  numericOrder: DistributorNumericColumnKey[],
): DistributorSalesRow | null {
  const values = valuesByAnchors(line, header.anchors);
  const parsed = parseProductLine(line.text);
  const remarks: string[] = [];
  const base = blankRow(meta);

  for (const column of DATA_COLUMNS) {
    const value = values.get(column.key) ?? "";
    if (column.numeric) {
      base[column.key] = parseNumber(value) as never;
    } else {
      base[column.key] = cleanText(value) as never;
    }
  }

  if (parsed) {
    base.productCode = parsed.productCode;
    base.productName = parsed.productName;
    applyParsedNumbers(base, parsed.numbers, numericOrder);
  }

  if (!base.productCode || !base.productName) return null;
  base.rowType = "product";

  for (const column of DATA_COLUMNS) {
    if (!header.columnMap.has(column.key)) remarks.push(`Missing mapping: ${column.header}`);
  }
  base.remarksWarnings = remarks.join("; ");
  return base;
}

function valuesByAnchors(
  line: PdfLine,
  anchors: Array<{ key: keyof DistributorSalesRow; x: number }>,
): Map<keyof DistributorSalesRow, string> {
  const values = new Map<keyof DistributorSalesRow, string>();
  if (!anchors.length) return values;
  const sorted = anchors.filter((anchor) => anchor.x > 0).sort((a, b) => a.x - b.x);
  if (!sorted.length) return values;

  for (const item of line.items) {
    let target = sorted[0];
    for (let index = 0; index < sorted.length; index++) {
      const current = sorted[index];
      const next = sorted[index + 1];
      const end = next ? (current.x + next.x) / 2 : Number.POSITIVE_INFINITY;
      if (item.x >= current.x - 8 && item.x < end) {
        target = current;
        break;
      }
    }
    values.set(target.key, [values.get(target.key), item.text].filter(Boolean).join(" "));
  }

  return values;
}

function parseProductLine(text: string): ParsedProductLine | null {
  const clean = cleanText(text);
  const codeMatch = clean.match(/^([A-Z0-9][A-Z0-9./-]{1,})\s+(.+)$/i);
  if (!codeMatch) return null;

  const productCode = codeMatch[1];
  const rest = codeMatch[2];
  const numericMatches = [...rest.matchAll(/-?\d[\d,]*(?:\.\d+)?%?/g)];
  const tradePriceMatch =
    numericMatches.find((match) => {
      const value = match[0];
      const index = match.index ?? 0;
      const previous = rest[index - 1] ?? " ";
      const next = rest[index + value.length] ?? " ";
      return value.includes(".") && !/[A-Za-z]/.test(previous) && !/[A-Za-z]/.test(next);
    }) ?? numericMatches.find((match) => (match.index ?? 0) > 0);
  if (!tradePriceMatch) return null;

  const tradePriceIndex = tradePriceMatch.index ?? rest.length;
  const productName = cleanText(rest.slice(0, tradePriceIndex));
  if (!productName || productName.length < 2) return null;

  const numbers = [...rest.slice(tradePriceIndex).matchAll(/-?\d[\d,]*(?:\.\d+)?%?/g)]
    .map((match) => Number(match[0].replace(/,/g, "").replace("%", "")))
    .filter((value) => Number.isFinite(value));

  return { productCode, productName, numbers };
}

function applyParsedNumbers(
  row: DistributorSalesRow,
  numbers: number[],
  numericOrder: DistributorNumericColumnKey[],
) {
  applyParsedValues(row, numbers, numericOrder);
}

function applyParsedValues(
  row: DistributorSalesRow,
  values: Array<number | "">,
  numericOrder: DistributorNumericColumnKey[],
) {
  numericOrder.forEach((key, index) => {
    if (values[index] === undefined) return;
    row[key] = values[index] as never;
  });
}

function blankRow(
  meta: Pick<
    DistributorSalesRow,
    "distributorName" | "fromDate" | "toDate" | "sourcePdfFileName" | "groupName"
  >,
): DistributorSalesRow {
  return {
    rowType: "product",
    groupName: meta.groupName,
    distributorName: meta.distributorName,
    fromDate: meta.fromDate,
    toDate: meta.toDate,
    productCode: "",
    productName: "",
    tradePrice: "",
    openingQty: "",
    openingValue: "",
    purchaseQty: "",
    purchaseBonus: "",
    purchaseValue: "",
    totalStock: "",
    salesQty: "",
    salesBonus: "",
    returnQty: "",
    returnBonus: "",
    netSaleQty: "",
    netSaleBonus: "",
    netSaleValue: "",
    transferIn: "",
    transferOut: "",
    closingQty: "",
    closingBonus: "",
    closingValue: "",
    todaySales: "",
    previousMonthSalesQty: "",
    previousMonthSalesValue: "",
    varianceQty: "",
    variancePercent: "",
    sourcePdfFileName: meta.sourcePdfFileName,
    remarksWarnings: "",
  };
}

function isLikelyProductLine(text: string): boolean {
  const clean = text.trim();
  if (!clean || clean.length < 8) return false;
  if (TOTAL_LINE.test(clean) || FOOTER_LINE.test(clean)) return false;
  if (!/\d/.test(clean)) return false;
  if (!/[A-Za-z]/.test(clean)) return false;
  return /^\s*[A-Z0-9][A-Z0-9./-]{2,}\s+/i.test(clean) || /\s{2,}/.test(clean);
}

function isProductNameContinuation(text: string): boolean {
  const clean = cleanText(text);
  if (
    !clean ||
    isLikelyProductLine(clean) ||
    isGroupTotalLine(clean) ||
    detectGroupHeading(clean)
  ) {
    return false;
  }
  if (/\b(company|division|page|date|from|to|product|description|opening|closing)\b/i.test(clean)) {
    return false;
  }
  if (/-?\d[\d,]*(?:\.\d+)?/.test(clean)) return false;
  return /[A-Za-z]/.test(clean) && clean.length <= 80;
}

function isGroupTotalLine(text: string): boolean {
  return /\bgroup\s*total\b/i.test(text) || /^\s*total\s*:/i.test(text);
}

function detectGroupHeading(text: string): string | null {
  const clean = cleanText(text);
  if (!clean || clean.length < 4) return null;
  if (TOTAL_LINE.test(clean) || FOOTER_LINE.test(clean)) return null;
  if (isLikelyProductLine(clean)) return null;
  if (/\b(company|division|page|date|from|to|product|description|opening|closing)\b/i.test(clean)) {
    return null;
  }
  if (/\b(group|grp)\b/i.test(clean)) return clean;
  const mostlyCaps = clean === clean.toUpperCase() && /[A-Z]/.test(clean);
  if (mostlyCaps && clean.split(/\s+/).length <= 8 && !/\d{4,}/.test(clean)) return clean;
  return null;
}

function rowFromTotalLine(
  text: string,
  meta: Pick<
    DistributorSalesRow,
    "distributorName" | "fromDate" | "toDate" | "sourcePdfFileName" | "groupName"
  >,
  numericOrder: DistributorNumericColumnKey[],
): DistributorSalesRow | null {
  const numbers = [...text.matchAll(/-?\d[\d,]*(?:\.\d+)?%?/g)]
    .map((match) => Number(match[0].replace(/,/g, "").replace("%", "")))
    .filter((value) => Number.isFinite(value));
  if (!numbers.length) return null;
  const row = blankRow(meta);
  row.rowType = "groupTotal";
  row.productName = "Group Total";
  applyParsedNumbers(row, numbers, numericOrder);
  row.remarksWarnings = "Group total detected from PDF.";
  return row;
}

function rowsFromTextFallback(
  lines: PdfLine[],
  meta: Pick<
    DistributorSalesRow,
    "distributorName" | "fromDate" | "toDate" | "sourcePdfFileName" | "groupName"
  >,
): DistributorSalesRow[] {
  const rows: DistributorSalesRow[] = [];
  const sortedLines = [...lines].sort((a, b) => a.order - b.order);
  let currentGroup = meta.groupName || "Ungrouped";
  let lastProductRow: DistributorSalesRow | null = null;

  for (const line of sortedLines) {
    const text = cleanText(line.text);
    if (!text || FOOTER_LINE.test(text)) continue;
    if (isGroupTotalLine(text)) {
      const totalRow = rowFromTotalLine(
        text,
        { ...meta, groupName: currentGroup },
        FALLBACK_NUMERIC_ORDER,
      );
      if (totalRow) rows.push(totalRow);
      lastProductRow = null;
      continue;
    }

    const groupName = detectGroupHeading(text);
    if (groupName) {
      currentGroup = groupName;
      lastProductRow = null;
      continue;
    }

    const parsed = parseProductLine(text);
    if (parsed) {
      const row = blankRow({ ...meta, groupName: currentGroup });
      row.rowType = "product";
      row.productCode = parsed.productCode;
      row.productName = parsed.productName;
      applyParsedNumbers(row, parsed.numbers, FALLBACK_NUMERIC_ORDER);
      row.remarksWarnings = "Parsed with text-row fallback.";
      rows.push(row);
      lastProductRow = row;
      continue;
    }

    if (lastProductRow && isProductNameContinuation(text)) {
      lastProductRow.productName = cleanText(`${lastProductRow.productName} ${text}`);
    }
  }

  return rows;
}

function summarizeGroups(rows: DistributorSalesRow[]): DistributorGroupSummary[] {
  const groups = new Map<string, DistributorGroupSummary>();
  for (const row of rows) {
    const name = row.groupName || "Ungrouped";
    const current =
      groups.get(name) ??
      ({
        name,
        productCount: 0,
        totalSource: "missing",
      } satisfies DistributorGroupSummary);
    if (row.rowType === "product") current.productCount += 1;
    if (row.rowType === "groupTotal") current.totalSource = "pdf";
    groups.set(name, current);
  }
  return [...groups.values()].map((group) => ({
    ...group,
    totalSource: group.totalSource === "missing" ? "calculated" : group.totalSource,
  }));
}

function detectDistributorName(text: string, fileName: string): string {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(0, 12);
  const labelled = lines
    .join(" ")
    .match(/\b(?:distributor|stockist|customer|party)\s*(?:name)?\s*[:-]\s*([^|,]{3,60})/i);
  if (labelled) return cleanText(labelled[1]);
  return cleanText(
    lines.find((line) => /enterprise|trader|distributor|pharma|medical|care/i.test(line)) ??
      distributorNameFromFile(fileName),
  );
}

function distributorNameFromFile(fileName: string): string {
  return cleanText(fileName.replace(/\.pdf$/i, "").replace(/\([^)]*\)/g, " "));
}

function detectDateRange(text: string): { fromDate: string; toDate: string } {
  const compact = text.replace(/\s+/g, " ");
  const range = compact.match(
    /\b(?:from\s*date|from|period)\s*[:-]?\s*([0-9]{1,2}[-/ ][A-Za-z]{3,9}[-/ ][0-9]{2,4}|[0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{2,4})\s*(?:to\s*date|to|till|-)\s*[:-]?\s*([0-9]{1,2}[-/ ][A-Za-z]{3,9}[-/ ][0-9]{2,4}|[0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{2,4})/i,
  );
  if (range) return { fromDate: range[1], toDate: range[2] };
  return { fromDate: "", toDate: "" };
}

function parseNumber(value: string): number | "" {
  const clean = value.replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);
  return clean ? Number(clean[0]) : "";
}

function cleanText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function normalizeHeader(value: string): string {
  return cleanText(value)
    .toLowerCase()
    .replace(/[./()]/g, "")
    .replace(/&/g, "and");
}

function columnHeader(key: keyof DistributorSalesRow | string): string {
  return STANDARD_COLUMNS.find((column) => column.key === key)?.header ?? String(key);
}

export async function exportDistributorSalesExcel(rows: DistributorSalesRow[]): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  const groups = new Map<string, DistributorSalesRow[]>();
  for (const row of rows) {
    const key = row.distributorName || "Distributor Sales";
    groups.set(key, [...(groups.get(key) ?? []), row]);
  }

  for (const [distributorName, distributorRows] of groups.entries()) {
    const ws = wb.addWorksheet(safeSheetName(distributorName));
    buildTemplateSheet(ws, distributorName, distributorRows);
  }

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function buildTemplateSheet(
  ws: ExcelJS.Worksheet,
  distributorName: string,
  rows: DistributorSalesRow[],
) {
  const totalColumns = 23;
  ws.columns = [{ width: 6 }, ...TEMPLATE_COLUMNS.map((column) => ({ width: column.width }))];

  ws.mergeCells(1, 1, 1, totalColumns);
  ws.getCell(1, 1).value = distributorName.toUpperCase();
  ws.getCell(1, 1).font = { bold: true, size: 11 };
  ws.getCell(1, 1).alignment = centerAlignment();

  ws.mergeCells(2, 1, 2, 5);
  ws.getCell(2, 1).value = `FROM DATE: ${rows[0]?.fromDate ?? ""}`;
  ws.getCell(2, 1).font = { bold: true };
  ws.getCell(2, 1).alignment = centerAlignment();

  ws.mergeCells(2, 18, 2, 23);
  ws.getCell(2, 18).value = `TO DATE: ${rows[0]?.toDate ?? ""}`;
  ws.getCell(2, 18).font = { bold: true };
  ws.getCell(2, 18).alignment = centerAlignment();

  const headerRow = ws.getRow(3);
  const subHeaderRow = ws.getRow(4);
  ws.mergeCells(3, 1, 4, 1);
  ws.mergeCells(3, 2, 4, 2);
  ws.mergeCells(3, 3, 4, 3);
  ws.mergeCells(3, 4, 4, 4);
  ws.mergeCells(3, 5, 4, 5);
  ws.mergeCells(3, 6, 3, 7);
  ws.mergeCells(3, 8, 4, 8);
  ws.mergeCells(3, 9, 3, 10);
  ws.mergeCells(3, 11, 3, 12);
  ws.mergeCells(3, 13, 3, 14);
  ws.mergeCells(3, 15, 4, 15);
  ws.mergeCells(3, 16, 4, 16);
  ws.mergeCells(3, 17, 4, 17);
  ws.mergeCells(3, 18, 4, 18);
  ws.mergeCells(3, 19, 4, 19);
  ws.mergeCells(3, 20, 4, 20);
  ws.mergeCells(3, 21, 4, 21);
  ws.mergeCells(3, 22, 3, 23);

  const topHeaders: Array<[number, string]> = [
    [1, "Sr#"],
    [2, "Product code"],
    [3, "Product name"],
    [4, "Trade Price"],
    [5, "Open Stock"],
    [6, "Receipt"],
    [8, "Total Stock"],
    [9, "Sales"],
    [11, "Return"],
    [13, "Net.Sale"],
    [15, "Sale Value"],
    [16, "Transfer In"],
    [17, "Transfer Out"],
    [18, "Closing Stock"],
    [19, "Stock Value"],
    [20, "Today Sale"],
    [21, "Previous Month"],
    [22, "Variance"],
  ];
  for (const [col, value] of topHeaders) headerRow.getCell(col).value = value;
  const subHeaders: Array<[number, string]> = [
    [6, "Qty"],
    [7, "Bns"],
    [9, "Qty"],
    [10, "Bns"],
    [11, "Qty"],
    [12, "Bns"],
    [13, "Qty"],
    [14, "Bns"],
    [22, "Qty"],
    [23, "Per%"],
  ];
  for (const [col, value] of subHeaders) subHeaderRow.getCell(col).value = value;

  for (let rowNumber = 1; rowNumber <= 4; rowNumber++) {
    const row = ws.getRow(rowNumber);
    row.height = rowNumber <= 2 ? 20 : 22;
    for (let col = 1; col <= totalColumns; col++) {
      const cell = row.getCell(col);
      cell.border = thinBorder("FF000000");
      cell.alignment = centerAlignment();
      if (rowNumber >= 3) cell.font = { bold: true };
    }
  }

  let outputRowNumber = 5;
  let serial = 1;
  for (const [groupName, groupRows] of groupRowsForExport(rows).entries()) {
    const headingRow = ws.getRow(outputRowNumber++);
    ws.mergeCells(headingRow.number, 1, headingRow.number, totalColumns);
    headingRow.getCell(1).value = `Group: ${groupName}`;
    headingRow.getCell(1).font = { bold: true };
    headingRow.getCell(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF99" },
    };
    headingRow.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
    for (let col = 1; col <= totalColumns; col++) {
      headingRow.getCell(col).border = thinBorder("FF000000");
    }

    const productRows = groupRows.filter((row) => row.rowType === "product");
    for (const item of productRows) {
      writeTemplateDataRow(ws.getRow(outputRowNumber++), item, serial++);
    }

    const detectedTotal = groupRows.find((row) => row.rowType === "groupTotal");
    const totalRow = detectedTotal ?? calculateGroupTotal(groupName, productRows, rows[0]);
    writeTemplateDataRow(ws.getRow(outputRowNumber++), totalRow, "");
    const excelTotalRow = ws.getRow(outputRowNumber - 1);
    excelTotalRow.getCell(2).value = "";
    excelTotalRow.getCell(3).value = "Group Total";
    excelTotalRow.font = { bold: true };
    excelTotalRow.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFEFEF" } };
    });
  }

  ws.views = [{ state: "frozen", ySplit: 4 }];
  ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
}

function writeTemplateDataRow(row: ExcelJS.Row, item: DistributorSalesRow, serial: number | "") {
  const values = [
    serial,
    item.productCode,
    item.productName,
    item.tradePrice,
    item.openingQty,
    item.purchaseQty,
    item.purchaseBonus,
    item.totalStock,
    item.salesQty,
    item.salesBonus,
    item.returnQty,
    item.returnBonus,
    item.netSaleQty,
    item.netSaleBonus,
    item.netSaleValue,
    item.transferIn,
    item.transferOut,
    item.closingQty,
    item.closingValue,
    item.todaySales,
    item.previousMonthSalesQty,
    item.varianceQty,
    item.variancePercent,
  ];
  values.forEach((value, colIndex) => {
    const cell = row.getCell(colIndex + 1);
    cell.value = value === "" ? null : value;
    cell.border = thinBorder("FF000000");
    cell.alignment =
      colIndex === 2
        ? { horizontal: "left", vertical: "middle", wrapText: true }
        : centerAlignment();
    if (typeof cell.value === "number" && colIndex > 2) cell.numFmt = "#,##0.00";
  });
}

function groupRowsForExport(rows: DistributorSalesRow[]): Map<string, DistributorSalesRow[]> {
  const groups = new Map<string, DistributorSalesRow[]>();
  for (const row of rows) {
    const key = row.groupName || "Ungrouped";
    groups.set(key, [...(groups.get(key) ?? []), row]);
  }
  return groups;
}

function calculateGroupTotal(
  groupName: string,
  rows: DistributorSalesRow[],
  sample?: DistributorSalesRow,
): DistributorSalesRow {
  const total = blankRow({
    distributorName: sample?.distributorName ?? "",
    fromDate: sample?.fromDate ?? "",
    toDate: sample?.toDate ?? "",
    sourcePdfFileName: sample?.sourcePdfFileName ?? "",
    groupName,
  });
  total.rowType = "groupTotal";
  total.productName = "Group Total";
  total.remarksWarnings = "Group total calculated from extracted product rows.";
  for (const column of TEMPLATE_COLUMNS) {
    if (!column.numeric) continue;
    const sum = rows.reduce((acc, row) => {
      const value = row[column.key];
      return acc + (typeof value === "number" ? value : 0);
    }, 0);
    total[column.key] = sum as never;
  }
  return total;
}

function centerAlignment(): Partial<ExcelJS.Alignment> {
  return { horizontal: "center", vertical: "middle", wrapText: true };
}

function safeSheetName(name: string): string {
  const clean = cleanText(name)
    .replace(/[:\\/?*[\]]/g, " ")
    .slice(0, 31);
  return clean || "Distributor Sales";
}

function thinBorder(color = "FF9CA3AF"): Partial<ExcelJS.Borders> {
  return {
    top: { style: "thin", color: { argb: color } },
    left: { style: "thin", color: { argb: color } },
    bottom: { style: "thin", color: { argb: color } },
    right: { style: "thin", color: { argb: color } },
  };
}
