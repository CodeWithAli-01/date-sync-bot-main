// Client-side PDF text extraction + employee-row parsing.
// Strong mode treats every numeric employee code as the start of a row.

export interface PdfRow {
  code: string | null;
  name: string;
  planned: number;
  unplanned: number;
  count: number;
  selfieText: string;
  total: number;
}

export interface PdfParseResult {
  date: string;
  day: number;
  fileName: string;
  fileHash: string;
  rows: PdfRow[];
  rawText: string;
  stats: {
    detectedRows: number;
    parsedRows: number;
    skippedRows: number;
  };
}

type PdfTextItem = {
  str?: string;
  transform?: number[];
};

type PdfJs = typeof import("pdfjs-dist");

type PdfLine = {
  text: string;
  items: { x: number; str: string }[];
};

type PositionedPdfItem = {
  x: number;
  y: number;
  str: string;
};

type DetectedPdfRow = {
  code: string;
  text: string;
  items: { x: number; str: string }[];
};

function extractDateFromFilename(fileName: string): string | null {
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
  const day = named[1].padStart(2, "0");
  const month = months[named[2].slice(0, 3).toLowerCase()];
  return month ? `${named[3]}-${month}-${day}` : null;
}

const MONTHS: Record<string, string> = {
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

function toIsoDate(
  year: string | number,
  month: string | number,
  day: string | number,
): string | null {
  const y = Number(year);
  const m = Number(month);
  const d = Number(day);
  if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return null;
  if (y < 2000 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) return null;
  const date = new Date(Date.UTC(y, m - 1, d));
  if (date.getUTCFullYear() !== y || date.getUTCMonth() !== m - 1 || date.getUTCDate() !== d) {
    return null;
  }
  return `${String(y).padStart(4, "0")}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

function extractDateFromPdfText(
  text: string,
  defaultYear = new Date().getFullYear(),
): string | null {
  const compact = text.replace(/\s+/g, " ").trim();
  const firstPageText = compact.slice(0, 4000);
  const labelled = firstPageText.match(
    /\b(?:report\s+date|date)\s*[:-]?\s*(\d{1,2})(?:st|nd|rd|th)?\s*(?:-|\/|\.|\s)\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s*(?:-|\/|,|\s)?\s*(20\d{2})\b/i,
  );
  if (labelled) {
    const month = MONTHS[labelled[2].slice(0, 3).toLowerCase()];
    const iso = month ? toIsoDate(labelled[3], month, labelled[1]) : null;
    if (iso) return iso;
  }

  const labelledNumeric = firstPageText.match(
    /\b(?:report\s+date|date)\s*[:-]?\s*(\d{1,2})[/.-](\d{1,2})[/.-](20\d{2})\b/i,
  );
  if (labelledNumeric) {
    const iso = toIsoDate(labelledNumeric[3], labelledNumeric[2], labelledNumeric[1]);
    if (iso) return iso;
  }

  const labelledShort = firstPageText.match(
    /\b(?:report\s+date|date)\s*[:-]?\s*(\d{1,2})(?:st|nd|rd|th)?\s*(?:-|\/|\.|\s)\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b/i,
  );
  if (labelledShort) {
    const month = MONTHS[labelledShort[2].slice(0, 3).toLowerCase()];
    const iso = month ? toIsoDate(defaultYear, month, labelledShort[1]) : null;
    if (iso) return iso;
  }

  const named = firstPageText.match(
    /\b(\d{1,2})(?:st|nd|rd|th)?\s*(?:-|\/|\.|\s)\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s*(?:-|\/|,|\s)?\s*(20\d{2})\b/i,
  );
  if (named) {
    const month = MONTHS[named[2].slice(0, 3).toLowerCase()];
    const iso = month ? toIsoDate(named[3], month, named[1]) : null;
    if (iso) return iso;
  }

  const monthFirst = firstPageText.match(
    /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(\d{1,2})(?:st|nd|rd|th)?(?:,|\s)+\s*(20\d{2})\b/i,
  );
  if (monthFirst) {
    const month = MONTHS[monthFirst[1].slice(0, 3).toLowerCase()];
    const iso = month ? toIsoDate(monthFirst[3], month, monthFirst[2]) : null;
    if (iso) return iso;
  }

  const namedShort = firstPageText.match(
    /\b(\d{1,2})(?:st|nd|rd|th)?\s*(?:-|\/|\.|\s)\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b/i,
  );
  if (namedShort) {
    const month = MONTHS[namedShort[2].slice(0, 3).toLowerCase()];
    const iso = month ? toIsoDate(defaultYear, month, namedShort[1]) : null;
    if (iso) return iso;
  }

  const numeric = firstPageText.match(/\b(\d{1,2})[/.-](\d{1,2})[/.-](20\d{2})\b/);
  return numeric ? toIsoDate(numeric[3], numeric[2], numeric[1]) : null;
}

function extractSelfieCount(text: string): number | null {
  const lower = text.toLowerCase();
  const selfie = lower.match(/(\d+)\s*selfies?\b/);
  if (selfie) return parseInt(selfie[1], 10);

  return null;
}

function extractSelfieText(text: string, count: number): string {
  const cleaned = cleanSelfieNoise(text);

  const selfieStart = cleaned.search(/\b\d{1,3}\s*selfies?\b/i);
  if (selfieStart >= 0) {
    const remark = cleaned.slice(selfieStart);
    const withGrp = remark.match(/^(\d{1,3}\s*selfies?\b.*?\bgrp\b)/i);
    if (withGrp) return normalizeSelfieText(withGrp[1], count);

    const withLocation = remark.match(
      /^(\d{1,3}\s*selfies?\b(?:\s+(?:with\s+)?\d{0,3}\s*locations?\b|\s+with\s+locations?\b)?)/i,
    );
    if (withLocation) return normalizeSelfieText(withLocation[1], count);
  }

  const specialRemark = extractSpecialRemark(cleaned);
  if (specialRemark && count === 0) return specialRemark;
  if (specialRemark && count > 0) return `${count} selfies with locations in grp`;

  return "";
}

function normalizeSelfieText(text: string, count: number): string {
  const normalized = cleanSelfieNoise(text)
    .replace(/[^\p{L}\p{N}\s]/gu, " ")
    .replace(
      /\bwith\s+(?:PSV|MIO|ASM|RSM|SM|AM|FM|TM)(?:\s+(?:PSV|MIO|ASM|RSM|SM|AM|FM|TM))*\s+locations?\b/gi,
      "with locations",
    )
    .replace(/\bwith\s+(?:PSV|MIO|ASM|RSM|SM|AM|FM|TM)\b.*$/i, "with locations in grp")
    .replace(/\b(?:PSV|MIO|ASM|RSM|SM|AM|FM|TM)\b.*$/i, "")
    .replace(/\blocations?\s+[A-Z][A-Z\s.'-]{3,}$/g, "locations")
    .replace(/\bgrp\s+[A-Z][A-Z\s.'-]{3,}$/g, "grp")
    .replace(/\s+/g, " ")
    .trim();

  if (!normalized || !/\bselfies?\b/i.test(normalized)) {
    return "";
  }

  return normalized;
}

function cleanSelfieNoise(text: string): string {
  return text
    .replace(/\b\d{1,2}:\d{2}(?:\s*[AP]M)?\b/gi, " ")
    .replace(/\b\d{4,}\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractSpecialRemark(text: string): string | null {
  const lower = text.toLowerCase();
  const special = lower.match(/\b(meeting|new joining|resigned)\b.*$/i);
  if (!special) return null;

  return special[0]
    .replace(/\b\d{1,3}\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractTotalValue(
  text: string,
  items: { x: number; str: string }[],
  totalX: number | null,
): number {
  if (totalX != null) {
    const numericItems = items
      .filter((item) => /^\d+$/.test(item.str.trim()))
      .map((item) => ({
        value: parseInt(item.str.trim(), 10),
        distance: Math.abs(item.x - totalX),
      }))
      .filter((item) => item.value <= 50)
      .sort((a, b) => a.distance - b.distance);

    const nearest = numericItems[0];
    if (nearest && nearest.distance <= 55) return nearest.value;
  }

  const strictTotal = extractTotalFromNumericColumns(text);
  return strictTotal ?? 0;
}

function extractTotalFromNumericColumns(text: string): number | null {
  const cleaned = text
    .replace(/\b\d{1,2}:\d{2}(?:\s*[AP]M)?\b/gi, " ")
    .replace(/\b\d{4,6}\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const tokens = cleaned.split(" ").filter(Boolean);

  for (let i = 0; i <= tokens.length - 3; i++) {
    if (!/^\d{1,2}$/.test(tokens[i])) continue;
    if (!/^\d{1,2}$/.test(tokens[i + 1])) continue;
    if (!/^\d{1,2}$/.test(tokens[i + 2])) continue;

    const total = parseInt(tokens[i + 2], 10);
    if (total <= 50) return total;
  }

  return null;
}

function extractDailyCallMetrics(text: string): {
  planned: number;
  unplanned: number;
  total: number;
} {
  const cleaned = text
    .replace(/\b\d{1,2}:\d{2}(?:\s*[AP]M)?\b/gi, " ")
    .replace(/\b\d{4,6}\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const tokens = cleaned.split(" ").filter(Boolean);

  for (let i = 0; i <= tokens.length - 5; i++) {
    const parts = tokens.slice(i, i + 5);
    if (!parts.every((token) => /^\d{1,2}$/.test(token))) continue;
    const [planned, unplanned, , , total] = parts.map((token) => parseInt(token, 10));
    if (planned <= 50 && unplanned <= 50 && total <= 50) return { planned, unplanned, total };
  }

  const total = extractTotalFromNumericColumns(text) ?? 0;
  return { planned: 0, unplanned: 0, total };
}

function extractRowStart(line: string): { code: string; rest: string } | null {
  const trimmed = line.replace(/\s+/g, " ").trim();
  if (!trimmed || /\b(employee|emp|code|name|total|region|city)\b/i.test(trimmed.slice(0, 35))) {
    return null;
  }

  const tokens = trimmed.split(" ").filter(Boolean);
  const maxScan = Math.min(tokens.length, 5);
  for (let i = 0; i < maxScan; i++) {
    const token = tokens[i].replace(/[^\d]/g, "");
    if (!/^\d{4,6}$/.test(token)) continue;
    const restTokens = tokens.slice(i + 1);
    if (!restTokens.some((part) => /\p{L}/u.test(part))) continue;
    return { code: token, rest: restTokens.join(" ") };
  }

  return null;
}

function detectRowsFromLines(lines: PdfLine[]): DetectedPdfRow[] {
  const rows: DetectedPdfRow[] = [];
  let current: DetectedPdfRow | null = null;

  for (const line of lines) {
    if (!line.text.trim()) continue;
    const rowStart = extractRowStart(line.text);
    if (rowStart) {
      if (current) rows.push(current);
      current = { code: rowStart.code, text: rowStart.rest, items: line.items };
      continue;
    }

    if (current) {
      current.text = `${current.text} ${line.text}`.replace(/\s+/g, " ").trim();
      current.items.push(...line.items);
    }
  }

  if (current) rows.push(current);
  return rows;
}

function groupItemsIntoLines(items: PositionedPdfItem[]): PdfLine[] {
  const yTolerance = 2.5;
  const lines: { y: number; items: PositionedPdfItem[] }[] = [];

  for (const item of [...items].sort((a, b) => b.y - a.y || a.x - b.x)) {
    let line = lines.find((candidate) => Math.abs(candidate.y - item.y) <= yTolerance);
    if (!line) {
      line = { y: item.y, items: [] };
      lines.push(line);
    }
    line.items.push(item);
    line.y = line.items.reduce((sum, part) => sum + part.y, 0) / line.items.length;
  }

  return lines
    .sort((a, b) => b.y - a.y)
    .map((line) => {
      const parts = line.items.sort((a, b) => a.x - b.x);
      return {
        text: parts
          .map((p) => p.str)
          .join(" ")
          .replace(/\s+/g, " ")
          .trim(),
        items: parts.map((p) => ({ x: p.x, str: p.str })),
      };
    })
    .filter((line) => line.text.length > 0);
}

async function sha256Hex(buf: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", buf);
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function loadPdfJs(): Promise<PdfJs> {
  const [pdfjsLib, worker] = await Promise.all([
    import("pdfjs-dist"),
    import("pdfjs-dist/build/pdf.worker.min.mjs?url"),
  ]);
  pdfjsLib.GlobalWorkerOptions.workerSrc = worker.default;
  return pdfjsLib;
}

export async function parsePdf(file: File): Promise<PdfParseResult> {
  const filenameDate = extractDateFromFilename(file.name);

  const buf = await file.arrayBuffer();
  const fileHash = await sha256Hex(buf);
  const pdfjsLib = await loadPdfJs();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

  const allLines: PdfLine[] = [];
  let rawText = "";
  const totalHeaderXs: number[] = [];

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const pageItems: PositionedPdfItem[] = [];

    for (const item of content.items as PdfTextItem[]) {
      const str = item.str;
      if (!str || !str.trim() || !item.transform) continue;
      const y = item.transform[5];
      const x = item.transform[4];
      pageItems.push({ x, y, str });
    }

    for (const line of groupItemsIntoLines(pageItems)) {
      for (const part of line.items) {
        if (/^total$/i.test(part.str.trim())) totalHeaderXs.push(part.x);
      }
      allLines.push(line);
      rawText += `${line.text}\n`;
    }
  }

  const fallbackYear = filenameDate ? Number(filenameDate.slice(0, 4)) : undefined;
  const date = extractDateFromPdfText(rawText, fallbackYear) ?? filenameDate;
  if (!date) {
    throw new Error(
      `Unable to detect a report date inside "${file.name}". Add a date to the PDF title/content or filename.`,
    );
  }
  const day = parseInt(date.slice(8, 10), 10);
  const totalX = totalHeaderXs.length
    ? totalHeaderXs.reduce((sum, x) => sum + x, 0) / totalHeaderXs.length
    : null;
  const detectedRows = detectRowsFromLines(allLines);
  const rowsByCode = new Map<string, PdfRow>();
  const skippedRows = 0;

  for (const detected of detectedRows) {
    const metrics = extractDailyCallMetrics(detected.text);
    const total = metrics.total || extractTotalValue(detected.text, detected.items, totalX);
    const count = extractSelfieCount(detected.text) ?? total;
    const selfieText = extractSelfieText(detected.text, count);
    const name = cleanName(detected.text) || detected.code;
    const key = detected.code;

    const existing = rowsByCode.get(key);
    if (existing) {
      rowsByCode.set(key, {
        code: detected.code,
        name: existing.name || name,
        planned: Math.max(existing.planned, metrics.planned),
        unplanned: Math.max(existing.unplanned, metrics.unplanned),
        count: Math.max(existing.count, count),
        selfieText: chooseBetterSelfieText(existing.selfieText, selfieText),
        total: Math.max(existing.total, total),
      });
      continue;
    }

    rowsByCode.set(key, {
      code: detected.code,
      name,
      planned: metrics.planned,
      unplanned: metrics.unplanned,
      count,
      selfieText,
      total,
    });
  }

  const rows = [...rowsByCode.values()];

  console.info("[PDF parser]", {
    fileName: file.name,
    filenameDate,
    reportDate: date,
    totalPages: pdf.numPages,
    totalRowsDetected: detectedRows.length,
    parsedRows: rows.length,
    skippedRows,
  });

  return {
    date,
    day,
    fileName: file.name,
    fileHash,
    rows,
    rawText,
    stats: {
      detectedRows: detectedRows.length,
      parsedRows: rows.length,
      skippedRows,
    },
  };
}

function chooseBetterSelfieText(current: string, next: string): string {
  if (!current) return next;
  if (!next) return current;
  const currentHasSelfie = /\b\d+\s*selfies?\b/i.test(current);
  const nextHasSelfie = /\b\d+\s*selfies?\b/i.test(next);
  if (!currentHasSelfie && nextHasSelfie) return next;
  if (currentHasSelfie && !nextHasSelfie) return current;
  return next.length > current.length ? next : current;
}

function cleanName(raw: string): string {
  const tokens = raw.replace(/\s+/g, " ").trim().split(" ").filter(Boolean);
  for (let i = 0; i <= tokens.length - 3; i++) {
    if (!/^\d{1,2}$/.test(tokens[i])) continue;
    if (!/^\d{1,2}$/.test(tokens[i + 1])) continue;
    if (!/^\d{1,2}$/.test(tokens[i + 2])) continue;

    const identity = tokens.slice(0, i);
    if (identity.length >= 3) {
      const nameOnly = identity.slice(0, -2).join(" ");
      if (nameOnly.trim()) return cleanNameText(nameOnly);
    }
  }

  return cleanNameText(raw);
}

function cleanNameText(raw: string): string {
  return raw
    .replace(/\b\d+\s*selfies?\b.*$/i, "")
    .replace(/\btotal\b.*$/i, "")
    .replace(/^\s*\d+[.): -]?\s*/, "")
    .replace(/\s+/g, " ")
    .replace(/[^\p{L}\p{M}\s.'-]/gu, "")
    .trim();
}
