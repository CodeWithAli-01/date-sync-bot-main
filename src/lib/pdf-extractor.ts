// Client-side PDF text extraction + employee-row parsing.
// Strong mode treats every numeric employee code as the start of a row.
import * as pdfjsLib from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

export interface PdfRow {
  code: string | null;
  name: string;
  count: number;
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

type DetectedPdfRow = {
  code: string;
  text: string;
  items: { x: number; str: string }[];
};

function extractDateFromFilename(fileName: string): string | null {
  const m = fileName.match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : null;
}

function extractSelfieCount(text: string): number | null {
  const lower = text.toLowerCase();
  const selfie = lower.match(/(\d+)\s*selfies?\b/);
  if (selfie) return parseInt(selfie[1], 10);

  return null;
}

function extractTotalValue(
  text: string,
  items: { x: number; str: string }[],
  totalX: number | null,
): number {
  const lower = text.toLowerCase();
  const labeled = lower.match(/\btotal\b\D{0,12}(\d+)/);
  if (labeled) return parseInt(labeled[1], 10);

  if (totalX == null) return 0;

  const numericItems = items
    .filter((item) => /^\d+$/.test(item.str.trim()))
    .map((item) => ({ value: parseInt(item.str.trim(), 10), distance: Math.abs(item.x - totalX) }))
    .sort((a, b) => a.distance - b.distance);

  return numericItems[0]?.value ?? 0;
}

function extractRowStart(line: string): { code: string; rest: string } | null {
  const trimmed = line.replace(/\s+/g, " ").trim();
  const match = trimmed.match(/^(?:\d{1,3}[.)-]?\s+)?(\d{4,6})\b\s*(.*)$/);
  if (!match) return null;
  return { code: match[1], rest: match[2] ?? "" };
}

function detectRowsFromLines(
  lines: { text: string; items: { x: number; str: string }[] }[],
): DetectedPdfRow[] {
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

async function sha256Hex(buf: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", buf);
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

export async function parsePdf(file: File): Promise<PdfParseResult> {
  const date = extractDateFromFilename(file.name);
  if (!date) {
    throw new Error(`Filename "${file.name}" must contain a YYYY-MM-DD date.`);
  }
  const day = parseInt(date.slice(8, 10), 10);

  const buf = await file.arrayBuffer();
  const fileHash = await sha256Hex(buf);
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

  const allLines: { text: string; items: { x: number; str: string }[] }[] = [];
  let rawText = "";
  const totalHeaderXs: number[] = [];

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const lineMap = new Map<number, { x: number; str: string }[]>();

    for (const item of content.items as PdfTextItem[]) {
      const str = item.str;
      if (!str || !str.trim() || !item.transform) continue;
      const y = Math.round(item.transform[5]);
      const x = item.transform[4];
      const arr = lineMap.get(y) ?? [];
      arr.push({ x, str });
      lineMap.set(y, arr);
    }

    const sortedY = [...lineMap.keys()].sort((a, b) => b - a);
    for (const y of sortedY) {
      const parts = lineMap.get(y)!.sort((a, b) => a.x - b.x);
      const line = parts
        .map((p) => p.str)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim();
      if (!line) continue;
      for (const part of parts) {
        if (/^total$/i.test(part.str.trim())) totalHeaderXs.push(part.x);
      }
      allLines.push({ text: line, items: parts });
      rawText += `${line}\n`;
    }
  }

  const totalX = totalHeaderXs.length
    ? totalHeaderXs.reduce((sum, x) => sum + x, 0) / totalHeaderXs.length
    : null;
  const detectedRows = detectRowsFromLines(allLines);
  const rows: PdfRow[] = [];
  const seen = new Set<string>();
  let skippedRows = 0;

  for (const detected of detectedRows) {
    const count = extractSelfieCount(detected.text) ?? 0;
    const total = extractTotalValue(detected.text, detected.items, totalX);
    const name = cleanName(detected.text) || detected.code;
    const key = detected.code;

    if (seen.has(key)) {
      skippedRows++;
      continue;
    }

    seen.add(key);
    rows.push({ code: detected.code, name, count, total });
  }

  console.info("[PDF parser]", {
    fileName: file.name,
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

function cleanName(raw: string): string {
  return raw
    .replace(/\b\d+\s*selfies?\b.*$/i, "")
    .replace(/\b\d+\s*calls?\b.*$/i, "")
    .replace(/\btotal\b.*$/i, "")
    .replace(/^\s*\d+[.): -]?\s*/, "")
    .replace(/\s+/g, " ")
    .replace(/[^\p{L}\p{M}\s.'-]/gu, "")
    .trim();
}
