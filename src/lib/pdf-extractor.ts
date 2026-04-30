// Client-side PDF text extraction + selfie count parsing.
// Extracts (employee_code?, name, count) per row from messy daily PDFs.
import * as pdfjsLib from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

export interface PdfRow {
  code: string | null; // employee code if detectable
  name: string;
  count: number;
}

export interface PdfParseResult {
  date: string; // YYYY-MM-DD
  day: number; // 1..31
  fileName: string;
  fileHash: string; // sha-256 hex of bytes
  rows: PdfRow[];
  rawText: string;
}

function extractDateFromFilename(fileName: string): string | null {
  const m = fileName.match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : null;
}

// "11 selfies" / "0 selfie" / "selfies - 5" / "Total: 7"
function extractSelfieCount(text: string): number | null {
  const lower = text.toLowerCase();
  const m1 = lower.match(/(\d+)\s*selfie/);
  if (m1) return parseInt(m1[1], 10);
  const m2 = lower.match(/selfie[^\d]{0,15}(\d+)/);
  if (m2) return parseInt(m2[1], 10);
  const m3 = lower.match(/total[\s:]+(\d+)/);
  if (m3) return parseInt(m3[1], 10);
  return null;
}

// Detect a leading employee code on a row. Common patterns:
//   "EMP001 Mashad Hussain ..."   "1234 Mashad ..."   "P-23 Mashad ..."
function extractCode(line: string): string | null {
  const m = line.match(/^\s*([A-Z]{1,5}[-_]?\d{2,6}|\d{3,6})\b/i);
  return m ? m[1].toUpperCase() : null;
}

async function sha256Hex(buf: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", buf);
  return [...new Uint8Array(digest)]
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
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

  const linesByPage: string[][] = [];
  let rawText = "";

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const lineMap = new Map<number, { x: number; str: string }[]>();
    for (const item of content.items as any[]) {
      const str = item.str;
      if (!str || !str.trim()) continue;
      const y = Math.round(item.transform[5]);
      const x = item.transform[4];
      const arr = lineMap.get(y) ?? [];
      arr.push({ x, str });
      lineMap.set(y, arr);
    }
    const sortedY = [...lineMap.keys()].sort((a, b) => b - a);
    const lines: string[] = [];
    for (const y of sortedY) {
      const parts = lineMap.get(y)!.sort((a, b) => a.x - b.x);
      lines.push(parts.map((p) => p.str).join(" ").replace(/\s+/g, " ").trim());
    }
    linesByPage.push(lines);
    rawText += lines.join("\n") + "\n";
  }

  const rows: PdfRow[] = [];
  const seen = new Set<string>();
  const flatLines = linesByPage.flat();

  for (let i = 0; i < flatLines.length; i++) {
    const line = flatLines[i];
    if (!line) continue;

    const count = extractSelfieCount(line);
    if (count === null) continue;

    // Try same-line: "<code?> <name> ... <n> selfies"
    let name = "";
    let code: string | null = null;
    const sameLineMatch = line.match(/^(.+?)[\s\-:|,]+\d+\s*selfie/i);
    if (sameLineMatch) {
      const head = sameLineMatch[1];
      code = extractCode(head);
      name = code ? head.replace(/^[^\s]+\s+/, "") : head;
    } else {
      // fall back to previous non-empty line
      for (let j = i - 1; j >= 0; j--) {
        if (flatLines[j] && !/^\s*$/.test(flatLines[j])) {
          const prev = flatLines[j];
          code = extractCode(prev);
          name = code ? prev.replace(/^[^\s]+\s+/, "") : prev;
          break;
        }
      }
    }

    name = cleanName(name);
    if (!name) continue;

    // Skip "0 0 0" style empty rows (no selfie count AND zero numbers only)
    if (count === 0 && /^[\s0]+$/.test(name)) continue;

    const key = (code ?? "") + "|" + name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);

    rows.push({ code, name, count });
  }

  return { date, day, fileName: file.name, fileHash, rows, rawText };
}

function cleanName(raw: string): string {
  return raw
    .replace(/^\s*\d+[\.\):\-]?\s*/, "")
    .replace(/\s+/g, " ")
    .replace(/[^\p{L}\p{M}\s\.\-']/gu, "")
    .trim();
}
