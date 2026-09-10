/**
 * Format-consistency detection within a single column (spec §2.2): mixed
 * date formats, text where the rest of the column is numeric, and
 * inconsistent leading/trailing whitespace or casing. Deterministic, no
 * Office.js.
 */

export type CellShape = "empty" | "number" | "date-iso" | "date-us" | "date-text" | "text";

const NUMBER_RE = /^-?\d+(\.\d+)?$/;
const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const US_DATE_RE = /^\d{1,2}\/\d{1,2}\/\d{2,4}$/;
const TEXT_DATE_RE = /^[A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4}$/;

const DATE_SHAPE_LABEL: Record<string, string> = {
  "date-iso": "YYYY-MM-DD",
  "date-us": "MM/DD/YYYY",
  "date-text": "Month D, YYYY",
};

function classify(raw: unknown): CellShape {
  if (raw === null || raw === undefined) return "empty";
  if (typeof raw === "number") return "number";
  const s = String(raw).trim();
  if (s === "") return "empty";
  if (NUMBER_RE.test(s.replace(/,/g, ""))) return "number";
  if (ISO_DATE_RE.test(s)) return "date-iso";
  if (US_DATE_RE.test(s)) return "date-us";
  if (TEXT_DATE_RE.test(s)) return "date-text";
  return "text";
}

type Casing = "lower" | "upper" | "title" | "mixed";

function classifyCasing(s: string): Casing | null {
  const letters = s.replace(/[^A-Za-z]/g, "");
  if (letters.length === 0) return null;
  if (letters === letters.toLowerCase()) return "lower";
  if (letters === letters.toUpperCase()) return "upper";
  const isTitle = s
    .split(/\s+/)
    .filter(Boolean)
    .every(
      (word) => word[0] === word[0].toUpperCase() && word.slice(1) === word.slice(1).toLowerCase()
    );
  return isTitle ? "title" : "mixed";
}

function mostCommon<T>(counts: Map<T, number>): { value: T; count: number } | null {
  let best: { value: T; count: number } | null = null;
  for (const [value, count] of counts) {
    if (!best || count > best.count) best = { value, count };
  }
  return best;
}

export interface FormatOptions {
  ignoreRows?: number[];
}

export interface FormatFlag {
  ruleId: "format-inconsistency";
  row: number;
  reason: string;
}

/**
 * Check one column's values (already extracted from the sheet range) for
 * internal format inconsistencies. `columnValues[i]` corresponds to sheet
 * row `i` (0-indexed within the checked range).
 */
export function detectFormatInconsistencies(
  columnValues: unknown[],
  options: FormatOptions = {}
): FormatFlag[] {
  const ignore = new Set(options.ignoreRows ?? []);
  const flags: FormatFlag[] = [];

  const shapes = columnValues.map((v, i) => (ignore.has(i) ? "empty" : classify(v)));
  const nonEmptyIdx = shapes
    .map((shape, i) => ({ shape, i }))
    .filter(({ shape, i }) => shape !== "empty" && !ignore.has(i));

  // Need at least two data points to judge "consistency" — a single-row
  // sheet (or a column with only one populated cell) has nothing to compare.
  if (nonEmptyIdx.length < 2) return flags;

  const numberCount = nonEmptyIdx.filter(({ shape }) => shape === "number").length;
  const dateEntries = nonEmptyIdx.filter(({ shape }) => shape.startsWith("date-"));
  const textCount = nonEmptyIdx.filter(({ shape }) => shape === "text").length;

  // Rule A: column is mostly numeric, but a minority of cells are plain text.
  if (numberCount > 0 && textCount > 0 && numberCount >= textCount) {
    for (const { i } of nonEmptyIdx.filter(({ shape }) => shape === "text")) {
      flags.push({
        ruleId: "format-inconsistency",
        row: i,
        reason: "Column is mostly numeric, but this cell contains text",
      });
    }
  }

  // Rule B: mixed date formats within the same column.
  if (dateEntries.length >= 2) {
    const shapeCounts = new Map<string, number>();
    for (const { shape } of dateEntries) shapeCounts.set(shape, (shapeCounts.get(shape) ?? 0) + 1);
    const majority = mostCommon(shapeCounts);
    if (majority && shapeCounts.size > 1) {
      for (const { shape, i } of dateEntries) {
        if (shape === majority.value) continue;
        flags.push({
          ruleId: "format-inconsistency",
          row: i,
          reason: `Column mostly uses ${DATE_SHAPE_LABEL[majority.value]} dates, but this cell uses ${DATE_SHAPE_LABEL[shape]}`,
        });
      }
    }
  }

  // Rule C: stray leading/trailing whitespace (always worth flagging).
  columnValues.forEach((raw, i) => {
    if (ignore.has(i) || typeof raw !== "string") return;
    if (raw !== raw.trim() && raw.trim() !== "") {
      flags.push({
        ruleId: "format-inconsistency",
        row: i,
        reason: "Cell has leading/trailing whitespace",
      });
    }
  });

  // Rule D: inconsistent casing among plain-text cells.
  const textEntries = nonEmptyIdx
    .filter(({ shape }) => shape === "text")
    .map(({ i }) => ({ i, casing: classifyCasing(String(columnValues[i]).trim()) }))
    .filter((e): e is { i: number; casing: Casing } => e.casing !== null);
  if (textEntries.length >= 2) {
    const casingCounts = new Map<Casing, number>();
    for (const { casing } of textEntries)
      casingCounts.set(casing, (casingCounts.get(casing) ?? 0) + 1);
    const majority = mostCommon(casingCounts);
    if (majority && majority.count > textEntries.length / 2 && casingCounts.size > 1) {
      for (const { i, casing } of textEntries) {
        if (casing === majority.value) continue;
        flags.push({
          ruleId: "format-inconsistency",
          row: i,
          reason: `Column mostly uses ${majority.value} case text, but this cell uses ${casing} case`,
        });
      }
    }
  }

  return flags.sort((a, b) => a.row - b.row);
}
