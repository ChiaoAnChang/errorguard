/**
 * Duplicate-row detection (spec §2.2, §8): deterministic, no Office.js.
 *
 * v1 default is full-row exact match (per the spec's own recommendation —
 * "start with full-row exact match; add key-column matching once the basic
 * engine is proven"). Passing `keyColumns` switches to matching on just
 * those columns instead, for callers that already want that.
 */

export interface DuplicateOptions {
  /** Column indices to match on. Omit/empty for full-row exact match (v1 default). */
  keyColumns?: number[];
  /** Row indices to exclude from matching (e.g. a header row). */
  ignoreRows?: number[];
}

export interface DuplicateFlag {
  ruleId: "duplicate-row";
  row: number;
  reason: string;
  /** Every row index sharing this row's key, including this one. */
  groupRows: number[];
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") return String(value);
  return String(value).trim().toLowerCase();
}

function isBlankRow(row: unknown[]): boolean {
  return row.every((v) => v === null || v === undefined || String(v).trim() === "");
}

function rowKey(row: unknown[], keyColumns?: number[]): string {
  const cols = keyColumns && keyColumns.length > 0 ? keyColumns : row.map((_, i) => i);
  return cols.map((c) => normalizeCell(row[c])).join("");
}

/**
 * Flag every row that exact-matches at least one other row (full-row by
 * default, or on `options.keyColumns` if given). Fully blank rows are never
 * flagged against each other — an empty row isn't a duplicate data entry.
 */
export function detectDuplicateRows(
  rows: unknown[][],
  options: DuplicateOptions = {}
): DuplicateFlag[] {
  const ignore = new Set(options.ignoreRows ?? []);
  const groups = new Map<string, number[]>();

  rows.forEach((row, i) => {
    if (ignore.has(i) || isBlankRow(row)) return;
    const key = rowKey(row, options.keyColumns);
    const list = groups.get(key);
    if (list) {
      list.push(i);
    } else {
      groups.set(key, [i]);
    }
  });

  const basis =
    options.keyColumns && options.keyColumns.length > 0
      ? "the selected key column(s)"
      : "all columns";

  const flags: DuplicateFlag[] = [];
  for (const groupRows of groups.values()) {
    if (groupRows.length < 2) continue;
    for (const row of groupRows) {
      const others = groupRows.filter((r) => r !== row).map((r) => r + 1);
      flags.push({
        ruleId: "duplicate-row",
        row,
        reason: `Duplicate row: matches row${others.length > 1 ? "s" : ""} ${others.join(", ")} on ${basis}`,
        groupRows,
      });
    }
  }
  return flags.sort((a, b) => a.row - b.row);
}
