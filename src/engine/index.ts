/**
 * Pure detection engine entry point (spec §4): takes a plain 2D array of
 * cell values in, returns a flat list of flagged cells/reasons out. Zero
 * dependency on Office.js — this is what makes the engine unit-testable
 * without a live Excel instance (spec §3, "Test coverage").
 */

import { detectDuplicateRows } from "./duplicates";
import type { DuplicateOptions } from "./duplicates";
import { detectFormatInconsistencies } from "./formatConsistency";
import type { FormatOptions } from "./formatConsistency";
import { detectOutliers } from "./outliers";
import type { OutlierOptions } from "./outliers";

// Re-exported explicitly as `export type` (rather than a blanket `export *`)
// so the webpack/babel build — which transpiles per-file without full
// cross-module type information — can tell these are type-only and safely
// erase them, instead of warning that a type name isn't a real runtime export.
export { detectDuplicateRows } from "./duplicates";
export type { DuplicateOptions, DuplicateFlag } from "./duplicates";
export { detectFormatInconsistencies } from "./formatConsistency";
export type { FormatOptions, FormatFlag, CellShape } from "./formatConsistency";
export { detectOutliers } from "./outliers";
export type { OutlierOptions, OutlierFlag } from "./outliers";

export type RuleId = "duplicate-row" | "format-inconsistency" | "outlier";

export interface ColumnRuleConfig {
  /** `false` disables the rule for this column; an options object enables it with overrides. */
  format?: boolean | FormatOptions;
  outlier?: boolean | OutlierOptions;
}

export interface EngineOptions {
  /** `false` disables duplicate-row detection entirely; an options object enables it with overrides. */
  duplicates?: boolean | DuplicateOptions;
  /** Global default for the format-inconsistency rule (default: on). A per-column entry in `columns` overrides this. */
  format?: boolean;
  /** Global default for the outlier rule (default: on). A per-column entry in `columns` overrides this. */
  outlier?: boolean;
  /** Per-column rule overrides, keyed by 0-indexed column number. */
  columns?: Record<number, ColumnRuleConfig>;
  /** Row indices to exclude everywhere (e.g. a header row). */
  ignoreRows?: number[];
}

export interface CellFlag {
  /** Stable id for dismiss/mark-reviewed tracking. */
  id: string;
  ruleId: RuleId;
  row: number;
  /** Absent for whole-row flags (duplicate-row). */
  col?: number;
  reason: string;
}

/**
 * Run all enabled rules over `rows` (a plain 2D array shaped like an Excel
 * range's `.values`) and return every flag, sorted by row then column.
 * Zero-configuration defaults: duplicates + format + outlier all on.
 */
export function runEngine(rows: unknown[][], options: EngineOptions = {}): CellFlag[] {
  const flags: CellFlag[] = [];
  const ignoreRows = options.ignoreRows ?? [];

  if (options.duplicates !== false) {
    const overrides = typeof options.duplicates === "object" ? options.duplicates : {};
    for (const f of detectDuplicateRows(rows, { ignoreRows, ...overrides })) {
      flags.push({
        id: `duplicate-row:${f.row}`,
        ruleId: "duplicate-row",
        row: f.row,
        reason: f.reason,
      });
    }
  }

  const colCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
  for (let col = 0; col < colCount; col++) {
    const columnValues = rows.map((row) => row[col]);
    const colConfig = options.columns?.[col] ?? {};
    const formatSetting = colConfig.format ?? options.format ?? true;
    const outlierSetting = colConfig.outlier ?? options.outlier ?? true;

    if (formatSetting !== false) {
      const overrides = typeof formatSetting === "object" ? formatSetting : {};
      for (const f of detectFormatInconsistencies(columnValues, { ignoreRows, ...overrides })) {
        flags.push({
          id: `format:${f.row}:${col}`,
          ruleId: "format-inconsistency",
          row: f.row,
          col,
          reason: f.reason,
        });
      }
    }

    if (outlierSetting !== false) {
      const overrides = typeof outlierSetting === "object" ? outlierSetting : {};
      for (const f of detectOutliers(columnValues, { ignoreRows, ...overrides })) {
        flags.push({
          id: `outlier:${f.row}:${col}`,
          ruleId: "outlier",
          row: f.row,
          col,
          reason: f.reason,
        });
      }
    }
  }

  return flags.sort((a, b) => a.row - b.row || (a.col ?? -1) - (b.col ?? -1));
}
