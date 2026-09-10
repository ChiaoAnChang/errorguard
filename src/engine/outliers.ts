/**
 * Out-of-range / outlier detection within a single column (spec §2.2): IQR
 * or z-score statistical outliers, plus optional user-defined min/max or
 * allow-list. Deterministic, no Office.js.
 */

export interface OutlierOptions {
  method?: "iqr" | "zscore";
  zScoreThreshold?: number;
  iqrMultiplier?: number;
  min?: number;
  max?: number;
  allowList?: (string | number)[];
  ignoreRows?: number[];
}

export interface OutlierFlag {
  ruleId: "outlier";
  row: number;
  reason: string;
}

const DEFAULT_Z_THRESHOLD = 3;
const DEFAULT_IQR_MULTIPLIER = 1.5;
// Below this many numeric samples, a statistical outlier check is more
// likely to flag noise than a real defect (spec's "single-row sheets" edge
// case), so it's skipped entirely; min/max/allow-list checks still run.
const MIN_SAMPLES_FOR_STATS = 4;

function quantile(sortedValues: number[], q: number): number {
  const pos = (sortedValues.length - 1) * q;
  const base = Math.floor(pos);
  const rest = pos - base;
  const next = sortedValues[base + 1];
  return next !== undefined
    ? sortedValues[base] + rest * (next - sortedValues[base])
    : sortedValues[base];
}

function toNumber(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "string" && value.trim() !== "") {
    const n = Number(value);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

/**
 * Check one column's values for range/allow-list/statistical-outlier
 * violations. `columnValues[i]` corresponds to sheet row `i` (0-indexed
 * within the checked range).
 */
export function detectOutliers(
  columnValues: unknown[],
  options: OutlierOptions = {}
): OutlierFlag[] {
  const ignore = new Set(options.ignoreRows ?? []);
  const flags: OutlierFlag[] = [];

  if (options.allowList && options.allowList.length > 0) {
    const allowed = new Set(options.allowList.map((v) => String(v).trim().toLowerCase()));
    columnValues.forEach((v, i) => {
      if (ignore.has(i) || v === null || v === undefined || String(v).trim() === "") return;
      if (!allowed.has(String(v).trim().toLowerCase())) {
        flags.push({
          ruleId: "outlier",
          row: i,
          reason: `Value "${v}" is not in the allowed list for this column`,
        });
      }
    });
  }

  const numericEntries: { row: number; value: number }[] = [];
  columnValues.forEach((v, i) => {
    if (ignore.has(i)) return;
    const n = toNumber(v);
    if (n !== null) numericEntries.push({ row: i, value: n });
  });

  if (options.min !== undefined || options.max !== undefined) {
    for (const { row, value } of numericEntries) {
      if (options.min !== undefined && value < options.min) {
        flags.push({
          ruleId: "outlier",
          row,
          reason: `Value ${value} is below the configured minimum (${options.min})`,
        });
      } else if (options.max !== undefined && value > options.max) {
        flags.push({
          ruleId: "outlier",
          row,
          reason: `Value ${value} is above the configured maximum (${options.max})`,
        });
      }
    }
  }

  if (numericEntries.length >= MIN_SAMPLES_FOR_STATS) {
    const method = options.method ?? "iqr";
    if (method === "iqr") {
      const sorted = [...numericEntries].sort((a, b) => a.value - b.value).map((e) => e.value);
      const q1 = quantile(sorted, 0.25);
      const q3 = quantile(sorted, 0.75);
      const iqr = q3 - q1;
      if (iqr > 0) {
        const mult = options.iqrMultiplier ?? DEFAULT_IQR_MULTIPLIER;
        const lower = q1 - mult * iqr;
        const upper = q3 + mult * iqr;
        for (const { row, value } of numericEntries) {
          if (value < lower || value > upper) {
            flags.push({
              ruleId: "outlier",
              row,
              reason: `Value ${value} is a statistical outlier (outside [${lower.toFixed(2)}, ${upper.toFixed(2)}] by IQR)`,
            });
          }
        }
      }
    } else {
      const values = numericEntries.map((e) => e.value);
      const mean = values.reduce((a, b) => a + b, 0) / values.length;
      const variance = values.reduce((a, b) => a + (b - mean) ** 2, 0) / values.length;
      const std = Math.sqrt(variance);
      if (std > 0) {
        const threshold = options.zScoreThreshold ?? DEFAULT_Z_THRESHOLD;
        for (const { row, value } of numericEntries) {
          const z = Math.abs((value - mean) / std);
          if (z > threshold) {
            flags.push({
              ruleId: "outlier",
              row,
              reason: `Value ${value} is a statistical outlier (z-score ${z.toFixed(2)} > ${threshold})`,
            });
          }
        }
      }
    }
  }

  return flags.sort((a, b) => a.row - b.row);
}
