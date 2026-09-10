import { detectDuplicateRows } from "./duplicates";
import { detectFormatInconsistencies } from "./formatConsistency";
import { detectOutliers } from "./outliers";
import { runEngine } from "./index";

describe("detectDuplicateRows", () => {
  it("flags exact full-row duplicates by default", () => {
    const rows = [
      ["Alice", 30],
      ["Bob", 40],
      ["Alice", 30],
    ];
    const flags = detectDuplicateRows(rows);
    expect(flags.map((f) => f.row)).toEqual([0, 2]);
    expect(flags[0].groupRows).toEqual([0, 2]);
  });

  it("does not flag rows that differ in any column", () => {
    const rows = [
      ["Alice", 30],
      ["Alice", 31],
    ];
    expect(detectDuplicateRows(rows)).toHaveLength(0);
  });

  it("matches on selected key columns only when provided", () => {
    const rows = [
      ["Alice", 30, "NY"],
      ["Alice", 31, "CA"],
    ];
    const flags = detectDuplicateRows(rows, { keyColumns: [0] });
    expect(flags.map((f) => f.row)).toEqual([0, 1]);
  });

  it("normalizes whitespace and casing before comparing", () => {
    const rows = [["Alice"], [" alice "]];
    expect(detectDuplicateRows(rows)).toHaveLength(2);
  });

  it("never flags fully blank rows against each other", () => {
    const rows = [
      ["", null],
      ["", undefined],
      ["Alice", 30],
    ];
    expect(detectDuplicateRows(rows)).toHaveLength(0);
  });

  it("respects ignoreRows (e.g. a header row)", () => {
    const rows = [
      ["Name", "Age"],
      ["Name", "Age"],
    ];
    expect(detectDuplicateRows(rows, { ignoreRows: [0, 1] })).toHaveLength(0);
  });

  it("handles a single-row sheet without error", () => {
    expect(detectDuplicateRows([["Alice", 30]])).toHaveLength(0);
  });

  it("groups three-way duplicates together", () => {
    const rows = [["X"], ["X"], ["X"]];
    const flags = detectDuplicateRows(rows);
    expect(flags).toHaveLength(3);
    expect(flags[0].groupRows).toEqual([0, 1, 2]);
    expect(flags[0].reason).toContain("rows");
  });
});

describe("detectFormatInconsistencies", () => {
  it("flags text cells in an otherwise-numeric column", () => {
    const column = [1, 2, 3, "N/A", 5];
    const flags = detectFormatInconsistencies(column);
    expect(flags.map((f) => f.row)).toEqual([3]);
    expect(flags[0].reason).toMatch(/mostly numeric/);
  });

  it("does not flag a numeric column with no text at all", () => {
    expect(detectFormatInconsistencies([1, 2, 3, 4])).toHaveLength(0);
  });

  it("flags the minority date format in a mixed-format column", () => {
    const column = ["2024-01-01", "2024-01-02", "2024-01-03", "01/04/2024"];
    const flags = detectFormatInconsistencies(column);
    expect(flags.map((f) => f.row)).toEqual([3]);
    expect(flags[0].reason).toMatch(/MM\/DD\/YYYY/);
  });

  it("does not flag a column that consistently uses one date format", () => {
    const column = ["2024-01-01", "2024-01-02", "2024-01-03"];
    expect(detectFormatInconsistencies(column)).toHaveLength(0);
  });

  it("flags leading/trailing whitespace", () => {
    const column = ["clean", " dirty", "also-clean"];
    const flags = detectFormatInconsistencies(column);
    expect(flags.some((f) => f.row === 1 && f.reason.includes("whitespace"))).toBe(true);
  });

  it("flags minority casing in a mostly-lowercase text column", () => {
    const column = ["apple", "banana", "cherry", "DURIAN"];
    const flags = detectFormatInconsistencies(column);
    expect(flags.some((f) => f.row === 3 && f.reason.includes("case"))).toBe(true);
  });

  it("treats empty cells as absent, not as a format defect", () => {
    const column = [1, 2, null, undefined, "", 3];
    expect(detectFormatInconsistencies(column)).toHaveLength(0);
  });

  it("does nothing on a single populated cell (nothing to compare)", () => {
    expect(detectFormatInconsistencies([42])).toHaveLength(0);
  });

  it("handles a fully empty column without error", () => {
    expect(detectFormatInconsistencies([null, undefined, ""])).toHaveLength(0);
  });
});

describe("detectOutliers", () => {
  it("flags values outside a configured min/max", () => {
    const column = [10, 20, 30, -5, 999];
    const flags = detectOutliers(column, { min: 0, max: 100 });
    // Row 4 (999) is far enough outside [0, 100] that it also trips the
    // default IQR statistical check — two independent rules, two flags for
    // the same cell, which is expected (each carries its own reason).
    expect(new Set(flags.map((f) => f.row))).toEqual(new Set([3, 4]));
    expect(flags.some((f) => f.row === 3 && f.reason.includes("minimum"))).toBe(true);
    expect(flags.some((f) => f.row === 4 && f.reason.includes("maximum"))).toBe(true);
  });

  it("flags values not present in an allow-list, ignoring case/whitespace", () => {
    const column = ["Active", "active", " Closed ", "Bogus"];
    const flags = detectOutliers(column, { allowList: ["active", "closed"] });
    expect(flags.map((f) => f.row)).toEqual([3]);
  });

  it("flags IQR statistical outliers", () => {
    const column = [10, 11, 12, 11, 10, 12, 500];
    const flags = detectOutliers(column, { method: "iqr" });
    expect(flags.map((f) => f.row)).toContain(6);
  });

  it("flags z-score statistical outliers", () => {
    const column = [10, 11, 12, 11, 10, 12, 500];
    const flags = detectOutliers(column, { method: "zscore", zScoreThreshold: 2 });
    expect(flags.map((f) => f.row)).toContain(6);
  });

  it("skips statistical checks below the minimum sample size (avoids noise on tiny sheets)", () => {
    expect(detectOutliers([10, 20, 30])).toHaveLength(0);
  });

  it("does not divide by zero / flag anything when every value is identical", () => {
    expect(detectOutliers([5, 5, 5, 5, 5])).toHaveLength(0);
  });

  it("ignores non-numeric text cells for statistical checks", () => {
    const column = ["n/a", 10, 11, 12, 10, 11];
    expect(() => detectOutliers(column)).not.toThrow();
  });
});

describe("runEngine (orchestrator)", () => {
  it("runs all three rules with zero configuration", () => {
    const rows = [
      ["Name", "Amount"],
      ["Alice", 100],
      ["Alice", 100],
      ["Bob", "oops"],
    ];
    const flags = runEngine(rows, { ignoreRows: [0] });
    const ruleIds = new Set(flags.map((f) => f.ruleId));
    expect(ruleIds.has("duplicate-row")).toBe(true);
    expect(ruleIds.has("format-inconsistency")).toBe(true);
  });

  it("lets a caller disable a rule entirely", () => {
    const rows = [
      ["Alice", 100],
      ["Alice", 100],
    ];
    const flags = runEngine(rows, { duplicates: false });
    expect(flags.some((f) => f.ruleId === "duplicate-row")).toBe(false);
  });

  it("lets a caller disable a rule for one column only", () => {
    const rows = [
      [1, "n/a"],
      [2, "n/a"],
      [3, "n/a"],
      [4, "n/a"],
    ];
    const flags = runEngine(rows, {
      duplicates: false,
      columns: { 1: { format: false } },
    });
    expect(flags.some((f) => f.col === 1)).toBe(false);
  });

  it("produces stable, unique ids per flag", () => {
    const rows = [
      ["Alice", 100],
      ["Alice", 100],
    ];
    const flags = runEngine(rows);
    const ids = flags.map((f) => f.id);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("handles a single-row sheet without error", () => {
    expect(() => runEngine([["Alice", 30, "2024-01-01"]])).not.toThrow();
  });

  it("handles an empty sheet without error", () => {
    expect(runEngine([])).toEqual([]);
  });
});
