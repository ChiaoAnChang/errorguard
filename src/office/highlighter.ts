/**
 * Applies/clears cell highlighting for flagged cells (spec §4, office/).
 *
 * Flags come from `src/engine`'s plain JS logic, not from a formula Excel
 * could evaluate itself — so this uses direct cell fill formatting rather
 * than Excel's ConditionalFormat feature (which needs a worksheet-evaluable
 * rule, not an externally-computed set of cells). Only fill color is
 * touched, never font color: we don't record each cell's original font
 * color before overriding it, so we can't restore it on clear — fill alone
 * is fully reversible via `.clear()` without losing the user's formatting.
 * All set/clear calls are queued before a single `context.sync()` per pass,
 * so a highlight refresh is one round-trip regardless of flag count.
 */

/* global Excel */

const FLAG_FILL_COLOR = "#FFC7CE"; // matches Excel's built-in "Bad" cell style

export interface RangeAnchor {
  row: number;
  col: number;
}

function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

/**
 * Expand `flags` into the full set of "row:col" cell keys that should be
 * highlighted, relative to the checked range's own (row, col) origin.
 * Whole-row flags (duplicate-row, which has no `col`) expand across every
 * column in the checked range.
 */
export function flagsToCellKeys(
  flags: { row: number; col?: number }[],
  columnCount: number
): Set<string> {
  const keys = new Set<string>();
  for (const flag of flags) {
    if (flag.col !== undefined) {
      keys.add(cellKey(flag.row, flag.col));
    } else {
      for (let c = 0; c < columnCount; c++) keys.add(cellKey(flag.row, c));
    }
  }
  return keys;
}

/**
 * Sync highlighting to exactly match `flags`: apply fill to newly-flagged
 * cells, clear fill from cells that were highlighted last pass but aren't
 * flagged anymore. Returns the new highlighted-cell-key set so the caller
 * can pass it back in on the next pass (see `office/eventBinding.ts`).
 */
export async function applyHighlights(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  anchor: RangeAnchor,
  columnCount: number,
  flags: { row: number; col?: number }[],
  previouslyHighlighted: Set<string>
): Promise<Set<string>> {
  const nextHighlighted = flagsToCellKeys(flags, columnCount);

  for (const key of previouslyHighlighted) {
    if (nextHighlighted.has(key)) continue;
    const [row, col] = key.split(":").map(Number);
    const cell = sheet.getRangeByIndexes(anchor.row + row, anchor.col + col, 1, 1);
    cell.format.fill.clear();
  }
  for (const key of nextHighlighted) {
    if (previouslyHighlighted.has(key)) continue;
    const [row, col] = key.split(":").map(Number);
    const cell = sheet.getRangeByIndexes(anchor.row + row, anchor.col + col, 1, 1);
    cell.format.fill.color = FLAG_FILL_COLOR;
  }

  await context.sync();
  return nextHighlighted;
}

/** Clear every currently-highlighted cell (e.g. on dispose, or when checks are turned off). */
export async function clearAllHighlights(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  anchor: RangeAnchor,
  highlighted: Set<string>
): Promise<void> {
  for (const key of highlighted) {
    const [row, col] = key.split(":").map(Number);
    const cell = sheet.getRangeByIndexes(anchor.row + row, anchor.col + col, 1, 1);
    cell.format.fill.clear();
  }
  await context.sync();
}
