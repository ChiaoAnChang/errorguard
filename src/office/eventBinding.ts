/**
 * Wires the pure engine (src/engine) to a live worksheet via
 * `worksheet.onChanged`, debounced so a burst of keystrokes/pastes triggers
 * one check, not one per event (spec §3, "Performance").
 *
 * This file is the one part of ErrorGuard that genuinely needs a live Excel
 * host to verify — Office.js has no meaningful test double for
 * `worksheet.onChanged` / `getUsedRangeOrNullObject`, which is exactly why
 * `src/engine` is kept free of this dependency (see that directory's own
 * comment). Manually sideload-test this against the synthetic workbook in
 * `test-data/` before relying on it (spec execution plan, phase 5).
 */

/* global Excel */

import { runEngine } from "../engine";
import type { EngineOptions, CellFlag } from "../engine";
import { applyHighlights, clearAllHighlights } from "./highlighter";
import type { RangeAnchor } from "./highlighter";

const DEBOUNCE_MS = 800;

export type FlagsChangedListener = (flags: CellFlag[]) => void;

export interface WatcherHandle {
  /** Stop watching and clear any highlights this watcher applied. */
  dispose: () => Promise<void>;
  /** Re-run the check immediately (e.g. after the user changes rule settings). */
  recheck: () => Promise<void>;
}

/**
 * Watch `sheetName`'s used range for changes and keep highlighting + the
 * task pane's flag list in sync.
 *
 * @param onFlagsChanged called with the FULL flag list (including dismissed
 *   ones — the task pane needs those to render "dismissed" state) after
 *   every check.
 * @param getEngineOptions called fresh on every check, so live settings
 *   changes (rule toggles, per-column overrides) take effect on the next run.
 * @param isDismissed used to filter which flags actually get highlighted on
 *   the sheet — a dismissed flag is still reported to the pane, just not
 *   re-highlighted, so it doesn't "come back" until the cell itself changes
 *   (which gives it a new flag id — see src/engine/index.ts).
 */
export async function watchWorksheet(
  sheetName: string,
  onFlagsChanged: FlagsChangedListener,
  getEngineOptions: () => EngineOptions,
  isDismissed: (flagId: string) => boolean
): Promise<WatcherHandle> {
  let debounceTimer: ReturnType<typeof setTimeout> | undefined;
  let highlighted = new Set<string>();
  let lastAnchor: RangeAnchor = { row: 0, col: 0 };
  let disposed = false;

  const runCheck = async (): Promise<void> => {
    if (disposed) return;
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const usedRange = sheet.getUsedRangeOrNullObject(true);
      // Check isNullObject on its own sync before loading the (potentially
      // large) values/navigational properties below — no point loading a
      // used range's contents just to discover the sheet is empty. (This
      // still trips eslint-plugin-office-addins' no-navigational-load
      // warning, which flags any isNullObject load; there's no alternative
      // to loading it when using the getXOrNullObject pattern, so it's left
      // as a warning rather than suppressed.)
      usedRange.load("isNullObject");
      await context.sync();

      if (usedRange.isNullObject) {
        highlighted = await applyHighlights(context, sheet, lastAnchor, 0, [], highlighted);
        onFlagsChanged([]);
        return;
      }

      usedRange.load(["values", "rowIndex", "columnIndex", "columnCount"]);
      await context.sync();

      lastAnchor = { row: usedRange.rowIndex, col: usedRange.columnIndex };
      const allFlags = runEngine(usedRange.values as unknown[][], getEngineOptions());
      const visibleFlags = allFlags.filter((f) => !isDismissed(f.id));

      highlighted = await applyHighlights(
        context,
        sheet,
        lastAnchor,
        usedRange.columnCount,
        visibleFlags,
        highlighted
      );
      onFlagsChanged(allFlags);
    });
  };

  const scheduleCheck = (): void => {
    if (debounceTimer) clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
      runCheck().catch((error) => console.error("ErrorGuard: check failed", error));
    }, DEBOUNCE_MS);
  };

  const eventResult = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const handler = sheet.onChanged.add(async () => {
      scheduleCheck();
    });
    await context.sync();
    return handler;
  });

  // Initial pass so the pane isn't empty until the user's first edit.
  await runCheck();

  return {
    recheck: runCheck,
    dispose: async () => {
      disposed = true;
      if (debounceTimer) clearTimeout(debounceTimer);
      await Excel.run(eventResult.context, async (context) => {
        eventResult.remove();
        const sheet = context.workbook.worksheets.getItem(sheetName);
        await clearAllHighlights(context, sheet, lastAnchor, highlighted);
      });
    },
  };
}
