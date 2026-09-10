/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

import type { CellFlag, EngineOptions } from "../engine";
import { watchWorksheet } from "../office/eventBinding";
import type { WatcherHandle } from "../office/eventBinding";

const RULE_CHECKBOX_IDS = {
  duplicates: "rule-duplicates",
  format: "rule-format",
  outlier: "rule-outlier",
} as const;

// Dismissed flag ids for the current session. Intentionally in-memory only
// (not persisted to Office roaming settings) for v1 — a flag id encodes its
// rule + cell location, so it naturally "comes back" once the underlying
// cell changes again; see src/engine/index.ts's id scheme.
const dismissed = new Set<string>();
let latestFlags: CellFlag[] = [];
let watcher: WatcherHandle | undefined;

function readEngineOptions(): EngineOptions {
  const duplicatesOn = (document.getElementById(RULE_CHECKBOX_IDS.duplicates) as HTMLInputElement)
    .checked;
  const formatOn = (document.getElementById(RULE_CHECKBOX_IDS.format) as HTMLInputElement).checked;
  const outlierOn = (document.getElementById(RULE_CHECKBOX_IDS.outlier) as HTMLInputElement)
    .checked;
  return {
    ignoreRows: [0], // v1 assumes row 1 of the used range is a header row.
    duplicates: duplicatesOn,
    format: formatOn,
    outlier: outlierOn,
  };
}

function ruleLabel(ruleId: CellFlag["ruleId"]): string {
  switch (ruleId) {
    case "duplicate-row":
      return "Duplicate";
    case "format-inconsistency":
      return "Format";
    case "outlier":
      return "Outlier";
  }
}

function cellLocation(flag: CellFlag): string {
  const rowLabel = `Row ${flag.row + 2}`; // +1 for 0-index, +1 for the assumed header row.
  return flag.col === undefined ? rowLabel : `${rowLabel}, Col ${flag.col + 1}`;
}

function renderFlags(flags: CellFlag[]): void {
  latestFlags = flags;
  const list = document.getElementById("flag-list")!;
  const emptyState = document.getElementById("empty-state")!;
  const statusText = document.getElementById("status-text")!;

  const visibleCount = flags.filter((f) => !dismissed.has(f.id)).length;
  statusText.textContent =
    flags.length === 0
      ? "No issues found"
      : `${visibleCount} issue${visibleCount === 1 ? "" : "s"} found` +
        (flags.length !== visibleCount ? ` (${flags.length - visibleCount} dismissed)` : "");

  list.innerHTML = "";
  emptyState.style.display = flags.length === 0 ? "block" : "none";

  for (const flag of flags) {
    const item = document.createElement("li");
    item.className = "eg-flag-item" + (dismissed.has(flag.id) ? " dismissed" : "");

    const badge = document.createElement("span");
    badge.className = "eg-flag-badge";
    badge.textContent = ruleLabel(flag.ruleId);

    const body = document.createElement("div");
    body.className = "eg-flag-body";
    const location = document.createElement("div");
    location.className = "eg-flag-location";
    location.textContent = cellLocation(flag);
    const reason = document.createElement("div");
    reason.className = "eg-flag-reason";
    reason.textContent = flag.reason;
    body.append(location, reason);

    const dismissButton = document.createElement("button");
    dismissButton.className = "eg-dismiss-button";
    dismissButton.textContent = dismissed.has(flag.id) ? "Restore" : "Dismiss";
    dismissButton.onclick = () => {
      if (dismissed.has(flag.id)) {
        dismissed.delete(flag.id);
      } else {
        dismissed.add(flag.id);
      }
      renderFlags(latestFlags);
      watcher?.recheck().catch((error) => console.error("ErrorGuard: recheck failed", error));
    };

    item.append(badge, body, dismissButton);
    list.appendChild(item);
  }
}

function isDismissed(flagId: string): boolean {
  return dismissed.has(flagId);
}

async function startWatching(): Promise<void> {
  watcher = await watchWorksheet("Sheet1", renderFlags, readEngineOptions, isDismissed);
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  document.getElementById("sideload-msg")!.style.display = "none";
  document.getElementById("app-body")!.style.display = "flex";

  const rulesToggle = document.getElementById("rules-toggle")!;
  const rulesPanel = document.getElementById("rules-panel")!;
  rulesToggle.onclick = () => {
    const expanded = rulesPanel.style.display !== "none";
    rulesPanel.style.display = expanded ? "none" : "flex";
    rulesToggle.setAttribute("aria-expanded", String(!expanded));
  };

  for (const id of Object.values(RULE_CHECKBOX_IDS)) {
    document.getElementById(id)!.addEventListener("change", () => {
      watcher?.recheck().catch((error) => console.error("ErrorGuard: recheck failed", error));
    });
  }

  document.getElementById("recheck-button")!.onclick = () => {
    watcher?.recheck().catch((error) => console.error("ErrorGuard: recheck failed", error));
  };

  startWatching().catch((error) =>
    console.error("ErrorGuard: failed to start watching the sheet", error)
  );
});
