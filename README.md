# ErrorGuard

> **Status:** Pre-alpha / v0.1. A free, open-source Excel add-in — not affiliated with, and not a companion product of, any other project by the same author.

ErrorGuard watches an Excel workbook for changes and flags, in real time as you type, three well-documented sources of spreadsheet error:

- **Duplicate rows** — exact full-row match (v1; key-column matching is a natural v1.1 extension).
- **Format inconsistencies** — mixed date formats, text where the rest of a column is numeric, inconsistent leading/trailing whitespace or casing.
- **Out-of-range / outlier values** — IQR or z-score statistical outliers, plus optional user-defined min/max or allow-lists.

All three are **deterministic rules and statistics — no AI, no network calls, no telemetry**. Nothing in your workbook ever leaves your machine. See [Privacy](#privacy--synthetic-data-only).

## Why this exists

Manual spreadsheet entry is one of the most common, best-documented sources of data error, in any domain. Independent surveys put the business-spreadsheet error rate at roughly 88–94%; manual data entry itself runs roughly 1–3% errors per field. The failure mode is always the same shape — a duplicate row, a format that silently drifted, an out-of-range value nobody noticed — whether it shows up in a nonprofit's donor spreadsheet, a school's grading roster, a small business's inventory sheet, or (at a very different scale) a bank's transaction ledger. ErrorGuard catches it at the point of capture, directly inside Excel, for any spreadsheet in any domain.

## Install / sideload

ErrorGuard isn't on AppSource yet (v1 is GitHub-only — see [Distribution](#distribution)), so you sideload the manifest directly.

### Excel Desktop (Windows/Mac)

```bash
git clone https://github.com/ChiaoAnChang/errorguard.git
cd errorguard
npm install
npm start
```

`npm start` builds the add-in, launches a local dev server, and sideloads it into Excel Desktop automatically (via `office-addin-debugging`). The first run may prompt you to trust a locally-generated dev certificate — accept it, that's what lets Excel load `https://localhost:3000` as a trusted source.

To stop: `npm stop`.

### Excel on the web

1. `npm run build` then `npm run dev-server` (or keep `npm start` running) to serve the add-in locally.
2. In Excel on the web, go to **Insert → Add-ins → Upload My Add-in** and upload `manifest.xml` from this repo.

See Microsoft's [sideloading guide](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins) for the general mechanics if either of the above doesn't match your Office build.

### Try it

Use [`test-data/sample_data.csv`](test-data/sample_data.csv) — a small, fully synthetic workbook built to trigger all three rules at once. See [`test-data/README.md`](test-data/README.md) for exactly what it exercises and a manual test checklist.

## How it works

1. Task pane registers `worksheet.onChanged` on the active sheet and debounces bursts of edits (typing, paste) into a single check ~800ms after the last change.
2. On each check, the current used range is read as plain values and handed to a pure, Office.js-free detection engine (`src/engine/`).
3. Flagged cells get a highlighted fill; the task pane lists every flag with a one-line reason and a dismiss control. A dismissed flag doesn't re-trigger until the underlying cell actually changes again.
4. A settings panel lets you turn each rule on/off. Zero configuration is required for a first-time user — sensible defaults are on by default.

## Architecture

```text
errorguard/
├── manifest.xml           # Office Add-in manifest
├── src/
│   ├── taskpane/           # Task pane UI (HTML/CSS/TS) — flag list, dismiss, rule toggles
│   ├── commands/            # Ribbon command entry point (scaffold default; unused by v1's UI)
│   ├── office/               # Office.js integration layer
│   │   ├── eventBinding.ts   # worksheet.onChanged registration + debouncing
│   │   └── highlighter.ts    # applies/clears cell fill on flagged cells
│   └── engine/                # PURE detection logic — zero Office.js imports
│       ├── duplicates.ts
│       ├── formatConsistency.ts
│       ├── outliers.ts
│       ├── index.ts           # runEngine() orchestrator
│       └── engine.test.ts     # unit tests, no live Excel instance needed
├── test-data/                # synthetic test workbook + manual sideload checklist
├── docs/screenshots/          # synthetic-data screenshots only
├── .github/workflows/ci.yml   # lint + type-check + unit tests + build, on every push
├── README.md
└── LICENSE
```

**The one rule that matters most:** `src/engine/` stays free of any Office.js dependency. It takes plain arrays in and returns flagged cells/reasons out — which is what makes it unit-testable without a live Excel instance, and what would let the same engine be reused elsewhere later (a CLI version, say) without rewriting the detection logic.

## Development

```bash
npm install       # installs via .npmrc's legacy-peer-deps=true — see note below
npm test           # runs src/engine's unit tests (Jest + ts-jest)
npx tsc --noEmit    # type-check
npm run lint         # eslint (via office-addin-lint) + prettier
npm run build         # production webpack build
```

**Why `.npmrc` sets `legacy-peer-deps=true`:** the Office Add-in generator's template pins `@babel/core@^8`, which several jest-ecosystem packages (`ts-jest`'s optional peer, a couple of `jest`'s own transitive deps) still declare a `<8` peer range against. `ts-jest` doesn't actually require `@babel/core` at runtime — it compiles via the TypeScript compiler directly — so this is a safe, deliberate override, not a real incompatibility. If a future dependency bump resolves the conflict, this can be removed.

`src/office/eventBinding.ts` and `src/office/highlighter.ts` are the one part of this codebase that genuinely needs a live Excel host to verify — Office.js has no meaningful test double for `worksheet.onChanged` or a live range. That's exactly why the detection logic itself lives in `src/engine/` instead, fully covered by `npm test` with zero Excel dependency. Manually sideload-test the office/ layer against `test-data/` before relying on it — see `test-data/README.md`.

## Privacy / synthetic data only

- **No telemetry, no network calls, by default.** Core detection (Section "How it works") runs 100% locally. Nothing about your workbook is ever sent anywhere unless you explicitly opt into the optional v1.1 AI layer described below — and even then, only the specific flagged rows, never the full sheet.
- **All test/demo data in this repo is synthetic**, clearly labeled as such — see [`test-data/README.md`](test-data/README.md). Never commit real personal or organizational data anywhere in this repo, including in screenshots.

## v1.1 (planned, not yet built): optional AI-assisted layer

Once the rule engine here is solid and has real sideload mileage, a natural v1.1 addition is an **optional, off-by-default** layer for fuzzier cases the deterministic rules can't easily catch (e.g. a likely-typo of a name that appears elsewhere in the column, or a cross-column logical inconsistency). Planned shape, not yet implemented:

- **Bring-your-own-key (BYOK).** You'd enter your own API key (OpenAI / Anthropic / Azure OpenAI) in settings; usage is billed to your account, never a shared backend.
- **Local-only key storage**, sent only to the provider you chose.
- **Minimal data exposure**: only the specific ambiguous rows already shortlisted by the rule engine, never the whole sheet — and the settings UI would say exactly that.
- **Never required.** Every v1 feature above keeps working with this fully disabled.

## Distribution

**v1: GitHub only**, open-source under Apache-2.0, sideload as described above — no app-store step, no paid infrastructure. Investigating a Microsoft AppSource/Partner Center listing is a possible later step (Partner Center's enrollment language suggests it may be oriented toward registered businesses rather than individuals — this needs checking directly with Microsoft before assuming it's available solo); not a blocker for v1.

### Production hosting (GitHub Pages)

Office Add-ins must be served over HTTPS even for sideloaded distribution, so the manifest's production build points at a GitHub Pages project site rather than `localhost`. [`.github/workflows/deploy-pages.yml`](.github/workflows/deploy-pages.yml) builds and deploys `dist/` on every push to `main`; `webpack.config.js`'s `urlProd` (`https://chiaoanchang.github.io/errorguard/`) is both the manifest's production source location and the build's asset `publicPath` (a GitHub Pages *project* site is served under `/errorguard/`, not domain root, so every asset URL needs that prefix or the browser 404s looking for them at the root).

One-time setup this repo still needs before that workflow can actually publish: **Settings → Pages → Source: “GitHub Actions”** (the workflow's `actions/deploy-pages` step will fail until Pages is enabled this way). The repo-root `manifest.xml` always points at `localhost` and is what you sideload locally per [Install / sideload](#install--sideload) above; a production manifest pointing at the Pages URL is generated at `dist/manifest.xml` by `npm run build`, for anyone who wants to sideload the hosted version instead of running a local dev server.

## License

Apache-2.0. See [LICENSE](LICENSE).
