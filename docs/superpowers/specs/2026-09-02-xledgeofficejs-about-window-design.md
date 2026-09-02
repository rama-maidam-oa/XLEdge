# XLEdgeOfficeJS — About Window (Sub-Project 1, Screen 2) Design

**Status:** Approved for planning
**Parent effort:** Same feasibility exploration as the Server Configuration dialog (`2026-09-02-xledgeofficejs-server-config-design.md`) — this is the second screen ported into the same `XLEdgeOfficeJS` repo, reusing its established patterns (task pane → `displayDialogAsync` dialog, Fluent UI React v9, `localStorage`-backed data).

## Goal

Port XLEdge's `XLEdgeAbout` WPF window into the same Office.js repo as a second dialog, reusing the Server Configuration screen's existing server-instance data. The interesting feasibility question this screen adds beyond the first one: **can a browser-hosted add-in make the same real HTTP compatibility check the WPF app does (`GET {instance}/rest/public/orbit-version`), given CORS and mixed-content restrictions the WPF app's `HttpClient` never had?** This is deliberately attempted for real (not stubbed), because the failure mode itself is real feasibility data for the rewrite decision.

## Context: what it's replacing

`XLEdge\Views\XLEdgeAbout.xaml` (+ `.xaml.cs`): a borderless WPF window (same chrome pattern as the Server Configuration window) showing:
- A header (icon + "About Orbit XLEdge" + close button).
- A product-info panel: logo, version (`XLEdgeAppConstants.DefaultVersion`), a static description line, build date (`XLEdgeAppConstants.DefaultCommitDate`), and a support link (`https://www.orbitanalytics.com`) that opens in the default browser.
- A read-only, non-editable `DataGrid` of "Instance Compatibility": for every server instance already configured in the same XML file the Server Configuration window manages, it makes an HTTP GET to `{address}/rest/public/orbit-version`, parses the JSON response's `verionInfo.xlEdgeVersion`, and shows a green check or red X depending on whether it matches the app's own version. Per-instance errors are swallowed (logged, row still gets a red X) — the whole check never fails visibly.
- An indeterminate progress bar + "Checking X... (i/N)" status text while the checks run.
- A footer with copyright text and a close button.

## Decisions (resolved during brainstorming)

| Decision | Choice | Why |
|---|---|---|
| Compatibility check | Real `fetch()` per instance, not a stub | The CORS/mixed-content failure mode is itself useful feasibility evidence for the rewrite decision, not something to hide behind a stub |
| Version/build-date source | `package.json`'s `version` field (currently `0.0.1`, generator default — not bumped for this POC), build date as a static placeholder string | No real build pipeline exists yet to generate a real build date; matching that exactly isn't the point of this screen |
| Data source | Reuses `serverConfigStore.ts`'s existing `loadServerInstances()` — no new persistence | Same instances the user already configured via Server Configuration; this screen is read-only |

## Architecture

- New page, mirroring `src/serverconfig/`: `src/about/about.html`, `src/about/index.tsx`, `src/about/components/AboutDialog.tsx`. Registered in `webpack.config.js` exactly like `serverconfig` was (new `entry` block depending on `react`, new `HtmlWebpackPlugin`).
- The task pane (`App.tsx`) gets a second button, "Open About", opening `about.html` via `Office.context.ui.displayDialogAsync` — the same mechanism as "Open Server Configuration". This requires generalizing (or duplicating) the existing single-dialog open/message-handling logic in `App.tsx` so two independent dialogs can each be opened, tracked, and closed without interfering with each other (each needs its own `Office.Dialog` ref and its own `DialogMessageReceived`/`DialogEventReceived` handlers, keyed by which dialog is open).
- `AboutDialog.tsx`'s own `closeDialog()` follows the exact same `{ type: "close" }` / `messageParent` contract as `ServerConfigDialog.tsx`.

## Version constant

`package.json`'s version isn't natively importable as TypeScript without a tsconfig change (`resolveJsonModule` isn't currently enabled, and `@types/node` isn't installed for a `process.env` alternative). Use `webpack.DefinePlugin` instead — already a pattern this project's `webpack.config.js` uses (`webpack.ProvidePlugin` for `Promise`):

- Add to `webpack.config.js`'s `plugins` array: `new webpack.DefinePlugin({ __APP_VERSION__: JSON.stringify(require("./package.json").version) })`.
- Add a small ambient declaration file, `src/globals.d.ts`: `declare const __APP_VERSION__: string;` so TypeScript recognizes the global without needing `@types/node`.
- `AboutDialog.tsx` references `__APP_VERSION__` directly (webpack replaces it with the literal string at build time).

## Dialog UI components

- **Header**: icon (reuse `@fluentui/react-icons`'s `Info24Regular` or similar) + "About Orbit XLEdge" + Close.
- **Product info section**: existing generator logo asset (`assets/logo-filled.png`, already in the repo) + `Version: {__APP_VERSION__}` + a static description line ("Financial Analytics and Reporting Tool") + a static build-date placeholder ("Build Date: POC build") + a support link (`https://www.orbitanalytics.com`) rendered as a clickable `Link`/styled text that calls `window.open(url, "_blank", "noopener,noreferrer")`.
- **Instance Compatibility grid**: Fluent `Table`/`TableHeader`/`TableHeaderCell`/`TableBody`/`TableRow`/`TableCell` (same primitives as the Server Configuration grid, for consistency), columns: Instance (name + address) and Compatible (icon-only: `CheckmarkCircle24Filled` in green / `DismissCircle24Filled` in red, using `tokens.colorPaletteGreenForeground1`/`tokens.colorPaletteRedForeground1`), with a tooltip on the icon showing the failure reason when incompatible/failed.
- **Progress state**: while checks are running, show an indeterminate Fluent `ProgressBar` + "Checking {name}... ({i}/{N})" text, matching the legacy sequencing (one check at a time, not all in parallel — this also avoids hammering multiple servers simultaneously).
- **Footer**: copyright text ("Copyright © 2015 Orbit Analytics. All Rights Reserved.") + Close button.

## Compatibility check logic

For each instance loaded via `loadServerInstances()`, sequentially (matching the legacy one-at-a-time behavior):
1. `fetch(`${address}/rest/public/orbit-version`)`.
2. On network/CORS/mixed-content failure (the `fetch` promise itself rejects — this is exactly how browsers surface CORS and mixed-content blocks, no special-case detection needed): mark the row incompatible, tooltip = the caught error's message.
3. On a non-OK HTTP status or a response body that isn't valid JSON: mark incompatible, tooltip = status code or "invalid response".
4. On valid JSON: read `verionInfo.xlEdgeVersion` (case-insensitive key match, matching the legacy `StringComparison.OrdinalIgnoreCase` lookup), compare to `__APP_VERSION__`; mark compatible only on an exact match, incompatible otherwise (tooltip shows both versions on mismatch).
5. Update the "Checking X... (i/N)" progress text before each fetch starts, same as legacy.

No instance ever throws out of this loop — every failure mode above resolves to an incompatible row with a tooltip, matching the legacy "never show a global error, just per-row status" behavior.

## Testing / verification

- `npm run build` after each change (no automated test suite, same as Server Configuration — spec-approved).
- Manual pass in Excel: open About from the task pane, watch the progress sequence run through all configured instances, confirm each settles to a check or an X with a sensible tooltip (expect most/all real instances to fail via CORS or mixed-content — that's an expected, informative result, not a bug to chase), click the support link (confirm it opens in the default browser), Close.
- Record the actual compatibility-check outcome (which failure modes were hit, if any) as a new section in `FINDINGS.md` — this is real feasibility data about whether Office.js add-ins can make direct server calls the way the current VSTO app does, relevant to the rewrite decision beyond just this one screen.

## Out of scope (explicitly deferred)

- Any workaround for CORS/mixed-content failures (e.g., a proxy) — this screen's job is to observe and report the failure mode, not solve it. A real fix (if needed) belongs to a later sub-project once the extent of the problem is known.
- A real build pipeline generating an actual build date/version.
- Porting any other XLEdge window (Login, Options, GL Accounts, Drilldown) or the ribbon itself.
