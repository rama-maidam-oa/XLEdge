# XLEdge AIPowered Stage 7 — QA & Cutover Support Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give QA and the developer what they need to run spec §7's regression on a real machine and decide the cutover:
a health-check script, a QA & cutover runbook, and a GLSense side-by-side checklist. The QA itself and the cutover are the
developer's (Excel and installs cannot be driven from this session).

**Architecture:** `AIPowered\XLEdge\tools\XLEdgeHealthCheck.ps1` reads the machine's XLEdge state the way the add-in does
(registry entries adxregistrator writes, the install root from the registered adxloader, `AddinCore\` as PathProvider lays it
out, the daily log `XLEdge_Logs_dd-MMM-yyyy.log`) and reports PASS/WARN/FAIL lines (or JSON). The runbook consolidates spec
§7 with every stage's Excel check items into one pass/fail matrix, plus rollback, uninstall and documented cutover steps.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§5 row 7, §7).

## Developer's decisions (2026-11-03)

- Deliverables: health-check script, QA & cutover runbook, side-by-side GLSense checks.
- Cutover is **documented only**; the developer decides when and how.
- The orbitbootapp `XLEdgeAddinController` is built by a separate engineer from `2026-11-03-xledge-addin-online-update-server.md`.

## Global Constraints

- The script only reads (no registry or file writes, no Excel automation); Windows PowerShell 5.1; exit code 1 when any FAIL.
- Registry it reads (HKCU): `Software\Microsoft\Office\Excel\Addins\XLEdge.AddinModule` (LoadBehavior 3 = load at start),
  `Software\Classes\CLSID\{80B0FB76-A5F4-41E0-B283-C175B7374B60}\InprocServer32` (default = the adxloader path → install root),
  `...\Addins\GLSense.AddinModule` (side by side, information only). `-RegistryRoot` (default `HKCU:\Software`) for tests.
- Paths: `<install root>\AddinCore\{Versions\, Manifest\, ReleaseHistory.json, SelectedRelease.txt}`; logs
  `%LOCALAPPDATA%\ORBIT\Excel_Logs\XLEdge_Logs\Logs\XLEdge_Logs_dd-MMM-yyyy.log` (`-DataRoot` overrides `%LOCALAPPDATA%\ORBIT\Excel_Logs`).
- Tests write only a unique temporary HKCU key (`Software\XLEdgeHealthCheckTests_<guid>`), removed in `finally`.
- 0 errors / 0 warnings; suite ×3; mutants; no Claude attribution; push `11.1.2-aipowered` at the end (as Stage 6).

## Review Focus

1. A machine with the old 11.1.2 add-in registered (InprocServer32 → its folder, no AddinCore) → reported clearly, not a crash.
2. Registered but the loader file missing (deleted folder) → FAIL naming the path.
3. SelectedRelease.txt naming a folder that is gone → FAIL (the next start falls back - say so).
4. Unsigned dev build → WARN, not FAIL.
5. No log yet today / log locked by Excel → reported, script still finishes.

---

### Task 1: The health-check script

**Files:** create `AIPowered\XLEdge\tools\XLEdgeHealthCheck.ps1`; test `XLEdge.Tests\HealthCheckTests.cs` (runs the script with
`powershell -NoProfile -ExecutionPolicy Bypass -File … -Json`, against a temp install root built from the shell's `bin\Debug`
DLLs, a temp data root with a log, and a temp registry root).

**Checks** (name → status rules):
- `Registration` — Addins key missing → FAIL; LoadBehavior ≠ 3 → WARN (`"LoadBehavior <n>: Excel does not load XLEdge at start (3 expected)"`); else PASS.
- `Loader` — InprocServer32 missing or its file missing → FAIL naming the path; else PASS with the path (its folder is the install root unless `-InstallRoot`).
- `Shell` — XLEdge.dll missing → FAIL; PASS with FileVersion; a separate `Signature` check: not Valid → WARN (`"not signed (a development build?)"`).
- `SharedAssemblies` — XLEdge.Contracts/Shared/Loader.Core missing → FAIL; FileVersion ≠ the shell's → WARN.
- `AddinCore` — no AddinCore folder → WARN (`"not set up: the old single-DLL XLEdge, or never started"`); ReleaseHistory.json unreadable → FAIL; PASS with the release list.
- `SelectedRelease` — file names a folder not under Versions → FAIL (`"… is gone: the next start falls back to the newest catalogued release"`); absent → PASS ("not chosen yet").
- `StagedRelease` — `Manifest\` holds a zip → PASS (`"staged: installed at the next Excel start"`).
- `Log` — today's log missing → WARN; else PASS with the count of `ERROR` lines today and the last `-LogErrors` (default 10) ones in the detail.
- `GLSense` — Addins\GLSense.AddinModule present → PASS with its LoadBehavior (side by side); absent → PASS ("not installed").

- [ ] Step 1: failing tests: `HealthCheck_HealthyInstall_HasNoFail` (every check present, statuses as above, exit 0);
  `HealthCheck_NotRegistered_Fails` (exit 1, Registration FAIL); `HealthCheck_LoaderMissing_FailsWithThePath`;
  `HealthCheck_SelectedReleaseGone_Fails`; `HealthCheck_LogErrors_AreCountedAndListed` (2 ERROR lines → "2", both texts);
  `HealthCheck_OldSingleDllInstall_WarnsNotSetUp` (no AddinCore → WARN, exit 0).
- [ ] Step 2: run → fail. Step 3: implement. Step 4: green + `bt.sh`. Step 5: commit `XLEdge AIPowered Stage 7: health-check script for QA machines`.

### Task 2: The QA & cutover runbook (with the GLSense side-by-side checks)

**Files:** create `docs/superpowers/qa/2026-11-03-xledge-aipowered-qa-runbook.md`:
- Setup: machine prerequisites; build/sign/MSI steps (Stage 5 release check); health check before and after each install.
- Installs: fresh install; upgrade over 11.1.2 (with `msiexec /l*v`); repair.
- Regression matrix (ID, area, steps, expected, pass/fail, notes) consolidating spec §7 and the Excel check items 1–11 in
  `AIPoweredXLEdge.md` (login manual/SSO/from GLSense, logout all panes, two windows, reports, title commands, XLSX,
  refresh/param refresh, control sheet, options, about, drilldown, attachments, pickers, sheet-delete cleanup, regional
  formats, broadcast, reload offline/online/release history, busy refusal, rollback, Excel close with a window open,
  workbook switch during a run, update extraction).
- GLSense side by side: both installed and loaded; login in each → the other follows; logout each → both out, one message;
  no-permission user; two Excel windows; reload of either add-in while the other is logged in; the log lines to look for in
  each add-in's log (XLEdge: `CallGLSense`, `LoginFromGLSense`, `NotifyGLSenseOfLogin`; GLSense: `LoginFromXLEdge`, `LogoutFromXLEdge`).
- Rollback: Release History (older core); back to 11.1.2 (uninstall 11.1.3 first - DetectNewerInstalledVersion blocks a
  downgrade; `AddinCore\` is removed by the uninstall); what the health check shows after each.
- Uninstall: what must be gone (files, `AddinCore\`, registration).
- Cutover (documented, not done): the decision points and steps (release branch, MSI build/sign, orbit-artifacts.xml
  `<version>`/`<release_date>` in both orbitbootapp branches, the orbitbootapp controller, Jira, the twin's TwinSubstitutions →
  OISRs for 11.1.2, retiring the 11.1.2 single-DLL line).
- [ ] Commit `XLEdge AIPowered Stage 7: QA & cutover runbook (regression matrix, GLSense side by side, rollback, uninstall, cutover)`.

### Task 3: Close the stage

- [ ] Fresh opus reviewer (script + runbook); one fix pass; `bt.sh` ×3; mutants on the script (LoadBehavior check, loader path,
  selected release check, error counting, exit code); `AIPoweredXLEdge.md`: Stage 7 done, how QA uses the runbook; commit;
  push `11.1.2-aipowered`.
