# AIPowered XLEdge — migration status and how to continue

Resume with: **"continue AIPoweredXLEdge.md"**. Read this file first, then the spec and the next stage's plan, and pick up at
"Next step" below. Last updated: 2026-11-03, after Stage 7 (QA & cutover support: health check, runbook), the
build check in the new folder layout, and OISR-22766 / OISR-22697 carried in from the 11.1.2 release of 03-Nov-2026.

## What this is

XLEdge C# 11.1.2 (`<checkout>\XLEdge`, the reference — never edited here) is rebuilt as a twin of GLSense AIPowered
11.1.2 under `<checkout>\AIPowered\XLEdge\`: a thin Add-in Express **shell** (`XLEdge`) loads a hot-reloadable **core**
(`XLEdge.Addin.Core`) into a separate AppDomain through .NET Remoting contracts (`XLEdge.Contracts`), with `XLEdge.Shared`,
`XLEdge.Loader.Core`, `XLEdge.Build` and the test project `XLEdge.Tests`. Hot reload, Release History / version switching, same
look as 11.1.2. Built from scratch ("no smell from GLSense"); 11.1.2's business code is moved into the core **verbatim**.

- Spec: `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§5 stage table, §6 deviations, §7 regression list)
- Plans (one per stage): `docs/superpowers/plans/2026-10-2*-xledge-aipowered-stage-*.md`
- Branch: **`11.1.3`** (pushed to origin) — the XLEdge 11.1.3 release line, GLSense-style version branch (developer's
  decision, 2026-11-03: 11.1.2 is frozen). Created from `11.1.2-aipowered` (Stages 0-7, kept on origin as history);
  `11.1.2` itself is untouched.
- **Folder layout** (developer's decision, 2026-11-03; `<checkout>` above = the 11.1.3 checkout):
  ```
  D:\SQLLite_Test\XLEdge\
    .bare\                         the git repository (core.bare = true; every checkout is a worktree)
    BusinessLayers\11.1.3\         branch 11.1.3 = <checkout>; solution AIPowered\XLEdge\XLEdge.sln (tests: XLEdge.Tests.sln)
    SingleLayer\Worktrees\11.1.2\  branch 11.1.2 (frozen classic XLEdge)
    SingleLayer\Worktrees\11.1.0\  branch 11.1.0
  ```
  `git --git-dir=D:/SQLLite_Test/XLEdge/.bare worktree list`; a new version = `git -C <checkout> worktree add
  ..\11.1.4 -b 11.1.4`. There is no 11.1.1 branch in this repo. Set up by `D:\SQLLite_Test\Move-XLEdgeLayout.ps1`.
  **A new checkout restores NuGet first** (packages are git-ignored; after the move `AIPowered\XLEdge\packages` held 10 of 17
  and the build failed with 884 errors): `MSBuild.exe AIPowered\XLEdge\XLEdge.Tests.sln -t:restore -p:RestorePackagesConfig=true` (the tests solution: it has every project).
  Verified 2026-11-03 in the new layout after the restore: Debug + Release 0 errors / 0 warnings, suite as below in both.
- **The classic single-DLL `XLEdge\` project is frozen at 11.1.2** on this branch: it is the parity baseline the twin's core
  is tested against (`CorePortParityTests`), not a product - never edit or version-bump it. 11.1.3 fixes go into the twin;
  a fix inside ported (verbatim) code is a counted `CorePortParityTests.TwinSubstitutions` entry, as for the 14+1 so far.

## Way of working (agreed with the developer)

- Stage by stage; the developer says **"proceed"** to start each stage. Inline execution (superpowers:executing-plans), TDD,
  one fresh **opus** reviewer for the whole stage at the end, then one fix pass. Each stage ends with: build Debug+Release with
  **0 errors / 0 warnings**, full suite green **×3**, mutation check (each fix broken on purpose must fail its test), commit.
- **Fix every finding immediately** (bugs, warnings, review minors) — nothing deferred unless it is a ruling with a reason.
- **No Claude attribution** in commits (no Co-Authored-By, no "Generated with Claude Code").
- Never read/print signing credentials, `secrets.env`, `.snk/.pfx`, `sign_file.cmd`, connection strings with passwords.
- Registering / unregistering the COM add-in (`adxregistrator`) is done **by the developer** (`!` prefix), never by Claude.
- Fixes to 11.1.2 production behaviour go to 11.1.2 first via OISR tickets (developer's call); the twin keeps 11.1.2's behaviour.
- Every decision taken on the developer's behalf is reported as a **"Ruling"** with what it costs if wrong.

### Tools (kept in `.superpowers\tools\`, git-ignored)
- **Solutions** (developer's requests, 03-Nov-2026): `XLEdge.sln` = the 6 product projects + the MSI project `OrbitXLEdge`
  (build and register in VS; no tests; the MSI builds in Release only), as GLSense AIPowered. `XLEdge.Tests.sln` = the
  product projects + `XLEdge.Tests`, MSBuild-only (what bt.sh / rt.sh build). `OrbitXLEdgeSetup.sln` is gone. The MSI
  project is the developer's Add-in Express setup project (`XLEdge\OrbitXLEdgeSetups\OrbitXLEdge.vdproj`) with the old
  project's settings ported in (11.1.2's UpgradeCode, the 11.1.3 Product/PackageCode, signing pre-build, core zip + manifest
  into `AddinCore\Manifest`, AddinCore cleanup on a real uninstall, no unregister during an upgrade, signed loaders from
  bin\Release, GAC System.IO.Compression 4.0.0.0 excluded); the wizard kept 11.1.2's component keys.
  `SetupProjectTests` pin all of it; `TestProject_LivesOnlyInTheTestsSolution_WhichHasEveryProjectOfXLEdgeSln` keeps the
  solutions in step.
- **Files Visual Studio changes on "Register ADX Project"** (ruling, 04-Nov-2026): `XLEdge.csproj` `RegisterForComInterop`
  true is a developer-machine setting - never committed (the repo keeps false; registration needs admin). The loaders
  `XLEdge\Loader\adxloader*.dll` are re-stamped with the shell's version resource (company, version): commit them when that
  information changes (it did on 04-Nov-2026: Add-in Express's placeholder 1.0.0.0 / "Default Company" -> Orbit Analytics
  11.1.3.0), as they are what the MSI installs. **Open (found 03-Nov-2026, 20:xx):** `CoreWindowsTests.Windows_OpenFromTheCore_AreListed_AndNothingSurvivesTheUnload`
  fails intermittently on the Release build only (2 of 3 alone; Debug 3 of 3 pass; it passed in every earlier run that day): an
  `HwndWrapper[CoreWindows;...]` outlives the 5 s wait after `AddinEntry.Shutdown` although every core window raised Closed. Not
  caused by the MSI/solution change (no code changed); to investigate (a window leak at shutdown, or only the test's wait). The
  scripts pass `-p:RegisterForComInterop=false`: VS's "Register ADX Project" turns it on in the shell csproj (a local change,
  never committed).
- **Logging** (developer's request, 03-Nov-2026): versions are shown as 11.1.3 (`XLEdge.Shared.VersionText.Display`; catalog,
  folder names and comparisons keep 11.1.3.0); the loader / core lifecycle lines are Debug, so with Debug logging off the log
  is the environment snapshot followed by `===== Start of XLEdge version 11.1.3 ...`; start and reload run in a log scope
  (`CoreHost.Scope`), so with Debug on their lines follow the marker as one block.
- `bash .superpowers/tools/bt.sh <tag>` — builds Debug and Release, runs the whole suite, prints errors/warnings/results.
- `bash .superpowers/tools/rt.sh <tag> <Test1,Test2|category:Excel>` — Debug build + only the named tests.
- Outputs in `.superpowers\tools\out\`. Per-stage workspace/ledger `.superpowers\sdd\<plan>\` is deleted when a stage ends.
- Python: use `py`. Write scripts with the Write tool (bash heredocs corrupt `\b \n \X` escapes).

## Stages done

| Stage | What | Last commit |
|---|---|---|
| 0a–0e | Solution and projects; contracts; Shared (Logger, PathProvider, ReleaseHistoryStore, VersionParser); Loader.Core (AddinDomainLoader, RemoteLoader, UpdateBootstrapper); shell host (AddinModule, ribbon, reload / source picker / progress / Release History windows), core packaging (zip + manifest) | `ffc5486` |
| 1 | WebView2 pane spike: core window re-parented into the ADX pane; reload with pane open; focus back to cells — **passed in Excel** | `a1c1291` |
| 2a | Core's own Excel RCW from an IDispatch pointer (contract 2); LogUtility / XLEdgeAppPaths / XLEdgeAppState / ribbon helper / Excel helpers / WPF facades | `e889041` |
| 2b | 22 helpers/models/utilities verbatim; golden values pin them to 11.1.2's built DLL; preferences | `5d17377` |
| 2c | Windows, theme, WPF-UI, ApiHelper in the core; session bag (contract 3) so session-only options survive a reload | `1eb92ed` |
| 2d | Report engine verbatim (ReportGenerator, ParamsControlSheetBuilder, XLEdgeParamsBuilder, DrilldownRequestBuilder, ReportParameterRequestSerializer, ProgressCoordinator); core `XLEdge.AddinModule` with 11.1.2's handlers (method-parity test); ribbon/Excel-event dispatch; `ExcelThreadContext` (continuations on Excel's thread, IsBusy while handlers/reports run) | `e7fa5aa` |
| 2e | 11.1.2's task pane (`ADXExcelTaskPane1`) as the core's pane content; Server Configuration + About; login, logout of all panes, Show/Hide, Param Refresh; contract 4 (pane of a given Excel window); title commands only from the instance's origin; login survives a reload (session value `XLEdge.Login`) | `fd0f2c9` |
| 22777 | 11.1.2 merged into this branch (reference = 11.1.2 at `eb78362`); OISR-22777 **Broadcast** ribbon button: shell ribbon regenerated (`make_designer.py` / `make_resx.py`), click forwarded to the core's verbatim `RibEdgeBroadcast_OnClick`, state lists (core helper + shell `RibbonControlIds`), `RibbonInitialize`; new test: every shell ribbon click forwards its own id | see `git log` |
| 2f | XLEdge ↔ GLSense login/logout sync. XLEdge core: 11.1.2's `InvokedFromGLSense` / `NotifyGLSenseOfLogin` / `NotifyGLSenseOfLogout` verbatim; `GetGLSenseAddinObject` returns `GLSenseAddin`, which asks the shell to make the COM call (contract 5); GLSense-started logoff counted as busy; a reload re-attaches GLSense. GLSense AIPowered (11.1.2 worktree `b693cf5`, `67f1e51`, `9d64975` - re-applied 04-Nov-2026 after a rebase had dropped the unpushed originals 8879b28/7489d20/2dbe4e9, caught by `GLSenseSourceTests`; 11.1.3 main checkout `02adf7d`, `b16b76a`): shell `GetGLCubeInformation` / `LogoutSession` → core `LoginFromXLEdge` / `LogoutFromXLEdge` | see `git log` |
| 4 | Remoting hardening (audit: leases and older-core compatibility already complete). A failing core start fails the load (rollback + error); Shutdown runs every step (`ShutdownSteps`); a stuck unload (5 s cap) is reported, the next core still loads; busy covers the pane's background work (`ExcelThreadContext.Track`, 15-min stale limit) and async-void WebView2 handlers (`Counted`); guard: the core never uses the shell's COM proxies. **Developer's request:** a report / refresh keeps writing to the workbook it started in (`RunWorkbookScope`, per run via AsyncLocal; verbatim files keep parity through `CorePortParityTests.TwinSubstitutions` — the twin now differs from 11.1.2 by those 14 changes until 11.1.2 gets the OISR) | `01637bd` |
| 5 | Packaging & MSI **11.1.3** (developer's decisions: version 11.1.3; signing through `sign_file.cmd` at `AIPowered\XLEdge\`; setup project drafted by Claude). `XLEdge.Build\sign.ps1` (Release only; `XLEDGE_SKIP_SIGNING=1` skips; missing `sign_file.cmd` fails the build); Contracts/Shared/Loader.Core sign in their post-builds, the core before its zip; Release copies the core zip to `SetupFiles\Release\InstallerPayload\OrbitXLEdge.zip`; `ReleaseNotes.txt` line 1 = manifest notes; setup project `XLEdge\OrbitXLEdgeSetups\OrbitXLEdge.vdproj` (from 11.1.2's: same UpgradeCode/folder/registration; generated by `.superpowers\tools\make_vdproj.py`) in its own `OrbitXLEdgeSetup.sln`; PreBuild signs the shell (**obj\Release** copy - that is what the Primary Output packages - and bin), the loaders, then `verify-signed.ps1` refuses any unsigned shipped file; uninstall removes `AddinCore\` (not on upgrade); the uninstall registration is skipped on upgrade | see `git log` |
| 6 | Online update, **client side** (developer: orbitbootapp untouched for now; server = spec `docs/superpowers/specs/2026-11-03-xledge-addin-online-update-server.md`, a copy-style `XLEdgeAddinController` with live tests). The core reports its login (`GetLoginInfo`); `XLEdge\Online\OnlineReleaseClient` calls `{LoginUrl}/rest/secure/xledge-addin/...` with the Bearer token (TLS 1.2/1.3, no proxy, as the core); the Reload picker's Online lists releases (loaded / already downloaded / needs a newer install / new), downloads every ticked one (cancellable), catalogues each through `ExtractAndCatalog` (and checks it is the build the index row describes), then reloads into the newest (GLSense's multi-release design, developer's decision) | see `git log` |
| 7 | QA & cutover **support** (developer: cutover documented only, decided later; the orbitbootapp controller by a separate engineer). `AIPowered\XLEdge\tools\XLEdgeHealthCheck.ps1` (read-only, PS 5.1, PASS/WARN/FAIL or `-Json`, exit 1 on FAIL: registration/LoadBehavior, registered loader → install root, shell version + signature, shared assemblies (the 11.1.2 single-DLL layout = WARN), catalogued releases, selected/staged release, today's ERROR lines, GLSense); `docs/superpowers/qa/2026-11-03-xledge-aipowered-qa-runbook.md` (installs, regression matrix R1-R25, GLSense side by side G1-G10 with the log lines - **Debug logging on in both add-ins**, rollback, uninstall, cutover steps) | see `git log` |

Test suite now: **386 tests — 379 pass, 7 Inconclusive** (the real-Excel tests, see "Machine notes").

## Next step (where we stopped)

**QA uses the runbook** `docs/superpowers/qa/2026-11-03-xledge-aipowered-qa-runbook.md` - it consolidates items 1 and 2
below (and spec §7) into one pass/fail matrix; run `tools\XLEdgeHealthCheck.ps1` before/after every install and attach its
output (and the XLEdge log) to any FAIL. Items 1-2 are kept as the developer's own quick list.

1. **Developer's manual check in Excel of 2d + 2e** (not done yet). Excel closed while building; shell + core in `bin\Debug`:
   1. Login → Server Configuration → instance → pane shows Home, ribbon LoggedIn.
   2. Run a report from the pane; Refresh Sheet / Refresh Book / Param Refresh (sheet, book).
   3. Second Excel window: Show/Hide acts on that window's pane only.
   4. Drilldown link, attachment link, "Goto Report Data"; Calendar / segment pickers; Options; About; Control Sheet;
      deleting a report sheet removes its `P_` sheet a moment later.
   5. Hide the pane → Reload Add-in → Show/Hide: pane back, logged in. Reload with the pane open: logs in by itself.
   6. Logout: every pane hides; next Login asks for credentials.
   7. Number/date formats on a non-English region (OISR-22663/22696); empty attachment cells (OISR-22766/22767).
   8. Broadcast (OISR-22777): disabled logged out, enabled after Login; click shows the instance's messages or
      "No broadcast messages available."; still enabled after a Reload Add-in while logged in.
   9. GLSense sync (2f) — GLSense AIPowered built and registered as usual (its post-build signs; Claude only compiled it
      unsigned in a temp export): both add-ins in one Excel:
      - Login in GLSense → XLEdge logged in, ribbon LoggedIn with the instance name, Broadcast enabled; if XLEdge's pane is
        open it navigates with the shared login.
      - Login in XLEdge by itself → GLSense `PartialLoggedIn` with its cubes (no Excel freeze while they load).
      - Logout in GLSense → XLEdge logged out too; GLSense shows **one** logout message and no "unexpected error".
      - Logout in XLEdge → GLSense LoggedOut.
      - GLSense login for a user without XLEdge permission → XLEdge stays logged out, its Login disabled.
      - GLSense not loaded → XLEdge login/logout work, nothing shown (log has a warning).
      - Logged in through GLSense → Reload Add-in (XLEdge) → Logout in XLEdge → GLSense logged out too.
   10. Stage 4: Reload while a report started from the pane runs, during an XLSX download, and while the pane is logging in
       → refused with the busy message. Switch to another workbook while a report loads, and while Refresh All runs →
       the report / refresh ends in the workbook it started in (that workbook comes back to the front when it writes).
       Two windows on one workbook ("Book1:2"): running a report keeps you in your window.
   11. Stage 6 (needs the ORBIT server with the XLEdge controller - not built yet): logged in, Reload Add-in → Online →
       Check for updates. Without the controller: "This XLEdge server does not offer XLEdge updates (yet)." With it (spec's
       publish steps): the list, Reload loads the newest ticked version (About), Cancel during a download stops it, log out →
       Online disabled with its reason. Check over HTTPS (the shell's own TLS settings).
   Fix whatever it finds (with a test) before Stage 7.
2. **Developer's release check of Stage 5** (Claude cannot sign or install):
   1. Copy `sign_file.cmd` to `AIPowered\XLEdge\` (git-ignored). Never set `XLEDGE_SKIP_SIGNING` for a release build.
   2. Open `AIPowered\XLEdge\XLEdge.sln` in VS once: check the setup project's Detected Dependencies (VS adds
      Contracts/Shared/Loader.Core entries - expected; nothing from the core: MahApps, WebView2, Wpf.Ui, Xaml.Behaviors,
      XLEdge.Addin.Core); discard VS's GUID/order drift in the vdproj (GLSense §65.3) but keep the component keys (pinned by
      `SetupProjectTests.SetupProject_KeepsTheKeysUpgradesDependOn`).
   3. Build: `devenv.com AIPowered\XLEdge\XLEdge.sln /Rebuild "Release"` (always **/Rebuild**: a plain Build can
      reuse unsigned outputs - `verify-signed.ps1` then fails it with XLEDGE0005). MSI: `AIPowered\XLEdge\XLEdge\OrbitXLEdgeSetups\Release\OrbitXLEdge.msi`;
      sign the `.msi` itself with signtool as today.
   4. On a machine with XLEdge 11.1.2 installed: install the 11.1.3 MSI with `msiexec /i ... /l*v upgrade.log` - one XLEdge in
      Add/Remove Programs, `XLEdge.dll` 11.1.3.0 and signed in `%LOCALAPPDATA%\Orbit Analytics\XLEdge`, `AddinCore\Manifest\`
      present before the first start, its zip gone after it (core extracted to `AddinCore\Versions\`; `manifest.json` stays),
      Excel loads XLEdge (About: shell 11.1.3.0, core 11.1.3.0). The log shows 11.1.2's `adxregistrator /uninstall` **running**
      during the removal (11.1.2's own package has no `NOT UPGRADINGPRODUCTCODE` guard - 11.1.3's guard applies from the
      next upgrade on), then 11.1.3's `/install` registering again - so XLEdge must load after the upgrade, twice.
      (The adxloader files stay 11.1.2's on an upgrade - same binary, unsigned; accepted.)
   5. Uninstall: the folder's `AddinCore\` is gone, XLEdge no longer in Excel's COM add-ins.
   6. Release day: `XLEdgeAppConstants.DefaultCommitDate` needs a `CorePortParityTests.TwinSubstitutions` entry (verbatim file).
3. Then: fix what QA finds (with a test); **Stage 3** only if it finds pane/login gaps; the **cutover** per the runbook §7
   when the developer decides (nothing of it is done); the orbitbootapp controller (spec above) by the separate engineer.
   Health check (after the Stage 7 review): reads every part of today's log (rollovers, `XLEdge_Error_*.log`), checks the
   loader for Excel's bitness (32-bit = WOW6432Node / adxloader.dll), and follows UpdateBootstrapper's rules for staged,
   refused and selected releases and an empty/unreadable catalog (runbook §0).

## Excel stays running after close (04-Nov-2026, from GLSense AIPowered)
Applied GLSense's `AIPowered\GLSense\docs\excel-stays-running-after-close.md`. XLEdge already had most of it: the core's own
Excel RCW (contract 2, released with FinalReleaseComObject at Shutdown; guard `TheCore_NeverUsesTheShellsComProxies`),
`ExcelThreadContext` around ribbon clicks, Excel events and the GLSense calls (Stage 2d), and a ribbon that resets on a new
core (no cube/ledger combos; LoggedOut resets the Login caption). **Found and fixed:** the pane's `SafeFireAndForget` (twin-only)
ran everything a page title starts - reports, multi-reports, Logs report, file downloads - and the WebView2 start-up on a pool
(MTA) thread through `Task.Run`, so a report run from the pane wrote the sheet from the MTA: the cross-apartment Excel
references that keep EXCEL.EXE alive. It now starts the work on Excel's thread under `ExcelThreadContext` (still tracked as
busy); the pane entry points (`CreateTaskPaneContent`, `CloseTaskPaneContent`, `OnTaskPaneShown`) run under it too. Safety
net: `CoreServices.ExcelApp` / `ExcelAppIfCreated` used off Excel's thread logs one WARN per calling method ("CoreServices:
Excel used off Excel's thread ... by <method>"). Tests: `PaneWork_StartsAndResumesOnExcelsThread_...`,
`PaneWork_StartedFromAPoolThread_StillRunsOnExcelsThread`, `ExcelUsedOffItsThread_IsLoggedOncePerCaller`. Trade-off (as GLSense):
the synchronous parts of a pane-started report now run on Excel's thread, as ribbon refreshes already did (e.g. the
synchronous image download of image columns).
**Developer's check in Excel** (close Excel after each; EXCEL.EXE must leave Task Manager within seconds, ~75 s with
Bitdefender): open/close; login/close; a report from the pane, then close (Save and Don't Save); Refresh Sheet/Book; a
drilldown; an attachment download; a Logs report; Reload Add-in then a report then close. Any "Excel used off Excel's thread"
WARN in the log names a path still to fix.

## OISR fixes carried into 11.1.3
**Standing rule (developer, 10-Nov-2026): every fix gets a same-issue sweep.** Name the bug's pattern, search every other
window/handler/code path here and in GLSense (both trees) for the same shape, fix the matches or list them as open with the
reason, and record it in the fix's entry ("Same-issue sweep:"). A fix without a recorded sweep is not finished.
- **11.1.2 release of 03-Nov-2026** (`5da5d37`, MSI in orbitbootapp `release/ORBIT11.1.2` `600efa7d08`; both tickets in QA
  Review) merged into 11.1.3 (`e52f964`), so the classic baseline (`XLEdge\`) = released 11.1.2; the twin core carries the
  same changes, so `CorePortParityTests` stay exact (no new TwinSubstitution). VB not ported (developer's decision).
- **OISR-22766** (Fusion attachment link: 400 / "Failed to download the attachment."): `AttachmentLinkHelper.TryParseAttachmentLink(raw,
  businessObjectId, …)` replaces the formula's `{businessObjectId}` placeholder with the report's id, as the web grid does;
  `ReportGenerator.AddAttachmentColumn` passes `reportMeta.BusinessObjectId`. `AddinModule.HandleAttachmentDownloadAsync` shows a
  "Downloading attachment" wait window and closes it (`CloseWaitWindowAsync`, added to `AddinModuleMethods`) before any message.
- **OISR-22697** (dates before 1900 as `####`): `XLEdgeValueFormatter.ToExcelDateValue` writes a pre-1900 date as text in the
  column's format (`ExcelDateFormatToDotNet`, leading apostrophe) and 1-Jan..28-Feb-1900 one serial lower (Excel's 1900 leap-year
  bug); `ReportGenerator`'s three write paths pass `ResolveColumnNumberFormat(column)`.
- Core `XLEdgeAppConstants.DefaultCommitDate` = 03-Nov-2026 (verbatim with 11.1.2; `DefaultVersion` stays the 11.1.3 substitution).
- Golden values regenerated from the released 11.1.2 `bin\Release\XLEdge.dll` (5da5d37): new cases for both tickets; one earlier
  value changed on purpose - th-TH `ToExcelCellValue("28.10.2026","DATE")` was the serial -152004 and is now the text
  `'28-Oct-1483` (the pinned th-TH defect below, now visible as text instead of `####`). `CoreHelperGoldenTests` asserts the QA cases.
- **OISR-22869** (Reports does nothing in a workbook opened after login; 11.1.3 only): 11.1.2's `XLEdgeLogin(true)` toggled
  ADX's `TaskPaneInstance`, which ADX had already made for every Excel window. The core's `TaskPanes.TaskPaneInstance` is null
  until a window's content is embedded, so the toggle returned without doing anything. A null pane now falls through to the
  create-and-show path. Test: `CorePaneTests.ShowHide_ActiveWindowWithoutAPane_ShowsANewPane` (fails without the fix).
  Same-issue sweep (pattern: active window's pane is null → silent return): every `GetPaneInstance()`/`TaskPaneInstance`
  caller. Fine as is: `InvokedFromGLSense`, `adxExcelAppEvents1_WorkbookActivate`, `ReportGenerator.ReleaseTaskPaneFocusToExcelAsync`
  (they only act on a pane that is already visible), `LaunchExcelPane` (creates one). **Open:** `RibEdgeParamRefreshBook_OnClick`
  and `RibEdgeParamRefresh_OnClick` return silently when the active window has no pane (e.g. Book2 opened after login).
  Creating the pane there isn't enough: the re-run is a script on the pane's loaded web page, which a new pane hasn't loaded
  yet. Needs the developer's decision (see the 10-Nov summary).
- **OISR-22777 QA (09-Nov)**, `XLEdgeMessageWindow` (the bug is in 11.1.2 too; not ported): long broadcast lines were hidden
  under the message scrollbar (the ScrollViewer's template draws it over the content; the text grid now has an 18px right
  margin), and the OK button was cut off at 125% scaling (`AutoFitOnce` measured the height with unlimited width; it now
  measures again at the final width). Checked by rendering the window from the built core (long message, before/after).
  Same-issue sweep: (a) overlaid scrollbar - it comes from the Wpf.Ui theme XLEdge loads for every window
  (`WpfUiBootstrapper`), so every plain `ScrollViewer` is affected. `XLEdgeGLAccountsWindow`'s segment list (scrolls past
  350px, ~8+ segments; the High boxes are at the right edge) got the same 18px right margin. `XLEdgeAbout` only has the
  card border under it - nothing hidden, left as is. GLSense no longer uses Wpf.Ui, so it isn't affected. (b) height measured
  before the final width - `XLEdgeMessageWindow` is the only window with its own auto-fit; GLSense's `GLMessageWindow` stays
  on `SizeToContent`, which re-measures at the real width.
- **OISR-22143 QA (09-Nov)**, report OISR-21508 after Book Refresh with DEPTNO = 10: the request and the rows were right
  (log: payload `deptno = 10`, every row 10), but Excel showed 12 rows against the web's 11. `ReportGenerator.ParseCsv` read
  the response with `StringReader.ReadLine`, so a quoted value containing a line break (ENAME "CRLF" typed with Enter) was
  split: `10,777,"CRLF` lost its SAL and `",2100` became a junk row (Record Count 12). It now appends the next line while
  the record has an odd number of quotes. Checked on the built DLL by reflection: the reported record, LF-only, escaped
  `""`, a comma inside quotes, an unclosed quote. Same-issue sweep: `ParseCsv` is XLEdge's only CSV parser (first run,
  Sheet/Book Refresh, Process screen Excel data); GLSense doesn't parse CSV. The legacy `11.1.3\XLEdge` project has the
  same code and was not changed.
- **OISR-22130 QA (09-Nov)**, GL Accounts Segment Selection: the scrollbar still covered the column header row ("Segment
  Name / Low / High"), since the header was inside the ScrollViewer. The header row now sits outside it (above the
  scrolling rows); header and rows both keep the 18px right margin so the columns line up. QA also attached `XLEdgeAbout`
  with the scrollbar over the cards' right edge and the progress bar: its content now has the same 18px right margin
  (supersedes "left as is" above). Checked by rendering the GL Accounts window from the built core (3 and 9 segments).
  Same-issue sweep: the remaining plain `ScrollViewer`s were checked in the 22777 sweep; only these two windows
  had content under the scrollbar.
  Issue 3 (wording): the Options checkbox already reads "Show segment selection window when selecting segments", but the
  ORACLE_GL_SEGMENT_VALUES cell's input/error messages (`ParamsControlSheetBuilder`) still said "...on selection"; both
  now use the Options wording (221/225 and 208/255 characters, Excel's limits). Existing Control Sheets keep the old text
  until they are created again. Issue 2 (dropdowns): not applicable - the LOVs depend on the responsibility and ledger.
- Both are deliberate differences from 11.1.2: `AddinModuleMethods_MatchXLEdge11_1_2` (XLEdgeLogin) and
  `PortedFiles_AreByteIdenticalToXLEdge11_1_2` (XLEdgeMessageWindow.xaml/.cs) report them; no TwinSubstitution added.
- **OISR-22874** (closing Excel with a core dialog open - e.g. the GL Accounts picker - showed "Thread was being aborted" and
  Office then disabled the add-in, so the ribbon was gone on the next start): `AddinBeginShutdown` ran inside the dialog's
  message loop, and `CoreHost.Shutdown` unloaded the core's AppDomain from a thread that still had core frames on its stack, so
  the runtime aborted it. `CoreHost.Shutdown` now asks `IsBusy` first; when busy it still runs the core's `Shutdown` but skips
  the unload (Excel is exiting). Shell change (needs an MSI). Test: `CoreHostTests.Shutdown_WhileTheCoreIsBusy_ShutsItDownWithoutUnloading`.
  The report-runner 500 in the same ticket ("Could not find Entity for id: 123") is a server error.
  Same-issue sweep (pattern: AppDomain unload while the current thread is inside a core call): `CoreHost.Reload` already refuses
  while busy; GLSense's `AddinModule_AddinBeginShutdown` had the same bug and is fixed (GLSense CLAUDE.md §103); its
  `ReloadAddinCore` can't run inside a core dialog. Nothing left open.
- **OISR-22872** (QA on GLSense's Reload / Version History windows; the developer asked for every change to be ported here):
  - Version dates are shown like the About window's date (`ReleaseDates.Display`, "dd-MMM-yyyy hh:mm tt", display only)
    in the history grid, the loaded-version line, the Online list and details, and the offline/no-update messages.
  - A 404 on the release index (`NoUpdatesHere`) or nothing new now logs "No updates found." and shows a "No Updates
    Found" box instead of an error line.
  - Version History has a read-only notes box for the selected version. The Source tooltips explain
    Install/Online/Offline. The Remove tooltip and confirm say it deletes this computer's copy only (the server is not
    changed, and Online lists it again).
  - Ribbon ScreenTip "XLEdge Version History". SuperTips and `CoreHost` errors no longer say "core version".
  - Tests updated: `OfflineReleaseScannerTests` (message text), `CoreHostTests` (error text), `RibbonParityTests` (ScreenTip).
  - Targeted run 73/73. Shell change (needs an MSI).
  - Follow-up (11-Nov-2026, developer): the NOTES column and the notes box were removed from Version History again
    (GLSense too); `ReleaseRow.Notes` dropped. Notes stay in ReleaseHistory.json.
- **OISR-22129** follow-up (QA 09-Nov: copy/paste still changed the Responsibility Operator/Value1 on the Parameters Control
  Sheet; 11.1.3 only, developer's call): the row's lock was data validation only (`SetupOracleRespIdValidation`), which stops
  typing but not a paste, Delete or fill; `Locked` does nothing on an unprotected sheet. Run sends Value1 as `ORACLE_RESP_ID`.
  - The shell now handles `SheetChange` and forwards it only for "Parameters Control Sheet" when the change reaches columns
    I:K. `ParamsControlSheetBuilder.RevertResponsibilityEdit` undoes a change that reaches I:K of the ORACLE_RESP_ID row
    (`Application.Undo` with events off) and shows "Responsibility edits are strictly prohibited". Skipped while
    `ShowOrRebuild` writes the sheet. A script/macro write can't be undone: logged, and the message asks to recreate the sheet.
  - Older cores log "no handler" for the new event. Shell change (needs an MSI).
  - Verified in Excel (Debug): a keyboard paste into J7 was undone (log "undid a change to J7"), J7 back to 591663.
  - Same-issue sweep (pattern: validation-only lock): the GL Accounts row's Value1 (`SetupGLSegmentValidation`) has the same
    paste gap - left open: the GL Accounts window writes that cell itself, so it needs a skip flag first. 11.1.2 not changed.
- **OISR-22887** (every core window slow the first time, fast after; host windows not affected): the core AppDomain was
  created with `LoaderOptimization.SingleDomain`, so .NET Framework/WPF assemblies loaded domain-bound there and were
  JIT-compiled instead of using their native images; each new window type hit new WPF code. `AddinDomainLoader` now uses
  `MultiDomainHost`: only GAC assemblies are shared; the core's own assemblies (release folder) still unload with the domain.
  - Measured (fresh Excel, About): first open 2.16 s -> 0.63 s; warm opens 0.53 s either way. Version History (host) 0.65 s.
  - Loader/CoreHost/CorePane tests 47/47. Host change (loader runs in the shell), needs an MSI.
- **OISR-22913** (Server Configuration: the Default tick moved to whichever row was clicked): the Default column was bound
  to `UrlInstance.IsSelected`, which `DgInstances_SelectionChanged` sets on the clicked row. The saved default
  (`IsDefault`, changed only by Set as Default) was never wrong; only the display was. The column now binds
  `IsDefault` one-way (the cell is read-only: `IsHitTestVisible=False`, and `DgInstances_BeginningEdit` cancels edits).
  `IsSelected` and the SelectionChanged loop are left in place, unused by the view. Core change only (no MSI).
  - 11.1.3 only (developer's call): `PortedFiles_AreByteIdenticalToXLEdge11_1_2` reports `XLEdgeServerConfiguration.xaml`
    as a deliberate difference; no TwinSubstitution added. 11.1.2 (`XLEdge\Views`) has the same binding, not changed.
  - Same-issue sweep (pattern: a checkbox column bound to selection state instead of the data it labels): no other
    `DataGridCheckBoxColumn` in the core binds `IsSelected`; the `Delete` re-sync workaround (`BtnDelete_Click`) now
    simply reloads a view that was already correct.
- **OISR-22914** (Server Configuration: an empty instance could be kept and made default but not deleted; developer: blank
  entries must not be kept, a blank row must be deletable, GO on a blank row warns): clicking into the grid's add-row
  committed an empty `UrlInstance`; `UpdateCachedConfiguration` copied it into the static `CachedConfiguration`, which
  `GetConfigurationSnapshots` reads before the XML, so it came back on every reopen in the session (never on disk:
  `SaveConfiguration` rejects it). `SetDefaultInstance` moved the in-memory default to it before Save failed, and
  `BtnDelete_Click` refused rows with no name and URL.
  - `UpdateCachedConfiguration` skips rows with no name and no URL. Set as Default refuses a row missing either.
    Delete accepts a blank row (same confirm; `CommitEdit(Row)` first, since Remove throws during AddNew).
  - GO: falls back to the default only when nothing is selected; a selected row without a URL (or the add-row, which
    is not a `UrlInstance` and used to fall through to the default) shows "Server address is empty for the selected row."
  - 11.1.3 only; `XLEdgeServerConfiguration.xaml.cs` is now also a deliberate difference reported by the parity test.
  - Same-issue sweep (pattern: placeholder/blank grid row treated as data): the other core grids are read-only lists
    (GL Accounts, drilldown picker, About); nothing else adds rows.
- **OISR-22915** (Server Configuration window changed size on Delete / Save / GO): every `UpdateStatus` called
  `AdjustWindowWidthToContent`, which measured the window at unlimited width (so the status text never wrapped) and set
  `Width` to that, clamped to 850-900 - each message length gave a different width. Ported unchanged from 11.1.2 (Stage 2e).
  The call and the method are removed; the window opens at a fixed `Width="850"` (what the short load message gave) and
  the status text wraps in its `*` column. Manual resizing within 850-900 still works.
  - 11.1.3 only; reported by the parity test with the 22913/22914 changes to the same view.
  - Same-issue sweep (pattern: window width recomputed on every content/status change): `DpiAwareWindow.FitToAvailableWorkArea`
    runs once on Loaded; `XLEdgeMessageWindow` fits once (`AutoFitOnce`); `XLEdgeDrilldownReports` sizes once from its list.
    No other window resizes after it is shown.

### Reopened 13-Nov-2026 (QA) - second round, 14-Nov-2026
- **OISR-22913 / 22914 / 22915 retest ran a pre-fix core.** QA's 13-Nov videos show messages that only the pre-fix code has
  ("Selected row is empty...", "The selected instance has no URL address configured."); the shipped 12-Nov MSI
  (`d18a24b158`, SHA256 59AC03...) contains the fixed core (checked by `msiexec /a` + strings in the DLL). VM (ExcelSmoke
  `Tests\UpgradeCore-XLEdge.ps1`): 11-Nov MSI + Excel, uninstall, 12-Nov MSI -> the 12-Nov core loads (the uninstall
  removes `AddinCore`, incl. `SelectedRelease.txt`). Installing over the 11-Nov MSI is refused (msiexec 1638, same
  ProductCode), so QA must have uninstalled. Most likely QA reopened Excel while the previous `EXCEL.EXE` was still
  running (it can linger 60-75 s, GLSense CLAUDE.md 71, or after a login, GLSense CLAUDE.md 111) - the new window then
  runs in the old process with the old core. No code defect found for 22913/22915.
- **OISR-22914** (QA: an empty instance should not be addable at all): clicking into the grid's add-row and leaving it
  still committed a blank `UrlInstance` that stayed listed for that window session (not saved, deletable). New
  `DgInstances_RowEditEnding`: a new row (`IEditableCollectionView.IsAddingNew`, it is `CurrentAddItem`) with neither a
  name nor a URL is cancelled (`CancelEdit(Row)` + `CancelNew`) instead of added. The VM test showed that alone was not
  enough: `CellEditEnding`'s "Instance Name is a mandatory field." cancelled the cell edit first, so the row commit never
  ran and the blank row stayed. `CellEditEnding` now drops a still-blank new row (nothing typed) instead of showing that
  error; a new row with partial data still gets the mandatory messages. Delete on the blank new row cancels the add
  (Remove throws during AddNew). `LoadConfiguration` no longer lists a blank entry from `ORBIT_URLS.xml` (an older build
  could have written one). Existing rows are never touched.
  - Same-issue sweep (pattern: blank add-row committed): the other core grids are read-only (see 22914 above).
- **OISR-22777** (QA01, 125%: the broadcast popup's bottom and OK button below the taskbar; the size was right): WPF's
  `CenterOwner` placed `XLEdgeMessageWindow` for its XAML size (340x200), then `AutoFitOnce` grew it right and down
  (only Width/Height set) and `DisableAutoSizing` turned off `DpiAwareWindow`'s re-centering and work-area clamp. Now
  `AutoFitOnce` caps `MaxHeight` at the work area of Excel's monitor and, after fitting, centers the window over Excel
  inside that work area (Win32 `MonitorFromWindow`/`GetMonitorInfo`; `System.Windows.Forms.Screen` was tried first and
  left SystemEvents' hidden `.NET-BroadcastEventWindow` alive after the core unloaded - caught by
  `CoreWindowsTests.Windows_OpenFromTheCore_AreListed_AndNothingSurvivesTheUnload`). A minimized Excel (rect at
  -32000) centers on the work area instead. Applies to every XLEdge message.
- **OISR-22915** (developer, 14-Nov: a wrapped status made the footer buttons taller): the status text and the buttons
  shared one footer card, so a two-line status grew the card and stretched the buttons. They are now two cards in the
  same row: the status card (`*`) grows with its text; the buttons card is `Auto`, `VerticalAlignment=Top`, own height.
  Layout only (`XLEdgeServerConfiguration.xaml`). Commit `4f2875a`.
  - Same-issue sweep (pattern: wrapping status sharing a card with buttons): GLSense's Server Configuration status is a
    single non-wrapping line; no other XLEdge window has a wrapping status beside buttons.
- VM (ExcelSmoke `Tests\Offline-XLEdge.ps1`, no login, Debug build 14-Nov 02:06): 22913 tick stays on the default, 22914
  no blank row after leaving the add-row, 22915 width 850 through Save / Delete->No / Set as Default, 22777 message
  centered on Excel (dx=0, dy=0) inside the work area - all pass. Login-dependent checks (22129 Control Sheet, 22937
  report run, 22777 with real broadcast messages) wait for QA02 (GLSense CLAUDE.md 111).
  - Same-issue sweep (pattern: window resized after placement with re-centering disabled): only `XLEdgeMessageWindow`
    sets `DisableAutoSizing`. GLSense's message window re-centers after resizing (GLSense CLAUDE.md 33).
- **OISR-22129** (QA 13-Nov: the "Responsibility edits are strictly prohibited" message shown ~8 times; video: an
  earlier, unrelated edit in I14 was undone too): `RevertResponsibilityEdit` called `Application.Undo` on every Change
  that touched I:K of the ORACLE_RESP_ID row without checking the values changed. Excel raises Change for J7 after its
  validation box is cancelled although nothing changed, so each Undo removed the next entry on Excel's undo stack (the
  user's other edits) and showed the message. Now `CreateControlSheet` remembers the row's I:K values in hidden sheet
  Names `XLEdge_RespLock_I/J/K` (`n:`/`s:`/`e:` = number/text/empty; they survive Reload and save); a Change with the
  values unchanged does nothing (no Undo, no message). A real change is that edit, so `Undo` reverts all of it (a paste
  over A:K too, and the validation it replaced); if the row still differs, the remembered values are written back (text
  with `NumberFormat "@"`, as the sheet is built, so a numeric-looking ID stays text) and I:K re-locked. One message per
  real edit. A sheet built by an older version (no Names) behaves as before (Undo).
  - Same-issue sweep (pattern: `Application.Undo` as a revert): `RevertResponsibilityEdit` was the only caller and the
    only SheetChange consumer in XLEdge; GLSense has none.
- **OISR-22937** (ribbon listed only some drilldown columns; not a regression - 11.1.2 and VB do the same): `UpdateTabLabel`
  looked only at the table's first data row, and links are added only to non-empty cells, so a column whose first value
  is blank was left out. It now asks each column of the whole data body for its links (one COM call per column); the
  caption is cut at 1000 characters with "..." (`CapRibbonCaption`, Office labels hold 1024).
  - Same-issue sweep (pattern: first data row taken as representative): `ReportGenerator.CaptureFirstRowFormulas` and the
    RefreshSync clean-up read the first row for table formula columns, which fill the whole column - correct there.
- **OISR-22887** (QA: still slow): measured in the VM (ExcelSmoke `Tests\Perf-XLEdge.ps1`, Debug build, no login):
  Server Configuration first open ~1.2 s, second ~0.04 s; About and Version History ~0.03-0.07 s. The ~1.2 s is WPF's
  one-time start-up in the core AppDomain (the first core window of a session). `AddinEntry.StartWpfUi` now posts
  `WarmUpWindows` at `ApplicationIdle`: builds a never-shown `XLEdgeMessageWindow` with a `DataGrid`, measures, closes.
  First Server Configuration open is now ~0.34 s. Not measured (needs a login): login page, report open/refresh,
  Reload - mostly server and WebView2 time. Left as is by decision: the strict online revocation check on each HTTPS
  call (`StrictCertificateValidator`).
  - Same-issue sweep: GLSense has no such first-window cost (checked for the 11-Nov fix).
- Targeted tests (CorePortParityTests, CoreHandlersTests, CoreHandlersExcelTests, CoreWindowsTests): same 4 failures as
  before these changes (the 3 known 11.1.2-parity tests, `UpdateTabLabel` now also listed, and
  `Broadcast_ShowsTheInstancesMessages_AsInXLEdge11_1_2`); 7 Excel tests skipped (Inconclusive on this machine).

## Pending stages

- **3** — remaining pane/login items not covered by 2e, if the Excel check finds any (spec §5 phase 3 was pulled into 2e).
- **5 — Packaging & MSI**: done (see Stages done; the developer's release check is open). Was: core post-build (sign Release, zip, manifest SHA256, ReleaseNotes), seed `InstallerPayload\OrbitXLEdge.zip`,
  `OrbitXLEdge.vdproj` on the GLSense pattern (shell as plain File, uninstall cleanup of `AddinCore\`), upgrade from today's MSI.
- **6 — Online update**: client done (see Stages done); the orbitbootapp `XLEdgeAddinController` is a written spec, to be built later.
- **7 — QA & cutover**: support done (see Stages done); running the QA and the cutover are the developer's.

## Contracts (`XLEdge.Contracts`, `ContractInfo.Version` = 5)
- 1: Stage 0b. 2: `IXLEdgeContext.GetExcelDispatch` (2a). 3: session bag `GetSessionValue/SetSessionValue` (2c).
  4: `IRibbonController` pane of an Excel window — `GetActiveTaskPaneWindow`, `Get/SetTaskPaneVisible`, `SetTaskPaneCaption(IntPtr,…)`,
  `Get/SetTaskPaneWidth` (2e). 5: `IXLEdgeContext.IsGLSenseConnected()`, `CallGLSense(method, args)` — the shell looks GLSense
  up and invokes it in its own AppDomain (2f); `GetGLSenseAddinInstance()` stays but the core no longer uses it.
- Additive only; every contract change ships with a new MSI. The loader refuses a core newer than the shell. `CoreHost.TryLoad`
  clears the session values for cores < 3 and drops `XLEdge.Login` for cores < 4.

## Key design decisions (details in the spec §6 and the plans)
- Verbatim ports guarded by `CorePortParityTests` (whole files; method-level for `AddinModule.cs` and `ADXExcelTaskPane1.cs`;
  only pack-URI and ADX type-name substitutions allowed). Golden values (`GoldenCases.cs` run against 11.1.2's built
  `XLEdge\bin\Debug\XLEdge.dll` by `Tools\GoldenValuesGenerator.cs`) pin helper behaviour.
- The shell forwards Excel events as names (sheet, workbook last, cell, hyperlink texts, shape); the core looks the COM objects up
  through its own RCW.
- Every core call into the shell's ribbon/panes runs on Excel's thread (`ExcelThreadContext.OnExcelThread`).
- Shell's build-order reference to the core has `ReferenceOutputAssembly=false` (both define `XLEdge.AddinModule`).
- Not ported: `LogHelper` (shell Logger owns the log + environment header), `XLEdgeRibbonReflectionHelper`, 11.1.2's pane
  BeforeTaskPaneShow / close-button handlers (the shell does both).
- Each core starts in 11.1.2's ribbon-load state (LoggedOut, or LoggedIn when the session login came back) but keeps the shell's
  Debug toggle. Temp report files are cleared at every core start.

## Machine notes
- **Real-Excel tests are Inconclusive here**: from any process other than Excel, QueryInterface for Excel's typed interfaces fails
  with `0x80029C4A TYPE_E_CANTLOADLIBRARY` (Click-to-Run Office 16.0.20527.20112 x64; the 32-bit Interface view points at TypeLib
  1.8, not registered). Inside Excel the add-in is unaffected. An **Office Quick Repair** would probably let the 6
  `[TestCategory("Excel")]` tests run — the developer's call. After that, run them ×3 and their mutants.
- Twin and today's XLEdge share assembly name / key / ProgId `XLEdge.AddinModule`: only one can be registered at a time.
- adxregistrator needed the ADX DLLs copied into `bin\Debug` (dev-only step).
- **Bitdefender quarantined `XLEdge.Addin.Core\post_build.ps1`** three times (30-Oct, 06-Nov, 13-Nov-2026;
  `CMD:Heur.BZC.PZQ.Boxter.985`, records in `C:\Program Files\Bitdefender\Endpoint Security\Quarantine`), so it showed as
  deleted in git. Cause, by elimination: its child `powershell -Command '[Reflection.Assembly]::LoadFile(...)'` that read
  `ContractInfo.Version` (sign.ps1, run the same `-ExecutionPolicy Bypass -File` way, was never flagged). Fixed 13-Nov-2026:
  the version is read as text from `XLEdge.Contracts\ContractInfo.cs`; `CorePackageTests` copies that file into its layout.
  Build + 15 `CorePackageTests` passed, nothing quarantined. An IT exclusion for `D:\SQLLite_Test\` was also requested.
  Same-issue sweep: no other `[Reflection.Assembly]::Load*`/`-EncodedCommand`/`IEX` in any build script under
  `D:\SQLLite_Test`. If it happens again: `git checkout -- XLEdge.Addin.Core/post_build.ps1` and check the quarantine.
- An untracked `.github/` folder (someone's Copilot instructions) exists in `<checkout>` — not created by this work, left alone.
- **GLSense here:** `D:\SQLLite_Test\GLSense` (main checkout) is on **11.1.3**; 11.1.2 is the worktree
  `D:\SQLLite_Test\worktrees\GLSense-11.1.2`. MSBuild cannot build GLSense from this session (its strong-name key
  `GLSense.Contracts.pfx` is password-protected and not installed; Claude never imports it) — Claude compiles a git export
  unsigned with no post-build events in `%TEMP%\claude\g2f` (short path: the XAML compiler fails past 260 characters).
  In the main checkout both `CLAUDE.md` files are **skip-worktree** (developer's setting): 11.1.3's fix-log section 91
  (the 2f change) is only a working-copy note there. `AddinEntry.cs` is CRLF in 11.1.2's index and LF in 11.1.3's, so
  cherry-picks between them conflict on every line — changes are applied to each branch separately.

## 11.1.2 defects found (pinned in the twin on purpose; tickets are the developer's call)
- `XLEdgeValueFormatter.InferDrilldownDataType` treats "12.5" / "1,234" as dates → drilldowns (and stored drilldown parameters)
  send ISO dates (e.g. "1,234" → 0234-01-01).
- Under th-TH, `ToExcelCellValue` reads a date that only the current-culture fallback parses (e.g. "28.10.2026") as a Buddhist-era
  year (1483 AD): a negative serial until 03-Nov-2026, since OISR-22697 the text `28-Oct-1483` (still the wrong year).
- Preferences file write race.
- Latent: a BigInteger serialized as a property object (2b); `ReportParameterRequestSerializer` quotes reportId and BigInteger
  values while the System.Text.Json path writes them as numbers (2d).
- Pane sends the login Bearer token to **every** origin it requests (`WebResourceRequested`, filter "*") — proposed OISR to limit
  it to the instance's origin.
- GLSense AIPowered lacked `GetGLCubeInformation` / `LogoutSession` (added in 2f on 11.1.2 and 11.1.3).
- GLSense AIPowered **11.1.2** `Logout()` touches ribbon controls from a pool thread after its `applogout` await — fixed on
  11.1.3 by OISR-22460, not on 11.1.2 (proposed OISR port; 2f's own new code marshals).
- (Stage 4 audit, proposed OISRs) a report writes into the workbook active after its HTTP awaits — switching workbooks during
  a run writes it into the other one (`BuildReportTable`, `BuildLogsSheet`, `GetParameterSheet`, Refresh All) — **fixed in
  the twin** (`CorePortParityTests.TwinSubstitutions` is the exact change set for 11.1.2); title-command reports run on a
  pool thread (cross-apartment COM); `ImageDownloadHelper` blocks on synchronous HTTP; the wait window is non-modal with no
  owner.

## Rulings still open / carried
- Broadcast messages are fetched again for each open pane after a reload (as 11.1.2 per pane load) — popup per pane per reload.
- An async-void started inside a WPF/WinForms continuation is not counted in IsBusy (none in today's code).
- Shell-only code (ADX panes: per-window members, re-attach after reload, active-window pane creation) is checked only in Excel.
- Stage 4 accepted risks: Excel closing during a report (no bounded wait, as GLSense); Excel's thread waits up to 5 s during
  an unload; `IsBusy` answers "not busy" when the call fails (old/broken core); background work stuck for 15+ minutes stops
  counting as busy; a failing newest core on Reload leaves no core (roll back through Release History).
- Stage 4 deferred minors: stuck-unload message only when the next load succeeds; cleanup after a failed start is one try;
  the busy message does not name background work; the COM-proxy guard only matches receivers named Context/context/_context;
  `WebView_CoreWebView2InitializationCompleted` not counted; a counted handler's escaping exception is logged by
  ExcelThreadContext instead of WinForms; a title-command run reads its workbook from a pool thread.
- Stage 5 deferred minors: the cleanup action (copied Windows 11 cmd.exe, exit code not ignored) could fail an uninstall on an
  older Windows - test on Windows 10; every Release build re-signs Contracts/Shared/Loader.Core/core (metered, as GLSense);
  sign.ps1 untested with spaced paths; test fixtures still use ShellVersion
  "11.1.2.0"; a core post-build failing before it writes leaves the previous manifest+payload pair.
- Stage 6 deferred minors: "newest" by release date only; an index row with a blank folderName fails as a 404; a history
  entry whose folder is gone cannot be re-downloaded from Online; an unreadable ReleaseHistory.json makes every row "New"
  silently; Mandatory/Notes not shown; RbOnline enabled during an Offline scan; classifier tests sit in OnlineReleaseClientTests.cs.
- Branch `11.1.2-aipowered` is pushed to origin (developer's decision, end of Stage 6; no PR); since 2026-11-03 work
  continues on `11.1.3` (see the top).
- GLSense 11.1.3's MSI very likely ships an **unsigned GLSense.dll**: its setup project's PreBuild signs bin\Release\GLSense.dll,
  but a Primary Output packages obj\Release (proven for the twin by making the two differ). Proposed GLSense fix: sign the obj
  copy too and verify the extracted MSI's DLL with Get-AuthenticodeSignature.
- Tools: a mutant script must `touch` a file it restores (`mv file.bak file` gives it the backup's older time and MSBuild keeps
  the mutant's DLL - this happened after Stage 2f, fixed with a full Debug Rebuild).
