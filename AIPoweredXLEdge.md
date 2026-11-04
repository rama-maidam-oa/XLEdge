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
  solutions in step. **Open (found 03-Nov-2026, 20:xx):** `CoreWindowsTests.Windows_OpenFromTheCore_AreListed_AndNothingSurvivesTheUnload`
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
| 2f | XLEdge ↔ GLSense login/logout sync. XLEdge core: 11.1.2's `InvokedFromGLSense` / `NotifyGLSenseOfLogin` / `NotifyGLSenseOfLogout` verbatim; `GetGLSenseAddinObject` returns `GLSenseAddin`, which asks the shell to make the COM call (contract 5); GLSense-started logoff counted as busy; a reload re-attaches GLSense. GLSense AIPowered (11.1.2 worktree `8879b28`, `7489d20`, `2dbe4e9`; 11.1.3 main checkout `02adf7d`, `b16b76a`): shell `GetGLCubeInformation` / `LogoutSession` → core `LoginFromXLEdge` / `LogoutFromXLEdge` | see `git log` |
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
