# AIPowered XLEdge — migration status and how to continue

Resume with: **"continue AIPoweredXLEdge.md"**. Read this file first, then the spec and the next stage's plan, and pick up at
"Next step" below. Last updated: 2026-11-03, after Stage 4 (remoting hardening; a run keeps its workbook).

## What this is

XLEdge C# 11.1.2 (`D:\SQLLite_Test\XLEdge\XLEdge`, the reference — never edited here) is rebuilt as a twin of GLSense AIPowered
11.1.2 under `D:\SQLLite_Test\XLEdge\AIPowered\XLEdge\`: a thin Add-in Express **shell** (`XLEdge`) loads a hot-reloadable **core**
(`XLEdge.Addin.Core`) into a separate AppDomain through .NET Remoting contracts (`XLEdge.Contracts`), with `XLEdge.Shared`,
`XLEdge.Loader.Core`, `XLEdge.Build` and the test project `XLEdge.Tests`. Hot reload, Release History / version switching, same
look as 11.1.2. Built from scratch ("no smell from GLSense"); 11.1.2's business code is moved into the core **verbatim**.

- Spec: `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§5 stage table, §6 deviations, §7 regression list)
- Plans (one per stage): `docs/superpowers/plans/2026-10-2*-xledge-aipowered-stage-*.md`
- Branch: **`11.1.2-aipowered`** in `D:\SQLLite_Test\XLEdge` — local commits only, **never pushed**.

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

Test suite now: **318 tests — 311 pass, 7 Inconclusive** (the real-Excel tests, see "Machine notes").

## Next step (where we stopped)

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
   Fix whatever it finds (with a test) before Stage 5.
2. Then **Stage 3** (only if the Excel check finds pane/login gaps) or **Stage 5** — on the developer's "proceed".

## Pending stages

- **3** — remaining pane/login items not covered by 2e, if the Excel check finds any (spec §5 phase 3 was pulled into 2e).
- **5 — Packaging & MSI**: core post-build (sign Release, zip, manifest SHA256, ReleaseNotes), seed `InstallerPayload\OrbitXLEdge.zip`,
  `OrbitXLEdge.vdproj` on the GLSense pattern (shell as plain File, uninstall cleanup of `AddinCore\`), upgrade from today's MSI.
- **6 — Online update**: orbitbootapp `XLEdgeAddinController` mirroring GLSense's.
- **7 — QA & cutover**: full regression (spec §7) side by side with GLSense, rollback, uninstall.

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
- An untracked `.github/` folder exists in `D:\SQLLite_Test\XLEdge` — not created by this work, left alone.
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
- Under th-TH, `ToExcelCellValue` gives a negative serial (Buddhist calendar).
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
- Tools: a mutant script must `touch` a file it restores (`mv file.bak file` gives it the backup's older time and MSBuild keeps
  the mutant's DLL - this happened after Stage 2f, fixed with a full Debug Rebuild).
