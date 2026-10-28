# XLEdge AIPowered: rebuild XLEdge on the GLSense AIPowered 11.1.2 architecture

Date: 2026-10-28 · Source of truth for the architecture: `D:\SQLLite_Test\GLSense\AIPowered\GLSense` (branch `11.1.2`)
· Source of the logic: `D:\SQLLite_Test\XLEdge\XLEdge` (branch `11.1.2`) · Status: draft for review

## 1. Intent

Recreate XLEdge as a twin of GLSense AIPowered: the same solution layout, the same shell / contracts /
shared / loader / core split, the same hot reload, Release History and version switching, the same
reload windows and update flow. In practice: start from the GLSense AIPowered skeleton, rename it to
XLEdge, strip what is GLSense-only, and pour the existing XLEdge C# logic into the core.

After this, a fix to either add-in ships the same way: build the core → zip + `manifest.json` → user
reloads (offline folder, online from orbitbootapp, or picks a version in Release History). The MSI is
only needed when the shell changes.

**Build approach (developer decision, 2026-10-28): new projects written from scratch, not copied.**
The XLEdge projects are created fresh with the same project types, settings and responsibilities as
GLSense AIPowered, and every class is written with XLEdge names. GLSense is the *reference design*,
not a source to copy — so no GLSense names, paths, comments or dead code leak in. To avoid losing
GLSense's hard-won fixes, each stage works from a **parity checklist** (behaviours + lessons pulled
from the GLSense code and its CLAUDE.md / PORTING_GUIDE.md) and ends with a "no GLSense smell" check:
a search for `GLSense`, `GL` prefixes, GLSense paths/GUIDs/mutex names in `AIPowered\XLEdge\`.

### Goals
- Project-for-project and folder-for-folder parity with GLSense AIPowered, so a developer who knows
  one knows the other, and a loader/reload fix can be copied between them.
- Same user experience for updates: Reload button, source picker (offline/online), progress window,
  Release History browser, rollback to any extracted version, without restarting Excel.
- XLEdge behaviour unchanged for users: ribbon, task pane login, reports, refresh, parameters,
  drilldown, attachments, GLSense↔XLEdge login sync, formats.

### Non-goals
- No redesign of XLEdge's own screens (task pane web UI, WPF windows keep their current look).
- No MVVM rewrite of XLEdge views; they move as they are.
- No shared runtime library between GLSense and XLEdge (two copies of the same design, XLEdge-named).
- No change to the VB xledge line.

### Anti-goals
- Diverging from GLSense's architecture "because XLEdge is different" without writing it down (§6).
- Excel left orphaned on close, locked DLLs blocking the next update, cryptic remoting errors to users.
- A swapped core reporting the old version in About / logs.

## 2. Target layout (mirrors GLSense)

```
D:\SQLLite_Test\XLEdge\
├─ AIPowered\XLEdge\                      ← NEW, mirrors AIPowered\GLSense\
│  ├─ XLEdge.sln
│  ├─ XLEdgeSharedVersion.cs              ← one version for all projects (like GLSenseSharedVersion.cs)
│  ├─ CLAUDE.md, PORTING_GUIDE.md         ← fix log + porting rules (copied from GLSense, XLEdge-specific edits)
│  ├─ sign_file.cmd                       ← copied as-is (never read/printed)
│  ├─ XLEdge\                             ← shell: ADX AddinModule, task pane host, RibbonController,
│  │                                         XLEdgeContext, GlobalsEx, RibbonControlIds, Views\ (reload windows),
│  │                                         Loader\ (adxloader), Interops\, OrbitXLEdgeSetups\OrbitXLEdge.vdproj, post_build.cmd
│  ├─ XLEdge.Contracts\                   ← IXLEdgeAddin, IXLEdgeContext, IRibbonController, ILogger, IPathProvider, LoginInfo, VersionInfo
│  ├─ XLEdge.Shared\                      ← Logger, PathProvider, ReleaseEntry, ReleaseHistoryStore, VersionParser
│  ├─ XLEdge.Loader.Core\                 ← AddinDomainLoader, RemoteLoader, ResolvedRelease, UpdateBootstrapper
│  ├─ XLEdge.Addin.Core\                  ← AddinEntry + all XLEdge logic and WPF; post_build.cmd zips + manifest
│  └─ XLEdge.Build\                       ← empty "build all" project, post_build summary
├─ XLEdge\  (existing project)            ← stays as the reference/monolith (role of GLSense FinalWorkingCode)
└─ MIGRATION_STATUS.md, docs\
```

Existing `XLEdge\` keeps receiving 11.1.2 fixes until cutover; each fix is re-applied in
`AIPowered\XLEdge\XLEdge.Addin.Core` (same discipline GLSense uses for its two trees).

### Project mapping

| GLSense AIPowered | XLEdge AIPowered | Reference design in GLSense (re-implemented, not copied) | Fill from XLEdge |
|---|---|---|---|
| `GLSense` (shell) | `XLEdge` | AddinModule skeleton (OnRibbonLoaded load, reload, shutdown), `GLSenseContext`, `RibbonController`, `GlobalsEx`, `RibbonControlIds`, Views `GLReloadSourcePicker`, `GLReleaseHistoryBrowser`, `GLReloadProgressWindow`, `OnlineReleaseClassifier/Row`, `WindowChromeHelper`, `post_build.cmd`, vdproj pattern | Ribbon layout + 14 buttons, Excel event list, task pane, `InvokedFromGLSense` / `GetGLSenseAddinObject` COM entry points, ProgId/GUID/snk (must stay `XLEdge.AddinModule`, PKT `cd5e60a954374975`) |
| `GLSense.Contracts` | `XLEdge.Contracts` | all files, renamed | XLEdge members (§3); drop `UdfSentinels`, `ExecuteUdf` |
| `GLSense.Shared` | `XLEdge.Shared` | all files, renamed | paths `Excel_Logs\XLEdge_Logs\…`, log file name `XLEdge_Logs_dd-MMM-yyyy.log` (as today) |
| `GLSense.Loader.Core` | `XLEdge.Loader.Core` | all files, renamed | — |
| `GLSense.Addin.Core` | `XLEdge.Addin.Core` | `AddinEntry` shape, `Infrastructure\ServiceLocator`, `Utilities\WpfAppManager`, `Views\BaseWindow` + pane host pattern, csproj/post_build/packaging | everything from XLEdge `Helpers`, `Models`, `Views`, `Utilities`, `Converters`, `Themes`, `Images`, `XLEdgeAppState/Paths/Constants`, `XLApp` |
| `GLSense.Build` | `XLEdge.Build` | as-is, renamed | — |
| — (not copied) | | `GLSense.LocalUpdateHost` (dead), `UpdateManager.cs` (unused), `GLSenseExcelFunctions` (UDF/XLL), SQLite + `sqlite_postbuild.cmd`, `RibbonComboPrimer` (XLEdge ribbon has no combos), cube/GL-specific views | |

## 3. Contracts (GLSense members, XLEdge names)

`IXLEdgeAddin` (host → core): `Initialize(IXLEdgeContext)`, `OnRibbonAction(string action, object parameter)`,
`bool OnExcelEvent(string eventName, object[] args)`, `Shutdown()`, pane bridge
`IntPtr CreateTaskPaneContent()`, `RelaunchTaskPane(bool)`, `ResetTaskPaneReference()`, `CloseTaskPaneContent()`,
`LoginInfo GetLoginInfo()`, `string[] GetOpenWindowTitles()` — plus XLEdge-only
`InvokedFromGLSense(name, url, token, user, hasPermission)` and `bool IsBusy()` (report running → block reload).

`IXLEdgeContext` (core reads): `ExcelApp`, `ExcelHandle`, `Logger`, `Paths`, `RibbonController`, `DebugMode`,
`Version`, `ReleaseDate`, `ActiveFolderName`, `AllVersions`, `SetRibbonController`, and
`GetGLSenseAddinInstance()` (mirror of GLSense's `GetEdgeAddinInstance()`).

`IRibbonController`, `ILogger`, `IPathProvider`, `LoginInfo`, `VersionInfo`: same members as GLSense.

Events forwarded (primitives only): SheetSelectionChange, SheetActivate, WorkbookActivate,
SheetFollowHyperlink (sheet, cell address, ScreenTip, SubAddress, text), SheetBeforeDelete.

## 4. Task pane — same pattern as GLSense, with an early spike

GLSense: the core creates a borderless off-screen WPF window with the pane content and returns its HWND;
the shell's `ADXExcelTaskPane` re-parents it (`SetParent`, `WS_CHILD`, `MoveWindow`).

XLEdge's pane content is a WebView2 that hosts the whole web UI (login, report list, parameters), with
`DocumentTitleChanged` as the command bus, `SourceChanged` for login/logout and cookie scrape,
`WebResourceRequested` for the Bearer header. In the twin, all of that lives in the core inside the
re-parented window (WPF `WebView2` control in a `BaseWindow`-style host) — so login/command logic is
hot-swappable too.

Risk: XLEdge moved its WebView2 from WPF/ElementHost to WinForms on 2026-08-20 because of airspace/trim
bugs. A re-parented top-level window is a different setup (no ElementHost), but it is unproven with
WebView2 + DPI + multiple Excel windows. **Phase 1 is a 2-day spike** to prove it. Fallback if it fails:
WebView2 stays in the shell pane and forwards strings to the core (`OnWebCommand`, `OnWebSourceChanged`,
`IPaneHost`), recorded as a deviation in §6.

## 5. Plan and effort

| Phase | Work | Days |
|---|---|---|
| 0 Foundation (from scratch) | 0a solution + empty projects with GLSense's project types/settings; 0b `XLEdge.Contracts`; 0c `XLEdge.Shared` (Logger, PathProvider, ReleaseHistoryStore, VersionParser); 0d `XLEdge.Loader.Core` (AddinDomainLoader, RemoteLoader, UpdateBootstrapper); 0e shell host (AddinModule skeleton with XLEdge GUID/ProgId/snk, context, RibbonController, reload/source-picker/progress/Release History windows) + empty `AddinEntry`. Each sub-stage against its parity checklist; ends with Excel loading an empty core and Reload/Release History working with a dummy zip | 5–7 |
| 1 Pane spike | WebView2 inside core window re-parented into the ADX pane: login page, resize, DPI, two Excel windows, focus to Excel, reload while pane open | 2 |
| 2 Port logic | Move XLEdge Helpers/Models/Views/Utilities/Themes into Addin.Core; replace `AddinModule.CurrentInstance` (12 sites), static ribbon access, ADX-typed fields with `ServiceLocator` / context; logger via `ILogger` proxy behind existing `LogUtility` API; move ~2,400 lines of handler logic from `AddinModule` into `AddinEntry` + handler classes; version constants into core | 6–8 |
| 3 Pane + login | Port `ADXExcelTaskPane1` logic (WebView2 env, title command bus, cookies/token, GLSense sync, logout all panes) into the core pane host | 3–4 |
| 4 Remoting hardening | Infinite leases, cached Workbook/Sheet in async chains, fire-and-forget entry points, `Shutdown` (WPF dispatcher, `_deleteTimer`, running reports, COM), 5 s unload cap, `IsBusy` reload guard | 3–4 |
| 5 Packaging & MSI | Core post_build (sign Release, zip, `manifest.json` SHA256, `ReleaseNotes.txt`), seed `InstallerPayload\OrbitXLEdge.zip`, `OrbitXLEdge.vdproj` on GLSense pattern (shell as plain File, uninstall cleanup of `AddinCore\`), upgrade from current single-DLL MSI | 2–3 |
| 6 Online update | orbitbootapp `XLEdgeAddinController` mirroring `GLSenseAddinController` (`/rest/secure/xledge-addin/manifests`, `…/versions/{folder}/manifest|zip`, `download-artifacts\addons\xledge\updates\`); ideally generalise the GLSense controller with a product segment | 2–3 |
| 7 QA & cutover | Full regression (§7) on the twin, side-by-side with GLSense in one Excel, rollback, uninstall; switch the release line to `AIPowered\XLEdge` | 5–8 |

**Total: ~28–39 developer days (6–8 weeks, one developer with AI assist), QA overlapping phases 5–7.**
Writing the foundation from scratch adds ~3–4 days over copying (about 3k lines of infrastructure:
contracts ~450, shared ~1.5k, loader ~950, shell host + reload windows); the main cost stays in
phase 2 (carving `AddinModule`) and phases 1/3 (the WebView2 pane in the core).

## 6. Deviations from GLSense (deliberate, my calls — veto any)
1. **Contract assembly versions frozen** at `1.0.0.0` (`XLEdgeContractsVersion.cs`; file version follows
   the product). GLSense stamps them from the shared version file; a core-only version bump could break
   the host/core type identity. Recommend back-porting to GLSense.
   Contracts, Shared and Loader.Core still **ship inside the core zip** (the core's AppDomain resolves
   from its release folder, so it needs them there). Because the identity no longer changes, a core built
   against different contracts would load silently and fail only at call time. So:
   - contract changes are additive only and always ship with a new MSI (the shell must know them too);
   - **Stage 0d:** before creating the AppDomain, the loader compares the release folder's
     `XLEdge.Contracts.dll` (and `XLEdge.Shared.dll`) with the shell's copies by SHA256 and refuses
     the release with a clear log line and message if they differ.
2. **`AppDomainSetup.ApplicationName` set**, so `CachePath` (shadow cache) is honoured.
3. **Shell referenced as a plain File in the vdproj** from the start (GLSense CLAUDE.md §67/§68 lesson).
4. No UDF/XLL module, no SQLite, no ribbon combo primer — XLEdge doesn't use them.
5. `InvokedFromGLSense` stays a public method on the shell's COM-visible `AddinModule` and forwards
   to the core (GLSense calls it by ProgId through `COMAddIns`).
6. Task pane fallback (§4) only if the spike fails.

7. Intended setting differences found in the Stage 0a review: `LangVersion latest` on Loader.Core and
   `Deterministic true` on the shell (GLSense has neither); the shell's `app.manifest` omits GLSense's
   `dpiAware` block (GLSense's own comment says it has no effect inside a DLL — do not "restore" it).

Open items for Stage 0e (from the Stage 0a review):
- The twin shell has the same assembly name, key and ProgId `XLEdge.AddinModule` as the existing add-in, so
  only one of them can be registered on a machine at a time. Dev boxes unregister one before registering
  the other.
- Loader DLL name: `adxloader.XLEdge.dll` (GLSense pattern) vs the existing `adxloader.dll`.
- `stdole.dll` is copied next to the shell (Office interop dependency): include it in the MSI file list or drop it deliberately.

Not adopted (same as GLSense today): Authenticode check of the core DLL at load — SHA256 vs manifest only.
Worth adding to both later.

## 7. Regression checklist
- Login (manual, SSO, from GLSense), logout (all panes), two Excel windows, GLSense + XLEdge loaded together.
- Run report from pane; title commands Edge/Process/EdgeWorkbook/Logs/Excel; XLSX download.
- Refresh Sheet / Book, Param Refresh (sheet/book), Control Sheet, Options, About (core + shell version).
- Drilldown hyperlink, "Goto Report Data", Fusion/EBS attachment download, empty attachment cells (OISR-22766/22767).
- Calendar / segment pickers on cell selection; sheet-delete cleanup of named ranges and caches.
- Number/date formats on non-English regions (OISR-22663/22696).
- Reload offline / online / Release History; reload blocked while a report runs; rollback to an older core.
- Excel close with a window open → no orphan Excel.exe; next update extracts cleanly; uninstall removes `AddinCore\`.

## 8. Risks (from GLSense's recorded lessons)
| Risk | Mitigation |
|---|---|
| `Task`/async across Remoting, unobserved exceptions kill Excel | Shell→core calls return immediately; core entry points catch everything (PORTING_GUIDE §1-2) |
| `InvalidCastException` on COM objects after `await`s (`CrossAppDomainSink`) | Cache Workbook/Worksheet at operation start (GLSense §24.2/§25) |
| "Object has been disconnected" after idle | `InitializeLifetimeService() => null` on every MarshalByRef |
| Unload with work/windows alive → orphan Excel, locked DLLs | `IsBusy`, open-window warning, explicit Shutdown, 5 s cap (GLSense §29/§71) |
| Older core loaded from Release History lacks new contract members | try/catch `MissingMethodException`/`RemotingException` (GLSense §40.3) |
| WebView2 in a re-parented core window | Phase 1 spike, documented fallback |
| Dual maintenance until cutover | Every `11.1.2` fix re-applied to the twin; MIGRATION_STATUS.md tracks it |

## 9. Noticed along the way (separate from this work)
- GLSense `OrbitGLSense.vdproj` (~line 2556): `GLSense.dll` SourcePath is `..\obj\Debug\GLSense.dll` — looks wrong for a Release MSI.
- GLSense comments `ConfiguratorPaneHost.cs:16`, `IGLSenseAddin.cs:66` say "own dedicated WPF thread"; the code uses Excel's thread.
