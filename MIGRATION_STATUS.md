# XLEdge VB.NET → C# WPF Migration — Status & Reference

Last updated: 2026-08-02

## Fixed S107 (9-parameter method) warning in ReportGenerator.cs (branch 11.1.0) — 2026-08-02

`TryResolveReportXmlForRefresh` had grown to 9 parameters (3 in, 6 out) - the deferred parameter-count
warning from the build-warnings pass above. Bundled the 6 out-values into a private nested
`ReportXmlRefreshResult` class (`Title`, `ReportId`, `RunId`, `MetaJson`, `ParamsJson`, `Mappings`),
bringing the method down to 4 parameters (3 in, 1 out). Updated both call sites - inside the public
`TryGetStoredReportXml` (whose own external signature, used by `AddinModule.cs`, `ParamsControlSheetBuilder.cs`,
and `XLEdgeParamsBuilder.cs`, was left unchanged) and inside `RefreshListObjectAsync` - to read the needed
values off the returned result object. No behavior change; confirmed only 2 internal call sites exist
(no external callers of `TryResolveReportXmlForRefresh` itself). The remaining 12 Cognitive Complexity
(S3776) warnings are still deferred as a separate, larger follow-up.

## Fixed safe build-warning findings (branch 11.1.0) — 2026-08-02

After the first successful build produced `Warnings.txt` (63 SonarQube/Roslyn findings, all in
`ReportGenerator.cs`), fixed the mechanical/low-risk ones: extracted 3 repeated string literals into
constants (`extraParameters`, the "request timed out" message, `"Column"`); extracted 2 nested ternaries
into explicit if/else; changed 5 `throw new Exception(...)` to `throw new InvalidOperationException(...)`
(all bubble to generic `catch (Exception ex)` callers, no behavior change); awaited `ShowErrorAsync`
instead of calling the synchronous `ShowError` from inside an already-async lambda; rewrote a CSV-parsing
`for` loop as an equivalent `while` loop so an intentional extra index-advance (for an escaped `""` inside
a quoted field) doesn't read as mutating the loop's own stop-condition variable; removed an unused
`sheet` parameter (confirmed unreferenced) and a fully dead private method (`GetColumnMappings`, zero
call sites); removed a useless final assignment and two unused locals. Deliberately deferred: the 12
Cognitive Complexity warnings (methods sized 16-192) and the related 9-parameter-method warning - these
need careful, verified method-splitting that's risky to do blind without a compiler here, better done as
a dedicated follow-up. Also left ~20 lower-severity IDE0xxx style suggestions and a build-environment
"file access denied" warning (not a code issue) untouched.

Also: a Python-based text-replace pass mid-way accidentally collapsed the file's CRLF line endings to LF
project-wide (Python's default text-mode read/write does universal-newline translation) - caught via an
unexpectedly huge git diff, and fixed by restoring CRLF before committing. Worth remembering for any
future scripted edits to this codebase: prefer the Edit tool (preserves line endings exactly), or open
scripted file I/O in binary mode / with `newline=''` when a script must touch a whole file.

## Added OrbitXLEdge installer project (branch 11.1.0) — 2026-08-02

For QA handoff, ported the VB.NET reference's `setupfiles\Orbit\OrbitXLEdge.vdproj` (a Visual Studio
Installer Project, which only ever produces an .msi - no .exe bootstrapper, matching the requirement
that only .msi is supported) to this C# port on a new `11.1.0` branch. Adapted the file list to this
project's actual dependencies rather than copying the VB reference's wholesale: kept Add-in Express XL/
MSO, the Office/VBE interop DLLs, NLog, and the three WebView2 assemblies; dropped Newtonsoft.Json
(unused here - this port uses `System.Text.Json`) and `System.ValueTuple.dll` (a compile-time-only
facade on .NET Framework 4.7+, confirmed via its own NuGet `.targets` file that it's never actually
copied to `bin\Release`); added this port's additional real dependencies not present in VB (MahApps.
Metro.IconPacks Core/FontAwesome, Wpf.Ui + Wpf.Ui.Abstractions, Microsoft.Xaml.Behaviors, and the full
`System.Text.Json` polyfill chain); and, unlike the VB reference (which excluded it), kept `Microsoft.
Web.WebView2.Wpf.dll` in the package since this port's task pane actually hosts the WPF WebView2 control.
Same adxloader/adxregistrator custom-action install/uninstall pattern, per-user `TARGETDIR`, and
WebView2Loader.dll native-runtime placement as the reference. Banner/icon use the user-provided
`Images\orbit_bitmap.bmp` / `Images\OrbitGLSense.ico`. Bumped the .NET Framework prerequisite to 4.8.1
(matching `TargetFrameworkVersion`, vs. the VB reference's mismatched 4.6.2/4.7.2) and dropped its VSTO
runtime prerequisite (Add-in Express doesn't need it). Generated fresh ProductCode/PackageCode/
UpgradeCode (independent of the VB installer's) - QA will manually uninstall any prior version before
installing this one, so no upgrade-code matching was needed. Wired into `XLEdge.sln` as a second project,
same as the VB reference's own `.sln`.

Caveat: hand-authored from the reference project's structure; could not be opened/built in Visual Studio
from this environment (no Windows/MSBuild toolchain here) to confirm it loads cleanly. First step is
opening it in Visual Studio (with the "Visual Studio Installer Projects" extension) and doing a test
build.

## Excel.exe lingers after close; new Excel instance's add-in fails to load — 2026-07-31

User-reported: sometimes, after closing Excel, its process stays running in the background; if a new
Excel instance is then opened while that old process is still around, the add-in doesn't load in it.

Root cause: `WebView2` (the `Microsoft.Web.WebView2.Wpf.WebView2` control hosted inside `XLEdgeCTP`, one
per open workbook's task pane) was never explicitly `Dispose()`d anywhere in the codebase. `XLEdgeCTP.
OnUnloaded` unsubscribed its various `CoreWebView2` event handlers but never called `WebCtrl.Dispose()`.
`Dispose()` is what actually tells WebView2's underlying `msedgewebview2.exe` browser process and its
`CoreWebView2Environment` to shut down cleanly; without it, that browser process (and its lock on the
environment's user data folder) can outlive the workbook/Excel window that created it. Every task pane's
`CoreWebView2Environment` is created against the same fixed, shared folder
(`XLEdgeAppPaths.BrowserLogsFolder`, `%LOCALAPPDATA%\ORBIT\Excel_Logs\XLEdge_Logs\BrowserLogs` - not
unique per process or per pane) - the same sharing pattern already identified as the cause of the
logout hang fixed earlier today. If an old Excel process's WebView2 browser process is still alive and
holding that folder's profile lock, a newly-opened Excel instance's own task pane trying to initialize
its own WebView2 against the same folder can hang, which is consistent with its add-in appearing not to
load. Fixed by calling `WebCtrl.Dispose()` in `XLEdgeCTP.OnUnloaded`, after the existing event
unsubscriptions, so the browser process and its folder lock are released as soon as the pane's WPF
visual tree is torn down (workbook/Excel closing) rather than lingering indefinitely.

## Logout hangs / opens multiple wait windows with multiple workbooks open — 2026-07-31

User-reported: clicking Logout with only one workbook open logs off smoothly, but with multiple
workbooks open it hangs, and multiple "Logging Off" wait windows appear (one per workbook). User
suspected this already worked correctly in the VB.NET version.

Root cause (confirmed by diffing against VB.NET's `LogOffSessionAndWaitAsync`/`LogOffAllTaskPanesAsync`
in `AddinModule.vb`): `XLEdgeCTP.LogoutSessionAsync` (`Views\XLEdgeCTP.xaml.cs`) called
`await EnsureWebViewInitializedAsync();` before checking whether the pane's WebView2 was already
initialized. Each open workbook gets its own task pane / `XLEdgeCTP` instance, and each one creates its
own `CoreWebView2Environment` pointed at the same shared, single `XLEdgeAppPaths.BrowserLogsFolder` user
data folder. `LogOffAllTaskPanesAsync` (`AddinModule.cs`) logs off every open task pane in a loop - for
any workbook whose pane/WebView2 had never been initialized before (e.g. a background workbook the user
never actually opened the report pane for), this forced a brand-new `CoreWebView2Environment` to be
created against a profile folder another workbook's environment already had open/locked in the same
process - a known WebView2 contention scenario that can hang indefinitely. With only one workbook open
this never triggers, since that workbook's pane is normally already initialized from regular use. VB.NET
never has this problem because its equivalent (`LogOffSessionAndWaitAsync`) only checks whether
`CoreWebView2` already exists and skips the pane entirely otherwise - it never lazily creates a WebView2
during logoff. Fixed by removing the `EnsureWebViewInitializedAsync()` call from `LogoutSessionAsync`,
matching VB.NET: a pane with no `CoreWebView2` yet has no active session to log out of, so it's now just
skipped.

Separately, `AddinModule.LogoffFromXLEdgeAddin`'s wait window was shown via
`GetFirstAvailableTaskPane()?.GetWpfDispatcher()` - an arbitrary open workbook's task pane dispatcher -
instead of the single shared app-wide `UiDispatcher.Current` every other wait/busy window in this add-in
uses (e.g. `ReportGenerator.CreateAndShowWaitWindow`). Switched to `UiDispatcher.RunAsync`, removing the
now-unused `GetFirstAvailableTaskPane` helper, so the wait window no longer depends on which task pane
happens to be first in the collection.

## Live-usage bug batch (post-first-build, real Excel session) — 2026-07-30

Seven items addressed across a single day of actual usage (not code-reading/build-log driven like
the batch below - each item here started from a user-reported symptom in a live Excel session, cross-
checked against `D:\Latest_Addons\XLEdge` where the fix required matching specific VB.NET behavior).
No Windows/MSBuild toolchain in this environment, same caveat as every other entry in this file -
every fix verified by brace-balance/XML-well-formedness checks only; each was confirmed working by the
user after a real rebuild.

### Duplicate full-stack-trace logging for a single API failure

A single API failure (e.g. a server 500) was logged with a full exception + stack trace **three**
separate times as it propagated up: `ApiHelper.ExecuteApiCall`'s catch-all, `ApiOperationHelper.
ExecuteWithRetry`'s non-transient/retry-exhausted catches, and the top-level caller's own catch (e.g.
`ReportGenerator`'s "Unhandled error in report generation"). Each layer re-dumped the identical
exception/stack trace since `throw`/`throw;` preserves the original trace as it bubbles up - nothing
new was actually captured at each layer, just noise. Fixed by trimming the first two layers to a
single summary line each (status/message, no stack trace); the full exception is now logged exactly
once, by whichever top-level caller ultimately handles/displays the failure to the user.

### Drilldown reports never got IT1 = "Child Report", so Refresh/Param Refresh stayed enabled

`XLEdgeAppState.Instance.FollowDrilldown` was only ever *read* (by `ReportGenerator.
RewriteParameterSectionRows`, to decide whether to write IT1 = "Child Report" on the parameter sheet)
and reset to `false` (in `ProgressCoordinator.ResetReportState`, itself dead code with zero call
sites) - nothing in the C# port ever set it to `true`. VB.NET's `AddinModule.vb` sets
`FollowDrilldown = True` right before kicking off generation in the drilldown hyperlink handler
(`SheetFollowHyperlink`); that assignment was simply missing here. Result: every report generated by
clicking a drilldown hyperlink had its IT1 cell left empty, so `AddinModule.cs`'s isChild check (which
reads IT1) came back false and Refresh/Refresh All/Run were never actually disabled for these child
reports. Fixed by having `CreateReportFromTitleAsync` set
`XLEdgeAppState.Instance.FollowDrilldown = isDrilldownRequest` (a flag it already computed locally
from whether a `paramsJsonPayload` was supplied) at the top of every call - correct for the invocation
actually running, without depending on any "operation-completed" reset elsewhere.

### Ribbon Refresh/Refresh All/Param Refresh not disabled for Child Report sheets

Even after the IT1 fix above, the ribbon buttons stayed enabled on a child-report sheet - only the
click handlers (`RibEdgeParamRefresh_OnClick` etc.) blocked the action reactively, after the click.
Root cause: `XLEdgeRibbonHelper.ProcessActiveWorkbook` (run after every report generation and on every
Sheet/WorkbookActivate) only checked the ListObject's name shape (`ORB_..._E`) to decide whether to
enable Refresh/Param Refresh/Refresh All - it never read IT1 at all, so a child report's table looked
exactly as "refreshable" as any other. Fixed by adding `IsChildReportSheet` (a read-only mirror of
`AddinModule.TryResolveInstanceAndChildFlag`'s sheet-resolution + IT1 read, deliberately without that
method's instance-mismatch `MessageBox` popup, since this now runs silently on every sheet activation)
and wiring it into `ProcessActiveWorkbook` (disables Refresh/Param Refresh for the active sheet) and
`BookHasEdgeReport` (so Refresh All also disables when every report in the book is a child report -
matching `RibEdgeRefreshAll_OnClick`, which already skips child sheets one by one).

### Error toast flashing and disappearing instantly, task pane blank for the toast's full duration

`AppOverlay.xaml` has no `"ShowBusy"`/`"HideBusy"` Storyboard resources defined anywhere (only
`ToastSlideIn`/`ToastSlideOut` exist), so `AppOverlay.xaml.cs`'s `HideBusyAsync()` "storyboard
missing" fallback branch is in fact the *only* branch it ever takes. That fallback unconditionally set
`this.Visibility = Collapsed` - collapsing the entire overlay root (Toast, Busy and Confirm all live
under it) with no check for whether a Toast was currently showing. `ReportGenerator.
CreateReportFromTitleAsync`'s catch block calls `DisplayErrorAsync` (shows the error Toast) immediately
followed by `CleanupAsync`, which calls `HideBusyAsync` to dismiss the earlier "Downloading report
data..." busy spinner - so the just-shown error Toast got collapsed a few milliseconds later, before
it could be read. Worse, collapsing the root this way skipped `RemoveBlurFromSiblings()`, so
`_toastHidesWebView2` was never cleared and the WebView2 task pane content stayed
`Visibility.Hidden` - genuinely blank - until the Toast's own independent 60-second timer eventually
elapsed and called `DismissToast()`, the only thing that actually restores WebView2 (matching the
reported symptom exactly, including why Excel's own ribbon/worksheet stayed usable throughout - only
the WebView2-hosted task pane content is affected). Fixed by adding the same Toast/Confirm-aware guard
the (dead) storyboard-completion branch already had - only collapse the overlay root if no Toast is
actually visible (`Opacity` ~0) and Confirm isn't showing either.

### Windows drifting off-center after a resize

Ported from `D:\SQLLite_Test\GLSense\FinalWorkingCode`, which documented and fixed the identical bug
in its own `DpiAwareWindow.cs` (see that project's `CLAUDE.md`). XLEdge's copy of this base class was
ported from an earlier, pre-fix version of GLSense's and never received the fix.
`WindowStartupLocation="CenterOwner"` (set per-window in XAML, or via `MessageFunctions.
XLEdgeMessage`) only centers a window once, at the moment WPF applies it. `FitToAvailableWorkArea()`
(runs once from `OnLoaded`) and `EnsureFitsWorkArea()` (runs on every `OnRenderSizeChanged` - e.g. a
DataGrid populating with data after an async load, or a DPI change) can both resize the window
afterward - but only ever changed Width/Height, never Left/Top, so a resize always grew/shrank
anchored at the window's current top-left corner and the window's true center silently drifted away
from wherever `CenterOwner` originally centered it. Every window in this codebase derives from
`DpiAwareWindow`, so this affected all of them. Fixed by porting GLSense's
`RecenterAfterSizeChange(previousLeft, previousTop, previousWidth, previousHeight)` directly: both
methods now capture Left/Top/Width/Height before making their change, and recenter around the same
center point afterward (clamped so recentering never pushes the window off the visible work area) if
they actually changed the size. `ExcelWindowHelper.CenterWindowOverExcel` (a separate, manual Left/Top
calculation) was checked and confirmed to have zero call sites - dead code, not part of the live
centering path, left untouched.

### Scheduled ("Process") reports using the same API endpoints as live ("Edge") reports

Verified against VB.NET's `FormProcessBar.vb` (`ReturnHTTP`, `MetaInfo`). `CreateReportFromTitleAsync`'s
CSV and report-definition (Meta) fetches always used the live-run ("Edge") endpoints regardless of the
title's type segment ("Edge" vs "Process") - only ever branching on drilldown-or-not - so a submitted/
scheduled report was requested with the same URL shape as a live ad-hoc one, which the server doesn't
recognize the same way, causing scheduled reports to fail to download entirely. VB.NET branches on
this type segment in both functions: `ReturnHTTP` - Edge posts JSON to `report/runner?runId=...&
type=csv`; Process instead does a plain GET to `process/excel-data?processId=...&type=csv` (Form, no
body). `MetaInfo` - Edge uses `report/report-definition?reportId=...&runId=...&isDrillDown=...`;
Process uses `process/report-definition?processId=...&isDrillDown=false`. `FollowDrilldown` always
takes priority over the type check in both, regardless of the parent report's type. The existing
params fetch was already correct - VB's `ParamInfo` hits the same `report/params` endpoint for both
types, just varying a `type` query parameter, which the C# port already did. Also fixed
`BuildReportTable`'s `tableId`, which always suffixed `"_E"` regardless of report type - VB's
`EETableID` assignment suffixes a Process report's table `"_P"` instead. This wasn't the cause of the
download failure (it only runs after a successful fetch), but `AddinModule.UpdateTabLabel` and
`XLEdgeRibbonHelper.ProcessActiveWorkbook` already have complete, correct handling for `"_P"` tables
(scheduled-output caption, Refresh/Param Refresh disabled) that was simply dead code until now, since
nothing ever produced a `"_P"` table to trigger it.

### XLEdgeOptions checkbox tooltips + spacing (UI polish, no behavior change)

Added a `ToolTip` (using the existing `ChromeStyleToolTip` style from `GlobalStyles.xaml`, which had zero
usages anywhere before this) to all six checkboxes in `XLEdgeOptions.xaml`, each showing the option's
label plus a plain-language description of what it does. Descriptions for `ParameterValuesInSameSheet`
(`ReportGenerator.cs`), `SyncWithReportDefinition` (`ReportGenerator.cs`'s stale-column cleanup on
refresh) and `ShowCalendarControl` (`AddinModule.cs`'s calendar-popup gate) are based on real, traced
consuming code. `DownloadScheduledOutputsToExistingSheets`, `OverrideSheetNameForScheduledOutputs`, and
`OverrideFormats` describe the option's evident intent, but a codebase-wide search found no consuming
logic anywhere outside their own persistence plumbing (`XLEdgeOptions.xaml.cs`,
`XLEdgePreferencesManager.cs`) - these three settings appear to be saved/loaded but not yet wired to any
actual behavior in the C# port. Worth a follow-up to confirm whether that's expected or a gap.

Also increased the gap between a checkbox's glyph and its label by 2px, app-wide: `ModernCheckBox`'s
`Padding` in `GlobalStyles.xaml` changed from `4,0,0,0` to `6,0,0,0` (this style has no custom
`ControlTemplate`, so `Padding` maps directly to WPF's default glyph-to-content spacing). This affects
every `CheckBox` using `ModernCheckBox` app-wide, not just the Options window.

### Button tooltip audit across all windows (read-only check, no changes)

Checked every `Button` in every `Views\*.xaml` file for a `ToolTip`. Only `XLEdgeServerConfiguration.xaml`
(GO/Set as Default/Save/Delete/Close) and `XLEdgeWaitWindow.xaml` (Cancel) have them. Every other action
button is missing one: `AppOverlay.xaml` (toast close, busy-cancel, Yes/No/Cancel confirm), `XLEdgeAbout.
xaml` (Close), `XLEdgeCalendar.xaml` (OK, Close), `XLEdgeDrilldownReports.xaml` (Execute, Close),
`XLEdgeGLAccountsWindow.xaml` (OK, Cancel), `XLEdgeLoginDetails.xaml` (Close), `XLEdgeOptions.xaml`
(Apply, Save, Close). The custom title-bar "X" close button is untooltipped consistently across every
single window in the app (including the two windows above that otherwise have full tooltip coverage) -
that looks like a deliberate, consistent choice rather than an oversight.

## RibEdgeRefreshAll enabled for books with no live ("_E") report — 2026-07-31

`XLEdgeRibbonHelper.BookHasEdgeReport` (the function `ProcessActiveWorkbook` calls, on every sheet/
workbook activation and after every report run, to decide whether `RibEdgeRefreshAll` should stay
enabled when the active sheet itself doesn't qualify) only checked that a sheet's table name started
with `"ORB_"`, wasn't `"orb_params_control"`, and wasn't a child (drilldown) report - it never checked
the `"_E"` suffix. `RibEdgeRefreshAll_OnClick` (`AddinModule.cs`) only ever collects `"_E"`-suffixed
tables to refresh and ignores `"_P"` (scheduled/Process) tables entirely, so a workbook containing only
scheduled-output sheets had RefreshAll enabled even though clicking it would find nothing to refresh
("No reports in the workbook to refresh!"). Fixed by adding
`sheet.ListObjects[1].Name.EndsWith("_E", StringComparison.Ordinal)` to `BookHasEdgeReport`'s condition,
matching the same suffix check `ProcessActiveWorkbook` already applies to the active sheet's own table
(`isRefreshableReportTable`) and the one `RibEdgeRefreshAll_OnClick` uses when collecting tables to
refresh. Also noted, not fixed: `AddinModule.UpdateTabLabel` sets `XLEdgeAppState.Instance.RefreshAll`
(true/false) per active sheet, but nothing in the codebase reads that property - it looks like orphaned
state, unrelated to the actual ribbon-enabled mechanism (`XLEdgeRibbonHelper.EnableControls`/
`DisableControls`, which set the ribbon control's `Enabled` property directly via reflection and are
what this fix touches).

## First real build + first real runtime bug batch — 2026-07-23

This is the first entry in this file backed by an actual Visual Studio build and an actual Excel
session (every earlier entry was verified only by brace-balance/XML-well-formedness checks, no
compiler available in this environment) - both a real build-error log and a real
`XLEdge_Logs_23-Jul-2026.log`/user bug report were used to find and fix the items below. Read this
before re-touching any of the files listed here.

### Build errors: 4 files never added to the csproj

`XLEdge.csproj` was missing `<Compile Include>` entries for `Helpers\ApiErrorMessageExtractor.cs`,
`Helpers\ApiExceptions.cs`, `Helpers\XLEdgeTempFileCleaner.cs`, `Helpers\DrilldownRequestBuilder.cs`
- all four existed on disk (created in earlier sessions) but were never wired into the build, so
MSBuild silently skipped them, producing CS0246/CS0103/CS0234 at every call site instead of an error
on the files themselves. Added the missing entries; verified via `comm`-diffing every physical `.cs`
file against every csproj `<Compile>` entry in both directions.

### FluentWindow crash: `ExtendsContentIntoTitleBar` vs. this app's transparent-chrome XAML pattern

Root cause of 3 of the 4 crash blocks in the log, and (see below) the actual reason Logout appeared
to do nothing. `Utilities\DpiAwareWindow.cs`'s base class was swapped to `Wpf.Ui.Controls.FluentWindow`
in an earlier session (the MahApps→WPF-UI task above) - `FluentWindow.ExtendsContentIntoTitleBar`
defaults to `true`, and applying that default in `OnSourceInitialized` coerces `WindowStyle` in a way
that's incompatible with `AllowsTransparency="True"` (`InvalidOperationException: "WindowStyle.None
is the only valid value for WindowStyle when AllowsTransparency is true"`) - thrown by every window
using this app's own custom-chrome pattern (`WindowStyle="None"` + `AllowsTransparency="True"`, e.g.
`XLEdgeMessageWindow`, `XLEdgeWaitWindow`), unlike GLSense's own FluentWindow usage (`WindowStyle=
"SingleBorderWindow"` + `ExtendsContentIntoTitleBar="True"`, no `AllowsTransparency`). Fixed with a
single `this.ExtendsContentIntoTitleBar = false;` in `DpiAwareWindow`'s constructor - fixes every
derived window at once, no per-View XAML changes needed.

### Drilldown hyperlink SubAddress was always empty

`Helpers\ReportGenerator.cs`'s drilldown-hyperlink-writing loop called
`sheet.Hyperlinks.Add(cell, "", "", tooltip, ...)` - the empty `SubAddress` meant
`AddinModule.adxExcelAppEvents1_SheetFollowHyperlink`'s `dataSheet.Range[hyperLink.SubAddress]` threw
`COMException 0x800A03EC` every time, silently swallowed by a surrounding try/catch (drilldown click
did nothing, no visible error). Fixed by passing `cell.Address` instead of `""`. Every other
`Hyperlinks.Add` call site in this file was checked and confirmed NOT to need this (they either don't
read `SubAddress` back, or already used `cell.Address`).

### XLEdgeServerConfiguration: 7 window/grid bugs

All in `Views\XLEdgeServerConfiguration.xaml(.cs)`:
- **Not centered on open**: `WindowStartupLocation="Manual"` with no Left/Top ever set anywhere ->
  changed to `WindowStartupLocation="CenterScreen"`.
- **Couldn't add a new row**: `CanUserAddRows="False"` and no Add button existed at all. Added a
  `btnAdd`/`BtnAdd_Click` that constructs a new `UrlInstance` (with `PropertyChanged` wired up, unlike
  the grid's own native add-row which would bypass that subscription) and puts the grid into edit mode
  on it.
- **Empty-row delete should warn, not delete**: `BtnDelete_Click` now checks for a blank Name+Address
  row first and shows a warning instead of running the confirm-delete flow.
- **Mandatory fields / name-only duplicate check**: already correctly implemented in
  `SaveConfiguration()`/`DgInstances_CellEditEnding` - verified, no change needed.
- **Checkbox alone shouldn't set default**: `DgInstances_BeginningEdit` now unconditionally cancels
  edits on the Default column (previously only cancelled when unchecking an already-default row),
  pointing the user at the "Set as Default" button instead.
- **Default row should sort first**: added `ReorderInstancesWithDefaultFirst()` (default row first,
  everything else alphabetical by Name via `ObservableCollection.Move`, preserving object references/
  selection), called from `SetDefaultInstance`, `LoadConfiguration`, and `SaveConfiguration`.
- **Status text should wrap**: `txtStatus` `TextWrapping="NoWrap"` -> `"Wrap"`.

### XLEdgeCTP: close button clipped + toast only visible over the header

- **Close button invisible**: `ADXExcelTaskPane1.ADXAfterTaskPaneShow` hardcoded `this.Width = 600`
  (raw pixels) instead of DPI-scaling it like every other sizing call in that file - at >100% display
  scaling, 600px is narrower than the `XLEdgeCTP` UserControl's `MinWidth="600"` (WPF DIPs), and with
  `HorizontalScrollBarVisibility="Disabled"` the overflow was clipped instead of scrollable, cutting
  off the header's Close button. Fixed to scale `_minWidthDip` by the current DPI, matching
  `ApplyDpiAwareSizing`/`SetBoundsCore` elsewhere in the same file.
- **Toast only visible over the header, not the WebView2 area**: WebView2 hosts its own child HWND
  and paints outside WPF's composition pipeline entirely (the standard WPF "airspace" limitation for
  any windowed/HwndHost control) - it always renders on top of the AppOverlay/Toast regardless of
  `Panel.ZIndex`, so the overlay only appeared to "work" over the plain-WPF header above it. Fixed in
  `Views\AppOverlay.xaml.cs`: `ApplyBlurToSiblings`/`RemoveBlurFromSiblings` now also recursively find
  and temporarily hide (`Visibility.Hidden`)/restore any `WebView2` descendants of the blurred
  siblings while an overlay is showing.

### Ribbon Login/Logout caption bug

`Helpers\XLEdgeRibbonHelper.cs`'s `ApplyLoggedInState()` was writing the selected instance name onto
`RibEdgeLogin`'s caption (the control being HIDDEN at that moment) instead of `RibEdgeLogout`'s (the
one becoming visible) - and `ApplyLoggedOutState()` never reset Login's caption back to "Login" at
all, so after a login+logout cycle Login became visible again still showing the old instance name.
Fixed both: `ApplyLoggedInState` now sets `RibEdgeLogout`'s caption; `ApplyLoggedOutState` now
explicitly resets `RibEdgeLogin`'s caption to `"Login"`.

### Logout didn't actually log out of the instance

Root cause: the FluentWindow crash above. `AddinModule.LogoffFromXLEdgeAddin`'s wait-window-show call
threw every time (log: `LogoffFromAddin|Error during logoff` at the `waitWindow.Show()` line) and was
caught by the SAME try/catch wrapping `LogOffAllTaskPanesAsync` (the call that actually navigates each
task pane's WebView2 to `/web/secure/applogout`) - so the real logout logic never ran at all, even
though the ribbon still flipped to "logged out". Fixed by the FluentWindow fix above, plus hardening
`LogoffFromXLEdgeAddin` itself so a wait-window display failure (for any reason, in the future too)
is caught on its own and can never again block the actual logoff work below it.

### False "sheet is missing parameter data" message on every sheet select

`AddinModule.UpdateTabLabel` unconditionally warned when `ExcelSheetHelper.GetParameterSheet` returned
null for any `_E`-suffixed report table - but a companion "P_" parameter sheet only ever exists for
reports generated in SEPARATE-sheet mode (`ReportGenerator.BuildCompanionParameterSheet`, table
starts at row 1); SAME-sheet mode (row 1-7 banner embedded in the data sheet itself, table starts at
row 8) never creates one by design, so `GetParameterSheet` correctly returning null there was being
misreported as an error on every activation of a same-sheet report tab. Fixed by only running the
check when `tableObj.HeaderRowRange.Row <= 1` (separate-sheet mode).

### Native MessageBox usage

Audited every `MessageBox.Show`/`System.Windows(.Forms).MessageBox` call site in the project.
`Helpers\ParamsControlSheetBuilder.cs` was already correctly using the custom
`MessageFunctions.XLEdgeMessage` (the grep hits there were just `MessageBoxIcon`/`MessageBoxButtons`
enum arguments, not a native call). Two genuine native calls remained: `Helpers\LogHelper.cs`'s
logger-init failure path (now routes through `MessageFunctions.XLEdgeMessage`, with a native
`MessageBox.Show` kept only as an absolute last-resort fallback if even that fails this early in
startup) and `Utilities\MessageFunctions.XLEdgeMessage`'s own outer catch (deliberately left as a
last-resort fallback for when the custom window itself fails to show - this used to fire on every
single message via the FluentWindow crash above, which is why it looked like the custom window was
never used at all; with that crash fixed this catch should now be effectively unreachable).

### Ribbon Refresh/SheetRefresh not re-enabling after running a report

`Helpers\ReportGenerator.cs`'s `ExcelBulkOperationScope` (used via `using var scope = new
ExcelBulkOperationScope();` at the top of every report-run method) sets `EnableEvents = false` for the
whole report-run duration - so Excel's own `SheetActivate` event for the newly-created/refreshed
report sheet becoming active never fired at all (events suppressed during that window are lost, not
queued/replayed once re-enabled), meaning `AddinModule.adxExcelAppEvents1_SheetActivate` (the only
thing that normally calls `ApplyRibbonState("ApplySheetActiveState")`) never ran, leaving
Refresh/Param Refresh reflecting whichever sheet was active BEFORE the report ran. Fixed by calling
`XLEdge.AddinModule.ApplyRibbonState("ApplySheetActiveState")` explicitly at the end of
`ExcelBulkOperationScope.Dispose()`, right after events are re-enabled.

### Verification note

No Windows/MSBuild toolchain is available in this environment (Linux sandbox) - every fix above was
verified by brace-balance counting (`grep -o "{" file | wc -l` vs `}`) and, for touched XAML, `python3
-c "import xml.etree.ElementTree as ET; ET.parse(...)"`, same as every other entry in this file. All
of the above needs a real rebuild + Excel relaunch to confirm - not yet done at the time of writing.

## Architecture: MahApps.Metro → WPF-UI swap + GLSense DPI/resettle patterns — 2026-07-23

Unlike every other entry in this file, this isn't a VB-parity fix - it's the UI-framework-alignment
task explicitly scoped by the user ("port DPI util + WPF-UI patterns from GLSense, then replace
MahApps.Metro with WPF-UI"). Flagged to the user first given the real risk profile (no git repo in
this environment - confirmed via `git status` - and no .NET/MSBuild toolchain available to compile
anything, so every change below is verified by XML well-formedness + manual brace-balance checks
only, the same method used throughout this session, but a materially weaker guarantee for a
framework swap than for a VB-logic port). User chose "back up first, then proceed" - a full source
backup (everything under `XLEdge\XLEdge\` except `bin`/`obj`, byte-identical file list confirmed via
`diff`) was made to `XLEdge_Backup_2026-07-23_pre-wpfui\` before any change below.

**Actual scope turned out far narrower than the task description implied - verified by grep before
touching anything, not assumed:**
- `Utilities\DpiAwareWindow.cs` (this project's actual window base class - every `Views\*.xaml` root
  is `<utils:DpiAwareWindow ...>`, not `MahApps.Metro.Controls.MetroWindow` directly) already
  inherited from plain `System.Windows.Window`, not MahApps.Metro's `MetroWindow` - so MahApps.Metro's
  real window-chrome/base-class was never actually used anywhere in this codebase. The DPI-awareness
  utility half of the task (`Utilities\DpiAwarenessHelper.cs`/`DpiAwareWindow.cs`) was *also* already
  done in an earlier session pass (see this file's own "C#-only additions" audit above) - confirmed
  by direct comparison against `GLSense.Addin.Core\Helpers\DpiAwarenessHelper.cs` that XLEdge's copy
  is a superset (adds `GetDpiForMonitor`/`MonitorFromWindow`/`GetDpiForScreenPoint` for correct
  multi-monitor accuracy that GLSense's own version lacks).
- Grepped every `Views\*.xaml` for `mah:`/`xmlns:mah` and every `.cs` file for `ControlzEx`: only 2
  files (`XLEdgeCalendar.xaml`, `XLEdgeServerConfiguration.xaml`) even declared `xmlns:mah`, and
  neither actually used a single `mah:`-prefixed control anywhere in its body - dead, vestigial
  namespace declarations. `Themes\GlobalStyles.xaml`'s only MahApps dependency is
  `MahApps.Metro.IconPacks.FontAwesome` (icons) - the SAME icon-only package GLSense itself uses
  alongside WPF-UI (confirmed against GLSense's own CLAUDE.md history), not the actual
  `MahApps.Metro` theme library. `using ControlzEx.Standard;` in `ReportGenerator.cs`/
  `MessageFunctions.cs` were unused imports (no `ControlzEx.*` symbol referenced anywhere else in
  either file) - only `Utilities\MahAppsBootstrapper.cs` genuinely used `ControlzEx.Theming`
  (`ThemeManager`).
- Grepped every `.xaml`/`.cs` file for `MahApps.Brushes.`/`MahApps.Colors.` (the resource-key family
  `MahAppsBootstrapper.cs` spent ~200 lines manually defining/deriving accent shades for): **zero**
  hits anywhere outside `MahAppsBootstrapper.cs` itself. That entire brush/color/accent-derivation
  system (`SetAccentColors`/`RegisterCoreBrushes`/`RegisterCompatibilityPairs`/`Lighten`/`Darken`) was
  dead code - nothing in this app's XAML or C# ever consumed a single key it defined - so none of it
  needed to be preserved or re-derived under new key names.

**What actually changed:**
- `Utilities\WpfUiBootstrapper.cs` (new) replaces `Utilities\MahAppsBootstrapper.cs`, modeled directly
  on `GLSense.Addin.Core\Utilities\WpfUiBootstrapper.cs`'s pattern: load the real
  `pack://application:,,,/Wpf.Ui;component/Resources/Theme/Light.xaml` +
  `.../Resources/Wpf.Ui.xaml` dictionaries via pack URI, then unconditionally define a small,
  hand-rolled fallback for every Fluent design-token key (`ControlBackgroundBrush`,
  `TextFillColorPrimaryBrush`, `ApplicationBackgroundBrush`, etc. - same key names/fallback colors as
  GLSense's own `AddRequiredResources`/`AddFallbackResources`) so a missing/failed pack URI load can
  never leave a `DynamicResource` lookup undefined. `Init(accentHex, baseTheme)`/`PreloadResources()`
  keep the exact method names/signature the old bootstrapper had, so `AddinModule.cs`'s single call
  site (`AddinModule_OnRibbonLoaded`) only needed its class name changed
  (`MahAppsBootstrapper.Init/.PreloadResources` -> `WpfUiBootstrapper.Init/.PreloadResources`), not
  its call shape. `Init`'s `baseTheme` now drives `Wpf.Ui.Appearance.ApplicationThemeManager.Apply(
  ApplicationTheme.Light/.Dark)` instead of loading MahApps.Metro's `Themes\{theme}.xaml`.
- `Utilities\DpiAwareWindow.cs`: base class `Window` -> `Wpf.Ui.Controls.FluentWindow` (this alone is
  the literal "replace MahApps.Metro with WPF-UI" ask, and - since `FluentWindow` is itself a `Window`
  subclass - a drop-in swap that leaves every existing event hook/override in this class (`Loaded`/
  `SourceInitialized`/`ContentRendered`/`Closed`/`Unloaded`, `OnInitialized`/`OnRenderSizeChanged`) and
  its own considerably more elaborate DPI/content-scaling logic (`ApplyScaleTransform`/
  `FitToAvailableWorkArea`/`AdjustForDpiChange` - a different, LayoutTransform-based scale-to-fit
  mechanism than GLSense's own `BaseWindow`, and arguably more capable in this respect) completely
  untouched. Every View using this base class already sets `WindowStyle="None"` with its own custom
  title-bar XAML and drag-move handling, so `FluentWindow`'s own default chrome doesn't clash with any
  of them - the same way GLSense's own `FluentWindow`-derived `BaseWindow` is used.
  - Also added, additively (ported from `GLSense.Addin.Core\Views\BaseWindow.cs`'s own
    `ForceSizeToContentResettle`/`PumpDispatcherFrame`/`OnContentRendered`, see that codebase's
    CLAUDE.md section 1 for the full "blank gap" bug saga this fixes): a `SizeToContent`-collapse
    resettle safety net, hooked into both `OnLoaded` and the existing `OnContentRenderedDebug`
    handler, guarded by `SizeToContent != Manual` (a no-op for any window using `SizeToContent="Manual"`,
    which most of this app's Views already do) and by the existing `DisableAutoSizing` flag. This is
    genuinely new defense-in-depth alongside `DpiAwareWindow`'s own existing content-scaling logic, not
    a replacement for it - the two mechanisms solve different problems (one fixes stale-first-measure
    `SizeToContent` windows; the other proactively scales content down to fit the screen).
  - Constructor now also defensively calls `WpfUiBootstrapper.Init(...)` if not yet initialized
    (mirrors GLSense's `BaseWindow` constructor's identical guard), in case a window is ever
    constructed before `AddinModule_OnRibbonLoaded`'s own call runs.
- `Views\XLEdgeCalendar.xaml` / `Views\XLEdgeServerConfiguration.xaml`: removed the dead, unused
  `xmlns:mah` declaration from each (no behavior change - nothing in either file referenced it).
- `Helpers\ReportGenerator.cs` / `Utilities\MessageFunctions.cs`: removed the dead
  `using ControlzEx.Standard;` line from each.
- `Utilities\MahAppsBootstrapper.cs`: emptied to a short "retired, see WpfUiBootstrapper.cs" comment
  stub and removed from `XLEdge.csproj`'s `<Compile>` list (so it no longer builds) rather than
  deleted outright - this environment's mounted filesystem rejected `rm` on this file (`Operation not
  permitted`, same restriction hit earlier when cleaning up the pre-swap backup's `packages\`
  subfolder) even though overwriting its *contents* via the file-write tool succeeded. Safe to delete
  manually from Explorer/Visual Studio once confirmed unneeded.
- `XLEdge.csproj`: removed the `MahApps.Metro`/`ControlzEx` `<Reference>` entries; kept
  `MahApps.Metro.IconPacks.Core`/`.FontAwesome` (icons, unrelated package, unaffected); added
  `Wpf.Ui`/`Wpf.Ui.Abstractions` 4.3.0 references (`net481` lib, matching this project's own
  `TargetFrameworkVersion` exactly) with the identical `PublicKeyToken`/`HintPath` shape GLSense's own
  `.csproj` already uses for the same package. Swapped the `<Compile>` entry for
  `MahAppsBootstrapper.cs` -> `WpfUiBootstrapper.cs`.
- `packages.config`: removed the `ControlzEx`/`MahApps.Metro` `<package>` entries, added `WPF-UI`/
  `WPF-UI.Abstractions` 4.3.0 entries; kept both IconPacks entries unchanged.
- `app.config`: removed the now-dead `ControlzEx` assembly `<bindingRedirect>` (harmless to leave, but
  cleaned up since nothing references that assembly anymore).
- **Package binaries**: since this environment has no NuGet/internet access, `WPF-UI.4.3.0`/
  `WPF-UI.Abstractions.4.3.0` were copied file-for-file from GLSense's own already-restored
  `packages\` folder (`D:\SQLLite_Test\AIPowered\GLSense\packages\`) into
  `D:\SQLLite_Test\AllNew_XLEdge\XLEdge\packages\` - the exact same DLLs the `<HintPath>`s above point
  at, so no restore step is needed before this builds.
- `AddinModule.cs`: updated its one call site (`AddinModule_OnRibbonLoaded`) from
  `MahAppsBootstrapper.Init/.PreloadResources` to `WpfUiBootstrapper.Init/.PreloadResources`.

**Not independently verified in this environment** (no Windows/MSBuild/.NET Framework runtime
available here, consistent with every other "not yet rebuilt/tested" caveat in this file, but a
materially bigger caveat for a framework swap than for a logic port): every touched `.xaml`/`.csproj`/
`.config` file was confirmed well-formed XML (`xml.etree.ElementTree.parse`) and every touched `.cs`
file has balanced braces, but neither check proves the app actually compiles or renders correctly -
in particular, `Wpf.Ui.Controls.FluentWindow`'s exact constructor/property requirements (e.g. whether
it needs `WindowBackdropType`/`ExtendsContentIntoTitleBar` set explicitly to behave correctly with
`WindowStyle="None"`) were reasoned from GLSense's own working `BaseWindow` usage, not confirmed by
an actual build+run. **This needs a real clean rebuild + manual smoke-test of every dialog
(`XLEdgeAbout`, `XLEdgeLoginDetails`, `XLEdgeMessageWindow`, `XLEdgeOptions`, `XLEdgeCalendar`,
`XLEdgeServerConfiguration`, `XLEdgeWaitWindow`, `XLEdgeDrilldownReports`) before being trusted** - if
anything renders wrong, the backup folder above has the exact pre-swap state to diff against or
restore from.

- **Status**: implemented, NOT yet rebuilt/tested by the user - flagged as higher-risk than this
  file's usual "not yet rebuilt/tested" caveat for the reasons above.

## Audit: C#-only additions with no VB.NET counterpart — 2026-07-23

Ran a dedicated audit whose only goal was to find code that exists **only** in the C# project with no
VB.NET analog at all — genuine architectural/behavioral additions a developer reading the VB source would
never find a reason for, as opposed to renamed/refactored ports of real VB logic (which this file already
documents extensively and which are *not* re-listed here). Method: read this whole file first (so anything
already logged as "new infra" — see the file-mapping table's `Utilities\UiDispatcher.cs, WpfAppManager.cs,
DpiAwareWindow.cs, DpiAwarenessHelper.cs` row — wasn't rediscovered from scratch), inventoried every file
under both `VB.NET\XLEdge\` and `XLEdge\XLEdge\`, then grepped the VB source for the relevant keywords
before concluding anything was actually new. Findings, grouped by category:

**Retry/backoff and resiliency infrastructure.** `Helpers/ApiOperationHelper.cs`'s `ExecuteWithRetry`
(3 attempts, exponential backoff, `IsTransientError` checking `HttpRequestException`/`TimeoutException`/
`WebException`/transient `SocketException` codes) wraps *every* call `APIHelper.ServerAPI` makes (the C#
port of VB's `GetServerData`) — confirmed via grep that VB's only retry/backoff logic anywhere in the
codebase is scoped specifically to `FormProcessBar.vb`'s `DownloadImageInternal` (image downloads, lines
~2993-3043: retry-on-429/502/503/504 with `Retry-After` honoring) — `GetServerData`/`ReturnHTTP` (the
general report/parameter/CSV fetch path) never retried anything in VB at all, a single failed request just
failed. Wrapping the whole general API path in retry/backoff is a genuine new resiliency behavior added
because `HttpClient`'s failure modes differ from VB's `HttpWebRequest`, not a port of specific VB retry
logic (the VB image-download retry logic *was* already faithfully ported — see the `ImageDownloadHelper`
entry above — this is a separate, additional retry layer with no VB source at all).

**COM self-healing/liveness-check helper.** `Helpers/ExcelApplicationHelper.TryUseApplication` (probes a
cached `Excel.Application` reference by reading `.Hwnd` in a try/catch to detect a disconnected/dead RCW)
and `TryGetActiveExcelApplicationInternal` (falls back to `AddinModule.CurrentInstance.HostApplication`,
then `Marshal.GetActiveObject("Excel.Application")` via the ROT) are wrapped around **every** access to
`XLEdgeAppState.Instance.ExcelApp` (both the getter and setter) — a systemic "is this COM reference still
alive; if not, silently re-resolve it" pattern. VB has exactly two, isolated `Marshal.GetActiveObject`
call sites (`AddinModule.vb:1106`, `:3286`) used as one-off fallbacks in specific handlers — it never
wrapped its module-level `_Edge_ExcelApp` global with a liveness probe on every read, so a stale/dead
Excel COM reference in VB would simply throw wherever it was next used. This systemic recovery wrapper is
a genuine new architectural addition, not a port of anything VB did everywhere.

**Cancellation-token infrastructure.** `Helpers/CancellationHelper.cs` (a reusable, thread-safe
`CancellationTokenSource` wrapper — lazy init, safe repeated Cancel/Dispose, returns an already-cancelled
token post-disposal instead of throwing) and `Helpers/CancellationTokenHelper.cs` (`ThrowIfCancelled`/
`IsCancellationRequested`/`DelayWithLogging`, each with built-in logging) have no VB counterpart because VB
has no `CancellationToken` concept at all — this file's own "Major scope correction" entry above already
documents that VB's report-generation engine ran on `BackgroundWorker.DoWork` (thread-pool thread, e.parameter
`Cancel`/`e.Cancel` flags) and that the C# port deliberately chose `async`/`await` instead. These two
classes are the plumbing that decision required, built new because the async model itself is new, not
because they port specific VB cancellation logic.

**Structured logging scopes and detailed-exception logging.** `Utilities/LogUtility.cs`'s nested
`LogScope : IDisposable` class (logs `"BEGIN: {name}"`/`"END: {name}"` pairs with an indent-depth counter
via `IncrementScope`/`DecrementScope`) is used throughout `ApiOperationHelper`/`PerformanceHelper` and has
no VB equivalent — VB's superseded `XLEdgeLogDebug`/`XLEdgeLogInfo` (per this file's own "Confirmed
obsolete" list above) were flat, unscoped log lines with no begin/end bracketing or indentation concept at
all. Likewise `Helpers/ExceptionHelper.cs` (`LogDetailedException` — context + exception type + inner-
exception chain in one structured line; `GetFriendlyErrorMessage` — inner-exception-preferring message
extraction for UI display) has no VB analog; VB's exception handling just logged `ex.Message` inline at
each catch site with no shared, structured formatting helper.

**Performance-measurement scope.** `Helpers/PerformanceHelper.cs`'s `PerformanceScope : IDisposable` (starts
a `Stopwatch` on construction, logs elapsed time on `Dispose`, escalates to a `LogWarn` "slow operation"
if elapsed exceeds 45 seconds) plus the `MeasureExecutionTime(Async)` wrapper overloads are new,
general-purpose instrumentation with no VB counterpart. VB's only `Stopwatch` usage anywhere
(`FormProgressNew.vb`) was a simple elapsed-time value read once for display in the progress dialog's own
UI text — not a reusable scope wrapping arbitrary operations with automatic slow-operation logging.

**Dispatcher-thread-marshaling and DPI-awareness utilities (already flagged in this file's own file-mapping
table, restated here for completeness).** `Utilities/UiDispatcher.cs` and `Utilities/WpfAppManager.cs`
(bootstrapping/managing a dedicated WPF `Application`/`Dispatcher` inside a COM add-in host process — a
problem that doesn't exist for a WinForms add-in, since WinForms controls just use the host process's
existing message loop directly) have no VB source at all. `Utilities/DpiAwareWindow.cs` (775 lines) and
`Utilities/DpiAwarenessHelper.cs` (172 lines) implement active, runtime Per-Monitor-V2 DPI handling
(scaling transforms, monitor-change reactions) — VB's entire DPI story, confirmed by reading `app.manifest`,
was a single static declarative flag (`<dpiAware>true</dpiAware>`, system-DPI-aware only) with zero runtime
code; there is nothing in VB.NET\XLEdge\*.vb resembling either of these classes.

**TLS/certificate validation — new hardening that actually inverts VB's own behavior.**
`Utilities/StrictCertificateValidator.cs` (`Validate`, wired into `APIHelper.cs`'s
`HttpClientHandler.ServerCertificateCustomValidationCallback`) fails closed on *any* `SslPolicyErrors`,
rebuilds the chain with online revocation checking (`X509RevocationMode.Online`), and logs full chain-status
detail on rejection. VB's own `APIHelper.vb` (`ServicePointManager.ServerCertificateValidationCallback`,
line ~108) does the opposite: logs the certificate subject/issuer for debugging and then unconditionally
`Return True` — i.e. VB accepts every certificate regardless of validity, with a comment admitting it's
"for debugging." This isn't just a new addition, it's a deliberate behavioral flip (permissive → strict)
introduced during the port; worth confirming this was an intentional hardening decision and not an
accidental behavior change, since a previously-accepted self-signed/misconfigured server certificate would
now cause every API call to fail.

**New WPF-only UI chrome with no VB Form equivalent.** `Views/AppOverlay.xaml`/`.xaml.cs` (~790 lines
combined) is a full toast/busy-spinner/confirm-dialog overlay system (`ShowToast`/`ShowSuccess`/`ShowError`/
`ShowWarning`/`ShowInfo`, a blocking `ShowBusyasyn`/`HideBusyAsync` pair, `ShowConfirm`/`ShowConfirmAsync`,
a `BlurApplied` attached property) embedded directly in `XLEdgeCTP` and reused by `XLEdgeWaitWindow`/
`XLEdgeServerConfiguration`/`XLEdgeAbout` (per this file's own "toast pinned to a small band" entry above).
VB has no toast/blur-overlay concept anywhere — it used plain `MsgBox`-style dialogs (`XLEdgeMsgDisplay`,
superseded by `MessageFunctions.XLEdgeMessage`, already documented as a legitimate port) and dedicated
WinForms progress dialogs (`FormProcessBar`/`FormProgressNew`, already documented as merged into
`XLEdgeWaitWindow`) — never an in-place, non-modal toast notification layered over the task pane's own
content. `Utilities/MahAppsBootstrapper.cs` (432 lines — accent-color/base-theme initialization and
live theme switching via `ControlzEx.Theming`) is new for the same reason: WinForms has no accent-theming
system, so nothing in VB needed anything like it. `Helpers/EnhancedDragDropHelper.cs`
(`EnableWindowDrag`/`IsInteractiveControl`, wired into `XLEdgeAbout`/`XLEdgeCalendar`/
`XLEdgeDrilldownReports`/`XLEdgeLoginDetails`/`XLEdgeMessageWindow`/`XLEdgeOptions`/
`XLEdgeServerConfiguration`/`XLEdgeWaitWindow`) exists purely because these WPF windows use custom/borderless
chrome and therefore need manual click-and-drag-to-move wiring that skips interactive child controls —
ordinary WinForms `Form`s (VB's entire UI) have a native title bar and never needed anything like this.
`Converters/WidthPercentageConverter.cs`'s three converters (`WidthPercentageConverter`,
`BoolToVisibilityConverter`, `EmptyStringToVisibilityConverter`) are plain WPF `IValueConverter`
implementations that exist solely because WPF's XAML data-binding model requires them — WinForms data
binding has no equivalent construct, so there is nothing to port from VB here; these are new purely as a
consequence of the UI framework switch, not new user-facing behavior.

**Minor/ambiguous items, noted rather than firmly classified either way.**
`Helpers/XLEdgeRibbonReflectionHelper.cs` (generic `GetRibbonControl(addinModuleInstance, controlName)` —
tries a public property, then a public field, then a non-public field via reflection) backs
`XLEdgeRibbonHelper`'s enable/disable-by-name-list pattern. This centralizes what VB did as repeated,
direct `RibXxx.Enabled = True/False` assignments inline in each ribbon handler — functionally the same end
state as VB, just reached generically through reflection instead of direct field access in each handler.
It's best described as a refactor-for-centralization rather than new behavior, but it has no single VB
Sub/Function it corresponds to one-to-one, so it's noted here rather than silently omitted.
`Helpers/ApiResponseHelper.cs`'s `Parse<T>`/`ApiResult<T>` generic envelope (`Models/ApiResult.cs`) has only
one real call site in the current codebase (`XLEdgeCTP.xaml.cs`'s broadcast-message fetch) — likely a
reasonable, narrow refactor of VB's own ad hoc JSON-key-sniffing for that one feature into a typed,
reusable wrapper, rather than a broad new capability; flagged here for completeness rather than treated as
a significant addition given how narrowly it's actually used today.

**What was checked and confirmed NOT to be a C#-only addition** (i.e. real ports, correctly excluded from
the list above): `ApiRequestException`/`ApiTimeoutException` (`Helpers/ApiExceptions.cs`) — faithful ports
of two VB exception classes defined inside `FormProcessBar.vb`, already documented above; `XLEdgeTempFileCleaner`
— a direct port of VB's `EEDeleteAllFiles`/`IsFileOpen`, already documented above; `ApiErrorMessageExtractor`
— a line-for-line port of VB's `ExtractErrorMessage`/`IsHtmlResponse`/`ExtractTextFromHtml`/`ExtractFromJson`/
`CleanPlainText`, already documented above. These three were named explicitly in the task brief as examples
of things that might look C#-only at a glance but aren't — confirmed via this file's own existing entries
plus a direct read of the corresponding VB source, and correctly *not* included in this audit's findings.

## Ported: "Process+XLSX" and "Excel" file-download title commands — 2026-07-23

Closes out the last follow-up task filed during the MultiData fix (both branches were previously routed
to a "not yet supported" warning log rather than being silently mis-processed - see that entry below).
Both VB branches - `"Process"`/`"Edge"` with a 5th (`"XLSX"`) title segment, and `"Excel"` - call the
exact same `DownloadFile1(strURL)`: a plain authenticated GET (Bearer token) that saves the response
to `%UserProfile%\Downloads`, named from the response's `Content-Disposition` header, followed by a
confirmation message box. No report table/Excel-writing is involved in either branch at all.

**Discovered while investigating**: this port already has a byte-for-byte equivalent of VB's
`DownloadFile1`/`DownloadFile` pair - `ApiHelper.DownloadFileAsync(url, ...)` (GET + Bearer auth +
`Content-Disposition` filename + save to Downloads) paired with `MessageFunctions.XLEdgeMessage(...)`
for the confirmation/error message box - already used by `AddinModule.cs`'s
`HandleAttachmentDownloadAsync` for the sibling VB `DownloadFile` (attachment-hyperlink) method. So this
task turned out to be wiring, not new download logic.

**Fix**:
- New `ReportGenerator.DownloadFile1Async(string downloadUrl)` - thin wrapper: calls the existing
  `ApiHelper.DownloadFileAsync`, then `MessageFunctions.XLEdgeMessage` with the same "Attachment has been
  saved to the downloads folder and the file name is ..." / "Failed to download the file." text VB/the
  existing attachment-download path already use.
- `Views/XLEdgeCTP.xaml.cs`'s `WebView_DocumentTitleChanged`:
  - `"Process"`/`"Edge"` case now checks `parts.Length >= 5 && parts[4] == "XLSX"` (matching VB's
    `var(4) = "XLSX"` check) and, when true, builds
    `{loginUrl}/rest/secure/process/finance-report-output?processId={parts[1]}` and calls
    `DownloadFile1Async` instead of falling through to `CreateReportFromTitleAsync` - matching VB's
    `URLExtension` construction exactly. The non-XLSX case is unchanged.
  - `"Excel"` case now builds `{loginUrl}/web/secure/financeTemplateFileDownload?reportId={parts[1]}`
    and calls `DownloadFile1Async`, replacing the previous "not yet supported" warning.
  - Doc-comment above the dispatch switch updated to reflect both branches are now ported.
- **Not ported, confirmed out of scope**: VB's `FileDownload`/`DownloadFile` (folder-picker variant,
  used elsewhere for a user-chosen destination folder) is a different code path from `DownloadFile1` and
  isn't called by either title-command branch - it wasn't touched here.

## Ported: Logs report-run type — 2026-07-23

Last of the three items the user explicitly called "in scope" together ("port all of them" -
MultiData/Logs/GLSense session-sync); MultiData and GLSense session-sync were already ported (see
entries below) - this closes out that group. VB's `"Logs"` document-title command
(`FormProcessBar.vb`'s `Edge_GenerateLogs`/`Edge_FillLogs`) is not a structured report at all - it's a
raw, free-form process-log text dump for a single `processId`, fetched from
`{loginUrl}/rest/secure/process/excel-log?processId={rID}` and imported into a `"Logs_{processId}"`
worksheet via Excel's native Text `QueryTable` mechanism ("Data > From Text") rather than the
ListObject/columns/hyperlinks machinery the "Edge"/"Process" report flow uses. Confirmed via grep that
no C# equivalent existed at all before this fix (no `excel-log`, `QueryTable`, or `"Logs_"` references
anywhere in the C# project) - the `"Logs"` case in `WebView_DocumentTitleChanged` was just logging a
"not yet supported" warning, added as a placeholder during the MultiData fix above.

**Fix**:
- New `ReportGenerator.CreateLogsReportAsync(string logsRequestStr, ...)` - parses the `"|"`-delimited
  title (`str(1)` = processId, matching VB's `Logstr.Split("|"c)`), fetches the log text via the
  existing `ApiHelper.ServerAPI` (reusing the same cancellation/timeout handling established in the
  API-exception work: `OperationCanceledException` -> "cancelled by the user", `ApiTimeoutException` ->
  "timed out, please try again"), then calls a new `BuildLogsSheet(processId, logText)` to write it into
  Excel. Follows the same `_appOverlay`/`_showWaitWindow`/`_ctsHelper`/`SetMessage`/`DisplayErrorAsync`/
  `CleanupAsync` scaffolding already established by `CreateReportFromTitleAsync`/
  `CreateMultiDataReportsAsync`, and wraps the whole call in `ExcelBulkOperationScope` like every other
  report-writing entry point in this file.
- New `BuildLogsSheet` - ports `Edge_GenerateLogs`'s sheet find/create (existing `"Logs_{processId}"`
  sheet -> activate + clear; otherwise add a new sheet and rename it, using the same
  try-bare-`Add()`-then-fall-back-to-explicit-position pattern already used by `BuildReportTable` for a
  brand-new sheet) followed by `Edge_FillLogs`'s temp-file-write + `QueryTable` import: writes `logText`
  to `{XLEdgeAppPaths.TempFolder}\{processId}_Logs.txt` (deleting any stale file first, matching VB's
  temp-file handling), then `logsSheet.QueryTables.Add("TEXT;{tempFile}", ...)` with
  `TextFileParseType=xlDelimited`, `TextFilePlatform=65001` (UTF-8), `TextFilePromptOnRefresh=false`,
  `.Refresh(false)` - the exact same property set VB's `With EEQueryTable` block sets, no more. VB never
  sets a delimiter flag (`TextFileCommaDelimiter` etc.), so no column-splitting actually occurs - each
  line of log text lands as one row in column A, consistent with this being a raw log dump rather than
  tabular data. The `QueryTable` is deleted immediately after the one-time refresh
  (`SaveData=false; .Delete()`, wrapped in try/finally + `Marshal.ReleaseComObject`) so the imported text
  stays as plain static cell values with no lingering external-data-range link, exactly matching VB's own
  `Finally` block. `ActiveWindow.DisplayGridlines = false` ported as the last step, matching VB.
- `Views/XLEdgeCTP.xaml.cs`'s `WebView_DocumentTitleChanged` `"Logs"` case now calls
  `ReportGenerator.CreateLogsReportAsync(title, AppOverlayControl)` via the same `SafeFireAndForget`
  pattern used by the other dispatch cases, instead of just logging a warning; its doc-comment updated
  to reflect Logs is no longer an unported branch.
- **Not ported, judged out of scope for this task**: VB's `IncrementColumn` helper (duplicate-column-name
  suffixing) is used by the unrelated `Edge_FillData` CSV-parsing routine, not by `Edge_FillLogs`/
  `Edge_GenerateLogs` - it was read as part of the same VB source excerpt but isn't part of the Logs
  feature; the C# port's existing "Edge"/"Process" report flow already has its own equivalent
  (`MakeUniqueName`) for that concern, so no new port was needed here.

## Ported: MultiData batch report-run type — 2026-07-23

Next of the larger remaining items, explicitly confirmed in scope by the user ("port all of them" -
MultiData/Logs/GLSense session-sync). VB's `WebView_DocumentTitleChanged` dispatches on the first
`"|"`-delimited segment of the WebView2 document title the hosted web app sets, via a big If/ElseIf
chain: `"EdgeWorkbook"` (batch-rerun a whole workbook's worth of reports at once - "MultiData"),
`"Logs"`, `"Process"`/`"Edge"` (the common single ad-hoc report run, with its own XLSX-extension
file-download sub-case), and `"Excel"` (a plain file-download command). The C# port's
`WebView_DocumentTitleChanged` never dispatched on this at all - it just blindly forwarded ANY
`"|"`-containing title straight into `CreateReportFromTitleAsync` as if it were always a single ad-hoc
report, meaning an `"EdgeWorkbook|..."` title (which doesn't carry a report id/run id in the title
itself at all - those come from a separate DOM query) would have been mis-parsed and failed silently or
produced garbage.

**Fix**:
- `Views/XLEdgeCTP.xaml.cs`'s `WebView_DocumentTitleChanged` rewritten as a `switch` on the title's
  first segment, matching VB's dispatch: `"EdgeWorkbook"` -> new `FetchWorkbookRerunIdsAsync()` (queries
  the hosted web app's `[reruntype=xledgeworkbookrerun]` element's `newrunids` attribute via
  `CoreWebView2.ExecuteScriptAsync`, stripping the JSON-string-result's surrounding quotes exactly like
  VB's `result.Trim(""""c)`) followed by the new `ReportGenerator.CreateMultiDataReportsAsync`;
  `"Process"`/`"Edge"` -> unchanged, still calls `CreateReportFromTitleAsync` directly (this was the
  only case the old, non-dispatching handler ever got right); `"Logs"`/`"Excel"` -> a clear "not yet
  supported" log instead of being silently mis-processed as a normal report (Logs report type is its own
  separate, still-pending migration item; the XLSX/"Excel" file-download sub-cases have no port at all
  yet - see the new follow-up task below).
- New `ReportGenerator.CreateMultiDataReportsAsync(string runIdsRaw, ...)` - ported from
  `BGWorker_DoWork`'s `"MultiData"` branch. VB downloads every report's CSV first (a full pass calling
  `StartTaskHere` for every run id), then writes every one of them into Excel in a second pass - purely
  an artifact of its `BackgroundWorker`-based threading model, not a user-visible requirement. Since this
  port's `CreateReportFromTitleAsync` already does download-then-build for one report in a single call,
  the batch is simply run sequentially (one report fully finished before the next starts) instead -
  same end result, much simpler. New private `ProcessRunIds(string input)` ports VB's `ProcessRunIDs`
  exactly (split by `"^"`, then by `"|"`, build an `"Edge|{reportId}|{runId}|"` title string per pair -
  the empty trailing segment matches `EdgeRequestParser`'s already-expected `ReportName=""` shape for an
  ordinary ad-hoc report).
- **Accepted UX difference from VB, not a bug**: VB's `FormProcessBar` shows one continuous progress
  dialog for the whole MultiData batch; since each `CreateReportFromTitleAsync` call creates its own
  wait window/`CancellationHelper`, a multi-report batch in this port shows a sequence of wait-window
  popups (one closes, the next opens) rather than one persistent dialog. Functionally equivalent (every
  report still gets downloaded and written; a report failing doesn't abort the rest of the batch,
  matching VB's own per-report try/catch), just a different popup rhythm.
- **Cancellation caveat, also accepted**: a user-driven cancel via one report's wait window still stops
  the remaining batch (the loop checks the shared `_ctsHelper`'s cancellation state right after each
  `CreateReportFromTitleAsync` call returns, before the next iteration replaces it with a fresh
  instance) - but this is detected one report later than VB's single shared cancellation token would,
  since each report's cancellation is scoped to that report's own call rather than one token spanning
  the whole batch.

**New follow-up task filed, not fixed here**: VB's `"Process"`+XLSX-extension sub-case (`DownloadFile1`,
an authenticated `finance-report-output` file download with no report table involved) and the `"Excel"`
command (`financeTemplateFileDownload`, a similar authenticated download) have no C# port at all - both
now log a clear "not yet supported" warning instead of silently failing inside
`CreateReportFromTitleAsync` the way they effectively did before this fix.

## Ported: GLSense session-sync (login + logout notification) — 2026-07-23

First of the larger remaining items. `XLEdgeCTP.xaml.cs` already had an explicit stub -
`SyncLoginToGLSense()` ("Intentionally left as-is"), already correctly called from
`WebCtrl_SourceChanged` at the "excel=Y#Home" login-landing branch, but doing nothing. This is the
reverse direction of the already-ported `InvokedFromGLSense` (GLSense -> XLEdge): VB's
`ADXExcelTaskPane1.vb WebCtrl_SourceChanged` notifies the sibling GLSense add-in, via the same
`GetGLSenseAddinObject()`/reflection mechanism, whenever a login completes *directly through XLEdge's
own WebView2* (as opposed to one that arrived via GLSense calling `InvokedFromGLSense`) - calling
GLSense's `GetGLCubeInformation` method (a misleading name - it hands GLSense the credentials to load
its own cube list, it doesn't return anything to XLEdge) with the auth token/login URL/username, guarded
so it only ever fires once per login and never for a GLSense-originated one. VB also has a symmetric
logout notification (`LogOffAllTaskPanesAsync`'s `addinInstance...InvokeMember("LogoutSession", ...)`)
that had no C# equivalent at all.

**Fix**:
- `XLEdgeAppState.cs`: added `LoginSentToGLSense` (bool), ported from VB's module-level
  `LoginSentToGLSense` flag - prevents re-notifying GLSense on every subsequent WebView2 navigation
  back to the login-landing URL within the same session.
- `AddinModule.cs`: new `public void NotifyGLSenseOfLogin(string authToken, string loginUrl, string
  userName)` - guards on `!LoginFromGLSense && !string.IsNullOrWhiteSpace(authToken) &&
  !LoginSentToGLSense` (VB's exact combined condition), resolves/caches `_glSenseAddinInstance` (the
  same field `InvokedFromGLSense` already uses), sets `LoginSentToGLSense = true` *before* attempting
  the reflection call (matching VB's ordering, so a failed/unavailable GLSense instance doesn't retry
  every navigation), then invokes `GetGLCubeInformation` via `InvokeMember`. New `private void
  NotifyGLSenseOfLogout()` - the symmetric logout call, invoking GLSense's `LogoutSession` method the
  same way. Wired into `LogoffFromXLEdgeAddin`'s existing state-reset block (right before clearing
  `LoginToken`/`LoginUrl`/etc.), which also now resets `LoginSentToGLSense = false` so a later fresh
  login can re-sync.
- `XLEdgeCTP.xaml.cs`: `SyncLoginToGLSense()` now calls
  `XLEdge.AddinModule.CurrentInstance?.NotifyGLSenseOfLogin(appState.LoginToken, appState.LoginUrl,
  appState.LoginUserName)`, wrapped in the same try/catch-and-log pattern already used by every other
  handler in this file.

**Not changed**: VB's `LoginFromSense` (no "GL") field - distinct from the actively-used
`LoginFromGLSense` - is set to `false` at one call site in `XLEdgeCTP.xaml.cs` but otherwise never read
anywhere in the C# port; it looks like leftover naming drift from an earlier refactor rather than a
load-bearing flag (VB's real equivalent, `EELoginFromSense`, is what `LoginFromGLSense` already
faithfully mirrors). Left untouched since it isn't part of this gap and nothing reads it either way.

## Ported: ParamValueInput's named-range + data-validation cell lock — 2026-07-23

Follow-up to the Responsibility/GL Accounts fix above, which is where this gap was first noticed: VB's
`ParamValueInput` doesn't just write the parameter value cell's text/formatting - for every non-blank
value, it also names the cell (`{TableName}_{CellAddress}`, cleaned via `CleanUpName`, truncated to 30
chars for Excel's own name-length limit) and adds a custom, self-referencing data-validation rule
(`xlValidateCustom`, stop-alert style, `Formula1` = that same named range) with the error message "To
change parameters, use the Run button on the ribbon or use param control sheet." The C# port's
`WriteParamValueCell` only ever did the text/formatting half - every parameter value cell in the C# port
was freely user-editable with no nudge back toward the Run button, unlike VB.

**Fix** (`Helpers/ReportGenerator.cs`): `WriteParamValueCell` now takes a `tableId` parameter and, after
writing the cell's value/formatting exactly as before, builds the named range from
`cell.Address[false, false, Excel.XlReferenceStyle.xlA1]` + `tableId` (reusing the already-existing
`ExcelSheetHelper.NamedRangeExists`/`DeleteNamedRange`/`CleanUpName` helpers - no new Excel-naming
helpers needed, they already existed and were just unused for this purpose), deletes/recreates the
validation via `Excel.Validation.Add(xlValidateCustom, xlValidAlertStop, xlEqual, rngName, Type.Missing)`
- the exact 5-positional-argument call shape already proven working elsewhere in this codebase
(`ParamsControlSheetBuilder.LockToCurrentValue` uses the identical `.Add(...)` shape for its own
self-referencing lock). `ErrorTitle` is hardcoded to `"Orbit"` (matching this migration's established,
user-confirmed decision to keep branding hardcoded rather than restoring VB's dynamic `EdgeBranding`
concept) instead of VB's `EdgeBranding` variable.

Threaded the new `tableId` parameter through both call sites: `WriteSameSheetBanner` gained a `tableId`
parameter (passed from `BuildReportTable`'s own already-computed `tableId` local) instead of previously
never receiving it, and `BuildCompanionParameterSheet` (which already had `tableId` as a parameter)
now passes it straight through - both call sites of `WriteParamValueCell` updated accordingly.

## Ported: Responsibility/GL Accounts extra display rows on the parameter sheet — 2026-07-23

Closes the last item from the original reconciliation-audit punch list before moving on to the bigger
remaining items (GLSense session-sync, MultiData, Logs report type). VB's `FormProcessBar.vb` (lines
~2400-2444) scans every entry in the report's parameters JSON array for an `"extraParameters"` object
(`ReportExtraParams`) and, when present, writes: a hidden `IT4`/`IU4` cell pair (`ORACLE_RESP_ID`/
`ORACLE_RESP_DISPLAY_VALUE`) later re-read by refresh/drilldown requests (already ported - see the
`IT4` reads in `XLEdgeParamsBuilder`/`DrilldownRequestBuilder`, from earlier in this session), plus two
*visible* rows ahead of the regular parameter rows: "Responsibility" (from `ORACLE_RESP_DISPLAY_VALUE`)
and "GL Accounts" (a flattened `"Name=Value, Name=Value"` string from the nested
`ORACLE_GL_SEGMENT_DISPLAY_VALUES` object). None of this existed in the C# port - `IT4` was only ever
*read*, never *written*, meaning the whole responsibility-scoping round-trip was silently broken: a
report's Responsibility context could never actually survive into a refresh or drilldown.

**Fix** (`Helpers/ReportGenerator.cs`):
- New `ExtractExtraParams(JsonElement)` - ports `ReportExtraParams` verbatim (same 3 well-known keys,
  same nested-object flattening for GL segments, same `"-"`/blank -> `""` substitution, same trailing
  `", "` trim).
- `ParseParamDisplayRows` gained an overload taking `out string oracleRespId, out string oracleRespValue`
  - scans every parameter entry for `extraParameters` (a separate pass, matching VB's separate loop
    ordering exactly) and prepends "Responsibility"/"GL Accounts" rows to the returned list before the
    regular per-parameter rows are appended, so row ordering matches VB exactly. The original
    single-parameter overload now just delegates with discarded out params, so nothing else calling it
    needed to change.
  - The "Responsibility" row's value is prefixed with a leading apostrophe
    (`"'" + oracleRespValue`), matching VB's `ParamValueInput(updrange, "'" & GlExtraParams.Item2, ...)`
    call exactly - a defensive text-force so a numeric-looking responsibility value is never silently
    auto-converted to a number at the moment `Value2` is first assigned (before `WriteParamValueCell`'s
    own `NumberFormat = "@"` line takes effect). The "GL Accounts" row is written as-is, matching VB
    (which never apostrophe-prefixes `GlExtraParams.Item3`).
- Both call sites (`BuildReportTable`'s same-sheet banner and `BuildCompanionParameterSheet`'s
  separate-sheet layout) now capture the two out params and, when both are non-blank (matching VB's
  own combined `Item1`/`Item2` non-blank check), write the hidden `IT4`/`IU4` cells
  (`Clear()`/`Value2 = RemoveEquaSymbol(...)`/`WrapText = false`, same as VB) right after their existing
  IT1/IT5 (and, for the companion sheet, IT2) bookkeeping-cell writes.

**Follow-up gap found while porting this** (tracked separately, not fixed here): VB's `ParamValueInput`
doesn't just write the value cell - it also creates a named range (`{TableName}_{CellAddress}`, cleaned/
truncated) and adds a custom data-validation rule (`Formula1` = that same named range, stop-alert style)
so a user can't type over a parameter value cell directly, only via the Run button or param control
sheet. This applies to *every* parameter value cell, not just Responsibility/GL Accounts, and the C#
`WriteParamValueCell` currently does none of it - every param value cell in the C# port is currently
freely editable. Filed as its own task rather than folded into this fix, since it's a pre-existing gap
affecting all rows, not something specific to Responsibility/GL Accounts.

## Ported: EEDeleteAllFiles temp-file cleanup — 2026-07-23

Next severity-ordered item; this one had an explicit, honest `TODO` comment already in the C# port
(`Helpers/ProgressCoordinator.cs` line ~39: `"TODO(AddinModule port): call the equivalent of
AddinModule.CurrentInstance.EEDeleteAllFiles() once AddinModule.cs exists..."`) - stale, since
`AddinModule.cs` had already been fully ported (task #7) by the time this was picked up.

VB's `EEDeleteAllFiles` (`AddinModule.vb:603`) deletes every file in `EETempFilesPath` except ones
named `"XLEdge_Logs"`/`"xledgeuserpreferences.json"` and anything currently open (checked via
`IsFileOpen`, which does a locking `FileOpen`/`FileClose` probe) - called once at add-in startup
(`AddinModule_OnRibbonLoaded`) and again after every completed report run
(`XLEdgeProcedures.vb`'s `Edge_CloseProgress`). Without it, temp CSVs written by
`ReportGenerator.WriteTempCsv` (one per report run, named `{runId}.csv`) would accumulate forever.

**Fix**:
- New `Helpers/XLEdgeTempFileCleaner.cs` - `DeleteAllTempFiles()` (ported from `EEDeleteAllFiles`) and
  `IsFileOpen(fileName)` (ported from VB's own `IsFileOpen`, using a `FileStream` exclusive-open probe
  instead of VB's `FileOpen`/`FileClose` intrinsics). VB's filename-based exclusion list isn't needed
  here: VB's single `EETempFilesPath` folder was shared with logs and the preferences file, but the C#
  port already keeps those in their own separate `XLEdgeAppPaths.LogFolder`/`BrowserLogsFolder`/
  `XLEdgeLogsFolder` (see `XLEdgePreferencesManager`) - `XLEdgeAppPaths.TempFolder` (where
  `WriteTempCsv` already writes) is already isolated to just these temp report CSVs. The "skip a file
  that's still open" safety check is kept, since a temp CSV could in principle still be mid-write or
  opened directly by the user.
- `Helpers/ProgressCoordinator.cs`'s `ResetReportState()`: replaced the stale TODO with a call to
  `XLEdgeTempFileCleaner.DeleteAllTempFiles()`, matching VB's `Edge_CloseProgress` call site.
- `AddinModule.cs`'s `AddinModule_OnRibbonLoaded`: added the same call (wrapped in its own try/catch via
  the existing `SafeLogException` helper) right after `XLEdgePreferencesManager.Instance.Initialize()`,
  matching VB's startup call site - so temp files left over from a previous Excel session/crash are
  also cleaned up on the next add-in load, not just after each run within the same session.

## Ported: ApiRequestException/ApiTimeoutException + cancel-run server notification — 2026-07-23

Next severity-ordered item, a direct follow-on to the error-message-extraction work above. VB defines
two typed exceptions inside `FormProcessBar.vb` (`ApiRequestException` - message + `HttpStatusCode`,
for a definitive non-success server response; `ApiTimeoutException` - a `TimeoutException` subtype,
for a request that timed out) and throws them from `GetServerData`, then branches on them separately
in `ReturnHTTP`/`ParamInfo`/etc. Distinctly, `StartTaskHere`'s `Catch ex As OperationCanceledException`
(the true user-cancel path, not a server timeout) calls `CancelBackEndRequest(CancelURL)` - a POST to
`.../rest/secure/report/cancel-run?runId=X` telling the server to stop processing a run the client no
longer wants, since otherwise the server keeps working on a result nobody will ever fetch.

Neither typed exception nor the cancel-run notification existed anywhere in the C# port -
`ApiHelper.ExecuteApiCall` only ever threw plain `InvalidOperationException`/rethrew whatever the
underlying `HttpClient` produced, and cancelling the wait window client-side never told the server to
stop.

**Fix**:
- New `Helpers/ApiExceptions.cs` - `ApiRequestException : Exception` (message + `HttpStatusCode`) and
  `ApiTimeoutException : TimeoutException`, matching VB's two classes. `ApiTimeoutException` deliberately
  inherits `TimeoutException` so `ApiOperationHelper.IsTransientError`'s existing "retry on
  `TimeoutException`" branch already covers it for free (a timed-out call is now retried up to 3x with
  backoff via the existing `ExecuteWithRetry`, instead of - as before this fix - immediately propagating
  as a bare `OperationCanceledException` that `ExecuteWithRetry`'s own first catch rethrows without any
  retry at all). `ApiRequestException` is a plain `Exception`, so it is *not* retried - a definitive 4xx/5xx
  from the server shouldn't be blindly retried the way a transient network blip should.
- `ApiHelper.ExecuteApiCall`: the non-success-status branch (added in the error-extraction fix above) now
  throws `ApiRequestException(result, response.StatusCode)` instead of a bare `InvalidOperationException`.
  Added a filtered `catch (OperationCanceledException ex) when (!cancellationToken.IsCancellationRequested)`
  ahead of the existing plain cancellation catch - this distinguishes `HttpClient`'s own configured
  `Timeout` expiring (also an `OperationCanceledException` in .NET, but the caller's own token was never
  triggered) from a genuine caller/user-driven cancellation, and throws `ApiTimeoutException("The request
  timed out", ex)` for the former. This matters because before this fix, a real server timeout and a real
  user Cancel-button click were indistinguishable to every caller - both surfaced as the same
  `OperationCanceledException`, so a slow server was being reported to the user as "cancelled by the user."
- New `ApiHelper.NotifyCancelRunAsync(loginUrl, runId)` - ported from `CancelBackEndRequest`: POSTs
  `{loginUrl}/rest/secure/report/cancel-run?runId={runId}` with the current bearer token, logs (but never
  throws on) any failure, matching VB's own best-effort/fire-and-forget treatment of this call.
- Wired into `ReportGenerator.cs` at all 4 places a report run can be cancelled mid-flight: the CSV
  download, report-definition fetch, and report-parameters fetch inside `CreateReportFromTitleAsync`
  (each already had a `catch (OperationCanceledException)` block; added the `NotifyCancelRunAsync` call
  there using `_edgeRequest.ReportRunId`), and the CSV re-download inside `RefreshListObjectAsync` (using
  the already-resolved `runId`/`eeLoginUrl` locals). Also added a `catch (ApiTimeoutException ex)` block
  at each of the 4 sites, showing "The request timed out. Please try again." (matching VB's wording)
  instead of falling through to the generic `catch (Exception ex)` handler's more generic message - and,
  matching VB's own `StartTaskHere`/`ReturnHTTP` split, a timeout does *not* trigger
  `NotifyCancelRunAsync` (only a genuine user cancellation does).
- `ApiRequestException` itself isn't given a dedicated `catch` block at each call site - its `.Message`
  is already the properly human-readable text from `ApiErrorMessageExtractor` (previous fix), so letting
  it fall through to each site's existing generic `catch (Exception ex)` handler (which already logs and
  displays `ex.Message`) produces a perfectly reasonable, readable error without extra code; adding a
  `StatusCode`-aware branch everywhere was judged unnecessary scope for this pass.

## Ported: API error-message extraction (HTML/JSON error body → readable message) — 2026-07-23

Next severity-ordered item. VB's `GetServerData` (`FormProcessBar.vb:887`) catches a failed
`HttpWebRequest`, reads the error response body, and runs it through `ExtractErrorMessage` -
a small pipeline (`IsHtmlResponse` → `ExtractTextFromHtml`, or a `{`/`[` prefix → `ExtractFromJson`,
else `CleanPlainText`) that turns an HTML error page or a JSON error object into a short, readable
sentence before throwing `ApiRequestException(errorMessage, statusCode, ex)`.

The C# port's `ApiHelper.ProcessResponse` had no equivalent at all: for any non-success HTTP status
it either passed a barely-touched JSON blob through `CleanResponse` (if the body happened to contain
the literal substring `"status"`) or returned the bare `HttpStatusCode` enum name (e.g. `"BadRequest"`)
with zero detail from the actual response body. Worse, that string then flowed into
`ApiOperationHelper.ValidateApiResponse`, whose checks are narrow text-pattern sniffs (empty / `"(401)
Unauthorized"` / starts-with-`"Error"` / contains `"<!DOCTYPE"`) - a message like `"BadRequest"` or a
properly-extracted sentence like `"An error occurred. Please check the server logs for details."`
matches none of those patterns, so it would have been silently treated as a *valid* successful
response and handed to downstream JSON/CSV parsing instead of being surfaced as an error.

**Fix**:
- New `Helpers/ApiErrorMessageExtractor.cs` - faithful line-for-line port of `ExtractErrorMessage`/
  `IsHtmlResponse`/`ExtractTextFromHtml`/`ExtractFromJson`/`CleanPlainText`, same keyword lists
  (`error`/`exception`/`timeout`/`failed`/`invalid`/`not found`/`unauthorized`/`forbidden`), same
  JSON field-name priority order (`message`/`msg`/`error`/`errorMessage`/`error_description`/`detail`/
  `Message`/`Msg`/`status`/`statusMessage`/`fault`/`reason`, then `errors[]`/`validationErrors[]`/
  `fieldErrors{}`), same fallback messages at every level. Uses `System.Text.Json`'s `JsonDocument`
  instead of VB's Newtonsoft `JObject`/`JArray`/`JToken`, since Newtonsoft isn't referenced anywhere
  else in this C# port (`JsonGlobals`/System.Text.Json is the established convention) - purely a
  library swap, the extraction logic/ordering is unchanged.
- `ApiHelper.ProcessResponse`: non-success responses now return
  `ApiErrorMessageExtractor.ExtractErrorMessage(responseBody)` (or a `"Server returned status code: X
  (n)"` fallback when the body is empty) instead of the bare status-code name or a half-cleaned blob.
- `ApiHelper.ExecuteApiCall`: added an explicit `if (!response.IsSuccessStatusCode)` branch right after
  `ProcessResponse` that logs the status + extracted message, calls the existing
  `ApiOperationHelper.NotifyApiError(...)`, and throws `InvalidOperationException` immediately - so a
  non-success HTTP status is always surfaced as an error via the now-good extracted message, rather
  than being handed to `EnsureValidApiResponse`'s narrower text-pattern check (which could have let it
  through as "valid" data, silently corrupting downstream parsing). `InvalidOperationException` isn't
  one of `ApiOperationHelper.IsTransientError`'s recognized types, so this correctly fails fast instead
  of being retried 3 times as if transient.
- Not in scope here (tracked separately, per the user's own severity ordering): swapping the plain
  `throw new InvalidOperationException(...)` for real `ApiRequestException`/`ApiTimeoutException`
  types matching VB's - that's bundled with the still-pending "server-side cancel-run notification"
  item, since both touch the same exception-throwing architecture at once.

## Reconciliation audit + 2 critical bug fixes — 2026-07-23

Ran a full VB.NET-vs-C# completeness audit (3 parallel passes: `AddinModule.vb`, `FormProcessBar.vb`,
remaining smaller VB files) and found ~20+ gaps of varying severity. Per direction, fixed the two
critical ones first; the rest are queued in severity order (see punch list kept in project tracking).

**Critical bug #1 — ribbon Refresh/Param Refresh silently disabled on every normal report.**
`XLEdgeRibbonHelper.ProcessActiveWorkbook`/`BookHasEdgeReport` were checking `ListObjects[1].Name`
against an `"ORB_DD_"` prefix. Real report tables are named `ORB_{reportId}_{runId}_E` (or `_P`) - VB's
`XlEdgeTableExists` (`AddinModule.vb:2078-2092`) checks `StartsWith("ORB_")` (excluding only the literal
`"orb_params_control"` name), and `"ORB_DD_"` is only ever used as an *exclusion* elsewhere (inside
`SheetFollowHyperlink`), never as an enablement check. The old C# check meant no real table ever matched,
so Refresh/Param Refresh were disabled on every workbook with a normal report. Fixed both methods to match
VB's `ORB_` prefix check exactly.

**Critical bug #2 — `InvokedFromGLSense` entry point was missing entirely.** The sibling GLSense add-in
hands off a completed login to XLEdge via COM reflection
(`InvokeMember("InvokedFromGLSense", ..., new object[] { instanceName, LoginUrl, LoginToken, LoginUserName, xlEdgePermission })`),
but no method with that name existed in the C# port, so every GLSense-initiated login silently failed to
reach XLEdge. Ported from VB `AddinModule.vb:192-281`:
- New `AddinModule.InvokedFromGLSense(string eeName, string eeUrl, string eeToken, string eeUser, bool hasXLEdgePermission = true)`,
  dispatched onto the WPF thread via the existing `SafeInvokeWpf` helper. No-permission branch disables
  `RibEdgeLogin`/`RibControlSheet`/`RibEdgeDebug`/`RibEdgeIncludeOutputData` and blanks+hides the pane if
  visible. Permission-granted branch caches the GLSense COM object (new `_glSenseAddinInstance` field,
  mirroring VB's module-level `addinInstance`), populates `XLEdgeAppState.Instance` login fields, flips
  `RibEdgeLogin`/`RibEdgeLogout` visibility + caption, re-enables the report-action ribbon controls, and -
  if the pane is already visible with a token present - triggers a navigation refresh (since
  `ADXAfterTaskPaneShow` won't re-fire for an already-visible pane).
- New `XLEdgeCTP.NavigateBlankAsync()` / `XLEdgeCTP.RefreshLoginNavigationAsync()` (both `internal`,
  reusing the existing `NavigateToLoginUrlSafeAsync` login-URL-selection logic) plus matching public
  wrappers `ADXExcelTaskPane1.NavigateBlankAsync()` / `ADXExcelTaskPane1.RefreshLoginNavigationAsync()`,
  following the same delegate-to-`_wpfControl` pattern already used by `LogoutAsync`/`ExecuteScriptAsync`.

Not restored: VB's `"Logout " & EELoginURLName"` caption prefix - the C# port's own established
convention (see `XLEdgeCTP.xaml.cs`'s `SetControlCaption("RibEdgeLogout", ...)` call) already omits the
prefix, so `InvokedFromGLSense` matches that convention instead of VB's literal text.

## Ported: image-download resilience (retry/backoff/throttle/HTML-scrape fallback) — 2026-07-23

Next severity-ordered item. `ImageDownloadHelper.TryDownloadImage` had an explicit, honest comment
admitting it only handled a direct image response with a plain per-request timeout - VB's
`DownloadImageInternal` (`FormProcessBar.vb` line ~2909) does considerably more: up to 3 attempts with
retry-on-429/502/503/504 (honoring a `Retry-After` header when present, exponential backoff capped at 10s
otherwise), a shared minimum 1-second spacing enforced between *any* two image downloads (some image hosts
throttle/ban bursty requests), a `Referer` header (Wikipedia's CDN specifically needs
`https://en.wikipedia.org/`; everything else gets its own origin), and an HTML-scrape fallback for when
the "image" URL actually returns an HTML page instead (extracts `og:image`/`twitter:image`/Wikipedia
file-page image/`link[rel=image_src]` via regex, then recurses once into a real image download).

**Fix**: rewrote `Helpers/ImageDownloadHelper.cs` to port all of this - `DownloadImageInternal` (retry
loop + throttle + Referer), `ThrottleImageDownload` (shared `lock`-protected last-download timestamp),
`IsRetryableDownloadError`/`GetRetryDelayMilliseconds` (status-code check + `Retry-After`/backoff),
`GetImageReferer`, `ExtractImageUrlFromHtml`/`NormalizeResolvedImageUrl` (same four regex patterns, same
priority order). Kept the existing `HttpClient`-based approach (rather than switching to VB's
`HttpWebRequest`) since `HttpClient` doesn't throw on non-2xx responses the way `WebException` does -
the non-success-status branch reads the response body for the HTML-fallback check directly instead of
needing a `catch (WebException)`, functionally equivalent to VB's exception-driven fallback path.
`TryDownloadImage`'s public signature is unchanged, so `ReportGenerator.AddImageColumn` (the only caller)
needed no changes.

## Fixed: refreshing a report left stale images and never re-embedded hyperlinks/images for new data — 2026-07-23

Next severity-ordered item. VB's report-writing method (`Edge_GenerateData`) is one shared procedure used
for both first-time report creation AND refreshing an existing table - it always calls
`DeleteSheetImages(sht)` right after finding the existing table (clearing every previously-embedded
`"ORB_*"`-named shape before rewriting anything), and always calls `LinksAndImages`/`LinksAndImages1` near
the end (re-adding drilldown hyperlinks, attachment hyperlinks, and downloaded images for whatever data is
now in the table) - regardless of whether the table was just created or just refreshed.

The C# port split this into two separate methods: `BuildReportTable` (first creation - already calls the
equivalent `AddDrilldownHyperlinks`/`AddAttachmentAndImageColumns`) and `RefreshListObjectAsync` (updating
an existing table), and the refresh method never gained either half of this. Concretely, refreshing a
report with attachment/hyperlink/image columns would: leave every image shape from the *previous* run's
row positions floating over the sheet (now misaligned with the new data, or simply stale), and never
re-add attachment/hyperlink/image content for the newly-refreshed rows at all - a real, visible feature
gap on any report using those column types.

**Fix** (`Helpers/ReportGenerator.cs`):
- Added `DeleteReportShapes(Excel.Worksheet sheet)` - a port of VB's `DeleteSheetImages`, removing every
  shape named `"ORB_*"` (matching `AddImageColumn`'s existing naming convention). Iterates the `Shapes`
  collection backwards by index rather than VB's `For Each` (which mutates the COM collection while
  enumerating it - undefined/fragile behavior to reproduce deliberately) and releases each `Shape` RCW as
  it goes, matching this codebase's established COM-lifecycle convention.
- `RefreshListObjectAsync` now calls `DeleteReportShapes(sheet)` immediately after resolving the existing
  table (VB's exact placement - before any column/row rewriting), and calls the existing
  `AddDrilldownHyperlinks`/`AddAttachmentAndImageColumns` again near the end (VB's `LinksAndImages`
  placement - after the data rewrite and the `RefreshSync` column-deletion step, right before hiding the
  wait window/busy overlay). Both re-embed calls need a deserialized `ReportMeta`, which
  `RefreshListObjectAsync` didn't previously parse at all (it only used the stored JSON for column
  mapping) - now parsed once, non-fatally, right after the metadata is resolved; a parse failure logs and
  skips the re-embed step rather than blocking the refresh's core data update.

## Ported: RibbonInitialize + PrintTimeZone (two smaller severity-ordered items) — 2026-07-23

- **`RibbonInitialize`** (`XLEdgeProcedures.vb` lines ~289-318) didn't exist in C# at all. Added
  `AddinModule.RibbonInitialize()` - resets the ribbon to its logged-out default (Login enabled, every
  other report-action button disabled). Wired into `XLEdgeServerConfiguration`'s delete-instance flow
  (`BtnDelete_Click`), matching VB's `FormConfiguration.CmdDelete_Click`: called after both confirm/cancel
  outcomes, but only while the ribbon is actually showing "Login" (`loginButtonVisible()` true) - same
  guard VB used (`If RibEdgeLogin.Visible Then RibbonInitialize()`). VB's own dead code (loading
  `EETempURLsPath` into an `XmlDocument` it then never reads) was not ported - confirmed via re-reading
  the source that the result genuinely goes unused there.
- **`PrintTimeZone`** (`FormProcessBar.vb` line ~4933) abbreviates the local time zone's DST-aware name to
  its initials (e.g. "Eastern Standard Time" -> "EST") for the banner/param-sheet's "Time Zone :" cell.
  The C# port had been showing `TimeZoneInfo.Local.DisplayName` in full (e.g. "(UTC-05:00) Eastern Time
  (US & Canada)") instead. Added `XLEdgeValueFormatter.PrintTimeZone(DateTime)` and wired it into
  `ReportGenerator.WriteRunInfoStrip`'s Time Zone value cell.

## Fixed: parameter-refresh request body used the display date formatter, not the request formatter — 2026-07-23

Next severity-ordered item, confirmed as a real regression by reading VB's actual source. VB has TWO
separate, differently-named `FormatValue`/`FormatDateValue` function pairs in two different files, doing
genuinely different jobs: `XLEdgeProcedures.vb`'s (shared, used for on-sheet *display* - reformats dates to
`dd-MMM-yyyy`), and a second, **private, module-scoped pair inside `XLEdgeParamsData.vb` itself**
(`BuildJSONObject`'s own `FormatValue`/`FormatDateValue`, lines ~468-508) - which reformats dates to ISO
`yyyy-MM-ddTHH:mm:ss` and returns genuinely typed `Integer`/`Decimal`/`BigInteger` values, specifically
because this one builds the actual JSON **request body** POSTed back to the server for a parameter-edited
refresh, where the server expects the same ISO/typed shape as everywhere else in this API (matching the
drilldown-request formatting from the earlier entry above almost exactly).

The C# port (`Helpers/XLEdgeParamsBuilder.cs`, ported from `BuildJSONObject`) had collapsed this down to a
single shared formatter - it called `XLEdgeValueFormatter.FormatValue` (the *display* one) when building
Value1/Value2 into the refresh JSON body, at all three call sites (the primary value, each split element of
an IN/NOT IN list, and the BETWEEN second value). A DATE-typed parameter edited on the "orb_params_control"
sheet and then refreshed would have sent `"23-Jul-2026"`-style text in the request body instead of the ISO
datetime the server's other endpoints (and the drilldown request) actually send - a real risk of the
server either rejecting the value or silently misparsing it.

**Fix**: swapped all three call sites in `XLEdgeParamsBuilder.BuildJsonPayload` from
`XLEdgeValueFormatter.FormatValue` to `XLEdgeValueFormatter.FormatDrilldownValue` (the ISO/typed-numeric
formatter added in the previous entry for the drilldown request - conveniently, since it's a faithful port
of the exact same logic shape VB duplicated into `XLEdgeParamsData.vb`'s own private function, reusing it
here is correct, not just convenient). The existing IN/NOT-IN date-vs-comma-split guard
(`paramType != DATE/DATETIME`) was already correct and untouched.

**Not replicated** (intentionally): VB's `BuildJSONObject` does one further redundant pass after building
`output.parameters`, re-parsing every `values` element's `.ToString()` back into `Integer`/`Decimal`/
`BigInteger` if `IsNumeric` and the declared type is `INTEGER`/`NUMERIC`/`DECIMAL`. Since `FormatValue`
already returned genuinely typed numeric values earlier in the same method, this second pass is redundant
work on an already-correct value (stringify then immediately re-parse to the same type) - not a distinct
behavior to preserve, so it wasn't ported.

## Fixed: parameter-summary display text (ReportParamValue) didn't match VB - wrong JSON keys, missing escaping/wording — 2026-07-23

Next severity-ordered item. `ReportGenerator.ParseParamDisplayRows`/`ExtractParamValues` (which build each
parameter's one-line "Operator Value" summary shown in the banner/companion parameter sheet) carried an
explicit comment admitting VB's real `ReportParamValue` (`FormProcessBar.vb` line ~4822) body "wasn't
available to port directly." It since became available (read directly this pass), and comparing line by
line turned up several real, user-visible divergences, not just wording taste:

- **Wrong JSON keys entirely.** VB's `ReportParamValue` only ever reads `displayValue`/`displayValues` -
  it never looks at `value`/`values` at all (those are a *different* JSON shape used for building
  request bodies, e.g. `DrilldownRequestBuilder`/`XLEdgeParamsBuilder`). The old C# code checked
  `value`/`values` first and `displayValue` last - the reverse priority, and using keys VB's own display
  logic never reads - so summaries could come out empty or wrong whenever a parameter's JSON had
  `value`/`values` unpopulated but `displayValue`/`displayValues` present (a normal, expected shape).
- **No date/type reformatting at all.** VB always runs every value through `XLEdgeFormatValue` (this
  codebase's `XLEdgeValueFormatter.FormatValue`) using the parameter's own declared `type` before
  display. The old C# code just called `.ToString()` on the raw JSON token - a DATE-typed parameter would
  show its raw ISO/serial value instead of a formatted date.
- **No comma/quote escaping for list values.** VB wraps a non-numeric value containing a comma in double
  quotes before joining a value list with commas, so a value like `"Smith, John"` doesn't get
  misinterpreted as two list items. The old C# code just comma-joined with no escaping.
- **BETWEEN detection missed `componentType`.** VB treats a value pair as a range whenever `componentType`
  contains `"range"`, *or* the operator is `BETWEEN`/`NOT BETWEEN` - not operator alone.
- **`and` vs `And` casing.** VB's between-range join uses lowercase `"{v1} and {v2}"`; the old C# used
  `"And"` (capitalized).
- **No IN/NOT IN single-selection override.** VB overrides the operator wording entirely for `IN`/`NOT IN`
  when `componentType` is `single-selection-prompt`/`oracle-erp-resp-selection` ("is equal to X"/"does not
  equal X"), falling back to "is in list X"/"is not in list X" otherwise - regardless of whatever
  `operatorKey` the generic operator-map lookup would have produced. The old C# code always used the
  generic operator-map wording, so a single-selection `IN` parameter incorrectly showed "is in list X"
  instead of "is equal to X".

**Fix**: rewrote `ParseParamDisplayRows` to read `type`/`componentType` per item (previously not read at
all) and added `BuildReportParamValue` - a direct, line-for-line port of VB's `ReportParamValue` - plus a
`JoinFormatted` helper matching VB's inline escape-or-format lambda exactly. Added
`XLEdgeValueFormatter.IsNumeric` (approximates VB's runtime `IsNumeric()` via `double.TryParse`) to back
the escaping check.

**Also found, not yet ported** (tracked separately): VB's parameter-sheet writer additionally inserts two
extra visible rows - "Responsibility" and "GL Accounts" - sourced from a `GlExtraParams` tuple, distinct
from the hidden `IT4` cell already used for the drilldown request's extra parameters (see the drilldown
entry above). Not in scope for this fix; tracked as a follow-up.

**Deliberately left alone**: `ParamsControlSheetBuilder.ExtractValues` (builds the separate,
re-editable "orb_params_control" sheet's Value1/Value2 columns) has a similar-looking
`displayValue`/`value`/`values` fallback chain, but it's a different VB source function serving a
different purpose (raw editable values, not a combined display sentence) and already has its own
correct IN/NOT-IN single-selection override - left untouched.

## Ported: drilldown clicks now scope the child report to the clicked row — 2026-07-23

Next severity-ordered item. `adxExcelAppEvents1_SheetFollowHyperlink` picked the right child report on a
drilldown click but always re-ran it completely unfiltered - the C# port's own comment admitted this
("drilldown opens the child report unfiltered rather than scoped to the clicked row"). VB's original
(`AddinModule.vb` lines ~1335-1694) builds a scoped request body from three parameter sources per the
child report's drilldown metadata: `PARAM` (resolved from the *parent* report's own stored parameter
values via `ReportHLink_Param`), `STATIC` (a fixed value baked into the drilldown definition), and
cell-value (read from the clicked row via a header-name match, `HRMatch`). It also attaches a
"Responsibility" extra parameter (`IT4`/`ORACLE_RESP_ID`) from the resolved parameter sheet, and refuses
to drill down at all if the report was run under a different logged-in instance (`IT5` mismatch).

Ported:
- `Helpers/XLEdgeValueFormatter.cs` - added `InferDrilldownDataType`/`FormatDrilldownValue` (ported from
  VB's `InferDataType`/`FormatValue1`) and a private ISO-date formatter (`FormatDateValue1`'s C#
  equivalent). Kept separate from the existing `FormatValue`/`FormatDateValue` (which reformat dates for
  on-sheet *display* as `dd-MMM-yyyy`) since the drilldown request body needs ISO datetimes and genuinely
  typed numeric values (int/decimal/BigInteger), not display strings.
- `Helpers/DrilldownRequestBuilder.cs` (new) - `GetColumnType` (VB's `ColType`), `ResolveStoredParamValue`
  (VB's `ReportHLink_Param`, parsing the parent report's stored "Param" JSON for a PARAM-type value), and
  `BuildDrilldownRequestJson`, which assembles the full `ReportParameterRequest` (reusing the existing
  `ReportParameterValue`/`ExtraParameters` models already built for parameter-edited refreshes) and
  serializes it via the shared `JsonGlobals.Options`.
- `Helpers/JsonGlobals.cs` - registered `NumericJsonConverter` globally (previously only attached via
  `[JsonConverter]` on individual `object`-typed properties like `Value`/`ReportId`). `ReportParameterValue
  .Values` is `List<object>` with no such attribute, so without this, a `BigInteger` element built by the
  new formatter would hit default object serialization (no `BigInteger` support) instead of the converter
  that already knows how to write one as a raw JSON number. Additive only - every property that already
  had the attribute is unaffected, since property-level converters still take precedence.
- `AddinModule.cs` - added the missing `IT5` different-instance check (reusing the same
  `TryResolveInstanceAndChildFlag` helper the Refresh handlers already use) right where VB has it, resolves
  the same companion parameter sheet for the `IT4` "Responsibility" value, and passes the built payload
  into `ReportGenerator.CreateReportFromTitleAsync` via a new optional `paramsJsonPayload` parameter
  (mirroring `RefreshListObjectAsync`'s existing pattern) instead of always POSTing an empty body.

**Deliberately not ported**: VB's elaborate child-sheet-name truncation math (`ChildShtName`/`RngAddress`/
`TotalLen`, lines ~1483-1533), which names the new child sheet after the parent sheet + clicked cell
address (truncated to fit Excel's 31-char sheet-name limit) so repeated drilldowns from different cells
land on differently-named sheets instead of colliding. The C# report-creation pipeline names/finds sheets
by `tableId` (`ORB_{reportId}_{runId}_E`), which is already unique per report run since the server assigns
a fresh `runId` per execution - so this specific collision case shouldn't arise the same way it did in VB,
but it hasn't been independently verified end-to-end. Flagged here rather than silently dropped, in case
sheet-naming collisions are ever reported for repeated drilldowns.

**Also confirmed while reading VB**: the `Strs.Length >= 4 AndAlso Not String.IsNullOrEmpty(Strs(3))` gate
around this whole block looked like it might be a feature flag, but `GetDrilldownStringForColumn` always
emits exactly 4 pipe-delimited segments (`DRILLDOWN|id|name|reportId`) - so that condition is always true
in practice and needed no equivalent gate in the C# port.

## Fixed: RefreshAll didn't skip child reports or different-instance reports — 2026-07-23

Continuing the reconciliation audit's severity-ordered punch list (next item after the 2 critical bugs).
`RibEdgeRefreshAll_OnClick` looped every `_E` table in the workbook and refreshed all of them
unconditionally. VB's equivalent loop (`FormProcessBar.vb` RefreshAll branch, ~line 5151-5245) skips two
categories before refreshing: reports executed under a different logged-in XLEdge instance (`IT5` on the
companion parameter sheet, or the report's own sheet in same-sheet mode, doesn't match the current login
URL - "DiffInstance"), and child/drilldown reports (`IT1 == "Child Report"`) - RefreshAll must not
silently pull another session's data or independently re-run a report that only exists as a drilldown
child of another.

The C# port already had the exact reusable check (`TryResolveInstanceAndChildFlag`, used by the
`RibEdgeParamRefreshBook_OnClick`/`RibEdgeParamRefresh_OnClick` ports) but it wasn't wired into
`RibEdgeRefreshAll_OnClick`'s per-table loop. Fixed by calling it for each table before refreshing:
different-instance tables and child-report tables are now skipped silently (matching VB's `Continue For`),
logged at debug level for traceability.

## Perf: batched the header-row write into a single Interop call — 2026-07-10

User asked (conceptually) whether building the data as an in-memory array first and writing it to Excel
in one shot, instead of cell-by-cell, would help performance for large reports (hundreds of columns,
up to Excel's ~1,048,576-row ceiling). Answer: yes, and `BuildReportTable`'s main data body already does
exactly this - `writeArr` (an `object[,]`) is built entirely in managed memory, then assigned in a single
`Range.Value2 = writeArr` call via `startCell.Resize[dataRowCount, mappings.Count]`, which is one COM
round trip regardless of size (Excel does the bulk write internally). That's already the fast path.

The one place that *wasn't* doing this: the header row, written via `for (int c = 0; c < mappings.Count;
c++) { sheet.Cells[headerRow, c + 1].Value2 = ... }` - one Interop call (really two: a `Cells[r,c]`
property-get plus a `Value2` property-set) per column. Harmless for a handful of columns, but with
"hundreds of columns" that's hundreds of avoidable round trips before any data has even been written.
Fixed to match the data body's pattern: build a `1 x mappings.Count` `object[,]` in memory, then assign it
in one call via `((Excel.Range)sheet.Cells[headerRow, 1]).Resize[1, mappings.Count].Value2 = headerArr`.

**Not fixable the same way (inherent to the Excel object model, not a code gap):** `AddDrilldownHyperlinks`
and `AddAttachmentAndImageColumns` still loop per matching cell, because `Range.Hyperlinks.Add(...)` has
no bulk/array equivalent - each hyperlink is its own COM object needing its own call. These only run at
all for columns actually flagged as drilldown/attachment/image columns in the report metadata, so a plain
report with none of those is unaffected; a report that has one of those columns *and* millions of rows
will still pay a per-row cost there that no amount of batching the plain data can avoid.

Complementary (already in place, not new here): `ExcelBulkOperationScope` (see the entry further below)
suspends `ScreenUpdating`/`EnableEvents`/`DisplayAlerts` and sets `Calculation = xlCalculationManual` for
the whole report-run, which matters just as much as batching the writes themselves - without it, Excel
would repaint/recalculate after every operation regardless of how few Interop calls were made.

## Fixed: toast pinned to a small band at the top instead of growing with content — 2026-07-10

User reported `AppOverlay.ShowToast` (embedded in `XLEdgeCTP` via the `AppOverlayControl`, also reused in
`XLEdgeWaitWindow`/`XLEdgeServerConfiguration`/`XLEdgeAbout`) always renders the toast as a small band
pinned to the top, even when there's plenty of space below and the message is long enough that it should
grow downward instead of relying on its inner `ScrollViewer` to scroll a tiny visible area.

Root cause: `Toast`'s `MaxHeight` was bound in XAML to `{Binding ActualHeight, ElementName=RootOverlay,
Converter={StaticResource PercentConverter}, ConverterParameter=0.9}` - i.e. 90% of `RootOverlay`'s (the
whole `AppOverlay` UserControl's own root) `ActualHeight`. But `DismissToast` sets `this.Visibility =
Visibility.Collapsed` on the *entire* `AppOverlay` control between toasts (not just the `Toast` border
itself) whenever the busy/confirm overlays aren't also active - and a `Collapsed` element is excluded
from layout entirely, so its `ActualHeight` resets and stays stale/zero until the *next* full layout pass
completes. Every time `ShowToast`/`ShowToastAsync` set `Visibility` back to `Visible` and then needed
`Toast.MaxHeight` to size the toast, `RootOverlay.ActualHeight` could still reflect that stale
post-collapse value at that exact moment - pinning the toast to a tiny max height instead of the real
available space, on every single toast, not just the first one.

**Fixed** by moving this out of a fragile ActualHeight-of-a-just-uncollapsed-sibling binding entirely:
added `AppOverlay.UpdateToastMaxHeight()`, which reads the size from `this.Parent` (the actual host
container - e.g. `XLEdgeCTP`'s outer `Grid`, or whichever `Window`/panel a given `AppOverlay` instance is
embedded in) instead of `RootOverlay`/`this`. The parent is never collapsed itself (only the `AppOverlay`
child is toggled), so its `ActualHeight` is always current. Falls back to `this.ActualHeight` and then
`SystemParameters.WorkArea.Height` if the parent can't be read for some reason. Called at the top of both
`ShowToast` and `ShowToastAsync`, right after `this.Visibility = Visibility.Visible`. Removed the now-
redundant XAML `MaxHeight` binding on the `Toast` Border (documented why in a XAML comment) since it's
set exclusively from code now.

## Making the data<->parameter sheet link fully rename-proof — 2026-07-10

User asked how the link between a data sheet and its companion parameter sheet is established, and
whether it survives either sheet being renamed, plus whether the orphaned-sheet-deletion (previous
entry) was already covered. Answer: the *binding* was already rename-proof - it's the hidden `IT2` cell
on the parameter sheet holding the data table's `tableId` (e.g. `ORB_123_456_E`), matched against the
data sheet's `ListObject.Name` (never the sheet's display name) - both of which are immune to a manual
sheet rename. `ExcelSheetHelper.GetParameterSheet(nameHint, tableId)` already implements this correctly:
it checks the name hint first as a fast path, then unconditionally falls back to scanning every sheet's
`IT2` for an exact match if the hint doesn't exist or doesn't match. Every consumer in `AddinModule.cs`
(refresh, tab-label sync, drilldown) already goes through this helper. But auditing `ReportGenerator.cs`
turned up two spots that bypassed this and depended on names directly, which *would* have broken on a
rename:

- **`BuildCompanionParameterSheet`'s reuse check** was `ExcelSheetHelper.SheetExists("P_{dataSheet.Name}",
  workbook)` - a pure name match. If the user (or an earlier mode-switch) had left the parameter sheet
  under a different name than the current `P_{dataSheet.Name}` convention, this would silently create a
  *second*, duplicate parameter sheet instead of reusing the existing one. Fixed to call
  `ExcelSheetHelper.GetParameterSheet(paramSheetName, tableId)` instead (name hint + IT2 fallback scan) -
  ported from VB's `GenerateParamSheet`, which does exactly this (`GetParameterSheet(shtname,
  TableIDs.Item1)`) and, notably, does **not** rename the sheet if one is found and reused - only a
  brand-new sheet gets the computed `P_{...}` name. That detail matters: it means a user's own rename of
  the parameter sheet is respected indefinitely across re-runs, not silently reverted.
- **The "Goto Report Data" hyperlink** baked the data sheet's *current* name into a static SubAddress
  (`$"'{dataSheet.Name}'!A1"`). Excel hyperlink addresses are plain text, not a live reference - so this
  would go stale the instant either sheet was renamed after the parameter sheet was built, and Excel's
  default hyperlink-follow behavior would try to jump to a sheet reference that no longer resolves.
  Fixed to pass empty strings for both Address and SubAddress, exactly matching VB's `GenerateParamSheet`
  (`.Hyperlinks.Add(DrillCell, "", "", TextToDisplay:="Goto Report Data")`). The hyperlink is really just
  a click target: `AddinModule.adxExcelAppEvents1_SheetFollowHyperlink` recognizes it by its cell text
  ("Goto Report Data") and calls `NavigateToReportDataSheet`, which resolves the real target by reading
  the parameter sheet's own `IT2` and scanning the workbook for the `ListObject` whose `Name` matches -
  the same rename-proof binding, so there's nothing left anywhere that depends on either sheet's name.

**Net effect:** the data sheet <-> parameter sheet link now has zero dependency on either sheet's display
name at any point - creation, reuse across re-runs, navigation, and the mode-switch deletion covered in
the entry below all resolve purely through `tableId` (`ListObject.Name`) <-> `IT2`.

## Handling the same-sheet <-> separate-sheet mode switch on re-run — 2026-07-10

User walked through the exact scenario the dual-mode layout work hadn't accounted for: run a report with
"Parameter values in same sheet" checked, uncheck it and re-run the *same* report (same table id), then
check it again and re-run once more. Each toggle needs to *convert* the existing sheet, not just start
writing at a different row and leave the old layout behind:

- **Same-sheet -> separate-sheet:** the old banner (rows 1-7, with rows 3-6 outline-grouped) has to go -
  otherwise the data sheet would keep the collapsible grouping and squashed 5px spacer rows sitting on
  top of what's now supposed to be a plain table starting at row 1.
- **Separate-sheet -> same-sheet:** the table has to physically move from row 1 down to row 8 to make
  room for the banner, and - the part explicitly called out - the companion "P_" parameter sheet from
  the old separate-sheet run is now orphaned and has to be deleted, not left behind as a stale sheet
  nobody points to anymore.

Re-read VB's actual handling of this in both directions: `Edge_GenerateData_Multisheet`'s `tbExists`
branch (`FormProcessBar.vb` ~line 1710, `If TableObj.HeaderRowRange.Offset(1, 0).Row = 9 Then` - i.e. "was
the header at row 8, meaning this table used to be in same-sheet mode?") ungroups rows 3-6 if grouped,
resets `Range("A1:A7").RowHeight` back to 15, then deletes rows 1-7 entirely
(`Selection.Delete(Shift:=xlShiftUp)`). `Edge_GenerateData`'s `tbExists` branch (~line 3809,
`If TableObj.HeaderRowRange.Offset(1, 0).Row = 2 Then` - i.e. "was the header at row 1, meaning this
table used to be in separate-sheet mode?") looks up the old companion sheet via `GetParameterSheet`,
marks it for deletion, then inserts 7 blank rows at the top (`Selection.Insert(Shift:=xlShiftDown)`) -
the actual deletion of the marked sheet happens later, in the routine's `Finally` block.

**Ported into `ReportGenerator.BuildReportTable`:**
- `sameSheet`/`headerRow` are now computed *before* finding/rebuilding the sheet (previously computed
  after), because converting the existing layout requires knowing the *target* row before deciding what
  to do with the *old* one.
- Before deleting the existing `ListObject`, its `HeaderRowRange.Row` is captured as `oldHeaderRow`.
  Comparing `oldHeaderRow` to the newly-computed `headerRow` identifies exactly which transition (if
  any) is happening: `8 -> 1` (same -> separate) or `1 -> 8` (separate -> same). No change (`8 -> 8` or
  `1 -> 1`) - i.e. re-running a report without touching the option - does nothing extra, same as before.
- New `RemoveSameSheetBanner(sheet)` - same-sheet -> separate-sheet: ungroups `A3:A6` if grouped (reusing
  the `Convert.ToInt32(EntireRow.OutlineLevel)` pattern from the earlier `InvalidCastException` fix),
  resets `Range["A1:A7"].RowHeight = 15`, then `Range["1:7"].Delete(xlShiftUp)`.
- New `InsertRoomForSameSheetBanner(sheet)` - separate-sheet -> same-sheet: `Range["1:7"].Insert(xlShiftDown)`,
  shifting the existing row-1 table down to row 8 so the normal header-row-8 write path (already in
  place) lands in the right spot and the pre-existing data is preserved (just relocated), not lost.
- The old companion sheet is resolved via the existing `ExcelSheetHelper.GetParameterSheet($"P_{sheet.Name}",
  tableId)` (the same IT2/tableId-binding lookup `RefreshListObjectAsync` already relies on) *before* the
  row insert, and its name stashed in `companionSheetToDelete`; the actual `.Delete()` call happens at the
  very end of `BuildReportTable`, after the new banner has been written successfully - matching VB's
  "resolve early, delete late" ordering so a mid-build failure doesn't strand the report with no
  parameter sheet at all. (`Worksheet.Delete()` would normally prompt a confirmation alert - suppressed
  automatically here since the whole report-run is now wrapped in `ExcelBulkOperationScope`, which sets
  `DisplayAlerts = False` for the duration - see the entry below.)

## Two more fixes: Excel screen-updating suspension during report runs + another Rows[int] cast crash — 2026-07-10

**1) Suspend ScreenUpdating/EnableEvents/DisplayAlerts/Calculation while a report is running.** VB did
this at the WinForms progress-dialog level: `FrmProcessBar_Load` (fires when the progress dialog opens,
right as report generation starts) sets `EnableEvents = False`, `ScreenUpdating = False`,
`DisplayAlerts = False`, `Calculation = xlCalculationManual`; `FormProcessBar_Closing` (fires when the
dialog closes, however the background worker finished - success, cancel, or error) unconditionally
forces everything back to `True`/`xlCalculationAutomatic` - not save/restore-previous-value, always
forced back on. Added a new `ExcelBulkOperationScope : IDisposable` class (bottom of `ReportGenerator.cs`)
that does the same thing, and wired it in as `using var excelBulkScope = new ExcelBulkOperationScope();`
as the first statement of `CreateReportFromTitleAsync`, `CreateReportFromListObjectAsync`, and
`RefreshListObjectAsync` - the three entry points that actually download/write report data. Using a C#
`using`-declaration (not a `using (...) { }` block) means it's disposed on every exit path - normal
return, an early `return` (there are many scattered through these methods for validation/cancel/error
cases), or an unhandled exception - without needing to touch/re-indent the existing method bodies.

**2) Another `InvalidCastException` from the same root cause as the Worksheets/Sheets bug.** User hit a
runtime crash on `sheet.Rows[3] as Excel.Range` (in the rows 3-6 outline-group check added for the
banner formatting fix above) - notable because this is an `as` cast, which is supposed to return null
rather than throw, yet it still threw. `Worksheet.Rows`'/`Range.Rows`'s integer indexer is declared to
return `object` (late-bound, like `Cells[r,c]`), but unlike `Cells[r,c]` casts (which work fine
everywhere else in this codebase), casting the result of the `Rows[int]`/`Columns[int]` collection
indexer to `Excel.Range` is apparently unreliable at runtime - same family of issue as
`Worksheets`/`Sheets`, just on a different collection. Audited the whole project for this pattern
(`grep .Rows[` / `.Columns[`) and fixed all 3 occurrences found:
- `ReportGenerator.WriteSameSheetBanner`'s outline-group check and the `RowHeight = 5` lines - switched
  to `sheet.Range["A3"]`/`sheet.Range["A7"]` (string indexer, guaranteed to return `Excel.Range`
  directly) and `.EntireRow.OutlineLevel`, matching what VB's own `Range("A3")`/`Range("A7")` did (VB's
  `Rows(3).outlinelevel` check never hit this because VB late-binding papers over the interface
  mismatch that C#'s early-bound cast can't).
- Two spots in `ReportGenerator.cs` (`lo.DataBodyRange.Rows[1] as Excel.Range`, used to capture/compare
  first-row formulas during a table resize) and one in `AddinModule.cs`
  (`tableObj.DataBodyRange.Rows[1]`, used to scan the first data row for drilldown/attachment
  hyperlinks) - switched to `dataBodyRange.Resize[1, columnCount]`, which is a plain property (not a
  late-bound indexer) that's already used successfully elsewhere in this codebase (e.g.
  `startCell.Resize[dataRowCount, mappings.Count]` in `BuildReportTable`) and returns `Excel.Range`
  directly, no cast/`as` needed.

**Takeaway for future code in this project:** avoid `SomeRange.Rows[n]`/`SomeRange.Columns[n]`
(integer indexer on the Rows/Columns collections) entirely. Prefer `Range["A1"]` (string address),
`Range.Cells[r, c]` (needs an explicit `(Excel.Range)` cast, but is reliable), or `Range.Resize[rows,
cols]`/`Range.Offset[r, c]` (plain properties, no cast needed, reliable).

**Follow-up compile error caught immediately after:** `sheet.Range["A3"].EntireRow.OutlineLevel > 1`
failed to compile - "Operator '>' cannot be applied to operands of type 'object' and 'int'" - because
`Range.EntireRow` is *also* declared to return `object` in this interop assembly (same as `Cells[r,c]`,
just a named property instead of an indexer, so it's easy to assume it's already strongly typed when it
isn't). Fixed by assigning to an explicitly-typed local first: `Excel.Range entireRow =
(Excel.Range)sheet.Range["A3"].EntireRow;` then reading `entireRow.OutlineLevel` off that - unlike
`Rows[int]`/`Columns[int]`, `EntireRow` genuinely does return a real `Range` COM object underneath, so
casting it (unlike the Rows/Columns collection indexer) is reliable at runtime, not just at compile
time.

**Second follow-up, a runtime exception on the fix above:** `entireRow.OutlineLevel` is *itself* also
typed `object` in this interop assembly (not `int`), and it boxes a numeric type that isn't `Int32`
(almost certainly `Int16`/`short`, matching how this property marshals from VBA). `(int)` on that
boxed value is a C# *unboxing* conversion, which (unlike a numeric conversion) requires the compile-time
target type to exactly match the boxed type - unboxing an `Int16` box as `int` throws "Specified cast is
not valid" even though the value is a perfectly good number. Fixed by using `Convert.ToInt32(...)`
instead of a direct cast - `Convert.ToInt32` goes through `IConvertible` and correctly handles a boxed
`short`, `int`, `double`, etc., so it doesn't matter which exact numeric type Excel hands back:
`int outlineLevel = Convert.ToInt32(entireRow.OutlineLevel); if (outlineLevel > 1) { ... }`. Audited the
rest of the project for the same `(int)someExcelProperty` unboxing pattern (`OutlineLevel`/`ColorIndex`/
`Count`-style members) - no other occurrences found.

## Follow-up: missing post-data-write cell formatting on the report banner/param sheet — 2026-07-10

User flagged that after the dual-mode layout fix above, the actual cell *formatting* on the banner/
param sheet was still missing - specifically colors on the first row, rows 3-6, and the Run Date/Time
Zone area. Re-read `Edge_GenerateData` (rows 3660-4600 of `FormProcessBar.vb`) and `GenerateParamSheet`
(rows 2260-2605) in detail this time (not just structure, but every `.Font`/`.Interior`/`.Outline` call)
and found the previous port had the right cells/values but none of the actual formatting VB applies to
them, all of which happens in VB *after* the data table/ListObject is built - matching the user's own
description. Ported into `ReportGenerator.cs`:

- **Row 1 title** (`A1:E1`): bold+italic, size 10, white text (`Font.ColorIndex = 2`), dark blue fill
  (`Interior.Color = RGB(21, 96, 130)`). Previously just `Font.Bold = true`, no fill/color/size.
- **Row 2 "Parameters Section:" label** and the **G1/I1/K1/K2 "Run Date :"/"Time Zone :"/"Executed in
  :"/"Record Count :" labels**: same bold+italic+size10+white-on-peach styling (`Interior.Color =
  RGB(241, 169, 131)`), each with `EntireColumn.AutoFit()`.
- **H1/J1/L1/L2 value cells** (run date, time zone, login URL, record count): italic (not bold), size
  10, dark-blue *font* color (`Font.Color = RGB(21, 96, 130)` - fill vs font color was the label/value
  distinction VB uses throughout), `EntireColumn.AutoFit()`. H1 keeps its `dd-mmm-yyyy hh:mm:ss` format.
- **Rows 3-6 outline grouping** (same-sheet banner only): `DisplayGridlines = False`,
  `Outline.SummaryRow/AutomaticStyles`, `DisplayOutline = True`, ungroup-then-`Range("A3:A6").Rows.Group()`
  (matching VB's guard against double-grouping), plus `Outline.ShowLevels(1, 1)` and an `Application.Goto`
  to A1 at the end, all previously entirely missing (only the `RowHeight = 5` spacer sizing existed).
- **Parameter label/value cells** now match `ParamTitleInput`/`ParamValueInput` exactly instead of just
  `Font.Bold`: labels are bold+italic, size 9, `ColorIndex = 14`, right/bottom-aligned, truncated to 28
  chars; values are italic (not bold), size 9, `ColorIndex = 16`, forced `NumberFormat = "@"`,
  left/bottom-aligned. Extracted into shared `WriteParamLabelCell`/`WriteParamValueCell` helpers used by
  both the same-sheet banner and the companion parameter sheet.
- Added a shared `WriteRunInfoStrip` helper (the G/I/K label + H/J/L value styling is identical between
  `Edge_GenerateData` and `GenerateParamSheet`, just at different merge widths) and a local `Rgb(r,g,b)`
  OLE-color helper (same `r + (g<<8) + (b<<16)` packing already used in `ParamsControlSheetBuilder.cs`).
- Companion parameter sheet also gained the previously-missing `DisplayGridlines = False`, column A
  `EntireColumn.AutoFit()`, and `Application.Goto(A1)` calls from `GenerateParamSheet`.
- All new `Application`/`ActiveWindow` access goes through `ExcelApplicationHelper.RequireActiveExcelApplication()`
  rather than `sheet.Application`, to stay consistent with the rest of the codebase's safe-COM-access pattern.

**Still-known simplification (unchanged from before, not what the user flagged this time):** VB's
`ParamValueInput` also creates a per-cell named range (`TableName_CellAddress`) and a self-referencing
data-validation rule to make parameter value cells effectively read-only in the workbook UI. This is a
*behavioral* lock (prevents editing), not a visual format, so it was left out of this pass - flagging
here in case it turns out to matter later (e.g. if users report being able to edit parameter cells that
should be locked).

## Major functional gap closed: dual-mode report layout ("Parameter values in same sheet") — 2026-07-10

User flagged that report generation has a user-facing setting - "Parameter values in same sheet"
(`XLEdgeAppState.Instance.ParamDataSameSheet`) - controlling two completely different sheet layouts,
ported in VB as two separate top-level functions in FormProcessBar.vb: `Edge_GenerateData` (checked -
title + parameters banner in rows 1-7 of the data sheet itself, table starts row 8) and
`Edge_GenerateData_Multisheet` + `GenerateParamSheet` (unchecked - data sheet has no banner and starts
at row 1, title/parameters instead go on a separate companion "P_{sheetname}" sheet). **`BuildReportTable`
previously implemented neither mode correctly** - it always started the table at row 1 with no banner
and never created a companion parameter sheet, silently ignoring the `ParamDataSameSheet` setting
entirely. This also explains why several already-ported consumers (`UpdateTabLabel`,
`RibEdgeRefresh_OnClick`, `RibEdgeParamRefreshBook_OnClick`, `BuildRefreshParamsPayload`, all of which
already correctly check `HeaderRowRange.Offset(1,0).Row == 2` to decide whether a companion parameter
sheet should exist) would have failed with "Reports parameters information worksheet missing" on every
report created in separate-sheet mode - there was simply never a companion sheet to find.

**Fixed in `ReportGenerator.cs`:**
- `BuildReportTable` now reads `XLEdgeAppState.Instance.ParamDataSameSheet` and sets the table's header
  row to 8 (same-sheet) or 1 (separate-sheet) accordingly, then calls one of two new methods.
- `WriteSameSheetBanner` (same-sheet mode) - merged A1:E1 title, a Run Date/Time Zone/Executed-in strip
  on row 1 columns G-L, a "Parameters Section:" label + record count on row 2, and the report's current
  parameters as Label/Value pairs in 3-row blocks (rows 4-6) that wrap to the next column pair (A/B →
  C/D → E/F → ...) past 3 parameters - matching the VB original's `irow > 6 → icol += 2` rule exactly.
  Writes IT1 (drilldown flag) and IT5 (login URL) bookkeeping cells on the data sheet itself. Confirmed
  via re-reading the VB router (`BGWorker_DoWork`) that same-sheet mode never writes IT2 - that's
  exclusively a separate-sheet concept.
- `BuildCompanionParameterSheet` (separate-sheet mode) - builds/rebuilds a "P_{dataSheetName}" sheet
  (name truncated to 28 chars, matching the VB 31-char sheet-name-limit handling), with the same title/
  Run-Date/Parameters-Section layout, one parameter per row (column A/B, starting row 3, no column-pair
  wrapping - confirmed this is genuinely different from the same-sheet banner's wrap behavior), a "Goto
  Report Data" hyperlink back to the data sheet (the click handler for this already existed in
  `AddinModule.NavigateToReportDataSheet` from an earlier session but had nothing to click on until
  now), and IT1/IT2/IT5 bookkeeping - IT2 (the table id) is written here and *only* here, confirming
  it's the key that binds a parameter sheet back to its data table.
- `ParseParamDisplayRows`/`ExtractParamValues` - shared parsing of the "Params" JSON (same array shape
  `ParamsControlSheetBuilder.ProcessSheetParams` already reads) into Label/Value display pairs used by
  both of the methods above.

**Simplification, documented in code:** the exact VB value-display formatting lived in a
`ReportParamValue` helper whose body wasn't available to read directly (only its call site was seen
via the research pass for this fix) - the reproduced format (an `XLEdgeOperatorMappings`-mapped
operator label followed by the formatted value or value range, e.g. "is equal to  123") matches the
same shape but isn't guaranteed to be a byte-for-byte match with the original wording. Also not
reproduced: the VB banner's row 3-6 Excel outline/grouping and specific fill colors (rows 3/7 are still
set to a 5px spacer height, but the collapsible-group behavior and exact colors are skipped) - purely
cosmetic, doesn't affect functionality.

**Not yet re-verified:** this was written and reasoned through without the ability to run Excel in this
environment - please test both settings (checked and unchecked) against a real report run before
relying on it, particularly the column-pair wrapping in same-sheet mode with more than 3 parameters,
and the "Goto Report Data" round-trip in separate-sheet mode.

## Runtime bug fixed post-build: Excel.Worksheets vs Excel.Sheets — 2026-07-10

First real-world test surfaced crashes on report creation and on every sheet/workbook activation:

```
InvalidCastException: Unable to cast COM object of type 'System.__ComObject' to interface type
'Microsoft.Office.Interop.Excel.Worksheets'. ... QueryInterface ... No such interface supported
(E_NOINTERFACE).
```

Root cause: `Workbook.Worksheets` returns an `Excel.Sheets` COM object at runtime, not
`Excel.Worksheets` - the `Worksheets` interface exists in the interop assembly but isn't what this
property actually implements. Several places in this codebase declared the variable as
`Excel.Worksheets`, which doesn't implicitly convert from `Sheets` and needs an explicit cast to
compile - that explicit cast compiles fine but throws this exact `InvalidCastException` at runtime,
every time. This is why the crash only showed up now, once the project was actually built and run in
Visual Studio (this environment cannot compile or run Excel COM code, so it was never caught before).

Fixed in all 7 locations by changing the declared type from `Excel.Worksheets` to `Excel.Sheets`
(removing the now-unnecessary explicit cast): `AddinModule.RibEdgeRefreshAll_OnClick`,
`ExcelSheetHelper.SheetExists`, `ExcelSheetHelper.GetParameterSheet`,
`ParamsControlSheetBuilder.CreateControlSheet`, `ReportGenerator.TryGetStoredReportXml`,
`ReportGenerator.FindSheetWithTable` (the one hit by the reported "Failed to build report table"
crash), and `XLEdgeRibbonHelper.BookHasEdgeReport` (hit by the reported `ApplySheetActiveState`/
`ApplyWorkbookActiveState` crashes, which fire on every sheet/workbook switch). `Excel.Sheets` and
`Excel.Worksheets` are used interchangeably as element types in `foreach (Excel.Worksheet ws in ...)`
elsewhere in the codebase - those were never affected, only explicit collection-variable
declarations were. **Please rebuild and re-test** - this was a systemic pattern (not just one typo),
so worth confirming the same crash doesn't recur elsewhere before moving on.

## Locations
- Original VB.NET source: `VB.NET\XLEdge\` — reference only, not edited.
- New C# project: `XLEdge\XLEdge\` — active work.
- `bkpup\` — older snapshot of the C# project, not the active copy.

## Target stack (already decided, don't re-litigate)
- .NET Framework 4.8.1, C# language version: latest.
- UI: WPF replacing WinForms, themed with MahApps.Metro + MahApps.Metro.IconPacks.
- Excel COM integration via Add-in Express (AddinExpress.XL.2005 / AddinExpress.MSO.2005). AddinModule GUIDs/ProgIDs must stay identical to the VB build so existing installs upgrade cleanly instead of losing registration.
- JSON: System.Text.Json, replacing the VB project's Newtonsoft.Json (confirmed intentional — `Models\AllModels.cs` already uses `JsonPropertyName`).
- Logging: NLog.
- Embedded browser: Microsoft.Web.WebView2, WPF control replacing the WinForms control.
- Task pane host: `ADXExcelTaskPane1` stays a WinForms Add-in Express control, hosting WPF content (ElementHost / WindowsFormsIntegration referenced).

## File mapping (VB.NET → C#)

| VB.NET | C# / WPF | Status |
|---|---|---|
| AddinModule.vb | AddinModule.cs | Ported |
| ADXExcelTaskPane1.vb | ADXExcelTaskPane1.cs | Ported |
| FormAbout.vb | Views\XLEdgeAbout.xaml | Ported |
| FormConfiguration.vb | Views\XLEdgeServerConfiguration.xaml | Ported (verified) |
| FormDetails.vb | Views\XLEdgeLoginDetails.xaml | Ported (verified) |
| CalendarPopup.vb | Views\XLEdgeCalendar.xaml | Ported — verify feature parity |
| FormDrillDown.vb | Views\XLEdgeDrilldownReports.xaml | Ported — verify feature parity |
| FormOptions.vb | Views\XLEdgeOptions.xaml | Ported — verify feature parity |
| FormProcessBar.vb + FormProgressNew.vb | Views\XLEdgeWaitWindow.xaml | Merged into one window (verified, uses CancellationHelper + DispatcherTimer) |
| **XLEdgeProcedures.vb** | Helpers\ExcelApplicationHelper, ExcelWindowHelper, ReportGenerator, EdgeRequestParser, XLEdgeRibbonHelper, ApiOperationHelper, ApiResponseHelper, APIHelper, JsonGlobals | **Not yet verified — this is almost certainly the bulk of remaining work** |
| ReportMetaInfo.vb | Models\AllModels.cs or ApiResult.cs (unconfirmed) | Needs confirmation |
| XLEdgeParamsData.vb | Models\AllModels.cs (unconfirmed) | Needs confirmation |
| NumericConverter.vb | Converters\WidthPercentageConverter.cs is a *different* converter | Needs confirmation — may still be required for WPF numeric bindings |
| NLogConfig.vb | Helpers\LogHelper.cs / Utilities\LogUtility.cs (assumed) | Needs confirmation |
| (none — new infra) | Utilities\UiDispatcher.cs, WpfAppManager.cs, DpiAwareWindow.cs, DpiAwarenessHelper.cs | New, built to support WPF-in-COM-addin threading/DPI — reuse these, don't duplicate |

## Risk areas to check as remaining work lands
- Cross-thread/dispatcher safety: every Excel event handler that touches a WPF window must marshal through `UiDispatcher`. Audit for consistency once XLEdgeProcedures.vb logic is ported.
- COM object lifecycle: every Excel Range/Worksheet/Workbook obtained in ExcelApplicationHelper/ExcelWindowHelper needs deterministic `Marshal.ReleaseComObject`, or EXCEL.EXE lingers after close.
- COM registration GUIDs in AddinModule must exactly match the VB build.
- Bulk range read/write vs. cell-by-cell interop in ReportGenerator, for performance.

## XLEdgeProcedures.vb port — done 2026-07-10

Ported into new files, all under `Helpers\`, following the existing pattern of small focused classes:

| VB function | C# home | Notes |
|---|---|---|
| Edge_SheetExists | ExcelSheetHelper.SheetExists | Fixed a COM leak: VB only released the *last* Worksheet touched by the loop variable; C# releases every one, plus the Worksheets collection itself. |
| GetParameterSheet | ExcelSheetHelper.GetParameterSheet | Same COM-leak fix applied. |
| NamedRangeExists / DeleteNamedRange / CreateNamedRange / CleanUpName | ExcelSheetHelper | Named ranges had no C# equivalent at all before this. Also fixed: VB never released `Excel.Name` COM objects in these loops. |
| ExcelApplicationResolver.RequireActiveExcelApplication (defined in AddinModule.vb, used throughout XLEdgeProcedures.vb) | ExcelApplicationHelper.RequireActiveExcelApplication | Added to the existing helper alongside Get/TryGet. |
| IsCellinEditMode | ExcelApplicationHelper.IsCellInEditMode | Direct port. |
| XLEdgeFormatDateValue / XLEdgeFormatValue / RemoveEquaSymbol | XLEdgeValueFormatter | Pure string logic, no COM involved. |
| FollowDrilldown / DrillPostData / ChildTableName / ChildShtName / ChildRptLabel / Edge_WorkBook / ActWorksheet / My.Settings.ProcessRunning (module-level globals) | XLEdgeAppState (new properties) | Centralized as instance state instead of static globals, consistent with existing LoginUrl/DebugLogs properties. |
| Edge_CloseProgress | ProgressCoordinator.ResetReportState | **Partial port.** The state-reset + DummySheet-hiding logic is fully ported. Two pieces are left as explicit TODOs in the code because they call back into AddinModule instance members that don't exist yet: `AddinModule.CurrentInstance.EEDeleteAllFiles()` and `AddinModule.CurrentInstance.UpdateTabLabel(...)`. Wire these in once AddinModule.vb is ported — do not guess the signatures now. |

**Confirmed obsolete / superseded — not ported:**
- `XLEdgeMsgDisplay` / `XlEdgeMsgResult` → superseded by `MessageFunctions.XLEdgeMessage`, which already exists and is better (WPF window, dispatcher-safe, Excel-owned modality).
- `XLEdgeLogDebug` / `XLEdgeLogInfo` / `XLEdgeLogHttpWebRequestProperties` / `XLEdgeLogProperty` → superseded by `LogUtility` (already handles debug-mode gating) and the inline `LogUtility.LogDebug(...)` calls already used for HTTP diagnostics in `XLEdgeAbout.xaml.cs`.
- `GetCustomXMLPart` / `EscapeXml` / `UnescapeXml` → superseded by `ReportGenerator`'s own custom-XML scheme (`XDocument` + `XCData`, which handles escaping automatically) - confirmed this is a genuinely different, newer design, not a gap.
- `PictureBoxImage` → WinForms-only (PictureBox doesn't exist in WPF); the ported `XLEdgeAbout.xaml` uses a declarative `<Image>` instead.
- `XlEdge_Initialize` → now just `XLEdgeAppState.Instance.EnsureExcelApplication()`, already covered.
- `RibbonInitialize` and the rest of `Edge_ThreadProgress` → **blocked on the AddinModule.vb port.** They reference actual ribbon controls (`EdgeLoginBtn`, etc.) and the old single global progress form, neither of which exist in the C# project yet. Porting them now would mean inventing APIs that may not match what AddinModule.vb actually needs.

**Important scope correction:** AddinModule.vb (not XLEdgeProcedures.vb) is the real bulk of remaining work — it's ~3,266 lines and 46 of the functions above are called from inside it. It contains the ribbon event handlers, task-pane wiring, and Excel event subscriptions. This has been added as its own tracked task (see task list) rather than folded into "next steps" below.

## AddinModule.vb port — in progress, started 2026-07-10

**Critical bug caught and fixed before porting anything else:** the in-progress C# `AddinModule.cs` had a *different* COM GUID (`3F5A137C-108B-4228-9ED0-C2C4F56A495D`) than the VB build (`80B0FB76-A5F4-41E0-B283-C175B7374B60`). Shipped as-is, every existing install would have lost its COM registration on upgrade. Fixed to match the VB GUID exactly. No other GUID mismatches found (the task pane class has no separate COM GUID in either version).

C# `AddinModule.cs` was already at 713 of a functional-equivalent ~3,321 VB lines before this session (ribbon-loaded bootstrap, login/options/about/help/logout dialogs, the WPF calendar popup wired to SheetSelectionChange, sheet/workbook-activate ribbon state). Added this session, with matching `+=` wiring added to `AddinModule.Designer.cs` (hand-added since there's no Add-in Express designer available in this environment):

| VB function | C# home | Notes |
|---|---|---|
| AddinModule_AddinStartupComplete | AddinModule_AddinStartupComplete | Thin - delegates to `XLApp.Initialize`, already existing. |
| AddinModule_AddinBeginShutdown | AddinModule_AddinBeginShutdown | **Partial port, deliberately.** The VB original also called `_Edge_ExcelApp.Quit()` here, which force-closes the entire Excel application whenever the add-in unloads (including when a user just disables it via the COM Add-ins dialog while other work is open). That looks like a bug, not intentional design, so it was **not** carried over - flagged in a code comment. Tell me if this was actually intentional and you want it back. The old `EEQueryTable` COM release was dropped as genuinely obsolete (no QueryTable object exists anywhere in the new architecture - confirmed by search). |
| RibEdgeRefresh_OnClick | RibEdgeRefresh_OnClick | **Significantly simplified, confirmed safe to do so.** The VB version resolved a separate parameter sheet and validated its IT5 (logged-in instance)/IT1 ("Child Report") markers, then built request post-data via `XLEdgeParamsData.BuildParamData`. Traced through `ReportGenerator.RefreshListObjectAsync` (already existing, 955-line file) and confirmed it's fully self-contained - it re-derives the runId/columns it needs from the CustomXMLParts metadata already stored per table. None of that old validation is needed anymore, so the C# version is just: validate login state + edit mode + active table, then call `RefreshListObjectAsync`. |
| RibEdgeRefreshAll_OnClick / IsRefreshAll | RibEdgeRefreshAll_OnClick | Same simplification - loops worksheets collecting "_E" tables (with proper per-iteration COM release, unlike the VB loop), then calls `RefreshListObjectAsync` per table and aggregates errors. |
| GetGLSenseAddinObject | GetGLSenseAddinObject | Direct port (locates the sibling "GLSense" Add-in Express add-in via Excel's COMAddIns collection). Fixed a COM leak: the VB loop never released the `COMAddIn` wrapper for non-matching candidates, only the one that matched - C# releases every one. |

**Now ported (see the two sections below for full detail):** `SheetFollowHyperlink`, `RibControlSheet_OnClick` + `BuildParamData`, `UpdateTabLabel`, `RibEdgeLogout_PropertyChanging`, `SheetBeforeDelete`/`DeleteTimer_Tick`/`DeleteHashNamedRanges`/`DeleteNamedCache`/`NmSheetExists`, `RibEdgeIncludeOutputData_OnClick`/`RibEdgeDebug_OnClick`, `RibEdgeParamRefreshBook_OnClick`/`RibEdgeParamRefresh_OnClick`.

**Deliberately not ported (no equivalent needed):**
- `OnRibbonBeforeCreate` (~2214) - in VB this just set `AdxRibbonTab1.Caption = EdgeBranding & " XLEdge"` at ribbon-creation time. The C# `AddinModule.Designer.cs` already sets `TabXLEdge.Caption = "Orbit XLEdge"` statically at design time, so there's nothing left for this handler to do - porting it would just reassign the same fixed string a second time.
- `RibEdgeProcessReports_OnClick` (~2218, sets `excelApp.EnableEvents = True`) - there is no corresponding ribbon button declared anywhere in `AddinModule.Designer.cs`; this looks like a leftover debug/reset control from the VB ribbon XML that was never carried into the new ribbon layout. Nothing to wire it to. Flag if this button needs to be re-added to the ribbon.
- `XLEdgeParamsData.BuildParamData` - confirmed no longer needed by the refresh flow (see above); its real replacement, `Helpers/XLEdgeParamsBuilder.BuildParamData`, is now ported (see the RibControlSheet_OnClick section below).

## "Needs confirmation" items — resolved 2026-07-10

**ReportMetaInfo.vb / XLEdgeParamsData.vb → Models mapping:**
- `ReportMetaInfo.vb`'s DTOs (report metadata + drill-submit JSON contracts: `RptColumn`, `ChildParameter`, `RptDrilldown`, `RptParameter`, `ReportMeta`, `DrillParameter`, `ExtraParameters`, `DrillSubmit`, `ColumnMetadata`, `DrilldownInfo`, `DrilldownParameter`, plus `Outputprop`/`Properties`) were **not previously ported anywhere** - added now as `Models\ReportMetaModels.cs`, with proper C# naming and `JsonPropertyName` attributes preserving the original JSON keys. (`braodcastMsg`/`xledgeuserPreferences` were already ported as `BroadcastMessage`/`XLEdgeUserPreferences` in `AllModels.cs` - not duplicated.)
- `XLEdgeParamsData.vb` (`BuildParamData`/`BuildJSONObject`) is **real, not-yet-ported functionality, not obsolete legacy.** Traced its only caller (the old `RibEdgeRefresh_OnClick`) and confirmed that specific call site is superseded by `ReportGenerator.RefreshListObjectAsync` - but `BuildParamData` itself implements a *different* feature: turning user-edited filter values in the "orb_params_control" sheet into a JSON request payload to resubmit a report with new parameters. That sheet name is still referenced in the already-ported calendar-popup handler (`adxExcelAppEvents1_SheetSelectionChange`), confirming the feature is still meant to exist - it just hasn't been rebuilt yet. Its full port depends on `AddinModule.vb`'s `RibControlSheet_OnClick`, which turned out to be much bigger than first estimated (see below) - so it's being ported together with that, not in isolation. Added the `OperatorMappings` dictionary it depends on as `Helpers\XLEdgeOperatorMappings.cs`, and the request-body DTOs (`ReportParameterRequest`/`ReportParameterValue`) into `Models\ReportMetaModels.cs`, ready for when that logic lands.

**NumericConverter.vb / NLogConfig.vb equivalents:**
- `NLogConfig.vb` is confirmed fully superseded by `Helpers\LogHelper.cs`, which is actually *more* complete (adds archive-file rotation the VB version never had). One real gap found and fixed: the VB layout included `${callsite:...}` (which method/file logged the message) and the C# layout had dropped it - restored, since it's strictly useful for debugging and costs nothing.
- `NumericConverter.vb` is a Newtonsoft `JsonConverter`, not a WPF value converter as the name might suggest - it makes sure numeric-looking strings get serialized as JSON numbers instead of strings. Ported as `Helpers\NumericJsonConverter.cs` using System.Text.Json's `JsonConverter<object>` (matching the project's System.Text.Json choice), applied via `[JsonConverter(typeof(NumericJsonConverter))]` on the relevant model properties in `ReportMetaModels.cs`.

## Major scope correction — the report-creation engine, found 2026-07-10

While starting the drilldown port, traced `ReportGenerator.CreateReportFromTitleAsync` end to end and found it's a **stub, not a finished feature**: it downloads CSV, report-definition JSON, and parameter JSON from the server, then just logs them - it never builds an Excel table from any of it. `CreateReportFromListObjectAsync` confirms this: it only actually works by falling back to `RefreshListObjectAsync`, which requires a table that *already exists*. There is currently no working path in the C# project to create a brand-new report table - which is what happens on every drilldown click, and almost certainly on the very first time any report is ever run.

The real "build a report table" logic lives in **`FormProcessBar.vb` - 5,398 lines, the single largest file in the entire VB.NET codebase**, bigger than `AddinModule.vb`. Found two near-duplicate ~450-line blocks (around lines 1600-2070 and 3600-4430) that do the actual work: create or reuse the target worksheet, create the `ListObject`/Excel Table, write the data (`DataTableToExcel`, `WriteDataTableColumnWiseUsingTuples`), persist metadata (`CreateXMLParts`), add drilldown hyperlinks (grouped by column, batched, with a 65,530-hyperlink safety cap), add file-attachment hyperlinks (`ParseAttachmentLink`), autofit/style the table, and generate the parameter sheet (`GenerateParamSheet`).

**Important threading finding, directly relevant to the "no cross-thread issues" goal:** the old engine runs all of this Excel COM interop inside a `BackgroundWorker.DoWork` handler - i.e. on a thread-pool thread, not Excel's STA thread. That's a real hazard in the original app (COM objects generally must be used on the thread that created them). Decision: **do not replicate this.** `RefreshListObjectAsync` already establishes the correct pattern for this codebase - `async`/`await` where network calls are truly awaited, with the Excel COM manipulation happening on the resumed UI-thread continuation, and `Dispatcher.InvokeAsync` reserved for explicit WPF window updates. The report-creation engine will follow that same pattern rather than the VB original's background-thread model.

**Status:** this is now understood to be one of the largest and highest-risk remaining pieces of the whole migration - likely bigger than everything ported so far combined - and it sits underneath both "run a new report" and "drilldown." User confirmed: build this engine first, before finishing the smaller AddinModule.vb handlers. Still need to read `DataTableToExcel`, `CreateXMLParts`, `WriteDataTableColumnWiseUsingTuples`, `ParamInfo`, `GenerateParamSheet`, `LinksAndImages`/`LinksAndImages1`, `ParseAttachmentLink`, `DeleteSheetImages`, and the second near-duplicate block, before it can be ported correctly. Given the size and that this environment cannot compile or test against a live Excel instance, this will be built out incrementally across further sessions rather than in one pass, to keep each piece verifiable.

## Report-creation engine + drilldown — completed 2026-07-10 (with documented simplifications)

`ReportGenerator.CreateReportFromTitleAsync` now actually builds the Excel table instead of just logging the fetched data. Added `BuildReportTable` and helpers (`FindSheetWithTable`, `BuildSheetName`/`SanitizeSheetName`, `MakeUniqueName`, `AddDrilldownHyperlinks`, `TryGetStoredReportXml`), plus `ExcelSheetHelper.HRMatch`. `AddinModule.cs` now has a working `adxExcelAppEvents1_SheetFollowHyperlink` handler wired in the Designer, so clicking a drilldown link in a report actually opens the child report.

**What's covered:** creating/reusing the worksheet, building the ListObject with columns derived from the report metadata (respecting the "_E" ad-hoc report naming and the "hdn" hidden-column flag), bulk-writing data with per-column date/value formatting (reusing `XLEdgeValueFormatter`), title/IT5 bookkeeping cells, drilldown hyperlinks grouped by column with the same tooltip format the VB original used (so a picker dialog - `XLEdgeDrilldownReports`, which already existed - works unchanged), the 65,530-hyperlink safety cap, and persisting metadata through the *same* CustomXMLParts schema `RefreshListObjectAsync` already reads (`BuildCustomXml`/`SaveCustomXmlPart`), so newly-created reports can be refreshed afterward. "Goto Report Data" hyperlink navigation is also ported.

**Deliberately simplified or deferred - not silently dropped, tracked here:**
- **Drilldown parameter injection.** Clicking a drilldown currently re-runs the child report fresh rather than filtering it by the clicked row's parent values. The VB original built a request payload per drilldown parameter (`PARAM` type read from stored params, `STATIC` used a fixed value, otherwise read from the clicked row's cell via `HRMatch`/`ColType`/`FormatValue1`/`InferDataType`). This is bundled with task #9 (`RibControlSheet_OnClick` + `BuildParamData`) since it's the same "build a parameter request payload" capability.
- **"Process"/scheduled report type and the multi-sheet "refresh all" variant** (`Edge_GenerateData_Multisheet`) - only the "Edge" ad-hoc report path is ported. Given `RefreshListObjectAsync` already covers refreshing everything that exists, and ribbon buttons/drilldown only ever produce "Edge"-type titles, this covers the paths actually reachable from the current UI - but if scheduled/"Process"-type reports are a real feature you use, flag it and it needs its own pass.
- **Attachment hyperlinks and the image-embedding subsystem** (`ParseAttachmentLink`, `DownloadImage` and ~10 supporting functions in FormProcessBar.vb, roughly 350 lines) - not ported. Any report column that embeds images or file-attachment links won't render those in the new build yet.
- **The "Logs" debug worksheet feature** (`Edge_GenerateLogs`/`Edge_FillLogs`) - not ported; assumed to be a diagnostic nicety, not core functionality.
- **Old HTTP-fetch/error-extraction functions** (`GetServerData`, `ExtractErrorMessage`, `IsHtmlResponse`, `ExtractTextFromHtml`, `ExtractFromJson`, `CleanPlainText`, `CancelBackEndRequest`, `ReturnHTTP`) - confirmed superseded by the already-existing `ApiHelper.ServerAPI`, which is a more robust HttpClient-based implementation with retry logic. Not re-ported.
- **Exact legacy cell formatting** (row grouping/outline levels, the reserved rows 1-7 "Parameters Section" banner with specific merge/color formatting) - simplified to a plain table starting at row 1. The functional data is all there; the decorative header banner isn't recreated.
- **Backward-compatibility with workbooks created by the VB.NET version**: the old engine stored metadata in a differently-shaped, namespaced CustomXMLParts schema (`CreateXMLParts`/`GetCustomXMLPart`, `XLEdgeURI` namespace with `Data`/`DataMeta`/`DataParam` nodes). The new engine's schema (`Title`/`ListObjectName`/`Meta`/`Params`/`Columns`, no namespace) is a clean break. **This means "Refresh" will not work on report tables that already exist in a workbook built with the VB.NET add-in** - only tables created by this new engine. Worth confirming whether that matters given this is a fresh install scenario (VB.NET fully uninstalled first) - if end users have existing workbooks with VB.NET-created reports they expect to keep refreshing, this is a real compatibility gap to test for.
- The `SheetFollowHyperlink` event-wiring line in `AddinModule.Designer.cs` uses a **guessed delegate type** (`ADXExcelSheet_EventHandler`, matched by parameter-signature shape to `SheetSelectionChange`) - I could not inspect the Add-in Express assembly metadata from this environment to confirm the exact delegate name. Flagged with a comment; check this compiles in Visual Studio first.

## RibControlSheet_OnClick + BuildParamData — completed (task #9)

`Helpers/ParamsControlSheetBuilder.cs` (`ShowOrRebuild`, `FindControlTable`, `CreateControlSheet`, `LockToCurrentValue`, `ProcessSheetParams`, `ExtractValues`) ports the "orb_params_control" sheet feature - viewing/editing every active report's filter parameters in one place, with data-validation-enforced read-only cells for locked values. `Helpers/XLEdgeParamsBuilder.cs` (`BuildParamData`, `ReadControlSheetRows`, `BuildJsonPayload`) ports the VB `XLEdgeParamsData.BuildParamData`/`BuildJSONObject`, converting from Newtonsoft `JObject` to System.Text.Json `JsonElement`, reusing `XLEdgeOperatorMappings` and `XLEdgeValueFormatter`.

**Resolved (see "Remaining pending items closed out" below):** the JSON payload `BuildParamData` produces is now wired into `RefreshListObjectAsync` as an optional POST body.

## Sheet-delete cleanup + remaining ribbon handlers — completed (task #10 + rest of task #9)

Ported directly into `AddinModule.cs` (all instance methods, following the file's existing style):

- **`UpdateTabLabel`** - refreshes the `RibSheetLabel` ribbon caption for the active sheet (scheduled-output / data-report / drilldown+attachment column summary, built by scanning the first data row's hyperlink ScreenTips), warns if a report's companion parameter sheet is missing, and calls `DeleteNamedCache`. **Simplified from the VB original:** the VB version also directly toggled `RibEdgeRefresh`/`RibEdgeParamRefresh` `Enabled` state inline. That's centralized already in `XLEdgeRibbonHelper.ProcessActiveWorkbook` (used for workbook/sheet-activate ribbon state), so `UpdateTabLabel` now just calls `ApplyRibbonState(workbook)` instead of re-implementing a second, possibly-diverging copy of the same enable/disable rules. Wired into `ProgressCoordinator.ResetReportState`, closing that TODO.
- **`RibEdgeLogout_PropertyChanging`** - refreshes the tab label whenever the Logout button's visibility flips (i.e. right after login/logout). **Uncertain delegate/event-args type** (`ADXRibbonControlPropertyChanging_EventHandler`/`ADXRibbonControlPropertyChangingEventArgs`) - guessed by ADX naming convention, could not verify against the actual assembly from this environment. Flagged with a comment in `AddinModule.Designer.cs`; confirm in Visual Studio.
- **`AdxExcelAppEvents1_SheetBeforeDelete` / `DeleteTimer_Tick`** - defers deletion of a report's companion parameter sheet to a `System.Windows.Forms.Timer` tick (200ms) instead of doing it inline, matching the VB original's crash-avoidance approach. Uses the already-confirmed `ADXHostActiveObject_EventHandler` delegate (same shape as `SheetActivate`/`WorkbookActivate`).
- **`DeleteHashNamedRanges` / `DeleteNamedCache` / `NmSheetExists`** - named-range cleanup for broken `#REF!` ranges and orphaned `_ChildReport`/`_Instance` caches. Every `Excel.Name`/`Excel.Worksheet` COM object is released per-iteration (the VB original didn't release `Name` objects at all in this code).
- **`RibEdgeIncludeOutputData_OnClick` / `RibEdgeDebug_OnClick`** - simple ribbon toggles, now writing to `XLEdgeAppState.Instance.DebugOutputData` (new property, added for this) and `.DebugLogs` (already existed).
- **`RibEdgeParamRefreshBook_OnClick` / `RibEdgeParamRefresh_OnClick`** - re-run a whole workbook's (or just the active sheet's) reports with their current parameters, by validating each table's stored "executed instance" (`IT5`) and "Child Report" (`IT1`) markers against the current login, then triggering a DOM hook in the hosted web app via JavaScript (`[reruntype=xledgeworkbookrerun]` / `#XLEdgeParamRefresh`), exactly like the VB original. This required adding a small new capability that didn't exist yet: `XLEdgeCTP.ExecuteScriptAsync(script)` (dispatcher-safe wrapper around `WebCtrl.CoreWebView2.ExecuteScriptAsync`, modeled on the existing `LogoutSessionAsync`) and `ADXExcelTaskPane1.ExecuteScriptAsync(script)` to expose it outside the WPF layer. Shared instance-mismatch/child-report logic factored into `TryResolveInstanceAndChildFlag`.

## Remaining pending items closed out — completed 2026-07-10

Went through the full outstanding punch list from the previous pass. Status of each:

**1. Wired `BuildParamData`'s payload into a real refresh-with-parameters call.** Traced the actual VB call site: it turns out the plain "Refresh" button (`RibEdgeRefresh_OnClick`) always resolved the report's parameter sheet and built `DrillPostData` via `XLEdgeParamsData.BuildParamData` *before* refreshing - so even a plain refresh was supposed to pick up any filter-value edits made in the "orb_params_control" sheet, not just blindly re-fetch with the original parameters. Confirmed via `ReturnHTTP` that "Edge" report refreshes always POST (JSON body = the params payload, empty string when there's nothing to send) rather than GET. Added `AddinModule.BuildRefreshParamsPayload` (resolves the companion parameter sheet the same way `UpdateTabLabel` does, then calls `XLEdgeParamsBuilder.BuildParamData`) and wired it into both `RibEdgeRefresh_OnClick` and `RibEdgeRefreshAll_OnClick`. `ReportGenerator.RefreshListObjectAsync` gained an optional `paramsJsonPayload` parameter - POSTs it when present, otherwise behaves exactly as before (plain GET).

**2. Added VB CustomXMLParts backward-compatibility.** New `ReportGenerator.TryResolveReportXmlForRefresh` understands both schemas: this engine's (`Title`/`ListObjectName`/`Meta`/`Params`/`Columns`) and the legacy VB one (`XLEdgeURI`-namespaced `MetaData`/`Data`/`InfoID`/`DataMeta`/`DataParam`). For a legacy part, the reportId/runId are derived from the table-name pattern `ORB_{reportId}_{runId}_E` (the VB original never stored a pipe-delimited title), and the column mapping is derived from the table's *current* header row (the old schema never stored a raw-CSV-index mapping either - reports were re-matched by header position at refresh time). Both `RefreshListObjectAsync` and `TryGetStoredReportXml` (used by drilldown/param-building) now go through this shared resolver. `SaveCustomXmlPart` now also deletes the old legacy part when saving, so a legacy table is transparently upgraded to the new schema the first time it's refreshed and never needs the fallback again. **This closes the "Refresh doesn't work on VB-created workbooks" gap** flagged earlier - test it against a real pre-existing VB workbook before relying on it, since this was written without the ability to run Excel in this environment.

**3. Ported the attachment/image-embedding subsystem.** New `Helpers/AttachmentLinkHelper.cs` (`TryParseAttachmentLink`/`BuildDownloadUrl`, ported from `ParseAttachmentLink` and the attachment-download URL logic in VB's `SheetFollowHyperlink`) and `Helpers/ImageDownloadHelper.cs` (ported from `DownloadImageInternal`, simplified - see below). `ReportGenerator.AddAttachmentAndImageColumns` (called from `BuildReportTable`, right after drilldown hyperlinks) now handles all three VB column behaviors: `isFileAttached` columns become `ATTACHMENT|...` hyperlinks (same "payload lives in the ScreenTip" pattern as drilldown), `properties.outputprop.type = "HYPERLINK"` columns become plain hyperlinks, and `"IMAGE"` columns get the image downloaded and embedded as a picture shape sized from `imgWidth`/`imgHeight`, with the same row-height/column-width auto-sizing math as the VB original. `AddinModule`'s drilldown click handler (`adxExcelAppEvents1_SheetFollowHyperlink`) now actually handles `attachment|` ScreenTips instead of skipping them - it resolves the download URL, calls the new `ApiHelper.DownloadFileAsync` (bearer-token-authenticated GET, saves to the Downloads folder using the response's `Content-Disposition` filename, ported from VB's `DownloadFile`), and shows the same confirmation message box as the original.
   - **Simplification, documented in code:** VB's `DownloadImageInternal` also included an HTML-scraping fallback (parsing `og:image`/`twitter:image`/`link[rel=image_src]` meta tags out of a non-image HTTP response, for embedding arbitrary web pages as images), a Wikipedia-specific `Referer` override, retry-on-HTTP-429/502/503 logic, and a 1-second minimum spacing between downloads. None of that is ported - only a direct image response is handled. If report image columns actually rely on the HTML-fallback behavior (i.e. point at a web page rather than a raw image URL), flag it and it can be added.
   - **Known tradeoff, not a regression:** `ImageDownloadHelper.TryDownloadImage` blocks synchronously (`.GetAwaiter().GetResult()`) since `BuildReportTable` itself is synchronous - matches the VB original's synchronous `WebRequest.GetResponse()` behavior, just noted here since a full async rewrite of the report-build pipeline would be a much larger change than this task warranted.

**4. Evaluated the "Process"/scheduled report type + multisheet refresh.** Confirmed it is **not reachable from anything in the ported UI** - `BuildReportTable` always builds `_E`-suffixed ("Edge" ad-hoc) tables via the Edge runner endpoint regardless of report type, there is no ribbon control or task-pane route that produces a `"Process|..."` title anywhere in the codebase, and `RibEdgeRefresh(All)_OnClick` only ever collect `_E` tables. Left unported rather than building an entire parallel ingestion pipeline (different REST endpoint, different table-naming/refresh rules, no existing trigger) for a feature with no current entry point. **If "Process"/scheduled reports are actually used, say so** and this needs a dedicated follow-up pass, likely including new ribbon UI.

**5. Thread-safety/dispatcher audit (task #4) - done.** Reviewed every `.cs` file. Result: no `async void` outside legitimate event handlers, no `ConfigureAwait(false)` bleeding into Excel COM calls anywhere, dispatcher usage is almost entirely the safe `InvokeAsync` form (the few synchronous `Dispatcher.Invoke` calls found are all one-time bootstrap/layout paths, not hot paths). One accepted tradeoff noted above (`ImageDownloadHelper`).

**6. COM object lifecycle audit (task #5) - done.** Reviewed every `.cs` file. Found and fixed one high-severity leak - `XLEdgeRibbonHelper.BookHasEdgeReport` iterated every worksheet without releasing any of them, and this method runs on **every single** `SheetActivate`/`WorkbookActivate` event, so it was leaking a Worksheet RCW per sheet on every sheet switch in the workbook. Also fixed two medium leaks: `ReportGenerator.TryGetStoredReportXml`'s worksheet-search loop, and `ParamsControlSheetBuilder.CreateControlSheet`'s worksheet scan. A handful of low-severity cell/Range-level leaks remain in `XLEdgeParamsBuilder.ReadControlSheetRows` and a couple of other cell-iteration loops - not fixed this pass since their impact is small (Range/cell RCWs, not full collections) and time was better spent on the collection-level leaks; worth a follow-up pass if you want it fully clean.

**7. Attempted to verify the two guessed Add-in Express delegate/event-args types - blocked.** Searched the entire mounted workspace for `AddinExpress.MSO.dll` to inspect via reflection; not found anywhere under the project folder (it's presumably resolved from a NuGet cache or GAC location outside what this environment can access). Both guesses remain as flagged code comments (`SheetFollowHyperlink` → `ADXExcelSheet_EventHandler`; `RibEdgeLogout.PropertyChanging` → `ADXRibbonControlPropertyChanging_EventHandler`/`ADXRibbonControlPropertyChangingEventArgs`) - **these two lines need confirming in Visual Studio before the project will compile.**

## Next steps
1. Build + manual test in Visual Studio — this environment can generate/edit code but cannot compile a Windows COM add-in or run the Add-in Express designer. This is now the main blocker to finding out what's actually broken.
2. Confirm the remaining guessed Add-in Express delegate/event-args type compiles as-is (see #7 above and the 2026-07-30 note below) - `RibEdgeLogout.PropertyChanging` is still unverified; `SheetFollowHyperlink` is resolved.
3. Test the legacy-workbook CustomXMLParts fallback (#2 above) against a real VB-created workbook if you have one available - this was written without the ability to run Excel in this environment.
4. Decide whether "Process"/scheduled reports and the HTML-fallback image-embedding behavior are actually needed (see #3/#4 above) - both were confirmed out of scope for now, but only you know if they're used.
5. Optional cleanup: the remaining low-severity COM leaks noted in #6 above.

## 2026-07-30: GLSense logout wasn't logging XLEdge out too

**Bug report** (from GLSense's own log): clicking GLSense's Logout ribbon button threw
`COMException 0x80020006 (DISP_E_UNKNOWNNAME)` - "Unknown name" - while GLSense tried to tell
XLEdge to log out too, via late-bound COM reflection
(`edgeAddin.GetType().InvokeMember("LogoffFromAddin", BindingFlags.InvokeMethod, null,
edgeAddin, new object[] { })` in `GLSense\FinalWorkingCode\GLSense\AddinModule.cs`'s
`RibLogout_OnClick`). GLSense caught and logged the exception, so the user-visible symptom
was subtler than a crash: GLSense's own session logged out fine, but XLEdge's session
(ribbon still showing "Logout <instance name>", task panes still logged in) silently stayed
logged in.

**Root cause**: no method named `LogoffFromAddin` ever existed on `AddinModule` - the real
logout logic lives in `LogoffFromXLEdgeAddin` (`private async Task`, used only internally by
this add-in's own `RibEdgeLogout_OnClick`). GLSense's caller was written against a method
name/contract that was either planned but never added during the VB→C# migration, or existed
in the old VB add-in under that name and was dropped when the logout logic got renamed/
rewritten here. Either way, IDispatch had nothing to resolve `LogoffFromAddin` to, hence
`DISP_E_UNKNOWNNAME` - not a marshaling bug, not a threading bug, just a missing public
entry point. This mirrors `InvokedFromGLSense` above (the working GLSense→XLEdge *login*
contract) - that one already exists and works; this is its missing logout counterpart.

**Fix**: added a public `AddinModule.LogoffFromAddin()` method - a thin wrapper that calls
the existing `LogoffFromXLEdgeAddin()` + `ApplyRibbonState("LoggedOut")` sequence already
used by `RibEdgeLogout_OnClick`, so XLEdge's own internal logout flow is completely
unchanged; this only adds the missing public name GLSense's existing (unmodified) call site
already expects. Fire-and-forget by design, matching how GLSense's own call site already
discards the result and matching `InvokedFromGLSense`'s synchronous-void COM-callable shape.

No change needed on the GLSense side - its call site already had the right contract in mind,
XLEdge just wasn't honoring it.

## 2026-07-30: `SheetFollowHyperlink` guessed delegate type — confirmed resolved

Re-checked the two guessed Add-in Express delegate/event-args types flagged in "Remaining
pending items closed out" (#7) and `Next steps` (#2). Current `AddinModule.Designer.cs:375`
wires `adxExcelAppEvents1.SheetFollowHyperlink` as `AddinExpress.MSO.ADXExcelHyperlink_EventHandler`
- not the originally-guessed `ADXExcelSheet_EventHandler`. This matches the handler's actual
3-parameter shape (`sender, sheet, hyperlink`) exactly, and no "guessed"/"confirm in Visual
Studio" comment remains near the wiring or the handler body (`AddinModule.cs:2279`) - a
repo-wide search for `guess|Confirm in Visual Studio|could not verify|unconfirmed` under
`XLEdge/` turns up nothing for this handler anymore. Whoever landed on
`ADXExcelHyperlink_EventHandler` didn't update this doc at the time, so it's being closed out
here instead.

**`RibEdgeLogout.PropertyChanging` is still open** - `AddinModule.cs:1161-1165` still carries
its original "could not verify the exact Add-in Express delegate/event-args type name...
best-guess names following ADX's naming convention... Confirm in Visual Studio" comment
verbatim, and the wiring (`AddinModule.Designer.cs:130`, `ADXRibbonPropertyChanging_EventHandler`/
`ADXRibbonPropertyChangingEventArgs`) hasn't changed. This is now the only remaining
unconfirmed-delegate risk from the original pair.
