# XLEdge AIPowered Stage 1 — WebView2 Pane Spike Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Prove that a WebView2 hosted in a core-created window, re-parented into the shell's ADX task pane, works and survives
hot reload — or find out early that it doesn't and fall back (spec §4).

**Architecture:** The core (`XLEdge.Addin.Core`, its own AppDomain) creates one borderless, off-screen, non-activating top-level
window per Excel window, hosting a WebView2 that shows a spike page, and returns its HWND through the existing
`IXLEdgeAddin.CreateTaskPaneContent`. The shell's `XLEdgeTaskPane` re-parents it (embedding code moved to `Hosting\PaneEmbedding`
so tests use the real code). `CloseTaskPaneContent` / `Shutdown` destroy every window the core made before the domain unloads.

**Tech Stack:** .NET Framework 4.8.1, WinForms + Microsoft.Web.WebView2 1.0.4022.49 (today's XLEdge version and control),
Add-in Express 10.1, MSTest 3.11.1.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§4 pane spike, §5 phase 1, §6 deviations)

## Global Constraints

- No contract change: `CreateTaskPaneContent(IntPtr)`, `CloseTaskPaneContent(IntPtr)`, `OnTaskPaneShown(IntPtr)` exist; `ContractInfo.Version` stays `1`.
- WebView2 package `Microsoft.Web.WebView2` **1.0.4022.49** (same as today's `XLEdge\packages.config`), WinForms control.
- Environment: `userDataFolder = IPathProvider.BrowserDataPath`, `AllowSingleSignOnUsingOSPrimaryAccount = true` (today's options).
- Every core entry point catches and logs; nothing escapes across the AppDomain boundary.
- Core windows: `ShowInTaskbar=false`, `WS_EX_TOOLWINDOW`, never activated on show, off-screen (-32000,-32000) until embedded.
- 0 errors / 0 warnings, Debug and Release; full suite green; local commits only, no push; no Claude attribution in commits.
- No Excel registration (`adxregistrator`) or Excel start without the developer's explicit go-ahead (same ProgId as the installed XLEdge).

## Review Focus

1. A core window (WinForms parking window, marshaling control, form) outliving the unloaded domain — Excel crashes on its next message.
2. Reload while the pane is open: the old content must be gone before unload and the new core's content embedded after load.
3. Two Excel windows: one content window per Excel window; closing one never closes the other's.
4. A second WebView2 environment on the same user data folder after a reload (different options → `0x8007139F`).
5. WebView2 Runtime missing or the browser process crashing: logged, pane stays empty, Excel unaffected.

---

### Task 1: The core package carries WebView2

**Files:** Modify `XLEdge.Addin.Core\XLEdge.Addin.Core.csproj`, create `XLEdge.Addin.Core\packages.config`; Test `XLEdge.Tests\CorePackageTests.cs`

**Interfaces:** Produces — the core build output and zip contain `Microsoft.Web.WebView2.Core.dll`, `Microsoft.Web.WebView2.WinForms.dll`,
`runtimes\win-x86\native\WebView2Loader.dll`, `runtimes\win-x64\native\WebView2Loader.dll`.

- [ ] Step 1: test `BuiltPackage_CarriesWebView2` — zip entries include the four paths above (forward slashes).
- [ ] Step 2: run → FAIL (entries missing).
- [ ] Step 3: add the package (restore into `AIPowered\XLEdge\packages`), references `Core` + `WinForms` (+ `System.Windows.Forms`), the WebView2 `.targets` import.
- [ ] Step 4: build 0/0, suite green; `BuiltPackage_InstallsAndLoads` still passes (bootstrapper shares the WebView2 group).
- [ ] Step 5: commit `XLEdge AIPowered Stage 1: core package carries WebView2 1.0.4022.49`.

### Task 2: Shell embedding code is one tested helper

**Files:** Create `XLEdge\Hosting\PaneEmbedding.cs`; modify `XLEdge\XLEdgeTaskPane.cs`; Test `XLEdge.Tests\PaneEmbeddingTests.cs`

**Interfaces:** Produces `internal static class PaneEmbedding { bool Embed(IntPtr content, IntPtr parent, int width, int height); void Resize(IntPtr content, int width, int height); void Release(IntPtr content); bool IsAlive(IntPtr content); }`.

- [ ] Step 1: tests (STA, a WinForms `Form` as the content, another as the pane):
  `Embed_MakesTheContentAChildSizedToThePane` (WS_CHILD set, WS_POPUP/WS_CAPTION/WS_THICKFRAME cleared, `GetParent == pane`, client rect = width×height),
  `Resize_FollowsThePane`, `Release_UnparentsTheContent` (`GetParent == Zero`), `Embed_DeadHandle_ReturnsFalse`.
- [ ] Step 2: run → FAIL (type missing).
- [ ] Step 3: implement by moving the P/Invoke + style code out of `XLEdgeTaskPane`; the pane calls the helper.
- [ ] Step 4: suite green.
- [ ] Step 5: commit `XLEdge AIPowered Stage 1: shell pane embedding as a tested helper`.

### Task 3: The core hosts a WebView2 spike page per Excel window

**Files:** Create `XLEdge.Addin.Core\Pane\PaneContentWindow.cs`, `XLEdge.Addin.Core\Pane\TaskPaneContentHost.cs`, `XLEdge.Addin.Core\Pane\SpikePage.html` (embedded resource);
modify `XLEdge.Addin.Core\AddinEntry.cs`; Test `XLEdge.Tests\CorePaneTests.cs`

**Interfaces:**
- `internal sealed class TaskPaneContentHost { IntPtr Create(IntPtr excelWindow); void Close(IntPtr excelWindow); void CloseAll(); void Shown(IntPtr excelWindow); int Count { get; } }`
- `AddinEntry.CreateTaskPaneContent/CloseTaskPaneContent/OnTaskPaneShown` delegate to it; `Shutdown` calls `CloseAll`.
- `AddinEntry.OnRibbonAction`: `RibEdgeLogin` → `RibbonController.ShowTaskPane()`, `RibEdgeShowHide` → `ToggleTaskPane()` (spike wiring; Stage 3 replaces it).
- Spike page sets `document.title = "xledge:ready"` on load; the core logs `PaneContent: command 'ready'` (the title command bus) and
  `PaneContent: navigation completed`. It shows core version, size, `devicePixelRatio`, a text box and a URL box (to open the real login page).

- [ ] Step 1: tests through the real loader and the built package, on an STA thread that pumps messages:
  `CreateTaskPaneContent_ReturnsAHiddenToolWindow` (IsWindow, WS_EX_TOOLWINDOW, not visible in taskbar, not foreground, off-screen),
  `CreateTaskPaneContent_IsPerExcelWindowAndIdempotent`,
  `EmbeddedContent_LoadsTheSpikePage_AndTheTitleBusWorks` (embed with `PaneEmbedding` into a test form; within 30 s the log has both lines),
  `CloseTaskPaneContent_DestroysOnlyThatWindow`.
- [ ] Step 2: run → FAIL.
- [ ] Step 3: implement (`Form` with `CreateParams` adding WS_EX_TOOLWINDOW|WS_EX_NOACTIVATE until embedded, `ShowWithoutActivation`, one environment per core domain).
- [ ] Step 4: suite green ×1.
- [ ] Step 5: commit `XLEdge AIPowered Stage 1: core WebView2 pane content (spike page)`.

### Task 4: Nothing of the core survives an unload; a reload gets a new pane

**Files:** modify `XLEdge.Addin.Core\Pane\*`, `AddinEntry.cs` as the tests demand; Test `XLEdge.Tests\CorePaneTests.cs`

- [ ] Step 1: tests:
  `Unload_WithEmbeddedContent_LeavesNoCoreWindowOnTheThread` (EnumThreadWindows + EnumChildWindows of the test pane: the set after unload equals the set before load; then pump 3 s — no crash),
  `Reload_SameUserDataFolder_SecondCoreShowsThePageAgain` (load A → content → unload → load B → content → log has `ready` twice).
- [ ] Step 2: run → expected FAIL if WinForms leaves its parking window / marshaling control behind.
- [ ] Step 3: fix by the smallest change that passes (dispose forms, then the WinForms thread context's hidden windows without posting WM_QUIT; if that cannot be done safely, switch the content window to WPF `Window` + `Microsoft.Web.WebView2.Wpf` — ledger the ruling either way).
- [ ] Step 4: suite green ×3.
- [ ] Step 5: commit `XLEdge AIPowered Stage 1: core pane survives reload`.

### Task 5: Excel spike checks (needs the developer's go-ahead — combined with Stage 0e Task 6)

- [ ] Register the Debug shell (unregister the installed XLEdge first), start Excel; check: pane opens on Login; spike page; type in the text box (keyboard reaches WebView2);
  open the real login URL; resize pane; second Excel window gets its own page; move Excel to a monitor with another scale; click a cell (focus back to Excel);
  Reload Add-in with the pane open (page comes back, version line changes); Version History back to the previous core; close Excel (no orphan EXCEL.EXE).
- [ ] Record the verdict in spec §4 (pass → Stage 3 builds on it; fail → the §4 fallback, recorded in §6).
