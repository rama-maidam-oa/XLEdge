# XLEdge AIPowered Stage 2c — Windows, Themes, API Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move XLEdge 11.1.2's windows, theme, converters and window/API helpers into the core verbatim, so the windows open
from the core (hot-reloadable, owned by Excel), and close the two items carried from 2b (HTTP settings in the core's domain,
session-only options across a reload).

**Architecture:** XAML and code-behind copied unchanged except one forced substitution: pack URIs naming the assembly
`XLEdge` now name `XLEdge.Addin.Core` (the parity test allows exactly that). WPF-UI is initialised in `AddinEntry.Initialize`
as 11.1.2 did on ribbon load. A small **session bag** in the shell's context (contract 3) carries values a hot reload must not
lose — first user: session-only options; Stage 2e reuses it for the login.

**Tech Stack:** .NET Framework 4.8.1, WPF, WPF-UI 4.3.0 (+Abstractions), MahApps.Metro.IconPacks Core/FontAwesome 6.2.1,
Microsoft.Xaml.Behaviors.Wpf 1.1.142, MSTest 3.11.1.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md`

## Global Constraints

- Verbatim copies (BOM/line endings aside); the only allowed change is `/XLEdge;component/` → `/XLEdge.Addin.Core;component/` in
  pack URIs. Anything else is a ledgered ruling.
- Not in 2c: `XLEdgeServerConfiguration` (login server picker; calls AddinModule's pane navigation and ribbon init → 2e),
  `MahAppsBootstrapper` (retired stub, not compiled in 11.1.2).
- Contract 3: `IXLEdgeContext.GetSessionValue(string key)` / `SetSessionValue(string key, string value)` (shell-side, survives core
  reloads, gone when Excel closes).
- 0 errors / 0 warnings; full suite green; local commits only on `11.1.2-aipowered`; no Claude attribution; bugs found are fixed or
  ruled in this sub-stage.

## Review Focus

1. XAML resources resolving in a non-default AppDomain (pack URIs, merged dictionaries, WPF-UI theme, images) — a missing resource
   only fails when the window opens.
2. Windows owned by Excel's main window (ExcelWindowHelper) from the core's AppDomain; no window left after a reload.
3. `ServicePointManager` / TLS settings made by APIHelper's static constructor applying in the core's domain.
4. Session-only options surviving a reload, and saved options never overwritten by session ones.
5. `GetOpenWindowTitles` (reload confirmation) listing exactly the core's open windows.

---

### Task 1: Windows, theme and helpers compile in the core
Copy `Views\{AppOverlay, XLEdgeAbout, XLEdgeCalendar, XLEdgeDrilldownReports, XLEdgeGLAccountsWindow, XLEdgeLoginDetails,
XLEdgeMessageWindow, XLEdgeOptions, XLEdgeWaitWindow}.xaml(.cs)`, `Themes\GlobalStyles.xaml`, `Converters\WidthPercentageConverter.cs`,
`Utilities\{DpiAwareWindow, MessageFunctions, WpfUiBootstrapper}.cs`, `Helpers\{ApiOperationHelper, APIHelper, ExcelWindowHelper,
EnhancedDragDropHelper}.cs`, `Images\{orbit_logo.png, OrbitGLSense.ico, orbit_bitmap.bmp}`; packages as above.
- [ ] Parity test first (XAML included, URI substitution allowed); copy; build 0/0; commit.

### Task 2: The windows open from the core
- [ ] Tests (STA, core AppDomain via CoreProbe): WPF-UI initialised by `AddinEntry.Initialize`; each window that needs no Excel
  constructs, shows and closes (About, Calendar, LoginDetails, MessageWindow, Options, WaitWindow) with its title and the
  GlobalStyles / WPF-UI resources resolved; `GetOpenWindowTitles` lists open core windows; after Shutdown + unload no managed window.
- [ ] Implement start-up and GetOpenWindowTitles; commit.

### Task 3: HTTP in the core's domain
- [ ] Tests: after APIHelper is used in the core's domain, `ServicePointManager` has 11.1.2's settings (TLS 1.2|1.3, connection limit
  100, Expect100Continue off, idle 5 s); APIHelper against a local `HttpListener`: Bearer header, response, error mapping.

### Task 4: Session-only options survive a reload (contract 3)
- [ ] Tests: Options › Apply (session only) → reload → applied values are back, saved file unchanged; Options › Save → reload →
  saved values; Excel restart (new shell context) → saved values only.
- [ ] Implement session bag (shell) + AddinEntry capture at Shutdown / restore after preferences load; commit; ×3; mutants; review.
