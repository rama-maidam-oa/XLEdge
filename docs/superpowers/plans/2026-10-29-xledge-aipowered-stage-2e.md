# XLEdge AIPowered Stage 2e — Task Pane and Login Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** XLEdge 11.1.2's task pane (WebView2 login, title commands, cookies/token, broadcast messages, logout of every pane) and
its login windows (Server Configuration, About) run in the core, the login survives a hot reload, and the AddinModule handlers
left out of 2d (login, logout, Show/Hide, Param Refresh sheet/book, About, the pane half of WorkbookActivate) are ported.

**Architecture:** The Stage 1 spike content window is replaced by a core class named `XLEdge.ADXExcelTaskPane1` — the content
window of one Excel window's pane (WinForms `Form` holding `WebCtrl`, a WinForms WebView2, re-parented into the shell's ADX pane
as in Stage 1). Its methods are 11.1.2's, unchanged (method-parity test); the ADX members they use (`Visible`, `Activate()`,
`Text`, `Width`, `Show()`, the Before/AfterTaskPaneShow and close events) are replaced in a header by calls on the shell's pane
for that Excel window (contract 4). The core AddinModule gets `adxExcelTaskPanesCollectionItem1`, a stand-in for ADX's pane
collection (`TaskPaneInstance` = the active Excel window's pane, `CreateTaskPaneInstance()`, `TaskPaneInstances`), so 11.1.2's
XLEdgeLogin / LaunchExcelPane / LogOffAllTaskPanesAsync / ParamRefresh compile unchanged.

**Tech Stack:** .NET Framework 4.8.1, WinForms WebView2 1.0.4022.49, WPF-UI, MSTest.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§4, §6 Stage 3 origin check, §7)

## Global Constraints

- Verbatim copies: `Views\XLEdgeServerConfiguration.xaml(.cs)`, `Views\XLEdgeAbout.xaml(.cs)` (pack-URI substitution only).
- Ported methods unchanged (method-parity tests for ADXExcelTaskPane1.cs and AddinModule.cs); every other difference is ledgered.
- Contract 4 (additive, `ContractInfo.Version = 4`): `IRibbonController` gets the pane of a given Excel window — `IntPtr
  GetActiveTaskPaneWindow()`, `bool GetTaskPaneVisible(IntPtr)`, `void SetTaskPaneVisible(IntPtr, bool)`, `void
  SetTaskPaneCaption(IntPtr, string)`, `int GetTaskPaneWidth(IntPtr)`, `void SetTaskPaneWidth(IntPtr, int)`.
- Title commands act only when `CoreWebView2.Source` has the logged-in instance's origin (spec §6) — a ledgered deviation.
- WebView2 environment options and user data folder unchanged (spec §6: identical in every core version).
- GLSense calls (NotifyGLSenseOfLogin/Logout, InvokedFromGLSense) stay stubs that log — Stage 2f.
- 0 errors / 0 warnings; full suite green; local commits only; no Claude attribution.

## Review Focus

1. Pane per Excel window: the right window's pane shown/hidden/captioned with two Excel windows; logout hides every pane.
2. Re-entrancy: core → shell ShowTaskPane → shell creates the pane → core CreateTaskPaneContent, on one thread.
3. A reload with the pane open and logged in: the next core's content logs in again from the session (no second login).
4. Title commands from a page that is not the instance's origin are ignored.
5. WebView2 / fire-and-forget work still running when the core shuts down.

---

### Task 1: Server Configuration and About in the core
- [ ] Parity first; copy both windows; port AddinModule's `NavigateReportsToAddress`, `NormalizeAddress`, `loginButtonVisibility`,
  `RibbonInitialize`, `RibLogin_OnClick`, `RibEdgeAbout_OnClick` (method parity); build; windows open from the core (CoreWindowsTests).

### Task 2: Contract 4 — the pane of an Excel window
- [ ] Contract surface test, shell RibbonController tests (two panes: visible/caption/width per window, active window); implement.

### Task 3: The pane in the core
- [ ] ADXExcelTaskPane1 (content window; method parity with 11.1.2's), pane collection stand-in, TaskPaneContentHost creates it;
  ported handlers: XLEdgeLogin, LaunchExcelPane, LogOffAllTaskPanesAsync, LogoffFromXLEdgeAddin, LogoffFromAddin(Async),
  RibEdgeLogout_OnClick, RibEdgeShowHide_OnClick, ParamRefresh(Book) + scripts, WorkbookActivate (whole), RibEdgeLogout_PropertyChanging.
- [ ] Tests (CorePaneTests style, real WebView2 against a local page): login navigation, cookie token, title command dispatch,
  origin check, logout hides the pane, Shutdown leaves no window or browser process.

### Task 4: The login survives a reload
- [ ] Session bag: login state kept at Shutdown, restored at Initialize (LoggedIn ribbon), cleared at logout; tests; ×3; mutants; review.
