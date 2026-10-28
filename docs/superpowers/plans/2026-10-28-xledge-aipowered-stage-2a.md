# XLEdge AIPowered Stage 2a — Core Foundation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give the core the foundation today's XLEdge code stands on — Excel, logging, paths, app state, ribbon, WPF dispatcher —
re-implemented over the contracts with today's names and members, so Stages 2b–2f can move XLEdge 11.1.2 code in nearly
line for line.

**Architecture:** Ported code keeps calling `LogUtility`, `XLEdgeAppPaths`, `XLApp`, `XLEdgeAppState`, `XLEdgeRibbonHelper`,
`UiDispatcher` and `WpfAppManager`. In the core these are facades over `IXLEdgeContext` (held by `CoreServices`). Excel reaches
the core as a native RCW built from the shell's IDispatch pointer — not a remoting proxy — so COM-heavy code runs at native
speed and cannot hit CrossAppDomainSink cast failures.

**Tech Stack:** .NET Framework 4.8.1, Excel interop (same Interops\Microsoft.Office.Interop.Excel.dll as the shell), WPF, MSTest 3.11.1.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md`

## Stage 2 overview (sub-stages, each with its own plan, review and commits)

| Sub-stage | Content | Depends on |
|---|---|---|
| **2a** | Core foundation (this plan) | — |
| 2b | Pure helpers, models, JSON, API and formatting (no Excel, no windows) + characterization tests | 2a |
| 2c | Windows, themes, converters (WPF-UI, MahApps icons) in the core; owned by Excel; `GetOpenWindowTitles` | 2a |
| 2d | Report engine and sheet builders; AddinModule's business handlers (ribbon actions, Excel events, refresh, drilldown, attachments) as core handlers; `IsBusy` | 2b, 2c |
| 2e | Pane and login (spec phase 3, pulled forward: the GLSense calls need it) — title command bus, login/logout, cookies, Bearer header, broadcast messages, `Source` origin check | 2d |
| 2f | XLEdge ↔ GLSense in **both** AIPowered code bases, matching XLEdge 11.1.2 ↔ GLSense 11.1.2 (`FinalWorkingCode`): XLEdge core `OnLoginFromGLSense` / `OnLogoutFromGLSense` and its outgoing `GetGLCubeInformation` / `LogoutSession`; GLSense AIPowered shell COM methods `GetGLCubeInformation` / `LogoutSession` forwarded into its core; Excel check with both twins | 2e |

## Global Constraints

- Port, don't rewrite: XLEdge 11.1.2 (`D:\SQLLite_Test\XLEdge\XLEdge`) code moves with only the changes the AppDomain boundary forces; every other change is a ledgered ruling.
- The core references no Add-in Express assembly; only primitives (and the Excel IDispatch pointer) cross the boundary.
- `ContractInfo.Version` goes from 1 to 2 for the contract addition in Task 1; shell and core ship together.
- 0 errors / 0 warnings, Debug and Release; full suite green; local commits only on `11.1.2-aipowered`; no Claude attribution.
- Any bug or warning found is fixed in the same sub-stage, with a test.

## Review Focus

1. The Excel RCW in the core outliving the core (COM reference leak keeps EXCEL.EXE alive after close) or released twice.
2. Logging from ported code before `Initialize` or after `Shutdown` (static facades with no context) — must not throw.
3. `XLEdgeRibbonHelper` getters for state only the core sets (Enabled) after a reload: the new core starts with no shadow state.
4. A WPF dispatcher / Application created in the core domain surviving unload (same check as CorePaneTests).
5. Debug mode read by ported code (`LogUtility.DebugMode`) must follow the shell's Debug toggle.

---

### Task 1: The core gets its own Excel RCW

**Files:** modify `XLEdge.Contracts\IXLEdgeContext.cs`, `XLEdge.Contracts\ContractInfo.cs`, `XLEdge\XLEdgeContext.cs`; create
`XLEdge.Addin.Core\Infrastructure\CoreServices.cs`; Test `XLEdge.Tests\CoreExcelTests.cs`, `ContractSurfaceTests.cs`

**Interfaces:** `IntPtr IXLEdgeContext.GetExcelDispatch()` — an AddRef'ed IDispatch of the Excel Application (Zero if none), the
caller releases it. `CoreServices.Initialize(IXLEdgeContext)`, `CoreServices.Context`, `CoreServices.ExcelApp` (lazy native RCW
via `Marshal.GetObjectForIUnknown`, released in `CoreServices.Shutdown()`).

- [ ] Tests (COM object: `Scripting.Dictionary` stands in for Excel): the core's object is a COM object, not a transparent proxy;
  calls on it work; `Shutdown` releases it (the shell-side object's ref count returns to its value before); context without
  Excel → `ExcelApp` null, no throw; contract surface lists the member; `ContractInfo.Version == 2`.
- [ ] Implement; commit `XLEdge AIPowered Stage 2a: the core gets its own Excel RCW (contract 2)`.

### Task 2: Logging, paths and app state facades

**Files:** create in `XLEdge.Addin.Core`: `Utilities\LogUtility.cs`, `XLEdgeAppPaths.cs`, `XLApp.cs`, `XLEdgeAppState.cs`; Test
`XLEdge.Tests\CoreFacadeTests.cs` (run inside a core AppDomain through a test probe class compiled into the core: `Infrastructure\CoreSelfTest.cs`, internal, called via a `RemoteScopeCaller`-style helper)

**Interfaces:** today's public members unchanged (see XLEdge 11.1.2 files of the same names). `LogUtility.*` → `ILogger`
(exceptions as `ToString()`, `LogScope` → `BeginLogScope`, `DebugMode` → `context.DebugMode`); `XLEdgeAppPaths.*` →
`IPathProvider` (BaseFolder = folder of `UrlsFile`, XLEdgeLogsFolder = `Root`, LogFolder = `Logs`, TempFolder = `Temp`,
TempUrlsPath = `UrlsFile`, BrowserLogsFolder = `BrowserDataPath`, LoginBrowserLogsPath = `BrowserDataPath\Login`);
`XLEdgeAppState` = today's file minus the ADX pane field, `ExcelApp` → `CoreServices.ExcelApp`.

- [ ] Tests: each facade maps as above; LogUtility before Initialize / after Shutdown does not throw; scopes buffer through the
  shell logger; DebugMode follows the context.
- [ ] Implement by copying today's files and changing only the backing; commit.

### Task 3: Ribbon facade and WPF dispatcher in the core

**Files:** create `XLEdge.Addin.Core\Helpers\XLEdgeRibbonHelper.cs` (today's API over `IRibbonController`, shadow Enabled state +
`RestorePreviousState`), `Utilities\UiDispatcher.cs`, `Utilities\WpfAppManager.cs`; Test `CoreFacadeTests.cs`, `CorePaneTests.cs`

- [ ] Tests: every `XLEdgeRibbonHelper` member reaches the controller (fake controller); `ApplyState` delegates to `SetState`;
  shadow Enabled after Enable/Disable/ApplyState; `RestorePreviousState`; `WpfAppManager.EnsureApplication` + `UiDispatcher.Run`
  on the core domain's STA thread; unload afterwards leaves no managed window.
- [ ] Implement; commit; full suite ×3; mutants; review; fix pass.
