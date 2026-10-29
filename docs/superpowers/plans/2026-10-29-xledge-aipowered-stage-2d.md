# XLEdge AIPowered Stage 2d — Report Engine and Ribbon/Excel Handlers Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move XLEdge 11.1.2's report engine into the core verbatim, and the AddinModule handlers that need neither the task pane
nor the login flow (those are 2e), so the ribbon buttons and Excel events the shell forwards run 11.1.2's code in the core.

**Architecture:** The engine files are copied unchanged. The core gets a plain class `XLEdge.AddinModule` (not an ADX module):
it carries 11.1.2's handler methods with their bodies unchanged — only ADX signatures and members are replaced (`IRibbonControl
control` → `object control`; `RibSheetLabel.Caption` etc. through a small ribbon-control shim over `XLEdgeRibbonHelper`) — so the
engine's `XLEdge.AddinModule.CurrentInstance` / `ApplyRibbonState` / `UpdateTabLabel` / `GetPaneInstance` calls compile as they
are. `AddinEntry.OnRibbonAction` / `OnExcelEvent` dispatch to those handlers; event arguments arrive as names/addresses (the shell
sends primitives) and are resolved to COM objects through the core's own Excel RCW.

**Tech Stack:** .NET Framework 4.8.1, Excel interop, WPF, MSTest 3.11.1; real Excel (automation instance) in `[TestCategory("Excel")]` tests.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md`

## Global Constraints

- Verbatim engine copies: `Helpers\{ReportGenerator, ParamsControlSheetBuilder, XLEdgeParamsBuilder, DrilldownRequestBuilder,
  ReportParameterRequestSerializer, ProgressCoordinator}.cs` (parity test; only the pack-URI substitution allowed).
- `AddinModule.cs` in the core: every ported method's body identical to 11.1.2's (method-level parity test), the allowed
  signature substitution `IRibbonControl control` → `object control`; any other difference is a listed, ledgered adaptation.
- Not in 2d (→ 2e): login, pane, GLSense calls, Show/Hide, ParamRefresh (sheet/book — they script the pane), About,
  RibbonInitialize, the pane half of WorkbookActivate. Not ported: `LogHelper` (the shell's Logger owns the file and its
  environment header, Stage 0c), `XLEdgeRibbonReflectionHelper` (the ribbon is reached through `IRibbonController`).
- Async handlers resume on Excel's thread (never on a pool thread) — COM calls from the core stay on Excel's STA.
- 0 errors / 0 warnings; full suite green; local commits only on `11.1.2-aipowered`; no Claude attribution; bugs found are fixed or
  ruled in this sub-stage.

## Review Focus

1. Continuations of the `async void` handlers after an `await` in the core's AppDomain: must run on Excel's thread.
2. Event arguments resolved by name: a sheet name that exists in several open workbooks, a deleted sheet, an address with `$`.
3. A reload while a refresh/report runs (`IsBusy`), and the delete timer / WPF windows of a handler at Shutdown.
4. COM objects the handlers create through the core's RCW: no Excel process left after the tests' Excel quits.
5. BigInteger / decimal / int values through `ReportParameterRequestSerializer` — same JSON as 11.1.2.

---

### Task 1: Engine and AddinModule compile in the core
- [ ] Parity tests first: the six engine files in `Verbatim`; new `AddinModuleMethods_MatchXLEdge11_1_2` comparing each listed
  method's text (signature line to its closing brace) with 11.1.2's, after the one signature substitution.
- [ ] Copy the engine; write the core `AddinModule.cs` (handlers listed in the test, `CurrentInstance`, `ExcelApp`, ribbon shims,
  `GetPaneInstance()` returning null until 2e, `_deleteTimer`); build 0/0; commit.

### Task 2: Golden values for the request builders
- [ ] Golden cases (run against 11.1.2's XLEdge.dll, then the core): `ReportParameterRequestSerializer.Serialize` with int, long,
  decimal, double, bool, string, null, BigInteger, `Values` lists and extra parameters; `DrilldownRequestBuilder.GetColumnType`
  and `ResolveStoredParamValue` (scalar, list, missing, not-an-array, bad JSON).
- [ ] Regenerate GoldenValues.cs from 11.1.2; core matches; commit.

### Task 3: Dispatch, Excel's thread, busy, cleanup
- [ ] Tests (CoreProbe, STA): an async handler's continuation after `await` runs on the calling (Excel) thread; `IsBusy` true while
  an async handler is in flight and while `ProcessRunning`; Shutdown stops/disposes the delete timer and clears pending deletes;
  temp files in `XLEdgeAppPaths.TempFolder` are deleted at Initialize (11.1.2 did it on ribbon load); unknown actions/events are
  logged and ignored; each forwarded ribbon id reaches its handler.
- [ ] Implement: a synchronization context for Excel's thread that counts outstanding async-void operations; `OnRibbonAction` /
  `OnExcelEvent` dispatch; Shutdown order; the shell appends the sheet's workbook name to every sheet event's arguments (last,
  so a core that reads fewer arguments is unaffected); commit.

### Task 4: Handlers against a real Excel
- [ ] `[TestCategory("Excel")]` tests with an automation Excel (killed by PID if it outlives the test): SheetBeforeDelete queues the
  companion `P_` sheet and the timer deletes it; "Goto Report Data" activates the report's sheet; SheetActivate updates the tab
  label caption and DeleteNamedCache removes orphaned `ORB_` names; an event naming a sheet in a non-active workbook resolves to
  that workbook's sheet; Control Sheet builds `orb_params_control`; no Excel process left.
- [ ] Fix what they find; ×3; mutants; review.
