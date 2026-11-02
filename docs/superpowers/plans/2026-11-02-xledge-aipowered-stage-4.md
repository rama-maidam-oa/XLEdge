# XLEdge AIPowered Stage 4 — Remoting Hardening Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the twin-only gaps the AppDomain split leaves in core start, shutdown, unload and the reload busy guard, so
a reload or Release History never unloads the core under running work and a failing core start or stuck unload is visible.

**Architecture:** An audit of the six spec items (leases, COM across awaits, fire-and-forget entry points, Shutdown, 5 s
unload cap, IsBusy) against the code found leases and older-core compatibility already complete (tests exist). The gaps are
fixed where the split put them: `AddinEntry` (start/shutdown), `AddinDomainLoader`/`CoreHost` (unload outcome),
`ExcelThreadContext` + the pane's adapted members (busy accounting). Verbatim 11.1.2 code is not touched, except
`SafeFireAndForget`, which moves from the parity list to the adapted ledger.

**Tech Stack:** .NET Framework 4.8.1, .NET Remoting (AppDomains), WinForms WebView2, MSTest.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§5 row 4, §6, §8 risks).

## Global Constraints

- Verbatim ports stay byte/method-identical to 11.1.2 (parity tests) except the declared Task 6 substitutions;
  `SafeFireAndForget` becomes adapted (Task 4), noted in `CorePortParityTests.PaneMethods`' comment.
- No contract change (`ContractInfo.Version` stays 5). `ICoreLoader.Unload` is shell-internal.
- 0 errors / 0 warnings; full suite ×3; mutants killed; local commits only; no Claude attribution.

## Audit outcome (what is not a task)

- **Already correct (tests exist):** infinite leases on every cross-domain MarshalByRefObject (context, ribbon controller,
  logger, path provider, RemoteLoader/RemoteAddinProxy, AddinEntry); every `IXLEdgeAddin` member catches; no Task crosses;
  core→shell ADX calls on Excel's thread; `_deleteTimer`, WPF windows, pane WebView2s, Excel RCWs released at Shutdown;
  5 s unload cap; reload never inline from a core→shell call; older cores (contract 1 interface unchanged, every shell call
  site try/catch; newer cores refused).
- **11.1.2 behaviour — OISR proposals for the developer:** a report writes into the active workbook read after its HTTP
  awaits, so switching workbooks during a run writes into the other one — **fixed in the twin too (Task 6, the developer's
  request)**; not fixed in the twin: title-command reports run on a pool thread (cross-apartment COM),
  `ImageDownloadHelper` blocks on synchronous HTTP, the wait window is non-modal without an owner.
- **Accepted risks (documented in AIPoweredXLEdge.md):** Excel closing during a report (no bounded wait; as GLSense);
  Excel's thread waits up to 5 s during an unload; `IsBusy` answers "not busy" when the call fails (old/broken core);
  `AppOverlay.BtnHideBusy_Click` (verbatim file) is not counted — it only hides the overlay.

## Rulings

1. **A failing core start fails the load** (Task 1): `AddinEntry.Initialize` rethrows a failure of its core part
   (CoreServices, ribbon helper, ExcelThreadContext, panes, AddinModule) so the loader's rollback and CoreHost's error
   message run; the best-effort parts (session restore, preferences, temp files, WPF-UI) keep logging only. *If wrong:* a
   core that would have half-worked is not loaded — the user gets an error and can reload or roll back.
2. **A stuck unload still loads the next core, and says so** (Task 3): refusing would leave XLEdge without a core. *If
   wrong:* one message box the user may not need.
4. **The run's workbook is captured where the run starts** (Task 6): in `ExcelBulkOperationScope` (every report and refresh
   opens one before its first await), plus around the multi-report and Refresh All loops; only `ActiveWorkbook` reads
   change, through declared parity substitutions. *If wrong:* a run started while one workbook is active and meant for
   another (none known) would go to the first.
3. **Background pane work counts as busy** (Task 4): `SafeFireAndForget` (11.1.2's title-command reports, XLSX downloads,
   multi-report runs) tracks its task in `ExcelThreadContext`; this is the one verbatim method adapted. *If wrong:* a reload
   is refused while a background download finishes.

## Review Focus

1. A reload clicked while a title-command report or download runs on a pool thread → refused with the busy message.
2. A WebView2 handler awaiting (login cookies, broadcast fetch, workbook rerun ids) when Reload is clicked → refused.
3. A core whose start throws → error shown, previous state kept (no "loaded" core that ignores every click).
4. One pane's WebView2 failing to dispose → the other panes, WPF and the Excel RCW are still released.
5. An unload that does not finish in 5 s → the user is told, the next core still loads.
6. Switching workbooks while a report or Refresh All runs → the run finishes in the workbook it started in (Excel tests;
   Inconclusive on this machine, so also the Excel check).

---

### Task 1: A failing core start fails the load

**Files:** Modify `XLEdge.Addin.Core\AddinEntry.cs` (`Initialize`). Test: `XLEdge.Tests\AddinDomainLoaderTests.cs`.

- [ ] **Step 1: Failing test** `Load_CoreStartThrows_LoadThrowsAndDomainIsUnloaded`: the existing loader fixture with a
  context whose `RibbonController` getter throws `InvalidOperationException("no ribbon")` → `Assert.ThrowsExactly` on
  `Load` (message contains `"no ribbon"`), `f.Loader.CurrentDomainName == null`, the log has
  `"AddinDomainLoader.Load: failed to load"`.
- [ ] **Step 2:** run → fails (Load returns an addin).
- [ ] **Step 3:** `Initialize`: the core part (`CoreServices.Initialize` … `_module = new AddinModule(...)`) logs and
  rethrows as `InvalidOperationException("XLEdge core start failed: " + ex.Message)` (no inner exception: core types must
  not cross); `RestoreSessionLogin`, `ApplyStartRibbonState`, `LoadPreferences`, `DeleteTempFiles`, `StartWpfUi` keep
  their own catch-and-log.
- [ ] **Step 4:** loader tests + `bt.sh` green. **Step 5:** commit `XLEdge AIPowered Stage 4: a failing core start fails the load`.

### Task 2: Shutdown runs every step

**Files:** Create `XLEdge.Addin.Core\Infrastructure\ShutdownSteps.cs`; modify `AddinEntry.cs` (`Shutdown`),
`Pane\TaskPaneContentHost.cs` (`CloseAll`). Test: new `XLEdge.Tests\ShutdownStepsTests.cs` (CoreProbe in a domain).

**Interfaces — Produces:** `internal static class ShutdownSteps { public static void RunAll(string owner, params (string Name, Action Step)[] steps) }`
— runs every step in order; a throwing step is logged through `LogUtility.LogException(ex, owner + " (" + name + ")")` and
the next step still runs.

- [ ] **Step 1: Failing tests**
  - `RunAll_AStepThrows_LaterStepsStillRun`: steps a (records), b (throws), c (records) → records `["a", "c"]`, log has
    `"AddinEntry.Shutdown (b)"`.
  - `Shutdown_UsesTheStepRunner` (source test; a real pane cannot be made to fail its Dispose without a production test
    seam): `AddinEntry.Shutdown` calls `ShutdownSteps.RunAll("AddinEntry.Shutdown"` with the steps `"session options"`,
    `"session login"`, `"module"`, `"panes"`, `"wpf"`, `"core services"` in that order, and `TaskPaneContentHost.CloseAll`
    has a `catch (Exception` inside its loop.
- [ ] **Step 2:** run → fail. **Step 3:** implement. **Step 4:** green + `bt.sh`. **Step 5:** commit
  `XLEdge AIPowered Stage 4: Shutdown runs every step even when one fails`.

### Task 3: A stuck unload is reported

**Files:** Modify `XLEdge.Loader.Core\AddinDomainLoader.cs` (`bool Unload(IXLEdgeContext)` — true when the domain is gone
or there was none, false on timeout/failure), `XLEdge\Hosting\HostSeams.cs` (`ICoreLoader.Unload` returns bool),
`XLEdge\Hosting\CoreHost.cs` (remember a failed unload; after the next core loaded, `_ui.ShowError(StuckUnloadMessage, ReloadTitle)`),
`XLEdge.Tests\CoreHostTests.cs` fake (`UnloadResult` settable, default true).
`StuckUnloadMessage` = `"The previous XLEdge core version did not close within 5 seconds. XLEdge keeps working; restart Excel before installing or switching versions again."`

- [ ] **Step 1: Failing tests**
  - `AddinDomainLoaderTests.Unload_DomainThreadStuckInNativeCode_ReturnsFalseWithinTheCap`: the probe in the loaded domain
    starts a thread calling kernel32 `Sleep(9000)` (P/Invoke) → `Unload` returns false in < 6.5 s, log has
    `"did not unload within 5 s"`; a second `Load` of the same release then succeeds.
  - `AddinDomainLoaderTests.Unload_Normal_ReturnsTrue`, `Unload_NothingLoaded_ReturnsTrue`.
  - `CoreHostTests.Reload_PreviousCoreStuck_NextCoreLoadsAndTheUserIsTold`: fake `UnloadResult = false` → reload loads the
    new core and `ui.Errors` contains `StuckUnloadMessage` once; `UnloadResult = true` → no such error.
- [ ] **Step 2:** run → fail (compile). **Step 3:** implement. **Step 4:** green + `bt.sh`. **Step 5:** commit
  `XLEdge AIPowered Stage 4: a stuck unload is reported, the next core still loads`.

### Task 4: Pane work counts as busy

**Files:** Modify `XLEdge.Addin.Core\Infrastructure\ExcelThreadContext.cs`, `ADXExcelTaskPane1.cs` (`SafeFireAndForget`,
`HookWebViewEvents`), `AddinEntry.cs` (`IsBusy`), `XLEdge.Tests\CorePortParityTests.cs` (drop `"void SafeFireAndForget("`
from `PaneMethods`, comment why). Test: `XLEdge.Tests\CoreHandlersTests.cs`.

**Interfaces — Produces:**
- `public static int Background { get; }` — tracked background tasks not yet ended.
- `public static void Track(Task task)` — counts `task` until it ends (any outcome).
- `public static EventHandler<T> Counted<T>(EventHandler<T> handler)` — a handler that invokes `handler` under `Run`, so an
  async-void handler counts in `Pending` until it ends.
- `IsBusy` = dialog || `Pending > 0` || `Background > 0` || `ProcessRunning`; its log line names `background`.

- [ ] **Step 1: Failing tests**
  - `TrackedBackgroundWork_KeepsTheCoreBusyUntilItEnds`: probe tracks a task awaiting a TaskCompletionSource →
    `EntryIsBusy()` true; complete it → false (pumped). Same for a faulted task.
  - `CountedHandler_KeepsTheCoreBusyWhileTheAsyncVoidHandlerAwaits`: `Counted` around an async-void handler awaiting a TCS
    → busy until completed.
  - `PaneBackgroundWork_IsTracked` (source + behaviour): `SafeFireAndForget` passes its `Task.Run` task to
    `ExcelThreadContext.Track`; `HookWebViewEvents` subscribes `SourceChanged` (both handlers) and `DocumentTitleChanged`
    through `ExcelThreadContext.Counted(`.
- [ ] **Step 2:** run → fail. **Step 3:** implement (`Track`: `Interlocked.Increment`, `task.ContinueWith(_ => Decrement,
  TaskContinuationOptions.ExecuteSynchronously)`). **Step 4:** green + `bt.sh`. **Step 5:** commit
  `XLEdge AIPowered Stage 4: pane handlers and background pane work count as busy`.

### Task 5: The core never uses the shell's COM objects

**Files:** Test `XLEdge.Tests\CorePortParityTests.cs` (or a new source test in the same file).

- [ ] `TheCore_NeverUsesTheShellsComProxies`: no `.cs` under `XLEdge.Addin.Core` contains `.ExcelApp` on the context
  (`Context.ExcelApp`, `context.ExcelApp`, `_context.ExcelApp`) or `GetGLSenseAddinInstance(` — those return proxies of the
  shell's RCWs (leases cannot be extended; the 2a/2f rulings). Expected to pass at once (guard against regressions): verify
  it fails by a mutant (add a `Context.ExcelApp` read), then commit
  `XLEdge AIPowered Stage 4: guard - the core never uses the shell's COM proxies`.

### Task 6: A run keeps writing to the workbook it started in (developer's request, 2026-11-02)

11.1.2 reads `ActiveWorkbook` after the HTTP awaits, so switching workbooks while a report loads writes it into the other
one. 11.1.2 already has the field for the fix, never set: `XLEdgeAppState.ActiveWorkbookOverride` (cleared by
`ProgressCoordinator.ResetReportState`). The twin sets it for the length of a run and the post-await reads prefer it. This
changes verbatim files, so the parity tests apply a declared list of exact substitutions to 11.1.2's text before comparing
(everything else stays byte-identical); the same change is proposed to 11.1.2 as an OISR.

**Files:**
- Create `XLEdge.Addin.Core\Infrastructure\RunWorkbookScope.cs` — `internal sealed class RunWorkbookScope : IDisposable`,
  `public static RunWorkbookScope Begin()`: when `ActiveWorkbookOverride` is null, sets it to the active workbook (via
  `ExcelApplicationHelper.GetActiveExcelApplication()?.ActiveWorkbook`, any failure logged) and owns it; `Dispose` clears it
  only when it owns it (nested scopes keep the outer run's workbook).
- Modify `XLEdge.Addin.Core\Helpers\ExcelApplicationHelper.cs` (adapted) — `public static Excel.Workbook RunWorkbook(Excel.Application app)`
  = `XLEdgeAppState.Instance.ActiveWorkbookOverride ?? app?.ActiveWorkbook`.
- Twin substitutions (each an exact `old → new`, `count`), applied in the verbatim/method-parity comparison:
  - `ReportGenerator.cs`: `ExcelBulkOperationScope` holds a `RunWorkbookScope` (`Begin()` in its constructor before
    `ProcessRunning = true`, disposed in `Dispose`); `CreateMultiDataReportsAsync` wraps its loop in
    `using (RunWorkbookScope.Begin())`; `excelApp.ActiveWorkbook` → `ExcelApplicationHelper.RunWorkbook(excelApp)` in
    `BuildLogsSheet` and `BuildReportTable` (11.1.2 lines 798, 941).
  - `ExcelSheetHelper.cs`: every `excelApp.ActiveWorkbook` / `excelApp?.ActiveWorkbook` → `ExcelApplicationHelper.RunWorkbook(excelApp)`
    (outside a run it is the active workbook, as today).
  - `ProgressCoordinator.cs`: `excelApp.ActiveWorkbook.Worksheets["DummySheet"]` → `ExcelApplicationHelper.RunWorkbook(excelApp).Worksheets["DummySheet"]`.
  - `AddinModule.cs` `RibEdgeRefreshAll_OnClick`: the loop runs inside `using (RunWorkbookScope.Begin())`, and its
    `excelApp.ActiveWorkbook` reads → `ExcelApplicationHelper.RunWorkbook(excelApp)`.
  - Not changed: `ActiveWindow` / `ActiveCell` / `ActiveSheet` reads (they follow the report's own sheet activation, or are
    cosmetic end-of-run reads — audit 2f).
- Modify `XLEdge.Tests\CorePortParityTests.cs`: `TwinSubstitutions` — `(string File, string Old, string New, int Count)[]`;
  `Normalized(original)` applies the file's substitutions (asserting each `Old` occurs exactly `Count` times in 11.1.2's
  text) before the byte/method comparison; these three files stay in `Verbatim`.
- Test: `XLEdge.Tests\CoreHandlersExcelTests.cs` (real Excel; Inconclusive on this machine — see Machine notes).

- [ ] **Step 1: Failing tests**
  - `CorePortParityTests.TwinSubstitutions_AreExactAndComplete`: every substitution's `Old` occurs `Count` times in 11.1.2
    and `New` the same number of times in the core; no other difference (the existing byte/method tests, now applying them).
  - `CoreHandlersTests.RunWorkbookScope_WithoutExcel_SetsNothingAndNeverThrows` (no Excel here, so no Workbook object can
    exist): `Begin()` → override stays null, no exception, a warning logged; an `ExcelBulkOperationScope` constructed and
    disposed → `ProcessRunning` true then false, override null. (Ownership across nested scopes is covered by the Excel tests.)
  - `CoreHandlersExcelTests.Report_WorkbookSwitchedDuringFetch_IsWrittenToTheStartingWorkbook`: two workbooks A, B; a
    report run against `InstanceServer` whose CSV response is held by a gate; activate B while held; release → the report
    sheet is in A, B unchanged.
  - `CoreHandlersExcelTests.RefreshAll_WorkbookSwitchedMidLoop_RefreshesOnlyTheStartingWorkbook`.
- [ ] **Step 2:** run → the parity and no-Excel tests fail; the Excel tests are Inconclusive here (record it).
- [ ] **Step 3:** implement. **Step 4:** green + `bt.sh`. **Step 5:** commit
  `XLEdge AIPowered Stage 4: a report or refresh keeps writing to the workbook it started in`.

### Task 7: Close the stage

- [ ] Fresh opus reviewer for the whole stage; one fix pass.
- [ ] `bt.sh` ×3; mutants: Initialize swallowing again; `RunAll` stopping at the first failure; `Unload` always true;
  CoreHost not telling the user; `Track` not decrementing on fault; `Counted` not using `Run`; `IsBusy` ignoring `Background`.
- [ ] `AIPoweredXLEdge.md`: Stage 4 done; the OISR proposals and accepted risks above; Excel check items (Reload during a
  title-command report / XLSX download → refused; Reload during login → refused; switch workbooks while a report and
  while Refresh All runs → written to the starting workbook); the twin now differs from 11.1.2 by the Task 6 substitutions
  until 11.1.2 gets the OISR.
- [ ] Commit.
