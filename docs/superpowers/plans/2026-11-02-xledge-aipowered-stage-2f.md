# XLEdge AIPowered Stage 2f — XLEdge ↔ GLSense Calls Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** The login/logout sync between XLEdge and GLSense works in both directions between the two AIPowered add-ins, with
11.1.2 XLEdge's behaviour: GLSense → XLEdge (`InvokedFromGLSense`, `LogoffFromAddin`) reaches 11.1.2's code in the XLEdge core,
and XLEdge → GLSense (`GetGLCubeInformation`, `LogoutSession`) reaches GLSense AIPowered, which gains those two methods.

**Architecture:** XLEdge core keeps 11.1.2's `InvokedFromGLSense`, `NotifyGLSenseOfLogin` and `NotifyGLSenseOfLogout` verbatim
(method parity), including their late-bound `InvokeMember` on `_glSenseAddinInstance`. Only `GetGLSenseAddinObject()` is adapted:
instead of GLSense's COM object it returns a core-side stand-in (`GLSenseAddin`) whose `GetGLCubeInformation` / `LogoutSession`
ask the shell to make the COM call (contract 5). The shell looks GLSense up and invokes it in its own AppDomain, so no COM object
crosses the AppDomain boundary — only strings. GLSense AIPowered's shell gets the two public COM methods 11.1.2 XLEdge calls,
forwarding to its core through the existing `OnRibbonAction` dispatch (no GLSense contract change).

**Tech Stack:** .NET Framework 4.8.1, Add-in Express, .NET Remoting (AppDomains), MSTest.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§3 contracts, §6.5, §7 "Login … from GLSense",
§9 last item). References: 11.1.2 XLEdge `XLEdge\AddinModule.cs` (`InvokedFromGLSense`, `NotifyGLSenseOfLogin`,
`NotifyGLSenseOfLogout`, `GetGLSenseAddinObject`, `LogoffFromAddin`); GLSense `FinalWorkingCode\GLSense\AddinModule.cs`
(`GetGLCubeInformation`, `LogoutSession`, `LoggOff`).

## Global Constraints

- XLEdge: verbatim methods stay byte-identical to 11.1.2 (method parity, `IRibbonControl` → `object` only); every other
  difference is ledgered in this plan.
- Contract 5 (additive, `ContractInfo.Version = 5`) on `IXLEdgeContext`: `bool IsGLSenseConnected()`,
  `void CallGLSense(string method, object[] args)`. Only strings/bools cross; the loader already refuses a core newer than the shell.
- `IXLEdgeContext.GetGLSenseAddinInstance()` (contract 1) stays as it is (contracts are additive); the core does not call it.
- Neither add-in logs a token: XLEdge's shell logs the method name only; GLSense passes the arguments as `object[]`, which its
  `OnRibbonAction` entry log prints as `System.Object[]`.
- GLSense changes: AIPowered only, in the `11.1.2` worktree `D:\SQLLite_Test\worktrees\GLSense-11.1.2\AIPowered\GLSense`, local
  commit, never pushed; `FinalWorkingCode` is reference only (its uncommitted `CLAUDE.md` change is the developer's — untouched).
  GLSense's fix log (`AIPowered\GLSense\CLAUDE.md`) gets an entry, in its numbered style.
- XLEdge: 0 errors / 0 warnings, full suite green ×3, mutants killed. GLSense: Debug + Release build with no new warnings
  (baseline counted before the change). Local commits only; no Claude attribution.

## Rulings (decisions taken for the developer; veto any)

1. **The shell makes the COM call, not the core** (refines the resume notes' "through `GetGLSenseAddinInstance()`"): a COM
   object returned through remoting reaches the core as a proxy to the shell's RCW, and late-bound `InvokeMember` through it is
   not reliable. *If wrong:* nothing breaks — it is only more code than needed.
2. **The GLSense lookup survives a reload:** 11.1.2 cached GLSense's object in the module, which a reload drops, so a logout
   after a reload would not tell GLSense. When the restored session login was shared with GLSense (`LoginFromGLSense` or
   `LoginSentToGLSense`), the new core looks GLSense up again at start. *If wrong:* GLSense gets a `LogoutSession` it would also
   have got without a reload.
3. **GLSense, no cubes or a failed cube call after an XLEdge login → GLSense stays `LoggedOut`** (warning logged).
   FinalWorkingCode applied a `NoCubes` state that AIPowered GLSense does not have (it disabled even Login). *If wrong:* a user
   without cubes sees GLSense's Login enabled instead of a fully disabled ribbon.
4. **GLSense fetches the cubes without blocking Excel:** FinalWorkingCode waited synchronously (up to 30 min) inside the COM
   call; AIPowered runs it fire-and-forget like its `Logout`, so XLEdge's call returns at once. *If wrong:* none known — XLEdge
   ignores the result either way.

## Review Focus

1. Logout ping-pong: GLSense logout → XLEdge `LogoffFromAddin` → XLEdge's logout calls GLSense `LogoutSession` back while
   GLSense's own logout is in progress — GLSense must end `LoggedOut` once, with one message box (Task 3 test + Excel check).
2. GLSense not installed, disconnected, or an older GLSense without the two methods: XLEdge login/logout complete, one warning
   logged, nothing shown to the user (Task 1/2 tests: not connected, `CallGLSense` throws).
3. Login from GLSense while XLEdge's pane is open: the pane navigates with the shared login (11.1.2's
   `RefreshLoginNavigationAsync`) — Excel check, pane code is unchanged 11.1.2.
4. A user without XLEdge permission logged in through GLSense: XLEdge stays logged out with Login disabled (Task 2 test).
5. A reload between login and logout (Ruling 2) (Task 2 test).

---

### Task 1: Contract 5 — the shell calls GLSense

**Files:**
- Modify: `XLEdge.Contracts\IXLEdgeContext.cs`, `XLEdge.Contracts\ContractInfo.cs` (Version 5)
- Modify: `XLEdge\XLEdgeContext.cs` (the two members; lookup shared with `GetGLSenseAddinInstance`)
- Modify: `XLEdge.Tests\FakeContext.cs`, `XLEdge.Tests\CoreExcelTests.cs` (its own fake context): implement the two members
  (FakeContext records calls as `"method|arg1|arg2|…"` in `GLSenseCalls`, `GLSenseConnected` settable, `GLSenseThrows` settable)
- Test: `XLEdge.Tests\ContractSurfaceTests.cs`, `XLEdge.Tests\XLEdgeContextTests.cs`

**Interfaces — Produces:**
- `bool IXLEdgeContext.IsGLSenseConnected()` — GLSense's COM add-in is in `COMAddIns` and connected (looked up now).
- `void IXLEdgeContext.CallGLSense(string method, object[] args)` — looks GLSense up now; not connected → returns without a
  call; otherwise `InvokeMember(method, InvokeMethod, …, args)` on its `.Object`. A failure is logged (method name only) and
  rethrown as `InvalidOperationException("GLSense " + method + " failed: " + inner message)` (serializable across remoting).

- [ ] **Step 1: Failing tests**
  - `ContractSurfaceTests.IXLEdgeContext_Members`: adds `"bool IsGLSenseConnected()"`,
    `"void CallGLSense(string method, object[] args)"`; the contract-version test expects `5`.
  - `XLEdgeContextTests.Context_GLSense_IsConnected_FollowsConnect`: `FakeExcel` → true; `Only.Connect = false` → false;
    ProgId `"Other.AddinModule"` → false.
  - `XLEdgeContextTests.Context_CallGLSense_InvokesTheConnectedAddin`: `Only.Object = new GLSenseRecorder()` (public methods
    `GetGLCubeInformation(string, string, string)`, `LogoutSession()`, recording `"GetGLCubeInformation|tok|http://x|u"`,
    `"LogoutSession"`); both calls recorded in order; with `Connect = false` a call records nothing and does not throw.
  - `XLEdgeContextTests.Context_CallGLSense_FailureComesBackAsInvalidOperation`: recorder method throws
    `new COMException("boom")` → `InvalidOperationException` whose message contains `"LogoutSession"` and `"boom"`; the logger
    has a warning/error naming `LogoutSession` and no argument value (`"tok"` not in the log).
- [ ] **Step 2:** `rt.sh` the four tests → fail (compile error on the new members).
- [ ] **Step 3:** Implement; `GetGLSenseAddinInstance` and the two new members share one private lookup (`FindGLSense(out bool connected)`).
- [ ] **Step 4:** Tests pass; `bt.sh` whole suite green, 0 warnings.
- [ ] **Step 5:** Commit `XLEdge AIPowered Stage 2f: contract 5 - the shell calls GLSense`.

### Task 2: 11.1.2's GLSense sync in the XLEdge core

**Files:**
- Create: `XLEdge.Addin.Core\Infrastructure\GLSenseAddin.cs` — `internal sealed class GLSenseAddin` with public
  `void GetGLCubeInformation(string token, string url, string userName)` and `void LogoutSession()`, each
  `ExcelThreadContext.OnExcelThread(() => CoreServices.Context.CallGLSense(nameof(…), args))`. Public instance methods,
  because 11.1.2 calls them with `GetType().InvokeMember`.
- Modify: `XLEdge.Addin.Core\AddinModule.cs` — replace the two logging stubs with 11.1.2's `NotifyGLSenseOfLogin`,
  `NotifyGLSenseOfLogout` and add `InvokedFromGLSense` and `private object _glSenseAddinInstance;` (all verbatim); add the
  adapted `public object GetGLSenseAddinObject()` = `CoreServices.Context` connected (via `OnExcelThread`) ? `new GLSenseAddin()`
  : null, catching and logging like 11.1.2 (`LogDebug`, returns null); add the `RibbonControl` facade members 11.1.2's
  `InvokedFromGLSense` uses: `Pressed` (via `RibbonHelper.Get/SetControlPressed`), fields `RibEdgeDialogBoxLauncher`,
  `RibEdgeDebug`, `RibEdgeIncludeOutputData`, `RibControlSheet`, `RibLoginURL`.
- Modify: `XLEdge.Addin.Core\AddinModule.Dispatch.cs` — `internal void ReattachGLSense()`: when
  `XLEdgeAppState.Instance.LoginFromGLSense || LoginSentToGLSense`, `_glSenseAddinInstance = GetGLSenseAddinObject()` (Ruling 2).
- Modify: `XLEdge.Addin.Core\AddinEntry.cs` — `OnLoginFromGLSense` → `ExcelThreadContext.Run(() =>
  _module?.InvokedFromGLSense(instanceName, loginUrl, token, userName, hasXLEdgePermission))`; `OnLogoutFromGLSense` →
  `Run(() => _module?.LogoffFromAddin())`; both try/catch-logged like the other entry points; `ReattachGLSense()` right after
  `RestoreSessionLogin()`.
- Test: `XLEdge.Tests\CorePortParityTests.cs` (AddinModuleMethods adds `"void InvokedFromGLSense(string eeName"`,
  `"void NotifyGLSenseOfLogin(string authToken"`, `"void NotifyGLSenseOfLogout()"`), new `XLEdge.Tests\CoreGLSenseTests.cs`
  (CoreHandlersTests' `WithCore` pattern; `CoreProbe` gets what it needs to call module methods).

**Interfaces — Consumes:** Task 1's two members.

- [ ] **Step 1: Failing tests** (`CoreGLSenseTests`; `FakeContext.GLSenseConnected = true` unless stated)
  - `OwnLogin_TellsGLSenseOnce`: `NotifyGLSenseOfLogin("tok", "http://x", "u")` twice → `GLSenseCalls` ==
    `["GetGLCubeInformation|tok|http://x|u"]`; `LoginSentToGLSense` true.
  - `OwnLogin_GLSenseNotThere_TriesAgainNextTime`: connected false → no call, `LoginSentToGLSense` false; connected true, call
    again → one call.
  - `OwnLogin_NotRepeatedBack_ForALoginFromGLSense` and `…_WithoutAToken`: `LoginFromGLSense = true` / token `""` → no call.
  - `LoginFromGLSense_LogsXLEdgeIn_AsInXLEdge11_1_2`: entry `OnLoginFromGLSense("Prod", "http://x", "tok", "u", true)` →
    AppState `LoginUrlName "Prod"`, `LoginUrl`, `LoginToken`, `LoginUserName`, `LoginFromGLSense`, `IsLoginCompleted` all set;
    ribbon calls contain `Visible:RibEdgeLogin=False`, `Visible:RibEdgeLogout=True`, the `RibEdgeLogout` caption `Prod`,
    `Enabled:RibEdgeBroadcast=True`, `Enabled:RibEdgeDialogBoxLauncher=True`, `Pressed:RibEdgeDebug=False`, the `RibLoginURL`
    caption `http://x`; `GLSenseCalls` empty (not echoed back).
  - `LoginFromGLSense_WithoutXLEdgePermission_StaysLoggedOut`: `…, false)` → `IsLoginCompleted` false; `Enabled:…=False` for
    `RibEdgeLogin`, `RibControlSheet`, `RibEdgeDebug`, `RibEdgeIncludeOutputData`, `RibEdgeBroadcast`.
  - `LogoutFromGLSense_LogsXLEdgeOut_AndTellsGLSense`: login from GLSense, then entry `OnLogoutFromGLSense()` → (pump until)
    `IsLoginCompleted` false, `LoginFromGLSense` false, ribbon ends `LoggedOut`, `GLSenseCalls == ["LogoutSession"]` (11.1.2
    calls back), `EntryIsBusy()` false at the end.
  - `OwnLogout_TellsGLSense_OnlyAfterASharedLogin`: logout with no sync → no call; after `NotifyGLSenseOfLogin` → `LogoutSession`.
  - `GLSenseFailing_DoesNotStopXLEdge`: `GLSenseThrows = true` → login-from-own still sets `LoginSentToGLSense`, logout still
    completes; the log has 11.1.2's `"Failed to invoke GetGLCubeInformation on GLSense addin."` / `"…LogoutSession…"`.
  - `Reload_KeepsTellingGLSense` (Ruling 2): session value `XLEdge.Login` with `LoginSentToGLSense:true, IsLoginCompleted:true`
    → new core → logout → `GLSenseCalls == ["LogoutSession"]`.
  - Parity: the three methods listed in `AddinModuleMethods`.
- [ ] **Step 2:** `rt.sh` them → fail.
- [ ] **Step 3:** Implement as listed under Files.
- [ ] **Step 4:** Tests pass; `bt.sh` green, 0 warnings.
- [ ] **Step 5:** Commit `XLEdge AIPowered Stage 2f: 11.1.2's GLSense login and logout sync in the core`.

### Task 3: GLSense AIPowered exposes GetGLCubeInformation and LogoutSession

Working tree: `D:\SQLLite_Test\worktrees\GLSense-11.1.2\AIPowered\GLSense` (branch `11.1.2`). No test project exists there; the
proof is the build, a source check from the XLEdge suite, and the Excel check.

**Files:**
- Modify: `GLSense\AddinModule.cs` — two public methods next to the ribbon handlers, each try/catch-logged:
  `public void GetGLCubeInformation(string token, string url, string userName)` →
  `GlobalsEx.Addin?.OnRibbonAction("LoginFromXLEdge", new object[] { token, url, userName })`;
  `public void LogoutSession()` → `GlobalsEx.Addin?.OnRibbonAction("LogoutFromXLEdge", null)`. An older core logs the action
  as unknown and does nothing.
- Modify: `GLSense.Addin.Core\AddinEntry.cs` — `case "LoginFromXLEdge": _ = LoginFromXLEdge(parameter as object[]);` and
  `case "LogoutFromXLEdge": LogoutFromXLEdge();`
  - `private async Task LoginFromXLEdge(object[] args)`: FinalWorkingCode's `GetGLCubeInformation` on AIPowered types — blank
    token or url → return; `AppState.Instance.LoginToken/LoginUrl`; POST `{LoginUrl}{AppConstants.RestSecure}finance-cubes` with
    `ApiHelper.ServerAPI(url, "Form", "", "POST", token)` (awaited, Ruling 4); empty response, parse failure or zero cubes →
    `LogWarn` + `SetState("LoggedOut")` (Ruling 3); else `CubeCache.AllCubes` sorted by `CubeName`,
    `await CubeDataRepository.InsertCubeDataAsync()`, `LoginUserName = userName`, `SetState("PartialLoggedIn")`. Everything
    caught and logged.
  - `private void LogoutFromXLEdge()`: FinalWorkingCode's `LoggOff()` on AIPowered types — the reset block of AIPowered's own
    `Logout()` (`AppState.Instance.Reset()`, `RibGetCube` label, `Ribledger`/`RibSegS` combos cleared, `SetState("LoggedOut")`)
    moved into `private void ResetLoginState()` that both call. No XLEdge call back, no `applogout`, no message box.
- Modify: `AIPowered\GLSense\CLAUDE.md` — next numbered entry: GLSense AIPowered lacked the two methods 11.1.2 XLEdge calls,
  so XLEdge → GLSense sync failed silently; what was added; Rulings 3 and 4.
- Test (XLEdge suite): `XLEdge.Tests\GLSenseSourceTests.cs` — `GLSenseShell_ExposesTheMethodsXLEdgeCalls`: GLSense's
  `GLSense\AddinModule.cs` contains `public void GetGLCubeInformation(string token, string url, string userName)` and
  `public void LogoutSession()`, and its `AddinEntry.cs` contains `case "LoginFromXLEdge":` and `case "LogoutFromXLEdge":`;
  `Assert.Inconclusive` when the GLSense worktree is not on disk.

- [ ] **Step 1:** Count the baseline: build GLSense.sln Debug and Release with MSBuild, note warnings (`git stash` nothing — the
  worktree is clean apart from FinalWorkingCode's `CLAUDE.md`).
- [ ] **Step 2:** Write `GLSenseSourceTests` → fails.
- [ ] **Step 3:** Implement the GLSense changes and the CLAUDE.md entry.
- [ ] **Step 4:** GLSense Debug + Release build: 0 errors, no warning beyond the baseline; `GLSenseSourceTests` passes.
- [ ] **Step 5:** Commit in the GLSense worktree: `GLSense AIPowered: GetGLCubeInformation and LogoutSession for XLEdge
  (login/logout sync from XLEdge, as FinalWorkingCode)`; commit the XLEdge test.

### Task 4: Close the stage

- [ ] Fresh opus reviewer for the whole stage (both repositories' diffs), one fix pass.
- [ ] `bt.sh` ×3 green; mutants (each must fail a test): drop the `CallGLSense` invoke; `IsGLSenseConnected` always true;
  `ReattachGLSense` skipped; `OnLogoutFromGLSense` empty; `InvokedFromGLSense` without `RibEdgeBroadcast.Enabled = true`;
  `GLSenseAddin.LogoutSession` calling `GetGLCubeInformation`; GLSense `case "LogoutFromXLEdge"` removed.
- [ ] `AIPoweredXLEdge.md`: 2f done, Excel check items added — GLSense + XLEdge in one Excel: login in GLSense → XLEdge logged
  in (pane navigates if open); login in XLEdge → GLSense `PartialLoggedIn` with cubes; logout either side → both logged out,
  one message box; without XLEdge permission; GLSense not loaded.
- [ ] Commit `XLEdge AIPowered Stage 2f: review fixes …` / notes.
