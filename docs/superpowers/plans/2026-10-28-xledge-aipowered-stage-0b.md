# XLEdge AIPowered — Stage 0b (XLEdge.Contracts) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fill `XLEdge.Contracts` with the interfaces and serializable data types the XLEdge shell and the XLEdge core exchange across the AppDomain boundary, mirroring GLSense AIPowered's contract set, sized to what XLEdge actually uses.

**Architecture:** Same five interfaces + two `[Serializable]` DTOs as `GLSense.Contracts`, re-written with XLEdge names. Host → core: `IXLEdgeAddin`. Core reads host services through `IXLEdgeContext` (`ILogger`, `IPathProvider`, `IRibbonController`). Only primitives, `string[]`, `IntPtr`, the two DTOs and MarshalByRef-implemented interfaces cross the boundary. No implementations in this stage.

**Tech Stack:** C# (`latest`), .NET Framework 4.8.1, no external references; checks via Windows PowerShell 5.1 reflection + `BinaryFormatter` round-trip (same by-value path .NET Remoting uses).

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§3 contracts, §4 task pane, §6 deviations)

## Global Constraints

- Namespace `XLEdge.Contracts`; files in `AIPowered/XLEdge/XLEdge.Contracts/`, one type per file, added to `XLEdge.Contracts.csproj` `<Compile>` items.
- No references beyond the framework refs already in the csproj (Stage 0a).
- No GLSense-specific members: no UDF (`ExecuteUdf`, `UdfSentinels`), no configurator (`RelaunchConfiguratorPane`, `ResetConfiguratorPaneReference`, `HasSavedConfigurationSelected`), no combo boxes, no cube (`SetCubeViewBased`), no `Database`/`DrilldownBrowserPath`/`Resources` paths, no `RegisterAction(string, Action)` (a delegate cannot safely cross the domain).
- The only allowed occurrences of "GLSense" are the cross-add-in members `IXLEdgeAddin.OnLoginFromGLSense` and `IXLEdgeContext.GetGLSenseAddinInstance` (XLEdge talks to the GLSense add-in today). The smell check allow-lists exactly these two identifiers.
- Collections crossing the boundary use `string[]` (GLSense uses `IEnumerable<string>`; arrays are guaranteed serializable by value).
- Doc comments: one `<summary>` per type and per non-obvious member, stating the cross-domain rule that applies (what may be passed, what callers must tolerate). Every member added after the first release must be called by the host inside try/catch (Release History can load an older core) — say this once in `IXLEdgeAddin`'s type summary.
- Until the first MSI ships, the contract may still change (Stage 1 spike may reshape the pane members); after that, additive only (spec §6.1).
- Commits: `XLEdge AIPowered Stage 0b: <what>`, no AI attribution, no push.

## Review Focus

1. **A DTO without `[Serializable]` or without a public parameterless constructor** → `SerializationException` at the first cross-domain call. Pinned by the round-trip test (Task 1).
2. **A member signature drifts from what Stages 0c–0e will implement** (name, parameter order, `string[]` vs `IEnumerable`). Pinned by the exact member-list test (Tasks 1–2).
3. **An extra public type or an external reference slips in** (e.g. Excel interop or ADX types in a signature). Pinned by the exact public-type-set and reference checks (Task 2).
4. **`IntPtr` pane handles vs per-window panes**: XLEdge has one task pane per Excel window; the pane members take the Excel window handle. Pinned by the member-list test.
5. **GLSense names leak** beyond the two allow-listed cross-add-in members. Pinned by the smell check (Task 2).

---

## Contract surface (exact)

```
VersionInfo  [Serializable] class: string Version, string ReleaseDate, string DownloadUrl, string Checksum,
             string Notes, bool Mandatory, string FolderName, string FileName   (get; set;)
LoginInfo    [Serializable] class: string LoginUrl, string LoginToken, string LoginUserName, bool IsLoggedIn (get; set;)

ILogger
  void LogInfo(string message)
  void LogWarn(string message)
  void LogError(string message, Exception ex = null)
  void LogDebug(string message)
  void LogException(Exception ex, string context = "")
  void LogRawJson(string context, string rawJson)
  void FlushDebugLogs(string section = "Buffered Logs")
  void BeginReleaseSession(string version, string releaseDate)
  void EndReleaseSession(string version, string releaseDate)
  void LogMethodEntry([CallerMemberName] string methodName = "")
  void LogMethodExit([CallerMemberName] string methodName = "")
  IDisposable BeginLogScope(string scopeName)

IPathProvider
  string Root, Logs, Temp, BrowserDataPath, LoginBrowserPath, UrlsFile            { get; }
  string VersionsPath, SharedDependenciesPath, ManifestFile, ManifestDirectory,
         ReleaseHistoryFile                                                       { get; }
  string LatestVersion, LatestReleaseDate, LatestDownloadUrl, LatestChecksum,
         LatestNotes                                                              { get; }
  bool   LatestMandatory                                                          { get; }
  IReadOnlyList<VersionInfo> AllVersions                                          { get; }
  void Ensure()
  void Refresh()
  void WriteManifest(VersionInfo info)

IRibbonController
  void ExecuteAction(string controlId)
  void SetState(string stateName)
  void SetControlEnabled(string controlId, bool enabled)
  void SetControlVisible(string controlId, bool visible)
  void SetControlPressed(string controlId, bool pressed)
  bool GetControlPressed(string controlId)
  void SetControlLabel(string controlId, string label)
  void EnableControls(string[] controlIds)
  void DisableControls(string[] controlIds)
  void Invalidate(string controlId)
  void InvalidateAll()
  void ShowTaskPane()
  void ToggleTaskPane()
  void HideAllTaskPanes()
  bool IsTaskPaneVisible()

IXLEdgeContext
  object ExcelApp                          { get; }
  ILogger Logger                           { get; }
  IPathProvider Paths                      { get; }
  IRibbonController RibbonController       { get; }
  IntPtr ExcelHandle                       { get; set; }
  bool DebugMode                           { get; set; }
  string Version                           { get; set; }   // loaded core release
  string ReleaseDate                       { get; set; }
  string ActiveFolderName                  { get; set; }   // Versions\{ActiveFolderName}\
  string ShellVersion                      { get; }        // XLEdge.dll (MSI) version, shown in About (spec D7)
  IReadOnlyList<VersionInfo> AllVersions   { get; }
  void SetRibbonController(IRibbonController controller)
  object GetGLSenseAddinInstance()

IXLEdgeAddin
  void Initialize(IXLEdgeContext context)
  void OnRibbonAction(string action, object parameter)
  bool OnExcelEvent(string eventName, object[] args)
  void Shutdown()
  IntPtr CreateTaskPaneContent(IntPtr excelWindowHandle)
  void CloseTaskPaneContent(IntPtr excelWindowHandle)
  void OnLoginFromGLSense(string instanceName, string loginUrl, string token, string userName, bool hasXLEdgePermission)
  LoginInfo GetLoginInfo()
  string[] GetOpenWindowTitles()
  bool IsBusy()
```

`LoginInfo.LoginUserName` is added over GLSense because XLEdge's state and the GLSense login sync carry the user name.

## Parity checklist (GLSense.Contracts → XLEdge.Contracts)

| GLSense | XLEdge | Why |
|---|---|---|
| `IGLSenseAddin` Initialize / OnRibbonAction / OnExcelEvent / Shutdown / GetLoginInfo / GetOpenWindowTitles | same, `IXLEdgeAddin` | parity |
| `ExecuteUdf`, `UdfSentinels` | dropped | no UDFs in XLEdge |
| `CreateConfiguratorPaneContent()` / `CloseConfiguratorPaneContent()` | `CreateTaskPaneContent(IntPtr)` / `CloseTaskPaneContent(IntPtr)` | one pane per Excel window |
| `RelaunchConfiguratorPane`, `ResetConfiguratorPaneReference`, `HasSavedConfigurationSelected` | dropped | configurator-only |
| — | `OnLoginFromGLSense`, `IsBusy` | spec §3 (GLSense→XLEdge login sync; block reload during a report) |
| `IGLSenseContext` (+`EdgeAddinInstance` get/set, `GetEdgeAddinInstance`) | `IXLEdgeContext` (+`ShellVersion`, `GetGLSenseAddinInstance` only) | mirror image of the cross-add-in hook; D7 |
| `IRibbonController` combo / cube / configurator members, `RegisterAction` | dropped; `ShowTaskPane`, `IsTaskPaneVisible` added | XLEdge ribbon has buttons, check boxes, labels only |
| `ILogger` | identical member set | parity |
| `IPathProvider` (`Database`, `DrilldownBrowserPath`, `Resources`, `UrlsDirectory`) | dropped; `BrowserDataPath`, `UrlsFile` | XLEdge folders (`XLEdgeAppPaths`) |
| `LoginInfo`, `VersionInfo` | same (+`LoginUserName`) | parity |

---

### Task 1: Data types and host-service interfaces

**Files:**
- Create: `XLEdge.Contracts/VersionInfo.cs`, `LoginInfo.cs`, `ILogger.cs`, `IPathProvider.cs`, `IRibbonController.cs`
- Modify: `XLEdge.Contracts/XLEdge.Contracts.csproj` (5 `<Compile>` items)
- Test: workspace `verify-0b.ps1`

**Interfaces:**
- Produces: the five types exactly as in "Contract surface"; Task 2's `IXLEdgeContext` / `IXLEdgeAddin` consume them.

- [ ] **Step 1: Write the check** `verify-0b.ps1 -Config <Debug|Release> [-Part Services|All]`. It loads `XLEdge.Contracts\bin\<Config>\XLEdge.Contracts.dll` with `[Reflection.Assembly]::LoadFile` and asserts, for each type in the part: the type exists in `XLEdge.Contracts`; its public members rendered as `ReturnType Name(ParamType name=default, ...)` (properties as `Type Name {get;set;}`) equal the "Contract surface" list exactly (order-insensitive); for `VersionInfo`/`LoginInfo`: `IsSerializable` is true, a public parameterless ctor exists, and a `BinaryFormatter` serialize→deserialize round-trip of an instance with every property set (e.g. `Version="11.1.2.0"`, `Mandatory=$true`, `FolderName="v11.1.2.0_2026-10-28T10-00-00"`, `LoginUserName="qa.user"`, `IsLoggedIn=$true`) returns equal values. Prints `PASS`/`FAIL <reason>` lines, exits non-zero on any failure.
- [ ] **Step 2: Run it** `-Part Services` → expected: FAIL `type XLEdge.Contracts.VersionInfo not found` (and the other four).
- [ ] **Step 3:** Implement the five types per "Contract surface" and Global Constraints (doc comments: DTOs say why they are `[Serializable]`; `ILogger.BeginLogScope` says it returns `IDisposable` so the core never needs the concrete scope type; `IRibbonController` says controls are addressed by their ribbon control id string and that arrays are used because they serialize by value).
- [ ] **Step 4: Build** `XLEdge.sln` Debug and Release → 0 errors, 0 warnings.
- [ ] **Step 5: Run the check** `-Part Services` Debug and Release → all PASS.
- [ ] **Step 6: Commit** `XLEdge AIPowered Stage 0b: contract data types, logger, paths, ribbon controller`.

### Task 2: Context and add-in entry interfaces, full-surface checks

**Files:**
- Create: `XLEdge.Contracts/IXLEdgeContext.cs`, `XLEdge.Contracts/IXLEdgeAddin.cs`
- Modify: `XLEdge.Contracts/XLEdge.Contracts.csproj` (2 `<Compile>` items)
- Modify: workspace `verify-0b.ps1` (part `All`)

**Interfaces:**
- Consumes: Task 1's five types.
- Produces: `IXLEdgeContext`, `IXLEdgeAddin` exactly as in "Contract surface" — Stage 0d's `RemoteLoader` returns `IXLEdgeAddin`; Stage 0e's shell implements `IXLEdgeContext`/`IRibbonController`; Stage 0c implements `ILogger`/`IPathProvider`.

- [ ] **Step 1: Extend the check** with `-Part All`: the Task 1 assertions plus member lists for `IXLEdgeContext` and `IXLEdgeAddin`; the assembly's exported public types are exactly `{VersionInfo, LoginInfo, ILogger, IPathProvider, IRibbonController, IXLEdgeContext, IXLEdgeAddin}`; `GetReferencedAssemblies()` names are a subset of `{mscorlib, System, System.Core}`; no public member signature mentions a type outside `System.*` / `XLEdge.Contracts.*`.
- [ ] **Step 2: Run it** `-Part All` → expected: FAIL `type XLEdge.Contracts.IXLEdgeContext not found`, `type XLEdge.Contracts.IXLEdgeAddin not found` (Task 1 parts PASS).
- [ ] **Step 3:** Implement both interfaces. `IXLEdgeAddin` type summary: entry point the shell calls in the core's AppDomain; args are primitives/strings/`IntPtr` only (never ADX event args, never live `Range`/`Worksheet`); calls return quickly and the core catches every exception; members added after the first release must be called inside try/catch tolerating `MissingMethodException`/`RemotingException`. `OnExcelEvent` lists the five XLEdge events (SheetSelectionChange, SheetActivate, WorkbookActivate, SheetFollowHyperlink, SheetBeforeDelete) and that `false` means "cancel" where the event has a cancel flag. `CreateTaskPaneContent` says it returns the HWND of a borderless core window the shell re-parents into the ADX pane of that Excel window, idempotent per window. `IsBusy` says true while a report or refresh is running so the shell blocks a reload.
- [ ] **Step 4: Build** Debug and Release → 0 errors, 0 warnings; the rest of `XLEdge.sln` still builds.
- [ ] **Step 5: Run the check** `-Part All` Debug and Release → all PASS.
- [ ] **Step 6: Smell check** `grep -rniE "glsense" AIPowered/XLEdge --exclude-dir=obj --exclude-dir=bin` → only `OnLoginFromGLSense` and `GetGLSenseAddinInstance` lines (and their doc comments); `grep -rnE "\bGL[A-Z]" … --include=*.cs` → only those same identifiers.
- [ ] **Step 7: Commit** `XLEdge AIPowered Stage 0b: IXLEdgeContext and IXLEdgeAddin` (+ this plan file).
