# XLEdge AIPowered — Stage 0c (XLEdge.Shared + XLEdge.Tests) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement `XLEdge.Shared` (PathProvider, Logger, ReleaseEntry, ReleaseHistoryStore, VersionParser) as the XLEdge twin of `GLSense.Shared`, with MSTest unit tests in a new `XLEdge.Tests` project, and apply the contract decisions taken after Stage 0b.

**Architecture:** Shared runs in the shell's AppDomain; `PathProvider` and `Logger` are MarshalByRefObjects (infinite lease) that the core reaches through remoting proxies; `ReleaseHistoryStore` and `VersionParser` are used by the loader (0d) and the shell's reload windows (0e). Behaviour follows GLSense.Shared (values below), minus its dead members and with its known bugs fixed (listed per task).

**Tech Stack:** .NET Framework 4.8.1, C# latest, NLog 6.1.3, System.Text.Json 10.0.9 (XLEdge's current version), MSTest 3.11.1 (`MSTest.TestFramework` + `MSTest.TestAdapter`, PackageReference), `vstest.console.exe` from VS 2022.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§3 contract version, §6.1, Tests)

## Global Constraints

- Shared packages via `XLEdge.Shared\packages.config` restored into `AIPowered\XLEdge\packages\`: NLog 6.1.3, System.Text.Json 10.0.9, Microsoft.Bcl.AsyncInterfaces 10.0.9, System.IO.Pipelines 10.0.9, System.Text.Encodings.Web 10.0.9, System.Buffers 4.6.1, System.Memory 4.6.3, System.Numerics.Vectors 4.6.1, System.Runtime.CompilerServices.Unsafe 6.1.2, System.Threading.Tasks.Extensions 4.6.3, System.ValueTuple 4.6.2.
- The shell csproj references every one of those DLLs directly with `Private=True` (GLSense CLAUDE.md §46: a ProjectReference does not copy packages.config dependencies; missing System.Text.Json crashed Excel at startup).
- `XLEdge.Tests`: classic csproj, net481, not signed, PackageReference MSTest 3.11.1 + NLog 6.1.3 + System.Text.Json 10.0.9, ProjectReferences Contracts and Shared. No shipped project references it. Tests use a fresh temp folder per test and delete it afterwards; no test touches `%LOCALAPPDATA%\ORBIT`.
- Test command (the suite): `"C:/Program Files/Microsoft Visual Studio/2022/Community/Common7/IDE/Extensions/TestPlatform/vstest.console.exe" AIPowered/XLEdge/XLEdge.Tests/bin/Debug/XLEdge.Tests.dll` → expected `Passed!` with 0 failed.
- Every MarshalByRefObject in Shared overrides `InitializeLifetimeService()` to return `null`.
- No static singletons for paths or logger state that crosses instances (GLSense's `PathProvider.Instance` / static Latest* are not repeated).
- Names: product text "Orbit XLEdge", files `XLEdge_Logs_…`, `XLEdge_Error_…`, mutex prefix `XLEdge_ReleaseHistory_`. Smell check as in 0b (allow-list: `OnLoginFromGLSense`, `GetGLSenseAddinInstance`, the `UrlsFile` "shared with GLSense" comment).
- Build: `XLEdge.sln` Debug and Release, 0 errors, 0 warnings. Commits `XLEdge AIPowered Stage 0c: <what>`, no AI attribution, no push.

## Review Focus

1. **Non-English Windows culture** (e.g. de-DE) → GLSense builds the log file name and its "find today's file" glob with different cultures, so the release-session de-duplication silently fails. Expected: file name always `XLEdge_Logs_28-Oct-2026.log` (English month). Test in Task 5.
2. **Corrupt `ReleaseHistory.json`** → must surface as an exception to the caller (0d decides), never be treated as "no history" (that path wipes `Versions\` on a fresh install). Test in Task 3.
3. **Mutex held by another process / unusable** → `TimeoutException` with a clear message, never an unprotected read/write; abandoned mutex counts as acquired. Test in Task 3.
4. **Missing `manifest.json`** → Latest* are null and the file is **not** recreated with defaults (GLSense §40.2 bug). Test in Task 2.
5. **Logging an exception whose type exists only in the core** → impossible by contract now (text only); `LogException(null, …)` / `LogError(msg, null)` must not throw. Test in Task 5.

---

## Task 1: XLEdge.Tests project, contract version, contract clean-up

**Files:**
- Create: `XLEdge.Contracts/ContractInfo.cs`; `XLEdge.Tests/XLEdge.Tests.csproj`, `XLEdge.Tests/Properties/AssemblyInfo.cs`, `XLEdge.Tests/ContractSurfaceTests.cs`
- Modify: `XLEdge.Contracts/VersionInfo.cs`, `IPathProvider.cs`, `IXLEdgeContext.cs`, `ILogger.cs`, `IRibbonController.cs`, `IXLEdgeAddin.cs`, `XLEdge.Contracts.csproj`; `XLEdge.sln` (add XLEdge.Tests)

**Interfaces:**
- Produces: `public static class ContractInfo { public const int Version = 1; }`; `VersionInfo.ContractVersion` (`int`, JSON `contractVersion`, 0 = unknown); trimmed `IPathProvider` = `Root, Logs, Temp, BrowserDataPath, UrlsFile, VersionsPath, SharedDependenciesPath, ManifestFile, ManifestDirectory, ReleaseHistoryFile, LatestVersion, LatestReleaseDate, Ensure(), Refresh()`; `IXLEdgeContext` without `AllVersions`; `ILogger` without `LogMethodEntry/LogMethodExit`.

Contract clean-up (my call, before the first MSI freezes it): drop GLSense's dead members — `IPathProvider.LatestDownloadUrl/LatestChecksum/LatestNotes/LatestMandatory/AllVersions/WriteManifest/LoginBrowserPath`, `IXLEdgeContext.AllVersions`, `ILogger.LogMethodEntry/LogMethodExit` (0 uses in GLSense, no XLEdge use). Doc fixes from the 0b review: pane methods act on the active Excel window's pane; `OnExcelEvent` return value is reserved (none of the five events can be cancelled).

- [ ] **Step 1: Test** `ContractSurfaceTests` (ports verify-0b, stricter): per contract type an exact, **case-sensitive** member list (methods, properties, events, fields, implemented interfaces) matching the Produces block and the 0b surface; `IsInterface` for the six interfaces, `IsClass` for the DTOs, `IsAbstract && IsSealed` for `ContractInfo`; `ContractInfo.Version == 1`; DTO BinaryFormatter round-trip incl. `ContractVersion = 1`; exported types exactly `{ContractInfo, ILogger, IPathProvider, IRibbonController, IXLEdgeAddin, IXLEdgeContext, LoginInfo, VersionInfo}`; references `{mscorlib}` only.
- [ ] **Step 2:** Create `XLEdge.Tests` (csproj per Global Constraints), restore (`MSBuild XLEdge.sln -t:Restore`), build, run the suite → expected FAIL (ContractInfo missing, dead members still present).
- [ ] **Step 3:** Apply the contract changes.
- [ ] **Step 4:** Build Debug + Release (0/0), run the suite → all pass.
- [ ] **Step 5: Commit** `XLEdge AIPowered Stage 0c: XLEdge.Tests, contract version, contract clean-up`.

## Task 2: Packages and PathProvider

**Files:**
- Create: `XLEdge.Shared/packages.config`, `XLEdge.Shared/PathProvider.cs`, `XLEdge.Tests/PathProviderTests.cs`
- Modify: `XLEdge.Shared/XLEdge.Shared.csproj` (package refs, Compile), `XLEdge/XLEdge.csproj` (direct package refs, §46)

**Interfaces:**
- Produces: `public class PathProvider : MarshalByRefObject, IPathProvider` with `PathProvider(string installRoot)` (data base = `%LOCALAPPDATA%\ORBIT\Excel_Logs`) and `PathProvider(string installRoot, string dataBasePath)`.

Values: `Root = {base}\XLEdge_Logs`, `Logs = Root\Logs`, `Temp = Root\Temp`, `BrowserDataPath = Root\BrowserLogs`, `UrlsFile = {base}\ORBIT_URLS.xml` (all as today's `XLEdgeAppPaths`); `VersionsPath = {installRoot}\AddinCore\Versions`, `SharedDependenciesPath = …\AddinCore\SharedDependencies`, `ManifestDirectory = …\AddinCore\Manifest`, `ManifestFile = ManifestDirectory\manifest.json`, `ReleaseHistoryFile = …\AddinCore\ReleaseHistory.json`. Constructor only computes paths (no IO). `Ensure()` creates Logs, Temp, BrowserDataPath, Versions, SharedDependencies, Manifest. `Refresh()` reads `manifest.json` (JSON array of VersionInfo, case-insensitive) and sets `LatestVersion`/`LatestReleaseDate` from the entry with the highest `System.Version`; entries whose version does not parse are skipped. Missing, empty, corrupt or object-not-array → both null; never creates the file. Null/blank `installRoot` → `ArgumentException`.

- [ ] **Step 1: Tests** `PathProviderTests`: `Paths_FollowXLEdgeLayout` (all 10 paths for installRoot `T\inst`, base `T\data`); `Ctor_DoesNoIo`; `Ensure_CreatesFolders`; `Refresh_PicksHighestVersion` (`[{"version":"11.1.1","releaseDate":"a"},{"version":"11.1.10","releaseDate":"b"},{"version":"bad"}]` → `11.1.10`/`b`); `Refresh_MissingManifest_NullsAndDoesNotCreate`; `Refresh_CorruptOrObject_Nulls`; `Ctor_BlankInstallRoot_Throws`; `Lease_IsInfinite` (`InitializeLifetimeService()` is null).
- [ ] **Step 2:** Run suite → FAIL (PathProvider missing).
- [ ] **Step 3:** Add packages (packages.config + HintPaths `..\packages\…`), shell direct refs, implement PathProvider.
- [ ] **Step 4:** Build 0/0, suite green; shell `bin\Debug` contains NLog.dll, System.Text.Json.dll and its 8 dependencies.
- [ ] **Step 5: Commit** `XLEdge AIPowered Stage 0c: packages and PathProvider`.

## Task 3: ReleaseEntry and ReleaseHistoryStore

**Files:**
- Create: `XLEdge.Shared/ReleaseEntry.cs`, `XLEdge.Shared/ReleaseHistoryStore.cs`, `XLEdge.Shared/Properties/AssemblyInfo.cs` (add `InternalsVisibleTo("XLEdge.Tests, PublicKey=…")` — the test assembly is unsigned, so use `[assembly: InternalsVisibleTo("XLEdge.Tests")]` only if the compiler accepts it for a signed assembly; otherwise expose the timeout overload as public and note it), `XLEdge.Tests/ReleaseHistoryStoreTests.cs`

**Interfaces:**
- Produces: `[Serializable] public class ReleaseEntry { string Version, ReleaseDate, FolderName, Checksum, Notes, Source; int ContractVersion; }` (Source: `"Install"` default, `"Online"`, `"Offline"`); `public enum ReleaseDeleteResult { Deleted, NotFound, ActiveRelease, InvalidFolderName, FolderDeleteFailed, CatalogWriteFailed }`; `public static class ReleaseHistoryStore` with `List<ReleaseEntry> ReadAll(string file)`, `void Append(string file, ReleaseEntry entry)`, `List<ReleaseEntry> Reconcile(string file, string versionsPath)`, `ReleaseDeleteResult Delete(string file, string versionsPath, string folderName, string activeFolderName, out Exception error)`, `string BuildFolderName(string version, string releaseDate)`, `bool IsSafeFolderName(string folderName)`.

Behaviour (GLSense values): JSON array, WriteIndented, PascalCase on write, case-insensitive read. Missing / empty / whitespace / `null` → empty list; corrupt → `JsonException` propagates. Atomic write: `{file}.tmp` then `File.Replace` (target exists) or `File.Move`. Append dedupes on Version+ReleaseDate+FolderName (ordinal), no ordering. Reconcile keeps entries whose `versionsPath\FolderName` exists and holds ≥1 top-level `*.dll`, rewrites only if the count changed. Delete: order and results exactly as GLSense (find by FolderName ignore-case → NotFound; active → ActiveRelease; unsafe name or full path outside versions root → InvalidFolderName; folder delete failure → FolderDeleteFailed; catalog write failure → CatalogWriteFailed). `BuildFolderName`: invalid file-name chars of the date → `-`, `"v{version}_{safeDate}"`. Lock: named mutex **`Global\XLEdge_ReleaseHistory_{hash}`**, hash = first 16 hex of SHA-256 of the lower-cased full file path (fixes GLSense's single global name, which collides across Windows users and throws UnauthorizedAccess); 10 s timeout; abandoned = acquired; not acquired or `UnauthorizedAccessException` → `TimeoutException("Could not acquire the release history lock for '{file}' within 00:00:10 - refusing to read/write it unprotected.")`; released in `finally`.

- [ ] **Step 1: Tests**: `Read_MissingEmptyNull_ReturnsEmpty`; `Read_Corrupt_ThrowsJsonException`; `Append_WritesPascalCaseAndDedupes`; `Write_FirstTimeUsesMove_LeavesNoTmp`; `Reconcile_DropsMissingFolderAndFolderWithoutDll_RewritesOnlyOnChange`; `Delete_Results` (one assertion per result value; `..`, `a\b`, `C:\x` → InvalidFolderName; Deleted removes folder and entry); `BuildFolderName_ReplacesInvalidChars` (`("11.1.2","2026-10-28T14:05:09")` → `v11.1.2_2026-10-28T14-05-09`); `Lock_HeldElsewhere_TimesOut` (hold the same-named mutex on another thread; call with a 200 ms test timeout); `Lock_Abandoned_CountsAsAcquired`; `Lock_DifferentFiles_DoNotContend`.
- [ ] **Step 2:** Run suite → FAIL.
- [ ] **Step 3:** Implement (the timeout overload used by tests: `internal` + `InternalsVisibleTo`, or `public` — rule at implementation time and ledger it).
- [ ] **Step 4:** Build 0/0, suite green.
- [ ] **Step 5: Commit** `XLEdge AIPowered Stage 0c: ReleaseEntry and ReleaseHistoryStore`.

## Task 4: VersionParser

**Files:**
- Create: `XLEdge.Shared/VersionParser.cs`, `XLEdge.Tests/VersionParserTests.cs`

**Interfaces:**
- Produces: `public class VersionParseResult { bool Success; string Version; DateTime ReleaseDate; int ContractVersion; List<VersionInfo> AllVersions; string ErrorMessage; }`; `public class VersionParser` with `VersionParser(ILogger logger = null)`, `VersionParseResult ParseVersionJson(string json)`, `VersionParseResult ParseVersionFile(string path)`, `VersionInfo GetLatestVersion(IEnumerable<VersionInfo> versions)`. (GLSense's `IVersionParser`, `ParseVersionFromDirectory`, DownloadUrl/Checksum/Notes/Mandatory result fields are dropped — unused.)

Behaviour: JSON array only, options case-insensitive + camelCase; never throws; error messages exactly GLSense's (`"JSON content is empty or null"`, `"No version data found in JSON"`, `"Failed to determine latest version"`, `"Invalid JSON format: {msg}"`, `"Failed to parse version JSON: {msg}"`, `"Manifest file not found: {path}"`, `"Failed to read manifest file: {msg}"`), each also sent to `logger?.LogError`. Result fields come from the latest entry (highest `System.Version`, ties → first). Release date: `DateTime.TryParseExact` with `yyyy-MM-ddTHH:mm:ss`, `dd-MMM-yyyy`, `yyyy-MM-dd` (InvariantCulture) — **no current-culture fallback** (GLSense relied on it; ambiguous on non-English machines) — else `DateTime.MinValue` + `LogWarn`.

- [ ] **Step 1: Tests**: `Parse_Array_TakesLatest_WithContractVersion`; `Parse_Errors` (blank, `[]`, `{}`, `[{"version":"x"}]`, garbage) with the exact messages; `Parse_ReleaseDate_InvariantUnderDeDE` (set `CurrentCulture = de-DE`, `"2026-10-28T14:05:09"` parses to 28 Oct 2026 14:05:09); `ParseFile_Missing`.
- [ ] **Step 2:** Run suite → FAIL. **Step 3:** Implement. **Step 4:** Build 0/0, suite green.
- [ ] **Step 5: Commit** `XLEdge AIPowered Stage 0c: VersionParser`.

## Task 5: Logger

**Files:**
- Create: `XLEdge.Shared/Logger.cs`, `XLEdge.Tests/LoggerTests.cs`, `XLEdge.Tests/FakeContext.cs` (test-only `IXLEdgeContext` with settable `DebugMode`, `Paths`, `ExcelApp = null`)

**Interfaces:**
- Consumes: `IXLEdgeContext` (`DebugMode`, `Paths.Logs`, `ExcelApp`), `ILogger` (Task 1).
- Produces: `public sealed class Logger : MarshalByRefObject, ILogger, IDisposable` with `Logger(IXLEdgeContext context)`; nested `public sealed class LogScope : MarshalByRefObject, IDisposable`. `Dispose()` flushes open buffers and shuts down its private NLog `LogFactory` (shell calls it at AddinBeginShutdown).

Behaviour (GLSense values unless marked **fix**):
- **fix:** own `NLog.LogFactory` instance per Logger (not the process-global `LogManager`), path from `context.Paths.Logs` (no singleton).
- File `{Logs}\XLEdge_Logs_${date:format=dd-MMM-yyyy:culture=en-US}.log`; C# side builds the same name with `CultureInfo.GetCultureInfo("en-US")` (**fix**, Review Focus 1). FileTarget: layout `${longdate}|${message:withException=true:exceptionSeparator=|}`, AutoFlush, KeepFileOpen, ArchiveAboveSize 20 MB, MaxArchiveFiles 30, ArchiveSuffixFormat `({0})`, no ArchiveFileName, Header = environment snapshot (`===== Environment Snapshot =====`, `Product: Orbit XLEdge`, time zone, Excel version via reflection on `ExcelApp` or `unknown`, OS, .NET runtime, screen DPI, culture, machine/user, `=================================`). Rule Debug→Fatal.
- Lines `{indent}INFO  | HH:mm:ss | msg`, `WARN  |`, `ERROR |`, `DEBUG |`; indent = 2 spaces × scope depth. INFO immediate; WARN immediate + flush buffer (`"{root} - warning logged"`); ERROR immediate (+ `|{exceptionDetails}` when given) + flush (`"{root} - error logged"`); DEBUG only when `DebugMode`, buffered inside a scope else immediate.
- `LogException(details, context)`: block `========== Exception ==========`, `Context: …` (if any), the details text (or `<no details>`), `============================`, immediate + flush (`"{root} - exception logged"`); null details never throws (**fix**).
- `LogRawJson(context, raw)`: `----- Raw JSON ({context}) -----`, raw or `<empty>`, `----- End Raw JSON -----`, immediate.
- Scopes: AsyncLocal buffer, created only when DebugMode is on and no buffer is current; nested scopes reuse it; BEGIN/END lines; flush on owning scope dispose as one event `===== {label} | yyyy-MM-dd HH:mm:ss =====` + dashes; safety-net timer (15 s period) flushes buffers older than 30 s (`"{root} - safety-net flush, action still running"`); depth changes under the buffer lock (**fix**). `FlushDebugLogs(section)` flushes all open buffers (`"{root} - {section}"`).
- Release sessions: `===== Start of XLEdge version {v} (released {d}) logs as on {dddd, dd MMMM yyyy} =====` (en-US), skipped when today's newest `XLEdge_Logs_*.log` already ends its latest section with the same Start marker; `EndReleaseSession` flushes (`"release session ending"`) then writes `===== End of XLEdge version {v} (released {d}) logs =====`.
- Init failure: `{Logs}\XLEdge_Error_{dd-MMM-yyyy en-US}.log`, then `%TEMP%\XLEdge_Error_{yyyyMMdd_HHmmss}.log`; no never-drained fallback buffer (**fix**).

- [ ] **Step 1: Tests** (read the log with `FileShare.ReadWrite`): `FileName_EnglishMonthUnderDeDE`; `Header_HasProductAndSnapshot`; `Info_LineFormat` (regex `^\S+ \S+\|INFO  \| \d\d:\d\d:\d\d \| hello$`); `Debug_DroppedWhenDebugOff`; `Scope_BuffersDebugAndFlushesOneBlockOnDispose`; `Scope_NestedReusesBuffer`; `Warn_InsideScope_FlushesWithReason`; `Error_AppendsDetails`; `LogException_NullDetails_DoesNotThrow`; `RawJson_Format`; `ReleaseSession_BeginNotDuplicated_EndFlushes`; `Lease_IsInfinite` (Logger and LogScope); `Dispose_ReleasesFile` (file can be deleted after Dispose).
- [ ] **Step 2:** Run suite → FAIL. **Step 3:** Implement. **Step 4:** Build Debug + Release 0/0, full suite green.
- [ ] **Step 5: Smell check** (0b rules) → only allow-listed hits.
- [ ] **Step 6: Commit** `XLEdge AIPowered Stage 0c: Logger` (+ this plan, spec updates).
