# XLEdge AIPowered — Stage 0d (XLEdge.Loader.Core) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement `XLEdge.Loader.Core` — load the XLEdge core into its own AppDomain from `AddinCore\Versions\<release>\`, unload it, and decide/extract/catalog which release to load — as the XLEdge twin of `GLSense.Loader.Core`, proven by tests that load and unload real AppDomains.

**Architecture:** `AddinDomainLoader` (shell domain) creates `XLEdgeCore_<folder>` with the release folder as ApplicationBase, creates `RemoteLoader` inside it, which loads `XLEdge.Addin.Core.dll`, instantiates its `IXLEdgeAddin` entry type and hands back a `RemoteAddinProxy` (a Loader.Core type, so the shell never has to resolve the core assembly). `UpdateBootstrapper` picks the release (fresh-install seed → staged manifest+zip → newest catalogued release) and extracts/catalogues zips. A minimal no-op `AddinEntry` is added to `XLEdge.Addin.Core` now (it was planned for 0e) so the loader is tested against the real core assembly.

**Tech Stack:** .NET Framework 4.8.1 AppDomains + .NET Remoting, System.IO.Compression, XLEdge.Shared (ReleaseHistoryStore, VersionParser, PathProvider), MSTest 3.11.1.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§3 contract version, §6, §8 risks)

## Global Constraints

- All rules from Stage 0c (packages, test command, 0/0 build in Debug and Release, smell check, commits `XLEdge AIPowered Stage 0d: <what>`, no AI attribution, no push).
- **Every bug or warning found while working or in review is fixed in this stage with a test — nothing is deferred** (developer rule, 2026-10-28).
- Every MarshalByRefObject returns `null` from `InitializeLifetimeService()`.
- Loader never lets an exception escape `Unload`; `Load` rethrows after unloading any half-created domain.
- Log lines through `context.Logger` (null-safe), prefixed with the class name like GLSense (`AddinDomainLoader.Load: …`).
- Contract rule (spec §3): a release loads/extracts only if `1 <= contractVersion <= ContractInfo.Version`; `0` (unknown) is refused with "built without a contract version".

## Deliberate differences from GLSense (my calls — veto any)

1. **No shadow copying.** Release folders are immutable (a release is never re-extracted into an existing folder, and the active release cannot be deleted), so shadow copies buy nothing; GLSense's `CachePath` (without `ApplicationName`) is unproven and put a `ShadowCache` inside the release folder that its shared-dependency walk would then pick up. A test proves the folder is unlocked after `Unload`.
2. **No `PrivateBinPath`/`PrivateBinPathProbe`.** Absolute paths outside ApplicationBase are ignored by the runtime and a non-null `PrivateBinPathProbe` drops ApplicationBase from probing; SharedDependencies is reached through the resolver only (as it effectively is in GLSense).
3. **`ConfigurationFile`** = `<release>\XLEdge.Addin.Core.dll.config` when present (binding redirects of the core's own app.config apply in its domain; GLSense sets none).
4. **Domain name** `XLEdgeCore_<folder>` (GLSense uses the version, so a same-version reload reuses the name).
5. **Load failure unloads the half-created domain** (GLSense leaves it until the next Unload).
6. **Resolver**: exact full-name match among loaded assemblies first, then simple-name; on a failing candidate it tries the next one (GLSense returns null).
7. **Extraction is atomic**: into `Versions\<folder>.partial-<guid>` then `Directory.Move` to `Versions\<folder>` (GLSense deletes/retries and may extract into a dirty folder).
8. **Checksum required** (GLSense skips verification when the manifest has none).
9. **Seed failure keeps the seed** (GLSense deletes `Manifest\` even when extraction failed).
10. **Several zips staged**: use the manifest's `fileName`; otherwise the newest zip with a warning (GLSense takes an arbitrary first).
11. **Corrupt or empty catalog** (spec Review Focus from 0c): renamed to `ReleaseHistory.corrupt-<yyyyMMddHHmmss>.json` and rebuilt from the `manifest.json` snapshots inside `Versions\*\`; never treated as a fresh install (only a *missing* catalog is).

## Review Focus

1. **Core types leaking into the shell domain** (casting the proxy forces the shell to load `XLEdge.Addin.Core.dll`) → test asserts the shell domain never loads that assembly.
2. **Unload leaves the release folder locked** (shadow copy dropped) → test deletes the folder after Unload.
3. **A fresh install path that wipes `Versions\` when the catalog is corrupt/empty** → tests for both.
4. **Contract version newer than the shell / 0** → refused before any domain is created or any file is extracted.
5. **Reload into the same or another release in one session** (domain names, resolver statics per domain) → test loads A, unloads, loads B, unloads, loads A again.

---

## File Structure

```
XLEdge.Loader.Core\
  ResolvedRelease.cs        Version, ReleaseDate, FolderName, ContractVersion
  ContractCompatibility.cs  IsLoadable(int), Describe(int) → refusal text
  RemoteLoader.cs           RemoteLoader (+ nested RemoteAddinProxy), AssemblyResolve handler
  AddinDomainLoader.cs      Load(context, release) / Unload(context)
  UpdateBootstrapper.cs     ResolveVersionToLoad, ExtractAndCatalog, catalog recovery, shared dependencies
XLEdge.Addin.Core\
  AddinEntry.cs             minimal IXLEdgeAddin (no-op), MarshalByRefObject
XLEdge.Tests\
  RecordingLogger.cs        test-only MarshalByRef ILogger collecting lines
  ReleaseFolders.cs         stages a real release folder from XLEdge.Addin.Core\bin\<Config>; builds zips + manifests
  AddinDomainLoaderTests.cs, UpdateBootstrapperTests.cs
```

`XLEdge.Tests.csproj` gets a ProjectReference to `XLEdge.Addin.Core` (`ReferenceOutputAssembly=false`, build order only) and to `XLEdge.Loader.Core`.

---

### Task 1: ResolvedRelease, contract compatibility, minimal AddinEntry, test helpers

**Interfaces — Produces:**
- `public class ResolvedRelease { string Version; string ReleaseDate; string FolderName; int ContractVersion; }`
- `public static class ContractCompatibility { bool IsLoadable(int contractVersion); string RefusalReason(int contractVersion); }` — `RefusalReason` returns null when loadable, else `"built without a contract version"` (0) or `"needs a newer XLEdge installer (contract {v}, this installer supports up to {ContractInfo.Version})"`.
- `XLEdge.Addin.Core.AddinEntry : MarshalByRefObject, IXLEdgeAddin` — `Initialize` stores the context and logs `AddinEntry.Initialize: XLEdge core initialized in AppDomain '{FriendlyName}' from '{BaseDirectory}'.`; `Shutdown` logs `AddinEntry.Shutdown: XLEdge core shut down.`; `GetLoginInfo` → `new LoginInfo()`; `GetOpenWindowTitles` → empty; `IsBusy` → false; `CreateTaskPaneContent` → `IntPtr.Zero`; other members no-op; every member catches and logs.
- Test helpers: `RecordingLogger` (`Lines` snapshot, `Contains(string)`); `ReleaseFolders.Stage(string versionsPath, string folderName)` copies every `*.dll` (+ `.config`) from `XLEdge.Addin.Core\bin\<Config>\`; `ReleaseFolders.BuildZip(...)` / `Manifest(...)` build a zip and a one-entry manifest JSON with SHA-256 checksum.

- [ ] **Step 1: Tests** `ContractCompatibility_Rules` (0 refused with the "without" text, 1 loadable, `ContractInfo.Version + 1` refused with the installer text); `AddinEntry_IsMarshalByRefWithInfiniteLease`; `AddinEntry_InitializeLogsDomain` (in-domain, RecordingLogger).
- [ ] **Step 2:** Run → RED. **Step 3:** Implement. **Step 4:** Build 0/0, suite green. **Step 5: Commit** `XLEdge AIPowered Stage 0d: ResolvedRelease, contract compatibility, minimal AddinEntry`.

### Task 2: RemoteLoader and AddinDomainLoader

**Interfaces — Produces:**
- `public class AddinDomainLoader` — `IXLEdgeAddin Load(IXLEdgeContext context, ResolvedRelease release)`; `void Unload(IXLEdgeContext context)`; `string CurrentDomainName { get; }` (null when nothing loaded).
- `public class RemoteLoader : MarshalByRefObject` — `IXLEdgeAddin Create(string releaseFolder, string sharedDependenciesPath, ILogger logger)`; nested `public sealed class RemoteAddinProxy : MarshalByRefObject, IXLEdgeAddin` forwarding all 12 members.

Behaviour: `Load` refuses (throws `InvalidOperationException` with `ContractCompatibility.RefusalReason`) before creating anything; release folder `Paths.VersionsPath\release.FolderName` must contain `XLEdge.Addin.Core.dll` (`FileNotFoundException`); AppDomainSetup = ApplicationBase only (+ ConfigurationFile if present, + `ApplicationName = "XLEdgeCore"`); `CreateInstanceAndUnwrap(typeof(RemoteLoader).Assembly.FullName, typeof(RemoteLoader).FullName)`; `Create` registers the resolver once per domain, `Assembly.LoadFrom(<folder>\XLEdge.Addin.Core.dll)`, entry type = first concrete `IXLEdgeAddin` (handles `ReflectionTypeLoadException` by logging its loader exceptions), `Activator.CreateInstance`, wraps in the proxy; `Load` then calls `Initialize(context)` and returns the proxy. Any failure: log, unload the domain, rethrow. Resolver: skip `*.resources`; exact full-name loaded match, then simple-name loaded match; then `<release>\{name}.dll`, `<shared>\{name}.dll`, `<release>\{name}.exe`, `<shared>\{name}.exe`, trying the next candidate if `LoadFrom` throws; else warn and null. `Unload`: null-safe context; take and clear the domain; `Task.Run(AppDomain.Unload)` waited 5 s; timeout → `LogError` and return; any exception (incl. `AggregateException`) → `LogError`, never rethrown.

- [ ] **Step 1: Tests** (real domains, `ReleaseFolders.Stage`): `Load_RunsCoreInItsOwnDomain` (log has `AppDomain 'XLEdgeCore_v1'` and the release path; `GetLoginInfo()` works through the proxy; `AppDomain.CurrentDomain.GetAssemblies()` has no `XLEdge.Addin.Core`); `Load_NewerOrUnknownContract_RefusedWithoutCreatingDomain` (no `XLEdgeCore_` log line, `CurrentDomainName` null); `Load_MissingCoreDll_ThrowsAndLeavesNothingLoaded`; `Unload_ReleasesTheReleaseFolder` (Directory.Delete succeeds); `Reload_AThenBThenA`; `Unload_WhenNothingLoaded_IsSafe` (also with null context); `Resolver_FindsDependencyInSharedDependencies` (move `XLEdge.Shared.dll` from the release folder to SharedDependencies — loads and logs fine); `Proxy_IsMarshalByRefWithInfiniteLease`.
- [ ] **Step 2:** Run → RED. **Step 3:** Implement. **Step 4:** Build 0/0, suite green ×3 (domain tests must be stable). **Step 5: Commit** `XLEdge AIPowered Stage 0d: RemoteLoader and AddinDomainLoader`.

### Task 3: UpdateBootstrapper.ExtractAndCatalog and shared dependencies

**Interfaces — Produces:** `public class UpdateBootstrapper` — `ResolvedRelease ExtractAndCatalog(IXLEdgeContext context, string manifestJson, byte[] zipBytes, string source)` (null on refusal, logged; I/O exceptions other than refusals propagate).

Behaviour: parse with `VersionParser` (first entry of the array); blank version → null; folder = `FolderName` or `ReleaseHistoryStore.BuildFolderName`; refuse unsafe name (`IsSafeFolderName`), `== context.ActiveFolderName`, blank checksum, checksum mismatch (SHA-256 uppercase hex, compared ignore-case), contract not loadable. Folder already holding DLLs → skip extraction (still dedupe shared deps). Else extract via temp zip file into `<folder>.partial-<guid>`, write the raw manifest snapshot `manifest.json`, `Directory.Move` to the final name; on any failure delete the partial folder and rethrow. Shared dependencies (file name match anywhere under the release, kept relative path): `MahApps.Metro.IconPacks.Core.dll`, `MahApps.Metro.IconPacks.FontAwesome.dll`, `Microsoft.Web.WebView2.Core.dll`, `Microsoft.Web.WebView2.Wpf.dll`, `Microsoft.Web.WebView2.WinForms.dll`, `WebView2Loader.dll` (native, any subfolder), `Wpf.Ui.dll`, `Wpf.Ui.Abstractions.dll`, `Microsoft.Xaml.Behaviors.dll`, `NLog.dll`, `System.Text.Json.dll`, `System.Text.Encodings.Web.dll`, `System.Buffers.dll`, `System.Memory.dll`, `System.IO.Pipelines.dll`, `System.Numerics.Vectors.dll`, `System.Runtime.CompilerServices.Unsafe.dll`, `System.Threading.Tasks.Extensions.dll`, `Microsoft.Bcl.AsyncInterfaces.dll`, `System.ValueTuple.dll` — move to SharedDependencies when absent there, delete the local copy when SHA-256-identical, keep the local copy (warn) when different; then remove empty folders. Contracts / Shared / Loader.Core / Addin.Core are never shared. Append `ReleaseEntry` (incl. `ContractVersion`, `Source`).

- [ ] **Step 1: Tests**: `Extract_CataloguesAndSnapshotsManifest`; `Extract_Refusals` (one assertion each: `..`, `v1.`, active folder, blank checksum, wrong checksum, contract 0, contract too new — each leaves `Versions\` and the catalog untouched); `Extract_AlreadyExtracted_IsIdempotent`; `Extract_BadZip_LeavesNoPartialFolder`; `SharedDependencies_MovedDedupedOrKept` (two releases: identical `NLog.dll` → one shared copy, no local copies; different `NLog.dll` → second keeps its own + warning; `XLEdge.Shared.dll` never moved).
- [ ] **Step 2:** RED. **Step 3:** Implement. **Step 4:** 0/0, green. **Step 5: Commit** `XLEdge AIPowered Stage 0d: ExtractAndCatalog and shared dependencies`.

### Task 4: UpdateBootstrapper.ResolveVersionToLoad and catalog recovery

**Interfaces — Produces:** `ResolvedRelease ResolveVersionToLoad(IXLEdgeContext context, string source = "Install")` (null when nothing is loadable; never throws).

Behaviour: `Paths.Ensure()`, `Paths.Refresh()`. (1) Catalog **missing** and a seed is staged (manifest.json + zip): delete stray `Versions\` contents, extract with source `"Install"`; delete `Manifest\` **only on success**; return. (2) Catalog unreadable (`JsonException`) or present-but-empty while `Versions\` has releases → rename to `ReleaseHistory.corrupt-<timestamp>.json`, rebuild entries from each `Versions\<folder>\manifest.json` snapshot whose folder holds DLLs (Source `"Recovered"`), log a warning. (3) Staged manifest + zip: pick the zip (manifest `fileName`, else newest + warning); known reinstall (same Version+ReleaseDate in catalog) → `Reconcile`; extract with `source`; delete the zip after the catalog append; return. (4) Otherwise newest catalogued entry (parsed ReleaseDate) whose folder holds DLLs **and** whose contract is loadable. (5) Nothing → `LogError` and null. Never uses `Paths.LatestVersion` to choose the installed release.

- [ ] **Step 1: Tests**: `FreshInstall_Seed_ExtractsCataloguesAndRemovesManifest`; `FreshInstall_BadSeed_KeepsManifestAndReturnsNull`; `FreshInstall_WipesStrayVersions`; `Staged_ZipAndManifest_AddsSecondRelease_DeletesZipKeepsManifest`; `Staged_SeveralZips_UsesManifestFileName`; `KnownReinstall_Reconciles`; `NoZip_NewestLoadableCataloguedRelease` (skips folder without DLLs and contract-too-new entry); `CorruptCatalog_RecoveredFromSnapshots_VersionsKept`; `EmptyCatalog_RecoveredFromSnapshots_VersionsKept`; `Nothing_ReturnsNull`.
- [ ] **Step 2:** RED. **Step 3:** Implement. **Step 4:** 0/0, green ×3; smell check. **Step 5: Commit** `XLEdge AIPowered Stage 0d: ResolveVersionToLoad and catalog recovery` (+ this plan).
