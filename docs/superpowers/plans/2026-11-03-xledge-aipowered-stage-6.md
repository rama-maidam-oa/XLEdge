# XLEdge AIPowered Stage 6 — Online Update (Client) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reload Add-in → Online lists the XLEdge core releases the ORBIT server publishes, downloads the chosen one with the
user's login, catalogues it (SHA256 + contract checks) and reloads into it — against the endpoints a future orbitbootapp
`XLEdgeAddinController` will serve.

**Architecture:** GLSense 11.1.3's online reload (`GLSense\Views\GLReloadSourcePicker.xaml.cs`, `OnlineReleaseClassifier`,
`OnlineReleaseRow`; CLAUDE.md §44.1/75/76/83), XLEdge-named, in the shell. The core fills `IXLEdgeAddin.GetLoginInfo()` (no
contract change). The shell's `OnlineReleaseClient` calls `{LoginUrl}/rest/secure/xledge-addin/...` with the Bearer token; the
picker shows the classified rows; the chosen release is downloaded and catalogued **before** the core is unloaded, then
`Host.Reload(() => thatRelease)` (as Release History), because the twin's `ResolveVersionToLoad` prefers the last-loaded release
over the newest catalogued one.

**Tech Stack:** .NET Framework 4.8.1, WPF, System.Net.Http, MSTest.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§5 row 6).

## Developer's decisions (2026-11-03)

- **orbitbootapp is not touched in this stage.** The server side is written as a spec:
  `docs/superpowers/specs/2026-11-03-xledge-addin-online-update-server.md` — a copy-style `XLEdgeAddinController` (GLSense's
  controller copied, XLEdge-named) with **live** `@Test` JUnit tests, to be implemented later.
- At the end of the stage, **push `11.1.2-aipowered` to origin** (no PR).

## Global Constraints

- Endpoints (spec §5 row 6, GLSense §83): `GET {LoginUrl}/rest/secure/xledge-addin/manifests` (array of manifest entries),
  `GET …/versions/{folderName}/manifest`, `GET …/versions/{folderName}/zip`; `Authorization: Bearer {LoginToken}`;
  `LoginUrl` is the context root as the login stores it (e.g. `https://host/reporting`), trailing `/` trimmed; `folderName`
  URL-escaped (`Uri.EscapeDataString`).
- No contract change (`ContractInfo.Version` stays 5); the shell calls `GetLoginInfo()` in try/catch (older cores).
- `UpdateBootstrapper.ExtractAndCatalog` is the only way a downloaded release enters `Versions\` (its checksum/contract/folder checks stay).
- Messages: 401 → `"Your XLEdge session has expired. Log in again, then check for updates again."`; 404 → `"That release is not on the server (any more)."`; anything else → `"The XLEdge server could not be reached: <message>"`.
- 0 errors / 0 warnings; suite ×3; mutants (restored files touched); no Claude attribution.

## Rulings

1. **Download before unloading** — a failed download leaves the running core as it is (GLSense also fetches before reloading). *If wrong:* a release the user then cancels stays catalogued (shown as "Already downloaded").
2. **One release at a time** (GLSense can fetch several): the twin reloads into the one chosen; others can be fetched later. *If wrong:* one more round trip to get two releases.
3. **Online is enabled only when `GetLoginInfo().IsLoggedIn`** (`IsLoginCompleted` and a token) — GLSense §44.1 (a flag never set kept Online disabled): the test logs in through the real state, not a stub.

## Review Focus

1. The server answering 401 because the token expired while the picker is open → the message, no crash, the picker usable.
2. A server entry the user already has (same folder) or the loaded one → shown, not selectable.
3. A downloaded zip whose checksum does not match its manifest → refused by ExtractAndCatalog, the old core keeps running.
4. Logged out, or an older core without a working `GetLoginInfo` → Online disabled with a tooltip, Offline still works.
5. A slow server → the picker stays responsive (async), Cancel/close while checking does not crash.

---

### Task 1: The core reports its login

**Files:** `XLEdge.Addin.Core\AddinEntry.cs` (`GetLoginInfo`). Test: `XLEdge.Tests\CoreHandlersTests.cs` (CoreProbe).

- [ ] Step 1: failing test `GetLoginInfo_ReportsTheLogin`: before login → `IsLoggedIn` false, all strings empty/null; after
  setting `XLEdgeAppState.Instance` `LoginUrl = "https://host/reporting"`, `LoginToken = "tok"`, `LoginUserName = "u"`,
  `IsLoginCompleted = true` → `LoginUrl`, `LoginToken`, `LoginUserName` as set, `IsLoggedIn` true; token cleared → `IsLoggedIn` false.
- [ ] Step 2: run → fails (empty LoginInfo). Step 3: implement (`IsLoggedIn = IsLoginCompleted && token and URL not blank`;
  any exception logged, empty LoginInfo returned). Step 4: green + `bt.sh`. Step 5: commit
  `XLEdge AIPowered Stage 6: the core reports its login (GetLoginInfo)`.

### Task 2: The online release client and classifier

**Files:** create `XLEdge\Online\OnlineReleaseClient.cs`, `XLEdge\Online\OnlineReleaseException.cs`,
`XLEdge\Online\OnlineReleaseClassifier.cs`, `XLEdge\Online\OnlineReleaseRow.cs` (GLSense's two files, XLEdge-named);
tests `XLEdge.Tests\OnlineReleaseClientTests.cs` (a local TcpListener server like `InstanceServer`: the three endpoints,
a configurable status code, records request lines and Authorization headers), `XLEdge.Tests\OnlineReleaseClassifierTests.cs`.

**Interfaces — Produces:**
- `internal interface IOnlineReleaseSource { Task<IReadOnlyList<VersionInfo>> GetReleasesAsync(CancellationToken ct); Task<(string ManifestJson, byte[] Zip)> DownloadAsync(string folderName, CancellationToken ct); }`
- `internal sealed class OnlineReleaseClient : IOnlineReleaseSource, IDisposable` — `public OnlineReleaseClient(LoginInfo login, ILogger logger, HttpMessageHandler handler = null)`; 5-minute timeout; list parsed with `new VersionParser(logger).ParseVersionJson`.
- `internal sealed class OnlineReleaseException : Exception` — `StatusCode` (int?), user message per Global Constraints.
- `internal static class OnlineReleaseClassifier { public static List<OnlineReleaseRow> Classify(IEnumerable<VersionInfo> serverEntries, IReadOnlyList<ReleaseEntry> localEntries, string currentVersion, string currentReleaseDate) }` and `OnlineReleaseRow` (GLSense's members: Version, ReleaseDate, Mandatory, FolderName, FileName, Checksum, Notes, IsCurrentlyLoaded, IsAlreadyDownloaded, IsSelectable, StatusText), newest first.

- [ ] Step 1: failing tests:
  - `Client_ListsTheServersReleases_WithTheBearerToken`: server returns 2 entries → 2 VersionInfo (versions, folderNames);
    request line `GET /reporting/rest/secure/xledge-addin/manifests`; header `Authorization: Bearer tok`.
  - `Client_DownloadsManifestAndZip_OfTheEscapedFolder`: folder `v11.1.3.0_2026-11-03T10-00-00` → requests `…/versions/v11.1.3.0_2026-11-03T10-00-00/manifest` and `…/zip`; bytes equal the served zip.
  - `Client_Unauthorized_SaysTheSessionExpired` (401 → message), `Client_NotFound_SaysNotOnTheServer` (404), `Client_ServerDown_SaysItCouldNotBeReached` (closed port).
  - `Client_LoginUrlWithTrailingSlash_IsTrimmed`.
  - Classifier (GLSense's cases): loaded → `IsCurrentlyLoaded`, not selectable; same folder in history → `IsAlreadyDownloaded`; new → selectable; newest first; null/blank entries skipped.
- [ ] Step 2: run → fail. Step 3: implement. Step 4: green + `bt.sh`. Step 5: commit
  `XLEdge AIPowered Stage 6: online release client and classifier (xledge-addin endpoints)`.

### Task 3: Online in the reload picker

**Files:** `XLEdge\Views\XLEdgeReloadSourcePicker.xaml(.cs)` (Online radio enabled when logged in; "Check for updates" button;
a release list bound to `OnlineReleaseRow`; status text; Reload fetches the selected row); `XLEdge\AddinModule.cs`
(`ShowReloadWindow`: LoginInfo from `Host.Addin.GetLoginInfo()` in try/catch → picker; an online result →
`Host.Reload(() => resolved)`); tests in `XLEdge.Tests\HostWindowsTests.cs` (the existing `RbOnline disabled` test changes).

**Interfaces — Consumes:** Task 2's `IOnlineReleaseSource`, `OnlineReleaseClassifier`. **Produces:**
`XLEdgeReloadSourcePicker(IXLEdgeContext context, LoginInfo login, Func<LoginInfo, IOnlineReleaseSource> sourceFactory = null)`;
`public ResolvedRelease OnlineRelease { get; }` (set when the online fetch + `ExtractAndCatalog(context, manifest, zip, "Online")`
succeeded); `SelectedSource == "Online"` then.

- [ ] Step 1: failing tests (STA, a fake `IOnlineReleaseSource`; the picker driven through its methods/controls as the existing tests do):
  - `Picker_LoggedOut_OnlineDisabledWithTheReason` (tooltip `"Log in to XLEdge to check for online updates."`); `Picker_LoggedIn_OnlineEnabled`.
  - `Picker_Check_ListsClassifiedRows` (loaded row not selectable).
  - `Picker_ReloadOnline_DownloadsCataloguesAndReturnsTheRelease` (a real zip+manifest from `ReleaseFolders` helpers → `OnlineRelease.FolderName` set, folder in `Versions\`).
  - `Picker_ReloadOnline_BadChecksum_StaysOpenWithTheError` (no `OnlineRelease`, dialog not closed, message shown).
  - `Picker_Check_Unauthorized_ShowsTheSessionMessage`.
  - AddinModule wiring is shell-only ADX code: a source check that `ShowReloadWindow` passes `OnlineRelease` to `Host.Reload(() => …)` and calls `GetLoginInfo` inside try/catch.
- [ ] Step 2: run → fail. Step 3: implement (WPF: async handlers, `SetBusy` as today, logging via `SafeLog`). Step 4: green +
  `bt.sh`. Step 5: commit `XLEdge AIPowered Stage 6: Online in the reload picker`.

### Task 4: The server spec (orbitbootapp, for later)

**Files:** create `docs/superpowers/specs/2026-11-03-xledge-addin-online-update-server.md`: `XLEdgeAddinController` as a copy of
`orbit-app\orbit-web\src\main\java\com\orbit\reporting\cpm\controller\GLSenseAddinController.java` (`@RequestMapping("/rest/secure/xledge-addin")`,
`xledgeManifests.json`, `addons\xledge\updates\`), the same path-safety rules, errors and logging; `XLEdgeAddinControllerTest` with
**live** `@Test`s (index, manifest, zip header, 404 unlisted / `..` / escaping); the server folder layout and the publish steps
(copy `XLEdge.Addin.Core\SetupFiles\Release\Manifest\{v….zip, manifest.json}` into `updates\{folderName}\`, add the entry to
`xledgeManifests.json`); build/test command (`orbit-app\mvnw.cmd -q -pl :orbit-web -am test -Dtest=XLEdgeAddinControllerTest`);
the OISR/commit conventions of orbitbootapp's CLAUDE.md. No code in orbitbootapp.
- [ ] Commit `XLEdge AIPowered Stage 6: server spec for orbitbootapp's XLEdgeAddinController (not implemented here)`.

### Task 5: Close the stage

- [ ] Fresh opus reviewer; one fix pass; `bt.sh` ×3; mutants (touch restored files): GetLoginInfo ignoring the token; no Bearer
  header; 401 mapped to the generic message; trailing slash not trimmed; folder not escaped; classifier marking a downloaded
  folder selectable; picker reloading without cataloguing; `ShowReloadWindow` ignoring `OnlineRelease`.
- [ ] `AIPoweredXLEdge.md`: Stage 6 done (client); the server spec; Excel/live check items (needs the server: Check lists,
  Reload Online loads the new core, 401 after logout, older core → Online disabled).
- [ ] Commit; then **push `11.1.2-aipowered` to origin** (`git push -u origin 11.1.2-aipowered`; if the harness blocks it, give
  the developer the command).
