# XLEdge AIPowered 11.1.3 — QA & cutover runbook

For QA and the developer. Everything here runs on a real Windows machine with Excel (Claude cannot drive Excel or install
from its session). Fill the **Result** column with PASS / FAIL / N/A and a note; a FAIL gets a ticket (OISR) with the
health-check output and the XLEdge log attached.

- Build/test status at hand-over (2026-11-03): 371 automated tests, 364 pass, 7 Inconclusive (the real-Excel tests: this dev
  machine's Office cannot marshal Excel's typed interfaces); branch `11.1.2-aipowered` (pushed to origin).
- Logs: `%LOCALAPPDATA%\ORBIT\Excel_Logs\XLEdge_Logs\Logs\XLEdge_Logs_dd-MMM-yyyy.log` (GLSense: its own `GLSense_Logs`).
- Install folder: `%LOCALAPPDATA%\Orbit Analytics\XLEdge` (per user); cores under `AddinCore\Versions\`.

## 0. The health check

`AIPowered\XLEdge\tools\XLEdgeHealthCheck.ps1` — read-only; run it before and after every install, upgrade, rollback and
uninstall, and attach its output to any FAIL:

```
powershell -ExecutionPolicy Bypass -File XLEdgeHealthCheck.ps1            # PASS/WARN/FAIL lines, exit 1 on any FAIL
powershell -ExecutionPolicy Bypass -File XLEdgeHealthCheck.ps1 -Json > health.json
```

It reports: the Excel add-in registration (LoadBehavior 3), the registered loader (→ install root), XLEdge.dll's version and
signature, Contracts/Shared/Loader.Core, the catalogued core releases, the release Excel loads next, a staged release,
today's ERROR lines, and whether GLSense is registered. A development build shows `Signature: WARN` (unsigned) — a release
build must show PASS.

## 1. Setup

| ID | Step | Expected | Result |
|---|---|---|---|
| S1 | Release build per `AIPoweredXLEdge.md` → "Developer's release check of Stage 5" (sign_file.cmd in `AIPowered\XLEdge\`, `devenv … OrbitXLEdgeSetup.sln /Rebuild "Release"`, sign the `.msi`) | build succeeds; no `XLEDGE0005` (unsigned file) | |
| S2 | Test machine: Excel 64-bit (and one 32-bit machine if supported), XLEdge 11.1.2 installed from today's MSI for the upgrade tests; one clean machine for the fresh install | | |
| S3 | ORBIT instance reachable; a user with XLEdge permission and one without; a user with GLSense cubes | | |

## 2. Install, upgrade, repair

| ID | Step | Expected | Result |
|---|---|---|---|
| I1 | Fresh install (no XLEdge): `msiexec /i OrbitXLEdge.msi /l*v install.log` | one "XLEdge" in Apps; health check: Registration PASS, Loader PASS, Shell 11.1.3.0, Signature PASS, StagedRelease "Staged: …" | |
| I2 | Start Excel once | XLEdge tab; About: shell 11.1.3.0, core 11.1.3.0; health check: AddinCore 1 release, Manifest\ empty | |
| I3 | Upgrade over 11.1.2: `msiexec /i OrbitXLEdge.msi /l*v upgrade.log` | one XLEdge in Apps; XLEdge.dll 11.1.3.0 signed; first Excel start loads core 11.1.3.0; `upgrade.log`: 11.1.2's `adxregistrator /uninstall` action skipped during RemoveExistingProducts; XLEdge still loads after a second Excel start | |
| I4 | After I3: `adxloader.dll` / `adxloader64.dll` | still 11.1.2's files (same binary, unsigned) — accepted, note it | |
| I5 | Repair (Apps → Modify/Repair) | XLEdge still loads; catalogued releases kept | |

## 3. Regression matrix (spec §7 + every stage's Excel checks)

| ID | Area | Steps | Expected | Result |
|---|---|---|---|---|
| R1 | Login | Login → Server Configuration → instance | pane shows Home; ribbon LoggedIn with the instance name; Broadcast enabled | |
| R2 | Login (SSO) | instance with SSO | logs in without a password prompt | |
| R3 | Logout | Logout with panes open in two windows | every pane hides; next Login asks for credentials | |
| R4 | Two windows | open a second Excel window; Show/Hide in each | acts on that window's pane only | |
| R5 | Reports | run a report from the pane | data sheet + `P_` sheet; formats right | |
| R6 | Title commands | Edge / Process / EdgeWorkbook / Logs / Excel from the instance pages | each works; nothing from a non-instance page | |
| R7 | XLSX download | a download command from the pane | the file opens/saves as in 11.1.2 | |
| R8 | Refresh | Refresh Sheet, Refresh Book, Param Refresh (sheet, book) | values refresh in place, formulas kept | |
| R9 | Control sheet, Options, About | each from the ribbon | as 11.1.2; About shows shell + core versions | |
| R10 | Drilldown | drilldown hyperlink, "Goto Report Data" | opens the drilldown / activates the data sheet | |
| R11 | Attachments | Fusion/EBS attachment link; an empty attachment cell (OISR-22766/22767) | downloads; empty cell does nothing harmful | |
| R12 | Pickers | Calendar and segment pickers on cell selection | open, apply, close | |
| R13 | Sheet delete | delete a report sheet | its `P_` sheet and named ranges go a moment later | |
| R14 | Regional formats | non-English region (OISR-22663/22696) | numbers/dates right | |
| R15 | Broadcast | logged out / logged in / after a Reload Add-in | disabled / shows messages or "No broadcast messages available." / still enabled | |
| R16 | Reload offline | Reload Add-in → Offline → a folder with a release zip + manifest | loads it; Release History lists it | |
| R17 | Reload online | (needs orbitbootapp `XLEdgeAddinController`) Online → Check for updates | without the controller: "This XLEdge server does not offer XLEdge updates (yet)."; with it: list, Reload loads the newest ticked, Cancel during a download stops it, logged out → Online disabled with its reason; over HTTPS | |
| R18 | Release History | switch to an older core | loads it; survives an Excel restart | |
| R19 | Busy refusal | Reload while a report runs, during an XLSX download, while the pane logs in | refused with the busy message | |
| R20 | Hot reload + pane | hide the pane → Reload → Show/Hide; Reload with the pane open | pane back, logged in; logs in by itself | |
| R21 | Workbook switch | switch workbooks while a report loads, and while Refresh All runs | the run ends in the workbook it started in (that workbook comes back to the front) | |
| R22 | Two windows, one workbook | "Book1:2" active, run a report | stays in your window | |
| R23 | Excel close | close Excel with an XLEdge window open | no orphan EXCEL.EXE (Task Manager); next start clean | |
| R24 | Update extraction | after R16/R17, restart Excel | the chosen core loads; health check SelectedRelease PASS | |
| R25 | Stuck unload | (hard to provoke) a reload whose previous core does not unload in 5 s | message "did not close within 5 seconds … restart Excel"; XLEdge keeps working | |

## 4. GLSense side by side (both installed and loaded in one Excel)

Health check: `GLSense: PASS … LoadBehavior 3`. **Switch on Debug logging in both add-ins first** (the "Debug" toggle on
each ribbon tab): nearly every line below is a DEBUG line and is not written otherwise. Look in both logs after each row.

| ID | Steps | Expected | Log evidence | Result |
|---|---|---|---|---|
| G1 | Login in GLSense | XLEdge logs in (ribbon LoggedIn, Broadcast enabled); an open XLEdge pane navigates with the login | XLEdge: `AddinEntry.OnLoginFromGLSense`, `InvokedFromGLSense\|Login invoked from GLSense` | |
| G2 | Login in XLEdge by itself | GLSense goes PartialLoggedIn with its cubes; Excel does not freeze | XLEdge: `XLEdgeContext.CallGLSense: GetGLCubeInformation called`; GLSense: `GetGLCubeInformation called from XLEdge`, `LoginFromXLEdge: N cubes loaded` | |
| G3 | Logout in GLSense | XLEdge logs out too; GLSense shows **one** logout message, no "unexpected error" | XLEdge: `AddinEntry.OnLogoutFromGLSense`, `CallGLSense: LogoutSession called`; GLSense: `LogoutFromXLEdge: GLSense's own logout is running` | |
| G4 | Logout in XLEdge | GLSense LoggedOut | GLSense: `LogoutSession called from XLEdge`, `LogoutFromXLEdge: logged out with XLEdge` | |
| G5 | GLSense login, user without XLEdge permission | XLEdge stays logged out, its Login disabled | XLEdge: `does not have access to XLEdge` | |
| G6 | GLSense not loaded (disable it in COM Add-ins) | XLEdge login/logout work, nothing shown | XLEdge: `GLSense is not connected` | |
| G7 | Logged in through GLSense → Reload Add-in (XLEdge) → Logout in XLEdge | GLSense logged out too | XLEdge: `CallGLSense: LogoutSession called` after the reload | |
| G8 | Log out of XLEdge right after logging in (GLSense still loading cubes) | GLSense ends LoggedOut (no PartialLoggedIn afterwards) | GLSense: `the login changed while the cubes were loading` | |
| G9 | Two Excel windows, both add-ins | logins/logouts as above in each window; panes per window | | |
| G10 | Reload GLSense's core while XLEdge is logged in (and the other way round) | the other add-in unaffected; sync still works afterwards | | |

## 5. Rollback

| ID | Steps | Expected | Result |
|---|---|---|---|
| B1 | Release History → an older core | loads; survives a restart (SelectedRelease) | |
| B2 | Back to XLEdge 11.1.2: **uninstall 11.1.3 first** (DetectNewerInstalledVersion blocks installing 11.1.2 over it), then install 11.1.2's MSI | 11.1.2 loads; health check: AddinCore WARN "not set up" (the uninstall removed `AddinCore\`) | |
| B3 | Re-install 11.1.3 after B2 | as I3 | |

## 6. Uninstall

| ID | Steps | Expected | Result |
|---|---|---|---|
| U1 | Apps → XLEdge → Uninstall | `%LOCALAPPDATA%\Orbit Analytics\XLEdge` gone (incl. `AddinCore\`); XLEdge not in Excel's COM add-ins; health check: Registration FAIL (expected now) | |
| U2 | Logs kept | `%LOCALAPPDATA%\ORBIT\Excel_Logs\XLEdge_Logs` still there (not part of the install) | |

## 7. Cutover (documented — the developer decides when and how)

Nothing below is done yet. In order, each a decision:

1. **Go / no-go**: sections 2–6 PASS (or accepted), the Stage 5 release check done, open OISRs triaged.
2. **Release line**: make the AIPowered twin the XLEdge release line (e.g. a release branch from `11.1.2-aipowered`, or merge
   into the version branch the team ships from); retire the single-DLL `XLEdge\` project for new releases. The twin carries
   14+1 deliberate changes to 11.1.2's verbatim code (`CorePortParityTests.TwinSubstitutions`) — raise them as OISRs for
   11.1.2 or accept the divergence.
3. **Version/date**: `XLEdgeSharedVersion.cs` (11.1.3.0), `XLEdgeAppConstants.DefaultCommitDate` (release day; a
   TwinSubstitutions entry), the vdproj ProductVersion/ProductCode/PackageCode for the next MSI.
4. **MSI**: build/sign per the release memory (twin section): `OrbitXLEdgeSetup.sln /Rebuild "Release"`, sign the `.msi`,
   copy to `D:\Orbit Excel Add-ons\Setup Files\XLEdge\<version>\`.
5. **orbitbootapp** (both branches as for 11.1.2: `development` and `release/ORBIT<version>`): the MSI under
   `orbit-app\ORBIT_HOME\download-artifacts\addons\xledge\`; `orbit-artifacts.xml` XLEDGE `<version>` **and** `<release_date>`;
   the `XLEdgeAddinController` (separate engineer, spec `docs/superpowers/specs/2026-11-03-xledge-addin-online-update-server.md`)
   and the `updates\` folder with the first published core.
6. **Jira**: the release tickets to QA Review as usual.
7. **GLSense**: its 11.1.2/11.1.3 commits for the XLEdge sync (`GetGLCubeInformation` / `LogoutSession`) ship with a GLSense
   release; the likely unsigned `GLSense.dll` in GLSense 11.1.3's MSI (see `AIPoweredXLEdge.md` → Machine notes) fixed first.
8. **Rollback plan** for the field: section 5 (B2) as the documented way back to 11.1.2.
