# XLEdge AIPowered Stage 5 — Packaging & MSI Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A Release build of the twin produces signed DLLs and an `OrbitXLEdge.msi` (11.1.3) that upgrades today's XLEdge
11.1.2 install, ships the first core inside it, registers the shell, and removes `AddinCore\` on uninstall.

**Architecture:** GLSense 11.1.3's packaging (`D:\SQLLite_Test\GLSense\AIPowered\GLSense`, CLAUDE.md §41/53/55/62/63/66/74),
XLEdge-named: libraries sign their own DLL in their post-build (Release only), the core signs before it zips, a Release
build copies the published zip to a fixed-name installer payload, and the setup project (Primary Output of the shell with
core-only dependencies excluded) installs shell + loaders + payload + manifest, registers through adxregistrator and cleans
`AddinCore\` on a real uninstall. The setup project starts from today's 11.1.2 `OrbitXLEdge.vdproj` (same UpgradeCode, folder,
registration), so the upgrade path is the one that already works.

**Tech Stack:** .NET Framework 4.8.1, Add-in Express (adxloader, adxregistrator, adxpatch), VS 2022 Installer Projects, PowerShell.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md` (§5 row 5, §6 deviation 3, §7).

## Developer's decisions (2026-11-03)

- **Version 11.1.3** for the MSI, the shell, every twin assembly and the first core (above today's 11.1.2, same UpgradeCode).
- **Signing through `sign_file.cmd`** at the twin root (`AIPowered\XLEdge\sign_file.cmd`), copied there by the developer,
  called as GLSense does (`"<file>" "<Configuration>"`). Claude never opens it, and builds/tests with signing skipped.
- **Setup project drafted by Claude** on GLSense 11.1.3's pattern (Primary Output + Exclude flags), built with `devenv.com`
  here (signing skipped) and checked with `msiexec /a`; the developer opens it once in VS to confirm the excludes and runs the
  real install-over-11.1.2 test.

## Global Constraints

- Spec §6 deviation 3 ("shell as plain File") is superseded by GLSense §74 (Primary Output with UI excludes) — the
  developer's decision above; the spec gets a note.
- Signing happens only in **Release** and only when `XLEDGE_SKIP_SIGNING` is not `1`; the build tools (`bt.sh`, `rt.sh`) and
  every MSI build Claude runs set `XLEDGE_SKIP_SIGNING=1` (DigiCert signings are metered — GLSense §55). A Release build
  without `sign_file.cmd` and without the skip variable **fails** with `"sign_file.cmd not found at <path> - copy it there, or
  set XLEDGE_SKIP_SIGNING=1 for an unsigned build"` (an unsigned MSI must never be built by accident).
- `sign_file.cmd` is git-ignored in the twin (`/AIPowered/XLEdge/sign_file.cmd` in `.gitignore`); Claude never reads it.
- Loader file names stay `adxloader.dll` / `adxloader64.dll` (today's, simplest upgrade; spec open item closed).
- Verbatim parity: `XLEdgeAppConstants.DefaultVersion` `"11.1.2"` → `"11.1.3"` is one more `TwinSubstitutions` entry (the
  About window compares it with the server's advertised XLEdge version).
- 0 errors / 0 warnings; suite ×3; local commits only; no Claude attribution; `adxregistrator` / installs are the
  developer's (Claude only builds the MSI and runs `msiexec /a` into a temp folder — no install, no registration).

## Rulings

1. **Skip variable instead of "build in Debug"** — the Release half of `bt.sh` must keep running without spending signings;
   GLSense's reverted opt-out was for developers, this one is set only by Claude's tools. *If wrong:* the developer has one more
   variable to know (documented in the notes and the failure message).
2. **`sign_file.cmd` is git-ignored** — it may hold credentials; GLSense's tracking of its own copy is not checked (never opened).
   *If wrong:* the developer adds it to git themselves.
3. **Signed files:** XLEdge.Contracts, XLEdge.Shared, XLEdge.Loader.Core (own post-builds), XLEdge.Addin.Core (before the zip),
   XLEdge.dll and both `bin\Release` adxloaders (setup project PreBuildEvent, as GLSense). The `.msi` itself stays a manual
   signtool step (as today, release memory). *If wrong:* one more `sign_file.cmd` call.

## Review Focus

1. A Release build on a machine without `sign_file.cmd` and without the skip variable → fails with the message above.
2. The installer payload and the seeded `manifest.json` come from the same build (checksum matches) — else the first core never installs.
3. Upgrade over 11.1.2: one XLEdge registration, the old `XLEdge.dll` replaced, `adxloader*.dll` replaced (equal-version files
   dropped by component-GUID changes — verify by listing the installed folder).
4. Uninstall removes `AddinCore\`; a major upgrade does not (`NOT UPGRADINGPRODUCTCODE`).
5. No core-only DLL in the MSI's shell folder (MahApps, WebView2, Wpf.Ui, Xaml.Behaviors, XLEdge.Addin.Core, …).

---

### Task 1: Version 11.1.3

**Files:** `AIPowered\XLEdge\XLEdgeSharedVersion.cs` (11.1.3.0), `XLEdge.Tests\CorePortParityTests.cs` (`TwinSubstitutions` +
`(@"XLEdgeAppConstants.cs", null, "DefaultVersion = \"11.1.2\"", "DefaultVersion = \"11.1.3\"", 1)`), `XLEdge.Addin.Core\XLEdgeAppConstants.cs`,
tests pinning `"11.1.2.0"` as the shell/assembly version (`XLEdgeContextTests`, `FakeContext.ShellVersion`, others found by
`grep "11.1.2.0"` — only where the value means "this build's version"; release-fixture strings stay).

- [ ] Step 1: change the tests' expected shell version to `"11.1.3.0"` and add the substitution → run → `Context_ShellVersionLeaseAndDefaults`, `TwinSubstitutions_AreExactAndPresent`, `PortedFiles_AreByteIdentical…` fail.
- [ ] Step 2: bump `XLEdgeSharedVersion.cs`, apply the substitution → `bt.sh` green → commit `XLEdge AIPowered Stage 5: version 11.1.3`.

### Task 2: Release signing hook

**Files:** create `AIPowered\XLEdge\Build\sign.ps1` — `param([string]$File, [string]$Configuration, [string]$SignScript)`;
no-op unless `$Configuration -eq 'Release'`; no-op (prints `"XLEdge: signing skipped (XLEDGE_SKIP_SIGNING=1): <file>"`) when
`$env:XLEDGE_SKIP_SIGNING -eq '1'`; fails (exit 1, the Global Constraints message) when `$SignScript` is missing; else runs
`cmd /c "<SignScript>" "<File>" "<Configuration>"` and fails on a non-zero exit. Default `$SignScript` =
`<twin root>\sign_file.cmd`. Post-build calls: Contracts, Shared, Loader.Core csproj PostBuildEvent (own DLL), core
`post_build.ps1` (before the zip; `-SignScript` passed through for tests). `.superpowers\tools\bt.sh` / `rt.sh` export
`XLEDGE_SKIP_SIGNING=1`. `.gitignore` gets `/AIPowered/XLEdge/sign_file.cmd`. Test: `XLEdge.Tests\CorePackageTests.cs`.

- [ ] Step 1: failing tests (run `sign.ps1` through `powershell -NoProfile -File`, a fake sign script in a TempFolder that
  appends `"<file>|<config>"` to a log):
  - `Sign_Release_CallsTheSignScriptWithFileAndConfiguration` → log `["C:\x\a.dll|Release"]`, exit 0.
  - `Sign_Debug_DoesNothing`, `Sign_SkipVariable_DoesNothingAndSaysSo` (stdout has `signing skipped`).
  - `Sign_Release_WithoutTheScript_FailsWithTheMessage` (exit 1, stdout/stderr has `sign_file.cmd not found`).
  - `Sign_ScriptFails_TheBuildFails` (fake exits 3 → exit non-zero).
  - `PostBuilds_SignWhatGLSenseSigns` (source): Contracts/Shared/Loader.Core csproj PostBuildEvent and `post_build.ps1` call `sign.ps1` for their DLL; in `post_build.ps1` the sign call comes before `Compress-Archive`/zip creation.
- [ ] Step 2: run → fail. Step 3: implement. Step 4: `bt.sh` green (skip variable set by the tools). Step 5: commit
  `XLEdge AIPowered Stage 5: Release signing hook (sign_file.cmd at the twin root; skipped by XLEDGE_SKIP_SIGNING=1)`.

### Task 3: Installer payload and release notes

**Files:** `XLEdge.Addin.Core\post_build.ps1` (Release only: byte copy of the published zip to
`SetupFiles\Release\InstallerPayload\OrbitXLEdge.zip`, old payload deleted first); create `XLEdge.Addin.Core\ReleaseNotes.txt`
(empty first line + a comment line explaining line 1 becomes the manifest's notes — GLSense §75 trims). Test: `CorePackageTests.cs`.

- [ ] Step 1: failing tests (existing CorePackageTests style — run `post_build.ps1` against a staged TargetDir):
  - `ReleaseBuild_CopiesThePublishedZipToTheInstallerPayload`: payload bytes == published zip bytes; manifest checksum == SHA256(payload).
  - `DebugBuild_WritesNoInstallerPayload`.
  - `ReleaseNotes_FirstLine_BecomesTheNotes_Trimmed` (`"  fixes  "` → `"fixes"`); empty first line → `""`.
- [ ] Step 2–5 as Task 2; commit `XLEdge AIPowered Stage 5: installer payload (OrbitXLEdge.zip) and ReleaseNotes.txt`.

### Task 4: The setup project

**Files:** create `AIPowered\XLEdge\XLEdge\OrbitXLEdgeSetups\OrbitXLEdge.vdproj` from `XLEdge\setupfiles\OrbitXLEdge.vdproj`
(11.1.2) with these changes; add it to `AIPowered\XLEdge\XLEdge.sln` (Release builds it, Debug does not — as 11.1.2's sln).
- ProductVersion `11.1.3`; new ProductCode and PackageCode; UpgradeCode, ProductName, Manufacturer, TARGETDIR, InstallAllUsers,
  RemovePreviousVersions, DetectNewerInstalledVersion unchanged.
- Primary Output (Built) of the twin shell project `{98093E3C-A05D-4EE1-BF6D-AC742BE9F782}` instead of 11.1.2's; drop the stale
  `..\obj\Debug\XLEdge.dll` file entry.
- Exclude (`"Exclude" = "11:TRUE"`) every detected core-only dependency: XLEdge.Addin.Core, MahApps.Metro*, ControlzEx,
  Microsoft.Web.WebView2.*, Wpf.Ui*, Microsoft.Xaml.Behaviors, System.Net.Http — final list from Step 3's `msiexec /a` diff.
- `adxloader.dll`, `adxloader64.dll` from `..\bin\Release\` (the signed copies, GLSense §66.1).
- Folder `AddinCore\Manifest\` with `..\..\XLEdge.Addin.Core\SetupFiles\Release\InstallerPayload\OrbitXLEdge.zip` and
  `..\..\XLEdge.Addin.Core\SetupFiles\Release\Manifest\manifest.json`.
- Uninstall cleanup (GLSense §62): File `C:\Windows\System32\cmd.exe` as `XLEdgeUninstallCleanup.exe`; uninstall custom
  action, Condition `REMOVE="ALL" AND NOT UPGRADINGPRODUCTCODE`, Arguments
  `/c if exist "[TARGETDIR]AddinCore" rd /s /q "[TARGETDIR]AddinCore" & exit /b 0`.
- PreBuildEvent: today's `CleanOnlineAppCache`, then `powershell -NoProfile -File "..\..\Build\sign.ps1"` for `..\bin\Release\XLEdge.dll`,
  `adxloader.dll`, `adxloader64.dll` with `Release`. PostBuildEvent: today's adxpatch line.
- Test: new `XLEdge.Tests\SetupProjectTests.cs` (source checks): UpgradeCode = 11.1.2's `{D090F1E0-CA8E-4529-B78D-981652FFF708}`;
  ProductVersion `11.1.3`; ProductCode/PackageCode differ from 11.1.2's; the shell output is the twin project; loaders from
  `bin\Release`; payload + manifest under `AddinCore\Manifest`; the cleanup action and its condition; the PreBuild sign calls;
  each excluded name has `Exclude 11:TRUE`.

- [ ] Step 1: `SetupProjectTests` → fail (no file). Step 2: write the vdproj (byte-safe: VS writes UTF-8 with CRLF).
- [ ] Step 3: build it: `XLEDGE_SKIP_SIGNING=1` + `devenv.com AIPowered\XLEdge\XLEdge.sln /Rebuild "Release"` → MSI; then
  `msiexec /a <msi> /qn TARGETDIR=<temp>` and list the files; expected: shell + its own dependencies, adxloader*, manifest,
  `AddinCore\Manifest\{OrbitXLEdge.zip, manifest.json}`, `XLEdgeUninstallCleanup.exe`, no core-only DLL. Adjust excludes until
  so; record the file list in the ledger. The temp folder is deleted after.
- [ ] Step 4: `SetupProjectTests` + `bt.sh` green; commit `XLEdge AIPowered Stage 5: the setup project (OrbitXLEdge.msi 11.1.3)`.

### Task 5: Close the stage

- [ ] Fresh opus reviewer; one fix pass; `bt.sh` ×3; mutants (touching restored files): `sign.ps1` signing in Debug / ignoring
  the skip / passing a missing script; payload not copied; payload from another zip; vdproj UpgradeCode changed; cleanup
  condition without `NOT UPGRADINGPRODUCTCODE`.
- [ ] Spec §6 deviation 3 note (superseded by GLSense §74, developer's decision 2026-11-03).
- [ ] `AIPoweredXLEdge.md`: Stage 5 done; the developer's steps — copy `sign_file.cmd` to `AIPowered\XLEdge\`, open the setup
  project once in VS (confirm excludes; discard VS's GUID/order drift, GLSense §65.3), Release build, sign the `.msi`, install
  over 11.1.2 (registration count, installed files, Excel loads the twin with core 11.1.3), uninstall (AddinCore\ gone).
- [ ] Commit.
