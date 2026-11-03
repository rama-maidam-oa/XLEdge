# orbitbootapp: XLEdge add-in online update endpoints (spec, not implemented)

Date: 2026-11-03 · Status: **to be implemented later** — the developer decided (Stage 6) not to touch orbitbootapp now.
The XLEdge AIPowered client (Stage 6) already calls these endpoints; until they exist Check for updates gets a 404 and says
"This XLEdge server does not offer XLEdge updates (yet)." (a missing index file answers 404 too).

## What to build

A copy of GLSense's controller, XLEdge-named (developer's decision: copy, not a shared base class or a `{product}` path):

- Source to copy: `orbit-app\orbit-web\src\main\java\com\orbit\reporting\cpm\controller\GLSenseAddinController.java`
  (111 lines; identical on `development` and `release/ORBIT11.1.2` as of 2026-11-03, heads `28e9a28683` / `856edd6083`).
- New file: `…\cpm\controller\XLEdgeAddinController.java`
  - `@RestController("XLEdgeAddinController")`, `@RequestMapping("/rest/secure/xledge-addin")`
  - `static final String MANIFESTS_FILE = "xledgeManifests.json"`
  - home: `Paths.get(reportingProperties.getArtifactsLocation(), "addons", "xledge", "updates")`
    (= `ORBIT_HOME\download-artifacts\addons\xledge\updates\`)
  - `GET /manifests` → the index file (`application/json`)
  - `GET /versions/{folderName}/manifest` → `{folderName}\manifest.json` (`application/json`)
  - `GET /versions/{folderName}/zip` → `{folderName}\{fileName}` with `fileName` taken from the index entry, as an attachment
    (`Content-Disposition: attachment; filename=…`, `application/octet-stream`)
  - Same rules as GLSense's: a folder not listed in the index → 404; unreadable index → 500; every served path normalised
    and required to start with the home folder and be a regular file (a `..` or an escaping `fileName` → 404);
    `ResponseStatusException`; `@Slf4j` (debug on 404, error on 500). Error texts say "XLEdge".
  - Auth: as GLSense's — none on the controller; `SecurityConfiguration` (`/rest/secure/**` → `anyRequest().authenticated()`,
    `TokenAuthRequestFilter` Bearer JWT + active session). Any logged-in user may download (same as GLSense; the client sends
    the XLEdge login's `XL-AUTH-TOKEN`/`ORB-AUTH-TOKEN`, already accepted by other `/rest/secure` XLEdge calls).
- New test: `orbit-web\src\test\java\…\cpm\controller\XLEdgeAddinControllerTest.java`, a copy of
  `GLSenseAddinControllerTest.java` (JUnit 5, Mockito, `@TempDir`, `ReflectionTestUtils`) with **live `@Test`s**
  (GLSense's are commented out — developer's decision: XLEdge's run): index served; a listed release's manifest; the zip with
  its `Content-Disposition`; 404 for an unlisted folder, a `..` folder and an index entry whose `fileName` escapes the folder.
- Run: `cd orbit-app && .\mvnw.cmd -q -pl :orbit-web -am test -Dtest=XLEdgeAddinControllerTest -Dsurefire.failIfNoSpecifiedTests=false`
- Conventions (orbitbootapp `CLAUDE.md`): branch/commit/PR carry only the ticket's own OISR key; commit subject
  `OISR-##### summary` + Dev note; no AI attribution; base branch `development`.

## Server folder and publishing a release

```
ORBIT_HOME\download-artifacts\addons\xledge\
    OrbitXLEdge.msi                       (already there - the installer)
    updates\
        xledgeManifests.json              the index: a JSON array of manifest entries (newest last or any order)
        v11.1.3.0_2026-11-03T10-00-00\
            manifest.json                 the build's manifest (one-entry array)
            v11.1.3.0_2026-11-03T10-00-00.zip
```

To publish a core release (from a **signed Release** build, `OrbitXLEdgeSetup.sln /Rebuild "Release"`):
1. From `AIPowered\XLEdge\XLEdge.Addin.Core\SetupFiles\Release\Manifest\` copy `manifest.json` and the `v….zip` into
   `updates\<folderName>\` (`folderName` from the manifest).
2. Append the manifest's single entry to `updates\xledgeManifests.json` (create it as `[]` the first time).
3. The entry's fields (as the twin's `post_build.ps1` writes them): `version`, `releaseDate`, `folderName`, `fileName`,
   `checksum` (SHA-256 of the zip — the server never checks it, the client does), `notes`, `mandatory`, `contractVersion`.

## Client (already done in the twin, Stage 6)

`XLEdge\Online\OnlineReleaseClient.cs` → these three URLs under `{LoginUrl}` (the login's context root, e.g.
`https://host/reporting`), `Authorization: Bearer {token}`; 401 → "session expired, log in again"; 404 → "not on the server".
The reload picker lists the releases (loaded / already downloaded / new), downloads every ticked one, catalogues it after the
checksum and contract checks, and loads the newest.

## Live check once the controller exists

Server running with `updates\` filled → Excel with the twin, logged in → Reload Add-in → Online → Check for updates lists
the release → Reload loads it (About shows its version); log out and Check → the session message; a release folder removed
from the server but still listed → 404 message for that release, the others still load.
