# XLEdge AIPowered Stage 2b — Helpers, Models, JSON, API Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move XLEdge 11.1.2's helpers, models, JSON, API and formatting code — everything that needs neither Excel nor a
window — into the core verbatim, with tests that pin today's behaviour.

**Architecture:** 23 files (≈3,800 lines) copied unchanged into `XLEdge.Addin.Core` (same relative paths and namespaces); they
compile against the 2a facades. Behaviour is pinned by **golden values**: a throwaway harness runs today's 11.1.2 build
(`XLEdge\bin\Debug\XLEdge.dll`, built from the same source — its last commit `ffe8f7f` touched only AddinModule.cs) over input
tables; the tests assert the core's copy gives exactly those outputs (run inside a loader-like AppDomain via `CoreProbe`).

**Tech Stack:** .NET Framework 4.8.1, System.Text.Json 10.0.9 (same version as XLEdge.Shared and XLEdge 11.1.2), MSTest 3.11.1.

**Spec:** `docs/superpowers/specs/2026-10-28-xledge-hot-swappable-core-design.md`

## Global Constraints

- Files are copied verbatim (byte-for-byte apart from the BOM); any change is a ledgered ruling.
- Not in 2b: `ApiOperationHelper` (WinForms/WPF messages → 2c), `DrilldownRequestBuilder` and `XLEdgeParamsBuilder` (Excel +
  report engine → 2d), `LogHelper` (NLog setup — the shell owns logging; not ported), `ProgressCoordinator`, `ExcelWindowHelper`,
  `EnhancedDragDropHelper` (2c/2d).
- 0 errors / 0 warnings, Debug and Release; full suite green; local commits only on `11.1.2-aipowered`; no Claude attribution.
- Any bug or warning found is fixed in this sub-stage, with a test.

## Review Focus

1. A copied file that is not byte-identical to 11.1.2 (silent drift).
2. Golden inputs that do not reach the branches that matter (culture-dependent dates/numbers, empty/null, HTML vs JSON errors).
3. Preferences file location and format: the core must read and write the same file today's add-in uses.
4. HTTP code (APIHelper, ImageDownloadHelper, StrictCertificateValidator) running in the core's AppDomain: TLS setup and
   `ServicePointManager` settings are per AppDomain.
5. Static state in the copied helpers that 11.1.2 relied on living for the whole Excel session (now: per core).

---

### Task 1: The 23 files compile in the core

**Files:** copy `Helpers\{APIHelper, ApiErrorMessageExtractor, XLEdgeValueFormatter, ApiResponseHelper, ImageDownloadHelper,
XLEdgePreferencesManager, NumericJsonConverter, PerformanceHelper, JsonGlobals, CancellationHelper, AttachmentLinkHelper,
ReportParameterRequestSerializer, XLEdgeTempFileCleaner, CancellationTokenHelper, ExceptionHelper, ApiExceptions,
XLEdgeOperatorMappings, EdgeRequestParser, SerializationHelper}.cs`, `Models\{ReportMetaModels, AllModels, ApiResult}.cs`,
`Utilities\StrictCertificateValidator.cs`, `XLEdgeAppConstants.cs`; core csproj gets System.Text.Json and its dependencies.
Test: `XLEdge.Tests\CorePortParityTests.cs` — `PortedFiles_AreByteIdenticalToXLEdge11_1_2` (every file in the 2a+2b port list,
BOM and line endings ignored; the 2a files with ledgered changes are listed with their allowed differences).

- [ ] Test first (fails: files missing); copy; build 0/0; commit.

### Task 2: Golden values pin today's behaviour

**Files:** throwaway harness `golden.ps1` (workspace) → `XLEdge.Tests\GoldenValues.cs` (generated table) and
`XLEdge.Tests\CoreHelperGoldenTests.cs`.

Functions and inputs (each also with null/empty and, where formatting is involved, under en-US, de-DE and th-TH current culture):
`XLEdgeValueFormatter` (FormatDateValue, FormatValue, ToExcelCellValue, InferDrilldownDataType, FormatDrilldownValue,
IsNumeric, RemoveEquaSymbol), `ApiErrorMessageExtractor` (ExtractErrorMessage, IsHtmlResponse — JSON error bodies, HTML pages,
plain text), `AttachmentLinkHelper` (TryParseAttachmentLink, IsEmptyAttachmentAnchor, BuildDownloadUrl — incl. OISR-22767's
empty anchors), `EdgeRequestParser.Parse`, `ReportParameterRequestSerializer.Serialize`, `ExceptionHelper.GetFriendlyErrorMessage`,
`JsonGlobals` getters, `NumericJsonConverter` round trips, `XLEdgeOperatorMappings`.

- [ ] Generate the table from today's DLL; tests assert the core gives the same; mutate a ported function → test fails; commit.

### Task 3: Preferences and HTTP in the core's AppDomain

- [ ] `XLEdgePreferencesManager` reads/writes the same file as 11.1.2 (path from `XLEdgeAppPaths`), and its defaults equal
  today's; a preferences file written by today's DLL loads in the core unchanged.
- [ ] `APIHelper` / `ImageDownloadHelper` against a local `HttpListener`: request headers (Bearer token), timeout and error
  mapping as today; TLS / `ServicePointManager` settings that 11.1.2 set at start-up are applied in the core's domain too.
- [ ] Full suite ×3; mutants; review; fix pass.
