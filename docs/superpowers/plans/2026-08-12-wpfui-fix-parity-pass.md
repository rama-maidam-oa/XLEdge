# XLEdge WPF-UI Fix-Parity Pass Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Port the regression fixes GLSense's MahApps.Metro → WPF-UI migration needed (ToolTip lazy-resolution crash, DataGrid resize gap, message-window row collapse, bulk-write Excel freeze) into XLEdge, applying only the ones whose root-cause condition actually exists in XLEdge's code.

**Architecture:** This is NOT a base-window migration — XLEdge's `Utilities\DpiAwareWindow.cs` already derives from `Wpf.Ui.Controls.FluentWindow` and `MahApps.Metro` (core) has never existed in this repo's history (confirmed via `git log -p -G'"MahApps\.Metro"'` returning zero hits). Only the icon-only `MahApps.Metro.IconPacks.Core`/`.FontAwesome` packages remain, by the same deliberate design GLSense settled on. This plan instead audits XLEdge against each regression GLSense's migration (`GLSense` repo, branch `11.1.0_NewUI`, commits `65f4dee`, `eaa1b94`, `f7fe334`, `de39e31`, `b43938a`, `2cc7d65`) had to fix afterward, and ports forward only what's actually broken here.

**Tech Stack:** WPF (.NET Framework 4.8.1), Wpf.Ui 4.3.0, Add-in Express (Excel VSTO-style add-in), no local MSBuild/Windows build toolchain available in this environment (matches every prior fix in `MIGRATION_STATUS.md` — verified by brace/tag-balance and grep evidence, not a real compile, until the user builds it).

## Global Constraints

- No build toolchain available here — every change must be verified by direct inspection (grep/read) of the edited XAML/C#, not by compiling. Call this out explicitly when reporting each step's result.
- Do not touch `Utilities\DpiAwareWindow.cs`, base-class wiring, or any window's root element/chrome — that migration is already complete and out of scope.
- Do not modify `MahApps.Metro.IconPacks` usage (`iconPacks:PackIconFontAwesome`) — kept intentionally, matching GLSense's final state.
- Match this repo's own fix-documentation convention: after implementation, append an entry to `MIGRATION_STATUS.md` (not a new doc) describing what was found and fixed, in the same style as its existing entries.

---

## Investigation findings (already completed — recorded here so no step re-derives them)

Audited each of GLSense's 4 post-migration regression classes against XLEdge's actual code:

| GLSense regression | Root cause condition | Present in XLEdge? | Verdict |
|---|---|---|---|
| ToolTip `ResourceReferenceKeyNotFoundException` on hover | A `ToolTip`'s own `Style="{StaticResource Name}"` resolves lazily at popup-open time, not load time; the lookup intermittently fails in this VSTO host | `Themes\GlobalStyles.xaml` still declares `SimpleBrowserToolTip` (line 93) and `ChromeStyleToolTip` (line 105) as **named** (`x:Key`) styles, referenced by `Style="{StaticResource ...}"` at ~32 call sites (19 in `Views\XLEdgeAbout.xaml` / `XLEdgeLoginDetails.xaml` / `XLEdgeOptions.xaml` / `XLEdgeServerConfiguration.xaml`, 13 more inside `GlobalStyles.xaml` itself, 7 of those specifically `ChromeStyleToolTip` in `XLEdgeOptions.xaml`) | **PRESENT — needs the fix (Task 1 + Task 2)** |
| DataGrid trailing-empty-column on resize | `SizeToContent="WidthAndHeight"` measures a genuine `Width="*"` column against `Infinity`, forcing an `Auto`+fixed-pixel workaround that can't self-balance on manual resize | Every XLEdge window uses `SizeToContent="Manual"` (confirmed via `grep -n SizeToContent Views/*.xaml` — `XLEdgeMessageWindow.xaml` even has an inline comment explaining `WidthAndHeight` was deliberately rejected for exactly this class of self-fighting layout). Both DataGrids (`XLEdgeAbout.xaml` `dgInstances`, `XLEdgeServerConfiguration.xaml` `dgInstances`) already use genuine `Width="*"`/star-ratio fill columns, never `Auto` | **NOT PRESENT — no code change; Task 3 is a documentation-only note, not a fix** |
| Message-window row collapsing to 0 height under `SizeToContent` | A bare `RowDefinition Height="*"` content row measures to 0 during `SizeToContent`'s infinite-availSize pass | `XLEdgeMessageWindow.xaml` already uses `Auto`+`Auto` rows under `SizeToContent="Manual"`, with an existing comment documenting this was chosen specifically to avoid that bug | **NOT PRESENT — no action** |
| Bulk-write Excel freeze (`O(n²)` recalculation storm) + re-entrant submit click | A WPF button triggers a long loop of individual cell writes with `Calculation` left on `Automatic` and no re-entrancy guard | XLEdge's only comparable write path, `Views\XLEdgeGLAccountsWindow.xaml.cs BtnOk_Click`, writes at most 2 cells (trivial, no loop). The real bulk-write path, report generation (`Helpers\ReportGenerator.cs`), is already wrapped in `ExcelBulkOperationScope` which sets `Calculation = xlCalculationManual` for the whole operation and restores it on dispose | **NOT PRESENT — no action** |

Net result: **one real, unfixed bug class** (ToolTip lazy-resolution crash risk) exists in XLEdge today. The other three are structurally impossible given choices already made when this codebase was built (all discovered via direct grep/read evidence in this plan's investigation, not assumption).

---

### Task 1: Make `SimpleBrowserToolTip` implicit and update its call sites

**Files:**
- Modify: `XLEdge\Themes\GlobalStyles.xaml:93-102` (style definition)
- Modify: `XLEdge\Themes\GlobalStyles.xaml` (13 internal call sites: lines 293, 307, 320, 380, 427, 474, 537, 584, 631, 652, 673, 694, 1733)
- Modify: `XLEdge\Views\XLEdgeAbout.xaml` (8 call sites: lines 117, 125, 133, 142, 149, 162, 252, 266)
- Modify: `XLEdge\Views\XLEdgeLoginDetails.xaml` (2 call sites: lines 100, 138)
- Modify: `XLEdge\Views\XLEdgeServerConfiguration.xaml` (10 call sites: lines 146, 171, 184, 209, 250, 303, 311, 319, 327, 335)

**Interfaces:**
- Consumes: nothing from other tasks.
- Produces: an implicit `<Style TargetType="ToolTip">` (no `x:Key`) in `GlobalStyles.xaml`, which every plain `<ToolTip>` in the app now picks up automatically. Task 2 depends on this style occupying the app's one allowed implicit `ToolTip` style, which is why `ChromeStyleToolTip` (Task 2) cannot also be made implicit and must be inlined instead.

- [ ] **Step 1: Convert the `SimpleBrowserToolTip` style to implicit**

In `XLEdge\Themes\GlobalStyles.xaml`, replace:
```xml
	<!-- ==================== ToolTip Styles ==================== -->
	<!-- Simple Browser ToolTip (yellow background) -->
	<Style x:Key="SimpleBrowserToolTip" TargetType="ToolTip">
		<Setter Property="Background" Value="{StaticResource TooltipBackgroundBrush}"/>
		<Setter Property="Foreground" Value="#3C3C3C"/>
		<Setter Property="BorderBrush" Value="{StaticResource TooltipBorderBrush}"/>
		<Setter Property="BorderThickness" Value="1"/>
		<Setter Property="Padding" Value="10,8"/>
		<Setter Property="FontSize" Value="12"/>
		<Setter Property="FontFamily" Value="Segoe UI"/>
		<Setter Property="HasDropShadow" Value="True"/>
	</Style>
```
with:
```xml
	<!-- ==================== ToolTip Styles ====================
	     No x:Key - implicit/default style for every ToolTip in any window that merges this
	     dictionary. Previously had x:Key="SimpleBrowserToolTip" and every window referenced it
	     explicitly via Style="{StaticResource SimpleBrowserToolTip}" - but a ToolTip's content
	     (including its own Style property) is deferred and only resolved when the tooltip
	     popup actually opens on hover, not when the page/window is loaded. In this app's VSTO
	     hosting context that deferred StaticResource lookup can intermittently fail to find the
	     key and throw ResourceReferenceKeyNotFoundException on hover (confirmed root cause of an
	     identical crash across multiple, unrelated windows in the sibling GLSense project's own
	     WPF-UI migration - see GLSense repo, branch 11.1.0_NewUI, commit f7fe334). An implicit
	     style is looked up the same way, but a failed lookup silently falls back to the default
	     ToolTip appearance instead of throwing, so this can never crash on hover. -->
	<Style TargetType="ToolTip">
		<Setter Property="Background" Value="{StaticResource TooltipBackgroundBrush}"/>
		<Setter Property="Foreground" Value="#3C3C3C"/>
		<Setter Property="BorderBrush" Value="{StaticResource TooltipBorderBrush}"/>
		<Setter Property="BorderThickness" Value="1"/>
		<Setter Property="Padding" Value="10,8"/>
		<Setter Property="FontSize" Value="12"/>
		<Setter Property="FontFamily" Value="Segoe UI"/>
		<Setter Property="HasDropShadow" Value="True"/>
	</Style>
```

- [ ] **Step 2: Verify no other implicit `TargetType="ToolTip"` style exists yet**

Run: `grep -n 'TargetType="ToolTip"' XLEdge/Themes/GlobalStyles.xaml`
Expected: exactly 3 matches at this point — the now-implicit style from Step 1, `WarningToolTipStyle` (still named, unused, untouched — leave it), and `ChromeStyleToolTip` (still named, will be removed in Task 2). No second bare (implicit) `TargetType="ToolTip"` — WPF only allows one per dictionary scope, and a second one here would be a silent authoring mistake, not a runtime error, so this check is required rather than optional.

- [ ] **Step 3: Replace every `Style="{StaticResource SimpleBrowserToolTip}"` call site with a bare `<ToolTip>`**

In each of the 4 files listed above, replace every occurrence of:
```xml
<ToolTip Style="{StaticResource SimpleBrowserToolTip}">
```
with:
```xml
<ToolTip>
```
Use `replace_all` within each file (the string is identical at every call site in that file).

- [ ] **Step 4: Verify zero remaining references**

Run: `grep -rn "SimpleBrowserToolTip" XLEdge/`
Expected: zero matches anywhere (definition and all call sites removed/renamed).

- [ ] **Step 5: Verify XML well-formedness of every edited file**

No build toolchain is available here (see Global Constraints) — verify with a balanced-tag check instead:
Run (PowerShell): `[xml](Get-Content "XLEdge\Themes\GlobalStyles.xaml" -Raw)` and repeat for `XLEdgeAbout.xaml`, `XLEdgeLoginDetails.xaml`, `XLEdgeServerConfiguration.xaml`.
Expected: each command returns the parsed XML object with no exception. An exception here means a `<ToolTip>`/`</ToolTip>` tag mismatch was introduced and must be fixed before continuing.

- [ ] **Step 6: Commit**

```bash
git add XLEdge/Themes/GlobalStyles.xaml XLEdge/Views/XLEdgeAbout.xaml XLEdge/Views/XLEdgeLoginDetails.xaml XLEdge/Views/XLEdgeServerConfiguration.xaml
git commit -m "Fix ToolTip ResourceReferenceKeyNotFoundException risk: make SimpleBrowserToolTip implicit"
```

---

### Task 2: Inline `ChromeStyleToolTip`'s template at its 7 call sites, remove the named style

**Files:**
- Modify: `XLEdge\Views\XLEdgeOptions.xaml:88,105,122,139,155,171,187` (7 call sites)
- Modify: `XLEdge\Themes\GlobalStyles.xaml:104-159` (remove the named style, now dead)

**Interfaces:**
- Consumes: Task 1's implicit `TargetType="ToolTip"` style (the reason `ChromeStyleToolTip` can't also be implicit — WPF allows only one implicit style per `TargetType` per merged-dictionary scope).
- Produces: nothing consumed by later tasks.

- [ ] **Step 1: Read one call site to confirm current structure**

Run: `grep -n -A2 'ChromeStyleToolTip' XLEdge/Views/XLEdgeOptions.xaml` and open the file around line 88 to see the exact `<ToolTip Style="{StaticResource ChromeStyleToolTip}">...</ToolTip>` block and its inner content (a `TextBlock` or similar, per each of the 7 sites — content differs per site, only the wrapping `<ToolTip>` tag/style changes).

- [ ] **Step 2: Replace each of the 7 `<ToolTip Style="{StaticResource ChromeStyleToolTip}">` opening tags**

For each of the 7 call sites in `XLEdge\Views\XLEdgeOptions.xaml`, change only the opening tag from:
```xml
<ToolTip Style="{StaticResource ChromeStyleToolTip}">
```
to:
```xml
<ToolTip Background="Transparent" BorderBrush="Transparent" BorderThickness="0" Padding="0" HasDropShadow="True">
    <!-- ChromeStyleToolTip's Style/Template inlined directly rather than referenced via
         StaticResource: a ToolTip's own Style resolves lazily when the popup actually opens
         on hover, not at load time, and that deferred lookup can intermittently fail to find
         a named resource (see SimpleBrowserToolTip's identical issue in Task 1 - not an
         option here since this template must stay visually distinct from the app-wide
         default). -->
    <ToolTip.Template>
        <ControlTemplate TargetType="ToolTip">
            <Border Background="{StaticResource TooltipBackgroundBrush}"
                    BorderBrush="{StaticResource TooltipBorderBrush}"
                    BorderThickness="1"
                    CornerRadius="4"
                    MaxWidth="400">
                <Border.Effect>
                    <DropShadowEffect Color="Black" Opacity="0.15" BlurRadius="8" ShadowDepth="2" Direction="315"/>
                </Border.Effect>
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="{StaticResource TooltipBorderBrush}" CornerRadius="4,4,0,0" Padding="10,6">
                        <StackPanel Orientation="Horizontal">
                            <iconPacks:PackIconFontAwesome Kind="LightbulbRegular" Width="12" Height="12" Foreground="Black" Margin="0,0,6,0" VerticalAlignment="Center"/>
                            <TextBlock Text="TIP" FontWeight="Bold" FontSize="11" Foreground="Black" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>
                    <Border Grid.Row="1" Padding="12,10">
                        <ContentPresenter ContentSource="Content"/>
                    </Border>
                </Grid>
            </Border>
        </ControlTemplate>
    </ToolTip.Template>
```
and add the matching extra closing `</ToolTip.Template>` is already included above — only the existing inner content (whatever `TextBlock`/etc. was already inside each `<ToolTip>...</ToolTip>` pair) and the existing `</ToolTip>` closing tag stay exactly as they were; only the opening tag grows into this block. Do this individually for each of the 7 sites (their inner content differs, so this cannot be a single `replace_all`).

- [ ] **Step 3: Remove the now-dead `ChromeStyleToolTip` named style from `GlobalStyles.xaml`**

Delete the entire style block (`XLEdge\Themes\GlobalStyles.xaml:104-159`, the `<!-- Chrome-style ToolTip WITH Header -->` comment through its closing `</Style>`).

- [ ] **Step 4: Verify zero remaining references**

Run: `grep -rn "ChromeStyleToolTip" XLEdge/`
Expected: zero matches.

- [ ] **Step 5: Verify XML well-formedness**

Run (PowerShell): `[xml](Get-Content "XLEdge\Views\XLEdgeOptions.xaml" -Raw)` and `[xml](Get-Content "XLEdge\Themes\GlobalStyles.xaml" -Raw)`.
Expected: both parse with no exception.

- [ ] **Step 6: Commit**

```bash
git add XLEdge/Views/XLEdgeOptions.xaml XLEdge/Themes/GlobalStyles.xaml
git commit -m "Inline ChromeStyleToolTip template at its call sites, remove the now-unimplicit-incompatible named style"
```

---

### Task 3: Record the DataGrid/message-window/busy-overlay audit result in MIGRATION_STATUS.md

**Files:**
- Modify: `XLEdge\MIGRATION_STATUS.md` (prepend a new dated entry, matching the file's existing style)

**Interfaces:**
- Consumes: the investigation findings table above.
- Produces: nothing (documentation only).

- [ ] **Step 1: Add a dated entry**

Prepend (after the file's `# ... Status & Reference` header, before the existing most-recent entry) a new section titled `## WPF-UI fix-parity audit against GLSense's migration — 2026-08-12`, summarizing: (a) confirmation that XLEdge's `DpiAwareWindow` already derives from `FluentWindow` and `MahApps.Metro` core has never been present in this repo's history, (b) the ToolTip fix applied in Tasks 1-2, (c) the three other regression classes checked and found not applicable, with the one-line reason each (SizeToContent="Manual" everywhere, both DataGrids already star-sized, XLEdgeGLAccountsWindow's write is a 2-cell write not a bulk loop, report generation already uses ExcelBulkOperationScope's Calculation=Manual wrapper) - use the investigation-findings table above as the source content, written in this file's established prose style (see any existing entry for tone/structure).

- [ ] **Step 2: Commit**

```bash
git add XLEdge/MIGRATION_STATUS.md
git commit -m "Document WPF-UI fix-parity audit against GLSense's migration"
```

---

## Self-review notes

- Spec coverage: all 4 of GLSense's regression classes are addressed — 2 fixed (Tasks 1-2), 2 documented as not-applicable with evidence (Task 3). The user's explicit ask to "check DGV fixes in detail" is covered by the investigation table + Task 3's write-up, even though it results in no code change (the DataGrid layout was already safe).
- No base-window/chrome work is included — confirmed out of scope per the user's own decision after seeing the investigation.
- `WarningToolTipStyle` is deliberately left untouched (unused, non-colliding, no bug).
