# XLEdgeOfficeJS — Server Configuration Dialog (Sub-Project 1) Design

**Status:** Approved for planning
**Parent effort:** Feasibility exploration of porting XLEdge from a VSTO/COM Excel add-in to an Office.js web add-in. The full port is out of scope for a single spec; this document covers only the first sub-project.

## Goal

Stand up a new repository, `XLEdgeOfficeJS`, and build one real, functional screen — the server configuration window — as an Office.js add-in. The purpose is to answer a concrete feasibility question before committing to a full rewrite: **what does an Office.js dialog actually look and feel like compared to the existing WPF window, and are the gaps (native title bar, no inline-add DataGrid row, sandboxed storage) acceptable?**

This is not a UI mockup — it should behave like the real thing (add/edit/delete/save/set-default, validated), just backed by browser storage instead of the real XML file and add-in host.

## Context: what it's replacing

The current window is `XLEdge\Views\XLEdgeServerConfiguration.xaml` (+ `.xaml.cs`), a borderless (`WindowStyle="None"`) WPF window opened standalone (not docked in a task pane) from the ribbon. It has:
- A custom blue header bar with icon, title, and close button (app-drawn, not OS chrome).
- A `DataGrid` with three columns: Instance Name (text), URL Address (text), Default (checkbox) — including WPF DataGrid's built-in empty "new row" placeholder for inline adds, and per-column tooltips with formatting tips.
- Buttons: GO (connect to the selected/default server), Set as Default, Save, Delete, Close.
- A color-coded status message area (success/error/warning/info) below the grid.
- Validation: Name and Address mandatory, Address must start with `http://`/`https://`, no duplicate names (case-insensitive), exactly one default enforced at all times.
- Persistence to a local XML file (`ORBIT` root, `URL` elements with `Name`/`Address`/`DefaultURL`) at `XLEdgeAppPaths.TempUrlsPath`, with an in-memory cache and a "has unsaved changes" comparison against the last-saved snapshot.

## Decisions (resolved during brainstorming)

| Decision | Choice | Why |
|---|---|---|
| Repo location | `D:\SQLLite_Test\XLEdgeOfficeJS`, sibling to `XLEdge` | Keeps it separate from the VSTO codebase while easy to cross-reference |
| UI surface | Office Dialog API (`Office.context.ui.displayDialogAsync`) | Closest match to today's standalone floating window; a task pane would be a docked-sidebar UX change, not what's being evaluated |
| Scaffold | Yeoman `yo office` generator (Excel Task Pane project, TypeScript) | Official Microsoft tooling — manifest, webpack, dev-cert HTTPS server, and sideload scripts come pre-wired instead of hand-built |
| Manifest type | XML add-in-only manifest | Simpler, no Azure AD/Teams app registration needed for sideload testing; still fully supported for Excel and eventual AppSource submission |
| UI library | Fluent UI React | Authentic native Office look/feel "for free" vs. hand-rolling the current custom styling: this is exactly the visual question the POC exists to answer |
| Functional scope | Fully functional (not a static mockup) | Backed by `localStorage` so add/edit/delete/save/default and validation genuinely work, just not wired to the real add-in/XML file yet |

## Architecture

- `yo office` scaffolds a default Excel Task Pane project: `taskpane.html`/`.ts` (ribbon-launched sidebar) plus `manifest.xml`, webpack config, and dev-cert HTTPS dev server.
- The default task pane gets one addition: a single button, "Open Server Configuration," that calls `Office.context.ui.displayDialogAsync` to launch a second page, `serverconfig.html` / `serverconfig.tsx`, served from the same webpack dev server (same origin as the task pane).
- The dialog page renders the Fluent UI server configuration UI described below inside whatever chrome `displayDialogAsync` gives it.

**Known, expected divergence to validate, not fix:** Office's Dialog API always shows a native OS-drawn title bar; it cannot be suppressed the way WPF's `WindowStyle="None"` is today. The current custom header bar (icon + title + app-drawn close button) becomes content *inside* that native chrome, not a chrome replacement. This is exactly the kind of finding this sub-project exists to surface — no workaround is planned going in.

## Dialog UI components

Mirrors the current window's behavior, rebuilt with Fluent UI React:

- **Grid**: Fluent `DetailsList` (or `DataGrid`, whichever gives cleaner inline cell editing) with editable Name and URL Address columns and a Default checkbox column. Fluent grids have no built-in "empty new-row" placeholder like WPF's `DataGrid`, so an explicit **"+ Add Row"** button drives adding a new instance instead.
- **Buttons**: GO, Set as Default, Save, Delete, Close — same set and same intent as today.
- **Status area**: single-line, color-coded (success/error/warning/info) message region below the grid, same four states as `StatusMessageType` in the current code.
- **Validation**, ported as-is from `SaveConfiguration`/`DgInstances_CellEditEnding`:
  - Name and Address are mandatory for every row.
  - Address must start with `http://` or `https://`.
  - Name must be unique, case-insensitive.
  - Exactly one instance is marked default at all times (auto-assigned to the first instance if none is set; forced to a single instance if more than one is somehow set).

## Data

- Persisted to the dialog page's `localStorage` under a single key (e.g. `xledge.serverInstances`), as a JSON array of `{ name, address, isDefault }`.
- No backend, no roaming settings, no communication back to the task pane in this sub-project — the dialog is self-contained. Real persistence strategy (Office roaming settings vs. a backend service vs. something else) is an explicit **out-of-scope** decision for a later sub-project, once the visual/UX question this spec targets is answered.

## Testing / verification

- `npm start` runs the webpack dev server and sideloads the add-in into Excel desktop.
- Manual pass covering: add a row, edit Name/Address inline, mark a row default (and confirm only one stays default), delete a row (including the currently-default row), save, and reopen the dialog to confirm `localStorage` state survived.
- Manual pass covering each validation rule (empty Name, empty Address, bad URL scheme, duplicate name) and confirming the status area shows the right message/color.
- No automated test suite planned for this sub-project — it's a feasibility/visual POC, not production code.

## Out of scope (explicitly deferred)

- Wiring the dialog to real server data, the real XML config file, or the rest of the XLEdge add-in.
- Any decision about the long-term persistence mechanism for a production Office.js version.
- Porting any other XLEdge window (About, Login, Options, GL Accounts, Drilldown, ribbon).
- Any decision about whether to actually proceed with the full VSTO → Office.js rewrite — this sub-project's output is an answer to that question, not a commitment.
