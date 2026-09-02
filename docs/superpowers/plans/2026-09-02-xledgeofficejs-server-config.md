# XLEdgeOfficeJS Server Configuration Dialog Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Scaffold a new `XLEdgeOfficeJS` repository and build a fully functional Server Configuration dialog as an Office.js Excel add-in, to answer the feasibility question of whether Office.js can visually and functionally replace XLEdge's current WPF `XLEdgeServerConfiguration` window.

**Architecture:** A Yeoman-scaffolded Office Add-in (Excel Task Pane, React, TypeScript, XML manifest). The default task pane gets one button that opens a second page — the Server Configuration screen — via the Office Dialog API (`Office.context.ui.displayDialogAsync`). That dialog page is a Fluent UI React (v9, `@fluentui/react-components`) screen backed by `localStorage`, replicating the current WPF window's grid, validation, and button behavior.

**Tech Stack:** Node.js 24 LTS (via nvm-windows, `C:\nvm4w\nodejs`), TypeScript 5.4, React 18, webpack 5, `@fluentui/react-components` v9, `office-addin-debugging` for sideload.

**Spec:** `docs/superpowers/specs/2026-09-02-xledgeofficejs-server-config-design.md`

## Global Constraints

- **Node version blocker (resolved this session):** this machine's default Node was v14.17.3 (EOL, incompatible with `yo`/`generator-office`/current `typescript`/`webpack-dev-server`). `nvm-windows` (`CoreyButler.NVMforWindows`) is now installed and Node 24.20.0 is active via the symlink at `C:\nvm4w\nodejs`. **Every command in every task that invokes `node`, `npm`, or `npx` must prepend `export PATH="/c/nvm4w/nodejs:$PATH"` (bash) or `$env:Path = "C:\nvm4w\nodejs;$env:Path"` (PowerShell) first** — the host shell's own cached PATH predates this Node install and will not resolve `node`/`npm` otherwise. A brand-new terminal window opened outside this session does not need this (the machine-wide `NVM_HOME`/`NVM_SYMLINK` env vars are set), but assume this session's shells need it.
- Repo root for all tasks: `D:\SQLLite_Test\XLEdgeOfficeJS` (verified with `yo office` — see Task 1).
- UI library is Fluent UI React **v9** (`@fluentui/react-components`) only — this is what the generator installs. Do not use `@fluentui/react` (v8) or `DetailsList`/`DetailsList`-family components; they are a different, uninstalled package. Editable grid rows use the v9 `Table`/`TableRow`/`TableCell` primitives with `Input`/`Checkbox` inside cells (verified available in the installed package — v9's `DataGrid` does not support inline cell editing).
- Data persists to browser `localStorage` only (key `xledge.serverInstances`) — no backend, no roaming settings, no wiring to the real XLEdge add-in or its XML file. This is explicit, spec-approved scope, not a shortcut to fix later inside this plan.
- No automated test suite (spec-approved: "feasibility/visual POC, not production code"). Each task's gate is `npm run build` (a real TypeScript + webpack compile, catches real errors) rather than automated tests. Full interactive verification against Excel desktop happens once, in Task 6, covering everything built in Tasks 2–5 together — sideloading into Excel for every intermediate task would be far more expensive than the signal it adds.
- Manifest type is the XML add-in-only manifest (`manifest.xml`) — the generator also produces a `manifest.json`; leave it in place (it's an intermediate build artifact of the generator's `convert-to-single-host` step) but do not use or reference it.

---

## Investigation findings (already completed — recorded here so no task re-derives them)

- Verified in this environment: `yo office react <Name> excel xml` (positional args: projectType, name, host, manifestType) scaffolds a working TypeScript + React + Fluent UI v9 Excel task-pane project. Confirmed via a real scaffold + `npm install` + `npm run build` in a scratch directory — compiled successfully with webpack 5.110.3.
- Confirmed installed dependencies include `@fluentui/react-components@^9.55.1` with `Table`, `TableHeader`, `TableHeaderCell`, `TableBody`, `TableRow`, `TableCell`, `Input`, `Checkbox`, `Button`, `MessageBar`, `MessageBarBody`, `Title2`, `tokens`, `makeStyles` all present as real exports (checked via `require()` against the installed package, not assumed from docs).
- Confirmed generated file layout: `src/taskpane/{index.tsx,taskpane.html,taskpane.ts,components/{App,Header,HeroList,TextInsertion}.tsx}`, `src/commands/{commands.html,commands.ts}`, `webpack.config.js` (entry/plugin pattern for HTML pages), `tsconfig.json`, `manifest.xml`. `package.json` includes `build`, `start` (`office-addin-debugging start manifest.xml`), `validate` (`office-addin-manifest validate manifest.xml`) scripts already wired.
- No `.gitignore` is generated — Task 1 hand-writes one.

---

### Task 1: Scaffold the XLEdgeOfficeJS project and initialize the repo

**Files:**
- Create: `D:\SQLLite_Test\XLEdgeOfficeJS\` (entire generated project — package.json, manifest.xml, manifest.json, webpack.config.js, tsconfig.json, babel.config.json, .eslintrc.json, .hintrc, .vscode/, assets/, src/taskpane/, src/commands/)
- Create: `D:\SQLLite_Test\XLEdgeOfficeJS\.gitignore`

**Interfaces:**
- Consumes: nothing.
- Produces: the base project all later tasks modify. Later tasks assume `npm run build` and `npm start` work exactly as the generator wires them (see `package.json` scripts above).

- [ ] **Step 1: Run the Yeoman Office generator**

From `D:\SQLLite_Test`:

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test
npx --yes -p yo -p generator-office yo office react XLEdgeOfficeJS excel xml --skip-cache
```

This creates `D:\SQLLite_Test\XLEdgeOfficeJS\` (name arg becomes the output folder since `-o` isn't passed), running `npm install` automatically (default `--skip-install` is `false`) — expect this step to take several minutes.

- [ ] **Step 2: Verify the generated project builds**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm run build
```

Expected: ends with `webpack 5.x.x compiled successfully`.

- [ ] **Step 3: Add a `.gitignore`**

Create `D:\SQLLite_Test\XLEdgeOfficeJS\.gitignore`:

```gitignore
node_modules/
dist/
*.log
.DS_Store
```

- [ ] **Step 4: Initialize git and make the first commit**

```bash
cd /d/SQLLite_Test/XLEdgeOfficeJS
git init
git add -A
git commit -m "Scaffold XLEdgeOfficeJS via yo office (React, TypeScript, Excel, XML manifest)"
```

---

### Task 2: Wire up the Server Configuration dialog plumbing (open/close round trip)

**Files:**
- Create: `D:\SQLLite_Test\XLEdgeOfficeJS\src\serverconfig\serverconfig.html`
- Create: `D:\SQLLite_Test\XLEdgeOfficeJS\src\serverconfig\index.tsx`
- Create: `D:\SQLLite_Test\XLEdgeOfficeJS\src\serverconfig\components\ServerConfigDialog.tsx`
- Modify: `D:\SQLLite_Test\XLEdgeOfficeJS\webpack.config.js`
- Modify: `D:\SQLLite_Test\XLEdgeOfficeJS\src\taskpane\components\App.tsx` (full rewrite)
- Delete: `D:\SQLLite_Test\XLEdgeOfficeJS\src\taskpane\components\HeroList.tsx`
- Delete: `D:\SQLLite_Test\XLEdgeOfficeJS\src\taskpane\components\TextInsertion.tsx`

**Interfaces:**
- Consumes: nothing from Task 1 beyond the scaffolded project structure.
- Produces: `ServerConfigDialog` (default export, no props yet) — Task 3 rewrites this file's internals but keeps the same default-export shape and the `closeDialog()` message contract (`{ type: "close" }` sent via `Office.context.ui.messageParent`), which `App.tsx`'s `handleDialogMessage` in this task already depends on.

This task proves the core feasibility mechanism end-to-end — a task-pane button opening a real Office dialog showing Fluent UI content — before any grid logic is built, since that visual/UX question (native title bar, Fluent look) is the whole point of this sub-project.

- [ ] **Step 1: Create the dialog page shell**

Create `src/serverconfig/serverconfig.html`:

```html
<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license. -->
<!-- See LICENSE in the project root for license information -->

<!doctype html>
<html lang="en" data-framework="typescript">

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Server Configuration</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
</head>

<body style="width: 100%; height: 100%; margin: 0; padding: 0;">
    <div id="container"></div>
</body>

</html>
```

- [ ] **Step 2: Create the dialog page's React entry point**

Create `src/serverconfig/index.tsx`:

```tsx
import * as React from "react";
import { createRoot } from "react-dom/client";
import ServerConfigDialog from "./components/ServerConfigDialog";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <ServerConfigDialog />
    </FluentProvider>
  );
});
```

- [ ] **Step 3: Create a minimal ServerConfigDialog (header + Close only)**

Create `src/serverconfig/components/ServerConfigDialog.tsx`:

```tsx
import * as React from "react";
import { Button, Title2, makeStyles, tokens } from "@fluentui/react-components";

/* global Office */

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    fontFamily: tokens.fontFamilyBase,
  },
  header: {
    padding: "12px 16px",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  footer: {
    padding: "12px 16px",
    display: "flex",
    justifyContent: "flex-end",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
});

function closeDialog(): void {
  Office.context.ui.messageParent(JSON.stringify({ type: "close" }));
}

const ServerConfigDialog: React.FC = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Title2>Server Configuration</Title2>
      </div>
      <div className={styles.footer}>
        <Button appearance="secondary" onClick={closeDialog}>
          Close
        </Button>
      </div>
    </div>
  );
};

export default ServerConfigDialog;
```

- [ ] **Step 4: Register `serverconfig.html` as a webpack entry/output page**

In `webpack.config.js`, change the `entry` block from:

```js
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      react: ["react", "react-dom"],
      taskpane: {
        import: ["./src/taskpane/index.tsx", "./src/taskpane/taskpane.html"],
        dependOn: "react",
      },
      commands: "./src/commands/commands.ts",
    },
```

to:

```js
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      react: ["react", "react-dom"],
      taskpane: {
        import: ["./src/taskpane/index.tsx", "./src/taskpane/taskpane.html"],
        dependOn: "react",
      },
      serverconfig: {
        import: ["./src/serverconfig/index.tsx", "./src/serverconfig/serverconfig.html"],
        dependOn: "react",
      },
      commands: "./src/commands/commands.ts",
    },
```

And add a matching `HtmlWebpackPlugin` entry to the `plugins` array, right after the existing `taskpane.html` one:

```js
      new HtmlWebpackPlugin({
        filename: "serverconfig.html",
        template: "./src/serverconfig/serverconfig.html",
        chunks: ["polyfill", "serverconfig", "react"],
      }),
```

- [ ] **Step 5: Replace App.tsx with the "Open Server Configuration" button**

Replace the full contents of `src/taskpane/components/App.tsx` with:

```tsx
import * as React from "react";
import { useRef } from "react";
import Header from "./Header";
import { Button, makeStyles } from "@fluentui/react-components";

/* global Office, window */

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
  content: {
    padding: "16px",
  },
});

function handleDialogMessage(
  dialogRef: React.MutableRefObject<Office.Dialog | undefined>,
  arg: Office.DialogParentMessageReceivedEventArgs
): void {
  const message = JSON.parse(arg.message as string);
  if (message.type === "close") {
    dialogRef.current?.close();
    dialogRef.current = undefined;
  }
}

function handleDialogEvent(arg: Office.DialogEventReceivedEventArgs): void {
  console.log(`Server Configuration dialog closed with event ${arg.error}`);
}

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  const dialogRef = useRef<Office.Dialog | undefined>(undefined);

  const openServerConfigDialog = () => {
    const dialogUrl = `${window.location.origin}/serverconfig.html`;
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 60, width: 40, displayInIframe: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(`Dialog failed to open: ${asyncResult.error.message}`);
          return;
        }
        dialogRef.current = asyncResult.value;
        dialogRef.current.addEventHandler(Office.EventType.DialogMessageReceived, (arg) =>
          handleDialogMessage(dialogRef, arg as Office.DialogParentMessageReceivedEventArgs)
        );
        dialogRef.current.addEventHandler(Office.EventType.DialogEventReceived, (arg) =>
          handleDialogEvent(arg as Office.DialogEventReceivedEventArgs)
        );
      }
    );
  };

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="XLEdge Office.js POC" />
      <div className={styles.content}>
        <Button appearance="primary" onClick={openServerConfigDialog}>
          Open Server Configuration
        </Button>
      </div>
    </div>
  );
};

export default App;
```

- [ ] **Step 6: Delete the now-unused demo components**

```bash
cd /d/SQLLite_Test/XLEdgeOfficeJS
rm src/taskpane/components/HeroList.tsx src/taskpane/components/TextInsertion.tsx
```

- [ ] **Step 7: Verify it compiles**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm run build
```

Expected: `webpack 5.x.x compiled successfully`, with `serverconfig.html`/`serverconfig.js` now listed among the emitted assets alongside `taskpane.html`/`taskpane.js`.

- [ ] **Step 8: Commit**

```bash
cd /d/SQLLite_Test/XLEdgeOfficeJS
git add -A
git commit -m "Add Server Configuration dialog plumbing (open/close round trip)"
```

---

### Task 3: Data module + editable grid (in-memory)

**Files:**
- Create: `D:\SQLLite_Test\XLEdgeOfficeJS\src\serverconfig\serverConfigStore.ts`
- Modify: `D:\SQLLite_Test\XLEdgeOfficeJS\src\serverconfig\components\ServerConfigDialog.tsx` (full rewrite)

**Interfaces:**
- Consumes: nothing new from Task 2 besides the file it's replacing.
- Produces: `ServerInstance` type, `createEmptyInstance()`, `loadServerInstances()`, `saveServerInstances(instances)`, `validateServerInstances(instances)` (returns `{ isValid: boolean; message?: string }`), `ensureSingleDefault(instances)` — all from `serverConfigStore.ts`. Task 4 consumes `saveServerInstances`, `validateServerInstances`, `ensureSingleDefault` by these exact names/signatures.

- [ ] **Step 1: Create the data/validation module**

Create `src/serverconfig/serverConfigStore.ts`:

```ts
export interface ServerInstance {
  id: string;
  name: string;
  address: string;
  isDefault: boolean;
}

const STORAGE_KEY = "xledge.serverInstances";

export function loadServerInstances(): ServerInstance[] {
  const raw = window.localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return [];
  }
  try {
    const parsed = JSON.parse(raw) as ServerInstance[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function saveServerInstances(instances: ServerInstance[]): void {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(instances));
}

export interface ValidationResult {
  isValid: boolean;
  message?: string;
}

export function validateServerInstances(instances: ServerInstance[]): ValidationResult {
  const missingMandatory = instances.some((instance) => !instance.name.trim() || !instance.address.trim());
  if (missingMandatory) {
    return { isValid: false, message: "Instance Name and URL Address are mandatory for all entries." };
  }

  const invalidUrl = instances.find(
    (instance) => !instance.address.startsWith("http://") && !instance.address.startsWith("https://")
  );
  if (invalidUrl) {
    return { isValid: false, message: "URL Address must start with http:// or https://" };
  }

  const namesLower = instances.map((instance) => instance.name.trim().toLowerCase());
  const duplicate = namesLower.find((name, index) => namesLower.indexOf(name) !== index);
  if (duplicate) {
    return { isValid: false, message: `Instance Name '${duplicate}' already exists. Please use a unique name.` };
  }

  return { isValid: true };
}

export function ensureSingleDefault(instances: ServerInstance[]): ServerInstance[] {
  if (instances.length === 0) {
    return instances;
  }
  const hasDefault = instances.some((instance) => instance.isDefault);
  if (!hasDefault) {
    return instances.map((instance, index) => ({ ...instance, isDefault: index === 0 }));
  }
  let seenDefault = false;
  return instances.map((instance) => {
    if (instance.isDefault) {
      if (seenDefault) {
        return { ...instance, isDefault: false };
      }
      seenDefault = true;
    }
    return instance;
  });
}

export function createEmptyInstance(): ServerInstance {
  return {
    id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
    name: "",
    address: "",
    isDefault: false,
  };
}
```

- [ ] **Step 2: Replace ServerConfigDialog.tsx with the grid, wired to in-memory state**

Replace the full contents of `src/serverconfig/components/ServerConfigDialog.tsx` with:

```tsx
import * as React from "react";
import { useState } from "react";
import {
  Button,
  Title2,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  Input,
  Checkbox,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { ServerInstance, createEmptyInstance, loadServerInstances } from "../serverConfigStore";

/* global Office */

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    fontFamily: tokens.fontFamilyBase,
  },
  header: {
    padding: "12px 16px",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  toolbar: {
    padding: "8px 16px",
    display: "flex",
    gap: "8px",
  },
  body: {
    flex: 1,
    padding: "12px 16px",
    overflow: "auto",
  },
  footer: {
    padding: "12px 16px",
    display: "flex",
    justifyContent: "flex-end",
    gap: "8px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
});

function closeDialog(): void {
  Office.context.ui.messageParent(JSON.stringify({ type: "close" }));
}

const ServerConfigDialog: React.FC = () => {
  const styles = useStyles();
  const [instances, setInstances] = useState<ServerInstance[]>(() => loadServerInstances());

  const updateInstance = (id: string, changes: Partial<ServerInstance>) => {
    setInstances((current) =>
      current.map((instance) => (instance.id === id ? { ...instance, ...changes } : instance))
    );
  };

  const addRow = () => {
    setInstances((current) => [...current, createEmptyInstance()]);
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Title2>Server Configuration</Title2>
      </div>
      <div className={styles.toolbar}>
        <Button appearance="secondary" onClick={addRow}>
          + Add Row
        </Button>
      </div>
      <div className={styles.body}>
        <Table aria-label="Server instances">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Instance Name</TableHeaderCell>
              <TableHeaderCell>URL Address</TableHeaderCell>
              <TableHeaderCell>Default</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {instances.map((instance) => (
              <TableRow key={instance.id}>
                <TableCell>
                  <Input
                    value={instance.name}
                    onChange={(_, data) => updateInstance(instance.id, { name: data.value })}
                  />
                </TableCell>
                <TableCell>
                  <Input
                    value={instance.address}
                    onChange={(_, data) => updateInstance(instance.id, { address: data.value })}
                  />
                </TableCell>
                <TableCell>
                  <Checkbox
                    checked={instance.isDefault}
                    onChange={(_, data) => updateInstance(instance.id, { isDefault: data.checked === true })}
                  />
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>
      <div className={styles.footer}>
        <Button appearance="secondary" onClick={closeDialog}>
          Close
        </Button>
      </div>
    </div>
  );
};

export default ServerConfigDialog;
```

- [ ] **Step 3: Verify it compiles**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm run build
```

Expected: `webpack 5.x.x compiled successfully`.

- [ ] **Step 4: Commit**

```bash
cd /d/SQLLite_Test/XLEdgeOfficeJS
git add -A
git commit -m "Add server instance data module and editable grid (in-memory)"
```

---

### Task 4: Save, validation, and status messages

**Files:**
- Modify: `D:\SQLLite_Test\XLEdgeOfficeJS\src\serverconfig\components\ServerConfigDialog.tsx` (full rewrite)

**Interfaces:**
- Consumes: `saveServerInstances`, `validateServerInstances`, `ensureSingleDefault` from `serverConfigStore.ts` (Task 3).
- Produces: row-selection state (`selectedId`) and the `persist()` helper pattern that Task 5's Set Default/Delete/GO handlers build on — Task 5 replaces this file too and keeps `persist(next, successMessage): boolean` as the save+validate+status entry point.

- [ ] **Step 1: Replace ServerConfigDialog.tsx with Save, validation, and a status MessageBar**

Replace the full contents of `src/serverconfig/components/ServerConfigDialog.tsx` with:

```tsx
import * as React from "react";
import { useState } from "react";
import {
  Button,
  Title2,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  Input,
  Checkbox,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  ServerInstance,
  createEmptyInstance,
  loadServerInstances,
  saveServerInstances,
  validateServerInstances,
  ensureSingleDefault,
} from "../serverConfigStore";

/* global Office */

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    fontFamily: tokens.fontFamilyBase,
  },
  header: {
    padding: "12px 16px",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  toolbar: {
    padding: "8px 16px",
    display: "flex",
    gap: "8px",
  },
  body: {
    flex: 1,
    padding: "12px 16px",
    overflow: "auto",
  },
  status: {
    padding: "0 16px",
  },
  footer: {
    padding: "12px 16px",
    display: "flex",
    justifyContent: "flex-end",
    gap: "8px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  selectedRow: {
    backgroundColor: tokens.colorBrandBackground2,
  },
});

type StatusIntent = "success" | "error" | "warning" | "info";

interface StatusMessage {
  text: string;
  intent: StatusIntent;
}

function closeDialog(): void {
  Office.context.ui.messageParent(JSON.stringify({ type: "close" }));
}

const ServerConfigDialog: React.FC = () => {
  const styles = useStyles();
  const [instances, setInstances] = useState<ServerInstance[]>(() => ensureSingleDefault(loadServerInstances()));
  const [selectedId, setSelectedId] = useState<string | undefined>(undefined);
  const [status, setStatus] = useState<StatusMessage>({
    text: `Configuration loaded. ${loadServerInstances().length} instance(s) found.`,
    intent: "success",
  });

  const updateInstance = (id: string, changes: Partial<ServerInstance>) => {
    setInstances((current) =>
      current.map((instance) => (instance.id === id ? { ...instance, ...changes } : instance))
    );
  };

  const addRow = () => {
    const newInstance = createEmptyInstance();
    setInstances((current) => [...current, newInstance]);
    setSelectedId(newInstance.id);
  };

  const persist = (next: ServerInstance[], successMessage: string): boolean => {
    const validation = validateServerInstances(next);
    if (!validation.isValid) {
      setStatus({ text: validation.message ?? "Validation failed.", intent: "error" });
      return false;
    }
    const normalized = ensureSingleDefault(next);
    saveServerInstances(normalized);
    setInstances(normalized);
    setStatus({ text: successMessage, intent: "success" });
    return true;
  };

  const handleSave = () => {
    persist(instances, "Configuration saved successfully.");
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Title2>Server Configuration</Title2>
      </div>
      <div className={styles.toolbar}>
        <Button appearance="secondary" onClick={addRow}>
          + Add Row
        </Button>
      </div>
      <div className={styles.body}>
        <Table aria-label="Server instances">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Instance Name</TableHeaderCell>
              <TableHeaderCell>URL Address</TableHeaderCell>
              <TableHeaderCell>Default</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {instances.map((instance) => (
              <TableRow
                key={instance.id}
                className={instance.id === selectedId ? styles.selectedRow : undefined}
                onClick={() => setSelectedId(instance.id)}
              >
                <TableCell>
                  <Input
                    value={instance.name}
                    onChange={(_, data) => updateInstance(instance.id, { name: data.value })}
                  />
                </TableCell>
                <TableCell>
                  <Input
                    value={instance.address}
                    onChange={(_, data) => updateInstance(instance.id, { address: data.value })}
                  />
                </TableCell>
                <TableCell>
                  <Checkbox checked={instance.isDefault} disabled />
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>
      <div className={styles.status}>
        <MessageBar intent={status.intent}>
          <MessageBarBody>{status.text}</MessageBarBody>
        </MessageBar>
      </div>
      <div className={styles.footer}>
        <Button appearance="secondary" onClick={handleSave}>
          Save
        </Button>
        <Button appearance="secondary" onClick={closeDialog}>
          Close
        </Button>
      </div>
    </div>
  );
};

export default ServerConfigDialog;
```

Note: the Default checkbox becomes read-only (`disabled`) here, matching the current WPF window's behavior (`DgInstances_BeginningEdit` cancels direct edits on that column and tells the user to use "Set as Default" instead) — Task 5 adds that button.

- [ ] **Step 2: Verify it compiles**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm run build
```

Expected: `webpack 5.x.x compiled successfully`.

- [ ] **Step 3: Commit**

```bash
cd /d/SQLLite_Test/XLEdgeOfficeJS
git add -A
git commit -m "Add Save, validation, and status messages to Server Configuration dialog"
```

---

### Task 5: Set as Default, Delete, and GO

**Files:**
- Modify: `D:\SQLLite_Test\XLEdgeOfficeJS\src\serverconfig\components\ServerConfigDialog.tsx` (full rewrite)

**Interfaces:**
- Consumes: the `persist(next, successMessage): boolean` pattern and `selectedId` state introduced in Task 4.
- Produces: the complete, spec-final `ServerConfigDialog` — no further tasks modify this file.

- [ ] **Step 1: Replace ServerConfigDialog.tsx with the full button set**

Replace the full contents of `src/serverconfig/components/ServerConfigDialog.tsx` with:

```tsx
import * as React from "react";
import { useState } from "react";
import {
  Button,
  Title2,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  Input,
  Checkbox,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  ServerInstance,
  createEmptyInstance,
  loadServerInstances,
  saveServerInstances,
  validateServerInstances,
  ensureSingleDefault,
} from "../serverConfigStore";

/* global Office */

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    fontFamily: tokens.fontFamilyBase,
  },
  header: {
    padding: "12px 16px",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  toolbar: {
    padding: "8px 16px",
    display: "flex",
    gap: "8px",
  },
  body: {
    flex: 1,
    padding: "12px 16px",
    overflow: "auto",
  },
  status: {
    padding: "0 16px",
  },
  footer: {
    padding: "12px 16px",
    display: "flex",
    justifyContent: "flex-end",
    gap: "8px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  selectedRow: {
    backgroundColor: tokens.colorBrandBackground2,
  },
});

type StatusIntent = "success" | "error" | "warning" | "info";

interface StatusMessage {
  text: string;
  intent: StatusIntent;
}

function closeDialog(): void {
  Office.context.ui.messageParent(JSON.stringify({ type: "close" }));
}

const ServerConfigDialog: React.FC = () => {
  const styles = useStyles();
  const [instances, setInstances] = useState<ServerInstance[]>(() => ensureSingleDefault(loadServerInstances()));
  const [selectedId, setSelectedId] = useState<string | undefined>(undefined);
  const [status, setStatus] = useState<StatusMessage>({
    text: `Configuration loaded. ${loadServerInstances().length} instance(s) found.`,
    intent: "success",
  });

  const selectedInstance = instances.find((instance) => instance.id === selectedId);

  const updateInstance = (id: string, changes: Partial<ServerInstance>) => {
    setInstances((current) =>
      current.map((instance) => (instance.id === id ? { ...instance, ...changes } : instance))
    );
  };

  const addRow = () => {
    const newInstance = createEmptyInstance();
    setInstances((current) => [...current, newInstance]);
    setSelectedId(newInstance.id);
  };

  const persist = (next: ServerInstance[], successMessage: string): boolean => {
    const validation = validateServerInstances(next);
    if (!validation.isValid) {
      setStatus({ text: validation.message ?? "Validation failed.", intent: "error" });
      return false;
    }
    const normalized = ensureSingleDefault(next);
    saveServerInstances(normalized);
    setInstances(normalized);
    setStatus({ text: successMessage, intent: "success" });
    return true;
  };

  const handleSave = () => {
    persist(instances, "Configuration saved successfully.");
  };

  const handleSetDefault = () => {
    if (!selectedInstance) {
      setStatus({ text: "Please select an instance to set as default.", intent: "error" });
      return;
    }
    if (selectedInstance.isDefault) {
      setStatus({ text: `'${selectedInstance.name}' is already the default server.`, intent: "warning" });
      return;
    }
    const next = instances.map((instance) => ({ ...instance, isDefault: instance.id === selectedInstance.id }));
    persist(next, `Default set to '${selectedInstance.name}' and saved.`);
  };

  const handleDelete = () => {
    if (!selectedInstance) {
      setStatus({ text: "Please select an instance to delete.", intent: "error" });
      return;
    }
    const wasDefault = selectedInstance.isDefault;
    let remaining = instances.filter((instance) => instance.id !== selectedInstance.id);
    if (wasDefault && remaining.length > 0) {
      remaining = remaining.map((instance, index) => ({ ...instance, isDefault: index === 0 }));
    }
    saveServerInstances(remaining);
    setInstances(remaining);
    setSelectedId(undefined);
    setStatus({ text: `Instance '${selectedInstance.name}' deleted and saved.`, intent: "info" });
  };

  // GO is a deliberate stub, not a placeholder: the spec's "Out of scope" section
  // explicitly excludes wiring this dialog to the real XLEdge add-in/host in this
  // sub-project, so there is no real navigation target for it to call yet.
  const handleGo = () => {
    const target = selectedInstance ?? instances.find((instance) => instance.isDefault);
    if (!target) {
      setStatus({ text: "Please select an instance or set a default server before continuing.", intent: "warning" });
      return;
    }
    if (!target.address.trim()) {
      setStatus({ text: "The selected instance has no URL address configured.", intent: "warning" });
      return;
    }
    setStatus({ text: `Navigating to ${target.name}... (stub: no real host wired up yet)`, intent: "success" });
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Title2>Server Configuration</Title2>
      </div>
      <div className={styles.toolbar}>
        <Button appearance="secondary" onClick={addRow}>
          + Add Row
        </Button>
      </div>
      <div className={styles.body}>
        <Table aria-label="Server instances">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Instance Name</TableHeaderCell>
              <TableHeaderCell>URL Address</TableHeaderCell>
              <TableHeaderCell>Default</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {instances.map((instance) => (
              <TableRow
                key={instance.id}
                className={instance.id === selectedId ? styles.selectedRow : undefined}
                onClick={() => setSelectedId(instance.id)}
              >
                <TableCell>
                  <Input
                    value={instance.name}
                    onChange={(_, data) => updateInstance(instance.id, { name: data.value })}
                  />
                </TableCell>
                <TableCell>
                  <Input
                    value={instance.address}
                    onChange={(_, data) => updateInstance(instance.id, { address: data.value })}
                  />
                </TableCell>
                <TableCell>
                  <Checkbox checked={instance.isDefault} disabled />
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>
      <div className={styles.status}>
        <MessageBar intent={status.intent}>
          <MessageBarBody>{status.text}</MessageBarBody>
        </MessageBar>
      </div>
      <div className={styles.footer}>
        <Button appearance="secondary" onClick={handleGo}>
          GO &gt;&gt;
        </Button>
        <Button appearance="secondary" onClick={handleSetDefault}>
          Set as Default
        </Button>
        <Button appearance="secondary" onClick={handleSave}>
          Save
        </Button>
        <Button appearance="secondary" onClick={handleDelete}>
          Delete
        </Button>
        <Button appearance="secondary" onClick={closeDialog}>
          Close
        </Button>
      </div>
    </div>
  );
};

export default ServerConfigDialog;
```

- [ ] **Step 2: Verify it compiles**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm run build
```

Expected: `webpack 5.x.x compiled successfully`.

- [ ] **Step 3: Commit**

```bash
cd /d/SQLLite_Test/XLEdgeOfficeJS
git add -A
git commit -m "Add Set as Default, Delete, and GO to Server Configuration dialog"
```

---

### Task 6: Sideload into Excel, manually verify, and record findings

**Files:**
- Create: `D:\SQLLite_Test\XLEdgeOfficeJS\FINDINGS.md`

**Interfaces:**
- Consumes: the finished dialog from Tasks 2–5.
- Produces: the actual answer to this sub-project's feasibility question, recorded for whoever decides whether to proceed with the full VSTO → Office.js rewrite.

- [ ] **Step 1: Start the dev server and sideload into Excel desktop**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm start
```

This launches Excel desktop and sideloads the add-in. On first run, Windows may show a certificate-trust prompt for the localhost dev cert (`office-addin-dev-certs`) — accept it; this is a one-time manual step outside the plan's control.

- [ ] **Step 2: Manually exercise every flow from the spec's "Testing / verification" section**

In the sideloaded task pane, click **Open Server Configuration** and confirm/record for each:
1. The dialog opens showing the native OS title bar plus the Fluent UI header inside it (the exact divergence from the WPF window's borderless chrome the spec called out).
2. **+ Add Row** adds an empty row; typing into Name/Address updates the grid live.
3. Leaving Name or Address empty and clicking **Save** shows the "mandatory" error (red/error `MessageBar`).
4. Entering an Address without `http://`/`https://` and clicking **Save** shows the URL-format error.
5. Two rows with the same Name (case-insensitive) and clicking **Save** shows the duplicate-name error.
6. A valid row, **Save** shows the success message, and closing/reopening the dialog (or `npm start` again) shows the row persisted (`localStorage` survived).
7. Selecting a non-default row and clicking **Set as Default** flips the checkbox state (recall it's read-only/disabled by design — only this button changes it) and only one row is ever default.
8. **Delete** on the currently-default row promotes another row to default and saves.
9. **GO** with no selection and no default shows the "select an instance" warning; with a valid selection shows the stub navigating message.
10. **Close** closes the dialog window (round-trips through `messageParent`/`dialog.close()`).

- [ ] **Step 2: Stop the debug session**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm stop
```

- [ ] **Step 3: Record findings**

Create `FINDINGS.md` at the repo root with:
- Whether all 10 checks in Step 2 passed as expected (list any that didn't, with what actually happened).
- A direct comparison of the dialog's look/feel against the WPF `XLEdgeServerConfiguration` window (specifically: how disruptive the native title bar is, whether the Fluent UI styling reads as "close enough to native Office" or not).
- Any Office.js/Fluent UI API gaps hit that aren't in the spec's "Known, expected divergence" note.
- A recommendation: proceed with the full VSTO → Office.js rewrite, proceed with reservations (name them), or don't proceed (why).

- [ ] **Step 4: Commit**

```bash
cd /d/SQLLite_Test/XLEdgeOfficeJS
git add -A
git commit -m "Record Task 6 manual verification findings"
```
