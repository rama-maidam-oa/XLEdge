# XLEdgeOfficeJS About Window Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Port XLEdge's `XLEdgeAbout` WPF window into the `XLEdgeOfficeJS` repo as a second dialog, reusing the Server Configuration screen's existing patterns and data, including a real (not stubbed) HTTP compatibility check against each configured server instance.

**Architecture:** A third page (`src/about/`) mirroring `src/serverconfig/`'s structure, opened from the task pane via the same `Office.context.ui.displayDialogAsync` mechanism (generalized in `App.tsx` to support two independently-tracked dialogs). Reads server instances via the existing `serverConfigStore.ts` (no new persistence). The instance-compatibility grid makes a real `fetch()` per instance; any CORS/mixed-content/network failure resolves to a red X with the error in a tooltip rather than crashing — that failure mode is itself feasibility data for `FINDINGS.md`.

**Tech Stack:** Same as the Server Configuration sub-project — TypeScript, React 18, Fluent UI React v9 (`@fluentui/react-components`, `@fluentui/react-icons`), webpack 5, Office.js.

**Spec:** `docs/superpowers/specs/2026-09-02-xledgeofficejs-about-window-design.md`

## Global Constraints

- Repo root for all tasks: `D:\SQLLite_Test\XLEdgeOfficeJS`.
- Every command invoking `node`/`npm`/`npx` must prepend `export PATH="/c/nvm4w/nodejs:$PATH"` (bash) or `$env:Path = "C:\nvm4w\nodejs;$env:Path"` (PowerShell) — this session's shells predate the Node 24 LTS install (see the Server Configuration plan's Global Constraints for the full story).
- UI library is Fluent UI React v9 (`@fluentui/react-components`, `@fluentui/react-icons`) only. Every component/prop used in this plan's code (`Table`/`TableCell` family, `ProgressBar`, `Tooltip`, `Link`, `Image`, `Body1`/`Caption1`/`Title2`, `CheckmarkCircle24Filled`/`DismissCircle24Filled` and their `primaryFill` prop) was independently verified against the actually-installed package before this plan was written — not assumed from memory or docs.
- No automated test suite (spec-approved, same as Server Configuration). Each task's gate is `npm run build` succeeding. Full interactive verification happens once, in Task 4.
- Data source is exclusively the existing `loadServerInstances()` from `src/serverconfig/serverConfigStore.ts` — this task never writes new data, only reads.
- The compatibility check is a **real** `fetch()`, not a stub — this is a deliberate, spec-approved decision. A failed fetch (CORS, mixed content, network, timeout) must resolve to a visible incompatible row with a reason, never an unhandled promise rejection or a crashed dialog.
- `__APP_VERSION__` is a build-time global constant (via `webpack.DefinePlugin`, reading `package.json`'s `version` field — currently `0.0.1`) declared in `src/globals.d.ts`. Do not attempt to `import` `package.json` directly — `resolveJsonModule` is not enabled in `tsconfig.json` and enabling it is out of scope for this plan.

---

## Investigation findings (already completed — recorded here so no task re-derives them)

- `package.json`'s `version` field is `0.0.1` (generator default, never bumped) and `@types/node` is not installed, so `process.env.APP_VERSION` isn't a clean option — `webpack.DefinePlugin` with an ambient `.d.ts` declaration is the verified approach (matches the existing `webpack.ProvidePlugin` pattern already in `webpack.config.js`).
- Verified against the installed `@fluentui/react-components`/`@fluentui/react-icons` packages directly (not docs): `Table`, `TableHeader`, `TableHeaderCell`, `TableBody`, `TableRow`, `TableCell`, `ProgressBar`, `Tooltip`, `Link`, `Body1`, `Caption1`, `Image` all exist as real exports; `CheckmarkCircle24Filled`/`DismissCircle24Filled` exist and their prop type (`FluentIconsProps`, in `@fluentui/react-icons/lib/utils/FluentIconsProps.types.d.ts`) confirms `primaryFill?: string` is real.
- `TableCell`'s exact prop surface could not be located as a standalone `.d.ts` in this install (only compiled `.js` present at that path) — rather than guess whether `colSpan` is supported, this plan avoids it entirely: the "no instances configured" empty state renders as plain text outside the `Table`, not as a spanning row inside it.
- The existing `Header.tsx` (from the Server Configuration sub-project) already uses `<Image width="90" height="90" .../>` with **string** width/height props, not numbers — this plan's code matches that already-verified-working usage.
- `src/taskpane/components/App.tsx`'s current state (after the Server Configuration sub-project, unmodified since) is reproduced in full in Task 1 below as the "before" — Task 1 replaces it entirely with a generalized two-dialog version.

---

### Task 1: Version constant, About page scaffold, and two-dialog App.tsx

**Files:**
- Create: `D:\SQLLite_Test\XLEdgeOfficeJS\src\globals.d.ts`
- Create: `D:\SQLLite_Test\XLEdgeOfficeJS\src\about\about.html`
- Create: `D:\SQLLite_Test\XLEdgeOfficeJS\src\about\index.tsx`
- Create: `D:\SQLLite_Test\XLEdgeOfficeJS\src\about\components\AboutDialog.tsx`
- Modify: `D:\SQLLite_Test\XLEdgeOfficeJS\webpack.config.js`
- Modify: `D:\SQLLite_Test\XLEdgeOfficeJS\src\taskpane\components\App.tsx` (full rewrite)

**Interfaces:**
- Consumes: nothing new — `App.tsx`'s existing `displayDialogAsync`/`{type:"close"}` pattern from the Server Configuration sub-project is generalized, not replaced.
- Produces: `__APP_VERSION__` global constant (usable anywhere in `src/` after this task). `AboutDialog` (default export, no props) — Task 2 fully rewrites this file's internals but keeps the same default-export shape and the `closeDialog()` contract this task establishes. `openDialog(dialogName, pageName, dialogRef)` helper in `App.tsx` — no later task in this plan modifies `App.tsx` again, so this is the final state of that file.

This task proves the second dialog opens/closes correctly (mirroring how the Server Configuration sub-project front-loaded its own open/close proof) before building real content into it.

- [ ] **Step 1: Add the `__APP_VERSION__` ambient declaration**

Create `src/globals.d.ts`:

```ts
declare const __APP_VERSION__: string;
```

- [ ] **Step 2: Wire `__APP_VERSION__` into webpack**

In `webpack.config.js`, change the `entry` block from:

```js
      serverconfig: {
        import: ["./src/serverconfig/index.tsx", "./src/serverconfig/serverconfig.html"],
        dependOn: "react",
      },
      commands: "./src/commands/commands.ts",
```

to:

```js
      serverconfig: {
        import: ["./src/serverconfig/index.tsx", "./src/serverconfig/serverconfig.html"],
        dependOn: "react",
      },
      about: {
        import: ["./src/about/index.tsx", "./src/about/about.html"],
        dependOn: "react",
      },
      commands: "./src/commands/commands.ts",
```

Change the `plugins` array's `HtmlWebpackPlugin` block from:

```js
      new HtmlWebpackPlugin({
        filename: "serverconfig.html",
        template: "./src/serverconfig/serverconfig.html",
        chunks: ["polyfill", "serverconfig", "react"],
      }),
      new CopyWebpackPlugin({
```

to:

```js
      new HtmlWebpackPlugin({
        filename: "serverconfig.html",
        template: "./src/serverconfig/serverconfig.html",
        chunks: ["polyfill", "serverconfig", "react"],
      }),
      new HtmlWebpackPlugin({
        filename: "about.html",
        template: "./src/about/about.html",
        chunks: ["polyfill", "about", "react"],
      }),
      new CopyWebpackPlugin({
```

And change the end of the `plugins` array from:

```js
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
```

to:

```js
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
      new webpack.DefinePlugin({
        __APP_VERSION__: JSON.stringify(require("./package.json").version),
      }),
    ],
```

- [ ] **Step 3: Create the About page shell**

Create `src/about/about.html`:

```html
<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license. -->
<!-- See LICENSE in the project root for license information -->

<!doctype html>
<html lang="en" data-framework="typescript">

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>About Orbit XLEdge</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
</head>

<body style="width: 100%; height: 100%; margin: 0; padding: 0;">
    <div id="container"></div>
</body>

</html>
```

- [ ] **Step 4: Create the About page's React entry point**

Create `src/about/index.tsx`:

```tsx
import * as React from "react";
import { createRoot } from "react-dom/client";
import AboutDialog from "./components/AboutDialog";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <AboutDialog />
    </FluentProvider>
  );
});
```

- [ ] **Step 5: Create a minimal AboutDialog (header + version + Close)**

Create `src/about/components/AboutDialog.tsx`:

```tsx
import * as React from "react";
import { Button, Title2, Body1, makeStyles, tokens } from "@fluentui/react-components";

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
  body: {
    flex: 1,
    padding: "12px 16px",
    overflow: "auto",
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

const AboutDialog: React.FC = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Title2>About Orbit XLEdge</Title2>
      </div>
      <div className={styles.body}>
        <Body1>Version: {__APP_VERSION__}</Body1>
      </div>
      <div className={styles.footer}>
        <Button appearance="secondary" onClick={closeDialog}>
          Close
        </Button>
      </div>
    </div>
  );
};

export default AboutDialog;
```

- [ ] **Step 6: Generalize App.tsx to open two independent dialogs**

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
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    alignItems: "flex-start",
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

function handleDialogEvent(dialogName: string, arg: { error: number }): void {
  console.log(`${dialogName} dialog closed with event ${arg.error}`);
}

function openDialog(
  dialogName: string,
  pageName: string,
  dialogRef: React.MutableRefObject<Office.Dialog | undefined>
): void {
  const dialogUrl = `${window.location.origin}/${pageName}`;
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 60, width: 40, displayInIframe: false },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(`${dialogName} dialog failed to open: ${asyncResult.error.message}`);
        return;
      }
      dialogRef.current = asyncResult.value;
      dialogRef.current.addEventHandler(Office.EventType.DialogMessageReceived, (arg) =>
        handleDialogMessage(dialogRef, arg as Office.DialogParentMessageReceivedEventArgs)
      );
      dialogRef.current.addEventHandler(Office.EventType.DialogEventReceived, (arg) =>
        handleDialogEvent(dialogName, arg as { error: number })
      );
    }
  );
}

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  const serverConfigDialogRef = useRef<Office.Dialog | undefined>(undefined);
  const aboutDialogRef = useRef<Office.Dialog | undefined>(undefined);

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="XLEdge Office.js POC" />
      <div className={styles.content}>
        <Button
          appearance="primary"
          onClick={() => openDialog("Server Configuration", "serverconfig.html", serverConfigDialogRef)}
        >
          Open Server Configuration
        </Button>
        <Button appearance="primary" onClick={() => openDialog("About", "about.html", aboutDialogRef)}>
          Open About
        </Button>
      </div>
    </div>
  );
};

export default App;
```

- [ ] **Step 7: Verify it compiles**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm run build
```

Expected: `webpack ... compiled successfully`, with `about.html`/`about.js` now listed among the emitted assets alongside `taskpane.html`/`serverconfig.html`.

- [ ] **Step 8: Commit**

```bash
cd /d/SQLLite_Test/XLEdgeOfficeJS
git add -A
git commit -m "Add About dialog plumbing and __APP_VERSION__ build constant"
```

---

### Task 2: Product info panel + Instance Compatibility grid shell

**Files:**
- Modify: `D:\SQLLite_Test\XLEdgeOfficeJS\src\about\components\AboutDialog.tsx` (full rewrite)

**Interfaces:**
- Consumes: `ServerInstance`, `loadServerInstances` from `../../serverconfig/serverConfigStore` (already exists, unmodified).
- Produces: the grid's row/column structure Task 3 builds real status onto — Task 3 replaces this file again but keeps the same `Table`/`TableRow`/`TableCell` layout and the `styles.productInfo`/`styles.productDetails` class names.

- [ ] **Step 1: Replace AboutDialog.tsx with the product info panel and a static instance list**

Replace the full contents of `src/about/components/AboutDialog.tsx` with:

```tsx
import * as React from "react";
import { useState } from "react";
import {
  Button,
  Title2,
  Body1,
  Caption1,
  Link,
  Image,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { ServerInstance, loadServerInstances } from "../../serverconfig/serverConfigStore";

/* global Office, window */

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
  body: {
    flex: 1,
    padding: "12px 16px",
    overflow: "auto",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  productInfo: {
    display: "flex",
    gap: "16px",
    alignItems: "flex-start",
  },
  productDetails: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  footer: {
    padding: "12px 16px",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
});

function closeDialog(): void {
  Office.context.ui.messageParent(JSON.stringify({ type: "close" }));
}

function openSupportLink(): void {
  window.open("https://www.orbitanalytics.com", "_blank", "noopener,noreferrer");
}

const AboutDialog: React.FC = () => {
  const styles = useStyles();
  const [instances] = useState<ServerInstance[]>(() => loadServerInstances());

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Title2>About Orbit XLEdge</Title2>
      </div>
      <div className={styles.body}>
        <div className={styles.productInfo}>
          <Image width="90" height="90" src="assets/logo-filled.png" alt="Orbit XLEdge logo" />
          <div className={styles.productDetails}>
            <Body1>Version: {__APP_VERSION__}</Body1>
            <Caption1>Financial Analytics and Reporting Tool</Caption1>
            <Caption1>Build Date: POC build</Caption1>
            <Caption1>
              For Support please contact: <Link onClick={openSupportLink}>https://www.orbitanalytics.com</Link>
            </Caption1>
          </div>
        </div>
        {instances.length === 0 ? (
          <Caption1>No instances configured</Caption1>
        ) : (
          <Table aria-label="Instance compatibility">
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Instance</TableHeaderCell>
                <TableHeaderCell>Compatible</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {instances.map((instance) => (
                <TableRow key={instance.id}>
                  <TableCell>{instance.address}</TableCell>
                  <TableCell>Pending</TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        )}
      </div>
      <div className={styles.footer}>
        <Caption1>Copyright © 2015 Orbit Analytics. All Rights Reserved.</Caption1>
        <Button appearance="secondary" onClick={closeDialog}>
          Close
        </Button>
      </div>
    </div>
  );
};

export default AboutDialog;
```

Note: the "no instances configured" case renders as plain text outside the `Table`, not as a spanning row inside it — `TableCell`'s `colSpan` support couldn't be confirmed against the installed package's type declarations (see this plan's Investigation findings), so this avoids the uncertainty entirely rather than guessing.

- [ ] **Step 2: Verify it compiles**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm run build
```

Expected: `webpack ... compiled successfully`.

- [ ] **Step 3: Commit**

```bash
cd /d/SQLLite_Test/XLEdgeOfficeJS
git add -A
git commit -m "Add About product info panel and instance list shell"
```

---

### Task 3: Real compatibility check logic

**Files:**
- Modify: `D:\SQLLite_Test\XLEdgeOfficeJS\src\about\components\AboutDialog.tsx` (full rewrite)

**Interfaces:**
- Consumes: `checkCompatibility(address: string): Promise<{ compatible: boolean; reason?: string }>` — defined and used entirely within this file, not exported (no later task needs it).
- Produces: the final, spec-complete `AboutDialog` — no further tasks in this plan modify this file.

- [ ] **Step 1: Replace AboutDialog.tsx with the real sequential compatibility check**

Replace the full contents of `src/about/components/AboutDialog.tsx` with:

```tsx
import * as React from "react";
import { useEffect, useState } from "react";
import {
  Button,
  Title2,
  Body1,
  Caption1,
  Link,
  Image,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  ProgressBar,
  Tooltip,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { CheckmarkCircle24Filled, DismissCircle24Filled } from "@fluentui/react-icons";
import { ServerInstance, loadServerInstances } from "../../serverconfig/serverConfigStore";

/* global Office, window, fetch */

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
  body: {
    flex: 1,
    padding: "12px 16px",
    overflow: "auto",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  productInfo: {
    display: "flex",
    gap: "16px",
    alignItems: "flex-start",
  },
  productDetails: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  progress: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
  },
  compatIcon: {
    display: "flex",
    justifyContent: "center",
  },
  footer: {
    padding: "12px 16px",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
});

type CompatibilityStatus = "pending" | "compatible" | "incompatible";

interface CompatibilityResult {
  instance: ServerInstance;
  status: CompatibilityStatus;
  reason?: string;
}

function closeDialog(): void {
  Office.context.ui.messageParent(JSON.stringify({ type: "close" }));
}

function openSupportLink(): void {
  window.open("https://www.orbitanalytics.com", "_blank", "noopener,noreferrer");
}

async function checkCompatibility(address: string): Promise<{ compatible: boolean; reason?: string }> {
  try {
    const response = await fetch(`${address}/rest/public/orbit-version`);
    if (!response.ok) {
      return { compatible: false, reason: `HTTP ${response.status}` };
    }

    const body = await response.json();
    const versionInfoKey = Object.keys(body).find((key) => key.toLowerCase() === "verioninfo");
    const versionInfo = versionInfoKey ? body[versionInfoKey] : undefined;
    const versionKey = versionInfo
      ? Object.keys(versionInfo).find((key) => key.toLowerCase() === "xledgeversion")
      : undefined;
    const remoteVersion = versionKey ? versionInfo[versionKey] : undefined;

    if (typeof remoteVersion !== "string") {
      return { compatible: false, reason: "No version found in response" };
    }
    if (remoteVersion !== __APP_VERSION__) {
      return { compatible: false, reason: `Version mismatch: expected ${__APP_VERSION__}, got ${remoteVersion}` };
    }
    return { compatible: true };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return { compatible: false, reason: message };
  }
}

const AboutDialog: React.FC = () => {
  const styles = useStyles();
  const [instances] = useState<ServerInstance[]>(() => loadServerInstances());
  const [results, setResults] = useState<CompatibilityResult[]>(() =>
    instances.map((instance) => ({ instance, status: "pending" as CompatibilityStatus }))
  );
  const [progressText, setProgressText] = useState<string>(
    instances.length > 0 ? "Starting compatibility checks..." : "No instances configured"
  );
  const [checking, setChecking] = useState<boolean>(instances.length > 0);

  useEffect(() => {
    if (instances.length === 0) {
      return;
    }

    let cancelled = false;

    (async () => {
      for (let i = 0; i < instances.length; i++) {
        const instance = instances[i];
        if (cancelled) {
          return;
        }
        setProgressText(`Checking ${instance.name}... (${i + 1}/${instances.length})`);

        const { compatible, reason } = await checkCompatibility(instance.address);

        if (cancelled) {
          return;
        }
        setResults((current) =>
          current.map((result) =>
            result.instance.id === instance.id
              ? { instance, status: compatible ? "compatible" : "incompatible", reason }
              : result
          )
        );
      }

      if (!cancelled) {
        setProgressText("Compatibility check completed");
        setChecking(false);
      }
    })();

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Title2>About Orbit XLEdge</Title2>
      </div>
      <div className={styles.body}>
        <div className={styles.productInfo}>
          <Image width="90" height="90" src="assets/logo-filled.png" alt="Orbit XLEdge logo" />
          <div className={styles.productDetails}>
            <Body1>Version: {__APP_VERSION__}</Body1>
            <Caption1>Financial Analytics and Reporting Tool</Caption1>
            <Caption1>Build Date: POC build</Caption1>
            <Caption1>
              For Support please contact: <Link onClick={openSupportLink}>https://www.orbitanalytics.com</Link>
            </Caption1>
          </div>
        </div>
        {results.length === 0 ? (
          <Caption1>No instances configured</Caption1>
        ) : (
          <Table aria-label="Instance compatibility">
            <TableHeader>
              <TableRow>
                <TableHeaderCell>Instance</TableHeaderCell>
                <TableHeaderCell>Compatible</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {results.map((result) => (
                <TableRow key={result.instance.id}>
                  <TableCell>{result.instance.address}</TableCell>
                  <TableCell>
                    <div className={styles.compatIcon}>
                      {result.status === "pending" && <Caption1>...</Caption1>}
                      {result.status === "compatible" && (
                        <Tooltip content="Compatible" relationship="label">
                          <CheckmarkCircle24Filled primaryFill={tokens.colorPaletteGreenForeground1} />
                        </Tooltip>
                      )}
                      {result.status === "incompatible" && (
                        <Tooltip content={result.reason ?? "Incompatible"} relationship="label">
                          <DismissCircle24Filled primaryFill={tokens.colorPaletteRedForeground1} />
                        </Tooltip>
                      )}
                    </div>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        )}
        {checking && (
          <div className={styles.progress}>
            <ProgressBar />
            <Caption1>{progressText}</Caption1>
          </div>
        )}
      </div>
      <div className={styles.footer}>
        <Caption1>Copyright © 2015 Orbit Analytics. All Rights Reserved.</Caption1>
        <Button appearance="secondary" onClick={closeDialog}>
          Close
        </Button>
      </div>
    </div>
  );
};

export default AboutDialog;
```

- [ ] **Step 2: Verify it compiles**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm run build
```

Expected: `webpack ... compiled successfully`.

- [ ] **Step 3: Commit**

```bash
cd /d/SQLLite_Test/XLEdgeOfficeJS
git add -A
git commit -m "Add real compatibility check (fetch, progress, per-row status) to About dialog"
```

---

### Task 4: Manual verification in Excel and FINDINGS.md update

**Files:**
- Modify: `D:\SQLLite_Test\XLEdgeOfficeJS\FINDINGS.md`

**Interfaces:**
- Consumes: the finished About dialog from Tasks 1–3.
- Produces: the answer to this screen's specific feasibility question (can the add-in make real server calls at all) recorded for the rewrite decision.

- [ ] **Step 1: Sideload and manually verify**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm start
```

In the task pane, click **Open About** and confirm/record:
1. The dialog opens showing the logo, version (`0.0.1`), description, build date placeholder, and support link.
2. The compatibility grid lists every instance configured via Server Configuration, each starting in a pending state.
3. The progress bar + "Checking X... (i/N)" text advances through the instances one at a time, then disappears when done.
4. Each row settles to a green check or red X with a tooltip; for a red X, note what the tooltip actually says (CORS error, mixed-content block, HTTP error, timeout, version mismatch, etc.) — this is the real data point this task exists to capture.
5. Clicking the support link opens `https://www.orbitanalytics.com` in the default browser.
6. **Open Server Configuration** still works exactly as before (proves the two-dialog generalization in `App.tsx` didn't regress it).
7. Close closes the About dialog.

- [ ] **Step 2: Stop the debug session**

```bash
export PATH="/c/nvm4w/nodejs:$PATH"
cd /d/SQLLite_Test/XLEdgeOfficeJS
npm stop
```

- [ ] **Step 3: Record findings**

Add a new section to `FINDINGS.md` (after the existing sections, before `## Recommendation` — renumber if needed so headings stay sequential) titled `## About Window / Instance Compatibility Check`, covering:
- Whether all 7 checks in Step 1 passed as expected (list any that didn't).
- For each configured instance, what the actual compatibility-check outcome was and why (the tooltip reason from check 4 above) — this is the concrete answer to "can an Office.js add-in make direct calls to XLEdge servers the way the VSTO app's `HttpClient` does."
- Whether this changes the overall rewrite recommendation, or is a solvable/known issue (e.g., "would need the servers' CORS headers updated" or "would need a lightweight proxy") to flag as a prerequisite rather than a blocker.

- [ ] **Step 4: Commit**

```bash
cd /d/SQLLite_Test/XLEdgeOfficeJS
git add -A
git commit -m "Record About window / compatibility-check manual verification findings"
```
