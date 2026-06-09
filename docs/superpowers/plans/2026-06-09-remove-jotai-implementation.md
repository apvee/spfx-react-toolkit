# Remove Jotai Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove Jotai from runtime code, package metadata, and documentation while preserving public provider/hook behavior.

**Architecture:** Replace Jotai atoms with a provider-scoped internal runtime store. React hooks subscribe through selectors so unrelated state changes do not cause unnecessary re-renders. Storage hooks become local browser-storage hooks and do not use the provider runtime store.

**Tech Stack:** TypeScript, React 17, SPFx 1.21, Node verification scripts, TypeScript compiler, ESLint, Gulp bundle.

---

## File Structure

- Create `scripts/verify-runtime-store.cjs`: Node regression script for the pure runtime store.
- Create `scripts/verify-jotai-removal.cjs`: Node regression script for final package/source cleanup.
- Create `src/core/runtime-store.internal.ts`: pure non-React store implementation.
- Create `src/core/state.internal.tsx`: React context, selector hook, action helpers.
- Modify `src/core/provider-base.internal.tsx`: replace Jotai provider/store with runtime store and imperative property sync.
- Delete `src/core/atoms.internal.ts`: old Jotai atom module.
- Modify `src/hooks/useSPFxThemeInfo.ts`: select theme from runtime store.
- Modify `src/hooks/useSPFxDisplayMode.ts`: select display mode from runtime store.
- Modify `src/hooks/useSPFxProperties.ts`: select and update properties through runtime actions.
- Modify `src/hooks/useSPFxContainerInfo.ts`: select container state and update size through runtime actions.
- Modify `src/hooks/useSPFxTeams.ts`: select and update Teams runtime state.
- Modify `src/hooks/useSPFxStorage.ts`: replace `atomWithStorage` with local browser-storage state.
- Modify `src/utils/theme-subscription.internal.ts`: keep signature compatible with plain setter; update comments only if needed.
- Modify `package.json`: remove Jotai dependency/keyword and add `lib/services/**/*`, `lib/helpers/**/*`, and verification script entries.
- Modify `package-lock.json`: remove Jotai package entry and root dependency.
- Modify `README.md` and `docs/INTRODUCTION.md`: remove "Built on Jotai" language.

## Task 1: Baseline And Failing Regression Scripts

**Files:**
- Create: `scripts/verify-runtime-store.cjs`
- Create: `scripts/verify-jotai-removal.cjs`
- No production code changes.

- [ ] **Step 1: Capture baseline declarations**

Run:

```bash
npx tsc --emitDeclarationOnly --declaration --declarationMap false --outDir /tmp/spfx-react-toolkit-jotai-before --pretty false
```

Expected: exit 0.

- [ ] **Step 2: Write runtime store regression script**

Create `scripts/verify-runtime-store.cjs`:

```js
const assert = require('assert');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const root = path.resolve(__dirname, '..');
const outDir = path.join(os.tmpdir(), 'spfx-runtime-store-verify');

fs.rmSync(outDir, { recursive: true, force: true });
fs.mkdirSync(outDir, { recursive: true });

execFileSync(
  'npx',
  [
    'tsc',
    'src/core/runtime-store.internal.ts',
    '--module',
    'commonjs',
    '--target',
    'es2020',
    '--skipLibCheck',
    '--esModuleInterop',
    '--outDir',
    outDir,
    '--pretty',
    'false'
  ],
  { cwd: root, stdio: 'inherit' }
);

const {
  createSPFxRuntimeStore,
  createDefaultSPFxRuntimeState
} = require(path.join(outDir, 'runtime-store.internal.js'));

const defaultState = createDefaultSPFxRuntimeState();
assert.deepStrictEqual(defaultState.teams, { supported: false, initialized: false });
assert.strictEqual(defaultState.theme, undefined);
assert.strictEqual(defaultState.displayMode, undefined);

const store = createSPFxRuntimeStore(defaultState);
assert.deepStrictEqual(store.getState(), defaultState);

let notifications = 0;
const unsubscribe = store.subscribe(() => {
  notifications += 1;
});

store.setState({ theme: undefined });
assert.strictEqual(notifications, 0, 'unchanged partial state must not notify');

const properties = { title: 'Hello' };
store.setState({ properties });
assert.strictEqual(store.getState().properties, properties);
assert.strictEqual(notifications, 1, 'changed partial state must notify once');

store.setState(previous => previous);
assert.strictEqual(notifications, 1, 'same updater state must not notify');

let snapshotNotifications = 0;
const unsubscribeDuringNotify = store.subscribe(() => {
  snapshotNotifications += 1;
  unsubscribeDuringNotify();
});

store.setState({ displayMode: 1 });
assert.strictEqual(snapshotNotifications, 1, 'listener should run during snapshot notification');
store.setState({ displayMode: 2 });
assert.strictEqual(snapshotNotifications, 1, 'unsubscribed listener should not run again');
assert.strictEqual(notifications, 3, 'remaining listener should receive both display mode updates');

unsubscribe();
store.setState({ containerSize: { width: 10, height: 20 } });
assert.strictEqual(notifications, 3, 'unsubscribed listener should not receive later updates');

console.log('runtime store verification passed');
```

- [ ] **Step 3: Verify runtime store script fails before implementation**

Run:

```bash
node scripts/verify-runtime-store.cjs
```

Expected: FAIL because `src/core/runtime-store.internal.ts` does not exist yet.

- [ ] **Step 4: Write Jotai removal regression script**

Create `scripts/verify-jotai-removal.cjs`:

```js
const assert = require('assert');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const sourceFiles = [];

function walk(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      walk(fullPath);
    } else if (/\.(ts|tsx|js|json|md)$/.test(entry.name)) {
      sourceFiles.push(fullPath);
    }
  }
}

for (const dir of ['src', 'docs']) {
  const absoluteDir = path.join(root, dir);
  if (fs.existsSync(absoluteDir)) {
    walk(absoluteDir);
  }
}

for (const file of [
  'package.json',
  'package-lock.json',
  'README.md'
]) {
  sourceFiles.push(path.join(root, file));
}

const filteredSourceFiles = sourceFiles.filter(file => {
  const relative = path.relative(root, file);
  return !relative.startsWith(`docs${path.sep}superpowers${path.sep}`);
});

const forbidden = [
  'jotai',
  'atomWithStorage',
  'useAtomValue',
  'useSetAtom',
  'useAtom(',
  'spfxAtoms',
  'atoms.internal'
];

const offenders = [];
for (const file of filteredSourceFiles) {
  const text = fs.readFileSync(file, 'utf8');
  for (const token of forbidden) {
    if (text.includes(token)) {
      offenders.push(`${path.relative(root, file)} contains ${token}`);
    }
  }
}

assert.deepStrictEqual(offenders, []);

const packageJson = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
assert.ok(packageJson.files.includes('lib/services/**/*'), 'package files must include lib/services/**/*');
assert.ok(packageJson.files.includes('lib/helpers/**/*'), 'package files must include lib/helpers/**/*');

console.log('jotai removal verification passed');
```

- [ ] **Step 5: Verify Jotai removal script fails before implementation**

Run:

```bash
node scripts/verify-jotai-removal.cjs
```

Expected: FAIL and list current Jotai references.

## Task 2: Pure Runtime Store

**Files:**
- Create: `src/core/runtime-store.internal.ts`
- Test: `scripts/verify-runtime-store.cjs`

- [ ] **Step 1: Implement pure runtime store**

Create `src/core/runtime-store.internal.ts` with:

```ts
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { DisplayMode } from '@microsoft/sp-core-library';
import type { ContainerSize } from './types';

export type SPFxRuntimeTeamsTheme = 'default' | 'dark' | 'highContrast';

export interface SPFxRuntimeTeamsState {
  readonly supported: boolean;
  readonly context?: unknown;
  readonly theme?: SPFxRuntimeTeamsTheme;
  readonly initialized: boolean;
}

export interface SPFxRuntimeState {
  readonly theme: IReadonlyTheme | undefined;
  readonly displayMode: DisplayMode | undefined;
  readonly properties: unknown;
  readonly containerEl: HTMLElement | undefined;
  readonly containerSize: ContainerSize | undefined;
  readonly teams: SPFxRuntimeTeamsState;
}

export type SPFxRuntimeStateUpdater =
  | Partial<SPFxRuntimeState>
  | ((previous: SPFxRuntimeState) => SPFxRuntimeState);

export interface SPFxRuntimeStore {
  getState: () => SPFxRuntimeState;
  setState: (updater: SPFxRuntimeStateUpdater) => void;
  subscribe: (listener: () => void) => () => void;
}

export function createDefaultSPFxRuntimeState(): SPFxRuntimeState {
  return {
    theme: undefined,
    displayMode: undefined,
    properties: undefined,
    containerEl: undefined,
    containerSize: undefined,
    teams: {
      supported: false,
      initialized: false,
    },
  };
}

function mergeState(
  previous: SPFxRuntimeState,
  partial: Partial<SPFxRuntimeState>
): SPFxRuntimeState {
  let changed = false;
  const next: SPFxRuntimeState = {
    ...previous,
    ...partial,
  };

  for (const key in next) {
    const stateKey = key as keyof SPFxRuntimeState;
    if (!Object.is(previous[stateKey], next[stateKey])) {
      changed = true;
      break;
    }
  }

  return changed ? next : previous;
}

export function createSPFxRuntimeStore(
  initialState?: Partial<SPFxRuntimeState>
): SPFxRuntimeStore {
  let state = mergeState(createDefaultSPFxRuntimeState(), initialState ?? {});
  const listeners = new Set<() => void>();

  const getState = (): SPFxRuntimeState => state;

  const setState = (updater: SPFxRuntimeStateUpdater): void => {
    const nextState = typeof updater === 'function'
      ? updater(state)
      : mergeState(state, updater);

    if (Object.is(nextState, state)) {
      return;
    }

    state = nextState;
    const snapshot = Array.from(listeners);

    for (let i = 0; i < snapshot.length; i++) {
      snapshot[i]();
    }
  };

  const subscribe = (listener: () => void): (() => void) => {
    listeners.add(listener);

    return function unsubscribe(): void {
      listeners.delete(listener);
    };
  };

  return {
    getState,
    setState,
    subscribe,
  };
}
```

- [ ] **Step 2: Verify runtime store script passes**

Run:

```bash
node scripts/verify-runtime-store.cjs
```

Expected: PASS and print `runtime store verification passed`.

- [ ] **Step 3: Typecheck**

Run:

```bash
npx tsc --noEmit --pretty false
```

Expected: exit 0.

## Task 3: React Runtime Context And Selectors

**Files:**
- Create: `src/core/state.internal.tsx`
- Depends on: `src/core/runtime-store.internal.ts`

- [ ] **Step 1: Implement runtime React bridge**

Create `src/core/state.internal.tsx` with:

```tsx
import * as React from 'react';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import type { DisplayMode } from '@microsoft/sp-core-library';
import type { ContainerSize } from './types';
import type {
  SPFxRuntimeState,
  SPFxRuntimeStore,
  SPFxRuntimeTeamsState,
} from './runtime-store.internal';

export const SPFxRuntimeStoreContext = React.createContext<SPFxRuntimeStore | undefined>(undefined);

if (process.env.NODE_ENV !== 'production') {
  SPFxRuntimeStoreContext.displayName = 'SPFxRuntimeStoreContext';
}

function objectIs<T>(previous: T, next: T): boolean {
  return Object.is(previous, next);
}

export function useSPFxRuntimeStore(): SPFxRuntimeStore {
  const store = React.useContext(SPFxRuntimeStoreContext);

  if (!store) {
    throw new Error(
      'SPFx runtime state is not available. ' +
      'Make sure your component is wrapped with a host-specific SPFx provider.'
    );
  }

  return store;
}

export function useSPFxRuntimeSelector<T>(
  selector: (state: SPFxRuntimeState) => T,
  isEqual: (previous: T, next: T) => boolean = objectIs
): T {
  const store = useSPFxRuntimeStore();
  const selectorRef = React.useRef(selector);
  const equalityRef = React.useRef(isEqual);
  const [selected, setSelected] = React.useState<T>(() => selector(store.getState()));

  React.useEffect(() => {
    selectorRef.current = selector;
    equalityRef.current = isEqual;

    setSelected(function(previous): T {
      const next = selector(store.getState());
      return isEqual(previous, next) ? previous : next;
    });
  }, [selector, isEqual, store]);

  React.useEffect(() => {
    const checkForUpdates = (): void => {
      setSelected(function(previous): T {
        const next = selectorRef.current(store.getState());
        return equalityRef.current(previous, next) ? previous : next;
      });
    };

    const unsubscribe = store.subscribe(checkForUpdates);
    checkForUpdates();

    return unsubscribe;
  }, [store]);

  return selected;
}

export interface SPFxRuntimeActions {
  readonly setTheme: (theme: IReadonlyTheme | undefined) => void;
  readonly setDisplayMode: (displayMode: DisplayMode | undefined) => void;
  readonly setProperties: (updater: unknown | ((previous: unknown) => unknown)) => void;
  readonly setContainerElement: (element: HTMLElement | undefined) => void;
  readonly setContainerSize: (size: ContainerSize | undefined) => void;
  readonly setTeamsState: (teams: SPFxRuntimeTeamsState) => void;
}

export function createSPFxRuntimeActions(store: SPFxRuntimeStore): SPFxRuntimeActions {
  return {
    setTheme: function(theme: IReadonlyTheme | undefined): void {
      store.setState({ theme });
    },
    setDisplayMode: function(displayMode: DisplayMode | undefined): void {
      store.setState({ displayMode });
    },
    setProperties: function(updater: unknown | ((previous: unknown) => unknown)): void {
      store.setState(function(previous): SPFxRuntimeState {
        const properties = typeof updater === 'function'
          ? (updater as (previous: unknown) => unknown)(previous.properties)
          : updater;

        if (Object.is(properties, previous.properties)) {
          return previous;
        }

        return {
          ...previous,
          properties,
        };
      });
    },
    setContainerElement: function(element: HTMLElement | undefined): void {
      store.setState({ containerEl: element });
    },
    setContainerSize: function(size: ContainerSize | undefined): void {
      store.setState({ containerSize: size });
    },
    setTeamsState: function(teams: SPFxRuntimeTeamsState): void {
      store.setState({ teams });
    },
  };
}

export function useSPFxRuntimeActions(): SPFxRuntimeActions {
  const store = useSPFxRuntimeStore();
  return React.useMemo(() => createSPFxRuntimeActions(store), [store]);
}
```

- [ ] **Step 2: Typecheck**

Run:

```bash
npx tsc --noEmit --pretty false
```

Expected: exit 0.

## Task 4: Provider Migration

**Files:**
- Modify: `src/core/provider-base.internal.tsx`
- Modify: `src/utils/theme-subscription.internal.ts` comments only if needed.

- [ ] **Step 1: Replace Jotai store in provider**

Update `src/core/provider-base.internal.tsx` so it:

- removes imports from `jotai`;
- removes import of `spfxAtoms`;
- imports `SPFxRuntimeStoreContext`, `createSPFxRuntimeActions`;
- imports `createSPFxRuntimeStore`;
- creates one runtime store per provider;
- uses runtime actions for theme/properties/display mode/container element;
- uses an imperative store subscription for properties -> SPFx sync;
- renders `<SPFxRuntimeStoreContext.Provider value={store}>` instead of `<Provider store={store}>`.

- [ ] **Step 2: Preserve property sync behavior**

The runtime -> SPFx sync effect must preserve this behavior:

```ts
if (properties === undefined) {
  return;
}

if (properties !== lastPropertiesRef.current) {
  const target = instanceAny.properties as Record<string, unknown>;
  const source = properties as Record<string, unknown>;

  for (const key in target) {
    if (Object.prototype.hasOwnProperty.call(target, key)) {
      delete target[key];
    }
  }

  for (const key in source) {
    if (Object.prototype.hasOwnProperty.call(source, key)) {
      target[key] = source[key];
    }
  }

  lastPropertiesRef.current = properties;

  if (isWebPart(instance)) {
    const ctx = instanceAny.context as unknown as { propertyPane?: { refresh(): void } };
    if (ctx.propertyPane && typeof ctx.propertyPane.refresh === 'function') {
      ctx.propertyPane.refresh();
    }
  }
}
```

- [ ] **Step 3: Typecheck**

Run:

```bash
npx tsc --noEmit --pretty false
```

Expected: exit 0.

## Task 5: Runtime Hook Migration

**Files:**
- Modify: `src/hooks/useSPFxThemeInfo.ts`
- Modify: `src/hooks/useSPFxDisplayMode.ts`
- Modify: `src/hooks/useSPFxProperties.ts`
- Modify: `src/hooks/useSPFxContainerInfo.ts`
- Modify: `src/hooks/useSPFxTeams.ts`

- [ ] **Step 1: Migrate theme hook**

`useSPFxThemeInfo` should import `useSPFxRuntimeSelector` and return:

```ts
return useSPFxRuntimeSelector(state => state.theme);
```

- [ ] **Step 2: Migrate display mode hook**

`useSPFxDisplayMode` should select `displayMode` from runtime state and preserve default `DisplayMode.Read`.

- [ ] **Step 3: Migrate properties hook**

`useSPFxProperties` should:

- select `properties` from runtime state;
- use `useSPFxRuntimeActions().setProperties`;
- preserve `setProperties(updates)` shallow merge;
- preserve `updateProperties(updater)`.

- [ ] **Step 4: Migrate container hook**

`useSPFxContainerInfo` should:

- select `containerEl`;
- select `containerSize`;
- use `useSPFxRuntimeActions().setContainerSize`;
- keep `useResizeObserver(element, setContainerSize)`.

- [ ] **Step 5: Migrate Teams hook**

`useSPFxTeams` should:

- select `teams`;
- update via `useSPFxRuntimeActions().setTeamsState`;
- preserve v2 then v1 initialization behavior;
- preserve unsupported fallback;
- preserve theme normalization.

- [ ] **Step 6: Typecheck and lint migrated hooks**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/core src/hooks/useSPFxThemeInfo.ts src/hooks/useSPFxDisplayMode.ts src/hooks/useSPFxProperties.ts src/hooks/useSPFxContainerInfo.ts src/hooks/useSPFxTeams.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit 0.

## Task 6: Storage Hook Migration

**Files:**
- Modify: `src/hooks/useSPFxStorage.ts`

- [ ] **Step 1: Replace atom storage with browser storage helper**

`useSPFxStorage.ts` should:

- remove all Jotai imports;
- import `useMemo`, `useCallback`, `useEffect`, `useState`;
- read initial storage lazily;
- catch storage read/write/remove errors;
- support functional setter;
- reset to `defaultValue` on remove;
- re-read when scoped key or default value changes;
- listen for browser `storage` events for the same key.

- [ ] **Step 2: Preserve public return type**

Do not change:

```ts
export interface SPFxStorageHook<T> {
  readonly value: T;
  readonly setValue: (value: T | ((prev: T) => T)) => void;
  readonly remove: () => void;
}
```

- [ ] **Step 3: Typecheck migrated storage**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/hooks/useSPFxStorage.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit 0.

## Task 7: Remove Jotai From Package And Docs

**Files:**
- Delete: `src/core/atoms.internal.ts`
- Modify: `package.json`
- Modify: `package-lock.json`
- Modify: `README.md`
- Modify: `docs/INTRODUCTION.md`

- [ ] **Step 1: Delete old atom module**

Remove `src/core/atoms.internal.ts`.

- [ ] **Step 2: Update package metadata**

In `package.json`:

- remove `"jotai": "^2.6.0"` from `dependencies`;
- remove `"jotai"` from `keywords`;
- add `"lib/services/**/*"` and `"lib/helpers/**/*"` to `files` if missing;
- add scripts:

```json
"verify:runtime-store": "node scripts/verify-runtime-store.cjs",
"verify:jotai-removal": "node scripts/verify-jotai-removal.cjs"
```

- [ ] **Step 3: Update package lock**

Run:

```bash
npm install --package-lock-only --ignore-scripts --cache /tmp/spfx-react-toolkit-npm-cache
```

Expected: exit 0 and `package-lock.json` no longer contains the root Jotai dependency or `node_modules/jotai`.

- [ ] **Step 4: Update docs**

Replace README and `docs/INTRODUCTION.md` Jotai wording with provider-scoped runtime store wording.

- [ ] **Step 5: Run Jotai removal verification**

Run:

```bash
node scripts/verify-jotai-removal.cjs
```

Expected: PASS and print `jotai removal verification passed`.

## Task 8: Final Verification And API Compatibility

**Files:**
- No new implementation files unless fixing verification failures.

- [ ] **Step 1: Build declarations after refactor**

Run:

```bash
npx tsc --emitDeclarationOnly --declaration --declarationMap false --outDir /tmp/spfx-react-toolkit-jotai-after --pretty false
```

Expected: exit 0.

- [ ] **Step 2: Compare public hook declarations**

Run:

```bash
diff -rq /tmp/spfx-react-toolkit-jotai-before/hooks /tmp/spfx-react-toolkit-jotai-after/hooks
```

Expected: no public signature changes except import/comment differences that do not alter hook names, parameters, or return types. Inspect any reported file with `diff -u`.

- [ ] **Step 3: Full static verification**

Run:

```bash
node scripts/verify-runtime-store.cjs
node scripts/verify-jotai-removal.cjs
npx tsc --noEmit --pretty false
npx eslint src --ext .ts,.tsx --max-warnings=0
npm run build
npm --cache /tmp/spfx-react-toolkit-npm-cache pack --dry-run
git diff --check
```

Expected: every command exits 0.

- [ ] **Step 4: Packaging introspection**

Inspect `npm pack --dry-run` output and confirm:

- no `lib/core/atoms.internal.*`;
- no Jotai references;
- includes `lib/core/runtime-store.internal.*`;
- includes `lib/core/state.internal.*`;
- includes `lib/services/**/*`;
- includes `lib/helpers/**/*`.

- [ ] **Step 5: Final grep**

Run:

```bash
rg -n "jotai|atomWithStorage|useAtom|useAtomValue|useSetAtom|spfxAtoms|atoms\\.internal" src package.json package-lock.json README.md docs --glob '!docs/superpowers/**'
```

Expected: exit 1 with no matches.
