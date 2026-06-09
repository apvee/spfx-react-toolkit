# Remove Jotai From SPFx React Toolkit

Date: 2026-06-09

## Goal

Remove Jotai completely from the toolkit implementation and package while preserving the documented public behavior of providers and hooks.

The replacement must keep:

- One isolated runtime state container per SPFx provider instance.
- Selective hook updates, so unrelated runtime state changes do not cause unnecessary re-renders.
- Bidirectional `useSPFxProperties` synchronization between SPFx properties and React consumers.
- Theme, display mode, container, Teams, and storage hook behavior.
- Existing public provider and hook APIs.

## Non-Goals

- Do not introduce another state-management dependency.
- Do not expose the new runtime store as public API.
- Do not preserve unofficial deep imports of `src/core/atoms.internal`.
- Do not make `useSPFxStorage` part of the provider runtime store.
- Do not change services/helpers created by the previous refactor except where docs or package files require consistency.

## Current Jotai Responsibilities

Jotai is currently used for:

- Provider-scoped state isolation via `createStore()` and `<Provider store={store}>`.
- Runtime atoms for `theme`, `displayMode`, `properties`, `containerEl`, `containerSize`, and `teams`.
- Hook-level subscriptions with atom granularity.
- `atomWithStorage` in `useSPFxLocalStorage` and `useSPFxSessionStorage`.

These responsibilities are internal. The public barrel exports do not expose `spfxAtoms`, but the package includes `lib/core/**/*`, so consumers could have unofficial deep imports.

## Selected Approach

Use an internal provider-scoped runtime store with selector subscriptions.

The provider creates a stable store object once per provider instance and passes it through an internal React context. The context value is stable; updates happen inside the store and notify subscribers.

Each runtime hook subscribes through a selector:

- `useSPFxThemeInfo` selects `theme`.
- `useSPFxDisplayMode` selects `displayMode`.
- `useSPFxProperties` selects `properties` and uses runtime actions to update properties.
- `useSPFxContainerInfo` selects `containerEl` and `containerSize`.
- `useSPFxTeams` selects `teams`.

This preserves the useful Jotai behavior without depending on Jotai.

## Runtime Store Design

Add the internal module `src/core/state.internal.tsx`.

State shape:

```ts
interface SPFxRuntimeState {
  readonly theme: IReadonlyTheme | undefined;
  readonly displayMode: DisplayMode | undefined;
  readonly properties: unknown;
  readonly containerEl: HTMLElement | undefined;
  readonly containerSize: ContainerSize | undefined;
  readonly teams: {
    readonly supported: boolean;
    readonly context?: unknown;
    readonly theme?: TeamsTheme;
    readonly initialized: boolean;
  };
}
```

Store API:

```ts
interface SPFxRuntimeStore {
  getState(): SPFxRuntimeState;
  setState(
    updater:
      | Partial<SPFxRuntimeState>
      | ((previous: SPFxRuntimeState) => SPFxRuntimeState)
  ): void;
  subscribe(listener: () => void): () => void;
}
```

Internal hooks:

```ts
function useSPFxRuntimeSelector<T>(
  selector: (state: SPFxRuntimeState) => T
): T;

function useSPFxRuntimeActions(): {
  setTheme(theme: IReadonlyTheme | undefined): void;
  setDisplayMode(displayMode: DisplayMode | undefined): void;
  setProperties(
    updater: unknown | ((previous: unknown) => unknown)
  ): void;
  setContainerElement(element: HTMLElement | undefined): void;
  setContainerSize(size: ContainerSize | undefined): void;
  setTeamsState(state: SPFxRuntimeState["teams"]): void;
};
```

The selector hook should compare selected values with `Object.is`. A hook re-renders only when its selected slice changes.

## Provider Behavior

`SPFxProviderBase` keeps the same public inputs and rendering contract.

It changes internally from:

- creating a Jotai store;
- writing atoms;
- rendering `<Provider store={store}>`;

to:

- creating an `SPFxRuntimeStore`;
- writing runtime state through internal actions;
- rendering `<SPFxRuntimeStoreContext.Provider value={store}>`.

The existing `SPFxContext.Provider` remains, and provider-specific wrappers continue to pass only `instance`.

The provider still:

- waits for `serviceScope.whenFinished`;
- initializes `properties`;
- initializes `displayMode` and `containerEl` for web parts;
- subscribes to SPFx theme changes;
- syncs changed SPFx properties into runtime state;
- syncs runtime property changes back into `instance.properties`;
- refreshes the web part property pane when properties change from React.

## Storage Hooks

`useSPFxLocalStorage` and `useSPFxSessionStorage` should not use the runtime store.

They should be implemented as local hooks with:

- lazy initial read from storage;
- JSON parse fallback to `defaultValue`;
- functional setter support;
- `remove()` that deletes the browser storage item and resets to `defaultValue`;
- safe no-op behavior when `localStorage` or `sessionStorage` is unavailable;
- browser `storage` event handling to sync updates from other tabs/windows.

Do not add an in-memory per-key registry. Same-page cross-component synchronization for identical storage keys is not currently documented and is outside this refactor.

## Edge Cases And Resolutions

### React 17

React 17 does not include `useSyncExternalStore`. Use `useState` plus `useEffect` subscription inside `useSPFxRuntimeSelector`. Do not add `use-sync-external-store`.

### Unnecessary Re-Renders

The runtime store context value must be stable. Do not place the whole runtime state in context. Hooks must subscribe through selectors and update only when the selected value changes by `Object.is`.

### Provider Isolation

The store must be created inside each `SPFxProviderBase` instance with a stable `useMemo` or `useRef`. There must be no module-level singleton store.

### Properties Sync Loops

Keep `lastPropertiesRef`. Sync from runtime state to SPFx only when the runtime `properties` reference differs from the last known SPFx reference. Do not sync when `properties === undefined`.

### Properties Object Mutation

Preserve existing behavior: when React updates properties, mutate `instance.properties` by clearing existing keys and copying keys from the runtime object. Then refresh the property pane for web parts when available.

### Theme Subscription

Keep `useThemeSubscription`. It should receive `setTheme` from runtime actions instead of a Jotai setter. Cleanup and initial theme read remain unchanged.

### Container Resize

`useSPFxContainerInfo` should select `containerEl` and `containerSize`, and pass `setContainerSize` to `useResizeObserver`. Resize updates must not re-render hooks that only read theme, properties, display mode, or Teams.

### Teams Initialization

`useSPFxTeams` should select `teams` and update it through runtime actions. Initialization should run once per provider while `teams.initialized` is false. Fallback unsupported state must be set when no Teams SDK exists or initialization fails.

### Consumer Jotai Provider Behavior

Removing `<Jotai.Provider store={store}>` can affect consumers who relied on the toolkit providing a Jotai store for their own atoms. This was not documented and is not part of the public toolkit contract. Mention it in release notes as an internal behavior removal.

### Deep Imports

Deleting `atoms.internal.ts` can break unofficial imports from `@apvee/spfx-react-toolkit/lib/core/atoms.internal`. This is acceptable as an internal cleanup, but it should be called out in release notes.

## Package And Documentation Changes

Remove:

- `jotai` from `dependencies`.
- `jotai` from `keywords`.
- `node_modules/jotai` entry from `package-lock.json`.
- Jotai references from README and docs.

Replace "Built on Jotai" language with "provider-scoped runtime store" or equivalent wording.

## Verification Plan

Minimum checks:

- `npx tsc --noEmit --pretty false`
- `npx eslint src --ext .ts,.tsx --max-warnings=0`
- `npm run build`
- `npm --cache /tmp/spfx-react-toolkit-npm-cache pack --dry-run`
- `rg -n "jotai|atomWithStorage|useAtom|useAtomValue|useSetAtom|spfxAtoms|atoms\\.internal" src package.json README.md docs`

Behavior-focused introspection:

- Confirm `src/core/atoms.internal.ts` is gone.
- Confirm no runtime hook imports Jotai.
- Confirm no provider renders a Jotai provider.
- Confirm package tarball no longer includes Jotai-dependent code.
- Confirm hooks still export the same public names and return types.

Focused manual review:

- `useSPFxProperties`: verify SPFx -> runtime and runtime -> SPFx sync paths.
- `useSPFxContainerInfo`: verify resize updates only container subscribers.
- `useSPFxTeams`: verify initialized guard and unsupported fallback.
- `useSPFxStorage`: verify local/session storage unavailable, invalid JSON, remove, and functional setter cases.

## Recommended Implementation Slices

1. Add internal runtime store and selector hooks.
2. Convert `SPFxProviderBase` to use the runtime store.
3. Convert runtime consumers: theme, display mode, properties, container, Teams.
4. Rewrite storage hooks without Jotai.
5. Remove `atoms.internal.ts` and package dependency references.
6. Update docs and package metadata.
7. Run verification and compare declaration output for public hook compatibility.

## Release Note

Jotai is no longer used or installed by the toolkit. Public provider and hook APIs are intended to remain unchanged. Unofficial deep imports from internal Jotai atom modules and implicit reliance on the toolkit's Jotai provider are no longer supported.
