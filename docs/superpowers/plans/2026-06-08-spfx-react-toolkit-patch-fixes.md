# SPFx React Toolkit Patch Fixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix the runtime, packaging, and documentation issues found in the audit while preserving public TypeScript signatures and the current dependency model.

**Architecture:** Treat this as a `2.0.x` patch release. Fix only internal behavior, package file inclusion, and stale docs; do not move React/SPFx packages between `dependencies` and `peerDependencies`, do not add dependencies, and do not change exported function/type names or parameter lists. Verification relies on the existing SPFx build, TypeScript, ESLint, package dry-run, and npm audit observation.

**Tech Stack:** SPFx 1.21.1, React 17, TypeScript 5.3, Jotai, PnPjs v4, Microsoft SPFx service scope APIs, npm packaging.

---

## Non-Goals

- Do not change `dependencies`, `devDependencies`, or `peerDependencies`.
- Do not introduce a new test framework or package.
- Do not rename exported hooks, providers, interfaces, or constants.
- Do not remove public exports from `src/index.ts`, `src/core/index.ts`, or `src/hooks/index.ts`.
- Do not upgrade SPFx, React, TypeScript, PnPjs, Fluent UI, or npm packages.
- Do not address npm audit by dependency upgrades in this patch.

## Files

- Modify: `src/core/provider-base.internal.tsx`
  - Gate theme subscription until `ServiceScope` is ready.
  - Reset readiness safely when `serviceScope` changes.
- Modify: `src/utils/theme-subscription.internal.ts`
  - Add an internal readiness parameter.
  - Avoid consuming `ThemeProvider` before the scope is finished.
- Modify: `src/hooks/useSPFxAadHttpClient.ts`
  - Add request sequencing to ignore stale `getClient()` resolutions.
- Modify: `src/hooks/useSPFxPnPSearch.ts`
  - Respect the full supplied `PnPContextInfo` instead of recreating from `siteUrl`.
- Modify: `src/hooks/useSPFxPnPList.ts`
  - Skip debounced refetch when no previous query exists.
  - Make batch partial failures deterministic and reject after setting hook error state.
- Modify: `package.json`
  - Narrow npm `files` to published library artifacts only.
  - Keep dependency sections unchanged.
- Modify: `docs/api/hooks/INDEX.md`
  - Fix provider example to pass `instance={this}`.
- Modify as needed after search: `docs/**/*.md`
  - Fix stale examples that pass `context={this.context}` or pass `hrContext.sp` where a `PnPContextInfo` is expected.

---

### Task 1: Protect Theme Subscription Behind ServiceScope Readiness

**Files:**
- Modify: `src/core/provider-base.internal.tsx`
- Modify: `src/utils/theme-subscription.internal.ts`

- [ ] **Step 1: Inspect current lines**

Run:

```bash
nl -ba src/core/provider-base.internal.tsx | sed -n '56,94p'
nl -ba src/utils/theme-subscription.internal.ts | sed -n '52,90p'
```

Expected: `serviceScope.whenFinished` exists in the provider, but `useThemeSubscription(context, setTheme)` is called before the render guard.

- [ ] **Step 2: Update readiness handling in the provider**

In `src/core/provider-base.internal.tsx`, replace the current `useEffect` that waits for `serviceScope` with:

```tsx
  React.useEffect(() => {
    let disposed = false;

    setIsScopeReady(false);

    if (!serviceScope) {
      setIsScopeReady(true);
      return () => {
        disposed = true;
      };
    }

    serviceScope.whenFinished(() => {
      if (!disposed) {
        setIsScopeReady(true);
      }
    });

    return () => {
      disposed = true;
    };
  }, [serviceScope]);
```

Then change the theme subscription call from:

```tsx
  useThemeSubscription(context, setTheme);
```

to:

```tsx
  useThemeSubscription(context, setTheme, isScopeReady);
```

- [ ] **Step 3: Update the internal theme subscription hook**

In `src/utils/theme-subscription.internal.ts`, change the function signature from:

```ts
export function useThemeSubscription(
  spfxContext: unknown,
  setTheme: (theme: IReadonlyTheme | undefined) => void
): void {
```

to:

```ts
export function useThemeSubscription(
  spfxContext: unknown,
  setTheme: (theme: IReadonlyTheme | undefined) => void,
  isScopeReady: boolean
): void {
```

Then make the effect start with:

```ts
  useEffect(() => {
    if (!isScopeReady) {
      return;
    }

    const themeProvider = getThemeProvider(spfxContext);
```

And update the dependency array from:

```ts
  }, [spfxContext, setTheme]);
```

to:

```ts
  }, [spfxContext, setTheme, isScopeReady]);
```

- [ ] **Step 4: Verify compile and lint for Task 1**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/core src/utils --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 2: Ignore Stale AadHttpClient Initializations

**Files:**
- Modify: `src/hooks/useSPFxAadHttpClient.ts`

- [ ] **Step 1: Inspect current initialization effect**

Run:

```bash
nl -ba src/hooks/useSPFxAadHttpClient.ts | sed -n '314,370p'
```

Expected: async `factory.getClient(resourceUrl)` updates state without checking whether the request is still current.

- [ ] **Step 2: Add request sequencing ref**

Below the existing `isMountedRef` declaration, add:

```ts
  const requestIdRef = useRef<number>(0);
```

- [ ] **Step 3: Guard the initialization effect**

Inside the `useEffect`, immediately after clearing client and error, add:

```ts
    const requestId = requestIdRef.current + 1;
    requestIdRef.current = requestId;
```

Then replace the success guard:

```ts
        if (isMountedRef.current) {
```

with:

```ts
        if (isMountedRef.current && requestId === requestIdRef.current) {
```

Replace the catch guard the same way:

```ts
        if (isMountedRef.current && requestId === requestIdRef.current) {
```

The resulting effect must only call `setClient`, `setInitError`, or `setIsInitializing(false)` for the latest `resourceUrl`.

- [ ] **Step 4: Verify compile and lint for Task 2**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/hooks/useSPFxAadHttpClient.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 3: Respect Injected PnP Context in Search Hook

**Files:**
- Modify: `src/hooks/useSPFxPnPSearch.ts`

- [ ] **Step 1: Inspect current context logic**

Run:

```bash
nl -ba src/hooks/useSPFxPnPSearch.ts | sed -n '586,594p'
```

Expected: the hook calls `useSPFxPnPContext(pnpContext?.siteUrl)` and loses the supplied `sp`, `error`, and config behavior.

- [ ] **Step 2: Replace context selection**

Replace:

```ts
  // Get PnP context
  const context = useSPFxPnPContext(pnpContext?.siteUrl);
  const { sp } = context;
```

with:

```ts
  // Get PnP context (use provided context or create default)
  const defaultContext = useSPFxPnPContext();
  const context = pnpContext || defaultContext;
  const { sp } = context;
```

This mirrors the pattern in `useSPFxPnPList` and keeps the public function signature unchanged.

- [ ] **Step 3: Verify compile and lint for Task 3**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/hooks/useSPFxPnPSearch.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 4: Make PnP List Refetch and Batch Errors Deterministic

**Files:**
- Modify: `src/hooks/useSPFxPnPList.ts`

- [ ] **Step 1: Inspect current refetch and batch logic**

Run:

```bash
nl -ba src/hooks/useSPFxPnPList.ts | sed -n '728,755p'
nl -ba src/hooks/useSPFxPnPList.ts | sed -n '884,1012p'
```

Expected: `debouncedRefetch` calls `refetch()` even when there is no previous query, and batch methods set errors but do not reject for partial failures.

- [ ] **Step 2: Skip debounced refetch with no previous query**

Replace the `debouncedRefetch` callback with:

```ts
  const debouncedRefetch = useCallback(() => {
    if (!lastQueryBuilder) {
      return;
    }

    if (refetchTimeoutRef.current) {
      clearTimeout(refetchTimeoutRef.current);
    }

    refetchTimeoutRef.current = setTimeout(function() {
      refetch().catch(function(err) {
        const error = err as Error;
        console.error('[useSPFxPnPList] Debounced refetch error:', error);
        setError(error);
      });
    }, 100);
  }, [lastQueryBuilder, refetch]);
```

- [ ] **Step 3: Replace `createBatch` with deterministic settlement**

Replace the body of `createBatch` after the initialization guard with:

```ts
    try {
      const batchResult = sp.batched();
      const batchedSP = batchResult[0];
      const execute = batchResult[1];
      const list = batchedSP.web.lists.getByTitle(listTitle);

      const operations = itemsToCreate.map(function(itemToCreate) {
        return list.items.add(itemToCreate as Record<string, unknown>);
      });

      await execute();

      const settled = await Promise.allSettled(operations);
      const ids: number[] = [];
      const errors: unknown[] = [];

      settled.forEach(function(result) {
        if (result.status === 'fulfilled') {
          ids.push(result.value.data.Id);
        } else {
          errors.push(result.reason);
        }
      });

      if (errors.length > 0) {
        const batchError = new Error(`Batch create failed: ${errors.length} of ${itemsToCreate.length} items failed`);
        setError(batchError);
        console.error('Batch create summary:', errors);
        throw batchError;
      }

      debouncedRefetch();
      return ids;
    } catch (err) {
      const error = err as Error;
      setError(error);
      throw error;
    }
```

- [ ] **Step 4: Replace `updateBatch` with deterministic settlement**

Replace the body of `updateBatch` after the initialization guard with:

```ts
    try {
      const batchResult = sp.batched();
      const batchedSP = batchResult[0];
      const execute = batchResult[1];
      const list = batchedSP.web.lists.getByTitle(listTitle);

      const operations = updates.map(function(updateItem) {
        return list.items.getById(updateItem.id).update(updateItem.item as Record<string, unknown>);
      });

      await execute();

      const settled = await Promise.allSettled(operations);
      const errors = settled
        .filter(function(result): result is PromiseRejectedResult {
          return result.status === 'rejected';
        })
        .map(function(result) {
          return result.reason;
        });

      if (errors.length > 0) {
        const batchError = new Error(`Batch update failed: ${errors.length} of ${updates.length} items failed`);
        setError(batchError);
        console.error('Batch update summary:', errors);
        throw batchError;
      }

      debouncedRefetch();
    } catch (err) {
      const error = err as Error;
      setError(error);
      throw error;
    }
```

- [ ] **Step 5: Replace `removeBatch` with deterministic settlement**

Replace the body of `removeBatch` after the initialization guard with:

```ts
    try {
      const batchResult = sp.batched();
      const batchedSP = batchResult[0];
      const execute = batchResult[1];
      const list = batchedSP.web.lists.getByTitle(listTitle);

      const operations = ids.map(function(id) {
        return list.items.getById(id).delete();
      });

      await execute();

      const settled = await Promise.allSettled(operations);
      const errors = settled
        .filter(function(result): result is PromiseRejectedResult {
          return result.status === 'rejected';
        })
        .map(function(result) {
          return result.reason;
        });

      if (errors.length > 0) {
        const batchError = new Error(`Batch delete failed: ${errors.length} of ${ids.length} items failed`);
        setError(batchError);
        console.error('Batch delete summary:', errors);
        throw batchError;
      }

      debouncedRefetch();
    } catch (err) {
      const error = err as Error;
      setError(error);
      throw error;
    }
```

- [ ] **Step 6: Verify compile and lint for Task 4**

Run:

```bash
npx tsc --noEmit --pretty false
npx eslint src/hooks/useSPFxPnPList.ts --ext .ts,.tsx --max-warnings=0
```

Expected: both commands exit `0`.

---

### Task 5: Narrow Published npm Files Without Dependency Changes

**Files:**
- Modify: `package.json`

- [ ] **Step 1: Inspect current package files**

Run:

```bash
nl -ba package.json | sed -n '1,40p'
npm pack --dry-run --cache /tmp/codex-npm-pack-cache
```

Expected before the fix: dry-run includes `lib/webparts/spFxReactToolkitTest/**` and `lib/extensions/spFxReactToolkitTest/**`.

- [ ] **Step 2: Replace the `files` array only**

In `package.json`, replace:

```json
  "files": [
    "lib/**/*",
    "README.md",
    "LICENSE"
  ],
```

with:

```json
  "files": [
    "lib/index.*",
    "lib/core/**/*",
    "lib/hooks/**/*",
    "lib/utils/**/*",
    "README.md",
    "LICENSE"
  ],
```

Do not edit any dependency section.

- [ ] **Step 3: Verify package contents**

Run:

```bash
npm pack --dry-run --cache /tmp/codex-npm-pack-cache
```

Expected:
- Includes `lib/index.*`, `lib/core/**`, `lib/hooks/**`, `lib/utils/**`, `README.md`, `LICENSE`, `package.json`.
- Does not include `lib/webparts/**`.
- Does not include `lib/extensions/**`.

---

### Task 6: Fix Stale Documentation Examples

**Files:**
- Modify: `docs/api/hooks/INDEX.md`
- Modify as needed: `docs/**/*.md`

- [ ] **Step 1: Search stale examples**

Run:

```bash
rg -n "SPFxWebPartProvider context=|SPFxProvider\\b|useSPFxPnPList\\([^\\n]+\\.sp\\)" docs README.md src
```

Expected: at least `docs/api/hooks/INDEX.md` contains `<SPFxWebPartProvider context={this.context}>`.

- [ ] **Step 2: Fix provider example in hooks index**

In `docs/api/hooks/INDEX.md`, replace:

```tsx
    <SPFxWebPartProvider context={this.context}>
      <MyComponent />
    </SPFxWebPartProvider>
```

with:

```tsx
    <SPFxWebPartProvider instance={this}>
      <MyComponent />
    </SPFxWebPartProvider>
```

- [ ] **Step 3: Fix stale generic provider references if present**

For documentation examples that show `SPFxProvider` as a public import, replace them with the concrete provider that matches the example. For WebPart examples, use:

```tsx
import { SPFxWebPartProvider } from '@apvee/spfx-react-toolkit';

const element = (
  <SPFxWebPartProvider instance={this}>
    <MyComponent />
  </SPFxWebPartProvider>
);
```

Do not add an exported `SPFxProvider` alias in code.

- [ ] **Step 4: Fix PnP context injection examples if present**

For examples that pass `hrContext.sp` into `useSPFxPnPList`, replace with the existing three-argument shape:

```tsx
const hrContext = useSPFxPnPContext('/sites/hr');
const { items: hrItems } = useSPFxPnPList('Employees', undefined, hrContext);
```

Do not change the hook signature.

- [ ] **Step 5: Re-run stale-doc search**

Run:

```bash
rg -n "SPFxWebPartProvider context=|SPFxProvider\\b|useSPFxPnPList\\([^\\n]+\\.sp\\)" docs README.md src
```

Expected: no stale example remains, except prose that explicitly says an old API is unavailable.

---

### Task 7: Full Verification

**Files:**
- No source edits unless verification reveals a directly related failure.

- [ ] **Step 1: Check worktree scope**

Run:

```bash
git status --short
```

Expected: only files from this plan are modified.

- [ ] **Step 2: TypeScript**

Run:

```bash
npx tsc --noEmit --pretty false
```

Expected: exits `0`.

- [ ] **Step 3: ESLint**

Run:

```bash
npx eslint src --ext .ts,.tsx --max-warnings=0
```

Expected: exits `0`.

- [ ] **Step 4: Existing SPFx test command**

Run:

```bash
npm test -- --no-color
```

Expected: exits `0`. Note in the final report that this command runs SPFx build/lint/tsc/webpack, not a unit test suite.

- [ ] **Step 5: Existing build**

Run:

```bash
npm run build -- --no-color
```

Expected: exits `0`.

- [ ] **Step 6: npm package dry-run**

Run:

```bash
npm pack --dry-run --cache /tmp/codex-npm-pack-cache
```

Expected: exits `0` and excludes `lib/webparts/**` and `lib/extensions/**`.

- [ ] **Step 7: npm audit observation**

Run:

```bash
npm audit --omit=dev --audit-level=moderate
```

Expected: may still fail with the existing SPFx transitive vulnerabilities. Do not apply dependency changes in this patch. Record the exact result in the final report.

---

### Task 8: Final Review and Release Classification

**Files:**
- No edits required unless review finds a directly related issue.

- [ ] **Step 1: Confirm no public API signature changes**

Run:

```bash
git diff -- src/index.ts src/core/index.ts src/hooks/index.ts src/core/types.ts
```

Expected: no changes, unless documentation-only comments changed. No exported symbol should be added, removed, or renamed.

- [ ] **Step 2: Confirm no dependency changes**

Run:

```bash
git diff -- package.json package-lock.json
```

Expected:
- `package.json` changes only the `files` array.
- `package-lock.json` is unchanged.
- No dependency section is edited.

- [ ] **Step 3: Summarize patch release scope**

Final summary should state:

```text
This is suitable for a 2.0.x patch release because public TypeScript signatures and dependency declarations are unchanged. Runtime behavior is corrected for ServiceScope readiness, stale AAD clients, injected PnP search context, PnP list refetch, and batch partial failures. npm package contents are narrowed to library artifacts only.
```

- [ ] **Step 4: Optional commit**

Only if the user asks to commit:

```bash
git add src/core/provider-base.internal.tsx src/utils/theme-subscription.internal.ts src/hooks/useSPFxAadHttpClient.ts src/hooks/useSPFxPnPSearch.ts src/hooks/useSPFxPnPList.ts package.json docs/api/hooks/INDEX.md docs
git commit -m "fix: harden SPFx toolkit runtime behavior"
```

Expected: commit succeeds.

---

## Self-Review

- Spec coverage: Covers all requested patch-compatible fixes from the audit: ServiceScope/theme, AAD race, PnP search context, PnP list refetch, batch partial failures, packaging, stale docs, and verification. Dependency-model cleanup and audit upgrades are explicitly deferred.
- Placeholder scan: No `TBD`, `TODO`, or unspecified "handle edge cases" steps remain.
- Type consistency: All modified public functions keep their existing names and signatures. Internal-only signature change is limited to `useThemeSubscription`, which is not exported from the package public index.
