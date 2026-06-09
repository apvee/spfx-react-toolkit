# SPFx API Permission Precheck Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a lightweight SPFx API permission precheck for Microsoft Graph and custom API delegated scopes, with helper/service layers first, React hooks above them, and complete docs/demo coverage.

**Architecture:** Implement pure helpers for config normalization, JWT parsing, scope evaluation, and error classification. Build a non-React service that uses a mockable SPFx AAD token provider shape. Add `useSPFxAadTokenProvider` and `useSPFxApiPermissionPrecheck` React facades, then update public docs and the SPFx test webpart.

**Tech Stack:** SPFx 1.21.1, React 17, TypeScript 5.3, `AadTokenProviderFactory`, Node verification scripts, existing SPFx build/test/doc verification.

---

## Source Spec

Implement the approved design in:

`docs/superpowers/specs/2026-06-09-spfx-api-permission-precheck-design.md`

Do not broaden scope beyond that spec. Version 1 supports delegated `scp` scopes only. Application permissions, Graph admin grant inspection, direct MSAL usage, and custom API app roles are out of scope.

## File Structure

- Create `src/helpers/spfx-api-permission-precheck.helpers.ts`
  - Owns pure public types, config expansion, validation, JWT decoding, scope extraction/evaluation, result summarization, remediation text, and error classification.
- Modify `src/helpers/index.ts`
  - Exports the new helper module.
- Create `src/services/spfx-api-permission-precheck.service.ts`
  - Owns non-React token acquisition orchestration using a minimal `SPFxAadTokenProviderLike`.
- Modify `src/services/index.ts`
  - Exports the new service.
- Create `src/hooks/useSPFxAadTokenProvider.ts`
  - Owns SPFx `AadTokenProviderFactory` consumption and async provider readiness.
- Create `src/hooks/useSPFxApiPermissionPrecheck.ts`
  - Owns React state, auto/manual checks, passive-mode auth event guard, and UI bucket result shape.
- Modify `src/hooks/index.ts`
  - Exports the new hooks.
- Create `scripts/verify-api-permission-precheck.cjs`
  - Compiles helper/service to a temporary CommonJS output and verifies pure behavior with Node assertions.
- Modify `package.json`
  - Adds `verify:api-permission-precheck`.
- Modify public docs:
  - `README.md`
  - `docs/INTRODUCTION.md`
  - `docs/INDEX.md`
  - `docs/api/helpers/INDEX.md`
  - `docs/api/services/INDEX.md`
  - `docs/api/hooks/INDEX.md`
  - `docs/api/hooks/http-clients.md`
- Modify test webpart:
  - `src/webparts/spFxReactToolkitTest/components/demoRegistry.ts`
  - `src/webparts/spFxReactToolkitTest/components/panels/ClientsPanel.tsx`

---

### Task 1: Pure Helper Layer And Verification Script

**Files:**
- Create: `src/helpers/spfx-api-permission-precheck.helpers.ts`
- Modify: `src/helpers/index.ts`
- Create: `scripts/verify-api-permission-precheck.cjs`
- Modify: `package.json`

- [ ] **Step 1: Create the helper module with public types**

Create `src/helpers/spfx-api-permission-precheck.helpers.ts`.

The file must export these types:

```ts
export type SPFxApiPermissionPrecheckMode = 'passive' | 'interactiveAllowed';

export type SPFxApiPermissionConfigurationState =
  | 'idle'
  | 'checking'
  | 'ready'
  | 'actionRequired'
  | 'cannotDetermine';

export type SPFxApiPermissionCheckStatus =
  | 'available'
  | 'missingScope'
  | 'consentRequired'
  | 'interactionRequired'
  | 'claimsChallenge'
  | 'conditionalAccessBlocked'
  | 'accessRestricted'
  | 'resourceMisconfigured'
  | 'tokenProviderUnavailable'
  | 'audienceMismatch'
  | 'unsupportedToken'
  | 'unsupportedPermissionKind'
  | 'tokenExpired'
  | 'transientFailure'
  | 'timeout'
  | 'invalidRequirement'
  | 'unknown';

export type SPFxApiPermissionSeverity = 'ok' | 'warning' | 'error' | 'unknown';

export interface SPFxPermissionScopeOptions {
  readonly scope: string;
  readonly required?: boolean;
  readonly satisfies?: readonly string[];
  readonly label?: string;
}

export type SPFxPermissionScopeInput = string | SPFxPermissionScopeOptions;

export interface SPFxCustomApiPermissionInput {
  readonly id?: string;
  readonly name: string;
  readonly resource: string;
  readonly packageResource?: string;
  readonly scopes: readonly SPFxPermissionScopeInput[];
  readonly expectedAudiences?: readonly string[];
}

export interface SPFxApiPermissionPrecheckConfig {
  readonly graph?: readonly SPFxPermissionScopeInput[];
  readonly customApis?: readonly SPFxCustomApiPermissionInput[];
  readonly requirements?: readonly SPFxApiPermissionRequirement[];
}

export interface SPFxApiPermissionRequirement {
  readonly id: string;
  readonly resourceName: string;
  readonly resourceEndpoint: string;
  readonly packageResource: string;
  readonly scope: string;
  readonly kind: 'delegatedScope';
  readonly required: boolean;
  readonly satisfies?: readonly string[];
  readonly expectedAudiences?: readonly string[];
  readonly adminMessage: string;
  readonly label?: string;
}

export interface SPFxApiPermissionPackageSolutionEntry {
  readonly resource: string;
  readonly scope: string;
}

export interface SPFxApiPermissionSummary {
  readonly id: string;
  readonly resourceName: string;
  readonly resourceEndpoint: string;
  readonly scope: string;
  readonly required: boolean;
  readonly status: SPFxApiPermissionCheckStatus;
  readonly severity: SPFxApiPermissionSeverity;
  readonly message: string;
  readonly adminMessage: string;
  readonly packageSolutionEntry: SPFxApiPermissionPackageSolutionEntry;
  readonly matchedScope?: string;
  readonly errorCode?: string;
}

export interface SPFxApiPermissionCheckResult extends SPFxApiPermissionSummary {
  readonly detectedScopes: readonly string[];
  readonly audience?: string;
  readonly tenantId?: string;
  readonly caseMismatchScope?: string;
}

export interface SPFxJwtPayload {
  readonly aud?: string | readonly string[];
  readonly scp?: string;
  readonly roles?: readonly string[] | string;
  readonly tid?: string;
  readonly exp?: number;
  readonly nbf?: number;
  readonly [key: string]: unknown;
}
```

- [ ] **Step 2: Implement config normalization helpers**

Add these exported constants and functions:

```ts
export const SPFX_GRAPH_RESOURCE_NAME = 'Microsoft Graph';
export const SPFX_GRAPH_RESOURCE_ENDPOINT = 'https://graph.microsoft.com';
export const SPFX_GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';

export function normalizeSPFxApiPermissionRequirements(
  config: SPFxApiPermissionPrecheckConfig
): readonly SPFxApiPermissionRequirement[];

export function buildSPFxApiPermissionAdminMessage(
  packageResource: string,
  scope: string
): string;

export function createSPFxApiPermissionRequirementId(
  resourceEndpoint: string,
  scope: string,
  prefix?: string
): string;
```

Required behavior:

- Empty config returns `[]`.
- Graph shorthand expands to resource name `Microsoft Graph`, endpoint `https://graph.microsoft.com`, package resource `Microsoft Graph`, expected audiences `['https://graph.microsoft.com', '00000003-0000-0000-c000-000000000000']`.
- Custom API shorthand uses `name`, `resource`, `packageResource ?? name`, and each scope input.
- `required` defaults to `true`.
- String scopes are trimmed.
- IDs are stable and based on `graph:<scope>` for Graph and `api:<id-or-resource>:<scope>` for custom APIs.
- Duplicate `resourceEndpoint + scope` pairs are deduplicated, preserving the first occurrence.
- Explicit `requirements` are included after shorthand entries and deduplicated with the same rule.

- [ ] **Step 3: Implement validation, JWT, scope, and error helpers**

Add these exported functions:

```ts
export function validateSPFxApiPermissionRequirement(
  requirement: SPFxApiPermissionRequirement
): SPFxApiPermissionCheckResult | undefined;

export function decodeSPFxJwtPayload(token: string): SPFxJwtPayload | undefined;

export function extractSPFxDelegatedScopes(payload: SPFxJwtPayload): readonly string[];

export function evaluateSPFxApiPermissionRequirement(
  requirement: SPFxApiPermissionRequirement,
  payload: SPFxJwtPayload | undefined,
  options?: { readonly validateAudience?: boolean }
): SPFxApiPermissionCheckResult;

export function classifySPFxApiPermissionError(
  requirement: SPFxApiPermissionRequirement,
  error: unknown
): SPFxApiPermissionCheckResult;

export function summarizeSPFxApiPermissionResults(
  results: readonly SPFxApiPermissionCheckResult[],
  isChecking?: boolean
): {
  readonly configurationState: SPFxApiPermissionConfigurationState;
  readonly isConfigured: boolean;
  readonly available: readonly SPFxApiPermissionSummary[];
  readonly missing: readonly SPFxApiPermissionSummary[];
  readonly warnings: readonly SPFxApiPermissionSummary[];
  readonly unknown: readonly SPFxApiPermissionSummary[];
};
```

Required behavior:

- Invalid requirements return `invalidRequirement` without throwing.
- JWT decode uses base64url decoding and never throws for malformed input.
- Scope matching is exact and case-sensitive.
- `satisfies` is an explicit list of alternate scopes that satisfy the requirement.
- If a scope differs only by case, return `missingScope` plus `caseMismatchScope`.
- If `roles` exists without `scp`, return `unsupportedPermissionKind`.
- If `aud` mismatches `expectedAudiences` and `validateAudience !== false`, return `audienceMismatch`.
- Expired tokens return `tokenExpired`.
- Valid token with required scope returns `available`.
- Valid token missing required scope returns `missingScope`.
- Error classification maps:
  - `AADSTS65001` and consent strings to `consentRequired`.
  - interaction/login/MFA strings to `interactionRequired`.
  - claims challenge strings to `claimsChallenge`.
  - conditional access strings to `conditionalAccessBlocked`.
  - user assignment/access assignment strings to `accessRestricted`.
  - resource not found / invalid resource strings to `resourceMisconfigured`.
  - timeout strings to `timeout`.
  - network/throttle/temporarily unavailable strings to `transientFailure`.
  - fallback to `unknown`.
- `summarizeSPFxApiPermissionResults` returns:
  - `idle` for no results and not checking.
  - `checking` when `isChecking` is true.
  - `ready` when all required results are `available`.
  - `actionRequired` for required `missingScope` or `consentRequired`.
  - `cannotDetermine` for required unknown/transient/runtime/parser failures.

- [ ] **Step 4: Export the helper module**

Modify `src/helpers/index.ts`:

```ts
export * from './spfx-api-permission-precheck.helpers';
```

- [ ] **Step 5: Add focused Node verification script**

Create `scripts/verify-api-permission-precheck.cjs`.

The script must:

- Compile `src/helpers/spfx-api-permission-precheck.helpers.ts` and `src/services/spfx-api-permission-precheck.service.ts` once the service exists. For Task 1, compile only the helper if the service file is not present yet.
- Use `npx tsc`.
- Use a temporary output directory under `os.tmpdir()`.
- Assert:
  - Graph shorthand normalization.
  - Custom API normalization.
  - Deduplication.
  - JWT decode success and malformed token failure.
  - Exact `scp` match.
  - `satisfies` match.
  - Missing scope with case mismatch hint.
  - Audience mismatch.
  - App-role-only token classified as unsupported.
  - `AADSTS65001` classified as consent required.

Use this pattern:

```js
const assert = require('assert');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const root = path.resolve(__dirname, '..');
const outDir = path.join(os.tmpdir(), 'spfx-api-permission-precheck-verify');

fs.rmSync(outDir, { recursive: true, force: true });
fs.mkdirSync(outDir, { recursive: true });

const sourceFiles = [
  'src/helpers/spfx-api-permission-precheck.helpers.ts'
];

if (fs.existsSync(path.join(root, 'src/services/spfx-api-permission-precheck.service.ts'))) {
  sourceFiles.push('src/services/spfx-api-permission-precheck.service.ts');
}

execFileSync(
  'npx',
  [
    'tsc',
    ...sourceFiles,
    '--module',
    'commonjs',
    '--target',
    'es2020',
    '--skipLibCheck',
    '--esModuleInterop',
    '--rootDir',
    'src',
    '--outDir',
    outDir,
    '--pretty',
    'false'
  ],
  { cwd: root, stdio: 'inherit' }
);

const helpers = require(path.join(outDir, 'helpers/spfx-api-permission-precheck.helpers.js'));

function createJwt(payload) {
  const encoded = Buffer.from(JSON.stringify(payload))
    .toString('base64')
    .replace(/=/g, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');
  return `header.${encoded}.signature`;
}

const requirements = helpers.normalizeSPFxApiPermissionRequirements({
  graph: ['Sites.Read.All', { scope: 'Files.Read.All', satisfies: ['Files.ReadWrite.All'] }],
  customApis: [
    {
      id: 'orders',
      name: 'Orders API',
      resource: 'api://orders',
      scopes: ['Orders.Read']
    }
  ]
});

assert.strictEqual(requirements.length, 3);
assert.strictEqual(requirements[0].resourceName, 'Microsoft Graph');
assert.strictEqual(requirements[0].resourceEndpoint, 'https://graph.microsoft.com');
assert.strictEqual(requirements[2].packageResource, 'Orders API');

const deduped = helpers.normalizeSPFxApiPermissionRequirements({
  graph: ['User.Read', 'User.Read']
});
assert.strictEqual(deduped.length, 1);

const graphToken = createJwt({
  aud: 'https://graph.microsoft.com',
  scp: 'Sites.Read.All Files.ReadWrite.All User.Read',
  tid: 'tenant-id',
  exp: Math.floor(Date.now() / 1000) + 3600
});
const graphPayload = helpers.decodeSPFxJwtPayload(graphToken);
assert.ok(graphPayload);
assert.deepStrictEqual(
  helpers.extractSPFxDelegatedScopes(graphPayload),
  ['Sites.Read.All', 'Files.ReadWrite.All', 'User.Read']
);

assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(requirements[0], graphPayload).status,
  'available'
);
assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(requirements[1], graphPayload).matchedScope,
  'Files.ReadWrite.All'
);

const mismatchPayload = helpers.decodeSPFxJwtPayload(createJwt({
  aud: 'https://graph.microsoft.com',
  scp: 'sites.read.all',
  exp: Math.floor(Date.now() / 1000) + 3600
}));
const caseResult = helpers.evaluateSPFxApiPermissionRequirement(requirements[0], mismatchPayload);
assert.strictEqual(caseResult.status, 'missingScope');
assert.strictEqual(caseResult.caseMismatchScope, 'sites.read.all');

const wrongAudiencePayload = helpers.decodeSPFxJwtPayload(createJwt({
  aud: 'api://wrong',
  scp: 'Sites.Read.All',
  exp: Math.floor(Date.now() / 1000) + 3600
}));
assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(requirements[0], wrongAudiencePayload).status,
  'audienceMismatch'
);

const rolesOnlyPayload = helpers.decodeSPFxJwtPayload(createJwt({
  aud: 'https://graph.microsoft.com',
  roles: ['Sites.Read.All'],
  exp: Math.floor(Date.now() / 1000) + 3600
}));
assert.strictEqual(
  helpers.evaluateSPFxApiPermissionRequirement(requirements[0], rolesOnlyPayload).status,
  'unsupportedPermissionKind'
);

assert.strictEqual(
  helpers.classifySPFxApiPermissionError(requirements[0], new Error('AADSTS65001: consent required')).status,
  'consentRequired'
);

assert.strictEqual(helpers.decodeSPFxJwtPayload('not-a-jwt'), undefined);

console.log('api permission precheck verification passed');
```

- [ ] **Step 6: Add npm script**

Modify `package.json` scripts:

```json
"verify:api-permission-precheck": "node scripts/verify-api-permission-precheck.cjs"
```

- [ ] **Step 7: Run helper verification**

Run:

```bash
npm run verify:api-permission-precheck
```

Expected: exits `0` and prints `api permission precheck verification passed`.

- [ ] **Step 8: Commit Task 1**

```bash
git add src/helpers/spfx-api-permission-precheck.helpers.ts src/helpers/index.ts scripts/verify-api-permission-precheck.cjs package.json
git commit -m "feat: add api permission precheck helpers"
```

---

### Task 2: Non-React Service Layer

**Files:**
- Create: `src/services/spfx-api-permission-precheck.service.ts`
- Modify: `src/services/index.ts`
- Modify: `scripts/verify-api-permission-precheck.cjs`

- [ ] **Step 1: Create the service module**

Create `src/services/spfx-api-permission-precheck.service.ts`.

The service must export:

```ts
import {
  classifySPFxApiPermissionError,
  decodeSPFxJwtPayload,
  evaluateSPFxApiPermissionRequirement,
  normalizeSPFxApiPermissionRequirements,
  summarizeSPFxApiPermissionResults,
  validateSPFxApiPermissionRequirement,
  type SPFxApiPermissionCheckResult,
  type SPFxApiPermissionPrecheckConfig,
} from '../helpers/spfx-api-permission-precheck.helpers';

export interface SPFxAadTokenProviderLike {
  getToken: (
    resourceEndpoint: string,
    options?: { readonly useCachedToken?: boolean; readonly claims?: string }
  ) => Promise<string>;
}

export interface SPFxApiPermissionPrecheckServiceOptions {
  readonly useCachedToken?: boolean;
  readonly validateAudience?: boolean;
  readonly timeoutMs?: number;
}

export interface SPFxApiPermissionPrecheckService {
  check: (
    config: SPFxApiPermissionPrecheckConfig,
    options?: SPFxApiPermissionPrecheckServiceOptions
  ) => Promise<readonly SPFxApiPermissionCheckResult[]>;
}

export function createSPFxApiPermissionPrecheckService(
  tokenProvider: SPFxAadTokenProviderLike
): SPFxApiPermissionPrecheckService;
```

- [ ] **Step 2: Implement grouped token acquisition**

Service behavior:

- Normalize requirements with `normalizeSPFxApiPermissionRequirements(config)`.
- Validate requirements first. Invalid results are included and do not trigger token calls.
- Group valid requirements by `resourceEndpoint`.
- For each resource group, call `tokenProvider.getToken(resourceEndpoint, { useCachedToken })` once.
- Use `Promise.race` to classify timeout after `timeoutMs ?? 15000`.
- Decode the token with `decodeSPFxJwtPayload`.
- Evaluate each requirement in the group with `evaluateSPFxApiPermissionRequirement`.
- If token acquisition fails, classify each group requirement with `classifySPFxApiPermissionError`.
- Never expose the raw token.
- Do not call `summarizeSPFxApiPermissionResults` inside the service except in tests; the hook owns aggregation.

- [ ] **Step 3: Export the service**

Modify `src/services/index.ts`:

```ts
export * from './spfx-api-permission-precheck.service';
```

- [ ] **Step 4: Extend verification script for service behavior**

Modify `scripts/verify-api-permission-precheck.cjs` to require:

```js
const {
  createSPFxApiPermissionPrecheckService
} = require(path.join(outDir, 'services/spfx-api-permission-precheck.service.js'));
```

Add assertions:

- Two Graph scopes call `getToken` once.
- Custom API and Graph call `getToken` once per resource.
- Token provider rejection with `AADSTS65001` returns `consentRequired`.
- Timeout returns `timeout`.
- Invalid requirements are returned and do not call `getToken`.

- [ ] **Step 5: Run service verification**

Run:

```bash
npm run verify:api-permission-precheck
```

Expected: exits `0`.

- [ ] **Step 6: Commit Task 2**

```bash
git add src/services/spfx-api-permission-precheck.service.ts src/services/index.ts scripts/verify-api-permission-precheck.cjs
git commit -m "feat: add api permission precheck service"
```

---

### Task 3: SPFx AadTokenProvider Hook

**Files:**
- Create: `src/hooks/useSPFxAadTokenProvider.ts`
- Modify: `src/hooks/index.ts`

- [ ] **Step 1: Create the token provider hook**

Create `src/hooks/useSPFxAadTokenProvider.ts`.

The hook must import:

```ts
import { AadTokenProviderFactory } from '@microsoft/sp-http';
import type { AadTokenProvider } from '@microsoft/sp-http';
import { useEffect, useMemo, useRef, useState } from 'react';
import { useSPFxServiceScope } from './useSPFxServiceScope';
```

It must export:

```ts
export interface SPFxAadTokenProviderInfo {
  readonly tokenProvider: AadTokenProvider | undefined;
  readonly isInitializing: boolean;
  readonly initError: Error | undefined;
  readonly isReady: boolean;
}

export function useSPFxAadTokenProvider(): SPFxAadTokenProviderInfo;
```

Behavior:

- Consume `AadTokenProviderFactory.serviceKey` from `useSPFxServiceScope()`.
- Catch consume failures and surface `initError`.
- Call `factory.getTokenProvider()` asynchronously.
- Guard state updates with `isMountedRef`.
- Set `isInitializing` false when resolved or failed.
- Mirror the style of `useSPFxMSGraphClient` and `useSPFxAadHttpClient`.

- [ ] **Step 2: Export the hook**

Modify `src/hooks/index.ts`:

```ts
export * from './useSPFxAadTokenProvider';
```

- [ ] **Step 3: Run TypeScript/SPFx build check**

Run:

```bash
npm run build
```

Expected: exits `0`.

- [ ] **Step 4: Commit Task 3**

```bash
git add src/hooks/useSPFxAadTokenProvider.ts src/hooks/index.ts
git commit -m "feat: expose spfx aad token provider hook"
```

---

### Task 4: React Precheck Hook

**Files:**
- Create: `src/hooks/useSPFxApiPermissionPrecheck.ts`
- Modify: `src/hooks/index.ts`

- [ ] **Step 1: Create the precheck hook**

Create `src/hooks/useSPFxApiPermissionPrecheck.ts`.

The hook must export:

```ts
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { ISPEventObserver } from '@microsoft/sp-core-library';
import type { AadTokenProvider, BeforeRedirectEventArgs, PopupEventArgs } from '@microsoft/sp-http';
import {
  createSPFxApiPermissionPrecheckService,
} from '../services/spfx-api-permission-precheck.service';
import {
  summarizeSPFxApiPermissionResults,
  type SPFxApiPermissionCheckResult,
  type SPFxApiPermissionConfigurationState,
  type SPFxApiPermissionPrecheckConfig,
  type SPFxApiPermissionPrecheckMode,
  type SPFxApiPermissionSummary,
} from '../helpers/spfx-api-permission-precheck.helpers';
import { useSPFxAadTokenProvider } from './useSPFxAadTokenProvider';
import { useSPFxInstanceInfo } from './useSPFxInstanceInfo';
```

Types:

```ts
export interface SPFxApiPermissionPrecheckOptions {
  readonly autoCheck?: boolean;
  readonly mode?: SPFxApiPermissionPrecheckMode;
  readonly useCachedToken?: boolean;
  readonly validateAudience?: boolean;
  readonly timeoutMs?: number;
}

export interface SPFxApiPermissionPrecheckResult {
  readonly isChecking: boolean;
  readonly isConfigured: boolean;
  readonly configurationState: SPFxApiPermissionConfigurationState;
  readonly available: readonly SPFxApiPermissionSummary[];
  readonly missing: readonly SPFxApiPermissionSummary[];
  readonly warnings: readonly SPFxApiPermissionSummary[];
  readonly unknown: readonly SPFxApiPermissionSummary[];
  readonly results: readonly SPFxApiPermissionCheckResult[];
  readonly tokenProviderError: Error | undefined;
  readonly check: () => Promise<readonly SPFxApiPermissionCheckResult[]>;
  readonly retry: () => Promise<readonly SPFxApiPermissionCheckResult[]>;
  readonly retryWithoutCache: () => Promise<readonly SPFxApiPermissionCheckResult[]>;
}

export function useSPFxApiPermissionPrecheck(
  config: SPFxApiPermissionPrecheckConfig,
  options?: SPFxApiPermissionPrecheckOptions
): SPFxApiPermissionPrecheckResult;
```

- [ ] **Step 2: Implement passive auth event guard**

Inside the hook, implement a local helper that registers handlers only for the duration of one check when `mode === 'passive'`.

Required behavior:

- Use an `ISPEventObserver` object with `instanceId`, `componentId`, `isDisposed`, and `dispose()`.
- Register `tokenProvider.popupEvent.add(observer, popupHandler)` and call `eventArgs.cancel(new Error('SPFx API permission precheck passive mode blocked an authentication popup.'))`.
- Register `tokenProvider.onBeforeRedirectEvent.add(observer, redirectHandler)` and call `eventArgs.cancel()`.
- Remove both handlers and dispose the observer in `finally`.
- If event APIs are unavailable in a future runtime, skip registration and allow service error classification to handle the result.

- [ ] **Step 3: Implement check/retry behavior**

Required behavior:

- Default options:
  - `autoCheck: true`
  - `mode: 'passive'`
  - `useCachedToken: true`
  - `validateAudience: true`
  - `timeoutMs: 15000`
- If `tokenProvider` is unavailable, return a single `tokenProviderUnavailable` result for each normalized requirement, or an empty `idle` result for empty config.
- Use `createSPFxApiPermissionPrecheckService(tokenProvider)`.
- `check()` uses current `useCachedToken`.
- `retry()` delegates to `check()`.
- `retryWithoutCache()` runs the service with `useCachedToken: false`.
- `autoCheck` runs once when token provider is ready and the config/options key changes.
- Avoid stale state updates with a request id ref and mounted ref.
- Use `summarizeSPFxApiPermissionResults(results, isChecking)` to derive buckets and `configurationState`.

- [ ] **Step 4: Export the hook**

Modify `src/hooks/index.ts`:

```ts
export * from './useSPFxApiPermissionPrecheck';
```

- [ ] **Step 5: Run verification**

Run:

```bash
npm run verify:api-permission-precheck
npm run build
```

Expected: both exit `0`.

- [ ] **Step 6: Commit Task 4**

```bash
git add src/hooks/useSPFxApiPermissionPrecheck.ts src/hooks/index.ts
git commit -m "feat: add spfx api permission precheck hook"
```

---

### Task 5: Public Documentation

**Files:**
- Modify: `README.md`
- Modify: `docs/INTRODUCTION.md`
- Modify: `docs/INDEX.md`
- Modify: `docs/api/helpers/INDEX.md`
- Modify: `docs/api/services/INDEX.md`
- Modify: `docs/api/hooks/INDEX.md`
- Modify: `docs/api/hooks/http-clients.md`

- [ ] **Step 1: Update README feature list and example**

In `README.md`:

- Add API permission precheck to public services/hooks feature descriptions.
- Add a compact example:

```ts
import { useSPFxApiPermissionPrecheck } from '@apvee/spfx-react-toolkit';

const precheck = useSPFxApiPermissionPrecheck({
  graph: ['Sites.Read.All'],
  customApis: [
    {
      name: 'Orders API',
      resource: 'api://contoso-orders-api',
      packageResource: 'Orders API',
      scopes: ['Orders.Read']
    }
  ]
});
```

Include the warning:

```md
`available` means the current SPFx runtime obtained a token with the delegated scope. It does not read tenant grants and does not replace server-side authorization.
```

- [ ] **Step 2: Update docs overview pages**

In `docs/INTRODUCTION.md` and `docs/INDEX.md`:

- Add `useSPFxAadTokenProvider`.
- Add `useSPFxApiPermissionPrecheck`.
- Add `createSPFxApiPermissionPrecheckService`.
- Mention that the check maps remediation back to `webApiPermissionRequests`.

- [ ] **Step 3: Update helper docs**

In `docs/api/helpers/INDEX.md`, add an "API Permission Precheck Helpers" section.

It must mention every exported helper function from `src/helpers/spfx-api-permission-precheck.helpers.ts`, including:

- `normalizeSPFxApiPermissionRequirements`
- `buildSPFxApiPermissionAdminMessage`
- `createSPFxApiPermissionRequirementId`
- `validateSPFxApiPermissionRequirement`
- `decodeSPFxJwtPayload`
- `extractSPFxDelegatedScopes`
- `evaluateSPFxApiPermissionRequirement`
- `classifySPFxApiPermissionError`
- `summarizeSPFxApiPermissionResults`

- [ ] **Step 4: Update service docs**

In `docs/api/services/INDEX.md`:

- Add dependency matrix entry for `createSPFxApiPermissionPrecheckService`.
- Document `SPFxAadTokenProviderLike`, `SPFxApiPermissionPrecheckServiceOptions`, and `SPFxApiPermissionPrecheckService`.
- Explain grouping by resource and that raw tokens are never exposed.

- [ ] **Step 5: Update hook docs**

In `docs/api/hooks/INDEX.md` and `docs/api/hooks/http-clients.md`:

- Add `useSPFxAadTokenProvider`.
- Add `useSPFxApiPermissionPrecheck`.
- Include simple and manual examples.
- Explain passive mode and `retryWithoutCache`.
- Avoid the words `approved` or `notApproved` for result semantics.

- [ ] **Step 6: Run docs verification**

Run:

```bash
npm run verify:public-docs
```

Expected: exits `0`.

- [ ] **Step 7: Commit Task 5**

```bash
git add README.md docs/INTRODUCTION.md docs/INDEX.md docs/api/helpers/INDEX.md docs/api/services/INDEX.md docs/api/hooks/INDEX.md docs/api/hooks/http-clients.md
git commit -m "docs: document spfx api permission precheck"
```

---

### Task 6: SPFx Test Webpart Surface

**Files:**
- Modify: `src/webparts/spFxReactToolkitTest/components/demoRegistry.ts`
- Modify: `src/webparts/spFxReactToolkitTest/components/panels/ClientsPanel.tsx`

- [ ] **Step 1: Add demo registry coverage**

In `src/webparts/spFxReactToolkitTest/components/demoRegistry.ts`, add coverage entries under the `clients` panel:

```ts
{ symbol: 'useSPFxAadTokenProvider', kind: 'read', notes: 'AAD token provider initialization status.' },
{ symbol: 'useSPFxApiPermissionPrecheck', kind: 'read', notes: 'Passive Graph/custom API delegated scope precheck.' },
```

- [ ] **Step 2: Add token provider and precheck UI to ClientsPanel**

Modify `ClientsPanel.tsx`:

- Import `useSPFxAadTokenProvider` and `useSPFxApiPermissionPrecheck`.
- Initialize:

```ts
const aadTokenProvider = useSPFxAadTokenProvider();
const permissionPrecheck = useSPFxApiPermissionPrecheck(
  {
    graph: ['User.Read'],
    customApis: aadResource.trim()
      ? [
          {
            name: 'Custom API',
            resource: aadResource.trim(),
            scopes: ['user_impersonation']
          }
        ]
      : []
  },
  {
    autoCheck: false,
    mode: 'passive'
  }
);
```

- Add a status badge for `AadTokenProvider`.
- Add a `DemoCard` titled `API Permission Precheck`.
- Include a primary button calling `permissionPrecheck.check`.
- Include a default button calling `permissionPrecheck.retryWithoutCache`.
- Render `configurationState`, `isConfigured`, missing messages, warnings, and `JsonDetails` for detailed results.
- Do not auto-run the precheck in the demo.

- [ ] **Step 3: Run example coverage verification**

Run:

```bash
npm run verify:examples
```

Expected: exits `0`.

- [ ] **Step 4: Run build**

Run:

```bash
npm run build
```

Expected: exits `0`.

- [ ] **Step 5: Commit Task 6**

```bash
git add src/webparts/spFxReactToolkitTest/components/demoRegistry.ts src/webparts/spFxReactToolkitTest/components/panels/ClientsPanel.tsx
git commit -m "test: add api permission precheck demo"
```

---

### Task 7: Final Verification And Cleanup

**Files:**
- Review all files changed by Tasks 1-6.

- [ ] **Step 1: Run focused verification**

Run:

```bash
npm run verify:api-permission-precheck
```

Expected: exits `0`.

- [ ] **Step 2: Run public docs verification**

Run:

```bash
npm run verify:public-docs
```

Expected: exits `0`.

- [ ] **Step 3: Run demo coverage verification**

Run:

```bash
npm run verify:examples
```

Expected: exits `0`.

- [ ] **Step 4: Run SPFx build**

Run:

```bash
npm run build
```

Expected: exits `0`.

- [ ] **Step 5: Run SPFx test pipeline**

Run:

```bash
npm test
```

Expected: exits `0`.

- [ ] **Step 6: Check formatting whitespace**

Run:

```bash
git diff --check
```

Expected: exits `0`.

- [ ] **Step 7: Inspect final diff**

Run:

```bash
git status --short
git log --oneline -8
```

Expected:

- No unstaged generated files that should be ignored.
- Recent commits correspond to the feature tasks and docs design commit.

- [ ] **Step 8: Commit final verification fixes if needed**

If final verification required fixes, stage only the concrete files changed by those fixes and commit them with:

```bash
git commit -m "fix: stabilize api permission precheck verification"
```

If no fixes were needed, do not create an empty commit.
