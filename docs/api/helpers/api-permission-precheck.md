# API Permission Precheck Helpers

[Helpers and services](../../guides/helpers-and-services.md) · [Precheck hook](../hooks/http-clients/api-permission-precheck.md) · [Precheck service](../services/api-permission-precheck.md)


Import:

```ts
import {
  buildSPFxApiPermissionAdminMessage,
  classifySPFxApiPermissionError,
  createSPFxApiPermissionRequirementId,
  decodeSPFxJwtPayload,
  evaluateSPFxApiPermissionRequirement,
  extractSPFxDelegatedScopes,
  normalizeSPFxApiPermissionRequirements,
  summarizeSPFxApiPermissionResults,
  validateSPFxApiPermissionRequirement,
} from '@apvee/spfx-react-toolkit';
```

These helpers power the API permission precheck hooks and services without reading React context or acquiring tokens themselves. They normalize Graph and custom API configuration, decode delegated `scp` scopes from JWT payloads, classify token acquisition errors, and summarize results for UI buckets and remediation.

| Function | Purpose |
|----------|---------|
| `normalizeSPFxApiPermissionRequirements` | Expands Graph shorthand, custom API shorthand, and explicit requirements into deduplicated delegated-scope requirements. |
| `buildSPFxApiPermissionAdminMessage` | Builds administrator-facing remediation text for a package resource and scope. |
| `createSPFxApiPermissionRequirementId` | Creates stable requirement identifiers such as `graph:Sites.Read.All` or custom API IDs. |
| `validateSPFxApiPermissionRequirement` | Returns an invalid-result object when a requirement is incomplete or uses an unsupported permission kind. |
| `decodeSPFxJwtPayload` | Decodes the payload section of a JWT without validating or exposing the raw token. |
| `extractSPFxDelegatedScopes` | Reads delegated scopes from the JWT `scp` claim. |
| `evaluateSPFxApiPermissionRequirement` | Evaluates one requirement against a decoded token payload and returns a check result. |
| `classifySPFxApiPermissionError` | Maps token acquisition failures to precheck statuses such as consent, interaction, timeout, and transient failure. |
| `summarizeSPFxApiPermissionResults` | Groups check results into `available`, `missing`, `warnings`, and `unknown` summaries with an overall configuration state. |

`normalizeSPFxApiPermissionRequirements` also carries `packageResource` and `scope` into each requirement so remediation can point back to the same resource/scope values used in `webApiPermissionRequests`.


## Configuration types

All properties below are readonly. Import these types from `@apvee/spfx-react-toolkit` alongside the helpers.

### `SPFxApiPermissionPrecheckConfig`

Input to `normalizeSPFxApiPermissionRequirements`, with three optional collections:

| Property | Type | Meaning |
|----------|------|---------|
| `graph` | `readonly SPFxPermissionScopeInput[]` | Microsoft Graph delegated scopes; normalization supplies Graph resource names, endpoint and expected audiences. |
| `customApis` | `readonly SPFxCustomApiPermissionInput[]` | Custom API definitions expanded to one requirement per scope. |
| `requirements` | `readonly SPFxApiPermissionRequirement[]` | Fully specified requirements appended without shorthand normalization. |

Omitted collections contribute no requirements. Normalization processes `graph`, then `customApis`, then `requirements`; it deduplicates by the exact resource endpoint and scope pair, retaining the first occurrence. It does not merge duplicate metadata or validate requirements.

### `SPFxPermissionScopeInput` and `SPFxPermissionScopeOptions`

`SPFxPermissionScopeInput` is either a scope string or a `SPFxPermissionScopeOptions` object:

| Property | Type | Meaning/default |
|----------|------|-----------------|
| `scope` | `string` | Required delegated scope name. Shorthand normalization trims surrounding whitespace. |
| `required` | `boolean` (optional) | Defaults to `true`; `false` makes an unavailable permission an optional warning in the summary. |
| `satisfies` | `readonly string[]` (optional) | Additional scope names that may satisfy this requirement. Normalization trims entries and removes empty entries. No scope hierarchy is inferred. |
| `label` | `string` (optional) | Display metadata retained on the normalized requirement. |

A string is shorthand for `{ scope: value, required: true }`. Matching against token scopes is case-sensitive. A scope differing only in case remains missing and is recorded as `caseMismatchScope`.

### `SPFxCustomApiPermissionInput`

| Property | Type | Meaning/default |
|----------|------|-----------------|
| `id` | `string` (optional) | Prefix component used to generate requirement IDs; falls back to `resource`. |
| `name` | `string` | Required readable API name. |
| `resource` | `string` | Required resource endpoint passed to the token provider. |
| `packageResource` | `string` (optional) | Resource name used in `webApiPermissionRequests` and administrator remediation; falls back to `name`. |
| `scopes` | `readonly SPFxPermissionScopeInput[]` | Required scope collection, expanded to individual requirements. |
| `expectedAudiences` | `readonly string[]` (optional) | Accepted JWT `aud` values. No audience check is performed when absent or empty. |

Use the resource and audience values configured for your API; normalization does not infer custom API audiences. Generated IDs have the form `api:<id-or-resource>:<scope>`.

### `SPFxApiPermissionRequirement`

The normalized or explicitly supplied requirement consumed by validation and evaluation:

| Property | Type | Meaning |
|----------|------|---------|
| `id` | `string` | Nonempty identifier used in results. |
| `resourceName` | `string` | Nonempty readable API name. |
| `resourceEndpoint` | `string` | Nonempty resource used for token acquisition. |
| `packageResource` | `string` | Nonempty resource name for package permission requests and remediation. |
| `scope` | `string` | Nonempty delegated scope to check. |
| `kind` | `'delegatedScope'` | The supported permission kind; application permissions are unsupported. |
| `required` | `boolean` | Whether availability is required for the overall configured state. |
| `satisfies` | `readonly string[]` (optional) | Explicit alternative scopes accepted by evaluation. |
| `expectedAudiences` | `readonly string[]` (optional) | Accepted token audiences when audience validation is enabled. |
| `adminMessage` | `string` | Administrator-facing remediation text. Shorthand normalization generates it with `buildSPFxApiPermissionAdminMessage`. |
| `label` | `string` (optional) | Optional display label; not copied into result summaries. |

Explicit requirements must provide the required fields themselves. `validateSPFxApiPermissionRequirement` checks the nonempty identity/resource/scope fields and supported `kind`; it returns `undefined` on success or a result with `invalidRequirement` or `unsupportedPermissionKind` on failure. This is not a general runtime schema validator.

## Result and state types

### `SPFxApiPermissionPrecheckMode`

`'passive' | 'interactiveAllowed'` configures the [precheck hook](../hooks/http-clients/api-permission-precheck.md). Its default is `passive`, which registers cancellation handlers for authentication popup/redirect events when supported by the token provider. `interactiveAllowed` permits the provider's normal authentication behavior. The standalone helpers do not acquire tokens or enforce an interaction mode; the [precheck service](../services/api-permission-precheck.md) accepts a token provider supplied by its caller.

### `SPFxApiPermissionCheckStatus`

The complete diagnostic status union is:

| Status | Meaning |
|--------|---------|
| `available` | The token contains the requested scope or an explicit alternative. |
| `missingScope` | No acceptable delegated scope matched. |
| `consentRequired` | Token acquisition indicates consent is needed. |
| `interactionRequired` | Token acquisition requires user interaction. |
| `claimsChallenge` | Token acquisition indicates a claims challenge. |
| `conditionalAccessBlocked` | Token acquisition indicates a Conditional Access block. |
| `accessRestricted` | Token acquisition indicates an assignment restriction. |
| `resourceMisconfigured` | Token acquisition could not resolve the resource. |
| `tokenProviderUnavailable` | The hook could not obtain a token provider. |
| `audienceMismatch` | No token audience matches the configured accepted audiences. |
| `unsupportedToken` | The token payload could not be decoded. |
| `unsupportedPermissionKind` | The requirement kind is unsupported, or the token contains app roles without delegated scopes. |
| `tokenExpired` | The token is expired or not yet valid according to `exp`/`nbf`. |
| `transientFailure` | Token acquisition indicates a network, throttling or temporary failure. |
| `timeout` | Token acquisition timed out. |
| `invalidRequirement` | A required identity/resource/scope field or permission kind is missing. |
| `unknown` | The permission state could not be determined. |

`classifySPFxApiPermissionError` classifies error text and extracts an optional error code; these labels are diagnostics rather than a definitive diagnosis from the identity provider.

### `SPFxApiPermissionSeverity`

`'ok' | 'warning' | 'error' | 'unknown'`. An `available` result is `ok`. Statuses `unknown`, `transientFailure` and `timeout` have severity `unknown`; other unavailable results are `warning` for optional requirements and `error` for required ones.

### `SPFxApiPermissionConfigurationState`

`'idle' | 'checking' | 'ready' | 'actionRequired' | 'cannotDetermine'` describes the result of `summarizeSPFxApiPermissionResults`:

- `checking` takes precedence when its `isChecking` argument is truthy.
- `idle` means there are no results and no check is running.
- `actionRequired` means at least one required result is in the `missing` bucket.
- `ready` means all required results are `available`.
- `cannotDetermine` means required results remain unresolved without an actionable missing result.

The helper places successful results in `available`, all unavailable optional results in `warnings`, and required results with `missingScope`, `consentRequired`, `interactionRequired`, `claimsChallenge`, `conditionalAccessBlocked` or `accessRestricted` in `missing`. Other unavailable required results go in `unknown`. Buckets and severity are separate classifications. `isConfigured` is true when every required result is available, including when there are no required results; the state may still be `idle` or `checking`.

### `SPFxApiPermissionPackageSolutionEntry`

A readonly `{ resource: string; scope: string }` pair suitable for a `webApiPermissionRequests` entry. Results derive `resource` from the requirement's `packageResource` and retain its requested `scope`; this metadata does not add requests to the package or approve them.

### `SPFxApiPermissionSummary`

The summary shape used in each result bucket:

| Property | Type | Meaning |
|----------|------|---------|
| `id` | `string` | Requirement identifier. |
| `resourceName` | `string` | Readable API name. |
| `resourceEndpoint` | `string` | Token acquisition resource. |
| `scope` | `string` | Requested delegated scope. |
| `required` | `boolean` | Whether availability is required. |
| `status` | `SPFxApiPermissionCheckStatus` | Diagnostic outcome. |
| `severity` | `SPFxApiPermissionSeverity` | Severity derived from status and requirement. |
| `message` | `string` | Explanation of the outcome. |
| `adminMessage` | `string` | Administrator remediation text. |
| `packageSolutionEntry` | `SPFxApiPermissionPackageSolutionEntry` | Matching package resource/scope metadata. |
| `matchedScope` | `string` (optional) | Requested or alternative scope that matched. |
| `errorCode` | `string` (optional) | Code extracted from a token acquisition error. |

### `SPFxApiPermissionCheckResult`

Extends `SPFxApiPermissionSummary` with `detectedScopes: readonly string[]` (empty when none were detected), optional `audience: string` (the string audience or first array entry), optional `tenantId: string` (the `tid` claim), and optional `caseMismatchScope: string` (a detected scope matching only without case sensitivity). Audience evaluation considers all array entries even though the diagnostic `audience` field displays only the first.

## JWT payload and Graph constants

### `SPFxJwtPayload`

The decoded payload shape contains optional `aud: string | readonly string[]`, `scp: string`, `roles: readonly string[] | string`, `tid: string`, `exp: number` and `nbf: number`, plus a readonly string index signature with `unknown` values for additional claims. `exp` and `nbf` are interpreted as Unix seconds. `scp` is split on spaces into delegated scopes; app roles do not satisfy delegated requirements.

`decodeSPFxJwtPayload` parses a JWT payload and returns `undefined` for a decoding failure or a non-object payload. It does not verify signatures or validate the claim schema. Evaluation checks supplied expiry/not-before claims and, by default, configured audiences, but this remains diagnostic-only token interpretation. An `available` result is not proof that the target API will authorize a request, and the helpers must not serve as a security boundary. Acquire real tokens through the [hook](../hooks/http-clients/api-permission-precheck.md) or [service](../services/api-permission-precheck.md) and let the API enforce authorization.

| Constant | Value | Use |
|----------|-------|-----|
| `SPFX_GRAPH_RESOURCE_NAME` | `'Microsoft Graph'` | Readable and package resource name for Graph shorthand. |
| `SPFX_GRAPH_RESOURCE_ENDPOINT` | `'https://graph.microsoft.com'` | Graph token resource endpoint and an accepted audience. |
| `SPFX_GRAPH_APP_ID` | `'00000003-0000-0000-c000-000000000000'` | Additional accepted Graph audience. |

## Example: normalize and inspect requirements

```ts
import {
  evaluateSPFxApiPermissionRequirement,
  normalizeSPFxApiPermissionRequirements,
  summarizeSPFxApiPermissionResults,
  SPFX_GRAPH_APP_ID,
  type SPFxApiPermissionPrecheckConfig,
  type SPFxJwtPayload,
} from '@apvee/spfx-react-toolkit';

const config: SPFxApiPermissionPrecheckConfig = {
  graph: [
    'Sites.Read.All',
    { scope: 'User.Read', required: false, label: 'Profile access' },
  ],
  customApis: [{
    id: 'orders',
    name: 'Orders API',
    resource: 'api://11111111-1111-1111-1111-111111111111',
    packageResource: 'Orders API',
    scopes: [{ scope: 'Orders.Read', satisfies: ['Orders.ReadWrite'] }],
    expectedAudiences: ['11111111-1111-1111-1111-111111111111'],
  }],
};

const requirements = normalizeSPFxApiPermissionRequirements(config);
// Illustrative decoded claims only; use the hook/service for real tokens.
const graphPayload: SPFxJwtPayload = {
  aud: SPFX_GRAPH_APP_ID,
  scp: 'Sites.Read.All User.Read',
};
const graphResults = requirements
  .filter(requirement => requirement.id.startsWith('graph:'))
  .map(requirement =>
    evaluateSPFxApiPermissionRequirement(requirement, graphPayload)
  );
const graphSummary = summarizeSPFxApiPermissionResults(graphResults);
// Both Graph requirements are available: configurationState === 'ready'.
// Custom API requirements need their own resource token and are not checked here.
```
