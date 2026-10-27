# API permission precheck helpers

These helpers normalize requirements and interpret delegated token diagnostics without React or token acquisition. Import from the package root or `@apvee/spfx-react-toolkit/helpers`. The [hook](../hooks/http-clients.md#usespfxapipermissionprecheck) and [service](../services/INDEX.md#createspfxapipermissionprecheckservice) acquire tokens; these helpers operate on supplied values.

## Example

```ts
import {
  normalizeSPFxApiPermissionRequirements,
  evaluateSPFxApiPermissionRequirement,
  summarizeSPFxApiPermissionResults,
  SPFX_GRAPH_APP_ID,
} from '@apvee/spfx-react-toolkit/helpers';

const requirements = normalizeSPFxApiPermissionRequirements({
  graph: ['User.Read', { scope: 'Mail.Read', required: false }],
});
const payload = { aud: SPFX_GRAPH_APP_ID, scp: 'User.Read' };
const results = requirements.map(requirement =>
  evaluateSPFxApiPermissionRequirement(requirement, payload));
const summary = summarizeSPFxApiPermissionResults(results);
console.log(summary.configurationState); // ready; optional Mail.Read is a warning
```

The payload above is illustrative diagnostic data. Decoding does not verify signatures, issuer or identity. An available scope does not prove that an API request will succeed; server authorization and tenant policies still apply.

## Functions

```ts
export function normalizeSPFxApiPermissionRequirements(config: SPFxApiPermissionPrecheckConfig): readonly SPFxApiPermissionRequirement[];

export function buildSPFxApiPermissionAdminMessage(packageResource: string, scope: string): string;

export function createSPFxApiPermissionRequirementId(resourceEndpoint: string, scope: string, prefix?: string): string;

export function validateSPFxApiPermissionRequirement(requirement: SPFxApiPermissionRequirement): SPFxApiPermissionCheckResult | undefined;

export function decodeSPFxJwtPayload(token: string): SPFxJwtPayload | undefined;

export function extractSPFxDelegatedScopes(payload: SPFxJwtPayload): readonly string[];

export function evaluateSPFxApiPermissionRequirement(requirement: SPFxApiPermissionRequirement, payload: SPFxJwtPayload | undefined, options?: {
    readonly validateAudience?: boolean;
}): SPFxApiPermissionCheckResult;

export function classifySPFxApiPermissionError(requirement: SPFxApiPermissionRequirement, error: unknown): SPFxApiPermissionCheckResult;

export function summarizeSPFxApiPermissionResults(results: readonly SPFxApiPermissionCheckResult[], isChecking?: boolean): {
    readonly configurationState: SPFxApiPermissionConfigurationState;
    readonly isConfigured: boolean;
    readonly available: readonly SPFxApiPermissionSummary[];
    readonly missing: readonly SPFxApiPermissionSummary[];
    readonly warnings: readonly SPFxApiPermissionSummary[];
    readonly unknown: readonly SPFxApiPermissionSummary[];
};
```

| Function | Inputs, output and omission behavior |
| --- | --- |
| `normalizeSPFxApiPermissionRequirements` | Expands Graph, custom API, then explicit requirements in order. Omitted arrays contribute nothing. Duplicate resource-endpoint/scope pairs keep the first entry. Scope strings are trimmed and required by default; alternatives are explicit rather than inferred permission hierarchies. |
| `buildSPFxApiPermissionAdminMessage` | Returns a plain-text administrator instruction for the supplied package resource and scope. It does not edit package configuration. |
| `createSPFxApiPermissionRequirementId` | Graph uses `graph:<trimmed scope>`; custom APIs use `api:<trimmed prefix or endpoint>:<trimmed scope>`. Omitted prefix uses the endpoint. |
| `validateSPFxApiPermissionRequirement` | Returns undefined for complete delegated-scope metadata; otherwise returns `invalidRequirement` or `unsupportedPermissionKind` diagnostic data. No endpoint/consent check occurs. |
| `decodeSPFxJwtPayload` | Decodes the JWT payload locally. Invalid JSON, nonobject payloads and decoding failures return undefined. Claim types and signatures are not validated. Do not log raw tokens. |
| `extractSPFxDelegatedScopes` | Splits the string `scp` claim and removes empty entries; preserves case, order and duplicates. Missing/nonstring `scp` returns an empty array. Application roles do not count as delegated scopes. |
| `evaluateSPFxApiPermissionRequirement` | Checks metadata, numeric `exp`/`nbf` against the current clock, audience and exact case-sensitive scopes/alternatives. Audience validation defaults to enabled but needs configured expected audiences. Undefined payload returns `unsupportedToken`; a case-only mismatch remains `missingScope` with diagnostic detail. |
| `classifySPFxApiPermissionError` | Maps arbitrary acquisition failures using message/code heuristics. It returns diagnostic data without retrying, acquiring tokens or granting access. |
| `summarizeSPFxApiPermissionResults` | Returns token-free UI buckets and aggregate state. Omitted `isChecking` is false. Pending work is checking; no results is idle; optional failures do not block ready. `isConfigured` checks required results only, so it is true even for an empty result set. |

## Public configuration types

`SPFxPermissionScopeOptions.scope` is required and case-sensitive. Omitted `required` defaults to true; omitted `satisfies` accepts only the requested scope; `label` is optional display metadata. `SPFxPermissionScopeInput` accepts that object or a required-scope string.

`SPFxCustomApiPermissionInput.name`, `resource` and `scopes` are required. `resource` is passed to token acquisition; omitted/empty `id` uses that resource as the identity prefix, and omitted/empty `packageResource` uses name for package remediation. Omitted `expectedAudiences` skips custom API audience comparison. Graph normalization supplies both Graph audiences shown below.

`SPFxApiPermissionPrecheckConfig` accepts any combination of optional `graph`, `customApis` and `requirements`; an empty object is valid and produces no requirements. Explicit `SPFxApiPermissionRequirement` values must supply all required metadata shown in their declaration, including `kind: 'delegatedScope'`, `required` and administrator text.

```ts
export type SPFxApiPermissionPrecheckMode = 'passive' | 'interactiveAllowed';

export type SPFxApiPermissionConfigurationState = 'idle' | 'checking' | 'ready' | 'actionRequired' | 'cannotDetermine';

export type SPFxApiPermissionCheckStatus = 'available' | 'missingScope' | 'consentRequired' | 'interactionRequired' | 'claimsChallenge' | 'conditionalAccessBlocked' | 'accessRestricted' | 'resourceMisconfigured' | 'tokenProviderUnavailable' | 'audienceMismatch' | 'unsupportedToken' | 'unsupportedPermissionKind' | 'tokenExpired' | 'transientFailure' | 'timeout' | 'invalidRequirement' | 'unknown';

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

export const SPFX_GRAPH_RESOURCE_NAME = 'Microsoft Graph';

export const SPFX_GRAPH_RESOURCE_ENDPOINT = 'https://graph.microsoft.com';

export const SPFX_GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';
```

## Results and states

`SPFxApiPermissionSummary` carries the requirement identity, resource and scope, requiredness, status, severity, display message, administrator instruction and package request pair. `matchedScope` identifies the accepted requested/alternative scope, and `errorCode` is present only when extracted. `SPFxApiPermissionCheckResult` adds observed `detectedScopes`, optional audience/tenant claims and optional `caseMismatchScope`. Claims remain unverified; raw tokens are never included.

The `available` bucket contains successful scopes. Required failures in `missingScope`, `consentRequired`, `interactionRequired`, `claimsChallenge`, `conditionalAccessBlocked` or `accessRestricted` enter `missing` and produce `actionRequired`. Other unresolved required failures enter `unknown` and produce `cannotDetermine`. Optional failures enter `warnings`. Required actionable failures take precedence over unresolved failures.

Severity is separate from these buckets: available is ok; unknown/transientFailure/timeout have unknown severity; other failures are error when required and warning when optional. `SPFxApiPermissionPrecheckMode` is the token interaction policy used by the hook; these helpers do not launch interaction themselves.

`SPFxJwtPayload` retains additional claims as unknown values. Known optional claims are `aud`, `scp`, `roles`, `tid`, `exp` and `nbf`. Validate any application-specific claims before use. `SPFxApiPermissionPackageSolutionEntry` contains only resource and scope; listing that pair does not approve tenant API access.

[View source](../../../packages/spfx-react-toolkit/src/helpers/spfx-api-permission-precheck.helpers.ts)
