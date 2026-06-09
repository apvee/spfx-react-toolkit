# SPFx API Permission Precheck Design

## Goal

Add a lightweight SPFx runtime precheck that helps developers and administrators understand whether Microsoft Graph and custom API delegated permission scopes appear available to the SharePoint Framework runtime.

The feature is diagnostic and configuration-oriented. It must help a web part or extension display actionable messages such as "Approve Microsoft Graph / Sites.Read.All in SharePoint Admin Center API access" without requiring directory-wide Graph admin permissions.

## Non-Goals

- Do not read tenant permission grants from Microsoft Graph, SharePoint Admin PowerShell, or Entra admin APIs.
- Do not claim that a permission is formally approved in the tenant.
- Do not replace server-side authorization for custom APIs.
- Do not support application permissions or app roles in the first version.
- Do not introduce direct MSAL usage.
- Do not introduce a new test framework or package dependency.

## SPFx-Specific Semantics

The check must be framed around SPFx, not generic OAuth:

- The hook consumes SPFx `ServiceScope`.
- The token provider comes from `AadTokenProviderFactory`.
- Token acquisition uses `AadTokenProvider.getToken(resourceEndpoint, options)`.
- The implicit principal is the SharePoint Online Client Extensibility Web Application Principal for non-domain-isolated SPFx solutions.
- Microsoft Graph and custom API permissions map back to `config/package-solution.json` `solution.webApiPermissionRequests`.
- Messages must direct users to SharePoint Admin Center API access approval.

The result must not use words such as `approved` or `notApproved`. The correct meaning is:

> The current SPFx runtime can or cannot obtain an access token for this resource whose delegated `scp` claim contains the requested scope.

## Recommended Public API

### Hook

```ts
function useSPFxApiPermissionPrecheck(
  config: SPFxApiPermissionPrecheckConfig,
  options?: SPFxApiPermissionPrecheckOptions
): SPFxApiPermissionPrecheckResult;
```

Basic usage:

```tsx
const precheck = useSPFxApiPermissionPrecheck({
  graph: ['Sites.Read.All', 'User.Read'],
  customApis: [
    {
      id: 'orders-api',
      name: 'Orders API',
      resource: 'api://contoso-orders-api',
      packageResource: 'Orders API',
      scopes: ['Orders.Read']
    }
  ]
});
```

Manual, passive diagnostic usage:

```tsx
const precheck = useSPFxApiPermissionPrecheck(
  {
    graph: ['Sites.Read.All'],
    customApis: [
      {
        name: 'Orders API',
        resource: 'api://contoso-orders-api',
        scopes: ['Orders.Read']
      }
    ]
  },
  {
    autoCheck: false,
    mode: 'passive'
  }
);

return (
  <PrimaryButton
    text={precheck.isChecking ? 'Checking...' : 'Check API permissions'}
    disabled={precheck.isChecking}
    onClick={precheck.check}
  />
);
```

### Config Types

```ts
interface SPFxApiPermissionPrecheckConfig {
  readonly graph?: readonly SPFxPermissionScopeInput[];
  readonly customApis?: readonly SPFxCustomApiPermissionInput[];
  readonly requirements?: readonly SPFxApiPermissionRequirement[];
}

type SPFxPermissionScopeInput =
  | string
  | {
      readonly scope: string;
      readonly required?: boolean;
      readonly satisfies?: readonly string[];
      readonly label?: string;
    };

interface SPFxCustomApiPermissionInput {
  readonly id?: string;
  readonly name: string;
  readonly resource: string;
  readonly packageResource?: string;
  readonly scopes: readonly SPFxPermissionScopeInput[];
  readonly expectedAudiences?: readonly string[];
}

interface SPFxApiPermissionRequirement {
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
}
```

### Options

```ts
interface SPFxApiPermissionPrecheckOptions {
  readonly autoCheck?: boolean;
  readonly mode?: 'passive' | 'interactiveAllowed';
  readonly useCachedToken?: boolean;
  readonly validateAudience?: boolean;
  readonly timeoutMs?: number;
}
```

Defaults:

```ts
{
  autoCheck: true,
  mode: 'passive',
  useCachedToken: true,
  validateAudience: true,
  timeoutMs: 15000
}
```

### Result Types

```ts
interface SPFxApiPermissionPrecheckResult {
  readonly isChecking: boolean;
  readonly isConfigured: boolean;
  readonly configurationState:
    | 'idle'
    | 'checking'
    | 'ready'
    | 'actionRequired'
    | 'cannotDetermine';
  readonly available: readonly SPFxApiPermissionSummary[];
  readonly missing: readonly SPFxApiPermissionSummary[];
  readonly warnings: readonly SPFxApiPermissionSummary[];
  readonly unknown: readonly SPFxApiPermissionSummary[];
  readonly results: readonly SPFxApiPermissionCheckResult[];
  readonly check: () => Promise<readonly SPFxApiPermissionCheckResult[]>;
  readonly retry: () => Promise<readonly SPFxApiPermissionCheckResult[]>;
  readonly retryWithoutCache: () => Promise<readonly SPFxApiPermissionCheckResult[]>;
}
```

```ts
type SPFxApiPermissionCheckStatus =
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
```

`isConfigured` is true only when every required requirement is `available`. Optional requirements can produce warnings without making `isConfigured` false.

## Architecture

Implement the feature in layers so each part is testable and the hook stays small.

### Helper Layer

Create focused pure helpers before I/O services:

- Expand simple Graph/custom API config into normalized requirements.
- Validate requirements before token acquisition.
- Decode JWT payloads defensively.
- Extract delegated scopes from `scp`.
- Validate expected audience values.
- Classify AAD/SPFx errors into diagnostic statuses.
- Build SharePoint Admin Center remediation messages.

Helpers must not import React or SPFx runtime services.

### Service Layer

Create a non-React service:

```ts
function createSPFxApiPermissionPrecheckService(
  tokenProvider: SPFxAadTokenProviderLike
): SPFxApiPermissionPrecheckService;
```

The token provider shape must be minimal and mockable:

```ts
interface SPFxAadTokenProviderLike {
  getToken(
    resourceEndpoint: string,
    options?: { useCachedToken?: boolean; claims?: string }
  ): Promise<string>;
}
```

The service groups requirements by `resourceEndpoint`, calls `getToken` once per resource, then evaluates all scopes for that token. It returns detailed per-requirement results and never exposes the raw token.

### Token Provider Hook

Create `useSPFxAadTokenProvider` as a small SPFx hook that consumes `AadTokenProviderFactory.serviceKey` from `ServiceScope` and initializes `AadTokenProvider` asynchronously.

The hook mirrors existing async client hooks:

- `tokenProvider`
- `isInitializing`
- `initError`
- `isReady`

### Precheck Hook

Create `useSPFxApiPermissionPrecheck` as the React facade:

- Gets `AadTokenProvider` from `useSPFxAadTokenProvider`.
- Creates the service with `useMemo`.
- Runs checks on demand, or automatically when `autoCheck` is true and token provider is ready.
- Exposes simplified buckets for UI: `available`, `missing`, `warnings`, `unknown`.
- Exposes the detailed `results` for diagnostics.

## Edge Cases And Required Behavior

### Configuration

| Case | Behavior |
| --- | --- |
| Empty config | `configurationState: 'idle'`; no token calls |
| Empty Graph scope | `invalidRequirement` |
| Empty custom API resource | `invalidRequirement` |
| Empty custom API name | `invalidRequirement` |
| Duplicate resource/scope | Deduplicate before token calls and return stable results |
| Same resource with different names | Group by resource, add warning metadata |
| `packageResource` omitted | Use `name` for custom APIs and `Microsoft Graph` for Graph |
| Optional scope missing | Warning bucket, does not fail `isConfigured` |
| Requirement kind not `delegatedScope` | `unsupportedPermissionKind` |

### SPFx Runtime

| Case | Behavior |
| --- | --- |
| Hook used outside provider | `tokenProviderUnavailable` |
| `ServiceScope` unavailable | `tokenProviderUnavailable` |
| `AadTokenProviderFactory` consume fails | `tokenProviderUnavailable` |
| `AadTokenProviderFactory.getTokenProvider()` fails | `tokenProviderUnavailable` or classified error |
| Local workbench cannot authenticate | `interactionRequired`, `tokenProviderUnavailable`, or classified error |
| Teams webview blocks auth | Classified auth failure, not `missingScope` |
| Domain-isolated SPFx | Documented as not the primary target; result still reflects the token provider used by that runtime |

### Token Acquisition

| Case | Behavior |
| --- | --- |
| Consent missing (`AADSTS65001`) | `consentRequired` |
| MFA/login/interaction needed | `interactionRequired` |
| Claims challenge | `claimsChallenge` |
| Conditional Access block | `conditionalAccessBlocked` |
| User assignment required | `accessRestricted` |
| Resource not found | `resourceMisconfigured` |
| Network or temporary AAD failure | `transientFailure` |
| Token call exceeds timeout | `timeout` |
| Cached token may be stale | Provide `retryWithoutCache()` |

### Passive Mode

The default `mode: 'passive'` must avoid surprising authentication UI.

When SPFx exposes popup or redirect events, the hook/service integration should cancel or classify interactive requirements where practical. If a fully passive guarantee is not possible in a specific SPFx runtime, documentation must state that consumers should prefer `autoCheck: false` and run the check from an explicit admin action.

In passive mode:

- Do not intentionally launch popup.
- Do not intentionally continue redirect.
- Classify interaction as `interactionRequired` or `claimsChallenge`.
- Clean up temporary event handlers.
- Avoid concurrent token acquisition fan-out.

`mode: 'interactiveAllowed'` permits standard SPFx token acquisition behavior and is opt-in.

### Token Parsing

| Case | Behavior |
| --- | --- |
| Token is not JWT-shaped | `unsupportedToken` |
| Payload cannot be base64url-decoded | `unsupportedToken` |
| Payload JSON is invalid | `unsupportedToken` |
| Token expired or not yet valid | `tokenExpired` or `unknown` depending on claims |
| `aud` mismatches expected audience | `audienceMismatch` |
| `scp` contains requested scope | `available` |
| `scp` is valid but requested scope absent | `missingScope` |
| `roles` exists without `scp` | `unsupportedPermissionKind` with app-role warning |
| Scope differs only by case | `missingScope` with case mismatch hint |

### Scope Evaluation

Do exact scope matching by default. Do not infer broader scopes automatically.

`satisfies` is the only supported equivalence mechanism:

```ts
{
  scope: 'Files.Read.All',
  satisfies: ['Files.ReadWrite.All']
}
```

Graph `Sites.Selected`, resource-specific consent, application permissions, and custom API app roles are not first-version supported. The result must explain that this precheck validates delegated scopes in `scp` only.

### Privacy And Security

- Never expose or log the raw access token.
- Never expose the full token payload.
- Expose only diagnostic fields such as `detectedScopes`, `audience`, `tenantId`, `errorCode`, and generated messages.
- Document that custom APIs must validate tokens server-side.
- Document that `available` does not guarantee the user can access every downstream resource.

## User-Facing Remediation

Each missing required permission should provide:

```ts
{
  resource: 'Microsoft Graph',
  scope: 'Sites.Read.All',
  packageSolutionEntry: {
    resource: 'Microsoft Graph',
    scope: 'Sites.Read.All'
  },
  adminMessage: 'Approve Microsoft Graph / Sites.Read.All in SharePoint Admin Center API access.'
}
```

For custom APIs:

```ts
{
  resource: 'api://contoso-orders-api',
  scope: 'Orders.Read',
  packageSolutionEntry: {
    resource: 'Orders API',
    scope: 'Orders.Read'
  },
  adminMessage: 'Approve Orders API / Orders.Read in SharePoint Admin Center API access.'
}
```

## Files Expected In The Implementation Plan

The implementation plan should create or modify these areas in this order:

1. Pure helpers:
   - `src/helpers/spfx-api-permission-precheck.helpers.ts`
   - `src/helpers/index.ts`
2. Non-React service:
   - `src/services/spfx-api-permission-precheck.service.ts`
   - `src/services/index.ts`
3. Token provider hook:
   - `src/hooks/useSPFxAadTokenProvider.ts`
   - `src/hooks/index.ts`
4. Precheck hook:
   - `src/hooks/useSPFxApiPermissionPrecheck.ts`
   - `src/hooks/index.ts`
5. Verification scripts:
   - Add focused Node verification for pure helper/service behavior if feasible without a browser.
   - Update package scripts only if a new script is added.
6. Public docs:
   - `README.md`
   - `docs/INTRODUCTION.md`
   - `docs/INDEX.md`
   - `docs/api/helpers/INDEX.md`
   - `docs/api/services/INDEX.md`
   - `docs/api/hooks/INDEX.md`
   - `docs/api/hooks/http-clients.md` or `docs/api/hooks/permissions.md`, whichever reads better after implementation.
7. Test webpart surface:
   - `src/webparts/spFxReactToolkitTest/components/demoRegistry.ts`
   - Existing clients panel or a focused API permissions panel.

## Repository Constraints

- New services exported from `src/services/index.ts` must be documented in `docs/api/services/INDEX.md`, or `npm run verify:public-docs` will fail.
- New helper exports must be documented in `docs/api/helpers/INDEX.md`.
- New hooks exported from `src/hooks/index.ts` must be represented in the demo registry, or `npm run verify:examples` will fail.
- Keep package dependencies unchanged.
- Preserve existing public APIs.
- Do not refactor unrelated hooks or panels.

## Verification Strategy

The implementation plan should include checks for:

- Config expansion for Graph and custom API shorthand.
- Requirement validation.
- JWT base64url payload decoding.
- `scp` exact matching.
- `satisfies` matching.
- Audience mismatch.
- Missing scope.
- Unsupported token.
- App-role-only token warning.
- Error classification for common AAD strings such as `AADSTS65001`.
- Grouping by resource so multiple scopes produce one token request.
- Hook build/typecheck through existing SPFx build.
- Public docs verification.
- Demo coverage verification.

Expected commands:

```bash
npm run verify:public-docs
npm run verify:examples
npm run build
npm test
```

## Open Decision Locked For V1

Version 1 supports delegated scopes from `scp` only. Tokens that expose only `roles` are detected and classified as unsupported for this precheck. This keeps the first implementation focused on SPFx `webApiPermissionRequests` delegated permission configuration.

