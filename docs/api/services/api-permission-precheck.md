# API Permission Precheck Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

Import this factory and its public types from `@apvee/spfx-react-toolkit/services` or the package root. Callers supply the initialized client/context; these services require no React provider.

## `createSPFxApiPermissionPrecheckService`

```ts
function createSPFxApiPermissionPrecheckService(
  tokenProvider: SPFxAadTokenProviderLike
): SPFxApiPermissionPrecheckService;
```

Required public types:

- `SPFxAadTokenProviderLike`
- `SPFxApiPermissionPrecheckServiceOptions`
- `SPFxApiPermissionPrecheckService`

`SPFxAadTokenProviderLike` is the minimal token provider shape used by the service:

```ts
interface SPFxAadTokenProviderLike {
  getToken(
    resourceEndpoint: string,
    options?: { readonly useCachedToken?: boolean; readonly claims?: string }
  ): Promise<string>;
}
```

`SPFxApiPermissionPrecheckServiceOptions` supports:

| Option | Purpose |
|--------|---------|
| `useCachedToken` | Uses SPFx cached tokens unless set to `false`. |
| `validateAudience` | Validates token audience against expected audiences unless set to `false`. |
| `timeoutMs` | Token acquisition timeout; defaults to `15000`. |
| `sequentialResourceAcquisition` | Acquires resource-group tokens one at a time when set to `true`; defaults to concurrent acquisition. |

Returned `SPFxApiPermissionPrecheckService` method:

| Method | Description |
|--------|-------------|
| `check(config, options?)` | Normalizes Graph/custom API requirements, obtains one token per resource endpoint, and returns per-scope check results. |

The service groups requirements by `resourceEndpoint`, so multiple scopes for Microsoft Graph or the same custom API share a single token acquisition attempt. It decodes only the JWT payload needed to evaluate delegated `scp` scopes; raw tokens are never exposed in the returned results.

Each result includes a `packageSolutionEntry` with the resource and scope to add to `webApiPermissionRequests`, making remediation copyable into `package-solution.json`.

## Usage, results and failure behavior

```ts
const diagnostics = createSPFxApiPermissionPrecheckService(tokenProvider);
const results = await diagnostics.check({
  graph: ['User.Read', { scope: 'Mail.Read', required: false }],
}, { useCachedToken: false, timeoutMs: 10000 });
for (const result of results) console.log(result.scope, result.status, result.packageSolutionEntry);
```

`SPFxApiPermissionPrecheckService.check(config, options?)` returns `Promise<readonly SPFxApiPermissionCheckResult[]>`. An empty configuration returns an empty array without acquisition. See the [helper reference](../helpers/api-permission-precheck.md) for every configuration/result type, normalization rules, diagnostics and summary buckets. Invalid normalized requirements become results without a token request. Valid requirements are grouped by the exact resource endpoint; invalid results appear first, then groups in first-seen order with each group's scopes in order. Concurrent acquisition does not reorder output by completion time.

Construction makes no request. A caller-owned `SPFxAadTokenProviderLike` must supply `getToken(resourceEndpoint, options?)`; the service passes its resolved `useCachedToken` preference. Although the provider interface permits optional claims, this service does not request challenge claims or retry. Timeout defaults to 15000ms and maps to per-scope timeout diagnostics; it does not cancel the provider's underlying request. Audience validation defaults to enabled but requires expected audiences; the helpers do not verify signatures, issuer or identity.

Acquisition, timeout and token-evaluation failures inside a resource group are converted to diagnostic results; they are not thrown as normal permission failures. Malformed untyped configuration that fails normalization can still reject. The service does not register passive popup/redirect cancellation or choose an interaction mode: the caller's token provider controls interaction. The [React hook](../hooks/http-clients/api-permission-precheck.md) additionally handles passive/interactive policy and provider lifecycle. A reported available scope is diagnostic and never authorizes the target API operation.
