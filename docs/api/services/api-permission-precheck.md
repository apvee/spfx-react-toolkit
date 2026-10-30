# API Permission Precheck Service

[Services overview](../INDEX.md#non-react-services) · [Non-React composition](../../guides/helpers-and-services.md)

```ts
import { createSPFxApiPermissionPrecheckService } from '@apvee/spfx-react-toolkit/services';
```


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
