# useSPFxApiPermissionPrecheck

[HTTP client overview](../../INDEX.md#http-client-hooks) · [Usage guide](../../../guides/http-clients.md)

```ts
import { useSPFxApiPermissionPrecheck } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).


Check whether the current SPFx runtime can obtain delegated tokens for Microsoft Graph and custom API scopes.

## Signature

```typescript
function useSPFxApiPermissionPrecheck(
  config: SPFxApiPermissionPrecheckConfig,
  options?: SPFxApiPermissionPrecheckOptions
): SPFxApiPermissionPrecheckResult
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `config` | `SPFxApiPermissionPrecheckConfig` | Yes | Graph, custom API, or explicit delegated-scope requirements |
| `options` | `SPFxApiPermissionPrecheckOptions` | No | Check behavior such as passive mode, cache usage, audience validation, and timeout |

## Options

| Option | Default | Description |
|--------|---------|-------------|
| `autoCheck` | `true` | Runs a check after the token provider is ready. Set to `false` for manual checks. |
| `mode` | `'passive'` | In passive mode, the hook blocks SPFx popup and redirect auth events during the check. Use `'interactiveAllowed'` only when an interactive token flow is acceptable. |
| `useCachedToken` | `true` | Uses cached tokens unless set to `false`. |
| `validateAudience` | `true` | Validates token audience against expected audiences unless set to `false`. |
| `timeoutMs` | `15000` | Token acquisition timeout in milliseconds. |

## Returns

```typescript
interface SPFxApiPermissionPrecheckResult {
  readonly isChecking: boolean;
  readonly isConfigured: boolean;
  readonly configurationState: 'idle' | 'checking' | 'ready' | 'actionRequired' | 'cannotDetermine';
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
```

## Description

The hook normalizes Graph and custom API requirements, uses the SPFx `AadTokenProvider` to acquire one token per resource, and evaluates delegated `scp` scopes in the token payload. Remediation entries include the `resource` and `scope` values that map back to `webApiPermissionRequests`.

`available` means the current SPFx runtime obtained a token with the delegated scope. It does not read tenant grants and does not replace server-side authorization. Requirements or token-provider replacement clear prior check results and invalidate older requests. With `autoCheck` enabled, the current provider is checked again when ready; manual mode waits for `check()`/`retry()`.

## Example: Simple Precheck

```tsx
import { useSPFxApiPermissionPrecheck } from '@apvee/spfx-react-toolkit';

function OrdersPermissionGate() {
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

  if (precheck.isChecking) return <Spinner label="Checking API access..." />;
  if (precheck.configurationState === 'ready') return <OrdersView />;

  return (
    <MessageBar>
      {precheck.missing.map(item => (
        <div key={item.id}>{item.adminMessage}</div>
      ))}
    </MessageBar>
  );
}
```

## Example: Manual Passive Check

```tsx
import { useSPFxApiPermissionPrecheck } from '@apvee/spfx-react-toolkit';

function ManualPermissionCheck() {
  const precheck = useSPFxApiPermissionPrecheck(
    { graph: ['Sites.Read.All'] },
    { autoCheck: false, mode: 'passive' }
  );

  return (
    <>
      <button onClick={() => precheck.check()} disabled={precheck.isChecking}>
        Check now
      </button>
      <button onClick={() => precheck.retryWithoutCache()} disabled={precheck.isChecking}>
        Recheck without cache
      </button>
      <pre>{precheck.configurationState}</pre>
    </>
  );
}
```

Use `retryWithoutCache()` after SharePoint admin center API access changes or when a cached token may not include a newly granted delegated scope.

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxApiPermissionPrecheck.ts)

---
