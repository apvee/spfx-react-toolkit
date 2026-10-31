# useSPFxAadHttpClient

[HTTP client overview](../../INDEX.md#http-client-hooks) · [Usage guide](../../../guides/http-clients.md)

```ts
import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).

Access Azure AD secured API client.

## Signature

```typescript
function useSPFxAadHttpClient(initialResourceUrl?: string): SPFxAadHttpClientInfo
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `initialResourceUrl` | `string` | No | Initial resource URI or application ID; omitted/empty leaves the client unavailable |

## Returns

```typescript
interface SPFxAadHttpClientInfo {
  /** Native AadHttpClient (undefined until initialization completes) */
  readonly client: AadHttpClient | undefined;

  /** Invoke API call with state management */
  readonly invoke: <T>(fn: (client: AadHttpClient) => Promise<T>) => Promise<T>;

  /** Loading state during invoke() calls */
  readonly isLoading: boolean;

  /** Last error from invoke() calls */
  readonly error: Error | undefined;

  /** Clear the current error */
  readonly clearError: () => void;

  /** Set or change the resource URL (triggers client re-initialization) */
  readonly setResourceUrl: (url: string) => void;

  /** Current Azure AD resource URL or App ID */
  readonly resourceUrl: string | undefined;

  /** True while the AAD client is being initialized */
  readonly isInitializing: boolean;

  /** Error that occurred during client initialization */
  readonly initError: Error | undefined;

  /** True when client is ready for use */
  readonly isReady: boolean;
}
```

## Description

Provides `AadHttpClient` for calling custom Azure AD secured APIs. Requires API permissions to be configured in SharePoint admin center.

`initialResourceUrl` is captured once; use `setResourceUrl` for later changes. The resource selects token acquisition; callers still construct the API request URL.

**Async Initialization**: Unlike sync hooks, the AAD client requires async initialization via `factory.getClient()`. Use `isInitializing`, `initError`, and `isReady` to track initialization state.

## Prerequisites

1. Register an Azure AD app for your API
2. Request API permissions in `package-solution.json`:
```json
{
  "webApiPermissionRequests": [
    {
      "resource": "contoso-api",
      "scope": "user_impersonation"
    }
  ]
}
```

3. Approve permissions in SharePoint admin center

## Example: Custom API with State Handling

```tsx
import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
import { AadHttpClient } from '@microsoft/sp-http';

function CustomApiWidget() {
  const {
    invoke,
    isLoading,
    error,
    resourceUrl,
    isInitializing,
    initError,
    isReady
  } = useSPFxAadHttpClient('https://contoso-api.azurewebsites.net');

  const [data, setData] = React.useState<unknown>(undefined);

  const loadData = () => {
    invoke(client =>
      client.get(
        `${resourceUrl}/api/data`,
        AadHttpClient.configurations.v1
      ).then(res => {
        if (!res.ok) throw new Error(`Request failed: ${res.status}`);
        return res.json();
      })
    ).then(setData).catch(() => undefined);
  };

  React.useEffect(() => {
    if (isReady) loadData();
  }, [isReady]);

  // Handle initialization states
  if (isInitializing) return <Spinner label="Initializing AAD client..." />;
  if (initError) return <ErrorMessage message={`Init failed: ${initError.message}`} />;
  if (!isReady) return <Spinner label="Waiting for client..." />;

  // Handle operation states
  if (isLoading) return <Spinner label="Loading data..." />;
  if (error) return <ErrorMessage message={error.message} />;

  return <div>{JSON.stringify(data)}</div>;
}
```

`initialResourceUrl` seeds local resource state only on mount; later argument changes do not select another resource. Use `setResourceUrl(url)` and wait for initialization. With no initial resource, the hook starts without a client and `isInitializing` is false. Resource or factory changes trigger replacement initialization; `initError` describes that acquisition separately from invocation `error`. Native factory consumption can throw during render.

## Example: Dynamic Resource URL

```tsx
import { useSPFxAadHttpClient } from '@apvee/spfx-react-toolkit';
import { AadHttpClient } from '@microsoft/sp-http';

function DynamicApiSelector() {
  const {
    invoke,
    setResourceUrl,
    resourceUrl,
    isInitializing,
    isReady,
    clearError
  } = useSPFxAadHttpClient();

  const [data, setData] = React.useState<unknown>(undefined);

  const loadFromApi = (apiUrl: string) => {
    clearError();
    setResourceUrl(apiUrl);
  };

  // Load data when client becomes ready
  React.useEffect(() => {
    if (isReady && resourceUrl) {
      invoke(client =>
        client.get(
          `${resourceUrl}/api/data`,
          AadHttpClient.configurations.v1
        ).then(res => {
        if (!res.ok) throw new Error(`Request failed: ${res.status}`);
        return res.json();
      })
      ).then(setData).catch(() => undefined);
    }
  }, [isReady, resourceUrl]);

  return (
    <>
      <button onClick={() => loadFromApi('https://api1.contoso.com')} disabled={isInitializing}>
        API 1
      </button>
      <button onClick={() => loadFromApi('https://api2.contoso.com')} disabled={isInitializing}>
        API 2
      </button>
      {isInitializing && <Spinner label="Switching API..." />}
      {data && <pre>{JSON.stringify(data, null, 2)}</pre>}
    </>
  );
}
```

## Invocation state and failures

`invoke(fn)` requires a ready client and returns the callback result. While any invocation on the current client is pending, `isLoading` remains true. Starting another invocation clears `error`; only the latest-started invocation may publish a failure. Client replacement and unmount suppress obsolete UI updates while caller promises still settle. Callback/request rejections are normalized to `Error` and rethrown; catch returned promises. An unavailable client rejects before invocation tracking starts.

`clearError()` clears the invocation error without cancelling work or clearing initialization errors. Direct `client` calls bypass invocation tracking. The hook does not inspect `response.ok`; check unsuccessful HTTP responses and reject them inside the callback. A resolved HTTP error response can otherwise look successful to the hook.

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxAadHttpClient.ts)
