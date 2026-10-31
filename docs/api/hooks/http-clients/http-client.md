# useSPFxHttpClient

[HTTP client overview](../../INDEX.md#http-client-hooks) · [Usage guide](../../../guides/http-clients.md)

```ts
import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).

Access generic HTTP client for external API calls.

## Signature

```typescript
function useSPFxHttpClient(): SPFxHttpClientInfo
```

## Returns

```typescript
interface SPFxHttpClientInfo {
  /** Native HttpClient from SPFx (undefined if ServiceScope unavailable) */
  readonly client: HttpClient | undefined;

  /** Invoke HTTP call with state management */
  readonly invoke: <T>(fn: (client: HttpClient) => Promise<T>) => Promise<T>;

  /** Loading state during invoke() calls */
  readonly isLoading: boolean;

  /** Last error from invoke() calls */
  readonly error: Error | undefined;

  /** Clear the current error */
  readonly clearError: () => void;

  /** True when client is ready for use */
  readonly isReady: boolean;
}
```

## Description

Provides native `HttpClient` for generic HTTP requests to external APIs, webhooks, or any non-SharePoint endpoints.

**Two usage patterns:**

1. **invoke()** - Automatic state management (loading + error tracking)
2. **client** - Direct access for full control

## Example: External API with isReady

```tsx
import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
import { HttpClient } from '@microsoft/sp-http';

function WeatherWidget() {
  const { invoke, isLoading, error, clearError, isReady } = useSPFxHttpClient();
  const [weather, setWeather] = React.useState<{ main?: { temp?: number } }>();

  const loadWeather = () => {
    invoke(client =>
      client.get(
        'https://api.openweathermap.org/data/2.5/weather?q=London&appid=YOUR_KEY',
        HttpClient.configurations.v1
      ).then(res => {
        if (!res.ok) throw new Error(`Request failed: ${res.status}`);
        return res.json();
      })
    ).then(data => setWeather(data)).catch(() => undefined);
  };

  React.useEffect(() => {
    if (isReady) loadWeather();
  }, [isReady]);

  if (!isReady) return <Spinner label="Initializing..." />;
  if (isLoading) return <Spinner label="Loading weather..." />;
  if (error) return <ErrorMessage message={error.message} onRetry={() => { clearError(); loadWeather(); }} />;

  return <div>Temperature: {weather?.main?.temp}°K</div>;
}
```

## Example: POST to Webhook

```tsx
import { useSPFxHttpClient } from '@apvee/spfx-react-toolkit';
import { HttpClient } from '@microsoft/sp-http';

function SlackNotifier() {
  const { invoke, isLoading, isReady } = useSPFxHttpClient();

  const sendNotification = (message: string) => {
    invoke(client =>
      client.post(
        'https://hooks.slack.com/services/YOUR/WEBHOOK/URL',
        HttpClient.configurations.v1,
        {
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ text: message })
        }
      ).then(response => {
        if (!response.ok) throw new Error(`Request failed: ${response.status}`);
        return response;
      })
    ).then(() => console.log('Notification sent')).catch(console.error);
  };

  return (
    <button onClick={() => sendNotification('Hello!')} disabled={!isReady || isLoading}>
      Send to Slack
    </button>
  );
}
```

## Invocation state and failures

`invoke(fn)` requires a ready client and returns the callback result. While any invocation on the current client is pending, `isLoading` remains true. Starting another invocation clears `error`; only the latest-started invocation may publish a failure. Client replacement and unmount suppress obsolete UI updates while caller promises still settle. Callback/request rejections are normalized to `Error` and rethrown; catch returned promises. An unavailable client rejects before invocation tracking starts.

`clearError()` clears the invocation error without cancelling work or clearing initialization errors. Direct `client` calls bypass invocation tracking. The hook does not inspect `response.ok`; check unsuccessful HTTP responses and reject them inside the callback. A resolved HTTP error response can otherwise look successful to the hook.

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxHttpClient.ts)
