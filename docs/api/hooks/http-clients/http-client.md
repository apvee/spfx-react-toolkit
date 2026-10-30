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
  const [weather, setWeather] = React.useState<any>(null);

  const loadWeather = () => {
    invoke(client =>
      client.get(
        'https://api.openweathermap.org/data/2.5/weather?q=London&appid=YOUR_KEY',
        HttpClient.configurations.v1
      ).then(res => res.json())
    ).then(data => setWeather(data));
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
      )
    ).then(() => console.log('Notification sent'));
  };

  return (
    <button onClick={() => sendNotification('Hello!')} disabled={!isReady || isLoading}>
      Send to Slack
    </button>
  );
}
```

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxHttpClient.ts)

---
