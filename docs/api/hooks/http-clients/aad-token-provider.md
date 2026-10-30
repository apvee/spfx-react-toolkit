# useSPFxAadTokenProvider

[HTTP client overview](../../INDEX.md#http-client-hooks) · [Usage guide](../../../guides/http-clients.md)

```ts
import { useSPFxAadTokenProvider } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).

Access the SPFx `AadTokenProvider` directly when you need token acquisition behavior that is not covered by `AadHttpClient`.

## Signature

```typescript
function useSPFxAadTokenProvider(): SPFxAadTokenProviderInfo
```

## Returns

```typescript
interface SPFxAadTokenProviderInfo {
  readonly tokenProvider: AadTokenProvider | undefined;
  readonly isInitializing: boolean;
  readonly initError: Error | undefined;
  readonly isReady: boolean;
}
```

## Description

The hook consumes `AadTokenProviderFactory` from the SPFx service scope and initializes the native provider asynchronously. It does not request a token by itself.

## Example: Manual Token Request

```tsx
import {
  decodeSPFxJwtPayload,
  useSPFxAadTokenProvider,
} from '@apvee/spfx-react-toolkit';

function TokenProbe() {
  const { tokenProvider, isReady, isInitializing, initError } = useSPFxAadTokenProvider();
  const [scopeText, setScopeText] = React.useState<string>('');

  const inspectGraphToken = async () => {
    if (!tokenProvider) {
      return;
    }

    const token = await tokenProvider.getToken('https://graph.microsoft.com');
    const payload = decodeSPFxJwtPayload(token);
    setScopeText(payload?.scp || '');
  };

  if (isInitializing) return <Spinner label="Preparing token provider..." />;
  if (initError) return <ErrorMessage message={initError.message} />;

  return (
    <>
      <button onClick={inspectGraphToken} disabled={!isReady}>
        Inspect Graph scopes
      </button>
      <pre>{scopeText}</pre>
    </>
  );
}
```

`tokenProvider` starts undefined and `isInitializing` starts true. Factory consumption and asynchronous provider acquisition failures populate `initError` and end initialization; `isReady` requires a provider with neither initialization nor error. Factory replacement resets acquisition and obsolete completions cannot publish. This hook does not request a token, check scopes or expose invocation state; direct `getToken` promises follow the native provider contract.

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxAadTokenProvider.ts)
