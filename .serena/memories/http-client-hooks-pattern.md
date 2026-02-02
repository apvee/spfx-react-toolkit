# HTTP Client Hooks Pattern

## Overview

The SPFx React Toolkit provides 4 HTTP client hooks with consistent patterns:

| Hook | Init Type | Use Case |
|------|-----------|----------|
| `useSPFxHttpClient` | Sync | External APIs, webhooks |
| `useSPFxSPHttpClient` | Sync | SharePoint REST API |
| `useSPFxAadHttpClient` | Async | Azure AD secured APIs |
| `useSPFxMSGraphClient` | Async | Microsoft Graph API |

## State Management Pattern

### Sync Hooks (HttpClient, SPHttpClient)

```typescript
interface SyncHookResult {
  client: T | undefined;        // Native SPFx client
  invoke: <R>(fn) => Promise<R>; // Execute with state tracking
  isLoading: boolean;           // Loading during invoke()
  error: Error | undefined;     // Last invoke() error
  clearError: () => void;       // Clear error state
  isReady: boolean;             // client !== undefined
}
```

### Async Hooks (MSGraphClient, AadHttpClient)

```typescript
interface AsyncHookResult {
  client: T | undefined;        // Native SPFx client
  invoke: <R>(fn) => Promise<R>; // Execute with state tracking
  isLoading: boolean;           // Loading during invoke()
  error: Error | undefined;     // Last invoke() error
  clearError: () => void;       // Clear error state
  isInitializing: boolean;      // Client being initialized
  initError: Error | undefined; // Initialization error
  isReady: boolean;             // client && !isInitializing && !initError
}
```

## Memory Leak Prevention

All async hooks use:
- `isMountedRef` to track component mounted state
- Check `isMountedRef.current` before any `setState` in async callbacks
- Cleanup effect: `useEffect(() => () => { isMountedRef.current = false; }, [])`

## Usage Pattern

```tsx
function MyComponent() {
  const { invoke, isLoading, isReady, initError } = useSPFxMSGraphClient();
  
  React.useEffect(() => {
    if (isReady) {
      invoke(client => client.api('/me').get()).then(setUser);
    }
  }, [isReady]);
  
  if (initError) return <Error />;
  if (!isReady || isLoading) return <Spinner />;
  return <Content />;
}
```

## Documentation

- Source: `src/hooks/useSPFx*.ts`
- Docs: `docs/api/hooks/http-clients.md`, `docs/api/hooks/storage.md`
