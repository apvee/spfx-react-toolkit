# useSPFxTenantProperty

[Storage overview](../../INDEX.md#storage-hooks) · [Usage guide](../../../guides/storage.md)

```ts
import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).

Access tenant-wide custom properties (read-only).

This hook and `createSPFxTenantPropertyService` expose reads only; neither exports write/remove operations. Manage StorageEntity provisioning through your administrator’s supported tooling. For toolkit-managed read/write tenant configuration, use [`useSPFxTenantKeyValueStore`](tenant-key-value-store.md).

## Signature

```typescript
function useSPFxTenantProperty<T = unknown>(
  key: string,
  autoFetch?: boolean
): SPFxTenantPropertyResult<T>
```

## Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `key` | `string` | Yes | — | Tenant property key |
| `autoFetch` | `boolean` | No | `true` | Auto-load on mount |

## Returns

```typescript
interface SPFxTenantPropertyResult<T> {
  /** Current value; undefined initially, after identity reset or for an absent value */
  readonly data: T | undefined;

  /** Property description metadata */
  readonly description: string | undefined;

  /** Loading state */
  readonly isLoading: boolean;

  /** Error if fetch failed */
  readonly error: Error | undefined;

  /** Manually load/reload the property */
  readonly load: () => Promise<void>;

  /** True when data is defined, with no pending load or current read error */
  readonly isReady: boolean;
}
```

## Request behavior

Changing key, service or catalog-discovery identity invalidates previous request ownership and clears old data/description/error. `autoFetch` defaults to true: enabling it starts a load when prerequisites exist; disabling it alone retains current state and does not cancel an already dispatched load. Only the latest load for the active identity can publish data, description, error and loading state. The previous value/description remain during a same-identity reload, and can remain after a failed reload; render `error` before treating retained data as current. `load()` records read failures in `error` and resolves rather than rethrowing them. If no client or key is available it logs and resolves without dispatch. `isReady` is exactly `!isLoading && !error && data !== undefined`, so an absent value is not ready. Requests are not cancelled on identity change or unmount. The read uses the StorageEntity JSON-first decoder; `T` does not validate its shape. There is no write or remove operation.

## Example: Feature Flags

```tsx
import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';

function FeatureGatedComponent() {
  const { data: featureFlags, isLoading } = useSPFxTenantProperty<{ enableNewUI?: boolean; enableBetaFeatures?: boolean }>('FeatureFlags');

  if (isLoading) return <Spinner />;

  const flags = featureFlags ?? {};

  return (
    <div>
      {flags.enableNewUI && <NewUIComponent />}
      {flags.enableBetaFeatures && <BetaFeatures />}
    </div>
  );
}
```

## Example: Tenant Configuration

```tsx
import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';

function TenantConfiguredComponent() {
  const { data: apiEndpoint, isLoading, error } = useSPFxTenantProperty<string>('CustomApiEndpoint');
  const { data: apiKey } = useSPFxTenantProperty<string>('CustomApiKey');

  if (isLoading) return <Spinner label="Loading configuration..." />;
  if (error) return <MessageBar messageBarType={MessageBarType.error}>{error.message}</MessageBar>;

  if (!apiEndpoint || !apiKey) {
    return (
      <MessageBar messageBarType={MessageBarType.warning}>
        Tenant properties not configured. Contact your administrator.
      </MessageBar>
    );
  }

  return <ApiClient endpoint={apiEndpoint} apiKey={apiKey} />;
}
```

## Example: Multi-Property Configuration

```tsx
import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit';

function ConfiguredWidget() {
  const logo = useSPFxTenantProperty<string>('CompanyLogo');
  const theme = useSPFxTenantProperty<{ primary: string }>('CompanyTheme');
  const helpUrl = useSPFxTenantProperty<string>('HelpDeskUrl');

  const isLoading = logo.isLoading || theme.isLoading || helpUrl.isLoading;

  if (isLoading) return <Spinner />;

  const themeColors = theme.data ?? { primary: '#0078d4' };

  return (
    <div style={{ '--primary-color': themeColors.primary } as React.CSSProperties}>
      {logo.data && <img src={logo.data} alt="Company Logo" />}
      {helpUrl.data && <a href={helpUrl.data}>Help</a>}
    </div>
  );
}
```

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxTenantProperty.ts)
