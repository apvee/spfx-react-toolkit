# useSPFxTenantProperty

[Storage overview](../../INDEX.md#storage-hooks) · [Usage guide](../../../guides/storage.md)

```ts
import { useSPFxTenantProperty } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).


Access tenant-wide custom properties (read-only).

> **Note:** Write and remove operations have been removed in v2.0.0. Microsoft has blocked the SetStorageEntity and RemoveStorageEntity REST API endpoints. Tenant properties can only be managed via PowerShell (`Set-PnPStorageEntity`, `Remove-PnPStorageEntity`). For a read/write key-value store at tenant level, use [`useSPFxTenantKeyValueStore`](tenant-key-value-store.md#usespfxtenantkeyvaluestore).

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
  /** Property value (undefined while loading or if not found) */
  readonly data: T | undefined;

  /** Property description metadata */
  readonly description: string | undefined;

  /** Loading state */
  readonly isLoading: boolean;

  /** Error if fetch failed */
  readonly error: Error | undefined;

  /** Manually load/reload the property */
  readonly load: () => Promise<void>;

  /** True if data is loaded successfully */
  readonly isReady: boolean;
}
```

## Request behavior

Changing key, service/catalog identity or fetch mode invalidates previous request ownership and clears old data. Only the latest load can publish data, error and loading state. `load()` records read failures in `error`; it resolves rather than rethrowing them. `isReady` indicates a successfully loaded current value. Requests are not cancelled on identity change or unmount.

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

---
