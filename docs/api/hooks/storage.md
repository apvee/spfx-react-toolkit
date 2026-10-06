# Storage Hooks

> Hooks for data persistence across browser storage, OneDrive, tenant properties, and tenant-level key-value store

## Overview

These hooks provide access to browser storage (localStorage, sessionStorage), OneDrive app-specific data storage, tenant properties, and tenant-level key-value store.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxLocalStorage`](#usespfxlocalstorage) | `SPFxStorageHook<T>` | Browser localStorage with namespace |
| [`useSPFxSessionStorage`](#usespfxsessionstorage) | `SPFxStorageHook<T>` | Browser sessionStorage with namespace |
| [`useSPFxOneDriveAppData`](#usespfxonedriveappdata) | `SPFxOneDriveAppDataResult` | OneDrive app-specific storage |
| [`useSPFxTenantProperty`](#usespfxtenantproperty) | `SPFxTenantPropertyResult` | Tenant properties (read-only) |
| [`useSPFxTenantKeyValueStore`](#usespfxtenantkeyvaluestore) | `SPFxTenantKeyValueStoreResult` | Tenant-level key-value store |

---

## useSPFxLocalStorage

Persistent storage that survives browser restarts. Data is namespaced per web part instance.

### Signature

```typescript
function useSPFxLocalStorage<T>(
  key: string,
  defaultValue: T
): SPFxStorageHook<T>
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `key` | `string` | Yes | Storage key (automatically namespaced) |
| `defaultValue` | `T` | Yes | Fallback when the current key is missing or unreadable |

### Returns

An object containing:

- `value`: Current value of type `T`
- `setValue`: Setter accepting a value or updater function
- `remove`: Removes the stored key and resets `value` to the current default

### Default and persistence behavior

The default seeds the current key when no readable JSON exists. Changing only `defaultValue` preserves the current value, including when the default is an inline object. `remove()` and a matching browser storage deletion event reset to the latest default. Changing key or instance ID reloads that scoped key, using the current default if needed. Persistence is best effort: blocked storage, invalid JSON and quota failures do not throw from this hook. Browser storage event rules apply; this is not a same-document synchronization bus.

### Namespacing

Keys are automatically namespaced with the web part instance ID to prevent conflicts between multiple instances of the same web part.

### Example: User Preferences

```tsx
import { useSPFxLocalStorage } from '@apvee/spfx-react-toolkit';

interface UserPreferences {
  theme: 'light' | 'dark';
  itemsPerPage: number;
  showSidebar: boolean;
}

function SettingsPanel() {
  const { value: preferences, setValue: setPreferences } = useSPFxLocalStorage<UserPreferences>(
    'userPrefs',
    { theme: 'light', itemsPerPage: 10, showSidebar: true }
  );
  
  const updateTheme = (theme: 'light' | 'dark') => {
    setPreferences(prev => ({ ...prev, theme }));
  };
  
  const updateItemsPerPage = (itemsPerPage: number) => {
    setPreferences(prev => ({ ...prev, itemsPerPage }));
  };
  
  return (
    <div className="settings">
      <label>
        Theme:
        <select 
          value={preferences.theme} 
          onChange={e => updateTheme(e.target.value as 'light' | 'dark')}
        >
          <option value="light">Light</option>
          <option value="dark">Dark</option>
        </select>
      </label>
      
      <label>
        Items per page:
        <input
          type="number"
          value={preferences.itemsPerPage}
          onChange={e => updateItemsPerPage(Number(e.target.value))}
          min={5}
          max={50}
        />
      </label>
    </div>
  );
}
```

### Example: Search History

```tsx
import { useSPFxLocalStorage } from '@apvee/spfx-react-toolkit';

function SearchWithHistory() {
  const { value: history, setValue: setHistory } = useSPFxLocalStorage<string[]>('searchHistory', []);
  const [query, setQuery] = React.useState('');
  
  const handleSearch = () => {
    if (query && !history.includes(query)) {
      setHistory(prev => [query, ...prev.slice(0, 9)]); // Keep last 10
    }
    // Perform search...
  };
  
  const clearHistory = () => setHistory([]);
  
  return (
    <div>
      <input 
        value={query} 
        onChange={e => setQuery(e.target.value)}
        placeholder="Search..."
        list="search-history"
      />
      <datalist id="search-history">
        {history.map(term => (
          <option key={term} value={term} />
        ))}
      </datalist>
      <button onClick={handleSearch}>Search</button>
      <button onClick={clearHistory}>Clear History</button>
    </div>
  );
}
```

### Source

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxStorage.ts)

---

## useSPFxSessionStorage

Temporary storage that persists only for the current browser session/tab. It has the same default, removal and best-effort behavior as [local storage](#default-and-persistence-behavior).

### Signature

```typescript
function useSPFxSessionStorage<T>(
  key: string,
  defaultValue: T
): SPFxStorageHook<T>
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `key` | `string` | Yes | Storage key (automatically namespaced) |
| `defaultValue` | `T` | Yes | Fallback when the current key is missing or unreadable |

### Returns

An object containing:

- `value`: Current value of type `T`
- `setValue`: Setter accepting a value or updater function
- `remove`: Removes the stored key and resets `value` to the current default

### Example: Form Draft

```tsx
import { useSPFxSessionStorage } from '@apvee/spfx-react-toolkit';

interface FormData {
  title: string;
  description: string;
  category: string;
}

function FormWithDraft() {
  const { value: draft, setValue: setDraft } = useSPFxSessionStorage<FormData>('formDraft', {
    title: '',
    description: '',
    category: ''
  });
  
  const handleChange = (field: keyof FormData, value: string) => {
    setDraft(prev => ({ ...prev, [field]: value }));
  };
  
  const handleSubmit = async () => {
    // Submit form...
    setDraft({ title: '', description: '', category: '' }); // Clear draft
  };
  
  return (
    <form onSubmit={e => { e.preventDefault(); handleSubmit(); }}>
      <input
        value={draft.title}
        onChange={e => handleChange('title', e.target.value)}
        placeholder="Title"
      />
      <textarea
        value={draft.description}
        onChange={e => handleChange('description', e.target.value)}
        placeholder="Description"
      />
      <button type="submit">Submit</button>
      <p className="hint">Draft auto-saved for this session</p>
    </form>
  );
}
```

### Example: Wizard State

```tsx
import { useSPFxSessionStorage } from '@apvee/spfx-react-toolkit';

interface WizardState {
  currentStep: number;
  completedSteps: number[];
  data: Record<string, unknown>;
}

function MultiStepWizard() {
  const { value: wizard, setValue: setWizard } = useSPFxSessionStorage<WizardState>('wizard', {
    currentStep: 0,
    completedSteps: [],
    data: {}
  });
  
  const goToStep = (step: number) => {
    setWizard(prev => ({ ...prev, currentStep: step }));
  };
  
  const completeStep = (stepData: Record<string, unknown>) => {
    setWizard(prev => ({
      currentStep: prev.currentStep + 1,
      completedSteps: [...prev.completedSteps, prev.currentStep],
      data: { ...prev.data, ...stepData }
    }));
  };
  
  return (
    <div className="wizard">
      <nav>
        {[0, 1, 2, 3].map(step => (
          <button
            key={step}
            onClick={() => goToStep(step)}
            disabled={!wizard.completedSteps.includes(step - 1) && step !== 0}
            className={wizard.currentStep === step ? 'active' : ''}
          >
            Step {step + 1}
          </button>
        ))}
      </nav>
      <main>
        {/* Render current step */}
      </main>
    </div>
  );
}
```

### Source

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxStorage.ts)

---

## useSPFxOneDriveAppData

Cloud-synced JSON storage using OneDrive app-specific data folder.

### Prerequisites

Requires Microsoft Graph permissions:
- `Files.ReadWrite` or `Files.ReadWrite.AppFolder`

### Signature

```typescript
function useSPFxOneDriveAppData<T = unknown>(
  fileName: string,
  options?: SPFxOneDriveAppDataOptions<T>
): SPFxOneDriveAppDataResult<T>
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `fileName` | `string` | Yes | File name in app data folder (e.g., `'config.json'`) |
| `options` | `SPFxOneDriveAppDataOptions<T>` | No | Configuration options |

### Options

```typescript
interface SPFxOneDriveAppDataOptions<T> {
  /** Optional folder/namespace for file organization */
  folder?: string;
  
  /** Whether to auto-load on mount. Default: true */
  autoFetch?: boolean;
  
  /** Default value when file is missing (404) */
  defaultValue?: T;
  
  /** If true, create file with defaultValue when missing */
  createIfMissing?: boolean;
}
```

### Returns

```typescript
interface SPFxOneDriveAppDataResult<T> {
  /** Current data value (undefined if not loaded) */
  readonly data: T | undefined;
  
  /** Loading state during load() calls */
  readonly isLoading: boolean;
  
  /** Last error from load() calls */
  readonly error: Error | undefined;
  
  /** True if file does not exist (404 response) */
  readonly isNotFound: boolean;
  
  /** Loading state during write() calls */
  readonly isWriting: boolean;
  
  /** Last error from write() calls */
  readonly writeError: Error | undefined;
  
  /** Load/reload file from OneDrive */
  readonly load: () => Promise<void>;
  
  /** Write data to OneDrive (upsert) */
  readonly write: (content: T) => Promise<void>;
  
  /** True when data is loaded and ready */
  readonly isReady: boolean;
}
```

### Request ownership and lazy defaults

`defaultValue` seeds `data` even with `autoFetch: false`; changing file, folder or Graph client seeds the new identity from the current default without forcing a lazy fetch. Changing an inline default alone does not reload. A current missing-file response is required before `createIfMissing` writes a default.

Latest-started reads and writes own the local data/error state together. An older read, 404 or write completion cannot replace a newer write's local data. `isLoading` and `isWriting` track pending reads and writes separately for the current identity. Identity changes and unmount ignore obsolete completions. Requests still settle for their callers; remote writes are not cancelled or serialized. `load()` records read errors in `error`; `write()` records `writeError` and rejects on failure. These guards do not provide remote transaction ordering.

### Example: Basic Usage with Auto-Fetch

```tsx
import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';

interface MyConfig {
  theme: 'light' | 'dark';
  language: string;
}

function ConfigPanel() {
  const { data, isLoading, error, write, isWriting, isReady } = 
    useSPFxOneDriveAppData<MyConfig>('config.json');
  
  const handleSave = async (newConfig: MyConfig) => {
    try {
      await write(newConfig);
      console.log('Saved!');
    } catch (err) {
      console.error('Save failed:', err);
    }
  };
  
  if (isLoading) return <Spinner label="Loading configuration..." />;
  if (error) return <MessageBar messageBarType={MessageBarType.error}>{error.message}</MessageBar>;
  if (!isReady) return <Spinner />;
  
  return (
    <div>
      <Toggle 
        label="Dark Mode"
        checked={data?.theme === 'dark'}
        onChange={(_, checked) => handleSave({ ...data!, theme: checked ? 'dark' : 'light' })}
        disabled={isWriting}
      />
      {isWriting && <Spinner label="Saving..." />}
    </div>
  );
}
```

### Example: With Folder Namespace

```tsx
import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';

function MyApp() {
  // Files stored in appRoot:/my-app-v2/
  const { data, write, isReady } = useSPFxOneDriveAppData<AppState>(
    'state.json',
    { folder: 'my-app-v2' }
  );
  
  if (!isReady) return <Spinner />;
  return <div>State: {JSON.stringify(data)}</div>;
}
```

### Example: Create If Missing

```tsx
import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';

interface UserPrefs {
  favorites: string[];
  notifications: boolean;
}

const defaultPrefs: UserPrefs = {
  favorites: [],
  notifications: true
};

function UserPreferences() {
  const { 
    data, 
    isLoading, 
    isNotFound, 
    isReady,
    write 
  } = useSPFxOneDriveAppData<UserPrefs>('prefs.json', {
    folder: 'user-settings',
    defaultValue: defaultPrefs,
    createIfMissing: true  // Auto-create file if missing
  });
  
  // If file is missing, it will be auto-created with defaultValue
  if (isLoading) return <Spinner label="Loading preferences..." />;
  if (!isReady) return <Spinner />;
  
  const addFavorite = async (id: string) => {
    await write({
      ...data!,
      favorites: [...data!.favorites, id]
    });
  };
  
  return (
    <div>
      <h3>Favorites ({data?.favorites.length})</h3>
      {/* ... */}
    </div>
  );
}
```

### Example: Manual Load (Lazy)

```tsx
import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';

function LazyLoader() {
  const { data, load, isLoading, isReady } = useSPFxOneDriveAppData<CacheData>(
    'cache.json',
    { autoFetch: false }  // Don't auto-load
  );
  
  return (
    <div>
      <button onClick={load} disabled={isLoading}>
        {isLoading ? 'Loading...' : 'Load Cache'}
      </button>
      {isReady && <pre>{JSON.stringify(data, null, 2)}</pre>}
    </div>
  );
}
```

### Example: CRUD-like Operations

```tsx
import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';

interface TodoList {
  items: Array<{ id: string; text: string; done: boolean }>;
}

function TodoApp() {
  const { data, write, isLoading, isWriting, isReady } = 
    useSPFxOneDriveAppData<TodoList>('todos.json', {
      folder: 'todo-app',
      defaultValue: { items: [] },
      createIfMissing: true
    });
  
  const addTodo = async (text: string) => {
    const newItem = { id: crypto.randomUUID(), text, done: false };
    await write({
      items: [...(data?.items ?? []), newItem]
    });
  };
  
  const toggleTodo = async (id: string) => {
    await write({
      items: data?.items.map(item => 
        item.id === id ? { ...item, done: !item.done } : item
      ) ?? []
    });
  };
  
  const deleteTodo = async (id: string) => {
    await write({
      items: data?.items.filter(item => item.id !== id) ?? []
    });
  };
  
  if (isLoading) return <Spinner />;
  if (!isReady) return <Spinner label="Initializing..." />;
  
  return (
    <div>
      <TodoList 
        items={data?.items ?? []} 
        onToggle={toggleTodo}
        onDelete={deleteTodo}
      />
      <AddTodoForm onAdd={addTodo} disabled={isWriting} />
      {isWriting && <span>Saving...</span>}
    </div>
  );
}
```

### Source

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxOneDriveAppData.ts)

---

## useSPFxTenantProperty

Access tenant-wide custom properties (read-only).

> **Note:** Write and remove operations have been removed in v2.0.0. Microsoft has blocked the
> SetStorageEntity and RemoveStorageEntity REST API endpoints. Tenant properties can only be
> managed via PowerShell (`Set-PnPStorageEntity`, `Remove-PnPStorageEntity`).
> For a read/write key-value store at tenant level, use [`useSPFxTenantKeyValueStore`](#usespfxtenantkeyvaluestore).

### Signature

```typescript
function useSPFxTenantProperty<T = unknown>(
  key: string,
  autoFetch?: boolean
): SPFxTenantPropertyResult<T>
```

### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `key` | `string` | Yes | — | Tenant property key |
| `autoFetch` | `boolean` | No | `true` | Auto-load on mount |

### Returns

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

### Request behavior

Changing key, service/catalog identity or fetch mode invalidates previous request ownership and clears old data. Only the latest load can publish data, error and loading state. `load()` records read failures in `error`; it resolves rather than rethrowing them. `isReady` indicates a successfully loaded current value. Requests are not cancelled on identity change or unmount.

### Example: Feature Flags

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

### Example: Tenant Configuration

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

### Example: Multi-Property Configuration

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

### Source

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxTenantProperty.ts)

---

## useSPFxTenantKeyValueStore

Tenant-level key-value store backed by a hidden SharePoint list in the tenant app catalog. Provides CRUD operations (get, list, save, remove) using the existing string/JSON serialization format.

This hook is an alternative to tenant properties (StorageEntity) for scenarios requiring read/write access via REST, since Microsoft has blocked the SetStorageEntity and RemoveStorageEntity REST endpoints.

### Signature

```typescript
function useSPFxTenantKeyValueStore(): SPFxTenantKeyValueStoreResult
```

### Returns

```typescript
interface SPFxTenantKeyValueStoreResult {
  /** Loading state for read operations */
  readonly isLoading: boolean;
  /** Last error from read operations */
  readonly error: Error | undefined;
  /** Loading state for write operations */
  readonly isWriting: boolean;
  /** Last error from write operations */
  readonly writeError: Error | undefined;
  /** Whether the current user can write (Site Collection Admin on app catalog) */
  readonly canWrite: boolean;
  /** True when SPHttpClient is available */
  readonly isReady: boolean;

  /** Get a single item by key */
  readonly get: <T = unknown>(key: string) => Promise<SPFxTenantKeyValueStoreItem<T> | undefined>;
  /** List all items (sorted by key) */
  readonly list: () => Promise<SPFxTenantKeyValueStoreItem<unknown>[]>;
  /** Create or update a key-value pair (auto-provisions list on first write) */
  readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;
  /** Remove an item by key (no-op if not found) */
  readonly remove: (key: string) => Promise<void>;
}

interface SPFxTenantKeyValueStoreItem<T = unknown> {
  /** Property key */
  readonly key: string;
  /** Deserialized property value */
  readonly value: T;
  /** Optional description metadata */
  readonly description: string | undefined;
  /** SharePoint list item ID */
  readonly id: number;
}
```

### Error and loading behavior

Read operations remain pending until their service promises settle. `get()` returns `undefined` and `list()` returns `[]` on read failures while publishing `error`; these fallbacks alone cannot distinguish absence from failure. Writes publish `writeError` and reject their returned promises on failure. Handle write rejections in the caller.

Read and write channels have independent pending counts: their loading flags remain true while any current-identity operation in that channel is pending. Starting an operation clears its channel error; only the latest-started operation may publish that error. Old catalog/service and unmounted completions do not update the current hook. The server still controls authorization and remote ordering.

### Storage Details

The store uses a hidden list named `TenantKeyValueStore` in the tenant app catalog:
- **Title** column: key (indexed, unique)
- **Value** column: multiline text (stores serialized data)
- **Description** column: multiline text (optional metadata)

The list is auto-provisioned on the first `save()` call. Read operations (`get`, `list`) return `undefined`/`[]` gracefully if the list doesn't exist yet.

### Serialization

| Input type | Stored as |
|------------|-----------|
| `string` | Raw string |
| `number`, `boolean`, `bigint` | `String(value)` |
| `null` | `"null"` |
| `Date` | ISO 8601 string |
| Objects/arrays | `JSON.stringify()` |

Deserialization attempts `JSON.parse()` first; falls back to raw string. This legacy format does not preserve every input type: strings such as `"123"`, `"true"` or `"null"` become JSON primitives; a bigint may become a lossy number; dates remain ISO strings. The generic `T` is a compile-time annotation, not runtime validation. Existing stored values retain this format; validate read values in application code.

### Requirements

- Tenant app catalog must be provisioned
- **Read**: Requires the current user to have access to the catalog/list; actual tenant permissions apply
- **Write/Remove**: Site Collection Administrator role on the tenant app catalog site

### Example: Basic CRUD

```tsx
import { useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit';

function TenantConfigEditor() {
  const store = useSPFxTenantKeyValueStore();
  const [value, setValue] = React.useState('');

  const handleLoad = async () => {
    const item = await store.get<string>('apiEndpoint');
    if (item) setValue(item.value);
  };

  const handleSave = async () => {
    await store.save<string>('apiEndpoint', value, 'Production API endpoint');
  };

  const handleDelete = async () => {
    await store.remove('apiEndpoint');
    setValue('');
  };

  return (
    <Stack tokens={{ childrenGap: 10 }}>
      <TextField value={value} onChange={(_, v) => setValue(v ?? '')} />
      <Stack horizontal tokens={{ childrenGap: 5 }}>
        <PrimaryButton onClick={handleLoad}>Load</PrimaryButton>
        <PrimaryButton onClick={handleSave} disabled={store.isWriting}>Save</PrimaryButton>
        <DefaultButton onClick={handleDelete}>Delete</DefaultButton>
      </Stack>
      {store.error && <MessageBar messageBarType={MessageBarType.error}>{store.error.message}</MessageBar>}
    </Stack>
  );
}
```

### Example: Heterogeneous Types

```tsx
const store = useSPFxTenantKeyValueStore();

// String
await store.save<string>('appVersion', '2.1.0');

// Number
await store.save<number>('maxUploadSize', 10485760);

// Complex object
interface FeatureFlags { enableChat: boolean; maxUsers: number; }
await store.save<FeatureFlags>('featureFlags', { enableChat: true, maxUsers: 500 });

// Read back with type safety
const flags = await store.get<FeatureFlags>('featureFlags');
if (flags?.value.enableChat) { /* ... */ }
```

### Example: Property Dashboard

```tsx
function AllPropertiesView() {
  const store = useSPFxTenantKeyValueStore();
  const [items, setItems] = React.useState<SPFxTenantKeyValueStoreItem<unknown>[]>([]);

  React.useEffect(() => {
    store.list().then(setItems);
  }, []);

  if (store.isLoading) return <Spinner />;

  return (
    <DetailsList
      items={items.map(i => ({
        key: i.key,
        value: typeof i.value === 'object' ? JSON.stringify(i.value) : String(i.value),
        description: i.description ?? '',
      }))}
      columns={[
        { key: 'key', name: 'Key', fieldName: 'key', minWidth: 150 },
        { key: 'value', name: 'Value', fieldName: 'value', minWidth: 200 },
        { key: 'desc', name: 'Description', fieldName: 'description', minWidth: 200 },
      ]}
    />
  );
}
```

### Source

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxTenantKeyValueStore.ts)

---

## See Also

- [HTTP Client Hooks](./http-clients.md) - API access
- [PnPjs Hooks](./pnpjs.md) - SharePoint data access
- [Context Hooks](./context.md) - SPFx context
- [Performance Hooks](./performance.md) - Performance measurement & diagnostics

---

*Generated from JSDoc comments. Last updated: April 10, 2026*
