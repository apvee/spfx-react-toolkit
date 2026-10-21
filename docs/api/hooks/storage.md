# Storage Hooks

> Hooks for data persistence across browser storage, OneDrive, tenant properties, site collection and tenant-level key-value stores

## Overview

These hooks provide access to browser storage (localStorage, sessionStorage), OneDrive app-specific data storage, tenant properties, site collection and tenant-level key-value stores.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxLocalStorage`](#usespfxlocalstorage) | `SPFxStorageHook<T>` | Browser localStorage with namespace |
| [`useSPFxSessionStorage`](#usespfxsessionstorage) | `SPFxStorageHook<T>` | Browser sessionStorage with namespace |
| [`useSPFxOneDriveAppData`](#usespfxonedriveappdata) | `SPFxOneDriveAppDataResult` | OneDrive app-specific storage |
| [`useSPFxTenantProperty`](#usespfxtenantproperty) | `SPFxTenantPropertyResult` | Tenant properties (read-only) |
| [`useSPFxTenantKeyValueStore`](#usespfxtenantkeyvaluestore) | `SPFxTenantKeyValueStoreResult` | Tenant-level key-value store |
| [`useSPFxSiteKeyValueStore`](#usespfxsitekeyvaluestore) | `SPFxSiteKeyValueStoreResult` | Collection root-web key-value store shared by subsites |

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

### Browser downloads

Reads request the file's complete Graph metadata and immediately download its JSON from `@microsoft.graph.downloadUrl` using native `fetch` with `credentials: 'omit'`. The path lookup uses no `$select` projection, which can omit the download annotation. Graph authorization headers are not forwarded to the download host. This avoids the authenticated `/content` redirect that fails browser CORS preflight, as described in [Microsoft's download guidance](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0#downloading-files-in-javascript-apps).

Download URLs must be absolute HTTPS URLs and are used immediately rather than cached or logged. Only a missing-file response from Graph metadata sets `isNotFound`. Denied Graph requests, failed downloads (including download URL 404s) and invalid JSON remain errors, preventing `createIfMissing` from acting on an unavailable download URL. Writes continue to use the Graph `/content` endpoint. The consumer runtime must provide native `fetch`.

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

Reads and writes shared tenant configuration in the hidden `TenantKeyValueStore` list in the tenant app catalog. Requires a matching host-specific SPFx provider, its authenticated `SPHttpClient`, page context, and a provisioned tenant app catalog. All applications share the same key namespace; use application-specific key prefixes.

### Signature and result

```typescript
function useSPFxTenantKeyValueStore(): SPFxTenantKeyValueStoreResult

interface SPFxTenantKeyValueStoreResult {
  readonly isLoading: boolean;
  readonly error: Error | undefined;
  readonly isWriting: boolean;
  readonly writeError: Error | undefined;
  readonly canWrite: boolean;
  readonly isReady: boolean;
  readonly get: <T = unknown>(key: string) => Promise<SPFxTenantKeyValueStoreItem<T> | undefined>;
  readonly list: () => Promise<SPFxTenantKeyValueStoreItem<unknown>[]>;
  readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;
  readonly remove: (key: string) => Promise<void>;
}

interface SPFxTenantKeyValueStoreItem<T = unknown> {
  readonly key: string;
  readonly value: T;
  readonly description: string | undefined;
  readonly id: number;
}
```

### Initialization, permissions and errors

Mount discovers the catalog and checks permissions using read-only requests. `get`, `list` and `remove` never create the list or repair its schema. Only `save` (or the non-React service's explicit `ensureListReady`) provisions/repairs it.

`isReady` retains its client-availability meaning: it does **not** mean discovery has finished, the list exists or the user can write. `canWrite` starts false and checks effective Add/Edit/Delete Items permissions on the list. If the list is confirmed absent, it checks those permissions plus Manage Lists on the catalog web. It refreshes after operations; failed checks return false and older checks cannot overwrite newer results. Schema repair can require additional permissions even when item-write permissions are present. SharePoint authorizes every request; this flag is only an indicator, not a security boundary or a guarantee of success.

Only a confirmed missing list/key is considered absent. Access denials, server errors, incompatible fields, duplicate lookup results, malformed data and paging failures publish `error`; `get` returns `undefined` and `list` returns `[]` for compatibility. Inspect `error` to distinguish failure from absence. Missing client availability rejects the operation. Writes/removal publish `writeError` and reject; callers must handle rejections. Background discovery failures are retried and surfaced by explicit operations.

Read/write loading flags independently count pending operations. Starting an operation clears its channel error, and only the latest-started operation may publish an error. Obsolete contexts and unmounted completions cannot update hook state. Remote requests already dispatched are not cancelled.

### Storage, paging and concurrency

The list keeps its existing schema: `Title` (Text/key), `Value` (Note/serialized text), `Description` (Note/optional metadata). Setup verifies field types and requires indexed, unique Title keys on both existing and newly created lists. Failed uniqueness/schema checks reject and are retryable; existing data is never deleted to repair schema. Resolve pre-existing duplicate keys administratively before retrying a write. Concurrent setup is shared within each service instance per normalized catalog URL; separate instances recover creation races only after bounded schema rereads.

`list()` follows all supported SharePoint continuation pages in server Title order. Links must stay on the same origin and list-items path, without credentials/fragments or repeated URLs. Failed later pages reject in the service and produce the hook's empty fallback plus `error`; partial data is never presented as a complete list. SharePoint list thresholds and the user's actual access still apply.

`save()` updates existing keys and preserves an omitted description; `description: ''` clears it. Updates and deletion use the observed item's ETag, fetching a single-item snapshot if collection metadata omits it. A missing ETag fails safely; HTTP 412 conflicts reject without wildcard overwrite or blind retry. This detects changes between the service's snapshot and mutation, **not** edits since a consumer's earlier `get()` call. A concurrent creation can recover through one validated conditional update. Removal of a concurrently deleted key is accepted only after confirming absence.

### Serialization and saved-data compatibility

No migration, new column or tagged envelope is required. Existing raw values keep their historical JSON-first interpretation, including REST Note `null` reading as legacy `null`. Previously ambiguous data cannot be reconstructed automatically: a historical raw `123` still reads as number `123`, even if the original caller supplied a string. Existing public legacy serialization helpers and StorageEntity reads remain unchanged. Both list stores use the safe encoding below for new writes, while preserving their respective historical read behavior.

New tenant and site `save()` values use:

| Input | Storage/readback |
|-------|------------------|
| Empty text, or text that parses as JSON, such as `"123"`, `"true"`, `"null"` or JSON-looking text | JSON-quoted text; reads back as the original string |
| Other text | Raw text; reads back unchanged |
| `number`, `boolean`, legacy `null` input | Legacy string representation, decoded with `JSON.parse` |
| `bigint` | JSON-quoted decimal text; reads as an exact **string**, without numeric precision loss |
| `Date` | ISO text; reads as a **string**, not a Date instance |
| Objects/arrays | `JSON.stringify`, then `JSON.parse` on read |

Older JSON-first readers can read new strings and bigint text correctly. Generics describe the expected shape and do not validate runtime data. This remains a text/JSON store, not a lossless arbitrary-JavaScript serializer: nested bigint, circular objects and invalid Dates throw before provisioning. Non-JSON values such as undefined/functions, non-finite numbers and custom prototypes are not guaranteed to round-trip. Validate application configuration after reading.

### Example: save and read text safely

```tsx
import * as React from 'react';
import { useSPFxTenantKeyValueStore } from '@apvee/spfx-react-toolkit';

function TenantConfigEditor() {
  const { save, get, isReady, isWriting, error, writeError } = useSPFxTenantKeyValueStore();
  const [text, setText] = React.useState<string>('');
  const run = async (): Promise<void> => {
    try {
      await save('my-app:version', '123', 'Application version');
      const item = await get<unknown>('my-app:version');
      if (typeof item?.value === 'string') setText(item.value);
    } catch {
      // Write failures are rendered through writeError; reads use error + fallback.
    }
  };
  return <>
    <button disabled={!isReady || isWriting} onClick={run}>Save and read</button>
    <p>{text}</p>
    {(error ?? writeError) && <p role="alert">{(error ?? writeError)?.message}</p>}
  </>;
}
```

For a bigint, `await save('my-app:identifier', BigInt('9007199254740993'))` is read using `get<string>()`. Reconstruct a BigInt or Date explicitly only after validating the returned text. Existing legacy numeric values may already have lost precision and are not silently rewritten.

### Demo and verification

The **Tenant → Tenant storage** sample runs CRUD only on explicit Run. Select Text/JSON/Big integer/Date to exercise serialization and inspect the read value's type. Empty description clears metadata. The list result reports the total fetched count and previews the first ten entries. Use disposable keys and remove them afterward. Local automated tests execute the real hook and panel with simulated service boundaries; authenticated tenant checks are tracked separately in [SharePoint validation](../../SHAREPOINT-VALIDATION.md).

### Source

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxTenantKeyValueStore.ts)

---

## useSPFxSiteKeyValueStore

Shares a hidden `SiteKeyValueStore` in `pageContext.site.absoluteUrl`, the root web of the current site collection. Root pages and every subsite use the same data; another site collection has its own store. The current web URL and tenant app catalog are not used as this store's target. New writes use the same safe encoding and ETag preconditions as the tenant store; historical site data keeps its previous interpretation without migration.

### Signature and public types

```typescript
function useSPFxSiteKeyValueStore(): SPFxSiteKeyValueStoreResult

interface SPFxSiteKeyValueStoreItem<T = unknown> {
  readonly key: string;
  readonly value: T;
  readonly description: string | undefined;
  readonly id: number;
}

interface SPFxSiteKeyValueStoreResult {
  readonly isLoading: boolean;
  readonly error: Error | undefined;
  readonly isWriting: boolean;
  readonly writeError: Error | undefined;
  readonly canWrite: boolean;
  readonly isReady: boolean;
  readonly get: <T = unknown>(key: string) => Promise<SPFxSiteKeyValueStoreItem<T> | undefined>;
  readonly list: () => Promise<SPFxSiteKeyValueStoreItem<unknown>[]>;
  readonly save: <T = unknown>(key: string, value: T, description?: string) => Promise<void>;
  readonly remove: (key: string) => Promise<void>;
}

```

### Readiness, errors and lifecycle

Mount checks permissions without provisioning or fetching values. `isReady` means the SPHttpClient and a nonempty normalized collection URL are available; it does not mean the list exists or the user has access. Operations invoked before readiness reject.

`get` and `list` never provision. Confirmed absent list/key returns `undefined`; an absent list returns `[]`. Other read failures set `error` and resolve with those same fallbacks. A fallback alone cannot establish absence. Display the current render's `error` before displaying a missing result; inspecting the error captured by a callback immediately after `await` can read stale React state. The standalone [service](../services/INDEX.md#createspfxsitekeyvaluestoreservice) rejects read failures instead.

`save` provisions or repairs if needed. `remove` never provisions and confirmed absence is a no-op. Failed writes set `writeError` and reject: catch the returned promise. Only `save` and the standalone service's explicit `ensureListReady` perform setup.

`isLoading` and `isWriting` independently track all pending operations for the current identity. Starting a read/write clears that channel's error; only its latest-started operation can publish a failure. Client or collection changes reset visible state and permission results; subsite navigation within the collection retains storage identity. Old callbacks keep their original collection/client, and obsolete completions or unmount do not publish state. Dispatched requests still settle and remote writes are not cancelled or serialized.

### Storage, grants and serialization

The root list uses generic list template 100, `Hidden: true`, `NoCrawl: true`, indexed unique `Title`, and multiline text `Value` and `Description`. Reads validate field types without repair. Writes reject incompatible schema or failed uniqueness enforcement without deleting existing data. `list` follows all continuation pages in server Title order and rejects incomplete, foreign or repeated continuations; a later-page failure becomes the hook's error/fallback, not partial success.

Read access depends on actual root-web/list permissions. `canWrite` checks effective Add, Edit and Delete Items on an existing list, including unique list permissions. An absent list also requires Manage Lists on the root web. The hook refreshes this advisory indicator on mount and after each current-identity read/write settles, including failures. Failed checks yield false; only the latest permission check can update state, and obsolete/unmounted operations never start a refresh. SharePoint enforces authorization, schema-management rights and item restrictions even when `canWrite` is true; subsite membership alone is insufficient.

Keys follow SharePoint's case-insensitive server collation and unique-column comparison; returned keys retain stored Title casing. Updates and deletion require the observed item ETag. If collection metadata omits it, the service retrieves a fresh item snapshot and its header ETag, preserving metadata from that same snapshot. Missing/invalid ETags and HTTP 412 conflicts reject without a wildcard overwrite or blind retry. This protects the service's snapshot-to-mutation interval, not changes since a consumer's previous get(). Creation recovery allows one validated conditional update and exposes that update's actual failure. Snapshot/deletion 404 succeeds only after verifying absence; a recreated key is not deleted. Omitted update descriptions preserve existing metadata; `''` clears it.

New site writes use the [safe list-store encoding](#serialization-and-saved-data-compatibility): empty and JSON-like strings are JSON-quoted so their text survives readback; bigint becomes exact decimal **string**, Date remains ISO **string**, and JSON-compatible objects/arrays retain JSON encoding. Historical JSON-first readers can read these new values. Circular objects, nested bigint and invalid Dates throw before any setup request. Other non-JSON values/prototypes have no complete round-trip guarantee; `T` is not runtime validation.

No existing rows, fields or public signatures are migrated. Historical raw `123`/`true` still decode as number/boolean, and any previously lost bigint precision cannot be reconstructed. The site-specific blank Note contract is preserved: REST `Value: null` reads as `''`, while the stored text `'null'` decodes to legacy `null`. New empty text is stored as `'""'`, avoiding SharePoint's empty-Note normalization. The tenant's historical REST-null behavior remains unchanged.

### Observable sample

The **Site** panel offers Text, JSON, Big integer and Date inputs for explicit Save. Read the key afterward to inspect the value and its runtime type; empty description explicitly clears metadata. Try Text `123`/empty text and Big integer `9007199254740993` to observe string-preserving/exact readback. Date reads as ISO string. List reports the complete count and previews the first ten entries. Conflicts/errors never appear as successful saves or missing records. Tests execute the actual panel/hook locally with simulated service boundaries; [authenticated root/subsite acceptance](../../SHAREPOINT-VALIDATION.md) remains separate.

### Example: Manual read and save

```tsx
import * as React from 'react';
import { useSPFxSiteKeyValueStore, SPFxSiteKeyValueStoreItem } from '@apvee/spfx-react-toolkit';

function SiteSettings() {
  const store = useSPFxSiteKeyValueStore();
  const [item, setItem] = React.useState<SPFxSiteKeyValueStoreItem<unknown>>();
  const [hasRead, setHasRead] = React.useState(false);
  const [message, setMessage] = React.useState('');

  const load = async () => {
    setHasRead(false);
    setMessage('');
    try {
      setItem(await store.get('toolkit-site-demo-setting'));
      setHasRead(true);
    } catch (failure) {
      setMessage(String(failure)); // For example, an operation before readiness.
    }
  };

  const save = async () => {
    setMessage('');
    try {
      await store.save('toolkit-site-demo-setting', { enabled: true }, 'Collection setting');
      setMessage('Saved');
    } catch (failure) {
      setMessage(String(failure));
    }
  };

  return (
    <div>
      <button onClick={load} disabled={!store.isReady || store.isLoading}>Read</button>
      <button onClick={save} disabled={!store.isReady || !store.canWrite || store.isWriting}>Save</button>
      {store.error ? <p role="alert">{store.error.message}</p> :
        hasRead && !store.isLoading && <p>{item ? JSON.stringify(item.value) : 'Key not found'}</p>}
      {store.writeError && <p role="alert">{store.writeError.message}</p>}
      {message && <p>{message}</p>}
    </div>
  );
}
```

Use a disposable key prefix for testing and catch `remove` rejections just as for `save`. The [Site sample and real-host checklist](../../SHAREPOINT-VALIDATION.md#site-collection-key-value-store) cover CRUD, sharing and permissions.

### Source

[View hook source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxSiteKeyValueStore.ts)

---

## See Also

- [HTTP Client Hooks](./http-clients.md) - API access
- [PnPjs Hooks](./pnpjs.md) - SharePoint data access
- [Context Hooks](./context.md) - SPFx context
- [Performance Hooks](./performance.md) - Performance measurement & diagnostics

---

*Maintained against public source declarations. Last updated: October 6, 2026*
