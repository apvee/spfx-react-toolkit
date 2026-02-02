# Storage Hooks

> Hooks for data persistence across browser storage and OneDrive

## Overview

These hooks provide access to browser storage (localStorage, sessionStorage) and OneDrive app-specific data storage.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxLocalStorage`](#usespfxlocalstorage) | `[value, setValue]` | Browser localStorage with namespace |
| [`useSPFxSessionStorage`](#usespfxsessionstorage) | `[value, setValue]` | Browser sessionStorage with namespace |
| [`useSPFxOneDriveAppData`](#usespfxonedriveappdata) | `SPFxOneDriveAppDataResult` | OneDrive app-specific storage |

---

## useSPFxLocalStorage

Persistent storage that survives browser restarts. Data is namespaced per web part instance.

### Signature

```typescript
function useSPFxLocalStorage<T>(
  key: string,
  defaultValue: T
): [T, (value: T | ((prev: T) => T)) => void]
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `key` | `string` | Yes | Storage key (automatically namespaced) |
| `defaultValue` | `T` | Yes | Default value when key doesn't exist |

### Returns

A tuple containing:
- `[0]`: Current value of type `T`
- `[1]`: Setter function (accepts value or updater function)

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
  const [preferences, setPreferences] = useSPFxLocalStorage<UserPreferences>(
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
  const [history, setHistory] = useSPFxLocalStorage<string[]>('searchHistory', []);
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

[View source](../../src/hooks/useSPFxLocalStorage.ts)

---

## useSPFxSessionStorage

Temporary storage that persists only for the current browser session.

### Signature

```typescript
function useSPFxSessionStorage<T>(
  key: string,
  defaultValue: T
): [T, (value: T | ((prev: T) => T)) => void]
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `key` | `string` | Yes | Storage key (automatically namespaced) |
| `defaultValue` | `T` | Yes | Default value when key doesn't exist |

### Returns

A tuple containing:
- `[0]`: Current value of type `T`
- `[1]`: Setter function (accepts value or updater function)

### Example: Form Draft

```tsx
import { useSPFxSessionStorage } from '@apvee/spfx-react-toolkit';

interface FormData {
  title: string;
  description: string;
  category: string;
}

function FormWithDraft() {
  const [draft, setDraft] = useSPFxSessionStorage<FormData>('formDraft', {
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
  const [wizard, setWizard] = useSPFxSessionStorage<WizardState>('wizard', {
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

[View source](../../src/hooks/useSPFxSessionStorage.ts)

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

[View source](../../../src/hooks/useSPFxOneDriveAppData.ts)

---

## See Also

- [HTTP Client Hooks](./http-clients.md) - API access
- [PnPjs Hooks](./pnpjs.md) - SharePoint data access
- [Context Hooks](./context.md) - SPFx context

---

*Generated from JSDoc comments. Last updated: February 2, 2026*
