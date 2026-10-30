# useSPFxOneDriveAppData

[Storage overview](../../INDEX.md#storage-hooks) · [Usage guide](../../../guides/storage.md)

```ts
import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).


Cloud-synced JSON storage using OneDrive app-specific data folder.

## Prerequisites

Requires Microsoft Graph permissions:
- `Files.ReadWrite` or `Files.ReadWrite.AppFolder`

## Signature

```typescript
function useSPFxOneDriveAppData<T = unknown>(
  fileName: string,
  options?: SPFxOneDriveAppDataOptions<T>
): SPFxOneDriveAppDataResult<T>
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `fileName` | `string` | Yes | File name in app data folder (e.g., `'config.json'`) |
| `options` | `SPFxOneDriveAppDataOptions<T>` | No | Configuration options |

## Options

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

## Returns

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

## Request ownership and lazy defaults

`defaultValue` seeds `data` even with `autoFetch: false`; changing file, folder or Graph client seeds the new identity from the current default without forcing a lazy fetch. Changing an inline default alone does not reload. A current missing-file response is required before `createIfMissing` writes a default.

Latest-started reads and writes own the local data/error state together. An older read, 404 or write completion cannot replace a newer write's local data. `isLoading` and `isWriting` track pending reads and writes separately for the current identity. Identity changes and unmount ignore obsolete completions. Requests still settle for their callers; remote writes are not cancelled or serialized. `load()` records read errors in `error`; `write()` records `writeError` and rejects on failure. These guards do not provide remote transaction ordering.

## Example: Basic Usage with Auto-Fetch

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

## Example: With Folder Namespace

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

## Example: Create If Missing

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

## Example: Manual Load (Lazy)

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

## Example: CRUD-like Operations

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

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxOneDriveAppData.ts)

---
