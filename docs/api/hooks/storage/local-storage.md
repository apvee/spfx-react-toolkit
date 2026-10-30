# useSPFxLocalStorage

[Storage overview](../../INDEX.md#storage-hooks) · [Usage guide](../../../guides/storage.md)

```ts
import { useSPFxLocalStorage } from '@apvee/spfx-react-toolkit/hooks';
```

Requires the matching SPFx host provider. See [provider selection](../../../getting-started/choosing-a-provider.md).


Persistent storage that survives browser restarts. Data is namespaced per web part instance.

## Signature

```typescript
function useSPFxLocalStorage<T>(
  key: string,
  defaultValue: T
): SPFxStorageHook<T>
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `key` | `string` | Yes | Storage key (automatically namespaced) |
| `defaultValue` | `T` | Yes | Fallback when the current key is missing or unreadable |

## Returns

An object containing:

- `value`: Current value of type `T`
- `setValue`: Setter accepting a value or updater function
- `remove`: Removes the stored key and resets `value` to the current default

## Default and persistence behavior

The default seeds the current key when no readable JSON exists. Changing only `defaultValue` preserves the current value, including when the default is an inline object. `remove()` and a matching browser storage deletion event reset to the latest default. Changing key or instance ID reloads that scoped key, using the current default if needed. Persistence is best effort: blocked storage, invalid JSON and quota failures do not throw from this hook. Browser storage event rules apply; this is not a same-document synchronization bus.

## Namespacing

Keys are automatically namespaced with the web part instance ID to prevent conflicts between multiple instances of the same web part.

## Example: User Preferences

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

## Example: Search History

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

## Source

[View source](../../../../packages/spfx-react-toolkit/src/hooks/useSPFxStorage.ts)

---
