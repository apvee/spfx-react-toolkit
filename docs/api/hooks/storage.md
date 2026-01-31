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

Cloud-synced storage using OneDrive app-specific data folder.

### Prerequisites

Requires Microsoft Graph permissions:
- `Files.ReadWrite.AppFolder`

### Signature

```typescript
function useSPFxOneDriveAppData<T>(
  fileName: string,
  defaultValue: T
): SPFxOneDriveAppDataResult<T>
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `fileName` | `string` | Yes | File name in app data folder |
| `defaultValue` | `T` | Yes | Default value when file doesn't exist |

### Returns

```typescript
interface SPFxOneDriveAppDataResult<T> {
  /** Current data value */
  readonly data: T;
  
  /** Loading state */
  readonly isLoading: boolean;
  
  /** Saving state */
  readonly isSaving: boolean;
  
  /** Error if any operation failed */
  readonly error: Error | undefined;
  
  /** Save data to OneDrive */
  readonly save: (value: T | ((prev: T) => T)) => Promise<void>;
  
  /** Reload data from OneDrive */
  readonly reload: () => Promise<void>;
}
```

### Example: Cloud-Synced Settings

```tsx
import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';

interface CloudSettings {
  favorites: string[];
  recentItems: { id: string; title: string; date: string }[];
  preferences: { notifications: boolean };
}

function CloudSyncedFavorites() {
  const { 
    data, 
    isLoading, 
    isSaving, 
    save, 
    error 
  } = useSPFxOneDriveAppData<CloudSettings>('app-settings.json', {
    favorites: [],
    recentItems: [],
    preferences: { notifications: true }
  });
  
  const addToFavorites = async (itemId: string) => {
    await save(prev => ({
      ...prev,
      favorites: [...prev.favorites, itemId]
    }));
  };
  
  const removeFromFavorites = async (itemId: string) => {
    await save(prev => ({
      ...prev,
      favorites: prev.favorites.filter(id => id !== itemId)
    }));
  };
  
  if (isLoading) return <Spinner label="Loading settings..." />;
  if (error) return <MessageBar messageBarType={MessageBarType.error}>{error.message}</MessageBar>;
  
  return (
    <div className="favorites">
      {isSaving && <span className="sync-indicator">Syncing...</span>}
      <h3>Favorites ({data.favorites.length})</h3>
      <ul>
        {data.favorites.map(id => (
          <li key={id}>
            {id}
            <button onClick={() => removeFromFavorites(id)}>Remove</button>
          </li>
        ))}
      </ul>
    </div>
  );
}
```

### Example: Cross-Device Sync

```tsx
import { useSPFxOneDriveAppData } from '@apvee/spfx-react-toolkit';

interface UserState {
  lastVisited: { url: string; title: string; timestamp: number }[];
  bookmarks: { id: string; url: string; title: string }[];
}

function CrossDeviceState() {
  const { data, save, reload, isLoading, isSaving } = useSPFxOneDriveAppData<UserState>(
    'user-state.json',
    { lastVisited: [], bookmarks: [] }
  );
  
  // Track page visits
  React.useEffect(() => {
    const visit = {
      url: window.location.href,
      title: document.title,
      timestamp: Date.now()
    };
    
    save(prev => ({
      ...prev,
      lastVisited: [visit, ...prev.lastVisited.slice(0, 9)]
    }));
  }, []);
  
  return (
    <div className="sync-status">
      <button onClick={reload} disabled={isLoading}>
        🔄 Sync from Cloud
      </button>
      {isSaving && <span>Saving to OneDrive...</span>}
      <p>Last visited ({data.lastVisited.length} items synced across devices)</p>
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxOneDriveAppData.ts)

---

## See Also

- [HTTP Client Hooks](./http-clients.md) - API access
- [PnPjs Hooks](./pnpjs.md) - SharePoint data access
- [Context Hooks](./context.md) - SPFx context

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
