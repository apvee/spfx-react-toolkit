# Properties & Display Hooks

> Hooks for managing component properties and display mode

## Overview

These hooks provide access to SPFx component properties with bidirectional synchronization and display mode detection.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxProperties`](#usespfxproperties) | `SPFxPropertiesInfo<T>` | Bidirectional property management |
| [`useSPFxDisplayMode`](#usespfxdisplaymode) | `SPFxDisplayModeInfo` | Read/Edit mode detection |
| [`useSPFxIsEdit`](#usespfxisedit) | `boolean` | Shortcut for edit mode check |

---

## useSPFxProperties

Access and manage SPFx component properties with bidirectional synchronization.

### Signature

```typescript
function useSPFxProperties<TProps = unknown>(): SPFxPropertiesInfo<TProps>
```

### Type Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `TProps` | `unknown` | The properties interface type |

### Returns

```typescript
interface SPFxPropertiesInfo<TProps = unknown> {
  /** Current properties object */
  readonly properties: TProps | undefined;
  
  /** Update properties with partial updates (shallow merge) */
  readonly setProperties: (updates: Partial<TProps>) => void;
  
  /** Update properties using updater function */
  readonly updateProperties: (updater: (current: TProps | undefined) => TProps) => void;
}
```

### Description

Properties are the configuration values for WebParts/Extensions that:
- Are set via Property Pane
- Persist across page loads
- Are specific to each instance

**Synchronization is automatic:**
- Property Pane changes → Atom → Hook (automatic)
- Hook updates → Atom → SPFx properties (automatic)
- Property Pane refresh for WebParts (automatic)

### Example: Basic Usage

```tsx
import { useSPFxProperties } from '@apvee/spfx-react-toolkit';

interface IMyWebPartProps {
  title: string;
  description: string;
  showHeader: boolean;
}

function MyComponent() {
  const { properties, setProperties } = useSPFxProperties<IMyWebPartProps>();
  
  return (
    <div>
      <h1>{properties?.title ?? 'Default Title'}</h1>
      <p>{properties?.description}</p>
      
      <button onClick={() => setProperties({ title: 'New Title' })}>
        Update Title
      </button>
    </div>
  );
}
```

### Example: With Updater Function

```tsx
import { useSPFxProperties } from '@apvee/spfx-react-toolkit';

interface ICounterProps {
  count: number;
  lastUpdated: string;
}

function CounterComponent() {
  const { properties, updateProperties } = useSPFxProperties<ICounterProps>();
  
  const increment = () => {
    updateProperties(prev => ({
      count: (prev?.count ?? 0) + 1,
      lastUpdated: new Date().toISOString()
    }));
  };
  
  return (
    <div>
      <p>Count: {properties?.count ?? 0}</p>
      <p>Last Updated: {properties?.lastUpdated ?? 'Never'}</p>
      <button onClick={increment}>Increment</button>
    </div>
  );
}
```

### Example: Form with Properties

```tsx
import { useSPFxProperties, useSPFxDisplayMode } from '@apvee/spfx-react-toolkit';

interface IFormProps {
  formTitle: string;
  submitUrl: string;
  showLabels: boolean;
}

function ConfigurableForm() {
  const { properties, setProperties } = useSPFxProperties<IFormProps>();
  const { isEdit } = useSPFxDisplayMode();
  
  if (isEdit) {
    return (
      <div className="config-panel">
        <label>
          Form Title:
          <input 
            value={properties?.formTitle ?? ''} 
            onChange={(e) => setProperties({ formTitle: e.target.value })}
          />
        </label>
        <label>
          <input 
            type="checkbox"
            checked={properties?.showLabels ?? true}
            onChange={(e) => setProperties({ showLabels: e.target.checked })}
          />
          Show Labels
        </label>
      </div>
    );
  }
  
  return (
    <form>
      <h2>{properties?.formTitle}</h2>
      {/* Form fields */}
    </form>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxProperties.ts)

---

## useSPFxDisplayMode

Access SPFx display mode (Read/Edit).

### Signature

```typescript
function useSPFxDisplayMode(): SPFxDisplayModeInfo
```

### Returns

```typescript
interface SPFxDisplayModeInfo {
  /** Current display mode (Read/Edit) */
  readonly mode: DisplayMode;
  
  /** Whether currently in Edit mode */
  readonly isEdit: boolean;
  
  /** Whether currently in Read mode */
  readonly isRead: boolean;
}
```

### Description

Display mode controls whether the WebPart/Extension is in:
- **Read mode** (`DisplayMode.Read`): Normal viewing mode
- **Edit mode** (`DisplayMode.Edit`): Editing/configuration mode

> **Note:** `displayMode` is readonly in SPFx and controlled by SharePoint. It changes when the user clicks the Edit button on the page.

**Use cases:**
- Showing/hiding edit controls
- Conditional rendering based on mode
- Different layouts for read vs edit

### Example

```tsx
import { useSPFxDisplayMode } from '@apvee/spfx-react-toolkit';

function MyComponent() {
  const { mode, isEdit, isRead } = useSPFxDisplayMode();
  
  return (
    <div>
      <p>Mode: {isEdit ? 'Editing' : 'Reading'}</p>
      
      {isEdit && (
        <div className="edit-toolbar">
          <button>Add Item</button>
          <button>Configure</button>
        </div>
      )}
      
      <div className="content">
        {/* Main content */}
      </div>
    </div>
  );
}
```

### Example: Edit/Read Views

```tsx
import { useSPFxDisplayMode } from '@apvee/spfx-react-toolkit';

function MyComponent() {
  const { isEdit } = useSPFxDisplayMode();
  
  return isEdit ? <EditView /> : <ReadView />;
}

function EditView() {
  return (
    <div className="edit-view">
      <p>You are in edit mode. Configure your web part.</p>
      {/* Edit controls */}
    </div>
  );
}

function ReadView() {
  return (
    <div className="read-view">
      {/* Normal content display */}
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxDisplayMode.ts)

---

## useSPFxIsEdit

Shortcut hook to check if currently in Edit mode.

### Signature

```typescript
function useSPFxIsEdit(): boolean
```

### Returns

`boolean` - `true` if in Edit mode, `false` otherwise

### Description

A convenience hook equivalent to `useSPFxDisplayMode().isEdit`. Use this for simple edit mode checks when you don't need the full display mode info.

### Example

```tsx
import { useSPFxIsEdit } from '@apvee/spfx-react-toolkit';

function MyComponent() {
  const isEdit = useSPFxIsEdit();
  
  return (
    <div>
      <h1>My Web Part</h1>
      {isEdit && <ConfigButton />}
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxDisplayMode.ts#L89)

---

## See Also

- [Context Hooks](./context.md) - Context access
- [Storage Hooks](./storage.md) - Persistent storage
- [Providers](../core/providers.md) - Provider components

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
