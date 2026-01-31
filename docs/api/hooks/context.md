# Context Hooks

> Hooks for accessing SPFx context and instance metadata

## Overview

Context hooks provide access to the core SPFx runtime context, page context, service scope, and instance metadata.

| Hook | Returns | Description |
|------|---------|-------------|
| [`useSPFxContext`](#usespfxcontext) | `SPFxContextValue` | Core SPFx context |
| [`useSPFxPageContext`](#usespfxpagecontext) | `PageContext` | SharePoint page context |
| [`useSPFxServiceScope`](#usespfxservicescope) | `SPFxServiceScopeInfo` | Service scope for DI |
| [`useSPFxInstanceInfo`](#usespfxinstanceinfo) | `SPFxInstanceInfo` | Instance metadata |

---

## useSPFxContext

Access the core SPFx context metadata.

### Signature

```typescript
function useSPFxContext(): SPFxContextValue
```

### Returns

```typescript
interface SPFxContextValue {
  /** Unique identifier for this SPFx instance */
  readonly instanceId: string;
  
  /** SPFx context object (WebPartContext, etc.) */
  readonly spfxContext: SPFxContextType;
  
  /** Type of host component */
  readonly kind: HostKind;
}
```

### Description

Returns core SPFx context metadata:
- `instanceId` - Unique identifier for this SPFx instance
- `spfxContext` - The native SPFx context object
- `kind` - Type of host component ('WebPart', 'AppCustomizer', etc.)

### Example

```tsx
import { useSPFxContext } from '@apvee/spfx-react-toolkit';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

function MyComponent() {
  const { instanceId, spfxContext, kind } = useSPFxContext();
  
  // Type narrowing for component-specific properties
  if (kind === 'WebPart') {
    const wpContext = spfxContext as WebPartContext;
    console.log(wpContext.domElement); // WebPart-specific
  }
  
  return (
    <div>
      <p>Instance: {instanceId}</p>
      <p>Type: {kind}</p>
      <p>Site: {spfxContext.pageContext.web.title}</p>
    </div>
  );
}
```

### Throws

- `Error` if used outside of an SPFx provider

### Source

[View source](../../src/hooks/useSPFxContext.ts)

---

## useSPFxPageContext

Access SharePoint page context.

### Signature

```typescript
function useSPFxPageContext(): PageContext
```

### Returns

`PageContext` - SharePoint page context object from `@microsoft/sp-page-context`

### Description

Provides access to SharePoint page context containing information about:
- Site collection and web
- Current user
- Current list and list item (if applicable)
- Teams context (if running in Teams)
- Culture and locale settings
- Permissions and capabilities

This hook consumes `PageContext` from SPFx ServiceScope using dependency injection. The service is consumed lazily and cached for performance.

### Example

```tsx
import { useSPFxPageContext } from '@apvee/spfx-react-toolkit';

function MyComponent() {
  const pageContext = useSPFxPageContext();
  
  return (
    <div>
      <h2>Site Information</h2>
      <p>Site: {pageContext.web.title}</p>
      <p>URL: {pageContext.web.absoluteUrl}</p>
      
      <h2>User Information</h2>
      <p>User: {pageContext.user.displayName}</p>
      <p>Email: {pageContext.user.email}</p>
      
      <h2>Locale</h2>
      <p>Language: {pageContext.cultureInfo.currentUICultureName}</p>
    </div>
  );
}
```

### Source

[View source](../../src/hooks/useSPFxPageContext.ts)

---

## useSPFxServiceScope

Access SPFx ServiceScope for dependency injection.

### Signature

```typescript
function useSPFxServiceScope(): SPFxServiceScopeInfo
```

### Returns

```typescript
interface SPFxServiceScopeInfo {
  /** Native ServiceScope instance from SPFx */
  readonly serviceScope: ServiceScope | undefined;
  
  /** Consume a service from the ServiceScope */
  readonly consume: <T>(serviceKey: ServiceKey<T>) => T;
}
```

### Description

ServiceScope is SPFx's dependency injection container that provides:
- Access to built-in SPFx services
- Access to custom registered services
- Service lifecycle management
- Service isolation per scope

**Common built-in services:**
- `PageContext` (via @microsoft/sp-page-context)
- `HttpClient` (via @microsoft/sp-http)
- `MSGraphClientFactory` (via @microsoft/sp-http)
- `SPPermission` (via @microsoft/sp-page-context)
- `EventAggregator` (via @microsoft/sp-core-library)

> **Note:** Most common services have dedicated hooks (`useSPFxHttpClient`, `useSPFxMSGraphClient`). Use this hook for custom services or advanced scenarios.

### Example: Consuming a Custom Service

```tsx
import { useSPFxServiceScope } from '@apvee/spfx-react-toolkit';
import { ServiceKey } from '@microsoft/sp-core-library';

// Define service interface
interface IMyService {
  getData(): Promise<string[]>;
}

// Service key (typically defined in service file)
const MyServiceKey = ServiceKey.create<IMyService>('my-solution:IMyService', MyService);

function MyComponent() {
  const { consume } = useSPFxServiceScope();
  const [data, setData] = React.useState<string[]>([]);
  
  React.useEffect(() => {
    // Consume the custom service
    const myService = consume<IMyService>(MyServiceKey);
    myService.getData().then(setData);
  }, [consume]);
  
  return (
    <ul>
      {data.map((item, i) => <li key={i}>{item}</li>)}
    </ul>
  );
}
```

### Example: Accessing EventAggregator

```tsx
import { useSPFxServiceScope } from '@apvee/spfx-react-toolkit';
import { useEffect } from 'react';

function MyComponent() {
  const { serviceScope } = useSPFxServiceScope();
  
  useEffect(() => {
    if (!serviceScope) return;
    
    // Subscribe to cross-component events
    const subscription = serviceScope.consume(EventAggregator.serviceKey)
      .subscribe('ItemSelected', (args: { itemId: number }) => {
        console.log('Item selected:', args.itemId);
      });
    
    return () => subscription.dispose();
  }, [serviceScope]);
  
  return <div>Listening for events...</div>;
}
```

### Source

[View source](../../src/hooks/useSPFxServiceScope.ts)

---

## useSPFxInstanceInfo

Access SPFx instance metadata.

### Signature

```typescript
function useSPFxInstanceInfo(): SPFxInstanceInfo
```

### Returns

```typescript
interface SPFxInstanceInfo {
  /** Unique identifier for this SPFx instance */
  readonly id: string;
  
  /** Type of SPFx component (WebPart, AppCustomizer, etc.) */
  readonly kind: HostKind;
}
```

### Description

Provides simplified access to instance metadata:
- `id` - Unique identifier for this SPFx instance
- `kind` - Type of component ('WebPart', 'AppCustomizer', etc.)

**Use cases:**
- Logging and telemetry
- Conditional logic based on host type
- Scoped storage keys
- Debug information

### Example

```tsx
import { useSPFxInstanceInfo } from '@apvee/spfx-react-toolkit';

function MyComponent() {
  const { id, kind } = useSPFxInstanceInfo();
  
  // Use instance ID for scoped logging
  const log = (message: string) => {
    console.log(`[${kind}:${id}] ${message}`);
  };
  
  return (
    <div>
      <p>Instance ID: {id}</p>
      <p>Component Type: {kind}</p>
      <button onClick={() => log('Button clicked')}>
        Log Message
      </button>
    </div>
  );
}
```

### Example: Conditional Rendering

```tsx
import { useSPFxInstanceInfo } from '@apvee/spfx-react-toolkit';

function MyComponent() {
  const { kind } = useSPFxInstanceInfo();
  
  // Different UI based on host type
  if (kind === 'WebPart') {
    return <FullWebPartUI />;
  }
  
  if (kind === 'AppCustomizer') {
    return <CompactHeaderUI />;
  }
  
  return <DefaultUI />;
}
```

### Source

[View source](../../src/hooks/useSPFxInstanceInfo.ts)

---

## See Also

- [Providers](../core/providers.md) - Provider components
- [Types](../core/types.md) - Type definitions
- [Properties Hooks](./properties.md) - Property management

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
