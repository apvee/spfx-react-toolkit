# Core Types

> TypeScript type definitions for SPFx React Toolkit

## Overview

The core module exports essential type definitions used throughout the library. These types provide full TypeScript support with public TypeScript declarations.

---

## HostKind

Type of SPFx host component.

### Definition

```typescript
type HostKind = 
  | 'WebPart' 
  | 'AppCustomizer' 
  | 'FieldCustomizer' 
  | 'CommandSet' 
  | 'ACE';
```

### Values

| Value | Description |
|-------|-------------|
| `'WebPart'` | Client-side web part |
| `'AppCustomizer'` | Application customizer extension |
| `'FieldCustomizer'` | Field customizer extension |
| `'CommandSet'` | ListView command set extension |
| `'ACE'` | Reserved host discriminator; no ACE provider is exported |

### Example

```typescript
import { useSPFxContext } from '@apvee/spfx-react-toolkit';

function MyComponent() {
  const { kind } = useSPFxContext();
  
  switch (kind) {
    case 'WebPart':
      return <WebPartUI />;
    case 'AppCustomizer':
      return <HeaderUI />;
    default:
      return <DefaultUI />;
  }
}
```

### Source

[View source](../../../packages/spfx-react-toolkit/src/core/types.ts)

---

## SPFxComponent

Union type for all SPFx component instances.

### Definition

```typescript
type SPFxComponent<TProps extends {} = {}> = 
  | BaseClientSideWebPart<TProps>
  | BaseApplicationCustomizer<TProps>
  | BaseListViewCommandSet<TProps>
  | BaseFieldCustomizer<TProps>;
```

### Description

This type represents any SPFx component instance that can be passed to a provider. It uses actual SPFx base classes for full type safety and API access.

### Example

```typescript
import type { SPFxComponent } from '@apvee/spfx-react-toolkit';

// The union describes provider inputs; read metadata through public hooks.
type TitledHost = SPFxComponent<{ title: string }>;
```

### Source

[View source](../../../packages/spfx-react-toolkit/src/core/types.ts)

---

## SPFxContextType

Union type for all SPFx context types.

### Definition

```typescript
type SPFxContextType = 
  | BaseClientSideWebPart<any>['context']
  | BaseApplicationCustomizer<any>['context']
  | BaseListViewCommandSet<any>['context']
  | BaseFieldCustomizer<any>['context'];
```

### Description

Provides type-safe access to common SPFx context properties across all component types:

- `pageContext` - SharePoint page context
- `serviceScope` - Service locator for SPFx services
- `instanceId` - Unique identifier for the component instance

### Example

```typescript
import { useSPFxContext } from '@apvee/spfx-react-toolkit';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

function MyComponent() {
  const { spfxContext, kind } = useSPFxContext();
  
  // The matching WebPart provider establishes this explicit assertion.
  if (kind === 'WebPart') {
    const wpContext = spfxContext as WebPartContext;
    console.log(wpContext.domElement); // WebPart-specific
  }
  
  // Common properties available on all contexts
  console.log(spfxContext.pageContext.web.title);
}
```

### Source

[View source](../../../packages/spfx-react-toolkit/src/core/types.ts)

---

## ContainerSize

Container size information interface.

### Definition

```typescript
interface ContainerSize {
  readonly width: number;
  readonly height: number;
}
```

### Properties

| Property | Type | Description |
|----------|------|-------------|
| `width` | `number` | Container width in pixels |
| `height` | `number` | Container height in pixels |

### Example

```typescript
import { useSPFxContainerInfo } from '@apvee/spfx-react-toolkit';

function MyComponent() {
  const { size } = useSPFxContainerInfo();
  
  return (
    <div>
      <p>{size ? `Container: ${size.width}x${size.height}px` : 'No container'}</p>
    </div>
  );
}
```

### Source

[View source](../../../packages/spfx-react-toolkit/src/core/types.ts)

---

## SPFxProviderProps

Props accepted by SPFx providers.

### Definition

```typescript
interface SPFxProviderProps<TProps extends {} = {}> {
  /** SPFx component instance (WebPart, ApplicationCustomizer, etc.) */
  readonly instance: SPFxComponent<TProps>;
  
  /** Children to render */
  readonly children?: React.ReactNode;
}
```

### Type Parameters

| Parameter | Constraint | Default | Description |
|-----------|------------|---------|-------------|
| `TProps` | `extends {}` | `{}` | Properties type for the SPFx component |

### Properties

| Property | Type | Required | Description |
|----------|------|----------|-------------|
| `instance` | `SPFxComponent<TProps>` | Yes | The SPFx component instance |
| `children` | `React.ReactNode` | No | Children to render |

### Example

```typescript
import { SPFxWebPartProvider } from '@apvee/spfx-react-toolkit';

interface IMyWebPartProps {
  title: string;
  showHeader: boolean;
}

// In WebPart render():
const element = React.createElement(
  SPFxWebPartProvider<IMyWebPartProps>,
  { instance: this },
  React.createElement(MyComponent)
);
```

### Source

[View source](../../../packages/spfx-react-toolkit/src/core/types.ts)

---

## SPFxContextValue

Context value provided by SPFx providers.

### Definition

```typescript
interface SPFxContextValue {
  /** Unique identifier for this SPFx instance */
  readonly instanceId: string;
  
  /** SPFx context object with full type safety */
  readonly spfxContext: SPFxContextType;
  
  /** Type of host component */
  readonly kind: HostKind;
}
```

### Properties

| Property | Type | Description |
|----------|------|-------------|
| `instanceId` | `string` | Unique identifier for this SPFx instance |
| `spfxContext` | `SPFxContextType` | SPFx context with common properties |
| `kind` | `HostKind` | Type of host component |

### Description

Contains only static metadata, no reactive state. The `spfxContext` property provides type-safe access to common SPFx context properties like `pageContext`, `serviceScope`, and `instanceId`.

`kind` and `spfxContext` are separate fields, not a discriminated union: checking `kind` does not narrow the context type in TypeScript. Where the matching provider is known, explicitly assert its native SPFx context type, as in the `SPFxContextType` example above. No host-specific context hooks are exported.

### Example

```typescript
import { useSPFxContext } from '@apvee/spfx-react-toolkit';

function MyComponent() {
  const { instanceId, spfxContext, kind } = useSPFxContext();
  
  return (
    <div>
      <p>Instance: {instanceId}</p>
      <p>Type: {kind}</p>
      <p>Site: {spfxContext.pageContext.web.title}</p>
    </div>
  );
}
```

### Source

[View source](../../../packages/spfx-react-toolkit/src/core/types.ts)

---

## Type Imports

Import types from the main package:

```typescript
import type { 
  HostKind,
  SPFxComponent,
  SPFxContextType,
  ContainerSize,
  SPFxProviderProps,
  SPFxContextValue
} from '@apvee/spfx-react-toolkit';
```

---

## See Also

- [Providers](./providers.md) - Provider components
- [Context Hooks](../hooks/context.md) - Hooks using these types
