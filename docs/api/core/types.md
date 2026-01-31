# Core Types

> TypeScript type definitions for SPFx React Toolkit

## Overview

The core module exports essential type definitions used throughout the library. These types provide full TypeScript support with strict typing and no `any` usage.

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
| `'ACE'` | Adaptive Card Extension (Viva Connections) |

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

[View source](../../src/core/types.ts#L11)

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

function getInstanceId<T extends {}>(component: SPFxComponent<T>): string {
  return component.context.instanceId;
}
```

### Source

[View source](../../src/core/types.ts#L22)

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
  
  // Type narrowing for WebPart-specific properties
  if (kind === 'WebPart') {
    const wpContext = spfxContext as WebPartContext;
    console.log(wpContext.domElement); // WebPart-specific
  }
  
  // Common properties available on all contexts
  console.log(spfxContext.pageContext.web.title);
}
```

### Source

[View source](../../src/core/types.ts#L48)

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
  const { width, height } = useSPFxContainerInfo();
  
  return (
    <div>
      <p>Container: {width}x{height}px</p>
    </div>
  );
}
```

### Source

[View source](../../src/core/types.ts#L61)

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

[View source](../../src/core/types.ts#L84)

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

For component-specific properties, use type narrowing with the `kind` property.

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

[View source](../../src/core/types.ts#L101)

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

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
