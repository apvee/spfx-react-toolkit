# Core Module

> Providers and types for SPFx React integration

## Overview

The Core module provides the foundation for SPFx React Toolkit:

- **Providers** - React context providers for each SPFx component type
- **Types** - TypeScript type definitions

## Providers

React context providers that wrap your SPFx components and enable hook usage.

| Provider | SPFx Component Type | Documentation |
|----------|---------------------|---------------|
| `SPFxWebPartProvider` | WebParts | [providers.md](./providers.md#spfxwebpartprovider) |
| `SPFxApplicationCustomizerProvider` | Application Customizers | [providers.md](./providers.md#spfxapplicationcustomizerprovider) |
| `SPFxFieldCustomizerProvider` | Field Customizers | [providers.md](./providers.md#spfxfieldcustomizerprovider) |
| `SPFxListViewCommandSetProvider` | ListView Command Sets | [providers.md](./providers.md#spfxlistviewcommandsetprovider) |

## Types

Core TypeScript type definitions.

| Type | Description | Documentation |
|------|-------------|---------------|
| `HostKind` | SPFx component type discriminator | [types.md](./types.md#hostkind) |
| `SPFxComponent` | Union of all SPFx component types | [types.md](./types.md#spfxcomponent) |
| `SPFxContextType` | Union of all SPFx context types | [types.md](./types.md#spfxcontexttype) |
| `ContainerSize` | Container dimensions interface | [types.md](./types.md#containersize) |
| `SPFxProviderProps` | Provider component props | [types.md](./types.md#spfxproviderprops) |
| `SPFxContextValue` | Context value interface | [types.md](./types.md#spfxcontextvalue) |

## Quick Start

### 1. Choose Your Provider

```tsx
// WebPart
import { SPFxWebPartProvider } from '@apvee/spfx-react-toolkit';

// Application Customizer
import { SPFxApplicationCustomizerProvider } from '@apvee/spfx-react-toolkit';

// Field Customizer
import { SPFxFieldCustomizerProvider } from '@apvee/spfx-react-toolkit';

// ListView Command Set
import { SPFxListViewCommandSetProvider } from '@apvee/spfx-react-toolkit';
```

### 2. Wrap Your Component

```tsx
// In WebPart render():
const element = React.createElement(
  SPFxWebPartProvider,
  { instance: this },
  React.createElement(MyComponent)
);
ReactDom.render(element, this.domElement);
```

### 3. Use Hooks

```tsx
import { useSPFxContext, useSPFxPageContext } from '@apvee/spfx-react-toolkit';

function MyComponent() {
  const { kind, instanceId } = useSPFxContext();
  const pageContext = useSPFxPageContext();
  
  return (
    <div>
      <p>Component Type: {kind}</p>
      <p>Site: {pageContext.web.title}</p>
    </div>
  );
}
```

## Import Example

```typescript
// Import providers
import { 
  SPFxWebPartProvider,
  SPFxApplicationCustomizerProvider,
  SPFxFieldCustomizerProvider,
  SPFxListViewCommandSetProvider
} from '@apvee/spfx-react-toolkit';

// Import types
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

- [Hooks Module](../hooks/INDEX.md) - React hooks
- [Introduction](../../INTRODUCTION.md) - Getting started

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
