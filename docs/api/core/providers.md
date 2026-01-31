# Core Providers

> React context providers for SPFx components

## Overview

SPFx React Toolkit provides specialized provider components for each SPFx component type. These providers wrap your React components and enable access to all toolkit hooks.

**Choose the provider matching your SPFx component type:**

| Provider | SPFx Component |
|----------|----------------|
| [`SPFxWebPartProvider`](#spfxwebpartprovider) | `BaseClientSideWebPart` |
| [`SPFxApplicationCustomizerProvider`](#spfxapplicationcustomizerprovider) | `BaseApplicationCustomizer` |
| [`SPFxFieldCustomizerProvider`](#spfxfieldcustomizerprovider) | `BaseFieldCustomizer` |
| [`SPFxListViewCommandSetProvider`](#spfxlistviewcommandsetprovider) | `BaseListViewCommandSet` |

---

## SPFxWebPartProvider

SPFx context provider for WebParts.

### Signature

```typescript
function SPFxWebPartProvider<TProps extends {} = {}>(
  props: SPFxWebPartProviderProps<TProps>
): JSX.Element
```

### Props

```typescript
interface SPFxWebPartProviderProps<TProps extends {} = {}> {
  /** The SPFx WebPart instance */
  instance: BaseClientSideWebPart<TProps>;
  
  /** Children to render within the provider */
  children?: React.ReactNode;
}
```

### Example

```tsx
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFxWebPartProvider } from '@apvee/spfx-react-toolkit';

interface IMyWebPartProps {
  title: string;
  description: string;
}

export default class MyWebPart extends BaseClientSideWebPart<IMyWebPartProps> {
  public render(): void {
    const element = React.createElement(
      SPFxWebPartProvider,
      { instance: this },
      React.createElement(MyComponent)
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
```

### Source

[View source](../../src/core/provider-webpart.tsx)

---

## SPFxApplicationCustomizerProvider

SPFx context provider for Application Customizers.

### Signature

```typescript
function SPFxApplicationCustomizerProvider<TProps extends {} = {}>(
  props: SPFxApplicationCustomizerProviderProps<TProps>
): JSX.Element
```

### Props

```typescript
interface SPFxApplicationCustomizerProviderProps<TProps extends {} = {}> {
  /** The SPFx Application Customizer instance */
  instance: BaseApplicationCustomizer<TProps>;
  
  /** Children to render within the provider */
  children?: React.ReactNode;
}
```

### Example

```tsx
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import { SPFxApplicationCustomizerProvider } from '@apvee/spfx-react-toolkit';

interface IMyCustomizerProps {
  headerMessage: string;
}

export default class MyApplicationCustomizer extends BaseApplicationCustomizer<IMyCustomizerProps> {
  public onInit(): Promise<void> {
    // Get header placeholder
    const placeholder = this.context.placeholderProvider.tryCreateContent(
      PlaceholderName.Top
    );

    if (placeholder) {
      const element = React.createElement(
        SPFxApplicationCustomizerProvider,
        { instance: this },
        React.createElement(HeaderComponent)
      );
      ReactDom.render(element, placeholder.domElement);
    }

    return Promise.resolve();
  }

  protected onDispose(): void {
    // Clean up React components
  }
}
```

### Source

[View source](../../src/core/provider-application-customizer.tsx)

---

## SPFxFieldCustomizerProvider

SPFx context provider for Field Customizers.

### Signature

```typescript
function SPFxFieldCustomizerProvider<TProps extends {} = {}>(
  props: SPFxFieldCustomizerProviderProps<TProps>
): JSX.Element
```

### Props

```typescript
interface SPFxFieldCustomizerProviderProps<TProps extends {} = {}> {
  /** The SPFx Field Customizer instance */
  instance: BaseFieldCustomizer<TProps>;
  
  /** Children to render within the provider */
  children?: React.ReactNode;
}
```

### Example

```tsx
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { 
  BaseFieldCustomizer, 
  IFieldCustomizerCellEventParameters 
} from '@microsoft/sp-listview-extensibility';
import { SPFxFieldCustomizerProvider } from '@apvee/spfx-react-toolkit';

interface IMyFieldProps {
  colorMapping: Record<string, string>;
}

export default class MyFieldCustomizer extends BaseFieldCustomizer<IMyFieldProps> {
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const element = React.createElement(
      SPFxFieldCustomizerProvider,
      { instance: this },
      React.createElement(FieldRenderer, {
        value: event.fieldValue,
        listItem: event.listItem
      })
    );
    ReactDom.render(element, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDom.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
```

### Source

[View source](../../src/core/provider-field-customizer.tsx)

---

## SPFxListViewCommandSetProvider

SPFx context provider for ListView Command Sets.

### Signature

```typescript
function SPFxListViewCommandSetProvider<TProps extends {} = {}>(
  props: SPFxListViewCommandSetProviderProps<TProps>
): JSX.Element
```

### Props

```typescript
interface SPFxListViewCommandSetProviderProps<TProps extends {} = {}> {
  /** The SPFx ListView Command Set instance */
  instance: BaseListViewCommandSet<TProps>;
  
  /** Children to render within the provider */
  children?: React.ReactNode;
}
```

### Example

```tsx
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { 
  BaseListViewCommandSet, 
  IListViewCommandSetExecuteEventParameters 
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { SPFxListViewCommandSetProvider } from '@apvee/spfx-react-toolkit';

interface IMyCommandSetProps {
  dialogTitle: string;
}

export default class MyCommandSet extends BaseListViewCommandSet<IMyCommandSetProps> {
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'SHOW_DETAILS':
        // Render dialog with provider
        const dialogContainer = document.createElement('div');
        document.body.appendChild(dialogContainer);
        
        const element = React.createElement(
          SPFxListViewCommandSetProvider,
          { instance: this },
          React.createElement(DetailsDialog, {
            selectedItems: this.context.listView.selectedRows,
            onClose: () => {
              ReactDom.unmountComponentAtNode(dialogContainer);
              document.body.removeChild(dialogContainer);
            }
          })
        );
        ReactDom.render(element, dialogContainer);
        break;
    }
  }
}
```

### Source

[View source](../../src/core/provider-listview-commandset.tsx)

---

## Provider Features

All providers share these capabilities:

### Instance Isolation

Each SPFx instance gets its own isolated state store. Multiple instances of the same WebPart on a page do not share state.

### Automatic Synchronization

- **Property Pane → React**: Changes in the Property Pane automatically update React state
- **React → SPFx**: Property updates via `useSPFxProperties` sync back to SPFx

### Theme Subscription

Theme changes (light/dark mode) are automatically detected and propagated to hooks.

### Display Mode Tracking

Edit/Read mode changes are tracked and available via `useSPFxDisplayMode`.

### Container Observation

Container size changes are observed and available via `useSPFxContainerSize`.

---

## See Also

- [Types](./types.md) - Core type definitions
- [Context Hooks](../hooks/context.md) - Hooks for accessing context
- [Properties Hooks](../hooks/properties.md) - Property management hooks

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
