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

[View source](../../../packages/spfx-react-toolkit/src/core/provider-webpart.tsx)

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

Placeholders can appear after initialization and be disposed during navigation. Subscribe to `changedEvent`, retry creation when it changes, and bind disposal to the placeholder captured by each callback. A late disposal callback from an old placeholder must not unmount a replacement.

```tsx
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { SPFxApplicationCustomizerProvider, useSPFxProperties } from '@apvee/spfx-react-toolkit';

interface IMyCustomizerProps {
  headerMessage: string;
}

const HeaderComponent: React.FC = () => {
  const { properties } = useSPFxProperties<IMyCustomizerProps>();
  return <header>{properties?.headerMessage ?? 'Welcome'}</header>;
};

export default class MyApplicationCustomizer extends BaseApplicationCustomizer<IMyCustomizerProps> {
  private topPlaceholder?: PlaceholderContent;

  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this.renderTopPlaceholder);
    this.renderTopPlaceholder();
    return Promise.resolve();
  }

  public onDispose(): void {
    this.context.placeholderProvider.changedEvent.remove(this, this.renderTopPlaceholder);
    const placeholder = this.topPlaceholder;
    this.topPlaceholder = undefined;
    if (placeholder) {
      ReactDom.unmountComponentAtNode(placeholder.domElement);
      placeholder.dispose();
    }
  }

  private renderTopPlaceholder = (): void => {
    if (!this.topPlaceholder) {
      const capturedPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        {
          onDispose: () => {
            if (capturedPlaceholder) {
              ReactDom.unmountComponentAtNode(capturedPlaceholder.domElement);
              if (this.topPlaceholder === capturedPlaceholder) {
                this.topPlaceholder = undefined;
              }
            }
          }
        }
      );
      this.topPlaceholder = capturedPlaceholder;
    }

    if (!this.topPlaceholder) return;
    const element = React.createElement(
      SPFxApplicationCustomizerProvider,
      { instance: this },
      React.createElement(HeaderComponent)
    );
    ReactDom.render(element, this.topPlaceholder.domElement);
  };
}
```

### Source

[View source](../../../packages/spfx-react-toolkit/src/core/provider-application-customizer.tsx)

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

[View source](../../../packages/spfx-react-toolkit/src/core/provider-field-customizer.tsx)

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

[View source](../../../packages/spfx-react-toolkit/src/core/provider-listview-commandset.tsx)

---

## Provider Features

All providers use the common runtime. The sample mounts WebPart and Application Customizer providers in their real host entry points; Field Customizer and ListView Command Set exports are checked but require separate real-host validation. See [SharePoint validation](../../SHAREPOINT-VALIDATION.md).

### Instance Isolation

Each SPFx instance gets its own isolated state store. Provider runtime state is scoped to each SPFx instance. Explicitly shared external clients, tenant data and application objects remain shared by their own contracts.

### Automatic Synchronization

- **Property Pane → React**: A host render reconciles top-level values, removals and changed references using a shallow snapshot; nested in-place mutations are not deep-observed
- **React → SPFx**: Property updates via `useSPFxProperties` sync back to SPFx

### Theme Subscription

Theme changes (light/dark mode) are automatically detected and propagated to hooks.

### Display Mode Tracking

Edit/Read mode changes are tracked and available via `useSPFxDisplayMode`.

### Lifecycle

The host must unmount React when it disposes its rendered surface. The provider removes its theme subscription and runtime listeners on unmount; scope replacement waits for the current ServiceScope and ignores obsolete readiness callbacks. Local lifecycle regression tests cover these paths, but they do not establish a universal memory or tenant-host guarantee.

### Container Observation

Container size changes are observed and available via `useSPFxContainerSize`.

---

## See Also

- [Types](./types.md) - Core type definitions
- [Context Hooks](../hooks/context.md) - Hooks for accessing context
- [Properties Hooks](../hooks/properties.md) - Property management hooks

---

*Generated from JSDoc comments. Last updated: January 31, 2026*
