// types.ts
// Core type definitions for SPFx React Toolkit

import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import type { BaseListViewCommandSet, BaseFieldCustomizer } from '@microsoft/sp-listview-extensibility';

/**
 * Type of SPFx host component
 */
export type HostKind = 
  | 'WebPart' 
  | 'AppCustomizer' 
  | 'FieldCustomizer' 
  | 'CommandSet' 
  | 'ACE';

/**
 * Union type for all SPFx component instances
 * Uses actual SPFx base classes for full type safety and API access
 */
export type SPFxComponent<TProps extends {} = {}> = 
  | BaseClientSideWebPart<TProps>
  | BaseApplicationCustomizer<TProps>
  | BaseListViewCommandSet<TProps>
  | BaseFieldCustomizer<TProps>;

/**
 * Union type for all SPFx context types
 * Provides type-safe access to common context properties across all SPFx components
 * 
 * @remarks
 * All SPFx contexts include these common properties:
 * - `pageContext` - SharePoint page context
 * - `serviceScope` - Service locator for SPFx services
 * - `instanceId` - Unique identifier for the component instance
 * 
 * For component-specific properties, use type narrowing or casting:
 * ```typescript
 * const ctx = useSPFxContext();
 * if (ctx.kind === 'WebPart') {
 *   const wpContext = ctx.spfxContext as WebPartContext;
 *   // Access WebPart-specific properties
 * }
 * ```
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type SPFxContextType = 
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  | BaseClientSideWebPart<any>['context']
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  | BaseApplicationCustomizer<any>['context']
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  | BaseListViewCommandSet<any>['context']
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  | BaseFieldCustomizer<any>['context'];

/**
 * Container size information
 */
export interface ContainerSize {
  readonly width: number;
  readonly height: number;
}

/**
 * Props accepted by SPFxProvider (instance-based API)
 * 
 * @template TProps - The properties type for the SPFx component (WebPart props, Extension props, etc.)
 * 
 * @example
 * ```tsx
 * // In your WebPart render():
 * public render(): void {
 *   const element = React.createElement(
 *     SPFxProvider,
 *     { instance: this },
 *     React.createElement(MyComponent)
 *   );
 *   ReactDom.render(element, this.domElement);
 * }
 * ```
 */
export interface SPFxProviderProps<TProps extends {} = {}> {
  /** SPFx component instance (WebPart, ApplicationCustomizer, etc.) */
  readonly instance: SPFxComponent<TProps>;
  
  /** Children to render */
  readonly children?: React.ReactNode;
}

/**
 * Context value provided by SPFxProvider
 * Contains only static metadata, no reactive state
 * 
 * @remarks
 * The `spfxContext` property provides type-safe access to common SPFx context properties
 * like `pageContext`, `serviceScope`, and `instanceId`. For component-specific properties,
 * use type narrowing with the `kind` property.
 */
export interface SPFxContextValue {
  /** Unique identifier for this SPFx instance */
  readonly instanceId: string;
  
  /** 
   * SPFx context object with full type safety
   * Provides access to common properties: pageContext, serviceScope, instanceId
   * For component-specific properties, use type narrowing based on `kind`
   */
  readonly spfxContext: SPFxContextType;
  
  /** Type of host component */
  readonly kind: HostKind;
}
