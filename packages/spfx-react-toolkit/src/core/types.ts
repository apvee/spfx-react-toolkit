// types.ts
// Core type definitions for SPFx React Toolkit

import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import type { BaseListViewCommandSet, BaseFieldCustomizer } from '@microsoft/sp-listview-extensibility';

/**
 * Host discriminator returned by provider metadata. Current exported providers cover WebPart, AppCustomizer, FieldCustomizer and CommandSet; ACE is a reserved type member without an exported ACE provider.
 *
 * @example
 * ```ts
 * const kind: HostKind = 'WebPart';
 * ```
 */
export type HostKind = 
  | 'WebPart' 
  | 'AppCustomizer' 
  | 'FieldCustomizer' 
  | 'CommandSet' 
  | 'ACE';

/**
 * Union of the four supported SPFx component instances. TProps is the host properties type. Pass an initialized instance to its matching host-specific provider; access context and properties through the public React hooks rather than assuming shared public members on this union.
 *
 * @typeParam TProps - Properties type of the supplied host instance.
 *
 * @example
 * ```tsx
 * type ComponentWithTitle = SPFxComponent<{ title: string }>;
 * ```
 */
export type SPFxComponent<TProps extends {} = {}> = 
  | BaseClientSideWebPart<TProps>
  | BaseApplicationCustomizer<TProps>
  | BaseListViewCommandSet<TProps>
  | BaseFieldCustomizer<TProps>;

/**
 * Union of the four supported SPFx host context types. Common context fields are available directly; kind is a separate metadata value and does not automatically narrow this union. Prefer host-specific context hooks when accessing WebPart or extension-only APIs.
 *
 * @example
 * ```ts
 * function webUrl(context: SPFxContextType): string {
 *   return context.pageContext.web.absoluteUrl;
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
 * Width and height measurement in CSS pixels. This shape does not perform measurement; consumers obtain host measurements from useSPFxContainerSize.
 *
 * @example
 * ```ts
 * const size: ContainerSize = { width: 720, height: 300 };
 * ```
 */
export interface ContainerSize {
  /**
   * Measured container width in CSS pixels.
   */
  readonly width: number;
  /**
   * Measured container height in CSS pixels.
   */
  readonly height: number;
}

/**
 * Props accepted by SPFx providers (instance-based API)
 * 
 * @template TProps - The properties type for the SPFx component (WebPart props, Extension props, etc.)
 * 
 * @example
 * ```tsx
 * // In your WebPart render():
 * public render(): void {
 *   const element = React.createElement(
 *     SPFxWebPartProvider,
 *     { instance: this },
 *     React.createElement(MyComponent)
 *   );
 *   ReactDom.render(element, this.domElement);
 * }
 * ```
 */
export interface SPFxProviderProps<TProps extends {} = {}> {
  /**
   * Required, initialized SPFx component instance. Use the provider matching its host type; provider descendants render after its service scope is ready.
   */
  readonly instance: SPFxComponent<TProps>;
  
  /**
   * Optional React subtree receiving instance metadata, runtime state and host theme. Omission renders no application content.
   */
  readonly children?: React.ReactNode;
}

/**
 * Static host metadata returned by useSPFxContext beneath a host-specific provider. Property and theme updates are exposed through dedicated runtime hooks. kind does not form a discriminated union with spfxContext; use host-specific context hooks for typed host fields.
 *
 * @example
 * ```tsx
 * function HostLabel() {
 *   const metadata = useSPFxContext();
 *   return <span>{metadata.kind}: {metadata.instanceId}</span>;
 * }
 * ```
 */
export interface SPFxContextValue {
  /** Unique identifier for this SPFx instance */
  readonly instanceId: string;
  
  /**
   * Context reference of the supplied SPFx instance. Common fields are directly accessible; use a host-specific context hook for host-only fields.
   */
  readonly spfxContext: SPFxContextType;
  
  /** Type of host component */
  readonly kind: HostKind;
}
