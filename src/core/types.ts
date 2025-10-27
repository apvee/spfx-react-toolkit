// types.ts
// Core type definitions for SPFx React Toolkit

import type { DisplayMode } from '@microsoft/sp-core-library';

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
 * Structural types for SPFx components (no imports needed)
 * These match the actual SPFx base class structures
 */

/**
 * Minimal SPFx Context structure (common to all)
 */
export interface SPFxContextLike {
  readonly instanceId?: string;
  readonly pageContext?: unknown;
  readonly serviceScope?: unknown;
}

/**
 * WebPart-like structure (structural typing)
 */
export interface WebPartLike<TProps = unknown> {
  readonly context: SPFxContextLike & {
    readonly instanceId: string;
  };
  readonly properties: TProps;
  readonly displayMode: DisplayMode;
  readonly domElement: HTMLElement;
  render(): void;
}

/**
 * ApplicationCustomizer-like structure
 */
export interface ApplicationCustomizerLike<TProps = unknown> {
  readonly context: SPFxContextLike & {
    readonly instanceId: string;
    readonly placeholderProvider: {
      readonly changedEvent: unknown;
      tryCreateContent(name: unknown, options?: unknown): unknown;
    };
  };
  readonly properties: TProps;
}

/**
 * ListViewCommandSet-like structure
 */
export interface ListViewCommandSetLike<TProps = unknown> {
  readonly context: SPFxContextLike & {
    readonly instanceId: string;
    readonly listView: {
      readonly selectedRows?: ReadonlyArray<unknown>;
      readonly rows?: ReadonlyArray<unknown>;
      readonly list?: { title?: string };
      listViewStateChangedEvent?: unknown;
    };
  };
  readonly properties: TProps;
  onExecute(event: unknown): void;
  tryGetCommand(id: string): unknown;
}

/**
 * FieldCustomizer-like structure
 */
export interface FieldCustomizerLike<TProps = unknown> {
  readonly context: SPFxContextLike & {
    readonly instanceId: string;
    readonly field: {
      readonly listId: string;
      readonly internalName: string;
    };
    readonly itemId?: number;
  };
  readonly properties: TProps;
  onRenderCell(event: unknown): void;
}

/**
 * Union type for all SPFx component-like structures
 */
export type SPFxComponentLike<TProps = unknown> = 
  | WebPartLike<TProps>
  | ApplicationCustomizerLike<TProps>
  | ListViewCommandSetLike<TProps>
  | FieldCustomizerLike<TProps>;

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
export interface SPFxProviderProps<TProps = unknown> {
  /** SPFx component instance (WebPart, ApplicationCustomizer, etc.) */
  readonly instance: SPFxComponentLike<TProps>;
  
  /** Children to render */
  readonly children?: React.ReactNode;
}

/**
 * Context value provided by SPFxProvider
 * Contains only static metadata, no reactive state
 */
export interface SPFxContextValue<TContext extends SPFxContextLike = SPFxContextLike> {
  /** Unique identifier for this SPFx instance */
  readonly instanceId: string;
  
  /** SPFx context object */
  readonly spfxContext: TContext;
  
  /** Type of host component */
  readonly kind: HostKind;
}
