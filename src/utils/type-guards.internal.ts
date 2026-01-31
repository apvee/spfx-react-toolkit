// type-guards.ts
// Type guards for SPFx component detection using structural typing (duck typing)

import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import type { BaseListViewCommandSet, BaseFieldCustomizer } from '@microsoft/sp-listview-extensibility';
import type {
  SPFxComponent,
  HostKind,
} from '../core/types';

/**
 * Type guard: Check if instance is a WebPart
 *
 * Uses duck typing (structural typing) to detect WebPart instances
 * by checking for WebPart-specific properties: displayMode, domElement, render, context, properties.
 *
 * @template TProps - WebPart properties type
 * @param instance - SPFx component instance to check
 * @returns True if instance is a BaseClientSideWebPart, narrowing the type
 *
 * @example
 * ```typescript
 * if (isWebPart(instance)) {
 *   // instance is now typed as BaseClientSideWebPart<TProps>
 *   console.log(instance.displayMode);
 * }
 * ```
 *
 * @internal
 */
export function isWebPart<TProps extends {} = {}>(
  instance: unknown
): instance is BaseClientSideWebPart<TProps> {
  if (!instance || typeof instance !== 'object') return false;
  
  const obj = instance as Record<string, unknown>;
  
  return (
    'displayMode' in obj &&
    'domElement' in obj &&
    'render' in obj &&
    typeof obj.render === 'function' &&
    'context' in obj &&
    'properties' in obj
  );
}

/**
 * Type guard: Check if instance is an ApplicationCustomizer
 *
 * Uses duck typing (structural typing) to detect ApplicationCustomizer instances
 * by checking for placeholderProvider in context and absence of displayMode.
 *
 * @template TProps - ApplicationCustomizer properties type
 * @param instance - SPFx component instance to check
 * @returns True if instance is a BaseApplicationCustomizer, narrowing the type
 *
 * @example
 * ```typescript
 * if (isApplicationCustomizer(instance)) {
 *   // instance is now typed as BaseApplicationCustomizer<TProps>
 *   console.log(instance.context.placeholderProvider);
 * }
 * ```
 *
 * @internal
 */
export function isApplicationCustomizer<TProps extends {} = {}>(
  instance: unknown
): instance is BaseApplicationCustomizer<TProps> {
  if (!instance || typeof instance !== 'object') return false;
  
  const obj = instance as Record<string, unknown>;
  
  if (!('context' in obj) || !obj.context || typeof obj.context !== 'object') {
    return false;
  }
  
  const context = obj.context as Record<string, unknown>;
  
  return (
    'placeholderProvider' in context &&
    'properties' in obj &&
    !('displayMode' in obj) // Not a WebPart
  );
}

/**
 * Type guard: Check if instance is a ListViewCommandSet
 *
 * Uses duck typing (structural typing) to detect ListViewCommandSet instances
 * by checking for onExecute and tryGetCommand methods.
 *
 * @template TProps - ListViewCommandSet properties type
 * @param instance - SPFx component instance to check
 * @returns True if instance is a BaseListViewCommandSet, narrowing the type
 *
 * @example
 * ```typescript
 * if (isListViewCommandSet(instance)) {
 *   // instance is now typed as BaseListViewCommandSet<TProps>
 *   instance.tryGetCommand('COMMAND_ID');
 * }
 * ```
 *
 * @internal
 */
export function isListViewCommandSet<TProps extends {} = {}>(
  instance: unknown
): instance is BaseListViewCommandSet<TProps> {
  if (!instance || typeof instance !== 'object') return false;
  
  const obj = instance as Record<string, unknown>;
  
  return (
    'onExecute' in obj &&
    typeof obj.onExecute === 'function' &&
    'tryGetCommand' in obj &&
    typeof obj.tryGetCommand === 'function' &&
    'context' in obj &&
    'properties' in obj
  );
}

/**
 * Type guard: Check if instance is a FieldCustomizer
 *
 * Uses duck typing (structural typing) to detect FieldCustomizer instances
 * by checking for field in context and onRenderCell method.
 *
 * @template TProps - FieldCustomizer properties type
 * @param instance - SPFx component instance to check
 * @returns True if instance is a BaseFieldCustomizer, narrowing the type
 *
 * @example
 * ```typescript
 * if (isFieldCustomizer(instance)) {
 *   // instance is now typed as BaseFieldCustomizer<TProps>
 *   console.log(instance.context.field);
 * }
 * ```
 *
 * @internal
 */
export function isFieldCustomizer<TProps extends {} = {}>(
  instance: unknown
): instance is BaseFieldCustomizer<TProps> {
  if (!instance || typeof instance !== 'object') return false;
  
  const obj = instance as Record<string, unknown>;
  
  if (!('context' in obj) || !obj.context || typeof obj.context !== 'object') {
    return false;
  }
  
  const context = obj.context as Record<string, unknown>;
  
  return (
    'field' in context &&
    'onRenderCell' in obj &&
    typeof obj.onRenderCell === 'function' &&
    'properties' in obj
  );
}

/**
 * Detect the kind of SPFx component from an instance
 *
 * Checks the instance against all known SPFx component types and returns
 * the corresponding HostKind discriminator.
 *
 * @template TProps - SPFx component properties type
 * @param instance - SPFx component instance to detect
 * @returns HostKind ('WebPart' | 'AppCustomizer' | 'CommandSet' | 'FieldCustomizer')
 * @throws Error if unable to detect component type
 *
 * @example
 * ```typescript
 * const kind = detectComponentKind(this); // 'WebPart'
 * ```
 *
 * @internal
 */
export function detectComponentKind<TProps extends {} = {}>(
  instance: SPFxComponent<TProps>
): HostKind {
  if (isWebPart(instance)) return 'WebPart';
  if (isApplicationCustomizer(instance)) return 'AppCustomizer';
  if (isListViewCommandSet(instance)) return 'CommandSet';
  if (isFieldCustomizer(instance)) return 'FieldCustomizer';
  
  throw new Error(
    '[SPFxProvider] Unable to detect SPFx component type. ' +
    'Instance must be a WebPart, ApplicationCustomizer, CommandSet, or FieldCustomizer.'
  );
}
