// type-guards.ts
// Type guards for SPFx component detection using structural typing (duck typing)

import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import type { BaseListViewCommandSet, BaseFieldCustomizer } from '@microsoft/sp-listview-extensibility';
import type {
  SPFxComponent,
  HostKind,
} from './types';

/**
 * Type guard: Check if instance is a WebPart
 * Uses duck typing - checks for WebPart-specific properties
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
 * Uses duck typing - checks for ApplicationCustomizer-specific properties
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
 * Uses duck typing - checks for CommandSet-specific properties
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
 * Uses duck typing - checks for FieldCustomizer-specific properties
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
 * Throws if unable to detect
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

/**
 * Derive instanceId from SPFx context
 * Fallback to generic ID if not found
 */
export function deriveInstanceId(context: unknown): string {
  if (!context || typeof context !== 'object') {
    return 'spfx-instance-unknown';
  }
  
  const ctx = context as Record<string, unknown>;
  
  if ('instanceId' in ctx && typeof ctx.instanceId === 'string') {
    return ctx.instanceId;
  }
  
  // Fallback
  return 'spfx-instance-fallback';
}
