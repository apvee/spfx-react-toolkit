// useSPFxInstanceInfo.ts
// Hook to access SPFx instance metadata

import { useSPFxContext } from './useSPFxContext';
import type { HostKind } from '../core/types';

/**
 * Return type for useSPFxInstanceInfo hook
 */
export interface SPFxInstanceInfo {
  /** Unique identifier for this SPFx instance */
  readonly id: string;
  
  /** Type of SPFx component (WebPart, AppCustomizer, etc.) */
  readonly kind: HostKind;
}

/**
 * Hook to access SPFx instance metadata
 * 
 * Provides:
 * - id: Unique identifier for this SPFx instance
 * - kind: Type of component ('WebPart', 'AppCustomizer', etc.)
 * 
 * Useful for:
 * - Logging and telemetry
 * - Conditional logic based on host type
 * - Scoped storage keys
 * - Debug information
 * 
 * @returns Instance metadata
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { id, kind } = useSPFxInstanceInfo();
 *   
 *   return (
 *     <div>
 *       <p>Instance: {id}</p>
 *       <p>Type: {kind}</p>
 *     </div>
 *   );
 * }
 * ```
 */
export function useSPFxInstanceInfo(): SPFxInstanceInfo {
  const { instanceId, kind } = useSPFxContext();
  
  return {
    id: instanceId,
    kind,
  };
}
