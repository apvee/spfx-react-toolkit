// useSPFxProperties.ts
// Hook to access and manage SPFx properties

import { useCallback, useMemo } from 'react';
import { useSPFxRuntimeActions, useSPFxRuntimeSelector } from './../core/state.internal';

/**
 * Return type for useSPFxProperties hook
 */
export interface SPFxPropertiesInfo<TProps = unknown> {
  /** Current properties object */
  readonly properties: TProps | undefined;
  
  /** 
   * Update properties with partial updates (shallow merge)
   * Changes are automatically synced back to SPFx by the Provider
   */
  readonly setProperties: (updates: Partial<TProps>) => void;
  
  /**
   * Update properties using updater function
   * Useful for complex updates based on current state
   * Changes are automatically synced back to SPFx by the Provider
   */
  readonly updateProperties: (updater: (current: TProps | undefined) => TProps) => void;
}

/**
 * Hook to access and manage SPFx properties
 * 
 * Properties are the configuration values for WebParts/Extensions that:
 * - Are set via Property Pane
 * - Persist across page loads
 * - Are specific to each instance
 * 
 * This hook provides:
 * - Type-safe access to properties
 * - Partial updates (merge with existing)
 * - Updater function pattern (like React setState)
 * - Automatic bidirectional sync with SPFx (managed by Provider)
 * 
 * The SPFx provider automatically handles synchronization:
 * - Property Pane changes → runtime state → Hook (automatic)
 * - Hook updates → runtime state → SPFx properties (automatic)
 * - Property Pane refresh for WebParts (automatic)
 * 
 * @returns Properties and setter functions
 * 
 * @example
 * ```tsx
 * interface IMyWebPartProps {
 *   title: string;
 *   description: string;
 *   listId?: string;
 * }
 * 
 * function MyComponent() {
 *   const { properties, setProperties, updateProperties } = useSPFxProperties<IMyWebPartProps>();
 *   
 *   return (
 *     <div>
 *       <h1>{properties?.title ?? 'Default Title'}</h1>
 *       <p>{properties?.description}</p>
 *       
 *       <button onClick={() => setProperties({ title: 'New Title' })}>
 *         Update Title
 *       </button>
 *       
 *       <button onClick={() => updateProperties(prev => ({
 *         ...prev,
 *         title: (prev?.title ?? '') + ' Updated'
 *       }))}>
 *         Append to Title
 *       </button>
 *     </div>
 *   );
 * }
 * 
 * // In WebPart render():
 * // Just pass the instance - sync is automatic!
 * const element = React.createElement(
 *   SPFxWebPartProvider,
 *   { instance: this },
 *   React.createElement(MyComponent)
 * );
 * ```
 */
export function useSPFxProperties<TProps = unknown>(): SPFxPropertiesInfo<TProps> {
  const properties = useSPFxRuntimeSelector(state => state.properties) as TProps | undefined;
  const { setProperties: setRuntimeProperties } = useSPFxRuntimeActions();
  
  // Setter with partial merge (functional update for stable dependencies)
  const setProperties = useCallback((updates: Partial<TProps>): void => {
    setRuntimeProperties((prev: unknown) => ({
      ...(prev ?? {} as TProps),
      ...updates,
    }));
  }, [setRuntimeProperties]);
  
  // Updater function pattern (like React setState)
  const updateProperties = useCallback((updater: (current: TProps | undefined) => TProps): void => {
    setRuntimeProperties((prev: unknown) => updater(prev as TProps | undefined));
  }, [setRuntimeProperties]);
  
  return useMemo(() => ({
    properties,
    setProperties,
    updateProperties,
  }), [properties, setProperties, updateProperties]);
}
