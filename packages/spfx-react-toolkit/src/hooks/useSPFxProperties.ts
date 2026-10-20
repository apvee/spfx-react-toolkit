// useSPFxProperties.ts
// Hook to access and manage SPFx properties

import { useCallback, useMemo } from 'react';
import { useSPFxRuntimeActions, useSPFxRuntimeSelector } from './../core/state.internal';

/**
 * Return type for useSPFxProperties hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const props: SPFxPropertiesInfo<{ title: string }> = useSPFxProperties();
 * const onRename = () => props.setProperties({ title: 'Projects' });
 * ```
 * @template TProps - Application properties shape; no runtime validation is performed.
 */
export interface SPFxPropertiesInfo<TProps = unknown> {
  /** Current properties object */
  readonly properties: TProps | undefined;
  
  /**
   * Shallow-merges updates into current provider properties. Provider host synchronization determines persistence/refresh behavior.
   * @param updates - Partial properties to merge.
   * @returns No value.
   */
  readonly setProperties: (updates: Partial<TProps>) => void;
  
  /**
   * Replaces properties with the updater result; missing current properties are passed as undefined.
   * @param updater - Pure updater returning the complete next property object.
   * @returns No value.
   * @throws Updater errors propagate.
   */
  readonly updateProperties: (updater: (current: TProps | undefined) => TProps) => void;
}

/**
 * Reads and updates the nearest provider properties. setProperties shallow-merges partial updates; updateProperties replaces the complete object and receives undefined before properties are available. The provider synchronizes host properties and refreshes WebPart property panes. This hook does not validate the TProps shape; render under a matching initialized provider using that host property type.
 *
 * @typeParam TProps - Host properties shape; no runtime validation is performed.
 * @returns Current optional properties and partial-merge/complete-replacement setters.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example
 * ```tsx
 * interface MyProperties { title: string; description: string; }
 *
 * function RenameButton() {
 *   const { properties, updateProperties } = useSPFxProperties<MyProperties>();
 *   return <button onClick={() => updateProperties(previous => ({
 *     ...previous,
 *     title: `${previous?.title ?? ''} Updated`,
 *     description: previous?.description ?? ''
 *   }))}>{properties?.title ?? 'Rename'}</button>;
 * }
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
