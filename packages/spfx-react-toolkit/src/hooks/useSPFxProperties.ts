/**
 * Host property edits synchronize through the nearest provider runtime.
 *
 * @packageDocumentation
 * @docPage hooks/properties
 */
// useSPFxProperties.ts
// Hook to access and manage SPFx properties

import { useCallback, useMemo } from 'react';
import { useSPFxRuntimeActions, useSPFxRuntimeSelector } from './../core/state.internal';

/**
 * Nearest provider property bag and shallow-merge/complete-replacement update operations.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxPropertiesInfo, useSPFxProperties } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const props: SPFxPropertiesInfo<{ title: string }> = useSPFxProperties();
 *   const onRename = () => props.setProperties({ title: 'Projects' });
 *   return <button onClick={()=>void onRename()}>Run operation</button>;
 * }
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
  readonly updateProperties: (updater: /**
   * Derives the complete next host property bag.
   * @param current - Current properties, or undefined before provider initialization.
   * @returns Complete next property bag; it replaces the previous object.
   */ (current: TProps | undefined) => TProps) => void;
}

/**
 * Reads and updates the nearest provider properties. setProperties shallow-merges partial updates; updateProperties replaces the complete object and receives undefined before properties are available. The provider synchronizes host properties and refreshes WebPart property panes. This hook does not validate the TProps shape; render under a matching initialized provider using that host property type.
 *
 * @typeParam TProps - Host properties shape; no runtime validation is performed.
 * @returns Current optional properties and partial-merge/complete-replacement setters.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxProperties } from '@apvee/spfx-react-toolkit';
 *
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
 *
 * @useWhen React controls must edit the host property bag and reflect host render updates.
 *
 *
 * @example Update selected properties
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxProperties } from '@apvee/spfx-react-toolkit';
 *
 * interface IMyWebPartProps {
 *   title: string;
 *   description: string;
 *   showHeader: boolean;
 * }
 *
 * function MyComponent() {
 *   const { properties, setProperties } = useSPFxProperties<IMyWebPartProps>();
 *
 *   return (
 *     <div>
 *       <h1>{properties?.title ?? 'Default Title'}</h1>
 *       <p>{properties?.description}</p>
 *
 *       <button onClick={() => setProperties({ title: 'New Title' })}>
 *         Update Title
 *       </button>
 *     </div>
 *   );
 * }
 * ```
 *
 *
 * @example Update properties from the previous snapshot
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxProperties } from '@apvee/spfx-react-toolkit';
 *
 * interface ICounterProps {
 *   count: number;
 *   lastUpdated: string;
 * }
 *
 * function CounterComponent() {
 *   const { properties, updateProperties } = useSPFxProperties<ICounterProps>();
 *
 *   const increment = () => {
 *     updateProperties(prev => ({
 *       count: (prev?.count ?? 0) + 1,
 *       lastUpdated: new Date().toISOString()
 *     }));
 *   };
 *
 *   return (
 *     <div>
 *       <p>Count: {properties?.count ?? 0}</p>
 *       <p>Last Updated: {properties?.lastUpdated ?? 'Never'}</p>
 *       <button onClick={increment}>Increment</button>
 *     </div>
 *   );
 * }
 * ```
 *
 *
 * @example Edit configuration in native edit mode
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxProperties, useSPFxDisplayMode } from '@apvee/spfx-react-toolkit';
 *
 * interface IFormProps {
 *   formTitle: string;
 *   submitUrl: string;
 *   showLabels: boolean;
 * }
 *
 * function ConfigurableForm() {
 *   const { properties, setProperties } = useSPFxProperties<IFormProps>();
 *   const { isEdit } = useSPFxDisplayMode();
 *
 *   if (isEdit) {
 *     return (
 *       <div className="config-panel">
 *         <label>
 *           Form Title:
 *           <input
 *             value={properties?.formTitle ?? ''}
 *             onChange={(e) => setProperties({ formTitle: e.target.value })}
 *           />
 *         </label>
 *         <label>
 *           <input
 *             type="checkbox"
 *             checked={properties?.showLabels ?? true}
 *             onChange={(e) => setProperties({ showLabels: e.target.checked })}
 *           />
 *           Show Labels
 *         </label>
 *       </div>
 *     );
 *   }
 *
 *   return (
 *     <form>
 *       <h2>{properties?.formTitle}</h2>
 *
 *     </form>
 *   );
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
  const updateProperties = useCallback((updater: /**
   * Derives the complete next host property bag.
   * @param current - Current properties, or undefined before provider initialization.
   * @returns Complete next property bag; it replaces the previous object.
   */ (current: TProps | undefined) => TProps): void => {
    setRuntimeProperties((prev: unknown) => updater(prev as TProps | undefined));
  }, [setRuntimeProperties]);
  
  return useMemo(() => ({
    properties,
    setProperties,
    updateProperties,
  }), [properties, setProperties, updateProperties]);
}
