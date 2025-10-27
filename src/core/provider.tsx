// provider.tsx
// SPFxProvider - Instance-based, auto-detecting, pragmatic approach

import * as React from 'react';
import { useSetAtom, useStore } from 'jotai';
import { SPFxContext } from './context';
import { spfxAtoms } from './atoms';
import type { SPFxProviderProps, SPFxContextValue } from './types';
import {
  detectComponentKind,
  deriveInstanceId,
  isWebPart,
} from './type-guards';
import { useThemeSubscription } from '../utils/theme-subscription';

/**
 * SPFxProvider - Instance-based provider for SPFx React Toolkit
 * 
 * Automatically detects component type and extracts all necessary properties.
 * No manual prop passing needed - just pass `this` from your SPFx component!
 * 
 * Features:
 * - Auto-detects: WebPart, ApplicationCustomizer, CommandSet, FieldCustomizer
 * - Extracts: context, properties, displayMode, containerEl, etc.
 * - Manages: Jotai atoms initialization and cleanup
 * - Type-safe: Full TypeScript support with generics
 * 
 * @example
 * ```tsx
 * // In your WebPart:
 * export default class MyWebPart extends BaseClientSideWebPart<IMyProps> {
 *   public render(): void {
 *     const element = React.createElement(
 *       SPFxProvider,
 *       { instance: this },
 *       React.createElement(MyApp)
 *     );
 *     ReactDom.render(element, this.domElement);
 *   }
 * }
 * ```
 */
export function SPFxProvider<TProps = unknown>(
  props: SPFxProviderProps<TProps>
): JSX.Element {
  const { instance, children } = props;
  
  // Detect component kind (WebPart, AppCustomizer, etc.)
  const kind = React.useMemo(() => detectComponentKind(instance), [instance]);
  
  // Extract context and instanceId
  const context = instance.context;
  const instanceId = React.useMemo(
    () => deriveInstanceId(context),
    [context]
  );
  
  // Get Jotai store for subscription
  const store = useStore();
  
  // Get atom setters for this instance
  const setProperties = useSetAtom(spfxAtoms.properties(instanceId));
  const setDisplayMode = useSetAtom(spfxAtoms.displayMode(instanceId));
  const setContainerEl = useSetAtom(spfxAtoms.containerEl(instanceId));
  const setTheme = useSetAtom(spfxAtoms.theme(instanceId));
  
  // Ref to track last known properties value (prevents loop)
  const lastPropertiesRef = React.useRef<unknown>(instance.properties);
  
  // Subscribe to theme changes (single subscription per instance)
  useThemeSubscription(context, setTheme);
  
  // Initialize atoms based on component type
  React.useEffect(() => {
    // Properties (common to all)
    setProperties(instance.properties);
    lastPropertiesRef.current = instance.properties;
    
    // WebPart-specific
    if (isWebPart(instance)) {
      setDisplayMode(instance.displayMode);
      setContainerEl(instance.domElement);
    }
  }, [
    instance,
    setProperties,
    setDisplayMode,
    setContainerEl,
  ]);
  
  // Sync properties when they change (SPFx → Atom)
  // Property Pane changes will trigger this via instance.properties reference change
  React.useEffect(() => {
    if (instance.properties !== lastPropertiesRef.current) {
      setProperties(instance.properties);
      lastPropertiesRef.current = instance.properties;
    }
  }, [instance.properties, setProperties]);
  
  // Sync properties when atom changes (Atom → SPFx)
  // Hook updates will trigger this via atom subscription
  React.useEffect(() => {
    const propertiesAtom = spfxAtoms.properties(instanceId);
    
    const unsubscribe = store.sub(propertiesAtom, () => {
      const atomValue = store.get(propertiesAtom);
      
      // Only sync if atom value is different from last known value
      if (atomValue !== lastPropertiesRef.current) {
        // Mutate SPFx properties object (copy all properties from atom to instance)
        const target = instance.properties as Record<string, unknown>;
        const source = atomValue as Record<string, unknown>;
        
        // Clear existing properties
        for (const key in target) {
          if (Object.prototype.hasOwnProperty.call(target, key)) {
            delete target[key];
          }
        }
        
        // Copy new properties
        for (const key in source) {
          if (Object.prototype.hasOwnProperty.call(source, key)) {
            target[key] = source[key];
          }
        }
        
        lastPropertiesRef.current = atomValue;
        
        // Refresh Property Pane for WebParts (if propertyPane exists)
        if (isWebPart(instance)) {
          const ctx = instance.context as unknown as { propertyPane?: { refresh(): void } };
          if (ctx.propertyPane && typeof ctx.propertyPane.refresh === 'function') {
            ctx.propertyPane.refresh();
          }
        }
      }
    });
    
    return unsubscribe;
  }, [store, instanceId, instance]);
  
  // WebPart: Sync displayMode when it changes
  React.useEffect(() => {
    if (isWebPart(instance)) {
      setDisplayMode(instance.displayMode);
    }
  }, [instance, setDisplayMode]);
  
  // Cleanup atoms when component unmounts (memory leak prevention)
  React.useEffect(() => {
    return () => {
      const families = [
        spfxAtoms.theme,
        spfxAtoms.displayMode,
        spfxAtoms.properties,
        spfxAtoms.containerEl,
        spfxAtoms.containerSize,
        spfxAtoms.teams,
        spfxAtoms.dynamicData,
      ];
      
      families.forEach(family => {
        family.remove(instanceId);
      });
    };
  }, [instanceId]);
  
  // Create context value (memoized to prevent re-renders)
  const contextValue = React.useMemo<SPFxContextValue>(
    () => ({
      instanceId,
      spfxContext: context,
      kind,
    }),
    [instanceId, context, kind]
  );
  
  return (
    <SPFxContext.Provider value={contextValue}>
      {children}
    </SPFxContext.Provider>
  );
}
