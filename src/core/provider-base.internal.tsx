// provider-base.tsx
// SPFxProviderBase - Internal base implementation with shared logic

import * as React from 'react';
import { Provider, createStore, useSetAtom, useAtomValue } from 'jotai';
import { SPFxContext } from './context.internal';
import { spfxAtoms } from './atoms.internal';
import type { SPFxProviderProps, SPFxContextValue } from './types';
import {
  detectComponentKind,
  isWebPart,
} from '../utils/type-guards.internal';
import { useThemeSubscription } from '../utils/theme-subscription.internal';

/**
 * SPFxProviderBase - Internal base provider with shared logic
 * 
 * Creates an isolated Jotai store for each instance, ensuring complete
 * state isolation between multiple SPFx components on the same page.
 * 
 * The provider ensures SPFx ServiceScope is finished before rendering children,
 * guaranteeing that all services are available for consumption via dependency
 * injection in hooks. This prevents race conditions and ensures type-safe access
 * to SPFx services.
 * 
 * DO NOT use directly - use type-specific providers instead:
 * - SPFxWebPartProvider
 * - SPFxApplicationCustomizerProvider
 * - SPFxListViewCommandSetProvider
 * - SPFxFieldCustomizerProvider
 * 
 * Or use the generic SPFxProvider for backward compatibility.
 * 
 * @internal
 */
export function SPFxProviderBase<TProps extends {} = {}>(
  props: SPFxProviderProps<TProps>
): React.ReactElement {
  const { instance, children } = props;
  
  // Cast to 'any' to access protected/private properties
  // This is safe because we're inside the provider and know the structure
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const instanceAny = instance as any;
  
  // Detect component kind (WebPart, AppCustomizer, etc.)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const kind = React.useMemo(() => detectComponentKind(instance as any), [instance]);
  
  // Extract context (BaseComponentContext - has instanceId)
  const context = instanceAny.context;
  
  // Extract instanceId type-safe (all SPFx contexts extend BaseComponentContext)
  const instanceId = React.useMemo(() => context.instanceId, [context]);
  
  // Extract serviceScope and ensure it's finished before rendering
  const serviceScope = React.useMemo(() => context.serviceScope, [context]);
  const [isScopeReady, setIsScopeReady] = React.useState(false);
  
  // Wait for serviceScope to be finished before rendering children
  React.useEffect(() => {
    if (!serviceScope) {
      // Fallback: if no serviceScope, proceed (shouldn't happen in valid SPFx)
      setIsScopeReady(true);
      return;
    }
    
    // whenFinished callback fires immediately if already finished (synchronous)
    // or later when finished (asynchronous) - handles both scenarios
    serviceScope.whenFinished(() => {
      setIsScopeReady(true);
    });
  }, [serviceScope]);
  
  // Create isolated Jotai store for this Provider instance
  // Each store is independent, ensuring complete state isolation
  const store = React.useMemo(() => createStore(), []);
  
  // Get atom setters (using the isolated store)
  const setProperties = useSetAtom(spfxAtoms.properties, { store });
  const setDisplayMode = useSetAtom(spfxAtoms.displayMode, { store });
  const setContainerEl = useSetAtom(spfxAtoms.containerEl, { store });
  const setTheme = useSetAtom(spfxAtoms.theme, { store });
  
  // Ref to track last known properties value (prevents loop)
  const lastPropertiesRef = React.useRef<unknown>(instanceAny.properties);
  
  // Subscribe to theme changes (single subscription per instance)
  useThemeSubscription(context, setTheme);
  
  // Initialize atoms based on component type
  React.useEffect(() => {
    // Properties (common to all)
    setProperties(instanceAny.properties);
    lastPropertiesRef.current = instanceAny.properties;
    
    // WebPart-specific
    if (isWebPart(instance)) {
      setDisplayMode(instanceAny.displayMode);
      setContainerEl(instanceAny.domElement);
    }
  }, [
    instance,
    instanceAny,
    setProperties,
    setDisplayMode,
    setContainerEl,
  ]);
  
  // Sync properties when they change (SPFx → Atom)
  // Property Pane changes will trigger this via instance.properties reference change
  React.useEffect(() => {
    if (instanceAny.properties !== lastPropertiesRef.current) {
      setProperties(instanceAny.properties);
      lastPropertiesRef.current = instanceAny.properties;
    }
  }, [instanceAny.properties, setProperties, instanceAny]);
  
  // Sync properties when atom changes (Atom → SPFx)
  // Hook updates will trigger this via atom subscription
  const properties = useAtomValue(spfxAtoms.properties, { store });
  React.useEffect(() => {
    // Guard: Don't sync if atom is still undefined (initial state before initialization)
    // This prevents race condition where useAtomValue reads before initialization useEffect runs
    if (properties === undefined) {
      return;
    }
    
    // Only sync if atom value is different from last known value
    if (properties !== lastPropertiesRef.current) {
      // Mutate SPFx properties object (copy all properties from atom to instance)
      const target = instanceAny.properties as Record<string, unknown>;
      const source = properties as Record<string, unknown>;
      
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
      
      lastPropertiesRef.current = properties;
      
      // Refresh Property Pane for WebParts (if propertyPane exists)
      if (isWebPart(instance)) {
        const ctx = instanceAny.context as unknown as { propertyPane?: { refresh(): void } };
        if (ctx.propertyPane && typeof ctx.propertyPane.refresh === 'function') {
          ctx.propertyPane.refresh();
        }
      }
    }
  }, [properties, instance, instanceAny]);
  
  // WebPart: Sync displayMode when it changes
  React.useEffect(() => {
    if (isWebPart(instance)) {
      setDisplayMode(instanceAny.displayMode);
    }
  }, [instance, instanceAny, setDisplayMode]);
  
  // Create context value (memoized to prevent re-renders)
  const contextValue = React.useMemo<SPFxContextValue>(
    () => ({
      instanceId,
      spfxContext: context,
      kind,
    }),
    [instanceId, context, kind]
  );
  
  // Guard: Wait for serviceScope to be finished before rendering children
  // This ensures all hooks can safely consume services via serviceScope.consume()
  // If serviceScope is already finished, this guard passes immediately (no flash)
  if (!isScopeReady) {
    return React.createElement(React.Fragment, undefined);
  }
  
  return (
    <Provider store={store}>
      <SPFxContext.Provider value={contextValue}>
        {children}
      </SPFxContext.Provider>
    </Provider>
  );
}
