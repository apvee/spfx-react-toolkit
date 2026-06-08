// context.tsx
// React Context for SPFx metadata (static, non-reactive)

import * as React from 'react';
import type { SPFxContextValue } from './types';

/**
 * React Context for SPFx metadata
 * Contains only static reference data:
 * - instanceId: Unique identifier for this SPFx instance
 * - spfxContext: The SPFx context object (WebPartContext, etc.)
 * - kind: Type of host component
 * 
 * This context does NOT contain reactive state.
 * State is managed via Jotai atoms in isolated stores per Provider instance.
 * 
 * @internal
 */
export const SPFxContext = React.createContext<SPFxContextValue | null>(null);

if (process.env.NODE_ENV !== 'production') {
  SPFxContext.displayName = 'SPFxContext';
}

/**
 * Internal hook to access SPFx context
 *
 * Provides access to the SPFx context value containing instanceId, spfxContext, and kind.
 * Must be used within an SPFx provider component tree.
 *
 * @returns SPFxContextValue containing instanceId, spfxContext, and kind
 * @throws Error if used outside an SPFx provider - component must be wrapped with a host-specific provider.
 *
 * @internal
 */
export function useSPFxContext(): SPFxContextValue {
  const context = React.useContext(SPFxContext);
  
  if (!context) {
    throw new Error(
      'useSPFxContext must be used within an SPFx provider. ' +
      'Make sure your component is wrapped with a host-specific SPFx provider.'
    );
  }
  
  return context;
}
