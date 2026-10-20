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
 * State is managed via isolated runtime stores per Provider instance.
 * 
 * @internal
 */
export const SPFxContext = React.createContext<SPFxContextValue | null>(null);

if (process.env.NODE_ENV !== 'production') {
  SPFxContext.displayName = 'SPFxContext';
}

/**
 * Returns static metadata for the nearest host-specific SPFx provider. Call inside a React component or hook; theme, size and property changes are exposed through dedicated runtime hooks. Although its implementation lives in an internal module, this hook is exported from the public hooks entry.
 *
 * @returns Nearest provider metadata, including kind, instanceId and the SPFx context reference.
 * @throws Error when called outside a host-specific SPFx provider.
 *
 * @example
 * ```tsx
 * function HostLabel() {
 *   const { kind, instanceId } = useSPFxContext();
 *   return <span>{kind}: {instanceId}</span>;
 * }
 * ```
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
