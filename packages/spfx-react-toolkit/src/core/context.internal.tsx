/**
 * SPFx metadata access.
 *
 * @packageDocumentation
 * @docPage hooks/context
 */
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
 * @example Read Context metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxContext } from '@apvee/spfx-react-toolkit';
 *
 * function HostLabel() {
 *   const { kind, instanceId } = useSPFxContext();
 *   return (
 *     <span>
 *       {kind}: {instanceId}
 *     </span>
 *   );
 * }
 * ```
 *
 * @useWhen a descendant needs the native host context or instance identity without subscribing to runtime state.
 *
 * @scenario Inspect the current native host identity
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxContext } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxContext();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxContext result">
 *         <dt>Host kind</dt><dd>{info.kind}</dd>
 *         <dt>Instance ID</dt><dd>{info.instanceId}</dd>
 *         <dt>Native context available</dt><dd>{String(info.spfxContext !== undefined)}</dd>
 *       </dl>,
 *     diagnostics: <p>Provider metadata identifies this native host; mutable nested fields require a React render.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual Context values and explicit unavailable fields as the native host changes.' };
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
