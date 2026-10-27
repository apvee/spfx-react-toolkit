// useSPFxContext.ts
// Public hook to access SPFx context metadata

import { useSPFxContext } from '../core/context.internal';

/**
 * Returns static metadata from the nearest host-specific SPFx provider.
 * Call inside a React component or hook. Use dedicated runtime hooks for
 * reactive theme, property, display-mode and container-size state.
 * 
 * Returns:
 * - instanceId: Unique identifier for this SPFx instance
 * - spfxContext: The SPFx context object (WebPartContext, etc.)
 * - kind: Type of host component ('WebPart', 'AppCustomizer', etc.)
 * 
 * @throws Error if used outside an SPFx provider
 * @example
 * ```tsx
 * function HostLabel() {
 *   const { kind, instanceId } = useSPFxContext();
 *   return <span>{kind}: {instanceId}</span>;
 * }
 * ```
 */
export { useSPFxContext };
