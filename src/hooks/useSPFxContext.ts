// useSPFxContext.ts
// Internal hook to access SPFx context metadata

import { useSPFxContext } from '../core/context.internal';

/**
 * Internal hook to access SPFx context
 * 
 * Returns:
 * - instanceId: Unique identifier for this SPFx instance
 * - spfxContext: The SPFx context object (WebPartContext, etc.)
 * - kind: Type of host component ('WebPart', 'AppCustomizer', etc.)
 * 
 * @throws Error if used outside SPFxProvider
 * @internal - Not exported in public API
 */
export { useSPFxContext };
