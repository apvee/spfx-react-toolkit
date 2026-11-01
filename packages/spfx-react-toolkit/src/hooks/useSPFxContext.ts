/**
 * Static native context and identity are separate from reactive runtime state.
 *
 * @packageDocumentation
 * @docPage hooks/context
 */
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
 * @useWhen a descendant needs the native host context or instance identity without subscribing to runtime state.
 * @returns Static identity and native context from the nearest provider.
 * @throws Error if used outside an SPFx provider
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxContext } from '@apvee/spfx-react-toolkit';
 *
 * function HostLabel() {
 *   const { kind, instanceId } = useSPFxContext();
 *   return <span>{kind}: {instanceId}</span>;
 * }
 * ```
 *
 * @example Render host metadata and guarded host-specific content
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxContext, useSPFxInstanceInfo, useSPFxPageContext, useSPFxContainerInfo } from '@apvee/spfx-react-toolkit';
 * function HostStatus() {
 *   const context = useSPFxContext();
 *   const instance = useSPFxInstanceInfo();
 *   const page = useSPFxPageContext();
 *   const container = useSPFxContainerInfo();
 *   const host = context.spfxContext;
 *   const root = 'domElement' in host && host.domElement instanceof HTMLElement ? host.domElement : undefined;
 *   const label = context.kind === 'WebPart' ? 'Full web part' : context.kind === 'AppCustomizer' ? 'Compact header' : 'Default host';
 *   return <section><h2>{label}</h2><p>{context.instanceId}: {context.kind}</p>
 *     <p>{page.web.title}: {page.web.absoluteUrl}</p><p>{page.user.displayName} {page.user.email}</p>
 *     <p>{page.cultureInfo.currentUICultureName}</p><p>{root?.tagName ?? 'No web-part root'}</p>
 *     <p>{container.size ? `${container.size.width}x${container.size.height}px` : 'No container'}</p>
 *     <button onClick={() => console.log(`[${instance.kind}:${instance.id}] Button clicked`)}>Log instance</button>
 *   </section>;
 * }
 * ```
 */
export { useSPFxContext };
