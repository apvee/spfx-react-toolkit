/**
 * PageContext exposes metadata already supplied by the active SPFx host.
 *
 * @packageDocumentation
 * @docPage hooks/context
 */
// useSPFxPageContext.ts
// Hook to access SharePoint page context

import { useMemo } from 'react';
import { PageContext } from '@microsoft/sp-page-context';
import { useSPFxServiceScope } from './useSPFxServiceScope';

/**
 * Consumes the host PageContext service lazily from the provider's ready ServiceScope and memoizes it for that consume function. Returns the host reference; it does not clone context, fetch list data or independently subscribe to every field mutation. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns The PageContext service instance from the nearest provider scope.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example Read PageContext metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPageContext } from '@apvee/spfx-react-toolkit';
 *
 * function SiteLabel() {
 *   const context = useSPFxPageContext();
 *   return <span>{context.web.title}</span>;
 * }
 * ```
 *
 * @useWhen a descendant needs the native PageContext metadata instead of issuing a network request.
 *
 * @scenario Inspect selected native page metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPageContext } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxPageContext();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxPageContext result">
 *         <dt>Web title</dt><dd>{info.web.title}</dd>
 *         <dt>Web URL</dt><dd>{info.web.absoluteUrl}</dd>
 *         <dt>Site URL</dt><dd>{info.site.absoluteUrl}</dd>
 *         <dt>User</dt><dd>{info.user.displayName}</dd>
 *         <dt>Email</dt><dd>{info.user.email ?? 'Unavailable'}</dd>
 *         <dt>UI culture</dt><dd>{info.cultureInfo.currentUICultureName}</dd>
 *         <dt>List context</dt><dd>{info.list ? 'Available' : 'No list context'}</dd>
 *       </dl>,
 *     diagnostics: <p>This is the native PageContext service reference. It neither queries data nor independently subscribes to every mutable field.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual PageContext values and explicit unavailable fields as the native host changes.' };
 * ```
 */
export function useSPFxPageContext(): PageContext {
  const { consume } = useSPFxServiceScope();
  
  // Lazy consume PageContext from ServiceScope (cached by useMemo)
  // ServiceScope is guaranteed to be finished by the SPFx provider guard
  return useMemo(() => {
    return consume<PageContext>(PageContext.serviceKey);
  }, [consume]);
}
