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
 * @example
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
 */
export function useSPFxPageContext(): PageContext {
  const { consume } = useSPFxServiceScope();
  
  // Lazy consume PageContext from ServiceScope (cached by useMemo)
  // ServiceScope is guaranteed to be finished by the SPFx provider guard
  return useMemo(() => {
    return consume<PageContext>(PageContext.serviceKey);
  }, [consume]);
}
