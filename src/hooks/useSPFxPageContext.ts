// useSPFxPageContext.ts
// Hook to access SharePoint page context

import { useMemo } from 'react';
import { PageContext } from '@microsoft/sp-page-context';
import { useSPFxServiceScope } from './useSPFxServiceScope';

/**
 * Hook to access SharePoint page context
 * 
 * Provides access to SharePoint page context containing information about:
 * - Site collection and web
 * - Current user
 * - Current list and list item (if applicable)
 * - Teams context (if running in Teams)
 * - Culture and locale settings
 * - Permissions and capabilities
 * 
 * @returns SharePoint page context object
 * 
 * @remarks
 * This hook consumes PageContext from SPFx ServiceScope using dependency injection.
 * The service is consumed lazily (only when this hook is used) and cached for performance.
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const pageContext = useSPFxPageContext();
 *   
 *   return (
 *     <div>
 *       <p>Site: {pageContext.web.title}</p>
 *       <p>User: {pageContext.user.displayName}</p>
 *       <p>Locale: {pageContext.cultureInfo.currentUICultureName}</p>
 *     </div>
 *   );
 * }
 * ```
 * 
 * @see {@link useSPFxContext} for accessing the full SPFx context
 */
export function useSPFxPageContext(): PageContext {
  const { consume } = useSPFxServiceScope();
  
  // Lazy consume PageContext from ServiceScope (cached by useMemo)
  // ServiceScope is guaranteed to be finished by SPFxProvider guard
  return useMemo(() => {
    return consume<PageContext>(PageContext.serviceKey);
  }, [consume]);
}
