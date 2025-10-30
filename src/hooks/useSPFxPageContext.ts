// useSPFxPageContext.ts
// Hook to access SharePoint page context

import type { PageContext } from '@microsoft/sp-page-context';
import { useSPFxContext } from './useSPFxContext';

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
 * This hook extracts and returns the `pageContext` property from the SPFx context.
 * If you need access to the full SPFx context object, use `useSPFxContext` instead.
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
  const { spfxContext } = useSPFxContext();
  
  // Extract pageContext from SPFx context
  // All SPFx contexts have pageContext property
  const ctx = spfxContext as { pageContext?: PageContext };
  
  if (!ctx.pageContext) {
    throw new Error(
      'SPFx context does not contain pageContext. ' +
      'This should never happen with valid SPFx contexts.'
    );
  }
  
  return ctx.pageContext;
}
