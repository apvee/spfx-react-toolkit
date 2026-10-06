// useSPFxPageType.ts
// Hook for SharePoint page type detection

import { useMemo } from 'react';
import { getSPFxPageTypeInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * SharePoint page types
 */
export type SPFxPageType = 
  | 'sitePage'          // Modern site page
  | 'webPartPage'       // Classic web part page
  | 'listPage'          // List view page
  | 'listFormPage'      // List form page (new/edit/display)
  | 'profilePage'       // User profile page
  | 'searchPage'        // Search results page
  | 'unknown';          // Unknown page type

/**
 * Return type for useSPFxPageType hook
 */
export interface SPFxPageTypeInfo {
  /** Current page type */
  readonly pageType: SPFxPageType;
  
  /** Whether the page is a modern site page */
  readonly isModernPage: boolean;
  
  /** Whether the page is a site page (modern) */
  readonly isSitePage: boolean;
  
  /** Whether the page is a list page (list view) */
  readonly isListPage: boolean;
  
  /** Whether the page is a list form page */
  readonly isListFormPage: boolean;
  
  /** Whether the page is a classic web part page */
  readonly isWebPartPage: boolean;
}

/**
 * Hook for SharePoint page type detection
 * 
 * Detects the current SharePoint page type:
 * - sitePage: Modern site page (Site Pages library)
 * - webPartPage: Classic web part page
 * - listPage: List view page
 * - listFormPage: List form (new/edit/display item)
 * - profilePage: User profile page
 * - searchPage: Search results page
 * - unknown: Unable to determine
 * 
 * Helper flags provided:
 * - isModernPage: True for modern site pages
 * - isSitePage: True for site pages
 * - isListPage: True for list views
 * - isListFormPage: True for list forms
 * - isWebPartPage: True for classic web part pages
 * 
 * Use this hook for:
 * - Conditional rendering based on page type
 * - Feature availability checks (e.g., modern-only features)
 * - Page-specific behavior
 * - Analytics/telemetry
 * 
 * @returns Page type information
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { pageType, isModernPage, isSitePage } = useSPFxPageType();
 *   
 *   if (!isModernPage) {
 *     return <div>This feature requires a modern page</div>;
 *   }
 *   
 *   return (
 *     <div>
 *       <h3>Page Type: {pageType}</h3>
 *       {isSitePage && <ModernPageFeature />}
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Conditional features
 * ```tsx
 * function ConditionalUI() {
 *   const { isListPage, isListFormPage } = useSPFxPageType();
 *   
 *   return (
 *     <div>
 *       {isListPage && <ListViewCustomizer />}
 *       {isListFormPage && <FormFieldCustomizer />}
 *     </div>
 *   );
 * }
 * ```
 */
export function useSPFxPageType(): SPFxPageTypeInfo {
  const pageContext = useSPFxPageContext();
  const modernPage = (pageContext as unknown as {
    page?: {
      type?: string;
    };
  }).page;
  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      pageType?: string;
      listId?: string;
      formType?: string | number;
    };
  }).legacyPageContext;
  
  return useMemo(
    () => getSPFxPageTypeInfo(pageContext),
    [pageContext, modernPage?.type, legacy?.pageType, legacy?.listId, legacy?.formType]
  );
}
