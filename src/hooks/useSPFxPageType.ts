// useSPFxPageType.ts
// Hook for SharePoint page type detection

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
  
  // Try to get page type from modern pageContext
  const modernPage = (pageContext as unknown as {
    page?: {
      type?: string;
    };
  }).page;
  
  // Try to get page type from legacy context
  const legacy = (pageContext as unknown as {
    legacyPageContext?: {
      pageType?: string;
      listId?: string;
      formType?: string | number;
    };
  }).legacyPageContext;
  
  // Determine page type
  let pageType: SPFxPageType = 'unknown';
  
  // Check modern page type first
  const modernPageType = modernPage?.type?.toLowerCase();
  if (modernPageType) {
    if (modernPageType.indexOf('sitepage') !== -1) {
      pageType = 'sitePage';
    } else if (modernPageType.indexOf('webpartpage') !== -1) {
      pageType = 'webPartPage';
    }
  }
  
  // Check legacy page type
  if (pageType === 'unknown') {
    const legacyPageType = legacy?.pageType?.toLowerCase();
    if (legacyPageType) {
      if (legacyPageType.indexOf('sitepage') !== -1) {
        pageType = 'sitePage';
      } else if (legacyPageType.indexOf('webpartpage') !== -1) {
        pageType = 'webPartPage';
      } else if (legacyPageType.indexOf('list') !== -1) {
        // Check if it's a form or list view
        if (legacy?.formType !== undefined && legacy.formType !== null) {
          pageType = 'listFormPage';
        } else {
          pageType = 'listPage';
        }
      } else if (legacyPageType.indexOf('profile') !== -1) {
        pageType = 'profilePage';
      } else if (legacyPageType.indexOf('search') !== -1) {
        pageType = 'searchPage';
      }
    }
  }
  
  // If still unknown, try to infer from context
  if (pageType === 'unknown') {
    // If we have a list ID, likely a list page
    if (legacy?.listId) {
      if (legacy.formType !== undefined && legacy.formType !== null) {
        pageType = 'listFormPage';
      } else {
        pageType = 'listPage';
      }
    }
  }
  
  // Calculate helper flags
  const isSitePage = pageType === 'sitePage';
  const isWebPartPage = pageType === 'webPartPage';
  const isListPage = pageType === 'listPage';
  const isListFormPage = pageType === 'listFormPage';
  
  // Modern page = site page (not classic web part page)
  const isModernPage = isSitePage;
  
  return {
    pageType,
    isModernPage,
    isSitePage,
    isListPage,
    isListFormPage,
    isWebPartPage,
  };
}
