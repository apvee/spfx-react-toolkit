/**
 * Page categories are best-effort interpretations of host metadata.
 *
 * @packageDocumentation
 * @docPage hooks/environment
 */
// useSPFxPageType.ts
// Hook for SharePoint page type detection

import { useMemo } from 'react';
import { getSPFxPageTypeInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * SharePoint page types
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxPageType } from '@apvee/spfx-react-toolkit';
 *
 * const type: SPFxPageType = 'sitePage';
 * ```
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
 * Best-effort category and derived flags from modern or legacy host page metadata.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxPageTypeInfo, useSPFxPageType } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const page: SPFxPageTypeInfo = useSPFxPageType();
 *   console.log(page.pageType, page.isModernPage);
 *   return <span>{page.pageType}</span>;
 * }
 * ```
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
 * Classifies supplied modern page-type text, then legacy page type/list/form metadata. Missing or unrecognized metadata yields unknown. isModernPage is derived from the sitePage category and is not independent proof of page capabilities; no network detection occurs. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Best-effort page category and category-specific flags.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxPageType } from '@apvee/spfx-react-toolkit';
 *
 * function PageLabel() {
 *   const { pageType } = useSPFxPageType();
 *   return <span>{pageType}</span>;
 * }
 * ```
 *
 * @useWhen UI needs the page-layout classification available in legacy page metadata.
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
