// useSPFxListInfo.ts
// Hook to access list information (when in list context)

import { useMemo } from 'react';
import { getSPFxListInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Return type for useSPFxListInfo hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const list: SPFxListInfo | undefined = useSPFxListInfo();
 * console.log(list?.title);
 * ```
 */
export interface SPFxListInfo {
  /** List ID (GUID) */
  readonly id: string;
  
  /**
   * Current list title; defaults to Unknown List when host metadata omits it.
   */
  readonly title: string;
  
  /**
   * List server-relative URL; defaults to an empty string when host metadata omits it.
   */
  readonly serverRelativeUrl: string;
  
  /** List template type (e.g., 100 for Generic List, 101 for Document Library) */
  readonly baseTemplate?: number;
  
  /**
   * True only when baseTemplate is 101; missing or different templates produce false.
   */
  readonly isDocumentLibrary?: boolean;
}

/**
 * Maps the current PageContext list when it has an ID. Missing list/ID returns undefined. Missing titles fall back to Unknown List and missing URLs to an empty string. isDocumentLibrary is true only for template 101; no list data is fetched. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Current list metadata, or undefined when no list ID is available.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example
 * ```tsx
 * function ListLabel() {
 *   const list = useSPFxListInfo();
 *   return <span>{list?.title ?? 'No list context'}</span>;
 * }
 * ```
 */
export function useSPFxListInfo(): SPFxListInfo | undefined {
  const pageContext = useSPFxPageContext();
  const list = (pageContext as unknown as {
    list?: {
      id?: { toString: () => string };
      title?: string;
      serverRelativeUrl?: string;
      baseTemplate?: number;
    };
  }).list;
  
  return useMemo(
    () => getSPFxListInfo(pageContext),
    [pageContext, list]
  );
}
