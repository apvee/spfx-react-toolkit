// useSPFxListInfo.ts
// Hook to access list information (when in list context)

import { useMemo } from 'react';
import { getSPFxListInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Return type for useSPFxListInfo hook
 */
export interface SPFxListInfo {
  /** List ID (GUID) */
  readonly id: string;
  
  /** List title */
  readonly title: string;
  
  /** List server relative URL */
  readonly serverRelativeUrl: string;
  
  /** List template type (e.g., 100 for Generic List, 101 for Document Library) */
  readonly baseTemplate?: number;
  
  /** Whether list is a document library */
  readonly isDocumentLibrary?: boolean;
}

/**
 * Hook to access list information
 * 
 * Provides information about the current SharePoint list/library
 * when component is rendered in a list context (e.g., list view,
 * list web part, field customizer).
 * 
 * Returns undefined if not in a list context.
 * 
 * Information provided:
 * - id: Unique identifier
 * - title: List title
 * - serverRelativeUrl: Server-relative URL
 * - baseTemplate: List template type
 * - isDocumentLibrary: Whether it's a document library
 * 
 * Note: List context is available in Field Customizers and some
 * List View WebParts. Standard page WebParts typically don't have
 * list context. Always check for undefined return value.
 * 
 * Useful for:
 * - List-specific operations in Field Customizers
 * - Conditional rendering based on list type
 * - Building list URLs
 * - List metadata display
 * 
 * @returns List information or undefined if not in list context
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const list = useSPFxListInfo();
 *   
 *   if (!list) {
 *     return <div>Not in list context</div>;
 *   }
 *   
 *   return (
 *     <div>
 *       <h2>{list.title}</h2>
 *       <p>List ID: {list.id}</p>
 *       {list.isDocumentLibrary && <p>Document Library</p>}
 *     </div>
 *   );
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
