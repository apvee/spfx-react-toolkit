/**
 * Current-list metadata exists only when PageContext includes an identified list.
 *
 * @packageDocumentation
 * @docPage hooks/user-site
 */
// useSPFxListInfo.ts
// Hook to access list information (when in list context)

import { useMemo } from 'react';
import { getSPFxListInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Current PageContext list identity and best-effort template metadata; no list query is issued.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxListInfo, useSPFxListInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const list: SPFxListInfo | undefined = useSPFxListInfo();
 *   console.log(list?.title);
 *   return <span>{list?.title ?? 'No list'}</span>;
 * }
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
  
  /** List template type (100 for Generic List, 101 for Document Library); undefined when host metadata omits it. */
  readonly baseTemplate?: number;
  
  /**
   * Always supplied by the hook: true only when baseTemplate is 101; undefined is permitted by the public shape but the hook returns false for missing/different templates.
   */
  readonly isDocumentLibrary?: boolean;
}

/**
 * Maps the current PageContext list when it has an ID. Missing list/ID returns undefined. Missing titles fall back to Unknown List and missing URLs to an empty string. isDocumentLibrary is true only for template 101; no list data is fetched. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Current list metadata, or undefined when no list ID is available.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example Read ListInfo metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxListInfo } from '@apvee/spfx-react-toolkit';
 *
 * function ListLabel() {
 *   const list = useSPFxListInfo();
 *   return <span>{list?.title ?? 'No list context'}</span>;
 * }
 * ```
 *
 * @useWhen a list-hosted view needs the current list identity and template metadata.
 *
 * @example Inspect the optional current list context
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxListInfo } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxListInfo();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxListInfo result">
 *         <dt>List context</dt><dd>{info ? 'Available' : 'No list context'}</dd>
 *         <dt>List ID</dt><dd>{info?.id ?? 'Unavailable'}</dd>
 *         <dt>List title</dt><dd>{info?.title ?? 'Unavailable'}</dd>
 *         <dt>List URL</dt><dd>{info?.serverRelativeUrl || 'Unavailable'}</dd>
 *         <dt>Template</dt><dd>{info?.baseTemplate === undefined ? 'Unavailable' : String(info.baseTemplate)}</dd>
 *         <dt>Document library</dt><dd>{String(info?.isDocumentLibrary ?? false)}</dd>
 *       </dl>,
 *     diagnostics: <p>No list ID means no list context. Missing title becomes Unknown List, URL becomes empty, and template 101 alone identifies a library. No list request is issued.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual ListInfo values and explicit unavailable fields as the native host changes.' };
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
