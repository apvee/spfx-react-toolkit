/**
 * User display metadata comes from the native PageContext without I/O.
 *
 * @packageDocumentation
 * @docPage hooks/user-site
 */
// useSPFxUserInfo.ts
// Hook to access current user information

import { useMemo } from 'react';
import { getSPFxUserInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Current PageContext user names, optional email and external-guest flag without I/O.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxUserInfo, useSPFxUserInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const user: SPFxUserInfo = useSPFxUserInfo();
 *   console.log(user.displayName, user.email);
 *   return <span>{user.displayName}</span>;
 * }
 * ```
 */
export interface SPFxUserInfo {
  /** User login name (e.g., "domain\\user" or email) */
  readonly loginName: string;
  
  /** User display name */
  readonly displayName: string;
  
  /** User email address; undefined when the host omits it */
  readonly email?: string;
  
  /**
   * Host external-guest flag; defaults to false when unavailable.
   */
  readonly isExternal: boolean;
}

/**
 * Maps current PageContext user metadata without I/O. Email remains optional; a missing external-guest flag defaults to false. These display values do not authorize a server operation. Call in a React component or hook beneath the matching host-specific SPFx provider.
 *
 * @returns Login/display names, optional email and external-guest flag.
 * @throws Error when no host-specific SPFx provider is present.
 *
 * @example Read UserInfo metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxUserInfo } from '@apvee/spfx-react-toolkit';
 *
 * function UserLabel() {
 *   const { displayName } = useSPFxUserInfo();
 *   return <span>{displayName}</span>;
 * }
 * ```
 *
 * @useWhen UI needs the signed-in user display/login metadata already supplied by PageContext.
 *
 * @example Inspect the current user display metadata
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxUserInfo } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const info = useSPFxUserInfo();
 *   const regions: DemoRegions = {
 *     children: <dl aria-label="useSPFxUserInfo result">
 *         <dt>User</dt><dd>{info.displayName}</dd>
 *         <dt>Email</dt><dd>{info.email ?? 'Unavailable'}</dd>
 *         <dt>External user</dt><dd>{String(info.isExternal)}</dd>
 *         <dt>Login available</dt><dd>{String(Boolean(info.loginName))}</dd>
 *       </dl>,
 *     diagnostics: <p>Supplied display metadata performs no profile fetch and is not an authorization result.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Inspect the actual UserInfo values and explicit unavailable fields as the native host changes.' };
 * ```
 */
export function useSPFxUserInfo(): SPFxUserInfo {
  const pageContext = useSPFxPageContext();
  const user = pageContext.user;
  
  return useMemo(
    () => getSPFxUserInfo(pageContext),
    [pageContext, user.loginName, user.displayName, user.email, user.isExternalGuestUser]
  );
}
