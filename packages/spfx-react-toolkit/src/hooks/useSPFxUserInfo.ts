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
 * @example
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
 */
export function useSPFxUserInfo(): SPFxUserInfo {
  const pageContext = useSPFxPageContext();
  const user = pageContext.user;
  
  return useMemo(
    () => getSPFxUserInfo(pageContext),
    [pageContext, user.loginName, user.displayName, user.email, user.isExternalGuestUser]
  );
}
