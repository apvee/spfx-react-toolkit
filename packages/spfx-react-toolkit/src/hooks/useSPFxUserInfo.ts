// useSPFxUserInfo.ts
// Hook to access current user information

import { useMemo } from 'react';
import { getSPFxUserInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Return type for useSPFxUserInfo hook
 *
 * @example
 * ```tsx
 * // Inside a React component beneath a matching host-specific SPFx provider:
 * const user: SPFxUserInfo = useSPFxUserInfo();
 * console.log(user.displayName, user.email);
 * ```
 */
export interface SPFxUserInfo {
  /** User login name (e.g., "domain\\user" or email) */
  readonly loginName: string;
  
  /** User display name */
  readonly displayName: string;
  
  /** User email address (optional) */
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
 * function UserLabel() {
 *   const { displayName } = useSPFxUserInfo();
 *   return <span>{displayName}</span>;
 * }
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
