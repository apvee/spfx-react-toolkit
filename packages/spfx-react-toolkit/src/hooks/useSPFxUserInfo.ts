// useSPFxUserInfo.ts
// Hook to access current user information

import { useMemo } from 'react';
import { getSPFxUserInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Return type for useSPFxUserInfo hook
 */
export interface SPFxUserInfo {
  /** User login name (e.g., "domain\\user" or email) */
  readonly loginName: string;
  
  /** User display name */
  readonly displayName: string;
  
  /** User email address (optional) */
  readonly email?: string;
  
  /** Whether user is an external guest user */
  readonly isExternal: boolean;
}

/**
 * Hook to access current user information
 * 
 * Provides:
 * - loginName: Login identifier
 * - displayName: Display name
 * - email: Email address
 * - isExternal: Whether user is a guest
 * 
 * Useful for:
 * - Personalization
 * - Authorization checks
 * - User-specific logging
 * - Display user information
 * 
 * @returns User information
 * 
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { displayName, email, isExternal } = useSPFxUserInfo();
 *   
 *   return (
 *     <div>
 *       <h2>Welcome, {displayName}!</h2>
 *       {email && <p>Email: {email}</p>}
 *       {isExternal && <p>Guest User</p>}
 *     </div>
 *   );
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
