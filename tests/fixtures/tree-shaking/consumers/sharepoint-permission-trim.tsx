import * as React from 'react';
import { SPPermission } from '@microsoft/sp-page-context';
import { SharePointPermissionTrim } from '__TOOLKIT__';
export function Consumer() {
  return <SharePointPermissionTrim permissions={[SPPermission.viewListItems]}>Protected<SharePointPermissionTrim.Loading>Checking</SharePointPermissionTrim.Loading><SharePointPermissionTrim.Fallback>Denied</SharePointPermissionTrim.Fallback><SharePointPermissionTrim.ErrorFallback>Failed</SharePointPermissionTrim.ErrorFallback></SharePointPermissionTrim>;
}
