import * as React from 'react';
import { SharePointGroupTrim } from '__TOOLKIT__';
export function Consumer() {
  return <SharePointGroupTrim groups={[{ id: 2 }]}>Protected<SharePointGroupTrim.Loading>Checking</SharePointGroupTrim.Loading><SharePointGroupTrim.Fallback>Denied</SharePointGroupTrim.Fallback><SharePointGroupTrim.ErrorFallback>Failed</SharePointGroupTrim.ErrorFallback></SharePointGroupTrim>;
}
