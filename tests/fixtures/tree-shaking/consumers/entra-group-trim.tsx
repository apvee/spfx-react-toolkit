import * as React from 'react';
import { EntraGroupTrim } from '__TOOLKIT__';
export function Consumer() {
  return <EntraGroupTrim groups={[{ id: "aaaaaaaa-0000-0000-0000-000000000002" }]}>Protected<EntraGroupTrim.Loading>Checking</EntraGroupTrim.Loading><EntraGroupTrim.Fallback>Denied</EntraGroupTrim.Fallback><EntraGroupTrim.ErrorFallback>Failed</EntraGroupTrim.ErrorFallback></EntraGroupTrim>;
}
