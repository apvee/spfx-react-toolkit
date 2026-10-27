import * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { WebPartShell, getWebPartShellPropertiesMetadata } from '__TOOLKIT__';
import type { IWebPartShellProperties } from '__TOOLKIT__';
export const metadata = getWebPartShellPropertiesMetadata();
export function Example({ instance }: { instance: BaseClientSideWebPart<IWebPartShellProperties> }): React.ReactElement {
  return <WebPartShell instance={instance}><span>Stable content</span></WebPartShell>;
}
