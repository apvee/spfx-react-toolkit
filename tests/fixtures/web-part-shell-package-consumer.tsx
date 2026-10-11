import * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { WebPartShell as rootShell, getWebPartShellPropertiesMetadata as rootMetadata } from '@apvee/spfx-react-toolkit';
import { WebPartShell as domainShell, getWebPartShellPropertiesMetadata as domainMetadata } from '@apvee/spfx-react-toolkit/controls';
import { WebPartShell as controlShell, getWebPartShellPropertiesMetadata as controlMetadata } from '@apvee/spfx-react-toolkit/controls/web-part-shell';
import type { IWebPartShellProperties, WebPartShellProps, WebPartShellSurface, WebPartShellPadding } from '@apvee/spfx-react-toolkit';
import type { WebPartShellProps as DomainProps } from '@apvee/spfx-react-toolkit/controls';
import type { WebPartShellProps as ControlProps } from '@apvee/spfx-react-toolkit/controls/web-part-shell';
interface Properties extends IWebPartShellProperties { listId: string; }
declare const instance: BaseClientSideWebPart<Properties>;
export const aliases: typeof rootShell[] = [domainShell, controlShell];
export const metadataAliases: typeof rootMetadata[] = [domainMetadata, controlMetadata];
const surface: WebPartShellSurface = 'section'; const padding: WebPartShellPadding = 'small';
export const props: WebPartShellProps<Properties> = { instance, surface, padding, action: <button>Refresh</button> };
export const domainProps: DomainProps<Properties> = props;
export const controlProps: ControlProps<Properties> = props;
const Shell = controlShell;
export const element = <Shell {...props} ref={React.createRef<HTMLDivElement>()} onClick={event => event.currentTarget.focus()}>Content</Shell>;
// @ts-expect-error Title semantics come from the persisted HTML.
<Shell instance={instance} titleVariant="h2" />;
// @ts-expect-error The private engine is not configurable through Shell props.
<Shell instance={instance} editor={{}} />;
// @ts-expect-error Editor implementation paths are private.
import {} from '@apvee/spfx-react-toolkit/controls/web-part-shell/title/web-part-title-editor.internal';
// @ts-expect-error There is no public editor engine type.
import type { Editor } from '@apvee/spfx-react-toolkit/controls/web-part-shell';
