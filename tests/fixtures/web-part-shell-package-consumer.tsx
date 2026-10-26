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

// Packed generic JSX must infer instance properties even when a props spread retains all slot branches.
<Shell {...props} action={{ ref: React.createRef<HTMLDivElement>(), 'data-actions': 'slot', children: (Tag, slotProps) => <Tag {...slotProps}><button>Refresh</button></Tag> }} />;
<Shell {...props} action={false} />;
// Fluent and React 17 retain legacy null absence.
<Shell {...props} action={null} />;
declare const legacyAction: React.ReactNode;
<Shell {...props} action={legacyAction} />;
// @ts-expect-error The fixed slot tag cannot replace the Shell action div.
<Shell {...props} action={{ as: 'button', children: 'Action' }} />;
// @ts-expect-error Action slots preserve the raw HTML prohibition.
<Shell {...props} action={{ dangerouslySetInnerHTML: { __html: 'unsafe' } }} />;
// @ts-expect-error Functions remain children callbacks inside slot objects.
<Shell {...props} action={() => 'Action'} />;
// @ts-expect-error The root raw HTML prohibition remains intact for spread consumers.
<Shell {...props} dangerouslySetInnerHTML={{ __html: 'unsafe' }} />;

interface OtherProperties extends IWebPartShellProperties { otherId: string; }
// @ts-expect-error Explicit application generics must still match the instance property contract.
<Shell<OtherProperties> instance={instance} />;

// @ts-expect-error The shared editor implementation also remains private.
import {} from '@apvee/spfx-react-toolkit/controls/editable-text/editable-text-editor.internal';
// @ts-expect-error Retired title support paths do not become public aliases.
import {} from '@apvee/spfx-react-toolkit/controls/web-part-shell/title/web-part-title-color-picker.internal';
// @ts-expect-error Retired title toolbar paths remain private.
import {} from '@apvee/spfx-react-toolkit/controls/web-part-shell/title/web-part-title-toolbar.internal';
