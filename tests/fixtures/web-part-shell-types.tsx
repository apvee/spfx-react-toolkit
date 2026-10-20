import * as React from 'react';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { WebPartShell, getWebPartShellPropertiesMetadata } from '../../packages/spfx-react-toolkit/src/controls/web-part-shell/index';
import type { IWebPartShellProperties, WebPartShellProps } from '../../packages/spfx-react-toolkit/src/controls/web-part-shell/web-part-shell.types';
interface Properties extends IWebPartShellProperties { listId: string; }
declare const instance: BaseClientSideWebPart<Properties>;
const ref = React.createRef<HTMLDivElement>();
const props: WebPartShellProps<Properties> = { instance, surface: 'auto', padding: 'medium', dir: 'rtl', id: 'native' };
<WebPartShell {...props} ref={ref} style={{ width: 300 }} onClick={event => event.currentTarget.focus()} action={<button>Action</button>}>Child</WebPartShell>;
<WebPartShell<Properties> instance={instance} />;
const metadata = getWebPartShellPropertiesMetadata(); metadata.title.isSearchablePlainText;
// @ts-expect-error instance is required
<WebPartShell />;
// @ts-expect-error HTML has one owned pathway
<WebPartShell instance={instance} dangerouslySetInnerHTML={{ __html: '<h2>Unsafe</h2>' }} />;
// @ts-expect-error format derives from HTML
<WebPartShell instance={instance} titleVariant="h2" />;
// @ts-expect-error title visibility is persisted as showTitle
<WebPartShell instance={instance} isTitleEnabled />;
// @ts-expect-error no engine options on public API
<WebPartShell instance={instance} EditorOptions={{}} />;
// @ts-expect-error title is a persisted property, not a native div tooltip
<WebPartShell instance={instance} title="Title" />;

// @ts-expect-error heading semantics derive from HTML
<WebPartShell instance={instance} headingLevel={2} />;
<WebPartShell instance={instance} action={{ ref: React.createRef<HTMLDivElement>(), 'data-actions': 'slot', children: (Component, slotProps) => <Component {...slotProps}><button>Action</button></Component> }} />;
<WebPartShell instance={instance} action={false && <button>Hidden</button>} />;
// @ts-expect-error action tag is fixed
<WebPartShell instance={instance} action={{ as: 'button', children: 'Action' }} />;
// @ts-expect-error action raw HTML is unsupported
<WebPartShell instance={instance} action={{ dangerouslySetInnerHTML: { __html: 'unsafe' } }} />;
// @ts-expect-error callbacks belong inside slot objects
<WebPartShell instance={instance} action={() => 'Action'} />;
