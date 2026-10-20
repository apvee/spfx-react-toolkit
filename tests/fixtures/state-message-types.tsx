import * as React from 'react';
import { StateMessage } from '../../packages/spfx-react-toolkit/src/controls/state-message/state-message';
import type { StateMessageAction } from '../../packages/spfx-react-toolkit/src/controls/state-message/state-message.types';
const action: StateMessageAction = { label: 'Retry', onClick: event => event.currentTarget.disabled = true };
<StateMessage title="No results" preset="empty" description={<em>Details</em>} visual={false}
  primaryAction={action} tertiaryAction={{ label: 'Help', href: '/help', onClick: event => event.currentTarget.href }}
  ref={React.createRef<HTMLDivElement>()} onClick={event => event.currentTarget.focus()} />;
// @ts-expect-error title is required
<StateMessage preset="empty" />;
// @ts-expect-error no layout wrapper options
<StateMessage title="State" layout={{ direction: 'row' }} />;
// @ts-expect-error root raw HTML is unsupported
<StateMessage title="State" dangerouslySetInnerHTML={{ __html: 'unsafe' }} />;
// @ts-expect-error action raw HTML is unsupported
<StateMessage title="State" primaryAction={{ label: 'Retry', onClick() {}, dangerouslySetInnerHTML: { __html: 'unsafe' } }} />;
// @ts-expect-error callback action must have callback
<StateMessage title="State" primaryAction={{ label: 'Retry' }} />;
// @ts-expect-error anchor callback receives anchor, not button
<StateMessage title="State" primaryAction={{ label: 'Help', href: '/help', onClick: (event: React.MouseEvent<HTMLButtonElement>) => event.currentTarget }} />;
// @ts-expect-error appearance belongs to action position
<StateMessage title="State" primaryAction={{ label: 'Help', href: '/help', appearance: 'subtle' }} />;
// @ts-expect-error heading semantic level is bounded
<StateMessage title="State" headingLevel={7} />;
const spanRef = React.createRef<HTMLSpanElement>();
const condition: boolean = true;
<StateMessage title={{ ref: spanRef, children: (Component, props) => <Component {...props}>Literal</Component> }}
  description={condition && { children: '**Markdown**', 'data-description': 'slot' }} visual={{ ref: React.createRef<HTMLDivElement>(), children: 'Art' }}
  primaryAction={{ label: { ref: spanRef, children: 'Retry' }, onClick() {} }} />;
// @ts-expect-error title cannot replace the owned heading
<StateMessage title={{ as: 'h1', children: 'Title' }} />;
// @ts-expect-error slot HTML injection is unsupported
<StateMessage title="Safe" description={{ dangerouslySetInnerHTML: { __html: 'unsafe' } }} />;
// @ts-expect-error callbacks belong inside slot objects
<StateMessage title={() => 'Title'} />;
