import * as React from 'react';
import type { StageStrings } from '../../packages/spfx-react-toolkit/src/controls/stage';
import type {
  StageProps, StageHeaderProps, StageBodyProps, StageFooterProps, StageOpenChangeEvent,
  StageOpenChangeReason, StagePresentation, StageSize, StageDrawerPosition,
  StageDismissMode, StageHeadingLevel, StageStateProps, StagePopoverPositioning,
} from '../../packages/spfx-react-toolkit/src/controls/stage/stage.types';
import type { StageController, StageAdapterProps } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-controller.internal';
import type { StageContextValue } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-context.internal';

const dialog: StageProps = { presentation: 'dialog', children: 'Body', size: 'full', open: true,
  id: 'profile', className: 'consumer', style: { color: 'red' }, 'aria-label': 'Profile',
  onClick: event => event.currentTarget.focus(),
  onOpenChange: (event, data) => { event.preventDefault(); const open: boolean = data.open;
    const reason: StageOpenChangeReason = data.reason; void [open, reason]; },
};
const drawer: StageProps = { presentation: 'drawer', children: 'Body', drawerPosition: 'start', defaultOpen: true };
const popover: StageProps = { presentation: 'popover', children: 'Body', trigger: <button>Open</button>,
  positioning: { position: 'after', align: 'end' }, withArrow: true, mountNode: document.body, dismissMode: 'explicit', unmountOnClose: false };
const header: StageHeaderProps = { title: <em>Title</em>, description: 'Description', icon: <span />,
  action: <button>Save</button>, headingLevel: 6, showCloseButton: false, closeLabel: 'Close', id: 'header' };
// @ts-expect-error Stage.Header uses singular action; the old plural name has no alias.
const legacyHeaderAction: StageHeaderProps = { title: 'Title', actions: <button>Help</button> };
void legacyHeaderAction;
const body: StageBodyProps = { children: <input />, onClick: event => event.currentTarget.focus() };
const footer: StageFooterProps = { children: <button>Save</button>, className: 'footer' };
type Parts = [StageProps, StageHeaderProps, StageBodyProps, StageFooterProps];
type RefContracts = { [K in keyof Parts]: React.ForwardRefExoticComponent<React.PropsWithoutRef<Parts[K]> & React.RefAttributes<HTMLDivElement>> };
const refs: React.RefAttributes<HTMLDivElement> = { ref: React.createRef<HTMLDivElement>() };
const nativeEvents: StageOpenChangeEvent[] = [new MouseEvent('click'), new TouchEvent('touchstart'), new FocusEvent('focus'), new KeyboardEvent('keydown')];
const events = (mouse: React.MouseEvent<HTMLElement>, key: React.KeyboardEvent<HTMLElement>, focus: React.FocusEvent<HTMLElement>): StageOpenChangeEvent[] => [mouse, key, focus];
const variants: [StagePresentation, StageSize, StageDrawerPosition, StageDismissMode, StageHeadingLevel, StageStateProps, StagePopoverPositioning] =
  ['popover', 'medium', 'end', 'standard', 1, {}, { position: 'below', align: 'start' }];
void [dialog, drawer, popover, header, body, footer, refs, nativeEvents, events, variants];
export type { RefContracts, StageController, StageAdapterProps, StageContextValue };
// @ts-expect-error presentation is required
const missingPresentation: StageProps = { children: 'Body' };
// @ts-expect-error explicit mode only, no responsive auto
const auto: StageProps = { presentation: 'auto', children: 'Body' };
// @ts-expect-error Popover has no full size
const fullPopover: StageProps = { presentation: 'popover', trigger: <button />, size: 'full', children: 'Body' };
// @ts-expect-error Popover requires an integrated trigger
const missingTrigger: StageProps = { presentation: 'popover', children: 'Body' };
// @ts-expect-error controlled and uncontrolled initial values are exclusive
const mixedState: StageProps = { presentation: 'dialog', open: true, defaultOpen: true, children: 'Body' };
// @ts-expect-error drawer props do not belong to dialog
const dialogPosition: StageProps = { presentation: 'dialog', drawerPosition: 'end', children: 'Body' };
// @ts-expect-error popover positioning does not belong to drawer
const drawerPopover: StageProps = { presentation: 'drawer', positioning: { align: 'end' }, children: 'Body' };
// @ts-expect-error arrow belongs to Popover only
const dialogArrow: StageProps = { presentation: 'dialog', withArrow: true, children: 'Body' };
// @ts-expect-error arrow belongs to Popover only
const drawerArrow: StageProps = { presentation: 'drawer', withArrow: true, children: 'Body' };
// @ts-expect-error arrow visibility is boolean
const badArrow: StageProps = { presentation: 'popover', trigger: <button />, withArrow: 'yes', children: 'Body' };
// @ts-expect-error position has a bounded logical union
const badPosition: StagePopoverPositioning = { position: 'left' };
// @ts-expect-error alignment is bounded
const badAlignment: StagePopoverPositioning = { align: 'stretch' };
// @ts-expect-error raw HTML is owned
const rawRoot: StageProps = { presentation: 'dialog', children: 'Body', dangerouslySetInnerHTML: { __html: '' } };
// @ts-expect-error root native title is owned by Header
const rootTitle: StageProps = { presentation: 'dialog', children: 'Body', title: 'Title' };
// @ts-expect-error surface role is owned
const rootRole: StageProps = { presentation: 'dialog', children: 'Body', role: 'region' };
// @ts-expect-error modal semantics are owned
const rootModal: StageProps = { presentation: 'dialog', children: 'Body', 'aria-modal': false };
// @ts-expect-error title association is owned
const rootLabel: StageProps = { presentation: 'dialog', children: 'Body', 'aria-labelledby': 'external' };
// @ts-expect-error description association is owned
const rootDescription: StageProps = { presentation: 'dialog', children: 'Body', 'aria-describedby': 'external' };
// @ts-expect-error Header requires a title
const noHeaderTitle: StageHeaderProps = {};
// @ts-expect-error Header content is supplied by named slots
const headerChildren: StageHeaderProps = { title: 'Title', children: 'Raw' };
// @ts-expect-error Header heading level is bounded
const invalidHeading: StageHeaderProps = { title: 'Title', headingLevel: 7 };
// @ts-expect-error Header owns its HTML
const rawHeader: StageHeaderProps = { title: 'Title', dangerouslySetInnerHTML: { __html: '' } };
// @ts-expect-error Body owns its HTML
const rawBody: StageBodyProps = { dangerouslySetInnerHTML: { __html: '' } };
// @ts-expect-error Footer owns its HTML
const rawFooter: StageFooterProps = { dangerouslySetInnerHTML: { __html: '' } };

import { Stage } from '../../packages/spfx-react-toolkit/src/controls/stage/stage';
<Stage presentation="dialog" size="full" ref={React.createRef<HTMLDivElement>()}>
  <Stage.Header title="Profile" action={<button type="button">Help</button>} closeLabel="Chiudi" ref={React.createRef<HTMLDivElement>()} />
  <Stage.Body ref={React.createRef<HTMLDivElement>()}><input aria-label="Name" /></Stage.Body>
  <Stage.Footer ref={React.createRef<HTMLDivElement>()}><button type="button">Save</button></Stage.Footer>
</Stage>;
<Stage presentation="popover" trigger={<button type="button">Details</button>}><Stage.Header title="Details" /><Stage.Body>Context</Stage.Body></Stage>;
// @ts-expect-error Compound Header still requires its named title.
<Stage.Header />;
// @ts-expect-error Header has named content slots, not arbitrary children.
<Stage.Header title="Title">Content</Stage.Header>;
// @ts-expect-error Body forwards a div ref, not a button ref.
<Stage.Body ref={React.createRef<HTMLButtonElement>()} />;
// @ts-expect-error No root title duplicates Header's API.
<Stage presentation="dialog" title="Duplicate title">Content</Stage>;
// @ts-expect-error There is no responsive presentation resolver.
<Stage presentation={{ small: 'dialog', large: 'drawer' }}>Content</Stage>;
// @ts-expect-error Popover has no full-screen mode.
<Stage presentation="popover" size="full" trigger={<button>Open</button>}>Content</Stage>;

const conditionalSlot: React.ReactNode = false;
<Stage.Header title={conditionalSlot} description={0} icon={{ children: <span>Icon</span>, ref: React.createRef<HTMLSpanElement>() }} />;
<Stage.Header title={{ as: 'span', children: (Tag, props) => <Tag {...props}>Callback title</Tag> }} />;
// @ts-expect-error named title slots cannot replace the semantic heading
<Stage.Header title={{ as: 'h1', children: 'Title' }} />;
// @ts-expect-error raw HTML is forbidden in named slots
<Stage.Header title={{ dangerouslySetInnerHTML: { __html: 'raw' } }} />;
// @ts-expect-error direct render functions are not slot shorthand
<Stage.Header title={() => <span>Direct callback</span>} />;

<Stage.Header title="Icon sizing" icon={{ children: (Tag, props) => <Tag {...props}><svg width={80} height={80} /></Tag> }} description={{ children: <> {false} </> }} />;

const localizedStage: StageStrings = { closeLabel: 'Dismiss' };
const localizedHeader: StageHeaderProps = { title: 'Profile', strings: localizedStage, closeLabel: 'Close profile' };
const partialHeader: StageHeaderProps = { title: 'Profile', strings: { closeLabel: undefined } };
// @ts-expect-error Stage dictionary values remain strings.
const invalidStageStrings: StageHeaderProps = { title: 'Profile', strings: { closeLabel: 12 } };
void [localizedHeader, partialHeader, invalidStageStrings];
