import type * as React from 'react';
import type { DialogSurfaceProps } from '@fluentui/react-components';

/** Explicit surface presentation; no automatic responsive switching.
 * @example const presentation: StagePresentation = 'dialog'; */
export type StagePresentation = 'dialog' | 'drawer' | 'popover';

/** Surface width preset. Defaults to small; full is unavailable for Popover.
 * Dialog widths are 640/800/1120px, Drawer 320/592/940px, Popover 240/320/480px.
 * Full fills the available viewport; Dialog retains its responsive outer gap.
 * @example const size: StageSize = 'medium'; */
export type StageSize = 'small' | 'medium' | 'large' | 'full';

/** Standard permits native dismissal; explicit permits only trigger and close button requests.
 * @example const dismissMode: StageDismissMode = 'explicit'; */
export type StageDismissMode = 'standard' | 'explicit';

/** Logical Drawer edge; defaults to end and follows provider direction.
 * @example const drawerPosition: StageDrawerPosition = 'end'; */
export type StageDrawerPosition = 'start' | 'end';

/** Semantic title heading level, independent of typography; defaults to 2.
 * @example const headingLevel: StageHeadingLevel = 2; */
export type StageHeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/** Normalized request source shared by all presentations.
 * @example if (data.reason === 'closeButton') saveDraft(); */
export type StageOpenChangeReason =
  | 'triggerClick' | 'closeButton' | 'escapeKeyDown'
  | 'backdropClick' | 'outsideClick' | 'focusOutside';

/** Native or React event accompanying an open-state request.
 * Native Popover Escape events can already be prevented by Fluent; prevention is not a veto.
 * @example const onOpenChange = (event: StageOpenChangeEvent) => console.log(event.type); */
export type StageOpenChangeEvent =
  | React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>
  | React.FocusEvent<HTMLElement> | MouseEvent | TouchEvent | FocusEvent | KeyboardEvent;

/** Requested state; a controlled consumer accepts it by updating open.
 * @example onOpenChange={(_, data) => { if (!dirty) setOpen(data.open); }} */
export interface StageOpenChangeData {
  /** Requested next open state, which may differ from controlled open. */
  open: boolean;
  /** Interaction that requested the change. */
  reason: StageOpenChangeReason;
}

/** Bounded logical Popover placement; native flip/shift remains enabled.
 * @example const positioning: StagePopoverPositioning = { position: 'below', align: 'start' }; */
export interface StagePopoverPositioning {
  /** Preferred logical side; defaults to below. */
  position?: 'above' | 'below' | 'before' | 'after';
  /** Alignment along that side; defaults to start. */
  align?: 'start' | 'center' | 'end';
}

/** Select controlled or uncontrolled ownership once for the component lifetime.
 * Uncontrolled state starts closed unless defaultOpen is true. Prop changes do not notify.
 * @example const state: StageStateProps = { defaultOpen: true }; */
export type StageStateProps =
  | {
    /** Authoritative controlled state; retain it to veto dirty/pending close requests. */
    open: boolean;
    /** Unavailable when open is controlled. */
    defaultOpen?: never;
  }
  | {
    /** Omit open for uncontrolled state; changing ownership during mount is unsupported. */
    open?: never;
    /** Initial uncontrolled state only; defaults to false. */
    defaultOpen?: boolean;
  };

/** Native attributes forwarded to the surface div. Stage owns content and dialog semantics.
 * Consumer className and style remain escape hatches; refs are forwarded by Stage.
 * @example const attributes: StageSurfaceAttributes = { id: 'editor', 'aria-label': 'Editor' }; */
export type StageSurfaceAttributes = Omit<React.HTMLAttributes<HTMLDivElement>,
  'title' | 'children' | 'dangerouslySetInnerHTML' | 'role' | 'aria-modal' | 'aria-labelledby' | 'aria-describedby'>;

/** Shared root behavior and surface attributes. Content remains in the declared order.
 * @example <Stage presentation="dialog"><Stage.Header title="Profile" /><Stage.Body>Content</Stage.Body></Stage> */
interface StageCommonProps extends StageSurfaceAttributes {
  /** Caller-owned content; use one Header and one Body, with optional Footer. */
  children: React.ReactNode;
  /** Request notification, called once per accepted interaction. Retain controlled open to veto.
   * Calling preventDefault here does not veto; consumer trigger cancellation is respected separately. */
  onOpenChange?: (event: StageOpenChangeEvent, data: StageOpenChangeData) => void;
  /** Dismissal policy; defaults to standard. Hiding X does not change the policy. */
  dismissMode?: StageDismissMode;
  /** Defaults to true. False retains compatible content after first open; adapter changes can reset it. */
  unmountOnClose?: boolean;
  /** Optional portal destination; omitted inherits the Fluent host document and provider. */
  mountNode?: DialogSurfaceProps['mountNode'];
}

/** Root contract with exclusive state ownership and presentation-specific options.
 * Presentation is latched for an open session; size updates live. A pending full Dialog/Drawer
 * retains the last compatible size of an active Popover until the next session.
 * Trigger must be a ref-capable native button or Fluent Button with an accessible name.
 * @example <Stage presentation="popover" trigger={<button>Details</button>}><Stage.Header title="Details" /><Stage.Body>Content</Stage.Body></Stage> */
export type StageProps = StageCommonProps & StageStateProps & (
  | {
    /** Modal Dialog presentation; explicit and required. */
    presentation: 'dialog';
    /** Width preset; defaults to small. Full fills both viewport dimensions with an outer gap. */
    size?: StageSize;
    /** Optional integrated accessible trigger; native ref and event handlers are preserved. */
    trigger?: React.ReactElement;
    /** Drawer edge is unavailable for Dialog. */
    drawerPosition?: never;
    /** Popover placement is unavailable for Dialog. */
    positioning?: never;
    /** Popover arrow is unavailable for Dialog. */
    withArrow?: never;
  }
  | {
    /** Modal Drawer presentation; explicit and required. */
    presentation: 'drawer';
    /** Width preset; defaults to small. Full fills available width. */
    size?: StageSize;
    /** Optional integrated accessible trigger; native ref and event handlers are preserved. */
    trigger?: React.ReactElement;
    /** Logical edge; defaults to end. */
    drawerPosition?: StageDrawerPosition;
    /** Popover placement is unavailable for Drawer. */
    positioning?: never;
    /** Popover arrow is unavailable for Drawer. */
    withArrow?: never;
  }
  | {
    /** Nonmodal Popover presentation, opened by trigger click. */
    presentation: 'popover';
    /** Width preset; defaults to small. Full is unsupported. */
    size?: Exclude<StageSize, 'full'>;
    /** Required integrated accessible trigger; native ref and event handlers are preserved. */
    trigger: React.ReactElement;
    /** Drawer edge is unavailable for Popover. */
    drawerPosition?: never;
    /** Logical preferred placement; defaults to below/start. */
    positioning?: StagePopoverPositioning;
    /** Show Fluent's native anchor arrow; defaults to false and updates while open. */
    withArrow?: boolean;
  }
);

/** Header div attributes and named content slots; strings are literal, JSX is preserved.
 * The title must be nonempty phrasing content without nested headings.
 * @example <Stage.Header title="Profile" description="Edit your details" actions={<button>Help</button>} /> */
export interface StageHeaderProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'children' | 'dangerouslySetInnerHTML'> {
  /** Required nonempty title phrasing content; subtitle2/primary defaults apply to its heading wrapper. */
  title: React.ReactNode;
  /** Optional supporting phrasing content; caption1/subtle defaults, with zero text gap. */
  description?: React.ReactNode;
  /** Decorative 24px icon, hidden from assistive technology. */
  icon?: React.ReactNode;
  /** Caller-owned actions; they do not close Stage automatically. */
  actions?: React.ReactNode;
  /** Title heading level; defaults to 2. */
  headingLevel?: StageHeadingLevel;
  /** Show the final close button; defaults to true. When hidden, provide a visible modal exit. */
  showCloseButton?: boolean;
  /** Accessible close-button label; defaults to Close and is consumer-localizable. */
  closeLabel?: string;
}

/** Scrollable Body div and caller-owned children; raw HTML is unsupported.
 * @example <Stage.Body><input aria-label="Name" /></Stage.Body> */
export interface StageBodyProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'dangerouslySetInnerHTML'> {
  /** Body content grows within the available surface space. */
  children?: React.ReactNode;
}

/** Footer div containing free JSX; actions do not close automatically. Omit it to leave no gap.
 * @example <Stage.Footer><button>Save</button></Stage.Footer> */
export interface StageFooterProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'dangerouslySetInnerHTML'> {
  /** Caller-owned footer actions or content. */
  children?: React.ReactNode;
}
