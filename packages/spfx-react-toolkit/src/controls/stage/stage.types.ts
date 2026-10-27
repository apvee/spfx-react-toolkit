import type { ContentSlot } from '../content-slot.types.internal';
import type * as React from 'react';
import type { DialogSurfaceProps } from '@fluentui/react-components';

/** Explicit surface presentation; no automatic responsive switching.
 * @example
 * ```ts
 * const presentation: StagePresentation = 'dialog';
 * ```
 */
export type StagePresentation = 'dialog' | 'drawer' | 'popover';

/** Surface width preset. Defaults to small; full is unavailable for Popover.
 * Dialog widths are 640/800/1120px, Drawer 320/592/940px, Popover 320/480/640px.
 * Popover widths are preferred sizes; native auto-size reduces them to available viewport space.
 * Full fills the available viewport; Dialog retains its responsive outer gap.
 * @example
 * ```ts
 * const size: StageSize = 'medium';
 * ```
 */
export type StageSize = 'small' | 'medium' | 'large' | 'full';

/** Standard permits native dismissal; explicit permits only trigger and close button requests.
 * @example
 * ```ts
 * const dismissMode: StageDismissMode = 'explicit';
 * ```
 */
export type StageDismissMode = 'standard' | 'explicit';

/** Logical Drawer edge; defaults to end and follows provider direction.
 * @example
 * ```ts
 * const drawerPosition: StageDrawerPosition = 'end';
 * ```
 */
export type StageDrawerPosition = 'start' | 'end';

/** Semantic title heading level, independent of typography; defaults to 2.
 * @example
 * ```ts
 * const headingLevel: StageHeadingLevel = 2;
 * ```
 */
export type StageHeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/** Normalized request source shared by all presentations.
 * @example
 * ```ts
 * const reason: StageOpenChangeReason = 'closeButton';
 * ```
 */
export type StageOpenChangeReason =
  | 'triggerClick' | 'closeButton' | 'escapeKeyDown'
  | 'backdropClick' | 'outsideClick' | 'focusOutside';

/** Native or React event accompanying an open-state request.
 * Native Popover Escape events can already be prevented by Fluent; prevention is not a veto.
 * @example
 * ```ts
 * const onOpenChange = (event: StageOpenChangeEvent) => console.log(event.type);
 * ```
 */
export type StageOpenChangeEvent =
  | React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>
  | React.FocusEvent<HTMLElement> | MouseEvent | TouchEvent | FocusEvent | KeyboardEvent;

/** Requested state; a controlled consumer accepts it by updating open.
 * @example
 * ```ts
 * const request: StageOpenChangeData = { open: false, reason: 'closeButton' };
 * ```
 */
export interface StageOpenChangeData {
  /** Requested next open state, which may differ from controlled open. */
  open: boolean;
  /** Interaction that requested the change. */
  reason: StageOpenChangeReason;
}

/** Bounded logical Popover placement; native flip/shift remains enabled.
 * @example
 * ```ts
 * const positioning: StagePopoverPositioning = { position: 'below', align: 'start' };
 * ```
 */
export interface StagePopoverPositioning {
  /** Preferred logical side; defaults to below. */
  position?: 'above' | 'below' | 'before' | 'after';
  /** Alignment along that side; defaults to start. */
  align?: 'start' | 'center' | 'end';
}

/** Select controlled or uncontrolled ownership once for the component lifetime.
 * Uncontrolled state starts closed unless defaultOpen is true. Prop changes do not notify.
 * @example
 * ```ts
 * const state: StageStateProps = { defaultOpen: true };
 * ```
 */
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
 * @example
 * ```ts
 * const attributes: StageSurfaceAttributes = { id: 'editor', 'aria-label': 'Editor' };
 * ```
 */
export type StageSurfaceAttributes = Omit<React.HTMLAttributes<HTMLDivElement>,
  'title' | 'children' | 'dangerouslySetInnerHTML' | 'role' | 'aria-modal' | 'aria-labelledby' | 'aria-describedby'>;

/** Shared root behavior and surface attributes. Content remains in the declared order.
 * @example
 * ```tsx
 * <Stage presentation="dialog"><Stage.Header title="Profile" /><Stage.Body>Content</Stage.Body></Stage>
 * ```
 */
interface StageCommonProps extends StageSurfaceAttributes {
  /** Caller-owned content; use one Header and one Body, with optional Footer. */
  children: React.ReactNode;
  /** Called once per permitted interaction with the event and requested `{ open, reason }`.
   * Explicit-dismissal filtering and cancelled trigger clicks produce no notification.
   * Retain controlled `open` to veto; uncontrolled state updates before notification.
   * Returned promises are not awaited; omission keeps uncontrolled behavior functional.
   * Calling preventDefault here does not veto; consumer trigger cancellation is respected separately. */
  onOpenChange?: (event: StageOpenChangeEvent, data: StageOpenChangeData) => void;
  /** Dismissal policy; defaults to standard. Hiding X does not change the policy. */
  dismissMode?: StageDismissMode;
  /** Defaults to `true`, removing content after native closing presence/motion. With `false`,
   * content first mounts on opening, then stays hidden/inert on close with effects active;
   * adapter changes can reset it. Compatible size/theme rerenders preserve child state. */
  unmountOnClose?: boolean;
  /** Optional portal destination; omitted inherits the Fluent host document and provider. */
  mountNode?: DialogSurfaceProps['mountNode'];
}

/** Root contract with exclusive state ownership and presentation-specific options.
 * Presentation is latched for an open session; size updates live. A pending full Dialog/Drawer
 * retains the last compatible size of an active Popover until the next session.
 * Trigger must be a ref-capable native button or Fluent Button with an accessible name.
 * Uses the host Fluent theme, direction, document and Griffel renderer without an SPFx
 * provider. Put a DOM form inside the portaled surface when its footer submits that form.
 * @example
 * ```tsx
 * <Stage presentation="popover" trigger={<button>Details</button>}><Stage.Header title="Details" /><Stage.Body>Content</Stage.Body></Stage>
 * ```
 */
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
    /** Preferred width: small 320px (default), medium 480px, large 640px.
     * Native positioning/auto-size bounds the rendered width to available viewport space.
     * Full is unsupported. */
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

/** Owned Stage UI labels. SPFx selects the native resource baseline without a toolkit provider.
 * Supply a partial dictionary to `Stage.Header.strings`. Omitted/undefined members use
 * that baseline; explicit empty and whitespace strings are preserved. Header's direct
 * `closeLabel` prop takes precedence. Caller title/description/action content is unchanged.
 * @example
 * ```tsx
 * const strings: StageStrings = { closeLabel: 'Dismiss profile' };
 * <Stage presentation="dialog"><Stage.Header title="Profile" strings={strings} /></Stage>
 * ```
 */
export interface StageStrings {
  /** Accessible name of the close button. Omission through a partial dictionary uses the host resource. */
  readonly closeLabel: string;
}

/** Header div attributes and named content slots; strings are literal, JSX is preserved.
 * Named props accept shorthand, fixed-tag native props or children render callbacks.
 * Objects request inner containers; refs identify those containers, not the heading or ID shell.
 * Callbacks receive the fixed tag/composed props during rendering and must forward the props.
 * Empty static descriptions do not register naming; Close remains outside consumer action.
 * Direct strings and props-object string children are visually limited to one title line
 * or two description lines with ellipsis; full text remains in the DOM for naming.
 * JSX and render callbacks keep caller-owned presentation in every Stage variant.
 * The title must be nonempty phrasing content without nested headings.
 * @example
 * ```tsx
 * <Stage presentation="dialog"><Stage.Header title="Profile" description="Edit your details" action={<button>Help</button>} /></Stage>
 * ```
 */
export interface StageHeaderProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'children' | 'dangerouslySetInnerHTML'> {
  /** Required nonempty inner span slot; subtitle2/primary heading wrapper. Literal strings (direct or props children) use one line with ellipsis; JSX/callbacks own presentation. Full text remains accessible. */
  title: ContentSlot<'span'>;
  /** Optional inner span slot; caption1/subtle defaults with zero text gap. Literal strings (direct or props children) use at most two lines with ellipsis; JSX/callbacks own presentation. Full text remains accessible. */
  description?: ContentSlot<'span'>;
  /** Inner span slot for a decorative 24px icon, hidden from assistive technology. */
  icon?: ContentSlot<'span'>;
  /** Inner div slot for the caller's primary header action, before the separately owned Close button.
   * Prefer one visible command; an application-owned menu is allowed. No automatic dismissal,
   * child counting or menu conversion. Action and Close stay on one row, vertically centered
   * with the title/description block in every presentation and at narrow widths.
   * Omission removes the caller action region. Shorthand, native div props/ref and render callbacks remain supported. */
  action?: ContentSlot<'div'>;
  /** Title heading level; defaults to 2. */
  headingLevel?: StageHeadingLevel;
  /** Show the final close button; defaults to true. When hidden, provide a visible modal exit. */
  showCloseButton?: boolean;
  /** Accessible close-button label. Overrides strings and the host resource; explicit empty/whitespace values are preserved. */
  closeLabel?: string;
  /** Optional per-Header UI dictionary. Omitted/undefined members use the host-selected
   * native resource; explicit empty/whitespace strings are preserved. Direct closeLabel wins. */
  strings?: Partial<StageStrings>;
}

/** Scrollable Body div and caller-owned children; raw HTML is unsupported. Requires
 * a `Stage` ancestor; native attributes and ref identify the body div.
 * @example
 * ```tsx
 * <Stage presentation="dialog"><Stage.Body><input aria-label="Name" /></Stage.Body></Stage>
 * ```
 */
export interface StageBodyProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'dangerouslySetInnerHTML'> {
  /** Body content grows within the available surface space. */
  children?: React.ReactNode;
}

/** Footer div containing free JSX; actions do not close automatically. Requires a `Stage`
 * ancestor; native attributes and ref target the footer div. Omit it to leave no gap.
 * @example
 * ```tsx
 * <Stage presentation="dialog"><Stage.Footer><button>Save</button></Stage.Footer></Stage>
 * ```
 */
export interface StageFooterProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'dangerouslySetInnerHTML'> {
  /** Caller-owned footer actions or content. */
  children?: React.ReactNode;
}
