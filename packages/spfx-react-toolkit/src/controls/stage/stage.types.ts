/**
 * Explicit Dialog, Drawer and Popover surfaces sharing one caller-owned content structure.
 *
 * @packageDocumentation
 * @docPage controls/stage
 */
import type { ContentSlot } from '../content-slot.types.internal';
import type * as React from 'react';
import type { DialogSurfaceProps } from '@fluentui/react-components';

/**
 * Explicit surface presentation; no automatic responsive switching.
 * @example
 * ```ts
 * import { Stage, StagePresentation } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * const presentation: StagePresentation = 'dialog';
 * ```
 */
export type StagePresentation = 'dialog' | 'drawer' | 'popover';

/**
 * Surface width preset. Defaults to small; full is unavailable for Popover.
 * Dialog widths are 640/800/1120px, Drawer 320/592/940px, Popover 320/480/640px.
 * Popover widths are preferred sizes; native auto-size reduces them to available viewport space.
 * Full fills the available viewport; Dialog retains its responsive outer gap.
 * @example
 * ```ts
 * import { Stage, StageSize } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * const size: StageSize = 'medium';
 * ```
 */
export type StageSize = 'small' | 'medium' | 'large' | 'full';

/**
 * Standard permits native dismissal; explicit permits only trigger and close button requests.
 * @example
 * ```ts
 * import { Stage, StageDismissMode } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * const dismissMode: StageDismissMode = 'explicit';
 * ```
 */
export type StageDismissMode = 'standard' | 'explicit';

/**
 * Logical Drawer edge; defaults to end and follows provider direction.
 * @example
 * ```ts
 * import { Stage, StageDrawerPosition } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * const drawerPosition: StageDrawerPosition = 'end';
 * ```
 */
export type StageDrawerPosition = 'start' | 'end';

/**
 * Semantic title heading level, independent of typography; defaults to 2.
 * @example
 * ```ts
 * import { Stage, StageHeadingLevel } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * const headingLevel: StageHeadingLevel = 2;
 * ```
 */
export type StageHeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/**
 * Normalized request source shared by all presentations.
 * @example
 * ```ts
 * import { Stage, StageOpenChangeReason } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * const reason: StageOpenChangeReason = 'closeButton';
 * ```
 */
export type StageOpenChangeReason =
  | 'triggerClick' | 'closeButton' | 'escapeKeyDown'
  | 'backdropClick' | 'outsideClick' | 'focusOutside';

/**
 * Native or React event accompanying an open-state request.
 * Native Popover Escape events can already be prevented by Fluent; prevention is not a veto.
 * @example
 * ```ts
 * import { Stage, StageOpenChangeEvent } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * const onOpenChange = (event: StageOpenChangeEvent) => console.log(event.type);
 * ```
 */
export type StageOpenChangeEvent =
  | React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>
  | React.FocusEvent<HTMLElement> | MouseEvent | TouchEvent | FocusEvent | KeyboardEvent;

/**
 * Requested state; a controlled consumer accepts it by updating open.
 * @example
 * ```ts
 * import { Stage, StageOpenChangeData } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * const request: StageOpenChangeData = { open: false, reason: 'closeButton' };
 * ```
 */
export interface StageOpenChangeData {
  /** Requested next open state, which may differ from controlled open. */
  open: boolean;
  /** Interaction that requested the change. */
  reason: StageOpenChangeReason;
}

/**
 * Bounded logical Popover placement; native flip/shift remains enabled.
 * @example
 * ```ts
 * import { Stage, StagePopoverPositioning } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * const positioning: StagePopoverPositioning = { position: 'below', align: 'start' };
 * ```
 */
export interface StagePopoverPositioning {
  /**
   * Preferred logical side; defaults to below.
   * When omitted, preferred placement is below.
   * @defaultValue below
   */
  position?: 'above' | 'below' | 'before' | 'after';
  /**
   * Alignment along that side; defaults to start.
   * When omitted, preferred alignment is start.
   * @defaultValue start
   */
  align?: 'start' | 'center' | 'end';
}

/**
 * Select controlled or uncontrolled ownership once for the component lifetime.
 * Uncontrolled state starts closed unless defaultOpen is true. Prop changes do not notify.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { Stage, StageProps } from '@apvee/spfx-react-toolkit/controls/stage';
 * const stageProps = {
 *   presentation: 'dialog', defaultOpen: true,
 *   children: <><Stage.Header title="Details" /><Stage.Body>Initially open</Stage.Body></>,
 * } satisfies StageProps;
 * export function Details(): React.ReactElement {
 *   return <Stage {...stageProps} />;
 * }
 * ```
 */
export type StageStateProps =
  | {
    /** Authoritative controlled state; retain it to veto dirty/pending close requests. */
    open: boolean;
    /**
     * Unavailable when open is controlled.
     * Omit this property when open supplies controlled state.
     */
    defaultOpen?: never;
  }
  | {
    /** Omit open for uncontrolled state; changing ownership during mount is unsupported. */
    open?: never;
    /**
     * Initial uncontrolled state only; defaults to false.
     * Omit when open is controlled. Otherwise omission starts closed.
     */
    defaultOpen?: boolean;
  };

/**
 * Native attributes forwarded to the surface div. Stage owns content and dialog semantics.
 * Consumer className and style remain escape hatches; refs are forwarded by Stage.
 * @example
 * ```ts
 * import { StageProps } from '@apvee/spfx-react-toolkit/controls/stage';
 * const attributes: Pick<StageProps, 'id' | 'aria-label'> = { id: 'editor', 'aria-label': 'Editor' };
 * ```
 */
export type StageSurfaceAttributes = Omit<React.HTMLAttributes<HTMLDivElement>,
  'title' | 'children' | 'dangerouslySetInnerHTML' | 'role' | 'aria-modal' | 'aria-labelledby' | 'aria-describedby'>;

/**
 * Shared root behavior and surface attributes. Content remains in the declared order.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { Stage } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * <Stage presentation="dialog"><Stage.Header title="Profile" /><Stage.Body>Content</Stage.Body></Stage>
 * ```
 */
interface StageCommonProps extends StageSurfaceAttributes {
  /** Caller-owned content; use one Header and one Body, with optional Footer. */
  children: React.ReactNode;
  /**
   * Called once per permitted interaction with the event and requested `{ open, reason }`.
   * Explicit-dismissal filtering and cancelled trigger clicks produce no notification.
   * Retain controlled `open` to veto; uncontrolled state updates before notification.
   * Returned promises are not awaited; omission keeps uncontrolled behavior functional.
   * Calling preventDefault here does not veto; consumer trigger cancellation is respected separately.
   * When omitted, no application notification runs; uncontrolled state requests still update normally.
   * @param event - Native/React interaction event accompanying the notification.
   * @param data - Requested next open state and normalized reason. A controlled consumer accepts by updating open or refuses by retaining it; uncontrolled state updates before notification.
   */
  onOpenChange?: (event: StageOpenChangeEvent, data: StageOpenChangeData) => void;
  /**
   * Dismissal policy; defaults to standard. Hiding X does not change the policy.
   * When omitted, standard native dismissal applies.
   * @defaultValue standard
   */
  dismissMode?: StageDismissMode;
  /**
   * Defaults to `true`, removing content after native closing presence/motion. With `false`,
   * content first mounts on opening, then stays hidden/inert on close with effects active;
   * adapter changes can reset it. Compatible size/theme rerenders preserve child state.
   * When omitted, content unmounts after native closing presence/motion.
   * @defaultValue true
   */
  unmountOnClose?: boolean;
  /** Optional portal destination; omitted inherits the Fluent host document and provider. */
  mountNode?: DialogSurfaceProps['mountNode'];
}

/**
 * Root contract with exclusive state ownership and presentation-specific options.
 * Presentation is latched for an open session; size updates live. A pending full Dialog/Drawer
 * retains the last compatible size of an active Popover until the next session.
 * Trigger must be a ref-capable native button or Fluent Button with an accessible name.
 * Uses the host Fluent theme, direction, document and Griffel renderer without an SPFx
 * provider. Put a DOM form inside the portaled surface when its footer submits that form.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { Stage } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * <Stage presentation="popover" trigger={<button>Details</button>}><Stage.Header title="Details" /><Stage.Body>Content</Stage.Body></Stage>
 * ```
 * @remarks
 * ### Semantics, portals and sizing
 * Dialog and Drawer are native modal surfaces; Popover is a nonmodal group without a focus trap. Stage owns surface roles and heading/description associations. Native focus placement/restoration and nested portal relationships remain Fluent-owned; removed triggers cannot restore focus. An explicit mountNode is caller-owned: verify its owner document, theme, direction, renderer styles and portal relationships. Stage adds no global renderer or arbitrary z-index. Dialog uses equal clamp(8px, 1vw, 16px) outer gaps, modal padding is 24px and Popover padding is 12px. Full Drawer reaches its viewport edge. Small viewport heights can scroll the whole surface so Header/Footer remain reachable.
 * ### Dismissal and child popups
 * Controlled open accepts or refuses each permitted request; prop changes do not notify. Explicit mode filters native dismissal before notification but leaves trigger/Close requests. Child popup Escape consumption remains respected. The legacy compat DatePicker input consumes Escape even with its calendar closed, so close via X/footer or move focus first; Stage does not rewrite third-party key handlers.
 * ### Form composition
 * Put a native form or Form.Root inside the portaled surface so footer submit controls share actual DOM form ownership. React context from a form outside the portal does not establish native form ownership. Preserve compatible flex layout in caller wrappers. Footer actions do not dismiss automatically.
 */
export type StageProps = StageCommonProps & StageStateProps & (
  | {
    /** Modal Dialog presentation; explicit and required. */
    presentation: 'dialog';
    /**
     * Width preset; defaults to small. Full fills both viewport dimensions with an outer gap.
     * When omitted, the width preset is small.
     * @defaultValue small
     */
    size?: StageSize;
    /**
     * Optional integrated accessible trigger; native ref and event handlers are preserved.
     * When omitted for Dialog/Drawer, open the controlled surface externally; Popover requires a trigger.
     */
    trigger?: React.ReactElement;
    /**
     * Drawer edge is unavailable for Dialog.
     * Omit this property in the selected presentation.
     */
    drawerPosition?: never;
    /**
     * Popover placement is unavailable for Dialog.
     * Omit this property in the selected presentation.
     */
    positioning?: never;
    /**
     * Popover arrow is unavailable for Dialog.
     * Omit this property in the selected presentation.
     */
    withArrow?: never;
  }
  | {
    /** Modal Drawer presentation; explicit and required. */
    presentation: 'drawer';
    /**
     * Width preset; defaults to small. Full fills available width.
     * When omitted, the width preset is small.
     * @defaultValue small
     */
    size?: StageSize;
    /**
     * Optional integrated accessible trigger; native ref and event handlers are preserved.
     * When omitted for Dialog/Drawer, open the controlled surface externally; Popover requires a trigger.
     */
    trigger?: React.ReactElement;
    /**
     * Logical edge; defaults to end.
     * When omitted in Drawer, the logical end edge is used.
     * @defaultValue end
     */
    drawerPosition?: StageDrawerPosition;
    /**
     * Popover placement is unavailable for Drawer.
     * Omit this property in the selected presentation.
     */
    positioning?: never;
    /**
     * Popover arrow is unavailable for Drawer.
     * Omit this property in the selected presentation.
     */
    withArrow?: never;
  }
  | {
    /** Nonmodal Popover presentation, opened by trigger click. */
    presentation: 'popover';
    /**
     * Preferred width: small 320px (default), medium 480px, large 640px.
     * Native positioning/auto-size bounds the rendered width to available viewport space.
     * Full is unsupported.
     * When omitted, the width preset is small.
     * @defaultValue small
     */
    size?: Exclude<StageSize, 'full'>;
    /** Required integrated accessible trigger; native ref and event handlers are preserved. */
    trigger: React.ReactElement;
    /**
     * Drawer edge is unavailable for Popover.
     * Omit this property in the selected presentation.
     */
    drawerPosition?: never;
    /**
     * Logical preferred placement; defaults to below/start.
     * When omitted in Popover, preferred placement is below/start.
     */
    positioning?: StagePopoverPositioning;
    /**
     * Show Fluent's native anchor arrow; defaults to false and updates while open.
     * When omitted in Popover, no arrow is shown.
     * @defaultValue false
     */
    withArrow?: boolean;
  }
);

/**
 * Owned Stage UI labels. SPFx selects the native resource baseline without a toolkit provider.
 * Supply a partial dictionary to `Stage.Header.strings`. Omitted/undefined members use
 * that baseline; explicit empty and whitespace strings are preserved. Header's direct
 * `closeLabel` prop takes precedence. Caller title/description/action content is unchanged.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { Stage, StageStrings } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * const strings: StageStrings = { closeLabel: 'Dismiss profile' };
 * <Stage presentation="dialog"><Stage.Header title="Profile" strings={strings} /></Stage>
 * ```
 */
export interface StageStrings {
  /** Accessible name of the close button. Omission through a partial dictionary uses the host resource. */
  readonly closeLabel: string;
}

/**
 * Header div attributes and named content slots; strings are literal, JSX is preserved.
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
 * import * as React from 'react';
 * import { Stage } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * <Stage presentation="dialog"><Stage.Header title="Profile" description="Edit your details" action={<button>Help</button>} /></Stage>
 * ```
 * @remarks
 * Native attributes/ref target the outer div; the neutral inner Card isolates selectable ancestor Card context without enabling interaction styling. Use singular action; the removed actions name is not supported. Close remains separately owned and reserved at the logical end. Keep action controls shrinkable, retaining full accessible labels. An empty/whitespace static description object creates no surface-description association unless explicitly named; callbacks request association without probing output. Shorthand whitespace retains its existing presence policy. Raw HTML is unsupported.
 */
export interface StageHeaderProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'title' | 'children' | 'dangerouslySetInnerHTML'> {
  /** Required nonempty inner span slot; subtitle2/primary heading wrapper. Literal strings (direct or props children) use one line with ellipsis; JSX/callbacks own presentation. Full text remains accessible. */
  title: ContentSlot<'span'>;
  /**
   * Optional inner span slot; caption1/subtle defaults with zero text gap. Literal strings (direct or props children) use at most two lines with ellipsis; JSX/callbacks own presentation. Full text remains accessible.
   * When omitted, no description region or accessible association is registered.
   */
  description?: ContentSlot<'span'>;
  /**
   * Inner span slot for a decorative 24px icon, hidden from assistive technology.
   * When omitted, no decorative icon region is rendered.
   */
  icon?: ContentSlot<'span'>;
  /**
   * Inner div slot for the caller's primary header action, before the separately owned Close button.
   * Prefer one visible command; an application-owned menu is allowed. No automatic dismissal,
   * child counting or menu conversion. Action and Close stay on one row, vertically centered
   * with the title/description block in every presentation and at narrow widths.
   * Omission removes the caller action region. Shorthand, native div props/ref and render callbacks remain supported.
   * When omitted, no caller action region is rendered; Close remains independently controlled.
   */
  action?: ContentSlot<'div'>;
  /**
   * Title heading level; defaults to 2.
   * When omitted, the title wrapper is h2.
   * @defaultValue 2
   */
  headingLevel?: StageHeadingLevel;
  /**
   * Show the final close button; defaults to true. When hidden, provide a visible modal exit.
   * When omitted, the final Close button is visible.
   * @defaultValue true
   */
  showCloseButton?: boolean;
  /**
   * Accessible close-button label. Overrides strings and the host resource; explicit empty/whitespace values are preserved.
   * When omitted, strings.closeLabel then the native SPFx resource supplies the label.
   */
  closeLabel?: string;
  /** Optional per-Header UI dictionary. Omitted/undefined members use the host-selected
   * native resource; explicit empty/whitespace strings are preserved. Direct closeLabel wins. */
  strings?: Partial<StageStrings>;
}

/**
 * Scrollable Body div and caller-owned children; raw HTML is unsupported. Requires
 * a `Stage` ancestor; native attributes and ref identify the body div.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { Stage } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * <Stage presentation="dialog"><Stage.Body><input aria-label="Name" /></Stage.Body></Stage>
 * ```
 */
export interface StageBodyProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'dangerouslySetInnerHTML'> {
  /**
   * Body content grows within the available surface space.
   * When omitted, the part remains an empty native div.
   */
  children?: React.ReactNode;
}

/**
 * Footer div containing free JSX; actions do not close automatically. Requires a `Stage`
 * ancestor; native attributes and ref target the footer div. Omit it to leave no gap.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { Stage } from '@apvee/spfx-react-toolkit/controls/stage';
 *
 * <Stage presentation="dialog"><Stage.Footer><button>Save</button></Stage.Footer></Stage>
 * ```
 */
export interface StageFooterProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'dangerouslySetInnerHTML'> {
  /**
   * Caller-owned footer actions or content.
   * When omitted, the part remains an empty native div.
   */
  children?: React.ReactNode;
}
