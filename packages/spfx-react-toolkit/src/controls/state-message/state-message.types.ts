import type { ContentSlot } from '../content-slot.types.internal';
import type * as React from 'react';
import type { ButtonProps } from '@fluentui/react-components';

/** Selects decorative artwork and announcement defaults, without supplying application copy.
 * `empty`, `error` and `success` use `status`; `generic` and `welcome` have no default role.
 * @example
 * ```ts
 * const preset: StateMessagePreset = 'empty';
 * ```
 */
export type StateMessagePreset = 'generic' | 'empty' | 'error' | 'success' | 'welcome';

/** Semantic heading level, independent of the title's typography; defaults to 3.
 * @example
 * ```tsx
 * <StateMessage title="No results" headingLevel={2} />
 * ```
 */
export type StateMessageHeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/** Announcement role owned by the text region, excluding artwork and actions.
 * @example
 * ```tsx
 * <StateMessage title="Save failed" role="alert" />
 * ```
 */
export type StateMessageRole = 'status' | 'alert';

/** Shared native Fluent Button attributes with label/icon content, excluding shell-owned
 * appearance, size, shape, root tag and children. The concrete action selects button or anchor
 * attributes; label slots are literal phrasing content and callback output remains caller-owned.
 * @template Props Native Fluent button or anchor props inherited by the concrete action.
 * @example
 * ```ts
 * const label: ActionBase<Exclude<ButtonProps, { as: 'a' }>> = { label: 'Retry', disabled: true };
 * ```
 */
type ActionBase<Props> = Omit<Props,
  'appearance' | 'as' | 'children' | 'href' | 'icon' | 'onClick' | 'shape' | 'size' | 'type' | 'dangerouslySetInnerHTML'
> & {
  /** Literal phrasing content or span props/render slot inside the owned Button. No nested interactive elements; ref targets the inner span. */
  label: ContentSlot<'span'>;
  /** Optional native Fluent Button icon slot; omitted actions have no icon. */
  icon?: ButtonProps['icon'];
  /** Unsupported: label is the sole content pathway. */
  dangerouslySetInnerHTML?: never;
};

/** Callback action rendered as a native button with type=button.
 * @example
 * ```ts
 * const action: StateMessageClickAction = { label: 'Retry', onClick: retry };
 * ```
 */
export type StateMessageClickAction = ActionBase<Exclude<ButtonProps, {
  /** Type filter excluding anchor props; the action owns its button root. */
  as: 'a'
}>> & {
  /** Navigation destinations belong to link actions. */
  href?: never;
  /** Called for native button activation with the React mouse event. No automatic
   * loading, awaiting or error handling is added; the application owns asynchronous work. */
  onClick: React.MouseEventHandler<HTMLButtonElement>;
};

/** Navigation action rendered as a Fluent Button with a native anchor root.
 * @example
 * ```ts
 * const action: StateMessageLinkAction = { label: 'Help', href: '/help' };
 * ```
 */
export type StateMessageLinkAction = ActionBase<Extract<ButtonProps, {
  /** Type filter selecting native anchor props; consumers supply href instead of as. */
  as: 'a'
}>> & {
  /** Destination assigned to the anchor. */
  href: string;
  /** Browsing context for navigation; omission uses the current context. */
  target?: React.HTMLAttributeAnchorTarget;
  /** Native link relationship; omission adds no toolkit relationship tokens. */
  rel?: string;
  /** Optional native anchor activation callback receiving the React mouse event.
   * Prevent the event to cancel native navigation; any asynchronous return is not awaited. */
  onClick?: React.MouseEventHandler<HTMLAnchorElement>;
};

/** An application callback or navigation intention; appearance is owned by its position.
 * @example
 * ```tsx
 * <StateMessage title="No results" primaryAction={{ label: 'Retry', onClick: retry }} />
 * ```
 */
export type StateMessageAction = StateMessageClickAction | StateMessageLinkAction;

/** Bounded section state composed from toolkit text, Fluent buttons and private artwork.
 * Uses the host Fluent theme/Griffel renderer without an SPFx provider. Named slots use
 * fixed tags; object refs target inner containers and render callbacks must forward their
 * composed props. Omit optional actions/content to leave their regions absent.
 * @example
 * ```tsx
 * <StateMessage preset="empty" title="No results" description="Try **other filters**." />
 * ```
 */
export interface StateMessageProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'children' | 'title' | 'role' | 'dangerouslySetInnerHTML'> {
  /** Optional artwork/announcement preset. Omission renders no automatic illustration or role. */
  preset?: StateMessagePreset;
  /** Required literal phrasing content or span props/render slot inside the owned heading. Strings are not Markdown; ref targets the inner span. */
  title: ContentSlot<'span'>;
  /** Optional span content slot inside p. Static strings/fragments retain inline Markdown; render callbacks are opaque and own formatting. Ref targets the inner span. */
  description?: ContentSlot<'span'>;
  /** Title HTML heading level, default 3. Choose it for the surrounding document hierarchy. */
  headingLevel?: StateMessageHeadingLevel;
  /** Replaces preset artwork; false hides it, undefined uses the preset. Custom div props/render slots retain their semantics; ref targets the inner div. */
  visual?: ContentSlot<'div'>;
  /** Highest-emphasis action, with primary appearance. */
  primaryAction?: StateMessageAction;
  /** Supporting action, with secondary appearance. */
  secondaryAction?: StateMessageAction;
  /** Lowest-emphasis action, with subtle appearance. */
  tertiaryAction?: StateMessageAction;
  /** Text-region role override. Empty/error/success default to status; alert must be explicit. */
  role?: StateMessageRole;
  /** Unsupported: title and description are the owned content pathways. */
  dangerouslySetInnerHTML?: never;
}
