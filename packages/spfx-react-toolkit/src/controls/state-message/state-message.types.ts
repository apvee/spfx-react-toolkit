import type * as React from 'react';
import type { ButtonProps } from '@fluentui/react-components';

/** Selects decorative artwork and announcement defaults, without supplying application copy.
 * @example const preset: StateMessagePreset = 'empty';
 */
export type StateMessagePreset = 'generic' | 'empty' | 'error' | 'success' | 'welcome';

/** Semantic heading level, independent of the title's typography; defaults to 3.
 * @example <StateMessage title="No results" headingLevel={2} />
 */
export type StateMessageHeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

/** Announcement role owned by the text region, excluding artwork and actions.
 * @example <StateMessage title="Save failed" role="alert" />
 */
export type StateMessageRole = 'status' | 'alert';

type ActionBase<Props> = Omit<Props,
  'appearance' | 'as' | 'children' | 'href' | 'icon' | 'onClick' | 'shape' | 'size' | 'type' | 'dangerouslySetInnerHTML'
> & {
  /** Visible localized label. Supply phrasing content, without nested interactive elements. */
  label: React.ReactNode;
  /** Optional Fluent Button icon. */
  icon?: ButtonProps['icon'];
  /** Unsupported: label is the sole content pathway. */
  dangerouslySetInnerHTML?: never;
};

/** Callback action rendered as a native button with type=button.
 * @example const action: StateMessageClickAction = { label: 'Retry', onClick: retry };
 */
export type StateMessageClickAction = ActionBase<Exclude<ButtonProps, { as: 'a' }>> & {
  /** Navigation destinations belong to link actions. */
  href?: never;
  /** Called when the native button is activated. */
  onClick: React.MouseEventHandler<HTMLButtonElement>;
};

/** Navigation action rendered as a Fluent Button with a native anchor root.
 * @example const action: StateMessageLinkAction = { label: 'Help', href: '/help' };
 */
export type StateMessageLinkAction = ActionBase<Extract<ButtonProps, { as: 'a' }>> & {
  /** Destination assigned to the anchor. */
  href: string;
  /** Browsing context for navigation. */
  target?: React.HTMLAttributeAnchorTarget;
  /** Relationship to the destination. */
  rel?: string;
  /** Optional native anchor activation callback. */
  onClick?: React.MouseEventHandler<HTMLAnchorElement>;
};

/** An application callback or navigation intention; appearance is owned by its position.
 * @example <StateMessage title="No results" primaryAction={{ label: 'Retry', onClick: retry }} />
 */
export type StateMessageAction = StateMessageClickAction | StateMessageLinkAction;

/** Bounded section state composed from toolkit text, Fluent buttons and private artwork.
 * @example <StateMessage preset="empty" title="No results" description="Try **other filters**." />
 */
export interface StateMessageProps extends Omit<React.HTMLAttributes<HTMLDivElement>, 'children' | 'title' | 'role' | 'dangerouslySetInnerHTML'> {
  /** Optional artwork/announcement preset. Omission renders no automatic illustration or role. */
  preset?: StateMessagePreset;
  /** Required literal title, rendered with subtitle2Stronger typography. Strings are not Markdown. */
  title: React.ReactNode;
  /** Optional inline Markdown or caller-owned phrasing JSX, rendered inside p. */
  description?: React.ReactNode;
  /** Title HTML heading level, default 3. Choose it for the surrounding document hierarchy. */
  headingLevel?: StateMessageHeadingLevel;
  /** Replaces preset artwork; false hides it, undefined uses the preset. Custom visuals retain their semantics. */
  visual?: React.ReactNode;
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
