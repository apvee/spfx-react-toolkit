import * as React from 'react';
import { Button } from '@fluentui/react-components';
import { FormattedText } from '../formatted-text/formatted-text';
import { useSx } from '../../hooks/useSx';
import * as flex from '../../helpers/styles/flex';
import * as flexItem from '../../helpers/styles/flex-item';
import * as alignItems from '../../helpers/styles/align-items';
import * as justifyContent from '../../helpers/styles/justify-content';
import * as gap from '../../helpers/styles/gap';
import * as padding from '../../helpers/styles/padding';
import * as width from '../../helpers/styles/width';
import * as height from '../../helpers/styles/height';
import * as minWidth from '../../helpers/styles/min-width';
import * as maxWidth from '../../helpers/styles/max-width';
import * as responsive from '../../helpers/styles/responsive';
import * as container from '../../helpers/styles/container';
import { useStateMessageStyles } from './state-message.styles.internal';
import { useStateMessageThemeMode } from './state-message-theme.internal';
import { resolvePresetIllustration } from './illustrations/illustration-registry.internal';
import type { StateMessageAction, StateMessageProps, StateMessageRole } from './state-message.types';

const headings = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] as const;
const presetRoles: Readonly<Record<NonNullable<StateMessageProps['preset']>, StateMessageRole | undefined>> = {
  generic: undefined, empty: 'status', error: 'status', success: 'status', welcome: undefined,
};
// Immutable descriptors are shared data, not a shared renderer or responsive state.
const messageLarge = responsive.medium(maxWidth.px(600));
const actionsLarge = responsive.medium(maxWidth.px(280));

function Action({ action, appearance, className }: {
  action: StateMessageAction; appearance: 'primary' | 'secondary' | 'subtle'; className: string;
}): React.ReactElement {
  const sx = useSx();
  const { label, className: consumerClass, ...props } = action;
  delete props.dangerouslySetInnerHTML;
  if (typeof props.href === 'string') {
    return <Button {...props} as="a" href={props.href} appearance={appearance} size="medium" shape="rounded"
      className={sx(className, consumerClass)}>{label}</Button>;
  }
  return <Button {...props} as="button" type="button" appearance={appearance} size="medium" shape="rounded"
    className={sx(className, consumerClass)}>{label}</Button>;
}

/** Replace a bounded section with a state explanation, decorative artwork and optional actions.
 * Reuses FormattedText, Fluent buttons and useSx container layout; requires no SPFx provider.
 * The nearest Fluent theme selects light/dark artwork by foreground/background luminance,
 * with a light fallback for absent or uninterpretable tokens. Custom visual overrides it.
 * Text/action widths follow the toolkit medium (640px) container breakpoint.
 * Artwork has fluid width up to 280px and natural height from its SVG aspect ratio.
 * Children of the text region must be semantically valid phrasing content. No autofocus.
 * @param props Native div attributes, localized content, preset and button/link intentions.
 * @returns The div root containing text, optional artwork and actions outside the live region.
 * @example
 * <StateMessage preset="empty" title="No results" description="Try **other filters**."
 *   primaryAction={{ label: 'Clear filters', onClick: clearFilters }} />
 */
export const StateMessage = React.forwardRef<HTMLDivElement, StateMessageProps>(({
  preset, title, description, headingLevel = 3, visual, primaryAction, secondaryAction,
  tertiaryAction, role, className, ...nativeProps
}, ref) => {
  delete nativeProps.dangerouslySetInnerHTML;
  // Own all content, including input from untyped callers.
  Reflect.deleteProperty(nativeProps, 'children');
  const sx = useSx();
  const styles = useStateMessageStyles();
  const mode = useStateMessageThemeMode();
  const Illustration = visual === undefined && preset ? resolvePresetIllustration(preset) : undefined;
  const showVisual = Illustration !== undefined || React.Children.toArray(visual).length > 0;
  const actionClass = sx(width.full);
  return <div {...nativeProps} ref={ref} className={sx(styles.root, container.inlineSize, width.full, minWidth.zero,
    flex.column, alignItems.center, justifyContent.center, gap.extraLarge, padding.medium, className)}>
    {showVisual && <div data-state-message-visual="" className={sx(styles.visual, width.full, maxWidth.px(280),
      height.auto, flexItem.noShrink)}>{Illustration ? <Illustration mode={mode} /> : visual}</div>}
    <div data-state-message-text="" role={role ?? (preset ? presetRoles[preset] : undefined)}
      className={sx(flex.column, alignItems.center, gap.small, width.full, minWidth.zero, maxWidth.px(240), messageLarge)}>
      <FormattedText as={headings[headingLevel - 1] ?? 'h3'} variant="subtitle2Stronger" markdown={false} align="center">{title}</FormattedText>
      {description !== undefined && <FormattedText as="p" variant="body1" align="center">{description}</FormattedText>}
    </div>
    {(primaryAction || secondaryAction || tertiaryAction) && <div data-state-message-actions=""
      className={sx(flex.column, alignItems.stretch, gap.small, width.full, minWidth.zero, maxWidth.px(240), actionsLarge)}>
      {primaryAction && <Action action={primaryAction} appearance="primary" className={actionClass} />}
      {secondaryAction && <Action action={secondaryAction} appearance="secondary" className={actionClass} />}
      {tertiaryAction && <Action action={tertiaryAction} appearance="subtle" className={actionClass} />}
    </div>}
  </div>;
});
StateMessage.displayName = 'StateMessage';
