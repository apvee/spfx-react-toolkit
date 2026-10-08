import * as React from 'react';
import { useSx } from '../../hooks/useSx';
import * as typography from '../../helpers/styles/typography';
import * as text from '../../helpers/styles/text';
import * as textAlign from '../../helpers/styles/text-align';
import * as foreground from '../../helpers/styles/foreground';
import { none as noBlockMargin } from '../../helpers/styles/margin-block';
import { extraSmall as inlinePadding } from '../../helpers/styles/padding-inline';
import { small as smallRadius } from '../../helpers/styles/border-radius';
import { alternative as codeBackground } from '../../helpers/styles/background';
import { useFormattedTextStyles, useLineClampClass } from './formatted-text.styles.internal';
import { renderFormattedTextChildren } from './formatted-text-content.internal';
import type { FormattedTextProps } from './formatted-text.types';

/** Render toolkit typography with bounded inline Markdown or caller-owned JSX.
 * `as` defaults to span independently of variant. Native attributes/events and ref reach that root.
 * Uses the host's Fluent theme and Griffel renderer; no SPFx provider is required.
 * Valid maxLines applies block clamp layout without changing the tag and wins over truncate/wrap.
 * Unsafe Markdown links preserve their label; raw HTML and unsupported document syntax stay literal.
 * Arbitrary JSX remains untouched. Ellipsis needs constrained width and suitable element layout.
 * @param props Native text attributes, presentation options and owned content.
 * @returns A semantic text root, including when children render nothing.
 * @example
 * <FormattedText as="h2" variant="title2" maxLines={2}>
 *   Read the **updates** in [documentation](./docs).
 * </FormattedText>
 */
export const FormattedText = React.forwardRef<HTMLElement, FormattedTextProps>((props, ref) => {
  const {
    as: Root = 'span', variant = 'body1', font = 'base', align = 'start',
    markdown = true, subtle = false, highlight = false, italic = false,
    underline = false, strikethrough = false, wrap = true, truncate = false,
    maxLines, children, className, style, ...nativeProps
  } = props;
  // Never propagate an alternative content pathway, even from untyped callers.
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx();
  const styles = useFormattedTextStyles();
  const clamp = maxLines !== undefined && Number.isFinite(maxLines) && Number.isInteger(maxLines) && maxLines > 0;
  const lineClampClass = useLineClampClass(clamp ? maxLines : undefined);
  const content = renderFormattedTextChildren(children, {
    markdown,
    codeClassName: sx(codeBackground, inlinePadding, smallRadius, styles.code),
    linkClassName: sx(styles.link, highlight && styles.highlightedLink),
  });
  return React.createElement(Root, {
    ...nativeProps, ref,
    className: sx(
      styles.root, noBlockMargin, typography[variant],
      align === 'justify' ? styles.justify : textAlign[align],
      font === 'monospace' && styles.monospace, font === 'numeric' && styles.numeric,
      subtle && foreground.subtle, italic && styles.italic,
      underline && strikethrough ? styles.combinedDecoration : underline ? styles.underline : strikethrough && styles.strikethrough,
      clamp ? text.wrap : truncate ? text.truncate : wrap ? text.wrap : text.noWrap,
      clamp && styles.clamp, lineClampClass, className,
    ),
    style,
  }, highlight ? <mark className={sx(foreground.primary, inlinePadding, smallRadius, styles.mark)}>{content}</mark> : content);
});
FormattedText.displayName = 'FormattedText';
