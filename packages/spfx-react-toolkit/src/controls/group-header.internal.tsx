import * as React from 'react';
import { inspectContentSlot, hasContentSlotContent, renderContentSlot } from './content-slot.internal';
import type { ContentSlot } from './content-slot.types.internal';
import { makeStyles } from '@griffel/react';
import { useSx } from '../hooks/useSx';
import * as flex from '../helpers/styles/flex';
import * as flexItem from '../helpers/styles/flex-item';
import * as alignItems from '../helpers/styles/align-items';
import * as gap from '../helpers/styles/gap';
import * as minWidth from '../helpers/styles/min-width';
import * as foreground from '../helpers/styles/foreground';
import * as typography from '../helpers/styles/typography';
import * as marginBlock from '../helpers/styles/margin-block';

const headings = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] as const;
const useStyles = makeStyles({
  text: { overflowWrap: 'anywhere' },
  // Only the introduced object/callback container participates in owned icon sizing.
  iconContent: {
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    inlineSize: '100%', blockSize: '100%',
    '& > i': { fontSize: 'inherit' },
    '& > svg': { inlineSize: '100%', blockSize: '100%' },
  },
  icon: {
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    inlineSize: '24px', blockSize: '24px', fontSize: '24px', lineHeight: '1',
    '& > i': { fontSize: 'inherit' },
    '& > svg': { inlineSize: '100%', blockSize: '100%' },
  },
  action: { maxInlineSize: '100%' },
});

/** Detect visible slot content without cloning caller elements or allocating placeholder slots. */
export function hasGroupContent(content: unknown): boolean {
  const inspected = inspectContentSlot(content);
  if (inspected.kind === 'absent') return false;
  if (inspected.kind !== 'shorthand') return true;
  return hasLegacyGroupContent(inspected.children);
}

/** Semantic descriptions retain legacy whitespace/JSX presence but empty props do not name a control. */
export function hasGroupDescriptionContent(content: unknown): boolean {
  const inspected = inspectContentSlot(content);
  if (inspected.kind === 'absent') return false;
  if (inspected.kind === 'render') return true;
  return inspected.kind === 'props' ? hasContentSlotContent(content) : hasLegacyGroupContent(inspected.children);
}

function hasLegacyGroupContent(content: React.ReactNode): boolean {
  return React.Children.toArray(content).some(child => {
    if (typeof child === 'string') return child.length > 0;
    if (React.isValidElement<{ children?: React.ReactNode }>(child) && child.type === React.Fragment) {
      return hasLegacyGroupContent(child.props.children);
    }
    return true;
  });
}

/** Apply semantic text wrappers and inherited presets without parsing or cloning caller children. */
export function GroupText({ children, description = false, headingLevel, disabled = false, className }: {
  children: ContentSlot<'span'>; description?: boolean; headingLevel?: 1 | 2 | 3 | 4 | 5 | 6; disabled?: boolean; className?: string;
}): React.ReactElement {
  const sx = useSx();
  const styles = useStyles();
  const Root = headingLevel ? headings[headingLevel - 1] : 'span';
  return <Root className={sx(styles.text, marginBlock.none,
    description ? typography.caption1 : headingLevel ? typography.subtitle2 : typography.body1Strong,
    !disabled && (description ? foreground.subtle : foreground.primary), className)}>{renderContentSlot(children, 'span', sx, { shorthand: 'preserve' })}</Root>;
}

/** Matching group roots share only header semantics, text presets and layout. */
export function GroupHeader({ title, description, icon, action, headingLevel = 2 }: {
  title?: ContentSlot<'span'>; description?: ContentSlot<'span'>; icon?: ContentSlot<'span'>; action?: ContentSlot<'div'>;
  headingLevel?: 1 | 2 | 3 | 4 | 5 | 6;
}): React.ReactElement {
  const sx = useSx();
  const styles = useStyles();
  const titlePresent = hasGroupContent(title), descriptionPresent = hasGroupContent(description);
  const iconPresent = hasGroupContent(icon), actionPresent = hasGroupContent(action);
  if (!titlePresent && !descriptionPresent && !iconPresent && !actionPresent) return <></>;
  return <div data-group-header="" className={sx(flex.row, flex.wrap, alignItems.center, gap.medium, minWidth.zero)}>
    {iconPresent && <span data-group-icon="" aria-hidden="true" className={sx(styles.icon, flexItem.noShrink, foreground.subtle)}>{renderContentSlot(icon, 'span', sx, { shorthand: 'preserve', className: styles.iconContent })}</span>}
    {(titlePresent || descriptionPresent) && <div data-group-text="" className={sx(flex.column, flexItem.grow, minWidth.zero, gap.none)}>
      {titlePresent && <GroupText headingLevel={headingLevel}>{title}</GroupText>}
      {descriptionPresent && <GroupText description>{description}</GroupText>}
    </div>}
    {actionPresent && <div className={sx(styles.action, flex.row, flex.wrap, alignItems.center, flexItem.noShrink, gap.small)}>{renderContentSlot(action, 'div', sx, { shorthand: 'preserve' })}</div>}
  </div>;
}
