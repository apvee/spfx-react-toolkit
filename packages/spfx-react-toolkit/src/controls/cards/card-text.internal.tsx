import * as React from 'react';
import { inspectContentSlot, renderContentSlot } from '../content-slot.internal';
import type { ContentSlotClassComposer } from '../content-slot.internal';
import type { ContentSlot } from '../content-slot.types.internal';
import type { ContentCardHeadingLevel } from '../content-card/content-card.types';
import * as typography from '../../helpers/styles/typography';
import * as foreground from '../../helpers/styles/foreground';
import { useSx } from '../../hooks/useSx';
import { useCardStyles } from './card.styles.internal';

/** Strings use existing typography descriptors without retaining the Markdown parser. */
export function CardText({ value, headingLevel, disabled = false }: {
  value: ContentSlot<'span'> | ContentSlot<'div'>; headingLevel?: ContentCardHeadingLevel; disabled?: boolean;
}): React.ReactElement {
  const sx = useSx(); const styles = useCardStyles();
  const inspection = inspectContentSlot(value);
  const literal = (inspection.kind === 'props' || inspection.kind === 'shorthand') && typeof inspection.children === 'string';
  const headings = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'] as const;
  const Root = headingLevel ? headings[headingLevel - 1] : 'div';
  return <Root className={sx(styles.text, headingLevel ? typography.body1Strong : typography.body1,
    headingLevel ? foreground.primary : disabled ? undefined : foreground.subtle,
    disabled && !headingLevel && styles.disabled, !headingLevel && styles.description, literal && styles.clamp)}>
    {renderContentSlot(value, headingLevel ? 'span' : 'div', sx, { shorthand: 'preserve' })}
  </Root>;
}

/** Generic regions retain semantic shells even when an opaque callback returns no DOM. */
export function renderCardRegion(value: unknown, tag: 'div' | 'span', sx: ContentSlotClassComposer, className?: string): React.ReactNode {
  return renderContentSlot(value, tag, sx, { shorthand: 'preserve', className });
}
