import * as React from 'react';
import { Card, CardHeader, CardFooter, CardPreview } from '@fluentui/react-components';
import { hasContentSlot, hasContentSlotContent, renderContentSlot } from '../content-slot.internal';
import type { PreviewCardProps } from '../preview-card/preview-card.types';
import { useSx } from '../../hooks/useSx';
import * as foreground from '../../helpers/styles/foreground';
import { CardText, renderCardRegion } from './card-text.internal';
import { useCardStyles } from './card.styles.internal';

type CardFrameProps = PreviewCardProps & { previewRequested?: boolean; visual?: React.ReactNode };
/** Canonical private card composition; stable detail ancestry preserves caller control state. */
export const CardFrame = React.forwardRef<HTMLDivElement, CardFrameProps>(({
  title, headingLevel = 3, description, icon, headerActions, children, footer, actions,
  size = 'medium', appearance = 'filled', orientation = 'vertical', disabled,
  logo, previewRequested = false, visual, className, ...rootProps
}, ref) => {
  const sx = useSx(); const styles = useCardStyles();
  if (!hasContentSlotContent(title)) throw new TypeError('Card title must contain meaningful phrasing content or an accessible name.');
  const safeRootProps = { ...rootProps };
  // Owned roots retain native interaction but never accept template replacement or implicit routing.
  for (const key of ['as', 'dangerouslySetInnerHTML', 'href', 'image', 'preview', 'previewFallback']) delete (safeRootProps as Record<string, unknown>)[key];
  const headerAction = hasContentSlot(headerActions) ? <div data-card-header-actions="" className={sx(styles.actions)}>{renderCardRegion(headerActions, 'div', sx)}</div> : undefined;
  const footerAction = hasContentSlot(actions) ? <div data-card-actions="" className={sx(styles.actions)}>{renderCardRegion(actions, 'div', sx)}</div> : undefined;
  return <Card {...safeRootProps} ref={ref} size={size} appearance={appearance} orientation={orientation} disabled={disabled} className={sx(styles.root, className)}>
    {previewRequested && <CardPreview data-card-preview="" className={sx(styles.preview, orientation === 'horizontal' ? styles.previewHorizontal : styles.previewVertical)}
      logo={hasContentSlot(logo) ? { children: renderContentSlot(logo, 'div', sx, { shorthand: 'preserve', allowedTags: ['img'], className: styles.logoContent }) } : undefined}>{visual}</CardPreview>}
    <div data-card-detail="" className={sx(styles.detail, orientation === 'horizontal' && styles.detailHorizontal)}>
      <CardHeader header={<CardText value={title} headingLevel={headingLevel} disabled={disabled} />}
        description={hasContentSlot(description) ? <CardText value={description} disabled={disabled} /> : undefined}
        image={hasContentSlot(icon) ? <span aria-hidden="true" className={sx(styles.icon)}>{renderCardRegion(icon, 'span', sx)}</span> : undefined}
        action={headerAction} />
      {children !== undefined && <div data-card-body="">{children}</div>}
      {(hasContentSlot(footer) || footerAction !== undefined) && <CardFooter className={sx(styles.footer)} action={footerAction}>
        {hasContentSlot(footer) && <div data-card-footer="" className={sx(foreground.subtle)}>{renderCardRegion(footer, 'div', sx)}</div>}
      </CardFooter>}
    </div>
  </Card>;
});
CardFrame.displayName = 'CardFrame';
