import * as React from 'react';
import { inspectContentSlot, renderContentSlot } from '../content-slot.internal';
import { Button, Card, CardHeader } from '@fluentui/react-components';
import { Dismiss24Regular } from '@fluentui/react-icons';
import { GroupText, hasGroupContent, hasGroupDescriptionContent } from '../group-header.internal';
import { useSx } from '../../hooks/useSx';
import * as text from '../../helpers/styles/text';
import { useStageContext } from './stage-context.internal';
import { useStageStyles } from './stage.styles.internal';
import type { StageHeaderProps } from './stage.types';
import { nativeStageResource } from './localization/stage-native.internal';

/** Render Stage's semantic heading, supporting text, decorative icon and caller action.
 * Strings remain literal; caller JSX, refs and handlers are preserved. Requires an owning
 * Stage and Fluent theme/renderer. A single nonempty phrasing title is required.
 * Direct strings and props-object string children use one title line and two description
 * lines with ellipsis in every presentation. Full text and naming IDs remain in the DOM;
 * JSX and render callbacks own their presentation. Close remains separate from action.
 * Action and Close share one row and are vertically centered with the text block,
 * including narrow Dialog, Drawer and Popover headers.
 * @param props Native div attributes and named Header slots; Close defaults to visible.
 * @returns A native div wrapping a neutral Card and CardHeader, isolated from ancestor Cards.
 * @example
 * ```tsx
 * <Stage presentation="dialog"><Stage.Header title="Profile" description="Edit details" closeLabel="Chiudi" /></Stage>
 * ```
 */
export const StageHeader = React.forwardRef<HTMLDivElement, StageHeaderProps>(({
  title, description, icon, action, headingLevel = 2, showCloseButton = true, closeLabel, strings, className, ...nativeProps
}, ref) => {
  // The public type forbids raw HTML, but untyped callers still reach the DOM boundary.
  const attributes: React.HTMLAttributes<HTMLDivElement> = nativeProps;
  delete attributes.dangerouslySetInnerHTML;
  // The removed plural API is unsupported, including at the untyped DOM boundary.
  delete (attributes as React.HTMLAttributes<HTMLDivElement> & { actions?: unknown }).actions;
  const { titleId, descriptionId, registerDescription, controller } = useStageContext();
  const styles = useStageStyles(), sx = useSx();
  const descriptionPresent = hasGroupContent(description);
  const descriptionMeaningful = hasGroupDescriptionContent(description);
  const titleSlot = inspectContentSlot(title), descriptionSlot = inspectContentSlot(description);
  const literalTitle = (titleSlot.kind === 'props' || titleSlot.kind === 'shorthand') && typeof titleSlot.children === 'string';
  const literalDescription = (descriptionSlot.kind === 'props' || descriptionSlot.kind === 'shorthand') && typeof descriptionSlot.children === 'string';
  React.useLayoutEffect(() => registerDescription(descriptionMeaningful), [descriptionMeaningful, registerDescription]);
  // Native activation handlers belong to the public div. Passing them to Card would enable
  // Fluent interaction styling, including forced-color selectors affecting caller descendants.
  return <div {...attributes} ref={ref} role="presentation" data-stage-header="" className={sx(styles.header, className)}>
    <Card appearance="subtle" focusMode="off" role="presentation" className={styles.neutralCard}>
      <CardHeader className={sx(styles.cardHeader, showCloseButton && styles.withClose)}
        image={hasGroupContent(icon) ? { 'aria-hidden': true, className: styles.icon, children: renderContentSlot(icon, 'span', sx, { shorthand: 'preserve', className: styles.iconContent }) } : undefined}
        header={{ id: titleId, className: styles.text, children: <GroupText headingLevel={headingLevel} className={literalTitle ? sx(text.truncate) : undefined}>{title}</GroupText> }}
        description={descriptionPresent ? { id: descriptionId, className: styles.description, children: <GroupText description className={literalDescription ? styles.literalDescription : undefined}>{description}</GroupText> } : undefined}
        action={hasGroupContent(action) ? { className: styles.actions, children:
          renderContentSlot(action, 'div', sx, { shorthand: 'preserve', className: styles.actionContent })
        } : undefined} />
      {showCloseButton && <Button data-stage-close="" className={styles.close} appearance="subtle" aria-label={closeLabel ?? strings?.closeLabel ?? nativeStageResource.closeLabel} icon={<Dismiss24Regular />}
        onClick={event => controller.requestOpenChange(event, { open: false, reason: 'closeButton' })} />}
    </Card>
  </div>;
});
StageHeader.displayName = 'Stage.Header';
