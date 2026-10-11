import * as React from 'react';
import { Button, Card, CardHeader } from '@fluentui/react-components';
import { Dismiss24Regular } from '@fluentui/react-icons';
import { GroupText, hasGroupContent } from '../group-header.internal';
import { useSx } from '../../hooks/useSx';
import { useStageContext } from './stage-context.internal';
import { useStageStyles } from './stage.styles.internal';
import type { StageHeaderProps } from './stage.types';

/** Render Stage's semantic heading, supporting text, decorative icon and free actions.
 * Strings remain literal; caller JSX, refs and handlers are preserved. Requires an owning
 * Stage and Fluent theme/renderer. A single nonempty phrasing title is required.
 * @param props Native div attributes and named Header slots; Close defaults to visible.
 * @returns A native div wrapping a neutral Card and CardHeader, isolated from ancestor Cards.
 * @example <Stage.Header title="Profile" description="Edit details" closeLabel="Chiudi" />
 */
export const StageHeader = React.forwardRef<HTMLDivElement, StageHeaderProps>(({
  title, description, icon, actions, headingLevel = 2, showCloseButton = true, closeLabel = 'Close', className, ...nativeProps
}, ref) => {
  // The public type forbids raw HTML, but untyped callers still reach the DOM boundary.
  const attributes: React.HTMLAttributes<HTMLDivElement> = nativeProps;
  delete attributes.dangerouslySetInnerHTML;
  const { titleId, descriptionId, registerDescription, controller } = useStageContext();
  const styles = useStageStyles(), sx = useSx();
  const descriptionPresent = hasGroupContent(description);
  React.useLayoutEffect(() => registerDescription(descriptionPresent), [descriptionPresent, registerDescription]);
  // Native activation handlers belong to the public div. Passing them to Card would enable
  // Fluent interaction styling, including forced-color selectors affecting caller descendants.
  return <div {...attributes} ref={ref} role="presentation" data-stage-header="" className={sx(styles.header, className)}>
    <Card appearance="subtle" focusMode="off" role="presentation" className={styles.neutralCard}>
      <CardHeader className={styles.cardHeader}
        image={hasGroupContent(icon) ? { 'aria-hidden': true, className: styles.icon, children: icon } : undefined}
        header={{ id: titleId, className: styles.text, children: <GroupText headingLevel={headingLevel}>{title}</GroupText> }}
        description={descriptionPresent ? { id: descriptionId, className: styles.description, children: <GroupText description>{description}</GroupText> } : undefined}
        action={hasGroupContent(actions) || showCloseButton ? { className: styles.actions, children: <>
          {actions}
          {showCloseButton && <Button appearance="subtle" aria-label={closeLabel} icon={<Dismiss24Regular />}
            onClick={event => controller.requestOpenChange(event, { open: false, reason: 'closeButton' })} />}
        </> } : undefined} />
    </Card>
  </div>;
});
StageHeader.displayName = 'Stage.Header';
