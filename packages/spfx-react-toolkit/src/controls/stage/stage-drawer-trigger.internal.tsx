import * as React from 'react';
import { Button, useRestoreFocusTarget } from '@fluentui/react-components';
import { getReactElementRef, useMergedRefs } from '@fluentui/react-utilities';
import { useStageContext } from './stage-context.internal';

/** Drawer has no native trigger; compose only the documented semantic button contract. */
export function StageDrawerTrigger({ trigger, triggerRef }: { trigger: React.ReactElement; triggerRef?: React.Ref<HTMLButtonElement> }): React.ReactElement {
  const { controller, surfaceId } = useStageContext();
  const restoreTarget = useRestoreFocusTarget();
  const ref = useMergedRefs(getReactElementRef<HTMLButtonElement>(trigger), triggerRef);
  if (trigger.type !== 'button' && trigger.type !== Button) {
    throw new Error('Stage trigger must be a native button or Fluent Button.');
  }
  const button: React.ReactElement<React.ButtonHTMLAttributes<HTMLButtonElement> & React.RefAttributes<HTMLButtonElement>> = trigger;
  return React.cloneElement(button, {
    ...restoreTarget, ref,
    type: button.props.type ?? 'button',
    'aria-controls': surfaceId, 'aria-haspopup': 'dialog', 'aria-expanded': controller.open,
    onClick: (event: React.MouseEvent<HTMLButtonElement>) => {
      button.props.onClick?.(event);
      if (!button.props.disabled && button.props['aria-disabled'] !== true && button.props['aria-disabled'] !== 'true' && !event.defaultPrevented) {
        controller.requestOpenChange(event, { open: true, reason: 'triggerClick' });
      }
    },
  });
}
