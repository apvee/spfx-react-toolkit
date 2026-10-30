import * as React from 'react';
import { Button, Popover, PopoverTrigger, PopoverSurface } from '@fluentui/react-components';
import type { PopoverProps } from '@fluentui/react-components';
import { useSx } from '../../hooks/useSx';
import { useStageContext } from './stage-context.internal';
import { useStageStyles } from './stage.styles.internal';
import type { StageAdapterProps } from './stage-controller.internal';

/** Controlled nonmodal adapter. Native positioning, motion, focus and portals remain Fluent-owned.
 * Stage owns trigger controls/expanded/popup associations: a group surface has no aria-haspopup.
 * Caller refs, handlers, disabled state, accessible names and presentation props remain intact. */
export function StagePopover({ controller, options, surfaceProps, surfaceRef, children }: StageAdapterProps): React.ReactElement {
  const context = useStageContext();
  const styles = useStageStyles();
  const sx = useSx();
  const [hasOpened, setHasOpened] = React.useState(controller.open);
  const [exited, setExited] = React.useState(!controller.open);
  const [motionKey, setMotionKey] = React.useState(0);
  const previousUnmount = React.useRef(options.unmountOnClose !== false);
  React.useLayoutEffect(() => {
    const nextUnmount = options.unmountOnClose !== false;
    // Native presence retains its mounted state on policy-only updates. Reset its
    // instance only when retention is disabled after an already-completed exit.
    if (!previousUnmount.current && nextUnmount && !controller.open && exited) setMotionKey(key => key + 1);
    previousUnmount.current = nextUnmount;
    if (controller.open) setExited(false);
  }, [options.unmountOnClose, controller.open, exited]);
  const consumedEscapeEvent = React.useRef<React.KeyboardEvent<HTMLDivElement> | undefined>(undefined);
  if (controller.open && !hasOpened) setHasOpened(true);

  const trigger = options.trigger;
  if (!React.isValidElement<React.ButtonHTMLAttributes<HTMLButtonElement>>(trigger) ||
      (trigger.type !== 'button' && trigger.type !== Button)) {
    throw new Error('Stage Popover requires a native button or Fluent Button trigger.');
  }
  if (controller.size === 'full') throw new Error('Stage Popover does not support full size.');
  if (options.presentation !== 'popover') throw new Error('Stage Popover requires compatible Popover options.');

  const onOpenChange: NonNullable<PopoverProps['onOpenChange']> = (event, data) => {
    // Only skip the exact Escape a descendant consumed before the native handler.
    // Ordinary native Popover Escape is already prevented and must still be delivered.
    if (event === consumedEscapeEvent.current) return;
    const reason = event.type === 'keydown' ? 'escapeKeyDown' :
      event.type === 'focus' || event.type === 'focusin' || event.type === 'fuiframefocus' ? 'focusOutside' :
      event.type === 'click' && 'nativeEvent' in event ? 'triggerClick' : 'outsideClick';
    controller.requestOpenChange(event, { open: data.open, reason });
  };
  const sizeClass = controller.size === 'small' ? styles.popoverSmall :
    controller.size === 'medium' ? styles.popoverMedium : styles.popoverLarge;
  const closed = !controller.open;
  // Native presence motion normally mounts immediately when retention is enabled.
  // Gate it until first open so initially closed children have no mount side effects.
  const retain = options.unmountOnClose === false && hasOpened;

  return <Popover open={controller.open} onOpenChange={onOpenChange} trapFocus={false} size="small" withArrow={options.withArrow}
    mountNode={options.mountNode} surfaceMotion={{
      children: (Motion, motionProps) => <Motion {...motionProps} key={motionKey} unmountOnExit={!retain} />,
      onMotionFinish: (_, data) => { if (data.direction === 'exit') setExited(true); },
    }}
    positioning={{ position: options.positioning?.position ?? 'below', align: options.positioning?.align ?? 'start', autoSize: true }}>
    <PopoverTrigger disableButtonEnhancement>
      {React.cloneElement(trigger, { type: trigger.props.type ?? 'button', 'aria-controls': context.surfaceId,
        'aria-expanded': controller.open, 'aria-haspopup': undefined })}
    </PopoverTrigger>
    <PopoverSurface {...surfaceProps} ref={surfaceRef} id={context.surfaceId}
      className={sx(styles.surface, styles.popover, sizeClass, styles.popoverPadding,
        options.withArrow ? styles.popoverArrow : undefined, surfaceProps.className)}
      aria-labelledby={context.titleId} aria-describedby={context.hasDescription ? context.descriptionId : undefined}
      hidden={closed || surfaceProps.hidden} aria-hidden={closed ? true : surfaceProps['aria-hidden']}>
      {/* The installed Fluent native-prop allowlist predates inert. The public
          root slot preserves native children (including the arrow), refs and handlers. */}
      {(Component, slotProps) => <Component {...slotProps} {...{ inert: closed ? '' : undefined }} onKeyDown={(event: React.KeyboardEvent<HTMLDivElement>) => {
        // This public root slot runs after descendants and before Fluent's surface
        // handler prevents Escape. Forward that handler, including the caller callback.
        consumedEscapeEvent.current = event.key === 'Escape' && event.defaultPrevented ? event : undefined;
        try { slotProps.onKeyDown?.(event); }
        finally { consumedEscapeEvent.current = undefined; }
      }}>{slotProps.children}<div className={styles.popoverContent}>{children}</div></Component>}
    </PopoverSurface>
  </Popover>;
}
