import * as React from 'react';
import { Button, Dialog, DialogSurface, DialogTrigger } from '@fluentui/react-components';
import { getReactElementRef, useMergedRefs } from '@fluentui/react-utilities';
import { useSx } from '../../hooks/useSx';
import { useStageContext } from './stage-context.internal';
import { useStageStyles } from './stage.styles.internal';
import type { StageAdapterProps } from './stage-controller.internal';

/** Controlled native modal adapter; the Dialog owns focus, portal and close motion. */
export function StageDialog({ controller, options, surfaceProps, surfaceRef, children }: StageAdapterProps): React.ReactElement {
  const { surfaceId, titleId, descriptionId, hasDescription } = useStageContext();
  const styles = useStageStyles(), sx = useSx();
  const [hasOpened, setHasOpened] = React.useState(controller.open);
  const [exited, setExited] = React.useState(!controller.open);
  const [backdropExited, setBackdropExited] = React.useState(!controller.open);
  const [motionKey, setMotionKey] = React.useState(0);
  const previousUnmount = React.useRef(options.unmountOnClose !== false);
  const previousSurfaceOpen = React.useRef(controller.open);
  const sourceHadFocus = React.useRef(false);
  const sourceDocument = React.useRef<Document | undefined>(undefined);
  const triggerElement = React.useRef<HTMLButtonElement | undefined>(undefined);
  // Native semantic trigger callbacks use null for detach in React 17.
  const captureTrigger = React.useCallback((node: HTMLButtonElement | null) => { triggerElement.current = node ?? undefined; }, []);
  const triggerRef = useMergedRefs(getReactElementRef<HTMLButtonElement>(options.trigger), captureTrigger);
  React.useLayoutEffect(() => {
    const nextUnmount = options.unmountOnClose !== false;
    // Fluent's mounted state does not react to an unmountOnExit-only update. Reset
    // its native motion only after exit completed and retention was then disabled.
    if (!previousUnmount.current && nextUnmount && !controller.open && exited) setMotionKey(key => key + 1);
    previousUnmount.current = nextUnmount;
  }, [options.unmountOnClose, controller.open, exited]);
  // React 17 requires null in the native ref callback's detach contract.
  const captureSurface = React.useCallback((node: HTMLDivElement | null) => {
    if (node) {
      sourceDocument.current = node.ownerDocument;
      if (!controller.open && previousSurfaceOpen.current && options.unmountOnClose === false) {
        sourceHadFocus.current = node.contains(node.ownerDocument.activeElement);
      }
      if (controller.open) sourceHadFocus.current = false;
      previousSurfaceOpen.current = controller.open;
    }
    // Fluent filters inert out of native surface props. Synchronize the native div
    // at attachment too, so a changed portal destination cannot lose this attribute.
    // Default unmount needs focus to remain inside the native source until removal:
    // early inert would clear Tabster's restore history during mouse dismissal.
    if (controller.open || options.unmountOnClose !== false) node?.removeAttribute('inert');
    else node?.setAttribute('inert', '');
  }, [controller.open, options.unmountOnClose]);
  const restoreRetainedFocus = (): void => {
    const target = triggerElement.current, owner = sourceDocument.current;
    const frame = owner?.defaultView?.frameElement;
    // Native Restorer skips a still-connected source after mouse focus loss. Only
    // repair that retained-source gap; never redirect focus from another control/document.
    if (!controller.open && options.unmountOnClose === false && sourceHadFocus.current && owner &&
      owner.activeElement === owner.body && (!frame || frame.ownerDocument.activeElement === frame) &&
      target?.isConnected && target.ownerDocument === owner &&
      !target.disabled && target.getAttribute('aria-disabled') !== 'true') target.focus();
    sourceHadFocus.current = false;
  };
  const ref = useMergedRefs(surfaceRef, captureSurface);
  React.useLayoutEffect(() => {
    if (controller.open) { setHasOpened(true); setExited(false); setBackdropExited(false); }
  }, [controller.open]);
  // Native retention would mount an initially closed surface. Until the first opening,
  // retain the native unmount behavior without adding a second content wrapper.
  const unmountOnClose = options.unmountOnClose !== false || !hasOpened;
  const sizeClass = { small: styles.dialogSmall, medium: styles.dialogMedium, large: styles.dialogLarge, full: styles.dialogFull }[controller.size];
  const attributes: React.HTMLAttributes<HTMLDivElement> = { ...surfaceProps };
  delete attributes.dangerouslySetInnerHTML;
  delete attributes.role;
  delete attributes['aria-modal'];
  delete attributes['aria-labelledby'];
  delete attributes['aria-describedby'];
  const trigger = options.trigger;
  if (trigger && trigger.type !== 'button' && trigger.type !== Button) {
    throw new Error('Stage trigger must be a native button or Fluent Button.');
  }
  const surface = <DialogSurface key="surface" {...attributes} ref={ref} mountNode={options.mountNode} id={surfaceId}
      role="dialog" aria-modal="true" aria-labelledby={titleId} aria-describedby={hasDescription ? descriptionId : undefined}
      aria-hidden={!controller.open && options.unmountOnClose === false ? true : attributes['aria-hidden']}
      // Hide only after the backdrop's own native fade completes, independently of
      // retained content and subsequent unmount-policy changes.
      backdrop={{ hidden: !controller.open && backdropExited, className: styles.backdrop }}
      backdropMotion={{ onMotionFinish: (_, data) => { if (data.direction === 'exit') setBackdropExited(true); } }}
      hidden={attributes.hidden || (!controller.open && exited)}
      className={sx(styles.surface, styles.dialog, sizeClass, styles.modalPadding, attributes.className)}>
      {children}
    </DialogSurface>;
  // A retained, fully exited surface must release native modalizer/trap attributes.
  // Keep modal behavior throughout native closing motion and restore it on opening.
  return <Dialog open={controller.open} modalType={controller.open || !exited ? 'modal' : 'non-modal'} unmountOnClose={unmountOnClose}
    onOpenChange={(event, data) => controller.requestOpenChange(event, { open: data.open, reason: data.type })}
    surfaceMotion={{
      children: (Motion, motionProps) => <Motion {...motionProps} key={motionKey} />,
      onMotionFinish: (_, data) => { if (data.direction === 'exit') { setExited(true); restoreRetainedFocus(); } },
    }}>
    {trigger ? [<DialogTrigger key="trigger">{React.cloneElement(trigger, {
      ref: triggerRef, 'aria-controls': surfaceId, 'aria-haspopup': 'dialog', 'aria-expanded': controller.open,
    })}</DialogTrigger>, surface] : surface}
  </Dialog>;
}
