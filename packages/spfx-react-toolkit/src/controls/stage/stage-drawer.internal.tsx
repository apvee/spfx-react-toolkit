import * as React from 'react';
import { OverlayDrawer, useRestoreFocusSource } from '@fluentui/react-components';
import { useFluent_unstable } from '@fluentui/react-shared-contexts';
import { useMergedRefs } from '@fluentui/react-utilities';
import { useSx } from '../../hooks/useSx';
import { useStageContext } from './stage-context.internal';
import { StageDrawerTrigger } from './stage-drawer-trigger.internal';
import { useStageStyles } from './stage.styles.internal';
import type { StageAdapterProps } from './stage-controller.internal';

/** Controlled OverlayDrawer uses the host direction, native portal/focus and motion. */
export function StageDrawer({ controller, options, surfaceProps, surfaceRef, children }: StageAdapterProps): React.ReactElement {
  const { surfaceId, titleId, descriptionId, hasDescription } = useStageContext();
  const { dir } = useFluent_unstable();
  const restoreSource = useRestoreFocusSource();
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
  // React 17 native trigger refs use null for detach; internal absence is undefined.
  const captureTrigger = React.useCallback((node: HTMLButtonElement | null) => { triggerElement.current = node ?? undefined; }, []);
  React.useLayoutEffect(() => {
    const nextUnmount = options.unmountOnClose !== false;
    // Native presence ignores policy-only updates after a completed exit. Reset
    // that native motion instance only for disabling already-closed retention.
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
    // Fluent filters inert out of native surface props. Synchronize at attachment
    // and open changes without replacing focus behavior or interrupting exit motion.
    // Preserve native source focus during default-unmount exit; applying inert
    // before removal prevents Tabster from restoring after a mouse close.
    if (controller.open || options.unmountOnClose !== false) node?.removeAttribute('inert');
    else node?.setAttribute('inert', '');
  }, [controller.open, options.unmountOnClose]);
  const restoreRetainedFocus = (): void => {
    const target = triggerElement.current, owner = sourceDocument.current;
    const frame = owner?.defaultView?.frameElement;
    // Native mouse restoration skips connected retained sources. Repair only lost
    // focus from this source to its own body, with a connected enabled native target.
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
  const unmountOnClose = options.unmountOnClose !== false || !hasOpened;
  const logicalPosition = options.presentation === 'drawer' ? options.drawerPosition ?? 'end' : 'end';
  // Fluent's edge classes and motion follow provider dir. Map a native dir override
  // into its logical position rather than authoring physical inline positioning.
  const nativeDir = surfaceProps.dir === 'rtl' || surfaceProps.dir === 'ltr' ? surfaceProps.dir : dir;
  const position = nativeDir === dir ? logicalPosition : logicalPosition === 'start' ? 'end' : 'start';
  const sizeClass = { small: styles.drawerSmall, medium: styles.drawerMedium, large: styles.drawerLarge, full: styles.drawerFull }[controller.size];
  const attributes: React.HTMLAttributes<HTMLDivElement> = { ...surfaceProps };
  delete attributes.dangerouslySetInnerHTML;
  delete attributes.role;
  delete attributes['aria-modal'];
  delete attributes['aria-labelledby'];
  delete attributes['aria-describedby'];
  return <>
    {options.trigger && <StageDrawerTrigger trigger={options.trigger} triggerRef={captureTrigger} />}
    <OverlayDrawer {...attributes} {...restoreSource} ref={ref} mountNode={options.mountNode} id={surfaceId}
      open={controller.open} size={controller.size} position={position}
      modalType={controller.open || !exited ? 'modal' : 'non-modal'} unmountOnClose={unmountOnClose}
      role="dialog" aria-modal="true" aria-labelledby={titleId} aria-describedby={hasDescription ? descriptionId : undefined}
      aria-hidden={!controller.open && options.unmountOnClose === false ? true : attributes['aria-hidden']}
      // Keep its own native fade; remove its hit-test target only after that exit.
      backdrop={{ hidden: !controller.open && backdropExited, className: styles.backdrop }}
      backdropMotion={{ onMotionFinish: (_, data) => { if (data.direction === 'exit') setBackdropExited(true); } }}
      hidden={attributes.hidden || (!controller.open && exited)}
      className={sx(styles.surface, styles.drawer, sizeClass, styles.modalPadding, attributes.className)}
      surfaceMotion={{
        children: (Motion, motionProps) => <Motion {...motionProps} key={motionKey} />,
        onMotionFinish: (_, data) => { if (data.direction === 'exit') { setExited(true); restoreRetainedFocus(); } },
      }}
      onOpenChange={(event, data) => controller.requestOpenChange(event, { open: data.open, reason: data.type })}>
      {children}
    </OverlayDrawer>
  </>;
}
