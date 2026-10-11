import * as React from 'react';
import type { StageOpenChangeData, StageOpenChangeEvent, StagePresentation, StageProps, StageSize, StageSurfaceAttributes } from './stage.types';

/** Private authoritative state passed to every controlled Fluent adapter. */
export interface StageController {
  open: boolean;
  presentation: StagePresentation;
  size: StageSize;
  requestOpenChange(event: StageOpenChangeEvent, data: StageOpenChangeData): void;
}

/** Private adapter boundary: root separates allowed native attributes from original options. */
export interface StageAdapterProps {
  controller: StageController;
  options: StageProps;
  surfaceProps: StageSurfaceAttributes;
  surfaceRef: React.ForwardedRef<HTMLDivElement>;
  children: React.ReactNode;
}

interface StageSession {
  open: boolean;
  presentation: StagePresentation;
  /** Last compatible size prevents a pending full Drawer from making a latched Popover full. */
  size: StageSize;
}

/** Maintain state ownership and a presentation session; no native event-prevention veto. */
export function useStageController(props: StageProps): StageController {
  if (props.presentation !== 'dialog' && props.presentation !== 'drawer' && props.presentation !== 'popover') {
    throw new Error('Stage presentation must be dialog, drawer or popover.');
  }
  const requestedSize = props.size === undefined ? 'small' : props.size;
  if (requestedSize !== 'small' && requestedSize !== 'medium' && requestedSize !== 'large' && requestedSize !== 'full') {
    throw new Error('Stage size must be small, medium, large or full.');
  }
  if (props.presentation === 'popover' && requestedSize === 'full') {
    throw new Error('Stage Popover does not support full size.');
  }

  const controlled = React.useRef(props.open !== undefined).current;
  const [uncontrolledOpen, setUncontrolledOpen] = React.useState(() => props.defaultOpen ?? false);
  if (controlled !== (props.open !== undefined)) {
    throw new Error('Stage cannot change between controlled and uncontrolled state during its lifetime.');
  }
  const open = props.open ?? uncontrolledOpen;
  const [session, setSession] = React.useState<StageSession>(() => ({ open, presentation: props.presentation, size: requestedSize }));
  const presentation = open && session.open ? session.presentation : props.presentation;
  const size = presentation === 'popover' && requestedSize === 'full' ? session.size : requestedSize;

  // Adjust only this component's session during render so descendants never commit an
  // intermediate adapter on opening. A refused controlled close leaves the session intact.
  if (session.open !== open || session.presentation !== presentation || session.size !== size) {
    setSession({ open, presentation, size });
  }

  const requestOpenChange = React.useCallback((event: StageOpenChangeEvent, data: StageOpenChangeData): void => {
    if (props.dismissMode === 'explicit' && data.reason !== 'triggerClick' && data.reason !== 'closeButton') return;
    // Consumer trigger cancellation is distinct from Fluent's already-prevented Escape.
    if (data.reason === 'triggerClick' && event.defaultPrevented) return;
    if (!controlled) setUncontrolledOpen(data.open);
    props.onOpenChange?.(event, data);
  }, [controlled, props.dismissMode, props.onOpenChange]);

  return { open, presentation, size, requestOpenChange };
}
