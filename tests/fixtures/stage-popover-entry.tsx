import * as React from 'react';
import { useStageController } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-controller.internal';
import { StageContext, useStageContextValue } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-context.internal';
import { StagePopover } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-popover.internal';
import type { StageProps, StageSurfaceAttributes } from '../../packages/spfx-react-toolkit/src/controls/stage/stage.types';

/** A single compilation unit keeps real controller, context, parts and adapter identities shared. */
export function PopoverOwner({ options, surfaceProps = {}, surfaceRef, children }: {
  options: StageProps;
  surfaceProps?: StageSurfaceAttributes;
  surfaceRef?: React.ForwardedRef<HTMLDivElement>;
  children?: React.ReactNode;
}): React.ReactElement {
  const controller = useStageController(options);
  const value = useStageContextValue(controller, surfaceProps.id);
  // React 17 ForwardedRef represents an absent forwarded ref with null.
  const ref = surfaceRef ?? null;
  return <StageContext.Provider value={value}>
    <StagePopover controller={controller} options={options} surfaceProps={surfaceProps}
      surfaceRef={ref} children={children ?? options.children} />
  </StageContext.Provider>;
}

export { StageHeader } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-header';
export { StageBody } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-body';
export { StageFooter } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-footer';
