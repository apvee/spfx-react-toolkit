import * as React from 'react';
import { useStageController } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-controller.internal';
import { StageContext, useStageContextValue } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-context.internal';
import { StageDialog } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-dialog.internal';
import { StageDrawer } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-drawer.internal';
import { StageHeader } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-header';
import { StageBody } from '../../packages/spfx-react-toolkit/src/controls/stage/stage-body';
import type { StageProps, StageSurfaceAttributes } from '../../packages/spfx-react-toolkit/src/controls/stage/stage.types';

/** Single bundle preserves the actual private controller/context identities. */
export function ModalFixture({ options, surfaceProps = {}, surfaceRef, description, children }: {
  options: StageProps;
  surfaceProps?: StageSurfaceAttributes;
  surfaceRef?: React.ForwardedRef<HTMLDivElement>;
  description?: React.ReactNode;
  children?: React.ReactNode;
}): React.ReactElement {
  const controller = useStageController(options);
  const value = useStageContextValue(controller, surfaceProps.id);
  const Adapter = controller.presentation === 'drawer' ? StageDrawer : StageDialog;
  return <StageContext.Provider value={value}>
    <Adapter controller={controller} options={options} surfaceProps={surfaceProps} surfaceRef={surfaceRef ?? (() => {})}>
      <StageHeader title="Editor" description={description} />
      <StageBody>{children ?? <input aria-label="Name" defaultValue="Initial" />}</StageBody>
    </Adapter>
  </StageContext.Provider>;
}
