import * as React from 'react';
import { Stage as rootStage } from '@apvee/spfx-react-toolkit';
import { Stage as domainStage } from '@apvee/spfx-react-toolkit/controls';
import { Stage as controlStage } from '@apvee/spfx-react-toolkit/controls/stage';
import type { StageProps as RootProps } from '@apvee/spfx-react-toolkit';
import type { StageProps as DomainProps } from '@apvee/spfx-react-toolkit/controls';
import type { StageProps as ControlProps } from '@apvee/spfx-react-toolkit/controls/stage';

export const aliases: typeof rootStage[] = [domainStage, controlStage];
export const props: RootProps = { presentation: 'dialog', size: 'full', children: 'Content' };
export const domainProps: DomainProps = props;
export const controlProps: ControlProps = props;
const Stage = controlStage;
export const element = <Stage {...props} ref={React.createRef<HTMLDivElement>()}>
  <Stage.Header title="Profile" description="Details" action={<button>Refresh</button>} closeLabel="Close" />
  <Stage.Body ref={React.createRef<HTMLDivElement>()}><input aria-label="Name" /></Stage.Body>
  <Stage.Footer><button type="button">Cancel</button></Stage.Footer>
</Stage>;
<Stage presentation="popover" trigger={<button>Open</button>}><Stage.Header title="Context" /><Stage.Body>Context</Stage.Body></Stage>;
// @ts-expect-error Popover remains an anchored surface with bounded sizes.
<Stage presentation="popover" size="full" trigger={<button>Open</button>}>Content</Stage>;
// @ts-expect-error Popover requires its real trigger.
<Stage presentation="popover">Content</Stage>;
// @ts-expect-error State ownership cannot mix controlled and initial open.
<Stage presentation="dialog" open defaultOpen>Content</Stage>;
// @ts-expect-error Private implementation modules are not public aliases.
import {} from '@apvee/spfx-react-toolkit/controls/stage/stage-context.internal';
