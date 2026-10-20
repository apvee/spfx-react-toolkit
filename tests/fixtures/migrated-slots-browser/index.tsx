import { mountDemoFixture } from '../demo-fixture-host';
import * as React from 'react';
import { registerIcons } from '@fluentui/react';
import GroupsPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/GroupsPanel';
import StateMessagePanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/StateMessagePanel';
import RelativeTimePanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/RelativeTimePanel';
import FormattedTextPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/FormattedTextPanel';
import AutoSkeletonPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/AutoSkeletonPanel';
import FormPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/FormPanel';
// Host substitutes local v8 icons, matching the established sample browser fixture.
const icon = <svg width="1em" height="1em" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M4 5h16v14H4zm2 2v10h12V7z" /></svg>;
registerIcons({ icons: { Settings: icon, Mail: icon, Ringer: icon, Info: icon, Clock: icon } });
const panels = { groups: GroupsPanel, state: StateMessagePanel, relative: RelativeTimePanel, form: FormPanel, formatted: FormattedTextPanel, skeleton: AutoSkeletonPanel };
const key = new URLSearchParams(location.search).get('panel') || 'groups';
const Panel = panels[key as keyof typeof panels] || GroupsPanel;
mountDemoFixture(<Panel />);
