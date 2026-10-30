import { mountDemoFixture } from '../demo-fixture-host';
import * as React from 'react';
import { registerIcons } from '@fluentui/react';
import { PnPSearchSuggestionsDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/demos/PnPSearchSuggestionsDemo';
import StylesPanel from '../foundation-legacy/StylesPanel';
import ProductStylesPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/StylesPanel';
import RelativeTimePanel from '../foundation-legacy/RelativeTimePanel';
import ProductRelativeTimePanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/RelativeTimePanel';
import FormattedTextPanel from '../foundation-legacy/FormattedTextPanel';
import ProductFormattedTextPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/FormattedTextPanel';
import AutoSkeletonPanel from '../foundation-legacy/AutoSkeletonPanel';
import ProductAutoSkeletonPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/AutoSkeletonPanel';
import StateMessagePanel from '../foundation-legacy/StateMessagePanel';
import ProductStateMessagePanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/StateMessagePanel';
import FormPanel from '../demo-control-isolation/form-regressions';
import GroupsPanel from '../demo-control-isolation/groups-regressions';
import ProductImportsPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/ImportsPanel';
const groupFixtureIcon = <svg width="1em" height="1em" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M4 5h16v14H4zm2 2v10h12V7z" /></svg>;
registerIcons({ icons: { Color: 'color', Font: 'font', Clock: 'clock', Processing: 'loading', Info: groupFixtureIcon, Settings: groupFixtureIcon, Mail: groupFixtureIcon, Ringer: groupFixtureIcon } });
const products: Readonly<Record<string, React.ComponentType>> = {
  styles: ProductStylesPanel, 'relative-time': ProductRelativeTimePanel,
  'formatted-text': ProductFormattedTextPanel, 'auto-skeleton': ProductAutoSkeletonPanel,
  'state-message': ProductStateMessagePanel, imports: ProductImportsPanel,
};
const query = new URLSearchParams(window.location.search);
const Product = products[query.get('foundation-product') ?? ''];
mountDemoFixture(Product ? <Product /> : query.has('suggestions-only') ? <PnPSearchSuggestionsDemo /> : query.has('form-only') ? <section data-form-boundary="compiled"><FormPanel /></section> : <><StylesPanel /><FormattedTextPanel /><RelativeTimePanel /><AutoSkeletonPanel /><StateMessagePanel /><GroupsPanel /></>);
