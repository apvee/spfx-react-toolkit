import { mountDemoFixture } from '../demo-fixture-host';
import * as React from 'react';
import { registerIcons } from '@fluentui/react';
import { PnPSearchSuggestionsDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/demos/PnPSearchSuggestionsDemo';
import StylesPanel from '../foundation-legacy/StylesPanel';
import ProductStylesPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/StylesPanel';
import RelativeTimePanel from '../foundation-legacy/RelativeTimePanel';
import { RelativeTimeScenario as ProductRelativeTimePanel } from '../executable-examples/relative-time-scenario';
import FormattedTextPanel from '../foundation-legacy/FormattedTextPanel';
import { FormattedTextScenario as ProductFormattedTextPanel } from '../executable-examples/formatted-text-scenario';
import AutoSkeletonPanel from '../foundation-legacy/AutoSkeletonPanel';
import { AutoSkeletonScenario as ProductAutoSkeletonPanel } from '../executable-examples/auto-skeleton-scenario';
import StateMessagePanel from '../foundation-legacy/StateMessagePanel';
import { StateMessageScenario as ProductStateMessagePanel } from '../executable-examples/state-message-scenario';
import FormPanel from '../demo-control-isolation/form-regressions';
import GroupsPanel from '../demo-control-isolation/groups-regressions';
import { Demo as ImportsDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/832166d77b18a4d570db5f61f5b33c9e23a764fa19638b140c6f82c8789d8e88';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
const ProductImportsScenario: React.FC = () => <ImportsDemo render={regions => <DemoCard title="Compare root domain and historical imports" iconName="Info" contentPlacement="configuration" {...regions} />} />;
const groupFixtureIcon = <svg width="1em" height="1em" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M4 5h16v14H4zm2 2v10h12V7z" /></svg>;
registerIcons({ icons: { Color: 'color', Font: 'font', Clock: 'clock', Processing: 'loading', Info: groupFixtureIcon, Settings: groupFixtureIcon, Mail: groupFixtureIcon, Ringer: groupFixtureIcon } });
const products: Readonly<Record<string, React.ComponentType>> = {
  styles: ProductStylesPanel, 'relative-time': ProductRelativeTimePanel,
  'formatted-text': ProductFormattedTextPanel, 'auto-skeleton': ProductAutoSkeletonPanel,
  'state-message': ProductStateMessagePanel, imports: ProductImportsScenario,
};
const query = new URLSearchParams(window.location.search);
const Product = products[query.get('foundation-product') ?? ''];
mountDemoFixture(Product ? <Product /> : query.has('suggestions-only') ? <PnPSearchSuggestionsDemo /> : query.has('form-only') ? <section data-form-boundary="compiled"><FormPanel /></section> : <><StylesPanel /><FormattedTextPanel /><RelativeTimePanel /><AutoSkeletonPanel /><StateMessagePanel /><GroupsPanel /></>);
