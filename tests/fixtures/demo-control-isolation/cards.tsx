import * as React from 'react';
import { Demo as PreviewDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/e28438309bd3c4e9a1700761e8eab4eb4a08e5cd5d57d7a2504fdd526ae9e264';
import { Demo as ContentDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/60c33ada62e5de74ffb49f20cb3ab7db19420b87493206efe5343a47a1e8aab3';
import { Demo as RetentionDemo } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/generated/jsdoc-examples/664a82490443458f53895ecce9de0cbaa8fe734d84249cfd793dd43076762400';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { CardsLegacyIsolationFixture } from './cards-legacy-isolation';
/** Real emitted source scenarios and separate independent conformance fixtures. */
export const CardsIsolationFixture: React.FC = () => {
  const scenario = new URLSearchParams(window.location.search).get('cards-scenario');
  const Demo = scenario === 'content' ? ContentDemo : scenario === 'retention' ? RetentionDemo : PreviewDemo;
  const title = scenario === 'content' ? 'Configure content slots and native card interaction' : scenario === 'retention' ? 'Retain caller data when switching card types' : 'Inspect preview media and image recovery';
  return <><section data-cards-demo="panel"><Demo render={regions => <DemoCard title={title} iconName="Page" {...regions} />} /></section><CardsLegacyIsolationFixture /></>;
};
