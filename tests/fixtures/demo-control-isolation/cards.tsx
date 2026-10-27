import * as React from 'react';
import CardsPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/CardsPanel';
import { CardsLegacyIsolationFixture } from './cards-legacy-isolation';
/** The actual single example and independently mounted retained regression scenarios. */
export const CardsIsolationFixture: React.FC = () => <><CardsPanel /><CardsLegacyIsolationFixture /></>;
