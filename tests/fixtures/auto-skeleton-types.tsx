import * as React from 'react';
import { AutoSkeleton } from '../../packages/spfx-react-toolkit/src/controls/auto-skeleton/auto-skeleton';
const selectors = ['.ignored'] as const;
<AutoSkeleton loading animation="pulse" minSize={4} ignoreSelectors={selectors} ref={React.createRef<HTMLDivElement>()} id="profile" onClick={event => event.currentTarget.focus()}>Content</AutoSkeleton>;
// @ts-expect-error loading required
<AutoSkeleton>Content</AutoSkeleton>;
// @ts-expect-error no square rendering API
<AutoSkeleton loading shape="square">Content</AutoSkeleton>;
// @ts-expect-error no configuration wrapper
<AutoSkeleton loading options={{cache:true}}>Content</AutoSkeleton>;
// @ts-expect-error no watch option
<AutoSkeleton loading watch>Content</AutoSkeleton>;
// @ts-expect-error direct HTML injection unsupported
<AutoSkeleton loading dangerouslySetInnerHTML={{__html:'Content'}}>Content</AutoSkeleton>;
// @ts-expect-error only Fluent wave/pulse
<AutoSkeleton loading animation="none">Content</AutoSkeleton>;
