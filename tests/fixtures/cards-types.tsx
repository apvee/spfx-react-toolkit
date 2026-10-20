import * as React from 'react';
import { ContentCard } from '../../packages/spfx-react-toolkit/src/controls/content-card';
import { PreviewCard } from '../../packages/spfx-react-toolkit/src/controls/preview-card';
<ContentCard title={{ children: (Tag, props) => <Tag {...props}>Title</Tag>, ref: React.createRef<HTMLSpanElement>() }} headingLevel={6} selected disabled focusMode="tab-exit" onSelectionChange={(_, data) => data.selected} checkbox={{ 'aria-label': 'Select' }} />;
<PreviewCard title="Preview" preview={false} image={{ src: '/image', alt: '' }} logo={{ as: 'img', src: '/logo', alt: 'Brand', ref: React.createRef<HTMLImageElement>() }} orientation="horizontal" headerActions={{ children: <button>Action</button> }} footer={{ ref: React.createRef<HTMLDivElement>() }} />;
// @ts-expect-error title required
<ContentCard />;
// @ts-expect-error semantic heading bound
<ContentCard title="Title" headingLevel={7} />;
// @ts-expect-error content card orientation fixed
<ContentCard title="Title" orientation="horizontal" />;
// @ts-expect-error no implicit routing
<PreviewCard title="Title" href="/route" />;
// @ts-expect-error image alt required
<PreviewCard title="Title" image={{ src: '/image' }} />;
// @ts-expect-error logo tags bounded
<PreviewCard title="Title" logo={{ as: 'a', href: '/route' }} />;
// @ts-expect-error title cannot replace heading
<ContentCard title={{ as: 'h1', children: 'Title' }} />;
// @ts-expect-error raw html unsupported
<ContentCard title="Title" dangerouslySetInnerHTML={{ __html: 'No' }} />;
// @ts-expect-error slot raw html unsupported
<ContentCard title={{ dangerouslySetInnerHTML: { __html: 'No' } }} />;

// @ts-expect-error explicit boolean titles are not meaningful
<ContentCard title={false} />;
// @ts-expect-error absent title
<ContentCard title={undefined} />;
// @ts-expect-error legacy null title unsupported
<PreviewCard title={null} />;
