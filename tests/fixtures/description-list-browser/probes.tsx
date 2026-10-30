import * as React from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { DescriptionList, container, useSx, width } from '@apvee/spfx-react-toolkit';
import type { DescriptionListLayout } from '@apvee/spfx-react-toolkit';

const StatefulValue: React.FC = () => {
  const [value, setValue] = React.useState('initial');
  return <input aria-label="Retained value" value={value} onChange={event => setValue(event.target.value)} />;
};

/** Browser-only native input and ancestor fallback checks complement the actual demo. */
export const DescriptionListProbes: React.FC = () => {
  const sx = useSx();
  const [layout, setLayout] = React.useState<DescriptionListLayout>('auto');
  const [pixels, setPixels] = React.useState(640);
  const [reverse, setReverse] = React.useState(false);
  const rows = [
    { id: 'input', label: 'Editable caller content', value: <StatefulValue /> },
    { id: 'other', label: 'Other', value: 'Stable keyed sibling' },
  ];
  return <FluentProvider theme={webLightTheme}>
    <section data-description-list-probe="host">
      <button type="button" data-description-list-probe="resize" onClick={() => setPixels(value => value === 640 ? 320 : 640)}>Resize probe</button>
      <button type="button" data-description-list-probe="layout" onClick={() => setLayout(value => value === 'horizontal' ? 'vertical' : 'horizontal')}>Change probe layout</button>
      <button type="button" data-description-list-probe="reverse" onClick={() => setReverse(value => !value)}>Reorder probe</button>
      <div className={sx(container.inlineSize, width.px(pixels))}>
        <DescriptionList data-description-list-probe="state" layout={layout} items={reverse ? [...rows].reverse() : rows} />
      </div>
      <DescriptionList data-description-list-probe="fallback" items={[{ id: 'fallback', label: 'No named ancestor', value: 'Vertical fallback at every viewport' }]} />
    </section>
  </FluentProvider>;
};
