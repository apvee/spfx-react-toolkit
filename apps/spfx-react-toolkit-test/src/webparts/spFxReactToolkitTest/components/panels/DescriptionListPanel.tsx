import * as React from 'react';
import { Badge, Button, Field, Checkbox, FluentProvider, Link, Select } from '@fluentui/react-components';
import { DemoCard, useDemoStyles } from '../shared';
import { DescriptionList, container, flex, gap, useSx, width, background, minWidth } from '@apvee/spfx-react-toolkit';
import type { DescriptionListEntry, DescriptionListLayout } from '@apvee/spfx-react-toolkit';

const Counter: React.FC<{ name: string }> = ({ name }) => {
  const [count, setCount] = React.useState(0);
  return <Button type="button" data-description-list-demo="counter" onClick={() => setCount(value => value + 1)}>{name}: {count}</Button>;
};
const widths = [320, 479, 480, 639, 640, 900] as const;

export const DescriptionListExample: React.FC<{ instance: string; initialRtl?: boolean }> = ({ instance, initialRtl = false }) => {
  const sx = useSx();
  const styles = useDemoStyles();
  const [layout, setLayout] = React.useState<DescriptionListLayout>('auto');
  const [pixels, setPixels] = React.useState<number>(640);
  const [compound, setCompound] = React.useState(false);
  const [reversed, setReversed] = React.useState(false);
  const [hidden, setHidden] = React.useState(false);
  const [itemHidden, setItemHidden] = React.useState(false);
  const [empty, setEmpty] = React.useState(false);
  const [rtl, setRtl] = React.useState(initialRtl);
  const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
  const [observed, setObserved] = React.useState('not inspected');
  const [clicks, setClicks] = React.useState(0);
  // React 17 native ref objects use the legacy null initial value.
  const rootRef = React.useMemo(() => React.createRef<HTMLDListElement>(), []);
  const itemRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const entries: readonly DescriptionListEntry[] = [
    { id: 'owner', label: 'Owner', value: <Link href="#description-list-demo">Ada Rossi</Link> },
    { id: 'status', label: 'Status', value: <Badge appearance="tint" color="success">In progress</Badge> },
    { id: 'team', label: 'Team', value: 'Operations' },
    { id: 'updated', label: 'Updated', value: '13 October 2026' },
    { id: 'reviews', label: 'Revisions', value: <Counter name="Revisions" /> },
  ];
  const ordered = reversed ? [...entries].reverse() : entries;
  const inspect = (): void => {
    const root = rootRef.current;
    const item = itemRef.current;
    const cell = root?.querySelector('dt');
    setObserved(`root=${root?.tagName ?? 'absent'}; pairs=${root?.querySelectorAll('dt').length ?? 0}; root hidden=${String(root?.hidden)}; item=${item?.tagName ?? 'absent'}; item hidden=${String(item?.hidden)}; columns=${root ? getComputedStyle(root).gridTemplateColumns : 'absent'}; first term width=${cell?.getBoundingClientRect().width.toFixed(1) ?? 'absent'}`);
  };
  return <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}>
    <section data-description-list-demo="instance" data-description-list-instance={instance} className={sx(flex.column, gap.medium)}>
      <DemoCard title="DescriptionList" description="Project details with readable labels and values. Change layout or order while preserving the revision counter." iconName="List" controls={<>

        <Field label="Layout"><Select data-description-list-demo="layout" value={layout} onChange={event => { const value = event.target.value; if (value === 'auto' || value === 'horizontal' || value === 'vertical') setLayout(value); }}><option value="auto">Auto</option><option value="horizontal">Horizontal</option><option value="vertical">Vertical</option></Select></Field>
        <Field label="Container width"><Select data-description-list-demo="width" value={pixels} onChange={event => setPixels(Number(event.target.value))}>{widths.map(value => <option key={value} value={value}>{value}px</option>)}</Select></Field>
        <Field label="API"><Select data-demo-variant="API" data-description-list-demo="composition" value={compound ? 'compound' : 'data'} onChange={(_, data) => setCompound(data.value === 'compound')}><option value="data">Data</option><option value="compound">Compound JSX</option></Select></Field>
        <Checkbox data-description-list-demo="reverse" label="Reverse stable IDs" checked={reversed} onChange={(_event, data) => setReversed(data.checked === true)} />
      </>} status={<output aria-live="polite">Bubbled list clicks: {clicks}</output>} diagnostics={<>
        <Checkbox data-description-list-demo="hidden" label="Hide root" checked={hidden} onChange={(_event, data) => setHidden(data.checked === true)} />
        <Checkbox data-description-list-demo="item-hidden" label="Hide first compound pair" checked={itemHidden} disabled={!compound} onChange={(_event, data) => setItemHidden(data.checked === true)} />
        <Checkbox data-description-list-demo="empty" label="Empty input" checked={empty} onChange={(_event, data) => setEmpty(data.checked === true)} />
        <Checkbox data-description-list-demo="rtl" label="RTL override (unchecked inherits host)" checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} />

        <Button type="button" data-description-list-demo="inspect-ref" onClick={inspect}>Inspect native refs and layout</Button>
        <output data-description-list-demo="ref">{observed}</output>
        <output data-description-list-demo="clicks">Bubbled list clicks: {clicks}</output>
      </>}>
        <div className={styles.scroll} role="region" aria-label={`${instance} description list width`} tabIndex={0}>

          <div data-description-list-demo="container" className={sx(container.inlineSize, width.px(pixels))}>
            {compound ? <DescriptionList data-description-list-demo="list" ref={rootRef} layout={layout} hidden={hidden} aria-label={`${instance} compound details`} onClick={() => setClicks(value => value + 1)}>
              {empty ? <></> : <>{ordered.map((entry, index) => <DescriptionList.Item key={entry.id} label={entry.label} ref={index === 0 ? itemRef : undefined} hidden={index === 0 && itemHidden} data-description-list-demo="pair">{entry.value}</DescriptionList.Item>)}</>}
            </DescriptionList> : <DescriptionList data-description-list-demo="list" ref={rootRef} items={empty ? [] : ordered} layout={layout} hidden={hidden} aria-label={`${instance} data details`} onClick={() => setClicks(value => value + 1)} />}

          </div>
        </div>
      </DemoCard>
    </section>
  </FluentProvider>;
};

const DescriptionListPanel: React.FC = () => {
  const styles = useDemoStyles();
  return <section id="description-list-demo" data-description-list-demo="panel" className={styles.root}>

    <DescriptionListExample instance="primary" />

  </section>;
};
export default DescriptionListPanel;
