import * as React from 'react';
import { Badge, Button, Field, Checkbox, FluentProvider, Link, Select } from '@fluentui/react-components';
import { DemoCard, useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared';
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
  const [showConditional, setShowConditional] = React.useState(false);
  const [slotClicks, setSlotClicks] = React.useState(0);
  // React 17 native ref objects use the legacy null initial value.
  const rootRef = React.useMemo(() => React.createRef<HTMLDListElement>(), []);
  const itemRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const slotRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const entries: readonly DescriptionListEntry[] = [
    { id: 'owner', label: 'Owner', value: <Link href="#description-list-demo">Ada Lovelace</Link> },
    { id: 'status', label: 'Status', value: <Badge appearance="filled">Active</Badge> },
    { id: 'literal', label: '**Literal label**', value: '**Literal value**' },
    { id: 'zero', label: 'Count', value: 0 },
    { id: 'blank', label: 'Blank', value: '' },
    { id: 'missing', label: 'Missing', value: undefined },
    { id: 'long', label: 'LongLabel'.repeat(12), value: 'LongUnbrokenValue'.repeat(20) },
    {
      id: 'slot', label: { as: 'div', ref: slotRef, children: 'Slot label', title: 'Native label slot title', 'data-description-list-demo': 'slot-container' }, value: {
        onClick: () => setSlotClicks(value => value + 1),
        children: (Component, props) => <Component {...props}><p>Block content from a slot callback.</p><Link href="#description-list-demo">Read details</Link></Component>,
      }
    },
    { id: 'conditional', label: 'Conditional slot', value: showConditional && 'Conditional content is visible.' },
    { id: 'empty-slot', label: 'Explicit empty slot', value: { children: undefined, 'data-description-list-demo': 'empty-slot-container' } },
    { id: 'state-a', label: 'State A', value: <Counter name={`${instance} A`} /> },
    { id: 'state-b', label: 'State B', value: <Counter name={`${instance} B`} /> },
  ];
  const ordered = reversed ? [...entries].reverse() : entries;
  const inspect = (): void => {
    const root = rootRef.current;
    const item = itemRef.current;
    const cell = root?.querySelector('dt');
    setObserved(`root=${root?.tagName ?? 'absent'}; pairs=${root?.querySelectorAll('dt').length ?? 0}; root hidden=${String(root?.hidden)}; item=${item?.tagName ?? 'absent'}; item hidden=${String(item?.hidden)}; slot=${slotRef.current?.tagName ?? 'absent'} inside ${slotRef.current?.parentElement?.tagName ?? 'absent'}; columns=${root ? getComputedStyle(root).gridTemplateColumns : 'absent'}; first term width=${cell?.getBoundingClientRect().width.toFixed(1) ?? 'absent'}`);
  };
  return <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}>
    <section data-description-list-demo="instance" data-description-list-instance={instance} className={sx(flex.column, gap.medium)}>
      <DemoCard diagnosticsLabel="Advanced diagnostics" title={instance === 'primary' ? 'DescriptionList' : 'Independent DescriptionList'} description="Read-only labeled values preserve native terms and definitions. Auto stacks below 480px and aligns shared columns at 480px or wider in the nearest named container. Increment either counter, then change layout, width or order to observe retained child state. Switching between data and compound APIs creates a new example." iconName="List" controls={<>

        <Field label="Layout"><Select data-description-list-demo="layout" value={layout} onChange={event => { const value = event.target.value; if (value === 'auto' || value === 'horizontal' || value === 'vertical') setLayout(value); }}><option value="auto">Auto</option><option value="horizontal">Horizontal</option><option value="vertical">Vertical</option></Select></Field>
        <Field label="Container width"><Select data-description-list-demo="width" value={pixels} onChange={event => setPixels(Number(event.target.value))}>{widths.map(value => <option key={value} value={value}>{value}px</option>)}</Select></Field>
        <Checkbox data-description-list-demo="composition" label="Compound JSX" checked={compound} onChange={(_event, data) => setCompound(data.checked === true)} />
        <Checkbox data-description-list-demo="reverse" label="Reverse stable IDs" checked={reversed} onChange={(_event, data) => setReversed(data.checked === true)} />
        <Checkbox data-description-list-demo="hidden" label="Hide root" checked={hidden} onChange={(_event, data) => setHidden(data.checked === true)} />
        <Checkbox data-description-list-demo="item-hidden" label="Hide first compound pair" checked={itemHidden} disabled={!compound} onChange={(_event, data) => setItemHidden(data.checked === true)} />
        <Checkbox data-description-list-demo="empty" label="Empty input" checked={empty} onChange={(_event, data) => setEmpty(data.checked === true)} />
        <Checkbox data-description-list-demo="rtl" label="RTL override (unchecked inherits host)" checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} />
        <Checkbox data-description-list-demo="conditional-slot" label="Show conditional value" checked={showConditional} onChange={(_event, data) => setShowConditional(data.checked === true)} />
      </>} status={<output aria-live="polite">Bubbled list clicks: {clicks} · Callback slot clicks: {slotClicks}</output>} diagnostics={<>
        <Button type="button" data-description-list-demo="inspect-ref" onClick={inspect}>Inspect native refs and layout</Button>
        <output data-description-list-demo="ref">{observed}</output>
        <output data-description-list-demo="clicks">Bubbled list clicks: {clicks}</output>
        <output data-description-list-demo="slot-clicks">Callback slot clicks: {slotClicks}</output>
      </>}>
        <div className={styles.scroll} role="region" aria-label={`${instance} description list width`} tabIndex={0}>

          <div data-description-list-demo="container" className={sx(container.inlineSize, width.px(pixels))}>
            {compound ? <DescriptionList data-description-list-demo="list" ref={rootRef} layout={layout} hidden={hidden} aria-label={`${instance} compound details`} onClick={() => setClicks(value => value + 1)}>
              {empty ? <></> : <>{ordered.map((entry, index) => <DescriptionList.Item key={entry.id} label={entry.label} ref={index === 0 ? itemRef : undefined} hidden={index === 0 && itemHidden} data-description-list-demo="pair">{entry.value}</DescriptionList.Item>)}</>}
            </DescriptionList> : <DescriptionList data-description-list-demo="list" ref={rootRef} items={empty ? [] : ordered} layout={layout} hidden={hidden} aria-label={`${instance} data details`} onClick={() => setClicks(value => value + 1)} />}
            <h5>Nested 320px container</h5>
            <div data-description-list-demo="nested-container" className={sx(container.inlineSize, width.px(320))}>
              <DescriptionList aria-label={`${instance} nested details`} items={[{ id: 'nested', label: 'Nearest container', value: 'Auto remains vertical at 320px.' }]} />
            </div>
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
