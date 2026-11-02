/**
 * Read-only term/value metadata with native description-list semantics.
 *
 * @packageDocumentation
 * @docPage controls/description-list
 */
import * as React from 'react';
import { renderContentSlot } from '../content-slot.internal';
import { useSx } from '../../hooks/useSx';
import * as margin from '../../helpers/styles/margin';
import * as minWidth from '../../helpers/styles/min-width';
import * as gap from '../../helpers/styles/gap';
import * as typography from '../../helpers/styles/typography';
import * as foreground from '../../helpers/styles/foreground';
import { useDescriptionListStyles } from './description-list.styles.internal';
import type { DescriptionListLayout, DescriptionListProps, DescriptionListItemProps } from './description-list.types';

const DescriptionListContext = React.createContext<DescriptionListLayout | undefined>(undefined);

/**
 * One pair with preserved native semantic cells and caller-owned content/state.
 * @param props - Native grouping attributes and Fluent content slots.
 * @returns A grouping div containing one dt and one dd.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { DescriptionList } from '@apvee/spfx-react-toolkit';
 * <DescriptionList><DescriptionList.Item label="Owner">Ada</DescriptionList.Item></DescriptionList>
 * ```
 */
const DescriptionListItem = React.forwardRef<HTMLDivElement, DescriptionListItemProps>(({
  label, children, className, ...nativeProps
}, ref) => {
  const layout = React.useContext(DescriptionListContext);
  const sx = useSx(), styles = useDescriptionListStyles();
  if (layout === undefined) throw new Error('DescriptionList.Item must be rendered inside DescriptionList.');
  delete nativeProps.dangerouslySetInnerHTML;
  const itemLayout = layout === 'horizontal' ? styles.itemHorizontal : layout === 'vertical' ? styles.itemVertical : styles.itemAuto;
  return <div {...nativeProps} ref={ref} data-description-list-item="" className={sx(styles.item, itemLayout, minWidth.zero, className)}>
    <dt className={sx(styles.cell, margin.none, minWidth.zero)}>
      {renderContentSlot(label, 'div', sx, { shorthand: 'wrap', className: sx(typography.body1Strong, foreground.primary, minWidth.zero) })}
    </dt>
    <dd className={sx(styles.cell, margin.none, minWidth.zero)}>
      {renderContentSlot(children, 'div', sx, { shorthand: 'wrap', className: sx(typography.body1, foreground.primary, minWidth.zero) })}
    </dd>
  </div>;
});
DescriptionListItem.displayName = 'DescriptionList.Item';

const DescriptionListRoot = React.forwardRef<HTMLDListElement, DescriptionListProps>(({
  items, children, layout = 'auto', className, ...nativeProps
}, ref) => {
  const sx = useSx(), styles = useDescriptionListStyles();
  delete nativeProps.dangerouslySetInnerHTML;
  // React.ReactNode retains legacy null children; treat them like undefined without introducing null defaults.
  if (items !== undefined ? items.length === 0 : children === undefined || children === null || typeof children === 'boolean') return <></>;
  return <DescriptionListContext.Provider value={layout}>
    <dl {...nativeProps} ref={ref} data-description-list-layout={layout}
      className={sx(styles.root, margin.none, minWidth.zero, gap.medium, styles[layout], className)}>
      {items === undefined ? children : items.map(item => <DescriptionListItem key={item.id} label={item.label}>{item.value}</DescriptionListItem>)}
    </dl>
  </DescriptionListContext.Provider>;
});
DescriptionListRoot.displayName = 'DescriptionList';

/**
 * Display read-only metadata for one record as a native description list.
 * Supply immutable items with stable IDs or compound Item children. Strings remain literal;
 * formatting, missing-value labels, loading and empty states belong to the consumer.
 * Auto layout stacks without a named container and shares two columns from 480px of the
 * nearest apvee-sx ancestor. Explicit horizontal/vertical overrides the query.
 * Uses Fluent theme variables and the host Griffel renderer, with no SPFx provider requirement.
 * No generated inline styles, keyboard behavior, cloning or application state.
 * Empty data/absent children render nothing; empty fragments can leave an empty dl.
 * @param props - Native dl attributes, layout and one content pathway.
 * @returns A native dl of term/definition pairs, or an empty fragment.
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { DescriptionList } from '@apvee/spfx-react-toolkit';
 * <DescriptionList items={[{ id: 'owner', label: 'Owner', value: 'Ada Lovelace' }]} />
 * ```
 *
 * @useWhen One record has read-only named values requiring native term/definition semantics.
 * @avoidWhen Use DataGrid for tabular collections of records.
 *
 *
 * @example Use the nearest named size container
 * ```tsx
 * import * as React from 'react';
 * import { DescriptionList, container, useSx, width } from '@apvee/spfx-react-toolkit';
 * function RecordDetails() {
 *   const sx = useSx();
 *   return <section className={sx(container.inlineSize, width.full)}>
 *     <DescriptionList items={[{ id: 'owner', label: 'Owner', value: 'Ada' }]} />
 *   </section>;
 * }
 * ```
 *
 *
 * @example Compose conditional metadata pairs with inner slot refs
 * ```tsx
 * import * as React from 'react';
 * import { Badge } from '@fluentui/react-components';
 * import { DescriptionList } from '@apvee/spfx-react-toolkit';
 * function ConditionalDetails({ showStatus }: { showStatus: boolean }) {
 *   const valueRef = React.useRef<HTMLDivElement>(null); // React 17's legacy ref initializer.
 *   return <DescriptionList aria-label="Record details" layout="horizontal">
 *     <DescriptionList.Item label={{ title: 'Account owner', children: 'Owner' }}>
 *       {{ ref: valueRef, children: (Component, props) => <Component {...props}><p>Ada Lovelace</p></Component> }}
 *     </DescriptionList.Item>
 *     {showStatus && <DescriptionList.Item label="Status"><Badge appearance="filled">Active</Badge></DescriptionList.Item>}
 *   </DescriptionList>;
 * }
 * ```
 *
 * @example Inspect responsive metadata with data and compound APIs
 * ```tsx
 * import * as React from 'react';
 * import { Badge, Button, Field, Checkbox, FluentProvider, Link, Select, makeStyles } from '@fluentui/react-components';
 *
 * import { DescriptionList, container, flex, gap, useSx, width, background, minWidth } from '@apvee/spfx-react-toolkit';
 * import type { DescriptionListEntry, DescriptionListLayout } from '@apvee/spfx-react-toolkit';
 *
 * const Counter: React.FC<{ name: string; resetRevision: number }> = ({ name, resetRevision }) => {
 *   const [count, setCount] = React.useState(0);
 *   const previous = React.useRef(resetRevision);
 *   React.useEffect(() => { if (previous.current !== resetRevision) { previous.current = resetRevision; setCount(0); } }, [resetRevision]);
 *   return <Button type="button" data-description-list-demo="counter" onClick={() => setCount(value => value + 1)}>{name}: {count}</Button>;
 * };
 * const widths = [320, 479, 480, 639, 640, 900] as const;
 *
 * type DemoRegions = {
 *   controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode;
 *   children?: React.ReactNode; diagnostics?: React.ReactNode;
 * };
 * const standalone = (regions: DemoRegions): React.ReactElement => <>
 *   {regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}
 * </>;
 * const useExampleStyles = makeStyles({ scroll: { overflowX: 'auto', minWidth: 0, maxWidth: '100%' } });
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const instance = 'primary';
 *   const sx = useSx();
 *   const styles = useExampleStyles();
 *   const [layout, setLayout] = React.useState<DescriptionListLayout>('auto');
 *   const [pixels, setPixels] = React.useState<number>(640);
 *   const [compound, setCompound] = React.useState(false);
 *   const [reversed, setReversed] = React.useState(false);
 *   const [hidden, setHidden] = React.useState(false);
 *   const [itemHidden, setItemHidden] = React.useState(false);
 *   const [empty, setEmpty] = React.useState(false);
 *   const [rtl, setRtl] = React.useState(false);
 *   const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
 *   const [observed, setObserved] = React.useState('not inspected');
 *   const [clicks, setClicks] = React.useState(0);
 *   const [itemClicks, setItemClicks] = React.useState(0);
 *   // React 17 native ref objects use the legacy null initial value.
 *   const rootRef = React.useMemo(() => React.createRef<HTMLDListElement>(), []);
 *   const itemRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
 *   const entries: readonly DescriptionListEntry[] = [
 *     { id: 'owner', label: 'Owner', value: <Link href="#description-list-demo">Ada Rossi</Link> },
 *     { id: 'status', label: 'Status', value: <Badge appearance="tint" color="success">In progress</Badge> },
 *     { id: 'team', label: 'Team', value: 'Operations' },
 *     { id: 'updated', label: 'Updated', value: '13 October 2026' },
 *     { id: 'reviews', label: 'Revisions', value: <Counter name="Revisions" resetRevision={resetRevision} /> },
 *   ];
 *   const ordered = reversed ? [...entries].reverse() : entries;
 *   const inspect = (): void => {
 *     const root = rootRef.current;
 *     const item = itemRef.current;
 *     const cell = root?.querySelector('dt');
 *     setObserved(`root=${root?.tagName ?? 'absent'}; pairs=${root?.querySelectorAll('dt').length ?? 0}; root hidden=${String(root?.hidden)}; item=${item?.tagName ?? 'absent'}; item hidden=${String(item?.hidden)}; columns=${root ? root.ownerDocument.defaultView?.getComputedStyle(root).gridTemplateColumns : 'absent'}; first term width=${cell?.getBoundingClientRect().width.toFixed(1) ?? 'absent'}`);
 *   };
 *   const previousReset = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision;
 *     setLayout('auto'); setPixels(640); setCompound(false); setReversed(false);
 *     setHidden(false); setItemHidden(false); setEmpty(false); setRtl(false);
 *     setObserved('not inspected'); setClicks(0); setItemClicks(0);
 *   }, [resetRevision]);
 *   return <FluentProvider data-description-list-demo="provider" dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}>
 *     <section data-description-list-demo="instance" data-description-list-instance={instance} className={sx(flex.column, gap.medium)}>
 *       {render({ controls: <>
 *
 *         <Field label="Layout"><Select data-description-list-demo="layout" value={layout} onChange={event => { const value = event.target.value; if (value === 'auto' || value === 'horizontal' || value === 'vertical') setLayout(value); }}><option value="auto">Auto</option><option value="horizontal">Horizontal</option><option value="vertical">Vertical</option></Select></Field>
 *         <Field label="Container width"><Select data-description-list-demo="width" value={pixels} onChange={event => setPixels(Number(event.target.value))}>{widths.map(value => <option key={value} value={value}>{value}px</option>)}</Select></Field>
 *         <Field label="API"><Select data-demo-variant="API" data-description-list-demo="composition" value={compound ? 'compound' : 'data'} onChange={(_, data) => setCompound(data.value === 'compound')}><option value="data">Data</option><option value="compound">Compound JSX</option></Select></Field>
 *         <Checkbox data-description-list-demo="reverse" label="Reverse stable IDs" checked={reversed} onChange={(_event, data) => setReversed(data.checked === true)} />
 *       </>, status: <output aria-live="polite">Bubbled list clicks: {clicks} · Item clicks: {itemClicks}</output>, diagnostics: <>
 *         <Checkbox data-description-list-demo="hidden" label="Hide root" checked={hidden} onChange={(_event, data) => setHidden(data.checked === true)} />
 *         <Checkbox data-description-list-demo="item-hidden" label="Hide first compound pair" checked={itemHidden} disabled={!compound} onChange={(_event, data) => setItemHidden(data.checked === true)} />
 *         <Checkbox data-description-list-demo="empty" label="Empty input" checked={empty} onChange={(_event, data) => setEmpty(data.checked === true)} />
 *         <Checkbox data-description-list-demo="rtl" label="RTL override (unchecked inherits host)" checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} />
 *
 *         <Button type="button" data-description-list-demo="inspect-ref" onClick={inspect}>Inspect native refs and layout</Button>
 *         <output data-description-list-demo="ref">{observed}</output>
 *         <output data-description-list-demo="clicks">Bubbled list clicks: {clicks} · Item clicks: {itemClicks}</output>
 *         <p>Stable IDs preserve counters through layout, width and order. Changing API or removing data can remount children; reset restores the local baseline.</p>
 *       </>, children:
 *         <div className={styles.scroll} role="region" aria-label={`${instance} description list width`} tabIndex={0}>
 *
 *           <div data-description-list-demo="container" className={sx(container.inlineSize, width.px(pixels))}>
 *             {compound ? <DescriptionList data-description-list-demo="list" ref={rootRef} layout={layout} hidden={hidden} aria-label={`${instance} compound details`} onClick={() => setClicks(value => value + 1)}>
 *               {empty ? <></> : <>{ordered.map((entry, index) => <DescriptionList.Item key={entry.id} label={entry.label} ref={index === 0 ? itemRef : undefined} hidden={index === 0 && itemHidden} data-description-list-demo="pair" onClick={() => setItemClicks(value => value + 1)}>{entry.value}</DescriptionList.Item>)}</>}
 *             </DescriptionList> : <DescriptionList data-description-list-demo="list" ref={rootRef} items={empty ? [] : ordered} layout={layout} hidden={hidden} aria-label={`${instance} data details`} onClick={() => setClicks(value => value + 1)} />}
 *
 *           </div>
 *         </div>
 *       })}
 *     </section>
 *   </FluentProvider>;
 * }
 *
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Resize the named container across 480px, compare data and compound layouts, reorder stable IDs without losing the counter, and inspect native refs, hidden state and events.',
 * };
 * ```
 */
export const DescriptionList = Object.assign(DescriptionListRoot, {
  /**
   * Render one term/value pair inside DescriptionList, preserving native dt/dd cells.
   * Native attributes and ref target the grouping div; Fluent slots customize cell content.
   * @param props - Native grouping attributes, required label and associated value slot.
   * @returns A grouping div containing one dt and one dd; requires a DescriptionList parent.
   * @example
   * ```tsx
   * import * as React from 'react';
   * import { DescriptionList } from '@apvee/spfx-react-toolkit';
   * <DescriptionList><DescriptionList.Item label="Owner">Ada</DescriptionList.Item></DescriptionList>
   * ```
   */
  Item: DescriptionListItem,
});
