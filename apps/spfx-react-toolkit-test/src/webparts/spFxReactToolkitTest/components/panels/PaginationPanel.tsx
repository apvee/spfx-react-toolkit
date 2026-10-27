import * as React from 'react';
import { Button, Field, Checkbox, FluentProvider, Select } from '@fluentui/react-components';
import { DemoCard, useDemoStyles } from '../shared';
import { Pagination, useSx, flex, gap, width, maxWidth, background, minWidth } from '@apvee/spfx-react-toolkit';
import type { PaginationChangeData, PaginationProps, PaginationSize, PaginationState } from '@apvee/spfx-react-toolkit';

interface Order { readonly id: string; readonly title: string; readonly category: 'Hardware' | 'Software'; }
const orders: readonly Order[] = Array.from({ length: 23 }, (_value, index) => ({
  id: `order-${index + 1}`, title: `Order ${String(index + 1).padStart(2, '0')}`,
  category: index % 2 === 0 ? 'Hardware' : 'Software',
}));
type Mode = 'known' | 'unknown' | 'cursor' | 'loadMore';
const customEnglishStrings = { pagination: 'Pagination', firstPage: 'First page', previousPage: 'Previous page', nextPage: 'Next page', lastPage: 'Last page', loadMore: 'Load additional orders', loading: 'Loading' };

export const PaginationExample: React.FC<{ instance: string; initialRtl?: boolean }> = ({ instance, initialRtl = false }) => {
  const sx = useSx();
  const styles = useDemoStyles();
  const [mode, setMode] = React.useState<Mode>('known');
  const [pageIndex, setPageIndex] = React.useState(0);
  const [pageSize, setPageSize] = React.useState(5);
  const [category, setCategory] = React.useState('all');
  const [reduced, setReduced] = React.useState(false);
  const [manual, setManual] = React.useState(false);
  const [pending, setPending] = React.useState<PaginationChangeData>();
  const [request, setRequest] = React.useState('none');
  const [hidden, setHidden] = React.useState(false);
  const [rtl, setRtl] = React.useState(initialRtl);
  const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
  const [localized, setLocalized] = React.useState(true);
  const [size, setSize] = React.useState<PaginationSize>('medium');
  const [appearance, setAppearance] = React.useState<NonNullable<PaginationProps['appearance']>>('subtle');
  const [narrow, setNarrow] = React.useState(false);
  const [observedRef, setObservedRef] = React.useState('not inspected');
  // React 17 native ref objects use the legacy null initial value.
  const navRef = React.useMemo(() => React.createRef<HTMLElement>(), []);
  const filtered = orders.filter(order => category === 'all' || order.category === category);
  const available = reduced ? filtered.slice(0, 3) : filtered;
  const start = mode === 'loadMore' ? 0 : pageIndex * pageSize;
  const end = Math.min(available.length, (pageIndex + 1) * pageSize);
  const rows = available.slice(start, end);
  const hasNextPage = end < available.length;
  const pagination: PaginationState = mode === 'known'
    ? { type: 'pages', pageIndex, pageSize, totalItems: available.length }
    : mode === 'unknown' ? { type: 'pages', pageIndex, pageSize, hasNextPage }
      : mode === 'cursor' ? { type: 'cursor', hasPreviousPage: pageIndex > 0, hasNextPage }
        : { type: 'loadMore', hasNextPage };
  const reset = (): void => { setPageIndex(0); setPending(undefined); };
  const apply = (data: PaginationChangeData): void => {
    if (data.type === 'pages') setPageIndex(data.pageIndex);
    else if (data.type === 'cursor') setPageIndex(value => value + (data.direction === 'next' ? 1 : -1));
    else setPageIndex(value => value + 1);
  };
  return <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}>
    <section data-pagination-demo="instance" data-pagination-instance={instance} className={sx(flex.column, gap.medium)}>
      <DemoCard title="Pagination" description="Browse team orders. Requests update the same result list." iconName="List" controls={<>

        <Field label="Mode"><Select data-demo-variant="Mode" data-pagination-demo="mode" value={mode} onChange={event => { const value = event.target.value; if (value === 'known' || value === 'unknown' || value === 'cursor' || value === 'loadMore') setMode(value); reset(); }}>
          <option value="known">Pages · known total</option><option value="unknown">Pages · unknown total</option><option value="cursor">Cursor</option><option value="loadMore">Load more</option>
        </Select></Field>
        <Field label="Page size"><Select data-pagination-demo="page-size" value={pageSize} onChange={event => { setPageSize(Number(event.target.value)); reset(); }}>
          {[3, 5, 10].map(value => <option key={value} value={value}>{value}</option>)}
        </Select></Field>
        <Field label="Category"><Select data-pagination-demo="filter" value={category} onChange={event => { setCategory(event.target.value); reset(); }}>
          <option value="all">All orders</option><option value="Hardware">Hardware</option><option value="Software">Software</option><option value="missing">No matching orders</option>
        </Select></Field>
        <Field label="Control size"><Select data-pagination-demo="size" value={size} onChange={event => { const value = event.target.value; if (value === 'small' || value === 'medium' || value === 'large') setSize(value); }}>
          <option value="small">Small</option><option value="medium">Medium</option><option value="large">Large</option>
        </Select></Field>
        <Field label="Appearance"><Select data-pagination-demo="appearance" value={appearance} onChange={event => {
          const value = event.target.value;
          if (value === 'primary' || value === 'secondary' || value === 'outline' || value === 'subtle' || value === 'transparent') setAppearance(value);
        }}>
          <option value="primary">Primary</option><option value="secondary">Secondary</option><option value="outline">Outline</option><option value="subtle">Subtle</option><option value="transparent">Transparent</option>
        </Select></Field>
      </>} commands={<>
        <div className={styles.actions} role="group" aria-label="Consumer simulation controls">
          <Checkbox data-pagination-demo="reduce-total" label="Reduce total to three" checked={reduced} onChange={(_event, data) => { setReduced(data.checked === true); reset(); }} />
          <Checkbox data-pagination-demo="manual-loading" label="Complete requests manually" checked={manual} disabled={pending !== undefined} onChange={(_event, data) => setManual(data.checked === true)} />
        </div>
            <Button type="button" data-pagination-demo="complete" disabled={pending === undefined} onClick={() => { if (pending !== undefined) apply(pending); setPending(undefined); }}>Complete loading</Button>
      </>} status={<output aria-live="polite">Last request: {request}</output>} diagnostics={<>
        <Checkbox data-pagination-demo="hidden" label="Hide navigation" checked={hidden} onChange={(_event, data) => setHidden(data.checked === true)} />
        <Checkbox data-pagination-demo="rtl" label="RTL override (unchecked inherits host)" checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} />
        <Checkbox data-pagination-demo="localization" label="Custom labels (unchecked: SPFx defaults)" checked={localized} onChange={(_event, data) => setLocalized(data.checked === true)} />
        <Checkbox data-pagination-demo="narrow" label="320px navigation" checked={narrow} onChange={(_event, data) => setNarrow(data.checked === true)} />

            <Button type="button" data-pagination-demo="inspect-ref" onClick={() => setObservedRef(`${navRef.current?.tagName ?? 'absent'}; hidden=${String(navRef.current?.hidden)}; busy=${navRef.current?.getAttribute('aria-busy') ?? 'absent'}`)}>Inspect navigation ref</Button>

        <output data-pagination-demo="request">Last request: {request}</output>
        <output data-pagination-demo="ref">Native ref: {observedRef}</output>
      </>}>

        <div className={styles.scroll} role="region" aria-label={`${instance} order results`} tabIndex={0}>
          <ul data-pagination-demo="results" aria-label={`${instance} orders`}>
            {rows.map(row => <li key={row.id}>{row.title} · {row.category}</li>)}
          </ul></div>
        <p data-pagination-demo="summary" role="status" aria-live="polite" aria-atomic="true">
          {available.length === 0 ? 'No matching orders.' : `${mode === 'loadMore' ? 'Loaded' : `Page ${pageIndex + 1}:`} items ${start + 1}–${end} of ${available.length}.`}
          {pending !== undefined ? ' Loading requested results; current rows remain visible.' : ''}
        </p>
        <div className={sx(flex.column, gap.small)}>
          <div className={narrow ? sx(width.full, maxWidth.px(320)) : sx(width.full)}>
            <Pagination data-pagination-demo="navigation" ref={navRef} hidden={hidden} size={size} appearance={appearance}
              pagination={pagination} loading={pending !== undefined} strings={localized ? customEnglishStrings : undefined}
              aria-label={`${instance} order navigation`} onPaginationChange={(_event, data) => {
                setRequest(JSON.stringify(data));
                if (manual) setPending(data); else apply(data);
              }} />
          </div>
        </div>
      </DemoCard>
    </section>
  </FluentProvider>;
};

const PaginationPanel: React.FC = () => {
  const styles = useDemoStyles();
  return <section data-pagination-demo="panel" className={styles.root}>

    <PaginationExample instance="primary" />

  </section>;
};
export default PaginationPanel;
