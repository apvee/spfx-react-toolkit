import * as React from 'react';
import { Badge, Button, Checkbox, Field, Input, Select, FluentProvider, makeStyles, mergeClasses } from '@fluentui/react-components';
import { Apps24Regular, Document16Regular, ArrowRight20Regular, ArrowSync20Regular, CheckmarkCircle16Regular, ImageOff20Regular, Info20Regular, Share16Regular } from '@fluentui/react-icons';
import { DemoCard, useDemoStyles } from '../shared';

import { AutoSkeleton, ContentCard, PreviewCard, useSx, text, background, minWidth } from '@apvee/spfx-react-toolkit';
import type { PreviewCardProps, ContentCardProps } from '@apvee/spfx-react-toolkit';

const cover = (accent: string): string => 'data:image/svg+xml,' + encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="640" height="300" viewBox="0 0 640 300"><rect width="640" height="300" fill="${accent}"/><path d="M0 230L190 70l150 135L500 35l140 135v130H0Z" fill="#ffffff" opacity=".16"/><circle cx="514" cy="86" r="34" fill="#ffffff" opacity=".65"/><path d="M64 104h200M64 126h132" stroke="#ffffff" stroke-width="8" opacity=".7"/></svg>`);
const firstCover = cover('#245b83'), secondCover = cover('#635585');
const useCardStyles = makeStyles({ subject: { width: '100%', maxWidth: '520px' }, compact: { width: '100%', maxWidth: '280px' } });
const modes = ['image', 'custom', 'empty-object', 'empty-callback', 'false', 'empty-fragment'] as const;
type PreviewMode = typeof modes[number];

/** Private per-instance scenario, reused by the independent regression fixture. */
export const CardExamples: React.FC<{ name: string; configuration?: React.ReactNode }> = ({ name, configuration }) => {
  const sx = useSx(); const styles = useDemoStyles(); const cardStyles = useCardStyles();
  const [horizontal, setHorizontal] = React.useState(true), [narrow, setNarrow] = React.useState(false);
  const [loading, setLoading] = React.useState(false), [disabled, setDisabled] = React.useState(false);
  const [decorative, setDecorative] = React.useState(false), [customFallback, setCustomFallback] = React.useState(false);
  const [logoCallback, setLogoCallback] = React.useState(false);
  const [mode, setMode] = React.useState<PreviewMode>('image');
  const [title, setTitle] = React.useState('Quarterly report');
  const [description, setDescription] = React.useState('Team results, goals and next steps.');
  const [kind, setKind] = React.useState<'preview' | 'content'>('preview');
  const [notes, setNotes] = React.useState('To share at the team meeting.');
  const [src, setSrc] = React.useState(firstCover);
  const [headerClicks, setHeaderClicks] = React.useState(0), [actionClicks, setActionClicks] = React.useState(0);
  const [rootClicks, setRootClicks] = React.useState(0), [selectionRequests, setSelectionRequests] = React.useState(0);
  const [selectable, setSelectable] = React.useState(false);
  const [selected, setSelected] = React.useState(false);
  const [observations, setObservations] = React.useState('Not inspected.');
  const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []), titleRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
  const previewRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const logoDivRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const buttonRef = React.useMemo(() => React.createRef<HTMLButtonElement>(), []);
  const previewCalls = React.useRef(0), fallbackCalls = React.useRef(0), logoCalls = React.useRef(0);
  const preview: PreviewCardProps['preview'] = mode === 'image' ? undefined : mode === 'custom' ? {
    ref: previewRef, 'data-cards-demo': 'preview-slot', children: (Tag, props) => { previewCalls.current++; return <Tag {...props}><img src={secondCover} alt="Caller-owned custom cover" /></Tag>; },
  } : mode === 'empty-object' ? { ref: previewRef, 'data-cards-demo': 'preview-slot' }
    : mode === 'empty-callback' ? { children: () => { previewCalls.current++; return undefined; } }
      : mode === 'false' ? false : <></>;
  // The shared model keeps props and notes independent of the selected presentation.
  const commonProps: ContentCardProps & React.RefAttributes<HTMLDivElement> = {
    ref: rootRef, title: { ref: titleRef, 'data-cards-demo': 'title-slot', className: 'cards-consumer-title', children: title || 'Untitled report' },
    headingLevel: 2, description, disabled,
    icon: <Document16Regular aria-hidden="true" />,
    headerActions: { 'data-cards-demo': 'header-slot', children: <Button data-cards-demo="header-action" ref={buttonRef} icon={<Share16Regular aria-hidden="true" />} className={sx(text.noWrap)} appearance="subtle" size="small" onClick={() => setHeaderClicks(value => value + 1)}>Share</Button> },
    actions: { 'data-cards-demo': 'actions-slot', children: <Button data-cards-demo="footer-action" icon={<ArrowRight20Regular aria-hidden="true" />} appearance="primary" size="small" onClick={() => setActionClicks(value => value + 1)}>Open report</Button> },
    footer: { 'data-cards-demo': 'footer-slot', children: 'Updated today · Team Operations' },
    selected: selectable ? selected : undefined, onSelectionChange: selectable ? (_, data) => { setSelected(data.selected); setSelectionRequests(value => value + 1); } : undefined,
    checkbox: { 'aria-label': `Select ${name} content` }, focusMode: 'tab-exit',
    onClick: () => setRootClicks(value => value + 1),
    children: <Field label="Retained notes"><Input input={{ as: 'input', ...{ 'data-cards-demo': 'retained-input' } }} value={notes} onChange={(_, data) => setNotes(data.value)} /></Field>,
  };
  return <section data-cards-demo="instance" data-card-instance={name} className={styles.panel}>
    <DemoCard title="Cards" iconName="Page" description="Configure the card and inspect the result directly on the web part surface."
      controls={<>
        <Field label="Title"><Input input={{ as: 'input', ...{ 'data-cards-demo': 'title-input' } }} value={title} onChange={(_, data) => setTitle(data.value)} /></Field>
        <Field label="Description"><Input input={{ as: 'input', ...{ 'data-cards-demo': 'description-input' } }} value={description} onChange={(_, data) => setDescription(data.value)} /></Field>
        <Field label="Type"><Select data-demo-variant="card-type" data-cards-demo="kind" value={kind} onChange={event => setKind(event.target.value === 'content' ? 'content' : 'preview')}><option value="preview">PreviewCard</option><option value="content">ContentCard</option></Select></Field>
        <Field label="Orientation"><Select data-cards-demo="orientation" value={horizontal ? 'horizontal' : 'vertical'} onChange={event => setHorizontal(event.target.value === 'horizontal')} disabled={kind === 'content'}><option value="horizontal">Horizontal</option><option value="vertical">Vertical</option></Select></Field>
        <Field label="Width"><Select data-cards-demo="width" value={narrow ? 'compact' : 'standard'} onChange={event => setNarrow(event.target.value === 'compact')}><option value="standard">Standard · up to 520 px</option><option value="compact">Compact · up to 280 px</option></Select></Field>
        <div className={mergeClasses(styles.controlRow, styles.actions)}><Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'loading' } }} label="Loading (AutoSkeleton)" checked={loading} onChange={(_, data) => setLoading(data.checked === true)} /><Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'disabled' } }} label="Disabled" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} /></div>
      </>}
      commands={<><Button size="small" data-cards-demo="fail-image" disabled={kind === 'content'} icon={<ImageOff20Regular aria-hidden="true" />} onClick={() => setSrc('/cards-demo-missing-image.png')}>Simulate image error</Button><Button size="small" data-cards-demo="new-src" disabled={kind === 'content'} icon={<ArrowSync20Regular aria-hidden="true" />} onClick={() => setSrc(current => current === firstCover ? secondCover : firstCover)}>Reload image</Button></>}
      status={<div data-cards-demo="counters" className={styles.actions} aria-live="polite"><Badge appearance="tint" color={loading ? 'informative' : disabled ? 'warning' : 'success'} icon={<CheckmarkCircle16Regular aria-hidden="true" />}>{loading ? 'Loading' : disabled ? 'Disabled' : 'Ready'}</Badge><span>Image: {kind === 'content' ? 'not applicable' : src === '/cards-demo-missing-image.png' ? 'fallback' : 'available'}</span><span>Header:{headerClicks} Footer:{actionClicks}</span></div>}
      diagnostics={<>
        {configuration}
        <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'selectable' } }} label="Selectable card" checked={selectable} onChange={(_, data) => setSelectable(data.checked === true)} />
        <Field label="preview"><Select data-cards-demo="preview-mode" value={mode} disabled={kind === 'content'} onChange={event => { const value = event.target.value; if (modes.some(candidate => candidate === value)) setMode(value as PreviewMode); }}>{modes.map(value => <option key={value} value={value}>{value}</option>)}</Select></Field>
        <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'decorative' } }} label="image.alt: decorative" checked={decorative} onChange={(_, data) => setDecorative(data.checked === true)} />
        <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'logo-callback' } }} label="Logo render callback" checked={logoCallback} onChange={(_, data) => setLogoCallback(data.checked === true)} />
        <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'custom-fallback' } }} label="previewFallback: custom" checked={customFallback} onChange={(_, data) => setCustomFallback(data.checked === true)} />
        <Button data-cards-demo="inspect-refs" icon={<Info20Regular aria-hidden="true" />} onClick={() => setObservations(`root:${rootRef.current?.tagName ?? 'absent'} title:${titleRef.current?.tagName ?? 'absent'} preview:${previewRef.current?.tagName ?? 'absent'} logo:${logoDivRef.current?.tagName ?? 'absent'} button:${buttonRef.current?.tagName ?? 'absent'} preview callbacks:${previewCalls.current} fallback callbacks:${fallbackCalls.current} logo callbacks:${logoCalls.current}`)}>Inspect refs and callbacks</Button>
        <p data-cards-demo="observations">{observations}</p><p>Root bubbled:{rootClicks} Selection requests:{selectionRequests} Selected:{String(selected)}</p>
        <p>Notes are preserved when changing type, orientation, width and image. Direction follows the host.</p>
      </>}>
      <div data-cards-demo="preview-container" className={narrow ? cardStyles.compact : cardStyles.subject}>
        <AutoSkeleton data-cards-demo="skeleton" loading={loading}>
          {kind === 'preview' ? <PreviewCard {...commonProps} data-cards-demo="preview-card" orientation={horizontal ? 'horizontal' : 'vertical'} preview={preview}
            image={{ src, alt: decorative ? '' : 'Illustrated quarterly report cover' }}
            previewFallback={customFallback ? { 'data-cards-demo': 'fallback-slot', children: (Tag, props) => { fallbackCalls.current++; return <Tag {...props}>Custom unavailable cover</Tag>; } } : 'Cover unavailable'}
            logo={logoCallback ? { ref: logoDivRef, 'data-cards-demo': 'logo-slot', children: (Tag, props) => { logoCalls.current++; return <Tag {...props}><Apps24Regular aria-hidden="true" /></Tag>; } } : { ref: logoDivRef, 'data-cards-demo': 'logo-slot', children: <Apps24Regular aria-hidden="true" /> }} />
            : <ContentCard {...commonProps} data-cards-demo="content-card" />}
        </AutoSkeleton>
      </div>
    </DemoCard>
  </section>;
};
const CardsPanel: React.FC = () => {
  const [rtl, setRtl] = React.useState(false);
  const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
  return <section data-cards-demo="panel"><FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}>
    <CardExamples name="main" configuration={<Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'direction' } }} label="Main RTL override (unchecked inherits host)" checked={rtl} onChange={(_, data) => setRtl(data.checked === true)} />} />
  </FluentProvider></section>;
};
export default CardsPanel;
