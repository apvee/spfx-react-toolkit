import * as React from 'react';
import { Accordion, AccordionItem, AccordionHeader, Button, Card, CardHeader, Checkbox, Field, Input, Select, Text, FluentProvider } from '@fluentui/react-components';
import { ArrowRight20Regular, ArrowSync20Regular, Checkmark20Regular, ImageOff20Regular, Info20Regular, MoreHorizontal20Regular, Share16Regular } from '@fluentui/react-icons';
import { DemoCard, DemoAccordionPanel, DemoIcon, useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared';

import { AutoSkeleton, ContentCard, PreviewCard, FormattedText, useSx, width, text, background, minWidth } from '@apvee/spfx-react-toolkit';
import type { PreviewCardProps } from '@apvee/spfx-react-toolkit';

const cover = (color: string): string => `data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360"><rect width="640" height="360" fill="${color}"/><circle cx="320" cy="180" r="70" fill="white"/></svg>`)}`;
const firstCover = cover('#0078d4'), secondCover = cover('#8764b8');
// Official Apps24Regular SVG asset keeps the native img/ref scenario observable.
// The callback variant renders the same Fluent icon through DemoIcon.
const logoImage = 'data:image/svg+xml,' + encodeURIComponent("<svg fill=\"#0078d4\" width=\"24\" height=\"24\" viewBox=\"0 0 24 24\" xmlns=\"http://www.w3.org/2000/svg\"><path d=\"m18.5 2.33 3.17 3.18c.88.88.88 2.3 0 3.18l-2.58 2.59A2.25 2.25 0 0 1 21 13.5v5.25C21 20 20 21 18.75 21H5.25C4.01 21 3 20 3 18.75V5.25C3 4.01 4 3 5.25 3h5.25c1.13 0 2.06.83 2.23 1.92l2.58-2.59c.88-.88 2.3-.88 3.18 0Zm-14 16.42c0 .41.34.75.75.75h6v-6.75H4.5v6Zm8.25.75h6c.41 0 .75-.34.75-.75V13.5a.75.75 0 0 0-.75-.75h-6v6.75Zm-2.25-15H5.25a.75.75 0 0 0-.75.75v6h6.75v-6a.75.75 0 0 0-.75-.75Zm2.25 4.81v1.94h1.94l-1.94-1.94Zm3.62-5.92L13.2 6.57c-.3.3-.3.77 0 1.06l3.18 3.18c.3.3.77.3 1.06 0l3.18-3.18c.3-.3.3-.77 0-1.06l-3.18-3.18a.75.75 0 0 0-1.06 0Z\" fill=\"#0078d4\"></path></svg>");
const modes = ['image', 'custom', 'empty-object', 'empty-callback', 'false', 'empty-fragment'] as const;
type PreviewMode = typeof modes[number];

/** Private per-instance scenario, reused by the independent regression fixture. */
export const CardExamples: React.FC<{ name: string; configuration?: React.ReactNode }> = ({ name, configuration }) => {
  const sx = useSx(); const styles = useDemoStyles();
  const [horizontal, setHorizontal] = React.useState(false), [narrow, setNarrow] = React.useState(false);
  const [loading, setLoading] = React.useState(false), [disabled, setDisabled] = React.useState(false);
  const [decorative, setDecorative] = React.useState(false), [customFallback, setCustomFallback] = React.useState(false);
  const [logoCallback, setLogoCallback] = React.useState(false);
  const [mode, setMode] = React.useState<PreviewMode>('image');
  const [title, setTitle] = React.useState('**Literal report title**');
  const [description, setDescription] = React.useState('**Literal** description with enough words to demonstrate the two-line limit in narrow cards. JSX can own its presentation.');
  const [src, setSrc] = React.useState(firstCover);
  const [headerClicks, setHeaderClicks] = React.useState(0), [actionClicks, setActionClicks] = React.useState(0);
  const [rootClicks, setRootClicks] = React.useState(0), [selectionRequests, setSelectionRequests] = React.useState(0);
  const [selected, setSelected] = React.useState(false), [nativeSelected, setNativeSelected] = React.useState(false);
  const [observations, setObservations] = React.useState('Not inspected');
  const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []), titleRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
  const previewRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []), logoRef = React.useMemo(() => React.createRef<HTMLImageElement>(), []);
  const logoDivRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
  const buttonRef = React.useMemo(() => React.createRef<HTMLButtonElement>(), []);
  const previewCalls = React.useRef(0), fallbackCalls = React.useRef(0), logoCalls = React.useRef(0);
  const preview: PreviewCardProps['preview'] = mode === 'image' ? undefined : mode === 'custom' ? {
    ref: previewRef, 'data-cards-demo': 'preview-slot', children: (Tag, props) => { previewCalls.current++; return <Tag {...props}><img src={secondCover} alt="Caller-owned custom cover" /></Tag>; },
  } : mode === 'empty-object' ? { ref: previewRef, 'data-cards-demo': 'preview-slot' }
    : mode === 'empty-callback' ? { children: () => { previewCalls.current++; return undefined; } }
      : mode === 'false' ? false : <></>;
  return <section data-cards-demo="instance" data-card-instance={name} className={styles.panel}>
    <DemoCard diagnosticsLabel="Advanced diagnostics" title={name === 'main' ? 'PreviewCard' : `${name} cards`} iconName="Page" description="Configure PreviewCard props. Media commands and native comparisons exercise retained content, slots and selection."
      controls={<>
        {configuration}
        <Field label="title"><Input input={{ as: 'input', ...{ 'data-cards-demo': 'title-input' } }} value={title} onChange={(_event, data) => setTitle(data.value)} /></Field>
        <Field label="description"><Input input={{ as: 'input', ...{ 'data-cards-demo': 'description-input' } }} value={description} onChange={(_event, data) => setDescription(data.value)} /></Field>
        <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'orientation' } }} label="orientation: horizontal" checked={horizontal} onChange={(_event, data) => setHorizontal(data.checked === true)} />
        <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'width' } }} label="Preview width: 240px" checked={narrow} onChange={(_event, data) => setNarrow(data.checked === true)} />
        <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'loading' } }} label="AutoSkeleton loading" checked={loading} onChange={(_event, data) => setLoading(data.checked === true)} />
        <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'disabled' } }} label="disabled" checked={disabled} onChange={(_event, data) => setDisabled(data.checked === true)} />
        <Field label="preview"><Select data-cards-demo="preview-mode" value={mode} onChange={event => { const value = event.target.value; if (modes.some(candidate => candidate === value)) setMode(value as PreviewMode); }}>{modes.map(value => <option key={value} value={value}>{value}</option>)}</Select></Field>
        <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'decorative' } }} label="image.alt: decorative" checked={decorative} onChange={(_event, data) => setDecorative(data.checked === true)} />
        <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'logo-callback' } }} label="logo: div callback" checked={logoCallback} onChange={(_event, data) => setLogoCallback(data.checked === true)} />
        <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'custom-fallback' } }} label="previewFallback: custom" checked={customFallback} onChange={(_event, data) => setCustomFallback(data.checked === true)} />
      </>}
      commands={<>
        <div className={styles.actions} role="group" aria-label="Media simulation commands">
          <Button data-cards-demo="fail-image" icon={<ImageOff20Regular aria-hidden="true" />} onClick={() => setSrc('/cards-demo-missing-image.png')}>Use failing resource</Button>
          <Button data-cards-demo="new-src" icon={<ArrowSync20Regular aria-hidden="true" />} onClick={() => setSrc(current => current === firstCover ? secondCover : firstCover)}>Load new source</Button>
        </div>
      </>}
      status={<p data-cards-demo="counters">Header:{headerClicks} Footer:{actionClicks} Root bubbled:{rootClicks} Selection requests:{selectionRequests} Selected:{String(selected)} Native selected:{String(nativeSelected)}</p>}
      diagnostics={<>
          <Button data-cards-demo="inspect-refs" icon={<Info20Regular aria-hidden="true" />} onClick={() => setObservations(`root:${rootRef.current?.tagName ?? 'absent'} title:${titleRef.current?.tagName ?? 'absent'} preview:${previewRef.current?.tagName ?? 'absent'} logo:${(logoCallback ? logoDivRef.current : logoRef.current)?.tagName ?? 'absent'} button:${buttonRef.current?.tagName ?? 'absent'} preview callbacks:${previewCalls.current} fallback callbacks:${fallbackCalls.current} logo callbacks:${logoCalls.current}`)}>Inspect refs and callback counts</Button>
        <p data-cards-demo="observations">{observations}</p>
        <p>Root clicks may bubble from child controls. Disabled Card keeps child disabled flags unchanged; native overlay blocks pointer input. Compare Enter on child controls with the native selectable Card.</p>
      </>}>
      <div className={styles.scroll} role="region" aria-label={`${name} card preview width`} tabIndex={0}>
        <div data-cards-demo="preview-container" className={sx(width.px(narrow ? 240 : 520))}>
          <AutoSkeleton data-cards-demo="skeleton" loading={loading}>
            <PreviewCard data-cards-demo="preview-card" ref={rootRef} title={{ ref: titleRef, 'data-cards-demo': 'title-slot', className: 'cards-consumer-title', children: title }}
              headingLevel={2} description={description}
              orientation={horizontal ? 'horizontal' : 'vertical'} disabled={disabled} preview={preview}
              image={{ src, alt: decorative ? '' : 'White circle on a colored report cover' }}
              previewFallback={customFallback ? { 'data-cards-demo': 'fallback-slot', children: (Tag, props) => { fallbackCalls.current++; return <Tag {...props}>Custom unavailable cover</Tag>; } } : undefined}
              logo={logoCallback ? { ref: logoDivRef, 'data-cards-demo': 'logo-slot', children: (Tag, props) => { logoCalls.current++; return <Tag {...props}><DemoIcon name="AppIconDefault" /></Tag>; } } : { as: 'img', ref: logoRef, src: logoImage, alt: 'Brand', 'data-cards-demo': 'logo-slot' }}
              icon={<DemoIcon name="Page" />}
              headerActions={{ 'data-cards-demo': 'header-slot', children: <Button data-cards-demo="header-action" ref={buttonRef} icon={<Share16Regular aria-hidden="true" />} className={sx(text.noWrap)} appearance="subtle" size="small" onClick={() => setHeaderClicks(value => value + 1)}>Share</Button> }}
              actions={{ 'data-cards-demo': 'actions-slot', children: <Button data-cards-demo="footer-action" icon={<ArrowRight20Regular aria-hidden="true" />} appearance="primary" onClick={() => setActionClicks(value => value + 1)}>Open report</Button> }}
              footer={{ 'data-cards-demo': 'footer-slot', children: 'Supplemental update information' }} onClick={() => setRootClicks(value => value + 1)}>
              <Field label="Retained notes"><Input input={{ as: 'input', ...{ 'data-cards-demo': 'retained-input' } }} defaultValue="Edit before changing layout" /></Field>
              <FormattedText>Explicit **FormattedText** composition</FormattedText>
            </PreviewCard>
          </AutoSkeleton>
          <Accordion collapsible><AccordionItem value="comparisons"><AccordionHeader>Selection and native Card comparison</AccordionHeader><DemoAccordionPanel>
            <ContentCard data-cards-demo="content-card" title="**Literal content title**" description={{ children: 'Object description stays literal' }} icon={<DemoIcon name="ContactCard" />}
              selected={selected} onSelectionChange={(_, data) => { setSelected(data.selected); setSelectionRequests(value => value + 1); }}
              checkbox={{ 'aria-label': `Select ${name} content` }} focusMode="tab-exit" floatingAction={<Button data-cards-demo="floating-action" icon={<MoreHorizontal20Regular aria-hidden="true" />} appearance="subtle">Floating</Button>}
              actions={<Button data-cards-demo="selection-child" icon={<Checkmark20Regular aria-hidden="true" />}>Child control</Button>} footer="Selection delegates to Fluent">
              <Field label="Content notes"><Input input={{ as: 'input', ...{ 'data-cards-demo': 'content-input' } }} defaultValue="Stable content" /></Field>
            </ContentCard>
            <Card data-cards-demo="native-card" disabled={disabled} selected={nativeSelected} onSelectionChange={(_, data) => setNativeSelected(data.selected)} focusMode="tab-exit">
              <CardHeader image={<DemoIcon name="ContactCard" />} header={<Text as="h4" weight="semibold" className={styles.heading}>Native Card comparison</Text>} />
              <Button data-cards-demo="native-child" icon={<Checkmark20Regular aria-hidden="true" />}>Native child control</Button>
            </Card>
          </DemoAccordionPanel></AccordionItem></Accordion>
        </div></div>
    </DemoCard>
  </section>;
};
const CardsPanel: React.FC = () => {
  const [rtl, setRtl] = React.useState(false);
  const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
  return <section data-cards-demo="regression-panel"><FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}>
    <CardExamples name="regression" configuration={<Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'direction' } }} label="Main RTL override (unchecked inherits host)" checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} />} />
  </FluentProvider></section>;
};
export default CardsPanel;
