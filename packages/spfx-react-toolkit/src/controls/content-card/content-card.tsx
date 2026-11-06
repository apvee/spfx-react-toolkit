/**
 * Titled content and explicit caller actions composed with native Fluent Card interaction.
 *
 * @packageDocumentation
 * @docPage controls/content-card
 */
import * as React from 'react';
import { CardFrame } from '../cards/card-frame.internal';
import type { ContentCardProps } from './content-card.types';

/**
 * Group a content summary with explicit caller-owned actions and native Card interaction.
 * Literal strings retain their text; compose FormattedText explicitly for Markdown.
 * Requires the host Fluent theme and Griffel renderer; no SPFx provider is required.
 * Root ref identifies the native div. Disabled preserves child disabled flags; native Card may block pointer input with its overlay.
 * @param props - Native Fluent Card behavior, required title and optional summary/body/action slots.
 * @returns A native div Card with a heading, optional body and explicit footer regions.
 * @example Basic content card
 * ```tsx
 * import * as React from 'react';
 * import { ContentCard } from '@apvee/spfx-react-toolkit';
 *
 * <ContentCard title="Workspace" footer="Updated today" actions={<button>Open</button>}>
 *   Three members
 * </ContentCard>;
 * ```
 *
 * @useWhen A titled summary needs native Card interactions, caller content and explicit header/footer actions.
 * @avoidWhen Use PreviewCard when the layout requires a preview-media region.
 *
 * @scenario Configure content slots and native card interaction
 * ```tsx
 * import * as React from 'react';
 * import { Button, Checkbox, Field, FluentProvider, Input, Select, makeStyles } from '@fluentui/react-components';
 * import { Document16Regular, Share16Regular, ArrowRight20Regular } from '@fluentui/react-icons';
 * import { AutoSkeleton, ContentCard, FormattedText, useSx, background, minWidth, text } from '@apvee/spfx-react-toolkit';
 * import type { ContentCardProps } from '@apvee/spfx-react-toolkit';
 * type DemoRegions = {
 *   controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode;
 *   children?: React.ReactNode; diagnostics?: React.ReactNode;
 * };
 * const standalone = (regions: DemoRegions): React.ReactElement => <>
 *   {regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}
 * </>;
 * const useExampleStyles = makeStyles({ standard: { width: '100%', maxWidth: '520px' }, compact: { width: '100%', maxWidth: '280px' } });
 * type Format = 'literal' | 'formatted' | 'callback' | 'absent';
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * } = {}): React.ReactElement {
 *   const sx = useSx(); const styles = useExampleStyles();
 *   const [title, setTitle] = React.useState('Quarterly report');
 *   const [description, setDescription] = React.useState('Team results, goals and next steps.');
 *   const [note, setNote] = React.useState('To share at the team meeting.');
 *   const [format, setFormat] = React.useState<Format>('literal');
 *   const [narrow, setNarrow] = React.useState(false), [loading, setLoading] = React.useState(false);
 *   const [disabled, setDisabled] = React.useState(false), [rtl, setRtl] = React.useState(false);
 *   // Keep native selection controlled for this Card's lifetime; switching undefined/boolean warns.
 *   const [selectable, setSelectable] = React.useState(true), [selected, setSelected] = React.useState(false);
 *   const [headerClicks, setHeaderClicks] = React.useState(0), [footerClicks, setFooterClicks] = React.useState(0);
 *   const [rootClicks, setRootClicks] = React.useState(0), [requests, setRequests] = React.useState(0);
 *   const [observed, setObserved] = React.useState('Not inspected.');
 *   const rootRef = React.useMemo(() => React.createRef<HTMLDivElement>(), []);
 *   const titleRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
 *   const buttonRef = React.useMemo(() => React.createRef<HTMLButtonElement>(), []);
 *   const descriptionCalls = React.useRef(0);
 *   const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
 *   const summary: ContentCardProps['description'] = format === 'absent' ? undefined : format === 'formatted'
 *     ? <FormattedText>{description}</FormattedText> : format === 'callback'
 *       ? { children: (Tag, props) => { descriptionCalls.current++; return <Tag {...props}>{description}</Tag>; } }
 *       : { children: description };
 *   const previousReset = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision;
 *     setTitle('Quarterly report'); setDescription('Team results, goals and next steps.'); setNote('To share at the team meeting.');
 *     setFormat('literal'); setNarrow(false); setLoading(false); setDisabled(false); setRtl(false);
 *     setSelectable(true); setSelected(false); setHeaderClicks(0); setFooterClicks(0); setRootClicks(0); setRequests(0);
 *     setObserved('Not inspected.'); descriptionCalls.current = 0;
 *   }, [resetRevision]);
 *   return <FluentProvider data-cards-demo="provider" dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}>
 *     <section data-cards-demo="instance" data-card-instance="content-primary">
 *       {render({
 *         controls: <>
 *           <Field label="Title"><Input input={{ as: 'input', ...{ 'data-cards-demo': 'title-input' } }} value={title} onChange={(_, data) => setTitle(data.value)} /></Field>
 *           <Field label="Description"><Input input={{ as: 'input', ...{ 'data-cards-demo': 'description-input' } }} value={description} onChange={(_, data) => setDescription(data.value)} /></Field>
 *           <Field label="Description presentation"><Select data-cards-demo="format" data-demo-variant="content-format" value={format} onChange={event => {
 *             const value = event.target.value; if (value === 'literal' || value === 'formatted' || value === 'callback' || value === 'absent') setFormat(value);
 *           }}><option value="literal">Literal object children</option><option value="formatted">Explicit FormattedText</option><option value="callback">Render callback</option><option value="absent">Absent description</option></Select></Field>
 *           <Field label="Width"><Select data-cards-demo="width" value={narrow ? 'compact' : 'standard'} onChange={event => setNarrow(event.target.value === 'compact')}><option value="standard">Standard · up to 520 px</option><option value="compact">Compact · up to 280 px</option></Select></Field>
 *           <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'loading' } }} label="Loading (AutoSkeleton)" checked={loading} onChange={(_, data) => setLoading(data.checked === true)} />
 *           <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'disabled' } }} label="Disabled" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} />
 *         </>,
 *         status: <output data-cards-demo="counters" aria-live="polite">Header:{headerClicks} Footer:{footerClicks} Root bubbled:{rootClicks} Selection requests:{requests} Selected:{String(selected)}</output>,
 *         diagnostics: <>
 *           <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'selectable' } }} label="Accept selection requests" checked={selectable} onChange={(_, data) => setSelectable(data.checked === true)} />
 *           <Checkbox input={{ as: 'input', ...{ 'data-cards-demo': 'direction' } }} label="RTL override (unchecked inherits host)" checked={rtl} onChange={(_, data) => setRtl(data.checked === true)} />
 *           <Button data-cards-demo="inspect-refs" onClick={() => setObserved(`root:${rootRef.current?.tagName ?? 'absent'} title:${titleRef.current?.tagName ?? 'absent'} button:${buttonRef.current?.tagName ?? 'absent'} description callbacks:${descriptionCalls.current}`)}>Inspect refs and callbacks</Button>
 *           <output data-cards-demo="observations">{observed}</output>
 *           <p>Strings stay literal; FormattedText is explicit. Selection stays controlled; the caller can decline requests. Native Card owns focus and disabled pointer behavior; child buttons keep their own disabled flags. Reset restores local values without replacing this card.</p>
 *         </>,
 *         children: <div data-cards-demo="preview-container" className={narrow ? styles.compact : styles.standard}>
 *           <AutoSkeleton data-cards-demo="skeleton" loading={loading}>
 *             <ContentCard data-cards-demo="content-card" ref={rootRef} headingLevel={2}
 *               title={{ ref: titleRef, 'data-cards-demo': 'title-slot', children: title || 'Untitled report' }} description={summary}
 *               icon={<Document16Regular aria-hidden="true" />} disabled={disabled}
 *               selected={selected} onSelectionChange={(_, data) => { if (selectable) setSelected(data.selected); setRequests(value => value + 1); }}
 *               checkbox={{ 'aria-label': 'Select content report' }} focusMode="tab-exit" onClick={() => setRootClicks(value => value + 1)}
 *               headerActions={{ children: <Button data-cards-demo="header-action" ref={buttonRef} className={sx(text.noWrap)} icon={<Share16Regular aria-hidden="true" />} appearance="subtle" size="small" onClick={() => setHeaderClicks(value => value + 1)}>Share</Button> }}
 *               footer={{ children: 'Updated today · Team Operations' }}
 *               actions={{ children: <Button data-cards-demo="footer-action" icon={<ArrowRight20Regular aria-hidden="true" />} appearance="primary" size="small" onClick={() => setFooterClicks(value => value + 1)}>Open report</Button> }}>
 *               <Field label="Retained notes"><Input input={{ as: 'input', ...{ 'data-cards-demo': 'retained-input' } }} value={note} onChange={(_, data) => setNote(data.value)} /></Field>
 *             </ContentCard>
 *           </AutoSkeleton>
 *         </div>,
 *       })}
 *     </section>
 *   </FluentProvider>;
 * }
 * export const demo = {
 *   host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Edit the report, compare literal and formatted slots, activate header and footer actions, request native selection and retain the note through loading, presentation and reset.',
 * };
 * ```
 */
export const ContentCard = React.forwardRef<HTMLDivElement, ContentCardProps>((props, ref) => <CardFrame {...props} ref={ref} orientation="vertical" />);
ContentCard.displayName = 'ContentCard';
