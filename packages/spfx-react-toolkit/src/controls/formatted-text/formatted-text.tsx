/**
 * Semantic text, toolkit typography and bounded inline Markdown.
 *
 * @packageDocumentation
 * @docPage controls/formatted-text
 */
import * as React from 'react';
import { useSx } from '../../hooks/useSx';
import * as typography from '../../helpers/styles/typography';
import * as text from '../../helpers/styles/text';
import * as textAlign from '../../helpers/styles/text-align';
import * as foreground from '../../helpers/styles/foreground';
import { none as noBlockMargin } from '../../helpers/styles/margin-block';
import { extraSmall as inlinePadding } from '../../helpers/styles/padding-inline';
import { small as smallRadius } from '../../helpers/styles/border-radius';
import { alternative as codeBackground } from '../../helpers/styles/background';
import { useFormattedTextStyles, useLineClampClass } from './formatted-text.styles.internal';
import { renderFormattedTextChildren } from './formatted-text-content.internal';
import type { FormattedTextProps } from './formatted-text.types';

/**
 * Render toolkit typography with bounded inline Markdown or caller-owned JSX.
 * `as` defaults to span independently of variant. Native attributes/events and ref reach that root.
 * Uses the host's Fluent theme and Griffel renderer; no SPFx provider is required.
 * Valid maxLines applies block clamp layout without changing the tag and wins over truncate/wrap.
 * Unsafe Markdown links preserve their label; raw HTML and unsupported document syntax stay literal.
 * Strings in arrays and `React.Fragment` receive Markdown. Use a fragment to group strings with JSX
 * actions: a wrapper such as `<span>` leaves descendant strings unparsed, even with `markdown={true}`.
 * Arbitrary JSX remains untouched. Ellipsis needs constrained width and suitable element layout.
 * Each owned string above 64 KiB UTF-8 falls back literally before parsing. More than 2,000
 * original AST nodes or depth above 32 (root depth 1) also returns the complete string literally.
 * Node/depth checks run after parsing and do not bound parser work; see the FormattedText guide.
 * @param props - Native text attributes, presentation options and owned content.
 * @returns A semantic text root, including when children render nothing.
 * @example Inline Markdown heading
 * ```tsx
 * import * as React from 'react';
 * import { FormattedText } from '@apvee/spfx-react-toolkit';
 * <FormattedText as="h2" variant="title2" maxLines={2}>
 *   Read the **updates** in [documentation](./docs).
 * </FormattedText>
 * ```
 *
 * @example Mixed content in a fragment
 * ```tsx
 * import * as React from 'react';
 * import { FormattedText } from '@apvee/spfx-react-toolkit';
 * <FormattedText>
 *   <React.Fragment>
 *     {'**Project update** '}
 *     <button type="button">Open project</button>
 *     {' *review available*'}
 *   </React.Fragment>
 * </FormattedText>
 * ```
 *
 * @useWhen Semantic text needs toolkit typography, bounded inline Markdown or overflow presentation.
 * @avoidWhen Use EditableText when the user must edit and serialize rich text.
 *
 * @example Explore semantic text and inline Markdown
 * ```tsx
 * // This standalone React scenario uses the host Fluent theme and Griffel renderer.
 * import * as React from 'react';
 * import { FluentProvider } from '@fluentui/react-provider';
 * import { Button, Checkbox, Field, Input, Select, Textarea, makeStyles, tokens } from '@fluentui/react-components';
 * import {
 *   FormattedText, FormattedTextProps, FormattedTextVariant, FormattedTextSemanticElement,
 *   useSx, typography, width, background, minWidth,
 * } from '@apvee/spfx-react-toolkit';
 *
 * const useStyles = makeStyles({ frame: { maxWidth: '100%' },
 *   controls: { display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: tokens.spacingVerticalM, minWidth: 0,
 *     '& > *': { minWidth: 0, maxWidth: '100%' }, '@media (max-width: 600px)': { gridTemplateColumns: 'minmax(0, 1fr)' } },
 *   controlRow: { gridColumn: '1 / -1' },
 * });
 * function parsePreviewWidth(draft: string): number | undefined {
 *   const value = Number(draft);
 *   return draft.trim() !== '' && Number.isInteger(value) && value >= 160 && value <= 640 ? value : undefined;
 * }
 *
 * const variants = Object.keys(typography) as FormattedTextVariant[];
 * const elements: readonly FormattedTextSemanticElement[] = ['span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'];
 * const fonts: readonly NonNullable<FormattedTextProps['font']>[] = ['base', 'monospace', 'numeric'];
 * const alignments: readonly NonNullable<FormattedTextProps['align']>[] = ['start', 'center', 'end', 'justify'];
 * function Choice<T extends string>({ label, value, choices, onChange, ...attributes }: {
 *   readonly label: string; readonly value: T; readonly choices: readonly T[];
 *   readonly onChange: (value: T) => void; readonly 'data-formatted-text-demo': string;
 * }): React.ReactElement {
 *   return <Field label={label}><Select select={{ as: 'select', ...attributes, 'aria-label': label }} value={value} onChange={(_event, data) => {
 *     const selected = choices.find(choice => choice === data.value);
 *     if (selected !== undefined) onChange(selected);
 *   }}>{choices.map(choice => <option key={choice} value={choice}>{choice}</option>)}</Select></Field>;
 * }
 *
 * type DemoRegions = {
 *   controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode;
 *   diagnostics?: React.ReactNode; children?: React.ReactNode;
 * };
 * const standalone = (regions: DemoRegions): React.ReactElement => <>
 *   {regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}
 * </>;
 * export function Demo({ render = standalone, resetRevision = 0 }: {
 *   render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number;
 * }): React.ReactElement {
 *   const [rtl, setRtl] = React.useState(false);
 *   const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
 *   const directionControl = <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'rtl' } }} checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} label="RTL direction" />;
 *   const sx = useSx();
 *   const styles = useStyles();
 *   const [source, setSource] = React.useState('The **intranet project** is ready for review. Read the [work plan](https://example.com) and share your comments.');
 *   const [contentMode, setContentMode] = React.useState('text');
 *   const [variant, setVariant] = React.useState<FormattedTextVariant>('body1');
 *   const [as, setAs] = React.useState<FormattedTextSemanticElement>('span');
 *   const [font, setFont] = React.useState<NonNullable<FormattedTextProps['font']>>('base');
 *   const [align, setAlign] = React.useState<NonNullable<FormattedTextProps['align']>>('start');
 *   const [markdown, setMarkdown] = React.useState(true);
 *   const [subtle, setSubtle] = React.useState(false);
 *   const [highlight, setHighlight] = React.useState(false);
 *   const [italic, setItalic] = React.useState(false);
 *   const [underline, setUnderline] = React.useState(false);
 *   const [strikethrough, setStrikethrough] = React.useState(false);
 *   const [wrap, setWrap] = React.useState(true);
 *   const [truncate, setTruncate] = React.useState(false);
 *   const [maxLines, setMaxLines] = React.useState(0);
 *   const [previewWidth, setPreviewWidth] = React.useState(320);
 *   const [widthDraft, setWidthDraft] = React.useState('320');
 *   const rootRef = React.useMemo(() => React.createRef<HTMLElement>(), []);
 *   const [observedTag, setObservedTag] = React.useState('not inspected');
 *   const [clicks, setClicks] = React.useState(0);
 *   const demoStyles = useStyles();
 *   const previousReset = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision;
 *     setSource('The **intranet project** is ready for review. Read the [work plan](https://example.com) and share your comments.');
 *     setContentMode('text'); setVariant('body1'); setAs('span'); setFont('base'); setAlign('start');
 *     setMarkdown(true); setSubtle(false); setHighlight(false); setItalic(false); setUnderline(false);
 *     setStrikethrough(false); setWrap(true); setTruncate(false); setMaxLines(0); setPreviewWidth(320);
 *     setWidthDraft('320'); setObservedTag('not inspected'); setClicks(0); setRtl(false);
 *   }, [resetRevision]);
 *   return render({ controls: <><Field label="Content"><Select select={{ as: 'select', ...{ 'data-demo-variant': 'formatted-text', 'data-formatted-text-demo': 'content' } }} value={contentMode} onChange={(_, data) => setContentMode(data.value)}><option value="text">Text</option><option value="jsx">Text and JSX action</option></Select></Field>
 * <Field label="Text" className={demoStyles.controlRow}><Textarea textarea={{ as: 'textarea', ...{ 'data-formatted-text-demo': 'text' } }} rows={4} value={source}
 *         className={sx(width.full)} onChange={(_event, data) => setSource(data.value)} /></Field>
 * <Choice<FormattedTextVariant> data-formatted-text-demo="variant" label="Typography" value={variant} choices={variants} onChange={setVariant} />
 * <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'markdown' } }} checked={markdown} onChange={(_event, data) => setMarkdown(data.checked === true)} label="Markdown" />
 * <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'wrap' } }} checked={wrap} onChange={(_event, data) => setWrap(data.checked === true)} label="Wrap text" />
 * <Field label="Width" hint="Enter a whole number from 160 to 640. Invalid drafts keep the last valid preview and reset on blur."
 *           validationState={parsePreviewWidth(widthDraft) === undefined ? 'error' : 'none'}><Input input={{ as: 'input', ...{ 'data-formatted-text-demo': 'width' } }} type="number" min={160} max={640} step={20} value={widthDraft}
 *           onChange={(_event, data) => {
 *             setWidthDraft(data.value);
 *             const value = parsePreviewWidth(data.value);
 *             if (value !== undefined) setPreviewWidth(value);
 *           }} onBlur={() => setWidthDraft(String(previewWidth))} /> {previewWidth}px</Field></>,
 *     status: <p aria-live="polite">Element: {as} · Variant: {variant} · Preview clicks: {clicks}</p>,
 *     diagnostics: <><div className={demoStyles.controls}><Choice<FormattedTextSemanticElement> data-formatted-text-demo="as" label="HTML element" value={as} choices={elements} onChange={setAs} />
 * <Choice<NonNullable<FormattedTextProps['font']>> data-formatted-text-demo="font" label="Font" value={font} choices={fonts} onChange={setFont} />
 * <Choice<NonNullable<FormattedTextProps['align']>> data-formatted-text-demo="align" label="Alignment" value={align} choices={alignments} onChange={setAlign} />
 * <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'subtle' } }} checked={subtle} onChange={(_event, data) => setSubtle(data.checked === true)} label="Subtle color" />
 * <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'highlight' } }} checked={highlight} onChange={(_event, data) => setHighlight(data.checked === true)} label="Highlighted" />
 * <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'italic' } }} checked={italic} onChange={(_event, data) => setItalic(data.checked === true)} label="Italic" />
 * <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'underline' } }} checked={underline} onChange={(_event, data) => setUnderline(data.checked === true)} label="Underlined" />
 * <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'strikethrough' } }} checked={strikethrough} onChange={(_event, data) => setStrikethrough(data.checked === true)} label="Strikethrough" />
 * <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'truncate' } }} checked={truncate} onChange={(_event, data) => setTruncate(data.checked === true)} label="Truncation" />
 * <Field label="Maximum lines (0 disables)"><Input input={{ as: 'input', ...{ 'data-formatted-text-demo': 'maxLines' } }} type="number" min={0} step={1} value={String(maxLines)}
 *           onChange={(_event, data) => { const value = Number(data.value); if (Number.isInteger(value) && value >= 0) setMaxLines(value); }} /></Field>
 * <p>The HTML element is independent of the variant. A valid maxLines takes priority and adds block clamp layout even to a span. For single-line ellipsis, select p and constrain the width.</p></div>
 * {directionControl}
 *
 * <Button data-formatted-text-demo="inspect-ref" type="button" onClick={() => setObservedTag(rootRef.current ? rootRef.current.tagName.toLowerCase() : 'unmounted')}>Inspect root ref</Button>
 *
 * <p data-formatted-text-demo="observations" aria-live="polite">Ref element: {observedTag}. Preview clicks: {clicks}.</p>
 * </>,
 *     children: <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}><div className={sx(styles.frame, width.px(previewWidth))}>
 *         <FormattedText data-formatted-text-demo="preview" ref={rootRef} as={as} variant={variant} font={font} align={align}
 *           markdown={markdown} subtle={subtle} highlight={highlight} italic={italic} underline={underline} strikethrough={strikethrough}
 *           wrap={wrap} truncate={truncate} maxLines={maxLines || undefined}
 *           onClick={() => setClicks(value => value + 1)}>{contentMode === 'jsx' ? <React.Fragment>
 *             {'**Project update** '}
 *             <Button data-formatted-text-demo="mixed-jsx" type="button">Open project</Button>
 *             <React.Fragment>{' *review available*'}</React.Fragment>
 *           </React.Fragment> : source}</FormattedText>
 *       </div></FluentProvider> });
 * }
 *
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Change semantic tags, typography and overflow independently, and activate the retained mixed JSX action.' };
 * ```
 */
export const FormattedText = React.forwardRef<HTMLElement, FormattedTextProps>((props, ref) => {
  const {
    as: Root = 'span', variant = 'body1', font = 'base', align = 'start',
    markdown = true, subtle = false, highlight = false, italic = false,
    underline = false, strikethrough = false, wrap = true, truncate = false,
    maxLines, children, className, style, ...nativeProps
  } = props;
  // Never propagate an alternative content pathway, even from untyped callers.
  delete nativeProps.dangerouslySetInnerHTML;
  const sx = useSx();
  const styles = useFormattedTextStyles();
  const clamp = maxLines !== undefined && Number.isFinite(maxLines) && Number.isInteger(maxLines) && maxLines > 0;
  const lineClampClass = useLineClampClass(clamp ? maxLines : undefined);
  const content = renderFormattedTextChildren(children, {
    markdown,
    codeClassName: sx(codeBackground, inlinePadding, smallRadius, styles.code),
    linkClassName: sx(styles.link, highlight && styles.highlightedLink),
  });
  return React.createElement(Root, {
    ...nativeProps, ref,
    className: sx(
      styles.root, noBlockMargin, typography[variant],
      align === 'justify' ? styles.justify : textAlign[align],
      font === 'monospace' && styles.monospace, font === 'numeric' && styles.numeric,
      subtle && foreground.subtle, italic && styles.italic,
      underline && strikethrough ? styles.combinedDecoration : underline ? styles.underline : strikethrough && styles.strikethrough,
      clamp ? text.wrap : truncate ? text.truncate : wrap ? text.wrap : text.noWrap,
      clamp && styles.clamp, lineClampClass, className,
    ),
    style,
  }, highlight ? <mark className={sx(foreground.primary, inlinePadding, smallRadius, styles.mark)}>{content}</mark> : content);
});
FormattedText.displayName = 'FormattedText';
