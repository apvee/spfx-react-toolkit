import * as React from 'react';
import { FluentProvider } from '@fluentui/react-provider';
import { webLightTheme, webDarkTheme } from '@fluentui/react-theme';
import {
  FormattedText, FormattedTextProps, FormattedTextVariant, FormattedTextSemanticElement,
  useSx, typography, foreground, width, padding, gap, flex, borderWidth, borderStyle, borderColor,
} from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';

const variants = Object.keys(typography) as FormattedTextVariant[];
const elements: readonly FormattedTextSemanticElement[] = ['span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'];
const fonts: readonly NonNullable<FormattedTextProps['font']>[] = ['base', 'monospace', 'numeric'];
const alignments: readonly NonNullable<FormattedTextProps['align']>[] = ['start', 'center', 'end', 'justify'];
function Choice<T extends string>({ label, value, choices, onChange, ...attributes }: {
  readonly label: string; readonly value: T; readonly choices: readonly T[];
  readonly onChange: (value: T) => void; readonly 'data-formatted-text-demo': string;
}): React.ReactElement {
  return <label>{label}{' '}<select {...attributes} value={value} onChange={event => {
    const selected = choices.find(choice => choice === event.target.value);
    if (selected !== undefined) onChange(selected);
  }}>{choices.map(choice => <option key={choice} value={choice}>{choice}</option>)}</select></label>;
}

const FormattedTextExamples: React.FC = () => {
  const sx = useSx();
  const [source, setSource] = React.useState('**Formatted text** with *emphasis*, ~~removed text~~, `code`, and [a link](https://example.com). Edit this text to try long words and overflow.');
  const [variant, setVariant] = React.useState<FormattedTextVariant>('body1');
  const [as, setAs] = React.useState<FormattedTextSemanticElement>('span');
  const [font, setFont] = React.useState<NonNullable<FormattedTextProps['font']>>('base');
  const [align, setAlign] = React.useState<NonNullable<FormattedTextProps['align']>>('start');
  const [markdown, setMarkdown] = React.useState(true);
  const [subtle, setSubtle] = React.useState(false);
  const [highlight, setHighlight] = React.useState(false);
  const [italic, setItalic] = React.useState(false);
  const [underline, setUnderline] = React.useState(false);
  const [strikethrough, setStrikethrough] = React.useState(false);
  const [wrap, setWrap] = React.useState(true);
  const [truncate, setTruncate] = React.useState(false);
  const [maxLines, setMaxLines] = React.useState(0);
  const [previewWidth, setPreviewWidth] = React.useState(320);
  const rootRef = React.useRef<HTMLElement>(null);
  const [observedTag, setObservedTag] = React.useState('not inspected');
  const [clicks, setClicks] = React.useState(0);
  const controlsClass = sx(flex.row, flex.wrap, gap.medium);
  return <DemoCard title="FormattedText" iconName="Font" description="Edit content and change its typography, semantic element and overflow independently.">
    <div className={sx(flex.column, gap.medium, padding.medium)}>
      <label>Text{' '}<textarea data-formatted-text-demo="text" rows={4} value={source}
        className={sx(width.full)} onChange={event => setSource(event.target.value)} /></label>
      <div className={controlsClass}>
        <Choice<FormattedTextVariant> data-formatted-text-demo="variant" label="Variant" value={variant} choices={variants} onChange={setVariant} />
        <Choice<FormattedTextSemanticElement> data-formatted-text-demo="as" label="HTML element" value={as} choices={elements} onChange={setAs} />
        <Choice<NonNullable<FormattedTextProps['font']>> data-formatted-text-demo="font" label="Font family" value={font} choices={fonts} onChange={setFont} />
        <Choice<NonNullable<FormattedTextProps['align']>> data-formatted-text-demo="align" label="Alignment" value={align} choices={alignments} onChange={setAlign} />
      </div>
      <div className={controlsClass}>
        <label><input data-formatted-text-demo="markdown" type="checkbox" checked={markdown} onChange={event => setMarkdown(event.target.checked)} /> Markdown</label>
        <label><input data-formatted-text-demo="subtle" type="checkbox" checked={subtle} onChange={event => setSubtle(event.target.checked)} /> Subtle</label>
        <label><input data-formatted-text-demo="highlight" type="checkbox" checked={highlight} onChange={event => setHighlight(event.target.checked)} /> Highlight</label>
        <label><input data-formatted-text-demo="italic" type="checkbox" checked={italic} onChange={event => setItalic(event.target.checked)} /> Italic</label>
        <label><input data-formatted-text-demo="underline" type="checkbox" checked={underline} onChange={event => setUnderline(event.target.checked)} /> Underline</label>
        <label><input data-formatted-text-demo="strikethrough" type="checkbox" checked={strikethrough} onChange={event => setStrikethrough(event.target.checked)} /> Strikethrough</label>
        <label><input data-formatted-text-demo="wrap" type="checkbox" checked={wrap} onChange={event => setWrap(event.target.checked)} /> Wrap</label>
        <label><input data-formatted-text-demo="truncate" type="checkbox" checked={truncate} onChange={event => setTruncate(event.target.checked)} /> Truncate</label>
      </div>
      <div className={controlsClass}>
        <label>Max lines (0 disables){' '}<input data-formatted-text-demo="maxLines" type="number" min={0} step={1} value={maxLines}
          onChange={event => { const value = Number(event.target.value); if (Number.isInteger(value) && value >= 0) setMaxLines(value); }} /></label>
        <label>Preview width{' '}<input data-formatted-text-demo="width" type="range" min={160} max={640} step={20} value={previewWidth}
          onChange={event => setPreviewWidth(Number(event.target.value))} /> {previewWidth}px</label>
      </div>
      <p>The HTML element is independent of the variant. A valid maxLines takes priority and adds block clamp layout even to a span. For single-line ellipsis, select p and constrain the width.</p>
      <div style={{ maxWidth: '100%' }} className={sx(width.px(previewWidth), borderWidth.thin, borderStyle.solid, borderColor.primary)}>
        <FormattedText data-formatted-text-demo="preview" ref={rootRef} as={as} variant={variant} font={font} align={align}
          markdown={markdown} subtle={subtle} highlight={highlight} italic={italic} underline={underline} strikethrough={strikethrough}
          wrap={wrap} truncate={truncate} maxLines={maxLines || undefined}
          onClick={() => setClicks(value => value + 1)}>{source}</FormattedText>
      </div>
      <p className={sx(foreground.subtle)}>Comparison: this text uses the existing foreground.subtle descriptor.</p>
      <button data-formatted-text-demo="inspect-ref" type="button" onClick={() => setObservedTag(rootRef.current ? rootRef.current.tagName.toLowerCase() : 'unmounted')}>Inspect root ref</button>
      <p data-formatted-text-demo="observations" aria-live="polite">Ref element: {observedTag}. Preview clicks: {clicks}.</p>
      <FormattedText data-formatted-text-demo="mixed-jsx" as="p">{'**Markdown string** '}
        <button type="button" onClick={() => setClicks(value => value + 1)}>Authored JSX button</button>
        <React.Fragment>{' *fragment emphasis*'}</React.Fragment>
      </FormattedText>
    </div>
  </DemoCard>;
};

const FormattedTextPanel: React.FC = () => {
  const [rtl, setRtl] = React.useState(false);
  const [dark, setDark] = React.useState(false);
  return <>
    <label><input data-formatted-text-demo="rtl" type="checkbox" checked={rtl} onChange={event => setRtl(event.target.checked)} /> Right-to-left</label>{' '}
    <label><input data-formatted-text-demo="theme" type="checkbox" checked={dark} onChange={event => setDark(event.target.checked)} /> Dark theme</label>
    <FluentProvider theme={dark ? webDarkTheme : webLightTheme} dir={rtl ? 'rtl' : 'ltr'}><FormattedTextExamples /></FluentProvider>
  </>;
};
export default FormattedTextPanel;
