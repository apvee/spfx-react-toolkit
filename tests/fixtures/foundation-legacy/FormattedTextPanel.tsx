/** Regression snapshot captured before single-subject consolidation; not product UI. */
import * as React from 'react';
import { FluentProvider } from '@fluentui/react-provider';
import { Button, Checkbox, Field, Input, Select, Textarea, makeStyles } from '@fluentui/react-components';
import {
  FormattedText, FormattedTextProps, FormattedTextVariant, FormattedTextSemanticElement,
  useSx, typography, foreground, width, background, minWidth,
} from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/demoStyles';

const useStyles = makeStyles({ frame: { maxWidth: '100%' } });
function parsePreviewWidth(draft: string): number | undefined {
  const value = Number(draft);
  return draft.trim() !== '' && Number.isInteger(value) && value >= 160 && value <= 640 ? value : undefined;
}

const variants = Object.keys(typography) as FormattedTextVariant[];
const elements: readonly FormattedTextSemanticElement[] = ['span', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'];
const fonts: readonly NonNullable<FormattedTextProps['font']>[] = ['base', 'monospace', 'numeric'];
const alignments: readonly NonNullable<FormattedTextProps['align']>[] = ['start', 'center', 'end', 'justify'];
function Choice<T extends string>({ label, value, choices, onChange, ...attributes }: {
  readonly label: string; readonly value: T; readonly choices: readonly T[];
  readonly onChange: (value: T) => void; readonly 'data-formatted-text-demo': string;
}): React.ReactElement {
  return <Field label={label}><Select select={{ as: 'select', ...attributes, 'aria-label': label }} value={value} onChange={(_event, data) => {
    const selected = choices.find(choice => choice === data.value);
    if (selected !== undefined) onChange(selected);
  }}>{choices.map(choice => <option key={choice} value={choice}>{choice}</option>)}</Select></Field>;
}

const FormattedTextExamples: React.FC<{ directionControl: React.ReactNode }> = ({ directionControl }) => {
  const sx = useSx();
  const styles = useStyles();
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
  const [widthDraft, setWidthDraft] = React.useState('320');
  const rootRef = React.useMemo(() => React.createRef<HTMLElement>(), []);
  const [observedTag, setObservedTag] = React.useState('not inspected');
  const [clicks, setClicks] = React.useState(0);
  const demoStyles = useDemoStyles();
  const controlsClass = demoStyles.controls;
  return <DemoCard title="FormattedText" iconName="Font" description="Edit content and change its typography, semantic element and overflow independently."
    controls={<div className={demoStyles.controls}>
      {directionControl}
      <Field label="Text"><Textarea textarea={{ as: 'textarea', ...{ 'data-formatted-text-demo': 'text' } }} rows={4} value={source}
        className={sx(width.full)} onChange={(_event, data) => setSource(data.value)} /></Field>
      <div className={controlsClass}>
        <Choice<FormattedTextVariant> data-formatted-text-demo="variant" label="Variant" value={variant} choices={variants} onChange={setVariant} />
        <Choice<FormattedTextSemanticElement> data-formatted-text-demo="as" label="HTML element" value={as} choices={elements} onChange={setAs} />
        <Choice<NonNullable<FormattedTextProps['font']>> data-formatted-text-demo="font" label="Font family" value={font} choices={fonts} onChange={setFont} />
        <Choice<NonNullable<FormattedTextProps['align']>> data-formatted-text-demo="align" label="Alignment" value={align} choices={alignments} onChange={setAlign} />
      </div>
      <div className={controlsClass}>
        <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'markdown' } }} checked={markdown} onChange={(_event, data) => setMarkdown(data.checked === true)} label="Markdown" />
        <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'subtle' } }} checked={subtle} onChange={(_event, data) => setSubtle(data.checked === true)} label="Subtle" />
        <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'highlight' } }} checked={highlight} onChange={(_event, data) => setHighlight(data.checked === true)} label="Highlight" />
        <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'italic' } }} checked={italic} onChange={(_event, data) => setItalic(data.checked === true)} label="Italic" />
        <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'underline' } }} checked={underline} onChange={(_event, data) => setUnderline(data.checked === true)} label="Underline" />
        <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'strikethrough' } }} checked={strikethrough} onChange={(_event, data) => setStrikethrough(data.checked === true)} label="Strikethrough" />
        <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'wrap' } }} checked={wrap} onChange={(_event, data) => setWrap(data.checked === true)} label="Wrap" />
        <Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'truncate' } }} checked={truncate} onChange={(_event, data) => setTruncate(data.checked === true)} label="Truncate" />
      </div>
      <div className={controlsClass}>
        <Field label="Max lines (0 disables)"><Input input={{ as: 'input', ...{ 'data-formatted-text-demo': 'maxLines' } }} type="number" min={0} step={1} value={String(maxLines)}
          onChange={(_event, data) => { const value = Number(data.value); if (Number.isInteger(value) && value >= 0) setMaxLines(value); }} /></Field>
        <Field label="Preview width" hint="Enter a whole number from 160 to 640. Invalid drafts keep the last valid preview and reset on blur."
          validationState={parsePreviewWidth(widthDraft) === undefined ? 'error' : 'none'}><Input input={{ as: 'input', ...{ 'data-formatted-text-demo': 'width' } }} type="number" min={160} max={640} step={20} value={widthDraft}
          onChange={(_event, data) => {
            setWidthDraft(data.value);
            const value = parsePreviewWidth(data.value);
            if (value !== undefined) setPreviewWidth(value);
          }} onBlur={() => setWidthDraft(String(previewWidth))} /> {previewWidth}px</Field>
      </div>
      <p>The HTML element is independent of the variant. A valid maxLines takes priority and adds block clamp layout even to a span. For single-line ellipsis, select p and constrain the width.</p>
    </div>}
    status={<p aria-live="polite">Element: {as} · Variant: {variant} · Preview clicks: {clicks}</p>}
    diagnostics={<>
      <p className={sx(foreground.subtle)}>Comparison: this text uses the existing foreground.subtle descriptor.</p>
      <Button data-formatted-text-demo="inspect-ref" type="button" onClick={() => setObservedTag(rootRef.current ? rootRef.current.tagName.toLowerCase() : 'unmounted')}>Inspect root ref</Button>
      <p data-formatted-text-demo="observations" aria-live="polite">Ref element: {observedTag}. Preview clicks: {clicks}.</p>
      <FormattedText data-formatted-text-demo="mixed-jsx" as="p">{'**Markdown string** '}
        <Button type="button" onClick={() => setClicks(value => value + 1)}>Authored JSX button</Button>
        <React.Fragment>{' *fragment emphasis*'}</React.Fragment>
      </FormattedText>

      </>}>
      <div className={sx(styles.frame, width.px(previewWidth))}>
        <FormattedText data-formatted-text-demo="preview" ref={rootRef} as={as} variant={variant} font={font} align={align}
          markdown={markdown} subtle={subtle} highlight={highlight} italic={italic} underline={underline} strikethrough={strikethrough}
          wrap={wrap} truncate={truncate} maxLines={maxLines || undefined}
          onClick={() => setClicks(value => value + 1)}>{source}</FormattedText>
      </div>
  </DemoCard>;
};

const FormattedTextPanel: React.FC = () => {
  const [rtl, setRtl] = React.useState(false);
  const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
  return <>

    <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}><FormattedTextExamples directionControl={<Checkbox input={{ as: 'input', ...{ 'data-formatted-text-demo': 'rtl' } }} checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} label="Override host direction to RTL" />} /></FluentProvider>
  </>;
};
export default FormattedTextPanel;
