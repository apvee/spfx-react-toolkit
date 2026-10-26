import * as React from 'react';
import { Button, Field, Input, Select } from '@fluentui/react-components';
import { EditableText, type EditableTextBlock, type EditableTextChangeData } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';

const presets = ['all', 'paragraph', 'headings', 'h2-h4'] as const;
type Preset = typeof presets[number];
const blocks: Record<Preset, readonly EditableTextBlock[] | undefined> = {
  all: undefined,
  paragraph: ['p'],
  headings: ['h1', 'h2', 'h3', 'h4'],
  'h2-h4': ['h2', 'h3', 'h4'],
};
const initialHTML = '<h2><strong>Project</strong> update</h2>';

/** Pure React controlled example; SPFx persistence remains a caller concern. */
const EditableTextPanel: React.FC = () => {
  const [html, setHTML] = React.useState<string | undefined>(initialHTML);
  const [mode, setMode] = React.useState<'read' | 'edit'>('read');
  const [preset, setPreset] = React.useState<Preset>('all');
  const [fallback, setFallback] = React.useState('Literal <h2>fallback</h2>');
  const [placeholder, setPlaceholder] = React.useState('Write a project update');
  const [locale, setLocale] = React.useState('en-US');
  const [lastChange, setLastChange] = React.useState<EditableTextChangeData>();
  const [changes, setChanges] = React.useState(0);
  const onChange = React.useCallback((data: EditableTextChangeData) => {
    setHTML(data.html);
    setLastChange(data);
    setChanges(count => count + 1);
  }, []);
  return <DemoCard title="EditableText" iconName="Font" description="One controlled formatted block. Changing policy projects the same stored input without saving."
    controls={<>
      <Field label="Mode"><Select select={{ as: 'select', ...{ 'data-editable-text-demo': 'mode' } }} value={mode}
        onChange={event => setMode(event.target.value === 'edit' ? 'edit' : 'read')}><option value="read">Read</option><option value="edit">Edit</option></Select></Field>
      <Field label="Allowed blocks"><Select select={{ as: 'select', ...{ 'data-editable-text-demo': 'preset', 'data-demo-variant': 'editable-text' } }} value={preset}
        onChange={event => { const selected = presets.find(value => value === event.target.value); if (selected !== undefined) setPreset(selected); }}>
        {presets.map(value => <option key={value} value={value}>{value}</option>)}
      </Select></Field>
      <Field label="Literal fallback"><Input input={{ as: 'input', ...{ 'data-editable-text-demo': 'fallback-input' } }} value={fallback} onChange={(_event, data) => setFallback(data.value)} /></Field>
      <Field label="Empty editor hint"><Input input={{ as: 'input', ...{ 'data-editable-text-demo': 'placeholder' } }} value={placeholder} onChange={(_event, data) => setPlaceholder(data.value)} /></Field>
    </>}
    commands={<>
      <Button data-editable-text-demo="h2" onClick={() => setHTML(initialHTML)}>Load H2 HTML</Button>
      <Button data-editable-text-demo="fallback" onClick={() => setHTML(undefined)}>Use literal fallback</Button>
      <Button data-editable-text-demo="empty" onClick={() => setHTML('')}>Load empty value</Button>
    </>}
    status={<>
      <p data-editable-text-demo="status" aria-live="polite">Mode: {mode} · Policy: {preset} · Callbacks: {changes}</p>
      <p>Stored controlled input: <code data-editable-text-demo="stored">{html === undefined ? 'undefined (literal fallback)' : html === '' ? '"" (authoritative empty)' : html}</code></p>
      <p>The live subject below is the visible policy projection. Normalization does not rewrite the stored input; a user change saves its canonical HTML into React state.</p>
    </>}
    diagnostics={<>
      <Field label="Editor language"><Select select={{ as: 'select', ...{ 'data-editable-text-demo': 'locale' } }} value={locale} onChange={event => setLocale(event.target.value)}>
        <option value="en-US">English</option><option value="it-IT">Italiano</option><option value="fr-FR">Français</option>
      </Select></Field>
      <p>Last user callback: <code data-editable-text-demo="callback">{lastChange === undefined ? 'No user changes' : JSON.stringify(lastChange)}</code></p>
      <p>Preset changes retain this control and its source. Switching to Read destroys its editor session. Locale changes translate commands, not content.</p>
    </>}>
    <EditableText data-editable-text-demo="preview" aria-label="Project update visible policy projection" mode={mode} value={html}
      fallbackText={fallback} allowedBlocks={blocks[preset]} placeholder={placeholder} locale={locale} onChange={onChange} />
  </DemoCard>;
};
export default EditableTextPanel;
