import * as React from 'react';
import { FluentProvider, Checkbox, Select, Button, Field, Input } from '@fluentui/react-components';
import { RelativeTime, RelativeTimeProps, useSx, background, minWidth } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';
import { useDemoStyles } from '../shared/demoStyles';

const formats: readonly NonNullable<RelativeTimeProps['format']>[] = ['long', 'short', 'narrow'];
const modes = ['fixed', 'live', 'frozen'] as const;
const RelativeTimeExamples: React.FC<{ directionControl: React.ReactNode }> = ({ directionControl }) => {
  const demoStyles = useDemoStyles();
  const [date, setDate] = React.useState('2026-01-31T12:00:00Z');
  const [reference, setReference] = React.useState('2026-02-28T12:00:00Z');
  const [locale, setLocale] = React.useState('en-US');
  const [scenario, setScenario] = React.useState('event');
  const [format, setFormat] = React.useState<NonNullable<RelativeTimeProps['format']>>('long');
  const [mode, setMode] = React.useState<typeof modes[number]>('fixed');
  const [showTooltip, setShowTooltip] = React.useState(true);
  const [timeZone, setTimeZone] = React.useState('UTC');
  const [fallback, setFallback] = React.useState('Date unavailable');
  const fallbackRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
  const [fallbackClicks, setFallbackClicks] = React.useState(0);
  const [slotTag, setSlotTag] = React.useState('not inspected');
  const rootRef = React.useMemo(() => React.createRef<HTMLTimeElement>(), []);
  const [observedTag, setObservedTag] = React.useState('not inspected');
  const [clicks, setClicks] = React.useState(0);
  const props: RelativeTimeProps = {
    dateTime: scenario === 'invalid' ? '2026-02-30T12:00:00Z' : scenario === 'slot' ? undefined : date, locale, format, autoUpdate: mode !== 'frozen', showTooltip,
    timeZone: timeZone || undefined, fallback: scenario === 'slot' ? { ref: fallbackRef, 'data-relative-time-demo': 'fallback-slot', onClick: () => setFallbackClicks(value => value + 1), children: (Component, slotProps) => <Component {...slotProps}>{fallback}</Component> } : fallback, referenceTime: mode === 'fixed' ? reference : undefined,
  };
  return <DemoCard title="RelativeTime" iconName="Clock" description="Show the latest update with a configurable language and reference clock."
    controls={<><Field label="Scenario"><Select select={{ as: 'select', ...{ 'data-demo-variant': 'relative-time' } }} value={scenario} onChange={(_, data) => setScenario(data.value)}><option value="event">Latest update</option><option value="invalid">Invalid date</option><option value="slot">Fallback content</option></Select></Field>
<Field label="Updated at"><Input input={{ as: 'input', ...{ 'data-relative-time-demo': 'dateTime' } }} value={date} onChange={(_event, data) => setDate(data.value)} /></Field>
<Field label="Language"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'locale' } }} value={locale} onChange={(_event, data) => setLocale(data.value)}>
          {['en-US', 'it-IT', 'ar-SA'].map(value => <option key={value}>{value}</option>)}
        </Select></Field>
<Field label="Format"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'format' } }} value={format} onChange={(_event, data) => {
          const selected = formats.find(value => value === data.value); if (selected) setFormat(selected);
        }}>{formats.map(value => <option key={value}>{value}</option>)}</Select></Field>
<Field label="Reference clock"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'mode' } }} value={mode} onChange={(_event, data) => {
          const selected = modes.find(value => value === data.value); if (selected) setMode(selected);
        }}>{modes.map(value => <option key={value}>{value}</option>)}</Select></Field>
<Checkbox input={{ as: 'input', ...{ 'data-relative-time-demo': 'showTooltip' } }} checked={showTooltip} onChange={(_event, data) => setShowTooltip(data.checked === true)} label="Show full date" /></>}
    commands={<>
      <Button type="button" onClick={() => setDate(new Date(Date.now() - 60000).toISOString())}>Simulate one minute ago</Button>
    </>}
    status={<p aria-live="polite">Clock: {mode} · Preview clicks: {clicks}</p>}
    diagnostics={<><div className={demoStyles.controls}><Field label="Time zone"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'timeZone' } }} value={timeZone} onChange={(_event, data) => setTimeZone(data.value)}>
          <option value="">Runtime zone (explicit fallback)</option><option>UTC</option><option>Europe/Rome</option><option>America/Los_Angeles</option>
        </Select></Field>
<Field label="Reference date"><Input input={{ as: 'input', ...{ 'data-relative-time-demo': 'referenceTime' } }} value={reference} onChange={(_event, data) => setReference(data.value)} /></Field>
<Checkbox input={{ as: 'input', ...{ 'data-relative-time-demo': 'autoUpdate' } }} checked={mode !== 'frozen'} onChange={(_event, data) => setMode(data.checked === true ? 'live' : 'frozen')} label="Automatic updates" />
<Field label="Fallback text"><Input input={{ as: 'input', ...{ 'data-relative-time-demo': 'fallback' } }} value={fallback} onChange={(_event, data) => setFallback(data.value)} /></Field>
<p>Fixed reference takes priority over autoUpdate. Frozen mode samples after commit when the event instant or clock mode changes; changing presentation keeps that sample.</p></div>
{directionControl}

<Button data-relative-time-demo="inspect-slot" type="button" onClick={() => setSlotTag(fallbackRef.current?.tagName.toLowerCase() ?? 'unmounted')}>Inspect fallback slot ref</Button>

<p>Fallback element: {slotTag}. Fallback clicks: {fallbackClicks}.</p>

<Button data-relative-time-demo="inspect-ref" type="button" onClick={() => setObservedTag(rootRef.current?.tagName.toLowerCase() ?? 'unmounted')}>Inspect root ref</Button>

<p data-relative-time-demo="observations">Ref element: {observedTag}. Preview clicks: {clicks}.</p>
</>}>
      <RelativeTime {...props} data-relative-time-demo="preview" ref={rootRef} onClick={() => setClicks(value => value + 1)} />
  </DemoCard>;
};
const RelativeTimePanel: React.FC = () => {
  const [rtl, setRtl] = React.useState(false);
  const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
  return <>

    <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}><RelativeTimeExamples directionControl={<Checkbox checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} label="RTL direction" />} /></FluentProvider>
  </>;
};
export default RelativeTimePanel;
