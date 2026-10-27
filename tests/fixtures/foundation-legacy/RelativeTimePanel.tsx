/** Regression snapshot captured before single-subject consolidation; not product UI. */
import * as React from 'react';
import { Badge, FluentProvider, Checkbox, Select, Button, Field, Input } from '@fluentui/react-components';
import { RelativeTime, RelativeTimeProps, FormattedText, useSx, background, minWidth } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/demoStyles';

const formats: readonly NonNullable<RelativeTimeProps['format']>[] = ['long', 'short', 'narrow'];
const modes = ['fixed', 'live', 'frozen'] as const;
const RelativeTimeExamples: React.FC<{ directionControl: React.ReactNode }> = ({ directionControl }) => {
  const demoStyles = useDemoStyles();
  const [date, setDate] = React.useState('2026-01-31T12:00:00Z');
  const [reference, setReference] = React.useState('2026-02-28T12:00:00Z');
  const [locale, setLocale] = React.useState('en-US');
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
    dateTime: date, locale, format, autoUpdate: mode !== 'frozen', showTooltip,
    timeZone: timeZone || undefined, fallback, referenceTime: mode === 'fixed' ? reference : undefined,
  };
  return <DemoCard title="RelativeTime" iconName="Clock" description="Generic instant formatting; locale and time zone are chosen explicitly here."
    controls={<div className={demoStyles.controls}>
      {directionControl}
      <Field label="Event ISO instant"><Input input={{ as: 'input', ...{ 'data-relative-time-demo': 'dateTime' } }} value={date} onChange={(_event, data) => setDate(data.value)} /></Field>

      <div className={demoStyles.controls}>
        <Field label="Locale"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'locale' } }} value={locale} onChange={(_event, data) => setLocale(data.value)}>
          {['en-US', 'it-IT', 'ar-SA'].map(value => <option key={value}>{value}</option>)}
        </Select></Field>
        <Field label="Format"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'format' } }} value={format} onChange={(_event, data) => {
          const selected = formats.find(value => value === data.value); if (selected) setFormat(selected);
        }}>{formats.map(value => <option key={value}>{value}</option>)}</Select></Field>
        <Field label="Clock mode"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'mode' } }} value={mode} onChange={(_event, data) => {
          const selected = modes.find(value => value === data.value); if (selected) setMode(selected);
        }}>{modes.map(value => <option key={value}>{value}</option>)}</Select></Field>
        <Field label="Full-date zone"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'timeZone' } }} value={timeZone} onChange={(_event, data) => setTimeZone(data.value)}>
          <option value="">Runtime zone (explicit fallback)</option><option>UTC</option><option>Europe/Rome</option><option>America/Los_Angeles</option>
        </Select></Field>
      </div>
      <Field label="Fixed reference ISO instant"><Input input={{ as: 'input', ...{ 'data-relative-time-demo': 'referenceTime' } }} value={reference} onChange={(_event, data) => setReference(data.value)} /></Field>
      <Checkbox input={{ as: 'input', ...{ 'data-relative-time-demo': 'autoUpdate' } }} checked={mode !== 'frozen'} onChange={(_event, data) => setMode(data.checked === true ? 'live' : 'frozen')} label="Auto update (switch to live/frozen)" />
      <Checkbox input={{ as: 'input', ...{ 'data-relative-time-demo': 'showTooltip' } }} checked={showTooltip} onChange={(_event, data) => setShowTooltip(data.checked === true)} label="Full-date tooltip" />
      <Field label="Invalid-input fallback"><Input input={{ as: 'input', ...{ 'data-relative-time-demo': 'fallback' } }} value={fallback} onChange={(_event, data) => setFallback(data.value)} /></Field>
      <p>Fixed reference takes priority over autoUpdate. Frozen mode samples after commit when the event instant or clock mode changes; changing presentation keeps that sample.</p>
    </div>}
    commands={<>
      <Button type="button" onClick={() => setDate(new Date(Date.now() - 60000).toISOString())}>Use one minute ago</Button>
    </>}
    status={<p aria-live="polite">Clock: {mode} · Preview clicks: {clicks}</p>}
    diagnostics={<>
      <p>Invalid input: <RelativeTime dateTime="2026-02-30T12:00:00Z" fallback={fallback} /></p>
      <p>Fallback slot: <RelativeTime dateTime={undefined} fallback={{ ref: fallbackRef, 'data-relative-time-demo': 'fallback-slot',
        onClick: () => setFallbackClicks(value => value + 1), children: (Component, slotProps) => <Component {...slotProps}>{fallback}</Component> }} /></p>
      <Button data-relative-time-demo="inspect-slot" type="button" onClick={() => setSlotTag(fallbackRef.current?.tagName.toLowerCase() ?? 'unmounted')}>Inspect fallback slot ref</Button>
      <p>Fallback element: {slotTag}. Fallback clicks: {fallbackClicks}.</p>
      <Button data-relative-time-demo="inspect-ref" type="button" onClick={() => setObservedTag(rootRef.current?.tagName.toLowerCase() ?? 'unmounted')}>Inspect root ref</Button>
      <p data-relative-time-demo="observations">Ref element: {observedTag}. Preview clicks: {clicks}.</p>

      </>}>
      <FormattedText variant="caption1" subtle><RelativeTime {...props} data-relative-time-demo="preview" ref={rootRef} onClick={() => setClicks(value => value + 1)} /></FormattedText>
      <Badge appearance="outline"><RelativeTime {...props} showTooltip={false} /></Badge>
  </DemoCard>;
};
const RelativeTimePanel: React.FC = () => {
  const [rtl, setRtl] = React.useState(false);
  const directionSx = useSx({ dir: rtl ? 'rtl' : undefined });
  return <>

    <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}><RelativeTimeExamples directionControl={<Checkbox checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} label="Override host direction to RTL" />} /></FluentProvider>
  </>;
};
export default RelativeTimePanel;
