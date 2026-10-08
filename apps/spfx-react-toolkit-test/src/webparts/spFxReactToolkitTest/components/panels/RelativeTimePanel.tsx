import * as React from 'react';
import { Badge, FluentProvider, webDarkTheme, webLightTheme } from '@fluentui/react-components';
import { RelativeTime, RelativeTimeProps, FormattedText, useSx, flex, gap, padding } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';

const formats: readonly NonNullable<RelativeTimeProps['format']>[] = ['long', 'short', 'narrow'];
const modes = ['fixed', 'live', 'frozen'] as const;
const RelativeTimeExamples: React.FC = () => {
  const sx = useSx();
  const [date, setDate] = React.useState('2026-01-31T12:00:00Z');
  const [reference, setReference] = React.useState('2026-02-28T12:00:00Z');
  const [locale, setLocale] = React.useState('en-US');
  const [format, setFormat] = React.useState<NonNullable<RelativeTimeProps['format']>>('long');
  const [mode, setMode] = React.useState<typeof modes[number]>('fixed');
  const [showTooltip, setShowTooltip] = React.useState(true);
  const [timeZone, setTimeZone] = React.useState('UTC');
  const [fallback, setFallback] = React.useState('Date unavailable');
  const rootRef = React.useMemo(() => React.createRef<HTMLTimeElement>(), []);
  const [observedTag, setObservedTag] = React.useState('not inspected');
  const [clicks, setClicks] = React.useState(0);
  const props: RelativeTimeProps = {
    dateTime: date, locale, format, autoUpdate: mode !== 'frozen', showTooltip,
    timeZone: timeZone || undefined, fallback, referenceTime: mode === 'fixed' ? reference : undefined,
  };
  return <DemoCard title="RelativeTime" iconName="Clock" description="Generic instant formatting; locale and time zone are chosen explicitly here.">
    <div className={sx(flex.column, gap.medium, padding.medium)}>
      <label>Event ISO instant{' '}<input data-relative-time-demo="dateTime" value={date} onChange={event => setDate(event.target.value)} /></label>
      <button type="button" onClick={() => setDate(new Date(Date.now() - 60000).toISOString())}>Use one minute ago</button>
      <div className={sx(flex.row, flex.wrap, gap.medium)}>
        <label>Locale{' '}<select data-relative-time-demo="locale" value={locale} onChange={event => setLocale(event.target.value)}>
          {['en-US', 'it-IT', 'ar-SA'].map(value => <option key={value}>{value}</option>)}
        </select></label>
        <label>Format{' '}<select data-relative-time-demo="format" value={format} onChange={event => {
          const selected = formats.find(value => value === event.target.value); if (selected) setFormat(selected);
        }}>{formats.map(value => <option key={value}>{value}</option>)}</select></label>
        <label>Clock mode{' '}<select data-relative-time-demo="mode" value={mode} onChange={event => {
          const selected = modes.find(value => value === event.target.value); if (selected) setMode(selected);
        }}>{modes.map(value => <option key={value}>{value}</option>)}</select></label>
        <label>Full-date zone{' '}<select data-relative-time-demo="timeZone" value={timeZone} onChange={event => setTimeZone(event.target.value)}>
          <option value="">Runtime zone (explicit fallback)</option><option>UTC</option><option>Europe/Rome</option><option>America/Los_Angeles</option>
        </select></label>
      </div>
      <label>Fixed reference ISO instant{' '}<input data-relative-time-demo="referenceTime" value={reference} onChange={event => setReference(event.target.value)} /></label>
      <label><input data-relative-time-demo="autoUpdate" type="checkbox" checked={mode !== 'frozen'} onChange={event => setMode(event.target.checked ? 'live' : 'frozen')} /> Auto update (switch to live/frozen)</label>
      <label><input data-relative-time-demo="showTooltip" type="checkbox" checked={showTooltip} onChange={event => setShowTooltip(event.target.checked)} /> Full-date tooltip</label>
      <label>Invalid-input fallback{' '}<input data-relative-time-demo="fallback" value={fallback} onChange={event => setFallback(event.target.value)} /></label>
      <p>Fixed reference takes priority over autoUpdate. Frozen mode samples after commit when the event instant or clock mode changes; changing presentation keeps that sample.</p>
      <FormattedText variant="caption1" subtle><RelativeTime {...props} data-relative-time-demo="preview" ref={rootRef} onClick={() => setClicks(value => value + 1)} /></FormattedText>
      <Badge appearance="outline"><RelativeTime {...props} showTooltip={false} /></Badge>
      <p>Invalid input: <RelativeTime dateTime="2026-02-30T12:00:00Z" fallback={fallback} /></p>
      <button data-relative-time-demo="inspect-ref" type="button" onClick={() => setObservedTag(rootRef.current?.tagName.toLowerCase() ?? 'unmounted')}>Inspect root ref</button>
      <p data-relative-time-demo="observations">Ref element: {observedTag}. Preview clicks: {clicks}.</p>
    </div>
  </DemoCard>;
};
const RelativeTimePanel: React.FC = () => {
  const [dark, setDark] = React.useState(false);
  const [rtl, setRtl] = React.useState(false);
  return <>
    <label><input type="checkbox" checked={dark} onChange={event => setDark(event.target.checked)} /> Dark theme</label>{' '}
    <label><input type="checkbox" checked={rtl} onChange={event => setRtl(event.target.checked)} /> Right-to-left</label>
    <FluentProvider theme={dark ? webDarkTheme : webLightTheme} dir={rtl ? 'rtl' : 'ltr'}><RelativeTimeExamples /></FluentProvider>
  </>;
};
export default RelativeTimePanel;
