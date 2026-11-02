/**
 * Localized relative instants with controlled/live clocks and full-date descriptions.
 *
 * @packageDocumentation
 * @docPage controls/relative-time
 */
import * as React from 'react';
import { Tooltip, useId } from '@fluentui/react-components';
import { useFluent_unstable } from '@fluentui/react-shared-contexts';
import { renderContentSlot } from '../content-slot.internal';
import { mergeClasses } from '@griffel/react';
import { useStableCallback } from '../../hooks/useStableCallback';
import { parseInstant, getRelativeTime } from './relative-time.internal';
import type { RelativeTimeProps } from './relative-time.types';

/**
 * Localized elapsed/remaining time with a semantic timestamp and full-date description.
 * Inherits typography/color: compose FormattedText or Badge instead of selecting visual variants.
 * Below one minute displays localized now. Days are 24 hours; months/years use clamped UTC anniversaries.
 * referenceTime freezes the reference; autoUpdate false samples after commit on date/mode changes.
 * Live mode updates at label boundaries, resynchronizes hourly and pauses in a hidden host document.
 * Tooltip defaults to keyboard-focusable time; no live announcements or generated inline styles.
 * Invalid instants render only fallback. Valid instants with invalid Intl configuration throw.
 * Requires native Intl locale/time-zone data; server imports are safe, hydration is not guaranteed.
 * @param props - Input instant, localization, clock/tooltip options and native time attributes.
 * @returns A time root and description/Tooltip, or the supplied fallback for invalid input.
 * @example Basic usage
 * ```tsx
 * import * as React from 'react';
 * import { FormattedText, RelativeTime } from '@apvee/spfx-react-toolkit';
 * <FormattedText variant="caption1" subtle>
 *   <RelativeTime dateTime="2026-01-01T12:00:00Z" locale="it-IT" timeZone="Europe/Rome" />
 * </FormattedText>
 * ```
 *
 * @useWhen An absolute instant should display a localized relative label and optional full-date description.
 * @avoidWhen Use an explicit calendar formatter for date-only values, which are not instants.
 *
 *
 * @example Separate SPFx UI culture from the site time zone
 * ```tsx
 * import * as React from 'react';
 * import { RelativeTime, useSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
 * // Mount under the initialized matching provider. Unknown groups have no site-zone label.
 * function ModifiedTime({ modified }: { modified: Date }) {
 *   const info = useSPFxLocaleInfo();
 *   return info.ianaTimeZone === undefined ? <span>Site time zone unavailable</span> :
 *     <RelativeTime dateTime={modified} locale={info.uiLocale} timeZone={info.ianaTimeZone} />;
 * }
 * ```
 *
 *
 * @example Customize an invalid-date fallback without replacing the time ref
 * ```tsx
 * import * as React from 'react';
 * import { RelativeTime } from '@apvee/spfx-react-toolkit';
 * function OptionalDate({ dateTime }: { dateTime: Date | string | number | undefined }) {
 *   const fallbackRef = React.useRef<HTMLSpanElement>(null); // React 17's legacy ref initializer.
 *   return <RelativeTime dateTime={dateTime} fallback={{ ref: fallbackRef,
 *     children: (Component, props) => <Component {...props}>Date unavailable</Component>,
 *   }} />;
 * }
 * ```
 *
 * @example Explore relative clocks and fallback slots
 * ```tsx
 * // This standalone React scenario uses the host Fluent theme and Griffel renderer.
 * import * as React from 'react';
 * import { FluentProvider, Checkbox, Select, Button, Field, Input, makeStyles, tokens } from '@fluentui/react-components';
 * import { RelativeTime, RelativeTimeProps, useSx, background, minWidth } from '@apvee/spfx-react-toolkit';
 *
 * const useStyles = makeStyles({
 *   controls: { display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: tokens.spacingVerticalM, minWidth: 0,
 *     '& > *': { minWidth: 0, maxWidth: '100%' }, '@media (max-width: 600px)': { gridTemplateColumns: 'minmax(0, 1fr)' } },
 *   controlRow: { gridColumn: '1 / -1' },
 * });
 * const formats: readonly NonNullable<RelativeTimeProps['format']>[] = ['long', 'short', 'narrow'];
 * const modes = ['fixed', 'live', 'frozen'] as const;
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
 *   const directionControl = <Checkbox input={{ as: 'input', ...{ 'data-relative-time-demo': 'direction' } }} checked={rtl} onChange={(_event, data) => setRtl(data.checked === true)} label="RTL direction" />;
 *   const demoStyles = useStyles();
 *   const [date, setDate] = React.useState('2026-01-31T12:00:00Z');
 *   const [reference, setReference] = React.useState('2026-02-28T12:00:00Z');
 *   const [locale, setLocale] = React.useState('en-US');
 *   const [scenario, setScenario] = React.useState('event');
 *   const [format, setFormat] = React.useState<NonNullable<RelativeTimeProps['format']>>('long');
 *   const [mode, setMode] = React.useState<typeof modes[number]>('fixed');
 *   const [showTooltip, setShowTooltip] = React.useState(true);
 *   const [timeZone, setTimeZone] = React.useState('UTC');
 *   const [fallback, setFallback] = React.useState('Date unavailable');
 *   const fallbackRef = React.useMemo(() => React.createRef<HTMLSpanElement>(), []);
 *   const [fallbackClicks, setFallbackClicks] = React.useState(0);
 *   const [slotTag, setSlotTag] = React.useState('not inspected');
 *   const rootRef = React.useMemo(() => React.createRef<HTMLTimeElement>(), []);
 *   const [observedTag, setObservedTag] = React.useState('not inspected');
 *   const [clicks, setClicks] = React.useState(0);
 *   const props: RelativeTimeProps = {
 *     dateTime: scenario === 'invalid' ? '2026-02-30T12:00:00Z' : scenario === 'slot' ? undefined : date, locale, format, autoUpdate: mode !== 'frozen', showTooltip,
 *     timeZone: timeZone || undefined, fallback: scenario === 'slot' ? { ref: fallbackRef, 'data-relative-time-demo': 'fallback-slot', onClick: () => setFallbackClicks(value => value + 1), children: (Component, slotProps) => <Component {...slotProps}>{fallback}</Component> } : fallback, referenceTime: mode === 'fixed' ? reference : undefined,
 *   };
 *   const previousReset = React.useRef(resetRevision);
 *   React.useEffect(() => {
 *     if (previousReset.current === resetRevision) return;
 *     previousReset.current = resetRevision;
 *     setDate('2026-01-31T12:00:00Z'); setReference('2026-02-28T12:00:00Z');
 *     setLocale('en-US'); setScenario('event'); setFormat('long'); setMode('fixed'); setShowTooltip(true);
 *     setTimeZone('UTC'); setFallback('Date unavailable'); setFallbackClicks(0); setSlotTag('not inspected');
 *     setObservedTag('not inspected'); setClicks(0); setRtl(false);
 *   }, [resetRevision]);
 *   return render({ controls: <><Field label="Scenario"><Select select={{ as: 'select', ...{ 'data-demo-variant': 'relative-time', 'data-relative-time-demo': 'scenario' } }} value={scenario} onChange={(_, data) => setScenario(data.value)}><option value="event">Latest update</option><option value="invalid">Invalid date</option><option value="slot">Fallback content</option></Select></Field>
 * <Field label="Updated at"><Input input={{ as: 'input', ...{ 'data-relative-time-demo': 'dateTime' } }} value={date} onChange={(_event, data) => setDate(data.value)} /></Field>
 * <Field label="Language"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'locale' } }} value={locale} onChange={(_event, data) => setLocale(data.value)}>
 *           {['en-US', 'it-IT', 'ar-SA'].map(value => <option key={value}>{value}</option>)}
 *         </Select></Field>
 * <Field label="Format"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'format' } }} value={format} onChange={(_event, data) => {
 *           const selected = formats.find(value => value === data.value); if (selected) setFormat(selected);
 *         }}>{formats.map(value => <option key={value}>{value}</option>)}</Select></Field>
 * <Field label="Reference clock"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'mode' } }} value={mode} onChange={(_event, data) => {
 *           const selected = modes.find(value => value === data.value); if (selected) setMode(selected);
 *         }}>{modes.map(value => <option key={value}>{value}</option>)}</Select></Field>
 * <Checkbox input={{ as: 'input', ...{ 'data-relative-time-demo': 'showTooltip' } }} checked={showTooltip} onChange={(_event, data) => setShowTooltip(data.checked === true)} label="Show full date" /></>,
 *     commands: <>
 *       <Button type="button" onClick={() => setDate(new Date(Date.now() - 60000).toISOString())}>Simulate one minute ago</Button>
 *     </>,
 *     status: <p aria-live="polite">Clock: {mode} · Preview clicks: {clicks}</p>,
 *     diagnostics: <><div className={demoStyles.controls}><Field label="Time zone"><Select select={{ as: 'select', ...{ 'data-relative-time-demo': 'timeZone' } }} value={timeZone} onChange={(_event, data) => setTimeZone(data.value)}>
 *           <option value="">Runtime zone (explicit fallback)</option><option>UTC</option><option>Europe/Rome</option><option>America/Los_Angeles</option>
 *         </Select></Field>
 * <Field label="Reference date"><Input input={{ as: 'input', ...{ 'data-relative-time-demo': 'referenceTime' } }} value={reference} onChange={(_event, data) => setReference(data.value)} /></Field>
 * <Checkbox input={{ as: 'input', ...{ 'data-relative-time-demo': 'autoUpdate' } }} checked={mode !== 'frozen'} onChange={(_event, data) => setMode(data.checked === true ? 'live' : 'frozen')} label="Automatic updates" />
 * <Field label="Fallback text"><Input input={{ as: 'input', ...{ 'data-relative-time-demo': 'fallback' } }} value={fallback} onChange={(_event, data) => setFallback(data.value)} /></Field>
 * <p>Fixed reference takes priority over autoUpdate. Frozen mode samples after commit when the event instant or clock mode changes; changing presentation keeps that sample.</p></div>
 * {directionControl}
 *
 * <Button data-relative-time-demo="inspect-slot" type="button" onClick={() => setSlotTag(fallbackRef.current?.tagName.toLowerCase() ?? 'unmounted')}>Inspect fallback slot ref</Button>
 *
 * <p>Fallback element: {slotTag}. Fallback clicks: {fallbackClicks}.</p>
 *
 * <Button data-relative-time-demo="inspect-ref" type="button" onClick={() => setObservedTag(rootRef.current?.tagName.toLowerCase() ?? 'unmounted')}>Inspect root ref</Button>
 *
 * <p data-relative-time-demo="observations">Ref element: {observedTag}. Preview clicks: {clicks}.</p>
 * </>,
 *     children: <FluentProvider dir={rtl ? 'rtl' : undefined} className={directionSx(background.transparent, minWidth.zero)}><RelativeTime {...props} data-relative-time-demo="preview" ref={rootRef} onClick={() => setClicks(value => value + 1)} /></FluentProvider> });
 * }
 *
 * export const demo = { host: 'none', kind: 'snapshot', contentPlacement: 'subject',
 *   observe: 'Compare fixed, live and frozen clocks, then inspect native time and fallback slot refs and events.' };
 * ```
 */
export const RelativeTime = React.forwardRef<HTMLTimeElement, RelativeTimeProps>((props, ref) => {
  const {
    dateTime, locale, format = 'long', autoUpdate = true, showTooltip = true,
    timeZone, fallback, referenceTime, tabIndex, lang,
    'aria-describedby': describedBy, ...nativeProps
  } = props;
  delete nativeProps.children;
  delete nativeProps.dangerouslySetInnerHTML;
  const timestamp = parseInstant(dateTime);
  const fixed = referenceTime === undefined ? undefined : parseInstant(referenceTime);
  const valid = Number.isFinite(timestamp) && (fixed === undefined || Number.isFinite(fixed));
  const { targetDocument } = useFluent_unstable();
  const [clock, setClock] = React.useState(Date.now);
  const descriptionId = useId('relative-time-description');
  const refreshClock = useStableCallback(() => {
    const now = Date.now();
    setClock(now);
    const next = getRelativeTime(timestamp, now).nextUpdate;
    return Number.isFinite(next) ? Math.max(1, Math.min(next - now, 3600000)) : 3600000;
  });
  React.useEffect(() => {
    if (!valid || fixed !== undefined) return;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const update = (): void => {
      if (timer !== undefined) clearTimeout(timer);
      const delay = refreshClock();
      if (autoUpdate && !targetDocument?.hidden) timer = setTimeout(update, delay);
    };
    update();
    if (autoUpdate) targetDocument?.addEventListener('visibilitychange', update);
    return () => {
      if (timer !== undefined) clearTimeout(timer);
      if (autoUpdate) targetDocument?.removeEventListener('visibilitychange', update);
    };
  }, [timestamp, fixed, valid, autoUpdate, targetDocument, refreshClock]);

  const formatters = React.useMemo(() => valid ? {
    relative: new Intl.RelativeTimeFormat(locale, { style: format, numeric: 'always' }),
    now: new Intl.RelativeTimeFormat(locale, { style: format, numeric: 'auto' }),
    full: new Intl.DateTimeFormat(locale, {
      year: 'numeric', month: 'long', day: 'numeric', hour: 'numeric', minute: '2-digit',
      second: '2-digit', timeZoneName: 'short', timeZone,
    }),
  } : undefined, [valid, locale, format, timeZone]);
  if (!formatters) return <>{renderContentSlot(fallback, 'span', mergeClasses, { shorthand: 'preserve' })}</>;
  const relative = getRelativeTime(timestamp, fixed ?? clock);
  const text = (relative.unit === 'second' ? formatters.now : formatters.relative).format(relative.value, relative.unit);
  const fullDate = formatters.full.format(timestamp);
  const resolvedLanguage = lang ?? formatters.relative.resolvedOptions().locale;
  const time = <time {...nativeProps} ref={ref} dateTime={new Date(timestamp).toISOString()}
    lang={resolvedLanguage} tabIndex={tabIndex ?? (showTooltip ? 0 : undefined)}
    aria-describedby={[describedBy, descriptionId].filter(Boolean).join(' ')}>{text}</time>;
  return showTooltip ? <Tooltip relationship="description" content={{ id: descriptionId, children: fullDate, lang: resolvedLanguage }}>{time}</Tooltip>
    : <>{time}<span id={descriptionId} lang={resolvedLanguage} hidden>{fullDate}</span></>;
});
RelativeTime.displayName = 'RelativeTime';
