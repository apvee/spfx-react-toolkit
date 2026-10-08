import * as React from 'react';
import { Tooltip, useId } from '@fluentui/react-components';
import { useFluent_unstable } from '@fluentui/react-shared-contexts';
import { useStableCallback } from '../../hooks/useStableCallback';
import { parseInstant, getRelativeTime } from './relative-time.internal';
import type { RelativeTimeProps } from './relative-time.types';

/** Localized elapsed/remaining time with a semantic timestamp and full-date description.
 * Inherits typography/color: compose FormattedText or Badge instead of selecting visual variants.
 * Below one minute displays localized now. Days are 24 hours; months/years use clamped UTC anniversaries.
 * referenceTime freezes the reference; autoUpdate false samples after commit on date/mode changes.
 * Live mode updates at label boundaries, resynchronizes hourly and pauses in a hidden host document.
 * Tooltip defaults to keyboard-focusable time; no live announcements or generated inline styles.
 * Invalid instants render only fallback. Valid instants with invalid Intl configuration throw.
 * Requires native Intl locale/time-zone data; server imports are safe, hydration is not guaranteed.
 * @param props Input instant, localization, clock/tooltip options and native time attributes.
 * @returns A time root and description/Tooltip, or the supplied fallback for invalid input.
 * @example
 * <FormattedText variant="caption1" subtle>
 *   <RelativeTime dateTime="2026-01-01T12:00:00Z" locale="it-IT" timeZone="Europe/Rome" />
 * </FormattedText>
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
  if (!formatters) return <>{fallback}</>;
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
