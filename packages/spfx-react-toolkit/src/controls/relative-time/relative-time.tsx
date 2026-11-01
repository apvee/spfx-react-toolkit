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
