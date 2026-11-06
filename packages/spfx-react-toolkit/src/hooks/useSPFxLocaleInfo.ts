/**
 * Host cultures and audited regional identifiers guide formatting and direction.
 *
 * @packageDocumentation
 * @docPage hooks/environment
 */
// useSPFxLocaleInfo.ts
// Hook for locale and regional settings

import { useMemo } from 'react';
import { getSPFxLocaleInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * Raw SharePoint regional time-zone metadata read from optional web.timeZoneInfo.
 * The toolkit preserves this host object without normalizing offsets.
 * Availability depends on the SPFx host; this shape is not exported by the
 * pinned sp-page-context package. Prefer ianaTimeZone for Intl formatting.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxTimeZone, useSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const zone: SPFxTimeZone | undefined = useSPFxLocaleInfo().timeZone;
 *   console.log(zone?.id, zone?.description);
 *   return <span>{zone?.description ?? 'Unavailable'}</span>;
 * }
 * ```
 */
export interface SPFxTimeZone {
  /** Time zone ID (numeric identifier) */
  readonly id: number;
  
  /** Raw host regional offset in minutes, preserved without normalization. */
  readonly offset: number;
  
  /** Time zone description (e.g., "Pacific Standard Time") */
  readonly description: string;
  
  /** Raw host daylight offset in minutes, preserved without normalization. */
  readonly daylightOffset: number;
  
  /** Raw host standard offset in minutes, preserved without normalization. */
  readonly standardOffset: number;
}

/**
 * Host content/UI cultures, preserved regional time-zone metadata and an optional audited IANA representative.
 *
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { SPFxLocaleInfo, useSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Example() {
 *   const locale: SPFxLocaleInfo = useSPFxLocaleInfo();
 *   console.log(locale.locale, locale.ianaTimeZone);
 *   return <span>{locale.locale}</span>;
 * }
 * ```
 */
export interface SPFxLocaleInfo {
  /** Current locale (e.g., "en-US", "it-IT") */
  readonly locale: string;
  
  /** Current UI locale (may differ from content locale) */
  readonly uiLocale: string;
  
  /** Original optional web.timeZoneInfo host metadata; undefined when unavailable. */
  readonly timeZone: SPFxTimeZone | undefined;

  /**
   * Audited CLDR territory001 IANA representative for the standard SharePoint
   * timeZone.id group. Undefined for missing, unknown or unverified IDs.
   * This represents a group, not the site's physical city or user profile zone;
   * customized SharePoint Server TIMEZONE.XML rules are not verified.
   * No locale, offset, description or browser time-zone fallback is inferred.
   */
  readonly ianaTimeZone?: string;
  
  /** Whether the language is right-to-left */
  readonly isRtl: boolean;
}

/**
 * Hook for locale and regional settings
 * Call inside a React component or hook beneath a host-specific SPFx provider.
 * 
 * Provides locale and regional information from SPFx PageContext:
 * - locale: Current content locale (e.g., "en-US", "it-IT")
 * - uiLocale: Current UI language locale
 * - timeZone: Time zone information from SPWeb (preview API)
 * - ianaTimeZone: Optional audited IANA representative derived only from timeZone.id
 * - isRtl: Right-to-left language detection from CultureInfo
 * 
 * Uses native SPFx properties (no legacy context):
 * - cultureInfo.currentCultureName
 * - cultureInfo.currentUICultureName
 * - cultureInfo.isRightToLeft
 * - web.timeZoneInfo (preview API)
 * 
 * Useful for:
 * - Internationalization (i18n) with Intl APIs
 * - Date/time formatting with timezone awareness
 * - Regional number/currency formatting
 * - RTL layout detection
 * - Calendar widget configuration
 * - Multi-lingual applications
 * 
 * The locale string can be used directly with JavaScript Intl APIs:
 * - Intl.DateTimeFormat(locale, options)
 * - Intl.NumberFormat(locale, options)
 * - Intl.Collator(locale, options)
 * 
 * @throws Error when no host-specific SPFx provider is present.
 * @returns Locale and regional settings, preserving the original SDK timeZone object
 * 
 * @example Basic locale usage
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
 *
 * function MyComponent() {
 *   const { locale, isRtl } = useSPFxLocaleInfo();
 *
 *   const formatDate = (date: Date) => {
 *     return new Intl.DateTimeFormat(locale, {
 *       dateStyle: 'full',
 *       timeStyle: 'long'
 *     }).format(date);
 *   };
 *
 *   return (
 *     <div dir={isRtl ? 'rtl' : 'ltr'}>
 *       <p>{formatDate(new Date())}</p>
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Time zone aware formatting
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
 *
 * function Calendar() {
 *   const { locale, timeZone, ianaTimeZone } = useSPFxLocaleInfo();
 *
 *   if (!timeZone || ianaTimeZone === undefined) return <div>Site time zone unavailable</div>;
 *
 *   const formatWithTimeZone = (date: Date) => {
 *     return new Intl.DateTimeFormat(locale, {
 *       timeZone: ianaTimeZone,
 *       dateStyle: 'medium',
 *       timeStyle: 'short'
 *     }).format(date);
 *   };
 *
 *   return (
 *     <div>
 *       <h3>Time Zone: {timeZone.description}</h3>
 *       <p>Offset: {timeZone.offset} minutes from UTC</p>
 *       <p>{formatWithTimeZone(new Date())}</p>
 *     </div>
 *   );
 * }
 * ```
 * 
 * @example Multi-locale support
 * ```tsx
 * import * as React from 'react';
 * import { useSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
 *
 * function PriceDisplay({ amount }: { amount: number }) {
 *   const { locale } = useSPFxLocaleInfo();
 *
 *   const price = new Intl.NumberFormat(locale, {
 *     style: 'currency',
 *     currency: 'USD'
 *   }).format(amount);
 *
 *   return <p>Price: {price}</p>;
 * }
 * ```
 *
 * @useWhen formatting or layout must follow the host content/UI culture and audited regional metadata.
 *
 * @scenario Compare host cultures and seasonal site times
 * ```tsx
 * import * as React from 'react';
 * import { Field, Select } from '@fluentui/react-components';
 * import { RelativeTime, useSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
 *
 * type DemoRegions = { controls?: React.ReactNode; commands?: React.ReactNode; status?: React.ReactNode; diagnostics?: React.ReactNode; children?: React.ReactNode };
 * export function Demo({ render, resetRevision = 0 }: { render?: (regions: DemoRegions) => React.ReactElement; resetRevision?: number } = {}): React.ReactElement {
 *   const locale = useSPFxLocaleInfo();
 *   const [season, setSeason] = React.useState<'winter' | 'summer'>('winter');
 *   const previous = React.useRef(resetRevision);
 *   React.useEffect(() => { if (previous.current !== resetRevision) { previous.current = resetRevision; setSeason('winter'); } }, [resetRevision]);
 *   const dateTime = season === 'winter' ? '2026-01-15T12:00:00Z' : '2026-07-15T12:00:00Z';
 *   const referenceTime = season === 'winter' ? '2026-01-15T13:00:00Z' : '2026-07-15T13:00:00Z';
 *   const regions: DemoRegions = {
 *     controls: <Field label="Period"><Select aria-label="Period" value={season} onChange={(_, data) => setSeason(data.value === 'summer' ? 'summer' : 'winter')}><option value="winter">Winter</option><option value="summer">Summer</option></Select></Field>,
 *     children: <section aria-label="Locale result">
 *       <p>Formatting culture: {locale.locale}; UI culture: {locale.uiLocale}; RTL: {String(locale.isRtl)}</p>
 *       <p>Regional time zone: {locale.timeZone?.description ?? 'Unavailable'}; IANA representative: {locale.ianaTimeZone ?? 'Unavailable'}</p>
 *       {locale.ianaTimeZone === undefined ? <p>IANA mapping unavailable. Site time is not rendered.</p>
 *         : <RelativeTime dateTime={dateTime} referenceTime={referenceTime} locale={locale.uiLocale} timeZone={locale.ianaTimeZone} />}
 *     </section>,
 *     diagnostics: <p>Hover or focus the actual time to inspect its full date. The audited numeric SharePoint group selects a representative standard zone, not a user city or profile zone. Culture and offsets never choose the mapping; custom TIMEZONE.XML rules are not verified. Reset restores winter without changing native locale or zone.</p>,
 *   };
 *   return render ? render(regions) : <>{regions.controls}{regions.commands}{regions.status}{regions.children}{regions.diagnostics}</>;
 * }
 * export const demo = { host: 'webPart', kind: 'snapshot', contentPlacement: 'configuration', observe: 'Switch winter and summer to compare one actual RelativeTime using the host UI culture and audited site zone; missing mapping renders an explicit unavailable result.' };
 * ```
 */
export function useSPFxLocaleInfo(): SPFxLocaleInfo {
  const pageContext = useSPFxPageContext();
  const cultureInfo = pageContext.cultureInfo;
  const locale = cultureInfo.currentCultureName;
  const uiLocale = cultureInfo.currentUICultureName;
  const isRtl = cultureInfo.isRightToLeft;
  const timeZone = (pageContext.web as { timeZoneInfo?: SPFxTimeZone }).timeZoneInfo;
  const timeZoneId = timeZone?.id;
  
  return useMemo(
    () => getSPFxLocaleInfo(pageContext),
    [pageContext, locale, uiLocale, timeZone, timeZoneId, isRtl]
  );
}
