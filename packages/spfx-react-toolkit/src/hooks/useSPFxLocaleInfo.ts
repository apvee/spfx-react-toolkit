// useSPFxLocaleInfo.ts
// Hook for locale and regional settings

import { useMemo } from 'react';
import { getSPFxLocaleInfo } from '../helpers/spfx-page-context.helpers';
import { useSPFxPageContext } from './useSPFxPageContext';

/**
 * SPTimeZone interface (preview API in SPFx)
 * Based on Microsoft Learn documentation
 * @see https://learn.microsoft.com/en-us/javascript/api/sp-page-context/sptimezone
 * 
 * Note: This is a preview API that may change in future SPFx versions.
 * Type definition based on official docs since it's not yet exported in @microsoft/sp-page-context.
 */
export interface SPFxTimeZone {
  /** Time zone ID (numeric identifier) */
  readonly id: number;
  
  /** Offset in minutes from UTC */
  readonly offset: number;
  
  /** Time zone description (e.g., "Pacific Standard Time") */
  readonly description: string;
  
  /** Daylight savings time offset in minutes from UTC */
  readonly daylightOffset: number;
  
  /** Standard time offset in minutes from UTC */
  readonly standardOffset: number;
}

/**
 * Return type for useSPFxLocaleInfo hook
 */
export interface SPFxLocaleInfo {
  /** Current locale (e.g., "en-US", "it-IT") */
  readonly locale: string;
  
  /** Current UI locale (may differ from content locale) */
  readonly uiLocale: string;
  
  /** Time zone information from SPWeb (preview API) */
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
 * @returns Locale and regional settings, preserving the original SDK timeZone object
 * 
 * @example Basic locale usage
 * ```tsx
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
