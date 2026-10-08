import type * as React from 'react';

/** Localized relative timestamp with a native time root and optional full-date description.
 * Styling is inherited; compose FormattedText or Badge for presentation.
 * @example <RelativeTime dateTime="2026-01-01T12:00:00Z" locale="it-IT" format="short" />
 */
export interface RelativeTimeProps extends Omit<React.ComponentPropsWithoutRef<'time'>, 'children' | 'dateTime' | 'dangerouslySetInnerHTML'> {
  /** Valid Date, epoch milliseconds or ISO instant with Z/offset. Undefined or invalid input renders fallback. */
  dateTime: Date | string | number | undefined;
  /** BCP47 locale or ordered preferences; defaults to the runtime locale. */
  locale?: string | string[];
  /** Localized unit spelling, default long. Abbreviations depend on Intl locale data. */
  format?: 'long' | 'short' | 'narrow';
  /** Refresh at label boundaries and visibility resume, default true; ignored with referenceTime. */
  autoUpdate?: boolean;
  /** Fluent full-date tooltip and default keyboard tab stop, default true. */
  showTooltip?: boolean;
  /** Intl time zone for the full date only, default runtime zone. Invalid configuration throws for valid dates. */
  timeZone?: string;
  /** Consumer-localized content for invalid date/reference, default undefined; no time root or timer. */
  fallback?: React.ReactNode;
  /** Fixed reference instant; disables automatic updates. Invalid supplied references render fallback. */
  referenceTime?: Date | string | number;
  /** Not supported: the formatter owns content. */
  children?: never;
  /** Not supported: generated content is always React text. */
  dangerouslySetInnerHTML?: never;
}
