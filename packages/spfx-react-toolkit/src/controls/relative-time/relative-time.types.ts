import type { ContentSlot } from '../content-slot.types.internal';
import type * as React from 'react';

/** Localized relative timestamp with a native time root and optional full-date description.
 * Styling is inherited; compose FormattedText or Badge for presentation. Uses native
 * `Intl` formatters and the host Fluent tooltip context; no SPFx provider is required.
 * @example
 * ```tsx
 * <RelativeTime dateTime="2026-01-01T12:00:00Z" locale="it-IT" format="short" />
 * ```
 */
export interface RelativeTimeProps extends Omit<React.ComponentPropsWithoutRef<'time'>, 'children' | 'dateTime' | 'dangerouslySetInnerHTML'> {
  /** Valid `Date`, epoch milliseconds or ISO datetime with `Z`/numeric offset.
   * Date-only, timezone-less and impossible calendar values are invalid. `undefined` or invalid input renders fallback. */
  dateTime: Date | string | number | undefined;
  /** BCP47 locale or ordered preferences; defaults to the runtime locale. */
  locale?: string | string[];
  /** Localized unit spelling, default long. Abbreviations depend on Intl locale data. */
  format?: 'long' | 'short' | 'narrow';
  /** Refresh at label boundaries and host-document visibility resume; defaults to `true`.
   * With `false`, sample after commit on instant/mode changes; presentation changes do not resample.
   * A supplied `referenceTime` disables this clock and its listeners. */
  autoUpdate?: boolean;
  /** Show a Fluent full-date description tooltip; defaults to `true` and `tabIndex=0`
   * unless explicitly overridden. With `false`, keep a hidden full-date description without adding a tab stop. */
  showTooltip?: boolean;
  /** Intl time zone for the full date only, default runtime zone. Invalid configuration throws for valid dates. */
  timeZone?: string;
  /** Literal content or span props/render slot for invalid date/reference, default undefined; no time root or timer. Shorthand gains no wrapper; slot ref targets the span, separate from the time root ref. Callbacks run only for invalid input. */
  fallback?: ContentSlot<'span'>;
  /** Fixed reference instant, using the same validation as `dateTime`; disables timers
   * and automatic updates. Omission uses the local clock; invalid supplied references render fallback. */
  referenceTime?: Date | string | number;
  /** Not supported: the formatter owns content. */
  children?: never;
  /** Not supported: generated content is always React text. */
  dangerouslySetInnerHTML?: never;
}
