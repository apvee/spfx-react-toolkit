# RelativeTime

[Documentation](../../INDEX.md) > [API reference](../INDEX.md#react-controls) > RelativeTime

`RelativeTime` formats an explicit instant as localized elapsed time and exposes its full date through an accessible description. It renders a native `time` element and requires no SPFx provider. Its typography, color and direction inherit from the surrounding UI. Tooltip uses the existing Fluent UI 9 peer dependency and shared Fluent/Griffel host runtime; no extra date library is required.

```tsx
import { RelativeTime, FormattedText } from '@apvee/spfx-react-toolkit';

<FormattedText variant="caption1" subtle>
  <RelativeTime dateTime="2026-01-31T12:00:00Z" locale="it-IT" timeZone="Europe/Rome" />
</FormattedText>
```

Import `RelativeTime` and `RelativeTimeProps` from the package root, `/controls`, or `/controls/relative-time`. These paths share one implementation. Parsing and clock helpers are private.

## RelativeTimeProps

The props extend native `React.ComponentPropsWithoutRef<'time'>`, excluding formatter-owned content and replacing the native string dateTime type with the required instant input. The forwarded ref targets `HTMLTimeElement`. Native events, `className`, consumer `style`, ARIA attributes and explicit `tabIndex` are forwarded.

| Member | Type | Default and behavior |
| --- | --- | --- |
| `RelativeTimeProps.dateTime` | `Date \| string \| number \| undefined` | Required prop; undefined or invalid input renders fallback. Numbers are epoch milliseconds. Strings require ISO `Z` or numeric offset. |
| `RelativeTimeProps.locale` | `string \| string[]` | Runtime locale; ordered locale preferences accepted. |
| `RelativeTimeProps.format` | `'long' \| 'short' \| 'narrow'` | `'long'`; abbreviations depend on runtime Intl data. |
| `RelativeTimeProps.autoUpdate` | `boolean` | `true`; updates when the label changes. |
| `RelativeTimeProps.showTooltip` | `boolean` | `true`; full-date Fluent tooltip with default keyboard tab stop. |
| `RelativeTimeProps.timeZone` | `string` | Runtime zone; changes the full date, not elapsed time or UTC calendar anniversaries. |
| `RelativeTimeProps.fallback` | span content slot | `undefined`; consumer supplies localized invalid-input content. |
| `RelativeTimeProps.referenceTime` | `Date \| string \| number` | Undefined uses the local clock; a valid supplied instant fixes the reference and overrides autoUpdate. Invalid supplied references render fallback. |
| `RelativeTimeProps.children` | `never` | Unsupported; the formatter owns content. |
| `RelativeTimeProps.dangerouslySetInnerHTML` | `never` | Unsupported; formatted output uses React text. |

Unsupported content props are also removed from untyped runtime input. The old `date` prop is not an alias. The native time `dateTime` attribute is normalized UTC ISO derived from the input. Invalid event/reference input renders fallback without a time root, ref target, full-date description or timer. Instant validation occurs before Intl construction, so invalid input still renders fallback if locale or timeZone is invalid. With valid instants, invalid Intl configuration propagates `RangeError`.

## Instant and unit rules

Date-only and timezone-less strings, impossible calendar dates, invalid offset/hour/minute/second fields, nonfinite epochs and values outside the native Date range are invalid. Fractional ISO seconds truncate to JavaScript millisecond precision. Valid signed six-digit extended ISO years are supported within that range.

Less than one minute in either direction renders the localized equivalent of “now” using zero seconds and `numeric: 'auto'`. Other labels use `numeric: 'always'` and completed minutes, hours, 24-hour days and seven-day weeks. Months and years use UTC anniversaries anchored to the original event, clamping to the end of the target month. January 31 to February 28 is one completed month; March 31 is two months. Leap-day anniversaries similarly clamp. Months are not approximated by thirty days; days are not rendered as yesterday/tomorrow. Daylight-saving changes affect the parsed instant but do not alter absolute elapsed hours.

```tsx
<RelativeTime dateTime="2026-01-31T12:00:00Z"
  referenceTime="2026-02-28T12:00:00Z" locale="en-US" showTooltip={false} />
// One month ago, with a hidden full-date description.
<RelativeTime dateTime="invalid" fallback="Date unavailable" />
```

## Clock modes and cleanup

A supplied valid `referenceTime` is controlled: it creates no timer or visibility listener. With `autoUpdate={false}`, the local sample stays frozen until the parsed event instant or clock mode changes, then resamples after commit. Equivalent instant inputs and presentation changes do not reset that sample. This mode does not promise synchronous sampling during render.

Live mode uses one timeout per instance at the next label transition, with a minimum delay of 1ms and maximum of one hour for clock resynchronization. It pauses while the Fluent host document is hidden and resynchronizes when visible. Input, document or mode changes and unmount clean up timers/listeners. When no document exists, it uses timers without visibility listeners. At Date extrema, unrepresentable future boundaries use the hourly cap safely rather than an overflowing tight loop.

## Full date and accessibility

The Fluent tooltip uses a description relationship and stable full-date ID. The root merges its ID with consumer `aria-describedby` tokens. Tooltip-enabled valid times receive `tabIndex={0}` unless the consumer overrides it; hover/focus, Escape/blur and ref/handlers compose through Fluent. With `showTooltip={false}`, a hidden description remains and no tab stop is added. `lang` defaults to the resolved locale unless overridden.

There is no automatic `aria-live` announcement. If access to exact dates is essential on touch devices, provide an explicit consumer full-date surface. Use `FormattedText` or Fluent `Badge` for presentation; the control adds no typography/appearance API or generated inline styling.

Requires browser `Intl.RelativeTimeFormat`, `Intl.DateTimeFormat` and appropriate locale/time-zone data. No polyfill is installed automatically. Full-date spelling, punctuation and relative abbreviations vary by runtime.

## SPFx locale composition

```tsx
const { uiLocale, ianaTimeZone } = useSPFxLocaleInfo();
<RelativeTime dateTime={modified} locale={uiLocale} timeZone={ianaTimeZone} />
```

The [locale hook](../hooks/environment.md#usespfxlocaleinfo) requires an SPFx provider; RelativeTime itself does not. `ianaTimeZone` identifies a representative standard SharePoint numeric zone group. When undefined, this example uses the runtime zone: it does not establish the site's zone. A locale such as `it-IT` does not imply `Europe/Rome`; verified group 4 maps to `Europe/Berlin`. Consumers should choose their unavailable-mapping UI or runtime fallback explicitly.

The sample's lazy RelativeTime panel exercises all props, three clock modes, invalid input, composition and native ref/events. ContextPanel separately compares fixed winter/summer instants only for a mapped site zone. Local tests and browser checks do not establish authenticated SharePoint-host behavior; see [SharePoint validation](../../development/sharepoint-validation/checklists/relative-time.md#relativetime-and-site-time-zone).

## Invalid-input fallback slot

`fallback` accepts literal string/JSX shorthand or a fixed `span` props object using existing Fluent/Griffel peers. Shorthand gains no wrapper; objects supply native attributes, classes, a ref and static or Fluent callback children:

```tsx
<RelativeTime dateTime={undefined} fallback={{ ref: fallbackRef,
  children: (Component, props) => <Component {...props}>Date unavailable</Component>,
}} />
```

Callbacks run only for invalid event/reference input during rendering, without presence preflight. The fallback ref targets span; the component ref remains reserved for the valid time root. Booleans, undefined and legacy null render no fallback; empty props still request a span. Strings stay literal. `as` can only repeat span, raw HTML and unsupported attributes are discarded, and JSX/callback subtrees stay caller-owned. React 17 erased `ReactNode` types can bypass literal-object restrictions. Valid time, clock, Tooltip and localization remain independent of fallback customization. No new dependency or SPFx provider is required.
