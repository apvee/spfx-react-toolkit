import * as React from 'react';
import { Badge } from '@fluentui/react-components';
import { RelativeTime } from '../../../packages/spfx-react-toolkit/src/controls/relative-time/relative-time';
import { FormattedText } from '../../../packages/spfx-react-toolkit/src/controls/formatted-text/formatted-text';

const event = '2026-01-15T12:00:00Z';
const reference = '2026-01-15T13:00:00Z';

/** Deterministic checks mounted under the existing fixture FluentProvider. */
export function RelativeTimeFixture(): React.ReactElement {
  const [locale, setLocale] = React.useState('en-US');
  return <section id="rt-fixture" aria-label="RelativeTime browser checks">
    <label>Relative locale{' '}<select id="rt-locale" value={locale} onChange={change => setLocale(change.target.value)}>
      <option>en-US</option><option>it-IT</option><option>ar-SA</option>
    </select></label>
    <p id="rt-description">Consumer full-date description</p>
    <button id="rt-tab-before" type="button">Before relative times</button>{' '}
    <RelativeTime id="rt-tooltip" dateTime={event} referenceTime={reference} locale={locale} timeZone="UTC" aria-describedby="rt-description" />{' '}
    <RelativeTime id="rt-no-tooltip" dateTime={event} referenceTime={reference} locale={locale} timeZone="UTC" showTooltip={false} />{' '}
    <button id="rt-tab-after" type="button">After relative times</button>
    <div><FormattedText id="rt-parent" variant="caption1" subtle><RelativeTime id="rt-inherited" dateTime={event} referenceTime={reference} locale={locale} timeZone="UTC" showTooltip={false} /></FormattedText></div>
    <div><Badge appearance="outline"><RelativeTime id="rt-badge" dateTime={event} referenceTime={reference} locale={locale} timeZone="UTC" showTooltip={false} /></Badge></div>
    <div><RelativeTime id="rt-fixed" data-relative-time-fixture="fixed" dateTime="2026-01-31T12:00:00Z" referenceTime="2026-02-28T12:00:00Z" locale="en-US" timeZone="UTC" showTooltip={false} /></div>
    <div id="rt-fallback"><RelativeTime dateTime="2026-02-30T12:00:00Z" fallback="Date unavailable" /></div>
    <div><RelativeTime id="rt-utc" dateTime={event} referenceTime={reference} locale="en-US" timeZone="UTC" showTooltip={false} /></div>
    <div><RelativeTime id="rt-rome" dateTime={event} referenceTime={reference} locale="en-US" timeZone="Europe/Rome" showTooltip={false} /></div>
  </section>;
}
