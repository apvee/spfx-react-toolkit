import * as React from 'react';
import type { RelativeTimeProps } from '../../packages/spfx-react-toolkit/src/controls/relative-time/relative-time.types';
const props: RelativeTimeProps = {
  dateTime: new Date(), referenceTime: '2026-01-01T00:00:00Z', locale: ['it-IT', 'en'],
  format: 'short', timeZone: 'Europe/Rome', autoUpdate: false, showTooltip: false,
  fallback: <span>Unknown</span>, id: 'time', lang: 'it', tabIndex: -1,
  className: 'consumer', style: { color: 'red' }, onFocus: event => event.currentTarget.focus(),
};
void props;
// @ts-expect-error Date is an explicit required input, even if undefined.
const absent: RelativeTimeProps = {};
// @ts-expect-error Unsupported format.
const format: RelativeTimeProps = { dateTime: 0, format: 'compact' };
// @ts-expect-error No alternate HTML root.
const as: RelativeTimeProps = { dateTime: 0, as: 'span' };
// @ts-expect-error No separate display mode.
const inline: RelativeTimeProps = { dateTime: 0, inline: true };
// @ts-expect-error Typography comes from composition.
const variant: RelativeTimeProps = { dateTime: 0, variant: 'body1' };
// @ts-expect-error Content belongs to the formatter.
const children: RelativeTimeProps = { dateTime: 0, children: 'Custom' };
// @ts-expect-error No compatibility alias for the new input.
const oldDate: RelativeTimeProps = { date: 0 };
// @ts-expect-error No authored HTML injection.
const html: RelativeTimeProps = { dateTime: 0, dangerouslySetInnerHTML: { __html: '<b>unsafe</b>' } };

import { RelativeTime } from '../../packages/spfx-react-toolkit/src/controls/relative-time/relative-time';
const element = <RelativeTime dateTime={undefined} ref={React.createRef<HTMLTimeElement>()} />;
void element;
// @ts-expect-error New APIs represent absence with undefined, not null.
const legacyNull: RelativeTimeProps = { dateTime: null };
