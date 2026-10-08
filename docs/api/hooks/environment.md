# Environment Hooks

Hooks for host environment, Teams context, locale and page type. These hooks must run under the matching SPFx provider.

## useSPFxEnvironmentInfo

```typescript
function useSPFxEnvironmentInfo(): SPFxEnvironmentInfo;
type SPFxEnvironmentType = 'Local' | 'SharePoint' | 'SharePointOnPrem' | 'Teams' | 'Office' | 'Outlook';
interface SPFxEnvironmentInfo {
  readonly type: SPFxEnvironmentType;
  readonly isLocal: boolean;
  readonly isWorkbench: boolean;
  readonly isSharePoint: boolean;
  readonly isSharePointOnPrem: boolean;
  readonly isTeams: boolean;
  readonly isOffice: boolean;
  readonly isOutlook: boolean;
}
```

Maps page-context metadata, host URL and available SDK objects through `getSPFxEnvironmentInfo`. This is host detection, not a development/test/production classification or proof that a remote feature is authorized.

```tsx
import * as React from 'react';
import { useSPFxEnvironmentInfo } from '@apvee/spfx-react-toolkit';

function EnvironmentBanner() {
  const { type, isLocal, isWorkbench } = useSPFxEnvironmentInfo();
  return <p>{type}{isLocal || isWorkbench ? ' — workbench' : ''}</p>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxEnvironmentInfo.ts)

## useSPFxTeams

```typescript
function useSPFxTeams(): SPFxTeamsInfo;
type TeamsTheme = 'default' | 'dark' | 'highContrast';
interface SPFxTeamsInfo {
  readonly supported: boolean;
  readonly context: unknown | undefined;
  readonly theme: TeamsTheme | undefined;
}
```

The hook reads the SPFx Teams SDK wrapper (`sdks.microsoftTeams.teamsJs`) and direct SDK forms, trying the v2 `app.getContext()` API before the v1 callback API. It normalizes the acquired theme. The returned context is `unknown` because SDK versions have different context shapes; narrow it in consuming code. SPFx context replacement resets Teams state and ignores obsolete initialization callbacks; unmount cleanup ignores later completions. The hook does not register a live Teams theme-change subscription.

```tsx
import * as React from 'react';
import { useSPFxTeams } from '@apvee/spfx-react-toolkit';

function TeamsStatus() {
  const { supported, theme } = useSPFxTeams();
  return <p>{supported ? `Teams theme: ${theme}` : 'Teams context unavailable'}</p>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxTeams.ts)

## useSPFxLocaleInfo

```typescript
function useSPFxLocaleInfo(): SPFxLocaleInfo;
interface SPFxTimeZone {
  readonly id: number;
  readonly offset: number;
  readonly description: string;
  readonly daylightOffset: number;
  readonly standardOffset: number;
}
interface SPFxLocaleInfo {
  readonly locale: string;
  readonly uiLocale: string;
  readonly timeZone: SPFxTimeZone | undefined;
  readonly ianaTimeZone?: string;
  readonly isRtl: boolean;
}
```

### Locale result members

| Member | Meaning |
| --- | --- |
| `SPFxLocaleInfo.locale` | Content culture |
| `SPFxLocaleInfo.uiLocale` | User interface language |
| `SPFxLocaleInfo.timeZone` | Original optional SharePoint preview metadata |
| `SPFxLocaleInfo.ianaTimeZone` | Audited optional standard-group IANA representative |
| `SPFxLocaleInfo.isRtl` | Culture writing direction |
| `SPFxTimeZone.id` | SharePoint numeric time-zone identifier |
| `SPFxTimeZone.offset` | SharePoint UTC bias in minutes |
| `SPFxTimeZone.description` | SharePoint description, not a mapping key |
| `SPFxTimeZone.daylightOffset` | Additional daylight-time bias in minutes |
| `SPFxTimeZone.standardOffset` | Additional standard-time bias in minutes |

Maps SPFx culture information and optional `web.timeZoneInfo` preview metadata. `SPFxLocaleInfo.ianaTimeZone` is an optional representative IANA identifier derived only from `timeZone.id`. The original `timeZone` object is preserved by reference, without cloning or mutation. A changed ID on the same metadata object is observed on the next render.

The static mapping uses verified SharePoint standard groups and CLDR territory `001`, with no network requests, locale/offset/description inference or runtime Windows conversion. Unknown, absent, invalid or excluded IDs return undefined. It does not locate a user's city or guarantee all historical/custom rules. SharePoint Server `TIMEZONE.XML` can be customized; this standard mapping cannot verify those installations. Coverage and immutable provenance are documented in the [helper guide](../helpers/INDEX.md#getspfxlocaleinfo).

Passing undefined as Intl/RelativeTime timeZone uses the runtime zone, not a resolved site zone. Locale `it-IT` does not imply Rome; group 4 maps to `Europe/Berlin`.

```tsx
import * as React from 'react';
import { RelativeTime, useSPFxLocaleInfo } from '@apvee/spfx-react-toolkit';

function LocalizedDate({ date }: { date: Date }) {
  const { locale, isRtl, ianaTimeZone } = useSPFxLocaleInfo();
  if (!ianaTimeZone) return <span>Site time zone unavailable</span>;
  return <time dir={isRtl ? 'rtl' : 'ltr'} dateTime={date.toISOString()}>
    {new Intl.DateTimeFormat(locale, { dateStyle: 'full', timeZone: ianaTimeZone }).format(date)}
  </time>;
}
```

```tsx
const { uiLocale, ianaTimeZone } = useSPFxLocaleInfo();
// Explicit runtime-zone fallback when mapping is unavailable:
<RelativeTime dateTime={modified} locale={uiLocale} timeZone={ianaTimeZone} />
```

See [RelativeTime](../controls/relative-time.md) for clock and accessibility behavior.

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxLocaleInfo.ts)

## useSPFxPageType

```typescript
function useSPFxPageType(): SPFxPageTypeInfo;
type SPFxPageType = 'sitePage' | 'webPartPage' | 'listPage' | 'listFormPage' | 'profilePage' | 'searchPage' | 'unknown';
interface SPFxPageTypeInfo {
  readonly pageType: SPFxPageType;
  readonly isModernPage: boolean;
  readonly isSitePage: boolean;
  readonly isListPage: boolean;
  readonly isListFormPage: boolean;
  readonly isWebPartPage: boolean;
}
```

Maps modern/legacy page metadata; unknown or unavailable metadata can produce `unknown`.

```tsx
import * as React from 'react';
import { useSPFxPageType } from '@apvee/spfx-react-toolkit';

function PageStatus() {
  const { pageType, isModernPage } = useSPFxPageType();
  return <p>{pageType}: {isModernPage ? 'modern' : 'other page'}</p>;
}
```

[View source](../../../packages/spfx-react-toolkit/src/hooks/useSPFxPageType.ts)

## See Also

- [Context Hooks](./context.md)
- [Theming Hooks](./theming.md)
- [Page Context Helpers](../helpers/INDEX.md#page-context-helpers)
