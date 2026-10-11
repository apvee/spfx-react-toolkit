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

Resolves `app.host.name` from asynchronously acquired SPFx application SDK context: Teams/TeamsModern maps to Teams, Office to Office, and Outlook/OutlookWin32 to Outlook. SDK presence alone does not classify Office or Outlook as Teams. Until context is available, unknown hosts use SharePoint URL/on-premises fallback; local workbench retains precedence. The hook updates when acquisition completes and shares the matching provider runtime through `useSPFxTeams`. The pure `getSPFxEnvironmentInfo` helper remains a separate page-context mapping, not an asynchronous SDK hook. Host detection is not a development/test/production classification or proof that a remote feature is authorized.

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

The hook reads the validated SPFx SDK wrapper (`sdks.microsoftTeams.teamsJs`) or direct SDK form. It prefers v2 `app.getContext()` and `app.registerOnThemeChangeHandler`, with v1 callback/registration fallback, and normalizes default/dark/highContrast themes. Despite its historical name, the application SDK can serve Teams, Office and Outlook. `supported` means context acquisition succeeded; use `useSPFxEnvironmentInfo` to identify the actual host. Context is `unknown`; narrow it before reading host-specific fields.

Live theme subscriptions share one lazy dispatcher per SDK identity within this library bundle. Each provider runtime store acquires context once for its active SDK subscription and reference-counts its hooks. Replacing the wrapper with the same SDK does not reacquire; replacing the resolved SDK resets state. Recent theme events take precedence over stale initialization themes. Cleanup removes runtime listeners/references and ignores obsolete completions; remount reacquires context.

The SDK offers no public unregister API. The dispatcher can remain without listeners after cleanup, and another bundle or consumer can replace the SDK's single handler. Registration/acquisition failures are reported through `console.error`; successful initial context does not guarantee live updates after a handler failure or replacement. There is no global registry or monkeypatch. See [WebPartShell](../controls/web-part-shell.md#surface-theme-and-lifecycle) for provider-composed usage.

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
