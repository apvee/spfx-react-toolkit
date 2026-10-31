# PreviewCard

[Documentation](../../INDEX.md) > [API reference](../INDEX.md#react-controls) > PreviewCard

`PreviewCard` adds a generated image or custom preview to [ContentCard](content-card.md), using native Fluent `CardPreview`. Import from the package root, `/controls` or `/controls/preview-card`.

```tsx
import { Button } from '@fluentui/react-components';
import { PreviewCard, AutoSkeleton } from '@apvee/spfx-react-toolkit';

<AutoSkeleton loading={isLoading}>
  <PreviewCard
    title="Monthly report"
    headingLevel={2}
    orientation="horizontal"
    description="October results."
    image={{ src: thumbnail, alt: 'Report cover with a revenue chart' }}
    logo={{ as: 'img', src: brandLogo, alt: 'Brand' }}
    headerActions={<Button appearance="subtle">Share</Button>}
    footer="Updated today"
    actions={<Button as="a" href="/report" appearance="primary">Open report</Button>}
  >
    <input aria-label="Notes" defaultValue="My notes" />
  </PreviewCard>
</AutoSkeleton>
```

`PreviewCardProps` inherits `ContentCardProps`: `title`, `headingLevel`, `description`, `icon`, `headerActions`, `children`, `footer`, `actions`, `appearance` and native Card interactions apply. `headingLevel` defaults to `3`, `size` to `medium`, `appearance` to `filled`. Strings remain literal; JSX/callback presentation is caller-owned.

| Added prop | Contract |
| --- | --- |
| `PreviewCardProps.orientation` | `vertical` (default) or `horizontal`. |
| `PreviewCardProps.image` | Optional `PreviewCardImage`: `{ src: string; alt: string }`; alt is required, with `''` marking decorative content. |
| `PreviewCardProps.preview` | Optional fixed inner `div` slot; every defined value takes precedence over `image`. |
| `PreviewCardProps.previewFallback` | Optional fixed inner `div` slot after generated image failure; undefined uses meaningful alt text. |
| `PreviewCardProps.logo` | Optional bounded inner `div`/`img` slot in native 32px overlay, independent of the header icon. |

`PreviewCardImage.src` is the resource URL. `PreviewCardImage.alt` is its required accessible description; an empty string is decorative.

Public types: `PreviewCardProps`, `PreviewCardAppearance`, `PreviewCardHeadingLevel`, `PreviewCardImage`, `PreviewCardOrientation`.

`image` and `preview` can be supplied together. Only an undefined `preview` derives an image. A defined false, legacy null, empty string or empty fragment suppresses the entire region and logo. An explicit props object requests the preview region even with empty children. A render callback object also requests the region when its callback returns undefined; the toolkit does not call it ahead of rendering or fall back to the image. Suppressed preview, fallback and logo callbacks are never probed. Logo requires a requested preview region; its inner ref identifies the selected div/img, while native CardPreview owns the overlay bounds. Introduced containers fit the overlay; shorthand JSX retains caller sizing.

Generated images crop with cover fit. After failure, the default and custom fallback fill the same reserved media area as the image: width, height, card layout and logo placement stay stable in vertical, horizontal and wrapped layouts. Defined empty fallback content leaves that area reserved. Image failure is local to `src`: a new source starts cleanly, while changing alt, fallback, layout, theme or width cannot automatically retry the same failed source. Defined fallback overrides alt content, including explicit empty suppression. Informative failed media keeps `role="img"` and its alt accessible name, even for an empty custom fallback. Decorative generated/failed media is hidden from assistive technology. Custom fallback JSX controls live within the media's accessibility shell; put recovery controls in `actions` when they must be independently accessible.

Media keeps native Card padding and border-radius variables. Vertical and wrapped previews use 16:9 geometry. Horizontal media prefers an 8rem column stretched to details, with a preferred 12rem details minimum; flex wrapping moves the preview above details when the container narrows. This depends on container width, without viewport breakpoints or measured layout state. The same detail wrapper preserves input/control identity across orientation changes. Custom previews stretch into the region and own their internal presentation.

Dependencies, fixed-tag objects/callbacks, refs, raw-HTML guards and native Card interaction limits match [ContentCard](content-card.md) and the [slot convention](../../guides/control-slots.md). No new peers, provider, renderer, implicit loading, route, playback or recovery workflow is introduced.

The lazy CardsPanel follows the realistic report mockup: one PreviewCard by default, with selectable ContentCard, title/description, orientation and standard/compact widths, loading/disabled and explicit image failure/recovery commands. Shared notes survive ordinary properties and type changes. Its outline configuration holds compact status and collapsed diagnostics; those options affect the same card and never introduce native comparisons. The live card renders directly afterward without a padded demo surface. Full slot/empty-media/native/isolation regressions remain in [versioned conformance fixtures](../../../tests/fixtures/demo-control-isolation/cards-regressions.tsx). Browser and authenticated host execution must be reported separately.
