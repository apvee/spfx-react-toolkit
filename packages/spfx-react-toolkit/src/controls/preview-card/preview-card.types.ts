/**
 * Native Fluent cards with generated images or caller-owned previews.
 *
 * @packageDocumentation
 * @docPage controls/preview-card
 */
import type { ContentCardProps, ContentCardAppearance, ContentCardHeadingLevel } from '../content-card/content-card.types';
import type { AlternativeContentSlot, ContentSlot } from '../content-slot.types.internal';

/**
 * Native Fluent Card appearance. Defaults to filled.
 * @example
 * ```ts
 * import { PreviewCard, PreviewCardAppearance } from '@apvee/spfx-react-toolkit/controls/preview-card';
 *
 * const appearance: PreviewCardAppearance = 'filled';
 * ```
 */
export type PreviewCardAppearance = ContentCardAppearance;
/**
 * Owned outer heading level; defaults to 3.
 * @example
 * ```ts
 * import { PreviewCard, PreviewCardHeadingLevel } from '@apvee/spfx-react-toolkit/controls/preview-card';
 *
 * const headingLevel: PreviewCardHeadingLevel = 2;
 * ```
 */
export type PreviewCardHeadingLevel = ContentCardHeadingLevel;
/**
 * Preview placement; horizontal wraps based on available card width.
 * @example
 * ```ts
 * import { PreviewCard, PreviewCardOrientation } from '@apvee/spfx-react-toolkit/controls/preview-card';
 *
 * const orientation: PreviewCardOrientation = 'horizontal';
 * ```
 */
export type PreviewCardOrientation = 'vertical' | 'horizontal';
/**
 * Generated cover image; resource failure state resets only when src changes.
 * @example
 * ```ts
 * import { PreviewCard, PreviewCardImage } from '@apvee/spfx-react-toolkit/controls/preview-card';
 *
 * const image: PreviewCardImage = { src: '/cover.jpg', alt: 'Report cover' };
 * ```
 */
export interface PreviewCardImage {
  /** Image resource URL; a new value starts a fresh load. */
  src: string;
  /** Accessible description; empty string identifies a decorative image. */
  alt: string;
}
/**
 * Native Card with a generated image or caller-owned preview and bounded logo overlay.
 * Inherits ContentCard's literal text, semantic heading, content slots and native interaction.
 * Any defined preview overrides image, including legacy null and conditional booleans.
 * @example
 * ```tsx
 * import * as React from 'react';
 * import { PreviewCard } from '@apvee/spfx-react-toolkit/controls/preview-card';
 *
 * <PreviewCard title="Report" image={{ src: '/report.jpg', alt: 'Report cover' }} actions={<button>Open</button>} />
 * ```
 * @remarks
 * ### Media behavior
 * Generated images use cover fit. Failed media keeps the reserved area and logo placement; meaningful alt text retains role=img and its accessible name even with an empty custom fallback. Decorative media is hidden from assistive technology. Custom fallback controls remain inside that accessibility shell; use actions for independently accessible recovery controls. Only changing src retries a failed image.
 * ### Layout and slots
 * Vertical and wrapped media uses 16:9 geometry. Horizontal media prefers an 8rem column with 12rem detail minimum and wraps according to container width; detail controls retain ancestry across orientation changes. Custom preview content owns its presentation. The logo inner ref identifies the bounded div/img, while native CardPreview owns the 32px overlay. Props/callback slots request containers without probing callbacks; raw HTML is unsupported. Inherited title, description, icon, headerActions, footer and actions retain ContentCard’s fixed tags, literal formatting and native interaction contracts. Native Card size defaults to medium, appearance to filled and the owned heading to h3. Requires the shared host Fluent theme/Griffel renderer, with no SPFx provider or implicit loading/navigation behavior.
 */
export interface PreviewCardProps extends ContentCardProps {
  /**
   * Preview above details by default; horizontal wraps above at narrow container widths.
   * When omitted, preview is vertical; horizontal wraps above details at narrow container widths.
   * @defaultValue vertical
   */
  orientation?: PreviewCardOrientation;
  /**
   * Optional generated cover image with required alt; ignored when preview is defined.
   * When omitted, no generated image is requested.
   */
  image?: PreviewCardImage;
  /**
   * Fixed div slot overriding image. Empty shorthand suppresses the region; props/callback objects request it even when empty.
   * When omitted, image supplies the region. Any defined empty shorthand suppresses it; props/callback objects still request it.
   */
  preview?: ContentSlot<'div'>;
  /** Div slot after generated image failure, filling the same reserved media area.
   * Undefined uses meaningful alt text; defined empty values suppress fallback content without collapsing the preview. */
  previewFallback?: ContentSlot<'div'>;
  /**
   * Native 32px preview overlay with bounded inner div/img slot and inner ref; rendered only for a requested preview region.
   * When omitted, no preview logo overlay is rendered.
   */
  logo?: AlternativeContentSlot<'div', 'img'>;
}
