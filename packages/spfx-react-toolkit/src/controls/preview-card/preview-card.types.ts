import type { ContentCardProps, ContentCardAppearance, ContentCardHeadingLevel } from '../content-card/content-card.types';
import type { AlternativeContentSlot, ContentSlot } from '../content-slot.types.internal';

/** Native Fluent Card appearance. Defaults to filled.
 * @example
 * ```ts
 * const appearance: PreviewCardAppearance = 'filled';
 * ```
 */
export type PreviewCardAppearance = ContentCardAppearance;
/** Owned outer heading level; defaults to 3.
 * @example
 * ```ts
 * const headingLevel: PreviewCardHeadingLevel = 2;
 * ```
 */
export type PreviewCardHeadingLevel = ContentCardHeadingLevel;
/** Preview placement; horizontal wraps based on available card width.
 * @example
 * ```ts
 * const orientation: PreviewCardOrientation = 'horizontal';
 * ```
 */
export type PreviewCardOrientation = 'vertical' | 'horizontal';
/** Generated cover image; resource failure state resets only when src changes.
 * @example
 * ```ts
 * const image: PreviewCardImage = { src: '/cover.jpg', alt: 'Report cover' };
 * ```
 */
export interface PreviewCardImage {
  /** Image resource URL; a new value starts a fresh load. */
  src: string;
  /** Accessible description; empty string identifies a decorative image. */
  alt: string;
}
/** Native Card with a generated image or caller-owned preview and bounded logo overlay.
 * Inherits ContentCard's literal text, semantic heading, content slots and native interaction.
 * Any defined preview overrides image, including legacy null and conditional booleans.
 * @example
 * ```tsx
 * <PreviewCard title="Report" image={{ src: '/report.jpg', alt: 'Report cover' }} actions={<button>Open</button>} />
 * ```
 */
export interface PreviewCardProps extends ContentCardProps {
  /** Preview above details by default; horizontal wraps above at narrow container widths. */
  orientation?: PreviewCardOrientation;
  /** Optional generated cover image with required alt; ignored when preview is defined. */
  image?: PreviewCardImage;
  /** Fixed div slot overriding image. Empty shorthand suppresses the region; props/callback objects request it even when empty. */
  preview?: ContentSlot<'div'>;
  /** Div slot after generated image failure, filling the same reserved media area.
   * Undefined uses meaningful alt text; defined empty values suppress fallback content without collapsing the preview. */
  previewFallback?: ContentSlot<'div'>;
  /** Native 32px preview overlay with bounded inner div/img slot and inner ref; rendered only for a requested preview region. */
  logo?: AlternativeContentSlot<'div', 'img'>;
}
