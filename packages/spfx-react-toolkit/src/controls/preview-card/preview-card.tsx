/**
 * Native Fluent cards with generated images or caller-owned previews.
 *
 * @packageDocumentation
 * @docPage controls/preview-card
 */
import * as React from 'react';
import { CardFrame } from '../cards/card-frame.internal';
import { CardImage } from '../cards/card-image.internal';
import { hasContentSlot, renderContentSlot } from '../content-slot.internal';
import { useSx } from '../../hooks/useSx';
import { useCardStyles } from '../cards/card.styles.internal';
import type { PreviewCardProps } from './preview-card.types';

/**
 * Display a generated cover or custom preview with explicit actions and native Card behavior.
 * Defined preview overrides image; object slots request the region without probing callbacks.
 * Image errors preserve informative/decorative semantics and reset only for a changed src.
 * Uses the host Fluent theme/renderer without an SPFx provider. The forwarded ref targets
 * the native Card div; optional region/interaction behavior matches ContentCard.
 * @param props - ContentCard attributes/slots, orientation and optional preview/image/logo content.
 * @returns A native Card div with summary content and the requested preview region.
 * @example Basic preview card
 * ```tsx
 * import * as React from 'react';
 * import { PreviewCard } from '@apvee/spfx-react-toolkit/controls/preview-card';
 *
 * <PreviewCard title="Report" image={{ src: '/report.jpg', alt: 'Report cover' }} actions={<button>Open report</button>} />
 * ```
 * @useWhen A card needs generated cover media or custom preview content with explicit actions.
 * @example Report card with loading and notes
 * ```tsx
 * import * as React from 'react';
 * import { Button } from '@fluentui/react-components';
 * import { PreviewCard, AutoSkeleton } from '@apvee/spfx-react-toolkit';
 *
 * export function ReportCard({ thumbnail, brandLogo, isLoading, shareReport }: {
 *   thumbnail: string; brandLogo: string; isLoading: boolean; shareReport: () => void;
 * }): React.ReactElement {
 *   return <AutoSkeleton loading={isLoading}>
 *     <PreviewCard
 *       title="Monthly report" headingLevel={2} orientation="horizontal"
 *       description="October results."
 *       image={{ src: thumbnail, alt: 'Report cover with a revenue chart' }}
 *       logo={{ as: 'img', src: brandLogo, alt: 'Brand' }}
 *       headerActions={<Button appearance="subtle" onClick={shareReport}>Share</Button>}
 *       footer="Updated today"
 *       actions={<Button as="a" href="/report" appearance="primary">Open report</Button>}
 *     >
 *       <input aria-label="Notes" defaultValue="My notes" />
 *     </PreviewCard>
 *   </AutoSkeleton>;
 * }
 * ```
 */
export const PreviewCard = React.forwardRef<HTMLDivElement, PreviewCardProps>(({ image, preview, previewFallback, ...props }, ref) => {
  const sx = useSx(); const styles = useCardStyles();
  const requested = preview !== undefined ? hasContentSlot(preview) : image !== undefined;
  const visual = preview !== undefined ? requested ? renderContentSlot(preview, 'div', sx, { shorthand: 'wrap', className: styles.media }) : undefined :
    image !== undefined ? <CardImage key={image.src} image={image} fallback={previewFallback} /> : undefined;
  return <CardFrame {...props} ref={ref} previewRequested={requested} visual={visual} />;
});
PreviewCard.displayName = 'PreviewCard';
