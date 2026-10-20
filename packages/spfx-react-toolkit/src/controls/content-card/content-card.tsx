import * as React from 'react';
import { CardFrame } from '../cards/card-frame.internal';
import type { ContentCardProps } from './content-card.types';

/** Group a content summary with explicit caller-owned actions and native Card interaction.
 * Literal strings retain their text; compose FormattedText explicitly for Markdown.
 * Requires the host Fluent theme and Griffel renderer; no SPFx provider is required.
 * Root ref identifies the native div. Disabled preserves child disabled flags; native Card may block pointer input with its overlay.
 * @param props Native Fluent Card behavior, required title and optional summary/body/action slots.
 * @returns A native div Card with a heading, optional body and explicit footer regions.
 * @example
 * ```tsx
 * <ContentCard title="Workspace" footer="Updated today" actions={<button>Open</button>}>Three members</ContentCard>
 * ```
 */
export const ContentCard = React.forwardRef<HTMLDivElement, ContentCardProps>((props, ref) => <CardFrame {...props} ref={ref} orientation="vertical" />);
ContentCard.displayName = 'ContentCard';
