import * as React from 'react';
import { ContentCard as rootContentCard, PreviewCard as rootPreviewCard } from '@apvee/spfx-react-toolkit';
import { ContentCard as domainContentCard, PreviewCard as domainPreviewCard } from '@apvee/spfx-react-toolkit/controls';
import { ContentCard as controlContentCard } from '@apvee/spfx-react-toolkit/controls/content-card';
import { PreviewCard as controlPreviewCard } from '@apvee/spfx-react-toolkit/controls/preview-card';
import type { ContentCardProps as RootProps, PreviewCardProps as RootPreviewProps } from '@apvee/spfx-react-toolkit';
import type { ContentCardProps as DomainProps, PreviewCardProps as DomainPreviewProps } from '@apvee/spfx-react-toolkit/controls';
import type { ContentCardProps as ControlProps, ContentCardAppearance, ContentCardHeadingLevel } from '@apvee/spfx-react-toolkit/controls/content-card';
import type { PreviewCardProps as ControlPreviewProps, PreviewCardAppearance, PreviewCardHeadingLevel,
  PreviewCardImage, PreviewCardOrientation } from '@apvee/spfx-react-toolkit/controls/preview-card';

export const aliases: typeof rootContentCard[] = [domainContentCard, controlContentCard];
export const previewAliases: typeof rootPreviewCard[] = [domainPreviewCard, controlPreviewCard];
const appearance: ContentCardAppearance & PreviewCardAppearance = 'outline';
const headingLevel: ContentCardHeadingLevel & PreviewCardHeadingLevel = 2;
const orientation: PreviewCardOrientation = 'horizontal';
const image: PreviewCardImage = { src: '/cover.jpg', alt: 'Report cover' };
export const props: RootProps = { title: { ref: React.createRef<HTMLSpanElement>(), children: 'Literal title' },
  headingLevel, appearance, actions: { children: (Tag, slotProps) => <Tag {...slotProps}><button>Open</button></Tag> },
  selected: false, focusMode: 'tab-exit', checkbox: { 'aria-label': 'Select content' } };
export const domainProps: DomainProps = props;
export const controlProps: ControlProps = props;
export const previewProps: RootPreviewProps = { ...props, orientation, image,
  preview: { ref: React.createRef<HTMLDivElement>(), children: (Tag, slotProps) => <Tag {...slotProps}>Cover</Tag> },
  logo: { as: 'img', ref: React.createRef<HTMLImageElement>(), src: '/logo.svg', alt: 'Brand' } };
export const domainPreviewProps: DomainPreviewProps = previewProps;
export const controlPreviewProps: ControlPreviewProps = previewProps;
const ContentCard = controlContentCard, PreviewCard = controlPreviewCard;
export const cards = <><ContentCard {...props} ref={React.createRef<HTMLDivElement>()} /><PreviewCard {...previewProps} /></>;
// @ts-expect-error ContentCard has fixed vertical composition.
<ContentCard title="Title" orientation="horizontal" />;
// @ts-expect-error Generated images always require alt.
<PreviewCard title="Title" image={{ src: '/cover.jpg' }} />;
// @ts-expect-error Slot tags cannot replace the owned heading.
<ContentCard title={{ as: 'h2', children: 'Title' }} />;
// @ts-expect-error Shared frame has no public package path.
import {} from '@apvee/spfx-react-toolkit/controls/cards/card-frame.internal';
// @ts-expect-error Shared resolver has no public package path.
import {} from '@apvee/spfx-react-toolkit/controls/content-slot.internal';
// @ts-expect-error Implementation paths remain private.
import {} from '@apvee/spfx-react-toolkit/controls/content-card/content-card';
