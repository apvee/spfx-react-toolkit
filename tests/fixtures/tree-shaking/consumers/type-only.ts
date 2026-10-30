import type { SPFxPnPListSelector } from '__TOOLKIT__';
export const identity = (value: SPFxPnPListSelector): SPFxPnPListSelector => value;

import type { ContentCardProps, PreviewCardProps, PreviewCardImage } from '__TOOLKIT__';
export type Cards = ContentCardProps | PreviewCardProps | PreviewCardImage;
