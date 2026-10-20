import { makeStyles } from '@griffel/react';
import { cardCSSVars, cardPreviewClassNames, tokens } from '@fluentui/react-components';

/** Geometry native Card does not provide; styles use the host's Griffel renderer. */
export const useCardStyles = makeStyles({
  root: { flexWrap: 'wrap', overflowWrap: 'anywhere' },
  detail: { display: 'flex', flexDirection: 'column', minInlineSize: 0, gap: `var(${cardCSSVars.cardSizeVar})`, inlineSize: '100%' },
  detailHorizontal: { inlineSize: 'auto', flexBasis: `calc(100% - 8rem - var(${cardCSSVars.cardSizeVar}))`, flexGrow: 1, minInlineSize: 'min(100%, 12rem)' },
  preview: {
    minInlineSize: 0, flexShrink: 0, borderRadius: `var(${cardCSSVars.cardBorderRadiusVar})`, overflow: 'hidden',
    // Outrank native full-bleed selectors even when another Fluent bundle inserts
    // its rules later (for example, a second web part in the SharePoint host).
    '&&&&&&': { margin: 0 },
    [`> :not(.${cardPreviewClassNames.logo})`]: { height: 'auto' },
  },
  previewVertical: { inlineSize: '100%', aspectRatio: '16 / 9', display: 'grid' },
  previewHorizontal: {
    inlineSize: '8rem', maxInlineSize: '100%', flexGrow: 1, alignSelf: 'stretch', display: 'grid', aspectRatio: '16 / 9',
    [`&& > :not(.${cardPreviewClassNames.logo})`]: { display: 'grid', minBlockSize: 0 },
    [`&&& > img:not(.${cardPreviewClassNames.logo})`]: { position: 'absolute', inset: 0, height: '100%', aspectRatio: 'auto', objectFit: 'cover' },
    [`&& > div:not(.${cardPreviewClassNames.logo}) > img`]: { minBlockSize: 0, minInlineSize: 0, width: '100%', height: '100%', objectFit: 'cover' },
  },
  text: { overflowWrap: 'anywhere', margin: 0 },
  description: { marginBlockStart: tokens.spacingVerticalXXS },
  clamp: { display: '-webkit-box', WebkitBoxOrient: 'vertical', WebkitLineClamp: 2, overflow: 'hidden' },
  disabled: { color: tokens.colorNeutralForegroundDisabled, '@media (forced-colors: active)': { color: 'GrayText' } },
  icon: { display: 'inline-flex', flexShrink: 0 },
  footer: { alignItems: 'center', flexWrap: 'wrap', minInlineSize: 0 },
  actions: { display: 'flex', alignItems: 'center', flexWrap: 'wrap', justifyContent: 'flex-end', gap: tokens.spacingHorizontalS, minInlineSize: 0, maxInlineSize: '100%' },
  media: { boxSizing: 'border-box', minInlineSize: 0, overflow: 'hidden', inlineSize: '100%' },
  image: { display: 'block', inlineSize: '100%', aspectRatio: '16 / 9', objectFit: 'cover' },
  fallback: { boxSizing: 'border-box', display: 'flex', alignItems: 'center', justifyContent: 'center', inlineSize: '100%', blockSize: '100%', minBlockSize: 0, padding: tokens.spacingVerticalL, color: tokens.colorNeutralForeground2, backgroundColor: tokens.colorNeutralBackground3, textAlign: 'center', overflowWrap: 'anywhere' },
  // CardPreview owns the 32px overlay. Only introduced object containers inherit those bounds.
  logoContent: { display: 'flex', alignItems: 'center', justifyContent: 'center', inlineSize: '100%', blockSize: '100%', maxInlineSize: '100%', maxBlockSize: '100%', minInlineSize: 0, overflow: 'hidden', objectFit: 'contain', '& > img, & > svg': { maxInlineSize: '100%', maxBlockSize: '100%' } },
});
