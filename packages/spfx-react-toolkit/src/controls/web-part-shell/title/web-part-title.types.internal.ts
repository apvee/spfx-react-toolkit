/** Finite semantic foreground roles persisted by the title protocol. */
export type TitleColor =
  | 'primary' | 'subtle' | 'brand' | 'neutralMuted' | 'brandAlt'
  | 'neutralDark' | 'neutralLight' | 'neutralLighter' | 'brandDarkest' | 'brandDarker' | 'brandDark' | 'brandLight'
  | 'darkRed' | 'amber' | 'yellowGreen' | 'cyan' | 'navy'
  | 'red' | 'orange' | 'yellow' | 'green' | 'teal' | 'blue' | 'purple' | 'pink' | 'brown';
/** Supported heading levels. */
export type TitleHeadingLevel = 1 | 2 | 3 | 4;
/** Supported unitless line spacing values. */
export type TitleLineSpacing = 1 | 1.2 | 1.4 | 1.6 | 2 | 2.4;

/** Optional inline overrides; absence uses the block preset and theme. */
export interface TitleMarks {
  bold?: true;
  italic?: true;
  underline?: true;
  /** Integer pixels in the inclusive 8–96 range. */
  size?: number;
  color?: TitleColor;
}

/** An inline text run or a hard break within the single title block. */
export type TitleInline =
  | { type: 'text'; text: string; marks?: TitleMarks }
  | { type: 'hardBreak' };

interface TitleBlock {
  content: TitleInline[];
  /** Absence uses the block's Fluent typography preset line-height token. */
  lineSpacing?: TitleLineSpacing;
  /** Active spacing uses the host theme's spacingVerticalS. */
  spaceBefore?: true;
  /** Active spacing uses the host theme's spacingVerticalS. */
  spaceAfter?: true;
}

/** Private engine-independent title document containing exactly one block. */
export type TitleDocument = TitleBlock & (
  | { type: 'heading'; level: TitleHeadingLevel }
  | { type: 'paragraph' }
);

/** Invalid input carries a safe document; resolve replaces it with literal fallback. */
export interface TitleParseResult {
  document: TitleDocument;
  /** Valid means canonical input; normalized means recoverable noncanonical input. */
  kind: 'valid' | 'normalized' | 'invalid';
}
