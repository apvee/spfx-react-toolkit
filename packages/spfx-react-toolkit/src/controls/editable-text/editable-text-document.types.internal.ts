/** Finite semantic foreground roles persisted by the editable text protocol. */
export type EditableTextColor =
  | 'primary' | 'subtle' | 'brand' | 'neutralMuted' | 'brandAlt'
  | 'neutralDark' | 'neutralLight' | 'neutralLighter' | 'brandDarkest' | 'brandDarker' | 'brandDark' | 'brandLight'
  | 'darkRed' | 'amber' | 'yellowGreen' | 'cyan' | 'navy'
  | 'red' | 'orange' | 'yellow' | 'green' | 'teal' | 'blue' | 'purple' | 'pink' | 'brown';
/** Supported heading levels. */
export type EditableTextHeadingLevel = 1 | 2 | 3 | 4;
/** Supported unitless line spacing values. */
export type EditableTextLineSpacing = 1 | 1.2 | 1.4 | 1.6 | 2 | 2.4;

/** Optional inline overrides; absence uses the block preset and theme. */
export interface EditableTextMarks {
  bold?: true;
  italic?: true;
  underline?: true;
  /** Integer pixels in the inclusive 8–96 range. */
  size?: number;
  color?: EditableTextColor;
}

/** An inline text run or a hard break within the single editable text block. */
export type EditableTextInline =
  | { type: 'text'; text: string; marks?: EditableTextMarks }
  | { type: 'hardBreak' };

interface EditableTextBlockContent {
  content: EditableTextInline[];
  /** Absence uses the block's Fluent typography preset line-height token. */
  lineSpacing?: EditableTextLineSpacing;
  /** Active spacing uses the host theme's spacingVerticalS. */
  spaceBefore?: true;
  /** Active spacing uses the host theme's spacingVerticalS. */
  spaceAfter?: true;
}

/** Private engine-independent editable text document containing exactly one block. */
export type EditableTextDocument = EditableTextBlockContent & (
  | { type: 'heading'; level: EditableTextHeadingLevel }
  | { type: 'paragraph' }
);

/** Invalid input carries a safe document; resolve replaces it with literal fallback. */
export interface EditableTextParseResult {
  document: EditableTextDocument;
  /** Valid means canonical input; normalized means recoverable noncanonical input. */
  kind: 'valid' | 'normalized' | 'invalid';
}
