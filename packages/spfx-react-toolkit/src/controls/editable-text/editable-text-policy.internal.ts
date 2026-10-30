import type { EditableTextBlock } from './editable-text.types';
import type { EditableTextDocument, EditableTextHeadingLevel } from './editable-text-document.types.internal';

/** Resolved per-instance semantic membership; order matches the block menu. */
export interface EditableTextBlockPolicy {
  readonly blocks: readonly EditableTextBlock[];
  readonly defaultBlock: EditableTextBlock;
  /** Stable for equivalent sets; deliberately independent of defaultBlock. */
  readonly membershipKey: string;
}

const canonicalBlocks: readonly EditableTextBlock[] = ['h1', 'h2', 'h3', 'h4', 'p'];
function isBlock(value: unknown): value is EditableTextBlock {
  return typeof value === 'string' && canonicalBlocks.some(block => block === value);
}

/** Validate configuration without silently extending membership or retaining caller arrays. */
export function resolveEditableTextBlockPolicy(allowedBlocks: readonly EditableTextBlock[] | undefined,
  defaultBlock: EditableTextBlock | undefined): EditableTextBlockPolicy {
  if (allowedBlocks !== undefined && (!Array.isArray(allowedBlocks) || allowedBlocks.length === 0 ||
    !Array.from(allowedBlocks).every(isBlock))) {
    throw new TypeError('EditableText allowedBlocks must be a nonempty array containing only p, h1, h2, h3 or h4');
  }
  const blocks = canonicalBlocks.filter(block => allowedBlocks === undefined || allowedBlocks.includes(block));
  if (defaultBlock !== undefined && (!isBlock(defaultBlock) || !blocks.includes(defaultBlock))) {
    throw new TypeError('EditableText defaultBlock must belong to allowedBlocks');
  }
  return {
    blocks,
    defaultBlock: defaultBlock ?? (blocks.includes('p') ? 'p' : blocks[0]),
    membershipKey: blocks.join(',')
  };
}

/**
 * Convert only an excluded or empty block; retain inline content, marks and spacing.
 * A compliant document preserves reference identity. No mutation, persistence or engine work.
 */
export function normalizeEditableTextDocument(document: EditableTextDocument,
  policy: EditableTextBlockPolicy): EditableTextDocument {
  const current: EditableTextBlock = document.type === 'paragraph' ? 'p' : `h${document.level}`;
  const empty = document.content.every(inline => inline.type === 'hardBreak' || inline.text.trim() === '');
  const target = empty || !policy.blocks.includes(current) ? policy.defaultBlock : current;
  if (target === current) return document;
  if (target === 'p') {
    if (document.type === 'paragraph') return document;
    return {
      type: 'paragraph',
      content: document.content,
      ...('lineSpacing' in document ? { lineSpacing: document.lineSpacing } : {}),
      ...('spaceBefore' in document ? { spaceBefore: document.spaceBefore } : {}),
      ...('spaceAfter' in document ? { spaceAfter: document.spaceAfter } : {})
    };
  }
  // Membership validation restricts heading tags to exactly h1–h4.
  const level = Number(target.slice(1)) as EditableTextHeadingLevel;
  return { ...document, type: 'heading', level };
}
