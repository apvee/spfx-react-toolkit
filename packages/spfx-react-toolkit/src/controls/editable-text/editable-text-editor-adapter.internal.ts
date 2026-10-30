import type { JSONContent } from '@tiptap/core';
import { isEditableTextDocument } from './editable-text-codec.internal';
import type { EditableTextDocument } from './editable-text-document.types.internal';

/** Convert only validated private protocol documents to the engine schema. */
export function editableTextDocumentToEditorContent(document: EditableTextDocument): JSONContent {
  if (!isEditableTextDocument(document)) throw new TypeError('Invalid editable text document');
  return { type: 'doc', content: [{ type: document.type,
    attrs: { ...(document.type === 'heading' ? { level: document.level } : {}),
      lineSpacing: document.lineSpacing, spaceBefore: document.spaceBefore, spaceAfter: document.spaceAfter },
    content: document.content.flatMap(inline => {
      if (inline.type === 'hardBreak') return [{ type: 'hardBreak' }];
      if (!inline.text) return [];
      const marks: NonNullable<JSONContent['marks']> = [];
      for (const type of ['bold', 'italic', 'underline'] as const) if (inline.marks?.[type]) marks.push({ type });
      if (inline.marks?.size !== undefined || inline.marks?.color !== undefined) {
        marks.push({ type: 'titleStyle', attrs: { size: inline.marks?.size, color: inline.marks?.color } });
      }
      return [{ type: 'text', text: inline.text, ...(marks.length ? { marks } : {}) }];
    }) }] };
}

function checkAttributes(attributes: Record<string, unknown> | undefined, allowed: string[]): void {
  if (attributes && Object.keys(attributes).some(key => !allowed.includes(key))) throw new TypeError('Invalid editable text attributes');
}

/** Reject engine content outside the finite single-block protocol before publication. */
export function editorContentToEditableTextDocument(content: JSONContent): EditableTextDocument {
  if (content.type !== 'doc' || content.content?.length !== 1) throw new TypeError('Invalid editable text editor document');
  checkAttributes(content.attrs, []);
  const block = content.content[0];
  if (block.type !== 'heading' && block.type !== 'paragraph') throw new TypeError('Invalid editable text block');
  checkAttributes(block.attrs, [...(block.type === 'heading' ? ['level'] : []), 'lineSpacing', 'spaceBefore', 'spaceAfter']);
  const inlines: unknown[] = (block.content ?? []).map(inline => {
    checkAttributes(inline.attrs, []);
    if (inline.type === 'hardBreak') return { type: 'hardBreak' };
    if (inline.type !== 'text' || typeof inline.text !== 'string') throw new TypeError('Invalid editable text inline');
    const marks: Record<string, unknown> = {};
    for (const mark of inline.marks ?? []) {
      if (mark.type === 'bold' || mark.type === 'italic' || mark.type === 'underline') {
        checkAttributes(mark.attrs, []);
        marks[mark.type] = true;
      } else if (mark.type === 'titleStyle') {
        checkAttributes(mark.attrs, ['size', 'color']);
        if (mark.attrs?.size !== undefined) marks.size = mark.attrs.size;
        if (mark.attrs?.color !== undefined) marks.color = mark.attrs.color;
      } else throw new TypeError('Invalid editable text mark');
    }
    return { type: 'text', text: inline.text, ...(Object.keys(marks).length ? { marks } : {}) };
  });
  const candidate: unknown = { type: block.type, ...(block.type === 'heading' ? { level: block.attrs?.level } : {}),
    content: inlines, ...(block.attrs?.lineSpacing !== undefined ? { lineSpacing: block.attrs.lineSpacing } : {}),
    ...(block.attrs?.spaceBefore !== undefined ? { spaceBefore: block.attrs.spaceBefore } : {}),
    ...(block.attrs?.spaceAfter !== undefined ? { spaceAfter: block.attrs.spaceAfter } : {}) };
  if (!isEditableTextDocument(candidate)) throw new TypeError('Invalid editable text editor attributes');
  return candidate;
}
