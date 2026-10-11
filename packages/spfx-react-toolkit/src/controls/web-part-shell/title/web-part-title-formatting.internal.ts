import type { Editor } from '@tiptap/core';
import type { Mark } from '@tiptap/pm/model';
import { closeHistory } from '@tiptap/pm/history';
import type { PartialTheme } from '@fluentui/react-theme';
import { getTitleDefaultSize, isTitleColor, isTitleHeadingLevel, isTitleLineSpacing, isTitleSize } from './web-part-title-format.internal';
import type { TitleColor, TitleHeadingLevel, TitleLineSpacing, TitleMarks } from './web-part-title.types.internal';

interface TitleBlockFormat {
  type: 'heading' | 'paragraph';
  level?: TitleHeadingLevel;
  lineSpacing?: TitleLineSpacing;
  spaceBefore?: true;
  spaceAfter?: true;
}
export interface TitleSelectionFormat {
  block: TitleBlockFormat;
  bold: boolean | 'mixed'; italic: boolean | 'mixed'; underline: boolean | 'mixed';
  size: number | 'mixed'; color: TitleColor | 'mixed';
}
function blockFormat(editor: Editor): TitleBlockFormat {
  const block = editor.state.doc.firstChild;
  const attrs = block?.attrs ?? {};
  return { type: block?.type.name === 'heading' ? 'heading' : 'paragraph',
    level: isTitleHeadingLevel(attrs.level) ? attrs.level : undefined,
    lineSpacing: isTitleLineSpacing(attrs.lineSpacing) ? attrs.lineSpacing : undefined,
    spaceBefore: attrs.spaceBefore === true ? true : undefined, spaceAfter: attrs.spaceAfter === true ? true : undefined };
}
function readMarks(marks: readonly Mark[]): TitleMarks {
  const style = marks.find(mark => mark.type.name === 'titleStyle')?.attrs;
  return { bold: marks.some(mark => mark.type.name === 'bold') ? true : undefined,
    italic: marks.some(mark => mark.type.name === 'italic') ? true : undefined,
    underline: marks.some(mark => mark.type.name === 'underline') ? true : undefined,
    size: isTitleSize(style?.size) ? style.size : undefined,
    color: isTitleColor(style?.color) ? style.color : undefined };
}
/** Effective selection state, including inherited size/color and caret stored marks. */
export function readTitleSelectionFormat(editor: Editor, theme?: PartialTheme): TitleSelectionFormat {
  const block = blockFormat(editor);
  const defaultSize = getTitleDefaultSize(block, theme);
  const samples: TitleMarks[] = [];
  const { selection, doc } = editor.state;
  if (selection.empty) samples.push(readMarks(editor.state.storedMarks ?? selection.$from.marks()));
  else doc.nodesBetween(selection.from, selection.to, node => { if (node.isInline) samples.push(readMarks(node.marks)); });
  if (!samples.length) samples.push(readMarks(editor.state.storedMarks ?? selection.$from.marks()));
  function common<T>(read: (sample: TitleMarks) => T): T | 'mixed' {
    const first = read(samples[0]); return samples.every(sample => read(sample) === first) ? first : 'mixed';
  }
  return { block, bold: common(sample => sample.bold === true), italic: common(sample => sample.italic === true),
    underline: common(sample => sample.underline === true), size: common(sample => sample.size ?? defaultSize),
    color: common(sample => sample.color ?? 'primary') };
}
/** Clear inline overrides in one undo transaction, preserving text and heading level. */
export function clearTitleFormat(editor: Editor): boolean {
  if (editor.isDestroyed) return false;
  const block = blockFormat(editor);
  const { from, to, empty } = editor.state.selection;
  const tr = closeHistory(editor.state.tr);
  tr.setNodeMarkup(0, editor.schema.nodes[block.type], { level: block.level });
  if (empty) tr.setStoredMarks([]);
  else tr.removeMark(from, to);
  editor.view.dispatch(tr);
  return true;
}
/** Finite integer-only input; invalid drafts never dispatch. */
export function setTitleFontSize(editor: Editor, value: string): boolean {
  if (!/^\d+$/.test(value.trim())) return false;
  const size = Number(value);
  return isTitleSize(size) && editor.commands.setMark('titleStyle', { size });
}
