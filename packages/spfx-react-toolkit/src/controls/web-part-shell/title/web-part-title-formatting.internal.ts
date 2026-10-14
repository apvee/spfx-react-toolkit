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
  size: number | 'mixed';
  /** Undefined is Automatic (inherited); an explicit primary override remains distinct. */
  color: TitleColor | 'mixed' | undefined;
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
/** The title has one block; toolbar caret actions address its complete inline content. */
function toolbarRange(editor: Editor): { from: number; to: number } {
  const { selection, doc } = editor.state;
  const block = doc.firstChild;
  return selection.empty && block && block.content.size > 0
    ? { from: 1, to: block.content.size + 1 } : { from: selection.from, to: selection.to };
}
function readFormat(editor: Editor, theme: PartialTheme | undefined, toolbar: boolean): TitleSelectionFormat {
  const block = blockFormat(editor);
  const defaultSize = getTitleDefaultSize(block, theme);
  const samples: TitleMarks[] = [];
  const { selection, doc } = editor.state;
  const range = toolbar ? toolbarRange(editor) : selection;
  if (range.from === range.to) samples.push(readMarks(editor.state.storedMarks ?? selection.$from.marks()));
  else doc.nodesBetween(range.from, range.to, node => {
    // Breaks carry no persisted formatting; toolbar state describes authored text.
    if (toolbar ? node.isText : node.isInline) samples.push(readMarks(node.marks));
  });
  if (!samples.length) samples.push(readMarks(editor.state.storedMarks ?? selection.$from.marks()));
  function common<T>(read: (sample: TitleMarks) => T): T | 'mixed' {
    const first = read(samples[0]); return samples.every(sample => read(sample) === first) ? first : 'mixed';
  }
  return { block, bold: common(sample => sample.bold === true), italic: common(sample => sample.italic === true),
    underline: common(sample => sample.underline === true), size: common(sample => sample.size ?? defaultSize),
    color: common(sample => sample.color) };
}
/** Effective engine selection state, including caret stored marks for native keyboard commands. */
export function readTitleSelectionFormat(editor: Editor, theme?: PartialTheme): TitleSelectionFormat {
  return readFormat(editor, theme, false);
}
/** Toolbar state describes the entire nonempty title at a caret, or the explicit selected range. */
export function readTitleToolbarFormat(editor: Editor, theme?: PartialTheme): TitleSelectionFormat {
  return readFormat(editor, theme, true);
}
/** Mixed toolbar marks apply uniformly; a second click removes them uniformly. */
export function toggleTitleMark(editor: Editor, name: 'bold' | 'italic' | 'underline'): boolean {
  if (editor.isDestroyed || editor.view.composing) return false;
  const type = editor.schema.marks[name];
  const range = toolbarRange(editor);
  const apply = readTitleToolbarFormat(editor)[name] !== true;
  const caretMarks = editor.state.storedMarks ?? editor.state.selection.$from.marks();
  const tr = closeHistory(editor.state.tr);
  if (range.from !== range.to) {
    if (apply) tr.addMark(range.from, range.to, type.create());
    else tr.removeMark(range.from, range.to, type);
  }
  if (editor.state.selection.empty) {
    const remaining = caretMarks.filter(mark => mark.type !== type);
    tr.setStoredMarks(apply ? type.create().addToSet(remaining) : remaining);
  }
  editor.view.dispatch(tr);
  return true;
}
function setTitleStyle(editor: Editor, update: { size?: number; color?: TitleColor }): boolean {
  if (editor.isDestroyed || editor.view.composing) return false;
  const type = editor.schema.marks.titleStyle;
  const range = toolbarRange(editor);
  const caretMarks = editor.state.storedMarks ?? editor.state.selection.$from.marks();
  const tr = closeHistory(editor.state.tr);
  function updatedStyle(marks: readonly Mark[]): Mark | undefined {
    const attrs = { ...marks.find(mark => mark.type === type)?.attrs, ...update };
    const size = isTitleSize(attrs.size) ? attrs.size : undefined;
    const color = isTitleColor(attrs.color) ? attrs.color : undefined;
    return size === undefined && color === undefined ? undefined : type.create({ size, color });
  }
  if (range.from !== range.to) editor.state.doc.nodesBetween(range.from, range.to, (node, position) => {
    if (!node.isInline) return;
    const from = Math.max(range.from, position), to = Math.min(range.to, position + node.nodeSize);
    const style = updatedStyle(node.marks);
    const previous = node.marks.find(mark => mark.type === type);
    if ((!previous && !style) || (previous && style && previous.eq(style))) return;
    tr.removeMark(from, to, type);
    if (style) tr.addMark(from, to, style);
  });
  if (editor.state.selection.empty) {
    const remaining = caretMarks.filter(mark => mark.type !== type);
    const style = updatedStyle(caretMarks);
    tr.setStoredMarks(style ? style.addToSet(remaining) : remaining);
  }
  editor.view.dispatch(tr);
  return true;
}
/** Automatic removes only the color override, retaining per-run sizes and other marks. */
export function setTitleColor(editor: Editor, color: TitleColor | undefined): boolean {
  return (color === undefined || isTitleColor(color)) && setTitleStyle(editor, { color });
}
/** Clear toolbar-targeted inline overrides and block spacing in one undo transaction, retaining the caret and heading. */
export function clearTitleFormat(editor: Editor): boolean {
  if (editor.isDestroyed || editor.view.composing) return false;
  const block = blockFormat(editor);
  const { from, to } = toolbarRange(editor);
  const tr = closeHistory(editor.state.tr);
  if (block.lineSpacing !== undefined || block.spaceBefore || block.spaceAfter) {
    tr.setNodeMarkup(0, editor.schema.nodes[block.type], { level: block.level });
  }
  if (from !== to) tr.removeMark(from, to);
  if (editor.state.selection.empty) tr.setStoredMarks([]);
  editor.view.dispatch(tr);
  return true;
}
/** Finite integer-only input; invalid drafts never dispatch. */
export function setTitleFontSize(editor: Editor, value: string): boolean {
  if (!/^\d+$/.test(value.trim())) return false;
  const size = Number(value);
  return isTitleSize(size) && setTitleStyle(editor, { size });
}
