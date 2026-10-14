import { Extension, Mark, type Extensions } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { toggleTitleMark } from './web-part-title-formatting.internal';
import { editorContentToTitleDocument } from './web-part-title-editor-adapter.internal';
import Document from '@tiptap/extension-document';
import Text from '@tiptap/extension-text';
import Paragraph from '@tiptap/extension-paragraph';
import Heading from '@tiptap/extension-heading';
import HardBreak from '@tiptap/extension-hard-break';
import Bold from '@tiptap/extension-bold';
import Italic from '@tiptap/extension-italic';
import Underline from '@tiptap/extension-underline';
import { Placeholder, UndoRedo } from '@tiptap/extensions';
import { getTitleClassColor, getTitleClassLine, getTitleClassSize, getTitleLineClass, isTitleColor, isTitleLineSpacing, isTitleSize } from './web-part-title-format.internal';

const TitleStyle = Mark.create({
  name: 'titleStyle',
  addAttributes() {
    return { size: { default: undefined }, color: { default: undefined } };
  },
  parseHTML() {
    return [{ tag: 'span', getAttrs: element => {
      if (typeof element === 'string') return false;
      const classes = Array.from(element.classList);
      const size = classes.map(getTitleClassSize).find(value => value !== undefined);
      const color = classes.map(getTitleClassColor).find(value => value !== undefined);
      return size === undefined && color === undefined ? false : { size, color };
    } }];
  },
  renderHTML({ mark }) {
    const classes = [isTitleColor(mark.attrs.color) ? `wps-color-${mark.attrs.color}` : undefined,
      isTitleSize(mark.attrs.size) ? `wps-size-${mark.attrs.size}` : undefined].filter(Boolean);
    return ['span', { class: classes.join(' ') }, 0];
  }
});
const TitleBlockAttributes = Extension.create({
  name: 'titleBlockAttributes',
  addGlobalAttributes() {
    return [{ types: ['heading', 'paragraph'], attributes: {
      lineSpacing: { default: undefined, parseHTML: element => Array.from(element.classList).map(getTitleClassLine).find(value => value !== undefined),
        renderHTML: attrs => isTitleLineSpacing(attrs.lineSpacing) ? { class: getTitleLineClass(attrs.lineSpacing) } : {} },
      spaceBefore: { default: undefined, parseHTML: element => element.classList.contains('wps-space-before') ? true : undefined,
        renderHTML: attrs => attrs.spaceBefore === true ? { class: 'wps-space-before' } : {} },
      spaceAfter: { default: undefined, parseHTML: element => element.classList.contains('wps-space-after') ? true : undefined,
        renderHTML: attrs => attrs.spaceAfter === true ? { class: 'wps-space-after' } : {} }
    } }];
  }
});
const SingleBlockBreak = HardBreak.extend({
  addKeyboardShortcuts() {
    const insert = (): boolean => this.editor.view.composing ? false : this.editor.commands.setHardBreak();
    return { Enter: insert, 'Shift-Enter': insert, 'Mod-Enter': insert };
  }
});

// Preserve each native mark's commands and input rules; only keyboard scope
// follows the same title/range contract as its toolbar command.
const TitleBold = Bold.extend({ addKeyboardShortcuts() {
  const toggle = (): boolean => toggleTitleMark(this.editor, 'bold');
  return { 'Mod-b': toggle, 'Mod-B': toggle };
} });
const TitleItalic = Italic.extend({ addKeyboardShortcuts() {
  const toggle = (): boolean => toggleTitleMark(this.editor, 'italic');
  return { 'Mod-i': toggle, 'Mod-I': toggle };
} });
const TitleUnderline = Underline.extend({ addKeyboardShortcuts() {
  const toggle = (): boolean => toggleTitleMark(this.editor, 'underline');
  return { 'Mod-u': toggle, 'Mod-U': toggle };
} });

const TitleProtocol = Extension.create({
  name: 'titleProtocol',
  addProseMirrorPlugins() {
    return [new Plugin({ filterTransaction: transaction => {
      if (!transaction.docChanged) return true;
      try { editorContentToTitleDocument(transaction.doc.toJSON()); return true; }
      catch { return false; }
    } })];
  }
});

/** Minimal engine extensions; neither engine UI nor unrestricted style parsing is installed. */
export function createTitleExtensions(placeholder: string | (() => string)): Extensions {
  return [Document.extend({ content: '(heading|paragraph)' }), Text, Paragraph, Heading.configure({ levels: [1, 2, 3, 4] }),
    SingleBlockBreak, TitleBold, TitleItalic, TitleUnderline, TitleStyle, TitleBlockAttributes, TitleProtocol,
    UndoRedo, Placeholder.configure({ placeholder })];
}
