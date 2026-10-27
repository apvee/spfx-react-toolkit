import { Extension, InputRule, Mark, type Extensions } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { normalizeEditableTextDocument, resolveEditableTextBlockPolicy } from './editable-text-policy.internal';
import type { EditableTextBlock } from './editable-text.types';
import type { EditableTextBlockPolicy } from './editable-text-policy.internal';
import { toggleEditableTextMark } from './editable-text-formatting.internal';
import { editorContentToEditableTextDocument, editableTextDocumentToEditorContent } from './editable-text-editor-adapter.internal';
import Document from '@tiptap/extension-document';
import Text from '@tiptap/extension-text';
import Paragraph from '@tiptap/extension-paragraph';
import Heading from '@tiptap/extension-heading';
import HardBreak from '@tiptap/extension-hard-break';
import Bold from '@tiptap/extension-bold';
import Italic from '@tiptap/extension-italic';
import Underline from '@tiptap/extension-underline';
import { Placeholder, UndoRedo } from '@tiptap/extensions';
import { getEditableTextClassColor, getEditableTextClassLine, getEditableTextClassSize, getEditableTextLineClass, isEditableTextColor, isEditableTextHeadingLevel, isEditableTextLineSpacing, isEditableTextSize } from './editable-text-format.internal';

// Retain the established private mark identity while legacy title adapters share this engine.
const EditableTextStyle = Mark.create({
  name: 'titleStyle',
  addAttributes() {
    return { size: { default: undefined }, color: { default: undefined } };
  },
  parseHTML() {
    return [{ tag: 'span', getAttrs: element => {
      if (typeof element === 'string') return false;
      const classes = Array.from(element.classList);
      const size = classes.map(getEditableTextClassSize).find(value => value !== undefined);
      const color = classes.map(getEditableTextClassColor).find(value => value !== undefined);
      return size === undefined && color === undefined ? false : { size, color };
    } }];
  },
  renderHTML({ mark }) {
    const classes = [isEditableTextColor(mark.attrs.color) ? `wps-color-${mark.attrs.color}` : undefined,
      isEditableTextSize(mark.attrs.size) ? `wps-size-${mark.attrs.size}` : undefined].filter(Boolean);
    return ['span', { class: classes.join(' ') }, 0];
  }
});
const EditableTextBlockAttributes = Extension.create({
  name: 'editableTextBlockAttributes',
  addGlobalAttributes() {
    return [{ types: ['heading', 'paragraph'], attributes: {
      lineSpacing: { default: undefined, parseHTML: element => Array.from(element.classList).map(getEditableTextClassLine).find(value => value !== undefined),
        renderHTML: attrs => isEditableTextLineSpacing(attrs.lineSpacing) ? { class: getEditableTextLineClass(attrs.lineSpacing) } : {} },
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
// follows the same text/range contract as its toolbar command.
const EditableTextBold = Bold.extend({ addKeyboardShortcuts() {
  const toggle = (): boolean => toggleEditableTextMark(this.editor, 'bold');
  return { 'Mod-b': toggle, 'Mod-B': toggle };
} });
const EditableTextItalic = Italic.extend({ addKeyboardShortcuts() {
  const toggle = (): boolean => toggleEditableTextMark(this.editor, 'italic');
  return { 'Mod-i': toggle, 'Mod-I': toggle };
} });
const EditableTextUnderline = Underline.extend({ addKeyboardShortcuts() {
  const toggle = (): boolean => toggleEditableTextMark(this.editor, 'underline');
  return { 'Mod-u': toggle, 'Mod-U': toggle };
} });

/** Stable schema; each closure reads this instance's current membership. */
export function createEditableTextExtensions(placeholder: string | (() => string),
  getPolicy: () => EditableTextBlockPolicy = () => resolveEditableTextBlockPolicy(undefined, undefined),
  composing: () => boolean = () => false): Extensions {
  const permitsHeading = (level: number): boolean => isEditableTextHeadingLevel(level) && getPolicy().blocks.includes(`h${level}`);
  const PolicyParagraph = Paragraph.extend({
    addCommands() { return { setParagraph: () => ({ commands }) =>
      getPolicy().blocks.includes('p') && commands.setNode('paragraph') }; }
  });
  const PolicyHeading = Heading.extend({
    addCommands() { return {
      setHeading: attributes => ({ commands }) => permitsHeading(attributes.level) && commands.setNode('heading', attributes),
      toggleHeading: attributes => ({ commands, state }) => {
        if (!permitsHeading(attributes.level)) return false;
        if (getPolicy().blocks.includes('p')) return commands.toggleNode('heading', 'paragraph', attributes);
        const block = state.doc.firstChild;
        return (block?.type.name === 'heading' && block.attrs.level === attributes.level) || commands.setNode('heading', attributes);
      }
    }; },
    addInputRules() {
      return [new InputRule({ find: /^(#{1,4})\s$/, handler: ({ state, range, match }) => {
        const level = match[1].length;
        // Tiptap's InputRule API requires null to leave an excluded marker literal.
        if (!permitsHeading(level)) return null;
        state.tr.delete(range.from, range.to).setBlockType(range.from, range.from, this.type, { level }).setMeta('editableTextBlockChoice', true);
      } })];
    }
  });
  const EditableTextProtocol = Extension.create({
    name: 'editableTextProtocol',
    addCommands() { return { clearContent: (emitUpdate = true) => ({ commands }) => commands.setContent(
      editableTextDocumentToEditorContent(normalizeEditableTextDocument({ type: 'paragraph', content: [] }, getPolicy())), { emitUpdate }) }; },
    addProseMirrorPlugins() {
      return [new Plugin({ filterTransaction: transaction => {
        if (!transaction.docChanged) return true;
        try { editorContentToEditableTextDocument(transaction.doc.toJSON()); return true; }
        catch { return false; }
      }, appendTransaction: (transactions, oldState, state) => {
        if (composing() || !transactions.some(transaction => transaction.docChanged)) return;
        const document = editorContentToEditableTextDocument(state.doc.toJSON());
        const policy = getPolicy();
        const block: EditableTextBlock = document.type === 'paragraph' ? 'p' : `h${document.level}`;
        // An explicit allowed empty heading is a typing choice. Only deletion of
        // nonempty content or an excluded block projects to the effective default.
        const emptied = oldState.doc.textContent.length > 0 && state.doc.textContent.length === 0 &&
          !transactions.some(transaction => transaction.getMeta('editableTextBlockChoice') === true);
        if (!emptied && policy.blocks.includes(block)) return;
        const normalized = normalizeEditableTextDocument(document, policy);
        if (normalized === document) return;
        const attrs = { ...state.doc.firstChild?.attrs, ...(normalized.type === 'heading' ? { level: normalized.level } : {}) };
        if (normalized.type === 'paragraph') delete attrs.level;
        const tr = state.tr.setNodeMarkup(0, state.schema.nodes[normalized.type], attrs);
        if (state.storedMarks) tr.setStoredMarks(state.storedMarks);
        return tr;
      } })];
    }
  });
  // Minimal engine: no unrestricted style parsing or global CSS injection.
  return [Document.extend({ content: '(paragraph|heading)' }), Text, PolicyParagraph, PolicyHeading.configure({ levels: [1, 2, 3, 4] }),
    SingleBlockBreak, EditableTextBold, EditableTextItalic, EditableTextUnderline, EditableTextStyle, EditableTextBlockAttributes, EditableTextProtocol,
    UndoRedo, Placeholder.configure({ placeholder })];
}
