import * as React from 'react';
import { EditorContent, useEditor } from '@tiptap/react';
import type { Editor } from '@tiptap/core';
import Document from '@tiptap/extension-document';
import Text from '@tiptap/extension-text';
import Paragraph from '@tiptap/extension-paragraph';
import Heading from '@tiptap/extension-heading';
import HardBreak from '@tiptap/extension-hard-break';
import Bold from '@tiptap/extension-bold';
import Italic from '@tiptap/extension-italic';
import Underline from '@tiptap/extension-underline';
// Named exports preserve the existing NodeJs declaration resolver. The same
// extensions' exports-only subpaths are checked separately for runtime resolution.
import { UndoRedo, Placeholder } from '@tiptap/extensions';

// Engine types stay in this private test fixture, never in the library API.
export interface DependencyProbeProps {
  onEditor: (editor: Editor) => void;
  onDestroy: () => void;
}

export function DependencyProbe({ onEditor, onDestroy }: DependencyProbeProps): React.ReactElement {
  const editor = useEditor({
    extensions: [Document, Text, Paragraph, Heading.configure({ levels: [1, 2, 3, 4] }),
      HardBreak, Bold, Italic, Underline, UndoRedo, Placeholder.configure({ placeholder: 'Title' })],
    content: '<h2>Title</h2>',
    injectCSS: false,
    immediatelyRender: false,
    onCreate: ({ editor: createdEditor }) => onEditor(createdEditor),
    onDestroy,
  });
  return <EditorContent editor={editor} />;
}

export const dependencyIdentity = { React, useEditor, EditorContent, UndoRedo, Placeholder };
