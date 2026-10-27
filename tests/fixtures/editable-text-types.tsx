/// <reference path="../../packages/spfx-react-toolkit/src/localization/native-resources.internal.d.ts" />
import * as React from 'react';
import { EditableText } from '../../packages/spfx-react-toolkit/src/controls/editable-text/editable-text';
import type { EditableTextProps, EditableTextBlock, EditableTextChangeData } from '../../packages/spfx-react-toolkit/src/controls/editable-text/editable-text.types';
const read: EditableTextProps = {};
const blocks: readonly EditableTextBlock[] = ['h2', 'h4'];
const change = (data: EditableTextChangeData): void => { console.log(data.html, data.text); };
declare const mode: 'read' | 'edit';
const dynamic: EditableTextProps = { mode, onChange: change, 'aria-label': 'Summary', allowedBlocks: blocks };
const labelled: EditableTextProps = { mode: 'edit', onChange: change, 'aria-labelledby': 'summary-label' };
const ref = React.createRef<HTMLDivElement>();
const element = <EditableText {...dynamic} ref={ref} title="Summary" style={{ margin: 2 }} data-record="1" onClick={event => event.currentTarget.focus()} />;
// @ts-expect-error Edit requires a domain change callback.
const missingChange: EditableTextProps = { mode: 'edit', 'aria-label': 'Summary' };
// @ts-expect-error Edit requires an accessible name.
const missingName: EditableTextProps = { mode: 'edit', onChange: change };
// @ts-expect-error H5 is outside the persisted protocol.
const unsupported: EditableTextProps = { allowedBlocks: ['h5'] };
// @ts-expect-error The control owns children.
const children = <EditableText>raw</EditableText>;
// @ts-expect-error Raw markup injection is not a public capability.
const html: EditableTextProps = { dangerouslySetInnerHTML: { __html: '<p>raw</p>' } };
// @ts-expect-error The editor owns its contenteditable semantics.
const editable: EditableTextProps = { contentEditable: true };
// @ts-expect-error The editor owns its native role.
const role: EditableTextProps = { role: 'textbox' };
// @ts-expect-error Native uncontrolled defaults conflict with controlled value.
const uncontrolled: EditableTextProps = { defaultValue: 'raw' };
// @ts-expect-error The control does not expose contenteditable warning flags.
const suppress: EditableTextProps = { suppressContentEditableWarning: true };
// @ts-expect-error Direction belongs to the host FluentProvider.
const direction: EditableTextProps = { dir: 'rtl' };
// @ts-expect-error No editor internals are exposed.
const engine: EditableTextProps = { editor: {} };
// @ts-expect-error No SPFx provider or instance is required or accepted.
const instance: EditableTextProps = { instance: {} };
// @ts-expect-error Themes come from FluentProvider.
const theme: EditableTextProps = { theme: {} };
// @ts-expect-error Owner document comes from host context and the mounted root.
const owner: EditableTextProps = { ownerDocument: document };
void [read, labelled, element, missingChange, missingName, unsupported, children, html, editable, role, uncontrolled, suppress, direction, engine, instance, theme, owner];

import type { EditableTextStrings, EditableTextColorNames, EditableTextStringsOverride } from '../../packages/spfx-react-toolkit/src/controls/editable-text';
const colors: Partial<EditableTextColorNames> = { red: 'Danger', brand: 'Accent' };
const dictionary: EditableTextStringsOverride = { bold: 'Strong', colorNames: colors, headingLevel: level => `Heading ${level}`, formattingShortcut: (label, shortcut) => `${label}: ${shortcut}` };
declare const fullDictionary: EditableTextStrings;
const custom: EditableTextProps = { strings: dictionary };
const complete: EditableTextProps = { strings: fullDictionary };
// @ts-expect-error Runtime language selection was replaced by SPFx resources and dictionaries.
const removedLocale: EditableTextProps = { locale: 'it-IT' };
// @ts-expect-error Color names use public semantic fields, never arbitrary keys.
const invalidColor: EditableTextStringsOverride = { colorNames: { invented: 'Custom' } };
// @ts-expect-error Dictionary hints never supply an implicit accessible name.
const unnamedCustom: EditableTextProps = { mode: 'edit', strings: dictionary, onChange: change };
void [custom, complete, removedLocale, invalidColor, unnamedCustom];

// @ts-expect-error No dictionary can supply an implicit standalone textbox name.
const implicitName: EditableTextStringsOverride = { editorLabel: 'Text' };
void implicitName;
