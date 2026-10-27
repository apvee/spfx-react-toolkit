import * as React from 'react';
import { EditableText as rootText, type EditableTextProps as RootProps, type EditableTextBlock as RootBlock, type EditableTextChangeData as RootChange, type EditableTextStrings as RootStrings, type EditableTextColorNames as RootColors, type EditableTextStringsOverride as RootOverride } from '@apvee/spfx-react-toolkit';
import { EditableText as domainText, type EditableTextProps as DomainProps, type EditableTextBlock as DomainBlock, type EditableTextChangeData as DomainChange, type EditableTextStrings as DomainStrings, type EditableTextColorNames as DomainColors, type EditableTextStringsOverride as DomainOverride } from '@apvee/spfx-react-toolkit/controls';
import { EditableText as controlText, type EditableTextProps as ControlProps, type EditableTextBlock as ControlBlock, type EditableTextChangeData as ControlChange, type EditableTextStrings as ControlStrings, type EditableTextColorNames as ControlColors, type EditableTextStringsOverride as ControlOverride } from '@apvee/spfx-react-toolkit/controls/editable-text';
export const aliases: typeof rootText[] = [domainText, controlText];
export const props: RootProps = { mode: 'edit', value: '<h2>News</h2>', allowedBlocks: ['h2'], 'aria-label': 'Summary', onChange: data => void data.text };
export const domainProps: DomainProps = props;
export const controlProps: ControlProps = props;
export type Blocks = [RootBlock, DomainBlock, ControlBlock];
export type Changes = [RootChange, DomainChange, ControlChange];
const Text = controlText;
export const element = <Text {...props} ref={React.createRef<HTMLDivElement>()} />;
// @ts-expect-error Edit requires a change callback.
<Text mode="edit" aria-label="Text" />;
// @ts-expect-error Edit requires an accessible name.
<Text mode="edit" onChange={() => undefined} />;
// @ts-expect-error Arbitrary markup injection is private to the control.
<Text dangerouslySetInnerHTML={{ __html: 'unsafe' }} />;
// @ts-expect-error Private editor paths are not exported.
import {} from '@apvee/spfx-react-toolkit/controls/editable-text/editable-text-editor.internal';
// @ts-expect-error Internal document types are not public.
import type { EditableTextDocument } from '@apvee/spfx-react-toolkit/controls/editable-text';
// @ts-expect-error Engine types are implementation-only.
import type { Editor } from '@apvee/spfx-react-toolkit/controls/editable-text';

export type Dictionaries = [RootStrings, DomainStrings, ControlStrings];
export type Colors = [RootColors, DomainColors, ControlColors];
export type Overrides = [RootOverride, DomainOverride, ControlOverride];
export const overrides: RootOverride = { bold: 'Strong', colorNames: { red: 'Danger' }, headingLevel: level => `Heading ${level}`, formattingShortcut: (label, shortcut) => `${label}: ${shortcut}` };
export const partial: ControlProps = { strings: overrides };
declare const full: DomainStrings;
export const complete: DomainProps = { strings: full };
// @ts-expect-error Locale selection was removed; SPFx chooses native resources.
<Text locale="it-IT" />;
// @ts-expect-error Private AST color types remain private.
import type { EditableTextColor } from '@apvee/spfx-react-toolkit/controls/editable-text';
