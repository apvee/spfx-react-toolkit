import * as React from 'react';
import { Checkbox, Option, OptionGroup, Radio, TagPickerOption, TagPickerOptionGroup } from '@fluentui/react-components';
import type { FormOption, FormValueCodec } from './form-value.types';

interface SelectionSource {
  options?: readonly FormOption[];
  children?: React.ReactNode;
  entries: readonly FormOption[];
}
/** Reads explicit primitive identities without adapting arbitrary consumer children. */
export function getSelectionSource(options: readonly FormOption[] | undefined, children: React.ReactNode): SelectionSource {
  if (options !== undefined && children !== undefined) throw new TypeError('Form selection accepts options OR children, never both.');
  const entries: FormOption[] = [];
  function inspect(nodes: React.ReactNode, ancestorDisabled = false): void {
    React.Children.forEach(nodes, child => {
      if (!React.isValidElement<{ value?: unknown; label?: React.ReactNode; text?: string; disabled?: boolean; 'aria-disabled'?: boolean | 'true' | 'false'; children?: React.ReactNode }>(child)) return;
      const props = child.props;
      if (child.type === 'option' || child.type === Option || child.type === Radio || child.type === Checkbox || child.type === TagPickerOption) {
        if (typeof props.value !== 'string') throw new TypeError('Form option children require an explicit string value.');
        entries.push({ value: props.value, label: props.label ?? props.children, textValue: props.text, disabled: ancestorDisabled || props.disabled || props['aria-disabled'] === true || props['aria-disabled'] === 'true' });
      } else if (child.type === React.Fragment || child.type === OptionGroup || child.type === TagPickerOptionGroup || typeof child.type === 'string') inspect(props.children, ancestorDisabled || props.disabled === true);
    });
  }
  if (options === undefined) inspect(children);
  else entries.push(...options);
  const identities = new Set<string>();
  for (const entry of entries) {
    if (typeof entry.value !== 'string' || identities.has(entry.value)) throw new TypeError('Form options require unique string identities.');
    identities.add(entry.value);
  }
  return { options, children, entries };
}
export function getOptionText(option: FormOption): string {
  if (option.textValue !== undefined) return option.textValue;
  if (typeof option.label === 'string') return option.label;
  if (typeof option.label === 'number') return String(option.label);
  return '';
}
export function getSelectionText(source: SelectionSource, value: string): string {
  const option = source.entries.find(entry => entry.value === value);
  return option ? getOptionText(option) : '';
}
export function selectionIsDisabled(source: SelectionSource, value: string): boolean {
  return source.entries.find(entry => entry.value === value)?.disabled === true;
}
export function renderSelectionOptions(source: SelectionSource): React.ReactNode {
  return source.options === undefined ? source.children : source.options.map(option => React.createElement(Option, { key: option.value, value: option.value, text: getOptionText(option), disabled: option.disabled }, option.label));
}
function assertSelectionValue(value: unknown, multi: boolean): void {
  if (multi ? !Array.isArray(value) || value.some(item => typeof item !== 'string') : typeof value !== 'string') throw new TypeError('Form selection codec must produce the complete string/string-array control domain.');
}
export function readSelectionValue<TValue extends string | string[]>(value: unknown, codec: FormValueCodec<unknown, TValue> | undefined, empty: TValue): TValue {
  const control = codec === undefined ? value === undefined ? empty : value : codec.toControl(value);
  assertSelectionValue(control, Array.isArray(empty));
  return control as TValue;
}
export function writeSelectionValue<TValue extends string | string[]>(value: TValue, codec: FormValueCodec<unknown, TValue> | undefined): unknown {
  if (codec === undefined) return value;
  const field = codec.fromControl(value);
  const roundTrip = codec.toControl(field);
  assertSelectionValue(roundTrip, Array.isArray(value));
  if (Array.isArray(value) ? !Array.isArray(roundTrip) || value.length !== roundTrip.length || value.some((entry, index) => entry !== roundTrip[index]) : value !== roundTrip) throw new TypeError('Form selection codec must losslessly round-trip every writable identity.');
  return field;
}
