import * as React from 'react';
import { Checkbox, Option, Radio, TagPickerOption } from '@fluentui/react-components';
import { createForm } from '../../packages/spfx-react-toolkit/src/controls/form';
import type { FormValueCodec } from '../../packages/spfx-react-toolkit/src/controls/form/form-value.types';
const F = createForm<{ text: string; optional?: string; literal: 'a' | 'b'; number: number; object: { id: string }; many: string[]; numbers: number[]; nullable: string | null }>();
const options = [{ value: 'a', label: 'Alpha' }, { value: 'b', label: 'Beta' }];
const select = React.createRef<HTMLSelectElement>();
const input = React.createRef<HTMLInputElement>();
const button = React.createRef<HTMLButtonElement>();
const group = React.createRef<HTMLDivElement>();
const objectCodec: FormValueCodec<{ id: string }, string> = { toControl: value => value.id, fromControl: id => ({ id }) };
const numberCodec: FormValueCodec<number, string> = { toControl: String, fromControl: Number };
const numbersCodec: FormValueCodec<number[], string[]> = { toControl: value => value.map(String), fromControl: value => value.map(Number) };
<F.Select name="text" options={options} ref={select} onChange={(event, data) => { const native: HTMLSelectElement = event.currentTarget; data.value.toUpperCase(); void native; }} />;
<F.Select name="optional"><option value="a">Alpha</option></F.Select>;
<F.RadioGroup name="text" ref={group}><Radio value="a" label="Alpha" /></F.RadioGroup>;
<F.Dropdown name="object" options={options} codec={objectCodec} ref={button} />;
<F.Combobox name="number" options={options} codec={numberCodec} ref={input} onChange={event => { const native: HTMLInputElement = event.currentTarget; void native; }} />;
<F.Select name="literal" options={options} codec={{ toControl: value => value, fromControl: value => value === 'b' ? 'b' : 'a' }} />;
<F.CheckboxGroup name="many" ref={group}><Checkbox value="a" label="Alpha" /></F.CheckboxGroup>;
<F.DropdownMulti name="many" options={options} ref={button} />;
<F.ComboboxMulti name="numbers" options={options} codec={numbersCodec} ref={input} />;
<F.TagPicker name="many" label="Tags" options={options} ref={input} strings={{ noOptions: 'No matches' }} />;
<F.TagPicker name="many" aria-label="Tags"><TagPickerOption value="a" text="Alpha">Alpha</TagPickerOption></F.TagPicker>;
<F.Combobox name="text"><Option value="a">Alpha</Option></F.Combobox>;
// @ts-expect-error full string writer cannot target literal fields without an explicit codec
<F.Select name="literal" options={options} />;
// @ts-expect-error number values need an explicit codec
<F.Dropdown name="number" options={options} />;
// @ts-expect-error codec's input must correlate with this exact path
<F.Select name="object" options={options} codec={numberCodec} />;
// @ts-expect-error object codec must correlate with this exact path
<F.Combobox name="number" options={options} codec={objectCodec} />;
// @ts-expect-error nullable read domain cannot use direct string binding
<F.Select name="nullable" options={options} />;
// @ts-expect-error string[] writer cannot target numeric array without codec
<F.ComboboxMulti name="numbers" options={options} />;
// @ts-expect-error single controls cannot use array codecs
<F.Dropdown name="many" options={options} codec={numbersCodec} />;
// @ts-expect-error multi controls require string[] codec write domain
<F.DropdownMulti name="number" options={options} codec={numberCodec} />;
// @ts-expect-error options and children are exclusive
<F.Select name="text" options={options}><option value="a">Alpha</option></F.Select>;
// @ts-expect-error native select ref is not a group div ref
<F.RadioGroup name="text" options={options} ref={select} />;
// @ts-expect-error native combobox input ref is not a button ref
<F.Combobox name="text" options={options} ref={button} />;
// @ts-expect-error TagPicker requires a visible text label or explicit accessible name
<F.TagPicker name="many" options={options} />;
// @ts-expect-error rich label needs explicit accessible name
<F.TagPicker name="many" label={<span>Tags</span>} options={options} />;
// @ts-expect-error RHF owns controlled defaults
<F.Select name="text" options={options} defaultValue="a" />;
// @ts-expect-error Input deliberately has no codec
<F.Input name="number" codec={numberCodec} />;
