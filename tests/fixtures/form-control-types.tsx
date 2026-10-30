import * as React from 'react';
import { createForm } from '../../packages/spfx-react-toolkit/src/controls/form';
import type { StandardSchemaV1 } from '@standard-schema/spec';
import type { SpinButtonChangeEvent, SpinButtonOnChangeData } from '@fluentui/react-components';
const F = createForm<{ text: string; literal: 'a' | 'b'; optional?: string; bool: boolean; mixed: boolean | 'mixed'; onlyTrue: true; number: number; nullable: number | null; tooWide: number | null | string; maybeNumber?: number }>();
const input = React.createRef<HTMLInputElement>();
const textarea = React.createRef<HTMLTextAreaElement>();
<F.Root defaultValues={{ text: '', bool: false, mixed: 'mixed', number: 30, nullable: null }} onSubmit={values => { values.nullable?.toFixed(); }}>
 <F.Textarea name="text" ref={textarea} onChange={(event, data) => { const native: HTMLTextAreaElement = event.currentTarget; data.value.toUpperCase(); void native; }} onBlur={event => { const native: HTMLTextAreaElement = event.currentTarget; void native; }} />
 <F.SearchBox name="optional" ref={input} onChange={(event, data) => { const native: HTMLInputElement | HTMLSpanElement = event.currentTarget; data.value.toUpperCase(); void native; }} />
 <F.Checkbox name="bool" ref={input} onChange={(event, data) => { const native: HTMLInputElement = event.currentTarget; const checked: boolean | 'mixed' = data.checked; void native; void checked; }} />
 <F.Checkbox name="mixed" triState ref={input} />
 <F.Switch name="bool" ref={input} onChange={(event, data) => { const native: HTMLInputElement = event.currentTarget; const checked: boolean = data.checked; void native; void checked; }} />
 <F.SpinButton name="nullable" ref={input} onChange={(event, data) => { const native: SpinButtonChangeEvent = event; const commit: SpinButtonOnChangeData = data; void native; void commit; }} onBlur={event => { const native: HTMLInputElement = event.currentTarget; void native; }} />
 <F.Slider name="number" min={20} max={40} ref={input} onChange={(event, data) => { const native: HTMLInputElement = event.currentTarget; data.value.toFixed(); void native; }} />
 <F.Slider name="maybeNumber" />
</F.Root>;
// @ts-expect-error nullable writer cannot target nonnullable number
<F.SpinButton name="number" />;
// @ts-expect-error reads cannot include unrelated strings
<F.SpinButton name="tooWide" />;
// @ts-expect-error Slider reads the complete number domain, not nullable values
<F.Slider name="nullable" />;
// @ts-expect-error boolean writer cannot target narrower true-only field
<F.Checkbox name="onlyTrue" />;
// @ts-expect-error mixed reading requires explicit triState
<F.Checkbox name="mixed" />;
// @ts-expect-error mixed writing cannot target boolean-only field
<F.Checkbox name="bool" triState />;
// @ts-expect-error Switch does not read mixed
<F.Switch name="mixed" />;
// @ts-expect-error text writer cannot target literals
<F.Textarea name="literal" />;
// @ts-expect-error SearchBox does not have a hidden codec
<F.SearchBox name="literal" codec={{}} />;
// @ts-expect-error direct Input deliberately has no literal codec
<F.Input name="literal" />;
// @ts-expect-error wrong native ref
<F.Textarea name="text" ref={input} />;
// @ts-expect-error SpinButton cannot expose an independent display draft
<F.SpinButton name="nullable" displayValue="12 units" />;
// @ts-expect-error RHF owns defaults
<F.Slider name="number" defaultValue={25} />;
const schema: StandardSchemaV1<{ choice: string }, { choice: 'a' | 'b' }> = { '~standard': { version: 1, vendor: 'fixture', validate: () => ({ value: { choice: 'a' } }) } };
const Schema = createForm(schema);
<Schema.Root defaultValues={{ choice: '' }} onSubmit={values => { const selected: 'a' | 'b' = values.choice; void selected; }}><Schema.Input name="choice" /></Schema.Root>;

const labelContent = React.createRef<HTMLSpanElement>();
const hintContent = React.createRef<HTMLDivElement>();
<F.Input name="text" label={{ as: 'span', ref: labelContent, 'data-label': 'slot', children: 'Literal **label**' }} hint={{ as: 'div', ref: hintContent, children: (Component, props) => <Component {...props}>Callback hint</Component> }} />;
<F.Textarea name="text" label={{ children: (Component, props) => <Component {...props}>Text</Component> }} />;
// @ts-expect-error native label cannot be replaced with another tag
<F.Input name="text" label={{ as: 'label', children: 'Wrong' }} />;
// @ts-expect-error raw HTML is excluded from the internal hint container
<F.Input name="text" hint={{ dangerouslySetInnerHTML: { __html: 'Wrong' } }} />;
// @ts-expect-error callbacks belong in slot children, not as direct values
<F.Input name="text" label={() => 'Wrong'} />;
// @ts-expect-error fixed span refs cannot point at SVG elements
<F.Input name="text" label={{ ref: React.createRef<SVGSVGElement>() }} />;
