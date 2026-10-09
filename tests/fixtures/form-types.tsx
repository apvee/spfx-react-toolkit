import * as React from 'react';
import { createForm, Form } from '../../packages/spfx-react-toolkit/src/controls/form';
import type { StandardSchemaV1 } from '@standard-schema/spec';
const F = createForm<{ text: string; literal: 'a' | 'b'; count: number; optional?: string; maybe: string | undefined }>();
<F.Root defaultValues={{ text: '' }} onSubmit={values => { values.text.toUpperCase(); }}><F.Input name="text" /><F.Input name="optional" /></F.Root>;
// @ts-expect-error Input cannot write arbitrary string to a literal domain
<F.Input name="literal" />;
// @ts-expect-error numeric fields need numeric controls
<F.Input name="count" />;
// @ts-expect-error nonexistent path
<F.Input name="missing" />;
// @ts-expect-error Input owns its value
<F.Input name="text" value="bad" />;
// @ts-expect-error Input has no codec escape hatch
<F.Input name="text" codec={{}} />;
// @ts-expect-error methods cannot be combined with initialization
<F.Root methods={F.useForm()} defaultValues={{}} onSubmit={() => {}} />;
// @ts-expect-error owned form children cannot be replaced with raw HTML
<F.Root dangerouslySetInnerHTML={{ __html: 'x' }} onSubmit={() => {}} />;
const schema: StandardSchemaV1<{ text: string }, { size: number }> = { '~standard': { version: 1, vendor: 'fixture', validate: value => ({ value: { size: (value as { text: string }).text.length } }) } };
const ExplicitSchema = createForm<typeof schema>(schema);
<ExplicitSchema.Root onSubmit={value => { value.size.toFixed(); }} />;
const Transformed = createForm(schema);
<Transformed.Root defaultValues={{ text: '' }} onSubmit={values => { values.size.toFixed();
  // @ts-expect-error submitted output is not input
  values.text;
}}><Transformed.Input name="text" /></Transformed.Root>;
// @ts-expect-error bound validation cannot be replaced
Transformed.useForm({ resolver: async () => ({ values: { size: 1 }, errors: {} }) });
const methods = Transformed.useForm(); methods.getValues().text.toUpperCase(); methods.formState.errors.__formValidation?.message;
<Form.Root onSubmit={() => {}}><Form.Input name="dynamic" /></Form.Root>;
declare const scalarSchema: StandardSchemaV1<string, number>;
// @ts-expect-error Form initialization requires a record-shaped input model
createForm(scalarSchema);
void React;
