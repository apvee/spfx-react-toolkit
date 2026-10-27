import * as React from 'react';
import { createForm } from '__TOOLKIT__';
import type { StandardSchemaV1 } from '@standard-schema/spec';
const schema: StandardSchemaV1<{ title: string }, { length: number }> = {
  '~standard': { version: 1, vendor: 'bundle-fixture', validate: value => ({ value: { length: (value as { title: string }).title.length } }) }
};
// Measuring the exported suite avoids promising per-slot pruning.
export const family = createForm(schema);
export function Example(): React.ReactElement {
  return <family.Root defaultValues={{ title: '' }} onSubmit={value => { console.log(value.length); }}>
    <family.Input name="title" label="Title" /><family.Submit>Save</family.Submit>
  </family.Root>;
}
