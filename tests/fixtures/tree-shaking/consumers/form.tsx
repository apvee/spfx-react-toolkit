import * as React from 'react';
import { Form } from '__TOOLKIT__';
// Public compound identity intentionally exposes the complete adapter family.
export const family = Form;
export function Example(): React.ReactElement {
  return <Form.Root defaultValues={{ title: '' }} onSubmit={value => { console.log(value.title); }}>
    <Form.Input name="title" label="Title" /><Form.Submit>Save</Form.Submit>
  </Form.Root>;
}
