import * as React from 'react';
import { createForm } from '../../packages/spfx-react-toolkit/src/controls/form';
const F = createForm<{ customer: { text: string; rows: { sku: string; quantity: number }[] }; rows: { text: string }[]; tags: string[] }>();
<F.ObjectSection name="customer" title="Customer">{({ Fields }) => <>
  <Fields.Input name="text" />
  <Fields.ArraySection name="rows" title="Rows" createItem={() => ({ sku: '', quantity: 1 })} renderItem={({ Fields }) => <Fields.Input name="sku" />} />
  {/* @ts-expect-error the relative record has no root rows text field */}
  <Fields.Input name="rows.0.text" />
</>}</F.ObjectSection>;
<F.ArraySection name="rows" aria-label="Rows" createItem={() => ({ text: '' })} renderItem={({ Fields }) => <Fields.Input name="text" />} />;
// @ts-expect-error sections require an accessible name
<F.ObjectSection name="customer">{() => undefined}</F.ObjectSection>;
// @ts-expect-error primitive arrays are not structured sections
<F.ArraySection name="tags" title="Tags" createItem={() => ''} renderItem={() => undefined} />;
// @ts-expect-error complete item model is required
<F.ArraySection name="rows" title="Rows" createItem={() => ({ wrong: '' })} renderItem={() => undefined} />;
// @ts-expect-error no copied source layout configuration
<F.ArraySection name="rows" title="Rows" createItem={() => ({ text: '' })} renderItem={() => undefined} itemLayout="horizontal" />;

<F.ObjectSection name="customer" title={{ children: (Component, props) => <Component {...props}>Customer</Component> }} description={{ ref: React.createRef<HTMLDivElement>(), children: 'Description' }}>{({ Fields }) => <Fields.Input name="text" label={{ children: 'Name' }} />}</F.ObjectSection>;
<F.ArraySection name="rows" title={{ ref: React.createRef<HTMLSpanElement>(), children: 'Rows' }} description={{ children: 'Entries' }} createItem={() => ({ text: '' })} renderItem={({ Fields }) => <Fields.Input name="text" />} />;
// @ts-expect-error the native legend cannot be replaced
<F.ObjectSection name="customer" title={{ as: 'legend', children: 'Customer' }}>{() => undefined}</F.ObjectSection>;
// @ts-expect-error outer accessible name remains required
<F.ObjectSection name="customer" description={{ children: 'Only a description' }}>{() => undefined}</F.ObjectSection>;
// @ts-expect-error raw HTML is forbidden in section content slots
<F.ObjectSection name="customer" title="Customer" description={{ dangerouslySetInnerHTML: { __html: 'Wrong' } }}>{() => undefined}</F.ObjectSection>;
