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
