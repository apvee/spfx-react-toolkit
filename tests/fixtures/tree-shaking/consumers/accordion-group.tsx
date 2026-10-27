import * as React from 'react';
import { AccordionGroup } from '__TOOLKIT__';
export function Example(): React.ReactElement {
  return <AccordionGroup title="Preferences" defaultOpenItems={['profile']}>
    <AccordionGroup.Item value="profile" title="**Profile**" description="Account preferences">
      <input aria-label="Name" />
    </AccordionGroup.Item>
  </AccordionGroup>;
}
