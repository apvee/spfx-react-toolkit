import * as React from 'react';
import { SettingsGroup } from '../../packages/spfx-react-toolkit/src/controls/settings-group/settings-group';
import { AccordionGroup } from '../../packages/spfx-react-toolkit/src/controls/accordion-group/accordion-group';
import type { AccordionGroupProps } from '../../packages/spfx-react-toolkit/src/controls/accordion-group/accordion-group.types';
const toggle: AccordionGroupProps['onToggle'] = (event, data) => {
  event.preventDefault();
  const values: Array<string | number> = data.openItems;
  void values;
};
<SettingsGroup title="**Settings**" headingLevel={2} ref={React.createRef<HTMLDivElement>()} onClick={event => event.currentTarget.focus()}>
  <SettingsGroup.Item description={<em>Caller JSX</em>} ref={React.createRef<HTMLDivElement>()}><input aria-label="Control" /></SettingsGroup.Item>
</SettingsGroup>;
<AccordionGroup openItems={['a', 1]} defaultOpenItems="a" multiple onToggle={toggle}>
  <AccordionGroup.Item value={1} headingLevel={6} disabled unmountOnClose ref={React.createRef<HTMLDivElement>()} />
</AccordionGroup>;
// @ts-expect-error value is required
<AccordionGroup.Item title="Missing value" />;
// @ts-expect-error heading is bounded
<SettingsGroup headingLevel={7} />;
// @ts-expect-error heading is bounded
<AccordionGroup.Item value="a" headingLevel={0} />;
// @ts-expect-error appearance is bounded
<SettingsGroup appearance="outline" />;
// @ts-expect-error owned HTML pathway unsupported
<SettingsGroup dangerouslySetInnerHTML={{ __html: 'raw' }} />;
// @ts-expect-error owned item HTML pathway unsupported
<AccordionGroup.Item value="a" dangerouslySetInnerHTML={{ __html: 'raw' }} />;
// @ts-expect-error no arbitrary slots API
<AccordionGroup slots={{ title: 'slot' }} />;
// @ts-expect-error no invented layout API
<SettingsGroup.Item layout={{ direction: 'row' }} />;
