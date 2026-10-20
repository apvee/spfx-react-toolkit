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
<SettingsGroup appearance="filled" />;
<SettingsGroup appearance="alternate" />;
<SettingsGroup appearance="outline" />;
<AccordionGroup appearance="filled" />;
<AccordionGroup appearance="alternate" />;
<AccordionGroup appearance="outline" />;
// @ts-expect-error transparent is no longer a group appearance
<SettingsGroup appearance="transparent" />;
// @ts-expect-error transparent is no longer a group appearance
<AccordionGroup appearance="transparent" />;
// @ts-expect-error use the public alternate name
<SettingsGroup appearance="filled-alternative" />;
// @ts-expect-error use the public alternate name
<AccordionGroup appearance="filled-alternative" />;
// @ts-expect-error owned HTML pathway unsupported
<SettingsGroup dangerouslySetInnerHTML={{ __html: 'raw' }} />;
// @ts-expect-error owned item HTML pathway unsupported
<AccordionGroup.Item value="a" dangerouslySetInnerHTML={{ __html: 'raw' }} />;
// @ts-expect-error no arbitrary slots API
<AccordionGroup slots={{ title: 'slot' }} />;
// @ts-expect-error no invented layout API
<SettingsGroup.Item layout={{ direction: 'row' }} />;

const conditionalSlot: React.ReactNode = false;
<SettingsGroup title={conditionalSlot} description={0} icon={{ children: <span>Icon</span>, ref: React.createRef<HTMLSpanElement>() }} />;
<SettingsGroup title={{ as: 'span', children: (Tag, props) => <Tag {...props}>Callback title</Tag> }} />;
// @ts-expect-error named title slots cannot replace the semantic heading
<SettingsGroup title={{ as: 'h1', children: 'Title' }} />;
// @ts-expect-error raw HTML is forbidden in named slots
<SettingsGroup title={{ dangerouslySetInnerHTML: { __html: 'raw' } }} />;
// @ts-expect-error direct render functions are not slot shorthand
<SettingsGroup title={() => <span>Direct callback</span>} />;

<SettingsGroup title="Icon sizing" icon={{ children: (Tag, props) => <Tag {...props}><svg width={80} height={80} /></Tag> }} description={{ children: <> {false} </> }} />;
