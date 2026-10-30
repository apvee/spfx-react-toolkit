import * as React from 'react';
import { DescriptionList } from '../../packages/spfx-react-toolkit/src/controls/description-list/description-list';
import type { DescriptionListEntry, DescriptionListLayout, DescriptionListProps, DescriptionListItemProps } from '../../packages/spfx-react-toolkit/src/controls/description-list/description-list.types';
const layout: DescriptionListLayout = 'auto';
const entries: readonly DescriptionListEntry[] = [{ id: 0, label: 0, value: 0 }, { id: 'empty', label: 'Empty', value: undefined }];
const props: DescriptionListProps = { layout, items: entries };
<DescriptionList {...props} ref={React.createRef<HTMLDListElement>()} title="Native" onClick={event => event.currentTarget.focus()} />;
const item: DescriptionListItemProps = { label: { as: 'div', ref: React.createRef<HTMLDivElement>(), className: 'label', children: 'Owner' }, children: null };
<DescriptionList><DescriptionList.Item {...item} ref={React.createRef<HTMLDivElement>()} />
  <DescriptionList.Item label="Callback">{{ children: (Component, contentProps) => <Component {...contentProps}>Rich</Component> }}</DescriptionList.Item>
</DescriptionList>;
// @ts-expect-error mutually exclusive content pathways
<DescriptionList items={entries}><DescriptionList.Item label="Duplicate">Value</DescriptionList.Item></DescriptionList>;
// @ts-expect-error bounded layout
<DescriptionList layout="rows" />;
// @ts-expect-error label is required
<DescriptionList.Item>Value</DescriptionList.Item>;
// @ts-expect-error null labels are unsupported
<DescriptionList.Item label={null}>Value</DescriptionList.Item>;
// @ts-expect-error root raw HTML pathway unsupported
<DescriptionList dangerouslySetInnerHTML={{ __html: 'raw' }} />;
// @ts-expect-error item raw HTML pathway unsupported
<DescriptionList.Item label="Safe" dangerouslySetInnerHTML={{ __html: 'raw' }}>Value</DescriptionList.Item>;
// @ts-expect-error slot raw HTML pathway unsupported
<DescriptionList.Item label={{ dangerouslySetInnerHTML: { __html: 'raw' } }}>Value</DescriptionList.Item>;
// @ts-expect-error value slot raw HTML pathway unsupported
<DescriptionList.Item label="Safe">{{ dangerouslySetInnerHTML: { __html: 'raw' } }}</DescriptionList.Item>;
// @ts-expect-error stable identity is required
const invalid: DescriptionListEntry = { label: 'Missing key', value: 'Value' };
void invalid;

declare const legacyNode: React.ReactNode;
const legacyValue: DescriptionListItemProps['children'] = legacyNode;
const conditionalValue: DescriptionListItemProps['children'] = false;
<DescriptionList.Item label="Conditional">{true}</DescriptionList.Item>;
// @ts-expect-error fixed tag preserves semantic inner div
<DescriptionList.Item label={{ as: 'h1', children: 'Heading' }}>Value</DescriptionList.Item>;
// @ts-expect-error fixed value tag
<DescriptionList.Item label="Safe">{{ as: 'span', children: 'Value' }}</DescriptionList.Item>;
void [legacyValue, conditionalValue];
