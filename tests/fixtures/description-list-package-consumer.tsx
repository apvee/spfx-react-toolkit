import * as React from 'react';
import { DescriptionList as rootDescriptionList } from '@apvee/spfx-react-toolkit';
import { DescriptionList as domainDescriptionList } from '@apvee/spfx-react-toolkit/controls';
import { DescriptionList as controlDescriptionList } from '@apvee/spfx-react-toolkit/controls/description-list';
import type { DescriptionListProps as RootProps } from '@apvee/spfx-react-toolkit';
import type { DescriptionListProps as DomainProps } from '@apvee/spfx-react-toolkit/controls';
import type { DescriptionListProps as ControlProps, DescriptionListEntry, DescriptionListLayout,
  DescriptionListItemProps } from '@apvee/spfx-react-toolkit/controls/description-list';

export const aliases: typeof rootDescriptionList[] = [domainDescriptionList, controlDescriptionList];
const entries: readonly DescriptionListEntry[] = [{ id: 'owner', label: 'Owner', value: 'Ada' }];
const layout: DescriptionListLayout = 'auto';
export const props: RootProps = { items: entries, layout, 'aria-label': 'Details' };
export const domainProps: DomainProps = props;
export const controlProps: ControlProps = props;
export const itemProps: DescriptionListItemProps = { label: 'Owner', children: <strong>Ada</strong> };
const DescriptionList = controlDescriptionList;
export const element = <DescriptionList {...props} ref={React.createRef<HTMLDListElement>()} hidden />;
export const compound = <DescriptionList><DescriptionList.Item {...itemProps} ref={React.createRef<HTMLDivElement>()} /></DescriptionList>;
// @ts-expect-error Data and compound content are mutually exclusive.
<DescriptionList items={entries}><DescriptionList.Item label="Owner">Ada</DescriptionList.Item></DescriptionList>;
// @ts-expect-error Implementation paths remain private.
import {} from '@apvee/spfx-react-toolkit/controls/description-list/description-list';
