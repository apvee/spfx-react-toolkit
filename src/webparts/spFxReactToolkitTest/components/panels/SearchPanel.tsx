import * as React from 'react';
import { Label, MessageBar, MessageBarType, Stack } from '@fluentui/react';
import { SearchVerticals } from '../../../../hooks';
import {
  PnPSearchAdvancedDemo,
  PnPSearchBasicDemo,
  PnPSearchRefinersDemo,
  PnPSearchSuggestionsDemo,
} from '../demos';
import { DemoCard, InfoGrid } from '../shared';

const SearchPanel: React.FC = () => (
  <Stack tokens={{ childrenGap: 16 }}>
    <MessageBar messageBarType={MessageBarType.info}>
      Search demos cover text queries, builder queries, vertical source ids, refiners, paging, and suggestions.
    </MessageBar>

    <DemoCard title="Search Verticals" iconName="SearchAndApps">
      <Label>Built-in source ids are exported by SearchVerticals.</Label>
      <InfoGrid
        rows={Object.keys(SearchVerticals).map(key => ({
          label: key,
          value: SearchVerticals[key as keyof typeof SearchVerticals],
          icon: 'Search',
        }))}
      />
    </DemoCard>

    <PnPSearchBasicDemo />
    <PnPSearchAdvancedDemo />
    <PnPSearchRefinersDemo />
    <PnPSearchSuggestionsDemo />
  </Stack>
);

export default SearchPanel;
