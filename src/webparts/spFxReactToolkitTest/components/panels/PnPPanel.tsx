import * as React from 'react';
import { MessageBar, MessageBarType, Stack } from '@fluentui/react';
import {
  PnPContextDemo,
  PnPListDemo,
  PnPOperationsDemo,
} from '../demos';

const PnPPanel: React.FC = () => (
  <Stack tokens={{ childrenGap: 16 }}>
    <MessageBar messageBarType={MessageBarType.info}>
      These demos cover PnPjs context creation, state-managed invoke/batch operations, and list query/CRUD helpers.
    </MessageBar>
    <PnPContextDemo />
    <PnPOperationsDemo />
    <PnPListDemo />
  </Stack>
);

export default PnPPanel;
