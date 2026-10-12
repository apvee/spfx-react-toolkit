import * as React from 'react';
import { DescriptionList } from '__TOOLKIT__';

export function Example(): React.ReactElement {
  return <DescriptionList items={[{ id: 'owner', label: 'Owner', value: 'Ada' }]} />;
}
