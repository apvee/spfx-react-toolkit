import * as React from 'react';
import { StateMessage } from '__TOOLKIT__';
export function Example(): React.ReactElement {
  return <StateMessage preset="empty" title="No results" description="Try **filters**."
    primaryAction={{ label: 'Retry', onClick: () => undefined }} />;
}
