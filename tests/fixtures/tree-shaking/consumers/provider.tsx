import * as React from 'react';
import { SPFxWebPartProvider } from '__TOOLKIT__';
import { useSPFxContext } from '__CONTEXT_HOOK__';
function Child(): React.ReactElement { return <div>{useSPFxContext().instanceId}</div>; }
export function ProviderExample(props: React.ComponentProps<typeof SPFxWebPartProvider>): React.ReactElement {
  return <SPFxWebPartProvider {...props}><Child /></SPFxWebPartProvider>;
}
