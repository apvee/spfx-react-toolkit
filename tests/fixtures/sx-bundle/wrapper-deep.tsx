import * as React from 'react';
import { useSx } from '@apvee/spfx-react-toolkit/lib/hooks/useSx';
import { width, typography } from '@apvee/spfx-react-toolkit/lib/helpers/styles';
export function Example(): React.ReactElement {
  const sx = useSx();
  return <div className={sx(width.full, typography.body1)}>Equivalent minimum UI</div>;
}
