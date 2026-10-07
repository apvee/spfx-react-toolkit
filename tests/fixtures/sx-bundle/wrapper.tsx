import * as React from 'react';
import { useSx, width, typography } from '@apvee/spfx-react-toolkit';
export function Example(): React.ReactElement {
  const sx = useSx();
  return <div className={sx(width.full, typography.body1)}>Equivalent minimum UI</div>;
}
