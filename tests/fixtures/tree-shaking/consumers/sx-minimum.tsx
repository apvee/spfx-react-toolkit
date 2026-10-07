import * as React from 'react';
import { useSx } from '__HOOK__';
import { width, typography } from '__DESCRIPTORS__';
export function Example(): React.ReactElement { return <div className={useSx()(width.full, typography.body1)}>Equivalent minimum UI</div>; }
