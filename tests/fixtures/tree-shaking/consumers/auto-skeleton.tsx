import * as React from 'react';
import { AutoSkeleton } from '__TOOLKIT__';
export function Example(): React.ReactElement { return <AutoSkeleton loading><section><p>Profile</p><img alt="" data-skeleton-shape="circle" /></section></AutoSkeleton>; }
