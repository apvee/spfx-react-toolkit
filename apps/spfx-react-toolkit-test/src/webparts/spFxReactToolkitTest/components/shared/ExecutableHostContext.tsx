import * as React from 'react';
import type { ExecutableHostKind } from './selectExecutableGroupsForHost';

/** Set only by the actual private host adapter, never inferred from entry metadata. */
export const ExecutableHostContext = React.createContext<ExecutableHostKind | undefined>(undefined);
