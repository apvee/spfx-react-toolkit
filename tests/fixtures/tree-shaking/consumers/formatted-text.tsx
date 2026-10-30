import * as React from 'react';
import { FormattedText } from '__TOOLKIT__';
export function Example(): React.ReactElement {
  return <FormattedText as="p" maxLines={2}>Read **updates** in [docs](./docs).</FormattedText>;
}
