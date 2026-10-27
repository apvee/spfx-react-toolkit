import * as React from 'react';
import { EditableText } from '__TOOLKIT__';
export function Example(): React.ReactElement {
  return <EditableText value="<h2>Standalone summary</h2>" allowedBlocks={['h2', 'p']} />;
}
