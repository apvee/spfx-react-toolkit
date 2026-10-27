import * as React from 'react';
import { renderContentSlot } from '../../packages/spfx-react-toolkit/src/controls/content-slot.internal';
import { useSx } from '../../packages/spfx-react-toolkit/src/hooks/useSx';
import * as margin from '../../packages/spfx-react-toolkit/src/helpers/styles/margin';

/** Real host fixture keeps the renderer and direction supplied by each Fluent/Griffel provider. */
export function ContentSlotHost({ name }: { name: string }): React.ReactElement {
  const sx = useSx();
  return <section>{renderContentSlot({ children: name, className: 'consumer', 'data-slot-host': name }, 'div', sx,
    { shorthand: 'preserve', className: sx(margin.medium) })}</section>;
}
