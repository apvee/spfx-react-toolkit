import * as React from 'react';
import { useSx, width, typography, padding, gap } from '__TOOLKIT__';
const catalog = { full: width.full, body: typography.body1, padding: padding.medium, gap: gap.medium };
export function DynamicExample({ selection }: { selection: keyof typeof catalog }): React.ReactElement {
  return <div className={useSx()(catalog[selection])}>{selection}</div>;
}
