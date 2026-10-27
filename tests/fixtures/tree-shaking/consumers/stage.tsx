import * as React from 'react';
import { Stage } from '__TOOLKIT__';

export function Example(): React.ReactElement {
  return <Stage presentation="dialog" size="small" trigger={<button type="button">Edit profile</button>}>
    <Stage.Header title="Profile" description="**Literal** details" actions={<button type="button">Refresh</button>} />
    <Stage.Body><input aria-label="Display name" /></Stage.Body>
    <Stage.Footer><button type="button">Save profile</button></Stage.Footer>
  </Stage>;
}
