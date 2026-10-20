import * as React from 'react';
import { ContentCard } from '__TOOLKIT__';

export function Example(): React.ReactElement {
  return <ContentCard title="**Literal** summary" description={{ children: "Details" }} actions={<button>Open</button>}>Caller body</ContentCard>;
}
