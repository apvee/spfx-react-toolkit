import * as React from 'react';
import { Pagination } from '__TOOLKIT__';

export function Example(): React.ReactElement {
  return <Pagination pagination={{ type: 'pages', pageIndex: 0, pageSize: 10, totalItems: 42 }}
    onPaginationChange={(_event, data) => { console.log(data); }} />;
}
