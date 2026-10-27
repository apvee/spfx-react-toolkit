import * as React from 'react';
// The callback closes over a namespace; make its used exports explicit to Webpack.
const LazyExample = React.lazy(() => import(/* webpackExports: ["useSx", "width", "typography"] */ '__TOOLKIT__').then(toolkit => ({ default: function LazyFeature() {
  return <div className={toolkit.useSx()(toolkit.width.full, toolkit.typography.body1)}>Lazy minimum UI</div>;
} })));
export function LazyApp(): React.ReactElement { return <React.Suspense fallback={<div>Loading</div>}><LazyExample /></React.Suspense>; }
