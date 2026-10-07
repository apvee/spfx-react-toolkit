import * as React from 'react';
import { useStableCallback } from '__TOOLKIT__';
export function Counter(): React.ReactElement {
  const [count, setCount] = React.useState(0);
  const increment = useStableCallback(() => setCount(count + 1));
  return <button onClick={increment}>{count}</button>;
}
