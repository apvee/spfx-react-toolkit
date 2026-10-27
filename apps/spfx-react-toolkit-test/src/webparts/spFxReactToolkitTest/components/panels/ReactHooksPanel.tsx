import { useDemoStyles } from '../shared/demoStyles';
import { Button } from '@fluentui/react-components';
import * as React from 'react';
import { useStableCallback } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';
const ReactHooksPanel: React.FC = () => {
  const demoStyles = useDemoStyles();
  const [count, setCount] = React.useState(0);
  const [observed, setObserved] = React.useState<number>();
  const [sameIdentity, setSameIdentity] = React.useState<boolean>();
  const originalCallback = React.useRef<(() => number) | undefined>();
  const callback = useStableCallback(() => {
    setObserved(count);
    return count;
  });
  React.useEffect(() => {
    if (!originalCallback.current)
      originalCallback.current = callback;
  }, [callback]);
  const invokeOriginal = (): void => {
    const original = originalCallback.current;
    if (original) {
      setSameIdentity(original === callback);
      original();
    }
  };
  return (<DemoCard contentPlacement="configuration" title="useStableCallback" iconName="Code" description="Increment the counter and invoke the captured callback: it keeps the same identity and reads the updated value." commands={<>
    <div className={demoStyles.actions}>
      <Button onClick={() => setCount(current => current + 1)}>{"Increment counter"}</Button>
      <Button onClick={invokeOriginal}>{"Invoke captured callback"}</Button>
    </div>
  </>}>
    <div className={demoStyles.panel}>
      <p>Current counter: {count}</p>
      <div aria-live="polite">
        <p>Observed counter: {observed === undefined ? 'not invoked' : observed}</p>
        <p>Same callback identity: {sameIdentity === undefined ? 'not checked' : sameIdentity ? 'yes' : 'no'}</p>
      </div>
    </div>
  </DemoCard>);
};
export default ReactHooksPanel;
