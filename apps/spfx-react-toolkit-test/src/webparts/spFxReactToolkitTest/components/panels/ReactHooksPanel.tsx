import * as React from 'react';
import { DefaultButton, Stack } from '@fluentui/react';
import { useStableCallback } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';

const ReactHooksPanel: React.FC = () => {
  const [count, setCount] = React.useState(0);
  const [observed, setObserved] = React.useState<number>();
  const [sameIdentity, setSameIdentity] = React.useState<boolean>();
  const originalCallback = React.useRef<(() => number) | undefined>();

  const callback = useStableCallback(() => {
    setObserved(count);
    return count;
  });

  React.useEffect(() => {
    if (!originalCallback.current) originalCallback.current = callback;
  }, [callback]);

  const invokeOriginal = (): void => {
    const original = originalCallback.current;
    if (original) {
      setSameIdentity(original === callback);
      original();
    }
  };

  return (
    <DemoCard
      title="useStableCallback"
      iconName="Code"
      description="Increment the counter, then invoke the callback captured at mount. It reads the updated counter while keeping the same identity. No SPFx or Fluent provider is required by this hook."
    >
      <Stack tokens={{ childrenGap: 12 }}>
        <p>Current counter: {count}</p>
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Increment counter" onClick={() => setCount(current => current + 1)} />
          <DefaultButton text="Invoke original callback" onClick={invokeOriginal} />
        </Stack>
        <div aria-live="polite">
          <p>Observed counter: {observed === undefined ? 'not invoked' : observed}</p>
          <p>Same callback identity: {sameIdentity === undefined ? 'not checked' : sameIdentity ? 'yes' : 'no'}</p>
        </div>
      </Stack>
    </DemoCard>
  );
};

export default ReactHooksPanel;
