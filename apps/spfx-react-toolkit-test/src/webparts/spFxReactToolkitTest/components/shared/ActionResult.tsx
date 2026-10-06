import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react';

export interface ActionResultProps {
  readonly result: string | undefined;
  readonly error?: Error;
}

export const ActionResult: React.FC<ActionResultProps> = ({ result, error }) => {
  if (error) {
    return (
      <MessageBar messageBarType={MessageBarType.error}>
        {error.message}
      </MessageBar>
    );
  }

  if (!result) {
    return null;
  }

  return (
    <MessageBar messageBarType={MessageBarType.success}>
      {result}
    </MessageBar>
  );
};
