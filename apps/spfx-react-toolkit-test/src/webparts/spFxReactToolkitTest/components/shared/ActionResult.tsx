import * as React from 'react';
import { MessageBar, MessageBarBody } from '@fluentui/react-components';

export interface ActionResultProps {
  readonly result: string | undefined;
  readonly error?: Error;
}

export const ActionResult: React.FC<ActionResultProps> = ({ result, error }) => {
  if (error) {
    return (
      <MessageBar intent="error">
        <MessageBarBody>{error.message}</MessageBarBody>
      </MessageBar>
    );
  }

  if (!result) {
    return <></>;
  }

  return (
    <MessageBar intent="success">
      <MessageBarBody>{result}</MessageBarBody>
    </MessageBar>
  );
};
