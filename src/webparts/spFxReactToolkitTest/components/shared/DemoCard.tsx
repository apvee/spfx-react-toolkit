import * as React from 'react';
import {
  Icon,
  Label,
  MessageBar,
  MessageBarType,
  Separator,
  Stack,
} from '@fluentui/react';

export interface DemoCardProps {
  readonly title: string;
  readonly iconName: string;
  readonly children: React.ReactNode;
  readonly description?: string;
  readonly error?: Error;
}

export const DemoCard: React.FC<DemoCardProps> = ({
  title,
  iconName,
  description,
  error,
  children,
}) => (
  <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: 16, border: '1px solid #edebe9', borderRadius: 4 } }}>
    <h3>
      <Icon iconName={iconName} style={{ marginRight: 8 }} />
      {title}
    </h3>
    <Separator />
    {description && <Label>{description}</Label>}
    {error && (
      <MessageBar messageBarType={MessageBarType.error}>
        {error.message}
      </MessageBar>
    )}
    {children}
  </Stack>
);
