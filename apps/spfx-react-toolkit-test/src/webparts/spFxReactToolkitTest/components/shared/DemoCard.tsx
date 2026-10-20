import * as React from 'react';
import {
  Icon,
  Label,
  MessageBar,
  MessageBarType,
  Separator,
  Stack,
} from '@fluentui/react';
import { Accordion, AccordionHeader, AccordionItem } from '@fluentui/react-components';
import { DemoAccordionPanel } from './DemoAccordionPanel';

export interface DemoCardProps {
  readonly title: string;
  readonly iconName: string;
  readonly children?: React.ReactNode;
  readonly controls?: React.ReactNode;
  readonly commands?: React.ReactNode;
  readonly status?: React.ReactNode;
  readonly diagnostics?: React.ReactNode;
  readonly diagnosticsLabel?: string;
  readonly contentPlacement?: 'subject' | 'configuration';
  readonly description?: string;
  readonly error?: Error;
}

export const DemoCard: React.FC<DemoCardProps> = ({
  title,
  iconName,
  description,
  error,
  children, controls, commands, status, diagnostics, diagnosticsLabel = 'Advanced diagnostics', contentPlacement = 'subject',
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
    {controls !== undefined && <div role="group" aria-label={`${title} configuration`}>{controls}</div>}
    {commands !== undefined && <div role="group" aria-label={`${title} commands`}>{commands}</div>}
    {status}
    {diagnostics !== undefined && <Accordion collapsible><AccordionItem value="diagnostics"><AccordionHeader>{diagnosticsLabel}</AccordionHeader><DemoAccordionPanel>{diagnostics}</DemoAccordionPanel></AccordionItem></Accordion>}
    {controls !== undefined ? <div role="group" aria-label={`${title} ${contentPlacement === 'configuration' ? 'results' : 'preview'}`}>{children}</div> : children}
  </Stack>
);
