import * as React from 'react';
import { Accordion, AccordionHeader, AccordionItem, Card, Divider, MessageBar, MessageBarBody, Text } from '@fluentui/react-components';
import { DemoAccordionPanel } from './DemoAccordionPanel';
import { DemoIcon } from './DemoIcon';
import { useDemoStyles } from './demoStyles';

export interface DemoCardProps {
  readonly title: string;
  readonly iconName: string;
  readonly children?: React.ReactNode;
  readonly description?: string;
  readonly error?: Error;
  /** Configuration stays mounted and reachable independently of the preview state. */
  readonly controls?: React.ReactNode;
  /** Commands that drive the example, separate from the live control's own actions. */
  readonly commands?: React.ReactNode;
  /** Compact, always-visible outcome or status. */
  readonly status?: React.ReactNode;
  /** Detailed refs, callbacks, raw payloads and explanations; initially collapsed. */
  readonly diagnostics?: React.ReactNode;
  readonly diagnosticsLabel?: string;
  /** Nonvisual hook/service results can stay in the configuration card. */
  readonly contentPlacement?: 'subject' | 'configuration';
}

/** Shared presentation: outline tools card followed by the undecorated live subject. */
export const DemoCard: React.FC<DemoCardProps> = ({ title, iconName, description, error, controls, commands, status, diagnostics, diagnosticsLabel = 'Advanced diagnostics', contentPlacement = 'subject', children }) => {
  const styles = useDemoStyles();
  return (
    <section className={styles.section} aria-label={title} data-demo-section="">
      <div className={styles.sectionHeader}>
        <Text as="h3" size={500} weight="semibold" className={styles.heading}>{title}</Text>
        {description && <Text as="p" className={styles.description}>{description}</Text>}
      </div>
      <Card appearance="outline" className={styles.card} data-demo-configuration="">
        <div className={styles.configTitle}><DemoIcon name={iconName} /><Text weight="semibold">{contentPlacement === 'configuration' ? 'Configuration and results' : 'Configuration and simulation'}</Text></div>
        {controls !== undefined && <div className={styles.controls} role="group" aria-label={`${title} configuration`}>{controls}</div>}
        {commands !== undefined && <div className={styles.actions} role="group" aria-label={`${title} commands`}>{commands}</div>}
        {error && <MessageBar intent="error"><MessageBarBody>{error.message}</MessageBarBody></MessageBar>}
        {status !== undefined && <div className={styles.status} role="group" aria-label={`${title} status`}>{status}</div>}
        {contentPlacement === 'configuration' && children !== undefined && <div className={styles.results} data-demo-result="" role="group" aria-label={`${title} results`}>{children}</div>}
        {diagnostics !== undefined && <>
          <Divider />
          <Accordion collapsible><AccordionItem value="diagnostics"><AccordionHeader>{diagnosticsLabel}</AccordionHeader>
            <DemoAccordionPanel><div className={styles.observations} role="group" aria-label={`${title} diagnostics`}>{diagnostics}</div></DemoAccordionPanel>
          </AccordionItem></Accordion>
        </>}
      </Card>
      {contentPlacement === 'subject' && children !== undefined && <div className={styles.preview} data-demo-subject="" role="group" aria-label={`${title} preview`}>{children}</div>}
    </section>
  );
};
