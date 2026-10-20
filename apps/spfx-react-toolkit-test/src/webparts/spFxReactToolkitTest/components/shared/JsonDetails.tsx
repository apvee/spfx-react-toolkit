import * as React from 'react';
import { Accordion, AccordionItem, AccordionHeader } from '@fluentui/react-components';
import { DemoAccordionPanel } from './DemoAccordionPanel';
import { useDemoStyles } from './demoStyles';

export interface JsonDetailsProps {
  readonly label: string;
  readonly value: unknown;
}

function stringifySafe(value: unknown): string {
  const seen = new WeakSet<object>();

  return JSON.stringify(value, (key, item) => {
    if (typeof key === 'string' && (key.indexOf('_') === 0 || key === 'serviceScope' || key === 'service')) {
      return '[Private]';
    }

    if (typeof item === 'object' && item !== null) {
      if (seen.has(item)) {
        return '[Circular]';
      }

      seen.add(item);
    }

    return item;
  }, 2);
}

export const JsonDetails: React.FC<JsonDetailsProps> = ({ label, value }) => {
  const styles = useDemoStyles();
  return <Accordion collapsible>
    <AccordionItem value="json">
      <AccordionHeader>{label}</AccordionHeader>
      <DemoAccordionPanel>
        <div className={styles.scroll} role="region" aria-label={`${label} JSON`} tabIndex={0}>
          <pre className={styles.json}>{stringifySafe(value)}</pre>
        </div>
      </DemoAccordionPanel>
    </AccordionItem>
  </Accordion>;
};
