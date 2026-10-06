import * as React from 'react';
import styles from '../SpFxReactToolkitTest.module.scss';

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

export const JsonDetails: React.FC<JsonDetailsProps> = ({ label, value }) => (
  <details className={styles.detailsSection}>
    <summary>{label}</summary>
    <pre>{stringifySafe(value)}</pre>
  </details>
);
