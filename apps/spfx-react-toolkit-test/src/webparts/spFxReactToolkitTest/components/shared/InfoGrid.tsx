import * as React from 'react';
import { useDemoStyles } from './demoStyles';
import { InfoRow } from './InfoRow';

export interface InfoGridRow {
  readonly label: string;
  readonly value: string | undefined;
  readonly icon?: string;
}

export interface InfoGridProps {
  readonly rows: readonly InfoGridRow[];
}

export const InfoGrid: React.FC<InfoGridProps> = ({ rows }) => {
  const styles = useDemoStyles();
  return <div className={styles.infoGrid}>
    {rows.map(row => <InfoRow key={`${row.label}:${row.icon ?? ''}`} label={row.label} value={row.value} icon={row.icon} />)}
  </div>;
};
