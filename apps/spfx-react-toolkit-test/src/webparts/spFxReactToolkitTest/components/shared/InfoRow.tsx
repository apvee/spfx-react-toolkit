import * as React from 'react';
import { Text } from '@fluentui/react-components';
import { DemoIcon } from './DemoIcon';
import { useDemoStyles } from './demoStyles';

/**
 * Props for the InfoRow component
 */
export interface IInfoRowProps {
  /** Label text to display */
  label: string;
  /** Value to display (shows 'N/A' if undefined) */
  value: string | undefined;
  /** Optional Fluent UI icon name */
  icon?: string;
}

/**
 * Helper component for displaying labeled information rows
 * with optional icon prefix
 */
export const InfoRow: React.FC<IInfoRowProps> = ({ label, value, icon }) => {
  const styles = useDemoStyles();
  return <div className={styles.infoRow}>
    <span className={styles.infoLabel}>
      {icon && <DemoIcon name={icon} />}
      <Text weight="semibold">{label}:</Text>
    </span>
    <Text>{value || 'N/A'}</Text>
  </div>;
};
