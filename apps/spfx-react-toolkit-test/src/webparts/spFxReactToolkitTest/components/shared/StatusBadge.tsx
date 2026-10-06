import * as React from 'react';
import { Icon } from '@fluentui/react';
import styles from '../SpFxReactToolkitTest.module.scss';

/**
 * Props for the StatusBadge component
 */
export interface IStatusBadgeProps {
  /** Whether the status is available/active */
  available: boolean;
  /** Label text to display */
  label: string;
}

/**
 * Helper component for displaying status badges
 * with success/error styling based on availability
 */
export const StatusBadge: React.FC<IStatusBadgeProps> = ({ available, label }) => (
  <div className={`${styles.statusBadge} ${available ? styles.success : styles.error}`}>
    <Icon iconName={available ? 'Completed' : 'ErrorBadge'} />
    {label}
  </div>
);
