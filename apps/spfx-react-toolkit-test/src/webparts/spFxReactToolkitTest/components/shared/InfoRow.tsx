import * as React from 'react';
import { Label, Icon } from '@fluentui/react';
import styles from '../SpFxReactToolkitTest.module.scss';

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
export const InfoRow: React.FC<IInfoRowProps> = ({ label, value, icon }) => (
  <div className={styles.infoRow}>
    {icon && <Icon iconName={icon} />}
    <Label>{label}:</Label>
    <span>{value || 'N/A'}</span>
  </div>
);
