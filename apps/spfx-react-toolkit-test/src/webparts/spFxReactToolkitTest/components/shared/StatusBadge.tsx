import * as React from 'react';
import { Badge } from '@fluentui/react-components';
import { DemoIcon } from './DemoIcon';

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
  <Badge appearance="tint" color={available ? 'success' : 'danger'} icon={<DemoIcon name={available ? 'Completed' : 'ErrorBadge'} />}>
    {label}
  </Badge>
);
