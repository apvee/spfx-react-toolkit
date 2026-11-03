import * as React from 'react';
import { Badge } from '@fluentui/react-components';
import { Checkmark24Regular, ErrorCircle24Regular } from '@fluentui/react-icons';

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
  <Badge appearance="tint" color={available ? 'success' : 'danger'} icon={available ? <Checkmark24Regular aria-hidden="true" /> : <ErrorCircle24Regular aria-hidden="true" />}>
    {label}
  </Badge>
);
