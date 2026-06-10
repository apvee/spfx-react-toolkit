import * as React from 'react';
import { Stack } from '@fluentui/react';
import { InfoRow } from './InfoRow';

export interface InfoGridRow {
  readonly label: string;
  readonly value: string | undefined;
  readonly icon?: string;
}

export interface InfoGridProps {
  readonly rows: readonly InfoGridRow[];
}

export const InfoGrid: React.FC<InfoGridProps> = ({ rows }) => (
  <Stack tokens={{ childrenGap: 2 }}>
    {rows.map(row => (
      <InfoRow
        key={`${row.label}:${row.icon ?? ''}`}
        label={row.label}
        value={row.value}
        icon={row.icon}
      />
    ))}
  </Stack>
);
