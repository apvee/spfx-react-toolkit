import * as React from 'react';
import { SettingsGroup } from '__TOOLKIT__';
export function Example(): React.ReactElement {
  return <SettingsGroup title="Preferences" description="Your **settings**.">
    <SettingsGroup.Item title="Name" description="Visible name">
      <input aria-label="Name" />
    </SettingsGroup.Item>
  </SettingsGroup>;
}
