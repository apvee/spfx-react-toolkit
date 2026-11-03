import * as React from 'react';
import { makeStyles } from '@fluentui/react-components';
import { Checkmark24Regular, Info24Regular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  inline: { fontSize: 'inherit', width: '1em', height: '1em', verticalAlign: 'middle', flexShrink: 0 },
});

/** Decorative icons follow the surrounding text or Fluent icon slot size; unknown names use an information symbol. */
export const DemoIcon = ({ name }: { readonly name: string }): JSX.Element => {
  const styles = useStyles();
  const IconComponent = name === 'CompletedSolid' ? Checkmark24Regular : Info24Regular;
  // Owning toolkit icon slots size their direct SVG child (24px header/20px item).
  // Keep that native boundary so the host class, rather than a wrapper, owns sizing.
  return <IconComponent className={styles.inline} aria-hidden="true" />;
};
