import * as React from 'react';
import { ColorSwatch, SwatchPicker, Tooltip } from '@fluentui/react-components';
import { ThemeContext_unstable } from '@fluentui/react-shared-contexts';
import { tokens } from '@fluentui/react-theme';
import { makeStyles } from '@griffel/react';
import type { WebPartShellStrings } from '../localization/web-part-shell-strings.internal';
import { getTitleColorCss, getTitlePickerColors, isTitleColor } from './web-part-title-format.internal';
import type { TitleColor } from './web-part-title.types.internal';

const useStyles = makeStyles({
  root: { maxInlineSize: 'min(280px, calc(100vw - 40px))', display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS },
  picker: { display: 'flex', flexDirection: 'column', alignItems: 'stretch', gap: tokens.spacingVerticalS },
  section: { display: 'flex', flexDirection: 'column', minInlineSize: 0 },
  group: { display: 'flex', flexWrap: 'wrap', gap: tokens.spacingHorizontalS, padding: tokens.spacingVerticalXS },
  label: { fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold },
});
interface ColorPickerProps {
  strings: WebPartShellStrings;
  selected: TitleColor | 'mixed' | undefined;
  onSelect: (color: TitleColor | undefined) => void;
}

/** Private picker shared by the main toolbar and its overflow command. */
export function WebPartTitleColorPicker({ strings, selected, onSelect }: ColorPickerProps): React.ReactElement {
  const styles = useStyles();
  const theme = React.useContext(ThemeContext_unstable);
  const colors = getTitlePickerColors(selected, theme);
  const swatch = (id: TitleColor | 'automatic', name: string, color: string): React.ReactElement =>
    <Tooltip key={id} content={name} relationship="description">
      <ColorSwatch data-title-color={id} value={id} aria-label={name} color={color} />
    </Tooltip>;
  return <div className={styles.root} data-title-color-picker="">
    <SwatchPicker className={styles.picker} aria-label={selected === 'mixed' ? strings.mixedFormatting(strings.titleThemeColors) : strings.titleThemeColors}
      shape="circular" size="medium" selectedValue={selected === 'mixed' ? '' : selected ?? 'automatic'}
      onSelectionChange={(_event, data) => {
        if (data.selectedValue === 'automatic') onSelect(undefined);
        else if (isTitleColor(data.selectedValue)) onSelect(data.selectedValue);
      }}>
      <div className={styles.section} role="group" aria-label={strings.automatic} data-title-color-section="automatic">
        <div className={styles.label}>{strings.automatic}</div>
        <div className={styles.group} data-title-color-row="automatic">{swatch('automatic', strings.automatic, tokens.colorNeutralForeground1)}</div>
      </div>
      {(['theme', 'palette'] as const).map(group => <div key={group} className={styles.section} role="group" aria-label={group === 'theme' ? strings.themeColors : strings.paletteColors} data-title-color-section={group}>
        <div className={styles.label}>{group === 'theme' ? strings.themeColors : strings.paletteColors}</div>
        <div className={styles.group} data-title-color-row={group}>{colors.filter(item => item.group === group)
          .map(item => swatch(item.id, strings.colorNames[item.id], getTitleColorCss(item.id)))}</div>
      </div>)}
    </SwatchPicker>
  </div>;
}
