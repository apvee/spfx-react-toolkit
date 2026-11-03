import { makeStyles, tokens } from '@fluentui/react-components';
import { useDemoStyles as useHostDemoStyles } from '../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/demoStyles';
const useComparisonStyles = makeStyles({
  panel: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, minWidth: 0 },
  scroll: { overflowX: 'auto', minWidth: 0, maxWidth: '100%', borderRadius: tokens.borderRadiusMedium },
});
/** Additional comparison-only presentation, outside the selectable sample. */
export function useDemoStyles(): ReturnType<typeof useHostDemoStyles> & ReturnType<typeof useComparisonStyles> {
  return { ...useHostDemoStyles(), ...useComparisonStyles() };
}
