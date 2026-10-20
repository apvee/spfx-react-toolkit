import * as React from 'react';
import { FluentProvider } from '@fluentui/react-components';
import { RendererProvider } from '@griffel/react';
import { createDOMRenderer } from '@griffel/core';
import { useSPFxFluent9ThemeInfo } from '../../hooks/useSPFxFluent9ThemeInfo';
import { useSPFxContext } from '../../hooks/useSPFxContext';
import { AutoSkeleton } from '../../controls/auto-skeleton/auto-skeleton';
import type { PickerControlBoundaryProps } from '../../controls/pickers/picker-runtime-context.internal';
import { PickerControlBoundaryContext, PickerRuntimeContext } from '../../controls/pickers/picker-runtime-context.internal';
import type { PickerRequestCache } from '../../controls/pickers/picker-request-cache.internal';

/** Only initially unlabelled selections allocate measurement; the chosen tree stays mounted after hydration. */
function PropertyPanePickerControlBoundary({ pending, children }: PickerControlBoundaryProps): React.ReactElement {
  const [measureInitialSelection] = React.useState(pending);
  const control = <div data-skeleton-shape="rectangle">{children}</div>;
  return measureInitialSelection ? <AutoSkeleton data-property-pane-picker-control="" loading={pending}>{control}</AutoSkeleton> :
    <div data-property-pane-picker-control="" aria-busy={false}>{children}</div>;
}

/** Small native-host wrapper, intentionally without WebPartShell layout/title/property editing. */
export function PropertyPanePickerHost({ document, runtime, children }: { readonly document: Document; readonly runtime: PickerRequestCache; readonly children: React.ReactNode }): React.ReactElement {
  const renderer = React.useMemo(() => createDOMRenderer(document), [document]);
  const { theme } = useSPFxFluent9ThemeInfo(), { spfxContext } = useSPFxContext();
  const dir = spfxContext.pageContext.cultureInfo.isRightToLeft ? 'rtl' : 'ltr';
  return <RendererProvider renderer={renderer} targetDocument={document}>
    <FluentProvider theme={theme} dir={dir} targetDocument={document}>
      <PickerRuntimeContext.Provider value={runtime}>
        <PickerControlBoundaryContext.Provider value={PropertyPanePickerControlBoundary}>{children}</PickerControlBoundaryContext.Provider>
      </PickerRuntimeContext.Provider>
    </FluentProvider>
  </RendererProvider>;
}
