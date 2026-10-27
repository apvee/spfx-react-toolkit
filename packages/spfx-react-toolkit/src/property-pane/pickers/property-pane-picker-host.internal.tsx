import * as React from 'react';
import { FluentProvider } from '@fluentui/react-components';
import { RendererProvider } from '@griffel/react';
import { createDOMRenderer } from '@griffel/core';
import { usePropertyPanePickerTheme } from './usePropertyPanePickerTheme.internal';
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

/** Native-only bridge: reads pane tokens outside this React root and carries them into body portals.
 * Retains the owning document renderer, SPFx direction and mounted picker runtime. */
export function PropertyPanePickerHost({ container, runtime, children }: { readonly container: HTMLElement; readonly runtime: PickerRequestCache; readonly children: React.ReactNode }): React.ReactElement {
  const document = container.ownerDocument;
  const renderer = React.useMemo(() => createDOMRenderer(document), [document]);
  const theme = usePropertyPanePickerTheme(container), { spfxContext } = useSPFxContext();
  const dir = spfxContext.pageContext.cultureInfo.isRightToLeft ? 'rtl' : 'ltr';
  return <RendererProvider renderer={renderer} targetDocument={document}>
    <FluentProvider theme={theme} dir={dir} targetDocument={document}>
      <PickerRuntimeContext.Provider value={runtime}>
        <PickerControlBoundaryContext.Provider value={PropertyPanePickerControlBoundary}>{children}</PickerControlBoundaryContext.Provider>
      </PickerRuntimeContext.Provider>
    </FluentProvider>
  </RendererProvider>;
}
