import * as React from 'react';
import { useSPFxPnPContext } from '../../hooks/useSPFxPnPContext';
import type { SitePickerProps, SitePickerOption } from './picker.types';
import { createSitePickerSource } from './site-picker.source.internal';
import type { SiteSearchTransport } from './site-picker.source.internal';
import { usePickerSourceContext } from './usePickerSourceContext.internal';
import type { SitePickerValue } from './picker-values.types';
import { usePickerRuntime } from './usePickerRuntime.internal';
import { PickerRenderer } from './picker-renderer.internal';

/** Search site collections or hubs with the provider's authenticated clients and memory-only query caching.
 * Defaults: site collection mode, Dropdown (historical combobox presentation), single selection, Title ascending. Initial menu includes consumer options, current collection and associated hub; includeCurrentSite/includeHubSite both default true.
 * @param props Site mode/scope, models, selection and presentation options. KQL is supported only in mode="site".
 * @returns A native Fluent Field and site picker; changing display props retains its runtime.
 * @example <SitePicker label="Sito" mode="site" includeCurrentSite includeHubSite />
 */
export function SitePicker(props: SitePickerProps): React.ReactElement {
  const context = usePickerSourceContext(), pnp = useSPFxPnPContext(undefined, { cache: { enabled: false } });
  const search = React.useMemo<SiteSearchTransport | undefined>(() => {
    const sp = pnp.sp;
    if (sp === undefined) return undefined;
    return async (query, options) => {
      // The owning existing search service registers PnP search lazily when this feature is requested.
      const { createSPFxPnPSearchService } = await import(/* webpackChunkName: 'spfx-picker-search' */ '../../services/spfx-pnp-search.service');
      return createSPFxPnPSearchService(sp, { selectProperties: ['Title', 'Path', 'SiteId', 'DepartmentId'] }).search(builder => {
        const sorted = builder.text(query);
        sorted.toSearchQuery().TrimDuplicates = options.trimDuplicates;
        return sorted;
      }, { pageSize: options.pageSize, startRow: options.startRow });
    };
  }, [pnp.sp]);
  const source = createSitePickerSource(props, context, search), runtime = usePickerRuntime<SitePickerValue, SitePickerOption>(props, source, { multiselect: false });
  return <PickerRenderer<SitePickerValue, SitePickerOption> props={{ ...props, placeholder: props.searchPlaceholder ?? props.placeholder }} runtime={runtime} defaultPresentation="combobox" defaultMultiselect={false} />;
}
