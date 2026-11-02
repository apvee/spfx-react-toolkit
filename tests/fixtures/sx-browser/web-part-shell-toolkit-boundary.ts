// Actual source modules: the private SPFx host adapter supplies services, never hooks.
export * from '../../../packages/spfx-react-toolkit/src/controls/web-part-shell';
export { useSPFxProperties } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxProperties';
export { useSPFxContext } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxContext';
export { useSPFxDisplayMode } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxDisplayMode';
export { useSPFxFluent9ThemeInfo } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxFluent9ThemeInfo';
export { useSx } from '../../../packages/spfx-react-toolkit/src/hooks/useSx';
export * from '../../../packages/spfx-react-toolkit/src/helpers/styles';

export * from '../../../packages/spfx-react-toolkit/src/controls/editable-text';

export { useSPFxPnPContext } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPContext';
export { useSPFxPnPList } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPnPList';

// The real catalogue also emits the import-interoperability scenario.
export { useStableCallback } from '../../../packages/spfx-react-toolkit/src/hooks/useStableCallback';
export { Pagination, type PaginationChangeData, type PaginationProps, type PaginationSize, type PaginationState } from '../../../packages/spfx-react-toolkit/src/controls/pagination';

// The actual catalogue is typechecked and bundled even when this host selects Shell.
export * from '../../../packages/spfx-react-toolkit/src/controls/stage';
export * from '../../../packages/spfx-react-toolkit/src/controls/form';
export * from '../../../packages/spfx-react-toolkit/src/controls/content-card';
export * from '../../../packages/spfx-react-toolkit/src/controls/preview-card';
export * from '../../../packages/spfx-react-toolkit/src/controls/description-list';
export * from '../../../packages/spfx-react-toolkit/src/controls/settings-group';
export * from '../../../packages/spfx-react-toolkit/src/controls/accordion-group';
export * from '../../../packages/spfx-react-toolkit/src/controls/state-message';
export * from '../../../packages/spfx-react-toolkit/src/controls/auto-skeleton';
export * from '../../../packages/spfx-react-toolkit/src/controls/relative-time';
export * from '../../../packages/spfx-react-toolkit/src/controls/formatted-text';

// All generated cases are bundled by the Shell conformance fixture.
export { useSPFxSessionStorage, useSPFxLocalStorage } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxStorage';
export { useSPFxLogger } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxLogger';
export { useSPFxPerformance } from '../../../packages/spfx-react-toolkit/src/hooks/useSPFxPerformance';

// Prior generated DataGrid cases are also dependencies of this complete catalogue.
export * from '../../../packages/spfx-react-toolkit/src/controls/data-grid';
