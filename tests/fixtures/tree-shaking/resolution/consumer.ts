import { width as rootWidth, useSx as rootSx, useStableCallback as rootCallback } from '@apvee/spfx-react-toolkit';
import { SPFxWebPartProvider } from '@apvee/spfx-react-toolkit/core';
import { useSPFxProperties, useSPFxPnPListById } from '@apvee/spfx-react-toolkit/hooks';
import { createSPFxPnPListService } from '@apvee/spfx-react-toolkit/services';
import { width as helperWidth } from '@apvee/spfx-react-toolkit/helpers';
import { width, useSx } from '@apvee/spfx-react-toolkit/styles';
import type { SxDescriptor, SxFunction } from '@apvee/spfx-react-toolkit/styles';
import { useSx as leafSx } from '@apvee/spfx-react-toolkit/styles/useSx';
import { useStableCallback } from '@apvee/spfx-react-toolkit/hooks/useStableCallback';
import { width as directoryWidth } from '@apvee/spfx-react-toolkit/lib/helpers/styles';
import { useStableCallback as legacyCallback } from '@apvee/spfx-react-toolkit/lib/hooks/useStableCallback';
import { useSx as legacySx } from '@apvee/spfx-react-toolkit/lib/hooks/useSx.js';
import type { SPFxPnPListSelector } from '@apvee/spfx-react-toolkit/lib/services/spfx-pnp-list.service';
import { serializeValue } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxTenantKeyValueStore.serialization.internal';
import { useSPFxContext } from '@apvee/spfx-react-toolkit/lib';
// @ts-expect-error Private clean paths are deliberately unavailable.
import {} from '@apvee/spfx-react-toolkit/hooks/useAsyncInvoke.internal';
// @ts-expect-error Private descriptor implementation has no clean alias.
import {} from '@apvee/spfx-react-toolkit/styles/descriptor.internal';
// @ts-expect-error Types are public through styles, not the branded internal types module.
import {} from '@apvee/spfx-react-toolkit/styles/types';
// @ts-expect-error Private brands must not be re-exported by the public facade.
import { sxDescriptorBrand } from '@apvee/spfx-react-toolkit/styles';

export const contracts = { rootWidth, rootSx, rootCallback, SPFxWebPartProvider, createSPFxPnPListService,
  helperWidth, directoryWidth, legacyCallback, legacySx, serializeValue, useSPFxContext };
export const descriptor: SxDescriptor = width.full;
export const composer: typeof useSx = leafSx;
export function useContracts(): { title: string | undefined; text: string; pending: Promise<number>; sx: SxFunction } {
  const properties = useSPFxProperties<{ title: string }>();
  properties.setProperties({ title: 'Title' });
  // @ts-expect-error Property generic cannot become any.
  properties.setProperties({ title: 7 });
  const list = useSPFxPnPListById<{ Title: string }>('guid');
  // @ts-expect-error List item fields retain the generic shape.
  list.create({ Title: false });
  const sync = useStableCallback((id: number, label: string) => label + id);
  const asyncCallback = legacyCallback(async (id: number) => id);
  // @ts-expect-error Callback argument types remain inferred.
  sync('bad', 'label');
  // @ts-expect-error Callback return types remain inferred.
  const wrong: number = sync(1, 'label');
  void wrong;
  return { title: properties.properties?.title, text: sync(1, 'label'), pending: asyncCallback(1), sx: useSx() };
}
export function readonlySelector(selector: Extract<SPFxPnPListSelector, { kind: 'id' }>): void {
  // @ts-expect-error Selector discriminants remain readonly.
  selector.kind = 'id';
  // @ts-expect-error Selector values remain readonly.
  selector.id = 'changed';
}

import { FormattedText as rootText } from '@apvee/spfx-react-toolkit';
import { FormattedText as domainText } from '@apvee/spfx-react-toolkit/controls';
import { FormattedText as controlText } from '@apvee/spfx-react-toolkit/controls/formatted-text';
import type { FormattedTextProps } from '@apvee/spfx-react-toolkit/controls';
// @ts-expect-error Control implementation modules are private.
import {} from '@apvee/spfx-react-toolkit/controls/formatted-text/formatted-text-content.internal';
export const textComponents: [typeof rootText, typeof rootText] = [domainText, controlText];
export const textProps: FormattedTextProps = { as: 'span', maxLines: 2, variant: 'title2' };
// @ts-expect-error There is no separate inline mode.
export const removedInline: FormattedTextProps = { inline: true };

import { RelativeTime as rootRelative } from '@apvee/spfx-react-toolkit';
import { RelativeTime as domainRelative } from '@apvee/spfx-react-toolkit/controls';
import { RelativeTime as controlRelative } from '@apvee/spfx-react-toolkit/controls/relative-time';
import type { RelativeTimeProps } from '@apvee/spfx-react-toolkit/controls/relative-time';
import type { SPFxLocaleInfo } from '@apvee/spfx-react-toolkit';
export const relativeComponents: [typeof rootRelative, typeof rootRelative] = [domainRelative, controlRelative];
export const relativeProps: RelativeTimeProps = { dateTime: 0, referenceTime: 60000, locale: 'en' };
export function mappedZone(info: SPFxLocaleInfo): string | undefined { return info.ianaTimeZone; }
// @ts-expect-error The derived IANA field is readonly.
export function readonlyZone(info: SPFxLocaleInfo): void { info.ianaTimeZone = 'UTC'; }
// @ts-expect-error Private lookup has no clean alias.
import {} from '@apvee/spfx-react-toolkit/helpers/spfx-time-zone.internal';

import { AutoSkeleton as rootSkeleton } from '@apvee/spfx-react-toolkit';
import { AutoSkeleton as domainSkeleton } from '@apvee/spfx-react-toolkit/controls';
import { AutoSkeleton as controlSkeleton } from '@apvee/spfx-react-toolkit/controls/auto-skeleton';
import type { AutoSkeletonProps } from '@apvee/spfx-react-toolkit/controls/auto-skeleton';
export const skeletons = [rootSkeleton, domainSkeleton, controlSkeleton];
export const skeletonProps: AutoSkeletonProps = { loading: true, children: 'Profile' };

import { StateMessage as rootState } from '@apvee/spfx-react-toolkit';
import { StateMessage as domainState } from '@apvee/spfx-react-toolkit/controls';
import { StateMessage as controlState } from '@apvee/spfx-react-toolkit/controls/state-message';
import type { StateMessageProps } from '@apvee/spfx-react-toolkit/controls/state-message';
export const states: [typeof rootState, typeof rootState] = [domainState, controlState];
export const stateProps: StateMessageProps = { title: 'No results', preset: 'empty', visual: false };

import { SettingsGroup as rootSettings, AccordionGroup as rootAccordion } from '@apvee/spfx-react-toolkit';
import { SettingsGroup as domainSettings, AccordionGroup as domainAccordion } from '@apvee/spfx-react-toolkit/controls';
import { SettingsGroup as controlSettings } from '@apvee/spfx-react-toolkit/controls/settings-group';
import { AccordionGroup as controlAccordion } from '@apvee/spfx-react-toolkit/controls/accordion-group';
import type { SettingsGroupProps, SettingsGroupItemProps } from '@apvee/spfx-react-toolkit/controls/settings-group';
import type { AccordionGroupProps, AccordionGroupItemProps } from '@apvee/spfx-react-toolkit/controls/accordion-group';
export const settingsAliases: typeof rootSettings[] = [domainSettings, controlSettings];
export const accordionAliases: typeof rootAccordion[] = [domainAccordion, controlAccordion];
export const settingsGroupProps: SettingsGroupProps = { title: '**Preferences**' };
export const settingsItemProps: SettingsGroupItemProps = { description: 'Name' };
export const accordionGroupProps: AccordionGroupProps = { defaultOpenItems: ['profile'] };
export const accordionItemProps: AccordionGroupItemProps = { value: 'profile', title: 'Profile' };
export const settingsCompoundItem: typeof rootSettings.Item = controlSettings.Item;
export const accordionCompoundItem: typeof rootAccordion.Item = controlAccordion.Item;

import { Form as rootForm, createForm as rootFactory } from '@apvee/spfx-react-toolkit';
import { Form as domainForm, createForm as domainFactory } from '@apvee/spfx-react-toolkit/controls';
import { Form as controlForm, createForm as controlFactory } from '@apvee/spfx-react-toolkit/controls/form';
import type { FormMethods as RootMethods } from '@apvee/spfx-react-toolkit';
import type { FormMethods as DomainMethods } from '@apvee/spfx-react-toolkit/controls';
import type { FormMethods as ControlMethods } from '@apvee/spfx-react-toolkit/controls/form';
export const formAliases: typeof rootForm[] = [domainForm, controlForm];
export const factoryAliases: typeof rootFactory[] = [domainFactory, controlFactory];
export type FormMethodAliases = [RootMethods<{ text: string }>, DomainMethods<{ text: string }>, ControlMethods<{ text: string }>];
// @ts-expect-error Adapters are available through compound suite members only.
import { FormInput } from '@apvee/spfx-react-toolkit/controls/form';
// @ts-expect-error No clean implementation subpath.
import {} from '@apvee/spfx-react-toolkit/controls/form/form';
// @ts-expect-error No plugin subpath.
import {} from '@apvee/spfx-react-toolkit/controls/form/plugins';
// @ts-expect-error No advanced subpath.
import {} from '@apvee/spfx-react-toolkit/controls/form/advanced';
void FormInput;

import { DataGrid as rootGrid, createDataGrid as rootGridFactory } from '@apvee/spfx-react-toolkit';
import { DataGrid as domainGrid, createDataGrid as domainGridFactory } from '@apvee/spfx-react-toolkit/controls';
import { DataGrid as controlGrid, createDataGrid as controlGridFactory } from '@apvee/spfx-react-toolkit/controls/data-grid';
import type { DataGridProps as RootGridProps } from '@apvee/spfx-react-toolkit';
import type { DataGridProps as DomainGridProps } from '@apvee/spfx-react-toolkit/controls';
import type { DataGridProps as ControlGridProps } from '@apvee/spfx-react-toolkit/controls/data-grid';
export const gridAliases: typeof rootGrid[] = [domainGrid, controlGrid];
export const gridFactoryAliases: typeof rootGridFactory[] = [domainGridFactory, controlGridFactory];
export type GridPropAliases = [RootGridProps<{ id: string }>, DomainGridProps<{ id: string }>, ControlGridProps<{ id: string }>];
// @ts-expect-error Grid implementation has no public subpath.
import {} from '@apvee/spfx-react-toolkit/controls/data-grid/data-grid';
// @ts-expect-error Row model is private.
import {} from '@apvee/spfx-react-toolkit/controls/data-grid/prepared-rows.internal';
// @ts-expect-error Column cells are suite members, not flat public controls.
import { TextColumn } from '@apvee/spfx-react-toolkit/controls/data-grid';
// @ts-expect-error Table-only API has no View.
export const gridView = rootGrid.View;
void TextColumn;

import { WebPartShell as rootShell, getWebPartShellPropertiesMetadata as rootMetadata } from '@apvee/spfx-react-toolkit';
import { WebPartShell as domainShell, getWebPartShellPropertiesMetadata as domainMetadata } from '@apvee/spfx-react-toolkit/controls';
import { WebPartShell as controlShell, getWebPartShellPropertiesMetadata as controlMetadata } from '@apvee/spfx-react-toolkit/controls/web-part-shell';
import type { WebPartShellProps as RootShellProps } from '@apvee/spfx-react-toolkit';
import type { WebPartShellProps as DomainShellProps } from '@apvee/spfx-react-toolkit/controls';
import type { WebPartShellProps as ControlShellProps, IWebPartShellProperties } from '@apvee/spfx-react-toolkit/controls/web-part-shell';
export const shellAliases: typeof rootShell[] = [domainShell, controlShell];
export const metadataAliases: typeof rootMetadata[] = [domainMetadata, controlMetadata];
export type ShellPropAliases = [RootShellProps<IWebPartShellProperties>, DomainShellProps<IWebPartShellProperties>, ControlShellProps<IWebPartShellProperties>];
// @ts-expect-error Private editor has no clean subpath.
import {} from '@apvee/spfx-react-toolkit/controls/web-part-shell/title/web-part-title-editor.internal';

import { Stage as rootStage } from '@apvee/spfx-react-toolkit';
import { Stage as domainStage } from '@apvee/spfx-react-toolkit/controls';
import { Stage as controlStage } from '@apvee/spfx-react-toolkit/controls/stage';
import type { StageProps as RootStageProps } from '@apvee/spfx-react-toolkit';
import type { StageProps as DomainStageProps } from '@apvee/spfx-react-toolkit/controls';
import type { StageProps as ControlStageProps } from '@apvee/spfx-react-toolkit/controls/stage';
export const stageAliases: typeof rootStage[] = [domainStage, controlStage];
export type StagePropAliases = [RootStageProps, DomainStageProps, ControlStageProps];
export const stageHeader: typeof rootStage.Header = controlStage.Header;
// @ts-expect-error Stage implementation is private.
import {} from '@apvee/spfx-react-toolkit/controls/stage/stage';
// @ts-expect-error Stage context is private.
import {} from '@apvee/spfx-react-toolkit/controls/stage/stage-context.internal';

import { ContentCard as rootContentCard, PreviewCard as rootPreviewCard } from '@apvee/spfx-react-toolkit';
import { ContentCard as domainContentCard, PreviewCard as domainPreviewCard } from '@apvee/spfx-react-toolkit/controls';
import { ContentCard as controlContentCard } from '@apvee/spfx-react-toolkit/controls/content-card';
import { PreviewCard as controlPreviewCard } from '@apvee/spfx-react-toolkit/controls/preview-card';
import type { ContentCardProps as RootCardProps, PreviewCardProps as RootPreviewProps } from '@apvee/spfx-react-toolkit';
import type { ContentCardProps as DomainCardProps, PreviewCardProps as DomainPreviewProps } from '@apvee/spfx-react-toolkit/controls';
import type { ContentCardProps as ControlCardProps } from '@apvee/spfx-react-toolkit/controls/content-card';
import type { PreviewCardProps as ControlPreviewProps, PreviewCardImage, PreviewCardOrientation } from '@apvee/spfx-react-toolkit/controls/preview-card';
export const contentCards: typeof rootContentCard[] = [domainContentCard, controlContentCard];
export const previewCards: typeof rootPreviewCard[] = [domainPreviewCard, controlPreviewCard];
export const contentCardProps: RootCardProps & DomainCardProps & ControlCardProps = { title: 'Summary' };
export const previewCardImage: PreviewCardImage = { src: '/cover.jpg', alt: 'Cover' };
export const previewCardOrientation: PreviewCardOrientation = 'horizontal';
export const previewCardProps: RootPreviewProps & DomainPreviewProps & ControlPreviewProps = { title: 'Report', image: previewCardImage, orientation: previewCardOrientation };
// @ts-expect-error Private frame has no clean alias.
import {} from '@apvee/spfx-react-toolkit/controls/cards/card-frame.internal';
// @ts-expect-error Private resolver has no clean alias.
import {} from '@apvee/spfx-react-toolkit/controls/content-slot.internal';

import { SharePointPermissionTrim as rootPermissionTrim } from '@apvee/spfx-react-toolkit';
import type { SharePointPermissionTrimProps as RootPermissionTrimProps } from '@apvee/spfx-react-toolkit';

import { SharePointPermissionTrim as domainPermissionTrim } from '@apvee/spfx-react-toolkit/controls';
import type { SharePointPermissionTrimProps as DomainPermissionTrimProps } from '@apvee/spfx-react-toolkit/controls';

import { SharePointPermissionTrim as controlPermissionTrim } from '@apvee/spfx-react-toolkit/controls/sharepoint-permission-trim';
import type { SharePointPermissionTrimProps as ControlPermissionTrimProps } from '@apvee/spfx-react-toolkit/controls/sharepoint-permission-trim';
export const PermissionTrimAliases: typeof rootPermissionTrim[] = [domainPermissionTrim, controlPermissionTrim];
export type PermissionTrimPropAliases = [RootPermissionTrimProps, DomainPermissionTrimProps, ControlPermissionTrimProps];
export const PermissionTrimLoading: typeof rootPermissionTrim.Loading = controlPermissionTrim.Loading;

import { SharePointGroupTrim as rootSPGroupTrim } from '@apvee/spfx-react-toolkit';
import type { SharePointGroupTrimProps as RootSPGroupTrimProps } from '@apvee/spfx-react-toolkit';

import { SharePointGroupTrim as domainSPGroupTrim } from '@apvee/spfx-react-toolkit/controls';
import type { SharePointGroupTrimProps as DomainSPGroupTrimProps } from '@apvee/spfx-react-toolkit/controls';

import { SharePointGroupTrim as controlSPGroupTrim } from '@apvee/spfx-react-toolkit/controls/sharepoint-group-trim';
import type { SharePointGroupTrimProps as ControlSPGroupTrimProps } from '@apvee/spfx-react-toolkit/controls/sharepoint-group-trim';
export const SPGroupTrimAliases: typeof rootSPGroupTrim[] = [domainSPGroupTrim, controlSPGroupTrim];
export type SPGroupTrimPropAliases = [RootSPGroupTrimProps, DomainSPGroupTrimProps, ControlSPGroupTrimProps];
export const SPGroupTrimLoading: typeof rootSPGroupTrim.Loading = controlSPGroupTrim.Loading;

import { EntraGroupTrim as rootEntraTrim } from '@apvee/spfx-react-toolkit';
import type { EntraGroupTrimProps as RootEntraTrimProps } from '@apvee/spfx-react-toolkit';

import { EntraGroupTrim as domainEntraTrim } from '@apvee/spfx-react-toolkit/controls';
import type { EntraGroupTrimProps as DomainEntraTrimProps } from '@apvee/spfx-react-toolkit/controls';

import { EntraGroupTrim as controlEntraTrim } from '@apvee/spfx-react-toolkit/controls/entra-group-trim';
import type { EntraGroupTrimProps as ControlEntraTrimProps } from '@apvee/spfx-react-toolkit/controls/entra-group-trim';
export const EntraTrimAliases: typeof rootEntraTrim[] = [domainEntraTrim, controlEntraTrim];
export type EntraTrimPropAliases = [RootEntraTrimProps, DomainEntraTrimProps, ControlEntraTrimProps];
export const EntraTrimLoading: typeof rootEntraTrim.Loading = controlEntraTrim.Loading;
import { useSPFxCrossSitePermissions as rootScopedHook } from '@apvee/spfx-react-toolkit';
import { useSPFxCrossSitePermissions as hooksScopedHook } from '@apvee/spfx-react-toolkit/hooks';
import { useSPFxCrossSitePermissions as legacyScopedHook } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxCrossSitePermissions';
import type { SPFxScopedPermissionsInfo as rootScopedInfo } from '@apvee/spfx-react-toolkit';
import type { SPFxScopedPermissionsInfo as hooksScopedInfo } from '@apvee/spfx-react-toolkit/hooks';
import type { SPFxScopedPermissionsInfo as legacyScopedInfo } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxCrossSitePermissions';
import type { SPFxScopedPermissionsOptions as rootScopedOptions } from '@apvee/spfx-react-toolkit';
import type { SPFxScopedPermissionsOptions as hooksScopedOptions } from '@apvee/spfx-react-toolkit/hooks';
import type { SPFxScopedPermissionsOptions as legacyScopedOptions } from '@apvee/spfx-react-toolkit/lib/hooks/useSPFxCrossSitePermissions';

import type { SharePointPermissionTrimBranchProps as RootSharePointPermissionTrimBranchProps } from '@apvee/spfx-react-toolkit';
import type { SharePointPermissionTrimBranchProps as DomainSharePointPermissionTrimBranchProps } from '@apvee/spfx-react-toolkit/controls';
import type { SharePointPermissionTrimBranchProps as ControlSharePointPermissionTrimBranchProps } from '@apvee/spfx-react-toolkit/controls/sharepoint-permission-trim';
export type SharePointPermissionTrimBranchPropsAliases = [RootSharePointPermissionTrimBranchProps, DomainSharePointPermissionTrimBranchProps, ControlSharePointPermissionTrimBranchProps];

import type { SharePointGroupSelector as RootSharePointGroupSelector } from '@apvee/spfx-react-toolkit';
import type { SharePointGroupSelector as DomainSharePointGroupSelector } from '@apvee/spfx-react-toolkit/controls';
import type { SharePointGroupSelector as ControlSharePointGroupSelector } from '@apvee/spfx-react-toolkit/controls/sharepoint-group-trim';
export type SharePointGroupSelectorAliases = [RootSharePointGroupSelector, DomainSharePointGroupSelector, ControlSharePointGroupSelector];

import type { SharePointGroupTrimBranchProps as RootSharePointGroupTrimBranchProps } from '@apvee/spfx-react-toolkit';
import type { SharePointGroupTrimBranchProps as DomainSharePointGroupTrimBranchProps } from '@apvee/spfx-react-toolkit/controls';
import type { SharePointGroupTrimBranchProps as ControlSharePointGroupTrimBranchProps } from '@apvee/spfx-react-toolkit/controls/sharepoint-group-trim';
export type SharePointGroupTrimBranchPropsAliases = [RootSharePointGroupTrimBranchProps, DomainSharePointGroupTrimBranchProps, ControlSharePointGroupTrimBranchProps];

import type { EntraGroupSelector as RootEntraGroupSelector } from '@apvee/spfx-react-toolkit';
import type { EntraGroupSelector as DomainEntraGroupSelector } from '@apvee/spfx-react-toolkit/controls';
import type { EntraGroupSelector as ControlEntraGroupSelector } from '@apvee/spfx-react-toolkit/controls/entra-group-trim';
export type EntraGroupSelectorAliases = [RootEntraGroupSelector, DomainEntraGroupSelector, ControlEntraGroupSelector];

import type { EntraGroupType as RootEntraGroupType } from '@apvee/spfx-react-toolkit';
import type { EntraGroupType as DomainEntraGroupType } from '@apvee/spfx-react-toolkit/controls';
import type { EntraGroupType as ControlEntraGroupType } from '@apvee/spfx-react-toolkit/controls/entra-group-trim';
export type EntraGroupTypeAliases = [RootEntraGroupType, DomainEntraGroupType, ControlEntraGroupType];

import type { EntraGroupRelationship as RootEntraGroupRelationship } from '@apvee/spfx-react-toolkit';
import type { EntraGroupRelationship as DomainEntraGroupRelationship } from '@apvee/spfx-react-toolkit/controls';
import type { EntraGroupRelationship as ControlEntraGroupRelationship } from '@apvee/spfx-react-toolkit/controls/entra-group-trim';
export type EntraGroupRelationshipAliases = [RootEntraGroupRelationship, DomainEntraGroupRelationship, ControlEntraGroupRelationship];

import type { EntraGroupTrimBranchProps as RootEntraGroupTrimBranchProps } from '@apvee/spfx-react-toolkit';
import type { EntraGroupTrimBranchProps as DomainEntraGroupTrimBranchProps } from '@apvee/spfx-react-toolkit/controls';
import type { EntraGroupTrimBranchProps as ControlEntraGroupTrimBranchProps } from '@apvee/spfx-react-toolkit/controls/entra-group-trim';
export type EntraGroupTrimBranchPropsAliases = [RootEntraGroupTrimBranchProps, DomainEntraGroupTrimBranchProps, ControlEntraGroupTrimBranchProps];

// Picker symbol identity is shared across root/controls/domain; native factories remain subpath-only.
import { SitePicker as rootSitePicker, ListPicker as rootListPicker, FieldPicker as rootFieldPicker, ViewPicker as rootViewPicker, ListItemPicker as rootItemPicker, PrincipalPicker as rootPrincipalPicker, type PrincipalPickerValue as RootPrincipalValue } from '@apvee/spfx-react-toolkit';
import { SitePicker as controlSitePicker, ListPicker as controlListPicker, FieldPicker as controlFieldPicker, ViewPicker as controlViewPicker, ListItemPicker as controlItemPicker, PrincipalPicker as controlPrincipalPicker, type PrincipalPickerValue as ControlPrincipalValue } from '@apvee/spfx-react-toolkit/controls';
import { SitePicker as domainSitePicker, ListPicker as domainListPicker, FieldPicker as domainFieldPicker, ViewPicker as domainViewPicker, ListItemPicker as domainItemPicker, PrincipalPicker as domainPrincipalPicker, type PrincipalPickerValue as DomainPrincipalValue, type SitePickerValue } from '@apvee/spfx-react-toolkit/controls/pickers';
import { PropertyPaneSitePicker, PropertyPaneListPicker, PropertyPaneFieldPicker, PropertyPaneViewPicker, PropertyPaneListItemPicker, PropertyPanePrincipalPicker, type PrincipalPickerValue as NativePrincipalValue, type PickerPropertyKey } from '@apvee/spfx-react-toolkit/property-pane/pickers';
// @ts-expect-error native runtime is intentionally not a root export
import { PropertyPanePrincipalPicker as invalidRootNativePicker } from '@apvee/spfx-react-toolkit';
const pickerImports = [rootSitePicker, rootListPicker, rootFieldPicker, rootViewPicker, rootItemPicker, rootPrincipalPicker, controlSitePicker, controlListPicker, controlFieldPicker, controlViewPicker, controlItemPicker, controlPrincipalPicker, domainSitePicker, domainListPicker, domainFieldPicker, domainViewPicker, domainItemPicker, domainPrincipalPicker, PropertyPaneSitePicker, PropertyPaneListPicker, PropertyPaneFieldPicker, PropertyPaneViewPicker, PropertyPaneListItemPicker, PropertyPanePrincipalPicker];
type PickerPropsBag = { sites?: SitePickerValue[]; principals?: NativePrincipalValue[]; title: string };
const siteKey: PickerPropertyKey<PickerPropsBag, SitePickerValue> = 'sites';
// @ts-expect-error scalar property is not a picker JSON array
const invalidSiteKey: PickerPropertyKey<PickerPropsBag, SitePickerValue> = 'title';
const nativeValues: NativePrincipalValue[] = [{ kind: 'user', loginName: 'alice@tenant.com' }];
const rootValues: RootPrincipalValue[] = nativeValues, controlValues: ControlPrincipalValue[] = rootValues, domainValues: DomainPrincipalValue[] = controlValues;
void [pickerImports, siteKey, invalidSiteKey, domainValues, invalidRootNativePicker];

import { EditableText as rootEditable, type EditableTextProps as RootEditableProps, type EditableTextBlock as RootEditableBlock, type EditableTextChangeData as RootEditableChange } from '@apvee/spfx-react-toolkit';
import { EditableText as domainEditable, type EditableTextProps as DomainEditableProps, type EditableTextBlock as DomainEditableBlock, type EditableTextChangeData as DomainEditableChange } from '@apvee/spfx-react-toolkit/controls';
import { EditableText as controlEditable, type EditableTextProps as ControlEditableProps, type EditableTextBlock as ControlEditableBlock, type EditableTextChangeData as ControlEditableChange } from '@apvee/spfx-react-toolkit/controls/editable-text';
export const editableAliases: typeof rootEditable[] = [domainEditable, controlEditable];
export type EditableAliases = [RootEditableProps, DomainEditableProps, ControlEditableProps, RootEditableBlock, DomainEditableBlock, ControlEditableBlock, RootEditableChange, DomainEditableChange, ControlEditableChange];
