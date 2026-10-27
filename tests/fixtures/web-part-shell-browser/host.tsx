import * as React from 'react';
import * as ReactDOM from 'react-dom';
import type { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { RendererProvider } from '@griffel/react';
import { createDOMRenderer } from '@griffel/core';
import { mountDemoFixture } from '../demo-fixture-host';
import EditableTextPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/EditableTextPanel';
import { EditableTextIsolationFixture } from '../demo-control-isolation/editable-text-regressions';
import { createTheme } from '@fluentui/react';
import { WebPartShell, useSx, width, type IWebPartShellProperties, type WebPartShellProps, type EditableTextStrings, type EditableTextStringsOverride } from '@apvee/spfx-react-toolkit';
import { WebPartShellDemoHost } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/WebPartShellDemoHost';
import WebPartShellPanel from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/WebPartShellPanel';

type HostName = 'SharePoint' | 'SingleWebPartAppPage' | 'Teams' | 'Office' | 'Outlook';
type ThemeName = 'light' | 'dark' | 'highContrast';
type EditorDictionaryResource = Omit<EditableTextStrings, 'headingLevel' | 'lineSpacingValue' | 'formattingShortcut' | 'mixedFormatting'> & {
  headingLevelTemplate: string; lineSpacingValueTemplate: string;
  formattingShortcutTemplate: string; mixedFormattingTemplate: string;
};

// Test-only caller dictionaries derive from canonical JSON supplied by the runner.
// Native loader selection has separate AMD tests; changing this data preserves the UI instance.
function editorDictionary(resource: EditorDictionaryResource | undefined): EditableTextStringsOverride | undefined {
  if (!resource) return undefined;
  const template = (value: string, tokens: Record<string, string | number>): string =>
    value.replace(/\{([^{}]+)\}/g, (_match, key: string) => String(tokens[key]));
  return {
    editorLoadError: resource.editorLoadError,
    retry: resource.retry,
    savedFormattingWarning: resource.savedFormattingWarning,
    loadingEditor: resource.loadingEditor,
    placeholder: resource.placeholder,
    tooLargeError: resource.tooLargeError,
    moreFormatting: resource.moreFormatting,
    heading: resource.heading,
    normal: resource.normal,
    bold: resource.bold,
    italic: resource.italic,
    underline: resource.underline,
    themeColor: resource.themeColor,
    clearFormatting: resource.clearFormatting,
    spacing: resource.spacing,
    defaultLineSpacing: resource.defaultLineSpacing,
    spaceBefore: resource.spaceBefore,
    spaceAfter: resource.spaceAfter,
    formattingControls: resource.formattingControls,
    formatting: resource.formatting,
    fontSize: resource.fontSize,
    mixed: resource.mixed,
    themeColorGroup: resource.themeColorGroup,
    automatic: resource.automatic,
    themeColors: resource.themeColors,
    paletteColors: resource.paletteColors,
    colorNames: resource.colorNames,
    headingLevel: level => template(resource.headingLevelTemplate, { level }),
    lineSpacingValue: value => template(resource.lineSpacingValueTemplate, { value }),
    formattingShortcut: (label, shortcut) => template(resource.formattingShortcutTemplate, { label, shortcut }),
    mixedFormatting: label => template(resource.mixedFormattingTemplate, { label }),
  };
}

interface FixturePatch {
  mode?: 'Read' | 'Edit'; host?: HostName; theme?: ThemeName; dir?: 'ltr' | 'rtl';
  width?: number; otherWidth?: number;
  uiCulture?: string; otherUiCulture?: string; regionalCulture?: string;
  strings?: WebPartShellProps<IWebPartShellProperties>['strings'];
  otherStrings?: WebPartShellProps<IWebPartShellProperties>['strings'];
  titleEditorResource?: EditorDictionaryResource; otherTitleEditorResource?: EditorDictionaryResource;
  size?: WebPartShellProps<IWebPartShellProperties>['size'];
  appearance?: WebPartShellProps<IWebPartShellProperties>['appearance'];
  titlePosition?: WebPartShellProps<IWebPartShellProperties>['titlePosition'];
  surface?: WebPartShellProps<IWebPartShellProperties>['surface'];
  padding?: WebPartShellProps<IWebPartShellProperties>['padding'];
  fluentTheme?: WebPartShellProps<IWebPartShellProperties>['theme'];
  properties?: Partial<IWebPartShellProperties>; otherProperties?: Partial<IWebPartShellProperties>;
}
interface HostDouble {
  instanceId: string; domElement: HTMLElement; properties: IWebPartShellProperties & { listId: string };
  displayMode: number; render: () => void;
  context: {
    instanceId: string; domElement: HTMLElement;
    pageContext: { cultureInfo: { isRightToLeft: boolean; currentUICultureName: string; currentCultureName: string }; legacyPageContext: { pageLayoutType?: string } };
    sdks: { microsoftTeams?: { teamsJs: { app: { getContext: () => Promise<unknown>; registerOnThemeChangeHandler: (handler: (theme: string) => void) => void } } } };
    serviceScope: { whenFinished: (callback: () => void) => void; consume: (key: unknown) => unknown };
    propertyPane: { refresh: () => void };
  };
}
interface ShellFixture {
  render: (patch: FixturePatch) => void;
  observe: () => { properties: IWebPartShellProperties; otherProperties: IWebPartShellProperties; refreshes: number[]; mounts: number; cleanups: number; subscriptions: number[] };
}
declare global { interface Window { shellFixture: ShellFixture; } }

export function mountShellFixture(boundary: 'source' | 'compiled'): void {
  const scenario = new URLSearchParams(location.search).get('scenario');
  if (scenario === 'editable-text') { mountDemoFixture(<EditableTextPanel />); return; }
  const root = document.getElementById('root');
  if (!root) throw new Error('Shell fixture root missing');
  const renderer = createDOMRenderer(document);
  const containers = [document.createElement('div'), document.createElement('div')];
  const refreshes = [0, 0]; const subscriptions = [new Set<(args: { theme: ReturnType<typeof createTheme> }) => void>(), new Set<(args: { theme: ReturnType<typeof createTheme> }) => void>()];
  let mounts = 0, cleanups = 0;
  const state: Required<Pick<FixturePatch, 'mode' | 'host' | 'theme' | 'dir' | 'width' | 'otherWidth'>> & FixturePatch = { mode: 'Read', host: 'SharePoint', theme: 'light', dir: 'ltr', width: 900, otherWidth: 320, surface: 'auto', padding: 'auto', uiCulture: 'en-US', otherUiCulture: 'en-US', regionalCulture: 'it-IT' };
  const theme = (): ReturnType<typeof createTheme> => createTheme({ isInverted: state.theme !== 'light' });
  const sdkHandlers = new Set<(theme: string) => void>();
  const sdk = { app: { getContext: async (): Promise<unknown> => ({ app: { host: { name: state.host }, theme: state.theme === 'light' ? 'default' : state.theme === 'highContrast' ? 'contrast' : 'dark' } }), registerOnThemeChangeHandler: (handler: (theme: string) => void): void => { sdkHandlers.add(handler); } } };
  const hosts: HostDouble[] = containers.map((domElement, index) => {
    const instanceId = `shell-browser-${index}`;
    const service = { tryGetTheme: theme, themeChangedEvent: { add: (_observer: unknown, handler: (args: { theme: ReturnType<typeof createTheme> }) => void): void => { subscriptions[index].add(handler); }, remove: (_observer: unknown, handler: (args: { theme: ReturnType<typeof createTheme> }) => void): void => { subscriptions[index].delete(handler); } } };
    return { instanceId, domElement, properties: { title: index ? 'Independent title' : 'Alpha Beta Gamma', listId: `keep-${index}` }, displayMode: 1, render: () => {}, context: { instanceId, domElement, pageContext: { cultureInfo: { isRightToLeft: false, currentUICultureName: 'en-US', currentCultureName: 'it-IT' }, legacyPageContext: {} }, sdks: {}, serviceScope: { whenFinished: callback => callback(), consume: key => { if (key !== 'shell-browser-theme-service') throw new Error(`Unexpected SPFx service ${String(key)}`); return service; } }, propertyPane: { refresh: () => { refreshes[index]++; } } } };
  });
  function PersistentChild(): JSX.Element {
    React.useEffect(() => { mounts++; return () => { cleanups++; }; }, []);
    return <WebPartShellPanel />;
  }
  function Fixture(): JSX.Element {
    const sx = useSx();
    // Only SDK/SPFx services are doubled. This structural host supplies precisely
    // the BaseClientSideWebPart fields consumed by the real provider and Shell.
    const first = hosts[0] as unknown as BaseClientSideWebPart<IWebPartShellProperties>;
    const second = hosts[1] as unknown as BaseClientSideWebPart<IWebPartShellProperties>;
    return <RendererProvider renderer={renderer} targetDocument={document}>
      <div className={sx(width.px(state.width))} data-shell-region="first">
        <WebPartShellDemoHost instance={first} surface={state.surface} padding={state.padding} appearance={state.appearance} size={state.size} titlePosition={state.titlePosition} theme={state.fluentTheme} dir={state.dir} strings={state.strings} titleEditorStrings={editorDictionary(state.titleEditorResource)} data-shell-root="first"><PersistentChild /></WebPartShellDemoHost>
      </div>
      <div className={sx(width.px(state.otherWidth))} data-shell-region="second">
        <WebPartShell instance={second} strings={state.otherStrings} titleEditorStrings={editorDictionary(state.otherTitleEditorResource)} data-shell-root="second"><input aria-label="Other child" defaultValue="Other state" /></WebPartShell>
      </div>
      <p data-fixture-boundary="">{boundary}: real provider/store/hooks; SPFx services and SDK only. Local fixture, not authenticated SharePoint validation.</p>
      <button data-fixture-mode="" onClick={() => render({ mode: state.mode === 'Read' ? 'Edit' : 'Read' })}>{state.mode === 'Read' ? 'Enter page Edit mode' : 'Return to page Read mode'}</button>
      <button data-fixture-outside="">Outside title editor</button>
    </RendererProvider>;
  }
  const render = (patch: FixturePatch): void => {
    Object.assign(state, patch);
    hosts.forEach((host, index) => {
      host.displayMode = state.mode === 'Edit' ? 2 : 1;
      host.context.pageContext.cultureInfo.isRightToLeft = state.dir === 'rtl';
      host.context.pageContext.cultureInfo.currentUICultureName = (index === 0 ? state.uiCulture : state.otherUiCulture) ?? 'en-US';
      host.context.pageContext.cultureInfo.currentCultureName = state.regionalCulture ?? 'it-IT';
      host.context.pageContext.legacyPageContext.pageLayoutType = state.host === 'SingleWebPartAppPage' ? state.host : undefined;
      host.context.sdks.microsoftTeams = ['Teams', 'Office', 'Outlook'].includes(state.host) ? { teamsJs: sdk } : undefined;
      if (index === 0 && patch.properties) Object.assign(host.properties, patch.properties);
      if (index === 1 && patch.otherProperties) Object.assign(host.properties, patch.otherProperties);
    });
    root.className = `host-theme-${state.theme}`;
    ReactDOM.render(<Fixture />, root);
    for (const set of subscriptions) for (const handler of set) handler({ theme: theme() });
    for (const handler of sdkHandlers) handler(state.theme === 'light' ? 'default' : state.theme === 'highContrast' ? 'contrast' : 'dark');
  };
  window.shellFixture = { render, observe: () => ({ properties: { ...hosts[0].properties }, otherProperties: { ...hosts[1].properties }, refreshes: [...refreshes], mounts, cleanups, subscriptions: subscriptions.map(set => set.size) }) };
  if (scenario === 'editable-text-regressions') {
    hosts[0].displayMode = 2;
    mountDemoFixture(<EditableTextIsolationFixture instance={hosts[0] as unknown as BaseClientSideWebPart<IWebPartShellProperties>} />);
    return;
  }
  render({});
}
