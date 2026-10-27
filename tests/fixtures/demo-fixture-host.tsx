import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { FluentProvider, webLightTheme, teamsLightTheme, teamsDarkTheme, teamsHighContrastTheme } from '@fluentui/react-components';
import { createDOMRenderer } from '@griffel/core';
import { RendererProvider } from '@griffel/react';

interface FixtureHostPatch {
  theme?: 'light' | 'dark' | 'contrast' | 'default' | 'custom';
  dir?: 'ltr' | 'rtl';
  brandColor?: string;
}
declare global { interface Window { demoFixtureHost: { update(patch: FixtureHostPatch): void; readonly config: Required<FixtureHostPatch> }; } }
const FixtureThemeContext = React.createContext('light');
/** Only SDK data is simulated; the Fluent theme hook remains the real implementation. */
export function useFixtureTeams(): { supported: boolean; theme?: string } {
  const theme = React.useContext(FixtureThemeContext);
  return { supported: theme !== 'light', theme: theme === 'light' ? undefined : theme };
}
/** Stable fixture root: updates host tokens/direction without replacing panel instances. */
export function mountDemoFixture(children: React.ReactNode): void {
  const root = document.getElementById('root');
  if (!root) throw new Error('Demo fixture root missing');
  const renderer = createDOMRenderer(document);
  let state: Required<FixtureHostPatch> = { theme: 'light', dir: 'ltr', brandColor: '#7b2cbf' };
  const update = (patch: FixtureHostPatch): void => {
    state = { ...state, ...patch };
    const theme = state.theme === 'custom' ? { ...webLightTheme, colorBrandBackground: state.brandColor, colorBrandForeground1: state.brandColor } : state.theme === 'dark' ? teamsDarkTheme : state.theme === 'contrast' ? teamsHighContrastTheme : state.theme === 'default' ? teamsLightTheme : webLightTheme;
    ReactDOM.render(<RendererProvider renderer={renderer}><FixtureThemeContext.Provider value={state.theme}>
      <FluentProvider theme={theme} dir={state.dir}>{children}</FluentProvider>
    </FixtureThemeContext.Provider></RendererProvider>, root);
  };
  window.demoFixtureHost = { update, get config() { return { ...state }; } };
  update({});
  window.addEventListener('beforeunload', () => { ReactDOM.unmountComponentAtNode(root); });
}
