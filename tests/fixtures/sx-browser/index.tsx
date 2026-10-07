import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { createDOMRenderer, makeStyles, mergeClasses } from '@griffel/core';
import { RendererProvider, useRenderer_unstable } from '@griffel/react';
import { FluentProvider } from '@fluentui/react-provider';
import { webLightTheme, webDarkTheme, teamsHighContrastTheme } from '@fluentui/react-theme';
import { useSx } from '../../../packages/spfx-react-toolkit/src/hooks/useSx';
import { container, width, responsive, viewport, hover, active, focusVisible, foreground,
  background as catalogBackground, presets, typography, scrollbar, overflow } from '../../../packages/spfx-react-toolkit/src/helpers/styles';
import { createDeclaration, createRecipe, SxProperty } from '../../../packages/spfx-react-toolkit/src/helpers/styles/descriptor.internal';
import type { SxBaseDescriptor, SxInput } from '../../../packages/spfx-react-toolkit/src/helpers/styles/types';

interface FixtureConfig {
  readonly size: number;
  readonly reverse: boolean;
  readonly scopes: boolean;
  readonly enabled: boolean;
  readonly selected: boolean;
  readonly dir: 'ltr' | 'rtl';
  readonly theme: 'light' | 'dark' | 'highContrast';
  readonly scrollbarEnabled: boolean;
}
interface FixtureApi {
  render(patch: Partial<FixtureConfig>): void;
  readonly config: FixtureConfig;
}
declare global { interface Window { sxFixture: FixtureApi; } }

// Private recipes exercise the generic engine without expanding the public catalog.
const declaration = (property: SxProperty, value: string, fallback: string): SxBaseDescriptor =>
  createDeclaration(property, value, { property, fallback });
const row = createRecipe('fixture.row', [declaration('display', 'flex', 'block'), declaration('flexDirection', 'row', 'row')]);
const column = createRecipe('fixture.column', [declaration('display', 'flex', 'block'), declaration('flexDirection', 'column', 'row')]);
const gap = (value: number): SxBaseDescriptor => declaration('rowGap', `${value}px`, 'normal');
const background = (value: string): SxBaseDescriptor => declaration('backgroundColor', value, 'transparent');
const red = background('rgb(200, 0, 0)');
const green = background('rgb(0, 160, 0)');
const blue = background('rgb(0, 0, 220)');
const purple = background('rgb(120, 0, 180)');
const disabled = background('rgb(100, 100, 100)');
const selected = background('rgb(200, 100, 0)');
const neutral = background('rgb(240, 240, 240)');
const externalStyles = makeStyles({
  width: { width: '400px' }, hover: { ':hover': { width: '400px' } },
  scrollbarWidth: { scrollbarWidth: 'none' },
  scrollbarColor: { scrollbarColor: 'red transparent' },
  scrollbarHover: { ':hover': { scrollbarWidth: 'none', scrollbarColor: 'red transparent' } },
  active: { ':active': { width: '400px' } }, focus: { ':focus-visible': { width: '400px' } },
  viewport: { '@media (min-width: 640px)': { width: '400px' } },
  container: { '@container apvee-sx (min-inline-size: 640px)': { width: '400px' } },
  viewportHover: { '@media (min-width: 640px)': { ':hover': { width: '400px' } } },
  containerHover: { '@container apvee-sx (min-inline-size: 640px)': { ':hover': { width: '400px' } } }
});
const renderer = createDOMRenderer(document);
let config: FixtureConfig = {
  size: 639, reverse: new URLSearchParams(location.search).get('reverse') === '1',
  scopes: true, enabled: true, selected: false, dir: 'ltr', theme: 'light', scrollbarEnabled: true
};

function Fixture(props: FixtureConfig): React.ReactElement {
  const sx = useSx();
  const native = externalStyles({ renderer: useRenderer_unstable(), dir: props.dir });
  const external = native.width;
  const scrollbarInterop = [
    { name: 'width', external: native.scrollbarWidth, descriptor: scrollbar.fluent },
    { name: 'color', external: native.scrollbarColor, descriptor: scrollbar.fluent },
    { name: 'hover', external: native.scrollbarHover, descriptor: hover(scrollbar.fluent) }
  ];
  const ordered = (inputs: readonly SxInput[]): readonly SxInput[] => props.reverse ? [...inputs].reverse() : inputs;
  const probe = (id: string, inputs: readonly SxInput[], children?: React.ReactNode): React.ReactElement =>
    <div key={id} id={id} className={sx('probe', ...inputs)}>{children || id}</div>;
  const host = (id: string, size: number, children: React.ReactNode, vertical = false): React.ReactElement =>
    <div key={id} id={id} className={sx('host', vertical && 'vertical', container.inlineSize)}
      style={vertical ? { height: size } : { width: size }}>{children}</div>;
  const directions = [column, responsive.medium(row)];
  const directionChildren = <><span data-part="first">First flex item</span><span data-part="second">Second flex item</span></>;
  const widths = [width.px(100), responsive.small(width.px(180)), responsive.medium(width.px(240)), responsive.large(width.px(360))];
  const responsiveNodes = [
    probe('direction', ordered(directions), directionChildren),
    probe('responsive-width', ordered(widths)),
    probe('responsive-gap', ordered([column, gap(4), responsive.medium(gap(20)), responsive.large(gap(32))])),
    probe('viewport', [width.px(120), viewport.medium(width.px(420))]),
    probe('priority', ordered([width.px(100), viewport.medium(width.px(250)), responsive.medium(width.px(260))])),
    probe('scope-removed', [width.px(111), props.scopes && responsive.medium(width.px(222)), props.scopes && hover(width.px(333))])
  ];
  const separateBase = sx(width.px(360));
  const separateHover = sx(hover(width.px(240)));
  const separateQuery = sx(responsive.medium(width.px(280)));
  const separateQueryHover = sx(responsive.medium(hover(width.px(320))));
  const mergeOrdered = (classes: readonly string[]): string => mergeClasses(...(props.reverse ? [...classes].reverse() : classes));
  const stateWidths = [
    width.px(100), hover(width.px(110)), active(width.px(120)), focusVisible(width.px(130)),
    responsive.small(hover(width.px(140)), active(width.px(150)), focusVisible(width.px(160))),
    viewport.medium(hover(width.px(170)), active(width.px(180)), focusVisible(width.px(190))),
    responsive.medium(hover(width.px(200)), active(width.px(210)), focusVisible(width.px(220))),
    viewport.large(hover(width.px(230)), active(width.px(240)), focusVisible(width.px(250))),
    responsive.large(hover(width.px(260)), active(width.px(270)), focusVisible(width.px(280)))
  ];
  const foreignScopes: readonly [string, string, SxInput][] = [
    ['hover', native.hover, hover(width.px(360))],
    ['viewport', native.viewport, viewport.medium(width.px(360))],
    ['container', native.container, responsive.medium(width.px(360))],
    ['viewport-hover', native.viewportHover, viewport.medium(hover(width.px(360)))],
    ['container-hover', native.containerHover, responsive.medium(hover(width.px(360)))]
  ];
  const foreignNodes = foreignScopes.flatMap(([name, className, descriptor]) => [
    probe(`foreign-${name}-sx-last`, [external, className, descriptor]),
    probe(`foreign-${name}-native-last`, [external, descriptor, className]),
    probe(`foreign-${name}-base-preserved`, [external, descriptor])
  ]);
  return <>
    <h1>useSx real-browser checkpoint</h1>
    <p className="legend">Real React 17, Griffel and FluentProvider. Geometry styles belong only to fixture hosts.</p>
    <a id="focus-start" href="#states">Keyboard focus start</a>
    <section className="case"><h2>Responsive thresholds and independent regions</h2>
      {host('main-container', props.size, props.reverse ? [...responsiveNodes].reverse() : responsiveNodes)}
      {host('region-small', 639, probe('region-small-child', directions, directionChildren))}
      {host('region-medium', 640, probe('region-medium-child', directions, directionChildren))}
      {host('outer-container', 1050, <>{probe('outer-child', widths)}{host('inner-container', 479, probe('inner-child', widths))}</>)}
      {probe('outside-container', [width.px(100), responsive.medium(width.px(240))])}
      {probe('self-container', [container.inlineSize, width.px(1024), responsive.medium(width.px(111))])}
      {host('vertical-container', props.size, probe('vertical-child', directions), true)}
    </section>
    <section className="case"><h2>Isolation and independent sx strings</h2>
      {host('isolation-container', 1050,
        probe('parent', [width.px(500), responsive.medium(width.px(550)), hover(width.px(600)), hover(foreground.subtle)],
          probe('own-child', [width.px(75), declaration('color', 'var(--colorNeutralForeground1)', 'inherit')])))}
      <div className="inherited" id="inherit-parent">{probe('omitted-color', [width.px(75)])}</div>
      <div id="separate" className={mergeOrdered([separateBase, separateHover])}>Separate base and hover</div>
      <div id="separate-strings" className={sx(...(props.reverse ? [separateHover, separateBase] : [separateBase, separateHover]))}>Strings through sx</div>
      {host('separate-container', props.size,
        <div id="separate-query" className={mergeOrdered([separateBase, separateQuery, separateQueryHover])}>Separate container scopes</div>)}
    </section>
    <section className="case"><h2>External Griffel composition</h2>
      {probe('external-intercalated', [width.px(240), external, gap(12)])}
      {probe('external-later-descriptor', [width.px(240), external, gap(12), width.px(360)])}
      <div id="external-last" className={mergeClasses(sx(width.px(240)), external)}>External last</div>
      <div id="descriptor-last" className={mergeClasses(external, sx(width.px(240)))}>Descriptor last</div>
      {host('foreign-container', props.size, props.reverse ? [...foreignNodes].reverse() : foreignNodes)}
    </section>
    <section className="case" id="states"><h2>Keyboard focus, hover, active, disabled and selected</h2>
      <button id="state-base" className={sx(neutral, hover(red), active(green), focusVisible(blue))}>State base</button>
      {host('state-container', props.size,
        <button id="state-query" className={sx(...ordered([neutral, hover(red), active(green), focusVisible(blue),
          viewport.medium(hover(background('rgb(80, 80, 0)'))), responsive.medium(hover(background('rgb(150, 0, 0)'))),
          responsive.medium(active(background('rgb(0, 100, 0)'))), responsive.medium(focusVisible(purple))]))}>Query state</button>)}
      {host('state-priority-container', props.size,
        <button id="state-priority" className={sx(...ordered(stateWidths))}>State responsive priority</button>)}
      <button id="viewport-state" className={sx(neutral, hover(red), viewport.medium(hover(green)))}>Viewport state without query container</button>
      <button id="transition" disabled={!props.enabled} aria-pressed={props.selected}
        className={sx(props.enabled ? props.selected ? selected : neutral : disabled,
          props.enabled && !props.selected && hover(red), props.enabled && !props.selected && active(green),
          props.enabled && focusVisible(blue))}>Transition</button>
      <button id="foreign-separate-states" className={mergeOrdered(['foreign-state-probe',
        sx(width.px(100)), sx(hover(width.px(110))), sx(active(width.px(120))), sx(focusVisible(width.px(130)))
      ])}>Separate state strings</button>
      <button id="foreign-state-sx-last" className={mergeClasses('foreign-state-probe', sx(width.px(100)), native.hover, native.active, native.focus,
        sx(hover(width.px(360)), active(width.px(360)), focusVisible(width.px(360))))}>Native states before sx</button>
      <button id="foreign-state-native-last" className={mergeClasses('foreign-state-probe', sx(width.px(100)),
        sx(hover(width.px(360)), active(width.px(360)), focusVisible(width.px(360))), native.hover, native.active, native.focus)}>Native states after sx</button>
    </section>
    <section className="case" id="catalog"><h2>Fluent catalog, theme colors, and scrollbar appearance</h2>
      {(['canvas', 'alternative'] as const).map(surface => <div key={surface} id={`catalog-${surface}`}
        className={sx('catalog-surface', catalogBackground[surface], foreground.primary)}>
        <h3>{surface} underlying surface ({props.theme})</h3>
        {props.theme === 'highContrast' && <p>Known floor limitation: inverted below is black on black (1:1), unsupported for active content.</p>}
        <p>Disabled pair below demonstrates disabled presentation only.</p>
        {Object.entries(presets).map(([name, preset]) => <div key={name} id={`preset-${surface}-${name}`}
          className={sx('catalog-preset', preset, typography.body1)}>{name}: readable sample text</div>)}
        <div id={`scrollbar-${surface}`} className={sx('catalog-scrollbox', catalogBackground[surface], overflow.auto, scrollbar.fluent)}>
          <div className="catalog-scroll-content">Scrollable content: scrollbar styling does not select overflow or geometry.</div>
        </div>
      </div>)}
      <div id="catalog-inheritance-parent" className="inherited">
        <div id="catalog-typography-inherits" className={sx(typography.body1)}>Typography inherits ordinary parent color</div>
        <div id="catalog-transparent-inherits" className={sx(presets.transparent)}>Transparent preset inherits ordinary parent color</div>
      </div>
      <div id="catalog-preset-override" className={sx(presets.canvas, foreground.subtle)}>Property-wise foreground override</div>
      <div id="catalog-preset-reverse" className={sx(foreground.subtle, presets.canvas)}>Property-wise preset override</div>
      <div id="scrollbar-only" className={sx(scrollbar.fluent)}>Scrollbar appearance without overflow</div>
      <div id="scrollbar-interop">
        {scrollbarInterop.map(item => (['native-last', 'recipe-last'] as const).map(order =>
          (['segmented', 'merged'] as const).map(composition => {
            const id = `scrollbar-interop-${item.name}-${order}-${composition}`;
            const inputs = order === 'native-last' ? [item.descriptor, item.external] : [item.external, item.descriptor];
            const className = composition === 'segmented' ? sx('probe', ...inputs)
              : mergeClasses('probe', ...inputs.map(input => typeof input === 'string' ? input : sx(input)));
            return <div key={id} id={id} className={className}>{id}</div>;
          })))}
      </div>
      <div id="scrollbar-removable" className={sx('catalog-scrollbox', overflow.auto, props.scrollbarEnabled && scrollbar.fluent)}>
        <div className="catalog-scroll-content">Removing the recipe restores native scrollbar appearance.</div>
      </div>
      <div className={sx('scrollbar-scope-inactive', container.inlineSize, scrollbar.fluent)}>
        {probe('scrollbar-scoped-inactive', [responsive.medium(scrollbar.fluent)])}
      </div>
      <div className={sx('scrollbar-scope-active', container.inlineSize)}>
        {probe('scrollbar-scoped-query', [responsive.medium(scrollbar.fluent)])}
        {probe('scrollbar-scoped-query-hover', [responsive.medium(hover(scrollbar.fluent))])}
      </div>
      <div id="catalog-typography-body1" className={sx(typography.body1)}>Body1 theme font tokens</div>
    </section>
    <section className="case"><h2>Direction and logical longhands</h2>
      {probe('logical', [width.px(75), declaration('paddingInlineStart', '13px', '0px')])}
      {probe('physical', [width.px(75), declaration('paddingLeft', '17px', '0px')])}
    </section>
  </>;
}
function render(patch: Partial<FixtureConfig>): void {
  config = { ...config, ...patch };
  ReactDOM.render(<RendererProvider renderer={renderer}>
    <FluentProvider theme={config.theme === 'dark' ? webDarkTheme : config.theme === 'highContrast' ? teamsHighContrastTheme : webLightTheme} dir={config.dir}><Fixture {...config} /></FluentProvider>
  </RendererProvider>, document.getElementById('root'));
}
window.sxFixture = { render, get config() { return config; } };
render({});
