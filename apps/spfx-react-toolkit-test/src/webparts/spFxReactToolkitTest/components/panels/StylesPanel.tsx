import * as React from 'react';
import { FluentProvider } from '@fluentui/react-provider';
import {
  useSx, useSPFxFluent9ThemeInfo, getTeamsFluentTheme, SxBaseDescriptor,
  width, height, minWidth, padding, paddingBlockEnd, paddingInlineStart, gap, grid, flex, container,
  foreground, background, presets, typography, borderWidth, borderStyle, borderColor,
  overflow, scrollbar, responsive, viewport, hover, active, focusVisible,
} from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../shared/DemoCard';
import { stylesCatalog } from './stylesCatalog';

type ThemeChoice = 'host' | 'default' | 'dark' | 'contrast';
type Threshold = 'small' | 'medium' | 'large';
interface SelectProps {
  readonly id: string;
  readonly label: string;
  readonly value: string;
  readonly choices: readonly string[];
  readonly onChange: (value: string) => void;
}
const Choice: React.FC<SelectProps> = props => (
  <label>{props.label}{' '}
    <select data-sx-demo={props.id} value={props.value} onChange={event => props.onChange(event.target.value)}>
      {props.choices.map(choice => <option key={choice} value={choice}>{choice}</option>)}
    </select>
  </label>
);
interface NumberProps {
  readonly id: string;
  readonly label: string;
  readonly value: number;
  readonly minimum?: number;
  readonly integer?: boolean;
  readonly onChange: (value: number) => void;
}
const NumberChoice: React.FC<NumberProps> = props => (
  <label>{props.label}{' '}
    <input data-sx-demo={props.id} type="number" min={props.minimum || 0} step={props.integer ? 1 : 'any'} value={props.value}
      onChange={event => {
        if (!event.target.value.trim()) return;
        const value = Number(event.target.value);
        if (Number.isFinite(value) && value >= (props.minimum || 0) && (!props.integer || Number.isInteger(value))) props.onChange(value);
      }} />
  </label>
);
const keys = (values: object): string[] => Object.keys(values);

/** useSx deliberately runs beneath the active sample FluentProvider. */
const StylesExamples: React.FC<{ readonly highContrast: boolean }> = ({ highContrast }) => {
  const sx = useSx();
  const [layoutWidth, setLayoutWidth] = React.useState(240);
  const [layoutGap, setLayoutGap] = React.useState(12);
  const [columns, setColumns] = React.useState(2);
  const [foregroundRole, setForegroundRole] = React.useState<keyof typeof foreground>('primary');
  const [backgroundRole, setBackgroundRole] = React.useState<keyof typeof background>('canvas');
  const [preset, setPreset] = React.useState<keyof typeof presets>('canvas');
  const [familyIndex, setFamilyIndex] = React.useState(0);
  const family = stylesCatalog[familyIndex];
  const [memberName, setMemberName] = React.useState(keys(family.members)[0]);
  const [numericValue, setNumericValue] = React.useState(24);
  const [disabled, setDisabled] = React.useState(false);
  const [selected, setSelected] = React.useState(false);
  const [invocations, setInvocations] = React.useState(0);
  const [queryAWidth, setQueryAWidth] = React.useState(420);
  const [queryBWidth, setQueryBWidth] = React.useState(720);
  const [threshold, setThreshold] = React.useState<Threshold>('medium');
  const [viewportThreshold, setViewportThreshold] = React.useState<Threshold>('medium');
  const selection = family.members[memberName];
  const factoryValue = numericValue;
  const descriptor: SxBaseDescriptor = typeof selection === 'function' ? selection(factoryValue) : selection;
  const catalogUnsupported = highContrast && family.name === 'presets' && memberName === 'inverted';
  const catalogDisabled = memberName === 'disabled' && ['presets', 'foreground', 'background', 'borderColor'].indexOf(family.name) >= 0;
  const presetUnsupported = highContrast && preset === 'inverted';
  const controlsClass = sx(flex.row, flex.wrap, gap.medium, padding.small);
  const frameClass = sx(presets.canvas, padding.medium, borderWidth.thin, borderStyle.solid, borderColor.primary);
  const query = (id: string, size: number): React.ReactElement => (
    <div data-sx-demo={id} className={sx(container.inlineSize, width.px(size))}>
      <div data-sx-demo={`${id}-preview`} className={sx(flex.column, gap.small, presets.alternative, padding.small,
        width.px(180), responsive[threshold](width.px(300), flex.row))}>
        <span>First item</span><span>Second item</span>
      </div>
    </div>
  );

  return <div className={sx(presets.canvas, typography.body1, padding.medium)}>
    <DemoCard title="Layout and roles" iconName="Color">
      <p>Change real descriptor inputs. Values are pixels; columns are positive integers.</p>
      <div className={controlsClass}>
        <NumberChoice id="layout-width" label="Layout width (px)" value={layoutWidth} onChange={setLayoutWidth} />
        <NumberChoice id="layout-gap" label="Layout gap (px)" value={layoutGap} onChange={setLayoutGap} />
        <NumberChoice id="layout-columns" label="Grid columns" value={columns} minimum={1} integer onChange={setColumns} />
      </div>
      <div className={sx(width.full, minWidth.zero, overflow.horizontal.auto)}>
        <div data-sx-demo="layout-preview" className={sx(grid.columns(columns), width.px(layoutWidth), gap.px(layoutGap))}>
          {[1, 2, 3, 4].map(value => <div key={value} className={sx(presets.alternative, padding.small)}>Item {value}</div>)}
        </div>
      </div>
      <div className={controlsClass}>
        <Choice id="foreground" label="Foreground role" value={foregroundRole} choices={keys(foreground)}
          onChange={value => { if (value in foreground) setForegroundRole(value as keyof typeof foreground); }} />
        <Choice id="background" label="Background role" value={backgroundRole} choices={keys(background)}
          onChange={value => { if (value in background) setBackgroundRole(value as keyof typeof background); }} />
      </div>
      <p data-sx-demo="role-preview" aria-disabled={foregroundRole === 'disabled' || backgroundRole === 'disabled'}
        className={sx(padding.medium, background[backgroundRole], foreground[foregroundRole])}>
        Foreground and background are independent roles. Subtle background and alternative background use different tokens.
      </p>
      <Choice id="preset" label="Preset" value={preset} choices={keys(presets)}
        onChange={value => { if (value in presets) setPreset(value as keyof typeof presets); }} />
      <div data-sx-demo="preset-preview" data-unsupported={presetUnsupported ? true : undefined}
        aria-disabled={preset === 'disabled'} className={sx(padding.medium, presetUnsupported ? presets.canvas : presets[preset])}>
        {preset === 'disabled' ? 'Disabled presentation only; this preset adds no native disabled behavior.' : 'Preset background and foreground compose property by property.'}
      </div>
      <p data-sx-demo="preset-notice" role="status">{presetUnsupported
        ? 'Inverted is unsupported for active content in the installed Fluent high contrast theme (1:1). Canvas is shown instead.'
        : 'Presets add color only. Disabled colors are for disabled presentation.'}</p>
    </DemoCard>

    <DemoCard title="Every catalog selection" iconName="Color">
      <p>Choose a namespace, member, and numeric factory value. The chosen public descriptor is applied to the preview.</p>
      <div className={controlsClass}>
        <Choice id="catalog-family" label="Catalog namespace" value={family.name} choices={stylesCatalog.map(item => item.name)}
          onChange={value => {
            const index = stylesCatalog.findIndex(item => item.name === value);
            if (index >= 0) {
              setFamilyIndex(index); setMemberName(keys(stylesCatalog[index].members)[0]);
              if (value === 'grid' && (!Number.isInteger(numericValue) || numericValue < 1)) setNumericValue(1);
            }
          }} />
        <Choice id="catalog-member" label="Catalog member" value={memberName} choices={keys(family.members)} onChange={setMemberName} />
        <NumberChoice id="catalog-value" label="Factory value (px or columns)" value={numericValue} minimum={family.name === 'grid' ? 1 : 0} integer={family.name === 'grid'} onChange={setNumericValue} />
      </div>
      <p aria-live="polite">Applied: {family.name}.{memberName}{typeof selection === 'function' ? `(${factoryValue})` : ''}</p>
      <div className={sx(frameClass, height.px(180), flex.row, overflow.auto)}>
        <div data-sx-demo="catalog-preview" aria-disabled={catalogDisabled} data-unsupported={catalogUnsupported ? true : undefined}
          className={sx(minWidth.zero, width.px(260), borderWidth.thin, borderStyle.solid, borderColor.primary,
            family.name === 'alignItems' || family.name === 'justifyContent' ? flex.row : false,
            family.name === 'scrollbar' || family.name.indexOf('overflow') === 0 ? height.px(100) : false,
            family.name === 'scrollbar' && overflow.auto,
            catalogUnsupported ? presets.canvas : descriptor)}>
          <span>Catalog preview with enough text to observe wrapping and truncation. </span>
          <span>Second child. </span><span>Third child. </span>
          {(family.name === 'scrollbar' || family.name.indexOf('overflow') === 0) && <div className={sx(height.px(200), width.px(480))}>Overflow content</div>}
        </div>
      </div>
      {catalogUnsupported && <p role="status">Inverted is unsupported for active content in this high contrast theme (1:1); canvas is shown.</p>}
      {catalogDisabled && <p>Preview represents disabled presentation.</p>}
    </DemoCard>

    <DemoCard title="Interaction and direction" iconName="Color">
      <p>Use pointer, keyboard Tab, and Space. Hover, active and focus-visible are simultaneous scopes; native focus is retained.</p>
      <div className={controlsClass}>
        <label><input data-sx-demo="disabled" type="checkbox" checked={disabled} onChange={event => setDisabled(event.target.checked)} /> Disabled action</label>
        <label><input data-sx-demo="selected" type="checkbox" checked={selected} onChange={event => setSelected(event.target.checked)} /> Selected action</label>
      </div>
      <button data-sx-demo="state-preview" type="button" disabled={disabled} aria-pressed={selected}
        className={sx(padding.medium, borderWidth.thin, borderStyle.solid, borderColor.primary,
          disabled ? presets.disabled : presets.canvas,
          !disabled && selected && presets.brandTint,
          !disabled && hover(presets.alternative), !disabled && active(presets.brand),
          !disabled && focusVisible(borderColor.brand))}
        onClick={() => setInvocations(count => count + 1)}>Native state action</button>
      <p aria-live="polite">Action invocations: {invocations}</p>
      <div data-sx-demo="direction-preview" className={sx(frameClass, paddingInlineStart.px(32))}>Logical inline-start padding follows the active provider direction.</div>
    </DemoCard>

    <DemoCard title="Independent container and viewport queries" iconName="Color">
      <p>Each child queries its nearest apvee-sx ancestor. The two regions resize independently; viewport queries follow the browser width.</p>
      <div className={controlsClass}>
        <Choice id="responsive-threshold" label="Container threshold" value={threshold} choices={['small', 'medium', 'large']}
          onChange={value => { if (value === 'small' || value === 'medium' || value === 'large') setThreshold(value); }} />
        <NumberChoice id="query-a-width" label="Region A width (px)" value={queryAWidth} onChange={setQueryAWidth} />
        <NumberChoice id="query-b-width" label="Region B width (px)" value={queryBWidth} onChange={setQueryBWidth} />
        <Choice id="viewport-threshold" label="Viewport threshold" value={viewportThreshold} choices={['small', 'medium', 'large']}
          onChange={value => { if (value === 'small' || value === 'medium' || value === 'large') setViewportThreshold(value); }} />
      </div>
      <div className={sx(width.full, minWidth.zero, overflow.horizontal.auto)}>
        {query('query-a', queryAWidth)}{query('query-b', queryBWidth)}
        <div data-sx-demo="viewport-preview" className={sx(width.px(180), viewport[viewportThreshold](width.px(300)), presets.alternative, padding.small)}>Viewport query</div>
      </div>
    </DemoCard>

    <DemoCard title="The single Fluent scrollbar recipe" iconName="Color">
      <p>Overflow and height are explicit. Standard scrollbar support and platform overlay behavior determine appearance; forced colors retain system colors.</p>
      <div data-sx-demo="scroll-preview" tabIndex={0} aria-label="Scrollable recipe region"
        className={sx(height.px(140), overflow.vertical.auto, scrollbar.fluent, presets.alternative, padding.small)}>
        {Array.from({ length: 14 }, (_, index) => <p key={index}>Scrollable content row {index + 1}</p>)}
      </div>
    </DemoCard>
  </div>;
};

const StylesPanel: React.FC = () => {
  const sx = useSx();
  const host = useSPFxFluent9ThemeInfo();
  const [themeChoice, setThemeChoice] = React.useState<ThemeChoice>('host');
  const [dir, setDir] = React.useState<'ltr' | 'rtl'>('ltr');
  const theme = themeChoice === 'host' ? host.theme : getTeamsFluentTheme(themeChoice);
  const highContrast = themeChoice === 'contrast' || (themeChoice === 'host' && (host.teamsTheme === 'contrast' || host.teamsTheme === 'highcontrast'));
  return <>
    <div className={sx(flex.row, flex.wrap, gap.px(16), paddingBlockEnd.px(12))}>
      <Choice id="theme" label="Theme" value={themeChoice} choices={['host', 'default', 'dark', 'contrast']}
        onChange={value => { if (value === 'host' || value === 'default' || value === 'dark' || value === 'contrast') setThemeChoice(value); }} />
      <Choice id="direction" label="Provider direction" value={dir} choices={['ltr', 'rtl']}
        onChange={value => { if (value === 'ltr' || value === 'rtl') setDir(value); }} />
    </div>
    <FluentProvider data-sx-demo="provider" theme={theme} dir={dir}>
      <StylesExamples highContrast={highContrast} />
    </FluentProvider>
  </>;
};
export default StylesPanel;
