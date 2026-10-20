/** Regression snapshot captured before single-subject consolidation; not product UI. */
import * as React from 'react';
import { Button, Checkbox, Field, Input, Select, makeStyles } from '@fluentui/react-components';
import { FluentProvider } from '@fluentui/react-provider';
import { useSx, useSPFxFluent9ThemeInfo, SxBaseDescriptor, width, height, minWidth, maxWidth, padding, paddingInlineStart, gap, grid, flex, container, foreground, background, presets, typography, borderWidth, borderStyle, borderColor, overflow, scrollbar, responsive, viewport, hover, active, focusVisible, } from '@apvee/spfx-react-toolkit';
import { DemoCard } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/DemoCard';
import { useDemoStyles } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/shared/demoStyles';
import { stylesCatalog } from '../../../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/panels/stylesCatalog';
// The peer Button animates colors. Disable that transition in this recipe probe
// so hover/active/focus can be compared immediately without an interpolated color.
const useStateProbeStyles = makeStyles({ action: { transitionProperty: 'none' }, scroll: { boxSizing: 'border-box' } });

type Threshold = 'small' | 'medium' | 'large';
interface SelectProps {
  readonly id: string;
  readonly label: string;
  readonly value: string;
  readonly choices: readonly string[];
  readonly onChange: (value: string) => void;
}
const Choice: React.FC<SelectProps> = props => (<Field label={props.label}>
  <Select select={{ as: 'select', ...{ 'data-sx-demo': props.id } }} value={props.value} onChange={(_, data) => props.onChange(data.value)}>
    {props.choices.map(choice => <option key={choice} value={choice}>{choice}</option>)}
  </Select>
</Field>);
interface NumberProps {
  readonly id: string;
  readonly label: string;
  readonly value: number;
  readonly minimum?: number;
  readonly integer?: boolean;
  readonly onChange: (value: number) => void;
}
const NumberChoice: React.FC<NumberProps> = props => {
  const [draft, setDraft] = React.useState(String(props.value));
  React.useEffect(() => setDraft(String(props.value)), [props.value, props.minimum, props.integer]);
  const isValid = (text: string): boolean => text.trim().length > 0 && Number.isFinite(Number(text))
    && Number(text) >= (props.minimum ?? 0) && (!props.integer || Number.isInteger(Number(text)));
  return <Field label={props.label}>
    <Input input={{ as: 'input', ...{ 'data-sx-demo': props.id } }} type="number" min={props.minimum ?? 0} step={props.integer ? 1 : 'any'} value={draft} onChange={(_, data) => {
      setDraft(data.value);
      if (isValid(data.value))
        props.onChange(Number(data.value));
    }} onBlur={() => {
      if (!isValid(draft))
        setDraft(String(props.value));
    }} />
  </Field>;
};
const keys = (values: object): string[] => Object.keys(values);
/** useSx deliberately runs beneath the active sample FluentProvider. */
const StylesExamples: React.FC<{
  readonly highContrast: boolean;
}> = ({ highContrast }) => {
  const sx = useSx();
  const stateStyles = useStateProbeStyles();
  const demoStyles = useDemoStyles();
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
  const frameClass = sx(padding.medium, borderWidth.thin, borderStyle.solid, borderColor.primary);
  const query = (id: string, size: number): React.ReactElement => (<div data-sx-demo={id} className={sx(container.inlineSize, width.px(size))}>
    <div data-sx-demo={`${id}-preview`} className={sx(flex.column, gap.small, presets.alternative, padding.small, width.px(180), responsive[threshold](width.px(300), flex.row))}>
      <span>First item</span>
      <span>Second item</span>
    </div>
  </div>);
  return <div className={sx(demoStyles.root, typography.body1)}>
    <DemoCard title="Layout and roles" iconName="Color" controls={<>
      <Choice id="preset" label="Preset" value={preset} choices={keys(presets)} onChange={value => {
        if (value in presets)
          setPreset(value as keyof typeof presets);
      }} />
      <>
        <NumberChoice id="layout-width" label="Layout width (px)" value={layoutWidth} onChange={setLayoutWidth} />
        <NumberChoice id="layout-gap" label="Layout gap (px)" value={layoutGap} onChange={setLayoutGap} />
        <NumberChoice id="layout-columns" label="Grid columns" value={columns} minimum={1} integer onChange={setColumns} />
      </>
      <>
        <Choice id="foreground" label="Foreground role" value={foregroundRole} choices={keys(foreground)} onChange={value => {
          if (value in foreground)
            setForegroundRole(value as keyof typeof foreground);
        }} />
        <Choice id="background" label="Background role" value={backgroundRole} choices={keys(background)} onChange={value => {
          if (value in background)
            setBackgroundRole(value as keyof typeof background);
        }} />
      </>
    </>} status={<p data-sx-demo="preset-notice" role="status">{presetUnsupported
      ? 'Inverted is unsupported for active content in the installed Fluent high contrast theme (1:1). Canvas is shown instead.'
      : 'Presets add color only. Disabled colors are for disabled presentation.'}</p>}>

      <p>Change real descriptor inputs. Values are pixels; columns are positive integers.</p>
      <div data-sx-demo="layout-scroll" role="region" aria-label="Layout preview" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.horizontal.auto)}>
        <div data-sx-demo="layout-preview" className={sx(grid.columns(columns), width.px(layoutWidth), gap.px(layoutGap))}>
          {[1, 2, 3, 4].map(value => <div key={value} className={sx(presets.alternative, padding.small)}>Item {value}</div>)}
        </div>
      </div>
      <p data-sx-demo="role-preview" aria-disabled={foregroundRole === 'disabled' || backgroundRole === 'disabled'} className={sx(padding.medium, background[backgroundRole], foreground[foregroundRole])}>
        Foreground and background are independent roles. Subtle background and alternative background use different tokens.
      </p>
      <div data-sx-demo="preset-preview" data-unsupported={presetUnsupported ? true : undefined} aria-disabled={preset === 'disabled'} className={sx(padding.medium, presetUnsupported ? presets.canvas : presets[preset])}>
        {preset === 'disabled' ? 'Disabled presentation only; this preset adds no native disabled behavior.' : 'Preset background and foreground compose property by property.'}
      </div>
    </DemoCard>

    <DemoCard title="Every catalog selection" iconName="Color" controls={<>
      <>
        <Choice id="catalog-family" label="Catalog namespace" value={family.name} choices={stylesCatalog.map(item => item.name)} onChange={value => {
          const index = stylesCatalog.findIndex(item => item.name === value);
          if (index >= 0) {
            setFamilyIndex(index);
            setMemberName(keys(stylesCatalog[index].members)[0]);
            if (value === 'grid' && (!Number.isInteger(numericValue) || numericValue < 1))
              setNumericValue(1);
          }
        }} />
        <Choice id="catalog-member" label="Catalog member" value={memberName} choices={keys(family.members)} onChange={setMemberName} />
        <NumberChoice id="catalog-value" label="Factory value (px or columns)" value={numericValue} minimum={family.name === 'grid' ? 1 : 0} integer={family.name === 'grid'} onChange={setNumericValue} />
      </>
    </>} status={<>
      <p aria-live="polite">Applied: {family.name}.{memberName}{typeof selection === 'function' ? `(${factoryValue})` : ''}</p>
      {catalogUnsupported && <p role="status">Inverted is unsupported for active content in this high contrast theme (1:1); canvas is shown.</p>}
      {catalogDisabled && <p>Preview represents disabled presentation.</p>}
    </>}>
      <p>Choose a namespace, member, and numeric factory value. The chosen public descriptor is applied to the preview.</p>
      <div data-sx-demo="catalog-scroll" role="region" aria-label="Catalog preview" tabIndex={0} className={sx(frameClass, height.px(180), flex.row, overflow.auto)}>
        <div data-sx-demo="catalog-preview" aria-disabled={catalogDisabled} data-unsupported={catalogUnsupported ? true : undefined} className={sx(minWidth.zero, width.px(260), borderWidth.thin, borderStyle.solid, borderColor.primary, family.name === 'alignItems' || family.name === 'justifyContent' ? flex.row : false, family.name === 'scrollbar' || family.name.indexOf('overflow') === 0 ? height.px(100) : false, family.name === 'scrollbar' && overflow.auto, catalogUnsupported ? presets.canvas : descriptor)}>
          <span>Catalog preview with enough text to observe wrapping and truncation. </span>
          <span>Second child. </span>
          <span>Third child. </span>
          {(family.name === 'scrollbar' || family.name.indexOf('overflow') === 0) && <div className={sx(height.px(200), width.px(480))}>Overflow content</div>}
        </div>
      </div>
    </DemoCard>

    <DemoCard title="Interaction and direction" iconName="Color" controls={<>
      <>
        <Checkbox input={{ as: 'input', ...{ 'data-sx-demo': 'disabled' } }} label="Disabled action" checked={disabled} onChange={(_, data) => setDisabled(data.checked === true)} />
        <Checkbox input={{ as: 'input', ...{ 'data-sx-demo': 'selected' } }} label="Selected action" checked={selected} onChange={(_, data) => setSelected(data.checked === true)} />
      </>
    </>} status={<p aria-live="polite">Action invocations: {invocations}</p>}>
      <Button data-sx-demo="state-preview" type="button" disabled={disabled} aria-pressed={selected} className={sx(stateStyles.action, padding.medium, borderWidth.thin, borderStyle.solid, borderColor.primary, disabled ? presets.disabled : presets.canvas, !disabled && selected && presets.brandTint, !disabled && hover(presets.alternative), !disabled && active(presets.brand), !disabled && focusVisible(borderColor.brand))} onClick={() => setInvocations(count => count + 1)}>State action</Button>
      <p>Use pointer, keyboard Tab, and Space. Hover, active and focus-visible are simultaneous scopes. Fluent retains its own combined pressed treatment and focus indicator; the active descriptor remains applied in its computed scope.</p>
      <div data-sx-demo="direction-preview" className={sx(frameClass, paddingInlineStart.px(32))}>Logical inline-start padding follows the active provider direction.</div>
    </DemoCard>

    <DemoCard title="Independent container and viewport queries" iconName="Color" controls={<>
      <>
        <Choice id="responsive-threshold" label="Container threshold" value={threshold} choices={['small', 'medium', 'large']} onChange={value => {
          if (value === 'small' || value === 'medium' || value === 'large')
            setThreshold(value);
        }} />
        <NumberChoice id="query-a-width" label="Region A width (px)" value={queryAWidth} onChange={setQueryAWidth} />
        <NumberChoice id="query-b-width" label="Region B width (px)" value={queryBWidth} onChange={setQueryBWidth} />
        <Choice id="viewport-threshold" label="Viewport threshold" value={viewportThreshold} choices={['small', 'medium', 'large']} onChange={value => {
          if (value === 'small' || value === 'medium' || value === 'large')
            setViewportThreshold(value);
        }} />
      </>
    </>}>
      <p>Each child queries its nearest apvee-sx ancestor. The two regions resize independently; viewport queries follow the browser width.</p>
      <div data-sx-demo="query-scroll" role="region" aria-label="Container query previews" tabIndex={0} className={sx(width.full, minWidth.zero, overflow.horizontal.auto)}>
        {query('query-a', queryAWidth)}{query('query-b', queryBWidth)}
        <div data-sx-demo="viewport-preview" className={sx(width.px(180), viewport[viewportThreshold](width.px(300)), presets.alternative, padding.small)}>Viewport query</div>
      </div>
    </DemoCard>

    <DemoCard title="The single Fluent scrollbar recipe" iconName="Color" description="Explicit overflow and dimensions demonstrate the Fluent scrollbar recipe.">
      <p>Both axes use 6px rounded scrollbars in Edge/Chrome. Hover or drag either thumb to see its Fluent feedback, with native thin fallback elsewhere. Overflow and dimensions are explicit; forced colors retain system colors.</p>
      <div data-sx-demo="scroll-preview" tabIndex={0} aria-label="Scrollable recipe region" className={sx(stateStyles.scroll, height.px(140), width.full, maxWidth.px(320), overflow.auto, scrollbar.fluent, presets.alternative, padding.small)}>
        <div className={sx(width.px(600))}>
          {Array.from({ length: 14 }, (_, index) => <p key={index}>Scrollable content row {index + 1}</p>)}
        </div>
      </div>
    </DemoCard>
  </div>;
};
const StylesPanel: React.FC = () => {
  const host = useSPFxFluent9ThemeInfo();
  const [dir, setDir] = React.useState<'host' | 'ltr' | 'rtl'>('host');
  const directionSx = useSx({ dir: dir === 'host' ? undefined : dir });
  const hostThemeName = host.teamsTheme?.toLowerCase();
  const highContrast = hostThemeName === 'contrast' || hostThemeName === 'highcontrast';
  return <>
    <DemoCard title="Styles" iconName="Color" description="Compose descriptors using the host theme and renderer." contentPlacement="configuration" controls={<>
      <Choice id="direction" label="Provider direction" value={dir} choices={['host', 'ltr', 'rtl']} onChange={value => {
        if (value === 'host' || value === 'ltr' || value === 'rtl')
          setDir(value);
      }} />
    </>} />
    <FluentProvider data-sx-demo="provider" dir={dir === 'host' ? undefined : dir} className={directionSx(background.transparent, minWidth.zero)}>
      <StylesExamples highContrast={highContrast} />
    </FluentProvider>
  </>;
};
export default StylesPanel;
