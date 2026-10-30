import * as React from 'react';
import { Card, Switch, cardCSSVars, makeStyles } from '@fluentui/react-components';
import { SettingsGroup } from '../../../packages/spfx-react-toolkit/src/controls/settings-group/settings-group';
import { AccordionGroup } from '../../../packages/spfx-react-toolkit/src/controls/accordion-group/accordion-group';
import { useSx } from '../../../packages/spfx-react-toolkit/src/hooks/useSx';
import * as foreground from '../../../packages/spfx-react-toolkit/src/helpers/styles/foreground';
import * as width from '../../../packages/spfx-react-toolkit/src/helpers/styles/width';
import * as typography from '../../../packages/spfx-react-toolkit/src/helpers/styles/typography';

interface Config { contentWidth: number; long: boolean; hidden: boolean; cardSizeOverride: boolean; }
let config: Config = { contentWidth: 639, long: false, hidden: false, cardSizeOverride: false };
const useFixtureStyles = makeStyles({ cardSizeOverride: { '& > .fui-Card': { [cardCSSVars.cardSizeVar]: '16px' } } });
let update: ((value: Config) => void) | undefined;
declare global { interface Window { groupsFixture: { render(patch: Partial<Config>): void; }; } }
window.groupsFixture = { render(patch) { config = { ...config, ...patch }; update?.(config); } };

function StatefulInput(props: { id: string }): React.ReactElement {
  const [value, setValue] = React.useState('Initial');
  return <input id={props.id} aria-label={props.id} value={value} onChange={event => setValue(event.currentTarget.value)} />;
}

/** Real source controls inside the existing Fluent/renderer/direction host. */
export function GroupsFixture(): React.ReactElement {
  const [value, setValue] = React.useState(config);
  const [clicks, setClicks] = React.useState(0);
  const [nativeEvents, setNativeEvents] = React.useState(0);
  const [hasRef, setHasRef] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);
  const sx = useSx();
  const styles = useFixtureStyles();
  React.useEffect(() => { update = setValue; setHasRef(ref.current instanceof HTMLDivElement); return () => { update = undefined; }; }, []);
  const title = value.long ? 'Long **group title** with naturally wrapping content '.repeat(4) : 'Group **title**';
  const description = value.long ? 'Averylongunbrokentextvaluethatmuststaywithinthegroupboundary'.repeat(8) : 'A *subtle description*';
  const groupWidth = value.contentWidth;
  const icon = <svg width="1em" height="1em" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M4 5h16v14H4zm2 2v10h12V7z" /></svg>;
  return <section id="groups-fixture" aria-label="Groups browser checks">
    <h2>SettingsGroup and AccordionGroup browser fixture</h2>
    <output id="groups-ref">{String(hasRef)}</output><output id="groups-clicks">{clicks}</output><output id="groups-native-events">{nativeEvents}</output>
    <SettingsGroup id="groups-settings" ref={ref} title={title} description={description} hidden={value.hidden}
      className={sx(width.px(groupWidth), value.cardSizeOverride && styles.cardSizeOverride)} onDoubleClick={() => setNativeEvents(count => count + 1)}
      action={<button id="groups-action" onClick={() => setClicks(count => count + 1)}>{value.long ? 'Group action with a longer naturally wrapping label' : 'Group action'}</button>}>
      <SettingsGroup.Item id="groups-setting-row" title="Item **title**" description={description}
        action={<button id="groups-item-action">{value.long ? 'Item action with a longer naturally wrapping label' : 'Item action'}</button>}>
        <label htmlFor="groups-setting-input">{value.long ? 'Caller label with additional words for natural wrapping' : 'Caller label'}</label><StatefulInput id="groups-setting-input" />
      </SettingsGroup.Item>
      <SettingsGroup.Item id="groups-setting-second" title="Second setting"><button>Second control</button></SettingsGroup.Item>
    </SettingsGroup>
    <SettingsGroup id="groups-independent" title="Independent container" className={sx(width.px(640))}>
      <SettingsGroup.Item id="groups-independent-row" title="Independent item"><input aria-label="Independent input" /></SettingsGroup.Item>
    </SettingsGroup>
    <Card id="groups-card-reference" size="medium" className={sx(width.px(groupWidth))}><span id="groups-card-content">Native medium Card inset</span></Card>
    <span id="groups-subtitle-reference" className={sx(typography.subtitle2, foreground.primary)}>Reference</span>
    <span id="groups-body-reference" className={sx(typography.body1Strong, foreground.primary)}>Reference</span>
    <span id="groups-caption-reference" className={sx(typography.caption1, foreground.subtle)}>Reference</span>
    <AccordionGroup id="groups-accordion" title="Accordion **title**" description="Accordion *description*"
      className={sx(width.px(groupWidth), value.cardSizeOverride && styles.cardSizeOverride)} defaultOpenItems={['persistent', 'reset']} multiple collapsible>
      <AccordionGroup.Item id="groups-persistent" value="persistent" title="Persistent **item**" description="Retained state">
        <StatefulInput id="groups-persistent-input" /><button id="groups-persistent-button">Panel button</button>
      </AccordionGroup.Item>
      <AccordionGroup.Item id="groups-reset" value="reset" title="Reset item" unmountOnClose>
        <StatefulInput id="groups-reset-input" />
      </AccordionGroup.Item>
      <AccordionGroup.Item id="groups-disabled" value="disabled" title="Disabled item" description="Disabled description" disabled>Disabled content</AccordionGroup.Item>
    </AccordionGroup>
    <button id="groups-after">After accordion</button>
    <AccordionGroup id="groups-hidden" hidden title="Hidden accordion"><AccordionGroup.Item value="hidden" title="Hidden item">Hidden body</AccordionGroup.Item></AccordionGroup>
    <SettingsGroup id="groups-opaque" title={<span id="groups-authored-jsx">**Authored JSX**</span>}
      description={<span id="groups-jsx-description">Inherited description</span>} className={sx(width.px(groupWidth))}>
      <SettingsGroup.Item id="groups-jsx-row" title={<em id="groups-jsx-item">Caller JSX title</em>}><button>Caller child</button></SettingsGroup.Item>
    </SettingsGroup>
    <SettingsGroup id="groups-fluid" title="Fluid settings" description="Library icon defaults" icon={icon}
      className={sx(width.px(groupWidth), value.cardSizeOverride && styles.cardSizeOverride)}>
      <SettingsGroup.Item id="groups-fluid-switch" title="Account alerts" description="Receive updates" icon={icon}>
        <Switch aria-label="Fluid account alerts" />
      </SettingsGroup.Item>
      <SettingsGroup.Item id="groups-fluid-second" title="Account alerts" description="Receive updates" icon={icon}>
        <Switch aria-label="Second fluid account alerts" />
      </SettingsGroup.Item>
    </SettingsGroup>
    <AccordionGroup id="groups-symmetric" title="Symmetric disclosures" description="Equal text in closed native rows" icon={icon}
      className={sx(width.px(groupWidth), value.cardSizeOverride && styles.cardSizeOverride)}>
      <AccordionGroup.Item id="groups-symmetric-first" value="first" title="Account details" description="Supporting details" icon={icon}>First content</AccordionGroup.Item>
      <AccordionGroup.Item id="groups-symmetric-second" value="second" title="Account details" description="Supporting details" icon={icon}>Second content</AccordionGroup.Item>
    </AccordionGroup>
  </section>;
}
