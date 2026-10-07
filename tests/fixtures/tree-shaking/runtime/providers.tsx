import * as React from 'react';
import { SPFxWebPartProvider } from '__PROVIDER__';
import { useSPFxContext } from '__CONTEXT__';
import { useSPFxProperties } from '__PROPERTIES__';

// Only the unavailable SPFx host service is doubled; providers/stores/hooks are real.
export function run({ ReactDOM, act, assert }) {
  const observations = {}, controls = {};
  function instance(id, title) {
    const listeners = new Set();
    const theme = { tryGetTheme: () => undefined, themeChangedEvent: {
      add: (_, handler) => listeners.add(handler), remove: (_, handler) => listeners.delete(handler),
    }};
    return { properties: { title }, displayMode: 1, domElement: document.createElement('div'), render() {},
      context: { instanceId: id, serviceScope: { whenFinished: fn => fn(), consume: key => { assert.equal(key, 'production-theme'); return theme; } } }, listeners };
  }
  const first = instance('first', 'one'), second = instance('second', 'two');
  function Child({id}) {
    const context = useSPFxContext();
    const properties = useSPFxProperties();
    observations[id] = { id: context.instanceId, title: properties.properties?.title };
    controls[id] = properties.setProperties;
    return <span>{context.instanceId}:{properties.properties?.title}</span>;
  }
  const container = document.createElement('div');document.body.appendChild(container);
  const render = () => <><SPFxWebPartProvider instance={first}><Child id="first" /></SPFxWebPartProvider><SPFxWebPartProvider instance={second}><Child id="second" /></SPFxWebPartProvider></>;
  act(() => ReactDOM.render(render(), container));
  assert.deepEqual(observations, { first: { id:'first',title:'one' }, second: { id:'second',title:'two' } });
  assert.equal(first.listeners.size,1);assert.equal(second.listeners.size,1);
  act(() => controls.first({title:'changed'}));
  assert.equal(first.properties.title,'changed');assert.equal(second.properties.title,'two');
  second.properties.title='host-update';
  act(() => ReactDOM.render(render(), container));
  assert.equal(observations.first.title,'changed');assert.equal(observations.second.title,'host-update');
  const detachedSetter=controls.first;
  act(() => ReactDOM.unmountComponentAtNode(container));
  assert.equal(first.listeners.size,0);assert.equal(second.listeners.size,0);
  act(() => detachedSetter({title:'after-unmount'}));
  assert.equal(first.properties.title,'changed');container.remove();
  return {assertions:['mixed imports share canonical provider contexts', 'two provider stores isolated', 'hook and host updates synchronize', 'unmount removes subscriptions and property sync']};
}
