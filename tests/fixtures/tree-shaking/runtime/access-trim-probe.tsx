import * as React from 'react';
import { act } from 'react-dom/test-utils';
import { PageContext, SPPermission } from '@microsoft/sp-page-context';
import { SPHttpClient, MSGraphClientFactory } from '@microsoft/sp-http';

const web = 'https://tenant.test/sites/standalone';
const group = 'aaaaaaaa-0000-0000-0000-000000000002';
const otherGroup = 'aaaaaaaa-0000-0000-0000-000000000003';

// Only unavailable native host services and HTTP/Graph transports are fixtures.
// Provider, all hooks, controls, validation and request lifecycles are installed
// production package code. SPPermission is the installed native implementation.
export async function runTrimProbe(Control, Provider, family, { ReactDOM, assert, nativePermission }) {
  const calls = [], mounts = {}, cleanups = {}, renders = {}, commits = [];
  const deferred = () => { let resolve; const promise = new Promise(done => { resolve = done; }); return { promise, resolve }; };
  const firstState = { mode: 'pending', obsolete: deferred() };
  const secondState = { mode: 'denied', obsolete: deferred() };
  function client(state, id) {
    const response = (url, body) => {
      if (state.mode === 'error') throw new Error('native transport rejected');
      const grant = state.mode === 'granted';
      if (body) return { value: grant ? body.groupIds : [] };
      if (url.includes('/effectivebasepermissions')) return { High: '0', Low: grant ? '1' : '0' };
      if (url.includes('/currentuser')) return { Id: 17 };
      if (url.includes('/users')) return { value: grant ? [{ Id: 17 }] : [] };
      const selected = /sitegroups\((\d+)\)/.exec(url);
      if (selected) return { Id: Number(selected[1]), Title: 'Readers' };
      throw new Error('Unexpected native request: ' + url);
    };
    const send = async (url, body) => {
      calls.push({ id, url, body });
      if (state.mode === 'pending') {
        await state.obsolete.promise;
        // Each response is captured by the transport for this obsolete request.
        const grant = state.mode;
        state.mode = 'granted';
        try { return response(url, body); } finally { state.mode = grant; }
      }
      return response(url, body);
    };
    return {
      sp: { get: async (url, config, options) => {
        assert.equal(config, SPHttpClient.configurations.v1);
        assert.equal(options.headers.Accept, 'application/json;odata.metadata=none');
        return { ok: true, status: 200, json: async () => await send(url) };
      } },
      graph: { api: url => ({ get: () => send(url), post: body => send(url, body) }) },
    };
  }
  function instance(id, state, sharedClients) {
    const native = sharedClients ?? client(state, id);
    const page = { web: { absoluteUrl: web }, site: { absoluteUrl: web }, user: { loginName: 'same-user', email: 'same@test' }, aadInfo: { userId: 'aaaaaaaa-0000-0000-0000-000000000001' } };
    const theme = { tryGetTheme: () => undefined, themeChangedEvent: { add() {}, remove() {} } };
    const factory = { getClient: async version => { assert.equal(version, '3'); return native.graph; } };
    return { properties: {}, displayMode: 1, domElement: document.createElement('div'), render() {},
      context: { instanceId: id, serviceScope: { whenFinished: done => done(), consume: key => {
        if (key === PageContext.serviceKey) return page;
        if (key === SPHttpClient.serviceKey) return native.sp;
        if (key === MSGraphClientFactory.serviceKey) return factory;
        assert.equal(key, 'production-theme'); return theme;
      } } }, native, page };
  }
  const first = instance('first', firstState), second = instance('second', secondState);
  const props = target => family === 'permission'
    ? { permissions: [SPPermission.viewListItems], scope: 'list', listTitle: target === 'first' ? 'Tasks' : 'Other' }
    : family === 'sharepoint-group' ? { groups: [{ id: target === 'first' ? 2 : 3 }] }
      : { groups: [{ id: target === 'first' ? group : otherGroup }], groupType: 'All' };
  function Branch({ role, text = role }) {
    renders[role] = (renders[role] ?? 0) + 1;
    React.useEffect(() => { mounts[role] = (mounts[role] ?? 0) + 1; return () => { cleanups[role] = (cleanups[role] ?? 0) + 1; }; }, [role]);
    return <span data-branch={role}>{text}</span>;
  }
  const branches = label => <React.Fragment key="granted">
    <Branch role="granted" text={label ?? 'granted'} />
  </React.Fragment>;
  // Markers are direct children, never wrapped in a Fragment.
  const content = label => [branches(label),
    <Control.Loading key="loading"><Branch role="loading" /></Control.Loading>,
    <Control.Fallback key="denied"><Branch role="denied" /></Control.Fallback>,
    <Control.ErrorFallback key="error"><Branch role="error" /></Control.ErrorFallback>];
  const container = document.createElement('div'); document.body.appendChild(container);
  let host = first, target = 'first';
  function View() {
    React.useLayoutEffect(() => { commits.push(container.textContent); });
    return <Provider instance={host}><Control {...props(target)}>{content()}</Control></Provider>;
  }
  const render = () => act(() => { ReactDOM.render(<View />, container); });
  const flush = async fn => { await act(async () => { fn?.(); for (let count = 0; count < 50; count++) await Promise.resolve(); }); };
  try {
    render(); await flush();
    assert.equal(container.textContent, 'loading');
    assert.equal(renders.granted, undefined); assert.equal(renders.denied, undefined); assert.equal(renders.error, undefined);
    assert.equal(container.childElementCount, 1); assert.equal(container.firstElementChild.tagName, 'SPAN');
    assert.ok(calls.some(call => call.id === 'first'));
    firstState.mode = 'granted'; await flush(() => firstState.obsolete.resolve());
    assert.equal(container.textContent, 'granted'); assert.equal(mounts.granted, 1); assert.equal(cleanups.loading, 1);
    // Target change masks the previous grant in the same committed render.
    firstState.mode = 'pending'; firstState.obsolete = deferred(); target = 'other';
    render(); await flush(); assert.equal(container.textContent, 'loading'); assert.ok(!commits.at(-1).includes('granted'));
    // Replacing the real provider's native service identities isolates requests.
    host = second; render(); await flush(); assert.equal(container.textContent, 'denied');
    assert.ok(calls.some(call => call.id === 'second'));
    await flush(() => firstState.obsolete.resolve());
    assert.equal(container.textContent, 'denied', 'obsolete target/provider success cannot grant');
    secondState.mode = 'error'; target = 'first'; render(); await flush();
    assert.equal(container.textContent, 'error'); assert.ok(mounts.error > 0); assert.ok(cleanups.denied > 0);
    secondState.mode = 'granted'; target = 'other'; render(); await flush(); assert.equal(container.textContent, 'granted');
    // Two real providers using one native transport must not share inflight work.
    const sharedState = { mode: 'granted', obsolete: deferred() };
    const sharedClients = client(sharedState, 'shared');
    const one = instance('isolated-one', sharedState, sharedClients), two = instance('isolated-two', sharedState, sharedClients);
    const before = calls.length;
    act(() => { ReactDOM.render(<><Provider instance={one}><Control {...props('first')}>{content('one')}</Control></Provider><Provider instance={two}><Control {...props('first')}>{content('two')}</Control></Provider></>, container); });
    await flush(); assert.equal(container.textContent, 'onetwo');
    const key = family === 'permission' ? '/effectivebasepermissions' : family === 'sharepoint-group' ? 'sitegroups(2)?' : '/me/checkMemberGroups';
    assert.equal(calls.slice(before).filter(call => call.url.includes(key)).length, 2, 'provider instances do not deduplicate one another');
    // Nested actual controls are logical AND; inner access never executes when outer denied.
    sharedState.mode = 'denied';
    act(() => { ReactDOM.render(<Provider instance={one}><Control {...props('other')}>
      <Control {...props('first')}><Branch role="and-granted" /><Control.Fallback><Branch role="and-denied" /></Control.Fallback></Control>
      <Control.Fallback><Branch role="outer-denied" /></Control.Fallback>
    </Control></Provider>, container); });
    await flush(); assert.equal(container.textContent, 'outer-denied'); assert.equal(renders['and-granted'], undefined); assert.equal(renders['and-denied'], undefined);
    sharedState.mode = 'granted';
    act(() => { ReactDOM.render(<Provider instance={one}><Control {...props('first')}>
      <Control {...props('other')}><Branch role="and-granted" /><Control.Fallback><Branch role="and-denied" /></Control.Fallback></Control>
    </Control></Provider>, container); });
    await flush(); assert.equal(container.textContent, 'and-granted');
    const mounted = Object.values(mounts).reduce((sum, count) => sum + count, 0);
    act(() => { ReactDOM.unmountComponentAtNode(container); });
    assert.equal(Object.values(cleanups).reduce((sum, count) => sum + count, 0), mounted);
    return { assertions: ['loading hides inactive consumers with no wrapper', 'positive, denied and native error select distinct mounted branches', 'target change masks prior grant in the committed render', 'provider replacement while pending uses replacement native service and ignores obsolete success', 'provider instances with identical native client remain isolated', 'nested controls implement AND without inactive inner mounting', 'unmount cleans up every branch'], calls, mounts, cleanups, commits, nativePermission };
  } finally { act(() => { ReactDOM.unmountComponentAtNode(container); }); container.remove(); }
}
