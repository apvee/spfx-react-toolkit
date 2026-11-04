const test = require('node:test');
const assert = require('node:assert/strict');
const { scenario } = require('./fixtures/executable-examples/cards-description-list.cjs');
const cases = [
    ['SitePicker', 'site-picker', 'Explore site selection and contextual search', 'read', 'Start the site picker, then compare its current-site and associated-hub choices with typed search while the selected value remains visible.'],
    ['ListPicker', 'list-picker', 'Explore list catalogs and independent consumer choices', 'read', 'Start the list picker, change the initial catalog limit, then copy confirmed models into consumer options and clear selection without losing those candidates.'],
    ['FieldPicker', 'field-picker', 'Explore site columns and list fields', 'read', 'Start the field picker for site columns or an existing list, then search its complete catalog and compare Dropdown with removable selected tags.'],
    ['ContentTypePicker', 'content-type-picker', 'Explore site and list content type references', 'read', 'Start a site or list content-type picker, apply group and parent-ID filters, and observe that source changes clear demo selection while presentation changes retain it.'],
    ['ViewPicker', 'view-picker', 'Explore scoped view catalogs', 'read', 'Enter an existing list and start the view picker, then filter the complete catalog while the selected view stays visible independently of the query.'],
    ['ListItemPicker', 'list-item-picker', 'Explore list item labels and query pages', 'read', 'Enter an existing list and display field, start the item picker, and compare query pages, file-name labels and retained selected references.'],
    ['PrincipalPicker', 'principal-picker', 'Explore principal discovery and explicit identity resolution', 'write', 'Start principal discovery, compare user and group sources and ID modes, and observe that only explicit confirmation can use the optional ensureUser policy.'],
];
for (const [name, file, title, kind, observe] of cases)
    test(`5a availability ${name}`, () => {
        const entry = scenario(title, `packages/spfx-react-toolkit/src/controls/pickers/${file}.tsx#${name}`).entry;
        assert.equal(entry.metadata.host, 'webPart');
        assert.equal(entry.metadata.kind, kind);
        assert.equal(entry.metadata.contentPlacement, 'configuration');
        assert.equal(entry.metadata.observe, observe);
        assert.equal(entry.metadata.hostAdapter, undefined);
    });
const { createReactPickersFixture } = require('./fixtures/executable-examples/react-pickers-scenarios.cjs');
test('5a checkpoint actual Principal source reset retires held native ensureUser success', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker', { revision: 9, hold: r => r.url.endsWith('/ensureuser') });
    assert.equal(f.calls.length, 0);
    assert.equal(f.initializations.length, 0);
    f.change('User IDs', 'sharePoint');
    f.toggle('Allow ensureUser on explicit confirmation');
    assert.equal(f.calls.length, 0);
    f.click('Start');
    await f.ready();
    f.open();
    f.search('Alice');
    await f.wait();
    f.choose('Alice');
    await f.flush();
    assert.equal(f.pending.length, 1, 'real native ensureUser dispatched');
    const oldStart = f.action('Start');
    const oldRefresh = f.action('Refresh / retry');
    const writes = f.calls.filter(r => r.url.endsWith('/ensureuser')).length;
    f.reset();
    assert.equal(f.trigger() === null, true);
    assert.equal(f.diagnostics.started, false);
    await f.flush(() => { oldStart(); oldRefresh(); });
    assert.equal(f.trigger() === null, true);
    f.click('Start');
    await f.ready();
    f.open();
    f.search('Bob');
    await f.wait();
    f.choose('Bob');
    await f.flush();
    assert.equal(f.diagnostics.selectedValues[0]?.loginName, 'bob@tenant.test');
    await f.settle(f.pending[0]);
    assert.equal(f.diagnostics.selectedValues[0]?.loginName, 'bob@tenant.test');
    assert.equal(f.diagnostics.error, '');
    assert.equal(f.calls.filter(r => r.url.endsWith('/ensureuser')).length, writes);
    f.origins();
});
for (const rejected of [false, true])
    test(`5a checkpoint old ensureUser ${rejected ? 'failure' : 'success'} cannot publish into unarmed reset`, async (t) => {
        const f = createReactPickersFixture(t, 'PrincipalPicker', { hold: r => r.url.endsWith('/ensureuser') });
        f.change('User IDs', 'sharePoint');
        f.toggle('Allow ensureUser on explicit confirmation');
        f.click('Start');
        await f.ready();
        f.open();
        f.search('Alice');
        await f.wait();
        f.choose('Alice');
        await f.flush();
        assert.equal(f.pending.length, 1);
        f.reset();
        await f.settle(f.pending[0], rejected ? new Error('old ensure failure') : undefined);
        assert.deepEqual(f.diagnostics.selectedOptions, []);
        assert.equal(f.diagnostics.error, '');
        assert.equal(f.diagnostics.started, false);
        assert.equal(f.trigger() === null, true);
    });
test('5a checkpoint native latest-callback negative control proves same [] plus refresh is insufficient', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker', { hold: r => r.url.endsWith('/ensureuser'), mutate: source => source
            .replace('started && session.resetRevision === resetRevision &&', 'started &&')
            .replace('boundary.current = { resetRevision, owner: spfxContext, scope };', 'boundary.current.resetRevision = resetRevision;')
            .replace("setStarted(false); setConfig(baseline); setSelection(undefined); setAvailable(undefined); setError(''); setRefresh(0);", "setSelection(undefined); setError(''); setRefresh(value => value + 1);") });
    f.change('User IDs', 'sharePoint');
    f.toggle('Allow ensureUser on explicit confirmation');
    f.click('Start');
    await f.ready();
    f.open();
    f.search('Alice');
    await f.wait();
    f.choose('Alice');
    await f.flush();
    assert.equal(f.pending.length, 1);
    const input = f.trigger();
    f.reset();
    assert.equal(f.trigger() === input, true, 'same operational Subject retained by counterfactual');
    assert.deepEqual(f.diagnostics.selectedOptions, []);
    await f.settle(f.pending[0]);
    assert.equal(f.diagnostics.selectedValues[0]?.loginName, 'alice@tenant.test', 'real stable emit invokes latest callback despite same empty controlled keys and refresh');
});
const { lists, fields, views, contentTypes, listId, acceptsSiteSearchPost } = require('./fixtures/executable-examples/react-pickers-native.cjs');
function prepare(f, name) { if (['ViewPicker', 'ListItemPicker'].includes(name))
    f.change('Exact list title', name === 'ListItemPicker' ? 'Docs' : 'Tasks'); }
function options() { return [...document.querySelectorAll('[role="option"],[role="menuitemcheckbox"]')]; }
for (const [name] of cases)
    test('5a unarmed defaults native lifetime and retained commands ' + name, async (t) => {
        const f = createReactPickersFixture(t, name, { revision: 17 });
        const header = f.container.querySelector('[aria-label="Suggestion header"]');
        const retained = ['Start', 'Use confirmed models as consumer options', 'Empty consumer options', 'Clear controlled selection', 'Refresh / retry'].map(label => f.action(label));
        assert.equal(f.subject.mounts, 0);
        assert.equal(f.calls.length + f.initializations.length, 0);
        f.change('Suggestion header', 'unarmed edit');
        f.toggle('Custom search and status text');
        f.toggle('Customize status content');
        prepare(f, name);
        assert.equal(f.subject.mounts, 0);
        assert.equal(f.calls.length + f.initializations.length, 0);
        f.reset();
        assert.equal(f.container.querySelector('[aria-label="Suggestion header"]') === header, true, 'outer Demo native config ref survives Reset');
        await f.flush(() => retained.forEach(action => action()));
        assert.equal(f.subject.mounts, 0);
        assert.equal(f.calls.length + f.initializations.length, 0);
        prepare(f, name);
        f.click('Start');
        await f.ready();
        assert.equal(f.subject.mounts, 1);
        const p = f.subject.props;
        assert.equal(p.presentation, 'combobox');
        assert.equal(p.multiselect, false);
        assert.equal(p.searchable, true);
        assert.equal(p.colorfulTagAvatars, false);
        assert.deepEqual(p.selectedOptions, []);
        if (name === 'ListPicker')
            assert.equal(p.initialCatalogLimit, 30);
        if (name === 'SitePicker') {
            assert.equal(p.includeCurrentSite, true);
            assert.equal(p.includeHubSite, true);
            assert.equal(Object.hasOwn(p, 'webUrl'), false);
        }
        if (name === 'FieldPicker')
            assert.equal(p.list, undefined);
        if (name === 'ContentTypePicker') {
            assert.equal(p.scope, 'site');
            assert.equal(p.list, undefined);
            assert.equal(p.includeHidden, true);
            assert.equal(p.includeDescendants, false);
        }
        if (name === 'ListItemPicker')
            assert.equal(p.textField, 'Title');
        if (name === 'PrincipalPicker') {
            assert.deepEqual(p.principalTypes, ['user']);
            assert.deepEqual(p.userIdentity, { mode: 'available' });
            assert.equal(p.loadPersonalSuggestions, false);
            assert.equal(p.initialSuggestions.includeCurrentUser, false);
        }
        const trigger = f.trigger(), calls = f.calls.length;
        assert.ok(trigger);
        assert.ok(f.container.querySelector(`label[for="${trigger.id}"]`), 'native accessible label retained');
        f.change('Suggestion header', 'Retained header');
        f.toggle('Custom search and status text');
        f.click('Start');
        f.click('Use confirmed models as consumer options');
        f.click('Use confirmed models as consumer options');
        await f.flush();
        assert.equal(f.trigger() === trigger, true);
        assert.equal(f.subject.mounts, 1);
        assert.equal(f.calls.length, calls);
        f.verifyUniverse();
        f.origins();
        f.reset();
        assert.equal(f.subject.unmounts, 1);
        assert.equal(f.trigger() === null, true);
        assert.equal(f.diagnostics.started, false);
    });
for (const [name, expected, query] of [['ListPicker', 30, 'List 44'], ['FieldPicker', 210, 'column 209'], ['ViewPicker', 210, 'view 209'], ['ContentTypePicker', 210, 'content 209']])
    test('5a actual paged ' + name + ' catalog and local/remote matching', async (t) => {
        const f = createReactPickersFixture(t, name);
        prepare(f, name);
        f.click('Start');
        await f.ready();
        f.open();
        await f.flush();
        assert.equal(options().length, expected);
        const before = f.calls.length;
        f.search(query);
        await f.wait();
        assert.equal(options().length, 1);
        assert.match(options()[0].textContent, new RegExp(query, 'i'));
        assert.equal(f.calls.length > before, name === 'ListPicker', 'only List typed discovery is remote');
        const rows = name === 'ListPicker' ? lists : name === 'FieldPicker' ? fields : name === 'ViewPicker' ? views : contentTypes;
        if (name !== 'ListPicker')
            assert.equal(f.calls.filter(r => new URL(r.url).searchParams.get('$skiptoken') === '120').length, 1, 'complete metadata continuation read');
        f.choose(query.slice(0, 1).toUpperCase() + query.slice(1));
        await f.flush();
        assert.equal(f.diagnostics.selectedValues.length, 1);
        assert.equal(f.diagnostics.selectedOptions.length, 1);
        assert.equal(f.diagnostics.reason, 'selection');
        assert.equal(f.diagnostics.roundTrip, true);
        const selected = f.diagnostics.selectedOptions[0], model = f.diagnostics.selectedItems[0];
        assert.equal(model.text.toLowerCase(), query.toLowerCase());
        assert.ok(rows.length > 30);
        assert.equal(f.trigger().textContent.includes(model.text), true, 'selected label remains independent of query');
        f.click('Use confirmed models as consumer options');
        f.click('Clear controlled selection');
        f.open();
        await f.flush();
        assert.equal(f.diagnostics.selectedOptions.length, 0);
        assert.equal(f.diagnostics.consumerOptions[0].key, selected);
        assert.ok(options().some(row => row.textContent.includes(model.text)), 'copied candidates survive controlled clear');
        f.choose(model.text);
        await f.flush();
        f.click('Empty consumer options');
        await f.flush();
        assert.equal(f.diagnostics.selectedOptions[0], selected, 'empty options never clears selection');
    });
test('5a Site actual PnP200 serialization native paging and pinned contextual choices', async (t) => {
    const f = createReactPickersFixture(t, 'SitePicker');
    f.click('Start');
    await f.ready();
    f.open();
    await f.flush();
    assert.ok(options().some(row => row.textContent.includes('Current collection')));
    assert.ok(options().some(row => row.textContent.includes('Associated hub')));
    f.search('Site');
    await f.wait(380);
    await f.wait(80);
    const native = f.calls.filter(r => r.kind === 'pnp' && r.url.endsWith('/_api/search/postquery'));
    assert.deepEqual(native.map(r => JSON.parse(r.init.body).request.StartRow ?? 0), [0, 200, 400]);
    assert.ok(native.every(r => acceptsSiteSearchPost(new URL(r.url), r.init)));
    assert.ok(native.every(r => JSON.parse(r.init.body).request.SortList.length === 0));
    assert.ok(options().find(row => row.textContent.includes('Site 000')), 'bounded native pages sorted before UI slicing');
    assert.ok(options().find(row => row.textContent.includes('Current collection')));
    assert.ok(options().find(row => row.textContent.includes('Associated hub')));
    f.choose('Site 000');
    await f.wait(100);
    assert.equal(f.diagnostics.selectedValues[0]?.url, 'https://tenant.test/sites/site-0');
    f.open();
    f.search('no-match');
    await f.wait();
    assert.ok(options().some(row => row.textContent.includes('Current collection')));
    assert.ok(options().some(row => row.textContent.includes('Associated hub')));
    assert.equal(f.diagnostics.selectedValues[0]?.url, 'https://tenant.test/sites/site-0');
    const observed = f.origins();
    for (const stem of ['services/spfx-pnp-search.service', 'services/spfx-pnp-context.service'])
        assert.equal(observed.includes(require('node:path').join(f.root, stem + (f.compiled ? '.js' : '.ts'))), true, 'actual lazy Site search service/factory origin');
});
test('5a Item actual field metadata filename fallback query and confirmation projection', async (t) => {
    const f = createReactPickersFixture(t, 'ListItemPicker');
    assert.equal(f.subject.mounts, 0);
    f.click('Start');
    assert.equal(f.subject.mounts, 0);
    f.change('Exact list title', 'Docs');
    f.click('Start');
    await f.ready();
    f.open();
    f.search('quarterly');
    await f.wait();
    assert.ok(options().some(row => row.textContent.includes('Quarterly-Report.docx')));
    assert.equal(options().some(row => row.textContent === 'Folder'), false);
    f.choose('Quarterly-Report.docx');
    await f.flush();
    assert.equal(f.diagnostics.selectedItems[0]?.text, 'Quarterly-Report.docx');
    assert.equal(f.diagnostics.selectedValues[0]?.kind, 'file');
    assert.equal(f.diagnostics.selectedValues[0]?.itemId, 1);
    assert.ok(f.calls.some(r => r.url.includes('getByInternalNameOrTitle') && r.url.includes('Title')));
    assert.ok(f.calls.some(r => new URL(r.url).searchParams.get('$filter') === '(FSObjType eq 0) and (Id eq 1)'));
    f.change('Item display field', 'FileLeafRef');
    await f.flush();
    assert.equal(f.diagnostics.selectedOptions.length, 1, 'source policy edit retains authoritative keys');
});
test('5a Principal Graph failure leaves real SharePoint search and confirmation usable', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker');
    f.click('Start');
    await f.ready(new Error('Graph unavailable'));
    f.open();
    f.search('Bob');
    await f.wait();
    f.choose('Bob');
    await f.flush();
    assert.equal(f.diagnostics.selectedValues[0]?.loginName, 'bob@tenant.test');
    assert.equal(f.calls.some(r => r.kind === 'graph'), false);
    assert.equal(f.calls.some(r => r.url.endsWith('/ensureuser')), false);
});
test('5a Principal source confirms and sends opt-in ensureUser before custom validation', async (t) => {
    let validations = 0;
    const validate = () => { validations++; return 'Refused after identity confirmation'; };
    const f = createReactPickersFixture(t, 'PrincipalPicker', { hold: r => r.url.endsWith('/ensureuser'), subjectProps: props => ({ ...props, validateSelection: validate }) });
    f.change('User IDs', 'sharePoint');
    f.toggle('Allow ensureUser on explicit confirmation');
    f.click('Start');
    await f.ready();
    f.open();
    f.search('Alice');
    await f.wait();
    f.choose('Alice');
    await f.flush();
    assert.equal(f.pending.length, 1);
    assert.equal(validations, 0, 'validator cannot prevent a write already dispatched by source.confirm');
    f.choose('Alice');
    await f.flush();
    assert.equal(f.pending.length, 1, 'live runtime excludes duplicate confirmation while write is pending');
    await f.settle(f.pending[0]);
    assert.equal(validations, 1);
    assert.deepEqual(f.diagnostics.selectedOptions, []);
    assert.match(f.container.textContent, /Refused after identity confirmation/);
});
test('5a Principal group categories identity modes and no provisioning from configuration', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker');
    f.change('Principal types', 'groups');
    f.change('User IDs', 'both');
    f.toggle('Allow ensureUser on explicit confirmation');
    f.click('Start');
    await f.ready();
    assert.deepEqual(f.subject.props.principalTypes, ['sharePointGroup', 'microsoft365Group', 'securityGroup', 'distributionList']);
    assert.deepEqual(f.subject.props.userIdentity, { mode: 'both', webUrl: undefined, ensureUser: true });
    f.open();
    f.search('Group');
    await f.wait();
    assert.equal(options().length, 4);
    assert.equal(f.calls.some(r => r.url.endsWith('/ensureuser')), false);
    f.choose('Group SharePoint');
    await f.flush();
    assert.equal(f.diagnostics.selectedValues[0]?.kind, 'sharePointGroup');
    assert.equal(f.diagnostics.selectedValues[0]?.identity.id, 51);
    f.change('Principal types', 'all');
    assert.equal(f.subject.props.principalTypes.length, 5);
    f.change('User IDs', 'available');
    assert.deepEqual(f.subject.props.userIdentity, { mode: 'available' });
    assert.equal(f.calls.some(r => r.url.endsWith('/ensureuser')), false);
});
test('5a two actual providers isolate arming reset selection and late completion', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker', { hold: r => r.url.endsWith('/ensureuser') });
    const peer = f.mountPeer();
    f.change('User IDs', 'sharePoint');
    f.toggle('Allow ensureUser on explicit confirmation');
    f.click('Start');
    await f.ready();
    assert.equal(peer.diagnostics.started, false);
    assert.equal(f.initializations.length, 1);
    f.open();
    f.search('Alice');
    await f.wait();
    f.choose('Alice');
    await f.flush();
    assert.equal(f.pending.length, 1);
    assert.equal(peer.diagnostics.selectedOptions.length, 0);
    f.reset();
    peer.click('Start');
    await f.ready();
    await f.settle(f.pending[0]);
    assert.equal(peer.diagnostics.started, true);
    assert.deepEqual(peer.diagnostics.selectedOptions, []);
    assert.equal(peer.diagnostics.error, '');
    assert.equal(f.diagnostics.started, false);
    assert.equal(f.calls.filter(r => r.id === 'pickers-peer').length, 0);
});
test('5a real owner/client replacement retires pending action and true ABA cannot revive commands', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker', { hold: r => r.url.endsWith('/ensureuser') });
    const first = f.instance;
    f.change('User IDs', 'sharePoint');
    f.toggle('Allow ensureUser on explicit confirmation');
    f.click('Start');
    await f.ready();
    f.open();
    f.search('Alice');
    await f.wait();
    f.choose('Alice');
    await f.flush();
    const stale = f.action('Start');
    f.replace(f.host('replacement'));
    f.replace(first);
    await f.flush(() => stale());
    assert.equal(f.diagnostics.started, false);
    await f.settle(f.pending[0]);
    assert.deepEqual(f.diagnostics.selectedValues, []);
    assert.equal(f.diagnostics.error, '');
});
for (const mode of ['sharePoint', 'both'])
    test('5a Principal ' + mode + ' false opt-in fails read-only missing registration', async (t) => {
        const f = createReactPickersFixture(t, 'PrincipalPicker');
        f.change('User IDs', mode);
        f.click('Start');
        await f.ready();
        f.open();
        f.search('Alice');
        await f.wait();
        f.choose('Alice');
        await f.flush();
        assert.equal(f.calls.some(r => r.url.endsWith('/ensureuser')), false);
        assert.deepEqual(f.diagnostics.selectedValues, []);
        assert.match(f.diagnostics.error, /404/);
    });
test('5a ContentType retains validated ID choices and clears descendant policy on empty ID', async (t) => {
    const f = createReactPickersFixture(t, 'ContentTypePicker');
    f.click('Start');
    await f.ready();
    const event = (value) => { const field = f.regions.controls.props.children.find(node => node?.props?.label === 'Content type ID filter'); f.h.act(() => field.props.children.props.onChange(undefined, { value })); };
    f.change('Content type ID filter', '0x01');
    f.toggle('Include descendants of content type ID');
    assert.equal(f.subject.props.contentTypeId, '0x01');
    assert.equal(f.subject.props.includeDescendants, true);
    f.change('Content type ID filter', '');
    assert.equal(f.subject.props.contentTypeId, undefined);
    assert.equal(f.subject.props.includeDescendants, false);
    const toggle = [...f.container.querySelectorAll('input')].find(node => [...f.container.querySelectorAll('label')].some(label => label.htmlFor === node.id && label.textContent === 'Include descendants of content type ID'));
    assert.equal(toggle.disabled, true);
    f.change('Content type ID filter', '0x01');
    for (const invalid of ['0', '0x0', '0x00', '0x0101?x=1']) {
        event(invalid);
        assert.equal(f.subject.props.contentTypeId, '0x01');
    }
    f.change('Content type ID filter', '');
    f.open();
    await f.flush();
    f.choose('Content 002');
    await f.flush();
    const id = f.diagnostics.selectedValues[0].contentTypeId;
    assert.ok([...f.container.querySelector('[aria-label="Content type ID filter"]').options].some(option => option.value === id), 'confirmed metadata IDs become offered filter choices');
    f.change('Content type ID filter', id);
    assert.equal(f.subject.props.contentTypeId, id);
});
test('5a ContentType effective source changes clear selection/options while cosmetic and equivalent edits retain both', async (t) => {
    const f = createReactPickersFixture(t, 'ContentTypePicker');
    f.click('Start');
    await f.ready();
    f.open();
    await f.flush();
    f.choose('Content 002');
    await f.flush();
    f.click('Use confirmed models as consumer options');
    const key = f.diagnostics.selectedOptions[0], trigger = f.trigger(), count = f.calls.length;
    f.change('Web URL', ' HTTPS://TENANT.TEST/sites/root/sub/// ');
    f.change('Suggestion header', 'Cosmetic');
    f.change('Exact list title', 'Tasks');
    await f.flush();
    assert.equal(f.diagnostics.selectedOptions[0], key);
    assert.equal(f.diagnostics.consumerOptions[0].key, key);
    assert.equal(f.trigger() === trigger, true);
    assert.equal(f.calls.length, count);
    f.change('Content type scope', 'list');
    await f.flush();
    assert.deepEqual(f.diagnostics.selectedOptions, []);
    assert.equal(f.diagnostics.consumerOptions, undefined);
    assert.equal(f.subject.props.scope, 'list');
    assert.equal(f.subject.props.list, 'Tasks');
    f.open();
    await f.flush();
    f.choose('Content 002');
    await f.flush();
    assert.equal(f.diagnostics.selectedValues[0]?.scope, 'list');
    assert.equal(f.diagnostics.selectedValues[0]?.listId, listId);
    f.change('Exact list title', 'Docs');
    assert.deepEqual(f.diagnostics.selectedOptions, []);
});
for (const continuation of ['foreign', 'cycle', 'changed-query'])
    test('5a real complete catalog rejects ' + continuation + ' continuation', async (t) => {
        const f = createReactPickersFixture(t, 'FieldPicker', { route: r => {
                const url = new URL(r.url);
                if (url.pathname.endsWith('/fields'))
                    return new Response(JSON.stringify({ value: [fields[0]], 'odata.nextLink': continuation === 'foreign' ? 'https://other.invalid/_api/web/fields' : continuation === 'changed-query' ? url.origin + url.pathname + '?$select=Wrong' : r.url }), { headers: { 'Content-Type': 'application/json' } });
            } });
        f.click('Start');
        await f.ready();
        f.open();
        await f.flush();
        assert.match(f.diagnostics.error, /continuation/i);
        assert.equal(f.calls.length <= 2, true, 'bad continuation is bounded');
        assert.equal(f.calls.every(request => new URL(request.url).origin === 'https://tenant.test'), true, 'foreign continuation never dispatches');
    });
test('5a native dictionary status slots classes refs and input survive cosmetic rerenders', async (t) => {
    let headerRef, clicks = 0;
    const f = createReactPickersFixture(t, 'FieldPicker', { hold: r => r.url.includes('/fields?'), subjectProps: props => ({ ...props, suggestionsHeader: { children: props.suggestionsHeader, className: 'picker-header-proof', ref: node => { headerRef = node; }, onClick: () => { clicks++; } } }) });
    f.toggle('Custom search and status text');
    f.toggle('Customize status content');
    f.change('Suggestion header', 'Custom header');
    f.click('Start');
    await f.ready();
    assert.equal(f.trigger().textContent, 'Choose source content');
    f.open();
    await f.flush();
    assert.match(document.body.textContent, /Fetching live suggestions/);
    assert.equal(headerRef?.tagName, 'SPAN');
    assert.equal(headerRef.classList.contains('picker-header-proof'), true);
    f.h.act(() => { headerRef.click(); });
    assert.equal(clicks, 1);
    const trigger = f.trigger();
    f.hold(undefined);
    await f.settle(f.pending[0]);
    f.change('Suggestion header', 'Updated header');
    await f.flush();
    assert.equal(f.trigger() === trigger, true);
    assert.equal(headerRef?.textContent, 'Updated header');
    f.search('zzz');
    await f.wait();
    assert.match(document.body.textContent, /No candidates: type at least two characters/);
});
test('5a actual source fault shows custom error shorthand and refresh retries only while started', async (t) => {
    let failed = true;
    const f = createReactPickersFixture(t, 'FieldPicker', { route: r => failed && new URL(r.url).pathname.endsWith('/fields') ? new Response('{}', { status: 503, headers: { 'Content-Type': 'application/json' } }) : undefined });
    f.toggle('Customize status content');
    f.click('Refresh / retry');
    assert.equal(f.calls.length, 0);
    f.click('Start');
    await f.ready();
    f.open();
    await f.flush();
    assert.match(f.diagnostics.error, /503/);
    assert.match(document.body.textContent, /Source unavailable: inspect diagnostics and retry/);
    failed = false;
    f.click('Refresh / retry');
    f.open();
    await f.flush();
    assert.equal(options().length, 210);
    assert.equal(f.diagnostics.error, '');
});
test('5a witness detects an independent provider/context cache and restores matching identity', t => {
    const f = createReactPickersFixture(t, 'PrincipalPicker');
    const { createPickerHarness } = require('./fixtures/picker-source-harness.cjs');
    const foreign = createPickerHarness();
    const actual = f.toolkit.useSPFxContext;
    const separate = foreign.load(require('node:path').join(f.root, 'core/context.internal' + (f.compiled ? '.js' : '.tsx'))).useSPFxContext;
    assert.notEqual(actual, separate, 'same filename loaded through a distinct cache has a different context universe');
    f.toolkit.useSPFxContext = separate;
    assert.throws(() => f.verifyUniverse(), assert.AssertionError);
    f.toolkit.useSPFxContext = actual;
    f.verifyUniverse();
});
test('5a instrumented valid initial reference remains silent before Start and hydrates without a selection event', async (t) => {
    const value = { kind: 'user', loginName: 'bob@tenant.test', sharePoint: { siteCollectionUrl: 'https://tenant.test/sites/root', id: 43 } };
    const alias = 'picker:v1:' + encodeURIComponent(JSON.stringify(['user', value]));
    const f = createReactPickersFixture(t, 'PrincipalPicker', { revision: 8, mutate: source => source.replaceAll('selection?.selectedOptions ?? []', `selection?.selectedOptions ?? [${JSON.stringify(alias)}]`) });
    assert.equal(f.calls.length + f.initializations.length, 0);
    f.change('Suggestion header', 'Edit seeded wrapper');
    f.reset();
    assert.equal(f.calls.length + f.initializations.length, 0);
    f.click('Start');
    await f.ready();
    await f.flush();
    assert.ok(f.calls.some(r => r.url.includes('/siteusers/getById(43)')));
    assert.match(f.trigger().textContent, /Bob/);
    assert.equal(f.diagnostics.reason, 'No action');
    assert.deepEqual(f.diagnostics.selectedOptions, [alias]);
});
test('5a opaque unresolved alias remains authoritative until explicit native removal', async (t) => {
    const f = createReactPickersFixture(t, 'FieldPicker', { mutate: source => source.replaceAll('selection?.selectedOptions ?? []', "selection?.selectedOptions ?? ['opaque-alias']") });
    f.click('Start');
    await f.ready();
    assert.equal(f.trigger().textContent, 'Unresolved value');
    assert.deepEqual(f.diagnostics.selectedOptions, ['opaque-alias']);
    assert.deepEqual(f.diagnostics.selectedValues, []);
    assert.equal(f.diagnostics.reason, 'No action');
    const clear = f.container.querySelector('.fui-Dropdown__clearButton');
    assert.ok(clear);
    f.h.act(() => { clear.click(); });
    await f.flush();
    assert.deepEqual(f.diagnostics.selectedOptions, []);
    assert.equal(f.diagnostics.reason, 'clear');
});
test('5a Start-gate fault control exposes real seeded hydration on import', async (t) => {
    const value = { kind: 'user', loginName: 'bob@tenant.test', sharePoint: { siteCollectionUrl: 'https://tenant.test/sites/root', id: 43 } }, alias = 'picker:v1:' + encodeURIComponent(JSON.stringify(['user', value]));
    const f = createReactPickersFixture(t, 'PrincipalPicker', { mutate: source => source.replace('const active = started &&', 'const active = true &&').replaceAll('selection?.selectedOptions ?? []', `selection?.selectedOptions ?? [${JSON.stringify(alias)}]`) });
    await f.ready();
    assert.ok(f.calls.some(r => r.url.includes('/siteusers/getById(43)')), 'removing actual mount gate performs a real verification read');
    assert.equal(f.subject.mounts, 1);
});
for (const [label, kind] of [['Group Microsoft365', 'microsoft365Group'], ['Group Security', 'securityGroup'], ['Group Distribution', 'distributionList']])
    test('5a confirms scoped directory ' + kind + ' independently of user identity provisioning', async (t) => {
        const f = createReactPickersFixture(t, 'PrincipalPicker');
        f.change('Principal types', 'groups');
        f.change('User IDs', 'both');
        f.toggle('Allow ensureUser on explicit confirmation');
        f.click('Start');
        await f.ready();
        f.open();
        f.search('Group');
        await f.wait();
        f.choose(label);
        await f.flush();
        assert.equal(f.diagnostics.selectedValues[0]?.kind, kind);
        assert.equal(f.diagnostics.selectedValues[0]?.identity.tenantId, '11111111-1111-4111-8111-111111111111');
        assert.equal(f.calls.some(r => r.url.endsWith('/ensureuser')), false);
    });
test('5a personal and current-user opt-ins make only real read requests after Start/open', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker');
    f.toggle('Load Graph personal suggestions');
    f.toggle('Offer current user');
    f.change('Principal types', 'all');
    assert.equal(f.calls.length, 0);
    f.click('Start');
    await f.ready();
    assert.equal(f.calls.length, 0);
    f.open();
    await f.flush();
    assert.ok(f.calls.some(r => r.uri === '/me/people?%24top=50'));
    assert.ok(f.calls.some(r => r.uri === '/me/memberOf?%24top=50'));
    assert.ok(f.calls.some(r => r.url.includes('/_api/web/currentuser?')));
    assert.equal(f.calls.some(r => r.url.endsWith('/ensureuser')), false);
    assert.ok(options().some(row => row.textContent.includes('Current user')));
});
test('5a separate Site protocol accepts only its exact bounded native shape and preserves Search rejection', t => {
    const f = createReactPickersFixture(t, 'PrincipalPicker');
    const url = new URL('https://tenant.test/sites/root/sub/_api/search/postquery');
    const request = { Querytext: 'contentclass:STS_Site AND (Title:"Site*" OR Path:"Site*")', RowLimit: 200, HitHighlightedProperties: [], Properties: [], RefinementFilters: [], ReorderingRules: [], SelectProperties: ['Title', 'Path', 'SiteId', 'DepartmentId'], SortList: [], TrimDuplicates: true };
    const init = q => ({ method: 'POST', headers: { 'Content-Type': 'application/json', 'X-RequestDigest': 'native' }, body: JSON.stringify({ request: q }) });
    assert.equal(acceptsSiteSearchPost(url, init(request)), true);
    const existing = f.h.load(require('node:path').resolve('tests/fixtures/demo-presentation/pnp-search-protocol.ts'));
    assert.equal(existing.acceptsSearchPost(url, init(request), 'https://tenant.test'), false, 'existing strict Search mode remains unchanged');
    for (const patch of [{ RowLimit: 5 }, { StartRow: 50 }, { StartRow: 1000 }, { TrimDuplicates: 'true' }, { SelectProperties: ['Title', 'Path'] }, { SortList: [{ Property: 'Title', Direction: 0 }] }, { RefinementFilters: ['FileType:pdf'] }, { Querytext: '*' }, { Unknown: true }])
        assert.equal(acceptsSiteSearchPost(url, init({ ...request, ...patch })), false);
    assert.equal(acceptsSiteSearchPost(new URL('https://foreign.invalid/_api/search/postquery'), init(request)), false);
    assert.equal(acceptsSiteSearchPost(url, { ...init(request), method: 'GET' }), false);
});
test('5a native multiple selection removal TagPicker projection and cosmetic cache retention', async (t) => {
    const f = createReactPickersFixture(t, 'FieldPicker');
    f.toggle('Multiple selection');
    f.click('Start');
    await f.ready();
    f.open();
    await f.flush();
    f.choose('Column 000');
    await f.flush();
    f.choose('Column 001');
    await f.flush();
    assert.equal(f.diagnostics.selectedOptions.length, 2);
    f.choose('Column 000');
    await f.flush();
    assert.equal(f.diagnostics.selectedOptions.length, 1);
    assert.equal(f.diagnostics.reason, 'remove');
    const key = f.diagnostics.selectedOptions[0], count = f.calls.length;
    f.change('Presentation', 'tagPicker');
    await f.flush();
    assert.equal(f.subject.mounts, 1);
    assert.equal(f.diagnostics.selectedOptions[0], key);
    assert.equal(f.container.querySelectorAll('.fui-Tag').length, 1);
    assert.equal(f.calls.length, count);
    const input = f.trigger();
    f.h.act(() => { input.focus(); require('react-dom/test-utils').Simulate.keyDown(input, { key: 'ArrowDown' }); });
    await f.flush();
    assert.equal([...document.querySelectorAll('.fui-TagPickerOption')].some(row => row.textContent.includes('Column 001')), false, 'TagPicker excludes selected popup row');
    f.change('Suggestion header', 'Equivalent presentation');
    await f.flush();
    assert.equal(f.trigger() === input, true);
    assert.equal(f.calls.length, count);
    f.h.act(() => { require('react-dom/test-utils').Simulate.keyDown(input, { key: 'Escape' }); });
    assert.equal(document.activeElement === input, true);
});
for (const failure of [false, true])
    test('5a held real candidate verification ' + (failure ? 'failure' : 'success') + ' after Reset starts no follow-up identity write', async (t) => {
        const f = createReactPickersFixture(t, 'PrincipalPicker', { hold: r => r.method === 'POST' && r.init?.body && JSON.parse(r.init.body).queryParams?.QueryString === 'alice@tenant.test' });
        f.change('User IDs', 'sharePoint');
        f.toggle('Allow ensureUser on explicit confirmation');
        f.click('Start');
        await f.ready();
        f.open();
        f.search('Alice');
        await f.wait();
        f.choose('Alice');
        await f.flush();
        assert.equal(f.pending.length, 1);
        assert.equal(f.calls.some(r => r.url.endsWith('/ensureuser')), false);
        f.reset();
        await f.settle(f.pending[0], failure ? new Error('retired verification') : undefined);
        assert.equal(f.calls.some(r => r.url.endsWith('/ensureuser')), false);
        assert.equal(f.diagnostics.started, false);
        assert.equal(f.diagnostics.error, '');
        assert.deepEqual(f.diagnostics.selectedValues, []);
    });
for (const failure of [false, true])
    test('5a held real typed search ' + (failure ? 'failure' : 'success') + ' cannot replace new Subject selection', async (t) => {
        const f = createReactPickersFixture(t, 'PrincipalPicker', { hold: r => r.method === 'POST' && r.init?.body && JSON.parse(r.init.body).queryParams?.QueryString === 'Al' });
        f.click('Start');
        await f.ready();
        f.open();
        f.search('Al');
        await f.wait();
        assert.equal(f.pending.length, 1);
        f.reset();
        f.click('Start');
        await f.ready();
        f.open();
        f.search('Bob');
        await f.wait();
        f.choose('Bob');
        await f.flush();
        const current = f.diagnostics.selectedOptions[0];
        await f.settle(f.pending[0], failure ? new Error('retired search') : undefined);
        assert.equal(f.diagnostics.selectedOptions[0], current);
        assert.equal(f.diagnostics.selectedValues[0]?.loginName, 'bob@tenant.test');
        assert.equal(f.diagnostics.error, '');
    });
for (const failure of [false, true])
    test('5a held real hydration ' + (failure ? 'failure' : 'success') + ' settles silently after Subject retirement', async (t) => {
        const value = { kind: 'user', loginName: 'bob@tenant.test', sharePoint: { siteCollectionUrl: 'https://tenant.test/sites/root', id: 43 } }, alias = 'picker:v1:' + encodeURIComponent(JSON.stringify(['user', value]));
        const f = createReactPickersFixture(t, 'PrincipalPicker', { hold: r => r.url.includes('/siteusers/getById(43)'), mutate: source => source.replaceAll('selection?.selectedOptions ?? []', `selection?.selectedOptions ?? [${JSON.stringify(alias)}]`) });
        f.click('Start');
        await f.ready();
        assert.equal(f.pending.length, 1);
        f.reset();
        await f.settle(f.pending[0], failure ? new Error('retired hydration') : undefined);
        assert.equal(f.subject.unmounts, 1);
        assert.equal(f.diagnostics.started, false);
        assert.equal(f.diagnostics.reason, 'No action');
        assert.equal(f.diagnostics.error, '');
    });
test('5a same-value identity and cosmetic edits retain held live confirmation and native references', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker', { hold: r => r.url.endsWith('/ensureuser') });
    f.change('User IDs', 'sharePoint');
    f.toggle('Allow ensureUser on explicit confirmation');
    f.click('Start');
    await f.ready();
    f.open();
    f.search('Alice');
    await f.wait();
    f.choose('Alice');
    await f.flush();
    const trigger = f.trigger(), count = f.calls.length;
    f.change('User IDs', 'sharePoint');
    f.change('Suggestion header', 'Cosmetic during confirmation');
    f.toggle('Custom search and status text');
    await f.flush();
    assert.equal(f.trigger() === trigger, true);
    assert.equal(f.subject.mounts, 1);
    assert.equal(f.calls.length, count);
    await f.settle(f.pending[0]);
    assert.equal(f.diagnostics.selectedValues[0]?.loginName, 'alice@tenant.test');
    assert.equal(f.diagnostics.selectedValues[0]?.sharePoint.id, 42);
});
test('5a true source identity ABA retains Subject but retires old confirmation', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker', { hold: r => r.url.endsWith('/ensureuser') });
    f.change('User IDs', 'sharePoint');
    f.toggle('Allow ensureUser on explicit confirmation');
    f.click('Start');
    await f.ready();
    f.open();
    f.search('Alice');
    await f.wait();
    f.choose('Alice');
    await f.flush();
    f.change('User IDs', 'available');
    f.change('User IDs', 'sharePoint');
    assert.equal(f.subject.mounts, 1);
    await f.settle(f.pending[0]);
    assert.deepEqual(f.diagnostics.selectedValues, []);
    assert.equal(f.diagnostics.error, '');
});
test('5a List offers every real catalog cap with stable native page cache and independent typed discovery', async (t) => {
    const f = createReactPickersFixture(t, 'ListPicker');
    const choices = [...f.container.querySelector('[aria-label="Initial catalog limit"]').options].map(node => node.value);
    assert.deepEqual(choices, ['10', '20', '30', '50', '100', '200']);
    f.click('Start');
    await f.ready();
    assert.equal(f.calls.length, 0);
    f.open();
    await f.flush();
    const reads = f.calls.length;
    for (const limit of [10, 20, 30, 50, 100, 200]) {
        f.change('Initial catalog limit', String(limit));
        await f.flush();
        assert.equal(options().length, Math.min(limit, 45));
        assert.equal(f.calls.length, reads);
    }
    f.search('list 44');
    await f.wait();
    assert.equal(options().length, 1);
    assert.ok(f.calls.length > reads);
});
test('5a Field exact list scope and View/Item prerequisites never fabricate a target', async (t) => {
    const f = createReactPickersFixture(t, 'FieldPicker');
    f.change('Exact list title', '  Tasks  ');
    f.click('Start');
    await f.ready();
    f.open();
    await f.flush();
    assert.ok(f.calls.some(r => decodeURIComponent(r.url).includes("getByTitle('Tasks')")));
    assert.ok(f.calls.some(r => r.url.includes(`/lists(guid'${listId}')/fields`)));
    assert.equal(f.subject.props.list, 'Tasks');
});
test('5a retained selected-model Copy Empty Clear Refresh callbacks cannot mutate replacement Subject', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker');
    f.click('Start');
    await f.ready();
    f.open();
    f.search('Alice');
    await f.wait();
    f.choose('Alice');
    await f.flush();
    const old = ['Start', 'Use confirmed models as consumer options', 'Empty consumer options', 'Clear controlled selection', 'Refresh / retry'].map(label => f.action(label));
    f.reset();
    await f.flush(() => old.forEach(action => action()));
    assert.equal(f.diagnostics.started, false);
    f.click('Start');
    await f.ready();
    f.open();
    f.search('Bob');
    await f.wait();
    f.choose('Bob');
    await f.flush();
    f.click('Use confirmed models as consumer options');
    await f.flush();
    const key = f.diagnostics.selectedOptions[0], calls = f.calls.length;
    await f.flush(() => old.forEach(action => action()));
    assert.equal(f.diagnostics.selectedOptions[0], key);
    assert.equal(f.diagnostics.consumerOptions[0].value.loginName, 'bob@tenant.test');
    assert.equal(f.calls.length, calls);
    assert.equal(f.subject.mounts, 2);
});
test('5a retired ensureUser rejection cannot publish through latest replacement callback', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker', { hold: r => r.url.endsWith('/ensureuser') });
    f.change('User IDs', 'sharePoint');
    f.toggle('Allow ensureUser on explicit confirmation');
    f.click('Start');
    await f.ready();
    f.open();
    f.search('Alice');
    await f.wait();
    f.choose('Alice');
    await f.flush();
    f.reset();
    f.click('Start');
    await f.ready();
    f.open();
    f.search('Bob');
    await f.wait();
    f.choose('Bob');
    await f.flush();
    const key = f.diagnostics.selectedOptions[0];
    await f.settle(f.pending[0], new Error('Retired ensure failed'));
    assert.equal(f.diagnostics.selectedOptions[0], key);
    assert.equal(f.diagnostics.selectedValues[0]?.loginName, 'bob@tenant.test');
    assert.equal(f.diagnostics.error, '');
});
test('5a actual Explorer Reset keeps native focus and outer Demo while retiring Principal Subject', async (t) => {
    const f = createReactPickersFixture(t, 'PrincipalPicker', { explorer: true });
    await f.flush();
    assert.equal(f.calls.length + f.initializations.length, 0);
    const header = f.container.querySelector('[aria-label="Suggestion header"]');
    assert.ok(header);
    f.click('Start');
    await f.ready();
    assert.equal(f.subject.mounts, 1);
    const reset = [...f.container.querySelectorAll('button')].find(node => node.textContent === 'Reset scenario');
    assert.ok(reset);
    f.h.act(() => { reset.focus(); reset.click(); });
    await f.flush();
    assert.equal(document.activeElement === reset, true);
    assert.equal(f.container.querySelector('[aria-label="Suggestion header"]') === header, true);
    assert.equal(f.subject.unmounts, 1);
    assert.equal(f.trigger() === null, true);
    assert.match(f.container.textContent, /Not started/);
    assert.equal(f.calls.length, 0);
});

test('5a actual opposite-mode control load is rejected by the Node origin witness', t => {
    const f = createReactPickersFixture(t, 'PrincipalPicker');
    const path = require('node:path');
    const opposite = path.join(path.dirname(f.root), f.compiled ? 'src' : 'lib', 'controls/pickers/field-picker' + (f.compiled ? '.tsx' : '.js'));
    f.h.load(opposite);
    assert.equal(f.seen.includes(opposite), true, 'a real opposite control and its dependencies were loaded');
    assert.throws(() => f.origins(), assert.AssertionError, 'origin witness must inspect both source and compiled paths');
});

for (const [name, query, label] of [
    ['SitePicker', 'Site 000', 'Site 000'],
    ['ListItemPicker', 'quarterly', 'Quarterly-Report.docx'],
    ['PrincipalPicker', 'Alice', 'Alice'],
]) test('5a independent confirmed consumer candidates survive clear and empty options retains selection ' + name, async t => {
    const f = createReactPickersFixture(t, name);
    prepare(f, name);
    f.click('Start');
    await f.ready();
    f.open();
    f.search(query);
    if (name === 'SitePicker') { await f.wait(380); await f.wait(80); }
    else await f.wait();
    f.choose(label);
    await f.flush();
    const key = f.diagnostics.selectedOptions[0];
    assert.ok(key);
    f.click('Use confirmed models as consumer options');
    f.click('Clear controlled selection');
    f.open();
    f.search('');
    await f.flush();
    assert.deepEqual(f.diagnostics.selectedOptions, []);
    assert.equal(f.diagnostics.consumerOptions[0].key, key);
    assert.equal(options().filter(row => row.textContent.includes(label)).length, 1, 'confirmed candidate survives independently of remote query and selection');
    f.choose(label);
    await f.flush();
    f.click('Empty consumer options');
    await f.flush();
    assert.equal(f.diagnostics.selectedOptions[0], key);
    assert.equal(f.diagnostics.selectedValues.length, 1);
});

for (const [name] of cases) test('5a standalone fallback is complete and Start-gated ' + name, async t => {
    const f = createReactPickersFixture(t, name, { fallback: true, revision: 23 });
    await f.flush();
    const header = f.container.querySelector('[aria-label="Suggestion header"]');
    assert.ok(header);
    assert.ok(f.container.querySelector('pre'), 'standalone diagnostics are available without a private renderer');
    assert.equal(f.calls.length + f.initializations.length, 0);
    assert.equal(f.subject.mounts, 0);
    prepare(f, name);
    f.click('Start');
    await f.ready();
    assert.equal(f.subject.mounts, 1);
    assert.ok(f.trigger(), 'standalone fallback renders the actual native subject');
    const resetCount = f.calls.length;
    f.reset();
    assert.equal(f.container.querySelector('[aria-label="Suggestion header"]') === header, true);
    assert.equal(f.trigger() === null, true);
    assert.equal(f.diagnostics.started, false);
    assert.equal(f.calls.length, resetCount);
});

for (const aba of [false, true]) test('5a real CT source change retires retained commands ' + (aba ? 'including true ABA' : 'before replacement selection'), async t => {
    const f = createReactPickersFixture(t, 'ContentTypePicker');
    f.change('Exact list title', 'Tasks');
    f.click('Start');
    await f.ready();
    f.open();
    await f.flush();
    f.choose('Content 002');
    await f.flush();
    const original = ['Start', 'Use confirmed models as consumer options', 'Empty consumer options', 'Clear controlled selection', 'Refresh / retry'].map(label => f.action(label));
    const trigger = f.trigger();
    f.change('Content type scope', 'list');
    if (aba) f.change('Content type scope', 'site');
    assert.deepEqual(f.diagnostics.selectedOptions, []);
    assert.equal(f.diagnostics.consumerOptions, undefined);
    await f.flush(() => original[1]());
    assert.equal(f.diagnostics.consumerOptions === undefined, true, 'a retired Copy cannot restore models from the old effective binding');
    f.open();
    await f.flush();
    f.choose('Content 004');
    await f.flush();
    f.click('Use confirmed models as consumer options');
    const key = f.diagnostics.selectedOptions[0], candidate = f.diagnostics.consumerOptions[0].key, refresh = f.subject.props.refreshKey;
    await f.flush(() => original.forEach(command => command()));
    assert.equal(f.diagnostics.selectedOptions[0], key);
    assert.equal(f.diagnostics.consumerOptions[0].key, candidate);
    assert.equal(f.subject.props.refreshKey, refresh);
    assert.equal(f.trigger() === trigger, true);
    assert.equal(f.subject.mounts, 1, 'effective source edits retain the operational Subject');
});

for (const [name] of cases) test('5a effective source binding ABA retires unarmed commands ' + name, async t => {
    const f = createReactPickersFixture(t, name);
    prepare(f, name);
    if (name === 'ContentTypePicker') f.change('Exact list title', 'Tasks');
    const original = ['Start', 'Use confirmed models as consumer options', 'Empty consumer options', 'Clear controlled selection', 'Refresh / retry'].map(label => f.action(label));
    if (name === 'SitePicker') { f.toggle('Offer current site collection'); f.toggle('Offer current site collection'); }
    if (name === 'ListPicker') { f.change('Web URL', 'https://tenant.test/sites/other'); f.change('Web URL', ''); }
    if (name === 'FieldPicker') { f.change('Exact list title', 'Tasks'); f.change('Exact list title', ''); }
    if (name === 'ContentTypePicker') { f.change('Content type scope', 'list'); f.change('Content type scope', 'site'); }
    if (name === 'ViewPicker') { f.change('Exact list title', 'Docs'); f.change('Exact list title', 'Tasks'); }
    if (name === 'ListItemPicker') { f.change('Item display field', 'FileLeafRef'); f.change('Item display field', 'Title'); }
    if (name === 'PrincipalPicker') { f.change('User IDs', 'sharePoint'); f.change('User IDs', 'available'); }
    await f.flush(() => original.forEach(command => command()));
    assert.equal(f.subject.mounts, 0);
    assert.equal(f.calls.length + f.initializations.length, 0);
    assert.equal(f.diagnostics.started, false);
    assert.equal(f.diagnostics.consumerOptions === undefined, true);
    f.click('Start');
    await f.ready();
    assert.equal(f.subject.mounts, 1);
    assert.equal(f.subject.props.refreshKey, 0, 'retired Refresh cannot advance the current source revision');
});

test('5a same-value normalized equivalent and cosmetic CT edits preserve original commands and native Subject', async t => {
    const f = createReactPickersFixture(t, 'ContentTypePicker');
    const start = f.action('Start');
    f.change('Content type scope', 'site');
    f.change('Web URL', ' HTTPS://TENANT.TEST:443/sites/root/sub/// ');
    f.change('Exact list title', 'unused while site scoped');
    f.change('Suggestion header', 'Cosmetic before Start');
    f.toggle('Custom search and status text');
    await f.flush(() => start());
    assert.equal(f.subject.mounts, 1, 'equivalent effective source retains the original Start command');
    f.open();
    await f.flush();
    f.choose('Content 002');
    await f.flush();
    const copy = f.action('Use confirmed models as consumer options'), key = f.diagnostics.selectedOptions[0], trigger = f.trigger(), calls = f.calls.length;
    f.change('Web URL', 'https://tenant.test/sites/root/sub');
    f.change('Suggestion header', 'Another cosmetic label');
    f.change('Presentation', 'combobox');
    await f.flush(() => copy());
    assert.equal(f.diagnostics.consumerOptions[0].key, key, 'equivalent edits retain the original Copy command');
    assert.equal(f.diagnostics.selectedOptions[0], key);
    assert.equal(f.trigger() === trigger, true);
    assert.equal(f.subject.mounts, 1);
    assert.equal(f.calls.length, calls);
});

test('5a finite native list targets reject unknown titles IDs and foreign metadata scopes', () => {
    const native = require('./fixtures/executable-examples/react-pickers-native.cjs');
    const request = path => ({ kind: 'sp', method: 'GET', url: 'https://tenant.test' + path });
    for (const path of [
        "/sites/root/sub/_api/web/lists/getByTitle('Missing')?$select=Id,Title,BaseTemplate",
        "/sites/root/sub/_api/web/lists(guid'99999999-9999-4999-8999-999999999999')?$select=Id,Title,BaseTemplate,Hidden,RootFolder/ServerRelativeUrl&$expand=RootFolder",
        "/sites/unknown/_api/web/fields?$select=Id,InternalName,Title,TypeAsString,FieldTypeKind,Hidden,ReadOnlyField,Group&$orderby=Title%20asc",
    ]) assert.throws(() => native.route(request(path)), /finite native/);
});
