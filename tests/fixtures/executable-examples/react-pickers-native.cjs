// Independently specified finite native picker contracts. No scenario text is used as an oracle.
const assert = require('node:assert/strict');
const root = 'https://tenant.test/sites/root', hub = 'https://tenant.test/sites/hub';
const guid = n => `22222222-2222-4222-8222-${String(n).padStart(12, '0')}`;
const listId = guid(1), hubId = guid(2), siteId = guid(3), documentsId = guid(4);
const nativeWebPath = '/sites/root/sub/_api/web';
function isListChild(endpoint, suffix) {
    return [listId, documentsId].some(id => endpoint === `${nativeWebPath}/lists(guid'${id}')/${suffix}`);
}
const lists = Array.from({ length: 45 }, (_, i) => ({ Id: guid(i + 10), Title: `List ${String(i).padStart(2, '0')}`, BaseTemplate: 100, Hidden: false, RootFolder: { ServerRelativeUrl: `/sites/root/sub/lists/${i}` } }));
const fields = Array.from({ length: 210 }, (_, i) => ({ Id: guid(i + 100), InternalName: `Column${i}`, Title: `Column ${String(i).padStart(3, '0')}`, TypeAsString: 'Text', FieldTypeKind: 2, Hidden: i === 209, ReadOnlyField: i === 208, Group: 'Demo' }));
const views = Array.from({ length: 210 }, (_, i) => ({ Id: guid(i + 400), Title: `View ${String(i).padStart(3, '0')}`, DefaultView: i === 0, PersonalView: i === 209, ServerRelativeUrl: `/sites/root/sub/lists/Tasks/view${i}.aspx` }));
const contentTypes = Array.from({ length: 210 }, (_, i) => ({ StringId: i === 0 ? '0x01' : i === 1 ? '0x0101' : '0x0101' + (i + 1).toString(16).padStart(2, '0'), Name: `Content ${String(i).padStart(3, '0')}`, Group: i % 2 ? 'Demo' : 'Other', Description: 'Fixture type', Hidden: i === 209, ReadOnly: i === 208, Sealed: i === 207, Id: { StringValue: 'must never coerce this object' } }));
const items = [{ Id: 1, FSObjType: 0, Title: '', FileLeafRef: 'Quarterly-Report.docx', File: { UniqueId: guid(700), Name: 'Quarterly-Report.docx', ServerRelativeUrl: '/sites/root/sub/Docs/Quarterly-Report.docx' } }, { Id: 2, FSObjType: 0, Title: 'Alpha document', FileLeafRef: 'alpha.pdf', File: { UniqueId: guid(701), Name: 'alpha.pdf', ServerRelativeUrl: '/sites/root/sub/Docs/alpha.pdf' } }, { Id: 3, FSObjType: 1, Title: 'Folder', Folder: { UniqueId: guid(702), ServerRelativeUrl: '/sites/root/sub/Docs/Folder' } }];
const sites = Array.from({ length: 405 }, (_, i) => ({ Title: `Site ${String(404 - i).padStart(3, '0')}`, Path: `https://tenant.test/sites/site-${404 - i}`, SiteId: guid(1000 + 404 - i), DepartmentId: hubId }));
function json(value, status = 200) { return new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json' } }); }
function exactParams(url, values, optional = []) { const actual = Object.fromEntries(url.searchParams); for (const [key, value] of Object.entries(values))
    assert.equal(actual[key], value, key); assert.ok(Object.keys(actual).every(key => Object.hasOwn(values, key) || optional.includes(key)), 'finite query keys ' + url); }
function metadataPage(url, rows) { const offset = Number(url.searchParams.get('$skiptoken') ?? 0), top = Number(url.searchParams.get('$top')); assert.ok(top > 0 && top <= 200); assert.ok([0, 120].includes(offset), 'native metadata continuation offset'); const filtered = url.searchParams.get('$filter')?.match(/Id eq guid'([^']+)'/); if (filtered)
    return json({ value: rows.filter(row => row.Id === filtered[1]) }); const next = offset === 0 && rows.length > 120 ? `${url.origin}${url.pathname}?$skiptoken=120` : undefined; return json({ value: rows.slice(offset, offset ? undefined : 120), ...(next ? { 'odata.nextLink': next } : {}) }); }
function acceptsSiteSearchPost(target, init, origin = 'https://tenant.test') {
    if (target.origin !== origin || !target.pathname.endsWith('/_api/search/postquery') || target.search || init?.method !== 'POST')
        return false;
    const headers = new Headers(init.headers);
    if (!headers.get('X-RequestDigest') || !headers.get('Content-Type')?.includes('application/json'))
        return false;
    let body;
    try {
        body = JSON.parse(init.body);
    }
    catch {
        return false;
    }
    if (!body || Object.keys(body).join() !== 'request')
        return false;
    const q = body.request;
    if (!q || Object.keys(q).some(key => !['Querytext', 'RowLimit', 'StartRow', 'HitHighlightedProperties', 'Properties', 'RefinementFilters', 'ReorderingRules', 'SelectProperties', 'SortList', 'TrimDuplicates'].includes(key)))
        return false;
    return typeof q.Querytext === 'string' && /^contentclass:STS_Site(?: AND (?:\(Title:"[^"\n]*\*" OR Path:"[^"\n]*\*"\)|Path="https:\/\/tenant\.test\/sites\/[\w/-]+"))?$/.test(q.Querytext) && q.RowLimit === 200 && (q.StartRow === undefined || [200, 400, 600, 800].includes(q.StartRow)) && typeof q.TrimDuplicates === 'boolean' && JSON.stringify(q.SelectProperties) === '["Title","Path","SiteId","DepartmentId"]' && ['HitHighlightedProperties', 'Properties', 'RefinementFilters', 'ReorderingRules', 'SortList'].every(key => Array.isArray(q[key]) && q[key].length === 0);
}
function route(request) {
    const url = new URL(request.url), endpoint = decodeURIComponent(url.pathname), lower = endpoint.toLowerCase();
    if (request.kind === 'pnp') {
        assert.equal(url.origin, 'https://tenant.test');
        if (lower.endsWith('/_api/contextinfo')) {
            assert.equal(request.method, 'POST');
            assert.equal(url.search, '');
            assert.equal(new Headers(request.init.headers).get('X-PnPjs-NoDigest'), '1');
            return json({ FormDigestValue: 'picker-fixture-digest', FormDigestTimeoutSeconds: 1800 });
        }
        assert.ok(acceptsSiteSearchPost(url, request.init), 'exact separate Site native Search shape');
        const q = JSON.parse(request.init.body).request;
        const exact = /Path="([^"]+)"/.exec(q.Querytext), query = /Title:"([^"*]*)\*"/.exec(q.Querytext)?.[1].toLowerCase() ?? '';
        const candidates = exact ? [...sites, { Title: 'Current collection', Path: root, SiteId: siteId, DepartmentId: hubId }, { Title: 'Associated hub', Path: hub, SiteId: hubId, DepartmentId: hubId }].filter(row => row.Path === exact[1]) : sites.filter(row => row.Title.toLowerCase().includes(query) || row.Path.toLowerCase().includes(query));
        const start = q.StartRow ?? 0, rows = candidates.slice(start, start + 200);
        return json({ PrimaryQueryResult: { RelevantResults: { RowCount: rows.length, TotalRows: candidates.length, TotalRowsIncludingDuplicates: candidates.length, Table: { Rows: rows.map(row => ({ Cells: Object.entries(row).map(([Key, Value]) => ({ Key, Value })) })) } } } });
    }
    if (request.kind !== 'sp')
        return undefined;
    if (lower.endsWith('/_api/site/rootweb')) {
        assert.equal(request.method, 'GET');
        exactParams(url, { '$select': 'Url' });
        return json({ Url: root });
    }
    if (lower.endsWith('/_api/site') && url.searchParams.get('$select') === 'Id,Url,HubSiteId,IsHubSite') {
        assert.equal(request.method, 'GET');
        return json({ Id: endpoint.startsWith('/sites/hub/') ? hubId : siteId, Url: endpoint.startsWith('/sites/hub/') ? hub : root, HubSiteId: endpoint.startsWith('/sites/hub/') ? '00000000-0000-0000-0000-000000000000' : hubId, IsHubSite: endpoint.startsWith('/sites/hub/') });
    }
    if (lower.endsWith('/_api/web') && url.searchParams.get('$select') === 'Title,Id') {
        assert.equal(request.method, 'GET');
        return json({ Title: endpoint.startsWith('/sites/hub/') ? 'Associated hub' : 'Current collection', Id: guid(999) });
    }
    if (lower.includes('/_api/hubsites/getbyid(')) {
        assert.equal(request.method, 'GET');
        assert.equal(endpoint, `/sites/root/_api/HubSites/GetById('${hubId}')`);
        return json({ ID: hubId, SiteUrl: hub, Title: 'Associated hub' });
    }
    if (lower.endsWith('/_api/web/lists')) {
        assert.equal(endpoint, nativeWebPath + '/lists', 'finite native catalog scope');
        assert.equal(request.method, 'GET');
        exactParams(url, { '$select': 'Id,Title,BaseTemplate,Hidden,RootFolder/ServerRelativeUrl', '$expand': 'RootFolder', '$orderby': 'Title asc' }, ['$top', '$filter', '$skiptoken']);
        return metadataPage(url, lists);
    }
    if (/\/_api\/web\/lists\/(?:getByTitle)\('[^']+'\)$/.test(endpoint) || /\/_api\/web\/lists\(guid'[^']+'\)$/.test(endpoint)) {
        assert.equal(request.method, 'GET');
        const title = /getByTitle\('([^']+)'\)$/.exec(endpoint)?.[1];
        const id = /lists\(guid'([^']+)'\)$/.exec(endpoint)?.[1];
        const targets = [{ Id: listId, Title: 'Tasks', BaseTemplate: 100 }, { Id: documentsId, Title: 'Docs', BaseTemplate: 101 }, ...lists];
        const target = targets.find(row => title === row.Title || id === row.Id);
        assert.ok(endpoint.startsWith(nativeWebPath + '/lists') && target, 'finite native list target');
        const select = url.searchParams.get('$select');
        if (select === 'Id,Title,BaseTemplate') {
            exactParams(url, { '$select': select });
            return json({ Id: target.Id, Title: target.Title, BaseTemplate: target.BaseTemplate });
        }
        if (select === 'Id,Title,BaseTemplate,Hidden,RootFolder/ServerRelativeUrl') {
            exactParams(url, { '$select': select, '$expand': 'RootFolder' });
            assert.ok(lists.includes(target), 'finite native ListPicker confirmation target');
            return json(target);
        }
    }
    if (lower.endsWith('/fields') && !lower.includes('/fields/')) {
        assert.ok(endpoint === nativeWebPath + '/fields' || isListChild(endpoint, 'fields'), 'finite native field scope');
        assert.equal(request.method, 'GET');
        exactParams(url, { '$select': 'Id,InternalName,Title,TypeAsString,FieldTypeKind,Hidden,ReadOnlyField,Group', '$orderby': 'Title asc' }, ['$top', '$filter', '$skiptoken']);
        return metadataPage(url, fields);
    }
    if (lower.endsWith('/views')) {
        assert.equal(request.method, 'GET');
        assert.ok(isListChild(endpoint, 'views'), 'finite native view scope after actual list resolution');
        exactParams(url, { '$select': 'Id,Title,DefaultView,PersonalView,ServerRelativeUrl', '$orderby': 'Title asc' }, ['$top', '$filter', '$skiptoken']);
        return metadataPage(url, views);
    }
    if (lower.includes('/fields/getbyinternalnameortitle(')) {
        assert.ok([listId, documentsId].some(id => endpoint.startsWith(`${nativeWebPath}/lists(guid'${id}')/fields/`)), 'finite native display-field scope');
        assert.equal(request.method, 'GET');
        const field = /'([^']+)'/.exec(endpoint.slice(endpoint.indexOf('/fields/')))?.[1];
        assert.ok(['Title', 'FileLeafRef'].includes(field));
        exactParams(url, { '$select': 'InternalName,EntityPropertyName,TypeAsString,FieldTypeKind' });
        return json({ InternalName: field, EntityPropertyName: field, TypeAsString: 'Text', FieldTypeKind: 2 });
    }
    if (lower.endsWith('/items')) {
        assert.ok(isListChild(endpoint, 'items'), 'finite native item scope');
        assert.equal(request.method, 'GET');
        const select = url.searchParams.get('$select'), field = select?.split(',')[2];
        assert.ok(['Title', 'FileLeafRef'].includes(field));
        exactParams(url, { '$select': `Id,FSObjType,${field},File/UniqueId,File/Name,File/ServerRelativeUrl,Folder/UniqueId,Folder/ServerRelativeUrl`, '$expand': 'File,Folder', '$orderby': 'Id asc' }, ['$filter', '$top', '$skiptoken']);
        assert.ok(/^\(FSObjType eq 0\)(?: and \(Id eq \d+\))?$/.test(url.searchParams.get('$filter')));
        const id = /Id eq (\d+)/.exec(url.searchParams.get('$filter'))?.[1];
        return json({ value: id ? items.filter(row => row.Id === Number(id)) : items });
    }
    if (lower.endsWith('/contenttypes')) {
        assert.equal(request.method, 'GET');
        exactParams(url, { '$select': 'StringId,Name,Group,Description,Hidden,ReadOnly,Sealed' }, ['$top', '$skiptoken']);
        assert.ok(endpoint === '/sites/root/_api/web/contenttypes' || isListChild(endpoint, 'contenttypes'), 'finite native content-type scope');
        return metadataPage(url, contentTypes);
    }
    if (lower.includes('/contenttypes/getbyid(')) {
        const parent = endpoint.slice(0, endpoint.indexOf('/contenttypes/') + '/contenttypes'.length);
        assert.ok(parent === '/sites/root/_api/web/contenttypes' || isListChild(parent, 'contenttypes'), 'finite native content-type confirmation scope');
        assert.equal(request.method, 'GET');
        exactParams(url, { '$select': 'StringId,Name,Group,Description,Hidden,ReadOnly,Sealed' });
        const id = /'([^']+)'/.exec(endpoint.slice(endpoint.indexOf('/contenttypes/')))?.[1];
        return json(contentTypes.find(row => row.StringId.toLowerCase() === id) ?? {}, contentTypes.some(row => row.StringId.toLowerCase() === id) ? 200 : 404);
    }
    return undefined;
}
module.exports = { route, acceptsSiteSearchPost, lists, fields, views, contentTypes, items, sites, guid, listId, hubId, siteId };
