const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { createHarness } = require('./services-test-harness.cjs');
const { createTransport } = require('./fixtures/pnp-list-transport.cjs');
const { Simulate } = require('react-dom/test-utils');
const { FluentProvider, webLightTheme } = require('@fluentui/react-components');

// RefinementName is presentation text; RefinementToken is an FQL operand.
const pdfToken = '"ǂǂ706466"';
const authorToken = '"ǂǂ466162696f204672616e7a696e69"';
const docxToken = '"ǂǂ646f6378"';
const sizeToken = 'range(100,1000,from="GE",to="LT")';
const writeToken = 'range(2026-01-01T00:00:00Z,2027-01-01T00:00:00Z)';
function response() {
  return {
    PrimaryQueryResult: {
      RelevantResults: {
        RowCount: 1, TotalRows: 2, TotalRowsIncludingDuplicates: 2,
        Table: { Rows: [{ Cells: [
          { Key: 'DocId', Value: '1', ValueType: 'Edm.Int64' },
          { Key: 'Title', Value: 'PDF document', ValueType: 'Edm.String' },
          { Key: 'Path', Value: 'https://tenant.test/document.pdf', ValueType: 'Edm.String' },
          { Key: 'FileType', Value: 'pdf', ValueType: 'Edm.String' },
          { Key: 'Author', Value: 'Fabio Franzini', ValueType: 'Edm.String' }
        ] }] }
      },
      RefinementResults: { Refiners: [{ Name: 'FileType', Entries: [{
        RefinementName: 'pdf', RefinementValue: 'pdf',
        RefinementCount: '2', RefinementToken: pdfToken
      }, {
        RefinementName: 'docx', RefinementValue: 'docx',
        RefinementCount: '2', RefinementToken: docxToken
      }] }, { Name: 'DisplayAuthor', Entries: [{
        RefinementName: 'Fabio Franzini', RefinementValue: 'Fabio Franzini',
        RefinementCount: '2', RefinementToken: authorToken
      }] }, { Name: 'Size', Entries: [{
        RefinementName: 'Small', RefinementValue: 'Small',
        RefinementCount: '2', RefinementToken: sizeToken
      }] }, { Name: 'Write', Entries: [{
        RefinementName: '2026', RefinementValue: '2026',
        RefinementCount: '2', RefinementToken: writeToken
      }] }] }
    }
  };
}

function fixture(t) {
  const boundaries = {};
  const h = createHarness(boundaries);
  t.after(h.close);
  // JSDOM 15 lacks Web Crypto, required by Fluent's real focus manager.
  Object.defineProperty(window, 'crypto', { configurable: true, value: require('node:crypto').webcrypto });
  // Keep PnP's builder, HTTP serialization and result parser real; replace HTTP only.
  const transport = createTransport(h, 'https://tenant.test/sites/search',
    Array.from({ length: 24 }, response));
  const context = { sp: transport.sp, isInitialized: true };
  boundaries['./useSPFxPnPContext'] = { useSPFxPnPContext: () => context };
  const { useSPFxPnPSearch } = h.load('hooks/useSPFxPnPSearch.ts');
  const { createSPFxPnPSearchService } = h.load('services/spfx-pnp-search.service.ts');
  return { ...h, ...transport, boundaries, useSPFxPnPSearch, createSPFxPnPSearchService,
    request: index => JSON.parse(transport.calls[index].init.body).request };
}

for (const [name, value, expected] of [
  ['FileType', pdfToken, 'FileType:"ǂǂ706466"'],
  ['FileType', 'ǂǂ706466', 'FileType:"ǂǂ706466"'],
  ['DisplayAuthor', authorToken, 'DisplayAuthor:"ǂǂ466162696f204672616e7a696e69"'],
  ['Author', '"ǂǂ666162696f6672616e7a696e69"', 'Author:"ǂǂ666162696f6672616e7a696e69"'],
  ['FileType', 'pdf', 'FileType:equals("pdf")'],
  ['Author', "O'Brien", 'Author:equals("O\'Brien")'],
  ['Size', sizeToken, 'Size:range(100,1000,from="GE",to="LT")'],
  ['Size', 'range(min,100)', 'Size:range(min,100)'],
  ['Size', 'range(1000,max)', 'Size:range(1000,max)'],
  ['Write', writeToken, 'Write:range(2026-01-01T00:00:00Z,2027-01-01T00:00:00Z)'],
  ['Write', 'datetime(2026-01-01T00:00:00Z)', 'Write:datetime(2026-01-01T00:00:00Z)'],
  ['Write', '2026-01-01T00:00:00Z', 'Write:2026-01-01T00:00:00Z'],
  ['RefinableInt00', '0', 'RefinableInt00:0'],
  ['RefinableDecimal00', '-12.5', 'RefinableDecimal00:-12.5'],
  ['RefinableInt00', 'int(42)', 'RefinableInt00:int(42)'],
  ['RefinableString00', '"東京 Café"', 'RefinableString00:"東京 Café"'],
  ['RefinableString00', '"ǂǂe69db1e4baac"', 'RefinableString00:"ǂǂe69db1e4baac"'],
  ['RefinableString00', '"a\\"b"', 'RefinableString00:"a\\"b"'],
  ['RefinableString00', 'a"b\\c', 'RefinableString00:equals("a\\"b\\\\c")'],
  ['RefinableString00', '', 'RefinableString00:equals("")'],
  ['owstaxidmetadataalltagsinfo', '"ǂǂ4750307c23313233"', 'owstaxidmetadataalltagsinfo:"ǂǂ4750307c23313233"'],
  ['RefinableString00', 'or("red","blue")', 'RefinableString00:or("red","blue")'],
  ['RefinableString00', 'string("red blue",mode="AND")', 'RefinableString00:string("red blue",mode="AND")']
]) {
  test(`search service preserves refiner operands and legacy text for ${name}: ${value}`, async t => {
    const f = fixture(t);
    const service = f.createSPFxPnPSearchService(f.sp, { refiners: 'FileType' });
    const result = await service.search('*', {
      refinementFilters: new Map([[name, [value]]])
    });
    assert.deepEqual(f.request(0).RefinementFilters, [expected]);
    assert.equal(f.calls[0].url, 'https://tenant.test/sites/search/_api/search/postquery');
    assert.equal(f.request(0).Querytext, '*');
    assert.equal(result.refiners[0].entries[0].token, pdfToken);
  });
}

for (const operator of [undefined, 'and', 'or']) {
  test(`same-facet values use ${operator ?? 'default AND'} and distinct facets remain separate`, async t => {
    const f = fixture(t);
    const service = f.createSPFxPnPSearchService(f.sp, { refinerOperator: operator });
    await service.search('*', { refinementFilters: new Map([
      ['FileType', [pdfToken, docxToken]], ['DisplayAuthor', [authorToken]], ['Size', []]
    ]) });
    assert.deepEqual(f.request(0).RefinementFilters, operator === 'or'
      ? ['FileType:or("ǂǂ706466","ǂǂ646f6378")', 'DisplayAuthor:"ǂǂ466162696f204672616e7a696e69"']
      : ['FileType:"ǂǂ706466"', 'FileType:"ǂǂ646f6378"', 'DisplayAuthor:"ǂǂ466162696f204672616e7a696e69"']);
  });
}

test('request operator overrides the service default and groups native ranges with OR', async t => {
  const f = fixture(t);
  const service = f.createSPFxPnPSearchService(f.sp, { refinerOperator: 'and' });
  await service.search('*', { refinerOperator: 'or', refinementFilters: new Map([
    ['Size', ['range(min,100)', sizeToken]], ['Write', [writeToken]]
  ]) });
  assert.deepEqual(f.request(0).RefinementFilters, [
    'Size:or(range(min,100),range(100,1000,from="GE",to="LT"))',
    'Write:range(2026-01-01T00:00:00Z,2027-01-01T00:00:00Z)'
  ]);
  await service.search('*', { refinementFilters: new Map([['Size', []]]) });
  assert.deepEqual(f.request(1).RefinementFilters, []);
});

test('legacy refiner helper also preserves native bins without changing its signature', t => {
  const f = fixture(t);
  const { buildRefinementFilters } = f.load('hooks/useSPFxPnPSearch.results.internal.ts');
  assert.deepEqual(buildRefinementFilters(new Map([
    ['FileType', [pdfToken]], ['Size', [sizeToken]], ['RefinableString00', ['red']]
  ])), ['FileType:"ǂǂ706466"', 'Size:range(100,1000,from="GE",to="LT")', 'RefinableString00:equals("red")']);
});

test('OR hook retains all selected tokens on refetch, pages, and removing one value', async t => {
  const f = fixture(t);
  const view = f.mount(() => f.useSPFxPnPSearch({ pageSize: 1, refinerOperator: 'or' }));
  await f.act(async () => { await view.current.search('*'); });
  await f.act(async () => { await view.current.applyRefiner('FileType', pdfToken); });
  await f.act(async () => { await view.current.applyRefiner('FileType', docxToken); });
  await f.act(async () => { await view.current.applyRefiner('Size', sizeToken); });
  await f.act(async () => { await view.current.refetch(); });
  await f.act(async () => { await view.current.loadMore(); });
  for (const index of [3, 4, 5]) assert.deepEqual(f.request(index).RefinementFilters, [
    'FileType:or("ǂǂ706466","ǂǂ646f6378")', 'Size:range(100,1000,from="GE",to="LT")'
  ]);
  await f.act(async () => { await view.current.applyRefiner('FileType', pdfToken); });
  assert.deepEqual(f.request(6).RefinementFilters, [
    'FileType:"ǂǂ646f6378"', 'Size:range(100,1000,from="GE",to="LT")'
  ]);
});

test('operator change clears search identity while equivalent inline options preserve it', async t => {
  const f = fixture(t);
  const view = f.mount(p => f.useSPFxPnPSearch({ refinerOperator: p.operator }), { operator: 'or' });
  await f.act(async () => { await view.current.search('*'); });
  await f.act(async () => { await view.current.applyRefiner('FileType', pdfToken); });
  view.render({ operator: 'or' });
  await f.act(async () => { await view.current.refetch(); });
  assert.deepEqual(f.request(2).RefinementFilters, ['FileType:"ǂǂ706466"']);
  view.render({ operator: 'and' });
  assert.deepEqual(view.current.results, []);
  assert.deepEqual(view.current.refiners, []);
  await assert.rejects(async () => f.act(async () => { await view.current.refetch(); }), /No previous search/);
  await f.act(async () => { await view.current.search('*'); });
  assert.deepEqual(f.request(3).RefinementFilters, []);
});

test('refiner token survives refetch and pagination, toggles off and clears on a new search', async t => {
  const f = fixture(t);
  const view = f.mount(() => f.useSPFxPnPSearch({ pageSize: 1, refiners: 'FileType' }));
  await f.act(async () => { await view.current.search('*'); });
  assert.deepEqual(f.request(0).RefinementFilters, []);
  const entry = view.current.refiners[0].entries[0];
  await f.act(async () => { await view.current.applyRefiner('FileType', entry.token); });
  await f.act(async () => { await view.current.refetch(); });
  await f.act(async () => { await view.current.loadMore(); });
  for (const index of [1, 2, 3]) {
    assert.deepEqual(f.request(index).RefinementFilters, ['FileType:"ǂǂ706466"']);
  }
  assert.equal(f.request(3).StartRow, 1);
  await f.act(async () => { await view.current.applyRefiner('FileType', entry.token); });
  assert.deepEqual(f.request(4).RefinementFilters, []);
  await f.act(async () => { await view.current.applyRefiner('FileType', entry.token); });
  await f.act(async () => { await view.current.search('new query'); });
  assert.deepEqual(f.request(6).RefinementFilters, []);
  assert.equal(f.request(6).Querytext, 'new query');
});

test('Faceted Search demo combines PDF and DisplayAuthor tokens and removes only the toggled facet', async t => {
  const f = fixture(t);
  let pending;
  f.boundaries['@apvee/spfx-react-toolkit'] = { useSPFxPnPSearch: options => {
    const info = f.useSPFxPnPSearch(options);
    return { ...info,
      search: (...args) => (pending = info.search(...args)),
      applyRefiner: (...args) => (pending = info.applyRefiner(...args))
    };
  } };
  const file = path.resolve(__dirname, '../apps/spfx-react-toolkit-test/src/webparts/spFxReactToolkitTest/components/demos/PnPSearchRefinersDemo.tsx');
  const { PnPSearchRefinersDemo } = f.load(file);
  const view = f.mountComponent(() => f.React.createElement(FluentProvider,
    { theme: webLightTheme }, f.React.createElement(PnPSearchRefinersDemo)));
  async function click(text) {
    const button = Array.from(view.container.querySelectorAll('button')).find(b => b.textContent === text);
    assert.ok(button, text);
    await f.act(async () => {
      Simulate.click(button);
      await pending;
    });
  }
  await click('Search All Content');
  assert.equal(f.request(0).Refiners, 'FileType,DisplayAuthor,Size,Write');
  await click('pdf (2)');
  assert.deepEqual(f.request(1).RefinementFilters, ['FileType:"ǂǂ706466"']);
  assert.ok(view.container.textContent.includes('PDF document'));
  await click('Fabio Franzini (2)');
  assert.deepEqual(f.request(2).RefinementFilters, [
    'FileType:"ǂǂ706466"', 'DisplayAuthor:"ǂǂ466162696f204672616e7a696e69"'
  ]);
  await click('Fabio Franzini (2)');
  assert.deepEqual(f.request(3).RefinementFilters, ['FileType:"ǂǂ706466"']);
  await click('pdf (2)');
  assert.deepEqual(f.request(4).RefinementFilters, []);
  await click('pdf (2)');
  await click('docx (2)');
  await click('Small (2)');
  await click('2026 (2)');
  assert.deepEqual(f.request(8).RefinementFilters, [
    'FileType:or("ǂǂ706466","ǂǂ646f6378")',
    'Size:range(100,1000,from="GE",to="LT")',
    'Write:range(2026-01-01T00:00:00Z,2027-01-01T00:00:00Z)'
  ]);
  const select = view.container.querySelector('select');
  assert.ok(select, 'explicit same-facet operator selector');
  f.act(() => Simulate.change(select, { target: { value: 'and' } }));
  assert.ok(!view.container.textContent.includes('PDF document'));
  await click('Search All Content');
  await click('pdf (2)');
  await click('docx (2)');
  assert.deepEqual(f.request(11).RefinementFilters, ['FileType:"ǂǂ706466"', 'FileType:"ǂǂ646f6378"']);
});
