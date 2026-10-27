const test = require('node:test');
const assert = require('node:assert/strict');
const { createCoreHarness } = require('./fixtures/editable-text/core-harness.cjs');
function fixture(t) {
  const h = createCoreHarness();
  t.after(h.close);
  return { api: h.load('editable-text-codec.internal.ts'), policy: h.load('editable-text-policy.internal.ts'), doc: h.ownerDocument };
}

test('undefined and invalid values use literal fallback while empty HTML remains authoritative', t => {
  const { api, policy, doc } = fixture(t);
  const blocks = policy.resolveEditableTextBlockPolicy(undefined, undefined);
  for (const value of [undefined, 'x'.repeat(65537)]) {
    const result = api.resolveEditableTextValue(value, '<tag> & text', doc, blocks);
    assert.equal(result.kind, value === undefined ? 'valid' : 'invalid');
    assert.equal(api.serializeEditableTextDocument(result.document), '<p>&lt;tag&gt; &amp; text</p>');
  }
  const result = api.resolveEditableTextValue('', '<tag>', doc, blocks);
  assert.equal(api.serializeEditableTextDocument(result.document), '');
  assert.deepEqual(result.document, { type: 'paragraph', content: [] });
  assert.equal(doc.body.innerHTML, '<div>Visible host</div>');
});
test('safe parser retains first supported block semantics before applying membership conversion', t => {
  const { api, policy, doc } = fixture(t);
  const html = '<div><h2 class="wps-line-16 wps-space-before"><b>One</b></h2><h3>Two</h3><p>Three</p></div>';
  const parsed = api.parseEditableTextHTML(html, doc, 'p');
  assert.equal(parsed.document.type, 'heading');
  assert.equal(parsed.document.level, 2);
  const converted = api.resolveEditableTextValue(html, '', doc, policy.resolveEditableTextBlockPolicy(['h3'], undefined));
  assert.equal(converted.kind, 'normalized');
  assert.equal(api.serializeEditableTextDocument(converted.document), '<h3 class="wps-line-16 wps-space-before"><strong>One</strong><br>Two<br>Three</h3>');
  assert.equal(api.getEditableTextPlainText(converted.document), 'One Two Three');
});
test('defined canonical excluded blocks normalize while allowed blocks remain canonical', t => {
  const { api, policy, doc } = fixture(t);
  const blocks = policy.resolveEditableTextBlockPolicy(['h2', 'h4'], 'h4');
  assert.equal(api.resolveEditableTextValue('<h2>Title</h2>', '', doc, blocks).kind, 'valid');
  const excluded = api.resolveEditableTextValue('<p>Title</p>', '', doc, blocks);
  assert.equal(excluded.kind, 'normalized');
  assert.equal(api.serializeEditableTextDocument(excluded.document), '<h4>Title</h4>');
  assert.deepEqual(api.parseEditableTextHTML('', doc, 'h4').document, { type: 'heading', level: 4, content: [] });
});
test('core preserves the legacy format, escaping, palette and sanitized owner-document parsing', t => {
  const { api, doc } = fixture(t);
  const input = '<h1 class="wps-space-after" style="line-height:1.4" onclick="bad()"><script>bad</script><span class="fontSizeXLarge fontColorBlue"><b><i><u>&lt;tag&gt; &amp;</u></i></b></span></h1>';
  const result = api.parseEditableTextHTML(input, doc, 'p');
  assert.equal(api.serializeEditableTextDocument(result.document), '<h1 class="wps-line-14 wps-space-after"><strong><em><u><span class="wps-color-brand wps-size-20">&lt;tag&gt; &amp;</span></u></em></strong></h1>');
  assert.equal(api.getEditableTextPlainText(result.document), '<tag> &');
  assert.equal(doc.body.innerHTML, '<div>Visible host</div>');
});
test('core enforces the input, depth and canonical escaping byte limits', t => {
  const { api, doc } = fixture(t);
  for (const html of ['é'.repeat(32769), '<span>'.repeat(33) + 'A' + '</span>'.repeat(33), '<p>' + '&'.repeat(13105) + 'xxxxx</p>']) {
    const result = api.parseEditableTextHTML(html, doc, 'h3');
    assert.equal(result.kind, 'invalid');
    assert.deepEqual(result.document, { type: 'heading', level: 3, content: [] });
  }
  const exact = '<p>' + '&amp;'.repeat(13105) + 'xxxx</p>';
  assert.equal(Buffer.byteLength(exact), 65536);
  assert.equal(api.parseEditableTextHTML(exact, doc, 'p').kind, 'valid');
  const ast = { type: 'paragraph', content: [{ type: 'text', text: '&'.repeat(14000) }] };
  assert.throws(() => api.serializeEditableTextDocument(ast), { name: 'RangeError' });
});

test('policy conversion also respects the canonical byte budget after adding heading tags', t => {
  const { api, policy, doc } = fixture(t);
  const exactParagraph = '<p>' + 'x'.repeat(65529) + '</p>';
  const result = api.resolveEditableTextValue(exactParagraph, '<fallback>', doc, policy.resolveEditableTextBlockPolicy(['h3'], undefined));
  assert.equal(result.kind, 'invalid');
  assert.equal(api.serializeEditableTextDocument(result.document), '<h3>&lt;fallback&gt;</h3>');
});
