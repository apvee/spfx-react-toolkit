const test = require('node:test');
const assert = require('node:assert/strict');
const { JSDOM } = require('jsdom');
const { createHarness } = require('./services-test-harness.cjs');
const native = require('./fixtures/web-part-shell/native-title.json');

function fixture(t) {
  const h = createHarness();
  const dom = new JSDOM('<!doctype html><body><div>Visible content</div></body>');
  t.after(() => { h.close(); dom.window.close(); });
  return { api: h.load('controls/web-part-shell/title/web-part-title-codec.internal.ts'), doc: dom.window.document };
}

test('HTML undefined uses literal plain H2 without changing host content or inputs', t => {
  const { api, doc } = fixture(t);
  const props = Object.freeze({ title: '<b>Literal</b> & text', titleHTML: undefined });
  const result = api.resolveWebPartTitle(props.title, props.titleHTML, doc);
  assert.equal(result.kind, 'valid');
  assert.deepEqual(result.document, { type: 'heading', level: 2, content: [{ type: 'text', text: '<b>Literal</b> & text' }] });
  assert.equal(api.serializeTitleDocument(result.document), '<h2>&lt;b&gt;Literal&lt;/b&gt; &amp; text</h2>');
  assert.equal(doc.body.innerHTML, '<div>Visible content</div>');
  assert.equal(props.titleHTML, undefined);
});

test('defined empty HTML is authoritative including whitespace-only documents', t => {
  const { api, doc } = fixture(t);
  for (const html of ['', '<h2> \n&nbsp;<br></h2>']) {
    const result = api.resolveWebPartTitle('Old title', html, doc);
    assert.equal(api.getTitlePlainText(result.document), '');
    assert.equal(api.isTitleDocumentEmpty(result.document), true);
    assert.equal(api.serializeTitleDocument(result.document), '');
  }
});

test('native title preserves heading, allowed line spacing, B/I/U and intentional brand alias', t => {
  const { api, doc } = fixture(t);
  const result = api.parseTitleHTML(native.html, doc);
  assert.equal(result.kind, 'normalized');
  assert.deepEqual(result.document, { type: 'heading', level: 2, lineSpacing: 1.4, content: [
    { type: 'text', text: 'Titolo ' },
    { type: 'text', text: 'News', marks: { bold: true, italic: true, underline: true, color: 'brand' } }
  ] });
  assert.equal(api.getTitlePlainText(result.document), native.plain);
  assert.equal(api.serializeTitleDocument(result.document), native.canonicalHTML);
});

test('canonical roundtrip is idempotent across finite markers and mark ordering', t => {
  const { api, doc } = fixture(t);
  const html = '<h4 class="wps-line-24 wps-space-after wps-space-before"><strong><em><u><span class="wps-color-subtle wps-size-96">A &amp; &lt;B&gt;</span></u></em></strong><br>C</h4>';
  const first = api.parseTitleHTML(html, doc);
  assert.equal(first.kind, 'valid');
  assert.equal(api.serializeTitleDocument(first.document), html);
  assert.deepEqual(api.parseTitleHTML(html, doc).document, first.document);
  assert.equal(api.getTitlePlainText(first.document), 'A & <B> C');
});

test('multiblock normalizes into hard breaks retaining the first allowed block attributes', t => {
  const { api, doc } = fixture(t);
  const result = api.parseTitleHTML('<div><h3 class="wps-line-16 wps-space-before">One</h3><p><b>Two</b><br>Three</p><div>Four</div></div>', doc);
  assert.equal(result.kind, 'normalized');
  assert.equal(api.serializeTitleDocument(result.document), '<h3 class="wps-line-16 wps-space-before">One<br><strong>Two</strong><br>Three<br>Four</h3>');
  assert.equal(api.getTitlePlainText(result.document), 'One Two Three Four');
});

test('text escapes canonically while search whitespace derives from the AST', t => {
  const { api, doc } = fixture(t);
  const result = api.parseTitleHTML('<p>A&nbsp;  B\n &quot;C&quot; &amp; &lt;D&gt;<br> E</p>', doc);
  assert.equal(api.getTitlePlainText(result.document), 'A B "C" & <D> E');
  const html = api.serializeTitleDocument(result.document);
  assert.equal(html, '<p>A\u00a0  B\n &quot;C&quot; &amp; &lt;D&gt;<br> E</p>');
  assert.equal(api.serializeTitleDocument(api.parseTitleHTML(html, doc).document), html);
});

test('active markup and arbitrary attributes cannot survive while innocuous wrappers preserve text', t => {
  const { api, doc } = fixture(t);
  const input = '<h2 id="leak" onclick="alert(1)" style="color:#123456;background:url(x);line-height:1.6"><script>bad</script><style>bad</style><svg><text>bad</text></svg><iframe>bad</iframe><object>bad</object><img src="x"><a href="javascript:bad()"><span class="f-hash fontColorRed wps-size-7" style="font-size:50px">Safe</span></a></h2>';
  const result = api.parseTitleHTML(input, doc);
  assert.equal(result.kind, 'normalized');
  assert.equal(api.serializeTitleDocument(result.document), '<h2 class="wps-line-16">Safe</h2>');
  assert.equal(doc.body.innerHTML, '<div>Visible content</div>');
});

test('font sizes and palette are bounded and native line-height accepts only the enum', t => {
  const { api, doc } = fixture(t);
  for (const [classes, expected] of [
    ['wps-size-8 wps-color-primary', '<p><span class="wps-color-primary wps-size-8">X</span></p>'],
    ['wps-size-97 wps-color-custom', '<p>X</p>'],
    ['fontSizeXLarge', '<p><span class="wps-size-20">X</span></p>'],
    ['fontSizeLarge fontSize20', '<p>X</p>'],
    ['wps-size-8.5 wps-size-08', '<p>X</p>'],
  ]) assert.equal(api.serializeTitleDocument(api.parseTitleHTML('<p><span class="' + classes + '">X</span></p>', doc).document), expected);
  for (const [line, expected] of [['1', '10'], ['1.2', '12'], ['1.4', '14'], ['1.6', '16'], ['2', '20'], ['2.4', '24'], ['1.5', undefined], ['24px', undefined], ['140%', undefined]]) {
    assert.equal(api.serializeTitleDocument(api.parseTitleHTML('<h1 style="line-height:' + line + '">X</h1>', doc).document), '<h1' + (expected ? ' class="wps-line-' + expected + '"' : '') + '>X</h1>');
  }
});

test('limits return safe invalid results and resolution falls back to the literal title', t => {
  const { api, doc } = fixture(t);
  for (const input of ['é'.repeat(32769), '<span>'.repeat(33) + 'X' + '</span>'.repeat(33), 42, undefined, null]) {
    const result = api.parseTitleHTML(input, doc);
    assert.equal(result.kind, 'invalid');
    assert.equal(api.serializeTitleDocument(result.document), '');
  }
  assert.notEqual(api.parseTitleHTML('<p>' + 'x'.repeat(65529) + '</p>', doc).kind, 'invalid');
  assert.notEqual(api.parseTitleHTML('<span>'.repeat(32) + 'X' + '</span>'.repeat(32), doc).kind, 'invalid');
  const result = api.resolveWebPartTitle('<literal>', 'x'.repeat(65537), doc);
  assert.equal(result.kind, 'invalid');
  assert.equal(api.serializeTitleDocument(result.document), '<h2>&lt;literal&gt;</h2>');
});

test('serializer validates untrusted documents rather than publishing invalid attributes', t => {
  const { api } = fixture(t);
  for (const document of [undefined, null, {}, { type: 'heading', level: 5, content: [] },
    { type: 'paragraph', content: [{ type: 'text', text: 'X', marks: { size: '20', color: '#fff' } }] },
    { type: 'paragraph', lineSpacing: 1.5, content: [{ type: 'text', text: 'X' }] }]) {
    assert.throws(() => api.serializeTitleDocument(document), /Invalid title document/);
  }
});

test('empty pasted blocks preserve blank lines and first block semantics', t => {
  const { api, doc } = fixture(t);
  for (const [input, expected] of [
    ['<p>One</p><p></p><p>Two</p>', '<p>One<br><br>Two</p>'],
    ['<h1></h1><p>Two</p>', '<h1><br>Two</h1>'],
    ['<p>One</p><p></p>', '<p>One<br></p>']
  ]) assert.equal(api.serializeTitleDocument(api.parseTitleHTML(input, doc).document), expected);
});

test('serialization coalesces equivalent adjacent runs without mutating its document', t => {
  const { api, doc } = fixture(t);
  const document = Object.freeze({ type: 'paragraph', content: Object.freeze([
    Object.freeze({ type: 'text', text: 'A', marks: Object.freeze({ bold: true }) }),
    Object.freeze({ type: 'text', text: 'B', marks: Object.freeze({ bold: true }) })
  ]) });
  const html = api.serializeTitleDocument(document);
  assert.equal(html, '<p><strong>AB</strong></p>');
  assert.equal(api.serializeTitleDocument(api.parseTitleHTML(html, doc).document), html);
  assert.equal(document.content[0].text, 'A');
});

test('unknown document validation rejects sparse content arrays', t => {
  const { api } = fixture(t);
  const malformed = { type: 'paragraph', content: Array(1) };
  assert.equal(api.isTitleDocument(malformed), false);
  assert.throws(() => api.serializeTitleDocument(malformed), /Invalid title document/);
});

test('benign nested wrappers retain effective block boundaries including empty blocks', t => {
  const { api, doc } = fixture(t);
  for (const [input, expected] of [
    ['<span><p>One</p></span><span><p>Two</p></span>', '<p>One<br>Two</p>'],
    ['<span><p>One</p></span><span><span><p></p></span></span><span><p>Two</p></span>', '<p>One<br><br>Two</p>'],
    ['<span><h3></h3></span><span><p>Two</p></span>', '<h3><br>Two</h3>'],
    ['<span><p>One</p></span><span><p></p></span>', '<p>One<br></p>'],
    ['<span><span>Before<p>One</p></span></span><span>After</span>', '<p>Before<br>One<br>After</p>'],
    ['<span>One</span><span>Two</span>', '<h2>OneTwo</h2>']
  ]) {
    const result = api.parseTitleHTML(input, doc);
    assert.equal(api.serializeTitleDocument(result.document), expected);
    assert.equal(api.serializeTitleDocument(api.parseTitleHTML(expected, doc).document), expected);
  }
});

test('escaping expansion uses the canonical UTF-8 budget and remains roundtrip readable', t => {
  const { api, doc } = fixture(t);
  const accepted = '<p>' + '&'.repeat(13105) + 'xxxx</p>'; // 65,536 canonical bytes.
  const canonical = '<p>' + '&amp;'.repeat(13105) + 'xxxx</p>';
  assert.equal(Buffer.byteLength(canonical), 65536);
  const result = api.parseTitleHTML(accepted, doc);
  assert.equal(result.kind, 'normalized');
  assert.equal(api.serializeTitleDocument(result.document), canonical);
  assert.equal(api.parseTitleHTML(canonical, doc).kind, 'valid');
  for (const input of ['<p>' + '&'.repeat(13105) + 'xxxxx</p>', '<h2>' + '&'.repeat(14000) + '</h2>']) {
    const oversized = api.parseTitleHTML(input, doc);
    assert.equal(oversized.kind, 'invalid');
    assert.equal(api.serializeTitleDocument(oversized.document), '');
  }
});

test('mark wrapper expansion rejects canonical HTML that exceeds the readable budget', t => {
  const { api, doc } = fixture(t);
  const accepted = '<p>' + '<b>x</b>y'.repeat(3448) + '</p>'; // 65,519 canonical bytes.
  const result = api.parseTitleHTML(accepted, doc);
  assert.equal(result.kind, 'normalized');
  const canonical = api.serializeTitleDocument(result.document);
  assert.equal(Buffer.byteLength(canonical), 65519);
  assert.equal(api.parseTitleHTML(canonical, doc).kind, 'valid');
  const oversized = api.parseTitleHTML('<p>' + '<b>x</b>y'.repeat(3449) + '</p>', doc);
  assert.equal(oversized.kind, 'invalid');
  assert.equal(api.serializeTitleDocument(oversized.document), '');
});

test('serializer throws RangeError for oversized canonical output without modifying the AST', t => {
  const { api } = fixture(t);
  const document = Object.freeze({ type: 'paragraph', content: Object.freeze([
    Object.freeze({ type: 'text', text: '&'.repeat(14000) })
  ]) });
  assert.equal(api.isTitleDocument(document), true);
  assert.throws(() => api.serializeTitleDocument(document), { name: 'RangeError', message: 'Title HTML exceeds 65536 UTF-8 bytes' });
  assert.equal(document.content[0].text.length, 14000);
  assert.equal(api.getTitlePlainText(document), '&'.repeat(14000));
});

const semanticColorRoles = ['primary', 'subtle', 'brand', 'neutralMuted', 'brandAlt', 'red', 'orange', 'yellow', 'green', 'teal', 'blue', 'purple', 'pink', 'brown', 'neutralDark', 'neutralLight', 'neutralLighter', 'brandDarkest', 'brandDarker', 'brandDark', 'brandLight', 'darkRed', 'amber', 'yellowGreen', 'cyan', 'navy'];
test('every semantic foreground role round-trips canonical HTML without storing theme values', t => {
  const { api, doc } = fixture(t);
  for (const color of semanticColorRoles) {
    const html = `<h2><span class="wps-color-${color}">Title</span></h2>`;
    const result = api.parseTitleHTML(html, doc);
    assert.equal(result.kind, 'valid', color);
    assert.equal(result.document.content[0].marks.color, color);
    assert.equal(api.serializeTitleDocument(result.document), html);
    assert.deepEqual(api.parseTitleHTML(api.serializeTitleDocument(result.document), doc).document, result.document);
  }
});
test('semantic colors reject arbitrary CSS, token names and prototype keys, while Automatic omits the mark', t => {
  const { api, doc } = fixture(t);
  for (const color of ['custom', '#ff0000', 'rgb(1, 2, 3)', 'colorNeutralForeground1', 'constructor', '__proto__', 'RED']) {
    const invalid = { type: 'paragraph', content: [{ type: 'text', text: 'Title', marks: { color } }] };
    assert.equal(api.isTitleDocument(invalid), false, color);
    assert.throws(() => api.serializeTitleDocument(invalid), /Invalid/, color);
    const result = api.parseTitleHTML(`<p><span class="wps-color-${color}">Title</span></p>`, doc);
    assert.equal(result.document.content[0].marks?.color, undefined);
  }
  assert.equal(api.serializeTitleDocument({ type: 'paragraph', content: [{ type: 'text', text: 'Title' }] }), '<p>Title</p>');
});
