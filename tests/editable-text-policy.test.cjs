const test = require('node:test');
const assert = require('node:assert/strict');
const { createCoreHarness } = require('./fixtures/editable-text/core-harness.cjs');
function fixture(t) {
  const h = createCoreHarness();
  t.after(h.close);
  return h.load('editable-text-policy.internal.ts');
}

test('omission admits every semantic block and initializes paragraphs', t => {
  const api = fixture(t);
  const policy = api.resolveEditableTextBlockPolicy(undefined, undefined);
  assert.deepEqual(policy.blocks, ['h1', 'h2', 'h3', 'h4', 'p']);
  assert.equal(policy.defaultBlock, 'p');
});
test('membership removes duplicates, uses canonical order and chooses the lowest heading', t => {
  const api = fixture(t);
  const first = api.resolveEditableTextBlockPolicy(['h4', 'h2', 'h2'], undefined);
  const second = api.resolveEditableTextBlockPolicy(['h2', 'h4'], 'h4');
  assert.deepEqual(first.blocks, ['h2', 'h4']);
  assert.equal(first.defaultBlock, 'h2');
  assert.equal(second.defaultBlock, 'h4');
  assert.equal(first.membershipKey, second.membershipKey);
  assert.notEqual(first.membershipKey, api.resolveEditableTextBlockPolicy(['h2'], undefined).membershipKey);
});
test('invalid runtime block configurations fail instead of silently expanding the policy', t => {
  const api = fixture(t);
  assert.equal(typeof api.resolveEditableTextBlockPolicy, 'function');
  for (const blocks of [[], ['h5'], ['p', 'H2'], 'h2', {}, null, Array(1)]) {
    assert.throws(() => api.resolveEditableTextBlockPolicy(blocks, undefined), { name: 'TypeError' });
  }
  for (const block of ['p', 'h5', null]) {
    assert.throws(() => api.resolveEditableTextBlockPolicy(['h2', 'h4'], block), { name: 'TypeError' });
  }
});
test('normalizing an excluded H2 preserves every content and format field without mutating input', t => {
  const api = fixture(t);
  const content = Object.freeze([Object.freeze({ type: 'text', text: 'A', marks: Object.freeze({ bold: true, size: 24, color: 'brand' }) }), Object.freeze({ type: 'hardBreak' })]);
  const document = Object.freeze({ type: 'heading', level: 2, content, lineSpacing: 1.4, spaceBefore: true, spaceAfter: true });
  const result = api.normalizeEditableTextDocument(document, api.resolveEditableTextBlockPolicy(['p'], undefined));
  assert.deepEqual(result, { type: 'paragraph', content, lineSpacing: 1.4, spaceBefore: true, spaceAfter: true });
  assert.equal(result.content, content);
  assert.equal(document.type, 'heading');
  assert.equal(document.level, 2);
});
test('explicit default converts paragraphs while compliant nonempty blocks retain identity and format', t => {
  const api = fixture(t);
  const policy = api.resolveEditableTextBlockPolicy(['h2', 'h3'], 'h3');
  const paragraph = { type: 'paragraph', content: [{ type: 'text', text: 'A' }] };
  assert.deepEqual(api.normalizeEditableTextDocument(paragraph, policy), { type: 'heading', level: 3, content: paragraph.content });
  const heading = { type: 'heading', level: 2, content: [{ type: 'text', text: 'B' }], lineSpacing: 2 };
  assert.equal(api.normalizeEditableTextDocument(heading, policy), heading);
});
test('empty and whitespace-only documents adopt the effective default and preserve inline content', t => {
  const api = fixture(t);
  const policy = api.resolveEditableTextBlockPolicy(undefined, 'h3');
  for (const content of [[], [{ type: 'text', text: ' \n\u00a0' }, { type: 'hardBreak' }]]) {
    const original = { type: 'heading', level: 2, content, spaceAfter: true };
    assert.deepEqual(api.normalizeEditableTextDocument(original, policy), { type: 'heading', level: 3, content, spaceAfter: true });
  }
});
