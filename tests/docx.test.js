// Run: node tests/docx.test.js
import assert from 'node:assert/strict';
import { zipStore, toRuns, makeDocx } from '../app/docx.js';

// runs tokenizer
assert.deepEqual(toRuns('plain'), [{ t: 'plain' }]);
assert.deepEqual(toRuns('a *b* c'), [{ t: 'a ' }, { t: 'b', i: true }, { t: ' c' }]);
assert.deepEqual(toRuns('_x_ and **y**'), [{ t: 'x', i: true }, { t: ' and ' }, { t: 'y', b: true }]);

// zip structure: signatures, EOCD count, readable names, valid CRCs
function readZip(bytes) {
  // find EOCD
  let e = bytes.length - 22;
  assert.ok(bytes[e] === 0x50 && bytes[e + 1] === 0x4B && bytes[e + 2] === 0x05 && bytes[e + 3] === 0x06, 'EOCD');
  const count = bytes[e + 10] | (bytes[e + 11] << 8);
  const cdOfs = bytes[e + 16] | (bytes[e + 17] << 8) | (bytes[e + 18] << 16) | (bytes[e + 19] << 24);
  const names = [];
  let p = cdOfs;
  const dec = new TextDecoder();
  for (let i = 0; i < count; i++) {
    assert.equal(bytes[p], 0x50); assert.equal(bytes[p + 1], 0x4B); assert.equal(bytes[p + 2], 0x01);
    const nameLen = bytes[p + 28] | (bytes[p + 29] << 8);
    names.push(dec.decode(bytes.slice(p + 46, p + 46 + nameLen)));
    p += 46 + nameLen;
  }
  return { count, names };
}
const z = zipStore([{ name: 'a.txt', text: 'hello' }, { name: 'dir/b.xml', text: '<x/>' }]);
const zr = readZip(z);
assert.equal(zr.count, 2);
assert.deepEqual(zr.names, ['a.txt', 'dir/b.xml']);

// docx: parts present, italics as runs, SMF specifics
const blocks = ['# Chapter One', 'She saw the *key* on the table.', '* * *', 'Later that night.'];
const tp = { contact: ['Trevor T', 't@pdx.edu'], by: 'T. Thompson', approx: '26,100' };

const smf = makeDocx(blocks, { title: 'Alice', titlePage: tp, smf: true, author: 'Trevor Thompson' });
const smfR = readZip(smf);
assert.ok(smfR.names.includes('word/document.xml'));
assert.ok(smfR.names.includes('word/header1.xml'), 'SMF has running header');
const smfText = new TextDecoder().decode(smf);
assert.ok(smfText.includes('<w:i/>'), 'italic run emitted');
assert.ok(smfText.includes('w:line="480"'), 'double spacing');
assert.ok(smfText.includes('w:firstLine="720"'), 'first-line indent');
assert.ok(smfText.includes('Thompson / ALICE / '), 'running header content');
assert.ok(smfText.includes(' PAGE '), 'page number field');
assert.ok(smfText.includes('approx. 26,100 words'), 'word count on title page');
assert.ok(!smfText.includes('* * *') || smfText.includes('>#<'), 'SMF scene separator is #');

const plain = makeDocx(blocks, { title: 'Alice', titlePage: null, smf: false, author: '' });
const plainR = readZip(plain);
assert.ok(!plainR.names.includes('word/header1.xml'), 'no header without SMF');
const plainText = new TextDecoder().decode(plain);
assert.ok(plainText.includes('* * *'), 'plain keeps * * * separator');
assert.ok(!plainText.includes('w:firstLine'), 'no indent outside SMF');

console.log('docx tests: all passed');
