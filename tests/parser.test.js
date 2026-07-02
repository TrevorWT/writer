// Run: node tests/parser.test.js
import assert from 'node:assert/strict';
import { parse, serialize, serializeChapter, serializeSkeleton, treeHeight, chapterFileName, cleanText, wordCount, splitFrontmatter, buildFrontmatter } from '../app/parser.js';

// frontmatter: scalars, lists, CRLF, round-trip, absent
{
  const fm = splitFrontmatter('---\nauthor: Trevor T\ncontact:\n  - 123 Street\n  - t@pdx.edu\n---\n\nlogline\n\n# A\n');
  assert.equal(fm.meta.author, 'Trevor T');
  assert.deepEqual(fm.meta.contact, ['123 Street', 't@pdx.edu']);
  assert.ok(fm.body.startsWith('logline'));
  const rebuilt = buildFrontmatter(fm.meta);
  assert.deepEqual(splitFrontmatter(rebuilt + fm.body).meta, fm.meta);
  assert.deepEqual(splitFrontmatter('no frontmatter here').meta, {});
  assert.equal(buildFrontmatter({}), '');
  assert.equal(buildFrontmatter({ author: '  ' }), '');
  const crlf = splitFrontmatter('---\r\nauthor: X\r\n---\r\nbody');
  assert.equal(crlf.meta.author, 'X');
}

// basic structure + root body
const t = parse('logline\n\n# A\n\nbody [[Bob]]\n\n## C1\n\n# B\n', 'My Story');
assert.equal(t.title, 'My Story');
assert.equal(t.body, 'logline');
assert.equal(t.children.length, 2);
assert.equal(t.children[0].children[0].title, 'C1');
assert.equal(parse(serialize(t), 'My Story').children[1].title, 'B');

// CRLF input must parse identically (regression: titled headings failed on \r)
const crlf = parse('logline\r\n\r\n# A\r\n\r\nbody\r\n\r\n## C1\r\n\r\n# B\r\n', 'S');
assert.equal(crlf.children[0].title, 'A');
assert.equal(crlf.children[0].children[0].title, 'C1');
assert.equal(crlf.children[1].title, 'B');

// bare headings (untitled sections)
const u = parse('# P\n\n## C\n\n###\n\npage one\n\n###\n\npage two\n\n# P2\n', 'S');
assert.equal(u.children[0].children[0].children.length, 2);
assert.equal(u.children[0].children[0].children[0].title, '');

// section notes extraction + round-trip (two H1s so no unwrap)
const nt = parse('# A\n\nprose here\n\n%%\nremember the key\n%%\n\n# B\n', 'S');
assert.equal(nt.children[0].body, 'prose here');
assert.equal(nt.children[0].notes, 'remember the key');
assert.ok(serialize(nt).includes('%%\nremember the key\n%%'));

// inline annotations survive in the body, distinct from section notes
const an = parse('# A\n\nsee ==the key==%%foreshadow%% here\n\n%%\nsection note\n%%\n\n# B\n', 'S');
assert.equal(an.children[0].notes, 'section note');
assert.equal(an.children[0].body, 'see ==the key==%%foreshadow%% here');

// old-format single-H1 unwrap
const old = parse('# Old Title\n\nintro\n\n## Beginning\n\nx\n\n### Ch1\n', 'Folder');
assert.equal(old.children[0].depth, 1);
assert.equal(old.children[0].title, 'Beginning');

// skeleton + chapter round-trip at the seam
const story = parse('log\n\n# P1\n\nnote\n\n## Ch\n\nchbody\n\n### S1\n\nscene text\n\n# P2\n', 'S');
const H = treeHeight(story);
assert.equal(H, 3);
const skel = serializeSkeleton(story, H);
assert.ok(skel.includes('![[chapters/01 Ch]]'));
assert.ok(!skel.includes('chbody'), 'chapter body stays out of the skeleton');
const ch = story.children[0].children[0];
const chText = serializeChapter(ch);
assert.ok(chText.startsWith('chbody'));
assert.ok(chText.includes('# S1'), 'scenes serialize from # down in chapter files');
// chapter files parse with unwrap disabled — a one-scene chapter must NOT
// have its lone # scene unwrapped away (regression: scene text was dropped)
const reparsed = parse(chText, '', { unwrap: false });
assert.equal(reparsed.body, 'chbody');
assert.equal(reparsed.children[0].title, 'S1');
assert.equal(reparsed.children[0].body, 'scene text');

// filenames strip illegal characters
assert.equal(chapterFileName({ title: 'Who? Stole: the *Tarts*' }, 3), '03 Who Stole the Tarts.md');

// manuscript cleaning
assert.equal(cleanText('see ==the key==%%note%% and [[Alice]] %%gone%%'), 'see the key and Alice');

// word counts ignore notes
assert.equal(wordCount({ body: 'one two %%not counted%% three', children: [] }), 3);

console.log('parser tests: all passed');
