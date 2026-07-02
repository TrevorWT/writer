// Pure markdown <-> tree logic. No DOM, no filesystem — testable in Node.
//
// Tree node: { depth, title, body, notes, children }
// %%…%% blocks are per-section notes (Obsidian comment syntax), kept out of
// the prose body. Inline annotations (==text==%%note%%) stay in the body.

// minimal YAML frontmatter: `key: value` lines and `key:` + `- item` lists.
// Story metadata (author, byline, contact) lives here, Obsidian-compatible.
export function splitFrontmatter(md) {
  const meta = {};
  const m = md.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n?/);
  if (!m) return { meta, body: md };
  let curKey = null;
  for (const line of m[1].split(/\r?\n/)) {
    const li = line.match(/^\s*- (.*)$/);
    if (li && curKey) {
      if (!Array.isArray(meta[curKey])) meta[curKey] = [];
      meta[curKey].push(li[1]);
      continue;
    }
    const kv = line.match(/^([A-Za-z_][\w-]*):\s*(.*)$/);
    if (kv) { curKey = kv[1]; if (kv[2]) meta[curKey] = kv[2]; }
  }
  return { meta, body: md.slice(m[0].length).replace(/^(\r?\n)+/, '') };
}
export function buildFrontmatter(meta) {
  const keys = Object.keys(meta).filter(k =>
    Array.isArray(meta[k]) ? meta[k].length : String(meta[k] || '').trim());
  if (!keys.length) return '';
  let out = '---\n';
  for (const k of keys) {
    if (Array.isArray(meta[k])) { out += k + ':\n'; for (const v of meta[k]) out += '  - ' + v + '\n'; }
    else out += k + ': ' + meta[k] + '\n';
  }
  return out + '---\n\n';
}

export function parse(md, rootTitle, { unwrap = true } = {}) {
  const root = { depth: 0, title: rootTitle, body: '', notes: '', children: [] };
  const stack = [root];
  let cur = root;
  for (const line of md.split(/\r?\n/)) {
    const m = line.match(/^(#{1,6})(?:\s+(.*))?$/);
    if (m) {
      const node = { depth: m[1].length, title: (m[2] || '').trim(), body: '', notes: '', children: [] };
      while (stack[stack.length - 1].depth >= node.depth) stack.pop();
      stack[stack.length - 1].children.push(node);
      stack.push(node);
      cur = node;
    } else {
      cur.body += (cur.body ? '\n' : '') + line;
    }
  }
  (function trim(n) {
    n.notes = '';
    n.body = n.body
      .replace(/==[^=\n]+==%%[\s\S]*?%%/g, m => m.replace(/%%/g, '\x00'))
      .replace(/%%([\s\S]*?)%%/g, (m, t) => { n.notes += (n.notes ? '\n' : '') + t.trim(); return ''; })
      .replace(/\x00/g, '%%')
      .trim();
    n.children.forEach(trim);
  })(root);
  // single H1 = old-format title wrapper, unwrap it. NOT for chapter files —
  // a one-scene chapter is a legitimate single-# document (caught by tests).
  if (unwrap && root.children.length === 1 && root.children[0].depth === 1) {
    const h1 = root.children[0];
    root.body = root.body ? (h1.body ? root.body + '\n\n' + h1.body : root.body) : h1.body;
    if (!root.notes) root.notes = h1.notes;
    root.children = h1.children;
    root.children.forEach(function promote(n) { n.depth--; n.children.forEach(promote); });
  }
  return root;
}

export function serialize(node) {
  let out = '';
  if (node.depth > 0) out += ('#'.repeat(node.depth) + ' ' + node.title).trimEnd() + '\n\n';
  if (node.body) out += node.body + '\n\n';
  if (node.notes) out += '%%\n' + node.notes + '\n%%\n\n';
  for (const c of node.children) out += serialize(c);
  return out;
}

// depth-agnostic serialization: headings start at the given depth
export function serializeAt(node, depth) {
  let out = ('#'.repeat(depth) + ' ' + node.title).trimEnd() + '\n\n';
  if (node.body) out += node.body + '\n\n';
  if (node.notes) out += '%%\n' + node.notes + '\n%%\n\n';
  for (const c of node.children) out += serializeAt(c, depth + 1);
  return out;
}

export function treeHeight(n) {
  return n.children.length ? 1 + Math.max(...n.children.map(treeHeight)) : 0;
}

export function chapterFileName(node, num) {
  const safe = node.title.replace(/[\\/:*?"<>|[\]]/g, '').trim();
  return String(num).padStart(2, '0') + (safe ? ' ' + safe : '') + '.md';
}

// chapter files: body + notes, then scenes from # down (re-depthed on load)
export function serializeChapter(ch) {
  let out = '';
  if (ch.body) out += ch.body + '\n\n';
  if (ch.notes) out += '%%\n' + ch.notes + '\n%%\n\n';
  for (const sc of ch.children) out += serializeAt(sc, 1);
  return out;
}

// story.md skeleton: everything above the chapter seam (depth H-1), with
// ![[chapters/NN Title]] embeds where the chapters go
export function serializeSkeleton(tree, H) {
  let out = '';
  if (tree.body) out += tree.body + '\n\n';
  if (tree.notes) out += '%%\n' + tree.notes + '\n%%\n\n';
  let num = 0;
  (function walk(n) {
    for (const c of n.children) {
      if (H >= 2 && c.depth === H - 1) {
        out += `![[chapters/${chapterFileName(c, ++num).slice(0, -3)}]]\n`;
        continue;
      }
      out += ('#'.repeat(c.depth) + ' ' + c.title).trimEnd() + '\n\n';
      if (c.body) out += c.body + '\n\n';
      if (c.notes) out += '%%\n' + c.notes + '\n%%\n\n';
      walk(c);
      if (H >= 2 && c.depth === H - 2 && c.children.length) out += '\n';
    }
  })(tree);
  return out;
}

// manuscript compile: strip notes, annotations' notes, and [[brackets]]
export function cleanText(s) {
  return s
    .replace(/==([^=\n]+)==%%[\s\S]*?%%/g, '$1')
    .replace(/%%[\s\S]*?%%/g, '')
    .replace(/\[\[([^\]]+)\]\]/g, '$1')
    .trim();
}

export function wordCount(n) {
  let w = n.body ? n.body.replace(/%%[\s\S]*?%%/g, '').split(/\s+/).filter(Boolean).length : 0;
  for (const c of n.children) w += wordCount(c);
  return w;
}
