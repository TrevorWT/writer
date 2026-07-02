import { parse, serialize, treeHeight, chapterFileName, serializeChapter, serializeSkeleton, cleanText, wordCount, splitFrontmatter, buildFrontmatter } from './parser.js';
import { makeDocx } from './docx.js';
import { initFS } from './fs.js';
const FS = initFS();

import { icon } from './icons.js';

// static chrome icons
document.getElementById('sidetoggle').innerHTML = icon('sidebar', 18);
document.getElementById('settingsbtn').innerHTML = icon('gear', 17);
document.getElementById('exportbtn').innerHTML = icon('upload', 17);
document.getElementById('tsicon').innerHTML = icon('search', 13);
document.getElementById('openbtn').innerHTML = icon('folder', 14);
document.getElementById('newstory').innerHTML = icon('plus', 14);
document.getElementById('newtag').innerHTML = icon('plus', 14);
document.getElementById('undobtn').innerHTML = icon('undo', 14);
document.getElementById('redobtn').innerHTML = icon('redo', 14);
document.getElementById('modebtn').innerHTML = icon('pencil', 13) + ' editing';
document.getElementById('storyinfobtn').innerHTML = icon('pencil', 12);

const KIND = ['Story', 'Part', 'Chapter', 'Scene', 'Page', 'Section', 'Section'];
let libPath = null, storyPath = null, storyName = '', storyTree = null;
let storyMeta = {};   // frontmatter of story.md: author, byline, contact[]
let currentTagPath = null;
let unwatchStory = null, lastSaveAt = 0, watchTimer = null;
const openCats = new Set();   // categories are collapsed unless opened
let currentTag = null;     // {cat, name} when viewing a tag page, else null
let tree = null, path = [];
let readonly = false;
let tagIndex = new Map();  // name -> category
let allCats = [];          // categories present in this story

const $ = id => document.getElementById(id);
const kindName = depth => currentTag ? (depth === 0 ? currentTag.cat : 'Section') : (KIND[depth] || 'Section');

// generated numbering: layer names hang from the BOTTOM of the tree — leaves
// are Scenes, scene-holders are Chapters, and so on up. Nesting deeper
// anywhere relabels the whole story; no configuration.
let LADDER = JSON.parse(localStorage.getItem('writer-ladder') || 'null') || ['Scene', 'Chapter', 'Part', 'Book', 'Volume'];
let kindMap = new Map();
function computeKinds() {
  kindMap = new Map();
  if (!storyTree) return;
  const H = treeHeight(storyTree);
  const counts = {};
  (function walk(n) {
    for (const c of n.children) {
      const kind = LADDER[Math.min(Math.max(H - c.depth, 0), LADDER.length - 1)];
      counts[kind] = (counts[kind] || 0) + 1;
      kindMap.set(c, { kind, num: counts[kind], label: kind + ' ' + counts[kind] });
      walk(c);
    }
  })(storyTree);
}
const labelOf = n => n.title || (kindMap.get(n) || {}).label || '(untitled)';

// ---- hybrid storage: story.md is the skeleton, one file per chapter ----
let chapterCache = new Map();   // filename -> last written text (skip clean writes)
let storyBackup = null;         // raw story.md as loaded; trashed once on first save

async function trashFile(name, text) {
  const tp = FS.join(storyPath, '.trash');
  await FS.mkdir(tp);
  await FS.writeText(FS.join(tp, `${Date.now()} ${name}`), text);
}
async function save() {
  lastSaveAt = Date.now();
  if (currentTag) { await FS.writeText(currentTagPath, serialize(tree)); lastSaveAt = Date.now(); return; }
  if (storyBackup !== null) {
    if (storyBackup.trim()) await trashFile('story.md', storyBackup);
    storyBackup = null;
  }
  const H = treeHeight(tree);
  const chDir = FS.join(storyPath, 'chapters');
  await FS.mkdir(chDir);
  const used = new Set();
  let num = 0;
  const chapters = [];
  if (H >= 2) (function collect(n) {
    for (const c of n.children) {
      if (c.depth === H - 1) chapters.push(c); else collect(c);
    }
  })(tree);
  for (const ch of chapters) {
    const fname = chapterFileName(ch, ++num);
    used.add(fname);
    const text = serializeChapter(ch);
    if (chapterCache.get(fname) !== text) {
      await FS.writeText(FS.join(chDir, fname), text);
      chapterCache.set(fname, text);
    }
  }
  await FS.writeText(FS.join(storyPath, 'story.md'), buildFrontmatter(storyMeta) + serializeSkeleton(tree, H));
  // sweep chapter files that no longer correspond to a chapter — NEVER hard
  // delete: anything with content goes to .trash/ first
  for (const f of await FS.readDir(chDir)) {
    if (f.isDir || !f.name.endsWith('.md') || used.has(f.name)) continue;
    try {
      const text = await FS.readText(FS.join(chDir, f.name));
      if (text.trim()) await trashFile(f.name, text);
    } catch {}
    await FS.remove(FS.join(chDir, f.name));
    chapterCache.delete(f.name);
  }
  lastSaveAt = Date.now();
}
let saveTimer = null;
function queueSave() { clearTimeout(saveTimer); saveTimer = setTimeout(save, 600); }

// ---- rendering ----
function esc(s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
const escRe = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
function anchor(name) { return `<a class="entity" data-name="${name}">${name}</a>`; }
// edit=true keeps unknown [[names]] as literal text so the editable div's
// innerText round-trips to valid markdown (known tags need no brackets at all)
function linkify(s, edit) {
  let html = esc(s)
    .replace(/==([^=\n]+)==%%([\s\S]*?)%%/g, (m, t, note) =>
      `<mark class="annot" data-note="${note.replace(/"/g, '&quot;')}">${t}</mark>`)
    .replace(/\[\[([^\]]+)\]\]/g, (m, n) => (edit && !tagIndex.has(n)) ? m : anchor(n));
  const names = [...tagIndex.keys()].sort((a, b) => b.length - a.length);
  if (names.length) {
    const re = new RegExp('\\b(' + names.map(escRe).join('|') + ')\\b', 'g');
    // only auto-link outside existing anchors and annotations
    html = html.split(/(<a[^>]*>[^<]*<\/a>|<mark[^>]*>[^<]*<\/mark>)/g)
      .map((seg, i) => i % 2 ? seg : seg.replace(re, m => anchor(m)))
      .join('');
  }
  return html;
}
// editable-DOM -> markdown source (anchors flow as plain text, annotations
// reconstruct their ==text==%%note%% form, divs/brs become newlines)
function domToMd(root) {
  let out = '';
  (function walk(n) {
    for (const c of n.childNodes) {
      if (c.nodeType === 3) out += c.textContent;
      else if (c.tagName === 'BR') { if (c.parentNode.childNodes.length > 1) out += '\n'; }
      else if (c.classList && c.classList.contains('annot')) out += `==${c.textContent}==%%${c.dataset.note}%%`;
      else if (c.tagName === 'DIV') { out += '\n'; walk(c); }
      else walk(c);
    }
  })(root);
  return out.replace(/^\n+|\n+$/g, '');
}
const focus = () => path[path.length - 1];
const focusHasChildren = () => tree && focus().children.length > 0;
let panelsHidden = false;

function render() {
  const has = !!tree;
  $('hint').hidden = has;
  $('focus').hidden = !has;
  $('modebtn').hidden = !has;
  $('topsearch').hidden = !has;
  $('upbtn').hidden = !has;
  $('upbtn').style.visibility = has && (path.length > 1 || currentTag) ? 'visible' : 'hidden';
  $('histbtns').hidden = !has || !!currentTag;
  $('panelsbtn').hidden = !has || !focusHasChildren();
  $('crumbpath').innerHTML = '';
  $('panels').innerHTML = '';
  if (!has) { $('timeline').hidden = true; return; }
  computeKinds();

  // breadcrumbs (tag pages get the story as their first crumb)
  const crumbs = $('crumbpath');
  if (currentTag) {
    const c = document.createElement('span');
    c.textContent = storyName;
    c.onclick = () => openStory(storyName);
    crumbs.appendChild(c);
    const s = document.createElement('span'); s.className = 'sep'; s.textContent = '›';
    crumbs.appendChild(s);
  }
  path.forEach((n, i) => {
    if (i) { const s = document.createElement('span'); s.className = 'sep'; s.textContent = '›'; crumbs.appendChild(s); }
    const c = document.createElement('span');
    c.textContent = labelOf(n);
    c.onclick = () => { path = path.slice(0, i + 1); render(); };
    crumbs.appendChild(c);
  });

  const f = focus();
  const wc = wordCount(f);
  const kindLabel = currentTag ? kindName(f.depth)
    : f.depth === 0 ? 'Story' : ((kindMap.get(f) || {}).label || '');
  $('focuskind').textContent = kindLabel + (wc && !currentTag ? ` · ${wc.toLocaleString()} words` : '');
  $('main').classList.toggle('writing', !f.children.length && (f.depth > 0 || !!currentTag));
  $('main').classList.toggle('hidepanels', panelsHidden && f.children.length > 0);
  $('panelsbtn').innerHTML = (panelsHidden ? icon('square', 13) : icon('columns', 13)) + ' panels';
  $('gridbtn').hidden = !has || !focusHasChildren() || panelsHidden;
  $('gridbtn').innerHTML = layout.grid ? icon('rows', 13) + ' row' : icon('grid', 13) + ' grid';
  $('main').classList.toggle('grid', !!layout.grid);
  const ft = $('focustitle');
  const titleEditable = f.depth > 0 && !readonly;
  ft.contentEditable = titleEditable ? 'plaintext-only' : 'false';
  ft.textContent = f.title;
  ft.style.cursor = titleEditable ? 'text' : 'default';
  ft.onblur = titleEditable ? () => { f.title = ft.innerText.trim() || f.title; save(); render(); } : null;
  ft.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); ft.blur(); } };

  const fb = $('focusbody');
  fb.onclick = e => {
    if (e.target.classList.contains('entity')) openTagByName(e.target.dataset.name);
    else if (e.target.classList.contains('annot')) openAnnot(e.target);
  };
  if (readonly) {
    fb.contentEditable = 'false';
    fb.innerHTML = linkify(f.body);
    fb.oninput = fb.onblur = fb.onpaste = null;
  } else {
    fb.contentEditable = 'true';
    fb.innerHTML = linkify(f.body, true);
    fb.oninput = () => { f.body = domToMd(fb); queueSave(); };
    fb.onblur = () => { clearTimeout(saveTimer); save(); };
    fb.onpaste = e => { e.preventDefault(); document.execCommand('insertText', false, e.clipboardData.getData('text/plain')); };
  }

  const wrap = $('panels');
  wrap.appendChild(gapEl(f, 0));
  f.children.forEach((c, i) => {
    wrap.appendChild(panelEl(c));
    wrap.appendChild(gapEl(f, i + 1));
  });

  renderTimeline();
  renderOutline();
  renderBacklinks();
  renderNotes();
}

// ---- per-section notes popout ----
function renderNotes() {
  const f = tree && focus();
  $('notestab').hidden = !f;
  if (!f) { $('notespanel').classList.remove('open'); return; }
  $('noteslabel').textContent = labelOf(f);
  $('notestext').textContent = f.notes || '';

  // highlight notes in this section (Google Docs-style comment list)
  const list = $('annotlist');
  list.innerHTML = '';
  const annots = [...f.body.matchAll(/==([^=\n]+)==%%([\s\S]*?)%%/g)];
  $('notestab').classList.toggle('has', !!f.notes || !!annots.length);
  if (!annots.length) return;
  const h = document.createElement('div');
  h.className = 'alhead';
  h.textContent = `Highlight notes (${annots.length})`;
  list.appendChild(h);
  annots.forEach((m, i) => {
    const row = document.createElement('div');
    row.className = 'alrow';
    row.innerHTML = `<div class="altext">“${esc(m[1])}”</div>` + (m[2].trim() ? `<div class="alnote">${esc(m[2])}</div>` : '');
    row.onclick = () => {
      const mk = document.querySelectorAll('#focusbody mark.annot')[i];
      if (!mk) return;
      mk.scrollIntoView({ block: 'center', behavior: 'smooth' });
      mk.classList.add('flash');
      setTimeout(() => { mk.classList.remove('flash'); openAnnot(mk); }, 400);
    };
    list.appendChild(row);
  });
}
$('notestab').onclick = () => { $('notespanel').classList.add('open'); $('notestext').focus(); };
document.querySelector('#notespanel .nclose').onclick = () => $('notespanel').classList.remove('open');
$('notestext').oninput = () => { focus().notes = $('notestext').innerText.trim(); queueSave(); };
$('notestext').onblur = () => { clearTimeout(saveTimer); save(); $('notestab').classList.toggle('has', !!focus().notes); };

// ---- backlinks on tag pages ----
function renderBacklinks() {
  const bl = $('backlinks');
  bl.innerHTML = '';
  bl.hidden = true;
  if (!currentTag || currentTag.template || !storyTree) return;
  const re = new RegExp('\\b' + escRe(currentTag.name) + '\\b');
  const hits = [];
  (function walk(n, anc) {
    n.children.forEach(c => {
      const p = [...anc, c];
      if (re.test(c.body) || re.test(c.title)) hits.push({ node: c, path: p });
      walk(c, p);
    });
  })(storyTree, [storyTree]);
  if (!hits.length) return;

  bl.hidden = false;
  const det = document.createElement('details');
  const sum = document.createElement('summary');
  sum.textContent = `Appears in ${hits.length} place${hits.length > 1 ? 's' : ''} in ${storyName}`;
  det.appendChild(sum);
  for (const h of hits) {
    const row = document.createElement('div');
    row.className = 'blrow';
    const where = document.createElement('div');
    where.className = 'where';
    where.textContent = h.path.slice(1).map(labelOf).join(' › ');
    row.appendChild(where);
    const i = h.node.body.search(re);
    if (i >= 0) {
      const snip = document.createElement('div');
      snip.className = 'snippet';
      const start = Math.max(0, i - 60);
      const raw = (start > 0 ? '…' : '') + h.node.body.slice(start, i + currentTag.name.length + 60).replace(/\n/g, ' ') + '…';
      snip.innerHTML = esc(raw).replace(re, m => '<b>' + m + '</b>');
      row.appendChild(snip);
    }
    row.onclick = () => { currentTag = null; currentTagPath = null; tree = storyTree; path = h.path; render(); };
    det.appendChild(row);
  }
  bl.appendChild(det);
}

// ---- sidebar outline tree ----
$('outlinetoggle').onclick = () => {
  if (!storyTree) return;
  let anyOpen = false;
  (function scan(n) { n.children.forEach(c => { if (c._open && c.children.length) anyOpen = true; scan(c); }); })(storyTree);
  (function set(n) { n.children.forEach(c => { c._open = !anyOpen; set(c); }); })(storyTree);
  renderOutline();
};

let lastOutlineFocus = null;
function renderOutline() {
  const list = $('outlinelist');
  list.innerHTML = '';
  if (!storyTree) return;
  // reveal where you are — but only on navigation, so manual expand/contract
  // clicks aren't immediately overridden
  if (!currentTag && focus() !== lastOutlineFocus) {
    path.slice(1, -1).forEach(n => n._open = true);
    lastOutlineFocus = focus();
  }
  (function walk(n, anc, indent) {
    for (const c of n.children) {
      const p = [...anc, c];
      const row = document.createElement('div');
      row.className = 'orow' + (!currentTag && focus() === c ? ' active' : '');
      row.style.paddingLeft = (8 + indent * 14) + 'px';
      const arrow = document.createElement('span');
      arrow.className = 'arrow';
      arrow.textContent = c.children.length ? (c._open ? '▾' : '▸') : '';
      arrow.onclick = e => { e.stopPropagation(); c._open = !c._open; renderOutline(); };
      const label = document.createElement('span');
      label.textContent = labelOf(c);
      label.style.cssText = 'flex:1;overflow:hidden;text-overflow:ellipsis';
      const del = document.createElement('span');
      del.className = 'rowdel';
      del.textContent = '✕';
      del.title = 'Delete section';
      del.onclick = e => {
        e.stopPropagation();
        if (currentTag) { currentTag = null; currentTagPath = null; tree = storyTree; path = [storyTree]; }
        deleteNode(c);
      };
      row.append(arrow, label, del);
      row.onclick = () => {
        currentTag = null; currentTagPath = null; tree = storyTree;
        path = [storyTree, ...p];
        render();
      };
      makeDraggable(row, c);
      makeDropTarget(row, () => [c, c.children.length]);    // drop on an outline row = into it
      list.appendChild(row);
      if (c._open) walk(c, p, indent + 1);
    }
  })(storyTree, [], 0);
  let anyOpen = false;
  (function scan(n) { n.children.forEach(c => { if (c._open && c.children.length) anyOpen = true; scan(c); }); })(storyTree);
  $('outlinetoggle').innerHTML = anyOpen ? icon('collapse', 13) : icon('expand', 13);
}

function panelEl(node) {
  const el = document.createElement('div');
  el.className = 'panel';
  const info = kindMap.get(node);
  if (info) {
    const k = document.createElement('div');
    k.className = 'kind';
    k.textContent = info.label;
    el.appendChild(k);
  }
  const del = document.createElement('span');
  del.className = 'pdel';
  del.textContent = '✕';
  del.title = 'Delete';
  del.onclick = e => { e.stopPropagation(); deleteNode(node); };
  el.appendChild(del);
  const h = document.createElement('h4');
  h.textContent = node.title;
  h.ondblclick = e => { e.stopPropagation(); editTitle(node, h); };
  const p = document.createElement('div');
  p.className = 'preview';
  p.innerHTML = linkify(node.body || node.children.map(c => labelOf(c)).join(' · '));
  const fade = document.createElement('div');
  fade.className = 'fade';
  el.append(h, p, fade);
  const wc = wordCount(node);
  if (node.children.length || wc) {
    const c = document.createElement('div');
    c.className = 'count';
    let t = '';
    if (node.children.length) {
      const kinds = node.children.map(k => ((kindMap.get(k) || {}).kind || 'section').toLowerCase());
      const word = kinds.every(k => k === kinds[0]) ? kinds[0] : 'section';
      t = `${node.children.length} ${word}${node.children.length > 1 ? 's' : ''}`;
    }
    if (wc) t += (t ? ' · ' : '') + wc.toLocaleString() + ' words';
    c.textContent = t;
    el.appendChild(c);
  }
  el.onclick = e => {
    if (e.target.classList.contains('entity')) { openTagByName(e.target.dataset.name); return; }
    if (e.target.isContentEditable) return;
    path.push(node); render();
  };
  el._node = node;
  makeDraggable(el, node);
  // drop ON a panel = insert before it (reorder); Alt+drop = nest inside it.
  // Plain nesting-by-drop is too easy to do by accident — it reshapes the story.
  el.title = 'Drag to reorder · hold Alt while dropping to nest inside';
  makeDropTarget(el, e => e && e.altKey
    ? [node, node.children.length]
    : (() => { const p = focus(); return [p, p.children.indexOf(node)]; })());
  return el;
}

function gapEl(parent, index) {
  const g = document.createElement('div');
  g.className = 'gap';
  const plus = document.createElement('div');
  plus.className = 'plus';
  plus.textContent = '+';
  plus.title = 'Add section';
  plus.onclick = () => addChild(parent, index);
  g.appendChild(plus);
  makeDropTarget(g, () => [parent, index]);                 // drop BETWEEN panels = reorder
  return g;
}

function defaultTitle() {
  return currentTag ? 'New section' : '';   // story sections: the generated label names them
}
function addChild(parent, index) {
  pushUndo();
  const node = { depth: parent.depth + 1, title: defaultTitle(), body: '', notes: '', children: [] };
  parent.children.splice(index, 0, node);
  save(); render();
  const el = [...document.querySelectorAll('.panel')].find(p => p._node === node);
  if (el) editTitle(node, el.querySelector('h4'));
}

// in-place rename: the element itself becomes editable, no box
// ---- undo/redo for structural changes (moves, deletes, adds) ----
// text edits use the browser's native undo while the field is focused
const undoStack = [], redoStack = [];
function treeSnap() {
  const idx = [];
  if (!currentTag) for (let i = 1; i < path.length; i++) idx.push(path[i - 1].children.indexOf(path[i]));
  return { root: structuredClone(storyTree), idx };
}
function pushUndo() {
  if (currentTag || !storyTree) return;   // ponytail: undo covers the story doc; tag pages rarely restructure
  undoStack.push(treeSnap());
  if (undoStack.length > 50) undoStack.shift();
  redoStack.length = 0;
  updateHistBtns();
}
function applySnap(s) {
  storyTree = s.root;
  tree = storyTree;
  currentTag = null;
  currentTagPath = null;
  path = [storyTree];
  let node = storyTree;
  for (const i of s.idx) {
    node = node.children[i];
    if (!node) break;
    path.push(node);
  }
  save(); render(); updateHistBtns();
}
function undo() { if (undoStack.length) { redoStack.push(treeSnap()); applySnap(undoStack.pop()); } }
function redo() { if (redoStack.length) { undoStack.push(treeSnap()); applySnap(redoStack.pop()); } }
function updateHistBtns() {
  $('undobtn').disabled = !undoStack.length;
  $('redobtn').disabled = !redoStack.length;
}
$('undobtn').onclick = undo;
$('redobtn').onclick = redo;
document.addEventListener('keydown', e => {
  if (!(e.ctrlKey || e.metaKey) || e.target.matches('input,textarea,[contenteditable="true"],[contenteditable="plaintext-only"]')) return;
  if (e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); }
  else if (e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); }
});

// ---- drag & drop reordering ----
let dragNode = null;
const containsNode = (n, t) => n === t || n.children.some(c => containsNode(c, t));
function findParent(root, node) {
  for (const c of root.children) {
    if (c === node) return root;
    const r = findParent(c, node);
    if (r) return r;
  }
  return null;
}
function moveNode(node, newParent, index) {
  if (!node || node === newParent || containsNode(node, newParent)) return;
  if (currentTag && storyTree && containsNode(storyTree, node)) return;  // outline drag while a tag page is open would save into the tag file
  const oldParent = findParent(tree, node);
  if (!oldParent) return;
  pushUndo();
  const oldIdx = oldParent.children.indexOf(node);
  oldParent.children.splice(oldIdx, 1);
  if (oldParent === newParent && oldIdx < index) index--;
  newParent.children.splice(index, 0, node);
  (function fix(k, d) { k.depth = d; k.children.forEach(x => fix(x, d + 1)); })(node, newParent.depth + 1);
  save(); render();
}
function makeDraggable(el, node) {
  el.draggable = true;
  el.addEventListener('dragstart', e => {
    if (e.target.isContentEditable) { e.preventDefault(); return; }
    dragNode = node;
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', '');
    e.stopPropagation();
  });
  el.addEventListener('dragend', () => { dragNode = null; });
}
function makeDropTarget(el, getTarget) {   // getTarget() -> [parent, index]
  el.addEventListener('dragover', e => {
    if (!dragNode) return;
    e.preventDefault();
    e.stopPropagation();
    el.classList.add('dragover');
  });
  el.addEventListener('dragleave', () => el.classList.remove('dragover'));
  el.addEventListener('drop', e => {
    e.preventDefault();
    e.stopPropagation();
    el.classList.remove('dragover');
    const [parent, index] = getTarget(e);
    moveNode(dragNode, parent, index);
  });
}

// deletion: trivial sections get a confirm; anything substantial requires
// typing DELETE so a stray click can't erase real writing
function deleteNode(node) {
  const wc = wordCount(node);
  const label = labelOf(node);
  if (wc >= 100 || node.children.length) {
    const detail = `${label}${node.children.length ? ` and its ${node.children.length} section${node.children.length > 1 ? 's' : ''}` : ''} — ${wc.toLocaleString()} words`;
    if (prompt(`This permanently deletes ${detail}.\n\nType DELETE to confirm:`) !== 'DELETE') return;
  } else if (!confirm(`Delete ${label}?`)) return;
  pushUndo();
  (function remove(parent) {
    const i = parent.children.indexOf(node);
    if (i >= 0) { parent.children.splice(i, 1); return true; }
    return parent.children.some(remove);
  })(tree);
  save(); render();
}

function editTitle(node, el) {
  if (readonly) return;
  el.contentEditable = 'plaintext-only';
  el.focus();
  document.getSelection().selectAllChildren(el);
  el.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); el.blur(); } e.stopPropagation(); };
  el.onclick = e => e.stopPropagation();
  el.onblur = () => { node.title = el.innerText.trim() || node.title; save(); render(); };
}

// ---- text settings (persisted) ----
const SETTINGS = { measure: { unit: 'px', def: 720 }, fsize: { unit: 'px', def: 17 }, lheight: { unit: '', def: 1.7 }, ui: { unit: 'px', def: 13 } };
const stored = JSON.parse(localStorage.getItem('writer-text') || '{}');
for (const key in SETTINGS) {
  const s = SETTINGS[key], input = $('s-' + key);
  const apply = v => {
    document.documentElement.style.setProperty('--' + key, v + s.unit);
    $('v-' + key).textContent = v + s.unit;
  };
  input.value = stored[key] || s.def;
  apply(input.value);
  input.oninput = () => {
    apply(input.value);
    stored[key] = +input.value;
    localStorage.setItem('writer-text', JSON.stringify(stored));
  };
}
const themeSel = $('themesel');
themeSel.value = localStorage.getItem('writer-theme') || 'dark';
document.body.classList.toggle('light', themeSel.value === 'light');
themeSel.onchange = () => {
  localStorage.setItem('writer-theme', themeSel.value);
  document.body.classList.toggle('light', themeSel.value === 'light');
};

$('upbtn').onclick = () => {
  if (path.length > 1) { path.pop(); render(); }
  else if (currentTag) openStory(storyName);   // tag page root -> back to the story
};
$('panelsbtn').onclick = () => { panelsHidden = !panelsHidden; render(); };
$('gridbtn').onclick = () => { layout.grid = !layout.grid; saveLayout(); render(); };

// ---- layout: collapsible sidebar, drag-resizable sidebar & notes ----
const layout = JSON.parse(localStorage.getItem('writer-layout') || '{}');
function applyLayout() {
  if (layout.sidew) document.documentElement.style.setProperty('--sidew', layout.sidew + 'px');
  if (layout.notesw) document.documentElement.style.setProperty('--notesw', layout.notesw + 'px');
  document.body.classList.toggle('sidecollapsed', !!layout.sidecollapsed);
}
applyLayout();
function saveLayout() { localStorage.setItem('writer-layout', JSON.stringify(layout)); applyLayout(); }
$('sidetoggle').onclick = () => { layout.sidecollapsed = !layout.sidecollapsed; saveLayout(); };

function makeGrip(grip, compute, key, min, max) {
  grip.onpointerdown = e => {
    e.preventDefault();
    grip.setPointerCapture(e.pointerId);
    grip.classList.add('dragging');
    grip.onpointermove = ev => { layout[key] = Math.min(max, Math.max(min, compute(ev))); saveLayout(); };
    grip.onpointerup = () => { grip.onpointermove = null; grip.classList.remove('dragging'); };
  };
}
makeGrip($('sidegrip'), ev => ev.clientX, 'sidew', 160, 480);
makeGrip($('notesgrip'), ev => innerWidth - ev.clientX, 'notesw', 260, Math.round(innerWidth * 0.7));

$('modebtn').onclick = () => {
  readonly = !readonly;
  const b = $('modebtn');
  b.innerHTML = readonly ? icon('book', 13) + ' reading' : icon('pencil', 13) + ' editing';
  b.classList.toggle('reading', readonly);
  render();
};

// ---- timeline slider ----
let flat = [];
function renderTimeline() {
  flat = [];
  (function walk(n, anc) {
    n.children.forEach(c => {
      const p = [...anc, c];
      flat.push({ node: c, path: p });
      walk(c, p);
    });
  })(tree, [tree]);
  flat.forEach((e, i) => e.pct = (i + 0.5) / flat.length * 100);

  $('timeline').hidden = !flat.length;
  const track = $('track');
  track.innerHTML = '';
  const f = focus();
  for (const e of flat) {
    const d = document.createElement('div');
    d.className = `tnode tn${Math.min(e.node.depth, 6)}` + (e.node === f ? ' current' : '');
    d.style.left = e.pct + '%';
    d.title = labelOf(e.node);
    track.appendChild(d);
  }
  const thumb = document.createElement('div');
  thumb.id = 'thumb';
  const cur = flat.find(e => e.node === f);
  thumb.style.left = (cur ? cur.pct : 0) + '%';
  thumb.style.opacity = cur ? 1 : 0;
  track.appendChild(thumb);
}
function nearestEntry(clientX) {
  const r = $('track').getBoundingClientRect();
  const frac = Math.min(1, Math.max(0, (clientX - r.left) / r.width));
  return flat[Math.min(flat.length - 1, Math.floor(frac * flat.length))];
}
let tlDragging = false;
$('timeline').onpointerdown = e => { tlDragging = true; $('timeline').setPointerCapture(e.pointerId); tlPreview(e); };
$('timeline').onpointermove = e => { if (tlDragging) tlPreview(e); };
$('timeline').onpointerup = e => {
  if (!tlDragging) return;
  tlDragging = false;
  const label = $('tlabel'); if (label) label.remove();
  const ent = nearestEntry(e.clientX);
  if (ent) { path = ent.path.slice(); render(); }
};
function tlPreview(e) {
  const ent = nearestEntry(e.clientX);
  if (!ent) return;
  const thumb = $('thumb');
  thumb.style.opacity = 1;
  thumb.style.left = ent.pct + '%';
  let label = $('tlabel');
  if (!label) { label = document.createElement('div'); label.id = 'tlabel'; thumb.appendChild(label); }
  label.textContent = labelOf(ent.node);
}

// ---- scroll: plain wheel scrolls, ctrl+wheel zooms (up = dive in, down = out) ----
let acc = 0;
$('main').addEventListener('wheel', e => {
  if (!tree) return;
  if (e.ctrlKey) {
    e.preventDefault();
    acc += e.deltaY;
    if (acc < -80) {
      acc = 0;
      const p = e.target.closest('.panel');
      if (p) { path.push(p._node); render(); }
    } else if (acc > 80) {
      acc = 0;
      if (path.length > 1) { path.pop(); render(); }
    }
    return;
  }
  // plain vertical wheel over the panel row scrolls it horizontally (grid scrolls natively)
  const panels = $('main').classList.contains('grid') ? null : e.target.closest('#panels');
  if (panels && Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
    panels.scrollLeft += e.deltaY;
    e.preventDefault();
  } else if (!e.target.closest('#focusbody,#backlinks,#timeline')) {
    $('focusbody').scrollTop += e.deltaY;   // scroll the text from anywhere on the page
  }
}, { passive: false });
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  if (!$('searchdrop').hidden) { closeSearch(); return; }
  if ($('ctxmenu')) { closeCtx(); return; }
  if (curMark) { closeAnnot(); return; }
  if ($('notespanel').classList.contains('open')) { $('notespanel').classList.remove('open'); return; }
  if (tree && path.length > 1 && !e.target.matches('input,textarea,[contenteditable]')) { path.pop(); render(); }
});

// ---- search (quick switcher + full text) ----
let searchSel = 0, searchHits = [];
function openSearch() {
  if (!storyTree) return;
  $('searchdrop').hidden = false;
  doSearch($('searchinput').value);
  $('searchinput').focus();
}
function closeSearch() {
  $('searchdrop').hidden = true;
  $('searchinput').value = '';
  $('searchinput').blur();
}
function doSearch(q) {
  const ql = q.trim().toLowerCase();
  searchHits = [];
  searchSel = 0;
  if (ql) {
    (function walk(n, anc) {
      for (const c of n.children) {
        const p = [...anc, c];
        const inTitle = c.title.toLowerCase().includes(ql);
        const bodyIdx = c.body.toLowerCase().indexOf(ql);
        const noteIdx = c.notes.toLowerCase().indexOf(ql);
        if (inTitle || bodyIdx >= 0 || noteIdx >= 0) {
          const src = bodyIdx >= 0 ? c.body : noteIdx >= 0 ? c.notes : '';
          const idx = bodyIdx >= 0 ? bodyIdx : noteIdx;
          searchHits.push({
            kind: (kindMap.get(c) || {}).label || 'Section',
            label: p.map(labelOf).join(' › '),
            snippet: src ? src.slice(Math.max(0, idx - 40), idx + ql.length + 60).replace(/\n/g, ' ') : '',
            q,
            rank: inTitle ? 0 : noteIdx >= 0 && bodyIdx < 0 ? 2 : 1,
            go: () => { currentTag = null; currentTagPath = null; tree = storyTree; path = [storyTree, ...p]; render(); },
          });
        }
      }
      n.children.forEach(c => walk(c, [...anc, c]));
    })(storyTree, []);
    for (const [name, cat] of tagIndex)
      if (name.toLowerCase().includes(ql))
        searchHits.push({ kind: 'tag · ' + cat, label: name, snippet: '', q, rank: 0, go: () => openTag(cat, name) });
    searchHits.sort((a, b) => a.rank - b.rank);
    searchHits = searchHits.slice(0, 50);
  }
  const box = $('searchresults');
  box.innerHTML = '';
  if (!searchHits.length) {
    const d = document.createElement('div');
    d.id = 'searchempty';
    d.textContent = ql ? 'No matches.' : 'Type to search this story.';
    box.appendChild(d);
    return;
  }
  searchHits.forEach((h, i) => {
    const row = document.createElement('div');
    row.className = 'sres' + (i === searchSel ? ' sel' : '');
    const where = document.createElement('div');
    where.className = 'swhere';
    where.innerHTML = `<span class="skind">${esc(h.kind)}</span>` + esc(h.label);
    row.appendChild(where);
    if (h.snippet) {
      const sn = document.createElement('div');
      sn.className = 'ssnip';
      sn.innerHTML = esc(h.snippet).replace(new RegExp(escRe(esc(h.q.trim())), 'i'), m => '<b>' + m + '</b>');
      row.appendChild(sn);
    }
    // mousedown (not click) so the input's blur doesn't dismiss the row first
    row.onmousedown = e => { e.preventDefault(); closeSearch(); h.go(); };
    box.appendChild(row);
  });
}
$('searchinput').onfocus = openSearch;
$('searchinput').onblur = () => { $('searchdrop').hidden = true; };
$('searchinput').oninput = () => doSearch($('searchinput').value);
$('searchinput').onkeydown = e => {
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault();
    searchSel = Math.min(searchHits.length - 1, Math.max(0, searchSel + (e.key === 'ArrowDown' ? 1 : -1)));
    document.querySelectorAll('.sres').forEach((r, i) => r.classList.toggle('sel', i === searchSel));
    const sel = document.querySelector('.sres.sel');
    if (sel) sel.scrollIntoView({ block: 'nearest' });
  } else if (e.key === 'Enter' && searchHits[searchSel]) {
    closeSearch();
    searchHits[searchSel].go();
  } else if (e.key === 'Escape') {
    closeSearch();
  }
};
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') { e.preventDefault(); openSearch(); }
});

// ---- tags ----
const tagDirPath = cat => cat === 'entities'
  ? FS.join(storyPath, 'entities')            // legacy folder
  : FS.join(storyPath, 'tags', cat);
const tagFilePath = (cat, name) => FS.join(tagDirPath(cat), name + '.md');

async function refreshTags() {
  tagIndex = new Map();
  const cats = [];   // {cat, names: []}
  async function scanDir(cat, dirPath) {
    const names = [];
    for (const f of await FS.readDir(dirPath)) {
      if (f.isDir || !f.name.endsWith('.md') || f.name.startsWith('_')) continue;
      const name = f.name.slice(0, -3);
      names.push(name);
      tagIndex.set(name, cat);
    }
    if (names.length || cat !== 'entities') cats.push({ cat, names: names.sort() });
  }
  try { await scanDir('entities', FS.join(storyPath, 'entities')); } catch {}
  try {
    for (const d of await FS.readDir(FS.join(storyPath, 'tags')))
      if (d.isDir) await scanDir(d.name, tagDirPath(d.name));
  } catch {}

  allCats = cats.map(c => c.cat);

  // explorer
  $('tagsection').hidden = false;
  const list = $('taglist');
  list.innerHTML = '';
  for (const { cat, names } of cats.sort((a, b) => a.cat.localeCompare(b.cat))) {
    const head = document.createElement('div');
    head.className = 'cat';
    const arrow = document.createElement('span');
    arrow.className = 'arrow';
    arrow.textContent = openCats.has(cat) ? '▾' : '▸';
    const catName = document.createElement('span');
    catName.textContent = cat;
    head.append(arrow, catName);
    head.onclick = () => { openCats.has(cat) ? openCats.delete(cat) : openCats.add(cat); refreshTags(); };
    const tbtn = document.createElement('span');
    tbtn.className = 'tbtn';
    tbtn.textContent = 'template';
    tbtn.title = `Edit the ${cat} template — new ${cat} tags start from it`;
    tbtn.onclick = e => { e.stopPropagation(); openTemplate(cat); };
    head.appendChild(tbtn);
    list.appendChild(head);
    if (!openCats.has(cat)) continue;
    for (const name of names) {
      const div = document.createElement('div');
      div.className = 'tag' + (currentTag && currentTag.name === name ? ' active' : '');
      const label = document.createElement('span');
      label.textContent = name;
      const del = document.createElement('span');
      del.className = 'rowdel';
      del.textContent = '✕';
      del.title = 'Delete tag';
      del.onclick = e => { e.stopPropagation(); deleteTag(cat, name); };
      div.append(label, del);
      div.onclick = () => openTag(cat, name);
      list.appendChild(div);
    }
  }
}

async function openTag(cat, name) {
  await FS.mkdir(tagDirPath(cat));
  currentTagPath = tagFilePath(cat, name);
  currentTag = { cat, name };
  let text = (await FS.exists(currentTagPath)) ? await FS.readText(currentTagPath) : '';
  if (!text.trim()) {
    text = await templateText(cat);
    await FS.writeText(currentTagPath, text);
  }
  tree = parse(text, name);
  path = [tree];
  await refreshTags();
  render();
}
async function templateText(cat) {
  try {
    const t = await FS.readText(FS.join(tagDirPath(cat), '_template.md'));
    if (t.trim()) return t;
  } catch {}
  return `type: ${cat}\n`;
}
async function openTemplate(cat) {
  await FS.mkdir(tagDirPath(cat));
  currentTagPath = FS.join(tagDirPath(cat), '_template.md');
  currentTag = { cat, name: cat + ' template', template: true };
  const text = (await FS.exists(currentTagPath)) ? await FS.readText(currentTagPath) : '';
  if (!text.trim()) await FS.writeText(currentTagPath, `type: ${cat}\n`);
  tree = parse(text || `type: ${cat}\n`, cat + ' template');
  path = [tree];
  render();
}
async function openTagByName(name) {
  let cat = tagIndex.get(name);
  if (!cat) {
    cat = prompt(`New tag "${name}" — category? (character, location, item, …)`, 'character');
    if (!cat) return;
    cat = cat.trim().toLowerCase();
  }
  openTag(cat, name);
}
// tags themselves are created from the writing (highlight -> right-click);
// the explorer builds the scaffolding: categories and their templates
$('newtag').onclick = async () => {
  const cat = prompt('New category name (character, location, item, …):');
  if (!cat) return;
  await openTemplate(cat.trim().toLowerCase());
  await refreshTags();
};

async function deleteTag(cat, name) {
  if (!confirm(`Delete tag "${name}"?`)) return;
  try {
    const p = tagFilePath(cat, name);
    const text = await FS.readText(p);
    if (text.trim()) await trashFile(`tag-${name}.md`, text);
    await FS.remove(p);
  } catch {}
  if (currentTag && currentTag.name === name) await openStory(storyName);
  else { await refreshTags(); render(); }
}

// create a tag file without navigating away from where you're writing
async function createTag(cat, name) {
  await FS.mkdir(tagDirPath(cat));
  const p = tagFilePath(cat, name);
  const existing = (await FS.exists(p)) ? await FS.readText(p) : '';
  if (!existing.trim()) await FS.writeText(p, await templateText(cat));
  await refreshTags();
  render();
}

// ---- highlight -> right-click -> tag or annotate ----
function closeCtx() { const m = $('ctxmenu'); if (m) m.remove(); }
document.addEventListener('contextmenu', e => {
  closeCtx();
  if (!tree || !storyPath) return;
  const sel = document.getSelection();
  const name = sel.toString().trim();
  if (!name || name.includes('\n') || !e.target.closest('#main')) return;

  const menu = document.createElement('div');
  menu.id = 'ctxmenu';
  const addItem = (label, fn) => {
    const it = document.createElement('div');
    it.className = 'item';
    it.style.textTransform = 'none';
    it.textContent = label;
    it.onclick = () => { closeCtx(); fn(); };
    menu.appendChild(it);
  };

  // clipboard basics (the custom menu replaces the native one)
  const selText = sel.toString();
  const savedRange = sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
  const restore = () => { const s = getSelection(); s.removeAllRanges(); if (savedRange) s.addRange(savedRange); };
  const inBody = !readonly && $('focusbody').contains(sel.anchorNode);
  addItem('Copy', () => navigator.clipboard.writeText(selText));
  if (inBody) {
    addItem('Cut', async () => {
      await navigator.clipboard.writeText(selText);
      restore();
      document.execCommand('delete');   // fires input -> body syncs + autosaves
    });
    addItem('Paste', async () => {
      const t = await navigator.clipboard.readText();
      restore();
      document.execCommand('insertText', false, t);
    });
  }

  // annotate: selection inside the editable body, not already annotated
  if (inBody && !sel.anchorNode.parentElement.closest('mark.annot')) {
    addItem('🗒 Add note to selection', () => { restore(); annotateSelection(getSelection()); });
  }

  if (name.length <= 60 && !tagIndex.has(name)) {
    const head = document.createElement('div');
    head.className = 'head';
    head.textContent = `Tag "${name}" as`;
    menu.appendChild(head);
    const cats = [...new Set([...allCats.filter(c => c !== 'entities'), 'character', 'location', 'item'])].sort();
    for (const cat of cats) {
      const it = document.createElement('div');
      it.className = 'item';
      it.textContent = cat;
      it.onclick = () => { closeCtx(); createTag(cat, name); };
      menu.appendChild(it);
    }
    const nw = document.createElement('div');
    nw.className = 'item new';
    nw.textContent = 'new category…';
    nw.onclick = () => {
      closeCtx();
      const cat = prompt('Category name:');
      if (cat) createTag(cat.trim().toLowerCase(), name);
    };
    menu.appendChild(nw);
  }

  if (!menu.children.length) return;
  e.preventDefault();
  document.body.appendChild(menu);
  const r = menu.getBoundingClientRect();
  menu.style.left = Math.min(e.clientX, innerWidth - r.width - 8) + 'px';
  menu.style.top = Math.min(e.clientY, innerHeight - r.height - 8) + 'px';
});
document.addEventListener('mousedown', e => { if (!e.target.closest('#ctxmenu')) closeCtx(); });

// ---- inline annotations ----
let curMark = null;
function annotateSelection(sel) {
  const range = sel.getRangeAt(0);
  const mark = document.createElement('mark');
  mark.className = 'annot';
  mark.dataset.note = '';
  try { range.surroundContents(mark); }
  catch { alert('Select plain text only (a selection can\'t cut across a link or another note).'); return; }
  sel.removeAllRanges();
  focus().body = domToMd($('focusbody'));
  queueSave();
  openAnnot(mark);
}
function openAnnot(mark) {
  curMark = mark;
  const pop = $('annotpop');
  pop.hidden = false;
  $('annottext').value = mark.dataset.note;
  const r = mark.getBoundingClientRect();
  pop.style.left = Math.max(8, Math.min(r.left, innerWidth - 340)) + 'px';
  pop.style.top = Math.min(r.bottom + 8, innerHeight - 190) + 'px';
  $('annottext').focus();
}
function closeAnnot() {
  if (!curMark) return;
  curMark = null;
  $('annotpop').hidden = true;
  clearTimeout(saveTimer);
  save();
  renderNotes();
}
$('annottext').oninput = () => {
  if (!curMark) return;
  const f = focus();
  const next = $('annottext').value;
  f.body = f.body.replace(`==${curMark.textContent}==%%${curMark.dataset.note}%%`, () => `==${curMark.textContent}==%%${next}%%`);
  curMark.dataset.note = next;
  queueSave();
};
$('annotdone').onclick = closeAnnot;
$('annotdel').onclick = () => {
  const f = focus();
  f.body = f.body.replace(`==${curMark.textContent}==%%${curMark.dataset.note}%%`, () => curMark.textContent);
  curMark = null;
  $('annotpop').hidden = true;
  save(); render();
};
document.addEventListener('mousedown', e => {
  if (curMark && !e.target.closest('#annotpop') && e.target !== curMark) closeAnnot();
});

// ---- library / stories ----
async function openLibrary(path) {
  libPath = path;
  if (FS.native) localStorage.setItem('writer-lib', path);
  $('newstory').hidden = false;
  await listStories();
}
$('openbtn').onclick = async () => {
  const p = await FS.pickFolder();
  if (p) await openLibrary(p);
};
async function listStories() {
  const list = $('storylist');
  list.innerHTML = '';
  for (const entry of await FS.readDir(libPath)) {
    if (!entry.isDir || entry.name.startsWith('.')) continue;
    const div = document.createElement('div');
    div.className = 'story';
    const label = document.createElement('span');
    label.textContent = entry.name;
    const del = document.createElement('span');
    del.className = 'rowdel';
    del.textContent = '✕';
    del.title = 'Delete story';
    del.onclick = async e => {
      e.stopPropagation();
      if (prompt(`This permanently deletes the story "${entry.name}" and ALL its files.\n\nType the story name to confirm:`) !== entry.name) return;
      await FS.removeDir(FS.join(libPath, entry.name));
      await listStories();
    };
    div.append(label, del);
    div.onclick = () => openStory(entry.name, div);
    list.appendChild(div);
  }
}
// swap ![[chapters/...]] embed lines in part bodies for the loaded chapter
// trees; old single-file stories have no embeds and load unchanged (their
// inline ## chapters convert to files on the next save)
async function resolveChapterEmbeds(sPath, root) {
  chapterCache = new Map();
  const chDir = FS.join(sPath, 'chapters');
  if (!(await FS.exists(chDir))) return;
  const nodes = [root];
  (function all(n) { n.children.forEach(c => { nodes.push(c); all(c); }); })(root);
  for (const node of nodes) {
    if (!node.body.includes('![[chapters/')) continue;
    const refs = [], keep = [];
    for (const line of node.body.split('\n')) {
      const m = line.trim().match(/^!\[\[chapters\/(.+?)\]\]$/);
      if (m) refs.push(m[1] + '.md'); else keep.push(line);
    }
    node.body = keep.join('\n').trim();
    const loaded = [];
    for (const fname of refs) {
      try {
        const text = await FS.readText(FS.join(chDir, fname));
        const c = parse(text, '', { unwrap: false });
        const chapter = { depth: 0, title: fname.replace(/^\d+\s*/, '').replace(/\.md$/, ''), body: c.body, notes: c.notes, children: c.children };
        (function fix(k, d) { k.depth = d; k.children.forEach(x => fix(x, d + 1)); })(chapter, node.depth + 1);
        chapterCache.set(fname, serializeChapter(chapter));
        loaded.push(chapter);
      } catch {
        // keep a placeholder so a missing file is visible and its slot survives saves
        loaded.push({ depth: node.depth + 1, title: fname.replace(/\.md$/, ''), body: '(missing chapter file: ' + fname + ')', notes: '', children: [] });
      }
    }
    node.children = [...loaded, ...node.children];
  }
}

async function openStory(name, el) {
  document.querySelectorAll('.story').forEach(s => s.classList.remove('active'));
  if (el) el.classList.add('active');
  storyName = name;
  storyPath = FS.join(libPath, name);
  $('librarysection').hidden = true;
  $('storyheader').hidden = false;
  $('storyname').textContent = name;
  currentTag = null;
  currentTagPath = null;
  const sf = FS.join(storyPath, 'story.md');
  const text = (await FS.exists(sf)) ? await FS.readText(sf) : '';
  storyBackup = text;
  const fm = splitFrontmatter(text);
  storyMeta = fm.meta;
  tree = parse(fm.body, name);
  await resolveChapterEmbeds(storyPath, tree);
  undoStack.length = 0; redoStack.length = 0; updateHistBtns();
  storyTree = tree;
  path = [tree];
  await refreshTags();
  render();
  // watch for external edits (Obsidian, scripts) — reload unless we just saved
  if (unwatchStory) { unwatchStory(); unwatchStory = null; }
  unwatchStory = await FS.watch(storyPath, () => {
    if (Date.now() - lastSaveAt < 2500) return;
    clearTimeout(watchTimer);
    watchTimer = setTimeout(async () => {
      const idx = [];
      for (let i = 1; i < path.length; i++) idx.push(path[i - 1].children.indexOf(path[i]));
      const wasTag = currentTag;
      await openStory(storyName);
      if (!wasTag) {
        let node = storyTree;
        for (const i of idx) { node = node.children[i]; if (!node) break; path.push(node); }
        render();
      }
    }, 400);
  });
}
$('backlib').onclick = () => {
  $('librarysection').hidden = false;
  $('storyheader').hidden = true;
  $('tagsection').hidden = true;
  if (unwatchStory) { unwatchStory(); unwatchStory = null; }
  storyPath = null; storyName = ''; tree = null; storyTree = null; currentTag = null;
  render();
};
$('storyname').onclick = () => openStory(storyName);
$('storyinfobtn').onclick = () => {
  $('si-author').value = storyMeta.author || '';
  $('si-byline').value = storyMeta.byline || '';
  $('si-contact').value = [].concat(storyMeta.contact || []).join('\n');
  $('storyinfodlg').showModal();
};
$('storyinfoform').onsubmit = () => {
  storyMeta.author = $('si-author').value.trim();
  storyMeta.byline = $('si-byline').value.trim();
  storyMeta.contact = $('si-contact').value.split('\n').map(s => s.trim()).filter(Boolean);
  save();
};
makeDropTarget($('storyname'), () => [storyTree, storyTree.children.length]);   // drop = promote to top level
$('newstory').onclick = () => { $('ns-name').value = ''; $('newstorydlg').showModal(); };
$('ns-template').onchange = () => { $('ns-custom').hidden = $('ns-template').value !== 'custom'; };
const TEMPLATES = { blank: [], short: [3], novel: [3, 4, 2], epic: [3, 3, 3, 2] };
function buildLevels(counts, depth = 1) {
  if (!counts.length) return [];
  return Array.from({ length: counts[0] }, () =>
    ({ depth, title: '', body: '', notes: '', children: buildLevels(counts.slice(1), depth + 1) }));
}
$('newstoryform').onsubmit = async e => {
  e.preventDefault();
  const name = $('ns-name').value.trim();
  if (!name) return;
  const tpl = $('ns-template').value;
  const counts = tpl === 'custom'
    ? Array(Math.min(5, +$('ns-depth').value || 1)).fill(Math.min(8, +$('ns-count').value || 1))
    : TEMPLATES[tpl];
  $('newstorydlg').close();
  await FS.mkdir(FS.join(libPath, name));
  await listStories();
  await openStory(name);
  if (counts.length && !tree.children.length) {
    tree.children = buildLevels(counts);
    await save();
    render();
  }
};

// ---- export / compile ----
// notes, annotations' notes, and [[brackets]] never reach the manuscript
function compileNode(n, base, opts) {
  const out = [];
  if (n.depth > base) {   // the export root's own title becomes the document title instead
    const layer = opts.layers[opts.H - n.depth] || { title: true, contents: true };
    if (layer.title) {
      const label = (kindMap.get(n) || {}).label || '';
      const title = layer.numbers
        ? label + (n.title ? ' — ' + n.title : '')       // "Chapter 1 — Down the Rabbit-Hole"
        : (n.title || label || 'Untitled');
      out.push('#'.repeat(Math.min(6, n.depth - base + 1)) + ' ' + title);
    }
    const body = cleanText(n.body);
    if (body && layer.contents) out.push(body);
  }
  let prevLeaf = false;
  for (const c of n.children) {
    const isLeaf = !c.children.length;
    if (isLeaf && prevLeaf && opts.sep) out.push(opts.sep);
    out.push(...compileNode(c, base, opts));
    prevLeaf = isLeaf;
  }
  return out;
}
function titlePageData(root) {
  const wc = wordCount(root);
  return {
    approx: (Math.max(100, Math.round(wc / 100) * 100)).toLocaleString(),
    by: storyMeta.byline || storyMeta.author || '',
    contact: [].concat(storyMeta.contact || []),
    author: storyMeta.author || '',
  };
}
function mdToHtml(md, title, tp) {
  const blocks = md.split('\n\n').map(b => {
    const h = b.match(/^(#{1,6}) (.*)$/s);
    if (h) return `<h${h[1].length}>${esc(h[2])}</h${h[1].length}>`;
    if (b.trim() === '* * *') return '<p class="sep">&#42; &#42; &#42;</p>';   // entity-escaped so the emphasis regex can't eat it
    return b.split(/\n\n+/).map(p => '<p>' + esc(p).replace(/\n/g, '<br>') + '</p>').join('');
  }).join('\n');
  const titlePage = tp ? `<div class="titlepage">
    <div class="tp-head"><div class="tp-contact">${tp.contact.map(esc).join('<br>')}</div><div class="tp-wc">approx. ${tp.approx} words</div></div>
    <div class="tp-mid"><h1>${esc(title)}</h1>${tp.by ? `<p class="tp-by">by ${esc(tp.by)}</p>` : ''}</div>
  </div>` : `<h1>${esc(title)}</h1>`;
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${esc(title)}</title><style>
  body{font-family:Georgia,serif;max-width:640px;margin:40px auto;padding:0 20px;line-height:1.7;color:#222}
  h1,h2,h3,h4{line-height:1.3} .sep{text-align:center;margin:2em 0}
  em{font-style:italic}
  .titlepage{min-height:90vh;display:flex;flex-direction:column}
  .tp-head{display:flex;justify-content:space-between;align-items:flex-start;line-height:1.5}
  .tp-contact{text-align:left} .tp-wc{text-align:right}
  .tp-mid{flex:1;display:flex;flex-direction:column;justify-content:center;text-align:center}
  .tp-by{margin-top:.4em;font-size:1.15em}
  @media print{body{margin:0;max-width:none}.titlepage{min-height:auto;height:95vh;page-break-after:always}}
  </style></head><body>${titlePage}\n${blocks
    .replace(/\*\*([^*\n]+)\*\*/g, '<strong>$1</strong>')
    .replace(/\*([^*\n]+)\*/g, '<em>$1</em>')
    .replace(/_([^_\n]+)_/g, '<em>$1</em>')}\n</body></html>`;
}
function download(name, text, type) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([text], { type }));
  a.download = name;
  a.click();
  URL.revokeObjectURL(a.href);
}
$('exportbtn').onclick = () => {
  if (!storyTree) return;
  // one row per layer in this story: check what compiles (titles as headings,
  // contents as text). Defaults: containers = title only, scenes = prose only.
  const H = treeHeight(storyTree);
  const box = $('ex-layers');
  box.innerHTML = '';
  for (let li = H - 1; li >= 0; li--) {
    const row = document.createElement('div');
    row.className = 'exrow';
    const name = document.createElement('span');
    name.textContent = LADDER[Math.min(li, LADDER.length - 1)] + 's';
    row.appendChild(name);
    for (const [key, label, def] of [['title', 'titles', li > 0], ['numbers', 'numbers', li > 0], ['contents', 'contents', li === 0]]) {
      const l = document.createElement('label');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.id = `ex-l-${key}-${li}`;
      cb.checked = def;
      l.append(cb, label);
      row.appendChild(l);
    }
    box.appendChild(row);
  }
  $('exportdlg').showModal();
};
$('exportform').onsubmit = e => {
  e.preventDefault();
  const scope = $('ex-scope').value;
  const root = scope === 'current' && !currentTag ? focus() : storyTree;
  const title = root === storyTree ? storyTree.title : (root.title || labelOf(root));
  const H = treeHeight(storyTree);
  const layers = {};
  for (let li = H - 1; li >= 0; li--) {
    const t = $(`ex-l-title-${li}`), c = $(`ex-l-contents-${li}`), nu = $(`ex-l-numbers-${li}`);
    if (t) layers[li] = { title: t.checked, contents: c.checked, numbers: nu.checked };
  }
  const opts = { sep: $('ex-sep').value, H, layers };
  const blocks = compileNode(root, root.depth, opts);
  const md = blocks.join('\n\n');
  $('exportdlg').close();
  const safe = title.replace(/[\\/:*?"<>|]/g, '');
  const tp = $('ex-titlepage').checked ? titlePageData(root) : null;
  if ($('ex-format').value === 'docx') {
    const bytes = makeDocx(blocks, {
      title,
      titlePage: tp,
      smf: $('ex-smf').checked,
      author: storyMeta.author || (tp && tp.by) || '',
    });
    download(safe + '.docx', bytes, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  } else if ($('ex-format').value === 'html') {
    download(safe + '.html', mdToHtml(md, title, tp), 'text/html');
  } else {
    const head = tp
      ? (tp.contact.length ? tp.contact.join('\n') + '\n\n' : '') +
        `approx. ${tp.approx} words\n\n# ${title}\n\n${tp.by ? `by ${tp.by}\n\n` : ''}---\n\n`
      : `# ${title}\n\n`;
    download(safe + '.md', head + md + '\n', 'text/markdown');
  }
};

// ---- settings ----
$('settingsbtn').onclick = () => { $('ladderinput').value = LADDER.join(', '); $('settingsdlg').showModal(); };
$('settingsform').onsubmit = () => {
  const names = $('ladderinput').value.split(',').map(s => s.trim()).filter(Boolean);
  if (names.length) {
    LADDER = names;
    localStorage.setItem('writer-ladder', JSON.stringify(LADDER));
  }
  render();
};

// ---- startup: reopen the last library automatically (native only) ----
(async function init() {
  if (!FS.native) return;
  const saved = localStorage.getItem('writer-lib');
  if (saved && await FS.exists(saved).catch(() => false)) await openLibrary(saved);
})();
