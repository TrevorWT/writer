import { parse, serialize, treeHeight, chapterFileName, serializeChapter, serializeSkeleton, cleanText, wordCount, splitFrontmatter, buildFrontmatter } from './parser.js';
import { makeDocx } from './docx.js';
import { initFS } from './fs.js';
const FS = initFS();

import { icon } from './icons.js';

// static chrome icons — null-safe so a stale cached page can't kill startup
const setIcon = (id, html) => { const el = document.getElementById(id); if (el) el.innerHTML = html; };
setIcon('sidetoggle', icon('sidebar', 18));
setIcon('settingsbtn', icon('gear', 17));
setIcon('guidebtn', icon('help', 17));
setIcon('exportbtn', icon('upload', 17));
setIcon('statsbtn', icon('chart', 17));
setIcon('tsicon', icon('search', 13));
setIcon('samplebtn', icon('book', 14));
setIcon('openbtn', icon('folder', 14));
setIcon('hintopen', icon('folder', 14) + ' Open a folder');
setIcon('hintsample', icon('book', 14) + ' Try the sample');
setIcon('hintnew', icon('plus', 14) + ' New Story');
setIcon('newstory', icon('plus', 14));
setIcon('newtag', icon('plus', 14));
setIcon('undobtn', icon('undo', 14));
setIcon('redobtn', icon('redo', 14));
setIcon('modebtn', icon('pencil', 15) + ' Edit');
setIcon('view-row', icon('columns', 14));
setIcon('view-grid', icon('grid', 14));
setIcon('view-hidden', icon('square', 14));
setIcon('storyinfobtn', icon('pencil', 12));
setIcon('historybtn', icon('history', 12));
setIcon('snapbtn', icon('camera', 13) + ' Snapshot now');
setIcon('upzoneicon', icon('arrowup', 16));
setIcon('upbtn', icon('arrowup', 18));
setIcon('prevbtn', icon('chevleft', 18));
setIcon('nextbtn', icon('chevright', 18));
document.querySelectorAll('#guidedlg .gicon').forEach(el => { el.innerHTML = icon(el.dataset.i, 17); });
{
  const gb = document.getElementById('guidebtn'), gd = document.getElementById('guidedlg');
  if (gb && gd) {
    gb.onclick = () => gd.showModal();
    document.getElementById('guideclose').onclick = () => { localStorage.setItem('writer-toured', '1'); gd.close(); };
    if (!localStorage.getItem('writer-toured')) setTimeout(() => gd.showModal(), 400);
  }
}

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

// in-app replacements for confirm/prompt/alert — Tauri makes the native ones
// async or unavailable, so window.* versions are unusable in the packaged app
function askDialog({ message, input = false, def = '', alertOnly = false }) {
  return new Promise(resolve => {
    const dlg = $('askdlg');
    $('askmsg').textContent = message;
    const inp = $('askinput');
    inp.hidden = !input;
    inp.value = def;
    $('askcancel').hidden = alertOnly;
    let ok = false;
    $('askform').onsubmit = () => { ok = true; };
    $('askcancel').onclick = () => { ok = false; dlg.close(); };
    dlg.onclose = () => resolve(ok ? (input ? inp.value : true) : (input ? null : false));
    dlg.showModal();
    if (input) { inp.focus(); inp.select(); }
  });
}
const appConfirm = m => askDialog({ message: m });
const appPrompt = (m, def = '') => askDialog({ message: m, input: true, def });
const appAlert = m => askDialog({ message: m, alertOnly: true });

function setStatus(text, err = false) {
  $('savestatus').textContent = text;
  $('savestatus').classList.toggle('err', err);
}
function updateSectionStatus(kindLabel, wc) {
  const el = $('sectionstatus');
  if (!tree) { el.hidden = true; return; }
  el.hidden = false;
  const words = `${wc.toLocaleString()} word${wc === 1 ? '' : 's'}`;
  el.textContent = `${kindLabel || 'Section'} - ${words}`;
  el.title = storyPath || libPath || '';
}

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
      if (c.nonum) {
        kindMap.set(c, { kind, num: null, label: kind });   // prologues etc: unnumbered, doesn't consume a number
      } else {
        counts[kind] = (counts[kind] || 0) + 1;
        kindMap.set(c, { kind, num: counts[kind], label: kind + ' ' + counts[kind] });
      }
      walk(c);
    }
  })(storyTree);
}
const STATUSES = ['', 'outline', 'draft', 'revised', 'done'];
function setNodeStatus(node, status) {
  node.status = status;
  save(); render();
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

// ---- history: snapshots + automatic backups in .trash/ ----
const trashDir = () => FS.join(storyPath, '.trash');
async function listHistory() {
  if (!storyPath || !(await FS.exists(trashDir()).catch(() => false))) return [];
  const out = [];
  for (const f of await FS.readDir(trashDir())) {
    if (f.isDir) continue;
    const m = f.name.match(/^(\d{13}) (.+)$/);
    if (!m) continue;
    const ts = +m[1], rest = m[2];
    let type = 'chapter', label = rest.replace(/\.md$/, '');
    if (rest === 'story.md') { type = 'skeleton'; label = 'story.md backup'; }
    else if (rest.startsWith('snapshot')) { type = 'snapshot'; label = rest.replace(/^snapshot ?/, '').replace(/\.md$/, '') || 'snapshot'; }
    else if (rest.startsWith('tag')) { type = 'tag'; label = rest.replace(/^tag[ -]/, '').replace(/\.md$/, ''); }
    out.push({ file: f.name, ts, rest, type, label });
  }
  return out.sort((a, b) => b.ts - a.ts);
}
async function takeSnapshot(label = '') {
  if (!storyTree) return;
  await FS.mkdir(trashDir());
  const text = buildFrontmatter(storyMeta) + serialize(storyTree);
  await FS.writeText(FS.join(trashDir(), `${Date.now()} snapshot${label ? ' ' + label : ''}.md`), text);
}
async function pruneHistory() {
  try {
    const days = Math.max(1, +prefs.historyDays || 30);
    const cutoff = Date.now() - days * 86400000;
    const autos = (await listHistory()).filter(e => e.type !== 'snapshot');
    const keep = new Set(autos.slice(0, 10).map(e => e.file));   // always keep the 10 newest
    for (const e of autos)
      if (e.ts < cutoff && !keep.has(e.file)) await FS.remove(FS.join(trashDir(), e.file)).catch(() => {});
  } catch {}
}
async function restoreHistory(ent) {
  const text = await FS.readText(FS.join(trashDir(), ent.file));
  if (ent.type === 'snapshot' || ent.type === 'skeleton') {
    if (!(await appConfirm('Restore this version? The current story is snapshotted first, so nothing is lost.'))) return;
    await takeSnapshot('before restore');
    if (ent.type === 'skeleton') {
      await FS.writeText(FS.join(storyPath, 'story.md'), text);
      storyBackup = null;
      $('historydlg').close();
      await openStory(storyName);
      return;
    }
    const fm = splitFrontmatter(text);
    storyMeta = fm.meta;
    const t = parse(fm.body, storyName);
    storyTree = t; tree = t; currentTag = null; currentTagPath = null; path = [t];
    chapterCache = new Map();
    undoStack.length = 0; redoStack.length = 0; updateHistBtns();
    await save();
    $('historydlg').close();
    render();
    setStatus('Snapshot restored');
  } else if (ent.type === 'chapter') {
    if (!(await appConfirm('Add this backed-up chapter to the end of the story?'))) return;
    const cparsed = parse(text, '', { unwrap: false });
    const node = {
      depth: 1,
      title: (ent.rest.replace(/^\d+\s*/, '').replace(/\.md$/, '') || 'Restored chapter'),
      body: cparsed.body, notes: cparsed.notes, children: cparsed.children,
    };
    const host = storyTree.children.length ? storyTree.children[storyTree.children.length - 1] : storyTree;
    (function fix(k, d) { k.depth = d; k.children.forEach(x => fix(x, d + 1)); })(node, host.depth + 1);
    pushUndo();
    host.children.push(node);
    currentTag = null; currentTagPath = null; tree = storyTree;
    await save();
    $('historydlg').close();
    render();
  } else {   // tag backup
    const m = ent.rest.match(/^tag (\S+) (.+)\.md$/);
    const cat = m ? m[1] : 'restored';
    const name = m ? m[2] : ent.rest.replace(/^tag-?/, '').replace(/\.md$/, '');
    await FS.mkdir(tagDirPath(cat));
    await FS.writeText(tagFilePath(cat, name), text);
    await refreshTags();
    render();
    setStatus(`Tag "${name}" restored`);
    renderHistory();
  }
}
function timeAgo(ts) {
  const m = (Date.now() - ts) / 60000;
  if (m < 1) return 'just now';
  if (m < 60) return Math.round(m) + ' min ago';
  if (m < 1440) return Math.round(m / 60) + ' h ago';
  if (m < 10080) return Math.round(m / 1440) + ' d ago';
  return new Date(ts).toLocaleDateString();
}
function historyRow(ent, compact) {
  const row = document.createElement('div');
  row.className = 'hrow' + (compact ? ' hsub' : '');
  if (!compact) {
    const type = document.createElement('span');
    type.className = 'htype t-' + ent.type;
    type.textContent = ent.type === 'skeleton' ? 'backup' : ent.type;
    row.appendChild(type);
  }
  const label = document.createElement('span');
  label.className = 'hlabel';
  label.textContent = compact ? timeAgo(ent.ts) : ent.label;
  row.appendChild(label);
  const date = document.createElement('span');
  date.className = 'hdate';
  date.textContent = compact ? new Date(ent.ts).toLocaleString() : timeAgo(ent.ts);
  row.appendChild(date);
  const rst = document.createElement('button');
  rst.type = 'button';
  rst.textContent = 'Restore';
  rst.onclick = () => restoreHistory(ent);
  row.appendChild(rst);
  const del = document.createElement('span');
  del.className = 'rowdel';
  del.style.visibility = 'visible';
  del.textContent = '✕';
  del.title = 'Delete permanently';
  del.onclick = async () => {
    if (!(await appConfirm('Delete this history entry permanently?'))) return;
    await FS.remove(FS.join(trashDir(), ent.file));
    renderHistory();
  };
  row.appendChild(del);
  return row;
}
async function renderHistory() {
  const list = $('historylist');
  list.innerHTML = '';
  const ents = await listHistory();
  if (!ents.length) {
    list.innerHTML = '<p class="setting-note">Nothing here yet. Take a snapshot before a big rewrite - it saves the whole story as one file.</p>';
    return;
  }
  const snaps = ents.filter(e => e.type === 'snapshot');
  const autos = ents.filter(e => e.type !== 'snapshot');
  if (snaps.length) {
    const h = document.createElement('div');
    h.className = 'hgrouphead';
    h.textContent = 'Snapshots - kept until you delete them';
    list.appendChild(h);
    for (const s of snaps) list.appendChild(historyRow(s, false));
  }
  if (autos.length) {
    const h = document.createElement('div');
    h.className = 'hgrouphead';
    h.textContent = 'Automatic backups - older copies of overwritten or deleted files';
    list.appendChild(h);
    const groups = new Map();
    for (const e of autos) {
      const key = e.type + '|' + e.label;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(e);
    }
    for (const [, group] of groups) {
      const latest = group[0];   // listHistory sorts newest first
      if (group.length === 1) {
        list.appendChild(historyRow(latest, false));
        continue;
      }
      const det = document.createElement('details');
      det.className = 'hgroup';
      const sum = document.createElement('summary');
      const type = document.createElement('span');
      type.className = 'htype t-' + latest.type;
      type.textContent = latest.type === 'skeleton' ? 'backup' : latest.type;
      const label = document.createElement('span');
      label.className = 'hlabel';
      label.textContent = latest.label;
      const meta = document.createElement('span');
      meta.className = 'hdate';
      meta.textContent = `${group.length} versions · latest ${timeAgo(latest.ts)}`;
      sum.append(type, label, meta);
      det.appendChild(sum);
      for (const e of group) det.appendChild(historyRow(e, true));
      list.appendChild(det);
    }
  }
}
$('clearbackupsbtn').onclick = async () => {
  const autos = (await listHistory()).filter(e => e.type !== 'snapshot');
  if (!autos.length) return;
  if (!(await appConfirm(`Permanently delete all ${autos.length} automatic backups? Snapshots are kept.`))) return;
  for (const e of autos) await FS.remove(FS.join(trashDir(), e.file)).catch(() => {});
  renderHistory();
};
$('historybtn').onclick = () => { $('historydlg').showModal(); renderHistory(); };
$('snapbtn').onclick = async () => {
  const label = await appPrompt('Label this snapshot (optional):', '');
  if (label === null) return;
  await takeSnapshot(label.trim().replace(/[\\/:*?"<>|]/g, ''));
  setStatus('Snapshot saved');
  renderHistory();
};

async function save() {
  setStatus('Saving...');
  try {
    lastSaveAt = Date.now();
    if (currentTag) {
      await FS.writeText(currentTagPath, serialize(tree));
      lastSaveAt = Date.now();
      setStatus('Saved');
      return;
    }
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
    // sweep chapter files that no longer correspond to a chapter - NEVER hard
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
    setStatus('Saved');
    updateGoal();
    relinkBody();
  } catch (e) {
    console.error(e);
    setStatus('Save failed - check console', true);
    throw e;
  }
}
let saveTimer = null;
function queueSave() { setStatus('Unsaved changes'); clearTimeout(saveTimer); saveTimer = setTimeout(save, 600); }

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
// ---- live linkify while typing: re-render links on save, preserving caret ----
// offsets mirror domToMd: text node chars count, DIVs and meaningful BRs = 1
function caretOffsetIn(el) {
  const sel = getSelection();
  if (!sel.rangeCount || !el.contains(sel.anchorNode)) return null;
  const r = sel.getRangeAt(0);
  let off = 0, found = false;
  (function walk(n) {
    if (found) return;
    if (n.nodeType === 3) {
      if (n === r.startContainer) { off += r.startOffset; found = true; }
      else off += n.textContent.length;
      return;
    }
    if (n.tagName === 'BR') { if (n.parentNode.childNodes.length > 1) off += 1; return; }
    if (n.tagName === 'DIV' && n !== el) off += 1;
    if (n === r.startContainer) {
      for (let i = 0; i < r.startOffset && !found; i++) walk(n.childNodes[i]);
      found = true;
      return;
    }
    for (const c of n.childNodes) { walk(c); if (found) return; }
  })(el);
  return found ? off : null;
}
function setCaretAt(el, off) {
  let remaining = off, done = false;
  (function walk(n) {
    if (done) return;
    if (n.nodeType === 3) {
      if (remaining <= n.textContent.length) {
        const r = document.createRange();
        r.setStart(n, remaining);
        r.collapse(true);
        const s = getSelection();
        s.removeAllRanges();
        s.addRange(r);
        done = true;
        return;
      }
      remaining -= n.textContent.length;
      return;
    }
    for (const c of n.childNodes) { walk(c); if (done) return; }
  })(el);
  if (!done) {
    const r = document.createRange();
    r.selectNodeContents(el);
    r.collapse(false);
    const s = getSelection();
    s.removeAllRanges();
    s.addRange(r);
  }
}
function relinkBody() {
  const fb = $('focusbody');
  if (readonly || !tree || document.activeElement !== fb) return;
  const f = focus();
  const html = linkify(f.body, true);
  if (fb.innerHTML === html) return;
  const off = caretOffsetIn(fb);
  fb.innerHTML = html;
  if (off !== null) setCaretAt(fb, off);
}

const focus = () => path[path.length - 1];
const focusHasChildren = () => tree && focus().children.length > 0;
let panelsHidden = false;

function render() {
  $('statsview').hidden = true;   // any navigation dismisses the stats overlay
  const has = !!tree;
  $('main').classList.toggle('empty', !has);
  $('hint').hidden = has;
  $('panels').hidden = !has;
  if (!has) {
    $('hinttitle').textContent = libPath ? 'Choose a story or create a new one.' : 'Open a folder to begin.';
    $('hinttext').textContent = libPath
      ? 'Pick an existing story, or create a new Markdown story folder.'
      : 'Stories are Markdown folders with layered sections.';
    $('hintopen').hidden = !!libPath;
    $('hintsample').hidden = !!libPath;
    $('hintnew').hidden = !libPath;
    $('sectionstatus').hidden = true;
  }
  $('focus').hidden = !has;
  $('modebtn').hidden = !has;
  $('topbtns').hidden = !has;
  $('topsearch').hidden = !has;
  $('navbtns').hidden = !has;
  $('navbtns').style.visibility = has && (path.length > 1 || currentTag) ? 'visible' : 'hidden';
  {
    const par = path.length > 1 ? path[path.length - 2] : null;
    const i = par ? par.children.indexOf(path[path.length - 1]) : -1;
    $('prevbtn').disabled = !par || i <= 0;
    $('nextbtn').disabled = !par || i < 0 || i >= par.children.length - 1;
    const sceneView = has && path.length > 1 && !focus().children.length;
    $('edgeleft').hidden = !sceneView;
    $('edgeright').hidden = !sceneView;
  }
  $('histbtns').hidden = !has || !!currentTag;
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
  const fk = $('focuskind');
  fk.textContent = kindLabel + (wc && !currentTag ? ` · ${wc.toLocaleString()} words` : '');
  if (!currentTag && f.depth > 0) {
    const chip = document.createElement('span');
    chip.className = 'statuschip focus-status ' + (f.status ? 'st-' + f.status : 'st-none');
    chip.textContent = f.status || 'no status';
    chip.title = 'Click to change status';
    chip.onclick = () => setNodeStatus(f, STATUSES[(STATUSES.indexOf(f.status || '') + 1) % STATUSES.length]);
    fk.appendChild(chip);
  }
  updateSectionStatus(kindLabel, wc);
  $('main').classList.toggle('writing', !f.children.length && (f.depth > 0 || !!currentTag));
  $('main').classList.toggle('hidepanels', panelsHidden && f.children.length > 0);
  $('viewseg').hidden = !has || !focusHasChildren();
  $('view-row').classList.toggle('active', !panelsHidden && !layout.grid);
  $('view-grid').classList.toggle('active', !panelsHidden && !!layout.grid);
  $('view-hidden').classList.toggle('active', panelsHidden);
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
  updateGoal();
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
$('notestab').onclick = () => {
  const open = $('notespanel').classList.toggle('open');
  if (open) $('notestext').focus();
};
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
      const more = document.createElement('span');
      more.className = 'rowmore';
      more.innerHTML = icon('more', 13);
      more.title = 'Section actions';
      more.onclick = e => {
        e.stopPropagation();
        const r = more.getBoundingClientRect();
        showSectionMenu(c, r.left, r.bottom + 4);
      };
      const del = document.createElement('span');
      del.className = 'rowdel';
      del.textContent = '✕';
      del.title = 'Delete section';
      del.onclick = e => {
        e.stopPropagation();
        if (currentTag) { currentTag = null; currentTagPath = null; tree = storyTree; path = [storyTree]; }
        deleteNode(c);
      };
      row.append(arrow, label, more, del);
      row.onclick = () => {
        currentTag = null; currentTagPath = null; tree = storyTree;
        path = [storyTree, ...p];
        render();
      };
      makeDraggable(row, c);
      // drop zones: top third = before, bottom third = after, middle = nest into
      row.addEventListener('dragover', e => {
        if (!dragNode) return;
        e.preventDefault();
        e.stopPropagation();
        const r = row.getBoundingClientRect();
        const y = e.clientY - r.top;
        row.classList.remove('drop-top', 'drop-bottom', 'dragover');
        if (y < r.height * 0.3) row.classList.add('drop-top');
        else if (y > r.height * 0.7) row.classList.add('drop-bottom');
        else row.classList.add('dragover');
      });
      row.addEventListener('dragleave', () => row.classList.remove('drop-top', 'drop-bottom', 'dragover'));
      row.addEventListener('drop', e => {
        e.preventDefault();
        e.stopPropagation();
        const r = row.getBoundingClientRect();
        const y = e.clientY - r.top;
        row.classList.remove('drop-top', 'drop-bottom', 'dragover');
        const parent = findParent(storyTree, c) || storyTree;
        const idx = parent.children.indexOf(c);
        if (y < r.height * 0.3) moveNode(dragNode, parent, idx);
        else if (y > r.height * 0.7) moveNode(dragNode, parent, idx + 1);
        else moveNode(dragNode, c, c.children.length);
      });
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
  const more = document.createElement('span');
  more.className = 'pmore';
  more.innerHTML = icon('more', 14);
  more.title = 'Section actions';
  more.onclick = e => {
    e.stopPropagation();
    const r = more.getBoundingClientRect();
    showSectionMenu(node, r.left, r.bottom + 4);
  };
  el.appendChild(more);
  const h = document.createElement('h4');
  h.textContent = node.title;
  h.ondblclick = e => { e.stopPropagation(); editTitle(node, h); };
  const p = document.createElement('div');
  p.className = 'preview';
  p.innerHTML = linkify(node.body || node.children.map(c => labelOf(c)).join(' · '));
  const fade = document.createElement('div');
  fade.className = 'fade';
  el.append(h, p, fade);
  if (node.status) {
    const st = document.createElement('div');
    st.className = 'statuschip st-' + node.status;
    st.textContent = node.status;
    st.title = 'Click to change status';
    st.onclick = e => {
      e.stopPropagation();
      setNodeStatus(node, STATUSES[(STATUSES.indexOf(node.status) + 1) % STATUSES.length]);
    };
    el.appendChild(st);
  }
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
  if (layout.grid) {
    const mkPlus = (cls, title, at) => {
      const gp = document.createElement('div');
      gp.className = 'gplus' + cls;
      gp.innerHTML = icon('plus', 12);
      gp.title = 'Add section here';
      gp.onclick = e => {
        e.stopPropagation();
        const par = focus();
        addChild(par, at(par));
      };
      el.appendChild(gp);
    };
    mkPlus('', 'Add section after this one', par => par.children.indexOf(node) + 1);
    if (focus().children[0] === node) mkPlus(' left', 'Add section before this one', () => 0);
  }
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
  if (parent.children.length === 0) {
    // the only affordance on an empty section - make it unmissable
    plus.classList.add('big');
    plus.textContent = '+ Add a ' + (currentTag ? 'section' : LADDER[0].toLowerCase()) + ' inside';
    g.classList.add('bigwrap');
  } else {
    plus.innerHTML = icon('plus', 12);
  }
  plus.title = 'Add section here';
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
    // offer "up a level" whenever the current view has a parent to go to
    if (tree && path.length > 1) document.getElementById('upzone').hidden = false;
  });
  el.addEventListener('dragend', () => {
    dragNode = null;
    document.getElementById('upzone').hidden = true;
  });
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

// mid-drag target: drop to move the section up beside its current container
makeDropTarget(document.getElementById('upzone'), () => {
  const gp = path[path.length - 2];
  const idx = gp.children.indexOf(focus()) + 1;
  return [gp, idx];
});

// deletion: trivial sections get a confirm; anything substantial requires
// typing DELETE so a stray click can't erase real writing
async function deleteNode(node) {
  const wc = wordCount(node);
  const label = labelOf(node);
  const empty = !wc && !node.children.length && !(node.notes || '').trim();
  if (empty) {
    // nothing to lose, and undo covers regret
  } else if (wc >= 100 || node.children.length) {
    const detail = `${label}${node.children.length ? ` and its ${node.children.length} section${node.children.length > 1 ? 's' : ''}` : ''} — ${wc.toLocaleString()} words`;
    if (await appPrompt(`This permanently deletes ${detail}.\n\nType DELETE to confirm:`) !== 'DELETE') return;
  } else if (!(await appConfirm(`Delete ${label}?`))) return;
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

const defaultPrefs = () => ({
  newTemplate: 'novel',
  customDepth: 3,
  customCount: 3,
  goal: 0,
  historyDays: 30,
  exportFormat: 'docx',
  exportSep: '* * *',
  exportDir: '',
  exportSmf: true,
  exportTitlepage: true,
  author: '',
  byline: '',
  contact: [],
});
let prefs = { ...defaultPrefs(), ...JSON.parse(localStorage.getItem('writer-prefs') || '{}') };
function savePrefs() { localStorage.setItem('writer-prefs', JSON.stringify(prefs)); }
function countsForTemplate(tpl, depth, count) {
  return tpl === 'custom'
    ? Array(Math.min(5, +depth || 1)).fill(Math.min(8, +count || 1))
    : (TEMPLATES[tpl] || TEMPLATES.novel);
}

// ---- text settings (persisted) ----
const SETTINGS = { measure: { unit: 'px', def: 720 }, fsize: { unit: 'px', def: 17 }, lheight: { unit: '', def: 1.7 }, panelw: { unit: 'px', def: 420 }, gridw: { unit: 'px', def: 180 }, pfsize: { unit: 'px', def: 14 } };
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
function gotoSibling(dir) {
  if (path.length < 2) return;
  const par = path[path.length - 2];
  const i = par.children.indexOf(path[path.length - 1]);
  const target = par.children[i + dir];
  if (!target) return;
  path[path.length - 1] = target;
  render();
}
$('prevbtn').onclick = () => gotoSibling(-1);
$('nextbtn').onclick = () => gotoSibling(1);
function addSibling(after) {
  if (path.length < 2) return;
  pushUndo();
  const par = path[path.length - 2];
  const i = par.children.indexOf(path[path.length - 1]);
  const node = { depth: par.depth + 1, title: '', body: '', notes: '', children: [] };
  par.children.splice(after ? i + 1 : i, 0, node);
  path[path.length - 1] = node;    // land in the fresh section, ready to write
  save(); render();
}
$('edgeleft').onclick = () => addSibling(false);
$('edgeright').onclick = () => addSibling(true);
$('view-row').onclick = () => { panelsHidden = false; layout.grid = false; saveLayout(); render(); };
$('view-grid').onclick = () => { panelsHidden = false; layout.grid = true; saveLayout(); render(); };
$('view-hidden').onclick = () => { panelsHidden = true; render(); };

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
  b.innerHTML = readonly ? icon('book', 15) + ' Read' : icon('pencil', 15) + ' Edit';
  b.title = readonly ? 'Reading mode - click to edit' : 'Editing mode - click for read-only';
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
  if (!$('statsview').hidden) { $('statsview').hidden = true; return; }
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
  replaceMode = false;
  $('replacerow').hidden = true;
  $('replaceinput').value = '';
}
function doSearch(q) {
  const ql = q.trim().toLowerCase();
  searchHits = [];
  searchSel = 0;
  if (ql && storyTree) {
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
// ---- find & replace ----
let replaceMode = false;
function openReplace() {
  replaceMode = true;
  $('replacerow').hidden = false;
  openSearch();
}
function doReplaceAll() {
  const q = $('searchinput').value.trim();
  const rep = $('replaceinput').value;
  if (!storyTree) return;
  if (!q) { setStatus('Type what to find in the search box first', true); $('searchinput').focus(); return; }
  const re = new RegExp(escRe(q), 'gi');
  pushUndo();
  let count = 0;
  const apply = (n, withTitle) => {
    for (const key of withTitle ? ['title', 'body', 'notes'] : ['body', 'notes']) {
      if (!n[key]) continue;
      const m = n[key].match(re);
      if (m) { count += m.length; n[key] = n[key].replace(re, rep); }
    }
    n.children.forEach(c => apply(c, true));
  };
  apply(storyTree, false);   // never rewrite the story's root title (it's the folder name)
  const wasTag = currentTag;
  if (wasTag) { currentTag = null; }   // make save() write the story, not the open tag page
  const wasTree = tree;
  tree = storyTree;
  save().finally(async () => {
    if (wasTag) { currentTag = wasTag; tree = wasTree; }
    // replacing a tagged name should carry the tag file along, or the tag
    // page and explorer keep the old name
    const newName = rep.trim();
    const tagHit = newName && [...tagIndex.keys()].find(n => n.toLowerCase() === q.toLowerCase());
    if (tagHit && !tagIndex.has(newName) &&
        await appConfirm(`"${tagHit}" is a tag. Rename the tag itself to "${newName}" too?`)) {
      const cat = tagIndex.get(tagHit);
      try {
        await FS.writeText(tagFilePath(cat, newName), await FS.readText(tagFilePath(cat, tagHit)));
        await FS.remove(tagFilePath(cat, tagHit));
        if (currentTag && currentTag.name === tagHit) { currentTag = null; currentTagPath = null; tree = storyTree; path = [storyTree]; }
        await refreshTags();
      } catch (err) { reportErr('tag rename failed: ' + (err.stack || err)); }
    }
    render();
    setStatus(`Replaced ${count} occurrence${count === 1 ? '' : 's'}`);
    if (count > 0) closeSearch();          // done — dismiss the bar
    else doSearch(q);                      // nothing matched: leave it up to adjust
  });
}
$('replaceall').onmousedown = e => {
  e.preventDefault();
  setStatus('Replace: clicked');
  try { doReplaceAll(); } catch (err) { reportErr('replace failed: ' + (err.stack || err)); }
};
$('replaceinput').onkeydown = e => {
  if (e.key === 'Enter') { e.preventDefault(); doReplaceAll(); }
  else if (e.key === 'Escape') closeSearch();
};

$('searchinput').onfocus = openSearch;
$('searchinput').onblur = () => {
  setTimeout(() => {
    const a = document.activeElement;
    if (!a || !a.closest || !a.closest('#topsearch')) { $('searchdrop').hidden = true; }
  }, 0);
};
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
  if (!(e.ctrlKey || e.metaKey)) return;
  const k = e.key.toLowerCase();
  if (k === 'p' || k === 'f') { e.preventDefault(); replaceMode = false; $('replacerow').hidden = true; openSearch(); }
  else if (k === 'h') { e.preventDefault(); openReplace(); }
  else if (k === 's') {                       // autosave makes this a no-op, but the reflex deserves reassurance
    e.preventDefault();
    if (tree) { clearTimeout(saveTimer); save(); }
  }
  else if ((k === 'b' || k === 'i') && document.activeElement === $('focusbody') && !readonly) {
    e.preventDefault();
    const sel = getSelection().toString();
    if (sel && !sel.includes('\n')) {
      const m = k === 'b' ? '**' : '*';
      document.execCommand('insertText', false, m + sel + m);   // fires input -> body syncs
    }
  }
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
      const ren = document.createElement('span');
      ren.className = 'rowdel rowren';
      ren.innerHTML = icon('pencil', 11);
      ren.title = 'Rename tag';
      ren.onclick = e => { e.stopPropagation(); renameTag(cat, name); };
      const del = document.createElement('span');
      del.className = 'rowdel';
      del.textContent = '✕';
      del.title = 'Delete tag';
      del.onclick = e => { e.stopPropagation(); deleteTag(cat, name); };
      div.append(label, ren, del);
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
    cat = await appPrompt(`New tag "${name}" — category? (character, location, item, …)`, 'character');
    if (!cat) return;
    cat = cat.trim().toLowerCase();
  }
  openTag(cat, name);
}
// tags themselves are created from the writing (highlight -> right-click);
// the explorer builds the scaffolding: categories and their templates
$('newtag').onclick = async () => {
  const cat = await appPrompt('New category name (character, location, item, …):');
  if (!cat) return;
  await openTemplate(cat.trim().toLowerCase());
  await refreshTags();
};

async function renameTag(cat, oldName) {
  const newName = ((await appPrompt(`Rename tag "${oldName}" to:`, oldName)) || '').trim();
  if (!newName || newName === oldName) return;
  if (tagIndex.has(newName)) { await appAlert('A tag with that name already exists.'); return; }
  const oldP = tagFilePath(cat, oldName), newP = tagFilePath(cat, newName);
  try {
    await FS.writeText(newP, await FS.readText(oldP));
    await FS.remove(oldP);
  } catch (e) { await appAlert('Rename failed: ' + e.message); return; }
  // [[bracket]] references always follow the rename
  const reB = new RegExp('\\[\\[' + escRe(oldName) + '\\]\\]', 'g');
  (function wb(n) {
    n.body = n.body.replace(reB, '[[' + newName + ']]');
    n.notes = n.notes.replace(reB, '[[' + newName + ']]');
    n.children.forEach(wb);
  })(storyTree);
  // plain prose mentions only on request — that's a real find & replace
  const reW = new RegExp('\\b' + escRe(oldName) + '\\b', 'g');
  let plain = 0;
  (function count(n) { plain += (n.body.match(reW) || []).length; n.children.forEach(count); })(storyTree);
  if (plain && await appConfirm(`Also replace ${plain} plain mention${plain === 1 ? '' : 's'} of "${oldName}" with "${newName}" in the story text?`)) {
    pushUndo();
    (function rw(n) { n.body = n.body.replace(reW, newName); n.children.forEach(rw); })(storyTree);
  }
  const viewingIt = currentTag && currentTag.name === oldName;
  const wasTag = currentTag, wasTree = tree;
  currentTag = null; tree = storyTree;          // save the story, not the open tag page
  await save();
  if (wasTag && !viewingIt) { currentTag = wasTag; tree = wasTree; }
  await refreshTags();
  if (viewingIt) await openTag(cat, newName); else render();
}

async function deleteTag(cat, name) {
  if (!(await appConfirm(`Delete tag "${name}"?`))) return;
  try {
    const p = tagFilePath(cat, name);
    const text = await FS.readText(p);
    if (text.trim()) await trashFile(`tag ${cat} ${name}.md`, text);
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
function placeCtx(menu, x, y) {
  document.body.appendChild(menu);
  const r = menu.getBoundingClientRect();
  menu.style.left = Math.min(x, innerWidth - r.width - 8) + 'px';
  menu.style.top = Math.min(y, innerHeight - r.height - 8) + 'px';
}
function showSectionMenu(node, x, y) {
  closeCtx();
  const menu = document.createElement('div');
  menu.id = 'ctxmenu';
  const head = document.createElement('div');
  head.className = 'head';
  head.textContent = labelOf(node);
  menu.appendChild(head);
  for (const s of STATUSES) {
    const it = document.createElement('div');
    it.className = 'item';
    it.style.textTransform = 'none';
    it.textContent = ((node.status || '') === s ? '* ' : '  ') + (s || 'no status');
    it.onclick = () => { closeCtx(); setNodeStatus(node, s); };
    menu.appendChild(it);
  }
  const nn = document.createElement('div');
  nn.className = 'item';
  nn.style.textTransform = 'none';
  nn.textContent = node.nonum ? 'Include in numbering' : 'Exclude from numbering';
  nn.onclick = () => { closeCtx(); node.nonum = !node.nonum; save(); render(); };
  menu.appendChild(nn);
  const dl = document.createElement('div');
  dl.className = 'item';
  dl.style.textTransform = 'none';
  dl.textContent = 'Delete...';
  dl.onclick = () => { closeCtx(); deleteNode(node); };
  menu.appendChild(dl);
  placeCtx(menu, x, y);
}
document.addEventListener('contextmenu', e => {
  closeCtx();
  if (!tree || !storyPath) return;
  const sel = document.getSelection();
  const name = sel.toString().trim();

  // right-click on a panel (no text selected): section menu
  const panelHit = !name && e.target.closest('.panel');
  if (panelHit && e.target.closest('#panels')) {
    e.preventDefault();
    showSectionMenu(panelHit._node, e.clientX, e.clientY);
    return;
  }
  if (false && panelHit && e.target.closest('#panels')) {
    const node = panelHit._node;
    e.preventDefault();
    const menu = document.createElement('div');
    menu.id = 'ctxmenu';
    const head = document.createElement('div');
    head.className = 'head';
    head.textContent = labelOf(node);
    menu.appendChild(head);
    for (const s of STATUSES) {
      const it = document.createElement('div');
      it.className = 'item';
      it.style.textTransform = 'none';
      it.textContent = ((node.status || '') === s ? '● ' : '   ') + (s || 'no status');
      it.onclick = () => { closeCtx(); setNodeStatus(node, s); };
      menu.appendChild(it);
    }
    const nn = document.createElement('div');
    nn.className = 'item';
    nn.style.textTransform = 'none';
    nn.textContent = node.nonum ? 'Include in numbering' : 'Exclude from numbering (prologue etc.)';
    nn.onclick = () => { closeCtx(); node.nonum = !node.nonum; save(); render(); };
    menu.appendChild(nn);
    const dl = document.createElement('div');
    dl.className = 'item';
    dl.style.textTransform = 'none';
    dl.textContent = 'Delete…';
    dl.onclick = () => { closeCtx(); deleteNode(node); };
    menu.appendChild(dl);
    document.body.appendChild(menu);
    const r = menu.getBoundingClientRect();
    menu.style.left = Math.min(e.clientX, innerWidth - r.width - 8) + 'px';
    menu.style.top = Math.min(e.clientY, innerHeight - r.height - 8) + 'px';
    return;
  }

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

  const short = name.length > 24 ? name.slice(0, 24) + '...' : name;
  addItem(`Find "${short}"`, () => {
    openSearch();
    $('searchinput').value = name;
    doSearch(name);
  });
  addItem(`Replace "${short}"...`, () => {
    openReplace();
    $('searchinput').value = name;
    doSearch(name);
    $('replaceinput').focus();
  });
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
    nw.onclick = async () => {
      closeCtx();
      const cat = await appPrompt('Category name:');
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
  catch { appAlert('Select plain text only (a selection can\'t cut across a link or another note).'); return; }
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

// ---- writing goals & session stats ----
const dateKey = (d = new Date()) =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
let dayStart = null;
function trackDay() {
  if (!storyTree || !storyName) { dayStart = null; return; }
  const key = 'writer-day:' + storyName;
  const today = dateKey();
  const wcNow = wordCount(storyTree);
  let d = JSON.parse(localStorage.getItem(key) || 'null');
  if (!d || d.date !== today) d = { date: today, start: wcNow };
  localStorage.setItem(key, JSON.stringify(d));
  dayStart = d.start;
}
function saveHistoryPoint(story, day, words) {
  const h = JSON.parse(localStorage.getItem('writer-history') || '{}');
  h[story] = h[story] || {};
  if (words) h[story][day] = words; else delete h[story][day];
  localStorage.setItem('writer-history', JSON.stringify(h));
}
function updateGoal() {
  const el = $('goalpill');
  if (!storyTree || dayStart === null) { el.hidden = true; return; }
  const delta = wordCount(storyTree) - dayStart;
  saveHistoryPoint(storyName, dateKey(), delta);
  const goal = +prefs.goal || 0;
  el.hidden = false;
  el.textContent = `${delta >= 0 ? '+' : ''}${delta.toLocaleString()} today` +
    (goal ? ` · ${Math.min(999, Math.round(100 * Math.max(0, delta) / goal))}% of ${goal.toLocaleString()}` : '');
  el.classList.toggle('met', goal > 0 && delta >= goal);
}

// ---- statistics view ----
async function statsTree(name) {
  if (name === storyName && storyTree) return storyTree;   // open story: use live tree
  const sp = FS.join(libPath, name);
  const sf = FS.join(sp, 'story.md');
  if (!(await FS.exists(sf).catch(() => false))) return null;
  const root = parse(splitFrontmatter(await FS.readText(sf)).body, name);
  const chDir = FS.join(sp, 'chapters');
  if (await FS.exists(chDir).catch(() => false)) {
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
          const c = parse(await FS.readText(FS.join(chDir, fname)), '', { unwrap: false });
          const ch = { depth: 0, title: '', body: c.body, notes: c.notes, children: c.children };
          (function fix(k, d) { k.depth = d; k.children.forEach(x => fix(x, d + 1)); })(ch, node.depth + 1);
          loaded.push(ch);
        } catch {}
      }
      node.children = [...loaded, ...node.children];
    }
  }
  return root;
}
function kindCounts(root) {
  const H = treeHeight(root);
  const counts = {};
  (function walk(n) {
    for (const c of n.children) {
      const kind = LADDER[Math.min(Math.max(H - c.depth, 0), LADDER.length - 1)];
      counts[kind] = (counts[kind] || 0) + 1;
      walk(c);
    }
  })(root);
  return counts;
}
function streaks(days) {   // days: map dateKey -> words
  const worked = new Set(Object.keys(days).filter(k => days[k] > 0));
  let longest = 0;
  for (const k of worked) {
    const prev = new Date(k); prev.setDate(prev.getDate() - 1);
    if (worked.has(dateKey(prev))) continue;   // not a streak start
    let len = 0;
    const d = new Date(k);
    while (worked.has(dateKey(d))) { len++; d.setDate(d.getDate() + 1); }
    longest = Math.max(longest, len);
  }
  let current = 0;
  const d = new Date();
  if (!worked.has(dateKey(d))) d.setDate(d.getDate() - 1);   // today not written yet doesn't break it
  while (worked.has(dateKey(d))) { current++; d.setDate(d.getDate() - 1); }
  return { current, longest };
}
function calendarEl(days) {
  const cal = document.createElement('div');
  cal.className = 'cal';
  const end = new Date();
  const start = new Date();
  start.setDate(start.getDate() - 364 - end.getDay());
  const d = new Date(start);
  let col = null;
  while (d <= end) {
    if (d.getDay() === 0 || !col) { col = document.createElement('div'); col.className = 'calcol'; cal.appendChild(col); }
    const cell = document.createElement('div');
    const w = Math.max(0, days[dateKey(d)] || 0);
    const cls = w >= 800 ? ' c4' : w >= 400 ? ' c3' : w >= 150 ? ' c2' : w > 0 ? ' c1' : '';
    cell.className = 'calcell' + cls;
    cell.title = `${dateKey(d)}: ${w.toLocaleString()} words`;
    col.appendChild(cell);
    d.setDate(d.getDate() + 1);
  }
  return cal;
}
async function renderStats() {
  const selName = $('statssel').value;
  const body = $('statsbody');
  body.innerHTML = '<p style="color:var(--dim)">Crunching…</p>';
  const names = [];
  if (selName === '*') {
    for (const e of await FS.readDir(libPath)) if (e.isDir && !e.name.startsWith('.')) names.push(e.name);
  } else names.push(selName);
  const history = JSON.parse(localStorage.getItem('writer-history') || '{}');
  const days = {};
  for (const n of names) for (const [day, w] of Object.entries(history[n] || {})) days[day] = (days[day] || 0) + w;
  let words = 0;
  const counts = {};
  for (const n of names) {
    const t = await statsTree(n);
    if (!t) continue;
    words += wordCount(t);
    for (const [k, v] of Object.entries(kindCounts(t))) counts[k] = (counts[k] || 0) + v;
  }
  const st = streaks(days);
  const daysWorked = Object.values(days).filter(w => w > 0).length;
  const wordsAllTime = Object.values(days).reduce((a, w) => a + Math.max(0, w), 0);

  body.innerHTML = '';
  const cards = document.createElement('div');
  cards.className = 'statcards';
  const card = (num, label) => {
    const c = document.createElement('div');
    c.className = 'statcard';
    c.innerHTML = `<div class="statnum">${num}</div><div class="statlabel">${esc(label)}</div>`;
    cards.appendChild(c);
  };
  card(words.toLocaleString(), 'words');
  for (const kind of [...LADDER].reverse()) if (counts[kind]) card(counts[kind].toLocaleString(), kind + (counts[kind] === 1 ? '' : 's'));
  card(daysWorked.toLocaleString(), 'days worked');
  card(st.current.toLocaleString(), 'current streak');
  card(st.longest.toLocaleString(), 'longest streak');
  card(wordsAllTime.toLocaleString(), 'words tracked');
  body.appendChild(cards);

  const calHead = document.createElement('h3');
  calHead.textContent = 'Last 12 months';
  calHead.className = 'calhead';
  body.appendChild(calHead);
  body.appendChild(calendarEl(days));
  const note = document.createElement('p');
  note.className = 'calnote';
  note.textContent = 'Daily words are tracked from the day this feature was added; earlier writing shows in the totals but not the calendar.';
  body.appendChild(note);
}
async function openStats() {
  if (!libPath) return;
  const sel = $('statssel');
  sel.innerHTML = '';
  const all = document.createElement('option');
  all.value = '*';
  all.textContent = 'All stories';
  sel.appendChild(all);
  for (const e of await FS.readDir(libPath)) {
    if (!e.isDir || e.name.startsWith('.')) continue;
    const o = document.createElement('option');
    o.value = e.name;
    o.textContent = e.name;
    sel.appendChild(o);
  }
  sel.value = storyName && [...sel.options].some(o => o.value === storyName) ? storyName : '*';
  $('statsview').hidden = false;
  renderStats();
}
$('statsbtn').onclick = () => { if (!$('statsview').hidden) $('statsview').hidden = true; else openStats(); };
$('statssel').onchange = renderStats;

// ---- library / stories ----
async function openLibrary(path) {
  libPath = path;
  if (FS.native) localStorage.setItem('writer-lib', path);
  $('newstory').hidden = false;
  await listStories();
  setStatus('');
  render();
}
$('samplebtn').onclick = async () => {
  if (await FS.exists('sample-library').catch(() => false)) await openLibrary('sample-library');
  else appAlert('Sample library was not found next to the app.');
};
$('openbtn').onclick = async () => {
  const p = await FS.pickFolder();
  if (p) await openLibrary(p);
};
$('hintopen').onclick = () => $('openbtn').click();
$('hintsample').onclick = () => $('samplebtn').click();
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
      if (await appPrompt(`This permanently deletes the story "${entry.name}" and ALL its files.\n\nType the story name to confirm:`) !== entry.name) return;
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
  trackDay();
  await refreshTags();
  render();
  $('historybtn').hidden = false;
  pruneHistory();
  setStatus('');
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
  $('historybtn').hidden = true;

  setStatus('');
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
$('newstory').onclick = () => {
  $('ns-name').value = '';
  $('ns-template').value = prefs.newTemplate;
  $('ns-depth').value = prefs.customDepth;
  $('ns-count').value = prefs.customCount;
  $('ns-custom').hidden = $('ns-template').value !== 'custom';
  $('newstorydlg').showModal();
};
$('hintnew').onclick = () => $('newstory').click();
$('ns-template').onchange = () => { $('ns-custom').hidden = $('ns-template').value !== 'custom'; };
const TEMPLATES = { blank: [], short: [3], novel: [2, 1, 1], epic: [2, 2, 1, 1] };
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
  const counts = countsForTemplate(tpl, $('ns-depth').value, $('ns-count').value);
  $('newstorydlg').close();
  await FS.mkdir(FS.join(libPath, name));
  await listStories();
  await openStory(name);
  if (!storyMeta.author && !storyMeta.byline && !storyMeta.contact) {
    storyMeta = { author: prefs.author, byline: prefs.byline, contact: prefs.contact };
  }
  if (counts.length && !tree.children.length) {
    tree.children = buildLevels(counts);
    await save();
    render();
  } else {
    await save();
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
function exportDir() {
  return FS.native ? (prefs.exportDir || storyPath || libPath) : '';
}
function showExportDir() {
  const value = FS.native ? (exportDir() || '') : 'Browser downloads';
  if ($('ex-dir')) $('ex-dir').value = value;
  if ($('set-export-dir')) $('set-export-dir').value = prefs.exportDir || '';
}
async function chooseExportDir() {
  if (!FS.native) { appAlert('In the browser version, exports use the browser downloads folder.'); return; }
  const p = await FS.pickFolder();
  if (!p) return;
  prefs.exportDir = p;
  savePrefs();
  showExportDir();
}
async function writeExport(name, data, type) {
  try {
    if (!FS.native) { download(name, data, type); return; }
    const dir = exportDir();
    if (!dir) { download(name, data, type); return; }
    const path = FS.join(dir, name);
    if (data instanceof Uint8Array) await FS.writeBytes(path, data);
    else await FS.writeText(path, data);
    setStatus('Exported to ' + path);
  } catch (e) {
    console.error(e);
    setStatus('Export failed - check console', true);
    throw e;
  }
}
$('exportbtn').onclick = () => {
  if (!storyTree) return;
  $('ex-format').value = prefs.exportFormat;
  $('ex-sep').value = prefs.exportSep;
  $('ex-smf').checked = !!prefs.exportSmf;
  $('ex-titlepage').checked = !!prefs.exportTitlepage;
  showExportDir();
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
$('exportform').onsubmit = async e => {
  e.preventDefault();
  prefs.exportFormat = $('ex-format').value;
  prefs.exportSep = $('ex-sep').value;
  prefs.exportSmf = $('ex-smf').checked;
  prefs.exportTitlepage = $('ex-titlepage').checked;
  savePrefs();
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
    await writeExport(safe + '.docx', bytes, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  } else if ($('ex-format').value === 'html') {
    await writeExport(safe + '.html', mdToHtml(md, title, tp), 'text/html');
  } else {
    const head = tp
      ? (tp.contact.length ? tp.contact.join('\n') + '\n\n' : '') +
        `approx. ${tp.approx} words\n\n# ${title}\n\n${tp.by ? `by ${tp.by}\n\n` : ''}---\n\n`
      : `# ${title}\n\n`;
    await writeExport(safe + '.md', head + md + '\n', 'text/markdown');
  }
};
$('ex-dir-btn').onclick = chooseExportDir;

// ---- settings ----
document.querySelectorAll('#settingsnav button').forEach(btn => {
  btn.onclick = () => {
    const tab = btn.dataset.settingsTab;
    document.querySelectorAll('#settingsnav button').forEach(b => b.classList.toggle('active', b === btn));
    document.querySelectorAll('.settings-page').forEach(p => p.classList.toggle('active', p.dataset.settingsPage === tab));
  };
});
$('set-new-template').onchange = () => { $('set-custom-row').hidden = $('set-new-template').value !== 'custom'; };
$('set-export-dir-btn').onclick = chooseExportDir;
$('set-export-dir-clear').onclick = () => {
  prefs.exportDir = '';
  savePrefs();
  showExportDir();
};
$('settingsbtn').onclick = () => {
  document.querySelector('#settingsnav button[data-settings-tab="files"]').click();
  $('set-libpath').value = libPath || '';
  $('set-storypath').value = storyPath || '';
  $('set-new-template').value = prefs.newTemplate;
  $('set-depth').value = prefs.customDepth;
  $('set-count').value = prefs.customCount;
  $('set-custom-row').hidden = prefs.newTemplate !== 'custom';
  $('set-export-format').value = prefs.exportFormat;
  $('set-export-sep').value = prefs.exportSep;
  showExportDir();
  $('set-export-smf').checked = !!prefs.exportSmf;
  $('set-export-titlepage').checked = !!prefs.exportTitlepage;
  $('set-goal').value = +prefs.goal || 0;
  $('set-historydays').value = +prefs.historyDays || 30;
  $('set-author').value = prefs.author || '';
  $('set-byline').value = prefs.byline || '';
  $('set-contact').value = [].concat(prefs.contact || []).join('\n');
  $('ladderinput').value = LADDER.join(', ');
  $('settingsdlg').showModal();
};
$('settingsform').onsubmit = () => {
  const names = $('ladderinput').value.split(',').map(s => s.trim()).filter(Boolean);
  if (names.length) {
    LADDER = names;
    localStorage.setItem('writer-ladder', JSON.stringify(LADDER));
  }
  prefs = {
    ...prefs,
    newTemplate: $('set-new-template').value,
    customDepth: Math.min(5, +$('set-depth').value || 1),
    customCount: Math.min(8, +$('set-count').value || 1),
    goal: Math.max(0, +$('set-goal').value || 0),
    historyDays: Math.max(1, +$('set-historydays').value || 30),
    exportFormat: $('set-export-format').value,
    exportSep: $('set-export-sep').value,
    exportDir: prefs.exportDir || '',
    exportSmf: $('set-export-smf').checked,
    exportTitlepage: $('set-export-titlepage').checked,
    author: $('set-author').value.trim(),
    byline: $('set-byline').value.trim(),
    contact: $('set-contact').value.split('\n').map(s => s.trim()).filter(Boolean),
  };
  savePrefs();
  render();
};

// ---- startup: reopen the last library automatically (native only) ----
(async function init() {
  if (!FS.native) return;
  const saved = localStorage.getItem('writer-lib');
  if (saved && await FS.exists(saved).catch(() => false)) await openLibrary(saved);
})();
