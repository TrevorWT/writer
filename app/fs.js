// Filesystem abstraction: path-string API with two backends.
// - Tauri (native): real paths, dialogs, file watching, remembered library.
// - Web (fallback for browser dev): File System Access API behind path tokens.

export let FS = null;

export function initFS() {
  FS = window.__TAURI__ ? tauriBackend() : webBackend();
  return FS;
}

function tauriBackend() {
  const t = window.__TAURI__;
  return {
    native: true,
    join: (...parts) => parts.join('/'),
    async pickFolder() {
      return await t.dialog.open({ directory: true });
    },
    readText: p => t.fs.readTextFile(p),
    writeText: (p, text) => t.fs.writeTextFile(p, text),
    async readDir(p) {
      return (await t.fs.readDir(p)).map(e => ({ name: e.name, isDir: e.isDirectory }));
    },
    mkdir: p => t.fs.mkdir(p, { recursive: true }).catch(() => {}),
    remove: p => t.fs.remove(p),
    removeDir: p => t.fs.remove(p, { recursive: true }),
    exists: p => t.fs.exists(p),
    async watch(p, cb) {
      try { return await t.fs.watch(p, cb, { recursive: true, delayMs: 800 }); }
      catch (e) { console.warn('watch failed', e); return () => {}; }
    },
  };
}

function webBackend() {
  const roots = new Map();
  let n = 0;
  // resolve a path like "root1/story/chapters/01.md" to its parent dir handle
  async function walkTo(path, create) {
    const parts = path.split('/').filter(Boolean);
    let h = roots.get(parts[0]);
    if (!h) throw new Error('unknown root: ' + path);
    for (let i = 1; i < parts.length - 1; i++) h = await h.getDirectoryHandle(parts[i], { create });
    return { dir: h, leaf: parts[parts.length - 1], rootOnly: parts.length === 1 };
  }
  async function dirHandle(path, create) {
    const parts = path.split('/').filter(Boolean);
    let h = roots.get(parts[0]);
    if (!h) throw new Error('unknown root: ' + path);
    for (let i = 1; i < parts.length; i++) h = await h.getDirectoryHandle(parts[i], { create });
    return h;
  }
  return {
    native: false,
    join: (...parts) => parts.join('/'),
    async pickFolder() {
      const h = await showDirectoryPicker({ mode: 'readwrite' });
      const token = 'library-' + (++n);
      roots.set(token, h);
      return token;
    },
    async readText(p) {
      const { dir, leaf } = await walkTo(p, false);
      return await (await (await dir.getFileHandle(leaf)).getFile()).text();
    },
    async writeText(p, text) {
      const { dir, leaf } = await walkTo(p, true);
      const w = await (await dir.getFileHandle(leaf, { create: true })).createWritable();
      await w.write(text);
      await w.close();
    },
    async readDir(p) {
      const h = await dirHandle(p, false);
      const out = [];
      for await (const e of h.values()) out.push({ name: e.name, isDir: e.kind === 'directory' });
      return out;
    },
    mkdir: p => dirHandle(p, true).then(() => {}),
    async remove(p) {
      const { dir, leaf } = await walkTo(p, false);
      await dir.removeEntry(leaf);
    },
    async removeDir(p) {
      const { dir, leaf } = await walkTo(p, false);
      await dir.removeEntry(leaf, { recursive: true });
    },
    async exists(p) {
      try {
        const { dir, leaf } = await walkTo(p, false);
        try { await dir.getFileHandle(leaf); return true; } catch {}
        await dir.getDirectoryHandle(leaf);
        return true;
      } catch { return false; }
    },
    watch: () => () => {},   // no watching in the browser fallback
  };
}
