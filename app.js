/* HTML 嵌入 —— PowerPoint 内容加载项
 * HTML 内容保存在 Office.context.document.settings 里，也就是存进 .pptx 文件。
 * 每个插入的框有自己独立的 settings，所以一页可以放多个。
 */
(function () {
  'use strict';

  const KEY_HTML = 'htmlContent';
  const KEY_NAME = 'htmlName';
  const KEY_W = 'designWidth';
  const WARN_BYTES = 2 * 1024 * 1024; // Office 对加载项设置大小有限制，超过 2MB 给出提示

  const $ = (id) => document.getElementById(id);
  const els = {
    picker: $('picker'), viewer: $('viewer'), frame: $('frame'), status: $('status'),
    docName: $('docName'), fileInput: $('fileInput'), folderInput: $('folderInput'),
    pasteBox: $('pasteBox'), pasteArea: $('pasteArea'), btnBack: $('btnBack'),
  };

  // ---------- 存储：PowerPoint 中用文档设置，普通浏览器里退回 localStorage 方便测试 ----------
  let store = null;

  function makeOfficeStore() {
    const s = Office.context.document.settings;
    return {
      get: (k) => s.get(k),
      set: (k, v) => s.set(k, v),
      save: () => new Promise((resolve, reject) => {
        s.saveAsync((r) => r.status === Office.AsyncResultStatus.Succeeded
          ? resolve() : reject(r.error || new Error('保存失败')));
      }),
    };
  }

  function makeLocalStore() {
    const mem = {};
    const ls = (() => { try { return window.localStorage; } catch (e) { return null; } })();
    return {
      get: (k) => { try { const v = ls && ls.getItem('pptHtml:' + k); return v == null ? mem[k] : v; } catch (e) { return mem[k]; } },
      set: (k, v) => {
        mem[k] = v;
        try { if (!ls) return; if (v == null) ls.removeItem('pptHtml:' + k); else ls.setItem('pptHtml:' + k, v); } catch (e) { /* 忽略 */ }
      },
      save: () => Promise.resolve(),
    };
  }

  // ---------- 显示 ----------
  // 网页按固定宽度 designW 排版，再整体等比缩放到框的大小。
  // 这样不管框多大、编辑还是放映，看起来都一样。
  let designW = null;

  function viewSize() {
    const r = els.viewer.getBoundingClientRect();
    return { w: Math.max(r.width, 1), h: Math.max(r.height, 1) };
  }

  function layout() {
    const { w, h } = viewSize();
    const W = designW || w;
    const s = w / W;
    els.frame.style.width = W + 'px';
    els.frame.style.height = (h / s) + 'px';
    els.frame.style.transform = `scale(${s})`;
  }

  // 用 setTimeout 而不是 requestAnimationFrame：后台或隐藏时 rAF 可能不触发
  const settle = () => new Promise((resolve) => setTimeout(resolve, 60));

  function saveDesignWidth() {
    store.set(KEY_W, designW ? Math.round(designW) : null);
    store.save().catch(() => { /* 放映时可能无法保存，忽略 */ });
  }

  // 从框的实际宽度开始，逐步加大排版宽度，直到整页内容都装得下
  async function fitToContent() {
    let doc = null;
    try { doc = els.frame.contentDocument; } catch (e) { /* 无法访问时跳过 */ }
    if (!doc || !doc.documentElement) return;
    const { w } = viewSize();
    let W = w;
    for (let i = 0; i < 12; i++) {
      designW = W;
      layout();
      await settle();
      const viewH = parseFloat(els.frame.style.height);
      const contentH = Math.max(doc.documentElement.scrollHeight, doc.body ? doc.body.scrollHeight : 0);
      if (contentH <= viewH + 2 || W >= 4000) break;
      W = Math.min(4000, Math.ceil(W * Math.max(1.05, Math.sqrt(contentH / viewH))));
    }
    saveDesignWidth();
  }

  function zoom(factor) {
    const { w } = viewSize();
    designW = Math.max(200, Math.min(6000, (designW || w) / factor));
    layout();
    saveDesignWidth();
  }

  function showViewer(html, name) {
    els.picker.hidden = true;
    els.viewer.hidden = false;
    els.viewer.title = name || '';
    designW = Number(store.get(KEY_W)) || null;
    layout();
    els.frame.onload = () => { if (!designW) setTimeout(fitToContent, 400); };
    els.frame.srcdoc = html;
  }

  function showPicker(canGoBack) {
    els.viewer.hidden = true;
    els.picker.hidden = false;
    els.pasteBox.hidden = true;
    els.btnBack.hidden = !canGoBack;
    setStatus('');
  }

  function setStatus(msg, isError) {
    els.status.textContent = msg || '';
    els.status.classList.toggle('error', !!isError);
  }

  async function useHtml(html, name) {
    const bytes = new Blob([html]).size;
    if (bytes > WARN_BYTES) {
      setStatus(`内容有 ${(bytes / 1048576).toFixed(1)} MB，可能超过 PPT 能保存的上限，正在尝试…`);
    } else {
      setStatus('正在保存…');
    }
    store.set(KEY_HTML, html);
    store.set(KEY_NAME, name || '');
    store.set(KEY_W, null); // 新内容重新自动适应
    try {
      await store.save();
    } catch (err) {
      setStatus('保存到 PPT 失败：' + (err.message || err) + '。可以试试压缩图片或减小文件。', true);
      return;
    }
    showViewer(html, name);
  }

  // ---------- 读取文件 ----------
  const readText = (file) => file.text();
  const readDataUrl = (file) => new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(file);
  });

  async function handleSingleFile(file) {
    if (!file) return;
    setStatus('正在读取 ' + file.name + '…');
    const html = await readText(file);
    await useHtml(html, file.name);
  }

  async function handleFolder(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    setStatus('正在打包文件夹…');

    // 路径 -> File，去掉最外层文件夹名
    const map = new Map();
    for (const f of files) {
      const rel = (f.webkitRelativePath || f.name).split('/').slice(1).join('/') || f.name;
      map.set(rel, f);
    }
    const htmlPaths = [...map.keys()].filter((p) => /\.html?$/i.test(p))
      .sort((a, b) => a.split('/').length - b.split('/').length || a.localeCompare(b));
    if (!htmlPaths.length) { setStatus('文件夹里没有找到 .html 文件。', true); return; }
    const entry = htmlPaths.find((p) => /(^|\/)index\.html?$/i.test(p)) || htmlPaths[0];

    try {
      const { html, missing } = await inlineHtml(entry, map);
      const name = entry.split('/').pop();
      await useHtml(html, name);
      if (missing.length) console.warn('未找到的资源：', missing);
    } catch (err) {
      setStatus('打包失败：' + (err.message || err), true);
    }
  }

  // ---------- 把文件夹打包成单个 HTML ----------
  const isExternal = (u) => !u || /^(?:[a-z][a-z0-9+.-]*:|\/\/|#)/i.test(u.trim());

  function resolvePath(fromFile, ref) {
    const clean = ref.trim().split('#')[0].split('?')[0];
    const url = new URL(clean, 'https://local.invalid/' + fromFile);
    return decodeURIComponent(url.pathname.slice(1));
  }

  async function inlineHtml(entryPath, map) {
    const missing = [];
    const dataCache = new Map();

    async function toDataUrl(fromFile, ref) {
      if (isExternal(ref)) return null;
      const p = resolvePath(fromFile, ref);
      const f = map.get(p);
      if (!f) { missing.push(p); return null; }
      if (!dataCache.has(p)) dataCache.set(p, readDataUrl(f));
      return dataCache.get(p);
    }

    async function inlineCss(css, fromFile, depth) {
      // @import "x.css" / @import url(x.css)
      const importRe = /@import\s+(?:url\(\s*)?["']?([^"')\s;]+)["']?\s*\)?\s*([^;]*);/g;
      const imports = [...css.matchAll(importRe)];
      for (const m of imports) {
        if (isExternal(m[1]) || depth > 5) continue;
        const p = resolvePath(fromFile, m[1]);
        const f = map.get(p);
        if (!f) { missing.push(p); continue; }
        let sub = await inlineCss(await readText(f), p, depth + 1);
        if (m[2].trim()) sub = `@media ${m[2].trim()} {\n${sub}\n}`;
        css = css.replace(m[0], sub);
      }
      // url(...)
      const urlRe = /url\(\s*(['"]?)([^'")]+)\1\s*\)/g;
      const parts = [];
      let last = 0;
      for (const m of css.matchAll(urlRe)) {
        parts.push(css.slice(last, m.index));
        const data = await toDataUrl(fromFile, m[2]);
        parts.push(data ? `url("${data}")` : m[0]);
        last = m.index + m[0].length;
      }
      parts.push(css.slice(last));
      return parts.join('');
    }

    const source = await readText(map.get(entryPath));
    const doc = new DOMParser().parseFromString(source, 'text/html');

    // 外部样式表 -> <style>
    for (const link of [...doc.querySelectorAll('link[rel~="stylesheet"][href]')]) {
      const href = link.getAttribute('href');
      if (isExternal(href)) continue;
      const p = resolvePath(entryPath, href);
      const f = map.get(p);
      if (!f) { missing.push(p); continue; }
      const style = doc.createElement('style');
      if (link.media) style.media = link.media;
      style.textContent = await inlineCss(await readText(f), p, 0);
      link.replaceWith(style);
    }
    // <style> 里的 url()
    for (const style of doc.querySelectorAll('style')) {
      style.textContent = await inlineCss(style.textContent, entryPath, 0);
    }
    // style="" 属性里的 url()
    for (const el of doc.querySelectorAll('[style*="url("]')) {
      el.setAttribute('style', await inlineCss(el.getAttribute('style'), entryPath, 0));
    }
    // 外部脚本 -> 内联脚本
    for (const s of [...doc.querySelectorAll('script[src]')]) {
      const src = s.getAttribute('src');
      if (isExternal(src)) continue;
      const p = resolvePath(entryPath, src);
      const f = map.get(p);
      if (!f) { missing.push(p); continue; }
      const code = (await readText(f)).replace(/<\/script/gi, '<\\/script');
      s.removeAttribute('src');
      s.textContent = code;
    }
    // 图片、音视频等资源 -> data URL
    const attrTargets = [
      ['img[src]', 'src'], ['source[src]', 'src'], ['video[src]', 'src'], ['audio[src]', 'src'],
      ['video[poster]', 'poster'], ['track[src]', 'src'], ['embed[src]', 'src'], ['object[data]', 'data'],
      ['input[type="image"][src]', 'src'], ['link[rel~="icon"][href]', 'href'],
      ['image[href]', 'href'], ['use[href]', 'href'],
    ];
    for (const [sel, attr] of attrTargets) {
      for (const el of doc.querySelectorAll(sel)) {
        const val = el.getAttribute(attr);
        if (attr === 'href' && val && val.startsWith('#')) continue;
        const data = await toDataUrl(entryPath, val);
        if (data) el.setAttribute(attr, data);
      }
    }
    // srcset
    for (const el of doc.querySelectorAll('[srcset]')) {
      const items = el.getAttribute('srcset').split(',').map((s) => s.trim()).filter(Boolean);
      const out = [];
      for (const item of items) {
        const [u, ...rest] = item.split(/\s+/);
        const data = await toDataUrl(entryPath, u);
        out.push([data || u, ...rest].join(' '));
      }
      el.setAttribute('srcset', out.join(', '));
    }

    const doctype = /^\s*<!doctype/i.test(source) ? '<!DOCTYPE html>\n' : '';
    return { html: doctype + doc.documentElement.outerHTML, missing: [...new Set(missing)] };
  }

  // ---------- 放映时隐藏工具栏 ----------
  function watchView() {
    if (!(window.Office && Office.context && Office.context.document)) return;
    const apply = (view) => document.body.classList.toggle('slideshow', view === 'read');
    const doc = Office.context.document;
    if (doc.getActiveViewAsync) {
      doc.getActiveViewAsync((r) => { if (r.status === Office.AsyncResultStatus.Succeeded) apply(r.value); });
    }
    if (doc.addHandlerAsync && Office.EventType.ActiveViewChanged) {
      doc.addHandlerAsync(Office.EventType.ActiveViewChanged, (e) => apply(e.activeView));
    }
  }

  // ---------- 事件 ----------
  function bindUi() {
    $('btnFile').onclick = () => { els.fileInput.value = ''; els.fileInput.click(); };
    $('btnFolder').onclick = () => { els.folderInput.value = ''; els.folderInput.click(); };
    $('btnPaste').onclick = () => { els.pasteBox.hidden = false; els.pasteArea.focus(); };
    $('btnPasteCancel').onclick = () => { els.pasteBox.hidden = true; };
    $('btnPasteOk').onclick = () => {
      const html = els.pasteArea.value.trim();
      if (!html) { setStatus('请先粘贴 HTML 代码。', true); return; }
      useHtml(html, '粘贴的代码');
    };
    els.fileInput.onchange = () => handleSingleFile(els.fileInput.files[0]).catch((e) => setStatus(String(e), true));
    els.folderInput.onchange = () => handleFolder(els.folderInput.files).catch((e) => setStatus(String(e), true));

    $('btnReload').onclick = () => { const h = store.get(KEY_HTML); if (h) { els.frame.srcdoc = ''; els.frame.srcdoc = h; } };
    $('btnChange').onclick = () => showPicker(true);
    $('btnSmaller').onclick = () => zoom(1 / 1.15);
    $('btnBigger').onclick = () => zoom(1.15);
    $('btnFit').onclick = () => fitToContent();
    if (window.ResizeObserver) new ResizeObserver(() => { if (!els.viewer.hidden) layout(); }).observe(els.viewer);
    window.addEventListener('resize', () => { if (!els.viewer.hidden) layout(); });
    // 兜底：PowerPoint 切换放映时不一定触发 resize 事件，定时检查框的尺寸
    let lastSize = '';
    setInterval(() => {
      if (els.viewer.hidden) return;
      const { w, h } = viewSize();
      const size = w + 'x' + h;
      if (size !== lastSize) { lastSize = size; layout(); }
    }, 500);
    els.btnBack.onclick = () => { const h = store.get(KEY_HTML); if (h) showViewer(h, store.get(KEY_NAME)); };

    // 拖放文件
    const pz = els.picker;
    ['dragenter', 'dragover'].forEach((t) => pz.addEventListener(t, (e) => { e.preventDefault(); pz.classList.add('drag'); }));
    ['dragleave', 'drop'].forEach((t) => pz.addEventListener(t, (e) => { e.preventDefault(); pz.classList.remove('drag'); }));
    pz.addEventListener('drop', (e) => {
      const f = e.dataTransfer && e.dataTransfer.files[0];
      if (f && /\.html?$/i.test(f.name)) handleSingleFile(f).catch((err) => setStatus(String(err), true));
      else setStatus('请拖入 .html 文件。', true);
    });
  }

  function start(inOffice) {
    store = inOffice ? makeOfficeStore() : makeLocalStore();
    bindUi();
    watchView();
    const html = store.get(KEY_HTML);
    if (html) showViewer(html, store.get(KEY_NAME));
    else showPicker(false);
  }

  // ---------- 离线缓存 ----------
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch((e) => console.warn('Service worker 注册失败', e));
  }

  // ---------- 启动 ----------
  let started = false;
  const go = (inOffice) => { if (!started) { started = true; start(inOffice); } };
  if (window.Office && Office.onReady) {
    Office.onReady((info) => go(!!(info && info.host && Office.context && Office.context.document && Office.context.document.settings)));
  } else {
    // office.js 没加载到（例如离线且没有缓存）：显示界面，并提示联网后重开
    go(false);
    setStatus('Office 组件没有加载成功，请联网后重新打开这个 PPT。', true);
  }
})();
