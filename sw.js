/* 离线缓存：联网打开过一次之后，插件本身和 Office 组件都从缓存加载。
 * 策略：先用缓存（快、离线可用），同时在后台联网更新缓存。
 */
const CACHE = 'ppt-html-embed-v1';
const OFFICE_JS = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
const CORE = ['./content.html', './app.js', './style.css', './icon-32.png', './icon-80.png'];

self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    await cache.addAll(CORE);
    try { await cache.add(new Request(OFFICE_JS, { mode: 'no-cors' })); } catch (e) { /* 下次再缓存 */ }
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    for (const key of await caches.keys()) if (key !== CACHE) await caches.delete(key);
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (!/^https?:$/.test(url.protocol)) return;

  // PowerPoint 打开插件时会在网址后面加参数（?_host_Info=...），页面请求要忽略参数来匹配缓存
  const isPage = req.mode === 'navigate' || req.destination === 'document';
  const key = isPage ? new Request(url.origin + url.pathname) : req;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const cached = await cache.match(key, { ignoreSearch: isPage });
    const network = fetch(req).then((res) => {
      if (res && (res.ok || res.type === 'opaque')) cache.put(key, res.clone()).catch(() => {});
      return res;
    }).catch(() => null);

    if (cached) {
      event.waitUntil(network);
      return cached;
    }
    const res = await network;
    return res || Response.error();
  })());
});
