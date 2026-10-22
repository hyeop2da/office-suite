// 오프라인 지원 서비스 워커. 앱 파일(코드·화면·샘플·보기 엔진)만 이 기기 캐시에 보관한다.
// 사용자가 연 문서와 편집 내용은 여기에 들어오지 않는다(문서는 브라우저 밖으로 나가지 않고, 임시 보존은 IndexedDB에만 있다).
const VERSION = '__VERSION__';
const SHELL = __SHELL__;

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()),
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  e.respondWith((async () => {
    const cache = await caches.open(VERSION);
    // 화면 이동 요청은 캐시된 첫 화면으로(오프라인에서도 열림). 온라인이면 네트워크 우선.
    if (req.mode === 'navigate') {
      try { const res = await fetch(req); if (res.ok) cache.put('index.html', res.clone()); return res; }
      catch { return (await cache.match('index.html')) ?? Response.error(); }
    }
    const hit = await cache.match(req, { ignoreSearch: true });
    if (hit) return hit;
    const res = await fetch(req); // 보기 엔진 등 큰 파일은 처음 쓸 때 받아 캐시에 둔다
    if (res.ok) cache.put(req, res.clone());
    return res;
  })());
});
