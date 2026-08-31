/* Lingel Factory QR service worker — NETWORK-FIRST for our own files
   so a new deploy shows up immediately (cache is only an offline fallback). */
const C='lingel-factory-qr-v26';
const CORE=['index.html','./','qr-quality.html','qr-store.html','qr-dispatch.html','qr-box-lookup.html','qr-receive.html','manifest.webmanifest','icon.svg'];
self.addEventListener('install',e=>{self.skipWaiting();e.waitUntil(caches.open(C).then(c=>c.addAll(CORE).catch(()=>{})));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==C).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{const req=e.request,u=req.url;
 if(req.method!=='GET')return;
 if(u.indexOf('script.google.com')>=0||u.indexOf('googleusercontent.com')>=0)return; // backend + Drive always live
 if(u.indexOf(self.location.origin)===0){ // OUR files: network-first, cache fallback (offline)
  e.respondWith(fetch(req).then(r=>{if(r&&r.status===200){const cp=r.clone();caches.open(C).then(c=>c.put(req,cp));}return r;}).catch(()=>caches.match(req)));
  return;}
 // CDN libs + fonts: cache-first for speed/offline
 e.respondWith(caches.match(req).then(cached=>cached||fetch(req).then(r=>{if(r&&r.status===200){const cp=r.clone();caches.open(C).then(c=>c.put(req,cp));}return r;}).catch(()=>cached)));
});
