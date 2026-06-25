const C="ersa-portal-v29";
const A=["./","index.html","manifest.webmanifest","logo-white.png","icon-192.png","icon-512.png","apple-touch-icon.png"];
self.addEventListener("install",e=>{self.skipWaiting();e.waitUntil(caches.open(C).then(c=>c.addAll(A).catch(()=>{})));});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x)))).then(()=>self.clients.claim()));});
self.addEventListener("fetch",e=>{
 const u=new URL(e.request.url);
 if(u.href.indexOf("script.google.com")>=0||e.request.method!=="GET")return; // API'yi cache'leme
 if(u.pathname.endsWith("config.js")){e.respondWith(fetch(e.request).catch(()=>caches.match(e.request)));return;}
 e.respondWith(fetch(e.request).then(r=>{const cp=r.clone();caches.open(C).then(c=>c.put(e.request,cp).catch(()=>{}));return r;}).catch(()=>caches.match(e.request).then(m=>m||caches.match("index.html"))));
});
