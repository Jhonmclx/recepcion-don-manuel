// Recepción Don Manuel — funciona sin internet: guarda la app y la lista de productos en el celular
const CACHE='rdm-v3';
const CORE=['./','./index.html','./inventario/','./inventario/index.html','./inventario/manifest.json','./pedidos/','./pedidos/index.html','./pedidos/manifest.json','https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js','./catalogo.json','./manifest.json','./icon-192.png','./icon-512.png',
  'https://unpkg.com/@zxing/library@0.21.3/umd/index.min.js'];
self.addEventListener('install',e=>{ e.waitUntil(caches.open(CACHE).then(c=>Promise.allSettled(CORE.map(u=>c.add(new Request(u,{mode:u.startsWith('http')?'no-cors':'same-origin'}))))).then(()=>self.skipWaiting())); });
self.addEventListener('activate',e=>{ e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())); });
self.addEventListener('fetch',e=>{
  const req=e.request; if(req.method!=='GET') return;
  const url=new URL(req.url);
  if(url.hostname.endsWith('script.google.com')||url.hostname.endsWith('googleusercontent.com')) return; // servidor: siempre en línea
  e.respondWith(caches.open(CACHE).then(async c=>{
    const hit=await c.match(req,{ignoreSearch:url.origin===location.origin});
    const net=fetch(req).then(r=>{ if(r && (r.ok||r.type==='opaque')) c.put(req,r.clone()); return r; }).catch(()=>null);
    if(hit){ e.waitUntil(net); return hit; }
    const r=await net; if(r) return r;
    if(req.mode==='navigate') return (await c.match('./index.html')) || Response.error();
    return Response.error();
  }));
});
