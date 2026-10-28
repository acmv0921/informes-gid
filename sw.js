const V='gid-v63';
const CORE=['./','index.html','presets.json','fuentes.json','manifest.json','icon-192.png','icon-512.png','py/engine.py','tpl/F06_base.docx','tpl/F82_base.docx','wheels/python_docx-1.2.0-py3-none-any.whl'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>Promise.all(CORE.map(u=>fetch(new Request(u,{cache:'reload'})).then(r=>{if(!r.ok)throw new Error(u);return c.put(u,r);})))).then(()=>self.skipWaiting()));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==V).map(x=>caches.delete(x)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{const r=e.request;if(r.method!=='GET')return;const u=new URL(r.url);
 if(u.hostname==='cdn.jsdelivr.net'){e.respondWith(caches.open(V).then(c=>c.match(r).then(h=>h||fetch(r).then(n=>{if(n.ok)c.put(r,n.clone());return n;}))));return;}
 if(u.origin===location.origin){e.respondWith(fetch(r.mode==='navigate'?new Request(r.url,{cache:'no-cache'}):r).then(n=>{if(n.ok){const k=n.clone();caches.open(V).then(c=>c.put(r,k));}return n;}).catch(()=>caches.match(r)));}});
