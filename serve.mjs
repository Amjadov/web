import http from 'node:http';
import {readFile,stat} from 'node:fs/promises';
import path from 'node:path';
const root=process.cwd();
const allowed=new Set(['index.html','portfolio.css','portfolio.js','scene.js','DIR_Amjad_Masoud_PMP_Resume.pdf','assets/architecture.png','assets/favicon.svg','vendor/three.min.js']);
const types={'.html':'text/html; charset=utf-8','.css':'text/css; charset=utf-8','.js':'text/javascript; charset=utf-8','.png':'image/png','.webp':'image/webp','.svg':'image/svg+xml','.pdf':'application/pdf','.woff2':'font/woff2'};
const server=http.createServer(async(req,res)=>{try{const pathname=decodeURIComponent(new URL(req.url,'http://localhost').pathname);const relative=pathname==='/'?'index.html':pathname.slice(1);const file=path.resolve(root,relative);if(!file.startsWith(root+path.sep)||!allowed.has(relative)){res.writeHead(404);res.end('Not found');return;}const info=await stat(file);if(!info.isFile())throw Error();res.writeHead(200,{'Content-Type':types[path.extname(file)]||'application/octet-stream','Cache-Control':'no-store'});res.end(await readFile(file));}catch{res.writeHead(404);res.end('Not found');}});
server.listen(4173,'127.0.0.1',()=>console.log('Local: http://127.0.0.1:4173'));
