import {readFile,mkdir,copyFile,stat} from 'node:fs/promises';
import {Script} from 'node:vm';
import path from 'node:path';
const root=process.cwd();
const files=['index.html','portfolio.css','portfolio.js','scene.js','DIR_Amjad_Masoud_PMP_Resume.pdf','assets/architecture.png','assets/favicon.svg','vendor/three.min.js','vendor/THREE-LICENSE.txt'];
const html=await readFile(path.join(root,'index.html'),'utf8');
const ids=[...html.matchAll(/\bid="([^"]+)"/g)].map(m=>m[1]);
if(new Set(ids).size!==ids.length)throw Error('Duplicate HTML IDs');
for(const match of html.matchAll(/\b(?:href|src)="([^"]+)"/g)){
  const ref=match[1];
  if(ref.startsWith('#')){if(!ids.includes(ref.slice(1)))throw Error(`Missing anchor: ${ref}`);continue;}
  if(/^(?:https?:|mailto:|tel:)/.test(ref))continue;
  if(!files.includes(ref))throw Error(`Unpackaged local asset: ${ref}`);
  if(!(await stat(path.join(root,ref))).isFile())throw Error(`Missing local asset: ${ref}`);
}
for(const file of ['portfolio.js','scene.js','vendor/three.min.js'])new Script(await readFile(path.join(root,file),'utf8'),{filename:file});
for(const match of html.matchAll(/\baria-controls="([^"]+)"/g))if(!ids.includes(match[1]))throw Error(`Missing controlled element: ${match[1]}`);
const forms=[...html.matchAll(/data-scene="([^"]+)"/g)].map(m=>m[1]);
if(forms.length!==3||new Set(forms).size!==3)throw Error('Missing sculpture controls');
if([...html.matchAll(/class="career-step"/g)].length!==5)throw Error('Career history is incomplete');
if([...html.matchAll(/data-category=/g)].length!==5)throw Error('Expertise inventory is incomplete');
for(const file of files){await mkdir(path.dirname(path.join(root,'dist',file)),{recursive:true});await copyFile(path.join(root,file),path.join(root,'dist',file));}
console.log(`Validated ${ids.length} unique anchors, all local references, 3D controls, five career chapters, five expertise disciplines, and JavaScript syntax.`);
console.log(`Static site ready: ${files.length} files in dist/`);
