import fs from 'node:fs/promises';
import fsSync from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '../..');
const output = path.join(root, 'artifacts/script-libraries');
const npm = path.join(output, 'npm');
const hash = value => crypto.createHash('sha256').update(value).digest('hex').toUpperCase();
const manifest = JSON.parse(await fs.readFile(path.join(here, 'package.json'), 'utf8'));
await fs.mkdir(npm, {recursive:true});
// Native/compiler/Web builds can request the same assets concurrently.
const lockPath=path.join(output,'build.lock');
let buildLock;
for(let attempt=0;!buildLock;attempt++) {
  try { buildLock=await fs.open(lockPath,'wx'); await buildLock.writeFile(String(process.pid)); }
  catch(error) {
    if(error.code!=='EEXIST'||attempt>=120)throw error;
    const owner=Number(await fs.readFile(lockPath,'utf8').catch(()=>''));
    if(owner>0){try{process.kill(owner,0);}catch(e){if(e.code==='ESRCH')await fs.unlink(lockPath).catch(()=>{});}}
    await new Promise(resolve=>setTimeout(resolve,1000));
  }
}
let released=false;
process.on('exit',()=>{if(!released){try{fsSync.unlinkSync(lockPath);}catch{}}});
const lock = await fs.readFile(path.join(here, 'package-lock.json'));
const installed = await fs.readFile(path.join(npm, '.installed-lock')).catch(()=>null);
if (!installed || installed.toString() !== hash(lock)) {
  for (const file of ['package.json','package-lock.json']) await fs.copyFile(path.join(here,file),path.join(npm,file));
  const install = spawnSync(process.platform === 'win32' ? 'npm.cmd' : 'npm', ['ci','--ignore-scripts','--no-audit','--no-fund','--cache',path.join(output,'npm-cache')], {cwd:npm,stdio:'inherit',shell:process.platform==='win32'});
  if (install.status !== 0) throw new Error('Pinned script-library dependency restore failed');
  await fs.writeFile(path.join(npm,'.installed-lock'),hash(lock));
}
const require = createRequire(path.join(npm,'package.json'));
const {build} = require('esbuild');
const tsFile=path.join(root,'artifacts/jsc/typescript/package/lib/typescript.js');
if(!fsSync.existsSync(tsFile)) {
  const prepared=spawnSync(process.platform==='win32'?'python':'python3',[path.join(root,'tools/QuickExcel.JscProbe/prepare_compiler.py'),'--sources-only'],{cwd:root,stdio:'inherit'});
  if(prepared.status!==0)throw new Error('Pinned TypeScript source preparation failed');
}
const ts = require(tsFile);
const names = ['lodash','decimal.js','dayjs','papaparse','simple-statistics','zod','fuse.js','arquero','mathjs'];
const bundles = {};
const lodashExports = [...(await fs.readFile(path.join(npm,'node_modules/lodash-es/lodash.js'),'utf8')).matchAll(/export \{ default as (\w+) \}/g)].map(m=>m[1]);
for (const name of names) {
  let contents = `export * from '${name}';`;
  if (name === 'lodash') contents = `import upstream from 'lodash'; export default upstream; export const lodash=upstream; export const {${lodashExports.filter(n=>n!=='lodash').join(',')}}=upstream;`;
  else if (['dayjs','papaparse','arquero'].includes(name)) contents = await fs.readFile(path.join(here,'runtime',name+'.js'),'utf8');
  else if (['decimal.js','fuse.js'].includes(name)) contents += `export {default} from '${name}';`;
  const browser = ['papaparse','arquero','mathjs'].includes(name);
  const result = await build({stdin:{contents,resolveDir:npm},bundle:true,format:'esm',platform:browser?'browser':'neutral',mainFields:browser?['browser','module','main']:['module','main'],conditions:['import'],target:'es2022',minify:true,write:false,legalComments:'inline',metafile:true});
  if (result.outputFiles.length !== 1 || Object.values(result.metafile.outputs).some(o=>o.imports.length)) throw new Error(`Unresolved runtime dependency: ${name}`);
  bundles[name] = result.outputFiles[0].text;
}
const supportResult=await build({stdin:{contents:await fs.readFile(path.join(here,'runtime/support.js'),'utf8'),resolveDir:npm},bundle:true,format:'iife',platform:'neutral',mainFields:['module','main'],target:'es2022',minify:true,write:false,legalComments:'inline'});
const support=supportResult.outputFiles[0].text;

// Preserve upstream declaration structure and transitive imports in a closed VFS.
const plugins=[...(await fs.readFile(path.join(here,'runtime/dayjs.js'),'utf8')).matchAll(/from 'dayjs\/plugin\/(\w+)\.js'/g)].map(m=>m[1]);
const typesRoot=path.join(output,'type-entry.ts');
await fs.writeFile(typesRoot,names.map((n,i)=>`import * as lib${i} from '${n}';`).join('\n')+'\n'+plugins.map(p=>`import 'dayjs/plugin/${p}';`).join('\n'));
const options={target:ts.ScriptTarget.ES2022,module:ts.ModuleKind.ESNext,moduleResolution:ts.ModuleResolutionKind.Bundler,types:[],lib:['lib.es2022.d.ts'],allowSyntheticDefaultImports:true,skipLibCheck:false,noEmit:true};
const compilerHost=ts.createCompilerHost(options);
const standardResolve=ts.resolveModuleName;
compilerHost.resolveModuleNames=(ids,file)=>ids.map(id=>standardResolve(id,path.join(npm,'entry.ts'),options,compilerHost).resolvedModule ?? standardResolve(id,file,options,compilerHost).resolvedModule);
const program=ts.createProgram([typesRoot],options,compilerHost);
const types={};
for(const file of program.getSourceFiles()) {
  const relative=path.relative(path.join(npm,'node_modules'),file.fileName).replaceAll('\\','/');
  if(relative.startsWith('../') || !file.isDeclarationFile || relative.startsWith('@types/node/')) continue;
  let text=file.text;
  if(relative==='arquero/dist/types/table/ColumnTable.d.ts') text=text.replace('...values?:','...values:');
  if(relative==='@uwdata/flechette/dist/types/table.d.ts') text=text.replace('getFactory;', 'getFactory: (batch: number) => (index: number) => Record<string, unknown>;');
  if(/^arquero\/dist\/types\/format\/from-(csv|json|fixed)\.d\.ts$/.test(relative)) text=text.replaceAll('ReadableStream<string>','ReadableStream<string | Uint8Array>');
  if(relative==='@types/papaparse/index.d.ts') {
    // Narrow upstream declarations to the synchronous in-memory API exposed above.
    const tree=ts.createSourceFile(relative,text,ts.ScriptTarget.ES2022,true);
    const removed=[];
    for(const node of tree.statements) {
      if(ts.isImportDeclaration(node) && node.moduleSpecifier.text==='stream' ||
          ts.isFunctionDeclaration(node) && node.name?.text==='parse' && !(node.parameters[0]?.name.getText(tree)==='csvString' && node.parameters[1]?.type?.getText(tree).startsWith('ParseConfig<')) ||
          ts.isVariableStatement(node) && node.declarationList.declarations.some(d=>d.name.getText(tree)==='NODE_STREAM_INPUT')) removed.push([node.getFullStart(),node.end]);
    }
    for(const [start,end] of removed.reverse()) text=text.slice(0,start)+text.slice(end);
    text=text.replace(/\/\/\/ <reference types="node" \/>/g,'').replace('export type LocalFile = File | NodeJS.ReadableStream;','export type LocalFile = never;');
    text+='\ndeclare const Papa: {parse: typeof parse; unparse: typeof unparse; WORKERS_SUPPORTED: false};\nexport default Papa;\n';
  }
  types['/builtin/node_modules/'+relative]=text;
}
const paths={};
for(const name of names){
  const resolved=standardResolve(name,path.join(npm,'entry.ts'),options,compilerHost).resolvedModule;
  if(!resolved)throw new Error(`Missing declarations for ${name}`);
  paths[name]='/builtin/node_modules/'+path.relative(path.join(npm,'node_modules'),resolved.resolvedFileName).replaceAll('\\','/');
}
// Internal declaration references must keep pointing to their bundled package
// even when a workbook declares its own module called dayjs/decimal.js/etc.
for (const [file,source] of Object.entries(types)) {
  const tree=ts.createSourceFile(file,source,ts.ScriptTarget.ES2022,true), edits=[];
  function visit(node) {
    const parent=node.parent;
    if(ts.isStringLiteral(node) && Object.hasOwn(paths,node.text) &&
      (ts.isImportDeclaration(parent) && parent.moduleSpecifier===node ||
       ts.isExportDeclaration(parent) && parent.moduleSpecifier===node ||
       ts.isModuleDeclaration(parent) && parent.name===node ||
       ts.isLiteralTypeNode(parent) && ts.isImportTypeNode(parent.parent))) {
      let target=path.posix.relative(path.posix.dirname(file),paths[node.text]).replace(/\.d\.[cm]?ts$/,'');
      if(!target.startsWith('.')) target='./'+target;
      edits.push([node.getStart(tree),node.end,JSON.stringify(target)]);
    }
    ts.forEachChild(node,visit);
  }
  visit(tree); let text=source;
  for(const [start,end,value] of edits.reverse()) text=text.slice(0,start)+value+text.slice(end);
  types[file]=text;
}
// Package metadata is needed for transitive type resolution, never runtime loading.
for(const key of Object.keys(types)) {
  const rel=key.slice('/builtin/node_modules/'.length).split('/');
  const packageName=rel[0].startsWith('@')?rel.slice(0,2).join('/'):rel[0];
  const file='/builtin/node_modules/'+packageName+'/package.json';
  if(!(file in types)) types[file]=await fs.readFile(path.join(npm,'node_modules',packageName,'package.json'),'utf8');
}
types['/builtin/environment.d.ts']=(await fs.readFile(path.join(here,'runtime/environment.d.ts'),'utf8')).replaceAll('\r\n','\n');
const typeRoots=['/builtin/environment.d.ts',...plugins.map(p=>'/builtin/node_modules/dayjs/plugin/'+p+'.d.ts')];
const runtime={schemaVersion:1,modules:bundles,support};
const declarations={schemaVersion:1,paths,files:types,roots:typeRoots};
const catalog={schemaVersion:1,packages:Object.fromEntries(names.map(name=>[name,{version:manifest.dependencies[name],sha256:hash(bundles[name])}])),supportSha256:hash(support),declarationsSha256:hash(JSON.stringify(declarations))};
const identity=hash(JSON.stringify(catalog));
async function write(name,data){const target=path.join(output,name);if((await fs.readFile(target).catch(()=>null))?.toString()!==data){const temp=target+'.'+crypto.randomUUID()+'.tmp';await fs.writeFile(temp,data);await fs.rename(temp,target);}}
await write('runtime.json',JSON.stringify(runtime));
await write('declarations.json',JSON.stringify(declarations));
await write('catalog.json',JSON.stringify(catalog,null,2));
await write('identity.txt',identity);
await write('definition-index.json',JSON.stringify({paths,files:Object.keys(types).filter(n=>/\.d\.[cm]?ts$/.test(n))}));

// Include upstream notices for runtime, transitive and declaration packages.
const locks=JSON.parse(lock).packages;
const notices=[];
for(const [location,pkg] of Object.entries(locks)){
  if(!location || pkg.optional || location.includes('/@esbuild/') || location==='node_modules/esbuild')continue;
  const directory=path.join(npm,location);
  const list=await fs.readdir(directory).catch(()=>[]);
  const files=list.filter(n=>/^(licen[cs]e|copying|notice)(\.|$)/i.test(n));
  if(location==='node_modules/javascript-natural-sort') {
    const mit=(await fs.readFile(path.join(npm,'node_modules/decimal.js/LICENCE.md'),'utf8')).split('Permission is hereby granted')[1];
    notices.push(`## javascript-natural-sort ${pkg.version}\n\n${(await fs.readFile(path.join(directory,'naturalSort.js'),'utf8')).match(/^\/\*[\s\S]*?\*\//)[0]}\n\nUpstream identifies this work as MIT; standard MIT license text follows.\n\nPermission is hereby granted${mit}\n`);
    continue;
  }
  if(location==='node_modules/seedrandom') {
    notices.push(`## seedrandom ${pkg.version}\n\n${(await fs.readFile(path.join(directory,'seedrandom.js'),'utf8')).match(/^\/\*[\s\S]*?\*\//)[0]}\n`);
    for(const file of (await fs.readdir(path.join(directory,'lib'))).filter(f=>f.endsWith('.js')&&!f.endsWith('.min.js')).sort())
      notices.push(`### seedrandom/lib/${file}\n\n${(await fs.readFile(path.join(directory,'lib',file),'utf8')).match(/^(?:\/\/[^\n]*\n|\s*\n)+/)?.[0]??''}\n`);
    continue;
  }
  if(!files.length)throw new Error(`No license found for ${location}`);
  for(const file of files)notices.push(`## ${location.replace(/^node_modules\//,'')} ${pkg.version} — ${file}\n\n${await fs.readFile(path.join(directory,file),'utf8')}\n`);
}
await write('THIRD-PARTY-SCRIPT-LIBRARIES.txt',notices.join('\n'));
await fs.writeFile(path.join(output,'build.stamp'),identity);
console.log(`Script libraries: ${names.length}; runtime ${Buffer.byteLength(JSON.stringify(runtime))} bytes; types ${Buffer.byteLength(JSON.stringify(declarations))} bytes; ${identity}`);
await buildLock.close(); await fs.unlink(lockPath); released=true;
