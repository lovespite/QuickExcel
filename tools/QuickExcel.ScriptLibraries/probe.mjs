// Development-only compatibility probe. Dependencies are pinned under artifacts.
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const out = path.join(root, 'artifacts/script-libraries');
const modules = path.join(out, 'npm/node_modules');
const { build } = await import(pathToFileURL(path.join(modules, 'esbuild/lib/main.js')));
const names = ['lodash', 'decimal.js', 'dayjs', 'papaparse', 'simple-statistics', 'zod', 'fuse.js', 'arquero', 'mathjs'];
const result=[];
const cases=JSON.parse(await fs.readFile(path.join(root,'tools/QuickExcel.ScriptLibraries/cases.json'),'utf8'));
for(const name of names){
  try {
    const browser = ['papaparse','arquero','mathjs'].includes(name);
    const output=await build({stdin:{contents:`export * from '${name}';${['lodash','decimal.js','dayjs','papaparse','fuse.js'].includes(name)?`export {default} from '${name}';`:''}`,resolveDir:path.join(out,'npm')},bundle:true,format:'esm',platform:browser?'browser':'neutral',mainFields:browser?['browser','module','main']:['module','main'],conditions:['import'],target:'es2022',minify:true,write:false,legalComments:'inline'});
    const source=output.outputFiles[0].text;
    await fs.writeFile(path.join(out,name+'.mjs'),source);
    result.push({name,bytes:Buffer.byteLength(source),source});
    result.at(-1).probe=`import * as lib from '${name}'; export function PROBE(){return JSON.stringify([${Object.entries(cases[name]).map(([id,body])=>`(()=>{try{return {id:${JSON.stringify(id)},value:(()=>{${body}})()}}catch(e){return {id:${JSON.stringify(id)},error:String(e)}}})()`).join(',')}]);}`;
    console.log(`${name}: ${Buffer.byteLength(source)} bytes`);
  }catch(error){result.push({name,error:String(error)});console.log(`${name}: ${error.message}`);}
}
await fs.writeFile(path.join(out,'probes.json'),JSON.stringify(result));
