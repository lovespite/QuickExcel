// Fast diagnostics during development; isolated-worker tests remain authoritative.
import fs from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import {fileURLToPath} from 'node:url';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'../..');
const payload=JSON.parse(await fs.readFile(path.join(root,'artifacts/jsc/compiler/typescript.bundle.json'),'utf8'));
payload.builtins=JSON.parse(await fs.readFile(path.join(root,'artifacts/script-libraries/declarations.json'),'utf8'));
payload.frontend=await fs.readFile(path.join(root,'QuickExcel.Scripting/TypeScript/frontend.js'),'utf8');
const context=vm.createContext({});
vm.runInContext(payload.compiler+'\n'+payload.frontend,context);
for(const name of Object.keys(payload.builtins.paths)) {
  const result=context.qxCompile({mode:'check',apiVersion:4,sourceMapMode:'none',modules:[{id:'main',language:'typescript',source:`import * as lib from '${name}'; export const n=lib;`}],entries:['main'],libraries:payload.libraries,declarations:payload.declarations,builtins:payload.builtins});
  console.log(name,JSON.stringify(result.diagnostics));
  if(result.diagnostics.length)process.exitCode=1;
}
const selfCheck=await fs.readFile(path.join(root,'QuickExcel.Desktop.Avalonia/Views/MainWindow.SelfCheck.cs'),'utf8');
const source=selfCheck.match(/var libraries = [\s\S]*?"""([\s\S]*?)"""/)[1];
const result=context.qxCompile({mode:'check',apiVersion:4,sourceMapMode:'none',modules:[{id:'main',language:'typescript',source}],entries:['main'],libraries:payload.libraries,declarations:payload.declarations,builtins:payload.builtins});
console.log('production self-check',JSON.stringify(result.diagnostics));
if(result.diagnostics.length)process.exitCode=1;
