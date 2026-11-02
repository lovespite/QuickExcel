import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs/promises';
import crypto from 'node:crypto';
import { patchPdfWorker } from './patch-pdfjs.mjs';

const source = path.dirname(fileURLToPath(import.meta.url));
const root = path.dirname(source);
const dependencies = path.join(root, 'artifacts/avalonia/npm/node_modules');
const require = createRequire(path.join(dependencies, '../package.json'));
const { build } = require('esbuild');
const testBuild = process.argv.includes('--tests');
const probeBuild = process.argv.includes('--probe');
const artifacts = path.resolve(root, 'artifacts/avalonia');
const destination = path.join(artifacts, 'web');
await fs.mkdir(artifacts, { recursive: true });
// Each manifest describes a fresh build, never stale assets left by a prior version.
const output = await fs.mkdtemp(path.join(artifacts, 'web-build-'));
const settings = { bundle: true, format: 'esm', target: 'es2022', minify: true, sourcemap: false, define: { __QUICKEXCEL_WEB_TEST__: String(testBuild) },
  nodePaths: [dependencies], outdir: output, loader: { '.ttf': 'file', '.svg': 'file', '.png': 'file', '.gif': 'file' }, legalComments: 'linked',
  plugins: [{ name: 'declaration-text', setup(build) { build.onLoad({ filter: /\.d\.ts$/ }, async args => ({ contents: await fs.readFile(args.path, 'utf8'), loader: 'text' })); } }] };
await build({ ...settings, entryPoints: { ...(probeBuild ? { probe: path.join(source, 'src/probe.ts') } : {
  script: path.join(source, 'src/script.ts'), preview: path.join(source, 'src/preview.ts') }),
  'editor.worker': path.join(dependencies, 'monaco-editor/esm/vs/editor/editor.worker.js'),
  'ts.worker': path.join(dependencies, 'monaco-editor/esm/vs/language/typescript/ts.worker.js') } });
const worker = patchPdfWorker(await fs.readFile(path.join(dependencies, 'pdfjs-dist/build/pdf.worker.mjs'), 'utf8'));
await build({ stdin: { contents: worker, resolveDir: dependencies, sourcefile: 'pdf.worker.mjs', loader: 'js' },
  bundle: true, format: 'esm', target: 'es2022', minify: true, sourcemap: false, legalComments: 'linked', outfile: path.join(output, 'pdf.worker.mjs') });
for (const page of probeBuild ? ['probe'] : ['script', 'preview'])
  await fs.copyFile(path.join(source, `src/${page}.html`), path.join(output, `${page}.html`));
await fs.cp(path.join(dependencies, 'pdfjs-dist/cmaps'), path.join(output, 'cmaps'), { recursive: true });
await fs.cp(path.join(dependencies, 'pdfjs-dist/standard_fonts'), path.join(output, 'standard_fonts'), { recursive: true });
await fs.cp(path.join(dependencies, 'pdfjs-dist/wasm'), path.join(output, 'wasm'), {
  recursive: true, filter: file => !path.basename(file).startsWith('quickjs-eval.')
}); // PDF scripting is disabled; do not ship its separate JavaScript interpreter.
await fs.mkdir(path.join(output, 'licenses'), { recursive: true });
await fs.copyFile(path.join(source, 'patches/pdfjs-actualtext.md'), path.join(output, 'licenses/pdfjs-actualtext.md'));
for (const [name, file] of [['pdfjs-dist', 'LICENSE'], ['monaco-editor', 'LICENSE'], ['esbuild', 'LICENSE.md']]) {
  await fs.copyFile(path.join(dependencies, name, file), path.join(output, 'licenses', `${name}.txt`));
}
await fs.copyFile(path.join(dependencies, 'monaco-editor/ThirdPartyNotices.txt'), path.join(output, 'licenses/monaco-ThirdPartyNotices.txt'));
const files = [];
async function walk(directory) {
  for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
    const file = path.join(directory, entry.name);
    if (entry.isDirectory()) await walk(file);
    else if (entry.name !== 'manifest.json') {
      const bytes = await fs.readFile(file);
      files.push({ name: path.relative(output, file).replaceAll('\\', '/'), size: bytes.length,
        sha256: crypto.createHash('sha256').update(bytes).digest('hex') });
    }
  }
}
await walk(output);
await fs.writeFile(path.join(output, 'manifest.json'), JSON.stringify({ schemaVersion: 1, testBuild, probeBuild,
  packages: JSON.parse(await fs.readFile(path.join(source, 'package.json'), 'utf8')).dependencies,
  files: files.sort((a, b) => a.name.localeCompare(b.name)) }, null, 2));
try {
  const previous = await fs.lstat(destination);
  if (!previous.isDirectory() || previous.isSymbolicLink()) throw new Error('Web output must be a regular artifact directory');
  // Retain the previous build as evidence; both absolute paths are inside artifacts.
  await fs.rename(destination, path.join(artifacts, `web-previous-${crypto.randomUUID()}`));
} catch (error) { if (error.code !== 'ENOENT') throw error; }
await fs.rename(output, destination);
console.log(`Built ${files.length} offline resources: ${destination}`);
