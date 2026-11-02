import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs/promises';
import crypto from 'node:crypto';

const source = path.dirname(fileURLToPath(import.meta.url));
const root = path.dirname(source);
const dependencies = path.join(root, 'artifacts/avalonia/npm/node_modules');
const require = createRequire(path.join(dependencies, '../package.json'));
const { build } = require('esbuild');
const artifacts = path.resolve(root, 'artifacts/avalonia');
const destination = path.join(artifacts, 'web');
await fs.mkdir(artifacts, { recursive: true });
// Each manifest describes a fresh build, never stale assets left by a prior version.
const output = await fs.mkdtemp(path.join(artifacts, 'web-build-'));
const settings = { bundle: true, format: 'esm', target: 'es2022', minify: true, sourcemap: false,
  nodePaths: [dependencies], outdir: output, loader: { '.ttf': 'file' }, legalComments: 'linked' };
await build({ ...settings, entryPoints: { probe: path.join(source, 'src/probe.ts'),
  'editor.worker': path.join(dependencies, 'monaco-editor/esm/vs/editor/editor.worker.js'),
  'ts.worker': path.join(dependencies, 'monaco-editor/esm/vs/language/typescript/ts.worker.js') } });
await fs.copyFile(path.join(dependencies, 'pdfjs-dist/build/pdf.worker.min.mjs'), path.join(output, 'pdf.worker.mjs'));
await fs.copyFile(path.join(source, 'src/probe.html'), path.join(output, 'probe.html'));
await fs.cp(path.join(dependencies, 'pdfjs-dist/cmaps'), path.join(output, 'cmaps'), { recursive: true });
await fs.cp(path.join(dependencies, 'pdfjs-dist/standard_fonts'), path.join(output, 'standard_fonts'), { recursive: true });
await fs.cp(path.join(dependencies, 'pdfjs-dist/wasm'), path.join(output, 'wasm'), { recursive: true });
await fs.mkdir(path.join(output, 'licenses'), { recursive: true });
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
await fs.writeFile(path.join(output, 'manifest.json'), JSON.stringify({ schemaVersion: 1,
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
