#!/usr/bin/env node
// Generates QuickExcel script declarations and the capability manifest from the
// pinned Office Scripts API definition. The generated declaration is a modified
// work under CC BY 4.0; see third_party/office-scripts-docs/NOTICE.md.
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import ts from 'typescript';

const repoRoot = path.resolve(import.meta.dirname, '..', '..', '..');
const defaults = {
  source: path.join(repoRoot, 'third_party/office-scripts-docs/office-scripts-docs.d.ts'),
  out: path.join(repoRoot, 'artifacts/api-gen/quickexcel.d.ts'),
  capabilities: path.join(repoRoot, 'artifacts/api-gen/capabilities.json'),
  coverage: path.join(repoRoot, 'artifacts/api-gen/coverage.json'),
};
const lockPath = path.join(repoRoot, 'third_party/office-scripts-docs/source.lock.json');
const overridesPath = path.resolve(argument('--overrides', path.join(import.meta.dirname, 'capabilities.overrides.json')));
const extensionsPath = path.join(import.meta.dirname, 'quickexcel-extensions.d.ts');

function argument(name, fallback) {
  const index = process.argv.indexOf(name);
  return index > 0 && process.argv[index + 1] ? process.argv[index + 1] : fallback;
}

const sourcePath = path.resolve(argument('--source', defaults.source));
const outPath = path.resolve(argument('--out', defaults.out));
const capabilitiesPath = path.resolve(argument('--capabilities', defaults.capabilities));
const coveragePath = path.resolve(argument('--coverage', defaults.coverage));
const check = process.argv.includes('--check');
const verify = !process.argv.includes('--no-verify');

if (!fs.existsSync(sourcePath)) {
  console.error(`Source definition not found: ${sourcePath}`);
  console.error('Run "node tools/QuickExcel.ApiGen/src/fetch-source.mjs --update-lock" first.');
  process.exit(1);
}

function fail(message) {
  console.error(message);
  process.exit(1);
}

function sourceIdentity() {
  const buffer = fs.readFileSync(sourcePath);
  const sha256 = crypto.createHash('sha256').update(buffer).digest('hex');
  const gitBlobSha1 = crypto.createHash('sha1').update(`blob ${buffer.length}\0`).update(buffer).digest('hex');
  let lock = null;
  try {
    lock = JSON.parse(fs.readFileSync(lockPath, 'utf8'));
  } catch {
    // A local fixture has no lock; identity is still recorded from the bytes.
  }
  const pinned = lock && path.resolve(repoRoot, lock.localPath) === sourcePath;
  if (sourcePath === defaults.source && !pinned) fail('Pinned declaration source lock is missing or invalid.');
  if (pinned && (lock.sha256 !== sha256 || lock.gitBlobSha1 !== gitBlobSha1)) fail('Pinned declaration source hash mismatch.');
  return {
    repository: pinned ? lock.repository : 'local fixture',
    commit: pinned ? lock.commit : null,
    file: pinned ? lock.file : path.basename(sourcePath),
    gitBlobSha1,
    sha256,
    bytes: buffer.length,
  };
}

const identity = sourceIdentity();
const text = fs.readFileSync(sourcePath, 'utf8');
const sourceFile = ts.createSourceFile('office-scripts-docs.d.ts', text, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);

const types = new Map();
for (const statement of sourceFile.statements) {
  if (!ts.isModuleDeclaration(statement) || !ts.isIdentifier(statement.name) || statement.name.text !== 'ExcelScript') {
    continue;
  }
  if (!statement.body || !ts.isModuleBlock(statement.body)) continue;
  for (const member of statement.body.statements) {
    if (ts.isInterfaceDeclaration(member)) types.set(member.name.text, { kind: 'interface', node: member });
    else if (ts.isEnumDeclaration(member)) types.set(member.name.text, { kind: 'enum', node: member });
    else if (ts.isTypeAliasDeclaration(member)) types.set(member.name.text, { kind: 'type', node: member });
    else if (ts.isClassDeclaration(member)) fail(`Unsupported class declaration: ${member.name?.text ?? '<anonymous>'}`);
  }
}
if (types.size === 0) fail('No ExcelScript type declarations were found in the source file.');

function declarationName(member) {
  if (member.name && (ts.isIdentifier(member.name) || ts.isStringLiteral(member.name) || ts.isNumericLiteral(member.name))) {
    return member.name.text;
  }
  return null;
}

function memberNames(node) {
  return new Set(node.members.map(declarationName).filter((name) => name !== null));
}

const printer = ts.createPrinter({ newLine: ts.NewLineKind.LineFeed, removeComments: true });

function emitDeclaration(node) {
  return printer
    .printNode(ts.EmitHint.Unspecified, node, sourceFile)
    .split('\n')
    .map((line) => `  ${line}`)
    .join('\n');
}

const declarations = [];
for (const entry of types.values()) declarations.push(emitDeclaration(entry.node));

const extensions = fs.existsSync(extensionsPath) ? fs.readFileSync(extensionsPath, 'utf8').trim() : '';
const sourceReference = identity.commit
  ? ` * https://github.com/${identity.repository}/blob/${identity.commit}/${identity.file}`
  : ` * Local fixture: ${identity.file}`;
const header = `/*!
 * QuickExcel script declarations. Generated by tools/QuickExcel.ApiGen; do not edit.
 *
 * Derived from the Office Scripts API definition:
${sourceReference}
 * Copyright (c) Microsoft Corporation. Licensed under CC BY 4.0.
 * https://creativecommons.org/licenses/by/4.0/legalcode
 *
 * Modified for QuickExcel: the namespace was renamed to QuickExcel, documentation
 * comments were removed, the OfficeScript namespace was removed, and QuickExcel
 * host extensions were appended. This file contains declarations only; all
 * behavior is implemented independently by QuickExcel.
 *
 * Source blob (git SHA-1): ${identity.gitBlobSha1}
 * Source SHA-256: ${identity.sha256}
 */

declare namespace QuickExcel {
${declarations.join('\n\n')}
}

${extensions}
`;

const overrides = JSON.parse(fs.readFileSync(overridesPath, 'utf8'));
const typesOverrides = overrides.types ?? {};
for (const [typeName, typeSpec] of Object.entries(typesOverrides)) {
  const entry = types.get(typeName);
  if (!entry || entry.kind !== 'interface') fail(`Capability override references unknown interface ${typeName}.`);
  const names = memberNames(entry.node);
  for (const [memberName, spec] of Object.entries(typeSpec.members ?? {})) {
    if (!names.has(memberName)) fail(`Capability override references unknown member ${typeName}.${memberName}.`);
    if (spec.kind !== 'local' && spec.kind !== 'host') fail(`Capability ${typeName}.${memberName} has an invalid kind.`);
    if (spec.kind === 'host') {
      if (typeof spec.host !== 'string' || !spec.host) fail(`Capability ${typeName}.${memberName} needs a host method.`);
      if (!Number.isInteger(spec.arity) || spec.arity < 0) fail(`Capability ${typeName}.${memberName} needs a non-negative arity.`);
      if (!Number.isFinite(spec.cost) || spec.cost <= 0) fail(`Capability ${typeName}.${memberName} needs a positive cost.`);
    }
  }
}

const enums = {};
for (const name of overrides.enums ?? []) {
  const entry = types.get(name);
  if (!entry || entry.kind !== 'enum') fail(`Unknown runtime enum ${name}.`);
  let next = 0;
  enums[name] = {};
  for (const member of entry.node.members) {
    if (member.initializer) {
      if (!ts.isNumericLiteral(member.initializer)) fail(`Unsupported enum initializer ${name}.`);
      next = Number(member.initializer.text);
    }
    enums[name][declarationName(member)] = next++;
  }
}
const manifest = {
  apiVersion: overrides.apiVersion ?? 3,
  source: identity,
  typeCount: types.size,
  types: typesOverrides,
  enums,
  limits: overrides.limits ?? { rows: 100000, columns: 1024 },
};
const manifestText = `${JSON.stringify(manifest, null, 2)}\n`;

function writeOrCheck(target, content, label) {
  fs.mkdirSync(path.dirname(target), { recursive: true });
  if (!check) {
    fs.writeFileSync(target, content);
    console.log(`Wrote ${path.relative(repoRoot, target)} (${Buffer.byteLength(content)} bytes, ${label}).`);
    return;
  }
  const current = fs.existsSync(target) ? fs.readFileSync(target, 'utf8') : null;
  if (current !== content) fail(`${path.relative(repoRoot, target)} is out of date. Run the generator.`);
  console.log(`${path.relative(repoRoot, target)} is up to date.`);
}

writeOrCheck(outPath, `${header}`, 'declarations');
writeOrCheck(capabilitiesPath, manifestText, 'capabilities');

if (!check) {
  const coverageTypes = {};
  let totalMembers = 0;
  let implementedMembers = 0;
  let implementedTypes = 0;
  for (const [name, entry] of types) {
    const members = entry.kind === 'interface' ? memberNames(entry.node).size : 0;
    const implemented = Object.keys(typesOverrides[name]?.members ?? {}).length;
    totalMembers += members;
    implementedMembers += implemented;
    if (implemented > 0) implementedTypes++;
    coverageTypes[name] = { kind: entry.kind, members, implemented };
  }
  const coverage = {
    apiVersion: manifest.apiVersion,
    totals: { types: types.size, members: totalMembers, implementedTypes, implementedMembers },
    types: coverageTypes,
  };
  fs.mkdirSync(path.dirname(coveragePath), { recursive: true });
  fs.writeFileSync(coveragePath, `${JSON.stringify(coverage, null, 2)}\n`);
  console.log(
    `Coverage: ${implementedTypes}/${types.size} types, ${implementedMembers}/${totalMembers} members -> ${path.relative(repoRoot, coveragePath)}.`,
  );
}

if (verify) {
  const program = ts.createProgram({
    rootNames: [outPath],
    options: {
      noEmit: true,
      strict: true,
      target: ts.ScriptTarget.ES2020,
      lib: ['lib.es2020.d.ts'],
      types: [],
      skipLibCheck: false,
    },
  });
  const diagnostics = ts.getPreEmitDiagnostics(program);
  if (diagnostics.length > 0) {
    const host = {
      getCanonicalFileName: (name) => name,
      getCurrentDirectory: () => process.cwd(),
      getNewLine: () => '\n',
    };
    fail(ts.formatDiagnosticsWithColorAndContext(diagnostics, host));
  }
  console.log('Declaration verification passed.');
}
