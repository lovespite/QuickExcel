import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const here = import.meta.dirname;
const toolRoot = path.resolve(here, '..');
const generator = path.join(toolRoot, 'src/generate.mjs');
const fixture = path.join(toolRoot, 'fixtures/mini.d.ts');
const output = fs.mkdtempSync(path.join(os.tmpdir(), 'quickexcel-apigen-'));
const declarations = path.join(output, 'quickexcel.d.ts');
const capabilities = path.join(output, 'capabilities.json');
const coverage = path.join(output, 'coverage.json');

// The Linux sandbox's V8 optimizer crashes on repeated TypeScript parses; jitless
// keeps the fixture test deterministic and is harmless elsewhere.
execFileSync(process.execPath, [
  '--jitless',
  generator,
  '--source', fixture,
  '--overrides', path.join(toolRoot, 'fixtures/capabilities.json'),
  '--out', declarations,
  '--capabilities', capabilities,
  '--coverage', coverage,
], { stdio: 'inherit' });

// A second run in check mode must be clean when nothing changed.
execFileSync(process.execPath, [
  '--jitless',
  generator,
  '--source', fixture,
  '--overrides', path.join(toolRoot, 'fixtures/capabilities.json'),
  '--out', declarations,
  '--capabilities', capabilities,
  '--coverage', coverage,
  '--check',
], { stdio: 'inherit' });

const dts = fs.readFileSync(declarations, 'utf8');
const manifest = JSON.parse(fs.readFileSync(capabilities, 'utf8'));

assert.ok(dts.includes('declare namespace QuickExcel'), 'namespace was not renamed');
assert.ok(!dts.includes('namespace OfficeScript'), 'OfficeScript namespace was not removed');
assert.ok(!dts.includes('Represents the workbook'), 'documentation comments were not stripped');
assert.ok(dts.includes('getWorksheet(key: string): Worksheet;'), 'Workbook.getWorksheet is missing');
assert.ok(dts.includes('getPivotTables(): PivotTable[];'), 'unsupported members must remain in the skeleton');
assert.ok(dts.includes('getValueTypes(): RangeValueType[][];'), 'RangeValueType members are missing');
assert.ok(dts.includes('interface VirtualFiles'), 'host extensions were not appended');
const selected = dts.split('namespace Selected {')[1].split('// QuickExcel host extensions')[0];
assert.ok(selected.includes('getWorksheet(key: string): Worksheet;'));
assert.ok(selected.includes('getRange('));
assert.ok(!selected.includes('getPivotTables') && !selected.includes('getFormat'));
assert.deepEqual(manifest.selectedModel.Workbook, ['getWorksheet']);
assert.ok(dts.includes('Source SHA-256:'), 'source identity header is missing');

assert.equal(manifest.apiVersion, 3);
assert.equal(manifest.types.Workbook.members.getWorksheet.kind, 'host');
assert.equal(manifest.types.Workbook.members.getWorksheet.host, 'sheet');
assert.equal(manifest.types.Workbook.members.getWorksheet.signature, 'getWorksheet(key: string): Worksheet');
assert.equal(manifest.types.Range.members.clear.signature, 'clear(applyTo?: string): void');
assert.equal(manifest.types.RangeFormat.members.getFill.kind, 'local');
assert.ok(!manifest.types.Workbook.members.getWorksheet.signature.includes('/**'), 'signatures must not carry documentation comments');
assert.equal(manifest.types.Workbook.members.getPivotTables, undefined, 'unsupported members must not be in the manifest');
assert.equal(manifest.types.RangeFill.members.setColor.cost, 20000);
assert.ok(manifest.typeCount >= 10, 'type count should include every fixture declaration');
assert.ok(manifest.source.sha256.length === 64, 'source SHA-256 is missing');

const report = JSON.parse(fs.readFileSync(coverage, 'utf8'));
assert.equal(report.totals.types, manifest.typeCount);
assert.ok(report.totals.members > 0, 'coverage should count interface members');
assert.equal(report.types.Workbook.implemented, 6);
assert.equal(report.types.Worksheet.implemented, 4);

console.log('ApiGen fixture test passed.');
