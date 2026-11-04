#!/usr/bin/env node
// Downloads the pinned Office Scripts API definition and verifies its identity
// before it is written into third_party. Run once per source update.
import crypto from 'node:crypto';
import fs from 'node:fs';
import https from 'node:https';
import path from 'node:path';

const repoRoot = path.resolve(import.meta.dirname, '..', '..', '..');
const lockPath = path.join(repoRoot, 'third_party/office-scripts-docs/source.lock.json');
const lock = JSON.parse(fs.readFileSync(lockPath, 'utf8'));
const target = path.join(repoRoot, lock.localPath);
const updateLock = process.argv.includes('--update-lock');

const url = `https://raw.githubusercontent.com/${lock.repository}/${lock.commit}/${lock.file}`;

function download(address, redirects = 0) {
  return new Promise((resolve, reject) => {
    const request = https.get(address, { headers: { 'user-agent': 'quickexcel-api-gen' } }, (response) => {
      if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location && redirects < 5) {
        response.resume();
        resolve(download(new URL(response.headers.location, address).toString(), redirects + 1));
        return;
      }
      if (response.statusCode !== 200) {
        response.resume();
        reject(new Error(`HTTP ${response.statusCode} for ${address}`));
        return;
      }
      const chunks = [];
      response.on('data', (chunk) => chunks.push(chunk));
      response.on('end', () => resolve(Buffer.concat(chunks)));
    });
    request.on('error', reject);
  });
}

const buffer = await download(url);
if (buffer.length !== lock.size) {
  throw new Error(`Source size mismatch: expected ${lock.size}, received ${buffer.length}.`);
}

const gitBlobSha1 = crypto.createHash('sha1').update(`blob ${buffer.length}\0`).update(buffer).digest('hex');
if (gitBlobSha1 !== lock.gitBlobSha1) {
  throw new Error(`Git blob identity mismatch: expected ${lock.gitBlobSha1}, received ${gitBlobSha1}.`);
}

const sha256 = crypto.createHash('sha256').update(buffer).digest('hex');
if (!lock.sha256) {
  if (!updateLock) {
    throw new Error('Source lock has no sha256 yet. Run with --update-lock once, review the change, and commit the lock.');
  }
  lock.sha256 = sha256;
  fs.writeFileSync(lockPath, `${JSON.stringify(lock, null, 2)}\n`);
} else if (lock.sha256 !== sha256) {
  throw new Error(`SHA-256 mismatch: expected ${lock.sha256}, received ${sha256}.`);
}

fs.mkdirSync(path.dirname(target), { recursive: true });
fs.writeFileSync(target, buffer);
console.log(`Wrote ${path.relative(repoRoot, target)} (${buffer.length} bytes, sha256 ${sha256}).`);
