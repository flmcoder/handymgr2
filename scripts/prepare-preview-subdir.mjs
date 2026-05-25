import { cp, mkdir, readdir, rm } from 'node:fs/promises';
import path from 'node:path';

const distDir = path.resolve('dist');
const subDir = path.join(distDir, 'handymgr2');

await mkdir(distDir, { recursive: true });
await rm(subDir, { recursive: true, force: true });
await mkdir(subDir, { recursive: true });

const entries = await readdir(distDir, { withFileTypes: true });

for (const entry of entries) {
  if (entry.name === 'handymgr2') continue;
  const from = path.join(distDir, entry.name);
  const to = path.join(subDir, entry.name);
  await cp(from, to, { recursive: true });
}

console.log('Prepared preview subdir: dist/handymgr2');