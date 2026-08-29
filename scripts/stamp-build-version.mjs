import { readFile, writeFile } from 'node:fs/promises';

const packageJson = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
const version = String(packageJson.version || '').trim();

if (!/^\d+\.\d+\.\d+$/.test(version)) {
  throw new Error(`Invalid package version: ${version || '(empty)'}`);
}

const serviceWorkerUrl = new URL('../dist/sw.js', import.meta.url);
const serviceWorker = await readFile(serviceWorkerUrl, 'utf8');
const stampedServiceWorker = serviceWorker.replace(
  /const HM_CACHE_VERSION = "[^"]+";/,
  `const HM_CACHE_VERSION = "hm-static-v${version}";`,
);

if (stampedServiceWorker === serviceWorker && !serviceWorker.includes(`hm-static-v${version}`)) {
  throw new Error('Could not stamp the service-worker cache version');
}

await writeFile(serviceWorkerUrl, stampedServiceWorker);
console.log(`Build stamped as v${version}`);