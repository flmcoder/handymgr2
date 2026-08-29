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

const indexUrl = new URL('../dist/index.html', import.meta.url);
const indexHtml = await readFile(indexUrl, 'utf8');
const stampedIndexHtml = indexHtml
  .replace(/(id="vaultSessionVersion">)Secured Session: v[^<]+/, `$1Secured Session: v${version}`)
  .replace(/(id="buildBadgeTag">)[^<]+/, `$1${version}`)
  .replace(/(id="sidebarBrandVer">)v[^<]+/, `$1v${version}`);

if (!stampedIndexHtml.includes(`Secured Session: v${version}`)
  || !stampedIndexHtml.includes(`id="buildBadgeTag">${version}`)
  || !stampedIndexHtml.includes(`id="sidebarBrandVer">v${version}`)) {
  throw new Error('Could not stamp all generated HTML version labels');
}

await writeFile(indexUrl, stampedIndexHtml);
console.log(`Build stamped as v${version}`);