import fs from 'node:fs';

const root = new URL('../', import.meta.url);
const readJson = name => JSON.parse(fs.readFileSync(new URL(name, root), 'utf8'));
const writeJson = (name, value) => fs.writeFileSync(new URL(name, root), `${JSON.stringify(value, null, 2)}\n`);

const packageJson = readJson('package.json');
const parts = String(packageJson.version).replace(/^v/i, '').split('.').map(Number);
parts[2] = (parts[2] || 0) + 1;
const version = parts.join('.');
packageJson.version = version;
writeJson('package.json', packageJson);

const lock = readJson('package-lock.json');
lock.version = version;
if (lock.packages?.['']) lock.packages[''].version = version;
writeJson('package-lock.json', lock);

const replacements = [
  ['src/app.js', /var APP_VERSION = '[^']+';/, `var APP_VERSION = 'v${version}';`],
  ['index.html', /(<div class="sidebar-brand-ver" id="sidebarBrandVer">)v[^<]+/, `$1v${version}`],
  ['afproxy/config.ts', /(PROXY_APP_VERSION = env\("PROXY_APP_VERSION", ")[^"]+/, `$1v${version}`],
];
for (const [name, pattern, replacement] of replacements) {
  const file = new URL(name, root);
  fs.writeFileSync(file, fs.readFileSync(file, 'utf8').replace(pattern, replacement));
}

console.log(`Bumped application version to ${version}`);
