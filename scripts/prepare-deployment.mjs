import { readFile, writeFile, access, cp, rm } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { XMLValidator } from 'fast-xml-parser';

// Keep manifest.xml for local sideloading; publish a separate copy with HTTPS URLs.
const origin = 'https://wpsimagecompat.fogce.workers.dev';
const root = new URL('../', import.meta.url);
const template = await readFile(new URL('manifest.xml', root), 'utf8');
const manifest = template.replaceAll('https://localhost:3000', origin);
if (manifest.includes('localhost') || XMLValidator.validate(manifest) !== true) {
  throw new Error('Production manifest must be valid XML without localhost URLs.');
}
for (const path of [
  'src/taskpane/taskpane.html', 'support.html', 'index.html', 'version.json',
  'assets/icon-16.png', 'assets/icon-32.png', 'assets/icon-80.png',
]) {
  await access(new URL(`dist/${path}`, root));
}
const wpsSource = new URL('../wps-et/dist/', import.meta.url);
const wpsTarget = new URL('../dist/wps-et/', import.meta.url);
await rm(wpsTarget, { recursive: true, force: true });
await cp(wpsSource, wpsTarget, { recursive: true });
for (const path of ['manifest.xml', 'ribbon.xml', 'index.html', 'main.js', 'publish.html', 'ui/taskpane.html']) {
  await access(new URL(`../dist/wps-et/${path}`, import.meta.url));
}
for (const path of ['PHASE1.md', 'PHASE2.md', 'PHASE3.md', 'PHASE4.md', 'RELEASE_0_6.md', 'RELEASE_0_7.md', 'RELEASE_0_8.md']) {
  try {
    await access(new URL(`dist/${path}`, root));
    throw new Error(`Internal project record must not be deployed: ${path}`);
  } catch (error) {
    if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') continue;
    throw error;
  }
}
const target = new URL('dist/manifest.xml', root);
await writeFile(target, manifest);
console.log(`Production manifest prepared: ${fileURLToPath(target)}`);
