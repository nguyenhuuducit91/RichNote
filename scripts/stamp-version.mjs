// Single source of truth for the version = package.json "version".
// Stamps it into ext.json, ext.dev.json and the About badge in index.html so they
// can never drift apart again. Runs automatically as part of `npm run build`.
import { readFileSync, writeFileSync } from 'node:fs';

const v = JSON.parse(readFileSync('package.json', 'utf8')).version;
if (!/^\d+\.\d+\.\d+$/.test(v)) {
  console.error('Invalid package.json version:', v);
  process.exit(1);
}

// Targeted regex replaces so we preserve each file's existing formatting.
function bumpJson(file) {
  const s = readFileSync(file, 'utf8');
  const out = s.replace(/("version":\s*")\d+\.\d+\.\d+(")/, `$1${v}$2`);
  if (out !== s) writeFileSync(file, out);
}
bumpJson('ext.json');
bumpJson('ext.dev.json');

const html = readFileSync('index.html', 'utf8');
const outHtml = html.replace(/(id="aboutVer">)v\d+\.\d+\.\d+(<)/, `$1v${v}$2`);
if (outHtml !== html) writeFileSync('index.html', outHtml);

console.log('Stamped version', v, '→ ext.json, ext.dev.json, index.html badge');
