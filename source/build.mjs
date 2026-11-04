// Build: node source/build.mjs  (run from the app folder, needs `npm install` in source/)
import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'fs';
import { createHash } from 'crypto';
import { execSync } from 'child_process';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');
await build({
  entryPoints: [join(here, 'app.js')], bundle: true, minify: true, format: 'iife',
  target: ['chrome100'], outfile: join(root, 'app.js'), legalComments: 'none',
});
const files = ['./', 'index.html', 'styles.css', 'app.js', 'config.js', 'manifest.webmanifest', 'recipes.xlsx',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png'];
const h = createHash('sha256');
for (const f of files.slice(1)) h.update(readFileSync(join(root, f)));
h.update(readFileSync(join(here, 'sw.template.js')));
const hash = h.digest('hex').slice(0, 10);
// Version number counts up: wcwb-v-001, wcwb-v-002, ... It goes up by one from the last committed
// source/version.json whenever the built files differ from that commit, so rebuilding before a commit
// doesn't skip numbers. Commit source/version.json with app.js and sw.js.
let last = { n: 0, hash: '' };
try { last = JSON.parse(execSync('git show HEAD:source/version.json', { cwd: root, stdio: ['ignore', 'pipe', 'ignore'] }).toString()); } catch { /* first numbered build */ }
const n = last.hash === hash ? last.n : last.n + 1;
writeFileSync(join(here, 'version.json'), JSON.stringify({ n, hash }) + '\n');
const version = 'wcwb-v-' + String(n).padStart(3, '0');
const app = readFileSync(join(root, 'app.js'), 'utf8');
writeFileSync(join(root, 'app.js'), `window.APP_VERSION=${JSON.stringify(version)};\n` + app);
writeFileSync(join(root, 'sw.js'), readFileSync(join(here, 'sw.template.js'), 'utf8')
  .replace('__VERSION__', version).replace('__FILES__', JSON.stringify(files)));
console.log('built', version, (app.length / 1024).toFixed(0) + ' KB');
