// Build: node source/build.mjs  (run from the app folder, needs `npm install` in source/)
import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'fs';
import { createHash } from 'crypto';
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
const version = h.digest('hex').slice(0, 10);
const app = readFileSync(join(root, 'app.js'), 'utf8');
writeFileSync(join(root, 'app.js'), `window.APP_VERSION=${JSON.stringify(version)};\n` + app);
writeFileSync(join(root, 'sw.js'), readFileSync(join(here, 'sw.template.js'), 'utf8')
  .replace('__VERSION__', version).replace('__FILES__', JSON.stringify(files)));
console.log('built', version, (app.length / 1024).toFixed(0) + ' KB');
