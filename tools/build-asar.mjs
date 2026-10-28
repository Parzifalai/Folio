// Собирает только сам Folio (main, preload, app/, иконка) в dist\next\...\app.asar —
// его ставит tools\swap-update.ps1. Electron и всё остальное не трогаем.
// (npm run pack на Windows не применял --ignore: cmd съедает «^», и в app.asar попадала прошлая сборка целиком.)
import { createPackage } from '@electron/asar';
import { cp, mkdir, rm, writeFile, readFile, stat } from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, '$1')), '..');
const stage = path.join(os.tmpdir(), 'folio-asar-stage');
const out = path.join(root, 'dist', 'next', 'Folio-win32-x64', 'resources', 'app.asar');

await rm(stage, { recursive: true, force: true });
await mkdir(path.join(stage, 'build'), { recursive: true });
for (const f of ['main.js', 'preload.js']) await cp(path.join(root, f), path.join(stage, f));
await cp(path.join(root, 'app'), path.join(stage, 'app'), { recursive: true });
await cp(path.join(root, 'build', 'icon.png'), path.join(stage, 'build', 'icon.png'));
const pkg = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
delete pkg.devDependencies;
delete pkg.scripts;
await writeFile(path.join(stage, 'package.json'), JSON.stringify(pkg, null, 2));

await mkdir(path.dirname(out), { recursive: true });
await createPackage(stage, out);
await rm(stage, { recursive: true, force: true });
console.log(out, ((await stat(out)).size / 1048576).toFixed(1) + ' MB');
