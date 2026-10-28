// Кладёт build\icon.ico в Folio.exe; новая копия — в dist\next (её ставит tools\swap-update.ps1,
// пока Folio закрыт). Так иконку можно сменить без полной пересборки Electron.
import * as ResEdit from 'resedit';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import path from 'node:path';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, '$1')), '..');
const src = path.join(root, 'dist', 'Folio-win32-x64', 'Folio.exe');
const out = path.join(root, 'dist', 'next', 'Folio-win32-x64', 'Folio.exe');

const exe = ResEdit.NtExecutable.from(await readFile(src), { ignoreCert: true });
const res = ResEdit.NtExecutableResource.from(exe);
const ico = ResEdit.Data.IconFile.from(await readFile(path.join(root, 'build', 'icon.ico')));
const groups = ResEdit.Resource.IconGroupEntry.fromEntries(res.entries);
if (!groups.length) throw new Error('в Folio.exe нет иконки');
for (const g of groups) {
  ResEdit.Resource.IconGroupEntry.replaceIconsForResource(res.entries, g.id, g.lang, ico.icons.map((i) => i.data));
}
res.outputResource(exe);
await mkdir(path.dirname(out), { recursive: true });
await writeFile(out, Buffer.from(exe.generate()));
console.log(out, groups.map((g) => g.id + '/' + g.lang).join(', '));
