import { readdir, readFile } from 'node:fs/promises';
import { resolve, relative, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
const root = fileURLToPath(new URL('../', import.meta.url));
const source = resolve(root, 'cws-agent/api');
const built = resolve(root, 'cws-agent/dist/api');
async function phpFiles(dir) {
  const groups = await Promise.all((await readdir(dir, { withFileTypes: true })).map(async e => {
    const path = join(dir, e.name);
    return e.isDirectory() ? phpFiles(path) : e.isFile() && e.name.endsWith('.php') ? [path] : [];
  }));
  return groups.flat();
}
const files = await phpFiles(source);
const mismatches = [];
for (const file of files) {
  const name = relative(source, file);
  const digest = data => createHash('sha256').update(data).digest('hex');
  try {
    if (digest(await readFile(file)) !== digest(await readFile(join(built, name)))) mismatches.push(name);
  } catch { mismatches.push(name); }
}
if (mismatches.length) {
  console.error(`CWS API build is stale (${mismatches.length} files). Run npm --prefix cws-agent run build before deployment.\n${mismatches.join('\n')}`);
  process.exitCode = 1;
} else console.log(`CWS API build matches all ${files.length} source files.`);
