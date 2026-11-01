import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const sourceRoots = ['src', 'scripts', 'cws-agent/src', 'cws-agent/api', 'cws-agent/scripts', 'cws-agent/test'];
const extensions = new Set(['.js', '.mjs', '.cjs', '.ts', '.tsx', '.jsx', '.css', '.scss', '.php', '.py', '.sh']);
const files = [];
async function walk(directory) {
  for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
    const filename = path.join(directory, entry.name);
    if (entry.isDirectory()) await walk(filename);
    else if (entry.isFile() && extensions.has(path.extname(filename))) {
      const source = await fs.readFile(filename, 'utf8');
      const lines = source.length ? source.split('\n').length - Number(source.endsWith('\n')) : 0;
      files.push({ file: path.relative(root, filename), lines });
    }
  }
}
for (const directory of sourceRoots) await walk(path.join(root, directory));
files.sort((a, b) => b.lines - a.lines);
const oversized = files.filter(file => file.lines >= 10000);
if (oversized.length) {
  console.error('Authored source files must contain fewer than 10,000 lines:', oversized);
  process.exitCode = 1;
} else {
  console.log(`Source size check passed: ${files.length} files; largest ${files[0].file} (${files[0].lines} lines).`);
}
if (process.argv.includes('--report')) console.log(JSON.stringify(files, null, 2));
