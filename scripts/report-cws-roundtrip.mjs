#!/usr/bin/env node

import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);
const outPath = consumeOption("--out");
const title = consumeOption("--title") || "CWS HTML Round-Trip Report";
const testId = consumeOption("--test-id") || "manual";
const maxDiffs = consumeOption("--max-diff") || "80";

if (args.includes("--help") || args.includes("-h") || args.length !== 3) {
  console.log(`Usage:
  node scripts/report-cws-roundtrip.mjs [--out report.md] [--title "..."] [--test-id ID] [--max-diff 80] original.html edited.html resaved.html`);
  process.exit(args.length === 3 ? 0 : 2);
}

const [originalPath, editedPath, resavedPath] = args.map((file) => path.resolve(file));
const report = [
  `# ${title}`,
  "",
  `Test ID: ${testId}`,
  "",
  `Generated: ${new Date().toISOString()}`,
  "",
  "## Files",
  "",
  `- Original: \`${relative(originalPath)}\``,
  `- Edited: \`${relative(editedPath)}\``,
  `- Resaved: \`${relative(resavedPath)}\``,
  "",
  "## Exit Code Notes",
  "",
  "- Validation exit code `0` means the file passed validation.",
  "- Diff exit code `0` means no model differences were found.",
  "- Diff exit code `1` means model differences were found and listed below.",
  "",
  "## Validation",
  "",
  validationSection("Original", originalPath),
  validationSection("Edited", editedPath),
  validationSection("Resaved", resavedPath),
  "## Model Diffs",
  "",
  diffSection("Original to Edited", originalPath, editedPath),
  diffSection("Edited to Resaved", editedPath, resavedPath),
  "## Result",
  "",
  "- [ ] Pass",
  "- [ ] Pass with notes",
  "- [ ] Fail",
  "",
  "## Notes",
  "",
  "Add human review notes here.",
  "",
].join("\n");

if (outPath) {
  const target = path.resolve(outPath);
  await mkdir(path.dirname(target), { recursive: true });
  await writeFile(target, report, "utf8");
  console.log(`Wrote ${relative(target)}`);
} else {
  process.stdout.write(report);
}

function consumeOption(name) {
  const index = args.indexOf(name);
  if (index < 0) return "";
  const value = args[index + 1] || "";
  args.splice(index, 2);
  return value;
}

function validationSection(label, filePath) {
  const result = runNodeScript("validate-cws-html.mjs", [filePath]);
  return [
    `### ${label}`,
    "",
    `Exit code: ${result.status}`,
    "",
    "```text",
    trimOutput(result.output) || "(no output)",
    "```",
    "",
  ].join("\n");
}

function diffSection(label, beforePath, afterPath) {
  const result = runNodeScript("diff-cws-model.mjs", [beforePath, afterPath, "--max", maxDiffs]);
  return [
    `### ${label}`,
    "",
    `Exit code: ${result.status}`,
    "",
    "```text",
    trimOutput(result.output) || "(no output)",
    "```",
    "",
  ].join("\n");
}

function runNodeScript(scriptName, scriptArgs) {
  const scriptPath = path.join(projectRoot, "scripts", scriptName);
  const result = spawnSync(process.execPath, [scriptPath, ...scriptArgs], {
    cwd: projectRoot,
    encoding: "utf8",
    maxBuffer: 1024 * 1024 * 20,
  });
  return {
    status: result.status ?? 1,
    output: [result.stdout || "", result.stderr || ""].filter(Boolean).join("\n"),
  };
}

function trimOutput(value) {
  return String(value || "").trim();
}

function relative(filePath) {
  return path.relative(projectRoot, filePath) || filePath;
}
