#!/usr/bin/env node

import { readFile } from "node:fs/promises";
import path from "node:path";
import { gunzipSync } from "node:zlib";
import { createHash } from "node:crypto";

const args = process.argv.slice(2);
const maxIndex = args.indexOf("--max");
let maxDiffs = 200;
if (maxIndex >= 0) {
  maxDiffs = Math.max(1, Number(args[maxIndex + 1]) || maxDiffs);
  args.splice(maxIndex, 2);
}

if (args.includes("--help") || args.includes("-h") || args.length !== 2) {
  console.log(`Usage:
  node scripts/diff-cws-model.mjs [--max 200] before.html after.html`);
  process.exit(args.length === 2 ? 0 : 2);
}

const [beforePath, afterPath] = args.map((file) => path.resolve(file));
const before = await readCwsModel(beforePath);
const after = await readCwsModel(afterPath);
const diffs = [];
diffValue(before, after, "$", diffs, maxDiffs);

if (!diffs.length) {
  console.log("No model differences.");
  process.exit(0);
}

console.log(`Model differences: ${diffs.length}${diffs.length >= maxDiffs ? " (truncated)" : ""}`);
diffs.forEach((diff) => {
  console.log(`${diff.kind} ${diff.path}`);
  if (diff.kind !== "added") console.log(`  - ${summarizeValue(diff.before)}`);
  if (diff.kind !== "removed") console.log(`  + ${summarizeValue(diff.after)}`);
});
process.exit(1);

async function readCwsModel(filePath) {
  const source = await readFile(filePath, "utf8");
  const script = findCwsModelScript(source);
  if (!script) throw new Error(`Cannot find script#websheet-model in ${filePath}.`);
  const parsed = JSON.parse(script.body.trim());
  if (parsed && typeof parsed === "object" && parsed.encoding === "gzip-base64" && typeof parsed.payload === "string") {
    return JSON.parse(gunzipSync(Buffer.from(parsed.payload, "base64")).toString("utf8"));
  }
  return parsed;
}

function findCwsModelScript(source) {
  const html = String(source || "");
  const masked = maskHtmlComments(html);
  const match = masked.match(/<script\b(?=[^>]*\bid=["']websheet-model["'])(?=[^>]*\btype=["']application\/json["'])[^>]*>([\s\S]*?)<\/script\s*>/i);
  if (!match) return null;
  return { body: html.slice(match.index, match.index + match[0].length).replace(/^<script\b[^>]*>/i, "").replace(/<\/script\s*>$/i, "") };
}

function maskHtmlComments(source) {
  return String(source || "").replace(/<!--[\s\S]*?-->/g, (comment) => " ".repeat(comment.length));
}

function diffValue(before, after, currentPath, diffs, maxDiffs) {
  if (diffs.length >= maxDiffs) return;
  if (Object.is(before, after)) return;
  if (typeof before !== typeof after || before == null || after == null || typeof before !== "object" || typeof after !== "object") {
    diffs.push({ kind: "changed", path: currentPath, before, after });
    return;
  }
  if (Array.isArray(before) || Array.isArray(after)) {
    if (!Array.isArray(before) || !Array.isArray(after)) {
      diffs.push({ kind: "changed", path: currentPath, before, after });
      return;
    }
    const length = Math.max(before.length, after.length);
    for (let index = 0; index < length && diffs.length < maxDiffs; index += 1) {
      if (index >= before.length) {
        diffs.push({ kind: "added", path: `${currentPath}[${index}]`, after: after[index] });
      } else if (index >= after.length) {
        diffs.push({ kind: "removed", path: `${currentPath}[${index}]`, before: before[index] });
      } else {
        diffValue(before[index], after[index], `${currentPath}[${index}]`, diffs, maxDiffs);
      }
    }
    return;
  }
  const keys = [...new Set([...Object.keys(before), ...Object.keys(after)])].sort();
  for (const key of keys) {
    if (diffs.length >= maxDiffs) return;
    const childPath = /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(key)
      ? `${currentPath}.${key}`
      : `${currentPath}[${JSON.stringify(key)}]`;
    if (!(key in before)) {
      diffs.push({ kind: "added", path: childPath, after: after[key] });
    } else if (!(key in after)) {
      diffs.push({ kind: "removed", path: childPath, before: before[key] });
    } else {
      diffValue(before[key], after[key], childPath, diffs, maxDiffs);
    }
  }
}

function summarizeValue(value) {
  if (typeof value === "string") {
    if (value.length > 120) {
      return `<string length=${value.length} sha256=${sha256(value)}>`;
    }
    return JSON.stringify(value);
  }
  if (value == null || typeof value !== "object") return JSON.stringify(value);
  const json = JSON.stringify(value);
  if (json.length > 160) return `<${Array.isArray(value) ? "array" : "object"} length=${json.length} sha256=${sha256(json)}>`;
  return json;
}

function sha256(value) {
  return createHash("sha256").update(String(value)).digest("hex").slice(0, 16);
}
