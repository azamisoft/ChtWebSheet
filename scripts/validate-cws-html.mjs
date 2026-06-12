#!/usr/bin/env node

import { readFile, readdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import Ajv2020 from "ajv/dist/2020.js";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const CWS_HTML_FORMAT = "CWS_HTML";
const CWS_SCHEMA_VERSION = "1";
const MODEL_SCRIPT_ID = "websheet-model";
const AI_INSTRUCTIONS_SCRIPT_ID = "cws-ai-instructions";
const schemaPath = path.join(projectRoot, "public", "schema", "cws-html-workbook-model-v1.schema.json");
const validateWorkbookSchema = createSchemaValidator(JSON.parse(await readFile(schemaPath, "utf8")));

const args = process.argv.slice(2);
const expectInvalid = consumeFlag("--expect-invalid");
const validateExamples = consumeFlag("--examples");

if (args.includes("--help") || args.includes("-h")) {
  printHelp();
  process.exit(0);
}

let failures = 0;

if (validateExamples) {
  const validDir = path.join(projectRoot, "examples", "cws-html", "valid");
  const invalidDir = path.join(projectRoot, "examples", "cws-html", "invalid");
  const validFiles = await listHtmlFiles(validDir);
  const invalidFiles = await listHtmlFiles(invalidDir);
  failures += await validateFiles(validFiles, { expectInvalid: false });
  failures += await validateFiles(invalidFiles, { expectInvalid: true });
} else {
  const files = args.map((file) => path.resolve(file));
  if (!files.length) {
    printHelp();
    process.exit(2);
  }
  failures += await validateFiles(files, { expectInvalid });
}

process.exit(failures ? 1 : 0);

function consumeFlag(flag) {
  const index = args.indexOf(flag);
  if (index < 0) return false;
  args.splice(index, 1);
  return true;
}

function printHelp() {
  console.log(`Usage:
  node scripts/validate-cws-html.mjs --examples
  node scripts/validate-cws-html.mjs <file.html> [...]
  node scripts/validate-cws-html.mjs --expect-invalid <file.html> [...]`);
}

async function listHtmlFiles(dir) {
  const entries = await readdir(dir, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...await listHtmlFiles(fullPath));
    } else if (/\.(html?|cws\.html)$/i.test(entry.name)) {
      files.push(fullPath);
    }
  }
  return files.sort();
}

async function validateFiles(files, options = {}) {
  let failed = 0;
  for (const file of files) {
    const source = await readFile(file, "utf8");
    const result = validateCwsHtml(source);
    const displayPath = path.relative(projectRoot, file) || file;
    const isInvalid = result.errors.length > 0;
    if (options.expectInvalid) {
      if (isInvalid) {
        console.log(`OK invalid ${displayPath}`);
      } else {
        console.error(`FAIL ${displayPath}: expected invalid, but no errors were found`);
        failed += 1;
      }
      continue;
    }
    if (isInvalid) {
      console.error(`FAIL ${displayPath}`);
      result.errors.forEach((error) => console.error(`  - ${error}`));
      failed += 1;
    } else {
      console.log(`OK ${displayPath}`);
      result.warnings.forEach((warning) => console.warn(`  warning: ${warning}`));
    }
  }
  return failed;
}

function validateCwsHtml(source) {
  const errors = [];
  const warnings = [];
  const html = String(source || "");

  const htmlAttrs = parseFirstTagAttributes(html, "html");
  if (!htmlAttrs) {
    errors.push("Missing html root element.");
  } else {
    requireAttr(errors, htmlAttrs, "data-cws-format", CWS_HTML_FORMAT, "html root");
    requireAttr(errors, htmlAttrs, "data-cws-schema-version", CWS_SCHEMA_VERSION, "html root");
    requireAttr(errors, htmlAttrs, "data-cws-app", "CWS", "html root");
  }

  const metaTags = parseMetaTags(html);
  requireMeta(errors, metaTags, "cws:format", CWS_HTML_FORMAT);
  requireMeta(errors, metaTags, "cws:schema-version", CWS_SCHEMA_VERSION);
  requireMeta(errors, metaTags, "cws:guide");
  requireMeta(errors, metaTags, "cws:schema");

  const modelScripts = parseScripts(html).filter((script) => script.attrs.id === MODEL_SCRIPT_ID);
  if (modelScripts.length !== 1) {
    errors.push(`Expected exactly one script#${MODEL_SCRIPT_ID}; found ${modelScripts.length}.`);
    return { errors, warnings };
  }

  const modelScript = modelScripts[0];
  requireAttr(errors, modelScript.attrs, "type", "application/json", `script#${MODEL_SCRIPT_ID}`);
  requireAttr(errors, modelScript.attrs, "data-cws-model", "workbook", `script#${MODEL_SCRIPT_ID}`);
  requireAttr(errors, modelScript.attrs, "data-cws-format", CWS_HTML_FORMAT, `script#${MODEL_SCRIPT_ID}`);
  requireAttr(errors, modelScript.attrs, "data-cws-schema-version", CWS_SCHEMA_VERSION, `script#${MODEL_SCRIPT_ID}`);
  requireAttr(errors, modelScript.attrs, "data-cws-editable", "true", `script#${MODEL_SCRIPT_ID}`);

  const aiInstructionScripts = parseScripts(html).filter((script) => script.attrs.id === AI_INSTRUCTIONS_SCRIPT_ID);
  if (aiInstructionScripts.length !== 1) {
    warnings.push(`Expected one script#${AI_INSTRUCTIONS_SCRIPT_ID}; found ${aiInstructionScripts.length}.`);
  } else {
    const aiJson = parseJson(aiInstructionScripts[0].body);
    if (!aiJson.ok) warnings.push(`script#${AI_INSTRUCTIONS_SCRIPT_ID} is not valid JSON: ${aiJson.message}`);
  }

  const parsedModel = parseJson(modelScript.body);
  if (!parsedModel.ok) {
    errors.push(`script#${MODEL_SCRIPT_ID} is not valid JSON: ${parsedModel.message}`);
    return { errors, warnings };
  }

  validateWorkbookModelSchema(parsedModel.value, errors);
  validateWorkbookModel(parsedModel.value, errors, warnings);
  return { errors, warnings };
}

function createSchemaValidator(schema) {
  const ajv = new Ajv2020({ allErrors: true, strict: false });
  return ajv.compile(schema);
}

function validateWorkbookModelSchema(model, errors) {
  const ok = validateWorkbookSchema(model);
  if (ok) return;
  (validateWorkbookSchema.errors || []).slice(0, 20).forEach((error) => {
    errors.push(`Schema ${error.instancePath || "/"} ${error.message || "validation failed"}.`);
  });
  if ((validateWorkbookSchema.errors || []).length > 20) {
    errors.push(`Schema validation found ${(validateWorkbookSchema.errors || []).length - 20} more error(s).`);
  }
}

function parseFirstTagAttributes(source, tagName) {
  const re = new RegExp(`<${tagName}\\b([^>]*)>`, "i");
  const match = source.match(re);
  return match ? parseAttributes(match[1]) : null;
}

function parseMetaTags(source) {
  const tags = [];
  const re = /<meta\b([^>]*)>/gi;
  let match;
  while ((match = re.exec(source))) {
    tags.push(parseAttributes(match[1]));
  }
  return tags;
}

function parseScripts(source) {
  const scripts = [];
  const html = stripHtmlComments(source);
  const re = /<script\b([^>]*)>([\s\S]*?)<\/script\s*>/gi;
  let match;
  while ((match = re.exec(html))) {
    scripts.push({
      attrs: parseAttributes(match[1]),
      body: match[2],
    });
  }
  return scripts;
}

function stripHtmlComments(source) {
  return String(source || "").replace(/<!--[\s\S]*?-->/g, "");
}

function parseAttributes(text) {
  const attrs = {};
  const re = /([^\s=/"'>]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g;
  let match;
  while ((match = re.exec(text || ""))) {
    attrs[match[1].toLowerCase()] = match[2] ?? match[3] ?? match[4] ?? "";
  }
  return attrs;
}

function requireAttr(errors, attrs, name, expected, label) {
  const actual = attrs[String(name).toLowerCase()];
  if (actual == null) {
    errors.push(`Missing ${label} attribute ${name}.`);
    return;
  }
  if (expected != null && String(actual) !== String(expected)) {
    errors.push(`Invalid ${label} attribute ${name}: expected ${expected}, got ${actual}.`);
  }
}

function requireMeta(errors, metaTags, name, expected = null) {
  const tag = metaTags.find((item) => item.name === name);
  if (!tag) {
    errors.push(`Missing meta[name="${name}"].`);
    return;
  }
  if (expected != null && String(tag.content || "") !== String(expected)) {
    errors.push(`Invalid meta[name="${name}"] content: expected ${expected}, got ${tag.content || ""}.`);
  }
}

function parseJson(text) {
  try {
    return { ok: true, value: JSON.parse(String(text || "").trim()) };
  } catch (error) {
    return { ok: false, message: error?.message || String(error) };
  }
}

function validateWorkbookModel(model, errors, warnings) {
  if (!model || typeof model !== "object" || Array.isArray(model)) {
    errors.push("Workbook model must be a JSON object.");
    return;
  }
  if (model.encoding === "gzip-base64" || typeof model.payload === "string" && model.format === "cht-websheet") {
    errors.push("Encoded workbook payloads are not valid CWS HTML 1.0 AI-editable profile models; store the plain workbook object in script#websheet-model.");
    return;
  }
  if (model.version == null) errors.push("Workbook model is missing version.");
  if (typeof model.sourceName !== "string" || !model.sourceName.trim()) errors.push("Workbook model sourceName must be a non-empty string.");
  if (!Array.isArray(model.sheets) || !model.sheets.length) {
    errors.push("Workbook model sheets must be a non-empty array.");
    return;
  }
  if (model.activeSheetIndex != null && (!Number.isInteger(model.activeSheetIndex) || model.activeSheetIndex < 0 || model.activeSheetIndex >= model.sheets.length)) {
    errors.push(`Workbook activeSheetIndex must be an integer from 0 to ${model.sheets.length - 1}.`);
  }

  const imageAssets = validateWorkbookImageAssets(model.assets?.images, errors, warnings);
  const sheetNames = new Set();
  model.sheets.forEach((sheet, index) => {
    const label = `sheets[${index}]`;
    validateSheet(sheet, label, errors, warnings, imageAssets);
    if (sheet && typeof sheet === "object") {
      if (sheetNames.has(sheet.name)) errors.push(`Duplicate sheet name: ${sheet.name}.`);
      sheetNames.add(sheet.name);
    }
  });
}

function validateWorkbookImageAssets(sourceAssets, errors, warnings) {
  const assets = new Map();
  if (sourceAssets == null) return assets;
  if (Array.isArray(sourceAssets)) {
    sourceAssets.forEach((asset, index) => {
      const id = String(asset?.id || `img_${String(index + 1).padStart(4, "0")}`);
      validateImageAsset(asset, `assets.images[${index}]`, errors, warnings);
      assets.set(id, asset);
    });
    return assets;
  }
  if (typeof sourceAssets !== "object") {
    errors.push("assets.images must be an object map or array.");
    return assets;
  }
  Object.entries(sourceAssets).forEach(([id, asset]) => {
    validateImageAsset(asset, `assets.images["${id}"]`, errors, warnings);
    assets.set(String(id), asset);
  });
  return assets;
}

function validateImageAsset(asset, label, errors, warnings) {
  if (!asset || typeof asset !== "object" || Array.isArray(asset)) {
    errors.push(`${label} must be an object.`);
    return;
  }
  if (asset.mimeType != null && (typeof asset.mimeType !== "string" || !asset.mimeType.startsWith("image/"))) {
    errors.push(`${label}.mimeType must begin with image/.`);
  }
  if (asset.dataEncoding != null && asset.dataEncoding !== "gzip-base64") {
    errors.push(`${label}.dataEncoding has unsupported value ${asset.dataEncoding}.`);
  }
  if (typeof asset.data !== "string" || !asset.data.trim()) {
    errors.push(`${label}.data must be a non-empty base64 string.`);
    return;
  }
  if (!isLikelyBase64(asset.data)) errors.push(`${label}.data is not valid-looking base64.`);
  if (asset.dataEncoding === "gzip-base64" && !String(asset.data).startsWith("H4sI")) {
    warnings.push(`${label}.dataEncoding is gzip-base64, but data does not have the common gzip base64 prefix.`);
  }
  if (asset.uncompressedByteLength != null && (!Number.isInteger(asset.uncompressedByteLength) || asset.uncompressedByteLength < 0)) {
    errors.push(`${label}.uncompressedByteLength must be a non-negative integer.`);
  }
}

function validateSheet(sheet, label, errors, warnings, imageAssets = new Map()) {
  if (!sheet || typeof sheet !== "object" || Array.isArray(sheet)) {
    errors.push(`${label} must be an object.`);
    return;
  }
  if (typeof sheet.name !== "string" || !sheet.name.trim()) {
    errors.push(`${label}.name must be a non-empty string.`);
  } else if (sheet.name.length > 31 || /[:\\/?*[\]]/.test(sheet.name)) {
    errors.push(`${label}.name is not a valid worksheet name: ${sheet.name}.`);
  }
  if (!isPositiveInteger(sheet.rowCount)) errors.push(`${label}.rowCount must be a positive integer.`);
  if (!isPositiveInteger(sheet.colCount)) errors.push(`${label}.colCount must be a positive integer.`);
  if (!sheet.cells || typeof sheet.cells !== "object" || Array.isArray(sheet.cells)) {
    errors.push(`${label}.cells must be an object keyed by row:col.`);
    return;
  }

  Object.entries(sheet.cells).forEach(([key, cell]) => {
    const match = key.match(/^([1-9][0-9]*):([1-9][0-9]*)$/);
    if (!match) {
      errors.push(`${label}.cells has invalid key "${key}"; expected row:col.`);
      return;
    }
    if (!cell || typeof cell !== "object" || Array.isArray(cell)) {
      errors.push(`${label}.cells["${key}"] must be an object.`);
      return;
    }
    const row = Number(match[1]);
    const col = Number(match[2]);
    if (cell.row != null && Number(cell.row) !== row) errors.push(`${label}.cells["${key}"].row does not match key.`);
    if (cell.col != null && Number(cell.col) !== col) errors.push(`${label}.cells["${key}"].col does not match key.`);
    if (cell.formula != null && (typeof cell.formula !== "string" || !cell.formula.startsWith("="))) {
      errors.push(`${label}.cells["${key}"].formula must be a string beginning with =.`);
    }
    if (row > sheet.rowCount || col > sheet.colCount) {
      warnings.push(`${label}.cells["${key}"] is outside rowCount/colCount.`);
    }
  });

  validateRanges(sheet.merges, `${label}.merges`, errors);
  validateRange(sheet.autoFilter?.range, `${label}.autoFilter.range`, errors, { optional: true });
  validateRange(sheet.pageLayout?.printArea?.range || sheet.pageLayout?.printArea, `${label}.pageLayout.printArea`, errors, { optional: true });
  validateSheetImages(sheet.images, `${label}.images`, errors, warnings, imageAssets);
}

function validateSheetImages(images, label, errors, warnings, imageAssets) {
  if (images == null) return;
  if (!Array.isArray(images)) {
    errors.push(`${label} must be an array.`);
    return;
  }
  images.forEach((image, index) => {
    const imageLabel = `${label}[${index}]`;
    if (!image || typeof image !== "object" || Array.isArray(image)) {
      errors.push(`${imageLabel} must be an object.`);
      return;
    }
    const assetId = String(image.assetId || "").trim();
    const src = String(image.src || "").trim();
    if (assetId && !imageAssets.has(assetId)) {
      errors.push(`${imageLabel}.assetId references missing assets.images entry: ${assetId}.`);
    }
    if (!assetId && src && !src.startsWith("data:image/") && !/^https?:\/\//i.test(src)) {
      warnings.push(`${imageLabel}.src is neither a data:image URL nor an http(s) URL.`);
    }
    if (!assetId && !src && !image.shape && !image.chart) {
      warnings.push(`${imageLabel} has no src or assetId.`);
    }
  });
}

function validateRanges(ranges, label, errors) {
  if (ranges == null) return;
  if (!Array.isArray(ranges)) {
    errors.push(`${label} must be an array.`);
    return;
  }
  ranges.forEach((range, index) => validateRange(range, `${label}[${index}]`, errors));
}

function validateRange(range, label, errors, options = {}) {
  if (range == null) {
    if (!options.optional) errors.push(`${label} is missing.`);
    return;
  }
  if (!range || typeof range !== "object" || Array.isArray(range)) {
    errors.push(`${label} must be an object.`);
    return;
  }
  for (const key of ["top", "left", "bottom", "right"]) {
    if (!isPositiveInteger(range[key])) errors.push(`${label}.${key} must be a positive integer.`);
  }
  if (isPositiveInteger(range.top) && isPositiveInteger(range.bottom) && range.top > range.bottom) {
    errors.push(`${label}.top must be less than or equal to bottom.`);
  }
  if (isPositiveInteger(range.left) && isPositiveInteger(range.right) && range.left > range.right) {
    errors.push(`${label}.left must be less than or equal to right.`);
  }
}

function isPositiveInteger(value) {
  return Number.isInteger(value) && value >= 1;
}

function isLikelyBase64(value) {
  const compact = String(value || "").replace(/\s+/g, "");
  if (!compact || compact.length % 4 !== 0) return false;
  return /^[A-Za-z0-9+/]*={0,2}$/.test(compact);
}
