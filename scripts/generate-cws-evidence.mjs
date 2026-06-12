#!/usr/bin/env node

import { chromium } from "@playwright/test";
import ExcelJS from "exceljs";
import { access, mkdir, readFile, writeFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { gunzipSync } from "node:zlib";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);
const DEFAULT_APP_URL = process.env.WEBSHEET_URL || "http://127.0.0.1:5173/";
const DEFAULT_CHROME_PATH = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const MODEL_SCRIPT_RE = /<script\b(?=[^>]*\bid=["']websheet-model["'])(?=[^>]*\btype=["']application\/json["'])([^>]*)>([\s\S]*?)<\/script\s*>/i;

const help = consumeFlag("--help") || consumeFlag("-h");
const appUrl = consumeOption("--app-url") || DEFAULT_APP_URL;
const chromePath = consumeOption("--chrome-path") || process.env.CHROME_PATH || DEFAULT_CHROME_PATH;
const workbookOption = consumeOption("--workbook");
const outDir = path.resolve(consumeOption("--out-dir") || "docs/evidence/generated/real-export-smoke");
const testId = consumeOption("--test-id") || "real-export-smoke";

if (help) {
  printHelp();
  process.exit(0);
}

await mkdir(outDir, { recursive: true });

const fixtureWorkbookPath = workbookOption
  ? path.resolve(workbookOption)
  : path.join(outDir, "source.xlsx");
if (workbookOption) {
  await access(fixtureWorkbookPath);
} else {
  await writeFixtureWorkbook(fixtureWorkbookPath);
}

await waitForApp(appUrl);

const exportedPath = path.join(outDir, "exported.cws.html");
const editedPath = path.join(outDir, "edited.cws.html");
const resavedPath = path.join(outDir, "resaved.cws.html");
const reportPath = path.join(outDir, "roundtrip-report.md");
const summaryPath = path.join(outDir, "summary.json");

const browser = await chromium.launch(await launchOptions(chromePath));
try {
  const context = await browser.newContext({
    viewport: { width: 1280, height: 720 },
    acceptDownloads: true,
    permissions: ["clipboard-read", "clipboard-write"],
  });

  await exportWorkbookFromApp(context, fixtureWorkbookPath, exportedPath);
  await writeAiEditedHtml(exportedPath, editedPath);
  const resaveInfo = await resaveStandaloneHtml(context, editedPath, resavedPath);
  await context.close();

  const exportedModel = await readCwsModel(exportedPath);
  const editedModel = await readCwsModel(editedPath);
  const resavedModel = await readCwsModel(resavedPath);
  const assertions = evidenceAssertions(exportedModel, editedModel, resavedModel, resaveInfo);
  const validation = {
    exported: runNodeScript("validate-cws-html.mjs", [exportedPath]),
    edited: runNodeScript("validate-cws-html.mjs", [editedPath]),
    resaved: runNodeScript("validate-cws-html.mjs", [resavedPath]),
  };
  const report = runNodeScript("report-cws-roundtrip.mjs", [
    "--out",
    reportPath,
    "--title",
    "CWS HTML Real Export Smoke Round-Trip Report",
    "--test-id",
    testId,
    exportedPath,
    editedPath,
    resavedPath,
  ]);

  const summary = {
    testId,
    generatedAt: new Date().toISOString(),
    appUrl,
    files: {
      sourceWorkbook: relative(fixtureWorkbookPath),
      exported: relative(exportedPath),
      edited: relative(editedPath),
      resaved: relative(resavedPath),
      report: relative(reportPath),
    },
    resave: resaveInfo,
    validation,
    report,
    assertions,
  };
  await writeFile(summaryPath, `${JSON.stringify(summary, null, 2)}\n`, "utf8");

  const failures = [
    ...Object.entries(validation)
      .filter(([, result]) => result.status !== 0)
      .map(([name, result]) => `${name} validation failed with exit code ${result.status}`),
    ...(report.status === 0 ? [] : [`round-trip report generation failed with exit code ${report.status}`]),
    ...assertions.filter((item) => !item.ok).map((item) => item.message),
  ];
  if (failures.length) {
    throw new Error(`CWS evidence generation failed:\n- ${failures.join("\n- ")}`);
  }

  console.log(JSON.stringify({
    ok: true,
    files: summary.files,
    summary: relative(summaryPath),
  }, null, 2));
} finally {
  await browser.close();
}

function printHelp() {
  console.log(`Usage:
  node scripts/generate-cws-evidence.mjs [--app-url http://127.0.0.1:5173/] [--workbook file.xlsx] [--out-dir docs/evidence/generated/real-export-smoke]

This command opens Cht WebSheet in a running dev server, imports an xlsx workbook,
saves it as CWS HTML, applies a controlled AI-style JSON edit to
script#websheet-model, reopens/resaves the edited HTML, validates all files, and
writes a round-trip evidence report.`);
}

function consumeFlag(name) {
  const index = args.indexOf(name);
  if (index < 0) return false;
  args.splice(index, 1);
  return true;
}

function consumeOption(name) {
  const index = args.indexOf(name);
  if (index < 0) return "";
  const value = args[index + 1] || "";
  args.splice(index, 2);
  return value;
}

async function launchOptions(candidateChromePath) {
  const options = { headless: true };
  if (candidateChromePath) {
    try {
      await access(candidateChromePath);
      options.executablePath = candidateChromePath;
    } catch {
      // Fall back to the browser bundled or configured for Playwright.
    }
  }
  return options;
}

async function waitForApp(url, timeoutMs = 30000) {
  const deadline = Date.now() + timeoutMs;
  let lastError = null;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(url, { redirect: "manual" });
      if (response.ok || (response.status >= 300 && response.status < 400)) return;
      lastError = new Error(`HTTP ${response.status}`);
    } catch (error) {
      lastError = error;
    }
    await delay(500);
  }
  throw new Error(`Cannot reach Cht WebSheet at ${url}. Start it with: npm run dev\n${lastError?.message || ""}`);
}

async function writeFixtureWorkbook(filePath) {
  await mkdir(path.dirname(filePath), { recursive: true });
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "CWS HTML evidence generator";
  workbook.created = new Date("2026-06-12T00:00:00.000Z");
  const sheet = workbook.addWorksheet("Evidence");
  sheet.mergeCells("A1:D1");
  sheet.getCell("A1").value = "CWS HTML evidence workbook";
  sheet.getCell("A1").font = { bold: true, size: 14, color: { argb: "FF1F2937" } };
  sheet.getCell("A1").alignment = { horizontal: "center" };
  sheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9EAF7" } };
  sheet.columns = [
    { width: 20 },
    { width: 16 },
    { width: 16 },
    { width: 14 },
    { width: 10 },
  ];
  sheet.addRows([
    ["Task", "Owner", "Status", "Score"],
    ["Draft profile", "CWS", "Todo", 72],
    ["Collect evidence", "CWS", "Todo", 18],
    ["Total", "", "", { formula: "SUM(D3:D4)", result: 90 }],
  ]);
  sheet.getRow(2).font = { bold: true };
  sheet.getCell("C3").dataValidation = {
    type: "list",
    allowBlank: false,
    formulae: ['"Todo,Done"'],
  };
  sheet.autoFilter = "A2:D4";
  sheet.views = [{ state: "frozen", ySplit: 2 }];
  sheet.getCell("A6").value = "The evidence generator edits C3 inside script#websheet-model.";
  sheet.getCell("A6").font = { italic: true, color: { argb: "FF475569" } };
  const imageId = workbook.addImage({
    base64: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
    extension: "png",
  });
  sheet.addImage(imageId, {
    tl: { col: 4, row: 1 },
    ext: { width: 48, height: 48 },
  });
  await workbook.xlsx.writeFile(filePath);
}

async function exportWorkbookFromApp(context, workbookPath, outputPath) {
  const page = await context.newPage();
  const messages = [];
  page.on("console", (message) => {
    if (["error", "warning"].includes(message.type())) messages.push(`${message.type()}: ${message.text()}`);
  });
  page.on("pageerror", (error) => messages.push(`pageerror: ${error.stack || error.message}`));
  try {
    await page.goto(withCacheBuster(appUrl, "cws-evidence"), { waitUntil: "load" });
    await page.setInputFiles("#fileInput", workbookPath);
    await page.waitForFunction(
      (name) => {
        const status = document.querySelector("#statusText")?.textContent || "";
        const state = window.__webSheetDevWorkbookState?.() || {};
        return (
          document.title.includes(name) &&
          state.activeSheetName &&
          /読み込みました|読み込み済み|已加载/.test(status)
        );
      },
      path.basename(workbookPath),
      { timeout: 30000 },
    );
    const downloadPromise = page.waitForEvent("download", { timeout: 30000 });
    await page.keyboard.press("Control+S");
    const download = await downloadPromise;
    await download.saveAs(outputPath);
    const html = await readFile(outputPath, "utf8");
    if (!looksLikeCwsHtml(html)) {
      throw new Error(`Saved file does not look like CWS HTML: ${download.suggestedFilename()}`);
    }
  } catch (error) {
    if (messages.length) {
      error.message = `${error.message}\nBrowser messages:\n${messages.slice(-20).join("\n")}`;
    }
    throw error;
  } finally {
    await page.close();
  }
}

async function writeAiEditedHtml(beforePath, afterPath) {
  const source = await readFile(beforePath, "utf8");
  const model = decodeCwsModelFromHtml(source);
  const sheet = model.sheets?.[0];
  if (!sheet?.cells) throw new Error("Cannot find first sheet cells in exported CWS HTML.");
  const statusCell = sheet.cells["3:3"] || { row: 3, col: 3 };
  sheet.cells["3:3"] = {
    ...statusCell,
    row: 3,
    col: 3,
    raw: "Done",
    display: "Done",
  };
  model.sourceName = path.basename(afterPath);
  model.saveType = "html";
  model.cwsEvidence = {
    testId,
    purpose: "AI-style edit: only script#websheet-model was changed.",
    expectedPreservation: "This unknown top-level field should remain after standalone reopen/save.",
  };
  await writeFile(afterPath, replaceCwsModel(source, model), "utf8");
}

async function resaveStandaloneHtml(context, inputPath, outputPath) {
  const page = await context.newPage();
  const messages = [];
  page.on("console", (message) => {
    if (["error", "warning"].includes(message.type())) messages.push(`${message.type()}: ${message.text()}`);
  });
  page.on("pageerror", (error) => messages.push(`pageerror: ${error.stack || error.message}`));
  let localSaveResolve;
  const localSavePromise = new Promise((resolve) => {
    localSaveResolve = resolve;
  });
  await page.route("**/*", async (route) => {
    const url = route.request().url();
    if (/^http:\/\/(?:127\.0\.0\.1|localhost):\d+\/__websheet\/local-save/i.test(url)) {
      let payload = {};
      try {
        payload = JSON.parse(route.request().postData() || "{}");
      } catch {
        payload = {};
      }
      localSaveResolve({
        mode: "local-save-bridge",
        path: payload.path || "",
        fileName: payload.fileName || "",
        hasHtml: typeof payload.html === "string" && looksLikeCwsHtml(payload.html),
        html: payload.html || "",
      });
      await route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ ok: true }) });
      return;
    }
    await route.continue();
  });
  try {
    await page.goto(pathToFileURL(inputPath).href, { waitUntil: "load" });
    await page.waitForFunction(
      () => document.querySelector("#sheetTabs .sheet-tab")?.textContent?.trim(),
      null,
      { timeout: 30000 },
    );
    await page.evaluate(() => {
      window.__cwsEvidencePickerCalls = [];
      Object.defineProperty(window, "showSaveFilePicker", {
        configurable: true,
        value: async (options = {}) => {
          window.__cwsEvidencePickerCalls.push({ suggestedName: String(options.suggestedName || "") });
          throw new DOMException("Picker is disabled during evidence generation", "AbortError");
        },
      });
    });
    const downloadPromise = page.waitForEvent("download", { timeout: 15000 })
      .then((download) => ({ mode: "download", download }))
      .catch(() => null);
    await page.keyboard.press("Control+S");
    const saveResult = await Promise.race([
      localSavePromise,
      downloadPromise,
      delay(16000).then(() => null),
    ]);
    if (saveResult?.mode === "local-save-bridge" && saveResult.html) {
      await writeFile(outputPath, saveResult.html, "utf8");
      const pickerCalls = await page.evaluate(() => window.__cwsEvidencePickerCalls || []);
      return {
        mode: saveResult.mode,
        targetPath: saveResult.path,
        fileName: saveResult.fileName,
        pickerCalls,
      };
    }
    if (saveResult?.mode === "download" && saveResult.download) {
      await saveResult.download.saveAs(outputPath);
      const pickerCalls = await page.evaluate(() => window.__cwsEvidencePickerCalls || []);
      return {
        mode: "download",
        suggestedFilename: saveResult.download.suggestedFilename(),
        pickerCalls,
      };
    }
    throw new Error("Standalone edited HTML did not produce a local-save request or a download.");
  } catch (error) {
    if (messages.length) {
      error.message = `${error.message}\nBrowser messages:\n${messages.slice(-20).join("\n")}`;
    }
    throw error;
  } finally {
    await page.close();
  }
}

function evidenceAssertions(exportedModel, editedModel, resavedModel, resaveInfo) {
  return [
    {
      ok: exportedModel?.sheets?.[0]?.cells?.["3:3"]?.raw !== "Done",
      message: "Exported model should start before the AI-style edit.",
    },
    {
      ok: editedModel?.sheets?.[0]?.cells?.["3:3"]?.raw === "Done",
      message: "Edited model should set Evidence!C3 to Done.",
    },
    {
      ok: resavedModel?.sheets?.[0]?.cells?.["3:3"]?.raw === "Done",
      message: "Resaved model should preserve the AI-style cell edit.",
    },
    {
      ok: Boolean(resavedModel?.cwsEvidence?.expectedPreservation),
      message: "Resaved model should preserve the unknown top-level cwsEvidence field.",
    },
    {
      ok: ["local-save-bridge", "download"].includes(resaveInfo?.mode || ""),
      message: "Standalone edited HTML should be saveable after reopening.",
    },
  ];
}

function looksLikeCwsHtml(source) {
  const html = String(source || "");
  return (
    html.includes('data-cws-format="CWS_HTML"') &&
    html.includes('id="websheet-model"') &&
    html.includes('name="cws:schema"')
  );
}

async function readCwsModel(filePath) {
  return decodeCwsModelFromHtml(await readFile(filePath, "utf8"));
}

function decodeCwsModelFromHtml(source) {
  const script = findCwsModelScript(source);
  if (!script) throw new Error("Cannot find script#websheet-model.");
  const parsed = JSON.parse(script.body.trim());
  if (parsed && typeof parsed === "object" && parsed.encoding === "gzip-base64" && typeof parsed.payload === "string") {
    return JSON.parse(gunzipSync(Buffer.from(parsed.payload, "base64")).toString("utf8"));
  }
  return parsed;
}

function replaceCwsModel(source, model) {
  const html = String(source || "");
  const script = findCwsModelScript(html);
  if (!script) throw new Error("Cannot find script#websheet-model.");
  return `${html.slice(0, script.index)}<script${script.attrs}>${safeJson(model)}</script>${html.slice(script.end)}`;
}

function findCwsModelScript(source) {
  const html = String(source || "");
  const masked = maskHtmlComments(html);
  const match = MODEL_SCRIPT_RE.exec(masked);
  if (!match) return null;
  const index = match.index;
  const end = index + match[0].length;
  const full = html.slice(index, end);
  const openEnd = html.indexOf(">", index);
  const closeTag = full.match(/<\/script\s*>$/i)?.[0] || "</script>";
  const closeStart = end - closeTag.length;
  return {
    index,
    end,
    attrs: full.match(/^<script\b([^>]*)>/i)?.[1] || "",
    body: html.slice(openEnd + 1, closeStart),
  };
}

function maskHtmlComments(source) {
  return String(source || "").replace(/<!--[\s\S]*?-->/g, (comment) => " ".repeat(comment.length));
}

function safeJson(value) {
  return JSON.stringify(value, null, 2)
    .replace(/</g, "\\u003c")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
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
    output: [result.stdout || "", result.stderr || ""].filter(Boolean).join("\n").trim(),
  };
}

function withCacheBuster(url, key) {
  const separator = String(url).includes("?") ? "&" : "?";
  return `${url}${separator}${encodeURIComponent(key)}=${Date.now()}`;
}

function relative(filePath) {
  return path.relative(projectRoot, filePath) || filePath;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
