import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { chromium } from "@playwright/test";
import { preview } from "vite";

test("distant and full-column Excel validation stays usable and survives a save and reopen", async (t) => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Input");
  sheet.getCell("A1").value = "Select a colour";
  sheet.getCell("B2").dataValidation = { type: "list", formulae: ['"Red,Blue"'] };
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  const part = "xl/worksheets/sheet1.xml";
  const xml = await zip.file(part).async("string");
  assert.ok(xml.includes('sqref="B2"'));
  zip.file(part, xml.replace('sqref="B2"', 'sqref="B2 D1:D1048576 XFD983086"'));

  const server = await preview({ logLevel: "error", preview: { host: "127.0.0.1", port: 0 },
    build: { outDir: process.env.CWS_DIST_DIR || "dist" } });
  t.after(() => server.close());
  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  const origin = `http://127.0.0.1:${server.httpServer.address().port}`;
  await page.route(/^https?:/, route => new URL(route.request().url()).origin === origin ? route.continue() : route.abort());
  await page.goto(origin);
  async function open(buffer) {
    await page.setInputFiles("#fileInput", { name: "validation.xlsx", mimeType: "application/octet-stream", buffer });
    await page.waitForFunction(() => document.title === "validation.xlsx" && document.querySelector("#fileInput").files.length === 0);
    await page.locator('.sheet-cell[data-address="A1"]').getByText("Select a colour").waitFor();
  }
  await open(await zip.generateAsync({ type: "nodebuffer" }));
  await page.reload();
  const validated = page.locator('.sheet-cell[data-address="B2"]');
  await validated.click();
  await page.keyboard.press("Alt+ArrowDown");
  await page.locator('[data-validation-dropdown-choice="Blue"]').click();
  await page.click('[data-ribbon-tab="file"]');
  await page.click('#fileSideMenu [data-ribbon-command="save-as"]');
  await page.selectOption('[data-data-field="saveAsType"]', "xlsx");
  const download = page.waitForEvent("download");
  await page.click('[data-data-dialog-action="save-as-ok"]');
  const saved = await readFile(await (await download).path());
  const result = await JSZip.loadAsync(saved);
  const savedXml = await result.file(part).async("string");
  const sqrefs = [...savedXml.matchAll(/<dataValidation\b[^>]*\bsqref="([^"]+)"/g)].flatMap(m => m[1].split(/\s+/));
  assert.ok(sqrefs.includes("B2") && sqrefs.includes("XFD983086"));
  const spans = sqrefs.filter(s => /^D\d+(?::D\d+)?$/.test(s)).map(s => s.match(/\d+/g).map(Number)).sort((a,b) => a[0]-b[0]);
  let next = 1;
  for (const [start,end = start] of spans) { assert.equal(start, next); next = end + 1; }
  assert.equal(next, 1048577);
  await open(saved);
});
