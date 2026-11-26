import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { JSDOM } from "jsdom";
import { chromium } from "@playwright/test";
import { preview } from "vite";
import { hasJapaneseEraTokens } from "../../src/modules/workbook/japanese-era-format.js";

test("era detection respects locale, escaped or quoted text, and numeric exponents", () => {
  for (const format of ["General", "0.00E+00", "0.0e-00", '"ggge"yyyy', "\\g yyyy", "[$-409]e", "[Red]0.0"]) {
    assert.equal(hasJapaneseEraTokens(format), false, format);
  }
  for (const format of ["[$-411]ee", "ggge/m/d", "[$-ja-JP]gge"]) {
    assert.equal(hasJapaneseEraTokens(format), true, format);
  }
});

function xml(text) {
  return new JSDOM(text, { contentType: "text/xml" }).window.document;
}

const dateCases = [
  ["1989-01-07", '[$-411]ggge"年"m"月"d"日";@', "昭和64年1月7日"],
  ["1989-01-08", '[$-411]ggge"年"m"月"d"日";@', "平成1年1月8日"],
  ["1993-04-09", '[$-411]ggge"年"m"月"d"日";@', "平成5年4月9日"],
  ["2019-04-30", "[$-411]gee.mm.dd", "H31.04.30"],
  ["2019-05-01", "[$-411]ggee.mm.dd", "令01.05.01"],
  ["2026-01-02", 'yyyy"年"m"月";@', "2026年1月"],
  ["2026-01-02", '"～"yyyy"/"mm', "～2026/01"],
  ["2026-01-02T13:07:05Z", 'yyyy-mm-dd hh:mm:ss', "2026-01-02 13:07:05"],
  ["2026-01-02", '"ggge "yyyy/mm/dd', "ggge 2026/01/02"],
];

async function fixture() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Test");
  workbook.addWorksheet("Plain").getCell("A1").value = "日本";
  for (const addr of ["A1", "B1", "C1"]) sheet.getCell(addr).value = "日本";
  for (const [index, [iso, format]] of dateCases.entries()) {
    const cell = sheet.getCell(index + 2, 1);
    cell.value = new Date(iso).getTime() / 86400000 + 25569;
    cell.numFmt = format;
  }
  sheet.getCell("A11").value = 1.5;
  sheet.getCell("A11").numFmt = "[h]:mm";
  sheet.getColumn(1).width = 32;
  // Keep the sample cells in the DOM as well as the canvas for rendering assertions.
  sheet.views = [{ state: "frozen", xSplit: 3 }];

  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  const path = "xl/worksheets/sheet1.xml";
  let sheetXml = await zip.file(path).async("string");
  for (const [addr, reading] of [["A1", "ニホン"], ["B1", "ニッポン"]]) {
    const inline = '<is><t>日本</t><rPh sb="0" eb="2"><t>' + reading +
      '</t></rPh><phoneticPr fontId="0" type="fullwidthKatakana" alignment="center"/></is>';
    const pattern = new RegExp('<c r="' + addr + '" t="s"><v>\\d+</v></c>');
    assert.match(sheetXml, pattern);
    sheetXml = sheetXml.replace(pattern, '<c r="' + addr + '" t="inlineStr"' + (addr === "A1" ? ' ph="1"' : '') + '>' + inline + '</c>');
  }
  zip.file(path, sheetXml);
  return zip.generateAsync({ type: "nodebuffer" });
}

test("Japanese era dates and independent phonetic guides survive Excel Save As and reopen", async t => {
  const server = await preview({ logLevel: "error", preview: { host: "127.0.0.1", port: 0 },
    build: { outDir: process.env.CWS_DIST_DIR || "dist" } });
  t.after(() => server.close());
  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 1300, height: 900 } });
  page.setDefaultTimeout(4500);
  const origin = "http://127.0.0.1:" + server.httpServer.address().port;
  await page.route(/^https?:/, r => new URL(r.request().url()).origin === origin ? r.continue() : r.abort());
  await page.goto(origin);
  async function open(buffer, name) {
    await page.setInputFiles("#fileInput", { name, mimeType: "application/octet-stream", buffer });
    await page.waitForFunction(n => document.title === n && !document.querySelector("#fileInput").files.length, name);
    await page.locator(".sheet-tab").filter({ hasText: /^Test$/ }).click();
    for (const [index, [, , expected]] of dateCases.entries()) {
      const address = "A" + (index + 2);
      assert.equal(await page.locator('.sheet-cell[data-address="' + address + '"] .cell-content').textContent(), expected, address);
    }
    assert.equal(await page.locator('.sheet-cell[data-address="A11"] .cell-content').textContent(), "36:00");
    for (const addr of ["A1", "B1", "C1"]) {
      const cell = page.locator('.sheet-cell[data-address="' + addr + '"]');
      assert.equal(await cell.locator(".cell-content").textContent(), "日本");
      assert.equal(await cell.evaluate(el => el.classList.contains("has-phonetic")), addr !== "C1");
      assert.equal(await cell.evaluate(el => el.classList.contains("show-phonetic")), addr === "A1");
    }
  }
  await open(await fixture(), "source.xlsx");
  await page.click('[data-ribbon-tab="file"]');
  await page.click('#fileSideMenu [data-ribbon-command="save-as"]');
  await page.selectOption('[data-data-field="saveAsType"]', "xlsx");
  const download = page.waitForEvent("download");
  await page.click('[data-data-dialog-action="save-as-ok"]');
  const saved = await readFile(await (await download).path());
  const result = await JSZip.loadAsync(saved);
  const read = async path => xml(await result.file(path).async("string"));
  const strings = (await read("xl/sharedStrings.xml")).querySelectorAll("sst > si");
  const fonts = (await read("xl/styles.xml")).querySelectorAll("fonts > font");
  for (const [i, readings] of [[1, { A1: "ニホン", B1: "ニッポン", C1: null }], [2, { A1: null }]]) {
    const sheet = await read("xl/worksheets/sheet" + i + ".xml");
    for (const [address, expected] of Object.entries(readings)) {
      const cell = sheet.querySelector('c[r="' + address + '"]');
      const entry = strings[Number(cell.querySelector("v").textContent)];
      assert.equal(entry.querySelector("rPh > t")?.textContent ?? null, expected);
      if (expected) {
        const run = entry.querySelector("rPh"), props = entry.querySelector("phoneticPr");
        assert.deepEqual([run.getAttribute("sb"), run.getAttribute("eb")], ["0", "2"]);
        assert.equal(props.getAttribute("alignment"), "center");
        assert.equal(props.getAttribute("type"), "fullwidthKatakana");
        assert.ok(fonts[Number(props.getAttribute("fontId"))]);
      }
    }
  }
  await open(saved, "roundtrip.xlsx");
});
