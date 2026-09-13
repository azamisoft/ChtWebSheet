import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { chromium } from "@playwright/test";
import ExcelJS from "exceljs";
import JSZip from "jszip";

const appUrl = process.env.WEBSHEET_URL || "http://127.0.0.1:5173/";
const browser = await chromium.launch({
  headless: true,
  executablePath: process.env.CHROME_PATH || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
});

function noteText(note) {
  return typeof note === "string" ? note : (note?.texts || []).map((run) => run.text || "").join("");
}

function workbookNotes(workbook) {
  return workbook.worksheets.map((sheet) => {
    const notes = [];
    sheet.eachRow({ includeEmpty: true }, (row) => row.eachCell({ includeEmpty: true }, (cell) => {
      if (cell.note) notes.push({ row: cell.row, col: cell.col, text: noteText(cell.note) });
    }));
    return { name: sheet.name, notes };
  });
}

async function makeFixture(page, { prefixed = false, comments = true } = {}) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Notes");
  sheet.mergeCells("A1:C1");
  sheet.getCell("A1").value = "Comment import regression";
  sheet.getCell("A1").font = { bold: true };
  sheet.getCell("B2").value = 7;
  sheet.getCell("C2").value = { formula: "B2*2", result: 14 };
  sheet.getCell("B3").value = "Annotated value";
  if (comments) {
    sheet.getCell("A1").note = "A note on a merged cell";
    sheet.getCell("B3").note = { texts: [{ text: "A formatted ", font: { bold: true } }, { text: "note" }] };
  }
  const imageId = workbook.addImage({
    base64: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aE8sAAAAASUVORK5CYII=",
    extension: "png",
  });
  sheet.addImage(imageId, { tl: { col: 3, row: 1 }, ext: { width: 24, height: 24 } });
  workbook.addWorksheet("Plain").getCell("A1").value = "No comments";
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  const relsPath = "xl/worksheets/_rels/sheet1.xml.rels";
  const relsXml = await page.evaluate((xml) => {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    // ExcelJS reconciles VML only when the comments relationship comes first.
    const comment = [...doc.documentElement.children].find((node) => node.getAttribute("Type").endsWith("/comments"));
    if (comment) doc.documentElement.prepend(comment);
    return new XMLSerializer().serializeToString(doc);
  }, await zip.file(relsPath).async("text"));
  zip.file(relsPath, relsXml);
  if (prefixed) {
    const sheetPath = "xl/worksheets/sheet1.xml";
    const xml = await page.evaluate((source) => {
      const doc = new DOMParser().parseFromString(source, "application/xml");
      const namespace = doc.documentElement.namespaceURI;
      for (const node of [...doc.getElementsByTagNameNS(namespace, "*")].reverse()) {
        const replacement = doc.createElementNS(namespace, `x:${node.localName}`);
        for (const attr of node.attributes) replacement.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
        while (node.firstChild) replacement.appendChild(node.firstChild);
        node.replaceWith(replacement);
      }
      return new XMLSerializer().serializeToString(doc);
    }, await zip.file(sheetPath).async("text"));
    zip.file(sheetPath, xml);
  }
  return { buffer: await zip.generateAsync({ type: "nodebuffer" }), expected: workbookNotes(workbook) };
}

async function checkImport(page, { buffer, expected, name, fallback = false, synthetic = false }) {
  const warnings = [];
  const errors = [];
  const onConsole = (message) => {
    if (message.type() === "warning") warnings.push(message.text());
    if (message.type() === "error" && !message.text().includes("404")) errors.push(message.text());
  };
  const onError = (error) => errors.push(error.message);
  page.on("console", onConsole);
  page.on("pageerror", onError);
  try {
    const imported = await page.evaluate(async ({ base64, name }) => {
      const bytes = Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
      await window.__webSheetDevLoadBlob(new Blob([bytes]), name);
      return window.__webSheetDevWorkbookState();
    }, { base64: buffer.toString("base64"), name });
    assert.equal(imported.sourceName, name, "The requested workbook must finish importing");
    for (const sheet of expected) {
      for (const note of sheet.notes) {
        const cell = await page.evaluate(({ name, row, col }) => window.__webSheetDevCellModel(name, row, col), { name: sheet.name, ...note });
        assert.equal(cell?.note, note.text, "Imported note text must be preserved");
      }
    }
    if (synthetic) {
      const cells = await page.evaluate(() => ({
        title: window.__webSheetDevCellModel("Notes", 1, 1),
        input: window.__webSheetDevCellModel("Notes", 2, 2),
        formula: window.__webSheetDevCellModel("Notes", 2, 3),
        display: window.__webSheetDevCellDisplay("Notes", 2, 3),
        images: window.__webSheetDevSheetModel("Notes").images.length,
      }));
      assert.equal(cells.title.raw, "Comment import regression");
      assert.equal(cells.title.css.fontWeight, "700");
      assert.equal(cells.input.raw, 7);
      assert.match(cells.formula.formula, /B2\*2/);
      assert.equal(String(cells.display), "14");
      assert.equal(cells.images, 1, "Drawing extraction must still preserve images");
    }
    const fallbackWarnings = warnings.filter((warning) => warning.includes("falling back to prepared workbook buffer"));
    assert.equal(fallbackWarnings.length, fallback ? 1 : 0, "Check both the fast and fallback import paths");
    assert.deepEqual(errors, [], "Import must not report browser errors");
    const exported = await page.evaluate(() => window.__webSheetDevExportXlsxBase64());
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(Buffer.from(exported, "base64"));
    const savedNotes = workbookNotes(reopened);
    assert.deepEqual(savedNotes.map((sheet) => sheet.name), expected.map((sheet) => sheet.name));
    for (let index = 0; index < expected.length; index += 1) {
      const actual = savedNotes[index].notes;
      assert.deepEqual(actual.map(({ row, col }) => ({ row, col })), expected[index].notes.map(({ row, col }) => ({ row, col })));
      expected[index].notes.forEach((note, noteIndex) => {
        // CWS export adds an author heading before the existing note body.
        assert.ok(actual[noteIndex].text === note.text || actual[noteIndex].text.endsWith(`:\n${note.text}`), "Saved note body must be preserved");
      });
    }
    console.log(JSON.stringify({ case: name, sheets: expected.length, notes: expected.reduce((count, sheet) => count + sheet.notes.length, 0), fallback, passed: true }));
  } finally {
    page.off("console", onConsole);
    page.off("pageerror", onError);
  }
}

try {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
  await page.goto(appUrl, { waitUntil: "load" });
  await page.waitForFunction(() => window.__webSheetDevLoadBlob && window.__webSheetDevWorkbookState);
  for (const options of [{}, { prefixed: true }, { comments: false }]) {
    await checkImport(page, {
      ...await makeFixture(page, options),
      name: options.prefixed ? "comments-fallback.xlsx" : options.comments === false ? "no-comments.xlsx" : "comments-fast.xlsx",
      fallback: Boolean(options.prefixed),
      synthetic: true,
    });
  }
  for (const filePath of process.argv.slice(2)) {
    const buffer = await readFile(filePath);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    await checkImport(page, { buffer, expected: workbookNotes(workbook), name: path.basename(filePath) });
  }
} finally {
  await browser.close();
}
