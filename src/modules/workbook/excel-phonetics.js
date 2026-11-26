import { excelCellAddress } from "./export-model-and-measurement.js";
import { sharedStringBaseText } from "./ooxml-drawings.js";
import { firstXmlChild, normalizePhoneticRuns, parseXmlDocument, xmlChildren } from "./tables-and-validation.js";

export async function patchExcelPhonetics(zip, model) {
  const sheets = model.sheets.map(sheet => new Map(Object.values(sheet.cells || {})
    .filter(cell => cell.phonetics?.length)
    .map(cell => [excelCellAddress(cell.row, cell.col), cell])));
  if (!sheets.some(cells => cells.size)) return;

  const path = "xl/sharedStrings.xml";
  const file = zip.file(path);
  if (!file) return; // ExcelJS emits shared strings for editable text cells.
  const doc = parseXmlDocument(await file.async("text"));
  const root = doc.documentElement;
  const templates = xmlChildren(root, "si");
  const variants = new Map();
  const serializer = new XMLSerializer();
  const stylesPath = "xl/styles.xml";
  const styles = parseXmlDocument(await zip.file(stylesPath).async("text"));
  const fonts = firstXmlChild(styles.documentElement, "fonts");
  const fontIds = new Map(xmlChildren(fonts, "font").map((font, i) => [serializer.serializeToString(font), i]));
  const originalFontCount = fonts.childElementCount;
  let nextIndex = templates.length;

  // ExcelJS coalesces identical text, even when different cells have different readings.
  // Assign a separate string entry to each variant, including the unannotated variant.
  for (const [index, cells] of sheets.entries()) {
    const sheetPath = 'xl/worksheets/sheet' + (index + 1) + '.xml';
    const worksheet = parseXmlDocument(await zip.file(sheetPath).async("text"));
    let changed = false;
    for (const ref of worksheet.querySelectorAll('sheetData c[t="s"]')) {
      const value = firstXmlChild(ref, "v");
      const original = Number(value.textContent);
      const cell = cells.get(ref.getAttribute("r"));
      const template = templates[original];
      const runs = cell ? normalizePhoneticRuns(cell.phonetics, sharedStringBaseText(template)) : [];
      const signature = runs.length ? JSON.stringify([runs, cell.phoneticProperties]) : "";
      if (!variants.has(original)) variants.set(original, new Map());
      const indices = variants.get(original);
      if (!indices.has(signature)) {
        const first = !indices.size;
        const id = first ? original : nextIndex++;
        indices.set(signature, id);
        if (!first || runs.length) {
          const entry = template.cloneNode(true);
          if (runs.length) appendPhoneticRuns(entry, runs, cell.phoneticProperties, fonts, fontIds);
          if (first) root.replaceChild(entry, template);
          else root.appendChild(entry);
        }
      }
      if (indices.get(signature) !== original) {
        value.textContent = String(indices.get(signature));
        changed = true;
      }
      if (runs.length && cell.showPhonetic) {
        ref.setAttribute("ph", "1");
        changed = true;
      }
    }
    if (changed) zip.file(sheetPath, serializer.serializeToString(worksheet));
  }
  root.setAttribute("uniqueCount", String(nextIndex));
  zip.file(path, serializer.serializeToString(doc));
  if (fonts.childElementCount !== originalFontCount) {
    fonts.setAttribute("count", String(fonts.childElementCount));
    zip.file(stylesPath, serializer.serializeToString(styles));
  }
}

function appendPhoneticRuns(entry, runs, properties, fonts, fontIds) {
  const create = name => entry.ownerDocument.createElementNS(entry.namespaceURI, name);
  for (const { start, end, text } of runs) {
    const run = create("rPh");
    run.setAttribute("sb", String(start));
    run.setAttribute("eb", String(end));
    const reading = create("t");
    if (text.trim() !== text) reading.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    reading.textContent = text;
    run.appendChild(reading);
    entry.appendChild(run);
  }
  if (!properties) return;
  const props = create("phoneticPr");
  const font = properties.fontXml;
  if (font && !fontIds.has(font)) {
    const element = parseXmlDocument(font).documentElement;
    if (element.localName === "font" && element.namespaceURI === fonts.namespaceURI) {
      fontIds.set(font, fonts.childElementCount);
      fonts.appendChild(fonts.ownerDocument.importNode(element, true));
    }
  }
  props.setAttribute("fontId", String(fontIds.get(font) ?? 0));
  if (properties.type) props.setAttribute("type", properties.type);
  if (properties.alignment) props.setAttribute("alignment", properties.alignment);
  entry.appendChild(props);
}
