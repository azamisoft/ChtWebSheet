// Feature module; shared bindings are imported explicitly.
import {
  CUSTOM_SHAPE_TYPE,
  DEFAULT_SHAPE_TYPE,
  DEFAULT_THEME_COLORS,
  EMU_PER_PIXEL,
  INDEXED_COLORS,
  JSZip,
  SORT_CUSTOM_LISTS,
  state,
} from "../../app-core.js";
import {
  autoFilterDateForSerial,
  autoFilterDateRangeForPreset,
  autoFilterLocalDate,
} from "../data/filter-and-text-tools.js";
import {
  normalizeCustomSortText,
  normalizeSortColor,
} from "../data/sort-and-autofilter.js";
import {
  chartSvg,
  lineBoxAndPointsFromEndpoints,
} from "../drawing/charts-and-insertion.js";
import {
  clampNumber,
  isLineShapeType,
  normalizeLineShapeConnections,
  normalizeRotationDegrees,
  normalizeShapeAdjustments,
  normalizeShapeArrowSize,
  normalizeShapeArrowType,
  normalizeShapeCustomPath,
  normalizeShapeFillPattern,
  normalizeShapeGlow,
  normalizeShapeGradient,
  normalizeShapeLineCap,
  normalizeShapeLineCompound,
  normalizeShapeLinePoints,
  normalizeShapeOutlineWidth,
  normalizeShapeReflection,
  normalizeShapeSoftEdges,
  normalizeShapeTextOutlineWidth,
  normalizeShapeTextRuns,
  shapeDescriptorForType,
} from "../drawing/shape-rendering.js";
import {
  rotatePointAround,
} from "../drawing/transforms-and-connectors.js";
import {
  normalizeColorValue,
  svgDataUrl,
} from "../editing/ink-and-clipboard.js";
import {
  ensureCellModel,
} from "../rendering/grid-and-canvas.js";
import {
  decodeRange,
} from "../ui/help-and-utilities.js";
import {
  clampZoom,
} from "../ui/sheet-protection-and-translation.js";
import {
  rangeFromCellLabel,
} from "./export-model-and-measurement.js";
import {
  applyTint,
  fontFamilyForExcel,
  hexToRgb,
} from "./formulas-and-sheet-tabs.js";
import {
  clampWorkbookSheetIndex,
  normalizeCustomTopLeftCell,
  normalizeCustomViewGuid,
  normalizeWorkbookCalcProperties,
  normalizeWorkbookProperties,
  normalizedSheetShowRuler,
  normalizedSheetViewMode,
} from "./import-model.js";
import {
  autoFilterStdDevCount,
  bytesToBase64,
  clampOutlineLevel,
  customSheetViewNodeToModel,
  customSheetViewPaneFromNode,
  customSheetViewSelectionFromNode,
  customViewScrollFromSheetView,
  excelFitToPageValue,
  firstXmlChild,
  firstXmlChildText,
  imageMimeType,
  normalizeOutlineSettings,
  normalizePackageRelationshipPath,
  normalizePhoneticRuns,
  normalizeWorkbookRelationshipPath,
  parseValidationOptions,
  parseWorksheetPhonetics,
  parseXmlDocument,
  workbookWindowViewFromCustomWorkbookViewNode,
  xmlBoolAttribute,
  xmlChildren,
  xmlElements,
  xmlLocalName,
  xmlTextContent,
} from "./tables-and-validation.js";


function preservedOoxmlEntryPath(path) {
  const normalized = String(path || "").replace(/^\/+/, "");
  return /^xl\/vbaProject(?:Signature(?:Agile)?)?\.bin$/i.test(normalized);
}



function parsePreservedWorkbookRelationships(xml) {
  const doc = parseXmlDocument(xml);
  return xmlElements(doc.documentElement, "Relationship")
    .map((node) => ({
      id: node.getAttribute("Id") || "",
      type: node.getAttribute("Type") || "",
      target: node.getAttribute("Target") || "",
      targetMode: node.getAttribute("TargetMode") || "",
    }))
    .filter((relationship) => /vbaProject/i.test(relationship.type) || /vbaProject/i.test(relationship.target));
}



function parsePreservedContentTypeOverrides(xml) {
  const doc = parseXmlDocument(xml);
  return xmlElements(doc.documentElement, "Override")
    .map((node) => ({
      partName: node.getAttribute("PartName") || "",
      contentType: node.getAttribute("ContentType") || "",
    }))
    .filter((override) => /vbaProject/i.test(override.partName) || /vbaProject/i.test(override.contentType));
}



async function extractWorkbookDefinedNames(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
    return workbookXml ? parseWorkbookDefinedNamesXml(workbookXml) : [];
  } catch (error) {
    console.warn("Failed to parse workbook defined names.", error);
    return [];
  }
}



function parseWorkbookDefinedNamesXml(workbookXml) {
  const doc = parseXmlDocument(workbookXml);
  return xmlElements(doc.documentElement, "definedName")
    .map((node) => {
      const name = String(node.getAttribute("name") || "").trim();
      const refersTo = String(node.textContent || "").trim();
      if (!name || !refersTo) return null;
      const localSheetIdText = node.getAttribute("localSheetId");
      const localSheetId = localSheetIdText == null || localSheetIdText === "" ? null : Number.parseInt(localSheetIdText, 10);
      const isSheetScoped = Number.isInteger(localSheetId) && localSheetId >= 0;
      return {
        name,
        scope: isSheetScoped ? "sheet" : "workbook",
        sheetIndex: isSheetScoped ? localSheetId : null,
        refersTo: refersTo.startsWith("=") ? refersTo : `=${refersTo}`,
        comment: node.getAttribute("comment") || "",
        hidden:
          String(node.getAttribute("hidden") || "").toLowerCase() === "1" ||
          String(node.getAttribute("hidden") || "").toLowerCase() === "true" ||
          name.startsWith("_xlnm."),
      };
    })
    .filter(Boolean);
}



async function extractWorkbookAutoFilters(buffer, themeColors = DEFAULT_THEME_COLORS) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const [workbookXml, workbookRelsXml, stylesXml] = await Promise.all([
      zip.file("xl/workbook.xml")?.async("text"),
      zip.file("xl/_rels/workbook.xml.rels")?.async("text"),
      zip.file("xl/styles.xml")?.async("text"),
    ]);
    if (!workbookXml || !workbookRelsXml) return {};
    const styleDxfs = parseWorkbookStyleDxfsXml(stylesXml, themeColors);
    const sheetPaths = workbookSheetPaths(workbookXml, workbookRelsXml);
    const entries = await Promise.all(
      sheetPaths.map(async ({ name, path }) => {
        const worksheetXml = await zip.file(path)?.async("text");
        const worksheetState = parseWorksheetAutoFilterXml(worksheetXml, styleDxfs);
        const tableStates = await parseWorksheetTableStates(zip, path, worksheetXml, styleDxfs);
        const tableState = preferredImportedTableAutoFilterState(tableStates);
        const autoFilter = mergeImportedAutoFilterStates(worksheetState, tableState);
        const tables = tableStates;
        return {
          name,
          autoFilter: {
            ...(autoFilter || {}),
            ...(tables.length ? { tables } : {}),
          },
        };
      }),
    );
    return Object.fromEntries(
      entries
        .filter((entry) => entry.autoFilter?.range || entry.autoFilter?.sortState || entry.autoFilter?.tables?.length)
        .map((entry) => [entry.name, entry.autoFilter]),
    );
  } catch (error) {
    console.warn("Failed to parse workbook AutoFilter criteria.", error);
    return {};
  }
}



async function extractWorkbookDataValidationDetails(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const [workbookXml, workbookRelsXml] = await Promise.all([
      zip.file("xl/workbook.xml")?.async("text"),
      zip.file("xl/_rels/workbook.xml.rels")?.async("text"),
    ]);
    if (!workbookXml || !workbookRelsXml) return {};
    const entries = await Promise.all(
      workbookSheetPaths(workbookXml, workbookRelsXml).map(async ({ name, path }) => ({
        name,
        details: parseWorksheetDataValidationDetailsXml(await zip.file(path)?.async("text")),
      })),
    );
    return Object.fromEntries(entries.filter((entry) => entry.details.length).map((entry) => [entry.name, entry.details]));
  } catch (error) {
    console.warn("Failed to parse workbook data validation details.", error);
    return {};
  }
}



function parseWorksheetDataValidationDetailsXml(xml) {
  if (!xml) return [];
  const doc = parseXmlDocument(xml);
  return xmlElements(doc.documentElement, "dataValidation")
    .map((node) => {
      const sqref = String(node.getAttribute("sqref") || "").trim();
      if (!sqref) return null;
      const showDropDown = parseDataValidationShowDropDownAttribute(node.getAttribute("showDropDown"));
      const imeMode = String(node.getAttribute("imeMode") || "").trim();
      const validation = dataValidationRuleFromXmlElement(node, { showDropDown, imeMode });
      const details = {
        ranges: sqref
          .split(/\s+/)
          .map((ref) => decodeRange(ref.replace(/\$/g, "")) || rangeFromCellLabel(ref))
          .filter(Boolean),
      };
      if (showDropDown != null) details.showDropDown = showDropDown;
      if (imeMode) details.imeMode = imeMode;
      if (validation) details.validation = validation;
      const hasDetails =
        Object.prototype.hasOwnProperty.call(details, "showDropDown") ||
        Object.prototype.hasOwnProperty.call(details, "imeMode") ||
        Object.prototype.hasOwnProperty.call(details, "validation");
      return hasDetails && details.ranges.length ? details : null;
    })
    .filter(Boolean);
}



function parseDataValidationShowDropDownAttribute(value) {
  if (value == null || value === "") return null;
  const text = String(value).toLowerCase();
  if (text === "1" || text === "true") return false;
  if (text === "0" || text === "false") return true;
  return null;
}



async function extractWorkbookCellHyperlinks(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const [workbookXml, workbookRelsXml] = await Promise.all([
      zip.file("xl/workbook.xml")?.async("text"),
      zip.file("xl/_rels/workbook.xml.rels")?.async("text"),
    ]);
    if (!workbookXml || !workbookRelsXml) return {};
    const entries = await Promise.all(
      workbookSheetPaths(workbookXml, workbookRelsXml).map(async ({ name, path }) => ({
        name,
        hyperlinks: await extractWorksheetCellHyperlinks(zip, path),
      })),
    );
    return Object.fromEntries(entries.filter((entry) => entry.hyperlinks.length).map((entry) => [entry.name, entry.hyperlinks]));
  } catch (error) {
    console.warn("Failed to extract workbook hyperlinks.", error);
    return {};
  }
}



async function extractWorksheetCellHyperlinks(zip, worksheetPath) {
  const worksheetXml = await zip.file(worksheetPath)?.async("text");
  if (!worksheetXml || !/<hyperlink\b/i.test(worksheetXml)) return [];
  const worksheetDoc = parseXmlDocument(worksheetXml);
  const relationships = await worksheetRelationshipTargets(zip, worksheetPath);
  return xmlElements(worksheetDoc.documentElement, "hyperlink")
    .map((node) => worksheetHyperlinkToModel(node, relationships))
    .filter(Boolean);
}



async function worksheetRelationshipTargets(zip, worksheetPath) {
  const parts = String(worksheetPath || "").split("/");
  const fileName = parts.pop();
  if (!fileName) return new Map();
  const relsPath = `${parts.join("/")}/_rels/${fileName}.rels`.replace(/^\/+/, "");
  const relsXml = await zip.file(relsPath)?.async("text");
  if (!relsXml) return new Map();
  const relsDoc = parseXmlDocument(relsXml);
  return new Map(
    xmlChildren(relsDoc.documentElement, "Relationship")
      .filter((rel) => /\/hyperlink$/i.test(rel.getAttribute("Type") || ""))
      .map((rel) => [rel.getAttribute("Id") || "", rel.getAttribute("Target") || ""])
      .filter(([id, target]) => id && target),
  );
}



function worksheetHyperlinkToModel(node, relationships) {
  const ref = String(node.getAttribute("ref") || "").trim();
  if (!ref) return null;
  const relationshipId =
    node.getAttribute("r:id") ||
    node.getAttribute("id") ||
    node.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") ||
    "";
  const relationshipTarget = relationshipId ? relationships.get(relationshipId) || "" : "";
  const location = String(node.getAttribute("location") || "").trim();
  const hyperlink = normalizeWorksheetHyperlinkTarget(relationshipTarget, location);
  if (!hyperlink) return null;
  return {
    ref,
    hyperlink,
    display: node.getAttribute("display") || "",
    tooltip: node.getAttribute("tooltip") || "",
  };
}



function normalizeWorksheetHyperlinkTarget(target, location) {
  const targetText = String(target || "").trim();
  const locationText = String(location || "").trim();
  if (!targetText && locationText) return locationText.startsWith("#") ? locationText : `#${locationText}`;
  if (!targetText) return "";
  if (!locationText) return targetText;
  if (targetText.startsWith("#")) return targetText;
  if (/^(?:https?|ftp|file):/i.test(targetText)) return `${targetText}#${locationText.replace(/^#/, "")}`;
  return targetText;
}



function applyWorkbookCellHyperlinksToModel(model, workbookCellHyperlinks = {}) {
  if (!model?.sheets?.length || !workbookCellHyperlinks || typeof workbookCellHyperlinks !== "object") return;
  model.sheets.forEach((sheet) => {
    const hyperlinks = workbookCellHyperlinks[sheet.name] || [];
    if (!hyperlinks.length) return;
    hyperlinks.forEach((hyperlink) => applyWorksheetCellHyperlinkToModel(sheet, hyperlink));
  });
}



function applyWorksheetCellHyperlinkToModel(sheet, hyperlink) {
  const range = decodeRange(String(hyperlink?.ref || "").replace(/\$/g, "")) || rangeFromCellLabel(hyperlink?.ref);
  if (!range || !hyperlink?.hyperlink) return;
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      const cell = ensureCellModel(sheet, row, col);
      cell.hyperlink = hyperlink.hyperlink;
      const display = String(hyperlink.display || "").trim();
      if (display && cell.raw == null && !cell.formula) {
        cell.raw = display;
        cell.cached = display;
        cell.display = display;
      }
    }
  }
}



function dataValidationRuleFromXmlElement(node, options = {}) {
  const type = String(node?.getAttribute("type") || "any").trim() || "any";
  if (type === "none" || type === "any") return null;
  const formula1 = normalizeDataValidationFormulaFromXml(type, String(firstXmlChild(node, "formula1")?.textContent || "").trim());
  const formula2 = normalizeDataValidationFormulaFromXml(type, String(firstXmlChild(node, "formula2")?.textContent || "").trim());
  const base = {
    type,
    allowBlank: xmlBooleanAttribute(node, "allowBlank"),
    showInputMessage: xmlBooleanAttribute(node, "showInputMessage"),
    inputTitle: node.getAttribute("promptTitle") || "",
    inputMessage: node.getAttribute("prompt") || "",
    showErrorMessage: xmlBooleanAttribute(node, "showErrorMessage", true),
    errorStyle: node.getAttribute("errorStyle") || "stop",
    errorTitle: node.getAttribute("errorTitle") || "",
    errorMessage: node.getAttribute("error") || "",
    ...(options.imeMode ? { imeMode: options.imeMode } : {}),
  };
  if (type === "list") {
    return {
      ...base,
      formula: formula1,
      options: parseValidationOptions(formula1),
      showDropDown: options.showDropDown == null ? true : options.showDropDown !== false,
    };
  }
  return {
    ...base,
    operator: type === "custom" ? "" : node.getAttribute("operator") || "between",
    formulae: [formula1, formula2].filter((formula) => formula !== ""),
  };
}



function normalizeDataValidationFormulaFromXml(type, formula) {
  const text = String(formula || "").trim();
  if (!text || text.startsWith("=")) return text;
  if (type === "list") {
    if (isQuotedValidationLiteral(text)) return text;
    return `=${text}`;
  }
  if (type === "custom") return `=${text}`;
  return shouldTreatDataValidationFormulaAsReference(text) ? `=${text}` : text;
}



function isQuotedValidationLiteral(text) {
  const value = String(text || "").trim();
  return (value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"));
}



function shouldTreatDataValidationFormulaAsReference(text) {
  const value = String(text || "").trim();
  if (!value || /^-?\d+(?:\.\d+)?$/.test(value)) return false;
  if (/^\d{1,2}:\d{2}(?::\d{2})?(?:\s*(?:AM|PM))?$/i.test(value)) return false;
  if (/^\d{4}[/-]\d{1,2}[/-]\d{1,2}$/.test(value)) return false;
  return /[A-Za-z_]|[!:()]/.test(value);
}



function parseWorksheetAutoFilterXml(xml, styleDxfs = []) {
  if (!xml) return null;
  const doc = parseXmlDocument(xml);
  const autoFilter = xmlElements(doc.documentElement, "autoFilter")[0];
  const state = parseAutoFilterStateFromElement(autoFilter, styleDxfs);
  const worksheetSortState = parseWorksheetSortStateXml(xmlElements(doc.documentElement, "sortState")[0], state?.range || null, styleDxfs);
  if (!state?.range && !state?.sortState && !worksheetSortState) return null;
  return {
    ...(state || {}),
    ...(worksheetSortState && !state?.sortState ? { sortState: worksheetSortState } : {}),
  };
}



function parseAutoFilterStateFromElement(autoFilter, styleDxfs = []) {
  if (!autoFilter) return null;
  const range = decodeRange(String(autoFilter.getAttribute("ref") || "").replace(/\$/g, ""));
  const criteria = {};
  const hiddenButtons = {};
  if (autoFilter && range) {
    xmlChildren(autoFilter, "filterColumn").forEach((columnNode) => {
      const colId = Number.parseInt(columnNode.getAttribute("colId") || "", 10);
      if (!Number.isInteger(colId) || colId < 0) return;
      const absoluteCol = range.left + colId;
      const showButtonAttr = columnNode.getAttribute("showButton");
      if (xmlBooleanAttribute(columnNode, "hiddenButton") || showButtonAttr === "0" || String(showButtonAttr || "").toLowerCase() === "false") {
        hiddenButtons[absoluteCol] = true;
      }
      const criterion = parseAutoFilterColumnCriterion(columnNode, styleDxfs);
      if (criterion) criteria[absoluteCol] = criterion;
    });
  }
  const sortState = parseWorksheetSortStateXml(firstXmlChild(autoFilter, "sortState"), range, styleDxfs);
  if (!range && !sortState) return null;
  return {
    ...(range
      ? {
          range: { sheetIndex: 0, ...range },
          criteria,
          ...(Object.keys(hiddenButtons).length ? { hiddenButtons } : {}),
        }
      : {}),
    ...(sortState ? { sortState } : {}),
  };
}



async function parseWorksheetTableStates(zip, worksheetPath, worksheetXml, styleDxfs = []) {
  if (!worksheetXml) return [];
  const relsXml = await zip.file(worksheetRelsPath(worksheetPath))?.async("text");
  if (!relsXml) return [];
  const doc = parseXmlDocument(worksheetXml);
  const rels = worksheetRelationships(relsXml, worksheetPath);
  const tablePaths = xmlElements(doc.documentElement, "tablePart")
    .map((node) => node.getAttribute("r:id") || node.getAttribute("id") || node.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"))
    .map((relId) => rels.get(relId))
    .filter(Boolean);
  const tableStates = await Promise.all(
    tablePaths.map(async (tablePath, index) => parseTableStateXml(await zip.file(tablePath)?.async("text"), index, tablePath, styleDxfs)),
  );
  return tableStates.filter(Boolean);
}



function preferredImportedTableAutoFilterState(tableStates) {
  return (
    tableStates.find((table) => table?.autoFilterState?.range && Object.keys(table.autoFilterState.criteria || {}).length) ||
    tableStates.find((table) => table?.autoFilterState?.range) ||
    null
  )?.autoFilterState || null;
}



function parseTableStateXml(tableXml, index = 0, tablePath = "", styleDxfs = []) {
  if (!tableXml) return null;
  const doc = parseXmlDocument(tableXml);
  const table = doc.documentElement;
  const tableRange = decodeRange(String(table.getAttribute("ref") || "").replace(/\$/g, ""));
  if (!tableRange) return null;
  const styleInfoNode = firstXmlChild(table, "tableStyleInfo");
  const columnsNode = firstXmlChild(table, "tableColumns");
  const columns = xmlChildren(columnsNode, "tableColumn").map((columnNode, columnIndex) => parseTableColumnXml(columnNode, columnIndex));
  const autoFilterState = parseAutoFilterStateFromElement(firstXmlChild(table, "autoFilter"), styleDxfs) || {
    range: { sheetIndex: 0, ...tableRange },
    criteria: {},
  };
  const tableId = table.getAttribute("id") || stableObjectId(tablePath) || String(index + 1);
  const name = table.getAttribute("displayName") || table.getAttribute("name") || `Table${index + 1}`;
  const totalsRowCount = Number.parseInt(table.getAttribute("totalsRowCount") || "", 10);
  const totalsRowShown = table.getAttribute("totalsRowShown");
  const headerRowCount = Number.parseInt(table.getAttribute("headerRowCount") || "1", 10);
  return {
    id: `table-${stableObjectId(`${tableId}-${name}`) || index + 1}`,
    name,
    displayName: table.getAttribute("displayName") || name,
    ref: table.getAttribute("ref") || "",
    range: { sheetIndex: 0, ...tableRange },
    showHeaderRow: headerRowCount !== 0,
    showTotalsRow: Number.isInteger(totalsRowCount) && totalsRowCount > 0,
    totalsRowShown: totalsRowShown == null ? "" : totalsRowShown,
    showFilterButton: Boolean(firstXmlChild(table, "autoFilter")),
    style: styleInfoNode?.getAttribute("name") || "",
    styleInfo: parseTableStyleInfoXml(styleInfoNode),
    columns,
    autoFilterState,
  };
}



function parseTableColumnXml(columnNode, index = 0) {
  const formulaNode = firstXmlChild(columnNode, "calculatedColumnFormula");
  const totalFormulaNode = firstXmlChild(columnNode, "totalsRowFormula");
  return {
    id: columnNode?.getAttribute("id") || String(index + 1),
    name: columnNode?.getAttribute("name") || `Column${index + 1}`,
    uniqueName: columnNode?.getAttribute("uniqueName") || "",
    totalsRowFunction: columnNode?.getAttribute("totalsRowFunction") || "",
    totalsRowLabel: columnNode?.getAttribute("totalsRowLabel") || "",
    calculatedColumnFormula: formulaNode?.textContent || "",
    totalsRowFormula: totalFormulaNode?.textContent || "",
  };
}



function parseTableStyleInfoXml(styleInfoNode) {
  if (!styleInfoNode) {
    return {
      name: "",
      showFirstColumn: false,
      showLastColumn: false,
      showRowStripes: true,
      showColumnStripes: false,
    };
  }
  return {
    name: styleInfoNode.getAttribute("name") || "",
    showFirstColumn: xmlBooleanAttribute(styleInfoNode, "showFirstColumn"),
    showLastColumn: xmlBooleanAttribute(styleInfoNode, "showLastColumn"),
    showRowStripes: xmlBooleanAttribute(styleInfoNode, "showRowStripes", true),
    showColumnStripes: xmlBooleanAttribute(styleInfoNode, "showColumnStripes"),
  };
}



function mergeImportedAutoFilterStates(worksheetState, tableState) {
  const worksheetHasCriteria = Boolean(worksheetState?.range && Object.keys(worksheetState.criteria || {}).length);
  if (worksheetHasCriteria || worksheetState?.sortState) return worksheetState;
  const tableHasCriteria = Boolean(tableState?.range && Object.keys(tableState.criteria || {}).length);
  const tableHasSortState = Boolean(tableState?.sortState?.range);
  if (tableHasCriteria) return tableState;
  if (tableState?.range && !tableHasSortState) return tableState;
  return worksheetState || null;
}



function parseWorksheetSortStateXml(sortStateNode, autoFilterRange = null, styleDxfs = []) {
  if (!sortStateNode) return null;
  const rawRange = decodeRange(String(sortStateNode.getAttribute("ref") || "").replace(/\$/g, ""));
  const columnSort = xmlBooleanAttribute(sortStateNode, "columnSort");
  const range = normalizeImportedSortRange(rawRange, autoFilterRange, columnSort);
  if (!range) return null;
  const keys = xmlChildren(sortStateNode, "sortCondition")
    .map((node) => parseWorksheetSortConditionXml(node, columnSort, styleDxfs))
    .filter(Boolean);
  if (!keys.length) return null;
  return {
    orientation: columnSort ? "columns" : "rows",
    range: { sheetIndex: 0, ...range.range },
    hasHeader: range.hasHeader,
    caseSensitive: xmlBooleanAttribute(sortStateNode, "caseSensitive"),
    keys,
  };
}



function normalizeImportedSortRange(rawRange, autoFilterRange = null, columnSort = false) {
  if (!rawRange && !autoFilterRange) return null;
  if (!rawRange) return { range: autoFilterRange, hasHeader: Boolean(autoFilterRange) && !columnSort };
  if (!columnSort && autoFilterRange && importedSortBodyMatchesFilter(rawRange, autoFilterRange)) {
    return { range: autoFilterRange, hasHeader: true };
  }
  if (columnSort && autoFilterRange && importedColumnSortBodyMatchesFilter(rawRange, autoFilterRange)) {
    return { range: autoFilterRange, hasHeader: true };
  }
  return { range: rawRange, hasHeader: false };
}



function importedSortBodyMatchesFilter(sortRange, filterRange) {
  return Boolean(
    sortRange &&
      filterRange &&
      sortRange.top === filterRange.top + 1 &&
      sortRange.bottom === filterRange.bottom &&
      sortRange.left === filterRange.left &&
      sortRange.right === filterRange.right,
  );
}



function importedColumnSortBodyMatchesFilter(sortRange, filterRange) {
  return Boolean(
    sortRange &&
      filterRange &&
      sortRange.left === filterRange.left + 1 &&
      sortRange.right === filterRange.right &&
      sortRange.top === filterRange.top &&
      sortRange.bottom === filterRange.bottom,
  );
}



function parseWorksheetSortConditionXml(node, columnSort = false, styleDxfs = []) {
  const refRange = decodeRange(String(node?.getAttribute("ref") || "").replace(/\$/g, ""));
  if (!refRange) return null;
  const direction = xmlBooleanAttribute(node, "descending") ? "desc" : "asc";
  const sortBy = String(node.getAttribute("sortBy") || "").trim();
  const color = importedSortConditionColor(node, sortBy, styleDxfs);
  const customList = importedSortConditionCustomList(node?.getAttribute("customList") || "");
  const key = columnSort ? { row: refRange.top, direction } : { col: refRange.left, direction };
  if (sortBy === "cellColor") key.sortOn = "cellColor";
  if (sortBy === "fontColor") key.sortOn = "fontColor";
  if (color) key.color = color;
  if (customList) key.customList = customList;
  return key;
}



function importedSortConditionColor(node, sortBy, styleDxfs = []) {
  const dxfId = Number.parseInt(node?.getAttribute("dxfId") || "", 10);
  if (Number.isInteger(dxfId)) {
    const dxf = styleDxfs[dxfId] || null;
    const dxfColor = sortBy === "fontColor" ? dxf?.fontColor : dxf?.fillColor;
    const normalizedDxfColor = normalizeSortColor(dxfColor);
    if (normalizedDxfColor) return normalizedDxfColor;
  }
  return normalizeSortColor(excelSortConditionColorToCss(node?.getAttribute("color") || ""));
}



function excelSortConditionColorToCss(value) {
  const text = String(value || "").trim().replace(/^#/, "");
  if (/^[0-9a-f]{8}$/i.test(text)) return `#${text.slice(2)}`;
  if (/^[0-9a-f]{6}$/i.test(text)) return `#${text}`;
  return String(value || "");
}



function importedSortConditionCustomList(value) {
  const items = String(value || "")
    .split(",")
    .map((item) => normalizeCustomSortText(item))
    .filter(Boolean);
  if (!items.length) return "";
  const list = SORT_CUSTOM_LISTS.find((candidate) => {
    const expected = candidate.values.map((item) => normalizeCustomSortText(item));
    return items.length <= expected.length && items.every((item, index) => item === expected[index]);
  });
  return list?.id || "";
}



function parseAutoFilterColumnCriterion(columnNode, styleDxfs = []) {
  const colorFilter = firstXmlChild(columnNode, "colorFilter");
  if (colorFilter) {
    const dxfId = Number.parseInt(colorFilter.getAttribute("dxfId") || "", 10);
    const target = xmlBooleanAttribute(colorFilter, "cellColor", true) ? "fill" : "font";
    const dxf = Number.isInteger(dxfId) ? styleDxfs[dxfId] || null : null;
    const color = normalizeSortColor(target === "font" ? dxf?.fontColor : dxf?.fillColor);
    if (color) {
      return {
        kind: "color",
        target,
        color,
      };
    }
  }
  const filters = firstXmlChild(columnNode, "filters");
  if (filters) {
    const values = xmlChildren(filters, "filter").map((node) => node.getAttribute("val") ?? "").filter((value) => value !== "");
    const dateGroups = xmlChildren(filters, "dateGroupItem").map(parseAutoFilterDateGroupCriterion).filter(Boolean);
    const includeBlank = xmlBooleanAttribute(filters, "blank");
    if (!values.length && !includeBlank && dateGroups.length === 1) return dateGroups[0];
    if (dateGroups.length) {
      return {
        kind: "dateGroups",
        values,
        includeBlank,
        dateGroups: dateGroups.map((criterion) => criterion.dateGroup).filter(Boolean),
      };
    }
    if (includeBlank) values.push("");
    if (values.length) return values;
  }
  const customFilters = firstXmlChild(columnNode, "customFilters");
  if (customFilters) {
    const filtersList = xmlChildren(customFilters, "customFilter");
    const first = filtersList[0];
    if (first) {
      const second = filtersList[1];
      const firstCriterion = parseExcelCustomFilterXml(first);
      const secondCriterion = second ? parseExcelCustomFilterXml(second) : null;
      return {
        kind: "custom",
        operator1: firstCriterion.operator,
        value1: firstCriterion.value,
        logic: xmlBooleanAttribute(customFilters, "and") ? "and" : "or",
        operator2: secondCriterion ? secondCriterion.operator : "equal",
        value2: secondCriterion ? secondCriterion.value : "",
      };
    }
  }
  const top10 = firstXmlChild(columnNode, "top10");
  if (top10) {
    const count = Number(top10.getAttribute("val"));
    return {
      kind: "custom",
      top10Source: {
        direction: xmlBooleanAttribute(top10, "top", true) ? "top" : "bottom",
        mode: xmlBooleanAttribute(top10, "percent") ? "percent" : "items",
        count: Number.isFinite(count) ? count : 10,
      },
    };
  }
  const dynamicFilter = firstXmlChild(columnNode, "dynamicFilter");
  if (dynamicFilter) return parseDynamicAutoFilterCriterion(dynamicFilter.getAttribute("type") || "", dynamicFilter);
  return null;
}



function parseAutoFilterDateGroupCriterion(node) {
  const year = Number.parseInt(node?.getAttribute("year") || "", 10);
  if (!Number.isInteger(year)) return null;
  const monthAttr = node.getAttribute("month");
  const dayAttr = node.getAttribute("day");
  const hourAttr = node.getAttribute("hour");
  const minuteAttr = node.getAttribute("minute");
  const secondAttr = node.getAttribute("second");
  const month = monthAttr == null ? null : Number.parseInt(monthAttr, 10);
  const day = dayAttr == null ? null : Number.parseInt(dayAttr, 10);
  const hour = hourAttr == null ? null : Number.parseInt(hourAttr, 10);
  const minute = minuteAttr == null ? null : Number.parseInt(minuteAttr, 10);
  const second = secondAttr == null ? null : Number.parseInt(secondAttr, 10);
  const grouping = node.getAttribute("dateTimeGrouping") || "year";
  let start = null;
  let end = null;
  if (grouping === "year") {
    start = autoFilterLocalDate(year, 0, 1);
    end = autoFilterLocalDate(year, 11, 31);
  } else if (grouping === "month" && month >= 1 && month <= 12) {
    start = autoFilterLocalDate(year, month - 1, 1);
    end = autoFilterLocalDate(year, month, 0);
  } else if (["day", "hour", "minute", "second"].includes(grouping) && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
    start = autoFilterLocalDate(year, month - 1, day);
    end = autoFilterLocalDate(year, month - 1, day);
  }
  if (!start || !end) return null;
  if (["hour", "minute", "second"].includes(grouping) && !(hour >= 0 && hour <= 23)) return null;
  if (["minute", "second"].includes(grouping) && !(minute >= 0 && minute <= 59)) return null;
  if (grouping === "second" && !(second >= 0 && second <= 59)) return null;
  return {
    kind: "custom",
    operator1: "greaterThanOrEqual",
    value1: autoFilterDateForSerial(start),
    logic: "and",
    operator2: "lessThanOrEqual",
    value2: autoFilterDateForSerial(end),
    dateGroup: {
      year,
      month: Number.isInteger(month) ? month : null,
      day: Number.isInteger(day) ? day : null,
      hour: Number.isInteger(hour) ? hour : null,
      minute: Number.isInteger(minute) ? minute : null,
      second: Number.isInteger(second) ? second : null,
      grouping,
    },
  };
}



function normalizeAutoFilterXmlOperator(operator) {
  const value = String(operator || "").trim();
  return {
    greaterThan: "greaterThan",
    greaterThanOrEqual: "greaterThanOrEqual",
    lessThan: "lessThan",
    lessThanOrEqual: "lessThanOrEqual",
    notEqual: "notEqual",
    equal: "equal",
  }[value] || "equal";
}



function parseExcelCustomFilterXml(node) {
  const operator = normalizeAutoFilterXmlOperator(node?.getAttribute("operator") || "equal");
  const value = node?.getAttribute("val") ?? "";
  if (operator === "equal" && value === "") return { operator: "blank", value: "" };
  if (operator === "notEqual" && value === "") return { operator: "nonBlank", value: "" };
  return decodeExcelCustomFilterWildcard(operator, value);
}



function decodeExcelCustomFilterWildcard(operator, value) {
  if (operator !== "equal" && operator !== "notEqual") return { operator, value };
  const pattern = String(value ?? "");
  const tokens = excelWildcardPatternTokens(pattern);
  if (!tokens.some((token) => token.wildcard)) return { operator, value: pattern };
  const hasQuestion = tokens.some((token) => token.wildcard === "?");
  if (hasQuestion) return { operator, value: pattern };
  const wildcardIndexes = tokens
    .map((token, index) => (token.wildcard === "*" ? index : -1))
    .filter((index) => index >= 0);
  const firstIsWildcard = wildcardIndexes.includes(0);
  const lastIsWildcard = wildcardIndexes.includes(tokens.length - 1);
  const middleWildcard = wildcardIndexes.some((index) => index !== 0 && index !== tokens.length - 1);
  const literal = tokens.filter((token) => !token.wildcard).map((token) => token.text).join("");
  if (!literal || middleWildcard) return { operator, value: pattern };
  if (firstIsWildcard && lastIsWildcard && wildcardIndexes.length === 2) {
    return { operator: operator === "notEqual" ? "notContains" : "contains", value: literal };
  }
  if (!firstIsWildcard && lastIsWildcard && wildcardIndexes.length === 1) {
    return { operator: operator === "notEqual" ? "notBeginsWith" : "beginsWith", value: literal };
  }
  if (firstIsWildcard && !lastIsWildcard && wildcardIndexes.length === 1) {
    return { operator: operator === "notEqual" ? "notEndsWith" : "endsWith", value: literal };
  }
  return { operator, value: pattern };
}



function excelWildcardPatternTokens(pattern) {
  const chars = Array.from(String(pattern ?? ""));
  const tokens = [];
  for (let index = 0; index < chars.length; index += 1) {
    const char = chars[index];
    if (char === "~" && index + 1 < chars.length && ["*", "?", "~"].includes(chars[index + 1])) {
      tokens.push({ text: chars[index + 1] });
      index += 1;
      continue;
    }
    if (char === "*" || char === "?") {
      tokens.push({ wildcard: char });
      continue;
    }
    tokens.push({ text: char });
  }
  return tokens;
}



function parseDynamicAutoFilterCriterion(type, node = null) {
  const value = String(type || "");
  if (value === "aboveAverage" || value === "belowAverage") {
    return { kind: "custom", dynamicSource: { type: value } };
  }
  if (value === "aboveStdDev" || value === "belowStdDev") {
    const count = autoFilterStdDevCount(node?.getAttribute?.("val"));
    return { kind: "custom", dynamicSource: { type: value, val: count } };
  }
  if (/^(?:M(?:[1-9]|1[0-2])|Q[1-4])$/.test(value)) {
    return {
      kind: "custom",
      datePeriod: value,
      dynamicSource: { type: value },
    };
  }
  const datePreset = {
    today: "dateToday",
    yesterday: "dateYesterday",
    tomorrow: "dateTomorrow",
    thisWeek: "dateThisWeek",
    lastWeek: "dateLastWeek",
    nextWeek: "dateNextWeek",
    thisMonth: "dateThisMonth",
    lastMonth: "dateLastMonth",
    nextMonth: "dateNextMonth",
    thisQuarter: "dateThisQuarter",
    lastQuarter: "dateLastQuarter",
    nextQuarter: "dateNextQuarter",
    thisYear: "dateThisYear",
    lastYear: "dateLastYear",
    nextYear: "dateNextYear",
    yearToDate: "dateYearToDate",
  }[value];
  if (!datePreset) return null;
  const dateRange = autoFilterDateRangeForPreset(datePreset);
  if (!dateRange) return null;
  return {
    kind: "custom",
    operator1: "greaterThanOrEqual",
    value1: autoFilterDateForSerial(dateRange.start),
    logic: "and",
    operator2: "lessThanOrEqual",
    value2: autoFilterDateForSerial(dateRange.end),
    datePreset,
  };
}



function parseWorkbookStyleDxfsXml(stylesXml, themeColors = DEFAULT_THEME_COLORS) {
  if (!stylesXml) return [];
  try {
    const doc = parseXmlDocument(stylesXml);
    const dxfs = firstXmlChild(doc.documentElement, "dxfs");
    return xmlChildren(dxfs, "dxf").map((dxf) => ({
      fillColor: parseDxfFillColor(dxf, themeColors),
      fontColor: parseDxfFontColor(dxf, themeColors),
    }));
  } catch (error) {
    console.warn("Failed to parse workbook differential styles.", error);
    return [];
  }
}



function parseDxfFillColor(dxfNode, themeColors = DEFAULT_THEME_COLORS) {
  const fill = firstXmlChild(dxfNode, "fill");
  const patternFill = firstXmlChild(fill, "patternFill") || fill;
  const colorNode = firstXmlChild(patternFill, "fgColor") || firstXmlChild(patternFill, "bgColor");
  return excelStyleColorNodeToCss(colorNode, themeColors);
}



function parseDxfFontColor(dxfNode, themeColors = DEFAULT_THEME_COLORS) {
  const font = firstXmlChild(dxfNode, "font");
  return excelStyleColorNodeToCss(firstXmlChild(font, "color"), themeColors);
}



function excelStyleColorNodeToCss(colorNode, themeColors = DEFAULT_THEME_COLORS) {
  if (!colorNode) return "";
  const rgb = colorNode.getAttribute("rgb");
  if (rgb) {
    const clean = String(rgb).replace(/^#/, "");
    const value = clean.length === 8 ? clean.slice(2) : clean;
    return normalizeColorValue(value);
  }
  const indexed = Number.parseInt(colorNode.getAttribute("indexed") || "", 10);
  if (Number.isInteger(indexed)) return normalizeColorValue(INDEXED_COLORS[indexed]);
  const theme = Number.parseInt(colorNode.getAttribute("theme") || "", 10);
  if (Number.isInteger(theme)) {
    const base = themeColors[theme] || DEFAULT_THEME_COLORS[theme] || "";
    if (!base) return "";
    const tint = Number(colorNode.getAttribute("tint") || 0);
    return normalizeColorValue(applyTint(`#${base}`, Number.isFinite(tint) ? tint : 0));
  }
  if (xmlBooleanAttribute(colorNode, "auto")) return "#000000";
  return "";
}



function xmlBooleanAttribute(node, name, defaultValue = false) {
  const value = node?.getAttribute(name);
  if (value == null || value === "") return Boolean(defaultValue);
  return value === "1" || value.toLowerCase() === "true";
}



async function extractWorkbookPhonetics(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const [sharedStringsXml, workbookXml, workbookRelsXml] = await Promise.all([
      zip.file("xl/sharedStrings.xml")?.async("text"),
      zip.file("xl/workbook.xml")?.async("text"),
      zip.file("xl/_rels/workbook.xml.rels")?.async("text"),
    ]);
    if (!sharedStringsXml || !workbookXml || !workbookRelsXml) return {};
    const sharedStrings = parseSharedStringPhonetics(sharedStringsXml);
    if (!sharedStrings.some((item) => item.phonetics.length)) return {};
    const sheetPaths = workbookSheetPaths(workbookXml, workbookRelsXml);
    const entries = await Promise.all(
      sheetPaths.map(async ({ name, path }) => ({
        name,
        cells: await parseWorksheetPhonetics(zip, path, sharedStrings),
      })),
    );
    return Object.fromEntries(entries.filter((entry) => Object.keys(entry.cells).length).map((entry) => [entry.name, entry.cells]));
  } catch (error) {
    console.warn("Failed to parse workbook phonetics.", error);
    return {};
  }
}



async function extractWorkbookDrawingShapes(buffer, themeColors = DEFAULT_THEME_COLORS) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const [workbookXml, workbookRelsXml] = await Promise.all([
      zip.file("xl/workbook.xml")?.async("text"),
      zip.file("xl/_rels/workbook.xml.rels")?.async("text"),
    ]);
    if (!workbookXml || !workbookRelsXml) return {};
    const sheetPaths = workbookSheetPaths(workbookXml, workbookRelsXml);
    const entries = await Promise.all(
      sheetPaths.map(async ({ name, path }) => ({
        name,
        shapes: await parseWorksheetDrawingShapes(zip, path, themeColors),
      })),
    );
    return Object.fromEntries(entries.filter((entry) => entry.shapes.length).map((entry) => [entry.name, entry.shapes]));
  } catch (error) {
    console.warn("Failed to parse workbook drawing shapes.", error);
    return {};
  }
}



async function parseWorksheetDrawingShapes(zip, worksheetPath, themeColors = DEFAULT_THEME_COLORS) {
  const [worksheetXml, relsXml] = await Promise.all([
    zip.file(worksheetPath)?.async("text"),
    zip.file(worksheetRelsPath(worksheetPath))?.async("text"),
  ]);
  if (!worksheetXml || !relsXml) return [];
  const worksheetDoc = parseXmlDocument(worksheetXml);
  const rels = worksheetRelationships(relsXml, worksheetPath);
  const drawingPaths = xmlElements(worksheetDoc.documentElement, "drawing")
    .map((drawing) => drawing.getAttribute("r:id") || drawing.getAttribute("id") || drawing.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"))
    .map((relId) => rels.get(relId))
    .filter(Boolean);

  const groups = await Promise.all(
    drawingPaths.map(async (drawingPath) => {
      try {
        return await parseDrawingXmlShapes(zip, drawingPath, themeColors);
      } catch (error) {
        console.warn("Skipped an unsupported worksheet drawing.", drawingPath, error);
        return [];
      }
    }),
  );
  return groups.flat();
}



function worksheetRelsPath(worksheetPath) {
  const parts = String(worksheetPath || "").split("/");
  const fileName = parts.pop();
  return `${parts.join("/")}/_rels/${fileName}.rels`.replace(/^\/+/, "");
}



function worksheetRelationships(relsXml, basePath) {
  const relsDoc = parseXmlDocument(relsXml);
  return new Map(
    xmlChildren(relsDoc.documentElement, "Relationship").map((rel) => [
      rel.getAttribute("Id"),
      normalizePackageRelationshipPath(basePath, rel.getAttribute("Target") || ""),
    ]),
  );
}



async function parseDrawingXmlShapes(zip, drawingPath, themeColors = DEFAULT_THEME_COLORS) {
  const [xml, resources] = await Promise.all([
    zip.file(drawingPath)?.async("text"),
    drawingRelationshipResources(zip, drawingPath, themeColors),
  ]);
  if (!xml) return [];
  const doc = parseXmlDocument(xml);
  const anchors = xmlChildren(doc.documentElement).filter((node) => ["oneCellAnchor", "twoCellAnchor", "absoluteAnchor"].includes(xmlLocalName(node)));
  let index = 0;
  const shapes = [];
  anchors.forEach((anchor) => {
    const parsedShapes = safeDrawingShapesFromAnchor(anchor, index, themeColors, resources, drawingPath);
    parsedShapes.forEach((shape) => {
      shapes.push(shape);
      index += 1;
    });
  });
  return resolveDrawingShapeConnections(shapes);
}



function safeDrawingShapesFromAnchor(anchor, index, themeColors, resources, drawingPath = "") {
  try {
    return drawingShapesFromAnchor(anchor, index, themeColors, resources);
  } catch (error) {
    console.warn("Skipped an unsupported DrawingML object.", { drawingPath, index, error });
    return [];
  }
}



function resolveDrawingShapeConnections(shapes) {
  const idMap = new Map(
    shapes
      .filter((shape) => shape._drawingObjectId)
      .map((shape) => [String(shape._drawingObjectId), shape.id]),
  );
  shapes.forEach((shape) => {
    if (shape._drawingConnections && shape.shape && isLineShapeType(shape.shape.type)) {
      const connections = {};
      ["start", "end"].forEach((side) => {
        const raw = shape._drawingConnections?.[side];
        const imageId = raw?.drawingId ? idMap.get(String(raw.drawingId)) : "";
        if (imageId && raw.site) connections[side] = { imageId, site: raw.site };
      });
      const normalized = normalizeLineShapeConnections(connections);
      if (normalized) shape.shape.connections = normalized;
    }
    delete shape._drawingObjectId;
    delete shape._drawingConnections;
  });
  return shapes;
}



async function drawingRelationshipResources(zip, drawingPath, themeColors = DEFAULT_THEME_COLORS) {
  const relsXml = await zip.file(worksheetRelsPath(drawingPath))?.async("text");
  const resources = { media: new Map(), charts: new Map(), targets: new Map() };
  if (!relsXml) return resources;
  const relsDoc = parseXmlDocument(relsXml);
  await Promise.all(xmlChildren(relsDoc.documentElement, "Relationship").map(async (rel) => {
    const id = rel.getAttribute("Id");
    const type = String(rel.getAttribute("Type") || "");
    const target = rel.getAttribute("Target") || "";
    const targetPath = normalizePackageRelationshipPath(drawingPath, target);
    if (!id || !targetPath) return;
    resources.targets.set(id, { type, path: targetPath });
    if (rel.getAttribute("TargetMode") === "External") return;
    const file = zip.file(targetPath);
    if (!file) return;
    try {
      if (type.includes("/image")) {
        const extension = drawingMediaExtension(targetPath);
        const mimeType = imageMimeType(extension);
        const bytes = await file.async("uint8array");
        resources.media.set(id, {
          name: drawingMediaName(targetPath),
          extension,
          mimeType,
          src: `data:${mimeType};base64,${bytesToBase64(bytes)}`,
        });
      } else if (type.includes("/chart")) {
        const xml = await file.async("text");
        const chart = drawingChartFromXml(xml, themeColors);
        if (chart) resources.charts.set(id, { path: targetPath, chart });
      }
    } catch (error) {
      console.warn("Skipped an unsupported DrawingML relationship.", { drawingPath, targetPath, type, error });
    }
  }));
  return resources;
}



async function drawingRelationshipMedia(zip, drawingPath) {
  return (await drawingRelationshipResources(zip, drawingPath)).media;
}



function drawingMediaName(targetPath) {
  const fileName = String(targetPath || "").split("/").pop() || "image";
  return fileName.replace(/\.[^.]+$/, "") || "image";
}



function drawingMediaExtension(targetPath) {
  const match = String(targetPath || "").match(/\.([^.\/]+)$/);
  return String(match?.[1] || "png").toLowerCase();
}



function drawingShapesFromAnchor(anchorNode, startIndex, themeColors = DEFAULT_THEME_COLORS, resources = {}) {
  const media = resources instanceof Map ? resources : resources.media || new Map();
  const directShapeNode = xmlChildren(anchorNode).find((node) => ["sp", "cxnSp", "pic", "graphicFrame"].includes(xmlLocalName(node)));
  if (directShapeNode) {
    const localName = xmlLocalName(directShapeNode);
    const geometry = drawingShapeGeometryFromAnchor(anchorNode, directShapeNode);
    const shape = localName === "pic"
      ? drawingPictureToModelWithGeometry(directShapeNode, startIndex, geometry, media, {}, themeColors)
      : localName === "graphicFrame"
        ? drawingGraphicFrameToModelWithGeometry(directShapeNode, startIndex, geometry, resources)
        : drawingShapeToModelWithGeometry(directShapeNode, startIndex, geometry, themeColors);
    return shape ? [shape] : [];
  }

  const groupNode = xmlChildren(anchorNode).find((node) => xmlLocalName(node) === "grpSp");
  if (!groupNode) return [];
  return drawingGroupShapesToModels(anchorNode, groupNode, startIndex, themeColors, resources);
}



function drawingShapeGeometryFromAnchor(anchorNode, shapeNode = null) {
  const isAbsolute = xmlLocalName(anchorNode) === "absoluteAnchor";
  const from = isAbsolute ? anchorPointFromAbsoluteDrawingNode(firstXmlChild(anchorNode, "pos")) : anchorPointFromDrawingNode(firstXmlChild(anchorNode, "from"));
  if (!from) return null;
  const extNode = firstXmlChild(anchorNode, "ext") || firstXmlChild(firstXmlChild(firstXmlChild(shapeNode, "spPr"), "xfrm"), "ext");
  const to = anchorPointFromDrawingNode(firstXmlChild(anchorNode, "to"));
  const ext = extNode
    ? {
        width: Math.max(1, emuToPixels(xmlNumber(extNode, "cx"))),
        height: Math.max(1, emuToPixels(xmlNumber(extNode, "cy"))),
      }
    : null;
  return {
    editAs: isAbsolute ? "absolute" : xmlLocalName(anchorNode) === "twoCellAnchor" ? "twoCell" : "oneCell",
    tl: from,
    br: ext ? null : to,
    ext,
  };
}



function drawingGroupShapesToModels(anchorNode, groupNode, startIndex, themeColors = DEFAULT_THEME_COLORS, resources = {}) {
  const media = resources instanceof Map ? resources : resources.media || new Map();
  const from = xmlLocalName(anchorNode) === "absoluteAnchor"
    ? anchorPointFromAbsoluteDrawingNode(firstXmlChild(anchorNode, "pos"))
    : anchorPointFromDrawingNode(firstXmlChild(anchorNode, "from"));
  const xfrm = firstXmlChild(firstXmlChild(groupNode, "grpSpPr"), "xfrm");
  const extNode = firstXmlChild(anchorNode, "ext") || firstXmlChild(xfrm, "ext") || firstXmlChild(xfrm, "chExt");
  if (!from || !extNode) return [];

  const anchorCx = Math.max(1, xmlNumber(extNode, "cx"));
  const anchorCy = Math.max(1, xmlNumber(extNode, "cy"));
  const chOff = firstXmlChild(xfrm, "chOff") || firstXmlChild(xfrm, "off");
  const chExt = firstXmlChild(xfrm, "chExt") || firstXmlChild(xfrm, "ext");
  const childOrigin = {
    x: xmlNumber(chOff, "x"),
    y: xmlNumber(chOff, "y"),
  };
  const childSize = {
    cx: Math.max(1, xmlNumber(chExt, "cx") || anchorCx),
    cy: Math.max(1, xmlNumber(chExt, "cy") || anchorCy),
  };
  const scale = {
    x: anchorCx / childSize.cx,
    y: anchorCy / childSize.cy,
  };
  const groupTransform = {
    childSize,
    rotation: normalizeRotationDegrees(xmlNumber(xfrm, "rot") / 60000),
    flipH: xfrm?.getAttribute("flipH") === "1" || xfrm?.getAttribute("flipH") === "true",
    flipV: xfrm?.getAttribute("flipV") === "1" || xfrm?.getAttribute("flipV") === "true",
  };
  const shapeNodes = drawingShapeNodesInGroup(groupNode);
  const groupName = drawingShapeName(groupNode, startIndex);
  const groupId = `shape-group-${startIndex + 1}-${stableObjectId(groupName) || "group"}`;
  return shapeNodes
    .map((shapeNode, offset) => {
      const geometry = drawingShapeGeometryFromGroupChild(from, shapeNode, childOrigin, scale, groupTransform);
      if (!geometry) return null;
      const options = { groupId, groupName, groupTransform };
      const localName = xmlLocalName(shapeNode);
      try {
        return localName === "pic"
          ? drawingPictureToModelWithGeometry(shapeNode, startIndex + offset, geometry, media, options, themeColors)
          : localName === "graphicFrame"
            ? drawingGraphicFrameToModelWithGeometry(shapeNode, startIndex + offset, geometry, resources, options)
          : drawingShapeToModelWithGeometry(shapeNode, startIndex + offset, geometry, themeColors, options);
      } catch (error) {
        console.warn("Skipped an unsupported grouped DrawingML object.", { groupName, index: startIndex + offset, error });
        return null;
      }
    })
    .filter(Boolean);
}



function drawingShapeNodesInGroup(groupNode) {
  const result = [];
  xmlChildren(groupNode).forEach((child) => {
    const localName = xmlLocalName(child);
    if (["sp", "cxnSp", "pic", "graphicFrame"].includes(localName)) {
      result.push(child);
    } else if (localName === "grpSp") {
      result.push(...drawingShapeNodesInGroup(child));
    }
  });
  return result;
}



function drawingShapeGeometryFromGroupChild(anchorFrom, shapeNode, childOrigin, scale, groupTransform = null) {
  const xfrm = firstXmlChild(firstXmlChild(shapeNode, "spPr"), "xfrm") || firstXmlChild(shapeNode, "xfrm");
  const off = firstXmlChild(xfrm, "off");
  const ext = firstXmlChild(xfrm, "ext");
  if (!off || !ext) return null;
  const localLeft = xmlNumber(off, "x") - childOrigin.x;
  const localTop = xmlNumber(off, "y") - childOrigin.y;
  const localWidth = Math.max(1, xmlNumber(ext, "cx"));
  const localHeight = Math.max(1, xmlNumber(ext, "cy"));
  const transformed = transformGroupedShapeBox(localLeft, localTop, localWidth, localHeight, groupTransform);
  const leftEmu = Math.round(anchorFrom.nativeColOff + transformed.left * scale.x);
  const topEmu = Math.round(anchorFrom.nativeRowOff + transformed.top * scale.y);
  const widthEmu = Math.max(1, Math.round(localWidth * scale.x));
  const heightEmu = Math.max(1, Math.round(localHeight * scale.y));
  return {
    editAs: "oneCell",
    tl: {
      col: anchorFrom.col,
      row: anchorFrom.row,
      nativeCol: anchorFrom.nativeCol,
      nativeRow: anchorFrom.nativeRow,
      nativeColOff: leftEmu,
      nativeRowOff: topEmu,
    },
    br: null,
    ext: {
      width: Math.max(1, emuToPixels(widthEmu)),
      height: Math.max(1, emuToPixels(heightEmu)),
    },
  };
}



function transformGroupedShapeBox(left, top, width, height, groupTransform = null) {
  const size = groupTransform?.childSize;
  if (!size || (!groupTransform.rotation && !groupTransform.flipH && !groupTransform.flipV)) {
    return { left, top, width, height };
  }
  const groupCenter = { x: size.cx / 2, y: size.cy / 2 };
  let center = { x: left + width / 2, y: top + height / 2 };
  if (groupTransform.flipH) center.x = size.cx - center.x;
  if (groupTransform.flipV) center.y = size.cy - center.y;
  if (groupTransform.rotation) center = rotatePointAround(center, groupCenter, groupTransform.rotation);
  return {
    left: center.x - width / 2,
    top: center.y - height / 2,
    width,
    height,
  };
}



function drawingShapeToModel(anchorNode, shapeNode, index, themeColors = DEFAULT_THEME_COLORS) {
  const geometry = drawingShapeGeometryFromAnchor(anchorNode, shapeNode);
  return drawingShapeToModelWithGeometry(shapeNode, index, geometry, themeColors);
}



function drawingShapeToModelWithGeometry(shapeNode, index, geometry, themeColors = DEFAULT_THEME_COLORS, options = {}) {
  const preset = drawingShapePreset(shapeNode);
  const type = drawingShapeType(shapeNode, preset);
  if (!geometry?.tl) return null;
  const name = drawingShapeName(shapeNode, index);
  const style = drawingLineStyleWithBakedFlip(
    drawingShapeStyleWithGroup(drawingShapeStyle(shapeNode, type, themeColors), options.groupTransform),
    type,
  );
  const text = drawingShapeText(shapeNode);
  const textRuns = drawingShapeTextRuns(shapeNode, themeColors);
  const textStyle = drawingShapeTextStyle(shapeNode, themeColors);
  const descriptorStyle = drawingShapeDescriptorStyle(type, geometry, { ...style, ...textStyle, text, ...(textRuns.length ? { textRuns } : {}) });
  const descriptor = shapeDescriptorForType(type, { ...descriptorStyle, label: name });
  return {
    id: `shape-import-${index + 1}-${stableObjectId(name)}`,
    name,
    src: descriptor.src,
    mimeType: descriptor.mimeType,
    extension: descriptor.extension,
    editAs: geometry.editAs,
    tl: descriptorStyle.tl || geometry.tl,
    br: geometry.br,
    ext: descriptorStyle.ext || geometry.ext,
    hyperlink: "",
    _drawingObjectId: drawingShapeDrawingId(shapeNode),
    _drawingConnections: drawingShapeConnectorConnections(shapeNode),
    ...(options.groupId ? { groupId: options.groupId, groupName: options.groupName || "" } : {}),
    shape: { ...descriptor.shape, sourcePreset: preset || "" },
  };
}



function drawingShapeDescriptorStyle(type, geometry, style = {}) {
  const width = Math.max(1, Math.round(Number(geometry.ext?.width) || 1));
  const height = Math.max(1, Math.round(Number(geometry.ext?.height) || 1));
  if (!isLineShapeType(type)) return { ...style, width, height };
  const points = normalizeShapeLinePoints(style.points);
  const endpoints = {
    start: { x: points.x1 * width, y: points.y1 * height },
    end: { x: points.x2 * width, y: points.y2 * height },
  };
  const result = lineBoxAndPointsFromEndpoints(endpoints.start, endpoints.end, { shape: { type, ...style, points } });
  const tl = drawingAnchorWithPixelOffset(geometry.tl, result.box.left, result.box.top);
  return {
    ...style,
    width: result.box.width,
    height: result.box.height,
    points: normalizeShapeLinePoints(result.points),
    tl,
    ext: {
      width: Math.max(1, Math.round(result.box.width)),
      height: Math.max(1, Math.round(result.box.height)),
    },
  };
}



function drawingAnchorWithPixelOffset(anchor, dx = 0, dy = 0) {
  return {
    ...anchor,
    nativeColOff: Number(anchor?.nativeColOff || 0) + Math.round(dx * EMU_PER_PIXEL),
    nativeRowOff: Number(anchor?.nativeRowOff || 0) + Math.round(dy * EMU_PER_PIXEL),
  };
}



function drawingGraphicFrameToModelWithGeometry(graphicFrameNode, index, geometry, resources = {}, options = {}) {
  if (!geometry?.tl) return null;
  const name = drawingShapeName(graphicFrameNode, index);
  const kind = drawingGraphicFrameKind(graphicFrameNode);
  const style = drawingShapeStyleWithGroup(drawingGraphicFrameStyle(graphicFrameNode), options.groupTransform);
  const chartResource = drawingGraphicFrameChartResource(graphicFrameNode, resources);
  if (chartResource?.chart?.series?.length) {
    const width = Math.max(1, Math.round(Number(geometry.ext?.width) || 520));
    const height = Math.max(1, Math.round(Number(geometry.ext?.height) || 320));
    const chart = {
      ...chartResource.chart,
      type: chartResource.chart.type || "column",
      title: chartResource.chart.title || name,
      sourcePath: chartResource.path || "",
    };
    const descriptor = shapeDescriptorForType("rect", {
      ...style,
      label: name,
      width,
      height,
      fillColor: "none",
      outlineColor: "none",
      outlineWidth: 0,
      text: "",
    });
    return {
      id: `shape-import-${index + 1}-${stableObjectId(name)}`,
      name,
      src: svgDataUrl(chartSvg(chart, chart.type, { width, height, title: chart.title })),
      mimeType: "image/svg+xml",
      extension: "svg",
      editAs: geometry.editAs,
      tl: geometry.tl,
      br: geometry.br,
      ext: geometry.ext,
      hyperlink: "",
      chart,
      _drawingObjectId: drawingShapeDrawingId(graphicFrameNode),
      ...(options.groupId ? { groupId: options.groupId, groupName: options.groupName || "" } : {}),
      shape: { ...descriptor.shape, sourcePreset: "graphicFrame", graphicFrameKind: kind },
    };
  }
  const descriptor = shapeDescriptorForType("rect", {
    ...style,
    label: name,
    width: geometry.ext?.width,
    height: geometry.ext?.height,
    fillColor: "#f8fbff",
    outlineColor: "#64748b",
    outlineWidth: 0.75,
    lineDash: "dash",
    text: kind,
    textColor: "#334155",
    textFontSize: 11,
    textAlign: "center",
    textVAlign: "middle",
  });
  return {
    id: `shape-import-${index + 1}-${stableObjectId(name)}`,
    name,
    src: descriptor.src,
    mimeType: descriptor.mimeType,
    extension: descriptor.extension,
    editAs: geometry.editAs,
    tl: geometry.tl,
    br: geometry.br,
    ext: geometry.ext,
    hyperlink: "",
    _drawingObjectId: drawingShapeDrawingId(graphicFrameNode),
    ...(options.groupId ? { groupId: options.groupId, groupName: options.groupName || "" } : {}),
    shape: { ...descriptor.shape, sourcePreset: "graphicFrame", graphicFrameKind: kind },
  };
}



function drawingGraphicFrameChartResource(graphicFrameNode, resources = {}) {
  const relId = drawingGraphicFrameRelationshipId(graphicFrameNode);
  return relId ? resources.charts?.get(relId) || null : null;
}



function drawingGraphicFrameRelationshipId(graphicFrameNode) {
  const chartNode = xmlElements(graphicFrameNode, "chart")[0];
  return (
    chartNode?.getAttribute("r:id") ||
    chartNode?.getAttribute("id") ||
    chartNode?.getAttributeNS?.("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") ||
    ""
  );
}



function drawingGraphicFrameKind(graphicFrameNode) {
  const uri = String(xmlElements(graphicFrameNode, "graphicData")[0]?.getAttribute("uri") || "");
  if (uri.includes("/chart")) return "グラフ";
  if (uri.includes("/diagram")) return "SmartArt";
  if (uri.includes("/lockedCanvas")) return "グラフィック";
  if (uri.includes("/table")) return "表";
  return "オブジェクト";
}



function drawingGraphicFrameStyle(graphicFrameNode) {
  const xfrm = firstXmlChild(graphicFrameNode, "xfrm");
  return {
    rotation: normalizeRotationDegrees(xmlNumber(xfrm, "rot") / 60000),
    flipH: xfrm?.getAttribute("flipH") === "1" || xfrm?.getAttribute("flipH") === "true",
    flipV: xfrm?.getAttribute("flipV") === "1" || xfrm?.getAttribute("flipV") === "true",
  };
}



function drawingChartFromXml(xml, themeColors = DEFAULT_THEME_COLORS) {
  if (!xml) return null;
  try {
    const doc = parseXmlDocument(xml);
    const chartRoot = xmlElements(doc.documentElement, "chart")[0] || doc.documentElement;
    const chartNode = drawingChartPlotNode(chartRoot);
    if (!chartNode) return null;
    const type = drawingChartType(chartNode);
    const rawSeries = xmlChildren(chartNode, "ser");
    const series = rawSeries
      .map((seriesNode, index) => drawingChartSeries(seriesNode, index, type, themeColors))
      .filter((item) => item && item.values.some((value) => Number.isFinite(value)));
    if (!series.length) return null;
    const labels = drawingChartLabels(rawSeries[0], series[0], type);
    const labelsRef = drawingChartLabelsReference(rawSeries[0], type);
    const titleRef = drawingChartTitleReference(chartRoot);
    return {
      type,
      title: drawingChartTitle(chartRoot),
      ...(titleRef ? { titleRef } : {}),
      titleStyle: drawingChartTitleStyle(chartRoot, themeColors),
      labels,
      ...(labelsRef ? { labelsRef } : {}),
      series: type === "pie" && series.length > 1 ? [series[0]] : series,
      grouping: drawingChartGrouping(chartNode),
      ...drawingChartAreaStyles(doc.documentElement, chartRoot, themeColors),
      axes: drawingChartAxes(chartRoot, themeColors),
      dataLabels: drawingChartDataLabels(chartNode, themeColors),
      legend: drawingChartLegend(chartRoot),
    };
  } catch (error) {
    console.warn("Failed to parse DrawingML chart", error);
    return null;
  }
}



function drawingChartPlotNode(chartRoot) {
  const candidates = ["barChart", "lineChart", "pieChart", "scatterChart", "areaChart"];
  return candidates.map((name) => xmlElements(chartRoot, name)[0]).find(Boolean) || null;
}



function drawingChartType(chartNode) {
  if (xmlLocalName(chartNode) === "barChart") {
    const direction = firstXmlChild(chartNode, "barDir")?.getAttribute("val") || "";
    return direction === "bar" ? "bar" : "column";
  }
  switch (xmlLocalName(chartNode)) {
    case "lineChart":
    case "areaChart":
      return "line";
    case "pieChart":
      return "pie";
    case "scatterChart":
      return "scatter";
    default:
      return "column";
  }
}



function drawingChartGrouping(chartNode) {
  const value = firstXmlChild(chartNode, "grouping")?.getAttribute("val") || "";
  if (value === "stacked" || value === "percentStacked") return value;
  return "clustered";
}



function drawingChartTitle(chartRoot) {
  const titleNode = xmlElements(chartRoot, "title")[0];
  return drawingChartText(titleNode) || "";
}



function drawingChartTitleReference(chartRoot) {
  return drawingChartReference(xmlElements(chartRoot, "title")[0]);
}



function drawingChartTitleStyle(chartRoot, themeColors = DEFAULT_THEME_COLORS) {
  const titleNode = xmlElements(chartRoot, "title")[0];
  return drawingChartTextStyle(titleNode, themeColors);
}



function drawingChartAreaStyles(chartSpaceRoot, chartRoot, themeColors = DEFAULT_THEME_COLORS) {
  const plotArea = firstXmlChild(chartRoot, "plotArea");
  const chartArea = drawingChartShapeStyle(firstXmlChild(chartSpaceRoot, "spPr"), themeColors, "area");
  const plot = drawingChartShapeStyle(firstXmlChild(plotArea, "spPr"), themeColors, "plot");
  return {
    ...(Object.keys(chartArea).length ? { chartArea } : {}),
    ...(Object.keys(plot).length ? { plotArea: plot } : {}),
  };
}



function drawingChartShapeStyle(spPr, themeColors = DEFAULT_THEME_COLORS) {
  if (!spPr) return {};
  const fill = firstXmlChild(spPr, "noFill")
    ? "none"
    : drawingColorToCss(firstXmlChild(spPr, "solidFill"), themeColors) ||
      drawingColorToCss(firstXmlChild(firstXmlChild(spPr, "gradFill"), "gs"), themeColors);
  const line = firstXmlChild(spPr, "ln");
  const outlineColor = line
    ? drawingShapeOutlineColor(line, null, themeColors)
    : "";
  const outlineWidth = line
    ? normalizeShapeOutlineWidth(emuToPoints(xmlNumber(line, "w")))
    : null;
  const lineDash = line ? drawingShapeLineDash(line) : "";
  return {
    ...(fill ? { fillColor: fill } : {}),
    ...(outlineColor ? { outlineColor } : {}),
    ...(outlineWidth ? { outlineWidth } : {}),
    ...(lineDash && lineDash !== "solid" ? { lineDash } : {}),
  };
}



function drawingChartLegend(chartRoot) {
  const legend = firstXmlChild(chartRoot, "legend");
  if (!legend) return { visible: false };
  const deleted = firstXmlChild(legend, "delete")?.getAttribute("val");
  if (deleted === "1" || deleted === "true") return { visible: false };
  const position = firstXmlChild(legend, "legendPos")?.getAttribute("val") || "";
  const map = {
    r: "right",
    l: "left",
    t: "top",
    b: "bottom",
    tr: "right",
  };
  return {
    visible: true,
    position: map[position] || "bottom",
  };
}



function drawingChartAxes(chartRoot, themeColors = DEFAULT_THEME_COLORS) {
  const plotArea = firstXmlChild(chartRoot, "plotArea");
  const categoryAxis = drawingChartAxisStyle(firstXmlChild(plotArea, "catAx"), themeColors);
  const valueAxis = drawingChartAxisStyle(firstXmlChild(plotArea, "valAx"), themeColors);
  const axes = {};
  if (Object.keys(categoryAxis).length) axes.category = categoryAxis;
  if (Object.keys(valueAxis).length) axes.value = valueAxis;
  return axes;
}



function drawingChartAxisStyle(axisNode, themeColors = DEFAULT_THEME_COLORS) {
  if (!axisNode) return {};
  const line = drawingChartLineStyle(firstXmlChild(firstXmlChild(axisNode, "spPr"), "ln"), themeColors);
  const majorGridlines = firstXmlChild(axisNode, "majorGridlines");
  const majorGridline = majorGridlines
    ? drawingChartLineStyle(firstXmlChild(firstXmlChild(majorGridlines, "spPr"), "ln"), themeColors)
    : {};
  const titleNode = firstXmlChild(axisNode, "title");
  const title = drawingChartText(titleNode);
  const titleStyle = drawingChartTextStyle(titleNode, themeColors);
  return {
    ...(title ? { title } : {}),
    ...(Object.keys(titleStyle).length ? { titleStyle } : {}),
    ...(Object.keys(line).length ? { line } : {}),
    ...(majorGridlines ? { majorGridline } : {}),
  };
}



function drawingChartLineStyle(line, themeColors = DEFAULT_THEME_COLORS) {
  if (!line) return {};
  const color = drawingShapeOutlineColor(line, null, themeColors);
  const width = normalizeShapeOutlineWidth(emuToPoints(xmlNumber(line, "w")));
  const dash = drawingShapeLineDash(line);
  return {
    ...(color ? { color } : {}),
    ...(width ? { width } : {}),
    ...(dash && dash !== "solid" ? { dash } : {}),
  };
}



function drawingChartDataLabels(chartNode, themeColors = DEFAULT_THEME_COLORS) {
  const labelsNode = firstXmlChild(chartNode, "dLbls");
  if (!labelsNode) return null;
  const enabled = (name) => {
    const node = firstXmlChild(labelsNode, name);
    const value = node?.getAttribute("val");
    return value === "1" || value === "true";
  };
  const model = {
    showValue: enabled("showVal"),
    showCategoryName: enabled("showCatName"),
    showSeriesName: enabled("showSerName"),
    showPercent: enabled("showPercent"),
    separator: firstXmlChild(labelsNode, "separator")?.textContent || "",
    position: firstXmlChild(labelsNode, "dLblPos")?.getAttribute("val") || "",
    textStyle: drawingChartTextStyle(labelsNode, themeColors),
  };
  return model.showValue || model.showCategoryName || model.showSeriesName || model.showPercent ? model : null;
}



function drawingChartSeries(seriesNode, index, type, themeColors = DEFAULT_THEME_COLORS) {
  if (!seriesNode) return null;
  const nameRef = drawingChartSeriesNameReference(seriesNode);
  if (type === "scatter") {
    const xNode = firstXmlChild(seriesNode, "xVal");
    const yNode = firstXmlChild(seriesNode, "yVal");
    const values = drawingChartCacheValues(xNode, true);
    const yValues = drawingChartCacheValues(yNode, true);
    return {
      name: drawingChartSeriesName(seriesNode, index),
      ...(nameRef ? { nameRef } : {}),
      values,
      ...(drawingChartReference(xNode) ? { valuesRef: drawingChartReference(xNode) } : {}),
      yValues: yValues.length ? yValues : values,
      ...(drawingChartReference(yNode) ? { yValuesRef: drawingChartReference(yNode) } : {}),
      ...drawingChartSeriesStyle(seriesNode, themeColors),
      dataLabels: drawingChartDataLabels(seriesNode, themeColors),
    };
  }
  const valueNode = firstXmlChild(seriesNode, "val");
  const valuesRef = drawingChartReference(valueNode);
  return {
    name: drawingChartSeriesName(seriesNode, index),
    ...(nameRef ? { nameRef } : {}),
    values: drawingChartCacheValues(valueNode, true),
    ...(valuesRef ? { valuesRef } : {}),
    ...drawingChartSeriesStyle(seriesNode, themeColors),
    dataLabels: drawingChartDataLabels(seriesNode, themeColors),
  };
}



function drawingChartSeriesStyle(seriesNode, themeColors = DEFAULT_THEME_COLORS) {
  const spPr = firstXmlChild(seriesNode, "spPr");
  const color =
    drawingColorToCss(firstXmlChild(spPr, "solidFill"), themeColors) ||
    drawingColorToCss(firstXmlChild(firstXmlChild(spPr, "ln"), "solidFill"), themeColors);
  const marker = drawingChartMarker(firstXmlChild(seriesNode, "marker"), themeColors);
  const pointColors = {};
  xmlChildren(seriesNode, "dPt").forEach((pointNode) => {
    const rawIndex = Number(firstXmlChild(pointNode, "idx")?.getAttribute("val"));
    const pointSpPr = firstXmlChild(pointNode, "spPr");
    const pointColor =
      drawingColorToCss(firstXmlChild(pointSpPr, "solidFill"), themeColors) ||
      drawingColorToCss(firstXmlChild(firstXmlChild(pointSpPr, "ln"), "solidFill"), themeColors);
    if (Number.isInteger(rawIndex) && pointColor) pointColors[rawIndex] = pointColor;
  });
  return {
    ...(color ? { color } : {}),
    ...(marker ? { marker } : {}),
    ...(Object.keys(pointColors).length ? { pointColors } : {}),
  };
}



function drawingChartMarker(markerNode, themeColors = DEFAULT_THEME_COLORS) {
  if (!markerNode) return null;
  const symbol = firstXmlChild(markerNode, "symbol")?.getAttribute("val") || "auto";
  if (symbol === "none") return { symbol: "none" };
  const rawSize = Number(firstXmlChild(markerNode, "size")?.getAttribute("val"));
  const spPr = firstXmlChild(markerNode, "spPr");
  const fillColor = drawingColorToCss(firstXmlChild(spPr, "solidFill"), themeColors);
  const outlineColor = drawingColorToCss(firstXmlChild(firstXmlChild(spPr, "ln"), "solidFill"), themeColors);
  const outlineWidth = normalizeShapeOutlineWidth(emuToPoints(xmlNumber(firstXmlChild(spPr, "ln"), "w")));
  return {
    symbol,
    ...(Number.isFinite(rawSize) ? { size: Math.max(2, Math.min(18, rawSize)) } : {}),
    ...(fillColor ? { fillColor } : {}),
    ...(outlineColor ? { outlineColor } : {}),
    ...(outlineWidth ? { outlineWidth } : {}),
  };
}



function drawingChartSeriesName(seriesNode, index) {
  return drawingChartText(firstXmlChild(seriesNode, "tx")) || `系列 ${index + 1}`;
}



function drawingChartSeriesNameReference(seriesNode) {
  return drawingChartReference(firstXmlChild(seriesNode, "tx"));
}



function drawingChartLabels(seriesNode, series, type) {
  if (!seriesNode) return series.values.map((_value, index) => String(index + 1));
  if (type === "scatter") {
    const xLabels = drawingChartCacheValues(firstXmlChild(seriesNode, "xVal"), false);
    return xLabels.length ? xLabels.map(String) : series.values.map((value) => String(value ?? ""));
  }
  const labels = drawingChartCacheValues(firstXmlChild(seriesNode, "cat"), false);
  return labels.length ? labels.map(String) : series.values.map((_value, index) => String(index + 1));
}



function drawingChartLabelsReference(seriesNode, type) {
  if (!seriesNode) return "";
  return drawingChartReference(type === "scatter" ? firstXmlChild(seriesNode, "xVal") : firstXmlChild(seriesNode, "cat"));
}



function drawingChartReference(node) {
  if (!node) return "";
  return (xmlElements(node, "f")[0]?.textContent || "").trim();
}



function drawingChartText(node) {
  const texts = drawingChartCacheValues(node, false);
  if (texts.length) return texts.join("");
  const direct = firstXmlChildText(node, "v");
  if (direct) return direct;
  return xmlElements(node, "t")
    .map((textNode) => textNode.textContent || "")
    .join("");
}



function drawingChartTextStyle(node, themeColors = DEFAULT_THEME_COLORS) {
  if (!node) return {};
  const paragraph =
    xmlElements(node, "p").find((item) => firstXmlChild(item, "r") || firstXmlChild(firstXmlChild(item, "pPr"), "defRPr")) ||
    firstXmlChild(node, "p");
  const paragraphProperties = firstXmlChild(paragraph, "pPr");
  const defaultRunProperties = firstXmlChild(paragraphProperties, "defRPr") || firstXmlChild(paragraph, "endParaRPr") || null;
  const runProperties = firstXmlChild(xmlElements(paragraph, "r")[0], "rPr") || defaultRunProperties || xmlElements(node, "defRPr")[0] || null;
  return drawingShapeTextRunStyle(runProperties, defaultRunProperties, themeColors);
}



function drawingChartCacheValues(node, numeric = false) {
  if (!node) return [];
  const points = xmlElements(node, "pt")
    .map((point) => ({
      index: Number(point.getAttribute("idx")),
      value: firstXmlChildText(point, "v"),
    }))
    .sort((left, right) => (Number.isFinite(left.index) ? left.index : 0) - (Number.isFinite(right.index) ? right.index : 0));
  const values = points.length ? points.map((point) => point.value) : xmlChildren(node, "v").map((valueNode) => valueNode.textContent || "");
  if (!numeric) return values.filter((value) => value !== "");
  return values.map((value) => {
    const number = Number(String(value ?? "").replace(/[,¥$€£%]/g, ""));
    return Number.isFinite(number) ? number : null;
  });
}



function drawingPictureToModelWithGeometry(pictureNode, index, geometry, media = new Map(), options = {}, themeColors = DEFAULT_THEME_COLORS) {
  if (!geometry?.tl) return null;
  const relId = drawingPictureRelationshipId(pictureNode);
  const source = relId ? media.get(relId) : null;
  if (!source?.src) return null;
  const name = drawingShapeName(pictureNode, index);
  const pictureStyle = drawingPictureStyleWithGroup(pictureNode, options.groupTransform, themeColors);
  return {
    id: `image-import-${index + 1}-${stableObjectId(name)}`,
    name,
    src: source.src,
    mimeType: source.mimeType,
    extension: source.extension,
    editAs: geometry.editAs,
    tl: geometry.tl,
    br: geometry.br,
    ext: geometry.ext,
    hyperlink: "",
    ...pictureStyle,
    _drawingObjectId: drawingShapeDrawingId(pictureNode),
    ...(options.groupId ? { groupId: options.groupId, groupName: options.groupName || "" } : {}),
  };
}



function drawingPictureRelationshipId(pictureNode) {
  const blip = xmlElements(pictureNode, "blip")[0];
  return (
    blip?.getAttribute("r:embed") ||
    blip?.getAttribute("embed") ||
    blip?.getAttributeNS?.("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed") ||
    ""
  );
}



function drawingPictureStyleWithGroup(pictureNode, groupTransform = null, themeColors = DEFAULT_THEME_COLORS) {
  const xfrm = firstXmlChild(firstXmlChild(pictureNode, "spPr"), "xfrm");
  const rotation = normalizeRotationDegrees((xmlNumber(xfrm, "rot") || 0) / 60000 + (groupTransform?.rotation || 0));
  const flipH = (xfrm?.getAttribute("flipH") === "1" || xfrm?.getAttribute("flipH") === "true") !== Boolean(groupTransform?.flipH);
  const flipV = (xfrm?.getAttribute("flipV") === "1" || xfrm?.getAttribute("flipV") === "true") !== Boolean(groupTransform?.flipV);
  const crop = drawingPictureCrop(pictureNode);
  return {
    ...(rotation ? { rotation } : {}),
    ...(flipH ? { flipH } : {}),
    ...(flipV ? { flipV } : {}),
    ...(crop ? { crop } : {}),
    ...drawingPictureVisualStyle(pictureNode, themeColors),
  };
}



function drawingPictureVisualStyle(pictureNode, themeColors = DEFAULT_THEME_COLORS) {
  const spPr = firstXmlChild(pictureNode, "spPr");
  if (!spPr) return {};
  const line = firstXmlChild(spPr, "ln");
  const outline = line
    ? {
        outlineColor: drawingShapeOutlineColor(line, null, themeColors),
        outlineWidth: normalizeShapeOutlineWidth(emuToPoints(xmlNumber(line, "w"))),
        lineDash: drawingShapeLineDash(line),
        lineCompound: normalizeShapeLineCompound(line.getAttribute("cmpd")),
      }
    : {};
  return {
    ...outline,
    ...drawingShapeEffects(spPr, themeColors),
  };
}



function drawingPictureCrop(pictureNode) {
  const srcRect = xmlElements(pictureNode, "srcRect")[0];
  if (!srcRect) return null;
  return normalizePictureCrop({
    left: drawingCropPercent(srcRect, "l"),
    top: drawingCropPercent(srcRect, "t"),
    right: drawingCropPercent(srcRect, "r"),
    bottom: drawingCropPercent(srcRect, "b"),
  });
}



function drawingCropPercent(node, attr) {
  const raw = Number(node?.getAttribute(attr));
  return Number.isFinite(raw) ? raw / 100000 : 0;
}



function normalizePictureCrop(crop = null) {
  if (!crop || typeof crop !== "object") return null;
  let left = clampNumber(crop.left, 0, 0.98, 0);
  let top = clampNumber(crop.top, 0, 0.98, 0);
  let right = clampNumber(crop.right, 0, 0.98, 0);
  let bottom = clampNumber(crop.bottom, 0, 0.98, 0);
  const horizontal = left + right;
  if (horizontal >= 0.98) {
    const scale = 0.98 / horizontal;
    left *= scale;
    right *= scale;
  }
  const vertical = top + bottom;
  if (vertical >= 0.98) {
    const scale = 0.98 / vertical;
    top *= scale;
    bottom *= scale;
  }
  return left || top || right || bottom ? { left, top, right, bottom } : null;
}



function drawingShapeDrawingId(shapeNode) {
  return xmlElements(shapeNode, "cNvPr")[0]?.getAttribute("id") || "";
}



function drawingShapeConnectorConnections(shapeNode) {
  if (xmlLocalName(shapeNode) !== "cxnSp") return null;
  const connectionProps = xmlElements(shapeNode, "cNvCxnSpPr")[0];
  if (!connectionProps) return null;
  const start = drawingShapeConnectionEndpoint(firstXmlChild(connectionProps, "stCxn"));
  const end = drawingShapeConnectionEndpoint(firstXmlChild(connectionProps, "endCxn"));
  return start || end ? { ...(start ? { start } : {}), ...(end ? { end } : {}) } : null;
}



function drawingShapeConnectionEndpoint(node) {
  if (!node) return null;
  const drawingId = node.getAttribute("id") || "";
  const site = drawingConnectionSite(node.getAttribute("idx"));
  return drawingId && site ? { drawingId, site } : null;
}



function drawingConnectionSite(index) {
  const raw = Number(index);
  if (!Number.isFinite(raw)) return "connection-0";
  const normalized = Math.max(0, Math.trunc(raw));
  return `connection-${normalized}`;
}



function drawingShapeStyleWithGroup(style, groupTransform = null) {
  if (!groupTransform) return style;
  return {
    ...style,
    rotation: normalizeRotationDegrees((style.rotation || 0) + (groupTransform.rotation || 0)),
    flipH: Boolean(style.flipH) !== Boolean(groupTransform.flipH),
    flipV: Boolean(style.flipV) !== Boolean(groupTransform.flipV),
  };
}



function drawingLineStyleWithBakedFlip(style, type) {
  if (!isLineShapeType(type) || (!style.flipH && !style.flipV)) return style;
  const points = normalizeShapeLinePoints(style.points);
  const nextPoints = {
    x1: style.flipH ? 1 - points.x1 : points.x1,
    y1: style.flipV ? 1 - points.y1 : points.y1,
    x2: style.flipH ? 1 - points.x2 : points.x2,
    y2: style.flipV ? 1 - points.y2 : points.y2,
  };
  return {
    ...style,
    flipH: false,
    flipV: false,
    points: normalizeShapeLinePoints(nextPoints),
  };
}



function drawingShapeName(shapeNode, index) {
  const props = xmlElements(shapeNode, "cNvPr")[0];
  return props?.getAttribute("name") || `図形 ${index + 1}`;
}



function drawingShapeText(shapeNode) {
  const body = firstXmlChild(shapeNode, "txBody");
  if (!body) return "";
  return xmlChildren(body, "p")
    .map((paragraph) => xmlTextContent(paragraph))
    .filter((line, index, lines) => line || index < lines.length - 1)
    .join("\n")
    .trim();
}



function drawingShapeTextRuns(shapeNode, themeColors = DEFAULT_THEME_COLORS) {
  const body = firstXmlChild(shapeNode, "txBody");
  if (!body) return [];
  const paragraphs = xmlChildren(body, "p");
  const runs = [];
  paragraphs.forEach((paragraph, paragraphIndex) => {
    const paragraphProperties = firstXmlChild(paragraph, "pPr");
    const defaultRunProperties = firstXmlChild(paragraphProperties, "defRPr") || firstXmlChild(paragraph, "endParaRPr") || null;
    xmlChildren(paragraph).forEach((child) => {
      const name = xmlLocalName(child);
      if (name === "r" || name === "fld") {
        const text = firstXmlChild(child, "t")?.textContent || "";
        if (!text) return;
        runs.push({
          text,
          ...drawingShapeTextRunStyle(firstXmlChild(child, "rPr"), defaultRunProperties, themeColors),
        });
      } else if (name === "br") {
        runs.push({ text: "\n" });
      }
    });
    if (paragraphIndex < paragraphs.length - 1) runs.push({ text: "\n" });
  });
  return normalizeShapeTextRuns(runs);
}



function drawingShapeTextRunStyle(runProperties, defaultRunProperties = null, themeColors = DEFAULT_THEME_COLORS) {
  const attr = (name) => runProperties?.getAttribute(name) ?? defaultRunProperties?.getAttribute(name) ?? "";
  const color = drawingColorToCss(firstXmlChild(runProperties, "solidFill"), themeColors) ||
    drawingColorToCss(firstXmlChild(defaultRunProperties, "solidFill"), themeColors);
  const size = Number(attr("sz")) / 100;
  const fontFamily = drawingShapeTextFontFamily(runProperties) || drawingShapeTextFontFamily(defaultRunProperties);
  const underline = drawingRunUnderline(runProperties) || drawingRunUnderline(defaultRunProperties);
  const style = {};
  if (color) style.color = color;
  if (Number.isFinite(size) && size > 0) style.fontSize = Math.max(6, Math.min(72, size));
  if (fontFamily) style.fontFamily = fontFamily;
  if (attr("b") === "1") style.bold = true;
  if (attr("b") === "0") style.bold = false;
  if (attr("i") === "1") style.italic = true;
  if (attr("i") === "0") style.italic = false;
  if (underline) style.underline = underline;
  return style;
}



function drawingShapeTextStyle(shapeNode, themeColors = DEFAULT_THEME_COLORS) {
  const body = firstXmlChild(shapeNode, "txBody");
  if (!body) return {};
  const bodyPr = firstXmlChild(body, "bodyPr");
  const firstParagraph = firstXmlChild(body, "p");
  const paragraphProperties = firstXmlChild(firstParagraph, "pPr");
  const runProperties = firstXmlChild(xmlElements(firstParagraph, "r")[0], "rPr") || firstXmlChild(paragraphProperties, "defRPr") || xmlElements(firstParagraph, "defRPr")[0] || firstXmlChild(paragraphProperties, "endParaRPr") || null;
  const textColor = drawingColorToCss(firstXmlChild(runProperties, "solidFill"), themeColors) || "";
  const textFontSize = xmlNumber(runProperties, "sz") / 100;
  const textFontFamily = drawingShapeTextFontFamily(runProperties);
  const textOutline = drawingShapeTextOutline(runProperties, themeColors);
  const paragraphAlign = paragraphProperties?.getAttribute("algn") || "";
  const alignMap = { l: "left", r: "right", ctr: "center", just: "center" };
  const verticalMap = { t: "top", ctr: "middle", b: "bottom", just: "middle", dist: "middle" };
  const vertical = verticalMap[bodyPr?.getAttribute("anchor")] || "";
  const underline = drawingRunUnderline(runProperties);
  const textMargin = drawingShapeTextMargin(bodyPr);
  return {
    ...(textColor ? { textColor } : {}),
    ...(Number.isFinite(textFontSize) && textFontSize > 0 ? { textFontSize: Math.max(6, Math.min(72, textFontSize)) } : {}),
    ...(textFontFamily ? { textFontFamily } : {}),
    ...(textOutline || {}),
    ...(runProperties?.getAttribute("b") === "1" ? { textBold: true } : {}),
    ...(runProperties?.getAttribute("i") === "1" ? { textItalic: true } : {}),
    ...(underline ? { textUnderline: underline } : {}),
    ...(alignMap[paragraphAlign] ? { textAlign: alignMap[paragraphAlign] } : {}),
    ...(vertical ? { textVAlign: vertical } : {}),
    ...(textMargin ? { textMargin } : {}),
  };
}



function drawingShapeTextOutline(runProperties, themeColors = DEFAULT_THEME_COLORS) {
  const line = firstXmlChild(runProperties, "ln");
  if (!line) return null;
  if (firstXmlChild(line, "noFill")) return { textOutlineColor: "none" };
  const color = drawingColorToCss(firstXmlChild(line, "solidFill"), themeColors);
  return color ? {
    textOutlineColor: color,
    textOutlineWidth: normalizeShapeTextOutlineWidth(emuToPoints(xmlNumber(line, "w"))),
  } : null;
}



function drawingShapeTextMargin(bodyPr) {
  if (!bodyPr) return null;
  const fields = [
    ["left", "lIns"],
    ["right", "rIns"],
    ["top", "tIns"],
    ["bottom", "bIns"],
  ];
  const margin = {};
  fields.forEach(([key, attr]) => {
    if (!bodyPr.hasAttribute(attr)) return;
    margin[key] = Math.max(0, Math.round(emuToPixels(xmlNumber(bodyPr, attr)) * 10) / 10);
  });
  return Object.keys(margin).length ? margin : null;
}



function drawingShapeTextFontFamily(runProperties) {
  const typeface =
    firstXmlChild(runProperties, "latin")?.getAttribute("typeface") ||
    firstXmlChild(runProperties, "ea")?.getAttribute("typeface") ||
    firstXmlChild(runProperties, "cs")?.getAttribute("typeface") ||
    "";
  const clean = String(typeface || "").trim();
  return clean && clean !== "+mn-lt" && clean !== "+mn-ea" && clean !== "+mj-lt" && clean !== "+mj-ea"
    ? fontFamilyForExcel(clean)
    : "";
}



function drawingRunUnderline(runProperties) {
  const value = String(runProperties?.getAttribute("u") || "").toLowerCase();
  if (!value || value === "none") return "";
  return value.includes("dbl") || value.includes("double") ? "underline double" : "underline";
}



function drawingShapePreset(shapeNode) {
  if (firstXmlChild(firstXmlChild(shapeNode, "spPr"), "custGeom")) return CUSTOM_SHAPE_TYPE;
  return xmlElements(shapeNode, "prstGeom")[0]?.getAttribute("prst") || "";
}



function drawingShapeType(shapeNode, preset) {
  const localName = xmlLocalName(shapeNode);
  if (localName === "cxnSp") {
    const line = firstXmlChild(firstXmlChild(shapeNode, "spPr"), "ln");
    const hasHead = drawingLineEndHasArrow(firstXmlChild(line, "headEnd"));
    const hasTail = drawingLineEndHasArrow(firstXmlChild(line, "tailEnd"));
    if (hasHead && hasTail) return "doubleArrow";
    if (hasHead || hasTail) return "arrow";
    if (/curved/i.test(preset)) return "curvedConnector";
    if (/bent|elbow/i.test(preset)) return "elbowConnector";
    return "line";
  }
  if (xmlElements(shapeNode, "cNvSpPr")[0]?.getAttribute("txBox") === "1") {
    return "textBox";
  }
  if (preset === CUSTOM_SHAPE_TYPE) return CUSTOM_SHAPE_TYPE;
  const map = {
    rect: "rect",
    roundRect: "roundRect",
    snip1Rect: "snipRect",
    snip2SameRect: "snipRect",
    snip2DiagRect: "snipRect",
    round1Rect: "roundSnipRect",
    round2SameRect: "roundSnipRect",
    round2DiagRect: "roundSnipRect",
    frame: "frame",
    ellipse: "oval",
    oval: "oval",
    triangle: "triangle",
    rtTriangle: "rightTriangle",
    diamond: "diamond",
    pentagon: "pentagon",
    hexagon: "hexagon",
    heptagon: "heptagon",
    octagon: "octagon",
    decagon: "decagon",
    dodecagon: "dodecagon",
    parallelogram: "parallelogram",
    trapezoid: "trapezoid",
    plus: "cross",
    can: "can",
    cube: "cube",
    bevel: "bevel",
    donut: "donut",
    noSmoking: "noSymbol",
    blockArc: "blockArc",
    foldedCorner: "foldedCorner",
    smileyFace: "smileyFace",
    heart: "heart",
    lightningBolt: "lightningBolt",
    sun: "sun",
    moon: "moon",
    cloud: "cloud",
    arc: "arc",
    pie: "pie",
    chord: "chord",
    teardrop: "teardrop",
    plaque: "plaque",
    corner: "corner",
    halfFrame: "halfFrame",
    diagStripe: "diagStripe",
    leftBracket: "leftBracket",
    rightBracket: "rightBracket",
    bracketPair: "bracketPair",
    leftBrace: "leftBrace",
    rightBrace: "rightBrace",
    bracePair: "bracePair",
    rightArrow: "rightBlockArrow",
    leftArrow: "leftBlockArrow",
    upArrow: "upBlockArrow",
    downArrow: "downBlockArrow",
    leftRightArrow: "leftRightBlockArrow",
    leftRightUpArrow: "leftRightUpArrow",
    quadArrow: "quadArrow",
    notchedRightArrow: "notchedRightArrow",
    bentUpArrow: "bentUpArrow",
    uturnArrow: "uturnArrow",
    curvedRightArrow: "curvedRightArrow",
    leftUpArrow: "leftUpArrow",
    stripedRightArrow: "stripedRightArrow",
    homePlate: "homePlate",
    chevron: "chevron",
    leftRightCircularArrow: "leftRightCircularArrow",
    mathPlus: "mathPlus",
    mathMinus: "mathMinus",
    mathMultiply: "mathMultiply",
    mathDivide: "mathDivide",
    mathEqual: "mathEqual",
    mathNotEqual: "mathNotEqual",
    flowChartProcess: "flowProcess",
    flowChartDecision: "flowDecision",
    flowChartInputOutput: "flowData",
    flowChartDocument: "flowDocument",
    flowChartMultidocument: "flowMultiDocument",
    flowChartTerminator: "flowTerminator",
    flowChartPredefinedProcess: "flowPredefinedProcess",
    flowChartStoredData: "flowStoredData",
    flowChartInternalStorage: "flowInternalStorage",
    flowChartManualInput: "flowManualInput",
    flowChartManualOperation: "flowManualOperation",
    flowChartPreparation: "flowPreparation",
    flowChartConnector: "flowConnector",
    flowChartOffpageConnector: "flowOffpageConnector",
    flowChartDelay: "flowDelay",
    flowChartDisplay: "flowDisplay",
    flowChartOr: "flowOr",
    flowChartSummingJunction: "flowSummingJunction",
    star4: "star4",
    star5: "star5",
    star6: "star6",
    star7: "star7",
    star8: "star8",
    star10: "star10",
    star12: "star12",
    star16: "star16",
    star24: "star24",
    star32: "star32",
    seal: "seal",
    irregularSeal1: "seal",
    irregularSeal2: "explosion",
    ribbon: "ribbon",
    wave: "wave",
    wedgeRectCallout: "calloutRect",
    wedgeRoundRectCallout: "calloutRoundRect",
    wedgeEllipseCallout: "calloutOval",
    cloudCallout: "calloutCloud",
  };
  return map[preset] || DEFAULT_SHAPE_TYPE;
}



function drawingShapeStyle(shapeNode, type, themeColors = DEFAULT_THEME_COLORS) {
  const spPr = firstXmlChild(shapeNode, "spPr");
  const style = firstXmlChild(shapeNode, "style");
  const xfrm = firstXmlChild(spPr, "xfrm");
  const fillColor = drawingShapeFillColor(spPr, style, type, themeColors);
  const fillGradient = drawingShapeFillGradient(spPr, themeColors);
  const fillPattern = drawingShapeFillPattern(spPr, style, themeColors);
  const line = firstXmlChild(spPr, "ln");
  const outlineColor = drawingShapeOutlineColor(line, style, themeColors);
  const outlineWidth = normalizeShapeOutlineWidth(emuToPoints(xmlNumber(line, "w")));
  const lineDash = drawingShapeLineDash(line);
  const lineCompound = normalizeShapeLineCompound(line?.getAttribute("cmpd"));
  const lineCap = normalizeShapeLineCap(line?.getAttribute("cap"));
  const lineJoin = drawingShapeLineJoin(line);
  const rotation = normalizeRotationDegrees(xmlNumber(xfrm, "rot") / 60000);
  const flipH = xfrm?.getAttribute("flipH") === "1" || xfrm?.getAttribute("flipH") === "true";
  const flipV = xfrm?.getAttribute("flipV") === "1" || xfrm?.getAttribute("flipV") === "true";
  const adjust = drawingShapeAdjustments(shapeNode, type);
  const points = drawingShapeLinePoints(shapeNode, type);
  const customPath = type === CUSTOM_SHAPE_TYPE ? drawingShapeCustomPath(shapeNode) : null;
  const lineStartArrowStyle = drawingLineEndStyle(firstXmlChild(line, "headEnd"));
  const lineEndArrowStyle = drawingLineEndStyle(firstXmlChild(line, "tailEnd"));
  const effects = drawingShapeEffects(spPr, themeColors);
  return {
    fillColor,
    ...(fillGradient ? { fillGradient } : {}),
    ...(fillPattern ? { fillPattern } : {}),
    outlineColor,
    outlineWidth,
    lineDash,
    lineCompound,
    lineCap,
    lineJoin,
    rotation,
    flipH,
    flipV,
    lineStartArrow: Boolean(lineStartArrowStyle),
    lineEndArrow: Boolean(lineEndArrowStyle),
    ...(lineStartArrowStyle ? { lineStartArrowType: lineStartArrowStyle.type } : {}),
    ...(lineEndArrowStyle ? { lineEndArrowType: lineEndArrowStyle.type } : {}),
    ...(lineStartArrowStyle ? { lineStartArrowSize: lineStartArrowStyle.size } : {}),
    ...(lineEndArrowStyle ? { lineEndArrowSize: lineEndArrowStyle.size } : {}),
    ...(points ? { points } : {}),
    ...(adjust ? { adjust } : {}),
    ...(customPath ? { customPath } : {}),
    ...effects,
  };
}



function drawingShapeLinePoints(shapeNode, type) {
  if (!isLineShapeType(type)) return null;
  return {
    x1: 0,
    y1: 0,
    x2: 1,
    y2: 1,
  };
}



function drawingShapeCustomPath(shapeNode) {
  const customGeometry = firstXmlChild(firstXmlChild(shapeNode, "spPr"), "custGeom");
  const pathNode = firstXmlChild(firstXmlChild(customGeometry, "pathLst"), "path");
  if (!pathNode) return null;
  const guideValues = drawingShapeGuideValues(customGeometry);
  const sourceWidth = Math.max(1, drawingShapePathCoordinate(pathNode.getAttribute("w"), 21600, guideValues, 21600));
  const sourceHeight = Math.max(1, drawingShapePathCoordinate(pathNode.getAttribute("h"), 21600, guideValues, 21600));
  const commands = [];
  xmlChildren(pathNode).forEach((node) => {
    const name = xmlLocalName(node);
    if (name === "close") {
      commands.push({ cmd: "Z", points: [] });
      return;
    }
    const points = xmlChildren(node, "pt")
      .map((point) => ({
        x: drawingShapePathCoordinate(point.getAttribute("x"), sourceWidth, guideValues, 0) / sourceWidth,
        y: drawingShapePathCoordinate(point.getAttribute("y"), sourceHeight, guideValues, 0) / sourceHeight,
      }))
      .filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
    if (name === "moveTo" && points[0]) {
      commands.push({ cmd: "M", points: [points[0]] });
    } else if (name === "lnTo" && points[0]) {
      commands.push({ cmd: "L", points: [points[0]] });
    } else if (name === "quadBezTo" && points.length >= 2) {
      commands.push({ cmd: "Q", points: points.slice(0, 2) });
    } else if (name === "cubicBezTo" && points.length >= 3) {
      commands.push({ cmd: "C", points: points.slice(0, 3) });
    }
  });
  return normalizeShapeCustomPath({ commands });
}



function drawingShapeGuideValues(customGeometry) {
  const values = new Map();
  xmlElements(firstXmlChild(customGeometry, "gdLst"), "gd").forEach((guide) => {
    const name = guide.getAttribute("name");
    const formula = String(guide.getAttribute("fmla") || "").trim();
    const match = formula.match(/^val\s+(-?\d+(?:\.\d+)?)$/i);
    if (name && match) values.set(name, Number(match[1]));
  });
  return values;
}



function drawingShapePathCoordinate(rawValue, axisSize, guideValues, fallback) {
  const raw = String(rawValue || "").trim();
  if (!raw) return fallback;
  const number = Number(raw);
  if (Number.isFinite(number)) return number;
  if (guideValues?.has(raw)) return Number(guideValues.get(raw));
  if (raw === "l" || raw === "t") return 0;
  if (raw === "r") return axisSize;
  if (raw === "b") return axisSize;
  if (raw === "hc" || raw === "vc") return axisSize / 2;
  return fallback;
}



function drawingLineEndHasArrow(node) {
  return Boolean(drawingLineEndStyle(node));
}



function drawingLineEndStyle(node) {
  if (!node) return null;
  const type = String(node.getAttribute("type") || "").toLowerCase();
  if (!type || type === "none") return null;
  return {
    type: normalizeShapeArrowType(type) || "triangle",
    size: normalizeShapeArrowSize(node.getAttribute("w") || node.getAttribute("len") || "medium"),
  };
}



function drawingShapeFillColor(spPr, style, type, themeColors = DEFAULT_THEME_COLORS) {
  if (firstXmlChild(spPr, "noFill")) return "none";
  const color =
    drawingColorToCss(firstXmlChild(spPr, "solidFill"), themeColors) ||
    drawingColorToCss(firstXmlChild(firstXmlChild(spPr, "gradFill"), "gs"), themeColors) ||
    drawingColorToCss(firstXmlChild(firstXmlChild(style, "fillRef"), "schemeClr"), themeColors);
  return color || (isLineShapeType(type) ? "none" : "#4472c4");
}



function drawingShapeFillGradient(spPr, themeColors = DEFAULT_THEME_COLORS) {
  const gradFill = firstXmlChild(spPr, "gradFill");
  if (!gradFill) return null;
  const stops = xmlElements(firstXmlChild(gradFill, "gsLst"), "gs")
    .map((stop) => {
      const color = drawingColorToCss(stop, themeColors);
      const rawPos = Number(stop.getAttribute("pos"));
      return color ? { pos: Number.isFinite(rawPos) ? Math.max(0, Math.min(1, rawPos / 100000)) : 0, color } : null;
    })
    .filter(Boolean)
    .sort((a, b) => a.pos - b.pos);
  if (stops.length < 2) return null;
  const angle = normalizeRotationDegrees(xmlNumber(firstXmlChild(gradFill, "lin"), "ang") / 60000);
  return normalizeShapeGradient({ type: "linear", angle, stops });
}



function drawingShapeFillPattern(spPr, style, themeColors = DEFAULT_THEME_COLORS) {
  const pattern = firstXmlChild(spPr, "pattFill");
  if (!pattern) return null;
  const preset = pattern.getAttribute("prst") || "pct20";
  const foreground = drawingColorToCss(firstXmlChild(pattern, "fgClr"), themeColors) ||
    drawingColorToCss(firstXmlChild(firstXmlChild(style, "fillRef"), "schemeClr"), themeColors) ||
    "#4472c4";
  const background = drawingColorToCss(firstXmlChild(pattern, "bgClr"), themeColors) || "rgba(255,255,255,0)";
  return normalizeShapeFillPattern({ preset, foreground, background });
}



function drawingShapeOutlineColor(line, style, themeColors = DEFAULT_THEME_COLORS) {
  if (firstXmlChild(line, "noFill")) return "none";
  return (
    drawingColorToCss(firstXmlChild(line, "solidFill"), themeColors) ||
    drawingColorToCss(firstXmlChild(firstXmlChild(style, "lnRef"), "schemeClr"), themeColors) ||
    "#2f5597"
  );
}



function drawingShapeLineDash(line) {
  const value = firstXmlChild(line, "prstDash")?.getAttribute("val") || "";
  const map = {
    solid: "solid",
    sysDash: "shortDash",
    dash: "dash",
    dashDot: "dashDot",
    lgDash: "longDash",
    lgDashDot: "longDashDot",
    lgDashDotDot: "longDashDotDot",
    sysDashDot: "dashDot",
    sysDashDotDot: "longDashDotDot",
    dot: "dot",
    sysDot: "dot",
    lgDashDotDot: "longDashDotDot",
    dashDotDot: "longDashDotDot",
    squareDot: "squareDot",
  };
  return map[value] || "solid";
}



function drawingShapeLineJoin(line) {
  if (!line) return "round";
  if (firstXmlChild(line, "bevel")) return "bevel";
  if (firstXmlChild(line, "miter")) return "miter";
  if (firstXmlChild(line, "round")) return "round";
  return "round";
}



function drawingShapeEffects(spPr, themeColors = DEFAULT_THEME_COLORS) {
  const effectList = firstXmlChild(spPr, "effectLst");
  if (!effectList) return {};
  const shadow = drawingShapeShadowEffect(effectList, themeColors);
  const glow = drawingShapeGlowEffect(effectList, themeColors);
  const softEdges = drawingShapeSoftEdgeEffect(effectList);
  const reflection = drawingShapeReflectionEffect(effectList);
  return {
    ...(shadow ? { shadow } : {}),
    ...(glow ? { glow } : {}),
    ...(softEdges ? { softEdges } : {}),
    ...(reflection ? { reflection } : {}),
  };
}



function drawingShapeShadowEffect(effectList, themeColors = DEFAULT_THEME_COLORS) {
  const shadow = firstXmlChild(effectList, "outerShdw");
  if (!shadow) return null;
  const distance = Math.max(0, emuToPixels(xmlNumber(shadow, "dist")));
  const blur = Math.max(0, emuToPixels(xmlNumber(shadow, "blurRad")));
  const degrees = normalizeRotationDegrees(xmlNumber(shadow, "dir") / 60000);
  const radians = (degrees * Math.PI) / 180;
  const dx = Math.round(Math.cos(radians) * distance * 10) / 10;
  const dy = Math.round(Math.sin(radians) * distance * 10) / 10;
  const colorNode = xmlChildren(shadow).find((child) => /Clr$/.test(xmlLocalName(child)));
  const color = drawingColorToCss(colorNode, themeColors) || "rgba(0,0,0,0.35)";
  return { type: "outer", dx, dy, blur, color };
}



function drawingShapeGlowEffect(effectList, themeColors = DEFAULT_THEME_COLORS) {
  const glow = firstXmlChild(effectList, "glow");
  if (!glow) return null;
  const radius = Math.max(1, Math.min(24, emuToPixels(xmlNumber(glow, "rad")) || 5));
  const colorNode = xmlChildren(glow).find((child) => /Clr$/.test(xmlLocalName(child)));
  const color = drawingColorToCss(colorNode, themeColors) || "#5b9bd5";
  return normalizeShapeGlow({ radius, color });
}



function drawingShapeSoftEdgeEffect(effectList) {
  const softEdge = firstXmlChild(effectList, "softEdge");
  if (!softEdge) return 0;
  return normalizeShapeSoftEdges(emuToPoints(xmlNumber(softEdge, "rad")));
}



function drawingShapeReflectionEffect(effectList) {
  const reflection = firstXmlChild(effectList, "reflection");
  if (!reflection) return "";
  const blur = Math.max(0, emuToPixels(xmlNumber(reflection, "blurRad")));
  const distance = Math.max(0, emuToPixels(xmlNumber(reflection, "dist")));
  const alpha = drawingAlphaValue(reflection);
  return normalizeShapeReflection({ preset: alpha > 0.72 ? "full" : alpha > 0.45 ? "half" : "tight", blur, distance, alpha });
}



function legacyDrawingShapeShadow(spPr, themeColors = DEFAULT_THEME_COLORS) {
  const shadow = firstXmlChild(firstXmlChild(spPr, "effectLst"), "outerShdw");
  if (!shadow) return null;
  const distance = Math.max(0, emuToPixels(xmlNumber(shadow, "dist")));
  const blur = Math.max(0, emuToPixels(xmlNumber(shadow, "blurRad")));
  const degrees = normalizeRotationDegrees(xmlNumber(shadow, "dir") / 60000);
  const radians = (degrees * Math.PI) / 180;
  const dx = Math.round(Math.cos(radians) * distance * 10) / 10;
  const dy = Math.round(Math.sin(radians) * distance * 10) / 10;
  const colorNode = xmlChildren(shadow).find((child) => /Clr$/.test(xmlLocalName(child)));
  const color = drawingColorToCss(colorNode, themeColors) || "rgba(0,0,0,0.35)";
  return { type: "outer", dx, dy, blur, color };
}



function drawingShapeShadow(spPr, themeColors = DEFAULT_THEME_COLORS) {
  return legacyDrawingShapeShadow(spPr, themeColors);
}



/* legacy shim retained for snapshots created before effect parsing was expanded */
function drawingShapeLineDashLegacy(line) {
  const value = firstXmlChild(line, "prstDash")?.getAttribute("val") || "";
  const map = {
    solid: "solid",
    dash: "dash",
    sysDash: "dash",
    dot: "dot",
    sysDot: "dot",
    dashDot: "dashDot",
    sysDashDot: "dashDot",
    lgDash: "longDash",
    lgDashDot: "longDash",
  };
  return map[value] || "solid";
}



function drawingShapeAdjustments(shapeNode, type) {
  const guides = Object.fromEntries(
    xmlElements(firstXmlChild(firstXmlChild(shapeNode, "spPr"), "prstGeom"), "gd")
      .map((node) => [node.getAttribute("name") || "", drawingGuideValue(node.getAttribute("fmla") || "")])
      .filter(([name, value]) => name && Number.isFinite(value)),
  );
  if (!Object.keys(guides).length) return null;
  const guide = (name, fallback = null) => (Number.isFinite(guides[name]) ? guides[name] : fallback);
  const pct = (value) => (Number.isFinite(value) ? value / 100000 : null);
  switch (type) {
    case "roundRect":
      return normalizeShapeAdjustments(type, { corner: pct(guide("adj")) });
    case "snipRect":
    case "roundSnipRect":
      return normalizeShapeAdjustments(type, { cut: pct(guide("adj", guide("adj1"))), corner: pct(guide("adj2")) });
    case "parallelogram":
    case "trapezoid":
    case "flowData":
    case "flowManualInput":
      return normalizeShapeAdjustments(type, { skew: pct(guide("adj", guide("adj1"))), inset: pct(guide("adj", guide("adj1"))) });
    case "cross":
    case "leftRightBlockArrow":
    case "quadArrow":
    case "curvedRightArrow":
      return normalizeShapeAdjustments(type, { thickness: pct(guide("adj", guide("adj1"))) });
    case "rightBlockArrow":
    case "leftBlockArrow":
    case "upBlockArrow":
    case "downBlockArrow":
      return normalizeShapeAdjustments(type, {
        thickness: pct(guide("adj1", guide("adj"))),
        head: pct(guide("adj2")),
      });
    case "can":
      return normalizeShapeAdjustments(type, { depth: pct(guide("adj")) });
    case "star4":
    case "star5":
    case "star6":
    case "star7":
    case "star8":
    case "star10":
    case "star12":
    case "star16":
    case "star24":
    case "star32":
      return normalizeShapeAdjustments(type, { inner: drawingStarInnerFromGuide(guide("adj")) });
    case "calloutRect":
    case "calloutRoundRect":
    case "calloutOval":
    case "calloutCloud":
      return normalizeShapeAdjustments(type, {
        tailX: drawingCalloutPositionFromGuide(guide("adj1")),
        tailY: drawingCalloutPositionFromGuide(guide("adj2")),
      });
    default:
      return null;
  }
}



function drawingStarInnerFromGuide(value) {
  if (!Number.isFinite(value)) return null;
  return value / 50000;
}



function drawingCalloutPositionFromGuide(value) {
  if (!Number.isFinite(value)) return null;
  return 0.5 + value / 100000;
}



function drawingGuideValue(formula) {
  const match = String(formula || "").match(/(?:^|\s)val\s+(-?\d+(?:\.\d+)?)/i);
  return match ? Number(match[1]) : Number.NaN;
}



function drawingColorToCss(node, themeColors = DEFAULT_THEME_COLORS) {
  if (!node) return "";
  const colorNode = ["solidFill", "gs", "outerShdw", "fgClr", "bgClr", "glow", "reflection"].includes(xmlLocalName(node))
    ? xmlChildren(node).find((child) => /Clr$/.test(xmlLocalName(child)))
    : node;
  if (!colorNode) return "";
  const localName = xmlLocalName(colorNode);
  if (localName === "srgbClr") {
    const value = colorNode.getAttribute("val");
    return /^[0-9a-f]{6}$/i.test(value || "") ? drawingColorTransforms(`#${value}`, colorNode) : "";
  }
  if (localName === "schemeClr") {
    const base = themeColorForScheme(colorNode.getAttribute("val"), themeColors);
    return base ? drawingColorTransforms(base, colorNode) : "";
  }
  if (localName === "prstClr") {
    const base = presetDrawingColor(colorNode.getAttribute("val"));
    return base ? drawingColorTransforms(base, colorNode) : "";
  }
  if (localName === "sysClr") {
    const lastColor = colorNode.getAttribute("lastClr");
    const base = /^[0-9a-f]{6}$/i.test(lastColor || "")
      ? `#${lastColor}`
      : systemDrawingColor(colorNode.getAttribute("val"));
    return base ? drawingColorTransforms(base, colorNode) : "";
  }
  return "";
}



function drawingColorTransforms(color, node) {
  let nextColor = color;
  const shade = Number(firstXmlChild(node, "shade")?.getAttribute("val"));
  if (Number.isFinite(shade) && shade > 0) {
    nextColor = applyTint(nextColor, -(1 - shade / 100000));
  }
  const tint = Number(firstXmlChild(node, "tint")?.getAttribute("val"));
  if (Number.isFinite(tint) && tint > 0) {
    nextColor = applyTint(nextColor, 1 - tint / 100000);
  }
  const lumMod = Number(firstXmlChild(node, "lumMod")?.getAttribute("val"));
  const lumOff = Number(firstXmlChild(node, "lumOff")?.getAttribute("val"));
  if (Number.isFinite(lumMod) || Number.isFinite(lumOff)) {
    const mod = Number.isFinite(lumMod) ? lumMod / 100000 : 1;
    const off = Number.isFinite(lumOff) ? lumOff / 100000 : 0;
    nextColor = adjustColorLuminosity(nextColor, mod, off);
  }
  return applyDrawingColorAlpha(nextColor, drawingAlphaValue(node));
}



function drawingAlphaValue(node) {
  const alpha = Number(firstXmlChild(node, "alpha")?.getAttribute("val"));
  const alphaMod = Number(firstXmlChild(node, "alphaMod")?.getAttribute("val"));
  const alphaOff = Number(firstXmlChild(node, "alphaOff")?.getAttribute("val"));
  let value = 1;
  if (Number.isFinite(alpha)) value = alpha / 100000;
  if (Number.isFinite(alphaMod)) value *= alphaMod / 100000;
  if (Number.isFinite(alphaOff)) value += alphaOff / 100000;
  return Math.max(0, Math.min(1, value));
}



function applyDrawingColorAlpha(color, alpha = 1) {
  if (!Number.isFinite(alpha) || alpha >= 0.999) return color;
  const rgb = hexToRgb(color);
  if (!rgb) return color;
  return `rgba(${rgb[0]},${rgb[1]},${rgb[2]},${Math.round(alpha * 1000) / 1000})`;
}



function themeColorForScheme(name, themeColors = DEFAULT_THEME_COLORS) {
  const index = {
    lt1: 0,
    dk1: 1,
    lt2: 2,
    dk2: 3,
    accent1: 4,
    accent2: 5,
    accent3: 6,
    accent4: 7,
    accent5: 8,
    accent6: 9,
    hlink: 10,
    folHlink: 11,
    tx1: 1,
    tx2: 3,
    bg1: 0,
    bg2: 2,
  }[name];
  const hex = Number.isInteger(index) ? themeColors[index] || DEFAULT_THEME_COLORS[index] : "";
  return hex ? `#${hex}` : "";
}



function presetDrawingColor(name) {
  const colors = {
    black: "#000000",
    white: "#ffffff",
    red: "#ff0000",
    green: "#008000",
    blue: "#0000ff",
    yellow: "#ffff00",
  };
  return colors[String(name || "").toLowerCase()] || "";
}



function systemDrawingColor(name) {
  const colors = {
    windowtext: "#000000",
    window: "#ffffff",
    text: "#000000",
    background: "#ffffff",
    btntext: "#000000",
    btnface: "#f0f0f0",
    highlight: "#3399ff",
    highlighttext: "#ffffff",
    graytext: "#808080",
  };
  return colors[String(name || "").toLowerCase()] || "";
}



function adjustColorLuminosity(color, mod = 1, off = 0) {
  const rgb = hexToRgb(color);
  if (!rgb) return color;
  const next = rgb.map((channel) => Math.max(0, Math.min(255, Math.round(channel * mod + 255 * off))));
  return `#${next.map((value) => value.toString(16).padStart(2, "0")).join("")}`;
}



function anchorPointFromDrawingNode(node) {
  if (!node) return null;
  const col = xmlNumber(firstXmlChild(node, "col"));
  const row = xmlNumber(firstXmlChild(node, "row"));
  if (!Number.isFinite(col) || !Number.isFinite(row)) return null;
  return {
    col,
    row,
    nativeCol: col,
    nativeRow: row,
    nativeColOff: xmlNumber(firstXmlChild(node, "colOff")) || 0,
    nativeRowOff: xmlNumber(firstXmlChild(node, "rowOff")) || 0,
  };
}



function anchorPointFromAbsoluteDrawingNode(node) {
  if (!node) return null;
  return {
    col: 0,
    row: 0,
    nativeCol: 0,
    nativeRow: 0,
    nativeColOff: xmlNumber(node, "x") || 0,
    nativeRowOff: xmlNumber(node, "y") || 0,
  };
}



function xmlNumber(node, attr = "") {
  const value = attr ? node?.getAttribute(attr) : node?.textContent;
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}



function emuToPixels(value) {
  return Math.round((Number(value) || 0) / 9525);
}



function emuToPoints(value) {
  return (Number(value) || 0) / 12700;
}



function stableObjectId(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9_-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 24);
}



function parseSharedStringPhonetics(xml) {
  const doc = parseXmlDocument(xml);
  return xmlChildren(doc.documentElement, "si").map((si) => {
    const text = sharedStringBaseText(si);
    const phonetics = xmlChildren(si, "rPh")
      .map((node) => ({
        start: Number.parseInt(node.getAttribute("sb") || "0", 10),
        end: Number.parseInt(node.getAttribute("eb") || "0", 10),
        text: xmlTextContent(node),
      }))
      .filter((run) => Number.isFinite(run.start) && Number.isFinite(run.end) && run.end > run.start && run.text.trim());
    return { text, phonetics: normalizePhoneticRuns(phonetics, text) };
  });
}



function sharedStringBaseText(si) {
  let text = "";
  xmlChildren(si).forEach((child) => {
    if (xmlLocalName(child) === "t") {
      text += child.textContent || "";
    } else if (xmlLocalName(child) === "r") {
      xmlChildren(child, "t").forEach((textNode) => {
        text += textNode.textContent || "";
      });
    }
  });
  return text;
}



function workbookSheetPaths(workbookXml, workbookRelsXml) {
  const workbookDoc = parseXmlDocument(workbookXml);
  const relsDoc = parseXmlDocument(workbookRelsXml);
  const rels = new Map(
    xmlChildren(relsDoc.documentElement, "Relationship").map((rel) => [rel.getAttribute("Id"), normalizeWorkbookRelationshipPath(rel.getAttribute("Target") || "")]),
  );
  return xmlElements(workbookDoc.documentElement, "sheet")
    .map((sheet) => {
      const relId = sheet.getAttribute("r:id") || sheet.getAttribute("id") || sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
      const path = rels.get(relId);
      return path ? { name: sheet.getAttribute("name") || "", path } : null;
    })
    .filter(Boolean);
}



async function extractWorkbookSheetViewDetails(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
    const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
    if (!workbookXml || !workbookRelsXml) return {};
    const details = {};
    for (const [sheetIndex, sheet] of workbookSheetPaths(workbookXml, workbookRelsXml).entries()) {
      const view = await extractWorksheetViewDetails(zip, sheet.path, sheetIndex);
      if (view) details[sheet.name] = view;
    }
    return details;
  } catch (error) {
    console.warn("Failed to extract worksheet view details.", error);
    return {};
  }
}



async function extractWorkbookOutlineDetails(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
    const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
    if (!workbookXml || !workbookRelsXml) return {};
    const details = {};
    for (const sheet of workbookSheetPaths(workbookXml, workbookRelsXml)) {
      const outline = await extractWorksheetOutlineDetails(zip, sheet.path);
      if (outline) details[sheet.name] = outline;
    }
    return details;
  } catch (error) {
    console.warn("Failed to extract worksheet outline details.", error);
    return {};
  }
}



async function extractWorksheetOutlineDetails(zip, path) {
  const xml = await zip.file(path)?.async("text");
  if (!xml) return null;
  const doc = parseXmlDocument(xml);
  const rows = {};
  const columns = {};
  const outlinePr = xmlElements(doc.documentElement, "outlinePr")[0] || null;
  const settings = outlinePr ? worksheetOutlineSettingsFromXml(outlinePr) : null;
  let rowExtent = 0;
  let columnExtent = 0;
  xmlElements(doc.documentElement, "row").forEach((node) => {
    const index = Math.max(1, Math.round(Number(node.getAttribute("r")) || 0));
    if (!index) return;
    const detail = worksheetOutlineNodeDetail(node);
    if (detail) {
      rows[index] = detail;
      rowExtent = Math.max(rowExtent, index);
    }
  });
  xmlElements(doc.documentElement, "col").forEach((node) => {
    const min = Math.max(1, Math.round(Number(node.getAttribute("min")) || 0));
    const max = Math.max(min, Math.round(Number(node.getAttribute("max")) || min));
    const detail = worksheetOutlineNodeDetail(node);
    if (!detail) return;
    columnExtent = Math.max(columnExtent, max);
    for (let index = min; index <= max; index += 1) {
      columns[index] = { ...(columns[index] || {}), ...detail };
    }
  });
  if (!Object.keys(rows).length && !Object.keys(columns).length && !settings) return null;
  return { rows, columns, rowExtent, columnExtent, ...(settings ? { settings } : {}) };
}



function worksheetOutlineSettingsFromXml(node) {
  return normalizeOutlineSettings({
    summaryBelow: xmlBoolAttribute(node, "summaryBelow", true),
    summaryRight: xmlBoolAttribute(node, "summaryRight", true),
    applyStyles: xmlBoolAttribute(node, "applyStyles", false),
  });
}



function worksheetOutlineNodeDetail(node) {
  if (!node) return null;
  const outlineLevel = Number(node.getAttribute("outlineLevel"));
  const hidden = xmlBoolAttribute(node, "hidden", false);
  const collapsed = xmlBoolAttribute(node, "collapsed", false);
  if (!outlineLevel && !hidden && !collapsed) return null;
  return {
    ...(Number.isFinite(outlineLevel) ? { outlineLevel: clampOutlineLevel(outlineLevel) } : {}),
    ...(hidden ? { hidden: true, outlineHidden: Number.isFinite(outlineLevel) && outlineLevel > 0 } : {}),
    ...(collapsed ? { outlineCollapsed: true } : {}),
  };
}



function worksheetOutlineDetailForIndex(details, axis, index) {
  const collection = details?.[axis];
  if (!collection) return null;
  const normalized = Math.max(1, Math.round(Number(index) || 0));
  return collection[normalized] || collection[String(normalized)] || null;
}



function worksheetOutlineDetailsDimensions(details) {
  if (!details) return { rows: 0, cols: 0 };
  const maxObjectIndex = (items) => Math.max(
    0,
    ...Object.keys(items || {})
      .map((key) => Number(key))
      .filter(Number.isFinite),
  );
  return {
    rows: Math.max(Number(details.rowExtent) || 0, maxObjectIndex(details.rows)),
    cols: Math.max(Number(details.columnExtent) || 0, maxObjectIndex(details.columns)),
  };
}



async function extractWorkbookWindowViews(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
    if (!workbookXml) return [];
    const doc = parseXmlDocument(workbookXml);
    return xmlElements(doc.documentElement, "workbookView")
      .map((node) => ({
        activeSheetIndex: Number(node.getAttribute("activeTab") ?? node.getAttribute("activeSheetId") ?? 0),
        firstSheet: Number(node.getAttribute("firstSheet") ?? 0),
        visibility: node.getAttribute("visibility") || "visible",
        x: Number(node.getAttribute("xWindow") ?? node.getAttribute("x") ?? 0),
        y: Number(node.getAttribute("yWindow") ?? node.getAttribute("y") ?? 0),
        width: Number(node.getAttribute("windowWidth") ?? node.getAttribute("width") ?? 0),
        height: Number(node.getAttribute("windowHeight") ?? node.getAttribute("height") ?? 0),
        maximized: xmlBoolAttribute(node, "maximized", false),
        minimized: xmlBoolAttribute(node, "minimized", false),
        showHorizontalScroll: xmlBoolAttribute(node, "showHorizontalScroll", true),
        showVerticalScroll: xmlBoolAttribute(node, "showVerticalScroll", true),
        showSheetTabs: xmlBoolAttribute(node, "showSheetTabs", true),
      }))
      .filter(Boolean);
  } catch (error) {
    console.warn("Failed to extract workbook window views.", error);
    return [];
  }
}



async function extractWorkbookMetadata(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
    if (!workbookXml) return {};
    const doc = parseXmlDocument(workbookXml);
    const workbookPr = firstXmlChild(doc.documentElement, "workbookPr");
    const calcPr = firstXmlChild(doc.documentElement, "calcPr");
    return {
      workbookProperties: normalizeWorkbookProperties({
        date1904: xmlBoolAttribute(workbookPr, "date1904", false),
        codeName: workbookPr?.getAttribute("codeName") || "",
        filterPrivacy: xmlBoolAttribute(workbookPr, "filterPrivacy", false),
        showObjects: workbookPr?.getAttribute("showObjects") || "all",
        defaultThemeVersion: workbookPr?.getAttribute("defaultThemeVersion") || "",
      }),
      calcProperties: normalizeWorkbookCalcProperties({
        calcId: calcPr?.getAttribute("calcId") || "",
        calcMode: calcPr?.getAttribute("calcMode") || "",
        refMode: calcPr?.getAttribute("refMode") || "",
        fullCalcOnLoad: xmlBoolAttribute(calcPr, "fullCalcOnLoad", true),
        forceFullCalc: xmlBoolAttribute(calcPr, "forceFullCalc", false),
        calcOnSave: xmlBoolAttribute(calcPr, "calcOnSave", true),
        concurrentCalc: xmlBoolAttribute(calcPr, "concurrentCalc", true),
        fullPrecision: xmlBoolAttribute(calcPr, "fullPrecision", true),
        iterate: xmlBoolAttribute(calcPr, "iterate", false),
        iterateCount: calcPr?.getAttribute("iterateCount") || "",
        iterateDelta: calcPr?.getAttribute("iterateDelta") || "",
      }),
    };
  } catch (error) {
    console.warn("Failed to extract workbook metadata.", error);
    return {};
  }
}



async function extractWorksheetViewDetails(zip, path, sheetIndex = 0) {
  const xml = await zip.file(path)?.async("text");
  if (!xml) return null;
  const doc = parseXmlDocument(xml);
  const sheetPr = firstXmlChild(doc.documentElement, "sheetPr");
  const pageSetupPr = sheetPr ? firstXmlChild(sheetPr, "pageSetUpPr") : null;
  const pageSetupNode = firstXmlChild(doc.documentElement, "pageSetup");
  const pageSetupDetails = {
    pageSetupFitToPage: xmlBoolAttribute(pageSetupPr, "fitToPage", false),
    pageSetupFitToWidth: pageSetupNode?.hasAttribute("fitToWidth") ? excelFitToPageValue(pageSetupNode.getAttribute("fitToWidth")) : "auto",
    pageSetupFitToHeight: pageSetupNode?.hasAttribute("fitToHeight") ? excelFitToPageValue(pageSetupNode.getAttribute("fitToHeight")) : "auto",
  };
  const sheetView = xmlElements(doc.documentElement, "sheetView")[0];
  if (!sheetView) return pageSetupDetails;
  const pane = firstXmlChild(sheetView, "pane");
  const viewAttr = String(sheetView.getAttribute("view") || "");
  const zoomScaleAttr = sheetView.getAttribute("zoomScale") || sheetView.getAttribute("scale");
  const selection = customSheetViewSelectionFromNode(sheetView, sheetIndex);
  return {
    viewMode: viewAttr === "pageLayout"
      ? "page-layout"
      : viewAttr === "pageBreakPreview"
        ? "page-break"
        : "normal",
    zoomScale: zoomScaleAttr ? clampZoom(zoomScaleAttr) : null,
    showGridLines: xmlBoolAttribute(sheetView, "showGridLines", true),
    showHeadings: xmlBoolAttribute(sheetView, "showRowColHeaders", true),
    showRuler: xmlBoolAttribute(sheetView, "showRuler", false),
    showFormulas: sheetView.getAttribute("showFormulas") === "1",
    showOutlineSymbols: sheetView.getAttribute("showOutlineSymbols") !== "0",
    showZeros: sheetView.getAttribute("showZeros") !== "0",
    showWhiteSpace: sheetView.getAttribute("showWhiteSpace") !== "0",
    rightToLeft: sheetView.getAttribute("rightToLeft") === "1",
    tabSelected: sheetView.getAttribute("tabSelected") === "1",
    activePane: pane?.getAttribute("activePane") || "",
    selectedCell: selection.selectedCell,
    selectionRange: selection.selectionRange,
    selectionRanges: selection.selectionRanges,
    topLeftCell: normalizeCustomTopLeftCell(sheetView.getAttribute("topLeftCell")) || normalizeCustomTopLeftCell(pane?.getAttribute("topLeftCell")),
    freezePanes: customSheetViewPaneFromNode(sheetView, "frozen"),
    splitPane: customSheetViewPaneFromNode(sheetView, "split"),
    ...pageSetupDetails,
  };
}



async function extractWorkbookCustomViews(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
    const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
    if (!workbookXml || !workbookRelsXml) return [];
    const sheetPaths = workbookSheetPaths(workbookXml, workbookRelsXml);
    const workbookDoc = parseXmlDocument(workbookXml);
    const workbookViewNodes = xmlElements(workbookDoc.documentElement, "customWorkbookView");
    if (!workbookViewNodes.length) return [];
    const sheetViewsByGuid = await extractWorkbookCustomSheetViews(zip, sheetPaths);
    return workbookViewNodes
      .map((node, index) => customWorkbookViewNodeToModel(node, sheetPaths, sheetViewsByGuid, index))
      .filter(Boolean);
  } catch (error) {
    console.warn("Failed to extract workbook custom views.", error);
    return [];
  }
}



async function extractWorkbookCustomSheetViews(zip, sheetPaths) {
  const byGuid = new Map();
  await Promise.all(sheetPaths.map(async (sheetInfo, sheetIndex) => {
    const xml = await zip.file(sheetInfo.path)?.async("text");
    if (!xml) return;
    const doc = parseXmlDocument(xml);
    xmlElements(doc.documentElement, "customSheetView").forEach((node) => {
      const guid = normalizeCustomViewGuid(node.getAttribute("guid") || `sheet-${sheetIndex}`);
      if (!byGuid.has(guid)) byGuid.set(guid, new Map());
      byGuid.get(guid).set(sheetIndex, customSheetViewNodeToModel(node, sheetInfo.name, sheetIndex));
    });
  }));
  return byGuid;
}



function customWorkbookViewNodeToModel(node, sheetPaths, sheetViewsByGuid, index = 0) {
  const name = String(node.getAttribute("name") || `ビュー ${index + 1}`).trim();
  if (!name) return null;
  const guid = normalizeCustomViewGuid(node.getAttribute("guid") || name);
  const activeSheetIndex = clampWorkbookSheetIndex(
    node.getAttribute("activeSheetId") ?? node.getAttribute("activeTab") ?? 0,
    sheetPaths.length || 1,
  );
  const sheetSnapshots = sheetViewsByGuid.get(guid) || new Map();
  const activeSheetSnapshot = sheetSnapshots.get(activeSheetIndex) || null;
  const selectedCell = activeSheetSnapshot?.selectedCell || null;
  const selectionRange = activeSheetSnapshot?.selectionRange || null;
  return {
    name,
    guid,
    createdAt: new Date().toISOString(),
    activeSheetIndex,
    autoUpdate: xmlBoolAttribute(node, "autoUpdate", false),
    mergeInterval: Math.max(0, Math.round(Number(node.getAttribute("mergeInterval")) || 0)),
    changesSavedWin: xmlBoolAttribute(node, "changesSavedWin", false),
    onlySync: xmlBoolAttribute(node, "onlySync", false),
    personalView: xmlBoolAttribute(node, "personalView", false),
    showFormulaBar: xmlBoolAttribute(node, "showFormulaBar", true),
    includePrintSettings: xmlBoolAttribute(node, "includePrintSettings", true),
    includeHiddenRowsColumnsAndFilters: xmlBoolAttribute(node, "includeHiddenRowCol", true),
    selectedCell,
    selectionRange,
    selectionRanges: activeSheetSnapshot?.selectionRanges || (selectionRange ? [selectionRange] : []),
    activeScroll: customViewScrollFromSheetView(activeSheetSnapshot),
    topLeftCell: normalizeCustomTopLeftCell(activeSheetSnapshot?.topLeftCell),
    activeWindowId: "main",
    mainWindowHidden: xmlBoolAttribute(node, "hidden", xmlBoolAttribute(node, "minimized", false)),
    mainWindowView: workbookWindowViewFromCustomWorkbookViewNode(node, activeSheetIndex, activeSheetSnapshot, sheetPaths.length || 1),
    windowViews: [],
    workbookWindowsCaptured: false,
    viewSideBySide: null,
    synchronousScrolling: false,
    windowArrangement: null,
    sheets: sheetPaths.map((sheetInfo, sheetIndex) => {
      const snapshot = sheetSnapshots.get(sheetIndex) || {};
      return {
        name: sheetInfo.name,
        viewMode: normalizedSheetViewMode(snapshot.viewMode),
        zoomScale: snapshot.zoomScale || 100,
        showGridLines: snapshot.showGridLines !== false,
        showHeadings: snapshot.showHeadings !== false,
        showRuler: normalizedSheetShowRuler(snapshot.showRuler, snapshot.viewMode),
        showObjects: snapshot.showObjects !== false,
        showOutlineSymbols: snapshot.showOutlineSymbols !== false,
        showFormulas: snapshot.showFormulas === true,
        showZeros: snapshot.showZeros !== false,
        showWhiteSpace: snapshot.showWhiteSpace !== false,
        rightToLeft: snapshot.rightToLeft === true,
        selectedCell: snapshot.selectedCell || null,
        selectionRange: snapshot.selectionRange || null,
        selectionRanges: snapshot.selectionRanges || [],
        activeScroll: customViewScrollFromSheetView(snapshot),
        topLeftCell: normalizeCustomTopLeftCell(snapshot.topLeftCell),
        freezePanes: snapshot.freezePanes || null,
        splitPane: snapshot.splitPane || null,
        pageLayout: snapshot.pageLayout || null,
        headerFooter: snapshot.headerFooter || null,
        sheetViews: [],
        activeSheetViewId: null,
        hiddenRows: [],
        hiddenColumns: [],
        autoFilter: null,
        sortState: null,
      };
    }),
  };
}

export {
  adjustColorLuminosity,
  anchorPointFromAbsoluteDrawingNode,
  anchorPointFromDrawingNode,
  applyDrawingColorAlpha,
  applyWorkbookCellHyperlinksToModel,
  applyWorksheetCellHyperlinkToModel,
  customWorkbookViewNodeToModel,
  dataValidationRuleFromXmlElement,
  decodeExcelCustomFilterWildcard,
  drawingAlphaValue,
  drawingAnchorWithPixelOffset,
  drawingCalloutPositionFromGuide,
  drawingChartAreaStyles,
  drawingChartAxes,
  drawingChartAxisStyle,
  drawingChartCacheValues,
  drawingChartDataLabels,
  drawingChartFromXml,
  drawingChartGrouping,
  drawingChartLabels,
  drawingChartLabelsReference,
  drawingChartLegend,
  drawingChartLineStyle,
  drawingChartMarker,
  drawingChartPlotNode,
  drawingChartReference,
  drawingChartSeries,
  drawingChartSeriesName,
  drawingChartSeriesNameReference,
  drawingChartSeriesStyle,
  drawingChartShapeStyle,
  drawingChartText,
  drawingChartTextStyle,
  drawingChartTitle,
  drawingChartTitleReference,
  drawingChartTitleStyle,
  drawingChartType,
  drawingColorToCss,
  drawingColorTransforms,
  drawingConnectionSite,
  drawingCropPercent,
  drawingGraphicFrameChartResource,
  drawingGraphicFrameKind,
  drawingGraphicFrameRelationshipId,
  drawingGraphicFrameStyle,
  drawingGraphicFrameToModelWithGeometry,
  drawingGroupShapesToModels,
  drawingGuideValue,
  drawingLineEndHasArrow,
  drawingLineEndStyle,
  drawingLineStyleWithBakedFlip,
  drawingMediaExtension,
  drawingMediaName,
  drawingPictureCrop,
  drawingPictureRelationshipId,
  drawingPictureStyleWithGroup,
  drawingPictureToModelWithGeometry,
  drawingPictureVisualStyle,
  drawingRelationshipMedia,
  drawingRelationshipResources,
  drawingRunUnderline,
  drawingShapeAdjustments,
  drawingShapeConnectionEndpoint,
  drawingShapeConnectorConnections,
  drawingShapeCustomPath,
  drawingShapeDescriptorStyle,
  drawingShapeDrawingId,
  drawingShapeEffects,
  drawingShapeFillColor,
  drawingShapeFillGradient,
  drawingShapeFillPattern,
  drawingShapeGeometryFromAnchor,
  drawingShapeGeometryFromGroupChild,
  drawingShapeGlowEffect,
  drawingShapeGuideValues,
  drawingShapeLineDash,
  drawingShapeLineDashLegacy,
  drawingShapeLineJoin,
  drawingShapeLinePoints,
  drawingShapeName,
  drawingShapeNodesInGroup,
  drawingShapeOutlineColor,
  drawingShapePathCoordinate,
  drawingShapePreset,
  drawingShapeReflectionEffect,
  drawingShapeShadow,
  drawingShapeShadowEffect,
  drawingShapeSoftEdgeEffect,
  drawingShapeStyle,
  drawingShapeStyleWithGroup,
  drawingShapeText,
  drawingShapeTextFontFamily,
  drawingShapeTextMargin,
  drawingShapeTextOutline,
  drawingShapeTextRunStyle,
  drawingShapeTextRuns,
  drawingShapeTextStyle,
  drawingShapeToModel,
  drawingShapeToModelWithGeometry,
  drawingShapeType,
  drawingShapesFromAnchor,
  drawingStarInnerFromGuide,
  emuToPixels,
  emuToPoints,
  excelSortConditionColorToCss,
  excelStyleColorNodeToCss,
  excelWildcardPatternTokens,
  extractWorkbookAutoFilters,
  extractWorkbookCellHyperlinks,
  extractWorkbookCustomSheetViews,
  extractWorkbookCustomViews,
  extractWorkbookDataValidationDetails,
  extractWorkbookDefinedNames,
  extractWorkbookDrawingShapes,
  extractWorkbookMetadata,
  extractWorkbookOutlineDetails,
  extractWorkbookPhonetics,
  extractWorkbookSheetViewDetails,
  extractWorkbookWindowViews,
  extractWorksheetCellHyperlinks,
  extractWorksheetOutlineDetails,
  extractWorksheetViewDetails,
  importedColumnSortBodyMatchesFilter,
  importedSortBodyMatchesFilter,
  importedSortConditionColor,
  importedSortConditionCustomList,
  isQuotedValidationLiteral,
  legacyDrawingShapeShadow,
  mergeImportedAutoFilterStates,
  normalizeAutoFilterXmlOperator,
  normalizeDataValidationFormulaFromXml,
  normalizeImportedSortRange,
  normalizePictureCrop,
  normalizeWorksheetHyperlinkTarget,
  parseAutoFilterColumnCriterion,
  parseAutoFilterDateGroupCriterion,
  parseAutoFilterStateFromElement,
  parseDataValidationShowDropDownAttribute,
  parseDrawingXmlShapes,
  parseDxfFillColor,
  parseDxfFontColor,
  parseDynamicAutoFilterCriterion,
  parseExcelCustomFilterXml,
  parsePreservedContentTypeOverrides,
  parsePreservedWorkbookRelationships,
  parseSharedStringPhonetics,
  parseTableColumnXml,
  parseTableStateXml,
  parseTableStyleInfoXml,
  parseWorkbookDefinedNamesXml,
  parseWorkbookStyleDxfsXml,
  parseWorksheetAutoFilterXml,
  parseWorksheetDataValidationDetailsXml,
  parseWorksheetDrawingShapes,
  parseWorksheetSortConditionXml,
  parseWorksheetSortStateXml,
  parseWorksheetTableStates,
  preferredImportedTableAutoFilterState,
  preservedOoxmlEntryPath,
  presetDrawingColor,
  resolveDrawingShapeConnections,
  safeDrawingShapesFromAnchor,
  sharedStringBaseText,
  shouldTreatDataValidationFormulaAsReference,
  stableObjectId,
  systemDrawingColor,
  themeColorForScheme,
  transformGroupedShapeBox,
  workbookSheetPaths,
  worksheetHyperlinkToModel,
  worksheetOutlineDetailForIndex,
  worksheetOutlineDetailsDimensions,
  worksheetOutlineNodeDetail,
  worksheetOutlineSettingsFromXml,
  worksheetRelationshipTargets,
  worksheetRelationships,
  worksheetRelsPath,
  xmlBooleanAttribute,
  xmlNumber,
};
