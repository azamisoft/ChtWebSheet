// Feature module; shared bindings are imported explicitly.
import {
  COLUMN_HEADER_HEIGHT,
  DEFAULT_THEME_COLORS,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  PAGE_LAYOUT_MARGIN_PRESETS,
  WORKSHEET_IMAGE_OBJECT_URL_MIN_DATA_URL_LENGTH,
  state,
  worksheetImageRenderObjectUrls,
} from "../../app-core.js";
import {
  reapplyAutoFilter,
} from "../data/sort-and-autofilter.js";
import {
  normalizedHeaderFooterSettings,
} from "../drawing/shape-rendering.js";
import {
  shiftFormulaReferences,
} from "../editing/autofill.js";
import {
  normalizeColorValue,
} from "../editing/ink-and-clipboard.js";
import {
  getCellRawInput,
  getDisplayForCell,
  isCellInRange,
  rangeToLabel,
  sheetIndexFromPixel,
  updateCellModel,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  cellHasFill,
  coordinateToPixels,
  normalizeFontName,
} from "../editing/selection-and-cell-text.js";
import {
  cloneJsonLikeValue,
  cloneTables,
  compactSheetDefaultCellCss,
  internCellsInRange,
  mutableBordersObject,
  mutableCssObject,
  storeCellModelForSheet,
} from "../editing/structure-and-history.js";
import {
  ensureCellModel,
  rangesIntersect,
} from "../rendering/grid-and-canvas.js";
import {
  cellKey,
  columnName,
  columnNumber,
  columnWidthToPixels,
  decodeCell,
  decodeRange,
  excelDateObjectToSerial,
  excelSerial1904To1900,
  isDateFormat,
  rowHeightToPixels,
  workbookUses1904DateSystem,
} from "../ui/help-and-utilities.js";
import {
  ensureSheetSize,
  rowHeaderWidthForSheet,
} from "../ui/ribbon-and-links.js";
import {
  clampZoom,
  defaultSheetProtectionOptions,
} from "../ui/sheet-protection-and-translation.js";
import {
  bordersToModel,
  cleanStructuredReferencePart,
  colorToCss,
  explicitCellFontColor,
  extractMerges,
  findWorksheetDimensions,
  fontFamilyForExcel,
  mergedEdgeBorders,
  splitStructuredReferenceParts,
  structuredReferenceColumnRangePart,
  structuredReferenceNameMatches,
  structuredReferenceRows,
  structuredReferenceTokenAt,
  styleToCss,
  tableColumnNamesForStructuredRef,
  unescapeStructuredReferencePart,
  unquoteSheetName,
  validationDefinedName,
  validationDefinedNameVisitKey,
} from "./formulas-and-sheet-tabs.js";
import {
  clampWorkbookSheetIndex,
  debugWorkbookLoadTrace,
  normalizeCustomTopLeftCell,
  normalizeCustomViewScroll,
  normalizePaneState,
  normalizeWorkbookWindowView,
  normalizedSheetShowRuler,
  normalizedSheetViewMode,
} from "./import-model.js";
import {
  sharedStringBaseText,
  worksheetOutlineDetailForIndex,
  worksheetOutlineDetailsDimensions,
} from "./ooxml-drawings.js";
import {
  activeSheet,
  cellDisplayText,
  displayValue,
  getSheetId,
  isOccupiedCell,
} from "./saving-and-export.js";
import {
  excelImportedPageMargins,
  normalizedPageLayoutSettings,
} from "./vba-and-printing.js";


function customSheetViewNodeToModel(node, sheetName, sheetIndex) {
  const viewAttr = String(node.getAttribute("view") || "");
  const viewMode = viewAttr === "pageLayout"
    ? "page-layout"
    : viewAttr === "pageBreakPreview"
      ? "page-break"
      : "normal";
  const selection = customSheetViewSelectionFromNode(node, sheetIndex);
  const pageLayout = customSheetViewPageLayoutFromNode(node);
  const headerFooter = customSheetViewHeaderFooterFromNode(node, pageLayout);
  return {
    name: sheetName,
    viewMode,
    zoomScale: clampZoom(node.getAttribute("scale") || node.getAttribute("zoomScale") || 100),
    showGridLines: xmlBoolAttribute(node, "showGridLines", true),
    showHeadings: xmlBoolAttribute(node, "showRowCol", true),
    showRuler: normalizedSheetShowRuler(xmlBoolAttribute(node, "showRuler", false), viewMode),
    showObjects: true,
    showOutlineSymbols: xmlBoolAttribute(node, "outlineSymbols", true),
    showFormulas: xmlBoolAttribute(node, "showFormulas", false),
    showZeros: xmlBoolAttribute(node, "zeroValues", true),
    showWhiteSpace: xmlBoolAttribute(node, "showWhiteSpace", true),
    rightToLeft: xmlBoolAttribute(node, "rightToLeft", false),
    selectedCell: selection.selectedCell,
    selectionRange: selection.selectionRange,
    selectionRanges: selection.selectionRanges,
    freezePanes: customSheetViewPaneFromNode(node, "frozen"),
    splitPane: customSheetViewPaneFromNode(node, "split"),
    pageLayout,
    headerFooter,
    topLeftCell: decodeCell(node.getAttribute("topLeftCell") || ""),
  };
}



function customSheetViewPageLayoutFromNode(node) {
  const marginsNode = firstXmlChild(node, "pageMargins");
  const setupNode = firstXmlChild(node, "pageSetup");
  const printOptionsNode = firstXmlChild(node, "printOptions");
  if (!marginsNode && !setupNode && !printOptionsNode) return null;
  const margins = marginsNode ? {
    top: marginsNode.getAttribute("top"),
    bottom: marginsNode.getAttribute("bottom"),
    left: marginsNode.getAttribute("left"),
    right: marginsNode.getAttribute("right"),
    header: marginsNode.getAttribute("header"),
    footer: marginsNode.getAttribute("footer"),
  } : {};
  const importedMargins = excelImportedPageMargins(margins);
  return normalizedPageLayoutSettings({
    margins: importedMargins,
    marginPreset: pageLayoutMarginPresetForMargins(importedMargins),
    orientation: setupNode?.getAttribute("orientation") === "landscape" ? "landscape" : "portrait",
    paperSize: excelPaperSizeValue(setupNode?.getAttribute("paperSize")),
    fitToPage: xmlBoolAttribute(node, "fitToPage", false),
    fitToWidth: excelFitToPageValue(setupNode?.getAttribute("fitToWidth")),
    fitToHeight: excelFitToPageValue(setupNode?.getAttribute("fitToHeight")),
    scale: setupNode?.getAttribute("scale"),
    pageOrder: setupNode?.getAttribute("pageOrder"),
    firstPageNumber: xmlBoolAttribute(setupNode, "useFirstPageNumber", false) ? setupNode?.getAttribute("firstPageNumber") : 1,
    centerHorizontal: xmlBoolAttribute(printOptionsNode, "horizontalCentered", false),
    centerVertical: xmlBoolAttribute(printOptionsNode, "verticalCentered", false),
    blackAndWhite: xmlBoolAttribute(setupNode, "blackAndWhite", false),
    draftQuality: xmlBoolAttribute(setupNode, "draft", false),
    printGridLines: xmlBoolAttribute(printOptionsNode, "gridLines", false),
    printHeadings: xmlBoolAttribute(printOptionsNode, "headings", false),
    comments: excelPageCommentsValue(setupNode?.getAttribute("cellComments")),
    cellErrors: excelPageErrorsValue(setupNode?.getAttribute("errors")),
  });
}



function customSheetViewHeaderFooterFromNode(node, pageLayout = null) {
  const headerFooterNode = firstXmlChild(node, "headerFooter");
  if (!headerFooterNode) return null;
  const text = (name) => firstXmlChild(headerFooterNode, name)?.textContent || "";
  return normalizedHeaderFooterSettings({
    oddHeader: excelHeaderFooterSection(text("oddHeader")),
    oddFooter: excelHeaderFooterSection(text("oddFooter")),
    evenHeader: excelHeaderFooterSection(text("evenHeader")),
    evenFooter: excelHeaderFooterSection(text("evenFooter")),
    firstHeader: excelHeaderFooterSection(text("firstHeader")),
    firstFooter: excelHeaderFooterSection(text("firstFooter")),
    differentFirst: xmlBoolAttribute(headerFooterNode, "differentFirst", false),
    differentOddEven: xmlBoolAttribute(headerFooterNode, "differentOddEven", false),
    scaleWithDoc: xmlBoolAttribute(headerFooterNode, "scaleWithDoc", true),
    alignWithMargins: xmlBoolAttribute(headerFooterNode, "alignWithMargins", true),
  }, pageLayout || {});
}



function customSheetViewSelectionFromNode(node, sheetIndex) {
  const selectionNodes = xmlChildren(node, "selection");
  const ranges = [];
  let selectedCell = null;
  selectionNodes.forEach((selection) => {
    const activeCell = decodeCell(selection.getAttribute("activeCell") || "");
    if (!selectedCell && activeCell) selectedCell = { sheetIndex, row: activeCell.row, col: activeCell.col };
    const sqref = String(selection.getAttribute("sqref") || "").trim();
    const refs = sqref ? sqref.split(/\s+/) : [];
    refs.forEach((ref) => {
      const range = decodeSheetViewSqrefRange(ref, sheetIndex);
      if (range) ranges.push({ sheetIndex, ...range });
    });
  });
  const fallbackRange = selectedCell
    ? { sheetIndex, top: selectedCell.row, left: selectedCell.col, bottom: selectedCell.row, right: selectedCell.col }
    : null;
  const selectionRange = ranges[0] || fallbackRange;
  if (!selectedCell && selectionRange) selectedCell = { sheetIndex, row: selectionRange.top, col: selectionRange.left };
  return {
    selectedCell,
    selectionRange,
    selectionRanges: ranges.length ? ranges : selectionRange ? [selectionRange] : [],
  };
}



function decodeSheetViewSqrefRange(ref, sheetIndex = 0) {
  const clean = String(ref || "").replace(/\$/g, "").trim().toUpperCase();
  if (!clean) return null;
  const rowOnly = clean.match(/^(\d+):(\d+)$/);
  if (rowOnly) {
    const top = Math.max(1, Math.min(Number(rowOnly[1]), Number(rowOnly[2])));
    const bottom = Math.max(top, Math.max(Number(rowOnly[1]), Number(rowOnly[2])));
    return { sheetIndex, top, left: 1, bottom, right: EXCEL_MAX_COLS, fullRow: true };
  }
  const colOnly = clean.match(/^([A-Z]+):([A-Z]+)$/);
  if (colOnly) {
    const left = Math.max(1, Math.min(columnNumber(colOnly[1]), columnNumber(colOnly[2])));
    const right = Math.max(left, Math.max(columnNumber(colOnly[1]), columnNumber(colOnly[2])));
    return { sheetIndex, top: 1, left, bottom: EXCEL_MAX_ROWS, right, fullColumn: true };
  }
  const range = decodeRange(clean);
  return range ? { sheetIndex, ...range } : null;
}



function customSheetViewPaneFromNode(node, expectedState) {
  const pane = firstXmlChild(node, "pane");
  if (!pane) return null;
  const state = String(pane.getAttribute("state") || "split");
  if (expectedState === "frozen" && !state.includes("frozen")) return null;
  if (expectedState === "split" && state.includes("frozen")) return null;
  const rawRows = Number(pane.getAttribute("ySplit")) || 0;
  const rawColumns = Number(pane.getAttribute("xSplit")) || 0;
  const rows = expectedState === "split"
    ? Math.round(excelSplitUnitsToPixels(rawRows) / rowHeightToPixels(15))
    : Math.round(rawRows);
  const columns = expectedState === "split"
    ? Math.round(excelSplitUnitsToPixels(rawColumns) / columnWidthToPixels(8.43))
    : Math.round(rawColumns);
  if (!rows && !columns) return null;
  return normalizePaneState({
    rows: Math.max(0, rows),
    columns: Math.max(0, columns),
    row: rows + 1,
    col: columns + 1,
    activePane: pane.getAttribute("activePane") || "",
  });
}



function customViewScrollFromSheetView(sheetSnapshot) {
  const topLeftCell = sheetSnapshot?.topLeftCell;
  if (!topLeftCell) return { top: 0, left: 0 };
  return {
    top: Math.max(0, topLeftCell.row - 1),
    left: Math.max(0, topLeftCell.col - 1),
    unit: "cell",
  };
}



function scrollRecordFromTopLeftCell(sheet, topLeftCell) {
  const cell = normalizeCustomTopLeftCell(topLeftCell);
  if (!sheet || !cell) return null;
  const top = Math.max(0, coordinateToPixels(sheet, "row", cell.row - 1) - COLUMN_HEADER_HEIGHT);
  const left = Math.max(0, coordinateToPixels(sheet, "col", cell.col - 1) - rowHeaderWidthForSheet(sheet));
  return {
    scrollTop: Math.round(top),
    scrollLeft: Math.round(left),
  };
}



function topLeftCellFromScrollRecord(sheet, scroll) {
  if (!sheet) return null;
  const normalized = normalizeCustomViewScroll(scroll);
  const rowPixel = normalized.top + COLUMN_HEADER_HEIGHT;
  const colPixel = normalized.left + rowHeaderWidthForSheet(sheet);
  const row = sheetIndexFromPixel(sheet, "row", rowPixel) || 1;
  const col = sheetIndexFromPixel(sheet, "col", colPixel) || 1;
  return { row: Math.max(1, row), col: Math.max(1, col) };
}



function restoreScrollRecordForView(sheet, scroll, topLeftCell = null) {
  const normalized = normalizeCustomViewScroll(scroll);
  if (normalized.unit === "cell") {
    const cellScroll = scrollRecordFromTopLeftCell(sheet, topLeftCell);
    return cellScroll ? {
      scrollTop: Math.max(cellScroll.scrollTop, normalized.top),
      scrollLeft: Math.max(cellScroll.scrollLeft, normalized.left),
    } : {
      scrollTop: normalized.top,
      scrollLeft: normalized.left,
    };
  }
  return {
    scrollTop: normalized.top,
    scrollLeft: normalized.left,
  };
}



function workbookWindowViewFromCustomWorkbookViewNode(node, activeSheetIndex, sheetSnapshot, sheetCount = 1) {
  const activeScroll = customViewScrollFromSheetView(sheetSnapshot);
  const viewMode = normalizedSheetViewMode(sheetSnapshot?.viewMode);
  return normalizeWorkbookWindowView({
    id: "main",
    name: node.getAttribute("name") || "Book1",
    sheetIndex: activeSheetIndex,
    firstSheet: clampWorkbookSheetIndex(node.getAttribute("firstSheet") ?? 0, sheetCount),
    maximized: xmlBoolAttribute(node, "maximized", false),
    minimized: xmlBoolAttribute(node, "minimized", false),
    showHorizontalScroll: xmlBoolAttribute(node, "showHorizontalScroll", true),
    showVerticalScroll: xmlBoolAttribute(node, "showVerticalScroll", true),
    showSheetTabs: xmlBoolAttribute(node, "showSheetTabs", true),
    viewMode,
    zoomScale: sheetSnapshot?.zoomScale || 100,
    showGridLines: sheetSnapshot?.showGridLines !== false,
    showHeadings: sheetSnapshot?.showHeadings !== false,
    showRuler: normalizedSheetShowRuler(sheetSnapshot?.showRuler, viewMode),
    showObjects: sheetSnapshot?.showObjects !== false,
    showOutlineSymbols: sheetSnapshot?.showOutlineSymbols !== false,
    freezePanes: sheetSnapshot?.freezePanes || null,
    splitPane: sheetSnapshot?.splitPane || null,
    scrollTop: activeScroll.top,
    scrollLeft: activeScroll.left,
    topLeftCell: sheetSnapshot?.topLeftCell || null,
    scrollUnit: "cell",
    selectedCell: sheetSnapshot?.selectedCell || null,
    selectionRange: sheetSnapshot?.selectionRange || null,
    selectionRanges: sheetSnapshot?.selectionRanges || [],
    hidden: xmlBoolAttribute(node, "hidden", xmlBoolAttribute(node, "minimized", false)),
    x: Number(node.getAttribute("xWindow")) || undefined,
    y: Number(node.getAttribute("yWindow")) || undefined,
    width: Number(node.getAttribute("windowWidth")) || undefined,
    height: Number(node.getAttribute("windowHeight")) || undefined,
  });
}



function xmlBoolAttribute(node, name, defaultValue = false) {
  if (!node || !node.hasAttribute(name)) return defaultValue;
  const value = String(node.getAttribute(name) || "").toLowerCase();
  if (value === "1" || value === "true") return true;
  if (value === "0" || value === "false") return false;
  return defaultValue;
}



function normalizeWorkbookRelationshipPath(target) {
  const value = String(target || "").replace(/\\/g, "/");
  if (!value) return "";
  if (value.startsWith("/")) return value.replace(/^\/+/, "");
  if (value.startsWith("xl/")) return value;
  return `xl/${value.replace(/^\.?\//, "")}`.replace(/\/+/g, "/");
}



async function parseWorksheetPhonetics(zip, path, sharedStrings) {
  const xml = await zip.file(path)?.async("text");
  if (!xml) return {};
  const doc = parseXmlDocument(xml);
  const mergeRanges = xmlElements(doc.documentElement, "mergeCell")
    .map((node) => decodeRange(node.getAttribute("ref") || ""))
    .filter(Boolean);
  const cells = {};
  xmlElements(doc.documentElement, "c").forEach((cellNode) => {
    const address = cellNode.getAttribute("r");
    const point = decodeCell(address);
    if (!point) return;
    const type = cellNode.getAttribute("t");
    const stringInfo = type === "inlineStr" ? inlineStringPhonetics(cellNode) : sharedStringForCell(cellNode, type, sharedStrings);
    if (!stringInfo?.phonetics?.length) return;
    const master = mergeMasterForPoint(mergeRanges, point) || point;
    cells[cellKey(master.row, master.col)] = stringInfo.phonetics;
  });
  return cells;
}



function mergeMasterForPoint(mergeRanges, point) {
  const merge = mergeRanges.find((range) => point.row >= range.top && point.row <= range.bottom && point.col >= range.left && point.col <= range.right);
  return merge ? { row: merge.top, col: merge.left } : null;
}



function sharedStringForCell(cellNode, type, sharedStrings) {
  if (type !== "s") return null;
  const index = Number.parseInt(firstXmlChildText(cellNode, "v") || "", 10);
  return Number.isFinite(index) ? sharedStrings[index] || null : null;
}



function inlineStringPhonetics(cellNode) {
  const inline = xmlChildren(cellNode, "is")[0];
  if (!inline) return null;
  const text = sharedStringBaseText(inline);
  const phonetics = xmlChildren(inline, "rPh")
    .map((node) => ({
      start: Number.parseInt(node.getAttribute("sb") || "0", 10),
      end: Number.parseInt(node.getAttribute("eb") || "0", 10),
      text: xmlTextContent(node),
    }))
    .filter((run) => Number.isFinite(run.start) && Number.isFinite(run.end) && run.end > run.start && run.text.trim());
  return { text, phonetics: normalizePhoneticRuns(phonetics, text) };
}



function normalizePhoneticRuns(runs, baseText = "") {
  const textLength = String(baseText || "").length;
  return (runs || [])
    .map((run) => {
      const start = Math.max(0, Number(run?.start ?? run?.sb ?? 0));
      const rawEnd = Number(run?.end ?? run?.eb ?? start);
      const end = textLength ? Math.min(textLength, Math.max(start, rawEnd)) : Math.max(start, rawEnd);
      return {
        start,
        end,
        text: String(run?.text || ""),
      };
    })
    .filter((run) => run.end > run.start && run.text.trim());
}



function parseXmlDocument(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const error = doc.querySelector("parsererror");
  if (error) throw new Error(error.textContent || "XML parse error");
  return doc;
}



function xmlChildren(node, localName = "") {
  return Array.from(node?.children || []).filter((child) => !localName || xmlLocalName(child) === localName);
}



function firstXmlChild(node, localName = "") {
  return xmlChildren(node, localName)[0] || null;
}



function xmlElements(node, localName) {
  return Array.from(node?.getElementsByTagName("*") || []).filter((child) => xmlLocalName(child) === localName);
}



function xmlLocalName(node) {
  return node?.localName || String(node?.nodeName || "").split(":").pop();
}



function xmlTextContent(node) {
  return xmlChildren(node)
    .filter((child) => xmlLocalName(child) === "t" || xmlLocalName(child) === "r")
    .map((child) => {
      if (xmlLocalName(child) === "t") return child.textContent || "";
      return xmlChildren(child, "t")
        .map((textNode) => textNode.textContent || "")
        .join("");
    })
    .join("");
}



function firstXmlChildText(node, localName) {
  return xmlChildren(node, localName)[0]?.textContent || "";
}



function normalizePackageRelationshipPath(basePath, target) {
  const value = String(target || "").replace(/\\/g, "/");
  if (!value) return "";
  if (value.startsWith("/")) return value.replace(/^\/+/, "");
  const baseParts = String(basePath || "")
    .replace(/\\/g, "/")
    .split("/");
  baseParts.pop();
  const parts = [...baseParts, value].join("/").split("/");
  const stack = [];
  parts.forEach((part) => {
    if (!part || part === ".") return;
    if (part === "..") stack.pop();
    else stack.push(part);
  });
  return stack.join("/");
}



function worksheetToModel(
  worksheet,
  themeColors,
  media = [],
  widthMetrics = defaultWorkbookWidthMetrics(),
  phoneticsByCell = {},
  drawingShapes = [],
  importedAutoFilter = null,
  importedDataValidationDetails = null,
  sheetViewDetails = null,
  outlineDetails = null,
  workbookProperties = {},
) {
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:start`);
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:merges:start`);
  const merges = extractMerges(worksheet);
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:merges:done`);
  const normalizedDrawingShapes = Array.isArray(drawingShapes) ? drawingShapes : [];
  const importedTables = normalizeImportedWorksheetTables(importedAutoFilter?.tables);
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:dimensions:start`);
  const dimensions = findWorksheetDimensions(worksheet, merges, widthMetrics, normalizedDrawingShapes);
  const tableDimensions = importedTableDimensions(importedTables);
  const dataValidationDimensions = importedDataValidationDimensions(importedDataValidationDetails);
  const outlineDimensions = worksheetOutlineDetailsDimensions(outlineDetails);
  const rowCount = Math.max(dimensions.rows, tableDimensions.rows, dataValidationDimensions.rows, outlineDimensions.rows, 1);
  const colCount = Math.max(dimensions.cols, tableDimensions.cols, dataValidationDimensions.cols, outlineDimensions.cols, 1);
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:dimensions:done:${rowCount}x${colCount}`);
  const warnings = [];
  const worksheetView = worksheet.views?.[0] || {};
  const showGridLines = Object.prototype.hasOwnProperty.call(sheetViewDetails || {}, "showGridLines")
    ? sheetViewDetails.showGridLines !== false
    : worksheetView.showGridLines !== false;
  const showHeadings = Object.prototype.hasOwnProperty.call(sheetViewDetails || {}, "showHeadings")
    ? sheetViewDetails.showHeadings !== false
    : worksheetView.showRowColHeaders !== false;
  const showFormulas = sheetViewDetails?.showFormulas === true || worksheetView.showFormulas === true;
  const showZeros = sheetViewDetails?.showZeros !== false;
  const showWhiteSpace = sheetViewDetails?.showWhiteSpace !== false;
  const rightToLeft = sheetViewDetails?.rightToLeft === true || worksheetView.rightToLeft === true;
  const tabSelected = sheetViewDetails?.tabSelected === true || worksheetView.tabSelected === true;
  const viewMode = sheetViewDetails?.viewMode || (worksheetView.style === "pageLayout"
    ? "page-layout"
    : worksheetView.style === "pageBreakPreview"
      ? "page-break"
      : "normal");
  const freezePanes = sheetViewDetails?.freezePanes || (worksheetView.state === "frozen"
    ? normalizePaneState({
        rows: worksheetView.ySplit || 0,
        columns: worksheetView.xSplit || 0,
        activePane: worksheetView.activePane || sheetViewDetails?.activePane || "",
      })
    : null);
  const splitPaneView = worksheetView.state === "split" ? worksheetView : null;
  const zoomScale = sheetViewDetails?.zoomScale || worksheetZoomScale(worksheet);
  const displayScale = 1;
  const pageLayout = worksheetPageLayoutToModel(worksheet, { rowCount, colCount }, sheetViewDetails);
  const headerFooter = excelHeaderFooterToModel(worksheet.headerFooter, pageLayout);

  const defaultColWidth = worksheet.properties?.defaultColWidth || 8.43;
  const defaultColumnWidthPixels = columnWidthToPixels(defaultColWidth, displayScale, widthMetrics.maxDigitWidth);
  const columns = [];
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:columns:start:${colCount}`);
  for (let col = 1; col <= colCount; col += 1) {
    const column = worksheet.getColumn(col);
    const outlineDetail = worksheetOutlineDetailForIndex(outlineDetails, "columns", col);
    const outlineLevel = clampOutlineLevel(outlineDetail?.outlineLevel ?? column.outlineLevel ?? column.model?.outlineLevel);
    const hidden = Object.prototype.hasOwnProperty.call(outlineDetail || {}, "hidden")
      ? outlineDetail.hidden === true
      : Boolean(column.hidden);
    const width = columnWidthToPixels(column.width ?? defaultColWidth, displayScale, widthMetrics.maxDigitWidth);
    columns.push({
      width,
      hidden,
      outlineLevel,
      outlineHidden: Boolean((outlineDetail?.outlineHidden || hidden) && outlineLevel),
      outlineCollapsed: outlineDetail?.outlineCollapsed === true,
    });
  }
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:columns:done`);

  const defaultRowHeight = worksheet.properties?.defaultRowHeight || 15;
  const defaultRowHeightPixels = rowHeightToPixels(defaultRowHeight, displayScale);
  const mergeByMaster = new Map(merges.map((merge) => [cellKey(merge.top, merge.left), merge]));
  const rows = [];
  const cells = {};

  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:rows:start:${rowCount}x${colCount}`);
  for (let rowNumber = 1; rowNumber <= rowCount; rowNumber += 1) {
    if (rowNumber === 1 || rowNumber % 50 === 0 || rowNumber === rowCount) {
      debugWorkbookLoadTrace(`worksheet:${worksheet.name}:rows:at:${rowNumber}/${rowCount}`);
    }
    const row = worksheet.getRow(rowNumber);
    const outlineDetail = worksheetOutlineDetailForIndex(outlineDetails, "rows", rowNumber);
    const outlineLevel = clampOutlineLevel(outlineDetail?.outlineLevel ?? row.outlineLevel ?? row.model?.outlineLevel);
    const hidden = Object.prototype.hasOwnProperty.call(outlineDetail || {}, "hidden")
      ? outlineDetail.hidden === true
      : Boolean(row.hidden);
    const height = rowHeightToPixels(row.height ?? defaultRowHeight, displayScale);
    rows.push({
      height,
      hidden,
      outlineLevel,
      outlineHidden: Boolean((outlineDetail?.outlineHidden || hidden) && outlineLevel),
      outlineCollapsed: outlineDetail?.outlineCollapsed === true,
    });

    const hfRow = [];
    for (let colNumber = 1; colNumber <= colCount; colNumber += 1) {
      const cell = worksheet.getCell(rowNumber, colNumber);
      const cellModel = cellToModel(cell, rowNumber, colNumber, themeColors, workbookProperties, worksheet.name);
      const phonetics = normalizePhoneticRuns(phoneticsByCell[cellKey(rowNumber, colNumber)]);
      if (phonetics.length) {
        cellModel.phonetics = phonetics;
        cellModel.showPhonetic = false;
      }
      const merge = mergeByMaster.get(cellKey(rowNumber, colNumber));
      if (merge) {
        cellModel.borders = mergedEdgeBorders(worksheet, merge, themeColors);
      }

      if (shouldStoreCell(cellModel)) {
      cells[cellKey(rowNumber, colNumber)] = cellModel;
      }
    }
  }
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:rows:done:${Object.keys(cells).length}`);

  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:views-and-images:start`);
  const splitPane = sheetViewDetails?.splitPane || (splitPaneView
    ? splitPaneFromWorksheetView(splitPaneView, { columns, rows, defaultColumnWidthPixels, defaultRowHeightPixels })
    : null);
  const worksheetImages = worksheetImagesToModel(worksheet, media, displayScale);
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:views-and-images:done`);
  const autoFilter = importedAutoFilter?.range
    ? {
        range: importedAutoFilter.range,
        criteria: cloneJsonLikeValue(importedAutoFilter.criteria || {}),
        ...(importedAutoFilter.hiddenButtons ? { hiddenButtons: cloneJsonLikeValue(importedAutoFilter.hiddenButtons) } : {}),
      }
    : worksheet.autoFilter
      ? { range: excelAutoFilterRangeToModel(worksheet.autoFilter), criteria: {} }
      : null;
  const sortState = importedAutoFilter?.sortState ? cloneJsonLikeValue(importedAutoFilter.sortState) : null;
  const protection = worksheetProtectionToModel(worksheet);
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:sheet-model:start`);
  const sheetModel = {
    name: worksheet.name,
    state: worksheet.state || "visible",
    tabColor: colorToCss(worksheet.properties?.tabColor, themeColors) || "",
    rowCount,
    colCount,
    baseRowCount: rowCount,
    baseColCount: colCount,
    defaultColumnWidth: defaultColumnWidthPixels,
    defaultRowHeight: defaultRowHeightPixels,
    zoomScale,
    baseZoomScale: zoomScale,
    displayScale,
    widthMetrics,
    showGridLines,
    showHeadings,
    showRuler: normalizedSheetShowRuler(
      Object.prototype.hasOwnProperty.call(sheetViewDetails || {}, "showRuler") ? sheetViewDetails.showRuler : worksheetView.showRuler,
      viewMode,
    ),
    showObjects: workbookProperties.showObjects !== "none",
    showOutlineSymbols: sheetViewDetails?.showOutlineSymbols !== false && worksheetView.showOutlineSymbols !== false,
    showFormulas,
    showZeros,
    showWhiteSpace,
    rightToLeft,
    tabSelected,
    viewMode,
    freezePanes,
    splitPane,
    selectedCell: sheetViewDetails?.selectedCell || null,
    selectionRange: sheetViewDetails?.selectionRange || null,
    selectionRanges: sheetViewDetails?.selectionRanges || [],
    topLeftCell: normalizeCustomTopLeftCell(sheetViewDetails?.topLeftCell),
    outlineSettings: normalizeOutlineSettings({
      ...worksheetOutlineSettings(worksheet),
      ...(outlineDetails?.settings || {}),
    }),
    columns,
    rows,
    merges: merges.filter((merge) => merge.bottom <= rowCount && merge.right <= colCount),
    images: mergeWorksheetImagesAndDrawingObjects(worksheetImages, normalizedDrawingShapes),
    tables: importedTables,
    autoFilter,
    sortState,
    protection,
    pageLayout,
    headerFooter,
    cells,
    hfData: [],
    warnings,
  };
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:sheet-model:done`);
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:postprocess:start`);
  applyImportedDataValidationDetails(sheetModel, importedDataValidationDetails);
  applyImportedTableVisualStyles(sheetModel, importedTables);
  applyImportedTableCalculatedColumns(sheetModel, importedTables);
  applyImportedTableTotals(sheetModel, importedTables);
  materializeImportedAutoFilterCriteria(sheetModel);
  compactSheetDefaultCellCss(sheetModel);
  debugWorkbookLoadTrace(`worksheet:${worksheet.name}:postprocess:done`);
  return sheetModel;
}



function worksheetProtectionToModel(worksheet) {
  const protection = worksheet?.sheetProtection || worksheet?.model?.sheetProtection;
  if (!protection?.sheet) return null;
  return defaultSheetProtectionOptions({
    password: "",
    allowSelectLockedCells: protection.selectLockedCells !== false,
    allowSelectUnlockedCells: protection.selectUnlockedCells !== false,
    allowFormatCells: Boolean(protection.formatCells),
    allowFormatColumns: Boolean(protection.formatColumns),
    allowFormatRows: Boolean(protection.formatRows),
    allowInsertColumns: Boolean(protection.insertColumns),
    allowInsertRows: Boolean(protection.insertRows),
    allowDeleteColumns: Boolean(protection.deleteColumns),
    allowDeleteRows: Boolean(protection.deleteRows),
    allowSort: Boolean(protection.sort),
    allowAutoFilter: Boolean(protection.autoFilter),
    allowPivotTables: Boolean(protection.pivotTables),
    allowEditObjects: protection.objects !== false,
    allowEditScenarios: protection.scenarios !== false,
  });
}



function clampOutlineLevel(value) {
  return Math.max(0, Math.min(8, Number(value) || 0));
}



function normalizeOutlineSettings(settings = {}) {
  return {
    summaryBelow: settings?.summaryBelow !== false,
    summaryRight: settings?.summaryRight !== false,
    applyStyles: settings?.applyStyles === true || settings?.autoStyle === true,
  };
}



function worksheetOutlineSettings(worksheet) {
  const outline = worksheet?.properties?.outlineProperties || {};
  return normalizeOutlineSettings(outline);
}



function worksheetPageLayoutToModel(worksheet, sheet = null, sheetViewDetails = null) {
  const setup = worksheet?.pageSetup || {};
  const margins = excelImportedPageMargins(setup.margins || worksheet?.pageMargins || {});
  const fitToWidth = Object.prototype.hasOwnProperty.call(sheetViewDetails || {}, "pageSetupFitToWidth")
    ? sheetViewDetails.pageSetupFitToWidth
    : excelFitToPageValue(setup.fitToWidth);
  const fitToHeight = Object.prototype.hasOwnProperty.call(sheetViewDetails || {}, "pageSetupFitToHeight")
    ? sheetViewDetails.pageSetupFitToHeight
    : excelFitToPageValue(setup.fitToHeight);
  return normalizedPageLayoutSettings({
    margins: {
      top: margins.top,
      bottom: margins.bottom,
      left: margins.left,
      right: margins.right,
      header: margins.header,
      footer: margins.footer,
    },
    marginPreset: pageLayoutMarginPresetForMargins(margins),
    orientation: setup.orientation === "landscape" ? "landscape" : "portrait",
    paperSize: excelPaperSizeValue(setup.paperSize),
    fitToPage: (setup.fitToPage === true || sheetViewDetails?.pageSetupFitToPage === true) && (fitToWidth !== "auto" || fitToHeight !== "auto"),
    fitToWidth,
    fitToHeight,
    scale: setup.scale,
    pageOrder: setup.pageOrder,
    firstPageNumber: setup.useFirstPageNumber === false ? 1 : setup.firstPageNumber,
    centerHorizontal: Boolean(setup.horizontalCentered),
    centerVertical: Boolean(setup.verticalCentered),
    blackAndWhite: Boolean(setup.blackAndWhite),
    draftQuality: Boolean(setup.draft),
    printGridLines: Boolean(setup.showGridLines),
    printHeadings: Boolean(setup.showRowColHeaders),
    comments: excelPageCommentsValue(setup.cellComments),
    cellErrors: excelPageErrorsValue(setup.errors),
    pageBreaks: {
      rows: excelManualBreakIndexes(worksheet?.rowBreaks),
      cols: excelManualBreakIndexes(worksheet?.colBreaks),
    },
  }, sheet);
}



function pageLayoutMarginPresetForMargins(margins = {}) {
  const numeric = (value) => Number.isFinite(Number(value)) ? Math.round(Number(value) * 1000) / 1000 : null;
  const keys = ["top", "bottom", "left", "right", "header", "footer"];
  const current = Object.fromEntries(keys.map((key) => [key, numeric(margins[key])]));
  const exact = Object.entries(PAGE_LAYOUT_MARGIN_PRESETS).find(([_key, preset]) =>
    keys.every((field) => current[field] != null && Math.abs(current[field] - numeric(preset[field])) < 0.001),
  );
  return exact?.[0] || "normal";
}



function excelManualBreakIndexes(breaks) {
  return [...new Set((Array.isArray(breaks) ? breaks : [])
    .map((item) => Number(item?.id ?? item))
    .filter((id) => Number.isFinite(id) && id >= 1)
    .map((id) => Math.max(2, Math.round(id) + 1)))]
    .sort((a, b) => a - b);
}



function excelPaperSizeValue(value) {
  const code = Number(value);
  if (code === 1) return "Letter";
  if (code === 8) return "A3";
  if (code === 13) return "B5";
  return "A4";
}



function excelFitToPageValue(value) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return "auto";
  return String(Math.max(1, Math.min(99, Math.round(number))));
}



function excelPageCommentsValue(value) {
  const text = String(value || "").toLowerCase();
  if (text === "atend") return "atEnd";
  if (text === "asdisplayed" || text === "asdisplyed") return "asDisplayed";
  return "none";
}



function excelPageErrorsValue(value) {
  const text = String(value || "").toLowerCase();
  if (text === "blank") return "blank";
  if (text === "dash") return "dash";
  if (text === "na") return "na";
  return "displayed";
}



function excelHeaderFooterToModel(headerFooter = null, pageLayout = {}) {
  const source = headerFooter || {};
  const normalized = normalizedHeaderFooterSettings({
    oddHeader: excelHeaderFooterSection(source.oddHeader),
    oddFooter: excelHeaderFooterSection(source.oddFooter),
    firstHeader: excelHeaderFooterSection(source.firstHeader),
    firstFooter: excelHeaderFooterSection(source.firstFooter),
    evenHeader: excelHeaderFooterSection(source.evenHeader),
    evenFooter: excelHeaderFooterSection(source.evenFooter),
    differentFirst: Boolean(source.differentFirst),
    differentOddEven: Boolean(source.differentOddEven),
    scaleWithDoc: source.scaleWithDoc !== false,
    alignWithMargins: source.alignWithMargins !== false,
  }, pageLayout);
  return normalized;
}



function excelHeaderFooterSection(text) {
  const value = String(text || "");
  if (!value) return { left: "", center: "", right: "" };
  const section = { left: "", center: "", right: "" };
  let key = "center";
  let index = 0;
  while (index < value.length) {
    if (value[index] === "&" && index + 1 < value.length) {
      const code = value[index + 1];
      if (code === "L" || code === "C" || code === "R") {
        key = code === "L" ? "left" : code === "R" ? "right" : "center";
        index += 2;
        continue;
      }
      if (code === "&") {
        section[key] += "&";
        index += 2;
        continue;
      }
    }
    section[key] += value[index];
    index += 1;
  }
  return section;
}



function normalizeImportedWorksheetTables(tables) {
  return cloneTables(Array.isArray(tables) ? tables : []).filter((table) => table?.range);
}



function importedDataValidationDimensions(details = []) {
  const maxCellsForDimensionExpansion = 20000;
  return (Array.isArray(details) ? details : []).reduce(
    (dimensions, detail) => {
      (Array.isArray(detail?.ranges) ? detail.ranges : []).forEach((range) => {
        if (!range) return;
        const rows = Math.max(0, (Number(range.bottom) || 0) - (Number(range.top) || 0) + 1);
        const cols = Math.max(0, (Number(range.right) || 0) - (Number(range.left) || 0) + 1);
        if (rows * cols > maxCellsForDimensionExpansion) return;
        dimensions.rows = Math.max(dimensions.rows, Number(range.bottom) || 0);
        dimensions.cols = Math.max(dimensions.cols, Number(range.right) || 0);
      });
      return dimensions;
    },
    { rows: 0, cols: 0 },
  );
}



function applyImportedDataValidationDetails(sheet, details = []) {
  if (!sheet || !Array.isArray(details) || !details.length) return;
  details.forEach((detail) => {
    const ranges = Array.isArray(detail?.ranges) ? detail.ranges : [];
    if (!ranges.length) return;
    const importedValidation = detail.validation && typeof detail.validation === "object" ? detail.validation : null;
    const hasShowDropDown = Object.prototype.hasOwnProperty.call(detail, "showDropDown");
    const hasImeMode = Object.prototype.hasOwnProperty.call(detail, "imeMode") && detail.imeMode;
    if (!importedValidation && !hasShowDropDown && !hasImeMode) return;
    ranges.forEach((range) => {
      if (!range) return;
      const top = Math.max(1, Number(range.top) || 1);
      const left = Math.max(1, Number(range.left) || 1);
      const bottom = Math.min(Number(sheet.rowCount) || top, Math.max(top, Number(range.bottom) || top));
      const right = Math.min(Number(sheet.colCount) || left, Math.max(left, Number(range.right) || left));
      for (let row = top; row <= bottom; row += 1) {
        for (let col = left; col <= right; col += 1) {
          const key = cellKey(row, col);
          const cell = sheet.cells?.[key] || (importedValidation ? ensureCellModel(sheet, row, col) : null);
          if (!cell?.validation && !importedValidation) continue;
          cell.validation = {
            ...(cell.validation || {}),
            ...(importedValidation || {}),
            ...(hasShowDropDown ? { showDropDown: detail.showDropDown !== false } : {}),
            ...(hasImeMode ? { imeMode: detail.imeMode } : {}),
          };
        }
      }
    });
  });
}



function importedTableDimensions(tables) {
  return (tables || []).reduce(
    (dimensions, table) => ({
      rows: Math.max(dimensions.rows, Number(table?.range?.bottom) || 0),
      cols: Math.max(dimensions.cols, Number(table?.range?.right) || 0),
    }),
    { rows: 0, cols: 0 },
  );
}



function applyImportedTableVisualStyles(sheet, tables) {
  (tables || []).forEach((table) => {
    if (!table?.range || !table.style) return;
    applyImportedTableVisualStyle(sheet, table);
  });
}



function applyImportedTableVisualStyle(sheet, table) {
  const range = table.range;
  const visualCellCount = (range.bottom - range.top + 1) * (range.right - range.left + 1);
  const rowLimit = visualCellCount > 50000 ? new Set([range.top, table.showTotalsRow === true ? range.bottom : range.top]) : null;
  const palette = importedTableStylePalette(table.style || table.styleInfo?.name || "");
  const showHeader = table.showHeaderRow !== false;
  const showTotals = table.showTotalsRow === true;
  const showRowStripes = table.styleInfo?.showRowStripes !== false;
  const showColumnStripes = table.styleInfo?.showColumnStripes === true;
  const border = { width: 1, style: "solid", color: palette.border, weight: 2 };
  for (let row = range.top; row <= range.bottom; row += 1) {
    if (rowLimit && !rowLimit.has(row)) continue;
    for (let col = range.left; col <= range.right; col += 1) {
      const cell = ensureCellModel(sheet, row, col);
      cell.css.fontFamily = cell.css.fontFamily || sheet.defaultCellCss?.fontFamily || fontFamilyForExcel("Calibri");
      const isHeader = showHeader && row === range.top;
      const isTotals = showTotals && row === range.bottom;
      const bodyRowIndex = row - range.top - (showHeader ? 1 : 0);
      if (isHeader) {
        applyImportedTableFillIfEmpty(cell, palette.headerFill);
        if (!cell.hasFontColor && !cell.css.color) {
          cell.css.color = palette.headerText;
          cell.hasFontColor = true;
        }
        cell.css.fontWeight = cell.css.fontWeight || "700";
      } else if (isTotals) {
        applyImportedTableFillIfEmpty(cell, palette.totalFill);
        cell.css.fontWeight = cell.css.fontWeight || "700";
      } else if (showRowStripes && bodyRowIndex >= 0 && bodyRowIndex % 2 === 1) {
        applyImportedTableFillIfEmpty(cell, palette.stripeFill);
      } else if (showColumnStripes && (col - range.left) % 2 === 1) {
        applyImportedTableFillIfEmpty(cell, palette.stripeFill);
      }
      cell.borders = mergeMissingTableBorders(cell.borders, border);
    }
  }
  if (rowLimit) {
    rowLimit.forEach((row) => internCellsInRange(sheet, { ...range, top: row, bottom: row }));
  } else {
    internCellsInRange(sheet, range);
  }
}



function refreshImportedTableVisualStyleForTableMutation(sheet, table) {
  if (!sheet || !table?.range || !table.style) return;
  const range = table.range;
  const visualCellCount = (range.bottom - range.top + 1) * (range.right - range.left + 1);
  if (visualCellCount > 50000) return;
  const palette = importedTableStylePalette(table.style || table.styleInfo?.name || "");
  const tableFillColors = new Set(
    [palette.headerFill, palette.stripeFill, palette.totalFill]
      .map((color) => normalizeColorValue(color))
      .filter(Boolean),
  );
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      const key = cellKey(row, col);
      const cell = sheet.cells[key];
      const fill = normalizeColorValue(cell?.css?.backgroundColor);
      if (!cell || !fill || !tableFillColors.has(fill)) continue;
      cell.css = mutableCssObject({ ...(cell.css || {}) });
      delete cell.css.backgroundColor;
      cell.hasFill = false;
      storeCellModelForSheet(sheet, key, cell);
    }
  }
  applyImportedTableVisualStyle(sheet, table);
}



function applyImportedTableFillIfEmpty(cell, fill) {
  if (!fill || cellHasFill(cell)) return;
  cell.css.backgroundColor = fill;
  cell.hasFill = true;
}



function mergeMissingTableBorders(existing = {}, border) {
  return {
    top: existing.top || { ...border },
    right: existing.right || { ...border },
    bottom: existing.bottom || { ...border },
    left: existing.left || { ...border },
  };
}



function importedTableStylePalette(styleName = "") {
  const text = String(styleName || "");
  const mediumMatch = text.match(/Medium(\d+)/i);
  const lightMatch = text.match(/Light(\d+)/i);
  const darkMatch = text.match(/Dark(\d+)/i);
  const index = Number(mediumMatch?.[1] || lightMatch?.[1] || darkMatch?.[1] || 2);
  const palettes = [
    { headerFill: "#5b9bd5", headerText: "#ffffff", stripeFill: "#ddebf7", totalFill: "#bdd7ee", border: "#9cc2e5" },
    { headerFill: "#70ad47", headerText: "#ffffff", stripeFill: "#e2f0d9", totalFill: "#c6e0b4", border: "#a9d18e" },
    { headerFill: "#ed7d31", headerText: "#ffffff", stripeFill: "#fce4d6", totalFill: "#f8cbad", border: "#f4b183" },
    { headerFill: "#a5a5a5", headerText: "#ffffff", stripeFill: "#ededed", totalFill: "#d9d9d9", border: "#bfbfbf" },
    { headerFill: "#ffc000", headerText: "#1f2937", stripeFill: "#fff2cc", totalFill: "#ffe699", border: "#ffd966" },
    { headerFill: "#4472c4", headerText: "#ffffff", stripeFill: "#d9eaf7", totalFill: "#cfe2f3", border: "#9cc2e5" },
  ];
  const palette = palettes[Math.max(0, index - 1) % palettes.length];
  if (darkMatch) return { ...palette, stripeFill: palette.totalFill };
  if (lightMatch) return { ...palette, headerFill: palette.stripeFill, headerText: "#1f2937", totalFill: palette.stripeFill };
  return palette;
}



function applyImportedTableTotals(sheet, tables) {
  (tables || []).forEach((table) => {
    if (!table?.range || table.showTotalsRow !== true) return;
    const totalRow = table.range.bottom;
    const columns = tableColumnNamesForStructuredRef(sheet, table);
    (table.columns || []).forEach((column, index) => {
      const col = table.range.left + index;
      if (col > table.range.right) return;
      const label = String(column?.totalsRowLabel || "");
      const fn = String(column?.totalsRowFunction || "").trim();
      const existing = sheet.cells[cellKey(totalRow, col)];
      if (label && !isOccupiedCell(existing)) {
        setImportedTableTotalCell(sheet, totalRow, col, label);
        return;
      }
      const customFormula = normalizeTableFormula(column?.totalsRowFormula || "");
      if (customFormula && !existing?.formula) {
        setImportedTableTotalCell(sheet, totalRow, col, customFormula);
        return;
      }
      const formulaCode = totalsRowSubtotalCode(fn);
      if (!formulaCode || existing?.formula) return;
      const columnNameText = columns[index] || `Column${index + 1}`;
      if (!columnNameText) return;
      setImportedTableTotalCell(sheet, totalRow, col, `=SUBTOTAL(${formulaCode},[${columnNameText}])`);
    });
  });
}



function applyImportedTableCalculatedColumns(sheet, tables) {
  (tables || []).forEach((table) => {
    if (!table?.range) return;
    const dataRows = structuredReferenceRows(table, "data");
    if (!dataRows) return;
    const columns = ensureTableColumnMetadata(sheet, table);
    columns.forEach((column, index) => {
      const formula = normalizeTableFormula(column?.calculatedColumnFormula || "");
      if (!formula) return;
      const col = table.range.left + index;
      if (col > table.range.right) return;
      for (let row = dataRows.top; row <= dataRows.bottom; row += 1) {
        materializeImportedTableCalculatedCell(sheet, table, row, col, formula);
      }
    });
  });
}



function materializeImportedTableCalculatedCell(sheet, table, row, col, formula) {
  const key = cellKey(row, col);
  const existing = sheet.cells?.[key];
  if (existing?.formula || (typeof existing?.raw === "string" && existing.raw.trim().startsWith("="))) return false;
  if (isOccupiedCell(existing)) return false;
  const cell = ensureCellModel(sheet, row, col);
  cell.raw = formula;
  cell.formula = formula;
  cell.cached = null;
  cell.display = "";
  delete cell.richText;
  delete cell.phonetics;
  delete cell.showPhonetic;
  delete cell.sparkline;
  storeCellModelForSheet(sheet, key, cell);
  return true;
}



function totalsRowSubtotalCode(functionName = "") {
  switch (String(functionName || "").toLocaleLowerCase("en-US")) {
    case "average":
      return 101;
    case "countnums":
      return 102;
    case "count":
      return 103;
    case "max":
      return 104;
    case "min":
      return 105;
    case "stddev":
      return 107;
    case "var":
      return 110;
    case "sum":
      return 109;
    default:
      return null;
  }
}



function setImportedTableTotalCell(sheet, row, col, content) {
  const cell = ensureCellModel(sheet, row, col);
  cell.raw = content;
  cell.formula = typeof content === "string" && content.startsWith("=") ? content : null;
  cell.cached = cell.formula ? null : content;
  cell.display = cell.formula ? "" : String(content ?? "");
  storeCellModelForSheet(sheet, cellKey(row, col), cell);
}



function tableEditMayChangeMetadata(sheet, row, col, content) {
  if (!sheet?.tables?.length) return false;
  return Boolean(
    tableExpansionTargetForEdit(sheet, row, col, content) ||
      tableHeaderForCell(sheet, row, col) ||
      tableCalculatedColumnCandidateForEdit(sheet, row, col, content),
  );
}



function applyTableEditBehaviors(sheet, row, col, content) {
  const expansion = applyTableAutoExpansionForEdit(sheet, row, col, content);
  const header = syncTableHeaderNameForEdit(sheet, row, col);
  const calculatedColumn = applyTableCalculatedColumnForEdit(sheet, row, col, content);
  if (!expansion.changed && !header.changed && !calculatedColumn.changed) return { changed: false };
  if (expansion.changed) {
    reapplyAutoFilter({ sheet, sheetIndex: state.activeSheetIndex, silent: true });
  }
  const label = [expansion, header, calculatedColumn].filter((item) => item.changed).map((item) => item.label).join(" / ");
  const status = [expansion, header, calculatedColumn].filter((item) => item.changed).map((item) => item.status).join(" ");
  return { changed: true, label, status };
}



function tableRangeEditMayChangeMetadata(sheet, range, options = {}) {
  if (!sheet?.tables?.length || !range) return false;
  return Boolean(
    tableRangeExpansionTarget(sheet, range, options) ||
      tableHeaderIntersectsRange(sheet, range) ||
      (options.assumeFormula && tableDataRangeIntersectsRange(sheet, range)),
  );
}



function applyTableRangeEditBehaviors(sheet, range) {
  const expansions = [];
  let headerChanged = false;
  const calculatedColumns = [];
  while (true) {
    const target = tableRangeExpansionTarget(sheet, range);
    if (!target) break;
    const oldRange = { ...target.table.range };
    if (!expandTableToRangeForEdit(sheet, target.table, range, target.nextRange, oldRange)) break;
    syncTableAutoFilterRanges(sheet, target.table, oldRange);
    expansions.push(target.table.displayName || target.table.name || "Table");
  }
  (sheet?.tables || []).forEach((table) => {
    if (!table?.range || table.showHeaderRow === false || !rangesIntersect(table.range, range)) return;
    if (syncTableColumnNamesFromHeaders(sheet, table)) headerChanged = true;
  });
  applyTableCalculatedColumnsForRangeEdit(sheet, range).forEach((name) => calculatedColumns.push(name));
  if (!expansions.length && !headerChanged && !calculatedColumns.length) return { changed: false };
  reapplyAutoFilter({ sheet, sheetIndex: state.activeSheetIndex, silent: true });
  const uniqueNames = [...new Set(expansions)];
  const uniqueCalculated = [...new Set(calculatedColumns)];
  const labels = [];
  if (uniqueNames.length) labels.push(`${uniqueNames.join(", ")} テーブル範囲の拡張`);
  if (headerChanged) labels.push("テーブル見出し");
  if (uniqueCalculated.length) labels.push(`${uniqueCalculated.join(", ")} 計算列`);
  const statusParts = [];
  if (uniqueNames.length) statusParts.push(`${uniqueNames.join(", ")} のテーブル範囲を更新しました。`);
  if (headerChanged) statusParts.push("テーブルの見出しを更新しました。");
  if (uniqueCalculated.length) statusParts.push(`${uniqueCalculated.join(", ")} の計算列を更新しました。`);
  return {
    changed: true,
    label: labels.join(" / "),
    status: statusParts.join(" "),
  };
}



function applyTableAutoExpansionForEdit(sheet, row, col, content) {
  const target = tableExpansionTargetForEdit(sheet, row, col, content);
  if (!target) return { changed: false };
  const oldRange = { ...target.table.range };
  const changed = target.kind === "row"
    ? expandTableDownForEdit(sheet, target.table, row, col)
    : expandTableRightForEdit(sheet, target.table, row, col);
  if (!changed) return { changed: false };
  syncTableAutoFilterRanges(sheet, target.table, oldRange);
  const tableName = target.table.displayName || target.table.name || "Table";
  return {
    changed: true,
    label: `${tableName} テーブル範囲の拡張`,
    status: `${tableName} を ${rangeToLabel(target.table.range)} に拡張しました。`,
  };
}



function tableExpansionTargetForEdit(sheet, row, col, content) {
  if (!tableEditContentIsNonEmpty(content)) return null;
  const tables = sheet?.tables || [];
  const downTable = tables.find((table) => tableCanExpandDownForEdit(sheet, table, row, col));
  if (downTable) return { table: downTable, kind: "row" };
  const rightTable = tables.find((table) => tableCanExpandRightForEdit(sheet, table, row, col));
  return rightTable ? { table: rightTable, kind: "column" } : null;
}



function tableEditContentIsNonEmpty(content) {
  return content != null && String(content) !== "";
}



function tableRangeExpansionTarget(sheet, range, options = {}) {
  if (!range) return null;
  if (!options.assumeNonEmpty && !rangeHasEditableTableContent(sheet, range)) return null;
  for (const table of sheet?.tables || []) {
    const tableRange = table?.range;
    if (!tableRange || table.showTotalsRow === true) continue;
    let bottom = tableRange.bottom;
    let right = tableRange.right;
    const touchesBelow =
      range.bottom > tableRange.bottom &&
      range.top <= tableRange.bottom + 1 &&
      range.left <= tableRange.right &&
      range.right >= tableRange.left;
    const touchesRight =
      range.right > tableRange.right &&
      range.left <= tableRange.right + 1 &&
      range.top <= tableRange.bottom &&
      range.bottom >= tableRange.top;
    if (touchesBelow) {
      bottom = Math.max(bottom, range.bottom);
      if (range.left <= tableRange.right + 1) right = Math.max(right, range.right);
    }
    if (touchesRight) {
      right = Math.max(right, range.right);
      if (range.top <= tableRange.bottom + 1) bottom = Math.max(bottom, range.bottom);
    }
    if (bottom === tableRange.bottom && right === tableRange.right) continue;
    const nextRange = { ...tableRange, bottom, right };
    if (!tableExpansionRangeIsAvailable(sheet, table, nextRange)) continue;
    return { table, nextRange };
  }
  return null;
}



function tableDataRangeIntersectsRange(sheet, range) {
  return (sheet?.tables || []).some((table) => {
    const dataRows = structuredReferenceRows(table, "data");
    if (!dataRows) return false;
    const dataRange = { ...table.range, top: dataRows.top, bottom: dataRows.bottom };
    return rangesIntersect(dataRange, range);
  });
}



function rangeHasEditableTableContent(sheet, range) {
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      if (isOccupiedCell(sheet?.cells?.[cellKey(row, col)])) return true;
    }
  }
  return false;
}



function tableHeaderIntersectsRange(sheet, range) {
  return (sheet?.tables || []).some((table) => {
    const tableRange = table?.range;
    return tableRange && table.showHeaderRow !== false && range.top <= tableRange.top && range.bottom >= tableRange.top && range.left <= tableRange.right && range.right >= tableRange.left;
  });
}



function rangeIntersectsWorksheetTable(sheet, range) {
  if (!sheet?.tables?.length || !range) return false;
  return (sheet.tables || []).some((table) => table?.range && rangesIntersect(table.range, range));
}



function tableCanExpandDownForEdit(sheet, table, row, col) {
  const range = table?.range;
  if (!range || table.showTotalsRow === true) return false;
  if (row !== range.bottom + 1 || col < range.left || col > range.right) return false;
  return tableExpansionRangeIsAvailable(sheet, table, { ...range, bottom: row });
}



function tableCanExpandRightForEdit(sheet, table, row, col) {
  const range = table?.range;
  if (!range || col !== range.right + 1 || row < range.top || row > range.bottom) return false;
  return tableExpansionRangeIsAvailable(sheet, table, { ...range, right: col });
}



function tableExpansionRangeIsAvailable(sheet, table, nextRange) {
  const normalizedNext = { sheetIndex: table.range?.sheetIndex ?? state.activeSheetIndex, ...nextRange };
  return !(sheet?.tables || []).some((other) => other !== table && other?.range && rangesIntersect(normalizedNext, other.range));
}



function expandTableDownForEdit(sheet, table, row, editedCol) {
  const range = table.range;
  const sourceRow = tableLastDataRow(table);
  if (!sourceRow || sourceRow < range.top || sourceRow > range.bottom) return false;
  ensureSheetSize(sheet, row, range.right);
  setTableRange(table, { ...range, bottom: row });
  ensureTableColumnMetadata(sheet, table);
  for (let col = table.range.left; col <= table.range.right; col += 1) {
    copyTableCellFormatting(sheet, sourceRow, col, row, col);
  }
  for (let col = table.range.left; col <= table.range.right; col += 1) {
    if (col === editedCol) continue;
    copyTableCalculatedColumnFormula(sheet, table, sourceRow, row, col);
  }
  internCellsInRange(sheet, { ...table.range, top: row, bottom: row });
  return true;
}



function expandTableRightForEdit(sheet, table, row, col) {
  const range = table.range;
  const sourceCol = range.right;
  ensureSheetSize(sheet, range.bottom, col);
  setTableRange(table, { ...range, right: col });
  for (let targetRow = table.range.top; targetRow <= table.range.bottom; targetRow += 1) {
    copyTableCellFormatting(sheet, targetRow, sourceCol, targetRow, col);
  }
  ensureTableColumnMetadata(sheet, table);
  syncTableHeaderNameForEdit(sheet, table.range.top, col);
  internCellsInRange(sheet, { ...table.range, left: col, right: col });
  return true;
}



function expandTableToRangeForEdit(sheet, table, editedRange, nextRange, oldRange = table.range) {
  if (!table?.range || !nextRange) return false;
  const previousRange = { ...oldRange };
  ensureSheetSize(sheet, nextRange.bottom, nextRange.right);
  setTableRange(table, nextRange);
  if (nextRange.right > previousRange.right) {
    for (let col = previousRange.right + 1; col <= nextRange.right; col += 1) {
      for (let row = previousRange.top; row <= previousRange.bottom; row += 1) {
        copyTableCellFormatting(sheet, row, previousRange.right, row, col);
      }
    }
  }
  if (nextRange.bottom > previousRange.bottom) {
    const sourceRow = tableLastDataRow({ ...table, range: previousRange }) || previousRange.bottom;
    for (let row = previousRange.bottom + 1; row <= nextRange.bottom; row += 1) {
      for (let col = nextRange.left; col <= nextRange.right; col += 1) {
        copyTableCellFormatting(sheet, sourceRow, Math.min(col, previousRange.right), row, col);
      }
    }
  }
  ensureTableColumnMetadata(sheet, table);
  syncTableColumnNamesFromHeaders(sheet, table);
  if (nextRange.bottom > previousRange.bottom) {
    const sourceRow = tableLastDataRow({ ...table, range: previousRange }) || previousRange.bottom;
    for (let row = previousRange.bottom + 1; row <= nextRange.bottom; row += 1) {
      for (let col = nextRange.left; col <= nextRange.right; col += 1) {
        if (isCellInRange(editedRange, row, col) && isOccupiedCell(sheet.cells[cellKey(row, col)])) continue;
        copyTableCalculatedColumnFormula(sheet, table, sourceRow, row, col);
      }
    }
  }
  internCellsInRange(sheet, nextRange);
  return true;
}



function setTableRange(table, range) {
  table.range = { sheetIndex: table.range?.sheetIndex ?? state.activeSheetIndex, ...range };
  table.ref = rangeToLabel(table.range);
}



function tableLastDataRow(table) {
  const rows = structuredReferenceRows(table, "data");
  return rows?.bottom || null;
}



function copyTableCellFormatting(sheet, sourceRow, sourceCol, targetRow, targetCol) {
  const source = sheet.cells[cellKey(sourceRow, sourceCol)];
  const target = ensureCellModel(sheet, targetRow, targetCol);
  target.css = mutableCssObject(cloneJsonLikeValue(source?.css) || {});
  target.hasFill = Boolean(source?.hasFill);
  target.hasFontColor = Boolean(source?.hasFontColor);
  target.borders = mutableBordersObject(cloneJsonLikeValue(source?.borders) || {});
  target.numFmt = source?.numFmt || "";
  target.validation = cloneJsonLikeValue(source?.validation);
  storeCellModelForSheet(sheet, cellKey(targetRow, targetCol), target);
}



function copyTableCalculatedColumnFormula(sheet, table, sourceRow, targetRow, col) {
  const formula = calculatedTableColumnFormula(sheet, table, sourceRow, col);
  if (!formula) return false;
  const targetCell = sheet.cells[cellKey(targetRow, col)];
  if (isOccupiedCell(targetCell)) return false;
  const shifted = formula.includes("[") ? formula : shiftFormulaReferences(formula, targetRow - sourceRow, 0);
  const cell = ensureCellModel(sheet, targetRow, col);
  cell.raw = shifted;
  cell.formula = shifted;
  cell.cached = null;
  cell.display = "";
  delete cell.richText;
  delete cell.phonetics;
  delete cell.showPhonetic;
  delete cell.sparkline;
  storeCellModelForSheet(sheet, cellKey(targetRow, col), cell);
  return true;
}



function calculatedTableColumnFormula(sheet, table, sourceRow, col) {
  const source = sheet.cells[cellKey(sourceRow, col)];
  const sourceFormula = source?.formula || (typeof source?.raw === "string" && source.raw.trim().startsWith("=") ? source.raw.trim() : "");
  if (sourceFormula) return sourceFormula;
  const offset = col - table.range.left;
  return normalizeTableFormula(table.columns?.[offset]?.calculatedColumnFormula || "");
}



function tableCalculatedColumnCandidateForEdit(sheet, row, col, content) {
  const text = String(content || "").trim();
  if (!text.startsWith("=")) return null;
  return tableDataColumnForCell(sheet, row, col);
}



function tableDataColumnForCell(sheet, row, col) {
  return (sheet?.tables || [])
    .map((table) => {
      const dataRows = structuredReferenceRows(table, "data");
      if (!table?.range || !dataRows) return null;
      if (row < dataRows.top || row > dataRows.bottom || col < table.range.left || col > table.range.right) return null;
      return { table, row, col };
    })
    .find(Boolean) || null;
}



function applyTableCalculatedColumnForEdit(sheet, row, col, content) {
  const candidate = tableCalculatedColumnCandidateForEdit(sheet, row, col, content);
  if (!candidate) return { changed: false };
  const formula = normalizeTableFormula(content);
  return applyTableCalculatedColumnFormula(sheet, candidate.table, row, col, formula);
}



function applyTableCalculatedColumnsForRangeEdit(sheet, range) {
  if (!sheet?.tables?.length || !range) return [];
  const changedNames = [];
  (sheet.tables || []).forEach((table) => {
    if (!table?.range) return;
    const dataRows = structuredReferenceRows(table, "data");
    if (!dataRows) return;
    const top = Math.max(range.top, dataRows.top);
    const bottom = Math.min(range.bottom, dataRows.bottom);
    const left = Math.max(range.left, table.range.left);
    const right = Math.min(range.right, table.range.right);
    if (top > bottom || left > right) return;
    for (let col = left; col <= right; col += 1) {
      let sourceRow = null;
      let formula = "";
      for (let row = top; row <= bottom; row += 1) {
        const cell = sheet.cells[cellKey(row, col)];
        const sourceFormula = cell?.formula || (typeof cell?.raw === "string" && cell.raw.trim().startsWith("=") ? cell.raw : "");
        const candidate = normalizeTableFormula(sourceFormula);
        if (!candidate || !candidate.startsWith("=")) continue;
        sourceRow = row;
        formula = candidate;
        break;
      }
      if (!formula) continue;
      const result = applyTableCalculatedColumnFormula(sheet, table, sourceRow, col, formula);
      if (result.changed) changedNames.push(result.tableName);
    }
  });
  return changedNames;
}



function applyTableCalculatedColumnFormula(sheet, table, sourceRow, col, formula) {
  const normalized = normalizeTableFormula(formula);
  const dataRows = structuredReferenceRows(table, "data");
  if (!normalized || !dataRows) return { changed: false };
  const offset = col - table.range.left;
  const columns = ensureTableColumnMetadata(sheet, table);
  let changed = false;
  const storedFormula = normalized.startsWith("=") ? normalized.slice(1) : normalized;
  if (columns[offset] && columns[offset].calculatedColumnFormula !== storedFormula) {
    columns[offset].calculatedColumnFormula = storedFormula;
    changed = true;
  }
  for (let row = dataRows.top; row <= dataRows.bottom; row += 1) {
    const shifted = normalized.includes("[") ? normalized : shiftFormulaReferences(normalized, row - sourceRow, 0);
    const key = cellKey(row, col);
    const existing = sheet.cells[key];
    if (existing?.formula === shifted && existing.raw === shifted) continue;
    const cell = ensureCellModel(sheet, row, col);
    cell.raw = shifted;
    cell.formula = shifted;
    cell.cached = null;
    cell.display = "";
    delete cell.richText;
    delete cell.phonetics;
    delete cell.showPhonetic;
    delete cell.sparkline;
    storeCellModelForSheet(sheet, key, cell);
    changed = true;
  }
  if (!changed) return { changed: false };
  const tableName = table.displayName || table.name || "Table";
  const columnNameText = columns[offset]?.name || columnName(col);
  return {
    changed: true,
    tableName,
    label: `${tableName} 計算列`,
    status: `${tableName}[${columnNameText}] の計算列を更新しました。`,
  };
}



function normalizeTableFormula(formula) {
  const text = String(formula || "").trim();
  if (!text) return "";
  return text.startsWith("=") ? text : `=${text}`;
}



function ensureTableColumnMetadata(sheet, table) {
  const range = table?.range;
  if (!range) return [];
  const count = range.right - range.left + 1;
  const columns = Array.isArray(table.columns) ? table.columns.map((column) => ({ ...column })) : [];
  for (let index = 0; index < count; index += 1) {
    const existing = columns[index] || {};
    const fallback = `Column${index + 1}`;
    const preferred = existing.name || tableHeaderText(sheet, table, index) || fallback;
    columns[index] = {
      ...existing,
      id: existing.id || String(index + 1),
      name: preferred,
    };
  }
  table.columns = columns.slice(0, count);
  return table.columns;
}



function syncTableHeaderNameForEdit(sheet, row, col) {
  const table = tableHeaderForCell(sheet, row, col);
  if (!table) return { changed: false };
  const offset = col - table.range.left;
  const columns = ensureTableColumnMetadata(sheet, table);
  const previousName = columns[offset]?.name || "";
  const preferred = tableHeaderText(sheet, table, offset) || `Column${offset + 1}`;
  const nextName = uniqueTableColumnName(columns, preferred, offset);
  let changed = false;
  if (nextName !== previousName) {
    columns[offset].name = nextName;
    renameStructuredReferencesForTableColumn(sheet, table, previousName, nextName);
    changed = true;
  }
  const headerCell = sheet.cells[cellKey(row, col)];
  if (!isOccupiedCell(headerCell)) {
    updateCellModel(sheet, row, col, nextName);
    changed = true;
  }
  if (!changed) return { changed: false };
  const tableName = table.displayName || table.name || "Table";
  return {
    changed: true,
    label: `${tableName} 見出し`,
    status: `${tableName} の見出しを更新しました。`,
  };
}



function syncTableColumnNamesFromHeaders(sheet, table) {
  if (!table?.range || table.showHeaderRow === false) return false;
  let changed = false;
  const columns = ensureTableColumnMetadata(sheet, table);
  for (let col = table.range.left; col <= table.range.right; col += 1) {
    const offset = col - table.range.left;
    const previousName = columns[offset]?.name || "";
    const preferred = tableHeaderText(sheet, table, offset) || `Column${offset + 1}`;
    const nextName = uniqueTableColumnName(columns, preferred, offset);
    if (nextName !== previousName) {
      columns[offset].name = nextName;
      renameStructuredReferencesForTableColumn(sheet, table, previousName, nextName);
      changed = true;
    }
    const headerCell = sheet.cells[cellKey(table.range.top, col)];
    if (!isOccupiedCell(headerCell)) {
      updateCellModel(sheet, table.range.top, col, nextName);
      changed = true;
    }
  }
  return changed;
}



function tableHeaderForCell(sheet, row, col) {
  return (sheet?.tables || []).find((table) => {
    const range = table?.range;
    return range && table.showHeaderRow !== false && row === range.top && col >= range.left && col <= range.right;
  }) || null;
}



function tableHeaderText(sheet, table, offset) {
  if (!table?.range || table.showHeaderRow === false) return "";
  const cell = sheet.cells[cellKey(table.range.top, table.range.left + offset)];
  return String(cell?.raw ?? cell?.display ?? "").trim();
}



function uniqueTableColumnName(columns, preferred, ignoreIndex = -1) {
  const base = String(preferred || "").trim() || `Column${ignoreIndex + 1 || 1}`;
  const used = new Set(
    (columns || [])
      .map((column, index) => (index === ignoreIndex ? "" : String(column?.name || "").trim().toLocaleLowerCase("ja-JP")))
      .filter(Boolean),
  );
  if (!used.has(base.toLocaleLowerCase("ja-JP"))) return base;
  let index = 2;
  let candidate = `${base}${index}`;
  while (used.has(candidate.toLocaleLowerCase("ja-JP"))) {
    index += 1;
    candidate = `${base}${index}`;
  }
  return candidate;
}



function renameStructuredReferencesForTableColumn(sheet, table, previousName, nextName) {
  const oldName = String(previousName || "").trim();
  const newName = String(nextName || "").trim();
  if (!oldName || !newName || oldName === newName || !state.model?.sheets?.length) return false;
  const owningSheetIndex = state.model.sheets.indexOf(sheet);
  const tableSheetIndex = owningSheetIndex >= 0 ? owningSheetIndex : table.range?.sheetIndex;
  let changed = false;
  state.model.sheets.forEach((targetSheet, sheetIndex) => {
    Object.values(targetSheet.cells || {}).forEach((cell) => {
      const formula = typeof cell?.formula === "string" && cell.formula
        ? cell.formula
        : (typeof cell?.raw === "string" && cell.raw.trim().startsWith("=") ? cell.raw : "");
      if (!formula) return;
      const contextTable = sheetIndex === tableSheetIndex && isCellInRange(table.range, cell.row, cell.col);
      const result = rewriteStructuredReferencesForTableColumn(formula, table, oldName, newName, contextTable);
      if (!result.changed) return;
      cell.formula = result.formula;
      cell.raw = result.formula;
      cell.cached = null;
      cell.display = "";
      storeCellModelForSheet(targetSheet, cellKey(cell.row, cell.col), cell);
      changed = true;
    });
    (targetSheet.tables || []).forEach((targetTable) => {
      const contextTable = targetTable === table;
      (targetTable.columns || []).forEach((column) => {
        const formula = normalizeTableFormula(column?.calculatedColumnFormula || "");
        if (!formula) return;
        const result = rewriteStructuredReferencesForTableColumn(formula, table, oldName, newName, contextTable);
        if (!result.changed) return;
        column.calculatedColumnFormula = result.formula.startsWith("=") ? result.formula.slice(1) : result.formula;
        changed = true;
      });
    });
  });
  return changed;
}



function rewriteStructuredReferencesForTableColumn(formula, table, oldName, newName, contextTable = false) {
  const text = String(formula || "");
  const tableNames = [table.displayName, table.name]
    .filter(Boolean)
    .map(String)
    .sort((left, right) => right.length - left.length);
  if (!text || !tableNames.length) return { formula: text, changed: false };
  let output = "";
  let index = 0;
  let changed = false;
  while (index < text.length) {
    if (text[index] === '"') {
      const literal = formulaStringLiteralAt(text, index);
      output += literal.text;
      index = literal.end;
      continue;
    }
    let replaced = false;
    for (const tableName of tableNames) {
      if (!structuredReferenceNameMatches(text, index, tableName)) continue;
      const token = structuredReferenceTokenAt(text, index + tableName.length);
      if (!token) continue;
      const rewritten = rewriteStructuredReferenceTokenColumn(token.text, oldName, newName);
      output += tableName + rewritten.text;
      index = token.end;
      changed = changed || rewritten.changed;
      replaced = true;
      break;
    }
    if (replaced) continue;
    if (contextTable && text[index] === "[") {
      const token = structuredReferenceTokenAt(text, index);
      if (token) {
        const rewritten = rewriteStructuredReferenceTokenColumn(token.text, oldName, newName);
        output += rewritten.text;
        index = token.end;
        changed = changed || rewritten.changed;
        continue;
      }
    }
    output += text[index];
    index += 1;
  }
  return { formula: changed ? output : text, changed };
}



function formulaStringLiteralAt(text, start) {
  let index = start + 1;
  while (index < text.length) {
    if (text[index] === '"') {
      if (text[index + 1] === '"') {
        index += 2;
        continue;
      }
      index += 1;
      break;
    }
    index += 1;
  }
  return { text: text.slice(start, index), end: index };
}



function rewriteStructuredReferenceTokenColumn(token, oldName, newName) {
  const inner = String(token || "").slice(1, -1);
  let changed = false;
  const parts = splitStructuredReferenceParts(inner).map((part) => {
    const rewritten = rewriteStructuredReferencePartColumn(part, oldName, newName);
    changed = changed || rewritten.changed;
    return rewritten.text;
  });
  return { text: changed ? `[${parts.join(",")}]` : token, changed };
}



function rewriteStructuredReferencePartColumn(part, oldName, newName) {
  const raw = String(part || "");
  const leading = raw.match(/^\s*/)?.[0] || "";
  const trailing = raw.match(/\s*$/)?.[0] || "";
  const trimmed = raw.trim();
  const range = structuredReferenceColumnRangePart(trimmed);
  if (range) {
    const left = rewriteStructuredReferenceColumnName(range.start, oldName, newName);
    const right = rewriteStructuredReferenceColumnName(range.end, oldName, newName);
    if (!left.changed && !right.changed) return { text: raw, changed: false };
    return {
      text: `${leading}[${left.name}]:[${right.name}]${trailing}`,
      changed: true,
    };
  }
  const item = cleanStructuredReferencePart(trimmed);
  if (!item || item.replace(/^'/, "").startsWith("#")) return { text: raw, changed: false };
  const currentRow = item.startsWith("@");
  const columnNameText = currentRow ? cleanStructuredReferencePart(item.slice(1)) : item;
  const rewritten = rewriteStructuredReferenceColumnName(columnNameText, oldName, newName);
  if (!rewritten.changed) return { text: raw, changed: false };
  return {
    text: `${leading}${currentRow ? `@[${rewritten.name}]` : `[${rewritten.name}]`}${trailing}`,
    changed: true,
  };
}



function rewriteStructuredReferenceColumnName(name, oldName, newName) {
  const normalized = unescapeStructuredReferencePart(name).trim().toLocaleLowerCase("ja-JP");
  if (normalized !== String(oldName || "").trim().toLocaleLowerCase("ja-JP")) {
    return { name, changed: false };
  }
  return { name: escapeStructuredReferenceColumnName(newName), changed: true };
}



function escapeStructuredReferenceColumnName(name) {
  return String(name || "").replace(/(['\[\]#@])/g, "'$1");
}



function syncTableAutoFilterRanges(sheet, table, oldRange) {
  const nextFilterRange = tableAutoFilterRange(table);
  const previousFilterRange = tableAutoFilterRange({ ...table, range: oldRange });
  if (!nextFilterRange) return;
  table.autoFilterState = table.autoFilterState || { range: null, criteria: {} };
  table.autoFilterState.range = { ...nextFilterRange };
  table.autoFilterState.criteria = table.autoFilterState.criteria || {};
  if (table.showFilterButton !== false && !sheet.autoFilter) {
    sheet.autoFilter = { range: { ...nextFilterRange }, criteria: {}, hiddenButtons: {} };
  } else if (sheet.autoFilter?.range && (rangeRectEquals(sheet.autoFilter.range, previousFilterRange) || rangeRectEquals(sheet.autoFilter.range, oldRange))) {
    sheet.autoFilter.range = { ...sheet.autoFilter.range, ...nextFilterRange };
  }
  if (sheet.sortState?.range && (rangeRectEquals(sheet.sortState.range, previousFilterRange) || rangeRectEquals(sheet.sortState.range, oldRange))) {
    sheet.sortState.range = { ...sheet.sortState.range, ...nextFilterRange };
  }
}



function tableAutoFilterRange(table) {
  const range = table?.range;
  if (!range) return null;
  const bottom = table.showTotalsRow === true ? Math.max(range.top, range.bottom - 1) : range.bottom;
  return { sheetIndex: range.sheetIndex ?? state.activeSheetIndex, top: range.top, left: range.left, bottom, right: range.right };
}



function rangeRectEquals(a, b) {
  return Boolean(a && b && a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right);
}



function materializeImportedAutoFilterCriteria(sheet) {
  const criteria = sheet?.autoFilter?.criteria;
  if (!criteria || typeof criteria !== "object") return;
  Object.entries(criteria).forEach(([colText, criterion]) => {
    const col = Number(colText);
    if (!col || !criterion || typeof criterion !== "object" || Array.isArray(criterion)) return;
    if (criterion.top10Source) criteria[col] = materializeTop10AutoFilterCriterion(sheet, col, criterion.top10Source) || criterion;
    else if (criterion.dynamicSource) criteria[col] = materializeDynamicAutoFilterCriterion(sheet, col, criterion.dynamicSource) || criterion;
  });
}



function materializeTop10AutoFilterCriterion(sheet, col, source = {}) {
  const range = sheet?.autoFilter?.range;
  if (!range || !col) return null;
  const numericRows = importedAutoFilterNumericRows(sheet, col);
  if (!numericRows.length) return null;
  const direction = source.direction === "bottom" ? "bottom" : "top";
  const mode = source.mode === "percent" ? "percent" : "items";
  const count = Math.max(1, Math.min(500, Math.floor(Number(source.count) || 10)));
  const take = mode === "percent" ? Math.max(1, Math.ceil((numericRows.length * Math.min(100, count)) / 100)) : Math.min(numericRows.length, count);
  const sorted = [...numericRows].sort((left, right) => (direction === "top" ? right.value - left.value : left.value - right.value));
  const threshold = sorted[Math.max(0, take - 1)]?.value;
  if (!Number.isFinite(threshold)) return null;
  return {
    kind: "custom",
    operator1: direction === "top" ? "greaterThanOrEqual" : "lessThanOrEqual",
    value1: String(threshold),
    logic: "and",
    operator2: "equal",
    value2: "",
    top10: { direction, mode, count },
  };
}



function materializeDynamicAutoFilterCriterion(sheet, col, source = {}) {
  const type = source.type;
  const averageFilter = type === "aboveAverage" || type === "belowAverage";
  const stdDevFilter = type === "aboveStdDev" || type === "belowStdDev";
  if (!averageFilter && !stdDevFilter) return null;
  const values = importedAutoFilterNumericRows(sheet, col).map((item) => item.value);
  if (!values.length) return null;
  const average = values.reduce((total, value) => total + value, 0) / values.length;
  if (stdDevFilter) {
    const count = autoFilterStdDevCount(source.val);
    const variance = values.length > 1
      ? values.reduce((total, value) => total + ((value - average) ** 2), 0) / (values.length - 1)
      : 0;
    const threshold = average + (type === "aboveStdDev" ? 1 : -1) * Math.sqrt(variance) * count;
    return {
      kind: "custom",
      operator1: type === "aboveStdDev" ? "greaterThan" : "lessThan",
      value1: String(threshold),
      logic: "and",
      operator2: "equal",
      value2: "",
      dynamicSource: { type, val: count },
      stdDevOperator: type,
      stdDevCount: count,
    };
  }
  return {
    kind: "custom",
    operator1: type === "aboveAverage" ? "greaterThan" : "lessThan",
    value1: String(average),
    logic: "and",
    operator2: "equal",
    value2: "",
    averageOperator: type,
  };
}



function autoFilterStdDevCount(value) {
  const count = Math.round(Number(value));
  return Math.max(1, Math.min(3, Number.isFinite(count) ? count : 1));
}



function importedAutoFilterNumericRows(sheet, col) {
  const range = sheet?.autoFilter?.range;
  if (!range || !col) return [];
  const rows = [];
  for (let row = range.top + 1; row <= range.bottom; row += 1) {
    const value = Number(importedAutoFilterCellValue(sheet, row, col).replace(/,/g, ""));
    if (Number.isFinite(value)) rows.push({ row, value });
  }
  return rows;
}



function importedAutoFilterCellValue(sheet, row, col) {
  return String(cellDisplayText(sheet?.cells?.[cellKey(row, col)]) ?? "");
}



function excelAutoFilterRangeToModel(autoFilter) {
  if (!autoFilter) return null;
  if (typeof autoFilter === "string") {
    const decoded = decodeRange(autoFilter.replace(/\$/g, ""));
    return decoded ? { sheetIndex: 0, ...decoded } : null;
  }
  const from = excelAutoFilterPoint(autoFilter.from);
  const to = excelAutoFilterPoint(autoFilter.to || autoFilter.from);
  if (!from || !to) return null;
  return {
    sheetIndex: 0,
    top: Math.min(from.row, to.row),
    left: Math.min(from.col, to.col),
    bottom: Math.max(from.row, to.row),
    right: Math.max(from.col, to.col),
  };
}



function excelAutoFilterPoint(point) {
  if (!point) return null;
  if (typeof point === "string") return decodeCell(point.replace(/\$/g, ""));
  const row = Number(point.row);
  const col = Number(point.col ?? point.column);
  if (!Number.isFinite(row) || !Number.isFinite(col)) return null;
  return { row, col };
}



function workbookMediaToDataUrls(workbook) {
  return (workbook.model?.media || []).map((item) => {
    if (item?.type !== "image" || !item.buffer) return null;
    const extension = String(item.extension || "png").toLowerCase();
    const mimeType = item.contentType || imageMimeType(extension);
    return {
      name: item.name || "",
      extension,
      mimeType,
      src: `data:${mimeType};base64,${bytesToBase64(item.buffer)}`,
    };
  });
}



function worksheetImagesToModel(worksheet, media, displayScale = 1) {
  return (worksheet.getImages?.() || [])
    .map((image, index) => {
      const source = media[image.imageId];
      if (!source?.src || !image.range?.tl) return null;
      return {
        id: `image-${image.imageId}-${index}`,
        name: source.name || `image${index + 1}`,
        src: source.src,
        mimeType: source.mimeType,
        extension: source.extension,
        editAs: image.range.editAs || "",
        tl: anchorToModel(image.range.tl),
        br: image.range.br ? anchorToModel(image.range.br) : null,
        ext: image.range.ext
          ? {
              width: Math.max(1, Math.round((Number(image.range.ext.width) || 1) * displayScale)),
              height: Math.max(1, Math.round((Number(image.range.ext.height) || 1) * displayScale)),
            }
          : null,
        hyperlink: image.range.hyperlinks?.hyperlink || "",
      };
    })
    .filter(Boolean);
}



function mergeWorksheetImagesAndDrawingObjects(worksheetImages = [], drawingObjects = []) {
  const images = [...worksheetImages];
  (drawingObjects || []).forEach((object) => {
    if (!object || worksheetImageDuplicatesDrawingObject(images, object)) return;
    images.push(object);
  });
  return images;
}



function worksheetImageDuplicatesDrawingObject(images, object) {
  if (object.shape) return false;
  return images.some((image) => !image.shape && sameWorksheetImagePlacement(image, object));
}



function sameWorksheetImagePlacement(a, b) {
  if (!a?.src || a.src !== b?.src || !sameAnchorPoint(a.tl, b.tl)) return false;
  if (a.ext && b.ext) {
    return Math.abs(Number(a.ext.width || 0) - Number(b.ext.width || 0)) <= 2 &&
      Math.abs(Number(a.ext.height || 0) - Number(b.ext.height || 0)) <= 2;
  }
  if (a.br || b.br) return sameAnchorPoint(a.br, b.br);
  return true;
}



function sameAnchorPoint(a, b) {
  if (!a || !b) return false;
  const aCol = Number(a.nativeCol ?? a.col ?? 0);
  const bCol = Number(b.nativeCol ?? b.col ?? 0);
  const aRow = Number(a.nativeRow ?? a.row ?? 0);
  const bRow = Number(b.nativeRow ?? b.row ?? 0);
  return Math.abs(aCol - bCol) < 0.01 &&
    Math.abs(aRow - bRow) < 0.01 &&
    Math.abs(Number(a.nativeColOff || 0) - Number(b.nativeColOff || 0)) <= 20000 &&
    Math.abs(Number(a.nativeRowOff || 0) - Number(b.nativeRowOff || 0)) <= 20000;
}



function imageMimeType(extension) {
  const normalized = String(extension || "").toLowerCase();
  if (normalized === "jpg" || normalized === "jpeg") return "image/jpeg";
  if (normalized === "gif") return "image/gif";
  if (normalized === "webp") return "image/webp";
  if (normalized === "bmp") return "image/bmp";
  if (normalized === "svg") return "image/svg+xml";
  if (normalized === "ico" || normalized === "icon") return "image/x-icon";
  return "image/png";
}



function bytesToBase64(bufferLike) {
  const bytes = bufferLike instanceof Uint8Array ? bufferLike : new Uint8Array(bufferLike);
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return btoa(binary);
}



function base64ToBytes(value) {
  const binary = atob(String(value || ""));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}



function revokeWorksheetImageRenderObjectUrls() {
  if (typeof URL !== "undefined" && typeof URL.revokeObjectURL === "function") {
    worksheetImageRenderObjectUrls.forEach((url) => URL.revokeObjectURL(url));
  }
  worksheetImageRenderObjectUrls.clear();
}



function cachedWorksheetImageRenderSource(src) {
  const value = String(src || "");
  if (
    value.length < WORKSHEET_IMAGE_OBJECT_URL_MIN_DATA_URL_LENGTH ||
    !/^data:image\//i.test(value) ||
    typeof Blob === "undefined" ||
    typeof URL === "undefined" ||
    typeof URL.createObjectURL !== "function"
  ) {
    return value;
  }
  const cached = worksheetImageRenderObjectUrls.get(value);
  if (cached) return cached;
  const blob = imageDataUrlToBlob(value);
  if (!blob) return value;
  const objectUrl = URL.createObjectURL(blob);
  worksheetImageRenderObjectUrls.set(value, objectUrl);
  return objectUrl;
}



function imageDataUrlToBlob(src) {
  const match = String(src || "").match(/^data:([^,]+),([\s\S]*)$/i);
  if (!match) return null;
  const meta = match[1] || "image/png";
  const data = match[2] || "";
  const mimeType = meta.split(";")[0] || "image/png";
  try {
    const bytes = /;base64(?:;|$)/i.test(`;${meta}`) ? base64ToBytes(data) : new TextEncoder().encode(decodeURIComponent(data));
    return new Blob([bytes], { type: mimeType });
  } catch (error) {
    console.warn("Failed to create worksheet image object URL.", error);
    return null;
  }
}



function anchorToModel(anchor) {
  return {
    col: Number(anchor.col ?? anchor.nativeCol ?? 0),
    row: Number(anchor.row ?? anchor.nativeRow ?? 0),
    nativeCol: Number(anchor.nativeCol ?? 0),
    nativeRow: Number(anchor.nativeRow ?? 0),
    nativeColOff: Number(anchor.nativeColOff ?? 0),
    nativeRowOff: Number(anchor.nativeRowOff ?? 0),
  };
}



function defaultWorkbookWidthMetrics() {
  return { maxDigitWidth: 7 };
}



function inferWorkbookWidthMetrics(workbook) {
  let sampledFonts = 0;
  let wideDigitFonts = 0;
  const addFontName = (font) => {
    const name = String(font?.name || "").trim();
    if (!name) return;
    sampledFonts += 1;
    if (isWideExcelDigitFont(name)) wideDigitFonts += 1;
  };

  for (const worksheet of workbook.worksheets || []) {
    let sampled = 0;
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      if (sampled >= 500) return;
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (sampled >= 500) return;
        addFontName(cell.font);
        sampled += 1;
      });
    });
  }

  return { maxDigitWidth: sampledFonts > 0 && wideDigitFonts / sampledFonts >= 0.5 ? 8 : 7 };
}



function isWideExcelDigitFont(name) {
  const normalized = normalizeFontName(name);
  return /ms\s*p?gothic|ｍｓ\s*[pｐ]?ゴシック|ms\s*p?ゴシック|ゴシック|明朝|hiragino|yu\s*gothic|meiryo|noto\s*sans\s*cjk|source\s*han|ipaex|takao/i.test(
    normalized,
  );
}



function worksheetZoomScale(worksheet) {
  const view = worksheet.views?.[0] || {};
  const zoom = Number(view.zoomScale ?? view.zoomScaleNormal ?? 100);
  return Number.isFinite(zoom) && zoom > 0 ? zoom : 100;
}



function splitPaneFromWorksheetView(view, metrics = {}) {
  if (!view || view.state !== "split") return null;
  const columns = splitPaneIndexFromExcelSplitUnits(view.xSplit, metrics.columns || [], metrics.defaultColumnWidthPixels || columnWidthToPixels(8.43));
  const rows = splitPaneIndexFromExcelSplitUnits(view.ySplit, metrics.rows || [], metrics.defaultRowHeightPixels || rowHeightToPixels(15));
  if (!columns && !rows) return null;
  return normalizePaneState({
    rows,
    columns,
    row: rows + 1,
    col: columns + 1,
    activePane: view.activePane || "",
  });
}



function splitPaneIndexFromExcelSplitUnits(value, items = [], defaultSize = 0) {
  const pixels = excelSplitUnitsToPixels(value);
  if (pixels <= 0) return 0;
  let position = 0;
  const fallback = Math.max(1, Number(defaultSize) || 1);
  const count = Math.max(items.length, 1);
  for (let index = 0; index < count; index += 1) {
    const item = items[index] || {};
    const size = item.hidden ? 0 : Number(item.width ?? item.height ?? fallback) || fallback;
    if (size <= 0) continue;
    const next = position + size;
    if (pixels < (position + next) / 2) return index;
    if (pixels <= next) return index + 1;
    position = next;
  }
  return Math.max(1, count);
}



function excelSplitUnitsToPixels(value) {
  const units = Number(value) || 0;
  if (units <= 0) return 0;
  return units > 300 ? units / 15 : units * (96 / 72);
}



function cellToModel(cell, row, col, themeColors, workbookProperties = {}, sheetName = "") {
  if (isCoveredMergedCell(cell)) {
    return {
      row,
      col,
      raw: null,
      formula: null,
      cached: null,
      display: "",
      hyperlink: null,
      comment: null,
      note: null,
      noteVisible: false,
      validation: null,
      sparkline: null,
      richText: null,
      css: {},
      hasFill: false,
      hasFontColor: false,
      borders: {},
      numFmt: "",
    };
  }

  const valueInfo = normalizeCellValue(cell, workbookProperties, row, col, sheetName);
  const css = styleToCss(cell, themeColors);
  const hasFill = Boolean(css.backgroundColor);
  const hasFontColor = explicitCellFontColor(cell.font, themeColors);
  const richText = normalizeCellRichText(cell.value?.richText, themeColors);
  const protection = cellProtectionToModel(cell.protection);
  if (hasHardLineBreak(valueInfo.raw) || hasHardLineBreak(valueInfo.display)) {
    css.whiteSpace = "pre-wrap";
    css.overflowWrap = "anywhere";
  }
  return {
    row,
    col,
    raw: valueInfo.raw,
    formula: valueInfo.formula,
    cached: valueInfo.cached,
    display: valueInfo.display,
    hyperlink: valueInfo.hyperlink,
    comment: null,
    note: normalizeNote(cell.note),
    noteVisible: false,
    validation: normalizeValidation(cell.dataValidation),
    css,
    richText,
    hasFill,
    hasFontColor,
    borders: bordersToModel(cell.border, themeColors),
    numFmt: cell.numFmt || "",
    protection,
  };
}



function cellProtectionToModel(protection) {
  if (!protection) return null;
  const locked = protection.locked !== false;
  const hidden = Boolean(protection.hidden);
  return locked && !hidden ? null : { locked, hidden };
}



function shouldStoreCell(cell) {
  return Boolean(
    cell.raw != null ||
      cell.formula ||
      cell.display ||
      cell.hyperlink ||
      cell.comment ||
      cell.note ||
      cell.phonetics?.length ||
      cell.showPhonetic ||
      cell.validation ||
      cell.sparkline ||
      cell.quickAnalysisBar ||
      cell.quickAnalysisIcon ||
      cell.quickAnalysisStyle ||
      cell.richText?.length ||
      Object.keys(cell.css || {}).length ||
      cell.hasFill ||
      cell.hasFontColor ||
      Object.keys(cell.borders || {}).length ||
      cell.numFmt ||
      cell.protection,
  );
}



function normalizeCellValue(cell, workbookProperties = {}, row = 0, col = 0, sheetName = "") {
  const value = cell.value;
  const text = safeCellText(cell);
  const base = {
    raw: null,
    formula: null,
    cached: null,
    display: text,
    hyperlink: cell.hyperlink || null,
  };

  if (value == null) {
    return base;
  }

  if (value instanceof Date) {
    const serial = excelDateObjectToSerial(value);
    return {
      ...base,
      raw: serial,
      cached: serial,
      display: displayValue(serial, cell.numFmt),
    };
  }

  if (typeof value !== "object") {
    const raw = normalizeRawValue(value, cell.numFmt, workbookProperties);
    return {
      ...base,
      raw,
      cached: raw,
      display: text || displayValue(raw, cell.numFmt) || String(value),
    };
  }

  if (value.formula || value.sharedFormula || cell.formula) {
    const formula = String(value.formula || cell.formula || value.sharedFormula || "");
    const rawFormula = formula.startsWith("=") ? formula : `=${formula}`;
    const normalizedFormula = normalizeCurrentCellReferenceFormula(rawFormula, row, col, sheetName);
    const cached = normalizeRawValue(value.result ?? cell.result ?? null, cell.numFmt, workbookProperties);
    return {
      ...base,
      raw: normalizedFormula,
      formula: normalizedFormula,
      cached,
      display: text || displayValue(cached, cell.numFmt),
    };
  }

  if (value.richText) {
    const text = value.richText.map((part) => part.text || "").join("");
    return {
      ...base,
      raw: text,
      cached: text,
      display: safeCellText(cell) || text,
    };
  }

  if (value.text || value.hyperlink) {
    const text = String(value.text || value.hyperlink || "");
    return {
      ...base,
      raw: text,
      cached: text,
      display: safeCellText(cell) || text,
      hyperlink: value.hyperlink || base.hyperlink,
    };
  }

  if (value.error) {
    return {
      ...base,
      raw: value.error,
      cached: value.error,
      display: value.error,
    };
  }

  const fallbackText = safeCellText(cell) || String(value.result ?? "");
  return {
    ...base,
    raw: fallbackText,
    cached: fallbackText,
    display: fallbackText,
  };
}



function normalizeCurrentCellReferenceFormula(formula, row = 0, col = 0, sheetName = "") {
  const currentRef = `${columnName(Number(col) || 1)}${Number(row) || 1}`.toUpperCase();
  const quotedSheetName = sheetName ? `'${String(sheetName).replace(/'/g, "''")}'` : "";
  const sheetAliases = new Set([
    "",
    String(sheetName || "").toUpperCase(),
    quotedSheetName.toUpperCase(),
  ]);
  return String(formula || "").replace(/\b(ROW|COLUMN)\s*\(\s*((?:'(?:(?:''|[^'])*)'|[A-Za-z_][\w .]*)!)?(\$?[A-Z]{1,3}\$?\d+)\s*\)/gi, (match, fnName, sheetPrefix = "", ref = "") => {
    const normalizedRef = String(ref || "").replace(/\$/g, "").toUpperCase();
    if (normalizedRef !== currentRef) return match;
    const normalizedSheet = String(sheetPrefix || "").replace(/!$/, "").toUpperCase();
    if (!sheetAliases.has(normalizedSheet)) return match;
    return `${String(fnName || "").toUpperCase()}()`;
  });
}



function normalizeCellRichText(richText, themeColors = DEFAULT_THEME_COLORS) {
  if (!Array.isArray(richText) || !richText.length) return null;
  const runs = richText
    .map((part) => {
      const text = String(part?.text || "");
      if (!text) return null;
      const css = {};
      const font = part?.font || {};
      if (font.bold) css.fontWeight = "700";
      if (font.italic) css.fontStyle = "italic";
      if (font.underline) css.textDecoration = String(font.underline).toLowerCase().includes("double") ? "underline double" : "underline";
      if (font.size) css.fontSize = `${font.size}pt`;
      if (font.name) css.fontFamily = fontFamilyForExcel(font.name);
      const color = colorToCss(font.color, themeColors);
      if (color) css.color = color;
      return Object.keys(css).length ? { text, css } : { text };
    })
    .filter(Boolean);
  return runs.length ? runs : null;
}



function isCoveredMergedCell(cell) {
  return Boolean(cell?.isMerged && cell.master && cell.master.address !== cell.address);
}



function safeCellText(cell) {
  try {
    return cell.text || "";
  } catch {
    return "";
  }
}



function hasHardLineBreak(value) {
  return typeof value === "string" && /\r|\n/.test(value);
}



function normalizeRawValue(value, numFmt = "", workbookProperties = {}) {
  if (value == null) return null;
  if (value instanceof Date) return excelDateObjectToSerial(value);
  if (typeof value === "number" && workbookUses1904DateSystem(workbookProperties) && isDateFormat(numFmt)) {
    return excelSerial1904To1900(value);
  }
  if (typeof value === "object") {
    if (value.error) return value.error;
    if (value.text) return value.text;
    if (value.result !== undefined) return normalizeRawValue(value.result, numFmt, workbookProperties);
    return String(value);
  }
  return value;
}



function normalizeNote(note) {
  if (!note) return null;
  if (typeof note === "string") return note;
  if (Array.isArray(note.texts)) {
    return note.texts.map((part) => part.text || "").join("");
  }
  return String(note);
}



function normalizeValidation(validation) {
  if (!validation) return null;
  if (validation.type === "list") {
    const formula = String(validation.formulae?.[0] ?? "");
    return {
      type: "list",
      formula,
      options: parseValidationOptions(formula),
      allowBlank: Boolean(validation.allowBlank),
      showDropDown: validation.showDropDown !== false,
      showInputMessage: Boolean(validation.showInputMessage),
      inputTitle: validation.promptTitle || validation.inputTitle || "",
      inputMessage: validation.prompt || validation.inputMessage || "",
      showErrorMessage: validation.showErrorMessage !== false,
      errorStyle: validation.errorStyle || "stop",
      errorTitle: validation.errorTitle || "",
      errorMessage: validation.error || validation.errorMessage || "",
    };
  }
  return {
    type: validation.type || "custom",
    operator: validation.operator || "",
    formulae: validation.formulae || [],
    allowBlank: Boolean(validation.allowBlank),
    showInputMessage: Boolean(validation.showInputMessage),
    inputTitle: validation.promptTitle || validation.inputTitle || "",
    inputMessage: validation.prompt || validation.inputMessage || "",
    showErrorMessage: validation.showErrorMessage !== false,
    errorStyle: validation.errorStyle || "stop",
    errorTitle: validation.errorTitle || "",
    errorMessage: validation.error || validation.errorMessage || "",
  };
}



function parseValidationOptions(formula) {
  const trimmed = String(formula ?? "").trim();
  if (!trimmed) return [];
  const startsFormula = trimmed.startsWith("=");
  const literal = startsFormula ? trimmed.slice(1).trim() : trimmed;
  const quotedLiteral =
    (literal.startsWith('"') && literal.endsWith('"')) ||
    (literal.startsWith("'") && literal.endsWith("'"));
  if (startsFormula && !quotedLiteral) return [];
  const unquoted =
    quotedLiteral ? literal.slice(1, -1) : literal;

  if (/[!:]/.test(unquoted)) {
    return [];
  }

  return splitValidationListLiteral(unquoted)
    .map((option) => option.trim())
    .filter(Boolean);
}



function splitValidationListLiteral(value) {
  const text = String(value ?? "");
  const separator = text.includes(",") || !text.includes(";") ? "," : ";";
  const parts = [];
  let current = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === '"') {
      if (quoted && text[index + 1] === '"') {
        current += '"';
        index += 1;
      } else {
        quoted = !quoted;
      }
      continue;
    }
    if (char === separator && !quoted) {
      parts.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  parts.push(current);
  return parts;
}



function validationListOptions(sheet, rule, visitedNames = new Set()) {
  if (!rule || rule.type !== "list") return [];
  const direct = (rule.options || []).map((item) => String(item)).filter(Boolean);
  if (direct.length) return direct;
  const inline = parseValidationOptions(rule.formula || "");
  if (inline.length) return inline;
  return validationOptionsFromReference(sheet, rule.formula || "", visitedNames);
}



function validationOptionsFromReference(baseSheet, formula, visitedNames = new Set()) {
  const reference = validationReferenceFromFormula(baseSheet, formula, visitedNames);
  if (!reference?.sheet || !reference.range) return [];
  const sheetId = getSheetId(reference.sheet.name);
  const values = [];
  for (let row = reference.range.top; row <= reference.range.bottom; row += 1) {
    for (let col = reference.range.left; col <= reference.range.right; col += 1) {
      const value = getDisplayForCell(reference.sheet, sheetId, row, col);
      if (String(value ?? "") !== "") values.push(String(value));
    }
  }
  return values;
}



function validationReferenceFromFormula(baseSheet, formula, visitedNames = new Set()) {
  const text = String(formula || "").trim().replace(/^=/, "").trim();
  if (!text) return null;
  const direct = validationQualifiedRange(text, baseSheet);
  if (direct) return direct;
  const rangeExpression = validationRangeExpressionReference(baseSheet, text, visitedNames);
  if (rangeExpression) return rangeExpression;
  const calculated = validationFunctionReferenceFromFormula(baseSheet, text, visitedNames);
  if (calculated) return calculated;
  const definition = validationDefinedNameFromFormulaText(text, baseSheet);
  if (!definition?.refersTo) return null;
  const visitKey = validationDefinedNameVisitKey(definition);
  if (visitedNames.has(visitKey)) return null;
  visitedNames.add(visitKey);
  const firstRef = firstDefinedNameReference(definition.refersTo);
  const definitionSheet = definition.scope === "sheet"
    ? state.model?.sheets?.[definition.sheetIndex] || baseSheet
    : baseSheet;
  return validationReferenceFromFormula(definitionSheet, firstRef, visitedNames);
}



function validationQualifiedRange(text, baseSheet) {
  const bangIndex = text.lastIndexOf("!");
  let sheet = baseSheet;
  let rangeText = text;
  if (bangIndex >= 0) {
    const sheetName = unquoteSheetName(text.slice(0, bangIndex));
    sheet = (state.model?.sheets || []).find((item) => item.name === sheetName) || null;
    rangeText = text.slice(bangIndex + 1);
  }
  const range = decodeExcelRangeReference(rangeText.replace(/\$/g, ""), sheet);
  return sheet && range ? { sheet, range } : null;
}



function decodeExcelRangeReference(rangeText, sheet = activeSheet()) {
  const text = String(rangeText || "").trim();
  const cellRange = decodeRange(text);
  if (cellRange) return cellRange;
  const columnMatch = text.match(/^([A-Z]+):([A-Z]+)$/i);
  if (columnMatch) {
    const left = columnNumber(columnMatch[1]);
    const right = columnNumber(columnMatch[2]);
    return {
      top: 1,
      left: Math.min(left, right),
      bottom: validationReferenceMaxRow(sheet),
      right: Math.max(left, right),
    };
  }
  const rowMatch = text.match(/^([1-9][0-9]*):([1-9][0-9]*)$/);
  if (rowMatch) {
    const top = Number(rowMatch[1]);
    const bottom = Number(rowMatch[2]);
    return {
      top: Math.min(top, bottom),
      left: 1,
      bottom: Math.max(top, bottom),
      right: validationReferenceMaxCol(sheet),
    };
  }
  return null;
}



function validationReferenceMaxRow(sheet) {
  return Math.max(1, Number(sheet?.rowCount || sheet?.rows?.length || 1));
}



function validationReferenceMaxCol(sheet) {
  return Math.max(1, Number(sheet?.colCount || sheet?.columns?.length || 1));
}



function validationFunctionReferenceFromFormula(baseSheet, text, visitedNames = new Set()) {
  const call = parseFormulaFunctionCall(text);
  if (!call) return null;
  if (call.name === "OFFSET") return validationOffsetReference(baseSheet, call.args, visitedNames);
  if (call.name === "INDEX") return validationIndexReference(baseSheet, call.args, visitedNames);
  if (call.name === "INDIRECT") return validationIndirectReference(baseSheet, call.args, visitedNames);
  return null;
}



function firstDefinedNameReference(refersTo) {
  return splitFormulaArguments(String(refersTo || "").replace(/^=/, ""))[0]?.trim() || "";
}



function validationRangeExpressionReference(baseSheet, text, visitedNames = new Set()) {
  const parts = splitTopLevelFormulaOperator(text, ":");
  if (parts.length !== 2) return null;
  const start = validationReferenceEndpoint(baseSheet, parts[0], "start", visitedNames);
  const end = validationReferenceEndpoint(start?.sheet || baseSheet, parts[1], "end", visitedNames);
  if (!start?.sheet || !end?.sheet || start.sheet !== end.sheet) return null;
  return {
    sheet: start.sheet,
    range: {
      top: Math.min(start.row, end.row),
      left: Math.min(start.col, end.col),
      bottom: Math.max(start.row, end.row),
      right: Math.max(start.col, end.col),
    },
  };
}



function validationReferenceEndpoint(baseSheet, formula, edge = "start", visitedNames = new Set()) {
  const text = String(formula || "").trim().replace(/^=/, "").trim();
  if (!text) return null;
  const direct = validationQualifiedRange(text, baseSheet);
  if (direct?.sheet && direct.range) return validationEndpointFromReference(direct, edge);
  const calculated = validationFunctionReferenceFromFormula(baseSheet, text, visitedNames);
  if (calculated?.sheet && calculated.range) return validationEndpointFromReference(calculated, edge);
  const definition = validationDefinedNameFromFormulaText(text, baseSheet);
  if (!definition?.refersTo) return null;
  const visitKey = validationDefinedNameVisitKey(definition);
  if (visitedNames.has(visitKey)) return null;
  visitedNames.add(visitKey);
  const definitionSheet = definition.scope === "sheet"
    ? state.model?.sheets?.[definition.sheetIndex] || baseSheet
    : baseSheet;
  const reference = validationReferenceFromFormula(definitionSheet, firstDefinedNameReference(definition.refersTo), visitedNames);
  return reference?.sheet && reference.range ? validationEndpointFromReference(reference, edge) : null;
}



function validationEndpointFromReference(reference, edge = "start") {
  const range = reference?.range;
  if (!reference?.sheet || !range) return null;
  return {
    sheet: reference.sheet,
    row: edge === "end" ? range.bottom : range.top,
    col: edge === "end" ? range.right : range.left,
  };
}



function validationOffsetReference(baseSheet, args, visitedNames = new Set()) {
  if (!args?.length || args.length < 3) return null;
  const baseReference = validationReferenceFromFormula(baseSheet, args[0], visitedNames);
  if (!baseReference?.sheet || !baseReference.range) return null;
  const rowOffset = validationFormulaNumberValue(baseReference.sheet, args[1], visitedNames);
  const colOffset = validationFormulaNumberValue(baseReference.sheet, args[2], visitedNames);
  if (!Number.isFinite(rowOffset) || !Number.isFinite(colOffset)) return null;
  const sourceHeight = baseReference.range.bottom - baseReference.range.top + 1;
  const sourceWidth = baseReference.range.right - baseReference.range.left + 1;
  const height = args.length >= 4 ? validationFormulaNumberValue(baseReference.sheet, args[3], visitedNames) : sourceHeight;
  const width = args.length >= 5 ? validationFormulaNumberValue(baseReference.sheet, args[4], visitedNames) : sourceWidth;
  if (!Number.isFinite(height) || !Number.isFinite(width) || height <= 0 || width <= 0) return null;
  const top = Math.max(1, baseReference.range.top + Math.trunc(rowOffset));
  const left = Math.max(1, baseReference.range.left + Math.trunc(colOffset));
  const bottom = Math.max(top, top + Math.trunc(height) - 1);
  const right = Math.max(left, left + Math.trunc(width) - 1);
  return { sheet: baseReference.sheet, range: { top, left, bottom, right } };
}



function validationIndexReference(baseSheet, args, visitedNames = new Set()) {
  if (!args?.length || args.length < 2) return null;
  const baseReference = validationReferenceFromFormula(baseSheet, args[0], visitedNames);
  if (!baseReference?.sheet || !baseReference.range) return null;
  const rowText = String(args[1] || "").trim();
  const colText = String(args[2] || "").trim();
  const rowOffset = rowText === "" ? 0 : validationFormulaNumberValue(baseReference.sheet, rowText, visitedNames);
  const colOffset = args.length >= 3
    ? (colText === "" ? 0 : validationFormulaNumberValue(baseReference.sheet, colText, visitedNames))
    : 1;
  if (!Number.isFinite(rowOffset) || !Number.isFinite(colOffset)) return null;
  const rowIndex = Math.trunc(rowOffset);
  const colIndex = Math.trunc(colOffset);
  if (rowIndex < 0 || colIndex < 0 || (rowIndex === 0 && colIndex === 0)) return null;
  if (rowIndex === 0) {
    const col = baseReference.range.left + colIndex - 1;
    if (col > baseReference.range.right) return null;
    return {
      sheet: baseReference.sheet,
      range: { top: baseReference.range.top, left: col, bottom: baseReference.range.bottom, right: col },
    };
  }
  if (colIndex === 0) {
    const row = baseReference.range.top + rowIndex - 1;
    if (row > baseReference.range.bottom) return null;
    return {
      sheet: baseReference.sheet,
      range: { top: row, left: baseReference.range.left, bottom: row, right: baseReference.range.right },
    };
  }
  const row = baseReference.range.top + rowIndex - 1;
  const col = baseReference.range.left + colIndex - 1;
  if (row > baseReference.range.bottom || col > baseReference.range.right) return null;
  return { sheet: baseReference.sheet, range: { top: row, left: col, bottom: row, right: col } };
}



function validationIndirectReference(baseSheet, args, visitedNames = new Set()) {
  if (!args?.length) return null;
  const referenceText = validationFormulaTextValue(baseSheet, args[0], visitedNames);
  if (!referenceText) return null;
  const a1Mode = args.length >= 2 ? validationFormulaBooleanValue(baseSheet, args[1], visitedNames, true) : true;
  if (a1Mode === false) {
    const r1c1Reference = validationQualifiedR1C1Range(referenceText, baseSheet);
    if (r1c1Reference) return r1c1Reference;
  }
  return validationReferenceFromFormula(baseSheet, referenceText, visitedNames);
}



function validationFormulaNumberValue(baseSheet, formula, visitedNames = new Set()) {
  const text = String(formula || "").trim().replace(/^=/, "").trim();
  if (!text) return Number.NaN;
  if (/^[+-]?\d+(?:\.\d+)?$/.test(text)) return Number(text);
  const call = parseFormulaFunctionCall(text);
  if (!call) {
    const reference = validationReferenceFromFormula(baseSheet, text, visitedNames);
    if (reference?.sheet && reference.range?.top === reference.range.bottom && reference.range.left === reference.range.right) {
      return validationFirstCellNumber(reference);
    }
  }
  if (!call) return Number.NaN;
  if (call.name === "COUNTA" || call.name === "COUNT") {
    return call.args.reduce((total, arg) => {
      const reference = validationReferenceFromFormula(baseSheet, arg, visitedNames);
      if (!reference) return total;
      return total + validationCountA(reference, call.name === "COUNT");
    }, 0);
  }
  if (call.name === "ROWS" || call.name === "COLUMNS") {
    const reference = validationReferenceFromFormula(baseSheet, call.args[0] || "", visitedNames);
    if (!reference?.range) return Number.NaN;
    return call.name === "ROWS"
      ? reference.range.bottom - reference.range.top + 1
      : reference.range.right - reference.range.left + 1;
  }
  if (call.name === "MATCH") return validationMatchValue(baseSheet, call.args, visitedNames);
  return Number.NaN;
}



function validationFormulaBooleanValue(baseSheet, formula, visitedNames = new Set(), fallback = true) {
  const text = String(formula || "").trim().replace(/^=/, "").trim();
  if (!text) return fallback;
  if (/^TRUE$/iu.test(text)) return true;
  if (/^FALSE$/iu.test(text)) return false;
  if (/^[+-]?\d+(?:\.\d+)?$/.test(text)) return Number(text) !== 0;
  if (!parseFormulaFunctionCall(text)) {
    const reference = validationReferenceFromFormula(baseSheet, text, visitedNames);
    if (reference?.sheet && reference.range?.top === reference.range.bottom && reference.range.left === reference.range.right) {
      const value = validationFirstCellText(reference);
      if (/^TRUE$/iu.test(value)) return true;
      if (/^FALSE$/iu.test(value)) return false;
      const number = Number(String(value ?? "").replace(/,/g, ""));
      if (Number.isFinite(number)) return number !== 0;
    }
  }
  return fallback;
}



function validationFirstCellNumber(reference) {
  const text = validationFirstCellText(reference);
  const number = Number(String(text ?? "").replace(/,/g, ""));
  return Number.isFinite(number) ? number : Number.NaN;
}



function validationQualifiedR1C1Range(text, baseSheet) {
  const bangIndex = String(text || "").lastIndexOf("!");
  let sheet = baseSheet;
  let rangeText = String(text || "").trim();
  if (bangIndex >= 0) {
    const sheetName = unquoteSheetName(rangeText.slice(0, bangIndex));
    sheet = (state.model?.sheets || []).find((item) => item.name === sheetName) || null;
    rangeText = rangeText.slice(bangIndex + 1);
  }
  const range = decodeExcelR1C1RangeReference(rangeText.replace(/\$/g, ""));
  return sheet && range ? { sheet, range } : null;
}



function decodeExcelR1C1RangeReference(rangeText) {
  const parts = String(rangeText || "").trim().split(":");
  if (!parts.length || parts.length > 2) return null;
  const start = decodeExcelR1C1CellReference(parts[0]);
  const end = decodeExcelR1C1CellReference(parts[1] || parts[0]);
  if (!start || !end) return null;
  return {
    top: Math.min(start.row, end.row),
    left: Math.min(start.col, end.col),
    bottom: Math.max(start.row, end.row),
    right: Math.max(start.col, end.col),
  };
}



function decodeExcelR1C1CellReference(text) {
  const match = String(text || "").trim().match(/^R([1-9][0-9]*)C([1-9][0-9]*)$/i);
  if (!match) return null;
  return { row: Number(match[1]), col: Number(match[2]) };
}



function validationCountA(reference, numbersOnly = false) {
  const sheet = reference?.sheet;
  const range = reference?.range;
  if (!sheet || !range) return 0;
  const sheetId = getSheetId(sheet.name);
  let count = 0;
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      const cell = sheet.cells[cellKey(row, col)];
      const raw = getCellRawInput(sheet, row, col);
      const display = getDisplayForCell(sheet, sheetId, row, col);
      const hasValue = cell?.formula || cell?.dataTableFormula || raw !== "" || String(display ?? "") !== "";
      if (!hasValue) continue;
      if (numbersOnly) {
        const value = Number(String(display || raw).replace(/,/g, ""));
        if (!Number.isFinite(value)) continue;
      }
      count += 1;
    }
  }
  return count;
}



function validationFormulaTextValue(baseSheet, formula, visitedNames = new Set()) {
  const text = String(formula || "").trim().replace(/^=/, "").trim();
  if (!text) return "";
  const concatParts = splitTopLevelFormulaOperator(text, "&");
  if (concatParts.length > 1) {
    const values = concatParts.map((part) => validationFormulaTextValue(baseSheet, part, visitedNames));
    return values.some((value) => value == null) ? "" : values.join("");
  }
  const literal = unquoteFormulaStringLiteral(text);
  if (literal != null) return literal;
  const reference = validationReferenceFromFormula(baseSheet, text, visitedNames);
  if (reference?.sheet && reference.range) return validationFirstCellText(reference);
  return text;
}



function validationFirstCellText(reference) {
  const sheet = reference?.sheet;
  const range = reference?.range;
  if (!sheet || !range) return "";
  const sheetId = getSheetId(sheet.name);
  const raw = getCellRawInput(sheet, range.top, range.left);
  const display = getDisplayForCell(sheet, sheetId, range.top, range.left);
  const preferred = raw && !String(raw).startsWith("=") ? raw : display;
  return String(preferred ?? "").trim();
}



function validationMatchValue(baseSheet, args, visitedNames = new Set()) {
  if (!args?.length || args.length < 2) return Number.NaN;
  const lookup = normalizeValidationMatchText(validationFormulaTextValue(baseSheet, args[0], visitedNames));
  const reference = validationReferenceFromFormula(baseSheet, args[1], visitedNames);
  if (!lookup || !reference?.sheet || !reference.range) return Number.NaN;
  const matchTypeText = String(args[2] || "").trim();
  const matchType = matchTypeText === "" ? 1 : validationFormulaNumberValue(baseSheet, matchTypeText, visitedNames);
  if (Number.isFinite(matchType) && Math.trunc(matchType) !== 0) return Number.NaN;
  const sheetId = getSheetId(reference.sheet.name);
  for (let row = reference.range.top; row <= reference.range.bottom; row += 1) {
    for (let col = reference.range.left; col <= reference.range.right; col += 1) {
      const display = getDisplayForCell(reference.sheet, sheetId, row, col);
      const raw = getCellRawInput(reference.sheet, row, col);
      const value = normalizeValidationMatchText(display || raw);
      if (value !== lookup) continue;
      return reference.range.bottom === reference.range.top
        ? col - reference.range.left + 1
        : row - reference.range.top + 1;
    }
  }
  return Number.NaN;
}



function normalizeValidationMatchText(value) {
  return String(value ?? "").trim().toLocaleLowerCase("ja-JP");
}



function unquoteFormulaStringLiteral(value) {
  const text = String(value || "").trim();
  if (text.length < 2 || !text.startsWith('"') || !text.endsWith('"')) return null;
  return text.slice(1, -1).replace(/""/g, '"');
}



function splitTopLevelFormulaOperator(source, operator) {
  const parts = [];
  let current = "";
  let depth = 0;
  let singleQuoted = false;
  let doubleQuoted = false;
  const text = String(source || "");
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (doubleQuoted) {
      current += char;
      if (char === '"' && next === '"') {
        current += next;
        index += 1;
      } else if (char === '"') {
        doubleQuoted = false;
      }
      continue;
    }
    if (singleQuoted) {
      current += char;
      if (char === "'" && next === "'") {
        current += next;
        index += 1;
      } else if (char === "'") {
        singleQuoted = false;
      }
      continue;
    }
    if (char === '"') {
      doubleQuoted = true;
      current += char;
      continue;
    }
    if (char === "'") {
      singleQuoted = true;
      current += char;
      continue;
    }
    if (char === "(") {
      depth += 1;
      current += char;
      continue;
    }
    if (char === ")") {
      depth = Math.max(0, depth - 1);
      current += char;
      continue;
    }
    if (char === operator && depth === 0) {
      parts.push(current.trim());
      current = "";
      continue;
    }
    current += char;
  }
  parts.push(current.trim());
  return parts;
}



function parseFormulaFunctionCall(value) {
  const text = String(value || "").trim();
  const nameMatch = text.match(/^([A-Za-z_][A-Za-z0-9_.]*)\s*\(/);
  if (!nameMatch || !text.endsWith(")")) return null;
  const openIndex = text.indexOf("(", nameMatch[0].length - 1);
  const argsSource = text.slice(openIndex + 1, -1);
  return {
    name: nameMatch[1].toUpperCase(),
    args: splitFormulaArguments(argsSource),
  };
}



function splitFormulaTopLevel(source, separators) {
  const parts = [];
  let current = "";
  let depth = 0;
  let singleQuoted = false;
  let doubleQuoted = false;
  const text = String(source || "");
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (doubleQuoted) {
      current += char;
      if (char === '"' && next === '"') {
        current += next;
        index += 1;
      } else if (char === '"') {
        doubleQuoted = false;
      }
      continue;
    }
    if (singleQuoted) {
      current += char;
      if (char === "'" && next === "'") {
        current += next;
        index += 1;
      } else if (char === "'") {
        singleQuoted = false;
      }
      continue;
    }
    if (char === '"') {
      doubleQuoted = true;
      current += char;
      continue;
    }
    if (char === "'") {
      singleQuoted = true;
      current += char;
      continue;
    }
    if ("([{".includes(char)) {
      depth += 1;
      current += char;
      continue;
    }
    if (")]}".includes(char)) {
      depth = Math.max(0, depth - 1);
      current += char;
      continue;
    }
    if (separators.has(char) && depth === 0) {
      parts.push(current.trim());
      current = "";
      continue;
    }
    current += char;
  }
  parts.push(current.trim());
  return parts;
}



function splitFormulaArguments(source) {
  return splitFormulaTopLevel(source, new Set([",", ";"]));
}



function validationDefinedNameFromFormulaText(text, baseSheet) {
  const qualified = validationQualifiedDefinedName(text);
  if (qualified) {
    const sheet = (state.model?.sheets || []).find((item) => item.name === qualified.sheetName) || null;
    return sheet ? validationDefinedName(qualified.name, sheet) : null;
  }
  const name = unquoteFormulaDefinedName(text);
  return name ? validationDefinedName(name, baseSheet) : null;
}



function validationQualifiedDefinedName(text) {
  const source = String(text || "").trim();
  const bangIndex = source.lastIndexOf("!");
  if (bangIndex < 0) return null;
  const sheetName = unquoteSheetName(source.slice(0, bangIndex));
  const name = unquoteFormulaDefinedName(source.slice(bangIndex + 1));
  if (!sheetName || !name || decodeRange(name.replace(/\$/g, ""))) return null;
  return { sheetName, name };
}



function unquoteFormulaDefinedName(value) {
  const text = String(value || "").trim();
  if (text.startsWith("'") && text.endsWith("'")) return text.slice(1, -1).replace(/''/g, "'").trim();
  return text;
}

export {
  anchorToModel,
  applyImportedDataValidationDetails,
  applyImportedTableCalculatedColumns,
  applyImportedTableFillIfEmpty,
  applyImportedTableTotals,
  applyImportedTableVisualStyle,
  applyImportedTableVisualStyles,
  applyTableAutoExpansionForEdit,
  applyTableCalculatedColumnForEdit,
  applyTableCalculatedColumnFormula,
  applyTableCalculatedColumnsForRangeEdit,
  applyTableEditBehaviors,
  applyTableRangeEditBehaviors,
  autoFilterStdDevCount,
  base64ToBytes,
  bytesToBase64,
  cachedWorksheetImageRenderSource,
  calculatedTableColumnFormula,
  cellProtectionToModel,
  cellToModel,
  clampOutlineLevel,
  copyTableCalculatedColumnFormula,
  copyTableCellFormatting,
  customSheetViewHeaderFooterFromNode,
  customSheetViewNodeToModel,
  customSheetViewPageLayoutFromNode,
  customSheetViewPaneFromNode,
  customSheetViewSelectionFromNode,
  customViewScrollFromSheetView,
  decodeExcelR1C1CellReference,
  decodeExcelR1C1RangeReference,
  decodeExcelRangeReference,
  decodeSheetViewSqrefRange,
  defaultWorkbookWidthMetrics,
  ensureTableColumnMetadata,
  escapeStructuredReferenceColumnName,
  excelAutoFilterPoint,
  excelAutoFilterRangeToModel,
  excelFitToPageValue,
  excelHeaderFooterSection,
  excelHeaderFooterToModel,
  excelManualBreakIndexes,
  excelPageCommentsValue,
  excelPageErrorsValue,
  excelPaperSizeValue,
  excelSplitUnitsToPixels,
  expandTableDownForEdit,
  expandTableRightForEdit,
  expandTableToRangeForEdit,
  firstDefinedNameReference,
  firstXmlChild,
  firstXmlChildText,
  formulaStringLiteralAt,
  hasHardLineBreak,
  imageDataUrlToBlob,
  imageMimeType,
  importedAutoFilterCellValue,
  importedAutoFilterNumericRows,
  importedDataValidationDimensions,
  importedTableDimensions,
  importedTableStylePalette,
  inferWorkbookWidthMetrics,
  inlineStringPhonetics,
  isCoveredMergedCell,
  isWideExcelDigitFont,
  materializeDynamicAutoFilterCriterion,
  materializeImportedAutoFilterCriteria,
  materializeImportedTableCalculatedCell,
  materializeTop10AutoFilterCriterion,
  mergeMasterForPoint,
  mergeMissingTableBorders,
  mergeWorksheetImagesAndDrawingObjects,
  normalizeCellRichText,
  normalizeCellValue,
  normalizeCurrentCellReferenceFormula,
  normalizeImportedWorksheetTables,
  normalizeNote,
  normalizeOutlineSettings,
  normalizePackageRelationshipPath,
  normalizePhoneticRuns,
  normalizeRawValue,
  normalizeTableFormula,
  normalizeValidation,
  normalizeValidationMatchText,
  normalizeWorkbookRelationshipPath,
  pageLayoutMarginPresetForMargins,
  parseFormulaFunctionCall,
  parseValidationOptions,
  parseWorksheetPhonetics,
  parseXmlDocument,
  rangeHasEditableTableContent,
  rangeIntersectsWorksheetTable,
  rangeRectEquals,
  refreshImportedTableVisualStyleForTableMutation,
  renameStructuredReferencesForTableColumn,
  restoreScrollRecordForView,
  revokeWorksheetImageRenderObjectUrls,
  rewriteStructuredReferenceColumnName,
  rewriteStructuredReferencePartColumn,
  rewriteStructuredReferenceTokenColumn,
  rewriteStructuredReferencesForTableColumn,
  safeCellText,
  sameAnchorPoint,
  sameWorksheetImagePlacement,
  scrollRecordFromTopLeftCell,
  setImportedTableTotalCell,
  setTableRange,
  sharedStringForCell,
  shouldStoreCell,
  splitFormulaArguments,
  splitFormulaTopLevel,
  splitPaneFromWorksheetView,
  splitPaneIndexFromExcelSplitUnits,
  splitTopLevelFormulaOperator,
  splitValidationListLiteral,
  syncTableAutoFilterRanges,
  syncTableColumnNamesFromHeaders,
  syncTableHeaderNameForEdit,
  tableAutoFilterRange,
  tableCalculatedColumnCandidateForEdit,
  tableCanExpandDownForEdit,
  tableCanExpandRightForEdit,
  tableDataColumnForCell,
  tableDataRangeIntersectsRange,
  tableEditContentIsNonEmpty,
  tableEditMayChangeMetadata,
  tableExpansionRangeIsAvailable,
  tableExpansionTargetForEdit,
  tableHeaderForCell,
  tableHeaderIntersectsRange,
  tableHeaderText,
  tableLastDataRow,
  tableRangeEditMayChangeMetadata,
  tableRangeExpansionTarget,
  topLeftCellFromScrollRecord,
  totalsRowSubtotalCode,
  uniqueTableColumnName,
  unquoteFormulaDefinedName,
  unquoteFormulaStringLiteral,
  validationCountA,
  validationDefinedNameFromFormulaText,
  validationEndpointFromReference,
  validationFirstCellNumber,
  validationFirstCellText,
  validationFormulaBooleanValue,
  validationFormulaNumberValue,
  validationFormulaTextValue,
  validationFunctionReferenceFromFormula,
  validationIndexReference,
  validationIndirectReference,
  validationListOptions,
  validationMatchValue,
  validationOffsetReference,
  validationOptionsFromReference,
  validationQualifiedDefinedName,
  validationQualifiedR1C1Range,
  validationQualifiedRange,
  validationRangeExpressionReference,
  validationReferenceEndpoint,
  validationReferenceFromFormula,
  validationReferenceMaxCol,
  validationReferenceMaxRow,
  workbookMediaToDataUrls,
  workbookWindowViewFromCustomWorkbookViewNode,
  worksheetImageDuplicatesDrawingObject,
  worksheetImagesToModel,
  worksheetOutlineSettings,
  worksheetPageLayoutToModel,
  worksheetProtectionToModel,
  worksheetToModel,
  worksheetZoomScale,
  xmlBoolAttribute,
  xmlChildren,
  xmlElements,
  xmlLocalName,
  xmlTextContent,
};
