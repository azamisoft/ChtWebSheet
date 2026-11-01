// Feature module; shared bindings are imported explicitly.
import {
  $,
  $formulaInput,
  $horizontalScrollbar,
  $sheetHost,
  $sheetTabs,
  CJK_FONT_FALLBACKS,
  DEFAULT_THEME_COLORS,
  EXCEL_FONT_FALLBACKS,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  FORMULA_ENGINE_STATIC_CELL_INFO_TYPES,
  HYPERFORMULA_LICENSE,
  HyperFormula,
  IMPORTED_RANGE_DIMENSION_EXPANSION_CELL_LIMIT,
  INDEXED_COLORS,
  SHEET_INSERT_TEMPLATES,
  SPECIAL_SHEET_INSERT_META,
  WEBSHEET_FORMULA_ENGINE_CACHEABLE_SPILL_NAMES,
  state,
} from "../../app-core.js";
import {
  closeDataDialog,
  insertChartSheetFromSelection,
  recommendedChartTypeForSelection,
  uniqueSheetNameWithBase,
} from "../drawing/charts-and-insertion.js";
import {
  normalizeColorValue,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  isElementImeCompositionEvent,
  nowMs,
} from "../editing/keyboard-and-paste.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  beginFormulaBarEdit,
  cellPointFromElement,
  getCellRawInput,
  selectionRangeFromPoints,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  clearCellEditSelection,
  normalizeFontName,
  stopEditAutoScroll,
  syncDataRibbonControls,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  internCellModelStylesForSheet,
  pushHistory,
  pushSheetDeleteHistory,
  pushSheetDeleteManyHistory,
  pushSheetInsertHistory,
  pushSheetInsertManyHistory,
  pushSheetMoveHistory,
  pushSheetMoveManyHistory,
} from "../editing/structure-and-history.js";
import {
  formulaCaretIsInsideString,
  formulaTextStartsFormula,
  updateFormulaAutocompleteForEditor,
} from "../formula/autocomplete.js";
import {
  activeFormulaReferenceEditor,
  activeFunctionArgumentReferenceInput,
  closeFormulaAutocomplete,
  formulaEditorCaretOffset,
  formulaEditorSelectionOffsets,
  formulaEditorText,
  formulaReferenceEditorSheetIndex,
  isFunctionArgumentReferenceEditor,
  updateFormulaReferenceUi,
} from "../formula/dialogs-and-names.js";
import {
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  closeCellContextMenu,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  clearWarnings,
  columnName,
  columnWidthToPixels,
  decodeRange,
  escapeAttr,
  escapeHtml,
  excelSerialToDate,
  isDateFormat,
  rowHeightToPixels,
  setStatus,
  showAppAlert,
  showMessageBox,
  showWarnings,
} from "../ui/help-and-utilities.js";
import {
  syncDrawRibbonButtons,
} from "../ui/printing-and-home-tools.js";
import {
  isCellReference,
  quoteSheetNameForFormula,
  resetLazyScrollLocks,
} from "../ui/ribbon-and-links.js";
import {
  clearSheetSelectionState,
  cloneSheetModel,
  closeSheetTabContextMenu,
  closeTranslationPane,
  createBlankSheet,
  createSpecialSheet,
  moveSheetToIndex,
  nearestVisibleSheetIndex,
  openSheetTabContextMenuByKeyboard,
  refreshWorkbookAfterSheetMetadataChange,
  uniqueCopiedSheetName,
  uniqueCopiedSheetNameFromSet,
  uniqueSheetName,
  updateSheetFooterUi,
  validateSheetName,
} from "../ui/sheet-protection-and-translation.js";
import {
  renderWorkbookWindowBar,
  syncFullScreenViewClass,
  syncViewRibbonButtons,
  syncWorkbookChromeVisibility,
  syncWorkbookTitle,
} from "../ui/views-and-formula-tools.js";
import {
  scheduleWorkbookPersistence,
} from "./import-model.js";
import {
  defaultWorkbookWidthMetrics,
  parseFormulaFunctionCall,
  splitFormulaArguments,
} from "./tables-and-validation.js";
import {
  syncPageLayoutRibbonControls,
} from "./vba-and-printing.js";
import {
  restoreWorksheetFocusAfterCommandPopup,
  saveActiveWindowViewState,
} from "./views-and-macros.js";


function unquoteSheetName(value) {
  const text = String(value || "").trim();
  if (text.startsWith("'") && text.endsWith("'")) return text.slice(1, -1).replace(/''/g, "'");
  return text;
}



function validationDefinedName(name, baseSheet, options = {}) {
  const normalized = normalizeDefinedNameLookup(name);
  const sheetIndex = (state.model?.sheets || []).indexOf(baseSheet);
  const definitions = state.model?.definedNames || [];
  const local = definitions.find((definition) =>
    !definition.hidden &&
    definition.scope === "sheet" &&
    definition.sheetIndex === sheetIndex &&
    normalizeDefinedNameLookup(definition.name) === normalized,
  );
  if (local || options.localOnly) return local || null;
  return definitions.find((definition) =>
    !definition.hidden &&
    definition.scope !== "sheet" &&
    normalizeDefinedNameLookup(definition.name) === normalized,
  ) || null;
}



function normalizeDefinedNameLookup(name) {
  return String(name || "").toLocaleLowerCase("ja-JP");
}



function validationDefinedNameVisitKey(definition) {
  return `${definition?.scope || "workbook"}:${definition?.sheetIndex ?? ""}:${normalizeDefinedNameLookup(definition?.name)}`;
}



function styleToCss(cell, themeColors = DEFAULT_THEME_COLORS) {
  const css = {};
  const { font, fill, alignment } = cell;
  const bgColor = fillColorToCss(fill, themeColors);
  if (bgColor) css.backgroundColor = bgColor;

  if (font) {
    if (font.bold) css.fontWeight = "700";
    if (font.italic) css.fontStyle = "italic";
    if (font.underline) css.textDecoration = String(font.underline).toLowerCase().includes("double") ? "underline double" : "underline";
    if (font.size) css.fontSize = `${font.size}pt`;
    if (font.name) css.fontFamily = fontFamilyForExcel(font.name);
    const fontColor = colorToCss(font.color, themeColors);
    if (fontColor) css.color = fontColor;
  }

  if (alignment) {
    if (alignment.horizontal) css.textAlign = alignment.horizontal;
    if (alignment.vertical) {
      css.verticalAlign = alignment.vertical;
      css.alignItems = verticalAlignToFlex(alignment.vertical);
    }
    const textRotation = normalizeCellTextRotation(alignment.textRotation);
    if (textRotation) css.textRotation = textRotation;
    if (alignment.wrapText) {
      css.whiteSpace = "pre-wrap";
      css.overflowWrap = "anywhere";
    }
  }

  return css;
}



function explicitCellFontColor(font, themeColors = DEFAULT_THEME_COLORS) {
  if (!font?.color) return false;
  const color = normalizeColorValue(colorToCss(font.color, themeColors));
  if (!color) return false;
  const defaultBlack = color === "#000000";
  const defaultMinorFont =
    normalizeFontName(font.name || "") === "calibri" &&
    Number(font.size || 11) === 11 &&
    (font.scheme === "minor" || font.scheme == null) &&
    !font.bold &&
    !font.italic &&
    !font.underline &&
    !font.strike;
  return !(defaultBlack && defaultMinorFont);
}



function fillColorToCss(fill, themeColors = DEFAULT_THEME_COLORS) {
  if (!fill) return null;
  if (fill.pattern === "none") return null;
  if (fill.type === "pattern" && fill.pattern && fill.pattern !== "solid") {
    return colorToCss(fill.fgColor, themeColors) || colorToCss(fill.bgColor, themeColors);
  }
  if (fill.fgColor) return colorToCss(fill.fgColor, themeColors);
  if (fill.bgColor) return colorToCss(fill.bgColor, themeColors);
  return null;
}



function colorToCss(color, themeColors = DEFAULT_THEME_COLORS) {
  if (!color) return null;
  if (color.argb) {
    const argb = String(color.argb).replace(/^#/, "");
    const rgb = argb.length === 8 ? argb.slice(2) : argb;
    if (/^0{6}$/i.test(rgb)) return "#000000";
    if (/^[0-9a-f]{6}$/i.test(rgb)) return `#${rgb}`;
  }
  if (typeof color.theme === "number") {
    const base = themeColors[color.theme] || DEFAULT_THEME_COLORS[color.theme];
    return base ? applyTint(`#${base}`, Number(color.tint || 0)) : null;
  }
  if (typeof color.indexed === "number") {
    return INDEXED_COLORS[color.indexed] || null;
  }
  return null;
}



function fontFamilyForExcel(name) {
  const cleaned = String(name || "").replaceAll('"', "").trim();
  if (!cleaned) return "";
  const aliases = EXCEL_FONT_FALLBACKS[cleaned.toLowerCase()] || [cleaned];
  return uniqueFontFamilies([...aliases, ...CJK_FONT_FALLBACKS]).map(quoteFont).join(", ");
}



function uniqueFontFamilies(fonts) {
  const seen = new Set();
  return fonts.filter((font) => {
    const key = String(font).toLowerCase();
    if (!font || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}



function quoteFont(name) {
  const cleaned = String(name).replaceAll('"', "").trim();
  if (/^(serif|sans-serif|monospace|cursive|fantasy|system-ui)$/i.test(cleaned)) return cleaned;
  return `"${cleaned}"`;
}



function verticalAlignToFlex(value) {
  const normalized = String(value || "").toLowerCase();
  if (normalized === "top") return "flex-start";
  if (normalized === "bottom") return "flex-end";
  return "center";
}



function normalizeCellTextRotation(value) {
  if (value === undefined || value === null || value === 0 || value === "0") return "";
  const normalized = String(value).trim().toLowerCase();
  if (!normalized) return "";
  if (normalized === "vertical") return "vertical";
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric === 0) return "";
  return String(Math.max(-90, Math.min(90, numeric)));
}



function bordersToModel(border, themeColors = DEFAULT_THEME_COLORS) {
  if (!border) return {};
  const model = {};
  addBorderModel(model, "top", border.top, themeColors);
  addBorderModel(model, "right", border.right, themeColors);
  addBorderModel(model, "bottom", border.bottom, themeColors);
  addBorderModel(model, "left", border.left, themeColors);
  return model;
}



function addBorderModel(model, side, border, themeColors = DEFAULT_THEME_COLORS) {
  if (!border?.style) return;
  const { width, style } = borderStyleToCss(border.style);
  const color = colorToCss(border.color, themeColors) || "#000000";
  model[side] = { width, style, color, weight: borderWeight(border.style) };
}



function mergedEdgeBorders(worksheet, merge, themeColors = DEFAULT_THEME_COLORS) {
  const borders = {};
  const top = findEdgeBorder(worksheet, merge, "top");
  const right = findEdgeBorder(worksheet, merge, "right");
  const bottom = findEdgeBorder(worksheet, merge, "bottom");
  const left = findEdgeBorder(worksheet, merge, "left");

  addBorderModel(borders, "top", top, themeColors);
  addBorderModel(borders, "right", right, themeColors);
  addBorderModel(borders, "bottom", bottom, themeColors);
  addBorderModel(borders, "left", left, themeColors);
  return borders;
}



function findEdgeBorder(worksheet, merge, side) {
  const candidates = [];
  if (side === "top" || side === "bottom") {
    const row = side === "top" ? merge.top : merge.bottom;
    for (let col = merge.left; col <= merge.right; col += 1) {
      candidates.push(worksheet.getCell(row, col).border?.[side]);
    }
  } else {
    const col = side === "left" ? merge.left : merge.right;
    for (let row = merge.top; row <= merge.bottom; row += 1) {
      candidates.push(worksheet.getCell(row, col).border?.[side]);
    }
  }

  return candidates.filter(Boolean).sort((a, b) => borderWeight(b.style) - borderWeight(a.style))[0];
}



function borderWeight(style) {
  const weights = {
    hair: 1,
    dotted: 1,
    thin: 2,
    dashed: 2,
    dashDot: 2,
    dashDotDot: 2,
    medium: 3,
    mediumDashed: 3,
    mediumDashDot: 3,
    mediumDashDotDot: 3,
    slantDashDot: 3,
    thick: 4,
    double: 4,
  };
  return weights[style] || 0;
}



function parseThemeColors(workbook) {
  const xml = workbook.model?.themes?.theme1 || workbook._themes?.theme1 || "";
  if (!xml) return DEFAULT_THEME_COLORS;

  const order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  return order.map((name, index) => parseThemeColor(xml, name) || DEFAULT_THEME_COLORS[index]);
}



function parseThemeColor(xml, name) {
  const section = xml.match(new RegExp(`<a:${name}>[\\s\\S]*?<\\/a:${name}>`))?.[0];
  if (!section) return null;
  const srgb = section.match(/<a:srgbClr[^>]*val="([0-9A-Fa-f]{6})"/)?.[1];
  if (srgb) return srgb.toUpperCase();

  const system = section.match(/lastClr="([0-9A-Fa-f]{6})"/)?.[1];
  return system ? system.toUpperCase() : null;
}



function applyTint(hex, tint) {
  const rgb = hexToRgb(hex);
  if (!rgb || !tint) return hex.toUpperCase();

  const transformed = rgb.map((channel) => {
    const value = tint < 0 ? channel * (1 + tint) : channel * (1 - tint) + 255 * tint;
    return Math.max(0, Math.min(255, Math.round(value)));
  });

  return rgbToHex(transformed);
}



function hexToRgb(hex) {
  const clean = String(hex).replace("#", "");
  if (!/^[0-9a-f]{6}$/i.test(clean)) return null;
  return [0, 2, 4].map((index) => Number.parseInt(clean.slice(index, index + 2), 16));
}



function rgbToHex(rgb) {
  return `#${rgb.map((value) => value.toString(16).padStart(2, "0")).join("")}`.toUpperCase();
}



function borderStyleToCss(style) {
  const widthMap = {
    hair: 1,
    thin: 1,
    dotted: 1,
    dashDot: 1,
    dashDotDot: 1,
    dashed: 1,
    medium: 2,
    mediumDashed: 2,
    mediumDashDot: 2,
    mediumDashDotDot: 2,
    slantDashDot: 2,
    thick: 3,
    double: 3,
  };
  const styleMap = {
    dotted: "dotted",
    dashDot: "dashed",
    dashDotDot: "dashed",
    dashed: "dashed",
    mediumDashed: "dashed",
    mediumDashDot: "dashed",
    mediumDashDotDot: "dashed",
    slantDashDot: "dashed",
    double: "double",
  };
  return {
    width: widthMap[style] || 1,
    style: styleMap[style] || "solid",
  };
}



function extractMerges(worksheet) {
  const merges = [];
  const seen = new Set();
  const addMerge = (merge) => {
    if (!merge) return;
    let model = merge.model || merge;
    if (typeof model === "string") {
      model = decodeRange(model);
    } else if (model.range) {
      model = decodeRange(model.range);
    }
    if (!model) return;

    const normalized = {
      top: model.top,
      left: model.left,
      bottom: model.bottom,
      right: model.right,
    };
    const id = `${normalized.top}:${normalized.left}:${normalized.bottom}:${normalized.right}`;
    if (
      normalized.top &&
      normalized.left &&
      normalized.bottom &&
      normalized.right &&
      normalized.bottom > normalized.top - 1 &&
      normalized.right > normalized.left - 1 &&
      !seen.has(id)
    ) {
      seen.add(id);
      merges.push(normalized);
    }
  };

  (worksheet.model?.merges || []).forEach(addMerge);

  const privateMerges = worksheet._merges;
  if (privateMerges instanceof Map) {
    [...privateMerges.values()].forEach(addMerge);
  } else if (privateMerges) {
    Object.values(privateMerges).forEach(addMerge);
  }

  return merges;
}



function findWorksheetDimensions(worksheet, merges, widthMetrics = defaultWorkbookWidthMetrics(), drawingObjects = []) {
  const modelDimensions = worksheetExplicitDimensionExtents(worksheet);
  let rows = Math.max(worksheet.actualRowCount || 0, modelDimensions.rows, 1);
  let cols = Math.max(worksheet.actualColumnCount || 0, modelDimensions.cols, 1);
  const defaultColWidth = worksheet.properties?.defaultColWidth || 8.43;
  const defaultColumnWidthPixels = columnWidthToPixels(defaultColWidth, 1, widthMetrics.maxDigitWidth);
  const defaultRowHeight = worksheet.properties?.defaultRowHeight || 15;
  const defaultRowHeightPixels = rowHeightToPixels(defaultRowHeight);

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    rows = Math.max(rows, rowNumber);
    row.eachCell({ includeEmpty: false }, (_cell, colNumber) => {
      cols = Math.max(cols, colNumber);
    });
  });

  Object.keys(worksheet.dataValidations?.model || {}).forEach((ref) => {
    const range = decodeRange(String(ref || "").replace(/\$/g, ""));
    if (!range) return;
    const rangeRows = Math.max(0, (Number(range.bottom) || 0) - (Number(range.top) || 0) + 1);
    const rangeCols = Math.max(0, (Number(range.right) || 0) - (Number(range.left) || 0) + 1);
    if (rangeRows * rangeCols > IMPORTED_RANGE_DIMENSION_EXPANSION_CELL_LIMIT) return;
    rows = Math.max(rows, range.bottom);
    cols = Math.max(cols, range.right);
  });

  merges.forEach((merge) => {
    rows = Math.max(rows, merge.bottom);
    cols = Math.max(cols, merge.right);
  });

  (worksheet.getImages?.() || []).forEach((image) => {
    const tl = image.range?.tl;
    if (!tl) return;
    if (image.range?.br) {
      rows = Math.max(rows, Math.ceil(Number(image.range.br.row ?? image.range.br.nativeRow ?? 0)) + 1);
      cols = Math.max(cols, Math.ceil(Number(image.range.br.col ?? image.range.br.nativeCol ?? 0)) + 1);
      return;
    }
    const row = Number(tl.row ?? tl.nativeRow ?? 0);
    const col = Number(tl.col ?? tl.nativeCol ?? 0);
    const width = Number(image.range?.ext?.width || 0);
    const height = Number(image.range?.ext?.height || 0);
    rows = Math.max(rows, Math.ceil(row + height / defaultRowHeightPixels) + 1);
    cols = Math.max(cols, Math.ceil(col + width / defaultColumnWidthPixels) + 1);
  });

  (drawingObjects || []).forEach((image) => {
    const tl = image.tl;
    if (!tl) return;
    if (image.br) {
      rows = Math.max(rows, Math.ceil(Number(image.br.row ?? image.br.nativeRow ?? 0)) + 1);
      cols = Math.max(cols, Math.ceil(Number(image.br.col ?? image.br.nativeCol ?? 0)) + 1);
      return;
    }
    const row = Number(tl.row ?? tl.nativeRow ?? 0);
    const col = Number(tl.col ?? tl.nativeCol ?? 0);
    const width = Number(image.ext?.width || 0);
    const height = Number(image.ext?.height || 0);
    rows = Math.max(rows, Math.ceil(row + height / defaultRowHeightPixels) + 1);
    cols = Math.max(cols, Math.ceil(col + width / defaultColumnWidthPixels) + 1);
  });

  return { rows, cols };
}



function worksheetExplicitDimensionExtents(worksheet) {
  let rows = 0;
  let cols = 0;
  const usedExtent = worksheetUsedDimensionExtents(worksheet);
  const rowHasExplicitDimension = (row) =>
    row?.hidden === true ||
    Number(row?.outlineLevel) > 0 ||
    row?.collapsed === true ||
    Number.isFinite(Number(row?.height));
  const columnHasExplicitDimension = (column) =>
    column?.hidden === true ||
    Number(column?.outlineLevel) > 0 ||
    column?.collapsed === true ||
    Number.isFinite(Number(column?.width));

  (worksheet?.model?.rows || []).forEach((row, offset) => {
    const index = Math.max(1, Math.round(Number(row?.number ?? row?.r ?? offset + 1) || 0));
    if (index && rowHasExplicitDimension(row)) rows = Math.max(rows, index);
  });
  (worksheet?._rows || []).forEach((row, offset) => {
    const index = Math.max(1, Math.round(Number(row?.number ?? offset + 1) || 0));
    if (index && rowHasExplicitDimension(row)) rows = Math.max(rows, index);
  });
  (worksheet?.model?.cols || []).forEach((column) => {
    if (!columnHasExplicitDimension(column)) return;
    const max = importedColumnDimensionExtent(column, usedExtent.cols);
    if (max) cols = Math.max(cols, max);
  });
  (worksheet?._columns || []).forEach((column, offset) => {
    const index = Math.max(1, Math.round(Number(column?.number ?? offset + 1) || 0));
    if (shouldIgnoreTailColumnInstance(column, index, usedExtent.cols)) return;
    if (index && columnHasExplicitDimension(column)) cols = Math.max(cols, index);
  });
  return { rows, cols };
}



function worksheetUsedDimensionExtents(worksheet) {
  const dimension = worksheet?.dimensions?.model || worksheet?.model?.dimensions || null;
  return {
    rows: Math.max(0, Number(worksheet?.actualRowCount) || 0, Number(dimension?.bottom) || Number(dimension?.rows) || 0),
    cols: Math.max(0, Number(worksheet?.actualColumnCount) || 0, Number(dimension?.right) || Number(dimension?.cols) || 0),
  };
}



function importedColumnDimensionExtent(column, usedCols = 0) {
  const min = Math.max(1, Math.round(Number(column?.min) || 1));
  const max = Math.max(min, Math.round(Number(column?.max ?? min) || min));
  if (!usedCols || max < EXCEL_MAX_COLS || columnHasStructuralDimension(column)) return max;
  return min <= usedCols ? usedCols : 0;
}



function shouldIgnoreTailColumnInstance(column, index, usedCols = 0) {
  if (!usedCols || index <= usedCols) return false;
  return !columnHasStructuralDimension(column);
}



function columnHasStructuralDimension(column) {
  return column?.hidden === true || Number(column?.outlineLevel) > 0 || column?.collapsed === true;
}



function sheetHfDataDimensions(sheet) {
  const hfData = Array.isArray(sheet?.hfData) ? sheet.hfData : [];
  let cols = 0;
  hfData.forEach((row) => {
    cols = Math.max(cols, Array.isArray(row) ? row.length : 0);
  });
  return { rows: hfData.length, cols };
}



function migrateDenseFormulaDataToCells(sheet) {
  const hfData = Array.isArray(sheet?.hfData) ? sheet.hfData : [];
  if (!sheet || !hfData.length) {
    if (sheet) sheet.hfData = [];
    return sheet;
  }

  sheet.cells = sheet.cells || {};
  hfData.forEach((hfRow, rowIndex) => {
    if (!Array.isArray(hfRow)) return;
    hfRow.forEach((raw, colIndex) => {
      if (raw == null) return;
      const row = rowIndex + 1;
      const col = colIndex + 1;
      const key = cellKey(row, col);
      const existing = sheet.cells[key];
      if (existing) {
        if (existing.raw == null && !existing.formula) {
          existing.raw = raw;
          existing.formula = typeof raw === "string" && raw.startsWith("=") ? raw : null;
          existing.cached = existing.cached ?? (existing.formula ? null : raw);
          existing.display = existing.display || (existing.formula ? "" : String(raw));
        }
        return;
      }
      const formula = typeof raw === "string" && raw.startsWith("=") ? raw : null;
      sheet.cells[key] = internCellModelStylesForSheet(sheet, {
        row,
        col,
        raw,
        formula,
        cached: formula ? null : raw,
        display: formula ? "" : String(raw),
        hyperlink: null,
        comment: null,
        note: null,
        noteVisible: false,
        validation: null,
        css: {},
        hasFill: false,
        hasFontColor: false,
        borders: {},
        numFmt: "",
      });
    });
  });
  sheet.hfData = [];
  return sheet;
}



function formulaCellRawValue(cell) {
  if (!cell) return null;
  if (cell.raw != null) return formulaEngineValueForCell(cell, cell.raw);
  if (cell.formula) return cell.formula;
  return null;
}



function formulaEngineValueForCell(cell, value) {
  if (typeof value === "number" && isDateFormat(cell?.numFmt)) {
    return excelSerialToDate(value);
  }
  return value;
}



function webSheetFormulaTopLevelFunctionName(formula) {
  const text = String(formula || "").trim().replace(/^=/, "").trim();
  return normalizeFormulaEngineFunctionName(parseFormulaFunctionCall(text)?.name || "");
}



function normalizeFormulaEngineFunctionName(name) {
  let text = String(name || "").trim().toUpperCase();
  while (text.startsWith("_XLFN.") || text.startsWith("_XLWS.")) {
    text = text.replace(/^_(?:XLFN|XLWS)\./, "");
  }
  return text;
}



function cachedFormulaEngineValueForCell(cell) {
  if (cell?.cached != null) return formulaEngineValueForCell(cell, cell.cached);
  if (cell?.display != null) return formulaEngineValueForCell(cell, cell.display);
  return null;
}



function shouldUseCachedWebSheetSpillFormulaForEngine(cell, raw, options = {}) {
  if (!options.useCachedWebSheetDynamicArrays || typeof raw !== "string" || !raw.startsWith("=")) return false;
  return WEBSHEET_FORMULA_ENGINE_CACHEABLE_SPILL_NAMES.has(webSheetFormulaTopLevelFunctionName(raw));
}



function hyperFormulaBuildConfig() {
  return {
    licenseKey: HYPERFORMULA_LICENSE,
    leapYear1900: false,
    useArrayArithmetic: true,
    maxRows: EXCEL_MAX_ROWS,
    maxColumns: EXCEL_MAX_COLS,
  };
}



function formulaEngineCellEntry(sheet, cell, options = {}) {
  const row = Number(cell?.row);
  const col = Number(cell?.col);
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 1 || col < 1) return null;
  let raw = formulaCellRawValue(cell);
  if (options.useCachedFormulas && typeof raw === "string" && raw.startsWith("=")) {
    raw = cell.cached ?? null;
  } else if (shouldUseCachedWebSheetSpillFormulaForEngine(cell, raw, options)) {
    raw = cachedFormulaEngineValueForCell(cell);
  }
  if (raw == null || raw === "") return null;
  const sheetRows = Math.max(1, Number(sheet?.rowCount) || sheet?.rows?.length || row);
  const sheetCols = Math.max(1, Number(sheet?.colCount) || sheet?.columns?.length || col);
  if (row > sheetRows || col > sheetCols) return null;
  return { row, col, raw, cell };
}



function denseFormulaDataForSheet(sheet, options = {}) {
  const entries = [];
  let rowCount = 1;
  let colCount = 1;
  Object.values(sheet?.cells || {}).forEach((cell) => {
    const entry = formulaEngineCellEntry(sheet, cell, options);
    if (!entry) return;
    entries.push(entry);
    rowCount = Math.max(rowCount, entry.row);
    colCount = Math.max(colCount, entry.col);
  });
  const data = Array.from({ length: rowCount }, () => Array.from({ length: colCount }, () => null));
  entries.forEach(({ cell, row, col, raw }) => {
    if (typeof raw === "string" && raw.startsWith("=")) {
      raw = translateFormulaForFormulaEngine(raw, {
        model: options.model,
        definedNameAliases: options.definedNameAliases,
        sheetIndex: options.sheetIndex,
        row,
        col,
      });
    }
    data[row - 1][col - 1] = raw ?? null;
  });
  return data;
}



function denseFormulaSheetsForWorkbook(model, options = {}) {
  return Object.fromEntries((model?.sheets || []).map((sheet, sheetIndex) => [
    sheet.name,
    denseFormulaDataForSheet(sheet, { ...options, model, sheetIndex }),
  ]));
}



function applyFormulaEngineIntrinsicNames(hf) {
  [
    ["TRUE", "=TRUE()"],
    ["FALSE", "=FALSE()"],
  ].forEach(([name, formula]) => {
    try {
      if (hf.getNamedExpressionFormula?.(name) == null) hf.addNamedExpression(name, formula);
    } catch (error) {
      console.warn(`Formula intrinsic name skipped: ${name}`, error);
    }
  });
}



function buildFormulaEngine(model) {
  const definedNameAliases = formulaEngineDefinedNameAliases(model);
  const sourceSheets = denseFormulaSheetsForWorkbook(model, { definedNameAliases });

  try {
    state.formulaEngineFallback = false;
    const hf = HyperFormula.buildFromSheets(sourceSheets, hyperFormulaBuildConfig());
    applyFormulaEngineIntrinsicNames(hf);
    applyDefinedNamesToFormulaEngine(hf, model, definedNameAliases);
    return hf;
  } catch (error) {
    if (isFormulaEngineArrayResizeError(error)) {
      const dynamicArraySheets = denseFormulaSheetsForWorkbook(model, {
        definedNameAliases,
        useCachedWebSheetDynamicArrays: true,
      });
      try {
        const hf = HyperFormula.buildFromSheets(dynamicArraySheets, hyperFormulaBuildConfig());
        applyFormulaEngineIntrinsicNames(hf);
        applyDefinedNamesToFormulaEngine(hf, model, definedNameAliases);
        state.formulaEngineFallback = false;
        return hf;
      } catch (dynamicArrayError) {
        console.warn("Formula engine dynamic array cache fallback failed:", dynamicArrayError);
      }
    }
    console.warn("Formula engine fallback:", error);
    state.formulaEngineFallback = true;
    showWarnings([
      "一部の数式構文をブラウザーの数式エンジンで解析できなかったため、Excel ファイルに保存されていた計算結果を使って開きました。通常の数式編集は引き続き処理できます。",
    ]);

    const fallbackSheets = denseFormulaSheetsForWorkbook(model, { useCachedFormulas: true, definedNameAliases });

    try {
      const hf = HyperFormula.buildFromSheets(fallbackSheets, hyperFormulaBuildConfig());
      applyFormulaEngineIntrinsicNames(hf);
      applyDefinedNamesToFormulaEngine(hf, model, definedNameAliases);
      return hf;
    } catch (fallbackError) {
      console.warn("Formula engine empty fallback:", fallbackError);
      const hf = HyperFormula.buildFromSheets(
        denseFormulaSheetsForWorkbook(model, { useCachedFormulas: true, definedNameAliases }),
        hyperFormulaBuildConfig(),
      );
      applyFormulaEngineIntrinsicNames(hf);
      applyDefinedNamesToFormulaEngine(hf, model, definedNameAliases);
      return hf;
    }
  }
}



function isFormulaEngineArrayResizeError(error) {
  return /resizing to smaller array/i.test(String(error?.message || error || ""));
}



function applyDefinedNamesToFormulaEngine(hf, model, definedNameAliases = formulaEngineDefinedNameAliases(model)) {
  (model.definedNames || []).forEach((definition) => {
    if (!definition?.name || definition.hidden) return;
    const engineName = definedNameAliases.byKey.get(definedNameEngineKey(definition)) || definition.name;
    const expression = definedNameExpressionForFormulaEngine(definition.refersTo, definedNameAliases, definition.sheetIndex, model);
    if (!expression) return;
    const scope = undefined;
    try {
      hf.addNamedExpression(engineName, expression, scope);
    } catch (error) {
      console.warn(`Could not add defined name ${definition.name}:`, error);
    }
  });
}



function definedNameExpressionForFormulaEngine(refersTo, definedNameAliases = null, sheetIndex = null, model = null) {
  const text = String(refersTo || "").trim();
  if (!text) return "";
  const expression = qualifySheetLocalDefinedNameExpression(text.startsWith("=") ? text : `=${text}`, model, sheetIndex);
  const structured = translateStructuredReferencesForFormulaEngine(expression, model || state.model, sheetIndex, {});
  return translateDefinedNamesForFormulaEngine(structured, definedNameAliases, sheetIndex);
}



function formulaEngineDefinedNameAliases(model) {
  const aliases = { byKey: new Map(), byName: new Map(), bySheetName: new Map(), bySheetDisplayName: new Map() };
  const usedGeneratedAliases = new Set();
  (model?.definedNames || []).forEach((definition, index) => {
    if (!definition?.name || definition.hidden) return;
    const engineName = definition.scope === "sheet"
      ? generatedFormulaEngineDefinedName(index, usedGeneratedAliases)
      : hyperFormulaSafeDefinedName(definition.name, index, usedGeneratedAliases);
    usedGeneratedAliases.add(engineName.toLowerCase());
    aliases.byKey.set(definedNameEngineKey(definition), engineName);
    if (definition.scope === "sheet") {
      const sheetIndex = Number(definition.sheetIndex);
      if (!aliases.bySheetName.has(sheetIndex)) aliases.bySheetName.set(sheetIndex, new Map());
      aliases.bySheetName.get(sheetIndex).set(definition.name, engineName);
      const sheetName = model?.sheets?.[sheetIndex]?.name;
      if (sheetName) {
        const normalizedSheetName = normalizeFormulaSheetName(sheetName);
        if (!aliases.bySheetDisplayName.has(normalizedSheetName)) aliases.bySheetDisplayName.set(normalizedSheetName, new Map());
        aliases.bySheetDisplayName.get(normalizedSheetName).set(definition.name, engineName);
      }
    } else if (!aliases.byName.has(definition.name)) {
      aliases.byName.set(definition.name, engineName);
    }
  });
  return aliases;
}



function definedNameEngineKey(definition) {
  return `${definition?.scope || "workbook"}:${definition?.sheetIndex ?? ""}:${definition?.name || ""}`;
}



function hyperFormulaSafeDefinedName(name, index, usedGeneratedAliases) {
  const text = String(name || "");
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(text) && !isCellReference(text)) return text;
  return generatedFormulaEngineDefinedName(index, usedGeneratedAliases);
}



function generatedFormulaEngineDefinedName(index, usedGeneratedAliases) {
  let alias = `__WEBSHEET_NAME_${index + 1}`;
  while (usedGeneratedAliases.has(alias.toLowerCase())) {
    alias = `__WEBSHEET_NAME_${index + 1}_${usedGeneratedAliases.size + 1}`;
  }
  usedGeneratedAliases.add(alias.toLowerCase());
  return alias;
}



function translateDefinedNamesForFormulaEngine(formula, definedNameAliases, sheetIndex = null) {
  if (typeof formula !== "string" || !formula.startsWith("=")) return formula;
  const names = new Map(definedNameAliases?.byName || []);
  const localNames = definedNameAliases?.bySheetName?.get(Number(sheetIndex));
  if (localNames?.size) {
    localNames.forEach((alias, name) => names.set(name, alias));
  }
  if (!names.size) return formula;
  const replacements = [...names.entries()]
    .filter(([name, alias]) => name && alias && name !== alias)
    .sort((left, right) => right[0].length - left[0].length);
  if (!replacements.length) return formula;

  let output = "";
  let index = 0;
  while (index < formula.length) {
    const char = formula[index];
    if (char === '"') {
      const end = consumeFormulaDoubleQuotedText(formula, index);
      output += formula.slice(index, end);
      index = end;
      continue;
    }
    if (char === "'") {
      const end = consumeFormulaSingleQuotedSheetName(formula, index);
      output += formula.slice(index, end);
      index = end;
      continue;
    }
    const qualifier = formulaSheetQualifierBeforeIndex(formula, index);
    if (qualifier) {
      const qualifiedNames = definedNameAliases?.bySheetDisplayName?.get(normalizeFormulaSheetName(qualifier.sheetName));
      const replacement = definedNameReplacementAt(formula, index, qualifiedNames);
      if (replacement) {
        output = output.slice(0, Math.max(0, output.length - (index - qualifier.start)));
        output += replacement[1];
        index += replacement[0].length;
        continue;
      }
    }
    const replacement = replacements.find(([name]) =>
      formula.startsWith(name, index) &&
      isDefinedNameTokenBoundary(formula, index, name.length) &&
      formula[index + name.length] !== "!",
    );
    if (replacement) {
      output += replacement[1];
      index += replacement[0].length;
      continue;
    }
    output += char;
    index += 1;
  }
  return output;
}



function definedNameReplacementAt(formula, index, names) {
  if (!names?.size) return null;
  return [...names.entries()]
    .filter(([name, alias]) => name && alias)
    .sort((left, right) => right[0].length - left[0].length)
    .find(([name]) =>
      formula.startsWith(name, index) &&
      isDefinedNameTokenBoundary(formula, index, name.length) &&
      formula[index + name.length] !== "!",
    ) || null;
}



function formulaSheetQualifierBeforeIndex(formula, index) {
  if (formula[index - 1] !== "!") return null;
  const end = index - 1;
  if (formula[end - 1] === "'") {
    let cursor = end - 2;
    while (cursor >= 0) {
      if (formula[cursor] === "'") {
        if (formula[cursor - 1] === "'") {
          cursor -= 2;
          continue;
        }
        return { start: cursor, sheetName: formula.slice(cursor + 1, end - 1).replace(/''/g, "'") };
      }
      cursor -= 1;
    }
    return null;
  }
  let start = end - 1;
  while (start >= 0 && isDefinedNameTokenChar(formula[start])) start -= 1;
  const sheetName = formula.slice(start + 1, end);
  return sheetName ? { start: start + 1, sheetName } : null;
}



function normalizeFormulaSheetName(name) {
  return String(name || "").toLocaleLowerCase("ja-JP");
}



function qualifySheetLocalDefinedNameExpression(expression, model, sheetIndex) {
  if (sheetIndex == null) return expression;
  const sheetName = model?.sheets?.[Number(sheetIndex)]?.name;
  if (!sheetName || typeof expression !== "string" || !expression.startsWith("=")) return expression;
  const qualifier = `${quoteSheetNameForFormula(sheetName)}!`;
  let output = "";
  let index = 0;
  while (index < expression.length) {
    const char = expression[index];
    if (char === '"') {
      const end = consumeFormulaDoubleQuotedText(expression, index);
      output += expression.slice(index, end);
      index = end;
      continue;
    }
    if (char === "'") {
      const end = consumeFormulaSingleQuotedSheetName(expression, index);
      output += expression.slice(index, end);
      index = end;
      continue;
    }
    const ref = expression.slice(index).match(/^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}/)?.[0];
    if (ref && isCellReference(ref) && isFormulaReferenceBoundary(expression, index, ref.length) && expression[index - 1] !== "!") {
      output += `${qualifier}${ref}`;
      index += ref.length;
      continue;
    }
    output += char;
    index += 1;
  }
  return output;
}



function isFormulaReferenceBoundary(formula, start, length) {
  return !isFormulaReferenceTokenChar(formula[start - 1]) && !isFormulaReferenceTokenChar(formula[start + length]);
}



function isFormulaReferenceTokenChar(char) {
  return char === "$" || isDefinedNameTokenChar(char);
}



function consumeFormulaDoubleQuotedText(formula, start) {
  let index = start + 1;
  while (index < formula.length) {
    if (formula[index] === '"' && formula[index + 1] === '"') {
      index += 2;
      continue;
    }
    if (formula[index] === '"') return index + 1;
    index += 1;
  }
  return formula.length;
}



function consumeFormulaSingleQuotedSheetName(formula, start) {
  let index = start + 1;
  while (index < formula.length) {
    if (formula[index] === "'" && formula[index + 1] === "'") {
      index += 2;
      continue;
    }
    if (formula[index] === "'") return index + 1;
    index += 1;
  }
  return formula.length;
}



function isDefinedNameTokenBoundary(formula, start, length) {
  return !isDefinedNameTokenChar(formula[start - 1]) && !isDefinedNameTokenChar(formula[start + length]);
}



function isDefinedNameTokenChar(char) {
  return Boolean(char && /[A-Za-z0-9_.\\\u0080-\uffff]/.test(char));
}



function translateStructuredReferencesForFormulaEngine(formula, model, sheetIndex = 0, context = {}) {
  if (typeof formula !== "string" || !formula.includes("[") || !model?.sheets?.length) return formula;
  const refs = workbookStructuredTableRefs(model);
  if (!refs.length) return formula;
  const namedRefs = refs
    .filter((ref) => ref.name)
    .sort((left, right) => right.name.length - left.name.length);
  let output = "";
  let index = 0;
  while (index < formula.length) {
    const char = formula[index];
    if (char === '"') {
      const end = consumeFormulaDoubleQuotedText(formula, index);
      output += formula.slice(index, end);
      index = end;
      continue;
    }
    if (char === "'") {
      const end = consumeFormulaSingleQuotedSheetName(formula, index);
      output += formula.slice(index, end);
      index = end;
      continue;
    }
    if (char === "[") {
      const token = structuredReferenceTokenAt(formula, index);
      if (token) {
        const tableRef = structuredReferenceContextTable(refs, sheetIndex, context);
        const replacement = tableRef ? structuredReferenceToA1(token.text, tableRef, context) : "";
        if (replacement) {
          output += replacement;
          index = token.end;
          continue;
        }
      }
    }
    const namedRef = namedRefs.find((ref) =>
      structuredReferenceNameMatches(formula, index, ref.name) &&
      formula[index + ref.name.length] === "[" &&
      isDefinedNameTokenBoundary(formula, index, ref.name.length)
    );
    if (namedRef) {
      const token = structuredReferenceTokenAt(formula, index + namedRef.name.length);
      const replacement = token ? structuredReferenceToA1(token.text, namedRef, context) : "";
      if (replacement) {
        output += replacement;
        index = token.end;
        continue;
      }
    }
    output += char;
    index += 1;
  }
  return output;
}



function workbookStructuredTableRefs(model) {
  const refs = [];
  (model?.sheets || []).forEach((sheet, sheetIndex) => {
    (sheet.tables || []).forEach((table, tableIndex) => {
      if (!table?.range) return;
      refs.push({
        sheet,
        sheetIndex,
        table,
        tableIndex,
        name: table.displayName || table.name || `Table${tableIndex + 1}`,
        columns: tableColumnNamesForStructuredRef(sheet, table),
      });
    });
  });
  return refs;
}



function tableColumnNamesForStructuredRef(sheet, table) {
  if (Array.isArray(table?.columns) && table.columns.length) {
    return table.columns.map((column, index) => String(column?.name || `Column${index + 1}`));
  }
  const names = [];
  const range = table?.range;
  if (!range) return names;
  for (let col = range.left; col <= range.right; col += 1) {
    const raw = sheet?.cells?.[cellKey(range.top, col)]?.raw;
    names.push(String(raw || `Column${col - range.left + 1}`));
  }
  return names;
}



function structuredReferenceNameMatches(formula, index, name) {
  return formula.slice(index, index + name.length).toLocaleLowerCase("ja-JP") === String(name || "").toLocaleLowerCase("ja-JP");
}



function structuredReferenceTokenAt(formula, start) {
  if (formula[start] !== "[") return null;
  let depth = 0;
  for (let index = start; index < formula.length; index += 1) {
    const char = formula[index];
    if (char === "'" && index + 1 < formula.length) {
      index += 1;
      continue;
    }
    if (char === "[") depth += 1;
    else if (char === "]") {
      depth -= 1;
      if (depth === 0) return { text: formula.slice(start, index + 1), end: index + 1 };
    }
  }
  return null;
}



function structuredReferenceContextTable(refs, sheetIndex, context = {}) {
  const row = Number(context.row);
  const col = Number(context.col);
  return refs.find((ref) => {
    if (ref.sheetIndex !== Number(sheetIndex)) return false;
    const range = ref.table?.range;
    if (!range || !Number.isFinite(row)) return false;
    const rows = structuredReferenceRows(ref.table, "thisRow", row);
    return row >= rows.top && row <= rows.bottom && (!Number.isFinite(col) || (col >= range.left && col <= range.right));
  }) || refs.find((ref) => ref.sheetIndex === Number(sheetIndex));
}



function structuredReferenceToA1(token, tableRef, context = {}) {
  const spec = parseStructuredReferenceSpec(token);
  if (!spec || !tableRef?.table?.range) return "";
  const range = structuredReferenceResolvedRange(tableRef, spec, context);
  if (!range) return "";
  return structuredReferenceA1(tableRef.sheet, range);
}



function parseStructuredReferenceSpec(token) {
  const inner = String(token || "").slice(1, -1).trim();
  if (!inner) return null;
  const parts = splitStructuredReferenceParts(inner);
  const spec = { rowKind: "", rowKinds: [], currentRow: false, colStart: "", colEnd: "" };
  parts.forEach((part) => {
    const raw = part.trim();
    const currentRowRange = structuredReferenceCurrentRowColumnRangePart(raw);
    if (currentRowRange) {
      spec.currentRow = true;
      spec.rowKind = "thisRow";
      spec.colStart = cleanStructuredReferencePart(currentRowRange.start);
      spec.colEnd = cleanStructuredReferencePart(currentRowRange.end);
      return;
    }
    const range = structuredReferenceColumnRangePart(raw);
    if (range) {
      spec.colStart = cleanStructuredReferencePart(range.start);
      spec.colEnd = cleanStructuredReferencePart(range.end);
      return;
    }
    const wrappedCurrentRowColumn = structuredReferenceWrappedCurrentRowColumnPart(raw);
    if (wrappedCurrentRowColumn) {
      spec.currentRow = true;
      spec.rowKind = "thisRow";
      spec.colStart = cleanStructuredReferencePart(wrappedCurrentRowColumn);
      spec.colEnd = spec.colStart;
      return;
    }
    const escaped = structuredReferencePartHasEscapedSpecial(raw);
    let item = cleanStructuredReferencePart(raw);
    if (!item) return;
    if (/^@\[(.*)\]$/.test(item)) item = item.replace(/^@\[(.*)\]$/, "@$1");
    if (!escaped && item.startsWith("@")) {
      spec.currentRow = true;
      spec.rowKind = "thisRow";
      spec.colStart = cleanStructuredReferencePart(item.slice(1));
      spec.colEnd = spec.colStart;
      return;
    }
    const normalized = item.replace(/^'/, "").toLocaleLowerCase("ja-JP");
    if (!escaped && normalized === "#all") {
      spec.rowKind = "all";
      spec.rowKinds = ["all"];
    } else if (!escaped && normalized === "#data") {
      spec.rowKind = "data";
      spec.rowKinds.push("data");
    } else if (!escaped && normalized === "#headers") {
      spec.rowKind = "headers";
      spec.rowKinds.push("headers");
    } else if (!escaped && normalized === "#totals") {
      spec.rowKind = "totals";
      spec.rowKinds.push("totals");
    }
    else if (!escaped && normalized === "#this row") {
      spec.rowKind = "thisRow";
      spec.rowKinds = ["thisRow"];
      spec.currentRow = true;
    } else if (!spec.colStart) {
      spec.colStart = item;
      spec.colEnd = item;
    }
  });
  return spec;
}



function splitStructuredReferenceParts(inner) {
  const normalized = String(inner || "").trim();
  const parts = [];
  let start = 0;
  let depth = 0;
  for (let index = 0; index < normalized.length; index += 1) {
    const char = normalized[index];
    if (char === "'" && index + 1 < normalized.length) {
      index += 1;
      continue;
    }
    if (char === "[") depth += 1;
    else if (char === "]") depth -= 1;
    else if (char === "," && depth === 0) {
      parts.push(normalized.slice(start, index));
      start = index + 1;
    }
  }
  parts.push(normalized.slice(start));
  return parts;
}



function cleanStructuredReferencePart(part) {
  let text = String(part || "").trim();
  if (text.startsWith("[") && text.endsWith("]") && structuredReferenceTokenAt(text, 0)?.end === text.length) {
    text = text.slice(1, -1).trim();
  }
  return unescapeStructuredReferencePart(text);
}



function unescapeStructuredReferencePart(part) {
  return String(part || "").replace(/'([\[\]#'@])/g, "$1");
}



function structuredReferencePartHasEscapedSpecial(part) {
  let text = String(part || "").trim();
  if (text.startsWith("[") && text.endsWith("]") && structuredReferenceTokenAt(text, 0)?.end === text.length) {
    text = text.slice(1, -1).trim();
  }
  return /'[\[\]#'@]/.test(text);
}



function structuredReferenceColumnRangePart(part) {
  const text = String(part || "").trim();
  const first = structuredReferenceTokenAt(text, 0);
  if (!first) return null;
  const middle = text.slice(first.end).trimStart();
  if (!middle.startsWith(":")) return null;
  const afterColonOffset = first.end + text.slice(first.end).length - middle.length + 1;
  const afterColon = text.slice(afterColonOffset).trimStart();
  const secondOffset = afterColonOffset + text.slice(afterColonOffset).length - afterColon.length;
  const second = structuredReferenceTokenAt(text, secondOffset);
  if (!second || text.slice(second.end).trim()) return null;
  return {
    start: first.text.slice(1, -1),
    end: second.text.slice(1, -1),
  };
}



function structuredReferenceCurrentRowColumnRangePart(part) {
  const text = String(part || "").trim();
  if (!text.startsWith("@")) return null;
  return structuredReferenceColumnRangePart(text.slice(1).trimStart());
}



function structuredReferenceWrappedCurrentRowColumnPart(part) {
  const text = String(part || "").trim();
  if (!text.startsWith("@")) return "";
  const columnText = text.slice(1).trimStart();
  const token = structuredReferenceTokenAt(columnText, 0);
  if (!token || token.end !== columnText.length) return "";
  return token.text.slice(1, -1);
}



function structuredReferenceResolvedRange(tableRef, spec, context = {}) {
  const table = tableRef.table;
  const tableRange = table.range;
  const rowKind = spec.rowKinds?.length ? spec.rowKinds : (spec.rowKind || (spec.currentRow ? "thisRow" : "data"));
  const rowRange = structuredReferenceRows(table, rowKind, context.row);
  if (!rowRange) return null;
  const colRange = structuredReferenceColumns(tableRef, spec);
  if (!colRange) return null;
  return {
    top: rowRange.top,
    bottom: rowRange.bottom,
    left: colRange.left,
    right: colRange.right,
    sheetIndex: tableRange.sheetIndex ?? tableRef.sheetIndex,
  };
}



function structuredReferenceRows(table, rowKind = "data", currentRow = null) {
  if (Array.isArray(rowKind)) return structuredReferenceCompositeRows(table, rowKind, currentRow);
  const range = table?.range;
  if (!range) return null;
  const hasHeader = table.showHeaderRow !== false;
  const hasTotals = table.showTotalsRow === true;
  const dataTop = range.top + (hasHeader ? 1 : 0);
  const dataBottom = range.bottom - (hasTotals ? 1 : 0);
  switch (rowKind) {
    case "all":
      return { top: range.top, bottom: range.bottom };
    case "headers":
      return hasHeader ? { top: range.top, bottom: range.top } : null;
    case "totals":
      return hasTotals ? { top: range.bottom, bottom: range.bottom } : null;
    case "thisRow": {
      const row = Number(currentRow);
      if (!Number.isFinite(row)) return null;
      return { top: row, bottom: row };
    }
    case "data":
    default:
      return dataBottom >= dataTop ? { top: dataTop, bottom: dataBottom } : null;
  }
}



function structuredReferenceCompositeRows(table, rowKinds = [], currentRow = null) {
  const uniqueKinds = [...new Set((rowKinds || []).filter(Boolean))];
  if (!uniqueKinds.length) return structuredReferenceRows(table, "data", currentRow);
  if (uniqueKinds.includes("all")) return structuredReferenceRows(table, "all", currentRow);
  if (uniqueKinds.includes("thisRow")) return uniqueKinds.length === 1 ? structuredReferenceRows(table, "thisRow", currentRow) : null;
  const ranges = uniqueKinds
    .map((kind) => structuredReferenceRows(table, kind, currentRow))
    .filter(Boolean)
    .sort((left, right) => left.top - right.top || left.bottom - right.bottom);
  if (!ranges.length) return null;
  let top = ranges[0].top;
  let bottom = ranges[0].bottom;
  for (const range of ranges.slice(1)) {
    if (range.top > bottom + 1) return null;
    top = Math.min(top, range.top);
    bottom = Math.max(bottom, range.bottom);
  }
  return { top, bottom };
}



function structuredReferenceColumns(tableRef, spec) {
  const range = tableRef.table?.range;
  if (!range) return null;
  const names = tableRef.columns || [];
  if (!spec.colStart) return { left: range.left, right: range.right };
  const startOffset = structuredReferenceColumnOffset(names, spec.colStart);
  const endOffset = structuredReferenceColumnOffset(names, spec.colEnd || spec.colStart);
  if (startOffset < 0 || endOffset < 0) return null;
  return {
    left: range.left + Math.min(startOffset, endOffset),
    right: range.left + Math.max(startOffset, endOffset),
  };
}



function structuredReferenceColumnOffset(names, name) {
  const target = String(name || "").trim().toLocaleLowerCase("ja-JP");
  return (names || []).findIndex((columnNameText) => String(columnNameText || "").trim().toLocaleLowerCase("ja-JP") === target);
}



function structuredReferenceA1(sheet, range) {
  const sheetName = quoteSheetNameForFormula(sheet?.name || "");
  const start = `$${columnName(range.left)}$${range.top}`;
  const end = `$${columnName(range.right)}$${range.bottom}`;
  return `${sheetName}!${start === end ? start : `${start}:${end}`}`;
}



function rewriteStaticCellInfoFormulaReferencesForFormulaEngine(formula, context = {}) {
  if (typeof formula !== "string" || !formula.startsWith("=") || !/cell\s*\(/i.test(formula)) return formula;
  let output = "";
  let index = 0;
  while (index < formula.length) {
    const char = formula[index];
    if (char === '"') {
      const end = consumeFormulaDoubleQuotedText(formula, index);
      output += formula.slice(index, end);
      index = end;
      continue;
    }
    if (char === "'") {
      const end = consumeFormulaSingleQuotedSheetName(formula, index);
      output += formula.slice(index, end);
      index = end;
      continue;
    }
    const match = formula.slice(index).match(/^(?:_xlfn\.)?cell\s*\(/i);
    if (match && !isDefinedNameTokenChar(formula[index - 1])) {
      const openIndex = index + match[0].lastIndexOf("(");
      const closeIndex = formulaEngineFindClosingParen(formula, openIndex);
      if (closeIndex > openIndex) {
        const replacement = formulaEngineStaticCellInfoReplacement(formula.slice(openIndex + 1, closeIndex), context);
        if (replacement != null) {
          output += replacement;
          index = closeIndex + 1;
          continue;
        }
      }
    }
    output += char;
    index += 1;
  }
  return output;
}



function formulaEngineFindClosingParen(formula, openIndex) {
  let depth = 0;
  for (let index = openIndex; index < formula.length; index += 1) {
    const char = formula[index];
    if (char === '"') {
      index = consumeFormulaDoubleQuotedText(formula, index) - 1;
      continue;
    }
    if (char === "'") {
      index = consumeFormulaSingleQuotedSheetName(formula, index) - 1;
      continue;
    }
    if (char === "(") depth += 1;
    else if (char === ")") {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return -1;
}



function formulaEngineStaticCellInfoReplacement(argsSource, context = {}) {
  const args = splitFormulaArguments(argsSource);
  if (args.length < 1 || args.length > 2) return null;
  const infoType = formulaEngineStaticStringLiteral(args[0]);
  if (!FORMULA_ENGINE_STATIC_CELL_INFO_TYPES.has(infoType)) return null;
  const range = args.length > 1
    ? formulaEngineStaticA1ReferenceRange(args[1], context)
    : formulaEngineContextCellRange(context);
  if (!range) return null;
  if (infoType === "row") return String(range.top);
  if (infoType === "col") return String(range.left);
  if (infoType === "address") return formulaEngineStringLiteral(`$${columnName(range.left)}$${range.top}`);
  return null;
}



function formulaEngineStaticStringLiteral(value) {
  const text = String(value || "").trim();
  if (!/^"(?:[^"]|"")*"$/.test(text)) return "";
  return text.slice(1, -1).replace(/""/g, '"').trim().toLocaleLowerCase("ja-JP");
}



function formulaEngineStringLiteral(value) {
  return `"${String(value ?? "").replace(/"/g, '""')}"`;
}



function formulaEngineContextCellRange(context = {}) {
  const row = Number(context.row);
  const col = Number(context.col);
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 1 || col < 1) return null;
  return { top: row, left: col, bottom: row, right: col };
}



function formulaEngineStaticA1ReferenceRange(expression, context = {}) {
  let text = formulaEngineStripOuterParentheses(String(expression || "").trim());
  if (!text) return null;
  const bangIndex = formulaEngineLastSheetBangIndex(text);
  if (bangIndex >= 0) {
    const sheetName = formulaEngineReferenceSheetName(text.slice(0, bangIndex));
    if (sheetName && context.model?.sheets?.length && !formulaEngineReferenceSheetExists(sheetName, context.model)) return null;
    text = text.slice(bangIndex + 1).trim();
  }
  if (!/^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?$/.test(text)) return null;
  return decodeRange(text);
}



function formulaEngineReferenceSheetName(prefix) {
  let text = String(prefix || "").trim();
  if (!text) return "";
  if (text.startsWith("'") && text.endsWith("'")) text = text.slice(1, -1).replace(/''/g, "'");
  const bracketIndex = text.lastIndexOf("]");
  if (bracketIndex >= 0) text = text.slice(bracketIndex + 1);
  return text.trim();
}



function formulaEngineReferenceSheetExists(sheetName, model) {
  const target = normalizeFormulaSheetName(sheetName);
  return (model?.sheets || []).some((sheet) => normalizeFormulaSheetName(sheet.name) === target);
}



function formulaEngineStripOuterParentheses(expression) {
  let text = String(expression || "").trim();
  while (text.startsWith("(") && formulaEngineFindClosingParen(text, 0) === text.length - 1) {
    text = text.slice(1, -1).trim();
  }
  return text;
}



function formulaEngineLastSheetBangIndex(expression) {
  let last = -1;
  let doubleQuoted = false;
  let singleQuoted = false;
  const text = String(expression || "");
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (doubleQuoted) {
      if (char === '"' && next === '"') index += 1;
      else if (char === '"') doubleQuoted = false;
      continue;
    }
    if (singleQuoted) {
      if (char === "'" && next === "'") index += 1;
      else if (char === "'") singleQuoted = false;
      continue;
    }
    if (char === '"') doubleQuoted = true;
    else if (char === "'") singleQuoted = true;
    else if (char === "!") last = index;
  }
  return last;
}



function translateFormulaForFormulaEngine(formula, options = {}) {
  if (typeof formula !== "string" || !formula.startsWith("=")) return formula;
  const model = options.model || state.model;
  const sheetIndex = options.sheetIndex ?? state.activeSheetIndex;
  const structured = translateStructuredReferencesForFormulaEngine(formula, model, sheetIndex, { row: options.row, col: options.col });
  const definedNames = translateDefinedNamesForFormulaEngine(structured, options.definedNameAliases || formulaEngineDefinedNameAliases(model), sheetIndex);
  return rewriteStaticCellInfoFormulaReferencesForFormulaEngine(definedNames, { row: options.row, col: options.col, model });
}



function translateFormulaForCurrentFormulaEngine(formula, sheetIndex = state.activeSheetIndex, context = {}) {
  return translateFormulaForFormulaEngine(formula, {
    model: state.model,
    definedNameAliases: formulaEngineDefinedNameAliases(state.model),
    sheetIndex,
    row: context.row,
    col: context.col,
  });
}



function formulaEngineContentForCurrentModel(content, sheetIndex = state.activeSheetIndex, context = {}) {
  if (typeof content === "string" && content.startsWith("=")) {
    return translateFormulaForCurrentFormulaEngine(content, sheetIndex, context);
  }
  return content;
}



function formulaEngineSheetIndex(sheet) {
  const index = state.model?.sheets?.indexOf(sheet);
  return Number.isInteger(index) && index >= 0 ? index : state.activeSheetIndex;
}



function visibleSheetIndexes() {
  if (!state.model) return [];
  return state.model.sheets
    .map((sheet, index) => (isSheetVisible(sheet) ? index : -1))
    .filter((index) => index >= 0);
}



function hiddenSheetIndexes() {
  if (!state.model) return [];
  return state.model.sheets
    .map((sheet, index) => (!isSheetVisible(sheet) ? index : -1))
    .filter((index) => index >= 0);
}



function isSheetVisible(sheet) {
  return !["hidden", "veryHidden"].includes(String(sheet?.state || "visible"));
}



function workbookTabSelectedSheetIndexes(model = state.model, activeIndex = state.activeSheetIndex) {
  const indexes = (model?.sheets || [])
    .map((sheet, index) => (isSheetVisible(sheet) && sheet.tabSelected === true ? index : -1))
    .filter((index) => index >= 0);
  const active = Number.isFinite(Number(activeIndex)) ? Number(activeIndex) : 0;
  if (!indexes.includes(active) && isSheetVisible(model?.sheets?.[active])) indexes.push(active);
  const unique = [...new Set(indexes)].sort((left, right) => left - right);
  return unique.length > 1 ? unique : [];
}



function ensureActiveSheetVisible() {
  if (!state.model?.sheets?.length) return;
  const visibleIndexes = visibleSheetIndexes();
  if (!visibleIndexes.length) {
    state.model.sheets[0].state = "visible";
    state.activeSheetIndex = 0;
    state.selectedSheetIndexes = [];
    return;
  }
  if (!visibleIndexes.includes(state.activeSheetIndex)) {
    state.activeSheetIndex = visibleIndexes[0];
    state.selectedSheetIndexes = [];
  }
}



function renderWorkbook() {
  ensureActiveSheetVisible();
  syncWorkbookTitle();
  const warnings = state.model.sheets.flatMap((sheet) => sheet.warnings || []);
  if (state.formulaEngineFallback) {
    warnings.unshift("数式エンジンはフォールバック モードです。複雑な数式は保存済みの計算結果を表示します。");
  }
  if (warnings.length) {
    showWarnings(warnings);
  } else {
    clearWarnings();
  }

  renderSheetTabs();
  renderSheet();
  renderWorkbookWindowBar();
  updateSheetFooterUi();
  syncFullScreenViewClass();
  syncWorkbookChromeVisibility();
  syncViewRibbonButtons();
  syncPageLayoutRibbonControls();
  syncDrawRibbonButtons();
  syncDataRibbonControls();
}



function renderSheetTabs() {
  if (!state.model) return;
  const visibleIndexes = visibleSheetIndexes();
  const groupSelection = new Set((state.selectedSheetIndexes || []).filter((index) => visibleIndexes.includes(index)));
  const html = visibleIndexes
    .map((index) => {
      const sheet = state.model.sheets[index];
      const active = index === state.activeSheetIndex ? " is-active" : "";
      const group = groupSelection.has(index) ? " is-group-selected" : "";
      const renaming = state.sheetTabRename?.sheetIndex === index;
      const renameClass = renaming ? " is-renaming" : "";
      const styleParts = [];
      if (sheet.tabColor) styleParts.push(`--sheet-tab-color:${escapeAttr(sheet.tabColor)}`);
      if (renaming) {
        const widthPx = Number(state.sheetTabRename?.widthPx || 0);
        if (Number.isFinite(widthPx) && widthPx > 0) {
          styleParts.push(`width:${Math.ceil(widthPx)}px`);
        }
      }
      const tabColor = sheet.tabColor ? ` data-tab-color="true"` : "";
      const styleAttr = styleParts.length ? ` style="${styleParts.join(";")}"` : "";
      const className = `sheet-tab${active}${group}${renameClass}`;
      const attrs = `class="${className}" data-sheet-index="${index}"${tabColor}${styleAttr}`;
      if (renaming) {
        const value = state.sheetTabRename.value ?? sheet.name;
        return `<div ${attrs} role="button" tabindex="0" aria-label="${escapeAttr(sheet.name)} の名前を変更"><input class="sheet-tab-rename-input" type="text" value="${escapeAttr(value)}" maxlength="31" spellcheck="false" aria-label="シート名" /></div>`;
      }
      return `<button ${attrs} type="button">${escapeHtml(sheet.name)}</button>`;
    })
    .join("");
  $sheetTabs.html(html);
  requestAnimationFrame(() => {
    if (state.sheetTabRename) {
      focusSheetTabRenameInput();
      return;
    }
    $sheetTabs.find(".sheet-tab.is-active")[0]?.scrollIntoView({ block: "nearest", inline: "nearest" });
  });
}



function sheetTabButtons() {
  return Array.from($sheetTabs[0]?.querySelectorAll(".sheet-tab") || []).filter((tab) => tab.offsetParent !== null);
}



function sheetTabButtonByIndex(sheetIndex) {
  if (!Number.isInteger(Number(sheetIndex))) return null;
  return $sheetTabs[0]?.querySelector(`.sheet-tab[data-sheet-index="${Number(sheetIndex)}"]`) || null;
}



function focusSheetTabByIndex(sheetIndex) {
  const tab = sheetTabButtonByIndex(sheetIndex);
  if (!tab) return false;
  tab.scrollIntoView({ block: "nearest", inline: "nearest" });
  tab.focus({ preventScroll: true });
  return true;
}



function scheduleFocusSheetTabByIndex(sheetIndex, frames = 12, options = {}) {
  let remaining = Math.max(1, Number(frames) || 1);
  const forceUntil = Number(options.forceUntil || 0);
  const focus = () => {
    if (!shouldContinueSheetTabFocus(sheetIndex, { forceUntil })) return;
    focusSheetTabByIndex(sheetIndex);
    if (document.activeElement === sheetTabButtonByIndex(sheetIndex)) return;
    remaining -= 1;
    if (remaining > 0) requestAnimationFrame(focus);
  };
  focus();
  requestAnimationFrame(focus);
  [0, 50, 120, 240].forEach((delay) => {
    setTimeout(() => {
      if (shouldContinueSheetTabFocus(sheetIndex, { forceUntil })) focusSheetTabByIndex(sheetIndex);
    }, delay);
  });
}



function shouldContinueSheetTabFocus(sheetIndex, options = {}) {
  if (state.sheetTabRename) return false;
  const tab = sheetTabButtonByIndex(sheetIndex);
  if (!tab) return false;
  const active = document.activeElement;
  if (active === tab) return false;
  if (Number(options.forceUntil || 0) > nowMs()) return true;
  if (active && !active.isConnected) return true;
  if (!active || active === document.body) return true;
  return Boolean(active.closest?.(".sheet-tabs"));
}



function scheduleSheetTabKeyboardActivationStabilizer(sheetIndex, sheetName = "") {
  const token = uniqueWorksheetObjectId("sheet-tab-keyboard");
  const until = nowMs() + 800;
  state.sheetTabKeyboardActivation = { token, sheetIndex, until, sheetName };
  [0, 50, 120, 220, 360, 520, 760].forEach((delay) => {
    setTimeout(() => stabilizeSheetTabKeyboardActivation(token), delay);
  });
}



function stabilizeSheetTabKeyboardActivation(token) {
  const activation = state.sheetTabKeyboardActivation;
  if (!activation || activation.token !== token) return;
  if (nowMs() > activation.until || state.sheetTabRename) {
    state.sheetTabKeyboardActivation = null;
    return;
  }
  const sheetIndex = activation.sheetIndex;
  if (!state.model?.sheets?.[sheetIndex] || !isSheetVisible(state.model.sheets[sheetIndex])) {
    state.sheetTabKeyboardActivation = null;
    return;
  }
  if (state.activeSheetIndex !== sheetIndex) {
    setActiveSheet(sheetIndex);
    if (activation.sheetName) setStatus(`${activation.sheetName} シートを表示しました。`);
  }
  scheduleFocusSheetTabByIndex(sheetIndex, 4, { forceUntil: Math.min(activation.until, nowMs() + 120) });
  if (document.activeElement === sheetTabButtonByIndex(sheetIndex) && state.activeSheetIndex === sheetIndex) {
    state.sheetTabKeyboardActivation = null;
  }
}



function focusSheetTabRenameInput({ select = true } = {}) {
  const input = $sheetTabs[0]?.querySelector(".sheet-tab-rename-input");
  if (!input) return false;
  input.focus({ preventScroll: true });
  if (select) input.select?.();
  resizeSheetTabRenameInput(input);
  return true;
}



function resizeSheetTabRenameInput(input) {
  if (!input) return;
  const lockedWidthPx = Number(state.sheetTabRename?.widthPx || 0);
  if (Number.isFinite(lockedWidthPx) && lockedWidthPx > 0) {
    input.style.width = "100%";
    input.style.maxWidth = "none";
    return;
  }
  const textLength = Math.max(3, String(input.value || "").length);
  input.style.width = `${Math.min(220, Math.max(48, textLength * 9 + 24))}px`;
}



function beginSheetTabInlineRename(sheetIndex) {
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) return false;
  if (sheetGroupActionIndexes(sheetIndex).length > 1) {
    setStatus("複数のシートを選択している場合、シート名は変更できません。");
    return false;
  }
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeDataDialog({ keepState: true });
  const existingTab = sheetTabButtonByIndex(sheetIndex);
  const widthPx = existingTab ? existingTab.getBoundingClientRect().width : 0;
  state.activeSheetIndex = sheetIndex;
  state.selectedSheetIndexes = [];
  state.sheetTabRename = {
    sheetIndex,
    value: sheet.name,
    widthPx: Number.isFinite(widthPx) && widthPx > 0 ? widthPx : null,
  };
  renderSheetTabs();
  setStatus(`${sheet.name} の名前を変更します。`);
  return true;
}



function stopSheetTabRenameInputPropagation(event) {
  event.stopPropagation();
}



function handleSheetTabRenameInput(event) {
  if (!state.sheetTabRename) return;
  state.sheetTabRename.value = event.currentTarget.value;
  resizeSheetTabRenameInput(event.currentTarget);
}



function handleSheetTabRenameInputKeydown(event) {
  event.stopPropagation();
  if (isElementImeCompositionEvent(event)) return;
  if (event.key === "Enter") {
    event.preventDefault();
    commitSheetTabInlineRename();
  } else if (event.key === "Escape") {
    event.preventDefault();
    cancelSheetTabInlineRename();
  }
}



function handleSheetTabRenameInputBlur() {
  if (!state.sheetTabRename) return;
  setTimeout(() => {
    if (!state.sheetTabRename) return;
    if ($sheetTabs[0]?.contains(document.activeElement)) return;
    commitSheetTabInlineRename({ restoreFocus: false });
  }, 0);
}



function commitSheetTabRenameOnOutsidePointer(event) {
  if (!state.sheetTabRename) return;
  if (event.target?.closest?.(".sheet-tab.is-renaming")) return;
  commitSheetTabInlineRename({ restoreFocus: false });
}



function commitSheetTabInlineRename({ restoreFocus = true } = {}) {
  const rename = state.sheetTabRename;
  if (!rename) return false;
  const sheetIndex = rename.sheetIndex;
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) {
    state.sheetTabRename = null;
    renderSheetTabs();
    return false;
  }
  const input = $sheetTabs[0]?.querySelector(".sheet-tab-rename-input");
  const requestedName = input?.value ?? rename.value ?? "";
  const validation = validateSheetName(requestedName, sheetIndex);
  if (!validation.valid) {
    setStatus(validation.message);
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    requestAnimationFrame(() => focusSheetTabRenameInput());
    return false;
  }
  state.sheetTabRename = null;
  if (validation.name === sheet.name) {
    renderSheetTabs();
    if (restoreFocus) scheduleFocusSheetTabByIndex(sheetIndex);
    setStatus("ワークシート名は変更されていません。");
    return true;
  }
  renameSheetAt(sheetIndex, validation.name);
  if (restoreFocus) scheduleFocusSheetTabByIndex(sheetIndex);
  return true;
}



function cancelSheetTabInlineRename({ restoreFocus = true } = {}) {
  const sheetIndex = state.sheetTabRename?.sheetIndex ?? state.activeSheetIndex;
  if (!state.sheetTabRename) return false;
  state.sheetTabRename = null;
  renderSheetTabs();
  if (restoreFocus) scheduleFocusSheetTabByIndex(sheetIndex);
  setStatus("名前の変更をキャンセルしました。");
  return true;
}



function handleSheetTabKeydown(event) {
  const key = String(event.key || "");
  const normalizedKey = key.toLowerCase();
  const isContextShortcut = normalizedKey === "contextmenu" || normalizedKey === "apps" || (normalizedKey === "f10" && event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey);
  if (isContextShortcut) {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    openSheetTabContextMenuByKeyboard(event.currentTarget);
    return;
  }
  if (event.ctrlKey || event.metaKey || event.altKey) return;

  const tabs = sheetTabButtons();
  const currentPosition = tabs.indexOf(event.currentTarget);
  if (currentPosition < 0) return;

  const moveFocus = (nextPosition) => {
    const tab = tabs[Math.max(0, Math.min(tabs.length - 1, nextPosition))];
    if (!tab || tab === event.currentTarget) return;
    event.preventDefault();
    event.stopPropagation();
    tab.scrollIntoView({ block: "nearest", inline: "nearest" });
    tab.focus({ preventScroll: true });
    const sheetIndex = Number(tab.dataset.sheetIndex);
    const sheetName = state.model?.sheets?.[sheetIndex]?.name || tab.textContent.trim();
    if (sheetName) setStatus(`${sheetName} シート タブに移動しました。`);
  };

  switch (key) {
    case "ArrowLeft":
      moveFocus(currentPosition - 1);
      break;
    case "ArrowRight":
      moveFocus(currentPosition + 1);
      break;
    case "Home":
      moveFocus(0);
      break;
    case "End":
      moveFocus(tabs.length - 1);
      break;
    case "Enter":
    case " ": {
      event.preventDefault();
      event.stopPropagation();
      const index = Number(event.currentTarget.dataset.sheetIndex);
      if (!state.model?.sheets?.[index] || !isSheetVisible(state.model.sheets[index])) return;
      const sheetName = state.model.sheets[index].name;
      setActiveSheet(index);
      scheduleFocusSheetTabByIndex(index, 24, { forceUntil: nowMs() + 800 });
      scheduleSheetTabKeyboardActivationStabilizer(index, sheetName);
      setStatus(`${sheetName} シートを表示しました。`);
      break;
    }
    case "Escape":
      event.preventDefault();
      event.stopPropagation();
      restoreWorksheetFocusAfterCommandPopup();
      break;
    default:
      break;
  }
}



function handleSheetTabDoubleClick(event) {
  event.preventDefault();
  event.stopPropagation();
  const sheetIndex = Number(event.currentTarget.dataset.sheetIndex);
  if (!state.model?.sheets?.[sheetIndex] || !isSheetVisible(state.model.sheets[sheetIndex])) return;
  beginSheetTabInlineRename(sheetIndex);
}



function startSheetTabDrag(event) {
  if (event.isDefaultPrevented?.()) return;
  if (event.target?.closest?.(".sheet-tab-rename-input")) return;
  if (state.sheetTabRename) return;
  if (event.button !== 0 || !state.model?.sheets?.length) return;
  const sheetIndex = Number(event.currentTarget.dataset.sheetIndex);
  if (!state.model?.sheets?.[sheetIndex] || !isSheetVisible(state.model.sheets[sheetIndex])) return;
  const sourceIndexes = sheetGroupActionIndexes(sheetIndex);
  state.sheetTabDrag = {
    pointerId: event.pointerId,
    sourceIndex: sheetIndex,
    sourceIndexes,
    startX: event.clientX,
    startY: event.clientY,
    clientX: event.clientX,
    clientY: event.clientY,
    beforeIndex: sheetIndex,
    copy: event.ctrlKey || event.metaKey,
    dragging: false,
    sourceElement: event.currentTarget,
    indicator: null,
    autoScrollFrame: null,
    autoScrollVelocity: 0,
  };
  event.currentTarget.setPointerCapture?.(event.pointerId);
}



function handleSheetTabDragMove(event) {
  const drag = state.sheetTabDrag;
  if (!drag || (drag.pointerId != null && event.pointerId != null && drag.pointerId !== event.pointerId)) return;
  const distance = Math.max(Math.abs(event.clientX - drag.startX), Math.abs(event.clientY - drag.startY));
  if (!drag.dragging && distance < 6) return;
  if (!drag.dragging) {
    drag.dragging = true;
    document.body.classList.add("is-sheet-tab-dragging");
    markSheetTabDragSources(drag.sourceIndexes, true);
  }
  event.preventDefault();
  drag.copy = event.ctrlKey || event.metaKey;
  drag.clientX = event.clientX;
  drag.clientY = event.clientY;
  const target = sheetTabDropTargetFromClientX(event.clientX);
  if (!target) return;
  drag.beforeIndex = target.beforeIndex;
  updateSheetTabDropIndicator(target, drag.copy);
  updateSheetTabDragAutoScroll(event.clientX, event.clientY);
  const mode = drag.copy ? "コピー" : "移動";
  setStatus(`${drag.sourceIndexes.length > 1 ? `${drag.sourceIndexes.length} 個のシート` : state.model.sheets[drag.sourceIndex]?.name || "シート"}を${mode}する位置を指定してください。`);
}



function finishSheetTabDrag(event = null) {
  const drag = state.sheetTabDrag;
  if (!drag || (event && drag.pointerId != null && event.pointerId != null && drag.pointerId !== event.pointerId)) return;
  state.sheetTabDrag = null;
  markSheetTabDragSources(drag.sourceIndexes, false);
  document.body.classList.remove("is-sheet-tab-dragging");
  cancelSheetTabDragAutoScroll(drag);
  removeSheetTabDropIndicator(drag);
  drag.sourceElement?.releasePointerCapture?.(drag.pointerId);
  if (!drag.dragging || event?.type === "pointercancel") return;
  event?.preventDefault?.();
  event?.stopPropagation?.();
  state.suppressNextSheetTabClick = true;
  setTimeout(() => {
    state.suppressNextSheetTabClick = false;
  }, 0);
  if (drag.copy) {
    copySheetsToPosition(drag.sourceIndexes, drag.beforeIndex);
  } else {
    moveSheetsToBeforeIndex(drag.sourceIndexes, drag.beforeIndex);
  }
  scheduleFocusSheetTabByIndex(state.activeSheetIndex);
}



function updateSheetTabDragAutoScroll(clientX, clientY) {
  const drag = state.sheetTabDrag;
  const tabs = $sheetTabs[0];
  if (!drag || !tabs) return;
  drag.clientX = clientX;
  drag.clientY = clientY;
  const rect = tabs.getBoundingClientRect();
  const edge = 34;
  const maxVelocity = 16;
  const maxScroll = Math.max(0, tabs.scrollWidth - tabs.clientWidth);
  let velocity = 0;
  if (clientX < rect.left + edge && tabs.scrollLeft > 0) {
    velocity = -Math.max(4, Math.round(((rect.left + edge - clientX) / edge) * maxVelocity));
  } else if (clientX > rect.right - edge && tabs.scrollLeft < maxScroll) {
    velocity = Math.max(4, Math.round(((clientX - (rect.right - edge)) / edge) * maxVelocity));
  }
  drag.autoScrollVelocity = velocity;
  if (velocity && !drag.autoScrollFrame) {
    drag.autoScrollFrame = requestAnimationFrame(runSheetTabDragAutoScroll);
  } else if (!velocity) {
    cancelSheetTabDragAutoScroll(drag);
  }
}



function runSheetTabDragAutoScroll() {
  const drag = state.sheetTabDrag;
  const tabs = $sheetTabs[0];
  if (!drag || !tabs || !drag.dragging) return;
  drag.autoScrollFrame = null;
  const maxScroll = Math.max(0, tabs.scrollWidth - tabs.clientWidth);
  const before = tabs.scrollLeft;
  tabs.scrollLeft = Math.max(0, Math.min(maxScroll, before + drag.autoScrollVelocity));
  if (tabs.scrollLeft !== before) {
    const target = sheetTabDropTargetFromClientX(drag.clientX);
    if (target) {
      drag.beforeIndex = target.beforeIndex;
      updateSheetTabDropIndicator(target, drag.copy);
    }
  }
  updateSheetTabDragAutoScroll(drag.clientX, drag.clientY);
}



function cancelSheetTabDragAutoScroll(drag = state.sheetTabDrag) {
  if (drag?.autoScrollFrame) cancelAnimationFrame(drag.autoScrollFrame);
  if (drag) {
    drag.autoScrollFrame = null;
    drag.autoScrollVelocity = 0;
  }
}



function sheetTabDropTargetFromClientX(clientX) {
  const tabs = sheetTabButtons();
  if (!tabs.length || !state.model?.sheets?.length) return null;
  let position = tabs.length;
  for (let index = 0; index < tabs.length; index += 1) {
    const rect = tabs[index].getBoundingClientRect();
    if (clientX < rect.left + rect.width / 2) {
      position = index;
      break;
    }
  }
  const referenceTab = position >= tabs.length ? tabs[tabs.length - 1] : tabs[position];
  const referenceRect = referenceTab.getBoundingClientRect();
  const referenceIndex = Number(referenceTab.dataset.sheetIndex);
  const beforeIndex = position >= tabs.length
    ? Math.min(state.model.sheets.length, referenceIndex + 1)
    : Math.max(0, referenceIndex);
  const indicatorX = position >= tabs.length ? referenceRect.right : referenceRect.left;
  return {
    beforeIndex,
    x: Math.round(indicatorX),
    y: Math.round(referenceRect.top),
    height: Math.round(referenceRect.height),
  };
}



function updateSheetTabDropIndicator(target, copy) {
  const drag = state.sheetTabDrag;
  if (!drag) return;
  if (!drag.indicator) {
    drag.indicator = document.createElement("div");
    drag.indicator.className = "sheet-tab-drop-indicator";
    drag.indicator.setAttribute("aria-hidden", "true");
    document.body.appendChild(drag.indicator);
  }
  drag.indicator.classList.toggle("is-copy", Boolean(copy));
  drag.indicator.style.left = `${target.x}px`;
  drag.indicator.style.top = `${target.y}px`;
  drag.indicator.style.height = `${Math.max(16, target.height)}px`;
}



function removeSheetTabDropIndicator(drag = state.sheetTabDrag) {
  drag?.indicator?.remove?.();
  if (drag) drag.indicator = null;
}



function markSheetTabDragSources(sourceIndexes, active) {
  const sourceSet = new Set((sourceIndexes || []).map((index) => String(index)));
  sheetTabButtons().forEach((tab) => {
    tab.classList.toggle("is-drag-source", active && sourceSet.has(tab.dataset.sheetIndex));
  });
}



function handleSheetTabClick(event) {
  if (event.target?.closest?.(".sheet-tab-rename-input")) return;
  if (state.suppressNextSheetTabClick) {
    event.preventDefault();
    event.stopPropagation();
    state.suppressNextSheetTabClick = false;
    return;
  }
  const index = Number($(event.currentTarget).data("sheetIndex"));
  if (!state.model?.sheets?.[index] || !isSheetVisible(state.model.sheets[index])) return;
  if (state.formulaReferenceSheetNavigation?.suppressClickIndex === index) {
    state.formulaReferenceSheetNavigation.suppressClickIndex = null;
    const argumentInput = activeFunctionArgumentReferenceInput();
    if (argumentInput) {
      argumentInput.focus({ preventScroll: true });
      updateFormulaReferenceUi();
      return;
    }
    const input = $formulaInput[0];
    if (input && formulaTextStartsFormula(input.value || "")) {
      input.focus({ preventScroll: true });
      updateFormulaReferenceUi();
      updateFormulaAutocompleteForEditor(input);
    }
    return;
  }
  if (event.shiftKey) {
    selectSheetTabRange(index);
    return;
  }
  if (event.ctrlKey || event.metaKey) {
    toggleSheetTabGroupSelection(index);
    return;
  }
  if (index === state.activeSheetIndex && !(state.selectedSheetIndexes || []).length) return;
  setActiveSheet(index);
}



function handleFormulaReferenceSheetTabPointerDown(event) {
  if (event.target?.closest?.(".sheet-tab-rename-input")) return;
  if (state.sheetTabRename) return;
  if (event.button !== 0 || event.shiftKey || event.ctrlKey || event.metaKey) return;
  const index = Number($(event.currentTarget).data("sheetIndex"));
  if (!state.model?.sheets?.[index] || !isSheetVisible(state.model.sheets[index])) return;
  if (!switchActiveSheetForFormulaReference(index)) return;
  event.preventDefault();
  event.stopPropagation();
  state.formulaReferenceSheetNavigation = {
    ...(state.formulaReferenceSheetNavigation || {}),
    suppressClickIndex: index,
  };
}



function formulaReferenceSheetNavigationContext() {
  const editor = activeFormulaReferenceEditor();
  if (!editor || !state.model) return null;
  const text = formulaEditorText(editor);
  const caret = formulaEditorCaretOffset(editor);
  if (isFunctionArgumentReferenceEditor(editor)) {
    const selection = formulaEditorSelectionOffsets(editor);
    return {
      editor,
      isArgumentInput: true,
      text,
      selection,
      sourceSheetIndex: formulaReferenceEditorSheetIndex(editor),
    };
  }
  if (!formulaTextStartsFormula(text) || formulaCaretIsInsideString(text, caret)) return null;
  const selection = formulaEditorSelectionOffsets(editor);
  if (editor === $formulaInput[0]) {
    if (!state.formulaEdit) beginFormulaBarEdit();
    return {
      editor,
      text,
      selection,
      sourceSheetIndex: state.formulaEdit?.sheetIndex ?? state.selected?.sheetIndex ?? state.activeSheetIndex,
    };
  }

  const point = cellPointFromElement(editor);
  if (!point) return null;
  const sheet = state.model.sheets?.[point.sheetIndex];
  if (!sheet) return null;
  state.formulaEdit = {
    sheetIndex: point.sheetIndex,
    row: point.row,
    col: point.col,
    before: $(editor).data("beforeEdit") ?? getCellRawInput(sheet, point.row, point.col),
  };
  state.selected = { sheetIndex: point.sheetIndex, row: point.row, col: point.col };
  state.selectionRange = selectionRangeFromPoints(state.selected, state.selected);
  state.selectionRanges = [state.selectionRange];
  state.selectionAnchor = state.selected;
  state.selectionStart = state.selected;
  $formulaInput.val(text);
  return { editor: $formulaInput[0], text, selection, sourceSheetIndex: point.sheetIndex };
}



function switchActiveSheetForFormulaReference(index) {
  const context = formulaReferenceSheetNavigationContext();
  if (!context) return false;
  if (index === state.activeSheetIndex) {
    restoreFormulaReferenceNavigationEditor(context);
    updateFormulaReferenceUi();
    return true;
  }
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeTranslationPane();
  closeFormulaAutocomplete({ restoreFocus: false });
  state.formulaReferenceSheetNavigation = {
    ...(state.formulaReferenceSheetNavigation || {}),
    sourceSheetIndex: context.sourceSheetIndex,
  };
  detachFormulaReferenceCellEditor(context);
  state.activeSheetIndex = index;
  state.selectedSheetIndexes = [];
  state.selectedImage = null;
  state.selectedImages = [];
  state.selectionDrag = null;
  state.isSelecting = false;
  resetLazyScrollLocks();
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  restoreFormulaReferenceNavigationEditor(context);
  updateFormulaReferenceUi();
  if (!context.isArgumentInput) updateFormulaAutocompleteForEditor($formulaInput[0]);
  setStatus(`${state.model.sheets[index].name} で参照するセルを選択してください。`);
  scheduleWorkbookPersistence();
  return true;
}



function detachFormulaReferenceCellEditor(context) {
  if (!context || context.isArgumentInput || context.editor === $formulaInput[0]) return;
  const editor = context.editor?.closest?.(".sheet-cell[contenteditable='true']");
  if (!editor) return;
  const $editor = $(editor);
  $editor.data("skipBlurCommit", true);
  editor.removeAttribute("contenteditable");
  editor.classList.remove("is-formula-editing");
  clearCellEditSelection(editor);
  stopEditAutoScroll();
}



function restoreFormulaReferenceNavigationEditor(context) {
  if (!context) return;
  if (context.isArgumentInput) {
    const input = context.editor?.isConnected ? context.editor : activeFunctionArgumentReferenceInput();
    if (!input) return;
    input.focus({ preventScroll: true });
    input.setSelectionRange?.(context.selection.start, context.selection.end);
    return;
  }
  const input = $formulaInput[0];
  if (!input) return;
  if (Number.isInteger(context.sourceSheetIndex)) {
    input.dataset.formulaReferenceSourceSheetIndex = String(context.sourceSheetIndex);
  }
  input.value = context.text;
  input.focus({ preventScroll: true });
  input.setSelectionRange?.(context.selection.start, context.selection.end);
}



function selectSheetTabRange(index) {
  const visible = visibleSheetIndexes();
  const anchor = visible.includes(state.activeSheetIndex) ? state.activeSheetIndex : visible[0];
  const anchorPosition = visible.indexOf(anchor);
  const targetPosition = visible.indexOf(index);
  if (anchorPosition < 0 || targetPosition < 0) {
    setActiveSheet(index);
    return;
  }
  const [start, end] = anchorPosition <= targetPosition ? [anchorPosition, targetPosition] : [targetPosition, anchorPosition];
  state.activeSheetIndex = index;
  state.selectedSheetIndexes = visible.slice(start, end + 1);
  renderWorkbook();
  setStatus(`${state.selectedSheetIndexes.length} 個の表示ワークシートを選択しました。`);
}



function toggleSheetTabGroupSelection(index) {
  const visible = new Set(visibleSheetIndexes());
  const current = new Set((state.selectedSheetIndexes || []).filter((item) => visible.has(item)));
  if (!current.size && visible.has(state.activeSheetIndex)) current.add(state.activeSheetIndex);
  if (current.has(index) && current.size > 1) {
    current.delete(index);
    state.activeSheetIndex = current.has(state.activeSheetIndex) ? state.activeSheetIndex : [...current][0];
  } else {
    current.add(index);
    state.activeSheetIndex = index;
  }
  state.selectedSheetIndexes = [...current].sort((left, right) => left - right);
  renderWorkbook();
  setStatus(`${state.selectedSheetIndexes.length} 個の表示ワークシートを選択しました。`);
}



function setActiveSheet(index) {
  if (!state.model || !state.model.sheets[index] || !isSheetVisible(state.model.sheets[index])) return;
  saveActiveWindowViewState();
  state.activeSheetIndex = index;
  state.showFormulas = state.model.sheets[index].showFormulas === true;
  state.selectedSheetIndexes = [];
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeTranslationPane();
  state.annotationEditor = null;
  state.selected = null;
  state.selectedImage = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  state.selectionDrag = null;
  resetLazyScrollLocks();
  state.formulaEdit = null;
  state.formulaReferenceSheetNavigation = null;
  state.isSelecting = false;
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  saveActiveWindowViewState();
  renderWorkbook();
  scheduleWorkbookPersistence();
}



function addBlankSheet() {
  if (!state.model) return;
  const name = uniqueSheetName(state.model.sheets.map((sheet) => sheet.name));
  const insertIndex = state.model.sheets.length;
  const activeSheetBefore = state.activeSheetIndex;
  const sheet = createBlankSheet(name);
  state.model.sheets.push(sheet);
  state.hf = buildFormulaEngine(state.model);
  state.activeSheetIndex = insertIndex;
  state.selectedSheetIndexes = [];
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeTranslationPane();
  state.annotationEditor = null;
  state.selected = { sheetIndex: insertIndex, row: 1, col: 1 };
  state.selectedImage = null;
  state.selectionRange = { sheetIndex: insertIndex, top: 1, left: 1, bottom: 1, right: 1 };
  state.selectionRanges = [state.selectionRange];
  state.selectionAnchor = state.selected;
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.fillDrag = null;
  state.formulaEdit = null;
  state.annotationEditor = null;
  state.isSelecting = false;
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  pushSheetInsertHistory(insertIndex, sheet, activeSheetBefore, insertIndex, `${name} ワークシート追加`, {
    selectionAfter: state.selectionRange,
  });
  setStatus(`新しいワークシートを追加しました：${name}`);
  scheduleWorkbookPersistence();
}



function insertBlankSheetAt(sheetIndex) {
  if (!state.model) return;
  const insertAt = Math.max(0, Math.min(sheetIndex, state.model.sheets.length));
  const name = uniqueSheetName(state.model.sheets.map((sheet) => sheet.name));
  const activeSheetBefore = state.activeSheetIndex;
  const sheet = createBlankSheet(name);
  state.model.sheets.splice(insertAt, 0, sheet);
  state.activeSheetIndex = insertAt;
  clearSheetSelectionState();
  refreshWorkbookAfterSheetMetadataChange();
  pushSheetInsertHistory(insertAt, sheet, activeSheetBefore, insertAt, `${name} ワークシート挿入`, {
    selectionAfter: null,
    selectionAfterRanges: [],
  });
  setStatus(`ワークシートを挿入しました：${name}`);
}



function insertSpecialSheetAt(sheetIndex, templateType) {
  if (!state.model) return;
  const meta = SPECIAL_SHEET_INSERT_META[templateType];
  if (!meta) return;
  const insertAt = Math.max(0, Math.min(sheetIndex, state.model.sheets.length));
  const activeSheetBefore = state.activeSheetIndex;
  const name = uniqueSheetNameWithBase(meta.nameBase, state.model.sheets.map((sheet) => sheet.name));
  const sheet = createSpecialSheet(name, templateType);
  state.model.sheets.splice(insertAt, 0, sheet);
  state.activeSheetIndex = insertAt;
  clearSheetSelectionState();
  refreshWorkbookAfterSheetMetadataChange();
  pushSheetInsertHistory(insertAt, sheet, activeSheetBefore, insertAt, `${name} ${meta.statusNoun}挿入`, {
    selectionAfter: null,
    selectionAfterRanges: [],
  });
  setStatus(`${meta.statusNoun}を挿入しました：${name}`);
}



function openSheetInsertDialog(sheetIndex) {
  if (!state.model) return;
  const insertAt = Math.max(0, Math.min(sheetIndex, state.model.sheets.length));
  closeDataDialog({ keepState: true });
  const choices = SHEET_INSERT_TEMPLATES.map((template, index) => {
    const selected = template.type === "worksheet";
    const disabled = template.disabled ? " disabled aria-disabled=\"true\"" : "";
    const classes = `sheet-insert-template${selected ? " is-selected" : ""}${template.disabled ? " is-disabled" : ""}`;
    return `<button class="${classes}" type="button" data-data-dialog-action="sheet-insert-template" data-sheet-insert-template="${escapeAttr(template.type)}"${disabled}>
      <span class="sheet-insert-template-icon">${escapeHtml(template.icon)}</span>
      <span class="sheet-insert-template-label">${escapeHtml(template.label)}</span>
    </button>`;
  }).join("");
  const overlay = document.createElement("div");
  overlay.id = "dataDialog";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog sheet-insert-dialog" role="dialog" aria-modal="true" aria-labelledby="sheetInsertTitle" data-data-dialog="sheet-insert">
      <div class="excel-mini-titlebar" id="sheetInsertTitle">挿入</div>
      <div class="excel-mini-body data-dialog-body sheet-insert-body">
        <div class="sheet-insert-tabs" role="tablist">
          <button class="sheet-insert-tab is-active" type="button" role="tab" aria-selected="true">全般</button>
        </div>
        <div class="sheet-insert-template-grid" role="listbox" aria-label="挿入するシート">
          ${choices}
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  state.dataDialog = { type: "sheet-insert", sheetIndex: insertAt, template: "worksheet" };
  overlay.querySelector("[data-sheet-insert-template='worksheet']")?.focus({ preventScroll: true });
  setStatus("挿入するシートの種類を選択してください。");
}



function selectSheetInsertTemplate(button) {
  if (!button || button.disabled) return;
  const dialog = button.closest("[data-data-dialog='sheet-insert']");
  const template = button.dataset.sheetInsertTemplate || "worksheet";
  state.dataDialog = { ...(state.dataDialog || {}), template };
  dialog?.querySelectorAll("[data-sheet-insert-template]").forEach((item) => {
    item.classList.toggle("is-selected", item === button);
  });
}



function applySheetInsertDialog(dialog) {
  const template = state.dataDialog?.template || dialog?.querySelector(".sheet-insert-template.is-selected")?.dataset.sheetInsertTemplate || "worksheet";
  const templateInfo = SHEET_INSERT_TEMPLATES.find((item) => item.type === template) || SHEET_INSERT_TEMPLATES[0];
  if (templateInfo.disabled) {
    showMessageBox(`${templateInfo.label} は現在このブックに挿入できません。`, { title: "Cht WebSheet" });
    return;
  }
  const insertAt = Number(state.dataDialog?.sheetIndex ?? state.activeSheetIndex);
  closeDataDialog();
  if (template === "chart") {
    insertChartSheetFromSelection(recommendedChartTypeForSelection(), { insertIndex: insertAt });
    return;
  }
  if (SPECIAL_SHEET_INSERT_META[template]) {
    insertSpecialSheetAt(insertAt, template);
    return;
  }
  insertBlankSheetAt(insertAt);
}



function handleSheetInsertTemplateDoubleClick(event) {
  const button = event.currentTarget;
  if (!button || button.disabled) return;
  const dialog = button.closest("[data-data-dialog='sheet-insert']");
  if (!dialog) return;
  event.preventDefault();
  event.stopPropagation();
  selectSheetInsertTemplate(button);
  applySheetInsertDialog(dialog);
}



function deleteSheetAt(sheetIndex) {
  if (!state.model?.sheets[sheetIndex]) return;
  if (state.model.sheets.length <= 1 || visibleSheetIndexes().length <= 1) {
    setStatus("少なくとも 1 つの表示ワークシートを残す必要があります。");
    return;
  }

  const activeSheetBefore = state.activeSheetIndex;
  const [removed] = state.model.sheets.splice(sheetIndex, 1);
  state.activeSheetIndex = Math.min(sheetIndex, state.model.sheets.length - 1);
  ensureActiveSheetVisible();
  const activeSheetAfter = state.activeSheetIndex;
  clearSheetSelectionState();
  refreshWorkbookAfterSheetMetadataChange();
  pushSheetDeleteHistory(sheetIndex, removed, activeSheetBefore, activeSheetAfter, `${removed.name} ワークシート削除`);
  setStatus(`ワークシートを削除しました：${removed.name}`);
}



function sheetGroupActionIndexes(sheetIndex, { allowAllVisible = true } = {}) {
  if (!state.model?.sheets?.[sheetIndex]) return [];
  const visible = new Set(visibleSheetIndexes());
  const selected = (state.selectedSheetIndexes || [])
    .map((index) => Number(index))
    .filter((index, position, indexes) => Number.isInteger(index) && state.model?.sheets?.[index] && visible.has(index) && indexes.indexOf(index) === position)
    .sort((left, right) => left - right);
  if (selected.length > 1 && selected.includes(sheetIndex) && (allowAllVisible || selected.length < visible.size)) return selected;
  return [sheetIndex];
}



function deleteSheetsAt(sheetIndexes) {
  if (!state.model?.sheets?.length) return;
  const targets = [...new Set((sheetIndexes || []).map((index) => Number(index)).filter((index) => state.model?.sheets?.[index]))]
    .sort((left, right) => left - right);
  if (!targets.length) return;
  if (targets.length === 1) {
    deleteSheetAt(targets[0]);
    return;
  }
  const visibleTargets = targets.filter((index) => isSheetVisible(state.model.sheets[index]));
  if (visibleSheetIndexes().length <= visibleTargets.length) {
    showMessageBox("少なくとも 1 つの表示ワークシートを残す必要があります。", { title: "Cht WebSheet", tone: "warning" });
    return;
  }
  const activeSheetBefore = state.activeSheetIndex;
  const activeName = state.model.sheets[activeSheetBefore]?.name || "";
  const deleted = targets.map((index) => ({ index, sheet: cloneSheetModel(state.model.sheets[index]) }));
  targets.slice().reverse().forEach((index) => {
    state.model.sheets.splice(index, 1);
  });
  const activeByName = state.model.sheets.findIndex((sheet) => sheet.name === activeName);
  state.activeSheetIndex = activeByName >= 0 ? activeByName : Math.min(targets[0], state.model.sheets.length - 1);
  ensureActiveSheetVisible();
  const activeSheetAfter = state.activeSheetIndex;
  clearSheetSelectionState();
  refreshWorkbookAfterSheetMetadataChange();
  pushSheetDeleteManyHistory(deleted, activeSheetBefore, activeSheetAfter, `${deleted.length} シート削除`, {
    selectionAfter: null,
    selectionAfterRanges: [],
  });
  setStatus(`${deleted.length} 個のワークシートを削除しました。`);
}



function confirmDeleteSheetAt(sheetIndex) {
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) return;
  const targets = sheetGroupActionIndexes(sheetIndex);
  const visibleTargets = targets.filter((index) => isSheetVisible(state.model.sheets[index]));
  if (state.model.sheets.length <= 1 || visibleSheetIndexes().length <= visibleTargets.length) {
    showMessageBox("少なくとも 1 つの表示ワークシートを残す必要があります。", { title: "Cht WebSheet", tone: "warning" });
    return;
  }
  const message = targets.length > 1
    ? "選択した複数のシートは完全に削除されます。\n続けるには [削除] をクリックしてください。"
    : "選択したシートは完全に削除されます。\n続けるには [削除] をクリックしてください。";
  showMessageBox(message, {
    title: "Cht WebSheet",
    tone: "warning",
    buttons: [
      { label: "削除(D)", result: "delete", callback: () => deleteSheetsAt(targets) },
      { label: "キャンセル", result: "cancel" },
    ],
    cancelResult: "cancel",
  });
}



function renameSheetAt(sheetIndex, requestedName) {
  const sheet = state.model?.sheets[sheetIndex];
  if (!sheet) return;
  const validation = validateSheetName(requestedName, sheetIndex);
  if (!validation.valid) {
    setStatus(validation.message);
    showAppAlert(validation.message, "error");
    return;
  }

  const oldName = sheet.name;
  if (oldName === validation.name) {
    setStatus("ワークシート名は変更されていません。");
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  sheet.name = validation.name;
  state.activeSheetIndex = sheetIndex;
  refreshWorkbookAfterSheetMetadataChange();
  const selectionAfter = activeSelectionRange();
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore,
    selectionAfter,
    selectionBeforeRanges,
    selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
    label: `${oldName} 名前の変更`,
  });
  setStatus(`ワークシート名を変更しました：${oldName} → ${sheet.name}`);
}



function duplicateSheetAt(sheetIndex) {
  const source = state.model?.sheets[sheetIndex];
  if (!source) return;
  const activeSheetBefore = state.activeSheetIndex;
  const copy = cloneSheetModel(source);
  copy.name = uniqueCopiedSheetName(source.name);
  copy.state = "visible";
  state.model.sheets.splice(sheetIndex + 1, 0, copy);
  state.activeSheetIndex = sheetIndex + 1;
  clearSheetSelectionState();
  refreshWorkbookAfterSheetMetadataChange();
  pushSheetInsertHistory(sheetIndex + 1, copy, activeSheetBefore, sheetIndex + 1, `${source.name} ワークシートコピー`, {
    selectionAfter: null,
    selectionAfterRanges: [],
  });
  setStatus(`ワークシートをコピーしました：${copy.name}`);
}



function hideSheetAt(sheetIndex) {
  const sheet = state.model?.sheets[sheetIndex];
  if (!sheet) return;
  if (visibleSheetIndexes().length <= 1) {
    setStatus("少なくとも 1 つの表示ワークシートを残す必要があります。");
    return;
  }

  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  sheet.state = "hidden";
  state.activeSheetIndex = nearestVisibleSheetIndex(sheetIndex);
  clearSheetSelectionState();
  refreshWorkbookAfterSheetMetadataChange();
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore,
    selectionAfter: null,
    selectionBeforeRanges,
    selectionAfterRanges: [],
    label: `${sheet.name} 非表示`,
  });
  setStatus(`ワークシートを非表示にしました：${sheet.name}`);
}



function hideSheetsAt(sheetIndexes) {
  if (!state.model?.sheets?.length) return;
  const targets = [...new Set((sheetIndexes || []).map((index) => Number(index)).filter((index) => state.model?.sheets?.[index] && isSheetVisible(state.model.sheets[index])))]
    .sort((left, right) => left - right);
  if (!targets.length) return;
  if (targets.length === 1) {
    hideSheetAt(targets[0]);
    return;
  }
  if (visibleSheetIndexes().length <= targets.length) {
    showMessageBox("少なくとも 1 つの表示ワークシートを残す必要があります。", { title: "Cht WebSheet", tone: "warning" });
    return;
  }
  const sheetIndex = targets.includes(state.activeSheetIndex) ? targets[0] : state.activeSheetIndex;
  const sheetsBefore = captureSheetsForHistory(targets);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  targets.forEach((index) => {
    state.model.sheets[index].state = "hidden";
  });
  ensureActiveSheetVisible();
  clearSheetSelectionState();
  refreshWorkbookAfterSheetMetadataChange();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(targets),
    selectionBefore,
    selectionAfter: null,
    selectionBeforeRanges,
    selectionAfterRanges: [],
    label: `${targets.length} シート非表示`,
  });
  setStatus(`${targets.length} 個のワークシートを非表示にしました。`);
}



function unhideSheetFromMenu() {
  if (!state.model) return;
  const hidden = hiddenSheetIndexes();
  if (!hidden.length) {
    setStatus("再表示できるワークシートがありません。");
    return;
  }
  openUnhideSheetDialog(hidden);
}



function openUnhideSheetDialog(hidden = hiddenSheetIndexes()) {
  closeDataDialog({ keepState: true });
  const optionsHtml = hidden
    .map((index, optionIndex) => `<option value="${index}"${optionIndex === 0 ? " selected" : ""}>${escapeHtml(state.model.sheets[index].name)}</option>`)
    .join("");
  const overlay = document.createElement("div");
  overlay.id = "dataDialog";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog sheet-unhide-dialog" role="dialog" aria-modal="true" aria-labelledby="sheetUnhideTitle" data-data-dialog="sheet-unhide">
      <div class="excel-mini-titlebar" id="sheetUnhideTitle">再表示</div>
      <div class="excel-mini-body data-dialog-body sheet-unhide-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">再表示するシート(U):</span>
          <select size="${Math.min(8, Math.max(3, hidden.length))}" data-data-field="unhideSheetIndex">${optionsHtml}</select>
        </label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  state.dataDialog = { type: "sheet-unhide" };
  overlay.querySelector("[data-data-field='unhideSheetIndex']")?.focus({ preventScroll: true });
  setStatus("再表示するワークシートを選択してください。");
}



function applyUnhideSheetDialog(dialog) {
  const targetIndex = Number(dialog?.querySelector("[data-data-field='unhideSheetIndex']")?.value);
  if (!Number.isInteger(targetIndex) || !state.model?.sheets?.[targetIndex]) {
    showMessageBox("再表示するワークシートを選択してください。", { title: "Cht WebSheet" });
    return;
  }
  closeDataDialog();
  unhideSheetAt(targetIndex);
}



function handleSheetUnhideListConfirm(event) {
  if (event.type === "keydown" && (event.key !== "Enter" || isElementImeCompositionEvent(event))) return;
  const dialog = event.currentTarget.closest("[data-data-dialog='sheet-unhide']");
  if (!dialog) return;
  event.preventDefault();
  event.stopPropagation();
  applyUnhideSheetDialog(dialog);
}



function unhideSheetAt(targetIndex) {
  if (!state.model?.sheets?.[targetIndex]) return;
  const sheet = state.model.sheets[targetIndex];
  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  sheet.state = "visible";
  state.activeSheetIndex = targetIndex;
  clearSheetSelectionState();
  refreshWorkbookAfterSheetMetadataChange();
  pushHistory({
    sheetIndex: targetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore,
    selectionAfter: null,
    selectionBeforeRanges,
    selectionAfterRanges: [],
    label: `${sheet.name} 再表示`,
  });
  setStatus(`ワークシートを再表示しました：${sheet.name}`);
}



function openSheetMoveCopyDialog(sheetIndex) {
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) return;
  closeDataDialog({ keepState: true });
  const sourceIndexes = sheetGroupActionIndexes(sheetIndex);
  const workbookName = state.model.fileName || state.model.name || "作業中のブック";
  const beforeOptions = state.model.sheets
    .map((item, index) => `<option value="${index}"${index === sheetIndex + 1 ? " selected" : ""}>${escapeHtml(item.name)}</option>`)
    .join("");
  const endSelected = sheetIndex === state.model.sheets.length - 1 ? " selected" : "";
  const overlay = document.createElement("div");
  overlay.id = "dataDialog";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog sheet-move-copy-dialog" role="dialog" aria-modal="true" aria-labelledby="sheetMoveCopyTitle" data-data-dialog="sheet-move-copy">
      <div class="excel-mini-titlebar" id="sheetMoveCopyTitle">移動またはコピー</div>
      <div class="excel-mini-body data-dialog-body sheet-move-copy-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">移動先ブック名(T):</span>
          <select data-data-field="sheetMoveWorkbook" disabled>
            <option selected>${escapeHtml(workbookName)}</option>
          </select>
        </label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">挿入先(B):</span>
          <select size="${Math.min(8, Math.max(4, state.model.sheets.length + 1))}" data-data-field="sheetMoveBeforeIndex">
            ${beforeOptions}
            <option value="__end"${endSelected}>（末尾へ移動）</option>
          </select>
        </label>
        <label class="data-dialog-checkbox sheet-move-copy-checkbox">
          <input type="checkbox" data-data-field="sheetMoveCreateCopy" />
          <span>コピーを作成する(C)</span>
        </label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  state.dataDialog = { type: "sheet-move-copy", sheetIndex, sheetIndexes: sourceIndexes };
  overlay.querySelector("[data-data-field='sheetMoveBeforeIndex']")?.focus({ preventScroll: true });
  setStatus(sourceIndexes.length > 1 ? `${sourceIndexes.length} 個のシートを移動またはコピーします。` : `${sheet.name} を移動またはコピーします。`);
}



function applySheetMoveCopyDialog(dialog) {
  const sourceIndex = Number(state.dataDialog?.sheetIndex ?? state.activeSheetIndex);
  const source = state.model?.sheets?.[sourceIndex];
  if (!source) return;
  const sourceIndexes = (state.dataDialog?.sheetIndexes?.length ? state.dataDialog.sheetIndexes : [sourceIndex])
    .map((index) => Number(index))
    .filter((index, position, indexes) => Number.isInteger(index) && state.model?.sheets?.[index] && indexes.indexOf(index) === position)
    .sort((left, right) => left - right);
  const beforeValue = dialog?.querySelector("[data-data-field='sheetMoveBeforeIndex']")?.value ?? "__end";
  const beforeIndex = beforeValue === "__end" ? state.model.sheets.length : Number(beforeValue);
  const createCopy = Boolean(dialog?.querySelector("[data-data-field='sheetMoveCreateCopy']")?.checked);
  closeDataDialog();
  if (createCopy) {
    copySheetsToPosition(sourceIndexes, beforeIndex);
  } else {
    moveSheetsToBeforeIndex(sourceIndexes, beforeIndex);
  }
}



function handleSheetMoveCopyListConfirm(event) {
  if (event.type === "keydown" && (event.key !== "Enter" || isElementImeCompositionEvent(event))) return;
  const dialog = event.currentTarget.closest("[data-data-dialog='sheet-move-copy']");
  if (!dialog) return;
  event.preventDefault();
  event.stopPropagation();
  applySheetMoveCopyDialog(dialog);
}



function copySheetToPosition(sourceIndex, beforeIndex) {
  const source = state.model?.sheets?.[sourceIndex];
  if (!source) return;
  const activeSheetBefore = state.activeSheetIndex;
  const copy = cloneSheetModel(source);
  copy.name = uniqueCopiedSheetName(source.name);
  copy.state = "visible";
  const insertAt = Math.max(0, Math.min(Number(beforeIndex) || 0, state.model.sheets.length));
  state.model.sheets.splice(insertAt, 0, copy);
  state.activeSheetIndex = insertAt;
  clearSheetSelectionState();
  refreshWorkbookAfterSheetMetadataChange();
  pushSheetInsertHistory(insertAt, copy, activeSheetBefore, insertAt, `${source.name} ワークシートコピー`, {
    selectionAfter: null,
    selectionAfterRanges: [],
  });
  setStatus(`ワークシートをコピーしました：${copy.name}`);
}



function copySheetsToPosition(sourceIndexes, beforeIndex) {
  const indexes = [...new Set((sourceIndexes || []).map((index) => Number(index)).filter((index) => state.model?.sheets?.[index]))]
    .sort((left, right) => left - right);
  if (!indexes.length) return;
  if (indexes.length === 1) {
    copySheetToPosition(indexes[0], beforeIndex);
    return;
  }
  const activeSheetBefore = state.activeSheetIndex;
  const insertAt = Math.max(0, Math.min(Number(beforeIndex) || 0, state.model.sheets.length));
  const existingNames = new Set(state.model.sheets.map((sheet) => sheet.name));
  const copies = indexes.map((index) => {
    const source = state.model.sheets[index];
    const copy = cloneSheetModel(source);
    copy.name = uniqueCopiedSheetNameFromSet(source.name, existingNames);
    existingNames.add(copy.name);
    copy.state = "visible";
    return copy;
  });
  state.model.sheets.splice(insertAt, 0, ...copies);
  state.activeSheetIndex = insertAt;
  clearSheetSelectionState();
  state.selectedSheetIndexes = copies.map((_sheet, offset) => insertAt + offset);
  refreshWorkbookAfterSheetMetadataChange({ preserveSheetSelection: true });
  pushSheetInsertManyHistory(insertAt, copies, activeSheetBefore, insertAt, `${copies.length} シートコピー`, {
    selectedSheetNamesAfter: copies.map((sheet) => sheet.name),
    selectionAfter: null,
    selectionAfterRanges: [],
  });
  setStatus(`${copies.length} 個のワークシートをコピーしました。`);
}



function moveSheetToBeforeIndex(sourceIndex, beforeIndex) {
  const source = state.model?.sheets?.[sourceIndex];
  if (!source) return;
  const sheetCount = state.model.sheets.length;
  const normalizedBefore = Math.max(0, Math.min(Number(beforeIndex), sheetCount));
  let targetIndex = normalizedBefore;
  if (normalizedBefore > sourceIndex) targetIndex -= 1;
  targetIndex = Math.max(0, Math.min(targetIndex, sheetCount - 1));
  if (targetIndex === sourceIndex) {
    state.activeSheetIndex = sourceIndex;
    setStatus("ワークシートは移動されていません。");
    return;
  }
  const activeSheetBefore = state.activeSheetIndex;
  const movedName = source.name;
  const actualIndex = moveSheetToIndex(sourceIndex, targetIndex);
  state.activeSheetIndex = actualIndex;
  clearSheetSelectionState();
  refreshWorkbookAfterSheetMetadataChange();
  pushSheetMoveHistory(sourceIndex, actualIndex, activeSheetBefore, actualIndex, `${movedName} ワークシート移動`);
  setStatus(`ワークシートを移動しました：${movedName}`);
}



function moveSheetsToBeforeIndex(sourceIndexes, beforeIndex) {
  const indexes = [...new Set((sourceIndexes || []).map((index) => Number(index)).filter((index) => state.model?.sheets?.[index]))]
    .sort((left, right) => left - right);
  if (!indexes.length) return;
  if (indexes.length === 1) {
    moveSheetToBeforeIndex(indexes[0], beforeIndex);
    return;
  }
  const sheetsBefore = state.model.sheets.map((sheet) => sheet.name);
  const activeSheetBefore = state.activeSheetIndex;
  const movingSet = new Set(indexes);
  const movingSheets = indexes.map((index) => state.model.sheets[index]);
  const selectedNames = movingSheets.map((sheet) => sheet.name);
  const rawBefore = Number(beforeIndex);
  const targetBefore = Number.isFinite(rawBefore) ? Math.max(0, Math.min(rawBefore, state.model.sheets.length)) : state.model.sheets.length;
  const remaining = state.model.sheets.filter((_sheet, index) => !movingSet.has(index));
  const insertAt = targetBefore >= state.model.sheets.length
    ? remaining.length
    : state.model.sheets.slice(0, targetBefore).filter((_sheet, index) => !movingSet.has(index)).length;
  const reordered = [
    ...remaining.slice(0, insertAt),
    ...movingSheets,
    ...remaining.slice(insertAt),
  ];
  const sheetsAfter = reordered.map((sheet) => sheet.name);
  if (sheetsBefore.join("\u0000") === sheetsAfter.join("\u0000")) {
    setStatus("ワークシートは移動されていません。");
    return;
  }
  const activeName = state.model.sheets[activeSheetBefore]?.name || selectedNames[0];
  state.model.sheets = reordered;
  state.activeSheetIndex = Math.max(0, state.model.sheets.findIndex((sheet) => sheet.name === activeName));
  if (state.activeSheetIndex < 0) state.activeSheetIndex = insertAt;
  clearSheetSelectionState();
  state.selectedSheetIndexes = selectedNames
    .map((name) => state.model.sheets.findIndex((sheet) => sheet.name === name))
    .filter((index) => index >= 0);
  if (!state.selectedSheetIndexes.includes(state.activeSheetIndex)) state.activeSheetIndex = state.selectedSheetIndexes[0] ?? state.activeSheetIndex;
  refreshWorkbookAfterSheetMetadataChange({ preserveSheetSelection: true });
  pushSheetMoveManyHistory(sheetsBefore, sheetsAfter, selectedNames, activeSheetBefore, state.activeSheetIndex, `${selectedNames.length} シート移動`);
  setStatus(`${selectedNames.length} 個のワークシートを移動しました。`);
}

export {
  addBlankSheet,
  addBorderModel,
  applyDefinedNamesToFormulaEngine,
  applyFormulaEngineIntrinsicNames,
  applySheetInsertDialog,
  applySheetMoveCopyDialog,
  applyTint,
  applyUnhideSheetDialog,
  beginSheetTabInlineRename,
  borderStyleToCss,
  borderWeight,
  bordersToModel,
  buildFormulaEngine,
  cachedFormulaEngineValueForCell,
  cancelSheetTabDragAutoScroll,
  cancelSheetTabInlineRename,
  cleanStructuredReferencePart,
  colorToCss,
  columnHasStructuralDimension,
  commitSheetTabInlineRename,
  commitSheetTabRenameOnOutsidePointer,
  confirmDeleteSheetAt,
  consumeFormulaDoubleQuotedText,
  consumeFormulaSingleQuotedSheetName,
  copySheetToPosition,
  copySheetsToPosition,
  definedNameEngineKey,
  definedNameExpressionForFormulaEngine,
  definedNameReplacementAt,
  deleteSheetAt,
  deleteSheetsAt,
  denseFormulaDataForSheet,
  denseFormulaSheetsForWorkbook,
  detachFormulaReferenceCellEditor,
  duplicateSheetAt,
  ensureActiveSheetVisible,
  explicitCellFontColor,
  extractMerges,
  fillColorToCss,
  findEdgeBorder,
  findWorksheetDimensions,
  finishSheetTabDrag,
  focusSheetTabByIndex,
  focusSheetTabRenameInput,
  fontFamilyForExcel,
  formulaCellRawValue,
  formulaEngineCellEntry,
  formulaEngineContentForCurrentModel,
  formulaEngineContextCellRange,
  formulaEngineDefinedNameAliases,
  formulaEngineFindClosingParen,
  formulaEngineLastSheetBangIndex,
  formulaEngineReferenceSheetExists,
  formulaEngineReferenceSheetName,
  formulaEngineSheetIndex,
  formulaEngineStaticA1ReferenceRange,
  formulaEngineStaticCellInfoReplacement,
  formulaEngineStaticStringLiteral,
  formulaEngineStringLiteral,
  formulaEngineStripOuterParentheses,
  formulaEngineValueForCell,
  formulaReferenceSheetNavigationContext,
  formulaSheetQualifierBeforeIndex,
  generatedFormulaEngineDefinedName,
  handleFormulaReferenceSheetTabPointerDown,
  handleSheetInsertTemplateDoubleClick,
  handleSheetMoveCopyListConfirm,
  handleSheetTabClick,
  handleSheetTabDoubleClick,
  handleSheetTabDragMove,
  handleSheetTabKeydown,
  handleSheetTabRenameInput,
  handleSheetTabRenameInputBlur,
  handleSheetTabRenameInputKeydown,
  handleSheetUnhideListConfirm,
  hexToRgb,
  hiddenSheetIndexes,
  hideSheetAt,
  hideSheetsAt,
  hyperFormulaBuildConfig,
  hyperFormulaSafeDefinedName,
  importedColumnDimensionExtent,
  insertBlankSheetAt,
  insertSpecialSheetAt,
  isDefinedNameTokenBoundary,
  isDefinedNameTokenChar,
  isFormulaEngineArrayResizeError,
  isFormulaReferenceBoundary,
  isFormulaReferenceTokenChar,
  isSheetVisible,
  markSheetTabDragSources,
  mergedEdgeBorders,
  migrateDenseFormulaDataToCells,
  moveSheetToBeforeIndex,
  moveSheetsToBeforeIndex,
  normalizeCellTextRotation,
  normalizeDefinedNameLookup,
  normalizeFormulaEngineFunctionName,
  normalizeFormulaSheetName,
  openSheetInsertDialog,
  openSheetMoveCopyDialog,
  openUnhideSheetDialog,
  parseStructuredReferenceSpec,
  parseThemeColor,
  parseThemeColors,
  qualifySheetLocalDefinedNameExpression,
  quoteFont,
  removeSheetTabDropIndicator,
  renameSheetAt,
  renderSheetTabs,
  renderWorkbook,
  resizeSheetTabRenameInput,
  restoreFormulaReferenceNavigationEditor,
  rewriteStaticCellInfoFormulaReferencesForFormulaEngine,
  rgbToHex,
  runSheetTabDragAutoScroll,
  scheduleFocusSheetTabByIndex,
  scheduleSheetTabKeyboardActivationStabilizer,
  selectSheetInsertTemplate,
  selectSheetTabRange,
  setActiveSheet,
  sheetGroupActionIndexes,
  sheetHfDataDimensions,
  sheetTabButtonByIndex,
  sheetTabButtons,
  sheetTabDropTargetFromClientX,
  shouldContinueSheetTabFocus,
  shouldIgnoreTailColumnInstance,
  shouldUseCachedWebSheetSpillFormulaForEngine,
  splitStructuredReferenceParts,
  stabilizeSheetTabKeyboardActivation,
  startSheetTabDrag,
  stopSheetTabRenameInputPropagation,
  structuredReferenceA1,
  structuredReferenceColumnOffset,
  structuredReferenceColumnRangePart,
  structuredReferenceColumns,
  structuredReferenceCompositeRows,
  structuredReferenceContextTable,
  structuredReferenceCurrentRowColumnRangePart,
  structuredReferenceNameMatches,
  structuredReferencePartHasEscapedSpecial,
  structuredReferenceResolvedRange,
  structuredReferenceRows,
  structuredReferenceToA1,
  structuredReferenceTokenAt,
  structuredReferenceWrappedCurrentRowColumnPart,
  styleToCss,
  switchActiveSheetForFormulaReference,
  tableColumnNamesForStructuredRef,
  toggleSheetTabGroupSelection,
  translateDefinedNamesForFormulaEngine,
  translateFormulaForCurrentFormulaEngine,
  translateFormulaForFormulaEngine,
  translateStructuredReferencesForFormulaEngine,
  unescapeStructuredReferencePart,
  unhideSheetAt,
  unhideSheetFromMenu,
  uniqueFontFamilies,
  unquoteSheetName,
  updateSheetTabDragAutoScroll,
  updateSheetTabDropIndicator,
  validationDefinedName,
  validationDefinedNameVisitKey,
  verticalAlignToFlex,
  visibleSheetIndexes,
  webSheetFormulaTopLevelFunctionName,
  workbookStructuredTableRefs,
  workbookTabSelectedSheetIndexes,
  worksheetExplicitDimensionExtents,
  worksheetUsedDimensionExtents,
};
