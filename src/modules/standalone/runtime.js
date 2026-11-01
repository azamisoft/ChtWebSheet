// Feature module; shared bindings are imported explicitly.
import {
  $,
  $zoomLabel,
  $zoomSlider,
  CELL_EDIT_MIN_PADDING_TOP,
  CELL_EDIT_PADDING_BOTTOM,
  CELL_EDIT_TRAILING_LINE_SENTINEL,
  HyperFormula,
  state,
} from "../../app-core.js";
import {
  accessKeyFromKeyboardEvent,
  chartSvg,
} from "../drawing/charts-and-insertion.js";
import {
  insertLineBreakInTextControl,
  isAltModifierEvent,
  isEnterKeyEvent,
} from "../editing/autofill.js";
import {
  svgDataUrl,
} from "../editing/ink-and-clipboard.js";
import {
  isElementImeCompositionEvent,
  isImeCompositionEvent,
  keyboardDateTimeShortcutKind,
  keyboardDateTimeShortcutText,
  markElementImeComposition,
  markImeComposition,
  nowMs,
  pad2,
} from "../editing/keyboard-and-paste.js";
import {
  beginCellEdit,
  canvasAlignedTextLeft,
  canvasDevicePixelRatio,
  canvasFont,
  canvasFontDisplayScale,
  canvasFontSize,
  canvasLayerBounds,
  canvasLineBaseline,
  canvasTextLineHeight,
  cellAnnotationMarkersHtml,
  cellAnnotationTitle,
  cellContentStyle,
  cellHasComment,
  cellHasFill,
  cellHasNote,
  cellInnerHtml,
  coordinateToPixels,
  editableCellOverlayHeight,
  editableCellOverlayMetrics,
  editableCellPaddingTop,
  editableCellVerticalAlignment,
  editablePlainText,
  gridPositionHasFill,
  imageBoxToPixels,
  insertLineBreakInEditable,
  overlayCellCss,
  paintCanvasPhonetics,
  paintCanvasStrikethrough,
  paintCanvasUnderline,
  phoneticRunsForLine,
  placeCaretAtEnd,
  pointIsInsideElementRect,
  renderImageLayer,
  rtlColumnCoordinateToPixels,
  sheetColumnsRightToLeft,
  shouldHideGridBottom,
  shouldHideGridRight,
  stripCellEditSentinels,
  syncEditableCellAlignment,
  visualGridColumnStart,
  wrapCanvasText,
} from "../editing/selection-and-cell-text.js";
import {
  cellEffectiveCss,
} from "../editing/structure-and-history.js";
import {
  appendRichChar,
  canvasRichLineMetrics,
  canvasTextPaintRange,
  cellShowsValidationDropdown,
  crispCanvasLine,
  drawCanvasRichLine,
  measureRichLine,
  normalizeRichTextLineEndings,
  paintCanvasCell,
  paintCanvasCellText,
  paintCanvasCellTextLayout,
  paintCanvasCells,
  paintCanvasGridLines,
  paintCanvasRichCellText,
  paintCanvasVerticalCellText,
  paintSheetCanvas,
  rangeIntersects,
  rangePixelRect,
  renderSheet,
  richLinePlainText,
  richSegmentCss,
  richTextRunsForDisplay,
  splitRichRunsIntoHardLines,
  wrapCanvasRichText,
} from "../rendering/grid-and-canvas.js";
import {
  cellHasPhonetics,
  formatDecimalPlaces,
  normalizeCommentText,
} from "../ui/annotations-and-cell-format.js";
import {
  accessKeyLabelHtml,
  cellKey,
  columnName,
  currencySymbol,
  dateDetectionFormat,
  dateToExcelSerial,
  decodeCell,
  decodeRange,
  escapeAttr,
  escapeHtml,
  excelSerialToDate,
  isDateFormat,
  styleObjectToString,
} from "../ui/help-and-utilities.js";
import {
  rowHeaderWidthForSheet,
} from "../ui/ribbon-and-links.js";
import {
  applyZoomDialog,
  closeZoomDialog,
  closeZoomDialogOnEscape,
  closeZoomDialogOnOutsidePointer,
  createBlankSheet,
  displayScaleForZoom,
  handleZoomDialogAction,
  handleZoomDialogKeydown,
  handleZoomDialogOptionChange,
  isWorkbookWheelTarget,
  openZoomDialog,
  positionZoomDialog,
  scaleSheetDisplay,
  selectCustomZoomOption,
  sheetBaseZoomScale,
  uniqueSheetName,
  updateOpenZoomDialogValue,
  zoomSelectionToFit,
} from "../ui/sheet-protection-and-translation.js";
import {
  applyDefinedNamesToFormulaEngine,
  definedNameExpressionForFormulaEngine,
  denseFormulaDataForSheet,
  denseFormulaSheetsForWorkbook,
  formulaCellRawValue,
  formulaEngineValueForCell,
  hyperFormulaBuildConfig,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  addBorderLine,
  borderLineHtml,
  borderLinesForSheet,
  buildMergeMaps,
  canvasBorderDash,
  cellDisplayText,
  columnBoundaryPositions,
  compareBorder,
  dateFormatTokenIsMinute,
  displayValue,
  formatDateForFormat,
  formatElapsedTimeForFormat,
  formatNumberForFormat,
  getCellFormula,
  getSheetId,
  gridTrackStyle,
  implicitCellDisplayText,
  isBlockingCell,
  isElapsedTimeFormat,
  isOccupiedCell,
  numberFormatCanApplyToNumericString,
  numericStringValueForNumberFormat,
  paintCanvasBorders,
  primaryExcelFormatSection,
  rowBoundaryPositions,
  sheetGridClassName,
  shouldHideZeroValueForSheet,
  textOverflowWidth,
  updateHorizontalScrollbar,
  visibleColumnRange,
  visibleColumnRangeRightToLeft,
  visibleGridRange,
  visibleRowRange,
  widthBetweenColumns,
} from "../workbook/saving-and-export.js";
import {
  normalizePhoneticRuns,
  parseValidationOptions,
} from "../workbook/tables-and-validation.js";
import {
  menuAccessKeyFromLabel,
} from "../workbook/views-and-macros.js";


function standaloneRuntime() {
  const model = window.__WEBSHEET_MODEL__;
  const HF = window.HyperFormula?.HyperFormula || window.HyperFormula;
  const licenseKey = "gpl-v3";
  function hyperFormulaBuildConfig() {
    return {
      licenseKey,
      leapYear1900: false,
      useArrayArithmetic: true,
      maxRows: 1048576,
      maxColumns: 16384,
    };
  }
  const state = {
    active: 0,
    selected: null,
    selectedRange: null,
    selectedRanges: [],
    showFormulas: false,
    cellEditComposing: false,
    formulaComposing: false,
    editDrag: null,
    suppressCellSelectionUntil: 0,
    validationDropdownPopup: null,
  };
  const CELL_EDIT_TRAILING_LINE_SENTINEL = "\u200b";
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
  const WEBSHEET_STANDALONE_CACHEABLE_SPILL_NAMES = new Set([
    "UNIQUE",
    "SORT",
    "SORTBY",
    "FILTER",
    "SEQUENCE",
    "INDIRECT",
    "OFFSET",
    "INDEX",
    "TAKE",
    "DROP",
    "VSTACK",
    "HSTACK",
    "CHOOSEROWS",
    "CHOOSECOLS",
    "TOCOL",
    "TOROW",
    "WRAPROWS",
    "WRAPCOLS",
    "EXPAND",
    "TRANSPOSE",
    "CHOOSE",
    "TEXTSPLIT",
    "MAKEARRAY",
    "BYROW",
    "BYCOL",
    "MAP",
    "REDUCE",
    "SCAN",
  ]);
  function standaloneFormulaTopLevelFunctionName(formula) {
    let text = String(formula || "").trim().replace(/^=/, "").trim();
    const match = text.match(/^([A-Za-z_][A-Za-z0-9_.]*)\s*\(/);
    text = String(match?.[1] || "").trim().toUpperCase();
    while (text.startsWith("_XLFN.") || text.startsWith("_XLWS.")) {
      text = text.replace(/^_(?:XLFN|XLWS)\./, "");
    }
    return text;
  }
  function cachedStandaloneFormulaEngineValueForCell(cell) {
    if (cell?.cached != null) return formulaEngineValueForCell(cell, cell.cached);
    if (cell?.display != null) return formulaEngineValueForCell(cell, cell.display);
    return null;
  }
  function shouldUseCachedStandaloneSpillFormulaForEngine(cell, raw, options) {
    if (!options?.useCachedWebSheetDynamicArrays || typeof raw !== "string" || !raw.startsWith("=")) return false;
    return WEBSHEET_STANDALONE_CACHEABLE_SPILL_NAMES.has(standaloneFormulaTopLevelFunctionName(raw));
  }
  function standaloneFormulaEngineCellEntry(sheet, cell, options) {
    const row = Number(cell?.row);
    const col = Number(cell?.col);
    if (!Number.isInteger(row) || !Number.isInteger(col) || row < 1 || col < 1) return null;
    let raw = formulaCellRawValue(cell);
    if (options?.useCachedFormulas && typeof raw === "string" && raw.startsWith("=")) {
      raw = cell.cached ?? null;
    } else if (shouldUseCachedStandaloneSpillFormulaForEngine(cell, raw, options)) {
      raw = cachedStandaloneFormulaEngineValueForCell(cell);
    }
    if (raw == null || raw === "") return null;
    const sheetRows = Math.max(1, Number(sheet?.rowCount) || sheet?.rows?.length || row);
    const sheetCols = Math.max(1, Number(sheet?.colCount) || sheet?.columns?.length || col);
    if (row > sheetRows || col > sheetCols) return null;
    return { row, col, raw, cell };
  }
  function denseFormulaDataForSheet(sheet, options) {
    const entries = [];
    let rowCount = 1;
    let colCount = 1;
    Object.keys(sheet?.cells || {}).forEach(function (key) {
      const entry = standaloneFormulaEngineCellEntry(sheet, sheet.cells[key], options);
      if (!entry) return;
      entries.push(entry);
      rowCount = Math.max(rowCount, entry.row);
      colCount = Math.max(colCount, entry.col);
    });
    const data = Array.from({ length: rowCount }, function () {
      return Array.from({ length: colCount }, function () {
        return null;
      });
    });
    entries.forEach(function (entry) {
      const row = entry.row;
      const col = entry.col;
      let raw = entry.raw;
      if (typeof raw === "string" && raw.startsWith("=")) {
        raw = translateStandaloneFormulaForFormulaEngine(raw, {
          definedNameAliases: options?.definedNameAliases,
          sheetIndex: options?.sheetIndex,
          row,
          col,
        });
      }
      data[row - 1][col - 1] = raw ?? null;
    });
    return data;
  }
  function denseFormulaSheetsForWorkbook(options) {
    const sheets = {};
    model.sheets.forEach(function (sheet, sheetIndex) {
      sheets[sheet.name] = denseFormulaDataForSheet(sheet, Object.assign({}, options || {}, { sheetIndex }));
    });
    return sheets;
  }
  const definedNameAliases = standaloneFormulaEngineDefinedNameAliases(model);
  const sourceSheets = denseFormulaSheetsForWorkbook({ definedNameAliases });
  let hf;
  try {
    hf = HF.buildFromSheets(sourceSheets, hyperFormulaBuildConfig());
    applyDefinedNamesToFormulaEngine(hf, definedNameAliases);
  } catch (error) {
    if (/resizing to smaller array/i.test(String(error?.message || error || ""))) {
      try {
        const dynamicArraySheets = denseFormulaSheetsForWorkbook({ useCachedWebSheetDynamicArrays: true, definedNameAliases });
        hf = HF.buildFromSheets(dynamicArraySheets, hyperFormulaBuildConfig());
        applyDefinedNamesToFormulaEngine(hf, definedNameAliases);
      } catch {}
    }
    if (!hf) {
      const fallbackSheets = denseFormulaSheetsForWorkbook({ useCachedFormulas: true, definedNameAliases });
      hf = HF.buildFromSheets(fallbackSheets, hyperFormulaBuildConfig());
      applyDefinedNamesToFormulaEngine(hf, definedNameAliases);
    }
  }
  const $tabs = $("#wx-tabs");
  const $sheet = $("#wx-sheet");
  const $formula = $("#wx-formula-input");
  const $label = $("#wx-cell-label");
  const $apply = $("#wx-apply");
  const $status = $("#wx-status");
  const $toggle = $("#wx-toggle");
  const $footer = $("#wx-footer");
  const $hscroll = $("#wx-hscroll");
  const $hspacer = $("#wx-hspacer");
  const $add = $("#wx-add");
  const $splitter = $("#wx-splitter");
  const $zoomSlider = $("#wx-zoom-slider");
  const $zoomLabel = $("#wx-zoom-label");
  const $zoomOut = $("#wx-zoom-out");
  const $zoomIn = $("#wx-zoom-in");
  let syncingScroll = false;
  let canvasPaintFrame = 0;
  const standaloneMergeMapsCacheByMerges = new WeakMap();

  function applyDefinedNamesToFormulaEngine(engine, aliases) {
    (model.definedNames || []).forEach(function (definition) {
      if (!definition?.name || definition.hidden) return;
      const engineName = aliases.byKey.get(standaloneDefinedNameEngineKey(definition)) || definition.name;
      const expression = definedNameExpressionForFormulaEngine(definition.refersTo, aliases, definition.sheetIndex);
      if (!expression) return;
      const scope = undefined;
      try {
        engine.addNamedExpression(engineName, expression, scope);
      } catch {}
    });
  }

  function definedNameExpressionForFormulaEngine(refersTo, aliases, sheetIndex) {
    const text = String(refersTo || "").trim();
    if (!text) return "";
    const expression = standaloneQualifySheetLocalDefinedNameExpression(text.startsWith("=") ? text : "=" + text, sheetIndex);
    const structured = translateStandaloneStructuredReferencesForFormulaEngine(expression, sheetIndex, {});
    return translateStandaloneDefinedNamesForFormulaEngine(structured, aliases, sheetIndex);
  }

  function standaloneFormulaEngineDefinedNameAliases(workbookModel) {
    const aliases = { byKey: new Map(), byName: new Map(), bySheetName: new Map(), bySheetDisplayName: new Map() };
    const usedGeneratedAliases = new Set();
    (workbookModel?.definedNames || []).forEach(function (definition, index) {
      if (!definition?.name || definition.hidden) return;
      const engineName = definition.scope === "sheet"
        ? standaloneGeneratedFormulaEngineDefinedName(index, usedGeneratedAliases)
        : standaloneHyperFormulaSafeDefinedName(definition.name, index, usedGeneratedAliases);
      usedGeneratedAliases.add(engineName.toLowerCase());
      aliases.byKey.set(standaloneDefinedNameEngineKey(definition), engineName);
      if (definition.scope === "sheet") {
        const sheetIndex = Number(definition.sheetIndex);
        if (!aliases.bySheetName.has(sheetIndex)) aliases.bySheetName.set(sheetIndex, new Map());
        aliases.bySheetName.get(sheetIndex).set(definition.name, engineName);
        const sheetName = workbookModel?.sheets?.[sheetIndex]?.name;
        if (sheetName) {
          const normalizedSheetName = standaloneNormalizeFormulaSheetName(sheetName);
          if (!aliases.bySheetDisplayName.has(normalizedSheetName)) aliases.bySheetDisplayName.set(normalizedSheetName, new Map());
          aliases.bySheetDisplayName.get(normalizedSheetName).set(definition.name, engineName);
        }
      } else if (!aliases.byName.has(definition.name)) {
        aliases.byName.set(definition.name, engineName);
      }
    });
    return aliases;
  }

  function standaloneDefinedNameEngineKey(definition) {
    return `${definition?.scope || "workbook"}:${definition?.sheetIndex ?? ""}:${definition?.name || ""}`;
  }

  function standaloneHyperFormulaSafeDefinedName(name, index, usedGeneratedAliases) {
    const text = String(name || "");
    if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(text) && !standaloneIsCellReference(text)) return text;
    return standaloneGeneratedFormulaEngineDefinedName(index, usedGeneratedAliases);
  }

  function standaloneGeneratedFormulaEngineDefinedName(index, usedGeneratedAliases) {
    let alias = "__WEBSHEET_NAME_" + (index + 1);
    while (usedGeneratedAliases.has(alias.toLowerCase())) {
      alias = "__WEBSHEET_NAME_" + (index + 1) + "_" + (usedGeneratedAliases.size + 1);
    }
    usedGeneratedAliases.add(alias.toLowerCase());
    return alias;
  }

  function translateStandaloneDefinedNamesForFormulaEngine(formula, aliases, sheetIndex) {
    if (typeof formula !== "string" || !formula.startsWith("=")) return formula;
    const names = new Map(aliases?.byName || []);
    const localNames = aliases?.bySheetName?.get(Number(sheetIndex));
    if (localNames?.size) {
      localNames.forEach(function (alias, name) {
        names.set(name, alias);
      });
    }
    if (!names.size) return formula;
    const replacements = Array.from(names.entries())
      .filter(function (entry) {
        return entry[0] && entry[1] && entry[0] !== entry[1];
      })
      .sort(function (left, right) {
        return right[0].length - left[0].length;
      });
    if (!replacements.length) return formula;

    let output = "";
    let index = 0;
    while (index < formula.length) {
      const char = formula[index];
      if (char === '"') {
        const end = standaloneConsumeFormulaDoubleQuotedText(formula, index);
        output += formula.slice(index, end);
        index = end;
        continue;
      }
      if (char === "'") {
        const end = standaloneConsumeFormulaSingleQuotedSheetName(formula, index);
        output += formula.slice(index, end);
        index = end;
        continue;
      }
      const qualifier = standaloneFormulaSheetQualifierBeforeIndex(formula, index);
      if (qualifier) {
        const qualifiedNames = aliases?.bySheetDisplayName?.get(standaloneNormalizeFormulaSheetName(qualifier.sheetName));
        const qualifiedReplacement = standaloneDefinedNameReplacementAt(formula, index, qualifiedNames);
        if (qualifiedReplacement) {
          output = output.slice(0, Math.max(0, output.length - (index - qualifier.start)));
          output += qualifiedReplacement[1];
          index += qualifiedReplacement[0].length;
          continue;
        }
      }
      const replacement = replacements.find(function (entry) {
        const name = entry[0];
        return (
          formula.startsWith(name, index) &&
          standaloneIsDefinedNameTokenBoundary(formula, index, name.length) &&
          formula[index + name.length] !== "!"
        );
      });
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

  function standaloneDefinedNameReplacementAt(formula, index, names) {
    if (!names?.size) return null;
    return Array.from(names.entries())
      .filter(function (entry) {
        return entry[0] && entry[1];
      })
      .sort(function (left, right) {
        return right[0].length - left[0].length;
      })
      .find(function (entry) {
        const name = entry[0];
        return (
          formula.startsWith(name, index) &&
          standaloneIsDefinedNameTokenBoundary(formula, index, name.length) &&
          formula[index + name.length] !== "!"
        );
      }) || null;
  }

  function standaloneFormulaSheetQualifierBeforeIndex(formula, index) {
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
    while (start >= 0 && standaloneIsDefinedNameTokenChar(formula[start])) start -= 1;
    const sheetName = formula.slice(start + 1, end);
    return sheetName ? { start: start + 1, sheetName } : null;
  }

  function standaloneNormalizeFormulaSheetName(name) {
    return String(name || "").toLocaleLowerCase("ja-JP");
  }

  function standaloneQualifySheetLocalDefinedNameExpression(expression, sheetIndex) {
    if (sheetIndex == null) return expression;
    const sheetName = model?.sheets?.[Number(sheetIndex)]?.name;
    if (!sheetName || typeof expression !== "string" || !expression.startsWith("=")) return expression;
    const qualifier = standaloneQuoteSheetNameForFormula(sheetName) + "!";
    let output = "";
    let index = 0;
    while (index < expression.length) {
      const char = expression[index];
      if (char === '"') {
        const end = standaloneConsumeFormulaDoubleQuotedText(expression, index);
        output += expression.slice(index, end);
        index = end;
        continue;
      }
      if (char === "'") {
        const end = standaloneConsumeFormulaSingleQuotedSheetName(expression, index);
        output += expression.slice(index, end);
        index = end;
        continue;
      }
      const ref = expression.slice(index).match(/^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}/)?.[0];
      if (ref && standaloneIsCellReference(ref) && standaloneIsFormulaReferenceBoundary(expression, index, ref.length) && expression[index - 1] !== "!") {
        output += qualifier + ref;
        index += ref.length;
        continue;
      }
      output += char;
      index += 1;
    }
    return output;
  }

  function standaloneQuoteSheetNameForFormula(name) {
    const text = String(name || "");
    return /^[A-Za-z_][A-Za-z0-9_]*$/.test(text) ? text : "'" + text.replace(/'/g, "''") + "'";
  }

  function standaloneIsFormulaReferenceBoundary(formula, start, length) {
    return !standaloneIsFormulaReferenceTokenChar(formula[start - 1]) && !standaloneIsFormulaReferenceTokenChar(formula[start + length]);
  }

  function standaloneIsFormulaReferenceTokenChar(char) {
    return char === "$" || standaloneIsDefinedNameTokenChar(char);
  }

  function standaloneConsumeFormulaDoubleQuotedText(formula, start) {
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

  function standaloneConsumeFormulaSingleQuotedSheetName(formula, start) {
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

  function standaloneIsDefinedNameTokenBoundary(formula, start, length) {
    return !standaloneIsDefinedNameTokenChar(formula[start - 1]) && !standaloneIsDefinedNameTokenChar(formula[start + length]);
  }

  function standaloneIsDefinedNameTokenChar(char) {
    return Boolean(char && /[A-Za-z0-9_.\\\u0080-\uffff]/.test(char));
  }

  function standaloneRewriteStaticCellInfoFormulaReferencesForFormulaEngine(formula, context) {
    if (typeof formula !== "string" || !formula.startsWith("=") || !/cell\s*\(/i.test(formula)) return formula;
    let output = "";
    let index = 0;
    while (index < formula.length) {
      const char = formula[index];
      if (char === '"') {
        const end = standaloneConsumeFormulaDoubleQuotedText(formula, index);
        output += formula.slice(index, end);
        index = end;
        continue;
      }
      if (char === "'") {
        const end = standaloneConsumeFormulaSingleQuotedSheetName(formula, index);
        output += formula.slice(index, end);
        index = end;
        continue;
      }
      const match = formula.slice(index).match(/^(?:_xlfn\.)?cell\s*\(/i);
      if (match && !standaloneIsDefinedNameTokenChar(formula[index - 1])) {
        const openIndex = index + match[0].lastIndexOf("(");
        const closeIndex = standaloneFindFormulaClosingParen(formula, openIndex);
        if (closeIndex > openIndex) {
          const replacement = standaloneStaticCellInfoFormulaReplacement(formula.slice(openIndex + 1, closeIndex), context);
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

  function standaloneFindFormulaClosingParen(formula, openIndex) {
    let depth = 0;
    for (let index = openIndex; index < formula.length; index += 1) {
      const char = formula[index];
      if (char === '"') {
        index = standaloneConsumeFormulaDoubleQuotedText(formula, index) - 1;
        continue;
      }
      if (char === "'") {
        index = standaloneConsumeFormulaSingleQuotedSheetName(formula, index) - 1;
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

  function standaloneStaticCellInfoFormulaReplacement(argsSource, context) {
    const args = standaloneSplitFormulaArguments(argsSource);
    if (args.length < 1 || args.length > 2) return null;
    const infoType = standaloneFormulaStaticStringLiteral(args[0]);
    if (!["address", "col", "row"].includes(infoType)) return null;
    const range = args.length > 1
      ? standaloneFormulaStaticA1ReferenceRange(args[1])
      : standaloneFormulaContextCellRange(context);
    if (!range) return null;
    if (infoType === "row") return String(range.top);
    if (infoType === "col") return String(range.left);
    if (infoType === "address") return standaloneFormulaStringLiteral("$" + columnName(range.left) + "$" + range.top);
    return null;
  }

  function standaloneSplitFormulaArguments(source) {
    return standaloneSplitFormulaTopLevel(source, { ",": true, ";": true });
  }

  function standaloneSplitFormulaTopLevel(source, separators) {
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
      if (separators[char] && depth === 0) {
        parts.push(current.trim());
        current = "";
        continue;
      }
      current += char;
    }
    parts.push(current.trim());
    return parts;
  }

  function standaloneFormulaStaticStringLiteral(value) {
    const text = String(value || "").trim();
    if (!/^"(?:[^"]|"")*"$/.test(text)) return "";
    return text.slice(1, -1).replace(/""/g, '"').trim().toLocaleLowerCase("ja-JP");
  }

  function standaloneFormulaStringLiteral(value) {
    return '"' + String(value ?? "").replace(/"/g, '""') + '"';
  }

  function standaloneFormulaContextCellRange(context) {
    const row = Number(context?.row);
    const col = Number(context?.col);
    if (!Number.isInteger(row) || !Number.isInteger(col) || row < 1 || col < 1) return null;
    return { top: row, left: col, bottom: row, right: col };
  }

  function standaloneFormulaStaticA1ReferenceRange(expression) {
    let text = standaloneStripFormulaOuterParentheses(String(expression || "").trim());
    if (!text) return null;
    const bangIndex = standaloneLastFormulaSheetBangIndex(text);
    if (bangIndex >= 0) {
      const sheetName = standaloneFormulaReferenceSheetName(text.slice(0, bangIndex));
      if (sheetName && model?.sheets?.length && !standaloneFormulaReferenceSheetExists(sheetName)) return null;
      text = text.slice(bangIndex + 1).trim();
    }
    if (!/^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?$/.test(text)) return null;
    return standaloneDecodeA1Range(text);
  }

  function standaloneFormulaReferenceSheetName(prefix) {
    let text = String(prefix || "").trim();
    if (!text) return "";
    if (text.startsWith("'") && text.endsWith("'")) text = text.slice(1, -1).replace(/''/g, "'");
    const bracketIndex = text.lastIndexOf("]");
    if (bracketIndex >= 0) text = text.slice(bracketIndex + 1);
    return text.trim();
  }

  function standaloneFormulaReferenceSheetExists(sheetName) {
    const target = standaloneNormalizeFormulaSheetName(sheetName);
    return (model?.sheets || []).some(function (sheet) {
      return standaloneNormalizeFormulaSheetName(sheet.name) === target;
    });
  }

  function standaloneStripFormulaOuterParentheses(expression) {
    let text = String(expression || "").trim();
    while (text.startsWith("(") && standaloneFindFormulaClosingParen(text, 0) === text.length - 1) {
      text = text.slice(1, -1).trim();
    }
    return text;
  }

  function standaloneLastFormulaSheetBangIndex(expression) {
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

  function standaloneColumnNumber(name) {
    return String(name || "")
      .toUpperCase()
      .split("")
      .reduce(function (total, char) { return total * 26 + char.charCodeAt(0) - 64; }, 0);
  }

  function standaloneDecodeA1Cell(address) {
    const match = String(address || "").match(/^\$?([A-Z]+)\$?(\d+)$/i);
    return match ? { col: standaloneColumnNumber(match[1]), row: Number(match[2]) } : null;
  }

  function standaloneDecodeA1Range(range) {
    const parts = String(range || "").split(":");
    const start = standaloneDecodeA1Cell(parts[0]);
    const end = standaloneDecodeA1Cell(parts[1] || parts[0]);
    if (!start || !end) return null;
    return {
      top: Math.min(start.row, end.row),
      left: Math.min(start.col, end.col),
      bottom: Math.max(start.row, end.row),
      right: Math.max(start.col, end.col),
    };
  }

  function translateStandaloneFormulaForFormulaEngine(formula, options) {
    if (typeof formula !== "string" || !formula.startsWith("=")) return formula;
    const structured = translateStandaloneStructuredReferencesForFormulaEngine(formula, options?.sheetIndex || 0, {
      row: options?.row,
      col: options?.col,
    });
    const definedNames = translateStandaloneDefinedNamesForFormulaEngine(structured, options?.definedNameAliases, options?.sheetIndex);
    return standaloneRewriteStaticCellInfoFormulaReferencesForFormulaEngine(definedNames, { row: options?.row, col: options?.col });
  }

  function translateStandaloneStructuredReferencesForFormulaEngine(formula, sheetIndex, context) {
    if (typeof formula !== "string" || !formula.includes("[") || !model?.sheets?.length) return formula;
    const refs = standaloneWorkbookStructuredTableRefs();
    if (!refs.length) return formula;
    const namedRefs = refs
      .filter(function (ref) { return ref.name; })
      .sort(function (left, right) { return right.name.length - left.name.length; });
    let output = "";
    let index = 0;
    while (index < formula.length) {
      const char = formula[index];
      if (char === '"') {
        const end = standaloneConsumeFormulaDoubleQuotedText(formula, index);
        output += formula.slice(index, end);
        index = end;
        continue;
      }
      if (char === "'") {
        const end = standaloneConsumeFormulaSingleQuotedSheetName(formula, index);
        output += formula.slice(index, end);
        index = end;
        continue;
      }
      if (char === "[") {
        const token = standaloneStructuredReferenceTokenAt(formula, index);
        if (token) {
          const tableRef = standaloneStructuredReferenceContextTable(refs, sheetIndex, context || {});
          const replacement = tableRef ? standaloneStructuredReferenceToA1(token.text, tableRef, context || {}) : "";
          if (replacement) {
            output += replacement;
            index = token.end;
            continue;
          }
        }
      }
      const namedRef = namedRefs.find(function (ref) {
        return (
          standaloneStructuredReferenceNameMatches(formula, index, ref.name) &&
          formula[index + ref.name.length] === "[" &&
          standaloneIsDefinedNameTokenBoundary(formula, index, ref.name.length)
        );
      });
      if (namedRef) {
        const token = standaloneStructuredReferenceTokenAt(formula, index + namedRef.name.length);
        const replacement = token ? standaloneStructuredReferenceToA1(token.text, namedRef, context || {}) : "";
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

  function standaloneWorkbookStructuredTableRefs() {
    const refs = [];
    (model?.sheets || []).forEach(function (sheet, sheetIndex) {
      (sheet.tables || []).forEach(function (table, tableIndex) {
        if (!table?.range) return;
        refs.push({
          sheet,
          sheetIndex,
          table,
          tableIndex,
          name: table.displayName || table.name || "Table" + (tableIndex + 1),
          columns: standaloneTableColumnNamesForStructuredRef(sheet, table),
        });
      });
    });
    return refs;
  }

  function standaloneTableColumnNamesForStructuredRef(sheet, table) {
    if (Array.isArray(table?.columns) && table.columns.length) {
      return table.columns.map(function (column, index) { return String(column?.name || "Column" + (index + 1)); });
    }
    const names = [];
    const range = table?.range;
    if (!range) return names;
    for (let col = range.left; col <= range.right; col += 1) {
      const raw = sheet?.cells?.[cellKey(range.top, col)]?.raw;
      names.push(String(raw || "Column" + (col - range.left + 1)));
    }
    return names;
  }

  function standaloneStructuredReferenceNameMatches(formula, index, name) {
    return formula.slice(index, index + name.length).toLocaleLowerCase("ja-JP") === String(name || "").toLocaleLowerCase("ja-JP");
  }

  function standaloneStructuredReferenceTokenAt(formula, start) {
    if (formula[start] !== "[") return null;
    let depth = 0;
    for (let index = start; index < formula.length; index += 1) {
      if (formula[index] === "[") depth += 1;
      else if (formula[index] === "]") {
        depth -= 1;
        if (depth === 0) return { text: formula.slice(start, index + 1), end: index + 1 };
      }
    }
    return null;
  }

  function standaloneStructuredReferenceContextTable(refs, sheetIndex, context) {
    const row = Number(context?.row);
    const col = Number(context?.col);
    return refs.find(function (ref) {
      if (ref.sheetIndex !== Number(sheetIndex)) return false;
      const range = ref.table?.range;
      if (!range || !Number.isFinite(row)) return false;
      const rows = standaloneStructuredReferenceRows(ref.table, "thisRow", row);
      return row >= rows.top && row <= rows.bottom && (!Number.isFinite(col) || (col >= range.left && col <= range.right));
    }) || refs.find(function (ref) { return ref.sheetIndex === Number(sheetIndex); });
  }

  function standaloneStructuredReferenceToA1(token, tableRef, context) {
    const spec = standaloneParseStructuredReferenceSpec(token);
    if (!spec || !tableRef?.table?.range) return "";
    const range = standaloneStructuredReferenceResolvedRange(tableRef, spec, context || {});
    if (!range) return "";
    return standaloneStructuredReferenceA1(tableRef.sheet, range);
  }

  function standaloneParseStructuredReferenceSpec(token) {
    const inner = String(token || "").slice(1, -1).trim();
    if (!inner) return null;
    const parts = standaloneSplitStructuredReferenceParts(inner);
    const spec = { rowKind: "", currentRow: false, colStart: "", colEnd: "" };
    parts.forEach(function (part) {
      const raw = part.trim();
      const range = raw.match(/^\[([^\]]+)\]\s*:\s*\[([^\]]+)\]$/);
      if (range) {
        spec.colStart = standaloneCleanStructuredReferencePart(range[1]);
        spec.colEnd = standaloneCleanStructuredReferencePart(range[2]);
        return;
      }
      let item = standaloneCleanStructuredReferencePart(raw);
      if (!item) return;
      if (/^@\[(.*)\]$/.test(item)) item = item.replace(/^@\[(.*)\]$/, "@$1");
      if (item.startsWith("@")) {
        spec.currentRow = true;
        spec.rowKind = "thisRow";
        spec.colStart = standaloneCleanStructuredReferencePart(item.slice(1));
        spec.colEnd = spec.colStart;
        return;
      }
      const normalized = item.replace(/^'/, "").toLocaleLowerCase("ja-JP");
      if (normalized === "#all") spec.rowKind = "all";
      else if (normalized === "#data") spec.rowKind = "data";
      else if (normalized === "#headers") spec.rowKind = "headers";
      else if (normalized === "#totals") spec.rowKind = "totals";
      else if (normalized === "#this row") {
        spec.rowKind = "thisRow";
        spec.currentRow = true;
      } else if (!spec.colStart) {
        spec.colStart = item;
        spec.colEnd = item;
      }
    });
    return spec;
  }

  function standaloneSplitStructuredReferenceParts(inner) {
    const text = String(inner || "").trim();
    const token = standaloneStructuredReferenceTokenAt(text, 0);
    const normalized = text.startsWith("[") && text.endsWith("]") && token?.end === text.length ? text.slice(1, -1) : text;
    const parts = [];
    let start = 0;
    let depth = 0;
    for (let index = 0; index < normalized.length; index += 1) {
      const char = normalized[index];
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

  function standaloneCleanStructuredReferencePart(part) {
    let text = String(part || "").trim();
    const token = standaloneStructuredReferenceTokenAt(text, 0);
    if (text.startsWith("[") && text.endsWith("]") && token?.end === text.length) {
      text = text.slice(1, -1).trim();
    }
    return text.replace(/''/g, "'");
  }

  function standaloneStructuredReferenceResolvedRange(tableRef, spec, context) {
    const table = tableRef.table;
    const tableRange = table.range;
    const rowRange = standaloneStructuredReferenceRows(table, spec.rowKind || (spec.currentRow ? "thisRow" : "data"), context?.row);
    if (!rowRange) return null;
    const colRange = standaloneStructuredReferenceColumns(tableRef, spec);
    if (!colRange) return null;
    return {
      top: rowRange.top,
      bottom: rowRange.bottom,
      left: colRange.left,
      right: colRange.right,
      sheetIndex: tableRange.sheetIndex ?? tableRef.sheetIndex,
    };
  }

  function standaloneStructuredReferenceRows(table, rowKind, currentRow) {
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
        return Number.isFinite(row) ? { top: row, bottom: row } : null;
      }
      case "data":
      default:
        return dataBottom >= dataTop ? { top: dataTop, bottom: dataBottom } : null;
    }
  }

  function standaloneStructuredReferenceColumns(tableRef, spec) {
    const range = tableRef.table?.range;
    if (!range) return null;
    const names = tableRef.columns || [];
    if (!spec.colStart) return { left: range.left, right: range.right };
    const startOffset = standaloneStructuredReferenceColumnOffset(names, spec.colStart);
    const endOffset = standaloneStructuredReferenceColumnOffset(names, spec.colEnd || spec.colStart);
    if (startOffset < 0 || endOffset < 0) return null;
    return {
      left: range.left + Math.min(startOffset, endOffset),
      right: range.left + Math.max(startOffset, endOffset),
    };
  }

  function standaloneStructuredReferenceColumnOffset(names, name) {
    const target = String(name || "").trim().toLocaleLowerCase("ja-JP");
    return (names || []).findIndex(function (columnNameText) {
      return String(columnNameText || "").trim().toLocaleLowerCase("ja-JP") === target;
    });
  }

  function standaloneStructuredReferenceA1(sheet, range) {
    const sheetName = standaloneQuoteSheetNameForFormula(sheet?.name || "");
    const start = "$" + columnName(range.left) + "$" + range.top;
    const end = "$" + columnName(range.right) + "$" + range.bottom;
    return sheetName + "!" + (start === end ? start : start + ":" + end);
  }

  function standaloneIsCellReference(value) {
    return /^[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}$/.test(String(value || "").replace(/\$/g, ""));
  }

  function renderTabs() {
    $tabs.html(
      model.sheets
        .map((sheet, index) => {
          const active = index === state.active ? " class='active'" : "";
          return "<button type='button' data-index='" + index + "'" + active + ">" + escapeHtml(sheet.name) + "</button>";
        })
        .join(""),
    );
  }

  function renderSheet() {
    const sheet = model.sheets[state.active];
    const sheetId = hf.getSheetId(sheet.name);
    const maps = buildMergeMaps(sheet.merges);
    let html = "<div class='" + sheetGridClassName(sheet) + "' style='" + escapeAttr(gridTrackStyle(sheet)) + "'>";
    html += "<canvas class='sheet-canvas' aria-hidden='true'></canvas>";
    html += "<div class='grid-header corner-header' style='grid-column:1;grid-row:1'></div>";
    for (let col = 1; col <= sheet.colCount; col += 1) {
      html += "<div class='grid-header column-header' style='grid-column:" + visualGridColumnStart(sheet, col) + ";grid-row:1'>" + columnName(col) + "</div>";
    }
    for (let row = 1; row <= sheet.rowCount; row += 1) {
      const rowInfo = sheet.rows[row - 1] || {};
      html += "<div class='grid-header row-header' style='grid-column:1;grid-row:" + (row + 1) + ";" + (rowInfo.hidden ? "display:none" : "") + "'>" + row + "</div>";
      for (let col = 1; col <= sheet.colCount; col += 1) {
        const key = row + ":" + col;
        if (maps.covered.has(key)) continue;
        const cell = sheet.cells[key] || {};
        const formula = cell.formula || getFormula(sheetId, row, col);
        const value = state.showFormulas && formula ? formula : getValue(sheet, sheetId, row, col);
        const merge = maps.master.get(key);
        const rowSpan = merge ? merge.bottom - merge.top + 1 : 1;
        const colSpan = merge ? merge.right - merge.left + 1 : 1;
        const overflowWidth = textOverflowWidth(sheet, maps, row, col, colSpan, cell, value);
        const css = overlayCellCss(cellEffectiveCss(sheet, cell), {
          gridColumn: visualGridColumnStart(sheet, col, colSpan) + " / span " + colSpan,
          gridRow: row + 1 + " / span " + rowSpan,
        }, canvasFontDisplayScale(sheet));
        if (sheet.columns[col - 1]?.hidden || rowInfo.hidden) css.display = "none";
        const attrs = ["data-row='" + row + "'", "data-col='" + col + "'", "tabindex='0'", "style='" + escapeAttr(styleObjectToString(css)) + "'"];
        const annotationTitle = cellAnnotationTitle(cell);
        if (annotationTitle) attrs.push("title='" + escapeAttr(annotationTitle) + "'");
        const classNames = ["sheet-cell"];
        if (overflowWidth) classNames.push("text-overflow");
        if (cellHasFill(cell)) classNames.push("has-fill");
        if (shouldHideGridRight(sheet, maps, row, col, rowSpan, colSpan, cell)) classNames.push("hide-grid-right");
        if (shouldHideGridBottom(sheet, maps, row, col, rowSpan, colSpan, cell)) classNames.push("hide-grid-bottom");
        if (cellHasComment(cell)) classNames.push("has-comment");
        if (cellHasNote(cell)) classNames.push("has-note");
        if (cellHasPhonetics(cell)) classNames.push("has-phonetic");
        if (cell?.hyperlink) classNames.push("has-hyperlink");
        if (cellShowsValidationDropdown(cell) && standaloneValidationListOptions(cell).length) classNames.push("has-validation-dropdown");
        if (cell?.showPhonetic) classNames.push("show-phonetic");
        if (formula) classNames.push("formula-cell");
        if (state.selected?.row === row && state.selected?.col === col) classNames.push("selected");
        if (classNames.length) attrs.push("class='" + classNames.join(" ") + "'");
        html += "<div " + attrs.join(" ") + ">" + cellInnerHtml(cell, value, overflowWidth) + "</div>";
      }
    }
    html += renderStandaloneSelectionFrames(sheet);
    html += renderImageLayer(sheet, state.active);
    html += "</div>";
    $sheet.html(html);
    paintSheetCanvas(sheet, maps, sheetId);
    updateHorizontalScrollbar();
    updateFooterUi();
  }

  function cellInnerHtml(cell, value, overflowWidth) {
    const text = value == null ? "" : String(value);
    const indicators = cellAnnotationMarkersHtml(cell);
    if (cellShowsValidationDropdown(cell) && standaloneValidationListOptions(cell).length) {
      return indicators + "<span class='cell-content' " + cellContentStyle(overflowWidth) + ">" + escapeHtml(text) + "</span>" + standaloneValidationDropdownButtonHtml(cell);
    }
    if (cell.hyperlink) {
      const hyperlink = escapeAttr(cell.hyperlink);
      return indicators + "<a class='cell-link cell-content' " + cellContentStyle(overflowWidth) + " href='" + hyperlink + "' data-hyperlink='" + hyperlink + "' target='_blank' rel='noopener noreferrer'>" + escapeHtml(text) + "</a>";
    }
    return indicators + "<span class='cell-content' " + cellContentStyle(overflowWidth) + ">" + escapeHtml(text) + "</span>";
  }
  function standaloneValidationDropdownButtonHtml(cell) {
    const label = columnName(cell?.col || 1) + (cell?.row || 1) + " ドロップダウン リストを表示";
    return "<button class='cell-validation-dropdown-button' type='button' aria-label='" + escapeAttr(label) + "' title='ドロップダウン リストを表示'></button>";
  }
  function standaloneValidationListOptions(cell) {
    const rule = cell?.validation;
    if (!rule || rule.type !== "list" || rule.showDropDown === false) return [];
    if (Array.isArray(rule.options) && rule.options.length) {
      return rule.options.map(function (option) {
        return String(option ?? "");
      });
    }
    return parseValidationOptions(rule.formula || "").map(function (option) {
      return String(option ?? "");
    });
  }
  function handleStandaloneValidationDropdownButtonPointerDown(event) {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
  }
  function handleStandaloneValidationDropdownButtonClick(event) {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    openStandaloneValidationDropdown($(event.currentTarget).closest(".sheet-cell[data-row][data-col]"));
  }
  function openStandaloneValidationDropdown($cell) {
    if (!$cell?.length) return false;
    const row = Number($cell.data("row"));
    const col = Number($cell.data("col"));
    const sheet = model.sheets[state.active];
    const cell = sheet?.cells?.[row + ":" + col];
    const options = standaloneValidationListOptions(cell);
    if (!row || !col || !options.length) {
      $status.text("このセルには選択できるドロップダウン リストがありません。");
      return false;
    }
    select($cell);
    closeStandaloneValidationDropdownPopup();
    const current = String(cell?.raw ?? getValue(sheet, hf.getSheetId(sheet.name), row, col) ?? "");
    const popup = document.createElement("div");
    popup.id = "wxValidationDropdownPopup";
    popup.className = "validation-dropdown-popup";
    popup.setAttribute("role", "listbox");
    popup.setAttribute("aria-label", columnName(col) + row + " ドロップダウン リスト");
    popup.dataset.row = String(row);
    popup.dataset.col = String(col);
    popup.innerHTML = options.map(function (option) {
      const value = String(option);
      const selected = value === current ? " is-selected" : "";
      return "<button class='validation-dropdown-choice" + selected + "' type='button' role='option' aria-selected='" + (selected ? "true" : "false") + "' data-standalone-validation-dropdown-choice='" + escapeAttr(value) + "'>" + escapeHtml(value) + "</button>";
    }).join("");
    document.body.appendChild(popup);
    positionStandaloneValidationDropdownPopup(popup, $cell[0].getBoundingClientRect());
    state.validationDropdownPopup = { row, col };
    (popup.querySelector(".is-selected") || popup.querySelector("[data-standalone-validation-dropdown-choice]"))?.focus({ preventScroll: true });
    return true;
  }
  function positionStandaloneValidationDropdownPopup(popup, cellRect) {
    const margin = 6;
    popup.style.visibility = "hidden";
    popup.style.left = "0px";
    popup.style.top = "0px";
    const rect = popup.getBoundingClientRect();
    const width = Math.max(rect.width, cellRect.width, 120);
    popup.style.width = Math.round(width) + "px";
    const adjusted = popup.getBoundingClientRect();
    const left = Math.max(margin, Math.min(cellRect.left, window.innerWidth - adjusted.width - margin));
    const top = Math.max(margin, Math.min(cellRect.bottom, window.innerHeight - adjusted.height - margin));
    popup.style.left = Math.round(left) + "px";
    popup.style.top = Math.round(top) + "px";
    popup.style.visibility = "visible";
  }
  function handleStandaloneValidationDropdownChoice(event) {
    const popup = event.currentTarget.closest("#wxValidationDropdownPopup");
    const row = Number(popup?.dataset.row);
    const col = Number(popup?.dataset.col);
    const value = event.currentTarget.dataset.standaloneValidationDropdownChoice || "";
    closeStandaloneValidationDropdownPopup();
    if (!row || !col) return;
    commit(row, col, value);
    $status.text(columnName(col) + row + " に「" + value + "」を入力しました。");
  }
  function handleStandaloneValidationDropdownKeydown(event) {
    const popup = event.currentTarget.closest("#wxValidationDropdownPopup");
    const choices = Array.prototype.slice.call(popup?.querySelectorAll("[data-standalone-validation-dropdown-choice]") || []);
    if (!choices.length) return;
    const currentIndex = Math.max(0, choices.indexOf(event.currentTarget));
    const focusChoice = function (index) {
      const next = choices[Math.max(0, Math.min(choices.length - 1, index))];
      next?.focus({ preventScroll: true });
    };
    const consume = function () {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation?.();
    };
    if (event.key === "ArrowDown") {
      consume();
      focusChoice(currentIndex + 1);
    } else if (event.key === "ArrowUp") {
      consume();
      focusChoice(currentIndex - 1);
    } else if (event.key === "Home") {
      consume();
      focusChoice(0);
    } else if (event.key === "End") {
      consume();
      focusChoice(choices.length - 1);
    } else if (event.key === "PageDown") {
      consume();
      focusChoice(currentIndex + 8);
    } else if (event.key === "PageUp") {
      consume();
      focusChoice(currentIndex - 8);
    } else if (event.key === "Enter" || event.key === " ") {
      consume();
      event.currentTarget.click();
    }
  }
  function closeStandaloneValidationDropdownOnOutsidePointer(event) {
    if (!state.validationDropdownPopup) return;
    if (event.target.closest?.("#wxValidationDropdownPopup")) return;
    closeStandaloneValidationDropdownPopup();
  }
  function closeStandaloneValidationDropdownOnEscape(event) {
    if (event.key === "Escape" && state.validationDropdownPopup) closeStandaloneValidationDropdownPopup();
  }
  function closeStandaloneValidationDropdownPopup() {
    document.querySelector("#wxValidationDropdownPopup")?.remove();
    state.validationDropdownPopup = null;
  }
  function cellContentStyle(overflowWidth) {
    return overflowWidth ? "style='width:" + overflowWidth + "px'" : "";
  }
  function renderStandaloneSelectionFrames(sheet) {
    if (!state.selectedRange && !state.selectedRanges?.length) return "";
    const ranges = standaloneActiveSelectionRanges()
      .map(function (range) {
        return standaloneNormalizeRange(sheet, range);
      })
      .filter(Boolean);
    if (!ranges.length) return "";
    return ranges
      .map(function (range, index) {
        const rect = rangePixelRect(sheet, range);
        if (!rect || rect.width <= 0 || rect.height <= 0) return "";
        const activeClass = index === 0 ? " is-active" : "";
        return "<div class='wx-selection-frame" + activeClass + "' style='" +
          escapeAttr("left:" + rect.left + "px;top:" + rect.top + "px;width:" + rect.width + "px;height:" + rect.height + "px") +
          "' aria-hidden='true'></div>";
      })
      .join("");
  }
  function standaloneActiveSelectionRanges() {
    if (Array.isArray(state.selectedRanges) && state.selectedRanges.length) return state.selectedRanges;
    if (state.selectedRange) return [state.selectedRange];
    if (state.selected) return [{ top: state.selected.row, left: state.selected.col, bottom: state.selected.row, right: state.selected.col }];
    return [];
  }
  function standaloneNormalizeRange(sheet, range) {
    if (!sheet || !range) return null;
    const rowCount = Math.max(1, Number(sheet.rowCount) || (sheet.rows || []).length || 1);
    const colCount = Math.max(1, Number(sheet.colCount) || (sheet.columns || []).length || 1);
    const top = Math.max(1, Math.min(rowCount, Math.round(Number(range.top) || 1)));
    const left = Math.max(1, Math.min(colCount, Math.round(Number(range.left) || 1)));
    const bottom = Math.max(top, Math.min(rowCount, Math.round(Number(range.bottom || range.top) || top)));
    const right = Math.max(left, Math.min(colCount, Math.round(Number(range.right || range.left) || left)));
    return { top, left, bottom, right };
  }
  function cellAnnotationTitle(cell) {
    const lines = [];
    const comment = normalizeCommentText(cell?.comment);
    const note = normalizeNoteText(cell?.note);
    if (comment) lines.push("コメント:\\n" + comment);
    if (note) lines.push("メモ:\\n" + note);
    return lines.join("\\n\\n");
  }
  function normalizeCommentText(comment) {
    if (!comment) return "";
    if (typeof comment === "string") return comment;
    if (typeof comment.text === "string") return comment.text;
    if (Array.isArray(comment.replies)) return comment.replies.map((reply) => reply?.text || "").filter(Boolean).join("\\n");
    return String(comment);
  }
  function normalizeNoteText(note) {
    if (!note) return "";
    if (typeof note === "string") return note;
    if (Array.isArray(note.texts)) return note.texts.map((part) => part.text || "").join("");
    return String(note);
  }
  function cellHasComment(cell) {
    return Boolean(normalizeCommentText(cell?.comment));
  }
  function cellHasNote(cell) {
    return Boolean(normalizeNoteText(cell?.note));
  }
  function cellHasPhonetics(cell) {
    return Boolean(cell?.phonetics?.length);
  }
  function cellAnnotationMarkersHtml(cell) {
    const markers = [];
    if (cellHasComment(cell)) markers.push("<span class='cell-annotation-marker is-comment' aria-hidden='true'></span>");
    if (cellHasNote(cell)) markers.push("<span class='cell-annotation-marker is-note' aria-hidden='true'></span>");
    return markers.join("");
  }
  function overlayCellCss(effectiveCss, baseCss, fontScale) {
    const css = { ...baseCss, ...(effectiveCss || {}) };
    css.fontSize = canvasFontSize(css.fontSize || "13px", fontScale || 1) + "px";
    if (css.backgroundColor) css["--cell-fill-color"] = css.backgroundColor;
    delete css.backgroundColor;
    delete css.borderColor;
    delete css.borderTopColor;
    delete css.borderRightColor;
    delete css.borderBottomColor;
    delete css.borderLeftColor;
    return css;
  }
  function cellEffectiveCss(sheet, cell) {
    return { ...(sheet?.defaultCellCss || {}), ...(cell?.css || {}) };
  }
  function cellHasFill(cell) {
    return Boolean(cell?.hasFill || cell?.css?.backgroundColor);
  }
  function shouldHideGridRight(sheet, maps, row, col, rowSpan, colSpan, cell) {
    if (cellHasFill(cell)) return true;
    const rightCol = col + colSpan;
    for (let edgeRow = row; edgeRow < row + rowSpan; edgeRow += 1) {
      if (gridPositionHasFill(sheet, maps, edgeRow, rightCol)) return true;
    }
    return false;
  }
  function shouldHideGridBottom(sheet, maps, row, col, rowSpan, colSpan, cell) {
    if (cellHasFill(cell)) return true;
    const bottomRow = row + rowSpan;
    for (let edgeCol = col; edgeCol < col + colSpan; edgeCol += 1) {
      if (gridPositionHasFill(sheet, maps, bottomRow, edgeCol)) return true;
    }
    return false;
  }
  function gridPositionHasFill(sheet, maps, row, col) {
    if (row < 1 || col < 1 || row > sheet.rowCount || col > sheet.colCount) return false;
    const key = row + ":" + col;
    const masterKey = maps.coveredToMaster.get(key) || key;
    return cellHasFill(sheet.cells[masterKey]);
  }
  const runtimeImageObjectUrlMinDataUrlLength = 4096;
  const runtimeImageObjectUrls = new Map();
  function runtimeCachedImageSource(src) {
    const value = String(src || "");
    if (
      value.length < runtimeImageObjectUrlMinDataUrlLength ||
      !/^data:image\\//i.test(value) ||
      typeof Blob === "undefined" ||
      typeof URL === "undefined" ||
      typeof URL.createObjectURL !== "function"
    ) {
      return value;
    }
    const cached = runtimeImageObjectUrls.get(value);
    if (cached) return cached;
    const blob = runtimeImageDataUrlToBlob(value);
    if (!blob) return value;
    const objectUrl = URL.createObjectURL(blob);
    runtimeImageObjectUrls.set(value, objectUrl);
    return objectUrl;
  }
  function runtimeImageDataUrlToBlob(src) {
    const match = String(src || "").match(/^data:([^,]+),([\\s\\S]*)$/i);
    if (!match) return null;
    const meta = match[1] || "image/png";
    const data = match[2] || "";
    const mimeType = meta.split(";")[0] || "image/png";
    try {
      const bytes = /;base64(?:;|$)/i.test(";" + meta)
        ? runtimeBase64ToBytes(data)
        : new TextEncoder().encode(decodeURIComponent(data));
      return new Blob([bytes], { type: mimeType });
    } catch (error) {
      console.warn("Failed to create runtime image object URL.", error);
      return null;
    }
  }
  function runtimeBase64ToBytes(value) {
    const binary = atob(String(value || ""));
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  }
  window.addEventListener("pagehide", function () {
    if (typeof URL === "undefined" || typeof URL.revokeObjectURL !== "function") return;
    runtimeImageObjectUrls.forEach(function (url) {
      URL.revokeObjectURL(url);
    });
    runtimeImageObjectUrls.clear();
  });
	  function renderImageLayer(sheet, sheetIndex = state.active) {
	    if (sheet.showObjects === false) return "";
	    if (!sheet.images?.length) return "";
	    return sheet.images
	      .map((image) => {
	        if (image.hidden) return "";
	        if (image?.chart?.series?.length) {
	          refreshStandaloneChartDataFromSources(image.chart, sheetIndex);
	          refreshStandaloneChartSource(image);
	        }
	        const box = imageBoxToPixels(sheet, image);
	        if (!box) return "";
	        const style = runtimeImageBoxStyle(image, box);
	        const imageHtml = runtimeImageObjectHtml(image, style, sheet.displayScale);
	        if (!image.hyperlink) return imageHtml;
	        return "<a class='sheet-image-link' href='" + escapeAttr(image.hyperlink) + "' target='_blank' rel='noopener noreferrer' style='" + escapeAttr(style) + "'>" + imageHtml + "</a>";
	      })
	      .join("");
	  }
	  function runtimeImageObjectHtml(image, style, textDisplayScale = 1) {
	    const crop = runtimeNormalizePictureCrop(image?.crop);
	    const isLineShape = runtimeIsLineShapeObject(image);
	    const imageClass = "sheet-image" + (image.shape ? " sheet-shape" : "") + (isLineShape ? " is-line-shape" : "") + (crop ? " sheet-picture-crop" : "");
	    const alt = escapeAttr(image.name || "worksheet image");
	    const src = runtimeCachedImageSource(image.src);
	    if (image?.shape && !isLineShape && !crop) {
	      return "<div class='" + imageClass + "' style='" + escapeAttr(style) + "' role='img' aria-label='" + alt + "'><img class='sheet-shape-render' src='" + escapeAttr(src) + "' alt='' draggable='false' loading='lazy' decoding='async' />" + runtimeShapeTextLayerHtml(image, textDisplayScale) + "</div>";
	    }
	    if (!crop) return "<img class='" + imageClass + "' src='" + escapeAttr(src) + "' alt='" + alt + "' style='" + escapeAttr(style) + "' draggable='false' loading='lazy' decoding='async' />";
	    return "<div class='" + imageClass + "' style='" + escapeAttr(style) + "' role='img' aria-label='" + alt + "'><img class='sheet-picture-crop-content' src='" + escapeAttr(src) + "' alt='' style='" + escapeAttr(runtimePictureCropContentStyle(crop)) + "' draggable='false' loading='lazy' decoding='async' /></div>";
	  }
	  function runtimeIsLineShapeObject(image) {
	    return Boolean(image?.shape && runtimeIsLineShapeType(image.shape.type));
	  }
	  function runtimeIsLineShapeType(type) {
	    return ["line", "arrow", "doubleArrow", "elbowConnector", "curvedConnector", "freeform"].includes(String(type || ""));
	  }
	  function runtimeShapeTextLayerHtml(image, textDisplayScale = 1) {
	    const shape = image?.shape;
	    if (!shape || runtimeIsLineShapeObject(image)) return "";
	    const runs = runtimeNormalizeShapeTextRuns(shape.textRuns);
	    const text = Object.prototype.hasOwnProperty.call(shape, "text") ? String(shape.text ?? "") : runtimeShapePlainTextFromRuns(runs);
	    if (!text && !runs.length) return "";
	    const content = runs.length ? runtimeShapeTextRunsLayerHtml(runs) : runtimeShapeTextPlainLayerHtml(text);
	    return "<div class='sheet-shape-text-layer' style='" + escapeAttr(runtimeShapeTextLayerStyle(shape) + ";" + runtimeShapeTextLayerScaleStyle(textDisplayScale)) + "'><div class='sheet-shape-text-block' style='width:100%;min-width:0;flex:none'>" + content + "</div></div>";
	  }
	  function runtimeShapeTextLayerScaleStyle(value) {
	    const scale = Number(value);
	    if (!Number.isFinite(scale) || scale <= 0 || scale === 1) return "";
	    return "width:" + (100 / scale) + "%;height:" + (100 / scale) + "%;transform-origin:0 0;transform:scale(" + scale + ")";
	  }
	  function runtimeShapePlainTextFromRuns(runs) {
	    return runtimeNormalizeShapeTextRuns(runs).map(function (run) { return String(run.text ?? ""); }).join("");
	  }
	  function runtimeNormalizeShapeTextRuns(runs) {
	    return Array.isArray(runs) ? runs.filter(function (run) { return run && typeof run === "object"; }) : [];
	  }
	  function runtimeShapeTextLayerStyle(shape) {
	    const margin = runtimeNormalizeShapeTextMargin(shape?.textMargin, shape?.type === "textBox");
	    const vertical = ["top", "bottom", "middle"].includes(shape?.textVAlign) ? shape.textVAlign : shape?.type === "textBox" ? "top" : "middle";
	    const justifyContent = vertical === "top" ? "flex-start" : vertical === "bottom" ? "flex-end" : "center";
	    const underline = String(shape?.textUnderline || "");
	    return [
	      "padding:" + runtimeRound(margin.top) + "px " + runtimeRound(margin.right) + "px " + runtimeRound(margin.bottom) + "px " + runtimeRound(margin.left) + "px",
	      "color:" + (shape?.textColor || "#1f2937"),
	      "font-family:" + (shape?.textFontFamily || "Calibri, 'Yu Gothic UI', Meiryo, sans-serif"),
	      "font-size:" + Math.max(8, Math.round((Number(shape?.textFontSize) || 11) * 96 / 72)) + "px",
	      "font-weight:" + (shape?.textBold ? "700" : "400"),
	      "font-style:" + (shape?.textItalic ? "italic" : "normal"),
	      "text-align:" + (shape?.textAlign || "center"),
	      "justify-content:" + justifyContent,
	      underline ? "text-decoration-line:underline" : "text-decoration-line:none",
	      underline === "underline double" ? "text-decoration-style:double" : "",
	      shape?.textOutlineColor && shape.textOutlineColor !== "none" ? "-webkit-text-stroke:" + runtimeRound(Number(shape.textOutlineWidth) || 0.75) + "px " + shape.textOutlineColor : "",
	    ].filter(Boolean).join(";");
	  }
	  function runtimeNormalizeShapeTextMargin(margin, compact) {
	    const defaults = compact ? { left: 5, right: 5, top: 4, bottom: 4 } : { left: 10, right: 10, top: 8, bottom: 8 };
	    if (!margin || typeof margin !== "object") return defaults;
	    return {
	      left: runtimeClamp(margin.left, 0, 80, defaults.left),
	      right: runtimeClamp(margin.right, 0, 80, defaults.right),
	      top: runtimeClamp(margin.top, 0, 80, defaults.top),
	      bottom: runtimeClamp(margin.bottom, 0, 80, defaults.bottom),
	    };
	  }
	  function runtimeShapeTextPlainLayerHtml(text) {
	    return runtimeSplitShapeTextLines(text).map(function (line) { return line ? escapeHtml(line) : "&#8203;"; }).join("<br>");
	  }
	  function runtimeShapeTextRunsLayerHtml(runs) {
	    const html = runtimeNormalizeShapeTextRuns(runs).map(function (run) {
	      const style = runtimeShapeTextRunLayerStyle(run);
	      const attrs = style ? " style='" + escapeAttr(style) + "'" : "";
	      return runtimeSplitShapeTextLines(run.text).map(function (line) {
	        return "<span" + attrs + ">" + (line ? escapeHtml(line) : "&#8203;") + "</span>";
	      }).join("<br>");
	    }).join("");
	    return html || "&#8203;";
	  }
	  function runtimeSplitShapeTextLines(text) {
	    return String(text ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
	  }
	  function runtimeShapeTextRunLayerStyle(run) {
	    const underline = String(run?.underline || "");
	    return [
	      run?.color ? "color:" + run.color : "",
	      Number.isFinite(Number(run?.fontSize)) ? "font-size:" + Math.max(8, Math.round(Number(run.fontSize) * 96 / 72)) + "px" : "",
	      run?.fontFamily ? "font-family:" + run.fontFamily : "",
	      typeof run?.bold === "boolean" ? "font-weight:" + (run.bold ? "700" : "400") : "",
	      typeof run?.italic === "boolean" ? "font-style:" + (run.italic ? "italic" : "normal") : "",
	      underline ? "text-decoration-line:underline" : "",
	      underline === "underline double" ? "text-decoration-style:double" : "",
	    ].filter(Boolean).join(";");
	  }
	  function runtimeImageBoxStyle(image, box) {
	    const transforms = [];
	    const rotation = runtimeNormalizeRotation(image?.shape ? image.shape.rotation : image?.rotation);
	    if (rotation) transforms.push("rotate(" + runtimeRound(rotation) + "deg)");
	    if (image?.shape ? image.shape.flipH : image?.flipH) transforms.push("scaleX(-1)");
	    if (image?.shape ? image.shape.flipV : image?.flipV) transforms.push("scaleY(-1)");
	    const transform = transforms.length ? ";transform:" + transforms.join(" ") + ";transform-origin:center center" : "";
	    return "left:" + runtimeRound(box.left) + "px;top:" + runtimeRound(box.top) + "px;width:" + runtimeRound(box.width) + "px;height:" + runtimeRound(box.height) + "px" + transform + runtimeObjectVisualStyle(image);
		  }
		  function runtimeObjectVisualStyle(image) {
		    return runtimeChartObjectOutlineStyle(image) + runtimePictureOutlineStyle(image) + runtimeVisualEffectStyle(image?.shape || image, image?.src || "", image);
		  }
		  function runtimeChartObjectOutlineStyle(image) {
		    if (!image?.shape || !image?.chart?.series?.length || !image.shape.outlineColor || image.shape.outlineColor === "none") return "";
		    const width = runtimeNormalizeOutlineWidth(image.shape.outlineWidth);
		    return ";border:" + runtimeRound(width) + "px " + runtimePictureOutlineCssStyle(image.shape) + " " + image.shape.outlineColor;
		  }
		  function runtimePictureOutlineStyle(image) {
		    if (!image || image.shape || !image.outlineColor || image.outlineColor === "none") return "";
		    const width = runtimeNormalizeOutlineWidth(image.outlineWidth);
		    return ";border:" + runtimeRound(width) + "px " + runtimePictureOutlineCssStyle(image) + " " + image.outlineColor;
	  }
	  function runtimePictureOutlineCssStyle(image) {
	    const compound = image?.lineCompound || "single";
	    if (compound && compound !== "single") return "double";
	    const dash = image?.lineDash || "solid";
	    if (dash === "dot" || dash === "squareDot") return "dotted";
	    if (dash && dash !== "solid") return "dashed";
	    return "solid";
	  }
		  function runtimeVisualEffectStyle(style, source, image) {
		    const isChartGraphicFrame = Boolean(image?.chart?.series?.length && style?.sourcePreset === "graphicFrame");
		    const hasEmbeddedSvgEffects = Boolean(style?.type) && String(source || "").startsWith("data:image/svg+xml") && !isChartGraphicFrame;
	    const filters = hasEmbeddedSvgEffects ? [] : runtimeShapeFilterValues(style);
	    const filter = filters.length ? ";filter:" + filters.join(" ") : "";
	    const reflection = style?.reflection;
	    const reflectionStyle = reflection ? ";-webkit-box-reflect:below " + runtimeRound(reflection.distance || 0) + "px linear-gradient(transparent, rgba(255,255,255," + runtimeRound(reflection.alpha || 0.45) + "))" : "";
	    return filter + reflectionStyle;
	  }
	  function runtimeShapeFilterValues(style) {
	    const filters = [];
	    if (style?.shadow && typeof style.shadow === "object") {
	      filters.push("drop-shadow(" + runtimeRound(style.shadow.dx || 0) + "px " + runtimeRound(style.shadow.dy || 0) + "px " + runtimeRound(Math.max(0, style.shadow.blur || 0)) + "px " + (style.shadow.color || "rgba(0,0,0,0.35)") + ")");
	    }
	    if (style?.shadow === "outer1") filters.push("drop-shadow(2px 2px 2px rgba(0,0,0,0.28))");
	    else if (style?.shadow === "outer2") filters.push("drop-shadow(3px 4px 5px rgba(0,0,0,0.24))");
	    if (style?.glow?.radius) filters.push("drop-shadow(0 0 " + runtimeRound(style.glow.radius) + "px " + (style.glow.color || "#5b9bd5") + ")");
	    if (Number(style?.softEdges)) filters.push("blur(" + runtimeRound(Number(style.softEdges) * 0.55) + "px)");
	    return filters;
	  }
	  function runtimeNormalizeOutlineWidth(value) {
	    const number = Number(value);
	    return Number.isFinite(number) && number > 0 ? Math.max(0.6, Math.min(6, number)) : 0.75;
	  }
	  function runtimePictureCropContentStyle(crop) {
	    const normalized = runtimeNormalizePictureCrop(crop) || { left: 0, top: 0, right: 0, bottom: 0 };
	    const visibleWidth = Math.max(0.01, 1 - normalized.left - normalized.right);
	    const visibleHeight = Math.max(0.01, 1 - normalized.top - normalized.bottom);
	    return "position:absolute;display:block;left:" + runtimeRound((-normalized.left / visibleWidth) * 100) + "%;top:" + runtimeRound((-normalized.top / visibleHeight) * 100) + "%;width:" + runtimeRound((1 / visibleWidth) * 100) + "%;height:" + runtimeRound((1 / visibleHeight) * 100) + "%;max-width:none;max-height:none;pointer-events:none;user-select:none";
	  }
	  function runtimeNormalizePictureCrop(crop) {
	    if (!crop || typeof crop !== "object") return null;
	    let left = runtimeClamp(crop.left, 0, 0.98, 0);
	    let top = runtimeClamp(crop.top, 0, 0.98, 0);
	    let right = runtimeClamp(crop.right, 0, 0.98, 0);
	    let bottom = runtimeClamp(crop.bottom, 0, 0.98, 0);
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
	  function runtimeClamp(value, min, max, fallback) {
	    const number = Number(value);
	    return Number.isFinite(number) ? Math.max(min, Math.min(max, number)) : fallback;
	  }
	  function runtimeNormalizeRotation(value) {
	    const number = Number(value) || 0;
	    return ((number % 360) + 360) % 360;
	  }
	  function runtimeRound(value) {
	    return Math.round(Number(value) * 100) / 100;
	  }
	  function imageBoxToPixels(sheet, image) {
    if (!image?.tl) return null;
    let left = coordinateToPixels(sheet, "col", image.tl.col, image.tl.nativeColOff, image.tl.nativeCol);
    const top = coordinateToPixels(sheet, "row", image.tl.row, image.tl.nativeRowOff, image.tl.nativeRow);
    let width = Number(image.ext?.width || 0);
    let height = Number(image.ext?.height || 0);
    if ((!width || !height) && image.br) {
      const right = coordinateToPixels(sheet, "col", image.br.col, image.br.nativeColOff, image.br.nativeCol);
      width = Math.abs(right - left);
      left = Math.min(left, right);
      height = coordinateToPixels(sheet, "row", image.br.row, image.br.nativeRowOff, image.br.nativeRow) - top;
    }
    if (width <= 0 || height <= 0) return null;
    return { left, top, width, height };
  }
  function coordinateToPixels(sheet, axis, value, nativeOffset, nativeIndex) {
    const isColumn = axis === "col";
    if (isColumn && sheetColumnsRightToLeft(sheet)) return rtlColumnCoordinateToPixels(sheet, value, nativeOffset, nativeIndex);
    const headerOffset = isColumn ? rowHeaderWidthForSheet(sheet) : 22;
    const items = isColumn ? sheet.columns : sheet.rows;
    const defaultSize = isColumn ? sheet.defaultColumnWidth : sheet.defaultRowHeight;
    const hasNativeAnchor = nativeIndex != null && Number.isFinite(Number(nativeIndex));
    const coordinate = hasNativeAnchor ? Number(nativeIndex) : Number.isFinite(Number(value)) ? Number(value) : 0;
    const whole = Math.max(0, Math.floor(coordinate));
    const fraction = hasNativeAnchor ? 0 : Math.max(0, coordinate - whole);
    let pixels = headerOffset;
    for (let index = 0; index < whole; index += 1) {
      const item = items[index];
      pixels += item?.hidden ? 0 : item?.width || item?.height || defaultSize || 0;
    }
    const current = items[whole];
    const currentSize = current?.hidden ? 0 : current?.width || current?.height || defaultSize || 0;
    pixels += fraction * currentSize;
    if (hasNativeAnchor && nativeOffset) pixels += (Number(nativeOffset) / 9525) * (sheet.displayScale || 1);
    return Math.round(pixels * 100) / 100;
  }
  function sheetColumnsRightToLeft(sheet) {
    return sheet && sheet.rightToLeft === true;
  }
  function visualGridColumnStart(sheet, col, colSpan) {
    const safeCol = Math.max(1, Math.round(Number(col) || 1));
    const safeSpan = Math.max(1, Math.round(Number(colSpan) || 1));
    if (!sheetColumnsRightToLeft(sheet)) return safeCol + 1;
    const rightCol = Math.min(Math.max(1, Number(sheet && sheet.colCount || safeCol)), safeCol + safeSpan - 1);
    return Math.max(2, Number(sheet && sheet.colCount || rightCol) - rightCol + 2);
  }
  function rtlColumnCoordinateToPixels(sheet, value, nativeOffset, nativeIndex) {
    const headerOffset = rowHeaderWidthForSheet(sheet);
    const colCount = Math.max(1, Number(sheet && (sheet.colCount || (sheet.columns || []).length) || 1));
    const hasNativeAnchor = nativeIndex != null && Number.isFinite(Number(nativeIndex));
    if (hasNativeAnchor) {
      const nativeCol = Math.max(1, Math.min(colCount, Math.floor(Number(nativeIndex)) + 1));
      const offsetPx = nativeOffset ? (Number(nativeOffset) / 9525) * (sheet.displayScale || 1) : 0;
      return Math.round((headerOffset + widthBetweenColumns(sheet, nativeCol + 1, colCount) + offsetPx) * 100) / 100;
    }
    const coordinate = Number.isFinite(Number(value)) ? Number(value) : 0;
    const whole = Math.max(0, Math.min(colCount, Math.floor(coordinate)));
    const fraction = Math.max(0, Math.min(1, coordinate - whole));
    const boundary = function (count) {
      if (count <= 0) return headerOffset + widthBetweenColumns(sheet, 1, colCount);
      if (count >= colCount) return headerOffset;
      return headerOffset + widthBetweenColumns(sheet, count + 1, colCount);
    };
    const start = boundary(whole);
    const end = boundary(Math.min(colCount, whole + 1));
    return Math.round((start + (end - start) * fraction) * 100) / 100;
  }

  function scheduleCanvasPaint() {
    if (canvasPaintFrame) return;
    canvasPaintFrame = requestAnimationFrame(function () {
      canvasPaintFrame = 0;
      const sheet = model.sheets[state.active];
      if (!sheet) return;
      paintSheetCanvas(sheet, buildMergeMaps(sheet.merges), hf.getSheetId(sheet.name));
    });
  }

  function paintSheetCanvas(sheet, maps, sheetId) {
    const canvas = $sheet.find(".sheet-canvas")[0];
    if (!canvas) return;
    const range = visibleGridRange(sheet);
    const bounds = canvasLayerBounds(sheet, range);
    const ratio = canvasDevicePixelRatio();
    const width = Math.max(1, Math.ceil(bounds.width));
    const height = Math.max(1, Math.ceil(bounds.height));
    const pixelWidth = Math.max(1, Math.ceil(width * ratio));
    const pixelHeight = Math.max(1, Math.ceil(height * ratio));

    if (canvas.width !== pixelWidth) canvas.width = pixelWidth;
    if (canvas.height !== pixelHeight) canvas.height = pixelHeight;
    canvas.style.left = bounds.left + "px";
    canvas.style.top = bounds.top + "px";
    canvas.style.width = width + "px";
    canvas.style.height = height + "px";
    canvas.dataset.left = String(bounds.left);
    canvas.dataset.top = String(bounds.top);
    canvas.dataset.dpr = String(ratio);

    const context = canvas.getContext("2d");
    if (!context) return;
    context.setTransform(ratio, 0, 0, ratio, 0, 0);
    context.clearRect(0, 0, width, height);
    context.save();
    context.beginPath();
    context.rect(0, 0, width, height);
    context.clip();
    context.translate(-bounds.left, -bounds.top);
    paintCanvasCells(context, sheet, maps, sheetId, range, "base");
    paintCanvasBorders(context, sheet, maps, range);
    paintCanvasCells(context, sheet, maps, sheetId, range, "text");
    context.restore();
  }

  function canvasLayerBounds(sheet, range) {
    const rect = rangePixelRect(sheet, range);
    return { left: rect.left, top: rect.top, width: Math.max(1, rect.width), height: Math.max(1, rect.height) };
  }

  function legacyCanvasLayerBounds(sheet, range) {
    const left = coordinateToPixels(sheet, "col", range.left - 1);
    const top = coordinateToPixels(sheet, "row", range.top - 1);
    const right = coordinateToPixels(sheet, "col", range.right);
    const bottom = coordinateToPixels(sheet, "row", range.bottom);
    return {
      left,
      top,
      width: Math.max(1, right - left),
      height: Math.max(1, bottom - top),
    };
  }

  function canvasDevicePixelRatio() {
    return Math.max(1, Math.min(2, Number(window.devicePixelRatio || 1)));
  }

  function visibleGridRange(sheet) {
    const host = $sheet[0];
    if (!host) {
      return { top: 1, left: 1, bottom: Math.min(sheet.rowCount, 40), right: Math.min(sheet.colCount, 20) };
    }
    return {
      ...visibleRowRange(sheet, host.scrollTop, host.scrollTop + host.clientHeight, 8),
      ...visibleColumnRange(sheet, host.scrollLeft, host.scrollLeft + host.clientWidth, 4),
    };
  }

  function visibleColumnRange(sheet, viewportStart, viewportEnd, overscan) {
    if (sheetColumnsRightToLeft(sheet)) return visibleColumnRangeRightToLeft(sheet, viewportStart, viewportEnd, overscan);
    let left = 1;
    let right = sheet.colCount;
    const rowHeaderWidth = rowHeaderWidthForSheet(sheet);
    let x = rowHeaderWidth;
    for (let col = 1; col <= sheet.colCount; col += 1) {
      const width = sheet.columns[col - 1]?.hidden ? 0 : sheet.columns[col - 1]?.width || sheet.defaultColumnWidth || 0;
      const next = x + width;
      if (next >= viewportStart) {
        left = Math.max(1, col - overscan);
        break;
      }
      x = next;
    }
    x = rowHeaderWidth;
    for (let col = 1; col <= sheet.colCount; col += 1) {
      const width = sheet.columns[col - 1]?.hidden ? 0 : sheet.columns[col - 1]?.width || sheet.defaultColumnWidth || 0;
      const next = x + width;
      if (x > viewportEnd) {
        right = Math.min(sheet.colCount, col + overscan);
        break;
      }
      x = next;
    }
    return { left, right: Math.max(left, right) };
  }
  function visibleColumnRangeRightToLeft(sheet, viewportStart, viewportEnd, overscan) {
    let left = Number(sheet && sheet.colCount || 1);
    let right = 1;
    let found = false;
    for (let col = 1; col <= sheet.colCount; col += 1) {
      const start = coordinateToPixels(sheet, "col", col - 1);
      const end = coordinateToPixels(sheet, "col", col);
      const visualLeft = Math.min(start, end);
      const visualRight = Math.max(start, end);
      if (visualRight >= viewportStart && visualLeft <= viewportEnd) {
        found = true;
        left = Math.min(left, col);
        right = Math.max(right, col);
      }
    }
    if (!found) return { left: 1, right: Math.min(sheet.colCount, Math.max(1, overscan + 1)) };
    return { left: Math.max(1, left - overscan), right: Math.min(sheet.colCount, right + overscan) };
  }

  function visibleRowRange(sheet, viewportStart, viewportEnd, overscan) {
    let top = 1;
    let bottom = sheet.rowCount;
    let y = 22;
    for (let row = 1; row <= sheet.rowCount; row += 1) {
      const height = sheet.rows[row - 1]?.hidden ? 0 : sheet.rows[row - 1]?.height || sheet.defaultRowHeight || 0;
      const next = y + height;
      if (next >= viewportStart) {
        top = Math.max(1, row - overscan);
        break;
      }
      y = next;
    }
    y = 22;
    for (let row = 1; row <= sheet.rowCount; row += 1) {
      const height = sheet.rows[row - 1]?.hidden ? 0 : sheet.rows[row - 1]?.height || sheet.defaultRowHeight || 0;
      const next = y + height;
      if (y > viewportEnd) {
        bottom = Math.min(sheet.rowCount, row + overscan);
        break;
      }
      y = next;
    }
    return { top, bottom: Math.max(top, bottom) };
  }

  function paintCanvasCells(context, sheet, maps, sheetId, range, phase) {
    const showGridLines = sheet.showGridLines !== false;
    const rendered = new Set();
    const editedCellKey = phase === "base" ? "" : standaloneCanvasEditedCellKey();
    const paintRange = phase === "text" ? canvasTextPaintRange(sheet, range) : range;
    for (let row = paintRange.top; row <= paintRange.bottom; row += 1) {
      for (let col = paintRange.left; col <= paintRange.right; col += 1) {
        const key = row + ":" + col;
        const masterKey = maps.coveredToMaster.get(key) || key;
        if (rendered.has(masterKey)) continue;
        const parts = masterKey.split(":");
        const masterRow = Number(parts[0]);
        const masterCol = Number(parts[1]);
        const merge = maps.master.get(masterKey);
        const cellRange = merge || { top: masterRow, left: masterCol, bottom: masterRow, right: masterCol };
        if (phase !== "text" && !rangeIntersects(cellRange, range)) continue;
        rendered.add(masterKey);
        paintCanvasCell(context, sheet, maps, sheetId, cellRange.top, cellRange.left, cellRange, showGridLines, phase || "all", editedCellKey);
      }
    }
  }

  function canvasTextPaintRange(sheet, range) {
    if (!range || sheetColumnsRightToLeft(sheet)) return range;
    const visibleLeft = coordinateToPixels(sheet, "col", Math.max(0, range.left - 1));
    const visibleRight = coordinateToPixels(sheet, "col", range.right);
    const visibleWidth = Math.max(1, visibleRight - visibleLeft);
    const lookbackPx = Math.max(2400, Math.min(8192, visibleWidth * 2));
    const gridLeft = rowHeaderWidthForSheet(sheet);
    const scanLeft = Math.max(gridLeft, visibleLeft - lookbackPx);
    let left = range.left;
    let x = gridLeft;
    for (let col = 1; col <= range.left; col += 1) {
      const width = sheet.columns[col - 1]?.hidden ? 0 : sheet.columns[col - 1]?.width || sheet.defaultColumnWidth || 0;
      const next = x + width;
      if (next >= scanLeft) {
        left = Math.min(range.left, col);
        break;
      }
      x = next;
    }
    return { ...range, left };
  }

  function paintCanvasCell(context, sheet, maps, sheetId, row, col, range, showGridLines, phase, editedCellKey) {
    const key = row + ":" + col;
    const cell = sheet.cells[key] || {};
    const rowModel = sheet.rows[row - 1] || {};
    const column = sheet.columns[col - 1] || {};
    if (rowModel.hidden || column.hidden) return;
    const rect = rangePixelRect(sheet, range);
    if (rect.width <= 0 || rect.height <= 0) return;
    if (phase !== "text") {
      if (cellHasFill(cell)) {
        context.fillStyle = cell.css?.backgroundColor || "#ffffff";
        context.fillRect(rect.left, rect.top, rect.width, rect.height);
      }
      if (showGridLines) paintCanvasGridLines(context, sheet, maps, row, col, range, cell, rect);
    }
    if (phase === "base") return;
    if (editedCellKey && editedCellKey === key) return;
    paintCanvasCellText(context, sheet, maps, sheetId, row, col, range, cell, rect);
  }

  function standaloneCanvasEditedCellKey() {
    const editor = $sheet.find(".sheet-cell[contenteditable='true']")[0];
    if (!editor?.isConnected) return "";
    const row = Number(editor.dataset.row);
    const col = Number(editor.dataset.col);
    return row && col ? row + ":" + col : "";
  }

  function rangePixelRect(sheet, range) {
    const rawLeft = coordinateToPixels(sheet, "col", range.left - 1);
    const top = coordinateToPixels(sheet, "row", range.top - 1);
    const rawRight = coordinateToPixels(sheet, "col", range.right);
    const bottom = coordinateToPixels(sheet, "row", range.bottom);
    const left = Math.min(rawLeft, rawRight);
    const right = Math.max(rawLeft, rawRight);
    return { left, top, right, bottom, width: Math.max(0, right - left), height: Math.max(0, bottom - top) };
  }

  function rangeIntersects(a, b) {
    return a.left <= b.right && a.right >= b.left && a.top <= b.bottom && a.bottom >= b.top;
  }

  function paintCanvasGridLines(context, sheet, maps, row, col, range, cell, rect) {
    const rowSpan = range.bottom - range.top + 1;
    const colSpan = range.right - range.left + 1;
    context.save();
    context.strokeStyle = "#d9e0ea";
    context.lineWidth = 1;
    context.beginPath();
    if (!shouldHideGridRight(sheet, maps, row, col, rowSpan, colSpan, cell)) {
      crispCanvasLine(context, rect.right, rect.top, rect.right, rect.bottom);
    }
    if (!shouldHideGridBottom(sheet, maps, row, col, rowSpan, colSpan, cell)) {
      crispCanvasLine(context, rect.left, rect.bottom, rect.right, rect.bottom);
    }
    context.stroke();
    context.restore();
  }

  function paintCanvasCellText(context, sheet, maps, sheetId, row, col, range, cell, rect) {
    const formula = cell.formula || getFormula(sheetId, row, col);
    const display = state.showFormulas && formula ? formula : getValue(sheet, sheetId, row, col);
    if (display == null || display === "") return;
    const text = String(display);
    const css = cellEffectiveCss(sheet, cell);
    const clipWidth = textOverflowWidth(sheet, maps, row, col, range.right - range.left + 1, cell, display) || rect.width;
    let clipRect = { left: rect.left, top: rect.top, width: Math.max(1, clipWidth), height: rect.height };
    const displayScale = canvasFontDisplayScale(sheet);
    if (displayScale !== 1) {
      rect = Object.fromEntries(Object.entries(rect).map(([key, value]) => [key, value / displayScale]));
      clipRect = Object.fromEntries(Object.entries(clipRect).map(([key, value]) => [key, value / displayScale]));
    }
    context.save();
    try {
      if (displayScale !== 1) context.scale(displayScale, displayScale);
      paintCanvasCellTextLayout(context, text, css, cell, rect, clipRect);
    } finally {
      context.restore();
    }
  }

  function paintCanvasCellTextLayout(context, text, css, cell, rect, clipRect) {
    const paddingX = 3;
    const fontScale = 1;
    const fontSize = canvasFontSize(css.fontSize, fontScale);
    const lineHeight = canvasTextLineHeight(fontSize);
    const font = canvasFont(css, fontSize);
    if (css.textRotation === "vertical") {
      paintCanvasVerticalCellText(context, text, css, rect, clipRect, fontSize, font);
      return;
    }
    const isWrapped = css.whiteSpace === "pre-wrap";
    const maxTextWidth = Math.max(1, clipRect.width - paddingX * 2);
    const richTextRuns = !state.showFormulas ? richTextRunsForDisplay(cell, text) : null;
    if (richTextRuns?.length) {
      paintCanvasRichCellText(context, richTextRuns, css, rect, clipRect, fontSize, lineHeight, paddingX, isWrapped, fontScale);
      return;
    }
    const lines = isWrapped ? wrapCanvasText(context, text, maxTextWidth, font) : text.split(/\r\n|\r|\n/).slice(0, 1);
    const phoneticRuns = !state.showFormulas && lines.length === 1 ? phoneticRunsForLine(cell, lines[0]) : [];
    const phoneticFontSize = phoneticRuns.length ? Math.max(7, Math.min(fontSize * 0.58, fontSize - 2)) : 0;
    const phoneticLineHeight = phoneticRuns.length ? Math.ceil(phoneticFontSize * 1.1) : 0;
    const totalTextHeight = Math.max(lineHeight, lines.length * lineHeight + phoneticLineHeight);
    const vertical = css.verticalAlign || "middle";
    let y = rect.top + Math.max(0, (rect.height - totalTextHeight) / 2);
    if (vertical === "top") y = rect.top + 2;
    if (vertical === "bottom") y = rect.bottom - totalTextHeight - 2;

    context.save();
    context.beginPath();
    context.rect(clipRect.left, clipRect.top, clipRect.width, clipRect.height);
    context.clip();
    context.font = font;
    context.fillStyle = css.color || "#1d2433";
    context.textBaseline = "alphabetic";
    const align = css.textAlign || "left";
    context.textAlign = align === "right" ? "right" : align === "center" ? "center" : "left";
    let x = rect.left + paddingX;
    if (context.textAlign === "center") x = rect.left + rect.width / 2;
    if (context.textAlign === "right") x = rect.right - paddingX;
    lines.forEach(function (line, index) {
      const lineY = y + index * lineHeight;
      if (lineY > clipRect.top + clipRect.height) return;
      const mainLineY = lineY + (index === 0 ? phoneticLineHeight : 0);
      const baselineY = canvasLineBaseline(context, line, mainLineY, lineHeight, fontSize);
      if (index === 0 && phoneticRuns.length) paintCanvasPhonetics(context, css, line, x, context.textAlign, lineY, phoneticRuns, phoneticFontSize);
      context.fillText(line, x, baselineY);
      const textDecoration = String(css.textDecoration || "");
      if (textDecoration.includes("underline")) paintCanvasUnderline(context, line, x, baselineY, lineHeight, textDecoration.includes("double"));
      if (textDecoration.includes("line-through")) paintCanvasStrikethrough(context, line, x, baselineY, lineHeight);
    });
    context.restore();
  }

  function richTextRunsForDisplay(cell, displayText) {
    if (!cell?.richText?.length) return null;
    const runs = cell.richText
      .map(function (run) {
        return { text: normalizeRichTextLineEndings(run?.text), css: run?.css && typeof run.css === "object" ? run.css : {} };
      })
      .filter(function (run) {
        return run.text;
      });
    if (!runs.length) return null;
    const richText = normalizeRichTextLineEndings(
      runs
        .map(function (run) {
          return run.text;
        })
        .join(""),
    );
    return richText === normalizeRichTextLineEndings(displayText) ? runs : null;
  }

  function normalizeRichTextLineEndings(value) {
    return String(value ?? "").replace(/\r\n|\r/g, "\n");
  }

  function canvasRichLineMetrics(lines, css, fontScale = 1) {
    const base = canvasFontSize(css.fontSize, fontScale);
    const fonts = lines.map(line => Math.max(base, ...line.map(run => canvasFontSize(run.css?.fontSize || css.fontSize, fontScale))));
    const heights = fonts.map(canvasTextLineHeight);
    return {fonts,heights,totalHeight:heights.reduce((sum,n)=>sum+n,0)};
  }

  function paintCanvasRichCellText(context, runs, css, rect, clipRect, fontSize, lineHeight, paddingX, isWrapped, fontScale) {
    const maxTextWidth = Math.max(1, clipRect.width - paddingX * 2);
    const lines = isWrapped ? wrapCanvasRichText(context, runs, css, maxTextWidth, fontScale || 1) : splitRichRunsIntoHardLines(runs).slice(0, 1);
    const metrics = canvasRichLineMetrics(lines, css, fontScale || 1);
    const totalTextHeight = metrics.totalHeight;
    const vertical = css.verticalAlign || "middle";
    let y = rect.top + Math.max(0, (rect.height - totalTextHeight) / 2);
    if (vertical === "top") y = rect.top + 2;
    if (vertical === "bottom") y = rect.bottom - totalTextHeight - 2;

    context.save();
    context.beginPath();
    context.rect(clipRect.left, clipRect.top, clipRect.width, clipRect.height);
    context.clip();
    context.textBaseline = "alphabetic";
    const align = css.textAlign || "left";
    const textAlign = align === "right" ? "right" : align === "center" ? "center" : "left";
    lines.forEach(function (line, index) {
      const lineY = y + metrics.heights.slice(0,index).reduce((sum,n)=>sum+n,0);
      const actualLineHeight=metrics.heights[index],actualFontSize=metrics.fonts[index];
      if (lineY > clipRect.top + clipRect.height) return;
      const text = richLinePlainText(line);
      context.font = canvasFont(css, actualFontSize);
      const baselineY = canvasLineBaseline(context, text, lineY, actualLineHeight, actualFontSize);
      const lineWidth = measureRichLine(context, line, css, fontScale || 1);
      let x = rect.left + paddingX;
      if (textAlign === "center") x = rect.left + rect.width / 2 - lineWidth / 2;
      if (textAlign === "right") x = rect.right - paddingX - lineWidth;
      drawCanvasRichLine(context, line, css, x, baselineY, actualLineHeight, fontScale || 1);
    });
    context.restore();
  }

  function splitRichRunsIntoHardLines(runs) {
    const lines = [];
    let current = [];
    runs.forEach(function (run) {
      normalizeRichTextLineEndings(run.text)
        .split("\n")
        .forEach(function (part, index, parts) {
          if (part) current.push({ text: part, css: run.css || {} });
          if (index < parts.length - 1) {
            lines.push(current);
            current = [];
          }
        });
    });
    lines.push(current);
    return lines;
  }

  function wrapCanvasRichText(context, runs, css, maxWidth, fontScale) {
    const lines = [];
    splitRichRunsIntoHardLines(runs).forEach(function (hardLine) {
      if (!hardLine.length) {
        lines.push([]);
        return;
      }
      let current = [];
      hardLine.forEach(function (segment) {
        Array.from(segment.text || "").forEach(function (char) {
          const candidate = appendRichChar(current, char, segment.css);
          if (richLinePlainText(current) && measureRichLine(context, candidate, css, fontScale || 1) > maxWidth) {
            lines.push(current);
            current = appendRichChar([], char, segment.css);
          } else {
            current = candidate;
          }
        });
      });
      lines.push(current);
    });
    return lines.slice(0, 200);
  }

  function appendRichChar(line, char, css) {
    const next = line.map(function (segment) {
      return { text: segment.text, css: segment.css };
    });
    const last = next[next.length - 1];
    if (last && richCssKey(last.css) === richCssKey(css || {})) {
      last.text += char;
    } else {
      next.push({ text: char, css: css || {} });
    }
    return next;
  }

  function richCssKey(css) {
    return JSON.stringify(
      Object.keys(css || {})
        .sort()
        .map(function (key) {
          return [key, css[key]];
        }),
    );
  }

  function richLinePlainText(line) {
    return (line || [])
      .map(function (segment) {
        return segment.text || "";
      })
      .join("");
  }

  function richSegmentCss(baseCss, segmentCss) {
    return { ...(baseCss || {}), ...(segmentCss || {}) };
  }

  function measureRichLine(context, line, css, fontScale) {
    return (line || []).reduce(function (width, segment) {
      const mergedCss = richSegmentCss(css, segment.css);
      context.font = canvasFont(mergedCss, canvasFontSize(mergedCss.fontSize, fontScale || 1));
      return width + context.measureText(segment.text || "").width;
    }, 0);
  }

  function drawCanvasRichLine(context, line, css, x, baselineY, lineHeight, fontScale) {
    let cursor = x;
    (line || []).forEach(function (segment) {
      const mergedCss = richSegmentCss(css, segment.css);
      context.font = canvasFont(mergedCss, canvasFontSize(mergedCss.fontSize, fontScale || 1));
      context.fillStyle = mergedCss.color || "#1d2433";
      context.textAlign = "left";
      const text = segment.text || "";
      context.fillText(text, cursor, baselineY);
      const textDecoration = String(mergedCss.textDecoration || "");
      if (textDecoration.includes("underline")) {
        const width = context.measureText(text).width;
        const underlineY = baselineY + Math.max(1, Math.min(3, lineHeight * 0.12));
        const underlineGap = Math.max(2, Math.min(3, lineHeight * 0.16));
        context.save();
        context.strokeStyle = context.fillStyle;
        context.lineWidth = 1;
        context.beginPath();
        context.moveTo(cursor, underlineY);
        context.lineTo(cursor + width, underlineY);
        if (textDecoration.includes("double")) {
          context.moveTo(cursor, underlineY + underlineGap);
          context.lineTo(cursor + width, underlineY + underlineGap);
        }
        context.stroke();
        context.restore();
      }
      if (textDecoration.includes("line-through")) {
        const width = context.measureText(text).width;
        const strikeY = baselineY - Math.max(3, Math.min(lineHeight * 0.34, lineHeight - 2));
        context.save();
        context.strokeStyle = context.fillStyle;
        context.lineWidth = 1;
        context.beginPath();
        context.moveTo(cursor, strikeY);
        context.lineTo(cursor + width, strikeY);
        context.stroke();
        context.restore();
      }
      cursor += context.measureText(text).width;
    });
  }

  function paintCanvasVerticalCellText(context, text, css, rect, clipRect, fontSize, font) {
    const chars = Array.from(String(text).replace(/\r\n|\r|\n/g, ""));
    if (!chars.length) return;
    const advance = Math.max(10, fontSize * 1.08);
    const totalTextHeight = chars.length * advance;
    const vertical = css.verticalAlign || "middle";
    let y = rect.top + Math.max(0, (rect.height - totalTextHeight) / 2);
    if (vertical === "top") y = rect.top + 2;
    if (vertical === "bottom") y = rect.bottom - totalTextHeight - 2;
    const align = css.textAlign || "center";
    let x = rect.left + rect.width / 2;
    if (align === "left") x = rect.left + Math.min(rect.width / 2, fontSize * 0.7 + 2);
    if (align === "right") x = rect.right - Math.min(rect.width / 2, fontSize * 0.7 + 2);

    context.save();
    context.beginPath();
    context.rect(clipRect.left, clipRect.top, clipRect.width, clipRect.height);
    context.clip();
    context.font = font;
    context.fillStyle = css.color || "#1d2433";
    context.textAlign = "center";
    context.textBaseline = "middle";
    chars.forEach(function (char, index) {
      const charY = y + advance * index + advance / 2;
      if (charY < clipRect.top - advance || charY > clipRect.top + clipRect.height + advance) return;
      context.fillText(char, x, charY);
    });
    context.restore();
  }

  function phoneticRunsForLine(cell, line) {
    if (!cell?.showPhonetic || !cell?.phonetics?.length || !line) return [];
    return normalizePhoneticRuns(cell.phonetics, line);
  }

  function paintCanvasPhonetics(context, css, line, x, align, y, runs, fontSize) {
    const textLeft = canvasAlignedTextLeft(context, line, x, align);
    const placements = runs
      .map(function (run) {
        const start = Math.max(0, Math.min(line.length, run.start));
        const end = Math.max(start, Math.min(line.length, run.end));
        const base = line.slice(start, end);
        if (!base) return null;
        const baseLeft = textLeft + context.measureText(line.slice(0, start)).width;
        const baseWidth = context.measureText(base).width;
        return { text: run.text, center: baseLeft + baseWidth / 2 };
      })
      .filter(Boolean);
    if (!placements.length) return;
    const previousFont = context.font;
    const previousAlign = context.textAlign;
    const previousBaseline = context.textBaseline;
    context.font = canvasFont(css, fontSize);
    context.textAlign = "left";
    context.textBaseline = "top";
    placements.forEach(function (placement) {
      const width = context.measureText(placement.text).width;
      context.fillText(placement.text, placement.center - width / 2, y);
    });
    context.font = previousFont;
    context.textAlign = previousAlign;
    context.textBaseline = previousBaseline;
  }

  function canvasAlignedTextLeft(context, line, x, align) {
    const width = context.measureText(line).width;
    if (align === "right" || align === "end") return x - width;
    if (align === "center") return x - width / 2;
    return x;
  }

  function normalizePhoneticRuns(runs, baseText) {
    const textLength = String(baseText || "").length;
    return (runs || [])
      .map(function (run) {
        const start = Math.max(0, Number(run?.start ?? run?.sb ?? 0));
        const rawEnd = Number(run?.end ?? run?.eb ?? start);
        const end = textLength ? Math.min(textLength, Math.max(start, rawEnd)) : Math.max(start, rawEnd);
        return { start, end, text: String(run?.text || "") };
      })
      .filter(function (run) {
        return run.end > run.start && run.text.trim();
      });
  }

  function canvasFontDisplayScale(sheet) {
    const scale = Number(sheet?.displayScale || 1);
    return Number.isFinite(scale) && scale > 0 ? scale : 1;
  }

  function canvasFontSize(fontSize, scale) {
    const value = Number.parseFloat(fontSize);
    const base = !Number.isFinite(value) || value <= 0 ? 13 : String(fontSize || "").trim().endsWith("pt") ? (value * 96) / 72 : value;
    const safeScale = Number.isFinite(Number(scale)) && Number(scale) > 0 ? Number(scale) : 1;
    return Math.max(1, base * safeScale);
  }

  function canvasFont(css, fontSize) {
    const style = css?.fontStyle === "italic" ? "italic " : "";
    const weight = css?.fontWeight ? css.fontWeight + " " : "";
    const family = css?.fontFamily || 'Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
    return style + weight + fontSize + "px " + family;
  }

  function canvasTextLineHeight(fontSize) {
    return Math.max(10, fontSize * 1.2);
  }

  function canvasLineBaseline(context, line, lineTop, lineHeight, fontSize) {
    const metrics = context.measureText(line || "Mg");
    const ascent = Number(metrics.actualBoundingBoxAscent);
    const descent = Number(metrics.actualBoundingBoxDescent);
    const fallbackAscent = fontSize * 0.78;
    const fallbackDescent = fontSize * 0.22;
    const safeAscent = Number.isFinite(ascent) && ascent > 0 ? ascent : fallbackAscent;
    const safeDescent = Number.isFinite(descent) && descent >= 0 ? descent : fallbackDescent;
    const inkHeight = Math.max(1, safeAscent + safeDescent);
    return lineTop + Math.max(0, (lineHeight - inkHeight) / 2) + safeAscent;
  }

  function wrapCanvasText(context, text, maxWidth, font) {
    context.save();
    context.font = font;
    const lines = [];
    String(text)
      .split(/\r\n|\r|\n/)
      .forEach(function (hardLine) {
        if (!hardLine) {
          lines.push("");
          return;
        }
        let current = "";
        Array.from(hardLine).forEach(function (char) {
          const candidate = current + char;
          if (current && context.measureText(candidate).width > maxWidth) {
            lines.push(current);
            current = char;
          } else {
            current = candidate;
          }
        });
        lines.push(current);
      });
    context.restore();
    return lines.slice(0, 200);
  }

  function paintCanvasUnderline(context, line, x, y, lineHeight, doubleUnderline) {
    const width = context.measureText(line).width;
    let start = x;
    if (context.textAlign === "center") start = x - width / 2;
    if (context.textAlign === "right") start = x - width;
    const underlineY = y + Math.max(1, Math.min(3, lineHeight * 0.12));
    const underlineGap = Math.max(2, Math.min(3, lineHeight * 0.16));
    context.save();
    context.strokeStyle = context.fillStyle;
    context.lineWidth = 1;
    context.beginPath();
    context.moveTo(start, underlineY);
    context.lineTo(start + width, underlineY);
    if (doubleUnderline) {
      context.moveTo(start, underlineY + underlineGap);
      context.lineTo(start + width, underlineY + underlineGap);
    }
    context.stroke();
    context.restore();
  }

  function paintCanvasStrikethrough(context, line, x, y, lineHeight) {
    const width = context.measureText(line).width;
    let start = x;
    if (context.textAlign === "center") start = x - width / 2;
    if (context.textAlign === "right") start = x - width;
    const strikeY = y - Math.max(3, Math.min(lineHeight * 0.34, lineHeight - 2));
    context.save();
    context.strokeStyle = context.fillStyle;
    context.lineWidth = 1;
    context.beginPath();
    context.moveTo(start, strikeY);
    context.lineTo(start + width, strikeY);
    context.stroke();
    context.restore();
  }

  function paintCanvasBorders(context, sheet, maps, range) {
    const lines = borderLinesForSheet(sheet, maps, range);
    context.save();
    lines.forEach(function (line) {
      const width = line.border.width || 1;
      context.strokeStyle = line.border.color || "#000000";
      context.lineWidth = width;
      context.setLineDash(canvasBorderDash(line.border.style));
      context.beginPath();
      if (line.orientation === "h") {
        crispCanvasLine(context, line.start, line.position, line.end, line.position);
      } else {
        crispCanvasLine(context, line.position, line.start, line.position, line.end);
      }
      context.stroke();
    });
    context.restore();
  }

  function crispCanvasLine(context, x1, y1, x2, y2) {
    const crispX1 = Math.round(x1) + (Math.abs(x1 - x2) < 0.01 ? 0.5 : 0);
    const crispX2 = Math.round(x2) + (Math.abs(x1 - x2) < 0.01 ? 0.5 : 0);
    const crispY1 = Math.round(y1) + (Math.abs(y1 - y2) < 0.01 ? 0.5 : 0);
    const crispY2 = Math.round(y2) + (Math.abs(y1 - y2) < 0.01 ? 0.5 : 0);
    context.moveTo(crispX1, crispY1);
    context.lineTo(crispX2, crispY2);
  }

  function canvasBorderDash(style) {
    if (style === "dashed") return [6, 4];
    if (style === "dotted") return [1, 3];
    return [];
  }

	  function getValue(sheet, sheetId, row, col) {
	    const cell = sheet.cells[row + ":" + col];
	    if (!cell?.formula) return shouldHideZeroValueForSheet(sheet, cell) ? "" : cellDisplayText(cell);
	    try {
	      const value = hf.getCellValue({ sheet: sheetId, row: row - 1, col: col - 1 });
	      return shouldHideZeroValueForSheet(sheet, cell, value) ? "" : displayValue(value, cell?.numFmt);
	    } catch {
	      return shouldHideZeroValueForSheet(sheet, cell) ? "" : cellDisplayText(cell);
	    }
	  }
	  function shouldHideZeroValueForSheet(sheet, cell, value) {
	    if (!cell || sheet?.showZeros !== false) return false;
	    const candidates = value === undefined ? [cell.raw, cell.cached, cell.display] : [value, cell.raw, cell.cached, cell.display];
	    return candidates.some(function (candidate) {
	      return typeof candidate === "number" && Number.isFinite(candidate) && Math.abs(candidate) < 1e-12;
	    });
	  }
	  function cellDisplayText(cell) {
	    if (!cell) return "";
	    if (cell.display != null && cell.display !== "") return String(cell.display);
	    return implicitCellDisplayText(cell);
	  }
	  function implicitCellDisplayText(cell) {
	    if (!cell) return "";
	    if (cell.formula) return cell.cached != null ? displayValue(cell.cached, cell.numFmt) : "";
	    if (cell.raw != null) return displayValue(cell.raw, cell.numFmt);
	    if (cell.cached != null) return displayValue(cell.cached, cell.numFmt);
	    return "";
	  }

	  function refreshStandaloneWorkbookChartsFromSources() {
	    let changed = false;
	    (model.sheets || []).forEach(function (sheet, sheetIndex) {
	      (sheet.images || []).forEach(function (image) {
	        if (!image?.chart?.series?.length) return;
	        const dataChanged = refreshStandaloneChartDataFromSources(image.chart, sheetIndex);
	        const sourceChanged = refreshStandaloneChartSource(image);
	        if (dataChanged || sourceChanged) changed = true;
	      });
	    });
	    return changed;
	  }

	  function refreshStandaloneChartDataFromSources(chart, ownerSheetIndex) {
	    if (!chart?.series?.length) return false;
	    let changed = false;
	    const title = standaloneChartSingleTextValueForReference(chart.titleRef, ownerSheetIndex);
	    if (title != null && chart.title !== title) {
	      chart.title = title;
	      changed = true;
	    }
	    const labels = standaloneChartTextValuesForReference(chart.labelsRef, ownerSheetIndex);
	    if (labels && !standaloneChartArraysEqual(chart.labels, labels)) {
	      chart.labels = labels;
	      changed = true;
	    }
	    (chart.series || []).forEach(function (series) {
	      const name = standaloneChartSingleTextValueForReference(series.nameRef, ownerSheetIndex);
	      if (name != null && series.name !== name) {
	        series.name = name;
	        changed = true;
	      }
	      const values = standaloneChartNumericValuesForReference(series.valuesRef || series.xValuesRef, ownerSheetIndex);
	      if (values && !standaloneChartArraysEqual(series.values, values)) {
	        series.values = values;
	        changed = true;
	      }
	      const yValues = standaloneChartNumericValuesForReference(series.yValuesRef, ownerSheetIndex);
	      if (yValues && !standaloneChartArraysEqual(series.yValues, yValues)) {
	        series.yValues = yValues;
	        changed = true;
	      }
	    });
	    return changed;
	  }

	  function refreshStandaloneChartSource(image) {
	    if (!image?.chart?.series?.length || typeof chartSvg !== "function" || typeof svgDataUrl !== "function") return false;
	    const width = Math.max(1, Math.round(Number(image.ext?.width) || 520));
	    const height = Math.max(1, Math.round(Number(image.ext?.height) || 320));
	    const nextSrc = svgDataUrl(chartSvg(image.chart, image.chart.type || "column", { width, height, title: image.chart.title || image.name || "グラフ" }));
	    const changed = image.src !== nextSrc || image.mimeType !== "image/svg+xml" || image.extension !== "svg";
	    image.src = nextSrc;
	    image.mimeType = "image/svg+xml";
	    image.extension = "svg";
	    image.name = image.name || image.chart.title || "グラフ";
	    return changed;
	  }

	  function standaloneChartTextValuesForReference(reference, ownerSheetIndex) {
	    const parsed = standaloneChartReferenceRange(reference, ownerSheetIndex);
	    if (!parsed) return null;
	    const sheet = model.sheets?.[parsed.sheetIndex];
	    if (!sheet) return null;
	    const sheetId = hf.getSheetId(sheet.name);
	    const values = [];
	    for (let row = parsed.range.top; row <= parsed.range.bottom; row += 1) {
	      for (let col = parsed.range.left; col <= parsed.range.right; col += 1) {
	        values.push(String(getValue(sheet, sheetId, row, col) ?? ""));
	      }
	    }
	    return values;
	  }

	  function standaloneChartSingleTextValueForReference(reference, ownerSheetIndex) {
	    const values = standaloneChartTextValuesForReference(reference, ownerSheetIndex);
	    return values?.length ? values[0] : null;
	  }

	  function standaloneChartNumericValuesForReference(reference, ownerSheetIndex) {
	    const parsed = standaloneChartReferenceRange(reference, ownerSheetIndex);
	    if (!parsed) return null;
	    const sheet = model.sheets?.[parsed.sheetIndex];
	    if (!sheet) return null;
	    const sheetId = hf.getSheetId(sheet.name);
	    const values = [];
	    for (let row = parsed.range.top; row <= parsed.range.bottom; row += 1) {
	      for (let col = parsed.range.left; col <= parsed.range.right; col += 1) {
	        values.push(standaloneChartCellNumericValue(sheet, sheetId, row, col));
	      }
	    }
	    return values;
	  }

	  function standaloneChartCellNumericValue(sheet, sheetId, row, col) {
	    const cell = sheet?.cells?.[row + ":" + col];
	    let value = null;
	    if (cell?.formula && sheetId != null) {
	      try {
	        value = hf.getCellValue({ sheet: sheetId, row: row - 1, col: col - 1 });
	      } catch {
	        value = null;
	      }
	    }
	    if (value == null) value = cell?.cached ?? cell?.raw ?? getValue(sheet, sheetId, row, col);
	    if (value instanceof Date && !Number.isNaN(value.getTime()) && typeof dateToExcelSerial === "function") return dateToExcelSerial(value);
	    if (typeof value === "object" && value?.value != null) value = value.value;
	    const number = typeof value === "number" ? value : Number(String(value ?? "").replace(/[,¥$€£%]/g, ""));
	    return Number.isFinite(number) ? number : null;
	  }

	  function standaloneChartReferenceRange(reference, ownerSheetIndex) {
	    const parsed = standaloneParseChartReference(reference, ownerSheetIndex);
	    if (!parsed) return null;
	    const range = decodeRange(String(parsed.range || "").replace(/\$/g, ""));
	    if (!range) return null;
	    return { sheetIndex: parsed.sheetIndex, range };
	  }

	  function standaloneParseChartReference(reference, ownerSheetIndex) {
	    let text = String(reference || "").trim();
	    if (!text) return null;
	    if (text.startsWith("=")) text = text.slice(1).trim();
	    if (!text || /[\[\]]/.test(text) || text.includes(",")) return null;
	    let sheetIndex = Number.isInteger(Number(ownerSheetIndex)) ? Number(ownerSheetIndex) : state.active;
	    let rangeText = text;
	    const bangIndex = standaloneChartReferenceBangIndex(text);
	    if (bangIndex >= 0) {
	      const sheetName = standaloneUnquoteSheetName(text.slice(0, bangIndex));
	      const resolvedIndex = standaloneChartSheetIndexByName(sheetName);
	      if (resolvedIndex < 0) return null;
	      sheetIndex = resolvedIndex;
	      rangeText = text.slice(bangIndex + 1);
	    }
	    return /^[\$A-Z0-9:]+$/i.test(rangeText) ? { sheetIndex, range: rangeText } : null;
	  }

	  function standaloneChartReferenceBangIndex(text) {
	    const source = String(text || "");
	    if (!source.includes("!")) return -1;
	    if (source.startsWith("'")) {
	      let index = 1;
	      while (index < source.length) {
	        if (source[index] === "'" && source[index + 1] === "'") {
	          index += 2;
	          continue;
	        }
	        if (source[index] === "'") return source[index + 1] === "!" ? index + 1 : -1;
	        index += 1;
	      }
	      return -1;
	    }
	    return source.lastIndexOf("!");
	  }

	  function standaloneUnquoteSheetName(name) {
	    const text = String(name || "").trim();
	    if (text.startsWith("'") && text.endsWith("'")) return text.slice(1, -1).replace(/''/g, "'");
	    return text;
	  }

	  function standaloneChartSheetIndexByName(sheetName) {
	    const target = String(sheetName || "").trim().toLocaleLowerCase("ja-JP");
	    return (model.sheets || []).findIndex(function (sheet) {
	      return String(sheet.name || "").trim().toLocaleLowerCase("ja-JP") === target;
	    });
	  }

	  function standaloneChartArraysEqual(left, right) {
	    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) return false;
	    return left.every(function (value, index) {
	      const next = right[index];
	      if (value == null || next == null) return value == null && next == null;
	      if (typeof value === "number" || typeof next === "number") return Number(value) === Number(next);
	      return String(value) === String(next);
	    });
	  }

	  function getFormula(sheetId, row, col) {
	    try {
	      return hf.getCellFormula({ sheet: sheetId, row: row - 1, col: col - 1 });
    } catch {
      return null;
    }
  }

  function commit(row, col, value) {
    const sheet = model.sheets[state.active];
    const content = coerce(value, { preserveWhitespace: true });
    hf.setCellContents({ sheet: hf.getSheetId(sheet.name), row: row - 1, col: col - 1 }, [[standaloneFormulaEngineContentForCurrentModel(content, { row, col })]]);
    const key = row + ":" + col;
    sheet.cells[key] = sheet.cells[key] || { row, col, css: {} };
    sheet.cells[key].raw = content;
	    sheet.cells[key].formula = typeof content === "string" && content.startsWith("=") ? content : null;
		    sheet.cells[key].borders = sheet.cells[key].borders || {};
		    applyStandaloneMultilineCellDisplayBehavior(sheet, row, col, sheet.cells[key], content);
		    delete sheet.cells[key].richText;
		    state.selected = { row, col };
		    refreshStandaloneWorkbookChartsFromSources();
		    renderSheet();
		    $status.text(sheet.name + "!" + columnName(col) + row + " を更新しました。");
		  }

	  function applyStandaloneMultilineCellDisplayBehavior(sheet, row, col, cell, content) {
	    if (typeof content !== "string" || content.startsWith("=") || !/\r\n|\r|\n/.test(content) || !cell) return;
	    cell.css = { ...(cell.css || {}), whiteSpace: "pre-wrap", overflowWrap: "anywhere" };
	    const rowModel = sheet.rows[row - 1];
	    if (!rowModel || rowModel.hidden) return;
	    const currentHeight = Number(rowModel.height || sheet.defaultRowHeight || 20);
	    const requiredHeight = standaloneMultilineCellAutoHeight(sheet, row, col, cell, content);
	    if (requiredHeight > currentHeight + 0.5) rowModel.height = requiredHeight;
	  }

	  function standaloneMultilineCellAutoHeight(sheet, row, col, cell, text) {
	    const css = cellEffectiveCss(sheet, cell);
	    const fontSize = canvasFontSize(css.fontSize);
	    const lineHeight = canvasTextLineHeight(fontSize);
	    const range = standaloneMergedRangeForCell(sheet, row, col) || { top: row, left: col, bottom: row, right: col };
	    const rect = rangePixelRect(sheet, range);
	    const maxTextWidth = Math.max(1, rect.width - 6);
	    let lineCount = String(text).split(/\r\n|\r|\n/).length;
	    const context = typeof document === "undefined" ? null : document.createElement("canvas").getContext?.("2d");
	    if (context) lineCount = Math.max(lineCount, wrapCanvasText(context, text, maxTextWidth, canvasFont(css, fontSize)).length);
	    return Math.ceil(lineCount * lineHeight + 4);
	  }

	  function standaloneMergedRangeForCell(sheet, row, col) {
	    return (sheet.merges || []).find(function (merge) {
	      return row >= merge.top && row <= merge.bottom && col >= merge.left && col <= merge.right;
	    }) || null;
	  }

	  function standaloneFormulaEngineContentForCurrentModel(content, context) {
    if (typeof content === "string" && content.startsWith("=")) {
      return translateStandaloneFormulaForFormulaEngine(content, {
        definedNameAliases,
        sheetIndex: state.active,
        row: context?.row,
        col: context?.col,
      });
    }
    return content;
  }

  function select($cell) {
    state.selected = { row: Number($cell.data("row")), col: Number($cell.data("col")) };
    state.selectedRange = {
      top: state.selected.row,
      left: state.selected.col,
      bottom: state.selected.row,
      right: state.selected.col,
    };
    state.selectedRanges = [state.selectedRange];
    $sheet.find(".sheet-cell.selected").removeClass("selected");
    $cell.addClass("selected");
    $label.text(columnName(state.selected.col) + state.selected.row);
    const cell = model.sheets[state.active].cells[state.selected.row + ":" + state.selected.col];
    $formula.prop("disabled", false).val(cell?.formula || (cell?.raw == null ? "" : String(cell.raw)));
    $apply.prop("disabled", false);
  }
  function handleStandaloneCellHyperlinkClick(event) {
    const hyperlink = event.currentTarget?.dataset?.hyperlink || event.currentTarget?.getAttribute?.("href") || "";
    if (!openStandaloneHyperlinkTarget(hyperlink)) return;
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
  }
  function openStandaloneHyperlinkTarget(hyperlink) {
    const text = String(hyperlink || "").trim();
    if (!text) return false;
    if (openStandaloneInternalHyperlink(text)) {
      $status.text("リンク先へ移動しました：" + text);
      return true;
    }
    window.open(text, "_blank", "noopener,noreferrer");
    $status.text("リンクを開きました：" + text);
    return true;
  }
  function standaloneHyperlinkFromCellClickEvent(event) {
    if (event.defaultPrevented || event.button > 0) return "";
    const target = event.target;
    if (!target?.closest) return "";
    if (target.closest(".cell-validation-dropdown-button")) return "";
    const cellElement = target.closest(".sheet-cell[data-row][data-col]") || standaloneCellElementFromClientPoint(event.clientX, event.clientY);
    if (!cellElement) return "";
    if (cellElement.matches?.(".sheet-cell[contenteditable='true']")) return "";
    const link = cellElement.querySelector(".cell-link");
    if (!link || !standalonePointIsInsideElementClientRects(link, event.clientX, event.clientY)) return "";
    return link.dataset?.hyperlink || link.getAttribute("href") || "";
  }
  function standaloneCellElementFromClientPoint(clientX, clientY) {
    const hitElements = typeof document.elementsFromPoint === "function"
      ? document.elementsFromPoint(clientX, clientY)
      : [];
    return Array.prototype.find.call(hitElements, function (element) {
      return element.matches?.(".sheet-cell[data-row][data-col]");
    }) || null;
  }
  function standalonePointIsInsideElementClientRects(element, clientX, clientY) {
    if (!element || !Number.isFinite(clientX) || !Number.isFinite(clientY)) return false;
    const rects = element.getClientRects?.();
    if (!rects?.length) return false;
    return Array.prototype.some.call(rects, function (rect) {
      return clientX >= rect.left - 1 &&
        clientX <= rect.right + 1 &&
        clientY >= rect.top - 1 &&
        clientY <= rect.bottom + 1;
    });
  }
  function openStandaloneInternalHyperlink(hyperlink) {
    const text = String(hyperlink || "").trim();
    if (!standaloneIsInternalHyperlinkValue(text)) return false;
    const target = standaloneParseInternalHyperlinkTarget(standaloneStripInternalHyperlinkPrefix(text));
    if (!target) return false;
    const sheetIndex = target.sheetName
      ? model.sheets.findIndex(function (sheet) { return sheet.name === target.sheetName && sheet.hidden !== true; })
      : state.active;
    if (sheetIndex < 0 || !model.sheets[sheetIndex]) return false;
    if (sheetIndex !== state.active) {
      state.active = sheetIndex;
      refreshStandaloneWorkbookChartsFromSources();
      renderTabs();
      renderSheet();
    }
    const $cell = $sheet.find(".sheet-cell[data-row='" + target.row + "'][data-col='" + target.col + "']");
    if ($cell.length) {
      select($cell);
      scrollStandaloneCellIntoView(model.sheets[state.active], target.row, target.col);
      $cell[0]?.focus?.({ preventScroll: true });
      return true;
    }
    state.selected = { row: target.row, col: target.col };
    state.selectedRange = { top: target.row, left: target.col, bottom: target.row, right: target.col };
    state.selectedRanges = [state.selectedRange];
    renderSheet();
    scrollStandaloneCellIntoView(model.sheets[state.active], target.row, target.col);
    return true;
  }
  function standaloneIsInternalHyperlinkValue(hyperlink) {
    const text = String(hyperlink || "").trim();
    if (!text) return false;
    if (text.startsWith("#")) return true;
    if (/^[a-z][a-z0-9+.-]*:/i.test(text)) return false;
    return Boolean(standaloneParseInternalHyperlinkTarget(standaloneStripInternalHyperlinkPrefix(text)));
  }
  function standaloneStripInternalHyperlinkPrefix(hyperlink) {
    try {
      return decodeURIComponent(String(hyperlink || "").replace(/^#/, ""));
    } catch {
      return String(hyperlink || "").replace(/^#/, "");
    }
  }
  function standaloneParseInternalHyperlinkTarget(target) {
    const text = String(target || "").trim();
    if (!text) return null;
    const direct = standaloneParseDirectInternalHyperlinkTarget(text);
    if (direct) return direct;
    return standaloneDefinedNameInternalHyperlinkTarget(text);
  }
  function standaloneParseDirectInternalHyperlinkTarget(target) {
    const text = String(target || "").trim();
    const bangIndex = text.lastIndexOf("!");
    const sheetName = bangIndex >= 0 ? standaloneUnquoteSheetName(text.slice(0, bangIndex)) : "";
    const address = (bangIndex >= 0 ? text.slice(bangIndex + 1) : text).split(":")[0];
    const cell = decodeCell(address);
    return cell ? { sheetName, row: cell.row, col: cell.col } : null;
  }
  function standaloneDefinedNameInternalHyperlinkTarget(target) {
    const reference = standaloneDefinedNameReferenceFromTarget(target, state.active);
    if (!Number.isInteger(reference?.sheetIndex) || !reference.range) return null;
    const sheet = model.sheets[reference.sheetIndex];
    return sheet ? { sheetName: sheet.name, row: reference.range.top, col: reference.range.left } : null;
  }
  function standaloneDefinedNameReferenceFromTarget(target, baseSheetIndex = state.active, visitedNames = new Set()) {
    const definition = standaloneDefinedNameFromTarget(target, baseSheetIndex);
    if (!definition) return null;
    return standaloneDefinedNameReference(definition, baseSheetIndex, visitedNames);
  }
  function standaloneDefinedNameFromTarget(target, baseSheetIndex = state.active) {
    const text = String(target || "").trim();
    if (!text || standaloneParseDirectInternalHyperlinkTarget(text)) return null;
    const bangIndex = text.lastIndexOf("!");
    if (bangIndex >= 0) {
      const sheetName = standaloneUnquoteSheetName(text.slice(0, bangIndex));
      const name = text.slice(bangIndex + 1).trim();
      const sheetIndex = standaloneHyperlinkSheetIndexByName(sheetName);
      if (sheetIndex < 0 || !name || standaloneDecodeA1Range(name.replace(/\$/g, ""))) return null;
      const local = (model.definedNames || []).find(function (definition) {
        return !definition.hidden &&
          definition.scope === "sheet" &&
          Number(definition.sheetIndex) === sheetIndex &&
          standaloneNormalizeDefinedNameLookup(definition.name) === standaloneNormalizeDefinedNameLookup(name);
      });
      if (local) return local;
      return (model.definedNames || []).find(function (definition) {
        return !definition.hidden &&
          definition.scope !== "sheet" &&
          standaloneNormalizeDefinedNameLookup(definition.name) === standaloneNormalizeDefinedNameLookup(name);
      }) || null;
    }
    const normalized = standaloneNormalizeDefinedNameLookup(text);
    const local = (model.definedNames || []).find(function (definition) {
      return !definition.hidden &&
        definition.scope === "sheet" &&
        Number(definition.sheetIndex) === baseSheetIndex &&
        standaloneNormalizeDefinedNameLookup(definition.name) === normalized;
    });
    if (local) return local;
    return (model.definedNames || []).find(function (definition) {
      return !definition.hidden &&
        definition.scope !== "sheet" &&
        standaloneNormalizeDefinedNameLookup(definition.name) === normalized;
    }) || null;
  }
  function standaloneDefinedNameReference(definition, baseSheetIndex = state.active, visitedNames = new Set()) {
    if (!definition?.refersTo) return null;
    const key = [
      definition.scope || "workbook",
      definition.sheetIndex == null ? "" : definition.sheetIndex,
      standaloneNormalizeDefinedNameLookup(definition.name),
    ].join(":");
    if (visitedNames.has(key)) return null;
    visitedNames.add(key);
    const sheetIndex = definition.scope === "sheet" && Number.isInteger(Number(definition.sheetIndex))
      ? Number(definition.sheetIndex)
      : baseSheetIndex;
    const firstReference = standaloneSplitFormulaArguments(String(definition.refersTo || "").replace(/^=/, ""))[0]?.trim() || "";
    return standaloneReferenceFromFormula(sheetIndex, firstReference, visitedNames);
  }
  function standaloneReferenceFromFormula(baseSheetIndex, formula, visitedNames = new Set()) {
    const text = String(formula || "").trim().replace(/^=/, "").trim();
    if (!text) return null;
    const direct = standaloneDirectReferenceFromFormula(baseSheetIndex, text);
    if (direct) return direct;
    return standaloneDefinedNameReferenceFromTarget(text, baseSheetIndex, visitedNames);
  }
  function standaloneDirectReferenceFromFormula(baseSheetIndex, formula) {
    const text = String(formula || "").trim();
    const bangIndex = text.lastIndexOf("!");
    const sheetIndex = bangIndex >= 0
      ? standaloneHyperlinkSheetIndexByName(standaloneUnquoteSheetName(text.slice(0, bangIndex)))
      : baseSheetIndex;
    if (sheetIndex < 0 || !model.sheets[sheetIndex]) return null;
    const rangeText = (bangIndex >= 0 ? text.slice(bangIndex + 1) : text).replace(/\$/g, "");
    const range = standaloneDecodeA1Range(rangeText);
    return range ? { sheetIndex, range } : null;
  }
  function standaloneHyperlinkSheetIndexByName(sheetName) {
    const target = String(sheetName || "").trim().toLocaleLowerCase("ja-JP");
    return (model.sheets || []).findIndex(function (sheet) {
      return sheet.hidden !== true && String(sheet.name || "").trim().toLocaleLowerCase("ja-JP") === target;
    });
  }
  function standaloneNormalizeDefinedNameLookup(name) {
    return String(name || "").trim().toLocaleLowerCase("ja-JP");
  }
  function scrollStandaloneCellIntoView(sheet, row, col) {
    const host = $sheet[0];
    if (!host) return;
    const rect = rangePixelRect(sheet, { top: row, left: col, bottom: row, right: col });
    if (!rect) return;
    const margin = 24;
    if (rect.left < host.scrollLeft + margin) {
      host.scrollLeft = Math.max(0, rect.left - margin);
    } else if (rect.left + rect.width > host.scrollLeft + host.clientWidth - margin) {
      host.scrollLeft = Math.max(0, rect.left + rect.width - host.clientWidth + margin);
    }
    if (rect.top < host.scrollTop + margin) {
      host.scrollTop = Math.max(0, rect.top - margin);
    } else if (rect.top + rect.height > host.scrollTop + host.clientHeight - margin) {
      host.scrollTop = Math.max(0, rect.top + rect.height - host.clientHeight + margin);
    }
    if ($hscroll[0]) $hscroll[0].scrollLeft = host.scrollLeft;
    scheduleCanvasPaint();
  }
  function handleStandaloneKeyboardShortcut(event) {
    if (!isStandaloneWorkbookShortcutTarget(event.target)) return;
    const key = String(event.key || "").toLowerCase();
    const dateTimeShortcut = keyboardDateTimeShortcutKind(event, key);
    if (dateTimeShortcut && state.selected) {
      event.preventDefault();
      event.stopPropagation();
      beginStandaloneDateTimeEdit(dateTimeShortcut);
      return;
    }
    if ((event.ctrlKey || event.metaKey) && !event.altKey && key === "a") {
      event.preventDefault();
      event.stopPropagation();
      selectAllStandaloneByKeyboard();
    }
  }
  function isStandaloneWorkbookShortcutTarget(target) {
    const element = target?.nodeType === Node.TEXT_NODE ? target.parentElement : target;
    if (!element) return true;
    if (element.closest?.("#wx-zoom-dialog")) return false;
    if (element.closest?.("[contenteditable='true'],input,textarea,select,button")) return false;
    return true;
  }
  function selectAllStandaloneByKeyboard() {
    const sheet = model.sheets[state.active];
    if (!sheet) return;
    const activeRanges = standaloneActiveSelectionRanges();
    const activeRange = activeRanges.length === 1 ? activeRanges[0] : null;
    const activePoint = standaloneSelectAllActivePoint(activeRange);
    const candidates = standaloneSelectAllExpansionCandidates(sheet, activePoint);
    const target = standaloneSelectAllNextCandidate(candidates, activeRanges.length > 1 ? standaloneWorksheetSelectionRange(sheet) : activeRange);
    if (!target) return;
    standaloneSelectRangesByKeyboard([target.range], target.message, activePoint);
  }
  function standaloneSelectAllActivePoint(activeRange) {
    const sheet = model.sheets[state.active];
    if (state.selected) return state.selected;
    const fallback = standaloneNormalizeRange(sheet, activeRange || state.selectedRange) || { top: 1, left: 1 };
    return { row: fallback.top, col: fallback.left };
  }
  function standaloneSelectAllExpansionCandidates(sheet, activePoint) {
    const candidates = [
      ...standaloneSelectAllTableCandidates(sheet, activePoint),
      standaloneSelectAllCurrentRegionCandidate(sheet, activePoint),
      standaloneSelectAllUsedRangeCandidate(sheet),
      {
        kind: "worksheet",
        range: standaloneWorksheetSelectionRange(sheet),
        message: "ワークシート全体を選択しました。",
      },
    ].filter(Boolean);
    return standaloneUniqueSelectAllCandidates(candidates);
  }
  function standaloneSelectAllTableCandidates(sheet, activePoint) {
    if (!sheet?.tables?.length || !activePoint) return [];
    const table = (sheet.tables || [])
      .filter(function (item) {
        return item?.range && standaloneCellInRange(item.range, activePoint.row, activePoint.col);
      })
      .sort(function (left, right) {
        return standaloneRangeCellCount(left.range) - standaloneRangeCellCount(right.range);
      })[0];
    if (!table?.range) return [];
    const ranges = [];
    const dataRange = standaloneTableDataBodySelectionRange(table);
    if (dataRange && standaloneCellInRange(dataRange, activePoint.row, activePoint.col) && !standaloneRangesEqualLoose(dataRange, table.range)) {
      ranges.push({
        kind: "table-data",
        range: dataRange,
        message: standaloneRangeLabel(dataRange) + " のテーブル データを選択しました。",
      });
    }
    ranges.push({
      kind: "table",
      range: { ...table.range },
      message: standaloneRangeLabel(table.range) + " のテーブル全体を選択しました。",
    });
    return ranges;
  }
  function standaloneTableDataBodySelectionRange(table) {
    const range = table?.range;
    if (!range) return null;
    const top = range.top + (table.showHeaderRow === false ? 0 : 1);
    const bottom = table.showTotalsRow === true ? range.bottom - 1 : range.bottom;
    if (top > bottom) return null;
    return { top, left: range.left, bottom, right: range.right };
  }
  function standaloneSelectAllCurrentRegionCandidate(sheet, activePoint) {
    const region = standaloneCurrentRegionFromPoint(sheet, activePoint);
    if (!region) return null;
    const blankSingleton = standaloneIsSingleCellRange(region) && !standaloneCellHasDataForRegion(sheet, activePoint.row, activePoint.col);
    if (blankSingleton) return null;
    return {
      kind: "current-region",
      range: region,
      message: standaloneRangeLabel(region) + " を選択しました。",
    };
  }
  function standaloneSelectAllUsedRangeCandidate(sheet) {
    const range = standaloneSheetUsedRangeForSelectAll(sheet);
    if (!range) return null;
    return {
      kind: "used-range",
      range,
      message: standaloneRangeLabel(range) + " の使用範囲を選択しました。",
    };
  }
  function standaloneSheetUsedRangeForSelectAll(sheet) {
    const ranges = [];
    const used = standaloneSheetUsedRange(sheet);
    if (used) ranges.push(used);
    (sheet?.tables || []).forEach(function (table) {
      if (table?.range) ranges.push(table.range);
    });
    if (sheet?.autoFilter?.range) ranges.push(sheet.autoFilter.range);
    if (!ranges.length) return null;
    return standaloneExpandRangeForMerges(sheet, standaloneZoomSelectionUnionRanges(ranges));
  }
  function standaloneWorksheetSelectionRange(sheet) {
    return {
      top: 1,
      left: 1,
      bottom: Math.max(1, Number(sheet?.rowCount) || (sheet?.rows || []).length || 1),
      right: Math.max(1, Number(sheet?.colCount) || (sheet?.columns || []).length || 1),
    };
  }
  function standaloneSelectAllNextCandidate(candidates, activeRange) {
    const worksheet = candidates.find(function (candidate) {
      return candidate.kind === "worksheet";
    }) || candidates[candidates.length - 1] || null;
    if (!activeRange) return candidates[0] || worksheet;
    const normalizedActive = standaloneNormalizeRange(model.sheets[state.active], activeRange);
    const next = candidates.find(function (candidate) {
      return !standaloneRangesEqualLoose(candidate.range, normalizedActive) && standaloneRangeContains(candidate.range, normalizedActive);
    });
    if (next) return next;
    return worksheet && !standaloneRangesEqualLoose(worksheet.range, normalizedActive) ? worksheet : worksheet;
  }
  function standaloneUniqueSelectAllCandidates(candidates) {
    const seen = new Set();
    return (candidates || []).filter(function (candidate) {
      if (!candidate?.range) return false;
      const range = candidate.range;
      const key = range.top + ":" + range.left + ":" + range.bottom + ":" + range.right;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }
  function standaloneCurrentRegionFromPoint(sheet, point) {
    const used = standaloneSheetUsedRange(sheet);
    if (!used || !point) return null;
    if (point.row < used.top || point.row > used.bottom || point.col < used.left || point.col > used.right) {
      return { top: point.row, left: point.col, bottom: point.row, right: point.col };
    }
    let region = { top: point.row, left: point.col, bottom: point.row, right: point.col };
    let changed = true;
    while (changed) {
      changed = false;
      if (region.left > used.left && standaloneRangeColumnHasData(sheet, region.left - 1, region.top, region.bottom)) {
        region.left -= 1;
        changed = true;
      }
      if (region.right < used.right && standaloneRangeColumnHasData(sheet, region.right + 1, region.top, region.bottom)) {
        region.right += 1;
        changed = true;
      }
      if (region.top > used.top && standaloneRangeRowHasData(sheet, region.top - 1, region.left, region.right)) {
        region.top -= 1;
        changed = true;
      }
      if (region.bottom < used.bottom && standaloneRangeRowHasData(sheet, region.bottom + 1, region.left, region.right)) {
        region.bottom += 1;
        changed = true;
      }
    }
    return standaloneExpandRangeForMerges(sheet, region);
  }
  function standaloneRangeRowHasData(sheet, row, left, right) {
    for (let col = left; col <= right; col += 1) {
      if (standaloneCellHasDataForRegion(sheet, row, col)) return true;
    }
    return false;
  }
  function standaloneRangeColumnHasData(sheet, col, top, bottom) {
    for (let row = top; row <= bottom; row += 1) {
      if (standaloneCellHasDataForRegion(sheet, row, col)) return true;
    }
    return false;
  }
  function standaloneCellHasDataForRegion(sheet, row, col) {
    const cell = sheet?.cells?.[cellKey(row, col)];
    return Boolean(cell?.formula || cell?.richText?.length || String(cell?.display ?? cell?.cached ?? cell?.raw ?? "").trim());
  }
  function standaloneExpandRangeForMerges(sheet, range) {
    if (!sheet || !range) return range;
    let expanded = { ...range };
    let changed = true;
    while (changed) {
      changed = false;
      (sheet.merges || []).forEach(function (merge) {
        if (!merge || !rangeIntersects(expanded, merge)) return;
        const next = {
          top: Math.min(expanded.top, merge.top),
          left: Math.min(expanded.left, merge.left),
          bottom: Math.max(expanded.bottom, merge.bottom),
          right: Math.max(expanded.right, merge.right),
        };
        if (!standaloneRangesEqualLoose(expanded, next)) {
          expanded = next;
          changed = true;
        }
      });
    }
    return expanded;
  }
  function standaloneSelectRangesByKeyboard(ranges, message, activePoint) {
    const sheet = model.sheets[state.active];
    const normalized = (ranges || [])
      .map(function (range) {
        return standaloneNormalizeRange(sheet, range);
      })
      .filter(Boolean);
    if (!normalized.length) return;
    const primary = normalized[0];
    const nextActive = activePoint && standaloneCellInRange(primary, activePoint.row, activePoint.col)
      ? { row: activePoint.row, col: activePoint.col }
      : { row: primary.top, col: primary.left };
    state.selected = nextActive;
    state.selectedRange = primary;
    state.selectedRanges = normalized;
    renderSheet();
    $label.text(standaloneRangeLabel(primary));
    const cell = sheet.cells[nextActive.row + ":" + nextActive.col];
    $formula.prop("disabled", false).val(cell?.formula || (cell?.raw == null ? "" : String(cell.raw)));
    $apply.prop("disabled", false);
    $status.text(message);
    const activeCell = $sheet.find(".sheet-cell[data-row='" + nextActive.row + "'][data-col='" + nextActive.col + "']")[0];
    activeCell?.focus?.({ preventScroll: true });
  }
  function standaloneCellInRange(range, row, col) {
    return Boolean(range && row >= range.top && row <= range.bottom && col >= range.left && col <= range.right);
  }
  function standaloneRangeContains(container, item) {
    return Boolean(container && item && item.top >= container.top && item.bottom <= container.bottom && item.left >= container.left && item.right <= container.right);
  }
  function standaloneRangesEqualLoose(a, b) {
    return Boolean(a && b && a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right);
  }
  function standaloneIsSingleCellRange(range) {
    return Boolean(range && range.top === range.bottom && range.left === range.right);
  }
  function standaloneRangeCellCount(range) {
    if (!range) return 0;
    return Math.max(1, range.bottom - range.top + 1) * Math.max(1, range.right - range.left + 1);
  }
  function beginCellEdit($cell, options = {}) {
    if (!$cell.length) return;
    select($cell);
    const row = Number($cell.data("row"));
    const col = Number($cell.data("col"));
    const sheet = model.sheets[state.active];
    const cell = sheet.cells[row + ":" + col];
    const beforeValue = cell?.formula || (cell?.raw == null ? "" : String(cell.raw));
    const editValue = options.replaceText
      ? String(options.initialText ?? "")
      : beforeValue;
    $cell.data("before", beforeValue);
    $cell.attr("contenteditable", "true").text(editValue).focus();
    syncEditableCellAlignment($cell[0]);
    scheduleCanvasPaint();
    placeCaretAtEnd($cell[0]);
  }
  function beginStandaloneDateTimeEdit(kind) {
    const point = state.selected;
    if (!point) return false;
    const value = keyboardDateTimeShortcutText(kind, new Date());
    if (!value) return false;
    const selector = ".sheet-cell[data-row='" + point.row + "'][data-col='" + point.col + "']";
    let $cell = $sheet.find(selector);
    if (!$cell.length) {
      scrollStandaloneCellIntoView(model.sheets[state.active], point.row, point.col);
      renderSheet();
      $cell = $sheet.find(selector);
    }
    if (!$cell.length) return false;
    beginCellEdit($cell, { initialText: value, replaceText: true });
    $status.text((kind === "time" ? "時刻" : "日付") + "を入力しました。");
    return true;
  }
  const CELL_EDIT_MIN_PADDING_TOP = 2;
  const CELL_EDIT_PADDING_BOTTOM = 2;
  function syncEditableCellAlignment(cellElement) {
    if (!cellElement || cellElement.getAttribute("contenteditable") !== "true") return;
    cellElement.style.setProperty("--cell-edit-padding-top", CELL_EDIT_MIN_PADDING_TOP + "px");
    const overlay = editableCellOverlayMetrics(cellElement);
    const paddingTop = editableCellPaddingTop(cellElement, overlay);
    const shiftedTextBottom = Number.isFinite(overlay.textBottom)
      ? overlay.textBottom + paddingTop - CELL_EDIT_MIN_PADDING_TOP
      : overlay.height;
    cellElement.style.setProperty("--cell-edit-overlay-left", overlay.left + "px");
    cellElement.style.setProperty("--cell-edit-overlay-width", overlay.width + "px");
    cellElement.style.removeProperty("--cell-edit-overlay-height");
    cellElement.style.setProperty("--cell-edit-overlay-height", Math.ceil(Math.max(overlay.baseHeight || overlay.height, shiftedTextBottom + CELL_EDIT_PADDING_BOTTOM)) + "px");
    cellElement.style.setProperty("--cell-edit-padding-top", paddingTop + "px");
    cellElement.scrollTop = 0;
  }
  function editableCellOverlayHeight(cellElement) {
    return editableCellOverlayMetrics(cellElement).height;
  }
  function editableCellOverlayMetrics(cellElement) {
    const cellRect = cellElement.getBoundingClientRect();
    const baseHeight = Math.max(cellElement.offsetHeight || 0, cellElement.clientHeight || 0, cellRect.height || 0);
    const baseWidth = Math.max(cellElement.offsetWidth || 0, cellElement.clientWidth || 0, cellRect.width || 0);
    const range = document.createRange();
    range.selectNodeContents(cellElement);
    const rects = Array.prototype.slice.call(range.getClientRects()).filter((rect) => rect.width > 0 || rect.height > 0);
    if (range.detach) range.detach();
    if (!rects.length) return { left: 0, width: Math.ceil(baseWidth), height: Math.ceil(baseHeight), baseHeight: Math.ceil(baseHeight) };
    const textLeft = Math.min.apply(null, rects.map((rect) => rect.left)) - cellRect.left;
    const textRight = Math.max.apply(null, rects.map((rect) => rect.right)) - cellRect.left;
    const textTop = Math.min.apply(null, rects.map((rect) => rect.top)) - cellRect.top;
    const textBottom = Math.max.apply(null, rects.map((rect) => rect.bottom)) - cellRect.top;
    const left = Math.min(0, Math.floor(textLeft) - 3);
    const right = Math.max(baseWidth, Math.ceil(textRight) + 3);
    return {
      left,
      width: Math.ceil(right - left),
      height: Math.ceil(Math.max(baseHeight, textBottom + 3)),
      baseHeight: Math.ceil(baseHeight),
      textTop,
      textBottom,
    };
  }
  function editableCellPaddingTop(cellElement, overlay) {
    const baseHeight = Number(overlay?.baseHeight || overlay?.height || 0);
    const textTop = Number(overlay?.textTop);
    const textBottom = Number(overlay?.textBottom);
    const textHeight = Number.isFinite(textTop) && Number.isFinite(textBottom) ? Math.max(0, textBottom - textTop) : 0;
    if (!baseHeight || !textHeight || baseHeight <= textHeight + CELL_EDIT_MIN_PADDING_TOP + CELL_EDIT_PADDING_BOTTOM) {
      return CELL_EDIT_MIN_PADDING_TOP;
    }
    const vertical = editableCellVerticalAlignment(cellElement);
    if (vertical === "bottom") {
      return Math.max(CELL_EDIT_MIN_PADDING_TOP, Math.floor(CELL_EDIT_MIN_PADDING_TOP + baseHeight - textHeight - CELL_EDIT_PADDING_BOTTOM - textTop));
    }
    if (vertical === "middle") {
      return Math.max(CELL_EDIT_MIN_PADDING_TOP, Math.floor(CELL_EDIT_MIN_PADDING_TOP + (baseHeight - textHeight) / 2 - textTop));
    }
    return CELL_EDIT_MIN_PADDING_TOP;
  }
  function editableCellVerticalAlignment(cellElement) {
    const computed = getComputedStyle(cellElement);
    const alignItems = String(computed.alignItems || cellElement.style.alignItems || "").toLowerCase();
    const verticalAlign = String(computed.verticalAlign || cellElement.style.verticalAlign || "").toLowerCase();
    if (alignItems === "flex-start" || verticalAlign === "top") return "top";
    if (alignItems === "flex-end" || verticalAlign === "bottom") return "bottom";
    return "middle";
  }
  function placeCaretAtEnd(element) {
    const selection = window.getSelection?.();
    if (!selection || !element) return;
    const range = document.createRange();
    range.selectNodeContents(element);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  }
  function insertLineBreakInEditable(element) {
    const selection = window.getSelection?.();
    if (!selection || !element) return false;
    let range = selection.rangeCount ? selection.getRangeAt(0) : null;
    if (!range || !element.contains(range.startContainer) || !element.contains(range.endContainer)) {
      placeCaretAtEnd(element);
      range = selection.rangeCount ? selection.getRangeAt(0) : null;
    }
    if (!range) return false;
    range.deleteContents();
    const lineBreak = document.createTextNode("\n");
    const sentinel = document.createTextNode(CELL_EDIT_TRAILING_LINE_SENTINEL);
    const fragment = document.createDocumentFragment();
    fragment.appendChild(lineBreak);
    fragment.appendChild(sentinel);
    range.insertNode(fragment);
    range.setStartBefore(sentinel);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
    try {
      element.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertLineBreak" }));
    } catch {
      element.dispatchEvent(new Event("input", { bubbles: true }));
    }
    return true;
  }
  function insertTextIntoStandaloneEditable(element, text) {
    const selection = window.getSelection?.();
    if (!selection || !element) return false;
    let range = selection.rangeCount ? selection.getRangeAt(0) : null;
    if (!range || !element.contains(range.startContainer) || !element.contains(range.endContainer)) {
      placeCaretAtEnd(element);
      range = selection.rangeCount ? selection.getRangeAt(0) : null;
    }
    if (!range) return false;
    range.deleteContents();
    const node = document.createTextNode(String(text ?? ""));
    range.insertNode(node);
    range.setStartAfter(node);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
    try {
      element.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: String(text ?? "") }));
    } catch {
      element.dispatchEvent(new Event("input", { bubbles: true }));
    }
    return true;
  }
  function insertStandaloneDateTimeIntoEditable(element, kind) {
    const value = keyboardDateTimeShortcutText(kind, new Date());
    if (!value) return false;
    const inserted = insertTextIntoStandaloneEditable(element, value);
    if (!inserted) return false;
    syncEditableCellAlignment(element);
    scheduleCanvasPaint();
    $status.text((kind === "time" ? "時刻" : "日付") + "を挿入しました。");
    return true;
  }
  function editablePlainText(element) {
    return stripCellEditSentinels(element?.textContent || "");
  }
  function stripCellEditSentinels(text) {
    return String(text ?? "").replaceAll(CELL_EDIT_TRAILING_LINE_SENTINEL, "");
  }
  function insertLineBreakInTextControl(element) {
    if (!element) return false;
    const value = String(element.value ?? "");
    const start = Math.max(0, Math.min(value.length, Number(element.selectionStart ?? value.length)));
    const end = Math.max(start, Math.min(value.length, Number(element.selectionEnd ?? start)));
    element.value = value.slice(0, start) + "\n" + value.slice(end);
    const caret = start + 1;
    element.setSelectionRange?.(caret, caret);
    try {
      element.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertLineBreak" }));
    } catch {
      element.dispatchEvent(new Event("input", { bubbles: true }));
    }
    return true;
  }
  function startRuntimeEditDrag(event) {
    if (event.button !== 0) return;
    state.editDrag = {
      cell: event.currentTarget,
      pointerId: event.pointerId,
      x: event.clientX,
      y: event.clientY,
      outsideCell: false,
    };
  }
  function updateRuntimeEditDrag(event) {
    const drag = state.editDrag;
    if (!drag) return;
    if (drag.pointerId != null && event.pointerId != null && event.pointerId !== drag.pointerId) return;
    if (event.buttons !== undefined && (event.buttons & 1) !== 1) {
      finishRuntimeEditDrag(event);
      return;
    }
    drag.x = event.clientX;
    drag.y = event.clientY;
    drag.outsideCell = !pointIsInsideElementRect(drag.cell, event.clientX, event.clientY);
  }
  function finishRuntimeEditDrag(event) {
    const drag = state.editDrag;
    if (
      drag &&
      (!event || drag.pointerId == null || event.pointerId == null || event.pointerId === drag.pointerId)
    ) {
      const x = Number.isFinite(Number(event?.clientX)) ? event.clientX : drag.x;
      const y = Number.isFinite(Number(event?.clientY)) ? event.clientY : drag.y;
      if (drag.outsideCell || !pointIsInsideElementRect(drag.cell, x, y)) {
        state.suppressCellSelectionUntil = nowMs() + 140;
      }
    }
    state.editDrag = null;
  }
  function pointIsInsideElementRect(element, clientX, clientY) {
    if (!element?.isConnected) return false;
    const rect = element.getBoundingClientRect();
    return clientX >= rect.left && clientX <= rect.right && clientY >= rect.top && clientY <= rect.bottom;
  }
  function markElementImeComposition(target, active) {
    if (!target) return;
    const $target = $(target);
    $target.data("imeComposing", Boolean(active));
    if (!active) $target.data("imeCompositionEndedAt", nowMs());
  }
  function isElementImeCompositionEvent(event, target) {
    const nativeEvent = event?.originalEvent || event || {};
    const keyCode = Number(nativeEvent.keyCode || nativeEvent.which || event?.keyCode || event?.which || 0);
    const $target = target ? $(target) : null;
    const recentlyEndedAt = Number($target?.data("imeCompositionEndedAt") || 0);
    const isEscape = String(event?.key || nativeEvent.key || "") === "Escape" || String(event?.code || nativeEvent.code || "") === "Escape" || keyCode === 27;
    const justEndedForActionKey = (isEnterKeyEvent(event) || isEscape) && recentlyEndedAt > 0 && nowMs() - recentlyEndedAt < 80;
    return Boolean(nativeEvent.isComposing || event?.isComposing || keyCode === 229 || $target?.data("imeComposing") || justEndedForActionKey);
  }
  function markImeComposition(target, scope, active) {
    if (!target) return;
    markElementImeComposition(target, active);
    if (scope === "formula") state.formulaComposing = Boolean(active);
    else state.cellEditComposing = Boolean(active);
  }
  function isImeCompositionEvent(event, target, scope) {
    const scopeComposing = scope === "formula" ? state.formulaComposing : scope === "cell" ? state.cellEditComposing : false;
    return Boolean(isElementImeCompositionEvent(event, target) || scopeComposing);
  }
  function nowMs() {
    return window.performance?.now?.() || Date.now();
  }
  function isEnterKeyEvent(event) {
    const nativeEvent = event?.originalEvent || event || {};
    const key = String(event?.key || nativeEvent.key || "");
    const code = String(event?.code || nativeEvent.code || "");
    const keyCode = Number(nativeEvent.keyCode || nativeEvent.which || event?.keyCode || event?.which || 0);
    return key === "Enter" || key === "NumpadEnter" || code === "Enter" || code === "NumpadEnter" || keyCode === 13;
  }
  function isAltModifierEvent(event) {
    const nativeEvent = event?.originalEvent || event || {};
    return Boolean(event?.altKey || nativeEvent.altKey || event?.getModifierState?.("Alt") || nativeEvent.getModifierState?.("Alt"));
  }

  $tabs.on("click", "button", function () {
    state.active = Number($(this).data("index"));
    state.selected = null;
	  refreshStandaloneWorkbookChartsFromSources();
	  renderTabs();
	  renderSheet();
    $sheet.scrollLeft(0);
    $hscroll.scrollLeft(0);
  });
  $add.on("click", function () {
    const name = uniqueSheetName(model.sheets.map((sheet) => sheet.name));
    model.sheets.push(createBlankSheet(name));
    try {
      hf.addSheet(name);
    } catch {
      const sheets = denseFormulaSheetsForWorkbook();
      hf = HF.buildFromSheets(sheets, hyperFormulaBuildConfig());
    }
    state.active = model.sheets.length - 1;
    state.selected = null;
    renderTabs();
    renderSheet();
    $status.text("新しいワークシートを追加しました：" + name);
  });
  $sheet.on("scroll", function () {
    closeStandaloneValidationDropdownPopup();
    syncHorizontalScrollbarFromSheet();
    scheduleCanvasPaint();
  });
  $hscroll.on("scroll", function () {
    syncSheetFromHorizontalScrollbar();
    scheduleCanvasPaint();
  });
  $splitter.on("pointerdown", startFooterResize);
  $zoomSlider.on("input", function () {
    setZoom(Number(this.value));
  });
  $zoomOut.on("click", function () {
    setZoom(Number(model.sheets[state.active].zoomScale || 100) - 10);
  });
  $zoomIn.on("click", function () {
    setZoom(Number(model.sheets[state.active].zoomScale || 100) + 10);
  });
  $zoomLabel.on("click", openZoomDialog);
  $zoomLabel.on("keydown", function (event) {
    if (event.key !== "Enter" && event.key !== " ") return;
    event.preventDefault();
    openZoomDialog(event);
  });
  $(document).on("compositionstart", "input, textarea, [contenteditable='true']", function (event) {
    markElementImeComposition(event.currentTarget, true);
  });
  $(document).on("compositionend", "input, textarea, [contenteditable='true']", function (event) {
    markElementImeComposition(event.currentTarget, false);
  });
  $(document).on("keydown", handleStandaloneKeyboardShortcut);
  document.addEventListener("wheel", handleZoomWheel, { passive: false, capture: true });
  $(document).on("pointerdown", closeZoomDialogOnOutsidePointer);
  $(document).on("click", "[data-wx-zoom-dialog-action]", handleZoomDialogAction);
  $(document).on("change", 'input[name="wxZoomDialogOption"]', handleZoomDialogOptionChange);
  $(document).on("focusin", "[data-wx-zoom-custom-input]", selectCustomZoomOption);
  $(document).on("keydown", "#wx-zoom-dialog", handleZoomDialogKeydown);
  $(document).on("keydown", closeZoomDialogOnEscape);
  $("#wx-view-buttons").on("click", "[data-wx-view]", function () {
    const view = this.dataset.wxView;
    model.sheets[state.active].viewMode = view;
    renderSheet();
    updateFooterUi();
    $status.text(model.sheets[state.active].name + "：" + (this.getAttribute("aria-label") || this.title || view));
  });
  window.addEventListener("pointermove", updateRuntimeEditDrag);
  window.addEventListener("pointerup", finishRuntimeEditDrag);
  window.addEventListener("blur", finishRuntimeEditDrag);
  $sheet.on("pointerdown", ".sheet-cell[contenteditable='true']", startRuntimeEditDrag);
  $sheet.on("click", ".cell-link", handleStandaloneCellHyperlinkClick);
  $sheet.on("pointerdown", ".cell-validation-dropdown-button", handleStandaloneValidationDropdownButtonPointerDown);
  $sheet.on("click", ".cell-validation-dropdown-button", handleStandaloneValidationDropdownButtonClick);
  $(document).on("click", "[data-standalone-validation-dropdown-choice]", handleStandaloneValidationDropdownChoice);
  $(document).on("keydown", "[data-standalone-validation-dropdown-choice]", handleStandaloneValidationDropdownKeydown);
  $(document).on("pointerdown", closeStandaloneValidationDropdownOnOutsidePointer);
  $(document).on("keydown", closeStandaloneValidationDropdownOnEscape);
  $sheet.on("click focusin", ".sheet-cell[data-row][data-col]", function (event) {
    if (event.type === "click") {
      const hyperlink = standaloneHyperlinkFromCellClickEvent(event);
      if (hyperlink && openStandaloneHyperlinkTarget(hyperlink)) {
        event.preventDefault();
        event.stopPropagation();
        return;
      }
    }
    if (state.suppressCellSelectionUntil && nowMs() < state.suppressCellSelectionUntil) return;
    state.suppressCellSelectionUntil = 0;
    select($(this));
  });
  $sheet.on("dblclick", ".sheet-cell[data-row][data-col]", function () {
    beginCellEdit($(this));
  });
  $sheet.on("keydown", ".sheet-cell[data-row][data-col]", function (event) {
    const dateTimeShortcut = keyboardDateTimeShortcutKind(event, String(event.key || "").toLowerCase());
    if (dateTimeShortcut) {
      event.preventDefault();
      beginCellEdit($(this), { initialText: keyboardDateTimeShortcutText(dateTimeShortcut, new Date()), replaceText: true });
      $status.text((dateTimeShortcut === "time" ? "時刻" : "日付") + "を入力しました。");
      return;
    }
    if (event.key === "F2") {
      event.preventDefault();
      beginCellEdit($(this));
    }
  });
  $sheet.on("input", ".sheet-cell[contenteditable='true']", function () {
    syncEditableCellAlignment(this);
  });
  $sheet.on("keydown", ".sheet-cell[contenteditable='true']", function (event) {
    const dateTimeShortcut = keyboardDateTimeShortcutKind(event, String(event.key || "").toLowerCase());
    if (dateTimeShortcut) {
      if (isImeCompositionEvent(event, this, "cell")) return;
      event.preventDefault();
      insertStandaloneDateTimeIntoEditable(this, dateTimeShortcut);
      return;
    }
    if (isEnterKeyEvent(event)) {
      if (isAltModifierEvent(event)) {
        event.preventDefault();
        insertLineBreakInEditable(this);
        return;
      }
      if (isImeCompositionEvent(event, this, "cell")) return;
      event.preventDefault();
      this.blur();
    } else if (event.key === "Escape") {
      event.preventDefault();
      $(this).text(String($(this).data("before") || ""));
      this.blur();
    }
  });
  $sheet.on("compositionstart", ".sheet-cell[contenteditable='true']", function () {
    markImeComposition(this, "cell", true);
  });
  $sheet.on("compositionend", ".sheet-cell[contenteditable='true']", function () {
    markImeComposition(this, "cell", false);
  });
  $sheet.on("blur", ".sheet-cell[contenteditable='true']", function () {
    const before = String($(this).data("before") || "");
    const after = editablePlainText(this);
    $(this).removeAttr("contenteditable");
    if (before !== after) commit(Number($(this).data("row")), Number($(this).data("col")), after);
    else renderSheet();
  });
  $apply.on("click", function () {
    if (state.selected) commit(state.selected.row, state.selected.col, $formula.val());
  });
  $formula.on("keydown", function (event) {
    if (isEnterKeyEvent(event) && isAltModifierEvent(event)) {
      event.preventDefault();
      insertLineBreakInTextControl(this);
      return;
    }
    if (event.key === "Enter" && state.selected) {
      if (isImeCompositionEvent(event, this, "formula")) return;
      event.preventDefault();
      commit(state.selected.row, state.selected.col, $formula.val());
    }
  });
  $formula.on("compositionstart", function () {
    markImeComposition(this, "formula", true);
  });
  $formula.on("compositionend", function () {
    markImeComposition(this, "formula", false);
  });
  $toggle.on("click", function () {
    state.showFormulas = !state.showFormulas;
    $toggle.text(state.showFormulas ? "結果を表示" : "数式を表示");
    renderSheet();
  });

  renderTabs();
  renderSheet();
  $status.text("HTML ワークブックを読み込みました");

  function updateFooterUi() {
    const sheet = model.sheets[state.active];
    const zoom = Math.round(Number(sheet.zoomScale || 100));
    const view = sheet.viewMode || "normal";
    $zoomSlider.val(String(zoom));
    $zoomLabel.text(zoom + "%");
    updateOpenZoomDialogValue(zoom);
    $("#wx-view-buttons [data-wx-view]").each(function () {
      this.classList.toggle("active", this.dataset.wxView === view);
    });
  }
  function setZoom(value) {
    const sheet = model.sheets[state.active];
    const previousZoom = Number(sheet.zoomScale || 100);
    const nextZoom = Math.max(10, Math.min(400, Math.round(Number(value) || 100)));
    if (nextZoom === previousZoom) {
      updateFooterUi();
      return;
    }
    const previousScale = Number(sheet.displayScale || 1);
    const nextScale = displayScaleForZoom(sheet, nextZoom);
    const ratio = nextScale / previousScale;
    scaleSheetDisplay(sheet, ratio);
    sheet.zoomScale = nextZoom;
    sheet.displayScale = nextScale;
    renderSheet();
    $status.text(sheet.name + "：ズーム " + nextZoom + "%");
  }
  function handleZoomWheel(event) {
    if (!event.ctrlKey || !isWorkbookWheelTarget(event)) return;
    event.preventDefault();
    event.stopPropagation();
    const deltaY = Number(event.deltaY || 0);
    if (!deltaY) return;
    const direction = deltaY < 0 ? 1 : -1;
    const steps = Math.max(1, Math.min(5, Math.round(Math.abs(deltaY) / 100) || 1));
    setZoom(Number(model.sheets[state.active].zoomScale || 100) + direction * steps * 10);
  }
  function openZoomDialog(event) {
    event?.preventDefault();
    event?.stopPropagation();
    closeZoomDialog();
    const currentZoom = Math.round(Number(model.sheets[state.active].zoomScale || 100));
    const presets = [
      { value: 200, label: "200%(0)" },
      { value: 100, label: "100%(1)" },
      { value: 75, label: "75%(7)" },
      { value: 50, label: "50%(5)" },
      { value: 25, label: "25%(2)" },
    ];
    const presetValues = presets.map(function (preset) {
      return preset.value;
    });
    const selectedValue = presetValues.includes(currentZoom) ? String(currentZoom) : "custom";
    const optionsHtml = presets
      .map(function (preset) {
        return "<label class='wx-zoom-dialog-option'" +
          standaloneZoomDialogAccessKeyAttr(preset.label) +
          "><input type='radio' name='wxZoomDialogOption' value='" +
          preset.value +
          "'" +
          (selectedValue === String(preset.value) ? " checked" : "") +
          " /><span>" +
          accessKeyLabelHtml(preset.label) +
          "</span></label>";
      })
      .join("");
    const dialog = document.createElement("div");
    dialog.id = "wx-zoom-dialog";
    dialog.setAttribute("role", "dialog");
    dialog.setAttribute("aria-label", "ズーム");
    dialog.innerHTML =
      "<div class='wx-zoom-dialog-title'>ズーム</div>" +
      "<div class='wx-zoom-dialog-section-title'>倍率</div>" +
      "<div class='wx-zoom-dialog-options'>" +
      optionsHtml +
      "<label class='wx-zoom-dialog-option'" +
      standaloneZoomDialogAccessKeyAttr("選択範囲をズーム(F)") +
      "><input type='radio' name='wxZoomDialogOption' value='selection' /><span>" +
      accessKeyLabelHtml("選択範囲をズーム(F)") +
      "</span></label>" +
      "<label class='wx-zoom-dialog-option wx-zoom-dialog-custom'" +
      standaloneZoomDialogAccessKeyAttr("指定(C):") +
      "><input type='radio' name='wxZoomDialogOption' value='custom'" +
      (selectedValue === "custom" ? " checked" : "") +
      " /><span>" +
      accessKeyLabelHtml("指定(C):") +
      "</span><input class='wx-zoom-dialog-input' type='number' min='10' max='400' step='1' value='" +
      currentZoom +
      "' data-wx-zoom-custom-input aria-label='ズーム倍率' /><span>%</span></label>" +
      "</div><div class='wx-zoom-dialog-actions'><button class='wx-zoom-dialog-button primary' type='button' data-wx-zoom-dialog-action='ok'>OK</button><button class='wx-zoom-dialog-button' type='button' data-wx-zoom-dialog-action='cancel'>キャンセル</button></div>";
    document.body.append(dialog);
    positionZoomDialog(dialog);
    const initialFocus = selectedValue === "custom" ? dialog.querySelector("[data-wx-zoom-custom-input]") : dialog.querySelector("input:checked");
    initialFocus?.focus({ preventScroll: true });
    if (selectedValue === "custom") initialFocus?.select?.();
  }
  function standaloneZoomDialogAccessKeyAttr(label) {
    const accessKey = menuAccessKeyFromLabel(label);
    return accessKey ? " data-access-key='" + escapeAttr(accessKey) + "'" : "";
  }
  function positionZoomDialog(dialog) {
    const anchor = $zoomLabel[0];
    if (!anchor || !dialog) return;
    dialog.style.visibility = "hidden";
    dialog.style.left = "0px";
    dialog.style.top = "0px";
    const anchorRect = anchor.getBoundingClientRect();
    const dialogRect = dialog.getBoundingClientRect();
    const margin = 8;
    const left = Math.max(margin, Math.min(anchorRect.right - dialogRect.width, window.innerWidth - dialogRect.width - margin));
    const topAbove = anchorRect.top - dialogRect.height - margin;
    const top = topAbove >= margin ? topAbove : Math.min(anchorRect.bottom + margin, window.innerHeight - dialogRect.height - margin);
    dialog.style.left = Math.round(left) + "px";
    dialog.style.top = Math.round(Math.max(margin, top)) + "px";
    dialog.style.visibility = "visible";
  }
  function handleZoomDialogAction(event) {
    const button = event.currentTarget;
    if (!button.closest("#wx-zoom-dialog")) return;
    event.preventDefault();
    event.stopPropagation();
    if (button.dataset.wxZoomDialogAction === "cancel") {
      closeZoomDialog();
      return;
    }
    applyZoomDialog();
  }
  function handleZoomDialogKeydown(event) {
    if (isImeCompositionEvent(event, event.target, "dialog")) return;
    const accessKey = accessKeyFromKeyboardEvent(event);
    if (accessKey && event.altKey && !event.ctrlKey && !event.metaKey) {
      const target = [...event.currentTarget.querySelectorAll("[data-access-key]")]
        .find(function (element) {
          return String(element.dataset.accessKey || "").toUpperCase() === accessKey;
        });
      if (!target) return;
      event.preventDefault();
      activateZoomDialogAccessKeyTarget(target);
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      applyZoomDialog();
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      closeZoomDialog();
    }
  }
  function activateZoomDialogAccessKeyTarget(target) {
    const radio = target.querySelector?.("input[type='radio']") || target.closest?.("label")?.querySelector?.("input[type='radio']");
    if (!radio) return;
    radio.click();
    if (radio.value === "custom") {
      const input = target.querySelector?.("[data-wx-zoom-custom-input]") ||
        target.closest?.("label")?.querySelector?.("[data-wx-zoom-custom-input]");
      input?.focus({ preventScroll: true });
      input?.select?.();
      return;
    }
    radio.focus({ preventScroll: true });
  }
  function handleZoomDialogOptionChange(event) {
    const dialog = event.target.closest?.("#wx-zoom-dialog");
    if (!dialog) return;
    if (event.target.value === "custom") dialog.querySelector("[data-wx-zoom-custom-input]")?.focus({ preventScroll: true });
  }
  function selectCustomZoomOption(event) {
    const dialog = event.target.closest?.("#wx-zoom-dialog");
    if (!dialog) return;
    const customRadio = dialog.querySelector('input[name="wxZoomDialogOption"][value="custom"]');
    if (customRadio) customRadio.checked = true;
  }
  function applyZoomDialog() {
    const dialog = document.querySelector("#wx-zoom-dialog");
    if (!dialog) return;
    const selectedOption = dialog.querySelector('input[name="wxZoomDialogOption"]:checked')?.value || "custom";
    if (selectedOption === "selection") {
      if (zoomSelectionToFit()) closeZoomDialog();
      return;
    }
    const rawValue = selectedOption === "custom" ? dialog.querySelector("[data-wx-zoom-custom-input]")?.value : selectedOption;
    setZoom(Math.max(10, Math.min(400, Math.round(Number(rawValue) || 100))));
    closeZoomDialog();
  }
  function zoomSelectionToFit() {
    if (!state.selected) {
      $status.text("ズームするセル範囲を選択してください。");
      return false;
    }
    const sheet = model.sheets[state.active];
    const sheetEl = $sheet[0];
    const ranges = Array.isArray(state.selectedRanges) && state.selectedRanges.length
      ? state.selectedRanges
      : [state.selectedRange || { top: state.selected.row, left: state.selected.col, bottom: state.selected.row, right: state.selected.col }];
    const range = standaloneZoomSelectionFitRange(sheet, ranges);
    if (!range) {
      $status.text("ズームするセル範囲を選択してください。");
      return false;
    }
    const left = coordinateToPixels(sheet, "col", range.left - 1);
    const top = coordinateToPixels(sheet, "row", range.top - 1);
    const right = coordinateToPixels(sheet, "col", range.right);
    const bottom = coordinateToPixels(sheet, "row", range.bottom);
    const rangeWidth = Math.max(1, right - left);
    const rangeHeight = Math.max(1, bottom - top);
    const availableWidth = Math.max(1, sheetEl.clientWidth - 46 - 16);
    const availableHeight = Math.max(1, sheetEl.clientHeight - 22 - 16);
    const currentZoom = Number(sheet.zoomScale || 100);
    const nextZoom = Math.max(10, Math.min(400, Math.round(currentZoom * Math.min(availableWidth / rangeWidth, availableHeight / rangeHeight))));
    setZoom(nextZoom);
    $status.text(standaloneRangesLabel(ranges.length ? ranges : [range]) + "：ズーム " + nextZoom + "%");
    return true;
  }
  function standaloneZoomSelectionFitRange(sheet, ranges) {
    const normalized = (ranges || [])
      .map(function (range) {
        return standaloneZoomSelectionNormalizeRange(sheet, range);
      })
      .filter(Boolean);
    if (!normalized.length) return null;
    return standaloneZoomSelectionUnionRanges(normalized);
  }
  function standaloneZoomSelectionNormalizeRange(sheet, range) {
    if (!sheet || !range) return null;
    const used = standaloneSheetUsedRange(sheet);
    const fallback = used || {
      top: Math.min(Math.max(1, Number(range.top) || 1), sheet.rowCount || 1),
      left: Math.min(Math.max(1, Number(range.left) || 1), sheet.colCount || 1),
      bottom: Math.min(Math.max(1, Number(range.bottom || range.top) || 1), sheet.rowCount || 1),
      right: Math.min(Math.max(1, Number(range.right || range.left) || 1), sheet.colCount || 1),
    };
    const normalized = { ...range };
    if (range.fullRow && range.fullColumn) {
      normalized.top = fallback.top;
      normalized.bottom = fallback.bottom;
      normalized.left = fallback.left;
      normalized.right = fallback.right;
    } else if (range.fullRow) {
      normalized.left = fallback.left;
      normalized.right = fallback.right;
    } else if (range.fullColumn) {
      normalized.top = fallback.top;
      normalized.bottom = fallback.bottom;
    }
    normalized.top = Math.min(Math.max(1, Number(normalized.top) || 1), sheet.rowCount || 1);
    normalized.left = Math.min(Math.max(1, Number(normalized.left) || 1), sheet.colCount || 1);
    normalized.bottom = Math.min(Math.max(normalized.top, Number(normalized.bottom) || normalized.top), sheet.rowCount || normalized.top);
    normalized.right = Math.min(Math.max(normalized.left, Number(normalized.right) || normalized.left), sheet.colCount || normalized.left);
    return normalized;
  }
  function standaloneSheetUsedRange(sheet) {
    if (!sheet) return null;
    const ranges = [];
    Object.keys(sheet.cells || {}).forEach(function (key) {
      const cell = sheet.cells[key];
      if (!cell) return;
      const hasContent = cell.raw != null || cell.formula || cell.cached != null || cell.richText?.length || cell.hyperlink || cell.note || cell.commentThread?.length;
      const hasStyle = cell.css && Object.keys(cell.css).length;
      const hasBorders = cell.borders && Object.keys(cell.borders).length;
      if (!hasContent && !hasStyle && !hasBorders) return;
      const parts = key.split(":");
      const row = Number(parts[0]);
      const col = Number(parts[1]);
      if (Number.isFinite(row) && Number.isFinite(col)) ranges.push({ top: row, left: col, bottom: row, right: col });
    });
    (sheet.merges || []).forEach(function (merge) {
      if (merge) ranges.push(merge);
    });
    (sheet.images || []).forEach(function (image) {
      if (image?.range) ranges.push(image.range);
      else if (image?.from) ranges.push({ top: image.from.row || 1, left: image.from.col || 1, bottom: image.to?.row || image.from.row || 1, right: image.to?.col || image.from.col || 1 });
    });
    if (!ranges.length) return null;
    return standaloneZoomSelectionUnionRanges(ranges);
  }
  function standaloneZoomSelectionUnionRanges(ranges) {
    const items = (ranges || []).filter(Boolean);
    if (!items.length) return null;
    return {
      top: Math.min.apply(null, items.map(function (range) { return Number(range.top) || 1; })),
      left: Math.min.apply(null, items.map(function (range) { return Number(range.left) || 1; })),
      bottom: Math.max.apply(null, items.map(function (range) { return Number(range.bottom || range.top) || 1; })),
      right: Math.max.apply(null, items.map(function (range) { return Number(range.right || range.left) || 1; })),
    };
  }
  function standaloneRangesLabel(ranges) {
    return (ranges || []).map(standaloneRangeLabel).filter(Boolean).join(",") || "";
  }
  function standaloneRangeLabel(range) {
    if (!range) return "";
    if (range.fullColumn) return columnName(range.left) + ":" + columnName(range.right);
    if (range.fullRow) return range.top + ":" + range.bottom;
    const start = columnName(range.left) + range.top;
    const end = columnName(range.right) + range.bottom;
    return start === end ? start : start + ":" + end;
  }
  function closeZoomDialogOnOutsidePointer(event) {
    if (!document.querySelector("#wx-zoom-dialog")) return;
    if (event.target.closest?.("#wx-zoom-dialog") || event.target.closest?.("#wx-zoom-label")) return;
    closeZoomDialog();
  }
  function closeZoomDialogOnEscape(event) {
    if (event.key !== "Escape" || !document.querySelector("#wx-zoom-dialog")) return;
    if (isImeCompositionEvent(event, event.target, "dialog")) return;
    event.preventDefault();
    closeZoomDialog();
  }
  function closeZoomDialog() {
    document.querySelector("#wx-zoom-dialog")?.remove();
  }
  function updateOpenZoomDialogValue(zoom) {
    const input = document.querySelector("#wx-zoom-dialog [data-wx-zoom-custom-input]");
    if (input && document.activeElement !== input) input.value = String(zoom);
  }
  function isWorkbookWheelTarget(event) {
    const root = document.querySelector(".websheet");
    if (!root) return false;
    const path = event.composedPath?.() || [];
    return path.includes(root) || root.contains(event.target);
  }
  function sheetBaseZoomScale(sheet) {
    const base = Number(sheet?.baseZoomScale || sheet?.zoomScale || 100);
    return Number.isFinite(base) && base > 0 ? base : 100;
  }
  function displayScaleForZoom(sheet, zoom) {
    const targetZoom = Math.max(10, Math.min(400, Math.round(Number(zoom) || 100)));
    return targetZoom / sheetBaseZoomScale(sheet);
  }
  function scaleSheetDisplay(sheet, ratio) {
    if (!Number.isFinite(ratio) || ratio <= 0) return;
    sheet.defaultColumnWidth = Math.max(1, (sheet.defaultColumnWidth || 64) * ratio);
    sheet.defaultRowHeight = Math.max(1, (sheet.defaultRowHeight || 20) * ratio);
    sheet.columns.forEach(function (column) {
      column.width = Math.max(0, (column.width || 0) * ratio);
    });
    sheet.rows.forEach(function (row) {
      row.height = Math.max(0, (row.height || 0) * ratio);
    });
    (sheet.images || []).forEach(function (image) {
      if (!image.ext) return;
      image.ext.width = Math.max(Number.EPSILON, (Number(image.ext.width) || 1) * ratio);
      image.ext.height = Math.max(Number.EPSILON, (Number(image.ext.height) || 1) * ratio);
    });
  }
  function uniqueSheetName(existingNames) {
    const existing = new Set(existingNames);
    let index = existingNames.length + 1;
    let name = "Sheet" + index;
    while (existing.has(name)) {
      index += 1;
      name = "Sheet" + index;
    }
    return name;
  }
  function createBlankSheet(name) {
    const rows = 100;
    const cols = 26;
    return {
      name,
      state: "visible",
      rowCount: rows,
      colCount: cols,
      showGridLines: true,
      showHeadings: true,
      showObjects: true,
      showOutlineSymbols: true,
      zoomScale: 100,
      baseZoomScale: 100,
      displayScale: 1,
      defaultCellCss: {},
      columns: Array.from({ length: cols }, () => ({ width: 64, hidden: false })),
      rows: Array.from({ length: rows }, () => ({ height: 20, hidden: false })),
      merges: [],
      images: [],
      cells: {},
      hfData: [],
      warnings: [],
    };
  }
  function updateHorizontalScrollbar() {
    requestAnimationFrame(function () {
      const sheetEl = $sheet[0];
      const scrollEl = $hscroll[0];
      const spacerEl = $hspacer[0];
      if (!sheetEl || !scrollEl || !spacerEl) return;
      const maxScrollLeft = Math.max(0, sheetEl.scrollWidth - sheetEl.clientWidth);
      spacerEl.style.width = scrollEl.clientWidth + maxScrollLeft + "px";
      forceSyncHorizontalScrollbarFromSheet();
    });
  }
  function forceSyncHorizontalScrollbarFromSheet() {
    const sheetEl = $sheet[0];
    const scrollEl = $hscroll[0];
    if (!sheetEl || !scrollEl) return;
    if (scrollEl.scrollLeft === sheetEl.scrollLeft) {
      syncingScroll = false;
      return;
    }
    syncingScroll = true;
    scrollEl.scrollLeft = sheetEl.scrollLeft;
    requestAnimationFrame(function () {
      syncingScroll = false;
    });
  }
  function syncHorizontalScrollbarFromSheet() {
    if (syncingScroll) return;
    syncingScroll = true;
    $hscroll[0].scrollLeft = $sheet[0].scrollLeft;
    requestAnimationFrame(function () {
      syncingScroll = false;
    });
  }
  function syncSheetFromHorizontalScrollbar() {
    if (syncingScroll) return;
    syncingScroll = true;
    $sheet[0].scrollLeft = $hscroll[0].scrollLeft;
    requestAnimationFrame(function () {
      syncingScroll = false;
    });
  }
  function startFooterResize(event) {
    event.preventDefault();
    const footer = $footer[0];
    const startX = event.clientX;
    const tabsPane = footer.querySelector(".wx-tabs-pane");
    const startWidth = tabsPane?.getBoundingClientRect().width || 280;
    const footerWidth = footer.getBoundingClientRect().width;
    const minWidth = 140;
    const maxWidth = Math.max(minWidth, footerWidth - 180);
    const move = function (moveEvent) {
      const nextWidth = Math.max(minWidth, Math.min(maxWidth, startWidth + moveEvent.clientX - startX));
      footer.style.setProperty("--wx-tabs-width", Math.round(nextWidth) + "px");
      updateHorizontalScrollbar();
    };
    const up = function () {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      updateHorizontalScrollbar();
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  }
  function buildMergeMaps(merges) {
    const mergeList = Array.isArray(merges) ? merges : [];
    const cached = standaloneMergeMapsCacheByMerges.get(mergeList);
    if (cached?.length === mergeList.length) return cached.maps;
    const master = new Map();
    const covered = new Set();
    const coveredToMaster = new Map();
    mergeList.forEach((merge) => {
      const key = merge.top + ":" + merge.left;
      master.set(key, merge);
      for (let row = merge.top; row <= merge.bottom; row += 1) {
        for (let col = merge.left; col <= merge.right; col += 1) {
          const cellKey = row + ":" + col;
          if (cellKey !== key) {
            covered.add(cellKey);
            coveredToMaster.set(cellKey, key);
          }
        }
      }
    });
    const maps = { master, covered, coveredToMaster };
    standaloneMergeMapsCacheByMerges.set(mergeList, { length: mergeList.length, maps });
    return maps;
  }
  function borderLinesForSheet(sheet, maps, visibleRange) {
    const xs = columnBoundaryPositions(sheet);
    const ys = rowBoundaryPositions(sheet);
    const lines = new Map();
    Object.values(sheet.cells).forEach((cell) => {
      const borders = cell.borders || {};
      if (!Object.keys(borders).length) return;
      const key = cell.row + ":" + cell.col;
      if (maps.covered.has(key)) return;
      const merge = maps.master.get(key);
      if (visibleRange && !rangeIntersects(merge || { top: cell.row, left: cell.col, bottom: cell.row, right: cell.col }, visibleRange)) return;
      const leftCol = cell.col;
      const rightCol = merge ? merge.right : cell.col;
      const topRow = cell.row;
      const bottomRow = merge ? merge.bottom : cell.row;
      const x1 = Math.min(xs[leftCol], xs[rightCol + 1]);
      const x2 = Math.max(xs[leftCol], xs[rightCol + 1]);
      const y1 = ys[topRow];
      const y2 = ys[bottomRow + 1];
      addBorderLine(lines, "h", y1, x1, x2, borders.top);
      addBorderLine(lines, "h", y2, x1, x2, borders.bottom);
      addBorderLine(lines, "v", x1, y1, y2, borders.left);
      addBorderLine(lines, "v", x2, y1, y2, borders.right);
    });
    return Array.from(lines.values());
  }
  function columnBoundaryPositions(sheet) {
    const positions = [];
    for (let col = 1; col <= sheet.colCount + 1; col += 1) {
      positions[col] = coordinateToPixels(sheet, "col", col - 1);
    }
    return positions;
  }
  function rowBoundaryPositions(sheet) {
    const positions = [];
    let y = 22;
    for (let row = 1; row <= sheet.rowCount; row += 1) {
      positions[row] = y;
      y += sheet.rows[row - 1]?.hidden ? 0 : sheet.rows[row - 1]?.height || 0;
    }
    positions[sheet.rowCount + 1] = y;
    return positions;
  }
  function addBorderLine(lines, orientation, position, start, end, border) {
    if (!border || end <= start) return;
    const key = orientation + ":" + position + ":" + start + ":" + end;
    const candidate = { orientation, position, start, end, border };
    const existing = lines.get(key);
    if (!existing || compareBorder(candidate.border, existing.border) > 0) lines.set(key, candidate);
  }
  function compareBorder(a, b) {
    return (a.weight || 0) - (b.weight || 0) || (a.width || 0) - (b.width || 0);
  }
  function borderLineHtml(line) {
    const width = line.border.width || 1;
    const style = line.border.style || "solid";
    const color = line.border.color || "#000000";
    const css =
      line.orientation === "h"
        ? "left:" + line.start + "px;top:" + (line.position - width / 2) + "px;width:" + (line.end - line.start) + "px;border-top:" + width + "px " + style + " " + color
        : "left:" + (line.position - width / 2) + "px;top:" + line.start + "px;height:" + (line.end - line.start) + "px;border-left:" + width + "px " + style + " " + color;
    return "<div class='sheet-border-line' style='" + escapeAttr(css) + "'></div>";
  }
  function textOverflowWidth(sheet, maps, row, col, colSpan, cell, value) {
    if (!value || cellShowsValidationDropdown(cell)) return null;
    const css = cellEffectiveCss(sheet, cell);
    if (css.whiteSpace && css.whiteSpace !== "nowrap") return null;
    if (css.textAlign === "center" || css.textAlign === "right") return null;
    const startRight = col + colSpan - 1;
    let endCol = startRight;
    for (let nextCol = startRight + 1; nextCol <= sheet.colCount; nextCol += 1) {
      if (isBlockingCell(sheet, maps, row, nextCol)) break;
      endCol = nextCol;
    }
    return endCol === startRight ? null : widthBetweenColumns(sheet, col, endCol);
  }
  function isBlockingCell(sheet, maps, row, col) {
    const key = row + ":" + col;
    const masterKey = maps.coveredToMaster.get(key) || key;
    return isOccupiedCell(sheet.cells[masterKey]);
  }
  function isOccupiedCell(cell) {
    if (!cell) return false;
    if (cell.formula || cell.hyperlink) return true;
    if (cell.raw == null) return false;
    return String(cell.raw) !== "";
  }
  function widthBetweenColumns(sheet, startCol, endCol) {
    let width = 0;
    for (let col = startCol; col <= endCol; col += 1) {
      width += sheet.columns[col - 1]?.hidden ? 0 : sheet.columns[col - 1]?.width || 0;
    }
    return width;
  }
  function coerce(input, options = {}) {
    const source = String(input || "");
    const preserveWhitespace = options.preserveWhitespace === true;
    const text = preserveWhitespace ? source : source.trim();
    if (!text) return null;
    if (text.startsWith("'")) return text.slice(1);
    if (text.startsWith("=")) return text;
    if (preserveWhitespace && text !== text.trim()) return text;
    if (/^(true|false)$/i.test(text)) return /^true$/i.test(text);
    if (/^-?\\d+(\\.\\d+)?$/.test(text)) return Number(text);
    return text;
  }
  function columnName(index) {
    let name = "";
    let value = index;
    while (value > 0) {
      const modulo = (value - 1) % 26;
      name = String.fromCharCode(65 + modulo) + name;
      value = Math.floor((value - modulo) / 26);
    }
    return name;
  }
  function styleObjectToString(style) {
    return Object.entries(style || {})
      .map(([key, value]) => key.replace(/[A-Z]/g, (match) => "-" + match.toLowerCase()) + ":" + value)
      .join(";");
  }
  function rowHeaderWidthForSheet(sheet) {
    const host = document.getElementById("wx-sheet");
    const rowHeight = Math.max(1, Number(sheet?.defaultRowHeight || 20));
    const viewportHeight = Math.max(0, host?.clientHeight || rowHeight * 40);
    const rowsPerPage = Math.max(1, Math.ceil((viewportHeight - 22) / rowHeight));
    const referenceRow = Math.max(1, Math.floor(Number(sheet?.rowCount || 1) + rowsPerPage));
    const digits = String(referenceRow).length;
    return Math.max(28, Math.min(72, Math.ceil(14 + digits * 7)));
  }
  function gridTrackStyle(sheet) {
    const sheetColumns = sheetColumnsRightToLeft(sheet) ? sheet.columns.slice().reverse() : sheet.columns;
    const columns = [rowHeaderWidthForSheet(sheet) + "px"].concat(sheetColumns.map((column) => (column.hidden ? 0 : column.width) + "px"));
    const rows = ["22px"].concat(sheet.rows.map((row) => (row.hidden ? 0 : row.height) + "px"));
    return "grid-template-columns:" + columns.join(" ") + ";grid-template-rows:" + rows.join(" ");
  }
  function sheetGridClassName(sheet) {
    const viewMode = ["page-layout", "page-break"].includes(sheet?.viewMode) ? sheet.viewMode : "normal";
    return "sheet-grid view-" + viewMode + (sheet.showGridLines === false ? " hide-gridlines" : "") + (sheet.showWhiteSpace === false ? " hide-white-space" : "") + (sheet.rightToLeft === true ? " is-right-to-left" : "");
  }
  function displayValue(value, numFmt) {
    if (value == null) return "";
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      if (isElapsedTimeFormat(numFmt)) {
        var elapsedSerial = Date.UTC(
          value.getFullYear(),
          value.getMonth(),
          value.getDate(),
          value.getHours(),
          value.getMinutes(),
          value.getSeconds()
        ) / 86400000 + 25569;
        return formatElapsedTimeForFormat(elapsedSerial, numFmt || "");
      }
      return formatDateForFormat(value, numFmt || "");
    }
    if (typeof value === "object") return value.value || value.error || String(value);
    if (typeof value === "number") {
      var primaryFormat = primaryExcelFormatSection(numFmt || "");
      if (isElapsedTimeFormat(primaryFormat)) return formatElapsedTimeForFormat(value, primaryFormat);
      if (isDateFormat(primaryFormat)) return formatDateForFormat(excelSerialToDate(value), primaryFormat);
      const text = formatNumberForFormat(value, primaryFormat);
      if (/%/.test(primaryFormat)) return formatNumberForFormat(value * 100, primaryFormat) + "%";
      const currency = currencySymbol(primaryFormat);
      return currency ? currency + text : text;
    }
    if (typeof value === "string") {
      var numericValue = numericStringValueForNumberFormat(value, numFmt || "");
      if (numericValue != null) return displayValue(numericValue, numFmt || "");
    }
    return String(value);
  }
  function numericStringValueForNumberFormat(value, numFmt) {
    var primaryFormat = primaryExcelFormatSection(numFmt || "").trim();
    if (!primaryFormat || primaryFormat === "@") return null;
    if (!numberFormatCanApplyToNumericString(primaryFormat)) return null;
    var text = String(value == null ? "" : value).trim().replace(/,/g, "");
    if (!/^[+-]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:e[+-]?\\d+)?$/i.test(text)) return null;
    var number = Number(text);
    return Number.isFinite(number) ? number : null;
  }
  function numberFormatCanApplyToNumericString(format) {
    return isDateFormat(format) || /%/.test(format) || Boolean(currencySymbol(format)) || /[0#?]/.test(format);
  }
  function isElapsedTimeFormat(numFmt) {
    return /\[(?:h+|m+|s+)\]/i.test(primaryExcelFormatSection(numFmt || ""));
  }
  function formatElapsedTimeForFormat(value, numFmt) {
    var number = Number(value);
    if (!Number.isFinite(number)) return String(value);
    var format = primaryExcelFormatSection(numFmt || "");
    if (!isElapsedTimeFormat(format)) return String(value);
    var negative = number < 0;
    var absoluteSeconds = Math.abs(number) * 86400;
    var plainFormat = format.replace(/"[^"]*"/g, "").replace(/\\./g, " ");
    var hasSeconds = /\[(?:s+)\]|s{1,2}/i.test(plainFormat);
    var hasMinutes = /\[(?:m+)\]|m{1,2}/i.test(plainFormat);
    var totalSeconds = hasSeconds
      ? Math.round(absoluteSeconds)
      : hasMinutes
        ? Math.round(absoluteSeconds / 60) * 60
        : Math.round(absoluteSeconds / 3600) * 3600;
    var hours = Math.floor(totalSeconds / 3600);
    var minutes = Math.floor(totalSeconds / 60);
    var seconds = totalSeconds;
    var clockHour = hours % 24;
    var clockMinute = minutes % 60;
    var clockSecond = seconds % 60;
    var pad2 = function (item) { return String(item).padStart(2, "0"); };
    var rendered = format.replace(/"([^"]*)"|\\(.)|\[(h+|m+|s+)\]|hh|ss|mm|h|s|m/gi, function (match, quoted, escaped, elapsed) {
      if (quoted != null) return quoted;
      if (escaped != null) return escaped;
      if (elapsed) {
        var unit = elapsed[0].toLowerCase();
        return String(unit === "h" ? hours : unit === "m" ? minutes : seconds);
      }
      var token = match.toLowerCase();
      if (token === "hh") return pad2(clockHour);
      if (token === "h") return String(clockHour);
      if (token === "mm") return pad2(clockMinute);
      if (token === "m") return String(clockMinute);
      if (token === "ss") return pad2(clockSecond);
      return String(clockSecond);
    });
    return (negative ? "-" : "") + rendered;
  }
  function formatDateForFormat(date, numFmt) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
    var format = primaryExcelFormatSection(numFmt || "").replace(/\[[^\]]+\]/g, "");
    if (!isDateFormat(format)) return date.toLocaleDateString();
    var monthNamesShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    var monthNamesLong = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    var weekdayNamesShort = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    var weekdayNamesLong = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    var pad2 = function (number) { return String(number).padStart(2, "0"); };
    var uses12HourClock = /AM\/PM|A\/P/i.test(format);
    var hour24 = date.getHours();
    var hour12 = hour24 % 12 || 12;
    var replacements = {
      yyyy: String(date.getFullYear()),
      yy: pad2(date.getFullYear() % 100),
      mmmm: monthNamesLong[date.getMonth()],
      mmm: monthNamesShort[date.getMonth()],
      mm: pad2(date.getMonth() + 1),
      m: String(date.getMonth() + 1),
      dddd: weekdayNamesLong[date.getDay()],
      ddd: weekdayNamesShort[date.getDay()],
      dd: pad2(date.getDate()),
      d: String(date.getDate()),
      hh: pad2(uses12HourClock ? hour12 : hour24),
      h: String(uses12HourClock ? hour12 : hour24),
      ss: pad2(date.getSeconds()),
      s: String(date.getSeconds()),
      "AM/PM": hour24 < 12 ? "AM" : "PM",
      "am/pm": hour24 < 12 ? "am" : "pm",
      "A/P": hour24 < 12 ? "A" : "P",
      "a/p": hour24 < 12 ? "a" : "p"
    };
    return format.replace(/"([^"]*)"|\\(.)|AM\/PM|A\/P|yyyy|mmmm|dddd|mmm|ddd|yy|hh|ss|mm|dd|h|s|m|d/gi, function (match, quoted, escaped, offset) {
      if (quoted != null) return quoted;
      if (escaped != null) return escaped;
      var key = match.toLowerCase();
      if (key === "m" || key === "mm") {
        return key === "mm" && dateFormatTokenIsMinute(format, offset, match.length)
          ? pad2(date.getMinutes())
          : key === "m" && dateFormatTokenIsMinute(format, offset, match.length)
            ? String(date.getMinutes())
            : replacements[key];
      }
      return replacements[match] || replacements[key] || match;
    }).replace(/"/g, "");
  }
  function dateFormatTokenIsMinute(format, start, length) {
    var source = String(format || "");
    var before = source.slice(Math.max(0, start - 8), start);
    var after = source.slice(start + length, Math.min(source.length, start + length + 8));
    var previousMatch = source.slice(0, start).match(/\S$/);
    var nextMatch = source.slice(start + length).match(/\S/);
    var previous = previousMatch ? previousMatch[0] : "";
    var next = nextMatch ? nextMatch[0] : "";
    if (previous === ":" || next === ":") return true;
    return /h{1,2}[^ymdhs]*$/i.test(before) || /^[^ymdhms]*s{1,2}/i.test(after);
  }
  function primaryExcelFormatSection(numFmt) {
    var text = String(numFmt || "");
    var quoted = false;
    var escaped = false;
    for (var index = 0; index < text.length; index += 1) {
      var char = text[index];
      if (escaped) {
        escaped = false;
        continue;
      }
      if (char === "\\\\") {
        escaped = true;
        continue;
      }
      if (char === '"') {
        quoted = !quoted;
        continue;
      }
      if (char === ";" && !quoted) return text.slice(0, index);
    }
    return text;
  }
  function formatNumberForFormat(value, numFmt) {
    if (!Number.isFinite(value)) return String(value);
    const format = primaryExcelFormatSection(numFmt || "");
    const decimals = formatDecimalPlaces(format);
    return value.toLocaleString(undefined, {
      useGrouping: /#,##/.test(String(format || "")),
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals || 10,
    });
  }
  function formatDecimalPlaces(numFmt) {
    const match = String(numFmt || "").match(/\\.([0#]+)/);
    return match ? Math.min(6, match[1].length) : 0;
  }
  function isDateFormat(numFmt) {
    if (isElapsedTimeFormat(numFmt)) return true;
    const format = dateDetectionFormat(numFmt || "");
    if (/[年月日]/.test(format)) return true;
    if (/^\s*(?:y{1,4}|m{1,4}|d{1,4})\s*$/i.test(format)) return true;
    var hasYear = /(^|[^a-z])y{2,4}/i.test(format);
    var hasMonth = /(^|[^a-z])m{1,4}/i.test(format);
    var hasDay = /(^|[^a-z])d{1,4}/i.test(format);
    var hasHour = /(^|[^a-z])h{1,2}/i.test(format);
    var hasSecond = /(^|[^a-z])s{1,2}/i.test(format);
    var hasAmPm = /AM\/PM|A\/P/i.test(format);
    return (hasYear && hasMonth) || (hasMonth && hasDay) || (hasYear && hasDay) || hasHour || hasSecond || hasAmPm;
  }
  function dateDetectionFormat(numFmt) {
    return primaryExcelFormatSection(numFmt || "")
      .replace(/\[[^\]]+\]/g, "")
      .replace(/"[^"]*"/g, " ")
      .replace(/\\./g, " ");
  }
  function excelSerialToDate(serial) {
    return new Date(Math.round((serial - 25569) * 86400000));
  }
  function currencySymbol(numFmt) {
    const match = String(numFmt || "").match(/[$¥€£]/);
    return match ? match[0] : "";
  }
  function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char]);
  }
  function menuAccessKeyFromLabel(label) {
    var text = String(label || "");
    var match = text.match(/\(([A-Za-z0-9])\)|（([A-Za-z0-9])）/);
    if (match) return String(match[1] || match[2]).toUpperCase();
    var leading = text.match(/^\s*([0-9])\s+/);
    return leading ? leading[1] : "";
  }
  function accessKeyFromKeyboardEvent(event) {
    var code = String(event?.code || "");
    var codeMatch = code.match(/^(?:Key([A-Z])|Digit([0-9])|Numpad([0-9]))$/i);
    if (codeMatch) return (codeMatch[1] || codeMatch[2] || codeMatch[3] || "").toUpperCase();
    var key = String(event?.key || "");
    return /^[A-Za-z0-9]$/.test(key) ? key.toUpperCase() : "";
  }
  function accessKeyLabelHtml(value) {
    var escaped = escapeHtml(value);
    var withParentheses = escaped.replace(/\(([A-Za-z0-9])\)|（([A-Za-z0-9])）/g, function (_match, asciiKey, fullWidthKey) {
      if (asciiKey) return "<span class='menu-access-key-token'>(<span class='menu-access-key'>" + asciiKey + "</span>)</span>";
      return "<span class='menu-access-key-token'>（<span class='menu-access-key'>" + fullWidthKey + "</span>）</span>";
    });
    if (withParentheses !== escaped) return withParentheses;
    return escaped.replace(/^(\s*)([0-9])(\s+)/, function (_match, prefix, accessKey, suffix) {
      return prefix + "<span class='menu-access-key-token'><span class='menu-access-key'>" + accessKey + "</span></span>" + suffix;
    });
  }
  function escapeAttr(value) {
    return escapeHtml(value).replace(/`/g, "&#96;");
  }
}

export {
  standaloneRuntime,
};
