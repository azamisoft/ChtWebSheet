// Feature module; shared bindings are imported explicitly.
import {
  $,
  $horizontalScrollSpacer,
  $horizontalScrollbar,
  $sheetFooter,
  $sheetHost,
  $sheetWorkspace,
  $workbookName,
  COLUMN_HEADER_HEIGHT,
  CellError,
  DEFAULT_SHAPE_TYPE,
  ExcelJS,
  FALLBACK_VISIBLE_COLS,
  FALLBACK_VISIBLE_ROWS,
  SAVE_AS_FILE_TYPES,
  SAVE_AS_LOCATIONS,
  STANDALONE_IMAGE_ASSET_ENCODING_GZIP_BASE64,
  STANDALONE_IMAGE_ASSET_GZIP_MIN_BASE64_LENGTH,
  STANDALONE_IMAGE_ASSET_GZIP_MIN_SAVINGS_RATIO,
  STANDALONE_IMAGE_DEVICE_PIXEL_RATIO,
  STANDALONE_IMAGE_MAX_DIMENSION,
  STANDALONE_IMAGE_WEBP_QUALITY,
  VIRTUAL_COL_OVERSCAN,
  VIRTUAL_ROW_OVERSCAN,
  WEBSHEET_FORMULA_NA_SENTINEL,
  loadExcelJsModule,
  sheetBorderCellCacheBySheet,
  sheetMergeMapsCacheByMerges,
  state,
} from "../../app-core.js";
import {
  autoFilterDateSerial,
} from "../data/filter-and-text-tools.js";
import {
  applySummaryFunction,
  decodeWorkbookRangeOrName,
} from "../data/validation-and-consolidation.js";
import {
  closeDataDialog,
} from "../drawing/charts-and-insertion.js";
import {
  fileExtensionFromImageMimeType,
  imageDimensionsFromSource,
  normalizedHeaderFooterSettings,
  shapeDescriptorForType,
} from "../drawing/shape-rendering.js";
import {
  refreshWorkbookChartsFromSources,
} from "../editing/autofill.js";
import {
  pad2,
} from "../editing/keyboard-and-paste.js";
import {
  decodeFormulaRangeForSheet,
  flattenFormulaMatrix,
  formulaArgumentBoolean,
  formulaArgumentMatrix,
  formulaArgumentNumberValue,
  formulaArgumentScalarValue,
  formulaLambdaApply,
  formulaLambdaNumeric,
  formulaLambdaTopLevelOperator,
  formulaLambdaTruthy,
  formulaMatrixDimensions,
  formulaMatrixValue,
  formulaRangeValueMatrix,
  formulaScalarValueForCell,
  getDisplayForCell,
  invalidateSheetUsedRangeCache,
  normalizeFormulaMatrix,
  rangeToLabel,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  coordinateToPixels,
  sheetColumnsRightToLeft,
} from "../editing/selection-and-cell-text.js";
import {
  cellEffectiveCss,
  cloneJsonLikeValue,
  normalizeSheetStyleRanges,
  sheetStyleRangeCell,
  stableObjectKey,
} from "../editing/structure-and-history.js";
import {
  textPartFormulaValue,
  webSheetFormulaCompareForSort,
  webSheetFormulaCompareValues,
  webSheetFormulaIsEmptyValue,
  webSheetFormulaRawValue,
  webSheetFormulaRowKey,
} from "../formula/calculations.js";
import {
  cachedCoordinateToPixels,
  cellShowsValidationDropdown,
  crispCanvasLine,
  forEachCellInRange,
  rangeIntersects,
  roundCssPx,
} from "../rendering/grid-and-canvas.js";
import {
  formatDecimalPlaces,
  formatUsesSeparator,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  columnName,
  columnWidthToPixels,
  currencySymbol,
  dateToExcelSerial,
  decodeRange,
  escapeAttr,
  escapeHtml,
  excelSerialToDate,
  isDateFormat,
  rowHeightToPixels,
  sanitizeFileName,
  setStatus,
  showMessageBox,
  workbookUses1904DateSystem,
} from "../ui/help-and-utilities.js";
import {
  rowHeaderWidthForSheet,
} from "../ui/ribbon-and-links.js";
import {
  activeSheetView,
  sheetHeadingsVisible,
} from "../ui/views-and-formula-tools.js";
import {
  buildLegacyExcelHtml,
  buildStandaloneHtml,
  excelCellAddress,
  excelColorObjectFromCss,
  excelPageCommentsMode,
  excelPageErrorsMode,
  excelPaperSizeCode,
  excelWorksheetViewSelectionState,
  exportSheetCellsToExcel,
  exportSheetDimensionsToExcel,
  exportSheetImagesToExcel,
  exportSheetMergesToExcel,
  exportSheetTablesToExcel,
  hasPreservedVbaProject,
  loadImageElement,
  pixelsToExcelColumnWidth,
  pixelsToExcelRowHeight,
  postProcessExcelWorkbookBuffer,
  sanitizeWorksheetName,
} from "./export-model-and-measurement.js";
import {
  clampWorkbookSheetIndex,
  dateSerialFromPersistedValue,
  isJsDateDisplayString,
  markWorkbookSavedClean,
  normalizePaneState,
  normalizeWorkbookCalcProperties,
  normalizeWorkbookViewModel,
  normalizeWorkbookWindowView,
  normalizedSheetShowRuler,
  normalizedSheetViewMode,
  persistedDateDisplayValue,
  scheduleWorkbookPersistence,
} from "./import-model.js";
import {
  base64ToBytes,
  bytesToBase64,
  parseFormulaFunctionCall,
  splitFormulaArguments,
  splitTopLevelFormulaOperator,
} from "./tables-and-validation.js";
import {
  normalizedPageLayoutSettings,
} from "./vba-and-printing.js";
import {
  saveActiveWindowViewState,
} from "./views-and-macros.js";


function formulaLambdaStripParentheses(source) {
  let text = String(source || "").trim();
  while (text.startsWith("(") && text.endsWith(")")) {
    const inner = text.slice(1, -1);
    if (splitFormulaArguments(inner).length !== 1) break;
    text = inner.trim();
  }
  return text;
}



function formulaLambdaEvaluateBinary(text, env, sheet, sheetIndex, operators) {
  const split = formulaLambdaTopLevelOperator(text, operators);
  if (!split) return null;
  const left = formulaLambdaEvaluate(text.slice(0, split.index), env, sheet, sheetIndex);
  const right = formulaLambdaEvaluate(text.slice(split.index + 1), env, sheet, sheetIndex);
  if (split.operator === "+") return formulaLambdaNumeric(left) + formulaLambdaNumeric(right);
  if (split.operator === "-") return formulaLambdaNumeric(left) - formulaLambdaNumeric(right);
  if (split.operator === "*") return formulaLambdaNumeric(left) * formulaLambdaNumeric(right);
  if (split.operator === "/") {
    const divisor = formulaLambdaNumeric(right);
    return divisor === 0 ? "#DIV/0!" : formulaLambdaNumeric(left) / divisor;
  }
  if (split.operator === "^") return formulaLambdaNumeric(left) ** formulaLambdaNumeric(right);
  return null;
}



function formulaLambdaFunctionValues(args, env, sheet, sheetIndex) {
  return (args || []).map((arg) => formulaLambdaEvaluate(arg, env, sheet, sheetIndex));
}



function formulaLambdaFlatten(values) {
  const output = [];
  values.forEach((value) => {
    if (Array.isArray(value)) output.push(...flattenFormulaMatrix(value));
    else output.push(value);
  });
  return output;
}



function formulaLambdaNumericValues(values) {
  return formulaLambdaFlatten(values)
    .map((value) => formulaLambdaNumeric(value, Number.NaN))
    .filter((value) => Number.isFinite(value));
}



function formulaLambdaEvaluateFunction(name, args, env, sheet, sheetIndex) {
  const values = formulaLambdaFunctionValues(args, env, sheet, sheetIndex);
  if (name === "SUM") return formulaLambdaNumericValues(values).reduce((sum, value) => sum + value, 0);
  if (name === "PRODUCT") return formulaLambdaNumericValues(values).reduce((product, value) => product * value, 1);
  if (name === "AVERAGE") {
    const numbers = formulaLambdaNumericValues(values);
    return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : "#DIV/0!";
  }
  if (name === "COUNT") return formulaLambdaNumericValues(values).length;
  if (name === "COUNTA") return formulaLambdaFlatten(values).filter((value) => !webSheetFormulaIsEmptyValue(webSheetFormulaRawValue(value))).length;
  if (name === "MIN") {
    const numbers = formulaLambdaNumericValues(values);
    return numbers.length ? Math.min(...numbers) : 0;
  }
  if (name === "MAX") {
    const numbers = formulaLambdaNumericValues(values);
    return numbers.length ? Math.max(...numbers) : 0;
  }
  if (name === "ROWS") return Math.max(1, formulaMatrixDimensions(Array.isArray(values[0]) ? values[0] : [[values[0]]]).height);
  if (name === "COLUMNS") return Math.max(1, formulaMatrixDimensions(Array.isArray(values[0]) ? values[0] : [[values[0]]]).width);
  if (name === "INDEX") {
    const data = Array.isArray(values[0]) ? values[0] : [[values[0]]];
    const row = Math.max(1, Math.trunc(formulaLambdaNumeric(values[1], 1))) - 1;
    const col = Math.max(1, Math.trunc(formulaLambdaNumeric(values[2], 1))) - 1;
    return formulaMatrixValue(data, row, col, "");
  }
  if (name === "IF") {
    const condition = formulaLambdaTruthy(values[0]);
    return condition ? values[1] : (args.length >= 3 ? values[2] : false);
  }
  return formulaArgumentScalarValue(`${name}(${args.join(",")})`, sheet, sheetIndex);
}



function formulaLambdaEvaluate(expression, env, sheet, sheetIndex) {
  const text = formulaLambdaStripParentheses(String(expression || "").trim());
  if (!text) return "";
  const lowered = text.toLocaleLowerCase("en-US");
  if (env.has(lowered)) return env.get(lowered);
  const quoted = text.match(/^"((?:[^"]|"")*)"$/u);
  if (quoted) return quoted[1].replace(/""/g, "\"");
  if (/^(TRUE|FALSE)$/iu.test(text)) return /^TRUE$/iu.test(text);
  const number = Number(text.replace(/,/g, ""));
  if (Number.isFinite(number) && /^[-+]?\d[\d,]*(?:\.\d+)?$/u.test(text)) return number;
  for (const operators of [["+", "-"], ["*", "/"], ["^"]]) {
    const value = formulaLambdaEvaluateBinary(text, env, sheet, sheetIndex, operators);
    if (value !== null) return value;
  }
  const call = parseFormulaFunctionCall(text);
  if (call) return formulaLambdaEvaluateFunction(call.name, call.args, env, sheet, sheetIndex);
  const range = decodeFormulaRangeForSheet(text, sheet, sheetIndex);
  if (range) return formulaRangeValueMatrix(range);
  return formulaArgumentScalarValue(text, sheet, sheetIndex) ?? "";
}



function makeArrayLambdaFormulaMatrix(args, sheet, sheetIndex) {
  const rows = Math.max(1, Math.trunc(formulaArgumentNumberValue(args[0], sheet, sheetIndex, 1)));
  const cols = Math.max(1, Math.trunc(formulaArgumentNumberValue(args[1], sheet, sheetIndex, 1)));
  const lambda = args[2] || "";
  return Array.from({ length: rows }, (_row, rowIndex) =>
    Array.from({ length: cols }, (_col, colIndex) => formulaLambdaApply(lambda, [rowIndex + 1, colIndex + 1], sheet, sheetIndex)),
  );
}



function byRowLambdaFormulaMatrix(args, sheet, sheetIndex) {
  const data = normalizeFormulaMatrix(formulaArgumentMatrix(args[0], sheet, sheetIndex));
  const lambda = args[1] || "";
  return data.map((row) => [formulaLambdaApply(lambda, [[row]], sheet, sheetIndex)]);
}



function byColLambdaFormulaMatrix(args, sheet, sheetIndex) {
  const data = normalizeFormulaMatrix(formulaArgumentMatrix(args[0], sheet, sheetIndex));
  const { width } = formulaMatrixDimensions(data);
  return [Array.from({ length: width }, (_col, colIndex) => {
    const column = data.map((row) => [row[colIndex]]);
    return formulaLambdaApply(args[1] || "", [column], sheet, sheetIndex);
  })];
}



function mapLambdaFormulaMatrix(args, sheet, sheetIndex) {
  const lambda = args[args.length - 1] || "";
  const arrays = args.slice(0, -1).map((arg) => normalizeFormulaMatrix(formulaArgumentMatrix(arg, sheet, sheetIndex)));
  const height = Math.max(1, ...arrays.map((matrix) => formulaMatrixDimensions(matrix).height));
  const width = Math.max(1, ...arrays.map((matrix) => formulaMatrixDimensions(matrix).width));
  return Array.from({ length: height }, (_row, rowIndex) =>
    Array.from({ length: width }, (_col, colIndex) =>
      formulaLambdaApply(lambda, arrays.map((matrix) => formulaMatrixValueAt(matrix, rowIndex, colIndex)), sheet, sheetIndex),
    ),
  );
}



function reduceLambdaFormulaValue(args, sheet, sheetIndex) {
  let accumulator = formulaArgumentScalarValue(args[0], sheet, sheetIndex) ?? "";
  const values = flattenFormulaMatrix(formulaArgumentMatrix(args[1], sheet, sheetIndex));
  const lambda = args[2] || "";
  values.forEach((value) => {
    accumulator = formulaLambdaApply(lambda, [accumulator, value], sheet, sheetIndex);
  });
  return accumulator;
}



function scanLambdaFormulaMatrix(args, sheet, sheetIndex) {
  let accumulator = formulaArgumentScalarValue(args[0], sheet, sheetIndex) ?? "";
  const data = normalizeFormulaMatrix(formulaArgumentMatrix(args[1], sheet, sheetIndex));
  const lambda = args[2] || "";
  return data.map((row) => row.map((value) => {
    accumulator = formulaLambdaApply(lambda, [accumulator, value], sheet, sheetIndex);
    return accumulator;
  }));
}



function sequenceFormulaMatrix(rows = 1, columns = 1, start = 1, step = 1) {
  const height = Math.max(1, Math.trunc(Number(rows) || 1));
  const width = Math.max(1, Math.trunc(Number(columns) || 1));
  const first = Number.isFinite(Number(start)) ? Number(start) : 1;
  const increment = Number.isFinite(Number(step)) ? Number(step) : 1;
  return Array.from({ length: height }, (_row, rowIndex) =>
    Array.from({ length: width }, (_col, colIndex) => first + (rowIndex * width + colIndex) * increment),
  );
}



function offsetFormulaRange(args, sheet, sheetIndex) {
  const source = decodeFormulaRangeForSheet(args[0], sheet, sheetIndex);
  if (!source) return null;
  const rowOffset = Math.trunc(formulaArgumentNumberValue(args[1], sheet, sheetIndex, 0));
  const colOffset = Math.trunc(formulaArgumentNumberValue(args[2], sheet, sheetIndex, 0));
  const sourceHeight = source.bottom - source.top + 1;
  const sourceWidth = source.right - source.left + 1;
  const height = Math.trunc(formulaArgumentNumberValue(args[3], sheet, sheetIndex, sourceHeight));
  const width = Math.trunc(formulaArgumentNumberValue(args[4], sheet, sheetIndex, sourceWidth));
  const top = source.top + rowOffset;
  const left = source.left + colOffset;
  if (top < 1 || left < 1 || height < 1 || width < 1) return null;
  return {
    sheetIndex: source.sheetIndex,
    sheetName: source.sheetName,
    top,
    left,
    bottom: top + height - 1,
    right: left + width - 1,
  };
}



function formulaFilterIncludeMatrix(expression, sheet, sheetIndex) {
  const text = stripFormulaOuterParentheses(String(expression || "").trim());
  if (!text) return [[true]];
  const comparison = splitFormulaComparisonExpression(text);
  if (comparison) {
    const leftRange = decodeFormulaRangeForSheet(comparison.left, sheet, sheetIndex);
    const rightRange = decodeFormulaRangeForSheet(comparison.right, sheet, sheetIndex);
    if (leftRange || rightRange) {
      const leftData = leftRange ? formulaRangeValueMatrix(leftRange) : null;
      const rightData = rightRange ? formulaRangeValueMatrix(rightRange) : null;
      const height = Math.max(leftData?.length || 1, rightData?.length || 1);
      const width = Math.max(
        Math.max(0, ...(leftData || [[]]).map((row) => row.length)),
        Math.max(0, ...(rightData || [[]]).map((row) => row.length)),
        1,
      );
      const leftScalar = leftRange ? undefined : formulaArgumentScalarValue(comparison.left, sheet, sheetIndex);
      const rightScalar = rightRange ? undefined : formulaArgumentScalarValue(comparison.right, sheet, sheetIndex);
      return Array.from({ length: height }, (_row, rowIndex) =>
        Array.from({ length: width }, (_col, colIndex) => {
          const left = leftRange ? formulaMatrixValueAt(leftData, rowIndex, colIndex) : leftScalar;
          const right = rightRange ? formulaMatrixValueAt(rightData, rowIndex, colIndex) : rightScalar;
          return formulaComparisonResult(left, right, comparison.operator);
        }),
      );
    }
  }
  const orParts = splitTopLevelFormulaOperator(text, "+").filter((part) => part !== "");
  if (orParts.length > 1) {
    return combineFormulaFilterIncludeMatrices(
      orParts.map((part) => formulaFilterIncludeMatrix(part, sheet, sheetIndex)),
      "or",
    );
  }
  const andParts = splitTopLevelFormulaOperator(text, "*").filter((part) => part !== "");
  if (andParts.length > 1) {
    return combineFormulaFilterIncludeMatrices(
      andParts.map((part) => formulaFilterIncludeMatrix(part, sheet, sheetIndex)),
      "and",
    );
  }
  const range = decodeFormulaRangeForSheet(text, sheet, sheetIndex);
  if (range) return formulaRangeValueMatrix(range).map((row) => row.map(formulaTruthyValue));
  return [[formulaTruthyValue(formulaArgumentScalarValue(text, sheet, sheetIndex))]];
}



function stripFormulaOuterParentheses(value) {
  let source = String(value || "").trim();
  while (source.startsWith("(") && source.endsWith(")") && formulaOuterParenthesesWrapWholeExpression(source)) {
    source = source.slice(1, -1).trim();
  }
  return source;
}



function formulaOuterParenthesesWrapWholeExpression(source) {
  let depth = 0;
  let inString = false;
  let inSheetQuote = false;
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char === "\"" && !inSheetQuote) {
      if (inString && source[index + 1] === "\"") {
        index += 1;
        continue;
      }
      inString = !inString;
      continue;
    }
    if (char === "'" && !inString) {
      if (inSheetQuote && source[index + 1] === "'") {
        index += 1;
        continue;
      }
      inSheetQuote = !inSheetQuote;
      continue;
    }
    if (inString || inSheetQuote) continue;
    if (char === "(") depth += 1;
    else if (char === ")") {
      depth -= 1;
      if (depth === 0 && index < source.length - 1) return false;
    }
    if (depth < 0) return false;
  }
  return depth === 0;
}



function combineFormulaFilterIncludeMatrices(matrices, mode = "and") {
  const normalized = matrices.map(normalizeFormulaMatrix);
  const height = Math.max(1, ...normalized.map((matrix) => matrix.length));
  const width = Math.max(1, ...normalized.flatMap((matrix) => matrix.map((row) => row.length)));
  return Array.from({ length: height }, (_row, rowIndex) =>
    Array.from({ length: width }, (_col, colIndex) => {
      const values = normalized.map((matrix) => formulaTruthyValue(formulaMatrixValueAt(matrix, rowIndex, colIndex)));
      return mode === "or" ? values.some(Boolean) : values.every(Boolean);
    }),
  );
}



function formulaMatrixValueAt(data, rowIndex, colIndex) {
  return data?.[rowIndex]?.[colIndex] ?? data?.[rowIndex]?.[0] ?? data?.[0]?.[colIndex] ?? data?.[0]?.[0] ?? "";
}



function splitFormulaComparisonExpression(text) {
  const source = String(text || "");
  let inString = false;
  let inSheetQuote = false;
  let depth = 0;
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char === "\"" && !inSheetQuote) {
      if (inString && source[index + 1] === "\"") {
        index += 1;
        continue;
      }
      inString = !inString;
      continue;
    }
    if (char === "'" && !inString) {
      if (inSheetQuote && source[index + 1] === "'") {
        index += 1;
        continue;
      }
      inSheetQuote = !inSheetQuote;
      continue;
    }
    if (inString || inSheetQuote) continue;
    if ("([{".includes(char)) depth += 1;
    else if (")]}".includes(char)) depth = Math.max(0, depth - 1);
    if (depth > 0) continue;
    const two = source.slice(index, index + 2);
    if ([">=", "<=", "<>"].includes(two)) {
      return { left: source.slice(0, index).trim(), operator: two, right: source.slice(index + 2).trim() };
    }
    if ([">", "<", "="].includes(char)) {
      return { left: source.slice(0, index).trim(), operator: char, right: source.slice(index + 1).trim() };
    }
  }
  return null;
}



function formulaComparisonResult(left, right, operator) {
  const comparison = webSheetFormulaCompareValues(left, right);
  if (!Number.isFinite(comparison)) return false;
  if (operator === "=") return comparison === 0;
  if (operator === "<>") return comparison !== 0;
  if (operator === ">") return comparison > 0;
  if (operator === ">=") return comparison >= 0;
  if (operator === "<") return comparison < 0;
  if (operator === "<=") return comparison <= 0;
  return false;
}



function formulaTruthyValue(value) {
  const raw = webSheetFormulaRawValue(value);
  if (raw instanceof CellError) return false;
  if (webSheetFormulaIsEmptyValue(raw)) return false;
  if (typeof raw === "boolean") return raw;
  if (typeof raw === "number") return raw !== 0;
  const number = Number(String(raw).trim());
  if (Number.isFinite(number)) return number !== 0;
  return String(raw).trim().toLocaleLowerCase() === "true";
}



function filterFormulaMatrix(data, include, ifEmptyValue = undefined) {
  const height = data.length;
  const width = Math.max(0, ...data.map((row) => row.length));
  const includeHeight = include.length;
  const includeWidth = Math.max(0, ...include.map((row) => row.length));
  let result = [];
  if (includeHeight === height && includeWidth === 1) {
    result = data.filter((_row, rowIndex) => formulaTruthyValue(include[rowIndex]?.[0]));
  } else if (includeHeight === 1 && includeWidth === width) {
    const keptColumns = [];
    for (let colIndex = 0; colIndex < width; colIndex += 1) {
      if (formulaTruthyValue(include[0]?.[colIndex])) keptColumns.push(colIndex);
    }
    result = data.map((row) => keptColumns.map((colIndex) => row[colIndex]));
  } else if (includeHeight === height && includeWidth === width && width === 1) {
    result = data.filter((_row, rowIndex) => formulaTruthyValue(include[rowIndex]?.[0]));
  } else if (includeHeight === height && includeWidth === width && height === 1) {
    result = [data[0].filter((_value, colIndex) => formulaTruthyValue(include[0]?.[colIndex]))];
  } else {
    result = [["#VALUE!"]];
  }
  if (result.length && Math.max(0, ...result.map((row) => row.length)) > 0) return result;
  return [[ifEmptyValue === undefined ? "#CALC!" : ifEmptyValue]];
}



function takeFormulaMatrix(data, rows, columns = undefined) {
  const normalized = normalizeFormulaMatrix(data);
  const { height, width } = formulaMatrixDimensions(normalized);
  const rowCount = Math.trunc(Number(rows));
  const colCount = columns == null ? width : Math.trunc(Number(columns));
  if (!rowCount || !colCount) return [["#CALC!"]];
  const selectedRows = rowCount > 0 ? normalized.slice(0, rowCount) : normalized.slice(Math.max(0, height + rowCount));
  return selectedRows.map((row) => (colCount > 0 ? row.slice(0, colCount) : row.slice(Math.max(0, width + colCount))));
}



function dropFormulaMatrix(data, rows, columns = undefined) {
  const normalized = normalizeFormulaMatrix(data);
  const { height, width } = formulaMatrixDimensions(normalized);
  const rowDrop = Math.trunc(Number(rows) || 0);
  const colDrop = columns == null ? 0 : Math.trunc(Number(columns) || 0);
  let rowsData = normalized;
  if (rowDrop > 0) rowsData = rowsData.slice(Math.min(height, rowDrop));
  if (rowDrop < 0) rowsData = rowsData.slice(0, Math.max(0, height + rowDrop));
  let result = rowsData;
  if (colDrop > 0) result = rowsData.map((row) => row.slice(Math.min(width, colDrop)));
  if (colDrop < 0) result = rowsData.map((row) => row.slice(0, Math.max(0, width + colDrop)));
  return result.length && Math.max(0, ...result.map((row) => row.length)) > 0 ? result : [["#CALC!"]];
}



function chooseRowsFormulaMatrix(data, selectors) {
  const normalized = normalizeFormulaMatrix(data);
  const { height } = formulaMatrixDimensions(normalized);
  const rows = selectors.map((selector) => formulaSelectorIndex(selector, height));
  if (rows.some((index) => index == null)) return [["#VALUE!"]];
  return rows.map((index) => normalized[index]);
}



function chooseColsFormulaMatrix(data, selectors) {
  const normalized = normalizeFormulaMatrix(data);
  const { width } = formulaMatrixDimensions(normalized);
  const cols = selectors.map((selector) => formulaSelectorIndex(selector, width));
  if (cols.some((index) => index == null)) return [["#VALUE!"]];
  return normalized.map((row) => cols.map((index) => row[index]));
}



function formulaSelectorIndex(selector, size) {
  const value = Math.trunc(Number(selector));
  if (!value) return null;
  const index = value > 0 ? value - 1 : size + value;
  return index >= 0 && index < size ? index : null;
}



function vstackFormulaMatrices(matrices) {
  const normalizedMatrices = matrices.map((matrix) => normalizeFormulaMatrix(matrix, "#N/A"));
  const width = Math.max(0, ...normalizedMatrices.map((matrix) => formulaMatrixDimensions(matrix).width));
  const rows = [];
  normalizedMatrices.forEach((matrix) => {
    matrix.forEach((row) => rows.push(Array.from({ length: width }, (_item, colIndex) => row[colIndex] ?? "#N/A")));
  });
  return rows.length ? rows : [["#N/A"]];
}



function hstackFormulaMatrices(matrices) {
  const normalizedMatrices = matrices.map((matrix) => normalizeFormulaMatrix(matrix, "#N/A"));
  const height = Math.max(0, ...normalizedMatrices.map((matrix) => formulaMatrixDimensions(matrix).height));
  const widths = normalizedMatrices.map((matrix) => formulaMatrixDimensions(matrix).width);
  return Array.from({ length: height }, (_row, rowIndex) =>
    normalizedMatrices.flatMap((matrix, matrixIndex) =>
      Array.from({ length: widths[matrixIndex] }, (_item, colIndex) => matrix[rowIndex]?.[colIndex] ?? "#N/A"),
    ),
  );
}



function transposeFormulaMatrix(data) {
  const normalized = normalizeFormulaMatrix(data);
  const { height, width } = formulaMatrixDimensions(normalized);
  return Array.from({ length: width }, (_row, rowIndex) =>
    Array.from({ length: height }, (_col, colIndex) => normalized[colIndex]?.[rowIndex] ?? ""),
  );
}



function formulaShouldIgnoreArrayValue(value, ignore = 0) {
  const mode = Math.trunc(Number(ignore) || 0);
  const empty = webSheetFormulaIsEmptyValue(webSheetFormulaRawValue(value));
  const error = formulaIsErrorLikeValue(value);
  return (mode === 1 && empty) || (mode === 2 && error) || (mode === 3 && (empty || error));
}



function formulaIsErrorLikeValue(value) {
  return value instanceof CellError || /^#(?:N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|CALC!|SPILL!)$/iu.test(String(value ?? ""));
}



function tocolFormulaMatrix(data, ignore = 0, scanByColumn = false) {
  const values = flattenFormulaMatrix(data, scanByColumn).filter((value) => !formulaShouldIgnoreArrayValue(value, ignore));
  return values.length ? values.map((value) => [value]) : [["#CALC!"]];
}



function torowFormulaMatrix(data, ignore = 0, scanByColumn = false) {
  const values = flattenFormulaMatrix(data, scanByColumn).filter((value) => !formulaShouldIgnoreArrayValue(value, ignore));
  return values.length ? [values] : [["#CALC!"]];
}



function wrapRowsFormulaMatrix(values, wrapCount, padWith = "#N/A") {
  const width = Math.max(1, Math.trunc(Number(wrapCount) || 1));
  const result = [];
  for (let index = 0; index < values.length; index += width) {
    const row = values.slice(index, index + width);
    while (row.length < width) row.push(padWith);
    result.push(row);
  }
  return result.length ? result : [[padWith]];
}



function wrapColsFormulaMatrix(values, wrapCount, padWith = "#N/A") {
  const height = Math.max(1, Math.trunc(Number(wrapCount) || 1));
  const colCount = Math.max(1, Math.ceil(values.length / height));
  return Array.from({ length: height }, (_row, rowIndex) =>
    Array.from({ length: colCount }, (_col, colIndex) => values[colIndex * height + rowIndex] ?? padWith),
  );
}



function expandFormulaMatrix(data, rows, columns = undefined, padWith = "#N/A") {
  const normalized = normalizeFormulaMatrix(data, padWith);
  const { height, width } = formulaMatrixDimensions(normalized);
  const targetHeight = Math.trunc(Number(rows));
  const targetWidth = columns == null ? width : Math.trunc(Number(columns));
  if (targetHeight < height || targetWidth < width || targetHeight < 1 || targetWidth < 1) return [["#VALUE!"]];
  return Array.from({ length: targetHeight }, (_row, rowIndex) =>
    Array.from({ length: targetWidth }, (_col, colIndex) => normalized[rowIndex]?.[colIndex] ?? padWith),
  );
}



function uniqueFormulaMatrix(data, byCol = false, exactlyOnce = false) {
  if (!byCol) {
    const counts = new Map();
    data.forEach((rowData) => counts.set(webSheetFormulaRowKey(rowData), (counts.get(webSheetFormulaRowKey(rowData)) || 0) + 1));
    const seen = new Set();
    const rows = [];
    for (const rowData of data) {
      const key = webSheetFormulaRowKey(rowData);
      if (exactlyOnce && counts.get(key) !== 1) continue;
      if (!exactlyOnce && seen.has(key)) continue;
      seen.add(key);
      rows.push(rowData);
    }
    return rows.length ? rows : [["#N/A"]];
  }
  const width = Math.max(0, ...data.map((rowData) => rowData.length));
  const columns = Array.from({ length: width }, (_item, colIndex) => data.map((rowData) => rowData[colIndex]));
  const counts = new Map();
  columns.forEach((column) => counts.set(webSheetFormulaRowKey(column), (counts.get(webSheetFormulaRowKey(column)) || 0) + 1));
  const seen = new Set();
  const uniqueColumns = [];
  for (const column of columns) {
    const key = webSheetFormulaRowKey(column);
    if (exactlyOnce && counts.get(key) !== 1) continue;
    if (!exactlyOnce && seen.has(key)) continue;
    seen.add(key);
    uniqueColumns.push(column);
  }
  return uniqueColumns.length ? data.map((_row, rowIndex) => uniqueColumns.map((column) => column[rowIndex])) : [["#N/A"]];
}



function sortFormulaMatrix(data, sortIndex = 1, sortOrder = 1, byCol = false) {
  const order = Number(sortOrder) < 0 ? -1 : 1;
  const index = Math.max(0, Math.trunc(Number(sortIndex) || 1) - 1);
  if (!byCol) {
    return [...data].sort((left, right) => webSheetFormulaCompareForSort(left[index], right[index]) * order);
  }
  const width = Math.max(0, ...data.map((rowData) => rowData.length));
  const columns = Array.from({ length: width }, (_item, colIndex) => data.map((rowData) => rowData[colIndex]));
  columns.sort((left, right) => webSheetFormulaCompareForSort(left[index], right[index]) * order);
  return data.map((_row, rowIndex) => columns.map((column) => column[rowIndex]));
}



function sortByFormulaMatrix(data, descriptors = []) {
  const normalized = normalizeFormulaMatrix(data);
  const activeDescriptors = descriptors.length
    ? descriptors.map((descriptor) => ({
        data: normalizeFormulaMatrix(descriptor.data),
        order: Number(descriptor.order) < 0 ? -1 : 1,
      }))
    : [{ data: normalized.map((row) => [row[0]]), order: 1 }];
  return normalized
    .map((row, index) => ({ row, index }))
    .sort((left, right) => {
      for (const descriptor of activeDescriptors) {
        const comparison = webSheetFormulaCompareForSort(
          formulaMatrixValueAt(descriptor.data, left.index, 0),
          formulaMatrixValueAt(descriptor.data, right.index, 0),
        );
        if (comparison !== 0) return comparison * descriptor.order;
      }
      return left.index - right.index;
    })
    .map((entry) => entry.row);
}



function shouldHideZeroValueForSheet(sheet, cell, value = undefined) {
  if (!cell || sheet?.showZeros !== false) return false;
  const candidates = value === undefined
    ? [cell.raw, cell.cached, cell.display]
    : [value, cell.raw, cell.cached, cell.display];
  return candidates.some((candidate) => typeof candidate === "number" && Number.isFinite(candidate) && Math.abs(candidate) < 1e-12);
}



function excelCompatibleFormulaValue(sheet, cell) {
  const summary = evaluateExcelSummaryReferenceFormula(sheet, cell?.formula);
  if (summary.supported) return summary;
  const textPart = evaluateExcelTextPartFormula(sheet, cell?.formula);
  if (textPart.supported) return textPart;
  const datedif = evaluateExcelDatedifFormula(sheet, cell?.formula);
  if (datedif.supported) return datedif;
  return { supported: false, value: null };
}



function evaluateExcelTextPartFormula(sheet, formula) {
  const match = String(formula || "").trim().match(/^=\s*(TEXTBEFORE|TEXTAFTER)\s*\((.*)\)\s*$/i);
  if (!match) return { supported: false, value: null };
  const sheetIndex = state.model?.sheets?.indexOf(sheet) ?? state.activeSheetIndex;
  const args = splitFormulaArguments(match[2]);
  const value = textPartFormulaValue(
    formulaArgumentScalarValue(args[0], sheet, sheetIndex),
    formulaArgumentScalarValue(args[1], sheet, sheetIndex),
    formulaArgumentNumberValue(args[2], sheet, sheetIndex, 1),
    formulaArgumentNumberValue(args[3], sheet, sheetIndex, 0),
    formulaArgumentBoolean(args[4], false),
    args[5] == null || String(args[5]).trim() === "" ? WEBSHEET_FORMULA_NA_SENTINEL : formulaArgumentScalarValue(args[5], sheet, sheetIndex),
    match[1].toUpperCase() === "TEXTAFTER",
  );
  return { supported: true, value };
}



function evaluateExcelSummaryReferenceFormula(sheet, formula) {
  const match = String(formula || "").trim().match(/^=\s*(SUM|AVERAGE|MAX|MIN|PRODUCT|COUNT|COUNTA|STDEV(?:\.S|\.P)?|VAR(?:\.S|\.P)?)\s*\((.*)\)\s*$/i);
  if (!match) return { supported: false, value: null };
  const method = dataSummaryMethodForFormulaName(match[1]);
  if (!method) return { supported: false, value: null };
  const values = [];
  const invalid = [];
  splitFormulaArgumentList(match[2]).forEach((token) => {
    const range = decodeWorkbookRangeOrName(token, sheet);
    if (!range) {
      invalid.push(token);
      return;
    }
    const sourceSheet = state.model?.sheets?.[range.sheetIndex ?? state.activeSheetIndex];
    if (!sourceSheet) {
      invalid.push(token);
      return;
    }
    const sourceSheetId = getSheetId(sourceSheet.name);
    forEachCellInRange(range, (row, col) => {
      const cell = sourceSheet.cells[cellKey(row, col)];
      values.push(
        cell?.formula
          ? formulaScalarValueForCell(sourceSheet, sourceSheetId, row, col)
          : cell?.raw ?? cell?.cached ?? getDisplayForCell(sourceSheet, sourceSheetId, row, col),
      );
    });
  });
  if (invalid.length) return { supported: false, value: null };
  return { supported: true, value: applySummaryFunction(values, method) };
}



function dataSummaryMethodForFormulaName(name) {
  switch (String(name || "").toUpperCase()) {
    case "AVERAGE":
      return "average";
    case "MAX":
      return "max";
    case "MIN":
      return "min";
    case "PRODUCT":
      return "product";
    case "COUNT":
      return "countNums";
    case "COUNTA":
      return "count";
    case "STDEV":
    case "STDEV.S":
      return "stdDev";
    case "STDEV.P":
      return "stdDevP";
    case "VAR":
    case "VAR.S":
      return "var";
    case "VAR.P":
      return "varP";
    case "SUM":
      return "sum";
    default:
      return "";
  }
}



function splitFormulaArgumentList(text) {
  const tokens = [];
  let current = "";
  let depth = 0;
  let inString = false;
  let inQuotedSheet = false;
  const source = String(text || "");
  const push = () => {
    const token = current.trim();
    if (token) tokens.push(token);
    current = "";
  };
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char === "\"" && !inQuotedSheet) {
      if (inString && source[index + 1] === "\"") {
        current += "\"\"";
        index += 1;
        continue;
      }
      inString = !inString;
      current += char;
      continue;
    }
    if (char === "'" && !inString) {
      if (inQuotedSheet && source[index + 1] === "'") {
        current += "''";
        index += 1;
        continue;
      }
      inQuotedSheet = !inQuotedSheet;
      current += char;
      continue;
    }
    if (!inString && !inQuotedSheet && "([{".includes(char)) {
      depth += 1;
      current += char;
      continue;
    }
    if (!inString && !inQuotedSheet && ")]}".includes(char)) {
      depth = Math.max(0, depth - 1);
      current += char;
      continue;
    }
    if (!inString && !inQuotedSheet && depth === 0 && (char === "," || char === ";")) {
      push();
      continue;
    }
    current += char;
  }
  push();
  return tokens;
}



function evaluateExcelDatedifFormula(sheet, formula) {
  const match = String(formula || "").trim().match(/^=\s*DATEDIF\s*\(\s*([^,]+?)\s*,\s*([^,]+?)\s*,\s*"([YMD]{1,2})"\s*\)\s*([+-]\s*\d+(?:\.\d+)?)?\s*$/i);
  if (!match) return { supported: false, value: null };
  const start = formulaDateValue(sheet, match[1]);
  const end = formulaDateValue(sheet, match[2]);
  if (!start || !end) return { supported: false, value: null };
  const unit = match[3].toUpperCase();
  let value = excelDatedif(start, end, unit);
  if (value == null) return { supported: false, value: null };
  const adjustment = match[4] ? Number(match[4].replace(/\s+/g, "")) : 0;
  if (Number.isFinite(adjustment)) value += adjustment;
  return { supported: true, value };
}



function formulaDateValue(baseSheet, reference) {
  const ref = parseFormulaCellReference(baseSheet, reference);
  if (!ref) return null;
  const sheet = state.model?.sheets?.[ref.sheetIndex] || baseSheet;
  const cell = sheet?.cells?.[cellKey(ref.row, ref.col)];
  const value = cell?.raw ?? cell?.cached ?? cellDisplayText(cell);
  return normalizeDateForDatedif(value);
}



function parseFormulaCellReference(baseSheet, reference) {
  const text = String(reference || "").trim().replace(/\$/g, "");
  const match = text.match(/^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?([A-Z]+[0-9]+)$/i);
  if (!match) return null;
  const sheetName = match[1]?.replace(/''/g, "'") || match[2] || "";
  const range = decodeRange(match[3]);
  if (!range || range.top !== range.bottom || range.left !== range.right) return null;
  const baseIndex = state.model?.sheets?.indexOf(baseSheet) ?? state.activeSheetIndex;
  const sheetIndex = sheetName
    ? state.model?.sheets?.findIndex((sheet) => sheet.name === sheetName)
    : baseIndex;
  if (sheetIndex == null || sheetIndex < 0) return null;
  return { sheetIndex, row: range.top, col: range.left };
}



function normalizeDateForDatedif(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return datePartsForDatedif(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = excelSerialToDate(value);
    return datePartsForDatedif(date.getFullYear(), date.getMonth() + 1, date.getDate());
  }
  const serial = autoFilterDateSerial(value);
  if (serial == null) return null;
  const date = new Date(serial * 86400000);
  return datePartsForDatedif(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate());
}



function datePartsForDatedif(year, month, day) {
  return { year: Number(year), month: Number(month), day: Number(day) };
}



function excelDatedif(start, end, unit) {
  if (compareDateParts(start, end) > 0) return null;
  switch (unit) {
    case "D":
      return datePartsToUtcDay(end) - datePartsToUtcDay(start);
    case "M":
      return fullMonthsBetween(start, end);
    case "Y":
      return fullYearsBetween(start, end);
    case "MD":
      return daysIgnoringMonthsAndYears(start, end);
    case "YM":
      return monthsIgnoringYears(start, end);
    case "YD":
      return daysIgnoringYears(start, end);
    default:
      return null;
  }
}



function compareDateParts(left, right) {
  return datePartsToUtcDay(left) - datePartsToUtcDay(right);
}



function datePartsToUtcDay(parts) {
  return Math.floor(Date.UTC(parts.year, parts.month - 1, parts.day) / 86400000);
}



function fullMonthsBetween(start, end) {
  let months = (end.year - start.year) * 12 + (end.month - start.month);
  if (end.day < start.day) months -= 1;
  return months;
}



function fullYearsBetween(start, end) {
  let years = end.year - start.year;
  if (end.month < start.month || (end.month === start.month && end.day < start.day)) years -= 1;
  return years;
}



function monthsIgnoringYears(start, end) {
  let months = end.month - start.month;
  if (end.day < start.day) months -= 1;
  if (months < 0) months += 12;
  return months;
}



function daysIgnoringMonthsAndYears(start, end) {
  if (end.day >= start.day) return end.day - start.day;
  const previousMonth = new Date(Date.UTC(end.year, end.month - 1, 0));
  return previousMonth.getUTCDate() - start.day + end.day;
}



function daysIgnoringYears(start, end) {
  let adjustedEndYear = start.year;
  const endCandidate = datePartsForDatedif(adjustedEndYear, end.month, end.day);
  if (compareDateParts(endCandidate, start) < 0) adjustedEndYear += 1;
  return datePartsToUtcDay(datePartsForDatedif(adjustedEndYear, end.month, end.day)) - datePartsToUtcDay(start);
}



function cellDisplayText(cell) {
  if (!cell) return "";
  if (cell.display != null && cell.display !== "") {
    const normalizedDisplay = normalizedCellDisplayText(cell);
    if (normalizedDisplay != null) return normalizedDisplay;
  }
  return implicitCellDisplayText(cell);
}



function normalizedCellDisplayText(cell) {
  const display = cell?.display;
  if (display == null || display === "") return null;
  if (isDateFormat(cell?.numFmt) && numericStringValueForNumberFormat(display, cell.numFmt) != null) {
    return displayValue(display, cell.numFmt);
  }
  if (display instanceof Date && !Number.isNaN(display.getTime())) {
    const serial = dateToExcelSerial(display);
    return isDateFormat(cell?.numFmt) ? displayValue(serial, cell.numFmt) : displayValue(serial, "yyyy/m/d h:mm");
  }
  const persistedDate = persistedDateDisplayValue(display);
  if (persistedDate && (isDateFormat(cell?.numFmt) || isJsDateDisplayString(display))) {
    const serial = dateSerialFromPersistedValue(cell.raw, cell.numFmt)
      ?? dateSerialFromPersistedValue(cell.cached, cell.numFmt)
      ?? dateToExcelSerial(persistedDate);
    return isDateFormat(cell?.numFmt) ? displayValue(serial, cell.numFmt) : displayValue(serial, "yyyy/m/d h:mm");
  }
  return String(display);
}



function implicitCellDisplayText(cell) {
  if (!cell) return "";
  if (cell.formula) {
    return cell.cached != null ? displayValue(cell.cached, cell.numFmt) : "";
  }
  if (cell.raw != null) return displayValue(cell.raw, cell.numFmt);
  if (cell.cached != null) return displayValue(cell.cached, cell.numFmt);
  return "";
}



function getCellFormula(sheetId, row, col) {
  if (!state.hf || sheetId == null) return null;
  try {
    return state.hf.getCellFormula({ sheet: sheetId, row: row - 1, col: col - 1 });
  } catch {
    return null;
  }
}



function getSheetId(sheetName) {
  return state.hf?.getSheetId(sheetName);
}



function activeSheet() {
  return state.model?.sheets?.[state.activeSheetIndex] || null;
}



function buildMergeMaps(merges) {
  const mergeList = Array.isArray(merges) ? merges : [];
  if (!mergeList.length) {
    return { master: new Map(), covered: new Set(), coveredToMaster: new Map() };
  }
  const cached = sheetMergeMapsCacheByMerges.get(mergeList);
  if (cached?.length === mergeList.length) return cached.maps;
  const signature = mergeRangesSignature(mergeList);
  if (cached?.signature === signature) return cached.maps;

  const master = new Map();
  const covered = new Set();
  const coveredToMaster = new Map();
  mergeList.forEach((merge) => {
    const masterKey = cellKey(merge.top, merge.left);
    master.set(masterKey, merge);
    for (let row = merge.top; row <= merge.bottom; row += 1) {
      for (let col = merge.left; col <= merge.right; col += 1) {
        const key = cellKey(row, col);
        if (key !== masterKey) {
          covered.add(key);
          coveredToMaster.set(key, masterKey);
        }
      }
    }
  });
  const maps = { master, covered, coveredToMaster };
  sheetMergeMapsCacheByMerges.set(mergeList, { length: mergeList.length, signature, maps });
  return maps;
}



function mergeRangesSignature(merges) {
  let hash = 2166136261;
  merges.forEach((merge) => {
    hash = mergeRangeHashStep(hash, merge?.top);
    hash = mergeRangeHashStep(hash, merge?.left);
    hash = mergeRangeHashStep(hash, merge?.bottom);
    hash = mergeRangeHashStep(hash, merge?.right);
  });
  return `${merges.length}:${hash >>> 0}`;
}



function mergeRangeHashStep(hash, value) {
  const number = Math.max(0, Math.round(Number(value) || 0));
  return Math.imul((hash ^ number) >>> 0, 16777619) >>> 0;
}



function visibleGridRange(sheet, scrollOverride = null) {
  const host = $sheetHost[0];
  if (!host) {
    return {
      top: 1,
      left: 1,
      bottom: Math.min(sheet.rowCount, FALLBACK_VISIBLE_ROWS),
      right: Math.min(sheet.colCount, FALLBACK_VISIBLE_COLS),
    };
  }
  const scrollTop = Number.isFinite(scrollOverride?.scrollTop) ? scrollOverride.scrollTop : host.scrollTop;
  const scrollLeft = Number.isFinite(scrollOverride?.scrollLeft) ? scrollOverride.scrollLeft : host.scrollLeft;

  return {
    ...visibleRowRange(sheet, scrollTop, scrollTop + host.clientHeight, VIRTUAL_ROW_OVERSCAN),
    ...visibleColumnRange(sheet, scrollLeft, scrollLeft + host.clientWidth, VIRTUAL_COL_OVERSCAN),
  };
}



function visibleColumnRange(sheet, viewportStart, viewportEnd, overscan) {
  if (sheetColumnsRightToLeft(sheet)) {
    return visibleColumnRangeRightToLeft(sheet, viewportStart, viewportEnd, overscan);
  }
  let left = 1;
  let right = sheet.colCount;
  const rowHeaderWidth = sheetHeadingsVisible(sheet) ? rowHeaderWidthForSheet(sheet) : 0;
  let x = rowHeaderWidth;

  for (let col = 1; col <= sheet.colCount; col += 1) {
    const width = sheet.columns[col - 1]?.hidden ? 0 : sheet.columns[col - 1]?.width || 0;
    const next = x + width;
    if (next >= viewportStart) {
      left = Math.max(1, col - overscan);
      break;
    }
    x = next;
  }

  x = rowHeaderWidth;
  for (let col = 1; col <= sheet.colCount; col += 1) {
    const width = sheet.columns[col - 1]?.hidden ? 0 : sheet.columns[col - 1]?.width || 0;
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
  let left = Number(sheet?.colCount || 1);
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
  return {
    left: Math.max(1, left - overscan),
    right: Math.min(sheet.colCount, right + overscan),
  };
}



function visibleRowRange(sheet, viewportStart, viewportEnd, overscan) {
  let top = 1;
  let bottom = sheet.rowCount;
  const columnHeaderHeight = sheetHeadingsVisible(sheet) ? COLUMN_HEADER_HEIGHT : 0;
  let y = columnHeaderHeight;

  for (let row = 1; row <= sheet.rowCount; row += 1) {
    const height = sheet.rows[row - 1]?.hidden ? 0 : sheet.rows[row - 1]?.height || 0;
    const next = y + height;
    if (next >= viewportStart) {
      top = Math.max(1, row - overscan);
      break;
    }
    y = next;
  }

  y = columnHeaderHeight;
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    const height = sheet.rows[row - 1]?.hidden ? 0 : sheet.rows[row - 1]?.height || 0;
    const next = y + height;
    if (y > viewportEnd) {
      bottom = Math.min(sheet.rowCount, row + overscan);
      break;
    }
    y = next;
  }

  return { top, bottom: Math.max(top, bottom) };
}



function gridTrackStyle(sheet) {
  const sheetColumns = sheetColumnsRightToLeft(sheet) ? [...sheet.columns].reverse() : sheet.columns;
  const columns = [`${sheetHeadingsVisible(sheet) ? rowHeaderWidthForSheet(sheet) : 0}px`, ...sheetColumns.map((column) => `${column.hidden ? 0 : column.width}px`)];
  const rows = [`${sheetHeadingsVisible(sheet) ? COLUMN_HEADER_HEIGHT : 0}px`, ...sheet.rows.map((row) => `${row.hidden ? 0 : row.height}px`)];
  return `grid-template-columns:${columns.join(" ")};grid-template-rows:${rows.join(" ")}`;
}



function sheetGridClassName(sheet) {
  const viewMode = ["page-layout", "page-break"].includes(sheet?.viewMode) ? sheet.viewMode : "normal";
  return `sheet-grid view-${viewMode}${sheet.showGridLines === false ? " hide-gridlines" : ""}${sheetHeadingsVisible(sheet) ? "" : " hide-headings"}${sheet.showOutlineSymbols === false ? " hide-outline-symbols" : ""}${sheet.showWhiteSpace === false ? " hide-white-space" : ""}${normalizedSheetShowRuler(sheet.showRuler, viewMode) ? " show-ruler" : ""}${sheet.rightToLeft === true ? " is-right-to-left" : ""}${activeSheetView(sheet) ? " has-sheet-view" : ""}`;
}



function updateHorizontalScrollbar() {
  requestAnimationFrame(() => {
    const host = $sheetHost[0];
    const scrollbar = $horizontalScrollbar[0];
    const spacer = $horizontalScrollSpacer[0];
    if (!host || !scrollbar || !spacer || $sheetWorkspace.hasClass("hidden")) return;

    const maxScrollLeft = Math.max(0, host.scrollWidth - host.clientWidth);
    spacer.style.width = `${scrollbar.clientWidth + maxScrollLeft}px`;
    forceSyncHorizontalScrollbarFromHost();
  });
}



function forceSyncHorizontalScrollbarFromHost() {
  const host = $sheetHost[0];
  const scrollbar = $horizontalScrollbar[0];
  if (!host || !scrollbar) return;
  if (scrollbar.scrollLeft === host.scrollLeft) {
    state.syncingScroll = false;
    return;
  }
  state.syncingScroll = true;
  scrollbar.scrollLeft = host.scrollLeft;
  requestAnimationFrame(() => {
    state.syncingScroll = false;
  });
}



function syncHorizontalScrollbarFromHost() {
  if (state.syncingScroll) return;
  state.syncingScroll = true;
  $horizontalScrollbar[0].scrollLeft = $sheetHost[0].scrollLeft;
  requestAnimationFrame(() => {
    state.syncingScroll = false;
  });
}



function syncHostFromHorizontalScrollbar() {
  if (state.syncingScroll) return;
  state.syncingScroll = true;
  $sheetHost[0].scrollLeft = $horizontalScrollbar[0].scrollLeft;
  requestAnimationFrame(() => {
    state.syncingScroll = false;
  });
}



function startSheetFooterResize(event) {
  event.preventDefault();
  const footer = $sheetFooter[0];
  if (!footer) return;

  const startX = event.clientX;
  const tabsPane = footer.querySelector(".sheet-tabs-pane");
  const startWidth = tabsPane?.getBoundingClientRect().width || 280;
  const footerWidth = footer.getBoundingClientRect().width;
  const minWidth = 140;
  const maxWidth = Math.max(minWidth, footerWidth - 180);

  $sheetFooter.addClass("is-resizing");

  const onPointerMove = (moveEvent) => {
    const nextWidth = Math.max(minWidth, Math.min(maxWidth, startWidth + moveEvent.clientX - startX));
    footer.style.setProperty("--sheet-tabs-width", `${Math.round(nextWidth)}px`);
    updateHorizontalScrollbar();
  };

  const onPointerUp = () => {
    $sheetFooter.removeClass("is-resizing");
    window.removeEventListener("pointermove", onPointerMove);
    window.removeEventListener("pointerup", onPointerUp);
    updateHorizontalScrollbar();
  };

  window.addEventListener("pointermove", onPointerMove);
  window.addEventListener("pointerup", onPointerUp);
}



function renderBorderLayer(sheet, mergeMaps, visibleRange = null) {
  return borderLinesForSheet(sheet, mergeMaps, visibleRange).map(borderLineHtml).join("");
}



function paintCanvasBorders(context, sheet, mergeMaps, visibleRange = null, pixelCache = null) {
  const lines = borderLinesForSheet(sheet, mergeMaps, visibleRange, pixelCache);
  context.save();
  lines.forEach((line) => {
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



function canvasBorderDash(style) {
  if (style === "dashed") return [6, 4];
  if (style === "dotted") return [1, 3];
  return [];
}



function sheetBorderCellsByRow(sheet) {
  const revision = Math.max(0, Number(state.workbookRevision) || 0);
  const cached = sheetBorderCellCacheBySheet.get(sheet);
  if (cached?.revision === revision && cached.cells === sheet.cells) return cached;

  const byRow = new Map();
  const byKey = new Map();
  Object.values(sheet.cells || {}).forEach((cell) => {
    const borders = cell?.borders || {};
    if (!Object.keys(borders).length) return;
    const row = Math.max(1, Math.floor(Number(cell.row) || 0));
    const col = Math.max(1, Math.floor(Number(cell.col) || 0));
    if (!row || !col) return;
    if (!byRow.has(row)) byRow.set(row, []);
    byRow.get(row).push(cell);
    byKey.set(cellKey(row, col), cell);
  });
  const next = { revision, cells: sheet.cells, byRow, byKey };
  sheetBorderCellCacheBySheet.set(sheet, next);
  return next;
}



function borderLinesForSheet(sheet, mergeMaps, visibleRange = null, pixelCache = null) {
  const lines = new Map();
  const addCellBorders = (cell) => {
    const key = cellKey(cell.row, cell.col);
    const merge = mergeMaps.master.get(key);
    if (visibleRange && !rangeIntersects(merge || { top: cell.row, left: cell.col, bottom: cell.row, right: cell.col }, visibleRange)) {
      return;
    }
    addCellBorderLines(lines, sheet, mergeMaps, cell, pixelCache);
  };

  const borderCellCache = sheetBorderCellsByRow(sheet);
  if (visibleRange) {
    for (let row = visibleRange.top; row <= visibleRange.bottom; row += 1) {
      (borderCellCache.byRow.get(row) || []).forEach(addCellBorders);
    }
    (sheet.merges || []).forEach((merge) => {
      if (merge.top >= visibleRange.top && merge.top <= visibleRange.bottom) return;
      if (!rangeIntersects(merge, visibleRange)) return;
      const master = borderCellCache.byKey.get(cellKey(merge.top, merge.left));
      if (master) addCellBorders(master);
    });
  } else {
    borderCellCache.byRow.forEach((cells) => cells.forEach(addCellBorders));
  }

  if (visibleRange && Array.isArray(sheet.styleRanges) && sheet.styleRanges.length) {
    for (let row = visibleRange.top; row <= visibleRange.bottom; row += 1) {
      for (let col = visibleRange.left; col <= visibleRange.right; col += 1) {
        const key = cellKey(row, col);
        if (sheet.cells[key]) continue;
        const cell = sheetStyleRangeCell(sheet, row, col);
        if (cell) addCellBorders(cell);
      }
    }
  }

  return [...lines.values()];
}



function addCellBorderLines(lines, sheet, mergeMaps, cell, pixelCache = null) {
  const borders = cell?.borders || {};
  if (!Object.keys(borders).length) return;

  const key = cellKey(cell.row, cell.col);
  if (mergeMaps.covered.has(key)) return;

  const xPosition = (columnBoundary) => cachedCoordinateToPixels(sheet, pixelCache, "col", columnBoundary - 1);
  const yPosition = (rowBoundary) => cachedCoordinateToPixels(sheet, pixelCache, "row", rowBoundary - 1);
  const merge = mergeMaps.master.get(key);
  const leftCol = cell.col;
  const rightCol = merge ? merge.right : cell.col;
  const topRow = cell.row;
  const bottomRow = merge ? merge.bottom : cell.row;
  const leftX = xPosition(leftCol);
  const rightX = xPosition(rightCol + 1);
  const x1 = Math.min(leftX, rightX);
  const x2 = Math.max(leftX, rightX);
  const y1 = yPosition(topRow);
  const y2 = yPosition(bottomRow + 1);

  addBorderLine(lines, "h", y1, x1, x2, borders.top);
  addBorderLine(lines, "h", y2, x1, x2, borders.bottom);
  addBorderLine(lines, "v", x1, y1, y2, borders.left);
  addBorderLine(lines, "v", x2, y1, y2, borders.right);
}



function textOverflowWidth(sheet, mergeMaps, row, col, colSpan, cell, display) {
  if (!display || cellShowsValidationDropdown(cell, sheet)) return null;
  const css = cellEffectiveCss(sheet, cell);
  if (css.whiteSpace && css.whiteSpace !== "nowrap") return null;
  if (css.textAlign === "center" || css.textAlign === "right") return null;

  const startRight = col + colSpan - 1;
  let endCol = startRight;
  for (let nextCol = startRight + 1; nextCol <= sheet.colCount; nextCol += 1) {
    if (isBlockingCell(sheet, mergeMaps, row, nextCol)) break;
    endCol = nextCol;
  }

  if (endCol === startRight) return null;
  return widthBetweenColumns(sheet, col, endCol);
}



function isBlockingCell(sheet, mergeMaps, row, col) {
  const key = cellKey(row, col);
  const masterKey = mergeMaps.coveredToMaster.get(key) || key;
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



function heightBetweenRows(sheet, startRow, endRow) {
  let height = 0;
  for (let row = startRow; row <= endRow; row += 1) {
    height += sheet.rows[row - 1]?.hidden ? 0 : sheet.rows[row - 1]?.height || 0;
  }
  return height;
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
  let y = COLUMN_HEADER_HEIGHT;
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    positions[row] = y;
    y += sheet.rows[row - 1]?.hidden ? 0 : sheet.rows[row - 1]?.height || 0;
  }
  positions[sheet.rowCount + 1] = y;
  return positions;
}



function addBorderLine(lines, orientation, position, start, end, border) {
  if (!border || end <= start) return;
  const key = `${orientation}:${position}:${start}:${end}`;
  const candidate = { orientation, position, start, end, border };
  const existing = lines.get(key);
  if (!existing || compareBorder(candidate.border, existing.border) > 0) {
    lines.set(key, candidate);
  }
}



function compareBorder(a, b) {
  const weight = (a.weight || 0) - (b.weight || 0);
  if (weight) return weight;
  return (a.width || 0) - (b.width || 0);
}



function borderLineHtml(line, className = "sheet-border-line", offsetLeft = 0, offsetTop = 0) {
  const width = line.border.width || 1;
  const style = line.border.style || "solid";
  const color = line.border.color || "#000000";
  const css =
    line.orientation === "h"
      ? `left:${roundCssPx(line.start - offsetLeft)}px;top:${roundCssPx(line.position - offsetTop - width / 2)}px;width:${roundCssPx(line.end - line.start)}px;border-top:${width}px ${style} ${color}`
      : `left:${roundCssPx(line.position - offsetLeft - width / 2)}px;top:${roundCssPx(line.start - offsetTop)}px;height:${roundCssPx(line.end - line.start)}px;border-left:${width}px ${style} ${color}`;
  return `<div class="${escapeAttr(className)}" style="${escapeAttr(css)}"></div>`;
}



function displayValue(value, numFmt = "") {
  if (value == null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    if (isElapsedTimeFormat(numFmt)) return formatElapsedTimeForFormat(dateToExcelSerial(value), numFmt);
    return formatDateForFormat(value, numFmt);
  }
  if (typeof value === "object") {
    if (value.value) return String(value.value);
    if (value.error) return String(value.error);
    return String(value);
  }

  if (typeof value === "number") {
    const primaryFormat = primaryExcelFormatSection(numFmt);
    if (isElapsedTimeFormat(primaryFormat)) return formatElapsedTimeForFormat(value, primaryFormat);
    if (isDateFormat(primaryFormat)) return formatDateForFormat(excelSerialToDate(value), primaryFormat);
    if (/%/.test(primaryFormat)) return `${formatNumberForFormat(value * 100, primaryFormat)}%`;
    const currency = currencySymbol(primaryFormat);
    if (currency) return `${currency}${formatNumberForFormat(value, primaryFormat)}`;
    if (primaryFormat) return formatNumberForFormat(value, primaryFormat);
    return trimNumber(value);
  }

  if (typeof value === "string") {
    const numericValue = numericStringValueForNumberFormat(value, numFmt);
    if (numericValue != null) return displayValue(numericValue, numFmt);
  }

  return String(value);
}



function numericStringValueForNumberFormat(value, numFmt = "") {
  const primaryFormat = primaryExcelFormatSection(numFmt).trim();
  if (!primaryFormat || primaryFormat === "@") return null;
  if (!numberFormatCanApplyToNumericString(primaryFormat)) return null;
  const text = String(value ?? "").trim().replace(/,/g, "");
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(text)) return null;
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}



function numberFormatCanApplyToNumericString(format) {
  return isDateFormat(format) || /%/.test(format) || Boolean(currencySymbol(format)) || /[0#?]/.test(format);
}



function isElapsedTimeFormat(numFmt) {
  return /\[(?:h+|m+|s+)\]/i.test(primaryExcelFormatSection(numFmt || ""));
}



function formatElapsedTimeForFormat(value, numFmt = "") {
  const number = Number(value);
  if (!Number.isFinite(number)) return String(value);
  const format = primaryExcelFormatSection(numFmt);
  if (!isElapsedTimeFormat(format)) return String(value);
  const negative = number < 0;
  const absoluteSeconds = Math.abs(number) * 86400;
  const plainFormat = format.replace(/"[^"]*"/g, "").replace(/\\./g, " ");
  const hasSeconds = /\[(?:s+)\]|s{1,2}/i.test(plainFormat);
  const hasMinutes = /\[(?:m+)\]|m{1,2}/i.test(plainFormat);
  const totalSeconds = hasSeconds
    ? Math.round(absoluteSeconds)
    : hasMinutes
      ? Math.round(absoluteSeconds / 60) * 60
      : Math.round(absoluteSeconds / 3600) * 3600;
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds;
  const clockHour = hours % 24;
  const clockMinute = minutes % 60;
  const clockSecond = seconds % 60;
  const pad2 = (item) => String(item).padStart(2, "0");
  const rendered = format.replace(/"([^"]*)"|\\(.)|\[(h+|m+|s+)\]|hh|ss|mm|h|s|m/gi, (match, quoted, escaped, elapsed) => {
    if (quoted != null) return quoted;
    if (escaped != null) return escaped;
    if (elapsed) {
      const unit = elapsed[0].toLowerCase();
      return String(unit === "h" ? hours : unit === "m" ? minutes : seconds);
    }
    const token = match.toLowerCase();
    if (token === "hh") return pad2(clockHour);
    if (token === "h") return String(clockHour);
    if (token === "mm") return pad2(clockMinute);
    if (token === "m") return String(clockMinute);
    if (token === "ss") return pad2(clockSecond);
    return String(clockSecond);
  });
  return `${negative ? "-" : ""}${rendered}`;
}



function formatDateForFormat(date, numFmt = "") {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const format = primaryExcelFormatSection(numFmt).replace(/\[[^\]]+\]/g, "");
  if (!isDateFormat(format)) return date.toLocaleDateString();
  const monthNamesShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const monthNamesLong = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];
  const weekdayNamesShort = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const weekdayNamesLong = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const pad2 = (number) => String(number).padStart(2, "0");
  const uses12HourClock = /AM\/PM|A\/P/i.test(format);
  const hour24 = date.getHours();
  const hour12 = hour24 % 12 || 12;
  const replacements = {
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
    "a/p": hour24 < 12 ? "a" : "p",
  };
  return format.replace(/"([^"]*)"|\\(.)|AM\/PM|A\/P|yyyy|mmmm|dddd|mmm|ddd|yy|hh|ss|mm|dd|h|s|m|d/gi, (match, quoted, escaped, offset) => {
    if (quoted != null) return quoted;
    if (escaped != null) return escaped;
    const key = match.toLowerCase();
    if (key === "m" || key === "mm") {
      return key === "mm" && dateFormatTokenIsMinute(format, offset, match.length)
        ? pad2(date.getMinutes())
        : key === "m" && dateFormatTokenIsMinute(format, offset, match.length)
          ? String(date.getMinutes())
          : replacements[key];
    }
    return replacements[match] ?? replacements[key] ?? match;
  }).replace(/"/g, "");
}



function dateFormatTokenIsMinute(format, start, length) {
  const source = String(format || "");
  const before = source.slice(Math.max(0, start - 8), start);
  const after = source.slice(start + length, Math.min(source.length, start + length + 8));
  const previous = source.slice(0, start).match(/\S$/)?.[0] || "";
  const next = source.slice(start + length).match(/\S/)?.[0] || "";
  if (previous === ":" || next === ":") return true;
  return /h{1,2}[^ymdhs]*$/i.test(before) || /^[^ymdhms]*s{1,2}/i.test(after);
}



function primaryExcelFormatSection(numFmt = "") {
  const text = String(numFmt || "");
  let quoted = false;
  let escaped = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (char === "\\") {
      escaped = true;
      continue;
    }
    if (char === '"') {
      quoted = !quoted;
      continue;
    }
    if (char === ";" && !quoted) {
      return text.slice(0, index);
    }
  }
  return text;
}



function formatNumberForFormat(value, numFmt = "") {
  if (!Number.isFinite(value)) return String(value);
  const format = primaryExcelFormatSection(numFmt);
  const decimals = formatDecimalPlaces(format);
  return value.toLocaleString(undefined, {
    useGrouping: formatUsesSeparator(format),
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals || 10,
  });
}



function trimNumber(value) {
  if (!Number.isFinite(value)) return String(value);
  return Number(value.toFixed(10)).toLocaleString(undefined, {
    maximumFractionDigits: 10,
  });
}



function openSaveAsDialog() {
  if (!state.model) return;
  closeDataDialog({ keepState: true });
  state.dataDialog = {
    type: "save-as",
    fileHandle: null,
    directoryHandle: null,
    pickedName: "",
    location: "downloads",
  };
  const overlay = $(saveAsDialogHtml());
  $("body").append(overlay);
  syncSaveAsDialogType(overlay[0].querySelector("[data-data-dialog='save-as']"));
  overlay.find("[data-data-field='saveAsFileName']").trigger("focus").trigger("select");
}



function saveAsDialogHtml() {
  const baseName = saveAsBaseFileName();
  const defaultType = defaultSaveAsFileType();
  const typeOptions = SAVE_AS_FILE_TYPES.map(
    (type) => `<option value="${escapeAttr(type.value)}"${type.value === defaultType.value ? " selected" : ""}>${escapeHtml(type.label)}</option>`,
  ).join("");
  const locations = SAVE_AS_LOCATIONS.map((location) => `
    <button class="save-as-place${location.value === "downloads" ? " is-active" : ""}" type="button" data-save-as-location="${escapeAttr(location.value)}">
      <span>${escapeHtml(location.label)}</span>
      <small>${escapeHtml(location.detail)}</small>
    </button>`).join("");
  const pathOptions = SAVE_AS_LOCATIONS.map(
    (location) => `<option value="${escapeAttr(location.value)}"${location.value === "downloads" ? " selected" : ""}>${escapeHtml(location.path || location.label)}</option>`,
  ).join("");
  const fileRows = saveAsFileListRowsHtml(baseName, defaultType.value);
  return `
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog save-as-dialog" role="dialog" aria-modal="true" aria-labelledby="saveAsTitle" data-data-dialog="save-as">
        <div class="excel-mini-titlebar" id="saveAsTitle">名前を付けて保存</div>
        <div class="excel-mini-body data-dialog-body save-as-body">
          <aside class="save-as-places" aria-label="保存場所">${locations}</aside>
          <section class="save-as-form">
            <label class="excel-dialog-field save-as-location-row">
              <span class="excel-dialog-caption">保存先(I):</span>
              <input type="text" data-data-field="saveAsLocation" value="ダウンロード" readonly />
              <button class="excel-mini-button" type="button" data-data-dialog-action="save-as-browse">参照(B)...</button>
            </label>
            <div class="save-as-path-row">
              <span class="excel-dialog-caption">パス(P):</span>
              <select data-save-as-path data-data-field="saveAsPath">${pathOptions}</select>
              <button class="excel-mini-button" type="button" data-data-dialog-action="save-as-new-folder">新しいフォルダー(N)</button>
            </div>
            <div class="save-as-file-list" role="grid" aria-label="ファイル一覧">
              <div class="save-as-file-header" role="row">
                <span>名前</span>
                <span>更新日時</span>
                <span>種類</span>
                <span>サイズ</span>
              </div>
              <div class="save-as-file-rows">${fileRows}</div>
            </div>
            <label class="excel-dialog-field">
              <span class="excel-dialog-caption">ファイル名(N):</span>
              <input type="text" data-data-field="saveAsFileName" value="${escapeAttr(baseName)}" />
            </label>
            <label class="excel-dialog-field">
              <span class="excel-dialog-caption">ファイルの種類(T):</span>
              <select data-save-as-type data-data-field="saveAsType">${typeOptions}</select>
            </label>
            <div class="save-as-note" data-save-as-note></div>
          </section>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="save-as-ok">保存(S)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`;
}



function saveAsFileListRowsHtml(baseName = saveAsBaseFileName(), activeTypeValue = defaultSaveAsFileType().value) {
  const now = new Date();
  const updated = `${now.getFullYear()}/${String(now.getMonth() + 1).padStart(2, "0")}/${String(now.getDate()).padStart(2, "0")} ${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
  const rows = [
    { type: "html", name: fileNameWithExtension(baseName, "html"), kind: "WebSheet HTML", size: "小" },
    { type: "xlsx", name: fileNameWithExtension(baseName, "xlsx"), kind: "Excel ブック", size: "自動" },
    { type: "xlsm", name: fileNameWithExtension(baseName, "xlsm"), kind: "Excel マクロ有効ブック", size: hasPreservedVbaProject(state.model) ? "VBA 保持" : "自動" },
    { type: "xls", name: fileNameWithExtension(baseName, "xls"), kind: "Excel 97-2003 ブック", size: "互換" },
  ];
  return rows
    .map(
      (row) => `
        <button class="save-as-file-row${row.type === activeTypeValue ? " is-active" : ""}" type="button" role="row" data-save-as-file-row data-save-as-file-type="${escapeAttr(row.type)}" data-save-as-file-name="${escapeAttr(row.name)}">
          <span>${escapeHtml(row.name)}</span>
          <span>${escapeHtml(updated)}</span>
          <span>${escapeHtml(row.kind)}</span>
          <span>${escapeHtml(row.size)}</span>
        </button>`,
    )
    .join("");
}



function saveAsBaseFileName() {
  const sourceName = state.model?.sourceName || "workbook";
  return sanitizeFileName(String(sourceName).replace(/\.[^.]+$/, "")) || "workbook";
}



function saveTypeForSourceName(sourceName = "", fallback = "html") {
  const extension = String(sourceName || "").trim().split(".").pop()?.toLowerCase() || "";
  if (isWorkbookSaveType(extension)) return extension;
  return fallback;
}



function shouldUseSourceSaveTypeForCurrentDocument() {
  return Boolean(
    state.isStandalone &&
      (location.protocol === "file:" ||
        location.origin === "null" ||
        window.__WEBSHEET_STANDALONE_MODEL__ ||
        window.__WEBSHEET_STANDALONE_MODEL_RAW__),
  );
}



function defaultSaveTypeForLoadedFile(sourceName = "") {
  return shouldUseSourceSaveTypeForCurrentDocument() ? saveTypeForSourceName(sourceName, "xlsx") : "html";
}



function applyOpenedWorkbookSaveTarget(file, fileHandle) {
  if (!state.model || !fileHandle?.createWritable || !shouldUseSourceSaveTypeForCurrentDocument()) return false;
  const saveType = normalizeWorkbookSaveType(state.currentSaveType || state.model.saveType, file?.name || state.model.sourceName);
  const type = saveAsFileTypeByValue(saveType);
  const fileName = fileNameWithExtension(fileHandle.name || file?.name || state.model.sourceName || saveAsBaseFileName(), type.extension);
  state.model.sourceName = fileName;
  state.model.saveType = type.value;
  state.currentSaveType = type.value;
  state.currentSaveFileHandle = fileHandle;
  state.currentSaveDirectoryHandle = null;
  state.currentSaveFileName = fileName;
  return true;
}



function quickSaveTypeForCurrentContext() {
  return shouldUseSourceSaveTypeForCurrentDocument() ? normalizeWorkbookSaveType(state.currentSaveType || state.model?.saveType, state.model?.sourceName) : "html";
}



function defaultSaveAsFileType() {
  return saveAsFileTypeByValue(state.model ? quickSaveTypeForCurrentContext() : "html");
}



function normalizeWorkbookSaveType(value, sourceName = "") {
  const candidate = String(value || "").toLowerCase();
  if (isWorkbookSaveType(candidate)) return candidate;
  return saveTypeForSourceName(sourceName, "html");
}



function isWorkbookSaveType(value) {
  return value === "html" || value === "xlsx" || value === "xlsm" || value === "xls";
}



function currentWorkbookSaveType() {
  const type = quickSaveTypeForCurrentContext();
  if (state.model) state.model.saveType = type;
  state.currentSaveType = type;
  return type;
}



function saveAsFileTypeByValue(value) {
  return SAVE_AS_FILE_TYPES.find((type) => type.value === value) || SAVE_AS_FILE_TYPES[0];
}



function handleSaveAsTypeChange(event) {
  if (state.dataDialog?.fileHandle) {
    state.dataDialog = { ...state.dataDialog, fileHandle: null, pickedName: "" };
  }
  syncSaveAsDialogType(event.currentTarget.closest("[data-data-dialog='save-as']"));
}



function handleSaveAsPathChange(event) {
  const select = event.currentTarget;
  const dialog = select.closest("[data-data-dialog='save-as']");
  if (!dialog) return;
  setSaveAsLocation(dialog, select.value || "downloads", { browse: false });
}



function handleSaveAsLocationClick(event) {
  const button = event.currentTarget;
  const dialog = button.closest("[data-data-dialog='save-as']");
  if (!dialog) return;
  const locationValue = button.dataset.saveAsLocation || "downloads";
  setSaveAsLocation(dialog, locationValue, { browse: locationValue === "this-pc" });
}



function handleSaveAsFileRowClick(event) {
  const row = event.currentTarget;
  const dialog = row.closest("[data-data-dialog='save-as']");
  if (!dialog) return;
  const typeValue = row.dataset.saveAsFileType || "html";
  const fileName = row.dataset.saveAsFileName || "";
  const typeSelect = dialog.querySelector("[data-data-field='saveAsType']");
  const fileNameInput = dialog.querySelector("[data-data-field='saveAsFileName']");
  if (typeSelect) typeSelect.value = typeValue;
  if (fileNameInput && fileName) fileNameInput.value = fileName;
  dialog.querySelectorAll("[data-save-as-file-row]").forEach((item) => item.classList.toggle("is-active", item === row));
  syncSaveAsDialogType(dialog, { preserveFileName: true });
}



function setSaveAsLocation(dialog, locationValue, options = {}) {
  const location = saveAsLocationByValue(locationValue);
  const nextValue = location.value;
  state.dataDialog = {
    ...(state.dataDialog || {}),
    type: "save-as",
    location: nextValue,
    fileHandle: options.keepHandle ? state.dataDialog?.fileHandle || null : null,
    pickedName: options.keepHandle ? state.dataDialog?.pickedName || "" : "",
    directoryHandle: options.keepDirectory ? state.dataDialog?.directoryHandle || null : null,
  };
  updateSaveAsLocationUi(dialog, nextValue);
  syncSaveAsDialogType(dialog, { preserveFileName: true });
  if (options.browse) {
    void chooseSaveAsLocation(dialog);
  }
}



function saveAsLocationByValue(value) {
  return SAVE_AS_LOCATIONS.find((item) => item.value === value) || SAVE_AS_LOCATIONS[0];
}



function updateSaveAsLocationUi(dialog, locationValue, labelOverride = "") {
  if (!dialog) return;
  const location = saveAsLocationByValue(locationValue);
  dialog.querySelectorAll("[data-save-as-location]").forEach((item) => {
    item.classList.toggle("is-active", item.dataset.saveAsLocation === location.value);
  });
  const pathSelect = dialog.querySelector("[data-data-field='saveAsPath']");
  if (pathSelect) pathSelect.value = location.value;
  const locationInput = dialog.querySelector("[data-data-field='saveAsLocation']");
  if (locationInput) locationInput.value = labelOverride || location.label || "ダウンロード";
}



function syncSaveAsDialogType(dialog, options = {}) {
  if (!dialog) return;
  const type = saveAsFileTypeByValue(dialog.querySelector("[data-data-field='saveAsType']")?.value);
  const fileNameInput = dialog.querySelector("[data-data-field='saveAsFileName']");
  const locationInput = dialog.querySelector("[data-data-field='saveAsLocation']");
  const note = dialog.querySelector("[data-save-as-note]");
  if (fileNameInput && !options.preserveFileName) fileNameInput.value = fileNameWithExtension(fileNameInput.value || saveAsBaseFileName(), type.extension);
  if (locationInput && !state.dataDialog?.fileHandle && !state.dataDialog?.directoryHandle) {
    locationInput.value = saveAsLocationByValue(state.dataDialog?.location || "downloads").label;
  }
  dialog.querySelectorAll("[data-save-as-file-row]").forEach((row) => {
    row.classList.toggle("is-active", row.dataset.saveAsFileType === type.value);
  });
  if (note) {
    const locationNote =
      state.dataDialog?.fileHandle || state.dataDialog?.directoryHandle
        ? "選択した保存先に直接書き込みます。"
        : state.dataDialog?.location === "downloads"
          ? "保存先を直接選ばない場合はブラウザーのダウンロード先に保存します。"
          : "ブラウザー制限により、参照で保存先を選択しない場合はダウンロード先に保存します。";
    note.textContent =
      (type.value === "html"
        ? "WebSheet の共通ランタイムを利用する小さい HTML として保存します。"
          : type.value === "xls"
            ? "ブラウザー上では旧形式のバイナリ .xls は生成できないため、Excel で開ける HTML 互換 .xls として保存します。"
          : type.value === "xlsm"
            ? hasPreservedVbaProject(state.model)
              ? "現在のブック内容と元ファイルの VBA プロジェクトを保持して保存します。"
              : "現在のブック内容をマクロ有効 Open XML 形式で保存します。VBA プロジェクトが元ファイルに無い場合は新規作成されません。"
            : "現在のブック内容を Excel Open XML 形式で保存します。") + ` ${locationNote}`;
  }
}



async function chooseSaveAsLocation(dialog) {
  if (!dialog) return;
  const type = saveAsFileTypeByValue(dialog.querySelector("[data-data-field='saveAsType']")?.value);
  const suggestedName = fileNameWithExtension(dialog.querySelector("[data-data-field='saveAsFileName']")?.value || saveAsBaseFileName(), type.extension);
  if (!window.showSaveFilePicker) {
    showMessageBox("このブラウザーでは保存先の直接選択に対応していません。保存時はダウンロード フォルダーに保存します。", { title: "Cht WebSheet" });
    setStatus("保存先の直接選択に対応していないため、ダウンロードで保存します。");
    return;
  }
  try {
    const handle = await window.showSaveFilePicker({
      suggestedName,
      types: [{ description: type.description, accept: type.pickerAccept }],
      excludeAcceptAllOption: false,
    });
    state.dataDialog = {
      ...(state.dataDialog || {}),
      type: "save-as",
      location: "this-pc",
      fileHandle: handle,
      directoryHandle: null,
      pickedName: handle.name || suggestedName,
    };
    const location = dialog.querySelector("[data-data-field='saveAsLocation']");
    const fileName = dialog.querySelector("[data-data-field='saveAsFileName']");
    const locationLabel = handle.name ? `選択済み: ${handle.name}` : "選択済み";
    updateSaveAsLocationUi(dialog, "this-pc", locationLabel);
    if (location) location.value = locationLabel;
    if (fileName && handle.name) fileName.value = fileNameWithExtension(handle.name, type.extension);
    syncSaveAsDialogType(dialog, { preserveFileName: true });
  } catch (error) {
    if (error?.name === "AbortError") return;
    console.error(error);
    showMessageBox(`保存先を選択できませんでした：${error.message || error}`, { title: "Cht WebSheet" });
  }
}



async function createSaveAsFolder(dialog) {
  if (!dialog) return;
  if (!window.showDirectoryPicker) {
    showMessageBox("このブラウザーではフォルダーの選択または作成に対応していません。参照からファイル名を指定してください。", { title: "Cht WebSheet" });
    setStatus("フォルダー作成に対応していないため、参照またはダウンロードで保存してください。");
    return;
  }
  try {
    const baseDirectory = state.dataDialog?.directoryHandle || await window.showDirectoryPicker({ mode: "readwrite" });
    let folderName = "新しいフォルダー";
    let suffix = 1;
    let folderHandle = null;
    while (!folderHandle) {
      try {
        folderHandle = await baseDirectory.getDirectoryHandle(folderName, { create: true });
      } catch (error) {
        if (error?.name !== "InvalidModificationError" && error?.name !== "TypeMismatchError") throw error;
        suffix += 1;
        folderName = `新しいフォルダー (${suffix})`;
      }
    }
    state.dataDialog = {
      ...(state.dataDialog || {}),
      type: "save-as",
      location: "this-pc",
      fileHandle: null,
      directoryHandle: folderHandle,
      pickedName: folderHandle.name || folderName,
    };
    updateSaveAsLocationUi(dialog, "this-pc", folderHandle.name || folderName);
    syncSaveAsDialogType(dialog, { preserveFileName: true });
    setStatus(`${folderHandle.name || folderName} を保存先にしました。`);
  } catch (error) {
    if (error?.name === "AbortError") return;
    console.error(error);
    showMessageBox(`フォルダーを作成できませんでした：${error.message || error}`, { title: "Cht WebSheet" });
  }
}



async function applySaveAsDialog(dialog) {
  if (!dialog || !state.model) return false;
  const type = saveAsFileTypeByValue(dialog.querySelector("[data-data-field='saveAsType']")?.value);
  const fileName = fileNameWithExtension(dialog.querySelector("[data-data-field='saveAsFileName']")?.value || saveAsBaseFileName(), type.extension);
  const handle = state.dataDialog?.fileHandle || null;
  const directoryHandle = state.dataDialog?.directoryHandle || null;
  closeDataDialog();
  try {
    const blob = await workbookBlobForSaveType(type, { fileName });
    await saveBlobToUser(blob, fileName, handle, directoryHandle);
    setWorkbookSavedAs(fileName, type.value, { fileHandle: handle, directoryHandle });
    setStatus(`${fileName} を保存しました。`);
    return true;
  } catch (error) {
    console.error(error);
    showMessageBox(`保存できませんでした：${error.message || error}`, { title: "Cht WebSheet" });
    setStatus("保存に失敗しました。");
    return false;
  }
}



async function quickSaveWorkbook() {
  if (!state.model) return false;
  const type = saveAsFileTypeByValue(currentWorkbookSaveType());
  const fileName = fileNameWithExtension(state.model.sourceName || saveAsBaseFileName(), type.extension);
  try {
    let target = await currentSaveTargetForFile(type, fileName);
    if (target.aborted) return false;
    const blob = await workbookBlobForSaveType(type, { fileName });
    try {
      await saveBlobToUser(blob, fileName, target.fileHandle, target.directoryHandle, target.localSave);
    } catch (error) {
      if (!target.localSave?.allowFallback) throw error;
      console.warn("WebSheet local save bridge failed; falling back to browser save.", error);
      setStatus("ローカル保存ブリッジに接続できないため、ブラウザー保存に切り替えます。");
      target = await currentSaveFallbackTargetForFile(type, fileName);
      if (target.aborted) return false;
      await saveBlobToUser(blob, fileName, target.fileHandle, target.directoryHandle, target.localSave);
    }
    const savedName = target.fileHandle?.name || fileName;
    setWorkbookSavedAs(savedName, type.value, target.remember ? target : {});
    setStatus(target.remember ? `${savedName} を上書き保存しました。` : `${savedName} を保存しました。`);
    return true;
  } catch (error) {
    console.error(error);
    showMessageBox(`保存できませんでした：${error.message || error}`, { title: "Cht WebSheet" });
    setStatus("保存に失敗しました。");
    return false;
  }
}



async function currentSaveTargetForFile(type, fileName) {
  const normalizedName = fileNameWithExtension(fileName, type.extension);
  if (state.currentSaveFileHandle?.createWritable && state.currentSaveFileName === normalizedName) {
    return { fileHandle: state.currentSaveFileHandle, directoryHandle: null, remember: true };
  }
  if (state.currentSaveDirectoryHandle?.getFileHandle && state.currentSaveFileName === normalizedName) {
    return { fileHandle: null, directoryHandle: state.currentSaveDirectoryHandle, remember: true };
  }
  const localSave = standaloneLocalSaveTargetForFile(type, normalizedName);
  if (localSave) {
    return { fileHandle: null, directoryHandle: null, localSave, remember: true };
  }
  const persisted = await persistedStandaloneSaveTargetForFile(type, normalizedName);
  if (persisted) return persisted;
  return currentSaveFallbackTargetForFile(type, normalizedName);
}



async function currentSaveFallbackTargetForFile(type, fileName) {
  const normalizedName = fileNameWithExtension(fileName, type.extension);
  const persisted = await persistedStandaloneSaveTargetForFile(type, normalizedName);
  if (persisted) return persisted;
  if (shouldPickStandaloneQuickSaveTarget(type)) {
    return pickStandaloneQuickSaveTarget(type, normalizedName);
  }
  if (shouldRequireStandaloneDirectSave(type)) {
    throw new Error("このブラウザーではローカル HTML を直接保存できません。Chrome または Edge で開き直して保存先の HTML ファイルを選択してください。");
  }
  return { fileHandle: null, directoryHandle: null, remember: false };
}



function shouldPickStandaloneQuickSaveTarget(type) {
  return state.isStandalone && type?.value === "html" && Boolean(window.showSaveFilePicker);
}



function shouldRequireStandaloneDirectSave(type) {
  return state.isStandalone && type?.value === "html" && window.location.protocol === "file:";
}



async function pickStandaloneQuickSaveTarget(type, fileName) {
  try {
    setStatus("保存先の HTML ファイルを選択してください。");
    const pickerOptions = {
      suggestedName: fileNameWithExtension(fileName || standaloneOpenedFileName() || saveAsBaseFileName(), type.extension),
      types: [{ description: type.description, accept: type.pickerAccept }],
      excludeAcceptAllOption: false,
    };
    const startIn = standaloneOpenedFilePickerStartIn();
    if (startIn) pickerOptions.startIn = startIn;
    const handle = await window.showSaveFilePicker(pickerOptions);
    void rememberStandaloneSaveFileHandle(fileNameWithExtension(handle.name || fileName, type.extension), handle);
    return { fileHandle: handle, directoryHandle: null, remember: true };
  } catch (error) {
    if (error?.name === "AbortError") {
      setStatus("保存をキャンセルしました。");
      return { fileHandle: null, directoryHandle: null, remember: false, aborted: true };
    }
    throw error;
  }
}



function setWorkbookSavedAs(fileName, saveType, options = {}) {
  if (!state.model) return;
  state.model.sourceName = sanitizeFileName(fileName) || fileNameWithExtension(saveAsBaseFileName(), saveAsFileTypeByValue(saveType).extension);
  state.model.saveType = normalizeWorkbookSaveType(saveType, state.model.sourceName);
  state.currentSaveType = state.model.saveType;
  if ("fileHandle" in options || "directoryHandle" in options) {
    state.currentSaveFileHandle = options.fileHandle || null;
    state.currentSaveDirectoryHandle = options.directoryHandle || null;
    state.currentSaveFileName = state.currentSaveFileHandle || state.currentSaveDirectoryHandle ? state.model.sourceName : "";
  }
  document.title = state.model.sourceName;
  $workbookName.text(state.model.sourceName);
  markWorkbookSavedClean();
  scheduleWorkbookPersistence();
  if (state.isStandalone && state.model.saveType === "html" && state.currentSaveFileHandle?.createWritable) {
    void rememberStandaloneSaveFileHandle(state.model.sourceName, state.currentSaveFileHandle);
  }
}



function resetCurrentSaveTarget() {
  state.currentSaveFileHandle = null;
  state.currentSaveDirectoryHandle = null;
  state.currentSaveFileName = "";
}



function standaloneOpenedFileName() {
  if (!state.isStandalone) return "";
  try {
    if (window.location.protocol !== "file:") return "";
    const pathname = decodeURIComponent(window.location.pathname || "");
    const name = pathname.split(/[\\/]/).filter(Boolean).pop() || "";
    return sanitizeFileName(name);
  } catch {
    return "";
  }
}



function standaloneOpenedFilePath() {
  if (!state.isStandalone) return "";
  try {
    if (window.location.protocol !== "file:") return "";
    let pathname = decodeURIComponent(window.location.pathname || "");
    if (/^\/[A-Za-z]:\//.test(pathname)) pathname = pathname.slice(1);
    return pathname;
  } catch {
    return "";
  }
}



function standaloneOpenedFilePickerStartIn() {
  const openedPath = standaloneOpenedFilePath().replace(/\\/g, "/").toLowerCase();
  if (!openedPath) return "";
  if (/(^|\/)downloads\//.test(openedPath)) return "downloads";
  if (/(^|\/)documents\//.test(openedPath)) return "documents";
  if (/(^|\/)desktop\//.test(openedPath)) return "desktop";
  if (/(^|\/)pictures\//.test(openedPath)) return "pictures";
  if (/(^|\/)music\//.test(openedPath)) return "music";
  if (/(^|\/)videos\//.test(openedPath)) return "videos";
  return "";
}



function standaloneLocalSaveTargetForFile(type, fileName) {
  if (!state.isStandalone || type?.value !== "html") return null;
  const path = standaloneOpenedFilePath();
  const endpoint = standaloneLocalSaveEndpoint();
  if (!path || !endpoint) return null;
  return {
    endpoint,
    path,
    fileName: fileNameWithExtension(fileName || standaloneOpenedFileName() || saveAsBaseFileName(), type.extension),
    allowFallback: false,
  };
}



function standaloneLocalSaveEndpoint() {
  const configured = window.__WEBSHEET_STANDALONE_CONFIG__?.localSaveEndpoint || window.__WEBSHEET_LOCAL_SAVE_ENDPOINT__;
  if (configured) return String(configured).trim();
  return "";
}



async function persistedStandaloneSaveTargetForFile(type, fileName) {
  if (!state.isStandalone || type?.value !== "html") return null;
  const handle = await loadRememberedStandaloneSaveFileHandle(fileName);
  if (!handle?.createWritable) return null;
  const permission = await requestFileHandleWritePermission(handle);
  if (permission !== "granted") return null;
  state.currentSaveFileHandle = handle;
  state.currentSaveDirectoryHandle = null;
  state.currentSaveFileName = fileNameWithExtension(handle.name || fileName, type.extension);
  return { fileHandle: handle, directoryHandle: null, remember: true };
}



async function queryFileHandleWritePermission(handle) {
  try {
    if (typeof handle.queryPermission !== "function") return "granted";
    return await handle.queryPermission({ mode: "readwrite" });
  } catch {
    return "denied";
  }
}



async function requestFileHandleWritePermission(handle) {
  let permission = await queryFileHandleWritePermission(handle);
  if (permission === "granted") return permission;
  try {
    if (typeof handle.requestPermission === "function") {
      permission = await handle.requestPermission({ mode: "readwrite" });
    }
  } catch {
    permission = "denied";
  }
  return permission;
}



function standaloneSaveHandleKey(fileName) {
  const tokenKey = standaloneSaveHandleTokenKey(standaloneOpenedFileHandleToken());
  if (tokenKey) return tokenKey;
  const openedPath = standaloneOpenedFilePath();
  if (openedPath) return `direct-save-v2:${openedPath}`;
  if (window.location.protocol === "blob:") return "";
  return `direct-save-v2:${fileNameWithExtension(fileName || standaloneOpenedFileName() || saveAsBaseFileName(), "html")}`;
}



async function rememberStandaloneSaveFileHandle(fileName, handle, options = {}) {
  if ((!state.isStandalone && !options.force) || !handle?.createWritable) return;
  try {
    const key = options.key || standaloneSaveHandleKey(fileName);
    if (!key) return;
    const db = await openStandaloneSaveHandleDb();
    await standaloneSaveHandleStoreRequest(db, "readwrite", (store) => store.put({ key, handle }));
    db.close();
  } catch {
    // File handles are an optimization for repeated saves; failing to remember one should not block saving.
  }
}



async function loadRememberedStandaloneSaveFileHandle(fileName) {
  try {
    const key = standaloneSaveHandleKey(fileName);
    if (!key) return null;
    const db = await openStandaloneSaveHandleDb();
    const entry = await standaloneSaveHandleStoreRequest(db, "readonly", (store) => store.get(key));
    db.close();
    return entry?.handle || null;
  } catch {
    return null;
  }
}



function openStandaloneSaveHandleDb() {
  if (!window.indexedDB) return Promise.reject(new Error("IndexedDB is not available."));
  return new Promise((resolve, reject) => {
    const request = window.indexedDB.open("websheet-save-handles", 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains("handles")) db.createObjectStore("handles", { keyPath: "key" });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Cannot open save handle database."));
  });
}



function standaloneSaveHandleStoreRequest(db, mode, callback) {
  return new Promise((resolve, reject) => {
    const transaction = db.transaction("handles", mode);
    const request = callback(transaction.objectStore("handles"));
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Cannot access save handle database."));
    transaction.onerror = () => reject(transaction.error || new Error("Cannot update save handle database."));
  });
}



function standaloneOpenedFileHandleToken() {
  return String(window.__WEBSHEET_OPENED_FILE_HANDLE_TOKEN__ || "").trim();
}



function standaloneSaveHandleTokenKey(token) {
  const value = String(token || "").trim();
  return /^[A-Za-z0-9_-]{8,120}$/.test(value) ? `direct-save-token-v1:${value}` : "";
}



function fileNameWithExtension(name, extension) {
  const cleanName = sanitizeFileName(String(name || "").trim()) || saveAsBaseFileName();
  const cleanExtension = String(extension || "").replace(/^\./, "");
  return cleanName.replace(/\.[^.]*$/, "") + (cleanExtension ? `.${cleanExtension}` : "");
}



async function workbookBlobForSaveType(type, options = {}) {
  if (state.model) saveActiveWindowViewState();
  if (type.value === "html") {
    const standaloneModel = await modelForStandaloneSave(state.model, {
      saveType: "html",
      sourceName: options.fileName ? fileNameWithExtension(options.fileName, "html") : "",
    });
    return new Blob([await buildStandaloneHtml(standaloneModel)], { type: type.mime });
  }
  if (type.value === "xls") {
    return new Blob([buildLegacyExcelHtml(state.model)], { type: `${type.mime};charset=utf-8` });
  }
  const workbook = await buildExcelWorkbookFromModel(state.model);
  const buffer = await postProcessExcelWorkbookBuffer(await workbook.xlsx.writeBuffer(), state.model, type);
  return new Blob([buffer], { type: type.mime });
}



async function saveBlobToUser(blob, fileName, fileHandle = null, directoryHandle = null, localSaveTarget = null) {
  if (localSaveTarget?.endpoint && localSaveTarget?.path) {
    await saveBlobToLocalSaveBridge(blob, localSaveTarget);
    return;
  }
  if (fileHandle?.createWritable) {
    const writable = await fileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    return;
  }
  if (directoryHandle?.getFileHandle) {
    const targetHandle = await directoryHandle.getFileHandle(fileName, { create: true });
    const writable = await targetHandle.createWritable();
    await writable.write(blob);
    await writable.close();
    return;
  }
  downloadBlob(blob, fileName);
}



async function saveBlobToLocalSaveBridge(blob, target) {
  const html = typeof blob.text === "function" ? await blob.text() : "";
  const response = await fetch(target.endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path: target.path, fileName: target.fileName || "", html }),
  });
  if (!response.ok) {
    let message = "";
    try {
      message = (await response.json())?.message || "";
    } catch {
      message = await response.text().catch(() => "");
    }
    throw new Error(message || `Local save bridge failed (${response.status})`);
  }
}



function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}



async function exportStandaloneHtml() {
  if (!state.model) return;
  const type = saveAsFileTypeByValue("html");
  const fileName = `${saveAsBaseFileName()}.html`;
  downloadBlob(await workbookBlobForSaveType(type, { fileName }), fileName);
}



async function modelForStandaloneSave(model, options = {}) {
  if (model === state.model) refreshWorkbookChartsFromSources();
  const snapshot = cloneJsonLikeValue(model);
  if (options.sourceName) snapshot.sourceName = options.sourceName;
  if (options.saveType) snapshot.saveType = options.saveType;
  removeExternalDataBeforeStandaloneSave(snapshot);
  compactStandaloneWorkbookStyleRanges(snapshot);
  refreshStandaloneShapeSources(snapshot);
  await optimizeStandaloneWorkbookImages(snapshot);
  await compressStandaloneWorkbookImageAssets(snapshot);
  return snapshot;
}



function refreshStandaloneShapeSources(model) {
  if (!model?.sheets?.length) return;
  (model.sheets || []).forEach((sheet) => {
    (sheet.images || []).forEach((image) => {
      if (!image?.shape) return;
      const descriptor = shapeDescriptorForType(image.shape.type || DEFAULT_SHAPE_TYPE, {
        ...image.shape,
        width: image.ext?.width,
        height: image.ext?.height,
        includeTextInSvg: false,
      });
      image.src = descriptor.src;
      image.mimeType = descriptor.mimeType;
      image.extension = descriptor.extension;
      image.name = image.name || image.shape.label || descriptor.name || "図形";
      image.shape = { ...image.shape, ...descriptor.shape };
    });
  });
}



function compactStandaloneWorkbookStyleRanges(model) {
  (model?.sheets || []).forEach(compactStandaloneSheetStyleRanges);
  return model;
}



function compactWorkbookStyleOnlyCellsToStyleRanges(model) {
  (model?.sheets || []).forEach((sheet) => {
    compactSheetStyleOnlyCellsToStyleRanges(sheet);
    normalizeSheetStyleRanges(sheet);
  });
  return model;
}



function compactStandaloneSheetStyleRanges(sheet) {
  delete sheet.styleTable;
  compactSheetStyleOnlyCellsToStyleRanges(sheet);
  compactStandaloneSheetStyleTable(sheet);
  return sheet;
}



function compactSheetStyleOnlyCellsToStyleRanges(sheet) {
  if (!sheet?.cells || typeof sheet.cells !== "object") {
    normalizeSheetStyleRanges(sheet);
    return sheet;
  }
  const groups = new Map();
  let removedCells = false;
  Object.entries(sheet.cells).forEach(([key, cell]) => {
    const style = standaloneStyleRangePayloadForCell(cell);
    if (!style) return;
    const row = Number(cell.row);
    const col = Number(cell.col);
    if (!Number.isInteger(row) || !Number.isInteger(col) || row < 1 || col < 1) return;
    const styleKey = stableObjectKey(style);
    let group = groups.get(styleKey);
    if (!group) {
      group = { style, rows: new Map() };
      groups.set(styleKey, group);
    }
    if (!group.rows.has(row)) group.rows.set(row, []);
    group.rows.get(row).push(col);
    delete sheet.cells[key];
    removedCells = true;
  });
  if (removedCells) invalidateSheetUsedRangeCache(sheet);

  const styleRanges = Array.isArray(sheet.styleRanges) ? [...sheet.styleRanges] : [];
  groups.forEach(({ style, rows }) => {
    const openRuns = new Map();
    [...rows.keys()].sort((left, right) => left - right).forEach((row) => {
      const currentRunKeys = new Set();
      standaloneSortedNumberRuns(rows.get(row)).forEach(({ left, right }) => {
        const runKey = `${left}:${right}`;
        currentRunKeys.add(runKey);
        const previous = openRuns.get(runKey);
        if (previous && previous.bottom === row - 1) {
          previous.bottom = row;
        } else {
          const range = { top: row, left, bottom: row, right, ...cloneJsonLikeValue(style) };
          styleRanges.push(range);
          openRuns.set(runKey, range);
        }
      });
      [...openRuns.keys()].forEach((runKey) => {
        if (!currentRunKeys.has(runKey)) openRuns.delete(runKey);
      });
    });
  });
  if (styleRanges.length) sheet.styleRanges = styleRanges;
  else delete sheet.styleRanges;
  normalizeSheetStyleRanges(sheet);
  return sheet;
}



function standaloneStyleRangePayloadForCell(cell) {
  if (!cell || !standaloneCellIsStyleOnly(cell)) return null;
  const css = cloneJsonLikeValue(cell.css || {});
  const borders = cloneJsonLikeValue(cell.borders || {});
  const style = {};
  if (Object.keys(css).length) style.css = css;
  if (cell.hasFill || css.backgroundColor) style.hasFill = true;
  if (cell.hasFontColor) style.hasFontColor = true;
  if (Object.keys(borders).length) style.borders = borders;
  if (!style.css && !style.borders && !style.hasFill && !style.hasFontColor) return null;
  return style;
}



function standaloneCellIsStyleOnly(cell) {
  if (!cell) return false;
  // A style range carries only these fields. Keep number formats and unknown
  // metadata on their cells, including empty cells: future input still depends
  // on that information and recovery must not silently lose it on export.
  const rangeKeys = new Set(['row', 'col', 'css', 'borders', 'hasFill', 'hasFontColor']);
  if (Object.keys(cell).some(key => !rangeKeys.has(key))) return false;
  if (cell.raw != null && cell.raw !== "") return false;
  if (cell.formula || cell.cached != null || cell.display) return false;
  if (cell.hyperlink || cell.comment || cell.note || cell.noteVisible) return false;
  if (cell.validation || cell.sparkline || cell.quickAnalysisBar || cell.quickAnalysisIcon || cell.quickAnalysisStyle) return false;
  if (cell.dataTableFormula || cell.richText?.length || cell.phonetics?.length || cell.showPhonetic) return false;
  if (cell.protection && (cell.protection.locked === false || cell.protection.hidden === true)) return false;
  return true;
}



function standaloneSortedNumberRuns(values = []) {
  const numbers = [...new Set(values.map((value) => Number(value)).filter((value) => Number.isInteger(value) && value > 0))]
    .sort((left, right) => left - right);
  const runs = [];
  numbers.forEach((value) => {
    const last = runs[runs.length - 1];
    if (last && value === last.right + 1) {
      last.right = value;
    } else {
      runs.push({ left: value, right: value });
    }
  });
  return runs;
}



function compactStandaloneSheetStyleTable(sheet) {
  const styleTable = {};
  const styleIds = new Map();
  let nextStyleIndex = 1;
  const styleIdFor = (style) => {
    if (!style) return "";
    const key = stableObjectKey(style);
    let styleId = styleIds.get(key);
    if (!styleId) {
      styleId = `s${nextStyleIndex}`;
      nextStyleIndex += 1;
      styleIds.set(key, styleId);
      styleTable[styleId] = style;
    }
    return styleId;
  };
  const assignStyle = (target) => {
    const style = standaloneStylePayloadForRecord(target);
    const styleId = styleIdFor(style);
    if (!styleId) return;
    target.styleId = styleId;
    delete target.css;
    delete target.borders;
    delete target.hasFill;
    delete target.hasFontColor;
  };
  Object.values(sheet.cells || {}).forEach(assignStyle);
  (sheet.styleRanges || []).forEach(assignStyle);
  if (Object.keys(styleTable).length) sheet.styleTable = styleTable;
}



function standaloneStylePayloadForRecord(record) {
  if (!record || typeof record !== "object") return null;
  const css = cloneJsonLikeValue(record.css || {});
  const borders = cloneJsonLikeValue(record.borders || {});
  const style = {};
  if (Object.keys(css).length) style.css = css;
  if (record.hasFill || css.backgroundColor) style.hasFill = true;
  if (record.hasFontColor) style.hasFontColor = true;
  if (Object.keys(borders).length) style.borders = borders;
  return style.css || style.borders || style.hasFill || style.hasFontColor ? style : null;
}



async function optimizeStandaloneWorkbookImages(model) {
  if (!model?.sheets?.length) return model;
  const imagesByKey = new Map();
  const assets = {};
  let nextIndex = 1;
  for (const sheet of model.sheets || []) {
    if (!Array.isArray(sheet.images) || !sheet.images.length) continue;
    for (const image of sheet.images) {
      if (!standaloneImageCanUseAsset(image)) continue;
      const optimized = await optimizedStandaloneImageSource(image);
      const parsed = parseBase64ImageDataUrl(optimized?.src || image.src);
      if (!parsed) continue;
      const key = `${parsed.mimeType};${parsed.data}`;
      let assetId = imagesByKey.get(key);
      if (!assetId) {
        assetId = `img_${String(nextIndex).padStart(4, "0")}`;
        nextIndex += 1;
        imagesByKey.set(key, assetId);
        assets[assetId] = {
          mimeType: parsed.mimeType,
          extension: fileExtensionFromImageMimeType(parsed.mimeType, image.extension || parsed.extension),
          width: Math.max(1, Math.round(Number(optimized?.width) || Number(image.ext?.width) || 1)),
          height: Math.max(1, Math.round(Number(optimized?.height) || Number(image.ext?.height) || 1)),
          data: parsed.data,
        };
      }
      image.assetId = assetId;
      image.mimeType = assets[assetId].mimeType;
      image.extension = assets[assetId].extension;
      delete image.src;
    }
  }
  if (Object.keys(assets).length) {
    model.assets = {
      ...(model.assets && typeof model.assets === "object" ? model.assets : {}),
      images: assets,
    };
  }
  return model;
}



function standaloneImageCanUseAsset(image) {
  if (!image || image.hidden) return false;
  if (image.shape || image.chart) return false;
  const src = String(image.src || "").trim();
  if (!src || !src.startsWith("data:image/")) return false;
  if (/^data:image\/svg\+xml/i.test(src) || /^data:image\/gif/i.test(src)) return false;
  return Boolean(parseBase64ImageDataUrl(src));
}



async function optimizedStandaloneImageSource(image) {
  const originalSrc = String(image?.src || "");
  const parsed = parseBase64ImageDataUrl(originalSrc);
  if (!parsed) return { src: originalSrc };
  const dimensions = await imageDimensionsFromSource(originalSrc);
  const target = standaloneImageTargetSize(image, dimensions);
  const shouldRasterize = target.width < dimensions.width || target.height < dimensions.height || parsed.mimeType !== "image/webp";
  if (!shouldRasterize) return { src: originalSrc, width: dimensions.width, height: dimensions.height };
  const webp = await rasterizeImageToDataUrl(originalSrc, target.width, target.height, "image/webp", STANDALONE_IMAGE_WEBP_QUALITY);
  if (webp && webp.length < originalSrc.length * 0.9) {
    return { src: webp, width: target.width, height: target.height };
  }
  if (target.width < dimensions.width || target.height < dimensions.height) {
    const fallback = await rasterizeImageToDataUrl(originalSrc, target.width, target.height, parsed.mimeType === "image/jpeg" ? "image/jpeg" : "image/png", 0.9);
    if (fallback && fallback.length < originalSrc.length) return { src: fallback, width: target.width, height: target.height };
  }
  return { src: originalSrc, width: dimensions.width, height: dimensions.height };
}



function standaloneImageTargetSize(image, dimensions) {
  const naturalWidth = Math.max(1, Math.round(Number(dimensions?.width) || Number(image?.ext?.width) || 1));
  const naturalHeight = Math.max(1, Math.round(Number(dimensions?.height) || Number(image?.ext?.height) || 1));
  const displayWidth = Math.max(1, Number(image?.ext?.width) || naturalWidth);
  const displayHeight = Math.max(1, Number(image?.ext?.height) || naturalHeight);
  const maxWidth = Math.max(displayWidth, Math.min(STANDALONE_IMAGE_MAX_DIMENSION, displayWidth * STANDALONE_IMAGE_DEVICE_PIXEL_RATIO));
  const maxHeight = Math.max(displayHeight, Math.min(STANDALONE_IMAGE_MAX_DIMENSION, displayHeight * STANDALONE_IMAGE_DEVICE_PIXEL_RATIO));
  const scale = Math.min(1, STANDALONE_IMAGE_MAX_DIMENSION / naturalWidth, STANDALONE_IMAGE_MAX_DIMENSION / naturalHeight, maxWidth / naturalWidth, maxHeight / naturalHeight);
  return {
    width: Math.max(1, Math.round(naturalWidth * scale)),
    height: Math.max(1, Math.round(naturalHeight * scale)),
  };
}



async function rasterizeImageToDataUrl(src, width, height, mimeType, quality) {
  try {
    const image = await loadImageElement(src);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(Number(width) || image.naturalWidth || image.width || 1));
    canvas.height = Math.max(1, Math.round(Number(height) || image.naturalHeight || image.height || 1));
    const context = canvas.getContext("2d");
    if (!context) return "";
    context.clearRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, mimeType, quality));
    if (!blob) return "";
    const buffer = await blob.arrayBuffer();
    return `data:${blob.type || mimeType};base64,${bytesToBase64(buffer)}`;
  } catch {
    return "";
  }
}



function parseBase64ImageDataUrl(src) {
  const match = String(src || "").match(/^data:(image\/[a-z0-9.+-]+);base64,([\s\S]*)$/i);
  if (!match) return null;
  const mimeType = match[1].toLowerCase();
  const extension = fileExtensionFromImageMimeType(mimeType);
  const data = String(match[2] || "").trim();
  return data ? { mimeType, extension, data } : null;
}



async function standaloneModelScriptPayload(model) {
  return model;
}



async function compressStandaloneWorkbookImageAssets(model) {
  const assets = model?.assets?.images;
  if (!assets || typeof assets !== "object") return model;
  for (const asset of Object.values(assets)) {
    if (!standaloneImageAssetCanUseGzip(asset)) continue;
    let rawBytes = null;
    try {
      rawBytes = base64ToBytes(asset.data);
    } catch {
      continue;
    }
    if (!rawBytes?.length) continue;
    const compressedData = await gzipBase64Bytes(rawBytes);
    if (!compressedData || compressedData.length >= String(asset.data).length * STANDALONE_IMAGE_ASSET_GZIP_MIN_SAVINGS_RATIO) continue;
    asset.data = compressedData;
    asset.dataEncoding = STANDALONE_IMAGE_ASSET_ENCODING_GZIP_BASE64;
    asset.uncompressedByteLength = rawBytes.length;
  }
  return model;
}



function standaloneImageAssetCanUseGzip(asset) {
  if (!asset || typeof asset !== "object") return false;
  if (asset.dataEncoding) return false;
  const data = String(asset.data || "").trim();
  if (data.length < STANDALONE_IMAGE_ASSET_GZIP_MIN_BASE64_LENGTH) return false;
  return /^[A-Za-z0-9+/=\s]+$/.test(data);
}



async function gzipBase64Bytes(bytes) {
  if (typeof CompressionStream !== "function" || typeof Response !== "function" || typeof Blob !== "function") return "";
  try {
    const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream("gzip"));
    const buffer = await new Response(stream).arrayBuffer();
    return bytesToBase64(buffer);
  } catch {
    return "";
  }
}



function removeExternalDataBeforeStandaloneSave(model) {
  if (!model?.sheets?.length) return;
  (model.sheets || []).forEach((sheet) => {
    (sheet.queries || []).forEach((query) => {
      if (!query?.removeDataBeforeSave) return;
      removeQueryCachedDataBeforeSave(sheet, query);
    });
  });
}



function removeQueryCachedDataBeforeSave(sheet, query) {
  if (!sheet || !query) return;
  if (query.range) {
    for (let row = query.range.top; row <= query.range.bottom; row += 1) {
      for (let col = query.range.left; col <= query.range.right; col += 1) {
        delete sheet.cells[cellKey(row, col)];
      }
    }
  }
  if (query.sourceText != null) query.sourceText = "";
  if (Array.isArray(query.sourceRows)) query.sourceRows = [];
  query.savedDataRemoved = true;
  query.savedDataRemovedAt = new Date().toISOString();
}



async function buildExcelWorkbookFromModel(model) {
  await loadExcelJsModule();
  const workbook = new ExcelJS.Workbook();
  exportWorkbookMetadataToExcel(model, workbook);
  const exportOptions = { date1904: workbookUses1904DateSystem(model) };
  const tableNames = new Set();
  for (const [index, sheet] of (model?.sheets || []).entries()) {
    const view = excelWorksheetView(sheet, index);
    const worksheet = workbook.addWorksheet(sanitizeWorksheetName(sheet.name || `Sheet${index + 1}`), {
      state: excelWorksheetState(sheet),
      views: view ? [view] : [],
      properties: excelWorksheetProperties(sheet),
      pageSetup: excelWorksheetPageSetup(sheet),
      headerFooter: excelWorksheetHeaderFooter(sheet),
      autoFilter: excelWorksheetAutoFilter(sheet),
    });
    worksheet.properties.defaultColWidth = pixelsToExcelColumnWidth(sheet.defaultColumnWidth || columnWidthToPixels(8.43));
    exportSheetOutlineToExcel(sheet, worksheet);
    exportSheetDimensionsToExcel(sheet, worksheet);
    exportSheetTablesToExcel(sheet, worksheet, tableNames, exportOptions);
    exportSheetCellsToExcel(sheet, worksheet, exportOptions);
    exportSheetMergesToExcel(sheet, worksheet);
    await exportSheetImagesToExcel(sheet, workbook, worksheet);
    exportSheetProtectionToExcel(sheet, worksheet);
  }
  return workbook;
}



function exportWorkbookMetadataToExcel(model, workbook) {
  if (state.model === model) {
    saveActiveWindowViewState();
    model.activeSheetIndex = state.activeSheetIndex;
  }
  const workbookView = normalizeWorkbookViewModel(model?.workbookView, model?.sheets?.length || 1);
  const activeTab = clampWorkbookSheetIndex(state.model === model ? state.activeSheetIndex : model?.activeSheetIndex ?? workbookView.activeSheetIndex, model?.sheets?.length || 1);
  model.workbookView = { ...workbookView, activeSheetIndex: activeTab };
  workbook.creator = "WebSheet";
  workbook.lastModifiedBy = "WebSheet";
  workbook.created = new Date(model?.createdAt || Date.now());
  workbook.modified = new Date();
  workbook.title = String(model?.sourceName || "").replace(/\.[^.]+$/, "") || "WebSheet";
  workbook.subject = "WebSheet workbook";
  workbook.company = "WebSheet";
  workbook.properties = {
    ...(workbook.properties || {}),
    date1904: model?.workbookProperties?.date1904 === true,
  };
  const calcProperties = normalizeWorkbookCalcProperties(model?.calcProperties);
  workbook.calcProperties = {
    ...(workbook.calcProperties || {}),
    fullCalcOnLoad: calcProperties.fullCalcOnLoad,
    calcMode: calcProperties.calcMode,
  };
  workbook.views = excelWorkbookViewEntries(model, workbookView, activeTab);
}



function excelWorkbookViewEntries(model, workbookView, activeTab) {
  const sheetCount = model?.sheets?.length || 1;
  const main = normalizeWorkbookWindowView({
    id: "main",
    name: model?.sourceName || "Book1",
    x: workbookView.x,
    y: workbookView.y,
    width: workbookView.width,
    height: workbookView.height,
    firstSheet: workbookView.firstSheet,
    sheetIndex: activeTab,
    visibility: workbookView.visibility,
    showHorizontalScroll: workbookView.showHorizontalScroll,
    showVerticalScroll: workbookView.showVerticalScroll,
    showSheetTabs: workbookView.showSheetTabs,
    ...(model?.mainWindowView && typeof model.mainWindowView === "object" ? model.mainWindowView : {}),
    hidden: model?.mainWindowHidden === true,
  });
  const records = [
    main,
    ...((Array.isArray(model?.windowViews) ? model.windowViews : [])
      .map((view) => normalizeWorkbookWindowView(view))
      .filter(Boolean)),
  ];
  const visibleRecords = records.some((record) => !record.hidden)
    ? records
    : records.map((record, index) => index === 0 ? { ...record, hidden: false, visibility: "visible", minimized: false } : record);
  return visibleRecords.map((record, index) => ({
    x: Number.isFinite(Number(record.x)) ? Math.round(Number(record.x)) : workbookView.x + index * 360,
    y: Number.isFinite(Number(record.y)) ? Math.round(Number(record.y)) : workbookView.y + index * 360,
    width: Number.isFinite(Number(record.width)) && Number(record.width) > 0 ? Math.round(Number(record.width)) : workbookView.width,
    height: Number.isFinite(Number(record.height)) && Number(record.height) > 0 ? Math.round(Number(record.height)) : workbookView.height,
    firstSheet: clampWorkbookSheetIndex(record.firstSheet ?? workbookView.firstSheet, sheetCount),
    activeTab: clampWorkbookSheetIndex(record.sheetIndex ?? activeTab, sheetCount),
    visibility: record.hidden ? "hidden" : record.visibility || "visible",
    maximized: record.maximized === true,
    minimized: record.minimized === true,
    showHorizontalScroll: record.showHorizontalScroll !== false,
    showVerticalScroll: record.showVerticalScroll !== false,
    showSheetTabs: record.showSheetTabs !== false,
  }));
}



function excelWorksheetState(sheet = {}) {
  if (sheet.state === "veryHidden") return "veryHidden";
  if (sheet.hidden || sheet.state === "hidden") return "hidden";
  return "visible";
}



function excelWorksheetProperties(sheet = {}) {
  const tabColor = excelColorObjectFromCss(sheet.tabColor);
  return {
    defaultRowHeight: pixelsToExcelRowHeight(sheet.defaultRowHeight || rowHeightToPixels(15)),
    ...(tabColor ? { tabColor } : {}),
    outlineProperties: {
      summaryBelow: sheet.outlineSettings?.summaryBelow !== false,
      summaryRight: sheet.outlineSettings?.summaryRight !== false,
      applyStyles: sheet.outlineSettings?.applyStyles === true,
    },
  };
}



function excelWorksheetView(sheet = {}, sheetIndex = 0) {
  const viewMode = normalizedSheetViewMode(sheet.viewMode);
  const view = {
    zoomScale: Math.round(Number(sheet.zoomScale || 100)),
    showGridLines: sheet.showGridLines !== false,
    showRowColHeaders: sheet.showHeadings !== false,
    showRuler: normalizedSheetShowRuler(sheet.showRuler, viewMode),
    showOutlineSymbols: sheet.showOutlineSymbols !== false,
    showFormulas: sheet.showFormulas === true,
    showZeros: sheet.showZeros !== false,
    showWhiteSpace: sheet.showWhiteSpace !== false,
    showPageBreaks: viewMode === "page-break",
    rightToLeft: sheet.rightToLeft === true,
    tabSelected: excelWorksheetTabSelected(sheet, sheetIndex),
  };
  if (viewMode === "page-layout") view.style = "pageLayout";
  else if (viewMode === "page-break") view.style = "pageBreakPreview";
  const pane = normalizePaneState(sheet.freezePanes);
  if (pane && (pane.rows > 0 || pane.columns > 0)) {
    const rows = Math.max(0, Number(pane.rows) || 0);
    const columns = Math.max(0, Number(pane.columns) || 0);
    view.state = "frozen";
    view.ySplit = rows;
    view.xSplit = columns;
    view.topLeftCell = excelCellAddress(rows + 1, columns + 1);
    view.activePane = pane.activePane || excelSplitActivePane(rows, columns);
  } else {
    const split = normalizePaneState(sheet.splitPane);
    if (split && (split.rows > 0 || split.columns > 0)) {
      view.state = "split";
      view.ySplit = split.rows > 0 ? excelSplitUnitsForPaneIndex(sheet, "row", split.rows) : undefined;
      view.xSplit = split.columns > 0 ? excelSplitUnitsForPaneIndex(sheet, "col", split.columns) : undefined;
      view.topLeftCell = excelCellAddress(split.rows + 1, split.columns + 1);
      view.activePane = split.activePane || excelSplitActivePane(split.rows, split.columns);
    }
  }
  const selection = excelWorksheetViewSelectionState(sheet, sheetIndex);
  if (selection.selected) view.activeCell = excelCellAddress(selection.selected.row, selection.selected.col);
  if (!view.topLeftCell && selection.topLeftCell && (selection.topLeftCell.row > 1 || selection.topLeftCell.col > 1)) {
    view.topLeftCell = excelCellAddress(selection.topLeftCell.row, selection.topLeftCell.col);
  }
  return view;
}



function excelWorksheetTabSelected(sheet = {}, sheetIndex = 0) {
  if (state.model?.sheets?.[sheetIndex] === sheet) {
    const selected = state.selectedSheetIndexes?.length
      ? new Set(state.selectedSheetIndexes)
      : new Set([state.activeSheetIndex]);
    return selected.has(sheetIndex);
  }
  return sheet.tabSelected === true;
}



function excelSplitUnitsForPaneIndex(sheet, axis, count) {
  const maxCount = Math.max(0, Math.round(Number(count) || 0));
  if (!maxCount) return undefined;
  const isColumn = axis === "col";
  const items = isColumn ? sheet.columns || [] : sheet.rows || [];
  const defaultSize = isColumn
    ? sheet.defaultColumnWidth || columnWidthToPixels(8.43)
    : sheet.defaultRowHeight || rowHeightToPixels(15);
  let pixels = 0;
  for (let index = 0; index < maxCount; index += 1) {
    const item = items[index] || {};
    pixels += item.hidden ? 0 : Number(item.width ?? item.height ?? defaultSize) || defaultSize;
  }
  return Math.max(1, Math.round(pixels * 15));
}



function excelSplitActivePane(rows, columns) {
  if (rows > 0 && columns > 0) return "bottomRight";
  if (columns > 0) return "topRight";
  if (rows > 0) return "bottomLeft";
  return undefined;
}



function excelWorksheetPageSetup(sheet = {}) {
  const layout = normalizedPageLayoutSettings(sheet.pageLayout || {}, sheet);
  const fitToPage = layout.fitToPage === true;
  return {
    margins: { ...layout.margins },
    orientation: layout.orientation,
    paperSize: excelPaperSizeCode(layout.paperSize),
    horizontalDpi: 4294967295,
    verticalDpi: 4294967295,
    fitToPage,
    pageOrder: layout.pageOrder,
    blackAndWhite: layout.blackAndWhite,
    draft: layout.draftQuality,
    cellComments: excelPageCommentsMode(layout.comments),
    errors: excelPageErrorsMode(layout.cellErrors),
    scale: fitToPage ? undefined : layout.scale,
    fitToWidth: layout.fitToWidth === "auto" ? 0 : Number(layout.fitToWidth),
    fitToHeight: layout.fitToHeight === "auto" ? 0 : Number(layout.fitToHeight),
    firstPageNumber: layout.firstPageNumber,
    usePrinterDefaults: true,
    showRowColHeaders: layout.printHeadings,
    showGridLines: layout.printGridLines,
    horizontalCentered: layout.centerHorizontal,
    verticalCentered: layout.centerVertical,
    ...(layout.printArea?.range ? { printArea: rangeToLabel(layout.printArea.range) } : {}),
    ...(layout.printTitles?.rows ? { printTitlesRow: `${layout.printTitles.rows.top}:${layout.printTitles.rows.bottom}` } : {}),
    ...(layout.printTitles?.cols ? { printTitlesColumn: `${columnName(layout.printTitles.cols.left)}:${columnName(layout.printTitles.cols.right)}` } : {}),
  };
}



function excelWorksheetHeaderFooter(sheet = {}) {
  const settings = normalizedHeaderFooterSettings(sheet.headerFooter, sheet.pageLayout);
  return {
    differentFirst: Boolean(settings.differentFirst),
    differentOddEven: Boolean(settings.differentOddEven),
    oddHeader: excelHeaderFooterText(settings.oddHeader),
    oddFooter: excelHeaderFooterText(settings.oddFooter),
    evenHeader: excelHeaderFooterText(settings.evenHeader),
    evenFooter: excelHeaderFooterText(settings.evenFooter),
    firstHeader: excelHeaderFooterText(settings.firstHeader),
    firstFooter: excelHeaderFooterText(settings.firstFooter),
    alignWithMargins: settings.alignWithMargins !== false,
    scaleWithDoc: settings.scaleWithDoc !== false,
  };
}



function excelHeaderFooterText(section = {}) {
  const left = escapeExcelHeaderFooterText(section.left);
  const center = escapeExcelHeaderFooterText(section.center);
  const right = escapeExcelHeaderFooterText(section.right);
  return [
    left ? `&L${left}` : "",
    center ? `&C${center}` : "",
    right ? `&R${right}` : "",
  ].join("");
}



function escapeExcelHeaderFooterText(text) {
  return String(text || "").replace(/&/g, "&&");
}



function excelWorksheetAutoFilter(sheet = {}) {
  const range = sheet.autoFilter?.range;
  if (!range) return null;
  return {
    from: { row: Number(range.top), column: Number(range.left) },
    to: { row: Number(range.bottom), column: Number(range.right) },
  };
}



function exportSheetOutlineToExcel(sheet, worksheet) {
  const maxColumnOutline = Math.max(0, ...(sheet.columns || []).map((column) => Number(column.outlineLevel) || 0));
  const maxRowOutline = Math.max(0, ...(sheet.rows || []).map((row) => Number(row.outlineLevel) || 0));
  worksheet.properties.outlineLevelCol = maxColumnOutline;
  worksheet.properties.outlineLevelRow = maxRowOutline;
  worksheet.properties.outlineProperties = {
    summaryBelow: sheet.outlineSettings?.summaryBelow !== false,
    summaryRight: sheet.outlineSettings?.summaryRight !== false,
    applyStyles: sheet.outlineSettings?.applyStyles === true,
  };
  const layout = normalizedPageLayoutSettings(sheet.pageLayout || {}, sheet);
  worksheet.rowBreaks = (layout.pageBreaks?.rows || []).map((row) => ({ id: Math.max(1, Number(row) - 1), max: 16383, man: 1 }));
}



function exportSheetProtectionToExcel(sheet, worksheet) {
  if (!sheet?.protection?.enabled) return;
  const protection = sheet.protection;
  worksheet.sheetProtection = {
    sheet: true,
    selectLockedCells: protection.allowSelectLockedCells !== false,
    selectUnlockedCells: protection.allowSelectUnlockedCells !== false,
    formatCells: Boolean(protection.allowFormatCells),
    formatColumns: Boolean(protection.allowFormatColumns),
    formatRows: Boolean(protection.allowFormatRows),
    insertColumns: Boolean(protection.allowInsertColumns),
    insertRows: Boolean(protection.allowInsertRows),
    deleteColumns: Boolean(protection.allowDeleteColumns),
    deleteRows: Boolean(protection.allowDeleteRows),
    sort: Boolean(protection.allowSort),
    autoFilter: Boolean(protection.allowAutoFilter),
    pivotTables: Boolean(protection.allowPivotTables),
    objects: protection.allowEditObjects ? undefined : false,
    scenarios: protection.allowEditScenarios ? undefined : false,
  };
}

export {
  activeSheet,
  addBorderLine,
  addCellBorderLines,
  applyOpenedWorkbookSaveTarget,
  applySaveAsDialog,
  borderLineHtml,
  borderLinesForSheet,
  buildExcelWorkbookFromModel,
  buildMergeMaps,
  byColLambdaFormulaMatrix,
  byRowLambdaFormulaMatrix,
  canvasBorderDash,
  cellDisplayText,
  chooseColsFormulaMatrix,
  chooseRowsFormulaMatrix,
  chooseSaveAsLocation,
  columnBoundaryPositions,
  combineFormulaFilterIncludeMatrices,
  compactSheetStyleOnlyCellsToStyleRanges,
  compactStandaloneSheetStyleRanges,
  compactStandaloneSheetStyleTable,
  compactStandaloneWorkbookStyleRanges,
  compactWorkbookStyleOnlyCellsToStyleRanges,
  compareBorder,
  compareDateParts,
  compressStandaloneWorkbookImageAssets,
  createSaveAsFolder,
  currentSaveFallbackTargetForFile,
  currentSaveTargetForFile,
  currentWorkbookSaveType,
  dataSummaryMethodForFormulaName,
  dateFormatTokenIsMinute,
  datePartsForDatedif,
  datePartsToUtcDay,
  daysIgnoringMonthsAndYears,
  daysIgnoringYears,
  defaultSaveAsFileType,
  defaultSaveTypeForLoadedFile,
  displayValue,
  downloadBlob,
  dropFormulaMatrix,
  escapeExcelHeaderFooterText,
  evaluateExcelDatedifFormula,
  evaluateExcelSummaryReferenceFormula,
  evaluateExcelTextPartFormula,
  excelCompatibleFormulaValue,
  excelDatedif,
  excelHeaderFooterText,
  excelSplitActivePane,
  excelSplitUnitsForPaneIndex,
  excelWorkbookViewEntries,
  excelWorksheetAutoFilter,
  excelWorksheetHeaderFooter,
  excelWorksheetPageSetup,
  excelWorksheetProperties,
  excelWorksheetState,
  excelWorksheetTabSelected,
  excelWorksheetView,
  expandFormulaMatrix,
  exportSheetOutlineToExcel,
  exportSheetProtectionToExcel,
  exportStandaloneHtml,
  exportWorkbookMetadataToExcel,
  fileNameWithExtension,
  filterFormulaMatrix,
  forceSyncHorizontalScrollbarFromHost,
  formatDateForFormat,
  formatElapsedTimeForFormat,
  formatNumberForFormat,
  formulaComparisonResult,
  formulaDateValue,
  formulaFilterIncludeMatrix,
  formulaIsErrorLikeValue,
  formulaLambdaEvaluate,
  formulaLambdaEvaluateBinary,
  formulaLambdaEvaluateFunction,
  formulaLambdaFlatten,
  formulaLambdaFunctionValues,
  formulaLambdaNumericValues,
  formulaLambdaStripParentheses,
  formulaMatrixValueAt,
  formulaOuterParenthesesWrapWholeExpression,
  formulaSelectorIndex,
  formulaShouldIgnoreArrayValue,
  formulaTruthyValue,
  fullMonthsBetween,
  fullYearsBetween,
  getCellFormula,
  getSheetId,
  gridTrackStyle,
  gzipBase64Bytes,
  handleSaveAsFileRowClick,
  handleSaveAsLocationClick,
  handleSaveAsPathChange,
  handleSaveAsTypeChange,
  heightBetweenRows,
  hstackFormulaMatrices,
  implicitCellDisplayText,
  isBlockingCell,
  isElapsedTimeFormat,
  isOccupiedCell,
  isWorkbookSaveType,
  loadRememberedStandaloneSaveFileHandle,
  makeArrayLambdaFormulaMatrix,
  mapLambdaFormulaMatrix,
  mergeRangeHashStep,
  mergeRangesSignature,
  modelForStandaloneSave,
  monthsIgnoringYears,
  normalizeDateForDatedif,
  normalizeWorkbookSaveType,
  normalizedCellDisplayText,
  numberFormatCanApplyToNumericString,
  numericStringValueForNumberFormat,
  offsetFormulaRange,
  openSaveAsDialog,
  openStandaloneSaveHandleDb,
  optimizeStandaloneWorkbookImages,
  optimizedStandaloneImageSource,
  paintCanvasBorders,
  parseBase64ImageDataUrl,
  parseFormulaCellReference,
  persistedStandaloneSaveTargetForFile,
  pickStandaloneQuickSaveTarget,
  primaryExcelFormatSection,
  queryFileHandleWritePermission,
  quickSaveTypeForCurrentContext,
  quickSaveWorkbook,
  rasterizeImageToDataUrl,
  reduceLambdaFormulaValue,
  refreshStandaloneShapeSources,
  rememberStandaloneSaveFileHandle,
  removeExternalDataBeforeStandaloneSave,
  removeQueryCachedDataBeforeSave,
  renderBorderLayer,
  requestFileHandleWritePermission,
  resetCurrentSaveTarget,
  rowBoundaryPositions,
  saveAsBaseFileName,
  saveAsDialogHtml,
  saveAsFileListRowsHtml,
  saveAsFileTypeByValue,
  saveAsLocationByValue,
  saveBlobToLocalSaveBridge,
  saveBlobToUser,
  saveTypeForSourceName,
  scanLambdaFormulaMatrix,
  sequenceFormulaMatrix,
  setSaveAsLocation,
  setWorkbookSavedAs,
  sheetBorderCellsByRow,
  sheetGridClassName,
  shouldHideZeroValueForSheet,
  shouldPickStandaloneQuickSaveTarget,
  shouldRequireStandaloneDirectSave,
  shouldUseSourceSaveTypeForCurrentDocument,
  sortByFormulaMatrix,
  sortFormulaMatrix,
  splitFormulaArgumentList,
  splitFormulaComparisonExpression,
  standaloneCellIsStyleOnly,
  standaloneImageAssetCanUseGzip,
  standaloneImageCanUseAsset,
  standaloneImageTargetSize,
  standaloneLocalSaveEndpoint,
  standaloneLocalSaveTargetForFile,
  standaloneModelScriptPayload,
  standaloneOpenedFileHandleToken,
  standaloneOpenedFileName,
  standaloneOpenedFilePath,
  standaloneOpenedFilePickerStartIn,
  standaloneSaveHandleKey,
  standaloneSaveHandleStoreRequest,
  standaloneSaveHandleTokenKey,
  standaloneSortedNumberRuns,
  standaloneStylePayloadForRecord,
  standaloneStyleRangePayloadForCell,
  startSheetFooterResize,
  stripFormulaOuterParentheses,
  syncHorizontalScrollbarFromHost,
  syncHostFromHorizontalScrollbar,
  syncSaveAsDialogType,
  takeFormulaMatrix,
  textOverflowWidth,
  tocolFormulaMatrix,
  torowFormulaMatrix,
  transposeFormulaMatrix,
  trimNumber,
  uniqueFormulaMatrix,
  updateHorizontalScrollbar,
  updateSaveAsLocationUi,
  visibleColumnRange,
  visibleColumnRangeRightToLeft,
  visibleGridRange,
  visibleRowRange,
  vstackFormulaMatrices,
  widthBetweenColumns,
  workbookBlobForSaveType,
  wrapColsFormulaMatrix,
  wrapRowsFormulaMatrix,
};
