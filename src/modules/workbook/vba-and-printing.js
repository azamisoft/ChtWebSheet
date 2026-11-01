// Feature module; shared bindings are imported explicitly.
import {
  $,
  $pageBackgroundInput,
  CELL_BORDER_SIDES,
  COLUMN_HEADER_HEIGHT,
  DEFAULT_MACRO_SECURITY_LEVEL,
  EXCEL_DEFAULT_PAGE_MARGINS,
  MACRO_SECURITY_NAV_ITEMS,
  MACRO_SECURITY_OPTIONS,
  PAGE_LAYOUT_EFFECT_PRESETS,
  PAGE_LAYOUT_MARGIN_PRESETS,
  PAGE_LAYOUT_PAPER_SIZES,
  PAGE_LAYOUT_THEMES,
  state,
} from "../../app-core.js";
import {
  sheetUsedRange,
} from "../data/sort-and-autofilter.js";
import {
  closeDataDialog,
} from "../drawing/charts-and-insertion.js";
import {
  fileExtension,
  fileToDataUrl,
  imageDimensionsFromSource,
  isLineShapeObject,
  isLineShapeType,
  normalizeShapeGlow,
  normalizeShapeReflection,
  normalizeShapeShadow,
  normalizeShapeSoftEdges,
} from "../drawing/shape-rendering.js";
import {
  refreshShapeSource,
  updateConnectedLinesForTargets,
} from "../drawing/transforms-and-connectors.js";
import {
  normalizeColorValue,
  normalizedThemeColorPalette,
} from "../editing/ink-and-clipboard.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  cssEscape,
  expandRangeForMerges,
  getDisplayForCell,
  materializeSelectionRange,
  materializedRangesForSheet,
  rangeToLabel,
  selectableCellPoint,
  selectionRangeFromPoints,
  sheetIndexFromPixel,
  syncShapeFormatControls,
  unionMaterializedRanges,
  withDisplayCalculationCache,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  columnSelectionRange,
  imageBoxToPixels,
  roundCssNumber,
  rowSelectionRange,
  selectedImageObjects,
  selectionToolbarPoint,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  cellEffectiveCss,
  cloneJsonLikeValue,
  pushHistory,
  storeCellModelForSheet,
} from "../editing/structure-and-history.js";
import {
  webSheetFormulaWildcardRegex,
} from "../formula/calculations.js";
import {
  ensureCellModel,
  forEachCellInRange,
  pageBreakPreviewPageNumber,
  pageBreakPreviewPlan,
  rangesIntersect,
} from "../rendering/grid-and-canvas.js";
import {
  accessKeyLabelHtml,
  cellKey,
  columnName,
  columnWidthToPixels,
  escapeAttr,
  escapeHtml,
  excelSerialToDate,
  rowHeightToPixels,
  setStatus,
  showMessageBox,
  styleObjectToString,
} from "../ui/help-and-utilities.js";
import {
  closeHomeCommandMenu,
  homeCommandActionButtonHtml,
  openHomeCommandMenu,
  printSheetObjectsHtml,
  syncPrintObjectLayerPositions,
} from "../ui/printing-and-home-tools.js";
import {
  ensureSheetSize,
  rowHeaderWidthForSheet,
} from "../ui/ribbon-and-links.js";
import {
  ensureSheetProtectionPermission,
} from "../ui/sheet-protection-and-translation.js";
import {
  captureWorkbookMacroState,
  captureWorkbookViewState,
  commitWorkbookMacroChange,
  refreshActiveSheetViewSnapshot,
  sheetHeadingsVisible,
} from "../ui/views-and-formula-tools.js";
import {
  legacyExcelCellStyle,
} from "./export-model-and-measurement.js";
import {
  renderWorkbook,
  rgbToHex,
  visibleSheetIndexes,
} from "./formulas-and-sheet-tabs.js";
import {
  captureWorkbookThemeState,
  normalizeMacroSecurity,
  normalizeSheetCodeModule,
  scheduleWorkbookPersistence,
  sheetCodeModuleName,
} from "./import-model.js";
import {
  activeSheet,
  cellDisplayText,
  getSheetId,
  textOverflowWidth,
  widthBetweenColumns,
} from "./saving-and-export.js";
import {
  imageMimeType,
} from "./tables-and-validation.js";
import {
  defaultMacroName,
  evaluateVbaExpression,
  executeWorkbookMacroStatements,
  macroNameFromDialog,
  openMacrosDialog,
  pageBreakPositionLabel,
  parseVbaMacroStatements,
  parseVbaRangeTarget,
  rangeCellCount,
  saveActiveWindowViewState,
  selectedMacroIndexFromDialog,
  showUnsupportedMacroRunMessage,
  splitVbaTopLevelArguments,
  validateMacroName,
  vbaVariableKey,
} from "./views-and-macros.js";


function evaluateVbaBuiltInFunctionExpression(text, context) {
  const match = String(text || "").match(/^([A-Za-z_][A-Za-z0-9_]*)\(([\s\S]*)\)$/);
  if (!match) return undefined;
  const rawName = match[1];
  const name = rawName.toUpperCase();
  const args = splitVbaTopLevelArguments(match[2] || "").filter((part) => part !== "");
  const variable = context?.variables?.[vbaVariableKey(rawName)];
  if (variable?.__webSheetVbaArray) {
    const index = Math.round(Number(evaluateVbaExpression(args[0], context)));
    const lower = Number.isFinite(Number(variable.lowerBound)) ? Number(variable.lowerBound) : 0;
    return variable.values[index - lower];
  }
  if (name === "LBOUND" || name === "UBOUND") {
    const value = evaluateVbaExpression(args[0], context);
    if (value?.__webSheetVbaArray) {
      const lower = Number.isFinite(Number(value.lowerBound)) ? Number(value.lowerBound) : 0;
      return name === "LBOUND" ? lower : lower + Math.max(0, value.values.length - 1);
    }
    return 0;
  }
  const values = args.map((arg) => evaluateVbaExpression(arg, context));
  if (name === "IIF") return vbaTruthy(values[0]) ? values[1] : values[2];
  if (name === "LEN") return String(values[0] ?? "").length;
  if (name === "LEFT") return String(values[0] ?? "").slice(0, Math.max(0, Math.round(Number(values[1]) || 1)));
  if (name === "RIGHT") {
    const source = String(values[0] ?? "");
    const count = Math.max(0, Math.round(Number(values[1]) || 1));
    return source.slice(Math.max(0, source.length - count));
  }
  if (name === "MID") {
    const source = String(values[0] ?? "");
    const start = Math.max(1, Math.round(Number(values[1]) || 1));
    const count = Math.max(0, Math.round(Number(values[2]) || source.length));
    return source.slice(start - 1, start - 1 + count);
  }
  if (name === "TRIM") return String(values[0] ?? "").trim().replace(/\s+/g, " ");
  if (name === "UCASE") return String(values[0] ?? "").toLocaleUpperCase("ja-JP");
  if (name === "LCASE") return String(values[0] ?? "").toLocaleLowerCase("ja-JP");
  if (name === "REPLACE") {
    const source = String(values[0] ?? "");
    const find = String(values[1] ?? "");
    const replacement = String(values[2] ?? "");
    const start = Math.max(1, Math.round(Number(values[3]) || 1));
    const count = values[4] == null ? Number.POSITIVE_INFINITY : Math.max(0, Math.round(Number(values[4]) || 0));
    if (!find) return source;
    const prefix = source.slice(0, start - 1);
    let replaced = "";
    let rest = source.slice(start - 1);
    let done = 0;
    while (rest && done < count) {
      const index = rest.indexOf(find);
      if (index < 0) break;
      replaced += rest.slice(0, index) + replacement;
      rest = rest.slice(index + find.length);
      done += 1;
    }
    return `${prefix}${replaced}${rest}`;
  }
  if (name === "INSTR") {
    const offset = values.length >= 3 ? Math.max(1, Math.round(Number(values[0]) || 1)) : 1;
    const haystack = String(values.length >= 3 ? values[1] : values[0] ?? "");
    const needle = String(values.length >= 3 ? values[2] : values[1] ?? "");
    if (!needle) return offset;
    const index = haystack.indexOf(needle, offset - 1);
    return index < 0 ? 0 : index + 1;
  }
  if (["CSTR", "STR"].includes(name)) return String(values[0] ?? "");
  if (["CDbl", "CSNG", "CURRENCY", "VAL"].includes(rawName) || ["CDBL", "CSNG", "CCUR", "VAL"].includes(name)) {
    const matchNumber = String(values[0] ?? "").trim().match(/^[+-]?\d+(?:\.\d+)?/);
    const number = Number(matchNumber ? matchNumber[0] : values[0]);
    return Number.isFinite(number) ? number : 0;
  }
  if (["CLNG", "CINT"].includes(name)) return Math.round(Number(values[0]) || 0);
  if (name === "ABS") return Math.abs(Number(values[0]) || 0);
  if (name === "INT") return Math.floor(Number(values[0]) || 0);
  if (name === "ROUND") return vbaRoundNumber(values[0], values[1], "nearest");
  if (name === "ISNUMERIC") return Number.isFinite(Number(values[0]));
  if (name === "ISDATE") return !Number.isNaN(vbaDateFromValue(values[0]).getTime());
  if (name === "DATE") return new Date();
  if (name === "NOW") return new Date();
  if (name === "DATESERIAL") return new Date(Number(values[0]) || 0, Math.max(0, (Number(values[1]) || 1) - 1), Number(values[2]) || 1);
  if (name === "TIMESERIAL") return ((Number(values[0]) || 0) * 3600 + (Number(values[1]) || 0) * 60 + (Number(values[2]) || 0)) / 86400;
  if (["YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND"].includes(name)) {
    const date = vbaDateFromValue(values[0]);
    if (Number.isNaN(date.getTime())) return 0;
    if (name === "YEAR") return date.getFullYear();
    if (name === "MONTH") return date.getMonth() + 1;
    if (name === "DAY") return date.getDate();
    if (name === "HOUR") return date.getHours();
    if (name === "MINUTE") return date.getMinutes();
    return date.getSeconds();
  }
  if (name === "CHR") return String.fromCharCode(Math.max(0, Math.round(Number(values[0]) || 0)));
  if (name === "ASC") return String(values[0] ?? "").charCodeAt(0) || 0;
  return undefined;
}



function vbaDateFromValue(value) {
  if (value instanceof Date) return value;
  if (typeof value === "number" && Number.isFinite(value)) return excelSerialToDate(value);
  const date = new Date(String(value ?? ""));
  return date;
}



function splitVbaTopLevelWord(text, word) {
  const source = String(text || "");
  const parts = [];
  const target = String(word || "").toLowerCase();
  let start = 0;
  let depth = 0;
  let inString = false;
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char === '"') {
      if (inString && source[index + 1] === '"') index += 1;
      else inString = !inString;
      continue;
    }
    if (inString) continue;
    if (char === "(") {
      depth += 1;
      continue;
    }
    if (char === ")") {
      depth = Math.max(0, depth - 1);
      continue;
    }
    if (
      depth === 0 &&
      source.slice(index, index + target.length).toLowerCase() === target &&
      isVbaWordBoundary(source[index - 1]) &&
      isVbaWordBoundary(source[index + target.length])
    ) {
      parts.push(source.slice(start, index).trim());
      start = index + target.length;
      index += target.length - 1;
    }
  }
  if (!parts.length) return [source.trim()];
  parts.push(source.slice(start).trim());
  return parts.filter((part) => part !== "");
}



function isVbaWordBoundary(char) {
  return char == null || !/[A-Za-z0-9_]/.test(char);
}



function splitVbaComparisonExpression(text) {
  const source = String(text || "");
  const operators = ["<>", ">=", "<=", "=", ">", "<"];
  let depth = 0;
  let inString = false;
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char === '"') {
      if (inString && source[index + 1] === '"') index += 1;
      else inString = !inString;
      continue;
    }
    if (inString) continue;
    if (char === "(") {
      depth += 1;
      continue;
    }
    if (char === ")") {
      depth = Math.max(0, depth - 1);
      continue;
    }
    if (depth !== 0) continue;
    const operator = operators.find((candidate) => source.startsWith(candidate, index));
    if (!operator) continue;
    const left = source.slice(0, index).trim();
    const right = source.slice(index + operator.length).trim();
    if (!left || !right) continue;
    return { left, operator, right };
  }
  return null;
}



function compareVbaExpressionValues(left, right, operator) {
  const leftComparable = vbaComparableValue(left);
  const rightComparable = vbaComparableValue(right);
  switch (operator) {
    case "=":
      return leftComparable === rightComparable;
    case "<>":
      return leftComparable !== rightComparable;
    case ">":
      return leftComparable > rightComparable;
    case ">=":
      return leftComparable >= rightComparable;
    case "<":
      return leftComparable < rightComparable;
    case "<=":
      return leftComparable <= rightComparable;
    default:
      return false;
  }
}



function vbaComparableValue(value) {
  if (value instanceof Date) return value.getTime();
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const numeric = Number(value);
  if (String(value ?? "").trim() !== "" && Number.isFinite(numeric)) return numeric;
  return String(value ?? "");
}



function trimVbaOuterParentheses(text) {
  let source = String(text || "").trim();
  while (source.startsWith("(") && source.endsWith(")") && vbaOuterParenthesesWrapWholeExpression(source)) {
    source = source.slice(1, -1).trim();
  }
  return source;
}



function vbaOuterParenthesesWrapWholeExpression(source) {
  let depth = 0;
  let inString = false;
  const text = String(source || "");
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === '"') {
      if (inString && text[index + 1] === '"') index += 1;
      else inString = !inString;
      continue;
    }
    if (inString) continue;
    if (char === "(") depth += 1;
    else if (char === ")") {
      depth -= 1;
      if (depth === 0 && index < text.length - 1) return false;
      if (depth < 0) return false;
    }
  }
  return depth === 0;
}



function splitVbaArithmeticExpression(text, operators) {
  const source = String(text || "").trim();
  let inString = false;
  let depth = 0;
  let splitIndex = -1;
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char === '"') {
      if (inString && source[index + 1] === '"') index += 1;
      else inString = !inString;
      continue;
    }
    if (inString) continue;
    if (char === "(") {
      depth += 1;
      continue;
    }
    if (char === ")") {
      depth = Math.max(0, depth - 1);
      continue;
    }
    if (depth === 0 && operators.includes(char) && !isVbaUnaryArithmeticOperator(source, index)) {
      splitIndex = index;
    }
  }
  if (splitIndex < 0) return null;
  return {
    left: source.slice(0, splitIndex).trim(),
    operator: source[splitIndex],
    right: source.slice(splitIndex + 1).trim(),
  };
}



function isVbaUnaryArithmeticOperator(source, index) {
  const char = source[index];
  if (char !== "+" && char !== "-") return false;
  for (let previous = index - 1; previous >= 0; previous -= 1) {
    const prev = source[previous];
    if (/\s/.test(prev)) continue;
    return /[=(,+\-*/&]/.test(prev);
  }
  return true;
}



function evaluateVbaWorksheetFunctionExpression(text, context) {
  const match = String(text || "").match(/^(?:(?:Application\.)?WorksheetFunction\.|Application\.)([A-Za-z][A-Za-z0-9_.]*)\(([\s\S]*)\)$/i);
  if (!match) return undefined;
  const name = String(match[1] || "").toUpperCase();
  const args = splitVbaTopLevelArguments(match[2] || "").map((part) => part.trim()).filter((part) => part !== "");
  if (name === "COUNTIF") {
    const target = parseVbaRangeTarget(args[0], context);
    if (!target) return undefined;
    const criteria = evaluateVbaExpression(args[1], context);
    return vbaRangeValues(target).filter((value) => vbaCriteriaMatches(value, criteria)).length;
  }
  if (name === "COUNTIFS") {
    const criteriaPairs = vbaWorksheetCriteriaPairs(args, 0, context);
    if (!criteriaPairs) return undefined;
    return vbaCriteriaRowIndexes(criteriaPairs).length;
  }
  if (name === "SUMIF") {
    const criteriaTarget = parseVbaRangeTarget(args[0], context);
    if (!criteriaTarget) return undefined;
    const criteria = evaluateVbaExpression(args[1], context);
    const sumTarget = args[2] ? parseVbaRangeTarget(args[2], context) : criteriaTarget;
    const criteriaValues = vbaRangeValues(criteriaTarget);
    const sumValues = sumTarget ? vbaRangeValues(sumTarget) : criteriaValues;
    return criteriaValues.reduce((total, value, index) => {
      if (!vbaCriteriaMatches(value, criteria)) return total;
      const numeric = Number(sumValues[index]);
      return Number.isFinite(numeric) ? total + numeric : total;
    }, 0);
  }
  if (name === "SUMIFS") {
    const sumTarget = parseVbaRangeTarget(args[0], context);
    const criteriaPairs = vbaWorksheetCriteriaPairs(args, 1, context);
    if (!sumTarget || !criteriaPairs) return undefined;
    const sumValues = vbaRangeValues(sumTarget);
    return vbaCriteriaRowIndexes(criteriaPairs).reduce((total, index) => {
      const numeric = Number(sumValues[index]);
      return Number.isFinite(numeric) ? total + numeric : total;
    }, 0);
  }
  if (name === "AVERAGEIF") {
    const criteriaTarget = parseVbaRangeTarget(args[0], context);
    if (!criteriaTarget) return undefined;
    const criteria = evaluateVbaExpression(args[1], context);
    const averageTarget = args[2] ? parseVbaRangeTarget(args[2], context) : criteriaTarget;
    const criteriaValues = vbaRangeValues(criteriaTarget);
    const averageValues = averageTarget ? vbaRangeValues(averageTarget) : criteriaValues;
    const matches = criteriaValues
      .map((value, index) => (vbaCriteriaMatches(value, criteria) ? Number(averageValues[index]) : Number.NaN))
      .filter((value) => Number.isFinite(value));
    return matches.length ? matches.reduce((total, value) => total + value, 0) / matches.length : 0;
  }
  if (name === "AVERAGEIFS") {
    const averageTarget = parseVbaRangeTarget(args[0], context);
    const criteriaPairs = vbaWorksheetCriteriaPairs(args, 1, context);
    if (!averageTarget || !criteriaPairs) return undefined;
    const averageValues = vbaRangeValues(averageTarget);
    const matches = vbaCriteriaRowIndexes(criteriaPairs)
      .map((index) => Number(averageValues[index]))
      .filter((value) => Number.isFinite(value));
    return matches.length ? matches.reduce((total, value) => total + value, 0) / matches.length : 0;
  }
  if (name === "VLOOKUP" || name === "HLOOKUP") {
    const table = parseVbaRangeTarget(args[1], context);
    if (!table) return undefined;
    const lookupValue = evaluateVbaExpression(args[0], context);
    const index = Math.max(1, Math.round(Number(evaluateVbaExpression(args[2], context)) || 1));
    const approximate = args[3] == null || vbaTruthy(evaluateVbaExpression(args[3], context));
    return name === "VLOOKUP"
      ? vbaLookupVertical(lookupValue, vbaRangeMatrix(table), index, approximate)
      : vbaLookupHorizontal(lookupValue, vbaRangeMatrix(table), index, approximate);
  }
  if (name === "MATCH") {
    const lookupTarget = parseVbaRangeTarget(args[1], context);
    const lookupValues = lookupTarget ? vbaRangeValues(lookupTarget) : vbaExpressionListValue(args[1], context);
    return vbaMatchIndex(evaluateVbaExpression(args[0], context), lookupValues, args[2] == null ? 1 : Number(evaluateVbaExpression(args[2], context)));
  }
  if (name === "INDEX") {
    const target = parseVbaRangeTarget(args[0], context);
    const matrix = target ? vbaRangeMatrix(target) : vbaArrayToMatrix(evaluateVbaExpression(args[0], context));
    return vbaIndexValue(matrix, Number(evaluateVbaExpression(args[1] || "1", context)), Number(evaluateVbaExpression(args[2] || "1", context)));
  }
  const values = args.flatMap((arg) => {
    const target = parseVbaRangeTarget(arg, context);
    if (target) return vbaRangeValues(target);
    const value = evaluateVbaExpression(arg, context);
    return value?.__webSheetVbaArray ? value.values : [value];
  });
  const numericValues = values.map((value) => Number(value)).filter((value) => Number.isFinite(value));
  if (name === "SUM") return numericValues.reduce((total, value) => total + value, 0);
  if (name === "AVERAGE") return numericValues.length ? numericValues.reduce((total, value) => total + value, 0) / numericValues.length : 0;
  if (name === "MIN") return numericValues.length ? Math.min(...numericValues) : 0;
  if (name === "MAX") return numericValues.length ? Math.max(...numericValues) : 0;
  if (name === "COUNT") return numericValues.length;
  if (name === "COUNTA") return values.filter((value) => value != null && value !== "").length;
  if (name === "COUNTBLANK") return values.filter((value) => value == null || value === "").length;
  if (name === "PRODUCT") return numericValues.reduce((total, value) => total * value, numericValues.length ? 1 : 0);
  if (name === "MEDIAN") {
    if (!numericValues.length) return 0;
    const sorted = [...numericValues].sort((left, right) => left - right);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  }
  if (name === "ROUND") return vbaRoundNumber(values[0], values[1], "nearest");
  if (name === "ROUNDUP") return vbaRoundNumber(values[0], values[1], "up");
  if (name === "ROUNDDOWN") return vbaRoundNumber(values[0], values[1], "down");
  if (name === "ABS") return Math.abs(Number(values[0]) || 0);
  if (name === "INT") return Math.floor(Number(values[0]) || 0);
  if (name === "LEN") return String(values[0] ?? "").length;
  if (name === "LEFT") return String(values[0] ?? "").slice(0, Math.max(0, Math.round(Number(values[1]) || 1)));
  if (name === "RIGHT") {
    const source = String(values[0] ?? "");
    const count = Math.max(0, Math.round(Number(values[1]) || 1));
    return source.slice(Math.max(0, source.length - count));
  }
  if (name === "MID") {
    const source = String(values[0] ?? "");
    const start = Math.max(1, Math.round(Number(values[1]) || 1));
    const count = Math.max(0, Math.round(Number(values[2]) || source.length));
    return source.slice(start - 1, start - 1 + count);
  }
  if (name === "TRIM") return String(values[0] ?? "").trim().replace(/\s+/g, " ");
  if (name === "UPPER") return String(values[0] ?? "").toLocaleUpperCase("ja-JP");
  if (name === "LOWER") return String(values[0] ?? "").toLocaleLowerCase("ja-JP");
  return undefined;
}



function vbaWorksheetCriteriaPairs(args, startIndex, context) {
  const pairs = [];
  for (let index = startIndex; index < args.length; index += 2) {
    const target = parseVbaRangeTarget(args[index], context);
    if (!target || args[index + 1] == null) return null;
    pairs.push({ values: vbaRangeValues(target), criteria: evaluateVbaExpression(args[index + 1], context) });
  }
  return pairs.length ? pairs : null;
}



function vbaCriteriaRowIndexes(pairs) {
  if (!pairs?.length) return [];
  const length = Math.min(...pairs.map((pair) => pair.values.length));
  const indexes = [];
  for (let index = 0; index < length; index += 1) {
    if (pairs.every((pair) => vbaCriteriaMatches(pair.values[index], pair.criteria))) indexes.push(index);
  }
  return indexes;
}



function vbaRangeMatrix(target) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return [];
  const rows = [];
  for (let row = target.range.top; row <= target.range.bottom; row += 1) {
    const values = [];
    for (let col = target.range.left; col <= target.range.right; col += 1) {
      const cell = sheet.cells?.[cellKey(row, col)];
      values.push(cell?.raw ?? cell?.cached ?? cellDisplayText(cell));
    }
    rows.push(values);
  }
  return rows;
}



function vbaArrayToMatrix(value) {
  if (value?.__webSheetVbaArray) return [value.values];
  if (Array.isArray(value)) return Array.isArray(value[0]) ? value : [value];
  return [[value]];
}



function vbaExpressionListValue(expression, context) {
  const value = evaluateVbaExpression(expression, context);
  if (value?.__webSheetVbaArray) return value.values;
  if (Array.isArray(value)) return value.flat();
  return [value];
}



function vbaLookupVertical(lookupValue, matrix, columnIndex, approximate) {
  if (!matrix?.length || columnIndex < 1) return null;
  let fallback = null;
  for (const row of matrix) {
    const key = row?.[0];
    if (compareVbaExpressionValues(key, lookupValue, "=")) return row[columnIndex - 1] ?? null;
    if (approximate && compareVbaExpressionValues(key, lookupValue, "<=")) fallback = row[columnIndex - 1] ?? null;
  }
  return approximate ? fallback : null;
}



function vbaLookupHorizontal(lookupValue, matrix, rowIndex, approximate) {
  if (!matrix?.length || rowIndex < 1) return null;
  const header = matrix[0] || [];
  let fallbackColumn = -1;
  for (let col = 0; col < header.length; col += 1) {
    const key = header[col];
    if (compareVbaExpressionValues(key, lookupValue, "=")) return matrix[rowIndex - 1]?.[col] ?? null;
    if (approximate && compareVbaExpressionValues(key, lookupValue, "<=")) fallbackColumn = col;
  }
  return approximate && fallbackColumn >= 0 ? matrix[rowIndex - 1]?.[fallbackColumn] ?? null : null;
}



function vbaMatchIndex(lookupValue, values, matchType = 1) {
  const list = Array.isArray(values) ? values : [values];
  const mode = Number.isFinite(Number(matchType)) ? Number(matchType) : 1;
  if (mode === 0) {
    const index = list.findIndex((value) => vbaCriteriaMatches(value, lookupValue));
    return index >= 0 ? index + 1 : 0;
  }
  let bestIndex = -1;
  for (let index = 0; index < list.length; index += 1) {
    const operator = mode < 0 ? ">=" : "<=";
    if (compareVbaExpressionValues(list[index], lookupValue, operator)) bestIndex = index;
  }
  return bestIndex >= 0 ? bestIndex + 1 : 0;
}



function vbaIndexValue(matrix, rowNumber, columnNumberValue) {
  const rows = Array.isArray(matrix) ? matrix : [];
  if (!rows.length) return null;
  const row = Math.max(1, Math.round(Number(rowNumber) || 1));
  const col = Math.max(1, Math.round(Number(columnNumberValue) || 1));
  return rows[row - 1]?.[col - 1] ?? null;
}



function vbaRoundNumber(value, digits, mode = "nearest") {
  const numeric = Number(value) || 0;
  const places = Math.max(-15, Math.min(15, Math.round(Number(digits) || 0)));
  const factor = 10 ** Math.abs(places);
  const scaled = places >= 0 ? numeric * factor : numeric / factor;
  const rounded = mode === "up"
    ? Math.sign(scaled || 1) * Math.ceil(Math.abs(scaled))
    : mode === "down"
      ? Math.sign(scaled || 1) * Math.floor(Math.abs(scaled))
      : Math.round(scaled);
  return places >= 0 ? rounded / factor : rounded * factor;
}



function vbaCriteriaMatches(value, criteria) {
  const criteriaText = String(criteria ?? "").trim();
  const match = criteriaText.match(/^(<=|>=|<>|=|<|>)([\s\S]*)$/);
  if (match) {
    const target = match[2].trim();
    if ((match[1] === "=" || match[1] === "<>") && /[*?]/.test(target)) {
      const matched = webSheetFormulaWildcardRegex(target.replace(/~([*?])/gu, "$1")).test(String(value ?? ""));
      return match[1] === "<>" ? !matched : matched;
    }
    return compareVbaExpressionValues(value, target, match[1]);
  }
  if (/[*?]/.test(criteriaText)) return webSheetFormulaWildcardRegex(criteriaText.replace(/~([*?])/gu, "$1")).test(String(value ?? ""));
  return compareVbaExpressionValues(value, criteria, "=");
}



function vbaRangeValues(target) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return [];
  const values = [];
  forEachCellInRange(target.range, (row, col) => {
    const cell = sheet.cells?.[cellKey(row, col)];
    values.push(cell?.raw ?? cell?.cached ?? cellDisplayText(cell));
  });
  return values;
}



function readVbaRangeValue(target) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return null;
  const property = String(target.property || "Value").toLowerCase();
  if (property === "row") return target.range.top;
  if (property === "column") return target.range.left;
  if (property === "count" || property === "cells.count") return rangeCellCount(target.range);
  if (property === "rows.count") return Math.max(1, target.range.bottom - target.range.top + 1);
  if (property === "columns.count") return Math.max(1, target.range.right - target.range.left + 1);
  if (property === "address" || property === "addresslocal") return vbaRangeAddress(target.range);
  if (property === "formula" || property === "formula2" || property === "formulalocal" || property === "formula2local") {
    const formulaCell = sheet.cells?.[cellKey(target.range.top, target.range.left)];
    return formulaCell?.formula || "";
  }
  const cell = sheet.cells?.[cellKey(target.range.top, target.range.left)];
  return cell?.raw ?? cell?.cached ?? cellDisplayText(cell);
}



function vbaRangeAddress(range) {
  const start = `$${columnName(range.left)}$${range.top}`;
  const end = `$${columnName(range.right)}$${range.bottom}`;
  return start === end ? start : `${start}:${end}`;
}



function clearVbaRangeTarget(target, method, context) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return;
  ensureSheetSize(sheet, target.range.bottom, target.range.right);
  let changed = 0;
  forEachCellInRange(target.range, (row, col) => {
    const key = cellKey(row, col);
    const cell = sheet.cells?.[key];
    const action = String(method || "").toLowerCase();
    if (action === "clear") {
      if (cell) {
        delete sheet.cells[key];
        changed += 1;
      }
      return;
    }
    const model = cell || ensureCellModel(sheet, row, col);
    if (action === "clearformats") {
      model.css = {};
      model.borders = {};
      model.hasFill = false;
      model.hasFontColor = false;
      model.numFmt = "";
    } else {
      model.raw = null;
      model.formula = null;
      model.cached = null;
      model.display = "";
      delete model.richText;
      delete model.phonetics;
      delete model.showPhonetic;
      delete model.sparkline;
      delete model.dataTableFormula;
    }
    storeCellModelForSheet(sheet, key, model);
    changed += 1;
  });
  if (!changed) return;
  context.writeCount += changed;
  context.touchedSheetIndexes.add(target.sheetIndex);
}



function vbaColorToCss(value) {
  const normalized = normalizeColorValue(value);
  if (normalized) return normalized;
  const numeric = Number(value);
  if (Number.isFinite(numeric)) {
    const color = Math.max(0, Math.min(0xffffff, Math.trunc(numeric)));
    return rgbToHex([color & 0xff, (color >> 8) & 0xff, (color >> 16) & 0xff]);
  }
  return "";
}



function isVbaNoneValue(value) {
  return /^(xlNone|xlColorIndexNone|xlColorIndexAutomatic|xlAutomatic|False|-4105|-4142)$/i.test(String(value ?? "").trim());
}



function vbaTruthy(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  return /^(true|yes|on|1|xlunderlinestylesingle|xlunderlinestyledouble)$/i.test(String(value ?? "").trim());
}



function vbaUnderlineTextDecoration(value) {
  const text = String(value ?? "").trim();
  if (/^(false|0|xlnone|xlunderlinestylenone)$/i.test(text)) return "";
  if (/xlunderlinestyledouble/i.test(text)) return "underline double";
  if (vbaTruthy(value)) return "underline";
  return "";
}



function vbaHorizontalAlignment(value) {
  const text = String(value ?? "").trim().toLowerCase();
  if (!text || /xlgeneral|general/.test(text)) return "";
  if (/xl(?:h)?aligncenter|xlcenter|center|中央/.test(text)) return "center";
  if (/xl(?:h)?alignright|xlright|right|右/.test(text)) return "right";
  if (/xl(?:h)?alignjustify|xljustify|justify|両端/.test(text)) return "justify";
  if (/xl(?:h)?alignleft|xlleft|left|左/.test(text)) return "left";
  return "";
}



function vbaVerticalAlignment(value) {
  const text = String(value ?? "").trim().toLowerCase();
  if (!text) return "";
  if (/xlvaligntop|xltop|top|上/.test(text)) return "top";
  if (/xlvalignbottom|xlbottom|bottom|下/.test(text)) return "bottom";
  if (/xlvaligncenter|xlcenter|center|middle|中央/.test(text)) return "middle";
  return "";
}



function vbaConstantValue(text) {
  const constants = {
    vbcrlf: "\r\n",
    vblf: "\n",
    vbcr: "\r",
    vbtab: "\t",
    vbnullstring: "",
    empty: null,
    null: null,
    xlnone: "xlNone",
    xlcolorindexnone: "xlColorIndexNone",
    xlautomatic: "xlAutomatic",
    xlcenter: "xlCenter",
    xlleft: "xlLeft",
    xlright: "xlRight",
    xlhalignleft: "xlLeft",
    xlhaligncenter: "xlCenter",
    xlhalignright: "xlRight",
    xlhalignjustify: "xlJustify",
    xljustify: "xlJustify",
    xltop: "xlTop",
    xlbottom: "xlBottom",
    xlvaligntop: "xlTop",
    xlvaligncenter: "xlCenter",
    xlvalignbottom: "xlBottom",
    xlunderlinestylenone: "xlUnderlineStyleNone",
    xlunderlinestylesingle: "xlUnderlineStyleSingle",
    xlunderlinestyledouble: "xlUnderlineStyleDouble",
    xlunderlinestyledoubleaccounting: "xlUnderlineStyleDouble",
    xlunderlinestylesingleaccounting: "xlUnderlineStyleSingle",
    xlcontinuous: "xlContinuous",
    xllinestylenone: "xlLineStyleNone",
    xldash: "xlDash",
    xldot: "xlDot",
    xldouble: "xlDouble",
    xlthin: "xlThin",
    xlmedium: "xlMedium",
    xlthick: "xlThick",
    xledgeleft: "xlEdgeLeft",
    xledgetop: "xlEdgeTop",
    xledgeright: "xlEdgeRight",
    xledgebottom: "xlEdgeBottom",
  };
  const key = String(text || "").trim().toLowerCase();
  return Object.prototype.hasOwnProperty.call(constants, key) ? constants[key] : undefined;
}



function writeVbaValueToRange(target, value, context) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return;
  ensureSheetSize(sheet, target.range.bottom, target.range.right);
  const sheetId = getSheetId(sheet.name);
  const arrayValues = value?.__webSheetVbaArray ? value.values || [] : null;
  let arrayIndex = 0;
  for (let row = target.range.top; row <= target.range.bottom; row += 1) {
    for (let col = target.range.left; col <= target.range.right; col += 1) {
      const cellValue = arrayValues ? arrayValues[Math.min(arrayIndex, Math.max(0, arrayValues.length - 1))] : value;
      writeCellContent(sheet, sheetId, row, col, cellValue);
      arrayIndex += 1;
      context.writeCount += 1;
    }
  }
  context.touchedSheetIndexes.add(target.sheetIndex);
}



function writeVbaFormulaR1C1ToRange(target, value, context) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return;
  ensureSheetSize(sheet, target.range.bottom, target.range.right);
  const sheetId = getSheetId(sheet.name);
  const text = String(value ?? "");
  for (let row = target.range.top; row <= target.range.bottom; row += 1) {
    for (let col = target.range.left; col <= target.range.right; col += 1) {
      const content = text.startsWith("=") ? convertVbaR1C1FormulaToA1(text, row, col) : value;
      writeCellContent(sheet, sheetId, row, col, content);
      context.writeCount += 1;
    }
  }
  context.touchedSheetIndexes.add(target.sheetIndex);
}



function convertVbaR1C1FormulaToA1(formula, baseRow, baseCol) {
  return splitFormulaStringLiterals(formula).map((part) => {
    if (part.string) return part.text;
    return part.text.replace(/R(\[[-+]?\d+\]|\d*)C(\[[-+]?\d+\]|\d*)/gi, (_match, rowToken, colToken) => {
      const row = resolveVbaR1C1Index(rowToken, baseRow);
      const col = resolveVbaR1C1Index(colToken, baseCol);
      const rowAbsolute = /^\d+$/.test(String(rowToken || ""));
      const colAbsolute = /^\d+$/.test(String(colToken || ""));
      return `${colAbsolute ? "$" : ""}${columnName(col)}${rowAbsolute ? "$" : ""}${row}`;
    });
  }).join("");
}



function splitFormulaStringLiterals(text) {
  const parts = [];
  let start = 0;
  let inString = false;
  const source = String(text || "");
  for (let index = 0; index < source.length; index += 1) {
    if (source[index] !== '"') continue;
    if (inString && source[index + 1] === '"') {
      index += 1;
      continue;
    }
    if (start < index) parts.push({ text: source.slice(start, index), string: inString });
    start = index;
    inString = !inString;
    index += 1;
    while (inString && index < source.length) {
      if (source[index] === '"') {
        if (source[index + 1] === '"') {
          index += 2;
          continue;
        }
        index += 1;
        break;
      }
      index += 1;
    }
    parts.push({ text: source.slice(start, index), string: true });
    start = index;
    inString = false;
    index -= 1;
  }
  if (start < source.length) parts.push({ text: source.slice(start), string: false });
  return parts;
}



function resolveVbaR1C1Index(token, base) {
  const text = String(token || "");
  const relative = text.match(/^\[([-+]?\d+)\]$/);
  if (relative) return Math.max(1, Math.round(Number(base) + Number(relative[1])));
  if (/^\d+$/.test(text)) return Math.max(1, Math.round(Number(text)));
  return Math.max(1, Math.round(Number(base) || 1));
}



function selectVbaRangeTarget(target) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return;
  ensureSheetSize(sheet, target.range.bottom, target.range.right);
  state.activeSheetIndex = target.sheetIndex;
  const range = expandRangeForMerges(sheet, { ...target.range, sheetIndex: target.sheetIndex });
  state.selected = { sheetIndex: target.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionAnchor = state.selected;
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  state.formulaEdit = null;
}



function openSheetCodeDialog(sheetIndex = state.activeSheetIndex) {
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) return;
  const module = normalizeSheetCodeModule(sheet.codeModule, sheet) || {
    name: sheetCodeModuleName(sheet),
    objectName: "Worksheet",
    procedure: "",
    code: "",
    updatedAt: "",
  };
  closeDataDialog({ keepState: true });
  const overlay = $(`
    <div class="excel-dialog-overlay vba-editor-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog vba-code-dialog" role="dialog" aria-modal="true" aria-labelledby="sheetCodeTitle" data-data-dialog="sheet-code" data-sheet-index="${sheetIndex}">
        <div class="excel-mini-titlebar" id="sheetCodeTitle">Microsoft Visual Basic for Applications - ${escapeHtml(state.model?.sourceName || "Book1")} - [${escapeHtml(module.name)}]</div>
        <div class="excel-mini-body data-dialog-body vba-code-body">
          <aside class="vba-project-pane" aria-label="プロジェクト">
            <div class="vba-project-title">プロジェクト - VBAProject</div>
            <div class="vba-project-tree">
              <div class="vba-project-root">VBAProject (${escapeHtml(state.model?.sourceName || "Book1")})</div>
              <div class="vba-project-folder">Cht WebSheet Objects</div>
              ${(state.model?.sheets || []).map((item, index) => `<button type="button" class="${index === sheetIndex ? "is-active" : ""}" data-sheet-code-nav="${index}">${escapeHtml(item.name)} (${escapeHtml(item.name)})</button>`).join("")}
              <button type="button" disabled>ThisWorkbook</button>
            </div>
          </aside>
          <section class="vba-code-pane">
            <div class="vba-code-selectors">
              <select data-data-field="sheetCodeObject">
                <option value="Worksheet"${module.objectName === "Worksheet" ? " selected" : ""}>Worksheet</option>
                <option value="General"${module.objectName === "General" ? " selected" : ""}>General</option>
              </select>
	            <select data-data-field="sheetCodeProcedure">
	                <option value=""${module.procedure === "" ? " selected" : ""}>(Declarations)</option>
	                <option value="SelectionChange"${module.procedure === "SelectionChange" ? " selected" : ""}>SelectionChange</option>
	                <option value="Change"${module.procedure === "Change" ? " selected" : ""}>Change</option>
	                <option value="Activate"${module.procedure === "Activate" ? " selected" : ""}>Activate</option>
	                <option value="BeforeDoubleClick"${module.procedure === "BeforeDoubleClick" ? " selected" : ""}>BeforeDoubleClick</option>
	                <option value="BeforeRightClick"${module.procedure === "BeforeRightClick" ? " selected" : ""}>BeforeRightClick</option>
	              </select>
            </div>
            <textarea class="macro-code-editor vba-code-editor" data-data-field="sheetCode" spellcheck="false">${escapeHtml(module.code || "")}</textarea>
          </section>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="sheet-code-ok">OK</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "sheet-code", sheetIndex };
  overlay.find("[data-data-field='sheetCode']").trigger("focus");
  setStatus(`コード ウィンドウを表示しました：${module.name}`);
}



function handleSheetCodeProcedureChange(event) {
  const dialog = event.target.closest?.("[data-data-dialog='sheet-code']");
  if (!dialog) return;
  const procedure = String(event.target.value || "");
  if (!procedure) return;
  const textArea = dialog.querySelector("[data-data-field='sheetCode']");
  if (!textArea) return;
  const stub = sheetCodeProcedureStub(procedure);
  if (!stub || String(textArea.value || "").includes(stub.signature)) return;
  const prefix = String(textArea.value || "").trimEnd();
  textArea.value = `${prefix ? `${prefix}\n\n` : ""}${stub.code}`;
  textArea.focus();
  textArea.selectionStart = textArea.selectionEnd = textArea.value.length;
}



function sheetCodeProcedureStub(procedure) {
  const stubs = {
    SelectionChange: {
      signature: "Private Sub Worksheet_SelectionChange",
      code: "Private Sub Worksheet_SelectionChange(ByVal Target As Range)\n\nEnd Sub",
    },
    Change: {
      signature: "Private Sub Worksheet_Change",
      code: "Private Sub Worksheet_Change(ByVal Target As Range)\n\nEnd Sub",
    },
    Activate: {
      signature: "Private Sub Worksheet_Activate",
      code: "Private Sub Worksheet_Activate()\n\nEnd Sub",
    },
    BeforeDoubleClick: {
      signature: "Private Sub Worksheet_BeforeDoubleClick",
      code: "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)\n\nEnd Sub",
    },
    BeforeRightClick: {
      signature: "Private Sub Worksheet_BeforeRightClick",
      code: "Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)\n\nEnd Sub",
    },
  };
  return stubs[procedure] || null;
}



function applySheetCodeDialog(dialog) {
  const sheetIndex = Number(dialog?.dataset.sheetIndex ?? state.dataDialog?.sheetIndex);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const code = String(dialog?.querySelector("[data-data-field='sheetCode']")?.value || "");
  const objectName = String(dialog?.querySelector("[data-data-field='sheetCodeObject']")?.value || "Worksheet");
  const procedure = String(dialog?.querySelector("[data-data-field='sheetCodeProcedure']")?.value || "");
  sheet.codeModule = code || sheet.codeModule
    ? {
        name: sheetCodeModuleName(sheet),
        objectName,
        procedure,
        code,
        updatedAt: new Date().toISOString(),
      }
    : null;
  closeDataDialog();
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: `${sheet.name} コード表示`,
  });
  scheduleWorkbookPersistence();
  setStatus(`コードを保存しました：${sheetCodeModuleName(sheet)}`);
}



function openMacroRecordDialog() {
  closeDataDialog({ keepState: true });
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog macro-record-dialog" role="dialog" aria-modal="true" aria-labelledby="macroRecordTitle" data-data-dialog="macro-record">
        <div class="excel-mini-titlebar" id="macroRecordTitle">マクロの記録</div>
        <div class="excel-mini-body data-dialog-body">
          <label class="excel-dialog-field"><span class="excel-dialog-caption">マクロ名(M):</span>
            <input type="text" data-data-field="macroRecordName" value="${escapeAttr(defaultMacroName())}" />
          </label>
          <label class="excel-dialog-field macro-shortcut-field"><span class="excel-dialog-caption">ショートカット キー(K):</span><span class="macro-shortcut-input"><span>Ctrl+</span><input type="text" maxlength="1" data-data-field="macroRecordShortcut" value="" /></span></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">マクロの保存先(I):</span>
            <select data-data-field="macroRecordScope">
              <option value="workbook" selected>作業中のブック</option>
              <option value="newWorkbook">新しいブック</option>
              <option value="personal">個人用マクロ ブック</option>
            </select>
          </label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">説明(D):</span><textarea data-data-field="macroRecordDescription"></textarea></label>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="macro-record-ok">OK</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "macro-record" };
  overlay.find("[data-data-field='macroRecordName']").trigger("focus").trigger("select");
}



function startMacroRecordingFromDialog(dialog) {
  const validation = validateMacroName(dialog?.querySelector("[data-data-field='macroRecordName']")?.value || "");
  if (!validation.valid) {
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    return;
  }
  const macroStateBefore = captureWorkbookMacroState();
  state.model.macros = Array.isArray(state.model.macros) ? state.model.macros : [];
  const macro = {
    name: validation.name,
    code: "",
    shortcut: String(dialog?.querySelector("[data-data-field='macroRecordShortcut']")?.value || "").trim().slice(0, 1),
    description: String(dialog?.querySelector("[data-data-field='macroRecordDescription']")?.value || ""),
    scope: dialog?.querySelector("[data-data-field='macroRecordScope']")?.value || "workbook",
    createdAt: new Date().toISOString(),
    recordedAt: new Date().toISOString(),
  };
  state.model.macros.push(macro);
  state.model.macroRecording = {
    name: macro.name,
    index: state.model.macros.length - 1,
    relativeReferences: state.model.macroRelativeReferences === true,
    startedAt: new Date().toISOString(),
  };
  closeDataDialog();
  commitWorkbookMacroChange(macroStateBefore, `マクロの記録を開始しました：${macro.name}`, `マクロの記録 ${macro.name}`);
}



function stopMacroRecording() {
  const recording = state.model?.macroRecording;
  if (!recording) {
    setStatus("記録中のマクロはありません。");
    return;
  }
  const macro = state.model.macros?.[Number(recording.index)];
  const macroStateBefore = captureWorkbookMacroState();
  if (macro) macro.updatedAt = new Date().toISOString();
  state.model.macroRecording = null;
  commitWorkbookMacroChange(
    macroStateBefore,
    `マクロの記録を終了しました：${recording.name || macro?.name || ""}`,
    `マクロの記録終了 ${recording.name || macro?.name || ""}`,
  );
}



function toggleMacroRelativeReferences() {
  if (!state.model) return;
  const macroStateBefore = captureWorkbookMacroState();
  state.model.macroRelativeReferences = state.model.macroRelativeReferences !== true;
  if (state.model.macroRecording) {
    state.model.macroRecording.relativeReferences = state.model.macroRelativeReferences === true;
  }
  commitWorkbookMacroChange(
    macroStateBefore,
    state.model.macroRelativeReferences ? "相対参照で記録を有効にしました。" : "相対参照で記録を無効にしました。",
    "マクロ 相対参照で記録",
  );
}



function macroSecurityOptionRows(security) {
  return MACRO_SECURITY_OPTIONS.map((option) => `
    <label class="data-dialog-check macro-security-option">
      <input type="radio" name="macroSecurityLevel" data-data-field="macroSecurityLevel" value="${escapeAttr(option.value)}"${security.level === option.value ? " checked" : ""} />
      <span>
        <strong>${escapeHtml(option.label)}</strong>
        <small>${escapeHtml(option.description)}</small>
      </span>
    </label>
  `).join("");
}



function openMacroSecurityDialog() {
  closeDataDialog({ keepState: true });
  const security = normalizeMacroSecurity(state.model?.macroSecurity);
  const navItems = MACRO_SECURITY_NAV_ITEMS.map((label) => `<button class="${label === "マクロの設定" ? "is-active" : ""}" type="button" tabindex="-1">${escapeHtml(label)}</button>`).join("");
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog macro-security-dialog" role="dialog" aria-modal="true" aria-labelledby="macroSecurityTitle" data-data-dialog="macro-security">
        <div class="excel-mini-titlebar" id="macroSecurityTitle">セキュリティ センター</div>
        <div class="excel-mini-body data-dialog-body macro-security-body">
          <div class="macro-security-layout">
            <div class="macro-security-nav" aria-label="セキュリティ センターのカテゴリ">
              ${navItems}
            </div>
            <div class="macro-security-panel">
              <div class="macro-security-heading">マクロの設定</div>
              <fieldset class="data-dialog-fieldset macro-security-options">
                <legend>VBA マクロの設定</legend>
                ${macroSecurityOptionRows(security)}
              </fieldset>
              <label class="data-dialog-check macro-security-trust">
                <input type="checkbox" data-data-field="macroTrustVbaObjectModel"${security.trustVbaObjectModel ? " checked" : ""} />
                <span>VBA プロジェクト オブジェクト モデルへのアクセスを信頼する</span>
              </label>
	              <div class="data-dialog-summary is-warning">WebSheet は設定をブックに保存し、実行時は対応済みの安全な VBA マクロ サブセットだけを処理します。</div>
            </div>
          </div>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="macro-security-ok">OK</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "macro-security" };
  overlay.find("[data-data-field='macroSecurityLevel']:checked").trigger("focus");
}



function applyMacroSecurityDialog(dialog) {
  if (!state.model) return;
  const selectedLevel = dialog?.querySelector("input[name='macroSecurityLevel']:checked")?.value || DEFAULT_MACRO_SECURITY_LEVEL;
  const macroStateBefore = captureWorkbookMacroState();
  state.model.macroSecurity = normalizeMacroSecurity({
    level: selectedLevel,
    trustVbaObjectModel: dialog?.querySelector("[data-data-field='macroTrustVbaObjectModel']")?.checked === true,
    updatedAt: new Date().toISOString(),
  });
  closeDataDialog();
  const option = MACRO_SECURITY_OPTIONS.find((item) => item.value === state.model.macroSecurity.level);
  commitWorkbookMacroChange(
    macroStateBefore,
    `マクロのセキュリティを更新しました：${option?.label || "既定"}`,
    "マクロ セキュリティ",
  );
}



function createMacroFromDialog(dialog) {
  const validation = validateMacroName(macroNameFromDialog(dialog));
  if (!validation.valid) {
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    return;
  }
  const macroStateBefore = captureWorkbookMacroState();
  state.model.macros = Array.isArray(state.model.macros) ? state.model.macros : [];
  const macro = { name: validation.name, code: "", shortcut: "", description: "", createdAt: new Date().toISOString() };
  state.model.macros.push(macro);
  commitWorkbookMacroChange(macroStateBefore, `マクロを作成しました：${macro.name}`, `マクロ作成 ${macro.name}`);
  openMacrosDialog({ selectedIndex: state.model.macros.length - 1 });
}



function selectedMacroFromDialog(dialog) {
  const index = selectedMacroIndexFromDialog(dialog);
  const macro = Array.isArray(state.model?.macros) ? state.model.macros[index] : null;
  return macro ? { macro, index } : null;
}



function openMacroStepDialog(index) {
  const macro = state.model?.macros?.[index];
  if (!macro) return;
  const statements = parseVbaMacroStatements(macro.code || "");
  if (!statements.length) {
    showUnsupportedMacroRunMessage();
    return;
  }
  closeDataDialog({ keepState: true });
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog macro-step-dialog" role="dialog" aria-modal="true" aria-labelledby="macroStepTitle" data-data-dialog="macro-step" data-macro-index="${index}">
        <div class="excel-mini-titlebar" id="macroStepTitle">ステップ イン - ${escapeHtml(macro.name || `Macro${index + 1}`)}</div>
        <div class="excel-mini-body data-dialog-body macro-step-body">
          <div class="macro-step-summary" data-macro-step-summary></div>
          <ol class="macro-step-list" data-macro-step-list></ol>
          <div class="macro-step-log" data-macro-step-log aria-live="polite"></div>
        </div>
        <div class="excel-mini-actions macro-step-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="macro-step-next">ステップ イン(S)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="macro-step-continue">継続(C)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="macro-step-edit">編集(E)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">閉じる</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = {
    type: "macro-step",
    index,
    pointer: 0,
    statements,
    messages: [],
    unsupported: [],
  };
  const dialog = overlay[0].querySelector("[data-data-dialog='macro-step']");
  syncMacroStepDialog(dialog);
  overlay.find("[data-data-dialog-action='macro-step-next']").trigger("focus");
}



function syncMacroStepDialog(dialog) {
  if (!dialog || state.dataDialog?.type !== "macro-step") return;
  const pointer = Math.max(0, Math.min(state.dataDialog.statements.length, Number(state.dataDialog.pointer) || 0));
  const statements = state.dataDialog.statements || [];
  const completed = pointer >= statements.length;
  const summary = dialog.querySelector("[data-macro-step-summary]");
  if (summary) {
    summary.textContent = completed
      ? `${statements.length} / ${statements.length} ステートメントを実行しました。`
      : `${pointer + 1} / ${statements.length}：次のステートメントを実行します。`;
  }
  const list = dialog.querySelector("[data-macro-step-list]");
  if (list) list.innerHTML = macroStepListHtml(statements, pointer);
  const log = dialog.querySelector("[data-macro-step-log]");
  if (log) log.innerHTML = macroStepLogHtml(state.dataDialog.messages, state.dataDialog.unsupported);
  ["macro-step-next", "macro-step-continue"].forEach((action) => {
    const button = dialog.querySelector(`[data-data-dialog-action='${action}']`);
    if (button) button.disabled = completed;
  });
}



function macroStepListHtml(statements, pointer) {
  return statements.map((statement, index) => {
    const stateClass = index < pointer ? " is-complete" : index === pointer ? " is-current" : "";
    return `<li class="macro-step-item${stateClass}" data-macro-step-item="${index + 1}">
      <span class="macro-step-index">${index + 1}</span>
      <code>${escapeHtml(statement)}</code>
    </li>`;
  }).join("");
}



function macroStepLogHtml(messages = [], unsupported = []) {
  const rows = [
    ...messages.map((message) => `<div class="macro-step-log-row">${escapeHtml(message)}</div>`),
    ...unsupported.map((statement) => `<div class="macro-step-log-row is-warning">未対応: ${escapeHtml(statement)}</div>`),
  ];
  return rows.length ? rows.join("") : '<div class="macro-step-log-empty">実行メッセージはありません。</div>';
}



function executeMacroStepDialog(dialog, mode = "next") {
  if (!dialog || state.dataDialog?.type !== "macro-step") return;
  const index = Number(dialog.dataset.macroIndex ?? state.dataDialog.index);
  const macro = state.model?.macros?.[index];
  const statements = state.dataDialog.statements || [];
  const pointer = Math.max(0, Math.min(statements.length, Number(state.dataDialog.pointer) || 0));
  if (!macro || pointer >= statements.length) {
    syncMacroStepDialog(dialog);
    return;
  }
  const nextPointer = mode === "continue" ? statements.length : pointer + 1;
  const batch = statements.slice(pointer, nextPointer);
  const context = executeWorkbookMacroStatements(macro, index, batch, {
    historyLabel: `マクロ ステップ ${macro.name || `Macro${index + 1}`} ${pointer + 1}-${nextPointer}`,
    statusLabel: `${macro.name || "マクロ"} ステップ`,
    showMessages: false,
    updateLastRun: nextPointer >= statements.length,
  });
  state.dataDialog = {
    ...state.dataDialog,
    type: "macro-step",
    index,
    pointer: nextPointer,
    statements,
    messages: [...(state.dataDialog.messages || []), ...(context.messages || [])],
    unsupported: [...(state.dataDialog.unsupported || []), ...(context.unsupported || [])],
  };
  syncMacroStepDialog(dialog);
}



function openMacroEditorDialog(index) {
  const macro = state.model?.macros?.[index];
  if (!macro) return;
  closeDataDialog({ keepState: true });
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog macro-editor-dialog" role="dialog" aria-modal="true" aria-labelledby="macroEditorTitle" data-data-dialog="macro-editor" data-macro-index="${index}">
        <div class="excel-mini-titlebar" id="macroEditorTitle">マクロの編集 - ${escapeHtml(macro.name || `Macro${index + 1}`)}</div>
        <div class="excel-mini-body data-dialog-body">
          <textarea class="macro-code-editor" data-data-field="macroCode" spellcheck="false">${escapeHtml(macro.code || "")}</textarea>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="macro-editor-ok">OK</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "macro-editor", index };
  overlay.find("[data-data-field='macroCode']").trigger("focus");
}



function applyMacroEditorDialog(dialog) {
  const index = Number(dialog?.dataset.macroIndex ?? state.dataDialog?.index);
  const macro = state.model?.macros?.[index];
  if (!macro) return;
  const macroStateBefore = captureWorkbookMacroState();
  macro.code = String(dialog?.querySelector("[data-data-field='macroCode']")?.value || "");
  macro.updatedAt = new Date().toISOString();
  commitWorkbookMacroChange(macroStateBefore, `マクロを更新しました：${macro.name}`, `マクロ編集 ${macro.name}`);
  openMacrosDialog({ selectedIndex: index });
}



function openMacroOptionsDialog(index) {
  const macro = state.model?.macros?.[index];
  if (!macro) return;
  closeDataDialog({ keepState: true });
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog macro-options-dialog" role="dialog" aria-modal="true" aria-labelledby="macroOptionsTitle" data-data-dialog="macro-options" data-macro-index="${index}">
        <div class="excel-mini-titlebar" id="macroOptionsTitle">マクロ オプション</div>
        <div class="excel-mini-body data-dialog-body">
          <div class="macro-options-name">${escapeHtml(macro.name || `Macro${index + 1}`)}</div>
          <label class="excel-dialog-field macro-shortcut-field"><span class="excel-dialog-caption">ショートカット キー(K):</span><span class="macro-shortcut-input"><span>Ctrl+</span><input type="text" maxlength="1" data-data-field="macroShortcut" value="${escapeAttr(macro.shortcut || "")}" /></span></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">説明(D):</span><textarea data-data-field="macroDescription">${escapeHtml(macro.description || "")}</textarea></label>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="macro-options-ok">OK</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "macro-options", index };
  overlay.find("[data-data-field='macroShortcut']").trigger("focus");
}



function applyMacroOptionsDialog(dialog) {
  const index = Number(dialog?.dataset.macroIndex ?? state.dataDialog?.index);
  const macro = state.model?.macros?.[index];
  if (!macro) return;
  const macroStateBefore = captureWorkbookMacroState();
  macro.shortcut = String(dialog?.querySelector("[data-data-field='macroShortcut']")?.value || "").trim().slice(0, 1);
  macro.description = String(dialog?.querySelector("[data-data-field='macroDescription']")?.value || "");
  macro.updatedAt = new Date().toISOString();
  commitWorkbookMacroChange(macroStateBefore, `マクロ オプションを更新しました：${macro.name}`, `マクロ オプション ${macro.name}`);
  openMacrosDialog({ selectedIndex: index });
}



function deleteMacroFromDialog(dialog) {
  const selected = selectedMacroFromDialog(dialog);
  if (!selected) return;
  showMessageBox(`マクロ「${selected.macro.name || ""}」を削除しますか?`, {
    title: "Cht WebSheet",
    tone: "warning",
    buttons: [
      {
        label: "削除",
        result: "delete",
        callback: () => {
          const macroStateBefore = captureWorkbookMacroState();
          state.model.macros.splice(selected.index, 1);
          commitWorkbookMacroChange(macroStateBefore, "マクロを削除しました。", `マクロ削除 ${selected.macro.name || ""}`);
          openMacrosDialog({ selectedIndex: Math.min(selected.index, state.model.macros.length - 1) });
        },
      },
      { label: "キャンセル", result: "cancel" },
    ],
    cancelResult: "cancel",
  });
}



function normalizedPageLayoutSettings(value = {}, sheet = activeSheet()) {
  value = value || {};
  const preset = PAGE_LAYOUT_MARGIN_PRESETS[value.marginPreset] ? value.marginPreset : "normal";
  const paper = PAGE_LAYOUT_PAPER_SIZES.some((item) => item.value === value.paperSize) ? value.paperSize : "A4";
  const numberOrDefault = (raw, fallback, min, max) => {
    const number = Number(raw);
    return Number.isFinite(number) ? Math.max(min, Math.min(max, number)) : fallback;
  };
  const normalizedMargins = {
    ...PAGE_LAYOUT_MARGIN_PRESETS[preset],
    ...(value.margins && typeof value.margins === "object" ? value.margins : {}),
  };
  const fitToPage = value.fitToPage === true || value.fitToPage === "true" || value.fitToPage === 1 || value.fitToPage === "1";
  return {
    theme: PAGE_LAYOUT_THEMES.some((theme) => theme.value === value.theme) ? value.theme : "office",
    colorScheme: PAGE_LAYOUT_THEMES.some((theme) => theme.value === value.colorScheme) ? value.colorScheme : value.theme || "office",
    fontScheme: PAGE_LAYOUT_THEMES.some((theme) => theme.value === value.fontScheme) ? value.fontScheme : value.theme || "office",
    effectScheme: pageLayoutEffectPreset(value.effectScheme || value.theme)?.value || "office",
    marginPreset: preset,
    margins: {
      top: numberOrDefault(normalizedMargins.top, PAGE_LAYOUT_MARGIN_PRESETS[preset].top, 0, 3),
      bottom: numberOrDefault(normalizedMargins.bottom, PAGE_LAYOUT_MARGIN_PRESETS[preset].bottom, 0, 3),
      left: numberOrDefault(normalizedMargins.left, PAGE_LAYOUT_MARGIN_PRESETS[preset].left, 0, 3),
      right: numberOrDefault(normalizedMargins.right, PAGE_LAYOUT_MARGIN_PRESETS[preset].right, 0, 3),
      header: numberOrDefault(normalizedMargins.header, PAGE_LAYOUT_MARGIN_PRESETS[preset].header, 0, 3),
      footer: numberOrDefault(normalizedMargins.footer, PAGE_LAYOUT_MARGIN_PRESETS[preset].footer, 0, 3),
    },
    orientation: value.orientation === "landscape" ? "landscape" : "portrait",
    paperSize: paper,
    printArea: value.printArea?.range ? { range: { ...value.printArea.range }, label: rangeToLabel(value.printArea.range) } : null,
    pageBreaks: {
      rows: [...new Set((value.pageBreaks?.rows || []).map(Number).filter((row) => row > 1))].sort((a, b) => a - b),
      cols: [...new Set((value.pageBreaks?.cols || []).map(Number).filter((col) => col > 1))].sort((a, b) => a - b),
    },
    printTitles: {
      rows: normalizePageSetupPrintTitleRows(value.printTitles?.rows, sheet),
      cols: normalizePageSetupPrintTitleColumns(value.printTitles?.cols, sheet),
    },
    background: normalizePageLayoutBackground(value.background),
    fitToPage,
    fitToWidth: fitToPage ? normalizeFitToPageCount(value.fitToWidth) : "auto",
    fitToHeight: fitToPage ? normalizeFitToPageCount(value.fitToHeight) : "auto",
    scale: Math.round(numberOrDefault(value.scale, 100, 10, 400)),
    pageOrder: ["downThenOver", "overThenDown"].includes(value.pageOrder) ? value.pageOrder : "downThenOver",
    firstPageNumber: Number.isFinite(Number(value.firstPageNumber)) ? Math.max(1, Math.min(9999, Math.round(Number(value.firstPageNumber)))) : 1,
    centerHorizontal: Boolean(value.centerHorizontal),
    centerVertical: Boolean(value.centerVertical),
    blackAndWhite: Boolean(value.blackAndWhite),
    draftQuality: Boolean(value.draftQuality),
    printGridLines: Boolean(value.printGridLines),
    printHeadings: Boolean(value.printHeadings),
    comments: ["none", "atEnd", "asDisplayed"].includes(value.comments) ? value.comments : "none",
    cellErrors: ["displayed", "blank", "dash", "na"].includes(value.cellErrors) ? value.cellErrors : "displayed",
    headerFooter: {
      header: String(value.headerFooter?.header ?? value.header ?? ""),
      footer: String(value.headerFooter?.footer ?? value.footer ?? ""),
    },
  };
}



function excelImportedPageMargins(margins = {}) {
  const normalized = {};
  const keys = Object.keys(EXCEL_DEFAULT_PAGE_MARGINS);
  let hasInvalidDefaultRecord = false;
  keys.forEach((key) => {
    const number = Number(margins?.[key]);
    if (!Number.isFinite(number) || number <= 0) hasInvalidDefaultRecord = true;
    normalized[key] = number;
  });
  return hasInvalidDefaultRecord ? { ...EXCEL_DEFAULT_PAGE_MARGINS } : normalized;
}



function normalizePageLayoutBackground(value = null) {
  if (!value) return null;
  if (typeof value === "string") {
    const text = value.trim();
    if (!text || !text.startsWith("data:image/")) return null;
    return {
      src: text,
      name: "背景",
      mimeType: imageMimeType("png"),
      mode: "tile",
    };
  }
  if (typeof value !== "object") return null;
  const src = String(value.src || value.dataUrl || "").trim();
  if (!src || !src.startsWith("data:image/")) return null;
  const width = Number(value.width);
  const height = Number(value.height);
  return {
    src,
    name: String(value.name || "背景"),
    mimeType: String(value.mimeType || imageMimeType(fileExtension(value.name || "png")) || "image/png"),
    mode: value.mode === "stretch" ? "stretch" : "tile",
    ...(Number.isFinite(width) && width > 0 ? { width: Math.round(width) } : {}),
    ...(Number.isFinite(height) && height > 0 ? { height: Math.round(height) } : {}),
  };
}



function hasPageBackground(layout = ensurePageLayoutSettings(activeSheet())) {
  return Boolean(normalizePageLayoutBackground(layout?.background)?.src);
}



function normalizePageSetupPrintTitleRows(range, sheet = activeSheet()) {
  if (!range) return null;
  const top = Math.max(1, Math.round(Number(range.top) || 1));
  const bottom = Math.max(top, Math.round(Number(range.bottom) || top));
  return {
    sheetIndex: range.sheetIndex ?? state.activeSheetIndex,
    top,
    bottom,
    left: 1,
    right: Math.max(1, Math.round(Number(sheet?.colCount || range.right || 1))),
    fullRow: true,
    label: `$${top}:$${bottom}`,
  };
}



function normalizePageSetupPrintTitleColumns(range, sheet = activeSheet()) {
  if (!range) return null;
  const left = Math.max(1, Math.round(Number(range.left) || 1));
  const right = Math.max(left, Math.round(Number(range.right) || left));
  return {
    sheetIndex: range.sheetIndex ?? state.activeSheetIndex,
    top: 1,
    bottom: Math.max(1, Math.round(Number(sheet?.rowCount || range.bottom || 1))),
    left,
    right,
    fullColumn: true,
    label: `$${columnName(left)}:$${columnName(right)}`,
  };
}



function normalizeFitToPageCount(value) {
  if (String(value) === "auto" || value == null || value === "") return "auto";
  const number = Math.round(Number(value));
  return Number.isFinite(number) && number > 0 ? String(Math.max(1, Math.min(99, number))) : "auto";
}



function formatPageSetupMarginValue(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return "";
  const rounded = Math.round(number * 100) / 100;
  return String(rounded).replace(/(\.\d*?)0+$/, "$1").replace(/\.$/, "");
}



function parsePageSetupMarginValue(input, key) {
  const fallback = EXCEL_DEFAULT_PAGE_MARGINS[key] ?? 0;
  const raw = String(input?.value ?? "").trim();
  if (!raw) return fallback;
  const number = Number(raw);
  return Number.isFinite(number) ? Math.max(0, Math.min(3, Math.round(number * 100) / 100)) : fallback;
}



function ensurePageLayoutSettings(sheet = activeSheet()) {
  if (!sheet) return normalizedPageLayoutSettings({}, null);
  sheet.pageLayout = normalizedPageLayoutSettings(sheet.pageLayout, sheet);
  return sheet.pageLayout;
}



function commitPageLayoutChange(mutator, label, statusMessage, options = {}) {
  const sheet = activeSheet();
  if (!sheet) return;
  const workbookViewBefore = options.captureWorkbookView ? captureWorkbookViewState() : null;
  const before = captureSheetForHistory(sheet);
  const workbookThemeBefore = captureWorkbookThemeState();
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  const layout = ensurePageLayoutSettings(sheet);
  mutator(layout, sheet);
  sheet.pageLayout = normalizedPageLayoutSettings(layout, sheet);
  if (typeof options.afterMutate === "function") options.afterMutate(sheet.pageLayout, sheet);
  refreshActiveSheetViewSnapshot(sheet);
  saveActiveWindowViewState();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore: before,
    sheetAfter: captureSheetForHistory(sheet),
    workbookThemeBefore,
    workbookThemeAfter: captureWorkbookThemeState(),
    ...(workbookViewBefore ? {
      workbookViewBefore,
      workbookViewAfter: captureWorkbookViewState(),
    } : {}),
    selectionBefore,
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges,
    selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
    label,
  });
  if (options.render !== false) renderWorkbook();
  syncPageLayoutRibbonControls();
  scheduleWorkbookPersistence();
  setStatus(statusMessage || label);
}



function openPageLayoutThemeMenu(trigger = null) {
  const layout = ensurePageLayoutSettings(activeSheet());
  const rows = PAGE_LAYOUT_THEMES
    .map((theme) => layoutMenuButtonHtml(`theme:${theme.value}`, theme.label, "palette", {
      checkable: true,
      checked: layout.theme === theme.value,
      description: pageLayoutThemeDescription(theme, "テーマの色とフォントを適用します。"),
    }))
    .join("");
  openPageLayoutPresetMenu(trigger, "layout-themes", "テーマ", rows);
}



function openPageLayoutColorMenu(trigger = null) {
  const layout = ensurePageLayoutSettings(activeSheet());
  const rows = PAGE_LAYOUT_THEMES
    .map((theme) => layoutMenuButtonHtml(`colors:${theme.value}`, theme.label, "paint-bucket", {
      checkable: true,
      checked: layout.colorScheme === theme.value,
      description: pageLayoutThemeDescription(theme, "配色だけを変更します。"),
    }))
    .join("");
  openPageLayoutPresetMenu(trigger, "layout-colors", "配色", rows);
}



function openPageLayoutFontMenu(trigger = null) {
  const layout = ensurePageLayoutSettings(activeSheet());
  const rows = PAGE_LAYOUT_THEMES.map((theme) => {
    const font = `${theme.fonts.heading} / ${theme.fonts.body}`;
    return layoutMenuButtonHtml(`fonts:${theme.value}`, `${theme.label} - ${font}`, "type", {
      checkable: true,
      checked: layout.fontScheme === theme.value,
      description: "見出し / 本文フォントの組み合わせを適用します。",
    });
  }).join("");
  openPageLayoutPresetMenu(trigger, "layout-fonts", "フォント", rows);
}



function openPageLayoutEffectsMenu(trigger = null) {
  const layout = ensurePageLayoutSettings(activeSheet());
  const rows = PAGE_LAYOUT_EFFECT_PRESETS
    .map((effect) => layoutMenuButtonHtml(`effects:${effect.value}`, effect.label, "sparkles", {
      checkable: true,
      checked: layout.effectScheme === effect.value,
      description: pageLayoutEffectDescription(effect),
    }))
    .join("");
  openPageLayoutPresetMenu(trigger, "layout-effects", "効果", rows);
}



function openPageLayoutMarginsMenu(trigger = null) {
  const layout = ensurePageLayoutSettings(activeSheet());
  const rows = Object.entries(PAGE_LAYOUT_MARGIN_PRESETS)
    .map(([key, margin]) => layoutMenuButtonHtml(`margins:${key}`, `${margin.label}  上 ${margin.top}" 左 ${margin.left}"`, "panel-top", {
      checkable: true,
      checked: layout.marginPreset === key,
      description: `下 ${margin.bottom}" 右 ${margin.right}" ヘッダー ${margin.header}" フッター ${margin.footer}"`,
    }))
    .join("");
  openPageLayoutPresetMenu(trigger, "layout-margins", "余白", `${rows}<div class="home-command-menu-separator" role="separator"></div>${layoutMenuButtonHtml("page-setup", "ユーザー設定の余白(A)...", "settings")}`);
}



function openPageLayoutOrientationMenu(trigger = null) {
  const layout = ensurePageLayoutSettings(activeSheet());
  openPageLayoutPresetMenu(trigger, "layout-orientation", "印刷の向き", `
    ${layoutMenuButtonHtml("orientation:portrait", "縦(P)", "panel-top", { checkable: true, checked: layout.orientation === "portrait", description: "用紙を縦方向にします。" })}
    ${layoutMenuButtonHtml("orientation:landscape", "横(L)", "rotate-cw", { checkable: true, checked: layout.orientation === "landscape", description: "用紙を横方向にします。" })}
  `);
}



function openPageLayoutSizeMenu(trigger = null) {
  const layout = ensurePageLayoutSettings(activeSheet());
  const rows = PAGE_LAYOUT_PAPER_SIZES
    .map((paper) => layoutMenuButtonHtml(`paper:${paper.value}`, `${paper.label}  ${paper.width} x ${paper.height}`, "file-spreadsheet", {
      checkable: true,
      checked: layout.paperSize === paper.value,
      description: `幅 ${paper.width}" / 高さ ${paper.height}"`,
    }))
    .join("");
  openPageLayoutPresetMenu(trigger, "layout-size", "サイズ", rows);
}



function openPageLayoutPrintAreaMenu(trigger = null) {
  const layout = ensurePageLayoutSettings(activeSheet());
  const selection = activeSelectionRange();
  const selectionLabel = selection ? rangeToLabel(materializeSelectionRange(selection) || selection) : "";
  openPageLayoutPresetMenu(trigger, "layout-print-area", "印刷範囲", `
    ${layoutMenuButtonHtml("print-area:set", "印刷範囲の設定(S)", "square-dashed", { description: selectionLabel ? `${selectionLabel} を印刷範囲に設定します。` : "選択範囲を印刷範囲に設定します。" })}
    ${layoutMenuButtonHtml("print-area:clear", "印刷範囲のクリア(C)", "eraser", { disabled: !layout.printArea?.range, description: layout.printArea?.label ? `現在: ${layout.printArea.label}` : "設定済みの印刷範囲はありません。" })}
  `);
}



function openPageLayoutBreaksMenu(trigger = null) {
  const layout = ensurePageLayoutSettings(activeSheet());
  const hasBreaks = Boolean(layout.pageBreaks?.rows?.length || layout.pageBreaks?.cols?.length);
  const targets = pageBreakTargetsForSelection(activeSheet());
  const currentBreak = pageLayoutCurrentSelectionBreakState(layout, targets);
  const targetLabel = pageBreakTargetsLabel(targets);
  openPageLayoutPresetMenu(trigger, "layout-breaks", "改ページ", `
    ${layoutMenuButtonHtml("breaks:insert", "改ページの挿入(I)", "table-2", { disabled: !targets.canInsert || currentBreak, description: !targets.canInsert ? "先頭行または先頭列には改ページを挿入できません。" : currentBreak ? "選択位置には既に手動改ページがあります。" : `${targetLabel}に手動改ページを挿入します。` })}
    ${layoutMenuButtonHtml("breaks:remove", "改ページの解除(R)", "eraser", { disabled: !currentBreak, description: currentBreak ? `${targetLabel}の手動改ページを解除します。` : "選択位置に手動改ページはありません。" })}
    ${layoutMenuButtonHtml("breaks:reset", "すべての改ページを解除(A)", "delete", { disabled: !hasBreaks, description: hasBreaks ? "シート上のすべての手動改ページを解除します。" : "手動改ページはありません。" })}
  `);
}



function openPageLayoutPresetMenu(trigger, kind, label, rows) {
  if (!trigger) {
    if (kind === "layout-margins") applyPageMarginsPreset("normal");
    return;
  }
  openHomeCommandMenu(trigger, kind, label, rows);
}



function layoutMenuButtonHtml(action, label, icon, options = {}) {
  return homeCommandActionButtonHtml("data-layout-menu-action", action, label, icon, options);
}



function pageLayoutThemeDescription(theme, suffix = "") {
  const colors = (theme.colors || [])
    .slice(0, 6)
    .map((color) => normalizeColorValue(color) || `#${String(color || "").replace(/^#/, "")}`)
    .join(" ");
  const fonts = `${theme.fonts?.heading || ""} / ${theme.fonts?.body || ""}`.trim();
  return [colors, fonts, suffix].filter(Boolean).join("  ");
}



function pageLayoutEffectDescription(effect) {
  const patch = normalizedPageLayoutEffectShapePatch(effect?.shape);
  const parts = [];
  if (patch.shadow) parts.push("影");
  if (patch.glow) parts.push("光彩");
  if (patch.softEdges) parts.push("ぼかし");
  if (patch.reflection) parts.push("反射");
  if (!parts.length) parts.push("効果なし");
  return `${parts.join(" / ")}  ${effect?.description || "図形のテーマ効果を適用します。"}`;
}



function pageLayoutEffectPreset(value) {
  const key = String(value || "").trim();
  return PAGE_LAYOUT_EFFECT_PRESETS.find((effect) => effect.value === key) || PAGE_LAYOUT_EFFECT_PRESETS[0];
}



function normalizedPageLayoutThemeEffects(value = {}) {
  const preset = pageLayoutEffectPreset(typeof value === "string" ? value : value?.scheme || value?.value);
  return {
    scheme: preset.value,
    label: preset.label,
    shape: normalizedPageLayoutEffectShapePatch(value?.shape || preset.shape),
  };
}



function normalizedPageLayoutEffectShapePatch(shape = {}) {
  const source = shape && typeof shape === "object" ? shape : {};
  return {
    shadow: normalizeShapeShadow(source.shadow) || "",
    glow: normalizeShapeGlow(source.glow),
    softEdges: normalizeShapeSoftEdges(source.softEdges),
    reflection: normalizeShapeReflection(source.reflection),
  };
}



function pageLayoutEffectShapePatch(effectKey) {
  return cloneJsonLikeValue(normalizedPageLayoutThemeEffects(effectKey).shape || {});
}



function applyPageLayoutEffectToSelectedShapes(effectKey) {
  const sheet = activeSheet();
  const patch = pageLayoutEffectShapePatch(effectKey);
  selectedImageObjects().forEach((image) => {
    if (!image?.shape || isLineShapeObject(image)) return;
    image.shape = {
      ...(image.shape || {}),
      ...patch,
    };
    refreshShapeSource(image);
  });
  if (sheet) updateConnectedLinesForTargets(sheet, selectedImageObjects().filter((image) => image?.shape && !isLineShapeObject(image)).map((image) => image.id));
}



function defaultShapeThemeEffectPatchForType(type, style = {}) {
  if (isLineShapeType(type)) return {};
  if (!style?.useThemeEffects && !style?.applyThemeEffects) return {};
  const hasEffectOverride = ["shadow", "glow", "softEdges", "reflection"].some((key) => Object.prototype.hasOwnProperty.call(style, key));
  if (hasEffectOverride) return {};
  const sheet = state.model ? activeSheet() : null;
  const effectKey = sheet ? ensurePageLayoutSettings(sheet)?.effectScheme || "office" : "office";
  return pageLayoutEffectShapePatch(effectKey);
}



function pageLayoutCurrentSelectionBreakState(layout = ensurePageLayoutSettings(activeSheet()), targets = pageBreakTargetsForSelection(activeSheet())) {
  const rows = layout.pageBreaks?.rows || [];
  const cols = layout.pageBreaks?.cols || [];
  return targets.rows.some((row) => rows.includes(row)) || targets.cols.some((col) => cols.includes(col));
}



function pageBreakTargetsForSelection(sheet = activeSheet(), range = activeSelectionRanges()[0] || activeSelectionRange()) {
  if (!sheet) return { rows: [], cols: [], canInsert: false, range: null };
  const targetRange = range || selectionRangeFromPoints(
    selectionToolbarPoint() || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 },
    selectionToolbarPoint() || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 },
  );
  const point = selectionToolbarPoint() || { sheetIndex: state.activeSheetIndex, row: targetRange?.top || 1, col: targetRange?.left || 1 };
  const rows = new Set();
  const cols = new Set();
  if (targetRange?.fullRow && !targetRange.fullColumn) {
    const row = Math.max(1, Math.min(sheet.rowCount || 1, targetRange.top || point.row || 1));
    if (row > 1) rows.add(row);
  } else if (targetRange?.fullColumn && !targetRange.fullRow) {
    const col = Math.max(1, Math.min(sheet.colCount || 1, targetRange.left || point.col || 1));
    if (col > 1) cols.add(col);
  } else if (!targetRange?.fullRow && !targetRange?.fullColumn) {
    const row = Math.max(1, Math.min(sheet.rowCount || 1, point.row || targetRange?.top || 1));
    const col = Math.max(1, Math.min(sheet.colCount || 1, point.col || targetRange?.left || 1));
    if (row > 1) rows.add(row);
    if (col > 1) cols.add(col);
  }
  const normalizedRows = [...rows].sort((a, b) => a - b);
  const normalizedCols = [...cols].sort((a, b) => a - b);
  return {
    rows: normalizedRows,
    cols: normalizedCols,
    canInsert: Boolean(normalizedRows.length || normalizedCols.length),
    range: targetRange || null,
  };
}



function pageBreakTargetsForHeaderContext(kind, index) {
  const sheet = activeSheet();
  if (!sheet) return { rows: [], cols: [], canInsert: false, range: null };
  const safeIndex = Math.max(1, Number(index) || 1);
  if (kind === "row") {
    const row = Math.min(sheet.rowCount || safeIndex, safeIndex);
    return { rows: row > 1 ? [row] : [], cols: [], canInsert: row > 1, range: rowSelectionRange(row, row) };
  }
  const col = Math.min(sheet.colCount || safeIndex, safeIndex);
  return { rows: [], cols: col > 1 ? [col] : [], canInsert: col > 1, range: columnSelectionRange(col, col) };
}



function pageBreakTargetsLabel(targets, options = {}) {
  const rowLabels = (targets?.rows || []).map((row) => `${row} 行目の前`);
  const colLabels = (targets?.cols || []).map((col) => `${columnName(col)} 列の前`);
  const labels = [...rowLabels, ...colLabels];
  if (!labels.length) return options.empty || "選択位置";
  if (labels.length <= 2) return labels.join("、");
  return `${labels[0]} ほか ${labels.length - 1} 件`;
}



function selectPageBreakTargets(targets, sheet = activeSheet()) {
  if (!sheet || !targets?.canInsert) return;
  const sheetIndex = state.activeSheetIndex;
  const row = Math.max(1, Math.min(sheet.rowCount || 1, targets.rows?.[0] || targets.range?.top || state.selected?.row || 1));
  const col = Math.max(1, Math.min(sheet.colCount || 1, targets.cols?.[0] || targets.range?.left || state.selected?.col || 1));
  let range = null;
  if (targets.rows?.length && !targets.cols?.length) {
    range = rowSelectionRange(row, row);
  } else if (targets.cols?.length && !targets.rows?.length) {
    range = columnSelectionRange(col, col);
  } else {
    range = selectionRangeFromPoints({ sheetIndex, row, col }, { sheetIndex, row, col });
  }
  const point = selectableCellPoint(range.top, range.left);
  state.selected = point;
  state.selectionRange = { ...range };
  state.selectionRanges = [{ ...range }];
  state.selectionStart = point;
  state.selectionAnchor = point;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  state.formulaEdit = null;
}



function pageBreakTargetsFromAxisIndex(axis, index) {
  const sheet = activeSheet();
  if (!sheet) return { rows: [], cols: [], canInsert: false, range: null };
  const breakIndex = Math.max(2, Number(index) || 2);
  if (axis === "row") {
    const row = Math.min(sheet.rowCount || breakIndex, breakIndex);
    return { rows: [row], cols: [], canInsert: row > 1, range: rowSelectionRange(row, row) };
  }
  const col = Math.min(sheet.colCount || breakIndex, breakIndex);
  return { rows: [], cols: [col], canInsert: col > 1, range: columnSelectionRange(col, col) };
}



function pageBreakTargetListsChanged(beforeRows, beforeCols, afterRows, afterCols) {
  return JSON.stringify(beforeRows || []) !== JSON.stringify(afterRows || []) ||
    JSON.stringify(beforeCols || []) !== JSON.stringify(afterCols || []);
}



function handlePageLayoutMenuAction(event) {
  const action = event.currentTarget.dataset.layoutMenuAction || "";
  closeHomeCommandMenu();
  if (action.startsWith("theme:")) return applyPageTheme(action.slice(6), { includeColors: true, includeFonts: true, includeEffects: true });
  if (action.startsWith("colors:")) return applyPageTheme(action.slice(7), { includeColors: true });
  if (action.startsWith("fonts:")) return applyPageTheme(action.slice(6), { includeFonts: true });
  if (action.startsWith("effects:")) return applyPageTheme(action.slice(8), { includeEffects: true });
  if (action.startsWith("margins:")) return applyPageMarginsPreset(action.slice(8));
  if (action.startsWith("orientation:")) return applyPageOrientation(action.slice(12));
  if (action.startsWith("paper:")) return applyPaperSize(action.slice(6));
  if (action === "print-area:set") return setPrintAreaFromSelection();
  if (action === "print-area:clear") return clearPrintArea();
  if (action === "breaks:insert") return insertPageBreakFromSelection();
  if (action === "breaks:remove") return removePageBreakFromSelection();
  if (action === "breaks:reset") return resetAllPageBreaks();
  if (action === "page-setup") return openPageSetupDialog();
}



function applyPageTheme(themeKey, options = {}) {
  const theme = PAGE_LAYOUT_THEMES.find((item) => item.value === themeKey) || PAGE_LAYOUT_THEMES[0];
  const effect = pageLayoutEffectPreset(themeKey);
  const includeColors = Boolean(options.includeColors);
  const includeFonts = Boolean(options.includeFonts);
  const includeEffects = Boolean(options.includeEffects);
  const appliesFullTheme = includeColors && includeFonts && includeEffects;
  const label = appliesFullTheme ? `テーマ ${theme.label}` : includeColors ? `配色 ${theme.label}` : includeFonts ? `フォント ${theme.label}` : includeEffects ? `効果 ${effect.label}` : `テーマ ${theme.label}`;
  const status = appliesFullTheme
    ? `${theme.label} テーマを適用しました。`
    : includeColors
      ? `配色を「${theme.label}」に変更しました。`
      : includeFonts
        ? `フォントを「${theme.label}」に変更しました。`
        : includeEffects
          ? `効果を「${effect.label}」に変更しました。`
        : `${theme.label} テーマを適用しました。`;
  commitPageLayoutChange((layout) => {
    if (appliesFullTheme || (!includeColors && !includeFonts && !includeEffects)) {
      layout.theme = theme.value;
    }
    if (includeColors) {
      layout.colorScheme = theme.value;
      state.model.themeColors = normalizedThemeColorPalette(theme.colors);
    }
    if (includeFonts) {
      layout.fontScheme = theme.value;
      state.model.themeFonts = { ...theme.fonts };
    }
    if (includeEffects) {
      layout.effectScheme = effect.value;
      state.model.themeEffects = normalizedPageLayoutThemeEffects(effect.value);
      applyPageLayoutEffectToSelectedShapes(effect.value);
    }
  }, label, status);
  syncShapeFormatControls();
}



function applyPageMarginsPreset(key) {
  const margin = PAGE_LAYOUT_MARGIN_PRESETS[key] || PAGE_LAYOUT_MARGIN_PRESETS.normal;
  commitPageLayoutChange((layout) => {
    layout.marginPreset = PAGE_LAYOUT_MARGIN_PRESETS[key] ? key : "normal";
    layout.margins = { ...margin };
  }, `余白 ${margin.label}`, `余白を「${margin.label}」に設定しました。`);
}



function applyPageOrientation(orientation) {
  const next = orientation === "landscape" ? "landscape" : "portrait";
  commitPageLayoutChange((layout) => {
    layout.orientation = next;
  }, "印刷の向き", `印刷の向きを${next === "landscape" ? "横" : "縦"}に設定しました。`);
}



function applyPaperSize(paperSize) {
  const paper = PAGE_LAYOUT_PAPER_SIZES.find((item) => item.value === paperSize) || PAGE_LAYOUT_PAPER_SIZES[0];
  commitPageLayoutChange((layout) => {
    layout.paperSize = paper.value;
  }, `用紙サイズ ${paper.label}`, `用紙サイズを ${paper.label} に設定しました。`);
}



function setPrintAreaFromSelection() {
  const range = activeSelectionRange();
  const materialized = range ? materializeSelectionRange(range) : null;
  if (!materialized) {
    showMessageBox("印刷範囲に設定するセル範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  commitPageLayoutChange((layout) => {
    layout.printArea = { range: { ...materialized }, label: rangeToLabel(materialized) };
  }, "印刷範囲", `${rangeToLabel(materialized)} を印刷範囲に設定しました。`);
}



function clearPrintArea() {
  commitPageLayoutChange((layout) => {
    layout.printArea = null;
  }, "印刷範囲のクリア", "印刷範囲をクリアしました。");
}



function insertPageBreakFromSelection() {
  const sheet = activeSheet();
  const targets = pageBreakTargetsForSelection(sheet);
  if (!targets.canInsert) {
    showMessageBox("先頭行または先頭列には改ページを挿入できません。", { title: "Cht WebSheet" });
    return;
  }
  const layout = ensurePageLayoutSettings(sheet);
  const beforeRows = [...(layout.pageBreaks?.rows || [])];
  const beforeCols = [...(layout.pageBreaks?.cols || [])];
  const nextRows = [...new Set([...beforeRows, ...targets.rows])].sort((a, b) => a - b);
  const nextCols = [...new Set([...beforeCols, ...targets.cols])].sort((a, b) => a - b);
  if (!pageBreakTargetListsChanged(beforeRows, beforeCols, nextRows, nextCols)) {
    setStatus(`${pageBreakTargetsLabel(targets)}には既に手動改ページがあります。`);
    return;
  }
  commitPageLayoutChange((layout) => {
    layout.pageBreaks = layout.pageBreaks || { rows: [], cols: [] };
    layout.pageBreaks.rows = nextRows;
    layout.pageBreaks.cols = nextCols;
  }, "改ページの挿入", `${pageBreakTargetsLabel(targets)}に改ページを挿入しました。`, {
    captureWorkbookView: true,
    afterMutate: (_layout, sheet) => selectPageBreakTargets(targets, sheet),
  });
}



function insertPageBreakAt(axis, index) {
  const normalizedAxis = axis === "row" ? "row" : "col";
  const breakIndex = Math.max(2, Number(index) || 2);
  const layout = ensurePageLayoutSettings(activeSheet());
  const rows = [...(layout.pageBreaks?.rows || [])];
  const cols = [...(layout.pageBreaks?.cols || [])];
  const keyRows = normalizedAxis === "row";
  if ((keyRows ? rows : cols).includes(breakIndex)) {
    setStatus(`${pageBreakPositionLabel(normalizedAxis, breakIndex)}には既に手動改ページがあります。`);
    return;
  }
  commitPageLayoutChange((layout) => {
    layout.pageBreaks = layout.pageBreaks || { rows: [], cols: [] };
    const key = normalizedAxis === "row" ? "rows" : "cols";
    layout.pageBreaks[key] = [...new Set([...(layout.pageBreaks[key] || []), breakIndex])].sort((a, b) => a - b);
  }, "改ページの挿入", `${pageBreakPositionLabel(normalizedAxis, breakIndex)}に改ページを挿入しました。`, {
    captureWorkbookView: true,
    afterMutate: (_layout, sheet) => selectPageBreakTargets(pageBreakTargetsFromAxisIndex(normalizedAxis, breakIndex), sheet),
  });
}



function removePageBreakFromSelection() {
  const sheet = activeSheet();
  const targets = pageBreakTargetsForSelection(sheet);
  const layout = ensurePageLayoutSettings(sheet);
  const beforeRows = [...(layout.pageBreaks?.rows || [])];
  const beforeCols = [...(layout.pageBreaks?.cols || [])];
  const targetRows = new Set(targets.rows);
  const targetCols = new Set(targets.cols);
  const nextRows = beforeRows.filter((row) => !targetRows.has(row));
  const nextCols = beforeCols.filter((col) => !targetCols.has(col));
  if (!pageBreakTargetListsChanged(beforeRows, beforeCols, nextRows, nextCols)) {
    showMessageBox("選択位置に解除できる手動改ページはありません。", { title: "Cht WebSheet" });
    return;
  }
  commitPageLayoutChange((layout) => {
    layout.pageBreaks = layout.pageBreaks || { rows: [], cols: [] };
    layout.pageBreaks.rows = nextRows;
    layout.pageBreaks.cols = nextCols;
  }, "改ページの解除", `${pageBreakTargetsLabel(targets)}の改ページを解除しました。`, {
    captureWorkbookView: true,
    afterMutate: (_layout, sheet) => selectPageBreakTargets(targets, sheet),
  });
}



function removePageBreakAt(axis, index) {
  const normalizedAxis = axis === "row" ? "row" : "col";
  const breakIndex = Math.max(2, Number(index) || 2);
  const layout = ensurePageLayoutSettings(activeSheet());
  const key = normalizedAxis === "row" ? "rows" : "cols";
  if (!(layout.pageBreaks?.[key] || []).includes(breakIndex)) {
    showMessageBox("自動改ページは解除できません。", { title: "Cht WebSheet" });
    return;
  }
  commitPageLayoutChange((layout) => {
    layout.pageBreaks = layout.pageBreaks || { rows: [], cols: [] };
    layout.pageBreaks[key] = (layout.pageBreaks[key] || []).filter((value) => Number(value) !== breakIndex);
  }, "改ページの解除", `${pageBreakPositionLabel(normalizedAxis, breakIndex)}の改ページを解除しました。`, {
    captureWorkbookView: true,
    afterMutate: (_layout, sheet) => selectPageBreakTargets(pageBreakTargetsFromAxisIndex(normalizedAxis, breakIndex), sheet),
  });
}



function resetAllPageBreaks() {
  const layout = ensurePageLayoutSettings(activeSheet());
  if (!(layout.pageBreaks?.rows?.length || layout.pageBreaks?.cols?.length)) {
    setStatus("手動改ページはありません。");
    return;
  }
  commitPageLayoutChange((layout) => {
    layout.pageBreaks = { rows: [], cols: [] };
  }, "改ページのリセット", "すべての改ページを解除しました。", {
    captureWorkbookView: true,
  });
}



function setPageBackgroundFromSelection() {
  const sheet = activeSheet();
  if (!sheet) return;
  const layout = ensurePageLayoutSettings(sheet);
  if (hasPageBackground(layout)) {
    clearPageBackground();
    return;
  }
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートでは背景を変更できません。")) return;
  $pageBackgroundInput.trigger("click");
  setStatus("背景に使用する画像を選択してください。");
}



async function applyPageBackgroundFile(file) {
  if (!file || !state.model) return;
  const sheet = activeSheet();
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートでは背景を変更できません。")) return;
  const src = await fileToDataUrl(file);
  const size = await imageDimensionsFromSource(src);
  applyPageBackground({
    src,
    name: file.name || "背景",
    mimeType: file.type || imageMimeType(fileExtension(file.name || "png")),
    width: size.width,
    height: size.height,
    mode: "tile",
  });
}



function applyPageBackground(background) {
  const normalized = normalizePageLayoutBackground(background);
  if (!normalized) {
    showMessageBox("背景に使用できる画像を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  commitPageLayoutChange((layout) => {
    layout.background = normalized;
  }, "背景", `${normalized.name || "画像"} をシートの背景に設定しました。`);
}



function clearPageBackground() {
  const sheet = activeSheet();
  if (!sheet) return;
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートでは背景を変更できません。")) return;
  commitPageLayoutChange((layout) => {
    layout.background = null;
  }, "背景の削除", "シートの背景を削除しました。");
}



function setPrintTitlesFromSelection() {
  openPageSetupDialog({ activeTab: "sheet", focusField: "printTitleRows" });
}



function handlePageLayoutFieldChange(event) {
  const field = event.currentTarget.dataset.layoutField;
  const value = event.currentTarget.value;
  if (!field) return;
  commitPageLayoutChange((layout) => {
    if (field === "scale") {
      layout.scale = Math.round(Math.max(10, Math.min(400, Number(value) || 100)));
      layout.fitToPage = false;
      layout.fitToWidth = "auto";
      layout.fitToHeight = "auto";
    } else if (field === "fitToWidth" || field === "fitToHeight") {
      layout[field] = value || "auto";
      layout.fitToPage = layout.fitToWidth !== "auto" || layout.fitToHeight !== "auto";
    } else {
      layout[field] = value || "auto";
    }
  }, "拡大縮小印刷", "拡大縮小印刷を更新しました。");
}



function handlePageLayoutSheetOptionChange(event) {
  const option = event.currentTarget.dataset.layoutSheetOption;
  const checked = Boolean(event.currentTarget.checked);
  if (!option) return;
  commitPageLayoutChange((layout, sheet) => {
    if (option === "gridlines-view") sheet.showGridLines = checked;
    if (option === "headings-view") sheet.showHeadings = checked;
    if (option === "gridlines-print") layout.printGridLines = checked;
    if (option === "headings-print") layout.printHeadings = checked;
  }, "シートのオプション", "シートのオプションを更新しました。");
}



function syncPageLayoutRibbonControls() {
  const sheet = state.model ? activeSheet() : null;
  if (!sheet) return;
  const layout = ensurePageLayoutSettings(sheet);
  document.querySelector('[data-layout-field="fitToWidth"]') && (document.querySelector('[data-layout-field="fitToWidth"]').value = layout.fitToWidth);
  document.querySelector('[data-layout-field="fitToHeight"]') && (document.querySelector('[data-layout-field="fitToHeight"]').value = layout.fitToHeight);
  document.querySelector('[data-layout-field="scale"]') && (document.querySelector('[data-layout-field="scale"]').value = String(layout.scale));
  const setChecked = (name, value) => {
    const input = document.querySelector(`[data-layout-sheet-option="${name}"]`);
    if (input) input.checked = Boolean(value);
  };
  setChecked("gridlines-view", sheet.showGridLines !== false);
  setChecked("headings-view", sheetHeadingsVisible(sheet));
  setChecked("gridlines-print", layout.printGridLines);
  setChecked("headings-print", layout.printHeadings);
  const backgroundButton = document.querySelector('[data-ribbon-command="layout-background"]');
  if (backgroundButton) {
    const hasBackground = hasPageBackground(layout);
    const label = hasBackground ? "背景の削除" : "背景";
    const text = backgroundButton.querySelector("span");
    if (text) text.textContent = label;
    backgroundButton.classList.toggle("is-active", hasBackground);
    backgroundButton.setAttribute("title", hasBackground ? "シートの背景を削除" : "シートの背景画像を選択");
    backgroundButton.setAttribute("aria-pressed", String(hasBackground));
  }
}



function openPageSetupDialog(options = {}) {
  const sheet = activeSheet();
  const layout = ensurePageLayoutSettings(sheet);
  const activeTab = ["page", "margins", "header-footer", "sheet"].includes(options.activeTab) ? options.activeTab : "page";
  state.dataDialog = { type: "page-setup" };
  closeDataDialog({ keepState: true });
  const paperOptions = PAGE_LAYOUT_PAPER_SIZES.map((paper) => `<option value="${escapeAttr(paper.value)}"${paper.value === layout.paperSize ? " selected" : ""}>${escapeHtml(paper.label)}</option>`).join("");
  const printTitleRows = pageSetupPrintTitleRowsLabel(layout.printTitles?.rows);
  const printTitleColumns = pageSetupPrintTitleColumnsLabel(layout.printTitles?.cols);
  const tabClass = (tab) => tab === activeTab ? " is-active" : "";
  const panelClass = (tab) => tab === activeTab ? " is-active" : "";
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog page-setup-dialog" role="dialog" aria-modal="true" aria-labelledby="pageSetupTitle" data-data-dialog="page-setup">
      <div class="excel-mini-titlebar" id="pageSetupTitle">ページ設定</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="page-setup-tabs" role="tablist" aria-label="ページ設定">
          <button class="page-setup-tab${tabClass("page")}" type="button" role="tab" data-page-setup-tab="page">ページ</button>
          <button class="page-setup-tab${tabClass("margins")}" type="button" role="tab" data-page-setup-tab="margins">余白</button>
          <button class="page-setup-tab${tabClass("header-footer")}" type="button" role="tab" data-page-setup-tab="header-footer">ヘッダー/フッター</button>
          <button class="page-setup-tab${tabClass("sheet")}" type="button" role="tab" data-page-setup-tab="sheet">シート</button>
        </div>
        <div class="page-setup-panels">
          <section class="page-setup-panel${panelClass("page")}" data-page-setup-panel="page">
            <fieldset class="format-group">
              <legend>印刷の向き</legend>
              <div class="page-setup-radio-row">
                <label><input type="radio" name="pageSetupOrientation" value="portrait" data-page-setup-field="orientation"${layout.orientation === "portrait" ? " checked" : ""} />縦(P)</label>
                <label><input type="radio" name="pageSetupOrientation" value="landscape" data-page-setup-field="orientation"${layout.orientation === "landscape" ? " checked" : ""} />横(L)</label>
              </div>
            </fieldset>
            <div class="format-dialog-grid two">
              <label class="excel-dialog-field"><span class="excel-dialog-caption">拡大/縮小(S):</span><input type="number" min="10" max="400" data-page-setup-field="scale" value="${layout.scale}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">用紙サイズ(Z):</span><select data-page-setup-field="paperSize">${paperOptions}</select></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">横(F):</span><select data-page-setup-field="fitToWidth">${pageSetupFitToOptions(layout.fitToWidth)}</select></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">縦(T):</span><select data-page-setup-field="fitToHeight">${pageSetupFitToOptions(layout.fitToHeight)}</select></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">先頭ページ番号(R):</span><input type="number" min="1" max="9999" data-page-setup-field="firstPageNumber" value="${layout.firstPageNumber}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">ページの方向(O):</span><select data-page-setup-field="pageOrder"><option value="downThenOver"${layout.pageOrder === "downThenOver" ? " selected" : ""}>下から右</option><option value="overThenDown"${layout.pageOrder === "overThenDown" ? " selected" : ""}>右から下</option></select></label>
            </div>
          </section>
          <section class="page-setup-panel${panelClass("margins")}" data-page-setup-panel="margins">
            <div class="format-dialog-grid two">
              <label class="excel-dialog-field"><span class="excel-dialog-caption">上(T):</span><input type="number" step="0.01" data-page-setup-margin="top" value="${formatPageSetupMarginValue(layout.margins.top)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">下(B):</span><input type="number" step="0.01" data-page-setup-margin="bottom" value="${formatPageSetupMarginValue(layout.margins.bottom)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">左(L):</span><input type="number" step="0.01" data-page-setup-margin="left" value="${formatPageSetupMarginValue(layout.margins.left)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">右(R):</span><input type="number" step="0.01" data-page-setup-margin="right" value="${formatPageSetupMarginValue(layout.margins.right)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">ヘッダー(A):</span><input type="number" step="0.01" data-page-setup-margin="header" value="${formatPageSetupMarginValue(layout.margins.header)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">フッター(F):</span><input type="number" step="0.01" data-page-setup-margin="footer" value="${formatPageSetupMarginValue(layout.margins.footer)}" /></label>
            </div>
            <fieldset class="format-group page-setup-center-group">
              <legend>ページ中央</legend>
              <label class="data-dialog-checkbox"><input type="checkbox" data-page-setup-check="centerHorizontal"${layout.centerHorizontal ? " checked" : ""} />水平(H)</label>
              <label class="data-dialog-checkbox"><input type="checkbox" data-page-setup-check="centerVertical"${layout.centerVertical ? " checked" : ""} />垂直(V)</label>
            </fieldset>
          </section>
          <section class="page-setup-panel${panelClass("header-footer")}" data-page-setup-panel="header-footer">
            <div class="format-dialog-grid one">
              <label class="excel-dialog-field"><span class="excel-dialog-caption">ヘッダー(H):</span><input type="text" data-page-setup-field="headerText" value="${escapeAttr(layout.headerFooter?.header || "")}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">フッター(F):</span><input type="text" data-page-setup-field="footerText" value="${escapeAttr(layout.headerFooter?.footer || "")}" /></label>
            </div>
          </section>
          <section class="page-setup-panel${panelClass("sheet")}" data-page-setup-panel="sheet">
            <div class="format-dialog-grid one">
              ${pageSetupRangeFieldHtml("印刷範囲(A):", "printArea", layout.printArea?.label || "", "range")}
              ${pageSetupRangeFieldHtml("タイトル行(R):", "printTitleRows", printTitleRows, "rows")}
              ${pageSetupRangeFieldHtml("タイトル列(C):", "printTitleColumns", printTitleColumns, "columns")}
            </div>
            <label class="data-dialog-checkbox"><input type="checkbox" data-page-setup-check="printGridLines"${layout.printGridLines ? " checked" : ""} />枠線を印刷する(G)</label>
            <label class="data-dialog-checkbox"><input type="checkbox" data-page-setup-check="printHeadings"${layout.printHeadings ? " checked" : ""} />行列番号を印刷する(H)</label>
            <label class="data-dialog-checkbox"><input type="checkbox" data-page-setup-check="blackAndWhite"${layout.blackAndWhite ? " checked" : ""} />白黒印刷(B)</label>
            <label class="data-dialog-checkbox"><input type="checkbox" data-page-setup-check="draftQuality"${layout.draftQuality ? " checked" : ""} />簡易印刷(Q)</label>
            <div class="format-dialog-grid two page-setup-print-options">
              <label class="excel-dialog-field"><span class="excel-dialog-caption">コメント(M):</span><select data-page-setup-field="comments"><option value="none"${layout.comments === "none" ? " selected" : ""}>(なし)</option><option value="atEnd"${layout.comments === "atEnd" ? " selected" : ""}>シートの末尾</option><option value="asDisplayed"${layout.comments === "asDisplayed" ? " selected" : ""}>画面表示イメージ</option></select></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">セルのエラー(E):</span><select data-page-setup-field="cellErrors"><option value="displayed"${layout.cellErrors === "displayed" ? " selected" : ""}>表示する</option><option value="blank"${layout.cellErrors === "blank" ? " selected" : ""}>&lt;空白&gt;</option><option value="dash"${layout.cellErrors === "dash" ? " selected" : ""}>--</option><option value="na"${layout.cellErrors === "na" ? " selected" : ""}>#N/A</option></select></label>
            </div>
          </section>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  const focusSelector = options.focusField ? `[data-page-setup-field='${cssEscape(options.focusField)}']` : "";
  const focusTarget = focusSelector ? overlay.querySelector(focusSelector) : null;
  (focusTarget || overlay.querySelector("[data-data-dialog-action='ok']"))?.focus({ preventScroll: true });
  focusTarget?.select?.();
  if (activeTab === "sheet") setStatus("ページ設定のシート設定を開きました。");
}



function pageSetupRangeFieldHtml(caption, field, value, mode = "range") {
  return `<label class="excel-dialog-field data-range-field page-setup-range-field"><span class="excel-dialog-caption">${accessKeyLabelHtml(caption)}</span><span class="data-range-input"><input type="text" data-page-setup-field="${escapeAttr(field)}" data-data-field="${escapeAttr(field)}" value="${escapeAttr(value || "")}" /><button class="data-range-picker" type="button" title="現在の選択範囲を使用" aria-label="${escapeAttr(caption)}に現在の選択範囲を使用" data-data-dialog-action="range-from-selection" data-data-range-target="${escapeAttr(field)}" data-data-range-mode="${escapeAttr(mode)}"></button></span></label>`;
}



function pageSetupPrintTitleRowsLabel(range) {
  const normalized = normalizePageSetupPrintTitleRows(range);
  return normalized?.label || "";
}



function pageSetupPrintTitleColumnsLabel(range) {
  const normalized = normalizePageSetupPrintTitleColumns(range);
  return normalized?.label || "";
}



function pageSetupFitToOptions(selected) {
  const value = normalizeFitToPageCount(selected);
  const numbers = [...new Set(["1", "2", value].filter((item) => item !== "auto"))]
    .sort((left, right) => Number(left) - Number(right));
  return [
    `<option value="auto"${value === "auto" ? " selected" : ""}>自動</option>`,
    ...numbers.map((item) => `<option value="${escapeAttr(item)}"${value === item ? " selected" : ""}>${escapeHtml(item)} ページ</option>`),
  ].join("");
}



function openPrintDialog() {
  if (!state.model) return;
  const sheet = activeSheet();
  const layout = ensurePageLayoutSettings(sheet);
  state.dataDialog = { type: "print" };
  closeDataDialog({ keepState: true });
  const paperOptions = PAGE_LAYOUT_PAPER_SIZES
    .map((paper) => `<option value="${escapeAttr(paper.value)}"${paper.value === layout.paperSize ? " selected" : ""}>${escapeHtml(paper.label)}</option>`)
    .join("");
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog print-dialog" role="dialog" aria-modal="true" aria-labelledby="printDialogTitle" data-data-dialog="print">
      <div class="excel-mini-titlebar" id="printDialogTitle">印刷</div>
      <div class="excel-mini-body data-dialog-body print-dialog-body">
        <div class="format-dialog-grid two">
          <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("部数(C):")}</span><input type="number" min="1" max="999" value="1" data-data-field="printCopies" /></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("プリンター(P):")}</span><select data-data-field="printPrinter"><option>WebSheet PDF</option><option>既定のプリンター</option></select></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("設定(S):")}</span><select data-data-field="printScope"><option value="active">作業中のシートを印刷</option><option value="workbook">ブック全体を印刷</option><option value="selection">選択した部分を印刷</option></select></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("カラー モード(M):")}</span><select data-data-field="printColorMode"><option value="color"${layout.blackAndWhite ? "" : " selected"}>カラー</option><option value="blackWhite"${layout.blackAndWhite ? " selected" : ""}>白黒</option></select></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("印刷の向き(O):")}</span><select data-data-field="printOrientation"><option value="portrait"${layout.orientation === "portrait" ? " selected" : ""}>縦方向</option><option value="landscape"${layout.orientation === "landscape" ? " selected" : ""}>横方向</option></select></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("用紙サイズ(Z):")}</span><select data-data-field="printPaper">${paperOptions}</select></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("拡大/縮小(S):")}</span><input type="number" min="10" max="400" value="${layout.scale}" data-data-field="printScale" /></label>
        </div>
        <div class="print-dialog-preview" aria-label="印刷プレビュー">
          <div class="print-dialog-preview-toolbar">
            <button class="excel-mini-button print-preview-nav" type="button" data-data-dialog-action="print-preview-prev" aria-label="前のプレビュー">‹</button>
            <span class="print-dialog-preview-label" data-print-preview-label>プレビュー</span>
            <button class="excel-mini-button print-preview-nav" type="button" data-data-dialog-action="print-preview-next" aria-label="次のプレビュー">›</button>
          </div>
          <div class="print-dialog-preview-viewport" data-print-preview-viewport>
            <div class="print-dialog-preview-content" data-print-preview-content></div>
          </div>
          <div class="print-dialog-preview-summary" data-print-preview-summary></div>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="print-page-setup">ページ設定(U)...</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="print-now">印刷(P)</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelectorAll("[data-data-field^='print']").forEach((field) => {
    field.addEventListener("input", () => schedulePrintDialogPreviewSync(overlay, { resetIndex: field.dataset.dataField === "printScope" }));
    field.addEventListener("change", () => schedulePrintDialogPreviewSync(overlay, { resetIndex: field.dataset.dataField === "printScope" }));
  });
  syncPrintDialogPreview(overlay, { resetIndex: true });
  overlay.querySelector("[data-data-field='printCopies']")?.focus({ preventScroll: true });
  overlay.querySelector("[data-data-field='printCopies']")?.select?.();
  setStatus("印刷画面を開きました。");
}



function printDialogPageSettings(dialog) {
  const layout = ensurePageLayoutSettings(activeSheet());
  const paper = dialog?.querySelector("[data-data-field='printPaper']")?.value || layout.paperSize;
  const scale = Number(dialog?.querySelector("[data-data-field='printScale']")?.value);
  return {
    orientation: dialog?.querySelector("[data-data-field='printOrientation']")?.value === "landscape" ? "landscape" : "portrait",
    paperSize: PAGE_LAYOUT_PAPER_SIZES.some((item) => item.value === paper) ? paper : layout.paperSize,
    scale: Math.round(Math.max(10, Math.min(400, Number.isFinite(scale) ? scale : layout.scale))),
    blackAndWhite: dialog?.querySelector("[data-data-field='printColorMode']")?.value === "blackWhite",
  };
}



function printDialogSettingsKey(settings) {
  return JSON.stringify({
    orientation: settings.orientation,
    paperSize: settings.paperSize,
    scale: settings.scale,
    blackAndWhite: settings.blackAndWhite,
  });
}



function applyPrintDialogPageSettings(dialog, options = {}) {
  const sheet = activeSheet();
  if (!sheet || !dialog) return false;
  const current = ensurePageLayoutSettings(sheet);
  const next = printDialogPageSettings(dialog);
  if (printDialogSettingsKey(current) === printDialogSettingsKey(next)) return false;
  commitPageLayoutChange((layout) => {
    layout.orientation = next.orientation;
    layout.paperSize = next.paperSize;
    layout.scale = next.scale;
    layout.blackAndWhite = next.blackAndWhite;
  }, "印刷設定", options.statusMessage || "印刷設定を更新しました。", { render: options.render });
  return true;
}



function syncPrintDialogPreview(dialog) {
  renderPrintDialogPreview(dialog);
}



function schedulePrintDialogPreviewSync(dialog, options = {}) {
  if (!dialog) return;
  if (options.resetIndex) dialog.dataset.printPreviewIndex = "0";
  if (dialog.__printPreviewFrame) cancelAnimationFrame(dialog.__printPreviewFrame);
  dialog.__printPreviewFrame = requestAnimationFrame(() => {
    dialog.__printPreviewFrame = null;
    syncPrintDialogPreview(dialog);
  });
}



function renderPrintDialogPreview(dialog) {
  const content = dialog?.querySelector?.("[data-print-preview-content]");
  const label = dialog?.querySelector?.("[data-print-preview-label]");
  const summary = dialog?.querySelector?.("[data-print-preview-summary]");
  if (!content || !label || !summary) return;
  const pages = printPreviewPagesForDialog(dialog);
  if (!pages.length) {
    content.innerHTML = `<div class="print-dialog-preview-empty">印刷する範囲がありません。</div>`;
    label.textContent = "プレビュー";
    summary.textContent = "";
    syncPrintDialogPreviewNavigation(dialog, 0, 0);
    return;
  }
  const rawIndex = Math.round(Number(dialog.dataset.printPreviewIndex) || 0);
  const index = Math.max(0, Math.min(pages.length - 1, rawIndex));
  dialog.dataset.printPreviewIndex = String(index);
  const page = pages[index];
  const paper = printDialogPaperPixels(page.layout);
  const margins = printDialogMarginPixels(page.layout);
  const printableWidth = Math.max(24, paper.width - margins.left - margins.right);
  const printableHeight = Math.max(24, paper.height - margins.top - margins.bottom);
  const rangeLabel = rangeToLabel(page.range);
  content.innerHTML = `
    <div class="print-dialog-preview-paper-shell" data-paper-width="${roundCssNumber(paper.width)}" data-paper-height="${roundCssNumber(paper.height)}">
      <div class="print-dialog-preview-paper" style="width:${roundCssNumber(paper.width)}px;height:${roundCssNumber(paper.height)}px">
        <div class="print-dialog-preview-printable" style="left:${roundCssNumber(margins.left)}px;top:${roundCssNumber(margins.top)}px;width:${roundCssNumber(printableWidth)}px;height:${roundCssNumber(printableHeight)}px">
          ${printSheetHtml(page.sheet, page.range, page.layout, index)}
        </div>
      </div>
    </div>`;
  syncPrintObjectLayerPositions(content);
  label.textContent = pages.length > 1 ? `${index + 1} / ${pages.length}` : "1 / 1";
  const printRangeLabel = rangeToLabel(page.printRange || page.range);
  const rangeSummary = printRangeLabel === rangeLabel
    ? `印刷範囲 ${printRangeLabel}`
    : `印刷範囲 ${printRangeLabel} / 表示 ${rangeLabel}`;
  summary.textContent = `${page.sheet.name || "Sheet"}  ${rangeSummary}  ${page.pageNumber} ページ / ${page.totalPages}  ${page.layout.orientation === "landscape" ? "横" : "縦"} / ${page.layout.paperSize} / ${page.layout.scale}%`;
  syncPrintDialogPreviewNavigation(dialog, index, pages.length);
  requestAnimationFrame(() => {
    syncPrintObjectLayerPositions(content);
    fitPrintDialogPreviewPaper(dialog);
  });
}



function syncPrintDialogPreviewNavigation(dialog, index, count) {
  const prev = dialog?.querySelector?.("[data-data-dialog-action='print-preview-prev']");
  const next = dialog?.querySelector?.("[data-data-dialog-action='print-preview-next']");
  if (prev) prev.disabled = count <= 1 || index <= 0;
  if (next) next.disabled = count <= 1 || index >= count - 1;
}



function fitPrintDialogPreviewPaper(dialog) {
  const viewport = dialog?.querySelector?.("[data-print-preview-viewport]");
  const shell = dialog?.querySelector?.(".print-dialog-preview-paper-shell");
  const paper = dialog?.querySelector?.(".print-dialog-preview-paper");
  if (!viewport || !shell || !paper) return;
  const rawWidth = Number(shell.dataset.paperWidth) || paper.offsetWidth || 1;
  const rawHeight = Number(shell.dataset.paperHeight) || paper.offsetHeight || 1;
  const availableWidth = Math.max(80, viewport.clientWidth - 28);
  const scale = Math.max(0.08, Math.min(1, availableWidth / rawWidth));
  paper.style.transform = `scale(${roundCssNumber(scale)})`;
  shell.style.width = `${roundCssNumber(rawWidth * scale)}px`;
  shell.style.height = `${roundCssNumber(rawHeight * scale)}px`;
}



function movePrintDialogPreview(dialog, delta) {
  if (!dialog) return;
  const count = printPreviewPagesForDialog(dialog).length;
  if (!count) return;
  const current = Math.max(0, Math.min(count - 1, Math.round(Number(dialog.dataset.printPreviewIndex) || 0)));
  dialog.dataset.printPreviewIndex = String(Math.max(0, Math.min(count - 1, current + delta)));
  syncPrintDialogPreview(dialog);
}



function printDialogPaperPixels(layout) {
  const paper = PAGE_LAYOUT_PAPER_SIZES.find((item) => item.value === layout.paperSize) || PAGE_LAYOUT_PAPER_SIZES[0];
  const landscape = layout.orientation === "landscape";
  return {
    width: (landscape ? paper.height : paper.width) * 96,
    height: (landscape ? paper.width : paper.height) * 96,
  };
}



function printDialogMarginPixels(layout) {
  const margins = layout.margins || PAGE_LAYOUT_MARGIN_PRESETS.normal;
  return {
    top: Math.max(0, Number(margins.top) || 0) * 96,
    right: Math.max(0, Number(margins.right) || 0) * 96,
    bottom: Math.max(0, Number(margins.bottom) || 0) * 96,
    left: Math.max(0, Number(margins.left) || 0) * 96,
  };
}



function printPreviewPagesForDialog(dialog) {
  return printTargetsForDialog(dialog)
    .flatMap((target, sheetOrder) => printPreviewPagesForTarget(target, sheetOrder))
    .sort((left, right) => left.sheetOrder - right.sheetOrder || left.pageNumber - right.pageNumber);
}



function printPreviewPagesForTarget(target, sheetOrder = 0) {
  if (!target?.sheet || !target.range) return [];
  const plan = pageBreakPreviewPlan(target.sheet, target.layout, target.range);
  const rows = plan.rows?.segments || [];
  const cols = plan.cols?.segments || [];
  const totalRows = Math.max(1, rows.length);
  const totalCols = Math.max(1, cols.length);
  const pages = [];
  rows.forEach((rowSegment, rowIndex) => {
    cols.forEach((colSegment, colIndex) => {
      const pageNumber = pageBreakPreviewPageNumber(rowIndex + 1, colIndex + 1, totalRows, totalCols, target.layout);
      pages.push({
        sheet: target.sheet,
        sheetOrder,
        layout: target.layout,
        pageNumber,
        totalPages: totalRows * totalCols,
        printRange: { ...target.range },
        range: normalizePrintRange(target.sheet, {
          top: rowSegment.start,
          bottom: rowSegment.end,
          left: colSegment.start,
          right: colSegment.end,
        }),
      });
    });
  });
  return pages.filter((page) => page.range).sort((left, right) => left.pageNumber - right.pageNumber);
}



function printWorkbookFromDialog(dialog) {
  if (!state.model || !dialog) return;
  applyPrintDialogPageSettings(dialog, { render: false });
  if (!prepareWorkbookPrintSurface(dialog)) return;
  closeDataDialog();
  setStatus("印刷を開始しました。");
  requestAnimationFrame(() => {
    try {
      window.print?.();
    } catch (error) {
      console.error("Print failed.", error);
      cleanupWorkbookPrintSurface();
      showMessageBox("印刷を開始できませんでした。ブラウザーの印刷機能を確認してください。", { title: "Cht WebSheet" });
    }
  });
}



function prepareWorkbookPrintSurface(dialog) {
  const targets = printTargetsForDialog(dialog);
  const pages = printPagesForTargets(targets);
  if (!pages.length) {
    showMessageBox("印刷する範囲がありません。", { title: "Cht WebSheet" });
    return false;
  }
  cleanupWorkbookPrintSurface();
  const surface = document.createElement("div");
  surface.id = "workbookPrintSurface";
  surface.setAttribute("aria-hidden", "true");
  surface.innerHTML = pages.map((page, index) => printWorkbookPageHtml(page, index)).join("");
  document.body.appendChild(surface);
  syncPrintObjectLayerPositions(surface);

  const pageStyle = document.createElement("style");
  pageStyle.id = "workbookPrintStyle";
  pageStyle.textContent = workbookPrintPageStyle(pages[0].layout);
  document.head.appendChild(pageStyle);

  document.body.classList.add("is-printing-workbook");
  ensureWorkbookPrintAfterPrintListener();
  scheduleWorkbookPrintSurfaceCleanup();
  requestAnimationFrame(() => syncPrintObjectLayerPositions(surface));
  return true;
}



function printPagesForTargets(targets = []) {
  return (targets || [])
    .flatMap((target, sheetOrder) => printPreviewPagesForTarget(target, sheetOrder))
    .sort((left, right) => left.sheetOrder - right.sheetOrder || left.pageNumber - right.pageNumber);
}



function ensureWorkbookPrintAfterPrintListener() {
  if (ensureWorkbookPrintAfterPrintListener.bound) return;
  window.addEventListener("afterprint", cleanupWorkbookPrintSurface);
  ensureWorkbookPrintAfterPrintListener.bound = true;
}



function scheduleWorkbookPrintSurfaceCleanup() {
  if (state.printSurfaceCleanupTimer) clearTimeout(state.printSurfaceCleanupTimer);
  state.printSurfaceCleanupTimer = window.setTimeout(() => cleanupWorkbookPrintSurface(), 120000);
}



function cleanupWorkbookPrintSurface() {
  if (state.printSurfaceCleanupTimer) {
    clearTimeout(state.printSurfaceCleanupTimer);
    state.printSurfaceCleanupTimer = null;
  }
  document.body.classList.remove("is-printing-workbook");
  document.getElementById("workbookPrintSurface")?.remove();
  document.getElementById("workbookPrintStyle")?.remove();
}



function printTargetsForDialog(dialog) {
  const scope = dialog?.querySelector("[data-data-field='printScope']")?.value || "active";
  if (scope === "workbook") {
    return visibleSheetIndexes()
      .map((sheetIndex) => {
        const sheet = state.model?.sheets?.[sheetIndex];
        return sheet ? printTargetForSheet(sheet, dialog) : null;
      })
      .filter(Boolean);
  }
  const sheet = activeSheet();
  if (!sheet) return [];
  if (scope === "selection") {
    const range = printRangeForSelection(sheet);
    return range ? [{ sheet, range, layout: printLayoutForDialog(sheet, dialog) }] : [];
  }
  return selectedSheetIndexesForPrint()
    .map((sheetIndex) => {
      const targetSheet = state.model?.sheets?.[sheetIndex];
      return targetSheet ? printTargetForSheet(targetSheet, dialog) : null;
    })
    .filter(Boolean);
}



function selectedSheetIndexesForPrint() {
  if (!state.model?.sheets?.length) return [];
  const visible = new Set(visibleSheetIndexes());
  const selected = [...new Set(state.selectedSheetIndexes || [])]
    .filter((sheetIndex) => visible.has(sheetIndex) && state.model.sheets[sheetIndex]);
  if (selected.length > 1) return selected;
  return visible.has(state.activeSheetIndex) && state.model.sheets[state.activeSheetIndex]
    ? [state.activeSheetIndex]
    : [];
}



function printTargetForSheet(sheet, dialog = null) {
  const layout = printLayoutForDialog(sheet, dialog);
  const range = printRangeForSheet(sheet, layout);
  return range ? { sheet, range, layout: normalizedPageLayoutSettings(layout, sheet) } : null;
}



function printLayoutForDialog(sheet, dialog = null) {
  const layout = normalizedPageLayoutSettings(sheet?.pageLayout, sheet);
  if (!dialog) return layout;
  return {
    ...layout,
    ...printDialogPageSettings(dialog),
  };
}



function printRangeForSheet(sheet, layout = ensurePageLayoutSettings(sheet)) {
  return normalizePrintRange(sheet, layout.printArea?.range || sheetUsedRangeForPrint(sheet) || { top: 1, left: 1, bottom: 1, right: 1 });
}



function printRangeForSelection(sheet) {
  const ranges = activeSelectionRanges();
  const materialized = unionMaterializedRanges(materializedRangesForSheet(sheet, ranges, state.activeSheetIndex), state.activeSheetIndex);
  if (!materialized) return printRangeForSheet(sheet);
  const used = sheetUsedRangeForPrint(sheet);
  const range = { ...materialized };
  if (used) {
    if (range.fullRow) {
      range.left = used.left;
      range.right = used.right;
    }
    if (range.fullColumn) {
      range.top = used.top;
      range.bottom = used.bottom;
    }
  }
  return normalizePrintRange(sheet, range);
}



function sheetUsedRangeForPrint(sheet) {
  const ranges = [];
  const used = sheetUsedRange(sheet);
  if (used) ranges.push(used);
  if (sheet?.autoFilter?.range) ranges.push(sheet.autoFilter.range);
  (sheet?.images || []).forEach((image) => {
    if (image?.hidden) return;
    const range = imageRangeForPrint(sheet, image);
    if (range) ranges.push(range);
  });
  if (!ranges.length) return null;
  return {
    top: Math.min(...ranges.map((range) => range.top)),
    left: Math.min(...ranges.map((range) => range.left)),
    bottom: Math.max(...ranges.map((range) => range.bottom)),
    right: Math.max(...ranges.map((range) => range.right)),
  };
}



function imageRangeForPrint(sheet, image) {
  const box = imageBoxToPixels(sheet, image);
  if (!box) return null;
  const top = sheetIndexFromPixel(sheet, "row", box.top) || 1;
  const left = sheetIndexFromPixel(sheet, "col", box.left) || 1;
  const bottom = sheetIndexFromPixel(sheet, "row", box.top + box.height) || top;
  const right = sheetIndexFromPixel(sheet, "col", box.left + box.width) || left;
  return normalizePrintRange(sheet, {
    top: Math.min(top, bottom),
    left: Math.min(left, right),
    bottom: Math.max(top, bottom),
    right: Math.max(left, right),
  });
}



function normalizePrintRange(sheet, range) {
  if (!sheet || !range) return null;
  const rowCount = Math.max(1, Number(sheet.rowCount || sheet.rows?.length || 1));
  const colCount = Math.max(1, Number(sheet.colCount || sheet.columns?.length || 1));
  const top = Math.max(1, Math.min(rowCount, Math.round(Number(range.top) || 1)));
  const left = Math.max(1, Math.min(colCount, Math.round(Number(range.left) || 1)));
  const bottom = Math.max(top, Math.min(rowCount, Math.round(Number(range.bottom) || top)));
  const right = Math.max(left, Math.min(colCount, Math.round(Number(range.right) || left)));
  return { sheetIndex: state.model?.sheets?.indexOf(sheet) ?? state.activeSheetIndex, top, left, bottom, right };
}



function workbookPrintPageStyle(layout) {
  const paper = PAGE_LAYOUT_PAPER_SIZES.find((item) => item.value === layout.paperSize) || PAGE_LAYOUT_PAPER_SIZES[0];
  const size = `${paper.value} ${layout.orientation === "landscape" ? "landscape" : "portrait"}`;
  return `@page{size:${size};margin:0;}`;
}



function printWorkbookPageHtml(page, index = 0) {
  const paper = printDialogPaperPixels(page.layout);
  const margins = printDialogMarginPixels(page.layout);
  const printableWidth = Math.max(24, paper.width - margins.left - margins.right);
  const printableHeight = Math.max(24, paper.height - margins.top - margins.bottom);
  return `
    <section class="workbook-print-page" data-print-page="${index + 1}" style="--print-paper-width:${roundCssNumber(paper.width)}px;--print-paper-height:${roundCssNumber(paper.height)}px;width:${roundCssNumber(paper.width)}px;height:${roundCssNumber(paper.height)}px">
      <div class="workbook-print-page-printable" style="left:${roundCssNumber(margins.left)}px;top:${roundCssNumber(margins.top)}px;width:${roundCssNumber(printableWidth)}px;height:${roundCssNumber(printableHeight)}px">
        ${printSheetHtml(page.sheet, page.range, page.layout, index, { materializeScale: true })}
      </div>
    </section>`;
}



function printSheetHtml(sheet, range, layout, index = 0, options = {}) {
  return withDisplayCalculationCache(sheet, () => {
    const printRows = printRowIndexes(sheet, range);
    const printCols = printColumnIndexes(sheet, range);
    const rawWidth = printCols.reduce((total, col) => total + printColumnWidth(sheet, col), layout.printHeadings ? rowHeaderWidthForSheet(sheet) : 0);
    const rawHeight = printRows.reduce((total, row) => total + printRowHeight(sheet, row), layout.printHeadings ? COLUMN_HEADER_HEIGHT : 0);
    const scale = printSheetEffectiveScale(sheet, range, layout);
    const renderScale = options.materializeScale ? scale : 1;
    const transformScale = options.materializeScale ? 1 : scale;
    const width = Math.max(1, rawWidth * renderScale);
    const height = Math.max(1, rawHeight * renderScale);
    const scaledWidth = Math.max(1, rawWidth * scale);
    const scaledHeight = Math.max(1, rawHeight * scale);
    const classes = [
      "workbook-print-sheet",
      layout.blackAndWhite ? "is-black-white" : "",
      index === 0 ? "is-first" : "",
    ].filter(Boolean).join(" ");
    return `
      <section class="${classes}" data-print-sheet="${escapeAttr(sheet.name || "Sheet")}" style="--print-scale:${roundCssNumber(transformScale)};--print-render-scale:${roundCssNumber(renderScale)};--print-effective-scale:${roundCssNumber(scale)};--print-sheet-width:${roundCssNumber(scaledWidth)}px;--print-sheet-height:${roundCssNumber(scaledHeight)}px;width:${roundCssNumber(scaledWidth)}px;min-height:${roundCssNumber(scaledHeight)}px">
        <div class="workbook-print-sheet-grid" data-print-range-top="${range.top}" data-print-range-left="${range.left}" data-print-range-bottom="${range.bottom}" data-print-range-right="${range.right}" style="width:${roundCssNumber(width)}px;height:${roundCssNumber(height)}px">
          ${printSheetTableHtml(sheet, range, layout, printRows, printCols, renderScale)}
          ${printSheetObjectsHtml(sheet, range, layout, renderScale)}
        </div>
      </section>`;
  });
}



function printSheetEffectiveScale(sheet, printRange, layout) {
  const baseScale = Math.max(0.1, Math.min(4, Number(layout.scale || 100) / 100));
  const pageSize = printRangeContentSize(sheet, printRange, layout);
  const printable = printLayoutPrintablePixels(layout);
  let scale = baseScale;
  const previewPadding = 0.94;
  const { fitToWidth, fitToHeight, active } = pageLayoutFitTargets(layout);
  if (fitToWidth > 0 || (!active && pageSize.width * scale > printable.width * previewPadding)) {
    scale = Math.min(scale, (printable.width * previewPadding) / Math.max(1, pageSize.width));
  }
  if (fitToHeight > 0 || (!active && pageSize.height * scale > printable.height * previewPadding)) {
    scale = Math.min(scale, (printable.height * previewPadding) / Math.max(1, pageSize.height));
  }
  return Math.max(0.1, Math.min(4, scale));
}



function pageLayoutFitTargets(layout = {}) {
  if (layout?.fitToPage !== true) return { fitToWidth: 0, fitToHeight: 0, active: false };
  const fitToWidth = layout.fitToWidth !== "auto" ? Math.max(1, Number(layout.fitToWidth) || 0) : 0;
  const fitToHeight = layout.fitToHeight !== "auto" ? Math.max(1, Number(layout.fitToHeight) || 0) : 0;
  return { fitToWidth, fitToHeight, active: fitToWidth > 0 || fitToHeight > 0 };
}



function printRangeContentSize(sheet, range, layout) {
  const normalized = normalizePrintRange(sheet, range);
  if (!normalized) return { width: 1, height: 1 };
  const cols = printColumnIndexes(sheet, normalized);
  const rows = printRowIndexes(sheet, normalized);
  return {
    width: cols.reduce((total, col) => total + printColumnWidth(sheet, col), layout.printHeadings ? rowHeaderWidthForSheet(sheet) : 0),
    height: rows.reduce((total, row) => total + printRowHeight(sheet, row), layout.printHeadings ? COLUMN_HEADER_HEIGHT : 0),
  };
}



function printLayoutPrintablePixels(layout) {
  const paper = PAGE_LAYOUT_PAPER_SIZES.find((item) => item.value === layout.paperSize) || PAGE_LAYOUT_PAPER_SIZES[0];
  const landscape = layout.orientation === "landscape";
  const paperWidth = landscape ? paper.height : paper.width;
  const paperHeight = landscape ? paper.width : paper.height;
  const margins = layout.margins || PAGE_LAYOUT_MARGIN_PRESETS.normal;
  return {
    width: Math.max(24, (paperWidth - Number(margins.left || 0) - Number(margins.right || 0)) * 96),
    height: Math.max(24, (paperHeight - Number(margins.top || 0) - Number(margins.bottom || 0)) * 96),
  };
}



function printRowIndexes(sheet, range) {
  const rows = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    if (sheet.rows?.[row - 1]?.hidden) continue;
    rows.push(row);
  }
  return rows.length ? rows : [range.top];
}



function printColumnIndexes(sheet, range) {
  const cols = [];
  for (let col = range.left; col <= range.right; col += 1) {
    if (sheet.columns?.[col - 1]?.hidden) continue;
    cols.push(col);
  }
  return cols.length ? cols : [range.left];
}



function printRowHeight(sheet, row) {
  return Math.max(1, Number(sheet.rows?.[row - 1]?.height || sheet.defaultRowHeight || rowHeightToPixels(15)));
}



function printColumnWidth(sheet, col) {
  return Math.max(1, Number(sheet.columns?.[col - 1]?.width || sheet.defaultColumnWidth || columnWidthToPixels(8.43)));
}



function printRenderScale(value) {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : 1;
}



function printSheetTableHtml(sheet, range, layout, printRows, printCols, renderScale = 1) {
  const sheetId = getSheetId(sheet.name);
  const mergeMaps = buildPrintMergeMaps(sheet, range, printRows, printCols);
  const scaleValue = printRenderScale(renderScale);
  const colgroup = [
    layout.printHeadings ? `<col class="workbook-print-row-heading-col" data-print-col-heading="true" data-print-col-width="${roundCssNumber(rowHeaderWidthForSheet(sheet) * scaleValue)}" style="width:${roundCssNumber(rowHeaderWidthForSheet(sheet) * scaleValue)}px">` : "",
    ...printCols.map((col) => `<col data-print-col="${col}" data-print-col-width="${roundCssNumber(printColumnWidth(sheet, col) * scaleValue)}" style="width:${roundCssNumber(printColumnWidth(sheet, col) * scaleValue)}px">`),
  ].filter(Boolean).join("");
  const rows = [];
  if (layout.printHeadings) {
    const headingHeight = COLUMN_HEADER_HEIGHT * scaleValue;
    rows.push(`<tr class="workbook-print-column-heading-row" data-print-row-heading="true" data-print-row-height="${roundCssNumber(headingHeight)}" style="height:${roundCssNumber(headingHeight)}px"><th class="workbook-print-corner-heading"></th>${printCols.map((col) => `<th class="workbook-print-column-heading">${escapeHtml(columnName(col))}</th>`).join("")}</tr>`);
  }
  printRows.forEach((row) => {
    const cells = [];
    if (layout.printHeadings) cells.push(`<th class="workbook-print-row-heading">${row}</th>`);
    printCols.forEach((col) => {
      const cellHtml = printCellHtml(sheet, mergeMaps, sheetId, row, col, range, printRows, printCols, layout, scaleValue);
      if (cellHtml) cells.push(cellHtml);
    });
    const rowHeight = printRowHeight(sheet, row) * scaleValue;
    rows.push(`<tr data-print-row="${row}" data-print-row-height="${roundCssNumber(rowHeight)}" style="height:${roundCssNumber(rowHeight)}px">${cells.join("")}</tr>`);
  });
  const classes = `workbook-print-table${layout.printGridLines ? " has-gridlines" : ""}`;
  return `<table class="${classes}"><colgroup>${colgroup}</colgroup><tbody>${rows.join("")}</tbody></table>`;
}



function printCellHtml(sheet, mergeMaps, sheetId, row, col, range, printRows, printCols, layout, renderScale = 1) {
  const key = cellKey(row, col);
  if (mergeMaps.covered.has(key)) return "";
  const merge = mergeMaps.master.get(key);
  const rowSpan = merge ? countPrintIndexesInRange(printRows, row, Math.min(merge.bottom, range.bottom)) : 1;
  const colSpan = merge ? countPrintIndexesInRange(printCols, col, Math.min(merge.right, range.right)) : 1;
  if (rowSpan <= 0 || colSpan <= 0) return "";
  const sourceRow = merge?.sourceTop || row;
  const sourceCol = merge?.sourceLeft || col;
  const sourceKey = cellKey(sourceRow, sourceCol);
  const sourceCell = sheet.cells?.[sourceKey] || {};
  const cell = merge ? printCellForMergeSlice(sourceCell, merge) : sourceCell;
  const scaleValue = printRenderScale(renderScale);
  const cellHeight = printCellSpanHeight(sheet, printRows, row, rowSpan) * scaleValue;
  const showMergedContent = !merge || (sourceRow === row && sourceCol === col);
  const display = showMergedContent
    ? (state.showFormulas && cell.formula ? cell.formula : getDisplayForCell(sheet, sheetId, sourceRow, sourceCol))
    : "";
  const overflowWidth = printCellOverflowWidth(sheet, mergeMaps, row, col, colSpan, cell, display, range, layout, scaleValue);
  const suppressGridSides = merge ? printMergeSliceSuppressedSides(merge) : null;
  const attrs = [
    overflowWidth ? `class="workbook-print-cell-overflow"` : "",
    rowSpan > 1 ? `rowspan="${rowSpan}"` : "",
    colSpan > 1 ? `colspan="${colSpan}"` : "",
    `style="${escapeAttr(printCellStyle(sheet, cell, cellHeight, layout, scaleValue, suppressGridSides))}"`,
  ].filter(Boolean).join(" ");
  return `<td ${attrs}><div class="workbook-print-cell-content" style="${escapeAttr(printCellContentStyle(sheet, cell, cellHeight, overflowWidth))}"><span class="workbook-print-cell-text">${printCellContentHtml(display)}</span></div></td>`;
}



function buildPrintMergeMaps(sheet, range, printRows, printCols) {
  const master = new Map();
  const covered = new Set();
  const coveredToMaster = new Map();
  const rowSet = new Set(printRows || []);
  const colSet = new Set(printCols || []);
  (sheet?.merges || []).forEach((merge) => {
    if (!merge || !rangesIntersect(merge, range)) return;
    const visibleRows = [];
    for (let row = Math.max(merge.top, range.top); row <= Math.min(merge.bottom, range.bottom); row += 1) {
      if (rowSet.has(row)) visibleRows.push(row);
    }
    const visibleCols = [];
    for (let col = Math.max(merge.left, range.left); col <= Math.min(merge.right, range.right); col += 1) {
      if (colSet.has(col)) visibleCols.push(col);
    }
    if (!visibleRows.length || !visibleCols.length) return;
    const visible = {
      top: visibleRows[0],
      left: visibleCols[0],
      bottom: visibleRows[visibleRows.length - 1],
      right: visibleCols[visibleCols.length - 1],
      sourceTop: merge.top,
      sourceLeft: merge.left,
      sourceBottom: merge.bottom,
      sourceRight: merge.right,
    };
    const masterKey = cellKey(visible.top, visible.left);
    master.set(masterKey, visible);
    visibleRows.forEach((row) => {
      visibleCols.forEach((col) => {
        const key = cellKey(row, col);
        if (key === masterKey) return;
        covered.add(key);
        coveredToMaster.set(key, masterKey);
      });
    });
  });
  return { master, covered, coveredToMaster };
}



function printCellForMergeSlice(cell, merge) {
  if (!merge || !cell) return cell || {};
  const borders = { ...(cell.borders || {}) };
  if (merge.top > merge.sourceTop) delete borders.top;
  if (merge.bottom < merge.sourceBottom) delete borders.bottom;
  if (merge.left > merge.sourceLeft) delete borders.left;
  if (merge.right < merge.sourceRight) delete borders.right;
  return { ...cell, borders };
}



function printMergeSliceSuppressedSides(merge) {
  const sides = new Set();
  if (!merge) return sides;
  if (merge.top > merge.sourceTop) sides.add("top");
  if (merge.bottom < merge.sourceBottom) sides.add("bottom");
  if (merge.left > merge.sourceLeft) sides.add("left");
  if (merge.right < merge.sourceRight) sides.add("right");
  return sides;
}



function countPrintIndexesInRange(indexes, start, end) {
  return indexes.filter((index) => index >= start && index <= end).length;
}



function printCellSpanHeight(sheet, printRows, row, rowSpan) {
  const startIndex = printRows.indexOf(row);
  const rows = startIndex >= 0 ? printRows.slice(startIndex, startIndex + Math.max(1, rowSpan)) : [row];
  return Math.max(1, rows.reduce((total, rowIndex) => total + printRowHeight(sheet, rowIndex), 0));
}



function printCellSpanWidth(sheet, col, colSpan) {
  let width = 0;
  for (let colIndex = col; colIndex < col + Math.max(1, colSpan); colIndex += 1) {
    width += printColumnWidth(sheet, colIndex);
  }
  return Math.max(1, width);
}



function printCellOverflowWidth(sheet, mergeMaps, row, col, colSpan, cell, display, range, layout = {}, renderScale = 1) {
  const scaleValue = printRenderScale(renderScale);
  const overflowWidth = textOverflowWidth(sheet, mergeMaps, row, col, colSpan, cell, display);
  if (!overflowWidth) return null;
  const currentWidth = printCellSpanWidth(sheet, col, colSpan);
  if (!layout?.printArea?.range) return Math.max(currentWidth, overflowWidth) * scaleValue;
  const rangeRight = Math.max(col + colSpan - 1, Math.min(sheet.colCount, Number(range?.right) || col + colSpan - 1));
  const pageWidth = widthBetweenColumns(sheet, col, rangeRight);
  return Math.max(currentWidth, Math.min(overflowWidth, pageWidth || currentWidth)) * scaleValue;
}



function printCellStyle(sheet, cell, rowHeight, layout, renderScale = 1, suppressGridSides = null) {
  const scaleValue = printRenderScale(renderScale);
  const style = legacyExcelCellStyle(sheet, cell, rowHeight, scaleValue);
  if (!layout.printGridLines) return style;
  const borders = cell?.borders || {};
  const extra = {};
  CELL_BORDER_SIDES.forEach((side) => {
    if (suppressGridSides?.has?.(side)) return;
    if (borders[side]) return;
    const cssSide = side === "top" ? "borderTop" : side === "right" ? "borderRight" : side === "bottom" ? "borderBottom" : "borderLeft";
    extra[cssSide] = `${roundCssNumber(Math.max(0.5, scaleValue))}px solid #d9e2ec`;
  });
  return [style, styleObjectToString(extra)].filter(Boolean).join(";");
}



function printCellContentStyle(sheet, cell, height, overflowWidth = null) {
  const css = cellEffectiveCss(sheet, cell);
  const justifyContent = css.verticalAlign === "top"
    ? "flex-start"
    : css.verticalAlign === "bottom"
      ? "flex-end"
      : "center";
  const width = overflowWidth ? `${Math.max(1, Math.round(Number(overflowWidth) || 1))}px` : "100%";
  return styleObjectToString({
    height: `${Math.max(1, Math.round(Number(height) || 1))}px`,
    maxHeight: `${Math.max(1, Math.round(Number(height) || 1))}px`,
    justifyContent,
    lineHeight: "1.2",
    whiteSpace: css.whiteSpace === "pre-wrap" ? "pre-wrap" : "nowrap",
    width,
    maxWidth: overflowWidth ? "none" : "100%",
  });
}



function printCellContentHtml(display) {
  return escapeHtml(display ?? "").replace(/\r\n|\r|\n/g, "<br>");
}

export {
  applyMacroEditorDialog,
  applyMacroOptionsDialog,
  applyMacroSecurityDialog,
  applyPageBackground,
  applyPageBackgroundFile,
  applyPageLayoutEffectToSelectedShapes,
  applyPageMarginsPreset,
  applyPageOrientation,
  applyPageTheme,
  applyPaperSize,
  applyPrintDialogPageSettings,
  applySheetCodeDialog,
  buildPrintMergeMaps,
  cleanupWorkbookPrintSurface,
  clearPageBackground,
  clearPrintArea,
  clearVbaRangeTarget,
  commitPageLayoutChange,
  compareVbaExpressionValues,
  convertVbaR1C1FormulaToA1,
  countPrintIndexesInRange,
  createMacroFromDialog,
  defaultShapeThemeEffectPatchForType,
  deleteMacroFromDialog,
  ensurePageLayoutSettings,
  ensureWorkbookPrintAfterPrintListener,
  evaluateVbaBuiltInFunctionExpression,
  evaluateVbaWorksheetFunctionExpression,
  excelImportedPageMargins,
  executeMacroStepDialog,
  fitPrintDialogPreviewPaper,
  formatPageSetupMarginValue,
  handlePageLayoutFieldChange,
  handlePageLayoutMenuAction,
  handlePageLayoutSheetOptionChange,
  handleSheetCodeProcedureChange,
  hasPageBackground,
  imageRangeForPrint,
  insertPageBreakAt,
  insertPageBreakFromSelection,
  isVbaNoneValue,
  isVbaUnaryArithmeticOperator,
  isVbaWordBoundary,
  layoutMenuButtonHtml,
  macroSecurityOptionRows,
  macroStepListHtml,
  macroStepLogHtml,
  movePrintDialogPreview,
  normalizeFitToPageCount,
  normalizePageLayoutBackground,
  normalizePageSetupPrintTitleColumns,
  normalizePageSetupPrintTitleRows,
  normalizePrintRange,
  normalizedPageLayoutEffectShapePatch,
  normalizedPageLayoutSettings,
  normalizedPageLayoutThemeEffects,
  openMacroEditorDialog,
  openMacroOptionsDialog,
  openMacroRecordDialog,
  openMacroSecurityDialog,
  openMacroStepDialog,
  openPageLayoutBreaksMenu,
  openPageLayoutColorMenu,
  openPageLayoutEffectsMenu,
  openPageLayoutFontMenu,
  openPageLayoutMarginsMenu,
  openPageLayoutOrientationMenu,
  openPageLayoutPresetMenu,
  openPageLayoutPrintAreaMenu,
  openPageLayoutSizeMenu,
  openPageLayoutThemeMenu,
  openPageSetupDialog,
  openPrintDialog,
  openSheetCodeDialog,
  pageBreakTargetListsChanged,
  pageBreakTargetsForHeaderContext,
  pageBreakTargetsForSelection,
  pageBreakTargetsFromAxisIndex,
  pageBreakTargetsLabel,
  pageLayoutCurrentSelectionBreakState,
  pageLayoutEffectDescription,
  pageLayoutEffectPreset,
  pageLayoutEffectShapePatch,
  pageLayoutFitTargets,
  pageLayoutThemeDescription,
  pageSetupFitToOptions,
  pageSetupPrintTitleColumnsLabel,
  pageSetupPrintTitleRowsLabel,
  pageSetupRangeFieldHtml,
  parsePageSetupMarginValue,
  prepareWorkbookPrintSurface,
  printCellContentHtml,
  printCellContentStyle,
  printCellForMergeSlice,
  printCellHtml,
  printCellOverflowWidth,
  printCellSpanHeight,
  printCellSpanWidth,
  printCellStyle,
  printColumnIndexes,
  printColumnWidth,
  printDialogMarginPixels,
  printDialogPageSettings,
  printDialogPaperPixels,
  printDialogSettingsKey,
  printLayoutForDialog,
  printLayoutPrintablePixels,
  printMergeSliceSuppressedSides,
  printPagesForTargets,
  printPreviewPagesForDialog,
  printPreviewPagesForTarget,
  printRangeContentSize,
  printRangeForSelection,
  printRangeForSheet,
  printRenderScale,
  printRowHeight,
  printRowIndexes,
  printSheetEffectiveScale,
  printSheetHtml,
  printSheetTableHtml,
  printTargetForSheet,
  printTargetsForDialog,
  printWorkbookFromDialog,
  printWorkbookPageHtml,
  readVbaRangeValue,
  removePageBreakAt,
  removePageBreakFromSelection,
  renderPrintDialogPreview,
  resetAllPageBreaks,
  resolveVbaR1C1Index,
  schedulePrintDialogPreviewSync,
  scheduleWorkbookPrintSurfaceCleanup,
  selectPageBreakTargets,
  selectVbaRangeTarget,
  selectedMacroFromDialog,
  selectedSheetIndexesForPrint,
  setPageBackgroundFromSelection,
  setPrintAreaFromSelection,
  setPrintTitlesFromSelection,
  sheetCodeProcedureStub,
  sheetUsedRangeForPrint,
  splitFormulaStringLiterals,
  splitVbaArithmeticExpression,
  splitVbaComparisonExpression,
  splitVbaTopLevelWord,
  startMacroRecordingFromDialog,
  stopMacroRecording,
  syncMacroStepDialog,
  syncPageLayoutRibbonControls,
  syncPrintDialogPreview,
  syncPrintDialogPreviewNavigation,
  toggleMacroRelativeReferences,
  trimVbaOuterParentheses,
  vbaArrayToMatrix,
  vbaColorToCss,
  vbaComparableValue,
  vbaConstantValue,
  vbaCriteriaMatches,
  vbaCriteriaRowIndexes,
  vbaDateFromValue,
  vbaExpressionListValue,
  vbaHorizontalAlignment,
  vbaIndexValue,
  vbaLookupHorizontal,
  vbaLookupVertical,
  vbaMatchIndex,
  vbaOuterParenthesesWrapWholeExpression,
  vbaRangeAddress,
  vbaRangeMatrix,
  vbaRangeValues,
  vbaRoundNumber,
  vbaTruthy,
  vbaUnderlineTextDecoration,
  vbaVerticalAlignment,
  vbaWorksheetCriteriaPairs,
  workbookPrintPageStyle,
  writeVbaFormulaR1C1ToRange,
  writeVbaValueToRange,
};
