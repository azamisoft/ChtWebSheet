// Feature module; shared bindings are imported explicitly.
import {
  RECENT_DATA_SOURCES_LIMIT,
  RECENT_DATA_SOURCES_STORAGE_KEY,
  RECENT_DATA_SOURCE_TEXT_LIMIT,
  createIcons,
  icons,
  state,
} from "../../app-core.js";
import {
  closeWhatIfMenu,
} from "./analysis-and-scenarios.js";
import {
  activeWorkbookQueryRef,
  applyQuerySourceChange,
  cloneWorkbookDataModel,
  closeQueryPaneContextMenu,
  formatQueryTimestamp,
  htmlTableSources,
  loadWebRowsIntoSelection,
  nextQueryName,
  nextWorkbookQueryName,
  openActiveQueryPropertiesDialog,
  openExistingConnectionsDialog,
  openImportDataDialog,
  openQueriesConnectionsPane,
  openQueryPropertiesDialog,
  openWebNavigatorDialog,
  queryCommandText,
  querySortTime,
  querySourceChangeShouldClearCache,
  querySourceIconName,
  querySourceTypeUsesStoredRows,
  refreshActiveWorkbookConnection,
  refreshQueriesConnectionsPane,
  refreshSingleQuery,
  registerQueryInDataModel,
  registerQueryTableLoad,
  selectQueryRange,
  workbookQueries,
  workbookQueryRef,
  workbookQueryValueRows,
  writeWorkbookQueryRows,
  xmlElementToRowObject,
  xmlRowsFromElements,
} from "./connections-and-queries.js";
import {
  autoOutlineSelection,
  clearOutlineFromSelection,
  openOutlineGroupDialog,
  openOutlineSettingsDialog,
  outlineAxisForRange,
  outlineAxisSheetAllows,
} from "./forecast-and-statistics.js";
import {
  applyQueryTransformSteps,
  chooseDelimitedTextForImport,
  chooseFolderDataForImport,
  chooseStructuredDataForImport,
  chooseWorkbookDataForImport,
  openTextCsvImportDialog,
  queryTransformName,
  queryTransformResultForDescriptor,
  renderQueryTransformEditorDialog,
  stripTextBom,
} from "./query-editor.js";
import {
  cancelWorkbookRefreshes,
  detectDelimitedTextDelimiter,
  inferHeaderRow,
  openRefreshStatusDialog,
  parseDelimitedRows,
  querySourceForRefresh,
  rangeHasUserData,
  refreshWorkbookData,
  showDataOverwriteConfirmation,
} from "./sort-and-autofilter.js";
import {
  closeDataValidationMenu,
} from "./validation-and-consolidation.js";
import {
  closeDataDialog,
} from "../drawing/charts-and-insertion.js";
import {
  closeRibbonBorderMenu,
  closeRibbonColorPalette,
  closeRibbonInkEraserMenu,
  closeRibbonShapeMenu,
  closeRibbonUnderlineMenu,
  closeShapeFormatMenu,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  activeSelectionRange,
  getDisplayForCell,
  normalizeRange,
  rangeToLabel,
  selectionRangeFromPoints,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  dataRibbonRangeHasData,
  dataRibbonSelectionContext,
  outlineRangeHasLevel,
  rangesEqual,
  scrollRangeIntoView,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneJsonLikeValue,
  internCellsInRange,
  pushHistory,
} from "../editing/structure-and-history.js";
import {
  closeCellContextMenu,
  positionPopupMenu,
} from "../ui/annotations-and-cell-format.js";
import {
  accessKeyLabelHtml,
  cellKey,
  columnName,
  decodeRange,
  escapeAttr,
  escapeHtml,
  excelSerialToDate,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  commandMenuEscapeBlockedByDialog,
  consumeCommandMenuKey,
  enhanceCommandMenuAccessKeys,
} from "../ui/printing-and-home-tools.js";
import {
  applyExcelTableStyle,
  closeFileSideMenu,
  commitDataOperation,
  ensureSheetSize,
  nextTableName,
  openInternalHyperlink,
  selectedDataRange,
} from "../ui/ribbon-and-links.js";
import {
  closeSheetTabContextMenu,
  isSheetProtected,
} from "../ui/sheet-protection-and-translation.js";
import {
  buildFormulaEngine,
  renderWorkbook,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  activeSheet,
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
  clampOutlineLevel,
} from "../workbook/tables-and-validation.js";
import {
  closeViewCommandMenu,
} from "../workbook/views-and-macros.js";


function queryTransformMoveTargetIndex(fromIndex, length, position) {
  if (position === "start") return 0;
  if (position === "end") return Math.max(0, length - 1);
  if (position === "left") return Math.max(0, fromIndex - 1);
  if (position === "right") return Math.min(Math.max(0, length - 1), fromIndex + 1);
  return fromIndex;
}



function queryTransformStepColumnIndexes(columns, step) {
  const names = Array.isArray(step?.columns) && step.columns.length ? step.columns : [step?.columnName];
  return [...new Set(names
    .map((column) => columns.indexOf(column))
    .filter((index) => index >= 0))];
}



function moveQueryTransformArrayItem(items, fromIndex, toIndex) {
  const next = [...items];
  const [item] = next.splice(fromIndex, 1);
  next.splice(toIndex, 0, item);
  return next;
}



function queryTransformConditionMatches(value, operator = "equals", expected = "", dataType = "") {
  const actualText = String(value ?? "");
  const expectedText = String(expected ?? "");
  const actualLower = actualText.toLocaleLowerCase("ja-JP");
  const expectedLower = expectedText.toLocaleLowerCase("ja-JP");
  if (operator === "blank") return actualText.trim() === "";
  if (operator === "notBlank") return actualText.trim() !== "";
  if (operator === "notEquals") return actualLower !== expectedLower;
  if (operator === "contains") return actualLower.includes(expectedLower);
  if (operator === "notContains") return !actualLower.includes(expectedLower);
  if (operator === "beginsWith") return actualLower.startsWith(expectedLower);
  if (operator === "notBeginsWith") return !actualLower.startsWith(expectedLower);
  if (operator === "endsWith") return actualLower.endsWith(expectedLower);
  if (operator === "notEndsWith") return !actualLower.endsWith(expectedLower);
  if (operator === "greaterThan" || operator === "greaterOrEqual" || operator === "lessThan" || operator === "lessOrEqual") {
    const comparison = compareQueryTransformValues(actualText, expectedText, dataType || queryTransformInferComparisonType(actualText, expectedText));
    if (operator === "greaterThan") return comparison > 0;
    if (operator === "greaterOrEqual") return comparison >= 0;
    if (operator === "lessThan") return comparison < 0;
    return comparison <= 0;
  }
  return actualLower === expectedLower;
}



function queryTransformTextCaseValue(value, mode = "uppercase") {
  const text = String(value ?? "");
  if (mode === "lowercase") return text.toLocaleLowerCase("ja-JP");
  if (mode === "capitalize") {
    return text.replace(/[\p{L}\p{N}]+/gu, (word) => {
      const [first = "", ...rest] = [...word];
      return `${first.toLocaleUpperCase("ja-JP")}${rest.join("").toLocaleLowerCase("ja-JP")}`;
    });
  }
  return text.toLocaleUpperCase("ja-JP");
}



function queryTransformCharacterSetValue(value, characters = "", keep = true) {
  const characterSet = new Set([...String(characters ?? "")]);
  if (!characterSet.size) return String(value ?? "");
  return [...String(value ?? "")]
    .filter((char) => keep ? characterSet.has(char) : !characterSet.has(char))
    .join("");
}



function queryTransformNumberOperationValue(value, operation = "add", operand = 0) {
  const numeric = Number(String(value ?? "").replace(/,/g, ""));
  if (!Number.isFinite(numeric)) return String(value ?? "");
  const rhs = Number(operand);
  let result = numeric;
  if (operation === "add") result = numeric + (Number.isFinite(rhs) ? rhs : 0);
  else if (operation === "subtract") result = numeric - (Number.isFinite(rhs) ? rhs : 0);
  else if (operation === "multiply") result = numeric * (Number.isFinite(rhs) ? rhs : 1);
  else if (operation === "divide") result = Number.isFinite(rhs) && rhs !== 0 ? numeric / rhs : numeric;
  else if (operation === "integerDivide") result = Number.isFinite(rhs) && rhs !== 0 ? Math.trunc(numeric / rhs) : numeric;
  else if (operation === "modulo") result = Number.isFinite(rhs) && rhs !== 0 ? numeric % rhs : numeric;
  else if (operation === "round") result = Math.round(numeric);
  else if (operation === "floor") result = Math.floor(numeric);
  else if (operation === "ceil") result = Math.ceil(numeric);
  else if (operation === "absolute") result = Math.abs(numeric);
  else if (operation === "sign") result = Math.sign(numeric);
  else if (operation === "squareRoot") result = numeric < 0 ? Number.NaN : Math.sqrt(numeric);
  else if (operation === "logarithm") result = numeric <= 0 ? Number.NaN : Math.log10(numeric);
  else if (operation === "power") result = Math.pow(numeric, Number.isFinite(rhs) ? rhs : 2);
  else if (operation === "percentage") result = numeric / 100;
  else if (operation === "percentOf") result = Number.isFinite(rhs) && rhs !== 0 ? numeric / rhs : numeric;
  return queryTransformFormatNumber(result);
}



function queryTransformDatePartValue(value, part = "year") {
  const date = queryTransformParseDateValue(value);
  if (!date) return String(value ?? "");
  const year = date.getFullYear();
  const month = date.getMonth();
  const dayNames = ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"];
  if (part === "month") return String(date.getMonth() + 1);
  if (part === "day") return String(date.getDate());
  if (part === "quarter") return String(Math.floor(month / 3) + 1);
  if (part === "monthName") return `${month + 1}月`;
  if (part === "dayName") return dayNames[date.getDay()];
  if (part === "dayOfWeek") return String(date.getDay());
  if (part === "dayOfYear") {
    const start = Date.UTC(year, 0, 1);
    const current = Date.UTC(year, month, date.getDate());
    return String(Math.floor((current - start) / 86400000) + 1);
  }
  if (part === "startOfWeek") return queryTransformFormatDate(new Date(year, month, date.getDate() - date.getDay()));
  if (part === "endOfWeek") return queryTransformFormatDate(new Date(year, month, date.getDate() + (6 - date.getDay())));
  if (part === "daysInMonth") return String(new Date(year, month + 1, 0).getDate());
  if (part === "startOfMonth") return queryTransformFormatDate(new Date(year, month, 1));
  if (part === "endOfMonth") return queryTransformFormatDate(new Date(year, month + 1, 0));
  if (part === "startOfQuarter") return queryTransformFormatDate(new Date(year, Math.floor(month / 3) * 3, 1));
  if (part === "endOfQuarter") return queryTransformFormatDate(new Date(year, Math.floor(month / 3) * 3 + 3, 0));
  if (part === "startOfYear") return queryTransformFormatDate(new Date(year, 0, 1));
  if (part === "endOfYear") return queryTransformFormatDate(new Date(year, 11, 31));
  return String(year);
}



function queryTransformDatePartDataType(part = "year") {
  if (["monthName", "dayName"].includes(part)) return "text";
  if (["startOfWeek", "endOfWeek", "startOfMonth", "endOfMonth", "startOfQuarter", "endOfQuarter", "startOfYear", "endOfYear"].includes(part)) return "date";
  return "whole";
}



function queryTransformParseDateValue(value) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  const serial = Number(text.replace(/,/g, ""));
  if (Number.isFinite(serial) && serial > 0) {
    const serialDate = excelSerialToDate(serial);
    if (!Number.isNaN(serialDate.getTime())) return serialDate;
  }
  let match = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (match) {
    const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    return Number.isNaN(date.getTime()) ? null : date;
  }
  match = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (match) {
    const date = new Date(Number(match[3]), Number(match[1]) - 1, Number(match[2]));
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const fallback = new Date(text);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}



function queryTransformFormatDate(date) {
  const year = String(date.getFullYear()).padStart(4, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}



function queryTransformIsErrorValue(value) {
  const text = String(value ?? "").trim();
  return /^#(?:N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!|CALC!|SPILL!|FIELD!|BLOCKED!|UNKNOWN!)/i.test(text) ||
    /^error\b/i.test(text);
}



function queryTransformInferComparisonType(left, right) {
  const leftNumber = Number(String(left ?? "").replace(/,/g, ""));
  const rightNumber = Number(String(right ?? "").replace(/,/g, ""));
  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) return "number";
  const leftTime = Date.parse(String(left ?? ""));
  const rightTime = Date.parse(String(right ?? ""));
  if (Number.isFinite(leftTime) && Number.isFinite(rightTime)) return "date";
  return "";
}



function evaluateQueryCustomColumnFormula(formula, columns, row) {
  let expression = String(formula || "").trim();
  if (expression.startsWith("=")) expression = expression.slice(1).trim();
  if (!expression) return "";
  const conditional = evaluateQueryMIfExpression(expression, columns, row);
  if (conditional.matched) return String(conditional.value ?? "");
  const functionResult = evaluateQueryMFunctionExpression(expression, columns, row);
  if (functionResult.matched) return String(functionResult.value ?? "");
  if (expression.includes("&")) {
    return expression
      .split("&")
      .map((part) => queryTransformFormulaTextTerm(part, columns, row))
      .join("");
  }
  const referenceMatch = expression.match(/^\[([^\]]+)]$/);
  if (referenceMatch) return queryTransformFormulaColumnValue(referenceMatch[1], columns, row);
  const numericExpression = expression.replace(/\[([^\]]+)]/g, (_match, column) => {
    const value = queryTransformFormulaColumnValue(column, columns, row).replace(/,/g, "");
    const numeric = Number(value);
    return Number.isFinite(numeric) ? String(numeric) : "NaN";
  });
  if (/^[\d+\-*/().\s]+$/.test(numericExpression)) {
    try {
      const value = Function(`"use strict"; return (${numericExpression});`)();
      if (Number.isFinite(value)) return String(value);
    } catch (error) {
      console.warn("Failed to evaluate custom column formula", error);
    }
  }
  return expression.replace(/\[([^\]]+)]/g, (_match, column) => queryTransformFormulaColumnValue(column, columns, row));
}



function evaluateQueryMIfExpression(expression, columns, row) {
  const match = String(expression || "").match(/^if\s+(.+?)\s+then\s+(.+?)\s+else\s+(.+)$/i);
  if (!match) return { matched: false, value: "" };
  const condition = evaluateQueryMCondition(match[1], columns, row);
  return {
    matched: true,
    value: condition
      ? queryTransformFormulaValueTerm(match[2], columns, row)
      : queryTransformFormulaValueTerm(match[3], columns, row),
  };
}



function evaluateQueryMCondition(condition, columns, row) {
  const expression = String(condition || "").trim();
  const functionResult = evaluateQueryMFunctionExpression(expression, columns, row);
  if (functionResult.matched) return queryTransformBooleanValue(functionResult.value);
  if (/^(true|false)$/i.test(expression)) return /^true$/i.test(expression);
  const operatorMatch = expression.match(/^(.+?)\s*(<>|>=|<=|=|>|<)\s*(.+)$/);
  if (!operatorMatch) return Boolean(queryTransformFormulaValueTerm(expression, columns, row));
  const left = queryTransformFormulaValueTerm(operatorMatch[1], columns, row);
  const right = queryTransformFormulaValueTerm(operatorMatch[3], columns, row);
  const operator = operatorMatch[2];
  const comparisonType = queryTransformInferComparisonType(left, right);
  let compared = 0;
  if (comparisonType === "number") {
    compared = Number(String(left ?? "").replace(/,/g, "")) - Number(String(right ?? "").replace(/,/g, ""));
  } else if (comparisonType === "date") {
    compared = Date.parse(String(left ?? "")) - Date.parse(String(right ?? ""));
  } else {
    compared = String(left ?? "").localeCompare(String(right ?? ""), "ja-JP", { numeric: true, sensitivity: "base" });
  }
  if (operator === "=") return compared === 0;
  if (operator === "<>") return compared !== 0;
  if (operator === ">") return compared > 0;
  if (operator === "<") return compared < 0;
  if (operator === ">=") return compared >= 0;
  if (operator === "<=") return compared <= 0;
  return false;
}



function evaluateQueryMFunctionExpression(expression, columns, row) {
  const match = String(expression || "").trim().match(/^([A-Za-z]+(?:\.[A-Za-z]+)?)\s*\((.*)\)$/);
  if (!match) return { matched: false, value: "" };
  const name = match[1].toLowerCase();
  const args = splitQueryMArguments(match[2]).map((arg) => queryTransformFormulaValueTerm(arg, columns, row));
  const text = String(args[0] ?? "");
  const numberArg = (index, fallback = 0) => {
    const numeric = Number(String(args[index] ?? "").replace(/,/g, ""));
    return Number.isFinite(numeric) ? numeric : fallback;
  };
  const dateArg = (index = 0) => queryTransformParseDateValue(args[index]);
  const dateResult = (date) => ({ matched: true, value: date && !Number.isNaN(date.getTime()) ? queryTransformFormatDate(date) : "#VALUE!" });
  if (name === "text.upper") return { matched: true, value: text.toLocaleUpperCase("ja-JP") };
  if (name === "text.lower") return { matched: true, value: text.toLocaleLowerCase("ja-JP") };
  if (name === "text.proper") return { matched: true, value: text.replace(/\p{L}[\p{L}\p{M}'-]*/gu, (word) => word.charAt(0).toLocaleUpperCase("ja-JP") + word.slice(1).toLocaleLowerCase("ja-JP")) };
  if (name === "text.trim") return { matched: true, value: text.trim() };
  if (name === "text.clean") return { matched: true, value: text.replace(/[\u0000-\u001f\u007f]/g, "") };
  if (name === "text.length") return { matched: true, value: text.length };
  if (name === "text.startswith") return { matched: true, value: String(text.toLocaleLowerCase("ja-JP").startsWith(String(args[1] ?? "").toLocaleLowerCase("ja-JP"))).toLowerCase() };
  if (name === "text.endswith") return { matched: true, value: String(text.toLocaleLowerCase("ja-JP").endsWith(String(args[1] ?? "").toLocaleLowerCase("ja-JP"))).toLowerCase() };
  if (name === "text.positionof") return { matched: true, value: String(text.indexOf(String(args[1] ?? ""))) };
  if (name === "text.padstart") return { matched: true, value: text.padStart(Math.max(0, numberArg(1)), String(args[2] ?? " ").charAt(0) || " ") };
  if (name === "text.padend") return { matched: true, value: text.padEnd(Math.max(0, numberArg(1)), String(args[2] ?? " ").charAt(0) || " ") };
  if (name === "text.repeat") return { matched: true, value: text.repeat(Math.max(0, Math.min(32767, Math.round(numberArg(1, 1))))) };
  if (name === "text.select" || name === "text.remove") {
    const characters = new Set([...String(args[1] ?? "")]);
    if (!characters.size) return { matched: true, value: text };
    const keep = name === "text.select";
    return { matched: true, value: [...text].filter((char) => keep ? characters.has(char) : !characters.has(char)).join("") };
  }
  if (name === "text.start") return { matched: true, value: text.slice(0, Math.max(0, numberArg(1))) };
  if (name === "text.end") return { matched: true, value: text.slice(Math.max(0, text.length - Math.max(0, numberArg(1)))) };
  if (name === "text.range" || name === "text.middle") {
    const start = Math.max(0, numberArg(1));
    const count = args.length > 2 ? Math.max(0, numberArg(2)) : undefined;
    return { matched: true, value: count === undefined ? text.slice(start) : text.slice(start, start + count) };
  }
  if (name === "text.beforedelimiter") {
    const delimiter = String(args[1] ?? "");
    const index = delimiter ? text.indexOf(delimiter) : -1;
    return { matched: true, value: index >= 0 ? text.slice(0, index) : text };
  }
  if (name === "text.afterdelimiter") {
    const delimiter = String(args[1] ?? "");
    const index = delimiter ? text.indexOf(delimiter) : -1;
    return { matched: true, value: index >= 0 ? text.slice(index + delimiter.length) : "" };
  }
  if (name === "text.betweendelimiters") {
    const startDelimiter = String(args[1] ?? "");
    const endDelimiter = String(args[2] ?? "");
    const start = startDelimiter ? text.indexOf(startDelimiter) : -1;
    const from = start >= 0 ? start + startDelimiter.length : 0;
    const end = endDelimiter ? text.indexOf(endDelimiter, from) : -1;
    return { matched: true, value: end >= 0 ? text.slice(from, end) : "" };
  }
  if (name === "text.contains") return { matched: true, value: String(text.includes(String(args[1] ?? ""))).toLowerCase() };
  if (name === "text.replace") return { matched: true, value: text.replaceAll(String(args[1] ?? ""), String(args[2] ?? "")) };
  if (name === "number.from" || name === "number.fromtext") return { matched: true, value: queryTransformFormatNumber(numberArg(0, NaN)) || "#VALUE!" };
  if (name === "number.abs") return { matched: true, value: queryTransformFormatNumber(Math.abs(numberArg(0, NaN))) || "#VALUE!" };
  if (name === "number.round") return { matched: true, value: queryTransformFormatNumber(Number(numberArg(0, NaN).toFixed(Math.max(0, numberArg(1, 0))))) || "#VALUE!" };
  if (name === "number.rounddown") {
    const scale = 10 ** Math.max(0, numberArg(1, 0));
    return { matched: true, value: queryTransformFormatNumber(Math.floor(numberArg(0, NaN) * scale) / scale) || "#VALUE!" };
  }
  if (name === "number.roundup") {
    const scale = 10 ** Math.max(0, numberArg(1, 0));
    return { matched: true, value: queryTransformFormatNumber(Math.ceil(numberArg(0, NaN) * scale) / scale) || "#VALUE!" };
  }
  if (name === "number.integerdivide") return { matched: true, value: queryTransformFormatNumber(Math.trunc(numberArg(0, 0) / Math.max(1, numberArg(1, 1)))) };
  if (name === "number.mod") return { matched: true, value: queryTransformFormatNumber(numberArg(0, 0) % Math.max(1, numberArg(1, 1))) };
  if (name === "number.power") return { matched: true, value: queryTransformFormatNumber(Math.pow(numberArg(0, NaN), numberArg(1, 2))) || "#VALUE!" };
  if (name === "number.sign") return { matched: true, value: queryTransformFormatNumber(Math.sign(numberArg(0, NaN))) || "#VALUE!" };
  if (name === "number.sqrt") return { matched: true, value: queryTransformFormatNumber(Math.sqrt(numberArg(0, NaN))) || "#VALUE!" };
  if (name === "number.log10") return { matched: true, value: queryTransformFormatNumber(Math.log10(numberArg(0, NaN))) || "#VALUE!" };
  if (name === "date.from") {
    const date = queryTransformParseDateValue(text);
    return { matched: true, value: date ? queryTransformFormatDate(date) : "#VALUE!" };
  }
  if (name === "datetime.date") {
    const date = queryTransformParseDateValue(text);
    return { matched: true, value: date ? queryTransformFormatDate(date) : "#VALUE!" };
  }
  if (name === "datetime.from") {
    const date = queryTransformParseDateValue(text);
    return { matched: true, value: date ? queryTransformFormatDateTime(date) : "#VALUE!" };
  }
  if (name === "date.adddays") {
    const date = dateArg(0);
    return dateResult(date ? new Date(date.getFullYear(), date.getMonth(), date.getDate() + Math.trunc(numberArg(1, 0))) : null);
  }
  if (name === "date.addmonths") {
    const date = dateArg(0);
    return dateResult(date ? new Date(date.getFullYear(), date.getMonth() + Math.trunc(numberArg(1, 0)), date.getDate()) : null);
  }
  if (name === "date.addyears") {
    const date = dateArg(0);
    return dateResult(date ? new Date(date.getFullYear() + Math.trunc(numberArg(1, 0)), date.getMonth(), date.getDate()) : null);
  }
  if (name === "date.startofweek" || name === "date.endofweek") {
    const date = dateArg(0);
    if (!date) return dateResult(null);
    const offset = name === "date.startofweek" ? -date.getDay() : 6 - date.getDay();
    return dateResult(new Date(date.getFullYear(), date.getMonth(), date.getDate() + offset));
  }
  if (name === "date.daysinmonth") {
    const date = dateArg(0);
    return { matched: true, value: date ? String(new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate()) : "#VALUE!" };
  }
  if (name.startsWith("date.")) {
    const part = name.slice("date.".length);
    const aliases = {
      year: "year",
      month: "month",
      day: "day",
      quarter: "quarter",
      monthname: "monthName",
      dayofweekname: "dayName",
      dayofweek: "dayOfWeek",
      dayofyear: "dayOfYear",
      startofweek: "startOfWeek",
      endofweek: "endOfWeek",
      daysinmonth: "daysInMonth",
      startofmonth: "startOfMonth",
      endofmonth: "endOfMonth",
      startofquarter: "startOfQuarter",
      endofquarter: "endOfQuarter",
      startofyear: "startOfYear",
      endofyear: "endOfYear",
    };
    if (aliases[part]) return { matched: true, value: queryTransformDatePartValue(text, aliases[part]) };
  }
  return { matched: false, value: "" };
}



function splitQueryMArguments(argsText) {
  const parts = [];
  let current = "";
  let depth = 0;
  let quote = "";
  for (let index = 0; index < String(argsText || "").length; index += 1) {
    const char = argsText[index];
    const next = argsText[index + 1];
    if (quote) {
      current += char;
      if (char === quote) {
        if (next === quote) {
          current += next;
          index += 1;
        } else {
          quote = "";
        }
      }
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      current += char;
      continue;
    }
    if (char === "(" || char === "[" || char === "{") depth += 1;
    if (char === ")" || char === "]" || char === "}") depth = Math.max(0, depth - 1);
    if (char === "," && depth === 0) {
      parts.push(current.trim());
      current = "";
      continue;
    }
    current += char;
  }
  if (current.trim() || String(argsText || "").trim()) parts.push(current.trim());
  return parts;
}



function queryTransformFormulaValueTerm(part, columns, row) {
  const text = String(part || "").trim();
  const functionResult = evaluateQueryMFunctionExpression(text, columns, row);
  if (functionResult.matched) return functionResult.value;
  if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
    return text.slice(1, -1).replace(/""/g, '"').replace(/''/g, "'");
  }
  const referenceMatch = text.match(/^\[([^\]]+)]$/);
  if (referenceMatch) return queryTransformFormulaColumnValue(referenceMatch[1], columns, row);
  if (/^-?\d+(?:\.\d+)?$/.test(text)) return text;
  return text.replace(/\[([^\]]+)]/g, (_match, column) => queryTransformFormulaColumnValue(column, columns, row));
}



function queryTransformBooleanValue(value) {
  const text = String(value ?? "").trim().toLowerCase();
  if (!text || text === "false" || text === "0") return false;
  return true;
}



function queryTransformFormulaTextTerm(part, columns, row) {
  const text = String(part || "").trim();
  const value = queryTransformFormulaValueTerm(text, columns, row);
  if (value !== text) return String(value ?? "");
  if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
    return text.slice(1, -1).replace(/""/g, '"').replace(/''/g, "'");
  }
  const referenceMatch = text.match(/^\[([^\]]+)]$/);
  if (referenceMatch) return queryTransformFormulaColumnValue(referenceMatch[1], columns, row);
  return text;
}



function queryTransformFormulaColumnValue(column, columns, row) {
  const index = columns.indexOf(column);
  return index >= 0 ? String(row[index] ?? "") : "";
}



function queryTransformAggregateValue(rows, aggregateIndex, aggregateFunction = "sum", dataType = "") {
  const allowed = ["sum", "average", "count", "countNums", "countDistinct", "min", "max", "product", "median", "stdDev", "stdDevP", "var", "varP"];
  const fn = allowed.includes(aggregateFunction) ? aggregateFunction : "sum";
  if (fn === "count") return String(rows.length);
  if (aggregateIndex < 0) return "";
  const values = rows.map((row) => row[aggregateIndex] ?? "").filter((value) => String(value ?? "").trim() !== "");
  if (!values.length) return "";
  if (fn === "countDistinct") {
    return String(new Set(values.map((value) => String(value ?? "").trim().toLocaleLowerCase("ja-JP"))).size);
  }
  if (["sum", "average", "countNums", "product", "median", "stdDev", "stdDevP", "var", "varP"].includes(fn)) {
    const numbers = queryTransformAggregateNumbers(values);
    if (!numbers.length) return "";
    if (fn === "countNums") return String(numbers.length);
    const sum = numbers.reduce((total, value) => total + value, 0);
    if (fn === "sum") return queryTransformFormatNumber(sum);
    if (fn === "average") return queryTransformFormatNumber(sum / numbers.length);
    if (fn === "product") return queryTransformFormatNumber(numbers.reduce((total, value) => total * value, 1));
    if (fn === "median") return queryTransformFormatNumber(queryTransformMedian(numbers));
    if (fn === "stdDev" || fn === "var") {
      if (numbers.length < 2) return "";
      const variance = queryTransformVariance(numbers, true);
      return queryTransformFormatNumber(fn === "stdDev" ? Math.sqrt(variance) : variance);
    }
    if (fn === "stdDevP" || fn === "varP") {
      const variance = queryTransformVariance(numbers, false);
      return queryTransformFormatNumber(fn === "stdDevP" ? Math.sqrt(variance) : variance);
    }
  }
  const sorted = values.slice().sort((left, right) => compareQueryTransformValues(left, right, dataType));
  return String(fn === "max" ? sorted[sorted.length - 1] : sorted[0]);
}



function queryTransformAggregateNumbers(values) {
  return (values || [])
    .map((value) => Number(String(value ?? "").replace(/,/g, "")))
    .filter((value) => Number.isFinite(value));
}



function queryTransformMedian(numbers) {
  const sorted = (numbers || []).filter((value) => Number.isFinite(value)).sort((left, right) => left - right);
  if (!sorted.length) return NaN;
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}



function queryTransformVariance(numbers, sample = true) {
  const values = (numbers || []).filter((value) => Number.isFinite(value));
  if (!values.length || (sample && values.length < 2)) return NaN;
  const average = values.reduce((total, value) => total + value, 0) / values.length;
  const sumSquares = values.reduce((total, value) => total + (value - average) ** 2, 0);
  return sumSquares / (sample ? values.length - 1 : values.length);
}



function queryTransformFormatNumber(value) {
  if (!Number.isFinite(value)) return "";
  if (Number.isInteger(value)) return String(value);
  return String(Number(value.toFixed(12)));
}



function queryTransformAggregateDataType(aggregateFunction = "sum", sourceType = "") {
  if (["count", "countNums", "countDistinct", "sum", "average", "product", "median", "stdDev", "stdDevP", "var", "varP"].includes(aggregateFunction)) return "number";
  return sourceType || "any";
}



function queryTransformNormalizeDataType(value) {
  return ["text", "number", "whole", "date", "datetime", "any"].includes(value) ? value : "any";
}



function queryTransformInferDataType(values = []) {
  const samples = (values || [])
    .map((value) => String(value ?? "").trim())
    .filter(Boolean);
  if (!samples.length) return "any";
  const numericSamples = samples
    .map((value) => Number(value.replace(/,/g, "")))
    .filter((value) => Number.isFinite(value));
  if (numericSamples.length === samples.length) {
    return numericSamples.every((value) => Number.isInteger(value)) ? "whole" : "number";
  }
  const dateSamples = samples
    .map((value) => queryTransformParseDateValue(value))
    .filter(Boolean);
  if (dateSamples.length === samples.length) {
    const hasTime = samples.some((value) => /\d{1,2}:\d{2}/.test(value));
    return hasTime ? "datetime" : "date";
  }
  return "text";
}



function queryTransformConvertDataTypeValue(value, dataType = "any") {
  const type = queryTransformNormalizeDataType(dataType);
  if (type === "any") return String(value ?? "");
  if (type === "text") return String(value ?? "");
  const text = String(value ?? "").trim();
  if (!text) return "";
  if (type === "number" || type === "whole") {
    const numeric = Number(text.replace(/,/g, ""));
    if (!Number.isFinite(numeric)) return "#VALUE!";
    return queryTransformFormatNumber(type === "whole" ? Math.trunc(numeric) : numeric);
  }
  if (type === "date" || type === "datetime") {
    const date = queryTransformParseDateValue(text);
    if (!date) return "#VALUE!";
    return type === "datetime" ? queryTransformFormatDateTime(date) : queryTransformFormatDate(date);
  }
  return String(value ?? "");
}



function queryTransformFormatDateTime(date) {
  const time = [
    String(date.getHours()).padStart(2, "0"),
    String(date.getMinutes()).padStart(2, "0"),
    String(date.getSeconds()).padStart(2, "0"),
  ].join(":");
  return `${queryTransformFormatDate(date)} ${time}`;
}



function splitTextByCharacterCount(value, count) {
  const text = String(value ?? "");
  const size = Math.max(1, Math.round(Number(count) || 1));
  if (!text) return [""];
  const parts = [];
  for (let index = 0; index < text.length; index += size) {
    parts.push(text.slice(index, index + size));
  }
  return parts.length ? parts : [""];
}



function compareQueryTransformValues(left, right, dataType = "") {
  const leftText = String(left ?? "");
  const rightText = String(right ?? "");
  if (dataType === "number" || dataType === "whole") {
    const leftNumber = Number(leftText.replace(/,/g, ""));
    const rightNumber = Number(rightText.replace(/,/g, ""));
    if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) return leftNumber - rightNumber;
  }
  if (dataType === "date" || dataType === "datetime") {
    const leftTime = Date.parse(leftText);
    const rightTime = Date.parse(rightText);
    if (Number.isFinite(leftTime) && Number.isFinite(rightTime)) return leftTime - rightTime;
  }
  return leftText.localeCompare(rightText, "ja-JP", { numeric: true, sensitivity: "base" });
}



function sheetDisplayRowsForRange(sheet, range) {
  const sheetId = getSheetId(sheet.name);
  const rows = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    const values = [];
    for (let col = range.left; col <= range.right; col += 1) {
      values.push(String(getDisplayForCell(sheet, sheetId, row, col) ?? ""));
    }
    rows.push(values);
  }
  return rows;
}



function queryTransformDefaultColumns(width) {
  return Array.from({ length: width }, (_item, index) => `Column${index + 1}`);
}



function uniqueQueryTransformColumns(columns) {
  const counts = new Map();
  return columns.map((column, index) => {
    const base = String(column || `Column${index + 1}`).trim() || `Column${index + 1}`;
    const key = base.toLocaleLowerCase("ja-JP");
    const count = (counts.get(key) || 0) + 1;
    counts.set(key, count);
    return count === 1 ? base : `${base}.${count - 1}`;
  });
}



function padQueryTransformRow(row, width) {
  return Array.from({ length: width }, (_item, index) => String(row?.[index] ?? ""));
}



function queryTransformRowsForDescriptor(descriptor) {
  const result = queryTransformResultForDescriptor(descriptor);
  const preserveTrailingBlankRows = (descriptor?.transformSteps || []).some((step) => step?.type === "keepBlankRows");
  return workbookQueryValueRows(result.outputRows, descriptor?.typeDetection || "first200", {
    columnTypes: result.columnTypes,
    hasHeader: true,
    preserveTrailingBlankRows,
  });
}



function workbookQueryRowsWithTransforms(query, sourceRows) {
  if (!query?.transformSteps?.length) return workbookQueryValueRows(sourceRows, query?.typeDetection || "first200");
  const result = applyQueryTransformSteps(
    (sourceRows || []).map((row) => row.map((value) => String(value ?? ""))),
    query.transformSteps,
  );
  const preserveTrailingBlankRows = (query.transformSteps || []).some((step) => step?.type === "keepBlankRows");
  return workbookQueryValueRows(result.outputRows, query.typeDetection || "first200", {
    columnTypes: result.columnTypes,
    hasHeader: true,
    preserveTrailingBlankRows,
  });
}



function createQueryFromTransformDescriptor(descriptor, overrides = {}) {
  const sheet = activeSheet();
  const loadedAt = new Date().toISOString();
  return {
    id: uniqueWorksheetObjectId("query"),
    name: nextWorkbookQueryName(overrides.name || queryTransformName(descriptor) || nextQueryName(sheet)),
    sourceType: descriptor.sourceType || "textCsv",
    sourceName: descriptor.sourceName || "CSV",
    sourceText: descriptor.sourceText || "",
    sourceUrl: descriptor.sourceUrl || "",
    contentType: descriptor.contentType || "",
    sourceAuthMode: descriptor.sourceAuthMode || "",
    sourceAuthUsername: descriptor.sourceAuthUsername || "",
    sourceCredentialKey: descriptor.sourceCredentialKey || "",
    sourceItemIndex: descriptor.sourceItemIndex ?? 0,
    sourceItemName: descriptor.sourceItemName || "",
    sourceItemKind: descriptor.sourceItemKind || "document",
    sourceQueryId: descriptor.sourceQueryId || "",
    sourceQueryName: descriptor.sourceQueryName || "",
    sourceQuerySheetIndex: Number.isFinite(Number(descriptor.sourceQuerySheetIndex)) ? Number(descriptor.sourceQuerySheetIndex) : null,
    appendQueryRefs: Array.isArray(descriptor.appendQueryRefs) ? cloneJsonLikeValue(descriptor.appendQueryRefs) : [],
    mergeLeftRef: descriptor.mergeLeftRef ? cloneJsonLikeValue(descriptor.mergeLeftRef) : null,
    mergeRightRef: descriptor.mergeRightRef ? cloneJsonLikeValue(descriptor.mergeRightRef) : null,
    mergeLeftColumn: descriptor.mergeLeftColumn || "",
    mergeRightColumn: descriptor.mergeRightColumn || "",
    mergeJoinKind: descriptor.mergeJoinKind || "leftOuter",
    mergeRightPrefix: descriptor.mergeRightPrefix || "",
    sourceSheetIndex: Number.isFinite(Number(descriptor.sourceSheetIndex)) ? Number(descriptor.sourceSheetIndex) : null,
    sourceSheetName: descriptor.sourceSheetName || "",
    sourceRange: descriptor.sourceRange ? { ...descriptor.sourceRange } : null,
    sourceRows: Array.isArray(descriptor.sourceRows) ? cloneJsonLikeValue(descriptor.sourceRows) : null,
    origin: descriptor.origin || "65001",
    delimiter: descriptor.delimiter || ",",
    range: overrides.range ? { ...overrides.range } : null,
    hasHeader: true,
    typeDetection: descriptor.typeDetection || "first200",
    transformSteps: cloneJsonLikeValue(descriptor.transformSteps || []),
    loadMode: overrides.loadMode || "table",
    loadedAt,
    lastRefreshed: loadedAt,
    refreshWithAll: true,
  };
}



function loadQueryTransformEditorToSelection(options = {}) {
  const descriptor = state.dataDialog;
  if (!descriptor || descriptor.type !== "query-transform-editor") return false;
  const rows = queryTransformRowsForDescriptor(descriptor);
  if (!rows.height || !rows.width) {
    showMessageBox("読み込むデータがありません。", { title: "Cht WebSheet" });
    return false;
  }
  const targetPoint = descriptor.targetPoint || state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  const targetSheetIndex = targetPoint.sheetIndex ?? state.activeSheetIndex;
  const sheet = state.model?.sheets?.[targetSheetIndex];
  if (!sheet) return false;
  const targetRange = normalizeRange(
    { sheetIndex: targetSheetIndex, row: targetPoint.row, col: targetPoint.col },
    { sheetIndex: targetSheetIndex, row: targetPoint.row + rows.height - 1, col: targetPoint.col + rows.width - 1 },
  );
  if (!options.skipOverwriteCheck && rangeHasUserData(sheet, targetRange)) {
    const snapshot = cloneJsonLikeValue(descriptor);
    showDataOverwriteConfirmation(() => {
      state.dataDialog = snapshot;
      loadQueryTransformEditorToSelection({ skipOverwriteCheck: true });
    });
    return false;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  state.activeSheetIndex = targetSheetIndex;
  ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
  writeWorkbookQueryRows(sheet, rows, targetRange.top, targetRange.left);
  const existingRef = editingWorkbookQueryRef(descriptor);
  const query = existingRef?.query || createQueryFromTransformDescriptor(descriptor, { range: targetRange, name: queryTransformName(descriptor), loadMode: "table" });
  if (existingRef?.query) {
    applyQueryTransformDescriptorToQuery(query, descriptor, { range: targetRange, name: queryTransformName(descriptor), loadMode: "table" });
  } else {
    sheet.queries = Array.isArray(sheet.queries) ? sheet.queries : [];
    sheet.queries.push(query);
  }
  rememberRecentDataSourceFromQuery(query);
  registerQueryTableLoad(sheet, targetRange, query);
  internCellsInRange(sheet, targetRange);
  closeDataDialog();
  commitDataOperation(sheet, sheetBefore, targetRange, `${query.name || "Query"} 読み込み`, `${rangeToLabel(targetRange)} にクエリを読み込みました。`, targetRange);
  refreshQueriesConnectionsPane();
  scheduleWorkbookPersistence();
  return true;
}



function registerQueryTransformConnectionAndOpenLoadTo() {
  const descriptor = state.dataDialog;
  const sheet = activeSheet();
  if (!descriptor || descriptor.type !== "query-transform-editor" || !sheet) return;
  const existingRef = editingWorkbookQueryRef(descriptor);
  const targetSheet = existingRef?.sheet || sheet;
  const targetSheetIndex = existingRef?.sheetIndex ?? state.activeSheetIndex;
  const sheetBefore = captureSheetForHistory(targetSheet);
  const query = existingRef?.query || createQueryFromTransformDescriptor(descriptor, { name: queryTransformName(descriptor), loadMode: "connectionOnly" });
  if (existingRef?.query) {
    applyQueryTransformDescriptorToQuery(query, descriptor, { name: queryTransformName(descriptor), loadMode: query.loadMode || "connectionOnly" });
  } else {
    targetSheet.queries = Array.isArray(targetSheet.queries) ? targetSheet.queries : [];
    targetSheet.queries.push(query);
  }
  rememberRecentDataSourceFromQuery(query);
  pushHistory({
    sheetIndex: targetSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(targetSheet),
    selectionBefore: state.selectionRange || null,
    selectionAfter: state.selectionRange || null,
    label: `${query.name || "Query"} ${existingRef ? "クエリ編集" : "接続作成"}`,
  });
  refreshQueriesConnectionsPane();
  scheduleWorkbookPersistence();
  openImportDataDialog(targetSheetIndex, query.id, { defaultMode: "table" });
}



function editingWorkbookQueryRef(descriptor) {
  if (!descriptor?.editingQueryId) return null;
  return workbookQueryRef(Number(descriptor.editingQuerySheetIndex), descriptor.editingQueryId);
}



function applyQueryTransformDescriptorToQuery(query, descriptor, overrides = {}) {
  const loadedAt = new Date().toISOString();
  Object.assign(query, {
    name: overrides.name || queryTransformName(descriptor) || query.name || "Query",
    sourceType: descriptor.sourceType || query.sourceType || "textCsv",
    sourceName: descriptor.sourceName || query.sourceName || "",
    sourceText: descriptor.sourceText || "",
    sourceUrl: descriptor.sourceUrl || "",
    contentType: descriptor.contentType || "",
    sourceAuthMode: descriptor.sourceAuthMode || "",
    sourceAuthUsername: descriptor.sourceAuthUsername || "",
    sourceCredentialKey: descriptor.sourceCredentialKey || "",
    sourceItemIndex: descriptor.sourceItemIndex ?? 0,
    sourceItemName: descriptor.sourceItemName || "",
    sourceItemKind: descriptor.sourceItemKind || "document",
    sourceQueryId: descriptor.sourceQueryId || "",
    sourceQueryName: descriptor.sourceQueryName || "",
    sourceQuerySheetIndex: Number.isFinite(Number(descriptor.sourceQuerySheetIndex)) ? Number(descriptor.sourceQuerySheetIndex) : null,
    appendQueryRefs: Array.isArray(descriptor.appendQueryRefs) ? cloneJsonLikeValue(descriptor.appendQueryRefs) : [],
    mergeLeftRef: descriptor.mergeLeftRef ? cloneJsonLikeValue(descriptor.mergeLeftRef) : null,
    mergeRightRef: descriptor.mergeRightRef ? cloneJsonLikeValue(descriptor.mergeRightRef) : null,
    mergeLeftColumn: descriptor.mergeLeftColumn || "",
    mergeRightColumn: descriptor.mergeRightColumn || "",
    mergeJoinKind: descriptor.mergeJoinKind || "leftOuter",
    mergeRightPrefix: descriptor.mergeRightPrefix || "",
    sourceSheetIndex: Number.isFinite(Number(descriptor.sourceSheetIndex)) ? Number(descriptor.sourceSheetIndex) : null,
    sourceSheetName: descriptor.sourceSheetName || "",
    sourceRange: descriptor.sourceRange ? { ...descriptor.sourceRange } : null,
    sourceRows: Array.isArray(descriptor.sourceRows) ? cloneJsonLikeValue(descriptor.sourceRows) : null,
    origin: descriptor.origin || query.origin || "65001",
    delimiter: descriptor.delimiter || query.delimiter || ",",
    range: overrides.range ? { ...overrides.range } : query.range || null,
    hasHeader: true,
    typeDetection: descriptor.typeDetection || query.typeDetection || "first200",
    transformSteps: cloneJsonLikeValue(descriptor.transformSteps || []),
    loadMode: overrides.loadMode || query.loadMode || "table",
    loadedAt,
    lastRefreshed: loadedAt,
    refreshWithAll: query.refreshWithAll !== false,
  });
  return query;
}



function openGetDataFromTableRangeDialog() {
  const context = selectedDataRange("テーブルまたは範囲からデータを取得", { limit: 50000, expandCurrentRegion: true });
  if (!context) return;
  const { sheet, range } = context;
  state.dataDialog = { type: "get-table-range", range };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog get-table-range-dialog" role="dialog" aria-modal="true" aria-labelledby="getTableRangeTitle" data-data-dialog="get-table-range">
      <div class="excel-mini-titlebar" id="getTableRangeTitle">テーブルまたは範囲からデータを取得</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">範囲</span><input type="text" data-data-field="sourceRange" value="${escapeAttr(rangeToLabel(range))}" /></label>
        <label class="data-dialog-check"><input type="checkbox" data-data-field="hasHeader" ${inferHeaderRow(sheet, range) ? "checked" : ""} />先頭行をテーブルの見出しとして使用する</label>
        <div class="query-preview" aria-label="プレビュー">${queryPreviewHtml(sheet, range)}</div>
        <div class="data-dialog-note">読み込みを実行すると、選択範囲をテーブルとして登録し、このブック内のクエリとして保持します。</div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="transform">データの変換</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="load">読み込み</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
}



function openTableRangeTransformEditor(dialog) {
  const sourceText = dialog.querySelector("[data-data-field='sourceRange']")?.value || "";
  const range = decodeRange(sourceText) || state.dataDialog?.range;
  const sheet = activeSheet();
  if (!range || !sheet) {
    showMessageBox("範囲を正しく入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const targetPoint = state.selected || { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
  state.dataDialog = {
    type: "query-transform-editor",
    sourceType: "tableRange",
    sourceName: `${sheet.name}!${rangeToLabel(range)}`,
    sourceSheetIndex: state.activeSheetIndex,
    sourceSheetName: sheet.name,
    sourceRange: { ...range },
    sourceRows: sheetDisplayRowsForRange(sheet, range),
    targetPoint: { ...targetPoint },
    transformSteps: [],
  };
  renderQueryTransformEditorDialog();
}



function queryPreviewHtml(sheet, range) {
  const sheetId = getSheetId(sheet.name);
  const maxRows = Math.min(range.bottom, range.top + 9);
  const maxCols = Math.min(range.right, range.left + 5);
  const rows = [];
  for (let row = range.top; row <= maxRows; row += 1) {
    const cells = [];
    for (let col = range.left; col <= maxCols; col += 1) {
      cells.push(`<td>${escapeHtml(getDisplayForCell(sheet, sheetId, row, col))}</td>`);
    }
    rows.push(`<tr>${cells.join("")}</tr>`);
  }
  return `<table><tbody>${rows.join("")}</tbody></table>`;
}



function applyGetDataFromTableRangeDialog(dialog) {
  const sourceText = dialog.querySelector("[data-data-field='sourceRange']")?.value || "";
  const range = decodeRange(sourceText) || state.dataDialog?.range;
  const sheet = activeSheet();
  if (!range || !sheet) {
    showMessageBox("範囲を正しく入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const hasHeader = Boolean(dialog.querySelector("[data-data-field='hasHeader']")?.checked);
  closeDataDialog();
  const sheetBefore = captureSheetForHistory(sheet);
  ensureSheetSize(sheet, range.bottom, range.right);
  applyExcelTableStyle(sheet, range);
  sheet.tables = Array.isArray(sheet.tables) ? sheet.tables : [];
  const existingTable = sheet.tables.find((table) => table.range && rangesEqual(table.range, range));
  if (existingTable) {
    existingTable.showHeaderRow = hasHeader;
    existingTable.showFilterButton = true;
  } else {
    sheet.tables.push({
      id: uniqueWorksheetObjectId("table"),
      name: nextTableName(sheet),
      range: { ...range },
      showHeaderRow: hasHeader,
      showFilterButton: true,
      style: "TableStyleMedium2",
    });
  }
  sheet.queries = Array.isArray(sheet.queries) ? sheet.queries : [];
  sheet.queries.push({
    id: uniqueWorksheetObjectId("query"),
    name: nextQueryName(sheet),
    sourceType: "tableRange",
    range: { ...range },
    hasHeader,
    loadedAt: new Date().toISOString(),
  });
  commitDataOperation(sheet, sheetBefore, range, `${rangeToLabel(range)} クエリ`, `${rangeToLabel(range)} をクエリとして読み込みました。`, range);
}



function openDataGetMenu(trigger = null) {
  if (!trigger) {
    openGetDataFromTableRangeDialog();
    return;
  }
  const sheet = state.model ? activeSheet() : null;
  const queryRefs = state.model ? workbookQueries() : [];
  const recentSources = state.model ? loadRecentDataSources() : [];
  const context = sheet ? dataRibbonSelectionContext(sheet, { expandCurrentRegion: true }) : null;
  const canGetTable = Boolean(context?.range && dataRibbonRangeHasData(sheet, context.range) && !context.range.fullRow && !context.range.fullColumn);
  const disabled = state.model ? "" : " disabled";
  const noTable = canGetTable ? "" : " disabled";
  const noRecent = queryRefs.length || recentSources.length ? "" : " disabled";
  const noConnections = queryRefs.length ? "" : " disabled";
  openDataCommandMenu(trigger, "data-get", "データの取得", `
    <button type="button" role="menuitem" data-data-menu-action="get-workbook"${disabled}><i data-lucide="file-spreadsheet"></i><span>${accessKeyLabelHtml("Excel ブックから(B)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="get-folder"${disabled}><i data-lucide="folder-open"></i><span>${accessKeyLabelHtml("フォルダーから(O)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="get-text"${disabled}><i data-lucide="file-text"></i><span>${accessKeyLabelHtml("テキストまたは CSV から(T)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="get-web"${disabled}><i data-lucide="globe"></i><span>${accessKeyLabelHtml("Web から(W)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="get-odata"${disabled}><i data-lucide="database"></i><span>${accessKeyLabelHtml("OData フィードから(D)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="get-table"${noTable}><i data-lucide="table-2"></i><span>${accessKeyLabelHtml("テーブルまたは範囲から(F)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="get-json"${disabled}><i data-lucide="braces"></i><span>${accessKeyLabelHtml("JSON から(J)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="get-xml"${disabled}><i data-lucide="file-code-2"></i><span>${accessKeyLabelHtml("XML から(X)...")}</span></button>
    <div class="ribbon-data-menu-separator" role="separator"></div>
    <button type="button" role="menuitem" data-data-menu-action="recent-sources"${noRecent}><i data-lucide="clock-3"></i><span>${accessKeyLabelHtml("最近使ったソース(R)")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="existing-connections"${noConnections}><i data-lucide="plug"></i><span>${accessKeyLabelHtml("既存の接続(E)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="queries-connections"${noConnections}><i data-lucide="book-open"></i><span>${accessKeyLabelHtml("クエリと接続(Q)")}</span></button>
    <div class="ribbon-data-menu-separator" role="separator"></div>
    <button type="button" role="menuitem" data-data-menu-action="other-data-sources"${disabled}><i data-lucide="database"></i><span>${accessKeyLabelHtml("その他のデータ ソース(O)")}</span><small>›</small></button>`);
}



function openOtherDataSourcesMenu(trigger = null) {
  if (!trigger) {
    openODataImportDialog();
    return;
  }
  const disabled = state.model ? "" : " disabled";
  openDataCommandMenu(trigger, "other-data-sources", "その他のデータ ソース", `
    <button type="button" role="menuitem" data-data-menu-action="get-blank-query"${disabled}><i data-lucide="file-plus"></i><span>${accessKeyLabelHtml("空のクエリ(B)")}</span></button>
    <div class="ribbon-data-menu-separator" role="separator"></div>
    <button type="button" role="menuitem" data-data-menu-action="get-odata"${disabled}><i data-lucide="database"></i><span>${accessKeyLabelHtml("OData フィードから(D)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="get-web"${disabled}><i data-lucide="globe"></i><span>${accessKeyLabelHtml("Web から(W)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="get-json"${disabled}><i data-lucide="braces"></i><span>${accessKeyLabelHtml("JSON から(J)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="get-xml"${disabled}><i data-lucide="file-code-2"></i><span>${accessKeyLabelHtml("XML から(X)...")}</span></button>
    <div class="ribbon-data-menu-separator" role="separator"></div>
    <button type="button" role="menuitem" data-data-menu-action="get-sql-server"${disabled}><i data-lucide="plug"></i><span>${accessKeyLabelHtml("SQL Server データベースから(S)...")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="get-access"${disabled}><i data-lucide="table-2"></i><span>${accessKeyLabelHtml("Microsoft Access データベースから(A)...")}</span></button>`);
}



function openRecentSourcesMenu(trigger = null) {
  if (!trigger) {
    openExistingConnectionsDialog();
    return;
  }
  const queryRows = workbookQueries()
    .slice()
    .sort((left, right) => querySortTime(right.query) - querySortTime(left.query))
    .slice(0, 9);
  const queryButtons = queryRows.length
    ? queryRows.map(({ sheet, sheetIndex, query }) => {
      const range = query.range ? ` ${rangeToLabel(query.range)}` : "";
      return `<button type="button" role="menuitem" data-data-menu-action="recent-source" data-query-sheet-index="${sheetIndex}" data-query-id="${escapeAttr(query.id)}"><i data-lucide="${querySourceIconName(query)}"></i><span><strong>${escapeHtml(query.name || "Query")}</strong><small>${escapeHtml(sheet.name)}${escapeHtml(range)}</small></span></button>`;
    }).join("")
    : "";
  const recentSources = loadRecentDataSources();
  const sourceButtons = recentSources.length
    ? recentSources.map((source) => {
      const disabled = canOpenRecentDataSource(source) ? "" : " disabled";
      const detail = recentDataSourceDetailLabel(source);
      return `<button type="button" role="menuitem" data-data-menu-action="recent-source-open" data-recent-source-id="${escapeAttr(source.id)}"${disabled}><i data-lucide="${recentDataSourceIconName(source)}"></i><span><strong>${escapeHtml(source.label || source.sourceName || "Source")}</strong><small>${escapeHtml(detail)}</small></span></button>`;
    }).join("")
    : "";
  const sections = [];
  if (queryButtons) sections.push(queryButtons);
  if (queryButtons && sourceButtons) sections.push('<div class="ribbon-data-menu-separator" role="separator"></div>');
  if (sourceButtons) sections.push(sourceButtons);
  const rows = sections.length
    ? sections.join("")
    : '<button type="button" role="menuitem" class="is-disabled" disabled><i data-lucide="clock-3"></i><span>最近使ったソースはありません</span></button>';
  openDataCommandMenu(trigger, "recent-sources", "最近使ったソース", rows);
}



function loadRecentDataSources() {
  try {
    const parsed = JSON.parse(localStorage.getItem(RECENT_DATA_SOURCES_STORAGE_KEY) || "[]");
    return Array.isArray(parsed)
      ? parsed.filter((item) => item && typeof item === "object" && item.id).slice(0, RECENT_DATA_SOURCES_LIMIT)
      : [];
  } catch (error) {
    return [];
  }
}



function saveRecentDataSources(sources) {
  try {
    localStorage.setItem(RECENT_DATA_SOURCES_STORAGE_KEY, JSON.stringify((sources || []).slice(0, RECENT_DATA_SOURCES_LIMIT)));
  } catch (error) {
    // Recent sources are a convenience feature; importing itself should not fail if storage is unavailable.
  }
}



function rememberRecentDataSourceFromQuery(query) {
  const source = recentDataSourceFromQuery(query);
  if (!source) return;
  const sources = loadRecentDataSources().filter((item) => item.id !== source.id);
  saveRecentDataSources([source, ...sources]);
}



function recentDataSourceFromQuery(query) {
  if (!query || !["textCsv", "web", "odata", "json", "xml", "workbook", "folder"].includes(query.sourceType)) return null;
  const sourceName = query.sourceName || query.sourceUrl || query.name || query.sourceItemName || "Source";
  const sourceText = String(query.sourceText || "");
  const sourceRows = Array.isArray(query.sourceRows) ? query.sourceRows : null;
  const rowPayload = sourceRows ? JSON.stringify(sourceRows) : "";
  const storableRows = rowPayload && rowPayload.length <= RECENT_DATA_SOURCE_TEXT_LIMIT ? cloneJsonLikeValue(sourceRows) : null;
  const storableText = sourceText && sourceText.length <= RECENT_DATA_SOURCE_TEXT_LIMIT ? sourceText : "";
  const sessionCredential = query.sourceCredentialKey ? state.queryCredentials?.get?.(query.sourceCredentialKey) : null;
  const source = {
    id: recentDataSourceId(query),
    sourceType: query.sourceType,
    label: query.sourceItemName || sourceName,
    sourceName,
    sourceUrl: query.sourceUrl || "",
    sourceText: storableText,
    contentType: query.contentType || "",
    sourceRows: storableRows,
    sourceItemIndex: query.sourceItemIndex ?? 0,
    sourceItemName: query.sourceItemName || "",
    sourceItemKind: query.sourceItemKind || query.sourceType,
    origin: query.origin || "65001",
    delimiter: query.delimiter || "",
    typeDetection: query.typeDetection || "first200",
    sourceAuthMode: query.sourceAuthMode || "",
    sourceAuthUsername: query.sourceAuthUsername || sessionCredential?.username || "",
    sourceCredentialKey: sessionCredential ? query.sourceCredentialKey : "",
    lastUsedAt: new Date().toISOString(),
  };
  return canRememberRecentDataSource(source) ? source : null;
}



function canRememberRecentDataSource(source) {
  if (!source?.sourceType) return false;
  if (source.sourceText || source.sourceUrl || source.sourceRows?.length) return true;
  return false;
}



function recentDataSourceId(query) {
  const sourceType = query?.sourceType || "source";
  const source = query?.sourceUrl || query?.sourceName || query?.name || "";
  const item = query?.sourceItemName || query?.sourceItemIndex || "";
  return `${sourceType}:${source}:${item}`.slice(0, 500);
}



function recentDataSourceDetailLabel(source) {
  const parts = [recentDataSourceTypeLabel(source)];
  if (source.sourceUrl && source.sourceUrl !== source.sourceName) parts.push(source.sourceUrl);
  else if (source.sourceName) parts.push(source.sourceName);
  if (source.sourceAuthMode === "basic") parts.push("基本認証");
  if (source.lastUsedAt) parts.push(formatQueryTimestamp(source.lastUsedAt));
  if (!canOpenRecentDataSource(source)) parts.push("再選択が必要");
  return parts.filter(Boolean).join(" / ");
}



function recentDataSourceTypeLabel(source) {
  switch (source?.sourceType) {
    case "textCsv":
      return "テキスト/CSV";
    case "web":
      return "Web";
    case "odata":
      return "OData";
    case "json":
      return "JSON";
    case "xml":
      return "XML";
    case "workbook":
      return "Excel ブック";
    case "folder":
      return "フォルダー";
    default:
      return "データ ソース";
  }
}



function recentDataSourceIconName(source) {
  return querySourceIconName({ sourceType: source?.sourceType });
}



function canOpenRecentDataSource(source) {
  if (!source) return false;
  if (source.sourceText || source.sourceRows?.length) return true;
  if (source.sourceUrl && /^(?:https?:|data:)/i.test(source.sourceUrl)) return true;
  return false;
}



async function openRecentDataSource(sourceId) {
  const source = loadRecentDataSources().find((item) => item.id === sourceId);
  if (!source) {
    showMessageBox("最近使ったソースが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  if (!state.model) {
    setStatus("Excel ファイルを開いてからデータを取得してください。");
    return;
  }
  if (!canOpenRecentDataSource(source)) {
    showMessageBox("この最近使ったソースを開くには、元のファイルをもう一度選択してください。", { title: "Cht WebSheet" });
    return;
  }
  if (recentDataSourceNeedsCredential(source) && !recentDataSourceAuthOptions(source)) {
    reopenRecentDataSourceCredentialDialog(source);
    return;
  }
  const targetPoint = state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  if (source.sourceType === "textCsv" && source.sourceText) {
    openTextCsvImportDialog(source.sourceText, source.sourceName || source.label || "CSV", {
      origin: source.origin || "65001",
      recentSourceId: source.id,
    });
    setStatus(`${source.label || source.sourceName || "CSV"} を最近使ったソースから開きました。`);
    return;
  }
  let navigatorSource = null;
  try {
    navigatorSource = await recentDataSourceNavigatorSource(source, targetPoint);
  } catch (error) {
    console.error(error);
    showMessageBox("最近使ったソースを読み込めませんでした。URL とアクセス権を確認してください。", { title: "Cht WebSheet" });
    return;
  }
  if (!navigatorSource) {
    showMessageBox("読み込める表データが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  openWebNavigatorDialog(navigatorSource);
  setStatus(`${source.label || source.sourceName || "ソース"} を最近使ったソースから開きました。`);
}



function recentDataSourceNeedsCredential(source) {
  if (source?.sourceAuthMode !== "basic") return false;
  if (source.sourceText || source.sourceRows?.length) return false;
  return Boolean(source.sourceUrl && /^(?:https?:|data:)/i.test(source.sourceUrl));
}



function recentDataSourceAuthOptions(source) {
  if (source?.sourceAuthMode !== "basic") return null;
  const credential = source.sourceCredentialKey ? state.queryCredentials?.get?.(source.sourceCredentialKey) : null;
  if (!credential?.password && credential?.password !== "") return null;
  return {
    mode: "basic",
    username: credential.username || source.sourceAuthUsername || "",
    password: credential.password,
  };
}



function recentDataSourceAuthMetadata(source) {
  const auth = recentDataSourceAuthOptions(source);
  return {
    sourceAuthMode: source?.sourceAuthMode || "",
    sourceAuthUsername: source?.sourceAuthUsername || "",
    sourceCredentialKey: auth ? source.sourceCredentialKey || "" : "",
  };
}



function reopenRecentDataSourceCredentialDialog(source) {
  const options = {
    sourceUrl: source.sourceUrl || source.sourceName || "",
    sourceAuthMode: "basic",
    sourceAuthUsername: source.sourceAuthUsername || "",
  };
  if (source.sourceType === "odata") openODataImportDialog(options);
  else openWebImportDialog(options);
  setStatus("最近使ったソースの資格情報を入力してください。");
}



async function recentDataSourceNavigatorSource(source, targetPoint) {
  if (Array.isArray(source.sourceRows) && source.sourceRows.length) {
    return {
      sourceType: source.sourceType,
      sourceName: source.sourceName || source.label || "Source",
      sourceUrl: source.sourceUrl || source.sourceName || "",
      sourceText: source.sourceText || "",
      contentType: source.contentType || "",
      ...recentDataSourceAuthMetadata(source),
      sources: [{
        kind: source.sourceItemKind || source.sourceType,
        itemIndex: source.sourceItemIndex ?? 0,
        name: source.sourceItemName || source.label || source.sourceName || "Source",
        rows: cloneJsonLikeValue(source.sourceRows),
      }],
      targetPoint: { ...targetPoint },
    };
  }
  const sourceText = source.sourceText || await fetchRecentDataSourceText(source);
  const contentType = source.contentType || recentDataSourceContentType(source);
  const sources = source.sourceType === "odata"
    ? odataImportSourcesFromSource(sourceText, contentType, source.sourceUrl || source.sourceName || "")
    : webImportSourcesFromSource(sourceText, contentType, source.sourceUrl || source.sourceName || "");
  if (!sources.length) return null;
  const selected = querySourceForRefresh(source, sources) || sources[0];
  return {
    sourceType: source.sourceType || "web",
    sourceName: source.sourceName || source.sourceUrl || source.label || "Source",
    sourceUrl: source.sourceUrl || source.sourceName || "",
    sourceText,
    contentType,
    ...recentDataSourceAuthMetadata(source),
    sources,
    rows: selected?.rows || sources[0]?.rows || [],
    sourceItemIndex: selected?.itemIndex ?? source.sourceItemIndex ?? 0,
    sourceItemName: selected?.name || source.sourceItemName || "",
    sourceItemKind: selected?.kind || source.sourceItemKind || source.sourceType || "document",
    targetPoint: { ...targetPoint },
  };
}



async function fetchRecentDataSourceText(source) {
  if (!source?.sourceUrl || !/^(?:https?:|data:)/i.test(source.sourceUrl)) return "";
  const response = await fetch(source.sourceUrl, {
    headers: requestHeadersForImportAuth(recentDataSourceAuthOptions(source)),
  });
  if (!response.ok) throw new Error(`${response.status} ${response.statusText}`.trim());
  return response.text();
}



function recentDataSourceContentType(source) {
  if (source?.contentType) return source.contentType;
  if (source?.sourceType === "json") return "application/json";
  if (source?.sourceType === "xml") return "application/xml";
  if (source?.sourceType === "odata") return odataContentTypeFromUrl(source.sourceUrl || source.sourceName || "");
  return "";
}



function openDataRefreshMenu(trigger = null) {
  if (!trigger) {
    void refreshWorkbookData();
    return;
  }
  const queryRefs = workbookQueries();
  const activeRef = activeWorkbookQueryRef();
  const disabled = activeRef ? "" : " disabled";
  const hasConnections = queryRefs.length ? "" : " disabled";
  const hasRefreshAll = queryRefs.some(({ query }) => query.refreshWithAll !== false) ? "" : " disabled";
  const hasActiveRefresh = state.activeQueryRefreshes?.length ? "" : " disabled";
  openDataCommandMenu(trigger, "refresh", "すべて更新", `
    <button type="button" role="menuitem" data-data-menu-action="refresh-all"${hasRefreshAll}><i data-lucide="refresh-cw"></i><span>${accessKeyLabelHtml("すべて更新(A)")}</span></button>
    <button type="button" role="menuitem" data-data-menu-action="refresh-active"${disabled}><i data-lucide="rotate-cw"></i><span>${accessKeyLabelHtml("更新(R)")}</span></button>
    <div class="ribbon-data-menu-separator" role="separator"></div>
    <button type="button" role="menuitem" data-data-menu-action="refresh-status"${hasConnections}><i data-lucide="activity"></i><span>${accessKeyLabelHtml("更新の状態(S)")}</span></button>
	    <button type="button" role="menuitem" data-data-menu-action="cancel-refresh"${hasActiveRefresh}><i data-lucide="circle-stop"></i><span>${accessKeyLabelHtml("更新のキャンセル(C)")}</span></button>
	    <div class="ribbon-data-menu-separator" role="separator"></div>
	    <button type="button" role="menuitem" data-data-menu-action="connection-properties"${disabled}><i data-lucide="settings"></i><span>${accessKeyLabelHtml("接続のプロパティ(P)...")}</span></button>
	    <button type="button" role="menuitem" data-data-menu-action="existing-connections"${hasConnections}><i data-lucide="plug"></i><span>${accessKeyLabelHtml("接続(C)...")}</span></button>
	    <button type="button" role="menuitem" data-data-menu-action="workbook-links"><i data-lucide="link"></i><span>${accessKeyLabelHtml("ブックのリンク(L)...")}</span></button>`);
}



function openOutlineCommandMenu(trigger = null, delta = 1) {
  if (!trigger) {
    openOutlineGroupDialog(delta);
    return;
  }
  const sheet = state.model ? activeSheet() : null;
  const context = sheet ? dataRibbonSelectionContext(sheet, { expandCurrentRegion: true }) : null;
  const outlineAxis = context?.range ? outlineAxisForRange(context.range) : "rows";
  const canChangeOutlineAxis = Boolean(sheet && outlineAxisSheetAllows(sheet, outlineAxis));
  const canGroup = Boolean(context?.range && canChangeOutlineAxis);
  const canUngroup = Boolean(context?.range && canChangeOutlineAxis && outlineRangeHasLevel(sheet, context.range));
  const disabled = state.model ? "" : " disabled";
  const groupDisabled = canGroup ? "" : " disabled";
  const ungroupDisabled = canUngroup ? "" : " disabled";
  const hasOutline = Boolean(
    sheet?.rows?.some((row) => clampOutlineLevel(row?.outlineLevel)) ||
    sheet?.columns?.some((column) => clampOutlineLevel(column?.outlineLevel)),
  );
  const canClear = Boolean(hasOutline && sheet && (!isSheetProtected(sheet) || (outlineAxisSheetAllows(sheet, "rows") && outlineAxisSheetAllows(sheet, "columns"))));
  const clearDisabled = canClear ? "" : " disabled";
  const isGroup = delta > 0;
  openDataCommandMenu(trigger, isGroup ? "outline-group" : "outline-ungroup", isGroup ? "グループ化" : "グループ解除", isGroup
    ? `
      <button type="button" role="menuitem" data-data-menu-action="outline-group"${groupDisabled || disabled}><i data-lucide="folder-open"></i><span>${accessKeyLabelHtml("グループ化(G)...")}</span></button>
      <button type="button" role="menuitem" data-data-menu-action="outline-auto"${disabled}><i data-lucide="sparkles"></i><span>${accessKeyLabelHtml("自動アウトライン(A)")}</span></button>
      <div class="ribbon-data-menu-separator" role="separator"></div>
      <button type="button" role="menuitem" data-data-menu-action="outline-settings"${disabled}><i data-lucide="settings"></i><span>${accessKeyLabelHtml("設定(S)...")}</span></button>`
    : `
      <button type="button" role="menuitem" data-data-menu-action="outline-ungroup"${ungroupDisabled || disabled}><i data-lucide="folder-open"></i><span>${accessKeyLabelHtml("グループ解除(U)...")}</span></button>
      <button type="button" role="menuitem" data-data-menu-action="outline-clear"${clearDisabled || disabled}><i data-lucide="eraser"></i><span>${accessKeyLabelHtml("アウトラインのクリア(C)")}</span></button>
      <div class="ribbon-data-menu-separator" role="separator"></div>
      <button type="button" role="menuitem" data-data-menu-action="outline-settings"${disabled}><i data-lucide="settings"></i><span>${accessKeyLabelHtml("設定(S)...")}</span></button>`);
}



function openDataCommandMenu(trigger, kind, label, innerHtml) {
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonColorPalette();
  closeRibbonBorderMenu();
  closeRibbonUnderlineMenu();
  closeRibbonShapeMenu();
  closeRibbonInkEraserMenu();
  closeShapeFormatMenu();
  closeDataValidationMenu();
  closeWhatIfMenu();
  closeQueryPaneContextMenu();
  closeViewCommandMenu();
  closeDataCommandMenu();
  const menu = document.createElement("div");
  menu.id = "dataCommandMenu";
  menu.className = `ribbon-data-menu data-command-menu is-${kind}`;
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", label);
  menu.innerHTML = innerHtml;
  enhanceCommandMenuAccessKeys(menu);
  document.body.append(menu);
  if (window.lucide) window.lucide.createIcons({ icons });
  const rect = typeof trigger?.getBoundingClientRect === "function"
    ? trigger.getBoundingClientRect()
    : trigger || { left: 0, bottom: 0 };
  positionPopupMenu(menu, rect.left, rect.bottom + 2);
  state.dataCommandMenu = { kind };
  menu.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
}



function closeDataCommandMenu() {
  document.getElementById("dataCommandMenu")?.remove();
  state.dataCommandMenu = null;
}



function closeDataCommandMenuOnOutsidePointer(event) {
  if (!state.dataCommandMenu) return;
  if (event.target.closest?.("#dataCommandMenu")) return;
  if (event.target.closest?.("[data-ribbon-command='data-get-menu'], [data-ribbon-command='data-refresh-menu'], [data-ribbon-command='data-recent-sources'], [data-ribbon-command='data-group'], [data-ribbon-command='data-ungroup']")) return;
  closeDataCommandMenu();
}



function closeDataCommandMenuOnEscape(event) {
  if (event.key !== "Escape" || !document.querySelector("#dataCommandMenu")) return;
  if (commandMenuEscapeBlockedByDialog()) return;
  consumeCommandMenuKey(event);
  closeDataCommandMenu();
}



async function handleDataCommandMenuAction(event) {
  const button = event.currentTarget;
  if (button.disabled || button.classList.contains("is-disabled")) return;
  const action = button.dataset.dataMenuAction;
  const buttonRect = button.getBoundingClientRect();
  const sheetIndex = Number(button.dataset.querySheetIndex);
  const queryId = button.dataset.queryId;
  closeDataCommandMenu();
  if (action === "get-workbook") {
    await chooseWorkbookDataForImport();
    return;
  }
  if (action === "get-folder") {
    await chooseFolderDataForImport();
    return;
  }
  if (action === "get-text") {
    await chooseDelimitedTextForImport();
    return;
  }
  if (action === "get-web") {
    openWebImportDialog();
    return;
  }
  if (action === "get-odata") {
    openODataImportDialog();
    return;
  }
  if (action === "get-sql-server") {
    openConnectionOnlySourceDialog("sqlServer");
    return;
  }
  if (action === "get-access") {
    openConnectionOnlySourceDialog("access");
    return;
  }
  if (action === "get-blank-query") {
    openBlankQueryTransformEditor();
    return;
  }
  if (action === "get-table") {
    openGetDataFromTableRangeDialog();
    return;
  }
  if (action === "get-json" || action === "get-xml") {
    await chooseStructuredDataForImport(action === "get-json" ? "json" : "xml");
    return;
  }
  if (action === "other-data-sources") {
    openOtherDataSourcesMenu(buttonRect);
    return;
  }
  if (action === "recent-sources") {
    const trigger = document.querySelector("[data-ribbon-command='data-recent-sources']");
    if (trigger) openRecentSourcesMenu(trigger);
    else openExistingConnectionsDialog();
    return;
  }
  if (action === "recent-source") {
    selectQueryRange(sheetIndex, queryId);
    return;
  }
  if (action === "recent-source-open") {
    await openRecentDataSource(button.dataset.recentSourceId || "");
    return;
  }
  if (action === "existing-connections") {
    openExistingConnectionsDialog();
    return;
  }
  if (action === "queries-connections") {
    openQueriesConnectionsPane();
    return;
  }
  if (action === "refresh-all") {
    await refreshWorkbookData();
    refreshQueriesConnectionsPane();
    return;
  }
  if (action === "refresh-active") {
    await refreshActiveWorkbookConnection();
    return;
  }
  if (action === "refresh-status") {
    openRefreshStatusDialog();
    return;
  }
  if (action === "cancel-refresh") {
    cancelWorkbookRefreshes();
    return;
  }
  if (action === "connection-properties") {
    openActiveQueryPropertiesDialog();
    return;
  }
  if (action === "workbook-links") {
    openWorkbookLinksDialog();
    return;
  }
  if (action === "outline-group") {
    openOutlineGroupDialog(1);
    return;
  }
  if (action === "outline-ungroup") {
    openOutlineGroupDialog(-1);
    return;
  }
  if (action === "outline-auto") {
    autoOutlineSelection();
    return;
  }
  if (action === "outline-clear") {
    clearOutlineFromSelection();
    return;
  }
  if (action === "outline-settings") {
    openOutlineSettingsDialog();
  }
}



function openBlankQueryTransformEditor() {
  if (!state.model) {
    setStatus("Excel ファイルを開いてからデータを取得してください。");
    return;
  }
  const targetPoint = state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  state.dataDialog = {
    type: "query-transform-editor",
    sourceType: "blankQuery",
    sourceName: "空のクエリ",
    sourceRows: [],
    sourceItemIndex: 0,
    sourceItemName: "空のクエリ",
    sourceItemKind: "blankQuery",
    targetPoint: { ...targetPoint },
    transformSteps: [],
    queryName: nextWorkbookQueryName("空のクエリ"),
  };
  renderQueryTransformEditorDialog();
}



function openConnectionOnlySourceDialog(sourceType = "sqlServer") {
  if (!state.model) {
    setStatus("Excel ファイルを開いてからデータを取得してください。");
    return;
  }
  const isAccess = sourceType === "access";
  const title = isAccess ? "Microsoft Access データベースから" : "SQL Server データベースから";
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  state.dataDialog = { type: "connection-only-source", sourceType };
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog connection-source-dialog" role="dialog" aria-modal="true" aria-labelledby="connectionSourceTitle" data-data-dialog="connection-only-source" data-connection-source-type="${escapeAttr(sourceType)}">
      <div class="excel-mini-titlebar" id="connectionSourceTitle">${escapeHtml(title)}</div>
      <div class="excel-mini-body data-dialog-body">
        <fieldset class="data-dialog-fieldset">
          <legend>接続情報</legend>
          ${isAccess ? `
            <label class="excel-dialog-field"><span class="excel-dialog-caption">ファイル名</span><input type="text" data-data-field="accessFile" data-connection-source-sync placeholder="C:\\Data\\sample.accdb" /></label>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">テーブル名</span><input type="text" data-data-field="tableName" data-connection-source-sync placeholder="Table1" /></label>
          ` : `
            <label class="excel-dialog-field"><span class="excel-dialog-caption">サーバー名</span><input type="text" data-data-field="serverName" data-connection-source-sync placeholder="localhost" /></label>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">データベース名</span><input type="text" data-data-field="databaseName" data-connection-source-sync /></label>
          `}
          <label class="excel-dialog-field"><span class="excel-dialog-caption">コマンドの種類</span>
            <select data-data-field="commandType" data-connection-source-sync>
              <option value="table">テーブル</option>
              <option value="sql">SQL</option>
            </select>
          </label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">コマンド テキスト</span><textarea data-data-field="commandText" data-connection-source-sync rows="4" placeholder="${isAccess ? "SELECT * FROM Table1" : "SELECT * FROM dbo.Table1"}"></textarea></label>
        </fieldset>
        <fieldset class="data-dialog-fieldset">
          <legend>読み込み先</legend>
          <label class="data-dialog-check"><input type="checkbox" data-data-field="addToDataModel" data-connection-source-sync />このデータをデータ モデルに追加する</label>
          <p class="data-dialog-note">ブラウザー版では外部 DB ドライバーを直接実行できないため、ここでは Excel と同じ接続情報を「接続のみ」として保存します。</p>
        </fieldset>
        <div class="data-dialog-summary" data-connection-source-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncConnectionOnlySourceDialog(overlay.querySelector("[data-data-dialog='connection-only-source']"));
  overlay.querySelector("input")?.focus({ preventScroll: true });
}



function handleConnectionOnlySourceDialogSync(event) {
  syncConnectionOnlySourceDialog(event.currentTarget.closest("[data-data-dialog='connection-only-source']"));
}



function connectionOnlySourceDialogOptions(dialog) {
  const sourceType = dialog?.dataset.connectionSourceType === "access" ? "access" : "sqlServer";
  const field = (name) => dialog?.querySelector(`[data-data-field="${name}"]`);
  const commandType = field("commandType")?.value === "sql" ? "sql" : "table";
  return {
    sourceType,
    serverName: String(field("serverName")?.value || "").trim(),
    databaseName: String(field("databaseName")?.value || "").trim(),
    accessFile: String(field("accessFile")?.value || "").trim(),
    tableName: String(field("tableName")?.value || "").trim(),
    commandType,
    commandText: String(field("commandText")?.value || "").trim(),
    addToDataModel: Boolean(field("addToDataModel")?.checked),
  };
}



function syncConnectionOnlySourceDialog(dialog) {
  if (!dialog) return;
  const options = connectionOnlySourceDialogOptions(dialog);
  const summary = dialog.querySelector("[data-connection-source-summary]");
  const validation = connectionOnlySourceValidation(options);
  if (summary) {
    summary.textContent = validation.message;
    summary.classList.toggle("is-warning", !validation.valid);
  }
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = false;
}



function connectionOnlySourceValidation(options) {
  if (options.sourceType === "access") {
    if (!options.accessFile) return { valid: false, message: "Access データベース ファイルを指定してください。" };
    if (options.commandType === "table" && !options.tableName && !options.commandText) return { valid: false, message: "テーブル名またはコマンド テキストを指定してください。" };
  } else {
    if (!options.serverName) return { valid: false, message: "SQL Server のサーバー名を指定してください。" };
    if (!options.databaseName) return { valid: false, message: "データベース名を指定してください。" };
    if (options.commandType === "table" && !options.commandText) return { valid: false, message: "テーブル名または SQL コマンドを指定してください。" };
  }
  if (options.commandType === "sql" && !options.commandText) return { valid: false, message: "SQL コマンドを入力してください。" };
  return { valid: true, message: "接続のみとしてブックに保存します。読み込み先は後から変更できます。" };
}



function applyConnectionOnlySourceDialog(dialog) {
  const options = connectionOnlySourceDialogOptions(dialog);
  const validation = connectionOnlySourceValidation(options);
  if (!validation.valid) {
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    syncConnectionOnlySourceDialog(dialog);
    return;
  }
  const sheet = activeSheet();
  if (!sheet) return;
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.queries = Array.isArray(sheet.queries) ? sheet.queries : [];
  const isAccess = options.sourceType === "access";
  const sourceName = isAccess ? (options.accessFile.split(/[\\/]/).pop() || options.accessFile) : `${options.serverName}/${options.databaseName}`;
  const commandText = options.commandText || options.tableName;
  const query = {
    id: uniqueWorksheetObjectId("query"),
    name: nextWorkbookQueryName(isAccess ? "Access" : "SQL Server"),
    sourceType: options.sourceType,
    sourceName,
    sourceUrl: isAccess ? options.accessFile : options.serverName,
    sourceText: "",
    contentType: isAccess ? "application/vnd.ms-access" : "application/vnd.microsoft.sqlserver",
    connectionString: connectionOnlySourceConnectionString(options),
    commandType: options.commandType,
    commandText,
    sourceItemName: commandText,
    sourceItemKind: options.commandType === "sql" ? "sql" : "table",
    sourceRows: [],
    loadMode: "connectionOnly",
    range: null,
    loadedAt: new Date().toISOString(),
    refreshWithAll: false,
    addToDataModel: options.addToDataModel,
  };
  sheet.queries.push(query);
  if (query.addToDataModel) registerQueryInDataModel(state.activeSheetIndex, query);
  closeDataDialog();
  const range = activeSelectionRange() || { top: 1, left: 1, bottom: 1, right: 1, sheetIndex: state.activeSheetIndex };
  commitDataOperation(sheet, sheetBefore, range, `${query.name} 接続`, `${query.name} を接続のみとして作成しました。`, range);
}



function connectionOnlySourceConnectionString(options) {
  if (options.sourceType === "access") {
    return `OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=${options.accessFile}`;
  }
  return `OLEDB;Provider=MSOLEDBSQL;Data Source=${options.serverName};Initial Catalog=${options.databaseName};Integrated Security=SSPI`;
}



function workbookLinks() {
  const links = [];
  (state.model?.sheets || []).forEach((sheet, sheetIndex) => {
    const sheetId = getSheetId(sheet.name);
    Object.entries(sheet.cells || {}).forEach(([key, cell]) => {
      const point = pointFromCellKey(key, sheetIndex);
      if (!point) return;
      if (cell.hyperlink) {
        links.push({
          id: workbookLinkId("cellHyperlink", sheetIndex, key, cell.hyperlink),
          kind: "cellHyperlink",
          sheetIndex,
          row: point.row,
          col: point.col,
          location: `${sheet.name}!${columnName(point.col)}${point.row}`,
          source: String(cell.hyperlink),
          display: getDisplayForCell(sheet, sheetId, point.row, point.col) || String(cell.hyperlink),
        });
      }
      workbookCellFormulaTexts(cell).forEach((formula) => {
        externalFormulaLinkSources(formula).forEach((source) => {
          links.push({
            id: workbookLinkId("formulaExternal", sheetIndex, key, source),
            kind: "formulaExternal",
            sheetIndex,
            row: point.row,
            col: point.col,
            location: `${sheet.name}!${columnName(point.col)}${point.row}`,
            source,
            display: formula,
          });
        });
      });
    });
    (sheet.images || []).forEach((image) => {
      if (!image?.hyperlink) return;
      links.push({
        id: workbookLinkId("objectHyperlink", sheetIndex, image.id, image.hyperlink),
        kind: "objectHyperlink",
        sheetIndex,
        imageId: image.id,
        location: `${sheet.name}!${image.name || image.id || "Object"}`,
        source: String(image.hyperlink),
        display: image.name || image.shape?.label || "オブジェクト",
      });
    });
    (sheet.queries || []).forEach((query) => {
      if (!workbookQueryHasLinkSource(query)) return;
      const source = workbookQueryLinkSource(query, sheet);
      links.push({
        id: workbookLinkId("query", sheetIndex, query.id, source),
        kind: "query",
        sheetIndex,
        queryId: query.id,
        location: `${sheet.name}${query.range ? `!${rangeToLabel(query.range)}` : ""}`,
        source,
        display: query.name || "Query",
      });
    });
  });
  return links;
}



function pointFromCellKey(key, sheetIndex = state.activeSheetIndex) {
  const [rowText, colText] = String(key || "").split(":");
  const row = Number(rowText);
  const col = Number(colText);
  if (!Number.isFinite(row) || !Number.isFinite(col)) return null;
  return { sheetIndex, row, col };
}



function cellFormulaText(cell) {
  if (typeof cell?.formula === "string" && cell.formula.trim()) return cell.formula.trim();
  if (typeof cell?.raw === "string" && cell.raw.trim().startsWith("=")) return cell.raw.trim();
  return "";
}



function workbookCellFormulaTexts(cell) {
  const formulas = [];
  const mainFormula = cellFormulaText(cell);
  if (mainFormula) formulas.push(mainFormula);
  if (typeof cell?.dataTableFormula === "string" && cell.dataTableFormula.trim()) formulas.push(cell.dataTableFormula.trim());
  return [...new Set(formulas)];
}



function externalFormulaLinkSources(formula) {
  const text = String(formula || "");
  if (!text.startsWith("=") && !text.startsWith("{=")) return [];
  const sources = new Set();
  const workbookPattern = /\[([^\]]+)\]/g;
  let match = workbookPattern.exec(text);
  while (match) {
    if (match[1]) sources.add(match[1]);
    match = workbookPattern.exec(text);
  }
  const urlPattern = /https?:\/\/[^"'\s,)]+/gi;
  match = urlPattern.exec(text);
  while (match) {
    sources.add(match[0]);
    match = urlPattern.exec(text);
  }
  return [...sources];
}



function workbookLinksMatchingSource(link) {
  if (!link) return [];
  const target = normalizeWorkbookLinkSource(link.source);
  const kind = link.kind;
  return workbookLinks().filter((item) => item.kind === kind && normalizeWorkbookLinkSource(item.source) === target);
}



function normalizeWorkbookLinkSource(source) {
  return String(source || "").trim();
}



function workbookQueryHasLinkSource(query) {
  if (!query) return false;
  if (["tableRange", "blankQuery", "queryReference", "queryAppend", "queryMerge"].includes(query.sourceType)) return false;
  return Boolean(query.sourceUrl || query.sourceName || query.connectionString || query.sourceText);
}



function workbookQueryLinkSource(query, sheet) {
  if (["web", "odata", "json", "xml"].includes(query?.sourceType) && query.sourceUrl) return query.sourceUrl;
  return query.connectionString || query.sourceUrl || query.sourceName || queryCommandText(query, sheet) || query.name || "";
}



function workbookLinkId(kind, sheetIndex, key, source) {
  return `${kind}:${sheetIndex}:${key}:${encodeURIComponent(String(source || "")).slice(0, 240)}`;
}



function workbookLinkKindLabel(kind) {
  switch (kind) {
    case "formulaExternal": return "外部参照";
    case "cellHyperlink": return "セルのリンク";
    case "objectHyperlink": return "オブジェクトのリンク";
    case "query": return "接続";
    default: return "リンク";
  }
}



function openWorkbookLinksDialog(selectedId = "") {
  if (!state.model) {
    setStatus("先に Excel ファイルを開いてください。");
    return;
  }
  const links = workbookLinks();
  const selected = selectedId && links.some((link) => link.id === selectedId)
    ? selectedId
    : links[0]?.id || "";
  state.dataDialog = { type: "workbook-links", selectedId: selected };
  closeDataDialog({ keepState: true });
  const options = links.length
    ? links.map((link) => `<option value="${escapeAttr(link.id)}"${link.id === selected ? " selected" : ""}>${escapeHtml(workbookLinkOptionLabel(link))}</option>`).join("")
    : '<option value="">このブックにはリンクがありません</option>';
  const selectedLink = links.find((link) => link.id === selected) || null;
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog workbook-links-dialog" role="dialog" aria-modal="true" aria-labelledby="workbookLinksTitle" data-data-dialog="workbook-links">
      <div class="excel-mini-titlebar" id="workbookLinksTitle">ブックのリンク</div>
      <div class="excel-mini-body data-dialog-body workbook-links-body">
        <label class="workbook-links-list-wrap">
          <span class="excel-dialog-caption">リンク</span>
          <select class="workbook-links-list" size="9" data-data-field="workbookLinkList" data-workbook-link-list>${options}</select>
        </label>
        <div class="workbook-links-details">
          <span class="excel-dialog-caption">詳細</span>
          <div class="workbook-links-details-box" data-workbook-link-details>${workbookLinkDetailsHtml(selectedLink)}</div>
          <div class="workbook-links-actions">
            <button class="excel-mini-button" type="button" data-data-dialog-action="workbook-link-open" ${selectedLink ? "" : "disabled"}>開く(O)</button>
            <button class="excel-mini-button" type="button" data-data-dialog-action="workbook-link-select" ${workbookLinkCanSelect(selectedLink) ? "" : "disabled"}>場所へ移動(G)</button>
            <button class="excel-mini-button" type="button" data-data-dialog-action="workbook-link-update" ${selectedLink ? "" : "disabled"}>更新(U)</button>
            <button class="excel-mini-button" type="button" data-data-dialog-action="workbook-link-change" ${selectedLink ? "" : "disabled"}>ソースの変更(C)...</button>
            <button class="excel-mini-button" type="button" data-data-dialog-action="workbook-link-break" ${selectedLink ? "" : "disabled"}>リンクの解除(B)</button>
          </div>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">閉じる</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("[data-workbook-link-list]")?.focus({ preventScroll: true });
}



function workbookLinkOptionLabel(link) {
  if (!link) return "";
  const source = String(link.source || "").replace(/\s+/g, " ");
  return `${workbookLinkKindLabel(link.kind)} - ${link.location} - ${source}`;
}



function workbookLinkDetailsHtml(link) {
  if (!link) return "このブックにはリンクがありません。";
  const matches = workbookLinksMatchingSource(link);
  const locations = matches.map((item) => item.location).filter(Boolean);
  const lines = [
    `種類: ${workbookLinkKindLabel(link.kind)}`,
    `件数: ${matches.length.toLocaleString("ja-JP")}`,
    `場所: ${link.location}`,
    `ソース: ${link.source}`,
    `表示: ${link.display || ""}`,
    `状態: ${workbookLinkStatusLabel(link)}`,
  ];
  if (locations.length > 1) {
    lines.push("リンク先:");
    locations.slice(0, 12).forEach((location) => lines.push(`  ${location}`));
    if (locations.length > 12) lines.push(`  他 ${locations.length - 12} 件`);
  }
  return escapeHtml(lines.join("\n"));
}



function workbookLinkStatusLabel(link) {
  if (!link) return "";
  if (link.kind === "formulaExternal") return "外部ブック参照";
  if (link.kind === "query") return "接続情報あり";
  if (String(link.source || "").startsWith("#")) return "このドキュメント内";
  if (/^mailto:/i.test(link.source || "")) return "電子メール アドレス";
  if (/^(?:https?:|file:)/i.test(link.source || "")) return "リンク先あり";
  return "確認が必要";
}



function workbookLinkCanSelect(link) {
  return Boolean(link && (link.kind === "cellHyperlink" || link.kind === "formulaExternal" || link.kind === "query" || link.kind === "objectHyperlink"));
}



function selectedWorkbookLinkFromDialog(dialog) {
  const selectedId = dialog?.querySelector("[data-workbook-link-list]")?.value || state.dataDialog?.selectedId || "";
  state.dataDialog = state.dataDialog && typeof state.dataDialog === "object" ? state.dataDialog : {};
  state.dataDialog.selectedId = selectedId;
  return workbookLinks().find((link) => link.id === selectedId) || null;
}



function handleWorkbookLinksDialogSync(event) {
  const dialog = event.target.closest("[data-data-dialog='workbook-links']");
  if (!dialog) return;
  const link = selectedWorkbookLinkFromDialog(dialog);
  const details = dialog.querySelector("[data-workbook-link-details]");
  if (details) details.innerHTML = workbookLinkDetailsHtml(link);
  dialog.querySelectorAll("[data-data-dialog-action^='workbook-link-']").forEach((button) => {
    if (button.dataset.dataDialogAction === "workbook-link-select") {
      button.disabled = !workbookLinkCanSelect(link);
      return;
    }
    button.disabled = !link;
  });
}



async function handleWorkbookLinksDialogAction(dialog, action) {
  const link = selectedWorkbookLinkFromDialog(dialog);
  if (!link) return;
  if (action === "workbook-link-select") {
    selectWorkbookLink(link);
    return;
  }
  if (action === "workbook-link-open") {
    openWorkbookLink(link);
    return;
  }
  if (action === "workbook-link-update") {
    await updateWorkbookLink(link);
    openWorkbookLinksDialog(link.id);
    return;
  }
  if (action === "workbook-link-change") {
    openWorkbookLinkChangeDialog(link);
    return;
  }
  if (action === "workbook-link-break") {
    confirmBreakWorkbookLink(link);
  }
}



function selectWorkbookLink(link) {
  if (!link) return;
  closeDataDialog();
  state.activeSheetIndex = link.sheetIndex;
  if (link.kind === "objectHyperlink") {
    state.selectedImage = { id: link.imageId, ids: [link.imageId], groupId: null };
    state.selectionRange = null;
    state.selectionRanges = [];
  } else if (link.kind === "query") {
    selectQueryRange(link.sheetIndex, link.queryId);
    return;
  } else {
    state.selected = { sheetIndex: link.sheetIndex, row: link.row, col: link.col };
    state.selectionRange = selectionRangeFromPoints(state.selected, state.selected);
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectedImage = null;
  }
  renderWorkbook();
  if (state.selectionRange) scrollRangeIntoView(state.selectionRange);
  updateFormulaBarForSelection();
  setStatus(`${link.location} へ移動しました。`);
}



function openWorkbookLink(link) {
  if (!link) return;
  if (link.kind === "query") {
    openQueryPropertiesDialog(link.sheetIndex, link.queryId);
    return;
  }
  if (openInternalHyperlink(link.source)) {
    closeDataDialog();
    setStatus(`リンク先へ移動しました：${link.source}`);
    return;
  }
  if (/^(?:https?:|mailto:|file:)/i.test(link.source || "")) {
    window.open(link.source, "_blank", "noopener,noreferrer");
    setStatus(`リンクを開きました：${link.source}`);
    return;
  }
  showMessageBox("リンク先を開けません。", { title: "Cht WebSheet" });
}



async function updateWorkbookLink(link) {
  if (!link) return;
  if (link.kind === "query") {
    const targets = workbookLinksMatchingSource(link);
    const links = targets.length ? targets : [link];
    for (const item of links) {
      await refreshSingleQuery(item.sheetIndex, item.queryId);
    }
    return;
  }
  if (link.kind === "formulaExternal") {
    state.hf = buildFormulaEngine(state.model);
    renderWorkbook();
    setStatus("外部参照を再計算しました。");
    return;
  }
  setStatus(`${workbookLinkKindLabel(link.kind)} の状態を確認しました。`);
}



function openWorkbookLinkChangeDialog(link) {
  if (!link) return;
  state.dataDialog = {
    type: "workbook-link-change",
    linkId: link.id,
    previousType: "workbook-links",
  };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog workbook-link-change-dialog" role="dialog" aria-modal="true" aria-labelledby="workbookLinkChangeTitle" data-data-dialog="workbook-link-change">
      <div class="excel-mini-titlebar" id="workbookLinkChangeTitle">ソースの変更</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-note">${escapeHtml(workbookLinkKindLabel(link.kind))}: ${escapeHtml(link.location)}</div>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">新しいソース</span><input type="text" data-data-field="workbookLinkSource" value="${escapeAttr(link.source || "")}" /></label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  const input = overlay.querySelector("[data-data-field='workbookLinkSource']");
  input?.focus({ preventScroll: true });
  input?.select?.();
}



function applyWorkbookLinkSourceChange(dialog) {
  const linkId = state.dataDialog?.linkId || "";
  const link = workbookLinks().find((item) => item.id === linkId);
  if (!link) {
    showMessageBox("リンクが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const nextSource = String(dialog.querySelector("[data-data-field='workbookLinkSource']")?.value || "").trim();
  if (!nextSource) {
    showMessageBox("ソースを入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const affected = workbookLinkAffectedSheetIndexes(link);
  const sheetsBefore = captureSheetsForHistory(affected);
  const changed = changeWorkbookLinkSource(link, nextSource);
  if (!changed) {
    showMessageBox("リンクのソースを変更できませんでした。", { title: "Cht WebSheet" });
    return;
  }
  state.hf = buildFormulaEngine(state.model);
  pushHistory({
    sheetIndex: link.sheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(affected),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    label: "ブックのリンク ソース変更",
  });
  scheduleWorkbookPersistence();
  openWorkbookLinksDialog();
  setStatus("リンクのソースを変更しました。");
}



function changeWorkbookLinkSource(link, nextSource) {
  const targets = workbookLinksMatchingSource(link);
  const links = targets.length ? targets : [link];
  let changed = false;
  if (link.kind === "cellHyperlink") {
    links.forEach((item) => {
      const cell = state.model?.sheets?.[item.sheetIndex]?.cells?.[cellKey(item.row, item.col)];
      if (!cell) return;
      cell.hyperlink = nextSource;
      changed = true;
    });
    return changed;
  }
  if (link.kind === "objectHyperlink") {
    links.forEach((item) => {
      const image = state.model?.sheets?.[item.sheetIndex]?.images?.find((imageItem) => imageItem.id === item.imageId);
      if (!image) return;
      image.hyperlink = nextSource;
      changed = true;
    });
    return changed;
  }
  if (link.kind === "query") {
    links.forEach((item) => {
      const query = workbookQueryRef(item.sheetIndex, item.queryId)?.query;
      if (!query) return;
      applyQuerySourceChange(query, nextSource, { clearCache: querySourceChangeShouldClearCache(query) });
      query.loadedAt = new Date().toISOString();
      changed = true;
    });
    return changed;
  }
  if (link.kind === "formulaExternal") {
    (state.model?.sheets || []).forEach((sheet) => {
      Object.values(sheet.cells || {}).forEach((cell) => {
        const formula = cellFormulaText(cell);
        if (formula) {
          const updated = replaceWorkbookLinkSourceInFormula(formula, link.source, nextSource);
          if (updated !== formula) {
            cell.formula = updated;
            cell.raw = updated;
            cell.cached = null;
            cell.display = "";
            changed = true;
          }
        }
        if (typeof cell.dataTableFormula === "string" && cell.dataTableFormula) {
          const updated = replaceWorkbookLinkSourceInFormula(cell.dataTableFormula, link.source, nextSource);
          if (updated !== cell.dataTableFormula) {
            cell.dataTableFormula = updated;
            changed = true;
          }
        }
      });
    });
    return changed;
  }
  return false;
}



function replaceWorkbookLinkSourceInFormula(formula, previousSource, nextSource) {
  const previous = String(previousSource || "");
  if (!previous) return formula;
  let updated = String(formula || "");
  updated = updated.replaceAll(`[${previous}]`, `[${nextSource}]`);
  if (/^(?:https?:|file:|\\\\|[A-Za-z]:[\\/])/.test(previous)) {
    updated = updated.replaceAll(previous, nextSource);
  }
  return updated;
}



function confirmBreakWorkbookLink(link) {
  showMessageBox("リンクを解除すると、外部参照の数式は現在の値に置き換えられ、リンク情報は削除されます。この操作は元に戻せます。続行しますか?", {
    title: "Cht WebSheet",
    tone: "warning",
    buttons: [
      { label: "リンクの解除", result: "break", callback: () => breakWorkbookLink(link) },
      { label: "キャンセル", result: "cancel" },
    ],
    cancelResult: "cancel",
  });
}



function breakWorkbookLink(link) {
  if (!link) return;
  const affected = workbookLinkAffectedSheetIndexes(link);
  const sheetsBefore = captureSheetsForHistory(affected);
  const dataModelBefore = workbookLinkTouchesDataModel(link) ? cloneWorkbookDataModel() : null;
  const changed = removeWorkbookLink(link);
  if (!changed) {
    showMessageBox("リンクを解除できませんでした。", { title: "Cht WebSheet" });
    return;
  }
  state.hf = buildFormulaEngine(state.model);
  pushHistory({
    sheetIndex: link.sheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(affected),
    ...(dataModelBefore ? { dataModelBefore, dataModelAfter: cloneWorkbookDataModel() } : {}),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    label: "ブックのリンク解除",
  });
  scheduleWorkbookPersistence();
  openWorkbookLinksDialog();
  setStatus("リンクを解除しました。");
}



function workbookLinkAffectedSheetIndexes(link) {
  if (link?.kind === "formulaExternal") return (state.model?.sheets || []).map((_sheet, index) => index);
  const indexes = workbookLinksMatchingSource(link).map((item) => item.sheetIndex);
  if (!indexes.length) indexes.push(link?.sheetIndex ?? state.activeSheetIndex);
  return [...new Set(indexes)].filter((index) => state.model?.sheets?.[index]);
}



function workbookLinkTouchesDataModel(link) {
  if (link?.kind !== "query") return false;
  return workbookLinksMatchingSource(link).some((item) => workbookQueryRef(item.sheetIndex, item.queryId)?.query?.addToDataModel);
}



function removeWorkbookLink(link) {
  const targets = workbookLinksMatchingSource(link);
  const links = targets.length ? targets : [link];
  let changed = false;
  if (link.kind === "cellHyperlink") {
    links.forEach((item) => {
      const cell = state.model?.sheets?.[item.sheetIndex]?.cells?.[cellKey(item.row, item.col)];
      if (!cell?.hyperlink) return;
      cell.hyperlink = null;
      changed = true;
    });
    return changed;
  }
  if (link.kind === "objectHyperlink") {
    links.forEach((item) => {
      const image = state.model?.sheets?.[item.sheetIndex]?.images?.find((imageItem) => imageItem.id === item.imageId);
      if (!image?.hyperlink) return;
      image.hyperlink = "";
      changed = true;
    });
    return changed;
  }
  if (link.kind === "query") {
    links.forEach((item) => {
      const sheet = state.model?.sheets?.[item.sheetIndex];
      if (!sheet?.queries?.length) return;
      const query = sheet.queries.find((queryItem) => queryItem.id === item.queryId);
      const before = sheet.queries.length;
      sheet.queries = sheet.queries.filter((queryItem) => queryItem.id !== item.queryId);
      if (query?.addToDataModel && state.model?.dataModel?.tables) {
        state.model.dataModel.tables = state.model.dataModel.tables.filter((table) => table.queryId !== item.queryId);
      }
      changed = changed || sheet.queries.length !== before;
    });
    return changed;
  }
  if (link.kind === "formulaExternal") {
    (state.model?.sheets || []).forEach((sheet, sheetIndex) => {
      const sheetId = getSheetId(sheet.name);
      Object.entries(sheet.cells || {}).forEach(([key, cell]) => {
        const formulaTexts = workbookCellFormulaTexts(cell);
        if (!formulaTexts.some((formula) => formulaReferencesWorkbookLinkSource(formula, link.source))) return;
        const point = pointFromCellKey(key, sheetIndex);
        const value = point ? getDisplayForCell(sheet, sheetId, point.row, point.col) : cell.display || cell.cached || "";
        cell.formula = null;
        cell.raw = value;
        cell.cached = value;
        cell.display = String(value ?? "");
        cell.dataTableFormula = "";
        changed = true;
      });
    });
    return changed;
  }
  return false;
}



function formulaReferencesWorkbookLinkSource(formula, source) {
  const text = String(formula || "");
  const previous = String(source || "");
  if (!text || !previous) return false;
  return externalFormulaLinkSources(text).some((item) => normalizeWorkbookLinkSource(item) === normalizeWorkbookLinkSource(previous));
}



function openWebImportDialog(options = {}) {
  if (!state.model) {
    setStatus("Excel ファイルを開いてからデータを取得してください。");
    return;
  }
  const targetPoint = state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  state.dataDialog = { type: "web-import", targetPoint: { ...targetPoint } };
  closeDataDialog({ keepState: true });
  const destination = `${activeSheet()?.name || ""}!${columnName(targetPoint.col)}${targetPoint.row}`;
  const sourceUrl = options.sourceUrl || "";
  const authMode = options.sourceAuthMode === "basic" ? "basic" : "anonymous";
  const authUsername = options.sourceAuthUsername || "";
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog web-import-dialog" role="dialog" aria-modal="true" aria-labelledby="webImportTitle" data-data-dialog="web-import">
      <div class="excel-mini-titlebar" id="webImportTitle">Web から</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">URL</span><input type="text" data-data-field="sourceUrl" value="${escapeAttr(sourceUrl)}" placeholder="https://example.com/table.html" /></label>
        <fieldset class="data-dialog-fieldset web-import-access">
          <legend>アクセス</legend>
          <label class="data-dialog-check"><input type="radio" name="webImportAuth" value="anonymous" data-import-auth-sync ${authMode === "anonymous" ? "checked" : ""} />匿名</label>
          <label class="data-dialog-check"><input type="radio" name="webImportAuth" value="basic" data-import-auth-sync ${authMode === "basic" ? "checked" : ""} />基本認証</label>
        </fieldset>
        <div class="web-import-basic-auth hidden" data-import-auth-basic-fields>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">ユーザー名</span><input type="text" data-data-field="authUsername" data-import-auth-sync autocomplete="username" value="${escapeAttr(authUsername)}" /></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">パスワード</span><input type="password" data-data-field="authPassword" data-import-auth-sync autocomplete="current-password" /></label>
        </div>
        <div class="data-dialog-note">読み込み先: ${escapeHtml(destination)}</div>
        <div class="data-dialog-note">HTML テーブル、JSON、XML、CSV/TSV テキストを読み込みます。複数の表がある場合はナビゲーターで選択できます。</div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="transform">データの変換</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="load">読み込み</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncImportAuthDialog(overlay.querySelector("[data-data-dialog='web-import']"));
  const focusTarget = authMode === "basic" && sourceUrl
    ? overlay.querySelector("[data-data-field='authPassword']")
    : overlay.querySelector("[data-data-field='sourceUrl']");
  focusTarget?.focus({ preventScroll: true });
  if (focusTarget?.matches?.("[data-data-field='sourceUrl']")) focusTarget.select?.();
}



function openODataImportDialog(options = {}) {
  if (!state.model) {
    setStatus("Excel ファイルを開いてからデータを取得してください。");
    return;
  }
  const targetPoint = state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  state.dataDialog = { type: "odata-import", targetPoint: { ...targetPoint } };
  closeDataDialog({ keepState: true });
  const destination = `${activeSheet()?.name || ""}!${columnName(targetPoint.col)}${targetPoint.row}`;
  const sourceUrl = options.sourceUrl || "";
  const authMode = options.sourceAuthMode === "basic" ? "basic" : "anonymous";
  const authUsername = options.sourceAuthUsername || "";
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog web-import-dialog odata-import-dialog" role="dialog" aria-modal="true" aria-labelledby="odataImportTitle" data-data-dialog="odata-import">
      <div class="excel-mini-titlebar" id="odataImportTitle">OData フィードから</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">URL</span><input type="text" data-data-field="sourceUrl" value="${escapeAttr(sourceUrl)}" placeholder="https://services.odata.org/V4/Northwind/Northwind.svc/Products" /></label>
        <fieldset class="data-dialog-fieldset web-import-access">
          <legend>アクセス</legend>
          <label class="data-dialog-check"><input type="radio" name="odataImportAuth" value="anonymous" data-import-auth-sync ${authMode === "anonymous" ? "checked" : ""} />匿名</label>
          <label class="data-dialog-check"><input type="radio" name="odataImportAuth" value="basic" data-import-auth-sync ${authMode === "basic" ? "checked" : ""} />基本認証</label>
          <label class="data-dialog-check"><input type="radio" name="odataImportAuth" value="organization" data-import-auth-sync disabled />組織アカウント</label>
        </fieldset>
        <div class="web-import-basic-auth hidden" data-import-auth-basic-fields>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">ユーザー名</span><input type="text" data-data-field="authUsername" data-import-auth-sync autocomplete="username" value="${escapeAttr(authUsername)}" /></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">パスワード</span><input type="password" data-data-field="authPassword" data-import-auth-sync autocomplete="current-password" /></label>
        </div>
        <div class="data-dialog-note">読み込み先: ${escapeHtml(destination)}</div>
        <div class="data-dialog-note">OData JSON の value 配列、d.results、複数のエンティティ セットをナビゲーターで選択できます。</div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="transform">データの変換</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="load">読み込み</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncImportAuthDialog(overlay.querySelector("[data-data-dialog='odata-import']"));
  const focusTarget = authMode === "basic" && sourceUrl
    ? overlay.querySelector("[data-data-field='authPassword']")
    : overlay.querySelector("[data-data-field='sourceUrl']");
  focusTarget?.focus({ preventScroll: true });
  if (focusTarget?.matches?.("[data-data-field='sourceUrl']")) focusTarget.select?.();
}



function handleImportAuthDialogSync(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='web-import'], [data-data-dialog='odata-import']");
  syncImportAuthDialog(dialog);
}



function syncImportAuthDialog(dialog) {
  if (!dialog) return;
  const auth = importAuthOptionsFromDialog(dialog);
  const fields = dialog.querySelector("[data-import-auth-basic-fields]");
  fields?.classList.toggle("hidden", auth.mode !== "basic");
  fields?.querySelectorAll("input").forEach((input) => {
    input.disabled = auth.mode !== "basic";
  });
}



function importAuthOptionsFromDialog(dialog) {
  const type = dialog?.dataset?.dataDialog || "";
  const groupName = type === "odata-import" ? "odataImportAuth" : "webImportAuth";
  const mode = dialog?.querySelector(`input[name="${groupName}"]:checked`)?.value || "anonymous";
  const username = String(dialog?.querySelector("[data-data-field='authUsername']")?.value || "");
  const password = String(dialog?.querySelector("[data-data-field='authPassword']")?.value || "");
  return { mode: mode === "basic" ? "basic" : "anonymous", username, password };
}



function validateImportAuthOptions(auth) {
  if (auth?.mode !== "basic") return true;
  if (String(auth.username || "").trim()) return true;
  showMessageBox("基本認証のユーザー名を入力してください。", { title: "Cht WebSheet" });
  return false;
}



function basicAuthHeader(username, password) {
  const credentials = `${String(username || "")}:${String(password || "")}`;
  const encoded = btoa(unescape(encodeURIComponent(credentials)));
  return `Basic ${encoded}`;
}



function requestHeadersForImportAuth(auth, baseHeaders = {}) {
  const headers = { ...(baseHeaders || {}) };
  if (auth?.mode === "basic") {
    headers.Authorization = basicAuthHeader(auth.username, auth.password);
  }
  return headers;
}



function storeQueryCredential(auth) {
  if (auth?.mode !== "basic") return "";
  state.queryCredentials = state.queryCredentials instanceof Map ? state.queryCredentials : new Map();
  const credentialKey = uniqueWorksheetObjectId("credential");
  state.queryCredentials.set(credentialKey, {
    mode: "basic",
    username: String(auth.username || ""),
    password: String(auth.password || ""),
  });
  return credentialKey;
}



function credentialMetadataFromAuth(auth) {
  if (auth?.mode !== "basic") return {};
  return {
    sourceAuthMode: "basic",
    sourceAuthUsername: String(auth.username || ""),
    sourceCredentialKey: storeQueryCredential(auth),
  };
}



function queryCredentialAuthOptions(query) {
  if (query?.sourceAuthMode !== "basic") return null;
  const credential = query?.sourceCredentialKey ? state.queryCredentials?.get?.(query.sourceCredentialKey) : null;
  if (!credential?.password && credential?.password !== "") {
    throw new Error("基本認証の資格情報が現在のセッションにありません。データ ソースへ再接続してください。");
  }
  return {
    mode: "basic",
    username: credential?.username || query.sourceAuthUsername || "",
    password: credential.password,
  };
}



async function applyWebImportDialog(dialog) {
  const source = await fetchWebImportDialogSource(dialog);
  if (!source) return;
  if (source.sources.length > 1) {
    openWebNavigatorDialog(source);
    return;
  }
  loadWebRowsIntoSelection(source.rows, {
    sourceUrl: source.sourceUrl,
    sourceText: source.sourceText,
    contentType: source.contentType,
    sourceAuthMode: source.sourceAuthMode,
    sourceAuthUsername: source.sourceAuthUsername,
    sourceCredentialKey: source.sourceCredentialKey,
    sourceItemIndex: source.sourceItemIndex,
    sourceItemName: source.sourceItemName,
    sourceItemKind: source.sourceItemKind,
    targetPoint: source.targetPoint,
  });
}



async function applyODataImportDialog(dialog) {
  const source = await fetchODataImportDialogSource(dialog);
  if (!source) return;
  if (source.sources.length > 1) {
    openWebNavigatorDialog(source);
    return;
  }
  loadWebRowsIntoSelection(source.rows, {
    sourceType: source.sourceType,
    sourceName: source.sourceName,
    sourceUrl: source.sourceUrl,
    sourceText: source.sourceText,
    contentType: source.contentType,
    sourceAuthMode: source.sourceAuthMode,
    sourceAuthUsername: source.sourceAuthUsername,
    sourceCredentialKey: source.sourceCredentialKey,
    sourceItemIndex: source.sourceItemIndex,
    sourceItemName: source.sourceItemName,
    sourceItemKind: source.sourceItemKind,
    targetPoint: source.targetPoint,
  });
}



async function openWebTransformEditor(dialog) {
  const source = await fetchWebImportDialogSource(dialog);
  if (!source) return;
  if (source.sources.length > 1) {
    openWebNavigatorDialog(source);
    return;
  }
  openWebNavigatorTransformEditorFromSource(source, source.sources[0]);
}



async function openODataTransformEditor(dialog) {
  const source = await fetchODataImportDialogSource(dialog);
  if (!source) return;
  if (source.sources.length > 1) {
    openWebNavigatorDialog(source);
    return;
  }
  openWebNavigatorTransformEditorFromSource(source, source.sources[0]);
}



function openWebNavigatorTransformEditorFromSource(source, selectedSource) {
  const sourceType = source.sourceType || "web";
  const sourceName = source.sourceName || source.sourceUrl || "Web";
  const shouldPersistRows = querySourceTypeUsesStoredRows(sourceType);
  state.dataDialog = {
    type: "query-transform-editor",
    sourceType,
    sourceName: selectedSource?.name ? `${sourceName} - ${selectedSource.name}` : sourceName,
    sourceUrl: source.sourceUrl,
    sourceText: source.sourceText,
    contentType: source.contentType,
    sourceAuthMode: source.sourceAuthMode || "",
    sourceAuthUsername: source.sourceAuthUsername || "",
    sourceCredentialKey: source.sourceCredentialKey || "",
    sourceRows: shouldPersistRows ? (selectedSource?.rows || source.sourceRows || []) : source.sourceRows || null,
    sourceItemIndex: selectedSource?.itemIndex ?? source.sourceItemIndex ?? 0,
    sourceItemName: selectedSource?.name || source.sourceItemName || "",
    sourceItemKind: selectedSource?.kind || source.sourceItemKind || "document",
    targetPoint: { ...source.targetPoint },
    transformSteps: [],
  };
  renderQueryTransformEditorDialog();
}



async function fetchWebImportDialogSource(dialog) {
  const sourceUrl = String(dialog.querySelector("[data-data-field='sourceUrl']")?.value || "").trim();
  if (!sourceUrl) {
    showMessageBox("URL を入力してください。", { title: "Cht WebSheet" });
    return null;
  }
  const auth = importAuthOptionsFromDialog(dialog);
  if (!validateImportAuthOptions(auth)) return null;
  const buttons = [...dialog.querySelectorAll("[data-data-dialog-action]")];
  buttons.forEach((button) => {
    button.disabled = true;
  });
  try {
    const response = await fetch(sourceUrl, { headers: requestHeadersForImportAuth(auth) });
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`.trim());
    const contentType = response.headers.get("content-type") || "";
    const sourceText = await response.text();
    const sources = webImportSourcesFromSource(sourceText, contentType, sourceUrl);
    const selectedSource = sources[0] || null;
    const rows = selectedSource?.rows || [];
    if (!rows.length) {
      showMessageBox("読み込める表データが見つかりません。", { title: "Cht WebSheet" });
      return null;
    }
    const point = state.dataDialog?.targetPoint || state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
    return {
      sourceUrl,
      sourceText,
      contentType,
      sources,
      rows,
      ...credentialMetadataFromAuth(auth),
      sourceItemIndex: selectedSource?.itemIndex ?? 0,
      sourceItemName: selectedSource?.name || "",
      sourceItemKind: selectedSource?.kind || "document",
      targetPoint: point,
    };
  } catch (error) {
    console.error(error);
    showMessageBox("Web データを読み込めませんでした。URL とアクセス権を確認してください。", { title: "Cht WebSheet" });
    return null;
  } finally {
    buttons.forEach((button) => {
      button.disabled = false;
    });
  }
}



async function fetchODataImportDialogSource(dialog) {
  const sourceUrl = String(dialog.querySelector("[data-data-field='sourceUrl']")?.value || "").trim();
  if (!sourceUrl) {
    showMessageBox("URL を入力してください。", { title: "Cht WebSheet" });
    return null;
  }
  const auth = importAuthOptionsFromDialog(dialog);
  if (!validateImportAuthOptions(auth)) return null;
  const buttons = [...dialog.querySelectorAll("[data-data-dialog-action]")];
  buttons.forEach((button) => {
    button.disabled = true;
  });
  try {
    const response = await fetch(sourceUrl, {
      headers: requestHeadersForImportAuth(auth, { accept: "application/json, application/xml, text/xml, text/csv, */*" }),
    });
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`.trim());
    const contentType = response.headers.get("content-type") || odataContentTypeFromUrl(sourceUrl);
    const sourceText = await response.text();
    const sources = odataImportSourcesFromSource(sourceText, contentType, sourceUrl);
    const selectedSource = sources[0] || null;
    const rows = selectedSource?.rows || [];
    if (!rows.length) {
      showMessageBox("OData フィードから読み込めるデータが見つかりません。", { title: "Cht WebSheet" });
      return null;
    }
    const point = state.dataDialog?.targetPoint || state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
    return {
      sourceType: "odata",
      sourceName: sourceUrl,
      sourceUrl,
      sourceText,
      contentType,
      sources,
      rows,
      ...credentialMetadataFromAuth(auth),
      sourceItemIndex: selectedSource?.itemIndex ?? 0,
      sourceItemName: selectedSource?.name || "",
      sourceItemKind: selectedSource?.kind || "odata",
      targetPoint: point,
    };
  } catch (error) {
    console.error(error);
    showMessageBox("OData フィードを読み込めませんでした。URL とアクセス権を確認してください。", { title: "Cht WebSheet" });
    return null;
  } finally {
    buttons.forEach((button) => {
      button.disabled = false;
    });
  }
}



function webImportRowsFromSource(sourceText, contentType = "", sourceUrl = "", sourceItemIndex = 0) {
  const sources = webImportSourcesFromSource(sourceText, contentType, sourceUrl);
  const numericIndex = Math.max(0, Math.round(Number(sourceItemIndex) || 0));
  const source = sources.find((item) => Number(item.itemIndex) === numericIndex) || sources[numericIndex] || sources[0];
  return source?.rows || [];
}



function odataImportRowsFromSource(sourceText, contentType = "", sourceUrl = "", sourceItemIndex = 0) {
  const sources = odataImportSourcesFromSource(sourceText, contentType, sourceUrl);
  const numericIndex = Math.max(0, Math.round(Number(sourceItemIndex) || 0));
  const source = sources.find((item) => Number(item.itemIndex) === numericIndex) || sources[numericIndex] || sources[0];
  return source?.rows || [];
}



function webImportSourcesFromSource(sourceText, contentType = "", sourceUrl = "") {
  const text = stripTextBom(sourceText);
  if (/html/i.test(contentType) || /<table[\s>]/i.test(text)) {
    const tableSources = htmlTableSources(text);
    if (tableSources.length) return tableSources;
  }
  if (/json/i.test(contentType) || /^[\s\r\n]*[\[{]/.test(text)) {
    const jsonSources = jsonImportSourcesFromSource(text);
    if (jsonSources.length) return jsonSources;
  }
  if (/xml/i.test(contentType) || /^[\s\r\n]*(?:<\?xml\b|<[\w:-]+[\s>/])/.test(text)) {
    const xmlSources = xmlImportSourcesFromSource(text);
    if (xmlSources.length) return xmlSources;
  }
  const delimiter = detectDelimitedTextDelimiter(text, sourceUrl || "Web");
  const rows = parseDelimitedRows(text, delimiter).filter((row, index, list) => index < list.length - 1 || row.some((value) => value !== ""));
  return rows.length
    ? [{ kind: "document", itemIndex: 0, name: "Document", rows }]
    : [];
}



function odataImportSourcesFromSource(sourceText, contentType = "", sourceUrl = "") {
  const text = stripTextBom(sourceText);
  const sources = [];
  if (/json/i.test(contentType) || /^[\s\r\n]*[\[{]/.test(text)) {
    try {
      const value = JSON.parse(text);
      if (Array.isArray(value)) {
        odataPushSource(sources, "Document", value);
      } else if (isPlainObject(value)) {
        const contextName = odataEntitySetNameFromContext(value["@odata.context"] || value["odata.metadata"]);
        if (Array.isArray(value.value)) odataPushSource(sources, contextName || odataEntitySetNameFromUrl(sourceUrl) || "value", value.value);
        const d = value.d;
        if (isPlainObject(d) && Array.isArray(d.results)) odataPushSource(sources, contextName || odataEntitySetNameFromUrl(sourceUrl) || "results", d.results);
        if (isPlainObject(d)) {
          Object.entries(d).forEach(([key, entry]) => {
            const rowsValue = Array.isArray(entry) ? entry : isPlainObject(entry) && Array.isArray(entry.results) ? entry.results : null;
            if (rowsValue) odataPushSource(sources, key, rowsValue);
          });
        }
        Object.entries(value).forEach(([key, entry]) => {
          if (key === "value" || key === "d" || key.startsWith("@odata.") || key.startsWith("odata.")) return;
          const rowsValue = Array.isArray(entry) ? entry : isPlainObject(entry) && Array.isArray(entry.value) ? entry.value : null;
          if (rowsValue) odataPushSource(sources, key, rowsValue);
        });
        if (!sources.length) odataPushSource(sources, contextName || odataEntitySetNameFromUrl(sourceUrl) || "Document", value);
      }
    } catch (error) {
      console.warn("Failed to parse OData source", error);
    }
  }
  if (sources.length) return sources.map((source, index) => ({ ...source, itemIndex: index }));
  return webImportSourcesFromSource(text, contentType, sourceUrl).map((source, index) => ({
    ...source,
    kind: source.kind === "json" || source.kind === "document" ? "odata" : source.kind,
    itemIndex: index,
    name: source.name === "Document" ? (odataEntitySetNameFromUrl(sourceUrl) || source.name) : source.name,
  }));
}



function odataPushSource(sources, name, value) {
  const rows = odataRowsFromValue(value);
  if (!rows.some((row) => row.some((cell) => String(cell ?? "") !== ""))) return;
  const normalizedName = String(name || `Table ${sources.length + 1}`).trim() || `Table ${sources.length + 1}`;
  if (sources.some((source) => source.name === normalizedName && JSON.stringify(source.rows) === JSON.stringify(rows))) return;
  sources.push({ kind: "odata", itemIndex: sources.length, name: normalizedName, rows });
}



function odataRowsFromValue(value) {
  if (Array.isArray(value) && value.every((item) => isPlainObject(item))) {
    return rowsFromObjects(value.map((item) => flattenODataObject(item)));
  }
  if (isPlainObject(value)) return rowsFromObjects([flattenODataObject(value)]);
  return jsonRowsFromValue(value);
}



function flattenODataObject(value, prefix = "", result = {}) {
  if (!isPlainObject(value)) return result;
  Object.entries(value).forEach(([key, entry]) => {
    if (key === "__metadata" || key.startsWith("@odata.") || key.startsWith("odata.")) return;
    const path = prefix ? `${prefix}.${key}` : key;
    if (isPlainObject(entry) && !Array.isArray(entry.results) && !Array.isArray(entry.value)) {
      flattenODataObject(entry, path, result);
    } else if (isPlainObject(entry) && Array.isArray(entry.results)) {
      result[path] = JSON.stringify(entry.results.map((item) => isPlainObject(item) ? flattenODataObject(item) : item));
    } else if (isPlainObject(entry) && Array.isArray(entry.value)) {
      result[path] = JSON.stringify(entry.value.map((item) => isPlainObject(item) ? flattenODataObject(item) : item));
    } else {
      result[path] = jsonScalarToCell(entry);
    }
  });
  return result;
}



function odataContentTypeFromUrl(sourceUrl) {
  const text = String(sourceUrl || "");
  if (/^data:([^,;]+)/i.test(text)) return text.match(/^data:([^,;]+)/i)?.[1] || "";
  return "";
}



function odataEntitySetNameFromUrl(sourceUrl) {
  try {
    const url = new URL(sourceUrl, window.location.href);
    const pathname = url.pathname.replace(/\/+$/, "");
    const last = pathname.split("/").filter(Boolean).pop() || "";
    return decodeURIComponent(last).replace(/\([^)]*\)$/, "");
  } catch (error) {
    const path = String(sourceUrl || "").split(/[?#]/)[0].replace(/\/+$/, "");
    const last = path.split("/").filter(Boolean).pop() || "";
    return last ? decodeURIComponent(last).replace(/\([^)]*\)$/, "") : "";
  }
}



function odataEntitySetNameFromContext(context) {
  const text = String(context || "");
  if (!text) return "";
  const afterHash = text.includes("#") ? text.slice(text.lastIndexOf("#") + 1) : text;
  const clean = afterHash.replace(/\/\$entity$/i, "").replace(/\([^)]*\)/g, "").replace(/^Collection\((.*)\)$/i, "$1");
  const parts = clean.split(/[/.]/).filter(Boolean);
  const name = parts.pop() || "";
  return name.replace(/^[\w-]+:/, "");
}



function jsonImportSourcesFromSource(sourceText) {
  try {
    const value = JSON.parse(sourceText);
    const sources = [];
    if (Array.isArray(value)) {
      sources.push({ kind: "json", itemIndex: 0, name: "Document", rows: jsonRowsFromValue(value) });
    } else if (isPlainObject(value)) {
      Object.entries(value).forEach(([key, entry]) => {
        if (Array.isArray(entry) || isPlainObject(entry)) {
          const rows = jsonRowsFromValue(entry);
          if (rows.length) sources.push({ kind: "json", itemIndex: sources.length, name: key, rows });
        }
      });
      if (!sources.length) sources.push({ kind: "json", itemIndex: 0, name: "Document", rows: jsonRowsFromValue(value) });
    } else {
      sources.push({ kind: "json", itemIndex: 0, name: "Document", rows: [["Value"], [jsonScalarToCell(value)]] });
    }
    return sources
      .map((source, index) => ({ ...source, itemIndex: index }))
      .filter((source) => source.rows.some((row) => row.some((cell) => String(cell ?? "") !== "")));
  } catch (error) {
    console.warn("Failed to parse JSON source", error);
    return [];
  }
}



function jsonRowsFromValue(value) {
  if (Array.isArray(value)) return jsonRowsFromArray(value);
  if (isPlainObject(value)) return jsonRowsFromObject(value);
  return [["Value"], [jsonScalarToCell(value)]];
}



function jsonRowsFromArray(items) {
  const values = Array.isArray(items) ? items : [];
  if (!values.length) return [];
  if (values.every((item) => isPlainObject(item))) {
    const flattened = values.map((item) => flattenJsonObject(item));
    return rowsFromObjects(flattened);
  }
  if (values.every((item) => Array.isArray(item))) {
    const width = Math.max(...values.map((item) => item.length), 1);
    return [
      ...[Array.from({ length: width }, (_item, index) => `Column${index + 1}`)],
      ...values.map((item) => Array.from({ length: width }, (_unused, index) => jsonScalarToCell(item[index]))),
    ];
  }
  return [["Value"], ...values.map((item) => [jsonScalarToCell(item)])];
}



function jsonRowsFromObject(value) {
  const entries = Object.entries(value || {});
  if (entries.length && entries.every(([, entry]) => isPlainObject(entry))) {
    const flattened = entries.map(([key, entry]) => ({ Key: key, ...flattenJsonObject(entry) }));
    return rowsFromObjects(flattened);
  }
  return rowsFromObjects([flattenJsonObject(value)]);
}



function rowsFromObjects(objects) {
  const rows = Array.isArray(objects) ? objects : [];
  const headers = [];
  rows.forEach((row) => {
    Object.keys(row || {}).forEach((key) => {
      if (!headers.includes(key)) headers.push(key);
    });
  });
  if (!headers.length) return [];
  return [headers, ...rows.map((row) => headers.map((header) => jsonScalarToCell(row?.[header])))];
}



function flattenJsonObject(value, prefix = "", result = {}) {
  if (!isPlainObject(value)) return result;
  Object.entries(value).forEach(([key, entry]) => {
    const path = prefix ? `${prefix}.${key}` : key;
    if (isPlainObject(entry)) {
      flattenJsonObject(entry, path, result);
    } else {
      result[path] = jsonScalarToCell(entry);
    }
  });
  return result;
}



function jsonScalarToCell(value) {
  if (value == null) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return String(value);
  return JSON.stringify(value);
}



function isPlainObject(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}



function xmlImportSourcesFromSource(sourceText) {
  try {
    const doc = new DOMParser().parseFromString(sourceText, "application/xml");
    if (doc.querySelector("parsererror")) return [];
    const root = doc.documentElement;
    if (!root) return [];
    const groups = xmlRepeatedElementGroups(root);
    const sources = groups
      .map((group, index) => ({
        kind: "xml",
        itemIndex: index,
        name: group.name,
        rows: xmlRowsFromElements(group.elements),
      }))
      .filter((source) => source.rows.some((row) => row.some((cell) => String(cell ?? "") !== "")));
    if (sources.length) return sources.map((source, index) => ({ ...source, itemIndex: index }));
    const childElements = [...root.children];
    const rows = childElements.length ? xmlRowsFromElements(childElements) : rowsFromObjects([xmlElementToRowObject(root)]);
    return rows.length ? [{ kind: "xml", itemIndex: 0, name: root.localName || "Document", rows }] : [];
  } catch (error) {
    console.warn("Failed to parse XML source", error);
    return [];
  }
}



function xmlRepeatedElementGroups(root) {
  const groups = [];
  const visit = (element, path = "") => {
    const byName = new Map();
    [...element.children].forEach((child) => {
      const name = child.localName || child.tagName;
      if (!byName.has(name)) byName.set(name, []);
      byName.get(name).push(child);
    });
    byName.forEach((elements, name) => {
      if (elements.length > 1) {
        groups.push({ name: path ? `${path}/${name}` : name, elements });
      }
    });
    [...element.children].forEach((child) => visit(child, path ? `${path}/${child.localName || child.tagName}` : (child.localName || child.tagName)));
  };
  visit(root, root.localName || root.tagName || "Document");
  const seen = new Set();
  return groups.filter((group) => {
    const key = `${group.name}:${group.elements.length}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

export {
  applyConnectionOnlySourceDialog,
  applyGetDataFromTableRangeDialog,
  applyODataImportDialog,
  applyQueryTransformDescriptorToQuery,
  applyWebImportDialog,
  applyWorkbookLinkSourceChange,
  basicAuthHeader,
  breakWorkbookLink,
  canOpenRecentDataSource,
  canRememberRecentDataSource,
  cellFormulaText,
  changeWorkbookLinkSource,
  closeDataCommandMenu,
  closeDataCommandMenuOnEscape,
  closeDataCommandMenuOnOutsidePointer,
  compareQueryTransformValues,
  confirmBreakWorkbookLink,
  connectionOnlySourceConnectionString,
  connectionOnlySourceDialogOptions,
  connectionOnlySourceValidation,
  createQueryFromTransformDescriptor,
  credentialMetadataFromAuth,
  editingWorkbookQueryRef,
  evaluateQueryCustomColumnFormula,
  evaluateQueryMCondition,
  evaluateQueryMFunctionExpression,
  evaluateQueryMIfExpression,
  externalFormulaLinkSources,
  fetchODataImportDialogSource,
  fetchRecentDataSourceText,
  fetchWebImportDialogSource,
  flattenJsonObject,
  flattenODataObject,
  formulaReferencesWorkbookLinkSource,
  handleConnectionOnlySourceDialogSync,
  handleDataCommandMenuAction,
  handleImportAuthDialogSync,
  handleWorkbookLinksDialogAction,
  handleWorkbookLinksDialogSync,
  importAuthOptionsFromDialog,
  isPlainObject,
  jsonImportSourcesFromSource,
  jsonRowsFromArray,
  jsonRowsFromObject,
  jsonRowsFromValue,
  jsonScalarToCell,
  loadQueryTransformEditorToSelection,
  loadRecentDataSources,
  moveQueryTransformArrayItem,
  normalizeWorkbookLinkSource,
  odataContentTypeFromUrl,
  odataEntitySetNameFromContext,
  odataEntitySetNameFromUrl,
  odataImportRowsFromSource,
  odataImportSourcesFromSource,
  odataPushSource,
  odataRowsFromValue,
  openBlankQueryTransformEditor,
  openConnectionOnlySourceDialog,
  openDataCommandMenu,
  openDataGetMenu,
  openDataRefreshMenu,
  openGetDataFromTableRangeDialog,
  openODataImportDialog,
  openODataTransformEditor,
  openOtherDataSourcesMenu,
  openOutlineCommandMenu,
  openRecentDataSource,
  openRecentSourcesMenu,
  openTableRangeTransformEditor,
  openWebImportDialog,
  openWebNavigatorTransformEditorFromSource,
  openWebTransformEditor,
  openWorkbookLink,
  openWorkbookLinkChangeDialog,
  openWorkbookLinksDialog,
  padQueryTransformRow,
  pointFromCellKey,
  queryCredentialAuthOptions,
  queryPreviewHtml,
  queryTransformAggregateDataType,
  queryTransformAggregateNumbers,
  queryTransformAggregateValue,
  queryTransformBooleanValue,
  queryTransformCharacterSetValue,
  queryTransformConditionMatches,
  queryTransformConvertDataTypeValue,
  queryTransformDatePartDataType,
  queryTransformDatePartValue,
  queryTransformDefaultColumns,
  queryTransformFormatDate,
  queryTransformFormatDateTime,
  queryTransformFormatNumber,
  queryTransformFormulaColumnValue,
  queryTransformFormulaTextTerm,
  queryTransformFormulaValueTerm,
  queryTransformInferComparisonType,
  queryTransformInferDataType,
  queryTransformIsErrorValue,
  queryTransformMedian,
  queryTransformMoveTargetIndex,
  queryTransformNormalizeDataType,
  queryTransformNumberOperationValue,
  queryTransformParseDateValue,
  queryTransformRowsForDescriptor,
  queryTransformStepColumnIndexes,
  queryTransformTextCaseValue,
  queryTransformVariance,
  recentDataSourceAuthMetadata,
  recentDataSourceAuthOptions,
  recentDataSourceContentType,
  recentDataSourceDetailLabel,
  recentDataSourceFromQuery,
  recentDataSourceIconName,
  recentDataSourceId,
  recentDataSourceNavigatorSource,
  recentDataSourceNeedsCredential,
  recentDataSourceTypeLabel,
  registerQueryTransformConnectionAndOpenLoadTo,
  rememberRecentDataSourceFromQuery,
  removeWorkbookLink,
  reopenRecentDataSourceCredentialDialog,
  replaceWorkbookLinkSourceInFormula,
  requestHeadersForImportAuth,
  rowsFromObjects,
  saveRecentDataSources,
  selectWorkbookLink,
  selectedWorkbookLinkFromDialog,
  sheetDisplayRowsForRange,
  splitQueryMArguments,
  splitTextByCharacterCount,
  storeQueryCredential,
  syncConnectionOnlySourceDialog,
  syncImportAuthDialog,
  uniqueQueryTransformColumns,
  updateWorkbookLink,
  validateImportAuthOptions,
  webImportRowsFromSource,
  webImportSourcesFromSource,
  workbookCellFormulaTexts,
  workbookLinkAffectedSheetIndexes,
  workbookLinkCanSelect,
  workbookLinkDetailsHtml,
  workbookLinkId,
  workbookLinkKindLabel,
  workbookLinkOptionLabel,
  workbookLinkStatusLabel,
  workbookLinkTouchesDataModel,
  workbookLinks,
  workbookLinksMatchingSource,
  workbookQueryHasLinkSource,
  workbookQueryLinkSource,
  workbookQueryRowsWithTransforms,
  xmlImportSourcesFromSource,
  xmlRepeatedElementGroups,
};
