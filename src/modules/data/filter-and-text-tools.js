// Feature module; shared bindings are imported explicitly.
import {
  AUTO_FILTER_MONTH_NAME_INDEX,
  FLASH_FILL_REGEX_SEGMENTS,
  state,
} from "../../app-core.js";
import {
  syncDataAnalysisDialog,
  syncDataTableDialog,
  syncGoalSeekDialog,
  syncScenarioEditDialog,
  syncScenarioSummaryDialog,
} from "./analysis-and-scenarios.js";
import {
  autoFilterCellMatchesCriterion,
  autoFilterDatePeriodLabel,
  autoFilterRangeFromSelection,
  autoFilterStateForHeaderCell,
  autoFilterStatusMessage,
  comparableDateSerial,
  currentRegionFromPoint,
  normalizeSortColor,
  rangeColumnHasData,
  rangeHasUserData,
  reapplyAutoFilter,
  sheetUsedRange,
  showDataOverwriteConfirmation,
  sortRangeRows,
  sortRangeRowsByKeys,
} from "./sort-and-autofilter.js";
import {
  decodeWorkbookRangeOrName,
  syncConsolidateDialog,
  syncDataValidationDialog,
  translateValidationFormulaReferences,
  validationFormulaResultTruthy,
  workbookCellReference,
  workbookRangeReference,
} from "./validation-and-consolidation.js";
import {
  closeDataDialog,
} from "../drawing/charts-and-insertion.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  coerceUserInput,
  cssEscape,
  getCellRawInput,
  getDisplayForCell,
  isSingleCellRange,
  materializeSelectionRange,
  materializeSelectionRangeForSheet,
  normalizeRange,
  rangeToLabel,
  restoreGroupedSheetsAfterFailure,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  rawActiveSelectionRanges,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneCellModel,
  defaultColumnModel,
  defaultRowModel,
  internCellsInRange,
  pushHistory,
  writeClonedCellContent,
} from "../editing/structure-and-history.js";
import {
  cellKey,
  columnName,
  decodeCell,
  escapeAttr,
  escapeHtml,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  commitDataOperation,
  commitDataOperationForSheet,
  ensureSheetSize,
  selectedDataRange,
} from "../ui/ribbon-and-links.js";
import {
  ensureEditableRangeForProtection,
  ensureSheetProtectionPermission,
} from "../ui/sheet-protection-and-translation.js";
import {
  buildFormulaEngine,
  formulaEngineSheetIndex,
  renderWorkbook,
  translateFormulaForCurrentFormulaEngine,
  validationDefinedName,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  activeSheet,
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
  autoFilterStdDevCount,
  validationReferenceFromFormula,
} from "../workbook/tables-and-validation.js";


function autoFilterConditionMenuHtml(profile, customActive = false) {
  const numeric = Boolean(profile?.numeric);
  const date = Boolean(profile?.date);
  const label = date ? "日付フィルター" : numeric ? "数値フィルター" : "テキスト フィルター";
  const items = date
    ? [
        ["dateToday", "今日"],
        ["dateYesterday", "昨日"],
        ["dateTomorrow", "明日"],
        ["dateThisWeek", "今週"],
        ["dateLastWeek", "先週"],
        ["dateNextWeek", "来週"],
        ["dateThisMonth", "今月"],
        ["dateLastMonth", "先月"],
        ["dateNextMonth", "来月"],
        ["dateThisQuarter", "今四半期"],
        ["dateLastQuarter", "前四半期"],
        ["dateNextQuarter", "次四半期"],
        ["dateThisYear", "今年"],
        ["dateLastYear", "昨年"],
        ["dateNextYear", "来年"],
        ["dateYearToDate", "年初来"],
        ["datePeriodQ1", "第1四半期"],
        ["datePeriodQ2", "第2四半期"],
        ["datePeriodQ3", "第3四半期"],
        ["datePeriodQ4", "第4四半期"],
        ["datePeriodM1", "1月"],
        ["datePeriodM2", "2月"],
        ["datePeriodM3", "3月"],
        ["datePeriodM4", "4月"],
        ["datePeriodM5", "5月"],
        ["datePeriodM6", "6月"],
        ["datePeriodM7", "7月"],
        ["datePeriodM8", "8月"],
        ["datePeriodM9", "9月"],
        ["datePeriodM10", "10月"],
        ["datePeriodM11", "11月"],
        ["datePeriodM12", "12月"],
        ["equal", "指定の日付に等しい..."],
        ["notEqual", "指定の日付に等しくない..."],
        ["greaterThan", "指定の日付より後..."],
        ["greaterThanOrEqual", "指定の日付以降..."],
        ["lessThan", "指定の日付より前..."],
        ["lessThanOrEqual", "指定の日付以前..."],
        ["between", "指定の範囲内..."],
        ["notBetween", "指定の範囲外..."],
        ["blanks", "空白セル"],
        ["nonBlanks", "空白以外のセル"],
        ["custom", "ユーザー設定フィルター..."],
      ]
    : numeric
      ? [
          ["equal", "指定の値に等しい..."],
          ["notEqual", "指定の値に等しくない..."],
        ["greaterThan", "指定の値より大きい..."],
        ["greaterThanOrEqual", "指定の値以上..."],
        ["lessThan", "指定の値より小さい..."],
        ["lessThanOrEqual", "指定の値以下..."],
        ["between", "指定の範囲内..."],
        ["notBetween", "指定の範囲外..."],
        ["top10", "トップ テン..."],
        ["aboveAverage", "平均より上"],
        ["belowAverage", "平均より下"],
        ["aboveStdDev", "標準偏差より上"],
        ["belowStdDev", "標準偏差より下"],
        ["blanks", "空白セル"],
        ["nonBlanks", "空白以外のセル"],
        ["custom", "ユーザー設定フィルター..."],
      ]
    : [
        ["equal", "指定の値に等しい..."],
        ["notEqual", "指定の値に等しくない..."],
        ["beginsWith", "指定の値で始まる..."],
        ["notBeginsWith", "指定の値で始まらない..."],
        ["endsWith", "指定の値で終わる..."],
        ["notEndsWith", "指定の値で終わらない..."],
        ["contains", "指定の値を含む..."],
        ["notContains", "指定の値を含まない..."],
        ["blanks", "空白セル"],
        ["nonBlanks", "空白以外のセル"],
        ["custom", "ユーザー設定フィルター..."],
      ];
  return `
    <div class="auto-filter-condition">
      <button type="button" class="auto-filter-submenu-toggle${customActive ? " is-filtered" : ""}" data-auto-filter-action="condition-menu">
        <span>${escapeHtml(label)}</span><span aria-hidden="true">›</span>
      </button>
      <div class="auto-filter-submenu hidden" role="menu" aria-label="${escapeAttr(label)}">
        ${items.map(([operator, itemLabel]) => `<button type="button" role="menuitem" data-auto-filter-action="condition" data-auto-filter-operator="${escapeAttr(operator)}">${escapeHtml(itemLabel)}</button>`).join("")}
      </div>
    </div>`;
}



function autoFilterColorMenuHtml(profile, active = false) {
  const fillColors = profile?.fillColors || [];
  const fontColors = profile?.fontColors || [];
  if (!fillColors.length && !fontColors.length) return "";
  const fillItems = fillColors.length
    ? `<div class="auto-filter-submenu-heading">セルの色で並べ替え</div>${fillColors.map((item) => autoFilterColorButtonHtml("fill", item, "color-sort")).join("")}
       <div class="auto-filter-submenu-heading">セルの色でフィルター</div>${fillColors.map((item) => autoFilterColorButtonHtml("fill", item, "color-filter")).join("")}`
    : "";
  const fontItems = fontColors.length
    ? `<div class="auto-filter-submenu-heading">フォントの色で並べ替え</div>${fontColors.map((item) => autoFilterColorButtonHtml("font", item, "color-sort")).join("")}
       <div class="auto-filter-submenu-heading">フォントの色でフィルター</div>${fontColors.map((item) => autoFilterColorButtonHtml("font", item, "color-filter")).join("")}`
    : "";
  return `
    <div class="auto-filter-condition">
      <button type="button" class="auto-filter-submenu-toggle${active ? " is-filtered" : ""}" data-auto-filter-action="color-filter-menu">
        <span>色フィルター</span><span aria-hidden="true">›</span>
      </button>
      <div class="auto-filter-submenu hidden" role="menu" aria-label="色フィルター">
        ${fillItems}${fontItems}
      </div>
    </div>`;
}



function autoFilterColorButtonHtml(target, item, action = "color-filter") {
  const label = target === "font" ? "フォントの色" : "セルの色";
  const actionLabel = action === "color-sort" ? "を上に表示" : "";
  return `<button type="button" role="menuitem" data-auto-filter-action="${escapeAttr(action)}" data-auto-filter-color-target="${target}" data-auto-filter-color="${escapeAttr(item.color)}"><span class="auto-filter-color-swatch" style="--auto-filter-color:${escapeAttr(item.color)}"></span><span>${escapeHtml(`${label}${actionLabel}`)}</span></button>`;
}



function autoFilterUniqueValues(sheet, col) {
  const range = sheet.autoFilter?.range;
  if (!range) return [];
  const values = new Map();
  for (let row = range.top + 1; row <= range.bottom; row += 1) {
    const value = autoFilterCellValue(sheet, row, col);
    const key = autoFilterValueGroupKey(value);
    const current = values.get(key);
    if (current) current.count += 1;
    else values.set(key, { value, count: 1 });
  }
  return Array.from(values.values()).sort(compareAutoFilterValueItems);
}



function compareAutoFilterValueItems(a, b) {
  const left = String(a?.value ?? "");
  const right = String(b?.value ?? "");
  const leftDate = autoFilterDateSerial(left);
  const rightDate = autoFilterDateSerial(right);
  if (leftDate != null && rightDate != null && leftDate !== rightDate) return leftDate - rightDate;
  const leftNumber = Number(left.replace(/,/g, ""));
  const rightNumber = Number(right.replace(/,/g, ""));
  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber) && leftNumber !== rightNumber) return leftNumber - rightNumber;
  return left.localeCompare(right, "ja", { numeric: true, sensitivity: "base" });
}



function autoFilterCellValue(sheet, row, col) {
  return String(getDisplayForCell(sheet, getSheetId(sheet.name), row, col) ?? "");
}



function normalizeAutoFilterCriterion(criterion) {
  if (!criterion) return { kind: "all" };
  if (Array.isArray(criterion)) return { kind: "values", values: criterion };
  if (criterion.kind === "values") return { kind: "values", values: criterion.values || [] };
  if (criterion.kind === "dateGroups") {
    return {
      kind: "dateGroups",
      values: criterion.values || [],
      includeBlank: Boolean(criterion.includeBlank),
      dateGroups: Array.isArray(criterion.dateGroups) ? criterion.dateGroups : [],
    };
  }
  if (criterion.kind === "custom") return criterion;
  if (criterion.kind === "color") return criterion;
  return { kind: "all" };
}



function autoFilterValueMatchesCriterion(value, criterion) {
  const normalized = normalizeAutoFilterCriterion(criterion);
  if (normalized.kind === "all") return true;
  if (normalized.kind === "values") return autoFilterValueInSelection(value, normalized.values || []);
  if (normalized.kind === "dateGroups") return autoFilterValueInDateGroups(value, normalized);
  if (normalized.kind !== "custom") return true;
  if (normalized.dateGroup) return autoFilterValueMatchesDateGroup(value, normalized.dateGroup);
  if (normalized.datePeriod) return autoFilterValueMatchesDatePeriod(value, normalized.datePeriod);
  if (normalized.top10) return autoFilterTop10ValueMatches(value, normalized);
  const first = autoFilterConditionMatches(value, normalized.operator1, normalized.value1);
  const hasSecond = String(normalized.value2 ?? "").trim() !== "" || autoFilterOperatorUsesBlankValue(normalized.operator2);
  if (!hasSecond) return first;
  const second = autoFilterConditionMatches(value, normalized.operator2, normalized.value2);
  return normalized.logic === "or" ? first || second : first && second;
}



function autoFilterValueGroupKey(value) {
  const date = autoFilterDateSerial(value);
  if (date != null) return `date:${date}`;
  return `text:${String(value ?? "")}`;
}



function autoFilterValueInSelection(value, selectedValues) {
  const valueText = String(value ?? "");
  const valueDate = autoFilterDateSerial(valueText);
  return (selectedValues || []).some((selected) => {
    const selectedText = String(selected ?? "");
    if (selectedText === valueText) return true;
    const selectedDate = autoFilterDateSerial(selectedText);
    return valueDate != null && selectedDate != null && valueDate === selectedDate;
  });
}



function autoFilterValueInDateGroups(value, criterion = {}) {
  const valueText = String(value ?? "");
  if (valueText === "" && criterion.includeBlank) return true;
  if (autoFilterValueInSelection(valueText, criterion.values || [])) return true;
  return (criterion.dateGroups || []).some((group) => autoFilterValueMatchesDateGroup(valueText, group));
}



function autoFilterValueMatchesDateGroup(value, group = {}) {
  const parts = autoFilterDatePartsFromValue(value);
  if (!parts || parts.year !== Number(group.year)) return false;
  const grouping = group.grouping || "year";
  if (grouping === "year") return true;
  if (parts.month !== Number(group.month)) return false;
  if (grouping === "month") return true;
  if (parts.day !== Number(group.day)) return false;
  if (grouping === "day") return true;
  if (parts.hour !== Number(group.hour)) return false;
  if (grouping === "hour") return true;
  if (parts.minute !== Number(group.minute)) return false;
  if (grouping === "minute") return true;
  if (parts.second !== Number(group.second)) return false;
  return true;
}



function autoFilterValueMatchesDatePeriod(value, period) {
  const serial = autoFilterDateSerial(value);
  if (serial == null) return false;
  const parts = autoFilterDatePartsFromSerial(serial);
  if (!parts) return false;
  const type = String(period || "");
  const month = type.match(/^M([1-9]|1[0-2])$/);
  if (month) return parts.month === Number(month[1]);
  const quarter = type.match(/^Q([1-4])$/);
  if (quarter) return Math.floor((parts.month - 1) / 3) + 1 === Number(quarter[1]);
  return false;
}



function autoFilterTop10ValueMatches(value, criterion) {
  const actualNumber = Number(String(value ?? "").replace(/,/g, ""));
  const threshold = Number(String(criterion.value1 ?? "").replace(/,/g, ""));
  if (!Number.isFinite(actualNumber) || !Number.isFinite(threshold)) return false;
  return criterion.top10?.direction === "bottom"
    ? actualNumber <= threshold
    : actualNumber >= threshold;
}



function autoFilterConditionMatches(value, operator, expected) {
  const actualText = String(value ?? "");
  const expectedText = String(expected ?? "");
  const actualNumber = Number(actualText.replace(/,/g, ""));
  const expectedNumber = Number(expectedText.replace(/,/g, ""));
  const canCompareNumber = Number.isFinite(actualNumber) && Number.isFinite(expectedNumber);
  const actualDate = autoFilterDateSerial(actualText);
  const expectedDate = autoFilterDateSerial(expectedText);
  const canCompareDate = actualDate != null && expectedDate != null && ["equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual"].includes(operator);
  const actualFolded = actualText.toLocaleLowerCase("ja-JP");
  const expectedFolded = expectedText.toLocaleLowerCase("ja-JP");
  const wildcardMatch = () => autoFilterWildcardMatches(actualText, expectedText);
  switch (operator) {
    case "blank":
      return actualText === "";
    case "nonBlank":
      return actualText !== "";
    case "notEqual":
      return canCompareNumber ? actualNumber !== expectedNumber : canCompareDate ? actualDate !== expectedDate : !wildcardMatch();
    case "contains":
      return actualFolded.includes(expectedFolded);
    case "notContains":
      return !actualFolded.includes(expectedFolded);
    case "beginsWith":
      return actualFolded.startsWith(expectedFolded);
    case "notBeginsWith":
      return !actualFolded.startsWith(expectedFolded);
    case "endsWith":
      return actualFolded.endsWith(expectedFolded);
    case "notEndsWith":
      return !actualFolded.endsWith(expectedFolded);
    case "greaterThan":
      if (Number.isFinite(expectedNumber) && !Number.isFinite(actualNumber)) return false;
      return canCompareNumber ? actualNumber > expectedNumber : canCompareDate ? actualDate > expectedDate : actualText.localeCompare(expectedText, "ja", { numeric: true, sensitivity: "base" }) > 0;
    case "greaterThanOrEqual":
      if (Number.isFinite(expectedNumber) && !Number.isFinite(actualNumber)) return false;
      return canCompareNumber ? actualNumber >= expectedNumber : canCompareDate ? actualDate >= expectedDate : actualText.localeCompare(expectedText, "ja", { numeric: true, sensitivity: "base" }) >= 0;
    case "lessThan":
      if (Number.isFinite(expectedNumber) && !Number.isFinite(actualNumber)) return false;
      return canCompareNumber ? actualNumber < expectedNumber : canCompareDate ? actualDate < expectedDate : actualText.localeCompare(expectedText, "ja", { numeric: true, sensitivity: "base" }) < 0;
    case "lessThanOrEqual":
      if (Number.isFinite(expectedNumber) && !Number.isFinite(actualNumber)) return false;
      return canCompareNumber ? actualNumber <= expectedNumber : canCompareDate ? actualDate <= expectedDate : actualText.localeCompare(expectedText, "ja", { numeric: true, sensitivity: "base" }) <= 0;
    case "equal":
    default:
      return canCompareNumber ? actualNumber === expectedNumber : canCompareDate ? actualDate === expectedDate : wildcardMatch();
  }
}



function autoFilterDateSerial(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return autoFilterSerialFromDateParts(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }
  const text = String(value ?? "").trim();
  if (!text) return null;
  const namedDate = autoFilterMonthNameDateSerial(text);
  if (namedDate != null) return namedDate;
  const numbers = text.match(/\d+/g);
  if (!numbers || numbers.length < 3) return null;
  const parts = numbers.slice(0, 3).map((part) => Number.parseInt(part, 10));
  const [first, second, third] = parts;
  const candidates = [];
  if (first > 31) {
    candidates.push([first, second, third]);
  } else if (third > 31) {
    if (first <= 12) candidates.push([third, first, second]);
    if (second <= 12) candidates.push([third, second, first]);
  } else {
    if (first >= 20 && first <= 99 && second <= 12) candidates.push([normalizeTwoDigitYear(first), second, third]);
    if (first <= 12) candidates.push([normalizeTwoDigitYear(third), first, second]);
    if (second <= 12) candidates.push([normalizeTwoDigitYear(third), second, first]);
  }
  for (const [year, month, day] of candidates) {
    const serial = autoFilterSerialFromDateParts(year, month, day);
    if (serial != null) return serial;
  }
  return null;
}



function autoFilterMonthNameDateSerial(text) {
  const monthMatch = String(text).match(/\b(january|jan|february|feb|march|mar|april|apr|may|june|jun|july|jul|august|aug|september|sep|sept|october|oct|november|nov|december|dec)\b/i);
  if (!monthMatch) return null;
  const month = AUTO_FILTER_MONTH_NAME_INDEX[monthMatch[1].toLocaleLowerCase("en-US")];
  const numbers = String(text).match(/\d+/g)?.map((part) => Number.parseInt(part, 10)) || [];
  if (!month || numbers.length < 2) return null;
  const candidates = [];
  if (numbers[0] > 31) candidates.push([numbers[0], month, numbers[1]]);
  else candidates.push([numbers[1], month, numbers[0]]);
  if (numbers.length >= 3) candidates.push([numbers[2], month, numbers[1]]);
  for (const [year, candidateMonth, day] of candidates) {
    const serial = autoFilterSerialFromDateParts(normalizeTwoDigitYear(year), candidateMonth, day);
    if (serial != null) return serial;
  }
  return null;
}



function normalizeTwoDigitYear(year) {
  if (!Number.isFinite(year)) return year;
  return year < 100 ? year + (year < 30 ? 2000 : 1900) : year;
}



function autoFilterSerialFromDateParts(year, month, day) {
  year = normalizeTwoDigitYear(Number(year));
  month = Number(month);
  day = Number(day);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return null;
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  return Math.floor(date.getTime() / 86400000);
}



function autoFilterDatePartsFromSerial(serial) {
  if (serial == null || serial === "") return null;
  const numeric = Number(serial);
  if (!Number.isFinite(numeric)) return null;
  const date = new Date(Math.floor(numeric) * 86400000);
  if (Number.isNaN(date.getTime())) return null;
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
  };
}



function autoFilterDatePartsFromValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return {
      year: value.getFullYear(),
      month: value.getMonth() + 1,
      day: value.getDate(),
      hour: value.getHours(),
      minute: value.getMinutes(),
      second: value.getSeconds(),
    };
  }
  const serial = autoFilterDateSerial(value);
  const parts = autoFilterDatePartsFromSerial(serial);
  if (!parts) return null;
  return { ...parts, ...autoFilterTimePartsFromText(value) };
}



function autoFilterTimePartsFromText(value) {
  const text = String(value ?? "");
  const match = text.match(/(?:^|[^\d])(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?\b/i);
  if (!match) return { hour: 0, minute: 0, second: 0 };
  let hour = Number.parseInt(match[1], 10);
  const minute = Number.parseInt(match[2], 10);
  const second = match[3] == null ? 0 : Number.parseInt(match[3], 10);
  const meridiem = String(match[4] || "").toUpperCase();
  if (meridiem === "PM" && hour < 12) hour += 12;
  if (meridiem === "AM" && hour === 12) hour = 0;
  if (!(hour >= 0 && hour <= 23) || !(minute >= 0 && minute <= 59) || !(second >= 0 && second <= 59)) {
    return { hour: 0, minute: 0, second: 0 };
  }
  return { hour, minute, second };
}



function autoFilterDateForSerial(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}/${month}/${day}`;
}



function autoFilterLocalDate(year, month, day) {
  return new Date(year, month, day, 0, 0, 0, 0);
}



function autoFilterAddDays(date, days) {
  return autoFilterLocalDate(date.getFullYear(), date.getMonth(), date.getDate() + days);
}



function autoFilterDateRangeForPreset(preset, referenceDate = new Date()) {
  const today = autoFilterLocalDate(referenceDate.getFullYear(), referenceDate.getMonth(), referenceDate.getDate());
  const year = today.getFullYear();
  const month = today.getMonth();
  const day = today.getDate();
  const weekStart = autoFilterAddDays(today, -today.getDay());
  const quarterStartMonth = Math.floor(month / 3) * 3;
  const ranges = {
    dateToday: [today, today],
    dateYesterday: [autoFilterAddDays(today, -1), autoFilterAddDays(today, -1)],
    dateTomorrow: [autoFilterAddDays(today, 1), autoFilterAddDays(today, 1)],
    dateThisWeek: [weekStart, autoFilterAddDays(weekStart, 6)],
    dateLastWeek: [autoFilterAddDays(weekStart, -7), autoFilterAddDays(weekStart, -1)],
    dateNextWeek: [autoFilterAddDays(weekStart, 7), autoFilterAddDays(weekStart, 13)],
    dateThisMonth: [autoFilterLocalDate(year, month, 1), autoFilterLocalDate(year, month + 1, 0)],
    dateLastMonth: [autoFilterLocalDate(year, month - 1, 1), autoFilterLocalDate(year, month, 0)],
    dateNextMonth: [autoFilterLocalDate(year, month + 1, 1), autoFilterLocalDate(year, month + 2, 0)],
    dateThisQuarter: [autoFilterLocalDate(year, quarterStartMonth, 1), autoFilterLocalDate(year, quarterStartMonth + 3, 0)],
    dateLastQuarter: [autoFilterLocalDate(year, quarterStartMonth - 3, 1), autoFilterLocalDate(year, quarterStartMonth, 0)],
    dateNextQuarter: [autoFilterLocalDate(year, quarterStartMonth + 3, 1), autoFilterLocalDate(year, quarterStartMonth + 6, 0)],
    dateThisYear: [autoFilterLocalDate(year, 0, 1), autoFilterLocalDate(year, 11, 31)],
    dateLastYear: [autoFilterLocalDate(year - 1, 0, 1), autoFilterLocalDate(year - 1, 11, 31)],
    dateNextYear: [autoFilterLocalDate(year + 1, 0, 1), autoFilterLocalDate(year + 1, 11, 31)],
    dateYearToDate: [autoFilterLocalDate(year, 0, 1), autoFilterLocalDate(year, month, day)],
  };
  const range = ranges[preset];
  return range ? { start: range[0], end: range[1] } : null;
}



function autoFilterWildcardMatches(actualText, expectedText) {
  const expected = String(expectedText ?? "");
  if (/[*?~]/.test(expected)) return advancedWildcardToRegExp(expected).test(String(actualText ?? ""));
  return String(actualText ?? "").localeCompare(expected, "ja", { sensitivity: "base" }) === 0;
}



function handleAutoFilterSelectAllChange(event) {
  const checked = event.currentTarget.checked;
  const menu = event.currentTarget.closest(".auto-filter-menu");
  if (menu) menu.dataset.valueSelectionDirty = "true";
  menu?.querySelectorAll(".auto-filter-value:not(.is-hidden) [data-auto-filter-value]").forEach((input) => {
    input.checked = checked;
  });
  updateAutoFilterSelectAllState(menu);
  syncAutoFilterMenuState(menu);
}



function handleAutoFilterValueChange(event) {
  const menu = event.currentTarget.closest(".auto-filter-menu");
  if (menu) menu.dataset.valueSelectionDirty = "true";
  updateAutoFilterSelectAllState(menu);
  syncAutoFilterMenuState(menu);
}



function handleAutoFilterSearchInput(event) {
  const menu = event.currentTarget.closest(".auto-filter-menu");
  if (!menu) return;
  const query = String(event.currentTarget.value || "").trim().toLocaleLowerCase("ja-JP");
  let visibleCount = 0;
  menu.querySelectorAll(".auto-filter-values .auto-filter-value").forEach((label) => {
    const visible = !query || String(label.dataset.filterLabel || "").includes(query);
    label.classList.toggle("is-hidden", !visible);
    if (visible) visibleCount += 1;
  });
  menu.querySelector(".auto-filter-add-current")?.classList.toggle("is-hidden", !query || menu.dataset.hasValueFilter !== "true");
  menu.querySelector(".auto-filter-empty")?.classList.toggle("is-hidden", visibleCount > 0);
  updateAutoFilterSelectAllState(menu);
  syncAutoFilterMenuState(menu);
}



function autoFilterSelectedValuesFromMenu(menu, allValues, currentCriterion) {
  const query = String(menu?.querySelector("[data-auto-filter-search]")?.value || "").trim();
  const visibleSelection = autoFilterVisibleValueInputs(menu)
    .filter((input) => input.checked)
    .map((input) => input.dataset.autoFilterValue || "");
  if (query) {
    const addCurrent = Boolean(menu?.querySelector("[data-auto-filter-add-current]")?.checked);
    if (addCurrent && currentCriterion?.kind === "values") {
      return [...new Set([...(currentCriterion.values || []), ...visibleSelection])].filter((value) => allValues.includes(value));
    }
    return [...new Set(Array.from(menu.querySelectorAll("[data-auto-filter-value]:checked")).map((input) => input.dataset.autoFilterValue || ""))]
      .filter((value) => allValues.includes(value));
  }
  const selectAll = menu?.querySelector("[data-auto-filter-select-all]");
  if (selectAll?.checked && !selectAll.indeterminate && (currentCriterion?.kind === "all" || menu?.dataset.hasMoreValues !== "true")) {
    return allValues;
  }
  return [...new Set(Array.from(menu.querySelectorAll("[data-auto-filter-value]:checked")).map((input) => input.dataset.autoFilterValue || ""))]
    .filter((value) => allValues.includes(value));
}



function autoFilterVisibleValueInputs(menu) {
  if (!menu) return [];
  return Array.from(menu.querySelectorAll(".auto-filter-values .auto-filter-value:not(.is-hidden) [data-auto-filter-value]"));
}



function updateAutoFilterSelectAllState(menu) {
  const selectAll = menu?.querySelector("[data-auto-filter-select-all]");
  if (!selectAll) return;
  const inputs = autoFilterVisibleValueInputs(menu);
  const checkedCount = inputs.filter((input) => input.checked).length;
  selectAll.disabled = inputs.length === 0;
  selectAll.checked = inputs.length > 0 && checkedCount === inputs.length;
  selectAll.indeterminate = checkedCount > 0 && checkedCount < inputs.length;
}



function syncAutoFilterMenuState(menu) {
  const summary = menu?.querySelector("[data-auto-filter-menu-summary]");
  const sheet = activeSheet();
  const col = Number(menu.dataset.filterCol || state.autoFilterMenu?.col);
  const range = sheet?.autoFilter?.range;
  if (!sheet || !range || !col) {
    if (summary) summary.textContent = "";
    return;
  }
  const query = String(menu.querySelector("[data-auto-filter-search]")?.value || "").trim();
  const visibleInputs = autoFilterVisibleValueInputs(menu);
  const checkedVisible = visibleInputs.filter((input) => input.checked).length;
  const renderedTotal = Number(menu.dataset.renderedValues || visibleInputs.length || 0);
  const totalValues = Number(menu.dataset.totalValues || renderedTotal || 0);
  const allValues = autoFilterUniqueValues(sheet, col).map((item) => item.value);
  const currentCriterion = normalizeAutoFilterCriterion(sheet.autoFilter.criteria?.[col] || null);
  const selectedValues = autoFilterSelectedValuesFromMenu(menu, allValues, currentCriterion);
  const isEditingValueList = menu.dataset.valueSelectionDirty === "true" || query;
  const criterion = !isEditingValueList && currentCriterion.kind !== "all"
    ? currentCriterion
    : selectedValues.length === allValues.length
      ? { kind: "all" }
      : { kind: "values", values: selectedValues };
  const totalRows = Math.max(0, range.bottom - range.top);
  const matchCount = criterion.kind === "values" && !selectedValues.length ? 0 : countRowsMatchingAutoFilterCriterion(sheet, range, col, criterion);
  const prefix = query
    ? `検索結果 ${visibleInputs.length.toLocaleString("ja-JP")} 件中 ${checkedVisible.toLocaleString("ja-JP")} 件を選択`
    : totalValues > renderedTotal
      ? `表示中 ${renderedTotal.toLocaleString("ja-JP")} / ${totalValues.toLocaleString("ja-JP")} 件`
      : `${checkedVisible.toLocaleString("ja-JP")} / ${totalValues.toLocaleString("ja-JP")} 件を選択`;
  if (summary) {
    summary.textContent = `${prefix}。表示予定: ${matchCount.toLocaleString("ja-JP")} / ${totalRows.toLocaleString("ja-JP")} 行。`;
    summary.classList.toggle("is-warning", criterion.kind === "values" && selectedValues.length === 0);
  }
  const ok = menu.querySelector("[data-auto-filter-action='ok']");
  if (ok) ok.disabled = false;
}



function autoFilterFocusableControls(menu) {
  if (!menu) return [];
  return Array.from(menu.querySelectorAll("button:not([disabled]), input:not([disabled])")).filter((control) => {
    if (control.closest(".hidden") || control.closest(".is-hidden")) return false;
    return control.offsetParent !== null;
  });
}



function handleAutoFilterMenuKeydown(event) {
  if (!state.autoFilterMenu) return;
  if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    closeAutoFilterMenu();
    return;
  }
  if (event.key !== "ArrowDown" && event.key !== "ArrowUp") return;
  const controls = autoFilterFocusableControls(event.currentTarget);
  if (!controls.length) return;
  event.preventDefault();
  const activeIndex = controls.indexOf(document.activeElement);
  const delta = event.key === "ArrowDown" ? 1 : -1;
  const nextIndex = activeIndex < 0 ? (delta > 0 ? 0 : controls.length - 1) : (activeIndex + delta + controls.length) % controls.length;
  controls[nextIndex]?.focus({ preventScroll: true });
}



function handleAutoFilterMenuAction(event) {
  const action = event.currentTarget.dataset.autoFilterAction;
  const menu = event.currentTarget.closest(".auto-filter-menu");
  const sheet = activeSheet();
  const col = Number(menu?.dataset.filterCol || state.autoFilterMenu?.col);
  if (!action || !sheet?.autoFilter?.range || !col) return;
  if (action === "condition-menu" || action === "color-filter-menu") {
    const container = event.currentTarget.closest(".auto-filter-condition");
    const submenu = container?.querySelector(".auto-filter-submenu");
    submenu?.classList.toggle("hidden");
    event.currentTarget.setAttribute("aria-expanded", String(!submenu?.classList.contains("hidden")));
    return;
  }
  if (action === "cancel") {
    closeAutoFilterMenu();
    return;
  }
  if (action === "sort-asc" || action === "sort-desc") {
    const range = sheet.autoFilter.range;
    closeAutoFilterMenu();
    sortRangeRows(range, col, action === "sort-asc" ? "asc" : "desc", true);
    return;
  }
  if (action === "custom") {
    closeAutoFilterMenu();
    openCustomAutoFilterDialog(col);
    return;
  }
  if (action === "condition") {
    const operator = event.currentTarget.dataset.autoFilterOperator || "equal";
    closeAutoFilterMenu();
    if (operator === "aboveAverage" || operator === "belowAverage") {
      applyAverageAutoFilter(col, operator);
    } else if (operator === "aboveStdDev" || operator === "belowStdDev") {
      applyStdDevAutoFilter(col, operator);
    } else if (operator === "top10") {
      openTop10AutoFilterDialog(col);
    } else if (operator === "blanks" || operator === "nonBlanks") {
      applyBlankAutoFilter(col, operator);
    } else if (operator.startsWith("datePeriod")) {
      applyDatePeriodAutoFilter(col, operator);
    } else if (operator.startsWith("date")) {
      applyDatePresetAutoFilter(col, operator);
    } else {
      openCustomAutoFilterDialog(col, autoFilterPresetForOperator(operator));
    }
    return;
  }
  if (action === "color-filter") {
    const target = event.currentTarget.dataset.autoFilterColorTarget || "fill";
    const color = event.currentTarget.dataset.autoFilterColor || "";
    closeAutoFilterMenu();
    applyColorAutoFilter(col, target, color);
    return;
  }
  if (action === "color-sort") {
    const target = event.currentTarget.dataset.autoFilterColorTarget || "fill";
    const color = event.currentTarget.dataset.autoFilterColor || "";
    closeAutoFilterMenu();
    applyColorAutoFilterSort(col, target, color);
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  if (action === "clear") {
    if (!sheet.autoFilter.criteria?.[col]) {
      closeAutoFilterMenu();
      return;
    }
    delete sheet.autoFilter.criteria[col];
    reapplyAutoFilter({ silent: true });
    closeAutoFilterMenu();
    commitDataOperation(sheet, sheetBefore, sheet.autoFilter.range, "フィルター クリア", autoFilterStatusMessage(`${columnName(col)}列のフィルターをクリアしました`, sheet, sheet.autoFilter.range));
    return;
  }
  if (action === "ok") {
    const allValues = autoFilterUniqueValues(sheet, col).map((item) => item.value);
    const currentCriterion = normalizeAutoFilterCriterion(sheet.autoFilter.criteria?.[col] || null);
    const hasQuery = String(menu?.querySelector("[data-auto-filter-search]")?.value || "").trim() !== "";
    const valueSelectionDirty = menu?.dataset.valueSelectionDirty === "true" || hasQuery;
    if (!valueSelectionDirty) {
      closeAutoFilterMenu();
      return;
    }
    const checkedValues = autoFilterSelectedValuesFromMenu(menu, allValues, currentCriterion);
    if (checkedValues.length === allValues.length) delete sheet.autoFilter.criteria[col];
    else sheet.autoFilter.criteria[col] = checkedValues;
    reapplyAutoFilter({ silent: true });
    closeAutoFilterMenu();
    commitDataOperation(sheet, sheetBefore, sheet.autoFilter.range, "フィルター", autoFilterStatusMessage(`${columnName(col)}列にフィルターを適用しました`, sheet, sheet.autoFilter.range));
  }
}



function autoFilterPresetForOperator(operator) {
  if (operator === "between") {
    return { operator1: "greaterThanOrEqual", operator2: "lessThanOrEqual", logic: "and" };
  }
  if (operator === "notBetween") {
    return { operator1: "lessThan", operator2: "greaterThan", logic: "or" };
  }
  if (operator === "custom") return {};
  return { operator1: operator || "equal", operator2: "equal", logic: "and" };
}



function applyDatePresetAutoFilter(col, preset) {
  const sheet = activeSheet();
  const range = sheet?.autoFilter?.range;
  const dateRange = autoFilterDateRangeForPreset(preset);
  if (!sheet || !range || !col || !dateRange) return;
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.autoFilter.criteria[col] = {
    kind: "custom",
    operator1: "greaterThanOrEqual",
    value1: autoFilterDateForSerial(dateRange.start),
    logic: "and",
    operator2: "lessThanOrEqual",
    value2: autoFilterDateForSerial(dateRange.end),
    datePreset: preset,
  };
  reapplyAutoFilter({ silent: true });
  commitDataOperation(sheet, sheetBefore, range, "日付フィルター", autoFilterStatusMessage(`${columnName(col)}列に日付フィルターを適用しました`, sheet, range));
}



function autoFilterDynamicDatePeriodType(period) {
  const value = String(period || "");
  const direct = value.match(/^(M(?:[1-9]|1[0-2])|Q[1-4])$/);
  if (direct) return direct[1];
  const month = value.match(/^datePeriodM([1-9]|1[0-2])$/);
  if (month) return `M${Number(month[1])}`;
  const quarter = value.match(/^datePeriodQ([1-4])$/);
  if (quarter) return `Q${Number(quarter[1])}`;
  return "";
}



function applyDatePeriodAutoFilter(col, period) {
  const sheet = activeSheet();
  const range = sheet?.autoFilter?.range;
  const type = autoFilterDynamicDatePeriodType(period);
  if (!sheet || !range || !col || !type) return;
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.autoFilter.criteria[col] = {
    kind: "custom",
    datePeriod: type,
    dynamicSource: { type },
  };
  reapplyAutoFilter({ silent: true });
  commitDataOperation(sheet, sheetBefore, range, "日付フィルター", autoFilterStatusMessage(`${columnName(col)}列に${autoFilterDatePeriodLabel(type)}のフィルターを適用しました`, sheet, range));
}



function applyAverageAutoFilter(col, operator) {
  const sheet = activeSheet();
  const range = sheet?.autoFilter?.range;
  if (!sheet || !range || !col) return;
  const values = [];
  for (let row = range.top + 1; row <= range.bottom; row += 1) {
    const number = Number(autoFilterCellValue(sheet, row, col).replace(/,/g, ""));
    if (Number.isFinite(number)) values.push(number);
  }
  if (!values.length) {
    showMessageBox("平均を計算できる数値データがありません。", { title: "Cht WebSheet" });
    return;
  }
  const average = values.reduce((total, value) => total + value, 0) / values.length;
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.autoFilter.criteria[col] = {
    kind: "custom",
    operator1: operator === "aboveAverage" ? "greaterThan" : "lessThan",
    value1: String(average),
    logic: "and",
    operator2: "equal",
    value2: "",
    averageOperator: operator,
  };
  reapplyAutoFilter({ silent: true });
  commitDataOperation(
    sheet,
    sheetBefore,
    range,
    operator === "aboveAverage" ? "平均より上フィルター" : "平均より下フィルター",
    autoFilterStatusMessage(`${columnName(col)}列に${operator === "aboveAverage" ? "平均より上" : "平均より下"}のフィルターを適用しました`, sheet, range),
  );
}



function applyStdDevAutoFilter(col, operator, count = 1) {
  const sheet = activeSheet();
  const range = sheet?.autoFilter?.range;
  if (!sheet || !range || !col) return;
  const values = [];
  for (let row = range.top + 1; row <= range.bottom; row += 1) {
    const number = Number(autoFilterCellValue(sheet, row, col).replace(/,/g, ""));
    if (Number.isFinite(number)) values.push(number);
  }
  if (!values.length) {
    showMessageBox("標準偏差を計算できる数値データがありません。", { title: "Cht WebSheet" });
    return;
  }
  const stdDevCount = autoFilterStdDevCount(count);
  const average = values.reduce((total, value) => total + value, 0) / values.length;
  const variance = values.length > 1
    ? values.reduce((total, value) => total + ((value - average) ** 2), 0) / (values.length - 1)
    : 0;
  const threshold = average + (operator === "aboveStdDev" ? 1 : -1) * Math.sqrt(variance) * stdDevCount;
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.autoFilter.criteria[col] = {
    kind: "custom",
    operator1: operator === "aboveStdDev" ? "greaterThan" : "lessThan",
    value1: String(threshold),
    logic: "and",
    operator2: "equal",
    value2: "",
    dynamicSource: { type: operator, val: stdDevCount },
    stdDevOperator: operator,
    stdDevCount,
  };
  reapplyAutoFilter({ silent: true });
  commitDataOperation(
    sheet,
    sheetBefore,
    range,
    operator === "aboveStdDev" ? "標準偏差より上フィルター" : "標準偏差より下フィルター",
    autoFilterStatusMessage(`${columnName(col)}列に${operator === "aboveStdDev" ? "標準偏差より上" : "標準偏差より下"}のフィルターを適用しました`, sheet, range),
  );
}



function applyBlankAutoFilter(col, operator) {
  const sheet = activeSheet();
  const range = sheet?.autoFilter?.range;
  if (!sheet || !range || !col) return;
  const blank = operator === "blanks";
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.autoFilter.criteria[col] = {
    kind: "custom",
    operator1: blank ? "blank" : "nonBlank",
    value1: "",
    logic: "and",
    operator2: "equal",
    value2: "",
    blankOperator: operator,
  };
  reapplyAutoFilter({ silent: true });
  commitDataOperation(
    sheet,
    sheetBefore,
    range,
    blank ? "空白セル フィルター" : "空白以外のセル フィルター",
    autoFilterStatusMessage(`${columnName(col)}列に${blank ? "空白セル" : "空白以外のセル"}のフィルターを適用しました`, sheet, range),
  );
}



function applyColorAutoFilter(col, target, color) {
  const sheet = activeSheet();
  const range = sheet?.autoFilter?.range;
  const normalizedColor = normalizeSortColor(color);
  if (!sheet || !range || !col || !normalizedColor) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const filterTarget = target === "font" ? "font" : "fill";
  sheet.autoFilter.criteria[col] = {
    kind: "color",
    target: filterTarget,
    color: normalizedColor,
  };
  reapplyAutoFilter({ silent: true });
  commitDataOperation(
    sheet,
    sheetBefore,
    range,
    filterTarget === "font" ? "フォントの色フィルター" : "セルの色フィルター",
    autoFilterStatusMessage(`${columnName(col)}列に${filterTarget === "font" ? "フォントの色" : "セルの色"}フィルターを適用しました`, sheet, range),
  );
}



function applyColorAutoFilterSort(col, target, color) {
  const sheet = activeSheet();
  const range = sheet?.autoFilter?.range;
  const normalizedColor = normalizeSortColor(color);
  if (!sheet || !range || !col || !normalizedColor) return;
  sortRangeRowsByKeys(
    range,
    [{
      col,
      direction: "asc",
      sortOn: target === "font" ? "fontColor" : "cellColor",
      color: normalizedColor,
      colorPosition: "top",
    }],
    true,
  );
  setStatus(`${columnName(col)}列を${target === "font" ? "フォントの色" : "セルの色"}で並べ替えました。`);
}



function openTop10AutoFilterDialog(col) {
  const sheet = activeSheet();
  if (!sheet?.autoFilter?.range || !col) return;
  state.dataDialog = { type: "auto-filter-top10", range: { ...sheet.autoFilter.range }, col };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog top10-filter-dialog" role="dialog" aria-modal="true" aria-labelledby="top10FilterDialogTitle" data-data-dialog="auto-filter-top10">
      <div class="excel-mini-titlebar" id="top10FilterDialogTitle">トップ テン オートフィルター</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-note">${escapeHtml(columnName(col))} 列で表示する上位または下位の値を指定します。</div>
        <div class="top10-filter-row">
          <span>表示</span>
          <select data-data-field="top10Direction" data-auto-filter-sync aria-label="上位または下位">
            <option value="top">上位</option>
            <option value="bottom">下位</option>
          </select>
          <input type="number" min="1" max="500" step="1" value="10" data-data-field="top10Count" data-auto-filter-sync aria-label="数" />
          <select data-data-field="top10Mode" data-auto-filter-sync aria-label="単位">
            <option value="items">項目</option>
            <option value="percent">パーセント</option>
          </select>
        </div>
        <div class="data-dialog-summary auto-filter-summary" data-auto-filter-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncAutoFilterDialog(overlay.querySelector("[data-data-dialog='auto-filter-top10']"));
  overlay.querySelector("[data-data-field='top10Count']")?.focus({ preventScroll: true });
  overlay.querySelector("[data-data-field='top10Count']")?.select?.();
}



function applyTop10AutoFilterDialog(dialog) {
  const col = state.dataDialog?.col;
  const direction = dialog.querySelector("[data-data-field='top10Direction']")?.value === "bottom" ? "bottom" : "top";
  const mode = dialog.querySelector("[data-data-field='top10Mode']")?.value === "percent" ? "percent" : "items";
  const rawCount = Number(dialog.querySelector("[data-data-field='top10Count']")?.value);
  const count = Math.max(1, Math.min(500, Math.floor(Number.isFinite(rawCount) ? rawCount : 10)));
  closeDataDialog();
  applyTop10AutoFilter(col, { direction, mode, count });
}



function handleAutoFilterDialogSync(event) {
  syncAutoFilterDialog(event.currentTarget.closest("[data-data-dialog]"));
}



function syncAutoFilterDialog(dialog) {
  if (!dialog) return;
  if (dialog.dataset.dataDialog === "auto-filter-custom") syncCustomAutoFilterDialogControls(dialog);
  const summary = dialog.querySelector("[data-auto-filter-summary]");
  const preview = autoFilterDialogPreview(dialog);
  if (summary) {
    summary.textContent = preview.message;
    summary.classList.toggle("is-warning", preview.warning || preview.blocked);
  }
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = Boolean(preview.blocked);
}



function syncCustomAutoFilterDialogControls(dialog) {
  const pairs = [
    ["operator1", "value1"],
    ["operator2", "value2"],
  ];
  pairs.forEach(([operatorField, valueField]) => {
    const operator = dialog.querySelector(`[data-data-field='${operatorField}']`)?.value || "";
    const input = dialog.querySelector(`[data-data-field='${valueField}']`);
    if (input) input.disabled = autoFilterOperatorUsesBlankValue(operator);
  });
}



function autoFilterDialogPreview(dialog) {
  const type = dialog.dataset.dataDialog;
  const sheet = activeSheet();
  const range = sheet?.autoFilter?.range;
  const col = state.dataDialog?.col;
  if (!sheet || !range || !col) return { warning: true, message: "フィルター範囲がありません。" };
  if (type === "auto-filter-top10") {
    const options = top10OptionsFromDialog(dialog);
    const result = top10AutoFilterCriterionForColumn(sheet, range, col, options);
    if (!result) return { warning: true, blocked: true, message: "トップ テン フィルターに使用できる数値データがありません。" };
    return {
      warning: false,
      message: `${result.numericCount} 件の数値から ${options.count}${options.mode === "percent" ? "%" : "項目"}を基準に抽出します。実際に表示される行: ${result.matchCount} 件。`,
    };
  }
  if (type === "auto-filter-custom") {
    const criterion = customAutoFilterCriterionFromDialog(dialog);
    if (!criterion.valid) return { warning: true, blocked: true, message: criterion.message };
    const total = Math.max(0, range.bottom - range.top);
    const matched = countRowsMatchingAutoFilterCriterion(sheet, range, col, criterion.value);
    return { warning: false, message: `${total} 件中 ${matched} 件を表示します。` };
  }
  return { warning: false, message: "" };
}



function top10OptionsFromDialog(dialog) {
  const direction = dialog.querySelector("[data-data-field='top10Direction']")?.value === "bottom" ? "bottom" : "top";
  const mode = dialog.querySelector("[data-data-field='top10Mode']")?.value === "percent" ? "percent" : "items";
  const rawCount = Number(dialog.querySelector("[data-data-field='top10Count']")?.value);
  return { direction, mode, count: Math.max(1, Math.min(500, Math.floor(Number.isFinite(rawCount) ? rawCount : 10))) };
}



function customAutoFilterCriterionFromDialog(dialog) {
  const operator1 = dialog.querySelector("[data-data-field='operator1']")?.value || "equal";
  const operator2 = dialog.querySelector("[data-data-field='operator2']")?.value || "equal";
  const value1 = dialog.querySelector("[data-data-field='value1']")?.value || "";
  const value2 = dialog.querySelector("[data-data-field='value2']")?.value || "";
  if (!String(value1).trim() && !autoFilterOperatorUsesBlankValue(operator1)) {
    return { valid: false, message: "1つ目の抽出条件を入力してください。" };
  }
  return {
    valid: true,
    value: {
      kind: "custom",
      operator1,
      value1,
      logic: dialog.querySelector("[data-data-field='logic']")?.value || "and",
      operator2,
      value2,
    },
  };
}



function countRowsMatchingAutoFilterCriterion(sheet, range, col, criterion) {
  let count = 0;
  for (let row = range.top + 1; row <= range.bottom; row += 1) {
    if (autoFilterCellMatchesCriterion(sheet, row, col, criterion)) count += 1;
  }
  return count;
}



function applyTop10AutoFilter(col, options = {}) {
  const sheet = activeSheet();
  const range = sheet?.autoFilter?.range;
  if (!sheet || !range || !col) return;
  const result = top10AutoFilterCriterionForColumn(sheet, range, col, options);
  if (!result) {
    showMessageBox("トップ テン フィルターに使用できる数値データがありません。", { title: "Cht WebSheet" });
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.autoFilter.criteria[col] = result.criterion;
  reapplyAutoFilter({ silent: true });
  commitDataOperation(
    sheet,
    sheetBefore,
    range,
    "トップ テン フィルター",
    autoFilterStatusMessage(`${columnName(col)}列にトップ テン フィルターを適用しました`, sheet, range),
  );
}



function top10AutoFilterCriterionForColumn(sheet, range, col, options = {}) {
  const numericRows = [];
  for (let row = range.top + 1; row <= range.bottom; row += 1) {
    const value = Number(autoFilterCellValue(sheet, row, col).replace(/,/g, ""));
    if (Number.isFinite(value)) numericRows.push({ row, value });
  }
  if (!numericRows.length) return null;
  const direction = options.direction === "bottom" ? "bottom" : "top";
  const mode = options.mode === "percent" ? "percent" : "items";
  const count = Math.max(1, Math.min(500, Math.floor(Number(options.count) || 10)));
  const take = mode === "percent" ? Math.max(1, Math.ceil((numericRows.length * Math.min(100, count)) / 100)) : Math.min(numericRows.length, count);
  const sorted = [...numericRows].sort((left, right) => (direction === "top" ? right.value - left.value : left.value - right.value));
  const threshold = sorted[Math.max(0, take - 1)].value;
  const criterion = {
    kind: "custom",
    operator1: direction === "top" ? "greaterThanOrEqual" : "lessThanOrEqual",
    value1: String(threshold),
    logic: "and",
    operator2: "equal",
    value2: "",
    top10: { direction, mode, count },
  };
  return {
    criterion,
    threshold,
    take,
    numericCount: numericRows.length,
    matchCount: countRowsMatchingAutoFilterCriterion(sheet, range, col, criterion),
  };
}



function closeAutoFilterMenu() {
  document.getElementById("autoFilterMenu")?.remove();
  state.autoFilterMenu = null;
}



function closeAutoFilterMenuOnEscape(event) {
  if (event.key === "Escape" && state.autoFilterMenu) closeAutoFilterMenu();
}



function closeAutoFilterMenuOnOutsidePointer(event) {
  if (!state.autoFilterMenu) return;
  if (event.target.closest?.("#autoFilterMenu") || event.target.closest?.(".cell-filter-button")) return;
  closeAutoFilterMenu();
}



function openCustomAutoFilterDialog(col, preset = null) {
  const sheet = activeSheet();
  if (!sheet?.autoFilter?.range || !col) return;
  const criterion = normalizeAutoFilterCriterion(sheet.autoFilter.criteria?.[col] || null);
  const current = { ...(criterion.kind === "custom" ? criterion : {}), ...(preset || {}) };
  state.dataDialog = { type: "auto-filter-custom", range: { ...sheet.autoFilter.range }, col };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog custom-filter-dialog" role="dialog" aria-modal="true" aria-labelledby="customFilterDialogTitle" data-data-dialog="auto-filter-custom">
      <div class="excel-mini-titlebar" id="customFilterDialogTitle">オートフィルター オプション</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-note">${escapeHtml(columnName(col))} 列で表示する行を指定します。</div>
        <div class="custom-filter-grid">
          <select data-data-field="operator1" data-auto-filter-sync>${customAutoFilterOperatorOptions(current.operator1)}</select>
          <input type="text" data-data-field="value1" data-auto-filter-sync value="${escapeAttr(current.value1 || "")}" />
          <select data-data-field="logic" data-auto-filter-sync>
            <option value="and" ${current.logic !== "or" ? "selected" : ""}>AND</option>
            <option value="or" ${current.logic === "or" ? "selected" : ""}>OR</option>
          </select>
          <select data-data-field="operator2" data-auto-filter-sync>${customAutoFilterOperatorOptions(current.operator2)}</select>
          <input type="text" data-data-field="value2" data-auto-filter-sync value="${escapeAttr(current.value2 || "")}" />
        </div>
        <div class="data-dialog-summary auto-filter-summary" data-auto-filter-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="clear">クリア</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncAutoFilterDialog(overlay.querySelector("[data-data-dialog='auto-filter-custom']"));
  overlay.querySelector("[data-data-field='value1']")?.focus({ preventScroll: true });
}



function customAutoFilterOperatorOptions(selected) {
  return [
    ["equal", "等しい"],
    ["notEqual", "等しくない"],
    ["contains", "を含む"],
    ["notContains", "を含まない"],
    ["beginsWith", "で始まる"],
    ["notBeginsWith", "で始まらない"],
    ["endsWith", "で終わる"],
    ["notEndsWith", "で終わらない"],
    ["greaterThan", "より大きい"],
    ["greaterThanOrEqual", "以上"],
    ["lessThan", "より小さい"],
    ["lessThanOrEqual", "以下"],
    ["blank", "空白"],
    ["nonBlank", "空白以外"],
  ]
    .map(([value, label]) => `<option value="${value}" ${selected === value ? "selected" : ""}>${label}</option>`)
    .join("");
}



function applyCustomAutoFilterDialog(dialog, clear = false) {
  const sheet = activeSheet();
  const range = state.dataDialog?.range;
  const col = state.dataDialog?.col;
  if (!sheet?.autoFilter?.range || !range || !col) return;
  if (!ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではオートフィルターを使用できません。")) return;
  const sheetBefore = captureSheetForHistory(sheet);
  if (clear) {
    delete sheet.autoFilter.criteria[col];
  } else {
    const criterion = customAutoFilterCriterionFromDialog(dialog);
    if (!criterion.valid) {
      showMessageBox(criterion.message || "抽出条件を入力してください。", { title: "Cht WebSheet" });
      return;
    }
    sheet.autoFilter.criteria[col] = criterion.value;
  }
  reapplyAutoFilter({ silent: true });
  closeDataDialog();
  commitDataOperation(
    sheet,
    sheetBefore,
    sheet.autoFilter.range,
    clear ? "フィルター クリア" : "カスタム フィルター",
    autoFilterStatusMessage(clear ? `${columnName(col)}列のフィルターをクリアしました` : `${columnName(col)}列にカスタム フィルターを適用しました`, sheet, sheet.autoFilter.range),
  );
}



function autoFilterOperatorUsesBlankValue(operator) {
  return operator === "blank" || operator === "nonBlank";
}



function isAutoFilterHeaderCell(sheet, row, col) {
  const filter = autoFilterStateForHeaderCell(sheet, row, col);
  const range = filter?.range;
  if (!range || row !== range.top || col < range.left || col > range.right) return false;
  return filter?.hiddenButtons?.[col] !== true;
}



function openAdvancedFilterDialog(initialRange = null) {
  const context = initialRange
    ? { sheet: activeSheet(), range: initialRange, ranges: [initialRange] }
    : selectedDataRange("詳細設定フィルター", { limit: 50000, expandCurrentRegion: true });
  if (!context) return;
  if (!ensureSheetProtectionPermission(context.sheet, "allowAutoFilter", "保護されたシートでは詳細設定フィルターを使用できません。")) return;
  const selectedRange = autoFilterRangeFromSelection(context.sheet, context.range);
  const databaseRange = advancedFilterNamedRange("Database", context.sheet) || selectedRange;
  const criteriaRange = advancedFilterNamedRange("Criteria", context.sheet);
  const extractRange = advancedFilterNamedRange("Extract", context.sheet);
  const range = {
    ...databaseRange,
    sheetIndex: databaseRange.sheetIndex ?? state.activeSheetIndex,
  };
  state.dataDialog = { type: "advanced-filter", range };
  closeDataDialog({ keepState: true });
  const listRangeValue = workbookRangeReference(range.sheetIndex ?? state.activeSheetIndex, range);
  const criteriaRangeValue = criteriaRange ? workbookRangeReference(criteriaRange.sheetIndex ?? state.activeSheetIndex, criteriaRange) : "";
  const extractRangeValue = extractRange ? workbookRangeReference(extractRange.sheetIndex ?? state.activeSheetIndex, extractRange) : "";
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog advanced-filter-dialog" role="dialog" aria-modal="true" aria-labelledby="advancedFilterTitle" data-data-dialog="advanced-filter">
      <div class="excel-mini-titlebar" id="advancedFilterTitle">フィルター オプションの設定</div>
      <div class="excel-mini-body data-dialog-body advanced-filter-body">
        <fieldset class="data-dialog-fieldset advanced-filter-action">
          <legend>処理方法</legend>
          <label class="data-dialog-check"><input type="radio" name="advancedFilterAction" data-data-field="action" value="in-place" data-advanced-filter-sync checked />選択範囲内</label>
          <label class="data-dialog-check"><input type="radio" name="advancedFilterAction" data-data-field="action" value="copy" data-advanced-filter-sync />指定した範囲</label>
        </fieldset>
        ${dataRangeFieldHtml("リスト範囲", "listRange", listRangeValue, "advanced-filter")}
        ${dataRangeFieldHtml("検索条件範囲", "criteriaRange", criteriaRangeValue, "advanced-filter")}
        ${dataRangeFieldHtml("抽出範囲", "copyToRange", extractRangeValue, "advanced-filter")}
        <label class="data-dialog-check"><input type="checkbox" data-data-field="uniqueOnly" data-advanced-filter-sync />重複するレコードは無視する</label>
        <div class="data-dialog-summary advanced-filter-summary" data-advanced-filter-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncAdvancedFilterDialog(overlay.querySelector("[data-data-dialog='advanced-filter']"));
  overlay.querySelector("[data-data-field='criteriaRange']")?.focus({ preventScroll: true });
}



function dataRangeFieldHtml(caption, field, value, syncPrefix = "") {
  const syncAttr = syncPrefix ? ` data-${syncPrefix}-sync` : "";
  return `<label class="excel-dialog-field data-range-field"><span class="excel-dialog-caption">${escapeHtml(caption)}</span><span class="data-range-input"><input type="text" data-data-field="${escapeAttr(field)}" value="${escapeAttr(value || "")}"${syncAttr} /><button class="data-range-picker" type="button" title="現在の選択範囲を使用" aria-label="${escapeAttr(caption)}に現在の選択範囲を使用" data-data-dialog-action="range-from-selection" data-data-range-target="${escapeAttr(field)}"></button></span></label>`;
}



function dataCellFieldHtml(caption, field, value, syncPrefix = "") {
  const syncAttr = syncPrefix ? ` data-${syncPrefix}-sync` : "";
  return `<label class="excel-dialog-field data-range-field"><span class="excel-dialog-caption">${escapeHtml(caption)}</span><span class="data-range-input"><input type="text" data-data-field="${escapeAttr(field)}" value="${escapeAttr(value || "")}"${syncAttr} /><button class="data-range-picker" type="button" title="現在の選択セルを使用" aria-label="${escapeAttr(caption)}に現在の選択セルを使用" data-data-dialog-action="range-from-selection" data-data-range-target="${escapeAttr(field)}" data-data-range-mode="cell"></button></span></label>`;
}



function setDialogRangeFieldFromCurrentSelection(dialog, field, mode = "range") {
  if (!dialog || !field) return;
  const isPageSetupDialog = dialog.dataset.dataDialog === "page-setup";
  const rawRange = rawActiveSelectionRanges()[0] || (state.selectionRange?.sheetIndex === state.activeSheetIndex ? state.selectionRange : null);
  const range = isPageSetupDialog
    ? rawRange || activeSelectionRange()
    : materializeSelectionRange(activeSelectionRange());
  if (!range) return;
  const input = dialog.querySelector(`[data-data-field="${cssEscape(field)}"]`) || dialog.querySelector(`[data-page-setup-field="${cssEscape(field)}"]`);
  if (!input) return;
  const reference = isPageSetupDialog
    ? pageSetupReferenceFromSelectionRange(range, mode)
    : mode === "cell"
      ? workbookCellReference(state.activeSheetIndex, range.top, range.left)
      : workbookRangeReference(state.activeSheetIndex, range);
  input.value = dialog.dataset.dataDialog === "validation" && field === "source" ? `=${reference}` : reference;
  if (dialog.dataset.dataDialog === "advanced-filter") syncAdvancedFilterDialog(dialog);
  if (dialog.dataset.dataDialog === "consolidate") syncConsolidateDialog(dialog);
  if (dialog.dataset.dataDialog === "validation") syncDataValidationDialog(dialog);
  if (dialog.dataset.dataDialog === "what-if") syncGoalSeekDialog(dialog);
  if (dialog.dataset.dataDialog === "data-table") syncDataTableDialog(dialog);
  if (dialog.dataset.dataDialog === "data-analysis") syncDataAnalysisDialog(dialog);
  if (dialog.dataset.dataDialog === "scenario-edit") syncScenarioEditDialog(dialog);
  if (dialog.dataset.dataDialog === "scenario-summary") syncScenarioSummaryDialog(dialog);
}



function pageSetupReferenceFromSelectionRange(range, mode = "range") {
  const sheet = activeSheet();
  const normalized = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left },
    { sheetIndex: state.activeSheetIndex, row: range.bottom, col: range.right },
  );
  const projected = {
    ...normalized,
    ...(range.fullRow ? { left: 1, right: Math.max(1, Number(sheet?.colCount) || normalized.right), fullRow: true } : {}),
    ...(range.fullColumn ? { top: 1, bottom: Math.max(1, Number(sheet?.rowCount) || normalized.bottom), fullColumn: true } : {}),
  };
  if (mode === "rows") return `$${projected.top}:$${projected.bottom}`;
  if (mode === "columns") return `$${columnName(projected.left)}:$${columnName(projected.right)}`;
  const start = `$${columnName(projected.left)}$${projected.top}`;
  const end = `$${columnName(projected.right)}$${projected.bottom}`;
  return start === end ? start : `${start}:${end}`;
}



function handleAdvancedFilterDialogSync(event) {
  syncAdvancedFilterDialog(event.currentTarget.closest("[data-data-dialog='advanced-filter']"));
}



function syncAdvancedFilterDialog(dialog) {
  if (!dialog) return;
  const action = dialog.querySelector("[data-data-field='action']:checked")?.value || "in-place";
  const copyEnabled = action === "copy";
  const copyToInput = dialog.querySelector("[data-data-field='copyToRange']");
  const copyToButton = dialog.querySelector("[data-data-range-target='copyToRange']");
  if (copyToInput) copyToInput.disabled = !copyEnabled;
  if (copyToButton) copyToButton.disabled = !copyEnabled;
  const summary = dialog.querySelector("[data-advanced-filter-summary]");
  const preview = advancedFilterDialogPreview(dialog);
  if (summary) {
    summary.textContent = preview.message;
    summary.classList.toggle("is-warning", preview.warning || preview.blocked);
  }
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = Boolean(preview.blocked);
}



function advancedFilterDialogPreview(dialog) {
  const listRange = decodedDialogRange(dialog.querySelector("[data-data-field='listRange']")?.value, state.dataDialog?.range);
  const criteriaRangeText = dialog.querySelector("[data-data-field='criteriaRange']")?.value || "";
  const criteriaRange = decodedOptionalDialogRange(criteriaRangeText);
  const copyToText = dialog.querySelector("[data-data-field='copyToRange']")?.value || "";
  const copyToRange = decodedOptionalDialogRange(copyToText);
  const action = dialog.querySelector("[data-data-field='action']:checked")?.value || "in-place";
  const uniqueOnly = Boolean(dialog.querySelector("[data-data-field='uniqueOnly']")?.checked);
  const listSheet = sheetForWorkbookRange(listRange);
  if (!listSheet || !listRange || listRange.bottom <= listRange.top) {
    return { warning: true, blocked: true, message: "リスト範囲を正しく入力してください。" };
  }
  if (criteriaRangeText.trim() && !criteriaRange) {
    return { warning: true, blocked: true, message: "検索条件範囲を正しく入力してください。" };
  }
  if (action === "copy" && copyToText.trim() && !copyToRange) {
    return { warning: true, blocked: true, message: "抽出範囲を正しく入力してください。" };
  }
  if (action === "copy" && !copyToRange) {
    return { warning: true, blocked: true, message: "抽出範囲を入力してください。" };
  }
  const criteriaIssues = advancedFilterCriteriaHeaderIssues(listSheet, listRange, criteriaRange);
  if (criteriaIssues.unmatched.length) {
    return {
      warning: true,
      blocked: true,
      message: `検索条件範囲の列見出しがリスト範囲に見つかりません: ${criteriaIssues.unmatched.join(", ")}`,
    };
  }
  try {
    let rows = advancedFilterMatchingRows(listSheet, listRange, { criteriaRange, uniqueOnly });
    const total = Math.max(0, listRange.bottom - listRange.top);
    let message = `${total} 件中 ${rows.length} 件が条件に一致します。`;
    if (uniqueOnly) message += " 重複するレコードは無視します。";
    if (action === "copy") {
      const outputSheet = copyToRange ? sheetForWorkbookRange(copyToRange) : null;
      const outputDetails = outputSheet && copyToRange ? advancedFilterOutputColumnDetails(listSheet, listRange, outputSheet, copyToRange) : null;
      if (outputDetails?.unmatched?.length) {
        return { warning: true, blocked: true, message: `抽出範囲の列見出しが見つかりません: ${outputDetails.unmatched.join(", ")}` };
      }
      if (!outputDetails?.cols?.length) {
        return { warning: true, blocked: true, message: "抽出する列がありません。" };
      }
      rows = advancedFilterMatchingRows(listSheet, listRange, { criteriaRange, uniqueOnly, uniqueColumns: outputDetails.cols });
      message = `${total} 件中 ${rows.length} 件が条件に一致します。`;
      if (uniqueOnly) message += " 重複するレコードは無視します。";
      const targetInfo = advancedFilterCopyTargetInfo(copyToRange, outputDetails.cols.length, rows.length, outputDetails.usesHeaderRow);
      const overwrite = targetInfo?.overwriteRange && rangeHasUserData(outputSheet, targetInfo.overwriteRange);
      message += ` 抽出列: ${outputDetails.cols.length} 列。出力先: ${rangeToLabel(targetInfo.targetRange)}。`;
      if (outputDetails.usesHeaderRow) message += " 抽出範囲の見出し行を使用します。";
      if (overwrite) message += " 出力先には既存のデータがあります。OK時に確認します。";
      return { warning: Boolean(overwrite), message };
    }
    return { warning: false, message };
  } catch (error) {
    console.warn("Advanced filter preview failed:", error);
    return { warning: true, blocked: true, message: "条件のプレビューを作成できません。" };
  }
}



function applyAdvancedFilterDialog(dialog) {
  const preview = advancedFilterDialogPreview(dialog);
  if (preview.blocked) {
    showMessageBox(preview.message || "フィルター オプションを正しく入力してください。", { title: "Cht WebSheet" });
    syncAdvancedFilterDialog(dialog);
    return;
  }
  const listRange = decodedDialogRange(dialog.querySelector("[data-data-field='listRange']")?.value, state.dataDialog?.range);
  const criteriaRange = decodedOptionalDialogRange(dialog.querySelector("[data-data-field='criteriaRange']")?.value);
  const copyToRange = decodedOptionalDialogRange(dialog.querySelector("[data-data-field='copyToRange']")?.value);
  const action = dialog.querySelector("[data-data-field='action']:checked")?.value || "in-place";
  const uniqueOnly = Boolean(dialog.querySelector("[data-data-field='uniqueOnly']")?.checked);
  const listSheet = sheetForWorkbookRange(listRange);
  if (!listSheet || !listRange || listRange.bottom <= listRange.top) {
    showMessageBox("リスト範囲を正しく入力してください。", { title: "Cht WebSheet" });
    return;
  }
  if (!ensureSheetProtectionPermission(listSheet, "allowAutoFilter", "保護されたシートでは詳細設定フィルターを使用できません。")) return;
  if (action === "copy" && !copyToRange) {
    showMessageBox("抽出範囲を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const outputSheet = action === "copy" ? sheetForWorkbookRange(copyToRange) : null;
  if (action === "copy" && !outputSheet) {
    showMessageBox("抽出範囲を正しく入力してください。", { title: "Cht WebSheet" });
    return;
  }
  if (action === "copy") {
    const outputDetails = advancedFilterOutputColumnDetails(listSheet, listRange, outputSheet, copyToRange);
    if (outputDetails.unmatched.length) {
      showMessageBox(`抽出範囲の列見出しがリスト範囲に見つかりません。\n${outputDetails.unmatched.join(", ")}`, { title: "Cht WebSheet" });
      return;
    }
    const outputInfo = advancedFilterCopyTargetInfo(
      copyToRange,
      outputDetails.cols.length,
      advancedFilterMatchingRows(listSheet, listRange, { criteriaRange, uniqueOnly, uniqueColumns: outputDetails.cols }).length,
      outputDetails.usesHeaderRow,
    );
    if (!ensureEditableRangeForProtection(outputSheet, outputInfo.targetRange, "抽出")) return;
  }
  const settings = {
    criteriaRange: criteriaRange ? { ...criteriaRange } : null,
    uniqueOnly,
  };
  if (action === "copy") {
    const copied = copyAdvancedFilterResults(listSheet, listRange, outputSheet, copyToRange, settings, { afterOverwriteConfirm: closeDataDialog });
    if (copied) closeDataDialog();
  } else {
    closeDataDialog();
    applyAdvancedFilterInPlace(listSheet, listRange, settings);
  }
}



function advancedFilterCriteriaHeaderIssues(sheet, listRange, criteriaRange) {
  if (!sheet || !listRange || !criteriaRange) return { unmatched: [] };
  const criteriaSheet = sheetForWorkbookRange(criteriaRange) || sheet;
  const criteriaSheetId = getSheetId(criteriaSheet.name);
  const headerMap = advancedFilterHeaderMap(sheet, listRange);
  const unmatched = [];
  for (let col = criteriaRange.left; col <= criteriaRange.right; col += 1) {
    const displayHeader = String(getDisplayForCell(criteriaSheet, criteriaSheetId, criteriaRange.top, col) ?? "").trim();
    const normalizedHeader = normalizeAdvancedFilterHeader(displayHeader);
    if (headerMap.has(normalizedHeader)) continue;
    let hasNonFormulaCriterion = false;
    for (let row = criteriaRange.top + 1; row <= criteriaRange.bottom; row += 1) {
      const rawCriterion = String(getCellRawInput(criteriaSheet, row, col) ?? "").trim();
      const displayCriterion = String(getDisplayForCell(criteriaSheet, criteriaSheetId, row, col) ?? "").trim();
      const criterion = rawCriterion || displayCriterion;
      if (!criterion) continue;
      if (!String(criterion).trim().startsWith("=")) hasNonFormulaCriterion = true;
    }
    if (hasNonFormulaCriterion) unmatched.push(displayHeader || columnName(col));
  }
  return { unmatched };
}



function decodedDialogRange(value, fallback = null) {
  const text = String(value || "").trim();
  const decoded = decodeWorkbookRangeOrName(text, activeSheet());
  if (!decoded && fallback) return { ...fallback, sheetIndex: fallback.sheetIndex ?? state.activeSheetIndex };
  return decoded ? { sheetIndex: state.activeSheetIndex, ...decoded } : null;
}



function decodedOptionalDialogRange(value) {
  const text = String(value || "").trim();
  return text ? decodedDialogRange(text) : null;
}



function sheetForWorkbookRange(range) {
  const sheetIndex = range?.sheetIndex ?? state.activeSheetIndex;
  return state.model?.sheets?.[sheetIndex] || null;
}



function advancedFilterNamedRange(name, baseSheet = activeSheet()) {
  const definition = validationDefinedName(name, baseSheet);
  return definition ? workbookRangeFromFormulaReference(definition.refersTo, baseSheet) : null;
}



function workbookRangeFromFormulaReference(value, baseSheet = activeSheet()) {
  const reference = validationReferenceFromFormula(baseSheet, value);
  if (!reference?.sheet || !reference.range) return null;
  const sheetIndex = (state.model?.sheets || []).indexOf(reference.sheet);
  if (sheetIndex < 0) return null;
  return { ...reference.range, sheetIndex, sheetName: reference.sheet.name };
}



function applyAdvancedFilterInPlace(sheet, listRange, settings) {
  const sheetBefore = captureSheetForHistory(sheet);
  ensureSheetSize(sheet, listRange.bottom, listRange.right);
  for (let row = listRange.top + 1; row <= listRange.bottom; row += 1) {
    if (sheet.rows[row - 1]) sheet.rows[row - 1].hidden = false;
  }
  const visibleRows = new Set(advancedFilterMatchingRows(sheet, listRange, settings).map((item) => item.row));
  for (let row = listRange.top + 1; row <= listRange.bottom; row += 1) {
    if (!visibleRows.has(row) && sheet.rows[row - 1]) sheet.rows[row - 1].hidden = true;
  }
  sheet.autoFilter = {
    range: { ...listRange },
    criteria: {},
    advancedCriteria: {
      criteriaRange: settings.criteriaRange ? { ...settings.criteriaRange } : null,
      uniqueOnly: settings.uniqueOnly,
    },
  };
  commitDataOperation(sheet, sheetBefore, listRange, `${rangeToLabel(listRange)} 詳細設定フィルター`, autoFilterStatusMessage("詳細設定フィルターを適用しました", sheet, listRange));
}



function copyAdvancedFilterResults(sheet, listRange, outputSheet, outputRange, settings, control = {}) {
  const outputDetails = advancedFilterOutputColumnDetails(sheet, listRange, outputSheet, outputRange);
  const outputCols = outputDetails.cols;
  const rows = advancedFilterMatchingRows(sheet, listRange, { ...settings, uniqueColumns: outputCols });
  const outputSheetIndex = outputRange.sheetIndex ?? state.activeSheetIndex;
  const targetInfo = advancedFilterCopyTargetInfo(outputRange, outputCols.length, rows.length, outputDetails.usesHeaderRow);
  const targetRange = targetInfo.targetRange;
  if (!control.skipOverwriteCheck && targetInfo.overwriteRange && rangeHasUserData(outputSheet, targetInfo.overwriteRange)) {
    showDataOverwriteConfirmation(() => {
      copyAdvancedFilterResults(sheet, listRange, outputSheet, outputRange, settings, { ...control, skipOverwriteCheck: true, afterOverwriteConfirm: null });
      control.afterOverwriteConfirm?.();
    });
    return false;
  }
  const sheetBefore = captureSheetForHistory(outputSheet);
  ensureSheetSize(outputSheet, targetRange.bottom, targetRange.right);
  copyAdvancedFilterRowModel(sheet, listRange.top, outputSheet, outputRange.top);
  outputCols.forEach((sourceCol, offset) => {
    copyAdvancedFilterColumnModel(sheet, sourceCol, outputSheet, outputRange.left + offset);
    writeClonedCellContent(outputSheet, outputRange.top, outputRange.left + offset, cloneCellModel(sheet.cells[cellKey(listRange.top, sourceCol)] || null));
  });
  rows.forEach((source, rowOffset) => {
    copyAdvancedFilterRowModel(sheet, source.row, outputSheet, outputRange.top + rowOffset + 1);
    outputCols.forEach((sourceCol, colOffset) => {
      writeClonedCellContent(
        outputSheet,
        outputRange.top + rowOffset + 1,
        outputRange.left + colOffset,
        cloneCellModel(sheet.cells[cellKey(source.row, sourceCol)] || null),
      );
    });
  });
  state.hf = buildFormulaEngine(state.model);
  commitDataOperationForSheet(outputSheetIndex, outputSheet, sheetBefore, listRange, `${rangeToLabel(listRange)} 詳細設定フィルター`, `${rows.length}件を抽出しました。`, targetRange);
  return true;
}



function copyAdvancedFilterColumnModel(sourceSheet, sourceCol, outputSheet, outputCol) {
  const sourceColumn = sourceSheet?.columns?.[sourceCol - 1];
  if (!sourceColumn || !outputSheet?.columns?.[outputCol - 1]) return;
  outputSheet.columns[outputCol - 1] = {
    ...defaultColumnModel(outputSheet),
    ...sourceColumn,
  };
}



function copyAdvancedFilterRowModel(sourceSheet, sourceRow, outputSheet, outputRow) {
  const sourceModel = sourceSheet?.rows?.[sourceRow - 1];
  if (!sourceModel || !outputSheet?.rows?.[outputRow - 1]) return;
  outputSheet.rows[outputRow - 1] = {
    ...defaultRowModel(outputSheet),
    ...sourceModel,
    hidden: false,
  };
}



function advancedFilterOutputColumns(sheet, listRange, outputSheet, outputRange) {
  return advancedFilterOutputColumnDetails(sheet, listRange, outputSheet, outputRange).cols;
}



function advancedFilterOutputColumnDetails(sheet, listRange, outputSheet, outputRange) {
  const requested = [];
  const unmatched = [];
  let sawHeaderText = false;
  const headerMap = advancedFilterHeaderMap(sheet, listRange);
  const outputSheetId = getSheetId(outputSheet.name);
  const singleDestinationCell = outputRange.top === outputRange.bottom && outputRange.left === outputRange.right;
  for (let col = outputRange.left; col <= outputRange.right; col += 1) {
    const displayHeader = String(getDisplayForCell(outputSheet, outputSheetId, outputRange.top, col) ?? "").trim();
    const header = normalizeAdvancedFilterHeader(displayHeader);
    if (!header) continue;
    sawHeaderText = true;
    const sourceCol = headerMap.get(header);
    if (sourceCol) requested.push(sourceCol);
    else unmatched.push(displayHeader);
  }
  if (singleDestinationCell && !requested.length) {
    const cols = [];
    for (let col = listRange.left; col <= listRange.right; col += 1) cols.push(col);
    return { cols, unmatched: [], usesHeaderRow: false };
  }
  if (requested.length || (sawHeaderText && unmatched.length)) return { cols: requested, unmatched, usesHeaderRow: true };
  const cols = [];
  for (let col = listRange.left; col <= listRange.right; col += 1) cols.push(col);
  return { cols, unmatched: [], usesHeaderRow: false };
}



function advancedFilterCopyTargetInfo(outputRange, colCount, rowCount, usesHeaderRow = false) {
  const width = Math.max(1, Number(colCount) || 1);
  const outputSheetIndex = outputRange.sheetIndex ?? state.activeSheetIndex;
  const targetRange = normalizeRange(
    { sheetIndex: outputSheetIndex, row: outputRange.top, col: outputRange.left },
    { sheetIndex: outputSheetIndex, row: outputRange.top + Math.max(1, Number(rowCount) + 1) - 1, col: outputRange.left + width - 1 },
  );
  const overwriteRange = usesHeaderRow
    ? (rowCount > 0
        ? normalizeRange(
            { sheetIndex: outputSheetIndex, row: outputRange.top + 1, col: outputRange.left },
            { sheetIndex: outputSheetIndex, row: outputRange.top + rowCount, col: outputRange.left + width - 1 },
          )
        : null)
    : targetRange;
  return { targetRange, overwriteRange };
}



function advancedFilterMatchingRows(sheet, listRange, settings = {}) {
  const rows = [];
  const criteriaRows = advancedFilterCriteriaRows(sheet, listRange, settings.criteriaRange);
  const seen = new Set();
  const uniqueColumns = Array.isArray(settings.uniqueColumns) && settings.uniqueColumns.length ? settings.uniqueColumns : null;
  for (let row = listRange.top + 1; row <= listRange.bottom; row += 1) {
    if (!advancedFilterRowMatches(sheet, row, criteriaRows)) continue;
    if (settings.uniqueOnly) {
      const signature = advancedFilterRowSignature(sheet, listRange, row, uniqueColumns);
      if (seen.has(signature)) continue;
      seen.add(signature);
    }
    rows.push({ row });
  }
  return rows;
}



function advancedFilterCriteriaRows(sheet, listRange, criteriaRange) {
  if (!criteriaRange) return [];
  const criteriaSheet = sheetForWorkbookRange(criteriaRange) || sheet;
  const headerMap = advancedFilterHeaderMap(sheet, listRange);
  const criteriaSheetId = getSheetId(criteriaSheet.name);
  const criterionColumns = [];
  for (let col = criteriaRange.left; col <= criteriaRange.right; col += 1) {
    const header = normalizeAdvancedFilterHeader(getDisplayForCell(criteriaSheet, criteriaSheetId, criteriaRange.top, col));
    criterionColumns.push({
      criteriaCol: col,
      listCol: headerMap.get(header) || null,
      allowFormula: !headerMap.has(header),
    });
  }
  const rows = [];
  for (let row = criteriaRange.top + 1; row <= criteriaRange.bottom; row += 1) {
    const conditions = [];
    criterionColumns.forEach((item) => {
      const rawCriterion = String(getCellRawInput(criteriaSheet, row, item.criteriaCol) ?? "").trim();
      const displayCriterion = String(getDisplayForCell(criteriaSheet, criteriaSheetId, row, item.criteriaCol) ?? "").trim();
      const criterion = rawCriterion || displayCriterion;
      if (!criterion) return;
      if (item.allowFormula && criterion.startsWith("=")) {
        conditions.push({ formula: criterion, anchorRow: listRange.top + 1 });
        return;
      }
      if (item.listCol) conditions.push({ col: item.listCol, criterion: advancedFilterCellCriterion(rawCriterion, displayCriterion) });
    });
    if (conditions.length) rows.push(conditions);
  }
  return rows;
}



function advancedFilterCellCriterion(rawCriterion, displayCriterion) {
  const rawText = String(rawCriterion ?? "").trim();
  const displayText = String(displayCriterion ?? "").trim();
  if (rawText.startsWith("'")) return rawText.slice(1);
  const formulaLiteral = advancedFilterFormulaStringLiteral(rawText);
  if (formulaLiteral != null) return formulaLiteral;
  return displayText || rawText;
}



function advancedFilterFormulaStringLiteral(value) {
  const text = String(value || "").trim();
  const match = text.match(/^=\s*"((?:[^"]|"")*)"$/);
  return match ? match[1].replace(/""/g, '"') : null;
}



function advancedFilterHeaderMap(sheet, listRange) {
  const map = new Map();
  const sheetId = getSheetId(sheet.name);
  for (let col = listRange.left; col <= listRange.right; col += 1) {
    const display = getDisplayForCell(sheet, sheetId, listRange.top, col);
    const normalized = normalizeAdvancedFilterHeader(display || columnName(col));
    if (!map.has(normalized)) map.set(normalized, col);
  }
  return map;
}



function normalizeAdvancedFilterHeader(value) {
  return String(value ?? "").trim().toLocaleLowerCase("ja-JP");
}



function advancedFilterRowMatches(sheet, row, criteriaRows) {
  if (!criteriaRows.length) return true;
  return criteriaRows.some((conditions) => conditions.every((condition) => {
    if (condition.formula) return advancedFilterFormulaMatches(sheet, row, condition);
    const value = autoFilterCellValue(sheet, row, condition.col);
    return advancedCriterionMatches(value, condition.criterion);
  }));
}



function advancedFilterFormulaMatches(sheet, row, condition) {
  const sheetId = getSheetId(sheet.name);
  if (!state.hf?.calculateFormula || sheetId == null) return false;
  const formula = String(condition.formula || "").trim();
  if (!formula) return true;
  const normalized = formula.startsWith("=") ? formula : `=${formula}`;
  const rowOffset = row - Number(condition.anchorRow || row);
  const translated = rowOffset ? translateValidationFormulaReferences(normalized, rowOffset, 0) : normalized;
  try {
    return validationFormulaResultTruthy(state.hf.calculateFormula(translateFormulaForCurrentFormulaEngine(translated, formulaEngineSheetIndex(sheet)), sheetId));
  } catch (error) {
    console.warn("Advanced filter formula criterion failed:", error);
    return false;
  }
}



function advancedCriterionMatches(value, criterion) {
  const text = String(criterion ?? "").trim();
  const match = text.match(/^(>=|<=|<>|>|<|=)(.*)$/);
  const explicitOperator = Boolean(match);
  const operator = match ? match[1] : "=";
  const expected = (match ? match[2] : text).trim();
  const actualText = String(value ?? "");
  const actualNumber = Number(actualText.replace(/,/g, ""));
  const expectedNumber = Number(expected.replace(/,/g, ""));
  const canCompareNumber = Number.isFinite(actualNumber) && Number.isFinite(expectedNumber);
  const actualDate = autoFilterDateSerial(actualText);
  const expectedDate = autoFilterDateSerial(expected);
  const canCompareDate = actualDate != null && expectedDate != null;
  if ((operator === "=" || operator === "<>") && /[*?]/.test(expected)) {
    const matches = advancedWildcardToRegExp(expected).test(actualText);
    return operator === "<>" ? !matches : matches;
  }
  if (operator === "=" && expected === "") return actualText === "";
  if (operator === "<>" && expected === "") return actualText !== "";
  if (!explicitOperator && !canCompareNumber && !canCompareDate) {
    return actualText.toLocaleLowerCase("ja-JP").startsWith(expected.toLocaleLowerCase("ja-JP"));
  }
  switch (operator) {
    case "<>":
      return canCompareNumber ? actualNumber !== expectedNumber : canCompareDate ? actualDate !== expectedDate : actualText.localeCompare(expected, "ja", { sensitivity: "base" }) !== 0;
    case ">":
      return canCompareNumber ? actualNumber > expectedNumber : canCompareDate ? actualDate > expectedDate : actualText.localeCompare(expected, "ja", { numeric: true, sensitivity: "base" }) > 0;
    case ">=":
      return canCompareNumber ? actualNumber >= expectedNumber : canCompareDate ? actualDate >= expectedDate : actualText.localeCompare(expected, "ja", { numeric: true, sensitivity: "base" }) >= 0;
    case "<":
      return canCompareNumber ? actualNumber < expectedNumber : canCompareDate ? actualDate < expectedDate : actualText.localeCompare(expected, "ja", { numeric: true, sensitivity: "base" }) < 0;
    case "<=":
      return canCompareNumber ? actualNumber <= expectedNumber : canCompareDate ? actualDate <= expectedDate : actualText.localeCompare(expected, "ja", { numeric: true, sensitivity: "base" }) <= 0;
    case "=":
    default:
      return canCompareNumber ? actualNumber === expectedNumber : canCompareDate ? actualDate === expectedDate : actualText.localeCompare(expected, "ja", { sensitivity: "base" }) === 0;
  }
}



function advancedWildcardToRegExp(pattern) {
  let escaped = "";
  const source = String(pattern);
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char === "~" && index + 1 < source.length) {
      escaped += source[index + 1].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      index += 1;
    } else if (char === "*") {
      escaped += ".*";
    } else if (char === "?") {
      escaped += ".";
    } else {
      escaped += char.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
  }
  return new RegExp(`^${escaped}$`, "i");
}



function advancedFilterRowSignature(sheet, listRange, row, columns = null) {
  const sheetId = getSheetId(sheet.name);
  const values = [];
  const signatureColumns = Array.isArray(columns) && columns.length
    ? columns
    : Array.from({ length: Math.max(0, listRange.right - listRange.left + 1) }, (_, index) => listRange.left + index);
  signatureColumns.forEach((col) => {
    const cell = sheet.cells[cellKey(row, col)];
    const rawDate = comparableDateSerial(cell?.formula ? cell?.cached : cell?.raw ?? cell?.cached);
    if (rawDate != null) {
      values.push(`date:${rawDate}`);
      return;
    }
    const display = String(getDisplayForCell(sheet, sheetId, row, col) ?? "");
    const date = comparableDateSerial(display);
    values.push(date == null ? display.toLocaleLowerCase("ja-JP") : `date:${date}`);
  });
  return values.join("\u0001");
}



function openTextToColumnsDialog() {
  const context = selectedDataRange("区切り位置", { singleColumn: true });
  if (!context) return;
  state.dataDialog = { type: "text-to-columns", range: context.range, step: 1, textColumnsSelectedColumn: 0, columnFormats: {} };
  closeDataDialog({ keepState: true });
  const previewRows = textToColumnsSourceRows(context.sheet, context.range, 8);
  const delimiterDefaults = initialTextColumnsDelimiterOptions(previewRows);
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog text-columns-dialog" role="dialog" aria-modal="true" aria-labelledby="textColumnsDialogTitle" data-data-dialog="text-to-columns">
      <div class="excel-mini-titlebar" id="textColumnsDialogTitle">区切り位置指定ウィザード</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="text-columns-stepbar">
          <strong data-text-columns-heading>元のデータ形式</strong>
          <span data-text-columns-counter>1 / 3</span>
        </div>
        <section class="text-columns-step is-active" data-text-columns-step="1">
          <div class="data-dialog-note">選択したデータに最も近いデータ形式を選択してください。</div>
          <fieldset class="data-dialog-fieldset text-columns-mode">
            <legend>元のデータの形式</legend>
            <label class="data-dialog-check"><input type="radio" name="textColumnsMode" data-data-field="mode" data-text-columns-sync value="delimited" checked />区切り記号付き</label>
            <label class="data-dialog-check"><input type="radio" name="textColumnsMode" data-data-field="mode" data-text-columns-sync value="fixed" />固定長</label>
          </fieldset>
          <div class="text-columns-preview-source">${textColumnsSourcePreviewHtml(previewRows)}</div>
        </section>
        <section class="text-columns-step" data-text-columns-step="2">
          <div class="text-columns-delimited-options" data-text-columns-mode-panel="delimited">
            <div class="data-dialog-note">データを区切る文字を指定してください。</div>
            <div class="text-columns-delimiters">
              <label class="data-dialog-check"><input type="checkbox" data-data-field="tab" data-text-columns-sync ${delimiterDefaults.tab ? "checked" : ""} />タブ</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="semicolon" data-text-columns-sync ${delimiterDefaults.semicolon ? "checked" : ""} />セミコロン</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="comma" data-text-columns-sync ${delimiterDefaults.comma ? "checked" : ""} />カンマ</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="space" data-text-columns-sync ${delimiterDefaults.space ? "checked" : ""} />スペース</label>
              <label class="data-dialog-check text-columns-other"><input type="checkbox" data-data-field="other" data-text-columns-sync />その他<input type="text" maxlength="1" data-data-field="otherDelimiter" data-text-columns-sync value="" /></label>
            </div>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="consecutive" data-text-columns-sync checked />連続した区切り文字は1文字として扱う</label>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">文字列の引用符</span><select data-data-field="qualifier" data-text-columns-sync><option value="&quot;">"</option><option value="'">'</option><option value="">なし</option></select></label>
          </div>
          <div class="text-columns-fixed-options" data-text-columns-mode-panel="fixed">
            <div class="data-dialog-note">列の区切り位置を文字数で指定してください。例: 8, 16, 24</div>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">区切り位置</span><input type="text" data-data-field="fixedBreaks" data-text-columns-sync value="" /></label>
          </div>
          <div class="text-columns-preview" data-text-columns-preview></div>
        </section>
        <section class="text-columns-step" data-text-columns-step="3">
          <div class="data-dialog-note">列を選択してデータ形式を指定します。</div>
          <fieldset class="data-dialog-fieldset text-columns-format">
            <legend>列のデータ形式</legend>
            <label class="data-dialog-check"><input type="radio" name="textColumnsFormat" data-data-field="columnFormat" data-text-columns-sync value="general" checked />G/標準</label>
            <label class="data-dialog-check"><input type="radio" name="textColumnsFormat" data-data-field="columnFormat" data-text-columns-sync value="text" />文字列</label>
            <label class="data-dialog-check text-columns-date-format">
              <input type="radio" name="textColumnsFormat" data-data-field="columnFormat" data-text-columns-sync value="date" />
              <span>日付</span>
              <select data-data-field="dateOrder" data-text-columns-sync aria-label="日付の順序">
                ${textColumnsDateOrderOptions("YMD")}
              </select>
            </label>
            <label class="data-dialog-check"><input type="radio" name="textColumnsFormat" data-data-field="columnFormat" data-text-columns-sync value="skip" />削除する</label>
          </fieldset>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">表示先</span><input type="text" data-data-field="destination" data-text-columns-sync value="${escapeAttr(`${columnName(context.range.left)}${context.range.top}`)}" /></label>
          <fieldset class="data-dialog-fieldset text-columns-advanced">
            <legend>詳細設定</legend>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">小数点の記号</span><input type="text" maxlength="1" data-data-field="decimalSeparator" data-text-columns-sync value="." /></label>
            <label class="excel-dialog-field"><span class="excel-dialog-caption">桁区切り記号</span><input type="text" maxlength="1" data-data-field="thousandsSeparator" data-text-columns-sync value="," /></label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="trailingMinus" data-text-columns-sync />末尾のマイナス符号を負の数として扱う</label>
          </fieldset>
          <div class="text-columns-preview" data-text-columns-preview></div>
        </section>
        <div class="data-dialog-summary text-columns-summary" data-text-columns-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="text-columns-back">戻る(&lt;B)</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="text-columns-next">次へ(&gt;N)</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="text-columns-finish">完了(F)</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncTextToColumnsDialog(overlay.querySelector("[data-data-dialog='text-to-columns']"));
}



function applyTextToColumnsDialog(dialog) {
  const context = state.dataDialog;
  if (!context?.range) return;
  saveTextColumnsColumnFormat(dialog);
  const validation = textColumnsDialogValidation(dialog);
  if (!validation.valid) {
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    syncTextToColumnsDialog(dialog);
    return;
  }
  const groupResult = textToColumnsForSheetGroup(context.range, validation.options, validation.destination, {
    activeTargetRange: validation.targetRange,
    afterOverwriteConfirm: closeDataDialog,
  });
  if (groupResult === "pending") return;
  if (groupResult) {
    closeDataDialog();
    return;
  }
  if (textToColumns(context.range, validation.options, validation.destination, {
    targetRange: validation.targetRange,
    outputRows: validation.outputRows,
    afterOverwriteConfirm: closeDataDialog,
  })) {
    closeDataDialog();
  }
}



function initialTextColumnsDelimiterOptions(rows) {
  const sample = (rows || []).map((row) => String(row ?? "")).filter(Boolean).join("\n");
  const counts = [
    ["\t", (sample.match(/\t/g) || []).length],
    [";", (sample.match(/;/g) || []).length],
    [",", (sample.match(/,/g) || []).length],
    [" ", (sample.match(/ /g) || []).length],
  ].sort((left, right) => right[1] - left[1]);
  const detected = counts[0]?.[1] > 0 ? counts[0][0] : ",";
  if (detected === "\t") return { tab: true, semicolon: false, comma: false, space: false };
  if (detected === ";") return { tab: false, semicolon: true, comma: false, space: false };
  if (detected === ",") return { tab: false, semicolon: false, comma: true, space: false };
  if (detected === " ") return { tab: false, semicolon: false, comma: false, space: true };
  return { tab: true, semicolon: false, comma: true, space: false };
}



function handleTextColumnsPreviewChange(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='text-to-columns']");
  if (!dialog) return;
  if (event.currentTarget.dataset.dataField === "fixedBreaks") {
    state.dataDialog.fixedBreaksTouched = true;
  }
  if (event.currentTarget.dataset.dataField === "columnFormat" || event.currentTarget.dataset.dataField === "dateOrder") {
    saveTextColumnsColumnFormat(dialog);
  }
  syncTextToColumnsDialog(dialog);
}



function syncTextToColumnsDialog(dialog) {
  if (!dialog || state.dataDialog?.type !== "text-to-columns") return;
  const context = state.dataDialog;
  const step = Math.min(3, Math.max(1, Number(context.step) || 1));
  context.step = step;
  ensureTextColumnsFixedBreaks(dialog, context);
  const options = readTextColumnsOptions(dialog);
  const headings = ["元のデータ形式", options.mode === "fixed" ? "フィールド幅の設定" : "区切り文字の設定", "列のデータ形式"];
  dialog.querySelector("[data-text-columns-heading]").textContent = headings[step - 1];
  dialog.querySelector("[data-text-columns-counter]").textContent = `${step} / 3`;
  dialog.querySelectorAll("[data-text-columns-step]").forEach((panel) => {
    panel.classList.toggle("is-active", Number(panel.dataset.textColumnsStep) === step);
  });
  dialog.querySelectorAll("[data-text-columns-mode-panel]").forEach((panel) => {
    panel.classList.toggle("is-hidden", panel.dataset.textColumnsModePanel !== options.mode);
  });
  const otherInput = dialog.querySelector("[data-data-field='otherDelimiter']");
  if (otherInput) otherInput.disabled = !dialog.querySelector("[data-data-field='other']")?.checked;
  const sourceRows = textToColumnsSourceRows(activeSheet(), context.range, 10).map((row) => String(row ?? ""));
  const rows = textToColumnsRowsForRange(activeSheet(), context.range, options).slice(0, 10);
  const maxColumns = Math.max(1, ...rows.map((row) => row.length));
  context.textColumnsSelectedColumn = Math.min(Math.max(0, Number(context.textColumnsSelectedColumn) || 0), maxColumns - 1);
  const selectedFormat = context.columnFormats?.[context.textColumnsSelectedColumn] || "general";
  const selectedFormatRadio = selectedFormat.startsWith("date") ? "date" : selectedFormat;
  dialog.querySelectorAll("[data-data-field='columnFormat']").forEach((input) => {
    input.checked = input.value === selectedFormatRadio;
  });
  const dateOrder = dialog.querySelector("[data-data-field='dateOrder']");
  if (dateOrder) {
    dateOrder.value = textColumnsDateOrderFromFormat(selectedFormat);
    dateOrder.disabled = selectedFormatRadio !== "date";
  }
  dialog.querySelectorAll("[data-text-columns-preview]").forEach((preview) => {
    preview.innerHTML = options.mode === "fixed" && step === 2
      ? textColumnsFixedPreviewHtml(sourceRows, options.fixedBreaks)
      : textColumnsPreviewHtml(rows, context.textColumnsSelectedColumn, context.columnFormats || {});
  });
  dialog.querySelector("[data-data-dialog-action='text-columns-back']").disabled = step <= 1;
  const stepValidation = textColumnsStepValidation(dialog, step, options);
  dialog.querySelector("[data-data-dialog-action='text-columns-next']").disabled = step >= 3 || !stepValidation.valid;
  const validation = textColumnsDialogValidation(dialog);
  const summary = dialog.querySelector("[data-text-columns-summary]");
  if (summary) {
    summary.textContent = stepValidation.valid ? validation.message : stepValidation.message;
    summary.classList.toggle("is-warning", !stepValidation.valid || !validation.valid || validation.warning);
  }
  const finish = dialog.querySelector("[data-data-dialog-action='text-columns-finish']");
  if (finish) finish.disabled = !stepValidation.valid || !validation.valid;
}



function textColumnsStepValidation(dialog, step, options = readTextColumnsOptions(dialog)) {
  if (step === 2 && options.mode === "delimited" && !options.delimiters.length) {
    return { valid: false, message: "区切り文字を選択してください。" };
  }
  if (step === 2 && options.mode === "fixed") {
    const sourceRows = textToColumnsSourceRows(activeSheet(), state.dataDialog?.range || {}, 20).map((row) => String(row ?? ""));
    const maxLength = Math.max(0, ...sourceRows.map((row) => row.length));
    const invalidBreaks = (options.fixedBreaks || []).filter((position) => position <= 0 || position >= maxLength);
    if (invalidBreaks.length) return { valid: false, message: "区切り位置はデータの文字数より小さい値を指定してください。" };
  }
  return { valid: true, message: "" };
}



function textColumnsDialogValidation(dialog) {
  const context = state.dataDialog;
  const sheet = activeSheet();
  if (!dialog || !context?.range || !sheet) return { valid: false, message: "変換する範囲がありません。" };
  const options = readTextColumnsOptions(dialog);
  const stepValidation = textColumnsStepValidation(dialog, Number(context.step) || 1, options);
  if (!stepValidation.valid) return { valid: false, options, message: stepValidation.message };
  if (options.decimalSeparator && options.thousandsSeparator && options.decimalSeparator === options.thousandsSeparator) {
    return { valid: false, options, message: "小数点の記号と桁区切り記号には異なる文字を指定してください。" };
  }
  if (options.mode === "delimited" && !options.delimiters.length) {
    return { valid: false, options, message: "区切り文字を選択してください。" };
  }
  const destination = textColumnsDestinationFromDialog(dialog);
  if (!destination) return { valid: false, options, message: "表示先を正しく入力してください。" };
  const rows = textToColumnsRowsForRange(sheet, context.range, options);
  const outputRows = textColumnsOutputRows(rows, options);
  const outputHeight = outputRows.length;
  const outputWidth = Math.max(1, ...outputRows.map((row) => row.length));
  if (!outputRows.some((row) => row.length)) {
    return { valid: false, options, destination, message: "出力する列がありません。列のデータ形式を確認してください。" };
  }
  const targetRange = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: destination.row, col: destination.col },
    { sheetIndex: state.activeSheetIndex, row: destination.row + outputHeight - 1, col: destination.col + outputWidth - 1 },
  );
  const overwrite = rangeHasUserData(sheet, targetRange, context.range);
  const oneColumn = outputWidth <= 1;
  const warning = overwrite || oneColumn;
  const overwriteText = overwrite ? " 表示先には既存のデータがあります。完了時に確認します。" : "";
  const oneColumnText = oneColumn ? " 変換結果は1列のままです。" : "";
  return {
    valid: true,
    warning,
    options,
    destination,
    targetRange,
    outputRows,
    message: `${outputHeight.toLocaleString("ja-JP")} 行 x ${outputWidth.toLocaleString("ja-JP")} 列を ${rangeToLabel(targetRange)} に出力します。${overwriteText}${oneColumnText}`.trim(),
  };
}



function textColumnsOutputRows(rows, options) {
  return (rows || []).map((row) => (row || [])
    .map((value, index) => ({ value, format: options.columnFormats?.[index] || "general" }))
    .filter((item) => item.format !== "skip"));
}



function textColumnsDestinationFromDialog(dialog) {
  const text = String(dialog?.querySelector("[data-data-field='destination']")?.value || "").trim();
  if (!text) return null;
  const workbookRange = decodeWorkbookRangeOrName(text, activeSheet());
  if (workbookRange) {
    const sheetIndex = workbookRange.sheetIndex ?? state.activeSheetIndex;
    if (sheetIndex !== state.activeSheetIndex) return null;
    return { row: workbookRange.top, col: workbookRange.left };
  }
  return decodeCell(text.replace(/\$/g, ""));
}



function ensureTextColumnsFixedBreaks(dialog, context) {
  const mode = dialog.querySelector("[data-data-field='mode']:checked")?.value || "delimited";
  if (mode !== "fixed" || context.fixedBreaksTouched) return;
  const input = dialog.querySelector("[data-data-field='fixedBreaks']");
  if (!input || String(input.value || "").trim()) return;
  const breaks = inferTextColumnsFixedBreaks(activeSheet(), context.range);
  if (breaks.length) input.value = breaks.join(", ");
}



function inferTextColumnsFixedBreaks(sheet, range) {
  if (!sheet || !range) return [];
  const rows = textToColumnsSourceRows(sheet, range, 20).map((value) => String(value ?? "")).filter(Boolean);
  if (!rows.length) return [];
  const counts = new Map();
  rows.forEach((row) => {
    const breaks = fixedBreakCandidatesForText(row);
    breaks.forEach((position) => counts.set(position, (counts.get(position) || 0) + 1));
  });
  const threshold = Math.max(1, Math.ceil(rows.length * 0.55));
  return [...counts.entries()]
    .filter(([position, count]) => position > 0 && count >= threshold)
    .map(([position]) => position)
    .sort((left, right) => left - right)
    .slice(0, 12);
}



function fixedBreakCandidatesForText(value) {
  const text = String(value ?? "");
  const candidates = new Set();
  for (let index = 1; index < text.length; index += 1) {
    const previous = fixedBreakCharacterClass(text[index - 1]);
    const current = fixedBreakCharacterClass(text[index]);
    if (!previous || !current || previous === current) continue;
    candidates.add(index);
  }
  text.replace(/ {2,}/g, (match, offset) => {
    if (offset > 0) candidates.add(offset);
    if (offset + match.length < text.length) candidates.add(offset + match.length);
    return match;
  });
  return [...candidates].sort((left, right) => left - right);
}



function fixedBreakCharacterClass(char) {
  if (/\s/.test(char)) return "";
  if (/[0-9０-９]/.test(char)) return "number";
  if (/[A-Za-zＡ-Ｚａ-ｚ]/.test(char)) return "latin";
  if (/[\u3040-\u30ff\u3400-\u9fff]/.test(char)) return "jp";
  return "symbol";
}



function saveTextColumnsColumnFormat(dialog) {
  if (state.dataDialog?.type !== "text-to-columns") return;
  const selected = Number(state.dataDialog.textColumnsSelectedColumn) || 0;
  let format = dialog?.querySelector("[data-data-field='columnFormat']:checked")?.value;
  if (!format) return;
  if (format === "date") format = `date${textColumnsDateOrderFromFormat(dialog?.querySelector("[data-data-field='dateOrder']")?.value || "YMD")}`;
  state.dataDialog.columnFormats = state.dataDialog.columnFormats || {};
  state.dataDialog.columnFormats[selected] = format;
}



function readTextColumnsOptions(dialog) {
  const mode = dialog.querySelector("[data-data-field='mode']:checked")?.value || "delimited";
  const delimiters = [];
  if (dialog.querySelector("[data-data-field='tab']")?.checked) delimiters.push("\t");
  if (dialog.querySelector("[data-data-field='semicolon']")?.checked) delimiters.push(";");
  if (dialog.querySelector("[data-data-field='comma']")?.checked) delimiters.push(",");
  if (dialog.querySelector("[data-data-field='space']")?.checked) delimiters.push(" ");
  if (dialog.querySelector("[data-data-field='other']")?.checked) {
    const other = dialog.querySelector("[data-data-field='otherDelimiter']")?.value || "";
    if (other) delimiters.push(other[0]);
  }
  return {
    mode,
    delimiters,
    consecutive: Boolean(dialog.querySelector("[data-data-field='consecutive']")?.checked),
    qualifier: dialog.querySelector("[data-data-field='qualifier']")?.value ?? '"',
    fixedBreaks: parseFixedWidthBreaks(dialog.querySelector("[data-data-field='fixedBreaks']")?.value || ""),
    columnFormats: { ...(state.dataDialog?.columnFormats || {}) },
    decimalSeparator: dialog.querySelector("[data-data-field='decimalSeparator']")?.value || ".",
    thousandsSeparator: dialog.querySelector("[data-data-field='thousandsSeparator']")?.value || ",",
    trailingMinus: Boolean(dialog.querySelector("[data-data-field='trailingMinus']")?.checked),
  };
}



function textToColumnsSourceRows(sheet, range, limit = 8) {
  const rows = [];
  for (let row = range.top; row <= Math.min(range.bottom, range.top + limit - 1); row += 1) {
    rows.push(getCellRawInput(sheet, row, range.left));
  }
  return rows;
}



function textColumnsSourcePreviewHtml(rows) {
  return `<div class="text-columns-raw-preview">${rows.map((row) => `<div>${escapeHtml(row || "\u00a0")}</div>`).join("")}</div>`;
}



function textColumnsFixedPreviewHtml(rows, breaks) {
  const cleanRows = (rows || []).map((row) => String(row ?? ""));
  const maxLength = Math.max(2, ...cleanRows.map((row) => row.length));
  const breakSet = new Set((breaks || []).filter((position) => position > 0 && position < maxLength));
  const markers = [];
  for (let position = 1; position < maxLength; position += 1) {
    markers.push(`
      <button class="text-columns-fixed-break${breakSet.has(position) ? " is-break" : ""}" type="button"
        data-data-dialog-action="text-columns-toggle-fixed-break" data-text-fixed-position="${position}"
        style="left:calc(${position}ch + 6px)" title="${breakSet.has(position) ? "区切り位置を削除" : "区切り位置を追加"}">
        <span aria-hidden="true"></span>
      </button>`);
  }
  const rowsHtml = cleanRows.map((row) => {
    const cells = [];
    for (let index = 0; index < maxLength; index += 1) {
      const position = index + 1;
      cells.push(`<span class="${breakSet.has(position) ? "is-break-after" : ""}">${escapeHtml(row[index] || "\u00a0")}</span>`);
    }
    return `<div class="text-columns-fixed-row">${cells.join("")}</div>`;
  }).join("");
  return `
    <div class="text-columns-fixed-preview" style="--text-columns-fixed-width:${maxLength}ch">
      <div class="text-columns-fixed-ruler" aria-label="固定長の区切り位置">${markers.join("")}</div>
      <div class="text-columns-fixed-rows">${rowsHtml}</div>
    </div>`;
}



function textColumnsPreviewHtml(rows, selectedColumn = 0, columnFormats = {}) {
  const width = Math.max(1, ...rows.map((row) => row.length));
  const headers = Array.from({ length: width }, (_item, index) => {
    const format = columnFormats[index] || "general";
    return `<th><button class="${index === selectedColumn ? "is-selected" : ""}" type="button" data-data-dialog-action="text-columns-select-column" data-text-column-index="${index}">${index + 1}<span>${escapeHtml(textColumnsFormatLabel(format))}</span></button></th>`;
  }).join("");
  const body = rows.map((row) => `<tr>${Array.from({ length: width }, (_item, index) => `<td>${escapeHtml(row[index] ?? "")}</td>`).join("")}</tr>`).join("");
  return `<table class="text-columns-preview-table"><thead><tr>${headers}</tr></thead><tbody>${body}</tbody></table>`;
}



function textColumnsFormatLabel(format) {
  switch (format) {
    case "text":
      return "文字列";
    case "skip":
      return "削除";
    case "general":
    default:
      if (String(format || "").startsWith("date")) return `日付(${textColumnsDateOrderFromFormat(format)})`;
      return "標準";
  }
}



function textColumnsDateOrderOptions(selected = "YMD") {
  return ["MDY", "DMY", "YMD", "MYD", "DYM", "YDM"]
    .map((order) => `<option value="${order}" ${textColumnsDateOrderFromFormat(selected) === order ? "selected" : ""}>${order}</option>`)
    .join("");
}



function textColumnsDateOrderFromFormat(format) {
  const value = String(format || "").replace(/^date/i, "").toUpperCase();
  return ["MDY", "DMY", "YMD", "MYD", "DYM", "YDM"].includes(value) ? value : "YMD";
}



function textToColumns(range, options, destination, control = {}) {
  const sheet = control.sheet || activeSheet();
  const sheetIndex = Number.isFinite(Number(control.sheetIndex)) ? Number(control.sheetIndex) : state.activeSheetIndex;
  const rows = textToColumnsRowsForRange(sheet, range, options);
  const outputRows = control.outputRows || textColumnsOutputRows(rows, options);
  if (!outputRows.some((row) => row.length)) {
    showMessageBox("出力する列がありません。列のデータ形式を確認してください。", { title: "Cht WebSheet" });
    return false;
  }
  const maxParts = Math.max(1, ...outputRows.map((row) => row.length));
  const targetRange = control.targetRange || normalizeRange(
    { sheetIndex, row: destination.row, col: destination.col },
    { sheetIndex, row: destination.row + outputRows.length - 1, col: destination.col + maxParts - 1 },
  );
  if (!ensureEditableRangeForProtection(sheet, targetRange, "分割")) return false;
  if (!control.skipOverwriteCheck && rangeHasUserData(sheet, targetRange, range)) {
    showDataOverwriteConfirmation(() => {
      textToColumns(range, options, destination, { ...control, skipOverwriteCheck: true, afterOverwriteConfirm: null });
      control.afterOverwriteConfirm?.();
    });
    return false;
  }
  const sheetBefore = control.commitHistory === false ? null : captureSheetForHistory(sheet);
  ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
  const sheetId = getSheetId(sheet.name);
  outputRows.forEach((parts, rowIndex) => {
    for (let colOffset = 0; colOffset < maxParts; colOffset += 1) {
      const item = parts[colOffset] || { value: "", format: "general" };
      writeCellContent(sheet, sheetId, targetRange.top + rowIndex, targetRange.left + colOffset, coerceTextColumnsValue(item.value, item.format, options));
    }
  });
  internCellsInRange(sheet, targetRange);
  if (control.commitHistory === false) return { targetRange, outputRows };
  commitDataOperationForSheet(sheetIndex, sheet, sheetBefore, targetRange, `${rangeToLabel(range)} 区切り位置`, `${rangeToLabel(range)} を分割しました。`, targetRange);
  return true;
}



function textToColumnsForSheetGroup(range, options, destination, control = {}) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !range || !destination) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    const sourceRange = materializeSelectionRangeForSheet(sheet, range, sheetIndex);
    if (!sourceRange) continue;
    const rows = textToColumnsRowsForRange(sheet, sourceRange, options);
    const outputRows = textColumnsOutputRows(rows, options);
    if (!outputRows.some((row) => row.length)) {
      showMessageBox(`${sheet.name}：出力する列がありません。列のデータ形式を確認してください。`, { title: "Cht WebSheet" });
      return true;
    }
    const maxParts = Math.max(1, ...outputRows.map((row) => row.length));
    const targetRange = normalizeRange(
      { sheetIndex, row: destination.row, col: destination.col },
      { sheetIndex, row: destination.row + outputRows.length - 1, col: destination.col + maxParts - 1 },
    );
    if (!ensureEditableRangeForProtection(sheet, targetRange, "分割")) return true;
    items.push({ sheetIndex, sheet, sourceRange, targetRange, outputRows });
  }
  if (!items.length) return true;
  const overwrite = items.find((item) => rangeHasUserData(item.sheet, item.targetRange, item.sourceRange));
  if (overwrite && !control.skipOverwriteCheck) {
    showDataOverwriteConfirmation(() => {
      const result = textToColumnsForSheetGroup(range, options, destination, { ...control, skipOverwriteCheck: true, afterOverwriteConfirm: null });
      if (result === true) control.afterOverwriteConfirm?.();
    });
    return "pending";
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      textToColumns(item.sourceRange, options, destination, {
        sheet: item.sheet,
        sheetIndex: item.sheetIndex,
        targetRange: item.targetRange,
        outputRows: item.outputRows,
        skipOverwriteCheck: true,
        commitHistory: false,
      });
    });
    state.hf = buildFormulaEngine(state.model);
    const activeTargetRange = control.activeTargetRange || items.find((item) => item.sheetIndex === state.activeSheetIndex)?.targetRange || items[0].targetRange;
    state.selected = { sheetIndex: state.activeSheetIndex, row: activeTargetRange.top, col: activeTargetRange.left };
    state.selectionRange = { ...activeTargetRange, sheetIndex: state.activeSheetIndex };
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectedImage = null;
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: range,
      selectionAfter: state.selectionRange,
      selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
      selectionAfterRanges: [state.selectionRange],
      label: `${items.length} シート 区切り位置`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートで${rangeToLabel(range)}を分割しました。`);
    return true;
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`区切り位置を実行できませんでした：${error.message || error}`);
    return true;
  }
}



function textToColumnsRowsForRange(sheet, range, options) {
  const rows = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    const text = getCellRawInput(sheet, row, range.left);
    rows.push(options.mode === "fixed" ? splitTextByFixedBreaks(text, options.fixedBreaks) : splitTextByDelimiters(text, options));
  }
  return rows;
}



function splitTextByDelimiters(value, options) {
  const source = String(value ?? "");
  const delimiters = new Set(options.delimiters || []);
  if (!delimiters.size) return [source];
  const qualifier = options.qualifier || "";
  const parts = [];
  let current = "";
  let inQuote = false;
  let lastWasDelimiter = false;
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (qualifier && char === qualifier) {
      if (inQuote && source[index + 1] === qualifier) {
        current += qualifier;
        index += 1;
      } else {
        inQuote = !inQuote;
      }
      lastWasDelimiter = false;
      continue;
    }
    if (!inQuote && delimiters.has(char)) {
      if (!(options.consecutive && lastWasDelimiter)) {
        parts.push(current);
        current = "";
      }
      lastWasDelimiter = true;
      continue;
    }
    current += char;
    lastWasDelimiter = false;
  }
  parts.push(current);
  return parts;
}



function parseFixedWidthBreaks(value) {
  return [...new Set(String(value || "")
    .split(/[,\s]+/)
    .map((item) => Number.parseInt(item, 10))
    .filter((item) => Number.isFinite(item) && item > 0))]
    .sort((left, right) => left - right);
}



function splitTextByFixedBreaks(value, breaks) {
  const source = String(value ?? "");
  if (!breaks?.length) return [source];
  const parts = [];
  let start = 0;
  breaks.forEach((position) => {
    const end = Math.min(source.length, position);
    parts.push(source.slice(start, end));
    start = end;
  });
  parts.push(source.slice(start));
  return parts;
}



function toggleTextColumnsFixedBreak(dialog, position) {
  if (!dialog || !Number.isFinite(position) || position <= 0) return;
  const input = dialog.querySelector("[data-data-field='fixedBreaks']");
  if (!input) return;
  const breaks = parseFixedWidthBreaks(input.value);
  const existingIndex = breaks.indexOf(position);
  if (existingIndex >= 0) breaks.splice(existingIndex, 1);
  else breaks.push(position);
  breaks.sort((left, right) => left - right);
  input.value = breaks.join(", ");
  state.dataDialog.fixedBreaksTouched = true;
  syncTextToColumnsDialog(dialog);
}



function coerceTextColumnsValue(value, format, options = {}) {
  const text = String(value ?? "");
  if (format === "text") return text;
  if (String(format || "").startsWith("date")) {
    const date = coerceTextColumnsDateValue(text, textColumnsDateOrderFromFormat(format));
    if (date != null) return date;
  }
  const numeric = coerceTextColumnsNumberValue(text, options);
  if (numeric != null) return numeric;
  return coerceUserInput(text);
}



function coerceTextColumnsNumberValue(value, options = {}) {
  let text = String(value ?? "").trim();
  if (!text) return null;
  let sign = 1;
  if (options.trailingMinus && /-$/.test(text)) {
    sign = -1;
    text = text.slice(0, -1).trim();
  }
  const decimalSeparator = String(options.decimalSeparator || ".")[0] || ".";
  const thousandsSeparator = String(options.thousandsSeparator || "")[0] || "";
  if (thousandsSeparator) {
    text = text.replace(new RegExp(escapeRegExp(thousandsSeparator), "g"), "");
  }
  if (decimalSeparator && decimalSeparator !== ".") {
    text = text.replace(new RegExp(escapeRegExp(decimalSeparator), "g"), ".");
  }
  if (!/^[+-]?(?:\d+|\d*\.\d+)$/.test(text)) return null;
  const valueNumber = Number(text);
  return Number.isFinite(valueNumber) ? sign * valueNumber : null;
}



function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}



function coerceTextColumnsDateValue(value, order = "YMD") {
  const text = String(value ?? "").trim();
  const namedDate = autoFilterMonthNameDateSerial(text);
  if (namedDate != null) return serialDayToYmd(namedDate);
  const parts = text.match(/\d+/g);
  if (!parts || parts.length < 3) return null;
  const mapped = {};
  textColumnsDateOrderFromFormat(order)
    .split("")
    .forEach((key, index) => {
      mapped[key] = Number.parseInt(parts[index], 10);
    });
  let year = mapped.Y;
  if (!Number.isFinite(year) || !Number.isFinite(mapped.M) || !Number.isFinite(mapped.D)) return null;
  if (year < 100) year += year < 30 ? 2000 : 1900;
  const month = mapped.M;
  const day = mapped.D;
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  return `${String(year).padStart(4, "0")}/${String(month).padStart(2, "0")}/${String(day).padStart(2, "0")}`;
}



function serialDayToYmd(serialDay) {
  if (!Number.isFinite(serialDay)) return null;
  const date = new Date(serialDay * 86400000);
  return `${String(date.getUTCFullYear()).padStart(4, "0")}/${String(date.getUTCMonth() + 1).padStart(2, "0")}/${String(date.getUTCDate()).padStart(2, "0")}`;
}



function flashFillSelection() {
  const context = selectedDataRange("フラッシュ フィル");
  if (!context) return;
  if (flashFillSelectionForSheetGroup(context)) return;
  const { sheet } = context;
  const range = flashFillTargetRangeForSelection(sheet, context.range);
  if (!ensureEditableRangeForProtection(sheet, range, "フラッシュ フィル")) return;
  if (range.left <= 1) {
    showMessageBox("左側の列に元データが必要です。", { title: "Cht WebSheet" });
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  const result = applyFlashFillToSheet(sheet, state.activeSheetIndex, range);
  if (!result.changed) {
    showMessageBox(result.skipped ? "フラッシュ フィルで使用できるパターンを検出できませんでした。" : "フラッシュ フィルで入力できる空白セルがありません。", { title: "Cht WebSheet" });
    return;
  }
  const skippedText = result.skipped ? ` ${result.skipped}個のセルはパターンを検出できませんでした。` : "";
  commitDataOperation(sheet, sheetBefore, range, `${rangeToLabel(range)} フラッシュ フィル`, `${result.changed}個のセルにフラッシュ フィルを適用しました。${skippedText}`);
}



function flashFillTargetRangeForSelection(sheet, range, sheetIndex = state.activeSheetIndex) {
  if (!sheet || !range) return range;
  const target = { ...range, sheetIndex };
  if (!isSingleCellRange(target) || target.left <= 1) return target;
  const sourcePoint = { sheetIndex, row: target.top, col: target.left - 1 };
  const sourceRegion = currentRegionFromPoint(sheet, sourcePoint, sheetIndex);
  if (!sourceRegion || sourceRegion.bottom < target.top) return target;
  return {
    ...target,
    top: Math.max(target.top, sourceRegion.top),
    bottom: Math.max(target.top, sourceRegion.bottom),
    left: target.left,
    right: target.right,
  };
}



function flashFillRangeHasSourceData(sheet, range) {
  if (!sheet || !range || range.left <= 1) return false;
  const used = sheetUsedRange(sheet);
  if (!used) return false;
  const top = Math.max(used.top, range.top);
  const bottom = Math.min(used.bottom, range.bottom);
  if (top > bottom) return false;
  for (let col = Math.max(1, range.left - 5); col < range.left; col += 1) {
    if (rangeColumnHasData(sheet, col, top, bottom)) return true;
  }
  return false;
}



function applyFlashFillToSheet(sheet, sheetIndex, range) {
  const sheetId = getSheetId(sheet.name);
  const rulesByColumn = new Map();
  for (let col = range.left; col <= range.right; col += 1) {
    rulesByColumn.set(col, inferFlashFillRule(sheet, sheetId, range, col));
  }
  let changed = 0;
  let skipped = 0;
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      if (getCellRawInput(sheet, row, col) !== "") continue;
      const rule = rulesByColumn.get(col);
      if (!rule) {
        skipped += 1;
        continue;
      }
      const next = applyFlashFillRuleToRow(sheet, sheetId, row, rule);
      if (next == null || String(next) === "") {
        skipped += 1;
        continue;
      }
      writeCellContent(sheet, sheetId, row, col, next);
      changed += 1;
    }
  }
  if (changed) internCellsInRange(sheet, range);
  return { sheetIndex, changed, skipped };
}



function flashFillSelectionForSheetGroup(context) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !context?.range) return false;
  if (context.range.left <= 1) {
    showMessageBox("左側の列に元データが必要です。", { title: "Cht WebSheet" });
    return true;
  }
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    const materialized = materializeSelectionRangeForSheet(sheet, context.range, sheetIndex);
    const range = flashFillTargetRangeForSelection(sheet, materialized, sheetIndex);
    if (!range) continue;
    if (!ensureEditableRangeForProtection(sheet, range, "フラッシュ フィル")) return true;
    items.push({ sheetIndex, sheet, range });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    const results = items.map((item) => applyFlashFillToSheet(item.sheet, item.sheetIndex, item.range));
    const changed = results.reduce((total, item) => total + item.changed, 0);
    const skipped = results.reduce((total, item) => total + item.skipped, 0);
    if (!changed) {
      restoreGroupedSheetsAfterFailure(sheetsBefore);
      showMessageBox(skipped ? "フラッシュ フィルで使用できるパターンを検出できませんでした。" : "フラッシュ フィルで入力できる空白セルがありません。", { title: "Cht WebSheet" });
      return true;
    }
    state.hf = buildFormulaEngine(state.model);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: context.range,
      selectionAfter: context.range,
      selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
      selectionAfterRanges: activeSelectionRanges().map((item) => ({ ...item })),
      label: `${items.length} シート フラッシュ フィル`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    const skippedText = skipped ? ` ${skipped}個のセルはパターンを検出できませんでした。` : "";
    setStatus(`${items.length} 個のシートで${changed}個のセルにフラッシュ フィルを適用しました。${skippedText}`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`フラッシュ フィルを実行できませんでした：${error.message || error}`);
  }
  return true;
}



function inferFlashFillRule(sheet, sheetId, range, targetCol) {
  const multiColumnRule = inferFlashFillMultiColumnRule(sheet, sheetId, range, targetCol);
  if (multiColumnRule) return multiColumnRule;
  const minSourceCol = Math.max(1, targetCol - 5);
  for (let sourceCol = targetCol - 1; sourceCol >= minSourceCol; sourceCol -= 1) {
    const examples = flashFillExamplesForSourceColumn(sheet, sheetId, range, targetCol, sourceCol);
    const rule = inferFlashFillRuleFromExamples(examples);
    if (rule && rule.type !== "copy") return { ...rule, sourceCol, exampleCount: examples.length };
  }
  const copyExamples = flashFillExamplesForSourceColumn(sheet, sheetId, range, targetCol, targetCol - 1);
  const copyRule = inferFlashFillRuleFromExamples(copyExamples);
  return copyRule ? { ...copyRule, sourceCol: targetCol - 1, exampleCount: copyExamples.length } : null;
}



function flashFillExamplesForSourceColumn(sheet, sheetId, range, targetCol, sourceCol) {
  if (sourceCol < 1 || sourceCol >= targetCol) return [];
  const selectedRows = [];
  for (let row = range.top; row <= range.bottom; row += 1) selectedRows.push(row);
  const selectedExamples = flashFillExamplesFromRows(sheet, sheetId, selectedRows, targetCol, sourceCol);
  if (selectedExamples.length) return selectedExamples;
  const rows = new Set(selectedRows);
  const used = sheetUsedRange(sheet);
  if (used) {
    let scanned = 0;
    for (let row = range.top - 1; row >= used.top && scanned < 80; row -= 1, scanned += 1) rows.add(row);
  }
  const examples = flashFillExamplesFromRows(sheet, sheetId, [...rows].sort((left, right) => left - right), targetCol, sourceCol);
  if (used && examples.length > 1) return examples.filter((item) => item.row !== used.top);
  return examples;
}



function flashFillExamplesFromRows(sheet, sheetId, rows, targetCol, sourceCol) {
  return (rows || [])
    .sort((left, right) => left - right)
    .map((row) => ({
      row,
      source: String(getDisplayForCell(sheet, sheetId, row, sourceCol) ?? ""),
      target: String(getDisplayForCell(sheet, sheetId, row, targetCol) ?? ""),
    }))
    .filter((item) => item.source && item.target);
}



function inferFlashFillMultiColumnRule(sheet, sheetId, range, targetCol) {
  const minSourceCol = Math.max(1, targetCol - 5);
  const sourceCols = [];
  for (let col = minSourceCol; col < targetCol; col += 1) sourceCols.push(col);
  const used = sheetUsedRange(sheet);
  const rows = new Set();
  for (let row = range.top; row <= range.bottom; row += 1) rows.add(row);
  if (used) {
    let scanned = 0;
    for (let row = range.top - 1; row >= used.top && scanned < 80; row -= 1, scanned += 1) rows.add(row);
  }
  const examples = [...rows]
    .sort((left, right) => left - right)
    .map((row) => ({
      row,
      target: String(getDisplayForCell(sheet, sheetId, row, targetCol) ?? ""),
      sources: sourceCols.map((col) => String(getDisplayForCell(sheet, sheetId, row, col) ?? "")),
    }))
    .filter((item) => item.target && item.sources.filter(Boolean).length >= 2);
  if (!examples.length) return null;
  const separators = [" ", "", "-", "/", "_", "・", "."];
  for (let width = Math.min(5, sourceCols.length); width >= 2; width -= 1) {
    for (let start = sourceCols.length - width; start >= 0; start -= 1) {
      const cols = sourceCols.slice(start, start + width);
      for (const separator of separators) {
        if (examples.every((item) => {
          const values = cols.map((col) => item.sources[sourceCols.indexOf(col)]).filter(Boolean);
          return values.length === width && item.target === values.join(separator);
        })) {
          return { type: "multiColumnJoin", sourceCols: cols, separator, exampleCount: examples.length };
        }
      }
    }
  }
  return null;
}



function inferFlashFillRuleFromExamples(examples) {
  if (!examples.length) return null;
  const first = examples[0];
  const candidates = [
    inferFlashFillCaseRule(examples),
    inferFlashFillDelimiterRule(examples),
    inferFlashFillJoinTokensRule(examples),
    inferFlashFillSliceFormatRule(examples),
    inferFlashFillRegexSegmentRule(examples),
    inferFlashFillBetweenCharactersRule(examples),
    inferFlashFillInitialsRule(examples),
    inferFlashFillTokenReorderRule(examples),
    inferFlashFillSubstringRule(examples, "prefix"),
    inferFlashFillSubstringRule(examples, "suffix"),
    inferFlashFillTokenRule(examples, "firstToken"),
    inferFlashFillTokenRule(examples, "lastToken"),
    inferFlashFillAffixRule(examples),
    first.source === first.target ? { type: "copy" } : null,
  ].filter(Boolean);
  return candidates[0] || null;
}



function inferFlashFillCaseRule(examples) {
  if (examples.every((item) => item.target === item.source.toUpperCase())) return { type: "upper" };
  if (examples.every((item) => item.target === item.source.toLowerCase())) return { type: "lower" };
  return null;
}



function inferFlashFillSubstringRule(examples, mode) {
  const lengths = examples.map((item) => item.target.length);
  const length = lengths[0];
  if (!lengths.every((item) => item === length)) return null;
  if (mode === "prefix" && examples.every((item) => item.source.startsWith(item.target))) return { type: "prefix", length };
  if (mode === "suffix" && examples.every((item) => item.source.endsWith(item.target))) return { type: "suffix", length };
  return null;
}



function inferFlashFillTokenRule(examples, type) {
  const tokenAt = (text, fromEnd = false) => {
    const tokens = String(text).trim().split(/\s+/).filter(Boolean);
    return fromEnd ? tokens[tokens.length - 1] || "" : tokens[0] || "";
  };
  if (type === "firstToken" && examples.every((item) => item.target === tokenAt(item.source))) return { type };
  if (type === "lastToken" && examples.every((item) => item.target === tokenAt(item.source, true))) return { type };
  return null;
}



function inferFlashFillDelimiterRule(examples) {
  const delimiters = ["@", ".", "-", "_", "/", "\\", "|", ":", ";", ",", " "];
  for (const delimiter of delimiters) {
    const maxParts = Math.max(...examples.map((item) => String(item.source).split(delimiter).length));
    for (let index = 0; index < maxParts; index += 1) {
      if (examples.every((item) => {
        const parts = String(item.source).split(delimiter);
        return parts.length > index && item.target === parts[index];
      })) {
        return { type: "delimiterPart", delimiter, index };
      }
    }
    for (let fromEnd = 1; fromEnd <= maxParts; fromEnd += 1) {
      if (examples.every((item) => {
        const parts = String(item.source).split(delimiter);
        return parts.length >= fromEnd && item.target === parts[parts.length - fromEnd];
      })) {
        return { type: "delimiterPartFromEnd", delimiter, fromEnd };
      }
    }
  }
  return null;
}



function inferFlashFillJoinTokensRule(examples) {
  const separators = ["-", "/", ".", "_", " ", "・"];
  for (const separator of separators) {
    if (examples.every((item) => {
      const tokens = flashFillJoinTokens(item.source);
      return tokens.length > 1 && item.target === tokens.join(separator);
    })) {
      return { type: "joinTokens", separator };
    }
  }
  return null;
}



function flashFillJoinTokens(text) {
  return String(text ?? "").match(/[A-Za-z0-9０-９]+/g) || [];
}



function inferFlashFillSliceFormatRule(examples) {
  const pattern = deriveFlashFillSliceFormat(examples[0]?.source, examples[0]?.target);
  if (!pattern) return null;
  if (examples.every((item) => applyFlashFillSliceFormat(pattern, item.source) === item.target)) {
    return { type: "sliceFormat", pattern };
  }
  return null;
}



function deriveFlashFillSliceFormat(sourceValue, targetValue) {
  const source = String(sourceValue ?? "");
  const target = String(targetValue ?? "");
  if (!source || !target || source === target) return null;
  const parts = [];
  let sourceIndex = 0;
  let targetIndex = 0;
  while (sourceIndex < source.length) {
    const match = bestFlashFillSliceMatch(source, target, sourceIndex, targetIndex);
    if (!match) return null;
    if (match.targetIndex > targetIndex) {
      parts.push({ literal: target.slice(targetIndex, match.targetIndex) });
    }
    parts.push({ length: match.length });
    sourceIndex += match.length;
    targetIndex = match.targetIndex + match.length;
  }
  if (targetIndex < target.length) parts.push({ literal: target.slice(targetIndex) });
  if (!parts.some((part) => part.literal)) return null;
  const totalSliceLength = parts.reduce((total, part) => total + (Number(part.length) || 0), 0);
  if (totalSliceLength !== source.length) return null;
  return applyFlashFillSliceFormat(parts, source) === target ? parts : null;
}



function bestFlashFillSliceMatch(source, target, sourceIndex, targetIndex) {
  const remaining = source.length - sourceIndex;
  for (let position = targetIndex; position < target.length; position += 1) {
    let bestLength = 0;
    const maxLength = Math.min(remaining, target.length - position);
    for (let length = 1; length <= maxLength; length += 1) {
      if (target.slice(position, position + length) === source.slice(sourceIndex, sourceIndex + length)) {
        bestLength = length;
      }
    }
    if (bestLength) return { targetIndex: position, length: bestLength };
  }
  return null;
}



function applyFlashFillSliceFormat(pattern, sourceValue) {
  const source = String(sourceValue ?? "");
  let sourceIndex = 0;
  let output = "";
  for (const part of pattern || []) {
    if (part.literal != null) {
      output += part.literal;
      continue;
    }
    const length = Number(part.length) || 0;
    if (length <= 0 || sourceIndex + length > source.length) return source;
    output += source.slice(sourceIndex, sourceIndex + length);
    sourceIndex += length;
  }
  return sourceIndex === source.length ? output : source;
}



function inferFlashFillRegexSegmentRule(examples) {
  for (const segment of FLASH_FILL_REGEX_SEGMENTS) {
    const maxParts = Math.max(...examples.map((item) => flashFillRegexSegmentMatches(item.source, segment.id).length));
    for (let index = 0; index < maxParts; index += 1) {
      if (examples.every((item) => flashFillRegexSegmentMatches(item.source, segment.id)[index] === item.target)) {
        return { type: "regexSegment", pattern: segment.id, index };
      }
    }
    for (let fromEnd = 1; fromEnd <= maxParts; fromEnd += 1) {
      if (examples.every((item) => {
        const matches = flashFillRegexSegmentMatches(item.source, segment.id);
        return matches.length >= fromEnd && matches[matches.length - fromEnd] === item.target;
      })) {
        return { type: "regexSegmentFromEnd", pattern: segment.id, fromEnd };
      }
    }
  }
  return null;
}



function flashFillRegexSegmentMatches(source, patternId) {
  const pattern = FLASH_FILL_REGEX_SEGMENTS.find((item) => item.id === patternId)?.regex;
  if (!pattern) return [];
  return String(source ?? "").match(new RegExp(pattern.source, "g")) || [];
}



function inferFlashFillBetweenCharactersRule(examples) {
  const first = examples[0];
  const candidates = flashFillBoundaryCandidates(first.source, first.target);
  return candidates.find((candidate) =>
    examples.every((item) => flashFillBetweenCharacters(item.source, candidate.left, candidate.right) === item.target),
  ) || null;
}



function flashFillBoundaryCandidates(source, target) {
  const text = String(source ?? "");
  const value = String(target ?? "");
  if (!value) return [];
  const candidates = [];
  let index = text.indexOf(value);
  while (index >= 0) {
    const left = text[index - 1] || "";
    const right = text[index + value.length] || "";
    if (left && right) candidates.push({ type: "betweenChars", left, right });
    index = text.indexOf(value, index + value.length);
  }
  return candidates;
}



function flashFillBetweenCharacters(source, left, right) {
  const text = String(source ?? "");
  const leftIndex = text.indexOf(left);
  if (leftIndex < 0) return "";
  const start = leftIndex + String(left).length;
  const end = text.indexOf(right, start);
  return end < 0 ? "" : text.slice(start, end);
}



function flashFillWordTokens(text) {
  return String(text).trim().split(/[\s._@\-]+/).filter(Boolean);
}



function inferFlashFillInitialsRule(examples) {
  if (examples.every((item) => {
    const tokens = flashFillWordTokens(item.source);
    return tokens.length > 1 && item.target === tokens.map((token) => token[0] || "").join("");
  })) {
    return { type: "initials", upper: examples.every((item) => item.target === item.target.toUpperCase()) };
  }
  return null;
}



function inferFlashFillTokenReorderRule(examples) {
  const patterns = [
    { type: "lastFirstComma", build: ([first, last]) => `${last}, ${first}` },
    { type: "lastFirstSpace", build: ([first, last]) => `${last} ${first}` },
    { type: "firstLastComma", build: ([first, last]) => `${first}, ${last}` },
  ];
  return patterns.find((pattern) =>
    examples.every((item) => {
      const tokens = String(item.source).trim().split(/\s+/).filter(Boolean);
      return tokens.length === 2 && item.target === pattern.build(tokens);
    }),
  ) || null;
}



function inferFlashFillAffixRule(examples) {
  const first = examples[0];
  const sourceIndex = first.target.indexOf(first.source);
  if (sourceIndex < 0) return null;
  const prefix = first.target.slice(0, sourceIndex);
  const suffix = first.target.slice(sourceIndex + first.source.length);
  if (examples.every((item) => item.target === `${prefix}${item.source}${suffix}`)) return { type: "affix", prefix, suffix };
  return null;
}



function applyFlashFillRuleToRow(sheet, sheetId, row, rule) {
  if (!rule) return null;
  if (rule.type === "multiColumnJoin") {
    const values = (rule.sourceCols || [])
      .map((col) => String(getDisplayForCell(sheet, sheetId, row, col) ?? ""))
      .filter(Boolean);
    return values.length === (rule.sourceCols || []).length ? values.join(rule.separator || "") : null;
  }
  const sourceCol = Number(rule.sourceCol);
  if (!Number.isFinite(sourceCol) || sourceCol < 1) return null;
  const source = String(getDisplayForCell(sheet, sheetId, row, sourceCol) ?? "");
  if (!source) return null;
  return applyFlashFillRule(rule, source);
}



function applyFlashFillRule(rule, source) {
  switch (rule?.type) {
    case "upper":
      return source.toUpperCase();
    case "lower":
      return source.toLowerCase();
    case "prefix":
      return source.slice(0, rule.length);
    case "suffix":
      return source.slice(Math.max(0, source.length - rule.length));
    case "firstToken":
      return source.trim().split(/\s+/).filter(Boolean)[0] || "";
    case "lastToken": {
      const tokens = source.trim().split(/\s+/).filter(Boolean);
      return tokens[tokens.length - 1] || "";
    }
    case "delimiterPart": {
      const parts = source.split(rule.delimiter);
      return parts[rule.index] ?? "";
    }
    case "delimiterPartFromEnd": {
      const parts = source.split(rule.delimiter);
      return parts[Math.max(0, parts.length - rule.fromEnd)] ?? "";
    }
    case "joinTokens":
      return flashFillJoinTokens(source).join(rule.separator || "");
    case "sliceFormat":
      return applyFlashFillSliceFormat(rule.pattern, source);
    case "regexSegment":
      return flashFillRegexSegmentMatches(source, rule.pattern)[rule.index] ?? "";
    case "regexSegmentFromEnd": {
      const matches = flashFillRegexSegmentMatches(source, rule.pattern);
      return matches[Math.max(0, matches.length - rule.fromEnd)] ?? "";
    }
    case "betweenChars":
      return flashFillBetweenCharacters(source, rule.left, rule.right);
    case "initials": {
      const value = flashFillWordTokens(source).map((token) => token[0] || "").join("");
      return rule.upper ? value.toUpperCase() : value;
    }
    case "lastFirstComma": {
      const [first, last] = source.trim().split(/\s+/).filter(Boolean);
      return first && last ? `${last}, ${first}` : source;
    }
    case "lastFirstSpace": {
      const [first, last] = source.trim().split(/\s+/).filter(Boolean);
      return first && last ? `${last} ${first}` : source;
    }
    case "firstLastComma": {
      const [first, last] = source.trim().split(/\s+/).filter(Boolean);
      return first && last ? `${first}, ${last}` : source;
    }
    case "affix":
      return `${rule.prefix || ""}${source}${rule.suffix || ""}`;
    case "copy":
    default:
      return source;
  }
}

export {
  advancedCriterionMatches,
  advancedFilterCellCriterion,
  advancedFilterCopyTargetInfo,
  advancedFilterCriteriaHeaderIssues,
  advancedFilterCriteriaRows,
  advancedFilterDialogPreview,
  advancedFilterFormulaMatches,
  advancedFilterFormulaStringLiteral,
  advancedFilterHeaderMap,
  advancedFilterMatchingRows,
  advancedFilterNamedRange,
  advancedFilterOutputColumnDetails,
  advancedFilterOutputColumns,
  advancedFilterRowMatches,
  advancedFilterRowSignature,
  advancedWildcardToRegExp,
  applyAdvancedFilterDialog,
  applyAdvancedFilterInPlace,
  applyAverageAutoFilter,
  applyBlankAutoFilter,
  applyColorAutoFilter,
  applyColorAutoFilterSort,
  applyCustomAutoFilterDialog,
  applyDatePeriodAutoFilter,
  applyDatePresetAutoFilter,
  applyFlashFillRule,
  applyFlashFillRuleToRow,
  applyFlashFillSliceFormat,
  applyFlashFillToSheet,
  applyStdDevAutoFilter,
  applyTextToColumnsDialog,
  applyTop10AutoFilter,
  applyTop10AutoFilterDialog,
  autoFilterAddDays,
  autoFilterCellValue,
  autoFilterColorButtonHtml,
  autoFilterColorMenuHtml,
  autoFilterConditionMatches,
  autoFilterConditionMenuHtml,
  autoFilterDateForSerial,
  autoFilterDatePartsFromSerial,
  autoFilterDatePartsFromValue,
  autoFilterDateRangeForPreset,
  autoFilterDateSerial,
  autoFilterDialogPreview,
  autoFilterDynamicDatePeriodType,
  autoFilterFocusableControls,
  autoFilterLocalDate,
  autoFilterMonthNameDateSerial,
  autoFilterOperatorUsesBlankValue,
  autoFilterPresetForOperator,
  autoFilterSelectedValuesFromMenu,
  autoFilterSerialFromDateParts,
  autoFilterTimePartsFromText,
  autoFilterTop10ValueMatches,
  autoFilterUniqueValues,
  autoFilterValueGroupKey,
  autoFilterValueInDateGroups,
  autoFilterValueInSelection,
  autoFilterValueMatchesCriterion,
  autoFilterValueMatchesDateGroup,
  autoFilterValueMatchesDatePeriod,
  autoFilterVisibleValueInputs,
  autoFilterWildcardMatches,
  bestFlashFillSliceMatch,
  closeAutoFilterMenu,
  closeAutoFilterMenuOnEscape,
  closeAutoFilterMenuOnOutsidePointer,
  coerceTextColumnsDateValue,
  coerceTextColumnsNumberValue,
  coerceTextColumnsValue,
  compareAutoFilterValueItems,
  copyAdvancedFilterColumnModel,
  copyAdvancedFilterResults,
  copyAdvancedFilterRowModel,
  countRowsMatchingAutoFilterCriterion,
  customAutoFilterCriterionFromDialog,
  customAutoFilterOperatorOptions,
  dataCellFieldHtml,
  dataRangeFieldHtml,
  decodedDialogRange,
  decodedOptionalDialogRange,
  deriveFlashFillSliceFormat,
  ensureTextColumnsFixedBreaks,
  escapeRegExp,
  fixedBreakCandidatesForText,
  fixedBreakCharacterClass,
  flashFillBetweenCharacters,
  flashFillBoundaryCandidates,
  flashFillExamplesForSourceColumn,
  flashFillExamplesFromRows,
  flashFillJoinTokens,
  flashFillRangeHasSourceData,
  flashFillRegexSegmentMatches,
  flashFillSelection,
  flashFillSelectionForSheetGroup,
  flashFillTargetRangeForSelection,
  flashFillWordTokens,
  handleAdvancedFilterDialogSync,
  handleAutoFilterDialogSync,
  handleAutoFilterMenuAction,
  handleAutoFilterMenuKeydown,
  handleAutoFilterSearchInput,
  handleAutoFilterSelectAllChange,
  handleAutoFilterValueChange,
  handleTextColumnsPreviewChange,
  inferFlashFillAffixRule,
  inferFlashFillBetweenCharactersRule,
  inferFlashFillCaseRule,
  inferFlashFillDelimiterRule,
  inferFlashFillInitialsRule,
  inferFlashFillJoinTokensRule,
  inferFlashFillMultiColumnRule,
  inferFlashFillRegexSegmentRule,
  inferFlashFillRule,
  inferFlashFillRuleFromExamples,
  inferFlashFillSliceFormatRule,
  inferFlashFillSubstringRule,
  inferFlashFillTokenReorderRule,
  inferFlashFillTokenRule,
  inferTextColumnsFixedBreaks,
  initialTextColumnsDelimiterOptions,
  isAutoFilterHeaderCell,
  normalizeAdvancedFilterHeader,
  normalizeAutoFilterCriterion,
  normalizeTwoDigitYear,
  openAdvancedFilterDialog,
  openCustomAutoFilterDialog,
  openTextToColumnsDialog,
  openTop10AutoFilterDialog,
  pageSetupReferenceFromSelectionRange,
  parseFixedWidthBreaks,
  readTextColumnsOptions,
  saveTextColumnsColumnFormat,
  serialDayToYmd,
  setDialogRangeFieldFromCurrentSelection,
  sheetForWorkbookRange,
  splitTextByDelimiters,
  splitTextByFixedBreaks,
  syncAdvancedFilterDialog,
  syncAutoFilterDialog,
  syncAutoFilterMenuState,
  syncCustomAutoFilterDialogControls,
  syncTextToColumnsDialog,
  textColumnsDateOrderFromFormat,
  textColumnsDateOrderOptions,
  textColumnsDestinationFromDialog,
  textColumnsDialogValidation,
  textColumnsFixedPreviewHtml,
  textColumnsFormatLabel,
  textColumnsOutputRows,
  textColumnsPreviewHtml,
  textColumnsSourcePreviewHtml,
  textColumnsStepValidation,
  textToColumns,
  textToColumnsForSheetGroup,
  textToColumnsRowsForRange,
  textToColumnsSourceRows,
  toggleTextColumnsFixedBreak,
  top10AutoFilterCriterionForColumn,
  top10OptionsFromDialog,
  updateAutoFilterSelectAllState,
  workbookRangeFromFormulaReference,
};
