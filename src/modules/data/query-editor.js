// Feature module; shared bindings are imported explicitly.
import {
  createIcons,
  icons,
  state,
} from "../../app-core.js";
import {
  loadWebRowsIntoSelection,
  nextQueryName,
  openWebNavigatorDialog,
  queryAppendSourceRows,
  queryMergeSourceRows,
  queryReferenceSourceRows,
} from "./connections-and-queries.js";
import {
  escapeRegExp,
  parseFixedWidthBreaks,
  splitTextByFixedBreaks,
} from "./filter-and-text-tools.js";
import {
  compareQueryTransformValues,
  evaluateQueryCustomColumnFormula,
  loadQueryTransformEditorToSelection,
  moveQueryTransformArrayItem,
  odataImportRowsFromSource,
  padQueryTransformRow,
  queryTransformAggregateDataType,
  queryTransformAggregateValue,
  queryTransformCharacterSetValue,
  queryTransformConditionMatches,
  queryTransformConvertDataTypeValue,
  queryTransformDatePartDataType,
  queryTransformDatePartValue,
  queryTransformDefaultColumns,
  queryTransformFormatDate,
  queryTransformInferDataType,
  queryTransformIsErrorValue,
  queryTransformMoveTargetIndex,
  queryTransformNormalizeDataType,
  queryTransformNumberOperationValue,
  queryTransformStepColumnIndexes,
  queryTransformTextCaseValue,
  registerQueryTransformConnectionAndOpenLoadTo,
  rememberRecentDataSourceFromQuery,
  sheetDisplayRowsForRange,
  splitTextByCharacterCount,
  uniqueQueryTransformColumns,
  webImportRowsFromSource,
  webImportSourcesFromSource,
} from "./query-transforms.js";
import {
  detectDelimitedTextDelimiter,
  parseDelimitedRows,
  rangeHasUserData,
  sheetUsedRange,
  showDataOverwriteConfirmation,
} from "./sort-and-autofilter.js";
import {
  closeDataDialog,
} from "../drawing/charts-and-insertion.js";
import {
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  isElementImeCompositionEvent,
} from "../editing/keyboard-and-paste.js";
import {
  activeSelectionRange,
  coerceUserInput,
  expandRangeForMerges,
  findMergeForCell,
  getCellRawInput,
  getDisplayForCell,
  normalizeRange,
  rangeToLabel,
  updateSelectionUi,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  rangesEqual,
  rawActiveSelectionRanges,
  scrollRangeIntoView,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  pushHistory,
} from "../editing/structure-and-history.js";
import {
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  columnName,
  escapeAttr,
  escapeHtml,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  openFindDialog,
} from "../ui/printing-and-home-tools.js";
import {
  applyExcelTableStyle,
  commitDataOperation,
  ensureSheetSize,
  nextTableName,
  parseCellKey,
} from "../ui/ribbon-and-links.js";
import {
  buildFormulaEngine,
  isSheetVisible,
  setActiveSheet,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  excelJsWorkbookFromFile,
} from "../workbook/import-model.js";
import {
  activeSheet,
  getSheetId,
} from "../workbook/saving-and-export.js";


function handleFindDialogKeydown(event) {
  if (isElementImeCompositionEvent(event)) return;
  if (event.key === "Enter") {
    event.preventDefault();
    findNextFromDialog(event.currentTarget);
  }
}



function handleFindDialogFieldChange(event) {
  const field = event.target?.dataset?.findField || "";
  if (!field || field === "replacement") return;
  clearFindAllResults(event.currentTarget.closest("#findDialog"));
}



function closeFindDialogOnEscape(event) {
  if (event.key !== "Escape" || !state.findDialog) return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  closeFindDialog();
}



function closeFindDialog() {
  document.querySelector("#findDialogOverlay")?.remove();
  state.findDialog = null;
}



function normalizeFindScope(scope) {
  return scope === "workbook" ? "workbook" : "sheet";
}



function findDialogOptionsFromElement(dialog) {
  return {
    caseSensitive: Boolean(dialog?.querySelector("[data-find-field='caseSensitive']")?.checked),
    wholeCell: Boolean(dialog?.querySelector("[data-find-field='wholeCell']")?.checked),
    scope: normalizeFindScope(dialog?.querySelector("[data-find-field='scope']")?.value),
  };
}



function findNextFromDialog(dialog) {
  if (!dialog) return;
  const query = String(dialog.querySelector("[data-find-field='query']")?.value || "");
  if (!query) {
    showMessageBox("検索する文字列を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const options = findDialogOptionsFromElement(dialog);
  state.lastFind = { query, options };
  const match = findNextCell(query, options);
  if (!match) {
    showMessageBox("検索対象が見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  selectSingleRange(match.range, `${rangeToLabel(match.range)} が見つかりました。`);
  dialog.querySelector("[data-find-field='query']")?.focus({ preventScroll: true });
}



function findAllFromDialog(dialog) {
  if (!dialog) return;
  const query = String(dialog.querySelector("[data-find-field='query']")?.value || "");
  if (!query) {
    showMessageBox("検索する文字列を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const options = findDialogOptionsFromElement(dialog);
  state.lastFind = { query, options };
  const matches = findMatchingCells(query, options);
  renderFindAllResults(dialog, matches, query, options);
  if (!matches.length) {
    setStatus("検索対象が見つかりません。");
    return;
  }
  setStatus(`${matches.length} 件が見つかりました。`);
}



function repeatLastFindShortcut() {
  const last = state.lastFind;
  if (!last?.query) {
    openFindDialog();
    setStatus("検索する文字列を入力してください。");
    return;
  }
  const match = findNextCell(last.query, last.options || {});
  if (!match) {
    showMessageBox("検索対象が見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  selectSingleRange(match.range, `${rangeToLabel(match.range)} が見つかりました。`);
}



function replaceFromDialog(dialog, options = {}) {
  if (!dialog) return;
  const query = String(dialog.querySelector("[data-find-field='query']")?.value || "");
  if (!query) {
    showMessageBox("検索する文字列を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const replacement = String(dialog.querySelector("[data-find-field='replacement']")?.value || "");
  const findOptions = findDialogOptionsFromElement(dialog);
  state.lastFind = { query, options: findOptions };
  const matches = options.all ? findMatchingCells(query, findOptions) : [findCurrentOrNextCell(query, findOptions)].filter(Boolean);
  if (!matches.length) {
    showMessageBox("検索対象が見つかりません。", { title: "Cht WebSheet" });
    return;
  }

  const activeSheetBefore = state.activeSheetIndex;
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = rawActiveSelectionRanges().map((range) => ({ ...range }));
  const affectedSheetIndexes = [...new Set(matches.map((match) => match.sheetIndex).filter((sheetIndex) => state.model?.sheets?.[sheetIndex]))];
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  let changed = 0;
  const changedMatches = [];
  matches.forEach((match) => {
    const sheet = state.model?.sheets?.[match.sheetIndex];
    const sheetId = getSheetId(sheet?.name);
    if (!sheet || sheetId == null) return;
    const source = getCellRawInput(sheet, match.row, match.col);
    const nextValue = replaceCellText(source, query, replacement, { ...findOptions, replaceAllInCell: options.all });
    if (nextValue === source) return;
    writeCellContent(sheet, sheetId, match.row, match.col, nextValue);
    changed += 1;
    changedMatches.push(match);
  });
  if (!changed) {
    showMessageBox("置換できる一致が見つかりません。", { title: "Cht WebSheet" });
    return;
  }

  state.hf = buildFormulaEngine(state.model);
  const last = changedMatches[changedMatches.length - 1];
  const selectedRange = findRangeForCell(last.sheetIndex, last.row, last.col);
  const historySheetIndexes = sheetsBefore.map((snapshot) => snapshot.index);
  pushHistory({
    sheetIndex: selectedRange.sheetIndex,
    activeSheetBefore,
    activeSheetAfter: selectedRange.sheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(historySheetIndexes),
    selectionBefore,
    selectionAfter: selectedRange,
    selectionBeforeRanges,
    selectionAfterRanges: [selectedRange],
    label: `${changed} 件の置換`,
  });
  clearFindAllResults(dialog);
  selectSingleRange(selectedRange, `${changed} 件を置換しました。`);
  setStatus(`${changed} 件を置換しました。`);
  dialog.querySelector("[data-find-field='query']")?.focus({ preventScroll: true });
}



function findCurrentOrNextCell(query, options = {}) {
  const scope = normalizeFindScope(options.scope);
  const selectedSheetIndex = state.selected?.sheetIndex;
  const sheet = state.model?.sheets?.[selectedSheetIndex];
  const canUseSelectedCell = sheet
    && isSheetVisible(sheet)
    && (scope === "workbook" || selectedSheetIndex === state.activeSheetIndex)
    && cellMatchesFindQuery(sheet, state.selected.row, state.selected.col, query, options);
  if (canUseSelectedCell) {
    const range = findRangeForCell(selectedSheetIndex, state.selected.row, state.selected.col);
    return { sheetIndex: selectedSheetIndex, row: state.selected.row, col: state.selected.col, range };
  }
  const match = findNextCell(query, options);
  return match ? { sheetIndex: match.sheetIndex, row: match.row, col: match.col, range: match.range } : null;
}



function findMatchingCells(query, options = {}) {
  const matches = [];
  const seen = new Set();
  findSearchSheetIndexes(options).forEach((sheetIndex) => {
    const sheet = state.model?.sheets?.[sheetIndex];
    const used = sheetUsedRange(sheet);
    if (!sheet || !used) return;
    for (let row = used.top; row <= used.bottom; row += 1) {
      for (let col = used.left; col <= used.right; col += 1) {
        if (!cellMatchesFindQuery(sheet, row, col, query, options)) continue;
        const range = findRangeForCell(sheetIndex, row, col);
        const key = `${range.sheetIndex}:${range.top}:${range.left}:${range.bottom}:${range.right}`;
        if (seen.has(key)) continue;
        seen.add(key);
        matches.push({ sheetIndex, row, col, range });
      }
    }
  });
  return matches;
}



function renderFindAllResults(dialog, matches, query, options = {}) {
  if (!dialog) return;
  const results = matches.map((match) => ({
    sheetIndex: match.sheetIndex,
    row: match.row,
    col: match.col,
    range: match.range,
  }));
  state.findDialog = {
    ...(state.findDialog || {}),
    query,
    options,
    results,
  };
  const panel = dialog.querySelector("[data-find-results]");
  const summary = dialog.querySelector("[data-find-results-summary]");
  const list = dialog.querySelector("[data-find-results-list]");
  if (!panel || !summary || !list) return;
  dialog.classList.add("is-results-visible");
  panel.classList.remove("hidden");
  summary.textContent = results.length ? `${results.length} 件見つかりました。` : "検索対象が見つかりません。";
  if (!results.length) {
    list.innerHTML = `<div class="find-dialog-results-empty">該当するセルはありません。</div>`;
    return;
  }
  list.innerHTML = results.map((match, index) => findResultButtonHtml(match, index)).join("");
}



function clearFindAllResults(dialog) {
  if (!dialog) return;
  dialog.classList.remove("is-results-visible");
  const panel = dialog.querySelector("[data-find-results]");
  const summary = dialog.querySelector("[data-find-results-summary]");
  const list = dialog.querySelector("[data-find-results-list]");
  panel?.classList.add("hidden");
  if (summary) summary.textContent = "";
  if (list) list.innerHTML = "";
  if (state.findDialog) state.findDialog.results = [];
}



function findResultButtonHtml(match, index) {
  const sheet = state.model?.sheets?.[match.sheetIndex];
  const sheetId = sheet ? getSheetId(sheet.name) : null;
  const range = match.range || findRangeForCell(match.sheetIndex, match.row, match.col);
  const address = rangeToLabel(range);
  const display = sheet && sheetId != null ? getDisplayForCell(sheet, sheetId, match.row, match.col) : "";
  const raw = sheet ? getCellRawInput(sheet, match.row, match.col) : "";
  const value = String(display ?? "") || String(raw ?? "");
  const sheetName = sheet?.name || `Sheet${Number(match.sheetIndex) + 1}`;
  const title = `${sheetName}!${address} ${value}`;
  return `
    <button class="find-dialog-result" type="button" role="option" aria-selected="false" title="${escapeAttr(title)}" data-find-dialog-action="select-result" data-find-result-index="${index}">
      <span>${escapeHtml(sheetName)}</span>
      <span>${escapeHtml(address)}</span>
      <span>${escapeHtml(String(value ?? ""))}</span>
    </button>`;
}



function selectFindResultFromDialog(dialog, target) {
  const index = Number(target?.dataset?.findResultIndex);
  const result = state.findDialog?.results?.[index];
  if (!dialog || !result) return;
  const range = result.range || findRangeForCell(result.sheetIndex, result.row, result.col);
  selectSingleRange(range, `${rangeToLabel(range)} が見つかりました。`);
  dialog.querySelectorAll("[data-find-result-index]").forEach((button) => {
    const selected = Number(button.dataset.findResultIndex) === index;
    button.classList.toggle("is-selected", selected);
    button.setAttribute("aria-selected", String(selected));
  });
  target.focus({ preventScroll: true });
}



function replaceCellText(value, query, replacement, options = {}) {
  const text = String(value ?? "");
  const needle = options.caseSensitive ? query : query.toLocaleLowerCase("ja-JP");
  const haystack = options.caseSensitive ? text : text.toLocaleLowerCase("ja-JP");
  if (options.wholeCell) return haystack === needle ? replacement : text;
  const index = haystack.indexOf(needle);
  if (index < 0) return text;
  if (!options.replaceAllInCell) {
    return `${text.slice(0, index)}${replacement}${text.slice(index + query.length)}`;
  }
  const flags = options.caseSensitive ? "g" : "gi";
  return text.replace(new RegExp(escapeRegExp(query), flags), replacement);
}



function findNextCell(query, options = {}) {
  const cells = findCellsInScope(options);
  if (!cells.length) return null;
  const startIndex = findSearchStartIndex(cells, options);
  for (let offset = 1; offset <= cells.length; offset += 1) {
    const cellPoint = cells[(startIndex + offset + cells.length) % cells.length];
    const sheet = state.model?.sheets?.[cellPoint.sheetIndex];
    if (cellMatchesFindQuery(sheet, cellPoint.row, cellPoint.col, query, options)) {
      return {
        sheetIndex: cellPoint.sheetIndex,
        row: cellPoint.row,
        col: cellPoint.col,
        range: findRangeForCell(cellPoint.sheetIndex, cellPoint.row, cellPoint.col),
      };
    }
  }
  return null;
}



function findSearchStartIndex(cells, options = {}) {
  const scope = normalizeFindScope(options.scope);
  const selected = state.selected;
  const selectedSheetIndex = selected && (scope === "workbook" || selected.sheetIndex === state.activeSheetIndex)
    ? selected.sheetIndex
    : state.activeSheetIndex;
  const selectedRow = selected?.sheetIndex === selectedSheetIndex ? selected.row : 0;
  const selectedCol = selected?.sheetIndex === selectedSheetIndex ? selected.col : 0;
  const currentIndex = cells.findIndex((cell) => (
    cell.sheetIndex === selectedSheetIndex
    && cell.row === selectedRow
    && cell.col === selectedCol
  ));
  if (currentIndex >= 0) return currentIndex;
  const nextIndex = cells.findIndex((cell) => (
    cell.sheetIndex > selectedSheetIndex
    || (
      cell.sheetIndex === selectedSheetIndex
      && (cell.row > selectedRow || (cell.row === selectedRow && cell.col > selectedCol))
    )
  ));
  return nextIndex >= 0 ? nextIndex - 1 : cells.length - 1;
}



function findSearchSheetIndexes(options = {}) {
  if (!state.model?.sheets?.length) return [];
  if (normalizeFindScope(options.scope) !== "workbook") {
    const sheet = state.model.sheets[state.activeSheetIndex];
    return sheet && isSheetVisible(sheet) ? [state.activeSheetIndex] : [];
  }
  return state.model.sheets
    .map((sheet, index) => (isSheetVisible(sheet) ? index : -1))
    .filter((index) => index >= 0);
}



function findCellsInScope(options = {}) {
  const cells = [];
  findSearchSheetIndexes(options).forEach((sheetIndex) => {
    const sheet = state.model?.sheets?.[sheetIndex];
    const used = sheetUsedRange(sheet);
    if (!sheet || !used) return;
    for (let row = used.top; row <= used.bottom; row += 1) {
      for (let col = used.left; col <= used.right; col += 1) {
        cells.push({ sheetIndex, row, col });
      }
    }
  });
  return cells;
}



function findRangeForCell(sheetIndex, row, col) {
  const targetSheetIndex = Number(sheetIndex);
  const sheet = state.model?.sheets?.[targetSheetIndex];
  const base = normalizeRange({ sheetIndex: targetSheetIndex, row, col }, { sheetIndex: targetSheetIndex, row, col });
  return { ...(expandRangeForMerges(sheet, base) || base), sheetIndex: targetSheetIndex };
}



function cellMatchesFindQuery(sheet, row, col, query, options = {}) {
  if (!sheet) return false;
  const sheetId = getSheetId(sheet.name);
  const raw = getCellRawInput(sheet, row, col);
  const display = getDisplayForCell(sheet, sheetId, row, col);
  const values = raw === display ? [raw] : [raw, display];
  const needle = options.caseSensitive ? String(query) : String(query).toLocaleLowerCase("ja-JP");
  return values.some((value) => {
    const haystack = options.caseSensitive ? String(value ?? "") : String(value ?? "").toLocaleLowerCase("ja-JP");
    return options.wholeCell ? haystack === needle : haystack.includes(needle);
  });
}



function selectCellsByPredicate(label, predicate) {
  const sheet = activeSheet();
  const ranges = matchingCellRanges(sheet, predicate);
  if (!ranges.length) {
    showMessageBox(`該当する${label}のセルが見つかりません。`, { title: "Cht WebSheet" });
    return;
  }
  selectRangesForHomeCommand(ranges, `${label}のセルを ${ranges.length} 個選択しました。`);
}



function matchingCellRanges(sheet, predicate) {
  const ranges = [];
  const seen = new Set();
  Object.entries(sheet?.cells || {}).forEach(([key, cell]) => {
    if (!predicate(cell)) return;
    const [row, col] = parseCellKey(key);
    const merge = findMergeForCell(sheet, row, col);
    const range = merge
      ? { sheetIndex: state.activeSheetIndex, ...merge }
      : normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col });
    const rangeKey = rangeToLabel(range);
    if (seen.has(rangeKey)) return;
    seen.add(rangeKey);
    ranges.push(range);
  });
  ranges.sort((left, right) => left.top - right.top || left.left - right.left);
  return ranges;
}



function selectSingleRange(range, statusMessage) {
  selectRangesForHomeCommand([range], statusMessage);
}



function selectRangesForHomeCommand(ranges, statusMessage) {
  const firstRange = ranges[0];
  if (!firstRange) return;
  const targetSheetIndex = Number.isInteger(Number(firstRange.sheetIndex)) ? Number(firstRange.sheetIndex) : state.activeSheetIndex;
  const targetSheet = state.model?.sheets?.[targetSheetIndex];
  if (!targetSheet || !isSheetVisible(targetSheet)) return;
  if (targetSheetIndex !== state.activeSheetIndex) setActiveSheet(targetSheetIndex);
  const targetRanges = ranges
    .map((range) => ({ ...range, sheetIndex: Number.isInteger(Number(range.sheetIndex)) ? Number(range.sheetIndex) : targetSheetIndex }))
    .filter((range) => range.sheetIndex === targetSheetIndex);
  const selectedRange = targetRanges[0] || { ...firstRange, sheetIndex: targetSheetIndex };
  state.selected = { sheetIndex: targetSheetIndex, row: selectedRange.top, col: selectedRange.left };
  state.selectionRange = { ...selectedRange };
  state.selectionRanges = targetRanges.length ? targetRanges.map((range) => ({ ...range })) : [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  scrollRangeIntoView(selectedRange);
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(statusMessage);
}



async function chooseDelimitedTextForImport() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".csv,.tsv,.txt,text/csv,text/tab-separated-values,text/plain";
  input.style.display = "none";
  document.body.appendChild(input);
  input.addEventListener("change", async () => {
    const [file] = input.files || [];
    input.remove();
    if (!file) return;
    try {
      const sourceBytes = new Uint8Array(await file.arrayBuffer());
      const origin = detectTextCsvOriginFromBytes(sourceBytes);
      const text = decodeTextCsvBytes(sourceBytes, origin);
      openTextCsvImportDialog(text, file.name || "CSV", { sourceBytes, origin });
    } catch (error) {
      console.error(error);
      showMessageBox("ファイルを読み込めませんでした。", { title: "Cht WebSheet" });
    }
  }, { once: true });
  input.click();
}



async function chooseStructuredDataForImport(kind = "json") {
  const normalizedKind = kind === "xml" ? "xml" : "json";
  const input = document.createElement("input");
  input.type = "file";
  input.accept = normalizedKind === "xml" ? ".xml,application/xml,text/xml" : ".json,application/json,text/json";
  input.style.display = "none";
  document.body.appendChild(input);
  input.addEventListener("change", async () => {
    const [file] = input.files || [];
    input.remove();
    if (!file) return;
    try {
      const sourceText = stripTextBom(await file.text());
      openStructuredImportFromText(sourceText, file.name || (normalizedKind === "xml" ? "XML" : "JSON"), normalizedKind);
    } catch (error) {
      console.error(error);
      showMessageBox("ファイルを読み込めませんでした。", { title: "Cht WebSheet" });
    }
  }, { once: true });
  input.click();
}



async function chooseWorkbookDataForImport() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".xlsx,.xlsm,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12,application/vnd.ms-excel";
  input.style.display = "none";
  document.body.appendChild(input);
  input.addEventListener("change", async () => {
    const [file] = input.files || [];
    input.remove();
    if (!file) return;
    try {
      await openWorkbookImportFromFile(file);
    } catch (error) {
      console.error(error);
      showMessageBox("ブックを読み込めませんでした。", { title: "Cht WebSheet" });
    }
  }, { once: true });
  input.click();
}



async function chooseFolderDataForImport() {
  const input = document.createElement("input");
  input.type = "file";
  input.multiple = true;
  if (!navigator.webdriver) {
    input.webkitdirectory = true;
    input.directory = true;
  }
  input.accept = ".csv,.tsv,.txt,.json,.xml,.xlsx,.xlsm,.xls,text/csv,text/tab-separated-values,text/plain,application/json,application/xml,text/xml,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12,application/vnd.ms-excel";
  input.style.display = "none";
  document.body.appendChild(input);
  input.addEventListener("change", async () => {
    const files = [...(input.files || [])];
    input.remove();
    if (!files.length) return;
    try {
      await openFolderImportFromFiles(files);
    } catch (error) {
      console.error(error);
      showMessageBox("フォルダーのデータを読み込めませんでした。", { title: "Cht WebSheet" });
    }
  }, { once: true });
  input.click();
}



async function openWorkbookImportFromFile(file) {
  if (!state.model) {
    setStatus("Excel ファイルを開いてからデータを取得してください。");
    return;
  }
  const workbook = await excelJsWorkbookFromFile(file);
  const sources = workbook.worksheets
    .map((worksheet, index) => ({
      kind: "workbook",
      itemIndex: index,
      name: worksheet.name || `Sheet${index + 1}`,
      rows: workbookImportRowsFromWorksheet(worksheet),
    }))
    .filter((source) => source.rows.some((row) => row.some((cell) => String(cell ?? "") !== "")))
    .map((source, index) => ({ ...source, itemIndex: index }));
  if (!sources.length) {
    showMessageBox("読み込めるワークシート データが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const targetPoint = state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  const source = {
    sourceType: "workbook",
    sourceName: file.name || "Excel ブック",
    sourceUrl: file.name || "Excel ブック",
    sourceText: "",
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    sources,
    rows: sources[0]?.rows || [],
    sourceItemIndex: sources[0]?.itemIndex ?? 0,
    sourceItemName: sources[0]?.name || "",
    sourceItemKind: "workbook",
    targetPoint: { ...targetPoint },
  };
  if (sources.length > 1) {
    openWebNavigatorDialog(source);
    return;
  }
  loadWebRowsIntoSelection(source.rows, {
    sourceType: source.sourceType,
    sourceName: source.sourceName,
    sourceUrl: source.sourceUrl,
    sourceText: source.sourceText,
    contentType: source.contentType,
    sourceRows: source.rows,
    sourceItemIndex: source.sourceItemIndex,
    sourceItemName: source.sourceItemName,
    sourceItemKind: source.sourceItemKind,
    targetPoint: source.targetPoint,
  });
}



function workbookImportRowsFromWorksheet(worksheet) {
  let maxRow = 0;
  let maxCol = 0;
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (workbookImportCellText(cell) === "") return;
      maxRow = Math.max(maxRow, rowNumber);
      maxCol = Math.max(maxCol, colNumber);
    });
  });
  if (!maxRow || !maxCol) return [];
  return Array.from({ length: maxRow }, (_row, rowIndex) =>
    Array.from({ length: maxCol }, (_col, colIndex) => workbookImportCellText(worksheet.getCell(rowIndex + 1, colIndex + 1))),
  );
}



function workbookImportCellText(cell) {
  const value = cell?.value;
  if (value == null) return "";
  if (value instanceof Date) return queryTransformFormatDate(value);
  if (typeof value !== "object") return String(value);
  if (Array.isArray(value.richText)) return value.richText.map((run) => run?.text || "").join("");
  if (value.result != null) return value.result instanceof Date ? queryTransformFormatDate(value.result) : String(value.result);
  if (value.text != null) return String(value.text);
  if (value.hyperlink && value.text != null) return String(value.text);
  if (value.formula) return value.result != null ? String(value.result) : `=${value.formula}`;
  return String(cell.text || "");
}



async function openFolderImportFromFiles(files) {
  if (!state.model) {
    setStatus("Excel ファイルを開いてからデータを取得してください。");
    return;
  }
  const fileListSource = folderImportFileListSource(files);
  const sources = fileListSource ? [fileListSource] : [];
  for (const file of files || []) {
    const fileSources = await folderImportSourcesFromFile(file);
    fileSources.forEach((source) => {
      if (source.rows.some((row) => row.some((cell) => String(cell ?? "") !== ""))) sources.push(source);
    });
  }
  const combinedSource = folderImportCombinedFilesSource(sources);
  if (combinedSource) sources.push(combinedSource);
  if (!sources.length) {
    showMessageBox("読み込めるフォルダー データが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const normalizedSources = sources.map((source, index) => ({ ...source, itemIndex: index }));
  const targetPoint = state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  const source = {
    sourceType: "folder",
    sourceName: "フォルダー",
    sourceUrl: "フォルダー",
    sourceText: "",
    contentType: "application/vnd.websheet.folder",
    sources: normalizedSources,
    rows: normalizedSources[0]?.rows || [],
    sourceItemIndex: normalizedSources[0]?.itemIndex ?? 0,
    sourceItemName: normalizedSources[0]?.name || "",
    sourceItemKind: normalizedSources[0]?.kind || "folder",
    targetPoint: { ...targetPoint },
  };
  if (normalizedSources.length > 1) {
    openWebNavigatorDialog(source);
    return;
  }
  loadWebRowsIntoSelection(source.rows, {
    sourceType: source.sourceType,
    sourceName: source.sourceName,
    sourceUrl: source.sourceUrl,
    sourceText: source.sourceText,
    contentType: source.contentType,
    sourceRows: source.rows,
    sourceItemIndex: source.sourceItemIndex,
    sourceItemName: source.sourceItemName,
    sourceItemKind: source.sourceItemKind,
    targetPoint: source.targetPoint,
  });
}



function folderImportFileListSource(files) {
  const rows = [...(files || [])].map((file) => {
    const fileName = folderImportDisplayName(file);
    const extension = folderImportFileExtension(fileName);
    const slashIndex = fileName.lastIndexOf("/");
    const folderPath = slashIndex >= 0 ? fileName.slice(0, slashIndex + 1) : "";
    const baseName = slashIndex >= 0 ? fileName.slice(slashIndex + 1) : fileName;
    const modified = file.lastModified ? queryTransformFormatDate(new Date(file.lastModified)) : "";
    return ["Binary", baseName, extension ? `.${extension}` : "", "", modified, modified, "[Record]", folderPath];
  });
  if (!rows.length) return null;
  return {
    kind: "folderTable",
    itemIndex: 0,
    name: "フォルダー ファイル一覧",
    rows: [["Content", "Name", "Extension", "Date accessed", "Date modified", "Date created", "Attributes", "Folder Path"], ...rows],
    sourceFileName: "フォルダー",
    sourceInnerKind: "folderTable",
  };
}



function folderImportCombinedFilesSource(sources = []) {
  const candidates = (sources || [])
    .filter((source) => source?.kind !== "folderTable" && source?.rows?.length > 1)
    .map((source) => ({
      source,
      header: (source.rows[0] || []).map((value) => String(value ?? "")),
    }))
    .filter((item) => item.header.some((value) => value.trim() !== ""));
  if (candidates.length < 2) return null;
  const groups = new Map();
  candidates.forEach((item) => {
    const key = item.header.map((value) => value.trim().toLocaleLowerCase("ja-JP")).join("\u0001");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item);
  });
  const bestGroup = [...groups.values()]
    .filter((group) => group.length >= 2)
    .sort((left, right) => right.length - left.length || folderImportCombinedRowCount(right) - folderImportCombinedRowCount(left))[0];
  if (!bestGroup) return null;
  const header = ["Source.Name", ...bestGroup[0].header];
  const rows = [header];
  bestGroup.forEach(({ source }) => {
    const sourceName = source.sourceFileName || source.name || "ファイル";
    source.rows.slice(1).forEach((row) => {
      if (!row.some((value) => String(value ?? "").trim() !== "")) return;
      rows.push([sourceName, ...padQueryTransformRow(row, bestGroup[0].header.length)]);
    });
  });
  if (rows.length <= 1) return null;
  return {
    kind: "folderCombined",
    itemIndex: 0,
    name: "結合されたファイル",
    rows,
    sourceFileName: "フォルダー",
    sourceInnerKind: "combined",
  };
}



function folderImportCombinedRowCount(group = []) {
  return group.reduce((total, item) => total + Math.max(0, (item.source?.rows?.length || 0) - 1), 0);
}



async function folderImportSourcesFromFile(file) {
  const fileName = folderImportDisplayName(file);
  const extension = folderImportFileExtension(fileName);
  if (["xlsx", "xlsm", "xls"].includes(extension)) return folderImportWorkbookSources(file, fileName);
  if (extension === "json") return folderImportStructuredSources(file, fileName, "json");
  if (extension === "xml") return folderImportStructuredSources(file, fileName, "xml");
  if (["csv", "tsv", "txt"].includes(extension) || String(file.type || "").startsWith("text/")) {
    return folderImportDelimitedTextSources(file, fileName);
  }
  return [];
}



function folderImportDisplayName(file) {
  return String(file?.webkitRelativePath || file?.name || "ファイル").replace(/\\/g, "/");
}



function folderImportFileExtension(fileName) {
  const match = String(fileName || "").toLocaleLowerCase("ja-JP").match(/\.([^.\/]+)$/);
  return match ? match[1] : "";
}



async function folderImportDelimitedTextSources(file, fileName) {
  const sourceBytes = new Uint8Array(await file.arrayBuffer());
  const origin = detectTextCsvOriginFromBytes(sourceBytes);
  const text = decodeTextCsvBytes(sourceBytes, origin);
  const delimiter = detectDelimitedTextDelimiter(text, fileName);
  const rows = parseDelimitedRows(text, delimiter).filter((row, index, list) => index < list.length - 1 || row.some((value) => value !== ""));
  return [{
    kind: "folder",
    itemIndex: 0,
    name: fileName,
    rows,
    sourceFileName: fileName,
    sourceInnerKind: "textCsv",
    delimiter,
    origin,
  }];
}



async function folderImportStructuredSources(file, fileName, kind) {
  const sourceText = stripTextBom(await file.text());
  const contentType = kind === "xml" ? "application/xml" : "application/json";
  return webImportSourcesFromSource(sourceText, contentType, fileName)
    .map((source, index, list) => ({
      ...source,
      kind,
      itemIndex: index,
      name: list.length > 1 ? `${fileName} - ${source.name || `${kind.toUpperCase()} ${index + 1}`}` : fileName,
      sourceFileName: fileName,
      sourceInnerKind: kind,
    }));
}



async function folderImportWorkbookSources(file, fileName) {
  const workbook = await excelJsWorkbookFromFile(file);
  return workbook.worksheets
    .map((worksheet, index) => ({
      kind: "workbook",
      itemIndex: index,
      name: `${fileName} - ${worksheet.name || `Sheet${index + 1}`}`,
      rows: workbookImportRowsFromWorksheet(worksheet),
      sourceFileName: fileName,
      sourceInnerKind: "workbook",
      sourceSheetName: worksheet.name || `Sheet${index + 1}`,
    }))
    .filter((source) => source.rows.some((row) => row.some((cell) => String(cell ?? "") !== "")));
}



function openStructuredImportFromText(sourceText, sourceName = "JSON", kind = "json") {
  if (!state.model) {
    setStatus("Excel ファイルを開いてからデータを取得してください。");
    return;
  }
  const contentType = kind === "xml" ? "application/xml" : "application/json";
  const targetPoint = state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  const sources = webImportSourcesFromSource(sourceText, contentType, sourceName);
  if (!sources.length) {
    showMessageBox("読み込める表データが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const source = {
    sourceType: kind,
    sourceUrl: sourceName,
    sourceText,
    contentType,
    sources,
    rows: sources[0]?.rows || [],
    sourceItemIndex: sources[0]?.itemIndex ?? 0,
    sourceItemName: sources[0]?.name || "",
    sourceItemKind: sources[0]?.kind || kind,
    targetPoint: { ...targetPoint },
  };
  if (sources.length > 1) {
    openWebNavigatorDialog(source);
    return;
  }
  loadWebRowsIntoSelection(source.rows, {
    sourceType: kind,
    sourceUrl: source.sourceUrl,
    sourceText: source.sourceText,
    contentType: source.contentType,
    sourceItemIndex: source.sourceItemIndex,
    sourceItemName: source.sourceItemName,
    sourceItemKind: source.sourceItemKind,
    targetPoint: source.targetPoint,
  });
}



function importDelimitedText(text, sourceName = "CSV") {
  importDelimitedTextIntoSelection(text, sourceName);
}



function importDelimitedTextIntoSelection(text, sourceName = "CSV", control = {}) {
  if (!state.model) return;
  const point = control.targetPoint || state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  const sheet = activeSheet();
  const delimiter = control.delimiter || detectDelimitedTextDelimiter(text, sourceName);
  const rows = parseDelimitedRows(text, delimiter).filter((row, index, list) => index < list.length - 1 || row.some((value) => value !== ""));
  if (!rows.length) {
    showMessageBox("インポートするデータがありません。", { title: "Cht WebSheet" });
    return;
  }
  const width = Math.max(...rows.map((row) => row.length), 1);
  const targetRange = normalizeRange(point, { sheetIndex: state.activeSheetIndex, row: point.row + rows.length - 1, col: point.col + width - 1 });
  if (!control.skipOverwriteCheck && rangeHasUserData(sheet, targetRange)) {
    showDataOverwriteConfirmation(() => {
      importDelimitedTextIntoSelection(text, sourceName, { ...control, skipOverwriteCheck: true });
    });
    return false;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
  const sheetId = getSheetId(sheet.name);
  rows.forEach((rowValues, rowIndex) => {
    for (let colIndex = 0; colIndex < width; colIndex += 1) {
      writeCellContent(sheet, sheetId, point.row + rowIndex, point.col + colIndex, coerceDelimitedImportValue(rowValues[colIndex] ?? "", control.typeDetection));
    }
  });
  if (control.registerQuery) {
    applyExcelTableStyle(sheet, targetRange);
    sheet.tables = Array.isArray(sheet.tables) ? sheet.tables : [];
    const existingTable = sheet.tables.find((table) => table.range && rangesEqual(table.range, targetRange));
    if (existingTable) {
      existingTable.showHeaderRow = control.hasHeader !== false;
      existingTable.showFilterButton = true;
    } else {
      sheet.tables.push({
        id: uniqueWorksheetObjectId("table"),
        name: nextTableName(sheet),
        range: { ...targetRange },
        showHeaderRow: control.hasHeader !== false,
        showFilterButton: true,
        style: "TableStyleMedium2",
      });
    }
    sheet.queries = Array.isArray(sheet.queries) ? sheet.queries : [];
    const query = {
      id: uniqueWorksheetObjectId("query"),
      name: nextQueryName(sheet),
      sourceType: "textCsv",
      sourceName,
      sourceText: text,
      origin: control.origin || "65001",
      delimiter,
      range: { ...targetRange },
      hasHeader: control.hasHeader !== false,
      typeDetection: control.typeDetection || "first200",
      loadedAt: new Date().toISOString(),
      refreshWithAll: true,
    };
    sheet.queries.push(query);
    rememberRecentDataSourceFromQuery(query);
  }
  commitDataOperation(sheet, sheetBefore, targetRange, `${sourceName} インポート`, `${rangeToLabel(targetRange)} にデータをインポートしました。`, targetRange);
  control.onSuccess?.(targetRange);
  return true;
}



function coerceDelimitedImportValue(value, typeDetection = "first200") {
  if (typeDetection === "none") return String(value ?? "");
  return coerceUserInput(value);
}



function textCsvBytesFromSource(source) {
  if (!source) return null;
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  if (ArrayBuffer.isView(source)) return new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  return null;
}



function stripTextBom(text) {
  return String(text ?? "").replace(/^\uFEFF/, "");
}



function detectTextCsvOriginFromBytes(sourceBytes) {
  const bytes = textCsvBytesFromSource(sourceBytes);
  if (!bytes || bytes.length < 2) return "65001";
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return "1200";
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return "65001";
  return "65001";
}



function textCsvEncodingLabel(origin) {
  switch (String(origin || "65001")) {
    case "932":
      return "shift_jis";
    case "1200":
      return "utf-16le";
    case "65001":
    default:
      return "utf-8";
  }
}



function decodeTextCsvBytes(sourceBytes, origin = "65001") {
  const bytes = textCsvBytesFromSource(sourceBytes);
  if (!bytes) return "";
  try {
    return stripTextBom(new TextDecoder(textCsvEncodingLabel(origin), { fatal: false }).decode(bytes));
  } catch (error) {
    console.warn("Failed to decode text/csv source", error);
    return stripTextBom(new TextDecoder("utf-8", { fatal: false }).decode(bytes));
  }
}



function textCsvDialogOrigin(dialog) {
  return dialog?.querySelector("[data-data-field='origin']")?.value || state.dataDialog?.origin || "65001";
}



function textCsvDialogText(dialog) {
  const bytes = textCsvBytesFromSource(state.dataDialog?.sourceBytes);
  if (bytes) return decodeTextCsvBytes(bytes, textCsvDialogOrigin(dialog));
  return stripTextBom(state.dataDialog?.sourceText || "");
}



function textCsvCurrentDetectedDelimiter(dialog) {
  return detectDelimitedTextDelimiter(textCsvDialogText(dialog), state.dataDialog?.sourceName || "CSV");
}



function openTextCsvImportDialog(text, sourceName = "CSV", options = {}) {
  if (!state.model) return;
  const targetPoint = state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  const origin = options.origin || "65001";
  const sourceBytes = textCsvBytesFromSource(options.sourceBytes);
  const sourceText = sourceBytes ? decodeTextCsvBytes(sourceBytes, origin) : stripTextBom(text);
  const detectedDelimiter = detectDelimitedTextDelimiter(sourceText, sourceName);
  state.dataDialog = {
    type: "text-csv-import",
    sourceText,
    sourceBytes,
    sourceName,
    origin,
    detectedDelimiter,
    targetPoint: { ...targetPoint },
  };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog text-csv-import-dialog" role="dialog" aria-modal="true" aria-labelledby="textCsvImportTitle" data-data-dialog="text-csv-import">
      <div class="excel-mini-titlebar" id="textCsvImportTitle">${escapeHtml(sourceName)}</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="text-csv-file-summary">
          <strong>${escapeHtml(sourceName)}</strong>
          <span data-text-csv-row-count>${escapeHtml(`${textCsvRowCount(sourceText, detectedDelimiter)} 行を検出しました`)}</span>
        </div>
        <div class="text-csv-options">
          <label class="excel-dialog-field"><span class="excel-dialog-caption">元のファイル</span>
            <select data-data-field="origin" data-text-csv-sync>
              <option value="65001"${origin === "65001" ? " selected" : ""}>65001: Unicode (UTF-8)</option>
              <option value="932"${origin === "932" ? " selected" : ""}>932: Japanese (Shift JIS)</option>
              <option value="1200"${origin === "1200" ? " selected" : ""}>1200: Unicode</option>
            </select>
          </label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">区切り記号</span>
            <select data-data-field="delimiter" data-text-csv-sync>${textCsvDelimiterOptions(detectedDelimiter)}</select>
          </label>
          <label class="excel-dialog-field text-csv-custom-delimiter hidden" data-text-csv-custom-row><span class="excel-dialog-caption">カスタム</span>
            <input type="text" maxlength="1" data-data-field="customDelimiter" data-text-csv-sync value="" />
          </label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">データ型検出</span>
            <select data-data-field="typeDetection">
              <option value="first200" selected>最初の 200 行に基づく</option>
              <option value="entire">データ セット全体に基づく</option>
              <option value="none">データ型を検出しない</option>
            </select>
          </label>
        </div>
        <div class="data-dialog-note">読み込み先: ${escapeHtml(`${activeSheet()?.name || ""}!${columnName(targetPoint.col)}${targetPoint.row}`)}</div>
        <div class="query-preview text-csv-preview" aria-label="プレビュー">${textCsvPreviewHtml(sourceText, detectedDelimiter)}</div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="transform">データの変換</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="load">読み込み</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
}



function textCsvDelimiterOptions(detectedDelimiter) {
  const options = [
    ["auto", `自動検出 (${delimiterLabel(detectedDelimiter)})`],
    [",", "コンマ"],
    ["\t", "タブ"],
    [";", "セミコロン"],
    [" ", "スペース"],
    ["custom", "カスタム"],
  ];
  return options.map(([value, label]) => `<option value="${escapeAttr(value)}"${value === "auto" ? " selected" : ""}>${escapeHtml(label)}</option>`).join("");
}



function delimiterLabel(delimiter) {
  switch (delimiter) {
    case "\t":
      return "タブ";
    case ";":
      return "セミコロン";
    case " ":
      return "スペース";
    case ",":
    default:
      return "コンマ";
  }
}



function textCsvRowCount(text, delimiter) {
  return parseDelimitedRows(text, delimiter).filter((row, index, list) => index < list.length - 1 || row.some((value) => value !== "")).length;
}



function updateTextCsvDialogSummary(dialog, text, detectedDelimiter) {
  const rowCount = dialog?.querySelector("[data-text-csv-row-count]");
  if (rowCount) rowCount.textContent = `${textCsvRowCount(text, detectedDelimiter)} 行を検出しました`;
  const autoOption = dialog?.querySelector("[data-data-field='delimiter'] option[value='auto']");
  if (autoOption) autoOption.textContent = `自動検出 (${delimiterLabel(detectedDelimiter)})`;
  if (state.dataDialog?.type === "text-csv-import") state.dataDialog.detectedDelimiter = detectedDelimiter;
}



function textCsvDelimiterFromDialog(dialog) {
  const value = dialog.querySelector("[data-data-field='delimiter']")?.value || "auto";
  if (value === "auto") return textCsvCurrentDetectedDelimiter(dialog) || ",";
  if (value === "custom") return dialog.querySelector("[data-data-field='customDelimiter']")?.value || ",";
  return value;
}



function textCsvRowsFromDialog(dialog) {
  const text = textCsvDialogText(dialog);
  const delimiter = textCsvDelimiterFromDialog(dialog);
  return parseDelimitedRows(text, delimiter).filter((row, index, list) => index < list.length - 1 || row.some((value) => value !== ""));
}



function textCsvPreviewHtml(text, delimiter) {
  const rows = parseDelimitedRows(text, delimiter)
    .filter((row, index, list) => index < list.length - 1 || row.some((value) => value !== ""))
    .slice(0, 50);
  if (!rows.length) return "<div class=\"query-preview-empty\">プレビューできるデータがありません。</div>";
  const width = Math.min(Math.max(...rows.map((row) => row.length), 1), 12);
  const body = rows
    .map((row) => `<tr>${Array.from({ length: width }, (_item, index) => `<td>${escapeHtml(row[index] ?? "")}</td>`).join("")}</tr>`)
    .join("");
  return `<table><tbody>${body}</tbody></table>`;
}



function handleTextCsvPreviewChange(event) {
  const dialog = event.target.closest("[data-data-dialog='text-csv-import']");
  if (!dialog || state.dataDialog?.type !== "text-csv-import") return;
  const customRow = dialog.querySelector("[data-text-csv-custom-row]");
  const delimiterValue = dialog.querySelector("[data-data-field='delimiter']")?.value || "auto";
  customRow?.classList.toggle("hidden", delimiterValue !== "custom");
  const text = textCsvDialogText(dialog);
  const detectedDelimiter = textCsvCurrentDetectedDelimiter(dialog);
  state.dataDialog.sourceText = text;
  state.dataDialog.origin = textCsvDialogOrigin(dialog);
  updateTextCsvDialogSummary(dialog, text, detectedDelimiter);
  const preview = dialog.querySelector(".text-csv-preview");
  if (preview) preview.innerHTML = textCsvPreviewHtml(text, textCsvDelimiterFromDialog(dialog));
}



function applyTextCsvImportDialog(dialog, options = {}) {
  const sourceText = textCsvDialogText(dialog);
  const sourceName = state.dataDialog?.sourceName || "CSV";
  const targetPoint = state.dataDialog?.targetPoint || state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  const delimiter = textCsvDelimiterFromDialog(dialog);
  const typeDetection = dialog.querySelector("[data-data-field='typeDetection']")?.value || "first200";
  const origin = textCsvDialogOrigin(dialog);
  const rows = textCsvRowsFromDialog(dialog);
  if (!rows.length) {
    showMessageBox("インポートするデータがありません。", { title: "Cht WebSheet" });
    return;
  }
  state.dataDialog.sourceText = sourceText;
  state.dataDialog.origin = origin;
  const hasHeader = rows.length > 1 && rows[0].some((value) => String(value || "").trim());
  const completed = importDelimitedTextIntoSelection(sourceText, sourceName, {
    delimiter,
    targetPoint,
    registerQuery: options.registerQuery !== false,
    hasHeader,
    typeDetection,
    origin,
    onSuccess: () => closeDataDialog(),
  });
  if (completed) closeDataDialog();
}



function openTextCsvTransformEditor(dialog) {
  const sourceText = textCsvDialogText(dialog);
  const sourceName = state.dataDialog?.sourceName || "CSV";
  const targetPoint = state.dataDialog?.targetPoint || state.selected || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  state.dataDialog = {
    type: "query-transform-editor",
    sourceType: "textCsv",
    sourceName,
    sourceText,
    origin: textCsvDialogOrigin(dialog),
    delimiter: textCsvDelimiterFromDialog(dialog),
    typeDetection: dialog.querySelector("[data-data-field='typeDetection']")?.value || "first200",
    targetPoint: { ...targetPoint },
    transformSteps: [],
  };
  renderQueryTransformEditorDialog();
}



function renderQueryTransformEditorDialog(focusSelector = "") {
  const descriptor = state.dataDialog;
  if (!descriptor || descriptor.type !== "query-transform-editor") return;
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay power-query-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = queryTransformEditorHtml(descriptor);
  document.body.appendChild(overlay);
  if (window.lucide) window.lucide.createIcons({ icons });
  const focusTarget = focusSelector ? overlay.querySelector(focusSelector) : overlay.querySelector("[data-data-dialog-action='query-transform-load']");
  focusTarget?.focus({ preventScroll: true });
}



function queryTransformEditorHtml(descriptor) {
  const result = queryTransformResultForDescriptor(descriptor);
  const destination = descriptor.targetPoint
    ? `${state.model?.sheets?.[descriptor.targetPoint.sheetIndex]?.name || activeSheet()?.name || ""}!${columnName(descriptor.targetPoint.col)}${descriptor.targetPoint.row}`
    : "";
  const columns = result.columns.length ? result.columns : ["Column1"];
  const hasPromotedHeaders = queryTransformHasStep(descriptor, "promoteHeaders");
  const selectedColumn = columns.includes(descriptor.selectedColumn) ? descriptor.selectedColumn : columns[0] || "";
  const selectedDataType = descriptor.selectedDataType || "text";
  const rowCount = Number.isFinite(Number(descriptor.pendingRowCount)) ? Math.max(1, Math.round(Number(descriptor.pendingRowCount))) : 10;
  const rangeStart = Number.isFinite(Number(descriptor.pendingRangeStart)) ? Math.max(1, Math.round(Number(descriptor.pendingRangeStart))) : 1;
  const rangeCount = Number.isFinite(Number(descriptor.pendingRangeCount)) ? Math.max(1, Math.round(Number(descriptor.pendingRangeCount))) : Math.max(1, rowCount);
  const groupByColumn = columns.includes(descriptor.pendingGroupByColumn) ? descriptor.pendingGroupByColumn : selectedColumn;
  const aggregateColumn = columns.includes(descriptor.pendingAggregateColumn) ? descriptor.pendingAggregateColumn : selectedColumn;
  const aggregateFunction = descriptor.pendingAggregateFunction || "sum";
  const pivotColumn = columns.includes(descriptor.pendingPivotColumn) ? descriptor.pendingPivotColumn : selectedColumn;
  const pivotValueColumn = columns.includes(descriptor.pendingPivotValueColumn)
    ? descriptor.pendingPivotValueColumn
    : columns.find((column) => column !== pivotColumn) || selectedColumn;
  const pivotAggregateFunction = descriptor.pendingPivotAggregateFunction || "sum";
  const conditionColumn = columns.includes(descriptor.pendingConditionColumn) ? descriptor.pendingConditionColumn : selectedColumn;
  const filterOperator = descriptor.pendingFilterOperator || "equals";
  return `
    <div class="excel-mini-dialog data-dialog query-transform-dialog" role="dialog" aria-modal="true" aria-labelledby="queryTransformTitle" data-data-dialog="query-transform-editor">
      <div class="excel-mini-titlebar" id="queryTransformTitle">Power Query エディター - ${escapeHtml(descriptor.sourceName || "CSV")}</div>
      <div class="power-query-ribbon">
        <div class="power-query-command-group">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="query-transform-load">閉じて読み込む</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-load-to">閉じて次に読み込む...</button>
        </div>
        <div class="power-query-command-group">
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-promote-headers" ${hasPromotedHeaders ? "disabled" : ""}>先頭行をヘッダーとして使用</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-demote-headers">ヘッダーを先頭行として使用</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-transpose">入れ替え</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-remove-columns">列の削除</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-keep-columns">列の保持</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-remove-empty-columns">空の列の削除</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-remove-duplicates">重複の削除</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-keep-duplicates">重複の保持</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-remove-blank-rows">空白行の削除</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-keep-blank-rows">空白行の保持</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-count-rows">行数のカウント</button>
        </div>
        <div class="power-query-command-group power-query-column-tools">
          <select data-data-field="queryActionColumn">${columns.map((column) => `<option value="${escapeAttr(column)}"${column === selectedColumn ? " selected" : ""}>${escapeHtml(column)}</option>`).join("")}</select>
          <input type="text" data-data-field="queryRenameTo" value="${escapeAttr(descriptor.pendingRenameTo || selectedColumn)}" aria-label="新しい列名" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-rename-column">列名の変更</button>
          <select data-data-field="queryDataType">
            ${queryTransformDataTypeOptions(selectedDataType)}
          </select>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-change-type">データ型</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-detect-type">データ型の検出</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-sort-asc">昇順</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-sort-desc">降順</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-duplicate-column">列の複製</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-move-start">先頭へ移動</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-move-left">左へ移動</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-move-right">右へ移動</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-move-end">末尾へ移動</button>
        </div>
        <div class="power-query-command-group power-query-row-tools">
          <input type="number" min="1" max="1048576" data-data-field="queryRowCount" value="${escapeAttr(String(rowCount))}" aria-label="行数" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-keep-top-rows">上位行の保持</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-remove-top-rows">上位行の削除</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-keep-bottom-rows">下位行の保持</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-remove-bottom-rows">下位行の削除</button>
          <input type="number" min="1" max="1048576" data-data-field="queryRangeStart" value="${escapeAttr(String(rangeStart))}" aria-label="開始行" />
          <input type="number" min="1" max="1048576" data-data-field="queryRangeCount" value="${escapeAttr(String(rangeCount))}" aria-label="保持行数" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-keep-range-rows">行範囲の保持</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-remove-range-rows">行範囲の削除</button>
          <input type="number" min="0" max="1048576" data-data-field="queryIndexStart" value="${escapeAttr(String(Number.isFinite(Number(descriptor.pendingIndexStart)) ? Math.max(0, Math.round(Number(descriptor.pendingIndexStart))) : 0))}" aria-label="開始インデックス" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-add-index-column">インデックス列</button>
        </div>
        <div class="power-query-command-group power-query-text-tools">
          <input type="text" data-data-field="queryReplaceFrom" value="${escapeAttr(descriptor.pendingReplaceFrom || "")}" placeholder="検索する値" />
          <input type="text" data-data-field="queryReplaceTo" value="${escapeAttr(descriptor.pendingReplaceTo || "")}" placeholder="置換後" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-replace-values">値の置換</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-replace-blank">空白の置換</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-uppercase">大文字</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-lowercase">小文字</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-capitalize">各単語を大文字</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-text-length">長さ</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-trim">トリム</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-clean">クリーン</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-fill-down">下方向へコピー</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-fill-up">上方向へコピー</button>
          <input type="number" min="1" max="32767" data-data-field="queryExtractCount" value="${escapeAttr(String(Number.isFinite(Number(descriptor.pendingExtractCount)) ? Math.max(1, Math.round(Number(descriptor.pendingExtractCount))) : 3))}" aria-label="文字数" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-extract-first">最初の文字</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-extract-last">最後の文字</button>
          <input type="number" min="0" max="32767" data-data-field="queryExtractStart" value="${escapeAttr(String(Number.isFinite(Number(descriptor.pendingExtractStart)) ? Math.max(0, Math.round(Number(descriptor.pendingExtractStart))) : 0))}" aria-label="開始位置" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-extract-range">指定範囲</button>
          <input type="text" data-data-field="querySplitDelimiter" value="${escapeAttr(descriptor.pendingSplitDelimiter || "")}" placeholder="区切り記号" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-split-column">列の分割</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-split-count">文字数で分割</button>
          <input type="text" data-data-field="querySplitPositions" value="${escapeAttr(descriptor.pendingSplitPositions || "")}" placeholder="位置 8,16" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-split-positions">位置で分割</button>
          <input type="text" data-data-field="queryExtractDelimiter" value="${escapeAttr(descriptor.pendingExtractDelimiter || "")}" placeholder="抽出区切り" />
          <input type="text" data-data-field="queryExtractEndDelimiter" value="${escapeAttr(descriptor.pendingExtractEndDelimiter || "")}" placeholder="終了区切り" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-extract-before">区切り記号の前</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-extract-after">区切り記号の後</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-extract-between">区切り記号の間</button>
          <input type="text" data-data-field="queryPrefixText" value="${escapeAttr(descriptor.pendingPrefixText || "")}" placeholder="プレフィックス" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-add-prefix">プレフィックス</button>
          <input type="text" data-data-field="querySuffixText" value="${escapeAttr(descriptor.pendingSuffixText || "")}" placeholder="サフィックス" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-add-suffix">サフィックス</button>
          <input type="text" data-data-field="queryCharacterSet" value="${escapeAttr(descriptor.pendingCharacterSet || "")}" placeholder="文字セット" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-keep-characters">文字の保持</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-remove-characters">文字の削除</button>
        </div>
        <div class="power-query-command-group power-query-table-tools">
          <input type="text" data-data-field="queryMergeDelimiter" value="${escapeAttr(descriptor.pendingMergeDelimiter || "-")}" placeholder="結合区切り" />
          <input type="text" data-data-field="queryMergeName" value="${escapeAttr(descriptor.pendingMergeName || "結合済み")}" placeholder="列名" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-merge-columns">列のマージ</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-unpivot-columns">列のピボット解除</button>
          <select data-data-field="queryPivotColumn" aria-label="ピボット列">
            ${columns.map((column) => `<option value="${escapeAttr(column)}"${column === pivotColumn ? " selected" : ""}>${escapeHtml(column)}</option>`).join("")}
          </select>
          <select data-data-field="queryPivotValueColumn" aria-label="値列">
            ${columns.map((column) => `<option value="${escapeAttr(column)}"${column === pivotValueColumn ? " selected" : ""}>${escapeHtml(column)}</option>`).join("")}
          </select>
          <select data-data-field="queryPivotAggregateFunction" aria-label="集計関数">
            ${queryTransformAggregateOptions(pivotAggregateFunction)}
          </select>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-pivot-column">列のピボット</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-reverse-rows">行の反転</button>
          <input type="number" min="1" max="1048576" data-data-field="queryAlternateStart" value="${escapeAttr(String(Number.isFinite(Number(descriptor.pendingAlternateStart)) ? Math.max(1, Math.round(Number(descriptor.pendingAlternateStart))) : 2))}" aria-label="開始行" />
          <input type="number" min="1" max="1048576" data-data-field="queryAlternateRemove" value="${escapeAttr(String(Number.isFinite(Number(descriptor.pendingAlternateRemove)) ? Math.max(1, Math.round(Number(descriptor.pendingAlternateRemove))) : 1))}" aria-label="削除行数" />
          <input type="number" min="1" max="1048576" data-data-field="queryAlternateKeep" value="${escapeAttr(String(Number.isFinite(Number(descriptor.pendingAlternateKeep)) ? Math.max(1, Math.round(Number(descriptor.pendingAlternateKeep))) : 1))}" aria-label="保持行数" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-remove-alternate-rows">代替行の削除</button>
        </div>
        <div class="power-query-command-group power-query-number-tools">
          <input type="number" step="any" data-data-field="queryNumberOperand" value="${escapeAttr(String(Number.isFinite(Number(descriptor.pendingNumberOperand)) ? Number(descriptor.pendingNumberOperand) : 1))}" aria-label="数値" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-add">加算</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-subtract">減算</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-multiply">乗算</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-divide">除算</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-integer-divide">整数除算</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-modulo">剰余</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-round">丸め</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-floor">切り捨て</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-ceil">切り上げ</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-absolute">絶対値</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-sign">符号</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-sqrt">平方根</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-log10">10 を底とする対数</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-power">べき乗</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-percentage">パーセント</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-number-percent-of">パーセント比</button>
          <input type="text" data-data-field="queryErrorReplacement" value="${escapeAttr(descriptor.pendingErrorReplacement || "")}" placeholder="エラー置換" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-replace-errors">エラーの置換</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-remove-errors">エラーの削除</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-keep-errors">エラーの保持</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-year">年</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-month">月</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-day">日</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-quarter">四半期</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-month-name">月の名前</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-day-name">曜日の名前</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-day-of-week">曜日</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-day-of-year">年内の日</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-start-week">週の開始</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-end-week">週の終了</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-days-in-month">月の日数</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-start-month">月の開始</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-end-month">月の終了</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-start-quarter">四半期の開始</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-end-quarter">四半期の終了</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-start-year">年の開始</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-date-end-year">年の終了</button>
        </div>
        <div class="power-query-command-group power-query-add-tools">
          <input type="text" data-data-field="queryCustomColumnName" value="${escapeAttr(descriptor.pendingCustomColumnName || "カスタム")}" placeholder="列名" />
          <input type="text" data-data-field="queryCustomFormula" value="${escapeAttr(descriptor.pendingCustomFormula || "")}" placeholder='[列] & "値"' />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-add-custom-column">カスタム列</button>
          <select data-data-field="queryGroupByColumn">
            ${columns.map((column) => `<option value="${escapeAttr(column)}"${column === groupByColumn ? " selected" : ""}>${escapeHtml(column)}</option>`).join("")}
          </select>
          <select data-data-field="queryAggregateColumn">
            ${columns.map((column) => `<option value="${escapeAttr(column)}"${column === aggregateColumn ? " selected" : ""}>${escapeHtml(column)}</option>`).join("")}
          </select>
          <select data-data-field="queryAggregateFunction">
            ${queryTransformAggregateOptions(aggregateFunction)}
          </select>
          <input type="text" data-data-field="queryAggregateName" value="${escapeAttr(descriptor.pendingAggregateName || "集計")}" placeholder="新しい列名" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-group-by">グループ化</button>
        </div>
        <div class="power-query-command-group power-query-condition-tools">
          <input type="text" data-data-field="queryConditionName" value="${escapeAttr(descriptor.pendingConditionName || "条件")}" placeholder="列名" />
          <select data-data-field="queryConditionColumn">
            ${columns.map((column) => `<option value="${escapeAttr(column)}"${column === conditionColumn ? " selected" : ""}>${escapeHtml(column)}</option>`).join("")}
          </select>
          <select data-data-field="queryConditionOperator">
            ${queryTransformConditionOperatorOptions(descriptor.pendingConditionOperator || "equals")}
          </select>
          <input type="text" data-data-field="queryConditionValue" value="${escapeAttr(descriptor.pendingConditionValue || "")}" placeholder="値" />
          <input type="text" data-data-field="queryConditionTrueValue" value="${escapeAttr(descriptor.pendingConditionTrueValue || "")}" placeholder="一致" />
          <input type="text" data-data-field="queryConditionFalseValue" value="${escapeAttr(descriptor.pendingConditionFalseValue || "")}" placeholder="不一致" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-add-conditional-column">条件列</button>
        </div>
        <div class="power-query-filter">
          <select data-data-field="queryFilterColumn">
            <option value="__all__">すべての列</option>
            ${columns.map((column) => `<option value="${escapeAttr(column)}">${escapeHtml(column)}</option>`).join("")}
          </select>
          <select data-data-field="queryFilterOperator">
            ${queryTransformConditionOperatorOptions(filterOperator)}
          </select>
          <input type="text" data-data-field="queryFilterText" value="${escapeAttr(descriptor.pendingFilterText || "")}" placeholder="フィルター テキスト" />
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-apply-filter">フィルター</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-clear-filter">クリア</button>
        </div>
      </div>
      <div class="excel-mini-body data-dialog-body power-query-body">
        <aside class="power-query-pane power-query-queries">
          <strong>クエリ</strong>
          <button type="button" class="is-active">${escapeHtml(queryTransformName(descriptor))}</button>
        </aside>
        <section class="power-query-main">
          <div class="power-query-formula-bar">= Source{${escapeHtml(JSON.stringify(descriptor.sourceName || "CSV"))}}</div>
          <div class="power-query-column-list" aria-label="列の選択">${queryTransformColumnChecklistHtml(columns)}</div>
          <div class="query-preview query-transform-preview" aria-label="プレビュー">${queryTransformPreviewHtml(result)}</div>
          <div class="data-dialog-note">読み込み先: ${escapeHtml(destination)} / ${result.outputRows.length} 行、${result.columns.length} 列</div>
        </section>
        <aside class="power-query-pane power-query-settings">
          <strong>クエリの設定</strong>
          <label class="excel-dialog-field power-query-name-field"><span class="excel-dialog-caption">名前</span><input type="text" data-data-field="queryName" value="${escapeAttr(queryTransformName(descriptor))}" /></label>
          <div class="power-query-steps">
            <span class="excel-dialog-caption">適用したステップ</span>
            ${queryTransformStepsHtml(descriptor)}
          </div>
        </aside>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="query-transform-load">閉じて読み込む</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="query-transform-load-to">閉じて次に読み込む...</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
}



function queryTransformColumnChecklistHtml(columns) {
  return columns
    .map((column) => `
      <label class="data-dialog-check">
        <input type="checkbox" data-query-transform-column value="${escapeAttr(column)}" />
        <span>${escapeHtml(column)}</span>
      </label>`)
    .join("");
}



function queryTransformDataTypeOptions(selected) {
  const options = [
    ["text", "テキスト"],
    ["number", "10 進数"],
    ["whole", "整数"],
    ["date", "日付"],
    ["datetime", "日付/時刻"],
    ["any", "任意"],
  ];
  return options
    .map(([value, label]) => `<option value="${escapeAttr(value)}"${value === selected ? " selected" : ""}>${escapeHtml(label)}</option>`)
    .join("");
}



function queryTransformAggregateOptions(selected) {
  const options = [
    ["sum", "合計"],
    ["average", "平均"],
    ["count", "行数"],
    ["countNums", "数値の個数"],
    ["countDistinct", "個別の個数"],
    ["min", "最小"],
    ["max", "最大"],
    ["product", "積"],
    ["median", "中央値"],
    ["stdDev", "標準偏差"],
    ["stdDevP", "標準偏差P"],
    ["var", "分散"],
    ["varP", "分散P"],
  ];
  return options
    .map(([value, label]) => `<option value="${escapeAttr(value)}"${value === selected ? " selected" : ""}>${escapeHtml(label)}</option>`)
    .join("");
}



function queryTransformConditionOperatorOptions(selected) {
  const options = [
    ["equals", "等しい"],
    ["notEquals", "等しくない"],
    ["contains", "含む"],
    ["notContains", "含まない"],
    ["beginsWith", "指定の値で始まる"],
    ["notBeginsWith", "指定の値で始まらない"],
    ["endsWith", "指定の値で終わる"],
    ["notEndsWith", "指定の値で終わらない"],
    ["greaterThan", "より大きい"],
    ["greaterOrEqual", "以上"],
    ["lessThan", "より小さい"],
    ["lessOrEqual", "以下"],
    ["blank", "空白"],
    ["notBlank", "空白ではない"],
  ];
  return options
    .map(([value, label]) => `<option value="${escapeAttr(value)}"${value === selected ? " selected" : ""}>${escapeHtml(label)}</option>`)
    .join("");
}



function queryTransformStepsHtml(descriptor) {
  const steps = Array.isArray(descriptor.transformSteps) ? descriptor.transformSteps : [];
  const sourceStep = '<div class="power-query-step is-source"><span>ソース</span></div>';
  if (!steps.length) return `${sourceStep}<div class="power-query-step"><span>変更された型</span></div>`;
  return `${sourceStep}${steps.map((step, index) => `
    <div class="power-query-step">
      <span>${escapeHtml(queryTransformStepLabel(step))}</span>
      <button type="button" data-data-dialog-action="query-transform-delete-step" data-query-transform-step-index="${index}" aria-label="ステップを削除">×</button>
    </div>`).join("")}`;
}



function queryTransformStepLabel(step) {
  const targetColumns = queryTransformStepColumnLabel(step);
  if (step?.type === "promoteHeaders") return "昇格されたヘッダー";
  if (step?.type === "demoteHeaders") return "降格されたヘッダー";
  if (step?.type === "transpose") return "入れ替えられたテーブル";
  if (step?.type === "removeColumns") return `削除された列: ${(step.columns || []).join(", ")}`;
  if (step?.type === "keepColumns") return `保持された列: ${(step.columns || []).join(", ")}`;
  if (step?.type === "removeEmptyColumns") return "削除された空の列";
  if (step?.type === "countRows") return "カウントされた行";
  if (step?.type === "renameColumn") return `名前が変更された列: ${step.from || ""} → ${step.to || ""}`;
  if (step?.type === "changeType") return `変更された型: ${targetColumns} = ${queryTransformDataTypeLabel(step.dataType)}`;
  if (step?.type === "detectTypes") return `検出された型: ${targetColumns || "すべての列"}`;
  if (step?.type === "numberOperation") return `変換された数値: ${targetColumns} ${queryTransformNumberOperationLabel(step.operation)}`;
  if (step?.type === "datePart") return `抽出された日付: ${targetColumns} ${queryTransformDatePartLabel(step.part)}`;
  if (step?.type === "replaceErrors") return `置換されたエラー: ${targetColumns}`;
  if (step?.type === "removeErrors") return `削除されたエラー: ${targetColumns}`;
  if (step?.type === "keepErrors") return `保持されたエラー: ${targetColumns}`;
  if (step?.type === "sortRows") return `並べ替えられた行: ${step.columnName || ""} ${step.direction === "desc" ? "降順" : "昇順"}`;
  if (step?.type === "duplicateColumn") return `複製された列: ${step.columnName || ""}`;
  if (step?.type === "moveColumn") return `移動された列: ${step.columnName || ""} ${queryTransformMoveLabel(step.position)}`;
  if (step?.type === "addIndexColumn") return `追加されたインデックス列: ${step.name || "インデックス"}`;
  if (step?.type === "removeDuplicates") return `削除された重複: ${(step.columns || []).join(", ") || "すべての列"}`;
  if (step?.type === "keepDuplicates") return `保持された重複: ${(step.columns || []).join(", ") || "すべての列"}`;
  if (step?.type === "removeBlankRows") return "削除された空白行";
  if (step?.type === "keepBlankRows") return "保持された空白行";
  if (step?.type === "keepTopRows") return `保持された上位行: ${step.count || 0}`;
  if (step?.type === "removeTopRows") return `削除された上位行: ${step.count || 0}`;
  if (step?.type === "keepBottomRows") return `保持された下位行: ${step.count || 0}`;
  if (step?.type === "removeBottomRows") return `削除された下位行: ${step.count || 0}`;
  if (step?.type === "keepRangeRows") return `保持された行範囲: ${step.start || 1} から ${step.count || 0} 行`;
  if (step?.type === "removeRangeRows") return `削除された行範囲: ${step.start || 1} から ${step.count || 0} 行`;
  if (step?.type === "replaceValues") return `置換された値: ${step.columnName || ""}`;
  if (step?.type === "replaceBlank") return `置換された空白: ${step.columnName || ""}`;
  if (step?.type === "transformTextCase") return `変更されたテキスト: ${step.columnName || ""} ${queryTransformTextCaseLabel(step.mode)}`;
  if (step?.type === "textLength") return `計算された長さ: ${targetColumns}`;
  if (step?.type === "trim") return `トリムされたテキスト: ${step.columnName || ""}`;
  if (step?.type === "clean") return `クリーンされたテキスト: ${step.columnName || ""}`;
  if (step?.type === "fillDown") return `下方向へコピー: ${step.columnName || ""}`;
  if (step?.type === "fillUp") return `上方向へコピー: ${step.columnName || ""}`;
  if (step?.type === "extractFirstCharacters") return `抽出された最初の文字: ${step.columnName || ""} ${step.count || 0}`;
  if (step?.type === "extractLastCharacters") return `抽出された最後の文字: ${step.columnName || ""} ${step.count || 0}`;
  if (step?.type === "extractRangeCharacters") return `抽出された範囲: ${step.columnName || ""} ${step.start || 0}, ${step.count || 0}`;
  if (step?.type === "splitColumnByDelimiter") return `区切り記号による列の分割: ${step.columnName || ""}`;
  if (step?.type === "splitColumnByCharacterCount") return `文字数による列の分割: ${step.columnName || ""} ${step.count || 0}`;
  if (step?.type === "splitColumnByPositions") return `位置による列の分割: ${step.columnName || ""} ${(step.positions || []).join(", ")}`;
  if (step?.type === "extractBeforeDelimiter") return `抽出されたテキスト: ${step.columnName || ""} 区切り記号の前`;
  if (step?.type === "extractAfterDelimiter") return `抽出されたテキスト: ${step.columnName || ""} 区切り記号の後`;
  if (step?.type === "extractBetweenDelimiters") return `抽出されたテキスト: ${step.columnName || ""} 区切り記号の間`;
  if (step?.type === "addPrefix") return `追加されたプレフィックス: ${step.columnName || ""}`;
  if (step?.type === "addSuffix") return `追加されたサフィックス: ${step.columnName || ""}`;
  if (step?.type === "keepCharacters") return `保持された文字: ${step.columnName || ""}`;
  if (step?.type === "removeCharacters") return `削除された文字: ${step.columnName || ""}`;
  if (step?.type === "mergeColumns") return `マージされた列: ${(step.columns || []).join(", ")}`;
  if (step?.type === "unpivotColumns") return `ピボット解除された列: ${(step.columns || []).join(", ")}`;
  if (step?.type === "pivotColumn") return `ピボットされた列: ${step.columnName || ""} / ${step.valueColumn || ""}`;
  if (step?.type === "reverseRows") return "反転された行";
  if (step?.type === "removeAlternateRows") return `削除された代替行: 開始 ${step.start || 1}`;
  if (step?.type === "addCustomColumn") return `追加されたカスタム列: ${step.name || ""}`;
  if (step?.type === "addConditionalColumn") return `追加された条件列: ${step.name || ""}`;
  if (step?.type === "groupBy") {
    const groupColumns = queryTransformGroupColumns(step);
    return `グループ化された行: ${groupColumns.join(", ")} / ${queryTransformAggregateLabel(step.aggregateFunction)} ${step.aggregateColumn || ""}`;
  }
  if (step?.type === "filterRows") {
    const column = step.columnName && step.columnName !== "__all__" ? step.columnName : "すべての列";
    return `フィルターされた行: ${column} ${queryTransformConditionOperatorLabel(step.operator || "equals")} ${step.text || ""}`.trim();
  }
  return "変更されたステップ";
}



function queryTransformStepColumnLabel(step) {
  const columns = Array.isArray(step?.columns) && step.columns.length
    ? step.columns
    : [step?.columnName].filter(Boolean);
  return columns.join(", ");
}



function queryTransformGroupColumns(step) {
  const columns = Array.isArray(step?.groupColumns) && step.groupColumns.length
    ? step.groupColumns
    : [step?.groupColumn].filter(Boolean);
  return columns.map((column) => String(column || "")).filter(Boolean);
}



function queryTransformMoveLabel(position) {
  return {
    start: "先頭",
    left: "左",
    right: "右",
    end: "末尾",
  }[position] || "";
}



function queryTransformDataTypeLabel(value) {
  return {
    text: "テキスト",
    number: "10 進数",
    whole: "整数",
    date: "日付",
    datetime: "日付/時刻",
    any: "任意",
  }[value] || "任意";
}



function queryTransformAggregateLabel(value) {
  return {
    sum: "合計",
    average: "平均",
    count: "行数",
    countNums: "数値の個数",
    countDistinct: "個別の個数",
    min: "最小",
    max: "最大",
    product: "積",
    median: "中央値",
    stdDev: "標準偏差",
    stdDevP: "標準偏差P",
    var: "分散",
    varP: "分散P",
  }[value] || "合計";
}



function queryTransformTextCaseLabel(value) {
  return {
    uppercase: "大文字",
    lowercase: "小文字",
    capitalize: "各単語を大文字",
  }[value] || "大文字";
}



function queryTransformNumberOperationLabel(value) {
  return {
    add: "加算",
    subtract: "減算",
    multiply: "乗算",
    divide: "除算",
    integerDivide: "整数除算",
    modulo: "剰余",
    round: "丸め",
    floor: "切り捨て",
    ceil: "切り上げ",
    absolute: "絶対値",
    sign: "符号",
    squareRoot: "平方根",
    logarithm: "対数",
    power: "べき乗",
    percentage: "パーセント",
    percentOf: "パーセント比",
  }[value] || "変換";
}



function queryTransformDatePartLabel(value) {
  return {
    year: "年",
    month: "月",
    day: "日",
    quarter: "四半期",
    monthName: "月の名前",
    dayName: "曜日の名前",
    dayOfWeek: "曜日",
    dayOfYear: "年内の日",
    startOfWeek: "週の開始",
    endOfWeek: "週の終了",
    daysInMonth: "月の日数",
    startOfMonth: "月の開始",
    endOfMonth: "月の終了",
    startOfQuarter: "四半期の開始",
    endOfQuarter: "四半期の終了",
    startOfYear: "年の開始",
    endOfYear: "年の終了",
  }[value] || "日付";
}



function queryTransformConditionOperatorLabel(value) {
  return {
    equals: "=",
    notEquals: "<>",
    contains: "を含む",
    notContains: "を含まない",
    beginsWith: "で始まる",
    notBeginsWith: "で始まらない",
    endsWith: "で終わる",
    notEndsWith: "で終わらない",
    greaterThan: ">",
    greaterOrEqual: ">=",
    lessThan: "<",
    lessOrEqual: "<=",
    blank: "空白",
    notBlank: "空白ではない",
  }[value] || "=";
}



function queryTransformPreviewHtml(result) {
  if (!result.outputRows.length || !result.columns.length) return '<div class="query-preview-empty">プレビューできるデータがありません。</div>';
  const headers = result.columns;
  const bodyRows = result.rows.slice(0, 100);
  const head = `<thead><tr>${headers.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr></thead>`;
  const body = bodyRows
    .map((row) => `<tr>${headers.map((_column, index) => `<td>${escapeHtml(row[index] ?? "")}</td>`).join("")}</tr>`)
    .join("");
  return `<table>${head}<tbody>${body}</tbody></table>`;
}



function handleQueryTransformEditorAction(dialog, action, event) {
  if (!state.dataDialog || state.dataDialog.type !== "query-transform-editor") return;
  syncQueryTransformEditorInputs(dialog);
  if (action === "query-transform-promote-headers") {
    if (!queryTransformHasStep(state.dataDialog, "promoteHeaders")) {
      state.dataDialog.transformSteps.push({ type: "promoteHeaders" });
    }
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-remove-columns']");
    return;
  }
  if (action === "query-transform-demote-headers") {
    state.dataDialog.transformSteps.push({ type: "demoteHeaders" });
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-demote-headers']");
    return;
  }
  if (action === "query-transform-transpose") {
    state.dataDialog.transformSteps.push({ type: "transpose" });
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-transpose']");
    return;
  }
  if (action === "query-transform-remove-columns") {
    const columns = [...dialog.querySelectorAll("[data-query-transform-column]:checked")].map((input) => input.value);
    if (!columns.length) {
      showMessageBox("削除する列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({ type: "removeColumns", columns });
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-remove-columns']");
    return;
  }
  if (action === "query-transform-keep-columns") {
    const columns = [...dialog.querySelectorAll("[data-query-transform-column]:checked")].map((input) => input.value);
    const keepColumns = columns.length ? columns : [state.dataDialog.selectedColumn].filter(Boolean);
    if (!keepColumns.length) {
      showMessageBox("保持する列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({ type: "keepColumns", columns: keepColumns });
    state.dataDialog.selectedColumn = keepColumns[0] || "";
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-keep-columns']");
    return;
  }
  if (action === "query-transform-remove-empty-columns") {
    state.dataDialog.transformSteps.push({ type: "removeEmptyColumns" });
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-remove-empty-columns']");
    return;
  }
  if (action === "query-transform-count-rows") {
    state.dataDialog.transformSteps.push({ type: "countRows", name: "行数" });
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-count-rows']");
    return;
  }
  if (action === "query-transform-rename-column") {
    const columnName = state.dataDialog.selectedColumn;
    const nextName = String(state.dataDialog.pendingRenameTo || "").trim();
    if (!columnName || !nextName) {
      showMessageBox("列名を入力してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({ type: "renameColumn", from: columnName, to: nextName });
    state.dataDialog.selectedColumn = nextName;
    state.dataDialog.pendingRenameTo = nextName;
    renderQueryTransformEditorDialog("[data-data-field='queryActionColumn']");
    return;
  }
  if (action === "query-transform-change-type") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({
      type: "changeType",
      columnName: targetColumns[0],
      columns: targetColumns,
      dataType: state.dataDialog.selectedDataType || "text",
    });
    renderQueryTransformEditorDialog("[data-data-field='queryDataType']");
    return;
  }
  if (action === "query-transform-detect-type") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({ type: "detectTypes", columnName: targetColumns[0], columns: targetColumns });
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-detect-type']");
    return;
  }
  if (
    action === "query-transform-number-add" ||
    action === "query-transform-number-subtract" ||
    action === "query-transform-number-multiply" ||
    action === "query-transform-number-divide" ||
    action === "query-transform-number-integer-divide" ||
    action === "query-transform-number-modulo" ||
    action === "query-transform-number-round" ||
    action === "query-transform-number-floor" ||
    action === "query-transform-number-ceil" ||
    action === "query-transform-number-absolute" ||
    action === "query-transform-number-sign" ||
    action === "query-transform-number-sqrt" ||
    action === "query-transform-number-log10" ||
    action === "query-transform-number-power" ||
    action === "query-transform-number-percentage" ||
    action === "query-transform-number-percent-of"
  ) {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const operation = {
      "query-transform-number-add": "add",
      "query-transform-number-subtract": "subtract",
      "query-transform-number-multiply": "multiply",
      "query-transform-number-divide": "divide",
      "query-transform-number-integer-divide": "integerDivide",
      "query-transform-number-modulo": "modulo",
      "query-transform-number-round": "round",
      "query-transform-number-floor": "floor",
      "query-transform-number-ceil": "ceil",
      "query-transform-number-absolute": "absolute",
      "query-transform-number-sign": "sign",
      "query-transform-number-sqrt": "squareRoot",
      "query-transform-number-log10": "logarithm",
      "query-transform-number-power": "power",
      "query-transform-number-percentage": "percentage",
      "query-transform-number-percent-of": "percentOf",
    }[action];
    state.dataDialog.transformSteps.push({
      type: "numberOperation",
      columnName: targetColumns[0],
      columns: targetColumns,
      operation,
      operand: Number(state.dataDialog.pendingNumberOperand),
    });
    renderQueryTransformEditorDialog("[data-data-field='queryNumberOperand']");
    return;
  }
  if (
    action === "query-transform-date-year" ||
    action === "query-transform-date-month" ||
    action === "query-transform-date-day" ||
    action === "query-transform-date-quarter" ||
    action === "query-transform-date-month-name" ||
    action === "query-transform-date-day-name" ||
    action === "query-transform-date-day-of-week" ||
    action === "query-transform-date-day-of-year" ||
    action === "query-transform-date-start-week" ||
    action === "query-transform-date-end-week" ||
    action === "query-transform-date-days-in-month" ||
    action === "query-transform-date-start-month" ||
    action === "query-transform-date-end-month" ||
    action === "query-transform-date-start-quarter" ||
    action === "query-transform-date-end-quarter" ||
    action === "query-transform-date-start-year" ||
    action === "query-transform-date-end-year"
  ) {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const part = {
      "query-transform-date-year": "year",
      "query-transform-date-month": "month",
      "query-transform-date-day": "day",
      "query-transform-date-quarter": "quarter",
      "query-transform-date-month-name": "monthName",
      "query-transform-date-day-name": "dayName",
      "query-transform-date-day-of-week": "dayOfWeek",
      "query-transform-date-day-of-year": "dayOfYear",
      "query-transform-date-start-week": "startOfWeek",
      "query-transform-date-end-week": "endOfWeek",
      "query-transform-date-days-in-month": "daysInMonth",
      "query-transform-date-start-month": "startOfMonth",
      "query-transform-date-end-month": "endOfMonth",
      "query-transform-date-start-quarter": "startOfQuarter",
      "query-transform-date-end-quarter": "endOfQuarter",
      "query-transform-date-start-year": "startOfYear",
      "query-transform-date-end-year": "endOfYear",
    }[action];
    state.dataDialog.transformSteps.push({ type: "datePart", columnName: targetColumns[0], columns: targetColumns, part });
    renderQueryTransformEditorDialog("[data-data-field='queryActionColumn']");
    return;
  }
  if (action === "query-transform-replace-errors") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({
      type: "replaceErrors",
      columnName: targetColumns[0],
      columns: targetColumns,
      value: String(state.dataDialog.pendingErrorReplacement || ""),
    });
    renderQueryTransformEditorDialog("[data-data-field='queryErrorReplacement']");
    return;
  }
  if (action === "query-transform-remove-errors") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({ type: "removeErrors", columnName: targetColumns[0], columns: targetColumns });
    renderQueryTransformEditorDialog("[data-data-field='queryActionColumn']");
    return;
  }
  if (action === "query-transform-keep-errors") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({ type: "keepErrors", columnName: targetColumns[0], columns: targetColumns });
    renderQueryTransformEditorDialog("[data-data-field='queryActionColumn']");
    return;
  }
  if (action === "query-transform-sort-asc" || action === "query-transform-sort-desc") {
    if (!state.dataDialog.selectedColumn) {
      showMessageBox("並べ替える列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({
      type: "sortRows",
      columnName: state.dataDialog.selectedColumn,
      direction: action === "query-transform-sort-desc" ? "desc" : "asc",
    });
    renderQueryTransformEditorDialog("[data-data-field='queryActionColumn']");
    return;
  }
  if (action === "query-transform-duplicate-column") {
    if (!state.dataDialog.selectedColumn) {
      showMessageBox("複製する列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const newName = `${state.dataDialog.selectedColumn} - コピー`;
    state.dataDialog.transformSteps.push({ type: "duplicateColumn", columnName: state.dataDialog.selectedColumn, newName });
    state.dataDialog.selectedColumn = newName;
    renderQueryTransformEditorDialog("[data-data-field='queryActionColumn']");
    return;
  }
  if (action === "query-transform-move-start" || action === "query-transform-move-left" || action === "query-transform-move-right" || action === "query-transform-move-end") {
    if (!state.dataDialog.selectedColumn) {
      showMessageBox("移動する列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const position = {
      "query-transform-move-start": "start",
      "query-transform-move-left": "left",
      "query-transform-move-right": "right",
      "query-transform-move-end": "end",
    }[action];
    state.dataDialog.transformSteps.push({ type: "moveColumn", columnName: state.dataDialog.selectedColumn, position });
    renderQueryTransformEditorDialog("[data-data-field='queryActionColumn']");
    return;
  }
  if (action === "query-transform-remove-duplicates") {
    const selectedColumns = [...dialog.querySelectorAll("[data-query-transform-column]:checked")].map((input) => input.value);
    const columns = selectedColumns.length ? selectedColumns : queryTransformResultForDescriptor(state.dataDialog).columns;
    state.dataDialog.transformSteps.push({ type: "removeDuplicates", columns });
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-remove-duplicates']");
    return;
  }
  if (action === "query-transform-keep-duplicates") {
    const selectedColumns = [...dialog.querySelectorAll("[data-query-transform-column]:checked")].map((input) => input.value);
    const columns = selectedColumns.length ? selectedColumns : queryTransformResultForDescriptor(state.dataDialog).columns;
    state.dataDialog.transformSteps.push({ type: "keepDuplicates", columns });
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-keep-duplicates']");
    return;
  }
  if (action === "query-transform-remove-blank-rows") {
    state.dataDialog.transformSteps.push({ type: "removeBlankRows" });
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-remove-blank-rows']");
    return;
  }
  if (action === "query-transform-keep-blank-rows") {
    state.dataDialog.transformSteps.push({ type: "keepBlankRows" });
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-keep-blank-rows']");
    return;
  }
  if (action === "query-transform-keep-top-rows" || action === "query-transform-remove-top-rows") {
    const count = Math.max(1, Math.min(1048576, Math.round(Number(state.dataDialog.pendingRowCount) || 1)));
    state.dataDialog.transformSteps.push({
      type: action === "query-transform-keep-top-rows" ? "keepTopRows" : "removeTopRows",
      count,
    });
    renderQueryTransformEditorDialog("[data-data-field='queryRowCount']");
    return;
  }
  if (action === "query-transform-keep-bottom-rows" || action === "query-transform-remove-bottom-rows") {
    const count = Math.max(1, Math.min(1048576, Math.round(Number(state.dataDialog.pendingRowCount) || 1)));
    state.dataDialog.transformSteps.push({
      type: action === "query-transform-keep-bottom-rows" ? "keepBottomRows" : "removeBottomRows",
      count,
    });
    renderQueryTransformEditorDialog("[data-data-field='queryRowCount']");
    return;
  }
  if (action === "query-transform-keep-range-rows") {
    const start = Math.max(1, Math.min(1048576, Math.round(Number(state.dataDialog.pendingRangeStart) || 1)));
    const count = Math.max(1, Math.min(1048576, Math.round(Number(state.dataDialog.pendingRangeCount) || 1)));
    state.dataDialog.transformSteps.push({ type: "keepRangeRows", start, count });
    renderQueryTransformEditorDialog("[data-data-field='queryRangeStart']");
    return;
  }
  if (action === "query-transform-remove-range-rows") {
    const start = Math.max(1, Math.min(1048576, Math.round(Number(state.dataDialog.pendingRangeStart) || 1)));
    const count = Math.max(1, Math.min(1048576, Math.round(Number(state.dataDialog.pendingRangeCount) || 1)));
    state.dataDialog.transformSteps.push({ type: "removeRangeRows", start, count });
    renderQueryTransformEditorDialog("[data-data-field='queryRangeStart']");
    return;
  }
  if (action === "query-transform-add-index-column") {
    const start = Math.max(0, Math.min(1048576, Math.round(Number(state.dataDialog.pendingIndexStart) || 0)));
    state.dataDialog.transformSteps.push({ type: "addIndexColumn", name: "インデックス", start });
    state.dataDialog.selectedColumn = "インデックス";
    renderQueryTransformEditorDialog("[data-data-field='queryIndexStart']");
    return;
  }
  if (action === "query-transform-replace-values") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const from = String(state.dataDialog.pendingReplaceFrom || "");
    if (!from) {
      showMessageBox("検索する値を入力してください。", { title: "Cht WebSheet" });
      return;
    }
    targetColumns.forEach((columnName) => state.dataDialog.transformSteps.push({
      type: "replaceValues",
      columnName,
      from,
      to: String(state.dataDialog.pendingReplaceTo || ""),
    }));
    renderQueryTransformEditorDialog("[data-data-field='queryReplaceFrom']");
    return;
  }
  if (action === "query-transform-replace-blank") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    targetColumns.forEach((columnName) => state.dataDialog.transformSteps.push({
      type: "replaceBlank",
      columnName,
      value: String(state.dataDialog.pendingReplaceTo || ""),
    }));
    renderQueryTransformEditorDialog("[data-data-field='queryReplaceTo']");
    return;
  }
  if (action === "query-transform-trim" || action === "query-transform-clean" || action === "query-transform-fill-down" || action === "query-transform-fill-up") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const typeMap = {
      "query-transform-trim": "trim",
      "query-transform-clean": "clean",
      "query-transform-fill-down": "fillDown",
      "query-transform-fill-up": "fillUp",
    };
    targetColumns.forEach((columnName) => state.dataDialog.transformSteps.push({ type: typeMap[action], columnName }));
    renderQueryTransformEditorDialog("[data-data-field='queryActionColumn']");
    return;
  }
  if (action === "query-transform-uppercase" || action === "query-transform-lowercase" || action === "query-transform-capitalize") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const mode = {
      "query-transform-uppercase": "uppercase",
      "query-transform-lowercase": "lowercase",
      "query-transform-capitalize": "capitalize",
    }[action];
    targetColumns.forEach((columnName) => state.dataDialog.transformSteps.push({ type: "transformTextCase", columnName, mode }));
    renderQueryTransformEditorDialog("[data-data-field='queryActionColumn']");
    return;
  }
  if (action === "query-transform-text-length") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({ type: "textLength", columnName: targetColumns[0], columns: targetColumns });
    renderQueryTransformEditorDialog("[data-data-field='queryActionColumn']");
    return;
  }
  if (action === "query-transform-extract-first" || action === "query-transform-extract-last") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const count = Math.max(1, Math.min(32767, Math.round(Number(state.dataDialog.pendingExtractCount) || 1)));
    targetColumns.forEach((columnName) => state.dataDialog.transformSteps.push({
      type: action === "query-transform-extract-first" ? "extractFirstCharacters" : "extractLastCharacters",
      columnName,
      count,
    }));
    renderQueryTransformEditorDialog("[data-data-field='queryExtractCount']");
    return;
  }
  if (action === "query-transform-extract-range") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const start = Math.max(0, Math.min(32767, Math.round(Number(state.dataDialog.pendingExtractStart) || 0)));
    const count = Math.max(1, Math.min(32767, Math.round(Number(state.dataDialog.pendingExtractCount) || 1)));
    targetColumns.forEach((columnName) => state.dataDialog.transformSteps.push({
      type: "extractRangeCharacters",
      columnName,
      start,
      count,
    }));
    renderQueryTransformEditorDialog("[data-data-field='queryExtractStart']");
    return;
  }
  if (action === "query-transform-split-column") {
    if (!state.dataDialog.selectedColumn) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const delimiter = String(state.dataDialog.pendingSplitDelimiter || "");
    if (!delimiter) {
      showMessageBox("区切り記号を入力してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({ type: "splitColumnByDelimiter", columnName: state.dataDialog.selectedColumn, delimiter });
    renderQueryTransformEditorDialog("[data-data-field='queryActionColumn']");
    return;
  }
  if (action === "query-transform-split-count") {
    if (!state.dataDialog.selectedColumn) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const count = Math.max(1, Math.min(32767, Math.round(Number(state.dataDialog.pendingExtractCount) || 1)));
    state.dataDialog.transformSteps.push({ type: "splitColumnByCharacterCount", columnName: state.dataDialog.selectedColumn, count });
    renderQueryTransformEditorDialog("[data-data-field='queryExtractCount']");
    return;
  }
  if (action === "query-transform-split-positions") {
    if (!state.dataDialog.selectedColumn) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const positions = parseFixedWidthBreaks(state.dataDialog.pendingSplitPositions || "");
    if (!positions.length) {
      showMessageBox("分割位置を入力してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({ type: "splitColumnByPositions", columnName: state.dataDialog.selectedColumn, positions });
    renderQueryTransformEditorDialog("[data-data-field='querySplitPositions']");
    return;
  }
  if (action === "query-transform-extract-before" || action === "query-transform-extract-after") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const delimiter = String(state.dataDialog.pendingExtractDelimiter || "");
    if (!delimiter) {
      showMessageBox("区切り記号を入力してください。", { title: "Cht WebSheet" });
      return;
    }
    targetColumns.forEach((columnName) => state.dataDialog.transformSteps.push({
      type: action === "query-transform-extract-before" ? "extractBeforeDelimiter" : "extractAfterDelimiter",
      columnName,
      delimiter,
    }));
    renderQueryTransformEditorDialog("[data-data-field='queryExtractDelimiter']");
    return;
  }
  if (action === "query-transform-extract-between") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const startDelimiter = String(state.dataDialog.pendingExtractDelimiter || "");
    const endDelimiter = String(state.dataDialog.pendingExtractEndDelimiter || "");
    if (!startDelimiter || !endDelimiter) {
      showMessageBox("開始と終了の区切り記号を入力してください。", { title: "Cht WebSheet" });
      return;
    }
    targetColumns.forEach((columnName) => state.dataDialog.transformSteps.push({
      type: "extractBetweenDelimiters",
      columnName,
      startDelimiter,
      endDelimiter,
    }));
    renderQueryTransformEditorDialog("[data-data-field='queryExtractDelimiter']");
    return;
  }
  if (action === "query-transform-add-prefix" || action === "query-transform-add-suffix") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const value = action === "query-transform-add-prefix"
      ? String(state.dataDialog.pendingPrefixText || "")
      : String(state.dataDialog.pendingSuffixText || "");
    if (!value) {
      showMessageBox("追加する文字列を入力してください。", { title: "Cht WebSheet" });
      return;
    }
    targetColumns.forEach((columnName) => state.dataDialog.transformSteps.push({
      type: action === "query-transform-add-prefix" ? "addPrefix" : "addSuffix",
      columnName,
      value,
    }));
    renderQueryTransformEditorDialog(action === "query-transform-add-prefix" ? "[data-data-field='queryPrefixText']" : "[data-data-field='querySuffixText']");
    return;
  }
  if (action === "query-transform-keep-characters" || action === "query-transform-remove-characters") {
    const targetColumns = selectedQueryTransformActionColumns(dialog);
    if (!targetColumns.length) {
      showMessageBox("列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const characters = String(state.dataDialog.pendingCharacterSet || "");
    if (!characters) {
      showMessageBox("対象の文字セットを入力してください。", { title: "Cht WebSheet" });
      return;
    }
    targetColumns.forEach((columnName) => state.dataDialog.transformSteps.push({
      type: action === "query-transform-keep-characters" ? "keepCharacters" : "removeCharacters",
      columnName,
      characters,
    }));
    renderQueryTransformEditorDialog("[data-data-field='queryCharacterSet']");
    return;
  }
  if (action === "query-transform-merge-columns") {
    const columns = [...dialog.querySelectorAll("[data-query-transform-column]:checked")].map((input) => input.value);
    if (columns.length < 2) {
      showMessageBox("マージする列を 2 つ以上選択してください。", { title: "Cht WebSheet" });
      return;
    }
    const name = String(state.dataDialog.pendingMergeName || "").trim() || "結合済み";
    state.dataDialog.transformSteps.push({
      type: "mergeColumns",
      columns,
      delimiter: String(state.dataDialog.pendingMergeDelimiter || ""),
      name,
    });
    state.dataDialog.selectedColumn = name;
    renderQueryTransformEditorDialog("[data-data-field='queryMergeName']");
    return;
  }
  if (action === "query-transform-unpivot-columns") {
    const columns = [...dialog.querySelectorAll("[data-query-transform-column]:checked")].map((input) => input.value);
    const unpivotColumns = columns.length ? columns : [state.dataDialog.selectedColumn].filter(Boolean);
    if (!unpivotColumns.length) {
      showMessageBox("ピボット解除する列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({ type: "unpivotColumns", columns: unpivotColumns, attributeName: "属性", valueName: "値" });
    state.dataDialog.selectedColumn = "属性";
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-unpivot-columns']");
    return;
  }
  if (action === "query-transform-pivot-column") {
    const columnName = state.dataDialog.pendingPivotColumn || state.dataDialog.selectedColumn;
    const valueColumn = state.dataDialog.pendingPivotValueColumn;
    if (!columnName || !valueColumn) {
      showMessageBox("ピボット列と値列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    if (columnName === valueColumn) {
      showMessageBox("ピボット列と値列には別の列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({
      type: "pivotColumn",
      columnName,
      valueColumn,
      aggregateFunction: state.dataDialog.pendingPivotAggregateFunction || "sum",
    });
    state.dataDialog.selectedColumn = valueColumn;
    renderQueryTransformEditorDialog("[data-data-field='queryPivotColumn']");
    return;
  }
  if (action === "query-transform-reverse-rows") {
    state.dataDialog.transformSteps.push({ type: "reverseRows" });
    renderQueryTransformEditorDialog("[data-data-dialog-action='query-transform-reverse-rows']");
    return;
  }
  if (action === "query-transform-remove-alternate-rows") {
    state.dataDialog.transformSteps.push({
      type: "removeAlternateRows",
      start: Math.max(1, Math.min(1048576, Math.round(Number(state.dataDialog.pendingAlternateStart) || 1))),
      remove: Math.max(1, Math.min(1048576, Math.round(Number(state.dataDialog.pendingAlternateRemove) || 1))),
      keep: Math.max(1, Math.min(1048576, Math.round(Number(state.dataDialog.pendingAlternateKeep) || 1))),
    });
    renderQueryTransformEditorDialog("[data-data-field='queryAlternateStart']");
    return;
  }
  if (action === "query-transform-add-custom-column") {
    const name = String(state.dataDialog.pendingCustomColumnName || "").trim();
    if (!name) {
      showMessageBox("列名を入力してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({
      type: "addCustomColumn",
      name,
      formula: String(state.dataDialog.pendingCustomFormula || ""),
    });
    state.dataDialog.selectedColumn = name;
    renderQueryTransformEditorDialog("[data-data-field='queryCustomFormula']");
    return;
  }
  if (action === "query-transform-add-conditional-column") {
    const name = String(state.dataDialog.pendingConditionName || "").trim();
    if (!name || !state.dataDialog.pendingConditionColumn) {
      showMessageBox("条件列の列名と判定列を入力してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({
      type: "addConditionalColumn",
      name,
      columnName: state.dataDialog.pendingConditionColumn,
      operator: state.dataDialog.pendingConditionOperator || "equals",
      value: String(state.dataDialog.pendingConditionValue || ""),
      trueValue: String(state.dataDialog.pendingConditionTrueValue || ""),
      falseValue: String(state.dataDialog.pendingConditionFalseValue || ""),
    });
    state.dataDialog.selectedColumn = name;
    renderQueryTransformEditorDialog("[data-data-field='queryConditionName']");
    return;
  }
  if (action === "query-transform-group-by") {
    const checkedGroupColumns = [...(dialog?.querySelectorAll("[data-query-transform-column]:checked") || [])]
      .map((input) => input.value)
      .filter(Boolean);
    const groupColumns = checkedGroupColumns.length ? checkedGroupColumns : [state.dataDialog.pendingGroupByColumn].filter(Boolean);
    if (!groupColumns.length) {
      showMessageBox("グループ化する列を選択してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({
      type: "groupBy",
      groupColumn: groupColumns[0],
      groupColumns,
      aggregateColumn: state.dataDialog.pendingAggregateColumn || groupColumns[0],
      aggregateFunction: state.dataDialog.pendingAggregateFunction || "sum",
      aggregateName: String(state.dataDialog.pendingAggregateName || "").trim() || queryTransformAggregateLabel(state.dataDialog.pendingAggregateFunction || "sum"),
    });
    renderQueryTransformEditorDialog("[data-data-field='queryGroupByColumn']");
    return;
  }
  if (action === "query-transform-apply-filter") {
    const text = String(dialog.querySelector("[data-data-field='queryFilterText']")?.value || "").trim();
    const operator = dialog.querySelector("[data-data-field='queryFilterOperator']")?.value || "equals";
    if (!text && operator !== "blank" && operator !== "notBlank") {
      showMessageBox("フィルター条件を入力してください。", { title: "Cht WebSheet" });
      return;
    }
    state.dataDialog.transformSteps.push({
      type: "filterRows",
      columnName: dialog.querySelector("[data-data-field='queryFilterColumn']")?.value || "__all__",
      operator,
      text,
    });
    state.dataDialog.pendingFilterText = text;
    renderQueryTransformEditorDialog("[data-data-field='queryFilterText']");
    return;
  }
  if (action === "query-transform-clear-filter") {
    state.dataDialog.transformSteps = (state.dataDialog.transformSteps || []).filter((step) => step.type !== "filterRows");
    state.dataDialog.pendingFilterText = "";
    renderQueryTransformEditorDialog("[data-data-field='queryFilterText']");
    return;
  }
  if (action === "query-transform-delete-step") {
    const index = Number(event.currentTarget.dataset.queryTransformStepIndex);
    if (Number.isInteger(index)) state.dataDialog.transformSteps.splice(index, 1);
    renderQueryTransformEditorDialog();
    return;
  }
  if (action === "query-transform-load") {
    loadQueryTransformEditorToSelection();
    return;
  }
  if (action === "query-transform-load-to") {
    registerQueryTransformConnectionAndOpenLoadTo();
  }
}



function selectedQueryTransformActionColumns(dialog) {
  const checked = [...(dialog?.querySelectorAll("[data-query-transform-column]:checked") || [])]
    .map((input) => input.value)
    .filter(Boolean);
  return checked.length ? checked : [state.dataDialog?.selectedColumn].filter(Boolean);
}



function syncQueryTransformEditorInputs(dialog) {
  const descriptor = state.dataDialog;
  if (!descriptor || descriptor.type !== "query-transform-editor") return;
  descriptor.queryName = String(dialog?.querySelector("[data-data-field='queryName']")?.value || descriptor.queryName || queryTransformName(descriptor)).trim();
  descriptor.pendingFilterText = String(dialog?.querySelector("[data-data-field='queryFilterText']")?.value || "");
  descriptor.selectedColumn = String(dialog?.querySelector("[data-data-field='queryActionColumn']")?.value || descriptor.selectedColumn || "");
  descriptor.pendingRenameTo = String(dialog?.querySelector("[data-data-field='queryRenameTo']")?.value || descriptor.pendingRenameTo || descriptor.selectedColumn || "");
  descriptor.selectedDataType = String(dialog?.querySelector("[data-data-field='queryDataType']")?.value || descriptor.selectedDataType || "text");
  descriptor.pendingRowCount = Math.max(1, Math.min(1048576, Math.round(Number(dialog?.querySelector("[data-data-field='queryRowCount']")?.value) || Number(descriptor.pendingRowCount) || 10)));
  descriptor.pendingRangeStart = Math.max(1, Math.min(1048576, Math.round(Number(dialog?.querySelector("[data-data-field='queryRangeStart']")?.value) || Number(descriptor.pendingRangeStart) || 1)));
  descriptor.pendingRangeCount = Math.max(1, Math.min(1048576, Math.round(Number(dialog?.querySelector("[data-data-field='queryRangeCount']")?.value) || Number(descriptor.pendingRangeCount) || descriptor.pendingRowCount || 10)));
  descriptor.pendingIndexStart = Math.max(0, Math.min(1048576, Math.round(Number(dialog?.querySelector("[data-data-field='queryIndexStart']")?.value) || Number(descriptor.pendingIndexStart) || 0)));
  descriptor.pendingReplaceFrom = String(dialog?.querySelector("[data-data-field='queryReplaceFrom']")?.value || descriptor.pendingReplaceFrom || "");
  descriptor.pendingReplaceTo = String(dialog?.querySelector("[data-data-field='queryReplaceTo']")?.value || descriptor.pendingReplaceTo || "");
  descriptor.pendingSplitDelimiter = String(dialog?.querySelector("[data-data-field='querySplitDelimiter']")?.value || descriptor.pendingSplitDelimiter || "");
  descriptor.pendingSplitPositions = String(dialog?.querySelector("[data-data-field='querySplitPositions']")?.value || descriptor.pendingSplitPositions || "");
  descriptor.pendingExtractDelimiter = String(dialog?.querySelector("[data-data-field='queryExtractDelimiter']")?.value || descriptor.pendingExtractDelimiter || "");
  descriptor.pendingExtractEndDelimiter = String(dialog?.querySelector("[data-data-field='queryExtractEndDelimiter']")?.value || descriptor.pendingExtractEndDelimiter || "");
  descriptor.pendingExtractCount = Math.max(1, Math.min(32767, Math.round(Number(dialog?.querySelector("[data-data-field='queryExtractCount']")?.value) || Number(descriptor.pendingExtractCount) || 3)));
  descriptor.pendingExtractStart = Math.max(0, Math.min(32767, Math.round(Number(dialog?.querySelector("[data-data-field='queryExtractStart']")?.value) || Number(descriptor.pendingExtractStart) || 0)));
  descriptor.pendingPrefixText = String(dialog?.querySelector("[data-data-field='queryPrefixText']")?.value || descriptor.pendingPrefixText || "");
  descriptor.pendingSuffixText = String(dialog?.querySelector("[data-data-field='querySuffixText']")?.value || descriptor.pendingSuffixText || "");
  descriptor.pendingCharacterSet = String(dialog?.querySelector("[data-data-field='queryCharacterSet']")?.value || descriptor.pendingCharacterSet || "");
  descriptor.pendingMergeDelimiter = String(dialog?.querySelector("[data-data-field='queryMergeDelimiter']")?.value || descriptor.pendingMergeDelimiter || "");
  descriptor.pendingMergeName = String(dialog?.querySelector("[data-data-field='queryMergeName']")?.value || descriptor.pendingMergeName || "結合済み");
  descriptor.pendingPivotColumn = String(dialog?.querySelector("[data-data-field='queryPivotColumn']")?.value || descriptor.pendingPivotColumn || descriptor.selectedColumn || "");
  descriptor.pendingPivotValueColumn = String(dialog?.querySelector("[data-data-field='queryPivotValueColumn']")?.value || descriptor.pendingPivotValueColumn || descriptor.selectedColumn || "");
  descriptor.pendingPivotAggregateFunction = String(dialog?.querySelector("[data-data-field='queryPivotAggregateFunction']")?.value || descriptor.pendingPivotAggregateFunction || "sum");
  descriptor.pendingAlternateStart = Math.max(1, Math.min(1048576, Math.round(Number(dialog?.querySelector("[data-data-field='queryAlternateStart']")?.value) || Number(descriptor.pendingAlternateStart) || 2)));
  descriptor.pendingAlternateRemove = Math.max(1, Math.min(1048576, Math.round(Number(dialog?.querySelector("[data-data-field='queryAlternateRemove']")?.value) || Number(descriptor.pendingAlternateRemove) || 1)));
  descriptor.pendingAlternateKeep = Math.max(1, Math.min(1048576, Math.round(Number(dialog?.querySelector("[data-data-field='queryAlternateKeep']")?.value) || Number(descriptor.pendingAlternateKeep) || 1)));
  descriptor.pendingNumberOperand = Number(dialog?.querySelector("[data-data-field='queryNumberOperand']")?.value ?? descriptor.pendingNumberOperand ?? 1);
  descriptor.pendingErrorReplacement = String(dialog?.querySelector("[data-data-field='queryErrorReplacement']")?.value || descriptor.pendingErrorReplacement || "");
  descriptor.pendingCustomColumnName = String(dialog?.querySelector("[data-data-field='queryCustomColumnName']")?.value || descriptor.pendingCustomColumnName || "カスタム");
  descriptor.pendingCustomFormula = String(dialog?.querySelector("[data-data-field='queryCustomFormula']")?.value || descriptor.pendingCustomFormula || "");
  descriptor.pendingGroupByColumn = String(dialog?.querySelector("[data-data-field='queryGroupByColumn']")?.value || descriptor.pendingGroupByColumn || descriptor.selectedColumn || "");
  descriptor.pendingAggregateColumn = String(dialog?.querySelector("[data-data-field='queryAggregateColumn']")?.value || descriptor.pendingAggregateColumn || descriptor.selectedColumn || "");
  descriptor.pendingAggregateFunction = String(dialog?.querySelector("[data-data-field='queryAggregateFunction']")?.value || descriptor.pendingAggregateFunction || "sum");
  descriptor.pendingAggregateName = String(dialog?.querySelector("[data-data-field='queryAggregateName']")?.value || descriptor.pendingAggregateName || "集計");
  descriptor.pendingConditionName = String(dialog?.querySelector("[data-data-field='queryConditionName']")?.value || descriptor.pendingConditionName || "条件");
  descriptor.pendingConditionColumn = String(dialog?.querySelector("[data-data-field='queryConditionColumn']")?.value || descriptor.pendingConditionColumn || descriptor.selectedColumn || "");
  descriptor.pendingConditionOperator = String(dialog?.querySelector("[data-data-field='queryConditionOperator']")?.value || descriptor.pendingConditionOperator || "equals");
  descriptor.pendingConditionValue = String(dialog?.querySelector("[data-data-field='queryConditionValue']")?.value || descriptor.pendingConditionValue || "");
  descriptor.pendingConditionTrueValue = String(dialog?.querySelector("[data-data-field='queryConditionTrueValue']")?.value || descriptor.pendingConditionTrueValue || "");
  descriptor.pendingConditionFalseValue = String(dialog?.querySelector("[data-data-field='queryConditionFalseValue']")?.value || descriptor.pendingConditionFalseValue || "");
  descriptor.pendingFilterOperator = String(dialog?.querySelector("[data-data-field='queryFilterOperator']")?.value || descriptor.pendingFilterOperator || "equals");
}



function queryTransformName(descriptor) {
  return descriptor?.queryName || String(descriptor?.sourceName || "Query").replace(/\.[^.]+$/, "") || "Query";
}



function queryTransformHasStep(descriptor, type) {
  return (descriptor?.transformSteps || []).some((step) => step.type === type);
}



function queryTransformResultForDescriptor(descriptor) {
  const sourceRows = queryTransformSourceRows(descriptor);
  return applyQueryTransformSteps(sourceRows, descriptor?.transformSteps || []);
}



function queryTransformSourceRows(descriptor) {
  if (!descriptor) return [];
  if (descriptor.sourceType === "blankQuery") {
    return [];
  }
  if (descriptor.sourceType === "queryReference") {
    return queryReferenceSourceRows(descriptor);
  }
  if (descriptor.sourceType === "queryAppend") {
    return queryAppendSourceRows(descriptor);
  }
  if (descriptor.sourceType === "queryMerge") {
    return queryMergeSourceRows(descriptor);
  }
  if (Array.isArray(descriptor.sourceRows)) {
    return descriptor.sourceRows.map((row) => (row || []).map((value) => String(value ?? "")));
  }
  if (descriptor.sourceType === "textCsv") {
    const delimiter = descriptor.delimiter || detectDelimitedTextDelimiter(descriptor.sourceText || "", descriptor.sourceName || "CSV");
    return parseDelimitedRows(descriptor.sourceText || "", delimiter)
      .filter((row, index, list) => index < list.length - 1 || row.some((value) => value !== ""))
      .map((row) => row.map((value) => String(value ?? "")));
  }
  if (descriptor.sourceType === "web") {
    return webImportRowsFromSource(
      descriptor.sourceText || "",
      descriptor.contentType || "",
      descriptor.sourceUrl || descriptor.sourceName || "Web",
      descriptor.sourceItemIndex ?? 0,
    )
      .map((row) => row.map((value) => String(value ?? "")));
  }
  if (descriptor.sourceType === "json" || descriptor.sourceType === "xml") {
    return webImportRowsFromSource(
      descriptor.sourceText || "",
      descriptor.contentType || (descriptor.sourceType === "xml" ? "application/xml" : "application/json"),
      descriptor.sourceUrl || descriptor.sourceName || descriptor.sourceType.toUpperCase(),
      descriptor.sourceItemIndex ?? 0,
    )
      .map((row) => row.map((value) => String(value ?? "")));
  }
  if (descriptor.sourceType === "odata") {
    return odataImportRowsFromSource(
      descriptor.sourceText || "",
      descriptor.contentType || "",
      descriptor.sourceUrl || descriptor.sourceName || "OData",
      descriptor.sourceItemIndex ?? 0,
    )
      .map((row) => row.map((value) => String(value ?? "")));
  }
  if (descriptor.sourceType === "tableRange") {
    const sheet = state.model?.sheets?.[descriptor.sourceSheetIndex] || state.model?.sheets?.find((candidate) => candidate.name === descriptor.sourceSheetName);
    if (sheet && descriptor.sourceRange) return sheetDisplayRowsForRange(sheet, descriptor.sourceRange);
  }
  return [];
}



function applyQueryTransformSteps(sourceRows, steps = []) {
  const sourceWidth = sourceRows.length ? Math.max(...sourceRows.map((row) => row.length), 1) : 0;
  let columns = queryTransformDefaultColumns(sourceWidth);
  let columnTypes = Array.from({ length: sourceWidth }, () => "");
  let rows = sourceRows.map((row) => padQueryTransformRow(row, sourceWidth));
  steps.forEach((step) => {
    if (step?.type === "promoteHeaders" && rows.length) {
      columns = uniqueQueryTransformColumns(rows[0].map((value, index) => String(value || "").trim() || `Column${index + 1}`));
      columnTypes = Array.from({ length: columns.length }, () => "");
      rows = rows.slice(1).map((row) => padQueryTransformRow(row, columns.length));
      return;
    }
    if (step?.type === "demoteHeaders") {
      rows = [columns, ...rows].map((row) => padQueryTransformRow(row, columns.length));
      columns = queryTransformDefaultColumns(columns.length);
      columnTypes = Array.from({ length: columns.length }, () => "");
      return;
    }
    if (step?.type === "transpose") {
      const matrix = [columns, ...rows];
      const height = matrix.length;
      const width = matrix.length ? Math.max(...matrix.map((row) => row.length), 1) : 0;
      const padded = matrix.map((row) => padQueryTransformRow(row, width));
      const transposed = Array.from({ length: width }, (_item, colIndex) =>
        Array.from({ length: height }, (_cell, rowIndex) => padded[rowIndex]?.[colIndex] ?? ""),
      );
      columns = queryTransformDefaultColumns(height);
      rows = transposed.map((row) => padQueryTransformRow(row, columns.length));
      columnTypes = Array.from({ length: columns.length }, () => "");
      return;
    }
    if (step?.type === "removeColumns") {
      const removeSet = new Set((step.columns || []).map((column) => String(column)));
      const keptIndexes = columns.map((column, index) => ({ column, index })).filter(({ column }) => !removeSet.has(column));
      columns = keptIndexes.map(({ column }) => column);
      columnTypes = keptIndexes.map(({ index }) => columnTypes[index] || "");
      rows = rows.map((row) => keptIndexes.map(({ index }) => row[index] ?? ""));
      return;
    }
    if (step?.type === "keepColumns") {
      const keepSet = new Set((step.columns || []).map((column) => String(column)));
      const keptIndexes = columns.map((column, index) => ({ column, index })).filter(({ column }) => keepSet.has(column));
      columns = keptIndexes.map(({ column }) => column);
      columnTypes = keptIndexes.map(({ index }) => columnTypes[index] || "");
      rows = rows.map((row) => keptIndexes.map(({ index }) => row[index] ?? ""));
      return;
    }
    if (step?.type === "removeEmptyColumns") {
      const keptIndexes = columns
        .map((column, index) => ({ column, index }))
        .filter(({ index }) => rows.some((row) => String(row[index] ?? "").trim() !== ""));
      columns = keptIndexes.map(({ column }) => column);
      columnTypes = keptIndexes.map(({ index }) => columnTypes[index] || "");
      rows = rows.map((row) => keptIndexes.map(({ index }) => row[index] ?? ""));
      return;
    }
    if (step?.type === "countRows") {
      columns = [String(step.name || "行数")];
      rows = [[String(rows.length)]];
      columnTypes = ["whole"];
      return;
    }
    if (step?.type === "renameColumn") {
      const index = columns.indexOf(step.from);
      if (index >= 0) {
        columns[index] = String(step.to || step.from || columns[index]);
        columns = uniqueQueryTransformColumns(columns);
      }
      return;
    }
    if (step?.type === "changeType") {
      const indexes = queryTransformStepColumnIndexes(columns, step);
      if (!indexes.length) return;
      const dataType = queryTransformNormalizeDataType(step.dataType);
      const indexSet = new Set(indexes);
      rows = rows.map((row) => row.map((value, valueIndex) =>
        indexSet.has(valueIndex) ? queryTransformConvertDataTypeValue(value, dataType) : value,
      ));
      indexes.forEach((index) => {
        columnTypes[index] = dataType;
      });
      return;
    }
    if (step?.type === "detectTypes") {
      const indexes = queryTransformStepColumnIndexes(columns, step);
      const targetIndexes = indexes.length ? indexes : columns.map((_column, index) => index);
      const inferredTypes = new Map(targetIndexes.map((index) => [
        index,
        queryTransformInferDataType(rows.map((row) => row[index])),
      ]));
      rows = rows.map((row) => row.map((value, valueIndex) => {
        const inferredType = inferredTypes.get(valueIndex);
        return inferredType ? queryTransformConvertDataTypeValue(value, inferredType) : value;
      }));
      inferredTypes.forEach((dataType, index) => {
        columnTypes[index] = dataType;
      });
      return;
    }
    if (step?.type === "numberOperation") {
      const indexes = queryTransformStepColumnIndexes(columns, step);
      if (!indexes.length) return;
      const indexSet = new Set(indexes);
      rows = rows.map((row) => row.map((value, valueIndex) =>
        indexSet.has(valueIndex) ? queryTransformNumberOperationValue(value, step.operation, step.operand) : value,
      ));
      indexes.forEach((index) => {
        columnTypes[index] = "number";
      });
      return;
    }
    if (step?.type === "datePart") {
      const indexes = queryTransformStepColumnIndexes(columns, step);
      if (!indexes.length) return;
      const indexSet = new Set(indexes);
      rows = rows.map((row) => row.map((value, valueIndex) =>
        indexSet.has(valueIndex) ? queryTransformDatePartValue(value, step.part) : value,
      ));
      indexes.forEach((index) => {
        columnTypes[index] = queryTransformDatePartDataType(step.part);
      });
      return;
    }
    if (step?.type === "replaceErrors") {
      const indexes = queryTransformStepColumnIndexes(columns, step);
      if (!indexes.length) return;
      const indexSet = new Set(indexes);
      rows = rows.map((row) => row.map((value, valueIndex) =>
        indexSet.has(valueIndex) && queryTransformIsErrorValue(value) ? String(step.value ?? "") : value,
      ));
      return;
    }
    if (step?.type === "removeErrors") {
      const indexes = queryTransformStepColumnIndexes(columns, step);
      if (!indexes.length) return;
      rows = rows.filter((row) => !indexes.some((index) => queryTransformIsErrorValue(row[index])));
      return;
    }
    if (step?.type === "keepErrors") {
      const indexes = queryTransformStepColumnIndexes(columns, step);
      if (!indexes.length) return;
      rows = rows.filter((row) => indexes.some((index) => queryTransformIsErrorValue(row[index])));
      return;
    }
    if (step?.type === "sortRows") {
      const index = columns.indexOf(step.columnName);
      if (index < 0) return;
      const direction = step.direction === "desc" ? -1 : 1;
      const dataType = columnTypes[index] || "";
      rows = rows.slice().sort((left, right) => compareQueryTransformValues(left[index], right[index], dataType) * direction);
      return;
    }
    if (step?.type === "duplicateColumn") {
      const index = columns.indexOf(step.columnName);
      if (index < 0) return;
      const newName = uniqueQueryTransformColumns([...columns, String(step.newName || `${columns[index]} - コピー`)]).at(-1);
      columns = [...columns.slice(0, index + 1), newName, ...columns.slice(index + 1)];
      columnTypes = [...columnTypes.slice(0, index + 1), columnTypes[index] || "", ...columnTypes.slice(index + 1)];
      rows = rows.map((row) => [...row.slice(0, index + 1), row[index] ?? "", ...row.slice(index + 1)]);
      if (step.newName !== newName) step.newName = newName;
      return;
    }
    if (step?.type === "moveColumn") {
      const fromIndex = columns.indexOf(step.columnName);
      if (fromIndex < 0) return;
      const toIndex = queryTransformMoveTargetIndex(fromIndex, columns.length, step.position);
      if (toIndex === fromIndex) return;
      columns = moveQueryTransformArrayItem(columns, fromIndex, toIndex);
      columnTypes = moveQueryTransformArrayItem(columnTypes, fromIndex, toIndex);
      rows = rows.map((row) => moveQueryTransformArrayItem(row, fromIndex, toIndex));
      return;
    }
    if (step?.type === "removeDuplicates") {
      const selectedIndexes = (step.columns || [])
        .map((column) => columns.indexOf(column))
        .filter((index) => index >= 0);
      const indexes = selectedIndexes.length ? selectedIndexes : columns.map((_column, index) => index);
      const seen = new Set();
      rows = rows.filter((row) => {
        const key = indexes.map((index) => String(row[index] ?? "")).join("\u0001");
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
      return;
    }
    if (step?.type === "keepDuplicates") {
      const selectedIndexes = (step.columns || [])
        .map((column) => columns.indexOf(column))
        .filter((index) => index >= 0);
      const indexes = selectedIndexes.length ? selectedIndexes : columns.map((_column, index) => index);
      const counts = new Map();
      rows.forEach((row) => {
        const key = indexes.map((index) => String(row[index] ?? "")).join("\u0001");
        counts.set(key, (counts.get(key) || 0) + 1);
      });
      rows = rows.filter((row) => {
        const key = indexes.map((index) => String(row[index] ?? "")).join("\u0001");
        return (counts.get(key) || 0) > 1;
      });
      return;
    }
    if (step?.type === "removeBlankRows") {
      rows = rows.filter((row) => row.some((value) => String(value ?? "").trim() !== ""));
      return;
    }
    if (step?.type === "keepBlankRows") {
      rows = rows.filter((row) => row.every((value) => String(value ?? "").trim() === ""));
      return;
    }
    if (step?.type === "keepTopRows") {
      rows = rows.slice(0, Math.max(0, Number(step.count) || 0));
      return;
    }
    if (step?.type === "removeTopRows") {
      rows = rows.slice(Math.max(0, Number(step.count) || 0));
      return;
    }
    if (step?.type === "keepBottomRows") {
      const count = Math.max(0, Number(step.count) || 0);
      rows = count >= rows.length ? rows : rows.slice(Math.max(0, rows.length - count));
      return;
    }
    if (step?.type === "removeBottomRows") {
      const count = Math.max(0, Number(step.count) || 0);
      rows = count >= rows.length ? [] : rows.slice(0, Math.max(0, rows.length - count));
      return;
    }
    if (step?.type === "keepRangeRows") {
      const startIndex = Math.max(0, Math.round(Number(step.start) || 1) - 1);
      const count = Math.max(0, Math.round(Number(step.count) || 0));
      rows = rows.slice(startIndex, startIndex + count);
      return;
    }
    if (step?.type === "removeRangeRows") {
      const startIndex = Math.max(0, Math.round(Number(step.start) || 1) - 1);
      const count = Math.max(0, Math.round(Number(step.count) || 0));
      rows = rows.filter((_row, rowIndex) => rowIndex < startIndex || rowIndex >= startIndex + count);
      return;
    }
    if (step?.type === "addIndexColumn") {
      const nextColumns = uniqueQueryTransformColumns([...columns, String(step.name || "インデックス").trim() || "インデックス"]);
      const columnName = nextColumns[nextColumns.length - 1];
      const start = Math.max(0, Math.round(Number(step.start) || 0));
      rows = rows.map((row, rowIndex) => [...row, String(start + rowIndex)]);
      columns = nextColumns;
      columnTypes = [...columnTypes, "whole"];
      if (step.name !== columnName) step.name = columnName;
      return;
    }
    if (step?.type === "replaceValues") {
      const index = columns.indexOf(step.columnName);
      const from = String(step.from ?? "");
      if (index < 0 || !from) return;
      const to = String(step.to ?? "");
      rows = rows.map((row) => row.map((value, valueIndex) => valueIndex === index ? String(value ?? "").split(from).join(to) : value));
      return;
    }
    if (step?.type === "replaceBlank") {
      const index = columns.indexOf(step.columnName);
      if (index < 0) return;
      const replacement = String(step.value ?? "");
      rows = rows.map((row) => row.map((value, valueIndex) =>
        valueIndex === index && String(value ?? "").trim() === "" ? replacement : value,
      ));
      return;
    }
    if (step?.type === "transformTextCase") {
      const index = columns.indexOf(step.columnName);
      if (index < 0) return;
      rows = rows.map((row) => row.map((value, valueIndex) => valueIndex === index ? queryTransformTextCaseValue(value, step.mode) : value));
      return;
    }
    if (step?.type === "textLength") {
      const indexes = queryTransformStepColumnIndexes(columns, step);
      if (!indexes.length) return;
      const indexSet = new Set(indexes);
      rows = rows.map((row) => row.map((value, valueIndex) => indexSet.has(valueIndex) ? String(String(value ?? "").length) : value));
      indexes.forEach((index) => {
        columnTypes[index] = "whole";
      });
      return;
    }
    if (step?.type === "trim") {
      const index = columns.indexOf(step.columnName);
      if (index < 0) return;
      rows = rows.map((row) => row.map((value, valueIndex) => valueIndex === index ? String(value ?? "").trim() : value));
      return;
    }
    if (step?.type === "clean") {
      const index = columns.indexOf(step.columnName);
      if (index < 0) return;
      rows = rows.map((row) => row.map((value, valueIndex) => valueIndex === index ? String(value ?? "").replace(/[\u0000-\u001f\u007f]/g, "") : value));
      return;
    }
    if (step?.type === "fillDown" || step?.type === "fillUp") {
      const index = columns.indexOf(step.columnName);
      if (index < 0) return;
      const nextRows = rows.map((row) => [...row]);
      const fillIndexes = step.type === "fillDown"
        ? nextRows.map((_row, rowIndex) => rowIndex)
        : nextRows.map((_row, rowIndex) => rowIndex).reverse();
      let carry = "";
      fillIndexes.forEach((rowIndex) => {
        const value = String(nextRows[rowIndex]?.[index] ?? "");
        if (value.trim() === "" && carry !== "") nextRows[rowIndex][index] = carry;
        else if (value.trim() !== "") carry = value;
      });
      rows = nextRows;
      return;
    }
    if (step?.type === "extractFirstCharacters" || step?.type === "extractLastCharacters") {
      const index = columns.indexOf(step.columnName);
      const count = Math.max(1, Math.min(32767, Math.round(Number(step.count) || 1)));
      if (index < 0) return;
      rows = rows.map((row) => row.map((value, valueIndex) => {
        if (valueIndex !== index) return value;
        const text = String(value ?? "");
        return step.type === "extractFirstCharacters" ? text.slice(0, count) : text.slice(Math.max(0, text.length - count));
      }));
      return;
    }
    if (step?.type === "extractRangeCharacters") {
      const index = columns.indexOf(step.columnName);
      const start = Math.max(0, Math.min(32767, Math.round(Number(step.start) || 0)));
      const count = Math.max(1, Math.min(32767, Math.round(Number(step.count) || 1)));
      if (index < 0) return;
      rows = rows.map((row) => row.map((value, valueIndex) => {
        if (valueIndex !== index) return value;
        return String(value ?? "").slice(start, start + count);
      }));
      return;
    }
    if (step?.type === "extractBetweenDelimiters") {
      const index = columns.indexOf(step.columnName);
      const startDelimiter = String(step.startDelimiter ?? "");
      const endDelimiter = String(step.endDelimiter ?? "");
      if (index < 0 || !startDelimiter || !endDelimiter) return;
      rows = rows.map((row) => row.map((value, valueIndex) => {
        if (valueIndex !== index) return value;
        const text = String(value ?? "");
        const startIndex = text.indexOf(startDelimiter);
        if (startIndex < 0) return "";
        const bodyStart = startIndex + startDelimiter.length;
        const endIndex = text.indexOf(endDelimiter, bodyStart);
        return endIndex < 0 ? "" : text.slice(bodyStart, endIndex);
      }));
      return;
    }
    if (step?.type === "addPrefix" || step?.type === "addSuffix") {
      const index = columns.indexOf(step.columnName);
      if (index < 0) return;
      const addition = String(step.value ?? "");
      rows = rows.map((row) => row.map((value, valueIndex) => {
        if (valueIndex !== index) return value;
        const text = String(value ?? "");
        return step.type === "addPrefix" ? `${addition}${text}` : `${text}${addition}`;
      }));
      return;
    }
    if (step?.type === "keepCharacters" || step?.type === "removeCharacters") {
      const index = columns.indexOf(step.columnName);
      if (index < 0) return;
      rows = rows.map((row) => row.map((value, valueIndex) => valueIndex === index
        ? queryTransformCharacterSetValue(value, step.characters, step.type === "keepCharacters")
        : value));
      return;
    }
    if (step?.type === "splitColumnByDelimiter") {
      const index = columns.indexOf(step.columnName);
      const delimiter = String(step.delimiter ?? "");
      if (index < 0 || !delimiter) return;
      const splitValues = rows.map((row) => String(row[index] ?? "").split(delimiter));
      const partCount = Math.max(1, ...splitValues.map((parts) => parts.length));
      const baseName = columns[index] || `Column${index + 1}`;
      const insertedColumns = Array.from({ length: partCount }, (_item, partIndex) => partIndex === 0 ? baseName : `${baseName}.${partIndex + 1}`);
      columns = uniqueQueryTransformColumns([...columns.slice(0, index), ...insertedColumns, ...columns.slice(index + 1)]);
      columnTypes = [
        ...columnTypes.slice(0, index),
        ...Array.from({ length: partCount }, () => "text"),
        ...columnTypes.slice(index + 1),
      ];
      rows = rows.map((row, rowIndex) => [
        ...row.slice(0, index),
        ...padQueryTransformRow(splitValues[rowIndex], partCount),
        ...row.slice(index + 1),
      ]);
      return;
    }
    if (step?.type === "splitColumnByCharacterCount") {
      const index = columns.indexOf(step.columnName);
      const count = Math.max(1, Math.min(32767, Math.round(Number(step.count) || 1)));
      if (index < 0) return;
      const splitValues = rows.map((row) => splitTextByCharacterCount(row[index], count));
      const partCount = Math.max(1, ...splitValues.map((parts) => parts.length));
      const baseName = columns[index] || `Column${index + 1}`;
      const insertedColumns = Array.from({ length: partCount }, (_item, partIndex) => partIndex === 0 ? baseName : `${baseName}.${partIndex + 1}`);
      columns = uniqueQueryTransformColumns([...columns.slice(0, index), ...insertedColumns, ...columns.slice(index + 1)]);
      columnTypes = [
        ...columnTypes.slice(0, index),
        ...Array.from({ length: partCount }, () => "text"),
        ...columnTypes.slice(index + 1),
      ];
      rows = rows.map((row, rowIndex) => [
        ...row.slice(0, index),
        ...padQueryTransformRow(splitValues[rowIndex], partCount),
        ...row.slice(index + 1),
      ]);
      return;
    }
    if (step?.type === "splitColumnByPositions") {
      const index = columns.indexOf(step.columnName);
      const positions = [...new Set((step.positions || []).map((position) => Math.round(Number(position))).filter((position) => Number.isFinite(position) && position > 0))]
        .sort((left, right) => left - right);
      if (index < 0 || !positions.length) return;
      const splitValues = rows.map((row) => splitTextByFixedBreaks(row[index], positions));
      const partCount = Math.max(1, ...splitValues.map((parts) => parts.length));
      const baseName = columns[index] || `Column${index + 1}`;
      const insertedColumns = Array.from({ length: partCount }, (_item, partIndex) => partIndex === 0 ? baseName : `${baseName}.${partIndex + 1}`);
      columns = uniqueQueryTransformColumns([...columns.slice(0, index), ...insertedColumns, ...columns.slice(index + 1)]);
      columnTypes = [
        ...columnTypes.slice(0, index),
        ...Array.from({ length: partCount }, () => "text"),
        ...columnTypes.slice(index + 1),
      ];
      rows = rows.map((row, rowIndex) => [
        ...row.slice(0, index),
        ...padQueryTransformRow(splitValues[rowIndex], partCount),
        ...row.slice(index + 1),
      ]);
      return;
    }
    if (step?.type === "mergeColumns") {
      const selectedIndexes = (step.columns || [])
        .map((column) => columns.indexOf(column))
        .filter((index) => index >= 0)
        .sort((left, right) => left - right);
      if (selectedIndexes.length < 2) return;
      const selectedSet = new Set(selectedIndexes);
      const insertIndex = selectedIndexes[0];
      const mergedName = uniqueQueryTransformColumns([...columns, String(step.name || "結合済み").trim() || "結合済み"]).at(-1);
      const delimiter = String(step.delimiter ?? "");
      const nextColumns = [];
      const nextTypes = [];
      columns.forEach((column, index) => {
        if (index === insertIndex) {
          nextColumns.push(mergedName);
          nextTypes.push("text");
        }
        if (!selectedSet.has(index)) {
          nextColumns.push(column);
          nextTypes.push(columnTypes[index] || "");
        }
      });
      rows = rows.map((row) => {
        const merged = selectedIndexes.map((index) => String(row[index] ?? "")).join(delimiter);
        const nextRow = [];
        row.forEach((value, index) => {
          if (index === insertIndex) nextRow.push(merged);
          if (!selectedSet.has(index)) nextRow.push(value);
        });
        return nextRow;
      });
      columns = uniqueQueryTransformColumns(nextColumns);
      columnTypes = nextTypes;
      if (step.name !== mergedName) step.name = mergedName;
      return;
    }
    if (step?.type === "unpivotColumns") {
      const selectedIndexes = (step.columns || [])
        .map((column) => columns.indexOf(column))
        .filter((index) => index >= 0)
        .sort((left, right) => left - right);
      if (!selectedIndexes.length) return;
      const selectedSet = new Set(selectedIndexes);
      const fixedIndexes = columns.map((_column, index) => index).filter((index) => !selectedSet.has(index));
      const attributeName = String(step.attributeName || "属性").trim() || "属性";
      const valueName = String(step.valueName || "値").trim() || "値";
      rows = rows.flatMap((row) => selectedIndexes.map((index) => [
        ...fixedIndexes.map((fixedIndex) => row[fixedIndex] ?? ""),
        columns[index],
        row[index] ?? "",
      ]));
      columns = uniqueQueryTransformColumns([...fixedIndexes.map((index) => columns[index]), attributeName, valueName]);
      columnTypes = [...fixedIndexes.map((index) => columnTypes[index] || ""), "text", "any"];
      return;
    }
    if (step?.type === "pivotColumn") {
      const pivotIndex = columns.indexOf(step.columnName);
      const valueIndex = columns.indexOf(step.valueColumn);
      if (pivotIndex < 0 || valueIndex < 0 || pivotIndex === valueIndex) return;
      const fixedIndexes = columns
        .map((_column, index) => index)
        .filter((index) => index !== pivotIndex && index !== valueIndex);
      const pivotValues = [];
      rows.forEach((row) => {
        const value = String(row[pivotIndex] ?? "").trim() || "(空白)";
        if (!pivotValues.includes(value)) pivotValues.push(value);
      });
      const groups = new Map();
      rows.forEach((row) => {
        const fixedValues = fixedIndexes.map((index) => row[index] ?? "");
        const key = JSON.stringify(fixedValues);
        if (!groups.has(key)) {
          groups.set(key, {
            fixedValues,
            buckets: new Map(),
          });
        }
        const pivotName = String(row[pivotIndex] ?? "").trim() || "(空白)";
        const group = groups.get(key);
        if (!group.buckets.has(pivotName)) group.buckets.set(pivotName, []);
        group.buckets.get(pivotName).push(row);
      });
      const aggregateFunction = step.aggregateFunction || "sum";
      rows = [...groups.values()].map((group) => [
        ...group.fixedValues,
        ...pivotValues.map((pivotName) =>
          queryTransformAggregateValue(group.buckets.get(pivotName) || [], valueIndex, aggregateFunction, columnTypes[valueIndex] || ""),
        ),
      ]);
      columns = uniqueQueryTransformColumns([
        ...fixedIndexes.map((index) => columns[index]),
        ...pivotValues,
      ]);
      columnTypes = [
        ...fixedIndexes.map((index) => columnTypes[index] || ""),
        ...pivotValues.map(() => queryTransformAggregateDataType(aggregateFunction, columnTypes[valueIndex] || "")),
      ];
      return;
    }
    if (step?.type === "extractBeforeDelimiter" || step?.type === "extractAfterDelimiter") {
      const index = columns.indexOf(step.columnName);
      const delimiter = String(step.delimiter ?? "");
      if (index < 0 || !delimiter) return;
      rows = rows.map((row) => row.map((value, valueIndex) => {
        if (valueIndex !== index) return value;
        const text = String(value ?? "");
        const delimiterIndex = text.indexOf(delimiter);
        if (delimiterIndex < 0) return text;
        return step.type === "extractBeforeDelimiter"
          ? text.slice(0, delimiterIndex)
          : text.slice(delimiterIndex + delimiter.length);
      }));
      return;
    }
    if (step?.type === "addCustomColumn") {
      const nextColumns = uniqueQueryTransformColumns([...columns, String(step.name || "カスタム").trim() || "カスタム"]);
      const columnName = nextColumns[nextColumns.length - 1];
      rows = rows.map((row) => [...row, evaluateQueryCustomColumnFormula(step.formula, columns, row)]);
      columns = nextColumns;
      columnTypes = [...columnTypes, "any"];
      if (step.name !== columnName) step.name = columnName;
      return;
    }
    if (step?.type === "addConditionalColumn") {
      const conditionIndex = columns.indexOf(step.columnName);
      if (conditionIndex < 0) return;
      const nextColumns = uniqueQueryTransformColumns([...columns, String(step.name || "条件").trim() || "条件"]);
      const columnName = nextColumns[nextColumns.length - 1];
      rows = rows.map((row) => [
        ...row,
        queryTransformConditionMatches(row[conditionIndex], step.operator, step.value, columnTypes[conditionIndex] || "")
          ? String(step.trueValue ?? "")
          : String(step.falseValue ?? ""),
      ]);
      columns = nextColumns;
      columnTypes = [...columnTypes, "text"];
      if (step.name !== columnName) step.name = columnName;
      return;
    }
    if (step?.type === "reverseRows") {
      rows = rows.slice().reverse();
      return;
    }
    if (step?.type === "removeAlternateRows") {
      const start = Math.max(1, Math.round(Number(step.start) || 1));
      const removeCount = Math.max(1, Math.round(Number(step.remove) || 1));
      const keepCount = Math.max(1, Math.round(Number(step.keep) || 1));
      rows = rows.filter((_row, rowIndex) => {
        const oneBased = rowIndex + 1;
        if (oneBased < start) return true;
        const cycleIndex = (oneBased - start) % (removeCount + keepCount);
        return cycleIndex >= removeCount;
      });
      return;
    }
    if (step?.type === "groupBy") {
      const groupIndexes = queryTransformGroupColumns(step)
        .map((column) => columns.indexOf(column))
        .filter((index) => index >= 0);
      const aggregateIndex = columns.indexOf(step.aggregateColumn);
      if (!groupIndexes.length) return;
      const groups = new Map();
      rows.forEach((row) => {
        const keyValues = groupIndexes.map((index) => String(row[index] ?? ""));
        const key = JSON.stringify(keyValues);
        if (!groups.has(key)) groups.set(key, { keyValues, rows: [] });
        groups.get(key).rows.push(row);
      });
      const aggregateName = String(step.aggregateName || queryTransformAggregateLabel(step.aggregateFunction) || "集計");
      rows = [...groups.values()].map((group) => [
        ...group.keyValues,
        queryTransformAggregateValue(group.rows, aggregateIndex, step.aggregateFunction, columnTypes[aggregateIndex] || ""),
      ]);
      columns = uniqueQueryTransformColumns([...groupIndexes.map((index) => columns[index]), aggregateName]);
      columnTypes = [
        ...groupIndexes.map((index) => columnTypes[index] || ""),
        queryTransformAggregateDataType(step.aggregateFunction, columnTypes[aggregateIndex] || ""),
      ];
      return;
    }
    if (step?.type === "filterRows") {
      const text = String(step.text || "");
      const operator = step.operator || "equals";
      if (!text && operator !== "blank" && operator !== "notBlank") return;
      const columnIndex = step.columnName && step.columnName !== "__all__" ? columns.indexOf(step.columnName) : -1;
      rows = rows.filter((row) => {
        const values = columnIndex >= 0 ? [row[columnIndex]] : row;
        return values.some((value, valueIndex) => {
          const dataType = columnIndex >= 0 ? columnTypes[columnIndex] || "" : columnTypes[valueIndex] || "";
          return queryTransformConditionMatches(value, operator, text, dataType);
        });
      });
    }
  });
  const outputRows = columns.length ? [columns, ...rows] : [];
  return {
    columns,
    rows,
    outputRows,
    columnTypes,
  };
}

export {
  applyQueryTransformSteps,
  applyTextCsvImportDialog,
  cellMatchesFindQuery,
  chooseDelimitedTextForImport,
  chooseFolderDataForImport,
  chooseStructuredDataForImport,
  chooseWorkbookDataForImport,
  clearFindAllResults,
  closeFindDialog,
  closeFindDialogOnEscape,
  coerceDelimitedImportValue,
  decodeTextCsvBytes,
  delimiterLabel,
  detectTextCsvOriginFromBytes,
  findAllFromDialog,
  findCellsInScope,
  findCurrentOrNextCell,
  findDialogOptionsFromElement,
  findMatchingCells,
  findNextCell,
  findNextFromDialog,
  findRangeForCell,
  findResultButtonHtml,
  findSearchSheetIndexes,
  findSearchStartIndex,
  folderImportCombinedFilesSource,
  folderImportCombinedRowCount,
  folderImportDelimitedTextSources,
  folderImportDisplayName,
  folderImportFileExtension,
  folderImportFileListSource,
  folderImportSourcesFromFile,
  folderImportStructuredSources,
  folderImportWorkbookSources,
  handleFindDialogFieldChange,
  handleFindDialogKeydown,
  handleQueryTransformEditorAction,
  handleTextCsvPreviewChange,
  importDelimitedText,
  importDelimitedTextIntoSelection,
  matchingCellRanges,
  normalizeFindScope,
  openFolderImportFromFiles,
  openStructuredImportFromText,
  openTextCsvImportDialog,
  openTextCsvTransformEditor,
  openWorkbookImportFromFile,
  queryTransformAggregateLabel,
  queryTransformAggregateOptions,
  queryTransformColumnChecklistHtml,
  queryTransformConditionOperatorLabel,
  queryTransformConditionOperatorOptions,
  queryTransformDataTypeLabel,
  queryTransformDataTypeOptions,
  queryTransformDatePartLabel,
  queryTransformEditorHtml,
  queryTransformGroupColumns,
  queryTransformHasStep,
  queryTransformMoveLabel,
  queryTransformName,
  queryTransformNumberOperationLabel,
  queryTransformPreviewHtml,
  queryTransformResultForDescriptor,
  queryTransformSourceRows,
  queryTransformStepColumnLabel,
  queryTransformStepLabel,
  queryTransformStepsHtml,
  queryTransformTextCaseLabel,
  renderFindAllResults,
  renderQueryTransformEditorDialog,
  repeatLastFindShortcut,
  replaceCellText,
  replaceFromDialog,
  selectCellsByPredicate,
  selectFindResultFromDialog,
  selectRangesForHomeCommand,
  selectSingleRange,
  selectedQueryTransformActionColumns,
  stripTextBom,
  syncQueryTransformEditorInputs,
  textCsvBytesFromSource,
  textCsvCurrentDetectedDelimiter,
  textCsvDelimiterFromDialog,
  textCsvDelimiterOptions,
  textCsvDialogOrigin,
  textCsvDialogText,
  textCsvEncodingLabel,
  textCsvPreviewHtml,
  textCsvRowCount,
  textCsvRowsFromDialog,
  updateTextCsvDialogSummary,
  workbookImportCellText,
  workbookImportRowsFromWorksheet,
};
