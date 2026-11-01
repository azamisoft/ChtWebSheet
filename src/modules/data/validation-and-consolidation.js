// Feature module; shared bindings are imported explicitly.
import {
  DATA_SUMMARY_FUNCTIONS,
  createIcons,
  icons,
  state,
} from "../../app-core.js";
import {
  applySubtotalGroupOutline,
  clearSubtotalOutlineMetadata,
  clearSubtotalPageBreaks,
  closeWhatIfMenu,
  insertSubtotalGrandTotalRow,
  insertSubtotalRowsForSheetGroup,
  removeSubtotalRowsWithoutHistory,
} from "./analysis-and-scenarios.js";
import {
  formatQueryTimestamp,
} from "./connections-and-queries.js";
import {
  autoFilterDateSerial,
  dataRangeFieldHtml,
  normalizeAdvancedFilterHeader,
  workbookRangeFromFormulaReference,
} from "./filter-and-text-tools.js";
import {
  comparableDateSerial,
  currentRegionFromPoint,
  inferHeaderRow,
  rangeHasUserData,
  showDataOverwriteConfirmation,
} from "./sort-and-autofilter.js";
import {
  closeDataDialog,
} from "../drawing/charts-and-insertion.js";
import {
  mergesIntersectingRange,
  rangeContains,
} from "../editing/autofill.js";
import {
  closeRibbonBorderMenu,
  closeRibbonColorPalette,
  closeRibbonShapeMenu,
  closeRibbonUnderlineMenu,
  closeShapeFormatMenu,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  firstPointForActiveSelection,
} from "../editing/keyboard-and-paste.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  affectedSheetIndexesForGroupedRanges,
  cellPointFromElement,
  commitCellValue,
  getCellRawInput,
  getDisplayForCell,
  isCellInRange,
  isSingleCellRange,
  materializeSelectionRangeForSheet,
  materializedRangesForSheet,
  normalizeRange,
  rangeToLabel,
  restoreGroupedSheetsAfterFailure,
  selectionRangeFromPoints,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  rangesEqual,
  scrollRangeIntoView,
  selectCellPoint,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneCellModel,
  cloneJsonLikeValue,
  defaultRowModel,
  deleteEntireRows,
  insertEntireRows,
  internCellsInRange,
  internCellsInRanges,
  pushHistory,
  restoreSheetFromHistory,
  storeCellModelForSheet,
  writeClonedCellContent,
} from "../editing/structure-and-history.js";
import {
  ensureCellModel,
  forEachCellInRange,
  rangeIntersects,
  rangesIntersect,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  closeCellContextMenu,
  positionPopupMenu,
} from "../ui/annotations-and-cell-format.js";
import {
  accessKeyLabelHtml,
  cellKey,
  columnName,
  columnNumber,
  dateToExcelSerial,
  decodeRange,
  escapeAttr,
  escapeHtml,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  consumeCommandMenuKey,
  enhanceCommandMenuAccessKeys,
  validationComparisonKey,
} from "../ui/printing-and-home-tools.js";
import {
  closeFileSideMenu,
  commitDataOperation,
  commitDataOperationForSheet,
  ensureSheetSize,
  parseCellKey,
  quoteSheetNameForFormula,
  selectedDataRange,
} from "../ui/ribbon-and-links.js";
import {
  closeSheetTabContextMenu,
  ensureActiveSheetFormatCellsAllowed,
  ensureEditableRangeForProtection,
  ensureSheetProtectionPermission,
  rangeHasLockedCellsForProtection,
} from "../ui/sheet-protection-and-translation.js";
import {
  buildFormulaEngine,
  formulaEngineContentForCurrentModel,
  formulaEngineSheetIndex,
  formulaEngineValueForCell,
  renderWorkbook,
  tableColumnNamesForStructuredRef,
  translateFormulaForCurrentFormulaEngine,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  activeSheet,
  getSheetId,
  visibleGridRange,
} from "../workbook/saving-and-export.js";
import {
  normalizeOutlineSettings,
  parseValidationOptions,
  rangeIntersectsWorksheetTable,
  validationListOptions,
  validationOptionsFromReference,
  validationReferenceFromFormula,
} from "../workbook/tables-and-validation.js";


function removeDuplicatesFromSelection() {
  const context = selectedDataRange("重複の削除");
  if (!context) return;
  if (!isSingleCellRange(context.range)) {
    openRemoveDuplicatesDialog(context.range);
    return;
  }
  const expandedRange = dataOperationRangeFromCurrentRegion(context.sheet, context.range);
  const openCurrent = () => openRemoveDuplicatesDialog(context.range);
  const openExpanded = () => openRemoveDuplicatesDialog(expandedRange);
  if (shouldShowDataExpansionWarning(context.sheet, context.range, expandedRange)) {
    showMessageBox("選択範囲の周囲にデータが見つかりました。重複の削除の対象をどうしますか?", {
      title: "Cht WebSheet",
      tone: "warning",
      buttons: [
        { label: "選択範囲を拡張する(E)", result: "expand", callback: openExpanded },
        { label: "現在選択されている範囲で続行する(C)", result: "current", callback: openCurrent },
        { label: "キャンセル", result: "cancel" },
      ],
      cancelResult: "cancel",
    });
    return;
  }
  openExpanded();
}



function dataOperationRangeFromCurrentRegion(sheet, range) {
  if (!sheet || !range) return range;
  const point = state.selected?.sheetIndex === state.activeSheetIndex && isCellInRange(range, state.selected.row, state.selected.col)
    ? state.selected
    : { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
  const region = currentRegionFromPoint(sheet, point);
  if (!region) return range;
  if (isSingleCellRange(range)) return region;
  if (rangeContains(region, range) && !rangesEqual(region, range)) return region;
  return range;
}



function shouldShowDataExpansionWarning(sheet, selectedRange, expandedRange) {
  if (!sheet || !selectedRange || !expandedRange) return false;
  if (isSingleCellRange(selectedRange) || rangesEqual(selectedRange, expandedRange)) return false;
  if (!rangeContains(expandedRange, selectedRange)) return false;
  const selectedCells = (selectedRange.bottom - selectedRange.top + 1) * (selectedRange.right - selectedRange.left + 1);
  const expandedCells = (expandedRange.bottom - expandedRange.top + 1) * (expandedRange.right - expandedRange.left + 1);
  return expandedCells > selectedCells;
}



function openRemoveDuplicatesDialog(initialRange = null) {
  const context = initialRange
    ? { sheet: activeSheet(), range: initialRange }
    : selectedDataRange("重複の削除", { expandCurrentRegion: true });
  if (!context) return;
  const { sheet, range } = context;
  if (range.bottom <= range.top) {
    showMessageBox("重複を削除するデータがありません。", { title: "Cht WebSheet" });
    return;
  }
  const hasHeader = inferHeaderRow(sheet, range);
  const sheetId = getSheetId(sheet.name);
  const columns = [];
  for (let col = range.left; col <= range.right; col += 1) {
    const header = hasHeader ? String(getDisplayForCell(sheet, sheetId, range.top, col) || columnName(col)) : columnName(col);
    columns.push({ col, label: header });
  }
  state.dataDialog = { type: "remove-duplicates", range };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog remove-duplicates-dialog" role="dialog" aria-modal="true" aria-labelledby="removeDuplicatesDialogTitle" data-data-dialog="remove-duplicates">
      <div class="excel-mini-titlebar" id="removeDuplicatesDialogTitle">重複の削除</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-note">重複を削除するには、重複している値を含む列を選択します。</div>
        <label class="data-dialog-check"><input type="checkbox" data-data-field="header" data-remove-duplicates-sync ${hasHeader ? "checked" : ""} />先頭行をデータの見出しとして使用する</label>
        <div class="data-dialog-inline-actions">
          <button class="excel-mini-button" type="button" data-data-dialog-action="select-all-columns">すべて選択</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="unselect-all-columns">すべて選択解除</button>
        </div>
        <div class="remove-duplicates-list" aria-label="列">
          ${columns
            .map((item) => `<label><input type="checkbox" data-data-duplicate-col="${item.col}" checked /><span data-duplicate-col-label="${item.col}">${escapeHtml(item.label)}</span></label>`)
            .join("")}
        </div>
        <div class="data-dialog-summary remove-duplicates-summary" data-remove-duplicates-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncRemoveDuplicatesDialog(overlay.querySelector("[data-data-dialog='remove-duplicates']"));
}



function handleRemoveDuplicatesDialogSync(event) {
  syncRemoveDuplicatesDialog(event.currentTarget.closest("[data-data-dialog='remove-duplicates']"));
}



function syncRemoveDuplicatesDialog(dialog) {
  const range = state.dataDialog?.type === "remove-duplicates" ? state.dataDialog.range : null;
  const sheet = activeSheet();
  if (!dialog || !range || !sheet) return;
  const hasHeader = Boolean(dialog.querySelector("[data-data-field='header']")?.checked);
  const sheetId = getSheetId(sheet.name);
  dialog.querySelectorAll("[data-duplicate-col-label]").forEach((label) => {
    const col = Number(label.dataset.duplicateColLabel);
    if (!Number.isFinite(col)) return;
    label.textContent = hasHeader ? String(getDisplayForCell(sheet, sheetId, range.top, col) || columnName(col)) : columnName(col);
  });
  const selectedCols = selectedRemoveDuplicateColumnsFromDialog(dialog);
  const allColumnInputs = Array.from(dialog.querySelectorAll("[data-data-duplicate-col]"));
  const selectedCount = selectedCols.length;
  const preview = removeDuplicatesPreview(sheet, range, selectedCols, hasHeader);
  const setDisabled = (action, disabled) => {
    const button = dialog.querySelector(`[data-data-dialog-action='${action}']`);
    if (button) button.disabled = Boolean(disabled);
  };
  setDisabled("select-all-columns", allColumnInputs.length > 0 && selectedCount === allColumnInputs.length);
  setDisabled("unselect-all-columns", selectedCount === 0);
  setDisabled("ok", preview.blocked);
  const summary = dialog.querySelector("[data-remove-duplicates-summary]");
  if (summary) {
    const selectedText = selectedCount ? `選択列: ${selectedCount} / ${allColumnInputs.length}。` : "";
    summary.textContent = `${selectedText}${preview.message}`;
    summary.classList.toggle("is-warning", preview.warning || preview.blocked);
  }
}



function applyRemoveDuplicatesDialog(dialog) {
  const range = state.dataDialog?.range;
  if (!range) return;
  if (!ensureEditableRangeForProtection(activeSheet(), range, "削除")) return;
  const selectedCols = selectedRemoveDuplicateColumnsFromDialog(dialog);
  if (!selectedCols.length) {
    showMessageBox("列を1つ以上選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const hasHeader = Boolean(dialog.querySelector("[data-data-field='header']")?.checked);
  const preview = removeDuplicatesPreview(activeSheet(), range, selectedCols, hasHeader);
  if (preview.blocked) {
    showMessageBox(preview.message, { title: "Cht WebSheet" });
    syncRemoveDuplicatesDialog(dialog);
    return;
  }
  closeDataDialog();
  if (removeDuplicateRowsForSheetGroup(range, selectedCols, hasHeader)) return;
  removeDuplicateRows(range, selectedCols, hasHeader);
}



function selectedRemoveDuplicateColumnsFromDialog(dialog) {
  return Array.from(dialog?.querySelectorAll("[data-data-duplicate-col]:checked") || [])
    .map((input) => Number(input.dataset.dataDuplicateCol))
    .filter((col) => Number.isFinite(col));
}



function removeDuplicatesPreview(sheet, range, selectedCols, hasHeader) {
  if (!selectedCols.length) return { warning: true, message: "比較する列を1つ以上選択してください。" };
  const mergeConflict = removeDuplicatesMergeConflict(sheet, range, hasHeader);
  if (mergeConflict) return { warning: true, blocked: true, message: mergeConflict };
  const startRow = hasHeader ? range.top + 1 : range.top;
  const total = Math.max(0, range.bottom - startRow + 1);
  if (!total) return { warning: true, message: "重複を削除するデータがありません。" };
  const sheetId = getSheetId(sheet.name);
  const seen = new Set();
  let removed = 0;
  for (let row = startRow; row <= range.bottom; row += 1) {
    const key = selectedCols.map((col) => removeDuplicatesComparableValue(sheet, sheetId, row, col)).join("\u0001");
    if (seen.has(key)) removed += 1;
    else seen.add(key);
  }
  return {
    warning: false,
    message: removed
      ? `${total} 行中 ${removed} 行が重複として削除され、${total - removed} 行が残ります。`
      : `${total} 行に重複は見つかりません。`,
  };
}



function removeDuplicateRows(range, selectedCols, hasHeader) {
  const sheet = activeSheet();
  const mergeConflict = removeDuplicatesMergeConflict(sheet, range, hasHeader);
  if (mergeConflict) {
    showMessageBox(mergeConflict, { title: "Cht WebSheet" });
    return;
  }
  const preview = collectRemoveDuplicateRows(sheet, range, selectedCols, hasHeader);
  if (!preview.removed) {
    showMessageBox("重複する値は見つかりませんでした。", { title: "Cht WebSheet" });
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  applyCollectedRemoveDuplicateRows(sheet, range, preview);
  commitDataOperation(sheet, sheetBefore, range, `${rangeToLabel(range)} 重複の削除`, `${preview.removed}個の重複値を削除しました。`);
  showMessageBox(`重複する値が ${preview.removed} 個見つかり、削除されました。\n一意の値が ${preview.uniqueRows.length} 個残っています。`, { title: "Cht WebSheet" });
}



function collectRemoveDuplicateRows(sheet, range, selectedCols, hasHeader) {
  const startRow = hasHeader ? range.top + 1 : range.top;
  const sheetId = getSheetId(sheet.name);
  const seen = new Set();
  const uniqueRows = [];
  let removed = 0;
  for (let row = startRow; row <= range.bottom; row += 1) {
    const cells = [];
    for (let col = range.left; col <= range.right; col += 1) {
      cells.push(cloneCellModel(sheet.cells[cellKey(row, col)]));
    }
    const signature = selectedCols.map((col) => removeDuplicatesComparableValue(sheet, sheetId, row, col));
    const key = signature.join("\u0001");
    if (seen.has(key)) {
      removed += 1;
    } else {
      seen.add(key);
      uniqueRows.push({
        cells,
        rowModel: removeDuplicatesCloneRowModel(sheet, row),
      });
    }
  }
  return { startRow, uniqueRows, removed };
}



function applyCollectedRemoveDuplicateRows(sheet, range, collected) {
  const startRow = collected.startRow;
  const height = range.bottom - startRow + 1;
  for (let offset = 0; offset < height; offset += 1) {
    const targetRow = startRow + offset;
    const source = collected.uniqueRows[offset] || null;
    const sourceCells = source?.cells || [];
    sheet.rows[targetRow - 1] = source?.rowModel ? { ...defaultRowModel(sheet), ...source.rowModel } : defaultRowModel(sheet);
    for (let col = range.left; col <= range.right; col += 1) {
      writeClonedCellContent(sheet, targetRow, col, sourceCells[col - range.left] || null);
    }
  }
  internCellsInRange(sheet, range);
}



function removeDuplicateRowsForSheetGroup(range, selectedCols, hasHeader) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    const targetRange = materializeSelectionRangeForSheet(sheet, range, sheetIndex);
    if (!targetRange) continue;
    if (!ensureEditableRangeForProtection(sheet, targetRange, "削除")) return true;
    const mergeConflict = removeDuplicatesMergeConflict(sheet, targetRange, hasHeader);
    if (mergeConflict) {
      showMessageBox(`${sheet.name}：${mergeConflict}`, { title: "Cht WebSheet" });
      return true;
    }
    items.push({ sheetIndex, sheet, range: targetRange, collected: collectRemoveDuplicateRows(sheet, targetRange, selectedCols, hasHeader) });
  }
  if (!items.length) return true;
  const removed = items.reduce((total, item) => total + item.collected.removed, 0);
  if (!removed) {
    showMessageBox("重複する値は見つかりませんでした。", { title: "Cht WebSheet" });
    return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => applyCollectedRemoveDuplicateRows(item.sheet, item.range, item.collected));
    state.hf = buildFormulaEngine(state.model);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
      selectionAfterRanges: activeSelectionRanges().map((item) => ({ ...item })),
      label: `${items.length} シート 重複の削除`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートで${removed}個の重複値を削除しました。`);
    const remaining = items.reduce((total, item) => total + item.collected.uniqueRows.length, 0);
    showMessageBox(`重複する値が ${removed} 個見つかり、削除されました。\n一意の値が ${remaining} 個残っています。`, { title: "Cht WebSheet" });
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`重複の削除を実行できませんでした：${error.message || error}`);
  }
  return true;
}



function removeDuplicatesMergeConflict(sheet, range, hasHeader = false) {
  if (!sheet?.merges?.length || !range) return "";
  const dataRange = { ...range, top: hasHeader ? range.top + 1 : range.top };
  if (dataRange.top > dataRange.bottom) return "";
  const merges = mergesIntersectingRange(sheet, dataRange);
  if (!merges.length) return "";
  const partial = merges.some((merge) => !rangeContains(dataRange, merge));
  if (partial) return "重複を削除するには、結合セルが選択範囲内に完全に含まれている必要があります。";
  return "結合セルを含む範囲から重複を削除することはできません。結合を解除してから実行してください。";
}



function removeDuplicatesCloneRowModel(sheet, row) {
  const rowModel = sheet?.rows?.[row - 1];
  return rowModel ? { ...rowModel } : defaultRowModel(sheet);
}



function removeDuplicatesComparableValue(sheet, sheetId, row, col) {
  const cell = sheet.cells[cellKey(row, col)];
  const rawValue = cell?.formula ? cell?.cached : cell?.raw ?? cell?.cached;
  const normalizedValue = removeDuplicatesNormalizedScalar(rawValue);
  if (normalizedValue) return normalizedValue;
  const display = String(getDisplayForCell(sheet, sheetId, row, col) ?? "");
  const displayDate = comparableDateSerial(display);
  if (displayDate != null) return `d:${displayDate}`;
  return `s:${display.toLocaleLowerCase("ja-JP")}`;
}



function removeDuplicatesNormalizedScalar(value) {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `d:${comparableDateSerial(value)}`;
  }
  if (typeof value === "number" && Number.isFinite(value)) return `n:${value}`;
  if (typeof value === "boolean") return `b:${value ? 1 : 0}`;
  return null;
}



function openDataValidationMenu(trigger = null) {
  if (!trigger) {
    openDataValidationDialog();
    return;
  }
  const hasInvalidMarks = Boolean(Object.keys(validationMarkStore()).length);
  const clearValidationDisabled = "";
  const clearCirclesDisabled = hasInvalidMarks ? "" : " disabled";
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonColorPalette();
  closeRibbonBorderMenu();
  closeRibbonUnderlineMenu();
  closeRibbonShapeMenu();
  closeShapeFormatMenu();
  closeWhatIfMenu();
  closeDataValidationMenu();
  const menu = document.createElement("div");
  menu.id = "dataValidationMenu";
  menu.className = "ribbon-data-menu data-validation-menu";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", "データの入力規則");
  menu.innerHTML = `
    <button type="button" role="menuitem" data-validation-menu-action="open"><i data-lucide="check"></i><span>${accessKeyLabelHtml("データの入力規則(V)...")}</span></button>
    <button type="button" role="menuitem" data-validation-menu-action="clear-validation"${clearValidationDisabled}><i data-lucide="eraser"></i><span>${accessKeyLabelHtml("入力規則のクリア(C)")}</span></button>
    <button type="button" role="menuitem" data-validation-menu-action="circle-invalid"><i data-lucide="badge-plus"></i><span>${accessKeyLabelHtml("無効データのマーク(C)")}</span></button>
    <button type="button" role="menuitem" data-validation-menu-action="clear-circles"${clearCirclesDisabled}><i data-lucide="eraser"></i><span>${accessKeyLabelHtml("入力規則マークのクリア(R)")}</span></button>`;
  enhanceCommandMenuAccessKeys(menu);
  document.body.append(menu);
  if (window.lucide) window.lucide.createIcons({ icons });
  const rect = trigger.getBoundingClientRect();
  positionPopupMenu(menu, rect.left, rect.bottom + 2);
  state.dataValidationMenu = true;
  menu.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
}



function closeDataValidationMenu() {
  document.getElementById("dataValidationMenu")?.remove();
  state.dataValidationMenu = null;
}



function closeDataValidationMenuOnOutsidePointer(event) {
  if (!state.dataValidationMenu) return;
  if (event.target.closest?.("#dataValidationMenu") || event.target.closest?.("[data-ribbon-command='data-validation']")) return;
  closeDataValidationMenu();
}



function closeDataValidationMenuOnEscape(event) {
  if (event.key === "Escape" && state.dataValidationMenu) {
    consumeCommandMenuKey(event);
    closeDataValidationMenu();
  }
}



function handleDataValidationMenuAction(event) {
  if (event.currentTarget.disabled || event.currentTarget.classList.contains("is-disabled")) return;
  const action = event.currentTarget.dataset.validationMenuAction;
  closeDataValidationMenu();
  if (action === "open") {
    openDataValidationDialog();
    return;
  }
  if (action === "clear-validation") {
    clearDataValidationFromSelection();
    return;
  }
  if (action === "circle-invalid") {
    markInvalidDataValidationCells();
    return;
  }
  if (action === "clear-circles") {
    clearInvalidDataValidationMarks();
  }
}



function rangeHasDataValidation(sheet, range) {
  if (!sheet || !range) return false;
  return Object.entries(sheet.cells || {}).some(([key, cell]) => {
    if (!cell?.validation) return false;
    const [row, col] = parseCellKey(key);
    return isCellInRange(range, row, col);
  });
}



function clearDataValidationFromSelection() {
  const context = selectedDataRange("入力規則のクリア");
  if (!context) return;
  const { sheet, range } = context;
  if (clearDataValidationFromSheetGroup(range)) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const marks = validationMarkStore();
  let changed = 0;
  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const cell = sheet.cells[key];
    if (!cell?.validation) return;
    delete cell.validation;
    delete marks[key];
    changed += 1;
  });
  if (!changed) {
    showMessageBox("選択範囲に入力規則はありません。", { title: "Cht WebSheet" });
    return;
  }
  internCellsInRange(sheet, range);
  commitDataOperation(sheet, sheetBefore, range, `${rangeToLabel(range)} 入力規則クリア`, "選択範囲の入力規則をクリアしました。", range);
}



function clearDataValidationFromSheetGroup(range) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートでは入力規則を変更できません。")) return true;
    const targetRanges = materializedRangesForSheet(sheet, [range], sheetIndex);
    if (targetRanges.length) items.push({ sheetIndex, sheet, ranges: targetRanges });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  let changed = 0;
  items.forEach((item) => {
    const marks = validationMarkStore(item.sheetIndex);
    item.ranges.forEach((targetRange) => {
      forEachCellInRange(targetRange, (row, col) => {
        const key = cellKey(row, col);
        const cell = item.sheet.cells[key];
        if (!cell?.validation) return;
        delete cell.validation;
        delete marks[key];
        changed += 1;
      });
      internCellsInRange(item.sheet, targetRange);
    });
  });
  if (!changed) {
    showMessageBox("選択範囲に入力規則はありません。", { title: "Cht WebSheet" });
    return true;
  }
  commitDataOperation(
    activeSheet(),
    null,
    range,
    `${rangeToLabel(range)} ${sheetIndexes.length} シート入力規則クリア`,
    `${sheetIndexes.length} 個のシートの選択範囲の入力規則をクリアしました。`,
    range,
    { sheetsBefore },
  );
  return true;
}



function validationMarkStore(sheetIndex = state.activeSheetIndex) {
  if (!state.validationInvalidMarks) state.validationInvalidMarks = {};
  if (!state.validationInvalidMarks[sheetIndex]) state.validationInvalidMarks[sheetIndex] = {};
  return state.validationInvalidMarks[sheetIndex];
}



function hasValidationInvalidMark(sheetIndex, row, col) {
  return Boolean(state.validationInvalidMarks?.[sheetIndex]?.[cellKey(row, col)]);
}



function markInvalidDataValidationCells() {
  const sheet = activeSheet();
  if (!sheet) return;
  const marks = {};
  const sheetId = getSheetId(sheet.name);
  Object.entries(sheet.cells || {}).forEach(([key, cell]) => {
    if (!cell?.validation) return;
    const [row, col] = parseCellKey(key);
    const value = cell.raw ?? getDisplayForCell(sheet, sheetId, row, col);
    if (!validateCellInput(sheet, row, col, value, { ignoreErrorAlertDisabled: true }).valid) marks[key] = true;
  });
  if (!state.validationInvalidMarks) state.validationInvalidMarks = {};
  state.validationInvalidMarks[state.activeSheetIndex] = marks;
  renderWorkbook();
  const count = Object.keys(marks).length;
  setStatus(count ? `${count} 個の無効なデータをマークしました。` : "無効なデータは見つかりませんでした。");
}



function clearInvalidDataValidationMarks() {
  const marks = validationMarkStore();
  const count = Object.keys(marks).length;
  state.validationInvalidMarks[state.activeSheetIndex] = {};
  renderWorkbook();
  setStatus(count ? "入力規則マークをクリアしました。" : "クリアする入力規則マークはありません。");
}



function handleValidationDropdownButtonPointerDown(event) {
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
}



function handleValidationDropdownButtonClick(event) {
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  const cellElement = event.currentTarget.closest(".sheet-cell[data-row][data-col]");
  const point = cellElement ? cellPointFromElement(cellElement) : null;
  if (!point || point.sheetIndex !== state.activeSheetIndex) return;
  selectCellPoint(point);
  openValidationDropdownPopup(point.row, point.col);
}



function openValidationDropdownPopup(row, col) {
  const sheet = activeSheet();
  const cell = sheet?.cells?.[cellKey(row, col)];
  const options = validationListOptions(sheet, cell?.validation);
  if (!sheet || !cell?.validation || !options.length) {
    showMessageBox("このセルには選択できるドロップダウン リストがありません。", { title: "Cht WebSheet" });
    return;
  }
  closeValidationDropdownPopup();
  const target = document.querySelector(`.sheet-cell[data-row="${row}"][data-col="${col}"]`);
  if (!target) {
    scrollRangeIntoView(normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col }));
    setStatus("ドロップダウン リストを表示するセルが画面外にあります。");
    return;
  }
  const sheetId = getSheetId(sheet.name);
  const current = String(cell.raw ?? getDisplayForCell(sheet, sheetId, row, col) ?? "");
  const popup = document.createElement("div");
  popup.id = "validationDropdownPopup";
  popup.className = "validation-dropdown-popup";
  popup.setAttribute("role", "listbox");
  popup.setAttribute("aria-label", `${columnName(col)}${row} ドロップダウン リスト`);
  popup.dataset.row = String(row);
  popup.dataset.col = String(col);
  popup.innerHTML = options
    .map((option) => {
      const value = String(option);
      const selected = value === current ? " is-selected" : "";
      return `<button class="validation-dropdown-choice${selected}" type="button" role="option" aria-selected="${selected ? "true" : "false"}" data-validation-dropdown-choice="${escapeAttr(value)}">${escapeHtml(value)}</button>`;
    })
    .join("");
  document.body.append(popup);
  const rect = target.getBoundingClientRect();
  positionValidationDropdownPopup(popup, rect);
  state.validationDropdownPopup = { row, col };
  (popup.querySelector(".is-selected") || popup.querySelector("[data-validation-dropdown-choice]"))?.focus({ preventScroll: true });
}



function openSelectedValidationDropdown() {
  const point = firstPointForActiveSelection();
  if (!point || point.sheetIndex !== state.activeSheetIndex) return false;
  const sheet = activeSheet();
  const cell = sheet?.cells?.[cellKey(point.row, point.col)];
  if (!validationListOptions(sheet, cell?.validation).length) return false;
  const range = selectionRangeFromPoints(point, point);
  if (!rangeIntersects(range, visibleGridRange(sheet)) || scrollRangeIntoView(range)) {
    renderSheet();
  }
  openValidationDropdownPopup(point.row, point.col);
  return true;
}



function positionValidationDropdownPopup(popup, cellRect) {
  const margin = 6;
  popup.style.visibility = "hidden";
  popup.style.left = "0px";
  popup.style.top = "0px";
  const rect = popup.getBoundingClientRect();
  const width = Math.max(rect.width, cellRect.width, 120);
  popup.style.width = `${Math.round(width)}px`;
  const adjusted = popup.getBoundingClientRect();
  const left = Math.max(margin, Math.min(cellRect.left, window.innerWidth - adjusted.width - margin));
  const top = Math.max(margin, Math.min(cellRect.bottom, window.innerHeight - adjusted.height - margin));
  popup.style.left = `${Math.round(left)}px`;
  popup.style.top = `${Math.round(top)}px`;
  popup.style.visibility = "visible";
}



function handleValidationDropdownChoice(event) {
  const popup = event.currentTarget.closest("#validationDropdownPopup");
  const row = Number(popup?.dataset.row);
  const col = Number(popup?.dataset.col);
  const value = event.currentTarget.dataset.validationDropdownChoice || "";
  closeValidationDropdownPopup();
  if (!row || !col) return;
  commitCellValue(row, col, value, { skipValidation: true });
  setStatus(`${columnName(col)}${row} に「${value}」を入力しました。`);
}



function handleValidationDropdownKeydown(event) {
  const popup = event.currentTarget.closest("#validationDropdownPopup");
  const choices = [...(popup?.querySelectorAll("[data-validation-dropdown-choice]") || [])];
  if (!choices.length) return;
  const currentIndex = Math.max(0, choices.indexOf(event.currentTarget));
  const focusChoice = (index) => {
    const next = choices[Math.max(0, Math.min(choices.length - 1, index))];
    next?.focus({ preventScroll: true });
  };
  const consume = () => {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
  };
  switch (event.key) {
    case "ArrowDown":
      consume();
      focusChoice(currentIndex + 1);
      break;
    case "ArrowUp":
      consume();
      focusChoice(currentIndex - 1);
      break;
    case "Home":
      consume();
      focusChoice(0);
      break;
    case "End":
      consume();
      focusChoice(choices.length - 1);
      break;
    case "PageDown":
      consume();
      focusChoice(currentIndex + 8);
      break;
    case "PageUp":
      consume();
      focusChoice(currentIndex - 8);
      break;
    case "Enter":
    case " ":
      consume();
      event.currentTarget.click();
      break;
    default:
      break;
  }
}



function closeValidationDropdownOnOutsidePointer(event) {
  if (!state.validationDropdownPopup) return;
  if (event.target.closest?.("#validationDropdownPopup")) return;
  closeValidationDropdownPopup();
}



function closeValidationDropdownOnEscape(event) {
  if (event.key === "Escape" && state.validationDropdownPopup) closeValidationDropdownPopup();
}



function closeValidationDropdownPopup() {
  document.querySelector("#validationDropdownPopup")?.remove();
  state.validationDropdownPopup = null;
}



function openDataValidationDialog() {
  const context = selectedDataRange("データの入力規則");
  if (!context) return;
  const firstCell = activeSheet().cells[cellKey(context.range.top, context.range.left)];
  const validation = firstCell?.validation || {};
  const originalValidationKey = firstCell?.validation ? validationComparisonKey(firstCell.validation) : "";
  state.dataDialog = { type: "validation", range: context.range, originalValidationKey };
  closeDataDialog({ keepState: true });
  const validationType = validation.type || "any";
  const validationOperator = validation.operator || "between";
  const formulae = validation.formulae || [];
  const sourceValue = validation.type === "list" ? validation.formula || (validation.options || []).join(",") : formulae[0] || "";
  const sourceValue2 = formulae[1] || "";
  const sameValidationDisabled = originalValidationKey ? "" : " disabled";
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog validation-dialog" role="dialog" aria-modal="true" aria-labelledby="validationDialogTitle" data-data-dialog="validation">
      <div class="excel-mini-titlebar" id="validationDialogTitle">データの入力規則</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-tabs" role="tablist">
          <button class="is-active" type="button" role="tab" data-validation-tab="settings" data-data-dialog-action="validation-tab">設定</button>
          <button type="button" role="tab" data-validation-tab="input" data-data-dialog-action="validation-tab">入力時メッセージ</button>
          <button type="button" role="tab" data-validation-tab="error" data-data-dialog-action="validation-tab">エラー メッセージ</button>
          <button type="button" role="tab" data-validation-tab="ime" data-data-dialog-action="validation-tab">日本語入力</button>
        </div>
        <div class="data-dialog-tab-panel is-active" data-validation-tab-panel="settings">
          <label class="excel-dialog-field"><span class="excel-dialog-caption">入力値の種類</span>
            <select data-data-field="allow" data-validation-sync>
              <option value="any" ${validationType === "any" ? "selected" : ""}>すべての値</option>
              <option value="whole" ${validationType === "whole" ? "selected" : ""}>整数</option>
              <option value="decimal" ${validationType === "decimal" ? "selected" : ""}>小数点数</option>
              <option value="list" ${validationType === "list" ? "selected" : ""}>リスト</option>
              <option value="date" ${validationType === "date" ? "selected" : ""}>日付</option>
              <option value="time" ${validationType === "time" ? "selected" : ""}>時刻</option>
              <option value="textLength" ${validationType === "textLength" ? "selected" : ""}>文字列（長さ指定）</option>
              <option value="custom" ${validationType === "custom" ? "selected" : ""}>ユーザー設定</option>
            </select>
          </label>
          <label class="excel-dialog-field" data-validation-row="operator"><span class="excel-dialog-caption">データ</span>
            <select data-data-field="operator" data-validation-sync>
              <option value="between" ${validationOperator === "between" ? "selected" : ""}>次の値の間</option>
              <option value="notBetween" ${validationOperator === "notBetween" ? "selected" : ""}>次の値の間以外</option>
              <option value="equal" ${validationOperator === "equal" ? "selected" : ""}>次の値に等しい</option>
              <option value="notEqual" ${validationOperator === "notEqual" ? "selected" : ""}>次の値に等しくない</option>
              <option value="greaterThan" ${validationOperator === "greaterThan" ? "selected" : ""}>次の値より大きい</option>
              <option value="lessThan" ${validationOperator === "lessThan" ? "selected" : ""}>次の値より小さい</option>
              <option value="greaterThanOrEqual" ${validationOperator === "greaterThanOrEqual" ? "selected" : ""}>次の値以上</option>
              <option value="lessThanOrEqual" ${validationOperator === "lessThanOrEqual" ? "selected" : ""}>次の値以下</option>
            </select>
          </label>
          <label class="excel-dialog-field data-range-field" data-validation-row="source"><span class="excel-dialog-caption" data-validation-caption="source">元の値 / 最小値</span><span class="data-range-input"><input type="text" data-data-field="source" data-validation-sync value="${escapeAttr(sourceValue)}" /><button class="data-range-picker" type="button" title="現在の選択範囲を使用" aria-label="元の値に現在の選択範囲を使用" data-data-dialog-action="range-from-selection" data-data-range-target="source"></button></span></label>
          <label class="excel-dialog-field" data-validation-row="source2"><span class="excel-dialog-caption" data-validation-caption="source2">最大値</span><input type="text" data-data-field="source2" data-validation-sync value="${escapeAttr(sourceValue2)}" /></label>
          <label class="data-dialog-check"><input type="checkbox" data-data-field="allowBlank" data-validation-sync ${validation.allowBlank !== false ? "checked" : ""} />空白を無視する</label>
          <label class="data-dialog-check" data-validation-row="dropdown"><input type="checkbox" data-data-field="showDropDown" data-validation-sync ${validation.showDropDown !== false ? "checked" : ""} />ドロップダウン リストから選択する</label>
          <label class="data-dialog-check"><input type="checkbox" data-data-field="applySameValidation"${sameValidationDisabled} />同じ入力規則が設定されたすべてのセルに変更を適用する</label>
          <div class="data-dialog-summary validation-summary" data-validation-summary></div>
        </div>
        <div class="data-dialog-tab-panel" data-validation-tab-panel="input">
          <label class="data-dialog-check"><input type="checkbox" data-data-field="showInputMessage" ${validation.showInputMessage ? "checked" : ""} />セルの選択時に入力時メッセージを表示する</label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">タイトル</span><input type="text" data-data-field="inputTitle" value="${escapeAttr(validation.inputTitle || "")}" /></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">入力時メッセージ</span><textarea data-data-field="inputMessage" rows="4">${escapeHtml(validation.inputMessage || "")}</textarea></label>
        </div>
        <div class="data-dialog-tab-panel" data-validation-tab-panel="error">
          <label class="data-dialog-check"><input type="checkbox" data-data-field="showErrorMessage" ${validation.showErrorMessage !== false ? "checked" : ""} />無効なデータが入力されたらエラー メッセージを表示する</label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">スタイル</span>
            <select data-data-field="errorStyle">
              <option value="stop" ${validation.errorStyle !== "warning" && validation.errorStyle !== "information" ? "selected" : ""}>停止</option>
              <option value="warning" ${validation.errorStyle === "warning" ? "selected" : ""}>注意</option>
              <option value="information" ${validation.errorStyle === "information" ? "selected" : ""}>情報</option>
            </select>
          </label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">タイトル</span><input type="text" data-data-field="errorTitle" value="${escapeAttr(validation.errorTitle || "")}" /></label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">エラー メッセージ</span><textarea data-data-field="errorMessage" rows="4">${escapeHtml(validation.errorMessage || "")}</textarea></label>
        </div>
        <div class="data-dialog-tab-panel" data-validation-tab-panel="ime">
          <label class="excel-dialog-field"><span class="excel-dialog-caption">日本語入力</span>
            <select data-data-field="imeMode">
              ${dataValidationImeModeOptions(validation.imeMode || "noControl")}
            </select>
          </label>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="clear">すべてクリア</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncDataValidationDialog(overlay.querySelector("[data-data-dialog='validation']"));
}



function dataValidationImeModeOptions(selected = "noControl") {
  const normalized = selected || "noControl";
  return [
    ["noControl", "コントロールなし"],
    ["on", "オン"],
    ["off", "オフ（英語モード）"],
    ["disabled", "無効"],
    ["hiragana", "ひらがな"],
    ["fullKatakana", "全角カタカナ"],
    ["halfKatakana", "半角カタカナ"],
    ["fullAlpha", "全角英数字"],
    ["halfAlpha", "半角英数字"],
  ]
    .map(([value, label]) => `<option value="${escapeAttr(value)}"${normalized === value ? " selected" : ""}>${escapeHtml(label)}</option>`)
    .join("");
}



function handleDataValidationDialogSync(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='validation']");
  syncDataValidationDialog(dialog);
}



function syncDataValidationDialog(dialog) {
  if (!dialog) return;
  const allow = dialog.querySelector("[data-data-field='allow']")?.value || "any";
  const operator = dialog.querySelector("[data-data-field='operator']")?.value || "between";
  const operatorRow = dialog.querySelector("[data-validation-row='operator']");
  const sourceRow = dialog.querySelector("[data-validation-row='source']");
  const source2Row = dialog.querySelector("[data-validation-row='source2']");
  const dropdownRow = dialog.querySelector("[data-validation-row='dropdown']");
  const sourceCaption = dialog.querySelector("[data-validation-caption='source']");
  const source2Caption = dialog.querySelector("[data-validation-caption='source2']");
  const operatorControl = dialog.querySelector("[data-data-field='operator']");
  const sourceControl = dialog.querySelector("[data-data-field='source']");
  const source2Control = dialog.querySelector("[data-data-field='source2']");
  const dropdownControl = dialog.querySelector("[data-data-field='showDropDown']");
  const sourcePicker = dialog.querySelector("[data-data-range-target='source']");
  const allowAny = allow === "any";
  const allowList = allow === "list";
  const allowCustom = allow === "custom";
  const hasSecondValue = !allowAny && !allowList && !allowCustom && (operator === "between" || operator === "notBetween");
  operatorRow?.classList.toggle("hidden", allowAny || allowList || allowCustom);
  sourceRow?.classList.toggle("hidden", allowAny);
  source2Row?.classList.toggle("hidden", !hasSecondValue);
  dropdownRow?.classList.toggle("hidden", !allowList);
  if (operatorControl) operatorControl.disabled = allowAny || allowList || allowCustom;
  if (sourceControl) sourceControl.disabled = allowAny;
  if (source2Control) source2Control.disabled = !hasSecondValue;
  if (dropdownControl) dropdownControl.disabled = !allowList;
  if (sourcePicker) sourcePicker.disabled = allowAny;
  if (sourceCaption) sourceCaption.textContent = validationSourceCaption(allow, operator, false);
  if (source2Caption) source2Caption.textContent = validationSourceCaption(allow, operator, true);
  const summary = dialog.querySelector("[data-validation-summary]");
  if (summary) {
    const preview = dataValidationDialogPreview(dialog);
    summary.textContent = preview.message;
    summary.classList.toggle("is-warning", preview.warning);
  }
  const validation = dataValidationDialogValidation(dialog);
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = !validation.valid;
}



function dataValidationDialogPreview(dialog) {
  const validation = dataValidationDialogValidation(dialog, { includeSelectionStats: true });
  if (!validation.valid) return { warning: true, message: validation.message };
  const allow = dialog.querySelector("[data-data-field='allow']")?.value || "any";
  const operator = dialog.querySelector("[data-data-field='operator']")?.value || "between";
  const source = String(dialog.querySelector("[data-data-field='source']")?.value || "").trim();
  const source2 = String(dialog.querySelector("[data-data-field='source2']")?.value || "").trim();
  const allowBlank = Boolean(dialog.querySelector("[data-data-field='allowBlank']")?.checked);
  const selectionText = validation.selectionMessage ? ` ${validation.selectionMessage}` : "";
  if (allow === "any") return { warning: false, message: `すべての値を許可します。入力規則は設定されません。${selectionText}` };
  if (allow === "list") {
    const inline = parseValidationOptions(source);
    const reference = validationReferenceFromFormula(activeSheet(), source);
    if (!inline.length && !reference) return { warning: true, message: "リストの元の値または参照範囲を入力してください。" };
    const count = reference ? validationOptionsFromReference(activeSheet(), source).length : inline.length;
    const dropDown = dialog.querySelector("[data-data-field='showDropDown']")?.checked !== false;
    return {
      warning: !count,
      message: reference
        ? `参照範囲から ${count} 件のリストを使用します。${dropDown ? "ドロップダウンを表示します。" : "ドロップダウンは表示しません。"}${selectionText}`
        : `直接入力された ${count} 件のリストを使用します。${dropDown ? "ドロップダウンを表示します。" : "ドロップダウンは表示しません。"}${selectionText}`,
    };
  }
  if (allow === "custom") {
    if (!source) return { warning: true, message: "ユーザー設定の数式を入力してください。" };
    return { warning: false, message: `ユーザー設定の数式 ${source} で入力値を判定します。${allowBlank ? "空白は許可します。" : "空白も検査します。"}${selectionText}` };
  }
  const needsSecond = operator === "between" || operator === "notBetween";
  if (!source || (needsSecond && !source2)) {
    return { warning: true, message: `${validationSourceCaption(allow, operator, needsSecond)}を入力してください。` };
  }
  const typeLabel = dataValidationTypeLabel(allow);
  const operatorLabel = dataValidationOperatorLabel(operator);
  const rangeText = needsSecond ? `${source} ～ ${source2}` : source;
  return { warning: false, message: `${typeLabel}が「${rangeText}」${operatorLabel}場合だけ入力できます。${allowBlank ? "空白は許可します。" : "空白も検査します。"}${selectionText}` };
}



function dataValidationDialogValidation(dialog, options = {}) {
  const allow = dialog?.querySelector("[data-data-field='allow']")?.value || "any";
  const operator = dialog?.querySelector("[data-data-field='operator']")?.value || "between";
  const source = String(dialog?.querySelector("[data-data-field='source']")?.value || "").trim();
  const source2 = String(dialog?.querySelector("[data-data-field='source2']")?.value || "").trim();
  if (allow === "any") return { valid: true, selectionMessage: dataValidationSelectionStatsMessage(null, options) };
  if (allow === "list") {
    const inline = parseValidationOptions(source);
    const reference = validationReferenceFromFormula(activeSheet(), source);
    const optionsCount = reference ? validationOptionsFromReference(activeSheet(), source).length : inline.length;
    if (!inline.length && !reference) return { valid: false, message: "リストの元の値または参照範囲を入力してください。" };
    if (reference && !optionsCount) return { valid: false, message: "参照範囲にリスト項目がありません。" };
    const rule = validationFromDialogFieldsForPreview(dialog);
    return { valid: true, selectionMessage: dataValidationSelectionStatsMessage(rule, options) };
  }
  if (allow === "custom") {
    if (!source) return { valid: false, message: "ユーザー設定の数式を入力してください。" };
    if (!String(source).trim().startsWith("=")) return { valid: false, message: "ユーザー設定には数式を入力してください。" };
    const rule = validationFromDialogFieldsForPreview(dialog);
    return { valid: true, selectionMessage: dataValidationSelectionStatsMessage(rule, options) };
  }
  const needsSecond = operator === "between" || operator === "notBetween";
  if (!source || (needsSecond && !source2)) return { valid: false, message: `${validationSourceCaption(allow, operator, needsSecond)}を入力してください。` };
  const first = source.startsWith("=") ? { valid: true } : validationComparableValue(allow, source, true);
  const second = !needsSecond || source2.startsWith("=") ? { valid: true } : validationComparableValue(allow, source2, true);
  if (!first.valid || !second.valid) return { valid: false, message: "入力規則の条件値を正しく入力してください。" };
  const rule = validationFromDialogFieldsForPreview(dialog);
  return { valid: true, selectionMessage: dataValidationSelectionStatsMessage(rule, options) };
}



function validationFromDialogFieldsForPreview(dialog) {
  const allow = dialog?.querySelector("[data-data-field='allow']")?.value || "any";
  if (allow === "any") return null;
  return validationFromDialogFields(
    allow,
    dialog?.querySelector("[data-data-field='operator']")?.value || "between",
    dialog?.querySelector("[data-data-field='source']")?.value || "",
    dialog?.querySelector("[data-data-field='source2']")?.value || "",
    {
      allowBlank: Boolean(dialog?.querySelector("[data-data-field='allowBlank']")?.checked),
      anchor: { row: state.dataDialog?.range?.top || 1, col: state.dataDialog?.range?.left || 1 },
      showDropDown: Boolean(dialog?.querySelector("[data-data-field='showDropDown']")?.checked),
      showInputMessage: Boolean(dialog?.querySelector("[data-data-field='showInputMessage']")?.checked),
      inputTitle: dialog?.querySelector("[data-data-field='inputTitle']")?.value || "",
      inputMessage: dialog?.querySelector("[data-data-field='inputMessage']")?.value || "",
      showErrorMessage: Boolean(dialog?.querySelector("[data-data-field='showErrorMessage']")?.checked),
      errorStyle: dialog?.querySelector("[data-data-field='errorStyle']")?.value || "stop",
      errorTitle: dialog?.querySelector("[data-data-field='errorTitle']")?.value || "",
      errorMessage: dialog?.querySelector("[data-data-field='errorMessage']")?.value || "",
      imeMode: dialog?.querySelector("[data-data-field='imeMode']")?.value || "noControl",
      silent: true,
    },
  );
}



function dataValidationSelectionStatsMessage(rule, options = {}) {
  if (!options.includeSelectionStats) return "";
  const sheet = activeSheet();
  const range = state.dataDialog?.range;
  if (!sheet || !range) return "";
  const original = [];
  const sheetId = getSheetId(sheet.name);
  let checked = 0;
  let invalid = 0;
  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const cell = sheet.cells[key];
    original.push({ key, cell: cell ? cloneCellModel(cell) : null });
    if (rule) {
      const target = ensureCellModel(sheet, row, col);
      target.validation = cloneJsonLikeValue(rule);
    } else if (cell) {
      delete cell.validation;
    }
  });
  try {
    forEachCellInRange(range, (row, col) => {
      checked += 1;
      const value = sheet.cells[cellKey(row, col)]?.raw ?? getDisplayForCell(sheet, sheetId, row, col);
      if (!validateCellInput(sheet, row, col, value, { ignoreErrorAlertDisabled: true }).valid) invalid += 1;
    });
  } finally {
    original.forEach(({ key, cell }) => {
      if (cell) storeCellModelForSheet(sheet, key, cloneCellModel(cell));
      else delete sheet.cells[key];
    });
  }
  return rule ? `対象範囲 ${checked.toLocaleString("ja-JP")} セル中、現在値が無効になるセルは ${invalid.toLocaleString("ja-JP")} 個です。` : `対象範囲 ${checked.toLocaleString("ja-JP")} セルの入力規則を解除します。`;
}



function dataValidationTypeLabel(type) {
  switch (type) {
    case "whole":
      return "整数";
    case "decimal":
      return "小数点数";
    case "date":
      return "日付";
    case "time":
      return "時刻";
    case "textLength":
      return "文字列の長さ";
    default:
      return "値";
  }
}



function dataValidationOperatorLabel(operator) {
  switch (operator) {
    case "between":
      return "の間にある";
    case "notBetween":
      return "の間にない";
    case "notEqual":
      return "と等しくない";
    case "greaterThan":
      return "より大きい";
    case "lessThan":
      return "より小さい";
    case "greaterThanOrEqual":
      return "以上である";
    case "lessThanOrEqual":
      return "以下である";
    case "equal":
    default:
      return "と等しい";
  }
}



function validationSourceCaption(allow, operator, second = false) {
  if (allow === "list") return "元の値";
  if (allow === "custom") return "数式";
  if (allow === "date") {
    if (operator === "between" || operator === "notBetween") return second ? "終了日" : "開始日";
    return "日付";
  }
  if (allow === "time") {
    if (operator === "between" || operator === "notBetween") return second ? "終了時刻" : "開始時刻";
    return "時刻";
  }
  if (operator === "between" || operator === "notBetween") return second ? "最大値" : "最小値";
  return "値";
}



function resetDataValidationDialog(dialog) {
  if (!dialog) return;
  const setField = (name, value) => {
    const field = dialog.querySelector(`[data-data-field='${name}']`);
    if (!field) return;
    if (field.type === "checkbox") field.checked = Boolean(value);
    else field.value = value;
  };
  setField("allow", "any");
  setField("operator", "between");
  setField("source", "");
  setField("source2", "");
  setField("allowBlank", true);
  setField("showDropDown", true);
  setField("showInputMessage", false);
  setField("inputTitle", "");
  setField("inputMessage", "");
  setField("showErrorMessage", true);
  setField("errorStyle", "stop");
  setField("errorTitle", "");
  setField("errorMessage", "");
  setField("imeMode", "noControl");
  syncDataValidationDialog(dialog);
}



function applyDataValidationDialog(dialog, clear = false) {
  const context = state.dataDialog;
  if (!context?.range) return;
  if (!clear) {
    const dialogValidation = dataValidationDialogValidation(dialog);
    if (!dialogValidation.valid) {
      showMessageBox(dialogValidation.message, { title: "Cht WebSheet" });
      syncDataValidationDialog(dialog);
      return;
    }
  }
  const applyToSameValidation = Boolean(dialog.querySelector("[data-data-field='applySameValidation']")?.checked);
  const allow = dialog.querySelector("[data-data-field='allow']")?.value || "any";
  const source = dialog.querySelector("[data-data-field='source']")?.value || "";
  const source2 = dialog.querySelector("[data-data-field='source2']")?.value || "";
  const operator = dialog.querySelector("[data-data-field='operator']")?.value || "between";
  const allowBlank = Boolean(dialog.querySelector("[data-data-field='allowBlank']")?.checked);
  const shouldClear = clear || allow === "any";
  const validation = shouldClear ? null : validationFromDialogFields(allow, operator, source, source2, {
    allowBlank,
    anchor: { row: context.range.top, col: context.range.left },
    showDropDown: Boolean(dialog.querySelector("[data-data-field='showDropDown']")?.checked),
    showInputMessage: Boolean(dialog.querySelector("[data-data-field='showInputMessage']")?.checked),
    inputTitle: dialog.querySelector("[data-data-field='inputTitle']")?.value || "",
    inputMessage: dialog.querySelector("[data-data-field='inputMessage']")?.value || "",
    showErrorMessage: Boolean(dialog.querySelector("[data-data-field='showErrorMessage']")?.checked),
    errorStyle: dialog.querySelector("[data-data-field='errorStyle']")?.value || "stop",
    errorTitle: dialog.querySelector("[data-data-field='errorTitle']")?.value || "",
    errorMessage: dialog.querySelector("[data-data-field='errorMessage']")?.value || "",
    imeMode: dialog.querySelector("[data-data-field='imeMode']")?.value || "noControl",
  });
  if (!shouldClear && !validation) return;
  if (applyDataValidationDialogToSheetGroup(context, validation, shouldClear, applyToSameValidation)) return;
  if (!ensureActiveSheetFormatCellsAllowed("保護されたシートでは入力規則を変更できません。")) return;
  const sheet = activeSheet();
  const sheetBefore = captureSheetForHistory(sheet);
  const targetPoints = validationTargetPointsForDialog(sheet, context.range, context.originalValidationKey, applyToSameValidation);
  applyDataValidationToSheetPoints(sheet, targetPoints, validation);
  const targetRange = boundingRangeForPoints(targetPoints, context.range);
  const status = shouldClear
    ? applyToSameValidation ? "同じ入力規則が設定されたセルの入力規則をクリアしました。" : "入力規則をクリアしました。"
    : applyToSameValidation ? "同じ入力規則が設定されたセルに入力規則を設定しました。" : "入力規則を設定しました。";
  commitDataOperation(sheet, sheetBefore, targetRange, `${rangeToLabel(context.range)} 入力規則`, status, context.range);
  closeDataDialog();
}



function applyDataValidationDialogToSheetGroup(context, validation, shouldClear, applyToSameValidation) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !context?.range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートでは入力規則を変更できません。")) return true;
    const targetRanges = materializedRangesForSheet(sheet, [context.range], sheetIndex);
    if (!targetRanges.length) continue;
    const points = [];
    targetRanges.forEach((range) => {
      points.push(...validationTargetPointsForDialog(sheet, range, context.originalValidationKey, applyToSameValidation));
    });
    const targetRange = boundingRangeForPoints(points, targetRanges[0], sheetIndex);
    items.push({ sheetIndex, sheet, points, targetRange });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => applyDataValidationToSheetPoints(item.sheet, item.points, validation));
    const status = shouldClear
      ? applyToSameValidation
        ? `${sheetIndexes.length} 個のシートの同じ入力規則が設定されたセルの入力規則をクリアしました。`
        : `${sheetIndexes.length} 個のシートの入力規則をクリアしました。`
      : applyToSameValidation
        ? `${sheetIndexes.length} 個のシートの同じ入力規則が設定されたセルに入力規則を設定しました。`
        : `${sheetIndexes.length} 個のシートに入力規則を設定しました。`;
    commitDataOperation(
      activeSheet(),
      null,
      context.range,
      `${rangeToLabel(context.range)} ${sheetIndexes.length} シート入力規則`,
      status,
      context.range,
      { sheetsBefore },
    );
    closeDataDialog();
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`入力規則を適用できませんでした：${error.message || error}`);
  }
  return true;
}



function applyDataValidationToSheetPoints(sheet, targetPoints, validation) {
  const touchedRanges = [];
  (targetPoints || []).forEach(({ row, col }) => {
    const cell = ensureCellModel(sheet, row, col);
    if (validation) cell.validation = cloneJsonLikeValue(validation);
    else delete cell.validation;
    touchedRanges.push({ top: row, bottom: row, left: col, right: col });
  });
  internCellsInRanges(sheet, touchedRanges);
}



function validationTargetPointsForDialog(sheet, range, originalValidationKey = "", applyToSameValidation = false) {
  const points = [];
  const seen = new Set();
  const push = (row, col) => {
    const key = cellKey(row, col);
    if (seen.has(key)) return;
    seen.add(key);
    points.push({ row, col });
  };
  if (applyToSameValidation && originalValidationKey) {
    Object.values(sheet?.cells || {}).forEach((cell) => {
      if (cell?.validation && validationComparisonKey(cell.validation) === originalValidationKey) {
        push(cell.row, cell.col);
      }
    });
  }
  forEachCellInRange(range, push);
  return points;
}



function boundingRangeForPoints(points, fallback, sheetIndex = fallback?.sheetIndex ?? state.activeSheetIndex) {
  if (!points?.length) return fallback;
  let top = points[0].row;
  let bottom = points[0].row;
  let left = points[0].col;
  let right = points[0].col;
  points.forEach((point) => {
    top = Math.min(top, point.row);
    bottom = Math.max(bottom, point.row);
    left = Math.min(left, point.col);
    right = Math.max(right, point.col);
  });
  return { sheetIndex, top, bottom, left, right };
}



function validationFromDialogFields(allow, operator, source, source2, options = {}) {
  const base = {
    allowBlank: Boolean(options.allowBlank),
    showInputMessage: Boolean(options.showInputMessage),
    inputTitle: options.inputTitle || "",
    inputMessage: options.inputMessage || "",
    showErrorMessage: options.showErrorMessage !== false,
    errorStyle: options.errorStyle || "stop",
    errorTitle: options.errorTitle || "",
    errorMessage: options.errorMessage || "",
    imeMode: options.imeMode || "noControl",
    ...(options.anchor ? { anchor: { row: options.anchor.row, col: options.anchor.col } } : {}),
  };
  if (allow === "list") {
    const listOptions = String(source)
      ? parseValidationOptions(source)
      : [];
    if (!listOptions.length && !validationReferenceFromFormula(activeSheet(), source)) {
      showMessageBox("リストの元の値を入力してください。", { title: "Cht WebSheet" });
      return null;
    }
    return { ...base, type: "list", formula: source, options: listOptions, showDropDown: options.showDropDown !== false };
  }
  const formulae = [source, source2].map((item) => String(item ?? "").trim()).filter(Boolean);
  return { ...base, type: allow, operator: operator || (formulae.length > 1 ? "between" : "equal"), formulae };
}



function validateCellInput(sheet, row, col, content, options = {}) {
  const rule = sheet.cells[cellKey(row, col)]?.validation;
  if (!rule?.type) return { valid: true };
  if (typeof content === "string" && content.startsWith("=")) return { valid: true };
  if ((content == null || content === "") && rule.allowBlank !== false) return { valid: true };
  if (rule.showErrorMessage === false && !options.ignoreErrorAlertDisabled) return { valid: true };
  const invalid = () => ({
    valid: false,
    title: rule.errorTitle || "Cht WebSheet",
    style: rule.errorStyle || "stop",
    message:
      rule.errorMessage ||
      "入力した値は、このセルに設定されているデータの入力規則を満たしていません。",
  });

  if (rule.type === "list") {
    const options = validationListOptions(sheet, rule);
    if (!options.length) return { valid: true };
    const candidate = normalizeValidationListValue(content);
    return options.map(normalizeValidationListValue).includes(candidate) ? { valid: true } : invalid();
  }
  if (rule.type === "custom") return validateCustomDataValidation(sheet, row, col, content, rule, invalid);

  const candidate = validationComparableValue(rule.type, content);
  if (!candidate.valid) return invalid();
  const formulae = rule.formulae || [];
  if (!formulae.length) return { valid: true };
  const bounds = formulae
    .map((formula) => validationComparableBoundValue(sheet, row, col, rule, formula))
    .filter((item) => item.valid)
    .map((item) => item.value);
  if (!bounds.length) return { valid: true };
  const value = candidate.value;
  const first = bounds[0];
  const second = bounds.length > 1 ? bounds[1] : first;
  switch (rule.operator || "equal") {
    case "between":
      return value >= Math.min(first, second) && value <= Math.max(first, second) ? { valid: true } : invalid();
    case "notBetween":
      return value < Math.min(first, second) || value > Math.max(first, second) ? { valid: true } : invalid();
    case "notEqual":
      return value !== first ? { valid: true } : invalid();
    case "greaterThan":
      return value > first ? { valid: true } : invalid();
    case "lessThan":
      return value < first ? { valid: true } : invalid();
    case "greaterThanOrEqual":
      return value >= first ? { valid: true } : invalid();
    case "lessThanOrEqual":
      return value <= first ? { valid: true } : invalid();
    case "equal":
    default:
      return value === first ? { valid: true } : invalid();
  }
}



function normalizeValidationListValue(value) {
  return String(value ?? "").trim().toLocaleLowerCase("ja-JP");
}



function validationComparableBoundValue(sheet, row, col, rule, formula) {
  const text = String(formula ?? "").trim();
  if (text.startsWith("=")) {
    const evaluated = evaluateValidationFormulaValue(sheet, row, col, rule, text);
    if (evaluated !== undefined) return validationComparableValue(rule.type, evaluated, false);
  }
  return validationComparableValue(rule.type, formula, true);
}



function evaluateValidationFormulaValue(sheet, row, col, rule, formula) {
  const sheetId = getSheetId(sheet.name);
  if (!state.hf?.calculateFormula || sheetId == null) return undefined;
  try {
    const result = state.hf.calculateFormula(
      translateFormulaForCurrentFormulaEngine(validationFormulaForCell(formula, rule, row, col), formulaEngineSheetIndex(sheet)),
      sheetId,
    );
    return validationScalarValue(result);
  } catch (error) {
    console.warn("Data validation bound formula failed:", error);
    return undefined;
  }
}



function validationScalarValue(value) {
  if (Array.isArray(value)) return validationScalarValue(value.flat()[0]);
  if (value && typeof value === "object") return undefined;
  return value;
}



function validateCustomDataValidation(sheet, row, col, content, rule, invalid) {
  const formula = String(rule.formulae?.[0] || rule.formula || "").trim();
  if (!formula) return { valid: true };
  const sheetId = getSheetId(sheet.name);
  if (!state.hf?.calculateFormula || sheetId == null) return { valid: true };
  const address = { sheet: sheetId, row: row - 1, col: col - 1 };
  const previous = state.hf.getCellSerialized(address);
  const currentCell = sheet.cells[cellKey(row, col)] || { row, col };
  const sheetIndex = formulaEngineSheetIndex(sheet);
  const formulaValue = formulaEngineContentForCurrentModel(formulaEngineValueForCell({ ...currentCell, raw: content }, content), sheetIndex);
  try {
    state.hf.setCellContents(address, [[formulaValue]]);
    const result = state.hf.calculateFormula(
      translateFormulaForCurrentFormulaEngine(validationFormulaForCell(formula, rule, row, col), sheetIndex),
      sheetId,
    );
    return validationFormulaResultTruthy(result) ? { valid: true } : invalid();
  } catch (error) {
    console.warn("Custom data validation formula failed:", error);
    return invalid();
  } finally {
    try {
      state.hf.setCellContents(address, [[previous ?? null]]);
    } catch (restoreError) {
      console.warn("Custom data validation restore failed:", restoreError);
      state.hf = buildFormulaEngine(state.model);
    }
  }
}



function validationFormulaForCell(formula, rule, row, col) {
  const normalized = String(formula || "").trim().startsWith("=") ? String(formula || "").trim() : `=${String(formula || "").trim()}`;
  const anchor = rule?.anchor;
  if (!anchor) return normalized;
  const rowOffset = row - Number(anchor.row || row);
  const colOffset = col - Number(anchor.col || col);
  if (!rowOffset && !colOffset) return normalized;
  return translateValidationFormulaReferences(normalized, rowOffset, colOffset);
}



function translateValidationFormulaReferences(formula, rowOffset, colOffset) {
  return String(formula || "").replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/gi, (match, colFixed, letters, rowFixed, rowText, offset, source) => {
    const before = source[offset - 1] || "";
    const after = source[offset + match.length] || "";
    if (/[A-Z0-9_]/i.test(before) || /[A-Z0-9_]/i.test(after)) return match;
    const nextCol = colFixed ? columnNumber(letters) : columnNumber(letters) + colOffset;
    const nextRow = rowFixed ? Number(rowText) : Number(rowText) + rowOffset;
    if (nextCol < 1 || nextRow < 1) return match;
    return `${colFixed}${columnName(nextCol)}${rowFixed}${nextRow}`;
  });
}



function validationFormulaResultTruthy(result) {
  if (Array.isArray(result)) return validationFormulaResultTruthy(result.flat()[0]);
  if (result && typeof result === "object") return false;
  if (typeof result === "boolean") return result;
  if (typeof result === "number") return result !== 0;
  const text = String(result ?? "").trim().toUpperCase();
  return text === "TRUE" || text === "1";
}



function validationComparableValue(type, value, fromFormula = false) {
  let text = String(value ?? "").trim();
  if (fromFormula && text.startsWith("=")) text = text.slice(1).trim();
  if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
    text = text.slice(1, -1);
  }
  if (type === "textLength") {
    if (fromFormula) {
      const length = Number(text);
      return Number.isFinite(length) ? { valid: true, value: length } : { valid: false };
    }
    return { valid: true, value: String(value ?? "").length };
  }
  if (type === "date") return validationDateValue(value, text);
  if (type === "time") return validationTimeValue(value, text);

  const number = typeof value === "number" && !fromFormula ? value : Number(text.replace(/,/g, ""));
  if (!Number.isFinite(number)) return { valid: false };
  if (type === "whole" && !Number.isInteger(number)) return { valid: false };
  return { valid: true, value: number };
}



function validationDateValue(value, text) {
  if (typeof value === "number") return { valid: true, value };
  const number = Number(text.replace(/,/g, ""));
  if (Number.isFinite(number)) return { valid: true, value: number };
  const serialDay = autoFilterDateSerial(text);
  if (serialDay != null) return { valid: true, value: serialDay + 25569 };
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? { valid: false } : { valid: true, value: dateToExcelSerial(date) };
}



function validationTimeValue(value, text) {
  if (typeof value === "number") return { valid: true, value: value % 1 };
  const normalized = normalizeValidationTimeText(text);
  const number = Number(normalized);
  if (Number.isFinite(number)) return { valid: true, value: number % 1 };
  const parsed = parseValidationClockText(normalized);
  if (parsed) return parsed;
  const match = normalized.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return { valid: false };
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3] || 0);
  if (hours > 23 || minutes > 59 || seconds > 59) return { valid: false };
  return { valid: true, value: (hours * 3600 + minutes * 60 + seconds) / 86400 };
}



function normalizeValidationTimeText(text) {
  return String(text ?? "")
    .trim()
    .replace(/[：]/g, ":")
    .replace(/\s+/g, " ");
}



function parseValidationClockText(text) {
  const source = String(text ?? "").trim();
  if (!source) return null;
  const meridiemMatch = source.match(/(?:^|\s)(午前|午後|am|pm|a\.m\.|p\.m\.)(?:\s|$)/i);
  const meridiem = meridiemMatch?.[1]?.toLocaleLowerCase("en-US") || "";
  const body = source
    .replace(/(?:^|\s)(午前|午後|am|pm|a\.m\.|p\.m\.)(?:\s|$)/ig, " ")
    .replace(/\s+/g, " ")
    .trim();
  let match = body.match(/^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?$/);
  if (!match) match = body.match(/^(\d{1,2})時(?:(\d{1,2})分?)?(?:(\d{1,2})秒?)?$/);
  if (!match) return null;
  let hours = Number(match[1]);
  const minutes = Number(match[2] || 0);
  const seconds = Number(match[3] || 0);
  if (minutes > 59 || seconds > 59) return { valid: false };
  if (meridiem) {
    if (hours < 1 || hours > 12) return { valid: false };
    if (meridiem.includes("p") || meridiem === "午後") hours = hours === 12 ? 12 : hours + 12;
    if (meridiem.includes("a") || meridiem === "午前") hours = hours === 12 ? 0 : hours;
  } else if (hours > 23) {
    return { valid: false };
  }
  return { valid: true, value: (hours * 3600 + minutes * 60 + seconds) / 86400 };
}



function workbookRelationships() {
  if (!state.model) return [];
  state.model.relationships = Array.isArray(state.model.relationships) ? state.model.relationships : [];
  return state.model.relationships;
}



function cloneRelationships(relationships) {
  return (relationships || []).map((relationship) => ({
    ...relationship,
    createdAt: relationship.createdAt || "",
    updatedAt: relationship.updatedAt || "",
  }));
}



function workbookTableRefs() {
  const refs = [];
  (state.model?.sheets || []).forEach((sheet, sheetIndex) => {
    (sheet.tables || []).forEach((table, tableIndex) => {
      if (!table?.range) return;
      refs.push({
        sheet,
        sheetIndex,
        table,
        tableIndex,
        key: relationshipTableKey(sheetIndex, table, tableIndex),
        name: table.name || `Table${tableIndex + 1}`,
        label: `${table.name || `Table${tableIndex + 1}`} (${sheet.name}!${rangeToLabel(table.range)})`,
      });
    });
  });
  return refs;
}



function relationshipTableKey(sheetIndex, table, tableIndex = 0) {
  return `${sheetIndex}:${table?.id || table?.name || tableIndex}`;
}



function tableRefByKey(key) {
  return workbookTableRefs().find((ref) => ref.key === key) || null;
}



function tableColumnNames(ref) {
  if (!ref?.table?.range || !ref.sheet) return [];
  const importedNames = tableColumnNamesForStructuredRef(ref.sheet, ref.table);
  if (importedNames.length) return importedNames;
  const { range } = ref.table;
  const names = [];
  for (let col = range.left; col <= range.right; col += 1) {
    const raw = getCellRawInput(ref.sheet, range.top, col);
    names.push(String(raw || `Column${col - range.left + 1}`));
  }
  return names;
}



function openRelationshipsDialog(selectedId = "") {
  if (!state.model) return;
  state.dataDialog = { type: "relationships", selectedId };
  closeDataDialog({ keepState: true });
  const relationships = workbookRelationships();
  const canCreate = canCreateWorkbookRelationship();
  const selected = relationships.some((item) => item.id === selectedId) ? selectedId : relationships[0]?.id || "";
  const rows = relationships.length
    ? relationships.map((relationship) => `<option value="${escapeAttr(relationship.id)}"${relationship.id === selected ? " selected" : ""}>${escapeHtml(relationshipSummaryLabel(relationship))}</option>`).join("")
    : '<option value="">リレーションシップはありません</option>';
  const active = relationships.find((item) => item.id === selected) || null;
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog relationships-dialog" role="dialog" aria-modal="true" aria-labelledby="relationshipsTitle" data-data-dialog="relationships">
      <div class="excel-mini-titlebar" id="relationshipsTitle">リレーションシップの管理</div>
      <div class="excel-mini-body data-dialog-body relationships-body">
        <label class="relationships-list-wrap">
          <span class="excel-dialog-caption">リレーションシップ</span>
          <select class="relationships-list" size="8" data-data-field="relationshipList" data-relationship-list>${rows}</select>
        </label>
        <div class="relationships-details">
          <span class="excel-dialog-caption">詳細</span>
          <div class="relationships-details-box" data-relationship-details>${relationshipDetailsHtml(active)}</div>
          <div class="data-dialog-summary relationships-summary${canCreate ? "" : " is-warning"}" data-relationship-summary>${canCreate ? "テーブル間の列を選択してリレーションシップを作成できます。" : relationshipUnavailableMessage()}</div>
          <div class="relationships-actions">
            <button class="excel-mini-button" type="button" data-data-dialog-action="relationship-new" ${canCreate ? "" : "disabled"}>新規(N)...</button>
            <button class="excel-mini-button" type="button" data-data-dialog-action="relationship-edit" ${active ? "" : "disabled"}>編集(E)...</button>
            <button class="excel-mini-button" type="button" data-data-dialog-action="relationship-delete" ${active ? "" : "disabled"}>削除(D)</button>
            <button class="excel-mini-button" type="button" data-data-dialog-action="relationship-auto-detect" ${canCreate ? "" : "disabled"}>自動検出(A)</button>
          </div>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">閉じる</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("[data-relationship-list]")?.focus({ preventScroll: true });
}



function relationshipSummaryLabel(relationship) {
  const source = relationshipTableLabel(relationship.sourceTableKey);
  const target = relationshipTableLabel(relationship.targetTableKey);
  return `${source}[${relationship.sourceColumn || ""}] → ${target}[${relationship.targetColumn || ""}]`;
}



function relationshipTableLabel(key) {
  const ref = tableRefByKey(key);
  return ref?.name || String(key || "");
}



function relationshipDetailsHtml(relationship) {
  if (!relationship) return "このブックにはリレーションシップがありません。";
  const stats = relationshipColumnStats(relationship);
  const lines = [
    `テーブル: ${relationshipTableLabel(relationship.sourceTableKey)}`,
    `列: ${relationship.sourceColumn || ""}`,
    `関連テーブル: ${relationshipTableLabel(relationship.targetTableKey)}`,
    `関連列: ${relationship.targetColumn || ""}`,
    `カーディナリティ: ${relationshipCardinalityLabel(stats)}`,
    `元データ行数: ${stats.sourceRows.toLocaleString("ja-JP")}`,
    `関連テーブル行数: ${stats.targetRows.toLocaleString("ja-JP")}`,
    `一致しない値: ${stats.unmatched.toLocaleString("ja-JP")}`,
  ];
  if (relationship.updatedAt || relationship.createdAt) lines.push(`更新日時: ${formatQueryTimestamp(relationship.updatedAt || relationship.createdAt)}`);
  return escapeHtml(lines.join("\n"));
}



function relationshipCardinalityLabel(stats) {
  const sourceUnique = stats.sourceRows > 0 && stats.sourceDistinct >= stats.sourceRows - stats.sourceBlank;
  const targetUnique = stats.targetRows > 0 && stats.targetDistinct >= stats.targetRows - stats.targetBlank;
  if (sourceUnique && targetUnique) return "一対一";
  if (targetUnique) return "多対一";
  if (sourceUnique) return "一対多";
  return "多対多";
}



function handleRelationshipsDialogSync(event) {
  const dialog = event.target.closest("[data-data-dialog='relationships']");
  if (!dialog) return;
  const relationship = selectedRelationshipFromDialog(dialog);
  const canCreate = canCreateWorkbookRelationship();
  const details = dialog.querySelector("[data-relationship-details]");
  if (details) details.innerHTML = relationshipDetailsHtml(relationship);
  const summary = dialog.querySelector("[data-relationship-summary]");
  if (summary) {
    summary.textContent = canCreate ? "テーブル間の列を選択してリレーションシップを作成できます。" : relationshipUnavailableMessage();
    summary.classList.toggle("is-warning", !canCreate);
  }
  dialog.querySelector("[data-data-dialog-action='relationship-new']").disabled = !canCreate;
  dialog.querySelector("[data-data-dialog-action='relationship-edit']").disabled = !relationship;
  dialog.querySelector("[data-data-dialog-action='relationship-delete']").disabled = !relationship;
  dialog.querySelector("[data-data-dialog-action='relationship-auto-detect']").disabled = !canCreate;
}



function selectedRelationshipFromDialog(dialog) {
  const id = dialog?.querySelector("[data-relationship-list]")?.value || "";
  return workbookRelationships().find((item) => item.id === id) || null;
}



function openRelationshipEditDialog(relationshipId = "") {
  const tables = workbookTableRefs();
  if (tables.length < 2) {
    showMessageBox(relationshipUnavailableMessage(), { title: "Cht WebSheet" });
    return;
  }
  const existing = workbookRelationships().find((item) => item.id === relationshipId) || null;
  const sourceKey = existing?.sourceTableKey || tables[1]?.key || tables[0].key;
  const targetKey = existing?.targetTableKey || tables.find((table) => table.key !== sourceKey)?.key || tables[0].key;
  const sourceColumns = tableColumnNames(tableRefByKey(sourceKey));
  const targetColumns = tableColumnNames(tableRefByKey(targetKey));
  state.dataDialog = { type: "relationship-edit", relationshipId };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog relationship-edit-dialog" role="dialog" aria-modal="true" aria-labelledby="relationshipEditTitle" data-data-dialog="relationship-edit">
      <div class="excel-mini-titlebar" id="relationshipEditTitle">${existing ? "リレーションシップの編集" : "リレーションシップの作成"}</div>
      <div class="excel-mini-body data-dialog-body relationship-edit-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">テーブル</span>
          <select data-data-field="sourceTableKey" data-relationship-edit-field>${relationshipTableOptionsHtml(sourceKey)}</select>
        </label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">列</span>
          <select data-data-field="sourceColumn" data-relationship-edit-field>${relationshipColumnOptionsHtml(sourceColumns, existing?.sourceColumn || sourceColumns[0] || "")}</select>
        </label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">関連テーブル</span>
          <select data-data-field="targetTableKey" data-relationship-edit-field>${relationshipTableOptionsHtml(targetKey)}</select>
        </label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">関連列</span>
          <select data-data-field="targetColumn" data-relationship-edit-field>${relationshipColumnOptionsHtml(targetColumns, existing?.targetColumn || targetColumns[0] || "")}</select>
        </label>
        <div class="data-dialog-note" data-relationship-edit-note>${escapeHtml(relationshipCardinalityNote(sourceKey, existing?.sourceColumn || sourceColumns[0] || "", targetKey, existing?.targetColumn || targetColumns[0] || ""))}</div>
        <div class="data-dialog-summary relationship-edit-summary" data-relationship-edit-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="relationship-ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="relationship-cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncRelationshipEditDialog(overlay.querySelector("[data-data-dialog='relationship-edit']"));
  overlay.querySelector("[data-data-field='sourceTableKey']")?.focus({ preventScroll: true });
}



function relationshipTableOptionsHtml(selectedKey = "") {
  return workbookTableRefs()
    .map((ref) => `<option value="${escapeAttr(ref.key)}"${ref.key === selectedKey ? " selected" : ""}>${escapeHtml(ref.label)}</option>`)
    .join("");
}



function relationshipColumnOptionsHtml(columns, selectedColumn = "") {
  return (columns || [])
    .map((column) => `<option value="${escapeAttr(column)}"${column === selectedColumn ? " selected" : ""}>${escapeHtml(column)}</option>`)
    .join("");
}



function handleRelationshipEditDialogSync(event) {
  const dialog = event.target.closest("[data-data-dialog='relationship-edit']");
  if (!dialog) return;
  const field = event.target.dataset.dataField;
  if (field === "sourceTableKey" || field === "targetTableKey") {
    const columnField = dialog.querySelector(`[data-data-field='${field === "sourceTableKey" ? "sourceColumn" : "targetColumn"}']`);
    const columns = tableColumnNames(tableRefByKey(event.target.value));
    if (columnField) columnField.innerHTML = relationshipColumnOptionsHtml(columns, columns[0] || "");
  }
  const note = dialog.querySelector("[data-relationship-edit-note]");
  if (note) {
    note.textContent = relationshipCardinalityNote(
      dialog.querySelector("[data-data-field='sourceTableKey']")?.value || "",
      dialog.querySelector("[data-data-field='sourceColumn']")?.value || "",
      dialog.querySelector("[data-data-field='targetTableKey']")?.value || "",
      dialog.querySelector("[data-data-field='targetColumn']")?.value || "",
    );
  }
  syncRelationshipEditDialog(dialog);
}



function syncRelationshipEditDialog(dialog) {
  if (!dialog) return;
  const definition = relationshipDefinitionFromDialog(dialog);
  const validation = validateRelationshipDefinition(definition, state.dataDialog?.relationshipId || "");
  const summary = dialog.querySelector("[data-relationship-edit-summary]");
  if (summary) {
    summary.textContent = validation.message || relationshipPreviewMessage(definition);
    summary.classList.toggle("is-warning", !validation.valid || Boolean(validation.warning));
  }
  const ok = dialog.querySelector("[data-data-dialog-action='relationship-ok']");
  if (ok) ok.disabled = Boolean(validation.disableOk);
}



function relationshipDefinitionFromDialog(dialog) {
  return {
    sourceTableKey: dialog?.querySelector("[data-data-field='sourceTableKey']")?.value || "",
    sourceColumn: dialog?.querySelector("[data-data-field='sourceColumn']")?.value || "",
    targetTableKey: dialog?.querySelector("[data-data-field='targetTableKey']")?.value || "",
    targetColumn: dialog?.querySelector("[data-data-field='targetColumn']")?.value || "",
  };
}



function relationshipCardinalityNote(sourceKey, sourceColumn, targetKey, targetColumn) {
  const sourceCount = tableDistinctValueCount(sourceKey, sourceColumn);
  const targetCount = tableDistinctValueCount(targetKey, targetColumn);
  if (!sourceColumn || !targetColumn) return "列を選択してください。";
  return `一意の値: ${relationshipTableLabel(sourceKey)}[${sourceColumn}] = ${sourceCount} / ${relationshipTableLabel(targetKey)}[${targetColumn}] = ${targetCount}`;
}



function relationshipPreviewMessage(definition) {
  const stats = relationshipColumnStats(definition);
  const sourceLabel = `${relationshipTableLabel(definition.sourceTableKey)}[${definition.sourceColumn || ""}]`;
  const targetLabel = `${relationshipTableLabel(definition.targetTableKey)}[${definition.targetColumn || ""}]`;
  const unmatchedText = stats.unmatched ? ` 一致しない値: ${stats.unmatched.toLocaleString("ja-JP")}。` : "";
  const blankText = stats.sourceBlank || stats.targetBlank
    ? ` 空白: 元 ${stats.sourceBlank.toLocaleString("ja-JP")} / 関連 ${stats.targetBlank.toLocaleString("ja-JP")}。`
    : "";
  return `${sourceLabel} から ${targetLabel} へのリレーションシップを作成します。元テーブル ${stats.sourceRows.toLocaleString("ja-JP")} 行、関連テーブル ${stats.targetRows.toLocaleString("ja-JP")} 行。${unmatchedText}${blankText}`;
}



function relationshipColumnStats(definition) {
  const sourceRef = tableRefByKey(definition.sourceTableKey);
  const targetRef = tableRefByKey(definition.targetTableKey);
  const sourceValues = tableColumnValues(sourceRef, definition.sourceColumn);
  const targetValues = tableColumnValues(targetRef, definition.targetColumn);
  const targetSet = new Set(targetValues.filter((value) => value !== ""));
  return {
    sourceRows: sourceValues.length,
    targetRows: targetValues.length,
    sourceDistinct: new Set(sourceValues.filter((value) => value !== "")).size,
    targetDistinct: new Set(targetValues.filter((value) => value !== "")).size,
    sourceBlank: sourceValues.filter((value) => value === "").length,
    targetBlank: targetValues.filter((value) => value === "").length,
    unmatched: sourceValues.filter((value) => value !== "" && !targetSet.has(value)).length,
    sourceType: inferRelationshipColumnType(sourceValues),
    targetType: inferRelationshipColumnType(targetValues),
  };
}



function tableColumnValues(ref, columnNameValue) {
  if (!ref?.table?.range || !columnNameValue) return [];
  const columns = tableColumnNames(ref);
  const offset = columns.indexOf(columnNameValue);
  if (offset < 0) return [];
  const values = [];
  const sheetId = getSheetId(ref.sheet.name);
  for (let row = ref.table.range.top + 1; row <= ref.table.range.bottom; row += 1) {
    values.push(String(getDisplayForCell(ref.sheet, sheetId, row, ref.table.range.left + offset) ?? "").trim());
  }
  return values;
}



function inferRelationshipColumnType(values) {
  const nonBlank = (values || []).filter((value) => String(value ?? "") !== "");
  if (!nonBlank.length) return "blank";
  const numericCount = nonBlank.filter((value) => Number.isFinite(Number(String(value).replace(/,/g, "")))).length;
  const dateCount = nonBlank.filter((value) => autoFilterDateSerial(value) != null || !Number.isNaN(new Date(value).getTime())).length;
  if (numericCount / nonBlank.length >= 0.85) return "number";
  if (dateCount / nonBlank.length >= 0.85) return "date";
  return "text";
}



function tableDistinctValueCount(tableKey, columnNameValue) {
  const ref = tableRefByKey(tableKey);
  if (!ref?.table?.range || !columnNameValue) return 0;
  const columns = tableColumnNames(ref);
  const offset = columns.indexOf(columnNameValue);
  if (offset < 0) return 0;
  const values = new Set();
  const sheetId = getSheetId(ref.sheet.name);
  for (let row = ref.table.range.top + 1; row <= ref.table.range.bottom; row += 1) {
    const value = getDisplayForCell(ref.sheet, sheetId, row, ref.table.range.left + offset);
    if (String(value ?? "") !== "") values.add(String(value));
  }
  return values.size;
}



function tableDataRowCount(tableKey) {
  const range = tableRefByKey(tableKey)?.table?.range;
  if (!range) return 0;
  return Math.max(0, range.bottom - range.top);
}



function canCreateWorkbookRelationship() {
  return workbookTableRefs().length >= 2;
}



function relationshipUnavailableMessage() {
  return "リレーションシップを作成するには、2つ以上のテーブルが必要です。";
}



function normalizeRelationshipColumnName(name) {
  return String(name || "").trim().toLocaleLowerCase("ja-JP");
}



function canonicalRelationshipColumnName(name) {
  return normalizeRelationshipColumnName(name)
    .replace(/[\s_\-./\\()[\]{}"'`~:;、。，．]+/g, "")
    .replace(/（.*?）|\(.*?\)/g, "")
    .replace(/^(fk|pk)/, "")
    .replace(/(id|ｉｄ|コード|code|番号|番號|no|№|キー|key)$/i, "");
}



function relationshipColumnNameTokens(name) {
  const normalized = normalizeRelationshipColumnName(name)
    .replace(/（.*?）|\(.*?\)/g, " ")
    .replace(/[_\-./\\()[\]{}"'`~:;、。，．]+/g, " ");
  return normalized.split(/\s+/).map((token) => token.trim()).filter(Boolean);
}



function relationshipColumnMatchScore(leftColumn, rightColumn, leftTable = null, rightTable = null) {
  const left = normalizeRelationshipColumnName(leftColumn);
  const right = normalizeRelationshipColumnName(rightColumn);
  if (!left || !right) return 0;
  if (left === right) return 100;
  const leftCanonical = canonicalRelationshipColumnName(leftColumn);
  const rightCanonical = canonicalRelationshipColumnName(rightColumn);
  if (leftCanonical && rightCanonical && leftCanonical === rightCanonical) return 92;
  const leftTokens = new Set(relationshipColumnNameTokens(leftColumn));
  const rightTokens = new Set(relationshipColumnNameTokens(rightColumn));
  const common = [...leftTokens].filter((token) => rightTokens.has(token));
  if (common.length && (leftTokens.has("id") || leftTokens.has("コード") || rightTokens.has("id") || rightTokens.has("コード"))) {
    return 86;
  }
  const leftTableName = canonicalRelationshipColumnName(leftTable?.name || leftTable?.table?.name || "");
  const rightTableName = canonicalRelationshipColumnName(rightTable?.name || rightTable?.table?.name || "");
  if (leftCanonical && rightCanonical) {
    if ((leftCanonical === rightTableName && ["id", "コード", "key", "キー", "no", "番号"].includes(right)) ||
      (rightCanonical === leftTableName && ["id", "コード", "key", "キー", "no", "番号"].includes(left))) {
      return 82;
    }
    if (leftCanonical.endsWith(rightCanonical) || rightCanonical.endsWith(leftCanonical)) return 76;
  }
  return 0;
}



function bestRelationshipColumnMatch(leftColumn, rightColumns, leftTable, rightTable) {
  return (rightColumns || [])
    .map((rightColumn) => ({
      rightColumn,
      score: relationshipColumnMatchScore(leftColumn, rightColumn, leftTable, rightTable),
    }))
    .filter((candidate) => candidate.score >= 76)
    .sort((a, b) => b.score - a.score || String(a.rightColumn).localeCompare(String(b.rightColumn), "ja-JP"))[0] || null;
}



function relationshipDefinitionExists(definition, relationships = workbookRelationships(), editingId = "") {
  return relationships.some((relationship) => {
    if (relationship.id && relationship.id === editingId) return false;
    const sameDirection =
      relationship.sourceTableKey === definition.sourceTableKey &&
      relationship.sourceColumn === definition.sourceColumn &&
      relationship.targetTableKey === definition.targetTableKey &&
      relationship.targetColumn === definition.targetColumn;
    const reverseDirection =
      relationship.sourceTableKey === definition.targetTableKey &&
      relationship.sourceColumn === definition.targetColumn &&
      relationship.targetTableKey === definition.sourceTableKey &&
      relationship.targetColumn === definition.sourceColumn;
    return sameDirection || reverseDirection;
  });
}



function autoDetectRelationshipDefinitions() {
  const tables = workbookTableRefs();
  const existing = workbookRelationships();
  const definitions = [];
  for (let leftIndex = 0; leftIndex < tables.length; leftIndex += 1) {
    for (let rightIndex = leftIndex + 1; rightIndex < tables.length; rightIndex += 1) {
      const left = tables[leftIndex];
      const right = tables[rightIndex];
      const leftColumns = tableColumnNames(left);
      const rightColumns = tableColumnNames(right);
      leftColumns.forEach((leftColumn) => {
        const rightColumn = bestRelationshipColumnMatch(leftColumn, rightColumns, left, right)?.rightColumn || "";
        if (!rightColumn) return;
        const leftRows = tableDataRowCount(left.key);
        const rightRows = tableDataRowCount(right.key);
        const leftUnique = tableDistinctValueCount(left.key, leftColumn);
        const rightUnique = tableDistinctValueCount(right.key, rightColumn);
        const leftLooksLookup = leftRows > 0 && leftUnique >= leftRows;
        const rightLooksLookup = rightRows > 0 && rightUnique >= rightRows;
        let definition;
        if (leftLooksLookup && !rightLooksLookup) {
          definition = { sourceTableKey: right.key, sourceColumn: rightColumn, targetTableKey: left.key, targetColumn: leftColumn };
        } else if (rightLooksLookup && !leftLooksLookup) {
          definition = { sourceTableKey: left.key, sourceColumn: leftColumn, targetTableKey: right.key, targetColumn: rightColumn };
        } else if (leftRows <= rightRows) {
          definition = { sourceTableKey: right.key, sourceColumn: rightColumn, targetTableKey: left.key, targetColumn: leftColumn };
        } else {
          definition = { sourceTableKey: left.key, sourceColumn: leftColumn, targetTableKey: right.key, targetColumn: rightColumn };
        }
        if (!relationshipDefinitionExists(definition, existing) && !relationshipDefinitionExists(definition, definitions)) {
          definitions.push(definition);
        }
      });
    }
  }
  return definitions;
}



function applyRelationshipEditDialog(dialog) {
  const definition = relationshipDefinitionFromDialog(dialog);
  const validation = validateRelationshipDefinition(definition, state.dataDialog?.relationshipId || "");
  if (!validation.valid) {
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    syncRelationshipEditDialog(dialog);
    return;
  }
  const { sourceTableKey, sourceColumn, targetTableKey, targetColumn } = definition;
  const relationships = workbookRelationships();
  const before = cloneRelationships(relationships);
  const id = state.dataDialog?.relationshipId || uniqueWorksheetObjectId("relationship");
  const now = new Date().toISOString();
  const existingIndex = relationships.findIndex((item) => item.id === id);
  const next = {
    id,
    sourceTableKey,
    sourceColumn,
    targetTableKey,
    targetColumn,
    createdAt: existingIndex >= 0 ? relationships[existingIndex].createdAt || now : now,
    updatedAt: now,
  };
  if (existingIndex >= 0) relationships[existingIndex] = next;
  else relationships.push(next);
  const after = cloneRelationships(relationships);
  closeDataDialog();
  pushRelationshipHistory(before, after, existingIndex >= 0 ? "リレーションシップの編集" : "リレーションシップの作成");
  openRelationshipsDialog(id);
  setStatus(`${relationshipSummaryLabel(next)} を${existingIndex >= 0 ? "更新" : "作成"}しました。`);
}



function validateRelationshipDefinition(definition, editingId = "") {
  if (!definition.sourceTableKey || !definition.targetTableKey || !definition.sourceColumn || !definition.targetColumn) {
    return { valid: false, disableOk: true, message: "テーブルと列を選択してください。" };
  }
  if (definition.sourceTableKey === definition.targetTableKey) {
    return { valid: false, disableOk: true, message: "同じテーブル間にはリレーションシップを作成できません。" };
  }
  const sourceColumns = tableColumnNames(tableRefByKey(definition.sourceTableKey));
  const targetColumns = tableColumnNames(tableRefByKey(definition.targetTableKey));
  if (!sourceColumns.includes(definition.sourceColumn) || !targetColumns.includes(definition.targetColumn)) {
    return { valid: false, disableOk: true, message: "選択した列がテーブルに存在しません。" };
  }
  if (relationshipDefinitionExists(definition, workbookRelationships(), editingId)) {
    return { valid: false, message: "同じリレーションシップが既に存在します。" };
  }
  const stats = relationshipColumnStats(definition);
  if (stats.sourceType !== "blank" && stats.targetType !== "blank" && stats.sourceType !== stats.targetType) {
    return { valid: true, warning: true, message: `列のデータ型が一致していません。元: ${relationshipColumnTypeLabel(stats.sourceType)} / 関連: ${relationshipColumnTypeLabel(stats.targetType)}。` };
  }
  if (stats.targetRows > 0 && stats.targetDistinct < stats.targetRows - stats.targetBlank) {
    return { valid: true, warning: true, message: "関連列に重複値があります。Excel では多対多として扱われる可能性があります。" };
  }
  if (stats.unmatched > 0) {
    return { valid: true, warning: true, message: `${stats.unmatched.toLocaleString("ja-JP")} 個の元データ値が関連列に見つかりません。` };
  }
  return { valid: true, message: relationshipPreviewMessage(definition) };
}



function relationshipColumnTypeLabel(type) {
  return {
    number: "数値",
    date: "日付",
    text: "テキスト",
    blank: "空白",
  }[type] || "テキスト";
}



function deleteRelationshipFromDialog(dialog) {
  const relationship = selectedRelationshipFromDialog(dialog);
  if (!relationship) return;
  showMessageBox(`リレーションシップ「${relationshipSummaryLabel(relationship)}」を削除しますか?`, {
    title: "Cht WebSheet",
    tone: "warning",
    buttons: [
      {
        label: "削除",
        result: "delete",
        callback: () => {
          const before = cloneRelationships(workbookRelationships());
          state.model.relationships = workbookRelationships().filter((item) => item.id !== relationship.id);
          const after = cloneRelationships(workbookRelationships());
          pushRelationshipHistory(before, after, "リレーションシップの削除");
          openRelationshipsDialog();
          setStatus("リレーションシップを削除しました。");
        },
      },
      { label: "キャンセル", result: "cancel" },
    ],
    cancelResult: "cancel",
  });
}



function autoDetectRelationshipsFromDialog() {
  const tables = workbookTableRefs();
  if (tables.length < 2) {
    showMessageBox("リレーションシップを検出するには、2つ以上のテーブルが必要です。", { title: "Cht WebSheet" });
    return;
  }
  const definitions = autoDetectRelationshipDefinitions();
  if (!definitions.length) {
    showMessageBox("検出できる新しいリレーションシップはありません。", { title: "Cht WebSheet" });
    return;
  }
  const validDefinitions = definitions.filter((definition) => validateRelationshipDefinition(definition).valid);
  if (validDefinitions.length === 1) {
    applyRelationshipDefinitions(validDefinitions, "リレーションシップの自動検出");
    return;
  }
  openRelationshipAutoDetectDialog(definitions);
}



function openRelationshipAutoDetectDialog(definitions) {
  state.dataDialog = { type: "relationship-auto-detect", definitions: definitions.map((definition) => ({ ...definition })) };
  closeDataDialog({ keepState: true });
  const rows = definitions.map((definition, index) => {
    const validation = validateRelationshipDefinition(definition);
    const warning = validation.warning ? " is-warning" : "";
    return `<label class="relationship-auto-row${warning}">
      <input type="checkbox" data-relationship-auto-index="${index}" checked />
      <span>
        <strong>${escapeHtml(relationshipDefinitionLabel(definition))}</strong>
        <small>${escapeHtml(validation.message || relationshipPreviewMessage(definition))}</small>
      </span>
    </label>`;
  }).join("");
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog relationship-auto-dialog" role="dialog" aria-modal="true" aria-labelledby="relationshipAutoTitle" data-data-dialog="relationship-auto-detect">
      <div class="excel-mini-titlebar" id="relationshipAutoTitle">リレーションシップの自動検出</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-note">検出されたリレーションシップを確認し、作成するものを選択してください。</div>
        <div class="relationship-auto-list">${rows}</div>
        <div class="data-dialog-summary" data-relationship-auto-summary>${definitions.length.toLocaleString("ja-JP")} 件の候補があります。</div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="relationship-auto-ok">作成</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="relationship-auto-back">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
}



function relationshipDefinitionLabel(definition) {
  return `${relationshipTableLabel(definition.sourceTableKey)}[${definition.sourceColumn || ""}] → ${relationshipTableLabel(definition.targetTableKey)}[${definition.targetColumn || ""}]`;
}



function handleRelationshipAutoDetectSync(event) {
  syncRelationshipAutoDetectDialog(event.target.closest("[data-data-dialog='relationship-auto-detect']"));
}



function syncRelationshipAutoDetectDialog(dialog) {
  if (!dialog) return;
  const total = dialog.querySelectorAll("[data-relationship-auto-index]").length;
  const selected = dialog.querySelectorAll("[data-relationship-auto-index]:checked").length;
  const summary = dialog.querySelector("[data-relationship-auto-summary]");
  if (summary) {
    summary.textContent = `${total.toLocaleString("ja-JP")} 件の候補中 ${selected.toLocaleString("ja-JP")} 件を作成します。`;
    summary.classList.toggle("is-warning", selected === 0);
  }
  const ok = dialog.querySelector("[data-data-dialog-action='relationship-auto-ok']");
  if (ok) ok.disabled = selected === 0;
}



function applyRelationshipAutoDetectDialog(dialog) {
  const definitions = state.dataDialog?.definitions || [];
  const selected = Array.from(dialog.querySelectorAll("[data-relationship-auto-index]:checked"))
    .map((input) => definitions[Number(input.dataset.relationshipAutoIndex)])
    .filter(Boolean)
    .filter((definition) => validateRelationshipDefinition(definition).valid);
  if (!selected.length) {
    showMessageBox("作成するリレーションシップを選択してください。", { title: "Cht WebSheet" });
    return;
  }
  applyRelationshipDefinitions(selected, "リレーションシップの自動検出");
}



function applyRelationshipDefinitions(selected, historyLabel) {
  const relationships = workbookRelationships();
  const before = cloneRelationships(relationships);
  const now = new Date().toISOString();
  const added = selected.map((definition) => ({
    id: uniqueWorksheetObjectId("relationship"),
    ...definition,
    createdAt: now,
    updatedAt: now,
  }));
  relationships.push(...added);
  const after = cloneRelationships(relationships);
  pushRelationshipHistory(before, after, historyLabel || "リレーションシップの作成");
  closeDataDialog();
  openRelationshipsDialog(added[0].id);
  setStatus(`${added.length} 件のリレーションシップを検出しました。`);
}



function pushRelationshipHistory(before, after, label) {
  const range = activeSelectionRange();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    relationshipsBefore: before,
    relationshipsAfter: after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
    selectionAfterRanges: activeSelectionRanges().map((item) => ({ ...item })),
    label,
  });
  scheduleWorkbookPersistence();
}



function dataSummaryFunctionOptions(selected = "sum") {
  return DATA_SUMMARY_FUNCTIONS.map((item) => `<option value="${item.value}" ${item.value === selected ? "selected" : ""}>${item.label}</option>`).join("");
}



function dataSummaryFunctionLabel(value) {
  return DATA_SUMMARY_FUNCTIONS.find((item) => item.value === value)?.label || "合計";
}



function dataSummarySubtotalCode(value) {
  return DATA_SUMMARY_FUNCTIONS.find((item) => item.value === value)?.subtotalCode || 9;
}



function applySummaryFunction(values, method) {
  const numeric = values.map((value) => Number(value)).filter((value) => Number.isFinite(value));
  switch (method) {
    case "count":
      return values.filter((value) => value != null && String(value) !== "").length;
    case "countNums":
      return numeric.length;
    case "average":
      return numeric.length ? numeric.reduce((total, value) => total + value, 0) / numeric.length : null;
    case "max":
      return numeric.length ? Math.max(...numeric) : null;
    case "min":
      return numeric.length ? Math.min(...numeric) : null;
    case "product":
      return numeric.length ? numeric.reduce((total, value) => total * value, 1) : null;
    case "stdDev": {
      if (numeric.length < 2) return null;
      const avg = numeric.reduce((total, value) => total + value, 0) / numeric.length;
      return Math.sqrt(numeric.reduce((total, value) => total + (value - avg) ** 2, 0) / (numeric.length - 1));
    }
    case "stdDevP": {
      if (!numeric.length) return null;
      const avg = numeric.reduce((total, value) => total + value, 0) / numeric.length;
      return Math.sqrt(numeric.reduce((total, value) => total + (value - avg) ** 2, 0) / numeric.length);
    }
    case "var": {
      if (numeric.length < 2) return null;
      const avg = numeric.reduce((total, value) => total + value, 0) / numeric.length;
      return numeric.reduce((total, value) => total + (value - avg) ** 2, 0) / (numeric.length - 1);
    }
    case "varP": {
      if (!numeric.length) return null;
      const avg = numeric.reduce((total, value) => total + value, 0) / numeric.length;
      return numeric.reduce((total, value) => total + (value - avg) ** 2, 0) / numeric.length;
    }
    case "sum":
    default:
      return numeric.length ? numeric.reduce((total, value) => total + value, 0) : null;
  }
}



function openConsolidateDialog() {
  const context = selectedDataRange("統合");
  if (!context) return;
  const reference = workbookRangeReference(state.activeSheetIndex, context.range);
  state.dataDialog = { type: "consolidate", range: context.range, references: [{ ...context.range }] };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog consolidate-dialog" role="dialog" aria-modal="true" aria-labelledby="consolidateDialogTitle" data-data-dialog="consolidate">
      <div class="excel-mini-titlebar" id="consolidateDialogTitle">統合</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">集計の方法</span><select data-data-field="function" data-consolidate-sync>${dataSummaryFunctionOptions("sum")}</select></label>
        ${dataRangeFieldHtml("参照範囲", "reference", reference, "consolidate")}
        <div class="data-dialog-inline-actions">
          <button class="excel-mini-button" type="button" data-data-dialog-action="add-reference">追加</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="delete-reference">削除</button>
        </div>
        <select class="data-reference-list" data-data-field="references" size="4" data-consolidate-sync>
          <option value="${escapeAttr(reference)}" selected>${escapeHtml(reference)}</option>
        </select>
        <fieldset class="data-dialog-fieldset">
          <legend>統合の基準</legend>
          <label class="data-dialog-check"><input type="checkbox" data-data-field="topRow" data-consolidate-sync checked />上端行</label>
          <label class="data-dialog-check"><input type="checkbox" data-data-field="leftColumn" data-consolidate-sync />左端列</label>
        </fieldset>
        <label class="data-dialog-check"><input type="checkbox" data-data-field="createLinks" data-consolidate-sync />統合元データとリンクする</label>
        <div class="data-dialog-summary consolidate-summary" data-consolidate-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncConsolidateDialog(overlay.querySelector("[data-data-dialog='consolidate']"));
}



function handleConsolidateDialogSync(event) {
  syncConsolidateDialog(event.currentTarget.closest("[data-data-dialog='consolidate']"));
}



function syncConsolidateDialog(dialog) {
  if (!dialog) return;
  const summary = dialog.querySelector("[data-consolidate-summary]");
  const preview = consolidateDialogPreview(dialog);
  if (summary) {
    summary.textContent = preview.message;
    summary.classList.toggle("is-warning", preview.warning || preview.blocked);
  }
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = Boolean(preview.blocked);
  const referenceInput = dialog.querySelector("[data-data-field='reference']");
  const addButton = dialog.querySelector("[data-data-dialog-action='add-reference']");
  if (addButton) addButton.disabled = !decodeWorkbookRangeOrName(referenceInput?.value || "", activeSheet());
  const deleteButton = dialog.querySelector("[data-data-dialog-action='delete-reference']");
  if (deleteButton) deleteButton.disabled = !dialog.querySelector("[data-data-field='references'] option:checked");
}



function sameSheetRangeIntersects(a, b) {
  if (!a || !b) return false;
  const aSheetIndex = a.sheetIndex ?? state.activeSheetIndex;
  const bSheetIndex = b.sheetIndex ?? state.activeSheetIndex;
  return aSheetIndex === bSheetIndex && rangesIntersect(a, b);
}



function consolidateDialogReferences(dialog) {
  return Array.from(dialog?.querySelectorAll("[data-data-field='references'] option") || [])
    .map((option) => ({ text: option.value || option.textContent || "", range: decodeWorkbookRangeOrName(option.value || option.textContent || "", activeSheet()) }))
    .map((item) => (item.range ? { ...item.range, sheetIndex: item.range.sheetIndex ?? state.activeSheetIndex, sourceText: item.text } : { sourceText: item.text, invalid: true }));
}



function consolidateDialogOptions(dialog) {
  return {
    method: dialog.querySelector("[data-data-field='function']")?.value || "sum",
    topRow: Boolean(dialog.querySelector("[data-data-field='topRow']")?.checked),
    leftColumn: Boolean(dialog.querySelector("[data-data-field='leftColumn']")?.checked),
    createLinks: Boolean(dialog.querySelector("[data-data-field='createLinks']")?.checked),
  };
}



function consolidateTargetValidation(sheet, range) {
  if (!sheet || !range) return { valid: false, message: "出力先を選択してください。" };
  if (rangeHasLockedCellsForProtection(sheet, range)) {
    return { valid: false, message: "出力先に保護されたロック済みセルが含まれています。" };
  }
  if (mergesIntersectingRange(sheet, range).length) {
    return { valid: false, message: "統合結果は結合されたセルを含む範囲には出力できません。" };
  }
  return { valid: true, message: "" };
}



function consolidateDialogPreview(dialog) {
  const references = consolidateDialogReferences(dialog);
  if (!references.length) return { warning: true, blocked: true, message: "参照範囲を追加してください。" };
  const invalid = references.filter((item) => item.invalid).map((item) => item.sourceText).filter(Boolean);
  if (invalid.length) return { warning: true, blocked: true, message: `参照範囲を正しく入力してください: ${invalid.join(", ")}` };
  const options = consolidateDialogOptions(dialog);
  const destinationRange = state.dataDialog?.range;
  const destinationSheet = activeSheet();
  if (!destinationRange || !destinationSheet) return { warning: true, blocked: true, message: "出力先を選択してください。" };
  const headers = consolidateHeaders(references, options);
  const dataCells = references.reduce((total, range) => {
    const rows = Math.max(0, range.bottom - range.top + 1 - (options.topRow ? 1 : 0));
    const cols = Math.max(0, range.right - range.left + 1 - (options.leftColumn ? 1 : 0));
    return total + rows * cols;
  }, 0);
  if (!dataCells) return { warning: true, blocked: true, message: "統合する値が参照範囲にありません。" };
  const placement = consolidateOutputPlacement(destinationSheet, destinationRange, references, options, headers);
  const targetValidation = consolidateTargetValidation(destinationSheet, placement.range);
  if (!targetValidation.valid) return { warning: true, blocked: true, message: targetValidation.message };
  const overwrite = placement?.range && rangeHasUserData(destinationSheet, placement.range);
  const overlapsSource = references.some((reference) => sameSheetRangeIntersects(placement?.range, reference));
  const linkText = options.createLinks ? " 統合元へのリンクを作成します。" : "";
  const overwriteText = overwrite ? " 出力先には既存のデータがあります。OK時に確認します。" : "";
  const overlapText = overlapsSource ? " 出力先が参照範囲と重なっています。" : "";
  return {
    warning: Boolean(overwrite || overlapsSource),
    message: `${references.length} 件の参照範囲、${dataCells} 個の値を ${dataSummaryFunctionLabel(options.method)} で統合します。出力先: ${placement?.range ? rangeToLabel(placement.range) : "選択範囲"}、${placement?.height || 0} 行 x ${placement?.width || 0} 列。${linkText}${overwriteText}${overlapText}`.trim(),
  };
}



function consolidateOutputPlacement(sheet, destinationRange, references, options, headers = consolidateHeaders(references, options), sheetIndex = state.activeSheetIndex) {
  const rowLabels = options.leftColumn ? consolidateLeftColumnLabels(references, options).length : 1;
  const height = Math.max(1, rowLabels) + 1;
  const width = headers.length + 1;
  const sameSheetReferences = (references || []).filter((reference) => (reference.sheetIndex ?? state.activeSheetIndex) === sheetIndex);
  const overlapsSource = sameSheetReferences.some((reference) => rangesIntersect(destinationRange, reference));
  const top = overlapsSource
    ? Math.max(destinationRange.bottom, ...sameSheetReferences.map((reference) => reference.bottom || destinationRange.bottom)) + 2
    : destinationRange.top;
  const left = destinationRange.left;
  const range = normalizeRange(
    { sheetIndex, row: top, col: left },
    { sheetIndex, row: top + height - 1, col: left + width - 1 },
  );
  return { top, left, width, height, range, shifted: overlapsSource };
}



function consolidateLeftColumnLabels(references, options) {
  const labels = [];
  const seen = new Set();
  references.forEach((range) => {
    const sourceSheet = state.model?.sheets?.[range.sheetIndex ?? state.activeSheetIndex];
    if (!sourceSheet) return;
    const sourceSheetId = getSheetId(sourceSheet.name);
    const dataStart = range.top + (options.topRow ? 1 : 0);
    for (let row = dataStart; row <= range.bottom; row += 1) {
      const label = String(getDisplayForCell(sourceSheet, sourceSheetId, row, range.left) || "(空白)");
      const key = normalizeAdvancedFilterHeader(label);
      if (seen.has(key)) continue;
      seen.add(key);
      labels.push(label);
    }
  });
  return labels;
}



function applyConsolidateDialog(dialog) {
  const context = state.dataDialog;
  const sheet = activeSheet();
  if (!context?.range || !sheet) return;
  const preview = consolidateDialogPreview(dialog);
  if (preview.blocked) {
    showMessageBox(preview.message || "統合の設定を正しく入力してください。", { title: "Cht WebSheet" });
    syncConsolidateDialog(dialog);
    return;
  }
  const referenceItems = consolidateDialogReferences(dialog);
  const invalid = referenceItems.filter((item) => item.invalid).map((item) => item.sourceText).filter(Boolean);
  if (invalid.length) {
    showMessageBox(`参照範囲を正しく入力してください。\n${invalid.join("\n")}`, { title: "Cht WebSheet" });
    return;
  }
  const references = referenceItems.filter((item) => !item.invalid);
  if (!references.length) {
    showMessageBox("参照範囲を追加してください。", { title: "Cht WebSheet" });
    return;
  }
  const { method, topRow, leftColumn, createLinks } = consolidateDialogOptions(dialog);
  const headers = consolidateHeaders(references, { method, topRow, leftColumn, createLinks });
  const placement = consolidateOutputPlacement(sheet, context.range, references, { method, topRow, leftColumn, createLinks }, headers);
  if (!ensureEditableRangeForProtection(sheet, placement.range, "統合")) return;
  const grouped = consolidateRangesForSheetGroup(context.range, references, {
    method,
    topRow,
    leftColumn,
    createLinks,
    afterOverwriteConfirm: closeDataDialog,
  });
  if (grouped === "pending") return;
  if (grouped === true) {
    closeDataDialog();
    return;
  }
  if (consolidateRanges(sheet, context.range, references, { method, topRow, leftColumn, createLinks, afterOverwriteConfirm: closeDataDialog })) {
    closeDataDialog();
  }
}



function consolidateRanges(sheet, sourceRange, references, options) {
  const sheetIndex = Number.isFinite(Number(options.sheetIndex)) ? Number(options.sheetIndex) : state.activeSheetIndex;
  const sheetId = getSheetId(sheet.name);
  const headers = consolidateHeaders(references, options);
  const placement = consolidateOutputPlacement(sheet, sourceRange, references, options, headers, sheetIndex);
  const outputTop = placement.top;
  const outputLeft = placement.left;
  const outputRows = [];
  if (options.leftColumn) {
    const grouped = new Map();
    references.forEach((range) => {
      const sourceSheet = state.model?.sheets?.[range.sheetIndex ?? state.activeSheetIndex];
      if (!sourceSheet) return;
      const sourceSheetId = getSheetId(sourceSheet.name);
      const dataStart = range.top + (options.topRow ? 1 : 0);
      const dataLeft = range.left + 1;
      for (let row = dataStart; row <= range.bottom; row += 1) {
        const key = String(getDisplayForCell(sourceSheet, sourceSheetId, row, range.left) || "(空白)");
        if (!grouped.has(key)) grouped.set(key, new Map(headers.map((header) => [header, []])));
        const valuesByHeader = grouped.get(key);
        for (let col = dataLeft; col <= range.right; col += 1) {
          const header = consolidateHeaderForColumn(sourceSheet, sourceSheetId, range, col, options, headers, dataLeft);
          if (!valuesByHeader.has(header)) valuesByHeader.set(header, []);
          valuesByHeader.get(header).push(consolidateEntry(sourceSheet, range.sheetIndex ?? state.activeSheetIndex, sourceSheetId, row, col));
        }
      }
    });
    grouped.forEach((values, key) => {
      outputRows.push([key, ...headers.map((header) => consolidateResultValue(values.get(header) || [], options.method, options.createLinks))]);
    });
  } else {
    const valuesByColumn = new Map(headers.map((header) => [header, []]));
    references.forEach((range) => {
      const sourceSheet = state.model?.sheets?.[range.sheetIndex ?? state.activeSheetIndex];
      if (!sourceSheet) return;
      const sourceSheetId = getSheetId(sourceSheet.name);
      const dataStart = range.top + (options.topRow ? 1 : 0);
      const dataLeft = range.left;
      for (let row = dataStart; row <= range.bottom; row += 1) {
        for (let col = dataLeft; col <= range.right; col += 1) {
          const header = consolidateHeaderForColumn(sourceSheet, sourceSheetId, range, col, options, headers, dataLeft);
          if (!valuesByColumn.has(header)) valuesByColumn.set(header, []);
          valuesByColumn.get(header).push(consolidateEntry(sourceSheet, range.sheetIndex ?? state.activeSheetIndex, sourceSheetId, row, col));
        }
      }
    });
    outputRows.push([dataSummaryFunctionLabel(options.method), ...headers.map((header) => consolidateResultValue(valuesByColumn.get(header) || [], options.method, options.createLinks))]);
  }

  const width = Math.max(1, headers.length + 1);
  const height = outputRows.length + 1;
  const targetRange = normalizeRange(
    { sheetIndex, row: outputTop, col: outputLeft },
    { sheetIndex, row: outputTop + height - 1, col: outputLeft + width - 1 },
  );
  const targetValidation = consolidateTargetValidation(sheet, targetRange);
  if (!targetValidation.valid) {
    showMessageBox(targetValidation.message, { title: "Cht WebSheet" });
    return false;
  }
  if (!ensureEditableRangeForProtection(sheet, targetRange, "統合")) return false;
  if (!options.skipOverwriteCheck && rangeHasUserData(sheet, targetRange)) {
    showDataOverwriteConfirmation(() => {
      consolidateRanges(sheet, sourceRange, references, { ...options, skipOverwriteCheck: true, afterOverwriteConfirm: null });
      options.afterOverwriteConfirm?.();
    });
    return false;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  ensureSheetSize(sheet, outputTop + height - 1, outputLeft + width - 1);
  writeCellContent(sheet, sheetId, outputTop, outputLeft, "統合");
  headers.forEach((header, index) => writeCellContent(sheet, sheetId, outputTop, outputLeft + index + 1, header));
  outputRows.forEach((rowValues, rowIndex) => {
    rowValues.forEach((value, colIndex) => writeCellContent(sheet, sheetId, outputTop + rowIndex + 1, outputLeft + colIndex, value));
  });
  forEachCellInRange(targetRange, (row, col) => {
    const cell = ensureCellModel(sheet, row, col);
    if (row === outputTop || col === outputLeft) cell.css.fontWeight = "700";
  });
  internCellsInRange(sheet, targetRange);
  const status = options.createLinks ? "統合元データとリンクする統合を作成しました。" : "統合しました。";
  if (options.commitHistory === false) return true;
  commitDataOperationForSheet(sheetIndex, sheet, sheetBefore, targetRange, `${rangeToLabel(sourceRange)} 統合`, status, targetRange);
  return true;
}



function consolidateRangesForSheetGroup(sourceRange, references, options = {}) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !sourceRange || !references?.length) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    const targetSourceRange = materializeSelectionRangeForSheet(sheet, sourceRange, sheetIndex);
    if (!targetSourceRange) continue;
    const targetReferences = references.map((reference) => {
      if ((reference.sheetIndex ?? state.activeSheetIndex) === state.activeSheetIndex) {
        return { ...reference, sheetIndex, sheetName: state.model?.sheets?.[sheetIndex]?.name };
      }
      return { ...reference };
    });
    const headers = consolidateHeaders(targetReferences, options);
    const placement = consolidateOutputPlacement(sheet, targetSourceRange, targetReferences, options, headers, sheetIndex);
    const targetRange = placement.range;
    const targetValidation = consolidateTargetValidation(sheet, targetRange);
    if (!targetValidation.valid) {
      showMessageBox(`${sheet.name}：${targetValidation.message}`, { title: "Cht WebSheet" });
      return true;
    }
    if (!ensureEditableRangeForProtection(sheet, targetRange, "統合")) return true;
    items.push({ sheetIndex, sheet, sourceRange: targetSourceRange, references: targetReferences, targetRange });
  }
  if (!items.length) return true;
  const overwrite = items.find((item) => rangeHasUserData(item.sheet, item.targetRange));
  if (overwrite && !options.skipOverwriteCheck) {
    showDataOverwriteConfirmation(() => {
      const result = consolidateRangesForSheetGroup(sourceRange, references, { ...options, skipOverwriteCheck: true, afterOverwriteConfirm: null });
      if (result === true) options.afterOverwriteConfirm?.();
    });
    return "pending";
  }
  const affectedSheetIndexes = affectedSheetIndexesForGroupedRanges(items);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  try {
    items.forEach((item) => {
      consolidateRanges(item.sheet, item.sourceRange, item.references, {
        ...options,
        sheetIndex: item.sheetIndex,
        skipOverwriteCheck: true,
        commitHistory: false,
      });
    });
    state.hf = buildFormulaEngine(state.model);
    const activeItem = items.find((item) => item.sheetIndex === state.activeSheetIndex) || items[0];
    state.selected = { sheetIndex: state.activeSheetIndex, row: activeItem.targetRange.top, col: activeItem.targetRange.left };
    state.selectionRange = { ...activeItem.targetRange, sheetIndex: state.activeSheetIndex };
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
    state.selectedImage = null;
    state.selectedImages = [];
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore,
      selectionAfter: state.selectionRange,
      selectionBeforeRanges,
      selectionAfterRanges: [state.selectionRange],
      label: `${rangeToLabel(sourceRange)} ${items.length} シート統合`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートで統合しました。`);
    return true;
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`統合を実行できませんでした：${error.message || error}`);
    return true;
  }
}



function consolidateHeaders(references, options) {
  const headers = [];
  const seen = new Set();
  const pushHeader = (header) => {
    const key = normalizeAdvancedFilterHeader(header);
    if (seen.has(key)) return;
    seen.add(key);
    headers.push(header);
  };
  if (options.topRow) {
    references.forEach((range) => {
      const sourceSheet = state.model?.sheets?.[range.sheetIndex ?? state.activeSheetIndex];
      if (!sourceSheet) return;
      const sourceSheetId = getSheetId(sourceSheet.name);
      const dataLeft = range.left + (options.leftColumn ? 1 : 0);
      for (let col = dataLeft; col <= range.right; col += 1) {
        pushHeader(String(getDisplayForCell(sourceSheet, sourceSheetId, range.top, col) || columnName(col)));
      }
    });
  } else {
    const firstReference = references[0];
    const firstDataLeft = firstReference.left + (options.leftColumn ? 1 : 0);
    const maxWidth = Math.max(
      1,
      ...references.map((range) => Math.max(0, range.right - (range.left + (options.leftColumn ? 1 : 0)) + 1)),
    );
    for (let offset = 0; offset < maxWidth; offset += 1) pushHeader(columnName(firstDataLeft + offset));
  }
  return headers.length ? headers : ["値"];
}



function consolidateHeaderForColumn(sheet, sheetId, range, col, options, headers, dataLeft) {
  if (options.topRow) return String(getDisplayForCell(sheet, sheetId, range.top, col) || columnName(col));
  return headers[col - dataLeft] || columnName(col);
}



function consolidateEntry(sheet, sheetIndex, sheetId, row, col) {
  return {
    value: sheet.cells[cellKey(row, col)]?.raw ?? getDisplayForCell(sheet, sheetId, row, col),
    ref: workbookCellReference(sheetIndex, row, col),
  };
}



function consolidateResultValue(entries, method, createLinks) {
  if (createLinks) {
    const formulaName = dataSummaryFormulaName(method);
    const refs = entries.map((entry) => entry.ref).filter(Boolean);
    if (refs.length) return `=${formulaName}(${refs.join(",")})`;
  }
  return applySummaryFunction(entries.map((entry) => entry.value), method);
}



function dataSummaryFormulaName(value) {
  switch (value) {
    case "count":
      return "COUNTA";
    case "countNums":
      return "COUNT";
    case "average":
      return "AVERAGE";
    case "max":
      return "MAX";
    case "min":
      return "MIN";
    case "product":
      return "PRODUCT";
    case "stdDev":
      return "STDEV.S";
    case "stdDevP":
      return "STDEV.P";
    case "var":
      return "VAR.S";
    case "varP":
      return "VAR.P";
    case "sum":
    default:
      return "SUM";
  }
}



function workbookCellReference(sheetIndex, row, col) {
  const sheet = state.model?.sheets?.[sheetIndex];
  return `${quoteSheetNameForFormula(sheet?.name || activeSheet()?.name || "Sheet1")}!${columnName(col)}${row}`;
}



function workbookRangeReference(sheetIndex, range) {
  const sheet = state.model?.sheets?.[sheetIndex];
  const start = `${columnName(range.left)}${range.top}`;
  const end = `${columnName(range.right)}${range.bottom}`;
  return `${quoteSheetNameForFormula(sheet?.name || activeSheet()?.name || "Sheet1")}!${start === end ? start : `${start}:${end}`}`;
}



function decodeWorkbookRange(value) {
  const text = String(value || "").trim();
  const match = text.match(/^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?(.+)$/);
  if (!match) return null;
  const sheetName = match[1]?.replace(/''/g, "'") || match[2] || "";
  const range = decodeRange(match[3]?.replace(/\$/g, "") || "");
  if (!range) return null;
  if (!sheetName) return { ...range, sheetIndex: state.activeSheetIndex };
  const sheetIndex = state.model?.sheets?.findIndex((sheet) => sheet.name === sheetName);
  if (sheetIndex == null || sheetIndex < 0) return null;
  return { ...range, sheetIndex, sheetName };
}



function decodeWorkbookRangeOrName(value, baseSheet = activeSheet()) {
  const text = String(value || "").trim();
  if (!text) return null;
  return decodeWorkbookRange(text) || workbookRangeFromFormulaReference(text, baseSheet);
}



function decodeWorkbookCell(value) {
  const range = decodeWorkbookRangeOrName(value);
  if (!range || range.top !== range.bottom || range.left !== range.right) return null;
  return {
    sheetIndex: range.sheetIndex ?? state.activeSheetIndex,
    sheetName: range.sheetName,
    row: range.top,
    col: range.left,
  };
}



function openSubtotalDialog() {
  const context = selectedDataRange("小計", { expandCurrentRegion: true });
  if (!context) return;
  const { sheet, range } = context;
  if (rangeIntersectsWorksheetTable(sheet, range)) {
    showMessageBox("テーブルを含む範囲には小計を設定できません。テーブルを範囲に変換してから実行してください。", { title: "Cht WebSheet" });
    return;
  }
  const sheetId = getSheetId(sheet.name);
  const hasHeader = inferHeaderRow(sheet, range);
  const headerRow = hasHeader ? range.top : null;
  const options = [];
  const fields = [];
  for (let col = range.left; col <= range.right; col += 1) {
    const label = headerRow ? String(getDisplayForCell(sheet, sheetId, headerRow, col) || columnName(col)) : columnName(col);
    options.push(`<option value="${col}">${escapeHtml(label)}</option>`);
    fields.push(`<label><input type="checkbox" data-data-subtotal-col="${col}" ${col !== range.left ? "checked" : ""} />${escapeHtml(label)}</label>`);
  }
  state.dataDialog = { type: "subtotal", range };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog subtotal-dialog" role="dialog" aria-modal="true" aria-labelledby="subtotalDialogTitle" data-data-dialog="subtotal">
      <div class="excel-mini-titlebar" id="subtotalDialogTitle">小計</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">グループの基準</span><select data-data-field="groupBy" data-subtotal-sync>${options.join("")}</select></label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">集計の方法</span><select data-data-field="function" data-subtotal-sync>${dataSummaryFunctionOptions("sum")}</select></label>
        <div class="excel-dialog-caption">集計するフィールド</div>
        <div class="remove-duplicates-list subtotal-field-list">${fields.join("")}</div>
        <label class="data-dialog-check"><input type="checkbox" data-data-field="replace" data-subtotal-sync checked />現在の小計をすべて置き換える</label>
        <label class="data-dialog-check"><input type="checkbox" data-data-field="pageBreak" data-subtotal-sync />グループごとに改ページを挿入する</label>
        <label class="data-dialog-check"><input type="checkbox" data-data-field="summaryBelow" data-subtotal-sync checked />集計行をデータの下に挿入する</label>
        <div class="data-dialog-summary subtotal-summary" data-subtotal-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="remove-all">すべて削除</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncSubtotalDialog(overlay.querySelector("[data-data-dialog='subtotal']"));
}



function handleSubtotalDialogSync(event) {
  syncSubtotalDialog(event.currentTarget.closest("[data-data-dialog='subtotal']"));
}



function selectedSubtotalColumnsFromDialog(dialog) {
  return Array.from(dialog?.querySelectorAll("[data-data-subtotal-col]:checked") || [])
    .map((input) => Number(input.dataset.dataSubtotalCol))
    .filter((col) => Number.isFinite(col));
}



function subtotalDialogOptions(dialog) {
  return {
    groupCol: Number(dialog.querySelector("[data-data-field='groupBy']")?.value),
    targetCols: selectedSubtotalColumnsFromDialog(dialog),
    method: dialog.querySelector("[data-data-field='function']")?.value || "sum",
    replace: Boolean(dialog.querySelector("[data-data-field='replace']")?.checked),
    pageBreak: Boolean(dialog.querySelector("[data-data-field='pageBreak']")?.checked),
    summaryBelow: Boolean(dialog.querySelector("[data-data-field='summaryBelow']")?.checked),
  };
}



function syncSubtotalDialog(dialog) {
  if (!dialog) return;
  const summary = dialog.querySelector("[data-subtotal-summary]");
  const preview = subtotalDialogPreview(dialog);
  if (summary) {
    summary.textContent = preview.message;
    summary.classList.toggle("is-warning", preview.warning || preview.blocked);
  }
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = Boolean(preview.blocked);
  const removeAll = dialog.querySelector("[data-data-dialog-action='remove-all']");
  if (removeAll) removeAll.disabled = !countSubtotalRowsInRange(activeSheet(), state.dataDialog?.range || {});
}



function subtotalDialogPreview(dialog) {
  const sheet = activeSheet();
  const range = state.dataDialog?.type === "subtotal" ? state.dataDialog.range : null;
  if (!sheet || !range) return { warning: true, blocked: true, message: "小計に使用する範囲がありません。" };
  if (rangeIntersectsWorksheetTable(sheet, range)) {
    return { warning: true, blocked: true, message: "テーブルを含む範囲には小計を設定できません。" };
  }
  const options = subtotalDialogOptions(dialog);
  if (!options.targetCols.length) return { warning: true, blocked: true, message: "集計するフィールドを1つ以上選択してください。" };
  if (!Number.isFinite(options.groupCol)) return { warning: true, blocked: true, message: "グループの基準を選択してください。" };
  const groups = subtotalGroupsForRange(sheet, range, options.groupCol);
  if (!groups.length) return { warning: true, blocked: true, message: "小計を挿入するデータがありません。" };
  const existing = countSubtotalRowsInRange(sheet, range);
  const inserted = groups.length + 1;
  const replaceText = options.replace && existing ? ` 既存の小計 ${existing} 行を置き換えます。` : "";
  const pageBreakText = options.pageBreak ? " グループごとに改ページを設定します。" : "";
  return {
    warning: false,
    message: `${groups.length} グループに ${dataSummaryFunctionLabel(options.method)} の小計を挿入します。追加予定: ${inserted} 行、対象フィールド: ${options.targetCols.length} 列。${replaceText}${pageBreakText}`,
  };
}



function applySubtotalDialog(dialog, removeAll = false) {
  const range = state.dataDialog?.range;
  if (!range) return;
  const sheet = activeSheet();
  if (!ensureSheetProtectionPermission(sheet, "allowInsertRows", "保護されたシートでは小計を挿入または削除できません。")) return;
  if (removeAll) {
    closeDataDialog();
    if (removeSubtotalRowsForSheetGroup(range)) return;
    removeSubtotalRows(sheet, range);
    return;
  }
  const preview = subtotalDialogPreview(dialog);
  if (preview.blocked) {
    showMessageBox(preview.message, { title: "Cht WebSheet" });
    syncSubtotalDialog(dialog);
    return;
  }
  const options = subtotalDialogOptions(dialog);
  if (!options.targetCols.length) {
    showMessageBox("集計するフィールドを1つ以上選択してください。", { title: "Cht WebSheet" });
    return;
  }
  closeDataDialog();
  if (insertSubtotalRowsForSheetGroup(range, options)) return;
  insertSubtotalRows(sheet, range, options);
}



function removeSubtotalRows(sheet, range, control = {}) {
  const sheetIndex = Number.isFinite(Number(control.sheetIndex)) ? Number(control.sheetIndex) : state.activeSheetIndex;
  const sheetBefore = captureSheetForHistory(sheet);
  clearSubtotalOutlineMetadata(sheet, range);
  clearSubtotalPageBreaks(sheet);
  let removed = 0;
  for (let row = Math.min(sheet.rowCount, range.bottom + 200); row >= range.top; row -= 1) {
    if (sheet.rows[row - 1]?.subtotal) {
      deleteEntireRows(sheet, row, 1);
      removed += 1;
    }
  }
  if (!removed) {
    if (!control.quiet) showMessageBox("削除する小計はありません。", { title: "Cht WebSheet" });
    return null;
  }
  const targetRange = normalizeRange(
    { sheetIndex, row: range.top, col: range.left },
    { sheetIndex, row: Math.max(range.top, range.bottom - removed), col: range.right },
  );
  if (control.commitHistory === false) return { removed, targetRange };
  commitDataOperationForSheet(sheetIndex, sheet, sheetBefore, targetRange, `${rangeToLabel(range)} 小計削除`, "小計を削除しました。", targetRange);
  return { removed, targetRange };
}



function removeSubtotalRowsForSheetGroup(range) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowInsertRows", "保護されたシートでは小計を挿入または削除できません。")) return true;
    const targetRange = materializeSelectionRangeForSheet(sheet, range, sheetIndex);
    if (!targetRange) continue;
    const count = countSubtotalRowsInRange(sheet, targetRange);
    if (count) items.push({ sheetIndex, sheet, range: targetRange, count });
  }
  if (!items.length) {
    showMessageBox("削除する小計はありません。", { title: "Cht WebSheet" });
    return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    const results = items
      .map((item) => removeSubtotalRows(item.sheet, item.range, { sheetIndex: item.sheetIndex, commitHistory: false, quiet: true }))
      .filter(Boolean);
    const removed = results.reduce((total, item) => total + item.removed, 0);
    if (!removed) {
      restoreGroupedSheetsAfterFailure(sheetsBefore);
      showMessageBox("削除する小計はありません。", { title: "Cht WebSheet" });
      return true;
    }
    state.hf = buildFormulaEngine(state.model);
    const activeTargetRange = results.find((item) => item.targetRange?.sheetIndex === state.activeSheetIndex)?.targetRange || results[0].targetRange || range;
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
      label: `${items.length} シート 小計削除`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートで小計を削除しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`小計を削除できませんでした：${error.message || error}`);
  }
  return true;
}



function countSubtotalRowsInRange(sheet, range) {
  let count = 0;
  for (let row = range.top; row <= Math.min(sheet.rowCount, range.bottom + 200); row += 1) {
    if (sheet.rows[row - 1]?.subtotal) count += 1;
  }
  return count;
}



function subtotalGroupsForRange(sheet, range, groupCol) {
  const hasHeader = inferHeaderRow(sheet, range);
  const firstDataRow = hasHeader ? range.top + 1 : range.top;
  if (!sheet || firstDataRow > range.bottom || !Number.isFinite(Number(groupCol))) return [];
  const sheetId = getSheetId(sheet.name);
  const groups = [];
  let groupStart = firstDataRow;
  let previous = getDisplayForCell(sheet, sheetId, firstDataRow, groupCol);
  for (let row = firstDataRow + 1; row <= range.bottom; row += 1) {
    const current = getDisplayForCell(sheet, sheetId, row, groupCol);
    if (String(current) !== String(previous)) {
      groups.push({ label: previous, start: groupStart, end: row - 1 });
      groupStart = row;
      previous = current;
    }
  }
  groups.push({ label: previous, start: groupStart, end: range.bottom });
  return groups;
}



function insertSubtotalRows(sheet, range, options, control = {}) {
  const sheetIndex = Number.isFinite(Number(control.sheetIndex)) ? Number(control.sheetIndex) : state.activeSheetIndex;
  const hasHeader = inferHeaderRow(sheet, range);
  const firstDataRow = hasHeader ? range.top + 1 : range.top;
  if (firstDataRow > range.bottom) {
    if (!control.quiet) showMessageBox("小計を挿入するデータがありません。", { title: "Cht WebSheet" });
    return null;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  const removedSubtotalRows = options.replace ? removeSubtotalRowsWithoutHistory(sheet, range) : 0;
  const workingRange = removedSubtotalRows
    ? { ...range, bottom: Math.max(range.top, range.bottom - removedSubtotalRows) }
    : range;
  const sheetId = getSheetId(sheet.name);
  const groups = subtotalGroupsForRange(sheet, workingRange, options.groupCol);
  if (!groups.length) {
    restoreSheetFromHistory(sheet, sheetBefore);
    if (!control.quiet) showMessageBox("小計を挿入するデータがありません。", { title: "Cht WebSheet" });
    return null;
  }

  const formulaCode = dataSummarySubtotalCode(options.method);
  let inserted = 0;
  const insertedGroups = [];
  sheet.outlineSettings = normalizeOutlineSettings({
    ...(sheet.outlineSettings || {}),
    summaryBelow: options.summaryBelow,
  });
  groups.forEach((group) => {
    const dataStart = options.summaryBelow ? group.start + inserted : group.start + inserted + 1;
    const dataEnd = options.summaryBelow ? group.end + inserted : group.end + inserted + 1;
    const insertAt = options.summaryBelow ? dataEnd + 1 : group.start + inserted;
    insertEntireRows(sheet, insertAt, 1);
    inserted += 1;
    const rowModel = sheet.rows[insertAt - 1] || defaultRowModel(sheet);
    rowModel.subtotal = true;
    rowModel.subtotalGroup = true;
    rowModel.outlineLevel = 0;
    sheet.rows[insertAt - 1] = rowModel;
    const label = `${group.label || "(空白)"} ${dataSummaryFunctionLabel(options.method)}`;
    writeCellContent(sheet, sheetId, insertAt, options.groupCol, label);
    for (let col = workingRange.left; col <= workingRange.right; col += 1) {
      const cell = ensureCellModel(sheet, insertAt, col);
      cell.css.fontWeight = "700";
      cell.borders.top = { width: 1, style: "solid", color: "#000000" };
    }
    insertedGroups.push({ dataStart, dataEnd, subtotalRow: insertAt });
  });
  const grandTotalRow = insertSubtotalGrandTotalRow(sheet, workingRange, {
    ...options,
    firstDataRow,
    insertedSubtotalRows: inserted,
    formulaCode,
    insertedGroups,
  });
  if (grandTotalRow) {
    inserted += 1;
    if (!options.summaryBelow) {
      insertedGroups.forEach((group) => {
        group.dataStart += 1;
        group.dataEnd += 1;
        group.subtotalRow += 1;
      });
    }
  }
  insertedGroups.forEach((group) => {
    applySubtotalGroupOutline(sheet, group, options);
    options.targetCols.forEach((col) => {
      writeCellContent(sheet, sheetId, group.subtotalRow, col, `=SUBTOTAL(${formulaCode},${columnName(col)}${group.dataStart}:${columnName(col)}${group.dataEnd})`);
    });
  });
  if (grandTotalRow) {
    options.targetCols.forEach((col) => {
      const refs = insertedGroups.map((group) => `${columnName(col)}${group.dataStart}:${columnName(col)}${group.dataEnd}`);
      writeCellContent(sheet, sheetId, grandTotalRow, col, `=SUBTOTAL(${formulaCode},${refs.join(",")})`);
    });
  }
  const targetRange = normalizeRange(
    { sheetIndex, row: workingRange.top, col: workingRange.left },
    { sheetIndex, row: workingRange.bottom + inserted, col: workingRange.right },
  );
  internCellsInRange(sheet, targetRange);
  if (control.commitHistory === false) return { inserted, targetRange, removedSubtotalRows };
  commitDataOperationForSheet(sheetIndex, sheet, sheetBefore, targetRange, `${rangeToLabel(range)} 小計`, "小計を挿入しました。", targetRange);
  return { inserted, targetRange, removedSubtotalRows };
}

export {
  applyCollectedRemoveDuplicateRows,
  applyConsolidateDialog,
  applyDataValidationDialog,
  applyDataValidationDialogToSheetGroup,
  applyDataValidationToSheetPoints,
  applyRelationshipAutoDetectDialog,
  applyRelationshipDefinitions,
  applyRelationshipEditDialog,
  applyRemoveDuplicatesDialog,
  applySubtotalDialog,
  applySummaryFunction,
  autoDetectRelationshipDefinitions,
  autoDetectRelationshipsFromDialog,
  bestRelationshipColumnMatch,
  boundingRangeForPoints,
  canCreateWorkbookRelationship,
  canonicalRelationshipColumnName,
  clearDataValidationFromSelection,
  clearDataValidationFromSheetGroup,
  clearInvalidDataValidationMarks,
  cloneRelationships,
  closeDataValidationMenu,
  closeDataValidationMenuOnEscape,
  closeDataValidationMenuOnOutsidePointer,
  closeValidationDropdownOnEscape,
  closeValidationDropdownOnOutsidePointer,
  closeValidationDropdownPopup,
  collectRemoveDuplicateRows,
  consolidateDialogOptions,
  consolidateDialogPreview,
  consolidateDialogReferences,
  consolidateEntry,
  consolidateHeaderForColumn,
  consolidateHeaders,
  consolidateLeftColumnLabels,
  consolidateOutputPlacement,
  consolidateRanges,
  consolidateRangesForSheetGroup,
  consolidateResultValue,
  consolidateTargetValidation,
  countSubtotalRowsInRange,
  dataOperationRangeFromCurrentRegion,
  dataSummaryFormulaName,
  dataSummaryFunctionLabel,
  dataSummaryFunctionOptions,
  dataSummarySubtotalCode,
  dataValidationDialogPreview,
  dataValidationDialogValidation,
  dataValidationImeModeOptions,
  dataValidationOperatorLabel,
  dataValidationSelectionStatsMessage,
  dataValidationTypeLabel,
  decodeWorkbookCell,
  decodeWorkbookRange,
  decodeWorkbookRangeOrName,
  deleteRelationshipFromDialog,
  evaluateValidationFormulaValue,
  handleConsolidateDialogSync,
  handleDataValidationDialogSync,
  handleDataValidationMenuAction,
  handleRelationshipAutoDetectSync,
  handleRelationshipEditDialogSync,
  handleRelationshipsDialogSync,
  handleRemoveDuplicatesDialogSync,
  handleSubtotalDialogSync,
  handleValidationDropdownButtonClick,
  handleValidationDropdownButtonPointerDown,
  handleValidationDropdownChoice,
  handleValidationDropdownKeydown,
  hasValidationInvalidMark,
  inferRelationshipColumnType,
  insertSubtotalRows,
  markInvalidDataValidationCells,
  normalizeRelationshipColumnName,
  normalizeValidationListValue,
  normalizeValidationTimeText,
  openConsolidateDialog,
  openDataValidationDialog,
  openDataValidationMenu,
  openRelationshipAutoDetectDialog,
  openRelationshipEditDialog,
  openRelationshipsDialog,
  openRemoveDuplicatesDialog,
  openSelectedValidationDropdown,
  openSubtotalDialog,
  openValidationDropdownPopup,
  parseValidationClockText,
  positionValidationDropdownPopup,
  pushRelationshipHistory,
  rangeHasDataValidation,
  relationshipCardinalityLabel,
  relationshipCardinalityNote,
  relationshipColumnMatchScore,
  relationshipColumnNameTokens,
  relationshipColumnOptionsHtml,
  relationshipColumnStats,
  relationshipColumnTypeLabel,
  relationshipDefinitionExists,
  relationshipDefinitionFromDialog,
  relationshipDefinitionLabel,
  relationshipDetailsHtml,
  relationshipPreviewMessage,
  relationshipSummaryLabel,
  relationshipTableKey,
  relationshipTableLabel,
  relationshipTableOptionsHtml,
  relationshipUnavailableMessage,
  removeDuplicateRows,
  removeDuplicateRowsForSheetGroup,
  removeDuplicatesCloneRowModel,
  removeDuplicatesComparableValue,
  removeDuplicatesFromSelection,
  removeDuplicatesMergeConflict,
  removeDuplicatesNormalizedScalar,
  removeDuplicatesPreview,
  removeSubtotalRows,
  removeSubtotalRowsForSheetGroup,
  resetDataValidationDialog,
  sameSheetRangeIntersects,
  selectedRelationshipFromDialog,
  selectedRemoveDuplicateColumnsFromDialog,
  selectedSubtotalColumnsFromDialog,
  shouldShowDataExpansionWarning,
  subtotalDialogOptions,
  subtotalDialogPreview,
  subtotalGroupsForRange,
  syncConsolidateDialog,
  syncDataValidationDialog,
  syncRelationshipAutoDetectDialog,
  syncRelationshipEditDialog,
  syncRemoveDuplicatesDialog,
  syncSubtotalDialog,
  tableColumnNames,
  tableColumnValues,
  tableDataRowCount,
  tableDistinctValueCount,
  tableRefByKey,
  translateValidationFormulaReferences,
  validateCellInput,
  validateCustomDataValidation,
  validateRelationshipDefinition,
  validationComparableBoundValue,
  validationComparableValue,
  validationDateValue,
  validationFormulaForCell,
  validationFormulaResultTruthy,
  validationFromDialogFields,
  validationFromDialogFieldsForPreview,
  validationMarkStore,
  validationScalarValue,
  validationSourceCaption,
  validationTargetPointsForDialog,
  validationTimeValue,
  workbookCellReference,
  workbookRangeReference,
  workbookRelationships,
  workbookTableRefs,
};
