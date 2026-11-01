// Feature module; shared bindings are imported explicitly.
import {
  DATA_ANALYSIS_HELP,
  DATA_ANALYSIS_METHODS,
  DATA_ANALYSIS_TWO_VARIABLE_METHODS,
  createIcons,
  icons,
  state,
} from "../../app-core.js";
import {
  dataCellFieldHtml,
  dataRangeFieldHtml,
} from "./filter-and-text-tools.js";
import {
  dataAnalysisAnovaRows,
  dataAnalysisAnovaTwoFactorNoReplicationRows,
  dataAnalysisAnovaTwoFactorReplicationRows,
  dataAnalysisCellNumericValue,
  dataAnalysisCorrelation,
  dataAnalysisCovarianceP,
  dataAnalysisDescriptiveRows,
  dataAnalysisExponentialSmoothingRows,
  dataAnalysisFTestRows,
  dataAnalysisFourierInputValues,
  dataAnalysisFourierRows,
  dataAnalysisHistogramRows,
  dataAnalysisInputMatrix,
  dataAnalysisMatrixRows,
  dataAnalysisMovingAverageRows,
  dataAnalysisRandomRows,
  dataAnalysisRankPercentileRows,
  dataAnalysisRegressionRows,
  dataAnalysisSamplingRows,
  dataAnalysisSeries,
  dataAnalysisTTestRows,
  dataAnalysisZTestRows,
  scenarioPointKey,
  scenarioPointLabel,
  scenarioSummaryAddress,
  scenarioSummaryValueMaps,
  scenarioValueForPoint,
  setScenarioValueForPoint,
} from "./forecast-and-statistics.js";
import {
  rangeHasUserData,
  showDataOverwriteConfirmation,
} from "./sort-and-autofilter.js";
import {
  decodeWorkbookCell,
  decodeWorkbookRangeOrName,
  insertSubtotalRows,
  workbookCellReference,
  workbookRangeReference,
} from "./validation-and-consolidation.js";
import {
  chartSvg,
  closeDataDialog,
  uniqueSheetNameWithBase,
} from "../drawing/charts-and-insertion.js";
import {
  clampNumber,
  endingCellForPixelSize,
  fitObjectSize,
} from "../drawing/shape-rendering.js";
import {
  mergesIntersectingRange,
} from "../editing/autofill.js";
import {
  closeRibbonBorderMenu,
  closeRibbonShapeMenu,
  closeRibbonUnderlineMenu,
  closeShapeFormatMenu,
  svgDataUrl,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  affectedSheetIndexesForGroupedRanges,
  coerceUserInput,
  getCellRawInput,
  getDisplayForCell,
  materializeSelectionRangeForSheet,
  normalizeRange,
  rangeToLabel,
  restoreGroupedSheetsAfterFailure,
  updateCellModel,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
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
  pushHistory,
  pushSheetInsertHistory,
  restoreSheetsFromHistory,
  storeCellModelForSheet,
} from "../editing/structure-and-history.js";
import {
  ensureCellModel,
  forEachCellInRange,
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
  decodeCell,
  escapeAttr,
  escapeHtml,
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
  commitDataOperationForSheet,
  ensureSheetSize,
  selectedDataRange,
} from "../ui/ribbon-and-links.js";
import {
  closeSheetTabContextMenu,
  createBlankSheet,
  ensureEditableRangeForProtection,
  ensureSheetProtectionPermission,
  isSheetProtected,
  rangeHasLockedCellsForProtection,
  sheetProtectionAllows,
  showProtectedSheetMessage,
  validateSheetName,
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
  ensurePageLayoutSettings,
  normalizedPageLayoutSettings,
} from "../workbook/vba-and-printing.js";
import {
  commandMenuAccessKeyAttr,
} from "../workbook/views-and-macros.js";


function insertSubtotalRowsForSheetGroup(range, options) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowInsertRows", "保護されたシートでは小計を挿入または削除できません。")) return true;
    const targetRange = materializeSelectionRangeForSheet(sheet, range, sheetIndex);
    if (!targetRange) continue;
    items.push({ sheetIndex, sheet, range: targetRange });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    const results = [];
    for (const item of items) {
      const result = insertSubtotalRows(item.sheet, item.range, options, {
        sheetIndex: item.sheetIndex,
        commitHistory: false,
        quiet: true,
      });
      if (!result) {
        restoreGroupedSheetsAfterFailure(sheetsBefore);
        showMessageBox(`${item.sheet.name}：小計を挿入するデータがありません。`, { title: "Cht WebSheet" });
        return true;
      }
      results.push(result);
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
      label: `${items.length} シート 小計`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートで小計を挿入しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`小計を挿入できませんでした：${error.message || error}`);
  }
  return true;
}



function insertSubtotalGrandTotalRow(sheet, range, options) {
  const sheetId = getSheetId(sheet.name);
  const row = options.summaryBelow
    ? range.bottom + Number(options.insertedSubtotalRows || 0) + 1
    : options.firstDataRow;
  insertEntireRows(sheet, row, 1);
  const rowModel = sheet.rows[row - 1] || defaultRowModel(sheet);
  rowModel.subtotal = true;
  rowModel.subtotalGrand = true;
  rowModel.outlineLevel = Math.max(0, Number(rowModel.outlineLevel || 0));
  sheet.rows[row - 1] = rowModel;
  writeCellContent(sheet, sheetId, row, options.groupCol, "総計");
  for (let col = range.left; col <= range.right; col += 1) {
    const cell = ensureCellModel(sheet, row, col);
    cell.css.fontWeight = "700";
    cell.borders.top = { width: 2, style: "solid", color: "#000000" };
  }
  return row;
}



function applySubtotalGroupOutline(sheet, group, options = {}) {
  const dataStart = Math.min(group.dataStart, group.dataEnd);
  const dataEnd = Math.max(group.dataStart, group.dataEnd);
  for (let row = dataStart; row <= dataEnd; row += 1) {
    const rowModel = sheet.rows[row - 1] || defaultRowModel(sheet);
    if (!rowModel.subtotal) {
      rowModel.subtotalDetail = true;
      rowModel.outlineLevel = Math.max(1, Number(rowModel.outlineLevel || 0));
      sheet.rows[row - 1] = rowModel;
    }
  }
  const subtotalModel = sheet.rows[group.subtotalRow - 1] || defaultRowModel(sheet);
  subtotalModel.subtotal = true;
  subtotalModel.subtotalGroup = true;
  subtotalModel.outlineLevel = 0;
  sheet.rows[group.subtotalRow - 1] = subtotalModel;

  if (options.pageBreak) {
    const breakBeforeRow = options.summaryBelow ? group.subtotalRow + 1 : dataEnd + 1;
    addSubtotalPageBreak(sheet, breakBeforeRow);
  }
}



function addSubtotalPageBreak(sheet, row) {
  const rowNumber = Math.round(Number(row));
  if (!sheet || !Number.isFinite(rowNumber) || rowNumber <= 1) return;
  const layout = ensurePageLayoutSettings(sheet);
  layout.pageBreaks = layout.pageBreaks || { rows: [], cols: [] };
  layout.pageBreaks.rows = [...new Set([...(layout.pageBreaks.rows || []), rowNumber])].sort((a, b) => a - b);
  sheet.pageLayout = normalizedPageLayoutSettings(layout, sheet);
  sheet.subtotalPageBreaks = [...new Set([...(sheet.subtotalPageBreaks || []), rowNumber])].sort((a, b) => a - b);
}



function clearSubtotalOutlineMetadata(sheet, range) {
  if (!sheet || !range) return 0;
  let changed = 0;
  const bottom = Math.min(sheet.rowCount, Math.max(range.bottom, range.top) + 200);
  for (let row = Math.max(1, range.top); row <= bottom; row += 1) {
    const rowModel = sheet.rows[row - 1];
    if (!rowModel?.subtotalDetail) continue;
    delete rowModel.subtotalDetail;
    rowModel.outlineLevel = 0;
    if (rowModel.outlineHidden) {
      rowModel.outlineHidden = false;
      rowModel.hidden = false;
    }
    changed += 1;
  }
  return changed;
}



function clearSubtotalPageBreaks(sheet) {
  const breaks = (sheet?.subtotalPageBreaks || []).map(Number).filter((row) => Number.isFinite(row));
  if (!sheet || !breaks.length) return 0;
  const breakSet = new Set(breaks);
  const layout = ensurePageLayoutSettings(sheet);
  layout.pageBreaks = layout.pageBreaks || { rows: [], cols: [] };
  const before = layout.pageBreaks.rows?.length || 0;
  layout.pageBreaks.rows = (layout.pageBreaks.rows || []).filter((row) => !breakSet.has(row));
  sheet.pageLayout = normalizedPageLayoutSettings(layout, sheet);
  sheet.subtotalPageBreaks = [];
  return before - layout.pageBreaks.rows.length;
}



function removeSubtotalRowsWithoutHistory(sheet, range) {
  clearSubtotalOutlineMetadata(sheet, range);
  clearSubtotalPageBreaks(sheet);
  let removed = 0;
  for (let row = Math.min(sheet.rowCount, range.bottom + 200); row >= range.top; row -= 1) {
    if (sheet.rows[row - 1]?.subtotal) {
      deleteEntireRows(sheet, row, 1);
      removed += 1;
    }
  }
  return removed;
}



function insertSummaryRow(range, label, statusMessage) {
  const sheet = activeSheet();
  const sheetBefore = captureSheetForHistory(sheet);
  const targetRow = range.bottom + 1;
  ensureSheetSize(sheet, targetRow, range.right);
  const sheetId = getSheetId(sheet.name);
  writeCellContent(sheet, sheetId, targetRow, range.left, label);
  for (let col = range.left + 1; col <= range.right; col += 1) {
    const values = [];
    for (let row = range.top + 1; row <= range.bottom; row += 1) {
      const value = Number(sheet.cells[cellKey(row, col)]?.raw ?? getDisplayForCell(sheet, sheetId, row, col));
      if (Number.isFinite(value)) values.push(value);
    }
    if (values.length) writeCellContent(sheet, sheetId, targetRow, col, `=SUBTOTAL(9,${columnName(col)}${range.top + 1}:${columnName(col)}${range.bottom})`);
  }
  const targetRange = { ...range, bottom: targetRow };
  commitDataOperation(sheet, sheetBefore, targetRange, `${rangeToLabel(range)} ${label}`, statusMessage, normalizeRange({ sheetIndex: state.activeSheetIndex, row: targetRow, col: range.left }, { sheetIndex: state.activeSheetIndex, row: targetRow, col: range.right }));
}



function openWhatIfMenu(trigger = null) {
  if (!trigger) {
    openWhatIfDialog();
    return;
  }
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonBorderMenu();
  closeRibbonUnderlineMenu();
  closeRibbonShapeMenu();
  closeShapeFormatMenu();
  closeWhatIfMenu();
  const menu = document.createElement("div");
  menu.id = "whatIfMenu";
  menu.className = "ribbon-data-menu what-if-menu";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", "What-If 分析");
  menu.innerHTML = `
    <button type="button" role="menuitem" data-what-if-menu-action="scenario-manager"${commandMenuAccessKeyAttr("シナリオの登録と管理(S)...")}><i data-lucide="list-filter"></i><span>${accessKeyLabelHtml("シナリオの登録と管理(S)...")}</span></button>
    <button type="button" role="menuitem" data-what-if-menu-action="goal-seek"${commandMenuAccessKeyAttr("ゴール シーク(G)...")}><i data-lucide="git-branch"></i><span>${accessKeyLabelHtml("ゴール シーク(G)...")}</span></button>
    <button type="button" role="menuitem" data-what-if-menu-action="data-table"${commandMenuAccessKeyAttr("データ テーブル(D)...")}><i data-lucide="table-2"></i><span>${accessKeyLabelHtml("データ テーブル(D)...")}</span></button>`;
  enhanceCommandMenuAccessKeys(menu);
  document.body.append(menu);
  if (window.lucide) window.lucide.createIcons({ icons });
  const rect = trigger.getBoundingClientRect();
  positionPopupMenu(menu, rect.left, rect.bottom + 2);
  state.whatIfMenu = true;
  menu.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
}



function closeWhatIfMenu() {
  document.getElementById("whatIfMenu")?.remove();
  state.whatIfMenu = null;
}



function closeWhatIfMenuOnOutsidePointer(event) {
  if (!state.whatIfMenu) return;
  if (event.target.closest?.("#whatIfMenu") || event.target.closest?.("[data-ribbon-command='data-what-if']")) return;
  closeWhatIfMenu();
}



function closeWhatIfMenuOnEscape(event) {
  if (event.key !== "Escape" || !document.querySelector("#whatIfMenu")) return;
  if (commandMenuEscapeBlockedByDialog()) return;
  consumeCommandMenuKey(event);
  closeWhatIfMenu();
}



function handleWhatIfMenuAction(event) {
  if (event.currentTarget.disabled || event.currentTarget.classList.contains("is-disabled")) return;
  const action = event.currentTarget.dataset.whatIfMenuAction;
  closeWhatIfMenu();
  if (action === "scenario-manager") {
    openScenarioManagerDialog();
    return;
  }
  if (action === "data-table") {
    openDataTableDialog();
    return;
  }
  openWhatIfDialog();
}



function openWhatIfDialog() {
  const point = state.selected || { row: 1, col: 1 };
  const selectedReference = workbookCellReference(point.sheetIndex ?? state.activeSheetIndex, point.row, point.col);
  const setCellReference = selectedReference || "";
  state.dataDialog = { type: "what-if" };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog what-if-dialog" role="dialog" aria-modal="true" aria-labelledby="whatIfDialogTitle" data-data-dialog="what-if">
      <div class="excel-mini-titlebar" id="whatIfDialogTitle">ゴール シーク</div>
      <div class="excel-mini-body data-dialog-body">
        ${dataCellFieldHtml("数式入力セル", "setCell", setCellReference, "goal-seek")}
        <label class="excel-dialog-field"><span class="excel-dialog-caption">目標値</span><input type="number" data-data-field="targetValue" value="0" data-goal-seek-sync /></label>
        ${dataCellFieldHtml("変化させるセル", "changingCell", "", "goal-seek")}
        <div class="data-dialog-summary goal-seek-summary" data-goal-seek-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  const dialog = overlay.querySelector("[data-data-dialog='what-if']");
  syncGoalSeekDialog(dialog);
  overlay.querySelector(setCellReference ? "[data-data-field='changingCell']" : "[data-data-field='setCell']")?.focus({ preventScroll: true });
}



function applyWhatIfDialog(dialog) {
  const validation = goalSeekDialogValidation(dialog);
  if (!validation.valid) {
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    syncGoalSeekDialog(dialog);
    return;
  }
  const grouped = runGoalSeekForSheetGroup(validation);
  if (grouped) {
    closeDataDialog();
    showGoalSeekResultDialog(grouped);
    return;
  }
  const result = runGoalSeek(validation.setCell, validation.targetValue, validation.changingCell);
  if (result) {
    closeDataDialog();
    showGoalSeekResultDialog(result);
  }
}



function handleGoalSeekDialogSync(event) {
  syncGoalSeekDialog(event.currentTarget.closest("[data-data-dialog='what-if']"));
}



function syncGoalSeekDialog(dialog) {
  if (!dialog) return;
  const validation = goalSeekDialogValidation(dialog, { preview: true });
  const summary = dialog.querySelector("[data-goal-seek-summary]");
  if (summary) {
    summary.textContent = validation.message;
    summary.classList.toggle("is-warning", !validation.valid || Boolean(validation.warning));
  }
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = false;
}



function goalSeekDialogValidation(dialog, options = {}) {
  const setCellText = dialog?.querySelector("[data-data-field='setCell']")?.value || "";
  const changingCellText = dialog?.querySelector("[data-data-field='changingCell']")?.value || "";
  const targetText = dialog?.querySelector("[data-data-field='targetValue']")?.value || "";
  const setCell = decodeWorkbookCell(setCellText);
  const changingCell = decodeWorkbookCell(changingCellText);
  const targetValue = Number(targetText);
  if (!setCell) return { valid: false, message: "数式入力セルを指定してください。" };
  if (!Number.isFinite(targetValue)) return { valid: false, message: "目標値を数値で入力してください。" };
  if (!changingCell) return { valid: false, message: "変化させるセルを指定してください。" };
  const setSheetIndex = setCell.sheetIndex ?? state.activeSheetIndex;
  const changingSheetIndex = changingCell.sheetIndex ?? state.activeSheetIndex;
  const setSheet = state.model?.sheets?.[setSheetIndex];
  const changingSheet = state.model?.sheets?.[changingSheetIndex];
  const setSheetId = getSheetId(setSheet?.name);
  const changingSheetId = getSheetId(changingSheet?.name);
  if (!setSheet || setSheetId == null || !changingSheet || changingSheetId == null) {
    return { valid: false, message: "セル参照を正しく入力してください。" };
  }
  if (setSheetIndex === changingSheetIndex && setCell.row === changingCell.row && setCell.col === changingCell.col) {
    return { valid: false, message: "数式入力セルと変化させるセルには別のセルを指定してください。" };
  }
  const formulaCell = setSheet.cells[cellKey(setCell.row, setCell.col)];
  if (!formulaCell?.formula) {
    return { valid: false, message: "数式入力セルには数式を指定してください。" };
  }
  const changingCellModel = changingSheet.cells[cellKey(changingCell.row, changingCell.col)];
  if (changingCellModel?.formula) {
    return { valid: false, message: "変化させるセルには数式を含まないセルを指定してください。" };
  }
  const changingRange = normalizeRange(
    { sheetIndex: changingSheetIndex, row: changingCell.row, col: changingCell.col },
    { sheetIndex: changingSheetIndex, row: changingCell.row, col: changingCell.col },
  );
  if (rangeHasLockedCellsForProtection(changingSheet, changingRange)) {
    return { valid: false, message: "保護されたシート上のロックされたセルは変化させるセルに指定できません。" };
  }
  const currentValue = Number(getDisplayForCell(setSheet, setSheetId, setCell.row, setCell.col));
  const changingValueRaw = changingCellModel?.raw ?? getDisplayForCell(changingSheet, changingSheetId, changingCell.row, changingCell.col);
  const changingValue = Number(changingValueRaw);
  const warning = !Number.isFinite(currentValue) || (!Number.isFinite(changingValue) && String(changingValueRaw ?? "").trim() !== "");
  const message = warning && options.preview
    ? "数値として計算できない値があります。実行時に Excel と同じく近似計算を試行します。"
    : `${workbookCellReference(setSheetIndex, setCell.row, setCell.col)} の現在値 ${formatGoalSeekNumber(currentValue)} を ${formatGoalSeekNumber(targetValue)} に近づけるため、${workbookCellReference(changingSheetIndex, changingCell.row, changingCell.col)} を変更します。`;
  return { valid: true, warning, message, setCell, changingCell, targetValue };
}



function runGoalSeek(setCell, targetValue, changingCell, control = {}) {
  const setSheetIndex = setCell.sheetIndex ?? state.activeSheetIndex;
  const changingSheetIndex = changingCell.sheetIndex ?? state.activeSheetIndex;
  const setSheet = state.model?.sheets?.[setSheetIndex];
  const changingSheet = state.model?.sheets?.[changingSheetIndex];
  const setSheetId = getSheetId(setSheet?.name);
  const changingSheetId = getSheetId(changingSheet?.name);
  if (!setSheet || !changingSheet || setSheetId == null || changingSheetId == null) {
    showMessageBox("セル参照を正しく入力してください。", { title: "Cht WebSheet" });
    return null;
  }
  if (setSheetIndex === changingSheetIndex && setCell.row === changingCell.row && setCell.col === changingCell.col) {
    showMessageBox("数式入力セルと変化させるセルには別のセルを指定してください。", { title: "Cht WebSheet" });
    return null;
  }
  const formulaCell = setSheet.cells[cellKey(setCell.row, setCell.col)];
  if (!formulaCell?.formula) {
    showMessageBox("数式入力セルには数式を指定してください。", { title: "Cht WebSheet" });
    return null;
  }
  const changingCellModel = changingSheet.cells[cellKey(changingCell.row, changingCell.col)];
  if (changingCellModel?.formula) {
    showMessageBox("変化させるセルには数式を含まないセルを指定してください。", { title: "Cht WebSheet" });
    return null;
  }
  const changingRange = normalizeRange(
    { sheetIndex: changingSheetIndex, row: changingCell.row, col: changingCell.col },
    { sheetIndex: changingSheetIndex, row: changingCell.row, col: changingCell.col },
  );
  if (control.commitHistory !== false && !ensureEditableRangeForProtection(changingSheet, changingRange, "ゴール シーク")) return null;
  const affectedSheetIndexes = [...new Set([setSheetIndex, changingSheetIndex])];
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  const activeSheetBefore = state.activeSheetIndex;
  const selectionBefore = state.selectionRange ? { ...state.selectionRange } : null;
  const selectionBeforeRanges = activeSelectionRanges().map((rangeItem) => ({ ...rangeItem }));
  const initial = Number(changingSheet.cells[cellKey(changingCell.row, changingCell.col)]?.raw ?? 0) || 0;
  const tolerance = Math.max(0.000001, Math.abs(targetValue) * 0.0000001);
  const maxIterations = 100;
  let x0 = initial;
  let y0 = goalSeekValue(setSheetId, setCell, changingSheet, changingSheetId, changingCell, x0) - targetValue;
  if (!Number.isFinite(y0)) {
    restoreSheetsFromHistory(sheetsBefore);
    state.hf = buildFormulaEngine(state.model);
    renderSheet();
    showMessageBox("目標値に使用できる数値結果がありません。", { title: "Cht WebSheet" });
    return null;
  }
  let x1 = goalSeekInitialSecondGuess(initial);
  let success = false;
  let iterations = 0;
  for (let iteration = 0; iteration < maxIterations; iteration += 1) {
    iterations = iteration + 1;
    const y1 = goalSeekValue(setSheetId, setCell, changingSheet, changingSheetId, changingCell, x1) - targetValue;
    if (!Number.isFinite(y1)) break;
    if (Math.abs(y1) <= tolerance) {
      success = true;
      y0 = y1;
      break;
    }
    const denominator = y1 - y0;
    const next = Math.abs(denominator) < 0.0000001 ? goalSeekFallbackGuess(x1, iteration) : x1 - y1 * (x1 - x0) / denominator;
    x0 = x1;
    y0 = y1;
    x1 = Number.isFinite(next) ? next : goalSeekFallbackGuess(x1, iteration);
  }
  writeCellContent(changingSheet, changingSheetId, changingCell.row, changingCell.col, x1);
  state.hf = buildFormulaEngine(state.model);
  const finalValue = Number(state.hf.getCellValue({ sheet: setSheetId, row: setCell.row - 1, col: setCell.col - 1 }));
  success = success || (Number.isFinite(finalValue) && Math.abs(finalValue - targetValue) <= tolerance);
  const range = normalizeRange({ sheetIndex: changingSheetIndex, row: changingCell.row, col: changingCell.col }, { sheetIndex: changingSheetIndex, row: changingCell.row, col: changingCell.col });
  const result = {
    success,
    targetValue,
    finalValue,
    changingValue: x1,
    iterations,
    setCell: workbookCellReference(setSheetIndex, setCell.row, setCell.col),
    changingCell: workbookCellReference(changingSheetIndex, changingCell.row, changingCell.col),
    affectedSheetIndexes,
    sheetsBefore,
    activeSheetBefore,
    selectionBefore,
    selectionBeforeRanges,
    selectionAfter: range,
  };
  if (control.commitHistory === false) return result;
  state.activeSheetIndex = changingSheetIndex;
  state.selected = { sheetIndex: changingSheetIndex, row: changingCell.row, col: changingCell.col };
  state.selectionRange = { ...range };
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  pushHistory({
    sheetIndex: changingSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(affectedSheetIndexes),
    selectionBefore,
    selectionAfter: range,
    selectionBeforeRanges,
    selectionAfterRanges: [range],
    label: "ゴール シーク",
  });
  renderWorkbook();
  updateFormulaBarForSelection();
  setStatus(success ? "ゴール シークが完了しました。" : "近似値を入力しました。");
  return result;
}



function runGoalSeekForSheetGroup(validation) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !validation?.setCell || !validation?.changingCell) return false;
  const activeIndex = state.activeSheetIndex;
  if ((validation.setCell.sheetIndex ?? activeIndex) !== activeIndex) return false;
  if ((validation.changingCell.sheetIndex ?? activeIndex) !== activeIndex) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    const sheetId = getSheetId(sheet?.name);
    if (!sheet || sheetId == null) continue;
    const setCell = { ...validation.setCell, sheetIndex };
    const changingCell = { ...validation.changingCell, sheetIndex };
    const formulaCell = sheet.cells[cellKey(setCell.row, setCell.col)];
    const changingCellModel = sheet.cells[cellKey(changingCell.row, changingCell.col)];
    if (!formulaCell?.formula) {
      showMessageBox(`${sheet.name}：数式入力セルには数式を指定してください。`, { title: "Cht WebSheet" });
      return true;
    }
    if (changingCellModel?.formula) {
      showMessageBox(`${sheet.name}：変化させるセルには数式を含まないセルを指定してください。`, { title: "Cht WebSheet" });
      return true;
    }
    const range = normalizeRange(
      { sheetIndex, row: changingCell.row, col: changingCell.col },
      { sheetIndex, row: changingCell.row, col: changingCell.col },
    );
    if (!ensureEditableRangeForProtection(sheet, range, "ゴール シーク")) return true;
    items.push({ sheetIndex, sheet, setCell, changingCell, range, ranges: [range] });
  }
  if (!items.length) return true;
  const affectedSheetIndexes = affectedSheetIndexesForGroupedRanges(items);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  try {
    const results = items.map((item) => runGoalSeek(item.setCell, validation.targetValue, item.changingCell, { commitHistory: false }));
    if (results.some((result) => !result)) throw new Error("一部のシートでゴール シークを実行できませんでした。");
    const activeItem = items.find((item) => item.sheetIndex === activeIndex) || items[0];
    const activeResult = results[items.indexOf(activeItem)] || results[0];
    state.hf = buildFormulaEngine(state.model);
    state.activeSheetIndex = activeIndex;
    state.selected = { sheetIndex: activeIndex, row: activeItem.changingCell.row, col: activeItem.changingCell.col };
    state.selectionRange = { ...activeItem.range, sheetIndex: activeIndex };
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
    state.selectedImage = null;
    state.selectedImages = [];
    pushHistory({
      sheetIndex: activeIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore,
      selectionAfter: state.selectionRange,
      selectionBeforeRanges,
      selectionAfterRanges: [state.selectionRange],
      label: `${items.length} シート ゴール シーク`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    const successCount = results.filter((result) => result?.success).length;
    setStatus(`${items.length} 個のシートでゴール シークを実行しました。`);
    return {
      ...activeResult,
      groupCount: items.length,
      successCount,
      setCell: workbookCellReference(activeIndex, activeItem.setCell.row, activeItem.setCell.col),
      changingCell: workbookCellReference(activeIndex, activeItem.changingCell.row, activeItem.changingCell.col),
    };
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`ゴール シークを実行できませんでした：${error.message || error}`);
    return true;
  }
}



function goalSeekValue(setSheetId, setCell, changingSheet, changingSheetId, changingCell, value) {
  writeCellContent(changingSheet, changingSheetId, changingCell.row, changingCell.col, value);
  const result = state.hf.getCellValue({ sheet: setSheetId, row: setCell.row - 1, col: setCell.col - 1 });
  return Number(result);
}



function goalSeekInitialSecondGuess(initial) {
  if (!Number.isFinite(initial) || initial === 0) return 1;
  return initial + Math.max(1, Math.abs(initial) * 0.1);
}



function goalSeekFallbackGuess(current, iteration) {
  const direction = iteration % 2 === 0 ? 1 : -1;
  const magnitude = 10 ** Math.min(6, Math.floor(iteration / 2));
  return (Number.isFinite(current) ? current : 0) + direction * magnitude;
}



function showGoalSeekResultDialog(result) {
  const groupLine = result.groupCount
    ? `対象シート数: ${result.groupCount} (解が見つかったシート: ${result.successCount ?? result.groupCount})\n`
    : "";
  showMessageBox(
    `${result.success ? "解が見つかりました。" : "近似値を入力しました。"}\n${groupLine}数式入力セル: ${result.setCell || ""}\n変化させるセル: ${result.changingCell || ""}\n目標値: ${formatGoalSeekNumber(result.targetValue)}\n現在の値: ${formatGoalSeekNumber(result.finalValue)}\n変化後の値: ${formatGoalSeekNumber(result.changingValue)}\n反復回数: ${result.iterations || 0}`,
    { title: "ゴール シークの結果" },
  );
}



function formatGoalSeekNumber(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return "";
  return String(Math.round(numeric * 1000000) / 1000000);
}



function openDataTableDialog() {
  const context = selectedDataRange("データ テーブル", { limit: 20000 });
  if (!context) return;
  if (context.range.bottom <= context.range.top || context.range.right <= context.range.left) {
    showMessageBox("データ テーブルを作成する範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  state.dataDialog = { type: "data-table", range: context.range };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog data-table-dialog" role="dialog" aria-modal="true" aria-labelledby="dataTableDialogTitle" data-data-dialog="data-table">
      <div class="excel-mini-titlebar" id="dataTableDialogTitle">データ テーブル</div>
      <div class="excel-mini-body data-dialog-body">
        ${dataCellFieldHtml("行の代入セル", "rowInput", "", "data-table")}
        ${dataCellFieldHtml("列の代入セル", "columnInput", "", "data-table")}
        <div class="data-dialog-note">選択範囲の上端行または左端列に数式、もう一方に代入値を配置してください。</div>
        <div class="data-dialog-summary data-table-summary" data-data-table-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncDataTableDialog(overlay.querySelector("[data-data-dialog='data-table']"));
  overlay.querySelector("[data-data-field='columnInput']")?.focus({ preventScroll: true });
}



function applyDataTableDialog(dialog) {
  const validation = dataTableDialogValidation(dialog);
  if (!validation.valid) {
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    syncDataTableDialog(dialog);
    return;
  }
  const grouped = applyDataTableForSheetGroup(validation);
  if (grouped === "pending") return;
  if (grouped === true) {
    closeDataDialog();
    return;
  }
  const execute = () => {
    if (applyDataTable(validation.range, { rowInput: validation.rowInput, columnInput: validation.columnInput })) closeDataDialog();
  };
  if (validation.overwriteCount > 0) {
    showMessageBox(`${validation.overwriteCount.toLocaleString("ja-JP")} 個のセルの内容が置き換えられます。続行しますか?`, {
      title: "Cht WebSheet",
      tone: "warning",
      buttons: [
        { label: "OK", result: "ok", callback: execute },
        { label: "キャンセル", result: "cancel" },
      ],
      cancelResult: "cancel",
    });
    return;
  }
  execute();
}



function handleDataTableDialogSync(event) {
  syncDataTableDialog(event.currentTarget.closest("[data-data-dialog='data-table']"));
}



function syncDataTableDialog(dialog) {
  if (!dialog) return;
  const validation = dataTableDialogValidation(dialog, { preview: true });
  const summary = dialog.querySelector("[data-data-table-summary]");
  if (summary) {
    summary.textContent = validation.message;
    summary.classList.toggle("is-warning", !validation.valid || Boolean(validation.warning));
  }
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = false;
}



function dataTableDialogValidation(dialog, options = {}) {
  const range = state.dataDialog?.range;
  const sheet = activeSheet();
  if (!range || !sheet) return { valid: false, message: "データ テーブルの範囲を選択してください。" };
  const rowInputText = dialog?.querySelector("[data-data-field='rowInput']")?.value || "";
  const columnInputText = dialog?.querySelector("[data-data-field='columnInput']")?.value || "";
  const rowInput = decodeWorkbookCell(rowInputText);
  const columnInput = decodeWorkbookCell(columnInputText);
  if (rowInputText.trim() && !rowInput) return { valid: false, message: "行の代入セルを正しく入力してください。" };
  if (columnInputText.trim() && !columnInput) return { valid: false, message: "列の代入セルを正しく入力してください。" };
  if (!rowInput && !columnInput) return { valid: false, message: "行の代入セルまたは列の代入セルを指定してください。" };
  if (range.bottom <= range.top || range.right <= range.left) {
    return { valid: false, message: "データ テーブルを作成する範囲を選択してください。" };
  }
  const rowInputTarget = dataTableInputTarget(rowInput);
  const columnInputTarget = dataTableInputTarget(columnInput);
  if ((rowInput && !rowInputTarget) || (columnInput && !columnInputTarget)) {
    return { valid: false, message: "代入セルを正しく入力してください。" };
  }
  const protectedInput = [rowInputTarget, columnInputTarget].filter(Boolean).find((target) => {
    const inputRange = normalizeRange(
      { sheetIndex: target.sheetIndex, row: target.row, col: target.col },
      { sheetIndex: target.sheetIndex, row: target.row, col: target.col },
    );
    return rangeHasLockedCellsForProtection(target.sheet, inputRange);
  });
  if (protectedInput) {
    return { valid: false, message: `保護されたシート上のロックされたセルは代入セルに指定できません: ${workbookCellReference(protectedInput.sheetIndex, protectedInput.row, protectedInput.col)}` };
  }
  if (rowInput && columnInput && rowInputTarget.snapshotKey === columnInputTarget.snapshotKey) {
    return { valid: false, message: "行の代入セルと列の代入セルには別のセルを指定してください。" };
  }
  const formulas = dataTableFormulaPoints(sheet, range, { rowInput, columnInput });
  if (!formulas.length) {
    const location = rowInput && columnInput ? "選択範囲の左上セル" : columnInput ? "選択範囲の上端行" : "選択範囲の左端列";
    return { valid: false, message: `${location}に数式がありません。` };
  }
  const resultCells = dataTableResultCells(range, { rowInput, columnInput }, formulas);
  const resultRange = dataTableResultRange(range, resultCells);
  if (resultRange && rangeHasLockedCellsForProtection(sheet, resultRange)) {
    return { valid: false, message: "保護されたシート上のロックされたセルにはデータ テーブルの結果を出力できません。" };
  }
  if (resultRange && mergesIntersectingRange(sheet, resultRange).length) {
    return { valid: false, message: "結合されたセルを含む範囲にはデータ テーブルの結果を出力できません。" };
  }
  const inputConflicts = dataTableInputResultConflicts([rowInputTarget, columnInputTarget], range, resultCells);
  if (inputConflicts.length) {
    return {
      valid: false,
      message: `代入セルはデータ テーブルの結果範囲外に指定してください: ${inputConflicts.join(", ")}`,
    };
  }
  const overwriteCount = countDataTableOverwriteCells(sheet, range, { rowInput, columnInput }, formulas);
  const resultCount = dataTableResultCellCount(range, { rowInput, columnInput }, formulas);
  const inputSummary = [
    rowInput ? `行: ${workbookCellReference(rowInput.sheetIndex ?? state.activeSheetIndex, rowInput.row, rowInput.col)}` : "",
    columnInput ? `列: ${workbookCellReference(columnInput.sheetIndex ?? state.activeSheetIndex, columnInput.row, columnInput.col)}` : "",
  ].filter(Boolean).join(" / ");
  const message = `${rangeToLabel(range)} に ${resultCount.toLocaleString("ja-JP")} 個の結果を作成します。数式: ${formulas.length.toLocaleString("ja-JP")} 個。${inputSummary}${overwriteCount ? ` 既存の ${overwriteCount.toLocaleString("ja-JP")} セルを置き換えます。` : ""}`;
  return { valid: true, warning: overwriteCount > 0 && options.preview, message, range, rowInput, columnInput, formulas, overwriteCount };
}



function applyDataTable(range, options, control = {}) {
  const sheetIndex = Number.isFinite(Number(control.sheetIndex)) ? Number(control.sheetIndex) : state.activeSheetIndex;
  const sheet = control.sheet || state.model?.sheets?.[sheetIndex] || activeSheet();
  if (!sheet) {
    showMessageBox("データ テーブルのシートを取得できません。", { title: "Cht WebSheet" });
    return false;
  }
  const sheetId = getSheetId(sheet.name);
  if (sheetId == null) {
    showMessageBox("データ テーブルのシートを取得できません。", { title: "Cht WebSheet" });
    return false;
  }
  if (control.commitHistory !== false && !ensureEditableRangeForProtection(sheet, range, "データ テーブルを作成")) return false;
  const sheetBefore = control.commitHistory === false ? null : captureSheetForHistory(sheet);
  const tableFormula = dataTableArrayFormula(options, sheetIndex);
  const rowInputTarget = dataTableInputTarget(options.rowInput);
  const columnInputTarget = dataTableInputTarget(options.columnInput);
  if ((options.rowInput && !rowInputTarget) || (options.columnInput && !columnInputTarget)) {
    showMessageBox("代入セルを正しく入力してください。", { title: "Cht WebSheet" });
    return false;
  }
  const inputSnapshots = [];
  [rowInputTarget, columnInputTarget].forEach((target) => {
    if (!target) return;
    const key = cellKey(target.row, target.col);
    const snapshotKey = `${target.sheetIndex}:${key}`;
    if (inputSnapshots.some((item) => item.snapshotKey === snapshotKey)) return;
    inputSnapshots.push({ snapshotKey, sheet: target.sheet, key, row: target.row, col: target.col, cell: cloneCellModel(target.sheet.cells[key] || null) });
  });

  const formulas = dataTableFormulaPoints(sheet, range, options);
  if (!formulas.length) {
    showMessageBox("データ テーブルに使用する数式が見つかりません。", { title: "Cht WebSheet" });
    return false;
  }
  const resultRange = dataTableResultRange(range, dataTableResultCells(range, options, formulas));
  if (resultRange && rangeHasLockedCellsForProtection(sheet, resultRange)) {
    showMessageBox("保護されたシート上のロックされたセルにはデータ テーブルの結果を出力できません。", { title: "Cht WebSheet" });
    return false;
  }
  if (resultRange && mergesIntersectingRange(sheet, resultRange).length) {
    showMessageBox("結合されたセルを含む範囲にはデータ テーブルの結果を出力できません。", { title: "Cht WebSheet" });
    return false;
  }

  try {
    if (options.rowInput && options.columnInput) {
      const formula = formulas[0];
      for (let row = range.top + 1; row <= range.bottom; row += 1) {
        const columnValue = cellDataTableInputValue(sheet, sheetId, row, range.left);
        writeCellContent(columnInputTarget.sheet, columnInputTarget.sheetId, columnInputTarget.row, columnInputTarget.col, columnValue);
        for (let col = range.left + 1; col <= range.right; col += 1) {
          const rowValue = cellDataTableInputValue(sheet, sheetId, range.top, col);
          writeCellContent(rowInputTarget.sheet, rowInputTarget.sheetId, rowInputTarget.row, rowInputTarget.col, rowValue);
          const result = dataTableFormulaResult(sheetId, formula);
          writeCellContent(sheet, sheetId, row, col, result);
          setDataTableFormulaForCell(sheet, row, col, tableFormula);
        }
      }
    } else if (options.columnInput) {
      for (let row = range.top + 1; row <= range.bottom; row += 1) {
        const inputValue = cellDataTableInputValue(sheet, sheetId, row, range.left);
        writeCellContent(columnInputTarget.sheet, columnInputTarget.sheetId, columnInputTarget.row, columnInputTarget.col, inputValue);
        formulas.forEach((formula) => {
          const result = dataTableFormulaResult(sheetId, formula);
          writeCellContent(sheet, sheetId, row, formula.col, result);
          setDataTableFormulaForCell(sheet, row, formula.col, tableFormula);
        });
      }
    } else if (options.rowInput) {
      for (let col = range.left + 1; col <= range.right; col += 1) {
        const inputValue = cellDataTableInputValue(sheet, sheetId, range.top, col);
        writeCellContent(rowInputTarget.sheet, rowInputTarget.sheetId, rowInputTarget.row, rowInputTarget.col, inputValue);
        formulas.forEach((formula) => {
          const result = dataTableFormulaResult(sheetId, formula);
          writeCellContent(sheet, sheetId, formula.row, col, result);
          setDataTableFormulaForCell(sheet, formula.row, col, tableFormula);
        });
      }
    }
  } finally {
    restoreDataTableInputCells(sheet, inputSnapshots);
    state.hf = buildFormulaEngine(state.model);
  }

  internCellsInRange(sheet, range);
  if (control.commitHistory === false) return true;
  commitDataOperationForSheet(sheetIndex, sheet, sheetBefore, range, `${rangeToLabel(range)} データ テーブル`, "データ テーブルを作成しました。", range);
  return true;
}



function applyDataTableForSheetGroup(validation, control = {}) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !validation?.range) return false;
  const activeIndex = state.activeSheetIndex;
  if ((validation.range.sheetIndex ?? activeIndex) !== activeIndex) return false;
  if (validation.rowInput && (validation.rowInput.sheetIndex ?? activeIndex) !== activeIndex) return false;
  if (validation.columnInput && (validation.columnInput.sheetIndex ?? activeIndex) !== activeIndex) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    const range = materializeSelectionRangeForSheet(sheet, validation.range, sheetIndex);
    if (!range) continue;
    const options = {
      rowInput: validation.rowInput ? { ...validation.rowInput, sheetIndex } : null,
      columnInput: validation.columnInput ? { ...validation.columnInput, sheetIndex } : null,
    };
    const rowInputTarget = dataTableInputTarget(options.rowInput);
    const columnInputTarget = dataTableInputTarget(options.columnInput);
    if ((options.rowInput && !rowInputTarget) || (options.columnInput && !columnInputTarget)) {
      showMessageBox(`${sheet.name}：代入セルを正しく入力してください。`, { title: "Cht WebSheet" });
      return true;
    }
    const protectedInput = [rowInputTarget, columnInputTarget].filter(Boolean).find((target) => {
      const inputRange = normalizeRange(
        { sheetIndex: target.sheetIndex, row: target.row, col: target.col },
        { sheetIndex: target.sheetIndex, row: target.row, col: target.col },
      );
      return rangeHasLockedCellsForProtection(target.sheet, inputRange);
    });
    if (protectedInput) {
      showMessageBox(`${sheet.name}：保護されたシート上のロックされたセルは代入セルに指定できません。`, { title: "Cht WebSheet" });
      return true;
    }
    if (rowInputTarget && columnInputTarget && rowInputTarget.snapshotKey === columnInputTarget.snapshotKey) {
      showMessageBox(`${sheet.name}：行の代入セルと列の代入セルには別のセルを指定してください。`, { title: "Cht WebSheet" });
      return true;
    }
    const formulas = dataTableFormulaPoints(sheet, range, options);
    if (!formulas.length) {
      showMessageBox(`${sheet.name}：データ テーブルに使用する数式が見つかりません。`, { title: "Cht WebSheet" });
      return true;
    }
    const resultCells = dataTableResultCells(range, options, formulas);
    const resultRange = dataTableResultRange(range, resultCells);
    if (resultRange && rangeHasLockedCellsForProtection(sheet, resultRange)) {
      showMessageBox(`${sheet.name}：保護されたシート上のロックされたセルにはデータ テーブルの結果を出力できません。`, { title: "Cht WebSheet" });
      return true;
    }
    if (resultRange && mergesIntersectingRange(sheet, resultRange).length) {
      showMessageBox(`${sheet.name}：結合されたセルを含む範囲にはデータ テーブルの結果を出力できません。`, { title: "Cht WebSheet" });
      return true;
    }
    const inputConflicts = dataTableInputResultConflicts([rowInputTarget, columnInputTarget], range, resultCells);
    if (inputConflicts.length) {
      showMessageBox(`${sheet.name}：代入セルはデータ テーブルの結果範囲外に指定してください: ${inputConflicts.join(", ")}`, { title: "Cht WebSheet" });
      return true;
    }
    if (!ensureEditableRangeForProtection(sheet, range, "データ テーブル")) return true;
    const overwriteCount = countDataTableOverwriteCells(sheet, range, options, formulas);
    items.push({ sheetIndex, sheet, range, options, formulas, overwriteCount, ranges: [range] });
  }
  if (!items.length) return true;
  const overwriteCount = items.reduce((total, item) => total + item.overwriteCount, 0);
  if (overwriteCount > 0 && !control.skipOverwriteCheck) {
    showMessageBox(`${overwriteCount.toLocaleString("ja-JP")} 個のセルの内容が置き換えられます。続行しますか?`, {
      title: "Cht WebSheet",
      tone: "warning",
      buttons: [
        {
          label: "OK",
          result: "ok",
          callback: () => {
            const result = applyDataTableForSheetGroup(validation, { ...control, skipOverwriteCheck: true });
            if (result === true) closeDataDialog();
          },
        },
        { label: "キャンセル", result: "cancel" },
      ],
      cancelResult: "cancel",
    });
    return "pending";
  }
  const affectedSheetIndexes = affectedSheetIndexesForGroupedRanges(items);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  try {
    items.forEach((item) => {
      applyDataTable(item.range, item.options, {
        sheet: item.sheet,
        sheetIndex: item.sheetIndex,
        commitHistory: false,
      });
    });
    state.hf = buildFormulaEngine(state.model);
    state.activeSheetIndex = activeIndex;
    const activeItem = items.find((item) => item.sheetIndex === activeIndex) || items[0];
    state.selected = { sheetIndex: activeIndex, row: activeItem.range.top, col: activeItem.range.left };
    state.selectionRange = { ...activeItem.range, sheetIndex: activeIndex };
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
    state.selectedImage = null;
    state.selectedImages = [];
    pushHistory({
      sheetIndex: activeIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore,
      selectionAfter: state.selectionRange,
      selectionBeforeRanges,
      selectionAfterRanges: [state.selectionRange],
      label: `${rangeToLabel(validation.range)} ${items.length} シートデータ テーブル`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートでデータ テーブルを作成しました。`);
    return true;
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`データ テーブルを作成できませんでした：${error.message || error}`);
    return true;
  }
}



function dataTableInputResultConflicts(targets, range, resultCells) {
  const resultSet = new Set((resultCells || []).map((point) => `${range.sheetIndex ?? state.activeSheetIndex}:${cellKey(point.row, point.col)}`));
  return (targets || [])
    .filter(Boolean)
    .filter((target) => resultSet.has(`${target.sheetIndex}:${cellKey(target.row, target.col)}`))
    .map((target) => workbookCellReference(target.sheetIndex, target.row, target.col));
}



function dataTableArrayFormula(options = {}, baseSheetIndex = state.activeSheetIndex) {
  const rowInput = options.rowInput ? dataTableInputFormulaReference(options.rowInput, baseSheetIndex) : "";
  const columnInput = options.columnInput ? dataTableInputFormulaReference(options.columnInput, baseSheetIndex) : "";
  return `{=TABLE(${rowInput},${columnInput})}`;
}



function dataTableInputFormulaReference(point, baseSheetIndex = state.activeSheetIndex) {
  const sheetIndex = point.sheetIndex ?? state.activeSheetIndex;
  if (sheetIndex !== baseSheetIndex) return workbookCellReference(sheetIndex, point.row, point.col);
  return `${columnName(point.col)}${point.row}`;
}



function dataTableInputTarget(point) {
  if (!point) return null;
  const sheetIndex = point.sheetIndex ?? state.activeSheetIndex;
  const sheet = state.model?.sheets?.[sheetIndex];
  const sheetId = getSheetId(sheet?.name);
  if (!sheet || sheetId == null) return null;
  return { ...point, sheetIndex, sheet, sheetId, snapshotKey: `${sheetIndex}:${cellKey(point.row, point.col)}` };
}



function setDataTableFormulaForCell(sheet, row, col, formula) {
  const cell = ensureCellModel(sheet, row, col);
  if (formula) cell.dataTableFormula = formula;
  storeCellModelForSheet(sheet, cellKey(row, col), cell);
}



function dataTableFormulaPoints(sheet, range, options) {
  if (options.rowInput && options.columnInput) {
    return sheet.cells[cellKey(range.top, range.left)]?.formula ? [{ row: range.top, col: range.left }] : [];
  }
  const points = [];
  if (options.columnInput) {
    for (let col = range.left + 1; col <= range.right; col += 1) {
      const cell = sheet.cells[cellKey(range.top, col)];
      if (cell?.formula) points.push({ row: range.top, col });
    }
  } else if (options.rowInput) {
    for (let row = range.top + 1; row <= range.bottom; row += 1) {
      const cell = sheet.cells[cellKey(row, range.left)];
      if (cell?.formula) points.push({ row, col: range.left });
    }
  }
  return points;
}



function dataTableResultCells(range, options, formulas) {
  const cells = [];
  if (options.rowInput && options.columnInput) {
    for (let row = range.top + 1; row <= range.bottom; row += 1) {
      for (let col = range.left + 1; col <= range.right; col += 1) cells.push({ row, col });
    }
    return cells;
  }
  if (options.columnInput) {
    for (let row = range.top + 1; row <= range.bottom; row += 1) {
      (formulas || []).forEach((formula) => cells.push({ row, col: formula.col }));
    }
    return cells;
  }
  if (options.rowInput) {
    for (let col = range.left + 1; col <= range.right; col += 1) {
      (formulas || []).forEach((formula) => cells.push({ row: formula.row, col }));
    }
  }
  return cells;
}



function dataTableResultCellCount(range, options, formulas) {
  return dataTableResultCells(range, options, formulas).length;
}



function dataTableResultRange(range, resultCells) {
  const cells = (resultCells || []).filter(Boolean);
  if (!range || !cells.length) return null;
  return normalizeRange(
    { sheetIndex: range.sheetIndex ?? state.activeSheetIndex, row: Math.min(...cells.map((cell) => cell.row)), col: Math.min(...cells.map((cell) => cell.col)) },
    { sheetIndex: range.sheetIndex ?? state.activeSheetIndex, row: Math.max(...cells.map((cell) => cell.row)), col: Math.max(...cells.map((cell) => cell.col)) },
  );
}



function countDataTableOverwriteCells(sheet, range, options, formulas) {
  return dataTableResultCells(range, options, formulas).filter(({ row, col }) => {
    const key = cellKey(row, col);
    const cell = sheet.cells[key];
    if (!cell) return false;
    if (cell.dataTableFormula) return false;
    return getCellRawInput(sheet, row, col) !== "";
  }).length;
}



function cellDataTableInputValue(sheet, sheetId, row, col) {
  const cell = sheet.cells[cellKey(row, col)];
  return cell?.raw ?? getDisplayForCell(sheet, sheetId, row, col);
}



function dataTableFormulaResult(sheetId, point) {
  const value = state.hf.getCellValue({ sheet: sheetId, row: point.row - 1, col: point.col - 1 });
  return typeof value === "object" && value ? String(value) : value;
}



function restoreDataTableInputCells(sheet, snapshots) {
  snapshots.forEach((snapshot) => {
    if (snapshot.cell) storeCellModelForSheet(snapshot.sheet || sheet, snapshot.key, cloneCellModel(snapshot.cell));
    else delete (snapshot.sheet || sheet).cells[snapshot.key];
  });
}



function openScenarioManagerDialog() {
  const scenarios = workbookScenarios();
  state.dataDialog = { type: "scenario-manager" };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = scenarioManagerDialogHtml(scenarios);
  document.body.appendChild(overlay);
}



function scenarioManagerDialogHtml(scenarios) {
  const rows = scenarios.length
    ? scenarios.map((scenario, index) => `<option value="${escapeAttr(scenario.id)}" ${index === 0 ? "selected" : ""}>${escapeHtml(scenario.name)}</option>`).join("")
    : '<option value="" disabled>シナリオはありません</option>';
  const first = scenarios[0];
  const canAdd = scenarioSheetAllowsEdit(state.activeSheetIndex);
  const canEditFirst = first && scenarioAllowsEdit(first);
  const canShowFirst = first && scenarioAllowsApply(first);
  return `
    <div class="excel-mini-dialog data-dialog scenario-manager-dialog" role="dialog" aria-modal="true" aria-labelledby="scenarioManagerTitle" data-data-dialog="scenario-manager">
      <div class="excel-mini-titlebar" id="scenarioManagerTitle">シナリオの登録と管理</div>
      <div class="excel-mini-body data-dialog-body scenario-manager-body">
        <div class="excel-dialog-caption">シナリオ</div>
        <div class="scenario-manager-grid">
          <select class="scenario-list" size="7" data-data-field="scenarioList">${rows}</select>
          <div class="scenario-manager-actions">
            <button class="excel-mini-button" type="button" data-data-dialog-action="scenario-show" ${canShowFirst ? "" : "disabled"}>表示</button>
            <button class="excel-mini-button" type="button" data-data-dialog-action="scenario-add" ${canAdd ? "" : "disabled"}>追加...</button>
            <button class="excel-mini-button" type="button" data-data-dialog-action="scenario-edit" ${canEditFirst ? "" : "disabled"}>編集...</button>
            <button class="excel-mini-button" type="button" data-data-dialog-action="scenario-delete" ${canEditFirst ? "" : "disabled"}>削除</button>
            <button class="excel-mini-button" type="button" data-data-dialog-action="scenario-summary" ${first ? "" : "disabled"}>結果の概要...</button>
          </div>
        </div>
        <div class="scenario-description" data-scenario-description>${escapeHtml(scenarioDescription(first))}</div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="cancel">閉じる</button>
      </div>
    </div>`;
}



function workbookScenarios() {
  if (!state.model) return [];
  if (!Array.isArray(state.model.scenarios)) state.model.scenarios = [];
  return state.model.scenarios;
}



function cloneScenarios(scenarios) {
  return (scenarios || []).map((scenario) => cloneJsonLikeValue(scenario));
}



function scenarioDescription(scenario) {
  if (!scenario) return "登録済みのシナリオはありません。";
  const sheetName = state.model?.sheets?.[scenario.sheetIndex]?.name || activeSheet()?.name || "";
  const cells = (scenario.cells || []).map((point) => scenarioPointLabel(point, scenario.sheetIndex)).join(", ");
  const permissionNote = !scenarioAllowsApply(scenario) ? "\nこのシナリオは保護されたシートでは表示できません。" : "";
  return `シート: ${sheetName}\n変化させるセル: ${cells || "(なし)"}${scenario.comment ? `\nコメント: ${scenario.comment}` : ""}${permissionNote}`;
}



function scenarioSheetIndex(scenario) {
  return Number.isFinite(Number(scenario?.sheetIndex))
    ? Number(scenario.sheetIndex)
    : Number.isFinite(Number(scenario?.cells?.[0]?.sheetIndex))
      ? Number(scenario.cells[0].sheetIndex)
      : state.activeSheetIndex;
}



function scenarioSheetAllowsEdit(sheetIndex = state.activeSheetIndex) {
  const sheet = state.model?.sheets?.[Number(sheetIndex)];
  if (!sheet) return false;
  return sheetProtectionAllows(sheet, "allowEditScenarios");
}



function scenarioAffectedSheetIndexes(scenario) {
  if (!scenario) return [state.activeSheetIndex];
  const indexes = (scenario.cells || [])
    .map((point) => point.sheetIndex ?? scenario.sheetIndex ?? state.activeSheetIndex)
    .filter((sheetIndex) => state.model?.sheets?.[sheetIndex]);
  return [...new Set(indexes.length ? indexes : [scenarioSheetIndex(scenario)])];
}



function scenarioAllowsEdit(scenario = null) {
  return scenarioSheetAllowsEdit(scenario ? scenarioSheetIndex(scenario) : state.activeSheetIndex);
}



function scenarioAllowsApply(scenario) {
  return scenarioAffectedSheetIndexes(scenario).every((sheetIndex) => scenarioSheetAllowsEdit(sheetIndex));
}



function ensureScenarioPermission(scenario = null, actionLabel = "シナリオを変更") {
  const allowed = scenario ? scenarioAllowsApply(scenario) : scenarioSheetAllowsEdit(state.activeSheetIndex);
  if (allowed) return true;
  showProtectedSheetMessage(`保護されたシートでは${actionLabel}できません。シート保護を解除するか、シナリオの編集を許可してください。`);
  return false;
}



function openScenarioSummaryDialog() {
  const scenarios = workbookScenarios();
  if (!scenarios.length) {
    showMessageBox("登録済みのシナリオはありません。", { title: "Cht WebSheet" });
    return;
  }
  const selection = activeSelectionRange();
  state.dataDialog = { type: "scenario-summary" };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog scenario-summary-dialog" role="dialog" aria-modal="true" aria-labelledby="scenarioSummaryTitle" data-data-dialog="scenario-summary">
      <div class="excel-mini-titlebar" id="scenarioSummaryTitle">シナリオの概要</div>
      <div class="excel-mini-body data-dialog-body">
        <fieldset class="data-dialog-fieldset">
          <legend>レポートの種類</legend>
          <label class="data-dialog-check"><input type="radio" name="scenarioReportType" data-data-field="reportType" data-scenario-summary-sync value="summary" checked />シナリオの概要</label>
          <label class="data-dialog-check"><input type="radio" name="scenarioReportType" data-data-field="reportType" data-scenario-summary-sync value="pivot" />シナリオ ピボットテーブル レポート</label>
        </fieldset>
        ${dataRangeFieldHtml("結果セル", "resultCells", selection ? rangeToLabel(selection) : "", "scenario-summary")}
        <div class="data-dialog-note">結果セルを指定すると、各シナリオを適用したときの値も概要に出力します。</div>
        <div class="data-dialog-summary scenario-summary-preview" data-scenario-summary-preview></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncScenarioSummaryDialog(overlay.querySelector("[data-data-dialog='scenario-summary']"));
  overlay.querySelector("[data-data-field='resultCells']")?.focus({ preventScroll: true });
}



function selectedScenarioFromDialog(dialog) {
  const id = dialog?.querySelector("[data-data-field='scenarioList']")?.value;
  return workbookScenarios().find((scenario) => scenario.id === id) || null;
}



function handleScenarioSummaryDialogSync(event) {
  syncScenarioSummaryDialog(event.currentTarget.closest("[data-data-dialog='scenario-summary']"));
}



function syncScenarioSummaryDialog(dialog) {
  if (!dialog) return;
  const validation = scenarioSummaryDialogValidation(dialog);
  const summary = dialog.querySelector("[data-scenario-summary-preview]");
  if (summary) {
    summary.textContent = validation.message;
    summary.classList.toggle("is-warning", !validation.valid || Boolean(validation.warning));
  }
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = !validation.valid;
}



function scenarioSummaryDialogValidation(dialog) {
  const scenarios = workbookScenarios();
  if (!scenarios.length) return { valid: false, message: "登録済みのシナリオはありません。" };
  const resultCellsText = dialog?.querySelector("[data-data-field='resultCells']")?.value || "";
  const reportType = dialog?.querySelector("[data-data-field='reportType']:checked")?.value || "summary";
  const resultCells = parseScenarioResultCells(resultCellsText);
  if (resultCellsText.trim() && !resultCells.length) {
    return { valid: false, message: "結果セルを正しく入力してください。" };
  }
  const changingCellCount = new Set(scenarios.flatMap((scenario) =>
    (scenario.cells || []).map((point) => scenarioPointKey({ ...point, sheetIndex: point.sheetIndex ?? scenario.sheetIndex ?? state.activeSheetIndex })),
  )).size;
  const message = `${scenarios.length.toLocaleString("ja-JP")} 個のシナリオ、${changingCellCount.toLocaleString("ja-JP")} 個の変更セルを出力します。結果セル: ${resultCells.length.toLocaleString("ja-JP")} 個。レポート: ${reportType === "pivot" ? "シナリオ ピボットテーブル レポート" : "シナリオの概要"}。`;
  return { valid: true, message, resultCellsText, reportType };
}



function handleScenarioListChange(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='scenario-manager']");
  if (!dialog) return;
  const scenario = selectedScenarioFromDialog(dialog);
  const description = dialog.querySelector("[data-scenario-description]");
  if (description) description.textContent = scenarioDescription(scenario);
  const disabled = !scenario;
  const canEdit = Boolean(scenario && scenarioAllowsEdit(scenario));
  const canShow = Boolean(scenario && scenarioAllowsApply(scenario));
  ["scenario-show", "scenario-edit", "scenario-delete", "scenario-summary"].forEach((action) => {
    const button = dialog.querySelector(`[data-data-dialog-action='${action}']`);
    if (!button) return;
    if (action === "scenario-show") button.disabled = !canShow;
    else if (action === "scenario-edit" || action === "scenario-delete") button.disabled = !canEdit;
    else button.disabled = disabled;
  });
  const addButton = dialog.querySelector("[data-data-dialog-action='scenario-add']");
  if (addButton) addButton.disabled = !scenarioSheetAllowsEdit(state.activeSheetIndex);
}



function openScenarioEditDialog(scenario = null) {
  if (!ensureScenarioPermission(scenario, scenario ? "シナリオを編集" : "シナリオを追加")) return;
  const sheet = activeSheet();
  const range = activeSelectionRange();
  const defaultCells = scenario
    ? scenario.cells || []
    : range
      ? rangeCellsForScenario(range)
      : state.selected
        ? [{ sheetIndex: state.selected.sheetIndex ?? state.activeSheetIndex, row: state.selected.row, col: state.selected.col }]
        : [{ sheetIndex: state.activeSheetIndex, row: 1, col: 1 }];
  const valuesText = scenarioValuesText(scenario, defaultCells);
  state.dataDialog = { type: "scenario-edit", scenarioId: scenario?.id || null };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog scenario-edit-dialog" role="dialog" aria-modal="true" aria-labelledby="scenarioEditTitle" data-data-dialog="scenario-edit">
      <div class="excel-mini-titlebar" id="scenarioEditTitle">${scenario ? "シナリオの編集" : "シナリオの追加"}</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">シナリオ名</span><input type="text" data-data-field="scenarioName" value="${escapeAttr(scenario?.name || "シナリオ 1")}" data-scenario-edit-sync /></label>
        ${dataRangeFieldHtml("変化させるセル", "scenarioCells", defaultCells.map((point) => scenarioPointLabel(point, state.activeSheetIndex)).join(","), "scenario-edit")}
        <label class="excel-dialog-field"><span class="excel-dialog-caption">値</span><textarea rows="6" data-data-field="scenarioValues" data-scenario-edit-sync>${escapeHtml(valuesText)}</textarea></label>
        <div class="data-dialog-inline-actions">
          <button class="excel-mini-button" type="button" data-data-dialog-action="scenario-load-values">現在値を読み込み</button>
        </div>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">コメント</span><textarea rows="3" data-data-field="scenarioComment" data-scenario-edit-sync>${escapeHtml(scenario?.comment || "")}</textarea></label>
        <fieldset class="data-dialog-fieldset">
          <legend>保護</legend>
          <label class="data-dialog-check"><input type="checkbox" data-data-field="scenarioLocked" data-scenario-edit-sync ${scenario?.locked === false ? "" : "checked"} />変更できないようにする</label>
          <label class="data-dialog-check"><input type="checkbox" data-data-field="scenarioHidden" data-scenario-edit-sync ${scenario?.hidden ? "checked" : ""} />表示しない</label>
        </fieldset>
        <div class="data-dialog-summary scenario-edit-summary" data-scenario-edit-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="scenario-save">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="scenario-manager">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncScenarioEditDialog(overlay.querySelector("[data-data-dialog='scenario-edit']"));
  overlay.querySelector("[data-data-field='scenarioName']")?.focus({ preventScroll: true });
}



function rangeCellsForScenario(range) {
  const cells = [];
  forEachCellInRange(range, (row, col) => {
    if (cells.length < 32) cells.push({ sheetIndex: range.sheetIndex ?? state.activeSheetIndex, row, col });
  });
  return cells;
}



function handleScenarioEditDialogSync(event) {
  syncScenarioEditDialog(event.currentTarget.closest("[data-data-dialog='scenario-edit']"));
}



function syncScenarioEditDialog(dialog) {
  if (!dialog) return;
  const validation = scenarioEditDialogValidation(dialog);
  const summary = dialog.querySelector("[data-scenario-edit-summary]");
  if (summary) {
    summary.textContent = validation.message;
    summary.classList.toggle("is-warning", !validation.valid || Boolean(validation.warning));
  }
  const ok = dialog.querySelector("[data-data-dialog-action='scenario-save']");
  if (ok) ok.disabled = false;
}



function scenarioEditDialogValidation(dialog) {
  const name = dialog?.querySelector("[data-data-field='scenarioName']")?.value.trim() || "";
  const cellsText = dialog?.querySelector("[data-data-field='scenarioCells']")?.value || "";
  const valuesText = dialog?.querySelector("[data-data-field='scenarioValues']")?.value || "";
  if (!name) return { valid: false, message: "シナリオ名を入力してください。" };
  const duplicate = workbookScenarios().some((item) =>
    item.id !== state.dataDialog?.scenarioId &&
    String(item.name || "").trim().toLocaleLowerCase("ja-JP") === name.toLocaleLowerCase("ja-JP"),
  );
  if (duplicate) return { valid: false, message: "同じ名前のシナリオが既に存在します。別の名前を入力してください。" };
  const cellStats = scenarioCellStats(cellsText);
  if (cellStats.invalid.length) {
    return { valid: false, message: `変化させるセルを正しく入力してください: ${cellStats.invalid.slice(0, 3).join(", ")}` };
  }
  if (cellStats.count > 32) {
    return { valid: false, message: "変化させるセルは最大 32 個までです。" };
  }
  const cells = parseScenarioCells(cellsText);
  if (!cells.length) return { valid: false, message: "変化させるセルを指定してください。" };
  const valueLines = String(valuesText || "").split(/\r?\n/).map((line) => line.trim()).filter(Boolean).length;
  const hasAddressedValues = String(valuesText || "").split(/\r?\n/).some((line) => line.includes("="));
  const warning = Boolean(valueLines && !hasAddressedValues && valueLines !== cells.length);
  const hidden = Boolean(dialog?.querySelector("[data-data-field='scenarioHidden']")?.checked);
  const locked = Boolean(dialog?.querySelector("[data-data-field='scenarioLocked']")?.checked);
  const valueWarning = warning ? " 値の行数が変更セル数と一致しません。空欄のセルには空白が設定されます。" : "";
  const message = `${cells.length} 個のセルを変更します。値の入力行: ${valueLines.toLocaleString("ja-JP")} 行。${locked ? "保護: 変更不可。" : "保護: 変更可。"}${hidden ? " このシナリオは非表示です。" : ""}${valueWarning}`;
  return { valid: true, warning, message, name, cells };
}



function loadScenarioCurrentValues(dialog) {
  const cells = parseScenarioCells(dialog?.querySelector("[data-data-field='scenarioCells']")?.value || "");
  if (!cells.length) {
    showMessageBox("変化させるセルを指定してください。", { title: "Cht WebSheet" });
    return;
  }
  const textarea = dialog.querySelector("[data-data-field='scenarioValues']");
  if (!textarea) return;
  textarea.value = scenarioValuesText(null, cells);
  syncScenarioEditDialog(dialog);
}



function scenarioValuesText(scenario, cells) {
  return (cells || [])
    .map((point) => {
      const sheetIndex = point.sheetIndex ?? scenario?.sheetIndex ?? state.activeSheetIndex;
      const sheet = state.model?.sheets?.[sheetIndex] || activeSheet();
      const sheetId = getSheetId(sheet?.name);
      const address = scenarioPointLabel({ ...point, sheetIndex }, state.activeSheetIndex);
      const value = scenarioValueForPoint(scenario, { ...point, sheetIndex }) ?? (sheet && sheetId != null ? cellDataTableInputValue(sheet, sheetId, point.row, point.col) : "") ?? "";
      return `${address}=${value}`;
    })
    .join("\n");
}



function saveScenarioFromDialog(dialog) {
  const existingScenario = workbookScenarios().find((item) => item.id === state.dataDialog?.scenarioId) || null;
  if (!ensureScenarioPermission(existingScenario, "シナリオを保存")) return;
  const validation = scenarioEditDialogValidation(dialog);
  if (!validation.valid) {
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    syncScenarioEditDialog(dialog);
    return;
  }
  const { name, cells } = validation;
  const values = parseScenarioValues(dialog.querySelector("[data-data-field='scenarioValues']")?.value || "", cells);
  const scenarios = workbookScenarios();
  const scenariosBefore = cloneScenarios(scenarios);
  const id = state.dataDialog?.scenarioId || uniqueWorksheetObjectId("scenario");
  const index = scenarios.findIndex((item) => item.id === id);
  const next = {
    id,
    name,
    sheetIndex: state.activeSheetIndex,
    cells,
    values,
    comment: dialog.querySelector("[data-data-field='scenarioComment']")?.value || "",
    locked: Boolean(dialog.querySelector("[data-data-field='scenarioLocked']")?.checked),
    hidden: Boolean(dialog.querySelector("[data-data-field='scenarioHidden']")?.checked),
  };
  if (index >= 0) scenarios[index] = next;
  else scenarios.push(next);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    scenariosBefore,
    scenariosAfter: cloneScenarios(scenarios),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: `シナリオ「${name}」${index >= 0 ? "編集" : "追加"}`,
  });
  scheduleWorkbookPersistence();
  openScenarioManagerDialog();
}



function parseScenarioCells(text) {
  return scenarioCellStats(text).cells;
}



function scenarioCellStats(text) {
  const cells = [];
  const invalid = [];
  const seen = new Set();
  let count = 0;
  splitWorkbookReferenceList(text).forEach((token) => {
    const range = decodeWorkbookRangeOrName(token, activeSheet());
    if (!range) {
      invalid.push(token);
      return;
    }
    const sheetIndex = range.sheetIndex ?? state.activeSheetIndex;
    forEachCellInRange(range, (row, col) => {
      const key = `${sheetIndex}:${cellKey(row, col)}`;
      if (seen.has(key)) return;
      seen.add(key);
      count += 1;
      if (cells.length < 32) cells.push({ sheetIndex, row, col });
    });
  });
  return { cells, count, invalid };
}



function splitWorkbookReferenceList(text) {
  const tokens = [];
  let current = "";
  let inQuotedSheet = false;
  const source = String(text || "");
  const push = () => {
    const token = current.trim();
    if (token) tokens.push(token);
    current = "";
  };
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char === "'") {
      if (inQuotedSheet && source[index + 1] === "'") {
        current += "''";
        index += 1;
        continue;
      }
      inQuotedSheet = !inQuotedSheet;
      current += char;
      continue;
    }
    if (!inQuotedSheet && /[,\s;]/.test(char)) {
      push();
      continue;
    }
    current += char;
  }
  push();
  return tokens;
}



function parseScenarioValues(text, cells) {
  const values = {};
  const positional = [];
  String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .forEach((line) => {
      const equals = line.indexOf("=");
      if (equals > 0) {
        const left = line.slice(0, equals).trim();
        const value = coerceUserInput(line.slice(equals + 1));
        const range = decodeWorkbookRangeOrName(left, activeSheet());
        if (range) {
          const sheetIndex = range.sheetIndex ?? state.activeSheetIndex;
          forEachCellInRange(range, (row, col) => {
            setScenarioValueForPoint(values, { sheetIndex, row, col }, value);
          });
        } else {
          const address = left.replace(/\$/g, "").toUpperCase();
          if (decodeCell(address)) values[address] = value;
        }
      } else {
        positional.push(coerceUserInput(line));
      }
    });
  cells.forEach((point, index) => {
    if (scenarioValueForPoint({ values }, point) === undefined) {
      setScenarioValueForPoint(values, point, positional[index] ?? "");
    }
  });
  return values;
}



function applyScenario(scenario) {
  if (!scenario) return;
  if (!ensureScenarioPermission(scenario, "シナリオを表示")) return;
  const first = scenario.cells?.[0] || { row: 1, col: 1 };
  const last = scenario.cells?.[scenario.cells.length - 1] || first;
  const affectedSheetIndexes = [...new Set((scenario.cells || []).map((point) => point.sheetIndex ?? scenario.sheetIndex ?? state.activeSheetIndex))]
    .filter((sheetIndex) => state.model?.sheets?.[sheetIndex]);
  if (!affectedSheetIndexes.length) return;
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  (scenario.cells || []).forEach((point) => {
    const sheetIndex = point.sheetIndex ?? scenario.sheetIndex ?? state.activeSheetIndex;
    const sheet = state.model.sheets[sheetIndex];
    const sheetId = getSheetId(sheet?.name);
    if (!sheet || sheetId == null) return;
    writeCellContent(sheet, sheetId, point.row, point.col, scenarioValueForPoint(scenario, { ...point, sheetIndex }) ?? "");
  });
  const firstSheetIndex = first.sheetIndex ?? scenario.sheetIndex ?? state.activeSheetIndex;
  state.activeSheetIndex = firstSheetIndex;
  const range = (last.sheetIndex == null || last.sheetIndex === firstSheetIndex)
    ? normalizeRange({ sheetIndex: firstSheetIndex, ...first }, { sheetIndex: firstSheetIndex, ...last })
    : normalizeRange({ sheetIndex: firstSheetIndex, ...first }, { sheetIndex: firstSheetIndex, ...first });
  state.selected = { sheetIndex: firstSheetIndex, row: first.row, col: first.col };
  state.selectionRange = { ...range };
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.hf = buildFormulaEngine(state.model);
  renderWorkbook();
  const selectionAfter = { ...range };
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(affectedSheetIndexes),
    selectionBefore,
    selectionAfter,
    selectionBeforeRanges,
    selectionAfterRanges: [selectionAfter],
    label: `シナリオ ${scenario.name}`,
  });
  setStatus(`シナリオ「${scenario.name}」を表示しました。`);
}



function createScenarioSummarySheet(options = {}) {
  const scenarios = workbookScenarios();
  if (!scenarios.length) {
    showMessageBox("登録済みのシナリオはありません。", { title: "Cht WebSheet" });
    return;
  }
  const cells = [];
  const seen = new Set();
  scenarios.forEach((scenario) => {
    (scenario.cells || []).forEach((point) => {
      const sheetIndex = point.sheetIndex ?? scenario.sheetIndex ?? state.activeSheetIndex;
      const key = `${sheetIndex}:${cellKey(point.row, point.col)}`;
      if (seen.has(key)) return;
      seen.add(key);
      cells.push({ sheetIndex, row: point.row, col: point.col });
    });
  });
  const resultCells = parseScenarioResultCells(options.resultCellsText);
  const valueMaps = scenarioSummaryValueMaps(scenarios, cells, resultCells);
  if (options.reportType === "pivot") {
    createScenarioPivotSummarySheet({ scenarios, changingCells: cells, resultCells, valueMaps });
    return;
  }
  const reportType = options.reportType === "pivot" ? "ピボット" : "概要";
  const baseName = options.reportType === "pivot" ? "シナリオ ピボット" : "シナリオの概要";
  const name = uniqueSheetNameWithBase(baseName, state.model.sheets.map((item) => item.name));
  const summary = createBlankSheet(name);
  const totalRows = Math.max(4, cells.length + resultCells.length + (resultCells.length ? 6 : 4));
  ensureSheetSize(summary, totalRows, scenarios.length + 3);
  updateCellModel(summary, 1, 1, `シナリオの${reportType}`);
  updateCellModel(summary, 3, 1, "変化させるセル");
  updateCellModel(summary, 3, 2, "現在値");
  scenarios.forEach((scenario, index) => updateCellModel(summary, 3, index + 3, scenario.name));
  cells.forEach((point, rowIndex) => {
    const row = rowIndex + 4;
    const sourceSheet = state.model.sheets[point.sheetIndex];
    const address = scenarioSummaryAddress(sourceSheet, point);
    updateCellModel(summary, row, 1, address);
    updateCellModel(summary, row, 2, valueMaps.current.get(scenarioPointKey(point)) ?? "");
    scenarios.forEach((scenario, scenarioIndex) => {
      const value = scenarioValueForPoint(scenario, point) ?? "";
      updateCellModel(summary, row, scenarioIndex + 3, value);
    });
  });
  let lastRow = cells.length + 3;
  if (resultCells.length) {
    const headerRow = lastRow + 2;
    updateCellModel(summary, headerRow, 1, "結果セル");
    updateCellModel(summary, headerRow, 2, "現在値");
    scenarios.forEach((scenario, index) => updateCellModel(summary, headerRow, index + 3, scenario.name));
    resultCells.forEach((point, index) => {
      const row = headerRow + index + 1;
      const sourceSheet = state.model.sheets[point.sheetIndex];
      updateCellModel(summary, row, 1, scenarioSummaryAddress(sourceSheet, point));
      updateCellModel(summary, row, 2, valueMaps.current.get(scenarioPointKey(point)) ?? "");
      scenarios.forEach((scenario, scenarioIndex) => {
        updateCellModel(summary, row, scenarioIndex + 3, valueMaps.scenarios.get(scenario.id)?.get(scenarioPointKey(point)) ?? "");
      });
    });
    lastRow = headerRow + resultCells.length;
  }
  applyExcelTableStyle(summary, normalizeRange({ sheetIndex: state.model.sheets.length, row: 3, col: 1 }, { sheetIndex: state.model.sheets.length, row: lastRow, col: scenarios.length + 2 }));
  summary.columns[0].width = 180;
  for (let col = 2; col <= scenarios.length + 2; col += 1) summary.columns[col - 1].width = 120;
  const activeSheetBefore = state.activeSheetIndex;
  const insertIndex = state.model.sheets.length;
  state.model.sheets.push(summary);
  state.activeSheetIndex = insertIndex;
  state.hf = buildFormulaEngine(state.model);
  renderWorkbook();
  setStatus(`${name} を作成しました。`);
  pushSheetInsertHistory(insertIndex, summary, activeSheetBefore, insertIndex, `${name} 作成`);
  scheduleWorkbookPersistence();
}



function createScenarioPivotSummarySheet({ scenarios, changingCells, resultCells, valueMaps }) {
  const points = [...changingCells, ...resultCells];
  const name = uniqueSheetNameWithBase("シナリオ ピボット", state.model.sheets.map((item) => item.name));
  const summary = createBlankSheet(name);
  const pointColumnCount = Math.max(1, points.length);
  const totalRows = Math.max(6, scenarios.length + 5);
  const totalCols = Math.max(3, pointColumnCount + 1);
  ensureSheetSize(summary, totalRows, totalCols);
  updateCellModel(summary, 1, 1, "シナリオ ピボットテーブル レポート");
  updateCellModel(summary, 2, 1, "値");
  if (changingCells.length) updateCellModel(summary, 2, 2, "変化させるセル");
  if (resultCells.length) updateCellModel(summary, 2, 2 + changingCells.length, "結果セル");
  updateCellModel(summary, 3, 1, "シナリオ");
  points.forEach((point, index) => {
    const sourceSheet = state.model.sheets[point.sheetIndex];
    updateCellModel(summary, 3, index + 2, scenarioSummaryAddress(sourceSheet, point));
  });
  updateCellModel(summary, 4, 1, "現在値");
  points.forEach((point, index) => {
    updateCellModel(summary, 4, index + 2, valueMaps.current.get(scenarioPointKey(point)) ?? "");
  });
  scenarios.forEach((scenario, scenarioIndex) => {
    const row = scenarioIndex + 5;
    updateCellModel(summary, row, 1, scenario.name);
    changingCells.forEach((point, index) => {
      updateCellModel(summary, row, index + 2, scenarioValueForPoint(scenario, point) ?? valueMaps.current.get(scenarioPointKey(point)) ?? "");
    });
    resultCells.forEach((point, index) => {
      updateCellModel(summary, row, changingCells.length + index + 2, valueMaps.scenarios.get(scenario.id)?.get(scenarioPointKey(point)) ?? "");
    });
  });
  styleScenarioReportSheet(summary, {
    titleRow: 1,
    headerRange: { top: 3, left: 1, bottom: totalRows, right: totalCols },
    groupRow: 2,
    groupStartCol: 2,
    groupEndCol: totalCols,
  });
  summary.columns[0].width = 170;
  for (let col = 2; col <= totalCols; col += 1) summary.columns[col - 1].width = 135;
  const activeSheetBefore = state.activeSheetIndex;
  const insertIndex = state.model.sheets.length;
  state.model.sheets.push(summary);
  state.activeSheetIndex = insertIndex;
  state.hf = buildFormulaEngine(state.model);
  renderWorkbook();
  setStatus(`${name} を作成しました。`);
  pushSheetInsertHistory(insertIndex, summary, activeSheetBefore, insertIndex, `${name} 作成`);
  scheduleWorkbookPersistence();
}



function styleScenarioReportSheet(sheet, options = {}) {
  const titleCell = ensureCellModel(sheet, options.titleRow || 1, 1);
  titleCell.css.fontWeight = "700";
  titleCell.css.fontSize = "14px";
  titleCell.css.color = "#1f2937";
  const groupRow = options.groupRow;
  if (groupRow) {
    for (let col = options.groupStartCol || 1; col <= (options.groupEndCol || sheet.colCount); col += 1) {
      const cell = ensureCellModel(sheet, groupRow, col);
      cell.css.backgroundColor = "#d9eaf7";
      cell.css.fontWeight = "700";
      cell.css.textAlign = "center";
      cell.hasFill = true;
    }
  }
  if (options.headerRange) applyExcelTableStyle(sheet, options.headerRange);
  internCellsInRange(sheet, { top: 1, left: 1, bottom: Math.max(sheet.rowCount, options.headerRange?.bottom || 1), right: Math.max(sheet.colCount, options.headerRange?.right || 1) });
}



function parseScenarioResultCells(text) {
  const cells = [];
  const seen = new Set();
  splitWorkbookReferenceList(text).forEach((token) => {
    const range = decodeWorkbookRangeOrName(token, activeSheet());
    if (!range) return;
    const sheetIndex = range.sheetIndex ?? state.activeSheetIndex;
    forEachCellInRange(range, (row, col) => {
      const point = { sheetIndex, row, col };
      const key = scenarioPointKey(point);
      if (seen.has(key)) return;
      seen.add(key);
      cells.push(point);
    });
  });
  return cells;
}



function openDataAnalysisDialog() {
  if (!state.model) {
    setStatus("Excel ファイルを開いてからデータ分析を実行してください。");
    return;
  }
  const context = selectedDataRange("データ分析", { expandCurrentRegion: true, limit: 120000 });
  if (!context) return;
  const method = state.dataDialog?.method || "descriptive";
  state.dataDialog = dataAnalysisDefaultDialogState(context, method);
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.id = "dataDialog";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = dataAnalysisPickerDialogHtml(method);
  document.body.appendChild(overlay);
  syncDataAnalysisPickerDialog(overlay.querySelector("[data-data-dialog='data-analysis-picker']"));
  overlay.querySelector("[data-data-analysis-picker]")?.focus({ preventScroll: true });
}



function dataAnalysisDefaultDialogState(context, method = "descriptive") {
  return {
    type: "data-analysis",
    method,
    inputRange: workbookRangeReference(state.activeSheetIndex, context.range),
    inputYRange: dataAnalysisDefaultYRange(context.range),
    inputXRange: dataAnalysisDefaultXRange(context.range),
    variable1Range: dataAnalysisDefaultVariableRange(context.range, 0),
    variable2Range: dataAnalysisDefaultVariableRange(context.range, 1),
    outputCell: dataAnalysisDefaultOutputCell(context.range),
  };
}



function dataAnalysisPickerDialogHtml(method = "descriptive") {
  const methodOptions = DATA_ANALYSIS_METHODS.map((item) =>
    `<option value="${escapeAttr(item.value)}"${item.value === method ? " selected" : ""}>${escapeHtml(item.label)}</option>`,
  ).join("");
  return `
    <div class="excel-mini-dialog data-dialog data-analysis-picker-dialog" role="dialog" aria-modal="true" aria-labelledby="dataAnalysisPickerTitle" data-data-dialog="data-analysis-picker">
      <div class="excel-mini-titlebar" id="dataAnalysisPickerTitle">データ分析</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="excel-dialog-field data-analysis-picker-field">
          <span class="excel-dialog-caption">分析ツール(A):</span>
          <select class="data-analysis-list" size="16" data-data-field="method" data-data-analysis-picker>${methodOptions}</select>
        </label>
        <div class="data-dialog-summary data-analysis-summary" data-data-analysis-picker-summary></div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="data-analysis-picker-ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="data-analysis-help">ヘルプ(H)</button>
      </div>
    </div>`;
}



function openDataAnalysisOptionsDialog(method, defaults = state.dataDialog || {}) {
  if (!state.model) return;
  const context = selectedDataRange("データ分析", { expandCurrentRegion: true, limit: 120000, quiet: true });
  const fallbackState = context ? dataAnalysisDefaultDialogState(context, method) : {};
  state.dataDialog = {
    ...fallbackState,
    ...defaults,
    type: "data-analysis",
    method,
  };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.id = "dataDialog";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = dataAnalysisDialogHtml(state.dataDialog);
  document.body.appendChild(overlay);
  syncDataAnalysisDialog(overlay.querySelector("[data-data-dialog='data-analysis']"));
  overlay.querySelector("[data-data-field='method']")?.focus({ preventScroll: true });
}



function selectedDataAnalysisPickerMethod(dialog) {
  return dialog?.querySelector("[data-data-analysis-picker]")?.value || state.dataDialog?.method || "descriptive";
}



function syncDataAnalysisPickerDialog(dialog) {
  if (!dialog) return;
  const method = selectedDataAnalysisPickerMethod(dialog);
  state.dataDialog = { ...(state.dataDialog || {}), method };
  const summary = dialog.querySelector("[data-data-analysis-picker-summary]");
  if (summary) summary.textContent = `${dataAnalysisMethodLabel(method)} のオプションを設定します。`;
}



function handleDataAnalysisPickerSync(event) {
  syncDataAnalysisPickerDialog(event.currentTarget.closest("[data-data-dialog='data-analysis-picker']"));
}



function handleDataAnalysisPickerDoubleClick(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='data-analysis-picker']");
  const method = selectedDataAnalysisPickerMethod(dialog);
  openDataAnalysisOptionsDialog(method, state.dataDialog);
}



function dataAnalysisDefaultOutputCell(range) {
  const col = Math.max(1, Math.min((range?.right || 1) + 2, (activeSheet()?.colCount || 1) + 2));
  const row = Math.max(1, range?.top || 1);
  return workbookCellReference(state.activeSheetIndex, row, col);
}



function dataAnalysisDefaultYRange(range) {
  if (!range) return "";
  return workbookRangeReference(state.activeSheetIndex, {
    top: range.top,
    bottom: range.bottom,
    left: range.left,
    right: range.left,
  });
}



function dataAnalysisDefaultXRange(range) {
  if (!range) return "";
  const left = Math.min(range.right, range.left + 1);
  return workbookRangeReference(state.activeSheetIndex, {
    top: range.top,
    bottom: range.bottom,
    left,
    right: Math.max(left, range.right),
  });
}



function dataAnalysisDefaultVariableRange(range, offset = 0) {
  if (!range) return "";
  const col = Math.min(range.right, range.left + offset);
  return workbookRangeReference(state.activeSheetIndex, {
    top: range.top,
    bottom: range.bottom,
    left: col,
    right: col,
  });
}



function dataAnalysisDialogStateValue(dialogState, field, fallback = "") {
  const value = dialogState?.[field];
  return value === undefined || value === null ? fallback : value;
}



function dataAnalysisCheckedAttr(dialogState, field, fallback = false) {
  return Boolean(dataAnalysisDialogStateValue(dialogState, field, fallback)) ? " checked" : "";
}



function dataAnalysisRadioCheckedAttr(dialogState, field, value, fallback) {
  return String(dataAnalysisDialogStateValue(dialogState, field, fallback)) === String(value) ? " checked" : "";
}



function dataAnalysisSelectedAttr(dialogState, field, value, fallback) {
  return String(dataAnalysisDialogStateValue(dialogState, field, fallback)) === String(value) ? " selected" : "";
}



function dataAnalysisNumberValue(dialogState, field, fallback) {
  return escapeAttr(dataAnalysisDialogStateValue(dialogState, field, fallback));
}



function dataAnalysisDialogStateFromOptions(options = {}) {
  return {
    type: "data-analysis",
    method: options.method || "descriptive",
    inputRange: options.inputText ?? options.inputRange ?? "",
    inputYRange: options.inputYText ?? options.inputYRange ?? "",
    inputXRange: options.inputXText ?? options.inputXRange ?? "",
    variable1Range: options.variable1Text ?? options.variable1Range ?? "",
    variable2Range: options.variable2Text ?? options.variable2Range ?? "",
    outputCell: options.outputText ?? options.outputCell ?? "",
    outputMode: options.outputMode || "range",
    outputSheetName: options.outputSheetName || "",
    binRange: options.binText ?? options.binRange ?? "",
    groupedBy: options.groupedBy || "columns",
    labels: Boolean(options.labels),
    confidence: options.confidence ?? 95,
    alpha: options.alpha ?? 0.05,
    interval: options.interval ?? 3,
    sampleCount: options.sampleCount ?? 10,
    samplingMethod: options.samplingMethod || "periodic",
    randomRows: options.randomRows ?? 10,
    randomCols: options.randomCols ?? 2,
    mean: options.mean ?? 0,
    stddev: options.stddev ?? 1,
    lowerBound: options.lowerBound ?? 0,
    upperBound: options.upperBound ?? 1,
    patternStep: options.patternStep ?? 1,
    probabilityRange: options.probabilityRangeText ?? options.probabilityRange ?? "",
    randomDistribution: options.randomDistribution || "normal",
    randomSeed: options.randomSeed || "",
    hypothesizedMeanDiff: options.hypothesizedMeanDiff ?? 0,
    knownVariance1: options.knownVariance1 ?? 1,
    knownVariance2: options.knownVariance2 ?? 1,
    kthLargest: options.kthLargest ?? 1,
    kthSmallest: options.kthSmallest ?? 1,
    summaryStats: options.summaryStats !== false,
    confidenceLevelEnabled: options.confidenceLevelEnabled !== false,
    kthLargestEnabled: Boolean(options.kthLargestEnabled),
    kthSmallestEnabled: Boolean(options.kthSmallestEnabled),
    chartOutput: Boolean(options.chartOutput),
    cumulativePercent: options.cumulativePercent !== false,
    pareto: Boolean(options.pareto),
    standardErrorOutput: Boolean(options.standardErrorOutput),
    constantZero: Boolean(options.constantZero),
    residuals: Boolean(options.residuals),
    standardizedResiduals: Boolean(options.standardizedResiduals),
    residualPlot: Boolean(options.residualPlot),
    lineFitPlot: Boolean(options.lineFitPlot),
    normalProbabilityPlot: Boolean(options.normalProbabilityPlot),
    inverseFourier: Boolean(options.inverseFourier),
  };
}



function persistDataAnalysisDialogState(dialog) {
  if (!dialog) return state.dataDialog || {};
  const options = dataAnalysisDialogOptions(dialog, { preview: true });
  state.dataDialog = {
    ...(state.dataDialog || {}),
    ...dataAnalysisDialogStateFromOptions(options),
  };
  return state.dataDialog;
}



function dataAnalysisDialogHtml(dialogState) {
  const methodLabel = dataAnalysisMethodLabel(dialogState.method);
  const outputMode = dataAnalysisDialogStateValue(dialogState, "outputMode", "range");
  const groupedBy = dataAnalysisDialogStateValue(dialogState, "groupedBy", "columns");
  const samplingMethod = dataAnalysisDialogStateValue(dialogState, "samplingMethod", "periodic");
  return `
    <div class="excel-mini-dialog data-dialog data-analysis-dialog data-analysis-options-dialog" role="dialog" aria-modal="true" aria-labelledby="dataAnalysisTitle" data-data-dialog="data-analysis">
      <div class="excel-mini-titlebar" id="dataAnalysisTitle">${escapeHtml(methodLabel)}</div>
      <div class="excel-mini-body data-dialog-body">
        <input type="hidden" data-data-field="method" value="${escapeAttr(dialogState.method)}" />
        <div class="data-analysis-options">
            <div class="data-analysis-section-title" data-data-analysis-method-title></div>
            ${dataRangeFieldHtml("入力範囲", "inputRange", dialogState.inputRange, "data-analysis")}
            ${dataRangeFieldHtml("入力 Y 範囲", "inputYRange", dialogState.inputYRange || "", "data-analysis")}
            ${dataRangeFieldHtml("入力 X 範囲", "inputXRange", dialogState.inputXRange || "", "data-analysis")}
            ${dataRangeFieldHtml("第 1 変数の入力範囲", "variable1Range", dialogState.variable1Range || "", "data-analysis")}
            ${dataRangeFieldHtml("第 2 変数の入力範囲", "variable2Range", dialogState.variable2Range || "", "data-analysis")}
            ${dataRangeFieldHtml("ビン範囲", "binRange", dataAnalysisDialogStateValue(dialogState, "binRange", ""), "data-analysis")}
            <fieldset class="data-dialog-fieldset data-analysis-output-options">
              <legend>出力オプション</legend>
              <label class="data-dialog-check"><input type="radio" name="dataAnalysisOutputMode" data-data-field="outputMode" value="range" data-data-analysis-sync${outputMode === "range" ? " checked" : ""} />出力先</label>
              ${dataCellFieldHtml("出力先", "outputCell", dialogState.outputCell, "data-analysis")}
              <label class="data-dialog-check"><input type="radio" name="dataAnalysisOutputMode" data-data-field="outputMode" value="newWorksheet" data-data-analysis-sync${outputMode === "newWorksheet" ? " checked" : ""} />新規ワークシート</label>
              <label class="excel-dialog-field data-analysis-new-sheet-name"><span class="excel-dialog-caption">シート名</span><input type="text" data-data-field="outputSheetName" data-data-analysis-sync value="${escapeAttr(dataAnalysisDialogStateValue(dialogState, "outputSheetName", ""))}" /></label>
              <label class="data-dialog-check"><input type="radio" name="dataAnalysisOutputMode" data-data-field="outputMode" value="newWorkbook" data-data-analysis-sync${outputMode === "newWorkbook" ? " checked" : ""} />新規ブック</label>
            </fieldset>
            <fieldset class="data-dialog-fieldset data-analysis-grouped-options">
              <legend>グループ化</legend>
              <label class="data-dialog-check data-analysis-grouped-choice"><input type="radio" name="dataAnalysisGroupedBy" data-data-field="groupedBy" value="columns" data-data-analysis-sync${groupedBy === "columns" ? " checked" : ""} />列</label>
              <label class="data-dialog-check data-analysis-grouped-choice"><input type="radio" name="dataAnalysisGroupedBy" data-data-field="groupedBy" value="rows" data-data-analysis-sync${groupedBy === "rows" ? " checked" : ""} />行</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="labels" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "labels", false)} />先頭行/左端列をラベルとして使用</label>
            </fieldset>
            <fieldset class="data-dialog-fieldset data-analysis-sampling-options">
              <legend>サンプリング方法</legend>
              <label class="data-dialog-check"><input type="radio" name="dataAnalysisSamplingMethod" data-data-field="samplingMethod" value="periodic" data-data-analysis-sync${samplingMethod === "periodic" ? " checked" : ""} />周期</label>
              <label class="data-dialog-check"><input type="radio" name="dataAnalysisSamplingMethod" data-data-field="samplingMethod" value="random" data-data-analysis-sync${samplingMethod === "random" ? " checked" : ""} />ランダム</label>
            </fieldset>
            <div class="data-analysis-options-grid">
              <label class="excel-dialog-field"><span class="excel-dialog-caption">信頼度</span><input type="number" min="1" max="99.9" step="0.1" data-data-field="confidence" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "confidence", 95)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">有意水準 α</span><input type="number" min="0.001" max="0.5" step="0.001" data-data-field="alpha" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "alpha", 0.05)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">間隔 / 期間</span><input type="number" min="2" max="999" step="1" data-data-field="interval" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "interval", 3)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">サンプル数</span><input type="number" min="1" max="100000" step="1" data-data-field="sampleCount" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "sampleCount", 10)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">乱数の行数</span><input type="number" min="1" max="10000" step="1" data-data-field="randomRows" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "randomRows", 10)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">乱数の列数</span><input type="number" min="1" max="64" step="1" data-data-field="randomCols" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "randomCols", 2)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">平均 / p / λ</span><input type="number" step="any" data-data-field="mean" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "mean", 0)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">標準偏差 / 試行回数</span><input type="number" min="0.000001" step="any" data-data-field="stddev" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "stddev", 1)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">下限</span><input type="number" step="any" data-data-field="lowerBound" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "lowerBound", 0)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">上限</span><input type="number" step="any" data-data-field="upperBound" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "upperBound", 1)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">ステップ</span><input type="number" min="0.000001" step="any" data-data-field="patternStep" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "patternStep", 1)}" /></label>
              ${dataRangeFieldHtml("値と確率の入力範囲", "probabilityRange", dataAnalysisDialogStateValue(dialogState, "probabilityRange", ""), "data-analysis")}
              <label class="excel-dialog-field"><span class="excel-dialog-caption">分布</span><select data-data-field="randomDistribution" data-data-analysis-sync><option value="normal"${dataAnalysisSelectedAttr(dialogState, "randomDistribution", "normal", "normal")}>正規</option><option value="uniform"${dataAnalysisSelectedAttr(dialogState, "randomDistribution", "uniform", "normal")}>一様</option><option value="bernoulli"${dataAnalysisSelectedAttr(dialogState, "randomDistribution", "bernoulli", "normal")}>ベルヌーイ</option><option value="binomial"${dataAnalysisSelectedAttr(dialogState, "randomDistribution", "binomial", "normal")}>二項</option><option value="poisson"${dataAnalysisSelectedAttr(dialogState, "randomDistribution", "poisson", "normal")}>ポアソン</option><option value="pattern"${dataAnalysisSelectedAttr(dialogState, "randomDistribution", "pattern", "normal")}>パターン</option><option value="discrete"${dataAnalysisSelectedAttr(dialogState, "randomDistribution", "discrete", "normal")}>離散</option></select></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">乱数シード</span><input type="number" step="1" data-data-field="randomSeed" data-data-analysis-sync value="${escapeAttr(dataAnalysisDialogStateValue(dialogState, "randomSeed", ""))}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">仮説平均との差異</span><input type="number" step="any" data-data-field="hypothesizedMeanDiff" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "hypothesizedMeanDiff", 0)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">第1標本の分散</span><input type="number" min="0.000001" step="any" data-data-field="knownVariance1" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "knownVariance1", 1)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">第2標本の分散</span><input type="number" min="0.000001" step="any" data-data-field="knownVariance2" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "knownVariance2", 1)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">K 番目の最大値</span><input type="number" min="1" step="1" data-data-field="kthLargest" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "kthLargest", 1)}" /></label>
              <label class="excel-dialog-field"><span class="excel-dialog-caption">K 番目の最小値</span><input type="number" min="1" step="1" data-data-field="kthSmallest" data-data-analysis-sync value="${dataAnalysisNumberValue(dialogState, "kthSmallest", 1)}" /></label>
            </div>
            <fieldset class="data-dialog-fieldset data-analysis-extra-options">
              <legend>オプション</legend>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="summaryStats" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "summaryStats", true)} />要約統計量</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="confidenceLevelEnabled" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "confidenceLevelEnabled", true)} />平均の信頼区間</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="kthLargestEnabled" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "kthLargestEnabled", false)} />K 番目の最大値</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="kthSmallestEnabled" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "kthSmallestEnabled", false)} />K 番目の最小値</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="chartOutput" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "chartOutput", false)} />グラフ作成</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="cumulativePercent" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "cumulativePercent", true)} />累積パーセント</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="pareto" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "pareto", false)} />パレート順</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="standardErrorOutput" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "standardErrorOutput", false)} />標準誤差</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="constantZero" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "constantZero", false)} />定数に 0 を使用</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="residuals" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "residuals", false)} />残差</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="standardizedResiduals" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "standardizedResiduals", false)} />標準化残差</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="residualPlot" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "residualPlot", false)} />残差グラフ</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="lineFitPlot" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "lineFitPlot", false)} />観測値グラフ</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="normalProbabilityPlot" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "normalProbabilityPlot", false)} />正規確率グラフ</label>
              <label class="data-dialog-check"><input type="checkbox" data-data-field="inverseFourier" data-data-analysis-sync${dataAnalysisCheckedAttr(dialogState, "inverseFourier", false)} />逆変換</label>
            </fieldset>
            <div class="data-dialog-summary data-analysis-summary" data-data-analysis-summary></div>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="data-analysis-options-back">戻る(B)</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="data-analysis-help">ヘルプ(H)</button>
      </div>
    </div>`;
}



function handleDataAnalysisDialogSync(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='data-analysis']");
  syncDataAnalysisDialog(dialog);
}



function syncDataAnalysisDialog(dialog) {
  if (!dialog) return;
  const options = dataAnalysisDialogOptions(dialog, { preview: true });
  state.dataDialog = {
    ...(state.dataDialog || {}),
    ...dataAnalysisDialogStateFromOptions(options),
  };
  const title = dialog.querySelector("[data-data-analysis-method-title]");
  if (title) title.textContent = dataAnalysisMethodLabel(options.method);
  const noInput = options.method === "random";
  const regressionInput = options.method === "regression";
  const twoVariableInput = DATA_ANALYSIS_TWO_VARIABLE_METHODS.has(options.method);
  const genericInputHidden = noInput || regressionInput || twoVariableInput;
  const inputRangeField = dialog.querySelector("[data-data-field='inputRange']")?.closest(".excel-dialog-field");
  inputRangeField?.classList.toggle("hidden", genericInputHidden);
  dialog.querySelector("[data-data-field='inputRange']")?.toggleAttribute("disabled", genericInputHidden);
  dialog.querySelector("[data-data-range-target='inputRange']")?.toggleAttribute("disabled", genericInputHidden);
  ["inputYRange", "inputXRange"].forEach((field) => {
    const control = dialog.querySelector(`[data-data-field='${field}']`);
    const hidden = !regressionInput;
    control?.closest(".excel-dialog-field")?.classList.toggle("hidden", hidden);
    control?.toggleAttribute("disabled", hidden);
    dialog.querySelector(`[data-data-range-target='${field}']`)?.toggleAttribute("disabled", hidden);
  });
  ["variable1Range", "variable2Range"].forEach((field) => {
    const control = dialog.querySelector(`[data-data-field='${field}']`);
    const hidden = !twoVariableInput;
    control?.closest(".excel-dialog-field")?.classList.toggle("hidden", hidden);
    control?.toggleAttribute("disabled", hidden);
    dialog.querySelector(`[data-data-range-target='${field}']`)?.toggleAttribute("disabled", hidden);
  });
  dialog.querySelector("[data-data-field='binRange']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "histogram");
  const outputMode = options.outputMode || "range";
  const outputCellField = dialog.querySelector("[data-data-field='outputCell']")?.closest(".excel-dialog-field");
  const outputCellControl = dialog.querySelector("[data-data-field='outputCell']");
  const outputCellButton = dialog.querySelector("[data-data-range-target='outputCell']");
  outputCellField?.classList.toggle("is-disabled", outputMode !== "range");
  if (outputCellControl) outputCellControl.disabled = outputMode !== "range";
  if (outputCellButton) outputCellButton.disabled = outputMode !== "range";
  const newSheetNameField = dialog.querySelector("[data-data-field='outputSheetName']")?.closest(".excel-dialog-field");
  const newSheetNameControl = dialog.querySelector("[data-data-field='outputSheetName']");
  newSheetNameField?.classList.toggle("hidden", outputMode === "range");
  if (newSheetNameControl) newSheetNameControl.disabled = outputMode === "range";
  dialog.querySelectorAll("[name='dataAnalysisGroupedBy'], [data-data-field='labels']").forEach((control) => {
    control.disabled = noInput;
  });
  const groupedOptions = dialog.querySelector(".data-analysis-grouped-options");
  groupedOptions?.classList.toggle("hidden", noInput);
  groupedOptions?.querySelectorAll(".data-analysis-grouped-choice").forEach((choice) => {
    choice.classList.toggle("hidden", regressionInput || twoVariableInput);
  });
  dialog.querySelector("[data-data-field='confidence']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "descriptive" && options.method !== "regression");
  dialog.querySelector("[data-data-field='alpha']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", !["exponentialSmoothing", "anova1", "anova2NoReplication", "anova2Replication", "fTest", "tTestPaired", "tTestEqual", "tTestUnequal", "zTest"].includes(options.method));
  dialog.querySelector("[data-data-field='interval']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", !["movingAverage", "sampling", "anova2Replication"].includes(options.method));
  dialog.querySelector("[data-data-field='sampleCount']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "sampling");
  dialog.querySelector(".data-analysis-sampling-options")?.classList.toggle("hidden", options.method !== "sampling");
  ["randomRows", "randomCols", "mean", "stddev", "lowerBound", "upperBound", "patternStep", "randomDistribution"].forEach((field) => {
    dialog.querySelector(`[data-data-field='${field}']`)?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "random");
  });
  const probabilityRangeField = dialog.querySelector("[data-data-field='probabilityRange']")?.closest(".excel-dialog-field");
  probabilityRangeField?.classList.toggle("hidden", options.method !== "random" || options.randomDistribution !== "discrete");
  dialog.querySelector("[data-data-field='patternStep']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "random" || options.randomDistribution !== "pattern");
  dialog.querySelector("[data-data-field='lowerBound']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "random" || !["uniform", "pattern"].includes(options.randomDistribution));
  dialog.querySelector("[data-data-field='upperBound']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "random" || !["uniform", "pattern"].includes(options.randomDistribution));
  dialog.querySelector("[data-data-field='mean']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "random" || ["uniform", "pattern", "discrete"].includes(options.randomDistribution));
  dialog.querySelector("[data-data-field='stddev']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "random" || ["uniform", "pattern", "poisson", "bernoulli", "discrete"].includes(options.randomDistribution));
  syncDataAnalysisRandomFieldCaptions(dialog, options);
  syncDataAnalysisMethodFieldCaptions(dialog, options);
  const randomSeedVisible = options.method === "random" || (options.method === "sampling" && options.samplingMethod === "random");
  dialog.querySelector("[data-data-field='randomSeed']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", !randomSeedVisible);
  dialog.querySelector("[data-data-field='hypothesizedMeanDiff']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", !["tTestPaired", "tTestEqual", "tTestUnequal", "zTest"].includes(options.method));
  dialog.querySelector("[data-data-field='knownVariance1']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "zTest");
  dialog.querySelector("[data-data-field='knownVariance2']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "zTest");
  dialog.querySelector("[data-data-field='kthLargest']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "descriptive");
  dialog.querySelector("[data-data-field='kthSmallest']")?.closest(".excel-dialog-field")?.classList.toggle("hidden", options.method !== "descriptive");
  const kthLargestControl = dialog.querySelector("[data-data-field='kthLargest']");
  const kthSmallestControl = dialog.querySelector("[data-data-field='kthSmallest']");
  if (kthLargestControl) kthLargestControl.disabled = options.method !== "descriptive" || !options.kthLargestEnabled;
  if (kthSmallestControl) kthSmallestControl.disabled = options.method !== "descriptive" || !options.kthSmallestEnabled;
  const toggleCheck = (field, visible) => {
    dialog.querySelector(`[data-data-field='${field}']`)?.closest(".data-dialog-check")?.classList.toggle("hidden", !visible);
  };
  toggleCheck("summaryStats", options.method === "descriptive");
  toggleCheck("confidenceLevelEnabled", options.method === "descriptive");
  toggleCheck("kthLargestEnabled", options.method === "descriptive");
  toggleCheck("kthSmallestEnabled", options.method === "descriptive");
  toggleCheck("chartOutput", ["histogram", "movingAverage", "exponentialSmoothing", "regression"].includes(options.method));
  toggleCheck("cumulativePercent", options.method === "histogram");
  toggleCheck("pareto", options.method === "histogram");
  toggleCheck("standardErrorOutput", ["movingAverage", "exponentialSmoothing"].includes(options.method));
  toggleCheck("constantZero", options.method === "regression");
  toggleCheck("residuals", options.method === "regression");
  toggleCheck("standardizedResiduals", options.method === "regression");
  toggleCheck("residualPlot", options.method === "regression");
  toggleCheck("lineFitPlot", options.method === "regression");
  toggleCheck("normalProbabilityPlot", options.method === "regression");
  toggleCheck("inverseFourier", options.method === "fourier");
  const summary = dialog.querySelector("[data-data-analysis-summary]");
  const preview = dataAnalysisPreview(options);
  if (summary) {
    summary.textContent = preview.message;
    summary.classList.toggle("is-warning", !preview.valid);
  }
  const ok = dialog.querySelector("[data-data-dialog-action='ok']");
  if (ok) ok.disabled = !preview.valid;
}



function syncDataAnalysisMethodFieldCaptions(dialog, options) {
  if (!dialog) return;
  const setCaption = (field, text) => {
    const caption = dialog.querySelector(`[data-data-field='${field}']`)?.closest(".excel-dialog-field")?.querySelector(".excel-dialog-caption");
    if (caption) caption.textContent = text;
  };
  setCaption("alpha", options.method === "exponentialSmoothing" ? "減衰係数" : "有意水準 α");
  if (options.method === "anova2Replication") {
    setCaption("interval", "標本あたりの行数");
  } else if (options.method === "movingAverage") {
    setCaption("interval", "区間");
  } else if (options.method === "sampling") {
    setCaption("interval", "間隔");
  } else {
    setCaption("interval", "間隔 / 期間");
  }
  setCaption("sampleCount", options.samplingMethod === "random" ? "サンプル数" : "サンプル数");
  setCaption("knownVariance1", "変数 1 の分散 (既知)");
  setCaption("knownVariance2", "変数 2 の分散 (既知)");
  setCaption("hypothesizedMeanDiff", "仮説平均との差異");
}



function syncDataAnalysisRandomFieldCaptions(dialog, options) {
  if (!dialog) return;
  const setCaption = (field, text) => {
    const caption = dialog.querySelector(`[data-data-field='${field}']`)?.closest(".excel-dialog-field")?.querySelector(".excel-dialog-caption");
    if (caption) caption.textContent = text;
  };
  if (options.method !== "random") {
    setCaption("mean", "平均 / p / λ");
    setCaption("stddev", "標準偏差 / 試行回数");
    setCaption("lowerBound", "下限");
    setCaption("upperBound", "上限");
    setCaption("patternStep", "ステップ");
    return;
  }
  switch (options.randomDistribution) {
    case "bernoulli":
      setCaption("mean", "成功率 p");
      setCaption("stddev", "標準偏差 / 試行回数");
      break;
    case "binomial":
      setCaption("mean", "成功率 p");
      setCaption("stddev", "試行回数");
      break;
    case "poisson":
      setCaption("mean", "λ");
      setCaption("stddev", "標準偏差 / 試行回数");
      break;
    case "uniform":
      setCaption("lowerBound", "下限");
      setCaption("upperBound", "上限");
      setCaption("mean", "平均 / p / λ");
      setCaption("stddev", "標準偏差 / 試行回数");
      break;
    case "pattern":
      setCaption("lowerBound", "開始値");
      setCaption("upperBound", "終了値");
      setCaption("patternStep", "ステップ");
      setCaption("mean", "平均 / p / λ");
      setCaption("stddev", "標準偏差 / 試行回数");
      break;
    case "normal":
    default:
      setCaption("mean", "平均");
      setCaption("stddev", "標準偏差");
      setCaption("lowerBound", "下限");
      setCaption("upperBound", "上限");
      setCaption("patternStep", "ステップ");
      break;
  }
}



function dataAnalysisDialogOptions(dialog, options = {}) {
  const method = dialog?.querySelector("[data-data-field='method']")?.value || state.dataDialog?.method || "descriptive";
  const inputText = dialog?.querySelector("[data-data-field='inputRange']")?.value || state.dataDialog?.inputRange || "";
  const inputYText = dialog?.querySelector("[data-data-field='inputYRange']")?.value || state.dataDialog?.inputYRange || "";
  const inputXText = dialog?.querySelector("[data-data-field='inputXRange']")?.value || state.dataDialog?.inputXRange || "";
  const variable1Text = dialog?.querySelector("[data-data-field='variable1Range']")?.value || state.dataDialog?.variable1Range || "";
  const variable2Text = dialog?.querySelector("[data-data-field='variable2Range']")?.value || state.dataDialog?.variable2Range || "";
  const outputText = dialog?.querySelector("[data-data-field='outputCell']")?.value || state.dataDialog?.outputCell || "";
  const outputMode = dialog?.querySelector("[name='dataAnalysisOutputMode']:checked")?.value || state.dataDialog?.outputMode || "range";
  const outputSheetName = dialog?.querySelector("[data-data-field='outputSheetName']")?.value || state.dataDialog?.outputSheetName || "";
  const binText = dialog?.querySelector("[data-data-field='binRange']")?.value || state.dataDialog?.binRange || "";
  const groupedBy = dialog?.querySelector("[name='dataAnalysisGroupedBy']:checked")?.value || state.dataDialog?.groupedBy || "columns";
  const numberField = (field, fallback) => {
    const control = dialog?.querySelector(`[data-data-field='${field}']`);
    const value = Number(control?.value ?? state.dataDialog?.[field]);
    return Number.isFinite(value) ? value : fallback;
  };
  const rawAlpha = numberField("alpha", 0.05);
  return {
    preview: Boolean(options.preview),
    method,
    inputText,
    inputYText,
    inputXText,
    variable1Text,
    variable2Text,
    outputText,
    outputMode,
    outputSheetName,
    binText,
    groupedBy,
    labels: dialog?.querySelector("[data-data-field='labels']")?.checked ?? Boolean(state.dataDialog?.labels),
    confidence: clampNumber(numberField("confidence", 95), 1, 99.9, 95),
    alpha: method === "exponentialSmoothing"
      ? clampNumber(rawAlpha, 0.001, 0.999, 0.05)
      : clampNumber(rawAlpha, 0.001, 0.5, 0.05),
    interval: Math.max(2, Math.round(numberField("interval", 3))),
    sampleCount: Math.max(1, Math.round(numberField("sampleCount", 10))),
    samplingMethod: (dialog?.querySelector("[name='dataAnalysisSamplingMethod']:checked")?.value || state.dataDialog?.samplingMethod) === "random" ? "random" : "periodic",
    randomRows: Math.max(1, Math.round(numberField("randomRows", 10))),
    randomCols: Math.max(1, Math.round(numberField("randomCols", 2))),
    mean: numberField("mean", 0),
    stddev: Math.max(0.000001, numberField("stddev", 1)),
    lowerBound: numberField("lowerBound", 0),
    upperBound: numberField("upperBound", 1),
    patternStep: Math.max(0.000001, Math.abs(numberField("patternStep", 1))),
    probabilityRangeText: dialog?.querySelector("[data-data-field='probabilityRange']")?.value || state.dataDialog?.probabilityRange || "",
    randomDistribution: dialog?.querySelector("[data-data-field='randomDistribution']")?.value || state.dataDialog?.randomDistribution || "normal",
    randomSeed: dialog?.querySelector("[data-data-field='randomSeed']")?.value || state.dataDialog?.randomSeed || "",
    hypothesizedMeanDiff: numberField("hypothesizedMeanDiff", 0),
    knownVariance1: Math.max(0.000001, numberField("knownVariance1", 1)),
    knownVariance2: Math.max(0.000001, numberField("knownVariance2", 1)),
    kthLargest: Math.max(1, Math.round(numberField("kthLargest", 1))),
    kthSmallest: Math.max(1, Math.round(numberField("kthSmallest", 1))),
    summaryStats: dialog?.querySelector("[data-data-field='summaryStats']")?.checked ?? state.dataDialog?.summaryStats ?? true,
    confidenceLevelEnabled: dialog?.querySelector("[data-data-field='confidenceLevelEnabled']")?.checked ?? state.dataDialog?.confidenceLevelEnabled ?? true,
    kthLargestEnabled: dialog?.querySelector("[data-data-field='kthLargestEnabled']")?.checked ?? Boolean(state.dataDialog?.kthLargestEnabled),
    kthSmallestEnabled: dialog?.querySelector("[data-data-field='kthSmallestEnabled']")?.checked ?? Boolean(state.dataDialog?.kthSmallestEnabled),
    chartOutput: dialog?.querySelector("[data-data-field='chartOutput']")?.checked ?? Boolean(state.dataDialog?.chartOutput),
    cumulativePercent: dialog?.querySelector("[data-data-field='cumulativePercent']")?.checked ?? state.dataDialog?.cumulativePercent ?? true,
    pareto: dialog?.querySelector("[data-data-field='pareto']")?.checked ?? Boolean(state.dataDialog?.pareto),
    standardErrorOutput: dialog?.querySelector("[data-data-field='standardErrorOutput']")?.checked ?? Boolean(state.dataDialog?.standardErrorOutput),
    constantZero: dialog?.querySelector("[data-data-field='constantZero']")?.checked ?? Boolean(state.dataDialog?.constantZero),
    residuals: dialog?.querySelector("[data-data-field='residuals']")?.checked ?? Boolean(state.dataDialog?.residuals),
    standardizedResiduals: dialog?.querySelector("[data-data-field='standardizedResiduals']")?.checked ?? Boolean(state.dataDialog?.standardizedResiduals),
    residualPlot: dialog?.querySelector("[data-data-field='residualPlot']")?.checked ?? Boolean(state.dataDialog?.residualPlot),
    lineFitPlot: dialog?.querySelector("[data-data-field='lineFitPlot']")?.checked ?? Boolean(state.dataDialog?.lineFitPlot),
    normalProbabilityPlot: dialog?.querySelector("[data-data-field='normalProbabilityPlot']")?.checked ?? Boolean(state.dataDialog?.normalProbabilityPlot),
    inverseFourier: dialog?.querySelector("[data-data-field='inverseFourier']")?.checked ?? Boolean(state.dataDialog?.inverseFourier),
  };
}



function dataAnalysisPreview(options) {
  const methodLabel = dataAnalysisMethodLabel(options.method);
  const outputValidation = dataAnalysisOutputValidation(options);
  if (!outputValidation.valid) return outputValidation;
  const finalize = (summary) => dataAnalysisPreviewWithOutputRange(options, summary);
  const outputDescription = dataAnalysisOutputDescription(options);
  if (options.method === "random") {
    if (options.randomDistribution === "discrete" && !decodeWorkbookRangeOrName(options.probabilityRangeText, activeSheet())) {
      return { valid: false, message: "離散分布の値と確率の入力範囲を指定してください。" };
    }
    const cells = Math.max(1, options.randomRows) * Math.max(1, options.randomCols);
    return finalize({ valid: true, message: `${methodLabel}: ${options.randomRows} 行 x ${options.randomCols} 列、${cells.toLocaleString("ja-JP")} 個の乱数を出力します。${outputDescription}` });
  }
  if (options.method === "regression") {
    const regression = dataAnalysisRegressionSeriesFromOptions(options);
    if (!regression.valid) return regression;
    const xCount = Math.max(0, regression.series.length - 1);
    const valueCount = regression.series.reduce((total, item) => total + item.values.length, 0);
    return finalize({ valid: true, message: `${methodLabel}: Y 系列 1 個、X 系列 ${xCount} 個、${valueCount.toLocaleString("ja-JP")} 個の数値を分析します。${outputDescription}` });
  }
  if (DATA_ANALYSIS_TWO_VARIABLE_METHODS.has(options.method)) {
    const pair = dataAnalysisTwoVariableSeriesFromOptions(options);
    if (!pair.valid) return pair;
    const valueCount = pair.series.reduce((total, item) => total + item.values.length, 0);
    return finalize({ valid: true, message: `${methodLabel}: 2 系列、${valueCount.toLocaleString("ja-JP")} 個の数値を分析します。${outputDescription}` });
  }
  const inputRange = decodeWorkbookRangeOrName(options.inputText, activeSheet());
  if (!inputRange) return { valid: false, message: "入力範囲を正しく入力してください。" };
  const sheet = state.model?.sheets?.[inputRange.sheetIndex ?? state.activeSheetIndex];
  if (!sheet) return { valid: false, message: "入力範囲のシートが見つかりません。" };
  if (options.method === "anova2NoReplication" || options.method === "anova2Replication") {
    const matrix = dataAnalysisInputMatrix(sheet, inputRange, options);
    const rows = matrix.values.length;
    const cols = matrix.values[0]?.length || 0;
    const replicationText = options.method === "anova2Replication" ? `、繰り返し数 ${options.interval}` : "";
    if (rows < 2 || cols < 2) return { valid: false, message: `${methodLabel}には 2 行 x 2 列以上の数値表が必要です。` };
    return finalize({ valid: true, message: `${methodLabel}: ${rows} 行 x ${cols} 列の数値表を分析します${replicationText}。${outputDescription}` });
  }
  const series = dataAnalysisSeries(sheet, inputRange, options);
  const minSeries = ["correlation", "covariance", "regression", "anova1", "fTest", "tTestPaired", "tTestEqual", "tTestUnequal", "zTest"].includes(options.method) ? 2 : 1;
  if (series.length < minSeries) return { valid: false, message: `${methodLabel}には数値系列が ${minSeries} 個以上必要です。` };
  const valueCount = series.reduce((total, item) => total + item.values.length, 0);
  if (!valueCount) return { valid: false, message: "入力範囲に数値データがありません。" };
  return finalize({ valid: true, message: `${methodLabel}: ${series.length} 系列、${valueCount.toLocaleString("ja-JP")} 個の数値を分析します。${outputDescription}` });
}



function dataAnalysisPreviewWithOutputRange(options, summary) {
  if (!summary?.valid || options.outputMode !== "range") return summary;
  const groupedValidation = dataAnalysisGroupedOutputPreviewValidation(options);
  if (groupedValidation && !groupedValidation.valid) return groupedValidation;
  if (options.method === "random") {
    const rowCount = Math.max(1, Math.round(options.randomRows || 1)) + 1;
    const colCount = Math.max(1, Math.round(options.randomCols || 1));
    const outputValidation = dataAnalysisValidateOutputSize(options, rowCount, colCount);
    return outputValidation.valid ? summary : outputValidation;
  }
  const result = dataAnalysisRows(options);
  if (!result.valid) return result;
  const outputValidation = dataAnalysisValidateOutputRange(options, result.rows);
  return outputValidation.valid ? summary : outputValidation;
}



function dataAnalysisGroupedOutputPreviewValidation(options) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || options.outputMode !== "range") return null;
  const randomSize = options.method === "random"
    ? {
        rowCount: Math.max(1, Math.round(options.randomRows || 1)) + 1,
        colCount: Math.max(1, Math.round(options.randomCols || 1)),
      }
    : null;
  for (const sheetIndex of sheetIndexes) {
    const sheetName = state.model?.sheets?.[sheetIndex]?.name || "Sheet";
    const targetOptions = rebaseDataAnalysisOptionsForSheet(options, sheetIndex);
    const outputValidation = randomSize
      ? dataAnalysisValidateOutputSize(targetOptions, randomSize.rowCount, randomSize.colCount, sheetIndex)
      : (() => {
          const result = dataAnalysisRows(targetOptions);
          if (!result.valid) return result;
          return dataAnalysisValidateOutputRange(targetOptions, result.rows, sheetIndex);
        })();
    if (!outputValidation.valid) {
      return { valid: false, message: `${sheetName}：${outputValidation.message}` };
    }
  }
  return { valid: true };
}



function dataAnalysisOutputValidation(options = {}, fallbackSheetIndex = state.activeSheetIndex) {
  if (options.outputMode === "newWorksheet" || options.outputMode === "newWorkbook") {
    const sheetName = dataAnalysisNewOutputSheetName(options);
    if (!sheetName.valid) return { valid: false, message: sheetName.message };
    return { valid: true };
  }
  const outputCell = decodeWorkbookCell(options.outputText);
  if (!outputCell) return { valid: false, message: "出力先セルを正しく入力してください。" };
  const outputSheetIndex = outputCell.sheetIndex ?? fallbackSheetIndex;
  const outputSheet = state.model?.sheets?.[outputSheetIndex];
  if (!outputSheet) return { valid: false, message: "出力先シートが見つかりません。" };
  const outputPointRange = normalizeRange(
    { sheetIndex: outputSheetIndex, row: outputCell.row, col: outputCell.col },
    { sheetIndex: outputSheetIndex, row: outputCell.row, col: outputCell.col },
  );
  if (isSheetProtected(outputSheet) && rangeHasLockedCellsForProtection(outputSheet, outputPointRange)) {
    return { valid: false, message: "保護されたシート上のロックされたセルには出力できません。" };
  }
  return { valid: true };
}



function dataAnalysisRowsOutputSize(rows = []) {
  const cleanRows = (rows || []).filter((row) => row && row.length);
  const width = cleanRows.length ? Math.max(1, ...cleanRows.map((row) => row.length)) : 0;
  return { rows: cleanRows, rowCount: cleanRows.length, colCount: width };
}



function dataAnalysisOutputRangeFromSize(options = {}, rowCount = 0, colCount = 0, rows = null, fallbackSheetIndex = state.activeSheetIndex) {
  const base = dataAnalysisOutputValidation(options, fallbackSheetIndex);
  if (!base.valid) return base;
  if (options.outputMode === "newWorksheet" || options.outputMode === "newWorkbook") {
    return { valid: true, rows: rows || [], rowCount, colCount };
  }
  if (rowCount < 1 || colCount < 1) return { valid: false, message: "出力する分析結果がありません。" };
  const outputCell = decodeWorkbookCell(options.outputText);
  if (!outputCell) return { valid: false, message: "出力先セルを正しく入力してください。" };
  const outputSheetIndex = outputCell.sheetIndex ?? fallbackSheetIndex;
  const outputSheet = state.model?.sheets?.[outputSheetIndex];
  if (!outputSheet) return { valid: false, message: "出力先シートが見つかりません。" };
  const targetRange = normalizeRange(
    { sheetIndex: outputSheetIndex, row: outputCell.row, col: outputCell.col },
    { sheetIndex: outputSheetIndex, row: outputCell.row + rowCount - 1, col: outputCell.col + colCount - 1 },
  );
  return { valid: true, outputCell, outputSheetIndex, outputSheet, targetRange, rows: rows || [], rowCount, colCount };
}



function dataAnalysisOutputRangeFromRows(options = {}, resultRows = [], fallbackSheetIndex = state.activeSheetIndex) {
  const size = dataAnalysisRowsOutputSize(resultRows);
  return dataAnalysisOutputRangeFromSize(options, size.rowCount, size.colCount, size.rows, fallbackSheetIndex);
}



function dataAnalysisValidateOutputPlacement(placement) {
  if (!placement?.valid || !placement.targetRange) return placement;
  if (rangeHasLockedCellsForProtection(placement.outputSheet, placement.targetRange)) {
    return { valid: false, message: "保護されたシート上のロックされたセルには出力できません。" };
  }
  if (mergesIntersectingRange(placement.outputSheet, placement.targetRange).length) {
    return { valid: false, message: "結合されたセルを含む範囲にはデータ分析の結果を出力できません。" };
  }
  return placement;
}



function dataAnalysisValidateOutputSize(options = {}, rowCount = 0, colCount = 0, fallbackSheetIndex = state.activeSheetIndex) {
  return dataAnalysisValidateOutputPlacement(dataAnalysisOutputRangeFromSize(options, rowCount, colCount, null, fallbackSheetIndex));
}



function dataAnalysisValidateOutputRange(options = {}, resultRows = [], fallbackSheetIndex = state.activeSheetIndex) {
  return dataAnalysisValidateOutputPlacement(dataAnalysisOutputRangeFromRows(options, resultRows, fallbackSheetIndex));
}



function dataAnalysisMethodLabel(methodValue) {
  return DATA_ANALYSIS_METHODS.find((method) => method.value === methodValue)?.label || "データ分析";
}



function showDataAnalysisHelp(methodValue = state.dataDialog?.method || "descriptive") {
  const method = DATA_ANALYSIS_METHODS.some((item) => item.value === methodValue) ? methodValue : "descriptive";
  const label = dataAnalysisMethodLabel(method);
  const body = DATA_ANALYSIS_HELP[method] || "選択した分析ツールの入力範囲、出力先、オプションを指定して分析結果を作成します。";
  showMessageBox(`${label}\n\n${body}`, { title: "データ分析 ヘルプ", tone: "info" });
}



function dataAnalysisOutputDescription(options) {
  if (options.outputMode === "newWorksheet") {
    return `出力先: 新規ワークシート${options.outputSheetName ? `「${options.outputSheetName}」` : ""}`;
  }
  if (options.outputMode === "newWorkbook") {
    return "出力先: 新規ブック";
  }
  return `出力先: ${options.outputText || "未指定"}`;
}



function applyDataAnalysisDialog(dialog) {
  const options = dataAnalysisDialogOptions(dialog);
  const result = dataAnalysisRows(options);
  if (!result.valid) {
    showMessageBox(result.message, { title: "Cht WebSheet" });
    syncDataAnalysisDialog(dialog);
    return;
  }
  if (options.outputMode === "newWorksheet" || options.outputMode === "newWorkbook") {
    applyDataAnalysisToNewSheet(options, result);
    return;
  }
  const grouped = applyDataAnalysisDialogToSheetGroup(options);
  if (grouped === "pending") return;
  if (grouped === true) return;
  const outputPlacement = dataAnalysisValidateOutputRange(options, result.rows);
  if (!outputPlacement.valid) {
    showMessageBox(outputPlacement.message, { title: "Cht WebSheet" });
    syncDataAnalysisDialog(dialog);
    return;
  }
  const { outputCell, outputSheet, rows, targetRange } = outputPlacement;
  if (rangeHasUserData(outputSheet, targetRange)) {
    const snapshot = cloneJsonLikeValue(options);
    showDataOverwriteConfirmation(() => {
      const restoredDialog = document.querySelector("[data-data-dialog='data-analysis']");
      if (restoredDialog) applyDataAnalysisWithOptions(snapshot, outputCell, rows, targetRange, result);
    });
    return;
  }
  applyDataAnalysisWithOptions(options, outputCell, rows, targetRange, result);
}



function applyDataAnalysisDialogToSheetGroup(options, control = {}) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || options.outputMode !== "range") return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const targetOptions = rebaseDataAnalysisOptionsForSheet(options, sheetIndex);
    const result = dataAnalysisRows(targetOptions);
    if (!result.valid) {
      showMessageBox(`${state.model?.sheets?.[sheetIndex]?.name || "Sheet"}：${result.message}`, { title: "Cht WebSheet" });
      return true;
    }
    const outputPlacement = dataAnalysisValidateOutputRange(targetOptions, result.rows, sheetIndex);
    if (!outputPlacement.valid) {
      showMessageBox(`${state.model?.sheets?.[sheetIndex]?.name || "Sheet"}：${outputPlacement.message}`, { title: "Cht WebSheet" });
      return true;
    }
    const { outputCell, outputSheetIndex, outputSheet, rows, targetRange } = outputPlacement;
    items.push({ sheetIndex: outputSheetIndex, sourceSheetIndex: sheetIndex, sheet: outputSheet, outputCell, rows, targetRange, result, options: targetOptions, ranges: [targetRange] });
  }
  if (!items.length) return true;
  const overwrite = items.find((item) => rangeHasUserData(item.sheet, item.targetRange));
  if (overwrite && !control.skipOverwriteCheck) {
    const snapshot = cloneJsonLikeValue(options);
    showDataOverwriteConfirmation(() => {
      const result = applyDataAnalysisDialogToSheetGroup(snapshot, { ...control, skipOverwriteCheck: true });
      if (result === true) closeDataDialog();
    });
    return "pending";
  }
  const affectedSheetIndexes = affectedSheetIndexesForGroupedRanges(items);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  try {
    items.forEach((item) => {
      applyDataAnalysisWithOptions(item.options, item.outputCell, item.rows, item.targetRange, item.result, {
        closeDialog: false,
        commitHistory: false,
      });
    });
    closeDataDialog();
    state.hf = buildFormulaEngine(state.model);
    const activeItem = items.find((item) => item.sourceSheetIndex === state.activeSheetIndex) || items[0];
    state.activeSheetIndex = activeItem.sheetIndex;
    state.selected = { sheetIndex: activeItem.sheetIndex, row: activeItem.targetRange.top, col: activeItem.targetRange.left };
    state.selectionRange = { ...activeItem.targetRange, sheetIndex: activeItem.sheetIndex };
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
      label: `${dataAnalysisMethodLabel(options.method)} ${items.length} シートデータ分析`,
    });
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus(`${items.length} 個のシートに${dataAnalysisMethodLabel(options.method)}を出力しました。`);
    return true;
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`データ分析を出力できませんでした：${error.message || error}`);
    return true;
  }
}



function rebaseDataAnalysisOptionsForSheet(options, sheetIndex) {
  const rebaseRange = (text) => rebaseWorkbookRangeTextForSheet(text, sheetIndex);
  return {
    ...options,
    inputText: rebaseRange(options.inputText),
    inputYText: rebaseRange(options.inputYText),
    inputXText: rebaseRange(options.inputXText),
    variable1Text: rebaseRange(options.variable1Text),
    variable2Text: rebaseRange(options.variable2Text),
    binText: rebaseRange(options.binText),
    probabilityRangeText: rebaseRange(options.probabilityRangeText),
    outputText: rebaseWorkbookCellTextForSheet(options.outputText, sheetIndex),
  };
}



function rebaseWorkbookRangeTextForSheet(text, sheetIndex) {
  const source = String(text || "").trim();
  if (!source) return source;
  const range = decodeWorkbookRangeOrName(source, activeSheet());
  if (!range || (range.sheetIndex ?? state.activeSheetIndex) !== state.activeSheetIndex) return source;
  return workbookRangeReference(sheetIndex, range);
}



function rebaseWorkbookCellTextForSheet(text, sheetIndex) {
  const source = String(text || "").trim();
  if (!source) return source;
  const cell = decodeWorkbookCell(source);
  if (!cell || (cell.sheetIndex ?? state.activeSheetIndex) !== state.activeSheetIndex) return source;
  return workbookCellReference(sheetIndex, cell.row, cell.col);
}



function applyDataAnalysisToNewSheet(options, result = {}) {
  const rows = result.rows.filter((row) => row && row.length);
  if (!rows.length) {
    showMessageBox("出力する分析結果がありません。", { title: "Cht WebSheet" });
    return;
  }
  const sheetNameResult = dataAnalysisNewOutputSheetName(options);
  if (!sheetNameResult.valid) {
    showMessageBox(sheetNameResult.message, { title: "Cht WebSheet" });
    return;
  }
  const sheetName = sheetNameResult.name;
  const insertIndex = state.model.sheets.length;
  const activeSheetBefore = state.activeSheetIndex;
  const outputSheet = createBlankSheet(sheetName);
  state.model.sheets.push(outputSheet);
  const outputCell = { sheetIndex: insertIndex, row: 1, col: 1 };
  const width = Math.max(1, ...rows.map((row) => row.length));
  const targetRange = normalizeRange(
    { sheetIndex: insertIndex, row: 1, col: 1 },
    { sheetIndex: insertIndex, row: rows.length, col: width },
  );
  ensureSheetSize(outputSheet, targetRange.bottom, targetRange.right);
  rows.forEach((rowValues, rowIndex) => {
    for (let colIndex = 0; colIndex < width; colIndex += 1) {
      const value = rowValues[colIndex] ?? "";
      updateCellModel(outputSheet, targetRange.top + rowIndex, targetRange.left + colIndex, value);
      const cell = outputSheet.cells[cellKey(targetRange.top + rowIndex, targetRange.left + colIndex)];
      if (cell && (rowIndex === 0 || colIndex === 0)) {
        cell.css = { ...(cell.css || {}) };
        cell.css.fontWeight = "700";
      }
    }
  });
  styleDataAnalysisOutputRange(outputSheet, targetRange, rows);
  insertDataAnalysisChartImages(outputSheet, insertIndex, targetRange, result.charts || []);
  closeDataDialog();
  state.activeSheetIndex = insertIndex;
  state.selected = { sheetIndex: insertIndex, row: outputCell.row, col: outputCell.col };
  state.selectionRange = { ...targetRange };
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.hf = buildFormulaEngine(state.model);
  renderWorkbook();
  updateFormulaBarForSelection();
  const label = dataAnalysisMethodLabel(options.method);
  pushSheetInsertHistory(insertIndex, outputSheet, activeSheetBefore, insertIndex, `${sheetName} ${label}`, {
    selectionAfter: state.selectionRange,
  });
  setStatus(`${sheetName} に${label}を出力しました。`);
}



function dataAnalysisNewOutputSheetName(options) {
  const requested = String(options.outputSheetName || "").trim();
  if (requested) return validateSheetName(requested, -1);
  const base = (options.outputMode === "newWorkbook" ? "新規ブック" : dataAnalysisMethodLabel(options.method)).slice(0, 24) || "データ分析";
  return { valid: true, name: uniqueSheetNameWithBase(base, state.model.sheets.map((sheet) => sheet.name)) };
}



function applyDataAnalysisWithOptions(options, outputCell, rows, targetRange, result = {}, control = {}) {
  const outputSheetIndex = outputCell.sheetIndex ?? state.activeSheetIndex;
  const outputSheet = state.model.sheets[outputSheetIndex];
  const outputPlacement = dataAnalysisValidateOutputRange(options, rows, outputSheetIndex);
  if (!outputPlacement.valid) {
    showMessageBox(outputPlacement.message, { title: "Cht WebSheet" });
    return false;
  }
  const sheetBefore = captureSheetForHistory(outputSheet);
  const sheetId = getSheetId(outputSheet.name);
  ensureSheetSize(outputSheet, targetRange.bottom, targetRange.right);
  rows.forEach((rowValues, rowIndex) => {
    for (let colIndex = 0; colIndex < targetRange.right - targetRange.left + 1; colIndex += 1) {
      const value = rowValues[colIndex] ?? "";
      writeCellContent(outputSheet, sheetId, targetRange.top + rowIndex, targetRange.left + colIndex, value);
      const cell = outputSheet.cells[cellKey(targetRange.top + rowIndex, targetRange.left + colIndex)];
      if (cell && (rowIndex === 0 || colIndex === 0)) {
        cell.css = { ...(cell.css || {}) };
        cell.css.fontWeight = "700";
      }
    }
  });
  styleDataAnalysisOutputRange(outputSheet, targetRange, rows);
  insertDataAnalysisChartImages(outputSheet, outputSheetIndex, targetRange, result.charts || []);
  if (control.closeDialog !== false) closeDataDialog();
  if (control.commitHistory === false) return true;
  const label = dataAnalysisMethodLabel(options.method);
  commitDataOperationForSheet(outputSheetIndex, outputSheet, sheetBefore, targetRange, `${label} データ分析`, `${label}を出力しました。`, targetRange);
  return true;
}



function styleDataAnalysisOutputRange(sheet, range, rows) {
  if (!sheet || !range || !Array.isArray(rows) || !rows.length) return;
  const widthByIndex = [];
  rows.forEach((row) => {
    (row || []).forEach((value, index) => {
      const text = String(value ?? "");
      const width = Math.min(220, Math.max(54, Math.round(24 + Math.min(28, text.length) * 7)));
      widthByIndex[index] = Math.max(widthByIndex[index] || 0, width);
    });
  });
  widthByIndex.forEach((width, index) => {
    const column = sheet.columns[range.left + index - 1];
    if (column) column.width = Math.max(column.hidden ? 0 : column.width || 0, width);
  });
  const sectionBorder = { width: 1, style: "solid", color: "#000000", weight: 2 };
  const lightBorder = { width: 1, style: "solid", color: "#d9e2f3", weight: 1 };
  rows.forEach((rowValues, rowIndex) => {
    const row = range.top + rowIndex;
    const nonEmptyValues = (rowValues || []).filter((value) => String(value ?? "") !== "");
    const isBlank = nonEmptyValues.length === 0;
    const isSectionTitle = nonEmptyValues.length === 1 && rowValues[0] != null;
    const isHeader = rowIndex === 0 || isSectionTitle || dataAnalysisOutputHeaderLabels(rowValues);
    for (let col = range.left; col <= range.right; col += 1) {
      const cell = sheet.cells[cellKey(row, col)];
      if (!cell || isBlank) continue;
      cell.css = { ...(cell.css || {}) };
      if (isHeader || col === range.left) cell.css.fontWeight = "700";
      const colIndex = col - range.left;
      const numFmt = dataAnalysisOutputNumberFormat(rowValues[colIndex], rows, rowIndex, colIndex);
      if (numFmt) {
        cell.numFmt = numFmt;
        if (col !== range.left) cell.css.textAlign = "right";
      }
      cell.borders = { ...(cell.borders || {}) };
      if (isHeader) {
        cell.borders.top = { ...sectionBorder };
        cell.borders.bottom = { ...sectionBorder };
      } else if (!cell.borders.bottom) {
        cell.borders.bottom = { ...lightBorder };
      }
    }
  });
}



function dataAnalysisOutputNumberFormat(value, rows, rowIndex, colIndex) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || value === "") return "";
  const title = String(rows?.[0]?.[0] ?? "");
  if (/サンプリング/.test(title)) return Number.isInteger(numeric) ? "0" : "0.000000";
  if (/乱数/.test(title)) return Number.isInteger(numeric) ? "0" : "0.000000";
  if (/順位と百分位数/.test(title)) {
    return colIndex === 3 ? "0.00%" : Number.isInteger(numeric) ? "0" : "0.000000";
  }
  if (String(rows?.[0]?.[0] ?? "") === "ビン") {
    return colIndex === 2 ? "0.00%" : Number.isInteger(numeric) ? "0" : "0.000000";
  }
  if (String(rows?.[0]?.[0] ?? "") === "k") return Number.isInteger(numeric) ? "0" : "0.000000";
  if (String(rows?.[0]?.[0] ?? "") === "概要") {
    const contextHeader = dataAnalysisOutputContextHeader(rows, rowIndex, colIndex);
    if (/標本数|データの個数|自由度/.test(contextHeader)) return "0";
    if (/百分位/.test(contextHeader)) return Number.isInteger(numeric) ? "0" : "0.000000";
    return Number.isInteger(numeric) ? "0" : "0.000000";
  }
  const header = String(rows?.[0]?.[colIndex] ?? "");
  const rowLabel = String(rows?.[rowIndex]?.[0] ?? "");
  const label = `${header} ${rowLabel}`;
  if (rows?.[0]?.[0] === "" && rows?.[0]?.length > 2) return Number.isInteger(numeric) ? "0" : "0.000000";
  if (/実績|移動平均|予測値|平滑値|標準誤差/.test(header)) return Number.isInteger(numeric) ? "0" : "0.000000";
  if (/累積\s*%|百分率/.test(label)) return "0.00%";
  if (/観測数|データの個数|標本数|自由度|標本番号|順位|期|^k$|乱数の行数|乱数の列数/.test(label)) return "0";
  return "0.000000";
}



function dataAnalysisOutputContextHeader(rows, rowIndex, colIndex) {
  for (let index = rowIndex; index >= 0; index -= 1) {
    const candidate = String(rows?.[index]?.[colIndex] ?? "").trim();
    if (!candidate) continue;
    if (/標本数|データの個数|合計|平均|分散|変動|自由度|F|P-値|P値|F 境界値|百分位/.test(candidate)) return candidate;
  }
  return "";
}



function dataAnalysisOutputHeaderLabels(rowValues) {
  const labels = (rowValues || []).map((value) => String(value ?? "").trim());
  if (!labels.some(Boolean)) return false;
  return labels.some((label) =>
    ["自由度", "変動", "分散", "標本分散", "F", "P値", "P-値", "係数", "標準誤差", "順位", "百分位", "頻度", "実数部", "虚数部", "出力", "実績", "予測値", "移動平均", "平滑値"].includes(label),
  );
}



function insertDataAnalysisChartImages(sheet, sheetIndex, targetRange, charts = []) {
  const cleanCharts = charts.filter((chart) => chart?.chart?.series?.length);
  if (!sheet || !cleanCharts.length) return;
  sheet.images = Array.isArray(sheet.images) ? sheet.images : [];
  const width = 420;
  const height = 260;
  let row = targetRange.top;
  const col = targetRange.right + 2;
  cleanCharts.forEach((chartInfo) => {
    const type = chartInfo.type || "column";
    const title = chartInfo.title || "データ分析グラフ";
    const size = fitObjectSize(width, height);
    ensureSheetSize(sheet, row, col);
    const end = endingCellForPixelSize(sheet, row, col, size.width, size.height);
    ensureSheetSize(sheet, end.row, end.col);
    const chart = {
      type,
      title,
      sourceSheet: state.model?.sheets?.[sheetIndex]?.name || sheet.name,
      sourceRange: rangeToLabel(targetRange),
      ...cloneJsonLikeValue(chartInfo.chart),
    };
    sheet.images.push({
      id: uniqueWorksheetObjectId("chart"),
      name: title,
      src: svgDataUrl(chartSvg(chart, type, { width, height, title })),
      mimeType: "image/svg+xml",
      extension: "svg",
      editAs: "oneCell",
      tl: {
        col: col - 1,
        row: row - 1,
        nativeCol: col - 1,
        nativeRow: row - 1,
        nativeColOff: 0,
        nativeRowOff: 0,
      },
      br: null,
      ext: { width: size.width, height: size.height },
      hyperlink: "",
      chart,
    });
    row = end.row + 1;
  });
}



function dataAnalysisRows(options) {
  if (options.method === "random") return dataAnalysisRandomRows(options);
  if (options.method === "regression") {
    const regression = dataAnalysisRegressionSeriesFromOptions(options);
    if (!regression.valid) return regression;
    return dataAnalysisRegressionRows(regression.series, options);
  }
  if (DATA_ANALYSIS_TWO_VARIABLE_METHODS.has(options.method)) {
    const pair = dataAnalysisTwoVariableSeriesFromOptions(options);
    if (!pair.valid) return pair;
    switch (options.method) {
      case "fTest":
        return dataAnalysisFTestRows(pair.series, options);
      case "tTestPaired":
        return dataAnalysisTTestRows(pair.series, options, "paired");
      case "tTestEqual":
        return dataAnalysisTTestRows(pair.series, options, "equal");
      case "tTestUnequal":
        return dataAnalysisTTestRows(pair.series, options, "unequal");
      case "zTest":
        return dataAnalysisZTestRows(pair.series, options);
      default:
        return pair;
    }
  }
  const inputRange = decodeWorkbookRangeOrName(options.inputText, activeSheet());
  if (!inputRange) return { valid: false, message: "入力範囲を正しく入力してください。" };
  const sheet = state.model?.sheets?.[inputRange.sheetIndex ?? state.activeSheetIndex];
  const sheetId = getSheetId(sheet?.name);
  if (!sheet || sheetId == null) return { valid: false, message: "入力範囲のシートが見つかりません。" };
  if (options.method === "fourier") {
    return dataAnalysisFourierRows(dataAnalysisFourierInputValues(sheet, inputRange, options), options);
  }
  const series = dataAnalysisSeries(sheet, inputRange, options).filter((item) => item.values.length);
  if (!series.length) return { valid: false, message: "入力範囲に数値データがありません。" };
  switch (options.method) {
    case "correlation":
      return dataAnalysisMatrixRows("相関", series, dataAnalysisCorrelation);
    case "covariance":
      return dataAnalysisMatrixRows("共分散", series, dataAnalysisCovarianceP);
    case "histogram":
      return dataAnalysisHistogramRows(sheet, options, series);
    case "movingAverage":
      return dataAnalysisMovingAverageRows(series, options);
    case "exponentialSmoothing":
      return dataAnalysisExponentialSmoothingRows(series, options);
    case "anova1":
      return dataAnalysisAnovaRows(series, options);
    case "anova2NoReplication":
      return dataAnalysisAnovaTwoFactorNoReplicationRows(sheet, inputRange, options);
    case "anova2Replication":
      return dataAnalysisAnovaTwoFactorReplicationRows(sheet, inputRange, options);
    case "rankPercentile":
      return dataAnalysisRankPercentileRows(series);
    case "sampling":
      return dataAnalysisSamplingRows(series, options);
    case "descriptive":
    default:
      return dataAnalysisDescriptiveRows(series, options);
  }
}



function dataAnalysisRegressionSeriesFromOptions(options) {
  if (!String(options.inputYText || "").trim() || !String(options.inputXText || "").trim()) {
    return dataAnalysisRegressionSeriesFromGenericInput(options);
  }
  const y = dataAnalysisSingleSeriesFromRangeText(options.inputYText, options, "Y");
  if (!y.valid) return y;
  const x = dataAnalysisSeriesFromRangeText(options.inputXText, { ...options, groupedBy: "columns" }, "X");
  if (!x.valid) return x;
  const xSeries = x.series.filter((item) => item.values.length);
  if (!y.series.values.length || !xSeries.length) {
    return { valid: false, message: "回帰分析には Y 範囲と X 範囲の数値データが必要です。" };
  }
  return { valid: true, series: [y.series, ...xSeries] };
}



function dataAnalysisRegressionSeriesFromGenericInput(options) {
  const inputRange = decodeWorkbookRangeOrName(options.inputText, activeSheet());
  if (!inputRange) return { valid: false, message: "入力 Y 範囲と入力 X 範囲を正しく入力してください。" };
  const sheet = state.model?.sheets?.[inputRange.sheetIndex ?? state.activeSheetIndex];
  if (!sheet) return { valid: false, message: "入力範囲のシートが見つかりません。" };
  const series = dataAnalysisSeries(sheet, inputRange, options).filter((item) => item.values.length);
  if (series.length < 2) return { valid: false, message: "回帰分析には Y 系列と X 系列が必要です。" };
  return { valid: true, series };
}



function dataAnalysisTwoVariableSeriesFromOptions(options) {
  if (!String(options.variable1Text || "").trim() || !String(options.variable2Text || "").trim()) {
    return dataAnalysisTwoVariableSeriesFromGenericInput(options);
  }
  const first = dataAnalysisSingleSeriesFromRangeText(options.variable1Text, options, "変数 1");
  if (!first.valid) return first;
  const second = dataAnalysisSingleSeriesFromRangeText(options.variable2Text, options, "変数 2");
  if (!second.valid) return second;
  if (!first.series.values.length || !second.series.values.length) {
    return { valid: false, message: `${dataAnalysisMethodLabel(options.method)}には 2 個の数値系列が必要です。` };
  }
  return { valid: true, series: [first.series, second.series] };
}



function dataAnalysisTwoVariableSeriesFromGenericInput(options) {
  const inputRange = decodeWorkbookRangeOrName(options.inputText, activeSheet());
  if (!inputRange) return { valid: false, message: "第 1 変数と第 2 変数の入力範囲を正しく入力してください。" };
  const sheet = state.model?.sheets?.[inputRange.sheetIndex ?? state.activeSheetIndex];
  if (!sheet) return { valid: false, message: "入力範囲のシートが見つかりません。" };
  const series = dataAnalysisSeries(sheet, inputRange, options).filter((item) => item.values.length).slice(0, 2);
  if (series.length < 2) return { valid: false, message: `${dataAnalysisMethodLabel(options.method)}には 2 個の数値系列が必要です。` };
  return { valid: true, series };
}



function dataAnalysisSeriesFromRangeText(rangeText, options, fallbackLabel = "系列") {
  const range = decodeWorkbookRangeOrName(rangeText, activeSheet());
  if (!range) return { valid: false, message: "入力範囲を正しく入力してください。" };
  const sheet = state.model?.sheets?.[range.sheetIndex ?? state.activeSheetIndex];
  if (!sheet) return { valid: false, message: "入力範囲のシートが見つかりません。" };
  const series = dataAnalysisSeries(sheet, range, options).filter((item) => item.values.length);
  if (!series.length) return { valid: false, message: `${fallbackLabel} 範囲に数値データがありません。` };
  return { valid: true, series };
}



function dataAnalysisSingleSeriesFromRangeText(rangeText, options, fallbackLabel = "系列") {
  const range = decodeWorkbookRangeOrName(rangeText, activeSheet());
  if (!range) return { valid: false, message: `${fallbackLabel} 範囲を正しく入力してください。` };
  const sheet = state.model?.sheets?.[range.sheetIndex ?? state.activeSheetIndex];
  const sheetId = getSheetId(sheet?.name);
  if (!sheet || sheetId == null) return { valid: false, message: `${fallbackLabel} 範囲のシートが見つかりません。` };
  const singleColumn = range.left === range.right;
  const singleRow = range.top === range.bottom;
  let label = fallbackLabel;
  let top = range.top;
  let left = range.left;
  if (options.labels) {
    const labelValue = getDisplayForCell(sheet, sheetId, range.top, range.left);
    if (String(labelValue || "").trim()) label = String(labelValue).trim();
    if (singleColumn || !singleRow) top += 1;
    else left += 1;
  }
  const values = [];
  const sourceIndexes = [];
  for (let row = top; row <= range.bottom; row += 1) {
    for (let col = left; col <= range.right; col += 1) {
      const value = dataAnalysisCellNumericValue(sheet, sheetId, row, col);
      if (Number.isFinite(value)) {
        values.push(value);
        sourceIndexes.push(singleColumn || !singleRow ? row : singleRow ? col : `${row}:${col}`);
      }
    }
  }
  return { valid: true, series: { label, values, sourceIndexes } };
}

export {
  addSubtotalPageBreak,
  applyDataAnalysisDialog,
  applyDataAnalysisDialogToSheetGroup,
  applyDataAnalysisToNewSheet,
  applyDataAnalysisWithOptions,
  applyDataTable,
  applyDataTableDialog,
  applyDataTableForSheetGroup,
  applyScenario,
  applySubtotalGroupOutline,
  applyWhatIfDialog,
  cellDataTableInputValue,
  clearSubtotalOutlineMetadata,
  clearSubtotalPageBreaks,
  cloneScenarios,
  closeWhatIfMenu,
  closeWhatIfMenuOnEscape,
  closeWhatIfMenuOnOutsidePointer,
  countDataTableOverwriteCells,
  createScenarioPivotSummarySheet,
  createScenarioSummarySheet,
  dataAnalysisCheckedAttr,
  dataAnalysisDefaultDialogState,
  dataAnalysisDefaultOutputCell,
  dataAnalysisDefaultVariableRange,
  dataAnalysisDefaultXRange,
  dataAnalysisDefaultYRange,
  dataAnalysisDialogHtml,
  dataAnalysisDialogOptions,
  dataAnalysisDialogStateFromOptions,
  dataAnalysisDialogStateValue,
  dataAnalysisGroupedOutputPreviewValidation,
  dataAnalysisMethodLabel,
  dataAnalysisNewOutputSheetName,
  dataAnalysisNumberValue,
  dataAnalysisOutputContextHeader,
  dataAnalysisOutputDescription,
  dataAnalysisOutputHeaderLabels,
  dataAnalysisOutputNumberFormat,
  dataAnalysisOutputRangeFromRows,
  dataAnalysisOutputRangeFromSize,
  dataAnalysisOutputValidation,
  dataAnalysisPickerDialogHtml,
  dataAnalysisPreview,
  dataAnalysisPreviewWithOutputRange,
  dataAnalysisRadioCheckedAttr,
  dataAnalysisRegressionSeriesFromGenericInput,
  dataAnalysisRegressionSeriesFromOptions,
  dataAnalysisRows,
  dataAnalysisRowsOutputSize,
  dataAnalysisSelectedAttr,
  dataAnalysisSeriesFromRangeText,
  dataAnalysisSingleSeriesFromRangeText,
  dataAnalysisTwoVariableSeriesFromGenericInput,
  dataAnalysisTwoVariableSeriesFromOptions,
  dataAnalysisValidateOutputPlacement,
  dataAnalysisValidateOutputRange,
  dataAnalysisValidateOutputSize,
  dataTableArrayFormula,
  dataTableDialogValidation,
  dataTableFormulaPoints,
  dataTableFormulaResult,
  dataTableInputFormulaReference,
  dataTableInputResultConflicts,
  dataTableInputTarget,
  dataTableResultCellCount,
  dataTableResultCells,
  dataTableResultRange,
  ensureScenarioPermission,
  formatGoalSeekNumber,
  goalSeekDialogValidation,
  goalSeekFallbackGuess,
  goalSeekInitialSecondGuess,
  goalSeekValue,
  handleDataAnalysisDialogSync,
  handleDataAnalysisPickerDoubleClick,
  handleDataAnalysisPickerSync,
  handleDataTableDialogSync,
  handleGoalSeekDialogSync,
  handleScenarioEditDialogSync,
  handleScenarioListChange,
  handleScenarioSummaryDialogSync,
  handleWhatIfMenuAction,
  insertDataAnalysisChartImages,
  insertSubtotalGrandTotalRow,
  insertSubtotalRowsForSheetGroup,
  insertSummaryRow,
  loadScenarioCurrentValues,
  openDataAnalysisDialog,
  openDataAnalysisOptionsDialog,
  openDataTableDialog,
  openScenarioEditDialog,
  openScenarioManagerDialog,
  openScenarioSummaryDialog,
  openWhatIfDialog,
  openWhatIfMenu,
  parseScenarioCells,
  parseScenarioResultCells,
  parseScenarioValues,
  persistDataAnalysisDialogState,
  rangeCellsForScenario,
  rebaseDataAnalysisOptionsForSheet,
  rebaseWorkbookCellTextForSheet,
  rebaseWorkbookRangeTextForSheet,
  removeSubtotalRowsWithoutHistory,
  restoreDataTableInputCells,
  runGoalSeek,
  runGoalSeekForSheetGroup,
  saveScenarioFromDialog,
  scenarioAffectedSheetIndexes,
  scenarioAllowsApply,
  scenarioAllowsEdit,
  scenarioCellStats,
  scenarioDescription,
  scenarioEditDialogValidation,
  scenarioManagerDialogHtml,
  scenarioSheetAllowsEdit,
  scenarioSheetIndex,
  scenarioSummaryDialogValidation,
  scenarioValuesText,
  selectedDataAnalysisPickerMethod,
  selectedScenarioFromDialog,
  setDataTableFormulaForCell,
  showDataAnalysisHelp,
  showGoalSeekResultDialog,
  splitWorkbookReferenceList,
  styleDataAnalysisOutputRange,
  styleScenarioReportSheet,
  syncDataAnalysisDialog,
  syncDataAnalysisMethodFieldCaptions,
  syncDataAnalysisPickerDialog,
  syncDataAnalysisRandomFieldCaptions,
  syncDataTableDialog,
  syncGoalSeekDialog,
  syncScenarioEditDialog,
  syncScenarioSummaryDialog,
  workbookScenarios,
};
