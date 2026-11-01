import { cwsApplyExistingShapeEdit } from './existing-shape-model.js';
// Feature module; shared bindings are imported explicitly.
import {
  $formulaInput,
  $horizontalScrollbar,
  $sheetHost,
  CWS_AGENT_MIN_SHAPE_SIZE,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  PAGE_LAYOUT_MARGIN_PRESETS,
  PAGE_LAYOUT_PAPER_SIZES,
  PAGE_LAYOUT_THEMES,
  state,
} from "../../app-core.js";
import {
  cwsAgentCloneRange,
  cwsAgentCloneRanges,
} from "./contracts.js";
import {
  cwsAgentShapeTextTarget,
} from "./normalization.js";
import {
  cwsAgentQuickAnalysisModeLabel,
  cwsAgentShapeBoxForOp,
  cwsAgentUniqueTableName,
} from "./operations.js";
import {
  applyDataTable,
  countDataTableOverwriteCells,
  dataTableFormulaPoints,
  dataTableInputResultConflicts,
  dataTableInputTarget,
  dataTableResultCells,
  dataTableResultRange,
  runGoalSeek,
} from "../data/analysis-and-scenarios.js";
import {
  createForecastSheetFromSelection,
  forecastEndValueFromDialogText,
  forecastPeriodsForOptions,
  forecastSeriesPointsForRange,
} from "../data/forecast-and-statistics.js";
import {
  rangeHasUserData,
} from "../data/sort-and-autofilter.js";
import {
  consolidateHeaders,
  consolidateOutputPlacement,
  consolidateRanges,
  consolidateTargetValidation,
  insertSubtotalRows,
  removeSubtotalRows,
  subtotalGroupsForRange,
  validateCellInput,
  validationMarkStore,
  workbookCellReference,
} from "../data/validation-and-consolidation.js";
import {
  addShapeAtPixelBoxToSheet,
  chartDataFromRange,
  chartDescriptor,
  chartTypeLabel,
  lineBoxAndPointsFromEndpoints,
  pivotSummaryFromRange,
  uniqueSheetNameWithBase,
  writePivotSummarySheet,
} from "../drawing/charts-and-insertion.js";
import {
  addWorksheetImageDescriptorsToSheet,
  headerFooterPlainText,
  headerFooterSectionHasText,
  isLineShapeType,
  normalizedHeaderFooterSettings,
  shapeDescriptorForType,
} from "../drawing/shape-rendering.js";
import {
  refreshShapeSource,
  shapeConnectionPointForSite,
} from "../drawing/transforms-and-connectors.js";
import {
  mergesIntersectingRange,
  refreshWorkbookChartsAfterSourceEdit,
} from "../editing/autofill.js";
import {
  normalizedThemeColorPalette,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  cloneCellForClipboard,
} from "../editing/keyboard-and-paste.js";
import {
  activateObjectFormatRibbonForSelection,
  activeSelectionRange,
  activeSelectionRanges,
  coerceUserInput,
  expandRangeForMerges,
  getCellRawInput,
  getDisplayForCell,
  initialCellDisplayText,
  isCellInRange,
  normalizeRange,
  rangeToLabel,
  scheduleSelectionUiSync,
  syncShapeFormatControls,
  updateSelectionUi,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  cellHasComment,
  cellHasNote,
  coordinateToPixels,
  imageBoxToPixels,
  scrollRangeIntoView,
  setSelectedImageIds,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureCellsForHistory,
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneDefinedNames,
  cloneJsonLikeValue,
  cwsAgentCaptureWorkbookSheetsSnapshot,
  internCellsInRange,
  internSheetCellStyles,
  mutableCssObject,
  pasteValueFromClipboardCell,
  pushHistory,
  pushSheetDeleteHistory,
  pushSheetInsertHistory,
  pushSheetMoveHistory,
  restoreSheetsFromHistory,
  storeCellModelForSheet,
  writeClonedClipboardCellContent,
} from "../editing/structure-and-history.js";
import {
  ensureCellModel,
  forEachCellInRange,
  rangesIntersect,
  removeIntersectingMerges,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  appendCommentReply,
  clearCellsInRange,
  clipboardCellEntries,
  clipboardHasFormula,
  closeCommentThreadPopup,
  isRangeInsideRange,
  makeCommentFromText,
  quickAnalysisTotalLabel,
  quickAnalysisTotalPlanForSheet,
  removeCellModelIfEmpty,
  writeQuickAnalysisTotalPlan,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  columnName,
  columnWidthToPixels,
  rowHeightToPixels,
  setStatus,
} from "../ui/help-and-utilities.js";
import {
  applyExcelTableStyle,
  definedNameDuplicateMessage,
  definedNameHasConflict,
  definedNameKey,
  ensureSheetSize,
  upsertDefinedName,
} from "../ui/ribbon-and-links.js";
import {
  clearSheetSelectionState,
  cloneSheetModel,
  createBlankSheet,
  defaultSheetProtectionOptions,
  ensureEditableRangeForProtection,
  ensureEditableRangesForProtection,
  ensureSheetProtectionPermission,
  isSheetProtected,
  moveSheetToIndex,
  nearestVisibleSheetIndex,
  quickAnalysisApplyColorScale,
  quickAnalysisApplyDataBars,
  quickAnalysisApplyDuplicateValues,
  quickAnalysisApplyGreaterThanAverage,
  quickAnalysisApplyIconSet,
  quickAnalysisClearFormatsForSheet,
  quickAnalysisNumericCells,
  quickAnalysisNumericColumns,
  rangeHasLockedCellsForProtection,
  setSheetTabColor,
} from "../ui/sheet-protection-and-translation.js";
import {
  captureWorkbookViewState,
  syncViewRibbonButtons,
} from "../ui/views-and-formula-tools.js";
import {
  buildFormulaEngine,
  ensureActiveSheetVisible,
  isSheetVisible,
  normalizeDefinedNameLookup,
  renderWorkbook,
  setActiveSheet,
  visibleSheetIndexes,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  normalizePaneState,
  normalizedSheetShowRuler,
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  parseTableStyleInfoXml,
} from "../workbook/ooxml-drawings.js";
import {
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
  activateCwsAgentPageSetupSheet,
  commitCwsAgentSheetViewChange,
  cwsAgentFreezePanesStatus,
  cwsAgentPageBreakPositionLabel,
  cwsAgentSheetViewHasChanges,
  cwsAgentSheetViewPatchLabels,
  cwsAgentSheetViewSnapshot,
  refreshCwsAgentWorkbookAfterSheetStructureChange,
} from "../workbook/startup-and-loading.js";
import {
  applyTableRangeEditBehaviors,
  rangeIntersectsWorksheetTable,
  tableRangeEditMayChangeMetadata,
} from "../workbook/tables-and-validation.js";
import {
  applyPageLayoutEffectToSelectedShapes,
  commitPageLayoutChange,
  ensurePageLayoutSettings,
  normalizedPageLayoutThemeEffects,
  pageLayoutEffectPreset,
} from "../workbook/vba-and-printing.js";
import {
  applySheetZoomWithoutHistory,
} from "../workbook/views-and-macros.js";


function applyCwsAgentCreateTableOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  if (rangeIntersectsWorksheetTable(sheet, range)) {
    throw new Error("指定範囲は既存のテーブルと重なっています。");
  }
  if (!ensureEditableRangeForProtection(sheet, range, "テーブルを作成")) {
    throw new Error("保護されたセルを含む範囲ではテーブルを作成できません。");
  }

  const tableName = cwsAgentUniqueTableName(sheet, op.name);
  const sheetBefore = captureSheetForHistory(sheet);
  ensureSheetSize(sheet, range.bottom, range.right);
  applyExcelTableStyle(sheet, range, { hasHeaders: op.hasHeaders });
  sheet.tables = Array.isArray(sheet.tables) ? sheet.tables : [];
  sheet.tables.push({
    id: uniqueWorksheetObjectId("table"),
    name: tableName,
    displayName: tableName,
    range: { ...range },
    showHeaderRow: true,
    showTotalsRow: false,
    showFilterButton: op.showFilterButton !== false,
    style: op.style || "TableStyleMedium2",
    styleInfo: parseTableStyleInfoXml(null),
    hasHeaders: op.hasHeaders !== false,
    autoFilterState: op.showFilterButton === false
      ? null
      : {
          range: { ...range },
          criteria: {},
          hiddenButtons: {},
        },
  });
  state.hf = buildFormulaEngine(state.model);
  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  pushHistory({
    sheetIndex: op.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent テーブル`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} にテーブル ${tableName} を作成しました。`);

  return {
    type: "createTable",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    tableName,
    hasHeaders: op.hasHeaders !== false,
    showFilterButton: op.showFilterButton !== false,
    style: op.style || "TableStyleMedium2",
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentCreateChartOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを挿入できません。")) {
    throw new Error("保護されたシートではグラフを挿入できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const chart = chartDataFromRange(sheet, range, op.chartType);
  if (!chart.series.length || !chart.series.some((series) => series.values.some((value) => Number.isFinite(value)))) {
    throw new Error("グラフを作成するための数値データが見つかりません。");
  }

  const descriptor = chartDescriptor(chart, op.chartType, range, {
    sourceSheet: sheet,
    sourceSheetName: sheet.name,
    title: op.title || undefined,
    width: op.width,
    height: op.height,
  });
  const anchor = {
    sheetIndex: op.sheetIndex,
    row: op.anchorRow,
    col: op.anchorCol,
  };
  const sheetBefore = captureSheetForHistory(sheet);
  const inserted = addWorksheetImageDescriptorsToSheet(sheet, [descriptor], anchor, { label: "グラフ" });
  if (!inserted.length) {
    throw new Error("グラフを挿入できませんでした。");
  }

  const image = inserted[0];
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  state.selectionDrag = null;
  state.formulaEdit = null;
  setSelectedImageIds([image.id], image.id);
  $formulaInput.prop("disabled", true).val("");
  pushHistory({
    sheetIndex: op.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: range,
    selectionAfter: null,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [],
    label: `${rangeToLabel(range)} CWS Agent グラフ`,
  });
  renderSheet();
  updateSelectionUi();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} から${chartTypeLabel(op.chartType)}を挿入しました。`);

  return {
    type: "createChart",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    chartType: op.chartType,
    title: descriptor.chart?.title || descriptor.name || "",
    imageId: image.id,
    anchorRow: op.anchorRow,
    anchorCol: op.anchorCol,
    seriesCount: chart.series.length,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentCreatePivotTableOp(op) {
  const sourceSheet = state.model.sheets[op.sheetIndex];
  if (!sourceSheet || !isSheetVisible(sourceSheet)) {
    throw new Error("ピボットテーブルの元シートが見つかりません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const summary = pivotSummaryFromRange(sourceSheet, range);
  if (!summary.rows.length) {
    throw new Error("ピボットテーブルを作成するためのデータが見つかりません。");
  }

  const existingNames = state.model.sheets.map((item) => item.name);
  const finalName = op.name && !existingNames.includes(op.name)
    ? op.name
    : uniqueSheetNameWithBase(op.name || "PivotTable", existingNames);
  const sheet = createBlankSheet(finalName);
  ensureSheetSize(sheet, summary.rows.length + 2, summary.headers.length + 1);
  const insertIndex = Math.max(0, Math.min(Number(op.insertIndex) || state.model.sheets.length, state.model.sheets.length));
  const activeSheetBefore = state.activeSheetIndex;
  const pivotRange = writePivotSummarySheet(sheet, summary, insertIndex);
  state.model.sheets.splice(insertIndex, 0, sheet);
  state.activeSheetIndex = insertIndex;
  state.selectedSheetIndexes = [];
  state.hf = buildFormulaEngine(state.model);
  const point = { sheetIndex: state.activeSheetIndex, row: pivotRange.top, col: pivotRange.left };
  state.selected = point;
  state.selectionRange = { ...pivotRange, sheetIndex: state.activeSheetIndex };
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = point;
  state.selectionAnchor = point;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  state.formulaEdit = null;
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  setStatus(`CWS Agent: ${finalName} にピボットテーブルを作成しました。`);
  pushSheetInsertHistory(insertIndex, sheet, activeSheetBefore, insertIndex, `${finalName} CWS Agent ピボットテーブル`, {
    selectionAfter: state.selectionRange,
    selectionAfterRanges: [state.selectionRange],
  });
  scheduleWorkbookPersistence();

  return {
    type: "createPivotTable",
    sheetIndex: op.sheetIndex,
    sheetName: sourceSheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    pivotSheetIndex: insertIndex,
    pivotSheetName: finalName,
    targetTop: pivotRange.top,
    targetLeft: pivotRange.left,
    targetBottom: pivotRange.bottom,
    targetRight: pivotRange.right,
    rows: summary.rows.length,
    columns: summary.headers.length,
    address: `${sourceSheet.name}!${rangeToLabel(range)}`,
    targetAddress: `${finalName}!${rangeToLabel(pivotRange)}`,
  };
}



function applyCwsAgentCreateForecastSheetOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("予測元シートを開けませんでした。");
  }

  const sourceSheet = state.model.sheets[op.sheetIndex];
  if (!sourceSheet || !isSheetVisible(sourceSheet)) {
    throw new Error("予測元シートが見つかりません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const options = {
    range,
    periods: op.periods,
    chartType: op.chartType,
    confidence: op.confidence,
    confidenceLevel: op.confidenceLevel,
    seasonality: op.seasonality,
    missingData: op.missingData,
    duplicateAggregate: op.duplicateAggregate,
  };
  if (op.forecastEndText) {
    const parsedEnd = forecastEndValueFromDialogText(op.forecastEndText, sourceSheet, range);
    if (!Number.isFinite(parsedEnd)) {
      throw new Error("予測終了には有効なタイムライン値を指定してください。");
    }
    options.forecastEnd = parsedEnd;
  } else if (Number.isFinite(Number(op.forecastEnd))) {
    options.forecastEnd = Number(op.forecastEnd);
  }

  const points = forecastSeriesPointsForRange(sourceSheet, range, options);
  if (points.length < 2) {
    throw new Error("予測を作成するには、2つ以上の数値データが必要です。");
  }
  if (Number.isFinite(Number(options.forecastEnd)) && options.forecastEnd <= points[points.length - 1].x) {
    throw new Error("予測終了には、最後のタイムライン値より後の値を指定してください。");
  }
  const periods = forecastPeriodsForOptions(points, options);
  const beforeCount = state.model.sheets.length;
  createForecastSheetFromSelection(options);
  if (state.model.sheets.length <= beforeCount) {
    throw new Error("予測ワークシートを作成できませんでした。");
  }
  const forecastSheetIndex = state.model.sheets.length - 1;
  const forecastSheet = state.model.sheets[forecastSheetIndex];
  const right = op.confidence === false ? 3 : 5;
  const targetRange = normalizeRange(
    { sheetIndex: forecastSheetIndex, row: 1, col: 1 },
    { sheetIndex: forecastSheetIndex, row: points.length + periods + 1, col: right },
  );

  return {
    type: "createForecastSheet",
    sheetIndex: op.sheetIndex,
    sheetName: sourceSheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    forecastSheetIndex,
    forecastSheetName: forecastSheet.name,
    targetTop: targetRange.top,
    targetLeft: targetRange.left,
    targetBottom: targetRange.bottom,
    targetRight: targetRange.right,
    points: points.length,
    periods,
    chartType: op.chartType,
    confidence: op.confidence !== false,
    confidenceLevel: op.confidenceLevel,
    seasonality: op.seasonality,
    missingData: op.missingData,
    duplicateAggregate: op.duplicateAggregate,
    address: `${sourceSheet.name}!${rangeToLabel(range)}`,
    targetAddress: `${forecastSheet.name}!${rangeToLabel(targetRange)}`,
  };
}



function applyCwsAgentGoalSeekOp(op) {
  const setCell = { sheetIndex: op.setSheetIndex, row: op.setRow, col: op.setCol };
  const changingCell = { sheetIndex: op.changingSheetIndex, row: op.changingRow, col: op.changingCol };
  const setSheet = state.model?.sheets?.[setCell.sheetIndex];
  const changingSheet = state.model?.sheets?.[changingCell.sheetIndex];
  if (!setSheet || !isSheetVisible(setSheet) || !changingSheet || !isSheetVisible(changingSheet)) {
    throw new Error("ゴール シーク対象シートが見つかりません。");
  }

  const changingRange = normalizeRange(
    { sheetIndex: changingCell.sheetIndex, row: changingCell.row, col: changingCell.col },
    { sheetIndex: changingCell.sheetIndex, row: changingCell.row, col: changingCell.col },
  );
  if (!ensureEditableRangeForProtection(changingSheet, changingRange, "ゴール シーク")) {
    throw new Error("保護されたシートではゴール シークを実行できません。");
  }

  const result = runGoalSeek(setCell, op.targetValue, changingCell);
  if (!result) {
    throw new Error("ゴール シークを実行できませんでした。");
  }

  return {
    type: "goalSeek",
    sheetIndex: changingCell.sheetIndex,
    sheetName: changingSheet.name,
    setCell: result.setCell,
    changingCell: result.changingCell,
    targetValue: result.targetValue,
    finalValue: result.finalValue,
    changingValue: result.changingValue,
    success: Boolean(result.success),
    iterations: result.iterations || 0,
    address: result.changingCell || `${changingSheet.name}!${rangeToLabel(changingRange)}`,
  };
}



function applyCwsAgentDataTableOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("データ テーブル対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error("データ テーブル対象シートが見つかりません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const rowInput = op.rowInputRow && op.rowInputCol
    ? { sheetIndex: op.rowInputSheetIndex ?? op.sheetIndex, row: op.rowInputRow, col: op.rowInputCol }
    : null;
  const columnInput = op.columnInputRow && op.columnInputCol
    ? { sheetIndex: op.columnInputSheetIndex ?? op.sheetIndex, row: op.columnInputRow, col: op.columnInputCol }
    : null;
  if (!rowInput && !columnInput) {
    throw new Error("データ テーブルの代入セルが指定されていません。");
  }

  const rowInputTarget = dataTableInputTarget(rowInput);
  const columnInputTarget = dataTableInputTarget(columnInput);
  if ((rowInput && !rowInputTarget) || (columnInput && !columnInputTarget)) {
    throw new Error("データ テーブルの代入セルが正しくありません。");
  }
  if (rowInputTarget && columnInputTarget && rowInputTarget.snapshotKey === columnInputTarget.snapshotKey) {
    throw new Error("行の代入セルと列の代入セルには別のセルを指定してください。");
  }
  const protectedInput = [rowInputTarget, columnInputTarget].filter(Boolean).find((target) => {
    const inputRange = normalizeRange(
      { sheetIndex: target.sheetIndex, row: target.row, col: target.col },
      { sheetIndex: target.sheetIndex, row: target.row, col: target.col },
    );
    return rangeHasLockedCellsForProtection(target.sheet, inputRange);
  });
  if (protectedInput) {
    throw new Error(`保護されたシート上のロックされたセルは代入セルに指定できません: ${workbookCellReference(protectedInput.sheetIndex, protectedInput.row, protectedInput.col)}`);
  }

  const options = { rowInput, columnInput };
  const formulas = dataTableFormulaPoints(sheet, range, options);
  if (!formulas.length) {
    throw new Error("データ テーブルに使用する数式が見つかりません。");
  }
  const resultCells = dataTableResultCells(range, options, formulas);
  const resultRange = dataTableResultRange(range, resultCells);
  if (resultRange && rangeHasLockedCellsForProtection(sheet, resultRange)) {
    throw new Error("保護されたシート上のロックされたセルにはデータ テーブルの結果を出力できません。");
  }
  if (resultRange && mergesIntersectingRange(sheet, resultRange).length) {
    throw new Error("結合されたセルを含む範囲にはデータ テーブルの結果を出力できません。");
  }
  const inputConflicts = dataTableInputResultConflicts([rowInputTarget, columnInputTarget], range, resultCells);
  if (inputConflicts.length) {
    throw new Error(`代入セルはデータ テーブルの結果範囲外に指定してください: ${inputConflicts.join(", ")}`);
  }
  const overwriteCount = countDataTableOverwriteCells(sheet, range, options, formulas);
  if (overwriteCount > 0 && !op.overwrite) {
    throw new Error("データ テーブルの出力先には既存のデータがあります。overwrite:true を指定してください。");
  }

  const ok = applyDataTable(range, options);
  if (!ok) {
    throw new Error("データ テーブルを実行できませんでした。");
  }

  const targetRange = resultRange || range;
  return {
    type: "dataTable",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    targetTop: targetRange.top,
    targetLeft: targetRange.left,
    targetBottom: targetRange.bottom,
    targetRight: targetRange.right,
    rowInput: rowInput ? workbookCellReference(rowInput.sheetIndex, rowInput.row, rowInput.col) : "",
    columnInput: columnInput ? workbookCellReference(columnInput.sheetIndex, columnInput.row, columnInput.col) : "",
    formulas: formulas.length,
    written: resultCells.length,
    overwritten: overwriteCount,
    address: `${sheet.name}!${rangeToLabel(range)}`,
    targetAddress: `${sheet.name}!${rangeToLabel(targetRange)}`,
  };
}



function applyCwsAgentExistingShapeEditOp(op) {
  const sheet = state.model?.sheets?.[op.sheetIndex];
  if (!sheet || !ensureSheetProtectionPermission(sheet, 'allowEditObjects', '保護されたシートでは図形を編集できません。')) throw new Error('図形は保護されています。');
  const sheetBefore = captureSheetForHistory(sheet);
  const image = cwsApplyExistingShapeEdit(sheet,op);
  if (op.type === 'setShapeStyle') refreshShapeSource(image);
  else setSelectedImageIds([]);
  pushHistory({sheetIndex:op.sheetIndex,sheetBefore,sheetAfter:captureSheetForHistory(sheet),label:`${image.name || '図形'} CWS Agent 編集`});
  renderSheet(); updateSelectionUi(); scheduleWorkbookPersistence();
  return {type:op.type,sheetIndex:op.sheetIndex,sheetName:sheet.name,shapeId:image.id,address:`${sheet.name}!${image.name || image.id}`};
}

function applyCwsAgentSetShapeTextOp(op) {
  const {sheet,image} = cwsAgentShapeTextTarget(op);
  if (!ensureSheetProtectionPermission(sheet, 'allowEditObjects', '保護されたシートでは図形を編集できません。')) throw new Error('図形は保護されています。');
  const sheetBefore = captureSheetForHistory(sheet);
  image.shape = {...image.shape, text:op.text};
  delete image.shape.textRuns;
  refreshShapeSource(image);
  pushHistory({sheetIndex:op.sheetIndex,sheetBefore,sheetAfter:captureSheetForHistory(sheet),label:`${image.name || '図形'} CWS Agent テキスト編集`});
  renderSheet(); updateSelectionUi(); scheduleWorkbookPersistence();
  return {type:'setShapeText',sheetIndex:op.sheetIndex,sheetName:sheet.name,shapeId:image.id,address:`${sheet.name}!${image.name || image.id}`};
}



function applyCwsAgentAddShapeOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("図形を挿入するシートを開けませんでした。");
  }
  const sheet = state.model.sheets[op.sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error("図形を挿入するシートが見つかりません。");
  }
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを挿入できません。")) {
    throw new Error("保護されたシートではオブジェクトを挿入できません。");
  }

  const sheetBefore = captureSheetForHistory(sheet);
  // New sheets need not carry default dimension fields. Resolve cell/range
  // coordinates only after materializing their rows and columns, otherwise
  // a shape beyond the existing grid can collapse to the minimum pixel size.
  cwsAgentEnsureShapeGrid(sheet, op);
  const shapeBox = cwsAgentShapeConnectorBoxForOp(op, sheet) || cwsAgentShapePlacementBoxForOp(op, sheet);
  const descriptor = shapeDescriptorForType(op.shapeType, {
    ...op.style,
    width: shapeBox.width,
    height: shapeBox.height,
  });
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  const image = addShapeAtPixelBoxToSheet(sheet, op.shapeType, shapeBox, descriptor);
  if (!image) {
    throw new Error("図形を挿入できませんでした。");
  }
  if (op.name) {
    image.name = op.name;
  }
  if (op.layoutRole || op.componentId || op.parentShape) image.cwsLayout={role:op.layoutRole,componentId:op.componentId,parentShape:op.parentShape,componentLabelShape:op.componentLabelShape};

  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  state.selectionDrag = null;
  state.formulaEdit = null;
  setSelectedImageIds([image.id], image.id);
  $formulaInput.prop("disabled", true).val("");
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore,
    selectionAfter: null,
    selectionBeforeRanges,
    selectionAfterRanges: [],
    label: `${image.name || descriptor.name || "図形"} CWS Agent 図形`,
  });
  renderSheet();
  updateSelectionUi();
  activateObjectFormatRibbonForSelection();
  scheduleWorkbookPersistence();
  setStatus(`CWS Agent: ${sheet.name} に${image.name || descriptor.name || "図形"}を挿入しました。`);

  return {
    type: "addShape",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    shapeType: image.shape?.type || op.shapeType,
    shapeId: image.id,
    name: image.name,
    left: shapeBox.left,
    top: shapeBox.top,
    width: shapeBox.width,
    height: shapeBox.height,
    address: `${sheet.name}!${op.anchor || "図形"}`,
  };
}



function applyCwsAgentCreateWorkbookTemplateOp(op) {
  const workbookSheetsBefore = cwsAgentCaptureWorkbookSheetsSnapshot();
  const activeSheetBefore = state.activeSheetIndex;
  const selectionBefore = cwsAgentCloneRange(state.selectionRange);
  const selectionBeforeRanges = cwsAgentCloneRanges(state.selectionRanges);
  const insertAt = op.replaceWorkbook
    ? 0
    : Math.max(0, Math.min(state.activeSheetIndex + 1, state.model.sheets.length));
  const sheets = op.sheets.map((sheetDef) => cwsAgentBuildTemplateSheet(sheetDef));
  if (op.replaceWorkbook) {
    state.model.sheets = sheets;
  } else {
    state.model.sheets.splice(insertAt, 0, ...sheets);
  }
  state.activeSheetIndex = insertAt;
  state.selectedSheetIndexes = sheets.map((_sheet, offset) => insertAt + offset);
  const selectionAfter = { sheetIndex: insertAt, top: 1, left: 1, bottom: 1, right: 1 };
  state.selected = { sheetIndex: insertAt, row: 1, col: 1 };
  state.selectionRange = selectionAfter;
  state.selectionRanges = [selectionAfter];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.formulaEdit = null;
  refreshCwsAgentWorkbookAfterSheetStructureChange({ preserveSheetSelection: true });
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    workbookSheetsBefore,
    workbookSheetsAfter: cwsAgentCaptureWorkbookSheetsSnapshot(),
    activeSheetBefore,
    activeSheetAfter: state.activeSheetIndex,
    selectionBefore,
    selectionAfter,
    selectionBeforeRanges,
    selectionAfterRanges: [selectionAfter],
    label: `${op.templateId || "Workbook template"} CWS Agent テンプレート作成`,
  });
  setStatus(`CWS Agent: ${sheets.length} シートのテンプレートを作成しました。`);
  return {
    type: "createWorkbookTemplate",
    templateId: op.templateId || "",
    sheetIndex: insertAt,
    sheetName: sheets[0]?.name || "",
    sheets: sheets.map((sheet) => sheet.name),
    address: sheets.map((sheet) => sheet.name).join(", "),
  };
}



function cwsAgentBuildTemplateSheet(sheetDef) {
  const rows = sheetDef.rows || [[]];
  const rowCount = Math.max(1, rows.length);
  const colCount = Math.max(1, ...rows.map((row) => row.length));
  const sheet = createBlankSheet(sheetDef.name);
  const defaultColumnWidth = sheet.defaultColumnWidth || columnWidthToPixels(8.43);
  const defaultRowHeight = sheet.defaultRowHeight || rowHeightToPixels(15);
  sheet.rowCount = rowCount;
  sheet.colCount = colCount;
  sheet.baseRowCount = rowCount;
  sheet.baseColCount = colCount;
  sheet.columns = Array.from({ length: colCount }, (_item, index) => ({
    width: Number(sheetDef.columnWidths?.[index]) || defaultColumnWidth,
    hidden: false,
  }));
  sheet.rows = Array.from({ length: rowCount }, (_item, index) => ({
    height: Number(sheetDef.rowHeights?.[index]) || defaultRowHeight,
    hidden: false,
  }));
  const titleRows = new Set(sheetDef.titleRows || [1]);
  const headerRows = new Set(sheetDef.headerRows || []);
  rows.forEach((rowValues, rowIndex) => {
    rowValues.forEach((value, colIndex) => {
      const row = rowIndex + 1;
      const col = colIndex + 1;
      const content = coerceUserInput(value, { preserveWhitespace: true });
      if (content === "") return;
      const isFormula = typeof content === "string" && content.startsWith("=");
      const cell = {
        row,
        col,
        raw: content,
        formula: isFormula ? content : null,
        cached: isFormula ? null : content,
        display: "",
        hyperlink: null,
        comment: null,
        note: null,
        noteVisible: false,
        validation: null,
        richText: null,
        css: cwsAgentTemplateCellCss(row, col, titleRows, headerRows),
        hasFill: titleRows.has(row) || headerRows.has(row),
        hasFontColor: titleRows.has(row),
        borders: cwsAgentTemplateCellBorders(row, col, rowValues.length),
        numFmt: "",
      };
      cell.display = initialCellDisplayText(cell, content, isFormula);
      storeCellModelForSheet(sheet, cellKey(row, col), cell);
    });
  });
  internSheetCellStyles(sheet);
  return sheet;
}



function cwsAgentTemplateCellCss(row, col, titleRows, headerRows) {
  if (titleRows.has(row)) {
    return {
      backgroundColor: "#1f4e79",
      color: "#ffffff",
      fontWeight: "700",
      textAlign: col === 1 ? "left" : "center",
      verticalAlign: "middle",
      whiteSpace: "pre-wrap",
      overflowWrap: "anywhere",
    };
  }
  if (headerRows.has(row)) {
    return {
      backgroundColor: "#d9eaf7",
      fontWeight: "700",
      verticalAlign: "middle",
      whiteSpace: "pre-wrap",
      overflowWrap: "anywhere",
    };
  }
  return {
    verticalAlign: "top",
    whiteSpace: "pre-wrap",
    overflowWrap: "anywhere",
  };
}



function cwsAgentTemplateCellBorders(row, col, width) {
  const color = "#b7c9d6";
  const side = { style: "solid", color, width: 1 };
  return {
    top: row === 1 ? side : null,
    right: col === width ? side : null,
    bottom: side,
    left: col === 1 ? side : null,
  };
}



function cwsAgentEnsureShapeGrid(sheet, op) {
  const location = op.placement?.location;
  const points = [location?.parsedCell, op.connector?.fromCell, op.connector?.toCell].filter(Boolean);
  const range = location?.parsedRange;
  const rows = Math.max(1, range?.bottom || 1, ...points.map(point => point.row));
  const cols = Math.max(1, range?.right || 1, ...points.map(point => point.col));
  if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows > EXCEL_MAX_ROWS || cols > EXCEL_MAX_COLS) {
    throw new Error("図形の配置範囲がシートの上限を超えています。");
  }
  ensureSheetSize(sheet, rows, cols);
}



function cwsAgentShapePlacementBoxForOp(op, sheet) {
  // A batch may resize rows/columns before inserting shapes. Resolve cell/range
  // placement against the current sheet at execution time, like line endpoints.
  return op.placement
    ? cwsAgentShapeBoxForOp(op.placement.dimensions, 0, sheet, op.placement.location, op.shapeType, op.style)
    : op.box;
}



function cwsAgentShapeConnectorBoxForOp(op, sheet) {
  if (!isLineShapeType(op?.shapeType) || !op?.connector) return null;
  if (op.connector.fromCell && op.connector.toCell) {
    const start = cwsAgentCellEndpointPoint(sheet, op.connector.fromCell, op.connector.fromSite);
    const end = cwsAgentCellEndpointPoint(sheet, op.connector.toCell, op.connector.toSite);
    return cwsAgentConnectorBoxFromEndpoints(op, start, end);
  }
  const from = cwsAgentFindShapeForConnector(sheet, op.connector.fromShape);
  const to = cwsAgentFindShapeForConnector(sheet, op.connector.toShape);
  if (!from || !to) {
    throw new Error(`線の接続先図形が見つかりません: ${op.connector.fromShape} -> ${op.connector.toShape}`);
  }
  const inferredSites = cwsAgentInferConnectorSites(sheet, from, to);
  const fromSite = op.connector.fromSite || inferredSites.fromSite;
  const toSite = op.connector.toSite || inferredSites.toSite;
  if (!fromSite || !toSite) {
    throw new Error(`線の接続点指定がありません: ${op.connector.fromShape} -> ${op.connector.toShape}`);
  }
  const start = shapeConnectionPointForSite(sheet, from, fromSite);
  const end = shapeConnectionPointForSite(sheet, to, toSite);
  if (!start || !end) {
    throw new Error(`線の接続点が見つかりません: ${fromSite} -> ${toSite}`);
  }
  return cwsAgentConnectorBoxFromEndpoints(op, start, end, {
    start: { imageId: from.id, site: fromSite },
    end: { imageId: to.id, site: toSite },
  });
}



function cwsAgentCellEndpointPoint(sheet, cell, site = "center") {
  if (cell.sheetName && cell.sheetName !== sheet.name) throw new Error("線の端点は同じシート内に指定してください。");
  const left = coordinateToPixels(sheet, "col", cell.col - 1);
  const right = coordinateToPixels(sheet, "col", cell.col);
  const top = coordinateToPixels(sheet, "row", cell.row - 1);
  const bottom = coordinateToPixels(sheet, "row", cell.row);
  return { x: site === "left" ? left : site === "right" ? right : (left + right) / 2,
    y: site === "top" ? top : site === "bottom" ? bottom : (top + bottom) / 2 };
}



function cwsAgentConnectorBoxFromEndpoints(op, start, end, connections = null) {
  if (Math.hypot(start.x - end.x, start.y - end.y) < 1) throw new Error("線の両端に異なる位置を指定してください。");
  const result = lineBoxAndPointsFromEndpoints(start, end, { shape: { type: op.shapeType, ...(op.style || {}) } });
  return {
    left: Math.max(0, Math.round(result.box.left)),
    top: Math.max(0, Math.round(result.box.top)),
    // Points are normalized against this exact box. Clamping the box alone
    // changes their absolute endpoints and disconnects long sequence messages.
    width: Math.max(CWS_AGENT_MIN_SHAPE_SIZE, Math.round(result.box.width)),
    height: Math.max(CWS_AGENT_MIN_SHAPE_SIZE, Math.round(result.box.height)),
    points: result.points,
    ...(connections ? { connections } : {}),
  };
}



function cwsAgentInferConnectorSites(sheet, from, to) {
  const fromBox = imageBoxToPixels(sheet, from);
  const toBox = imageBoxToPixels(sheet, to);
  if (!fromBox || !toBox) return { fromSite: "", toSite: "" };
  const fromCenter = {
    x: fromBox.left + fromBox.width / 2,
    y: fromBox.top + fromBox.height / 2,
  };
  const toCenter = {
    x: toBox.left + toBox.width / 2,
    y: toBox.top + toBox.height / 2,
  };
  const dx = toCenter.x - fromCenter.x;
  const dy = toCenter.y - fromCenter.y;
  if (Math.abs(dx) >= Math.abs(dy)) {
    return dx >= 0
      ? { fromSite: "right", toSite: "left" }
      : { fromSite: "left", toSite: "right" };
  }
  return dy >= 0
    ? { fromSite: "bottom", toSite: "top" }
    : { fromSite: "top", toSite: "bottom" };
}



function cwsAgentFindShapeForConnector(sheet, reference) {
  const text = String(reference || "").trim();
  if (!text || !sheet?.images?.length) return null;
  const normalizedText = text.toLowerCase();
  const images = [...sheet.images].reverse();
  return images.find((image) =>
    String(image?.id || "").trim() === text ||
    String(image?.name || "").trim() === text ||
    String(image?.shape?.label || "").trim() === text ||
    String(image?.shape?.text || "").trim() === text ||
    String(image?.id || "").trim().toLowerCase() === normalizedText ||
    String(image?.name || "").trim().toLowerCase() === normalizedText ||
    String(image?.shape?.label || "").trim().toLowerCase() === normalizedText ||
    String(image?.shape?.text || "").trim().toLowerCase() === normalizedText
  ) || null;
}



function applyCwsAgentSetQuickAnalysisOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではセルの書式を変更できません。")) {
    throw new Error("保護されたシートではクイック分析書式を変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);

  const applicable = cwsAgentQuickAnalysisApplicableCount(sheet, range, op.mode);
  if (!applicable) {
    throw new Error(`${cwsAgentQuickAnalysisModeLabel(op.mode)}を適用できるセルが見つかりません。`);
  }

  const before = captureCellsForHistory(sheet, range);
  quickAnalysisClearFormatsForSheet(sheet, range);
  const applied = cwsAgentApplyQuickAnalysisMode(sheet, range, op.mode);
  if (!applied) {
    throw new Error(`${cwsAgentQuickAnalysisModeLabel(op.mode)}を適用できませんでした。`);
  }
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent クイック分析`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} に${cwsAgentQuickAnalysisModeLabel(op.mode)}を適用しました。`);

  return {
    type: "setQuickAnalysis",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    mode: op.mode,
    applied,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentClearQuickAnalysisOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではセルの書式を変更できません。")) {
    throw new Error("保護されたシートではクイック分析書式を変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const before = captureCellsForHistory(sheet, range);
  const cleared = quickAnalysisClearFormatsForSheet(sheet, range);
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent クイック分析クリア`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(cleared ? `CWS Agent: ${sheet.name}!${rangeToLabel(range)} のクイック分析書式をクリアしました。` : "CWS Agent: 対象範囲にはクイック分析書式がありません。");

  return {
    type: "clearQuickAnalysis",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    cleared,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSetQuickAnalysisTotalOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);

  const numericColumns = quickAnalysisNumericColumns(sheet, range);
  if (!numericColumns.length && op.method !== "count") {
    throw new Error("集計に使用できる数値データが見つかりません。");
  }
  const plan = quickAnalysisTotalPlanForSheet(sheet, range, op.method, op.sheetIndex);
  if (!plan?.targetRange || !plan.writes.length) {
    throw new Error("集計に使用できるデータが見つかりません。");
  }
  if (plan.targetRange.bottom > EXCEL_MAX_ROWS || plan.targetRange.right > EXCEL_MAX_COLS) {
    throw new Error("集計の出力先がシート範囲を超えています。");
  }
  if (!ensureEditableRangesForProtection(sheet, [plan.targetRange], "クイック分析")) {
    throw new Error("保護されたシートではクイック分析の集計を追加できません。");
  }

  const sheetBefore = captureSheetForHistory(sheet);
  const sheetId = getSheetId(sheet.name);
  writeQuickAnalysisTotalPlan(sheet, sheetId, plan);
  state.hf = buildFormulaEngine(state.model);
  state.selected = { sheetIndex: op.sheetIndex, row: plan.targetRange.top, col: plan.targetRange.left };
  state.selectionRange = plan.targetRange;
  state.selectionRanges = [plan.targetRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: range,
    selectionAfter: plan.targetRange,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [plan.targetRange],
    label: `${rangeToLabel(range)} CWS Agent クイック分析 ${quickAnalysisTotalLabel(op.method)}`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(plan.targetRange)} に${quickAnalysisTotalLabel(op.method)}を追加しました。`);

  return {
    type: "setQuickAnalysisTotal",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    targetTop: plan.targetRange.top,
    targetLeft: plan.targetRange.left,
    targetBottom: plan.targetRange.bottom,
    targetRight: plan.targetRange.right,
    method: op.method,
    written: plan.writes.length,
    address: `${sheet.name}!${rangeToLabel(range)}`,
    targetAddress: `${sheet.name}!${rangeToLabel(plan.targetRange)}`,
  };
}



function cwsAgentApplyQuickAnalysisMode(sheet, range, mode) {
  if (mode === "dataBar") return quickAnalysisApplyDataBars(sheet, range);
  if (mode === "colorScale") return quickAnalysisApplyColorScale(sheet, range);
  if (mode === "iconSet") return quickAnalysisApplyIconSet(sheet, range);
  if (mode === "greaterThanAverage") return quickAnalysisApplyGreaterThanAverage(sheet, range);
  if (mode === "duplicateValues") return quickAnalysisApplyDuplicateValues(sheet, range);
  return 0;
}



function cwsAgentQuickAnalysisApplicableCount(sheet, range, mode) {
  const numericCells = ["dataBar", "colorScale", "iconSet", "greaterThanAverage"].includes(mode)
    ? quickAnalysisNumericCells(sheet, range)
    : [];
  if (mode === "dataBar" || mode === "colorScale" || mode === "iconSet") return numericCells.length;
  if (mode === "greaterThanAverage") {
    if (!numericCells.length) return 0;
    const average = numericCells.reduce((sum, cell) => sum + cell.value, 0) / numericCells.length;
    return numericCells.filter((cell) => cell.value > average).length;
  }
  if (mode === "duplicateValues") {
    const sheetId = getSheetId(sheet.name);
    const buckets = new Map();
    forEachCellInRange(range, (row, col) => {
      const text = String(getDisplayForCell(sheet, sheetId, row, col) ?? "").trim();
      if (!text) return;
      const key = text.toLocaleLowerCase("ja-JP");
      buckets.set(key, (buckets.get(key) || 0) + 1);
    });
    let count = 0;
    buckets.forEach((value) => {
      if (value > 1) count += value;
    });
    return count;
  }
  return 0;
}



function applyCwsAgentSetSubtotalOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  if (rangeIntersectsWorksheetTable(sheet, range)) {
    throw new Error("テーブルを含む範囲には小計を設定できません。");
  }
  if (!ensureSheetProtectionPermission(sheet, "allowInsertRows", "保護されたシートでは小計を挿入できません。")) {
    throw new Error("保護されたシートでは小計を挿入できません。");
  }
  const groups = subtotalGroupsForRange(sheet, range, op.groupCol);
  if (!groups.length) {
    throw new Error("小計を挿入するデータがありません。");
  }
  const result = insertSubtotalRows(sheet, range, {
    groupCol: op.groupCol,
    targetCols: op.targetCols,
    method: op.method,
    replace: op.replace,
    pageBreak: op.pageBreak,
    summaryBelow: op.summaryBelow,
  }, {
    sheetIndex: op.sheetIndex,
    quiet: true,
  });
  if (!result) {
    throw new Error("小計を挿入できませんでした。");
  }
  return {
    type: "setSubtotal",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    targetBottom: result.targetRange?.bottom || range.bottom,
    groupCol: op.groupCol,
    targetCols: [...op.targetCols],
    method: op.method,
    inserted: result.inserted,
    removedSubtotalRows: result.removedSubtotalRows || 0,
    address: `${sheet.name}!${rangeToLabel(result.targetRange || range)}`,
  };
}



function applyCwsAgentClearSubtotalsOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  if (!ensureSheetProtectionPermission(sheet, "allowInsertRows", "保護されたシートでは小計を削除できません。")) {
    throw new Error("保護されたシートでは小計を削除できません。");
  }
  const result = removeSubtotalRows(sheet, range, {
    sheetIndex: op.sheetIndex,
    quiet: true,
  });
  return {
    type: "clearSubtotals",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    removed: result?.removed || 0,
    address: `${sheet.name}!${rangeToLabel(result?.targetRange || range)}`,
  };
}



function applyCwsAgentConsolidateOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const destinationRange = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const references = (op.references || []).map((reference) => ({ ...reference }));
  const options = {
    method: op.method,
    topRow: Boolean(op.topRow),
    leftColumn: Boolean(op.leftColumn),
    createLinks: Boolean(op.createLinks),
    sheetIndex: op.sheetIndex,
  };
  const dataCells = references.reduce((total, range) => {
    const rows = Math.max(0, range.bottom - range.top + 1 - (options.topRow ? 1 : 0));
    const cols = Math.max(0, range.right - range.left + 1 - (options.leftColumn ? 1 : 0));
    return total + (rows * cols);
  }, 0);
  if (!dataCells) {
    throw new Error("統合する値が参照範囲にありません。");
  }

  const headers = consolidateHeaders(references, options);
  const placement = consolidateOutputPlacement(sheet, destinationRange, references, options, headers, op.sheetIndex);
  const validation = consolidateTargetValidation(sheet, placement.range);
  if (!validation.valid) {
    throw new Error(validation.message || "統合の出力先が正しくありません。");
  }
  if (!ensureEditableRangeForProtection(sheet, placement.range, "統合")) {
    throw new Error("保護されたシートでは統合できません。");
  }
  if (!op.overwrite && rangeHasUserData(sheet, placement.range)) {
    throw new Error("統合の出力先には既存のデータがあります。overwrite:true を指定してください。");
  }

  const ok = consolidateRanges(sheet, destinationRange, references, {
    ...options,
    skipOverwriteCheck: true,
  });
  if (!ok) {
    throw new Error("統合を実行できませんでした。");
  }

  return {
    type: "consolidateRanges",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: destinationRange.top,
    left: destinationRange.left,
    bottom: destinationRange.bottom,
    right: destinationRange.right,
    targetTop: placement.range.top,
    targetLeft: placement.range.left,
    targetBottom: placement.range.bottom,
    targetRight: placement.range.right,
    references: references.length,
    method: op.method,
    topRow: options.topRow,
    leftColumn: options.leftColumn,
    createLinks: options.createLinks,
    overwritten: Boolean(op.overwrite),
    address: `${sheet.name}!${rangeToLabel(placement.range)}`,
  };
}



function cwsAgentCellMatchesFindQuery(sheet, sheetId, row, col, query, options = {}) {
  const raw = getCellRawInput(sheet, row, col);
  const display = getDisplayForCell(sheet, sheetId, row, col);
  const values = options.lookIn === "formulas"
    ? [raw]
    : options.lookIn === "values"
      ? [display]
      : raw === display ? [raw] : [raw, display];
  const needle = options.caseSensitive ? String(query) : String(query).toLocaleLowerCase("ja-JP");
  return values.some((value) => {
    const haystack = options.caseSensitive ? String(value ?? "") : String(value ?? "").toLocaleLowerCase("ja-JP");
    return options.wholeCell ? haystack === needle : haystack.includes(needle);
  });
}



function applyCwsAgentSetDataValidationOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートでは入力規則を変更できません。")) {
    throw new Error("保護されたシートでは入力規則を変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);
  const before = captureCellsForHistory(sheet, range);
  const marks = validationMarkStore(op.sheetIndex);
  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const cell = ensureCellModel(sheet, row, col);
    cell.validation = cloneJsonLikeValue(op.rule);
    delete marks[key];
    storeCellModelForSheet(sheet, key, cell);
  });
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent 入力規則`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} に入力規則を設定しました。`);

  return {
    type: "setDataValidation",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentClearDataValidationOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートでは入力規則を変更できません。")) {
    throw new Error("保護されたシートでは入力規則を変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const before = captureCellsForHistory(sheet, range);
  const marks = validationMarkStore(op.sheetIndex);
  let cleared = 0;
  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const cell = sheet.cells[key];
    if (cell?.validation) {
      delete cell.validation;
      cleared += 1;
      storeCellModelForSheet(sheet, key, cell);
    }
    delete marks[key];
  });
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent 入力規則クリア`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(cleared ? `CWS Agent: ${sheet.name}!${rangeToLabel(range)} の入力規則をクリアしました。` : "CWS Agent: 対象範囲には入力規則がありません。");

  return {
    type: "clearDataValidation",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    cleared,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSetHyperlinkOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではリンクを変更できません。")) {
    throw new Error("保護されたシートではリンクを変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);
  const before = captureCellsForHistory(sheet, range);
  const sheetId = getSheetId(sheet.name);

  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    let displayText = op.text;
    if (displayText === null) {
      const raw = getCellRawInput(sheet, row, col);
      const display = getDisplayForCell(sheet, sheetId, row, col);
      if (String(raw ?? display ?? "") === "") {
        displayText = op.hyperlink;
      }
    }

    if (displayText !== null) {
      const content = coerceUserInput(displayText, { preserveWhitespace: true });
      const validation = validateCellInput(sheet, row, col, content);
      if (!validation.valid) {
        throw new Error(`${columnName(col)}${row}: ${validation.message || "入力規則に違反しています。"}`);
      }
      writeCellContent(sheet, sheetId, row, col, content, { preserveWhitespace: true });
    }

    const updated = ensureCellModel(sheet, row, col);
    updated.hyperlink = op.hyperlink;
    updated.css = mutableCssObject(updated.css);
    if (!updated.css.color) {
      updated.css.color = "#0563c1";
      updated.hasFontColor = true;
    }
    if (!updated.css.textDecoration) {
      updated.css.textDecoration = "underline";
    }
    storeCellModelForSheet(sheet, key, updated);
  });
  internCellsInRange(sheet, range);
  refreshWorkbookChartsAfterSourceEdit(op.sheetIndex, range);
  state.hf = buildFormulaEngine(state.model);
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent リンク`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} にリンクを設定しました。`);

  return {
    type: "setHyperlink",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    hyperlink: op.hyperlink,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentClearHyperlinkOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではリンクを変更できません。")) {
    throw new Error("保護されたシートではリンクを変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const before = captureCellsForHistory(sheet, range);
  let cleared = 0;
  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const cell = sheet.cells[key];
    if (cell?.hyperlink) {
      delete cell.hyperlink;
      cleared += 1;
      storeCellModelForSheet(sheet, key, cell);
    }
  });
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent リンク解除`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(cleared ? `CWS Agent: ${sheet.name}!${rangeToLabel(range)} のリンクを解除しました。` : "CWS Agent: 対象範囲にはリンクがありません。");

  return {
    type: "clearHyperlink",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    cleared,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSetNoteOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではメモを変更できません。")) {
    throw new Error("保護されたシートではメモを変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);
  const before = captureCellsForHistory(sheet, range);

  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const cell = ensureCellModel(sheet, row, col);
    cell.note = op.note;
    if (op.noteVisible !== null) {
      cell.noteVisible = Boolean(op.noteVisible);
    } else {
      cell.noteVisible = Boolean(cell.noteVisible);
    }
    storeCellModelForSheet(sheet, key, cell);
  });
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent メモ`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} にメモを設定しました。`);

  return {
    type: "setNote",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentClearNoteOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではメモを変更できません。")) {
    throw new Error("保護されたシートではメモを変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const before = captureCellsForHistory(sheet, range);
  let cleared = 0;
  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const cell = sheet.cells[key];
    if (cellHasNote(cell) || cell?.noteVisible) {
      cell.note = null;
      cell.noteVisible = false;
      cleared += 1;
      storeCellModelForSheet(sheet, key, cell);
      removeCellModelIfEmpty(sheet, row, col);
    }
  });
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent メモ削除`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(cleared ? `CWS Agent: ${sheet.name}!${rangeToLabel(range)} のメモを削除しました。` : "CWS Agent: 対象範囲にはメモがありません。");

  return {
    type: "clearNote",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    cleared,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSetCommentOp(op, append = false) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではコメントを変更できません。")) {
    throw new Error("保護されたシートではコメントを変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);
  const before = captureCellsForHistory(sheet, range);

  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const cell = ensureCellModel(sheet, row, col);
    cell.comment = append ? appendCommentReply(cell.comment, op.comment) : makeCommentFromText(op.comment, cell.comment);
    storeCellModelForSheet(sheet, key, cell);
  });
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  closeCommentThreadPopup();
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent コメント${append ? "返信" : ""}`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} にコメントを${append ? "追加" : "設定"}しました。`);

  return {
    type: append ? "appendComment" : "setComment",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentClearCommentOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではコメントを変更できません。")) {
    throw new Error("保護されたシートではコメントを変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const before = captureCellsForHistory(sheet, range);
  let cleared = 0;
  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const cell = sheet.cells[key];
    if (cellHasComment(cell)) {
      delete cell.comment;
      cleared += 1;
      storeCellModelForSheet(sheet, key, cell);
      removeCellModelIfEmpty(sheet, row, col);
    }
  });
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  closeCommentThreadPopup();
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent コメント削除`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(cleared ? `CWS Agent: ${sheet.name}!${rangeToLabel(range)} のコメントを削除しました。` : "CWS Agent: 対象範囲にはコメントがありません。");

  return {
    type: "clearComment",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    cleared,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSetNamedRangeOp(op) {
  const before = cloneDefinedNames(state.model.definedNames || []);
  const definition = {
    name: op.name,
    scope: op.scope === "sheet" ? "sheet" : "workbook",
    sheetIndex: op.scope === "sheet" ? op.scopeSheetIndex : null,
    refersTo: op.refersTo,
    comment: op.comment || "",
    hidden: false,
  };
  const existingIndex = (state.model.definedNames || []).findIndex((item) => definedNameKey(item) === definedNameKey(definition));
  if (definedNameHasConflict(definition, existingIndex)) {
    throw new Error(definedNameDuplicateMessage(op.name));
  }
  state.model.definedNames = upsertDefinedName(state.model.definedNames || [], definition);
  const after = cloneDefinedNames(state.model.definedNames || []);
  state.hf = buildFormulaEngine(state.model);

  const selection = activeSelectionRange();
  const selectionRanges = activeSelectionRanges().map((range) => ({ ...range }));
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    definedNamesBefore: before,
    definedNamesAfter: after,
    selectionBefore: selection,
    selectionAfter: selection,
    selectionBeforeRanges: selectionRanges,
    selectionAfterRanges: selectionRanges,
    label: `${op.name} CWS Agent 名前の定義`,
  });
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: 名前 ${op.name} を定義しました。`);

  return {
    type: "setNamedRange",
    name: op.name,
    scope: definition.scope,
    scopeSheetIndex: definition.sheetIndex,
    refersTo: definition.refersTo,
    address: op.name,
  };
}



function applyCwsAgentClearNamedRangeOp(op) {
  const before = cloneDefinedNames(state.model.definedNames || []);
  const normalizedName = normalizeDefinedNameLookup(op.name);
  const next = (state.model.definedNames || []).filter((definition) => {
    if (definition.hidden) return true;
    if (normalizeDefinedNameLookup(definition.name) !== normalizedName) return true;
    if (op.scope === "any") return false;
    if (op.scope === "sheet") {
      return !(definition.scope === "sheet" && definition.sheetIndex === op.scopeSheetIndex);
    }
    return definition.scope === "sheet";
  });
  const cleared = before.length - next.length;
  if (!cleared) {
    setStatus(`CWS Agent: 名前 ${op.name} は見つかりませんでした。`);
    return {
      type: "clearNamedRange",
      name: op.name,
      scope: op.scope,
      scopeSheetIndex: op.scopeSheetIndex,
      cleared: 0,
      address: op.name,
    };
  }

  state.model.definedNames = next;
  const after = cloneDefinedNames(state.model.definedNames || []);
  state.hf = buildFormulaEngine(state.model);
  const selection = activeSelectionRange();
  const selectionRanges = activeSelectionRanges().map((range) => ({ ...range }));
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    definedNamesBefore: before,
    definedNamesAfter: after,
    selectionBefore: selection,
    selectionAfter: selection,
    selectionBeforeRanges: selectionRanges,
    selectionAfterRanges: selectionRanges,
    label: `${op.name} CWS Agent 名前の削除`,
  });
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: 名前 ${op.name} を削除しました。`);

  return {
    type: "clearNamedRange",
    name: op.name,
    scope: op.scope,
    scopeSheetIndex: op.scopeSheetIndex,
    cleared,
    address: op.name,
  };
}



function applyCwsAgentProtectSheetOp(op) {
  const sheet = state.model.sheets[op.sheetIndex];
  if (!sheet) {
    throw new Error("対象シートを開けませんでした。");
  }
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.protection = defaultSheetProtectionOptions(op.options);
  state.activeSheetIndex = op.sheetIndex;
  pushHistory({
    sheetIndex: op.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: `${sheet.name} CWS Agent シートの保護`,
  });
  renderWorkbook();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name} を保護しました。`);

  return {
    type: "protectSheet",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    address: sheet.name,
  };
}



function applyCwsAgentUnprotectSheetOp(op) {
  const sheet = state.model.sheets[op.sheetIndex];
  if (!sheet) {
    throw new Error("対象シートを開けませんでした。");
  }
  if (!isSheetProtected(sheet)) {
    setStatus(`CWS Agent: ${sheet.name} は保護されていません。`);
    return {
      type: "unprotectSheet",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      protected: false,
      address: sheet.name,
    };
  }
  const expected = String(sheet.protection?.password || "");
  if (expected && String(op.password || "") !== expected) {
    throw new Error("シート保護を解除するパスワードが正しくありません。");
  }

  const sheetBefore = captureSheetForHistory(sheet);
  delete sheet.protection;
  state.activeSheetIndex = op.sheetIndex;
  pushHistory({
    sheetIndex: op.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: `${sheet.name} CWS Agent シート保護の解除`,
  });
  renderWorkbook();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name} のシート保護を解除しました。`);

  return {
    type: "unprotectSheet",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    protected: false,
    address: sheet.name,
  };
}



function applyCwsAgentAddSheetOp(op) {
  const activeSheetBefore = state.activeSheetIndex;
  const insertAt = Math.max(0, Math.min(Number(op.index) || 0, state.model.sheets.length));
  const sheet = createBlankSheet(op.name);
  state.model.sheets.splice(insertAt, 0, sheet);
  state.activeSheetIndex = insertAt;
  clearSheetSelectionState();
  state.selected = { sheetIndex: insertAt, row: 1, col: 1 };
  state.selectionRange = { sheetIndex: insertAt, top: 1, left: 1, bottom: 1, right: 1 };
  state.selectionRanges = [state.selectionRange];
  state.selectionAnchor = state.selected;
  state.selectionStart = state.selected;
  refreshCwsAgentWorkbookAfterSheetStructureChange();
  pushSheetInsertHistory(insertAt, sheet, activeSheetBefore, insertAt, `${sheet.name} CWS Agent ワークシート追加`, {
    selectionAfter: state.selectionRange,
    selectionAfterRanges: [state.selectionRange],
  });
  setStatus(`CWS Agent: ワークシートを追加しました：${sheet.name}`);

  return {
    type: "addSheet",
    sheetIndex: insertAt,
    sheetName: sheet.name,
    address: sheet.name,
  };
}



function applyCwsAgentRenameSheetOp(op) {
  const sheet = state.model.sheets[op.sheetIndex];
  if (!sheet) {
    throw new Error("対象シートを開けませんでした。");
  }
  const oldName = sheet.name;
  if (oldName === op.newName) {
    setStatus("CWS Agent: ワークシート名は変更されていません。");
    return {
      type: "renameSheet",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      oldName,
      newName: sheet.name,
      address: sheet.name,
    };
  }

  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  sheet.name = op.newName;
  state.activeSheetIndex = op.sheetIndex;
  refreshCwsAgentWorkbookAfterSheetStructureChange();
  const selectionAfter = activeSelectionRange();
  pushHistory({
    sheetIndex: op.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore,
    selectionAfter,
    selectionBeforeRanges,
    selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
    label: `${oldName} CWS Agent 名前の変更`,
  });
  setStatus(`CWS Agent: ワークシート名を変更しました：${oldName} → ${sheet.name}`);

  return {
    type: "renameSheet",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    oldName,
    newName: sheet.name,
    address: sheet.name,
  };
}



function applyCwsAgentDuplicateSheetOp(op) {
  const source = state.model.sheets[op.sourceSheetIndex];
  if (!source) {
    throw new Error("コピー元シートを開けませんでした。");
  }

  const activeSheetBefore = state.activeSheetIndex;
  const insertAt = Math.max(0, Math.min(Number(op.index) || 0, state.model.sheets.length));
  const copy = cloneSheetModel(source);
  copy.name = op.name;
  copy.state = "visible";
  state.model.sheets.splice(insertAt, 0, copy);
  state.activeSheetIndex = insertAt;
  clearSheetSelectionState();
  refreshCwsAgentWorkbookAfterSheetStructureChange();
  pushSheetInsertHistory(insertAt, copy, activeSheetBefore, insertAt, `${source.name} CWS Agent ワークシートコピー`, {
    selectionAfter: null,
    selectionAfterRanges: [],
  });
  setStatus(`CWS Agent: ワークシートをコピーしました：${copy.name}`);

  return {
    type: "duplicateSheet",
    sourceSheetIndex: op.sourceSheetIndex,
    sheetIndex: insertAt,
    sheetName: copy.name,
    sourceSheetName: source.name,
    address: copy.name,
  };
}



function applyCwsAgentMoveSheetOp(op) {
  const source = state.model.sheets[op.sourceSheetIndex];
  if (!source) {
    throw new Error("移動元シートを開けませんでした。");
  }

  const sheetCount = state.model.sheets.length;
  const normalizedBefore = Math.max(0, Math.min(Number(op.beforeIndex), sheetCount));
  let targetIndex = normalizedBefore;
  if (normalizedBefore > op.sourceSheetIndex) targetIndex -= 1;
  targetIndex = Math.max(0, Math.min(targetIndex, sheetCount - 1));
  if (targetIndex === op.sourceSheetIndex) {
    state.activeSheetIndex = op.sourceSheetIndex;
    setStatus("CWS Agent: ワークシートは移動されていません。");
    return {
      type: "moveSheet",
      sourceSheetIndex: op.sourceSheetIndex,
      sheetIndex: op.sourceSheetIndex,
      sheetName: source.name,
      beforeIndex: normalizedBefore,
      moved: false,
      address: source.name,
    };
  }

  const activeSheetBefore = state.activeSheetIndex;
  const movedName = source.name;
  const actualIndex = moveSheetToIndex(op.sourceSheetIndex, targetIndex);
  state.activeSheetIndex = actualIndex;
  clearSheetSelectionState();
  refreshCwsAgentWorkbookAfterSheetStructureChange();
  pushSheetMoveHistory(op.sourceSheetIndex, actualIndex, activeSheetBefore, actualIndex, `${movedName} CWS Agent ワークシート移動`);
  setStatus(`CWS Agent: ワークシートを移動しました：${movedName}`);

  return {
    type: "moveSheet",
    sourceSheetIndex: op.sourceSheetIndex,
    sheetIndex: actualIndex,
    sheetName: movedName,
    beforeIndex: normalizedBefore,
    moved: true,
    address: movedName,
  };
}



function applyCwsAgentSheetTabColorOp(op) {
  const sheet = state.model.sheets[op.sheetIndex];
  if (!sheet) {
    throw new Error("対象シートを開けませんでした。");
  }
  const previousColor = sheet.tabColor || "";
  setSheetTabColor(op.sheetIndex, op.color || "");
  return {
    type: op.color ? "setSheetTabColor" : "clearSheetTabColor",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    oldColor: previousColor,
    color: sheet.tabColor || "",
    address: sheet.name,
  };
}



function applyCwsAgentActivateSheetOp(op) {
  const sheet = state.model.sheets[op.sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error("対象シートを開けませんでした。");
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }
  setStatus(`CWS Agent: ${sheet.name} に移動しました。`);
  return {
    type: "activateSheet",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    address: sheet.name,
  };
}



function applyCwsAgentSelectRangeOp(op) {
  const sheet = state.model.sheets[op.sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error("対象シートを開けませんでした。");
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  if (op.fullRow) range.fullRow = true;
  if (op.fullColumn) range.fullColumn = true;
  const point = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selected = point;
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = point;
  state.selectionAnchor = point;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  state.formulaEdit = null;
  const didScroll = scrollRangeIntoView(range);
  if (didScroll) {
    renderSheet(null, { deferSelectionUiSync: true });
  }
  updateSelectionUi();
  updateFormulaBarForSelection();
  scheduleSelectionUiSync();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} を選択しました。`);
  return {
    type: "selectRange",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    fullRow: range.fullRow === true,
    fullColumn: range.fullColumn === true,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function cwsAgentCaptureRangeClipboard(sheetIndex, sheet, range, mode = "copy") {
  const cellEntries = [];
  const merges = (sheet.merges || [])
    .filter((merge) => isRangeInsideRange(merge, range))
    .map((merge) => ({
      top: merge.top - range.top + 1,
      left: merge.left - range.left + 1,
      bottom: merge.bottom - range.top + 1,
      right: merge.right - range.left + 1,
    }));
  Object.values(sheet.cells || {}).forEach((sourceCell) => {
    if (!sourceCell || !isCellInRange(range, sourceCell.row, sourceCell.col)) return;
    const cell = cloneCellForClipboard(sourceCell);
    if (!cell) return;
    cell.row = sourceCell.row - range.top + 1;
    cell.col = sourceCell.col - range.left + 1;
    cellEntries.push(cell);
  });
  cellEntries.sort((left, right) => left.row - right.row || left.col - right.col);
  return {
    text: "",
    mode,
    sourceSheetIndex: sheetIndex,
    sourceRange: { ...range },
    rowCount: range.bottom - range.top + 1,
    colCount: range.right - range.left + 1,
    rows: [],
    columns: [],
    merges,
    cellEntries,
  };
}



function applyCwsAgentCopyMoveRangeOp(op) {
  const sourceSheet = state.model.sheets[op.sourceSheetIndex];
  const targetSheet = state.model.sheets[op.targetSheetIndex];
  if (!sourceSheet || !isSheetVisible(sourceSheet)) {
    throw new Error("コピー元シートを開けませんでした。");
  }
  if (!targetSheet || !isSheetVisible(targetSheet)) {
    throw new Error("コピー先シートを開けませんでした。");
  }

  const sourceRange = normalizeRange(
    { sheetIndex: op.sourceSheetIndex, row: op.sourceTop, col: op.sourceLeft },
    { sheetIndex: op.sourceSheetIndex, row: op.sourceBottom, col: op.sourceRight },
  );
  const targetRange = normalizeRange(
    { sheetIndex: op.targetSheetIndex, row: op.targetTop, col: op.targetLeft },
    { sheetIndex: op.targetSheetIndex, row: op.targetBottom, col: op.targetRight },
  );
  if (op.type === "moveRange" && op.sourceSheetIndex === op.targetSheetIndex && rangesIntersect(sourceRange, targetRange)) {
    throw new Error("移動元と移動先が重なっているため、Agent 操作を適用しません。");
  }
  if (op.type === "moveRange" && !ensureEditableRangeForProtection(sourceSheet, sourceRange, "移動")) {
    throw new Error("保護されたセルは移動できません。");
  }
  if (!ensureEditableRangeForProtection(targetSheet, targetRange, "貼り付け")) {
    throw new Error("保護されたセルには貼り付けできません。");
  }

  const affectedSheetIndexes = [...new Set([op.sourceSheetIndex, op.targetSheetIndex])];
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  const activeSheetBefore = state.activeSheetIndex;
  const clipboard = cwsAgentCaptureRangeClipboard(op.sourceSheetIndex, sourceSheet, sourceRange, op.type === "moveRange" ? "cut" : "copy");
  const entriesByOffset = new Map(clipboardCellEntries(clipboard).map((sourceCell) => [cellKey(sourceCell.row, sourceCell.col), sourceCell]));
  const tableEditCandidate = tableRangeEditMayChangeMetadata(targetSheet, targetRange, {
    assumeNonEmpty: entriesByOffset.size > 0,
    assumeFormula: clipboardHasFormula(clipboard),
  });

  try {
    ensureSheetSize(targetSheet, targetRange.bottom, targetRange.right);
    removeIntersectingMerges(targetSheet, targetRange);
    clearCellsInRange(targetSheet, targetRange);
    for (let rowOffset = 1; rowOffset <= clipboard.rowCount; rowOffset += 1) {
      for (let colOffset = 1; colOffset <= clipboard.colCount; colOffset += 1) {
        const sourceOffset = { row: rowOffset, col: colOffset };
        const sourceCell = entriesByOffset.get(cellKey(rowOffset, colOffset));
        const row = targetRange.top + rowOffset - 1;
        const col = targetRange.left + colOffset - 1;
        if (op.keepSourceFormatting === false) {
          writeCellContent(targetSheet, getSheetId(targetSheet.name), row, col, pasteValueFromClipboardCell(sourceCell, {
            clipboard,
            targetRow: row,
            targetCol: col,
            sourceOffset,
          }));
        } else {
          writeClonedClipboardCellContent(targetSheet, row, col, clipboard, sourceCell, sourceOffset);
        }
      }
    }
    if (op.keepSourceFormatting !== false) {
      clipboard.merges?.forEach((merge) => {
        targetSheet.merges.push({
          top: targetRange.top + merge.top - 1,
          left: targetRange.left + merge.left - 1,
          bottom: targetRange.top + merge.bottom - 1,
          right: targetRange.left + merge.right - 1,
        });
      });
    }
    if (op.type === "moveRange") {
      removeIntersectingMerges(sourceSheet, sourceRange);
      clearCellsInRange(sourceSheet, sourceRange);
    }
    const tableEdit = tableEditCandidate ? applyTableRangeEditBehaviors(targetSheet, targetRange) : { changed: false, status: "" };
    state.hf = buildFormulaEngine(state.model);
    if (state.activeSheetIndex !== op.targetSheetIndex) {
      state.activeSheetIndex = op.targetSheetIndex;
      state.showFormulas = targetSheet.showFormulas === true;
      state.selectedSheetIndexes = [];
    }
    state.selected = { sheetIndex: op.targetSheetIndex, row: targetRange.top, col: targetRange.left };
    state.selectionRange = expandRangeForMerges(targetSheet, targetRange);
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
    state.selectedImage = null;
    state.selectedImages = [];
    state.formulaEdit = null;
    pushHistory({
      sheetIndex: op.targetSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(affectedSheetIndexes),
      activeSheetBefore,
      activeSheetAfter: op.targetSheetIndex,
      selectionBefore: sourceRange,
      selectionAfter: state.selectionRange,
      selectionBeforeRanges: [sourceRange],
      selectionAfterRanges: [state.selectionRange],
      label: `${rangeToLabel(sourceRange)} -> ${rangeToLabel(targetRange)} CWS Agent ${op.type === "moveRange" ? "移動" : "コピー"}`,
    });
    renderWorkbook();
    updateSelectionUi();
    updateFormulaBarForSelection();
    const actionLabel = op.type === "moveRange" ? "移動" : "コピー";
    const status = `CWS Agent: ${sourceSheet.name}!${rangeToLabel(sourceRange)} を ${targetSheet.name}!${rangeToLabel(targetRange)} に${actionLabel}しました。`;
    setStatus(tableEdit.changed ? `${tableEdit.status} ${status}` : status);
    return {
      type: op.type,
      sourceSheetIndex: op.sourceSheetIndex,
      sourceSheetName: sourceSheet.name,
      targetSheetIndex: op.targetSheetIndex,
      targetSheetName: targetSheet.name,
      sourceRange: rangeToLabel(sourceRange),
      targetRange: rangeToLabel(targetRange),
      address: `${targetSheet.name}!${rangeToLabel(targetRange)}`,
    };
  } catch (error) {
    restoreSheetsFromHistory(sheetsBefore);
    state.activeSheetIndex = Math.max(0, Math.min(activeSheetBefore, state.model.sheets.length - 1));
    state.hf = buildFormulaEngine(state.model);
    renderWorkbook();
    updateFormulaBarForSelection();
    throw error;
  }
}



function applyCwsAgentDeleteSheetOp(op) {
  const sheet = state.model.sheets[op.sheetIndex];
  if (!sheet) {
    throw new Error("対象シートを開けませんでした。");
  }
  if (state.model.sheets.length <= 1 || (isSheetVisible(sheet) && visibleSheetIndexes().length <= 1)) {
    throw new Error("少なくとも 1 つの表示ワークシートを残す必要があります。");
  }

  const activeSheetBefore = state.activeSheetIndex;
  const activeName = state.model.sheets[activeSheetBefore]?.name || "";
  const [removed] = state.model.sheets.splice(op.sheetIndex, 1);
  const activeByName = state.model.sheets.findIndex((candidate) => candidate.name === activeName);
  state.activeSheetIndex = activeByName >= 0 ? activeByName : Math.min(op.sheetIndex, state.model.sheets.length - 1);
  ensureActiveSheetVisible();
  const activeSheetAfter = state.activeSheetIndex;
  clearSheetSelectionState();
  refreshCwsAgentWorkbookAfterSheetStructureChange();
  pushSheetDeleteHistory(op.sheetIndex, removed, activeSheetBefore, activeSheetAfter, `${removed.name} CWS Agent ワークシート削除`, {
    selectionAfter: null,
    selectionAfterRanges: [],
  });
  setStatus(`CWS Agent: ワークシートを削除しました：${removed.name}`);

  return {
    type: "deleteSheet",
    sheetIndex: op.sheetIndex,
    sheetName: removed.name,
    address: removed.name,
  };
}



function applyCwsAgentSheetVisibilityOp(op, hidden) {
  const sheet = state.model.sheets[op.sheetIndex];
  if (!sheet) {
    throw new Error("対象シートを開けませんでした。");
  }
  if (hidden && isSheetVisible(sheet) && visibleSheetIndexes().length <= 1) {
    throw new Error("少なくとも 1 つの表示ワークシートを残す必要があります。");
  }
  if (hidden === !isSheetVisible(sheet)) {
    setStatus(`CWS Agent: ワークシートはすでに${hidden ? "非表示" : "表示"}です：${sheet.name}`);
    return {
      type: hidden ? "hideSheet" : "unhideSheet",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      address: sheet.name,
    };
  }

  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  sheet.state = hidden ? "hidden" : "visible";
  state.activeSheetIndex = hidden ? nearestVisibleSheetIndex(op.sheetIndex) : op.sheetIndex;
  clearSheetSelectionState();
  refreshCwsAgentWorkbookAfterSheetStructureChange();
  pushHistory({
    sheetIndex: op.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore,
    selectionAfter: null,
    selectionBeforeRanges,
    selectionAfterRanges: [],
    label: `${sheet.name} CWS Agent ${hidden ? "非表示" : "再表示"}`,
  });
  setStatus(`CWS Agent: ワークシートを${hidden ? "非表示" : "再表示"}にしました：${sheet.name}`);

  return {
    type: hidden ? "hideSheet" : "unhideSheet",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    address: sheet.name,
  };
}



function applyCwsAgentSetFreezePanesOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const pane = normalizePaneState({ rows: op.rows, columns: op.columns });
  if (!pane) {
    throw new Error("固定する行または列を指定してください。");
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.freezePanes = pane;
  sheet.splitPane = null;
  const address = `${sheet.name}!${columnName(pane.col)}${pane.row}`;
  commitCwsAgentSheetViewChange(
    sheet,
    sheetBefore,
    `CWS Agent: ${cwsAgentFreezePanesStatus(pane)}固定しました。`,
    `${sheet.name} CWS Agent ウィンドウ枠の固定`,
    workbookViewBefore,
  );

  return {
    type: "setFreezePanes",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    rows: pane.rows,
    columns: pane.columns,
    address,
  };
}



function applyCwsAgentClearFreezePanesOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const currentPane = normalizePaneState(sheet.freezePanes);
  const address = `${sheet.name}!A1`;
  if (!currentPane) {
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus("CWS Agent: 固定されたウィンドウ枠はありません。");
    return {
      type: "clearFreezePanes",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      address,
    };
  }

  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.freezePanes = null;
  commitCwsAgentSheetViewChange(
    sheet,
    sheetBefore,
    "CWS Agent: ウィンドウ枠の固定を解除しました。",
    `${sheet.name} CWS Agent ウィンドウ枠の固定解除`,
    workbookViewBefore,
  );

  return {
    type: "clearFreezePanes",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    address,
  };
}



function applyCwsAgentSetSheetViewOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const patch = op.patch || {};
  const current = cwsAgentSheetViewSnapshot(sheet);
  const next = { ...current, ...patch };
  if (Object.prototype.hasOwnProperty.call(patch, "viewMode") && !Object.prototype.hasOwnProperty.call(patch, "showRuler")) {
    next.showRuler = patch.viewMode === "page-layout";
  }
  if (Object.prototype.hasOwnProperty.call(patch, "showWhiteSpace") && patch.showWhiteSpace && next.viewMode === "normal") {
    next.viewMode = "page-layout";
  }
  if (Object.prototype.hasOwnProperty.call(patch, "showRuler") && patch.showRuler && next.viewMode === "normal") {
    next.viewMode = "page-layout";
  }
  next.showRuler = normalizedSheetShowRuler(next.showRuler, next.viewMode);

  if (!cwsAgentSheetViewHasChanges(current, next, patch)) {
    renderWorkbook();
    updateFormulaBarForSelection();
    syncViewRibbonButtons();
    setStatus("CWS Agent: 表示設定は変更されていません。");
    return {
      type: "setSheetView",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      patch: { ...patch },
      address: sheet.name,
    };
  }

  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.viewMode = next.viewMode;
  if (Object.prototype.hasOwnProperty.call(patch, "zoomScale")) {
    applySheetZoomWithoutHistory(sheet, next.zoomScale);
  }
  sheet.showGridLines = next.showGridLines !== false;
  sheet.showHeadings = next.showHeadings !== false;
  sheet.showRuler = normalizedSheetShowRuler(next.showRuler, sheet.viewMode);
  sheet.showFormulas = next.showFormulas === true;
  sheet.showZeros = next.showZeros !== false;
  sheet.showWhiteSpace = next.showWhiteSpace !== false;
  sheet.showOutlineSymbols = next.showOutlineSymbols !== false;
  sheet.showObjects = next.showObjects !== false;
  sheet.rightToLeft = next.rightToLeft === true;
  state.showFormulas = sheet.showFormulas === true;

  const labels = cwsAgentSheetViewPatchLabels(patch, sheet);
  commitCwsAgentSheetViewChange(
    sheet,
    sheetBefore,
    `CWS Agent: 表示設定を更新しました${labels ? `（${labels}）` : ""}。`,
    `${sheet.name} CWS Agent 表示設定`,
    workbookViewBefore,
  );

  return {
    type: "setSheetView",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    patch: { ...patch },
    address: sheet.name,
  };
}



function applyCwsAgentSetPrintAreaOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const range = { sheetIndex: op.sheetIndex, top: op.top, left: op.left, bottom: op.bottom, right: op.right };
  const label = rangeToLabel(range);
  commitPageLayoutChange((layout) => {
    layout.printArea = { range: { ...range }, label };
  }, `${sheet.name} CWS Agent 印刷範囲`, `CWS Agent: ${sheet.name}!${label} を印刷範囲に設定しました。`);

  return {
    type: "setPrintArea",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    address: `${sheet.name}!${label}`,
  };
}



function applyCwsAgentClearPrintAreaOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const current = ensurePageLayoutSettings(sheet).printArea?.range || null;
  if (!current) {
    renderWorkbook();
    setStatus("CWS Agent: 印刷範囲は設定されていません。");
    return {
      type: "clearPrintArea",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      address: sheet.name,
    };
  }

  commitPageLayoutChange((layout) => {
    layout.printArea = null;
  }, `${sheet.name} CWS Agent 印刷範囲のクリア`, "CWS Agent: 印刷範囲をクリアしました。");

  return {
    type: "clearPrintArea",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    address: `${sheet.name}!${rangeToLabel(current)}`,
  };
}



function applyCwsAgentSetPrintTitlesOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const rowLabel = op.rows ? rangeToLabel(op.rows) : "";
  const colLabel = op.cols ? rangeToLabel(op.cols) : "";
  const summary = [rowLabel ? `タイトル行 ${rowLabel}` : "", colLabel ? `タイトル列 ${colLabel}` : ""].filter(Boolean).join("、");
  commitPageLayoutChange((layout) => {
    layout.printTitles = layout.printTitles || { rows: null, cols: null };
    if (op.rows) layout.printTitles.rows = { ...op.rows };
    if (op.cols) layout.printTitles.cols = { ...op.cols };
  }, `${sheet.name} CWS Agent 印刷タイトル`, `CWS Agent: ${summary} を印刷タイトルに設定しました。`);

  return {
    type: "setPrintTitles",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    rows: rowLabel,
    cols: colLabel,
    address: `${sheet.name}!${rowLabel || colLabel}`,
  };
}



function applyCwsAgentClearPrintTitlesOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const layout = ensurePageLayoutSettings(sheet);
  const currentRows = layout.printTitles?.rows || null;
  const currentCols = layout.printTitles?.cols || null;
  if (!((op.rows && currentRows) || (op.cols && currentCols))) {
    renderWorkbook();
    setStatus("CWS Agent: 印刷タイトルは設定されていません。");
    return {
      type: "clearPrintTitles",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      address: sheet.name,
    };
  }

  commitPageLayoutChange((layout) => {
    layout.printTitles = layout.printTitles || { rows: null, cols: null };
    if (op.rows) layout.printTitles.rows = null;
    if (op.cols) layout.printTitles.cols = null;
  }, `${sheet.name} CWS Agent 印刷タイトルのクリア`, "CWS Agent: 印刷タイトルをクリアしました。");

  const label = [currentRows ? rangeToLabel(currentRows) : "", currentCols ? rangeToLabel(currentCols) : ""].filter(Boolean).join(",");
  return {
    type: "clearPrintTitles",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    rows: op.rows,
    cols: op.cols,
    address: label ? `${sheet.name}!${label}` : sheet.name,
  };
}



function applyCwsAgentSetHeaderFooterOp(op) {
  const sheet = activateCwsAgentPageSetupSheet(op);
  const labels = cwsAgentHeaderFooterPatchLabels(op.settings);
  commitPageLayoutChange((layout, sheet) => {
    const next = cwsAgentMergedHeaderFooterSettings(sheet, layout, op.settings);
    sheet.headerFooter = normalizedHeaderFooterSettings(next, layout);
    layout.headerFooter = {
      header: headerFooterPlainText(sheet.headerFooter.oddHeader),
      footer: headerFooterPlainText(sheet.headerFooter.oddFooter),
    };
  }, `${sheet.name} CWS Agent ヘッダー/フッター`, `CWS Agent: ヘッダー/フッターを更新しました${labels ? `（${labels}）` : ""}。`);
  return {
    type: "setHeaderFooter",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    settings: cloneJsonLikeValue(op.settings),
    address: sheet.name,
  };
}



function applyCwsAgentClearHeaderFooterOp(op) {
  const sheet = activateCwsAgentPageSetupSheet(op);
  const layout = ensurePageLayoutSettings(sheet);
  const current = normalizedHeaderFooterSettings(sheet.headerFooter, layout);
  if (!cwsAgentHeaderFooterHasTargetContent(current, op.target)) {
    renderWorkbook();
    setStatus("CWS Agent: クリア対象のヘッダー/フッターは設定されていません。");
    return {
      type: "clearHeaderFooter",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      target: op.target,
      address: sheet.name,
    };
  }
  const targetLabel = op.target === "header" ? "ヘッダー" : op.target === "footer" ? "フッター" : "ヘッダー/フッター";
  commitPageLayoutChange((layout, sheet) => {
    const next = normalizedHeaderFooterSettings(sheet.headerFooter, layout);
    cwsAgentClearHeaderFooterTarget(next, op.target);
    sheet.headerFooter = normalizedHeaderFooterSettings(next, layout);
    layout.headerFooter = {
      header: headerFooterPlainText(sheet.headerFooter.oddHeader),
      footer: headerFooterPlainText(sheet.headerFooter.oddFooter),
    };
  }, `${sheet.name} CWS Agent ${targetLabel}のクリア`, `CWS Agent: ${targetLabel}をクリアしました。`);
  return {
    type: "clearHeaderFooter",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    target: op.target,
    address: sheet.name,
  };
}



function cwsAgentMergedHeaderFooterSettings(sheet, layout, patch = {}) {
  const next = normalizedHeaderFooterSettings(sheet.headerFooter, layout);
  for (const key of ["oddHeader", "oddFooter", "firstHeader", "firstFooter", "evenHeader", "evenFooter"]) {
    if (patch[key]) next[key] = { ...next[key], ...patch[key] };
  }
  for (const key of ["differentFirst", "differentOddEven", "scaleWithDoc", "alignWithMargins"]) {
    if (Object.prototype.hasOwnProperty.call(patch, key)) next[key] = patch[key];
  }
  return next;
}



function cwsAgentHeaderFooterPatchLabels(patch = {}) {
  const labels = [];
  if (patch.oddHeader) labels.push("ヘッダー");
  if (patch.oddFooter) labels.push("フッター");
  if (patch.firstHeader || patch.firstFooter) labels.push("先頭ページ");
  if (patch.evenHeader || patch.evenFooter) labels.push("偶数ページ");
  if (Object.prototype.hasOwnProperty.call(patch, "differentFirst")) labels.push(`先頭ページ別指定${patch.differentFirst ? "オン" : "オフ"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "differentOddEven")) labels.push(`奇数/偶数別指定${patch.differentOddEven ? "オン" : "オフ"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "scaleWithDoc")) labels.push(`拡大縮小連動${patch.scaleWithDoc ? "オン" : "オフ"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "alignWithMargins")) labels.push(`余白に合わせる${patch.alignWithMargins ? "オン" : "オフ"}`);
  return labels.join("、");
}



function cwsAgentHeaderFooterHasTargetContent(settings, target) {
  const normalized = normalizedHeaderFooterSettings(settings);
  const hasHeader = ["oddHeader", "firstHeader", "evenHeader"].some((key) => headerFooterSectionHasText(normalized[key]));
  const hasFooter = ["oddFooter", "firstFooter", "evenFooter"].some((key) => headerFooterSectionHasText(normalized[key]));
  const hasFlags = normalized.differentFirst || normalized.differentOddEven || normalized.scaleWithDoc === false || normalized.alignWithMargins === false;
  if (target === "header") return hasHeader;
  if (target === "footer") return hasFooter;
  return hasHeader || hasFooter || hasFlags;
}



function cwsAgentClearHeaderFooterTarget(settings, target) {
  const empty = () => ({ left: "", center: "", right: "" });
  if (target === "header" || target === "all") {
    settings.oddHeader = empty();
    settings.firstHeader = empty();
    settings.evenHeader = empty();
  }
  if (target === "footer" || target === "all") {
    settings.oddFooter = empty();
    settings.firstFooter = empty();
    settings.evenFooter = empty();
  }
  if (target === "all") {
    settings.differentFirst = false;
    settings.differentOddEven = false;
    settings.scaleWithDoc = true;
    settings.alignWithMargins = true;
  }
}



function applyCwsAgentSetPageThemeOp(op) {
  const sheet = activateCwsAgentPageSetupSheet(op);
  const theme = PAGE_LAYOUT_THEMES.find((item) => item.value === op.theme) || PAGE_LAYOUT_THEMES[0];
  const effect = pageLayoutEffectPreset(op.theme);
  const includeColors = op.includeColors === true;
  const includeFonts = op.includeFonts === true;
  const includeEffects = op.includeEffects === true;
  const appliesFullTheme = includeColors && includeFonts && includeEffects;
  const label = appliesFullTheme
    ? `テーマ ${theme.label}`
    : includeColors
      ? `配色 ${theme.label}`
      : includeFonts
        ? `フォント ${theme.label}`
        : includeEffects
          ? `効果 ${effect.label}`
          : `テーマ ${theme.label}`;
  const status = appliesFullTheme
    ? `CWS Agent: ${theme.label} テーマを適用しました。`
    : includeColors
      ? `CWS Agent: 配色を「${theme.label}」に変更しました。`
      : includeFonts
        ? `CWS Agent: フォントを「${theme.label}」に変更しました。`
        : includeEffects
          ? `CWS Agent: 効果を「${effect.label}」に変更しました。`
          : `CWS Agent: ${theme.label} テーマを適用しました。`;

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
  }, `${sheet.name} CWS Agent ${label}`, status);
  syncShapeFormatControls();
  return {
    type: "setPageTheme",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    theme: op.theme,
    includeColors,
    includeFonts,
    includeEffects,
    address: sheet.name,
  };
}



function applyCwsAgentSetPageOrientationOp(op) {
  const sheet = activateCwsAgentPageSetupSheet(op);
  commitPageLayoutChange((layout) => {
    layout.orientation = op.orientation;
  }, `${sheet.name} CWS Agent 印刷の向き`, `CWS Agent: 印刷の向きを${op.orientation === "landscape" ? "横" : "縦"}に設定しました。`);
  return {
    type: "setPageOrientation",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    orientation: op.orientation,
    address: sheet.name,
  };
}



function applyCwsAgentSetPaperSizeOp(op) {
  const sheet = activateCwsAgentPageSetupSheet(op);
  const paper = PAGE_LAYOUT_PAPER_SIZES.find((item) => item.value === op.paperSize) || PAGE_LAYOUT_PAPER_SIZES[0];
  commitPageLayoutChange((layout) => {
    layout.paperSize = paper.value;
  }, `${sheet.name} CWS Agent 用紙サイズ`, `CWS Agent: 用紙サイズを ${paper.label} に設定しました。`);
  return {
    type: "setPaperSize",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    paperSize: paper.value,
    address: sheet.name,
  };
}



function applyCwsAgentSetPageMarginsOp(op) {
  const sheet = activateCwsAgentPageSetupSheet(op);
  const label = op.marginPreset ? PAGE_LAYOUT_MARGIN_PRESETS[op.marginPreset]?.label || op.marginPreset : "ユーザー設定";
  commitPageLayoutChange((layout) => {
    if (op.marginPreset) layout.marginPreset = op.marginPreset;
    const base = op.marginPreset ? PAGE_LAYOUT_MARGIN_PRESETS[op.marginPreset] || PAGE_LAYOUT_MARGIN_PRESETS.normal : layout.margins || PAGE_LAYOUT_MARGIN_PRESETS.normal;
    layout.margins = { ...base, ...op.margins };
    delete layout.margins.label;
  }, `${sheet.name} CWS Agent 余白`, `CWS Agent: 余白を「${label}」に設定しました。`);
  return {
    type: "setPageMargins",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    marginPreset: op.marginPreset || "",
    margins: { ...op.margins },
    address: sheet.name,
  };
}



function applyCwsAgentSetPageScaleOp(op) {
  const sheet = activateCwsAgentPageSetupSheet(op);
  const status = op.fitToPage
    ? `CWS Agent: 印刷を 横 ${op.fitToWidth} / 縦 ${op.fitToHeight} ページに合わせます。`
    : `CWS Agent: 印刷倍率を ${op.scale}% に設定しました。`;
  commitPageLayoutChange((layout) => {
    layout.fitToPage = Boolean(op.fitToPage);
    layout.fitToWidth = op.fitToPage ? op.fitToWidth : "auto";
    layout.fitToHeight = op.fitToPage ? op.fitToHeight : "auto";
    layout.scale = op.fitToPage ? 100 : op.scale;
  }, `${sheet.name} CWS Agent 拡大縮小印刷`, status);
  return {
    type: "setPageScale",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    fitToPage: Boolean(op.fitToPage),
    fitToWidth: op.fitToWidth,
    fitToHeight: op.fitToHeight,
    scale: op.scale,
    address: sheet.name,
  };
}



function applyCwsAgentSetPrintOptionsOp(op) {
  const sheet = activateCwsAgentPageSetupSheet(op);
  const labels = cwsAgentPrintOptionsPatchLabels(op.patch);
  commitPageLayoutChange((layout) => {
    Object.assign(layout, op.patch);
  }, `${sheet.name} CWS Agent 印刷オプション`, `CWS Agent: 印刷オプションを更新しました${labels ? `（${labels}）` : ""}。`);
  return {
    type: "setPrintOptions",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    patch: { ...op.patch },
    address: sheet.name,
  };
}



function cwsAgentPrintOptionsPatchLabels(patch = {}) {
  const labels = [];
  if (Object.prototype.hasOwnProperty.call(patch, "printGridLines")) labels.push(`枠線${patch.printGridLines ? "印刷" : "非印刷"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "printHeadings")) labels.push(`行列番号${patch.printHeadings ? "印刷" : "非印刷"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "centerHorizontal")) labels.push(`水平中央${patch.centerHorizontal ? "オン" : "オフ"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "centerVertical")) labels.push(`垂直中央${patch.centerVertical ? "オン" : "オフ"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "blackAndWhite")) labels.push(`白黒${patch.blackAndWhite ? "オン" : "オフ"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "draftQuality")) labels.push(`簡易印刷${patch.draftQuality ? "オン" : "オフ"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "comments")) labels.push(`コメント:${patch.comments}`);
  if (Object.prototype.hasOwnProperty.call(patch, "cellErrors")) labels.push(`エラー:${patch.cellErrors}`);
  if (Object.prototype.hasOwnProperty.call(patch, "pageOrder")) labels.push(`ページ順:${patch.pageOrder}`);
  if (Object.prototype.hasOwnProperty.call(patch, "firstPageNumber")) labels.push(`先頭ページ:${patch.firstPageNumber}`);
  return labels.join("、");
}



function applyCwsAgentPageBreakOp(op, insert) {
  const sheet = activateCwsAgentPageSetupSheet(op);
  const key = op.axis === "row" ? "rows" : "cols";
  const label = cwsAgentPageBreakPositionLabel(op.axis, op.index);
  const layout = ensurePageLayoutSettings(sheet);
  const current = layout.pageBreaks?.[key] || [];
  if (insert && current.includes(op.index)) {
    renderWorkbook();
    setStatus(`CWS Agent: ${label}には既に手動改ページがあります。`);
    return {
      type: "insertPageBreak",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      axis: op.axis,
      index: op.index,
      address: `${sheet.name}!${label}`,
    };
  }
  if (!insert && !current.includes(op.index)) {
    renderWorkbook();
    setStatus(`CWS Agent: ${label}に解除できる手動改ページはありません。`);
    return {
      type: "removePageBreak",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      axis: op.axis,
      index: op.index,
      address: `${sheet.name}!${label}`,
    };
  }
  commitPageLayoutChange((layout) => {
    layout.pageBreaks = layout.pageBreaks || { rows: [], cols: [] };
    const values = layout.pageBreaks[key] || [];
    layout.pageBreaks[key] = insert
      ? [...new Set([...values, op.index])].sort((a, b) => a - b)
      : values.filter((value) => Number(value) !== op.index);
  }, `${sheet.name} CWS Agent 改ページ${insert ? "の挿入" : "の解除"}`, insert ? `CWS Agent: ${label}に改ページを挿入しました。` : `CWS Agent: ${label}の改ページを解除しました。`, {
    captureWorkbookView: true,
  });
  return {
    type: insert ? "insertPageBreak" : "removePageBreak",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    axis: op.axis,
    index: op.index,
    address: `${sheet.name}!${label}`,
  };
}

export {
  applyCwsAgentExistingShapeEditOp,
  applyCwsAgentActivateSheetOp,
  applyCwsAgentAddShapeOp,
  applyCwsAgentAddSheetOp,
  applyCwsAgentClearCommentOp,
  applyCwsAgentClearDataValidationOp,
  applyCwsAgentClearFreezePanesOp,
  applyCwsAgentClearHeaderFooterOp,
  applyCwsAgentClearHyperlinkOp,
  applyCwsAgentClearNamedRangeOp,
  applyCwsAgentClearNoteOp,
  applyCwsAgentClearPrintAreaOp,
  applyCwsAgentClearPrintTitlesOp,
  applyCwsAgentClearQuickAnalysisOp,
  applyCwsAgentClearSubtotalsOp,
  applyCwsAgentConsolidateOp,
  applyCwsAgentCopyMoveRangeOp,
  applyCwsAgentCreateChartOp,
  applyCwsAgentCreateForecastSheetOp,
  applyCwsAgentCreatePivotTableOp,
  applyCwsAgentCreateTableOp,
  applyCwsAgentCreateWorkbookTemplateOp,
  applyCwsAgentDataTableOp,
  applyCwsAgentDeleteSheetOp,
  applyCwsAgentDuplicateSheetOp,
  applyCwsAgentGoalSeekOp,
  applyCwsAgentMoveSheetOp,
  applyCwsAgentPageBreakOp,
  applyCwsAgentProtectSheetOp,
  applyCwsAgentRenameSheetOp,
  applyCwsAgentSelectRangeOp,
  applyCwsAgentSetCommentOp,
  applyCwsAgentSetDataValidationOp,
  applyCwsAgentSetFreezePanesOp,
  applyCwsAgentSetHeaderFooterOp,
  applyCwsAgentSetHyperlinkOp,
  applyCwsAgentSetNamedRangeOp,
  applyCwsAgentSetNoteOp,
  applyCwsAgentSetPageMarginsOp,
  applyCwsAgentSetPageOrientationOp,
  applyCwsAgentSetPageScaleOp,
  applyCwsAgentSetPageThemeOp,
  applyCwsAgentSetPaperSizeOp,
  applyCwsAgentSetPrintAreaOp,
  applyCwsAgentSetPrintOptionsOp,
  applyCwsAgentSetPrintTitlesOp,
  applyCwsAgentSetQuickAnalysisOp,
  applyCwsAgentSetQuickAnalysisTotalOp,
  applyCwsAgentSetShapeTextOp,
  applyCwsAgentSetSheetViewOp,
  applyCwsAgentSetSubtotalOp,
  applyCwsAgentSheetTabColorOp,
  applyCwsAgentSheetVisibilityOp,
  applyCwsAgentUnprotectSheetOp,
  cwsAgentApplyQuickAnalysisMode,
  cwsAgentBuildTemplateSheet,
  cwsAgentCaptureRangeClipboard,
  cwsAgentCellEndpointPoint,
  cwsAgentCellMatchesFindQuery,
  cwsAgentClearHeaderFooterTarget,
  cwsAgentConnectorBoxFromEndpoints,
  cwsAgentEnsureShapeGrid,
  cwsAgentFindShapeForConnector,
  cwsAgentHeaderFooterHasTargetContent,
  cwsAgentHeaderFooterPatchLabels,
  cwsAgentInferConnectorSites,
  cwsAgentMergedHeaderFooterSettings,
  cwsAgentPrintOptionsPatchLabels,
  cwsAgentQuickAnalysisApplicableCount,
  cwsAgentShapeConnectorBoxForOp,
  cwsAgentShapePlacementBoxForOp,
  cwsAgentTemplateCellBorders,
  cwsAgentTemplateCellCss,
};
