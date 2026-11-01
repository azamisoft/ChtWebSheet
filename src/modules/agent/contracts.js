// Feature module; shared bindings are imported explicitly.
import {
  CWS_AGENT_CONTRACT_RANGE_OPS,
  CWS_AGENT_MAX_COLUMN_WIDTH,
  CWS_AGENT_MAX_FORMAT_CELLS,
  CWS_AGENT_MAX_OPS,
  CWS_AGENT_MAX_ROW_HEIGHT,
  CWS_AGENT_MAX_WORKBOOK_TEMPLATE_CELLS,
  CWS_AGENT_MAX_WORKBOOK_TEMPLATE_SHEETS,
  CWS_AGENT_MIN_COLUMN_WIDTH,
  CWS_AGENT_MIN_ROW_HEIGHT,
  CWS_AGENT_PROGRESSIVE_PREVIEW_DELAY_MS,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  HyperFormula,
  WEBSHEET_FORMULA_TRANSLATIONS,
  WebSheetFormulaPlugin,
  cellStyleInternPools,
  cwsAgentOperationSources,
  state,
} from "../../app-core.js";
import {
  cwsAgentBorderPreset,
  cwsAgentCellValue,
  cwsAgentColumnValue,
  cwsAgentMatrixValues,
  cwsAgentNumberFormat,
  cwsAgentNumberFormatLabel,
  cwsAgentPlainObject,
  cwsAgentPositiveInteger,
  cwsAgentRangeForOp,
  cwsAgentSheetIndexForOp,
  cwsAgentSortKeyColumn,
  cwsAgentSortRangeForOp,
  cwsAgentSparklineAxis,
  cwsAgentSparklineSourceRangeForOp,
  cwsAgentSparklineType,
  cwsAgentStylePatch,
  normalizeCwsAgentActivateSheetOp,
  normalizeCwsAgentAddShapeOp,
  normalizeCwsAgentAddSheetOp,
  normalizeCwsAgentAdvancedFilterOp,
  normalizeCwsAgentAppendCommentOp,
  normalizeCwsAgentAutoFitColumnsOp,
  normalizeCwsAgentAutoFitRowsOp,
  normalizeCwsAgentClearAutoFilterOp,
  normalizeCwsAgentClearCommentOp,
  normalizeCwsAgentClearDataValidationOp,
  normalizeCwsAgentClearFilterCriteriaOp,
  normalizeCwsAgentClearFreezePanesOp,
  normalizeCwsAgentClearHeaderFooterOp,
  normalizeCwsAgentClearHyperlinkOp,
  normalizeCwsAgentClearNamedRangeOp,
  normalizeCwsAgentClearNoteOp,
  normalizeCwsAgentClearPrintAreaOp,
  normalizeCwsAgentClearPrintTitlesOp,
  normalizeCwsAgentClearQuickAnalysisOp,
  normalizeCwsAgentClearRangeOp,
  normalizeCwsAgentClearSubtotalsOp,
  normalizeCwsAgentColumnsOp,
  normalizeCwsAgentConsolidateOp,
  normalizeCwsAgentCopyMoveRangeOp,
  normalizeCwsAgentCreateChartOp,
  normalizeCwsAgentCreateForecastSheetOp,
  normalizeCwsAgentCreatePivotTableOp,
  normalizeCwsAgentCreateTableOp,
  normalizeCwsAgentDataTableOp,
  normalizeCwsAgentDeleteSheetOp,
  normalizeCwsAgentDuplicateSheetOp,
  normalizeCwsAgentFindReplaceOp,
  normalizeCwsAgentGoalSeekOp,
  normalizeCwsAgentHeaderHiddenOp,
  normalizeCwsAgentMergeCellsOp,
  normalizeCwsAgentMoveSheetOp,
  normalizeCwsAgentOutlineOp,
  normalizeCwsAgentPageBreakOp,
  normalizeCwsAgentProtectSheetOp,
  normalizeCwsAgentRemoveDuplicatesOp,
  normalizeCwsAgentRenameSheetOp,
  normalizeCwsAgentResetPageBreaksOp,
  normalizeCwsAgentRowsOp,
  normalizeCwsAgentSelectRangeOp,
  normalizeCwsAgentSetAutoFilterOp,
  normalizeCwsAgentSetColumnWidthOp,
  normalizeCwsAgentSetCommentOp,
  normalizeCwsAgentSetDataValidationOp,
  normalizeCwsAgentSetFilterCriteriaOp,
  normalizeCwsAgentSetFreezePanesOp,
  normalizeCwsAgentSetHeaderFooterOp,
  normalizeCwsAgentSetHyperlinkOp,
  normalizeCwsAgentSetNamedRangeOp,
  normalizeCwsAgentSetNoteOp,
  normalizeCwsAgentSetPageMarginsOp,
  normalizeCwsAgentSetPageOrientationOp,
  normalizeCwsAgentSetPageScaleOp,
  normalizeCwsAgentSetPageThemeOp,
  normalizeCwsAgentSetPaperSizeOp,
  normalizeCwsAgentSetPrintAreaOp,
  normalizeCwsAgentSetPrintOptionsOp,
  normalizeCwsAgentSetPrintTitlesOp,
  normalizeCwsAgentSetQuickAnalysisOp,
  normalizeCwsAgentSetQuickAnalysisTotalOp,
  normalizeCwsAgentSetRowHeightOp,
  normalizeCwsAgentSetShapeTextOp,
  normalizeCwsAgentExistingShapeEditOp,
  normalizeCwsAgentSetSheetViewOp,
  normalizeCwsAgentSetSubtotalOp,
  normalizeCwsAgentSheetTabColorOp,
  normalizeCwsAgentSheetVisibilityOp,
  normalizeCwsAgentTextToColumnsOp,
  normalizeCwsAgentUnmergeCellsOp,
  normalizeCwsAgentUnprotectSheetOp,
} from "./normalization.js";
import {
  applyCwsAgentOp,
  cwsAgentCellProtectionPatch,
  cwsAgentClampDimension,
  cwsAgentHeaderMode,
  cwsAgentParseCellAddress,
  cwsAgentParseRangeAddress,
  cwsAgentSortDirection,
} from "./operations.js";
import {
  uniqueSheetNameWithBase,
} from "../drawing/charts-and-insertion.js";
import {
  lineShapeSegmentsInSheetPixels,
} from "../drawing/transforms-and-connectors.js";
import {
  rangeToLabel,
  updateSelectionUi,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  imageBoxToPixels,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  applyHistoryEntry,
  captureSheetsForHistory,
  cloneDefinedNames,
  cloneJsonLikeValue,
  cwsAgentCaptureWorkbookSheetsSnapshot,
  cwsAgentRestoreWorkbookSheetsSnapshot,
  pushHistory,
  restoreSheetsFromHistory,
} from "../editing/structure-and-history.js";
import {
  setStatus,
} from "../ui/help-and-utilities.js";
import {
  buildFormulaEngine,
  isSheetVisible,
  renderWorkbook,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  beginCwsAgentPreviewPersistence,
  currentWorkbookSnapshot,
  persistenceAvailable,
  scheduleWorkbookPersistence,
  writePersistedWorkbookSnapshot,
} from "../workbook/import-model.js";
import {
  compactSheetStyleOnlyCellsToStyleRanges,
} from "../workbook/saving-and-export.js";


function registerWebSheetFormulaPlugin() {
  if (typeof HyperFormula.getFunctionPlugin === "function" && HyperFormula.getFunctionPlugin("TEXTJOIN")) return;
  try {
    HyperFormula.registerFunctionPlugin(WebSheetFormulaPlugin, WEBSHEET_FORMULA_TRANSLATIONS);
  } catch (error) {
    console.warn("WebSheet formula plugin registration skipped:", error);
  }
}



function resetCellStyleInternPools() {
  cellStyleInternPools.css.clear();
  cellStyleInternPools.borderSide.clear();
  cellStyleInternPools.borders.clear();
}



function applyCwsAgentOps(ops = [], options = {}) {
  if (!Array.isArray(ops)) {
    throw new Error("Agent 操作が正しくありません。");
  }
  if (!state.model) {
    throw new Error("ブックが開かれていません。");
  }
  if (ops.length > CWS_AGENT_MAX_OPS) {
    throw new Error(`一度に適用できる Agent 操作は${CWS_AGENT_MAX_OPS}件までです。`);
  }

  if (!options.keepPreview) {
    clearCwsAgentPreview({ restore: true, quiet: true });
  }
  const normalizedOps = normalizeCwsAgentOpsBatch(ops, options);
  const results = normalizedOps.map((op) => applyCwsAgentOp(op));
  validateCwsAgentAppliedBranchGeometry(normalizedOps, results);
  if (results.length) {
    setStatus(`CWS Agent: ${results.map((result) => result.address).join(", ")} を更新しました。`);
  }
  return { ok: true, applied: results.length, results };
}



function previewCwsAgentOps(ops = [], options = {}) {
  if (!Array.isArray(ops)) {
    throw new Error("Agent 操作が正しくありません。");
  }
  if (!state.model) {
    throw new Error("ブックが開かれていません。");
  }
  if (ops.length > CWS_AGENT_MAX_OPS) {
    throw new Error(`一度にプレビューできる Agent 操作は${CWS_AGENT_MAX_OPS}件までです。`);
  }

  clearCwsAgentPreview({ restore: true, quiet: true });
  const workbookSheetsBefore = cwsAgentOpsMayChangeSheetStructure(ops) ? cwsAgentCaptureWorkbookSheetsSnapshot() : null;
  const normalizedOps = normalizeCwsAgentOpsBatch(ops, options);
  const sheetIndexes = state.model.sheets.map((_sheet, index) => index);
  beginCwsAgentPreviewPersistence();
  const sheetsBefore = captureSheetsForHistory(sheetIndexes);
  const definedNamesBefore = cloneDefinedNames(state.model.definedNames);
  const activeSheetBefore = state.activeSheetIndex;
  const selectionBefore = cwsAgentCloneRange(state.selectionRange);
  const selectionBeforeRanges = cwsAgentCloneRanges(state.selectionRanges);
  const revisionSnapshot = cwsAgentWorkbookRevisionSnapshot();

  try {
    const results = withCwsAgentPreviewMutation(() => normalizedOps.map((op) => applyCwsAgentOp(op)), revisionSnapshot);
    validateCwsAgentAppliedBranchGeometry(normalizedOps, results);
    const afterSheetIndexes = sheetIndexes.filter((index) => state.model?.sheets?.[index]);
    const workbookSheetsAfter = workbookSheetsBefore ? cwsAgentCaptureWorkbookSheetsSnapshot() : null;
    const historyEntry = {
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(afterSheetIndexes),
      definedNamesBefore,
      definedNamesAfter: cloneDefinedNames(state.model.definedNames),
      selectionBefore,
      selectionAfter: cwsAgentCloneRange(state.selectionRange),
      selectionBeforeRanges,
      selectionAfterRanges: cwsAgentCloneRanges(state.selectionRanges),
      activeSheetBefore,
      activeSheetAfter: state.activeSheetIndex,
      ...(workbookSheetsBefore ? { workbookSheetsBefore, workbookSheetsAfter } : {}),
      label: "CWS Agent プレビュー",
    };
    const previewId = `cws-agent-preview-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
    state.cwsAgentPreview = {
      id: previewId,
      messageId: String(options.messageId || ""),
      historyEntry,
      results,
      ops: normalizedOps,
      visible: true,
    };
    cwsAgentRestoreWorkbookRevisionSnapshot(revisionSnapshot);
    setStatus("CWS Agent: 変更プレビューを表示しています。");
    return { ok: true, previewId, visible: true, applied: results.length, results };
  } catch (error) {
    state.cwsAgentPersistenceSnapshot = null;
    if (workbookSheetsBefore) cwsAgentRestoreWorkbookSheetsSnapshot(workbookSheetsBefore);
    else restoreSheetsFromHistory(sheetsBefore);
    state.model.definedNames = cloneDefinedNames(definedNamesBefore);
    state.activeSheetIndex = Math.max(0, Math.min(activeSheetBefore, state.model.sheets.length - 1));
    cwsAgentRestoreSelectionSnapshot(selectionBefore, selectionBeforeRanges, state.activeSheetIndex);
    state.hf = buildFormulaEngine(state.model);
    cwsAgentRestoreWorkbookRevisionSnapshot(revisionSnapshot);
    renderWorkbook();
    updateSelectionUi();
    updateFormulaBarForSelection();
    throw error;
  }
}



async function previewCwsAgentOpsProgressively(ops = [], options = {}) {
  if (!Array.isArray(ops)) {
    throw new Error("Agent 操作が正しくありません。");
  }
  if (!state.model) {
    throw new Error("ブックが開かれていません。");
  }
  if (ops.length > CWS_AGENT_MAX_OPS) {
    throw new Error(`一度にプレビューできる Agent 操作は${CWS_AGENT_MAX_OPS}件までです。`);
  }

  clearCwsAgentPreview({ restore: true, quiet: true });
  const workbookSheetsBefore = cwsAgentOpsMayChangeSheetStructure(ops) ? cwsAgentCaptureWorkbookSheetsSnapshot() : null;
  const normalizedOps = normalizeCwsAgentOpsBatch(ops, options);
  const sheetIndexes = state.model.sheets.map((_sheet, index) => index);
  beginCwsAgentPreviewPersistence();
  const sheetsBefore = captureSheetsForHistory(sheetIndexes);
  const definedNamesBefore = cloneDefinedNames(state.model.definedNames);
  const activeSheetBefore = state.activeSheetIndex;
  const selectionBefore = cwsAgentCloneRange(state.selectionRange);
  const selectionBeforeRanges = cwsAgentCloneRanges(state.selectionRanges);
  const revisionSnapshot = cwsAgentWorkbookRevisionSnapshot();
  const results = [];
  const wasRestoringHistory = state.isRestoringHistory;

  try {
    state.isRestoringHistory = true;
    for (let index = 0; index < normalizedOps.length; index += 1) {
      const result = applyCwsAgentOp(normalizedOps[index]);
      results.push(result);
      cwsAgentRestoreWorkbookRevisionSnapshot(revisionSnapshot);
      renderCwsAgentProgressivePreviewStep();
      if (typeof options.onProgress === "function") {
        options.onProgress({
          index: index + 1,
          total: normalizedOps.length,
          result,
          results: results.slice(),
        });
      }
      if (index < normalizedOps.length - 1) {
        await cwsAgentDelay(cwsAgentProgressivePreviewDelayMs(options));
      }
    }
    validateCwsAgentAppliedBranchGeometry(normalizedOps, results);
  } catch (error) {
    state.cwsAgentPersistenceSnapshot = null;
    if (workbookSheetsBefore) cwsAgentRestoreWorkbookSheetsSnapshot(workbookSheetsBefore);
    else restoreSheetsFromHistory(sheetsBefore);
    state.model.definedNames = cloneDefinedNames(definedNamesBefore);
    state.activeSheetIndex = Math.max(0, Math.min(activeSheetBefore, state.model.sheets.length - 1));
    cwsAgentRestoreSelectionSnapshot(selectionBefore, selectionBeforeRanges, state.activeSheetIndex);
    state.hf = buildFormulaEngine(state.model);
    renderWorkbook();
    updateSelectionUi();
    updateFormulaBarForSelection();
    throw error;
  } finally {
    state.isRestoringHistory = wasRestoringHistory;
    cwsAgentRestoreWorkbookRevisionSnapshot(revisionSnapshot);
  }

  const afterSheetIndexes = sheetIndexes.filter((index) => state.model?.sheets?.[index]);
  const workbookSheetsAfter = workbookSheetsBefore ? cwsAgentCaptureWorkbookSheetsSnapshot() : null;
  const historyEntry = {
    sheetIndex: state.activeSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(afterSheetIndexes),
    definedNamesBefore,
    definedNamesAfter: cloneDefinedNames(state.model.definedNames),
    selectionBefore,
    selectionAfter: cwsAgentCloneRange(state.selectionRange),
    selectionBeforeRanges,
    selectionAfterRanges: cwsAgentCloneRanges(state.selectionRanges),
    activeSheetBefore,
    activeSheetAfter: state.activeSheetIndex,
    ...(workbookSheetsBefore ? { workbookSheetsBefore, workbookSheetsAfter } : {}),
    label: "CWS Agent プレビュー",
  };
  const previewId = `cws-agent-preview-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  state.cwsAgentPreview = {
    id: previewId,
    messageId: String(options.messageId || ""),
    historyEntry,
    results,
    ops: normalizedOps,
    visible: true,
  };
  cwsAgentRestoreWorkbookRevisionSnapshot(revisionSnapshot);
  setStatus("CWS Agent: 変更プレビューを表示しています。");
  return { ok: true, previewId, visible: true, applied: results.length, results };
}



async function appendCwsAgentPreviewOpsProgressively(ops = [], options = {}) {
  if (!Array.isArray(ops)) {
    throw new Error("Agent 操作が正しくありません。");
  }
  if (!state.model) {
    throw new Error("ブックが開かれていません。");
  }
  if (!state.cwsAgentPreview || (options.messageId && state.cwsAgentPreview.messageId !== String(options.messageId))) {
    return previewCwsAgentOpsProgressively(ops, options);
  }

  const preview = state.cwsAgentPreview;
  const existingResults = Array.isArray(preview.results) ? preview.results : [];
  const existingOps = Array.isArray(preview.ops) ? preview.ops : [];
  if (existingResults.length + ops.length > CWS_AGENT_MAX_OPS) {
    throw new Error(`一度にプレビューできる Agent 操作は${CWS_AGENT_MAX_OPS}件までです。`);
  }
  if (preview.visible === false) {
    applyHistoryEntry(preview.historyEntry, "redo");
    preview.visible = true;
  }

  const normalizedOps = normalizeCwsAgentOpsBatch(ops, {
    ...options,
    existingOps,
    indexOffset: existingOps.length,
  });
  const baseline = preview.historyEntry;
  const revisionSnapshot = cwsAgentWorkbookRevisionSnapshot();
  const results = [];
  const wasRestoringHistory = state.isRestoringHistory;

  try {
    state.isRestoringHistory = true;
    for (let index = 0; index < normalizedOps.length; index += 1) {
      const result = applyCwsAgentOp(normalizedOps[index]);
      results.push(result);
      cwsAgentRestoreWorkbookRevisionSnapshot(revisionSnapshot);
      renderCwsAgentProgressivePreviewStep();
      if (typeof options.onProgress === "function") {
        options.onProgress({
          index: index + 1,
          total: normalizedOps.length,
          result,
          results: results.slice(),
        });
      }
      if (index < normalizedOps.length - 1) {
        await cwsAgentDelay(cwsAgentProgressivePreviewDelayMs(options));
      }
    }
    validateCwsAgentAppliedBranchGeometry(existingOps.concat(normalizedOps), existingResults.concat(results));
  } catch (error) {
    if (baseline.workbookSheetsBefore) cwsAgentRestoreWorkbookSheetsSnapshot(baseline.workbookSheetsBefore);
    else restoreSheetsFromHistory(baseline.sheetsBefore || []);
    state.model.definedNames = cloneDefinedNames(baseline.definedNamesBefore);
    state.activeSheetIndex = Math.max(0, Math.min(Number(baseline.activeSheetBefore) || 0, state.model.sheets.length - 1));
    cwsAgentRestoreSelectionSnapshot(baseline.selectionBefore, baseline.selectionBeforeRanges, state.activeSheetIndex);
    state.hf = buildFormulaEngine(state.model);
    state.cwsAgentPreview = null;
    state.cwsAgentPersistenceSnapshot = null;
    renderWorkbook();
    updateSelectionUi();
    updateFormulaBarForSelection();
    throw error;
  } finally {
    state.isRestoringHistory = wasRestoringHistory;
    cwsAgentRestoreWorkbookRevisionSnapshot(revisionSnapshot);
  }

  const afterSheetIndexes = state.model.sheets.map((_sheet, index) => index);
  preview.results = existingResults.concat(results);
  preview.ops = existingOps.concat(normalizedOps);
  preview.historyEntry = {
    ...baseline,
    sheetsAfter: captureSheetsForHistory(afterSheetIndexes),
    definedNamesAfter: cloneDefinedNames(state.model.definedNames),
    selectionAfter: cwsAgentCloneRange(state.selectionRange),
    selectionAfterRanges: cwsAgentCloneRanges(state.selectionRanges),
    activeSheetAfter: state.activeSheetIndex,
  };
  cwsAgentRestoreWorkbookRevisionSnapshot(revisionSnapshot);
  setStatus("CWS Agent: 変更プレビューを表示しています。");
  return {
    ok: true,
    previewId: preview.id,
    visible: true,
    applied: preview.results.length,
    appended: results.length,
    results: preview.results,
    appendedResults: results,
  };
}



function normalizeCwsAgentOpsBatch(ops = [], options = {}) {
  const indexOffset = Math.max(0, Math.floor(Number(options.indexOffset) || 0));
  const normalizedOps = ops.map((op, index) => normalizeCwsAgentOp(op, indexOffset + index));
  if (cwsAgentOpsMayChangeSheetStructure(ops)) {
    normalizedOps.forEach((op, index) => {
      if (!cwsAgentOpsMayChangeSheetStructure([ops[index]])) {
        cwsAgentOperationSources.set(op, { source: cloneJsonLikeValue(ops[index]), index: indexOffset + index });
      }
    });
  }
  const contractOps = Array.isArray(options.existingOps)
    ? options.existingOps.concat(normalizedOps)
    : normalizedOps;
  validateCwsAgentOpsContract(contractOps, options);
  return normalizedOps;
}



function validateCwsAgentAppliedBranchGeometry(ops, results) {
  for (let index = 0; index < ops.length; index += 1) {
    const op = ops[index];
    if (!op.branchRelation) continue;
    const sheet = state.model.sheets.find(candidate => candidate.images?.some(image => image.id === results[index]?.shapeId));
    const images = sheet?.images || [];
    const label = images.find(image => image.id === results[index]?.shapeId);
    const from = images.filter(image => image.name === op.branchRelation.fromShape);
    const to = images.filter(image => image.name === op.branchRelation.toShape);
    const fail = () => { throw new Error(`図形の適用後検証に失敗しました: 分岐ラベル「${op.style?.text || ""}」が接続線に重なる、離れすぎる、または接続先がありません。作業を完了できません。`); };
    if (!label || from.length !== 1 || to.length !== 1) fail();
    const lines = images.filter(image => image.shape?.connections?.start && image.shape?.connections?.end);
    const own = lines.filter(image => image.shape.connections.start.imageId === from[0].id && image.shape.connections.end.imageId === to[0].id);
    if (own.length !== 1) fail();
    const box = imageBoxToPixels(sheet, label);
    const padded = { left: box.left - 4, top: box.top - 4, width: box.width + 8, height: box.height + 8 };
    for (const line of lines) {
      for (const [a, b] of lineShapeSegmentsInSheetPixels(sheet, line)) {
        if (cwsAgentSegmentCrossesBox(a, b, padded)) fail();
      }
    }
    const distance = Math.min(...lineShapeSegmentsInSheetPixels(sheet, own[0]).map(([a, b]) => Math.hypot(
      Math.max(box.left - Math.max(a.x, b.x), Math.min(a.x, b.x) - box.left - box.width, 0),
      Math.max(box.top - Math.max(a.y, b.y), Math.min(a.y, b.y) - box.top - box.height, 0),
    )));
    if (distance > 64) fail();
  }
}



function cwsAgentSegmentCrossesBox(a, b, box) {
  let entry = 0, exit = 1;
  for (const [axis, start, size] of [["x", "left", "width"], ["y", "top", "height"]]) {
    const low = box[start] + 0.01, high = box[start] + box[size] - 0.01, delta = b[axis] - a[axis];
    if (Math.abs(delta) < 0.0001) {
      if (a[axis] < low || a[axis] > high) return false;
    } else {
      const t1 = (low - a[axis]) / delta, t2 = (high - a[axis]) / delta;
      entry = Math.max(entry, Math.min(t1, t2)); exit = Math.min(exit, Math.max(t1, t2));
      if (entry > exit) return false;
    }
  }
  return true;
}



function cwsAgentOpsMayChangeSheetStructure(ops = []) {
  return (Array.isArray(ops) ? ops : []).some((op) => {
    const type = String(op?.type || op?.op || "").trim().toLowerCase().replace(/[_-]/g, "");
    return [
      "addsheet",
      "createsheet",
      "insertsheet",
      "newworksheet",
      "addworksheet",
      "createworksheet",
      "renamesheet",
      "renameworksheet",
      "duplicatesheet",
      "copysheet",
      "clonesheet",
      "duplicateworksheet",
      "copyworksheet",
      "cloneworksheet",
      "movesheet",
      "moveworksheet",
      "reordersheet",
      "reorderworksheet",
      "deletesheet",
      "removesheet",
      "deleteworksheet",
      "removeworksheet",
      "createworkbooktemplate",
      "applyworkbooktemplate",
      "workbooktemplate",
    ].includes(type);
  });
}



function validateCwsAgentOpsContract(ops = [], options = {}) {
  validateCwsAgentRequiredOpsContract(ops, options);
  validateCwsAgentAllowedRangeContract(ops, options);
  validateCwsAgentHeaderStyleScope(ops);
}



function validateCwsAgentRequiredOpsContract(ops = [], options = {}) {
  const requiredOps = cwsAgentContractRequiredOps(options);
  if (!requiredOps.length) return;
  const present = new Set(ops.map((op) => op?.type).filter(Boolean));
  const missing = requiredOps.find((type) => !present.has(type));
  if (!missing) return;
  throw new Error(`Agent 操作の計画契約に必要な ${missing} が含まれていません。`);
}



function cwsAgentContractRequiredOps(options = {}) {
  const contract = cwsAgentPlainObject(options.contract) || cwsAgentPlainObject(options.agentContract) || {};
  const raw = Array.isArray(options.requiredOps)
    ? options.requiredOps
    : Array.isArray(contract.requiredOps)
      ? contract.requiredOps
      : [];
  const supported = [
    "setCell",
    "setRange",
    "setStyle",
    "setCellProtection",
    "clearCellProtection",
    "setSparkline",
    "clearSparkline",
    "setBorders",
    "setNumberFormat",
    "sortRange",
    "setAutoFilter",
    "clearAutoFilter",
    "setFilterCriteria",
    "clearFilterCriteria",
    "advancedFilter",
    "removeDuplicates",
    "findReplace",
    "textToColumns",
    "createTable",
    "createChart",
    "createPivotTable",
    "createForecastSheet",
    "goalSeek",
    "dataTable",
    "addShape",
    "setShapeText",
    "setShapeStyle",
    "removeShape",
    "createWorkbookTemplate",
    "setPrintArea",
    "clearPrintArea",
    "setPrintTitles",
    "clearPrintTitles",
    "setHeaderFooter",
    "clearHeaderFooter",
    "setPageTheme",
    "setPageOrientation",
    "setPaperSize",
    "setPageMargins",
    "setPageScale",
    "setPrintOptions",
    "insertPageBreak",
    "removePageBreak",
    "resetPageBreaks",
    "setSheetView",
    "setQuickAnalysis",
    "clearQuickAnalysis",
    "setQuickAnalysisTotal",
    "setSubtotal",
    "clearSubtotals",
    "consolidateRanges",
    "setDataValidation",
    "clearDataValidation",
    "setHyperlink",
    "clearHyperlink",
    "setNote",
    "clearNote",
    "setComment",
    "appendComment",
    "clearComment",
    "setNamedRange",
    "clearNamedRange",
    "protectSheet",
    "unprotectSheet",
    "activateSheet",
    "selectRange",
    "copyRange",
    "moveRange",
    "addSheet",
    "renameSheet",
    "duplicateSheet",
    "moveSheet",
    "setSheetTabColor",
    "clearSheetTabColor",
    "deleteSheet",
    "hideSheet",
    "unhideSheet",
    "setFreezePanes",
    "clearFreezePanes",
    "setColumnWidth",
    "setRowHeight",
    "autoFitColumns",
    "autoFitRows",
    "hideRows",
    "unhideRows",
    "hideColumns",
    "unhideColumns",
    "groupRows",
    "ungroupRows",
    "groupColumns",
    "ungroupColumns",
    "clearRange",
    "mergeCells",
    "unmergeCells",
    "insertRows",
    "deleteRows",
    "insertColumns",
    "deleteColumns",
  ];
  const supportedByKey = new Map(supported.map((type) => [cwsAgentContractOpTypeKey(type), type]));
  return [...new Set(raw
    .map((type) => supportedByKey.get(cwsAgentContractOpTypeKey(type)))
    .filter(Boolean))];
}



function cwsAgentContractOpTypeKey(value) {
  return String(value || "").trim().toLowerCase().replace(/[_-]+/g, "");
}



function validateCwsAgentAllowedRangeContract(ops = [], options = {}) {
  const allowedRanges = cwsAgentContractAllowedRanges(options);
  if (!allowedRanges.length) return;
  for (const op of ops) {
    const opRange = cwsAgentRangeForContractOp(op);
    if (!opRange || cwsAgentRangeWithinAny(opRange, allowedRanges)) continue;
    throw new Error(`Agent 操作の範囲 ${rangeToLabel(opRange)} は計画契約の範囲外です。`);
  }
}



function cwsAgentContractAllowedRanges(options = {}) {
  const contract = cwsAgentPlainObject(options.contract) || cwsAgentPlainObject(options.agentContract) || {};
  const raw = Array.isArray(options.allowedRanges)
    ? options.allowedRanges
    : Array.isArray(contract.allowedRanges)
      ? contract.allowedRanges
      : [];
  return raw.map(cwsAgentNormalizeContractRange).filter(Boolean);
}



function cwsAgentNormalizeContractRange(value) {
  const parsed = typeof value === "string" ? cwsAgentParseRangeAddress(value) : null;
  const source = parsed || cwsAgentPlainObject(value);
  if (!source) return null;
  const top = cwsAgentPositiveInteger(source.top ?? source.row);
  const left = cwsAgentColumnValue(source.left ?? source.col);
  const bottom = cwsAgentPositiveInteger(source.bottom) || top;
  const right = cwsAgentColumnValue(source.right) || left;
  if (!top || !left || bottom < top || right < left) return null;
  return { top, left, bottom, right };
}



function cwsAgentRangeForContractOp(op) {
  if (!op || typeof op !== "object") return null;
  if (op.type === "setCell") {
    return { top: op.row, left: op.col, bottom: op.row, right: op.col };
  }
  if (op.type === "goalSeek") {
    return {
      top: op.changingRow,
      left: op.changingCol,
      bottom: op.changingRow,
      right: op.changingCol,
    };
  }
  if (!CWS_AGENT_CONTRACT_RANGE_OPS.has(op.type)) return null;
  const top = cwsAgentPositiveInteger(op.top);
  const left = cwsAgentColumnValue(op.left);
  const bottom = cwsAgentPositiveInteger(op.bottom) || top;
  const right = cwsAgentColumnValue(op.right) || left;
  if (!top || !left || bottom < top || right < left) return null;
  return { top, left, bottom, right };
}



function cwsAgentRangeWithinAny(candidate, ranges) {
  return ranges.some((range) =>
    candidate.top >= range.top &&
    candidate.left >= range.left &&
    candidate.bottom <= range.bottom &&
    candidate.right <= range.right,
  );
}



function validateCwsAgentHeaderStyleScope(ops = []) {
  const generatedRanges = ops.filter((op) =>
    op?.type === "setRange" &&
    op.bottom > op.top &&
    op.right >= op.left &&
    Array.isArray(op.values) &&
    op.values.length > 1 &&
    cwsAgentSetRangeValuesLookLikeTable(op.values),
  );
  if (!generatedRanges.length) return;
  for (const op of ops) {
    if (op?.type !== "setStyle" || !cwsAgentStyleLooksLikeHeaderStyle(op.style)) continue;
    for (const range of generatedRanges) {
      if (!cwsAgentStyleCoversGeneratedHeader(op, range)) continue;
      if (op.bottom === range.top) continue;
      throw new Error("Agent 操作の見出し書式が表全体に適用されています。見出し用の背景色や白文字は見出し行だけに適用してください。");
    }
  }
}



function cwsAgentSetRangeValuesLookLikeTable(values) {
  const nonempty = (row) => Array.isArray(row)
    ? row.filter((cell) => String(cell ?? "").trim() !== "").length : 0;
  return Array.isArray(values) && values.length > 1 && nonempty(values[0]) >= 2 &&
    values.slice(1).some((row) => nonempty(row) >= 2);
}



function cwsAgentStyleLooksLikeHeaderStyle(style) {
  if (!style || typeof style !== "object") return false;
  const fontWeight = String(style.fontWeight || "").toLowerCase();
  const headerSignals = [
    Boolean(style.backgroundColor),
    Boolean(style.color),
    fontWeight === "700" || fontWeight === "bold",
    String(style.textAlign || "").toLowerCase() === "center",
  ];
  return headerSignals.filter(Boolean).length >= 2;
}



function cwsAgentStyleCoversGeneratedHeader(styleOp, rangeOp) {
  if (styleOp.sheetIndex !== rangeOp.sheetIndex) return false;
  if (styleOp.top !== rangeOp.top) return false;
  if (styleOp.left > rangeOp.left || styleOp.right < rangeOp.right) return false;
  if (styleOp.bottom === rangeOp.top) return true;
  return styleOp.bottom >= rangeOp.bottom;
}



function cwsAgentProgressivePreviewDelayMs(options = {}) {
  const value = Number(options.stepDelayMs);
  if (Number.isFinite(value)) return Math.max(0, Math.min(500, value));
  return CWS_AGENT_PROGRESSIVE_PREVIEW_DELAY_MS;
}



function cwsAgentDelay(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}



function renderCwsAgentProgressivePreviewStep() {
  renderWorkbook();
  updateSelectionUi();
  updateFormulaBarForSelection();
}



function setCwsAgentPreviewVisible(previewId, visible) {
  const preview = cwsAgentPreviewForId(previewId);
  const nextVisible = visible !== false;
  if (preview.visible !== nextVisible) {
    applyHistoryEntry(preview.historyEntry, nextVisible ? "redo" : "undo");
    preview.visible = nextVisible;
  }
  setStatus(nextVisible ? "CWS Agent: 変更後を表示しています。" : "CWS Agent: 変更前を表示しています。");
  return { ok: true, previewId: preview.id, visible: preview.visible, applied: preview.results?.length || 0, results: preview.results || [] };
}



function commitCwsAgentPreview(previewId) {
  const preview = cwsAgentPreviewForId(previewId);
  if (preview.visible === false) {
    applyHistoryEntry(preview.historyEntry, "redo");
    preview.visible = true;
  }
  pushHistory(preview.historyEntry);
  state.cwsAgentPreview = null;
  state.cwsAgentPersistenceSnapshot = null;
  const results = preview.results || [];
  setStatus(`CWS Agent: ${results.length > 1 ? `${results.length}件の変更` : "変更"}を適用しました。`);
  return { ok: true, previewId: preview.id, visible: true, applied: results.length, results };
}



function captureCwsAgentTaskWorkbook() {
  if (!state.model) throw new Error("ブックが開かれていません。");
  return { documentId: state.model.cwsAgentDocumentId, sheets: cwsAgentCaptureWorkbookSheetsSnapshot(),
    definedNames: cloneDefinedNames(state.model.definedNames), activeSheetIndex: state.activeSheetIndex };
}



function normalizeCwsAgentTaskSnapshot(snapshot) {
  // Reuse export's style compaction while retaining unknown or meaningful
  // metadata that export must never be allowed to erase from recovery identity.
  const styleKeys=new Set(['row','col','css','borders','hasFill','hasFontColor']);
  for(const sheet of snapshot.sheets || []) {
    sheet.formulaAudits={precedents:[],dependents:[],arrows:[],ignoredErrors:[],...(sheet.formulaAudits || {})};
    const retained={},previousRanges=sheet.styleRanges || [];
    for(const [key,cell] of Object.entries(sheet.cells || {})) {
      if(Object.keys(cell).some(k=>!styleKeys.has(k))) {retained[key]=cell;delete sheet.cells[key];}
    }
    sheet.styleRanges=[];
    compactSheetStyleOnlyCellsToStyleRanges(sheet);
    sheet.cells={...sheet.cells,...retained};
    sheet.styleRanges=[...previousRanges,...(sheet.styleRanges || [])];
  }
  return snapshot;
}



function restoreCwsAgentTaskWorkbook(snapshot) {
  if (!snapshot || snapshot.documentId !== state.model?.cwsAgentDocumentId || !Array.isArray(snapshot.sheets) || !snapshot.sheets.length) {
    throw new Error("復元対象のブックが一致しません。");
  }
  clearCwsAgentPreview({ restore: true, quiet: true });
  cwsAgentRestoreWorkbookSheetsSnapshot(snapshot.sheets);
  state.model.definedNames = cloneDefinedNames(snapshot.definedNames);
  state.activeSheetIndex = Math.max(0, Math.min(snapshot.activeSheetIndex || 0, state.model.sheets.length - 1));
  state.hf = buildFormulaEngine(state.model);
  renderWorkbook(); updateSelectionUi(); updateFormulaBarForSelection();
  scheduleWorkbookPersistence();
  return { ok: true };
}



async function persistCwsAgentTaskWorkbook() {
  if (state.persistenceLoading || !state.model || !persistenceAvailable()) throw new Error("ブックの保存を確認できません。");
  await writePersistedWorkbookSnapshot(currentWorkbookSnapshot());
  return { ok: true };
}



function clearCwsAgentPreview(options = {}) {
  state.cwsAgentPersistenceSnapshot = null;
  const preview = state.cwsAgentPreview;
  if (!preview) return { ok: true, cleared: false };
  if (options.restore !== false && preview.visible !== false) {
    applyHistoryEntry(preview.historyEntry, "undo");
  }
  state.cwsAgentPreview = null;
  if (!options.quiet) {
    setStatus("CWS Agent: 変更プレビューを取り消しました。");
  }
  return { ok: true, cleared: true };
}



function cwsAgentPreviewForId(previewId) {
  const preview = state.cwsAgentPreview;
  if (!preview || (previewId && preview.id !== previewId)) {
    throw new Error("変更プレビューはもう利用できません。");
  }
  return preview;
}



function withCwsAgentPreviewMutation(callback, revisionSnapshot = cwsAgentWorkbookRevisionSnapshot()) {
  const wasRestoringHistory = state.isRestoringHistory;
  state.isRestoringHistory = true;
  try {
    return callback();
  } finally {
    state.isRestoringHistory = wasRestoringHistory;
    cwsAgentRestoreWorkbookRevisionSnapshot(revisionSnapshot);
  }
}



function cwsAgentWorkbookRevisionSnapshot() {
  return {
    workbookRevision: state.workbookRevision,
    workbookSavedRevision: state.workbookSavedRevision,
  };
}



function cwsAgentRestoreWorkbookRevisionSnapshot(snapshot = {}) {
  state.workbookRevision = Number(snapshot.workbookRevision) || 0;
  state.workbookSavedRevision = Number(snapshot.workbookSavedRevision) || 0;
}



function cwsAgentCloneRange(range) {
  return range ? { ...range } : null;
}



function cwsAgentCloneRanges(ranges) {
  return Array.isArray(ranges) ? ranges.map(cwsAgentCloneRange).filter(Boolean) : [];
}



function cwsAgentRestoreSelectionSnapshot(selection, ranges, sheetIndex) {
  if (selection) {
    const normalizedSelection = { ...selection, sheetIndex };
    state.selected = { sheetIndex, row: normalizedSelection.top || normalizedSelection.row || 1, col: normalizedSelection.left || normalizedSelection.col || 1 };
    state.selectionRange = normalizedSelection;
    state.selectionRanges = ranges?.length ? ranges.map((range) => ({ ...range, sheetIndex })) : [normalizedSelection];
    state.selectionStart = state.selected;
  } else {
    state.selected = null;
    state.selectionRange = null;
    state.selectionRanges = [];
    state.selectionStart = null;
  }
  state.selectionDrag = null;
  state.selectedImage = null;
}



function normalizeCwsAgentOp(op, index) {
  const type = String(op?.type || op?.op || "").trim().toLowerCase().replace(/[_-]/g, "");
  if (type === "createworkbooktemplate" || type === "applyworkbooktemplate" || type === "workbooktemplate") {
    return normalizeCwsAgentWorkbookTemplateOp(op, index);
  }
  if (type === "setcell") {
    return normalizeCwsAgentSetCellOp(op, index);
  }
  if (type === "setrange") {
    return normalizeCwsAgentSetRangeOp(op, index);
  }
  if (type === "setstyle") {
    return normalizeCwsAgentSetStyleOp(op, index);
  }
  if (type === "setcellprotection" || type === "cellprotection" || type === "setprotection" || type === "lockcells" || type === "unlockcells" || type === "hideformulas" || type === "showformulas" || type === "unhideformulas") {
    return normalizeCwsAgentSetCellProtectionOp(op, index, type);
  }
  if (type === "clearcellprotection" || type === "resetcellprotection" || type === "removecellprotection") {
    return normalizeCwsAgentClearCellProtectionOp(op, index);
  }
  if (type === "setsparkline" || type === "createsparkline" || type === "addsparkline" || type === "sparkline" || type === "insertsparkline") {
    return normalizeCwsAgentSetSparklineOp(op, index);
  }
  if (type === "clearsparkline" || type === "removesparkline" || type === "deletesparkline") {
    return normalizeCwsAgentClearSparklineOp(op, index);
  }
  if (type === "setborder" || type === "setborders") {
    return normalizeCwsAgentSetBordersOp(op, index);
  }
  if (type === "setnumberformat" || type === "numberformat" || type === "setnumfmt" || type === "numfmt") {
    return normalizeCwsAgentSetNumberFormatOp(op, index);
  }
  if (type === "sortrange" || type === "sortrows" || type === "sort") {
    return normalizeCwsAgentSortRangeOp(op, index);
  }
  if (type === "setautofilter" || type === "autofilter" || type === "filterrange" || type === "setfilter" || type === "enablefilter") {
    return normalizeCwsAgentSetAutoFilterOp(op, index);
  }
  if (type === "clearautofilter" || type === "removeautofilter" || type === "disablefilter" || type === "clearfilter") {
    return normalizeCwsAgentClearAutoFilterOp(op, index);
  }
  if (type === "setfiltercriteria" || type === "setautofiltercriteria" || type === "applyfiltercriteria" || type === "filtercriteria" || type === "filterrows") {
    return normalizeCwsAgentSetFilterCriteriaOp(op, index);
  }
  if (type === "clearfiltercriteria" || type === "clearautofiltercriteria" || type === "removefiltercriteria") {
    return normalizeCwsAgentClearFilterCriteriaOp(op, index);
  }
  if (type === "advancedfilter" || type === "applyadvancedfilter" || type === "advancedfiltercopy" || type === "filteradvanced" || type === "extractfilter") {
    return normalizeCwsAgentAdvancedFilterOp(op, index);
  }
  if (type === "removeduplicates" || type === "deduplicate" || type === "deduperows" || type === "deleteduplicates" || type === "dropduplicates" || type === "uniquerows") {
    return normalizeCwsAgentRemoveDuplicatesOp(op, index);
  }
  if (type === "findreplace" || type === "replaceall" || type === "replacecelltext" || type === "replacevalues" || type === "replace") {
    return normalizeCwsAgentFindReplaceOp(op, index);
  }
  if (type === "texttocolumns" || type === "textcolumns" || type === "splittexttocolumns" || type === "splitcolumn" || type === "splitcolumns" || type === "splitdelimited" || type === "splitfixedwidth") {
    return normalizeCwsAgentTextToColumnsOp(op, index);
  }
  if (type === "createtable" || type === "formattable" || type === "formatastable" || type === "maketable" || type === "converttotable" || type === "table" || type === "addtable") {
    return normalizeCwsAgentCreateTableOp(op, index);
  }
  if (type === "createchart" || type === "insertchart" || type === "addchart" || type === "chart" || type === "columnchart" || type === "barchart" || type === "linechart" || type === "piechart" || type === "scatterchart" || type === "mapchart") {
    return normalizeCwsAgentCreateChartOp(op, index, type);
  }
  if (type === "createpivottable" || type === "insertpivottable" || type === "addpivottable" || type === "pivottable" || type === "pivot" || type === "quickanalysispivottable") {
    return normalizeCwsAgentCreatePivotTableOp(op, index);
  }
  if (type === "createforecastsheet" || type === "forecastsheet" || type === "addforecastsheet" || type === "insertforecastsheet" || type === "forecastworksheet" || type === "createforecastworksheet") {
    return normalizeCwsAgentCreateForecastSheetOp(op, index);
  }
  if (type === "goalseek" || type === "whatifgoalseek" || type === "whatifanalysis" || type === "goalsearch" || type === "targetseek") {
    return normalizeCwsAgentGoalSeekOp(op, index);
  }
  if (type === "datatable" || type === "whatifdatatable" || type === "datatables" || type === "whatiftable" || type === "sensitivitytable" || type === "sensitivityanalysis") {
    return normalizeCwsAgentDataTableOp(op, index);
  }
  if (type === "addshape" || type === "createshape" || type === "insertshape" || type === "drawshape" || type === "shape" || type === "autoshape" || type === "textbox" || type === "textshape" || type === "addconnector" || type === "createconnector" || type === "insertconnector" || type === "drawconnector" || type === "connector" || type === "line" || type === "arrow") {
    return normalizeCwsAgentAddShapeOp(op, index, type);
  }
  if (type === "setshapestyle" || type === "removeshape") return normalizeCwsAgentExistingShapeEditOp(op);
  if (type === "setshapetext") return normalizeCwsAgentSetShapeTextOp(op, index);
  if (type === "setquickanalysis" || type === "quickanalysis" || type === "databar" || type === "databars" || type === "colorscale" || type === "iconset" || type === "greaterthanaverage" || type === "aboveaverage" || type === "duplicatevalues") {
    return normalizeCwsAgentSetQuickAnalysisOp(op, index, type);
  }
  if (type === "clearquickanalysis" || type === "removequickanalysis" || type === "clearquickanalysisformatting" || type === "clearconditionalstyle") {
    return normalizeCwsAgentClearQuickAnalysisOp(op, index);
  }
  if (type === "setquickanalysistotal" || type === "quickanalysistotal" || type === "addquickanalysistotal" || type === "quickanalysissum" || type === "quickanalysisaverage" || type === "quickanalysiscount" || type === "quickanalysispercent" || type === "quickanalysisrunningtotal") {
    return normalizeCwsAgentSetQuickAnalysisTotalOp(op, index, type);
  }
  if (type === "setsubtotal" || type === "subtotal" || type === "addsubtotal" || type === "createsubtotal" || type === "subtotals") {
    return normalizeCwsAgentSetSubtotalOp(op, index);
  }
  if (type === "clearsubtotal" || type === "clearsubtotals" || type === "removesubtotal" || type === "removesubtotals" || type === "deletesubtotal" || type === "deletesubtotals") {
    return normalizeCwsAgentClearSubtotalsOp(op, index);
  }
  if (type === "consolidateranges" || type === "consolidate" || type === "dataconsolidate" || type === "setconsolidation" || type === "createconsolidation" || type === "consolidation") {
    return normalizeCwsAgentConsolidateOp(op, index);
  }
  if (type === "setdatavalidation" || type === "datavalidation" || type === "setvalidation" || type === "setdropdown" || type === "setdropdownlist" || type === "dropdownlist") {
    return normalizeCwsAgentSetDataValidationOp(op, index);
  }
  if (type === "cleardatavalidation" || type === "clearvalidation" || type === "removedatavalidation" || type === "removevalidation" || type === "cleardropdown" || type === "removedropdown") {
    return normalizeCwsAgentClearDataValidationOp(op, index);
  }
  if (type === "sethyperlink" || type === "hyperlink" || type === "setlink" || type === "linkcell" || type === "linkcells") {
    return normalizeCwsAgentSetHyperlinkOp(op, index);
  }
  if (type === "clearhyperlink" || type === "removehyperlink" || type === "clearlink" || type === "removelink" || type === "unlink") {
    return normalizeCwsAgentClearHyperlinkOp(op, index);
  }
  if (type === "setnote" || type === "note" || type === "setmemo" || type === "memo" || type === "cellnote" || type === "cellmemo") {
    return normalizeCwsAgentSetNoteOp(op, index);
  }
  if (type === "clearnote" || type === "removenote" || type === "deletenote" || type === "clearmemo" || type === "removememo" || type === "deletememo") {
    return normalizeCwsAgentClearNoteOp(op, index);
  }
  if (type === "setcomment" || type === "comment" || type === "addcomment" || type === "cellcomment") {
    return normalizeCwsAgentSetCommentOp(op, index);
  }
  if (type === "appendcomment" || type === "replycomment" || type === "addcommentreply" || type === "commentreply") {
    return normalizeCwsAgentAppendCommentOp(op, index);
  }
  if (type === "clearcomment" || type === "removecomment" || type === "deletecomment") {
    return normalizeCwsAgentClearCommentOp(op, index);
  }
  if (type === "setnamedrange" || type === "addnamedrange" || type === "definename" || type === "setdefinedname" || type === "namedrange" || type === "definedname") {
    return normalizeCwsAgentSetNamedRangeOp(op, index);
  }
  if (type === "clearnamedrange" || type === "removenamedrange" || type === "deletenamedrange" || type === "cleardefinedname" || type === "removedefinedname" || type === "deletedefinedname") {
    return normalizeCwsAgentClearNamedRangeOp(op, index);
  }
  if (type === "protectsheet" || type === "setsheetprotection" || type === "sheetprotection") {
    return normalizeCwsAgentProtectSheetOp(op, index);
  }
  if (type === "unprotectsheet" || type === "clearsheetprotection" || type === "removesheetprotection") {
    return normalizeCwsAgentUnprotectSheetOp(op, index);
  }
  if (type === "activatesheet" || type === "selectsheet" || type === "gotosheet" || type === "opensheet" || type === "switchsheet") {
    return normalizeCwsAgentActivateSheetOp(op, index);
  }
  if (type === "selectrange" || type === "selectcell" || type === "selectcells" || type === "selectcellrange" || type === "gotorange" || type === "gotocell" || type === "jumptorange" || type === "jumpcell") {
    return normalizeCwsAgentSelectRangeOp(op, index);
  }
  if (type === "copyrange" || type === "copycell" || type === "copycells" || type === "copycellrange" || type === "duplicaterange") {
    return normalizeCwsAgentCopyMoveRangeOp(op, index, "copyRange");
  }
  if (type === "moverange" || type === "movecell" || type === "movecells" || type === "movecellrange" || type === "cutrange" || type === "cutcells") {
    return normalizeCwsAgentCopyMoveRangeOp(op, index, "moveRange");
  }
  if (type === "addsheet" || type === "createsheet" || type === "insertsheet" || type === "newworksheet" || type === "addworksheet" || type === "createworksheet") {
    return normalizeCwsAgentAddSheetOp(op, index);
  }
  if (type === "renamesheet" || type === "renameworksheet") {
    return normalizeCwsAgentRenameSheetOp(op, index);
  }
  if (type === "duplicatesheet" || type === "copysheet" || type === "clonesheet" || type === "duplicateworksheet" || type === "copyworksheet" || type === "cloneworksheet") {
    return normalizeCwsAgentDuplicateSheetOp(op, index);
  }
  if (type === "movesheet" || type === "moveworksheet" || type === "reordersheet" || type === "reorderworksheet") {
    return normalizeCwsAgentMoveSheetOp(op, index);
  }
  if (type === "setsheettabcolor" || type === "sheettabcolor" || type === "settabcolor" || type === "tabcolor") {
    return normalizeCwsAgentSheetTabColorOp(op, index, false);
  }
  if (type === "clearsheettabcolor" || type === "cleartabcolor" || type === "removesheettabcolor" || type === "removetabcolor") {
    return normalizeCwsAgentSheetTabColorOp(op, index, true);
  }
  if (type === "deletesheet" || type === "removesheet" || type === "deleteworksheet" || type === "removeworksheet") {
    return normalizeCwsAgentDeleteSheetOp(op, index);
  }
  if (type === "hidesheet" || type === "hideworksheet") {
    return normalizeCwsAgentSheetVisibilityOp(op, index, true);
  }
  if (type === "unhidesheet" || type === "unhideworksheet" || type === "showsheet" || type === "showworksheet") {
    return normalizeCwsAgentSheetVisibilityOp(op, index, false);
  }
  if (type === "setfreezepanes" || type === "freezepanes" || type === "freezepane" || type === "freeze" || type === "freezeview") {
    return normalizeCwsAgentSetFreezePanesOp(op, index);
  }
  if (type === "clearfreezepanes" || type === "unfreezepanes" || type === "unfreezepane" || type === "clearfreeze" || type === "unfreeze") {
    return normalizeCwsAgentClearFreezePanesOp(op, index);
  }
  if (type === "setsheetview" || type === "sheetview" || type === "setviewmode" || type === "viewmode" || type === "viewsheet" || type === "setzoom" || type === "setsheetzoom" || type === "zoom" || type === "setviewoptions" || type === "viewoptions" || type === "setdisplayoptions" || type === "displayoptions" || type === "normalview" || type === "pagelayoutview" || type === "pagebreakview" || type === "pagebreakpreview") {
    return normalizeCwsAgentSetSheetViewOp(op, index, type);
  }
  if (type === "setprintarea" || type === "printarea" || type === "defineprintarea") {
    return normalizeCwsAgentSetPrintAreaOp(op, index);
  }
  if (type === "clearprintarea" || type === "removeprintarea" || type === "resetprintarea") {
    return normalizeCwsAgentClearPrintAreaOp(op, index);
  }
  if (type === "setprinttitles" || type === "printtitles" || type === "setprinttitle" || type === "repeatprinttitles" || type === "repeatrowscolumns") {
    return normalizeCwsAgentSetPrintTitlesOp(op, index);
  }
  if (type === "clearprinttitles" || type === "removeprinttitles" || type === "resetprinttitles") {
    return normalizeCwsAgentClearPrintTitlesOp(op, index);
  }
  if (type === "setheaderfooter" || type === "headerfooter" || type === "setpageheaderfooter" || type === "setprintheaderfooter" || type === "setpageheader" || type === "setpagefooter" || type === "setprintheader" || type === "setprintfooter") {
    return normalizeCwsAgentSetHeaderFooterOp(op, index, type);
  }
  if (type === "clearheaderfooter" || type === "removeheaderfooter" || type === "resetheaderfooter" || type === "clearpageheaderfooter" || type === "clearprintheaderfooter" || type === "clearpageheader" || type === "clearpagefooter" || type === "removepageheader" || type === "removepagefooter") {
    return normalizeCwsAgentClearHeaderFooterOp(op, index, type);
  }
  if (type === "setpagetheme" || type === "pagetheme" || type === "settheme" || type === "setworkbooktheme" || type === "setlayouttheme" || type === "setthemecolors" || type === "setcolorscheme" || type === "setthemefonts" || type === "setfontscheme" || type === "setthemeeffects" || type === "seteffectscheme" || type === "resetpagetheme" || type === "resettheme") {
    return normalizeCwsAgentSetPageThemeOp(op, index, type);
  }
  if (type === "setpageorientation" || type === "pageorientation" || type === "setprintorientation" || type === "printorientation" || type === "orientation") {
    return normalizeCwsAgentSetPageOrientationOp(op, index);
  }
  if (type === "setpapersize" || type === "papersize" || type === "setpagepapersize" || type === "pagepapersize" || type === "printpapersize") {
    return normalizeCwsAgentSetPaperSizeOp(op, index);
  }
  if (type === "setpagemargins" || type === "pagemargins" || type === "setprintmargins" || type === "printmargins" || type === "margins") {
    return normalizeCwsAgentSetPageMarginsOp(op, index);
  }
  if (type === "setpagescale" || type === "pagescale" || type === "setprintscale" || type === "printscale" || type === "fitprinttopage" || type === "fitprinttopages" || type === "fitpage") {
    return normalizeCwsAgentSetPageScaleOp(op, index);
  }
  if (type === "setprintoptions" || type === "printoptions" || type === "setpageprintoptions" || type === "pageprintoptions" || type === "setpagesheetoptions") {
    return normalizeCwsAgentSetPrintOptionsOp(op, index);
  }
  if (type === "insertpagebreak" || type === "addpagebreak" || type === "setpagebreak" || type === "pagebreak") {
    return normalizeCwsAgentPageBreakOp(op, index, "insertPageBreak");
  }
  if (type === "removepagebreak" || type === "deletepagebreak" || type === "clearpagebreak") {
    return normalizeCwsAgentPageBreakOp(op, index, "removePageBreak");
  }
  if (type === "resetpagebreaks" || type === "clearallpagebreaks" || type === "removeallpagebreaks") {
    return normalizeCwsAgentResetPageBreaksOp(op, index);
  }
  if (type === "setcolumnwidth" || type === "setcolwidth" || type === "columnwidth" || type === "colwidth") {
    return normalizeCwsAgentSetColumnWidthOp(op, index);
  }
  if (type === "setrowheight" || type === "rowheight") {
    return normalizeCwsAgentSetRowHeightOp(op, index);
  }
  if (type === "autofitcolumns" || type === "autofitcolumn" || type === "fitcolumns" || type === "fitcolumn") {
    return normalizeCwsAgentAutoFitColumnsOp(op, index);
  }
  if (type === "autofitrows" || type === "autofitrow" || type === "fitrows" || type === "fitrow") {
    return normalizeCwsAgentAutoFitRowsOp(op, index);
  }
  if (type === "hiderow" || type === "hiderows") {
    return normalizeCwsAgentHeaderHiddenOp(op, index, "row", true);
  }
  if (type === "unhiderow" || type === "unhiderows" || type === "showrow" || type === "showrows") {
    return normalizeCwsAgentHeaderHiddenOp(op, index, "row", false);
  }
  if (type === "hidecolumn" || type === "hidecolumns" || type === "hidecol" || type === "hidecols") {
    return normalizeCwsAgentHeaderHiddenOp(op, index, "column", true);
  }
  if (type === "unhidecolumn" || type === "unhidecolumns" || type === "unhidecol" || type === "unhidecols" || type === "showcolumn" || type === "showcolumns" || type === "showcol" || type === "showcols") {
    return normalizeCwsAgentHeaderHiddenOp(op, index, "column", false);
  }
  if (type === "grouprow" || type === "grouprows" || type === "outlinerow" || type === "outlinerows") {
    return normalizeCwsAgentOutlineOp(op, index, "row", 1);
  }
  if (type === "ungrouprow" || type === "ungrouprows" || type === "unoutlinerow" || type === "unoutlinerows") {
    return normalizeCwsAgentOutlineOp(op, index, "row", -1);
  }
  if (type === "groupcolumn" || type === "groupcolumns" || type === "groupcol" || type === "groupcols" || type === "outlinecolumn" || type === "outlinecolumns" || type === "outlinecol" || type === "outlinecols") {
    return normalizeCwsAgentOutlineOp(op, index, "column", 1);
  }
  if (type === "ungroupcolumn" || type === "ungroupcolumns" || type === "ungroupcol" || type === "ungroupcols" || type === "unoutlinecolumn" || type === "unoutlinecolumns" || type === "unoutlinecol" || type === "unoutlinecols") {
    return normalizeCwsAgentOutlineOp(op, index, "column", -1);
  }
  if (type === "insertrow" || type === "insertrows") {
    return normalizeCwsAgentRowsOp(op, index, "insertRows");
  }
  if (type === "deleterow" || type === "deleterows") {
    return normalizeCwsAgentRowsOp(op, index, "deleteRows");
  }
  if (type === "insertcolumn" || type === "insertcolumns" || type === "insertcol" || type === "insertcols") {
    return normalizeCwsAgentColumnsOp(op, index, "insertColumns");
  }
  if (type === "deletecolumn" || type === "deletecolumns" || type === "deletecol" || type === "deletecols") {
    return normalizeCwsAgentColumnsOp(op, index, "deleteColumns");
  }
  if (type === "clearrange" || type === "clearcells" || type === "clearcell") {
    return normalizeCwsAgentClearRangeOp(op, index);
  }
  if (type === "mergecells" || type === "mergecell" || type === "mergecenter") {
    return normalizeCwsAgentMergeCellsOp(op, index, type === "mergecenter");
  }
  if (type === "unmergecells" || type === "unmergecell" || type === "unmerge") {
    return normalizeCwsAgentUnmergeCellsOp(op, index);
  }
  throw new Error(`対応していない Agent 操作です: ${type || index + 1}`);
}



function normalizeCwsAgentSetCellOp(op, index) {
  const address = cwsAgentParseCellAddress(op?.address || op?.ref || op?.cell || "");
  const row = cwsAgentPositiveInteger(op?.row) || address?.row || 0;
  const col = cwsAgentColumnValue(op?.col) || address?.col || 0;
  if (!row || !col || row > EXCEL_MAX_ROWS || col > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のセル位置が正しくありません。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, address?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }

  return {
    type: "setCell",
    sheetIndex,
    row,
    col,
    value: cwsAgentCellValue(op?.value),
  };
}



function normalizeCwsAgentSetRangeOp(op, index) {
  const values = cwsAgentMatrixValues(op?.values || op?.rows || op?.data);
  if (!values.length) {
    throw new Error(`Agent 操作 ${index + 1}: setRange には矩形の values 配列が必要です。書式だけの変更は setStyle を使ってください。`);
  }
  const range = cwsAgentRangeForOp(op, index, values);
  return {
    type: "setRange",
    ...range,
    values,
  };
}



function normalizeCwsAgentSetStyleOp(op, index) {
  const style = cwsAgentStylePatch(op?.style || op || {});
  if (!Object.keys(style).length) {
    throw new Error(`Agent 操作 ${index + 1} の書式が正しくありません。`);
  }
  const range = cwsAgentRangeForOp(op, index, null, CWS_AGENT_MAX_FORMAT_CELLS);
  return {
    type: "setStyle",
    ...range,
    style,
  };
}



function normalizeCwsAgentWorkbookTemplateOp(op, index) {
  const sourceSheets = Array.isArray(op?.sheets) ? op.sheets : [];
  if (!sourceSheets.length || sourceSheets.length > CWS_AGENT_MAX_WORKBOOK_TEMPLATE_SHEETS) {
    throw new Error(`Agent 操作 ${index + 1}: workbook template のシート数が正しくありません。`);
  }
  const mode = ["insert", "replaceBlankWorkbook", "replaceWorkbook"].includes(String(op.mode || ""))
    ? String(op.mode || "")
    : "insert";
  const replaceWorkbook = mode === "replaceWorkbook" || (mode === "replaceBlankWorkbook" && cwsAgentWorkbookLooksBlank());
  const usedNames = new Set(replaceWorkbook ? [] : (state.model?.sheets || []).map((sheet) => sheet.name));
  let totalCells = 0;
  const sheets = sourceSheets.map((source, sheetOffset) => {
    if (!source || typeof source !== "object" || Array.isArray(source)) {
      throw new Error(`Agent 操作 ${index + 1}: workbook template のシート定義が正しくありません。`);
    }
    const name = cwsAgentUniqueTemplateSheetName(source.name ?? source.sheetName ?? source.title, usedNames, sheetOffset + 1);
    usedNames.add(name);
    const rows = cwsAgentMatrixValues(source.rows || source.values || source.data);
    if (!rows.length) {
      throw new Error(`Agent 操作 ${index + 1}: ${name} のテンプレート内容が空です。`);
    }
    const width = Math.max(...rows.map((row) => row.length));
    const height = rows.length;
    totalCells += width * height;
    if (totalCells > CWS_AGENT_MAX_WORKBOOK_TEMPLATE_CELLS) {
      throw new Error(`Agent 操作 ${index + 1}: workbook template のセル数が多すぎます。`);
    }
    return {
      name,
      rows: rows.map((row) => row.concat(Array(Math.max(0, width - row.length)).fill(""))),
      columnWidths: cwsAgentTemplateDimensionList(source.columnWidths, width, CWS_AGENT_MIN_COLUMN_WIDTH, CWS_AGENT_MAX_COLUMN_WIDTH),
      rowHeights: cwsAgentTemplateDimensionList(source.rowHeights, height, CWS_AGENT_MIN_ROW_HEIGHT, CWS_AGENT_MAX_ROW_HEIGHT),
      titleRows: cwsAgentTemplateIndexList(source.titleRows, height),
      headerRows: cwsAgentTemplateIndexList(source.headerRows, height),
    };
  });
  return {
    type: "createWorkbookTemplate",
    templateId: String(op.templateId || op.id || "").slice(0, 120),
    mode,
    replaceWorkbook,
    sheets,
  };
}



function cwsAgentUniqueTemplateSheetName(value, usedNames, fallbackIndex) {
  const raw = String(value ?? "").trim().replace(/[:\\/?*\[\]]/g, "_").slice(0, 31);
  const base = raw || `Template${fallbackIndex}`;
  if (!usedNames.has(base)) return base;
  return uniqueSheetNameWithBase(base.slice(0, 25), [...usedNames]);
}



function cwsAgentTemplateDimensionList(value, maxItems, minValue, maxValue) {
  if (!Array.isArray(value)) return [];
  return value.slice(0, maxItems).map((item) => {
    const number = Number(item);
    return Number.isFinite(number) ? cwsAgentClampDimension(number, minValue, maxValue) : 0;
  });
}



function cwsAgentTemplateIndexList(value, maxIndex) {
  if (!Array.isArray(value)) return [];
  return [...new Set(value.map((item) => Number(item)).filter((item) => Number.isInteger(item) && item >= 1 && item <= maxIndex))];
}



function cwsAgentWorkbookLooksBlank() {
  const sheets = Array.isArray(state.model?.sheets) ? state.model.sheets : [];
  if (sheets.length !== 1) return false;
  const sheet = sheets[0];
  if (!sheet) return true;
  if ((sheet.merges || []).length || (sheet.images || []).length) return false;
  return !Object.values(sheet.cells || {}).some((cell) => (
    cell &&
    (cell.raw != null || cell.formula || cell.cached != null || cell.display || cell.richText?.length || cell.hyperlink || cell.note || cell.commentThread?.length)
  ));
}



function normalizeCwsAgentSetCellProtectionOp(op, index, aliasType = "") {
  const protection = cwsAgentCellProtectionPatch(op, aliasType);
  if (!Object.keys(protection).length) {
    throw new Error(`Agent 操作 ${index + 1} のセル保護設定が正しくありません。`);
  }
  const range = cwsAgentRangeForOp(op, index);
  return {
    type: "setCellProtection",
    ...range,
    protection,
  };
}



function normalizeCwsAgentClearCellProtectionOp(op, index) {
  const range = cwsAgentRangeForOp(op, index);
  return {
    type: "clearCellProtection",
    ...range,
  };
}



function normalizeCwsAgentSetSparklineOp(op, index) {
  const sourceText = String(op?.sourceRange ?? op?.dataRange ?? op?.valuesRange ?? op?.fromRange ?? op?.source ?? op?.data ?? op?.range ?? op?.address ?? op?.ref ?? "").trim();
  const source = cwsAgentSparklineSourceRangeForOp(op, index, sourceText);
  const axis = cwsAgentSparklineAxis(op?.axis ?? op?.orientation ?? op?.by);
  const type = cwsAgentSparklineType(op?.sparklineType ?? op?.sparkType ?? op?.kind ?? op?.chartType ?? op?.style ?? op?.sparklineStyle);
  const targetText = String(op?.targetRange ?? op?.outputRange ?? op?.locationRange ?? op?.destinationRange ?? op?.destRange ?? op?.target ?? op?.to ?? op?.location ?? "").trim();
  const parsedTarget = cwsAgentParseRangeAddress(targetText);
  const targetSheetIndex = cwsAgentSheetIndexForOp(op, parsedTarget?.sheetName || source.sheetName);
  if (targetSheetIndex !== source.sheetIndex) {
    throw new Error(`Agent 操作 ${index + 1} のスパークラインは同じシート内で指定してください。`);
  }

  const targetHeight = axis === "columns" ? 1 : source.bottom - source.top + 1;
  const targetWidth = axis === "columns" ? source.right - source.left + 1 : 1;
  const targetTop = cwsAgentPositiveInteger(op?.targetTop || op?.outputTop || op?.locationTop || op?.toTop || op?.targetRow) || parsedTarget?.top || source.top;
  const targetLeft = cwsAgentColumnValue(op?.targetLeft || op?.outputLeft || op?.locationLeft || op?.toLeft || op?.targetCol || op?.targetColumn) || parsedTarget?.left || (axis === "columns" ? source.left : source.right + 1);
  const targetBottom = cwsAgentPositiveInteger(op?.targetBottom || op?.outputBottom || op?.locationBottom || op?.toBottom) || parsedTarget?.bottom || (targetTop + targetHeight - 1);
  const targetRight = cwsAgentColumnValue(op?.targetRight || op?.outputRight || op?.locationRight || op?.toRight) || parsedTarget?.right || (targetLeft + targetWidth - 1);
  if (!targetTop || !targetLeft || targetBottom < targetTop || targetRight < targetLeft || targetBottom > EXCEL_MAX_ROWS || targetRight > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のスパークライン出力範囲が正しくありません。`);
  }
  if ((targetBottom - targetTop + 1) !== targetHeight || (targetRight - targetLeft + 1) !== targetWidth) {
    throw new Error(`Agent 操作 ${index + 1} のスパークライン出力範囲と元データ範囲が一致しません。`);
  }

  return {
    type: "setSparkline",
    sheetIndex: source.sheetIndex,
    sparklineType: type,
    axis,
    sourceTop: source.top,
    sourceLeft: source.left,
    sourceBottom: source.bottom,
    sourceRight: source.right,
    targetTop,
    targetLeft,
    targetBottom,
    targetRight,
  };
}



function normalizeCwsAgentClearSparklineOp(op, index) {
  const rangeOp = {
    ...op,
    range: op?.range ?? op?.targetRange ?? op?.outputRange ?? op?.locationRange ?? op?.address ?? op?.ref ?? op?.cell,
  };
  return {
    type: "clearSparkline",
    ...cwsAgentRangeForOp(rangeOp, index),
  };
}



function normalizeCwsAgentSetBordersOp(op, index) {
  const range = cwsAgentRangeForOp(op, index, null, CWS_AGENT_MAX_FORMAT_CELLS);
  return {
    type: "setBorders",
    ...range,
    preset: cwsAgentBorderPreset(op?.preset || op?.border || op?.borders || op?.style || op?.mode),
  };
}



function normalizeCwsAgentSetNumberFormatOp(op, index) {
  const range = cwsAgentRangeForOp(op, index, null, CWS_AGENT_MAX_FORMAT_CELLS);
  const numFmt = cwsAgentNumberFormat(op);
  if (numFmt === null) {
    throw new Error(`Agent 操作 ${index + 1} の表示形式が正しくありません。`);
  }
  return {
    type: "setNumberFormat",
    ...range,
    numFmt,
    label: cwsAgentNumberFormatLabel(numFmt),
  };
}



function normalizeCwsAgentSortRangeOp(op, index) {
  const range = cwsAgentSortRangeForOp(op, index);
  const sheet = state.model?.sheets?.[range.sheetIndex];
  const keyCol = cwsAgentSortKeyColumn(op, sheet, range);
  if (!keyCol || keyCol < range.left || keyCol > range.right) {
    throw new Error(`Agent 操作 ${index + 1} の並べ替えキー列が正しくありません。`);
  }
  return {
    type: "sortRange",
    ...range,
    keyCol,
    direction: cwsAgentSortDirection(op?.direction ?? op?.order ?? op?.sortOrder),
    hasHeader: cwsAgentHeaderMode(op?.hasHeader ?? op?.header ?? op?.headers),
  };
}

export {
  appendCwsAgentPreviewOpsProgressively,
  applyCwsAgentOps,
  captureCwsAgentTaskWorkbook,
  clearCwsAgentPreview,
  commitCwsAgentPreview,
  cwsAgentCloneRange,
  cwsAgentCloneRanges,
  cwsAgentContractAllowedRanges,
  cwsAgentContractOpTypeKey,
  cwsAgentContractRequiredOps,
  cwsAgentDelay,
  cwsAgentNormalizeContractRange,
  cwsAgentOpsMayChangeSheetStructure,
  cwsAgentPreviewForId,
  cwsAgentProgressivePreviewDelayMs,
  cwsAgentRangeForContractOp,
  cwsAgentRangeWithinAny,
  cwsAgentRestoreSelectionSnapshot,
  cwsAgentRestoreWorkbookRevisionSnapshot,
  cwsAgentSegmentCrossesBox,
  cwsAgentSetRangeValuesLookLikeTable,
  cwsAgentStyleCoversGeneratedHeader,
  cwsAgentStyleLooksLikeHeaderStyle,
  cwsAgentTemplateDimensionList,
  cwsAgentTemplateIndexList,
  cwsAgentUniqueTemplateSheetName,
  cwsAgentWorkbookLooksBlank,
  cwsAgentWorkbookRevisionSnapshot,
  normalizeCwsAgentClearCellProtectionOp,
  normalizeCwsAgentClearSparklineOp,
  normalizeCwsAgentOp,
  normalizeCwsAgentOpsBatch,
  normalizeCwsAgentSetBordersOp,
  normalizeCwsAgentSetCellOp,
  normalizeCwsAgentSetCellProtectionOp,
  normalizeCwsAgentSetNumberFormatOp,
  normalizeCwsAgentSetRangeOp,
  normalizeCwsAgentSetSparklineOp,
  normalizeCwsAgentSetStyleOp,
  normalizeCwsAgentSortRangeOp,
  normalizeCwsAgentTaskSnapshot,
  normalizeCwsAgentWorkbookTemplateOp,
  persistCwsAgentTaskWorkbook,
  previewCwsAgentOps,
  previewCwsAgentOpsProgressively,
  registerWebSheetFormulaPlugin,
  renderCwsAgentProgressivePreviewStep,
  resetCellStyleInternPools,
  restoreCwsAgentTaskWorkbook,
  setCwsAgentPreviewVisible,
  validateCwsAgentAllowedRangeContract,
  validateCwsAgentAppliedBranchGeometry,
  validateCwsAgentHeaderStyleScope,
  validateCwsAgentOpsContract,
  validateCwsAgentRequiredOpsContract,
  withCwsAgentPreviewMutation,
};
