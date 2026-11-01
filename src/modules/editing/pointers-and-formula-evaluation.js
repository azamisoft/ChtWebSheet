// Feature module; shared bindings are imported explicitly.
import {
  $,
  $formulaBar,
  $formulaExpandButton,
  $formulaInput,
  $horizontalScrollbar,
  $ribbonPanels,
  $selectedCellLabel,
  $selectionStatus,
  $sheetHost,
  COLUMN_HEADER_HEIGHT,
  DEFAULT_STATUS_BAR_ITEMS,
  EMU_PER_PIXEL,
  ErrorType,
  FORMULA_BAR_DEFAULT_EXPANDED_LINES,
  FORMULA_BAR_MAX_LINES,
  FORMULA_BAR_MIN_LINES,
  FORMULA_REFERENCE_COLORS,
  SELECTION_AUTOSCROLL_EDGE_PX,
  SELECTION_AUTOSCROLL_MAX_STEP_PX,
  SELECTION_AUTOSCROLL_MIN_STEP_PX,
  STATUS_BAR_ITEMS,
  STATUS_BAR_STORAGE_KEY,
  SimpleRangeValue,
  createIcons,
  formulaSpillAnchorCacheBySheet,
  icons,
  sheetBorderCellCacheBySheet,
  sheetMinimumExtentCacheBySheet,
  sheetPersistentNoteCellCacheBySheet,
  sheetUsedRangeCacheBySheet,
  state,
} from "../../app-core.js";
import {
  isAutoFilterHeaderCell,
  workbookRangeFromFormulaReference,
} from "../data/filter-and-text-tools.js";
import {
  sheetUsedRange,
} from "../data/sort-and-autofilter.js";
import {
  validateCellInput,
} from "../data/validation-and-consolidation.js";
import {
  isLineShapeObject,
} from "../drawing/shape-rendering.js";
import {
  mutateShapeObjectsForSheetGroup,
  pruneDetachedLineShapeConnectionsForIds,
  updateConnectedLinesForTargets,
} from "../drawing/transforms-and-connectors.js";
import {
  cellEditRichTextForCommit,
  chartSourceOwnerSheetIndexesForRange,
  normalizeCellEditRichTextForModel,
  rangesEqualLoose,
  refreshWorkbookChartsAfterSourceEdit,
  selectedWorksheetObjectSummary,
  syncAutoFillOptionsWithSelection,
} from "./autofill.js";
import {
  setRibbonUnderlinePreset,
  syncShapeFormatPane,
  syncShapeSelectionPane,
  underlinePresetForValue,
} from "./ink-and-clipboard.js";
import {
  scheduleSelectionKeyboardSinkFocus,
  selectedObjectSummary,
} from "./keyboard-and-paste.js";
import {
  applySelectionRange,
  canvasFont,
  canvasFontSize,
  canvasTextLineHeight,
  clearCellEditSelection,
  clipRangeToVisibleRange,
  columnSelectionRange,
  coordinateToPixels,
  copiedRangeForActiveSheet,
  editablePlainText,
  ensureCellElementForPoint,
  fontOptionFromCss,
  fontSizeOptionFromCss,
  imageRotationDegrees,
  isCopiedSourceRange,
  isPictureCropModeActive,
  rangesEqual,
  rawActiveSelectionRanges,
  rowSelectionRange,
  selectedImageGroups,
  selectedImageObject,
  selectedImageObjects,
  selectedInkStrokes,
  selectionToolbarPoint,
  setRibbonButtonActive,
  sheetColumnsRightToLeft,
  stopEditAutoScroll,
  syncDataRibbonControls,
  syncEditableCellAlignment,
  syncPictureCropSelectionState,
  unionImageBox,
  updateFormulaBarForCell,
  updateFormulaBarForSelection,
  updateHeaderSelectionFromClient,
  wrapCanvasText,
} from "./selection-and-cell-text.js";
import {
  captureCellsForHistory,
  captureSheetForHistory,
  captureSheetsForHistory,
  cellEffectiveCss,
  pushHistory,
  restoreSheetFromHistory,
  restoreSheetsFromHistory,
  storeCellModelForSheet,
} from "./structure-and-history.js";
import {
  formulaReferenceColorClass,
  formulaTextStartsFormula,
  formulaTextWithExcelAutoCompletion,
} from "../formula/autocomplete.js";
import {
  textSplitFormulaMatrix,
  webSheetFormulaError,
} from "../formula/calculations.js";
import {
  activeFormulaReferenceEditor,
  clearFormulaReferenceState,
  closeFormulaAutocomplete,
  formulaEditorText,
  formulaReferenceEditorSheetIndex,
  isFunctionArgumentReferenceEditor,
  updateFormulaReferenceUi,
} from "../formula/dialogs-and-names.js";
import {
  applyFormatPainterAtPoint,
  applyFormatPainterToRange,
  canvasRichLineMetrics,
  cellCanRecomputeDisplayText,
  cellShowsValidationDropdown,
  ensureCellModel,
  forEachCellInRange,
  measureRichLine,
  rangeIntersects,
  rangePixelRect,
  rangesIntersect,
  renderSheet,
  repaintActiveSheetCanvas,
  richTextRunsForDisplay,
  splitRichRunsIntoHardLines,
  wrapCanvasRichText,
} from "../rendering/grid-and-canvas.js";
import {
  closeCellContextMenu,
  positionPopupMenu,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  columnName,
  columnWidthToPixels,
  dateDetectionFormat,
  dateToExcelSerial,
  decodeRange,
  escapeAttr,
  escapeHtml,
  excelSerialToDate,
  isDateFormat,
  rowHeightToPixels,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  consumeCommandMenuKey,
  enhanceCommandMenuAccessKeys,
  parseFormulaReferences,
} from "../ui/printing-and-home-tools.js";
import {
  closeFileSideMenu,
  commitDataOperation,
  ensureSheetSize,
  parseCellKey,
  resetLazyScrollLocks,
  rowHeaderWidthForSheet,
  scheduleViewportRender,
  switchRibbonTab,
} from "../ui/ribbon-and-links.js";
import {
  closeSheetTabContextMenu,
  ensureEditableRangeForProtection,
  ensureEditableRangesForProtection,
  ensureHeaderFormatAllowed,
  ensureHeaderFormatAllowedForSheet,
  ensureSelectableRangeForProtection,
  ensureSheetProtectionPermission,
  updateSheetFooterUi,
} from "../ui/sheet-protection-and-translation.js";
import {
  syncFormulaWatchWindow,
} from "../ui/views-and-formula-tools.js";
import {
  buildFormulaEngine,
  formulaEngineContentForCurrentModel,
  formulaEngineSheetIndex,
  isSheetVisible,
  renderSheetTabs,
  renderWorkbook,
  sheetGroupActionIndexes,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  dateSerialFromPersistedValue,
  localDateFromPersistedParts,
  markWorkbookChanged,
  parsePersistedDateString,
} from "../workbook/import-model.js";
import {
  activeSheet,
  buildMergeMaps,
  byColLambdaFormulaMatrix,
  byRowLambdaFormulaMatrix,
  cellDisplayText,
  chooseColsFormulaMatrix,
  chooseRowsFormulaMatrix,
  displayValue,
  dropFormulaMatrix,
  excelCompatibleFormulaValue,
  expandFormulaMatrix,
  filterFormulaMatrix,
  forceSyncHorizontalScrollbarFromHost,
  formulaFilterIncludeMatrix,
  formulaLambdaEvaluate,
  formulaTruthyValue,
  getCellFormula,
  getSheetId,
  hstackFormulaMatrices,
  isElapsedTimeFormat,
  makeArrayLambdaFormulaMatrix,
  mapLambdaFormulaMatrix,
  reduceLambdaFormulaValue,
  scanLambdaFormulaMatrix,
  sequenceFormulaMatrix,
  shouldHideZeroValueForSheet,
  sortByFormulaMatrix,
  sortFormulaMatrix,
  takeFormulaMatrix,
  tocolFormulaMatrix,
  torowFormulaMatrix,
  transposeFormulaMatrix,
  uniqueFormulaMatrix,
  updateHorizontalScrollbar,
  visibleGridRange,
  vstackFormulaMatrices,
  wrapColsFormulaMatrix,
  wrapRowsFormulaMatrix,
} from "../workbook/saving-and-export.js";
import {
  applyTableEditBehaviors,
  parseFormulaFunctionCall,
  splitFormulaArguments,
  splitFormulaTopLevel,
  tableEditMayChangeMetadata,
} from "../workbook/tables-and-validation.js";
import {
  rangeCellCount,
} from "../workbook/views-and-macros.js";


function mutateShapeObjects(images, mutator, label, options = {}) {
  if (!images?.length) {
    setStatus(`${options.objectLabel || "図形"}を選択してください。`);
    return;
  }
  const sheet = activeSheet();
  if (mutateShapeObjectsForSheetGroup(images, mutator, label, options)) return;
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return;
  const sheetBefore = captureSheetForHistory(sheet);
  images.forEach((image) => mutator(image, sheet));
  if (options.updateConnectedLines) {
    updateConnectedLinesForTargets(sheet, images.filter((image) => image && !isLineShapeObject(image)).map((image) => image.id));
    pruneDetachedLineShapeConnectionsForIds(sheet, images.filter((image) => isLineShapeObject(image)).map((image) => image.id));
  }
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${selectedObjectSummary(images)} ${label}`,
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`${selectedObjectSummary(images)}に${label}を適用しました。`);
}



function startColumnHeaderResize(event) {
  startHeaderResize(event, "col", Number(event.currentTarget.dataset.col));
}



function startRowHeaderResize(event) {
  startHeaderResize(event, "row", Number(event.currentTarget.dataset.row));
}



function autoFitColumnHeaderFromDoubleClick(event) {
  autoFitHeaderFromDoubleClick(event, "col", Number(event.currentTarget.dataset.col));
}



function autoFitRowHeaderFromDoubleClick(event) {
  autoFitHeaderFromDoubleClick(event, "row", Number(event.currentTarget.dataset.row));
}



function autoFitHeaderFromDoubleClick(event, kind, index) {
  if (!index || !state.model) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  applyHeaderAutoFit(kind, selectedHeaderRangesForKind(kind, index));
}



function startHeaderResize(event, kind, index) {
  if (event.button !== 0 || !index) return;
  if (!ensureHeaderFormatAllowed(kind)) return;
  const sheet = activeSheet();
  const target = kind === "col" ? sheet.columns[index - 1] : sheet.rows[index - 1];
  if (!target) return;
  const sheetIndexes = activeSheetGroupIndexes();
  const groupedSheetIndexes = sheetIndexes.length > 1 ? sheetIndexes : [state.activeSheetIndex];
  for (const sheetIndex of groupedSheetIndexes) {
    const targetSheet = state.model?.sheets?.[sheetIndex];
    if (!targetSheet || !ensureHeaderFormatAllowedForSheet(targetSheet, kind)) return;
  }

  event.preventDefault();
  event.stopPropagation();
  const ranges = selectedHeaderRangesForKind(kind, index);
  const targets = headerResizeTargetsForRanges(sheet, kind, ranges);
  const startSize = kind === "col" ? target.width || sheet.defaultColumnWidth || 64 : target.height || sheet.defaultRowHeight || 20;
  state.headerResize = {
    kind,
    index,
    ranges,
    targets,
    sheetIndexes: groupedSheetIndexes,
    sheetsBefore: groupedSheetIndexes.length > 1 ? captureSheetsForHistory(groupedSheetIndexes) : null,
    startPointer: kind === "col" ? event.clientX : event.clientY,
    startSize,
  };
  document.body.classList.add("is-resizing-sheet");
}



function selectedHeaderRangesForKind(kind, index) {
  const isColumn = kind === "col" || kind === "column";
  const selectedRanges = activeSelectionRanges().filter((range) => (isColumn ? range.fullColumn : range.fullRow));
  const includesTarget = selectedRanges.some((range) =>
    isColumn ? index >= range.left && index <= range.right : index >= range.top && index <= range.bottom,
  );
  if (includesTarget) return selectedRanges.map((range) => ({ ...range }));
  return [isColumn ? columnSelectionRange(index, index) : rowSelectionRange(index, index)];
}



function headerResizeTargetsForRanges(sheet, kind, ranges) {
  const isColumn = kind === "col" || kind === "column";
  const seen = new Set();
  const targets = [];
  (ranges || []).forEach((range) => {
    const start = isColumn ? range.left : range.top;
    const end = isColumn ? range.right : range.bottom;
    for (let index = start; index <= end; index += 1) {
      if (seen.has(index)) continue;
      const target = isColumn ? sheet.columns[index - 1] : sheet.rows[index - 1];
      if (!target) continue;
      seen.add(index);
      targets.push({
        index,
        startSize: isColumn ? target.width || sheet.defaultColumnWidth || 64 : target.height || sheet.defaultRowHeight || 20,
      });
    }
  });
  return targets;
}



function updateHeaderResize(event) {
  const job = state.headerResize;
  if (!job || !state.model) return;

  const delta = (job.kind === "col" ? event.clientX : event.clientY) - job.startPointer;
  const minSize = job.kind === "col" ? 16 : 8;
  const maxSize = job.kind === "col" ? 640 : 360;
  const nextSize = Math.max(minSize, Math.min(maxSize, Math.round(job.startSize + delta)));
  const property = job.kind === "col" ? "width" : "height";
  let changed = false;
  (job.sheetIndexes || [state.activeSheetIndex]).forEach((sheetIndex) => {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) return;
    (job.targets || []).forEach((item) => {
      const target = job.kind === "col" ? sheet.columns[item.index - 1] : sheet.rows[item.index - 1];
      if (!target || target[property] === nextSize) return;
      target[property] = nextSize;
      changed = true;
    });
  });
  if (!changed) return;
  scheduleHeaderResizeRender();
}



function scheduleHeaderResizeRender() {
  if (state.headerResizeFrame) return;
  state.headerResizeFrame = requestAnimationFrame(() => {
    state.headerResizeFrame = null;
    renderSheet();
    updateHorizontalScrollbar();
  });
}



function finishHeaderResize() {
  const job = state.headerResize;
  if (!job) return;
  const sheet = activeSheet();
  const property = job.kind === "col" ? "width" : "height";
  const grouped = (job.sheetIndexes || []).length > 1 && job.sheetsBefore?.length;
  const changedDimensions = (job.targets || [])
    .map((item) => {
      const target = job.kind === "col" ? sheet.columns[item.index - 1] : sheet.rows[item.index - 1];
      const size = target?.[property];
      if (!size || size === item.startSize) return null;
      return {
        before: { kind: job.kind, index: item.index, size: item.startSize },
        after: { kind: job.kind, index: item.index, size },
      };
    })
    .filter(Boolean);
  const size = changedDimensions[0]?.after.size;
  const groupSize = grouped
    ? job.targets
        .map((item) => (job.kind === "col" ? sheet.columns[item.index - 1]?.width : sheet.rows[item.index - 1]?.height))
        .find((value) => Number.isFinite(Number(value)))
    : null;
  const groupChanged = grouped && headerResizeGroupChanged(job, property);
  state.headerResize = null;
  document.body.classList.remove("is-resizing-sheet");
  if (grouped) {
    if (groupChanged) {
      const selection = activeSelectionRange();
      pushHistory({
        sheetIndex: state.activeSheetIndex,
        sheetsBefore: job.sheetsBefore,
        sheetsAfter: captureSheetsForHistory(job.sheetsBefore.map((snapshot) => snapshot.index)),
        selectionBefore: selection,
        selectionAfter: selection,
        selectionBeforeRanges: activeSelectionRanges(),
        selectionAfterRanges: activeSelectionRanges(),
        label: `${rangesToLabel(job.ranges)} ${(job.sheetIndexes || []).length} シート${job.kind === "col" ? "列の幅" : "行の高さ"} 調整`,
      });
    }
    if (groupSize) {
      setStatus(`${(job.sheetIndexes || []).length} 個のシートの${rangesToLabel(job.ranges)} ${job.kind === "col" ? "列の幅" : "行の高さ"}を ${Math.round(groupSize)}px に調整しました。`);
    }
    return;
  }
  if (size) {
    const label = `${rangesToLabel(job.ranges)} ${job.kind === "col" ? "列宽" : "行高"}`;
    if (changedDimensions.length) {
      const selection = activeSelectionRange();
      pushHistory({
        sheetIndex: state.activeSheetIndex,
        dimensionsBefore: changedDimensions.map((item) => item.before),
        dimensionsAfter: changedDimensions.map((item) => item.after),
        selectionBefore: selection,
        selectionAfter: selection,
        selectionBeforeRanges: activeSelectionRanges(),
        selectionAfterRanges: activeSelectionRanges(),
        label: `${label} 調整`,
      });
    }
    setStatus(`${label} を ${Math.round(size)}px に調整しました。`);
  }
}



function headerResizeGroupChanged(job, property) {
  const beforeByIndex = new Map((job.sheetsBefore || []).map((snapshot) => [snapshot.index, snapshot.sheet]));
  return (job.sheetIndexes || []).some((sheetIndex) => {
    const sheet = state.model?.sheets?.[sheetIndex];
    const before = beforeByIndex.get(sheetIndex);
    if (!sheet || !before) return false;
    return (job.targets || []).some((item) => {
      const current = job.kind === "col" ? sheet.columns[item.index - 1]?.[property] : sheet.rows[item.index - 1]?.[property];
      const previous = job.kind === "col" ? before.columns?.[item.index - 1]?.[property] : before.rows?.[item.index - 1]?.[property];
      return Number(current) !== Number(previous);
    });
  });
}



function applyHeaderAutoFit(kind, ranges) {
  const isColumn = kind === "col" || kind === "column";
  if (!ranges?.length || !ensureHeaderFormatAllowed(kind)) return;
  const groupedSheetIndexes = activeSheetGroupEditIndexes();
  if (groupedSheetIndexes.length > 1 && applyHeaderAutoFitToSheetGroup(kind, ranges, groupedSheetIndexes)) return;

  const sheet = activeSheet();
  if (!sheet) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const changed = applyHeaderAutoFitToSheet(sheet, kind, ranges);
  if (!changed) return;
  const selection = activeSelectionRange();
  const clonedRanges = ranges.map((range) => ({ ...range }));
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: selection,
    selectionAfter: selection,
    selectionBeforeRanges: clonedRanges,
    selectionAfterRanges: clonedRanges,
    label: `${rangesToLabel(ranges)} ${isColumn ? "列の幅" : "行の高さ"} 自動調整`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  updateHorizontalScrollbar();
  setStatus(`${rangesToLabel(ranges)} の${isColumn ? "列の幅" : "行の高さ"}を自動調整しました。`);
}



function applyHeaderAutoFitToSheetGroup(kind, ranges, sheetIndexes) {
  const isColumn = kind === "col" || kind === "column";
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureHeaderFormatAllowedForSheet(sheet, kind)) return true;
    const targetRanges = materializedRangesForSheet(sheet, ranges, sheetIndex);
    if (targetRanges.length) items.push({ sheetIndex, sheet, ranges: targetRanges });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  let changed = 0;
  items.forEach((item) => {
    changed += applyHeaderAutoFitToSheet(item.sheet, kind, item.ranges);
  });
  if (!changed) return true;
  const selection = activeSelectionRange();
  const clonedRanges = ranges.map((range) => ({ ...range }));
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
    selectionBefore: selection,
    selectionAfter: selection,
    selectionBeforeRanges: clonedRanges,
    selectionAfterRanges: clonedRanges,
    label: `${rangesToLabel(ranges)} ${sheetIndexes.length} シート${isColumn ? "列の幅" : "行の高さ"} 自動調整`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  updateHorizontalScrollbar();
  setStatus(`${sheetIndexes.length} 個のシートの${rangesToLabel(ranges)}の${isColumn ? "列の幅" : "行の高さ"}を自動調整しました。`);
  return true;
}



function applyHeaderAutoFitToSheet(sheet, kind, ranges) {
  const isColumn = kind === "col" || kind === "column";
  const property = isColumn ? "width" : "height";
  const context = createHeaderAutoFitContext(sheet);
  const seen = new Set();
  let changed = 0;
  (ranges || []).forEach((range) => {
    const start = isColumn ? range.left : range.top;
    const end = isColumn ? range.right : range.bottom;
    for (let index = start; index <= end; index += 1) {
      if (!index || seen.has(index)) continue;
      seen.add(index);
      ensureSheetSize(sheet, isColumn ? sheet.rowCount : index, isColumn ? index : sheet.colCount);
      const target = isColumn ? sheet.columns[index - 1] : sheet.rows[index - 1];
      if (!target) continue;
      const nextSize = bestHeaderAutoFitSize(sheet, kind, index, context);
      if (!Number.isFinite(nextSize)) continue;
      if (Math.abs(Number(target[property] || 0) - nextSize) < 0.5 && target.hidden !== true) continue;
      target[property] = nextSize;
      target.hidden = false;
      changed += 1;
    }
  });
  return changed;
}



function bestHeaderAutoFitSize(sheet, kind, index, context = createHeaderAutoFitContext(sheet)) {
  return kind === "col" || kind === "column"
    ? bestColumnAutoFitWidth(sheet, index, context)
    : bestRowAutoFitHeight(sheet, index, context);
}



function createHeaderAutoFitContext(sheet) {
  const byRow = new Map();
  const byCol = new Map();
  Object.entries(sheet?.cells || {}).forEach(([key, cell]) => {
    const [row, col] = parseCellKey(key);
    if (!row || !col || !cell) return;
    const item = { row, col, cell };
    if (!byRow.has(row)) byRow.set(row, []);
    if (!byCol.has(col)) byCol.set(col, []);
    byRow.get(row).push(item);
    byCol.get(col).push(item);
  });
  return {
    byRow,
    byCol,
    mergeMaps: buildMergeMaps(sheet?.merges || []),
    measure: headerAutoFitMeasureContext(),
    sheetId: getSheetId(sheet?.name),
  };
}



function headerAutoFitMeasureContext() {
  if (typeof document === "undefined") return null;
  if (!headerAutoFitMeasureContext.context) {
    headerAutoFitMeasureContext.context = document.createElement("canvas").getContext?.("2d") || null;
  }
  return headerAutoFitMeasureContext.context;
}



function bestColumnAutoFitWidth(sheet, col, context) {
  const fallback = Number(sheet?.defaultColumnWidth || columnWidthToPixels(8.43));
  let best = fallback;
  (context.byCol.get(col) || []).forEach(({ row, cell }) => {
    if (sheet.rows[row - 1]?.hidden) return;
    const merge = mergedRangeForCell(sheet, row, col);
    if (merge && (merge.left !== col || merge.right !== col)) return;
    const width = autoFitCellTextWidth(sheet, context, row, col, cell);
    if (Number.isFinite(width)) best = Math.max(best, width);
  });
  return Math.max(16, Math.min(640, Math.ceil(best)));
}



function bestRowAutoFitHeight(sheet, row, context) {
  const fallback = Number(sheet?.defaultRowHeight || rowHeightToPixels(15));
  let best = fallback;
  (context.byRow.get(row) || []).forEach(({ col, cell }) => {
    if (sheet.columns[col - 1]?.hidden) return;
    const merge = mergedRangeForCell(sheet, row, col);
    if (merge && merge.top !== row) return;
    const height = autoFitCellTextHeight(sheet, context, row, col, cell, merge);
    if (Number.isFinite(height)) best = Math.max(best, height);
  });
  return Math.max(8, Math.min(360, Math.ceil(best)));
}



function autoFitCellDisplayText(sheet, context, row, col, cell) {
  const formula = cell?.formula || getCellFormula(context.sheetId, row, col);
  const display = state.showFormulas && formula ? formula : getDisplayForCell(sheet, context.sheetId, row, col);
  return display == null || display === "" ? "" : String(display);
}



function autoFitCellControlInset(sheet, row, col, cell) {
  const validation = cellShowsValidationDropdown(cell, sheet);
  const filter = isAutoFilterHeaderCell(sheet, row, col);
  return validation && filter ? 34 : validation || filter ? 18 : 0;
}



function autoFitCellTextWidth(sheet, context, row, col, cell) {
  const text = autoFitCellDisplayText(sheet, context, row, col, cell);
  if (!text) return null;
  const css = cellEffectiveCss(sheet, cell);
  const fontSize = canvasFontSize(css.fontSize);
  if (css.textRotation === "vertical") return Math.ceil(fontSize + 8);
  const measure = context.measure;
  if (!measure) return Math.ceil(String(text).length * Math.max(6, fontSize * 0.55) + 8);
  const richTextRuns = !state.showFormulas ? richTextRunsForDisplay(cell, text) : null;
  let width = 0;
  if (richTextRuns?.length) {
    splitRichRunsIntoHardLines(richTextRuns).forEach((line) => {
      width = Math.max(width, measureRichLine(measure, line, css));
    });
  } else {
    measure.font = canvasFont(css, fontSize);
    String(text)
      .split(/\r\n|\r|\n/)
      .forEach((line) => {
        width = Math.max(width, measure.measureText(line || " ").width);
      });
  }
  width += autoFitCellControlInset(sheet, row, col, cell);
  return Math.ceil(width + 8);
}



function autoFitCellTextHeight(sheet, context, row, col, cell, merge = null) {
  const text = autoFitCellDisplayText(sheet, context, row, col, cell);
  if (!text) return null;
  const css = cellEffectiveCss(sheet, cell);
  const fontSize = canvasFontSize(css.fontSize);
  const lineHeight = canvasTextLineHeight(fontSize);
  if (css.textRotation === "vertical") {
    const chars = Array.from(String(text).replace(/\r\n|\r|\n/g, ""));
    return Math.ceil(Math.max(lineHeight, chars.length * Math.max(10, fontSize * 1.08)) + 4);
  }
  const range = merge || { top: row, left: col, bottom: row, right: col };
  const rect = rangePixelRect(sheet, range);
  const maxTextWidth = Math.max(1, rect.width - 6 - autoFitCellControlInset(sheet, row, col, cell));
  const measure = context.measure;
  const richTextRuns = !state.showFormulas ? richTextRunsForDisplay(cell, text) : null;
  let lineCount = 1;
  if (richTextRuns?.length) {
    const lines = (css.whiteSpace === "pre-wrap" && measure)
      ? wrapCanvasRichText(measure, richTextRuns, css, maxTextWidth)
      : splitRichRunsIntoHardLines(richTextRuns);
    return Math.ceil(canvasRichLineMetrics(lines,css).totalHeight+4);
  } else if (css.whiteSpace === "pre-wrap" && measure) {
    lineCount = wrapCanvasText(measure, text, maxTextWidth, canvasFont(css, fontSize)).length;
  } else {
    lineCount = String(text).split(/\r\n|\r|\n/).length;
  }
  const phoneticLineHeight = cell?.showPhonetic && cell?.phonetics?.length ? Math.ceil(Math.max(7, Math.min(fontSize * 0.58, fontSize - 2)) * 1.1) : 0;
  return Math.ceil(Math.max(lineHeight, lineCount * lineHeight + phoneticLineHeight) + 4);
}



function extendCellRangeSelection(event) {
  const drag = state.selectionDrag;
  if (!state.isSelecting || !drag || (drag.kind !== "cell" && drag.kind !== "format-painter")) return;
  event.preventDefault();
  updateCellRangeSelectionFromClient(event.clientX, event.clientY, { clampToGrid: true });
  updateSelectionAutoScroll(event.clientX, event.clientY);
}



function updateCellRangeSelectionFromClient(clientX, clientY, options = {}) {
  const drag = state.selectionDrag;
  if (!state.isSelecting || !drag || (drag.kind !== "cell" && drag.kind !== "format-painter")) return false;
  const endpoint = cellPointFromClientPosition(clientX, clientY, options);
  if (!endpoint) return false;
  const range = selectionRangeFromPoints(drag.anchor, endpoint);
  if (!ensureSelectableRangeForProtection(activeSheet(), range)) return false;
  if (drag.kind === "format-painter" && (endpoint.row !== drag.anchor.row || endpoint.col !== drag.anchor.col)) {
    drag.moved = true;
  }
  applySelectionRange(range, endpoint, { additive: drag.additive, baseRanges: drag.baseRanges });
  state.suppressNextCellClick = true;
  updateSelectionUi();
  updateFormulaBarForSelection();
  return true;
}



function updateSelectionAutoScroll(clientX, clientY) {
  if (!selectionDragSupportsAutoScroll()) {
    stopSelectionAutoScroll();
    return;
  }
  const velocity = selectionDragAutoScrollVelocity(clientX, clientY);
  if (!velocity.x && !velocity.y) {
    stopSelectionAutoScroll();
    return;
  }

  const frame = state.selectionAutoScroll?.frame || null;
  state.selectionAutoScroll = { clientX, clientY, velocityX: velocity.x, velocityY: velocity.y, frame };
  if (!frame) {
    state.selectionAutoScroll.frame = requestAnimationFrame(runSelectionAutoScroll);
  }
}



function runSelectionAutoScroll() {
  const job = state.selectionAutoScroll;
  if (!job) return;
  job.frame = null;

  if (!selectionDragSupportsAutoScroll()) {
    stopSelectionAutoScroll();
    return;
  }

  const host = $sheetHost[0];
  if (!host) {
    stopSelectionAutoScroll();
    return;
  }

  const velocity = selectionDragAutoScrollVelocity(job.clientX, job.clientY);
  job.velocityX = velocity.x;
  job.velocityY = velocity.y;
  if (!velocity.x && !velocity.y) {
    stopSelectionAutoScroll();
    return;
  }

  const previousTop = host.scrollTop;
  const previousLeft = host.scrollLeft;
  const maxTop = Math.max(0, host.scrollHeight - host.clientHeight);
  const maxLeft = Math.max(0, host.scrollWidth - host.clientWidth);
  host.scrollTop = Math.max(0, Math.min(maxTop, previousTop + velocity.y));
  host.scrollLeft = Math.max(0, Math.min(maxLeft, previousLeft + velocity.x));

  const changed = host.scrollTop !== previousTop || host.scrollLeft !== previousLeft;
  if (changed) {
    forceSyncHorizontalScrollbarFromHost();
    scheduleViewportRender({ deferTrim: true });
  } else if ((velocity.y > 0 && !state.lazyScrollLocks.rows) || (velocity.x > 0 && !state.lazyScrollLocks.cols)) {
    scheduleViewportRender({ deferTrim: true });
  }

  if (state.selectionDrag?.kind === "col" || state.selectionDrag?.kind === "row") {
    updateHeaderSelectionFromClient(job.clientX, job.clientY, { clampToGrid: true });
  } else {
    updateCellRangeSelectionFromClient(job.clientX, job.clientY, { clampToGrid: true });
  }

  if (selectionAutoScrollCanContinue(host, velocity)) {
    job.frame = requestAnimationFrame(runSelectionAutoScroll);
  } else {
    state.selectionAutoScroll = null;
  }
}



function stopSelectionAutoScroll() {
  const frame = state.selectionAutoScroll?.frame;
  if (frame) cancelAnimationFrame(frame);
  state.selectionAutoScroll = null;
}



function selectionDragSupportsAutoScroll() {
  if (!state.isSelecting) return false;
  const kind = state.selectionDrag?.kind;
  return kind === "cell" || kind === "format-painter" || kind === "col" || kind === "row";
}



function selectionDragAutoScrollVelocity(clientX, clientY) {
  const velocity = selectionAutoScrollVelocity(clientX, clientY);
  if (state.selectionDrag?.kind === "col") return { x: velocity.x, y: 0 };
  if (state.selectionDrag?.kind === "row") return { x: 0, y: velocity.y };
  return velocity;
}



function selectionAutoScrollCanContinue(host, velocity) {
  const maxTop = Math.max(0, host.scrollHeight - host.clientHeight);
  const maxLeft = Math.max(0, host.scrollWidth - host.clientWidth);
  if (velocity.y < 0 && host.scrollTop > 0) return true;
  if (velocity.y > 0 && host.scrollTop < maxTop) return true;
  if (velocity.x < 0 && host.scrollLeft > 0) return true;
  if (velocity.x > 0 && host.scrollLeft < maxLeft) return true;
  if (velocity.y > 0 && !state.lazyScrollLocks.rows) return true;
  if (velocity.x > 0 && !state.lazyScrollLocks.cols) return true;
  return false;
}



function selectionAutoScrollVelocity(clientX, clientY) {
  const host = $sheetHost[0];
  if (!host || !state.model) return { x: 0, y: 0 };
  const rect = host.getBoundingClientRect();
  const sheet = activeSheet();
  const gridLeft = rect.left + rowHeaderWidthForSheet(sheet);
  const gridTop = rect.top + COLUMN_HEADER_HEIGHT;
  return {
    x: autoScrollDeltaForAxis(clientX, gridLeft, rect.right),
    y: autoScrollDeltaForAxis(clientY, gridTop, rect.bottom),
  };
}



function autoScrollDeltaForAxis(pointer, start, end) {
  if (end <= start) return 0;
  if (pointer < start + SELECTION_AUTOSCROLL_EDGE_PX) {
    return -autoScrollStep(start + SELECTION_AUTOSCROLL_EDGE_PX - pointer);
  }
  if (pointer > end - SELECTION_AUTOSCROLL_EDGE_PX) {
    return autoScrollStep(pointer - (end - SELECTION_AUTOSCROLL_EDGE_PX));
  }
  return 0;
}



function autoScrollStep(distance) {
  const ratio = Math.min(1, Math.max(0, distance) / SELECTION_AUTOSCROLL_EDGE_PX);
  return Math.round(SELECTION_AUTOSCROLL_MIN_STEP_PX + (SELECTION_AUTOSCROLL_MAX_STEP_PX - SELECTION_AUTOSCROLL_MIN_STEP_PX) * ratio);
}



function finishCellRangeSelection() {
  const drag = state.selectionDrag;
  const range = activeSelectionRange();
  if (drag?.kind === "format-painter") {
    const targetRange = drag.moved ? range : null;
    const targetPoint = drag.anchor;
    stopSelectionAutoScroll();
    state.isSelecting = false;
    state.selectionAnchor = null;
    state.selectionDrag = null;
    state.suppressNextCellClick = true;
    if (targetRange) applyFormatPainterToRange(targetRange);
    else applyFormatPainterAtPoint(targetPoint);
    return;
  }
  if (range && !isSingleCellRange(range)) {
    state.suppressNextCellClick = true;
  }
  stopSelectionAutoScroll();
  state.isSelecting = false;
  state.selectionAnchor = null;
  state.selectionDrag = null;
}



function cellPointFromElement(element) {
  const $cell = $(element);
  const row = Number($cell.data("row"));
  const col = Number($cell.data("col"));
  if (!row || !col) return null;
  const sheetIndex = Number($cell.data("sheetIndex"));
  return {
    sheetIndex: Number.isInteger(sheetIndex) ? sheetIndex : state.activeSheetIndex,
    row,
    col,
  };
}



function cellPointFromPointerEvent(event, options = {}) {
  if (!isCoordinateCellEvent(event, options)) return null;
  return cellPointFromClientPosition(event.clientX, event.clientY);
}



function formatPainterPointFromPointerEvent(event) {
  if (!state.model || !$sheetHost[0]) return null;
  const target = event.target;
  if (target !== $sheetHost[0] && !target?.closest?.(".sheet-grid")) return null;
  if (
    target.closest?.(
      [
        ".cell-validation-dropdown-button",
        ".sheet-image",
        ".sheet-image-link",
        ".sheet-image-selection-frame",
        ".sheet-image-resize-handle",
        ".grid-header",
        ".column-resize-handle",
        ".row-resize-handle",
        ".cell-annotation-editor",
        ".cell-comment-thread-popup",
        ".cell-note-popup",
      ].join(","),
    )
  ) {
    return null;
  }
  return cellPointFromClientPosition(event.clientX, event.clientY);
}



function isCoordinateCellEvent(event, options = {}) {
  if (!state.model || !$sheetHost[0]) return false;
  const target = event.target;
  if (target !== $sheetHost[0] && !target?.closest?.(".sheet-grid")) return false;
  if (!options.allowEditable && target.closest(".sheet-cell[contenteditable='true']")) return false;
  if (
    target.closest(
      [
        ".cell-validation-dropdown-button",
        ".sheet-image",
        ".sheet-image-link",
        ".sheet-image-selection-frame",
        ".sheet-image-resize-handle",
        ".grid-header",
        ".column-resize-handle",
        ".row-resize-handle",
        ".sheet-selection-frame",
        ".sheet-copy-frame",
        ".cell-annotation-editor",
        ".cell-comment-thread-popup",
        ".cell-note-popup",
      ].join(","),
    )
  ) {
    return false;
  }
  return true;
}



function cellPointFromClientPosition(clientX, clientY, options = {}) {
  const sheetPixel = sheetPixelFromClientPosition(clientX, clientY, options);
  if (!sheetPixel) return null;
  const row = sheetIndexFromPixel(activeSheet(), "row", sheetPixel.y);
  const col = sheetIndexFromPixel(activeSheet(), "col", sheetPixel.x);
  if (!row || !col) return null;
  return selectableCellPoint(row, col);
}



function sheetPixelFromClientPosition(clientX, clientY, options = {}) {
  if (!state.model) return null;
  const host = $sheetHost[0];
  if (!host) return null;
  const rect = host.getBoundingClientRect();
  const sheet = activeSheet();
  let pointX = clientX;
  let pointY = clientY;
  if (options.clampToGrid) {
    const minX = Math.min(rect.right - 1, rect.left + rowHeaderWidthForSheet(sheet) + 1);
    const maxX = Math.max(minX, rect.right - 1);
    const minY = Math.min(rect.bottom - 1, rect.top + COLUMN_HEADER_HEIGHT + 1);
    const maxY = Math.max(minY, rect.bottom - 1);
    pointX = Math.max(minX, Math.min(maxX, pointX));
    pointY = Math.max(minY, Math.min(maxY, pointY));
  } else if (clientX < rect.left || clientX > rect.right || clientY < rect.top || clientY > rect.bottom) {
    return null;
  }

  const x = pointX - rect.left + host.scrollLeft;
  const y = pointY - rect.top + host.scrollTop;
  if (x < rowHeaderWidthForSheet(sheet) || y < COLUMN_HEADER_HEIGHT) return null;
  return { x, y };
}



function nativeAnchorFromSheetPixel(sheet, axis, pixel) {
  const isColumn = axis === "col";
  const headerSize = isColumn ? rowHeaderWidthForSheet(sheet) : COLUMN_HEADER_HEIGHT;
  const items = isColumn ? sheet.columns : sheet.rows;
  const defaultSize = isColumn ? sheet.defaultColumnWidth : sheet.defaultRowHeight;
  const fallbackSize = defaultSize || (isColumn ? columnWidthToPixels(8.43) : rowHeightToPixels(15));
  const scale = sheet.displayScale || 1;
  if (isColumn && sheetColumnsRightToLeft(sheet)) {
    for (let col = 1; col <= sheet.colCount; col += 1) {
      const start = coordinateToPixels(sheet, "col", col - 1);
      const end = coordinateToPixels(sheet, "col", col);
      const visualLeft = Math.min(start, end);
      const visualRight = Math.max(start, end);
      if (pixel >= visualLeft && pixel <= visualRight) {
        return { index: col - 1, offsetEmu: Math.round(((pixel - visualLeft) / scale) * EMU_PER_PIXEL) };
      }
    }
    return { index: 0, offsetEmu: 0 };
  }
  let local = Math.max(0, Number(pixel) - headerSize);
  let index = 0;

  for (; index < items.length; index += 1) {
    const item = items[index];
    const size = item?.hidden ? 0 : item?.width || item?.height || fallbackSize;
    if (size <= 0) continue;
    if (local <= size) {
      return { index, offsetEmu: Math.round((local / scale) * EMU_PER_PIXEL) };
    }
    local -= size;
  }

  const extra = Math.max(0, Math.floor(local / Math.max(1, fallbackSize)));
  index += extra;
  const offset = local - extra * Math.max(1, fallbackSize);
  return { index, offsetEmu: Math.round((offset / scale) * EMU_PER_PIXEL) };
}



function sheetIndexFromPixel(sheet, axis, pixel) {
  const isColumn = axis === "col";
  const headerSize = isColumn ? rowHeaderWidthForSheet(sheet) : COLUMN_HEADER_HEIGHT;
  if (pixel < headerSize) return null;

  const items = isColumn ? sheet.columns : sheet.rows;
  const defaultSize = isColumn ? sheet.defaultColumnWidth : sheet.defaultRowHeight;
  const fallbackSize = defaultSize || (isColumn ? columnWidthToPixels(8.43) : rowHeightToPixels(15));
  if (isColumn && sheetColumnsRightToLeft(sheet)) {
    for (let col = 1; col <= sheet.colCount; col += 1) {
      const start = coordinateToPixels(sheet, "col", col - 1);
      const end = coordinateToPixels(sheet, "col", col);
      const visualLeft = Math.min(start, end);
      const visualRight = Math.max(start, end);
      if (pixel >= visualLeft && pixel <= visualRight) return col;
    }
    return null;
  }
  let position = headerSize;

  for (let index = 0; index < items.length; index += 1) {
    const item = items[index];
    const size = item?.hidden ? 0 : item?.width || item?.height || fallbackSize;
    if (size <= 0) continue;
    const next = position + size;
    if (pixel <= next) return index + 1;
    position = next;
  }

  return null;
}



function cellElementForPoint(point) {
  if (!point || point.sheetIndex !== state.activeSheetIndex) return null;
  return $sheetHost.find(`.sheet-cell[data-row="${point.row}"][data-col="${point.col}"]`)[0] || null;
}



function selectableCellPoint(row, col) {
  const sheet = activeSheet();
  const targetRow = Math.max(1, Number(row) || 1);
  const targetCol = Math.max(1, Number(col) || 1);
  const merge = findMergeForCell(sheet, targetRow, targetCol);
  return {
    sheetIndex: state.activeSheetIndex,
    row: merge?.top || targetRow,
    col: merge?.left || targetCol,
  };
}



function findMergeForCell(sheet, row, col) {
  if (!sheet?.merges?.length) return null;
  return mergeForCellFromMaps(buildMergeMaps(sheet.merges), row, col);
}



function mergeForCellFromMaps(mergeMaps, row, col) {
  if (!mergeMaps) return null;
  const key = cellKey(row, col);
  const directMerge = mergeMaps.master.get(key);
  if (directMerge) return directMerge;
  const masterKey = mergeMaps.coveredToMaster.get(key);
  return masterKey ? mergeMaps.master.get(masterKey) || null : null;
}



function normalizeRange(start, end) {
  if (!start || !end) return null;
  const sheetIndex = start.sheetIndex ?? state.activeSheetIndex;
  return {
    sheetIndex,
    top: Math.min(start.row, end.row),
    left: Math.min(start.col, end.col),
    bottom: Math.max(start.row, end.row),
    right: Math.max(start.col, end.col),
  };
}



function selectionRangeFromPoints(start, end) {
  return expandRangeForMerges(activeSheet(), normalizeRange(start, end));
}



function expandRangeForMerges(sheet, range) {
  if (!range || !sheet?.merges?.length) return range;

  const expanded = { ...range };
  let changed = true;
  while (changed) {
    changed = false;
    sheet.merges.forEach((merge) => {
      if (!rangesIntersect(expanded, merge)) return;
      const next = {
        ...expanded,
        top: Math.min(expanded.top, merge.top),
        left: Math.min(expanded.left, merge.left),
        bottom: Math.max(expanded.bottom, merge.bottom),
        right: Math.max(expanded.right, merge.right),
      };
      if (next.top !== expanded.top || next.left !== expanded.left || next.bottom !== expanded.bottom || next.right !== expanded.right) {
        Object.assign(expanded, next);
        changed = true;
      }
    });
  }
  return expanded;
}



function activeSelectionRange() {
  if (state.selectionRange?.sheetIndex === state.activeSheetIndex) {
    return materializeSelectionRange(state.selectionRange);
  }
  if (!state.selected || state.selected.sheetIndex !== state.activeSheetIndex) {
    return null;
  }
  return selectionRangeFromPoints(state.selected, state.selected);
}



function activeSelectionRanges() {
  const rawRanges = rawActiveSelectionRanges();
  if (rawRanges.length) {
    const ranges = rawRanges.map(materializeSelectionRange).filter(Boolean);
    if (ranges.length) return ranges;
  }
  const range = activeSelectionRange();
  return range ? [range] : [];
}



function activeFormulaReferenceEntries() {
  if (!state.model) return [];
  const editor = activeFormulaReferenceEditor();
  if (!editor) return [];
  const text = formulaEditorText(editor);
  const isArgumentInput = isFunctionArgumentReferenceEditor(editor);
  if (!isArgumentInput && !formulaTextStartsFormula(text)) return [];
  const baseSheetIndex = formulaReferenceEditorSheetIndex(editor);
  const baseSheet = state.model.sheets?.[baseSheetIndex] || activeSheet();
  const references = parseFormulaReferences(isArgumentInput ? `=${text}` : text, baseSheet);
  const entries = references
    .map((reference, index) => {
      const sheet = state.model.sheets?.[reference.sheetIndex];
      const materialized = materializeSelectionRangeForSheet(sheet, reference.range, reference.sheetIndex);
      if (!materialized) return null;
      return {
        ...materialized,
        colorIndex: index % FORMULA_REFERENCE_COLORS.length,
      };
    })
    .filter(Boolean);

  if (
    state.formulaReferenceRange &&
    !entries.some((entry) => entry.sheetIndex === state.formulaReferenceRange.sheetIndex && rangesEqualLoose(entry, state.formulaReferenceRange))
  ) {
    const sheet = state.model.sheets?.[state.formulaReferenceRange.sheetIndex];
    const materialized = materializeSelectionRangeForSheet(sheet, state.formulaReferenceRange, state.formulaReferenceRange.sheetIndex);
    if (materialized) {
      entries.push({
        ...materialized,
        colorIndex: entries.length % FORMULA_REFERENCE_COLORS.length,
      });
    }
  }

  return entries;
}



function activeFormulaReferenceEntriesForActiveSheet() {
  return activeFormulaReferenceEntries().filter((entry) => entry.sheetIndex === state.activeSheetIndex);
}



function activeFormulaReferenceEntryForCell(row, col) {
  return activeFormulaReferenceEntriesForActiveSheet().find((entry) => isCellInRange(entry, row, col)) || null;
}



function materializeSelectionRange(range) {
  if (!range || range.sheetIndex !== state.activeSheetIndex) return null;
  return materializeSelectionRangeForSheet(activeSheet(), range, state.activeSheetIndex);
}



function materializeSelectionRangeForSheet(sheet, range, sheetIndex = formulaEngineSheetIndex(sheet)) {
  if (!range || !sheet) return null;
  const materialized = { ...range };
  materialized.sheetIndex = sheetIndex;
  if (materialized.fullColumn) materialized.bottom = sheet.rowCount;
  if (materialized.fullRow) materialized.right = sheet.colCount;
  return materialized.fullColumn || materialized.fullRow ? materialized : expandRangeForMerges(sheet, materialized);
}



function activeSheetGroupIndexes() {
  return sheetGroupActionIndexes(state.activeSheetIndex)
    .filter((index) => state.model?.sheets?.[index] && isSheetVisible(state.model.sheets[index]));
}



function activeSheetGroupEditIndexes() {
  const indexes = activeSheetGroupIndexes();
  return indexes.length > 1 ? indexes : [];
}



function materializedRangesForSheet(sheet, ranges, sheetIndex = formulaEngineSheetIndex(sheet)) {
  return (ranges || [])
    .map((range) => materializeSelectionRangeForSheet(sheet, range, sheetIndex))
    .filter(Boolean);
}



function unionMaterializedRanges(ranges, sheetIndex = state.activeSheetIndex) {
  const materialized = (ranges || []).filter(Boolean);
  if (!materialized.length) return null;
  return {
    sheetIndex,
    top: Math.min(...materialized.map((range) => range.top)),
    left: Math.min(...materialized.map((range) => range.left)),
    bottom: Math.max(...materialized.map((range) => range.bottom)),
    right: Math.max(...materialized.map((range) => range.right)),
  };
}



function loadStatusBarItems() {
  try {
    const parsed = JSON.parse(localStorage.getItem(STATUS_BAR_STORAGE_KEY) || "{}");
    return Object.fromEntries(
      Object.entries(DEFAULT_STATUS_BAR_ITEMS).map(([key, value]) => [key, typeof parsed[key] === "boolean" ? parsed[key] : value]),
    );
  } catch {
    return { ...DEFAULT_STATUS_BAR_ITEMS };
  }
}



function saveStatusBarItems() {
  try {
    localStorage.setItem(STATUS_BAR_STORAGE_KEY, JSON.stringify(state.statusBarItems));
  } catch {
    // Ignore private-mode/localStorage failures; the visible setting still applies for this session.
  }
}



function updateSelectionStatusBar() {
  if (!$selectionStatus.length) return;
  const stats = selectionStatusStats();
  if (stats.selectedCellCount <= 1) {
    $selectionStatus.empty();
    return;
  }
  const items = [];
  const visible = state.statusBarItems || DEFAULT_STATUS_BAR_ITEMS;
  const add = (key, value) => {
    if (!visible[key] || value == null) return;
    const item = STATUS_BAR_ITEMS.find((entry) => entry.key === key);
    if (!item) return;
    items.push(`<span class="selection-status-item" data-status-stat="${escapeAttr(key)}">${escapeHtml(item.label)}：${escapeHtml(value)}</span>`);
  };

  add("average", stats.numericCount ? formatStatusBarNumber(stats.sum / stats.numericCount) : null);
  add("count", stats.count ? formatStatusBarNumber(stats.count) : null);
  add("numericCount", stats.numericCount ? formatStatusBarNumber(stats.numericCount) : null);
  add("min", stats.numericCount ? formatStatusBarNumber(stats.min) : null);
  add("max", stats.numericCount ? formatStatusBarNumber(stats.max) : null);
  add("sum", stats.numericCount ? formatStatusBarNumber(stats.sum) : null);
  $selectionStatus.html(items.join(""));
}



function scheduleSelectionUiSync() {
  if (state.deferredSelectionUiSyncTimer) {
    window.clearTimeout(state.deferredSelectionUiSyncTimer);
  }
  state.deferredSelectionUiSyncTimer = window.setTimeout(() => {
    state.deferredSelectionUiSyncTimer = null;
    if (!state.model) return;
    const sheet = activeSheet();
    withDisplayCalculationCache(sheet, () => {
      updateSelectionStatusBar();
      syncDataRibbonControls();
      syncFormulaWatchWindow();
    });
  }, 0);
}



function selectionStatusStats() {
  if (!state.model) return { selectedCellCount: 0, count: 0, numericCount: 0, sum: 0, min: null, max: null };
  const sheet = activeSheet();
  const ranges = activeSelectionRanges();
  const emptyStats = { selectedCellCount: 0, count: 0, numericCount: 0, sum: 0, min: null, max: null };
  if (!sheet || !ranges.length || state.selectedImage) return emptyStats;

  const sheetId = getSheetId(sheet.name);
  const mergeMaps = sheet?.merges?.length ? buildMergeMaps(sheet.merges) : null;
  const seen = new Set();
  const selectedCellCount = selectionStatusSelectedCellCount(ranges);
  if (selectedCellCount <= 1) return { ...emptyStats, selectedCellCount };
  let count = 0;
  let numericCount = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;

  Object.entries(sheet.cells || {}).forEach(([key, cell]) => {
    if (!statusBarCellMayHaveData(cell)) return;
    const point = cellPointFromKey(key);
    if (!point) return;
    const merge = mergeForCellFromMaps(mergeMaps, point.row, point.col);
    const targetRange = merge || { top: point.row, left: point.col, bottom: point.row, right: point.col };
    if (!ranges.some((range) => rangeIntersects(range, targetRange))) return;
    const masterRow = merge?.top || point.row;
    const masterCol = merge?.left || point.col;
    const masterKey = cellKey(masterRow, masterCol);
    if (seen.has(masterKey)) return;
    seen.add(masterKey);

    const masterCell = sheet.cells[masterKey] || cell;
    const value = statusBarCellValue(sheet, sheetId, masterRow, masterCol, masterCell);
    const display = statusBarCellDisplay(sheet, sheetId, masterRow, masterCol, masterCell);
    if (statusBarCellHasData(masterCell, value, display)) count += 1;
    if (typeof value === "number" && Number.isFinite(value)) {
      numericCount += 1;
      sum += value;
      min = Math.min(min, value);
      max = Math.max(max, value);
    }
  });

  return {
    selectedCellCount,
    count,
    numericCount,
    sum,
    min: numericCount ? min : null,
    max: numericCount ? max : null,
  };
}



function selectionStatusSelectedCellCount(ranges) {
  return (ranges || []).reduce((total, range) => total + rangeCellCount(range), 0);
}



function cellPointFromKey(key) {
  const [rowText, colText] = String(key || "").split(":");
  const row = Number(rowText);
  const col = Number(colText);
  return Number.isFinite(row) && Number.isFinite(col) && row > 0 && col > 0 ? { row, col } : null;
}



function statusBarCellMayHaveData(cell) {
  return Boolean(
    cell?.formula ||
    cell?.richText?.length ||
    (cell?.raw != null && cell.raw !== "") ||
    (cell?.cached != null && cell.cached !== "") ||
    (cell?.display != null && cell.display !== "")
  );
}



function statusBarCellValue(sheet, sheetId, row, col, cell) {
  if (state.hf && sheetId != null) {
    try {
      const value = state.hf.getCellValue({ sheet: sheetId, row: row - 1, col: col - 1 });
      if (value != null && typeof value !== "object") return value;
    } catch {
      // Fall back to cached model data below.
    }
  }
  if (!cell) return null;
  if (cell.formula && cell.cached != null) return cell.cached;
  if (cell.raw != null && !(typeof cell.raw === "string" && cell.raw.startsWith("="))) return cell.raw;
  if (cell.cached != null) return cell.cached;
  return null;
}



function statusBarCellDisplay(sheet, sheetId, row, col, cell) {
  if (!cell?.formula && cell?.display == null) return cell?.raw ?? cell?.cached ?? "";
  return getDisplayForCell(sheet, sheetId, row, col);
}



function statusBarCellHasData(cell, value, display) {
  if (value != null && value !== "") return true;
  if (cell?.formula) return true;
  return String(display ?? "").trim() !== "";
}



function formatStatusBarNumber(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return "";
  return new Intl.NumberFormat("ja-JP", { maximumFractionDigits: 10 }).format(Object.is(number, -0) ? 0 : number);
}



function updateStatusBarControlVisibility() {
  const visible = state.statusBarItems || DEFAULT_STATUS_BAR_ITEMS;
  document.querySelector(".sheet-view-buttons")?.classList.toggle("status-controlled-hidden", visible.viewShortcuts === false);
  document.querySelector("#zoomOutButton")?.classList.toggle("status-controlled-hidden", visible.zoomSlider === false);
  document.querySelector("#zoomSlider")?.classList.toggle("status-controlled-hidden", visible.zoomSlider === false);
  document.querySelector("#zoomInButton")?.classList.toggle("status-controlled-hidden", visible.zoomSlider === false);
  document.querySelector("#zoomLabel")?.classList.toggle("status-controlled-hidden", visible.zoom === false);
}



function openStatusBarContextMenu(event) {
  event.preventDefault();
  closeCellContextMenu();
  closeSheetTabContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeStatusBarContextMenu();
  const menu = document.createElement("div");
  menu.id = "statusBarContextMenu";
  menu.className = "status-bar-context-menu";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", "ステータス バーのユーザー設定");
  menu.innerHTML = STATUS_BAR_ITEMS.map((item) => {
    if (item.type === "separator") return '<div class="status-bar-menu-separator" role="separator"></div>';
    const checked = state.statusBarItems?.[item.key] !== false ? "✓" : "";
    return `<button class="status-bar-menu-item" type="button" role="menuitemcheckbox" aria-checked="${checked ? "true" : "false"}" data-status-bar-item="${escapeAttr(item.key)}"><span class="status-bar-menu-check" aria-hidden="true">${checked}</span><span>${escapeHtml(item.label)}</span></button>`;
  }).join("");
  enhanceCommandMenuAccessKeys(menu);
  document.body.append(menu);
  state.statusBarContextMenu = true;
  positionPopupMenu(menu, event.clientX, event.clientY);
  menu.querySelector("button")?.focus({ preventScroll: true });
}



function handleStatusBarMenuItem(event) {
  const key = event.currentTarget.dataset.statusBarItem;
  if (!key) return;
  state.statusBarItems[key] = state.statusBarItems[key] === false;
  saveStatusBarItems();
  updateStatusBarControlVisibility();
  updateSelectionStatusBar();
  const menu = event.currentTarget.closest("#statusBarContextMenu");
  if (menu) {
    const checked = state.statusBarItems[key] !== false;
    event.currentTarget.setAttribute("aria-checked", checked ? "true" : "false");
    event.currentTarget.querySelector(".status-bar-menu-check").textContent = checked ? "✓" : "";
  }
}



function closeStatusBarContextMenuOnOutsidePointer(event) {
  if (!state.statusBarContextMenu) return;
  if (event.target.closest?.("#statusBarContextMenu") || event.target.closest?.("#selectionStatus")) return;
  closeStatusBarContextMenu();
}



function closeStatusBarContextMenuOnEscape(event) {
  if (event.key === "Escape" && state.statusBarContextMenu) {
    consumeCommandMenuKey(event);
    closeStatusBarContextMenu();
  }
}



function closeStatusBarContextMenu() {
  document.querySelector("#statusBarContextMenu")?.remove();
  state.statusBarContextMenu = false;
}



function updateSelectionUi(options = {}) {
  const ranges = activeSelectionRanges();
  const imageObjects = selectedImageObjects();
  const inkStrokes = selectedInkStrokes();
  syncAutoFillOptionsWithSelection(ranges, imageObjects);
  const cropFocusChanged = syncPictureCropSelectionState();
  const formulaReferenceColorClasses = FORMULA_REFERENCE_COLORS.map((_color, index) => formulaReferenceColorClass(index)).join(" ");
  $sheetHost.find(".sheet-cell.is-selected, .sheet-cell.is-in-range").removeClass("is-selected is-in-range");
  $sheetHost.find(".sheet-cell.is-formula-reference").removeClass(`is-formula-reference ${formulaReferenceColorClasses}`);
  $sheetHost.find(".sheet-image.is-selected, .sheet-image-link.is-selected").removeClass("is-selected");
  if (!imageObjects.length || cropFocusChanged || state.selectedImage?.sheetIndex !== state.activeSheetIndex) {
    $sheetHost.find(".sheet-image-selection-frame").remove();
  }
  if (!imageObjects.length && state.shapePointEditIds?.length) {
    state.shapePointEditIds = [];
  }

  if (ranges.length) {
    $sheetHost.find(".sheet-cell[data-row][data-col]").each((_index, element) => {
      const row = Number(element.dataset.row);
      const col = Number(element.dataset.col);
      if (ranges.some((range) => isCellInRange(range, row, col))) {
        element.classList.add("is-in-range");
      }
      if (isActiveCell(row, col)) {
        element.classList.add("is-selected");
      }
    });
    $selectedCellLabel.text(ranges.map(rangeToLabel).join(","));
  } else if (imageObjects.length || inkStrokes.length) {
    imageObjects.forEach((image) => {
      $sheetHost.find(`[data-image-id="${cssEscape(image.id)}"]`).addClass("is-selected");
    });
    $selectedCellLabel.text(selectedWorksheetObjectSummary(imageObjects, inkStrokes));
  } else {
    $selectedCellLabel.text("A1");
  }
  const formulaReferences = activeFormulaReferenceEntriesForActiveSheet();
  if (formulaReferences.length) {
    $sheetHost.find(".sheet-cell[data-row][data-col]").each((_index, element) => {
      const row = Number(element.dataset.row);
      const col = Number(element.dataset.col);
      const entry = formulaReferences.find((item) => isCellInRange(item, row, col));
      if (entry) element.classList.add("is-formula-reference", formulaReferenceColorClass(entry.colorIndex));
    });
  }
  if (!options.skipStatusBar) updateSelectionStatusBar();
  updateSelectionFrames();
  updateValidationInputMessagePopup();
  updateContextualRibbonTabs();
  syncShapeFormatControls();
  syncShapeSelectionPane();
  scheduleSelectionKeyboardSinkFocus();
}



function updateValidationInputMessagePopup() {
  const grid = $sheetHost.find(".sheet-grid")[0];
  if (!grid || !state.model || state.selectedImage) return closeValidationInputMessagePopup();
  const point = selectionToolbarPoint();
  if (!point || point.sheetIndex !== state.activeSheetIndex) return closeValidationInputMessagePopup();
  const sheet = activeSheet();
  const cell = sheet.cells[cellKey(point.row, point.col)];
  const rule = cell?.validation;
  const title = String(rule?.inputTitle || "").trim();
  const message = String(rule?.inputMessage || "").trim();
  if (!rule?.showInputMessage || (!title && !message)) return closeValidationInputMessagePopup();
  const visibleRange = visibleGridRange(sheet);
  const merge = findMergeForCell(sheet, point.row, point.col) || { top: point.row, left: point.col, bottom: point.row, right: point.col };
  if (!rangeIntersects(merge, visibleRange)) return closeValidationInputMessagePopup();

  let popup = grid.querySelector("#cellValidationInputPopup");
  if (!popup) {
    popup = document.createElement("div");
    popup.id = "cellValidationInputPopup";
    popup.className = "cell-validation-input-popup";
    popup.setAttribute("role", "status");
    popup.setAttribute("aria-live", "polite");
    grid.append(popup);
  }
  popup.dataset.row = String(point.row);
  popup.dataset.col = String(point.col);
  popup.innerHTML = `
    ${title ? `<strong>${escapeHtml(title)}</strong>` : ""}
    ${message ? `<div>${escapeHtml(message).replace(/\n/g, "<br>")}</div>` : ""}
  `;
  positionValidationInputMessagePopup(popup, sheet, merge);
}



function closeValidationInputMessagePopup() {
  document.querySelector("#cellValidationInputPopup")?.remove();
}



function positionValidationInputMessagePopup(popup, sheet, range) {
  const host = $sheetHost[0];
  const rowHeaderWidth = rowHeaderWidthForSheet(sheet);
  const preferredLeft = coordinateToPixels(sheet, "col", range.right) + 6;
  const preferredTop = coordinateToPixels(sheet, "row", range.top - 1) + 2;
  if (!host) {
    popup.style.left = `${Math.max(rowHeaderWidth + 4, preferredLeft)}px`;
    popup.style.top = `${Math.max(COLUMN_HEADER_HEIGHT + 4, preferredTop)}px`;
    return;
  }
  const popupWidth = popup.offsetWidth || 240;
  const popupHeight = popup.offsetHeight || 80;
  const viewportLeft = host.scrollLeft + rowHeaderWidth + 4;
  const viewportRight = Math.max(viewportLeft, host.scrollLeft + host.clientWidth - popupWidth - 8);
  const viewportTop = host.scrollTop + COLUMN_HEADER_HEIGHT + 4;
  const viewportBottom = Math.max(viewportTop, host.scrollTop + host.clientHeight - popupHeight - 8);
  popup.style.left = `${Math.min(Math.max(preferredLeft, viewportLeft), viewportRight)}px`;
  popup.style.top = `${Math.min(Math.max(preferredTop, viewportTop), viewportBottom)}px`;
}



function updateContextualRibbonTabs() {
  const objects = selectedImageObjects();
  const objectActive = objects.length > 0;
  const tab = document.querySelector('[data-ribbon-tab="shape-format"]');
  if (!tab) return;
  syncContextualObjectFormatRibbon(objects);
  tab.classList.toggle("hidden", !objectActive);
  if (!objectActive && tab.classList.contains("is-active")) {
    switchRibbonTab("home");
  }
}



function activateObjectFormatRibbonForSelection() {
  const objects = selectedImageObjects();
  if (!objects.length) return false;
  const tab = document.querySelector('[data-ribbon-tab="shape-format"]');
  if (tab) tab.classList.remove("hidden");
  syncContextualObjectFormatRibbon(objects);
  switchRibbonTab("shape-format");
  return true;
}



function selectedObjectFormatMode(objects = selectedImageObjects()) {
  if (!objects.length) return "none";
  if (objects.every((image) => image?.chart?.series?.length)) return "chart";
  return objects.every((image) => !image?.shape) ? "picture" : "shape";
}



function syncContextualObjectFormatRibbon(objects = selectedImageObjects()) {
  const mode = selectedObjectFormatMode(objects);
  const pictureMode = mode === "picture";
  const chartMode = mode === "chart";
  const tab = document.querySelector('[data-ribbon-tab="shape-format"]');
  if (tab) tab.textContent = chartMode ? "グラフの書式" : pictureMode ? "図の形式" : "図形の書式";
  const panel = document.querySelector('[data-ribbon-panel="shape-format"]');
  if (panel) {
    panel.classList.toggle("is-picture-format", pictureMode);
    panel.classList.toggle("is-chart-format", chartMode);
    panel.classList.toggle("is-shape-format", mode === "shape");
  }
  const styleLabel = document.querySelector(".ribbon-shape-style-group .ribbon-group-label");
  if (styleLabel) styleLabel.textContent = chartMode ? "グラフのスタイル" : pictureMode ? "図のスタイル" : "図形のスタイル";
  const gallery = document.querySelector(".ribbon-shape-style-gallery");
  if (gallery) gallery.setAttribute("aria-label", chartMode ? "グラフのスタイル" : pictureMode ? "図のスタイル" : "図形のスタイル");
  const insertGroup = document.querySelector(".ribbon-shape-insert-group");
  if (insertGroup) insertGroup.classList.toggle("hidden", pictureMode || chartMode);
  const textGroup = document.querySelector(".ribbon-shape-text-group");
  if (textGroup) textGroup.classList.toggle("hidden", pictureMode || chartMode);
  const fillButton = document.querySelector("[data-shape-fill-menu]");
  if (fillButton) {
    fillButton.title = chartMode ? "グラフ エリアの塗りつぶし" : pictureMode ? "図の塗りつぶし" : "図形の塗りつぶし";
    fillButton.toggleAttribute("disabled", pictureMode || chartMode);
  }
  const outlineButton = document.querySelector("[data-shape-outline-menu]");
  if (outlineButton) outlineButton.title = chartMode ? "グラフの枠線" : pictureMode ? "図の枠線" : "図形の枠線";
  const effectsButton = document.querySelector("[data-shape-effects-menu]");
  if (effectsButton) effectsButton.title = chartMode ? "グラフの効果" : pictureMode ? "図の効果" : "図形の効果";
  const cropButton = document.querySelector('[data-ribbon-command="picture-crop"]');
  if (cropButton) {
    cropButton.classList.toggle("hidden", !pictureMode);
    cropButton.toggleAttribute("disabled", !pictureMode);
    cropButton.classList.toggle("is-active", Boolean(pictureMode && isPictureCropModeActive()));
  }
}



function syncShapeFormatControls() {
  const image = selectedImageObject();
  const objects = selectedImageObjects();
  const group = selectedImageGroups(objects)[0];
  const widthInput = document.querySelector('[data-shape-size="width"]');
  const heightInput = document.querySelector('[data-shape-size="height"]');
  const rotationInput = document.querySelector("[data-shape-rotation]");
  if (!widthInput || !heightInput) return;
  if (!image) {
    widthInput.value = "";
    heightInput.value = "";
    if (rotationInput) rotationInput.value = "";
    syncShapeFormatPane();
    return;
  }
  const groupBox = group && objects.length === group.images.length ? unionImageBox(activeSheet(), group.images) : null;
  widthInput.value = String(Math.round(Number(groupBox?.width || image.ext?.width || 0)));
  heightInput.value = String(Math.round(Number(groupBox?.height || image.ext?.height || 0)));
  if (rotationInput) rotationInput.value = String(Math.round(imageRotationDegrees(image)));
  if (image.shape) syncRibbonControlsForShape(image);
  syncShapeFormatPane();
}



function syncRibbonControlsForShape(image) {
  if (!image?.shape || isLineShapeObject(image)) return;
  const shape = image.shape || {};
  const fontOption = fontOptionFromCss(shape.textFontFamily);
  if (fontOption) $ribbonPanels.find('[data-ribbon-style="fontFamily"]').val(fontOption);
  const sizeOption = fontSizeOptionFromCss(shape.textFontSize ? `${shape.textFontSize}pt` : "");
  if (sizeOption) $ribbonPanels.find('[data-ribbon-style="fontSize"]').val(sizeOption);
  $ribbonPanels.find('[data-ribbon-style="numberFormat"]').val("general");
  setRibbonButtonActive("bold", Boolean(shape.textBold));
  setRibbonButtonActive("italic", Boolean(shape.textItalic));
  const underlinePreset = underlinePresetForValue(shape.textUnderline);
  setRibbonButtonActive("underline", Boolean(underlinePreset));
  if (underlinePreset) setRibbonUnderlinePreset(underlinePreset.value);
  setRibbonButtonActive("align-left", shape.textAlign === "left");
  setRibbonButtonActive("align-center", !shape.textAlign || shape.textAlign === "center");
  setRibbonButtonActive("align-right", shape.textAlign === "right");
  setRibbonButtonActive("valign-top", shape.textVAlign === "top");
  setRibbonButtonActive("valign-middle", !shape.textVAlign || shape.textVAlign === "middle");
  setRibbonButtonActive("valign-bottom", shape.textVAlign === "bottom");
  setRibbonButtonActive("wrap-text", false);
  setRibbonButtonActive("merge-center", false);
}



function updateSelectionFrames() {
  const grid = $sheetHost.find(".sheet-grid")[0];
  if (!grid || !state.model) return;

  const activeRanges = activeSelectionRanges();
  const copyRange = copiedRangeForActiveSheet();
  const formulaReferenceRanges = activeFormulaReferenceEntriesForActiveSheet();
  updateRangeFrameElements(grid, ".sheet-selection-frame", activeRanges, "sheet-selection-frame", isCopiedSourceRange);
  updateRangeFrameElements(grid, ".sheet-formula-reference-frame", formulaReferenceRanges, "sheet-formula-reference-frame");
  const shouldShowCopyFrame = copyRange && !activeRanges.some((range) => rangesEqual(copyRange, range));
  updateRangeFrameElements(grid, ".sheet-copy-frame", shouldShowCopyFrame ? [copyRange] : [], "sheet-copy-frame");
}



function updateRangeFrameElements(grid, selector, ranges, className, copySourcePredicate = () => false) {
  const existing = Array.from(grid.querySelectorAll(selector));
  const sheet = activeSheet();
  const visibleRange = visibleGridRange(sheet);
  const visibleRanges = (ranges || [])
    .map((range) => {
      const frameRange = clipRangeToVisibleRange(range, visibleRange);
      return frameRange ? { range, frameRange } : null;
    })
    .filter(Boolean);
  const isMultiRangeSelection = className === "sheet-selection-frame" && (ranges || []).length > 1;

  visibleRanges.forEach(({ range, frameRange }, index) => {
    const frame = existing[index] || document.createElement("div");
    const isCopySource = copySourcePredicate(range);
    const formulaColorClass = className === "sheet-formula-reference-frame" ? ` ${formulaReferenceColorClass(range.colorIndex)}` : "";
    frame.className = `${className}${formulaColorClass}${isCopySource ? " is-copy-source" : ""}${isMultiRangeSelection && !isCopySource ? " is-multi-range" : ""}`;
    frame.setAttribute("aria-hidden", "true");
    frame.dataset.rangeIndex = String(index);
    const hideHandle = isCopySource || isMultiRangeSelection || frameRange.bottom !== range.bottom || frameRange.right !== range.right;
    syncRangeFrameHandle(frame, className, hideHandle);
    positionRangeFrameElement(frame, sheet, range, frameRange);
    if (!existing[index]) grid.append(frame);
  });
  existing.slice(visibleRanges.length).forEach((frame) => frame.remove());
}



function syncRangeFrameHandle(frame, className, isCopySource) {
  const shouldHaveHandle = className === "sheet-selection-frame" && !isCopySource;
  const existing = frame.querySelector(":scope > .sheet-fill-handle");
  if (shouldHaveHandle && !existing) {
    const handle = document.createElement("span");
    handle.className = "sheet-fill-handle";
    handle.setAttribute("aria-hidden", "true");
    frame.append(handle);
  } else if (!shouldHaveHandle && existing) {
    existing.remove();
  }
}



function positionRangeFrameElement(frame, sheet, range, frameRange = range) {
  const rect = rangePixelRect(sheet, frameRange);
  frame.style.left = `${rect.left}px`;
  frame.style.top = `${rect.top}px`;
  frame.style.width = `${Math.max(1, rect.width)}px`;
  frame.style.height = `${Math.max(1, rect.height)}px`;
  frame.style.borderTopColor = frameRange.top > range.top ? "transparent" : "";
  frame.style.borderLeftColor = frameRange.left > range.left ? "transparent" : "";
  frame.style.borderBottomColor = frameRange.bottom < range.bottom ? "transparent" : "";
  frame.style.borderRightColor = frameRange.right < range.right ? "transparent" : "";
}



function isActiveCell(row, col) {
  return state.selected?.sheetIndex === state.activeSheetIndex && state.selected.row === row && state.selected.col === col;
}



function cssEscape(value) {
  return window.CSS?.escape ? window.CSS.escape(String(value)) : String(value).replace(/["\\]/g, "\\$&");
}



function isCellInActiveSelection(row, col) {
  return activeSelectionRanges().some((range) => isCellInRange(range, row, col));
}



function isCellInFormulaReferenceRange(row, col) {
  return Boolean(activeFormulaReferenceEntryForCell(row, col));
}



function isCellInRange(range, row, col) {
  return row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;
}



function isSingleCellRange(range) {
  return !range || (range.top === range.bottom && range.left === range.right);
}



function rangeToLabel(range) {
  if (!range) return "A1";
  if (range.fullColumn) return `${columnName(range.left)}:${columnName(range.right)}`;
  if (range.fullRow) return `${range.top}:${range.bottom}`;
  const start = `${columnName(range.left)}${range.top}`;
  const end = `${columnName(range.right)}${range.bottom}`;
  return start === end ? start : `${start}:${end}`;
}



function rangesToLabel(ranges) {
  return (ranges || []).map(rangeToLabel).join(",") || "A1";
}



function formulaBarLineHeight() {
  const input = $formulaInput[0];
  if (!input) return 18;
  const lineHeight = Number.parseFloat(getComputedStyle(input).lineHeight);
  return Number.isFinite(lineHeight) && lineHeight > 0 ? lineHeight : 18;
}



function setFormulaBarLines(lines, options = {}) {
  const nextLines = Math.max(FORMULA_BAR_MIN_LINES, Math.min(FORMULA_BAR_MAX_LINES, Math.round(Number(lines) || FORMULA_BAR_MIN_LINES)));
  state.formulaBarLines = nextLines;
  if (nextLines > FORMULA_BAR_MIN_LINES) {
    state.formulaBarExpandedLines = nextLines;
  } else if (options.resetExpandedLines) {
    state.formulaBarExpandedLines = FORMULA_BAR_DEFAULT_EXPANDED_LINES;
  }

  $formulaBar.css("--formula-bar-lines", String(nextLines));
  $formulaBar.toggleClass("is-expanded", nextLines > FORMULA_BAR_MIN_LINES);
  $formulaInput.attr("rows", String(nextLines));
  syncFormulaExpandButton();
  updateHorizontalScrollbar();
  updateSelectionUi();
}



function syncFormulaExpandButton() {
  const expanded = state.formulaBarLines > FORMULA_BAR_MIN_LINES;
  const previousExpanded = $formulaExpandButton.attr("aria-expanded") === "true";
  $formulaExpandButton
    .attr("aria-expanded", expanded ? "true" : "false")
    .attr("aria-label", expanded ? "数式バーを折りたたむ" : "数式バーを展開");
  if (previousExpanded !== expanded || !$formulaExpandButton.find("svg").length) {
    $formulaExpandButton.html(`<i data-lucide="${expanded ? "chevron-up" : "chevron-down"}"></i>`);
    createIcons({ icons });
  }
}



function toggleFormulaBarLines() {
  if (state.formulaBarLines > FORMULA_BAR_MIN_LINES) {
    setFormulaBarLines(FORMULA_BAR_MIN_LINES);
  } else {
    setFormulaBarLines(state.formulaBarExpandedLines || FORMULA_BAR_DEFAULT_EXPANDED_LINES);
  }
}



function startFormulaBarResize(event) {
  if (event.button !== 0) return;
  event.preventDefault();
  event.stopPropagation();
  state.formulaBarResize = {
    pointerId: event.pointerId,
    startY: event.clientY,
    startLines: state.formulaBarLines || FORMULA_BAR_MIN_LINES,
    lineHeight: formulaBarLineHeight(),
  };
  document.body.classList.add("is-resizing-formula-bar");
  event.currentTarget.setPointerCapture?.(event.pointerId);
}



function updateFormulaBarResize(event) {
  const resize = state.formulaBarResize;
  if (!resize) return;
  const lineDelta = Math.round((event.clientY - resize.startY) / resize.lineHeight);
  setFormulaBarLines(resize.startLines + lineDelta);
}



function finishFormulaBarResize() {
  if (!state.formulaBarResize) return;
  state.formulaBarResize = null;
  document.body.classList.remove("is-resizing-formula-bar");
}



function commitEditableCellToSelection(cellElement) {
  closeFormulaAutocomplete({ restoreFocus: false });
  clearFormulaReferenceState({ update: false });
  const $cell = $(cellElement);
  const row = Number($cell.data("row"));
  const col = Number($cell.data("col"));
  const editedValue = editablePlainText(cellElement);
  const beforeEditTextData = $cell.data("beforeEditText");
  const value = beforeEditTextData !== undefined && editedValue === String(beforeEditTextData)
    ? String($cell.data("beforeEdit") ?? "")
    : editedValue;
  const richText = cellEditRichTextForCommit(cellElement, value);
  $cell.data("skipBlurCommit", true);
  cellElement.removeAttribute("contenteditable");
  $cell.removeData("beforeEdit");
  $cell.removeData("beforeEditText");
  $cell.removeData("beforeRichText");
  clearCellEditSelection(cellElement);
  stopEditAutoScroll();
  state.deferredViewportRender = false;
  commitCellValueToSelection(row, col, value, { preserveWhitespace: true, richText });
}



function beginFormulaBarEdit() {
  if (!state.model || !state.selected || state.selected.sheetIndex !== state.activeSheetIndex || $formulaInput.prop("disabled")) {
    if (state.formulaReferenceSheetNavigation?.sourceSheetIndex != null && state.formulaEdit) {
      updateFormulaReferenceUi();
      repaintActiveSheetCanvas();
    }
    return;
  }

  if (state.formulaReferenceSheetNavigation?.sourceSheetIndex != null && state.formulaEdit) {
    updateFormulaReferenceUi();
    repaintActiveSheetCanvas();
    return;
  }

  const { row, col } = state.selected;
  if (!ensureEditableRangeForProtection(activeSheet(), selectionRangeFromPoints(state.selected, state.selected), "編集")) return;
  if (
    state.formulaEdit?.sheetIndex === state.activeSheetIndex &&
    state.formulaEdit.row === row &&
    state.formulaEdit.col === col
  ) {
    markFormulaEditCell();
    syncFormulaEditPreview();
    repaintActiveSheetCanvas();
    return;
  }

  state.formulaEdit = {
    sheetIndex: state.activeSheetIndex,
    row,
    col,
    before: getCellRawInput(activeSheet(), row, col),
    beforeText: getCellEditInput(activeSheet(), row, col),
  };
  markFormulaEditCell();
  syncFormulaEditPreview();
  repaintActiveSheetCanvas();
  updateFormulaReferenceUi();
}



function markFormulaEditCell() {
  const edit = state.formulaEdit;
  if (!edit || edit.sheetIndex !== state.activeSheetIndex) return;
  const cell = formulaEditCellElement() || ensureCellElementForPoint(edit);
  if (!cell) return;
  cell.setAttribute("contenteditable", "true");
  cell.classList.add("is-formula-editing");
  syncEditableCellAlignment(cell);
}



function formulaEditCellElement() {
  const edit = state.formulaEdit;
  if (!edit) return null;
  if (edit.sheetIndex !== state.activeSheetIndex) return null;
  return $sheetHost.find(`.sheet-cell[data-row="${edit.row}"][data-col="${edit.col}"]`)[0] || null;
}



function syncFormulaEditPreview() {
  const cell = formulaEditCellElement();
  if (!cell) return;
  const target = cell.querySelector(".cell-content") || cell;
  target.textContent = $formulaInput.val();
  syncEditableCellAlignment(cell);
}



function cancelFormulaBarEdit() {
  const edit = state.formulaEdit;
  if (!edit) return;
  closeFormulaAutocomplete({ restoreFocus: false });
  clearFormulaReferenceState({ update: false });
  state.formulaEdit = null;
  $formulaInput.val(edit.before);
  if (edit.sheetIndex !== state.activeSheetIndex) {
    activateSheetForFormulaEditResult(edit);
    renderWorkbook();
  } else {
    renderSheet();
  }
  updateSelectionUi();
  updateFormulaBarForCell(edit.row, edit.col);
  $formulaInput.blur();
  setStatus("今回の入力をキャンセルしました。");
}



function commitFormulaBar(options = {}) {
  if (!state.selected && !state.formulaEdit) return;
  closeFormulaAutocomplete({ restoreFocus: false });
  clearFormulaReferenceState({ update: false });
  const edit = state.formulaEdit || state.selected;
  const editedValue = $formulaInput.val();
  const editSheet = state.model?.sheets?.[edit.sheetIndex];
  const before = edit.before ?? getCellRawInput(editSheet, edit.row, edit.col);
  const beforeText = edit.beforeText ?? getCellEditInput(editSheet, edit.row, edit.col);
  const value = String(editedValue) === String(beforeText) ? before : editedValue;
  state.formulaEdit = null;
  const changedSheet = edit.sheetIndex !== state.activeSheetIndex;
  if (edit.sheetIndex !== state.activeSheetIndex) {
    activateSheetForFormulaEditResult(edit);
  }
  if (options.fillSelection) commitCellValueToSelection(edit.row, edit.col, value, { preserveWhitespace: true });
  else commitCellValue(edit.row, edit.col, value, { preserveWhitespace: true });
  if (changedSheet) {
    renderSheetTabs();
    updateSheetFooterUi();
  }
  $formulaInput.blur();
}



function activateSheetForFormulaEditResult(edit) {
  if (!edit || !state.model?.sheets?.[edit.sheetIndex]) return;
  state.activeSheetIndex = edit.sheetIndex;
  state.selectedSheetIndexes = [];
  state.selected = { sheetIndex: edit.sheetIndex, row: edit.row, col: edit.col };
  state.selectionRange = selectionRangeFromPoints(state.selected, state.selected);
  state.selectionRanges = [state.selectionRange];
  state.selectionAnchor = state.selected;
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  resetLazyScrollLocks();
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
}



function commitActiveCellValueToSelection() {
  const point = selectionToolbarPoint();
  if (!point) return false;
  const value = getCellRawInput(activeSheet(), point.row, point.col);
  commitCellValueToSelection(point.row, point.col, value, { fromActiveCell: true });
  return true;
}



function commitCellValueToSelection(activeRow, activeCol, value, options = {}) {
  if (!options.skipSheetGroup && commitCellValueToSheetGroup(activeRow, activeCol, value, options)) return;
  const sheet = activeSheet();
  const ranges = activeSelectionRanges().length ? activeSelectionRanges() : [
    normalizeRange(
      { sheetIndex: state.activeSheetIndex, row: activeRow, col: activeCol },
      { sheetIndex: state.activeSheetIndex, row: activeRow, col: activeCol },
    ),
  ];
  const targets = uniqueEditableTargetsForRanges(sheet, ranges);
  if (targets.length <= 1 && targets[0]?.row === activeRow && targets[0]?.col === activeCol) {
    commitCellValue(activeRow, activeCol, value, options);
    return;
  }
  const content = coerceCellInput(sheet, activeRow, activeCol, value, options);
  const invalid = options.skipValidation ? null : targets.find((target) => !validateCellInput(sheet, target.row, target.col, content).valid);
  if (invalid) {
    showValidationErrorForCell(invalid.row, invalid.col, content, validateCellInput(sheet, invalid.row, invalid.col, content));
    return;
  }
  const protectedRange = unionRanges(ranges) || normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: activeRow, col: activeCol },
    { sheetIndex: state.activeSheetIndex, row: activeRow, col: activeCol },
  );
  if (!ensureEditableRangeForProtection(sheet, protectedRange, "変更")) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  const sheetId = getSheetId(sheet.name);
  try {
    targets.forEach((target) => writeCellContent(sheet, sheetId, target.row, target.col, content, options));
    refreshWorkbookChartsAfterSourceEdit(state.activeSheetIndex, protectedRange);
    state.selected = { sheetIndex: state.activeSheetIndex, row: activeRow, col: activeCol };
    state.selectionRange = selectionBefore;
    state.selectionRanges = selectionBeforeRanges.map((range) => ({ ...range }));
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetBefore,
      sheetAfter: captureSheetForHistory(sheet),
      selectionBefore,
      selectionAfter: activeSelectionRange(),
      selectionBeforeRanges,
      selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
      label: `${rangeToLabel(protectedRange)} 一括入力`,
    });
    markWorkbookChanged();
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForCell(activeRow, activeCol);
    setStatus(`${targets.length} 個のセルに入力しました。`);
  } catch (error) {
    console.error(error);
    restoreSheetFromHistory(sheet, sheetBefore);
    state.hf = buildFormulaEngine(state.model);
    renderSheet();
    showMessageBox("選択範囲に入力できませんでした。", { title: "Cht WebSheet" });
  }
}



function uniqueEditableTargetsForRanges(sheet, ranges, options = {}) {
  const targets = [];
  const seen = new Set();
  const sheetIndex = Number.isFinite(Number(options.sheetIndex)) ? Number(options.sheetIndex) : formulaEngineSheetIndex(sheet);
  ranges.forEach((range) => {
    const materialized = materializeSelectionRangeForSheet(sheet, range, sheetIndex);
    if (!materialized) return;
    ensureSheetSize(sheet, materialized.bottom, materialized.right);
    forEachCellInRange(materialized, (row, col) => {
      const merge = findMergeForCell(sheet, row, col);
      const targetRow = merge ? merge.top : row;
      const targetCol = merge ? merge.left : col;
      const key = cellKey(targetRow, targetCol);
      if (seen.has(key)) return;
      seen.add(key);
      targets.push({ row: targetRow, col: targetCol });
    });
  });
  targets.sort((left, right) => left.row - right.row || left.col - right.col);
  return targets;
}



function unionRanges(ranges) {
  const materialized = (ranges || []).map((range) => materializeSelectionRange(range)).filter(Boolean);
  return unionMaterializedRanges(materialized, state.activeSheetIndex);
}



function fallbackSingleCellRangeForSheet(sheetIndex, row, col) {
  return normalizeRange(
    { sheetIndex, row, col },
    { sheetIndex, row, col },
  );
}



function activeSelectionRangesForGroupedOperation(activeRow, activeCol) {
  const ranges = activeSelectionRanges().map((range) => ({ ...range }));
  if (ranges.length) return ranges;
  return [fallbackSingleCellRangeForSheet(state.activeSheetIndex, activeRow, activeCol)];
}



function groupedRangesForSheet(sheet, sheetIndex, sourceRanges, activeRow, activeCol) {
  const ranges = materializedRangesForSheet(sheet, sourceRanges, sheetIndex);
  if (ranges.length) return ranges;
  ensureSheetSize(sheet, activeRow, activeCol);
  return [fallbackSingleCellRangeForSheet(sheetIndex, activeRow, activeCol)];
}



function affectedSheetIndexesForGroupedRanges(items) {
  const affected = new Set();
  (items || []).forEach((item) => {
    if (!state.model?.sheets?.[item.sheetIndex]) return;
    affected.add(item.sheetIndex);
    (item.ranges || item.targetRanges || [item.range]).filter(Boolean).forEach((range) => {
      chartSourceOwnerSheetIndexesForRange(item.sheetIndex, range).forEach((index) => affected.add(index));
    });
  });
  return [...affected].sort((left, right) => left - right);
}



function refreshChartsForGroupedRanges(items) {
  let changed = false;
  (items || []).forEach((item) => {
    (item.ranges || item.targetRanges || [item.range]).filter(Boolean).forEach((range) => {
      if (refreshWorkbookChartsAfterSourceEdit(item.sheetIndex, range)) changed = true;
    });
  });
  return changed;
}



function restoreGroupedSheetsAfterFailure(sheetsBefore) {
  if (sheetsBefore?.length) restoreSheetsFromHistory(sheetsBefore);
  state.hf = buildFormulaEngine(state.model);
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
}



function commitCellValueToSheetGroup(activeRow, activeCol, value, options = {}) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const content = coerceCellInput(activeSheet(), activeRow, activeCol, value, options);
  const selectionBefore = activeSelectionRange() || fallbackSingleCellRangeForSheet(state.activeSheetIndex, activeRow, activeCol);
  const selectionBeforeRanges = activeSelectionRangesForGroupedOperation(activeRow, activeCol);
  const items = [];

  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model.sheets[sheetIndex];
    const targetRanges = groupedRangesForSheet(sheet, sheetIndex, selectionBeforeRanges, activeRow, activeCol);
    const targets = uniqueEditableTargetsForRanges(sheet, targetRanges, { sheetIndex });
    if (!targets.length) continue;
    const invalid = options.skipValidation ? null : targets.find((target) => !validateCellInput(sheet, target.row, target.col, content).valid);
    if (invalid) {
      const validation = validateCellInput(sheet, invalid.row, invalid.col, content);
      if (sheetIndex === state.activeSheetIndex) {
        showValidationErrorForCell(invalid.row, invalid.col, content, validation);
      } else {
        showMessageBox(`${sheet.name}!${columnName(invalid.col)}${invalid.row}\n${validation.message || "入力した値は、このセルに設定されている入力規則に違反しています。"}`, {
          title: validation.title || "Cht WebSheet",
          tone: validation.style === "information" ? "info" : "warning",
        });
      }
      return true;
    }
    if (!ensureEditableRangesForProtection(sheet, targetRanges, "変更")) return true;
    items.push({
      sheetIndex,
      sheet,
      sheetId: getSheetId(sheet.name),
      targetRanges,
      targets,
      ranges: targetRanges,
    });
  }

  if (!items.length) return true;
  const affectedSheetIndexes = affectedSheetIndexesForGroupedRanges(items);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  try {
    items.forEach((item) => {
      item.targets.forEach((target) => writeCellContent(item.sheet, item.sheetId, target.row, target.col, content, options));
    });
    refreshChartsForGroupedRanges(items);
    state.selected = { sheetIndex: state.activeSheetIndex, row: activeRow, col: activeCol };
    state.selectionRange = selectionBefore;
    state.selectionRanges = selectionBeforeRanges.map((range) => ({ ...range, sheetIndex: state.activeSheetIndex }));
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore,
      selectionAfter: activeSelectionRange(),
      selectionBeforeRanges,
      selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
      label: `${rangesToLabel(selectionBeforeRanges)} ${sheetIndexes.length} シート一括入力`,
    });
    markWorkbookChanged();
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForCell(activeRow, activeCol);
    setStatus(`${sheetIndexes.length} 個のシートの${rangesToLabel(selectionBeforeRanges)}に入力しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    showMessageBox("選択したシート グループに入力できませんでした。", { title: "Cht WebSheet" });
  }
  return true;
}



function commitCellValue(row, col, value, options = {}) {
  const sheet = activeSheet();
  if (!options.skipSheetGroup && commitCellValueToSheetGroup(row, col, value, options)) return;
  if (!ensureEditableRangeForProtection(sheet, selectionRangeFromPoints({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col }), "変更")) return;
  const sheetId = getSheetId(sheet.name);
  const content = options.agentInput ? options.agentInput.content : coerceCellInput(sheet, row, col, value, options);
  const validation = options.skipValidation ? { valid: true } : validateCellInput(sheet, row, col, content);
  if (!validation.valid) {
    showValidationErrorForCell(row, col, content, validation);
    return;
  }
  const range = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row, col },
    { sheetIndex: state.activeSheetIndex, row, col },
  );
  const tableEditCandidate = tableEditMayChangeMetadata(sheet, row, col, content);
  const chartSourceOwnerIndexes = tableEditCandidate ? [] : chartSourceOwnerSheetIndexesForRange(state.activeSheetIndex, range);
  const chartHistoryIndexes = [...new Set([state.activeSheetIndex, ...chartSourceOwnerIndexes])];
  const sheetBefore = tableEditCandidate || (chartSourceOwnerIndexes.length && chartHistoryIndexes.length === 1) ? captureSheetForHistory(sheet) : null;
  const sheetsBefore = tableEditCandidate
    ? captureSheetsForHistory(state.model.sheets.map((_item, index) => index))
    : chartHistoryIndexes.length > 1
      ? captureSheetsForHistory(chartHistoryIndexes)
      : null;
  const before = captureCellsForHistory(sheet, range);
  const rowDimensionBefore = captureDimensionForHistory(sheet, "row", row);

  try {
    if (options.agentInput?.numFmt) {
      const cell = ensureCellModel(sheet, row, col);
      cell.numFmt = options.agentInput.numFmt;
      storeCellModelForSheet(sheet, cellKey(row, col), cell);
    }
    writeCellContent(sheet, sheetId, row, col, content, options);
    const chartsChanged = refreshWorkbookChartsAfterSourceEdit(state.activeSheetIndex, range);
    const tableEdit = applyTableEditBehaviors(sheet, row, col, content);
    if (tableEdit.changed) {
      state.hf = buildFormulaEngine(state.model);
      const selectionAfter = normalizeRange(
        { sheetIndex: state.activeSheetIndex, row, col },
        { sheetIndex: state.activeSheetIndex, row, col },
      );
      commitDataOperation(sheet, sheetBefore || captureSheetForHistory(sheet), range, tableEdit.label, tableEdit.status, selectionAfter, { sheetsBefore });
      return;
    }
    state.selected = { sheetIndex: state.activeSheetIndex, row, col };
    state.selectionRange = range;
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectionDrag = null;
    const after = captureCellsForHistory(sheet, range);
    const rowDimensionAfter = captureDimensionForHistory(sheet, "row", row);
    const rowHeightChanged =
      rowDimensionBefore &&
      rowDimensionAfter &&
      Math.abs(Number(rowDimensionBefore.size) - Number(rowDimensionAfter.size)) > 0.01;
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      ...(chartsChanged && sheetsBefore?.length
        ? {
            sheetsBefore,
            sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
          }
        : chartsChanged && sheetBefore
          ? {
              sheetBefore,
              sheetAfter: captureSheetForHistory(sheet),
            }
          : {
              before,
              after,
              dimensionsBefore: rowHeightChanged ? [rowDimensionBefore] : null,
              dimensionsAfter: rowHeightChanged ? [rowDimensionAfter] : null,
            }),
      selectionBefore: range,
      selectionAfter: range,
      label: `${columnName(col)}${row} 更新`,
    });
    markWorkbookChanged();
    renderSheet();
    $selectedCellLabel.text(`${columnName(col)}${row}`);
    $formulaInput.val(getCellEditInput(sheet, row, col));
    setStatus(`${sheet.name}!${columnName(col)}${row} を更新しました。`);
  } catch (error) {
    console.error(error);
    setStatus(`${columnName(col)}${row} に数式または値を書き込めませんでした：${error.message || error}`);
    renderSheet();
  }
}



function showValidationErrorForCell(row, col, content, validation) {
  const style = validation.style || "stop";
  if (style === "warning") {
    showMessageBox(validation.message, {
      title: validation.title || "Cht WebSheet",
      tone: "warning",
      buttons: [
        { label: "はい(Y)", result: "yes", callback: () => commitCellValue(row, col, content, { skipValidation: true }) },
        { label: "いいえ(N)", result: "no", callback: () => rejectInvalidCellInput(row, col) },
        { label: "キャンセル", result: "cancel", callback: () => rejectInvalidCellInput(row, col) },
      ],
      cancelResult: "cancel",
    });
    return;
  }
  if (style === "information") {
    showMessageBox(validation.message, {
      title: validation.title || "Cht WebSheet",
      tone: "info",
      buttons: [
        { label: "OK", result: "ok", callback: () => commitCellValue(row, col, content, { skipValidation: true }) },
        { label: "キャンセル", result: "cancel", callback: () => rejectInvalidCellInput(row, col) },
      ],
      cancelResult: "cancel",
    });
    return;
  }
  showMessageBox(validation.message, {
    title: validation.title || "Cht WebSheet",
    tone: "stop",
    buttons: [
      { label: "再試行(R)", result: "retry", callback: () => retryInvalidCellInput(row, col, content) },
      { label: "キャンセル", result: "cancel", callback: () => rejectInvalidCellInput(row, col) },
    ],
    cancelResult: "cancel",
  });
}



function retryInvalidCellInput(row, col, content) {
  const range = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row, col },
    { sheetIndex: state.activeSheetIndex, row, col },
  );
  state.selected = { sheetIndex: state.activeSheetIndex, row, col };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForCell(row, col);
  $formulaInput.val(String(content ?? ""));
  beginFormulaBarEdit();
  syncFormulaEditPreview();
  $formulaInput.trigger("focus").select();
  setStatus("入力値を修正してください。");
}



function rejectInvalidCellInput(row, col) {
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForCell(row, col);
}



function writeCellContent(sheet, sheetId, row, col, content, options = {}) {
  ensureSheetSize(sheet, row, col);
  let formulaEngineError = null;
  try {
    state.hf.setCellContents({ sheet: sheetId, row: row - 1, col: col - 1 }, [[formulaEngineContentForCurrentModel(content, formulaEngineSheetIndex(sheet), { row, col })]]);
  } catch (error) {
    formulaEngineError = error;
    console.warn("Formula engine rejected cell content; keeping workbook text model in sync.", error);
  }
  updateCellModel(sheet, row, col, content, options);
  if (formulaEngineError) {
    const cell = sheet.cells[cellKey(row, col)];
    if (cell) cell.formulaEngineError = formulaEngineError.message || String(formulaEngineError);
  }
}



function updateCellModel(sheet, row, col, content, options = {}) {
  const key = cellKey(row, col);
  const cell = sheet.cells[key] || ensureCellModel(sheet, row, col);
  const previousText = cell.raw == null ? cellDisplayText(cell) : cell.raw;
  const isFormula = typeof content === "string" && content.startsWith("=");
  cell.raw = content;
  cell.formula = isFormula ? content : null;
  cell.cached = isFormula ? null : content;
  cell.display = initialCellDisplayText(cell, content, isFormula);
  delete cell.dataTableFormula;
  const hasRichTextOption = Object.prototype.hasOwnProperty.call(options, "richText");
  const nextRichText = hasRichTextOption && !isFormula ? normalizeCellEditRichTextForModel(options.richText, content) : null;
  const textChanged = String(previousText ?? "") !== String(content ?? "");
  if (hasRichTextOption) {
    if (nextRichText) cell.richText = nextRichText;
    else delete cell.richText;
  } else if (textChanged) {
    delete cell.richText;
  }
  if (textChanged) {
    delete cell.phonetics;
    delete cell.showPhonetic;
    delete cell.sparkline;
  }
  applyMultilineCellDisplayBehavior(sheet, row, col, cell, content);
  cell.borders = cell.borders || {};
  storeCellModelForSheet(sheet, key, cell);
}



function initialCellDisplayText(cell, content, isFormula) {
  if (isFormula || content == null) return "";
  if (cell?.numFmt && cellCanRecomputeDisplayText({ raw: content, cached: content, formula: null })) return "";
  return String(content);
}



function applyMultilineCellDisplayBehavior(sheet, row, col, cell, content) {
  if (typeof content !== "string" || content.startsWith("=") || !hasCellLineBreak(content) || !cell) return false;
  const css = { ...(cell.css || {}) };
  css.whiteSpace = "pre-wrap";
  css.overflowWrap = "anywhere";
  cell.css = css;

  const rowModel = sheet.rows[row - 1];
  if (!rowModel || rowModel.hidden) return true;
  const currentHeight = Number(rowModel.height || sheet.defaultRowHeight || rowHeightToPixels(15));
  const requiredHeight = multilineCellAutoHeight(sheet, row, col, cell, content);
  if (requiredHeight > currentHeight + 0.5) {
    rowModel.height = requiredHeight;
  }
  return true;
}



function hasCellLineBreak(value) {
  return /\r\n|\r|\n/.test(String(value ?? ""));
}



function multilineCellAutoHeight(sheet, row, col, cell, text) {
  const css = cellEffectiveCss(sheet, cell);
  const fontSize = canvasFontSize(css.fontSize);
  const lineHeight = canvasTextLineHeight(fontSize);
  const range = mergedRangeForCell(sheet, row, col) || { top: row, left: col, bottom: row, right: col };
  const rect = rangePixelRect(sheet, range);
  const maxTextWidth = Math.max(1, rect.width - 6);
  let lineCount = String(text).split(/\r\n|\r|\n/).length;
  const context = multilineMeasureContext();
  if (context) {
    lineCount = Math.max(lineCount, wrapCanvasText(context, text, maxTextWidth, canvasFont(css, fontSize)).length);
  }
  return Math.ceil(lineCount * lineHeight + 4);
}



function multilineMeasureContext() {
  if (typeof document === "undefined") return null;
  const canvas = document.createElement("canvas");
  return canvas.getContext?.("2d") || null;
}



function mergedRangeForCell(sheet, row, col) {
  return (sheet.merges || []).find((merge) => row >= merge.top && row <= merge.bottom && col >= merge.left && col <= merge.right) || null;
}



function captureDimensionForHistory(sheet, kind, index) {
  if (!sheet || !index) return null;
  const isColumn = kind === "col";
  const item = isColumn ? sheet.columns[index - 1] : sheet.rows[index - 1];
  const fallback = isColumn ? sheet.defaultColumnWidth || columnWidthToPixels(8.43) : sheet.defaultRowHeight || rowHeightToPixels(15);
  return { kind, index, size: Number(item?.[isColumn ? "width" : "height"] || fallback) };
}



function coerceUserInput(input, options = {}) {
  const source = String(input ?? "");
  const preserveWhitespace = options.preserveWhitespace === true;
  const text = preserveWhitespace ? source : source.trim();
  if (!text) return null;
  if (text.startsWith("'")) return text.slice(1);
  if (text.startsWith("=")) return formulaTextWithExcelAutoCompletion(text);
  if (preserveWhitespace && text !== text.trim()) return text;
  if (/^(true|false)$/i.test(text)) return /^true$/i.test(text);
  if (/^-?\d+(\.\d+)?$/.test(text)) return Number(text);
  const formattedNumber = parseFormattedNumericInput(text);
  if (formattedNumber) return formattedNumber.number;
  return text;
}



function coerceCellInput(sheet, row, col, input, options = {}) {
  const cell = sheet?.cells?.[cellKey(row, col)];
  const dateSerial = editedDateSerialForCell(cell, input);
  return dateSerial == null ? coerceUserInput(input, options) : dateSerial;
}



function editedDateSerialForCell(cell, input) {
  if (!cell || cell.formula || cell.dataTableFormula || !isDateFormat(cell.numFmt) || isElapsedTimeFormat(cell.numFmt) || input == null || typeof input === "object") {
    return null;
  }
  const text = String(input).trim();
  if (!text || text.startsWith("'") || text.startsWith("=") || /[\r\n]/.test(text)) return null;

  const timeSerial = editedTimeSerialForFormat(text, cell.numFmt);
  if (timeSerial != null) return timeSerial;

  const completeDate = parsePersistedDateString(text);
  if (completeDate) return dateToExcelSerial(completeDate);

  const baselineSerial = dateSerialFromPersistedValue(cell.raw, cell.numFmt)
    ?? dateSerialFromPersistedValue(cell.cached, cell.numFmt);
  const baseline = baselineSerial == null ? null : excelSerialToDate(baselineSerial);
  const now = new Date();
  const parts = {
    year: baseline?.getFullYear() ?? now.getFullYear(),
    month: (baseline?.getMonth() ?? now.getMonth()) + 1,
    day: baseline?.getDate() ?? 1,
    hour: baseline?.getHours() ?? 0,
    minute: baseline?.getMinutes() ?? 0,
    second: baseline?.getSeconds() ?? 0,
  };
  const build = (next, dayIsExplicit = false) => {
    const values = { ...parts, ...next };
    if (!dayIsExplicit) {
      values.day = Math.min(values.day, new Date(values.year, values.month, 0).getDate());
    }
    return localDateFromPersistedParts(values.year, values.month, values.day, values.hour, values.minute, values.second);
  };

  let match = text.match(/^(\d{4})年\s*(\d{1,2})月$/);
  let date = match ? build({ year: Number(match[1]), month: Number(match[2]) }) : null;
  if (!date) {
    match = text.match(/^(\d{4})[-/.](\d{1,2})$/);
    date = match ? build({ year: Number(match[1]), month: Number(match[2]) }) : null;
  }
  if (!date) {
    match = text.match(/^(\d{1,2})月\s*(\d{1,2})日$/);
    date = match ? build({ month: Number(match[1]), day: Number(match[2]) }, true) : null;
  }
  if (!date) {
    match = text.match(/^(\d{1,2})[-/.](\d{1,2})$/);
    date = match ? build({ month: Number(match[1]), day: Number(match[2]) }, true) : null;
  }
  if (!date) {
    match = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    date = match ? build({ hour: Number(match[1]), minute: Number(match[2]), second: Number(match[3] || 0) }, true) : null;
  }

  const format = dateDetectionFormat(cell.numFmt).trim().toLowerCase();
  if (!date && /^y{1,4}$/.test(format) && /^\d{4}$/.test(text)) {
    date = build({ year: Number(text) });
  }
  if (!date && /^m{1,4}$/.test(format) && /^\d{1,2}$/.test(text)) {
    date = build({ month: Number(text) });
  }
  if (!date && /^d{1,4}$/.test(format) && /^\d{1,2}$/.test(text)) {
    date = build({ day: Number(text) }, true);
  }
  return date ? dateToExcelSerial(date) : null;
}



function editedTimeSerialForFormat(text, numFmt = "") {
  if (!isTimeOnlyDateFormat(numFmt)) return null;
  const match = String(text || "").trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?$/i);
  if (!match) return null;
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const second = Number(match[3] || 0);
  const meridiem = String(match[4] || "").toUpperCase();
  if (meridiem) {
    if (hour < 1 || hour > 12) return null;
    hour = hour % 12 + (meridiem === "PM" ? 12 : 0);
  }
  if (hour < 0 || hour > 23 || minute < 0 || minute > 59 || second < 0 || second > 59) return null;
  return (hour * 3600 + minute * 60 + second) / 86400;
}



function isTimeOnlyDateFormat(numFmt = "") {
  if (isElapsedTimeFormat(numFmt)) return false;
  const format = dateDetectionFormat(numFmt);
  const hasTime = /(^|[^a-z])h{1,2}|(^|[^a-z])s{1,2}|AM\/PM|A\/P/i.test(format);
  if (!hasTime) return false;
  const hasCalendarDate = /[年月日]|(^|[^a-z])y{1,4}|(^|[^a-z])d{1,4}/i.test(format);
  return !hasCalendarDate;
}



function parseFormattedNumericInput(input) {
  if (input === null || input === undefined || typeof input === "object") return null;
  let text = String(input).trim();
  if (!text || text.startsWith("'") || text.startsWith("=") || /[\r\n]/.test(text)) return null;
  text = text.replace(/\u00a0/g, " ");
  const currencyMatch = text.match(/[¥￥$€£]/);
  const currency = currencyMatch ? (currencyMatch[0] === "￥" ? "¥" : currencyMatch[0]) : "";
  const percent = /%$/.test(text);
  let negative = false;
  if (/^\s*[-−]/.test(text)) {
    negative = true;
    text = text.replace(/^\s*[-−]\s*/, "");
  }
  if (/^\(\s*.+\s*\)$/.test(text)) {
    negative = true;
    text = text.replace(/^\(\s*|\s*\)$/g, "");
  }
  if (currency) {
    text = text.replace(/[¥￥$€£]/g, "");
  }
  text = text.replace(/円|税込|税抜/g, "");
  if (currency && /-$/.test(text.trim())) {
    text = text.trim().slice(0, -1);
  }
  if (percent) {
    text = text.replace(/%$/, "");
  }
  text = text.replace(/\s+/g, "");
  const usedGrouping = text.includes(",");
  if (usedGrouping && !/^[+-]?(?:\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(text)) return null;
  const normalized = text.replace(/,/g, "");
  if (!/^[+]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(normalized)) return null;
  let number = Number(normalized);
  if (!Number.isFinite(number)) return null;
  if (negative) number *= -1;
  if (percent) number /= 100;
  if (!currency && !percent && !usedGrouping) return null;
  return { number, currency, percent, usedGrouping };
}



function getCellRawInput(sheet, row, col) {
  if (!sheet) return "";
  const cell = sheet.cells[cellKey(row, col)];
  if (!cell) return "";
  if (cell.dataTableFormula) return cell.dataTableFormula;
  if (cell.formula) return cell.formula;
  return cell.raw == null ? "" : String(cell.raw);
}



function getCellEditInput(sheet, row, col) {
  const rawInput = getCellRawInput(sheet, row, col);
  const cell = sheet?.cells?.[cellKey(row, col)];
  if (!cell || cell.dataTableFormula || cell.formula || !isDateFormat(cell.numFmt)) return rawInput;
  const serial = dateSerialFromPersistedValue(cell.raw, cell.numFmt)
    ?? dateSerialFromPersistedValue(cell.cached, cell.numFmt);
  return serial == null ? rawInput : displayValue(serial, cell.numFmt);
}



function getDisplayForCell(sheet, sheetId, row, col) {
  const cache = state.renderDisplayCache;
  const cacheKey = cache ? renderDisplayCacheKey(sheet, sheetId, row, col) : "";
  if (cache && cache.has(cacheKey)) return cache.get(cacheKey);
  const display = computeDisplayForCell(sheet, sheetId, row, col);
  if (cache) cache.set(cacheKey, display);
  return display;
}



function renderDisplayCacheKey(sheet, sheetId, row, col) {
  const sheetIndex = state.model?.sheets?.indexOf(sheet) ?? -1;
  const cell = sheet?.cells?.[cellKey(row, col)] || null;
  return [
    sheetIndex,
    sheetId ?? "",
    row,
    col,
    state.showFormulas ? 1 : 0,
    cell?.numFmt ?? "",
    cell?.raw ?? "",
    cell?.cached ?? "",
    cell?.display ?? "",
    cell?.formula ?? "",
  ].join(":");
}



function computeDisplayForCell(sheet, sheetId, row, col) {
  const cell = sheet.cells[cellKey(row, col)];
  if (!cell?.formula) {
    if (cell && (cell.raw != null || cell.cached != null || cell.display != null)) {
      return shouldHideZeroValueForSheet(sheet, cell) ? "" : cellDisplayText(cell);
    }
    const spill = excelCompatibleFormulaSpillValueAt(sheet, row, col);
    if (spill.supported) return shouldHideZeroValueForSheet(sheet, cell, spill.value) ? "" : displayValue(spill.value, spill.numFmt);
    if (state.hf && sheetId != null) {
      try {
        const value = state.hf.getCellValue({ sheet: sheetId, row: row - 1, col: col - 1 });
        if (value != null) return shouldHideZeroValueForSheet(sheet, cell, value) ? "" : displayValue(value, cell?.numFmt);
      } catch {}
    }
    return shouldHideZeroValueForSheet(sheet, cell) ? "" : cellDisplayText(cell);
  }
  const formulaSpill = evaluateExcelArraySpillFormula(sheet, state.model?.sheets?.indexOf(sheet) ?? state.activeSheetIndex, cell.formula);
  if (formulaSpill.supported) {
    const value = formulaSpill.data[0]?.[0] ?? "";
    return shouldHideZeroValueForSheet(sheet, cell, value) ? "" : displayValue(value, cell?.numFmt);
  }
  const compatibleValue = excelCompatibleFormulaValue(sheet, cell);
  if (compatibleValue.supported) return shouldHideZeroValueForSheet(sheet, cell, compatibleValue.value) ? "" : displayValue(compatibleValue.value, cell?.numFmt);
  if (!state.hf || sheetId == null) {
    return shouldHideZeroValueForSheet(sheet, cell) ? "" : cellDisplayText(cell);
  }

  try {
    const value = state.hf.getCellValue({ sheet: sheetId, row: row - 1, col: col - 1 });
    if (shouldHideZeroValueForSheet(sheet, cell, value)) return "";
    return displayValue(value, cell?.numFmt);
  } catch {
    return shouldHideZeroValueForSheet(sheet, cell) ? "" : cellDisplayText(cell);
  }
}



function createFormulaSpillRenderCache(sheet) {
  return {
    sheet,
    anchors: null,
    anchorSpills: new Map(),
    targetSpills: new Map(),
  };
}



function withDisplayCalculationCache(sheet, callback) {
  const previousRenderDisplayCache = state.renderDisplayCache;
  const previousRenderFormulaSpillCache = state.renderFormulaSpillCache;
  state.renderDisplayCache = previousRenderDisplayCache instanceof Map ? previousRenderDisplayCache : new Map();
  state.renderFormulaSpillCache = formulaSpillRenderCacheForSheet(sheet) || createFormulaSpillRenderCache(sheet);
  try {
    return callback();
  } finally {
    state.renderDisplayCache = previousRenderDisplayCache;
    state.renderFormulaSpillCache = previousRenderFormulaSpillCache;
  }
}



function formulaSpillRenderCacheForSheet(sheet) {
  return state.renderFormulaSpillCache?.sheet === sheet ? state.renderFormulaSpillCache : null;
}



function invalidateSheetUsedRangeCache(sheet) {
  if (!sheet || typeof sheet !== "object") return;
  sheetUsedRangeCacheBySheet.delete(sheet);
  sheetMinimumExtentCacheBySheet.delete(sheet);
  sheetBorderCellCacheBySheet.delete(sheet);
  sheetPersistentNoteCellCacheBySheet.delete(sheet);
}



function invalidateFormulaSpillAnchorCache(sheet) {
  if (!sheet || typeof sheet !== "object") return;
  formulaSpillAnchorCacheBySheet.delete(sheet);
  if (state.renderFormulaSpillCache?.sheet === sheet) {
    state.renderFormulaSpillCache.anchors = null;
    state.renderFormulaSpillCache.anchorSpills.clear();
    state.renderFormulaSpillCache.targetSpills.clear();
  }
}



function primeWorkbookFormulaSpillAnchorCache(model) {
  (model?.sheets || []).forEach((sheet) => {
    if (sheet) formulaSpillAnchorEntries(sheet);
  });
}



function primeWorkbookUsedRangeCache(model) {
  (model?.sheets || []).forEach((sheet) => {
    if (sheet) sheetUsedRange(sheet);
  });
}



function formulaSpillAnchorEntries(sheet) {
  const cache = formulaSpillRenderCacheForSheet(sheet);
  if (cache?.anchors) return cache.anchors;
  const persistent = formulaSpillAnchorCacheBySheet.get(sheet);
  if (persistent?.cells === sheet?.cells) {
    if (cache) cache.anchors = persistent.anchors;
    return persistent.anchors;
  }
  const entries = Object.entries(sheet?.cells || {})
    .map(([key, cell]) => {
      if (!cell?.formula) return null;
      const parsed = parseCellKey(key);
      const rawRow = Number(cell.row) || parsed[0];
      const rawCol = Number(cell.col) || parsed[1];
      if (!rawRow || !rawCol) return null;
      const row = Math.max(1, rawRow);
      const col = Math.max(1, rawCol);
      return { key, row, col, cell };
    })
    .filter(Boolean)
    .sort((left, right) => left.row - right.row || left.col - right.col);
  formulaSpillAnchorCacheBySheet.set(sheet, { cells: sheet?.cells, anchors: entries });
  if (cache) cache.anchors = entries;
  return entries;
}



function evaluateExcelArraySpillFormulaCached(sheet, sheetIndex, anchor) {
  const cache = formulaSpillRenderCacheForSheet(sheet);
  if (!cache) return evaluateExcelArraySpillFormula(sheet, sheetIndex, anchor.cell.formula);
  if (cache.anchorSpills.has(anchor.key)) return cache.anchorSpills.get(anchor.key);
  const spill = evaluateExcelArraySpillFormula(sheet, sheetIndex, anchor.cell.formula);
  cache.anchorSpills.set(anchor.key, spill);
  return spill;
}



function excelCompatibleFormulaSpillValueAt(sheet, row, col) {
  const cache = formulaSpillRenderCacheForSheet(sheet);
  const targetKey = cache ? cellKey(row, col) : "";
  if (cache?.targetSpills.has(targetKey)) return cache.targetSpills.get(targetKey);
  const sheetIndex = state.model?.sheets?.indexOf(sheet) ?? state.activeSheetIndex;
  const maxLookbackRows = 200;
  const maxLookbackCols = 100;
  const minAnchorRow = Math.max(1, row - maxLookbackRows);
  const minAnchorCol = Math.max(1, col - maxLookbackCols);
  for (const anchor of formulaSpillAnchorEntries(sheet)) {
    if (anchor.row < minAnchorRow || anchor.row > row || anchor.col < minAnchorCol || anchor.col > col) continue;
    const spill = evaluateExcelArraySpillFormulaCached(sheet, sheetIndex, anchor);
    if (!spill.supported) continue;
    const rowOffset = row - anchor.row;
    const colOffset = col - anchor.col;
    if (rowOffset < 0 || colOffset < 0 || rowOffset >= spill.data.length || colOffset >= (spill.data[rowOffset] || []).length) continue;
    const result = { supported: true, value: spill.data[rowOffset][colOffset], numFmt: anchor.cell.numFmt };
    if (cache) cache.targetSpills.set(targetKey, result);
    return result;
  }
  const result = { supported: false, value: null, numFmt: "" };
  if (cache) cache.targetSpills.set(targetKey, result);
  return result;
}



function evaluateExcelArraySpillFormula(sheet, sheetIndex, formula) {
  const expression = String(formula || "").trim().replace(/^=/, "").trim();
  const directRange = decodeFormulaRangeForSheet(expression, sheet, sheetIndex);
  if (directRange) {
    const data = formulaRangeValueMatrix(directRange);
    return { supported: true, data: data.length ? data : [[""]] };
  }
  const match = expression.match(/^(UNIQUE|SORT|SORTBY|FILTER|SEQUENCE|INDIRECT|OFFSET|INDEX|TAKE|DROP|VSTACK|HSTACK|CHOOSEROWS|CHOOSECOLS|TOCOL|TOROW|WRAPROWS|WRAPCOLS|EXPAND|TRANSPOSE|TEXTSPLIT|MAKEARRAY|BYROW|BYCOL|MAP|REDUCE|SCAN)\s*\((.*)\)\s*$/i);
  if (!match) return { supported: false, data: [] };
  const name = match[1].toUpperCase();
  const args = splitFormulaArguments(match[2]);
  if (name === "SEQUENCE") {
    return {
      supported: true,
      data: sequenceFormulaMatrix(
        Math.max(1, Math.trunc(formulaArgumentNumberValue(args[0], sheet, sheetIndex, 1))),
        Math.max(1, Math.trunc(formulaArgumentNumberValue(args[1], sheet, sheetIndex, 1))),
        formulaArgumentNumberValue(args[2], sheet, sheetIndex, 1),
        formulaArgumentNumberValue(args[3], sheet, sheetIndex, 1),
      ),
    };
  }
  if (name === "TEXTSPLIT") {
    return {
      supported: true,
      data: textSplitFormulaMatrix(
        formulaArgumentScalarValue(args[0], sheet, sheetIndex),
        formulaArgumentMatrix(args[1], sheet, sheetIndex),
        args[2] == null || String(args[2]).trim() === "" ? undefined : formulaArgumentMatrix(args[2], sheet, sheetIndex),
        formulaArgumentBoolean(args[3], false),
        formulaArgumentNumberValue(args[4], sheet, sheetIndex, 0),
        args[5] == null ? webSheetFormulaError(ErrorType.NA) : formulaArgumentScalarValue(args[5], sheet, sheetIndex),
      ),
    };
  }
  if (name === "MAKEARRAY") {
    return { supported: true, data: makeArrayLambdaFormulaMatrix(args, sheet, sheetIndex) };
  }
  if (name === "BYROW") {
    return { supported: true, data: byRowLambdaFormulaMatrix(args, sheet, sheetIndex) };
  }
  if (name === "BYCOL") {
    return { supported: true, data: byColLambdaFormulaMatrix(args, sheet, sheetIndex) };
  }
  if (name === "MAP") {
    return { supported: true, data: mapLambdaFormulaMatrix(args, sheet, sheetIndex) };
  }
  if (name === "REDUCE") {
    return { supported: true, data: [[reduceLambdaFormulaValue(args, sheet, sheetIndex)]] };
  }
  if (name === "SCAN") {
    return { supported: true, data: scanLambdaFormulaMatrix(args, sheet, sheetIndex) };
  }
  if (name === "INDIRECT" || name === "OFFSET" || name === "INDEX") {
    const range = decodeFormulaRangeForSheet(`${name}(${match[2]})`, sheet, sheetIndex);
    if (!range) return { supported: true, data: [["#REF!"]] };
    const data = formulaRangeValueMatrix(range);
    return { supported: true, data: data.length ? data : [[""]] };
  }
  if (name === "VSTACK") {
    return { supported: true, data: vstackFormulaMatrices(args.map((arg) => formulaArgumentMatrix(arg, sheet, sheetIndex))) };
  }
  if (name === "HSTACK") {
    return { supported: true, data: hstackFormulaMatrices(args.map((arg) => formulaArgumentMatrix(arg, sheet, sheetIndex))) };
  }
  const sourceData = formulaArgumentMatrix(args[0], sheet, sheetIndex);
  if (!sourceData.length) return { supported: true, data: [[""]] };
  if (name === "TRANSPOSE") {
    return { supported: true, data: transposeFormulaMatrix(sourceData) };
  }
  if (name === "UNIQUE") {
    return {
      supported: true,
      data: uniqueFormulaMatrix(
        sourceData,
        formulaArgumentBoolean(args[1], false),
        formulaArgumentBoolean(args[2], false),
      ),
    };
  }
  if (name === "FILTER") {
    return {
      supported: true,
      data: filterFormulaMatrix(
        sourceData,
        formulaFilterIncludeMatrix(args[1], sheet, sheetIndex),
        formulaArgumentScalarValue(args[2], sheet, sheetIndex),
      ),
    };
  }
  if (name === "SORTBY") {
    const descriptors = [];
    for (let index = 1; index < args.length; index += 2) {
      descriptors.push({
        data: formulaArgumentMatrix(args[index], sheet, sheetIndex),
        order: formulaArgumentNumberValue(args[index + 1], sheet, sheetIndex, 1),
      });
    }
    return { supported: true, data: sortByFormulaMatrix(sourceData, descriptors) };
  }
  if (name === "TAKE") {
    return {
      supported: true,
      data: takeFormulaMatrix(
        sourceData,
        formulaArgumentNumberValue(args[1], sheet, sheetIndex, 0),
        args[2] == null ? undefined : formulaArgumentNumberValue(args[2], sheet, sheetIndex, 0),
      ),
    };
  }
  if (name === "DROP") {
    return {
      supported: true,
      data: dropFormulaMatrix(
        sourceData,
        formulaArgumentNumberValue(args[1], sheet, sheetIndex, 0),
        args[2] == null ? undefined : formulaArgumentNumberValue(args[2], sheet, sheetIndex, 0),
      ),
    };
  }
  if (name === "CHOOSEROWS") {
    return { supported: true, data: chooseRowsFormulaMatrix(sourceData, args.slice(1).map((arg) => formulaArgumentNumberValue(arg, sheet, sheetIndex, 0))) };
  }
  if (name === "CHOOSECOLS") {
    return { supported: true, data: chooseColsFormulaMatrix(sourceData, args.slice(1).map((arg) => formulaArgumentNumberValue(arg, sheet, sheetIndex, 0))) };
  }
  if (name === "TOCOL") {
    return {
      supported: true,
      data: tocolFormulaMatrix(
        sourceData,
        formulaArgumentNumberValue(args[1], sheet, sheetIndex, 0),
        formulaArgumentBoolean(args[2], false),
      ),
    };
  }
  if (name === "TOROW") {
    return {
      supported: true,
      data: torowFormulaMatrix(
        sourceData,
        formulaArgumentNumberValue(args[1], sheet, sheetIndex, 0),
        formulaArgumentBoolean(args[2], false),
      ),
    };
  }
  if (name === "WRAPROWS") {
    return {
      supported: true,
      data: wrapRowsFormulaMatrix(
        flattenFormulaMatrix(sourceData),
        formulaArgumentNumberValue(args[1], sheet, sheetIndex, 1),
        args[2] == null ? "#N/A" : formulaArgumentScalarValue(args[2], sheet, sheetIndex),
      ),
    };
  }
  if (name === "WRAPCOLS") {
    return {
      supported: true,
      data: wrapColsFormulaMatrix(
        flattenFormulaMatrix(sourceData),
        formulaArgumentNumberValue(args[1], sheet, sheetIndex, 1),
        args[2] == null ? "#N/A" : formulaArgumentScalarValue(args[2], sheet, sheetIndex),
      ),
    };
  }
  if (name === "EXPAND") {
    return {
      supported: true,
      data: expandFormulaMatrix(
        sourceData,
        formulaArgumentNumberValue(args[1], sheet, sheetIndex, 0),
        args[2] == null ? undefined : formulaArgumentNumberValue(args[2], sheet, sheetIndex, 0),
        args[3] == null ? "#N/A" : formulaArgumentScalarValue(args[3], sheet, sheetIndex),
      ),
    };
  }
  return {
    supported: true,
    data: sortFormulaMatrix(
      sourceData,
      formulaArgumentNumber(args[1], 1),
      formulaArgumentNumber(args[2], 1),
      formulaArgumentBoolean(args[3], false),
    ),
  };
}



function decodeFormulaRangeForSheet(value, baseSheet, baseSheetIndex = state.activeSheetIndex) {
  const source = String(value || "").trim().replace(/^=/, "").trim();
  const base = baseSheet || state.model?.sheets?.[baseSheetIndex] || activeSheet();
  const resolved = workbookRangeFromFormulaReference(source, base);
  if (resolved) return resolved;
  const text = source.replace(/\$/g, "");
  const match = text.match(/^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?(.+)$/);
  if (!match) return null;
  const range = decodeRange(match[3] || "");
  if (!range) return null;
  const sheetName = match[1]?.replace(/''/g, "'") || match[2] || "";
  const sheetIndex = sheetName ? state.model?.sheets?.findIndex((candidate) => candidate.name === sheetName) : baseSheetIndex;
  if (sheetIndex == null || sheetIndex < 0) return null;
  return { ...range, sheetIndex, sheetName: sheetName || baseSheet?.name || "" };
}



function formulaRangeValueMatrix(range) {
  const sheet = state.model?.sheets?.[range.sheetIndex];
  if (!sheet) return [];
  const sheetId = getSheetId(sheet.name);
  const rows = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    const values = [];
    for (let col = range.left; col <= range.right; col += 1) {
      values.push(formulaScalarValueForCell(sheet, sheetId, row, col));
    }
    rows.push(values);
  }
  return rows;
}



function formulaScalarValueForCell(sheet, sheetId, row, col) {
  const cell = sheet.cells[cellKey(row, col)];
  if (cell?.formula && state.hf && sheetId != null) {
    try {
      const value = state.hf.getCellValue({ sheet: sheetId, row: row - 1, col: col - 1 });
      return value instanceof SimpleRangeValue ? value.data?.[0]?.[0] ?? "" : value;
    } catch {
      return cell.cached ?? cellDisplayText(cell);
    }
  }
  return cell?.raw ?? cell?.cached ?? cellDisplayText(cell);
}



function formulaArgumentNumber(value, fallback) {
  if (value == null || String(value).trim() === "") return fallback;
  const number = Number(String(value).trim());
  return Number.isFinite(number) ? number : fallback;
}



function formulaArgumentBoolean(value, fallback) {
  if (value == null || String(value).trim() === "") return fallback;
  const text = String(value).trim().toUpperCase();
  if (text === "TRUE") return true;
  if (text === "FALSE") return false;
  const number = Number(text);
  return Number.isFinite(number) ? Boolean(number) : fallback;
}



function formulaArrayConstantMatrix(value, sheet, sheetIndex) {
  const text = String(value || "").trim();
  if (!text.startsWith("{") || !text.endsWith("}")) return null;
  const inner = text.slice(1, -1);
  return splitFormulaTopLevel(inner, new Set([";"])).map((rowText) =>
    splitFormulaTopLevel(rowText, new Set([","])).map((cellText) => formulaArgumentScalarValue(cellText, sheet, sheetIndex) ?? ""),
  );
}



function formulaArgumentScalarValue(value, sheet, sheetIndex) {
  if (value == null || String(value).trim() === "") return undefined;
  const text = String(value).trim();
  const quoted = text.match(/^"((?:[^"]|"")*)"$/u);
  if (quoted) return quoted[1].replace(/""/g, "\"");
  if (/^(TRUE|FALSE)$/iu.test(text)) return /^TRUE$/iu.test(text);
  const number = Number(text.replace(/,/g, ""));
  if (Number.isFinite(number) && /^[-+]?\d[\d,]*(?:\.\d+)?$/u.test(text)) return number;
  const arrayConstant = formulaArrayConstantMatrix(text, sheet, sheetIndex);
  if (arrayConstant) return arrayConstant[0]?.[0] ?? "";
  const range = decodeFormulaRangeForSheet(text, sheet, sheetIndex);
  if (range && range.top === range.bottom && range.left === range.right) {
    return formulaRangeValueMatrix(range)[0]?.[0] ?? "";
  }
  return text;
}



function formulaArgumentNumberValue(value, sheet, sheetIndex, fallback) {
  if (value == null || String(value).trim() === "") return fallback;
  const scalar = formulaArgumentScalarValue(value, sheet, sheetIndex);
  if (typeof scalar === "number" && Number.isFinite(scalar)) return scalar;
  if (typeof scalar === "boolean") return scalar ? 1 : 0;
  const number = Number(String(scalar ?? "").replace(/,/g, "").trim());
  return Number.isFinite(number) ? number : fallback;
}



function formulaArgumentMatrix(value, sheet, sheetIndex) {
  if (value == null || String(value).trim() === "") return [[]];
  const text = String(value).trim();
  const arrayConstant = formulaArrayConstantMatrix(text, sheet, sheetIndex);
  if (arrayConstant) return arrayConstant;
  const range = decodeFormulaRangeForSheet(text, sheet, sheetIndex);
  if (range) return formulaRangeValueMatrix(range);
  const spill = evaluateExcelArraySpillFormula(sheet, sheetIndex, text.startsWith("=") ? text : `=${text}`);
  if (spill.supported) return spill.data;
  return [[formulaArgumentScalarValue(text, sheet, sheetIndex) ?? ""]];
}



function formulaMatrixDimensions(data) {
  return {
    height: data.length,
    width: Math.max(0, ...data.map((row) => row.length)),
  };
}



function formulaMatrixValue(data, rowIndex, colIndex, fallback = "") {
  return data?.[rowIndex]?.[colIndex] ?? fallback;
}



function normalizeFormulaMatrix(data, fill = "") {
  const { height, width } = formulaMatrixDimensions(data);
  return Array.from({ length: height }, (_row, rowIndex) =>
    Array.from({ length: width }, (_col, colIndex) => formulaMatrixValue(data, rowIndex, colIndex, fill)),
  );
}



function flattenFormulaMatrix(data, scanByColumn = false) {
  const normalized = normalizeFormulaMatrix(data);
  const { height, width } = formulaMatrixDimensions(normalized);
  const values = [];
  if (scanByColumn) {
    for (let colIndex = 0; colIndex < width; colIndex += 1) {
      for (let rowIndex = 0; rowIndex < height; rowIndex += 1) values.push(normalized[rowIndex][colIndex]);
    }
  } else {
    for (let rowIndex = 0; rowIndex < height; rowIndex += 1) {
      for (let colIndex = 0; colIndex < width; colIndex += 1) values.push(normalized[rowIndex][colIndex]);
    }
  }
  return values;
}



function parseFormulaLambda(value) {
  const call = parseFormulaFunctionCall(value);
  if (!call || call.name !== "LAMBDA" || call.args.length < 2) return null;
  const names = call.args.slice(0, -1).map((arg) => String(arg || "").trim()).filter(Boolean);
  if (names.length !== call.args.length - 1 || names.some((name) => !/^[A-Za-z_][A-Za-z0-9_.]*$/u.test(name))) return null;
  return { names: names.map((name) => name.toLocaleLowerCase("en-US")), body: call.args[call.args.length - 1] };
}



function formulaLambdaApply(lambdaText, values, sheet, sheetIndex) {
  const lambda = parseFormulaLambda(lambdaText);
  if (!lambda || lambda.names.length !== values.length) return "#VALUE!";
  const env = new Map();
  lambda.names.forEach((name, index) => env.set(name, values[index]));
  const result = formulaLambdaEvaluate(lambda.body, env, sheet, sheetIndex);
  return formulaLambdaScalar(result);
}



function formulaLambdaScalar(value) {
  if (Array.isArray(value)) return value[0]?.[0] ?? "";
  return value;
}



function formulaLambdaNumeric(value, fallback = 0) {
  const scalar = formulaLambdaScalar(value);
  if (typeof scalar === "number" && Number.isFinite(scalar)) return scalar;
  if (typeof scalar === "boolean") return scalar ? 1 : 0;
  const number = Number(String(scalar ?? "").replace(/,/g, "").trim());
  return Number.isFinite(number) ? number : fallback;
}



function formulaLambdaTruthy(value) {
  return formulaTruthyValue(formulaLambdaScalar(value));
}



function formulaLambdaTopLevelOperator(source, operators) {
  const text = String(source || "");
  let depth = 0;
  let doubleQuoted = false;
  let singleQuoted = false;
  for (let index = text.length - 1; index >= 0; index -= 1) {
    const char = text[index];
    if (doubleQuoted) {
      if (char === "\"" && text[index - 1] !== "\"") doubleQuoted = false;
      else if (char === "\"" && text[index - 1] === "\"") index -= 1;
      continue;
    }
    if (singleQuoted) {
      if (char === "'" && text[index - 1] !== "'") singleQuoted = false;
      else if (char === "'" && text[index - 1] === "'") index -= 1;
      continue;
    }
    if (char === "\"") {
      doubleQuoted = true;
      continue;
    }
    if (char === "'") {
      singleQuoted = true;
      continue;
    }
    if (char === ")") {
      depth += 1;
      continue;
    }
    if (char === "(") {
      depth = Math.max(0, depth - 1);
      continue;
    }
    if (depth > 0 || !operators.includes(char)) continue;
    if ((char === "+" || char === "-") && (index === 0 || /[+\-*/^,(]/u.test(text[index - 1] || ""))) continue;
    return { index, operator: char };
  }
  return null;
}

export {
  activateObjectFormatRibbonForSelection,
  activateSheetForFormulaEditResult,
  activeFormulaReferenceEntries,
  activeFormulaReferenceEntriesForActiveSheet,
  activeFormulaReferenceEntryForCell,
  activeSelectionRange,
  activeSelectionRanges,
  activeSelectionRangesForGroupedOperation,
  activeSheetGroupEditIndexes,
  activeSheetGroupIndexes,
  affectedSheetIndexesForGroupedRanges,
  applyHeaderAutoFit,
  applyHeaderAutoFitToSheet,
  applyHeaderAutoFitToSheetGroup,
  applyMultilineCellDisplayBehavior,
  autoFitCellControlInset,
  autoFitCellDisplayText,
  autoFitCellTextHeight,
  autoFitCellTextWidth,
  autoFitColumnHeaderFromDoubleClick,
  autoFitHeaderFromDoubleClick,
  autoFitRowHeaderFromDoubleClick,
  autoScrollDeltaForAxis,
  autoScrollStep,
  beginFormulaBarEdit,
  bestColumnAutoFitWidth,
  bestHeaderAutoFitSize,
  bestRowAutoFitHeight,
  cancelFormulaBarEdit,
  captureDimensionForHistory,
  cellElementForPoint,
  cellPointFromClientPosition,
  cellPointFromElement,
  cellPointFromKey,
  cellPointFromPointerEvent,
  closeStatusBarContextMenu,
  closeStatusBarContextMenuOnEscape,
  closeStatusBarContextMenuOnOutsidePointer,
  closeValidationInputMessagePopup,
  coerceCellInput,
  coerceUserInput,
  commitActiveCellValueToSelection,
  commitCellValue,
  commitCellValueToSelection,
  commitCellValueToSheetGroup,
  commitEditableCellToSelection,
  commitFormulaBar,
  computeDisplayForCell,
  createFormulaSpillRenderCache,
  createHeaderAutoFitContext,
  cssEscape,
  decodeFormulaRangeForSheet,
  editedDateSerialForCell,
  editedTimeSerialForFormat,
  evaluateExcelArraySpillFormula,
  evaluateExcelArraySpillFormulaCached,
  excelCompatibleFormulaSpillValueAt,
  expandRangeForMerges,
  extendCellRangeSelection,
  fallbackSingleCellRangeForSheet,
  findMergeForCell,
  finishCellRangeSelection,
  finishFormulaBarResize,
  finishHeaderResize,
  flattenFormulaMatrix,
  formatPainterPointFromPointerEvent,
  formatStatusBarNumber,
  formulaArgumentBoolean,
  formulaArgumentMatrix,
  formulaArgumentNumber,
  formulaArgumentNumberValue,
  formulaArgumentScalarValue,
  formulaArrayConstantMatrix,
  formulaBarLineHeight,
  formulaEditCellElement,
  formulaLambdaApply,
  formulaLambdaNumeric,
  formulaLambdaScalar,
  formulaLambdaTopLevelOperator,
  formulaLambdaTruthy,
  formulaMatrixDimensions,
  formulaMatrixValue,
  formulaRangeValueMatrix,
  formulaScalarValueForCell,
  formulaSpillAnchorEntries,
  formulaSpillRenderCacheForSheet,
  getCellEditInput,
  getCellRawInput,
  getDisplayForCell,
  groupedRangesForSheet,
  handleStatusBarMenuItem,
  hasCellLineBreak,
  headerAutoFitMeasureContext,
  headerResizeGroupChanged,
  headerResizeTargetsForRanges,
  initialCellDisplayText,
  invalidateFormulaSpillAnchorCache,
  invalidateSheetUsedRangeCache,
  isActiveCell,
  isCellInActiveSelection,
  isCellInFormulaReferenceRange,
  isCellInRange,
  isCoordinateCellEvent,
  isSingleCellRange,
  isTimeOnlyDateFormat,
  loadStatusBarItems,
  markFormulaEditCell,
  materializeSelectionRange,
  materializeSelectionRangeForSheet,
  materializedRangesForSheet,
  mergeForCellFromMaps,
  mergedRangeForCell,
  multilineCellAutoHeight,
  multilineMeasureContext,
  mutateShapeObjects,
  nativeAnchorFromSheetPixel,
  normalizeFormulaMatrix,
  normalizeRange,
  openStatusBarContextMenu,
  parseFormattedNumericInput,
  parseFormulaLambda,
  positionRangeFrameElement,
  positionValidationInputMessagePopup,
  primeWorkbookFormulaSpillAnchorCache,
  primeWorkbookUsedRangeCache,
  rangeToLabel,
  rangesToLabel,
  refreshChartsForGroupedRanges,
  rejectInvalidCellInput,
  renderDisplayCacheKey,
  restoreGroupedSheetsAfterFailure,
  retryInvalidCellInput,
  runSelectionAutoScroll,
  saveStatusBarItems,
  scheduleHeaderResizeRender,
  scheduleSelectionUiSync,
  selectableCellPoint,
  selectedHeaderRangesForKind,
  selectedObjectFormatMode,
  selectionAutoScrollCanContinue,
  selectionAutoScrollVelocity,
  selectionDragAutoScrollVelocity,
  selectionDragSupportsAutoScroll,
  selectionRangeFromPoints,
  selectionStatusSelectedCellCount,
  selectionStatusStats,
  setFormulaBarLines,
  sheetIndexFromPixel,
  sheetPixelFromClientPosition,
  showValidationErrorForCell,
  startColumnHeaderResize,
  startFormulaBarResize,
  startHeaderResize,
  startRowHeaderResize,
  statusBarCellDisplay,
  statusBarCellHasData,
  statusBarCellMayHaveData,
  statusBarCellValue,
  stopSelectionAutoScroll,
  syncContextualObjectFormatRibbon,
  syncFormulaEditPreview,
  syncFormulaExpandButton,
  syncRangeFrameHandle,
  syncRibbonControlsForShape,
  syncShapeFormatControls,
  toggleFormulaBarLines,
  unionMaterializedRanges,
  unionRanges,
  uniqueEditableTargetsForRanges,
  updateCellModel,
  updateCellRangeSelectionFromClient,
  updateContextualRibbonTabs,
  updateFormulaBarResize,
  updateHeaderResize,
  updateRangeFrameElements,
  updateSelectionAutoScroll,
  updateSelectionFrames,
  updateSelectionStatusBar,
  updateSelectionUi,
  updateStatusBarControlVisibility,
  updateValidationInputMessagePopup,
  withDisplayCalculationCache,
  writeCellContent,
};
