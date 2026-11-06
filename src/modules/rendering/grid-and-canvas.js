import { measuredLineRanges, sliceStyledLine } from "./text-wrap.js";
// Feature module; shared bindings are imported explicitly.
import { reconcileViewportGrid } from './viewport-reconciliation.js';
import {
  $horizontalScrollbar,
  $ribbonPanels,
  $sheetHost,
  CELL_BORDER_SIDES,
  COLUMN_HEADER_HEIGHT,
  CanvasRenderer,
  DATA_OPERATION_CELL_LIMIT,
  EMU_PER_PIXEL,
  MERGE_CONTENT_WARNING_MESSAGE,
  PAGE_LAYOUT_MARGIN_PRESETS,
  PAGE_LAYOUT_PAPER_SIZES,
  TEXT_OVERFLOW_LOOKBACK_MAX_PX,
  TEXT_OVERFLOW_LOOKBACK_MIN_PX,
  TEXT_OVERFLOW_LOOKBACK_VIEWPORT_FACTOR,
  state,
} from "../../app-core.js";
import {
  refreshQueriesConnectionsPane,
} from "../data/connections-and-queries.js";
import {
  buildOutlineControlMaps,
  renderCornerHeader,
} from "../data/forecast-and-statistics.js";
import {
  startShapeDrawing,
} from "../drawing/charts-and-insertion.js";
import {
  isLineShapeObject,
  normalizedHeaderFooterSettings,
  openHeaderFooterDialog,
  shapePlainTextFromRuns,
} from "../drawing/shape-rendering.js";
import {
  refreshShapeSource,
  removeLineConnectionPreviewOverlay,
} from "../drawing/transforms-and-connectors.js";
import {
  applyCellEditCssPatchToSelection,
  applyCellEditUnderlinePreset,
  normalizeCellEditInlineCss,
  normalizeCellEditRichTextForModel,
  renderAutoFillOptions,
  toggleCellEditCssForSelection,
} from "../editing/autofill.js";
import {
  normalizeColorValue,
  underlinePresetForValue,
} from "../editing/ink-and-clipboard.js";
import {
  selectedObjectSummary,
} from "../editing/keyboard-and-paste.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  autoFitCellControlInset,
  createFormulaSpillRenderCache,
  expandRangeForMerges,
  findMergeForCell,
  formulaSpillRenderCacheForSheet,
  getDisplayForCell,
  invalidateFormulaSpillAnchorCache,
  invalidateSheetUsedRangeCache,
  isSingleCellRange,
  materializeSelectionRangeForSheet,
  materializedRangesForSheet,
  normalizeRange,
  rangeToLabel,
  rangesToLabel,
  restoreGroupedSheetsAfterFailure,
  selectableCellPoint,
  sheetIndexFromPixel,
  sheetPixelFromClientPosition,
  syncShapeFormatControls,
  updateSelectionStatusBar,
  updateSelectionUi,
  updateValidationInputMessagePopup,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  canvasDevicePixelRatio,
  canvasFont,
  canvasFontDisplayScale,
  canvasFontSize,
  canvasLayerBounds,
  canvasLineBaseline,
  canvasTextLineHeight,
  cellHasDomContent,
  cellHasFill,
  coordinateToPixels,
  cssFontSizeToPoints,
  paintCanvasPhonetics,
  paintCanvasStrikethrough,
  paintCanvasUnderline,
  phoneticRunsForLine,
  renderCellHtml,
  renderCopyFrame,
  renderFormulaReferenceFrame,
  renderImageLayer,
  renderImageSelectionFrames,
  renderInkLayer,
  renderSelectionFrame,
  selectedImageObjects,
  selectionToolbarPoint,
  sheetColumnsRightToLeft,
  shouldHideGridBottom,
  shouldHideGridRight,
  syncDataRibbonControls,
  updateFormulaBarForSelection,
  visualGridColumnStart,
  wrapCanvasText,
} from "../editing/selection-and-cell-text.js";
import {
  captureCellsForHistory,
  captureCellsForHistoryRanges,
  captureSheetForHistory,
  captureSheetsForHistory,
  cellEffectiveCss,
  cloneCellModel,
  cloneJsonLikeValue,
  cloneMerges,
  internCellsInRange,
  internCellsInRanges,
  mutableBordersObject,
  mutableCssObject,
  pushHistory,
  removeSheetStyleRangesAtCell,
  sheetDisplayCell,
  sheetStyleRangeCell,
  stableObjectKey,
  storeCellModelForSheet,
} from "../editing/structure-and-history.js";
import {
  formatDecimalPlaces,
  formatUsesSeparator,
  isRangeInsideRange,
  renderAnnotationEditor,
  renderPersistentNotePopups,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  columnName,
  columnWidthToPixels,
  currencySymbol,
  escapeAttr,
  escapeHtml,
  isDateFormat,
  rowHeightToPixels,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  startInkDrawing,
} from "../ui/printing-and-home-tools.js";
import {
  ensureSheetCoversViewport,
  ensureSheetSize,
  indexForPixelExtent,
  parseCellKey,
  rowHeaderWidthForSheet,
} from "../ui/ribbon-and-links.js";
import {
  ensureActiveSheetFormatCellsAllowed,
  ensureSheetProtectionPermission,
} from "../ui/sheet-protection-and-translation.js";
import {
  captureWorkbookViewState,
  normalizeFormulaAudits,
  refreshActiveSheetViewSnapshot,
  sheetHeadingsVisible,
  splitPaneCounts,
  syncFormulaWatchWindow,
  toggleActiveSheetWhiteSpace,
} from "../ui/views-and-formula-tools.js";
import {
  buildFormulaEngine,
  ensureActiveSheetVisible,
  isSheetVisible,
  renderWorkbook,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  dateSerialFromPersistedValue,
  normalizePaneState,
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  activeSheet,
  borderLineHtml,
  borderLinesForSheet,
  buildMergeMaps,
  getCellFormula,
  getSheetId,
  gridTrackStyle,
  paintCanvasBorders,
  sheetGridClassName,
  textOverflowWidth,
  updateHorizontalScrollbar,
  visibleGridRange,
} from "../workbook/saving-and-export.js";
import {
  ensurePageLayoutSettings,
  normalizePageLayoutBackground,
  normalizePrintRange,
  normalizedPageLayoutSettings,
  pageBreakTargetsFromAxisIndex,
  pageLayoutFitTargets,
  printRangeForSheet,
  selectPageBreakTargets,
} from "../workbook/vba-and-printing.js";
import {
  saveActiveWindowViewState,
} from "../workbook/views-and-macros.js";


function finishSheetStructureHistoryRestore(options = {}) {
  state.selectedSheetIndexes = [];
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionStart = null;
  state.selectionAnchor = null;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.headerSizeDialog = null;
  state.nameDialog = null;
  state.linkDialog = null;
  state.symbolDialog = null;
  state.annotationEditor = null;
  ensureActiveSheetVisible();
  if (options.selectedSheetNames?.length) {
    state.selectedSheetIndexes = options.selectedSheetNames
      .map((name) => state.model.sheets.findIndex((sheet) => sheet.name === name))
      .filter((index) => index >= 0 && isSheetVisible(state.model.sheets[index]));
    if (state.selectedSheetIndexes.length && !state.selectedSheetIndexes.includes(state.activeSheetIndex)) {
      state.activeSheetIndex = state.selectedSheetIndexes[0];
    }
  }
  state.hf = buildFormulaEngine(state.model);
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  refreshQueriesConnectionsPane();
  updateSelectionUi();
  updateFormulaBarForSelection();
}



function restoreCellsFromHistory(sheet, cells) {
  cells.forEach(({ row, col, cell }) => {
    ensureSheetSize(sheet, row, col);
    const key = cellKey(row, col);
    const previousFormula = sheet.cells?.[key]?.formula || null;
    removeSheetStyleRangesAtCell(sheet, row, col);
    if (cell) {
      const nextCell = cloneCellModel(cell);
      nextCell.row = row;
      nextCell.col = col;
      storeCellModelForSheet(sheet, key, nextCell);
    } else {
      if (previousFormula) invalidateFormulaSpillAnchorCache(sheet);
      delete sheet.cells[key];
    }
  });
  invalidateSheetUsedRangeCache(sheet);
}



function restoreDimensionsFromHistory(sheet, dimensions) {
  const items = Array.isArray(dimensions) ? dimensions : [dimensions];
  items.forEach((item) => {
    if (!item?.kind || !item.index || !Number.isFinite(Number(item.size))) return;
    ensureSheetSize(sheet, item.kind === "row" ? item.index : sheet.rowCount, item.kind === "col" ? item.index : sheet.colCount);
    const target = item.kind === "col" ? sheet.columns[item.index - 1] : sheet.rows[item.index - 1];
    if (!target) return;
    target[item.kind === "col" ? "width" : "height"] = item.size;
  });
}



function applyCssPatchToSheetRanges(sheet, ranges, cssPatch) {
  const richTextOverrideKeys = cssPatchRichTextOverrideKeys(cssPatch);
  ranges.forEach((range) => {
    forEachCellInRange(range, (row, col) => {
      const cell = ensureCellModel(sheet, row, col);
      preserveCellRichTextForWholeCellCssPatch(cell, richTextOverrideKeys);
      Object.entries(cssPatch).forEach(([key, value]) => {
        if (value == null || value === "") {
          delete cell.css[key];
          if (key === "backgroundColor") cell.hasFill = false;
          if (key === "color") cell.hasFontColor = false;
        } else {
          cell.css[key] = value;
          if (key === "backgroundColor") cell.hasFill = true;
          if (key === "color") cell.hasFontColor = true;
        }
      });
    });
  });
  internCellsInRanges(sheet, ranges);
}



function applyCssToSheetGroup(cssPatch, label, targetRanges, activeRange) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !targetRanges.length) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model.sheets[sheetIndex];
    if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではセルの書式を変更できません。")) return true;
    const ranges = materializedRangesForSheet(sheet, targetRanges, sheetIndex);
    if (ranges.length) items.push({ sheetIndex, sheet, ranges });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => applyCssPatchToSheetRanges(item.sheet, item.ranges, cssPatch));
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: activeRange || targetRanges[0],
      selectionAfter: activeRange || targetRanges[0],
      selectionBeforeRanges: targetRanges.map((range) => ({ ...range })),
      selectionAfterRanges: targetRanges.map((range) => ({ ...range })),
      label: `${rangesToLabel(targetRanges)} ${sheetIndexes.length} シート書式：${label}`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${rangesToLabel(targetRanges)}に適用しました：${label}`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`書式を適用できませんでした：${error.message || error}`);
  }
  return true;
}



function applyCssToSelection(cssPatch, label) {
  if (applyCellEditCssPatchToSelection(cssPatch, label)) return;
  const activeRange = activeSelectionRange();
  const ranges = activeSelectionRanges();
  const targetRanges = ranges.length ? ranges : activeRange ? [activeRange] : [];
  if (!targetRanges.length) {
    if (applyShapeTextCssPatchToSelection(cssPatch, label)) return;
    setStatus("セルを選択してください。");
    return;
  }

  const sheet = activeSheet();
  if (!ensureActiveSheetFormatCellsAllowed()) return;
  if (applyCssToSheetGroup(cssPatch, label, targetRanges, activeRange)) return;
  const before = captureCellsForHistoryRanges(sheet, targetRanges);
  applyCssPatchToSheetRanges(sheet, targetRanges, cssPatch);
  const after = captureCellsForHistoryRanges(sheet, targetRanges);
  const labelText = rangesToLabel(targetRanges);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: activeRange || targetRanges[0],
    selectionAfter: activeRange || targetRanges[0],
    selectionBeforeRanges: targetRanges,
    selectionAfterRanges: targetRanges,
    label: `${labelText} 書式：${label}`,
  });
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(`${labelText} に適用しました：${label}`);
}



function cssPatchResetsCellRichText(cssPatch = {}) {
  return cssPatchRichTextOverrideKeys(cssPatch).length > 0;
}



function cssPatchRichTextOverrideKeys(cssPatch = {}) {
  return ["fontFamily", "fontSize", "color", "fontWeight", "fontStyle", "textDecoration"].filter((key) =>
    Object.prototype.hasOwnProperty.call(cssPatch, key),
  );
}



function preserveCellRichTextForWholeCellCssPatch(cell, overrideKeys = []) {
  if (!overrideKeys.length || !Array.isArray(cell?.richText) || !cell.richText.length) return;
  const keys = new Set(overrideKeys);
  const runs = cell.richText
    .map((run) => {
      const text = String(run?.text ?? "");
      if (!text) return null;
      const css = normalizeCellEditInlineCss(run?.css || {}, { keepNormal: true });
      keys.forEach((key) => delete css[key]);
      return Object.keys(css).length ? { text, css } : { text };
    })
    .filter(Boolean);
  const plainText = runs.map((run) => run.text).join("");
  const normalized = normalizeCellEditRichTextForModel(runs, plainText);
  if (normalized) cell.richText = normalized;
  else delete cell.richText;
}



function startFormatPainter(options = {}) {
  const sourceRanges = activeSelectionRanges();
  const sourceRange = sourceRanges[0] || activeSelectionRange();
  if (!sourceRange) {
    setStatus("書式をコピーするセルを選択してください。");
    return;
  }
  const cellCount = (sourceRange.bottom - sourceRange.top + 1) * (sourceRange.right - sourceRange.left + 1);
  if (cellCount > DATA_OPERATION_CELL_LIMIT) {
    showMessageBox("選択範囲が大きすぎるため、書式をコピーできません。", { title: "Cht WebSheet" });
    return;
  }
  state.formatPainter = {
    sheetIndex: state.activeSheetIndex,
    range: { ...sourceRange },
    formats: captureFormatPainterFormats(activeSheet(), sourceRange),
    merges: captureFormatPainterMerges(activeSheet(), sourceRange),
    rowCount: sourceRange.bottom - sourceRange.top + 1,
    colCount: sourceRange.right - sourceRange.left + 1,
    repeat: options.repeat === true,
  };
  setFormatPainterActive(true);
  setStatus(
    options.repeat === true
      ? `${rangeToLabel(sourceRange)} の書式をコピーしました。貼り付け先を続けて選択できます。Esc キーで終了します。`
      : `${rangeToLabel(sourceRange)} の書式をコピーしました。貼り付け先のセルを選択してください。`,
  );
}



function toggleFormatPainter(options = {}) {
  if (state.formatPainter) {
    if (options.repeat === true) {
      state.formatPainter.repeat = true;
      setFormatPainterActive(true);
      setStatus(`${rangeToLabel(state.formatPainter.range)} の書式を連続して貼り付けできます。Esc キーで終了します。`);
      return;
    }
    clearFormatPainter();
    return;
  }
  startFormatPainter(options);
}



function isFormatPainterRepeatTrigger(event) {
  const nativeEvent = event?.originalEvent || event;
  return Number(nativeEvent?.detail || 0) >= 2;
}



function setFormatPainterActive(active) {
  $ribbonPanels.find('[data-ribbon-command="format-painter"]').toggleClass("is-active", Boolean(active)).attr("aria-pressed", String(Boolean(active)));
}



function clearFormatPainter(options = {}) {
  if (!state.formatPainter) return;
  state.formatPainter = null;
  setFormatPainterActive(false);
  if (!options.silent) setStatus("書式のコピー/貼り付けをキャンセルしました。");
}



function captureFormatPainterFormats(sheet, range) {
  const formats = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    const line = [];
    for (let col = range.left; col <= range.right; col += 1) {
      line.push(cloneCellFormat(sheet.cells[cellKey(row, col)] || null));
    }
    formats.push(line);
  }
  return formats;
}



function captureFormatPainterMerges(sheet, range) {
  return (sheet?.merges || [])
    .filter((merge) => isRangeInsideRange(merge, range))
    .map((merge) => ({
      top: merge.top - range.top + 1,
      left: merge.left - range.left + 1,
      bottom: merge.bottom - range.top + 1,
      right: merge.right - range.left + 1,
    }));
}



function cloneCellFormat(cell) {
  return {
    css: { ...(cell?.css || {}) },
    hasFill: Boolean(cell?.hasFill),
    hasFontColor: Boolean(cell?.hasFontColor),
    borders: cloneCellBorders(cell?.borders || {}),
    numFmt: cell?.numFmt || "",
    protection: cloneJsonLikeValue(cell?.protection || null),
  };
}



function cloneCellBorders(borders = {}) {
  const next = {};
  CELL_BORDER_SIDES.forEach((side) => {
    if (borders?.[side]) next[side] = { ...borders[side] };
  });
  return next;
}



function applyFormatPainterAtPoint(point) {
  const painter = state.formatPainter;
  if (!painter || !point) return;
  const targetRange = normalizeRange(point, {
    sheetIndex: state.activeSheetIndex,
    row: point.row + painter.rowCount - 1,
    col: point.col + painter.colCount - 1,
  });
  applyFormatPainterToRange(targetRange);
}



function applyFormatPainterToRange(targetRange) {
  const painter = state.formatPainter;
  if (!painter || !targetRange) return;
  const sheet = activeSheet();
  const cellCount = Math.max(0, targetRange.bottom - targetRange.top + 1) * Math.max(0, targetRange.right - targetRange.left + 1);
  if (cellCount > DATA_OPERATION_CELL_LIMIT) {
    showMessageBox("選択範囲が大きすぎるため、書式を貼り付けできません。", { title: "Cht WebSheet" });
    return;
  }
  if (!ensureActiveSheetFormatCellsAllowed("保護されたシートでは書式を貼り付けできません。")) {
    clearFormatPainter({ silent: true });
    return;
  }
  ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
  const before = captureCellsForHistory(sheet, targetRange);
  const mergesBefore = cloneMerges(sheet.merges);
  sheet.merges = Array.isArray(sheet.merges) ? sheet.merges : [];
  removeIntersectingMerges(sheet, targetRange);
  for (let row = targetRange.top; row <= targetRange.bottom; row += 1) {
    for (let col = targetRange.left; col <= targetRange.right; col += 1) {
      const sourceRow = (row - targetRange.top) % painter.rowCount;
      const sourceCol = (col - targetRange.left) % painter.colCount;
      applyCellFormat(sheet, row, col, painter.formats[sourceRow]?.[sourceCol] || {});
    }
  }
  applyFormatPainterMerges(sheet, targetRange, painter);
  internCellsInRange(sheet, targetRange);
  const after = captureCellsForHistory(sheet, targetRange);
  const mergesAfter = cloneMerges(sheet.merges);
  const selectionBefore = painter.range;
  const keepActive = painter.repeat === true;
  if (!keepActive) clearFormatPainter({ silent: true });
  state.selected = { sheetIndex: state.activeSheetIndex, row: targetRange.top, col: targetRange.left };
  state.selectionRange = expandRangeForMerges(sheet, targetRange) || targetRange;
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    mergesBefore,
    mergesAfter,
    selectionBefore,
    selectionAfter: state.selectionRange,
    selectionBeforeRanges: [selectionBefore],
    selectionAfterRanges: [state.selectionRange],
    label: `${rangeToLabel(targetRange)} 書式のコピー/貼り付け`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(
    keepActive
      ? `${rangeToLabel(targetRange)} に書式を貼り付けました。続けて貼り付け先を選択できます。`
      : `${rangeToLabel(targetRange)} に書式を貼り付けました。`,
  );
}



function applyFormatPainterMerges(sheet, targetRange, painter) {
  const sourceMerges = Array.isArray(painter?.merges) ? painter.merges : [];
  if (!sourceMerges.length || !painter.rowCount || !painter.colCount) return;
  for (let rowBase = targetRange.top; rowBase <= targetRange.bottom; rowBase += painter.rowCount) {
    for (let colBase = targetRange.left; colBase <= targetRange.right; colBase += painter.colCount) {
      sourceMerges.forEach((merge) => {
        const nextMerge = {
          top: rowBase + merge.top - 1,
          left: colBase + merge.left - 1,
          bottom: rowBase + merge.bottom - 1,
          right: colBase + merge.right - 1,
        };
        if (!isSingleCellRange(nextMerge) && isRangeInsideRange(nextMerge, targetRange)) {
          sheet.merges.push(nextMerge);
        }
      });
    }
  }
}



function applyCellFormat(sheet, row, col, format = {}) {
  const cell = ensureCellModel(sheet, row, col);
  cell.css = { ...(format.css || {}) };
  cell.hasFill = Boolean(format.hasFill);
  cell.hasFontColor = Boolean(format.hasFontColor);
  cell.borders = cloneCellBorders(format.borders || {});
  cell.numFmt = format.numFmt || "";
  if (format.protection) cell.protection = cloneJsonLikeValue(format.protection);
  else delete cell.protection;
  if (cssPatchResetsCellRichText(cell.css || {})) delete cell.richText;
  storeCellModelForSheet(sheet, cellKey(row, col), cell);
}



function applyShapeTextCssPatchToSelection(cssPatch, label) {
  const objects = selectedImageObjects().filter((image) => image?.shape && !isLineShapeObject(image));
  if (!objects.length) return false;
  const shapePatch = shapeTextPatchFromCss(cssPatch);
  if (!Object.keys(shapePatch).length) return false;
  const sheet = activeSheet();
  const sheetBefore = captureSheetForHistory(sheet);
  objects.forEach((image) => {
    const text = image.shape?.text || shapePlainTextFromRuns(image.shape?.textRuns) || "";
    image.shape = {
      ...(image.shape || {}),
      ...shapePatch,
      text,
    };
    delete image.shape.textRuns;
    refreshShapeSource(image);
  });
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${selectedObjectSummary(objects)} テキスト書式：${label}`,
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`${selectedObjectSummary(objects)} のテキストに適用しました：${label}`);
  return true;
}



function shapeTextPatchFromCss(cssPatch = {}) {
  const patch = {};
  Object.entries(cssPatch).forEach(([key, value]) => {
    switch (key) {
      case "fontFamily":
        if (value) patch.textFontFamily = value;
        break;
      case "fontSize":
        {
          const pt = cssFontSizeToPoints(value);
          if (pt) patch.textFontSize = Math.max(6, Math.min(72, Math.round(pt * 10) / 10));
        }
        break;
      case "color":
        patch.textColor = normalizeColorValue(value) || "#1f2937";
        break;
      case "backgroundColor":
        patch.fillColor = normalizeColorValue(value) || "none";
        if (patch.fillColor !== "none") patch.fillGradient = null;
        break;
      case "fontWeight":
        patch.textBold = value === "700" || value === "bold";
        break;
      case "fontStyle":
        patch.textItalic = value === "italic";
        break;
      case "textDecoration":
        patch.textUnderline = underlinePresetForValue(value)?.value || "";
        break;
      case "textAlign":
        if (["left", "center", "right"].includes(value)) patch.textAlign = value;
        break;
      case "verticalAlign":
        if (value === "top") patch.textVAlign = "top";
        if (value === "middle") patch.textVAlign = "middle";
        if (value === "bottom") patch.textVAlign = "bottom";
        break;
      case "alignItems":
        if (value === "flex-start") patch.textVAlign = "top";
        if (value === "center") patch.textVAlign = "middle";
        if (value === "flex-end") patch.textVAlign = "bottom";
        break;
      default:
        break;
    }
  });
  return patch;
}



function toggleWrapTextForSelection() {
  const ranges = activeSelectionRanges();
  if (!ranges.length) {
    setStatus("セルを選択してください。");
    return;
  }

  const sheet = activeSheet();
  const point = selectionToolbarPoint() || { row: ranges[0].top, col: ranges[0].left };
  const css = cellEffectiveCss(sheet, sheet.cells[cellKey(point.row, point.col)] || {});
  const shouldWrap = css.whiteSpace !== "pre-wrap";
  applyCssToSelection(
    shouldWrap
      ? { whiteSpace: "pre-wrap", overflowWrap: "anywhere" }
      : { whiteSpace: "", overflowWrap: "" },
    "折り返して全体を表示する",
  );
}



function applyNumberFormatToSheetGroup(numFmt, label, targetRanges, activeRange) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !targetRanges.length) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model.sheets[sheetIndex];
    if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートでは表示形式を変更できません。")) return true;
    const ranges = materializedRangesForSheet(sheet, targetRanges, sheetIndex);
    if (ranges.length) items.push({ sheetIndex, sheet, ranges });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      item.ranges.forEach((range) => {
        forEachCellInRange(range, (row, col) => {
          const cell = ensureCellModel(item.sheet, row, col);
          applyCellNumberFormat(cell, numFmt);
        });
      });
      internCellsInRanges(item.sheet, item.ranges);
    });
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: activeRange || targetRanges[0],
      selectionAfter: activeRange || targetRanges[0],
      selectionBeforeRanges: targetRanges.map((range) => ({ ...range })),
      selectionAfterRanges: targetRanges.map((range) => ({ ...range })),
      label: `${rangesToLabel(targetRanges)} ${sheetIndexes.length} シート表示形式：${label}`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${rangesToLabel(targetRanges)}に表示形式を適用しました：${label}`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`表示形式を適用できませんでした：${error.message || error}`);
  }
  return true;
}



function applyNumberFormatToSelection(numFmt, label) {
  const activeRange = activeSelectionRange();
  const ranges = activeSelectionRanges();
  const targetRanges = ranges.length ? ranges : activeRange ? [activeRange] : [];
  if (!targetRanges.length) {
    setStatus("セルを選択してください。");
    return;
  }

  const sheet = activeSheet();
  if (!ensureActiveSheetFormatCellsAllowed("保護されたシートでは表示形式を変更できません。")) return;
  if (applyNumberFormatToSheetGroup(numFmt, label, targetRanges, activeRange)) return;
  const before = captureCellsForHistoryRanges(sheet, targetRanges);
  targetRanges.forEach((range) => {
    forEachCellInRange(range, (row, col) => {
      const cell = ensureCellModel(sheet, row, col);
      applyCellNumberFormat(cell, numFmt);
    });
  });
  internCellsInRanges(sheet, targetRanges);
  const after = captureCellsForHistoryRanges(sheet, targetRanges);
  const labelText = rangesToLabel(targetRanges);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: activeRange || targetRanges[0],
    selectionAfter: activeRange || targetRanges[0],
    selectionBeforeRanges: targetRanges,
    selectionAfterRanges: targetRanges,
    label: `${labelText} 表示形式：${label}`,
  });
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(`${labelText} に表示形式を適用しました：${label}`);
}



function applyCellNumberFormat(cell, numFmt) {
  if (!cell) return;
  const nextNumFmt = String(numFmt || "");
  const previousNumFmt = String(cell.numFmt || "");
  if (isDateFormat(nextNumFmt)) {
    normalizeCellDateValueForNumberFormat(cell, nextNumFmt);
  }
  cell.numFmt = nextNumFmt;
  if (nextNumFmt !== previousNumFmt && cellCanRecomputeDisplayText(cell)) {
    cell.display = "";
  }
}



function normalizeCellDateValueForNumberFormat(cell, numFmt) {
  if (!cell || !isDateFormat(numFmt)) return;
  const serial =
    dateSerialFromValueForNumberFormat(cell.raw, numFmt) ??
    dateSerialFromValueForNumberFormat(cell.cached, numFmt) ??
    dateSerialFromValueForNumberFormat(cell.display, numFmt);
  if (serial == null) return;
  if (cell.formula) cell.cached = serial;
  else cell.raw = serial;
  if (cell.cached != null && !cell.formula) cell.cached = serial;
  if (cell.display) cell.display = "";
}



function dateSerialFromValueForNumberFormat(value, numFmt = "") {
  if (value == null || value === "") return null;
  const serial = dateSerialFromPersistedValue(value, numFmt);
  if (serial == null) return null;
  const number = Number(serial);
  return Number.isFinite(number) ? number : null;
}



function cellCanRecomputeDisplayText(cell) {
  return Boolean(cell && (cell.formula || cell.raw != null || cell.cached != null));
}



function numberFormatForRibbon(value) {
  switch (value) {
    case "number":
      return "0";
    case "currency":
      return "¥#,##0";
    case "percent":
      return "0%";
    case "date":
      return "yyyy/mm/dd";
    case "text":
      return "@";
    default:
      return "";
  }
}



function adjustSelectionDecimalPlaces(delta) {
  const range = activeSelectionRange();
  if (!range) {
    setStatus("セルを選択してください。");
    return;
  }
  const sheet = activeSheet();
  const point = selectionToolbarPoint() || { row: range.top, col: range.left };
  const current = sheet.cells[cellKey(point.row, point.col)]?.numFmt || "0";
  const decimals = Math.max(0, Math.min(6, formatDecimalPlaces(current) + delta));
  const next = numberFormatWithDecimals(current, decimals);
  applyNumberFormatToSelection(next, `小数点以下 ${decimals} 桁`);
}



function numberFormatWithDecimals(numFmt, decimals) {
  const format = String(numFmt || "0");
  const zeros = decimals > 0 ? `.${"0".repeat(decimals)}` : "";
  if (/%/.test(format)) return `${formatUsesSeparator(format) ? "#,##" : ""}0${zeros}%`;
  const currency = currencySymbol(format);
  return `${currency || ""}${formatUsesSeparator(format) ? "#,##" : ""}0${zeros}`;
}



function toggleStyleForSelection(property, activeValue, inactiveValue, label) {
  if (toggleCellEditCssForSelection(property, activeValue, inactiveValue, label)) return;
  const range = activeSelectionRange();
  if (!range) {
    const image = selectedImageObjects().find((item) => item?.shape && !isLineShapeObject(item));
    if (image) {
      const active = property === "fontWeight" ? Boolean(image.shape?.textBold) : property === "fontStyle" ? Boolean(image.shape?.textItalic) : false;
      applyShapeTextCssPatchToSelection({ [property]: active ? inactiveValue : activeValue }, label);
      return;
    }
    setStatus("セルを選択してください。");
    return;
  }
  const sheet = activeSheet();
  const point = selectionToolbarPoint() || { row: range.top, col: range.left };
  const activeCell = sheet.cells[cellKey(point.row, point.col)];
  const nextValue = activeCell?.css?.[property] === activeValue ? inactiveValue : activeValue;
  applyCssToSelection({ [property]: nextValue }, label);
}



function toggleTextDecorationForSelection(token, label) {
  const normalizedToken = token === "line-through" ? "line-through" : "underline";
  if (normalizedToken === "underline" && applyCellEditUnderlinePreset("underline", { toggle: true }, label)) return;
  const range = activeSelectionRange();
  if (!range) {
    if (normalizedToken !== "underline") {
      setStatus("セルを選択してください。");
      return;
    }
    const image = selectedImageObjects().find((item) => item?.shape && !isLineShapeObject(item));
    if (!image) {
      setStatus("セルを選択してください。");
      return;
    }
    const nextValue = underlinePresetForValue(image.shape?.textUnderline)?.value ? "" : "underline";
    applyShapeTextCssPatchToSelection({ textDecoration: nextValue }, label);
    return;
  }

  const sheet = activeSheet();
  const point = selectionToolbarPoint() || { row: range.top, col: range.left };
  const activeCell = sheet.cells[cellKey(point.row, point.col)];
  const current = String(activeCell?.css?.textDecoration || "");
  const enabled = !textDecorationHasToken(current, normalizedToken);
  applyCssToSelection({ textDecoration: textDecorationWithToken(current, normalizedToken, enabled) }, label);
}



function textDecorationHasToken(value, token) {
  const text = String(value || "").toLowerCase();
  if (token === "underline") return text.includes("underline");
  if (token === "line-through") return text.includes("line-through");
  return false;
}



function textDecorationWithToken(value, token, enabled) {
  const current = String(value || "").toLowerCase();
  const hasUnderline = token === "underline" ? enabled : current.includes("underline");
  const hasLineThrough = token === "line-through" ? enabled : current.includes("line-through");
  const hasDouble = hasUnderline && current.includes("double") && token !== "underline";
  const parts = [];
  if (hasUnderline) parts.push("underline");
  if (hasLineThrough) parts.push("line-through");
  if (hasDouble) parts.push("double");
  return parts.join(" ");
}



function borderTargetForCell(sheet, mergeMaps, row, col) {
  const key = cellKey(row, col);
  const masterKey = mergeMaps.coveredToMaster.get(key) || key;
  const merge = mergeMaps.master.get(masterKey);
  if (merge) {
    return {
      key: masterKey,
      row: merge.top,
      col: merge.left,
      range: merge,
    };
  }
  return {
    key,
    row,
    col,
    range: { top: row, left: col, bottom: row, right: col },
  };
}



function applyBorderPresetToSheetRange(sheet, range, preset) {
  const thin = { width: 1, style: "solid", color: "#000000", weight: 2 };
  const thick = { width: 3, style: "solid", color: "#000000", weight: 4 };
  const boundaryRange = expandRangeForMerges(sheet, range) || range;
  const mergeMaps = buildMergeMaps(sheet.merges || []);
  const touchedKeys = new Set();
  const handledTargets = new Set();
  forEachCellInRange(boundaryRange, (row, col) => {
    const target = borderTargetForCell(sheet, mergeMaps, row, col);
    if (handledTargets.has(target.key)) return;
    handledTargets.add(target.key);

    const cell = ensureCellModel(sheet, target.row, target.col);
    cell.borders = cell.borders || {};
    touchedKeys.add(target.key);

    const visualRange = target.range;
    const onTop = visualRange.top === boundaryRange.top;
    const onRight = visualRange.right === boundaryRange.right;
    const onBottom = visualRange.bottom === boundaryRange.bottom;
    const onLeft = visualRange.left === boundaryRange.left;

    if (preset === "none") {
      cell.borders = {};
      return;
    }
    if (preset === "all") {
      cell.borders.top = { ...thin };
      cell.borders.right = { ...thin };
      cell.borders.bottom = { ...thin };
      cell.borders.left = { ...thin };
      return;
    }
    if (preset === "remove-outside") {
      if (onTop) delete cell.borders.top;
      if (onRight) delete cell.borders.right;
      if (onBottom) delete cell.borders.bottom;
      if (onLeft) delete cell.borders.left;
      return;
    }
    if (preset === "inside") {
      if (visualRange.bottom < boundaryRange.bottom) cell.borders.bottom = { ...thin };
      if (visualRange.right < boundaryRange.right) cell.borders.right = { ...thin };
      return;
    }
    if ((preset === "thick-outside" || preset === "thick-bottom") && onBottom) {
      cell.borders.bottom = { ...thick };
    }
    if (preset === "thick-outside") {
      if (onTop) cell.borders.top = { ...thick };
      if (onRight) cell.borders.right = { ...thick };
      if (onLeft) cell.borders.left = { ...thick };
      return;
    }
    if ((preset === "outside" || preset === "top") && onTop) cell.borders.top = { ...thin };
    if ((preset === "outside" || preset === "right") && onRight) cell.borders.right = { ...thin };
    if ((preset === "outside" || preset === "bottom") && onBottom) cell.borders.bottom = { ...thin };
    if ((preset === "outside" || preset === "left") && onLeft) cell.borders.left = { ...thin };
  });
  touchedKeys.forEach((key) => {
    if (sheet.cells?.[key]) storeCellModelForSheet(sheet, key, sheet.cells[key]);
  });
}



function applyBorderPresetToSheetGroup(preset, range) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model.sheets[sheetIndex];
    if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートでは罫線を変更できません。")) return true;
    const targetRange = materializeSelectionRangeForSheet(sheet, range, sheetIndex);
    if (targetRange) items.push({ sheetIndex, sheet, range: targetRange });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => applyBorderPresetToSheetRange(item.sheet, item.range, preset));
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: range,
      selectionAfter: range,
      label: `${rangeToLabel(range)} ${sheetIndexes.length} シート罫線`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${rangeToLabel(range)}に罫線を適用しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`罫線を適用できませんでした：${error.message || error}`);
  }
  return true;
}



function applyBorderPreset(preset) {
  const range = activeSelectionRange();
  if (!range) {
    setStatus("セルを選択してください。");
    return;
  }
  if (!ensureActiveSheetFormatCellsAllowed("保護されたシートでは罫線を変更できません。")) return;
  if (applyBorderPresetToSheetGroup(preset, range)) return;

  const sheet = activeSheet();
  const before = captureCellsForHistory(sheet, range);
  applyBorderPresetToSheetRange(sheet, range, preset);
  const after = captureCellsForHistory(sheet, range);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    label: `${rangeToLabel(range)} 罫線`,
  });
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(`${rangeToLabel(range)} に罫線を適用しました。`);
}



function mergeSheetRange(sheet, range, centerContent) {
  removeIntersectingMerges(sheet, range);
  const master = ensureCellModel(sheet, range.top, range.left);
  if (centerContent) {
    master.css.textAlign = "center";
    master.css.alignItems = "center";
    master.css.verticalAlign = "middle";
  }

  forEachCellInRange(range, (row, col) => {
    if (row === range.top && col === range.left) return;
    delete sheet.cells[cellKey(row, col)];
  });

  sheet.merges.push({ top: range.top, left: range.left, bottom: range.bottom, right: range.right });
  internCellsInRange(sheet, range);
}



function mergeRangeContentCellCount(sheet, range) {
  if (!sheet || !range) return 0;
  const mergeMaps = buildMergeMaps(sheet.merges || []);
  const seen = new Set();
  let count = 0;
  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const masterKey = mergeMaps.coveredToMaster.get(key) || key;
    if (seen.has(masterKey)) return;
    seen.add(masterKey);
    if (cellHasMergeWarningContent(sheet.cells[masterKey])) count += 1;
  });
  return count;
}



function cellHasMergeWarningContent(cell) {
  if (!cell) return false;
  if (cell.formula || cell.dataTableFormula || cell.richText?.length) return true;
  if (cell.raw != null && String(cell.raw) !== "") return true;
  if (cell.cached != null && String(cell.cached) !== "") return true;
  if (cell.display != null && String(cell.display) !== "") return true;
  return false;
}



function mergeRangeNeedsContentWarning(sheet, range) {
  return mergeRangeContentCellCount(sheet, range) > 1;
}



function showMergeContentWarning(callback) {
  showMessageBox(MERGE_CONTENT_WARNING_MESSAGE, {
    title: "Cht WebSheet",
    tone: "warning",
    buttons: [
      { label: "OK", result: "ok", callback },
      { label: "キャンセル", result: "cancel" },
    ],
    cancelResult: "cancel",
  });
}



function mergeSelectionForSheetGroup(centerContent, range, options = {}) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model.sheets[sheetIndex];
    if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではセルを結合できません。")) return true;
    const targetRange = materializeSelectionRangeForSheet(sheet, range, sheetIndex);
    if (targetRange && !isSingleCellRange(targetRange)) items.push({ sheetIndex, sheet, range: targetRange });
  }
  if (!items.length) return true;
  if (!options.skipContentWarning && items.some((item) => mergeRangeNeedsContentWarning(item.sheet, item.range))) {
    showMergeContentWarning(() => mergeSelection(centerContent, { skipContentWarning: true }));
    return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => mergeSheetRange(item.sheet, item.range, centerContent));
    state.hf = buildFormulaEngine(state.model);
    state.selected = { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
    state.selectionRange = range;
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectionDrag = null;
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: range,
      selectionAfter: range,
      label: `${rangeToLabel(range)} ${sheetIndexes.length} シートセルの結合`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートで${rangeToLabel(range)}を結合しました${centerContent ? "（中央揃え）" : ""}。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`セルを結合できませんでした：${error.message || error}`);
  }
  return true;
}



function mergeSelection(centerContent, options = {}) {
  const range = activeSelectionRange();
  if (!range || isSingleCellRange(range)) {
    setStatus("結合するには 2 つ以上のセルを選択してください。");
    return;
  }
  if (!ensureActiveSheetFormatCellsAllowed("保護されたシートではセルを結合できません。")) return;
  const sheet = activeSheet();
  if (!options.skipContentWarning && mergeRangeNeedsContentWarning(sheet, range)) {
    showMergeContentWarning(() => mergeSelection(centerContent, { skipContentWarning: true }));
    return;
  }
  if (mergeSelectionForSheetGroup(centerContent, range, options)) return;

  const before = captureCellsForHistory(sheet, range);
  const mergesBefore = cloneMerges(sheet.merges);
  mergeSheetRange(sheet, range, centerContent);
  state.hf = buildFormulaEngine(state.model);
  state.selected = { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  const after = captureCellsForHistory(sheet, range);
  const mergesAfter = cloneMerges(sheet.merges);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    mergesBefore,
    mergesAfter,
    selectionBefore: range,
    selectionAfter: range,
    label: `${rangeToLabel(range)} セルの結合`,
  });
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(`${rangeToLabel(range)} を結合しました${centerContent ? "（中央揃え）" : ""}。`);
}



function toggleMergeCenterSelection() {
  const range = activeSelectionRange();
  if (!range) {
    setStatus("セルを選択してください。");
    return;
  }
  const sheet = activeSheet();
  const point = selectionToolbarPoint() || { row: range.top, col: range.left };
  if (findMergeForCell(sheet, point.row, point.col)) {
    unmergeSelection();
    return;
  }
  mergeSelection(true);
}



function unmergeSelectionForSheetGroup(range) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model.sheets[sheetIndex];
    if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではセルの結合を解除できません。")) return true;
    const targetRange = materializeSelectionRangeForSheet(sheet, range, sheetIndex);
    if (targetRange) items.push({ sheetIndex, sheet, range: targetRange });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  let removed = 0;
  try {
    items.forEach((item) => {
      const before = item.sheet.merges.length;
      removeIntersectingMerges(item.sheet, item.range);
      removed += before - item.sheet.merges.length;
    });
    if (removed) {
      state.hf = buildFormulaEngine(state.model);
      pushHistory({
        sheetIndex: state.activeSheetIndex,
        sheetsBefore,
        sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
        selectionBefore: range,
        selectionAfter: range,
        label: `${rangeToLabel(range)} ${sheetIndexes.length} シート結合の解除`,
      });
    }
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(removed ? `${sheetIndexes.length} 個のシートで ${removed} 個の結合範囲を解除しました。` : "現在の選択範囲には結合セルがありません。");
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`結合を解除できませんでした：${error.message || error}`);
  }
  return true;
}



function unmergeSelection() {
  const range = activeSelectionRange();
  if (!range) {
    setStatus("結合を解除する範囲を選択してください。");
    return;
  }
  if (!ensureActiveSheetFormatCellsAllowed("保護されたシートではセルの結合を解除できません。")) return;
  if (unmergeSelectionForSheetGroup(range)) return;

  const sheet = activeSheet();
  const beforeCells = captureCellsForHistory(sheet, range);
  const mergesBefore = cloneMerges(sheet.merges);
  const before = sheet.merges.length;
  removeIntersectingMerges(sheet, range);
  const removed = before - sheet.merges.length;
  if (removed) {
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      before: beforeCells,
      after: captureCellsForHistory(sheet, range),
      mergesBefore,
      mergesAfter: cloneMerges(sheet.merges),
      selectionBefore: range,
      selectionAfter: range,
      label: `${rangeToLabel(range)} 結合の解除`,
    });
  }
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(removed ? `${removed} 個の結合範囲を解除しました。` : "現在の選択範囲には結合セルがありません。");
}



function removeIntersectingMerges(sheet, range) {
  sheet.merges = sheet.merges.filter((merge) => !rangesIntersect(merge, range));
}



function rangesIntersect(a, b) {
  return a.left <= b.right && a.right >= b.left && a.top <= b.bottom && a.bottom >= b.top;
}



function rangeIntersects(a, b) {
  return rangesIntersect(a, b);
}



function forEachCellInRange(range, callback) {
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      callback(row, col);
    }
  }
}



function ensureCellModel(sheet, row, col) {
  const key = cellKey(row, col);
  if (!sheet.cells[key]) {
    const styleCell = sheetStyleRangeCell(sheet, row, col);
    if (styleCell) removeSheetStyleRangesAtCell(sheet, row, col);
    invalidateSheetUsedRangeCache(sheet);
    sheet.cells[key] = {
      row,
      col,
      raw: null,
      formula: null,
      cached: null,
      display: "",
      hyperlink: null,
      comment: null,
      note: null,
      noteVisible: false,
      validation: null,
      richText: null,
      css: styleCell?.css ? { ...styleCell.css } : {},
      hasFill: Boolean(styleCell?.hasFill),
      hasFontColor: Boolean(styleCell?.hasFontColor),
      borders: styleCell?.borders ? { ...styleCell.borders } : {},
      numFmt: "",
    };
  }
  sheet.cells[key].css = mutableCssObject(sheet.cells[key].css);
  sheet.cells[key].borders = mutableBordersObject(sheet.cells[key].borders);
  return sheet.cells[key];
}



function renderSheet(scrollOverride = null, options = {}) {
  if (!state.model) {
    removeLineConnectionPreviewOverlay();
    return;
  }
  const previousRenderDisplayCache = state.renderDisplayCache;
  const previousRenderFormulaSpillCache = state.renderFormulaSpillCache;
  state.renderDisplayCache = previousRenderDisplayCache instanceof Map ? previousRenderDisplayCache : new Map();
  state.deferredViewportRender = false;
  const sheet = activeSheet();
  state.renderFormulaSpillCache = formulaSpillRenderCacheForSheet(sheet) || createFormulaSpillRenderCache(sheet);
  try {
  ensureSheetCoversViewport(sheet);
  const mergeMaps = buildMergeMaps(sheet.merges);
  const sheetId = getSheetId(sheet.name);
  const visibleRange = visibleGridRange(sheet, scrollOverride);
  const freezeRows = frozenPaneRowCount(sheet);
  const freezeColumns = frozenPaneColumnCount(sheet);
  const renderedRows = renderIndexesWithFrozenBand(visibleRange.top, visibleRange.bottom, freezeRows, sheet.rowCount);
  const renderedColumns = renderIndexesWithFrozenBand(visibleRange.left, visibleRange.right, freezeColumns, sheet.colCount);
  const renderBands = renderBandsForVisibleRange(visibleRange, freezeRows, freezeColumns);
  const outlineControls = buildOutlineControlMaps(sheet);
  const html = [];

  html.push(
    `<div class="${sheetGridClassName(sheet)}" data-sheet="${escapeAttr(sheet.name)}" style="${escapeAttr(gridTrackStyle(sheet))}">`,
  );
  html.push(renderSheetBackgroundLayer(sheet));
  html.push(CanvasRenderer.renderLayer(sheet, visibleRange));
  html.push(renderFrozenPaneBackgroundLayer(sheet, visibleRange, renderedRows, renderedColumns));
  html.push(renderCornerHeader(sheet, outlineControls));
  renderedColumns.forEach((col) => {
    const columnModel = sheet.columns[col - 1] || {};
    const outlineLevel = Math.max(0, Number(columnModel.outlineLevel || 0));
    const outlineClass = outlineLevel ? " has-outline-level" : "";
    const outlineMark = outlineLevel ? `<span class="outline-level-mark" title="アウトライン レベル ${outlineLevel}">${outlineLevel}</span>` : "";
    const controls = outlineControls.columns.get(col) || "";
    const frozenClass = col <= freezeColumns ? " is-frozen-pane is-frozen-column" : "";
    const frozenStyle = frozenColumnStickyStyle(sheet, col, freezeColumns);
    html.push(
      `<div class="grid-header column-header${outlineClass}${frozenClass}" data-col="${col}" data-outline-level="${outlineLevel}" style="grid-column:${visualGridColumnStart(sheet, col)};grid-row:1;${frozenStyle}">${outlineMark}${controls}<span class="header-label">${columnName(col)}</span><span class="column-resize-handle" data-col="${col}" title="ドラッグで列幅を調整、ダブルクリックで自動調整"></span></div>`,
    );
  });

  const renderedCells = new Set();
  const appendCell = (row, col) => {
    const cellHtml = renderCellHtml(sheet, mergeMaps, sheetId, row, col, visibleRange, renderBands);
    if (cellHtml) {
      renderedCells.add(cellKey(row, col));
      html.push(cellHtml);
    }
  };

  renderedRows.forEach((row) => {
    const rowModel = sheet.rows[row - 1] || {};
    const outlineLevel = Math.max(0, Number(rowModel.outlineLevel || 0));
    const outlineClass = outlineLevel ? " has-outline-level" : "";
    const outlineMark = outlineLevel ? `<span class="outline-level-mark" title="アウトライン レベル ${outlineLevel}">${outlineLevel}</span>` : "";
    const controls = outlineControls.rows.get(row) || "";
    const frozenClass = row <= freezeRows ? " is-frozen-pane is-frozen-row" : "";
    const frozenStyle = frozenRowStickyStyle(sheet, row, freezeRows);
    html.push(
      `<div class="grid-header row-header${outlineClass}${frozenClass}" data-row="${row}" data-outline-level="${outlineLevel}" style="grid-column:1;grid-row:${row + 1};${rowModel.hidden ? "display:none;" : ""}${frozenStyle}">${outlineMark}${controls}<span class="header-label">${row}</span><span class="row-resize-handle" data-row="${row}" title="ドラッグで行の高さを調整、ダブルクリックで自動調整"></span></div>`,
    );
    renderedColumns.forEach((col) => {
      appendCell(row, col);
    });
  });

  sheet.merges.forEach((merge) => {
    if (!rangeIntersectsRenderBands(merge, renderBands)) return;
    const key = cellKey(merge.top, merge.left);
    if (!renderedCells.has(key)) appendCell(merge.top, merge.left);
  });

  html.push(renderFrozenPaneBorderLayer(sheet, mergeMaps, visibleRange));
  html.push(renderViewPaneGuides(sheet, visibleRange));
  html.push(renderFormulaAuditOverlay(sheet, visibleRange));
  html.push(renderImageLayer(sheet, visibleRange));
  html.push(renderInkLayer(sheet, visibleRange));
  html.push(renderImageSelectionFrames(sheet, visibleRange));
  html.push(renderSelectionFrame(sheet, visibleRange));
  html.push(renderFormulaReferenceFrame(sheet, visibleRange));
  html.push(renderCopyFrame(sheet, visibleRange));
  html.push(renderAutoFillOptions(sheet, visibleRange));
  html.push(renderPersistentNotePopups(sheet, visibleRange));
  html.push(renderAnnotationEditor(sheet, visibleRange));
  html.push("</div>");
  const markup = html.join("");
  let reconciled = false;
  if (options.viewportOnly) {
    const template = document.createElement("template");
    template.innerHTML = markup;
    reconciled = reconcileViewportGrid($sheetHost[0]?.firstElementChild, template.content.firstElementChild);
  }
  if (!reconciled) $sheetHost.html(markup);
  if (!state.imageTransform || state.imageTransform.mode !== "line-point") {
    removeLineConnectionPreviewOverlay();
  }
  CanvasRenderer.paint(sheet, mergeMaps, sheetId, visibleRange);
  updateHorizontalScrollbar();
  updateValidationInputMessagePopup();
  if (!options.viewportOnly && !options.deferSelectionUiSync) {
    updateSelectionStatusBar();
    syncDataRibbonControls();
    syncFormulaWatchWindow();
  }
  } finally {
    state.renderDisplayCache = previousRenderDisplayCache;
    state.renderFormulaSpillCache = previousRenderFormulaSpillCache;
  }
}



function renderIndexesWithFrozenBand(start, end, frozenCount, maxIndex) {
  const indexes = new Set();
  const safeMax = Math.max(0, Math.round(Number(maxIndex) || 0));
  const frozenEnd = Math.min(safeMax, Math.max(0, Math.round(Number(frozenCount) || 0)));
  for (let index = 1; index <= frozenEnd; index += 1) indexes.add(index);
  const safeStart = Math.max(1, Math.round(Number(start) || 1));
  const safeEnd = Math.min(safeMax, Math.max(safeStart, Math.round(Number(end) || safeStart)));
  for (let index = safeStart; index <= safeEnd; index += 1) indexes.add(index);
  return [...indexes].sort((left, right) => left - right);
}



function renderBandsForVisibleRange(visibleRange, freezeRows, freezeColumns) {
  const rowBands = [{ start: visibleRange.top, end: visibleRange.bottom }];
  const colBands = [{ start: visibleRange.left, end: visibleRange.right }];
  if (freezeRows > 0) rowBands.push({ start: 1, end: freezeRows });
  if (freezeColumns > 0) colBands.push({ start: 1, end: freezeColumns });
  return { rowBands, colBands };
}



function rangeIntersectsRenderBands(range, bands) {
  if (!range || !bands) return false;
  const rowHit = bands.rowBands.some((band) => rangesOverlap(range.top, range.bottom, band.start, band.end));
  const colHit = bands.colBands.some((band) => rangesOverlap(range.left, range.right, band.start, band.end));
  return rowHit && colHit;
}



function renderSheetBackgroundLayer(sheet) {
  const background = normalizePageLayoutBackground(ensurePageLayoutSettings(sheet).background);
  if (!background?.src) return "";
  const left = sheetHeadingsVisible(sheet) ? rowHeaderWidthForSheet(sheet) : 0;
  const top = sheetHeadingsVisible(sheet) ? COLUMN_HEADER_HEIGHT : 0;
  const width = Math.max(0, coordinateToPixels(sheet, "col", sheet.colCount) - left);
  const height = Math.max(0, coordinateToPixels(sheet, "row", sheet.rowCount) - top);
  const size = background.mode === "stretch"
    ? "100% 100%"
    : background.width && background.height
      ? `${Math.max(1, background.width)}px ${Math.max(1, background.height)}px`
      : "auto";
  const repeat = background.mode === "stretch" ? "no-repeat" : "repeat";
  const style = [
    `left:${roundCssPx(left)}px`,
    `top:${roundCssPx(top)}px`,
    `width:${roundCssPx(width)}px`,
    `height:${roundCssPx(height)}px`,
    `background-image:url("${cssUrlForStyle(background.src)}")`,
    `background-size:${size}`,
    `background-repeat:${repeat}`,
    "background-position:0 0",
  ].join(";");
  return `<div class="sheet-background-layer" aria-hidden="true" data-background-name="${escapeAttr(background.name || "背景")}" style="${escapeAttr(style)}"></div>`;
}



function cssUrlForStyle(value) {
  return String(value || "")
    .replace(/[\r\n]/g, "")
    .replace(/\\/g, "\\\\")
    .replace(/"/g, "%22");
}



function rangesOverlap(startA, endA, startB, endB) {
  return Number(startA) <= Number(endB) && Number(endA) >= Number(startB);
}



function frozenPaneRowCount(sheet) {
  const pane = normalizePaneState(sheet?.freezePanes);
  return Math.min(Math.max(0, Number(sheet?.rowCount || 0)), Math.max(0, Number(pane?.rows || 0)));
}



function frozenPaneColumnCount(sheet) {
  const pane = normalizePaneState(sheet?.freezePanes);
  return Math.min(Math.max(0, Number(sheet?.colCount || 0)), Math.max(0, Number(pane?.columns || 0)));
}



function frozenRowStickyStyle(sheet, row, freezeRows = frozenPaneRowCount(sheet)) {
  if (row > freezeRows) return "";
  return `top:${roundCssPx(frozenRowStickyTop(sheet, row))}px;`;
}



function frozenColumnStickyStyle(sheet, col, freezeColumns = frozenPaneColumnCount(sheet)) {
  if (col > freezeColumns) return "";
  if (sheetColumnsRightToLeft(sheet)) return `right:${roundCssPx(frozenColumnStickyRight(sheet, col))}px;`;
  return `left:${roundCssPx(frozenColumnStickyLeft(sheet, col))}px;`;
}



function frozenRowStickyTop(sheet, row) {
  let top = sheetHeadingsVisible(sheet) ? COLUMN_HEADER_HEIGHT : 0;
  for (let current = 1; current < row; current += 1) {
    const rowModel = sheet.rows[current - 1] || {};
    top += rowModel.hidden ? 0 : rowModel.height || sheet.defaultRowHeight || 0;
  }
  return top;
}



function frozenColumnStickyLeft(sheet, col) {
  let left = sheetHeadingsVisible(sheet) ? rowHeaderWidthForSheet(sheet) : 0;
  for (let current = 1; current < col; current += 1) {
    const columnModel = sheet.columns[current - 1] || {};
    left += columnModel.hidden ? 0 : columnModel.width || sheet.defaultColumnWidth || 0;
  }
  return left;
}



function frozenColumnStickyRight(sheet, col) {
  let right = 0;
  for (let current = 1; current < col; current += 1) {
    const columnModel = sheet.columns[current - 1] || {};
    right += columnModel.hidden ? 0 : columnModel.width || sheet.defaultColumnWidth || 0;
  }
  return right;
}



function frozenCellStickyCss(sheet, row, col) {
  const freezeRows = frozenPaneRowCount(sheet);
  const freezeColumns = frozenPaneColumnCount(sheet);
  const css = {};
  if (row <= freezeRows) css.top = `${roundCssPx(frozenRowStickyTop(sheet, row))}px`;
  if (col <= freezeColumns) {
    if (sheetColumnsRightToLeft(sheet)) css.right = `${roundCssPx(frozenColumnStickyRight(sheet, col))}px`;
    else css.left = `${roundCssPx(frozenColumnStickyLeft(sheet, col))}px`;
  }
  return css;
}



function frozenCellClassNames(sheet, row, col) {
  const freezeRows = frozenPaneRowCount(sheet);
  const freezeColumns = frozenPaneColumnCount(sheet);
  const classes = [];
  if (row <= freezeRows || col <= freezeColumns) classes.push("is-frozen-pane");
  if (row <= freezeRows) classes.push("is-frozen-row");
  if (col <= freezeColumns) classes.push("is-frozen-column");
  if (row <= freezeRows && col <= freezeColumns) classes.push("is-frozen-corner");
  return classes;
}



function isFrozenPaneCell(sheet, row, col) {
  return row <= frozenPaneRowCount(sheet) || col <= frozenPaneColumnCount(sheet);
}



function renderFrozenPaneBackgroundLayer(sheet, visibleRange, renderedRows, renderedColumns) {
  const freezeRows = frozenPaneRowCount(sheet);
  const freezeColumns = frozenPaneColumnCount(sheet);
  if (!freezeRows && !freezeColumns) return "";
  const parts = [];
  const rowHeaderWidth = rowHeaderWidthForSheet(sheet);
  const headerHeight = sheetHeadingsVisible(sheet) ? COLUMN_HEADER_HEIGHT : 0;
  const dataLeft = sheetHeadingsVisible(sheet) ? rowHeaderWidth : 0;
  const dataTop = headerHeight;

  if (freezeRows > 0) {
    const height = Math.max(0, coordinateToPixels(sheet, "row", freezeRows) - dataTop);
    const lines = [];
    (renderedColumns || []).forEach((col) => {
      if (col < 1 || col > sheet.colCount) return;
      const column = sheet.columns[col - 1] || {};
      if (column.hidden) return;
      const x = coordinateToPixels(sheet, "col", col) - dataLeft;
      lines.push(`<span class="sheet-frozen-bg-line is-vertical" style="${escapeAttr(`left:${roundCssPx(x)}px;top:0;height:${roundCssPx(height)}px`)}"></span>`);
    });
    for (let row = 1; row <= freezeRows; row += 1) {
      const rowModel = sheet.rows[row - 1] || {};
      if (rowModel.hidden) continue;
      const y = coordinateToPixels(sheet, "row", row) - dataTop;
      lines.push(`<span class="sheet-frozen-bg-line is-horizontal" style="${escapeAttr(`top:${roundCssPx(y)}px;left:0;width:100%`)}"></span>`);
    }
    const directionStyle = sheetColumnsRightToLeft(sheet) ? "" : `left:${roundCssPx(dataLeft)}px`;
    parts.push(`<div class="sheet-frozen-pane-background is-frozen-rows" aria-hidden="true" style="${escapeAttr(`grid-column:2 / span ${sheet.colCount};grid-row:2 / span ${freezeRows};top:${roundCssPx(dataTop)}px;${directionStyle}`)}">${lines.join("")}</div>`);
  }

  if (freezeColumns > 0) {
    const width = Math.max(0, coordinateToPixels(sheet, "col", freezeColumns) - dataLeft);
    const lines = [];
    (renderedRows || []).forEach((row) => {
      if (row < 1 || row > sheet.rowCount) return;
      const rowModel = sheet.rows[row - 1] || {};
      if (rowModel.hidden) return;
      const y = coordinateToPixels(sheet, "row", row) - dataTop;
      lines.push(`<span class="sheet-frozen-bg-line is-horizontal" style="${escapeAttr(`top:${roundCssPx(y)}px;left:0;width:${roundCssPx(width)}px`)}"></span>`);
    });
    for (let col = 1; col <= freezeColumns; col += 1) {
      const column = sheet.columns[col - 1] || {};
      if (column.hidden) continue;
      const x = coordinateToPixels(sheet, "col", col) - dataLeft;
      lines.push(`<span class="sheet-frozen-bg-line is-vertical" style="${escapeAttr(`left:${roundCssPx(x)}px;top:0;height:100%`)}"></span>`);
    }
    const horizontalSticky = sheetColumnsRightToLeft(sheet)
      ? `right:${roundCssPx(frozenColumnStickyRight(sheet, 1))}px`
      : `left:${roundCssPx(dataLeft)}px`;
    parts.push(`<div class="sheet-frozen-pane-background is-frozen-columns" aria-hidden="true" style="${escapeAttr(`grid-column:2 / span ${freezeColumns};grid-row:2 / span ${sheet.rowCount};${horizontalSticky};width:${roundCssPx(width)}px`)}">${lines.join("")}</div>`);
  }
  return parts.join("");
}



function renderFrozenPaneBorderLayer(sheet, mergeMaps, visibleRange) {
  const freezeRows = frozenPaneRowCount(sheet);
  const freezeColumns = frozenPaneColumnCount(sheet);
  if (!freezeRows && !freezeColumns) return "";
  const dataLeft = sheetHeadingsVisible(sheet) ? rowHeaderWidthForSheet(sheet) : 0;
  const dataTop = sheetHeadingsVisible(sheet) ? COLUMN_HEADER_HEIGHT : 0;
  const parts = [];
  const renderLines = (range) => borderLinesForSheet(sheet, mergeMaps, range)
    .map((line) => borderLineHtml(line, "sheet-border-line sheet-frozen-border-line", dataLeft, dataTop))
    .join("");

  if (freezeRows > 0) {
    const range = {
      top: 1,
      bottom: freezeRows,
      left: Math.max(1, Number(visibleRange.left) || 1),
      right: Math.min(Number(sheet.colCount || visibleRange.right || 1), Math.max(1, Number(visibleRange.right) || 1)),
    };
    const lines = renderLines(range);
    if (lines) {
      const directionStyle = sheetColumnsRightToLeft(sheet) ? "" : `left:${roundCssPx(dataLeft)}px`;
      parts.push(`<div class="sheet-frozen-border-layer is-frozen-rows" aria-hidden="true" style="${escapeAttr(`grid-column:2 / span ${sheet.colCount};grid-row:2 / span ${freezeRows};top:${roundCssPx(dataTop)}px;${directionStyle}`)}">${lines}</div>`);
    }
  }

  if (freezeColumns > 0) {
    const range = {
      top: Math.max(1, Number(visibleRange.top) || 1),
      bottom: Math.min(Number(sheet.rowCount || visibleRange.bottom || 1), Math.max(1, Number(visibleRange.bottom) || 1)),
      left: 1,
      right: freezeColumns,
    };
    const lines = renderLines(range);
    if (lines) {
      const width = Math.max(0, coordinateToPixels(sheet, "col", freezeColumns) - dataLeft);
      const horizontalSticky = sheetColumnsRightToLeft(sheet)
        ? `right:${roundCssPx(frozenColumnStickyRight(sheet, 1))}px`
        : `left:${roundCssPx(dataLeft)}px`;
      parts.push(`<div class="sheet-frozen-border-layer is-frozen-columns" aria-hidden="true" style="${escapeAttr(`grid-column:2 / span ${freezeColumns};grid-row:2 / span ${sheet.rowCount};${horizontalSticky};width:${roundCssPx(width)}px`)}">${lines}</div>`);
    }
  }

  return parts.join("");
}



function renderViewPaneGuides(sheet, visibleRange) {
  const guides = [];
  const gridRight = coordinateToPixels(sheet, "col", sheet.colCount);
  const gridBottom = coordinateToPixels(sheet, "row", sheet.rowCount);
  const visibleLeft = coordinateToPixels(sheet, "col", Math.max(0, visibleRange.left - 1));
  const visibleRight = coordinateToPixels(sheet, "col", visibleRange.right);
  const visibleTop = coordinateToPixels(sheet, "row", Math.max(0, visibleRange.top - 1));
  const visibleBottom = coordinateToPixels(sheet, "row", visibleRange.bottom);
  const topOffset = sheetHeadingsVisible(sheet) ? COLUMN_HEADER_HEIGHT : 0;
  const leftOffset = sheetHeadingsVisible(sheet) ? rowHeaderWidthForSheet(sheet) : 0;
  const lineTop = Math.max(topOffset, visibleTop);
  const lineLeft = Math.max(leftOffset, visibleLeft);
  const lineHeight = Math.max(0, visibleBottom - lineTop);
  const lineWidth = Math.max(0, visibleRight - lineLeft);
  const addVertical = (x, className, label, attributes = "") => {
    if (x < visibleLeft - 4 || x > visibleRight + 4) return;
    guides.push(`<div class="sheet-view-guide is-vertical ${className}" aria-hidden="true" title="${escapeAttr(label)}"${attributes} style="left:${roundCssPx(x)}px;top:${roundCssPx(lineTop)}px;height:${roundCssPx(lineHeight)}px"></div>`);
  };
  const addHorizontal = (y, className, label, attributes = "") => {
    if (y < visibleTop - 4 || y > visibleBottom + 4) return;
    guides.push(`<div class="sheet-view-guide is-horizontal ${className}" aria-hidden="true" title="${escapeAttr(label)}"${attributes} style="top:${roundCssPx(y)}px;left:${roundCssPx(lineLeft)}px;width:${roundCssPx(lineWidth)}px"></div>`);
  };

  const freeze = normalizePaneState(sheet.freezePanes);
  if (freeze?.columns > 0) addVertical(coordinateToPixels(sheet, "col", freeze.columns), "is-freeze-pane", "ウィンドウ枠の固定");
  if (freeze?.rows > 0) addHorizontal(coordinateToPixels(sheet, "row", freeze.rows), "is-freeze-pane", "ウィンドウ枠の固定");

  const split = splitPaneCounts(sheet);
  if (split.columns > 0) addVertical(coordinateToPixels(sheet, "col", split.columns), "is-split-pane", "分割", ' data-split-pane-axis="columns"');
  if (split.rows > 0) addHorizontal(coordinateToPixels(sheet, "row", split.rows), "is-split-pane", "分割", ' data-split-pane-axis="rows"');

  const geometry = { gridRight, gridBottom, visibleLeft, visibleRight, visibleTop, visibleBottom, leftOffset, topOffset };
  if ((sheet.viewMode || "normal") === "page-layout") {
    renderPageLayoutGuides(sheet, visibleRange, guides, geometry);
    if (sheet.showRuler === true) {
      guides.push(renderSheetRulers(sheet, geometry));
    }
  }
  if ((sheet.viewMode || "normal") === "page-break") {
    renderPageBreakGuides(sheet, visibleRange, guides, addVertical, addHorizontal, geometry);
  }
  return guides.join("");
}



function renderFormulaAuditOverlay(sheet, visibleRange) {
  const audits = normalizeFormulaAudits(sheet?.formulaAudits);
  if (!audits.arrows.length) return "";
  const visibleLeft = coordinateToPixels(sheet, "col", Math.max(0, visibleRange.left - 1));
  const visibleRight = coordinateToPixels(sheet, "col", visibleRange.right);
  const visibleTop = coordinateToPixels(sheet, "row", Math.max(0, visibleRange.top - 1));
  const visibleBottom = coordinateToPixels(sheet, "row", visibleRange.bottom);
  const visibleBox = {
    left: visibleLeft,
    top: visibleTop,
    right: visibleRight,
    bottom: visibleBottom,
    width: Math.max(1, visibleRight - visibleLeft),
    height: Math.max(1, visibleBottom - visibleTop),
  };
  const lines = audits.arrows
    .map((arrow, index) => formulaAuditArrowSvg(sheet, arrow, visibleBox, index))
    .filter(Boolean)
    .join("");
  if (!lines) return "";
  return `
    <svg class="formula-audit-overlay" aria-hidden="true" style="left:${roundCssPx(visibleBox.left)}px;top:${roundCssPx(visibleBox.top)}px;width:${roundCssPx(visibleBox.width)}px;height:${roundCssPx(visibleBox.height)}px" viewBox="${visibleBox.left} ${visibleBox.top} ${visibleBox.width} ${visibleBox.height}">
      <defs>
        <marker id="formulaAuditArrowBlue" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto" markerUnits="strokeWidth"><path d="M0,0 L8,4 L0,8 Z" fill="#2563eb"></path></marker>
        <marker id="formulaAuditArrowRed" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto" markerUnits="strokeWidth"><path d="M0,0 L8,4 L0,8 Z" fill="#dc2626"></path></marker>
      </defs>
      ${lines}
    </svg>`;
}



function formulaAuditArrowSvg(sheet, arrow, visibleBox, index) {
  if (!arrow?.fromRange || !arrow?.toRange) return "";
  const fromRect = rangePixelRect(sheet, arrow.fromRange);
  const toRect = rangePixelRect(sheet, arrow.toRange);
  if (!pixelRectsIntersect(fromRect, visibleBox) && !pixelRectsIntersect(toRect, visibleBox) && !formulaAuditLineIntersectsBox(fromRect, toRect, visibleBox)) return "";
  const start = formulaAuditArrowPoint(fromRect, toRect, "start");
  const end = formulaAuditArrowPoint(toRect, fromRect, "end");
  const color = arrow.kind === "dependents" ? "#dc2626" : "#2563eb";
  const marker = arrow.kind === "dependents" ? "formulaAuditArrowRed" : "formulaAuditArrowBlue";
  const dash = arrow.kind === "dependents" ? ' stroke-dasharray="5 3"' : "";
  return `
    <g class="formula-audit-arrow is-${escapeAttr(arrow.kind)}" data-formula-audit-arrow="${index}">
      <line x1="${roundCssPx(start.x)}" y1="${roundCssPx(start.y)}" x2="${roundCssPx(end.x)}" y2="${roundCssPx(end.y)}" stroke="${color}"${dash} marker-end="url(#${marker})"></line>
      <circle cx="${roundCssPx(start.x)}" cy="${roundCssPx(start.y)}" r="3" fill="#ffffff" stroke="${color}"></circle>
    </g>`;
}



function formulaAuditArrowPoint(rect, oppositeRect, role) {
  const center = { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
  const target = { x: oppositeRect.left + oppositeRect.width / 2, y: oppositeRect.top + oppositeRect.height / 2 };
  const dx = target.x - center.x;
  const dy = target.y - center.y;
  if (Math.abs(dx) > Math.abs(dy)) {
    return { x: dx > 0 ? rect.right : rect.left, y: center.y };
  }
  if (Math.abs(dy) > 0) {
    return { x: center.x, y: dy > 0 ? rect.bottom : rect.top };
  }
  return role === "end" ? { x: rect.left, y: center.y } : { x: rect.right, y: center.y };
}



function formulaAuditLineIntersectsBox(fromRect, toRect, box) {
  const x1 = fromRect.left + fromRect.width / 2;
  const y1 = fromRect.top + fromRect.height / 2;
  const x2 = toRect.left + toRect.width / 2;
  const y2 = toRect.top + toRect.height / 2;
  const minX = Math.min(x1, x2);
  const maxX = Math.max(x1, x2);
  const minY = Math.min(y1, y2);
  const maxY = Math.max(y1, y2);
  return minX <= box.right && maxX >= box.left && minY <= box.bottom && maxY >= box.top;
}



function renderSheetRulers(sheet, geometry) {
  const leftOffset = sheetHeadingsVisible(sheet) ? rowHeaderWidthForSheet(sheet) : 0;
  const topOffset = sheetHeadingsVisible(sheet) ? COLUMN_HEADER_HEIGHT : 0;
  const horizontalStart = Math.max(leftOffset, geometry.visibleLeft);
  const horizontalEnd = Math.max(horizontalStart, geometry.visibleRight);
  const verticalStart = Math.max(topOffset, geometry.visibleTop);
  const verticalEnd = Math.max(verticalStart, geometry.visibleBottom);
  return `
    <div class="sheet-ruler sheet-ruler-horizontal" aria-hidden="true" style="grid-column:${sheetHeadingsVisible(sheet) ? "2 / -1" : "1 / -1"};grid-row:1;">
      ${renderSheetRulerTicks("horizontal", sheet, horizontalStart, horizontalEnd, leftOffset)}
    </div>
    <div class="sheet-ruler sheet-ruler-vertical" aria-hidden="true" style="grid-column:1;grid-row:${sheetHeadingsVisible(sheet) ? "2 / -1" : "1 / -1"};">
      ${renderSheetRulerTicks("vertical", sheet, verticalStart, verticalEnd, topOffset)}
    </div>`;
}



function renderSheetRulerTicks(axis, sheet, startPx, endPx, originPx) {
  const unit = sheetRulerUnitPixels(sheet);
  const quarter = Math.max(4, unit / 4);
  const minTick = Math.floor((startPx - originPx) / quarter) - 2;
  const maxTick = Math.ceil((endPx - originPx) / quarter) + 2;
  const ticks = [];
  for (let tick = minTick; tick <= maxTick; tick += 1) {
    const absolute = originPx + tick * quarter;
    const local = absolute - originPx;
    const major = tick % 4 === 0;
    const half = tick % 2 === 0;
    const label = major ? Math.max(0, Math.round(tick / 4)) : "";
    ticks.push(
      `<span class="sheet-ruler-tick${major ? " is-major" : half ? " is-half" : ""}" style="${axis === "horizontal" ? "left" : "top"}:${roundCssPx(local)}px">${label !== "" ? `<span>${escapeHtml(label)}</span>` : ""}</span>`,
    );
  }
  return ticks.join("");
}



function sheetRulerUnitPixels(sheet) {
  const layout = normalizedPageLayoutSettings(sheet?.pageLayout, sheet);
  const pageScale = Math.max(0.1, Number(layout.scale || 100) / 100);
  const displayScale = Math.max(0.1, Number(sheet?.displayScale || 1));
  return Math.max(24, (96 * displayScale) / pageScale);
}



function renderPageLayoutGuides(sheet, visibleRange, guides, geometry) {
  const layout = normalizedPageLayoutSettings(sheet.pageLayout, sheet);
  const metrics = pageLayoutPaperMetrics(sheet, layout);
  const showWhiteSpace = sheet.showWhiteSpace !== false;
  const maxPageCol = Math.max(0, Math.ceil(Math.max(1, geometry.gridRight - geometry.leftOffset) / metrics.paperWidth) - 1);
  const maxPageRow = Math.max(0, Math.ceil(Math.max(1, geometry.gridBottom - geometry.topOffset) / metrics.paperHeight) - 1);
  const firstCol = Math.max(0, Math.floor(Math.max(0, geometry.visibleLeft - geometry.leftOffset) / metrics.paperWidth) - 1);
  const firstRow = Math.max(0, Math.floor(Math.max(0, geometry.visibleTop - geometry.topOffset) / metrics.paperHeight) - 1);
  const lastCol = Math.min(maxPageCol, Math.ceil(Math.max(0, geometry.visibleRight - geometry.leftOffset) / metrics.paperWidth) + 1, firstCol + 8);
  const lastRow = Math.min(maxPageRow, Math.ceil(Math.max(0, geometry.visibleBottom - geometry.topOffset) / metrics.paperHeight) + 1, firstRow + 8);
  const viewBox = {
    left: geometry.visibleLeft,
    top: geometry.visibleTop,
    right: geometry.visibleRight,
    bottom: geometry.visibleBottom,
  };

  for (let pageRow = firstRow; pageRow <= lastRow; pageRow += 1) {
    for (let pageCol = firstCol; pageCol <= lastCol; pageCol += 1) {
      const left = geometry.leftOffset + pageCol * metrics.paperWidth;
      const top = geometry.topOffset + pageRow * metrics.paperHeight;
      const pageRect = { left, top, right: left + metrics.paperWidth, bottom: top + metrics.paperHeight };
      if (!pixelRectsIntersect(pageRect, viewBox)) continue;
      appendPageLayoutRect(guides, pageRect, `sheet-page-layout-page${showWhiteSpace ? "" : " is-white-space-hidden"}`, `ページ ${pageRow + 1}-${pageCol + 1}`, viewBox);
      appendPageLayoutWhiteSpaceEdges(guides, pageRect, viewBox);
      if (!showWhiteSpace) continue;
      const marginRect = {
        left: left + metrics.marginLeft,
        top: top + metrics.marginTop,
        right: left + metrics.paperWidth - metrics.marginRight,
        bottom: top + metrics.paperHeight - metrics.marginBottom,
      };
      appendPageLayoutRect(guides, marginRect, "sheet-page-layout-margin", "余白", viewBox);
      if (metrics.headerHeight > 0) {
        appendPageLayoutRect(guides, {
          left: marginRect.left,
          top: Math.max(top, top + metrics.marginTop - metrics.headerHeight),
          right: marginRect.right,
          bottom: Math.max(top + 1, top + metrics.marginTop),
        }, "sheet-page-layout-header", "ヘッダー", viewBox, pageLayoutHeaderFooterHtml(sheet, "header"));
      }
      if (metrics.footerHeight > 0) {
        appendPageLayoutRect(guides, {
          left: marginRect.left,
          top: Math.min(top + metrics.paperHeight - 1, top + metrics.paperHeight - metrics.marginBottom),
          right: marginRect.right,
          bottom: Math.min(top + metrics.paperHeight, top + metrics.paperHeight - metrics.marginBottom + metrics.footerHeight),
        }, "sheet-page-layout-footer", "フッター", viewBox, pageLayoutHeaderFooterHtml(sheet, "footer"));
      }
    }
  }

  if (layout.printArea?.range && rangeIntersects(layout.printArea.range, visibleRange)) {
    appendPageLayoutRect(guides, rangePixelRect(sheet, layout.printArea.range), "sheet-page-layout-print-area", "印刷範囲", viewBox);
  }
  if (layout.printTitles?.rows) {
    const rows = { ...layout.printTitles.rows, left: visibleRange.left, right: visibleRange.right };
    if (rangeIntersects(rows, visibleRange)) appendPageLayoutRect(guides, rangePixelRect(sheet, rows), "sheet-page-layout-title-row", "印刷タイトル行", viewBox);
  }
  if (layout.printTitles?.cols) {
    const cols = { ...layout.printTitles.cols, top: visibleRange.top, bottom: visibleRange.bottom };
    if (rangeIntersects(cols, visibleRange)) appendPageLayoutRect(guides, rangePixelRect(sheet, cols), "sheet-page-layout-title-col", "印刷タイトル列", viewBox);
  }
  const background = normalizePageLayoutBackground(layout.background);
  if (background?.src) {
    guides.push(`<div class="sheet-page-layout-background-label" aria-hidden="true" style="left:${roundCssPx(geometry.visibleLeft + 12)}px;top:${roundCssPx(geometry.visibleTop + 12)}px">背景: ${escapeHtml(background.name || "画像")}</div>`);
  }
}



function appendPageLayoutWhiteSpaceEdges(guides, pageRect, viewBox) {
  const edge = 8;
  [
    { left: pageRect.left, top: pageRect.top, right: pageRect.right, bottom: pageRect.top + edge },
    { left: pageRect.left, top: pageRect.bottom - edge, right: pageRect.right, bottom: pageRect.bottom },
    { left: pageRect.left, top: pageRect.top, right: pageRect.left + edge, bottom: pageRect.bottom },
    { left: pageRect.right - edge, top: pageRect.top, right: pageRect.right, bottom: pageRect.bottom },
  ].forEach((rect) => appendPageLayoutRect(guides, rect, "sheet-page-layout-white-space-edge", "ページ余白の表示/非表示", viewBox));
}



function pageLayoutPaperMetrics(sheet, layout) {
  const paper = PAGE_LAYOUT_PAPER_SIZES.find((item) => item.value === layout.paperSize) || PAGE_LAYOUT_PAPER_SIZES[0];
  const landscape = layout.orientation === "landscape";
  const paperWidth = landscape ? paper.height : paper.width;
  const paperHeight = landscape ? paper.width : paper.height;
  const margins = layout.margins || PAGE_LAYOUT_MARGIN_PRESETS.normal;
  const scale = Math.max(0.1, Number(layout.scale || 100) / 100);
  const displayScale = Math.max(0.1, Number(sheet.displayScale || 1));
  const toPx = (inches) => Math.max(0, (Number(inches || 0) * 96 * displayScale) / scale);
  const paperWidthPx = Math.max(180, toPx(paperWidth));
  const paperHeightPx = Math.max(180, toPx(paperHeight));
  const marginLeft = Math.min(paperWidthPx / 2 - 8, toPx(margins.left));
  const marginRight = Math.min(paperWidthPx / 2 - 8, toPx(margins.right));
  const marginTop = Math.min(paperHeightPx / 2 - 8, toPx(margins.top));
  const marginBottom = Math.min(paperHeightPx / 2 - 8, toPx(margins.bottom));
  return {
    paperWidth: paperWidthPx,
    paperHeight: paperHeightPx,
    marginLeft,
    marginRight,
    marginTop,
    marginBottom,
    headerHeight: Math.min(Math.max(8, toPx(margins.header)), Math.max(8, marginTop)),
    footerHeight: Math.min(Math.max(8, toPx(margins.footer)), Math.max(8, marginBottom)),
  };
}



function appendPageLayoutRect(guides, rect, className, label, viewBox, contentHtml = "") {
  const clipped = clipPixelRect(rect, viewBox);
  if (!clipped || clipped.width <= 0 || clipped.height <= 0) return;
  guides.push(`<div class="${className}" aria-hidden="true" title="${escapeAttr(label)}" style="left:${roundCssPx(clipped.left)}px;top:${roundCssPx(clipped.top)}px;width:${roundCssPx(clipped.width)}px;height:${roundCssPx(clipped.height)}px">${contentHtml}</div>`);
}



function pageLayoutHeaderFooterHtml(sheet, kind) {
  const settings = normalizedHeaderFooterSettings(sheet?.headerFooter, sheet?.pageLayout);
  const section = kind === "footer" ? settings.oddFooter : settings.oddHeader;
  const fallback = kind === "footer" ? "フッターを追加" : "ヘッダーを追加";
  const values = ["left", "center", "right"].map((key) => section?.[key] || "");
  const empty = values.every((value) => !String(value || "").trim());
  const slots = values
    .map((value, index) => `<span class="sheet-page-layout-header-footer-slot${value ? "" : " is-empty"}" data-header-footer-slot="${index}">${escapeHtml(value || (index === 1 && empty ? fallback : ""))}</span>`)
    .join("");
  return `<div class="sheet-page-layout-header-footer-content${empty ? " is-empty" : ""}" aria-hidden="true">${slots}</div>`;
}



function handlePageLayoutHeaderFooterDblClick(event) {
  if (!state.model) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  openHeaderFooterDialog();
}



function handlePageLayoutWhiteSpaceDblClick(event) {
  const sheet = activeSheet();
  if (!state.model || (sheet?.viewMode || "normal") !== "page-layout") return;
  const rect = event.currentTarget.getBoundingClientRect();
  const edgeDistance = Math.min(
    Math.abs(event.clientX - rect.left),
    Math.abs(event.clientX - rect.right),
    Math.abs(event.clientY - rect.top),
    Math.abs(event.clientY - rect.bottom),
  );
  if (edgeDistance > 10) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  toggleActiveSheetWhiteSpace();
}



function clipPixelRect(rect, clip) {
  const left = Math.max(rect.left, clip.left);
  const top = Math.max(rect.top, clip.top);
  const right = Math.min(rect.right, clip.right);
  const bottom = Math.min(rect.bottom, clip.bottom);
  if (right <= left || bottom <= top) return null;
  return { left, top, right, bottom, width: right - left, height: bottom - top };
}



function pixelRectsIntersect(a, b) {
  return a.left < b.right && a.right > b.left && a.top < b.bottom && a.bottom > b.top;
}



function renderPageBreakGuides(sheet, visibleRange, guides, addVertical, addHorizontal, geometry) {
  const layout = normalizedPageLayoutSettings(sheet.pageLayout, sheet);
  const manualRows = new Set(layout.pageBreaks.rows || []);
  const manualCols = new Set(layout.pageBreaks.cols || []);
  const plan = pageBreakPreviewPlan(sheet, layout);
  const printRect = plan.range ? rangePixelRect(sheet, plan.range) : null;
  appendPageBreakPreviewPages(sheet, visibleRange, guides, geometry, plan, layout);
  manualCols.forEach((col) => {
    if (!plan.cols.breakIndexes.has(Number(col))) return;
    const position = coordinateToPixels(sheet, "col", col - 1);
    if (pageBreakGuideCoincidesWithPrintAreaEdge("col", position, printRect)) return;
    addVertical(
      position,
      "is-page-break is-manual",
      "改ページ",
      pageBreakGuideAttributes("col", col, true),
    );
  });
  manualRows.forEach((row) => {
    if (!plan.rows.breakIndexes.has(Number(row))) return;
    const position = coordinateToPixels(sheet, "row", row - 1);
    if (pageBreakGuideCoincidesWithPrintAreaEdge("row", position, printRect)) return;
    addHorizontal(
      position,
      "is-page-break is-manual",
      "改ページ",
      pageBreakGuideAttributes("row", row, true),
    );
  });

  plan.cols.breaks.forEach((breakInfo) => {
    if (breakInfo.manual) return;
    const position = coordinateToPixels(sheet, "col", breakInfo.index - 1);
    if (pageBreakGuideCoincidesWithPrintAreaEdge("col", position, printRect)) return;
    addVertical(
      position,
      "is-page-break is-automatic",
      "自動改ページ",
      pageBreakGuideAttributes("col", breakInfo.index, false),
    );
  });
  plan.rows.breaks.forEach((breakInfo) => {
    if (breakInfo.manual) return;
    const position = coordinateToPixels(sheet, "row", breakInfo.index - 1);
    if (pageBreakGuideCoincidesWithPrintAreaEdge("row", position, printRect)) return;
    addHorizontal(
      position,
      "is-page-break is-automatic",
      "自動改ページ",
      pageBreakGuideAttributes("row", breakInfo.index, false),
    );
  });

  if (plan.range) {
    const rect = printRect;
    appendPageBreakOutsidePrintArea(guides, rect, geometry);
    if (rangeIntersects(plan.range, visibleRange)) {
      guides.push(`<div class="sheet-view-print-area" aria-hidden="true" data-print-area-range="${escapeAttr(rangeToLabel(plan.range))}" style="left:${roundCssPx(rect.left)}px;top:${roundCssPx(rect.top)}px;width:${roundCssPx(rect.width)}px;height:${roundCssPx(rect.height)}px"></div>`);
      guides.push(printAreaDragHandlesHtml(rect));
    }
  }
}



function pageBreakGuideCoincidesWithPrintAreaEdge(axis, position, printRect) {
  if (!printRect) return false;
  const value = Number(position);
  if (!Number.isFinite(value)) return false;
  const tolerance = 2.5;
  if (axis === "col") return Math.abs(value - printRect.left) <= tolerance || Math.abs(value - printRect.right) <= tolerance;
  return Math.abs(value - printRect.top) <= tolerance || Math.abs(value - printRect.bottom) <= tolerance;
}



function printAreaDragHandlesHtml(rect) {
  const handle = (edge, style) =>
    `<span class="sheet-view-print-area-handle is-${edge}" data-print-area-edge="${edge}" title="ドラッグして印刷範囲を変更" style="${escapeAttr(style)}"></span>`;
  return ["top", "right", "bottom", "left"]
    .map((edge) => {
      if (edge === "top") return handle(edge, `left:${roundCssPx(rect.left - 5)}px;top:${roundCssPx(rect.top - 6)}px;width:${roundCssPx(rect.width + 10)}px;height:10px`);
      if (edge === "bottom") return handle(edge, `left:${roundCssPx(rect.left - 5)}px;top:${roundCssPx(rect.bottom - 4)}px;width:${roundCssPx(rect.width + 10)}px;height:10px`);
      if (edge === "left") return handle(edge, `left:${roundCssPx(rect.left - 6)}px;top:${roundCssPx(rect.top - 5)}px;width:10px;height:${roundCssPx(rect.height + 10)}px`);
      return handle(edge, `left:${roundCssPx(rect.right - 4)}px;top:${roundCssPx(rect.top - 5)}px;width:10px;height:${roundCssPx(rect.height + 10)}px`);
    })
    .join("");
}



function appendPageBreakPreviewPages(sheet, visibleRange, guides, geometry, plan, layout) {
  const clip = { left: geometry.visibleLeft, top: geometry.visibleTop, right: geometry.visibleRight, bottom: geometry.visibleBottom };
  const rowSegments = plan?.rows?.segments || [];
  const colSegments = plan?.cols?.segments || [];
  const totalRows = Math.max(1, rowSegments.length);
  const totalCols = Math.max(1, colSegments.length);

  rowSegments.forEach((rowSegment, rowIndex) => {
    colSegments.forEach((colSegment, colIndex) => {
      const rect = {
        left: colSegment.startPx,
        top: rowSegment.startPx,
        right: colSegment.endPx,
        bottom: rowSegment.endPx,
      };
      const clipped = clipPixelRect(rect, clip);
      if (!clipped) return;
      const index = pageBreakPreviewPageNumber(rowIndex + 1, colIndex + 1, totalRows, totalCols, layout);
      const labelLeft = Math.max(clipped.left + 16, rect.left + 24);
      const labelTop = Math.max(clipped.top + 16, rect.top + 24);
      const labelStyle = `left:${roundCssPx(labelLeft - clipped.left)}px;top:${roundCssPx(labelTop - clipped.top)}px`;
      guides.push(`<div class="sheet-page-break-page" aria-hidden="true" style="left:${roundCssPx(clipped.left)}px;top:${roundCssPx(clipped.top)}px;width:${roundCssPx(clipped.width)}px;height:${roundCssPx(clipped.height)}px"><span class="sheet-page-break-label" style="${escapeAttr(labelStyle)}">${index} ページ</span></div>`);
    });
  });
}



function pageBreakPreviewPageNumber(pageRow, pageCol, totalRows, totalCols, layout) {
  const first = Math.max(1, Math.round(Number(layout?.firstPageNumber) || 1));
  const offset = layout?.pageOrder === "overThenDown"
    ? (pageRow - 1) * totalCols + pageCol
    : (pageCol - 1) * totalRows + pageRow;
  return first + offset - 1;
}



function pageBreakPreviewPlan(sheet, layout = normalizedPageLayoutSettings(sheet?.pageLayout, sheet), rangeOverride = null) {
  const range = normalizePrintRange(sheet, rangeOverride) || printRangeForSheet(sheet, layout) || normalizePrintRange(sheet, { top: 1, left: 1, bottom: 1, right: 1 });
  const page = pageBreakPreviewPrintablePageSize(sheet, layout, range);
  const rows = pageBreakAxisPlan(sheet, "row", range.top, range.bottom, page.height, layout.pageBreaks?.rows || []);
  const cols = pageBreakAxisPlan(sheet, "col", range.left, range.right, page.width, layout.pageBreaks?.cols || []);
  return { range, page, rows, cols };
}



function pageBreakPreviewPrintablePageSize(sheet, layout, range) {
  const page = printablePageSizePixels(sheet, layout);
  const rect = rangePixelRect(sheet, range);
  const { fitToWidth, fitToHeight } = pageLayoutFitTargets(layout);
  return {
    width: fitToWidth > 0 ? Math.max(page.width, rect.width / fitToWidth) : page.width,
    height: fitToHeight > 0 ? Math.max(page.height, rect.height / fitToHeight) : page.height,
  };
}



function pageBreakAxisPlan(sheet, axis, startIndex, endIndex, pageSize, manualBreaks = []) {
  const start = Math.max(1, Math.round(Number(startIndex) || 1));
  const end = Math.max(start, Math.round(Number(endIndex) || start));
  const page = Math.max(1, Number(pageSize) || 1);
  const edgeBreak = end + 1;
  const manual = [...new Set(
    (manualBreaks || [])
      .map((value) => Math.round(Number(value) || 0))
      .filter((value) => value > start && value <= edgeBreak),
  )].sort((a, b) => a - b);
  const breaks = [];

  if (manual.length) {
    manual.forEach((index) => breaks.push({ index, manual: true }));
    const lastManual = manual[manual.length - 1];
    if (lastManual <= end) {
      appendAutomaticPageBreaks(sheet, axis, lastManual, end, page, breaks);
    }
  } else {
    appendAutomaticPageBreaks(sheet, axis, start, end, page, breaks);
  }

  const deduped = [];
  const seen = new Set();
  breaks
    .sort((left, right) => left.index - right.index || Number(right.manual) - Number(left.manual))
    .forEach((breakInfo) => {
      if (seen.has(breakInfo.index)) return;
      seen.add(breakInfo.index);
      deduped.push(breakInfo);
    });

  const segments = [];
  let currentStart = start;
  deduped.forEach((breakInfo) => {
    if (breakInfo.index <= currentStart) return;
    const segmentEnd = Math.min(end, breakInfo.index - 1);
    if (currentStart <= segmentEnd) segments.push(pageBreakAxisSegment(sheet, axis, currentStart, segmentEnd));
    currentStart = breakInfo.index;
  });
  if (currentStart <= end) segments.push(pageBreakAxisSegment(sheet, axis, currentStart, end));
  if (!segments.length) segments.push(pageBreakAxisSegment(sheet, axis, start, end));
  return { breaks: deduped, breakIndexes: seen, segments };
}



function appendAutomaticPageBreaks(sheet, axis, start, end, pageSize, breaks) {
  let segmentStart = start;
  let used = 0;
  for (let index = start; index <= end; index += 1) {
    const size = Math.max(0, sheetAxisItemSize(sheet, axis, index - 1));
    if (used > 0 && used + size > pageSize && index > segmentStart) {
      breaks.push({ index, manual: false });
      segmentStart = index;
      used = 0;
    }
    used += size;
  }
}



function pageBreakAxisSegment(sheet, axis, start, end) {
  return {
    start,
    end,
    startPx: coordinateToPixels(sheet, axis, start - 1),
    endPx: coordinateToPixels(sheet, axis, end),
  };
}



function appendPageBreakOutsidePrintArea(guides, printRect, geometry) {
  const clip = { left: geometry.visibleLeft, top: geometry.visibleTop, right: geometry.visibleRight, bottom: geometry.visibleBottom };
  const grid = {
    left: geometry.leftOffset,
    top: geometry.topOffset,
    right: geometry.gridRight,
    bottom: geometry.gridBottom,
  };
  [
    { left: grid.left, top: grid.top, right: grid.right, bottom: printRect.top },
    { left: grid.left, top: printRect.bottom, right: grid.right, bottom: grid.bottom },
    { left: grid.left, top: printRect.top, right: printRect.left, bottom: printRect.bottom },
    { left: printRect.right, top: printRect.top, right: grid.right, bottom: printRect.bottom },
  ].forEach((rect) => {
    const clipped = clipPixelRect(rect, clip);
    if (!clipped || clipped.width <= 0 || clipped.height <= 0) return;
    const label = clipped.width >= 160 && clipped.height >= 52
      ? `<span class="sheet-page-break-outside-label">印刷範囲外</span>`
      : "";
    guides.push(`<div class="sheet-page-break-outside-print-area" aria-hidden="true" style="left:${roundCssPx(clipped.left)}px;top:${roundCssPx(clipped.top)}px;width:${roundCssPx(clipped.width)}px;height:${roundCssPx(clipped.height)}px">${label}</div>`);
  });
}



function pageBreakGuideAttributes(axis, index, manual) {
  return ` data-page-break-axis="${axis}" data-page-break-index="${Math.max(2, Number(index) || 2)}" data-page-break-manual="${manual ? "true" : "false"}"`;
}



function pageBreakIndexFromPixel(sheet, axis, pixel) {
  const isColumn = axis === "col";
  const headerSize = isColumn ? rowHeaderWidthForSheet(sheet) : COLUMN_HEADER_HEIGHT;
  const items = isColumn ? sheet.columns : sheet.rows;
  const defaultSize = isColumn ? sheet.defaultColumnWidth : sheet.defaultRowHeight;
  const fallbackSize = defaultSize || (isColumn ? columnWidthToPixels(8.43) : rowHeightToPixels(15));
  const maxIndex = Math.max(2, (isColumn ? sheet.colCount : sheet.rowCount) + 1);
  let position = headerSize;
  let bestIndex = 2;
  let bestDistance = Number.POSITIVE_INFINITY;

  for (let index = 1; index < maxIndex; index += 1) {
    const item = items[index - 1];
    const size = item?.hidden ? 0 : item?.width || item?.height || fallbackSize;
    position += Math.max(0, size);
    const distance = Math.abs(position - pixel);
    if (distance < bestDistance) {
      bestDistance = distance;
      bestIndex = index + 1;
    }
  }
  return Math.max(2, Math.min(maxIndex, bestIndex));
}



function startPageBreakDrag(event) {
  if (event.button !== 0 || !state.model) return;
  if (state.pendingInkTool) {
    startInkDrawing(event);
    return;
  }
  if (state.pendingShapeType) {
    startShapeDrawing(event);
    return;
  }
  const sheet = activeSheet();
  if ((sheet.viewMode || "normal") !== "page-break") return;
  const axis = event.currentTarget.dataset.pageBreakAxis === "row" ? "row" : "col";
  const layout = ensurePageLayoutSettings(sheet);
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  state.pageBreakDrag = {
    axis,
    sheetIndex: state.activeSheetIndex,
    startIndex: Math.max(2, Number(event.currentTarget.dataset.pageBreakIndex) || 2),
    currentIndex: Math.max(2, Number(event.currentTarget.dataset.pageBreakIndex) || 2),
    wasManual: event.currentTarget.dataset.pageBreakManual === "true",
    removing: false,
    startClientX: event.clientX,
    startClientY: event.clientY,
    moved: false,
    workbookViewBefore: captureWorkbookViewState(),
    sheetBefore: captureSheetForHistory(sheet),
    selectionBefore: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges().map((range) => ({ ...range })),
    originalRows: [...(layout.pageBreaks.rows || [])],
    originalCols: [...(layout.pageBreaks.cols || [])],
  };
  try {
    event.currentTarget.setPointerCapture?.(event.pointerId);
  } catch {}
  document.body.classList.add("is-dragging-page-break");
  setStatus(state.pageBreakDrag.wasManual ? "改ページをドラッグして移動または解除します。" : "自動改ページをドラッグすると手動改ページになります。");
}



function updatePageBreakDrag(event) {
  const job = state.pageBreakDrag;
  if (!job || !state.model || job.sheetIndex !== state.activeSheetIndex) return;
  const movedEnough = Math.abs(event.clientX - job.startClientX) + Math.abs(event.clientY - job.startClientY) >= 3;
  if (!movedEnough && !job.moved) return;
  const dropState = pageBreakDropState(job.axis, event.clientX, event.clientY, job);
  if (!dropState) return;
  event.preventDefault?.();
  job.moved = true;
  if (dropState.remove) {
    if (job.removing) return;
    const sheet = activeSheet();
    job.removing = true;
    job.currentIndex = null;
    applyDraggedPageBreakIndex(sheet, job, null);
    renderSheet();
    setStatus("改ページを解除します。");
    return;
  }
  const index = dropState.index;
  if (!index) return;
  if (!job.removing && index === job.currentIndex) return;
  const sheet = activeSheet();
  job.removing = false;
  job.currentIndex = index;
  applyDraggedPageBreakIndex(sheet, job, index);
  renderSheet();
  const label = job.axis === "col" ? `${columnName(index)}列の前` : `${index}行目の前`;
  setStatus(`改ページを ${label} に移動しています。`);
}



function finishPageBreakDrag() {
  const job = state.pageBreakDrag;
  if (!job) return;
  state.pageBreakDrag = null;
  document.body.classList.remove("is-dragging-page-break");
  if (!state.model || job.sheetIndex !== state.activeSheetIndex || !job.moved) {
    setStatus("改ページの移動をキャンセルしました。");
    return;
  }
  const sheet = activeSheet();
  const layout = ensurePageLayoutSettings(sheet);
  const changed =
    JSON.stringify(job.originalRows) !== JSON.stringify(layout.pageBreaks.rows || []) ||
    JSON.stringify(job.originalCols) !== JSON.stringify(layout.pageBreaks.cols || []);
  if (!changed) {
    setStatus("改ページは変更されませんでした。");
    return;
  }
  const targetIndex = job.currentIndex || job.startIndex;
  selectPageBreakTargets(pageBreakTargetsFromAxisIndex(job.axis, targetIndex), sheet);
  refreshActiveSheetViewSnapshot(sheet);
  saveActiveWindowViewState();
  pushHistory({
    sheetIndex: job.sheetIndex,
    sheetBefore: job.sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    workbookViewBefore: job.workbookViewBefore,
    workbookViewAfter: captureWorkbookViewState(),
    selectionBefore: job.selectionBefore,
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: job.selectionBeforeRanges,
    selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
    label: job.removing ? "改ページの解除" : "改ページの移動",
  });
  scheduleWorkbookPersistence();
  if (job.removing) {
    setStatus("改ページを解除しました。");
    return;
  }
  const label = job.axis === "col" ? `${columnName(job.currentIndex)}列の前` : `${job.currentIndex}行目の前`;
  setStatus(`改ページを ${label} に移動しました。`);
}



function startPrintAreaDrag(event) {
  if (event.button !== 0 || !state.model) return;
  if (state.pendingInkTool) {
    startInkDrawing(event);
    return;
  }
  if (state.pendingShapeType) {
    startShapeDrawing(event);
    return;
  }
  const sheet = activeSheet();
  if ((sheet?.viewMode || "normal") !== "page-break") return;
  const edge = String(event.currentTarget?.dataset?.printAreaEdge || "");
  if (!["top", "right", "bottom", "left"].includes(edge)) return;
  const layout = ensurePageLayoutSettings(sheet);
  const range = printRangeForSheet(sheet, layout);
  if (!range) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  state.printAreaDrag = {
    edge,
    sheetIndex: state.activeSheetIndex,
    startClientX: event.clientX,
    startClientY: event.clientY,
    moved: false,
    workbookViewBefore: captureWorkbookViewState(),
    sheetBefore: captureSheetForHistory(sheet),
    originalRange: { ...range },
    currentRange: { ...range },
  };
  try {
    event.currentTarget.setPointerCapture?.(event.pointerId);
  } catch {}
  document.body.classList.add("is-dragging-print-area");
  document.body.dataset.printAreaDragEdge = edge;
  setStatus("印刷範囲をドラッグして変更します。");
}



function updatePrintAreaDrag(event) {
  const job = state.printAreaDrag;
  if (!job || !state.model || job.sheetIndex !== state.activeSheetIndex) return;
  const movedEnough = Math.abs(event.clientX - job.startClientX) + Math.abs(event.clientY - job.startClientY) >= 3;
  if (!movedEnough && !job.moved) return;
  const sheet = activeSheet();
  const nextRange = printAreaRangeFromClientPosition(sheet, job, event.clientX, event.clientY);
  if (!nextRange) return;
  event.preventDefault?.();
  job.moved = true;
  if (sameRange(job.currentRange, nextRange)) return;
  job.currentRange = nextRange;
  applyDraggedPrintAreaRange(sheet, nextRange);
  renderSheet();
  setStatus(`印刷範囲：${rangeToLabel(nextRange)}`);
}



function finishPrintAreaDrag() {
  const job = state.printAreaDrag;
  if (!job) return;
  state.printAreaDrag = null;
  document.body.classList.remove("is-dragging-print-area");
  delete document.body.dataset.printAreaDragEdge;
  if (!state.model || job.sheetIndex !== state.activeSheetIndex || !job.moved) {
    setStatus("印刷範囲の変更をキャンセルしました。");
    return;
  }
  const sheet = activeSheet();
  const layout = ensurePageLayoutSettings(sheet);
  const finalRange = layout.printArea?.range || job.currentRange;
  if (sameRange(job.originalRange, finalRange)) {
    setStatus("印刷範囲は変更されませんでした。");
    return;
  }
  state.selected = selectableCellPoint(finalRange.top, finalRange.left);
  state.selectionRange = { ...finalRange };
  state.selectionRanges = [{ ...finalRange }];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  refreshActiveSheetViewSnapshot(sheet);
  saveActiveWindowViewState();
  pushHistory({
    sheetIndex: job.sheetIndex,
    sheetBefore: job.sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    workbookViewBefore: job.workbookViewBefore,
    workbookViewAfter: captureWorkbookViewState(),
    selectionBefore: job.originalRange,
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: [{ ...job.originalRange }],
    selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
    label: "印刷範囲の変更",
  });
  scheduleWorkbookPersistence();
  renderWorkbook();
  setStatus(`印刷範囲を ${rangeToLabel(finalRange)} に変更しました。`);
}



function printAreaRangeFromClientPosition(sheet, job, clientX, clientY) {
  if (!sheet || !job?.originalRange) return null;
  const point = sheetPixelFromClientPosition(clientX, clientY, { clampToGrid: true });
  if (!point) return null;
  const range = { ...job.currentRange };
  if (job.edge === "left" || job.edge === "right") {
    const col = sheetIndexFromPixelForPrintAreaDrag(sheet, "col", point.x);
    if (job.edge === "left") range.left = Math.min(Math.max(1, col), range.right);
    else range.right = Math.max(range.left, Math.min(Math.max(1, col), sheet.colCount || col));
  } else {
    const row = sheetIndexFromPixelForPrintAreaDrag(sheet, "row", point.y);
    if (job.edge === "top") range.top = Math.min(Math.max(1, row), range.bottom);
    else range.bottom = Math.max(range.top, Math.min(Math.max(1, row), sheet.rowCount || row));
  }
  return normalizePrintRange(sheet, range);
}



function sheetIndexFromPixelForPrintAreaDrag(sheet, axis, pixel) {
  const isColumn = axis === "col";
  const headerSize = isColumn ? rowHeaderWidthForSheet(sheet) : COLUMN_HEADER_HEIGHT;
  if (Number(pixel) <= headerSize) return 1;
  const index = sheetIndexFromPixel(sheet, axis, pixel);
  if (index != null) return index;
  return Math.max(1, isColumn ? Number(sheet?.colCount || 1) : Number(sheet?.rowCount || 1));
}



function applyDraggedPrintAreaRange(sheet, range) {
  const layout = ensurePageLayoutSettings(sheet);
  const normalized = normalizePrintRange(sheet, range);
  layout.printArea = normalized ? { range: { ...normalized }, label: rangeToLabel(normalized) } : null;
  sheet.pageLayout = normalizedPageLayoutSettings(layout, sheet);
}



function sameRange(left, right) {
  return Boolean(left && right) &&
    Number(left.top) === Number(right.top) &&
    Number(left.left) === Number(right.left) &&
    Number(left.bottom) === Number(right.bottom) &&
    Number(left.right) === Number(right.right);
}



function pageBreakIndexFromClientPosition(axis, clientX, clientY) {
  return pageBreakDropState(axis, clientX, clientY)?.index || null;
}



function pageBreakDropState(axis, clientX, clientY, job = null) {
  const host = $sheetHost[0];
  const sheet = activeSheet();
  if (!host || !sheet) return null;
  const rect = host.getBoundingClientRect();
  const headerLeft = rect.left + rowHeaderWidthForSheet(sheet);
  const headerTop = rect.top + COLUMN_HEADER_HEIGHT;
  const removeThreshold = 6;
  const canRemove = Boolean(job?.wasManual);
  if (canRemove) {
    const removeColumnBreak = axis === "col" && clientX <= headerLeft + removeThreshold;
    const removeRowBreak = axis === "row" && clientY <= headerTop + removeThreshold;
    if (removeColumnBreak || removeRowBreak) return { remove: true, index: null };
  }
  const point = sheetPixelFromClientPosition(clientX, clientY, { clampToGrid: true });
  if (!point) return null;
  const edgeIndex = pageBreakEdgeSnapIndex(sheet, axis, point);
  return { remove: false, index: edgeIndex || pageBreakIndexFromPixel(sheet, axis, axis === "col" ? point.x : point.y) };
}



function pageBreakEdgeSnapIndex(sheet, axis, point) {
  const layout = normalizedPageLayoutSettings(sheet?.pageLayout, sheet);
  const plan = pageBreakPreviewPlan(sheet, layout);
  const range = plan?.range;
  if (!range) return null;
  const rect = rangePixelRect(sheet, range);
  const threshold = 10;
  if (axis === "col") {
    const x = Number(point?.x);
    if (Number.isFinite(x) && Math.abs(x - rect.right) <= threshold) return range.right + 1;
  } else {
    const y = Number(point?.y);
    if (Number.isFinite(y) && Math.abs(y - rect.bottom) <= threshold) return range.bottom + 1;
  }
  return null;
}



function applyDraggedPageBreakIndex(sheet, job, index) {
  const layout = ensurePageLayoutSettings(sheet);
  const rows = [...job.originalRows];
  const cols = [...job.originalCols];
  const list = job.axis === "row" ? rows : cols;
  const next = list.filter((value) => !(job.wasManual && value === job.startIndex));
  if (index != null) next.push(index);
  const normalized = [...new Set(next.map(Number).filter((value) => value > 1))].sort((a, b) => a - b);
  if (job.axis === "row") layout.pageBreaks = { rows: normalized, cols };
  else layout.pageBreaks = { rows, cols: normalized };
  sheet.pageLayout = normalizedPageLayoutSettings(layout, sheet);
}



function printablePageSizePixels(sheet, layout) {
  const paper = PAGE_LAYOUT_PAPER_SIZES.find((item) => item.value === layout.paperSize) || PAGE_LAYOUT_PAPER_SIZES[0];
  const landscape = layout.orientation === "landscape";
  const paperWidth = landscape ? paper.height : paper.width;
  const paperHeight = landscape ? paper.width : paper.height;
  const margins = layout.margins || PAGE_LAYOUT_MARGIN_PRESETS.normal;
  const scale = Math.max(0.1, Number(layout.scale || 100) / 100);
  const displayScale = Math.max(0.1, Number(sheet.displayScale || 1));
  const width = Math.max(160, ((paperWidth - Number(margins.left || 0) - Number(margins.right || 0)) * 96 * displayScale) / scale);
  const height = Math.max(160, ((paperHeight - Number(margins.top || 0) - Number(margins.bottom || 0)) * 96 * displayScale) / scale);
  return { width, height };
}



function roundCssPx(value) {
  return Math.round(Number(value || 0) * 100) / 100;
}



function renderCanvasLayer(sheet, visibleRange) {
  const bounds = canvasLayerBounds(sheet, visibleRange);
  const ratio = canvasDevicePixelRatio();
  const width = Math.max(1, Math.ceil(bounds.width));
  const height = Math.max(1, Math.ceil(bounds.height));
  const pixelWidth = Math.max(1, Math.ceil(width * ratio));
  const pixelHeight = Math.max(1, Math.ceil(height * ratio));
  const style = `left:${bounds.left}px;top:${bounds.top}px;width:${width}px;height:${height}px`;
  return `<canvas class="sheet-canvas" width="${pixelWidth}" height="${pixelHeight}" style="${escapeAttr(style)}" aria-hidden="true" data-dpr="${ratio}" data-left="${bounds.left}" data-top="${bounds.top}"></canvas>`;
}



function paintSheetCanvas(sheet, mergeMaps, sheetId, visibleRange) {
  const canvas = $sheetHost.find(".sheet-canvas")[0];
  if (!canvas) return;
  const context = canvas.getContext("2d");
  if (!context) return;

  const ratio = Number(canvas.dataset.dpr || 1);
  const width = Number.parseFloat(canvas.style.width) || canvas.width / ratio;
  const height = Number.parseFloat(canvas.style.height) || canvas.height / ratio;
  const left = Number(canvas.dataset.left || 0);
  const top = Number(canvas.dataset.top || 0);
  context.setTransform(ratio, 0, 0, ratio, 0, 0);
  context.clearRect(0, 0, width, height);
  context.save();
  context.beginPath();
  context.rect(0, 0, width, height);
  context.clip();
  context.translate(-left, -top);

  const pixelCache = createPixelPositionCache(sheet);
  paintCanvasCells(context, sheet, mergeMaps, sheetId, visibleRange, pixelCache, "base");
  paintCanvasBorders(context, sheet, mergeMaps, visibleRange, pixelCache);
  paintCanvasCells(context, sheet, mergeMaps, sheetId, visibleRange, pixelCache, "text");
  context.restore();
}



function paintCanvasCells(context, sheet, mergeMaps, sheetId, visibleRange, pixelCache = null, phase = "all") {
  const showGridLines = sheet.showGridLines !== false;
  const rendered = new Set();
  const editedCellKey = phase === "base" ? "" : canvasEditedCellKey();
  const paintRange = phase === "text" ? canvasTextPaintRange(sheet, visibleRange, pixelCache) : visibleRange;

  for (let row = paintRange.top; row <= paintRange.bottom; row += 1) {
    for (let col = paintRange.left; col <= paintRange.right; col += 1) {
      const key = cellKey(row, col);
      const coveredMasterKey = mergeMaps.coveredToMaster.get(key);
      const masterKey = coveredMasterKey || key;
      if (rendered.has(masterKey)) continue;
      const merge = mergeMaps.master.get(masterKey);
      let range = null;
      if (merge) {
        range = merge;
      } else if (coveredMasterKey) {
        const [masterRow, masterCol] = parseCellKey(masterKey);
        range = { top: masterRow, left: masterCol, bottom: masterRow, right: masterCol };
      } else {
        range = { top: row, left: col, bottom: row, right: col };
      }
      if (phase !== "text" && !rangeIntersects(range, visibleRange)) continue;
      rendered.add(masterKey);
      paintCanvasCell(context, sheet, mergeMaps, sheetId, range.top, range.left, range, showGridLines, pixelCache, phase, editedCellKey);
    }
  }
}



function canvasTextPaintRange(sheet, visibleRange, pixelCache = null) {
  if (!visibleRange || sheetColumnsRightToLeft(sheet)) return visibleRange;
  const visibleLeft = cachedCoordinateToPixels(sheet, pixelCache, "col", Math.max(0, visibleRange.left - 1));
  const visibleRight = cachedCoordinateToPixels(sheet, pixelCache, "col", visibleRange.right);
  const visibleWidth = Math.max(1, visibleRight - visibleLeft);
  const lookbackPx = Math.max(
    TEXT_OVERFLOW_LOOKBACK_MIN_PX,
    Math.min(TEXT_OVERFLOW_LOOKBACK_MAX_PX, visibleWidth * TEXT_OVERFLOW_LOOKBACK_VIEWPORT_FACTOR),
  );
  const gridLeft = sheetHeadingsVisible(sheet) ? rowHeaderWidthForSheet(sheet) : 0;
  const scanLeft = Math.max(gridLeft, visibleLeft - lookbackPx);
  const left = Math.min(visibleRange.left, Math.max(1, indexForPixelExtent(sheet, "col", scanLeft)));
  return { ...visibleRange, left };
}



function paintCanvasCell(context, sheet, mergeMaps, sheetId, row, col, range, showGridLines, pixelCache = null, phase = "all", editedCellKey = "") {
  const key = cellKey(row, col);
  const sourceCell = sheet.cells[key] || null;
  const cell = sheetDisplayCell(sheet, row, col, sourceCell) || { row, col };
  const rowModel = sheet.rows[row - 1] || {};
  const column = sheet.columns[col - 1] || {};
  if (rowModel.hidden || column.hidden) return;

  const rect = rangePixelRect(sheet, range, pixelCache);
  if (rect.width <= 0 || rect.height <= 0) return;

  if (phase !== "text") {
    if (cellHasFill(cell)) {
      context.fillStyle = cell.css?.backgroundColor || "#ffffff";
      context.fillRect(rect.left, rect.top, rect.width, rect.height);
    }

    if (showGridLines) {
      paintCanvasGridLines(context, sheet, mergeMaps, row, col, range, cell, rect);
    }

    if (cell.quickAnalysisBar) {
      paintCanvasQuickAnalysisDataBar(context, cell.quickAnalysisBar, rect);
    }

    if (cell.sparkline) {
      paintCanvasSparkline(context, cell.sparkline, rect);
    }

    if (cell.quickAnalysisIcon) {
      paintCanvasQuickAnalysisIcon(context, cell.quickAnalysisIcon, rect);
    }
  }

  if (phase === "base") return;
  if (editedCellKey && editedCellKey === key) return;
  if (cellHasCanvasTextCandidate(cell) || (sourceCell && getCellFormula(sheetId, row, col))) {
    paintCanvasCellText(context, sheet, mergeMaps, sheetId, row, col, range, cell, rect);
  }
}



function canvasEditedCellKey() {
  if (state.formulaEdit?.sheetIndex === state.activeSheetIndex) {
    return cellKey(state.formulaEdit.row, state.formulaEdit.col);
  }
  const editor = $sheetHost.find(".sheet-cell[contenteditable='true']")[0];
  if (!editor?.isConnected) return "";
  const row = Number(editor.dataset.row);
  const col = Number(editor.dataset.col);
  return row && col ? cellKey(row, col) : "";
}



function repaintActiveSheetCanvas() {
  const sheet = activeSheet();
  if (!sheet) return;
  CanvasRenderer.paint(sheet, buildMergeMaps(sheet.merges), getSheetId(sheet.name), visibleGridRange(sheet));
}



function rangePixelRect(sheet, range, pixelCache = null) {
  const rawLeft = cachedCoordinateToPixels(sheet, pixelCache, "col", range.left - 1);
  const top = cachedCoordinateToPixels(sheet, pixelCache, "row", range.top - 1);
  const rawRight = cachedCoordinateToPixels(sheet, pixelCache, "col", range.right);
  const bottom = cachedCoordinateToPixels(sheet, pixelCache, "row", range.bottom);
  const left = Math.min(rawLeft, rawRight);
  const right = Math.max(rawLeft, rawRight);
  return {
    left,
    top,
    right,
    bottom,
    width: Math.max(0, right - left),
    height: Math.max(0, bottom - top),
  };
}



function createPixelPositionCache(sheet) {
  return {
    sheet,
    row: new Map([[0, COLUMN_HEADER_HEIGHT]]),
    col: new Map([[0, rowHeaderWidthForSheet(sheet)]]),
  };
}



function cachedCoordinateToPixels(sheet, pixelCache, axis, value, nativeOffset = 0, nativeIndex = null) {
  const isColumn = axis === "col";
  if (!pixelCache || pixelCache.sheet !== sheet || (isColumn && sheetColumnsRightToLeft(sheet))) {
    return coordinateToPixels(sheet, axis, value, nativeOffset, nativeIndex);
  }
  const hasNativeAnchor = nativeIndex !== null && Number.isFinite(Number(nativeIndex));
  const coordinate = hasNativeAnchor ? Number(nativeIndex) : Number.isFinite(Number(value)) ? Number(value) : 0;
  const whole = Math.max(0, Math.floor(coordinate));
  const fraction = hasNativeAnchor ? 0 : Math.max(0, coordinate - whole);
  const base = cachedBoundaryPosition(sheet, pixelCache, axis, whole);
  const currentSize = sheetAxisItemSize(sheet, axis, whole);
  let pixels = base + fraction * currentSize;
  if (hasNativeAnchor && nativeOffset) pixels += (Number(nativeOffset) / EMU_PER_PIXEL) * (sheet.displayScale || 1);
  return Math.round(pixels * 100) / 100;
}



function cachedBoundaryPosition(sheet, pixelCache, axis, boundary) {
  const safeBoundary = Math.max(0, Math.floor(Number(boundary) || 0));
  const map = axis === "col" ? pixelCache.col : pixelCache.row;
  if (map.has(safeBoundary)) return map.get(safeBoundary);

  let start = 0;
  let position = map.get(0) ?? (axis === "col" ? rowHeaderWidthForSheet(sheet) : COLUMN_HEADER_HEIGHT);
  map.forEach((value, key) => {
    if (key <= safeBoundary && key >= start) {
      start = key;
      position = value;
    }
  });

  for (let index = start; index < safeBoundary; index += 1) {
    position += sheetAxisItemSize(sheet, axis, index);
    map.set(index + 1, Math.round(position * 100) / 100);
  }
  return map.get(safeBoundary) ?? Math.round(position * 100) / 100;
}



function sheetAxisItemSize(sheet, axis, zeroBasedIndex) {
  const isColumn = axis === "col";
  const item = isColumn ? sheet.columns[zeroBasedIndex] : sheet.rows[zeroBasedIndex];
  const fallback = isColumn ? sheet.defaultColumnWidth : sheet.defaultRowHeight;
  return item?.hidden ? 0 : item?.width || item?.height || fallback || 0;
}



function cellHasCanvasTextCandidate(cell) {
  return cellHasDomContent(cell);
}



function paintCanvasGridLines(context, sheet, mergeMaps, row, col, range, cell, rect) {
  const rowSpan = range.bottom - range.top + 1;
  const colSpan = range.right - range.left + 1;
  context.save();
  context.strokeStyle = "#d9e0ea";
  context.lineWidth = 1;
  context.beginPath();
  if (!shouldHideGridRight(sheet, mergeMaps, row, col, rowSpan, colSpan, cell)) {
    crispCanvasLine(context, rect.right, rect.top, rect.right, rect.bottom);
  }
  if (!shouldHideGridBottom(sheet, mergeMaps, row, col, rowSpan, colSpan, cell)) {
    crispCanvasLine(context, rect.left, rect.bottom, rect.right, rect.bottom);
  }
  context.stroke();
  context.restore();
}



function paintCanvasQuickAnalysisDataBar(context, dataBar, rect) {
  const ratio = Math.max(0, Math.min(1, Number(dataBar?.ratio) || 0));
  if (ratio <= 0 || rect.width <= 6 || rect.height <= 6) return;
  context.save();
  context.fillStyle = dataBar.color || "#5b9bd5";
  context.globalAlpha = 0.68;
  const insetX = 2;
  const insetY = 3;
  context.fillRect(
    rect.left + insetX,
    rect.top + insetY,
    Math.max(1, (rect.width - insetX * 2) * ratio),
    Math.max(1, rect.height - insetY * 2),
  );
  context.restore();
}



function paintCanvasQuickAnalysisIcon(context, icon, rect) {
  if (rect.width <= 18 || rect.height <= 10) return;
  const color = icon?.color || "#4472c4";
  const size = Math.min(10, Math.max(6, rect.height - 8));
  const x = rect.left + 5 + size / 2;
  const y = rect.top + rect.height / 2;
  context.save();
  context.fillStyle = color;
  context.strokeStyle = "#ffffff";
  context.lineWidth = 1;
  context.beginPath();
  if (icon?.shape === "same") {
    context.arc(x, y, size / 2, 0, Math.PI * 2);
  } else if (icon?.shape === "down") {
    context.moveTo(x - size / 2, y - size / 3);
    context.lineTo(x + size / 2, y - size / 3);
    context.lineTo(x, y + size / 2);
    context.closePath();
  } else {
    context.moveTo(x, y - size / 2);
    context.lineTo(x + size / 2, y + size / 3);
    context.lineTo(x - size / 2, y + size / 3);
    context.closePath();
  }
  context.fill();
  context.stroke();
  context.restore();
}



function paintCanvasSparkline(context, sparkline, rect) {
  const values = (sparkline?.values || []).map((value) => Number(value)).filter(Number.isFinite);
  if (!values.length || rect.width <= 8 || rect.height <= 8) return;
  const type = sparkline.type || "line";
  const padX = 4;
  const padY = 3;
  const left = rect.left + padX;
  const right = rect.right - padX;
  const top = rect.top + padY;
  const bottom = rect.bottom - padY;
  const min = Math.min(...values, type === "winloss" ? -1 : 0);
  const max = Math.max(...values, type === "winloss" ? 1 : 0);
  const span = max === min ? 1 : max - min;
  const xFor = (index) => left + (values.length <= 1 ? 0 : (index / (values.length - 1)) * (right - left));
  const yFor = (value) => bottom - ((value - min) / span) * (bottom - top);

  context.save();
  context.beginPath();
  context.rect(rect.left + 1, rect.top + 1, Math.max(1, rect.width - 2), Math.max(1, rect.height - 2));
  context.clip();
  context.strokeStyle = "#4472c4";
  context.fillStyle = "#4472c4";
  context.lineWidth = 1.5;

  if (type === "column" || type === "winloss") {
    const zeroY = type === "winloss" ? rect.top + rect.height / 2 : yFor(0);
    const barWidth = Math.max(2, (right - left) / Math.max(1, values.length) - 2);
    values.forEach((value, index) => {
      const x = left + index * ((right - left) / Math.max(1, values.length));
      const positive = value >= 0;
      const y = type === "winloss" ? (positive ? top : zeroY) : Math.min(yFor(value), zeroY);
      const h = type === "winloss" ? Math.max(2, (bottom - top) / 2 - 1) : Math.max(1, Math.abs(zeroY - yFor(value)));
      context.fillStyle = positive ? "#4472c4" : "#c00000";
      context.fillRect(Math.round(x), Math.round(y), Math.round(barWidth), Math.round(h));
    });
    context.strokeStyle = "#9aa6b2";
    context.beginPath();
    crispCanvasLine(context, left, zeroY, right, zeroY);
    context.stroke();
    context.restore();
    return;
  }

  context.beginPath();
  values.forEach((value, index) => {
    const x = xFor(index);
    const y = yFor(value);
    if (index === 0) context.moveTo(x, y);
    else context.lineTo(x, y);
  });
  context.stroke();
  const lastX = xFor(values.length - 1);
  const lastY = yFor(values[values.length - 1]);
  context.fillStyle = "#4472c4";
  context.beginPath();
  context.arc(lastX, lastY, 2.4, 0, Math.PI * 2);
  context.fill();
  context.restore();
}



function crispCanvasLine(context, x1, y1, x2, y2) {
  const crispX1 = Math.round(x1) + (Math.abs(x1 - x2) < 0.01 ? 0.5 : 0);
  const crispX2 = Math.round(x2) + (Math.abs(x1 - x2) < 0.01 ? 0.5 : 0);
  const crispY1 = Math.round(y1) + (Math.abs(y1 - y2) < 0.01 ? 0.5 : 0);
  const crispY2 = Math.round(y2) + (Math.abs(y1 - y2) < 0.01 ? 0.5 : 0);
  context.moveTo(crispX1, crispY1);
  context.lineTo(crispX2, crispY2);
}



function cellShowsValidationDropdown(cell, sheet = activeSheet()) {
  return Boolean(cell?.validation?.type === "list" && cell.validation.showDropDown !== false);
}



function paintCanvasCellText(context, sheet, mergeMaps, sheetId, row, col, range, cell, rect) {
  const formula = cell.formula || getCellFormula(sheetId, row, col);
  const display = state.showFormulas && formula ? formula : getDisplayForCell(sheet, sheetId, row, col);
  if (display == null || display === "") return;

  const text = String(display);
  const css = cellEffectiveCss(sheet, cell);
  const controlInset = autoFitCellControlInset(sheet, row, col, cell) * canvasFontDisplayScale(sheet);
  const clipWidth = (textOverflowWidth(sheet, mergeMaps, row, col, range.right - range.left + 1, cell, display) || rect.width) - controlInset;
  let clipRect = { left: rect.left, top: rect.top, width: Math.max(1, clipWidth), height: rect.height };
  // Match the space reserved by the DOM filter/validation buttons and auto-fit.
  // Both wrapped text and right/center alignment must use the remaining area.
  if (controlInset) rect = { ...rect, width: Math.max(1, rect.width - controlInset), right: rect.right - controlInset };
  const displayScale = canvasFontDisplayScale(sheet);
  if (displayScale !== 1) {
    rect = Object.fromEntries(Object.entries(rect).map(([key, value]) => [key, value / displayScale]));
    clipRect = Object.fromEntries(Object.entries(clipRect).map(([key, value]) => [key, value / displayScale]));
  }
  context.save();
  try {
    if (displayScale !== 1) context.scale(displayScale, displayScale);
    paintCanvasCellTextLayout(context, text, css, cell, rect, clipRect);
  } finally {
    context.restore();
  }
}



function paintCanvasCellTextLayout(context, text, css, cell, rect, clipRect) {
  const paddingX = 3;
  const fontScale = 1;
  const fontSize = canvasFontSize(css.fontSize, fontScale);
  const lineHeight = canvasTextLineHeight(fontSize);
  const font = canvasFont(css, fontSize);
  if (css.textRotation === "vertical") {
    paintCanvasVerticalCellText(context, text, css, rect, clipRect, fontSize, font);
    return;
  }
  const isWrapped = css.whiteSpace === "pre-wrap";
  const maxTextWidth = Math.max(1, clipRect.width - paddingX * 2);
  const richTextRuns = !state.showFormulas ? richTextRunsForDisplay(cell, text) : null;
  if (richTextRuns?.length) {
    paintCanvasRichCellText(context, richTextRuns, css, rect, clipRect, fontSize, lineHeight, paddingX, isWrapped, fontScale);
    return;
  }
  const lines = isWrapped ? wrapCanvasText(context, text, maxTextWidth, font) : String(text).split(/\r\n|\r|\n/).slice(0, 1);
  const phoneticRuns = !state.showFormulas && lines.length === 1 ? phoneticRunsForLine(cell, lines[0]) : [];
  const phoneticFontSize = phoneticRuns.length ? Math.max(7, Math.min(fontSize * 0.58, fontSize - 2)) : 0;
  const phoneticLineHeight = phoneticRuns.length ? Math.ceil(phoneticFontSize * 1.1) : 0;
  const totalTextHeight = Math.max(lineHeight, lines.length * lineHeight + phoneticLineHeight);
  const vertical = css.verticalAlign || "middle";
  let y = rect.top + Math.max(0, (rect.height - totalTextHeight) / 2);
  if (vertical === "top") y = rect.top + 2;
  if (vertical === "bottom") y = rect.bottom - totalTextHeight - 2;

  context.save();
  context.beginPath();
  context.rect(clipRect.left, clipRect.top, clipRect.width, clipRect.height);
  context.clip();
  context.font = font;
  context.fillStyle = css.color || "#1d2433";
  context.textBaseline = "alphabetic";

  const align = css.textAlign || "left";
  context.textAlign = align === "right" ? "right" : align === "center" ? "center" : "left";
  let x = rect.left + paddingX;
  if (context.textAlign === "center") x = rect.left + rect.width / 2;
  if (context.textAlign === "right") x = rect.right - paddingX;

  lines.forEach((line, index) => {
    const lineY = y + index * lineHeight;
    if (lineY > clipRect.top + clipRect.height) return;
    const mainLineY = lineY + (index === 0 ? phoneticLineHeight : 0);
    const baselineY = canvasLineBaseline(context, line, mainLineY, lineHeight, fontSize);
    if (index === 0 && phoneticRuns.length) {
      paintCanvasPhonetics(context, css, line, x, context.textAlign, lineY, phoneticRuns, phoneticFontSize);
    }
    context.fillText(line, x, baselineY);
    const textDecoration = String(css.textDecoration || "");
    if (textDecoration.includes("underline")) {
      paintCanvasUnderline(context, line, x, baselineY, lineHeight, textDecoration.includes("double"));
    }
    if (textDecoration.includes("line-through")) {
      paintCanvasStrikethrough(context, line, x, baselineY, lineHeight);
    }
  });
  context.restore();
}



function richTextRunsForDisplay(cell, displayText) {
  if (!cell?.richText?.length) return null;
  const runs = cell.richText
    .map((run) => ({
      text: normalizeRichTextLineEndings(run?.text),
      css: run?.css && typeof run.css === "object" ? run.css : {},
    }))
    .filter((run) => run.text);
  if (!runs.length) return null;
  const richText = normalizeRichTextLineEndings(runs.map((run) => run.text).join(""));
  return richText === normalizeRichTextLineEndings(displayText) ? runs : null;
}



function normalizeRichTextLineEndings(value) {
  return String(value ?? "").replace(/\r\n|\r/g, "\n");
}



function canvasRichLineMetrics(lines, css, fontScale = 1) {
  const base = canvasFontSize(css.fontSize, fontScale);
  const fonts = lines.map(line => Math.max(base, ...line.map(run => canvasFontSize(run.css?.fontSize || css.fontSize, fontScale))));
  const heights = fonts.map(canvasTextLineHeight);
  return {fonts,heights,totalHeight:heights.reduce((sum,n)=>sum+n,0)};
}



function paintCanvasRichCellText(context, runs, css, rect, clipRect, fontSize, lineHeight, paddingX, isWrapped, fontScale = 1) {
  const maxTextWidth = Math.max(1, clipRect.width - paddingX * 2);
  const lines = isWrapped ? wrapCanvasRichText(context, runs, css, maxTextWidth, fontScale) : splitRichRunsIntoHardLines(runs).slice(0, 1);
  const metrics = canvasRichLineMetrics(lines, css, fontScale || 1);
  const totalTextHeight = metrics.totalHeight;
  const vertical = css.verticalAlign || "middle";
  let y = rect.top + Math.max(0, (rect.height - totalTextHeight) / 2);
  if (vertical === "top") y = rect.top + 2;
  if (vertical === "bottom") y = rect.bottom - totalTextHeight - 2;

  context.save();
  context.beginPath();
  context.rect(clipRect.left, clipRect.top, clipRect.width, clipRect.height);
  context.clip();
  context.textBaseline = "alphabetic";
  const align = css.textAlign || "left";
  const textAlign = align === "right" ? "right" : align === "center" ? "center" : "left";
  lines.forEach((line, index) => {
    const lineY = y + metrics.heights.slice(0,index).reduce((sum,n)=>sum+n,0);
    const actualLineHeight=metrics.heights[index],actualFontSize=metrics.fonts[index];
    if (lineY > clipRect.top + clipRect.height) return;
    const text = richLinePlainText(line);
    context.font = canvasFont(css, actualFontSize);
    const baselineY = canvasLineBaseline(context, text, lineY, actualLineHeight, actualFontSize);
    const lineWidth = measureRichLine(context, line, css, fontScale);
    let x = rect.left + paddingX;
    if (textAlign === "center") x = rect.left + rect.width / 2 - lineWidth / 2;
    if (textAlign === "right") x = rect.right - paddingX - lineWidth;
    drawCanvasRichLine(context, line, css, x, baselineY, actualLineHeight, fontScale);
  });
  context.restore();
}



function splitRichRunsIntoHardLines(runs) {
  const lines = [];
  let current = [];
  runs.forEach((run) => {
    normalizeRichTextLineEndings(run.text)
      .split("\n")
      .forEach((part, index, parts) => {
        if (part) current.push({ text: part, css: run.css || {} });
        if (index < parts.length - 1) {
          lines.push(current);
          current = [];
        }
      });
  });
  lines.push(current);
  return lines;
}



function wrapCanvasRichText(context, runs, css, maxWidth, fontScale = 1) {
  const lines = [];
  for (const hardLine of splitRichRunsIntoHardLines(runs)) {
    const text = richLinePlainText(hardLine);
    const ranges = measuredLineRanges(text, maxWidth,
      (start, end) => measureRichLine(context, sliceStyledLine(hardLine, start, end), css, fontScale));
    for (const [start, end] of ranges) lines.push(sliceStyledLine(hardLine, start, end));
  }
  return lines.slice(0, 200);
}



function appendRichChar(line, char, css) {
  const next = line.map((segment) => ({ text: segment.text, css: segment.css }));
  const last = next[next.length - 1];
  if (last && stableObjectKey(last.css || {}) === stableObjectKey(css || {})) {
    last.text += char;
  } else {
    next.push({ text: char, css: css || {} });
  }
  return next;
}



function richLinePlainText(line) {
  return (line || []).map((segment) => segment.text || "").join("");
}



function richSegmentCss(baseCss, segmentCss) {
  return { ...(baseCss || {}), ...(segmentCss || {}) };
}



function measureRichLine(context, line, css, fontScale = 1) {
  return (line || []).reduce((width, segment) => {
    const mergedCss = richSegmentCss(css, segment.css);
    context.font = canvasFont(mergedCss, canvasFontSize(mergedCss.fontSize, fontScale));
    return width + context.measureText(segment.text || "").width;
  }, 0);
}



function drawCanvasRichLine(context, line, css, x, baselineY, lineHeight, fontScale = 1) {
  let cursor = x;
  (line || []).forEach((segment) => {
    const mergedCss = richSegmentCss(css, segment.css);
    context.font = canvasFont(mergedCss, canvasFontSize(mergedCss.fontSize, fontScale));
    context.fillStyle = mergedCss.color || "#1d2433";
    context.textAlign = "left";
    const text = segment.text || "";
    context.fillText(text, cursor, baselineY);
    const textDecoration = String(mergedCss.textDecoration || "");
    if (textDecoration.includes("underline")) {
      const width = context.measureText(text).width;
      const underlineY = baselineY + Math.max(1, Math.min(3, lineHeight * 0.12));
      const underlineGap = Math.max(2, Math.min(3, lineHeight * 0.16));
      context.save();
      context.strokeStyle = context.fillStyle;
      context.lineWidth = 1;
      context.beginPath();
      context.moveTo(cursor, underlineY);
      context.lineTo(cursor + width, underlineY);
      if (textDecoration.includes("double")) {
        context.moveTo(cursor, underlineY + underlineGap);
        context.lineTo(cursor + width, underlineY + underlineGap);
      }
      context.stroke();
      context.restore();
    }
    if (textDecoration.includes("line-through")) {
      const width = context.measureText(text).width;
      const strikeY = baselineY - Math.max(3, Math.min(lineHeight * 0.34, lineHeight - 2));
      context.save();
      context.strokeStyle = context.fillStyle;
      context.lineWidth = 1;
      context.beginPath();
      context.moveTo(cursor, strikeY);
      context.lineTo(cursor + width, strikeY);
      context.stroke();
      context.restore();
    }
    cursor += context.measureText(text).width;
  });
}



function paintCanvasVerticalCellText(context, text, css, rect, clipRect, fontSize, font) {
  const chars = Array.from(String(text).replace(/\r\n|\r|\n/g, ""));
  if (!chars.length) return;
  const advance = Math.max(10, fontSize * 1.08);
  const totalHeight = chars.length * advance;
  const vertical = css.verticalAlign || "middle";
  let y = rect.top + Math.max(0, (rect.height - totalHeight) / 2);
  if (vertical === "top") y = rect.top + 2;
  if (vertical === "bottom") y = rect.bottom - totalHeight - 2;
  const align = css.textAlign || "center";
  let x = rect.left + rect.width / 2;
  if (align === "left") x = rect.left + Math.min(rect.width / 2, fontSize * 0.7 + 2);
  if (align === "right") x = rect.right - Math.min(rect.width / 2, fontSize * 0.7 + 2);

  context.save();
  context.beginPath();
  context.rect(clipRect.left, clipRect.top, clipRect.width, clipRect.height);
  context.clip();
  context.font = font;
  context.fillStyle = css.color || "#1d2433";
  context.textAlign = "center";
  context.textBaseline = "middle";
  chars.forEach((char, index) => {
    const charY = y + advance * index + advance / 2;
    if (charY < clipRect.top - advance || charY > clipRect.top + clipRect.height + advance) return;
    context.fillText(char, x, charY);
  });
  context.restore();
}

export {
  adjustSelectionDecimalPlaces,
  appendAutomaticPageBreaks,
  appendPageBreakOutsidePrintArea,
  appendPageBreakPreviewPages,
  appendPageLayoutRect,
  appendPageLayoutWhiteSpaceEdges,
  appendRichChar,
  applyBorderPreset,
  applyBorderPresetToSheetGroup,
  applyBorderPresetToSheetRange,
  applyCellFormat,
  applyCellNumberFormat,
  applyCssPatchToSheetRanges,
  applyCssToSelection,
  applyCssToSheetGroup,
  applyDraggedPageBreakIndex,
  applyDraggedPrintAreaRange,
  applyFormatPainterAtPoint,
  applyFormatPainterMerges,
  applyFormatPainterToRange,
  applyNumberFormatToSelection,
  applyNumberFormatToSheetGroup,
  applyShapeTextCssPatchToSelection,
  borderTargetForCell,
  cachedBoundaryPosition,
  cachedCoordinateToPixels,
  canvasEditedCellKey,
  canvasRichLineMetrics,
  canvasTextPaintRange,
  captureFormatPainterFormats,
  captureFormatPainterMerges,
  cellCanRecomputeDisplayText,
  cellHasCanvasTextCandidate,
  cellHasMergeWarningContent,
  cellShowsValidationDropdown,
  clearFormatPainter,
  clipPixelRect,
  cloneCellBorders,
  cloneCellFormat,
  createPixelPositionCache,
  crispCanvasLine,
  cssPatchResetsCellRichText,
  cssPatchRichTextOverrideKeys,
  cssUrlForStyle,
  dateSerialFromValueForNumberFormat,
  drawCanvasRichLine,
  ensureCellModel,
  finishPageBreakDrag,
  finishPrintAreaDrag,
  finishSheetStructureHistoryRestore,
  forEachCellInRange,
  formulaAuditArrowPoint,
  formulaAuditArrowSvg,
  formulaAuditLineIntersectsBox,
  frozenCellClassNames,
  frozenCellStickyCss,
  frozenColumnStickyLeft,
  frozenColumnStickyRight,
  frozenColumnStickyStyle,
  frozenPaneColumnCount,
  frozenPaneRowCount,
  frozenRowStickyStyle,
  frozenRowStickyTop,
  handlePageLayoutHeaderFooterDblClick,
  handlePageLayoutWhiteSpaceDblClick,
  isFormatPainterRepeatTrigger,
  isFrozenPaneCell,
  measureRichLine,
  mergeRangeContentCellCount,
  mergeRangeNeedsContentWarning,
  mergeSelection,
  mergeSelectionForSheetGroup,
  mergeSheetRange,
  normalizeCellDateValueForNumberFormat,
  normalizeRichTextLineEndings,
  numberFormatForRibbon,
  numberFormatWithDecimals,
  pageBreakAxisPlan,
  pageBreakAxisSegment,
  pageBreakDropState,
  pageBreakEdgeSnapIndex,
  pageBreakGuideAttributes,
  pageBreakGuideCoincidesWithPrintAreaEdge,
  pageBreakIndexFromClientPosition,
  pageBreakIndexFromPixel,
  pageBreakPreviewPageNumber,
  pageBreakPreviewPlan,
  pageBreakPreviewPrintablePageSize,
  pageLayoutHeaderFooterHtml,
  pageLayoutPaperMetrics,
  paintCanvasCell,
  paintCanvasCellText,
  paintCanvasCellTextLayout,
  paintCanvasCells,
  paintCanvasGridLines,
  paintCanvasQuickAnalysisDataBar,
  paintCanvasQuickAnalysisIcon,
  paintCanvasRichCellText,
  paintCanvasSparkline,
  paintCanvasVerticalCellText,
  paintSheetCanvas,
  pixelRectsIntersect,
  preserveCellRichTextForWholeCellCssPatch,
  printAreaDragHandlesHtml,
  printAreaRangeFromClientPosition,
  printablePageSizePixels,
  rangeIntersects,
  rangeIntersectsRenderBands,
  rangePixelRect,
  rangesIntersect,
  rangesOverlap,
  removeIntersectingMerges,
  renderBandsForVisibleRange,
  renderCanvasLayer,
  renderFormulaAuditOverlay,
  renderFrozenPaneBackgroundLayer,
  renderFrozenPaneBorderLayer,
  renderIndexesWithFrozenBand,
  renderPageBreakGuides,
  renderPageLayoutGuides,
  renderSheet,
  renderSheetBackgroundLayer,
  renderSheetRulerTicks,
  renderSheetRulers,
  renderViewPaneGuides,
  repaintActiveSheetCanvas,
  restoreCellsFromHistory,
  restoreDimensionsFromHistory,
  richLinePlainText,
  richSegmentCss,
  richTextRunsForDisplay,
  roundCssPx,
  sameRange,
  setFormatPainterActive,
  shapeTextPatchFromCss,
  sheetAxisItemSize,
  sheetIndexFromPixelForPrintAreaDrag,
  sheetRulerUnitPixels,
  showMergeContentWarning,
  splitRichRunsIntoHardLines,
  startFormatPainter,
  startPageBreakDrag,
  startPrintAreaDrag,
  textDecorationHasToken,
  textDecorationWithToken,
  toggleFormatPainter,
  toggleMergeCenterSelection,
  toggleStyleForSelection,
  toggleTextDecorationForSelection,
  toggleWrapTextForSelection,
  unmergeSelection,
  unmergeSelectionForSheetGroup,
  updatePageBreakDrag,
  updatePrintAreaDrag,
  wrapCanvasRichText,
};
