// Feature module; shared bindings are imported explicitly.
import {
  $,
  $formulaBar,
  $horizontalScrollbar,
  $sheetHost,
  CELL_BORDER_SIDES,
  INDEXED_COLORS,
  createIcons,
  icons,
  state,
} from "../../app-core.js";
import {
  closeDataCommandMenu,
} from "../data/query-transforms.js";
import {
  closeDataDialog,
} from "../drawing/charts-and-insertion.js";
import {
  normalizedHeaderFooterSettings,
} from "../drawing/shape-rendering.js";
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
  cloneCellForClipboard,
  isElementImeCompositionEvent,
} from "../editing/keyboard-and-paste.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  isSingleCellRange,
  materializeSelectionRange,
  selectableCellPoint,
  updateSelectionUi,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  restoreCellEditSelection,
  updateFormulaBarForCell,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  applyDeleteOperation,
  applyInsertOperation,
  captureSheetsForHistory,
  cloneJsonLikeValue,
  internCellsInRange,
  mutableBordersObject,
  mutableCssObject,
  normalizeSheetModelAfterStructureChange,
  pushHistory,
  storeCellModelForSheet,
  writeClonedCellContent,
} from "../editing/structure-and-history.js";
import {
  ensureCellModel,
  forEachCellInRange,
  removeIntersectingMerges,
  roundCssPx,
} from "../rendering/grid-and-canvas.js";
import {
  clearCellsInRange,
  closeCellContextMenu,
  isRangeInsideRange,
  positionPopupMenu,
} from "../ui/annotations-and-cell-format.js";
import {
  accessKeyLabelHtml,
  cellKey,
  columnName,
  columnNumber,
  columnWidthToPixels,
  decodeRange,
  escapeAttr,
  escapeHtml,
  rowHeightToPixels,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  closeHomeCommandMenu,
  consumeCommandMenuKey,
  enhanceCommandMenuAccessKeys,
} from "../ui/printing-and-home-tools.js";
import {
  activeCellEditElementForFormatting,
  closeFileSideMenu,
  ensureSheetSize,
} from "../ui/ribbon-and-links.js";
import {
  clampZoom,
  closeSheetTabContextMenu,
  displayScaleForZoom,
  scaleSheetDisplay,
  setActiveSheetView,
} from "../ui/sheet-protection-and-translation.js";
import {
  activateSheetView,
  captureWorkbookMacroState,
  captureWorkbookViewState,
  commitWorkbookViewChange,
  createSheetView,
  deleteActiveSheetView,
  duplicateActiveSheetView,
  ensureWorkbookWindows,
  exitActiveSheetView,
  freezePanesMenuKind,
  keepActiveSheetView,
  openSheetViewRenameDialog,
  setFreezePanes,
  sheetHeadingsVisible,
  syncViewRibbonButtons,
  syncWorkbookTitle,
  updateActiveSheetView,
  withWorkbookWindowDisplayNames,
  workbookWindowBaseName,
  workbookWindowDisplayName,
} from "../ui/views-and-formula-tools.js";
import {
  buildFormulaEngine,
  fontFamilyForExcel,
  renderWorkbook,
  rgbToHex,
  verticalAlignToFlex,
} from "./formulas-and-sheet-tabs.js";
import {
  normalizeCustomTopLeftCell,
  normalizeCustomViewPoint,
  normalizeCustomViewRange,
  normalizeCustomViewScroll,
  normalizePaneState,
  normalizeSheetView,
  normalizeWorkbookCustomView,
  normalizeWorkbookWindowView,
  normalizedSheetShowRuler,
  normalizedSheetViewMode,
  scheduleWorkbookPersistence,
} from "./import-model.js";
import {
  activeSheet,
  updateHorizontalScrollbar,
} from "./saving-and-export.js";
import {
  restoreScrollRecordForView,
  scrollRecordFromTopLeftCell,
  topLeftCellFromScrollRecord,
} from "./tables-and-validation.js";
import {
  clearVbaRangeTarget,
  compareVbaExpressionValues,
  evaluateVbaBuiltInFunctionExpression,
  evaluateVbaWorksheetFunctionExpression,
  insertPageBreakAt,
  isVbaNoneValue,
  normalizedPageLayoutSettings,
  openMacroRecordDialog,
  openMacroSecurityDialog,
  openPageSetupDialog,
  readVbaRangeValue,
  removePageBreakAt,
  resetAllPageBreaks,
  selectVbaRangeTarget,
  selectedMacroFromDialog,
  splitVbaArithmeticExpression,
  splitVbaComparisonExpression,
  splitVbaTopLevelWord,
  stopMacroRecording,
  toggleMacroRelativeReferences,
  trimVbaOuterParentheses,
  vbaColorToCss,
  vbaComparableValue,
  vbaConstantValue,
  vbaHorizontalAlignment,
  vbaTruthy,
  vbaUnderlineTextDecoration,
  vbaVerticalAlignment,
  writeVbaFormulaR1C1ToRange,
  writeVbaValueToRange,
} from "./vba-and-printing.js";


function normalizedViewSideBySidePair(visibleRecords = null) {
  if (!state.model?.viewSideBySide) return null;
  const visible = Array.isArray(visibleRecords) ? visibleRecords : workbookWindowRecords().filter((view) => !view.hidden);
  if (visible.length < 2) return null;
  const visibleIds = new Set(visible.map((view) => view.id));
  const pair = state.model.viewSideBySide;
  let left = visibleIds.has(pair.left) ? pair.left : "";
  let right = visibleIds.has(pair.right) && pair.right !== left ? pair.right : "";
  const activeId = state.model.activeWindowId || "main";
  if (visibleIds.has(activeId) && left !== activeId && right !== activeId) {
    left = activeId;
    right = visible.find((view) => view.id !== left && (view.id === pair.right || view.id === pair.left))?.id || "";
  }
  if (!left) left = visibleIds.has(activeId) ? activeId : visible[0]?.id || "";
  if (!right || right === left) right = visible.find((view) => view.id !== left)?.id || "";
  return left && right && left !== right ? { left, right } : null;
}



function repairViewSideBySideState(visibleRecords = null) {
  if (!state.model?.viewSideBySide) return null;
  const pair = normalizedViewSideBySidePair(visibleRecords);
  if (!pair) {
    state.model.viewSideBySide = null;
    state.model.synchronousScrolling = false;
    return null;
  }
  state.model.viewSideBySide = pair;
  return pair;
}



function workbookWindowChipHtml(view, index, activeId, sideBySide = null) {
  const isActive = view.id === activeId;
  const isCompared = sideBySide && (view.id === sideBySide.left || view.id === sideBySide.right);
  const label = workbookWindowDisplayName(view);
  const meta = workbookWindowMenuDescription(view);
  const closeLabel = `${label} を閉じる`;
  return `<div class="workbook-window-chip-wrap${isActive ? " is-active" : ""}${isCompared ? " is-compared" : ""}">
    <button class="workbook-window-chip${isActive ? " is-active" : ""}${isCompared ? " is-compared" : ""}" type="button" role="tab" aria-selected="${isActive ? "true" : "false"}" data-workbook-window-id="${escapeAttr(view.id)}" title="${escapeAttr(meta)}">
      <span class="workbook-window-index">${index + 1}</span>
      <span class="workbook-window-chip-name">${escapeHtml(label)}</span>
    </button>
    <button class="workbook-window-close" type="button" data-workbook-window-close="${escapeAttr(view.id)}" aria-label="${escapeAttr(closeLabel)}" title="${escapeAttr(closeLabel)}">×</button>
  </div>`;
}



function workbookWindowArrangementPreviewHtml(visibleRecords, arrangement, activeId, sideBySide = null) {
  const recordsById = new Map(visibleRecords.map((view) => [view.id, view]));
  const slots = (arrangement.windows || []).map((slot) => {
    const view = recordsById.get(slot.id);
    if (!view) return "";
    const isActive = view.id === activeId;
    const isCompared = sideBySide && (view.id === sideBySide.left || view.id === sideBySide.right);
    return `<button class="workbook-window-preview-slot${isActive ? " is-active" : ""}${isCompared ? " is-compared" : ""}" type="button" data-workbook-window-id="${escapeAttr(view.id)}" style="${workbookWindowSlotStyle(slot)}" title="${escapeAttr(workbookWindowMenuDescription(view))}">
      <span>${escapeHtml(workbookWindowDisplayName(view))}</span>
    </button>`;
  }).join("");
  return `<div class="workbook-window-preview is-${escapeAttr(arrangement.mode || "tiled")}" aria-label="${escapeAttr(workbookWindowArrangementModeLabel(arrangement.mode || "tiled"))}">${slots}</div>`;
}



function workbookWindowSlotStyle(slot) {
  const left = Math.max(0, Math.min(100, Number(slot.left) || 0));
  const top = Math.max(0, Math.min(100, Number(slot.top) || 0));
  const width = Math.max(8, Math.min(100, Number(slot.width) || 50));
  const height = Math.max(14, Math.min(100, Number(slot.height) || 50));
  const zIndex = Math.max(1, Math.round(Number(slot.zIndex) || 1));
  return `left:${roundCssPx(left)}%;top:${roundCssPx(top)}%;width:${roundCssPx(width)}%;height:${roundCssPx(height)}%;z-index:${zIndex}`;
}



function workbookWindowArrangementModeLabel(mode) {
  const labels = {
    tiled: "並べて表示",
    horizontal: "上下に並べて表示",
    vertical: "左右に並べて表示",
    cascade: "重ねて表示",
    "side-by-side": "並べて比較",
  };
  return labels[mode] || labels.tiled;
}



function handleWorkbookWindowBarClick(event) {
  const button = event.currentTarget;
  const id = button?.dataset?.workbookWindowId || "";
  if (!id || id === (state.model?.activeWindowId || "main")) return;
  event.preventDefault();
  switchWorkbookWindow(id);
}



function handleWorkbookWindowCloseClick(event) {
  event.preventDefault();
  event.stopPropagation();
  const id = event.currentTarget?.dataset?.workbookWindowClose || "";
  closeWorkbookWindowView(id);
}



function workbookWindowRecords() {
  if (!state.model) return [];
  const sourceName = workbookWindowBaseName();
  const activeId = state.model.activeWindowId || "main";
  const mainStoredView = normalizeWorkbookWindowView({
    id: "main",
    name: sourceName,
    ...(state.model.mainWindowView && typeof state.model.mainWindowView === "object" ? state.model.mainWindowView : {}),
    hidden: state.model.mainWindowHidden === true,
  });
  const mainLiveView = activeId === "main"
    ? {
      ...mainStoredView,
      ...currentWorkbookWindowSnapshot(),
      id: "main",
      name: sourceName,
      hidden: state.model.mainWindowHidden === true,
    }
    : mainStoredView;
  const liveSnapshot = activeId !== "main" ? currentWorkbookWindowSnapshot() : null;
  return withWorkbookWindowDisplayNames([
    mainLiveView,
    ...(state.model.windowViews || []).map(normalizeWorkbookWindowView).map((view, index) => ({
      ...view,
      ...(view.id === activeId && liveSnapshot ? liveSnapshot : {}),
      id: view.id,
      name: view.name || `${sourceName}:${index + 2}`,
      hidden: view.hidden === true,
    })),
  ]);
}



function storedWorkbookWindowRecordById(id = "main") {
  if (!state.model) return null;
  const sourceName = workbookWindowBaseName();
  if (id === "main") {
    return normalizeWorkbookWindowView({
      id: "main",
      name: sourceName,
      ...(state.model.mainWindowView && typeof state.model.mainWindowView === "object" ? state.model.mainWindowView : {}),
      hidden: state.model.mainWindowHidden === true,
    });
  }
  const view = state.model.windowViews?.find((item) => item.id === id);
  if (!view) return null;
  return normalizeWorkbookWindowView({
    ...view,
    name: view.name || sourceName,
    hidden: view.hidden === true,
  });
}



function storedVisibleWorkbookWindowRecords() {
  return storedWorkbookWindowRecords().filter((view) => view && !view.hidden);
}



function storedWorkbookWindowRecords() {
  if (!state.model) return [];
  return [
    storedWorkbookWindowRecordById("main"),
    ...(state.model.windowViews || []).map((view) => storedWorkbookWindowRecordById(view.id)),
  ].filter(Boolean);
}



function currentWorkbookWindowSnapshot() {
  const host = $sheetHost[0];
  const sheet = activeSheet();
  const viewMode = normalizedSheetViewMode(sheet?.viewMode);
  const selectionRanges = activeSelectionRanges().map((range) => ({ ...range }));
  const selectionRange = selectionRanges[0] || activeSelectionRange();
  const activeScroll = {
    top: Math.max(0, Math.round(host?.scrollTop || 0)),
    left: Math.max(0, Math.round(host?.scrollLeft || 0)),
  };
  return {
    sheetIndex: state.activeSheetIndex,
    viewMode,
    zoomScale: clampZoom(sheet?.zoomScale || 100),
    showGridLines: sheet?.showGridLines !== false,
    showHeadings: sheetHeadingsVisible(sheet),
    showRuler: normalizedSheetShowRuler(sheet?.showRuler, viewMode),
    showObjects: sheet?.showObjects !== false,
    showOutlineSymbols: sheet?.showOutlineSymbols !== false,
    showFormulas: sheet?.showFormulas === true,
    showZeros: sheet?.showZeros !== false,
    showWhiteSpace: sheet?.showWhiteSpace !== false,
    rightToLeft: sheet?.rightToLeft === true,
    freezePanes: cloneJsonLikeValue(sheet?.freezePanes),
    splitPane: cloneJsonLikeValue(sheet?.splitPane),
    showHorizontalScroll: activeWorkbookWindowShowOption("showHorizontalScroll"),
    showVerticalScroll: activeWorkbookWindowShowOption("showVerticalScroll"),
    showSheetTabs: activeWorkbookWindowShowOption("showSheetTabs"),
    scrollTop: activeScroll.top,
    scrollLeft: activeScroll.left,
    topLeftCell: topLeftCellFromScrollRecord(sheet, activeScroll),
    selectedCell: state.selected?.sheetIndex === state.activeSheetIndex ? { ...state.selected } : null,
    selectionRange: selectionRange ? { ...selectionRange } : null,
    selectionRanges,
  };
}



function activeWorkbookWindowShowOption(optionName) {
  if (!state.model) return true;
  const activeId = state.model.activeWindowId || "main";
  const source = activeId === "main"
    ? state.model.mainWindowView || state.model.workbookView || {}
    : state.model.windowViews?.find((view) => view.id === activeId) || {};
  return source?.[optionName] !== false;
}



function restoreWorkbookWindowSelection(record) {
  const sheet = activeSheet();
  const sheetIndex = state.activeSheetIndex;
  const ranges = (Array.isArray(record?.selectionRanges) && record.selectionRanges.length ? record.selectionRanges : [record?.selectionRange])
    .map((range) => normalizeCustomViewRange(range, sheetIndex))
    .filter((range) => range && range.sheetIndex === sheetIndex)
    .map((range) => ({
      ...range,
      top: Math.min(Math.max(1, range.top), sheet.rowCount),
      left: Math.min(Math.max(1, range.left), sheet.colCount),
      bottom: Math.min(Math.max(1, range.bottom), sheet.rowCount),
      right: Math.min(Math.max(1, range.right), sheet.colCount),
    }))
    .map((range) => materializeSelectionRange(range))
    .filter(Boolean);
  const firstRange = ranges[0] || { sheetIndex, top: 1, left: 1, bottom: 1, right: 1 };
  const selected = normalizeCustomViewPoint(record?.selectedCell, sheetIndex) || {
    sheetIndex,
    row: firstRange.top,
    col: firstRange.left,
  };
  const point = selectableCellPoint(
    Math.min(Math.max(1, selected.row), sheet.rowCount),
    Math.min(Math.max(1, selected.col), sheet.colCount),
  );
  state.selected = point;
  state.selectionRange = firstRange;
  state.selectionRanges = ranges.length ? ranges : [{ ...firstRange }];
  state.selectionStart = point;
  state.selectionAnchor = point;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  state.formulaEdit = null;
}



function saveActiveWindowViewState() {
  if (!state.model) return;
  const activeId = state.model.activeWindowId || "main";
  const snapshot = currentWorkbookWindowSnapshot();
  if (activeId === "main") {
    state.model.mainWindowView = normalizeWorkbookWindowView({
      id: "main",
      name: state.model.sourceName || "Book1",
      ...(state.model.mainWindowView || {}),
      ...snapshot,
      hidden: state.model.mainWindowHidden === true,
    });
    return;
  }
  const view = state.model.windowViews?.find((item) => item.id === activeId);
  if (!view) return;
  Object.assign(view, snapshot);
}



function updateActiveWorkbookWindowScrollState() {
  if (!state.model) return;
  const host = $sheetHost[0];
  if (!host) return;
  const activeId = state.model.activeWindowId || "main";
  const target = mutableWorkbookWindowRecordById(activeId);
  if (!target) return;
  updateWorkbookWindowScrollFields(target, activeSheet(), host.scrollTop, host.scrollLeft);
}



function syncPairedWorkbookWindowScrollState() {
  if (!state.model?.synchronousScrolling || !state.model.viewSideBySide) return;
  const activeId = state.model.activeWindowId || "main";
  const pair = state.model.viewSideBySide;
  const targetId = activeId === pair.left ? pair.right : activeId === pair.right ? pair.left : "";
  if (!targetId) return;
  const source = workbookWindowRecords().find((view) => view.id === activeId);
  const target = mutableWorkbookWindowRecordById(targetId);
  if (!source || !target) return;
  const targetSheet = state.model.sheets[target.sheetIndex] || activeSheet();
  updateWorkbookWindowScrollFields(target, targetSheet, source.scrollTop, source.scrollLeft);
}



function mutableWorkbookWindowRecordById(id) {
  if (!state.model) return null;
  const recordId = id || state.model.activeWindowId || "main";
  if (recordId === "main") {
    state.model.mainWindowView ||= normalizeWorkbookWindowView({ id: "main", name: state.model.sourceName || "Book1" });
    return state.model.mainWindowView;
  }
  state.model.windowViews = Array.isArray(state.model.windowViews) ? state.model.windowViews : [];
  return state.model.windowViews.find((view) => view.id === recordId) || null;
}



function updateWorkbookWindowScrollFields(record, sheet, scrollTop, scrollLeft) {
  if (!record) return;
  const top = Math.max(0, Math.round(Number(scrollTop) || 0));
  const left = Math.max(0, Math.round(Number(scrollLeft) || 0));
  record.scrollTop = top;
  record.scrollLeft = left;
  record.topLeftCell = topLeftCellFromScrollRecord(sheet, { top, left, scrollTop: top, scrollLeft: left });
  record.scrollUnit = "pixel";
}



function createWorkbookWindowView() {
  saveActiveWindowViewState();
  const workbookViewBefore = captureWorkbookViewState();
  state.model.windowViews = Array.isArray(state.model.windowViews) ? state.model.windowViews : [];
  const id = uniqueWorksheetObjectId("window");
  const name = `${state.model.sourceName || "Book1"}:${state.model.windowViews.length + 2}`;
  state.model.windowViews.push({
    id,
    name,
    ...currentWorkbookWindowSnapshot(),
    hidden: false,
  });
  state.model.activeWindowId = id;
  state.model.windowArrangement = computeWorkbookWindowArrangement(state.model.windowArrangement?.mode || "tiled");
  syncWorkbookTitle();
  commitWorkbookViewChange(workbookViewBefore, `新しいウィンドウを作成しました：${name}`, "新しいウィンドウ");
}



function openArrangeAllDialog() {
  ensureWorkbookWindows();
  closeDataDialog({ keepState: true });
  const currentMode = state.model?.windowArrangement?.mode || "tiled";
  const checked = (mode) => (currentMode === mode ? " checked" : "");
  const radio = (mode, label) => `<label class="data-dialog-check"${commandMenuAccessKeyAttr(label)}><input type="radio" name="viewArrangeMode" value="${escapeAttr(mode)}"${checked(mode)} />${accessKeyLabelHtml(label)}</label>`;
  const activeWorkbookLabel = "作業中のブックのウィンドウを整列する(W)";
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog view-arrange-dialog" role="dialog" aria-modal="true" aria-labelledby="arrangeWindowTitle" data-data-dialog="view-arrange-window">
        <div class="excel-mini-titlebar" id="arrangeWindowTitle">ウィンドウの整列</div>
        <div class="excel-mini-body data-dialog-body">
          <fieldset class="data-dialog-fieldset view-arrange-options">
            <legend>整列方法</legend>
            ${radio("tiled", "並べて表示(T)")}
            ${radio("horizontal", "上下に並べて表示(H)")}
            ${radio("vertical", "左右に並べて表示(V)")}
            ${radio("cascade", "重ねて表示(C)")}
          </fieldset>
          <label class="data-dialog-check"${commandMenuAccessKeyAttr(activeWorkbookLabel)}>
            <input type="checkbox" data-data-field="activeWorkbookOnly" checked disabled />${accessKeyLabelHtml(activeWorkbookLabel)}
          </label>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="view-arrange-ok">OK</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "view-arrange-window" };
  overlay.find("input[name='viewArrangeMode']:checked").trigger("focus");
}



function arrangeWorkbookWindowViews(mode = "tiled") {
  const workbookViewBefore = captureWorkbookViewState();
  const normalizedMode = ["tiled", "horizontal", "vertical", "cascade"].includes(mode) ? mode : "tiled";
  state.model.windowArrangement = computeWorkbookWindowArrangement(normalizedMode);
  const labels = { tiled: "並べて表示", horizontal: "上下に並べて表示", vertical: "左右に並べて表示", cascade: "重ねて表示" };
  commitWorkbookViewChange(workbookViewBefore, `${state.model.windowArrangement.count} 個のウィンドウを整列しました：${labels[normalizedMode]}`, "ウィンドウの整列");
}



function computeWorkbookWindowArrangement(mode = "tiled") {
  const normalizedMode = ["tiled", "horizontal", "vertical", "cascade", "side-by-side"].includes(mode) ? mode : "tiled";
  const visibleWindows = workbookWindowRecords().filter((view) => !view.hidden);
  let windows = visibleWindows;
  if (normalizedMode === "side-by-side") {
    const pair = normalizedViewSideBySidePair(visibleWindows);
    const byId = new Map(visibleWindows.map((view) => [view.id, view]));
    windows = pair ? [pair.left, pair.right].map((id) => byId.get(id)).filter(Boolean) : [];
    if (windows.length < 2) windows = visibleWindows.slice(0, 2);
  }
  const count = Math.max(1, windows.length);
  const columns = normalizedMode === "vertical" || normalizedMode === "side-by-side"
    ? count
    : normalizedMode === "horizontal"
      ? 1
      : Math.ceil(Math.sqrt(count));
  const rows = normalizedMode === "horizontal" ? count : Math.ceil(count / columns);
  const slots = windows.map((view, index) => {
    if (normalizedMode === "cascade") {
      const offset = Math.min(24, 8 + index * 4);
      return {
        id: view.id,
        name: workbookWindowDisplayName(view),
        left: Math.min(45, index * offset),
        top: Math.min(45, index * offset),
        width: Math.max(45, 92 - index * 4),
        height: Math.max(45, 88 - index * 4),
        zIndex: index + 1,
      };
    }
    const row = Math.floor(index / columns);
    const col = index % columns;
    return {
      id: view.id,
      name: workbookWindowDisplayName(view),
      left: Math.round((col * 10000) / columns) / 100,
      top: Math.round((row * 10000) / rows) / 100,
      width: Math.round((10000 / columns)) / 100,
      height: Math.round((10000 / rows)) / 100,
      zIndex: index + 1,
    };
  });
  return { mode: normalizedMode, count, rows, columns, windows: slots };
}



function hideWorkbookWindow() {
  const records = ensureWorkbookWindows();
  const visible = records.filter((view) => !view.hidden);
  if (visible.length <= 1) {
    showMessageBox("表示中の最後のウィンドウを非表示にすることはできません。", { title: "Cht WebSheet" });
    return;
  }
  saveActiveWindowViewState();
  const workbookViewBefore = captureWorkbookViewState();
  const activeId = state.model.activeWindowId || "main";
  if (activeId === "main") state.model.mainWindowHidden = true;
  else {
    const view = state.model.windowViews.find((item) => item.id === activeId);
    if (view) view.hidden = true;
  }
  const next = workbookWindowRecords().find((view) => !view.hidden);
  if (next) switchWorkbookWindow(next.id, { silent: true });
  const visibleAfter = workbookWindowRecords().filter((view) => !view.hidden);
  const pair = repairViewSideBySideState(visibleAfter);
  const mode = pair ? "side-by-side" : state.model.windowArrangement?.mode === "side-by-side" ? "tiled" : state.model.windowArrangement?.mode || "tiled";
  state.model.windowArrangement = computeWorkbookWindowArrangement(mode);
  commitWorkbookViewChange(workbookViewBefore, "ウィンドウを非表示にしました。", "ウィンドウを表示しない");
}



function closeWorkbookWindowView(id) {
  if (!state.model || !id) return;
  const records = ensureWorkbookWindows();
  const visible = records.filter((view) => !view.hidden);
  const target = visible.find((view) => view.id === id);
  if (!target) return;
  if (visible.length <= 1) {
    showMessageBox("表示中の最後のウィンドウを閉じることはできません。", { title: "Cht WebSheet" });
    return;
  }
  if (id === "main") {
    hideWorkbookWindow();
    return;
  }
  saveActiveWindowViewState();
  const workbookViewBefore = captureWorkbookViewState();
  const activeId = state.model.activeWindowId || "main";
  const targetName = workbookWindowDisplayName(target);
  state.model.windowViews = (state.model.windowViews || []).filter((view) => view.id !== id);
  const visibleAfter = workbookWindowRecords().filter((view) => !view.hidden);
  const next = visibleAfter.find((view) => view.id !== id && view.id === activeId)
    || visibleAfter.find((view) => view.id !== id)
    || visibleAfter[0];
  if (activeId === id && next) {
    switchWorkbookWindow(next.id, { silent: true });
  }
  const pair = repairViewSideBySideState(visibleAfter);
  const currentMode = state.model.windowArrangement?.mode || "tiled";
  state.model.windowArrangement = computeWorkbookWindowArrangement(pair ? "side-by-side" : currentMode === "side-by-side" ? "tiled" : currentMode);
  syncWorkbookTitle();
  commitWorkbookViewChange(workbookViewBefore, `ウィンドウを閉じました：${targetName}`, "ウィンドウを閉じる");
}



function openUnhideWorkbookWindowDialog() {
  const hidden = ensureWorkbookWindows().filter((view) => view.hidden);
  if (!hidden.length) {
    showMessageBox("再表示できるウィンドウはありません。", { title: "Cht WebSheet" });
    return;
  }
  closeDataDialog({ keepState: true });
  const options = hidden.map((view, index) => `<option value="${escapeAttr(view.id)}"${index === 0 ? " selected" : ""}>${escapeHtml(workbookWindowDisplayName(view))}</option>`).join("");
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog view-window-dialog" role="dialog" aria-modal="true" aria-labelledby="unhideWindowTitle" data-data-dialog="view-unhide-window">
        <div class="excel-mini-titlebar" id="unhideWindowTitle">再表示</div>
        <div class="excel-mini-body data-dialog-body">
          <label class="excel-dialog-field"><span class="excel-dialog-caption">再表示するウィンドウ(U):</span>
            <select size="${Math.min(6, Math.max(2, hidden.length))}" data-data-field="windowId">${options}</select>
          </label>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="view-window-unhide-ok">OK</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "view-unhide-window" };
  overlay.find("[data-data-field='windowId']").trigger("focus");
}



function unhideWorkbookWindow(id) {
  if (!state.model || !id) return;
  const workbookViewBefore = captureWorkbookViewState();
  if (id === "main") state.model.mainWindowHidden = false;
  else {
    const view = state.model.windowViews?.find((item) => item.id === id);
    if (view) view.hidden = false;
  }
  switchWorkbookWindow(id, { silent: true });
  repairViewSideBySideState();
  state.model.windowArrangement = computeWorkbookWindowArrangement(state.model.viewSideBySide ? "side-by-side" : state.model.windowArrangement?.mode || "tiled");
  commitWorkbookViewChange(workbookViewBefore, "ウィンドウを再表示しました。", "ウィンドウの再表示");
}



function toggleViewSideBySide() {
  const visible = ensureWorkbookWindows().filter((view) => !view.hidden);
  if (state.model.viewSideBySide) {
    const workbookViewBefore = captureWorkbookViewState();
    state.model.viewSideBySide = null;
    state.model.synchronousScrolling = false;
    state.model.windowArrangement = computeWorkbookWindowArrangement(state.model.windowArrangement?.mode || "tiled");
    commitWorkbookViewChange(workbookViewBefore, "並べて比較を解除しました。", "並べて比較");
    return;
  }
  if (visible.length < 2) {
    showMessageBox("並べて比較するには、少なくとも 2 つのウィンドウを開いてください。", { title: "Cht WebSheet" });
    return;
  }
  const activeId = state.model.activeWindowId || "main";
  const candidates = visible.filter((view) => view.id !== activeId);
  if (candidates.length > 1) {
    openSideBySideWindowDialog(candidates);
    return;
  }
  enableViewSideBySide(activeId, candidates[0]?.id || visible[0]?.id || "main");
}



function openSideBySideWindowDialog(candidates) {
  closeDataDialog({ keepState: true });
  const options = candidates
    .map((view, index) => `<option value="${escapeAttr(view.id)}"${index === 0 ? " selected" : ""}>${escapeHtml(workbookWindowDisplayName(view))}</option>`)
    .join("");
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog view-window-dialog" role="dialog" aria-modal="true" aria-labelledby="sideBySideWindowTitle" data-data-dialog="view-side-by-side-window">
        <div class="excel-mini-titlebar" id="sideBySideWindowTitle">並べて比較</div>
        <div class="excel-mini-body data-dialog-body">
          <label class="excel-dialog-field"><span class="excel-dialog-caption">比較するウィンドウ(C):</span>
            <select size="${Math.min(6, Math.max(2, candidates.length))}" data-data-field="sideBySideWindowId">${options}</select>
          </label>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="view-side-by-side-ok">OK</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "view-side-by-side-window" };
  overlay.find("[data-data-field='sideBySideWindowId']").trigger("focus");
}



function enableViewSideBySide(activeId, compareId) {
  const visibleIds = new Set(ensureWorkbookWindows().filter((view) => !view.hidden).map((view) => view.id));
  const left = visibleIds.has(activeId) ? activeId : "main";
  const right = visibleIds.has(compareId) && compareId !== left
    ? compareId
    : [...visibleIds].find((id) => id !== left) || left;
  if (!right || right === left) {
    showMessageBox("比較するウィンドウを選択できません。", { title: "Cht WebSheet" });
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  state.model.viewSideBySide = { left, right };
  state.model.synchronousScrolling = false;
  state.model.windowArrangement = computeWorkbookWindowArrangement("side-by-side");
  commitWorkbookViewChange(workbookViewBefore, "ウィンドウを並べて比較します。", "並べて比較");
}



function toggleSynchronousScrolling() {
  const pair = repairViewSideBySideState();
  if (!pair) {
    showMessageBox("同時にスクロールするには、先にウィンドウを並べて比較してください。", { title: "Cht WebSheet" });
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  state.model.synchronousScrolling = state.model.synchronousScrolling !== true;
  if (state.model.synchronousScrolling) {
    updateActiveWorkbookWindowScrollState();
    syncPairedWorkbookWindowScrollState();
  }
  commitWorkbookViewChange(
    workbookViewBefore,
    state.model.synchronousScrolling ? "同時にスクロールを有効にしました。" : "同時にスクロールを無効にしました。",
    "同時にスクロール",
  );
}



function resetWindowPosition() {
  const visible = ensureWorkbookWindows().filter((view) => !view.hidden);
  if (!state.model.viewSideBySide) {
    showMessageBox("ウィンドウ位置をリセットするには、先に並べて比較を有効にしてください。", { title: "Cht WebSheet" });
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const pair = repairViewSideBySideState(visible);
  if (!pair) {
    showMessageBox("比較する 2 つのウィンドウを選択できません。", { title: "Cht WebSheet" });
    return;
  }
  state.model.windowArrangement = computeWorkbookWindowArrangement("side-by-side");
  commitWorkbookViewChange(workbookViewBefore, "ウィンドウの位置をリセットしました。", "ウィンドウ位置のリセット");
}



function switchWorkbookWindow(id, options = {}) {
  if (!state.model || !id) return false;
  const recordHistory = !options.silent && options.recordHistory !== false;
  const selectionBefore = recordHistory ? activeSelectionRange() : null;
  const selectionBeforeRanges = recordHistory ? activeSelectionRanges().map((range) => ({ ...range })) : [];
  const workbookViewBefore = recordHistory ? captureWorkbookViewState() : null;
  saveActiveWindowViewState();
  const record = workbookWindowRecords().find((view) => view.id === id && !view.hidden);
  if (!record) return false;
  const previousWindowId = state.model.activeWindowId || "main";
  if (previousWindowId === id) {
    if (!options.silent) setStatus(`ウィンドウを切り替えました：${workbookWindowDisplayName(record)}`);
    return true;
  }
  state.model.activeWindowId = id;
  if (state.model.sheets[record.sheetIndex]) state.activeSheetIndex = record.sheetIndex;
  const sheet = activeSheet();
  if (sheet) {
    applyWorkbookWindowSheetView(sheet, record);
    restoreWorkbookWindowSelection(record);
  }
  restoreWorkbookWindowScrollStable(record);
  renderWorkbook();
  restoreWorkbookWindowScrollStable(record);
  updateSelectionUi();
  updateFormulaBarForSelection();
  syncViewRibbonButtons();
  if (recordHistory && workbookViewBefore) {
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      workbookViewBefore,
      workbookViewAfter: captureWorkbookViewState(),
      selectionBefore,
      selectionAfter: activeSelectionRange(),
      selectionBeforeRanges,
      selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
      label: `ウィンドウの切り替え：${workbookWindowDisplayName(record)}`,
    });
  }
  scheduleWorkbookPersistence();
  if (!options.silent) setStatus(`ウィンドウを切り替えました：${workbookWindowDisplayName(record)}`);
  return true;
}



function applyWorkbookWindowSheetView(sheet, record) {
  if (!sheet || !record) return;
  sheet.viewMode = normalizedSheetViewMode(record.viewMode || sheet.viewMode);
  if (record.zoomScale && record.zoomScale !== sheet.zoomScale) applySheetZoomWithoutHistory(sheet, record.zoomScale);
  sheet.showGridLines = record.showGridLines !== false;
  sheet.showHeadings = record.showHeadings !== false;
  sheet.showRuler = normalizedSheetShowRuler(record.showRuler, sheet.viewMode);
  sheet.showObjects = record.showObjects !== false;
  sheet.showOutlineSymbols = record.showOutlineSymbols !== false;
  sheet.showFormulas = record.showFormulas === true;
  sheet.showZeros = record.showZeros !== false;
  sheet.showWhiteSpace = record.showWhiteSpace !== false;
  sheet.rightToLeft = record.rightToLeft === true;
  state.showFormulas = sheet.showFormulas === true;
  sheet.freezePanes = normalizePaneState(record.freezePanes);
  sheet.splitPane = normalizePaneState(record.splitPane);
}



function applySheetZoomWithoutHistory(sheet, value) {
  if (!sheet) return;
  const nextZoom = clampZoom(value);
  const previousZoom = Number(sheet.zoomScale || 100);
  if (nextZoom === previousZoom) return;
  const previousScale = Number(sheet.displayScale || 1);
  const nextScale = displayScaleForZoom(sheet, nextZoom);
  const ratio = nextScale / previousScale;
  scaleSheetDisplay(sheet, ratio);
  sheet.zoomScale = nextZoom;
  sheet.displayScale = nextScale;
}



function restoreWorkbookWindowScroll(record) {
  const host = $sheetHost[0];
  if (!host || !record) return;
  const cellScroll = record.scrollUnit === "cell" ? scrollRecordFromTopLeftCell(activeSheet(), record.topLeftCell) : null;
  const left = Math.max(0, Math.round(Number(cellScroll?.scrollLeft ?? record.scrollLeft) || 0));
  const top = Math.max(0, Math.round(Number(cellScroll?.scrollTop ?? record.scrollTop) || 0));
  host.scrollTop = top;
  host.scrollLeft = left;
  if ($horizontalScrollbar[0]) $horizontalScrollbar[0].scrollLeft = left;
  updateHorizontalScrollbar();
}



function restoreWorkbookWindowScrollStable(record) {
  if (!record) return;
  restoreWorkbookWindowScroll(record);
  requestAnimationFrame(() => restoreWorkbookWindowScroll(record));
  requestAnimationFrame(() => requestAnimationFrame(() => restoreWorkbookWindowScroll(record)));
}



function openViewFreezeMenu(trigger = null) {
  const sheet = activeSheet();
  const current = freezePanesMenuKind(sheet);
  const hasFreeze = Boolean(current);
  const primaryAction = hasFreeze
    ? viewCommandMenuButtonHtml("unfreeze", "ウィンドウ枠固定の解除(U)", "unlock-keyhole", false, { description: "固定されている行と列を解除します。" })
    : viewCommandMenuButtonHtml("freeze-selection", "ウィンドウ枠の固定(F)", "panel-top", false, { checkable: true, checked: current === "selection", description: "現在の選択範囲に基づいて行と列を固定します。" });
  openViewCommandMenu(trigger, "freeze", "ウィンドウ枠の固定", `
    ${primaryAction}
    ${hasFreeze ? '<div class="home-command-menu-separator" role="separator"></div>' : ""}
    ${viewCommandMenuButtonHtml("freeze-top-row", "先頭行の固定(R)", "panel-top", false, { checkable: true, checked: current === "top-row", description: "ワークシートの先頭行を表示したままスクロールします。" })}
    ${viewCommandMenuButtonHtml("freeze-first-column", "最初の列の固定(C)", "panel-left", false, { checkable: true, checked: current === "first-column", description: "ワークシートの最初の列を表示したままスクロールします。" })}
  `);
}



function openSwitchWindowsMenu(trigger = null) {
  const records = ensureWorkbookWindows().filter((view) => !view.hidden);
  const rows = records.map((view, index) =>
    viewCommandMenuButtonHtml(`switch:${view.id}`, `${index + 1} ${workbookWindowDisplayName(view)}`, "monitor", false, {
      checkable: true,
      checked: view.id === (state.model.activeWindowId || "main"),
      description: workbookWindowMenuDescription(view),
    }),
  ).join("");
  const dialogAction = records.length > 1
    ? `${rows ? '<div class="home-command-menu-separator" role="separator"></div>' : ""}${viewCommandMenuButtonHtml("switch-dialog", "その他のウィンドウ(M)...", "list", false, { description: "表示中のウィンドウを一覧から選択します。" })}`
    : "";
  openViewCommandMenu(trigger, "switch-windows", "ウィンドウの切り替え", rows ? `${rows}${dialogAction}` : `<button type="button" role="menuitem" disabled>ウィンドウがありません</button>`);
}



function openSwitchWindowsDialog() {
  const records = ensureWorkbookWindows().filter((view) => !view.hidden);
  if (!records.length) {
    showMessageBox("切り替えできるウィンドウはありません。", { title: "Cht WebSheet" });
    return;
  }
  const activeId = state.model?.activeWindowId || "main";
  const options = records
    .map((view) => `<option value="${escapeAttr(view.id)}"${view.id === activeId ? " selected" : ""}>${escapeHtml(workbookWindowDisplayName(view))}</option>`)
    .join("");
  closeDataDialog({ keepState: true });
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog view-window-dialog" role="dialog" aria-modal="true" aria-labelledby="switchWindowTitle" data-data-dialog="view-switch-window">
        <div class="excel-mini-titlebar" id="switchWindowTitle">ウィンドウの切り替え</div>
        <div class="excel-mini-body data-dialog-body">
          <label class="excel-dialog-field"${commandMenuAccessKeyAttr("ウィンドウ(W):")}><span class="excel-dialog-caption">${accessKeyLabelHtml("ウィンドウ(W):")}</span>
            <select size="${Math.min(9, Math.max(3, records.length))}" data-data-field="switchWindowId">${options}</select>
          </label>
          <div class="data-dialog-summary">切り替えるウィンドウを選択してください。</div>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="view-switch-window-ok">OK</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "view-switch-window" };
  overlay.find("[data-data-field='switchWindowId']").trigger("focus");
}



function openPageBreakContextMenu(event) {
  if (!state.model) return;
  const sheet = activeSheet();
  if ((sheet?.viewMode || "normal") !== "page-break") return;
  const target = event.currentTarget;
  const axis = target.dataset.pageBreakAxis === "row" ? "row" : "col";
  const index = Math.max(2, Number(target.dataset.pageBreakIndex) || 2);
  const manual = target.dataset.pageBreakManual === "true";
  openPageBreakContextMenuAt(event, axis, index, manual);
}



function openPageBreakContextMenuFromPoint(event) {
  if (!state.model) return;
  const sheet = activeSheet();
  if ((sheet?.viewMode || "normal") !== "page-break") return;
  if (event.target?.closest?.("#viewCommandMenu, #cellContextMenu, #dataCommandMenu, #homeCommandMenu, .excel-dialog-overlay, .sheet-image-selection-frame, .sheet-image, .sheet-image-link")) return;
  const guide = nearestPageBreakGuideAtPoint(event.clientX, event.clientY);
  if (!guide) return;
  const axis = guide.dataset.pageBreakAxis === "row" ? "row" : "col";
  const index = Math.max(2, Number(guide.dataset.pageBreakIndex) || 2);
  const manual = guide.dataset.pageBreakManual === "true";
  openPageBreakContextMenuAt(event, axis, index, manual);
}



function nearestPageBreakGuideAtPoint(clientX, clientY) {
  let best = null;
  let bestDistance = Number.POSITIVE_INFINITY;
  document.querySelectorAll(".sheet-view-guide.is-page-break").forEach((guide) => {
    const rect = guide.getBoundingClientRect();
    if (!rect.width && !rect.height) return;
    const pad = 8;
    if (clientX < rect.left - pad || clientX > rect.right + pad || clientY < rect.top - pad || clientY > rect.bottom + pad) return;
    const centerX = rect.left + rect.width / 2;
    const centerY = rect.top + rect.height / 2;
    const axis = guide.dataset.pageBreakAxis === "row" ? "row" : "col";
    const distance = axis === "row" ? Math.abs(clientY - centerY) : Math.abs(clientX - centerX);
    if (distance >= bestDistance) return;
    best = guide;
    bestDistance = distance;
  });
  return best;
}



function openPageBreakContextMenuAt(event, axis, index, manual) {
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  const label = pageBreakPositionLabel(axis, index);
  openViewCommandMenu({ left: event.clientX, bottom: event.clientY }, "page-break", "改ページ", `
    ${viewCommandMenuButtonHtml(`page-break:insert:${axis}:${index}`, "改ページの挿入(I)", "table-2", manual, { description: manual ? `${label} には手動改ページがあります。` : `${label} に手動改ページを挿入します。` })}
    ${viewCommandMenuButtonHtml(`page-break:remove:${axis}:${index}`, "改ページの解除(R)", "eraser", !manual, { description: manual ? `${label} の手動改ページを解除します。` : "自動改ページは解除できません。挿入すると手動改ページになります。" })}
    ${viewCommandMenuButtonHtml("page-break:reset", "すべての改ページを解除(A)", "delete", false, { description: "このシートの手動改ページをすべて解除します。" })}
    <div class="home-command-menu-separator" role="separator"></div>
    ${viewCommandMenuButtonHtml("page-break:page-setup", "ページ設定(P)...", "settings", false, { description: "印刷の向き、余白、用紙サイズなどを変更します。" })}
    ${viewCommandMenuButtonHtml("page-break:close-preview", "改ページ プレビューを閉じる(V)", "grid-2x2", false, { description: "標準表示に戻します。" })}
  `);
}



function pageBreakPositionLabel(axis, index) {
  return axis === "row" ? `${index} 行目の前` : `${columnName(index)} 列の前`;
}



function workbookWindowMenuDescription(view) {
  const sheet = state.model?.sheets?.[view?.sheetIndex] || activeSheet();
  const sheetName = sheet?.name || "Sheet";
  const viewLabel = view?.viewMode === "page-layout" ? "ページ レイアウト"
    : view?.viewMode === "page-break" ? "改ページ プレビュー"
      : "標準";
  return `${sheetName} - ${viewLabel} - ${clampZoom(view?.zoomScale || 100)}%`;
}



function openViewCommandMenu(trigger, kind, label, contentHtml) {
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonColorPalette();
  closeRibbonBorderMenu();
  closeRibbonUnderlineMenu();
  closeRibbonShapeMenu();
  closeRibbonInkEraserMenu();
  closeShapeFormatMenu();
  closeViewCommandMenu();
  closeHomeCommandMenu();
  closeDataCommandMenu();
  closeViewCommandMenu();
  const menu = document.createElement("div");
  menu.id = "viewCommandMenu";
  menu.className = `home-command-menu view-command-menu is-${kind}`;
  menu.dataset.viewCommandKind = kind;
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", label);
  menu.innerHTML = contentHtml;
  enhanceCommandMenuAccessKeys(menu);
  document.body.append(menu);
  if (window.lucide) window.lucide.createIcons({ icons });
  const rect = typeof trigger?.getBoundingClientRect === "function"
    ? trigger.getBoundingClientRect()
    : trigger || { left: 0, bottom: 0 };
  positionPopupMenu(menu, rect.left, rect.bottom + 2);
  state.viewCommandMenu = kind;
  menu.querySelector("button:not(:disabled)")?.focus({ preventScroll: true });
}



function viewCommandMenuButtonHtml(action, label, icon, disabled = false, options = {}) {
  const disabledAttr = disabled ? " disabled" : "";
  const checkable = Boolean(options.checkable || options.checked);
  const checkedAttr = checkable ? ` aria-checked="${options.checked ? "true" : "false"}"` : "";
  const role = checkable ? "menuitemcheckbox" : "menuitem";
  const checkHtml = checkable ? `<span class="home-command-menu-check" aria-hidden="true">${options.checked ? "✓" : ""}</span>` : "";
  const description = String(options.description || "").trim();
  const classAttr = description ? ' class="has-description"' : "";
  const accessKey = menuAccessKeyFromLabel(label);
  const accessKeyAttr = accessKey ? ` data-access-key="${escapeAttr(accessKey)}"` : "";
  const textHtml = description
    ? `<span class="home-command-menu-text"><span class="home-command-menu-label">${accessKeyLabelHtml(label)}</span><small>${escapeHtml(description)}</small></span>`
    : `<span>${accessKeyLabelHtml(label)}</span>`;
  return `<button type="button"${classAttr} role="${role}" data-view-menu-action="${escapeAttr(action)}"${accessKeyAttr}${checkedAttr}${disabledAttr}>${checkHtml}<i data-lucide="${escapeAttr(icon)}"></i>${textHtml}</button>`;
}



function menuAccessKeyFromLabel(label) {
  const match = String(label || "").match(/\(([A-Za-z0-9])\)|（([A-Za-z0-9])）/);
  if (match) return (match[1] || match[2] || "").toUpperCase();
  const leading = String(label || "").trimStart().match(/^([0-9])\s+/);
  return leading ? leading[1] : "";
}



function commandMenuAccessKeyAttr(label) {
  const accessKey = menuAccessKeyFromLabel(label);
  return accessKey ? ` data-access-key="${escapeAttr(accessKey)}"` : "";
}



function closeViewCommandMenu() {
  document.querySelector("#viewCommandMenu")?.remove();
  state.viewCommandMenu = null;
}



function closeViewCommandMenuOnOutsidePointer(event) {
  if (!state.viewCommandMenu) return;
  if (event.target.closest?.("#viewCommandMenu")) return;
  if (event.target.closest?.("[data-ribbon-command='view-freeze-menu'], [data-ribbon-command='view-switch-windows'], [data-ribbon-command='view-sheet-view-options'], [data-ribbon-command='view-macros-menu']")) return;
  closeViewCommandMenu();
}



function closeViewCommandMenuOnEscape(event) {
  if (event.key !== "Escape" || !document.querySelector("#viewCommandMenu")) return;
  consumeCommandMenuKey(event);
  closeViewCommandMenu();
}



function handleViewCommandMenuAction(event) {
  const button = event.currentTarget;
  if (button.disabled) return;
  const action = button.dataset.viewMenuAction || "";
  closeViewCommandMenu();
  if (action === "freeze-selection") setFreezePanes("selection");
  else if (action === "freeze-top-row") setFreezePanes("top-row");
  else if (action === "freeze-first-column") setFreezePanes("first-column");
  else if (action === "unfreeze") setFreezePanes("unfreeze");
  else if (action.startsWith("switch:")) switchWorkbookWindow(action.slice("switch:".length));
  else if (action === "switch-dialog") openSwitchWindowsDialog();
  else if (action === "sheet-view:new") createSheetView();
  else if (action === "sheet-view:keep") keepActiveSheetView();
  else if (action === "sheet-view:exit") exitActiveSheetView();
  else if (action === "sheet-view:default") exitActiveSheetView();
  else if (action.startsWith("sheet-view:activate:")) activateSheetView(action.slice("sheet-view:activate:".length));
  else if (action === "sheet-view:update") updateActiveSheetView();
  else if (action === "sheet-view:duplicate") duplicateActiveSheetView();
  else if (action === "sheet-view:rename") openSheetViewRenameDialog();
  else if (action === "sheet-view:delete") deleteActiveSheetView();
  else if (action.startsWith("page-break:insert:")) {
    const [, , axis, index] = action.split(":");
    insertPageBreakAt(axis, Number(index));
  } else if (action.startsWith("page-break:remove:")) {
    const [, , axis, index] = action.split(":");
    removePageBreakAt(axis, Number(index));
  } else if (action === "page-break:reset") resetAllPageBreaks();
  else if (action === "page-break:page-setup") openPageSetupDialog();
  else if (action === "page-break:close-preview") setActiveSheetView("normal");
  else if (action === "macro-show") openMacrosDialog();
  else if (action === "macro-record") openMacroRecordDialog();
  else if (action === "macro-stop-recording") stopMacroRecording();
  else if (action === "macro-use-relative-refs") toggleMacroRelativeReferences();
  else if (action === "macro-security") openMacroSecurityDialog();
  restoreWorksheetFocusAfterViewCommand();
}



function restoreWorksheetFocusAfterViewCommand() {
  restoreWorksheetFocusAfterCommandPopup();
}



function restoreWorksheetFocusAfterCommandPopup() {
  const restore = () => {
    if (!state.model) return;
    if (state.sheetTabRename) return;
    if (document.querySelector("#dataDialog, #messageBox, #zoomDialog, #findDialog, #goToDialog, #specialDialog, #homeCommandMenu, #viewCommandMenu, #dataCommandMenu, #dataValidationMenu, #whatIfMenu, #ribbonColorPalette, #ribbonBorderMenu, #ribbonUnderlineMenu, #ribbonInkEraserMenu, #ribbonShapeMenu, #shapeFormatMenu, #sheetTabContextMenu, #sheetTabNavigatorMenu, #cellContextMenu, #statusBarContextMenu, #queryPaneContextMenu, #quickAnalysisPopup, #fileSideMenu:not(.hidden), #ribbonDisplayMenu:not(.hidden), #nameBoxMenu, #cellAnnotationEditor, #cellCommentThreadPopup, #cellNotePopup, #translationPane, #shapeTextEditor")) return;
    const editor = activeCellEditElementForFormatting();
    if (editor) {
      editor.focus?.({ preventScroll: true });
      restoreCellEditSelection(editor);
      return;
    }
    const host = $sheetHost[0];
    if (!host) return;
    if (!host.hasAttribute("tabindex")) host.setAttribute("tabindex", "-1");
    host.focus?.({ preventScroll: true });
  };
  restore();
  requestAnimationFrame(restore);
}



function scheduleWorksheetFocusAfterCommandPopup() {
  requestAnimationFrame(restoreWorksheetFocusAfterCommandPopup);
}



function handleWorkbookWindowListConfirm(event) {
  if (event.type === "keydown" && (event.key !== "Enter" || isElementImeCompositionEvent(event))) return;
  const dialog = event.currentTarget.closest("[data-data-dialog]");
  const type = dialog?.dataset.dataDialog || state.dataDialog?.type || "";
  if (type !== "view-unhide-window" && type !== "view-side-by-side-window" && type !== "view-switch-window") return;
  event.preventDefault();
  event.stopPropagation();
  if (type === "view-unhide-window") {
    const id = event.currentTarget.value || dialog?.querySelector("[data-data-field='windowId']")?.value || "";
    closeDataDialog();
    unhideWorkbookWindow(id);
    return;
  }
  if (type === "view-switch-window") {
    const id = event.currentTarget.value || dialog?.querySelector("[data-data-field='switchWindowId']")?.value || "";
    closeDataDialog();
    switchWorkbookWindow(id);
    return;
  }
  const id = event.currentTarget.value || dialog?.querySelector("[data-data-field='sideBySideWindowId']")?.value || "";
  const activeId = state.model?.activeWindowId || "main";
  closeDataDialog();
  enableViewSideBySide(activeId, id);
}



function customViewsUnavailableSummary(model = state.model) {
  const summary = workbookTableUsageSummary(model);
  const customViews = Array.isArray(model?.customViews) ? model.customViews : [];
  return summary.hasTables && !customViews.length ? summary : null;
}



function customViewsUnavailableMessage(summary = workbookTableUsageSummary()) {
  const sheetText = (summary?.sheets || [])
    .slice(0, 3)
    .map((item) => `${item.name}${item.count > 1 ? ` (${item.count})` : ""}`)
    .join("、");
  const suffix = sheetText ? `（${sheetText}${summary.sheets.length > 3 ? " ほか" : ""}）` : "";
  return `このブックにはテーブルが含まれているため、ユーザー設定のビューを使用できません。テーブルを範囲に変換してから実行してください。${suffix}`;
}



function showCustomViewsUnavailableIfNeeded() {
  const summary = customViewsUnavailableSummary();
  if (!summary) return false;
  showMessageBox(customViewsUnavailableMessage(summary), { title: "Cht WebSheet" });
  return true;
}



function openCustomViewsDialog(options = {}) {
  if (options.force !== true && showCustomViewsUnavailableIfNeeded()) return;
  state.model.customViews = Array.isArray(state.model.customViews) ? state.model.customViews : [];
  closeDataDialog({ keepState: true });
  const overlay = $(customViewsDialogHtml(options));
  $("body").append(overlay);
  state.dataDialog = { type: "custom-views" };
  syncCustomViewsDialog(overlay[0].querySelector("[data-data-dialog='custom-views']"));
}



function customViewsDialogHtml(options = {}) {
  const selectedName = String(options.selectedName || "");
  const tableRestriction = workbookTableUsageSummary();
  const addDisabled = tableRestriction.hasTables;
  const addTitle = addDisabled ? customViewTableRestrictionMessage(tableRestriction) : "新しいビューを追加します";
  const optionsHtml = (state.model.customViews || [])
    .map((view, index) => `<option value="${index}"${view.name === selectedName || (!selectedName && index === 0) ? " selected" : ""}>${escapeHtml(view.name)}</option>`)
    .join("");
  const restrictionHtml = addDisabled
    ? `<div class="data-dialog-summary custom-view-restriction">${escapeHtml(customViewTableRestrictionMessage(tableRestriction))}</div>`
    : "";
  return `
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog custom-views-dialog" role="dialog" aria-modal="true" aria-labelledby="customViewsTitle" data-data-dialog="custom-views">
        <div class="excel-mini-titlebar" id="customViewsTitle">ユーザー設定のビュー</div>
        <div class="excel-mini-body data-dialog-body">
          <label class="excel-dialog-field"><span class="excel-dialog-caption">ビュー(V):</span>
            <select size="7" data-data-field="customViewIndex">${optionsHtml}</select>
          </label>
          ${restrictionHtml}
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="custom-view-show">表示(S)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="custom-view-add"${addDisabled ? " disabled" : ""} title="${escapeAttr(addTitle)}">追加(A)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="custom-view-delete">削除(D)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">閉じる</button>
        </div>
      </div>
    </div>`;
}



function defaultCustomViewName() {
  const count = (state.model?.customViews || []).length + 1;
  return `ビュー ${count}`;
}



function workbookTableUsageSummary(model = state.model) {
  const sheets = (model?.sheets || [])
    .map((sheet) => {
      const count = (sheet?.tables || []).filter((table) => table?.range).length;
      return { name: sheet?.name || "Sheet", count };
    })
    .filter((item) => item.count > 0);
  return {
    hasTables: sheets.length > 0,
    tableCount: sheets.reduce((total, item) => total + item.count, 0),
    sheets,
  };
}



function customViewTableRestrictionMessage(summary = workbookTableUsageSummary()) {
  const sheetText = (summary.sheets || [])
    .slice(0, 3)
    .map((item) => `${item.name}${item.count > 1 ? ` (${item.count})` : ""}`)
    .join("、");
  const suffix = sheetText ? `（${sheetText}${summary.sheets.length > 3 ? " ほか" : ""}）` : "";
  return `このブックにはテーブルが含まれているため、ユーザー設定のビューを追加できません。テーブルを範囲に変換してから実行してください。${suffix}`;
}



function showCustomViewTableRestrictionIfNeeded() {
  const summary = workbookTableUsageSummary();
  if (!summary.hasTables) return false;
  showMessageBox(customViewTableRestrictionMessage(summary), { title: "Cht WebSheet" });
  return true;
}



function openAddCustomViewDialog() {
  if (showCustomViewTableRestrictionIfNeeded()) return;
  closeDataDialog({ keepState: true });
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog custom-view-add-dialog" role="dialog" aria-modal="true" aria-labelledby="customViewAddTitle" data-data-dialog="custom-view-add">
        <div class="excel-mini-titlebar" id="customViewAddTitle">ビューの追加</div>
        <div class="excel-mini-body data-dialog-body">
          <label class="excel-dialog-field custom-view-name-field"><span class="excel-dialog-caption">名前(N):</span>
            <input type="text" data-data-field="customViewName" value="${escapeAttr(defaultCustomViewName())}" />
          </label>
          <fieldset class="data-dialog-fieldset custom-view-include-options">
            <legend>ビューに含める設定</legend>
            <label class="data-dialog-checkbox"><input type="checkbox" data-data-field="customViewPrintSettings" checked />印刷の設定(P)</label>
            <label class="data-dialog-checkbox"><input type="checkbox" data-data-field="customViewHiddenRowsFilters" checked />非表示の行、列、フィルターの設定(R)</label>
          </fieldset>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="custom-view-add-ok">OK</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "custom-view-add" };
  overlay.find("[data-data-field='customViewName']").trigger("focus").trigger("select");
}



function syncCustomViewsDialog(dialog) {
  if (!dialog) return;
  const list = dialog.querySelector("[data-data-field='customViewIndex']");
  if (list && !list.querySelector("option:checked") && list.options.length) list.selectedIndex = 0;
  const hasSelection = Boolean(dialog.querySelector("[data-data-field='customViewIndex'] option:checked"));
  const tableRestriction = workbookTableUsageSummary();
  const addButton = dialog.querySelector("[data-data-dialog-action='custom-view-add']");
  dialog.querySelector("[data-data-dialog-action='custom-view-show']").disabled = !hasSelection;
  dialog.querySelector("[data-data-dialog-action='custom-view-delete']").disabled = !hasSelection;
  if (addButton) {
    addButton.disabled = tableRestriction.hasTables;
    addButton.title = tableRestriction.hasTables ? customViewTableRestrictionMessage(tableRestriction) : "新しいビューを追加します";
  }
  list?.focus({ preventScroll: true });
}



function addCustomViewFromDialog(dialog) {
  if (showCustomViewTableRestrictionIfNeeded()) return;
  const name = String(dialog?.querySelector("[data-data-field='customViewName']")?.value || "").trim();
  if (!name) {
    showMessageBox("ビュー名を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  state.model.customViews = Array.isArray(state.model.customViews) ? state.model.customViews : [];
  const existingIndex = state.model.customViews.findIndex((view) => view.name.toLocaleLowerCase("ja-JP") === name.toLocaleLowerCase("ja-JP"));
  const view = captureCustomView(name, {
    includePrintSettings: dialog?.querySelector("[data-data-field='customViewPrintSettings']")?.checked !== false,
    includeHiddenRowsColumnsAndFilters: dialog?.querySelector("[data-data-field='customViewHiddenRowsFilters']")?.checked !== false,
  });
  if (existingIndex >= 0) {
    showMessageBox(`'${name}' は既に存在します。置き換えますか?`, {
      title: "Cht WebSheet",
      tone: "warning",
      buttons: [
        { label: "はい(Y)", result: "yes", callback: () => saveCustomViewSnapshot(name, view, existingIndex) },
        { label: "いいえ(N)", result: "no" },
      ],
      cancelResult: "no",
    });
    return;
  }
  saveCustomViewSnapshot(name, view, existingIndex);
}



function saveCustomViewSnapshot(name, view, existingIndex = -1) {
  state.model.customViews = Array.isArray(state.model.customViews) ? state.model.customViews : [];
  const workbookViewBefore = captureWorkbookViewState();
  if (existingIndex >= 0) state.model.customViews[existingIndex] = view;
  else state.model.customViews.push(view);
  commitWorkbookViewChange(workbookViewBefore, `ユーザー設定のビューを保存しました：${name}`, `${name} ユーザー設定のビュー保存`, {
    render: false,
  });
  openCustomViewsDialog({ selectedName: name });
}



function captureCustomView(name, options = {}) {
  const includePrintSettings = options.includePrintSettings !== false;
  const includeHiddenRowsColumnsAndFilters = options.includeHiddenRowsColumnsAndFilters !== false;
  updateActiveWorkbookWindowScrollState();
  saveActiveWindowViewState();
  const host = $sheetHost[0];
  const selectionRanges = activeSelectionRanges().map((range) => ({ ...range }));
  const selectionRange = selectionRanges[0] || activeSelectionRange();
  const activeScroll = {
    top: Math.max(0, Math.round(host?.scrollTop || 0)),
    left: Math.max(0, Math.round(host?.scrollLeft || 0)),
  };
  const activeTopLeftCell = topLeftCellFromScrollRecord(activeSheet(), activeScroll);
  const savedWindowRecords = storedWorkbookWindowRecords();
  return {
    name,
    createdAt: new Date().toISOString(),
    activeSheetIndex: state.activeSheetIndex,
    showFormulaBar: !$formulaBar.hasClass("hidden"),
    includePrintSettings,
    includeHiddenRowsColumnsAndFilters,
    selectedCell: state.selected ? { ...state.selected } : null,
    selectionRange: selectionRange ? { ...selectionRange } : null,
    selectionRanges,
    activeScroll,
    topLeftCell: activeTopLeftCell,
    activeWindowId: state.model.activeWindowId || "main",
    mainWindowHidden: state.model.mainWindowHidden === true,
    mainWindowView: cloneJsonLikeValue(state.model.mainWindowView || null),
    windowViews: cloneJsonLikeValue(state.model.windowViews || []),
    workbookWindowsCaptured: true,
    viewSideBySide: cloneJsonLikeValue(state.model.viewSideBySide || null),
    synchronousScrolling: state.model.synchronousScrolling === true,
    windowArrangement: cloneJsonLikeValue(state.model.windowArrangement || null),
    sheets: state.model.sheets.map((sheet, sheetIndex) => {
      const viewMode = normalizedSheetViewMode(sheet.viewMode);
      const sheetScroll = customViewScrollSnapshotForSheet(sheet, sheetIndex, savedWindowRecords, activeScroll, activeTopLeftCell);
      const sheetSelectionRanges = selectionRanges.filter((range) => range.sheetIndex === sheetIndex).map((range) => ({ ...range }));
      const sheetSelectionRange = sheetSelectionRanges[0] || (selectionRange?.sheetIndex === sheetIndex ? { ...selectionRange } : null);
      const sheetSelectedCell = state.selected?.sheetIndex === sheetIndex ? { ...state.selected } : null;
      const saved = {
        name: sheet.name,
        viewMode,
        zoomScale: sheet.zoomScale || 100,
        showGridLines: sheet.showGridLines !== false,
        showHeadings: sheet.showHeadings !== false,
        showRuler: normalizedSheetShowRuler(sheet.showRuler, viewMode),
        showObjects: sheet.showObjects !== false,
        showOutlineSymbols: sheet.showOutlineSymbols !== false,
        showFormulas: sheet.showFormulas === true,
        showZeros: sheet.showZeros !== false,
        showWhiteSpace: sheet.showWhiteSpace !== false,
        rightToLeft: sheet.rightToLeft === true,
        selectedCell: sheetSelectedCell,
        selectionRange: sheetSelectionRange,
        selectionRanges: sheetSelectionRanges,
        activeScroll: sheetScroll.activeScroll,
        topLeftCell: sheetScroll.topLeftCell,
        freezePanes: cloneJsonLikeValue(sheet.freezePanes),
        splitPane: cloneJsonLikeValue(sheet.splitPane),
        sheetViews: cloneJsonLikeValue(sheet.sheetViews || []),
        activeSheetViewId: sheet.activeSheetViewId || null,
      };
      if (includePrintSettings) {
        saved.pageLayout = cloneJsonLikeValue(sheet.pageLayout);
        saved.headerFooter = cloneJsonLikeValue(sheet.headerFooter);
      }
      if (includeHiddenRowsColumnsAndFilters) {
        saved.hiddenRows = hiddenIndexesForModels(sheet.rows);
        saved.hiddenColumns = hiddenIndexesForModels(sheet.columns);
        saved.autoFilter = cloneJsonLikeValue(sheet.autoFilter);
        saved.sortState = cloneJsonLikeValue(sheet.sortState);
      }
      return saved;
    }),
  };
}



function customViewScrollSnapshotForSheet(sheet, sheetIndex, windowRecords = [], activeScroll = null, activeTopLeftCell = null) {
  if (sheetIndex === state.activeSheetIndex) {
    const active = normalizeCustomViewScroll(activeScroll || {});
    return {
      activeScroll: { top: active.top, left: active.left },
      topLeftCell: activeTopLeftCell || topLeftCellFromScrollRecord(sheet, active) || { row: 1, col: 1 },
    };
  }
  const record = (windowRecords || []).find((view) => view && view.sheetIndex === sheetIndex);
  if (record) {
    const fromCell = record.scrollUnit === "cell" ? scrollRecordFromTopLeftCell(sheet, record.topLeftCell) : null;
    const activeScrollForSheet = {
      top: Math.max(0, Math.round(Number(fromCell?.scrollTop ?? record.scrollTop) || 0)),
      left: Math.max(0, Math.round(Number(fromCell?.scrollLeft ?? record.scrollLeft) || 0)),
    };
    return {
      activeScroll: activeScrollForSheet,
      topLeftCell: normalizeCustomTopLeftCell(record.topLeftCell) || topLeftCellFromScrollRecord(sheet, activeScrollForSheet) || { row: 1, col: 1 },
    };
  }
  return {
    activeScroll: { top: 0, left: 0 },
    topLeftCell: { row: 1, col: 1 },
  };
}



function hiddenIndexesForModels(models = []) {
  return (models || [])
    .map((model, index) => (model?.hidden ? index + 1 : null))
    .filter(Boolean);
}



function selectedCustomView(dialog) {
  const index = Number(dialog?.querySelector("[data-data-field='customViewIndex']")?.value);
  return state.model?.customViews?.[index] || null;
}



function savedCustomViewSheetForSheet(customView, sheet, sheetIndex) {
  return customView?.sheets?.find((candidate) => candidate.name === sheet?.name) || customView?.sheets?.[sheetIndex] || null;
}



function applyCustomViewSheetState(sheet, savedSheet, customView) {
  if (!sheet || !savedSheet) return;
  sheet.viewMode = normalizedSheetViewMode(savedSheet.viewMode);
  applySheetZoomWithoutHistory(sheet, savedSheet.zoomScale || sheet.zoomScale || 100);
  sheet.showGridLines = savedSheet.showGridLines !== false;
  sheet.showHeadings = savedSheet.showHeadings !== false;
  sheet.showRuler = normalizedSheetShowRuler(savedSheet.showRuler, sheet.viewMode);
  sheet.showObjects = savedSheet.showObjects !== false;
  sheet.showOutlineSymbols = savedSheet.showOutlineSymbols !== false;
  sheet.showFormulas = savedSheet.showFormulas === true;
  sheet.showZeros = savedSheet.showZeros !== false;
  sheet.showWhiteSpace = savedSheet.showWhiteSpace !== false;
  sheet.rightToLeft = savedSheet.rightToLeft === true;
  sheet.freezePanes = normalizePaneState(savedSheet.freezePanes);
  sheet.splitPane = normalizePaneState(savedSheet.splitPane);
  restoreCustomViewActiveSheetView(sheet, savedSheet);
  if (customView?.includePrintSettings !== false) {
    if (savedSheet.pageLayout) {
      sheet.pageLayout = normalizedPageLayoutSettings(savedSheet.pageLayout, sheet);
    }
    if (Object.prototype.hasOwnProperty.call(savedSheet, "headerFooter")) {
      sheet.headerFooter = normalizedHeaderFooterSettings(savedSheet.headerFooter, sheet.pageLayout);
    }
  }
  if (customView?.includeHiddenRowsColumnsAndFilters !== false) {
    applyHiddenIndexesToModels(sheet.rows, savedSheet.hiddenRows);
    applyHiddenIndexesToModels(sheet.columns, savedSheet.hiddenColumns);
    sheet.autoFilter = cloneJsonLikeValue(savedSheet.autoFilter);
    sheet.sortState = cloneJsonLikeValue(savedSheet.sortState);
  }
}



function restoreCustomViewActiveSheetView(sheet, savedSheet) {
  if (!sheet || !savedSheet || !Object.prototype.hasOwnProperty.call(savedSheet, "activeSheetViewId")) return;
  sheet.sheetViews = Array.isArray(sheet.sheetViews) ? sheet.sheetViews.map(normalizeSheetView).filter(Boolean) : [];
  const savedActiveId = String(savedSheet.activeSheetViewId || "");
  sheet.activeSheetViewId = savedActiveId && sheet.sheetViews.some((view) => view.id === savedActiveId)
    ? savedActiveId
    : null;
}



function customViewActiveSheetIndex(customView) {
  if (!state.model?.sheets?.length) return 0;
  return Math.max(0, Math.min(
    state.model.sheets.length - 1,
    Math.round(Number(customView?.activeSheetIndex) || 0),
  ));
}



function customViewWindowRecordForCurrentWindow(customView, currentRecord) {
  if (!state.model) return null;
  const sheetIndex = customViewActiveSheetIndex(customView);
  const sheet = state.model.sheets[sheetIndex] || activeSheet();
  const savedSheet = savedCustomViewSheetForSheet(customView, sheet, sheetIndex);
  const scrollRecord = restoreScrollRecordForView(
    sheet,
    savedSheet?.activeScroll || customView.activeScroll,
    savedSheet?.topLeftCell || customView.topLeftCell,
  );
  const selectedCell = normalizeCustomViewPoint(savedSheet?.selectedCell || customView.selectedCell, sheetIndex);
  const selectionRange = normalizeCustomViewRange(savedSheet?.selectionRange || customView.selectionRange, sheetIndex);
  const selectionRanges = (Array.isArray(savedSheet?.selectionRanges) && savedSheet.selectionRanges.length ? savedSheet.selectionRanges : customView.selectionRanges || [])
    .map((range) => normalizeCustomViewRange(range, sheetIndex))
    .filter(Boolean);
  const baseRecord = customView.mainWindowView && typeof customView.mainWindowView === "object"
    ? customView.mainWindowView
    : {};
  const viewValue = (key, fallback) => (savedSheet && Object.prototype.hasOwnProperty.call(savedSheet, key) ? savedSheet[key] : fallback);
  const viewMode = normalizedSheetViewMode(viewValue("viewMode", baseRecord.viewMode || sheet?.viewMode || "normal"));
  const topLeftCell = normalizeCustomTopLeftCell(scrollRecord.topLeftCell) || topLeftCellFromScrollRecord(sheet, scrollRecord);
  return normalizeWorkbookWindowView({
    ...(currentRecord || {}),
    ...baseRecord,
    id: currentRecord?.id || state.model.activeWindowId || "main",
    name: currentRecord?.name || baseRecord.name || state.model.sourceName || "Book1",
    hidden: currentRecord?.hidden === true,
    sheetIndex,
    viewMode,
    zoomScale: viewValue("zoomScale", baseRecord.zoomScale || sheet?.zoomScale || 100),
    showGridLines: viewValue("showGridLines", baseRecord.showGridLines) !== false,
    showHeadings: viewValue("showHeadings", baseRecord.showHeadings) !== false,
    showRuler: normalizedSheetShowRuler(viewValue("showRuler", baseRecord.showRuler), viewMode),
    showObjects: viewValue("showObjects", baseRecord.showObjects) !== false,
    showOutlineSymbols: viewValue("showOutlineSymbols", baseRecord.showOutlineSymbols) !== false,
    showFormulas: viewValue("showFormulas", baseRecord.showFormulas) === true,
    showZeros: viewValue("showZeros", baseRecord.showZeros) !== false,
    showWhiteSpace: viewValue("showWhiteSpace", baseRecord.showWhiteSpace) !== false,
    rightToLeft: viewValue("rightToLeft", baseRecord.rightToLeft) === true,
    freezePanes: normalizePaneState(savedSheet?.freezePanes ?? baseRecord.freezePanes),
    splitPane: normalizePaneState(savedSheet?.splitPane ?? baseRecord.splitPane),
    scrollTop: scrollRecord.scrollTop,
    scrollLeft: scrollRecord.scrollLeft,
    scrollUnit: "pixel",
    topLeftCell,
    selectedCell,
    selectionRange,
    selectionRanges: selectionRanges.length ? selectionRanges : (selectionRange ? [selectionRange] : []),
  });
}



function applyCustomViewToCurrentWorkbookWindow(customView) {
  if (!state.model) return null;
  const activeId = state.model.activeWindowId || "main";
  const currentRecord = storedWorkbookWindowRecordById(activeId) || normalizeWorkbookWindowView({
    id: activeId,
    name: activeId === "main" ? state.model.sourceName || "Book1" : "",
  });
  const record = customViewWindowRecordForCurrentWindow(customView, currentRecord);
  if (!record) return null;
  if (activeId === "main") {
    state.model.mainWindowView = normalizeWorkbookWindowView({
      ...record,
      id: "main",
      name: state.model.sourceName || record.name || "Book1",
      hidden: state.model.mainWindowHidden === true,
    });
  } else {
    const target = mutableWorkbookWindowRecordById(activeId);
    if (target) Object.assign(target, record);
  }
  return record;
}



function applySelectedCustomView(dialog) {
  const view = selectedCustomView(dialog);
  if (!view) return;
  const normalizedView = normalizeWorkbookCustomView(view) || view;
  const workbookViewBefore = captureWorkbookViewState();
  const sheetIndexes = state.model.sheets.map((_sheet, index) => index);
  const sheetsBefore = captureSheetsForHistory(sheetIndexes);
  const selectionBefore = activeSelectionRange();
  const selectionRangesBefore = activeSelectionRanges();
  normalizedView.sheets?.forEach((savedSheet, index) => {
    const sheet = state.model.sheets.find((candidate) => candidate.name === savedSheet.name) || state.model.sheets[index];
    applyCustomViewSheetState(sheet, savedSheet, normalizedView);
  });
  const applyWorkbookWindows = normalizedView.workbookWindowsCaptured === true;
  if (applyWorkbookWindows) state.model.mainWindowHidden = normalizedView.mainWindowHidden === true;
  if (normalizedView.mainWindowView) {
    state.model.mainWindowView = normalizeWorkbookWindowView({
      id: "main",
      name: state.model.sourceName || "Book1",
      ...(applyWorkbookWindows ? {} : state.model.mainWindowView || {}),
      ...normalizedView.mainWindowView,
      hidden: applyWorkbookWindows ? state.model.mainWindowHidden : state.model.mainWindowHidden === true,
    });
  }
  if (applyWorkbookWindows) {
    state.model.windowViews = Array.isArray(normalizedView.windowViews) ? normalizedView.windowViews.map(normalizeWorkbookWindowView) : [];
    state.model.viewSideBySide = normalizedView.viewSideBySide ? cloneJsonLikeValue(normalizedView.viewSideBySide) : null;
    state.model.synchronousScrolling = normalizedView.synchronousScrolling === true;
    state.model.windowArrangement = normalizedView.windowArrangement ? cloneJsonLikeValue(normalizedView.windowArrangement) : state.model.windowArrangement;
    state.model.activeWindowId = normalizedView.activeWindowId || "main";
  } else {
    applyCustomViewToCurrentWorkbookWindow(normalizedView);
  }
  if (!storedVisibleWorkbookWindowRecords().some((record) => record.id === state.model.activeWindowId)) {
    state.model.activeWindowId = storedVisibleWorkbookWindowRecords()[0]?.id || "main";
    state.model.mainWindowHidden = false;
  }
  const restoreRecord = storedWorkbookWindowRecordById(state.model.activeWindowId || "main");
  const fallbackSheetIndex = Math.min(state.model.sheets.length - 1, Math.max(0, Number(normalizedView.activeSheetIndex) || 0));
  state.activeSheetIndex = state.model.sheets[restoreRecord?.sheetIndex]
    ? restoreRecord.sheetIndex
    : fallbackSheetIndex;
  if (restoreRecord && state.model.sheets[restoreRecord.sheetIndex]) {
    applyWorkbookWindowSheetView(activeSheet(), restoreRecord);
    applyCustomViewSheetState(activeSheet(), savedCustomViewSheetForSheet(normalizedView, activeSheet(), state.activeSheetIndex), normalizedView);
    restoreWorkbookWindowSelection(restoreRecord);
  } else {
    restoreCustomViewSelection(normalizedView);
  }
  const savedActiveSheet = savedCustomViewSheetForSheet(normalizedView, activeSheet(), state.activeSheetIndex);
  const scrollRecord = applyWorkbookWindows && restoreRecord
    ? restoreRecord
    : restoreScrollRecordForView(
      activeSheet(),
      savedActiveSheet?.activeScroll || normalizedView.activeScroll,
      savedActiveSheet?.topLeftCell || normalizedView.topLeftCell,
    );
  state.showFormulas = activeSheet()?.showFormulas === true;
  $formulaBar.toggleClass("hidden", normalizedView.showFormulaBar === false);
  state.model.showFormulaBar = normalizedView.showFormulaBar !== false;
  closeDataDialog();
  renderWorkbook();
  restoreWorkbookWindowScrollStable(scrollRecord);
  requestAnimationFrame(() => {
    updateActiveWorkbookWindowScrollState();
    syncPairedWorkbookWindowScrollState();
  });
  updateSelectionUi();
  updateFormulaBarForSelection();
  syncViewRibbonButtons();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(sheetIndexes),
    workbookViewBefore,
    workbookViewAfter: captureWorkbookViewState(),
    selectionBefore,
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: selectionRangesBefore,
    selectionAfterRanges: activeSelectionRanges(),
    label: `${view.name} ユーザー設定のビュー表示`,
  });
  scheduleWorkbookPersistence();
  setStatus(`ユーザー設定のビューを表示しました：${view.name}`);
}



function restoreCustomViewSelection(view) {
  const activeSheetIndex = Math.min(state.model.sheets.length - 1, Math.max(0, Number(view.activeSheetIndex) || 0));
  const ranges = (Array.isArray(view.selectionRanges) && view.selectionRanges.length ? view.selectionRanges : [view.selectionRange])
    .map((range) => normalizeCustomViewRange(range, activeSheetIndex))
    .filter((range) => range && range.sheetIndex === activeSheetIndex);
  const firstRange = ranges[0] || { sheetIndex: activeSheetIndex, top: 1, left: 1, bottom: 1, right: 1 };
  const selected = normalizeCustomViewPoint(view.selectedCell, activeSheetIndex) || {
    sheetIndex: activeSheetIndex,
    row: firstRange.top,
    col: firstRange.left,
  };
  state.selected = {
    sheetIndex: activeSheetIndex,
    row: Math.min(Math.max(1, selected.row), state.model.sheets[activeSheetIndex]?.rowCount || selected.row),
    col: Math.min(Math.max(1, selected.col), state.model.sheets[activeSheetIndex]?.colCount || selected.col),
  };
  state.selectionRange = { ...firstRange };
  state.selectionRanges = ranges.length ? ranges.map((range) => ({ ...range })) : [{ ...firstRange }];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  state.formulaEdit = null;
}



function applyHiddenIndexesToModels(models = [], indexes = []) {
  const hidden = new Set((indexes || []).map((value) => Math.max(1, Math.round(Number(value) || 0))).filter(Boolean));
  (models || []).forEach((model, index) => {
    if (model) model.hidden = hidden.has(index + 1);
  });
}



function deleteSelectedCustomView(dialog) {
  const index = Number(dialog?.querySelector("[data-data-field='customViewIndex']")?.value);
  if (!state.model?.customViews?.[index]) return;
  const removed = state.model.customViews[index];
  showMessageBox(`'${removed.name}' を削除しますか?`, {
    title: "Cht WebSheet",
    tone: "warning",
    buttons: [
      {
        label: "削除(D)",
        result: "delete",
        callback: () => confirmDeleteCustomView(index),
      },
      { label: "キャンセル", result: "cancel" },
    ],
    cancelResult: "cancel",
  });
}



function confirmDeleteCustomView(index) {
  if (!state.model?.customViews?.[index]) return;
  const workbookViewBefore = captureWorkbookViewState();
  const [removed] = state.model.customViews.splice(index, 1);
  commitWorkbookViewChange(workbookViewBefore, `ユーザー設定のビューを削除しました：${removed.name}`, `${removed.name} ユーザー設定のビュー削除`, {
    render: false,
  });
  openCustomViewsDialog();
}



function handleCustomViewListChange(event) {
  const dialog = event.target.closest?.("[data-data-dialog='custom-views']");
  if (!dialog) return;
  syncCustomViewsDialog(dialog);
}



function handleCustomViewListKeydown(event) {
  if (isElementImeCompositionEvent(event) || !["Delete", "Backspace"].includes(event.key)) return;
  const dialog = event.target.closest?.("[data-data-dialog='custom-views']");
  if (!dialog || !selectedCustomView(dialog)) return;
  event.preventDefault();
  event.stopPropagation();
  deleteSelectedCustomView(dialog);
}



function handleCustomViewListDoubleClick(event) {
  const dialog = event.target.closest?.("[data-data-dialog='custom-views']");
  if (!dialog || !selectedCustomView(dialog)) return;
  applySelectedCustomView(dialog);
}



function openMacrosDialog(options = {}) {
  closeDataDialog({ keepState: true });
  const macros = Array.isArray(state.model.macros) ? state.model.macros : [];
  const selectedIndex = Math.max(-1, Math.min(macros.length - 1, Number(options.selectedIndex ?? (macros.length ? 0 : -1))));
  const selectedMacro = selectedIndex >= 0 ? macros[selectedIndex] : null;
  const rows = macros.map((macro, index) => `<option value="${index}"${index === selectedIndex ? " selected" : ""}>${escapeHtml(macro.name || `Macro${index + 1}`)}</option>`).join("");
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog macros-dialog" role="dialog" aria-modal="true" aria-labelledby="macrosTitle" data-data-dialog="macros">
        <div class="excel-mini-titlebar" id="macrosTitle">マクロ</div>
        <div class="excel-mini-body data-dialog-body">
          <label class="excel-dialog-field"><span class="excel-dialog-caption">マクロ名(M):</span>
            <input type="text" data-data-field="macroName" data-macro-dialog-sync value="${escapeAttr(selectedMacro?.name || defaultMacroName())}" />
          </label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">マクロの一覧(A):</span>
            <select size="7" data-data-field="macroIndex" data-macro-dialog-sync>${rows}</select>
          </label>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">マクロの保存先(I):</span>
            <select data-data-field="macroScope" data-macro-dialog-sync>
              <option value="workbook" selected>作業中のブック</option>
              <option value="allOpen">開いているすべてのブック</option>
            </select>
          </label>
          <div class="data-dialog-note">VBA マクロはブック内に保持します。実行時は WebSheet が対応している安全な VBA サブセットだけを処理します。</div>
        </div>
        <div class="excel-mini-actions macro-dialog-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="macro-run">実行(R)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="macro-step">ステップ イン(S)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="macro-edit">編集(E)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="macro-create">作成(C)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="macro-delete">削除(D)</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="macro-options">オプション(O)...</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "macros" };
  const dialog = overlay[0].querySelector("[data-data-dialog='macros']");
  syncMacrosDialog(dialog);
  overlay.find("[data-data-field='macroName']").trigger("focus").trigger("select");
}



function defaultMacroName() {
  const macros = Array.isArray(state.model?.macros) ? state.model.macros : [];
  let index = macros.length + 1;
  const names = new Set(macros.map((macro) => String(macro.name || "").toLocaleLowerCase("ja-JP")));
  while (names.has(`macro${index}`.toLocaleLowerCase("ja-JP"))) index += 1;
  return `Macro${index}`;
}



function selectedMacroIndexFromDialog(dialog) {
  const value = Number(dialog?.querySelector("[data-data-field='macroIndex']")?.value);
  return Number.isInteger(value) && value >= 0 ? value : -1;
}



function macroNameFromDialog(dialog) {
  return String(dialog?.querySelector("[data-data-field='macroName']")?.value || "").trim();
}



function validateMacroName(name, existingIndex = -1) {
  const text = String(name || "").trim();
  if (!text) return { valid: false, message: "マクロ名を入力してください。" };
  if (/\s/.test(text)) return { valid: false, message: "マクロ名に空白を含めることはできません。" };
  const macros = Array.isArray(state.model?.macros) ? state.model.macros : [];
  const duplicate = macros.findIndex((macro, index) => index !== existingIndex && String(macro.name || "").toLocaleLowerCase("ja-JP") === text.toLocaleLowerCase("ja-JP"));
  if (duplicate >= 0) return { valid: false, message: "同じ名前のマクロが既に存在します。" };
  return { valid: true, name: text };
}



function syncMacrosDialog(dialog, options = {}) {
  if (!dialog) return;
  const macros = Array.isArray(state.model?.macros) ? state.model.macros : [];
  const selectedIndex = selectedMacroIndexFromDialog(dialog);
  const selectedMacro = selectedIndex >= 0 ? macros[selectedIndex] : null;
  const nameInput = dialog.querySelector("[data-data-field='macroName']");
  if (options.fromList && nameInput && selectedMacro) nameInput.value = selectedMacro.name || `Macro${selectedIndex + 1}`;
  const hasSelection = Boolean(selectedMacro);
  const name = macroNameFromDialog(dialog);
  const validNewName = validateMacroName(name, selectedIndex).valid;
  ["macro-run", "macro-step", "macro-edit", "macro-delete", "macro-options"].forEach((action) => {
    const button = dialog.querySelector(`[data-data-dialog-action='${action}']`);
    if (button) button.disabled = !hasSelection;
  });
  const createButton = dialog.querySelector("[data-data-dialog-action='macro-create']");
  if (createButton) createButton.disabled = !validNewName || (hasSelection && selectedMacro?.name === name);
}



function handleMacroDialogSync(event) {
  const dialog = event.target.closest?.("[data-data-dialog='macros']");
  if (!dialog) return;
  syncMacrosDialog(dialog, { fromList: event.target.matches("[data-data-field='macroIndex']") });
}



function handleMacroListDoubleClick(event) {
  const dialog = event.target.closest?.("[data-data-dialog='macros']");
  if (!dialog || selectedMacroIndexFromDialog(dialog) < 0) return;
  runMacroFromDialog(dialog);
}



function showUnsupportedMacroRunMessage() {
  showMessageBox("このマクロには、WebSheet がまだ実行できない VBA ステートメントが含まれています。編集またはステップ インでコードを確認できます。", { title: "Cht WebSheet" });
}



function runMacroFromDialog(dialog) {
  const selected = selectedMacroFromDialog(dialog);
  if (!selected) return;
  closeDataDialog();
  runWorkbookMacro(selected.macro, selected.index);
}



function runWorkbookMacro(macro, macroIndex = -1) {
  if (!state.model || !macro) return;
  const statements = parseVbaMacroStatements(macro.code || "");
  if (!statements.length) {
    showUnsupportedMacroRunMessage();
    return;
  }
  executeWorkbookMacroStatements(macro, macroIndex, statements, {
    historyLabel: `マクロ実行 ${macro.name || `Macro${macroIndex + 1}`}`,
    statusLabel: macro.name || "マクロ",
    updateLastRun: true,
  });
}



function executeWorkbookMacroStatements(macro, macroIndex = -1, statements = [], options = {}) {
  if (!state.model || !Array.isArray(statements) || !statements.length) {
    return {
      activeSheetIndex: state.activeSheetIndex,
      messages: [],
      unsupported: [],
      touchedSheetIndexes: new Set(),
      vbaClipboard: null,
      writeCount: 0,
    };
  }
  const sheetIndexes = state.model.sheets.map((_sheet, index) => index);
  const sheetsBefore = captureSheetsForHistory(sheetIndexes);
  const macroStateBefore = captureWorkbookMacroState();
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges();
  const context = {
    activeSheetIndex: state.activeSheetIndex,
    messages: [],
    unsupported: [],
    touchedSheetIndexes: new Set(),
    vbaClipboard: null,
    variables: Object.create(null),
    writeCount: 0,
  };

  executeVbaMacroStatementsSequence(statements, context);
  const storedMacro = state.model.macros?.[macroIndex];
  if (storedMacro && options.updateLastRun !== false) storedMacro.lastRunAt = new Date().toISOString();

  const touched = context.touchedSheetIndexes.size > 0;
  if (touched) {
    state.hf = buildFormulaEngine(state.model);
  }
  const selectionAfter = activeSelectionRange();
  const selectionAfterRanges = activeSelectionRanges();
  if (touched || storedMacro) {
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetIndexes),
      macroStateBefore,
      macroStateAfter: captureWorkbookMacroState(),
      selectionBefore,
      selectionAfter,
      selectionBeforeRanges,
      selectionAfterRanges,
      label: options.historyLabel || `マクロ実行 ${macro.name || `Macro${macroIndex + 1}`}`,
    });
  }
  renderWorkbook();
  updateSelectionUi();
  if (state.selected) updateFormulaBarForCell(state.selected.row, state.selected.col);
  scheduleWorkbookPersistence();

  const unsupportedText = context.unsupported.length
    ? `\n\n実行できなかったステートメント:\n${context.unsupported.slice(0, 5).join("\n")}`
    : "";
  const statusLabel = options.statusLabel || macro.name || "マクロ";
  setStatus(`${statusLabel} を実行しました。${context.writeCount ? `${context.writeCount} 個のセルを更新しました。` : ""}`.trim());
  if (options.showMessages !== false && context.messages.length) {
    showMessageBox(`${context.messages.join("\n")}${unsupportedText}`, { title: "Cht WebSheet" });
  } else if (options.showMessages !== false && context.unsupported.length) {
    showMessageBox(`対応しているステートメントを実行しました。${unsupportedText}`, { title: "Cht WebSheet" });
  }
  return context;
}



function parseVbaMacroStatements(code) {
  const normalized = String(code || "")
    .replace(/\\r\\n/g, "\n")
    .replace(/\\n/g, "\n")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n");
  const statements = [];
  let continuation = "";
  normalized.split("\n").forEach((rawLine) => {
    const stripped = stripVbaLineComment(rawLine).trim();
    if (!stripped) return;
    const continued = /_\s*$/.test(stripped);
    const fragment = stripped.replace(/_\s*$/, "").trim();
    continuation = continuation ? `${continuation} ${fragment}` : fragment;
    if (continued) return;
    splitOutsideVbaStrings(continuation, ":").forEach((statement) => {
      const text = statement.trim();
      if (text && !isIgnoredVbaStatement(text)) statements.push(text);
    });
    continuation = "";
  });
  if (continuation && !isIgnoredVbaStatement(continuation)) statements.push(continuation);
  return statements;
}



function stripVbaLineComment(line) {
  let inString = false;
  const text = String(line || "");
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === '"') {
      if (inString && text[index + 1] === '"') {
        index += 1;
      } else {
        inString = !inString;
      }
    } else if (!inString && char === "'") {
      return text.slice(0, index);
    }
  }
  return text;
}



function splitOutsideVbaStrings(text, delimiter) {
  const parts = [];
  let inString = false;
  let start = 0;
  const source = String(text || "");
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char === '"') {
      if (inString && source[index + 1] === '"') {
        index += 1;
      } else {
        inString = !inString;
      }
    } else if (!inString && char === delimiter && !(delimiter === ":" && source[index + 1] === "=")) {
      parts.push(source.slice(start, index));
      start = index + 1;
    }
  }
  parts.push(source.slice(start));
  return parts;
}



function splitVbaTopLevelArguments(text) {
  const parts = [];
  const source = String(text || "");
  let inString = false;
  let depth = 0;
  let start = 0;
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char === '"') {
      if (inString && source[index + 1] === '"') {
        index += 1;
      } else {
        inString = !inString;
      }
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
    if (char === "," && depth === 0) {
      parts.push(source.slice(start, index).trim());
      start = index + 1;
    }
  }
  parts.push(source.slice(start).trim());
  return parts;
}



function isIgnoredVbaStatement(statement) {
  return /^(?:Option\s+Explicit|(?:Private\s+|Public\s+)?Sub\b.*|End\s+Sub|Dim\b.*|Set\b.*|With\b.*|End\s+With)$/i.test(String(statement || "").trim());
}



function executeVbaMacroStatementsSequence(statements, context) {
  const source = Array.isArray(statements) ? statements : [];
  for (let index = 0; index < source.length; index += 1) {
    const statement = String(source[index] || "").trim();
    const ifStatement = parseVbaIfStatement(statement);
    if (ifStatement) {
      if (ifStatement.inlineStatement) {
        const signal = executeVbaInlineIfStatement(ifStatement, context);
        if (signal) return signal;
        continue;
      }
      const block = collectVbaIfBlock(source, index + 1, ifStatement.condition);
      if (!block) {
        context.unsupported.push(statement);
        continue;
      }
      const branch = block.branches.find((item) => item.elseBranch || vbaTruthy(evaluateVbaExpression(item.condition, context)));
      if (branch) {
        const signal = executeVbaMacroStatementsSequence(branch.statements, context);
        if (signal) return signal;
      }
      index = block.endIndex;
      continue;
    }
    if (/^(?:ElseIf\b|Else\b|End\s+If\b)/i.test(statement)) continue;
    const selectMatch = statement.match(/^Select\s+Case\s+([\s\S]+)$/i);
    if (selectMatch) {
      const block = collectVbaSelectCaseBlock(source, index + 1);
      if (!block) {
        context.unsupported.push(statement);
        continue;
      }
      const selectorValue = evaluateVbaExpression(selectMatch[1], context);
      const branch = block.branches.find((item) => item.elseBranch || vbaCaseClauseMatches(selectorValue, item.clause, context));
      if (branch) {
        const signal = executeVbaMacroStatementsSequence(branch.statements, context);
        if (signal) return signal;
      }
      index = block.endIndex;
      continue;
    }
    if (/^(?:Case\b|End\s+Select\b)/i.test(statement)) continue;
    const doMatch = statement.match(/^Do(?:\s+(While|Until)\s+([\s\S]+))?$/i);
    if (doMatch) {
      const block = collectVbaDoLoopBlock(source, index + 1);
      if (!block) {
        context.unsupported.push(statement);
        continue;
      }
      const signal = executeVbaDoLoop(doMatch[1], doMatch[2], block.loopMode, block.loopCondition, block.statements, context);
      if (signal && signal !== "exitDo") return signal;
      index = block.endIndex;
      continue;
    }
    if (/^Loop(?:\s+(?:While|Until)\b[\s\S]*)?$/i.test(statement)) continue;
    const forEachMatch = statement.match(/^For\s+Each\s+([A-Za-z_][A-Za-z0-9_]*)\s+In\s+([\s\S]+)$/i);
    if (forEachMatch) {
      const nextIndex = matchingVbaNextIndex(source, index + 1);
      if (nextIndex < 0) {
        context.unsupported.push(statement);
        continue;
      }
      const variableKey = vbaVariableKey(forEachMatch[1]);
      const items = vbaIterableValues(forEachMatch[2], context);
      const body = source.slice(index + 1, nextIndex);
      for (const item of items) {
        context.variables[variableKey] = item;
        const signal = executeVbaMacroStatementsSequence(body, context);
        if (signal === "exitFor") break;
        if (signal) return signal;
      }
      index = nextIndex;
      continue;
    }
    const forMatch = statement.match(/^For\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([\s\S]+?)\s+To\s+([\s\S]+?)(?:\s+Step\s+([\s\S]+))?$/i);
    if (!forMatch) {
      if (!/^Next(?:\s+[A-Za-z_][A-Za-z0-9_]*)?$/i.test(statement)) {
        const signal = executeVbaMacroStatement(statement, context);
        if (signal) return signal;
      }
      continue;
    }
    const nextIndex = matchingVbaNextIndex(source, index + 1);
    if (nextIndex < 0) {
      context.unsupported.push(statement);
      continue;
    }
    const variableName = forMatch[1];
    const variableKey = vbaVariableKey(variableName);
    const start = Number(evaluateVbaExpression(forMatch[2], context));
    const end = Number(evaluateVbaExpression(forMatch[3], context));
    const step = forMatch[4] == null || String(forMatch[4]).trim() === "" ? 1 : Number(evaluateVbaExpression(forMatch[4], context));
    if (!Number.isFinite(start) || !Number.isFinite(end) || !Number.isFinite(step) || step === 0) {
      context.unsupported.push(statement);
      index = nextIndex;
      continue;
    }
    const body = source.slice(index + 1, nextIndex);
    let guard = 0;
    for (let value = start; step > 0 ? value <= end : value >= end; value += step) {
      context.variables[variableKey] = value;
      const signal = executeVbaMacroStatementsSequence(body, context);
      if (signal === "exitFor") break;
      if (signal) return signal;
      guard += 1;
      if (guard > 10000) {
        context.unsupported.push(`${statement} ... Next`);
        break;
      }
    }
    index = nextIndex;
  }
  return null;
}



function parseVbaIfStatement(statement) {
  const match = String(statement || "").trim().match(/^If\s+([\s\S]+?)\s+Then(?:\s+([\s\S]+))?$/i);
  if (!match) return null;
  return {
    condition: match[1].trim(),
    inlineStatement: String(match[2] || "").trim(),
  };
}



function executeVbaInlineIfStatement(ifStatement, context) {
  const parts = splitVbaTopLevelWord(ifStatement.inlineStatement, "Else");
  const thenText = parts[0] || "";
  const elseText = parts.length > 1 ? parts.slice(1).join(" Else ") : "";
  const activeText = vbaTruthy(evaluateVbaExpression(ifStatement.condition, context)) ? thenText : elseText;
  for (const statement of splitOutsideVbaStrings(activeText, ":")
    .map((statement) => statement.trim())
    .filter(Boolean)) {
    const signal = executeVbaMacroStatement(statement, context);
    if (signal) return signal;
  }
  return null;
}



function collectVbaIfBlock(statements, startIndex, firstCondition) {
  const branches = [{ condition: firstCondition, statements: [], elseBranch: false }];
  let current = branches[0];
  let nestedDepth = 0;
  for (let index = startIndex; index < statements.length; index += 1) {
    const statement = String(statements[index] || "").trim();
    const nestedIf = parseVbaIfStatement(statement);
    if (nestedIf && !nestedIf.inlineStatement) {
      nestedDepth += 1;
      current.statements.push(statement);
      continue;
    }
    if (/^End\s+If$/i.test(statement)) {
      if (nestedDepth === 0) return { branches, endIndex: index };
      nestedDepth -= 1;
      current.statements.push(statement);
      continue;
    }
    if (nestedDepth === 0) {
      const elseIf = statement.match(/^ElseIf\s+([\s\S]+?)\s+Then\s*$/i);
      if (elseIf) {
        current = { condition: elseIf[1].trim(), statements: [], elseBranch: false };
        branches.push(current);
        continue;
      }
      if (/^Else\s*$/i.test(statement)) {
        current = { condition: "", statements: [], elseBranch: true };
        branches.push(current);
        continue;
      }
    }
    current.statements.push(statement);
  }
  return null;
}



function collectVbaSelectCaseBlock(statements, startIndex) {
  const branches = [];
  let current = null;
  let nestedDepth = 0;
  for (let index = startIndex; index < statements.length; index += 1) {
    const statement = String(statements[index] || "").trim();
    if (/^Select\s+Case\s+/i.test(statement)) {
      nestedDepth += 1;
      if (current) current.statements.push(statement);
      continue;
    }
    if (/^End\s+Select$/i.test(statement)) {
      if (nestedDepth === 0) return { branches, endIndex: index };
      nestedDepth -= 1;
      if (current) current.statements.push(statement);
      continue;
    }
    if (nestedDepth === 0) {
      const caseMatch = statement.match(/^Case\s+([\s\S]+)$/i);
      if (caseMatch) {
        current = {
          clause: caseMatch[1].trim(),
          elseBranch: /^Else$/i.test(caseMatch[1].trim()),
          statements: [],
        };
        branches.push(current);
        continue;
      }
    }
    if (current) current.statements.push(statement);
  }
  return null;
}



function vbaCaseClauseMatches(selectorValue, clause, context) {
  const parts = splitOutsideVbaStrings(clause, ",").map((part) => part.trim()).filter(Boolean);
  return parts.some((part) => {
    const isMatch = part.match(/^Is\s*(<=|>=|<>|=|<|>)\s*([\s\S]+)$/i);
    if (isMatch) return compareVbaExpressionValues(selectorValue, evaluateVbaExpression(isMatch[2], context), isMatch[1]);
    const operatorMatch = part.match(/^(<=|>=|<>|=|<|>)\s*([\s\S]+)$/);
    if (operatorMatch) return compareVbaExpressionValues(selectorValue, evaluateVbaExpression(operatorMatch[2], context), operatorMatch[1]);
    const rangeMatch = part.match(/^([\s\S]+?)\s+To\s+([\s\S]+)$/i);
    if (rangeMatch) {
      const value = vbaComparableValue(selectorValue);
      const min = vbaComparableValue(evaluateVbaExpression(rangeMatch[1], context));
      const max = vbaComparableValue(evaluateVbaExpression(rangeMatch[2], context));
      return value >= min && value <= max;
    }
    return compareVbaExpressionValues(selectorValue, evaluateVbaExpression(part, context), "=");
  });
}



function collectVbaDoLoopBlock(statements, startIndex) {
  const body = [];
  let nestedDepth = 0;
  for (let index = startIndex; index < statements.length; index += 1) {
    const statement = String(statements[index] || "").trim();
    if (/^Do(?:\s+(?:While|Until)\b[\s\S]*)?$/i.test(statement)) {
      nestedDepth += 1;
      body.push(statement);
      continue;
    }
    const loopMatch = statement.match(/^Loop(?:\s+(While|Until)\s+([\s\S]+))?$/i);
    if (loopMatch) {
      if (nestedDepth === 0) {
        return {
          statements: body,
          endIndex: index,
          loopMode: loopMatch[1] || "",
          loopCondition: loopMatch[2] || "",
        };
      }
      nestedDepth -= 1;
      body.push(statement);
      continue;
    }
    body.push(statement);
  }
  return null;
}



function executeVbaDoLoop(preMode, preCondition, postMode, postCondition, body, context) {
  let guard = 0;
  while (true) {
    if (preMode && !vbaDoLoopConditionAllows(preMode, preCondition, context)) break;
    const signal = executeVbaMacroStatementsSequence(body, context);
    if (signal === "exitDo") return null;
    if (signal) return signal;
    guard += 1;
    if (guard > 10000) {
      context.unsupported.push("Do ... Loop");
      return null;
    }
    if (postMode && !vbaDoLoopConditionAllows(postMode, postCondition, context)) break;
    if (!preMode && !postMode) continue;
  }
  return null;
}



function vbaDoLoopConditionAllows(mode, condition, context) {
  const truthy = vbaTruthy(evaluateVbaExpression(condition, context));
  return /^until$/i.test(String(mode || "")) ? !truthy : truthy;
}



function matchingVbaNextIndex(statements, startIndex) {
  let depth = 0;
  for (let index = startIndex; index < statements.length; index += 1) {
    const statement = String(statements[index] || "").trim();
    if (/^For\s+/i.test(statement)) depth += 1;
    else if (/^Next(?:\s+[A-Za-z_][A-Za-z0-9_]*)?$/i.test(statement)) {
      if (depth === 0) return index;
      depth -= 1;
    }
  }
  return -1;
}



function executeVbaMacroStatement(statement, context) {
  const text = String(statement || "").trim();
  if (!text) return;
  if (/^Exit\s+For$/i.test(text)) return "exitFor";
  if (/^Exit\s+Do$/i.test(text)) return "exitDo";
  if (/^Exit\s+Sub$/i.test(text)) return "exitSub";
  const msgBox = parseVbaMsgBoxStatement(text, context);
  if (msgBox) {
    context.messages.push(msgBox.message);
    return;
  }
  if (executeVbaSelectStatement(text, context)) return;
  if (executeVbaMethodStatement(text, context)) return;
  if (executeVbaAssignmentStatement(text, context)) return;
  context.unsupported.push(text);
  return null;
}



function parseVbaMsgBoxStatement(statement, context) {
  const match = String(statement || "").match(/^MsgBox(?:\s+|\s*\()([\s\S]*?)\)?$/i);
  if (!match) return null;
  const args = splitOutsideVbaStrings(match[1], ",").map((item) => item.trim()).filter(Boolean);
  const message = String(evaluateVbaExpression(args[0] || "\"\"", context) ?? "");
  return { message };
}



function executeVbaSelectStatement(statement, context) {
  const direct = String(statement || "").match(/^(.+?)\.Select$/i);
  const gotoMatch = String(statement || "").match(/^Application\.Goto\s+(.+)$/i);
  const targetText = direct?.[1] || gotoMatch?.[1] || "";
  if (!targetText) return false;
  const target = parseVbaRangeTarget(targetText, context);
  if (!target) return false;
  selectVbaRangeTarget(target);
  return true;
}



function executeVbaAssignmentStatement(statement, context) {
  const parts = splitOutsideVbaStrings(statement, "=");
  if (parts.length < 2) return false;
  const targetText = parts.shift();
  const expression = parts.join("=");
  const value = evaluateVbaExpression(expression, context);
  const target = parseVbaRangeTarget(targetText, context);
  if (!target) {
    const variableMatch = String(targetText || "").trim().match(/^([A-Za-z_][A-Za-z0-9_]*)$/);
    if (!variableMatch) return false;
    context.variables[vbaVariableKey(variableMatch[1])] = value;
    return true;
  }
  if (!applyVbaRangePropertyAssignment(target, value, context)) {
    context.unsupported.push(statement);
  }
  return true;
}



function vbaVariableKey(name) {
  return String(name || "").trim().toLocaleLowerCase("en-US");
}



function parseVbaRangeTarget(targetText, context) {
  let text = String(targetText || "").trim();
  let sheetIndex = state.activeSheetIndex;
  text = text.replace(/^(?:Application\.)?(?:ActiveSheet|ThisWorkbook\.ActiveSheet)\./i, "");
  const sheetMatch = text.match(/^(?:(?:ThisWorkbook|ActiveWorkbook)\.)?(?:Worksheets?|Sheets)\(\s*(?:"((?:[^"]|"")*)"|(\d+))\s*\)\.(.+)$/i);
  if (sheetMatch) {
    const sheetName = sheetMatch[1]?.replace(/""/g, '"');
    const sheetNumber = Number(sheetMatch[2]);
    const index = sheetName
      ? state.model?.sheets?.findIndex((sheet) => sheet.name === sheetName)
      : Math.max(0, sheetNumber - 1);
    if (index == null || index < 0) return null;
    sheetIndex = index;
    text = sheetMatch[3].trim();
  }
  const variableTargetMatch = text.match(/^([A-Za-z_][A-Za-z0-9_]*)(?:\.(.+))?$/);
  if (variableTargetMatch) {
    const variableTarget = cloneVbaRangeHandle(context?.variables?.[vbaVariableKey(variableTargetMatch[1])]);
    if (variableTarget) return applyVbaRangePostfix(variableTarget, variableTargetMatch[2], context);
  }
  const rowsMatch = text.match(/^Rows\(\s*(?:"(\d+)(?::(\d+))?"|(\d+))\s*\)(?:\.(.+))?$/i);
  if (rowsMatch) {
    const startRow = Number(rowsMatch[1] || rowsMatch[3]);
    const endRow = Number(rowsMatch[2] || rowsMatch[1] || rowsMatch[3]);
    const sheet = state.model?.sheets?.[sheetIndex];
    const right = Math.max(1, Number(sheet?.colCount || 1));
    return applyVbaRangePostfix({
      sheetIndex,
      range: { sheetIndex, top: Math.min(startRow, endRow), left: 1, bottom: Math.max(startRow, endRow), right, fullRow: true },
    }, rowsMatch[4], context);
  }
  const columnsMatch = text.match(/^Columns\(\s*(?:"([A-Z]+)(?::([A-Z]+))?"|(\d+))\s*\)(?:\.(.+))?$/i);
  if (columnsMatch) {
    const startCol = columnsMatch[1] ? columnNumber(columnsMatch[1]) : Number(columnsMatch[3]);
    const endCol = columnsMatch[2] ? columnNumber(columnsMatch[2]) : startCol;
    const sheet = state.model?.sheets?.[sheetIndex];
    const bottom = Math.max(1, Number(sheet?.rowCount || 1));
    return applyVbaRangePostfix({
      sheetIndex,
      range: { sheetIndex, top: 1, left: Math.min(startCol, endCol), bottom, right: Math.max(startCol, endCol), fullColumn: true },
    }, columnsMatch[4], context);
  }
  const rangeMatch = text.match(/^Range\(\s*"([^"]+)"\s*\)(?:\.(.+))?$/i);
  if (rangeMatch) {
    const range = decodeRange(rangeMatch[1].replace(/\$/g, ""));
    if (!range) return null;
    return applyVbaRangePostfix({ sheetIndex, range: { ...range, sheetIndex } }, rangeMatch[2], context);
  }
  const rangeCellsMatch = text.match(/^Range\(\s*(Cells\([\s\S]*?\))\s*,\s*(Cells\([\s\S]*?\))\s*\)(?:\.(.+))?$/i);
  if (rangeCellsMatch) {
    const start = parseVbaCellsPoint(rangeCellsMatch[1], sheetIndex, context);
    const end = parseVbaCellsPoint(rangeCellsMatch[2], sheetIndex, context);
    if (!start || !end) return null;
    return applyVbaRangePostfix({
      sheetIndex,
      range: {
        sheetIndex,
        top: Math.min(start.row, end.row),
        left: Math.min(start.col, end.col),
        bottom: Math.max(start.row, end.row),
        right: Math.max(start.col, end.col),
      },
    }, rangeCellsMatch[3], context);
  }
  const cellsMatch = text.match(/^Cells\(([\s\S]*)\)(?:\.(.+))?$/i);
  if (cellsMatch) {
    const point = parseVbaCellsPoint(`Cells(${cellsMatch[1]})`, sheetIndex, context);
    if (!point) return null;
    return applyVbaRangePostfix({ sheetIndex, range: { sheetIndex, top: point.row, left: point.col, bottom: point.row, right: point.col } }, cellsMatch[2], context);
  }
  const activeMatch = text.match(/^ActiveCell(?:\.(.+))?$/i);
  if (activeMatch) {
    const point = state.selected?.sheetIndex === state.activeSheetIndex ? state.selected : { sheetIndex, row: 1, col: 1 };
    return applyVbaRangePostfix({ sheetIndex: point.sheetIndex, range: { sheetIndex: point.sheetIndex, top: point.row, left: point.col, bottom: point.row, right: point.col } }, activeMatch[1], context);
  }
  const selectionMatch = text.match(/^Selection(?:\.(.+))?$/i);
  if (selectionMatch) {
    const range = activeSelectionRange();
    if (!range) return null;
    return applyVbaRangePostfix({ sheetIndex: state.activeSheetIndex, range: { ...range, sheetIndex: state.activeSheetIndex } }, selectionMatch[1], context);
  }
  return null;
}



function parseVbaCellsPoint(cellsText, sheetIndex, context) {
  const match = String(cellsText || "").trim().match(/^Cells\(([\s\S]*)\)$/i);
  if (!match) return null;
  const args = splitVbaTopLevelArguments(match[1]);
  if (args.length < 2) return null;
  const row = Math.max(1, Math.round(Number(evaluateVbaExpression(args[0], context)) || 0));
  const col = vbaColumnIndexFromExpression(args[1], context);
  if (!row || !col) return null;
  const sheet = state.model?.sheets?.[sheetIndex];
  ensureSheetSize(sheet, row, col);
  return { row, col };
}



function vbaColumnIndexFromExpression(expression, context) {
  const value = evaluateVbaExpression(expression, context);
  if (typeof value === "number" && Number.isFinite(value)) return Math.max(1, Math.round(value));
  const text = String(value ?? "").trim().replace(/^"|"$/g, "");
  if (/^[A-Z]+$/i.test(text)) return columnNumber(text);
  const numeric = Number(text);
  return Number.isFinite(numeric) ? Math.max(1, Math.round(numeric)) : 0;
}



function vbaRangeHandle(target) {
  if (!target?.range) return null;
  return {
    __webSheetVbaRangeTarget: true,
    sheetIndex: target.sheetIndex,
    range: { ...target.range, sheetIndex: target.sheetIndex },
    property: target.property || "Value",
  };
}



function cloneVbaRangeHandle(value) {
  if (!value?.__webSheetVbaRangeTarget || !value.range) return null;
  return {
    sheetIndex: value.sheetIndex,
    range: { ...value.range, sheetIndex: value.sheetIndex },
    property: value.property || "Value",
  };
}



function vbaIterableValues(expression, context) {
  const target = parseVbaRangeTarget(expression, context);
  if (target) {
    const items = [];
    forEachCellInRange(target.range, (row, col) => {
      items.push(vbaRangeHandle({
        sheetIndex: target.sheetIndex,
        range: { sheetIndex: target.sheetIndex, top: row, left: col, bottom: row, right: col },
      }));
    });
    return items;
  }
  const value = evaluateVbaExpression(expression, context);
  if (value?.__webSheetVbaArray) return [...value.values];
  if (Array.isArray(value)) return value;
  return value == null ? [] : [value];
}



function applyVbaRangePostfix(target, postfix, context) {
  if (!target) return null;
  let remaining = String(postfix || "").trim();
  if (!remaining) {
    target.property = "Value";
    return target;
  }
  while (remaining) {
    const offsetMatch = remaining.match(/^Offset\s*\(([\s\S]*?)\)(?:\.(.*))?$/i);
    if (offsetMatch) {
      applyVbaRangeOffset(target, offsetMatch[1], context);
      remaining = String(offsetMatch[2] || "").trim();
      continue;
    }
    const resizeMatch = remaining.match(/^Resize\s*\(([\s\S]*?)\)(?:\.(.*))?$/i);
    if (resizeMatch) {
      applyVbaRangeResize(target, resizeMatch[1], context);
      remaining = String(resizeMatch[2] || "").trim();
      continue;
    }
    const entireRowMatch = remaining.match(/^EntireRow(?:\.(.*))?$/i);
    if (entireRowMatch) {
      applyVbaRangeEntireAxis(target, "row");
      remaining = String(entireRowMatch[1] || "").trim();
      if (!remaining) remaining = "Value";
      continue;
    }
    const entireColumnMatch = remaining.match(/^EntireColumn(?:\.(.*))?$/i);
    if (entireColumnMatch) {
      applyVbaRangeEntireAxis(target, "column");
      remaining = String(entireColumnMatch[1] || "").trim();
      if (!remaining) remaining = "Value";
      continue;
    }
    target.property = normalizeVbaRangeProperty(remaining);
    return target;
  }
  target.property = "Value";
  return target;
}



function applyVbaRangeOffset(target, argsText, context) {
  const args = splitVbaTopLevelArguments(argsText);
  const rowOffset = Math.round(Number(evaluateVbaExpression(args[0] || "0", context)) || 0);
  const colOffset = Math.round(Number(evaluateVbaExpression(args[1] || "0", context)) || 0);
  const height = Math.max(1, target.range.bottom - target.range.top + 1);
  const width = Math.max(1, target.range.right - target.range.left + 1);
  target.range.top = Math.max(1, target.range.top + rowOffset);
  target.range.left = Math.max(1, target.range.left + colOffset);
  target.range.bottom = target.range.top + height - 1;
  target.range.right = target.range.left + width - 1;
  delete target.range.fullRow;
  delete target.range.fullColumn;
}



function applyVbaRangeResize(target, argsText, context) {
  const args = splitVbaTopLevelArguments(argsText);
  const currentHeight = Math.max(1, target.range.bottom - target.range.top + 1);
  const currentWidth = Math.max(1, target.range.right - target.range.left + 1);
  const rowCount = args[0] == null || String(args[0]).trim() === ""
    ? currentHeight
    : Math.max(1, Math.round(Number(evaluateVbaExpression(args[0], context)) || currentHeight));
  const colCount = args[1] == null || String(args[1]).trim() === ""
    ? currentWidth
    : Math.max(1, Math.round(Number(evaluateVbaExpression(args[1], context)) || currentWidth));
  target.range.bottom = target.range.top + rowCount - 1;
  target.range.right = target.range.left + colCount - 1;
  delete target.range.fullRow;
  delete target.range.fullColumn;
}



function applyVbaRangeEntireAxis(target, axis) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return;
  if (axis === "row") {
    target.range.left = 1;
    target.range.right = Math.max(1, Number(sheet.colCount || 1));
    target.range.fullRow = true;
    delete target.range.fullColumn;
    return;
  }
  target.range.top = 1;
  target.range.bottom = Math.max(1, Number(sheet.rowCount || 1));
  target.range.fullColumn = true;
  delete target.range.fullRow;
}



function normalizeVbaRangeProperty(property) {
  const text = String(property || "Value").trim();
  return text || "Value";
}



function executeVbaMethodStatement(statement, context) {
  const match = String(statement || "").trim().match(/^(.+?)\.(ClearContents|ClearFormats|Clear|Copy|Cut|PasteSpecial|Merge|UnMerge|Delete|Insert)(?:\s+([\s\S]*))?$/i);
  if (!match) return false;
  const target = parseVbaRangeTarget(match[1], context);
  if (!target) return false;
  const method = String(match[2] || "").toLowerCase();
  const args = String(match[3] || "").trim();
  if (method === "copy" || method === "cut") {
    executeVbaCopyCutStatement(target, method, args, context);
  } else if (method === "pastespecial") {
    executeVbaPasteSpecialStatement(target, context);
  } else if (method === "merge") {
    mergeVbaRangeTarget(target, false, context);
  } else if (method === "unmerge") {
    unmergeVbaRangeTarget(target, context);
  } else if (method === "delete" || method === "insert") {
    applyVbaRangeStructureMethod(target, method, args, context);
  } else {
    clearVbaRangeTarget(target, match[2], context);
  }
  return true;
}



function applyVbaRangePropertyAssignment(target, value, context) {
  const property = String(target?.property || "Value").toLowerCase();
  if (["value", "value2", "formula", "formula2", "formulalocal", "formula2local", "text"].includes(property)) {
    writeVbaValueToRange(target, value, context);
    return true;
  }
  if (["formular1c1", "formula2r1c1", "formular1c1local", "formula2r1c1local"].includes(property)) {
    writeVbaFormulaR1C1ToRange(target, value, context);
    return true;
  }
  if (property === "mergecells") {
    if (vbaTruthy(value)) {
      mergeVbaRangeTarget(target, false, context);
    } else {
      unmergeVbaRangeTarget(target, context);
    }
    return true;
  }
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return false;
  if (property === "columnwidth") {
    applyVbaColumnWidth(target, value, context);
    return true;
  }
  if (property === "rowheight") {
    applyVbaRowHeight(target, value, context);
    return true;
  }
  if (property === "hidden") {
    applyVbaHidden(target, value, context);
    return true;
  }
  const borderAssignment = parseVbaBordersProperty(property);
  if (borderAssignment) {
    applyVbaBordersPropertyAssignment(target, borderAssignment, value, context);
    return true;
  }
  const fillColor = property === "interior.color" ? vbaColorToCss(value) : "";
  const fontColor = property === "font.color" ? vbaColorToCss(value) : "";
  const isNoColor = isVbaNoneValue(value);
  ensureSheetSize(sheet, target.range.bottom, target.range.right);
  let changed = 0;
  forEachCellInRange(target.range, (row, col) => {
    const cell = ensureCellModel(sheet, row, col);
    const css = mutableCssObject(cell.css);
    cell.css = css;
    switch (property) {
      case "interior.color":
        if (fillColor && !isNoColor) {
          css.backgroundColor = fillColor;
          cell.hasFill = true;
        } else {
          delete css.backgroundColor;
          cell.hasFill = false;
        }
        break;
      case "interior.colorindex":
        if (isNoColor || Number(value) === 0) {
          delete css.backgroundColor;
          cell.hasFill = false;
        } else {
          const color = INDEXED_COLORS[Math.max(0, Number(value) || 0)] || "";
          if (!color) return;
          css.backgroundColor = color;
          cell.hasFill = true;
        }
        break;
      case "font.color":
        if (fontColor && !isNoColor) {
          css.color = fontColor;
          cell.hasFontColor = true;
        } else {
          delete css.color;
          cell.hasFontColor = false;
        }
        break;
      case "font.colorindex":
        if (isNoColor || Number(value) <= 0) {
          delete css.color;
          cell.hasFontColor = false;
        } else {
          const color = INDEXED_COLORS[Math.max(0, Number(value) || 0)] || "";
          if (!color) return;
          css.color = color;
          cell.hasFontColor = true;
        }
        break;
      case "font.bold":
        if (vbaTruthy(value)) css.fontWeight = "700";
        else delete css.fontWeight;
        break;
      case "font.italic":
        if (vbaTruthy(value)) css.fontStyle = "italic";
        else delete css.fontStyle;
        break;
      case "font.name":
        if (value == null || value === "") delete css.fontFamily;
        else css.fontFamily = fontFamilyForExcel(String(value));
        break;
      case "font.size": {
        const size = Number(value);
        if (Number.isFinite(size) && size > 0) css.fontSize = `${Math.max(1, Math.min(409, size))}pt`;
        break;
      }
      case "font.underline": {
        const underline = vbaUnderlineTextDecoration(value);
        if (underline) css.textDecoration = underline;
        else delete css.textDecoration;
        break;
      }
      case "font.strikethrough":
        setCssTextDecorationFlag(css, "line-through", vbaTruthy(value));
        break;
      case "numberformat":
      case "numberformatlocal":
        cell.numFmt = String(value ?? "");
        break;
      case "horizontalalignment": {
        const alignment = vbaHorizontalAlignment(value);
        if (alignment) css.textAlign = alignment;
        else delete css.textAlign;
        break;
      }
      case "verticalalignment": {
        const alignment = vbaVerticalAlignment(value);
        if (alignment) {
          css.verticalAlign = alignment;
          css.alignItems = verticalAlignToFlex(alignment);
        } else {
          delete css.verticalAlign;
          delete css.alignItems;
        }
        break;
      }
      case "wraptext":
        if (vbaTruthy(value)) css.whiteSpace = "pre-wrap";
        else delete css.whiteSpace;
        break;
      default:
        return;
    }
    storeCellModelForSheet(sheet, cellKey(row, col), cell);
    changed += 1;
  });
  if (!changed) return false;
  context.writeCount += changed;
  context.touchedSheetIndexes.add(target.sheetIndex);
  return true;
}



function setCssTextDecorationFlag(css, token, enabled) {
  const parts = String(css.textDecoration || "")
    .split(/\s+/)
    .map((part) => part.trim())
    .filter((part) => part && part !== "none");
  const set = new Set(parts);
  if (enabled) set.add(token);
  else set.delete(token);
  const next = [...set].join(" ");
  if (next) css.textDecoration = next;
  else delete css.textDecoration;
}



function applyVbaColumnWidth(target, value, context) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return;
  ensureSheetSize(sheet, target.range.bottom, target.range.right);
  const width = Math.max(0, Math.min(255, Number(value) || 0));
  for (let col = target.range.left; col <= target.range.right; col += 1) {
    const column = sheet.columns[col - 1] ||= {};
    column.width = columnWidthToPixels(width);
  }
  context.writeCount += Math.max(1, target.range.right - target.range.left + 1);
  context.touchedSheetIndexes.add(target.sheetIndex);
}



function applyVbaRowHeight(target, value, context) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return;
  ensureSheetSize(sheet, target.range.bottom, target.range.right);
  const height = Math.max(0, Math.min(409, Number(value) || 0));
  for (let row = target.range.top; row <= target.range.bottom; row += 1) {
    const rowModel = sheet.rows[row - 1] ||= {};
    rowModel.height = rowHeightToPixels(height);
  }
  context.writeCount += Math.max(1, target.range.bottom - target.range.top + 1);
  context.touchedSheetIndexes.add(target.sheetIndex);
}



function applyVbaHidden(target, value, context) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return;
  ensureSheetSize(sheet, target.range.bottom, target.range.right);
  const hidden = vbaTruthy(value);
  if (target.range.fullColumn && !target.range.fullRow) {
    for (let col = target.range.left; col <= target.range.right; col += 1) {
      const column = sheet.columns[col - 1] ||= {};
      column.hidden = hidden;
    }
    context.writeCount += Math.max(1, target.range.right - target.range.left + 1);
  } else {
    for (let row = target.range.top; row <= target.range.bottom; row += 1) {
      const rowModel = sheet.rows[row - 1] ||= {};
      rowModel.hidden = hidden;
    }
    context.writeCount += Math.max(1, target.range.bottom - target.range.top + 1);
  }
  context.touchedSheetIndexes.add(target.sheetIndex);
}



function parseVbaBordersProperty(property) {
  const match = String(property || "").match(/^borders(?:\(([^)]*)\))?\.(linestyle|color|weight)$/i);
  if (!match) return null;
  return {
    side: vbaBorderSide(match[1]),
    property: String(match[2] || "").toLowerCase(),
  };
}



function vbaBorderSide(value) {
  const text = String(value || "").trim().toLowerCase();
  if (/top/.test(text)) return "top";
  if (/bottom/.test(text)) return "bottom";
  if (/left/.test(text)) return "left";
  if (/right/.test(text)) return "right";
  return "all";
}



function applyVbaBordersPropertyAssignment(target, assignment, value, context) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return;
  ensureSheetSize(sheet, target.range.bottom, target.range.right);
  let changed = 0;
  forEachCellInRange(target.range, (row, col) => {
    const cell = ensureCellModel(sheet, row, col);
    cell.borders = mutableBordersObject(cell.borders);
    const sides = assignment.side === "all" ? CELL_BORDER_SIDES : [assignment.side];
    sides.forEach((side) => {
      if (assignment.property === "linestyle" && vbaBorderLineStyle(value) === "none") {
        delete cell.borders[side];
        return;
      }
      const border = cell.borders[side] ? { ...cell.borders[side] } : { width: 1, style: "solid", color: "#000000" };
      if (assignment.property === "linestyle") border.style = vbaBorderLineStyle(value);
      if (assignment.property === "color") border.color = vbaColorToCss(value) || border.color || "#000000";
      if (assignment.property === "weight") border.width = vbaBorderWeight(value);
      border.weight = border.width;
      cell.borders[side] = border;
    });
    storeCellModelForSheet(sheet, cellKey(row, col), cell);
    changed += 1;
  });
  context.writeCount += changed;
  context.touchedSheetIndexes.add(target.sheetIndex);
}



function vbaBorderLineStyle(value) {
  const text = String(value ?? "").trim().toLowerCase();
  if (/^(0|false|xlnone|xllinestylenone)$/.test(text)) return "none";
  if (/double/.test(text)) return "double";
  if (/dash|xl_dash|xldash/.test(text)) return "dashed";
  if (/dot|xldot/.test(text)) return "dotted";
  return "solid";
}



function vbaBorderWeight(value) {
  const text = String(value ?? "").trim().toLowerCase();
  if (/xlhairline/.test(text)) return 1;
  if (/xlthin/.test(text)) return 1;
  if (/xlmedium/.test(text)) return 2;
  if (/xlthick/.test(text)) return 3;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? Math.max(1, Math.min(6, Math.round(numeric))) : 1;
}



function executeVbaCopyCutStatement(sourceTarget, method, args, context) {
  const destinationText = vbaNamedArgument(args, "Destination") || (args && !/^\w+\s*:=/.test(args) ? args : "");
  if (destinationText) {
    const destinationTarget = parseVbaRangeTarget(destinationText, context);
    if (!destinationTarget) {
      context.unsupported.push(`${method.toUpperCase()} ${args}`);
      return;
    }
    copyVbaRangeToTarget(sourceTarget, destinationTarget, { cut: method === "cut" }, context);
    return;
  }
  context.vbaClipboard = captureVbaRangeClipboard(sourceTarget, { cut: method === "cut" });
}



function executeVbaPasteSpecialStatement(destinationTarget, context) {
  if (!context.vbaClipboard) {
    context.unsupported.push("PasteSpecial");
    return;
  }
  copyVbaRangeToTarget(context.vbaClipboard.sourceTarget, destinationTarget, {
    cut: context.vbaClipboard.cut,
    clipboard: context.vbaClipboard,
  }, context);
  if (context.vbaClipboard.cut) context.vbaClipboard = null;
}



function vbaNamedArgument(args, name) {
  const pattern = new RegExp(`(?:^|,)\\s*${name}\\s*:=\\s*([\\s\\S]+)$`, "i");
  return String(args || "").match(pattern)?.[1]?.trim() || "";
}



function captureVbaRangeClipboard(sourceTarget, options = {}) {
  const sourceSheet = state.model?.sheets?.[sourceTarget.sheetIndex];
  const sourceRange = sourceTarget.range;
  const cellEntries = [];
  forEachCellInRange(sourceRange, (row, col) => {
    const sourceCell = sourceSheet?.cells?.[cellKey(row, col)];
    const cell = cloneCellForClipboard(sourceCell);
    if (!cell) return;
    cell.row = row - sourceRange.top + 1;
    cell.col = col - sourceRange.left + 1;
    cellEntries.push(cell);
  });
  const merges = (sourceSheet?.merges || [])
    .filter((merge) => isRangeInsideRange(merge, sourceRange))
    .map((merge) => ({
      top: merge.top - sourceRange.top + 1,
      left: merge.left - sourceRange.left + 1,
      bottom: merge.bottom - sourceRange.top + 1,
      right: merge.right - sourceRange.left + 1,
    }));
  return {
    sourceTarget,
    cut: options.cut === true,
    rowCount: sourceRange.bottom - sourceRange.top + 1,
    colCount: sourceRange.right - sourceRange.left + 1,
    cellEntries,
    merges,
  };
}



function copyVbaRangeToTarget(sourceTarget, destinationTarget, options = {}, context) {
  const clipboard = options.clipboard || captureVbaRangeClipboard(sourceTarget, { cut: options.cut === true });
  const destinationSheet = state.model?.sheets?.[destinationTarget.sheetIndex];
  if (!destinationSheet) return;
  const targetRange = {
    sheetIndex: destinationTarget.sheetIndex,
    top: destinationTarget.range.top,
    left: destinationTarget.range.left,
    bottom: destinationTarget.range.top + clipboard.rowCount - 1,
    right: destinationTarget.range.left + clipboard.colCount - 1,
  };
  ensureSheetSize(destinationSheet, targetRange.bottom, targetRange.right);
  removeIntersectingMerges(destinationSheet, targetRange);
  clearCellsInRange(destinationSheet, targetRange);
  clipboard.cellEntries.forEach((sourceCell) => {
    const row = targetRange.top + sourceCell.row - 1;
    const col = targetRange.left + sourceCell.col - 1;
    writeClonedCellContent(
      destinationSheet,
      row,
      col,
      sourceCell,
      {
        clipboard: { sourceRange: sourceTarget.range },
        targetRow: row,
        targetCol: col,
        sourceOffset: { row: sourceCell.row, col: sourceCell.col },
      },
    );
  });
  clipboard.merges.forEach((merge) => {
    destinationSheet.merges.push({
      top: targetRange.top + merge.top - 1,
      left: targetRange.left + merge.left - 1,
      bottom: targetRange.top + merge.bottom - 1,
      right: targetRange.left + merge.right - 1,
    });
  });
  context.writeCount += Math.max(1, clipboard.rowCount * clipboard.colCount);
  context.touchedSheetIndexes.add(destinationTarget.sheetIndex);
  if (options.cut === true) {
    const sourceSheet = state.model?.sheets?.[sourceTarget.sheetIndex];
    if (sourceSheet) {
      clearCellsInRange(sourceSheet, sourceTarget.range);
      removeIntersectingMerges(sourceSheet, sourceTarget.range);
      context.touchedSheetIndexes.add(sourceTarget.sheetIndex);
    }
  }
}



function mergeVbaRangeTarget(target, centerContent, context) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  const range = target.range;
  if (!sheet || isSingleCellRange(range)) return;
  ensureSheetSize(sheet, range.bottom, range.right);
  removeIntersectingMerges(sheet, range);
  const master = ensureCellModel(sheet, range.top, range.left);
  if (centerContent) {
    master.css = mutableCssObject(master.css);
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
  context.writeCount += rangeCellCount(range);
  context.touchedSheetIndexes.add(target.sheetIndex);
}



function unmergeVbaRangeTarget(target, context) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return;
  const before = sheet.merges?.length || 0;
  removeIntersectingMerges(sheet, target.range);
  if ((sheet.merges?.length || 0) !== before) {
    context.writeCount += 1;
    context.touchedSheetIndexes.add(target.sheetIndex);
  }
}



function applyVbaRangeStructureMethod(target, method, args, context) {
  const sheet = state.model?.sheets?.[target.sheetIndex];
  if (!sheet) return;
  ensureSheetSize(sheet, target.range.bottom, target.range.right);
  const option = vbaStructureOption(method, args, target.range);
  if (!option) {
    context.unsupported.push(`${method} ${args}`.trim());
    return;
  }
  if (method === "insert") {
    applyInsertOperation(sheet, target.range, option);
  } else {
    applyDeleteOperation(sheet, target.range, option);
  }
  normalizeSheetModelAfterStructureChange(sheet);
  context.writeCount += rangeCellCount(target.range);
  context.touchedSheetIndexes.add(target.sheetIndex);
}



function vbaStructureOption(method, args, range) {
  if (range.fullRow) return "entire-row";
  if (range.fullColumn) return "entire-column";
  const shiftText = (vbaNamedArgument(args, "Shift") || args || "").toLowerCase();
  if (/xlshifttoleft|xltoleft|left/.test(shiftText)) return "shift-left";
  if (/xlshiftup|xlup|up/.test(shiftText)) return "shift-up";
  if (/xlshifttoright|xltoright|right/.test(shiftText)) return "shift-right";
  if (/xlshiftdown|xldown|down/.test(shiftText)) return "shift-down";
  return method === "insert" ? "shift-down" : "shift-up";
}



function rangeCellCount(range) {
  return Math.max(0, range.bottom - range.top + 1) * Math.max(0, range.right - range.left + 1);
}



function evaluateVbaExpression(expression, context) {
  let text = String(expression || "").trim();
  text = trimVbaOuterParentheses(text);
  const orParts = splitVbaTopLevelWord(text, "Or");
  if (orParts.length > 1) return orParts.some((part) => vbaTruthy(evaluateVbaExpression(part, context)));
  const andParts = splitVbaTopLevelWord(text, "And");
  if (andParts.length > 1) return andParts.every((part) => vbaTruthy(evaluateVbaExpression(part, context)));
  const notMatch = text.match(/^Not\s+([\s\S]+)$/i);
  if (notMatch) return !vbaTruthy(evaluateVbaExpression(notMatch[1], context));
  const comparison = splitVbaComparisonExpression(text);
  if (comparison) {
    return compareVbaExpressionValues(
      evaluateVbaExpression(comparison.left, context),
      evaluateVbaExpression(comparison.right, context),
      comparison.operator,
    );
  }
  const modParts = splitVbaTopLevelWord(text, "Mod");
  if (modParts.length > 1) {
    return modParts.slice(1).reduce((left, right) => {
      const leftNumber = Number(left);
      const rightNumber = Number(evaluateVbaExpression(right, context));
      return Number.isFinite(leftNumber) && Number.isFinite(rightNumber) && rightNumber !== 0 ? leftNumber % rightNumber : Number.NaN;
    }, Number(evaluateVbaExpression(modParts[0], context)));
  }
  const concatParts = splitOutsideVbaStrings(text, "&");
  if (concatParts.length > 1) {
    return concatParts.map((part) => String(evaluateVbaExpression(part, context) ?? "")).join("");
  }
  const additive = splitVbaArithmeticExpression(text, ["+", "-"]);
  if (additive) {
    const left = Number(evaluateVbaExpression(additive.left, context));
    const right = Number(evaluateVbaExpression(additive.right, context));
    if (Number.isFinite(left) && Number.isFinite(right)) return additive.operator === "+" ? left + right : left - right;
  }
  const multiplicative = splitVbaArithmeticExpression(text, ["*", "/"]);
  if (multiplicative) {
    const left = Number(evaluateVbaExpression(multiplicative.left, context));
    const right = Number(evaluateVbaExpression(multiplicative.right, context));
    if (Number.isFinite(left) && Number.isFinite(right)) return multiplicative.operator === "*" ? left * right : left / right;
  }
  const rgbMatch = text.match(/^RGB\(([\s\S]*)\)$/i);
  if (rgbMatch) {
    const parts = splitVbaTopLevelArguments(rgbMatch[1]).map((part) => Number(evaluateVbaExpression(part, context)));
    if (parts.length >= 3 && parts.every((part) => Number.isFinite(part))) {
      return rgbToHex(parts.slice(0, 3).map((part) => Math.max(0, Math.min(255, Math.round(part)))));
    }
  }
  const arrayMatch = text.match(/^Array\(([\s\S]*)\)$/i);
  if (arrayMatch) {
    return {
      __webSheetVbaArray: true,
      lowerBound: 0,
      values: splitVbaTopLevelArguments(arrayMatch[1]).map((part) => evaluateVbaExpression(part, context)),
    };
  }
  const builtInFunction = evaluateVbaBuiltInFunctionExpression(text, context);
  if (builtInFunction !== undefined) return builtInFunction;
  const worksheetFunction = evaluateVbaWorksheetFunctionExpression(text, context);
  if (worksheetFunction !== undefined) return worksheetFunction;
  const stringMatch = text.match(/^"((?:[^"]|"")*)"$/);
  if (stringMatch) return stringMatch[1].replace(/""/g, '"');
  if (/^-?\d+(?:\.\d+)?$/.test(text)) return Number(text);
  if (/^(True|False)$/i.test(text)) return /^true$/i.test(text);
  if (/^Date$/i.test(text)) return new Date();
  if (/^Now$/i.test(text)) return new Date();
  const constant = vbaConstantValue(text);
  if (constant !== undefined) return constant;
  const variable = context?.variables?.[vbaVariableKey(text)];
  if (variable !== undefined) return variable;
  const target = parseVbaRangeTarget(text, context);
  if (target) return readVbaRangeValue(target);
  return text;
}

export {
  activeWorkbookWindowShowOption,
  addCustomViewFromDialog,
  applyCustomViewSheetState,
  applyCustomViewToCurrentWorkbookWindow,
  applyHiddenIndexesToModels,
  applySelectedCustomView,
  applySheetZoomWithoutHistory,
  applyVbaBordersPropertyAssignment,
  applyVbaColumnWidth,
  applyVbaHidden,
  applyVbaRangeEntireAxis,
  applyVbaRangeOffset,
  applyVbaRangePostfix,
  applyVbaRangePropertyAssignment,
  applyVbaRangeResize,
  applyVbaRangeStructureMethod,
  applyVbaRowHeight,
  applyWorkbookWindowSheetView,
  arrangeWorkbookWindowViews,
  captureCustomView,
  captureVbaRangeClipboard,
  cloneVbaRangeHandle,
  closeViewCommandMenu,
  closeViewCommandMenuOnEscape,
  closeViewCommandMenuOnOutsidePointer,
  closeWorkbookWindowView,
  collectVbaDoLoopBlock,
  collectVbaIfBlock,
  collectVbaSelectCaseBlock,
  commandMenuAccessKeyAttr,
  computeWorkbookWindowArrangement,
  confirmDeleteCustomView,
  copyVbaRangeToTarget,
  createWorkbookWindowView,
  currentWorkbookWindowSnapshot,
  customViewActiveSheetIndex,
  customViewScrollSnapshotForSheet,
  customViewTableRestrictionMessage,
  customViewWindowRecordForCurrentWindow,
  customViewsDialogHtml,
  customViewsUnavailableMessage,
  customViewsUnavailableSummary,
  defaultCustomViewName,
  defaultMacroName,
  deleteSelectedCustomView,
  enableViewSideBySide,
  evaluateVbaExpression,
  executeVbaAssignmentStatement,
  executeVbaCopyCutStatement,
  executeVbaDoLoop,
  executeVbaInlineIfStatement,
  executeVbaMacroStatement,
  executeVbaMacroStatementsSequence,
  executeVbaMethodStatement,
  executeVbaPasteSpecialStatement,
  executeVbaSelectStatement,
  executeWorkbookMacroStatements,
  handleCustomViewListChange,
  handleCustomViewListDoubleClick,
  handleCustomViewListKeydown,
  handleMacroDialogSync,
  handleMacroListDoubleClick,
  handleViewCommandMenuAction,
  handleWorkbookWindowBarClick,
  handleWorkbookWindowCloseClick,
  handleWorkbookWindowListConfirm,
  hiddenIndexesForModels,
  hideWorkbookWindow,
  isIgnoredVbaStatement,
  macroNameFromDialog,
  matchingVbaNextIndex,
  menuAccessKeyFromLabel,
  mergeVbaRangeTarget,
  mutableWorkbookWindowRecordById,
  nearestPageBreakGuideAtPoint,
  normalizeVbaRangeProperty,
  normalizedViewSideBySidePair,
  openAddCustomViewDialog,
  openArrangeAllDialog,
  openCustomViewsDialog,
  openMacrosDialog,
  openPageBreakContextMenu,
  openPageBreakContextMenuAt,
  openPageBreakContextMenuFromPoint,
  openSideBySideWindowDialog,
  openSwitchWindowsDialog,
  openSwitchWindowsMenu,
  openUnhideWorkbookWindowDialog,
  openViewCommandMenu,
  openViewFreezeMenu,
  pageBreakPositionLabel,
  parseVbaBordersProperty,
  parseVbaCellsPoint,
  parseVbaIfStatement,
  parseVbaMacroStatements,
  parseVbaMsgBoxStatement,
  parseVbaRangeTarget,
  rangeCellCount,
  repairViewSideBySideState,
  resetWindowPosition,
  restoreCustomViewActiveSheetView,
  restoreCustomViewSelection,
  restoreWorkbookWindowScroll,
  restoreWorkbookWindowScrollStable,
  restoreWorkbookWindowSelection,
  restoreWorksheetFocusAfterCommandPopup,
  restoreWorksheetFocusAfterViewCommand,
  runMacroFromDialog,
  runWorkbookMacro,
  saveActiveWindowViewState,
  saveCustomViewSnapshot,
  savedCustomViewSheetForSheet,
  scheduleWorksheetFocusAfterCommandPopup,
  selectedCustomView,
  selectedMacroIndexFromDialog,
  setCssTextDecorationFlag,
  showCustomViewTableRestrictionIfNeeded,
  showCustomViewsUnavailableIfNeeded,
  showUnsupportedMacroRunMessage,
  splitOutsideVbaStrings,
  splitVbaTopLevelArguments,
  storedVisibleWorkbookWindowRecords,
  storedWorkbookWindowRecordById,
  storedWorkbookWindowRecords,
  stripVbaLineComment,
  switchWorkbookWindow,
  syncCustomViewsDialog,
  syncMacrosDialog,
  syncPairedWorkbookWindowScrollState,
  toggleSynchronousScrolling,
  toggleViewSideBySide,
  unhideWorkbookWindow,
  unmergeVbaRangeTarget,
  updateActiveWorkbookWindowScrollState,
  updateWorkbookWindowScrollFields,
  validateMacroName,
  vbaBorderLineStyle,
  vbaBorderSide,
  vbaBorderWeight,
  vbaCaseClauseMatches,
  vbaColumnIndexFromExpression,
  vbaDoLoopConditionAllows,
  vbaIterableValues,
  vbaNamedArgument,
  vbaRangeHandle,
  vbaStructureOption,
  vbaVariableKey,
  viewCommandMenuButtonHtml,
  workbookTableUsageSummary,
  workbookWindowArrangementModeLabel,
  workbookWindowArrangementPreviewHtml,
  workbookWindowChipHtml,
  workbookWindowMenuDescription,
  workbookWindowRecords,
  workbookWindowSlotStyle,
};
