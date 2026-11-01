// Feature module; shared bindings are imported explicitly.
import {
  $formulaInput,
  $horizontalScrollbar,
  $sheetFooter,
  $sheetHost,
  $sheetTabs,
  $sheetWorkspace,
  $zoomLabel,
  $zoomSlider,
  COLUMN_HEADER_HEIGHT,
  CUSTOM_SHAPE_TYPE,
  DATA_OPERATION_CELL_LIMIT,
  OFFLINE_TRANSLATION_DICTIONARIES,
  QUICK_ANALYSIS_ACTIONS,
  QUICK_ANALYSIS_TABS,
  SHEET_TAB_COLORS,
  SPECIAL_SHEET_INSERT_META,
  TRANSLATION_CELL_LIMIT,
  TRANSLATION_LANGUAGES,
  TRANSLATION_PREVIEW_LIMIT,
  TRANSLATION_SCAN_LIMIT,
  ZOOM_DIALOG_PRESETS,
  createIcons,
  icons,
  state,
} from "../../app-core.js";
import {
  escapeRegExp,
} from "../data/filter-and-text-tools.js";
import {
  closeDataCommandMenu,
} from "../data/query-transforms.js";
import {
  closeDataValidationMenu,
} from "../data/validation-and-consolidation.js";
import {
  accessKeyFromKeyboardEvent,
  activateDialogAccessKeyTarget,
  cellNumericValue,
  closeDataDialog,
  findDialogAccessKeyTarget,
  insertChartFromSelection,
  insertPivotSummaryFromSelection,
  insertSparklinesFromSelection,
  recommendedChartTypeForSelection,
} from "../drawing/charts-and-insertion.js";
import {
  isLineShapeObject,
  normalizeShapeCustomPath,
  normalizedHeaderFooterSettings,
  shapeCanEditPoints,
} from "../drawing/shape-rendering.js";
import {
  closeRibbonBorderMenu,
  closeRibbonColorPalette,
  closeRibbonInkEraserMenu,
  closeRibbonShapeMenu,
  closeRibbonUnderlineMenu,
  closeShapeFormatMenu,
} from "../editing/ink-and-clipboard.js";
import {
  isElementImeCompositionEvent,
  nowMs,
} from "../editing/keyboard-and-paste.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  affectedSheetIndexesForGroupedRanges,
  cellElementForPoint,
  cellPointFromElement,
  cellPointFromPointerEvent,
  cssEscape,
  expandRangeForMerges,
  getDisplayForCell,
  isCellInRange,
  materializeSelectionRange,
  materializedRangesForSheet,
  nativeAnchorFromSheetPixel,
  normalizeRange,
  rangeToLabel,
  rangesToLabel,
  restoreGroupedSheetsAfterFailure,
  selectedObjectFormatMode,
  sheetIndexFromPixel,
  updateSelectionUi,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  applySelectionRange,
  cellHasComment,
  cellHasNote,
  columnSelectionRange,
  coordinateToPixels,
  customPathEditablePointCount,
  customPathVertexForKey,
  imageBoxToPixels,
  isActiveImage,
  rowSelectionRange,
  scrollRangeIntoView,
  selectCellPoint,
  selectedImageObjectById,
  selectedImageObjects,
  setSelectedImageIds,
  shapeCustomPathIsClosed,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureCellsForHistory,
  captureCellsForHistoryRanges,
  captureSheetForHistory,
  captureSheetsForHistory,
  internCellsInRange,
  internCellsInRanges,
  internCssObject,
  internSheetCellStyles,
  mutableCssObject,
  pushHistory,
  restoreSheetFromHistory,
  storeCellModelForSheet,
} from "../editing/structure-and-history.js";
import {
  ensureCellModel,
  forEachCellInRange,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  applyQuickAnalysisTotal,
  cellHasPhonetics,
  cellPasteOptionsHtml,
  closeCellContextMenu,
  interpolateHexColor,
  positionPopupMenu,
  quickAnalysisAvailableForContext,
} from "./annotations-and-cell-format.js";
import {
  accessKeyLabelHtml,
  cellKey,
  columnName,
  columnWidthToPixels,
  escapeAttr,
  escapeHtml,
  rowHeightToPixels,
  setStatus,
  showMessageBox,
} from "./help-and-utilities.js";
import {
  consumeCommandMenuKey,
  enhanceCommandMenuAccessKeys,
} from "./printing-and-home-tools.js";
import {
  closeFileSideMenu,
  ensureSheetSize,
  insertTableFromSelection,
  parseCellKey,
  rowHeaderWidthForSheet,
} from "./ribbon-and-links.js";
import {
  activeSheetView,
  applyGroupedSheetViewOptionChange,
  captureWorkbookViewState,
  commitGroupedSheetViewOptionChange,
  commitSheetViewOptionChange,
  refreshActiveSheetViewSnapshot,
  syncViewRibbonButtons,
  syncWorkbookChromeVisibility,
} from "./views-and-formula-tools.js";
import {
  beginSheetTabInlineRename,
  buildFormulaEngine,
  confirmDeleteSheetAt,
  ensureActiveSheetVisible,
  hiddenSheetIndexes,
  hideSheetsAt,
  isSheetVisible,
  openSheetInsertDialog,
  openSheetMoveCopyDialog,
  renderSheetTabs,
  renderWorkbook,
  scheduleFocusSheetTabByIndex,
  setActiveSheet,
  sheetGroupActionIndexes,
  sheetTabButtonByIndex,
  unhideSheetFromMenu,
  visibleSheetIndexes,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  scheduleWorkbookPersistence,
  sheetCodeModuleName,
} from "../workbook/import-model.js";
import {
  activeSheet,
  cellDisplayText,
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
  validationListOptions,
} from "../workbook/tables-and-validation.js";
import {
  ensurePageLayoutSettings,
  normalizedPageLayoutSettings,
  openSheetCodeDialog,
  pageBreakTargetsForHeaderContext,
  pageBreakTargetsForSelection,
  pageBreakTargetsLabel,
  pageLayoutCurrentSelectionBreakState,
} from "../workbook/vba-and-printing.js";
import {
  applySheetZoomWithoutHistory,
  menuAccessKeyFromLabel,
  saveActiveWindowViewState,
} from "../workbook/views-and-macros.js";


function moveSheetToIndex(fromIndex, toIndex) {
  if (!state.model?.sheets?.length) return 0;
  const count = state.model.sheets.length;
  const from = Math.max(0, Math.min(Number(fromIndex) || 0, count - 1));
  const [sheet] = state.model.sheets.splice(from, 1);
  const insertAt = Math.max(0, Math.min(Number(toIndex) || 0, state.model.sheets.length));
  state.model.sheets.splice(insertAt, 0, sheet);
  return insertAt;
}



function setSheetTabColor(sheetIndex, color) {
  const sheet = state.model?.sheets[sheetIndex];
  if (!sheet) return;
  const nextColor = color || "";
  if ((sheet.tabColor || "") === nextColor) {
    setStatus(sheet.tabColor ? `${sheet.name} のシート見出し色は既に設定されています。` : `${sheet.name} のシート見出し色は既にクリアされています。`);
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  sheet.tabColor = nextColor;
  state.activeSheetIndex = sheetIndex;
  renderSheetTabs();
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore,
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges,
    selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
    label: `${sheet.name} シート見出しの色`,
  });
  setStatus(sheet.tabColor ? `${sheet.name} のシート見出し色を設定しました。` : `${sheet.name} のシート見出し色をクリアしました。`);
  scheduleWorkbookPersistence();
}



function setSheetTabColorForSheets(sheetIndexes, color) {
  if (!state.model?.sheets?.length) return;
  const targets = [...new Set((sheetIndexes || []).map((index) => Number(index)).filter((index) => state.model?.sheets?.[index]))]
    .sort((left, right) => left - right);
  if (!targets.length) return;
  if (targets.length === 1) {
    setSheetTabColor(targets[0], color);
    return;
  }
  const nextColor = color || "";
  if (targets.every((index) => (state.model.sheets[index].tabColor || "") === nextColor)) {
    setStatus(nextColor ? "選択したシート見出し色は既に設定されています。" : "選択したシート見出し色は既にクリアされています。");
    return;
  }
  const sheetsBefore = captureSheetsForHistory(targets);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  targets.forEach((index) => {
    state.model.sheets[index].tabColor = nextColor;
  });
  state.activeSheetIndex = targets.includes(state.activeSheetIndex) ? state.activeSheetIndex : targets[0];
  renderSheetTabs();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(targets),
    selectionBefore,
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges,
    selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
    label: `${targets.length} シート見出しの色`,
  });
  setStatus(nextColor ? `${targets.length} 個のシート見出し色を設定しました。` : `${targets.length} 個のシート見出し色をクリアしました。`);
  scheduleWorkbookPersistence();
}



function selectAllVisibleSheets(sheetIndex) {
  if (!state.model?.sheets[sheetIndex]) return;
  state.activeSheetIndex = sheetIndex;
  state.selectedSheetIndexes = visibleSheetIndexes();
  renderSheetTabs();
  setStatus(`${state.selectedSheetIndexes.length} 個の表示ワークシートを選択しました。`);
}



function rebuildWorkbookAfterSheetChange() {
  state.selectedSheetIndexes = [];
  state.selected = null;
  state.selectedImage = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  state.selectionDrag = null;
  state.formulaEdit = null;
  state.isSelecting = false;
  state.undoStack = [];
  state.redoStack = [];
  closeSheetTabContextMenu();
  closeCellContextMenu();
  state.hf = buildFormulaEngine(state.model);
  ensureActiveSheetVisible();
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  scheduleWorkbookPersistence();
}



function clearSheetSelectionState() {
  state.selectedSheetIndexes = [];
  state.selected = null;
  state.selectedImage = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  state.selectionDrag = null;
  state.formulaEdit = null;
  state.isSelecting = false;
}



function refreshWorkbookAfterSheetMetadataChange(options = {}) {
  if (!options.preserveSheetSelection) state.selectedSheetIndexes = [];
  closeSheetTabContextMenu();
  closeCellContextMenu();
  state.hf = buildFormulaEngine(state.model);
  ensureActiveSheetVisible();
  renderWorkbook();
  updateSelectionUi();
  scheduleWorkbookPersistence();
}



function cloneSheetModel(sheet) {
  const clone = JSON.parse(JSON.stringify(sheet));
  internSheetCellStyles(clone);
  return clone;
}



function uniqueCopiedSheetName(sourceName) {
  const base = String(sourceName || "Sheet").slice(0, 25);
  const existing = new Set(state.model.sheets.map((sheet) => sheet.name));
  return uniqueCopiedSheetNameFromSet(base, existing);
}



function uniqueCopiedSheetNameFromSet(sourceName, existing) {
  const base = String(sourceName || "Sheet").slice(0, 25);
  let index = 2;
  let name = `${base} (${index})`;
  while (existing.has(name) || name.length > 31) {
    index += 1;
    name = `${base.slice(0, Math.max(1, 31 - String(index).length - 3))} (${index})`;
  }
  return name;
}



function nearestVisibleSheetIndex(fromIndex) {
  const visible = visibleSheetIndexes();
  return visible.find((index) => index >= fromIndex) ?? visible.at(-1) ?? 0;
}



function validateSheetName(value, currentIndex = -1) {
  const name = String(value || "").trim();
  if (!name) return { valid: false, message: "ワークシート名を空にすることはできません。" };
  if (name.length > 31) return { valid: false, message: "ワークシート名は 31 文字以内にしてください。" };
  if (/[:\\/?*\[\]]/.test(name)) return { valid: false, message: "ワークシート名には : \\ / ? * [ ] を使用できません。" };
  const duplicateIndex = state.model.sheets.findIndex((sheet, index) => index !== currentIndex && sheet.name === name);
  if (duplicateIndex >= 0) return { valid: false, message: `${name} という名前のワークシートは既に存在します。` };
  return { valid: true, name };
}



function uniqueSheetName(existingNames) {
  const existing = new Set(existingNames);
  let index = existingNames.length + 1;
  let name = `Sheet${index}`;
  while (existing.has(name)) {
    index += 1;
    name = `Sheet${index}`;
  }
  return name;
}



function createBlankSheet(name) {
  const rowCount = 1;
  const colCount = 1;
  const columnWidth = columnWidthToPixels(8.43);
  const rowHeight = rowHeightToPixels(15);
  return {
    name,
    state: "visible",
    tabColor: "",
    rowCount,
    colCount,
    baseRowCount: rowCount,
    baseColCount: colCount,
    defaultColumnWidth: columnWidth,
    defaultRowHeight: rowHeight,
    zoomScale: 100,
    baseZoomScale: 100,
    displayScale: 1,
    defaultCellCss: {},
    showGridLines: true,
    showHeadings: true,
    showRuler: false,
    showObjects: true,
    showOutlineSymbols: true,
    showFormulas: false,
    showZeros: true,
    showWhiteSpace: true,
    rightToLeft: false,
    tabSelected: false,
    viewMode: "normal",
    freezePanes: null,
    splitPane: null,
    sheetViews: [],
    activeSheetViewId: null,
    codeModule: null,
    outlineSettings: { summaryBelow: true, summaryRight: true },
    formulaAudits: { precedents: [], dependents: [] },
    pageLayout: normalizedPageLayoutSettings(),
    headerFooter: normalizedHeaderFooterSettings(),
    columns: Array.from({ length: colCount }, () => ({ width: columnWidth, hidden: false })),
    rows: Array.from({ length: rowCount }, () => ({ height: rowHeight, hidden: false })),
    merges: [],
    images: [],
    inkStrokes: [],
    cells: {},
    autoFilter: null,
    hfData: [],
    warnings: [],
  };
}



function createSpecialSheet(name, templateType) {
  const meta = SPECIAL_SHEET_INSERT_META[templateType] || SPECIAL_SHEET_INSERT_META.macro;
  const sheet = createBlankSheet(name);
  sheet.sheetKind = meta.sheetKind;
  sheet.specialSheet = {
    type: meta.type,
    label: meta.label,
    createdAt: new Date().toISOString(),
  };
  sheet.macroSheet = meta.sheetKind === "macro" || meta.sheetKind === "intlMacro";
  sheet.xl4MacroSheet = meta.sheetKind === "macro";
  sheet.internationalMacroSheet = meta.sheetKind === "intlMacro";
  sheet.dialogSheet = meta.sheetKind === "dialog";
  sheet.codeModule = {
    name: sheetCodeModuleName(sheet),
    objectName: meta.codeObjectName,
    procedure: "",
    code: "",
    updatedAt: "",
  };
  if (sheet.dialogSheet) {
    sheet.showGridLines = false;
    sheet.showHeadings = false;
    sheet.defaultCellCss = internCssObject({
      backgroundColor: "#f0f0f0",
    });
  }
  return sheet;
}



function scrollSheetTabs(direction) {
  const tabs = $sheetTabs[0];
  if (!tabs) return;
  const distance = Math.max(120, Math.round(tabs.clientWidth * 0.7));
  tabs.scrollBy({ left: direction * distance, behavior: "auto" });
}



function updateSheetFooterUi() {
  const sheet = state.model ? activeSheet() : null;
  const zoom = Math.round(sheet?.zoomScale || 100);
  const viewMode = sheet?.viewMode || "normal";
  $zoomSlider.val(String(zoom));
  $zoomLabel.text(`${zoom}%`);
  updateOpenZoomDialogValue(zoom);
  $sheetFooter.find("[data-sheet-view]").each((_index, element) => {
    element.classList.toggle("is-active", element.dataset.sheetView === viewMode);
  });
  syncWorkbookChromeVisibility();
}



function openZoomDialog(source = null) {
  const event = typeof source?.preventDefault === "function" ? source : null;
  const anchor = event?.currentTarget || (typeof source?.getBoundingClientRect === "function" ? source : null);
  event?.preventDefault();
  event?.stopPropagation();
  if (!state.model) return;

  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonColorPalette();
  closeZoomDialog({ restoreFocus: false });

  const currentZoom = Math.round(Number(activeSheet().zoomScale || 100));
  const presetValues = new Set(ZOOM_DIALOG_PRESETS.map((preset) => preset.value));
  const selectedValue = presetValues.has(currentZoom) ? String(currentZoom) : "custom";
  const optionsHtml = ZOOM_DIALOG_PRESETS.map(
    (preset) => `<label class="zoom-dialog-option"${zoomDialogAccessKeyAttr(preset.label)}>
      <input type="radio" name="zoomDialogOption" value="${preset.value}"${selectedValue === String(preset.value) ? " checked" : ""} />
      <span>${accessKeyLabelHtml(preset.label)}</span>
    </label>`,
  ).join("");
  const dialog = document.createElement("div");
  dialog.id = "zoomDialog";
  dialog.className = "zoom-dialog";
  dialog.setAttribute("role", "dialog");
  dialog.setAttribute("aria-label", "ズーム");
  dialog.innerHTML = `
    <div class="zoom-dialog-title">ズーム</div>
    <div class="zoom-dialog-section-title">倍率</div>
    <div class="zoom-dialog-options">
      ${optionsHtml}
      <label class="zoom-dialog-option"${zoomDialogAccessKeyAttr("選択範囲をズーム(F)")}>
        <input type="radio" name="zoomDialogOption" value="selection" />
        <span>${accessKeyLabelHtml("選択範囲をズーム(F)")}</span>
      </label>
      <label class="zoom-dialog-option zoom-dialog-custom"${zoomDialogAccessKeyAttr("指定(C):")}>
        <input type="radio" name="zoomDialogOption" value="custom"${selectedValue === "custom" ? " checked" : ""} />
        <span>${accessKeyLabelHtml("指定(C):")}</span>
        <input class="zoom-dialog-input" type="number" min="10" max="400" step="1" value="${currentZoom}" data-zoom-custom-input aria-label="ズーム倍率" />
        <span>%</span>
      </label>
    </div>
    <div class="zoom-dialog-actions">
      <button class="zoom-dialog-button is-primary" type="button" data-zoom-dialog-action="ok">OK</button>
      <button class="zoom-dialog-button" type="button" data-zoom-dialog-action="cancel">キャンセル</button>
    </div>`;
  document.body.append(dialog);
  state.zoomDialog = { previousFocus: document.activeElement, anchor };
  positionZoomDialog(dialog);
  const initialFocus = selectedValue === "custom" ? dialog.querySelector("[data-zoom-custom-input]") : dialog.querySelector("input:checked");
  initialFocus?.focus({ preventScroll: true });
  if (selectedValue === "custom") initialFocus?.select?.();
}



function zoomDialogAccessKeyAttr(label) {
  const accessKey = menuAccessKeyFromLabel(label);
  return accessKey ? ` data-access-key="${escapeAttr(accessKey)}"` : "";
}



function positionZoomDialog(dialog) {
  const anchor = state.zoomDialog?.anchor || $zoomLabel[0];
  if (!anchor || !dialog) return;

  dialog.style.visibility = "hidden";
  dialog.style.left = "0px";
  dialog.style.top = "0px";
  const anchorRect = anchor.getBoundingClientRect();
  const dialogRect = dialog.getBoundingClientRect();
  const margin = 8;
  const isRibbonAnchor = Boolean(anchor.closest?.(".ribbon-panels"));
  const preferredLeft = isRibbonAnchor ? anchorRect.left : anchorRect.right - dialogRect.width;
  const left = Math.max(margin, Math.min(preferredLeft, window.innerWidth - dialogRect.width - margin));
  const topAbove = anchorRect.top - dialogRect.height - margin;
  const preferredTop = isRibbonAnchor ? anchorRect.bottom + 4 : (topAbove >= margin ? topAbove : anchorRect.bottom + margin);
  const top = Math.min(preferredTop, window.innerHeight - dialogRect.height - margin);
  dialog.style.left = `${Math.round(left)}px`;
  dialog.style.top = `${Math.round(Math.max(margin, top))}px`;
  dialog.style.visibility = "visible";
}



function handleZoomDialogAction(event) {
  const button = event.currentTarget;
  if (!button.closest("#zoomDialog")) return;
  event.preventDefault();
  event.stopPropagation();
  if (button.dataset.zoomDialogAction === "cancel") {
    closeZoomDialog();
    return;
  }
  applyZoomDialog();
}



function handleZoomDialogKeydown(event) {
  if (isElementImeCompositionEvent(event)) return;
  const accessKey = accessKeyFromKeyboardEvent(event);
  if (accessKey && event.altKey && !event.ctrlKey && !event.metaKey) {
    const target = findDialogAccessKeyTarget(event.currentTarget, accessKey);
    if (!target) return;
    event.preventDefault();
    activateDialogAccessKeyTarget(target);
    return;
  }
  if (event.key === "Enter") {
    event.preventDefault();
    applyZoomDialog();
    return;
  }
  if (event.key === "Escape") {
    event.preventDefault();
    closeZoomDialog();
  }
}



function handleZoomDialogOptionChange(event) {
  const dialog = event.target.closest?.("#zoomDialog");
  if (!dialog) return;
  if (event.target.value === "custom") {
    dialog.querySelector("[data-zoom-custom-input]")?.focus({ preventScroll: true });
  }
}



function selectCustomZoomOption(event) {
  const dialog = event.target.closest?.("#zoomDialog");
  if (!dialog) return;
  const customRadio = dialog.querySelector('input[name="zoomDialogOption"][value="custom"]');
  if (customRadio) customRadio.checked = true;
}



function applyZoomDialog() {
  const dialog = document.querySelector("#zoomDialog");
  if (!dialog || !state.model) return;
  const selectedOption = dialog.querySelector('input[name="zoomDialogOption"]:checked')?.value || "custom";

  if (selectedOption === "selection") {
    if (zoomSelectionToFit()) closeZoomDialog();
    return;
  }

  const rawValue = selectedOption === "custom" ? dialog.querySelector("[data-zoom-custom-input]")?.value : selectedOption;
  setActiveSheetZoom(clampZoom(rawValue));
  closeZoomDialog();
}



function zoomSelectionToFit({ defer = true } = {}) {
  const host = $sheetHost[0];
  const ranges = activeSelectionRanges();
  const fitRange = zoomSelectionFitRange(activeSheet(), ranges);
  if (!fitRange || !host) {
    setStatus("ズームするセル範囲を選択してください。");
    return false;
  }

  const sheet = activeSheet();
  const left = coordinateToPixels(sheet, "col", fitRange.left - 1);
  const top = coordinateToPixels(sheet, "row", fitRange.top - 1);
  const right = coordinateToPixels(sheet, "col", fitRange.right);
  const bottom = coordinateToPixels(sheet, "row", fitRange.bottom);
  const rangeWidth = Math.max(1, right - left);
  const rangeHeight = Math.max(1, bottom - top);
  const availableWidth = Math.max(1, host.clientWidth - rowHeaderWidthForSheet(sheet) - 16);
  const availableHeight = Math.max(1, host.clientHeight - COLUMN_HEADER_HEIGHT - 16);
  const currentZoom = Number(sheet.zoomScale || 100);
  const fitRatio = Math.min(availableWidth / rangeWidth, availableHeight / rangeHeight);
  let nextZoom = clampZoom(currentZoom * fitRatio);
  if (nextZoom === Math.round(currentZoom)) {
    const direction = fitRatio >= 1 ? 1 : -1;
    nextZoom = clampZoom(currentZoom + direction * 10);
  }
  const statusLabel = rangesToLabel(ranges.length ? ranges : [fitRange]);
  const applyZoom = () => {
    setActiveSheetZoom(nextZoom);
    requestAnimationFrame(() => {
      const activeRange = zoomSelectionFitRange(activeSheet(), activeSelectionRanges()) || activeSelectionRange();
      if (rangeWidth < host.clientWidth * 4 && rangeHeight < host.clientHeight * 4) {
        scrollRangeIntoView(activeRange);
      }
      updateSelectionUi();
    });
    setStatus(`${statusLabel}：ズーム ${nextZoom}%`);
  };
  if (defer) setTimeout(applyZoom, 0);
  else applyZoom();
  setStatus(`${statusLabel}：ズーム ${nextZoom}%`);
  return true;
}



function zoomSelectionFitRange(sheet, ranges = activeSelectionRanges()) {
  const normalizedRanges = (ranges || [])
    .map((range) => zoomSelectionNormalizeRange(sheet, range))
    .filter(Boolean);
  if (!normalizedRanges.length) return null;
  return zoomSelectionUnionRanges(normalizedRanges);
}



function zoomSelectionNormalizeRange(sheet, range) {
  if (!sheet || !range) return null;
  const used = sheetUsedRangeForZoom(sheet);
  const fallback = used || {
    top: Math.min(Math.max(1, range.top || 1), sheet.rowCount || 1),
    left: Math.min(Math.max(1, range.left || 1), sheet.colCount || 1),
    bottom: Math.min(Math.max(1, range.bottom || range.top || 1), sheet.rowCount || 1),
    right: Math.min(Math.max(1, range.right || range.left || 1), sheet.colCount || 1),
  };
  const normalized = { ...range, sheetIndex: state.activeSheetIndex };
  if (range.fullRow && range.fullColumn) {
    normalized.top = fallback.top;
    normalized.bottom = fallback.bottom;
    normalized.left = fallback.left;
    normalized.right = fallback.right;
  } else if (range.fullRow) {
    normalized.left = fallback.left;
    normalized.right = fallback.right;
  } else if (range.fullColumn) {
    normalized.top = fallback.top;
    normalized.bottom = fallback.bottom;
  }
  const rowLimit = Math.max(1, Number(sheet.rowCount) || 0, Number(fallback.bottom) || 0, Number(normalized.bottom) || 0);
  const colLimit = Math.max(1, Number(sheet.colCount) || 0, Number(fallback.right) || 0, Number(normalized.right) || 0);
  normalized.top = Math.min(Math.max(1, normalized.top || 1), rowLimit);
  normalized.left = Math.min(Math.max(1, normalized.left || 1), colLimit);
  normalized.bottom = Math.min(Math.max(normalized.top, normalized.bottom || normalized.top), rowLimit);
  normalized.right = Math.min(Math.max(normalized.left, normalized.right || normalized.left), colLimit);
  return normalized;
}



function sheetUsedRangeForZoom(sheet) {
  if (!sheet) return null;
  const ranges = [];
  Object.entries(sheet.cells || {}).forEach(([key, cell]) => {
    if (!cell) return;
    const hasContent = String(cellDisplayText(cell)).trim() ||
      cell.raw != null ||
      cell.formula ||
      cell.cached != null ||
      cell.richText?.length ||
      cell.hyperlink ||
      cell.note ||
      cell.commentThread?.length;
    const hasStyle = cell.css && Object.keys(cell.css).length;
    const hasBorders = cell.borders && Object.keys(cell.borders).length;
    if (!hasContent && !hasStyle && !hasBorders) return;
    const [row, col] = parseCellKey(key);
    if (Number.isFinite(row) && Number.isFinite(col)) ranges.push({ top: row, left: col, bottom: row, right: col });
  });
  (sheet.merges || []).forEach((merge) => {
    const normalized = normalizeZoomUsedRange(sheet, merge);
    if (normalized) ranges.push(normalized);
  });
  (sheet.tables || []).forEach((table) => {
    const normalized = normalizeZoomUsedRange(sheet, table?.range);
    if (normalized) ranges.push(normalized);
  });
  const filterRange = normalizeZoomUsedRange(sheet, sheet.autoFilter?.range);
  if (filterRange) ranges.push(filterRange);
  (sheet.images || []).forEach((image) => {
    if (image?.hidden) return;
    const normalized = imageUsedRangeForZoom(sheet, image);
    if (normalized) ranges.push(normalized);
  });
  return zoomSelectionUnionRanges(ranges.map((range) => normalizeZoomUsedRange(sheet, range)).filter(Boolean));
}



function normalizeZoomUsedRange(sheet, range) {
  if (!sheet || !range) return null;
  const top = Math.max(1, Math.round(Number(range.top ?? range.row) || 1));
  const left = Math.max(1, Math.round(Number(range.left ?? range.col) || 1));
  const bottom = Math.max(top, Math.round(Number(range.bottom) || top));
  const right = Math.max(left, Math.round(Number(range.right) || left));
  const rowCount = Math.max(1, Number(sheet.rowCount) || 0, bottom);
  const colCount = Math.max(1, Number(sheet.colCount) || 0, right);
  return {
    sheetIndex: state.activeSheetIndex,
    top: Math.min(top, rowCount),
    left: Math.min(left, colCount),
    bottom: Math.min(bottom, rowCount),
    right: Math.min(right, colCount),
  };
}



function imageUsedRangeForZoom(sheet, image) {
  if (!sheet || !image) return null;
  if (image.range) return normalizeZoomUsedRange(sheet, image.range);
  if (image.from) {
    return normalizeZoomUsedRange(sheet, {
      top: image.from.row || 1,
      left: image.from.col || 1,
      bottom: image.to?.row || image.from.row || 1,
      right: image.to?.col || image.from.col || 1,
    });
  }
  const box = imageBoxToPixels(sheet, image);
  if (!box) {
    const top = Math.max(1, Math.floor(Number(image.tl?.row ?? image.tl?.nativeRow ?? 0)) + 1);
    const left = Math.max(1, Math.floor(Number(image.tl?.col ?? image.tl?.nativeCol ?? 0)) + 1);
    return normalizeZoomUsedRange(sheet, { top, left, bottom: top, right: left });
  }
  const top = sheetIndexFromPixel(sheet, "row", Math.max(COLUMN_HEADER_HEIGHT, box.top)) ||
    Math.max(1, Math.floor(Number(image.tl?.row ?? image.tl?.nativeRow ?? 0)) + 1);
  const left = sheetIndexFromPixel(sheet, "col", Math.max(rowHeaderWidthForSheet(sheet), box.left)) ||
    Math.max(1, Math.floor(Number(image.tl?.col ?? image.tl?.nativeCol ?? 0)) + 1);
  const bottomAnchor = nativeAnchorFromSheetPixel(sheet, "row", box.top + Math.max(1, box.height));
  const rightAnchor = nativeAnchorFromSheetPixel(sheet, "col", box.left + Math.max(1, box.width));
  return normalizeZoomUsedRange(sheet, {
    top,
    left,
    bottom: Math.max(top, (bottomAnchor?.index ?? top - 1) + 1),
    right: Math.max(left, (rightAnchor?.index ?? left - 1) + 1),
  });
}



function zoomSelectionUnionRanges(ranges) {
  const items = (ranges || []).filter(Boolean);
  if (!items.length) return null;
  return {
    sheetIndex: items[0].sheetIndex ?? state.activeSheetIndex,
    top: Math.min(...items.map((range) => range.top)),
    left: Math.min(...items.map((range) => range.left)),
    bottom: Math.max(...items.map((range) => range.bottom)),
    right: Math.max(...items.map((range) => range.right)),
  };
}



function closeZoomDialogOnOutsidePointer(event) {
  if (!state.zoomDialog) return;
  const anchor = state.zoomDialog?.anchor;
  if (
    event.target.closest?.("#zoomDialog") ||
    event.target.closest?.("#zoomLabel") ||
    (anchor && (event.target === anchor || anchor.contains?.(event.target)))
  ) return;
  closeZoomDialog();
}



function closeZoomDialogOnEscape(event) {
  if (event.key !== "Escape" || !state.zoomDialog) return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  closeZoomDialog();
}



function closeZoomDialog(options = {}) {
  const dialog = document.querySelector("#zoomDialog");
  const previousFocus = state.zoomDialog?.previousFocus;
  dialog?.remove();
  state.zoomDialog = null;
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) {
    previousFocus.focus?.({ preventScroll: true });
  }
}



function updateOpenZoomDialogValue(zoom) {
  const dialog = document.querySelector("#zoomDialog");
  if (!dialog) return;
  const input = dialog.querySelector("[data-zoom-custom-input]");
  if (input && document.activeElement !== input) input.value = String(zoom);
}



function setActiveSheetView(viewMode) {
  if (!state.model || !viewMode) return;
  if (!["normal", "page-layout", "page-break"].includes(viewMode)) return;
  const sheet = activeSheet();
  const previousMode = sheet.viewMode || "normal";
  const previousRuler = sheet.showRuler === true;
  const labels = {
    normal: "標準",
    "page-layout": "ページレイアウト",
    "page-break": "改ページプレビュー",
  };
  if (applyGroupedSheetViewOptionChange((targetSheet) => {
    targetSheet.viewMode = viewMode;
    targetSheet.showRuler = viewMode === "page-layout";
  }, labels[viewMode], `${activeSheetGroupEditIndexes().length} シート 表示 ${labels[viewMode]}`)) {
    return;
  }
  if (previousMode === viewMode) {
    renderSheet();
    syncViewRibbonButtons();
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.viewMode = viewMode;
  sheet.showRuler = viewMode === "page-layout";
  const status = previousRuler && viewMode !== "page-layout"
    ? `${sheet.name}：${labels[viewMode]}（ルーラーを非表示にしました）`
    : `${sheet.name}：${labels[viewMode]}`;
  commitSheetViewOptionChange(sheet, sheetBefore, status, `${sheet.name} 表示 ${labels[viewMode]}`, workbookViewBefore);
}



function adjustActiveSheetZoom(delta) {
  if (!state.model) return;
  const current = Number(activeSheet().zoomScale || 100);
  setActiveSheetZoom(current + delta);
}



function handleSheetZoomWheel(event) {
  if (!state.model || !event.ctrlKey || !isWorkbookWheelTarget(event)) return;
  event.preventDefault();
  event.stopPropagation();
  const deltaY = Number(event.deltaY || 0);
  if (!deltaY) return;
  const direction = deltaY < 0 ? 1 : -1;
  const steps = Math.max(1, Math.min(5, Math.round(Math.abs(deltaY) / 100) || 1));
  adjustActiveSheetZoom(direction * steps * 10);
}



function isWorkbookWheelTarget(event) {
  const workspace = $sheetWorkspace[0];
  if (!workspace || workspace.classList.contains("hidden")) return false;
  const path = event.composedPath?.() || [];
  return path.includes(workspace) || workspace.contains(event.target);
}



function setActiveSheetZoom(value) {
  if (!state.model) return;
  const sheet = activeSheet();
  const previousZoom = Number(sheet.zoomScale || 100);
  const nextZoom = clampZoom(value);
  const groupedSheetIndexes = activeSheetGroupEditIndexes();
  if (groupedSheetIndexes.length > 1) {
    const hasChanges = groupedSheetIndexes.some((sheetIndex) => Number(state.model.sheets[sheetIndex]?.zoomScale || 100) !== nextZoom);
    if (!hasChanges) {
      updateSheetFooterUi();
      syncViewRibbonButtons();
      return;
    }
    const workbookViewBefore = captureWorkbookViewState();
    const sheetsBefore = captureSheetsForHistory(groupedSheetIndexes);
    groupedSheetIndexes.forEach((sheetIndex) => applySheetZoomWithoutHistory(state.model.sheets[sheetIndex], nextZoom));
    commitGroupedSheetViewOptionChange(
      groupedSheetIndexes,
      sheetsBefore,
      `${groupedSheetIndexes.length} 個のシート：ズーム ${nextZoom}%`,
      `${groupedSheetIndexes.length} シート ズーム ${nextZoom}%`,
      workbookViewBefore,
    );
    return;
  }
  const sheetBefore = activeSheetView(sheet) ? captureSheetForHistory(sheet) : null;
  const zoomBefore = captureZoomForHistory(sheet);
  if (nextZoom === previousZoom) {
    updateSheetFooterUi();
    syncViewRibbonButtons();
    return;
  }

  const previousScale = Number(sheet.displayScale || 1);
  const nextScale = displayScaleForZoom(sheet, nextZoom);
  const ratio = nextScale / previousScale;
  scaleSheetDisplay(sheet, ratio);
  sheet.zoomScale = nextZoom;
  sheet.displayScale = nextScale;
  const zoomAfter = captureZoomForHistory(sheet);
  refreshActiveSheetViewSnapshot(sheet);
  const selection = activeSelectionRange();
  const historyEntry = {
    sheetIndex: state.activeSheetIndex,
    selectionBefore: selection,
    selectionAfter: selection,
    label: `${sheet.name} ズーム ${Math.round(previousZoom)}% → ${nextZoom}%`,
  };
  if (sheetBefore) {
    historyEntry.sheetBefore = sheetBefore;
    historyEntry.sheetAfter = captureSheetForHistory(sheet);
  } else {
    historyEntry.zoomBefore = zoomBefore;
    historyEntry.zoomAfter = zoomAfter;
  }
  pushHistory(historyEntry);
  saveActiveWindowViewState();
  renderSheet();
  updateSheetFooterUi();
  syncViewRibbonButtons();
  setStatus(`${sheet.name}：ズーム ${nextZoom}%`);
  scheduleWorkbookPersistence();
}



function captureZoomForHistory(sheet) {
  const zoom = Number(sheet?.zoomScale || 100);
  return {
    zoomScale: zoom,
    baseZoomScale: sheetBaseZoomScale(sheet),
    displayScale: Number(sheet?.displayScale || 1),
  };
}



function restoreZoomFromHistory(sheet, zoomState) {
  if (!zoomState) return;
  const targetZoom = clampZoom(zoomState.zoomScale);
  const currentZoom = Number(sheet.zoomScale || 100);
  sheet.baseZoomScale = Number(zoomState.baseZoomScale || sheet.baseZoomScale || currentZoom || 100);
  const currentScale = Number(sheet.displayScale || 1);
  const targetScale = Number(zoomState.displayScale || displayScaleForZoom(sheet, targetZoom));
  const ratio = targetScale / currentScale;
  scaleSheetDisplay(sheet, ratio);
  sheet.zoomScale = targetZoom;
  sheet.displayScale = targetScale;
}



function sheetBaseZoomScale(sheet) {
  const base = Number(sheet?.baseZoomScale || sheet?.zoomScale || 100);
  return Number.isFinite(base) && base > 0 ? base : 100;
}



function displayScaleForZoom(sheet, zoom) {
  const targetZoom = clampZoom(zoom);
  return targetZoom / sheetBaseZoomScale(sheet);
}



function clampZoom(value) {
  const zoom = Math.round(Number(value) || 100);
  return Math.max(10, Math.min(400, zoom));
}



function scaleSheetDisplay(sheet, ratio) {
  if (!Number.isFinite(ratio) || ratio <= 0 || Math.abs(ratio - 1) < 0.0001) return;
  sheet.defaultColumnWidth = Math.max(1, sheet.defaultColumnWidth * ratio);
  sheet.defaultRowHeight = Math.max(1, sheet.defaultRowHeight * ratio);
  sheet.columns.forEach((column) => {
    column.width = Math.max(0, (column.width || 0) * ratio);
  });
  sheet.rows.forEach((row) => {
    row.height = Math.max(0, (row.height || 0) * ratio);
  });
  sheet.images?.forEach((image) => {
    if (image.ext) {
      image.ext.width = Math.max(Number.EPSILON, (Number(image.ext.width) || 1) * ratio);
      image.ext.height = Math.max(Number.EPSILON, (Number(image.ext.height) || 1) * ratio);
    }
  });
}



function openSheetTabContextMenu(event) {
  event.preventDefault();
  const sheetIndex = Number(event.currentTarget.dataset.sheetIndex);
  if (!state.model?.sheets[sheetIndex]) return;

  activateSheetTabForContextMenu(sheetIndex);
  renderSheetTabContextMenu(sheetIndex, event.clientX, event.clientY);
}



function activateSheetTabForContextMenu(sheetIndex) {
  if (!state.model?.sheets[sheetIndex]) return false;
  const groupIncludesSheet = (state.selectedSheetIndexes || []).includes(sheetIndex);
  if (!groupIncludesSheet && state.activeSheetIndex !== sheetIndex && isSheetVisible(state.model.sheets[sheetIndex])) {
    state.activeSheetIndex = sheetIndex;
    state.selectedSheetIndexes = [];
    state.selected = null;
    state.selectedImage = null;
    state.selectionRange = null;
    state.selectionRanges = [];
    state.selectionAnchor = null;
    state.selectionStart = null;
    state.selectionDrag = null;
    state.formulaEdit = null;
    state.isSelecting = false;
    $sheetHost.scrollTop(0);
    $sheetHost.scrollLeft(0);
    $horizontalScrollbar.scrollLeft(0);
    renderWorkbook();
  }
  return true;
}



function openSheetTabContextMenuByKeyboard(tab) {
  const sheetIndex = Number(tab?.dataset?.sheetIndex);
  if (!state.model?.sheets[sheetIndex] || !isSheetVisible(state.model.sheets[sheetIndex])) return false;
  activateSheetTabForContextMenu(sheetIndex);
  const currentTab = sheetTabButtonByIndex(sheetIndex) || tab;
  const rect = currentTab.getBoundingClientRect();
  const clientX = Math.min(window.innerWidth - 8, Math.max(8, rect.left + Math.min(rect.width / 2, 28)));
  const clientY = Math.min(window.innerHeight - 8, Math.max(8, rect.top + rect.height / 2));
  renderSheetTabContextMenu(sheetIndex, clientX, clientY);
  return true;
}



function renderSheetTabContextMenu(sheetIndex, clientX, clientY) {
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  const sheet = state.model?.sheets[sheetIndex];
  if (!sheet) return;

  state.sheetContextMenu = { sheetIndex };
  const html = sheetTabContextMenuItems(sheetIndex)
    .map((item) => {
      if (item.type === "separator") return '<div class="sheet-tab-menu-separator" role="separator"></div>';
      if (item.type === "color") return sheetTabColorSubmenuHtml(sheetIndex, item);
      const disabled = item.disabled ? " disabled" : "";
      const className = item.disabled ? "sheet-tab-menu-item is-disabled" : "sheet-tab-menu-item";
      return `<button class="${className}" type="button" role="menuitem" data-sheet-menu-action="${escapeAttr(item.action)}" data-sheet-index="${sheetIndex}"${disabled}>${accessKeyLabelHtml(item.label)}</button>`;
    })
    .join("");
  const menu = document.createElement("div");
  menu.id = "sheetTabContextMenu";
  menu.className = "sheet-tab-context-menu";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", `${sheet.name} menu`);
  menu.innerHTML = html;
  enhanceCommandMenuAccessKeys(menu);
  document.body.append(menu);
  positionSheetTabContextMenu(menu, clientX, clientY);
  menu.querySelector("button:not([disabled])")?.focus({ preventScroll: true });
}



function openSheetNavigatorMenu(event) {
  event.preventDefault();
  event.stopPropagation();
  if (!state.model?.sheets?.length) return;
  renderSheetNavigatorMenu(event.clientX, event.clientY);
}



function renderSheetNavigatorMenu(clientX, clientY) {
  closeSheetTabContextMenu();
  closeSheetNavigatorMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  const indexes = visibleSheetIndexes();
  if (!indexes.length) return;
  state.sheetNavigatorMenu = { open: true };
  const html = indexes.map((sheetIndex) => {
    const sheet = state.model.sheets[sheetIndex];
    const active = sheetIndex === state.activeSheetIndex;
    return `<button class="sheet-tab-menu-item sheet-navigator-menu-item${active ? " is-active" : ""}" type="button" role="menuitemradio" aria-checked="${active ? "true" : "false"}" data-sheet-navigator-index="${sheetIndex}" title="${escapeAttr(sheet.name)}"><span class="sheet-navigator-check" aria-hidden="true">${active ? "✓" : ""}</span><span class="sheet-navigator-name">${escapeHtml(sheet.name)}</span></button>`;
  }).join("");
  const menu = document.createElement("div");
  menu.id = "sheetTabNavigatorMenu";
  menu.className = "sheet-tab-context-menu sheet-tab-navigator-menu";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", "シートの選択");
  menu.innerHTML = html;
  document.body.append(menu);
  positionSheetTabContextMenu(menu, clientX, clientY);
  (menu.querySelector(".is-active") || menu.querySelector("button"))?.focus({ preventScroll: true });
}



function closeSheetNavigatorMenuOnOutsidePointer(event) {
  if (!state.sheetNavigatorMenu) return;
  if (event.target.closest?.("#sheetTabNavigatorMenu")) return;
  closeSheetNavigatorMenu();
}



function closeSheetNavigatorMenu() {
  document.querySelector("#sheetTabNavigatorMenu")?.remove();
  state.sheetNavigatorMenu = null;
}



function handleSheetNavigatorMenuChoice(event) {
  const button = event.currentTarget;
  if (!button.closest("#sheetTabNavigatorMenu")) return;
  event.preventDefault();
  event.stopPropagation();
  const sheetIndex = Number(button.dataset.sheetNavigatorIndex);
  const sheet = state.model?.sheets?.[sheetIndex];
  closeSheetNavigatorMenu();
  if (!sheet || !isSheetVisible(sheet)) return;
  if (sheetIndex !== state.activeSheetIndex || (state.selectedSheetIndexes || []).length) {
    setActiveSheet(sheetIndex);
  }
  scheduleFocusSheetTabByIndex(sheetIndex, 24, { forceUntil: nowMs() + 800 });
  setStatus(`${sheet.name} シートを表示しました。`);
}



function sheetTabContextMenuItems(sheetIndex) {
  const hiddenCount = hiddenSheetIndexes().length;
  const visibleCount = visibleSheetIndexes().length;
  const actionIndexes = sheetGroupActionIndexes(sheetIndex);
  const visibleActionCount = actionIndexes.filter((index) => isSheetVisible(state.model?.sheets?.[index])).length;
  const canRemoveVisibleSheet = visibleCount > visibleActionCount;
  const sheet = state.model?.sheets?.[sheetIndex];
  const protectedSheet = isSheetProtected(sheet);
  const grouped = actionIndexes.length > 1;
  return [
    { action: "insert", label: "挿入(I)..." },
    { action: "delete", label: "削除(D)", disabled: !canRemoveVisibleSheet },
    { action: "rename", label: "名前の変更(R)", disabled: grouped },
    { action: "move-copy", label: "移動またはコピー(M)..." },
    { action: "code", label: "コード表示(V)", disabled: grouped },
    { action: "protect", label: protectedSheet ? "シート保護の解除(P)..." : "シートの保護(P)...", disabled: grouped },
    { type: "color", label: "シート見出しの色(T)" },
    { type: "separator" },
    { action: "hide", label: "非表示(H)", disabled: !canRemoveVisibleSheet },
    { action: "unhide", label: "再表示(U)...", disabled: hiddenCount === 0 },
    { type: "separator" },
    { action: "select-all", label: "すべてのシートを選択(S)" },
  ];
}



function sheetTabColorSubmenuHtml(sheetIndex, item) {
  const colors = SHEET_TAB_COLORS.map((color) => {
    const style = color.value ? ` style="--menu-color:${escapeAttr(color.value)}"` : "";
    const swatchClass = color.value ? "sheet-tab-color-swatch" : "sheet-tab-color-swatch is-empty";
    return `<button class="sheet-tab-color-choice" type="button" role="menuitem" data-sheet-index="${sheetIndex}" data-sheet-tab-color="${escapeAttr(color.value)}"><span class="${swatchClass}"${style}></span>${accessKeyLabelHtml(color.label)}</button>`;
  }).join("");
  return `<div class="sheet-tab-menu-submenu">
    <button class="sheet-tab-menu-item has-submenu" type="button" role="menuitem" aria-haspopup="true">${accessKeyLabelHtml(item.label)}<span class="sheet-tab-menu-arrow">›</span></button>
    <div class="sheet-tab-color-menu" role="menu">${colors}</div>
  </div>`;
}



function positionSheetTabContextMenu(menu, clientX, clientY) {
  menu.style.visibility = "hidden";
  menu.style.left = "0px";
  menu.style.top = "0px";
  const rect = menu.getBoundingClientRect();
  const margin = 8;
  const left = Math.max(margin, Math.min(clientX, window.innerWidth - rect.width - margin));
  const top = Math.max(margin, Math.min(clientY, window.innerHeight - rect.height - margin));
  menu.classList.toggle("opens-left", left + rect.width + 180 > window.innerWidth);
  menu.style.left = `${Math.round(left)}px`;
  menu.style.top = `${Math.round(top)}px`;
  menu.style.visibility = "visible";
}



function closeSheetTabContextMenuOnOutsidePointer(event) {
  if (!state.sheetContextMenu) return;
  if (event.target.closest?.("#sheetTabContextMenu")) return;
  closeSheetTabContextMenu();
}



function closeSheetTabContextMenuOnEscape(event) {
  if (event.key !== "Escape") return;
  if (!document.querySelector("#sheetTabContextMenu, #sheetTabNavigatorMenu, #cellContextMenu")) return;
  consumeCommandMenuKey(event);
  closeSheetTabContextMenu();
  closeSheetNavigatorMenu();
  closeCellContextMenu();
}



function closeSheetTabContextMenu() {
  document.querySelector("#sheetTabContextMenu")?.remove();
  state.sheetContextMenu = null;
}



function handleSheetTabContextMenuAction(event) {
  const button = event.currentTarget;
  if (!button.closest("#sheetTabContextMenu")) return;
  event.preventDefault();
  event.stopPropagation();
  const action = button.dataset.sheetMenuAction;
  const sheetIndex = Number(button.dataset.sheetIndex ?? state.sheetContextMenu?.sheetIndex);
  closeSheetTabContextMenu();
  runSheetTabMenuAction(action, sheetIndex);
}



function handleSheetTabColorChoice(event) {
  const button = event.currentTarget;
  if (!button.closest("#sheetTabContextMenu")) return;
  event.preventDefault();
  event.stopPropagation();
  const sheetIndex = Number(button.dataset.sheetIndex ?? state.sheetContextMenu?.sheetIndex);
  const color = button.dataset.sheetTabColor || "";
  closeSheetTabContextMenu();
  setSheetTabColorForSheets(sheetGroupActionIndexes(sheetIndex), color);
}



function runSheetTabMenuAction(action, sheetIndex) {
  if (!state.model?.sheets[sheetIndex]) return;
  switch (action) {
    case "insert":
      openSheetInsertDialog(sheetIndex);
      break;
    case "delete":
      confirmDeleteSheetAt(sheetIndex);
      break;
    case "rename":
      beginSheetTabInlineRename(sheetIndex);
      break;
    case "move-copy":
      openSheetMoveCopyDialog(sheetIndex);
      break;
    case "code":
      openSheetCodeDialog(sheetIndex);
      break;
    case "protect":
      toggleSheetProtectionDialog(sheetIndex);
      break;
    case "hide":
      hideSheetsAt(sheetGroupActionIndexes(sheetIndex));
      break;
    case "unhide":
      unhideSheetFromMenu();
      break;
    case "select-all":
      selectAllVisibleSheets(sheetIndex);
      break;
    default:
      break;
  }
}



function toggleSheetProtectionDialog(sheetIndex = state.activeSheetIndex) {
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) return;
  if (isSheetProtected(sheet)) {
    openSheetUnprotectDialog(sheetIndex);
  } else {
    openSheetProtectionDialog(sheetIndex);
  }
}



function isSheetProtected(sheet = activeSheet()) {
  return Boolean(sheet?.protection?.enabled);
}



function defaultSheetProtectionOptions(existing = {}) {
  return {
    enabled: true,
    password: String(existing.password || ""),
    allowSelectLockedCells: existing.allowSelectLockedCells !== false,
    allowSelectUnlockedCells: existing.allowSelectUnlockedCells !== false,
    allowFormatCells: Boolean(existing.allowFormatCells),
    allowFormatColumns: Boolean(existing.allowFormatColumns),
    allowFormatRows: Boolean(existing.allowFormatRows),
    allowInsertColumns: Boolean(existing.allowInsertColumns),
    allowInsertRows: Boolean(existing.allowInsertRows),
    allowDeleteColumns: Boolean(existing.allowDeleteColumns),
    allowDeleteRows: Boolean(existing.allowDeleteRows),
    allowSort: Boolean(existing.allowSort),
    allowAutoFilter: Boolean(existing.allowAutoFilter),
    allowPivotTables: Boolean(existing.allowPivotTables),
    allowEditObjects: Boolean(existing.allowEditObjects),
    allowEditScenarios: Boolean(existing.allowEditScenarios),
  };
}



function openSheetProtectionDialog(sheetIndex = state.activeSheetIndex) {
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) return;
  closeDataDialog();
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  state.dataDialog = { type: "sheet-protection", sheetIndex };
  const options = defaultSheetProtectionOptions(sheet.protection || {});
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = sheetProtectionDialogHtml(sheet, options);
  document.body.appendChild(overlay);
  overlay.querySelector("[data-data-field='sheetPassword']")?.focus({ preventScroll: true });
}



function sheetProtectionDialogHtml(sheet, options) {
  const checked = (key) => (options[key] ? " checked" : "");
  return `
    <div class="excel-mini-dialog data-dialog sheet-protection-dialog" role="dialog" aria-modal="true" aria-labelledby="sheetProtectionTitle" data-data-dialog="sheet-protection">
      <div class="excel-mini-titlebar" id="sheetProtectionTitle">シートの保護</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="data-dialog-check"><input type="checkbox" data-data-field="protectSheetContents" checked />${accessKeyLabelHtml("シートとロックされたセルの内容を保護する(P)")}</label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("シート保護を解除するためのパスワード(O):")}</span>
          <input type="password" autocomplete="new-password" data-data-field="sheetPassword" value="" />
        </label>
        <fieldset class="data-dialog-fieldset">
          <legend>${accessKeyLabelHtml("このシートのすべてのユーザーに許可する操作(A):")}</legend>
          <div class="sheet-protection-permissions">
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowSelectLockedCells"${checked("allowSelectLockedCells")} />ロックされたセル範囲の選択</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowSelectUnlockedCells"${checked("allowSelectUnlockedCells")} />ロックされていないセル範囲の選択</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowFormatCells"${checked("allowFormatCells")} />セルの書式設定</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowFormatColumns"${checked("allowFormatColumns")} />列の書式設定</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowFormatRows"${checked("allowFormatRows")} />行の書式設定</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowInsertColumns"${checked("allowInsertColumns")} />列の挿入</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowInsertRows"${checked("allowInsertRows")} />行の挿入</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowDeleteColumns"${checked("allowDeleteColumns")} />列の削除</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowDeleteRows"${checked("allowDeleteRows")} />行の削除</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowSort"${checked("allowSort")} />並べ替え</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowAutoFilter"${checked("allowAutoFilter")} />オートフィルターの使用</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowPivotTables"${checked("allowPivotTables")} />ピボットテーブル レポートの使用</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowEditObjects"${checked("allowEditObjects")} />オブジェクトの編集</label>
            <label class="data-dialog-check"><input type="checkbox" data-data-field="allowEditScenarios"${checked("allowEditScenarios")} />シナリオの編集</label>
          </div>
        </fieldset>
        <div class="data-dialog-summary">${escapeHtml(sheet.name)} の保護設定を指定します。</div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
}



function collectSheetProtectionOptions(dialog) {
  const checked = (field) => Boolean(dialog?.querySelector(`[data-data-field="${field}"]`)?.checked);
  return defaultSheetProtectionOptions({
    password: dialog?.querySelector("[data-data-field='sheetPassword']")?.value || "",
    allowSelectLockedCells: checked("allowSelectLockedCells"),
    allowSelectUnlockedCells: checked("allowSelectUnlockedCells"),
    allowFormatCells: checked("allowFormatCells"),
    allowFormatColumns: checked("allowFormatColumns"),
    allowFormatRows: checked("allowFormatRows"),
    allowInsertColumns: checked("allowInsertColumns"),
    allowInsertRows: checked("allowInsertRows"),
    allowDeleteColumns: checked("allowDeleteColumns"),
    allowDeleteRows: checked("allowDeleteRows"),
    allowSort: checked("allowSort"),
    allowAutoFilter: checked("allowAutoFilter"),
    allowPivotTables: checked("allowPivotTables"),
    allowEditObjects: checked("allowEditObjects"),
    allowEditScenarios: checked("allowEditScenarios"),
  });
}



function applySheetProtectionDialog(dialog) {
  const sheetIndex = Number(state.dataDialog?.sheetIndex ?? state.activeSheetIndex);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) return;
  const options = collectSheetProtectionOptions(dialog);
  if (!dialog.querySelector("[data-data-field='protectSheetContents']")?.checked) {
    closeDataDialog();
    setStatus("シート保護は設定されませんでした。");
    return;
  }
  if (options.password) {
    state.dataDialog = { type: "sheet-protection-confirm", sheetIndex, pendingProtection: options };
    renderSheetProtectionConfirmDialog();
    return;
  }
  protectSheetWithOptions(sheetIndex, options);
}



function renderSheetProtectionConfirmDialog() {
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog sheet-protection-confirm-dialog" role="dialog" aria-modal="true" aria-labelledby="sheetProtectionConfirmTitle" data-data-dialog="sheet-protection-confirm">
      <div class="excel-mini-titlebar" id="sheetProtectionConfirmTitle">パスワードの確認</div>
      <div class="excel-mini-body data-dialog-body">
        <p class="data-dialog-summary">確認のため、もう一度パスワードを入力してください。</p>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("パスワードの再入力(R):")}</span>
          <input type="password" autocomplete="new-password" data-data-field="sheetPasswordConfirm" />
        </label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("[data-data-field='sheetPasswordConfirm']")?.focus({ preventScroll: true });
}



function applySheetProtectionConfirmDialog(dialog) {
  const expected = String(state.dataDialog?.pendingProtection?.password || "");
  const actual = String(dialog?.querySelector("[data-data-field='sheetPasswordConfirm']")?.value || "");
  if (actual !== expected) {
    showMessageBox("入力したパスワードが一致しません。もう一度入力してください。", { title: "Cht WebSheet" });
    dialog?.querySelector("[data-data-field='sheetPasswordConfirm']")?.select?.();
    return;
  }
  protectSheetWithOptions(Number(state.dataDialog?.sheetIndex ?? state.activeSheetIndex), state.dataDialog.pendingProtection);
}



function protectSheetWithOptions(sheetIndex, options) {
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) return;
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.protection = defaultSheetProtectionOptions(options);
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: `${sheet.name} シートの保護`,
  });
  closeDataDialog();
  renderWorkbook();
  updateFormulaBarForSelection();
  setStatus(`${sheet.name} を保護しました。`);
}



function openSheetUnprotectDialog(sheetIndex = state.activeSheetIndex) {
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) return;
  const password = String(sheet.protection?.password || "");
  if (!password) {
    unprotectSheet(sheetIndex);
    return;
  }
  closeDataDialog();
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  state.dataDialog = { type: "sheet-unprotect", sheetIndex };
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog sheet-unprotect-dialog" role="dialog" aria-modal="true" aria-labelledby="sheetUnprotectTitle" data-data-dialog="sheet-unprotect">
      <div class="excel-mini-titlebar" id="sheetUnprotectTitle">シート保護の解除</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("パスワード(P):")}</span>
          <input type="password" autocomplete="current-password" data-data-field="sheetUnprotectPassword" />
        </label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("[data-data-field='sheetUnprotectPassword']")?.focus({ preventScroll: true });
}



function applySheetUnprotectDialog(dialog) {
  const sheetIndex = Number(state.dataDialog?.sheetIndex ?? state.activeSheetIndex);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) return;
  const expected = String(sheet.protection?.password || "");
  const actual = String(dialog?.querySelector("[data-data-field='sheetUnprotectPassword']")?.value || "");
  if (expected && actual !== expected) {
    showMessageBox("入力したパスワードが正しくありません。", { title: "Cht WebSheet" });
    dialog?.querySelector("[data-data-field='sheetUnprotectPassword']")?.select?.();
    return;
  }
  unprotectSheet(sheetIndex);
}



function unprotectSheet(sheetIndex = state.activeSheetIndex) {
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) return;
  const sheetBefore = captureSheetForHistory(sheet);
  delete sheet.protection;
  pushHistory({
    sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: `${sheet.name} シート保護の解除`,
  });
  closeDataDialog();
  renderWorkbook();
  updateFormulaBarForSelection();
  setStatus(`${sheet.name} のシート保護を解除しました。`);
}



function sheetProtectionAllows(sheet, permission) {
  if (!isSheetProtected(sheet)) return true;
  return Boolean(sheet.protection?.[permission]);
}



function cellLockedForProtection(sheet, row, col) {
  return sheet?.cells?.[cellKey(row, col)]?.protection?.locked !== false;
}



function rangeHasLockedCellsForProtection(sheet, range) {
  if (!isSheetProtected(sheet) || !range) return false;
  let found = false;
  forEachCellInRange(range, (row, col) => {
    if (!found && cellLockedForProtection(sheet, row, col)) found = true;
  });
  return found;
}



function rangesHaveLockedCellsForProtection(sheet, ranges) {
  return (ranges || []).some((range) => rangeHasLockedCellsForProtection(sheet, range));
}



function showProtectedSheetMessage(message = "変更しようとしているセルまたはグラフは保護されたシート上にあります。変更するには、シート保護を解除してください。") {
  showMessageBox(message, { title: "Cht WebSheet", tone: "warning" });
}



function ensureEditableRangeForProtection(sheet, range, actionLabel = "変更") {
  if (!isSheetProtected(sheet)) return true;
  if (!rangeHasLockedCellsForProtection(sheet, range)) return true;
  showProtectedSheetMessage(`保護されたシート上のロックされたセルは${actionLabel}できません。シート保護を解除するか、セルのロックを解除してください。`);
  return false;
}



function ensureEditableRangesForProtection(sheet, ranges, actionLabel = "変更") {
  if (!isSheetProtected(sheet)) return true;
  if (!rangesHaveLockedCellsForProtection(sheet, ranges)) return true;
  showProtectedSheetMessage(`保護されたシート上のロックされたセルは${actionLabel}できません。シート保護を解除するか、セルのロックを解除してください。`);
  return false;
}



function rangeSelectionTypesForProtection(sheet, range) {
  const types = { locked: false, unlocked: false };
  if (!range) return types;
  forEachCellInRange(range, (row, col) => {
    if (cellLockedForProtection(sheet, row, col)) types.locked = true;
    else types.unlocked = true;
  });
  return types;
}



function ensureSelectableRangeForProtection(sheet, range) {
  if (!isSheetProtected(sheet)) return true;
  const types = rangeSelectionTypesForProtection(sheet, range);
  if (types.locked && !sheetProtectionAllows(sheet, "allowSelectLockedCells")) {
    showProtectedSheetMessage("保護されたシートではロックされたセルを選択できません。");
    return false;
  }
  if (types.unlocked && !sheetProtectionAllows(sheet, "allowSelectUnlockedCells")) {
    showProtectedSheetMessage("保護されたシートではロックされていないセルを選択できません。");
    return false;
  }
  return true;
}



function ensureSheetProtectionPermission(sheet, permission, message) {
  if (sheetProtectionAllows(sheet, permission)) return true;
  showProtectedSheetMessage(message);
  return false;
}



function ensureActiveSheetFormatCellsAllowed(message = "保護されたシートではセルの書式を変更できません。") {
  return ensureSheetProtectionPermission(activeSheet(), "allowFormatCells", message);
}



function ensureHeaderFormatAllowed(kind) {
  return ensureHeaderFormatAllowedForSheet(activeSheet(), kind);
}



function ensureStructureAllowedForProtection(kind, option) {
  return ensureStructureAllowedForSheet(activeSheet(), kind, option);
}



function ensureStructureAllowedForSheet(sheet, kind, option) {
  if (!isSheetProtected(sheet)) return true;
  const insert = kind === "insert";
  if (option === "entire-column") {
    return ensureSheetProtectionPermission(
      sheet,
      insert ? "allowInsertColumns" : "allowDeleteColumns",
      insert ? "保護されたシートでは列を挿入できません。" : "保護されたシートでは列を削除できません。",
    );
  }
  if (option === "entire-row") {
    return ensureSheetProtectionPermission(
      sheet,
      insert ? "allowInsertRows" : "allowDeleteRows",
      insert ? "保護されたシートでは行を挿入できません。" : "保護されたシートでは行を削除できません。",
    );
  }
  showProtectedSheetMessage(insert ? "保護されたシートではセルを挿入できません。" : "保護されたシートではセルを削除できません。");
  return false;
}



function ensureHeaderFormatAllowedForSheet(sheet, kind) {
  const isColumn = kind === "col" || kind === "column";
  return ensureSheetProtectionPermission(
    sheet,
    isColumn ? "allowFormatColumns" : "allowFormatRows",
    isColumn ? "保護されたシートでは列の書式を変更できません。" : "保護されたシートでは行の書式を変更できません。",
  );
}



function openCellContextMenu(event) {
  let point = cellPointFromPointerEvent(event, { allowEditable: true });
  if (!point && event.target?.closest?.(".cell-validation-dropdown-button")) {
    point = cellPointFromElement(event.target.closest(".sheet-cell[data-row][data-col]"));
  }
  if (!point) return;
  event.preventDefault();
  closeSheetTabContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeQuickAnalysisPopup();
  const { row, col } = point;
  if (!state.model || !row || !col) return;

  const currentRanges = activeSelectionRanges();
  if (!currentRanges.length || !currentRanges.some((range) => isCellInRange(range, row, col))) {
    selectCellPoint(point);
  }

  renderCellContextMenu(row, col, event.clientX, event.clientY);
}



function renderCellContextMenu(row, col, clientX, clientY) {
  closeCellContextMenu();
  state.cellContextMenu = { sheetIndex: state.activeSheetIndex, row, col };
  const html = cellContextMenuItems(row, col)
    .map((item) => {
      if (item.type === "separator") return '<div class="cell-context-menu-separator" role="separator"></div>';
      if (item.type === "description") return `<div class="cell-context-menu-note">${escapeHtml(item.text || "")}</div>`;
      if (item.type === "paste-options") return cellPasteOptionsHtml(row, col);
      const disabled = item.disabled ? " disabled" : "";
      const className = item.disabled ? "cell-context-menu-item is-disabled" : "cell-context-menu-item";
      const arrow = item.submenu ? '<span class="cell-context-menu-arrow">›</span>' : "";
      return `<button class="${className}" type="button" role="menuitem" data-cell-menu-action="${escapeAttr(item.action)}" data-row="${row}" data-col="${col}"${disabled}>${accessKeyLabelHtml(item.label)}${arrow}</button>`;
    })
    .join("");
  const menu = document.createElement("div");
  menu.id = "cellContextMenu";
  menu.className = "cell-context-menu";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", `${columnName(col)}${row} menu`);
  menu.innerHTML = html;
  enhanceCommandMenuAccessKeys(menu);
  document.body.append(menu);
  positionPopupMenu(menu, clientX, clientY);
  menu.querySelector("button:not([disabled])")?.focus({ preventScroll: true });
}



function openImageContextMenu(event, options = {}) {
  if (!state.model) return;
  const imageId = options.imageId || event.currentTarget?.dataset?.imageId || event.target?.closest?.("[data-image-id]")?.dataset.imageId || "";
  const image = activeSheet().images?.find((item) => item.id === imageId);
  if (!imageId || !image) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  closeSheetTabContextMenu();
  closeFileSideMenu({ restoreHome: true });
  if (!isActiveImage(imageId)) {
    state.selected = null;
    state.selectionRange = null;
    state.selectionRanges = [];
    state.selectionAnchor = null;
    state.selectionStart = null;
    state.selectionDrag = null;
    setSelectedImageIds([imageId], imageId);
    $formulaInput.prop("disabled", true).val("");
    renderSheet();
    updateSelectionUi();
  }
  const vertexHandle = event.target?.closest?.(".sheet-shape-adjust-handle.is-vertex");
  if (vertexHandle && state.shapePointEditIds?.includes(imageId) && image.shape?.type === CUSTOM_SHAPE_TYPE) {
    renderShapeVertexContextMenu(imageId, vertexHandle.dataset.shapeAdjust || "", event.clientX, event.clientY);
    return;
  }
  renderImageContextMenu(imageId, event.clientX, event.clientY);
}



function renderImageContextMenu(imageId, clientX, clientY) {
  closeCellContextMenu();
  const selectedObjects = selectedImageObjects();
  const objectName = selectedObjects.length === 1 ? selectedObjects[0].name || "オブジェクト" : `${selectedObjects.length} 個のオブジェクト`;
  state.cellContextMenu = { sheetIndex: state.activeSheetIndex, imageId, imageContext: true };
  const html = imageContextMenuItems()
    .map((item) => {
      if (item.type === "separator") return '<div class="cell-context-menu-separator" role="separator"></div>';
      const disabled = item.disabled ? " disabled" : "";
      const className = item.disabled ? "cell-context-menu-item is-disabled" : "cell-context-menu-item";
      return `<button class="${className}" type="button" role="menuitem" data-cell-menu-action="${escapeAttr(item.action)}"${disabled}>${accessKeyLabelHtml(item.label)}</button>`;
    })
    .join("");
  const menu = document.createElement("div");
  menu.id = "cellContextMenu";
  menu.className = "cell-context-menu image-context-menu";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", `${objectName} menu`);
  menu.innerHTML = html;
  enhanceCommandMenuAccessKeys(menu);
  document.body.append(menu);
  positionPopupMenu(menu, clientX, clientY);
  menu.querySelector("button:not([disabled])")?.focus({ preventScroll: true });
}



function renderShapeVertexContextMenu(imageId, vertexKey, clientX, clientY) {
  closeCellContextMenu();
  const image = selectedImageObjectById(imageId);
  state.cellContextMenu = { sheetIndex: state.activeSheetIndex, imageId, imageContext: true, vertexKey };
  const html = shapeVertexContextMenuItems(image, vertexKey)
    .map((item) => {
      if (item.type === "separator") return '<div class="cell-context-menu-separator" role="separator"></div>';
      const disabled = item.disabled ? " disabled" : "";
      const className = item.disabled ? "cell-context-menu-item is-disabled" : "cell-context-menu-item";
      return `<button class="${className}" type="button" role="menuitem" data-cell-menu-action="${escapeAttr(item.action)}"${disabled}>${accessKeyLabelHtml(item.label)}</button>`;
    })
    .join("");
  const menu = document.createElement("div");
  menu.id = "cellContextMenu";
  menu.className = "cell-context-menu image-context-menu shape-vertex-context-menu";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", "頂点 menu");
  menu.innerHTML = html;
  enhanceCommandMenuAccessKeys(menu);
  document.body.append(menu);
  positionPopupMenu(menu, clientX, clientY);
  menu.querySelector("button:not([disabled])")?.focus({ preventScroll: true });
}



function imageContextMenuItems() {
  const selectedObjects = selectedImageObjects();
  const canEditText = selectedObjects.length === 1 && selectedObjects[0]?.shape && !selectedObjects[0]?.chart && !isLineShapeObject(selectedObjects[0]);
  const shapeObjects = selectedObjects.filter((image) => image?.shape);
  const formatLabel = objectFormatMenuLabel(selectedObjects);
  const canGroup = selectedObjects.length >= 2;
  const canUngroup = selectedObjects.some((image) => image.groupId);
  const canEditPoints = shapeObjects.length === 1 && shapeCanEditPoints(shapeObjects[0]);
  const hasHyperlink = selectedObjects.length === 1 && Boolean(selectedObjects[0]?.hyperlink);
  return [
    { action: "object-cut", label: "切り取り(T)" },
    { action: "object-copy", label: "コピー(C)" },
    { action: "object-paste", label: "貼り付け(P)", disabled: !state.objectClipboard?.objects?.length },
    { type: "separator" },
    { action: "object-edit-text", label: "テキストの編集(X)", disabled: !canEditText },
    { action: "object-edit-points", label: "頂点の編集(E)", disabled: !canEditPoints },
    { type: "separator" },
    { action: "object-group", label: "グループ化(G)", disabled: !canGroup },
    { action: "object-ungroup", label: "グループ解除(U)", disabled: !canUngroup },
    { type: "separator" },
    { action: "object-bring-front", label: "最前面へ移動(R)" },
    { action: "object-bring-forward", label: "前面へ移動(F)" },
    { action: "object-send-backward", label: "背面へ移動(K)" },
    { action: "object-send-back", label: "最背面へ移動(B)" },
    { action: "object-selection-pane", label: "選択ウィンドウ(S)..." },
    { type: "separator" },
    { action: "object-link", label: "リンク(I)" },
    { action: "object-open-link", label: "ハイパーリンクを開く(O)", disabled: !hasHyperlink },
    { type: "separator" },
    { action: "object-delete", label: "削除(D)" },
    { type: "separator" },
    { action: "object-format", label: formatLabel },
  ];
}



function objectFormatMenuLabel(objects = selectedImageObjects()) {
  const mode = selectedObjectFormatMode(objects);
  if (mode === "chart") return "グラフ エリアの書式設定(O)...";
  if (mode === "picture") return "図の書式設定(O)...";
  return "図形の書式設定(O)...";
}



function objectFormatPaneTitle(objects = selectedImageObjects()) {
  const mode = selectedObjectFormatMode(objects);
  if (mode === "chart") return "グラフ エリアの書式設定";
  if (mode === "picture") return "図の書式設定";
  return "図形の書式設定";
}



function shapeVertexContextMenuItems(image, vertexKey) {
  const path = normalizeShapeCustomPath(image?.shape?.customPath);
  const vertex = customPathVertexForKey(path, vertexKey);
  const canDelete = Boolean(vertex && vertex.command?.cmd !== "M" && customPathEditablePointCount(path) > 2);
  const closed = shapeCustomPathIsClosed(path);
  return [
    { action: "object-add-vertex", label: "頂点の追加(A)" },
    { action: "object-delete-vertex", label: "頂点の削除(D)", disabled: !canDelete },
    { type: "separator" },
    { action: "object-toggle-open-path", label: closed ? "開いたパス(O)" : "パスを閉じる(C)" },
  ];
}



function openColumnHeaderContextMenu(event) {
  openHeaderContextMenu(event, "column", Number(event.currentTarget.dataset.col));
}



function openRowHeaderContextMenu(event) {
  openHeaderContextMenu(event, "row", Number(event.currentTarget.dataset.row));
}



function openHeaderContextMenu(event, kind, index) {
  if (!state.model || !index) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  closeSheetTabContextMenu();
  closeFileSideMenu({ restoreHome: true });

  const range = kind === "column" ? columnSelectionRange(index, index) : rowSelectionRange(index, index);
  const point = kind === "column"
    ? { sheetIndex: state.activeSheetIndex, row: 1, col: index }
    : { sheetIndex: state.activeSheetIndex, row: index, col: 1 };
  if (!headerIndexInActiveSelection(kind, index)) {
    applySelectionRange(range, point);
    state.selectionAnchor = point;
    state.selectionStart = point;
    state.selectionDrag = null;
    updateSelectionUi();
    updateFormulaBarForSelection();
  }

  renderHeaderContextMenu(kind, index, event.clientX, event.clientY);
}



function headerIndexInActiveSelection(kind, index) {
  return activeSelectionRanges().some((range) => {
    if (kind === "column") return range.fullColumn && index >= range.left && index <= range.right;
    return range.fullRow && index >= range.top && index <= range.bottom;
  });
}



function renderHeaderContextMenu(kind, index, clientX, clientY) {
  closeCellContextMenu();
  const row = kind === "row" ? index : 1;
  const col = kind === "column" ? index : 1;
  state.cellContextMenu = { sheetIndex: state.activeSheetIndex, row, col, headerKind: kind, headerIndex: index };
  const html = headerContextMenuItems(kind)
    .map((item) => {
      if (item.type === "separator") return '<div class="cell-context-menu-separator" role="separator"></div>';
      if (item.type === "description") return `<div class="cell-context-menu-note">${escapeHtml(item.text || "")}</div>`;
      if (item.type === "paste-options") return cellPasteOptionsHtml(row, col);
      const disabled = item.disabled ? " disabled" : "";
      const className = item.disabled ? "cell-context-menu-item is-disabled" : "cell-context-menu-item";
      return `<button class="${className}" type="button" role="menuitem" data-cell-menu-action="${escapeAttr(item.action)}" data-row="${row}" data-col="${col}"${disabled}>${accessKeyLabelHtml(item.label)}</button>`;
    })
    .join("");
  const menu = document.createElement("div");
  menu.id = "cellContextMenu";
  menu.className = "cell-context-menu header-context-menu";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", kind === "column" ? `${columnName(index)} column menu` : `${index} row menu`);
  menu.innerHTML = html;
  enhanceCommandMenuAccessKeys(menu);
  document.body.append(menu);
  positionPopupMenu(menu, clientX, clientY);
  menu.querySelector("button:not([disabled])")?.focus({ preventScroll: true });
}



function headerContextMenuItems(kind) {
  const isColumn = kind === "column";
  const hasCopiedHeaderCells = canInsertCopiedHeaderCells(kind);
  const insertItem = hasCopiedHeaderCells
    ? { action: isColumn ? "insert-copied-column" : "insert-copied-row", label: "コピーしたセルの挿入(E)" }
    : { action: isColumn ? "insert-entire-column" : "insert-entire-row", label: "挿入(I)" };
  const items = [
    { action: "cut", label: "切り取り(T)" },
    { action: "copy", label: "コピー(C)" },
    { type: "paste-options" },
    { type: "separator" },
    insertItem,
    { action: isColumn ? "delete-entire-column" : "delete-entire-row", label: "削除(D)" },
    { action: "clear", label: "数式と値のクリア(N)" },
    { type: "separator" },
    { action: "format-cells", label: "セルの書式設定(F)..." },
    { action: isColumn ? "column-width" : "row-height", label: isColumn ? "列の幅(C)..." : "行の高さ(R)..." },
    { type: "separator" },
    { action: isColumn ? "hide-column" : "hide-row", label: "非表示(H)" },
    { action: isColumn ? "unhide-column" : "unhide-row", label: "再表示(U)" },
  ];
  return addPageBreakContextItems(items, pageBreakTargetsForHeaderContext(kind, state.cellContextMenu?.headerIndex || 1));
}



function canInsertCopiedHeaderCells(kind) {
  const sourceRange = state.internalClipboard?.sourceRange;
  if (!sourceRange) return false;
  return kind === "column" ? Boolean(sourceRange.fullColumn) : Boolean(sourceRange.fullRow);
}



function cellContextMenuItems(row, col) {
  const cell = activeSheet().cells[cellKey(row, col)];
  const hasHyperlink = Boolean(cell?.hyperlink);
  const hasFurigana = contextMenuHasPhonetics(row, col);
  const hasValidationDropdown = contextMenuHasValidationDropdown(row, col);
  const hasComment = cellHasComment(cell);
  const commentItems = hasComment
    ? [
        { action: "reply-comment", label: "コメントに返信する(R)" },
        { action: "delete-comment", label: "コメントの削除(M)" },
      ]
    : [{ action: "new-comment", label: "新しいコメント(M)" }];
  const hasMemo = cellHasNote(cell);
  const memoItems = hasMemo
    ? [
        { action: "edit-note", label: "メモの編集(E)" },
        { action: "delete-note", label: "メモの削除(T)" },
        { action: "toggle-note", label: "メモの表示/非表示(O)" },
      ]
    : [{ action: "new-note", label: "新しいメモ(N)" }];
  const items = [
    { action: "cut", label: "切り取り(T)" },
    { action: "copy", label: "コピー(C)" },
    { type: "paste-options" },
    { type: "separator" },
    { action: "translate", label: "翻訳" },
    { type: "separator" },
    { action: "insert", label: "挿入(I)..." },
    { action: "delete", label: "削除(D)..." },
    { action: "clear", label: "数式と値のクリア(N)" },
    { type: "separator" },
    { action: "quick-analysis", label: "クイック分析(Q)", disabled: !quickAnalysisAvailableForContext(row, col) },
    { action: "filter", label: "フィルター(E)", submenu: true },
    { action: "sort", label: "並び替え(O)", submenu: true },
    { type: "separator" },
    { action: "get-data", label: "テーブルまたは範囲からデータを取得(G)..." },
    { type: "separator" },
    ...commentItems,
    ...memoItems,
    { type: "separator" },
    { action: "format-cells", label: "セルの書式設定(F)..." },
    { action: "select-dropdown", label: "ドロップダウン　リストから選択(K)...", disabled: !hasValidationDropdown },
    { action: "furigana", label: "ふりがなの表示(S)", disabled: !hasFurigana },
    { action: "define-name", label: "名前の定義(A)..." },
    { type: "separator" },
    { action: "link", label: "リンク(I)" },
    { action: "open-hyperlink", label: "ハイパーリンクを開く(O)", disabled: !hasHyperlink },
  ];
  return addPageBreakContextItems(items, pageBreakTargetsForSelection(activeSheet()));
}



function addPageBreakContextItems(items, targets) {
  const sheet = activeSheet();
  if ((sheet?.viewMode || "normal") !== "page-break") return items;
  const layout = ensurePageLayoutSettings(sheet);
  const hasManualBreak = pageLayoutCurrentSelectionBreakState(layout, targets);
  const hasAnyBreaks = Boolean(layout.pageBreaks?.rows?.length || layout.pageBreaks?.cols?.length);
  const label = pageBreakTargetsLabel(targets);
  return [
    ...items,
    { type: "separator" },
    { action: "insert-page-break-selection", label: "改ページの挿入(I)", disabled: !targets.canInsert || hasManualBreak },
    { action: "remove-page-break-selection", label: "改ページの解除(R)", disabled: !hasManualBreak },
    { action: "reset-page-breaks", label: "すべての改ページを解除(A)", disabled: !hasAnyBreaks },
    { action: "close-page-break-preview", label: "改ページ プレビューを閉じる(V)" },
    ...(targets.canInsert && !hasManualBreak ? [{ type: "description", text: `${label}に改ページを挿入します。` }] : []),
  ];
}



function contextMenuHasValidationDropdown(row, col) {
  const sheet = activeSheet();
  if (!sheet) return false;
  const cell = sheet.cells[cellKey(row, col)];
  return Boolean(cell?.validation?.type === "list" && validationListOptions(sheet, cell.validation).length);
}



function contextMenuHasPhonetics(row, col) {
  const sheet = activeSheet();
  const ranges = contextMenuTargetRanges(row, col);
  if (!sheet || !ranges.length) return false;
  return Object.values(sheet.cells || {}).some((cell) =>
    cellHasPhonetics(cell) && ranges.some((range) => isCellInRange(range, cell.row, cell.col)),
  );
}



function contextMenuTargetRanges(row, col) {
  const ranges = activeSelectionRanges();
  if (ranges.some((range) => isCellInRange(range, row, col))) {
    return ranges;
  }
  const point = { sheetIndex: state.activeSheetIndex, row, col };
  const range = normalizeRange(point, point);
  const expanded = expandRangeForMerges(activeSheet(), range);
  return expanded ? [expanded] : [];
}



function openTranslationPane(row, col) {
  const context = translationContextFromSelection(row, col);
  if (!context) return;
  const detectedLanguage = detectLanguageForTranslationItems(context.items);
  state.translationPane = {
    sheetIndex: state.activeSheetIndex,
    ranges: context.ranges.map((range) => ({ ...range })),
    sourceLang: "auto",
    targetLang: defaultTranslationTargetLanguage(detectedLanguage),
  };
  renderTranslationPane();
  setStatus(`${rangesToLabel(context.ranges)} の翻訳ペインを開きました。`);
}



function translationContextFromSelection(row = null, col = null) {
  if (!state.model) return null;
  const sheet = activeSheet();
  const ranges = (row && col ? contextMenuTargetRanges(row, col) : activeSelectionRanges())
    .map(materializeSelectionRange)
    .filter(Boolean);
  if (!ranges.length) {
    showMessageBox("翻訳するセル範囲を選択してください。", { title: "Cht WebSheet" });
    return null;
  }
  const scanCount = ranges.reduce((total, range) => total + rangeCellCountForTranslation(range), 0);
  if (scanCount > TRANSLATION_SCAN_LIMIT) {
    showMessageBox("選択範囲が大きすぎるため、翻訳するセル範囲を小さくしてください。", { title: "Cht WebSheet" });
    return null;
  }
  const items = translationItemsForRanges(sheet, ranges, { limit: TRANSLATION_CELL_LIMIT });
  if (!items.length) {
    showMessageBox("翻訳できる文字列が選択範囲にありません。", { title: "Cht WebSheet" });
    return null;
  }
  return { sheet, ranges, items };
}



function currentTranslationContext() {
  const pane = state.translationPane;
  if (!pane || pane.sheetIndex !== state.activeSheetIndex || !state.model) return null;
  const sheet = activeSheet();
  const ranges = (pane.ranges || []).map(materializeSelectionRange).filter(Boolean);
  const items = translationItemsForRanges(sheet, ranges, { limit: TRANSLATION_CELL_LIMIT });
  return ranges.length && items.length ? { sheet, ranges, items } : null;
}



function rangeCellCountForTranslation(range) {
  if (!range) return 0;
  return Math.max(0, range.bottom - range.top + 1) * Math.max(0, range.right - range.left + 1);
}



function translationItemsForRanges(sheet, ranges, options = {}) {
  const limit = options.limit || TRANSLATION_CELL_LIMIT;
  const items = [];
  const seen = new Set();
  for (const range of ranges || []) {
    for (let row = range.top; row <= range.bottom; row += 1) {
      for (let col = range.left; col <= range.right; col += 1) {
        const key = cellKey(row, col);
        if (seen.has(key)) continue;
        seen.add(key);
        const cell = sheet.cells[key];
        const source = translationSourceTextForCell(cell);
        if (source.trim()) {
          items.push({ row, col, address: `${columnName(col)}${row}`, source });
          if (items.length >= limit) return items;
        }
      }
    }
  }
  return items;
}



function translationSourceTextForCell(cell) {
  if (!cell) return "";
  return String(cellDisplayText(cell) || cell.raw || "");
}



function defaultTranslationTargetLanguage(sourceLang) {
  return sourceLang === "zh-CN" ? "ja" : "zh-CN";
}



function detectLanguageForTranslationItems(items) {
  const sample = (items || []).map((item) => item.source).join("\n").slice(0, 2000);
  return detectTranslationLanguage(sample);
}



function detectTranslationLanguage(text) {
  const value = String(text || "");
  if (!value) return "ja";
  if (/[\u3040-\u30ff]/.test(value)) return "ja";
  const knownJapanese = Object.keys(OFFLINE_TRANSLATION_DICTIONARIES["ja>zh-CN"] || {});
  if (knownJapanese.some((term) => term && value.includes(term))) return "ja";
  if (/[\u4e00-\u9fff]/.test(value)) return "zh-CN";
  if (/[a-z]/i.test(value)) return "en";
  if (/[\uac00-\ud7af]/.test(value)) return "ko";
  return "ja";
}



function renderTranslationPane() {
  document.querySelector("#translationPane")?.remove();
  const context = currentTranslationContext();
  if (!state.translationPane || !context) {
    state.translationPane = null;
    return;
  }
  const paneState = state.translationPane;
  const sourceLang = paneState.sourceLang || "auto";
  const targetLang = paneState.targetLang || "zh-CN";
  const detectedLang = detectLanguageForTranslationItems(context.items);
  const previewItems = context.items.slice(0, TRANSLATION_PREVIEW_LIMIT);
  const translatedPreview = previewItems.map((item) => ({
    ...item,
    translation: translateTextOffline(item.source, targetLang, sourceLang).text,
  }));
  const truncated = context.items.length > previewItems.length;
  const pane = document.createElement("aside");
  pane.id = "translationPane";
  pane.className = "translation-pane";
  pane.setAttribute("role", "complementary");
  pane.setAttribute("aria-label", "翻訳");
  pane.innerHTML = `
    <div class="translation-pane-header">
      <div>
        <strong>翻訳</strong>
        <span>${escapeHtml(rangesToLabel(context.ranges))}</span>
      </div>
      <button type="button" data-translation-pane-action="close" aria-label="閉じる">×</button>
    </div>
    <div class="translation-pane-body">
      <div class="translation-pane-controls">
        <label>
          <span>翻訳元</span>
          <select data-translation-pane-field="sourceLang">
            ${translationLanguageOptions(sourceLang, true)}
          </select>
        </label>
        <button type="button" class="translation-pane-swap" data-translation-pane-action="swap" title="言語を入れ替え"><i data-lucide="move-right"></i></button>
        <label>
          <span>翻訳先</span>
          <select data-translation-pane-field="targetLang">
            ${translationLanguageOptions(targetLang, false)}
          </select>
        </label>
      </div>
      <div class="translation-pane-detected">検出：${escapeHtml(languageLabel(detectedLang))} / ${escapeHtml(context.items.length)} 個のセル</div>
      <div class="translation-pane-preview" role="list">
        ${translatedPreview.map(translationPreviewRowHtml).join("")}
        ${truncated ? `<div class="translation-pane-more">ほか ${context.items.length - previewItems.length} 個のセル</div>` : ""}
      </div>
      <div class="translation-pane-note">オフライン辞書にない語句は原文を保持します。</div>
    </div>
    <div class="translation-pane-footer">
      <button type="button" data-translation-pane-action="insert-right">右側に挿入</button>
      <button type="button" data-translation-pane-action="replace">選択範囲を置換</button>
    </div>
  `;
  document.body.append(pane);
  createIcons({ icons });
}



function translationLanguageOptions(selected, includeAuto) {
  return TRANSLATION_LANGUAGES
    .filter((language) => includeAuto || language.value !== "auto")
    .map((language) => `<option value="${escapeAttr(language.value)}"${language.value === selected ? " selected" : ""}>${escapeHtml(language.label)}</option>`)
    .join("");
}



function languageLabel(value) {
  return TRANSLATION_LANGUAGES.find((language) => language.value === value)?.label || value || "";
}



function translationPreviewRowHtml(item) {
  return `
    <div class="translation-pane-row" role="listitem">
      <div class="translation-pane-address">${escapeHtml(item.address)}</div>
      <div class="translation-pane-text">
        <div class="translation-pane-source">${escapeHtml(item.source)}</div>
        <div class="translation-pane-result">${escapeHtml(item.translation)}</div>
      </div>
    </div>
  `;
}



function handleTranslationPaneFieldChange(event) {
  const field = event.currentTarget.dataset.translationPaneField;
  if (!state.translationPane || !field) return;
  state.translationPane[field] = event.currentTarget.value;
  renderTranslationPane();
}



function handleTranslationPaneAction(event) {
  const action = event.currentTarget.dataset.translationPaneAction;
  if (!action) return;
  event.preventDefault();
  event.stopPropagation();
  if (action === "close") {
    closeTranslationPane();
    return;
  }
  if (action === "swap") {
    swapTranslationPaneLanguages();
    return;
  }
  if (action === "insert-right" || action === "replace") {
    applyTranslationPaneResult(action);
  }
}



function swapTranslationPaneLanguages() {
  const pane = state.translationPane;
  const context = currentTranslationContext();
  if (!pane || !context) return;
  const source = pane.sourceLang === "auto" ? detectLanguageForTranslationItems(context.items) : pane.sourceLang;
  pane.sourceLang = pane.targetLang || "zh-CN";
  pane.targetLang = source === "auto" ? defaultTranslationTargetLanguage(detectLanguageForTranslationItems(context.items)) : source;
  renderTranslationPane();
}



function applyTranslationPaneResult(action) {
  const context = currentTranslationContext();
  const pane = state.translationPane;
  if (!context || !pane) return;
  const sheet = context.sheet;
  const sheetId = getSheetId(sheet.name);
  if (sheetId == null) return;
  const mode = action === "insert-right" ? "right" : "replace";
  const targetPlan = translationTargetPlan(context, mode);
  if (!targetPlan.items.length) return;
  if (applyTranslationPaneResultToSheetGroup(context, pane, mode)) return;
  if (!ensureEditableRangesForProtection(sheet, targetPlan.ranges, "翻訳")) return;

  const sheetBefore = mode === "right" ? captureSheetForHistory(sheet) : null;
  const before = mode === "replace" ? captureCellsForHistoryRanges(sheet, targetPlan.ranges) : null;
  const maxRow = Math.max(...targetPlan.items.map((item) => item.targetRow));
  const maxCol = Math.max(...targetPlan.items.map((item) => item.targetCol));
  try {
    ensureSheetSize(sheet, maxRow, maxCol);
    targetPlan.items.forEach((item) => {
      const translated = translateTextOffline(item.source, pane.targetLang, pane.sourceLang).text;
      writeCellContent(sheet, sheetId, item.targetRow, item.targetCol, translated);
    });
    internCellsInRanges(sheet, targetPlan.ranges);
    state.hf = buildFormulaEngine(state.model);
    const selectionBefore = activeSelectionRange();
    const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
    const selectionAfter = targetPlan.ranges[0];
    state.selected = { sheetIndex: state.activeSheetIndex, row: selectionAfter.top, col: selectionAfter.left };
    state.selectionRange = { ...selectionAfter };
    state.selectionRanges = targetPlan.ranges.map((range) => ({ ...range }));
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
    const history = {
      sheetIndex: state.activeSheetIndex,
      selectionBefore,
      selectionAfter,
      selectionBeforeRanges,
      selectionAfterRanges: targetPlan.ranges,
      label: `${rangesToLabel(targetPlan.ranges)} 翻訳`,
    };
    if (mode === "right") {
      history.sheetBefore = sheetBefore;
      history.sheetAfter = captureSheetForHistory(sheet);
    } else {
      history.before = before;
      history.after = captureCellsForHistoryRanges(sheet, targetPlan.ranges);
    }
    pushHistory(history);
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    renderTranslationPane();
    setStatus(`${targetPlan.items.length} 個のセルに翻訳を${mode === "right" ? "挿入" : "適用"}しました。`);
  } catch (error) {
    console.error(error);
    if (sheetBefore) restoreSheetFromHistory(sheet, sheetBefore);
    state.hf = buildFormulaEngine(state.model);
    renderSheet();
    setStatus(`翻訳に失敗しました：${error.message || error}`);
  }
}



function applyTranslationPaneResultToSheetGroup(context, pane, mode) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !context?.ranges?.length) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    const sheetId = getSheetId(sheet?.name);
    if (!sheet || sheetId == null) continue;
    const sourceRanges = materializedRangesForSheet(sheet, context.ranges, sheetIndex);
    if (!sourceRanges.length) continue;
    const sourceItems = translationItemsForRanges(sheet, sourceRanges, { limit: TRANSLATION_CELL_LIMIT });
    if (!sourceItems.length) continue;
    const plan = translationTargetPlan({ sheet, ranges: sourceRanges, items: sourceItems }, mode);
    if (!plan.items.length) continue;
    if (!ensureEditableRangesForProtection(sheet, plan.ranges, "翻訳")) return true;
    items.push({ sheetIndex, sheet, sheetId, plan, ranges: plan.ranges });
  }
  if (!items.length) return true;
  const affectedSheetIndexes = affectedSheetIndexesForGroupedRanges(items);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  const activeItem = items.find((item) => item.sheetIndex === state.activeSheetIndex) || items[0];
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  try {
    items.forEach((item) => {
      const maxRow = Math.max(...item.plan.items.map((entry) => entry.targetRow));
      const maxCol = Math.max(...item.plan.items.map((entry) => entry.targetCol));
      ensureSheetSize(item.sheet, maxRow, maxCol);
      item.plan.items.forEach((entry) => {
        const translated = translateTextOffline(entry.source, pane.targetLang, pane.sourceLang).text;
        writeCellContent(item.sheet, item.sheetId, entry.targetRow, entry.targetCol, translated);
      });
      internCellsInRanges(item.sheet, item.plan.ranges);
    });
    state.hf = buildFormulaEngine(state.model);
    const selectionAfter = activeItem?.plan?.ranges?.[0] || context.ranges[0];
    state.selected = { sheetIndex: state.activeSheetIndex, row: selectionAfter.top, col: selectionAfter.left };
    state.selectionRange = { ...selectionAfter, sheetIndex: state.activeSheetIndex };
    state.selectionRanges = (activeItem?.plan?.ranges || [selectionAfter]).map((range) => ({ ...range, sheetIndex: state.activeSheetIndex }));
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
      selectionAfter,
      selectionBeforeRanges,
      selectionAfterRanges: state.selectionRanges.map((range) => ({ ...range })),
      label: `${rangesToLabel(activeItem?.plan?.ranges || context.ranges)} ${sheetIndexes.length} シート翻訳`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    renderTranslationPane();
    const changed = items.reduce((total, item) => total + item.plan.items.length, 0);
    setStatus(`${sheetIndexes.length} 個のシートの ${changed} 個のセルに翻訳を${mode === "right" ? "挿入" : "適用"}しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`翻訳に失敗しました：${error.message || error}`);
  }
  return true;
}



function translationTargetPlan(context, mode) {
  const ranges = [];
  const items = [];
  const seen = new Set();
  context.ranges.forEach((sourceRange) => {
    const width = sourceRange.right - sourceRange.left + 1;
    const targetRange = mode === "right"
      ? { ...sourceRange, left: sourceRange.right + 1, right: sourceRange.right + width }
      : { ...sourceRange };
    ranges.push(targetRange);
    for (const item of context.items) {
      if (!isCellInRange(sourceRange, item.row, item.col)) continue;
      const targetRow = item.row;
      const targetCol = mode === "right" ? targetRange.left + (item.col - sourceRange.left) : item.col;
      const key = cellKey(targetRow, targetCol);
      if (seen.has(key)) continue;
      seen.add(key);
      items.push({ ...item, targetRow, targetCol });
    }
  });
  return { ranges, items };
}



function translateTextOffline(text, targetLang, sourceLang = "auto") {
  const source = sourceLang === "auto" ? detectTranslationLanguage(text) : sourceLang;
  if (!targetLang || source === targetLang) return { text: String(text ?? ""), changed: false };
  const direct = translateTextWithDictionary(text, `${source}>${targetLang}`);
  if (direct.changed) return direct;
  if (source !== "ja" && targetLang !== "ja") {
    const viaJapanese = translateTextWithDictionary(text, `${source}>ja`);
    if (viaJapanese.changed) {
      const final = translateTextWithDictionary(viaJapanese.text, `ja>${targetLang}`);
      if (final.changed) return final;
    }
  }
  return { text: String(text ?? ""), changed: false };
}



function translateTextWithDictionary(text, key) {
  const dictionary = OFFLINE_TRANSLATION_DICTIONARIES[key];
  const sourceText = String(text ?? "");
  if (!dictionary) return { text: sourceText, changed: false };
  let result = sourceText;
  let changed = false;
  Object.entries(dictionary)
    .sort((a, b) => b[0].length - a[0].length)
    .forEach(([from, to]) => {
      const next = result.replace(new RegExp(escapeRegExp(from), "g"), to);
      if (next !== result) changed = true;
      result = next;
    });
  return { text: result, changed };
}



function closeTranslationPaneOnEscape(event) {
  if (event.key === "Escape" && state.translationPane) closeTranslationPane();
}



function closeTranslationPane() {
  document.querySelector("#translationPane")?.remove();
  state.translationPane = null;
}



function quickAnalysisContext(row = null, col = null) {
  if (!state.model) return null;
  const ranges = row && col ? contextMenuTargetRanges(row, col) : activeSelectionRanges();
  if (ranges.length !== 1) return null;
  const range = materializeSelectionRange(ranges[0]);
  if (!range || range.fullRow || range.fullColumn) return null;
  const cellCount = (range.bottom - range.top + 1) * (range.right - range.left + 1);
  if (cellCount <= 1 || cellCount > DATA_OPERATION_CELL_LIMIT) return null;
  return { sheet: activeSheet(), range };
}



function activeQuickAnalysisContext() {
  if (!state.model) return null;
  const storedRange = state.quickAnalysisPopup?.range ? materializeSelectionRange(state.quickAnalysisPopup.range) : null;
  if (storedRange) return { sheet: activeSheet(), range: storedRange };
  return quickAnalysisContext();
}



function openQuickAnalysisPopup(row, col) {
  const context = quickAnalysisContext(row, col);
  if (!context) {
    showMessageBox("クイック分析を使用するデータ範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeRibbonColorPalette();
  closeRibbonBorderMenu();
  closeRibbonUnderlineMenu();
  closeRibbonShapeMenu();
  closeRibbonInkEraserMenu();
  closeShapeFormatMenu();
  closeDataCommandMenu();
  closeDataValidationMenu();
  closeQuickAnalysisPopup();

  const popup = document.createElement("div");
  popup.id = "quickAnalysisPopup";
  popup.className = "quick-analysis-popup";
  popup.setAttribute("role", "dialog");
  popup.setAttribute("aria-label", "クイック分析");
  state.quickAnalysisPopup = {
    tab: "format",
    range: { ...context.range },
    sheetIndex: state.activeSheetIndex,
  };
  popup.innerHTML = quickAnalysisPopupHtml("format");
  enhanceCommandMenuAccessKeys(popup);
  document.body.append(popup);
  createIcons({ icons });
  positionQuickAnalysisPopup(popup, context.range);
  popup.querySelector("[data-quick-analysis-action]")?.focus({ preventScroll: true });
}



function quickAnalysisPopupHtml(activeTab) {
  const tabs = QUICK_ANALYSIS_TABS.map((tab) => {
    const active = tab.id === activeTab;
    return `<button class="quick-analysis-tab${active ? " is-active" : ""}" type="button" role="tab" aria-selected="${active}" data-quick-analysis-tab="${escapeAttr(tab.id)}"><i data-lucide="${escapeAttr(tab.icon)}"></i><span>${escapeHtml(tab.label)}</span></button>`;
  }).join("");
  const actions = (QUICK_ANALYSIS_ACTIONS[activeTab] || QUICK_ANALYSIS_ACTIONS.format)
    .map((item) => `<button class="quick-analysis-action" type="button" data-quick-analysis-action="${escapeAttr(item.action)}" title="${escapeAttr(item.label)}"><i data-lucide="${escapeAttr(item.icon)}"></i><span>${escapeHtml(item.label)}</span></button>`)
    .join("");
  return `
    <div class="quick-analysis-tabs" role="tablist" aria-label="クイック分析の種類">${tabs}</div>
    <div class="quick-analysis-panel" role="tabpanel">${actions}</div>`;
}



function positionQuickAnalysisPopup(popup, range) {
  const frame = document.querySelector(".sheet-selection-frame:not(.is-copy-source)");
  const frameRect = frame?.getBoundingClientRect?.();
  if (frameRect && (frameRect.width || frameRect.height)) {
    positionPopupMenu(popup, frameRect.right - 28, frameRect.bottom + 6);
    return;
  }
  const cell = cellElementForPoint({ sheetIndex: state.activeSheetIndex, row: range.top, col: range.left });
  const rect = cell?.getBoundingClientRect?.();
  if (rect && (rect.width || rect.height)) {
    positionPopupMenu(popup, rect.left, rect.bottom + 6);
    return;
  }
  positionPopupMenu(popup, window.innerWidth / 2 - 140, window.innerHeight / 2 - 80);
}



function handleQuickAnalysisTabClick(event) {
  const popup = event.currentTarget.closest("#quickAnalysisPopup");
  if (!popup || !state.quickAnalysisPopup) return;
  const tab = event.currentTarget.dataset.quickAnalysisTab || "format";
  state.quickAnalysisPopup.tab = tab;
  popup.innerHTML = quickAnalysisPopupHtml(tab);
  enhanceCommandMenuAccessKeys(popup);
  createIcons({ icons });
  popup.querySelector(`[data-quick-analysis-tab="${cssEscape(tab)}"]`)?.focus({ preventScroll: true });
}



function handleQuickAnalysisAction(event) {
  const button = event.currentTarget;
  if (!button.closest("#quickAnalysisPopup") || button.disabled) return;
  event.preventDefault();
  event.stopPropagation();
  const action = button.dataset.quickAnalysisAction;
  runQuickAnalysisAction(action);
  closeQuickAnalysisPopup();
}



function runQuickAnalysisAction(action) {
  switch (action) {
    case "format-data-bars":
      applyQuickAnalysisDataBars();
      break;
    case "format-color-scale":
      applyQuickAnalysisColorScale();
      break;
    case "format-icon-set":
      applyQuickAnalysisIconSet();
      break;
    case "format-greater-than":
      applyQuickAnalysisGreaterThanAverage();
      break;
    case "format-duplicates":
      applyQuickAnalysisDuplicateValues();
      break;
    case "format-clear":
      clearQuickAnalysisFormats();
      break;
    case "chart-column":
      insertChartFromSelection("column");
      break;
    case "chart-bar":
      insertChartFromSelection("bar");
      break;
    case "chart-line":
      insertChartFromSelection("line");
      break;
    case "chart-pie":
      insertChartFromSelection("pie");
      break;
    case "chart-scatter":
      insertChartFromSelection("scatter");
      break;
    case "chart-recommended":
      insertChartFromSelection(recommendedChartTypeForSelection());
      break;
    case "total-sum":
      applyQuickAnalysisTotal("sum");
      break;
    case "total-average":
      applyQuickAnalysisTotal("average");
      break;
    case "total-count":
      applyQuickAnalysisTotal("count");
      break;
    case "total-percent":
      applyQuickAnalysisTotal("percent");
      break;
    case "total-running":
      applyQuickAnalysisTotal("running");
      break;
    case "table":
      insertTableFromSelection();
      break;
    case "pivot-table":
      insertPivotSummaryFromSelection();
      break;
    case "sparkline-line":
      insertSparklinesFromSelection("line");
      break;
    case "sparkline-column":
      insertSparklinesFromSelection("column");
      break;
    case "sparkline-winloss":
      insertSparklinesFromSelection("winloss");
      break;
    default:
      break;
  }
}



function closeQuickAnalysisPopup() {
  document.getElementById("quickAnalysisPopup")?.remove();
  state.quickAnalysisPopup = null;
}



function closeQuickAnalysisPopupOnOutsidePointer(event) {
  if (!state.quickAnalysisPopup) return;
  if (event.target.closest?.("#quickAnalysisPopup")) return;
  closeQuickAnalysisPopup();
}



function closeQuickAnalysisPopupOnEscape(event) {
  if (event.key === "Escape" && state.quickAnalysisPopup) {
    consumeCommandMenuKey(event);
    closeQuickAnalysisPopup();
  }
}



function quickAnalysisNumericCells(sheet, range) {
  const sheetId = getSheetId(sheet.name);
  const cells = [];
  forEachCellInRange(range, (row, col) => {
    const value = cellNumericValue(sheet, sheetId, row, col);
    if (Number.isFinite(value)) cells.push({ row, col, value });
  });
  return cells;
}



function quickAnalysisNumericColumns(sheet, range) {
  const sheetId = getSheetId(sheet.name);
  const columns = [];
  for (let col = range.left; col <= range.right; col += 1) {
    let count = 0;
    for (let row = range.top; row <= range.bottom; row += 1) {
      if (Number.isFinite(cellNumericValue(sheet, sheetId, row, col))) count += 1;
    }
    if (count) columns.push({ col, count });
  }
  return columns;
}



function applyQuickAnalysisDataBars() {
  const context = activeQuickAnalysisContext();
  if (!context || !ensureActiveSheetFormatCellsAllowed()) return;
  const { sheet, range } = context;
  const numericCells = quickAnalysisNumericCells(sheet, range);
  if (!numericCells.length) {
    showMessageBox("データ バーを適用できる数値データが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  applyQuickAnalysisCellStyle(range, "クイック分析 データ バー", quickAnalysisApplyDataBars, `${rangeToLabel(range)} にデータ バーを適用しました。`);
}



function applyQuickAnalysisColorScale() {
  const context = activeQuickAnalysisContext();
  if (!context || !ensureActiveSheetFormatCellsAllowed()) return;
  const { sheet, range } = context;
  const numericCells = quickAnalysisNumericCells(sheet, range);
  if (!numericCells.length) {
    showMessageBox("カラー スケールを適用できる数値データが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  applyQuickAnalysisCellStyle(range, "クイック分析 カラー スケール", quickAnalysisApplyColorScale, `${rangeToLabel(range)} にカラー スケールを適用しました。`);
}



function applyQuickAnalysisIconSet() {
  const context = activeQuickAnalysisContext();
  if (!context || !ensureActiveSheetFormatCellsAllowed()) return;
  const { sheet, range } = context;
  const numericCells = quickAnalysisNumericCells(sheet, range);
  if (!numericCells.length) {
    showMessageBox("アイコン セットを適用できる数値データが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  applyQuickAnalysisCellStyle(range, "クイック分析 アイコン セット", quickAnalysisApplyIconSet, `${rangeToLabel(range)} にアイコン セットを適用しました。`);
}



function applyQuickAnalysisGreaterThanAverage() {
  const context = activeQuickAnalysisContext();
  if (!context || !ensureActiveSheetFormatCellsAllowed()) return;
  const { sheet, range } = context;
  const numericCells = quickAnalysisNumericCells(sheet, range);
  if (!numericCells.length) {
    showMessageBox("比較できる数値データが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  applyQuickAnalysisCellStyle(range, "クイック分析 指定の値より大きい", quickAnalysisApplyGreaterThanAverage, `${rangeToLabel(range)} で平均より大きい値を強調しました。`);
}



function applyQuickAnalysisDuplicateValues() {
  const context = activeQuickAnalysisContext();
  if (!context || !ensureActiveSheetFormatCellsAllowed()) return;
  const { sheet, range } = context;
  const sheetId = getSheetId(sheet.name);
  const buckets = new Map();
  forEachCellInRange(range, (row, col) => {
    const text = String(getDisplayForCell(sheet, sheetId, row, col) ?? "").trim();
    if (!text) return;
    const key = text.toLocaleLowerCase("ja-JP");
    const list = buckets.get(key) || [];
    list.push({ row, col });
    buckets.set(key, list);
  });
  const duplicates = Array.from(buckets.values()).filter((list) => list.length > 1).flat();
  if (!duplicates.length) {
    showMessageBox("選択範囲に重複する値はありません。", { title: "Cht WebSheet" });
    return;
  }
  applyQuickAnalysisCellStyle(range, "クイック分析 重複する値", quickAnalysisApplyDuplicateValues, `${rangeToLabel(range)} の重複する値を強調しました。`);
}



function clearQuickAnalysisFormats() {
  const context = activeQuickAnalysisContext();
  if (!context || !ensureActiveSheetFormatCellsAllowed()) return;
  const { sheet, range } = context;
  applyQuickAnalysisCellStyle(range, "クイック分析 書式のクリア", quickAnalysisClearFormatsForSheet, `${rangeToLabel(range)} のクイック分析書式をクリアしました。`);
}



function applyQuickAnalysisCellStyle(range, label, mutator, statusMessage) {
  const sheet = activeSheet();
  if (applyQuickAnalysisCellStyleToSheetGroup(range, label, mutator, statusMessage)) return;
  const before = captureCellsForHistory(sheet, range);
  const changed = mutator(sheet, range, state.activeSheetIndex);
  if (!changed) {
    showMessageBox("適用できるセルが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
    selectionAfterRanges: [{ ...range }],
    label,
  });
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(statusMessage);
}



function applyQuickAnalysisCellStyleToSheetGroup(range, label, mutator, statusMessage) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではセルの書式を変更できません。")) return true;
    const targetRange = materializedRangesForSheet(sheet, [range], sheetIndex)[0];
    if (targetRange) items.push({ sheetIndex, sheet, range: targetRange });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  let changed = 0;
  try {
    items.forEach((item) => {
      changed += mutator(item.sheet, item.range, item.sheetIndex) || 0;
      internCellsInRange(item.sheet, item.range);
    });
    if (!changed) {
      showMessageBox("適用できるセルが見つかりません。", { title: "Cht WebSheet" });
      return true;
    }
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
      selectionAfterRanges: [{ ...range }],
      label: `${label} ${sheetIndexes.length} シート`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${statusMessage}`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`クイック分析を適用できませんでした：${error.message || error}`);
  }
  return true;
}



function quickAnalysisApplyDataBars(targetSheet, targetRange) {
  const numericCells = quickAnalysisNumericCells(targetSheet, targetRange);
  if (!numericCells.length) return 0;
  const values = numericCells.map((cell) => cell.value);
  const min = Math.min(0, ...values);
  const max = Math.max(...values);
  const span = max === min ? 1 : max - min;
  numericCells.forEach(({ row, col, value }) => {
    const cell = ensureCellModel(targetSheet, row, col);
    cell.quickAnalysisBar = {
      ratio: Math.max(0, Math.min(1, (value - min) / span)),
      color: "#5b9bd5",
    };
    delete cell.quickAnalysisIcon;
    cell.quickAnalysisStyle = { ...(cell.quickAnalysisStyle || {}), dataBar: true };
  });
  return numericCells.length;
}



function quickAnalysisApplyColorScale(targetSheet, targetRange) {
  const numericCells = quickAnalysisNumericCells(targetSheet, targetRange);
  if (!numericCells.length) return 0;
  const values = numericCells.map((cell) => cell.value);
  const min = Math.min(...values);
  const max = Math.max(...values);
  const span = max === min ? 1 : max - min;
  numericCells.forEach(({ row, col, value }) => {
    const ratio = Math.max(0, Math.min(1, (value - min) / span));
    const cell = ensureCellModel(targetSheet, row, col);
    cell.css.backgroundColor = quickAnalysisScaleColor(ratio);
    cell.hasFill = true;
    delete cell.quickAnalysisBar;
    cell.quickAnalysisStyle = { ...(cell.quickAnalysisStyle || {}), background: true };
  });
  return numericCells.length;
}



function quickAnalysisApplyIconSet(targetSheet, targetRange) {
  const numericCells = quickAnalysisNumericCells(targetSheet, targetRange);
  if (!numericCells.length) return 0;
  const values = numericCells.map((cell) => cell.value);
  const min = Math.min(...values);
  const max = Math.max(...values);
  const span = max === min ? 1 : max - min;
  numericCells.forEach(({ row, col, value }) => {
    const ratio = (value - min) / span;
    const cell = ensureCellModel(targetSheet, row, col);
    cell.quickAnalysisIcon = ratio >= 0.67
      ? { shape: "up", color: "#00b050" }
      : ratio >= 0.33
        ? { shape: "same", color: "#ffc000" }
        : { shape: "down", color: "#c00000" };
    delete cell.quickAnalysisBar;
    cell.quickAnalysisStyle = { ...(cell.quickAnalysisStyle || {}), iconSet: true };
  });
  return numericCells.length;
}



function quickAnalysisApplyGreaterThanAverage(targetSheet, targetRange) {
  const numericCells = quickAnalysisNumericCells(targetSheet, targetRange);
  if (!numericCells.length) return 0;
  const average = numericCells.reduce((sum, cell) => sum + cell.value, 0) / numericCells.length;
  let changed = 0;
  numericCells.forEach(({ row, col, value }) => {
    if (value <= average) return;
    const cell = ensureCellModel(targetSheet, row, col);
    cell.css.backgroundColor = "#c6efce";
    cell.css.color = "#006100";
    cell.hasFill = true;
    cell.hasFontColor = true;
    cell.quickAnalysisStyle = { ...(cell.quickAnalysisStyle || {}), background: true, fontColor: true };
    changed += 1;
  });
  return changed;
}



function quickAnalysisApplyDuplicateValues(targetSheet, targetRange) {
  const sheetId = getSheetId(targetSheet.name);
  const buckets = new Map();
  forEachCellInRange(targetRange, (row, col) => {
    const text = String(getDisplayForCell(targetSheet, sheetId, row, col) ?? "").trim();
    if (!text) return;
    const key = text.toLocaleLowerCase("ja-JP");
    const list = buckets.get(key) || [];
    list.push({ row, col });
    buckets.set(key, list);
  });
  const duplicates = Array.from(buckets.values()).filter((list) => list.length > 1).flat();
  duplicates.forEach(({ row, col }) => {
    const cell = ensureCellModel(targetSheet, row, col);
    cell.css.backgroundColor = "#ffc7ce";
    cell.css.color = "#9c0006";
    cell.hasFill = true;
    cell.hasFontColor = true;
    cell.quickAnalysisStyle = { ...(cell.quickAnalysisStyle || {}), background: true, fontColor: true };
  });
  return duplicates.length;
}



function quickAnalysisClearFormatsForSheet(targetSheet, targetRange) {
  let changed = 0;
  forEachCellInRange(targetRange, (row, col) => {
    const cell = targetSheet.cells[cellKey(row, col)];
    if (!cell?.quickAnalysisStyle && !cell?.quickAnalysisBar && !cell?.quickAnalysisIcon) return;
    cell.css = mutableCssObject(cell.css);
    if (cell.quickAnalysisStyle?.background) {
      delete cell.css.backgroundColor;
      cell.hasFill = false;
    }
    if (cell.quickAnalysisStyle?.fontColor) {
      delete cell.css.color;
      cell.hasFontColor = false;
    }
    delete cell.quickAnalysisBar;
    delete cell.quickAnalysisIcon;
    delete cell.quickAnalysisStyle;
    storeCellModelForSheet(targetSheet, cellKey(row, col), cell);
    changed += 1;
  });
  return changed;
}



function quickAnalysisScaleColor(ratio) {
  if (ratio <= 0.5) return interpolateHexColor("#f8696b", "#ffeb84", ratio / 0.5);
  return interpolateHexColor("#ffeb84", "#63be7b", (ratio - 0.5) / 0.5);
}

export {
  activateSheetTabForContextMenu,
  activeQuickAnalysisContext,
  addPageBreakContextItems,
  adjustActiveSheetZoom,
  applyQuickAnalysisCellStyle,
  applyQuickAnalysisCellStyleToSheetGroup,
  applyQuickAnalysisColorScale,
  applyQuickAnalysisDataBars,
  applyQuickAnalysisDuplicateValues,
  applyQuickAnalysisGreaterThanAverage,
  applyQuickAnalysisIconSet,
  applySheetProtectionConfirmDialog,
  applySheetProtectionDialog,
  applySheetUnprotectDialog,
  applyTranslationPaneResult,
  applyTranslationPaneResultToSheetGroup,
  applyZoomDialog,
  canInsertCopiedHeaderCells,
  captureZoomForHistory,
  cellContextMenuItems,
  cellLockedForProtection,
  clampZoom,
  clearQuickAnalysisFormats,
  clearSheetSelectionState,
  cloneSheetModel,
  closeQuickAnalysisPopup,
  closeQuickAnalysisPopupOnEscape,
  closeQuickAnalysisPopupOnOutsidePointer,
  closeSheetNavigatorMenu,
  closeSheetNavigatorMenuOnOutsidePointer,
  closeSheetTabContextMenu,
  closeSheetTabContextMenuOnEscape,
  closeSheetTabContextMenuOnOutsidePointer,
  closeTranslationPane,
  closeTranslationPaneOnEscape,
  closeZoomDialog,
  closeZoomDialogOnEscape,
  closeZoomDialogOnOutsidePointer,
  collectSheetProtectionOptions,
  contextMenuHasPhonetics,
  contextMenuHasValidationDropdown,
  contextMenuTargetRanges,
  createBlankSheet,
  createSpecialSheet,
  currentTranslationContext,
  defaultSheetProtectionOptions,
  defaultTranslationTargetLanguage,
  detectLanguageForTranslationItems,
  detectTranslationLanguage,
  displayScaleForZoom,
  ensureActiveSheetFormatCellsAllowed,
  ensureEditableRangeForProtection,
  ensureEditableRangesForProtection,
  ensureHeaderFormatAllowed,
  ensureHeaderFormatAllowedForSheet,
  ensureSelectableRangeForProtection,
  ensureSheetProtectionPermission,
  ensureStructureAllowedForProtection,
  ensureStructureAllowedForSheet,
  handleQuickAnalysisAction,
  handleQuickAnalysisTabClick,
  handleSheetNavigatorMenuChoice,
  handleSheetTabColorChoice,
  handleSheetTabContextMenuAction,
  handleSheetZoomWheel,
  handleTranslationPaneAction,
  handleTranslationPaneFieldChange,
  handleZoomDialogAction,
  handleZoomDialogKeydown,
  handleZoomDialogOptionChange,
  headerContextMenuItems,
  headerIndexInActiveSelection,
  imageContextMenuItems,
  imageUsedRangeForZoom,
  isSheetProtected,
  isWorkbookWheelTarget,
  languageLabel,
  moveSheetToIndex,
  nearestVisibleSheetIndex,
  normalizeZoomUsedRange,
  objectFormatMenuLabel,
  objectFormatPaneTitle,
  openCellContextMenu,
  openColumnHeaderContextMenu,
  openHeaderContextMenu,
  openImageContextMenu,
  openQuickAnalysisPopup,
  openRowHeaderContextMenu,
  openSheetNavigatorMenu,
  openSheetProtectionDialog,
  openSheetTabContextMenu,
  openSheetTabContextMenuByKeyboard,
  openSheetUnprotectDialog,
  openTranslationPane,
  openZoomDialog,
  positionQuickAnalysisPopup,
  positionSheetTabContextMenu,
  positionZoomDialog,
  protectSheetWithOptions,
  quickAnalysisApplyColorScale,
  quickAnalysisApplyDataBars,
  quickAnalysisApplyDuplicateValues,
  quickAnalysisApplyGreaterThanAverage,
  quickAnalysisApplyIconSet,
  quickAnalysisClearFormatsForSheet,
  quickAnalysisContext,
  quickAnalysisNumericCells,
  quickAnalysisNumericColumns,
  quickAnalysisPopupHtml,
  quickAnalysisScaleColor,
  rangeCellCountForTranslation,
  rangeHasLockedCellsForProtection,
  rangeSelectionTypesForProtection,
  rangesHaveLockedCellsForProtection,
  rebuildWorkbookAfterSheetChange,
  refreshWorkbookAfterSheetMetadataChange,
  renderCellContextMenu,
  renderHeaderContextMenu,
  renderImageContextMenu,
  renderShapeVertexContextMenu,
  renderSheetNavigatorMenu,
  renderSheetProtectionConfirmDialog,
  renderSheetTabContextMenu,
  renderTranslationPane,
  restoreZoomFromHistory,
  runQuickAnalysisAction,
  runSheetTabMenuAction,
  scaleSheetDisplay,
  scrollSheetTabs,
  selectAllVisibleSheets,
  selectCustomZoomOption,
  setActiveSheetView,
  setActiveSheetZoom,
  setSheetTabColor,
  setSheetTabColorForSheets,
  shapeVertexContextMenuItems,
  sheetBaseZoomScale,
  sheetProtectionAllows,
  sheetProtectionDialogHtml,
  sheetTabColorSubmenuHtml,
  sheetTabContextMenuItems,
  sheetUsedRangeForZoom,
  showProtectedSheetMessage,
  swapTranslationPaneLanguages,
  toggleSheetProtectionDialog,
  translateTextOffline,
  translateTextWithDictionary,
  translationContextFromSelection,
  translationItemsForRanges,
  translationLanguageOptions,
  translationPreviewRowHtml,
  translationSourceTextForCell,
  translationTargetPlan,
  uniqueCopiedSheetName,
  uniqueCopiedSheetNameFromSet,
  uniqueSheetName,
  unprotectSheet,
  updateOpenZoomDialogValue,
  updateSheetFooterUi,
  validateSheetName,
  zoomDialogAccessKeyAttr,
  zoomSelectionFitRange,
  zoomSelectionNormalizeRange,
  zoomSelectionToFit,
  zoomSelectionUnionRanges,
};
