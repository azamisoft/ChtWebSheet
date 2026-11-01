// Feature module; shared bindings are imported explicitly.
import {
  $fileSideMenu,
  $formulaInput,
  $horizontalScrollbar,
  $ribbonDisplayButton,
  $ribbonDisplayMenu,
  $ribbonPanels,
  $ribbonTabs,
  $selectedCellLabel,
  $sheetHost,
  COLUMN_HEADER_HEIGHT,
  CWS_WEB_APP_URL,
  DATA_OPERATION_CELL_LIMIT,
  DEFAULT_SHAPE_OUTLINE_WIDTH,
  DEFAULT_SHAPE_TYPE,
  FALLBACK_VISIBLE_COLS,
  FALLBACK_VISIBLE_ROWS,
  LINK_DIALOG_KINDS,
  RIBBON_LAYOUTS,
  RIBBON_LAYOUT_STORAGE_KEY,
  RIBBON_PANEL_FOCUSABLE_SELECTOR,
  RIBBON_VISIBILITY_MODES,
  RIBBON_VISIBILITY_STORAGE_KEY,
  ROW_HEADER_DIGIT_WIDTH,
  ROW_HEADER_MAX_WIDTH,
  ROW_HEADER_MIN_WIDTH,
  ROW_HEADER_PADDING_WIDTH,
  SCROLL_EXTEND_COLS,
  SCROLL_EXTEND_ROWS,
  SCROLL_EXTEND_THRESHOLD,
  SCROLL_TRIM_BUFFER_PX,
  SCROLL_TRIM_IDLE_MS,
  VIRTUAL_COL_OVERSCAN,
  VIRTUAL_ROW_OVERSCAN,
  WORKBOOK_OPTIONAL_RIBBON_COMMANDS,
  sheetMinimumExtentCacheBySheet,
  state,
} from "../../app-core.js";
import {
  openDataAnalysisDialog,
  openWhatIfMenu,
} from "../data/analysis-and-scenarios.js";
import {
  openDataModelDialog,
  openExistingConnectionsDialog,
  openQueriesConnectionsPane,
} from "../data/connections-and-queries.js";
import {
  flashFillSelection,
  openAdvancedFilterDialog,
  openTextToColumnsDialog,
} from "../data/filter-and-text-tools.js";
import {
  openForecastSheetDialog,
} from "../data/forecast-and-statistics.js";
import {
  chooseDelimitedTextForImport,
} from "../data/query-editor.js";
import {
  closeDataCommandMenu,
  openDataGetMenu,
  openDataRefreshMenu,
  openGetDataFromTableRangeDialog,
  openOutlineCommandMenu,
  openRecentSourcesMenu,
  openWebImportDialog,
} from "../data/query-transforms.js";
import {
  clearAutoFilterCriteria,
  currentRegionFromPoint,
  inferHeaderRow,
  openSortDialog,
  reapplyAutoFilter,
  reapplyAutoFilterForSheetGroup,
  refreshWorkbookData,
  sortSelectionByActiveColumn,
  toggleAutoFilterForSelection,
} from "../data/sort-and-autofilter.js";
import {
  openConsolidateDialog,
  openDataValidationMenu,
  openRelationshipsDialog,
  openSubtotalDialog,
  removeDuplicatesFromSelection,
} from "../data/validation-and-consolidation.js";
import {
  beginShapeDrawingMode,
  cancelShapeDrawing,
  choosePictureForInsertion,
  closeDataDialog,
  insertChartFromSelection,
  insertPivotSummaryFromSelection,
  insertSparklinesFromSelection,
  insertTextBoxObject,
  insertWordArtObject,
  openInsertObjectDialog,
  openSignatureLineDialog,
  recommendedChartTypeForSelection,
  toggleOutlineDetailFromSelection,
} from "../drawing/charts-and-insertion.js";
import {
  insert3dModelPlaceholder,
  insertEquationObject,
  insertScreenshotObject,
  insertSlicerObject,
  insertSmartArtObject,
  insertTimelineObject,
  openHeaderFooterDialog,
  openInsertIconDialog,
  openOfficeAddinsDialog,
} from "../drawing/shape-rendering.js";
import {
  applySelectedShapeStyle,
  groupSelectedShapes,
  reorderSelectedShape,
  resetSelectedShapeSize,
  togglePictureCropMode,
  ungroupSelectedShapes,
} from "../drawing/transforms-and-connectors.js";
import {
  groupedShapeObjectTargets,
  openShapeTextEditor,
} from "../editing/autofill.js";
import {
  applyUnderlinePreset,
  openLinkDialogForSelection,
  openRibbonInkEraserMenu,
  openShapeSelectionPane,
  openSymbolDialog,
  setRibbonUnderlinePreset,
  syncShapeFormatPane,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  copySelectedImages,
  copySelectedRange,
  cutSelectedImages,
  cutSelectedRange,
  executeWorkbookClipboardShortcut,
  isElementImeCompositionEvent,
  isTextEditingTarget,
  pasteFromRibbonButton,
  selectedObjectSummary,
} from "../editing/keyboard-and-paste.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  affectedSheetIndexesForGroupedRanges,
  cssEscape,
  getCellRawInput,
  isCellInRange,
  isSingleCellRange,
  materializeSelectionRange,
  materializeSelectionRangeForSheet,
  normalizeRange,
  rangeToLabel,
  rangesToLabel,
  refreshChartsForGroupedRanges,
  restoreGroupedSheetsAfterFailure,
  selectionRangeFromPoints,
  syncShapeFormatControls,
  updateSelectionUi,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  coordinateToPixels,
  imageBoxToPixels,
  rawActiveSelectionRanges,
  selectCellByKeyboard,
  selectedImageObjectById,
  selectedImageObjects,
  selectionToolbarPoint,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneDefinedNames,
  internCellsInRange,
  internCellsInRanges,
  pushHistory,
  restoreSheetFromHistory,
} from "../editing/structure-and-history.js";
import {
  formulaInsertFunctionDefinitionByName,
} from "../formula/autocomplete.js";
import {
  openDefineNameForSelection,
  openFormulaAutoSumMenu,
  openFormulaDefineNameMenu,
  openFormulaFunctionMenu,
  openInsertFunctionDialog,
  openNameManagerDialog,
  openPasteNameDialog,
  openUseInFormulaMenu,
  recordFormulaFunctionUse,
} from "../formula/dialogs-and-names.js";
import {
  adjustSelectionDecimalPlaces,
  applyCssToSelection,
  applyNumberFormatToSelection,
  clearFormatPainter,
  ensureCellModel,
  isFormatPainterRepeatTrigger,
  mergeSelection,
  renderSheet,
  toggleFormatPainter,
  toggleMergeCenterSelection,
  toggleStyleForSelection,
  toggleWrapTextForSelection,
  unmergeSelection,
} from "../rendering/grid-and-canvas.js";
import {
  closeCellContextMenu,
  linkDialogDocumentTargetsHtml,
  openCellStructureDialog,
  openFormatCellsDialog,
} from "./annotations-and-cell-format.js";
import {
  accessKeyLabelHtml,
  cellKey,
  columnName,
  columnWidthToPixels,
  decodeCell,
  escapeAttr,
  escapeHtml,
  openHelpDialog,
  openVersionDialog,
  rowHeightToPixels,
  setStatus,
  showMessageBox,
} from "./help-and-utilities.js";
import {
  autoSumOperationsForRanges,
  beginInkDrawingMode,
  beginInkEraserMode,
  cancelInkDrawingMode,
  clearInkObjects,
  closeHomeCommandMenu,
  commandMenuEscapeBlockedByDialog,
  consumeCommandMenuKey,
  convertSelectedInkToFreeformShape,
  enhanceCommandMenuAccessKeys,
  nextCommandMenuItemIndexByArrow,
  openFindSelectMenu,
  openHomeAutoSumMenu,
  openHomeClearMenu,
  openHomeFillMenu,
  openHomeSortFilterMenu,
  parsePageSetupRangeLabel,
  setInkColor,
  setInkWidth,
  syncDrawRibbonButtons,
} from "./printing-and-home-tools.js";
import {
  closeQuickAnalysisPopup,
  closeSheetTabContextMenu,
  contextMenuTargetRanges,
  ensureEditableRangesForProtection,
  ensureSheetProtectionPermission,
  openSheetTabContextMenuByKeyboard,
  openZoomDialog,
  renderCellContextMenu,
  renderHeaderContextMenu,
  renderImageContextMenu,
  setActiveSheetView,
  setActiveSheetZoom,
  zoomSelectionToFit,
} from "./sheet-protection-and-translation.js";
import {
  calculateActiveSheetNow,
  calculateWorkbookNow,
  createSheetView,
  exitActiveSheetView,
  keepActiveSheetView,
  openCalculationOptionsMenu,
  openCreateNamesDialog,
  openEvaluateFormulaDialog,
  openFormulaErrorCheckingMenu,
  openFormulaRemoveArrowsMenu,
  openFormulaWatchWindow,
  openSheetViewOptionsMenu,
  openViewMacrosMenu,
  refreshActiveSheetViewSnapshot,
  syncViewRibbonButtons,
  toggleActiveSheetGridlines,
  toggleActiveSheetHeadings,
  toggleActiveSheetObjectsVisibility,
  toggleActiveSheetOutlineSymbols,
  toggleActiveSheetRightToLeft,
  toggleActiveSheetRuler,
  toggleActiveSheetWhiteSpace,
  toggleActiveSheetZeroValues,
  toggleFormulaBarVisibility,
  toggleFreezePanes,
  toggleFullScreenView,
  toggleShowFormulasFromRibbon,
  toggleSplitPane,
  toggleWorkbookWindowShowOption,
  traceFormulaReferences,
} from "./views-and-formula-tools.js";
import {
  addBlankSheet,
  buildFormulaEngine,
  fontFamilyForExcel,
  isSheetVisible,
  normalizeDefinedNameLookup,
  renderWorkbook,
  setActiveSheet,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  openBlankWorkbook,
  workbookHasUnsavedChanges,
} from "../workbook/import-model.js";
import {
  parseTableStyleInfoXml,
} from "../workbook/ooxml-drawings.js";
import {
  activeSheet,
  getSheetId,
  heightBetweenRows,
  openSaveAsDialog,
  quickSaveWorkbook,
  syncHorizontalScrollbarFromHost,
  widthBetweenColumns,
} from "../workbook/saving-and-export.js";
import {
  openWorkbookFilePicker,
  showUnsavedWorkbookReplacementMessage,
} from "../workbook/startup-and-loading.js";
import {
  validationDefinedNameFromFormulaText,
  validationReferenceFromFormula,
} from "../workbook/tables-and-validation.js";
import {
  openPageLayoutBreaksMenu,
  openPageLayoutColorMenu,
  openPageLayoutEffectsMenu,
  openPageLayoutFontMenu,
  openPageLayoutMarginsMenu,
  openPageLayoutOrientationMenu,
  openPageLayoutPrintAreaMenu,
  openPageLayoutSizeMenu,
  openPageLayoutThemeMenu,
  openPageSetupDialog,
  openPrintDialog,
  setPageBackgroundFromSelection,
  setPrintTitlesFromSelection,
  syncPageLayoutRibbonControls,
} from "../workbook/vba-and-printing.js";
import {
  closeViewCommandMenu,
  createWorkbookWindowView,
  hideWorkbookWindow,
  openArrangeAllDialog,
  openCustomViewsDialog,
  openMacrosDialog,
  openSwitchWindowsMenu,
  openUnhideWorkbookWindowDialog,
  openViewFreezeMenu,
  resetWindowPosition,
  restoreWorksheetFocusAfterCommandPopup,
  showCustomViewsUnavailableIfNeeded,
  toggleSynchronousScrolling,
  toggleViewSideBySide,
} from "../workbook/views-and-macros.js";


function objectLinkDialogHtml(image, count, kind = "web", values = {}) {
  const title = image?.hyperlink ? "ハイパーリンクの編集" : "ハイパーリンクの挿入";
  const display = count === 1 ? image?.name || image?.shape?.label || "図形" : `${count} 個のオブジェクト`;
  const url = image?.hyperlink || "";
  const meta = linkKindMeta(kind);
  const documentSheetIndex = normalizedDocumentSheetIndex(values.documentSheetIndex);
  const buttons = LINK_DIALOG_KINDS.map(
    (item) =>
      `<button class="${item.kind === kind ? "is-active" : ""}" type="button" data-object-link-kind="${item.kind}" aria-pressed="${item.kind === kind ? "true" : "false"}">${escapeHtml(item.label)}</button>`,
  ).join("");
  return `
    <div id="objectLinkDialog" class="excel-mini-dialog link-dialog object-link-dialog" role="dialog" aria-modal="true" aria-labelledby="objectLinkDialogTitle">
      <div class="excel-mini-titlebar" id="objectLinkDialogTitle">${title}</div>
      <div class="excel-mini-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("表示文字列(T):")}</span>
          <input type="text" data-object-link-field="text" value="${escapeAttr(display)}"${count === 1 ? "" : " disabled"} />
        </label>
        <div class="link-dialog-content">
          <div class="link-dialog-kind" role="group" aria-label="リンク先">
            ${buttons}
          </div>
          <div class="link-dialog-main">
            <div class="link-dialog-address-label">${escapeHtml(display)}</div>
            <label class="excel-dialog-field"><span class="excel-dialog-caption" data-object-link-address-caption>${accessKeyLabelHtml(meta.fieldLabel)}</span>
              <input type="text" data-object-link-field="url" value="${escapeAttr(values[kind] ?? "")}" placeholder="${escapeAttr(meta.placeholder)}" />
            </label>
            <div class="link-dialog-document-panel" data-object-link-document-panel ${kind === "document" ? "" : "hidden"}>
              <div class="excel-dialog-caption">${accessKeyLabelHtml("またはドキュメント内の場所を選択してください(C):")}</div>
              <div class="link-dialog-document-tree" role="tree" aria-label="ドキュメント内の場所">
                ${linkDialogDocumentTargetsHtml(documentSheetIndex, values.documentName, {
                  sheetIndexAttribute: "data-object-link-document-sheet-index",
                  nameTargetAttribute: "data-object-link-document-name-target",
                })}
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="excel-mini-actions">
        ${url ? '<button class="excel-mini-button" type="button" data-object-link-dialog-action="remove">リンクの解除</button>' : ""}
        <button class="excel-mini-button is-primary" type="button" data-object-link-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-object-link-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
}



function handleObjectLinkKindClick(event) {
  const button = event.currentTarget;
  const dialog = button.closest("#objectLinkDialog");
  if (!dialog || !state.objectLinkDialog) return;
  event.preventDefault();
  const nextKind = button.dataset.objectLinkKind;
  if (!LINK_DIALOG_KINDS.some((item) => item.kind === nextKind)) return;
  const input = dialog.querySelector('[data-object-link-field="url"]');
  const previousKind = state.objectLinkDialog.kind || "web";
  state.objectLinkDialog.values = state.objectLinkDialog.values || {};
  state.objectLinkDialog.values[previousKind] = input?.value || "";
  state.objectLinkDialog.kind = nextKind;
  if (state.objectLinkDialog.values[nextKind] == null) {
    state.objectLinkDialog.values[nextKind] = defaultLinkDialogValue(nextKind, state.objectLinkDialog.point);
  }
  const meta = linkKindMeta(nextKind);
  dialog.querySelectorAll("[data-object-link-kind]").forEach((kindButton) => {
    const active = kindButton.dataset.objectLinkKind === nextKind;
    kindButton.classList.toggle("is-active", active);
    kindButton.setAttribute("aria-pressed", active ? "true" : "false");
  });
  const documentPanel = dialog.querySelector("[data-object-link-document-panel]");
  if (documentPanel) documentPanel.hidden = nextKind !== "document";
  const caption = dialog.querySelector("[data-object-link-address-caption]");
  if (caption) caption.innerHTML = accessKeyLabelHtml(meta.fieldLabel);
  if (input) {
    input.placeholder = meta.placeholder;
    input.value = state.objectLinkDialog.values[nextKind] || "";
    input.focus({ preventScroll: true });
    input.select?.();
  }
}



function handleObjectLinkDocumentSheetClick(event) {
  const button = event.currentTarget;
  const dialog = button.closest("#objectLinkDialog");
  if (!dialog || !state.objectLinkDialog) return;
  event.preventDefault();
  const index = Number(button.dataset.objectLinkDocumentSheetIndex);
  if (!Number.isInteger(index) || !state.model?.sheets[index] || !isSheetVisible(state.model.sheets[index])) return;
  state.objectLinkDialog.values = state.objectLinkDialog.values || {};
  state.objectLinkDialog.values.documentSheetIndex = index;
  state.objectLinkDialog.values.documentName = "";
  dialog.querySelectorAll("[data-object-link-document-sheet-index]").forEach((sheetButton) => {
    const active = Number(sheetButton.dataset.objectLinkDocumentSheetIndex) === index;
    sheetButton.classList.toggle("is-active", active);
    sheetButton.setAttribute("aria-selected", active ? "true" : "false");
  });
  dialog.querySelectorAll("[data-object-link-document-name-target]").forEach((nameButton) => {
    nameButton.classList.remove("is-active");
    nameButton.setAttribute("aria-selected", "false");
  });
  const input = dialog.querySelector('[data-object-link-field="url"]');
  if (input && !input.value.trim()) {
    input.value = defaultLinkDialogValue("document", state.objectLinkDialog.point);
  }
  input?.focus({ preventScroll: true });
}



function handleObjectLinkDocumentNameClick(event) {
  const button = event.currentTarget;
  const dialog = button.closest("#objectLinkDialog");
  if (!dialog || !state.objectLinkDialog) return;
  event.preventDefault();
  const target = button.dataset.objectLinkDocumentNameTarget || "";
  if (!target) return;
  state.objectLinkDialog.values = state.objectLinkDialog.values || {};
  state.objectLinkDialog.values.documentName = target;
  const reference = internalHyperlinkReferenceFromDefinedName(target, state.objectLinkDialog.values.documentSheetIndex);
  if (reference?.sheetName) {
    const index = state.model.sheets.findIndex((sheet) => sheet.name === reference.sheetName && isSheetVisible(sheet));
    if (index >= 0) state.objectLinkDialog.values.documentSheetIndex = index;
  }
  dialog.querySelectorAll("[data-object-link-document-name-target]").forEach((nameButton) => {
    const active = nameButton === button;
    nameButton.classList.toggle("is-active", active);
    nameButton.setAttribute("aria-selected", active ? "true" : "false");
  });
  dialog.querySelectorAll("[data-object-link-document-sheet-index]").forEach((sheetButton) => {
    const active = Number(sheetButton.dataset.objectLinkDocumentSheetIndex) === state.objectLinkDialog.values.documentSheetIndex;
    sheetButton.classList.toggle("is-active", active);
    sheetButton.setAttribute("aria-selected", active ? "true" : "false");
  });
  const input = dialog.querySelector('[data-object-link-field="url"]');
  if (input) input.value = target;
  input?.focus({ preventScroll: true });
}



function handleObjectLinkDialogAction(event) {
  const action = event.currentTarget.dataset.objectLinkDialogAction;
  if (action === "cancel") {
    closeObjectLinkDialog();
    return;
  }
  if (action === "remove") {
    if (removeObjectLinkFromDialog()) closeObjectLinkDialog();
    return;
  }
  if (applyObjectLinkDialog()) closeObjectLinkDialog();
}



function applyObjectLinkDialog() {
  const dialog = document.querySelector("#objectLinkDialog");
  const ids = state.objectLinkDialog?.imageIds || [];
  if (!dialog || !ids.length) return false;
  const rawUrl = dialog.querySelector('[data-object-link-field="url"]')?.value.trim() || "";
  const kind = state.objectLinkDialog?.kind || "web";
  const url = normalizeHyperlinkForKind(rawUrl, kind, state.objectLinkDialog?.values?.documentSheetIndex);
  if (!url) {
    showMessageBox(linkKindEmptyMessage(kind), { title: "Cht WebSheet" });
    return false;
  }
  const text = dialog.querySelector('[data-object-link-field="text"]')?.value.trim() || "";
  applyHyperlinkToObjects(ids, { hyperlink: url, text });
  setStatus(`${ids.length} 個のオブジェクトにリンクを設定しました。`);
  return true;
}



function removeObjectLinkFromDialog() {
  const ids = state.objectLinkDialog?.imageIds || [];
  if (!ids.length) return false;
  applyHyperlinkToObjects(ids, { remove: true });
  setStatus(`${ids.length} 個のオブジェクトのリンクを解除しました。`);
  return true;
}



function applyHyperlinkToObjects(ids, options = {}) {
  const sheet = activeSheet();
  const idSet = new Set(ids || []);
  const targets = (sheet.images || []).filter((image) => idSet.has(image.id));
  if (!targets.length) return false;
  if (applyHyperlinkToObjectsForSheetGroup(targets, options)) return true;
  const sheetBefore = captureSheetForHistory(sheet);
  targets.forEach((image, index) => {
    if (options.remove) {
      image.hyperlink = null;
    } else {
      image.hyperlink = options.hyperlink;
      if (index === 0 && options.text && targets.length === 1) image.name = options.text;
    }
  });
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${selectedObjectSummary(targets)} リンク`,
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  return true;
}



function applyHyperlinkToObjectsForSheetGroup(targets, options = {}) {
  const items = groupedShapeObjectTargets(targets);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      item.images.forEach((image, index) => {
        if (options.remove) {
          image.hyperlink = null;
        } else {
          image.hyperlink = options.hyperlink;
          if (index === 0 && options.text && item.images.length === 1) image.name = options.text;
        }
      });
    });
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${items.length} シート${selectedObjectSummary(targets)} リンク`,
    });
    renderSheet();
    updateSelectionUi();
    syncShapeFormatControls();
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`オブジェクトのリンクを設定できませんでした：${error.message || error}`);
  }
  return true;
}



function renameSelectedObjects(name) {
  const nextName = String(name || "").trim();
  if (!nextName) {
    syncShapeFormatPane();
    return false;
  }
  const objects = selectedImageObjects();
  if (objects.length !== 1) {
    syncShapeFormatPane();
    return false;
  }
  const image = objects[0];
  if (image.name === nextName) return false;
  if (renameObjectForSheetGroup(image, nextName)) return true;
  const sheet = activeSheet();
  const sheetBefore = captureSheetForHistory(sheet);
  image.name = nextName;
  if (image.shape) image.shape = { ...image.shape, label: nextName };
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${nextName} 名前の変更`,
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`オブジェクト名を ${nextName} に変更しました。`);
  return true;
}



function renameObjectForSheetGroup(image, nextName) {
  const items = groupedShapeObjectTargets([image]);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      const target = item.images[0];
      if (!target) return;
      target.name = nextName;
      if (target.shape) target.shape = { ...target.shape, label: nextName };
    });
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${items.length} シート ${nextName} 名前の変更`,
    });
    renderSheet();
    updateSelectionUi();
    syncShapeFormatControls();
    setStatus(`${items.length} 個のシートのオブジェクト名を ${nextName} に変更しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`オブジェクト名を変更できませんでした：${error.message || error}`);
  }
  return true;
}



function openSelectedObjectHyperlink() {
  const image = selectedImageObjects()[0];
  if (!image?.hyperlink) {
    setStatus("選択したオブジェクトにリンクがありません。");
    return;
  }
  if (openInternalHyperlink(image.hyperlink)) {
    setStatus(`リンク先へ移動しました：${image.hyperlink}`);
    return;
  }
  window.open(image.hyperlink, "_blank", "noopener,noreferrer");
  setStatus(`リンクを開きました：${image.hyperlink}`);
}



function closeObjectLinkDialog(options = {}) {
  document.querySelector("#objectLinkDialogOverlay")?.remove();
  const previousFocus = state.objectLinkDialog?.previousFocus;
  state.objectLinkDialog = null;
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) {
    previousFocus.focus?.({ preventScroll: true });
  }
}



function closeObjectLinkDialogOnEscape(event) {
  if (event.key !== "Escape" || !state.objectLinkDialog) return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  closeObjectLinkDialog();
}



function contextMenuActionPoint(row, col) {
  const ranges = contextMenuTargetRanges(row, col);
  const toolbarPoint = selectionToolbarPoint();
  if (toolbarPoint && ranges.some((range) => isCellInRange(range, toolbarPoint.row, toolbarPoint.col))) {
    return { sheetIndex: state.activeSheetIndex, row: toolbarPoint.row, col: toolbarPoint.col };
  }
  const range = ranges[0] || (row && col ? normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col }) : null);
  return range ? { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left } : null;
}



function suggestedDefinedNameForRange(sheet, range) {
  const cell = sheet.cells[cellKey(range.top, range.left)];
  const source = String(cell?.display ?? cell?.raw ?? "").trim();
  return sanitizeDefinedName(source) || `${columnName(range.left)}${range.top}`;
}



function sanitizeDefinedName(value) {
  const text = String(value || "")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^A-Za-z0-9_.\\\u0080-\uffff]+/g, "_")
    .replace(/^_+|_+$/g, "");
  if (!text) return "";
  const candidate = /^[A-Za-z_\\\u0080-\uffff]/.test(text) ? text : `_${text}`;
  return isCellReference(candidate) ? `_${candidate}` : candidate.slice(0, 255);
}



function validateDefinedName(name) {
  if (!name) return { ok: false, message: "名前を入力してください。" };
  if (name.length > 255) return { ok: false, message: "名前は 255 文字以内で入力してください。" };
  if (isCellReference(name)) return { ok: false, message: "セル参照と同じ名前は使用できません。" };
  if (!/^[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*$/.test(name)) {
    return { ok: false, message: "名前には空白や記号を使用できません。" };
  }
  return { ok: true };
}



function definedNameDuplicateMessage(name) {
  return `名前 "${name}" は既に存在します。別の名前を入力してください。`;
}



function isCellReference(value) {
  return /^[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}$/.test(String(value || "").replace(/\$/g, ""));
}



function normalizeDefinedNameRefersTo(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  return text.startsWith("=") ? text : `=${text}`;
}



function upsertDefinedName(definedNames = [], definition) {
  const key = definedNameKey(definition);
  const next = cloneDefinedNames(definedNames);
  const index = next.findIndex((item) => definedNameKey(item) === key);
  if (index >= 0) next[index] = definition;
  else next.push(definition);
  return next;
}



function definedNameKey(definition) {
  const scope = definition.scope === "sheet" ? `sheet:${definition.sheetIndex}` : "workbook";
  return `${scope}:${normalizeDefinedNameLookup(definition.name)}`;
}



function definedNameHasConflict(definition, ignoreIndex = -1, definedNames = state.model?.definedNames || []) {
  const targetKey = definedNameKey(definition);
  return (definedNames || []).some((item, index) => index !== ignoreIndex && !item.hidden && definedNameKey(item) === targetKey);
}



function absoluteRangeReference(sheet, range) {
  const sheetName = quoteSheetNameForFormula(sheet?.name || activeSheet().name);
  const start = `$${columnName(range.left)}$${range.top}`;
  const end = `$${columnName(range.right)}$${range.bottom}`;
  return `${sheetName}!${start === end ? start : `${start}:${end}`}`;
}



function quoteSheetNameForFormula(name) {
  const text = String(name || "Sheet1");
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(text)) return text;
  return `'${text.replace(/'/g, "''")}'`;
}



function hyperlinkKind(hyperlink) {
  const text = String(hyperlink || "").trim();
  if (isInternalHyperlinkValue(text)) return "document";
  if (/^mailto:/i.test(text)) return "email";
  return "web";
}



function linkKindMeta(kind) {
  return LINK_DIALOG_KINDS.find((item) => item.kind === kind) || LINK_DIALOG_KINDS[0];
}



function linkDialogInitialValues(hyperlink, kind, point) {
  const documentSource = kind === "document" ? stripInternalHyperlinkPrefix(hyperlink) : "";
  const documentTarget = kind === "document" ? parseInternalHyperlinkTarget(documentSource) : null;
  const documentIsCellReference = internalHyperlinkTextIsCellReference(documentSource);
  const documentSheetIndex = documentTarget?.sheetName
    ? state.model.sheets.findIndex((sheet) => sheet.name === documentTarget.sheetName && isSheetVisible(sheet))
    : state.activeSheetIndex;
  return {
    web: kind === "web" ? String(hyperlink || "") : "",
    document: documentTarget && documentIsCellReference
      ? `${columnName(documentTarget.col)}${documentTarget.row}`
      : documentSource || defaultLinkDialogValue("document", point),
    documentName: documentSource && !documentIsCellReference ? documentSource : "",
    documentSheetIndex: normalizedDocumentSheetIndex(documentSheetIndex),
    email: kind === "email" ? stripMailtoPrefix(hyperlink) : "",
  };
}



function normalizedDocumentSheetIndex(index) {
  const numericIndex = Number(index);
  if (Number.isInteger(numericIndex) && state.model?.sheets[numericIndex] && isSheetVisible(state.model.sheets[numericIndex])) {
    return numericIndex;
  }
  return state.activeSheetIndex;
}



function defaultLinkDialogValue(kind, point) {
  if (kind === "document") return "A1";
  return "";
}



function stripInternalHyperlinkPrefix(hyperlink) {
  const text = String(hyperlink || "").replace(/^#/, "");
  try {
    return decodeURIComponent(text);
  } catch {
    return text;
  }
}



function stripMailtoPrefix(hyperlink) {
  return String(hyperlink || "")
    .replace(/^mailto:/i, "")
    .replace(/\?.*$/, "");
}



function normalizeHyperlinkForKind(value, kind, documentSheetIndex) {
  if (kind === "document") return normalizeInternalHyperlinkUrl(value, documentSheetIndex);
  if (kind === "email") return normalizeEmailHyperlinkUrl(value);
  return normalizeHyperlinkUrl(value);
}



function linkKindEmptyMessage(kind) {
  if (kind === "document") return "参照先を入力してください。";
  if (kind === "email") return "電子メール アドレスを入力してください。";
  return "アドレスを入力してください。";
}



function normalizeHyperlinkUrl(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  if (/^(https?|ftp|mailto|tel|file):/i.test(text) || text.startsWith("#")) return text;
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text)) return `mailto:${text}`;
  return `https://${text.replace(/^\/\//, "")}`;
}



function normalizeInternalHyperlinkUrl(value, documentSheetIndex) {
  const raw = String(value || "").trim();
  const text = stripInternalHyperlinkPrefix(raw);
  if (!text) return "";
  const definedNameTarget = normalizedInternalDefinedNameHyperlinkTarget(text, documentSheetIndex);
  if (definedNameTarget) return `#${definedNameTarget}`;
  if (raw.startsWith("#")) return raw;
  if (text.includes("!")) return `#${text}`;
  const sheet = state.model?.sheets[normalizedDocumentSheetIndex(documentSheetIndex)];
  return sheet ? `#${quoteSheetNameForFormula(sheet.name)}!${text}` : `#${text}`;
}



function normalizeEmailHyperlinkUrl(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  return /^mailto:/i.test(text) ? text : `mailto:${text}`;
}



function openInternalHyperlink(hyperlink) {
  const text = String(hyperlink || "").trim();
  if (!state.model || !isInternalHyperlinkValue(text)) return false;
  const target = parseInternalHyperlinkTarget(stripInternalHyperlinkPrefix(text));
  if (!target) return false;
  const sheetIndex = target.sheetName
    ? state.model.sheets.findIndex((sheet) => sheet.name === target.sheetName && isSheetVisible(sheet))
    : state.activeSheetIndex;
  if (sheetIndex < 0 || !state.model.sheets[sheetIndex]) return false;
  if (sheetIndex !== state.activeSheetIndex) setActiveSheet(sheetIndex);
  selectCellByKeyboard(target.row, target.col);
  return true;
}



function isInternalHyperlinkValue(hyperlink) {
  const text = String(hyperlink || "").trim();
  if (!text) return false;
  if (text.startsWith("#")) return true;
  if (/^[a-z][a-z0-9+.-]*:/i.test(text)) return false;
  return Boolean(parseInternalHyperlinkTarget(stripInternalHyperlinkPrefix(text)));
}



function parseInternalHyperlinkTarget(target) {
  const text = String(target || "").trim();
  if (!text) return null;
  const bangIndex = text.lastIndexOf("!");
  const sheetName = bangIndex >= 0 ? unquoteFormulaSheetName(text.slice(0, bangIndex)) : "";
  const address = (bangIndex >= 0 ? text.slice(bangIndex + 1) : text).split(":")[0];
  const cell = decodeCell(address);
  if (cell) return { sheetName, row: cell.row, col: cell.col };
  return internalHyperlinkReferenceFromDefinedName(text);
}



function internalHyperlinkTextIsCellReference(target) {
  const text = String(target || "").trim();
  if (!text) return false;
  const bangIndex = text.lastIndexOf("!");
  const address = (bangIndex >= 0 ? text.slice(bangIndex + 1) : text).split(":")[0];
  return Boolean(decodeCell(address));
}



function normalizedInternalDefinedNameHyperlinkTarget(target, documentSheetIndex) {
  if (internalHyperlinkTextIsCellReference(target)) return "";
  const sheet = state.model?.sheets?.[normalizedDocumentSheetIndex(documentSheetIndex)] || activeSheet();
  const definition = linkDefinedNameFromTarget(target, sheet);
  if (!definition) return "";
  const reference = definedNameLinkReference(definition, sheet);
  if (!reference?.sheet || !reference.range) return "";
  if (definition.scope === "sheet") {
    const scopeSheet = state.model?.sheets?.[definition.sheetIndex];
    return scopeSheet ? `${quoteSheetNameForFormula(scopeSheet.name)}!${definition.name}` : "";
  }
  return definition.name;
}



function internalHyperlinkReferenceFromDefinedName(target, documentSheetIndex = state.activeSheetIndex) {
  if (internalHyperlinkTextIsCellReference(target)) return null;
  const sheet = state.model?.sheets?.[normalizedDocumentSheetIndex(documentSheetIndex)] || activeSheet();
  const definition = linkDefinedNameFromTarget(target, sheet);
  if (!definition) return null;
  const reference = definedNameLinkReference(definition, sheet);
  if (!reference?.sheet || !reference.range) return null;
  return {
    sheetName: reference.sheet.name,
    row: reference.range.top,
    col: reference.range.left,
  };
}



function linkDefinedNameFromTarget(target, baseSheet = activeSheet()) {
  const text = String(target || "").trim();
  if (!text) return null;
  return validationDefinedNameFromFormulaText(text, baseSheet);
}



function definedNameLinkReference(definition, baseSheet = activeSheet()) {
  if (!definition?.refersTo) return null;
  const definitionSheet = definition.scope === "sheet"
    ? state.model?.sheets?.[definition.sheetIndex] || baseSheet
    : baseSheet;
  return validationReferenceFromFormula(definitionSheet, definition.refersTo, new Set());
}



function unquoteFormulaSheetName(name) {
  const text = String(name || "").trim();
  if (text.startsWith("'") && text.endsWith("'")) return text.slice(1, -1).replace(/''/g, "'");
  return text;
}



function scheduleViewportRender({ deferTrim = false } = {}) {
  if (deferTrim) scheduleIdleViewportTrim();
  if (state.viewportRenderFrame) return;
  state.viewportRenderFrame = requestAnimationFrame(() => {
    state.viewportRenderFrame = null;
    if (!state.model) return;
    if (activeSheetEditElement()) {
      state.deferredViewportRender = true;
      return;
    }
    const scrollAnchor = resizeActiveSheetForScroll({ allowTrim: !deferTrim });
    renderSheet(scrollAnchor?.renderScroll, { viewportOnly: true });
    restoreViewportScroll(scrollAnchor);
  });
}



function scheduleIdleViewportTrim() {
  if (state.viewportTrimTimer) window.clearTimeout(state.viewportTrimTimer);
  state.viewportTrimTimer = window.setTimeout(() => {
    state.viewportTrimTimer = null;
    if (!state.model || state.restoringViewportScroll) return;
    if (activeSheetEditElement()) {
      state.deferredViewportRender = true;
      return;
    }
    const scrollAnchor = resizeActiveSheetForScroll({ allowExtend: false, allowTrim: true });
    renderSheet(scrollAnchor?.renderScroll, { viewportOnly: true });
    restoreViewportScroll(scrollAnchor);
  }, SCROLL_TRIM_IDLE_MS);
}



function cancelViewportRender() {
  if (!state.viewportRenderFrame) return;
  cancelAnimationFrame(state.viewportRenderFrame);
  state.viewportRenderFrame = null;
}



function cancelIdleViewportTrim() {
  if (!state.viewportTrimTimer) return;
  window.clearTimeout(state.viewportTrimTimer);
  state.viewportTrimTimer = null;
}



function activeSheetEditElement() {
  return $sheetHost.find(".sheet-cell[contenteditable='true']")[0] || null;
}



function activeCellEditElementForFormatting() {
  const editor = $sheetHost.find(".sheet-cell[contenteditable='true']:not(.is-formula-editing)")[0] || null;
  return editor?.isConnected ? editor : null;
}



function flushDeferredViewportRender() {
  if (!state.deferredViewportRender || activeSheetEditElement()) return;
  state.deferredViewportRender = false;
  scheduleViewportRender();
}



function restoreViewportScroll(anchor) {
  if (!anchor) return;
  const host = $sheetHost[0];
  if (!host) return;

  const nextTop = Number.isFinite(anchor.scrollTop)
    ? Math.max(0, Math.min(host.scrollHeight - host.clientHeight, anchor.scrollTop))
    : host.scrollTop;
  const nextLeft = Number.isFinite(anchor.scrollLeft)
    ? Math.max(0, Math.min(host.scrollWidth - host.clientWidth, anchor.scrollLeft))
    : host.scrollLeft;

  if (Math.abs(host.scrollTop - nextTop) < 0.5 && Math.abs(host.scrollLeft - nextLeft) < 0.5) return;

  state.restoringViewportScroll = true;
  host.scrollTop = nextTop;
  host.scrollLeft = nextLeft;
  if ($horizontalScrollbar[0]) $horizontalScrollbar[0].scrollLeft = nextLeft;
  requestAnimationFrame(() => {
    state.restoringViewportScroll = false;
  });
}



function resetLazyScrollLocks() {
  state.lazyScrollLocks.rows = false;
  state.lazyScrollLocks.cols = false;
  state.deferredViewportRender = false;
  cancelViewportRender();
  cancelIdleViewportTrim();
}



function ensureSheetCoversViewport(sheet) {
  const host = $sheetHost[0];
  const rowHeight = sheet.defaultRowHeight || rowHeightToPixels(15);
  const columnWidth = sheet.defaultColumnWidth || columnWidthToPixels(8.43);
  if (!host || !host.clientWidth || !host.clientHeight) {
    const viewportHeight = rowHeight * FALLBACK_VISIBLE_ROWS;
    const viewportWidth = columnWidth * FALLBACK_VISIBLE_COLS;
    const rowsNeeded = Math.max(1, Math.ceil((viewportHeight - COLUMN_HEADER_HEIGHT) / rowHeight) + 1);
    const colsNeeded = Math.max(1, Math.ceil((viewportWidth - rowHeaderWidthForSheet(sheet)) / columnWidth) + 1);
    ensureSheetSize(sheet, rowsNeeded, colsNeeded);
    return;
  }

  const viewportBottom = Math.max(COLUMN_HEADER_HEIGHT + 1, host.scrollTop + host.clientHeight);
  const viewportRight = Math.max(rowHeaderWidthForSheet(sheet) + 1, host.scrollLeft + host.clientWidth);
  const rowsNeeded = Math.max(1, indexForPixelExtent(sheet, "row", viewportBottom) + VIRTUAL_ROW_OVERSCAN);
  const colsNeeded = Math.max(1, indexForPixelExtent(sheet, "col", viewportRight) + VIRTUAL_COL_OVERSCAN);
  ensureSheetSize(sheet, rowsNeeded, colsNeeded);
}



function rowHeaderWidthForSheet(sheet) {
  const digits = String(rowHeaderReferenceRow(sheet)).length;
  const width = ROW_HEADER_PADDING_WIDTH + digits * ROW_HEADER_DIGIT_WIDTH;
  return Math.max(ROW_HEADER_MIN_WIDTH, Math.min(ROW_HEADER_MAX_WIDTH, Math.ceil(width)));
}



function rowHeaderReferenceRow(sheet) {
  const currentRows = Math.max(1, Number(sheet?.rowCount || 1));
  const host = $sheetHost[0];
  const rowsPerPage = host
    ? Math.max(1, indexForPixelExtent(sheet, "row", Math.max(COLUMN_HEADER_HEIGHT + 1, host.clientHeight)))
    : FALLBACK_VISIBLE_ROWS;
  return Math.max(1, Math.floor(currentRows + rowsPerPage));
}



function resizeActiveSheetForScroll({ allowExtend = true, allowTrim = true } = {}) {
  if (!state.model) return;
  const host = $sheetHost[0];
  if (!host) return;

  const bottomRemaining = host.scrollHeight - host.clientHeight - host.scrollTop;
  const rightRemaining = host.scrollWidth - host.clientWidth - host.scrollLeft;
  if (bottomRemaining > SCROLL_EXTEND_THRESHOLD * 2) state.lazyScrollLocks.rows = false;
  if (rightRemaining > SCROLL_EXTEND_THRESHOLD * 2) state.lazyScrollLocks.cols = false;
  const needsRows = allowExtend && bottomRemaining < SCROLL_EXTEND_THRESHOLD && !state.lazyScrollLocks.rows;
  const needsCols = allowExtend && rightRemaining < SCROLL_EXTEND_THRESHOLD && !state.lazyScrollLocks.cols;
  const sheet = activeSheet();
  const previousRowCount = sheet.rowCount;
  const previousColCount = sheet.colCount;
  const previousRowHeaderWidth = rowHeaderWidthForSheet(sheet);
  const previousScrollTop = host.scrollTop;
  const previousScrollLeft = host.scrollLeft;
  const viewportExtent = sheetExtentForViewport(sheet, host);
  const minimumExtent = minimumSheetExtent(sheet);
  let targetRows = Math.max(minimumExtent.rows, Math.min(sheet.rowCount, viewportExtent.rows));
  let targetCols = Math.max(minimumExtent.cols, Math.min(sheet.colCount, viewportExtent.cols));

  if (needsRows) targetRows = Math.max(targetRows, sheet.rowCount + SCROLL_EXTEND_ROWS);
  if (needsCols) targetCols = Math.max(targetCols, sheet.colCount + SCROLL_EXTEND_COLS);

  ensureSheetSize(sheet, targetRows, targetCols);
  if (allowTrim) trimSheetToViewportExtent(sheet, targetRows, targetCols);

  const addedRowsHeight = needsRows && sheet.rowCount > previousRowCount ? heightBetweenRows(sheet, previousRowCount + 1, sheet.rowCount) : 0;
  const addedColsWidth = needsCols && sheet.colCount > previousColCount ? widthBetweenColumns(sheet, previousColCount + 1, sheet.colCount) : 0;
  if (addedRowsHeight > 0) state.lazyScrollLocks.rows = true;
  if (addedColsWidth > 0) state.lazyScrollLocks.cols = true;
  const rowHeaderDelta = rowHeaderWidthForSheet(sheet) - previousRowHeaderWidth;
  const scrollTop = null;
  const scrollLeftDelta = previousScrollLeft > 0 ? rowHeaderDelta : 0;
  const scrollLeft = scrollLeftDelta ? previousScrollLeft + scrollLeftDelta : null;

  if (scrollTop == null && scrollLeft == null) return null;
  return {
    scrollTop,
    scrollLeft,
    renderScroll: {
      scrollTop: scrollTop ?? previousScrollTop,
      scrollLeft: scrollLeft ?? previousScrollLeft,
    },
  };
}



function ensureSheetSize(sheet, minRows, minCols) {
  const addRows = Math.max(0, minRows - sheet.rowCount);
  const addCols = Math.max(0, minCols - sheet.colCount);
  if (!addRows && !addCols) return;
  extendSheet(sheet, addRows, addCols);
}



function extendSheet(sheet, addRows, addCols) {
  const rowsToAdd = Math.max(0, addRows);
  const colsToAdd = Math.max(0, addCols);
  if (!rowsToAdd && !colsToAdd) return;

  const oldRowCount = sheet.rowCount;
  const oldColCount = sheet.colCount;
  const columnWidth = sheet.defaultColumnWidth || columnWidthToPixels(8.43);
  const rowHeight = sheet.defaultRowHeight || rowHeightToPixels(15);

  const formulaSheetExtended = tryExtendFormulaSheet(sheet, oldRowCount, oldColCount, rowsToAdd, colsToAdd);

  for (let i = 0; i < colsToAdd; i += 1) {
    sheet.columns.push({ width: columnWidth, hidden: false });
  }
  for (let i = 0; i < rowsToAdd; i += 1) {
    sheet.rows.push({ height: rowHeight, hidden: false });
  }

  sheet.rowCount += rowsToAdd;
  sheet.colCount += colsToAdd;
  sheet.hfData = [];

  if (!formulaSheetExtended) {
    state.hf = buildFormulaEngine(state.model);
  }
}



function trimSheetToViewportExtent(sheet, targetRows, targetCols) {
  const minimum = minimumSheetExtent(sheet);
  const nextRowCount = Math.max(1, minimum.rows, targetRows);
  const nextColCount = Math.max(1, minimum.cols, targetCols);
  if (nextRowCount >= sheet.rowCount && nextColCount >= sheet.colCount) return;

  const safeRowCount = Math.min(sheet.rowCount, nextRowCount);
  const safeColCount = Math.min(sheet.colCount, nextColCount);
  if (safeRowCount === sheet.rowCount && safeColCount === sheet.colCount) return;

  sheet.rows.length = safeRowCount;
  sheet.columns.length = safeColCount;
  sheet.hfData = [];

  Object.keys(sheet.cells || {}).forEach((key) => {
    const [row, col] = parseCellKey(key);
    if (row > safeRowCount || col > safeColCount) delete sheet.cells[key];
  });

  sheet.rowCount = safeRowCount;
  sheet.colCount = safeColCount;
  state.hf = buildFormulaEngine(state.model);
}



function sheetExtentForViewport(sheet, host) {
  const viewportBottom = host.scrollTop + host.clientHeight + SCROLL_TRIM_BUFFER_PX;
  const viewportRight = host.scrollLeft + host.clientWidth + SCROLL_TRIM_BUFFER_PX;
  return {
    rows: indexForPixelExtent(sheet, "row", viewportBottom) + VIRTUAL_ROW_OVERSCAN,
    cols: indexForPixelExtent(sheet, "col", viewportRight) + VIRTUAL_COL_OVERSCAN,
  };
}



function indexForPixelExtent(sheet, axis, pixel) {
  const isColumn = axis === "col";
  const headerSize = isColumn ? rowHeaderWidthForSheet(sheet) : COLUMN_HEADER_HEIGHT;
  const items = isColumn ? sheet.columns : sheet.rows;
  const defaultSize = isColumn ? sheet.defaultColumnWidth : sheet.defaultRowHeight;
  const fallbackSize = defaultSize || (isColumn ? columnWidthToPixels(8.43) : rowHeightToPixels(15));
  let position = headerSize;

  for (let index = 0; index < items.length; index += 1) {
    const item = items[index];
    const size = item?.hidden ? 0 : item?.width || item?.height || fallbackSize;
    const next = position + size;
    if (next >= pixel) return index + 1;
    position = next;
  }

  if (pixel <= position) return items.length;
  return items.length + Math.ceil((pixel - position) / Math.max(1, fallbackSize));
}



function minimumSheetExtent(sheet) {
  const revision = Math.max(0, Number(state.workbookRevision) || 0);
  const cached = sheetMinimumExtentCacheBySheet.get(sheet);
  if (
    cached &&
    cached.revision === revision &&
    cached.cells === sheet.cells &&
    cached.merges === sheet.merges &&
    cached.rows === sheet.rows &&
    cached.columns === sheet.columns &&
    cached.baseRowCount === sheet.baseRowCount &&
    cached.baseColCount === sheet.baseColCount &&
    cached.defaultRowHeight === sheet.defaultRowHeight &&
    cached.defaultColumnWidth === sheet.defaultColumnWidth
  ) {
    return cached.extent;
  }

  let rows = Math.max(1, Number(sheet.baseRowCount || 1));
  let cols = Math.max(1, Number(sheet.baseColCount || 1));
  const defaultRowHeight = sheet.defaultRowHeight || rowHeightToPixels(15);
  const defaultColumnWidth = sheet.defaultColumnWidth || columnWidthToPixels(8.43);

  Object.entries(sheet.cells || {}).forEach(([key, cell]) => {
    if (!isPersistentCell(cell)) return;
    const [row, col] = parseCellKey(key);
    rows = Math.max(rows, row);
    cols = Math.max(cols, col);
  });

  sheet.merges?.forEach((merge) => {
    rows = Math.max(rows, merge.bottom);
    cols = Math.max(cols, merge.right);
  });

  sheet.images?.forEach((image) => {
    const bottom = image.br?.row ?? image.br?.nativeRow ?? (image.tl?.row ?? image.tl?.nativeRow ?? 0) + 1;
    const right = image.br?.col ?? image.br?.nativeCol ?? (image.tl?.col ?? image.tl?.nativeCol ?? 0) + 1;
    rows = Math.max(rows, Math.ceil(Number(bottom) || 1));
    cols = Math.max(cols, Math.ceil(Number(right) || 1));
  });

  sheet.rows.forEach((row, index) => {
    if (row.hidden || Math.abs((row.height || 0) - defaultRowHeight) > 0.01) rows = Math.max(rows, index + 1);
  });
  sheet.columns.forEach((column, index) => {
    if (column.hidden || Math.abs((column.width || 0) - defaultColumnWidth) > 0.01) cols = Math.max(cols, index + 1);
  });

  const extent = { rows, cols };
  sheetMinimumExtentCacheBySheet.set(sheet, {
    revision,
    cells: sheet.cells,
    merges: sheet.merges,
    rows: sheet.rows,
    columns: sheet.columns,
    baseRowCount: sheet.baseRowCount,
    baseColCount: sheet.baseColCount,
    defaultRowHeight: sheet.defaultRowHeight,
    defaultColumnWidth: sheet.defaultColumnWidth,
    extent,
  });
  return extent;
}



function isPersistentCell(cell) {
  if (!cell) return false;
  return Boolean(
    cell.raw != null ||
      cell.formula ||
      cell.display ||
      cell.hyperlink ||
      cell.comment ||
      cell.note ||
      cell.phonetics?.length ||
      cell.showPhonetic ||
      cell.validation ||
      cell.sparkline ||
      cell.richText?.length ||
      Object.keys(cell.css || {}).length ||
      cell.hasFill ||
      cell.hasFontColor ||
      Object.keys(cell.borders || {}).length ||
      cell.numFmt,
  );
}



function parseCellKey(key) {
  const [row, col] = String(key).split(":").map((value) => Number(value) || 0);
  return [row, col];
}



function tryExtendFormulaSheet(sheet, oldRowCount, oldColCount, addRows, addCols) {
  const sheetId = getSheetId(sheet.name);
  if (!state.hf || sheetId == null) return true;
  try {
    if (addRows) {
      if (!state.hf.isItPossibleToAddRows?.(sheetId, [oldRowCount, addRows])) return false;
      state.hf.addRows(sheetId, [oldRowCount, addRows]);
    }
    if (addCols) {
      if (!state.hf.isItPossibleToAddColumns?.(sheetId, [oldColCount, addCols])) return false;
      state.hf.addColumns(sheetId, [oldColCount, addCols]);
    }
    return true;
  } catch (error) {
    console.warn("Formula sheet extension fallback:", error);
    return false;
  }
}



function switchRibbonTab(tabId) {
  if (!tabId) return;
  if (tabId === "file") {
    hideRibbonForAutoHide();
    openFileSideMenu();
    return;
  }
  const requestedTab = document.querySelector(`[data-ribbon-tab="${cssEscape(tabId)}"]`);
  if (requestedTab?.classList.contains("hidden")) return;

  closeFileSideMenu({ restoreHome: false });
  setActiveRibbonTab(tabId);
  showRibbonForAutoHide();
}



function setActiveRibbonTab(tabId) {
  $ribbonTabs.find(".ribbon-tab").each((_index, element) => {
    const isActive = element.dataset.ribbonTab === tabId;
    element.classList.toggle("is-active", isActive);
    element.setAttribute("aria-selected", String(isActive));
  });
  $ribbonPanels.find(".ribbon-panel").each((_index, element) => {
    element.classList.toggle("is-active", element.dataset.ribbonPanel === tabId);
  });
  syncViewRibbonButtons();
  syncPageLayoutRibbonControls();
  syncDrawRibbonButtons();
}



function toggleRibbonCollapsed(force = null) {
  state.ribbonCollapsed = force == null ? !state.ribbonCollapsed : Boolean(force);
  if (state.ribbonCollapsed) state.ribbonAutoOpen = false;
  syncRibbonCollapsed();
  closeFileSideMenu();
  closeRibbonDisplayMenu();
  closeHomeCommandMenu();
  closeViewCommandMenu();
  closeDataCommandMenu();
  requestAnimationFrame(() => {
    syncHorizontalScrollbarFromHost();
    scheduleViewportRender({ deferTrim: true });
  });
  setStatus(state.ribbonCollapsed ? "リボンを折りたたみました。" : "リボンを表示しました。");
}



function syncRibbonCollapsed() {
  const collapsed = state.ribbonCollapsed === true;
  $ribbonPanels.closest(".ribbon").toggleClass("is-collapsed", collapsed);
  syncRibbonDisplayState();
}



function normalizeRibbonLayout(value) {
  return RIBBON_LAYOUTS.has(value) ? value : "classic";
}



function normalizeRibbonVisibility(value) {
  return RIBBON_VISIBILITY_MODES.has(value) ? value : "always";
}



function loadRibbonDisplayPreferences() {
  try {
    state.ribbonLayout = normalizeRibbonLayout(localStorage.getItem(RIBBON_LAYOUT_STORAGE_KEY));
    state.ribbonVisibility = normalizeRibbonVisibility(localStorage.getItem(RIBBON_VISIBILITY_STORAGE_KEY));
  } catch {
    state.ribbonLayout = "classic";
    state.ribbonVisibility = "always";
  }
  state.ribbonAutoOpen = false;
}



function saveRibbonDisplayPreferences() {
  try {
    localStorage.setItem(RIBBON_LAYOUT_STORAGE_KEY, state.ribbonLayout);
    localStorage.setItem(RIBBON_VISIBILITY_STORAGE_KEY, state.ribbonVisibility);
  } catch {
    // Keep the visible state even when browser storage is unavailable.
  }
}



function ribbonPanelsAreVisible() {
  if (state.ribbonCollapsed === true) return false;
  if (state.ribbonVisibility === "auto" && !state.ribbonAutoOpen) return false;
  return true;
}



function syncRibbonDisplayState(options = {}) {
  const ribbon = $ribbonPanels.closest(".ribbon");
  const singleLine = state.ribbonLayout === "single";
  const autoHide = state.ribbonVisibility === "auto";
  ribbon
    .toggleClass("ribbon-layout-single", singleLine)
    .toggleClass("ribbon-layout-classic", !singleLine)
    .toggleClass("ribbon-visibility-auto", autoHide)
    .toggleClass("ribbon-visibility-always", !autoHide)
    .toggleClass("ribbon-auto-open", Boolean(state.ribbonAutoOpen));
  $ribbonPanels.attr("aria-hidden", ribbonPanelsAreVisible() ? "false" : "true");
  $ribbonDisplayMenu.find("[data-ribbon-layout-option]").each((_index, element) => {
    const checked = element.dataset.ribbonLayoutOption === state.ribbonLayout;
    element.setAttribute("aria-checked", String(checked));
  });
  $ribbonDisplayMenu.find("[data-ribbon-visibility-option]").each((_index, element) => {
    const checked = element.dataset.ribbonVisibilityOption === state.ribbonVisibility;
    element.setAttribute("aria-checked", String(checked));
  });
  if (options.layout !== false) {
    requestAnimationFrame(() => {
      syncHorizontalScrollbarFromHost();
      scheduleViewportRender({ deferTrim: true });
    });
  }
}



function setRibbonLayout(layout) {
  const nextLayout = normalizeRibbonLayout(layout);
  if (state.ribbonLayout === nextLayout) return;
  state.ribbonLayout = nextLayout;
  saveRibbonDisplayPreferences();
  syncRibbonDisplayState();
  setStatus(nextLayout === "single" ? "リボンを 1 行表示にしました。" : "クラシック リボンを表示しました。");
}



function setRibbonVisibility(visibility) {
  const nextVisibility = normalizeRibbonVisibility(visibility);
  if (state.ribbonVisibility === nextVisibility && state.ribbonAutoOpen === false && state.ribbonCollapsed === false) return;
  state.ribbonVisibility = nextVisibility;
  state.ribbonAutoOpen = false;
  state.ribbonCollapsed = false;
  saveRibbonDisplayPreferences();
  syncRibbonCollapsed();
  setStatus(nextVisibility === "auto" ? "リボンを自動的に非表示にします。" : "リボンを常に表示します。");
}



function showRibbonForAutoHide() {
  if (state.ribbonVisibility !== "auto" || state.ribbonCollapsed) return;
  if (state.ribbonAutoOpen) return;
  state.ribbonAutoOpen = true;
  syncRibbonDisplayState();
}



function hideRibbonForAutoHide() {
  if (state.ribbonVisibility !== "auto" || !state.ribbonAutoOpen) return;
  state.ribbonAutoOpen = false;
  syncRibbonDisplayState();
}



function openRibbonDisplayMenu(event = null) {
  event?.preventDefault?.();
  event?.stopPropagation?.();
  closeFileSideMenu({ restoreHome: false });
  showRibbonForAutoHide();
  syncRibbonDisplayState({ layout: false });
  $ribbonDisplayMenu.removeClass("hidden");
  $ribbonDisplayButton.attr("aria-expanded", "true");
  $ribbonDisplayMenu.find("button").first().trigger("focus");
}



function closeRibbonDisplayMenu() {
  $ribbonDisplayMenu.addClass("hidden");
  $ribbonDisplayButton.attr("aria-expanded", "false");
}



function toggleRibbonDisplayMenu(event) {
  if ($ribbonDisplayMenu.hasClass("hidden")) {
    openRibbonDisplayMenu(event);
    return;
  }
  event?.preventDefault?.();
  event?.stopPropagation?.();
  closeRibbonDisplayMenu();
}



function handleRibbonLayoutOptionClick(event) {
  event.preventDefault();
  setRibbonLayout(event.currentTarget.dataset.ribbonLayoutOption);
  closeRibbonDisplayMenu();
  $ribbonDisplayButton.trigger("focus");
}



function handleRibbonVisibilityOptionClick(event) {
  event.preventDefault();
  setRibbonVisibility(event.currentTarget.dataset.ribbonVisibilityOption);
  closeRibbonDisplayMenu();
  $ribbonDisplayButton.trigger("focus");
}



function ribbonDisplayMenuButtons() {
  return [...$ribbonDisplayMenu.find("button")].filter((button) => !button.disabled && button.offsetParent !== null);
}



function focusRibbonDisplayMenuButton(current, direction) {
  const buttons = ribbonDisplayMenuButtons();
  if (!buttons.length) return;
  const index = Math.max(0, buttons.indexOf(current));
  buttons[(index + direction + buttons.length) % buttons.length].focus({ preventScroll: true });
}



function handleRibbonDisplayMenuKeydown(event) {
  if (isElementImeCompositionEvent(event)) return;
  if (event.key === "Escape") {
    consumeCommandMenuKey(event);
    closeRibbonDisplayMenu();
    $ribbonDisplayButton.trigger("focus");
    return;
  }
  if (event.key === "ArrowDown" || event.key === "ArrowUp") {
    consumeCommandMenuKey(event);
    focusRibbonDisplayMenuButton(event.target, event.key === "ArrowDown" ? 1 : -1);
    return;
  }
  if (event.key === "Home" || event.key === "End") {
    const buttons = ribbonDisplayMenuButtons();
    if (!buttons.length) return;
    consumeCommandMenuKey(event);
    buttons[event.key === "Home" ? 0 : buttons.length - 1].focus({ preventScroll: true });
  }
}



function closeRibbonDisplayMenuOnOutsidePointer(event) {
  if ($ribbonDisplayMenu.hasClass("hidden")) return;
  if (event.target.closest?.("#ribbonDisplayMenu, #ribbonDisplayButton")) return;
  closeRibbonDisplayMenu();
}



function closeRibbonDisplayMenuOnEscape(event) {
  if (event.key !== "Escape" || $ribbonDisplayMenu.hasClass("hidden")) return;
  consumeCommandMenuKey(event);
  closeRibbonDisplayMenu();
  $ribbonDisplayButton.trigger("focus");
}



function hideRibbonForAutoHideOnOutsidePointer(event) {
  if (state.ribbonVisibility !== "auto" || !state.ribbonAutoOpen) return;
  if (event.target.closest?.(".ribbon, #fileSideMenu, #homeCommandMenu, #viewCommandMenu, #dataCommandMenu, #dataValidationMenu, #whatIfMenu, #ribbonColorPalette, #ribbonBorderMenu, #ribbonUnderlineMenu, #ribbonInkEraserMenu, #ribbonShapeMenu, #shapeFormatMenu")) {
    return;
  }
  hideRibbonForAutoHide();
}



function handleRibbonAccessKeyShortcut(event, key) {
  if (!event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return false;
  if (isTextEditingTarget(event.target)) return false;
  if (event.target?.closest?.(".excel-dialog-overlay, #fileSideMenu, #ribbonDisplayMenu, #autoFilterMenu, #homeCommandMenu, #viewCommandMenu, #dataCommandMenu, #whatIfMenu, #ribbonColorPalette, #ribbonBorderMenu, #ribbonUnderlineMenu, #ribbonInkEraserMenu, #ribbonShapeMenu, #shapeFormatMenu, #nameBoxMenu")) {
    return false;
  }
  const tabByKey = {
    f: "file",
    h: "home",
    n: "insert",
    d: "draw",
    p: "layout",
    m: "formula",
    a: "data",
    w: "view",
    l: "help",
  };
  const tabId = tabByKey[key];
  if (!tabId) return false;
  const tab = document.querySelector(`[data-ribbon-tab="${cssEscape(tabId)}"]`);
  if (!tab || tab.classList.contains("hidden")) return false;
  event.preventDefault();
  switchRibbonTab(tabId);
  if (tabId !== "file") setStatus(`${tab.textContent.trim()} タブを表示しました。`);
  return true;
}



function handleRibbonFocusShortcut(event, key) {
  if ((key !== "f10" && key !== "alt") || event.ctrlKey || event.metaKey || event.shiftKey) return false;
  if (isTextEditingTarget(event.target)) return false;
  if (event.target?.closest?.(".excel-dialog-overlay, #fileSideMenu, #ribbonDisplayMenu, #autoFilterMenu, #homeCommandMenu, #viewCommandMenu, #dataCommandMenu, #whatIfMenu, #ribbonColorPalette, #ribbonBorderMenu, #ribbonUnderlineMenu, #ribbonInkEraserMenu, #ribbonShapeMenu, #shapeFormatMenu, #nameBoxMenu")) {
    return false;
  }
  const tab = focusActiveRibbonTab();
  if (!tab) return false;
  event.preventDefault();
  setStatus(`${tab.textContent.trim()} タブを選択しました。`);
  return true;
}



function handleWorkbookPaneFocusShortcut(event, key) {
  if (key !== "f6" || event.ctrlKey || event.metaKey || event.altKey) return false;
  if (!state.model || activeSheetEditElement()) return false;
  if (commandMenuEscapeBlockedByDialog() || document.querySelector("#fileSideMenu:not(.hidden), #ribbonDisplayMenu:not(.hidden), #homeCommandMenu, #viewCommandMenu, #dataCommandMenu, #dataValidationMenu, #whatIfMenu, #ribbonColorPalette, #ribbonBorderMenu, #ribbonUnderlineMenu, #ribbonInkEraserMenu, #ribbonShapeMenu, #shapeFormatMenu, #sheetTabContextMenu, #sheetTabNavigatorMenu, #cellContextMenu, #statusBarContextMenu, #queryPaneContextMenu, #quickAnalysisPopup, #nameBoxMenu")) {
    return false;
  }
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  focusWorkbookPaneByKeyboard(event.shiftKey ? -1 : 1);
  return true;
}



function workbookPaneFocusZones() {
  const zones = ["worksheet", "name", "formula", "ribbon"];
  if ($formulaInput.prop("disabled")) return zones.filter((zone) => zone !== "formula");
  return zones;
}



function currentWorkbookPaneFocusZone() {
  const active = document.activeElement;
  if (active?.closest?.("#ribbonTabs, #ribbonPanels, .ribbon-display-options")) return "ribbon";
  if (active === $selectedCellLabel[0]) return "name";
  if (active === $formulaInput[0]) return "formula";
  return "worksheet";
}



function focusWorkbookPaneByKeyboard(direction = 1) {
  const zones = workbookPaneFocusZones();
  const current = currentWorkbookPaneFocusZone();
  const currentIndex = Math.max(0, zones.indexOf(current));
  const nextZone = zones[(currentIndex + direction + zones.length) % zones.length];
  focusWorkbookPaneZone(nextZone);
}



function focusWorkbookPaneZone(zone) {
  switch (zone) {
    case "name":
      $selectedCellLabel[0]?.focus?.({ preventScroll: true });
      setStatus("名前ボックスに移動しました。");
      break;
    case "formula":
      $formulaInput[0]?.focus?.({ preventScroll: true });
      setStatus("数式バーに移動しました。");
      break;
    case "ribbon": {
      const tab = focusActiveRibbonTab();
      if (tab) setStatus(`${tab.textContent.trim()} タブに移動しました。`);
      break;
    }
    case "worksheet":
    default:
      restoreWorksheetFocusAfterCommandPopup();
      setStatus("ワークシートに移動しました。");
      break;
  }
}



function handleKeyboardContextMenuShortcut(event, key) {
  const isContextMenuKey = key === "contextmenu" || key === "apps";
  const isShiftF10 = key === "f10" && event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey;
  if (!isContextMenuKey && !isShiftF10) return false;
  if (!state.model || isTextEditingTarget(event.target)) return false;
  if (commandMenuEscapeBlockedByDialog() || document.querySelector("#fileSideMenu:not(.hidden), #ribbonDisplayMenu:not(.hidden), #homeCommandMenu, #viewCommandMenu, #dataCommandMenu, #dataValidationMenu, #whatIfMenu, #ribbonColorPalette, #ribbonBorderMenu, #ribbonUnderlineMenu, #ribbonInkEraserMenu, #ribbonShapeMenu, #shapeFormatMenu, #sheetTabContextMenu, #sheetTabNavigatorMenu, #cellContextMenu, #statusBarContextMenu, #queryPaneContextMenu, #quickAnalysisPopup, #nameBoxMenu")) {
    return false;
  }
  const sheetTab = event.target?.closest?.(".sheet-tab") || document.activeElement?.closest?.(".sheet-tab");
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  if (sheetTab && openSheetTabContextMenuByKeyboard(sheetTab)) return true;
  return openSelectionContextMenuByKeyboard();
}



function handleWorkbookKeyboardCaptureShortcut(event) {
  if (event.defaultPrevented) return;
  const key = String(event.key || "").toLowerCase();
  if (key !== "f6" || event.ctrlKey || event.metaKey || event.altKey) return;
  const target = event.target;
  if (target !== $formulaInput[0] && target !== $selectedCellLabel[0]) return;
  handleWorkbookPaneFocusShortcut(event, key);
}



function updateWorkbookModifierState(event) {
  const key = String(event?.key || "").toLowerCase();
  state.autoFillModifierActive = Boolean(event?.ctrlKey || event?.metaKey || (event?.type === "keydown" && (key === "control" || key === "meta")));
}



function resetWorkbookModifierState() {
  state.autoFillModifierActive = false;
}



function openSelectionContextMenuByKeyboard() {
  const sheet = activeSheet();
  if (!sheet) return false;
  const objects = selectedImageObjects();
  if (objects.length) {
    const image = selectedImageObjectById(state.selectedImage?.id) || objects[0];
    const point = keyboardContextMenuPointForImage(sheet, image);
    closeSheetTabContextMenu();
    closeFileSideMenu({ restoreHome: true });
    renderImageContextMenu(image.id, point.clientX, point.clientY);
    return true;
  }

  const range = activeSelectionRange();
  if (!range) return false;
  closeSheetTabContextMenu();
  closeFileSideMenu({ restoreHome: true });
  closeQuickAnalysisPopup();
  if (range.fullColumn) {
    const index = state.selected?.sheetIndex === state.activeSheetIndex && state.selected.col >= range.left && state.selected.col <= range.right ? state.selected.col : range.left;
    const point = keyboardContextMenuPointForColumnHeader(sheet, index);
    renderHeaderContextMenu("column", index, point.clientX, point.clientY);
    return true;
  }
  if (range.fullRow) {
    const index = state.selected?.sheetIndex === state.activeSheetIndex && state.selected.row >= range.top && state.selected.row <= range.bottom ? state.selected.row : range.top;
    const point = keyboardContextMenuPointForRowHeader(sheet, index);
    renderHeaderContextMenu("row", index, point.clientX, point.clientY);
    return true;
  }

  const row = state.selected?.sheetIndex === state.activeSheetIndex && state.selected.row >= range.top && state.selected.row <= range.bottom ? state.selected.row : range.top;
  const col = state.selected?.sheetIndex === state.activeSheetIndex && state.selected.col >= range.left && state.selected.col <= range.right ? state.selected.col : range.left;
  const point = keyboardContextMenuPointForCell(sheet, row, col);
  renderCellContextMenu(row, col, point.clientX, point.clientY);
  return true;
}



function keyboardContextMenuPointForCell(sheet, row, col) {
  const left = coordinateToPixels(sheet, "col", col - 1);
  const top = coordinateToPixels(sheet, "row", row - 1);
  const right = coordinateToPixels(sheet, "col", col);
  const bottom = coordinateToPixels(sheet, "row", row);
  return sheetPixelToClientPoint(Math.min(right - 2, left + Math.max(18, (right - left) * 0.72)), Math.min(bottom - 2, top + Math.max(14, (bottom - top) * 0.72)));
}



function keyboardContextMenuPointForColumnHeader(sheet, col) {
  const left = coordinateToPixels(sheet, "col", col - 1);
  const right = coordinateToPixels(sheet, "col", col);
  return sheetPixelToClientPoint((left + right) / 2, Math.max(4, COLUMN_HEADER_HEIGHT - 4));
}



function keyboardContextMenuPointForRowHeader(sheet, row) {
  const top = coordinateToPixels(sheet, "row", row - 1);
  const bottom = coordinateToPixels(sheet, "row", row);
  return sheetPixelToClientPoint(Math.max(4, rowHeaderWidthForSheet(sheet) - 4), (top + bottom) / 2);
}



function keyboardContextMenuPointForImage(sheet, image) {
  const box = imageBoxToPixels(sheet, image);
  if (!box) return sheetPixelToClientPoint(rowHeaderWidthForSheet(sheet) + 24, COLUMN_HEADER_HEIGHT + 24);
  return sheetPixelToClientPoint(box.left + Math.min(box.width - 2, Math.max(20, box.width * 0.72)), box.top + Math.min(box.height - 2, Math.max(18, box.height * 0.72)));
}



function sheetPixelToClientPoint(x, y) {
  const host = $sheetHost[0];
  const rect = host?.getBoundingClientRect?.();
  if (!host || !rect) return { clientX: Math.max(8, Math.round(x || 8)), clientY: Math.max(8, Math.round(y || 8)) };
  const margin = 8;
  return {
    clientX: Math.round(Math.max(rect.left + margin, Math.min(rect.right - margin, rect.left + Number(x || 0) - host.scrollLeft))),
    clientY: Math.round(Math.max(rect.top + margin, Math.min(rect.bottom - margin, rect.top + Number(y || 0) - host.scrollTop))),
  };
}



function openFileSideMenu() {
  closeSheetTabContextMenu();
  closeCellContextMenu();
  syncFileSideMenuLocalStandaloneItems();
  positionFileSideMenu();
  enhanceCommandMenuAccessKeys($fileSideMenu[0]);
  $fileSideMenu.removeClass("hidden");
  $fileSideMenu.find("button:not([disabled])").first().trigger("focus");
}



function shouldShowFileSideMenuWebVersionItem() {
  return Boolean(state.isStandalone && window.location.protocol === "file:");
}



function syncFileSideMenuLocalStandaloneItems() {
  const show = shouldShowFileSideMenuWebVersionItem();
  $fileSideMenu.find("[data-local-standalone-menu]").each((_index, element) => {
    element.classList.toggle("hidden", !show);
    element.setAttribute("aria-hidden", show ? "false" : "true");
  });
}



function positionFileSideMenu() {
  const fileTab = document.querySelector('[data-ribbon-tab="file"]');
  const rect = fileTab?.getBoundingClientRect();
  const left = Math.max(8, Math.round(rect?.left || 8));
  const top = Math.max(8, Math.round(rect?.bottom || 44));
  $fileSideMenu.css({
    left: `${left}px`,
    top: `${top}px`,
    maxHeight: `${Math.max(160, window.innerHeight - top - 8)}px`,
  });
}



function closeFileSideMenu(options = {}) {
  $fileSideMenu.addClass("hidden");
}



function closeFileSideMenuOnOutsidePointer(event) {
  if ($fileSideMenu.hasClass("hidden")) return;
  if (event.target.closest?.("#fileSideMenu") || event.target.closest?.('[data-ribbon-tab="file"]')) return;
  closeFileSideMenu({ restoreHome: true });
}



function closeFileSideMenuOnEscape(event) {
  if (event.key !== "Escape" || $fileSideMenu.hasClass("hidden")) return;
  consumeCommandMenuKey(event);
  closeFileSideMenu({ restoreHome: true });
  restoreWorksheetFocusAfterCommandPopup();
}



function visibleRibbonTabs() {
  return [...document.querySelectorAll("#ribbonTabs .ribbon-tab")].filter((tab) => {
    if (tab.disabled || tab.classList.contains("hidden")) return false;
    return tab.offsetParent !== null;
  });
}



function activeRibbonTabElement() {
  const active = document.querySelector("#ribbonTabs .ribbon-tab.is-active:not(.hidden)");
  if (active?.offsetParent !== null) return active;
  return visibleRibbonTabs()[0] || null;
}



function focusRibbonTabAt(index) {
  const tabs = visibleRibbonTabs();
  if (!tabs.length) return null;
  const normalized = (index + tabs.length) % tabs.length;
  const tab = tabs[normalized];
  tab.focus({ preventScroll: true });
  return tab;
}



function focusActiveRibbonTab() {
  const active = activeRibbonTabElement();
  active?.focus?.({ preventScroll: true });
  return active;
}



function activeRibbonPanelElement() {
  const activeTab = activeRibbonTabElement();
  const tabId = activeTab?.dataset.ribbonTab || "home";
  let panel = document.querySelector(`#ribbonPanels .ribbon-panel[data-ribbon-panel="${cssEscape(tabId)}"]`);
  if (!panel?.classList.contains("is-active")) panel = document.querySelector("#ribbonPanels .ribbon-panel.is-active");
  if (panel?.offsetParent === null) return null;
  return panel || null;
}



function ribbonPanelFocusableControls(panel = activeRibbonPanelElement()) {
  if (!panel) return [];
  return [...panel.querySelectorAll(RIBBON_PANEL_FOCUSABLE_SELECTOR)].filter((control) => {
    if (control.disabled || control.getAttribute("aria-disabled") === "true") return false;
    if (control.classList.contains("hidden") || control.classList.contains("sr-only")) return false;
    if (control.closest(".hidden, .sr-only")) return false;
    return control.offsetParent !== null;
  });
}



function ribbonPanelFocusedControl(panel = activeRibbonPanelElement()) {
  const focused = document.activeElement?.closest?.(RIBBON_PANEL_FOCUSABLE_SELECTOR);
  if (!focused || !panel?.contains(focused)) return null;
  return focused;
}



function focusRibbonPanelControl(control) {
  control?.focus?.({ preventScroll: true });
}



function focusFirstRibbonPanelControl() {
  const control = ribbonPanelFocusableControls()[0] || null;
  focusRibbonPanelControl(control);
  return control;
}



function isRibbonPanelNativeKeyTarget(element) {
  return Boolean(element?.closest?.("input, select, textarea, [contenteditable='true']"));
}



function handleRibbonTabKeydown(event) {
  if (isElementImeCompositionEvent(event) || event.defaultPrevented) return;
  if (event.ctrlKey || event.metaKey || event.altKey) return;
  const tabs = visibleRibbonTabs();
  if (!tabs.length) return;
  const tab = event.currentTarget;
  const currentIndex = Math.max(0, tabs.indexOf(tab));
  if (event.key === "ArrowLeft" || event.key === "ArrowRight") {
    const direction = event.key === "ArrowRight" ? 1 : -1;
    focusRibbonTabAt(currentIndex + direction);
    consumeCommandMenuKey(event);
    return;
  }
  if (event.key === "Home" || event.key === "End") {
    focusRibbonTabAt(event.key === "Home" ? 0 : tabs.length - 1);
    consumeCommandMenuKey(event);
    return;
  }
  if (event.key === "ArrowDown") {
    if (tab.dataset.ribbonTab === "file") {
      openFileSideMenu();
    } else {
      switchRibbonTab(tab.dataset.ribbonTab);
      focusFirstRibbonPanelControl();
    }
    consumeCommandMenuKey(event);
    return;
  }
  if (event.key === "Enter" || event.key === " ") {
    switchRibbonTab(tab.dataset.ribbonTab);
    if (tab.dataset.ribbonTab !== "file") setStatus(`${tab.textContent.trim()} タブを表示しました。`);
    consumeCommandMenuKey(event);
  }
}



function handleRibbonPanelKeydown(event) {
  if (isElementImeCompositionEvent(event) || event.defaultPrevented) return;
  if (event.ctrlKey || event.metaKey || event.altKey) return;
  const panel = activeRibbonPanelElement();
  if (!panel || !panel.contains(event.target)) return;
  if (event.key === "Escape") {
    focusActiveRibbonTab();
    consumeCommandMenuKey(event);
    return;
  }
  if (isRibbonPanelNativeKeyTarget(event.target)) return;
  const keys = ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Home", "End"];
  if (!keys.includes(event.key)) return;
  const controls = ribbonPanelFocusableControls(panel);
  if (!controls.length) return;
  const focused = ribbonPanelFocusedControl(panel);
  const currentIndex = focused ? controls.indexOf(focused) : -1;
  const nextIndex =
    event.key === "Home"
      ? 0
      : event.key === "End"
        ? controls.length - 1
        : nextCommandMenuItemIndexByArrow(controls, currentIndex, event.key);
  focusRibbonPanelControl(controls[nextIndex] || controls[0]);
  consumeCommandMenuKey(event);
}



function releaseRibbonPointerFocus(trigger, event) {
  const nativeEvent = event?.originalEvent || event;
  if (!trigger || nativeEvent?.detail === 0) return;
  const active = document.activeElement;
  if (active === trigger || trigger.contains?.(active)) active.blur?.();
}



function openNewWorkbookFromFileMenu() {
  if (!state.model || !workbookHasUnsavedChanges()) {
    openBlankWorkbook({ status: "新しいブックを作成しました。" });
    return;
  }
  showUnsavedWorkbookReplacementMessage({
    message: "現在のブックに保存されていない変更があります。新しいブックを開く前に保存しますか?",
    saveStatus: "保存してから新しいブックを作成しました。",
    discardStatus: "保存せずに新しいブックを作成しました。",
    cancelStatus: "新しいブックの作成をキャンセルしました。",
    action: () => {
      openBlankWorkbook({ status: "新しいブックを作成しました。" });
      return true;
    },
  });
}



async function handleRibbonCommand(command, trigger = null, event = null) {
  if (!command) return;

  if (command === "open") {
    await openWorkbookFilePicker();
    return;
  }

  if (command === "open-web-version") {
    openWebVersionFromLocalFile();
    return;
  }

  if (!state.model && !WORKBOOK_OPTIONAL_RIBBON_COMMANDS.has(command)) {
    setStatus("先に Excel ファイルを開いてください。");
    return;
  }

  if (command !== "format-painter" && state.formatPainter) {
    clearFormatPainter({ silent: true });
  }

  switch (command) {
    case "help":
      openHelpDialog();
      break;
    case "version":
      openVersionDialog();
      break;
    case "save":
      await quickSaveWorkbook();
      break;
    case "save-as":
      openSaveAsDialog();
      break;
    case "print":
      openPrintDialog();
      break;
    case "new-workbook":
      openNewWorkbookFromFileMenu();
      break;
    case "new-sheet":
      addBlankSheet();
      break;
    case "insert-pivot-table":
    case "insert-recommended-pivot":
      insertPivotSummaryFromSelection();
      break;
    case "insert-table":
      openCreateTableDialog();
      break;
    case "insert-picture":
      await choosePictureForInsertion();
      break;
    case "insert-icons":
      openInsertIconDialog();
      break;
    case "insert-3d-model":
      insert3dModelPlaceholder();
      break;
    case "insert-smartart":
      insertSmartArtObject();
      break;
    case "insert-screenshot":
      insertScreenshotObject();
      break;
    case "insert-get-addins":
      openOfficeAddinsDialog("store");
      break;
    case "insert-my-addins":
      openOfficeAddinsDialog("mine");
      break;
    case "insert-recommended-chart":
      insertChartFromSelection(recommendedChartTypeForSelection());
      break;
    case "insert-column-chart":
      insertChartFromSelection("column");
      break;
    case "insert-line-chart":
      insertChartFromSelection("line");
      break;
    case "insert-pie-chart":
      insertChartFromSelection("pie");
      break;
    case "insert-scatter-chart":
      insertChartFromSelection("scatter");
      break;
    case "insert-map-chart":
      insertChartFromSelection("map");
      break;
    case "insert-pivot-chart":
      insertChartFromSelection("column", { pivot: true });
      break;
    case "insert-sparkline-line":
      insertSparklinesFromSelection("line");
      break;
    case "insert-sparkline-column":
      insertSparklinesFromSelection("column");
      break;
    case "insert-sparkline-winloss":
      insertSparklinesFromSelection("winloss");
      break;
    case "insert-slicer":
      insertSlicerObject();
      break;
    case "insert-timeline":
      insertTimelineObject();
      break;
    case "insert-shape":
      beginShapeDrawingMode(state.currentShapeType || DEFAULT_SHAPE_TYPE);
      break;
    case "insert-text-box":
      insertTextBoxObject();
      break;
    case "insert-header-footer":
      openHeaderFooterDialog();
      break;
    case "insert-wordart":
      insertWordArtObject();
      break;
    case "insert-signature-line":
      openSignatureLineDialog();
      break;
    case "insert-object":
      openInsertObjectDialog();
      break;
    case "insert-equation":
      insertEquationObject();
      break;
    case "insert-link":
      openLinkDialogForSelection();
      break;
    case "insert-symbol":
      openSymbolDialog();
      break;
    case "shape-fill-blue":
      applySelectedShapeStyle({ fillColor: "#dbeafe" }, "図形の塗りつぶし");
      break;
    case "shape-fill-none":
      applySelectedShapeStyle({ fillColor: "none" }, "塗りつぶしなし");
      break;
    case "shape-outline-blue":
      applySelectedShapeStyle({ outlineColor: "#4472c4", outlineWidth: DEFAULT_SHAPE_OUTLINE_WIDTH }, "図形の枠線");
      break;
    case "shape-outline-none":
      applySelectedShapeStyle({ outlineColor: "none" }, "枠線なし");
      break;
    case "shape-edit-text":
      openShapeTextEditor();
      break;
    case "shape-bring-front":
      reorderSelectedShape("forward");
      break;
    case "shape-send-back":
      reorderSelectedShape("backward");
      break;
    case "shape-selection-pane":
      openShapeSelectionPane();
      break;
    case "shape-group":
      groupSelectedShapes();
      break;
    case "shape-ungroup":
      ungroupSelectedShapes();
      break;
    case "shape-reset-size":
      resetSelectedShapeSize();
      break;
    case "picture-crop":
      togglePictureCropMode();
      break;
    case "copy":
      executeWorkbookClipboardShortcut("copy", () => {
        if (selectedImageObjects().length) void copySelectedImages({ systemClipboard: true });
        else void copySelectedRange({ systemClipboard: true });
      });
      break;
    case "cut":
      executeWorkbookClipboardShortcut("cut", () => {
        if (selectedImageObjects().length) void cutSelectedImages({ systemClipboard: true });
        else void cutSelectedRange({ systemClipboard: true });
      });
      break;
    case "paste":
      await pasteFromRibbonButton();
      break;
    case "format-painter":
      toggleFormatPainter({ repeat: isFormatPainterRepeatTrigger(event) });
      break;
    case "bold":
      toggleStyleForSelection("fontWeight", "700", "normal", "太字");
      break;
    case "italic":
      toggleStyleForSelection("fontStyle", "italic", "normal", "斜体");
      break;
    case "underline":
      setRibbonUnderlinePreset("underline");
      applyUnderlinePreset("underline", { toggle: true });
      break;
    case "align-left":
      applyCssToSelection({ textAlign: "left" }, "左揃え");
      break;
    case "align-center":
      applyCssToSelection({ textAlign: "center" }, "中央揃え");
      break;
    case "align-right":
      applyCssToSelection({ textAlign: "right" }, "右揃え");
      break;
    case "wrap-text":
      toggleWrapTextForSelection();
      break;
    case "valign-top":
      applyCssToSelection({ verticalAlign: "top", alignItems: "flex-start" }, "上揃え");
      break;
    case "valign-middle":
      applyCssToSelection({ verticalAlign: "middle", alignItems: "center" }, "上下中央揃え");
      break;
    case "valign-bottom":
      applyCssToSelection({ verticalAlign: "bottom", alignItems: "flex-end" }, "下揃え");
      break;
    case "merge":
      mergeSelection(false);
      break;
    case "merge-center":
      toggleMergeCenterSelection();
      break;
    case "unmerge":
      unmergeSelection();
      break;
    case "insert-cells":
      openCellStructureDialog("insert");
      break;
    case "delete-cells":
      openCellStructureDialog("delete");
      break;
    case "format-cells":
      openFormatCellsDialog();
      break;
    case "autosum":
      applyAutoSumToSelection();
      break;
    case "formula-autosum":
      openFormulaAutoSumMenu(trigger);
      break;
    case "formula-insert-function":
      openInsertFunctionDialog();
      break;
    case "formula-function-recent":
      openFormulaFunctionMenu(trigger, "recent");
      break;
    case "formula-function-financial":
      openFormulaFunctionMenu(trigger, "financial");
      break;
    case "formula-function-sum":
      openFormulaFunctionMenu(trigger, "math");
      break;
    case "formula-function-if":
      openFormulaFunctionMenu(trigger, "logical");
      break;
    case "formula-function-xlookup":
      openFormulaFunctionMenu(trigger, "lookup");
      break;
    case "formula-function-today":
      openFormulaFunctionMenu(trigger, "date");
      break;
    case "formula-function-textjoin":
      openFormulaFunctionMenu(trigger, "text");
      break;
    case "formula-function-more":
      openFormulaFunctionMenu(trigger, "more");
      break;
    case "formula-name-manager":
      openNameManagerDialog();
      break;
    case "formula-define-name": {
      if (event?.target?.closest?.(".ribbon-command-caret")) {
        openFormulaDefineNameMenu(trigger);
      } else {
        openDefineNameForSelection();
      }
      break;
    }
    case "formula-use-in-formula":
      openUseInFormulaMenu(trigger);
      break;
    case "formula-paste-name":
      openPasteNameDialog();
      break;
    case "formula-create-from-selection":
      openCreateNamesDialog();
      break;
    case "formula-trace-precedents":
      traceFormulaReferences("precedents");
      break;
    case "formula-trace-dependents":
      traceFormulaReferences("dependents");
      break;
    case "formula-remove-arrows":
      openFormulaRemoveArrowsMenu(trigger);
      break;
    case "formula-show-formulas":
      toggleShowFormulasFromRibbon();
      break;
    case "formula-error-checking":
      openFormulaErrorCheckingMenu(trigger);
      break;
    case "formula-evaluate":
      openEvaluateFormulaDialog();
      break;
    case "formula-watch-window":
      openFormulaWatchWindow();
      break;
    case "formula-calculation-options":
      openCalculationOptionsMenu(trigger);
      break;
    case "formula-calculate-now":
      calculateWorkbookNow();
      break;
    case "formula-calculate-sheet":
      calculateActiveSheetNow();
      break;
    case "view-normal":
      setActiveSheetView("normal");
      break;
    case "view-page-layout":
      setActiveSheetView("page-layout");
      break;
    case "view-page-break":
      setActiveSheetView("page-break");
      break;
    case "view-custom-views":
      if (showCustomViewsUnavailableIfNeeded()) break;
      openCustomViewsDialog();
      break;
    case "view-full-screen":
      await toggleFullScreenView();
      break;
    case "view-sheet-view-new":
      createSheetView();
      break;
    case "view-sheet-view-keep":
      keepActiveSheetView();
      break;
    case "view-sheet-view-exit":
      exitActiveSheetView();
      break;
    case "view-sheet-view-options":
      openSheetViewOptionsMenu(trigger);
      break;
    case "view-ruler":
      toggleActiveSheetRuler();
      break;
    case "view-gridlines":
      toggleActiveSheetGridlines();
      break;
    case "view-headings":
      toggleActiveSheetHeadings();
      break;
    case "view-formula-bar":
      toggleFormulaBarVisibility();
      break;
    case "view-white-space":
      toggleActiveSheetWhiteSpace();
      break;
    case "view-zero-values":
      toggleActiveSheetZeroValues();
      break;
    case "view-objects":
      toggleActiveSheetObjectsVisibility();
      break;
    case "view-outline-symbols":
      toggleActiveSheetOutlineSymbols();
      break;
    case "view-right-to-left":
      toggleActiveSheetRightToLeft();
      break;
    case "view-horizontal-scrollbar":
      toggleWorkbookWindowShowOption("showHorizontalScroll", "横スクロールバー");
      break;
    case "view-vertical-scrollbar":
      toggleWorkbookWindowShowOption("showVerticalScroll", "縦スクロールバー");
      break;
    case "view-sheet-tabs":
      toggleWorkbookWindowShowOption("showSheetTabs", "シート見出し");
      break;
    case "view-zoom":
      openZoomDialog(trigger);
      break;
    case "view-100":
      setActiveSheetZoom(100);
      break;
    case "view-zoom-selection":
      zoomSelectionToFit();
      break;
    case "view-freeze-panes":
      toggleFreezePanes();
      break;
    case "view-freeze-menu":
      openViewFreezeMenu(trigger);
      break;
    case "view-split":
      toggleSplitPane();
      break;
    case "view-new-window":
      createWorkbookWindowView();
      break;
    case "view-arrange-all":
      openArrangeAllDialog();
      break;
    case "view-hide-window":
      hideWorkbookWindow();
      break;
    case "view-unhide-window":
      openUnhideWorkbookWindowDialog();
      break;
    case "view-side-by-side":
      toggleViewSideBySide();
      break;
    case "view-sync-scrolling":
      toggleSynchronousScrolling();
      break;
    case "view-reset-window-position":
      resetWindowPosition();
      break;
    case "view-switch-windows":
      openSwitchWindowsMenu(trigger);
      break;
    case "view-macros":
      openMacrosDialog();
      break;
    case "view-macros-menu":
      openViewMacrosMenu(trigger);
      break;
    case "draw-select":
      cancelInkDrawingMode();
      cancelShapeDrawing({ silent: true });
      setStatus("選択モードに戻りました。");
      break;
    case "draw-pen":
      beginInkDrawingMode("pen");
      break;
    case "draw-pencil":
      beginInkDrawingMode("pencil");
      break;
    case "draw-highlighter":
      beginInkDrawingMode("highlighter");
      break;
    case "draw-eraser":
      if (event?.target?.closest?.(".ribbon-inline-chevron, .ribbon-command-caret")) openRibbonInkEraserMenu(trigger);
      else beginInkEraserMode();
      break;
    case "draw-lasso":
      beginInkDrawingMode("lasso");
      break;
    case "draw-color-black":
    case "draw-color-blue":
    case "draw-color-red":
    case "draw-color-yellow":
      setInkColor(command.replace("draw-color-", ""));
      break;
    case "draw-thin":
      setInkWidth(1.5);
      break;
    case "draw-thick":
      setInkWidth(5);
      break;
    case "draw-convert-shape":
      convertSelectedInkToFreeformShape();
      break;
    case "draw-clear-ink":
      clearInkObjects();
      break;
    case "layout-themes":
      openPageLayoutThemeMenu(trigger);
      break;
    case "layout-colors":
      openPageLayoutColorMenu(trigger);
      break;
    case "layout-fonts":
      openPageLayoutFontMenu(trigger);
      break;
    case "layout-effects":
      openPageLayoutEffectsMenu(trigger);
      break;
    case "layout-margins":
      openPageLayoutMarginsMenu(trigger);
      break;
    case "layout-orientation":
      openPageLayoutOrientationMenu(trigger);
      break;
    case "layout-size":
      openPageLayoutSizeMenu(trigger);
      break;
    case "layout-print-area":
      openPageLayoutPrintAreaMenu(trigger);
      break;
    case "layout-breaks":
      openPageLayoutBreaksMenu(trigger);
      break;
    case "layout-background":
      setPageBackgroundFromSelection();
      break;
    case "layout-print-titles":
      setPrintTitlesFromSelection();
      break;
    case "layout-page-setup":
      openPageSetupDialog();
      break;
    case "layout-bring-front":
      reorderSelectedShape("forward");
      break;
    case "layout-send-back":
      reorderSelectedShape("backward");
      break;
    case "layout-selection-pane":
      openShapeSelectionPane();
      break;
    case "autosum-menu":
      openHomeAutoSumMenu(trigger);
      break;
    case "fill-menu":
      openHomeFillMenu(trigger);
      break;
    case "clear-menu":
      openHomeClearMenu(trigger);
      break;
    case "sort-filter":
      openHomeSortFilterMenu(trigger);
      break;
    case "find-select":
      openFindSelectMenu(trigger);
      break;
    case "number-currency":
      applyNumberFormatToSelection("¥#,##0", "通貨");
      break;
    case "number-percent":
      applyNumberFormatToSelection("0%", "パーセント");
      break;
    case "number-comma":
      applyNumberFormatToSelection("#,##0", "桁区切り");
      break;
    case "increase-decimal":
      adjustSelectionDecimalPlaces(1);
      break;
    case "decrease-decimal":
      adjustSelectionDecimalPlaces(-1);
      break;
    case "data-get-menu":
      openDataGetMenu(trigger);
      break;
    case "data-get-text":
      await chooseDelimitedTextForImport();
      break;
    case "data-get-web":
      openWebImportDialog();
      break;
    case "data-get-table-range":
      openGetDataFromTableRangeDialog();
      break;
    case "data-recent-sources":
      openRecentSourcesMenu(trigger);
      break;
    case "data-existing-connections":
      openExistingConnectionsDialog();
      break;
    case "data-queries-connections":
      openQueriesConnectionsPane();
      break;
    case "data-refresh-menu":
      openDataRefreshMenu(trigger);
      break;
    case "data-refresh-all":
      void refreshWorkbookData();
      break;
    case "data-sort-asc":
      sortSelectionByActiveColumn("asc");
      break;
    case "data-sort-desc":
      sortSelectionByActiveColumn("desc");
      break;
    case "data-sort-dialog":
      openSortDialog();
      break;
    case "data-filter":
      toggleAutoFilterForSelection();
      break;
    case "data-clear-filter":
      clearAutoFilterCriteria();
      break;
    case "data-reapply-filter":
      if (!reapplyAutoFilterForSheetGroup({ includeSort: true })) reapplyAutoFilter({ includeSort: true });
      break;
    case "data-advanced-filter":
      openAdvancedFilterDialog();
      break;
    case "data-text-to-columns":
      openTextToColumnsDialog();
      break;
    case "data-flash-fill":
      flashFillSelection();
      break;
    case "data-remove-duplicates":
      removeDuplicatesFromSelection();
      break;
    case "data-validation":
      openDataValidationMenu(trigger);
      break;
    case "data-consolidate":
      openConsolidateDialog();
      break;
    case "data-relationships":
      openRelationshipsDialog();
      break;
    case "data-manage-data-model":
      openDataModelDialog();
      break;
    case "data-what-if":
      openWhatIfMenu(trigger);
      break;
    case "data-forecast-sheet":
      openForecastSheetDialog();
      break;
    case "data-analysis":
      openDataAnalysisDialog();
      break;
    case "data-group":
      openOutlineCommandMenu(trigger, 1);
      break;
    case "data-ungroup":
      openOutlineCommandMenu(trigger, -1);
      break;
    case "data-show-detail":
      toggleOutlineDetailFromSelection(false);
      break;
    case "data-hide-detail":
      toggleOutlineDetailFromSelection(true);
      break;
    case "data-subtotal":
      openSubtotalDialog();
      break;
    default:
      setStatus(`"${command}" はリボンに追加済みです。今後さらに動作を拡張できます。`);
  }
}



function openWebVersionFromLocalFile() {
  const opened = window.open(CWS_WEB_APP_URL, "_blank", "noopener");
  if (!opened) {
    showMessageBox("Web版を新しいタブで開けませんでした。ブラウザーのポップアップブロック設定を確認してください。", {
      title: "Cht WebSheet",
    });
    setStatus("Web版を新しいタブで開けませんでした。");
    return false;
  }
  setStatus("Web版を新しいタブで開きました。");
  return true;
}



function insertTableFromSelection() {
  const ranges = activeSelectionRanges();
  const range = ranges[0] || activeSelectionRange();
  if (!range) {
    setStatus("テーブルに変換するセル範囲を選択してください。");
    return;
  }
  if (ranges.length > 1 || range.fullRow || range.fullColumn) {
    showMessageBox("複数の選択範囲または行/列全体からはテーブルを作成できません。", { title: "Cht WebSheet" });
    return;
  }

  const sheet = activeSheet();
  const tableRange = materializeSelectionRange(range);
  return insertTableFromRange(tableRange, {
    selectionBefore: tableRange ? { ...tableRange } : null,
    selectionBeforeRanges: ranges.map((item) => ({ ...item })),
    hasHeaders: tableRange ? inferHeaderRow(sheet, tableRange) : true,
  });
}



function openCreateTableDialog() {
  const context = selectedDataRange("テーブル", { expandCurrentRegion: true });
  if (!context) return;
  const { sheet, range } = context;
  if (context.ranges.length > 1 || range.fullRow || range.fullColumn) {
    showMessageBox("複数の選択範囲または行/列全体からはテーブルを作成できません。", { title: "Cht WebSheet" });
    return;
  }
  const tableRange = materializeSelectionRange(range);
  if (!tableRange) return;
  const cellCount = (tableRange.bottom - tableRange.top + 1) * (tableRange.right - tableRange.left + 1);
  if (cellCount > 50000) {
    showMessageBox("選択範囲が大きすぎるため、テーブルを作成できません。", { title: "Cht WebSheet" });
    return;
  }
  const hasHeaders = inferHeaderRow(sheet, tableRange);
  state.dataDialog = { type: "create-table", range: { ...tableRange } };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog create-table-dialog" role="dialog" aria-modal="true" aria-labelledby="createTableTitle" data-data-dialog="create-table">
      <div class="excel-mini-titlebar" id="createTableTitle">テーブルの作成</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="excel-dialog-field">
          <span class="excel-dialog-caption">${accessKeyLabelHtml("テーブルに変換するデータ範囲(W):")}</span>
          <input type="text" data-data-field="createTableRange" value="${escapeAttr(rangeToLabel(tableRange))}" />
        </label>
        <label class="data-dialog-check"><input type="checkbox" data-data-field="createTableHasHeaders"${hasHeaders ? " checked" : ""} />先頭行をテーブルの見出しとして使用する(M)</label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="create-table-ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("[data-data-field='createTableRange']")?.focus({ preventScroll: true });
  overlay.querySelector("[data-data-field='createTableRange']")?.select?.();
}



function applyCreateTableDialog(dialog) {
  const sheet = activeSheet();
  const rawRange = dialog?.querySelector("[data-data-field='createTableRange']")?.value || "";
  const range = parsePageSetupRangeLabel(rawRange, sheet);
  if (!range || range.fullRow || range.fullColumn) {
    showMessageBox("テーブルに変換するデータ範囲が正しくありません。", { title: "Cht WebSheet" });
    return;
  }
  const hasHeaders = Boolean(dialog?.querySelector("[data-data-field='createTableHasHeaders']")?.checked);
  const cellCount = (range.bottom - range.top + 1) * (range.right - range.left + 1);
  if (cellCount > 50000) {
    showMessageBox("選択範囲が大きすぎるため、テーブルを作成できません。", { title: "Cht WebSheet" });
    return;
  }
  closeDataDialog();
  insertTableFromRange(range, { hasHeaders });
}



function insertTableFromRange(tableRange, options = {}) {
  if (!tableRange) return;
  if (!options.skipSheetGroup && insertTableToSheetGroup(tableRange, options)) return;
  const sheet = activeSheet();
  const hasHeaders = options.hasHeaders !== false;
  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = options.selectionBefore || activeSelectionRange() || { ...tableRange };
  const selectionBeforeRanges = options.selectionBeforeRanges || activeSelectionRanges().map((item) => ({ ...item }));
  try {
    ensureSheetSize(sheet, tableRange.bottom, tableRange.right);
    applyExcelTableStyle(sheet, tableRange, { hasHeaders });
    sheet.tables = Array.isArray(sheet.tables) ? sheet.tables : [];
    sheet.tables.push({
      id: uniqueWorksheetObjectId("table"),
      name: nextTableName(sheet),
      range: { ...tableRange },
      showHeaderRow: true,
      showTotalsRow: false,
      showFilterButton: true,
      style: "TableStyleMedium2",
      styleInfo: parseTableStyleInfoXml(null),
      hasHeaders,
    });
    state.hf = buildFormulaEngine(state.model);
    state.selected = { sheetIndex: state.activeSheetIndex, row: tableRange.top, col: tableRange.left };
    state.selectionRange = tableRange;
    state.selectionRanges = [tableRange];
    state.selectionStart = state.selected;
    state.selectionDrag = null;
    state.selectedImage = null;
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetBefore,
      sheetAfter: captureSheetForHistory(sheet),
      selectionBefore,
      selectionAfter: tableRange,
      selectionBeforeRanges,
      selectionAfterRanges: [tableRange],
      label: `${rangeToLabel(tableRange)} テーブル`,
    });
    renderSheet();
    updateFormulaBarForSelection();
    setStatus(`${rangeToLabel(tableRange)} をテーブルとして書式設定しました。`);
  } catch (error) {
    console.error(error);
    restoreSheetFromHistory(sheet, sheetBefore);
    state.hf = buildFormulaEngine(state.model);
    renderSheet();
    showMessageBox("テーブルを作成できません。", { title: "Cht WebSheet" });
  }
}



function insertTableToSheetGroup(tableRange, options = {}) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !tableRange) return false;
  const hasHeaders = options.hasHeaders !== false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    const targetRange = materializeSelectionRangeForSheet(sheet, tableRange, sheetIndex);
    if (!targetRange || targetRange.fullRow || targetRange.fullColumn) continue;
    const cellCount = (targetRange.bottom - targetRange.top + 1) * (targetRange.right - targetRange.left + 1);
    if (cellCount > 50000) {
      showMessageBox("選択範囲が大きすぎるため、テーブルを作成できません。", { title: "Cht WebSheet" });
      return true;
    }
    if (!ensureEditableRangesForProtection(sheet, [targetRange], "テーブルを作成")) return true;
    items.push({ sheetIndex, sheet, range: targetRange });
  }
  if (!items.length) return true;

  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  const selectionBefore = options.selectionBefore || activeSelectionRange() || { ...tableRange };
  const selectionBeforeRanges = options.selectionBeforeRanges || activeSelectionRanges().map((item) => ({ ...item }));
  const activeItem = items.find((item) => item.sheetIndex === state.activeSheetIndex) || items[0];
  try {
    items.forEach((item) => {
      ensureSheetSize(item.sheet, item.range.bottom, item.range.right);
      applyExcelTableStyle(item.sheet, item.range, { hasHeaders });
      item.sheet.tables = Array.isArray(item.sheet.tables) ? item.sheet.tables : [];
      item.sheet.tables.push({
        id: uniqueWorksheetObjectId("table"),
        name: nextTableName(item.sheet),
        range: { ...item.range },
        showHeaderRow: true,
        showTotalsRow: false,
        showFilterButton: true,
        style: "TableStyleMedium2",
        styleInfo: parseTableStyleInfoXml(null),
        hasHeaders,
      });
    });
    state.hf = buildFormulaEngine(state.model);
    state.selected = { sheetIndex: state.activeSheetIndex, row: activeItem.range.top, col: activeItem.range.left };
    state.selectionRange = { ...activeItem.range, sheetIndex: state.activeSheetIndex };
    state.selectionRanges = [{ ...activeItem.range, sheetIndex: state.activeSheetIndex }];
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
      selectionAfter: activeItem.range,
      selectionBeforeRanges,
      selectionAfterRanges: [{ ...activeItem.range, sheetIndex: state.activeSheetIndex }],
      label: `${rangeToLabel(tableRange)} ${sheetIndexes.length} シートテーブル`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${rangeToLabel(activeItem.range)} をテーブルとして書式設定しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    showMessageBox("テーブルを作成できません。", { title: "Cht WebSheet" });
  }
  return true;
}



function applyExcelTableStyle(sheet, range, options = {}) {
  const sheetId = getSheetId(sheet.name);
  const border = { width: 1, style: "solid", color: "#9cc2e5", weight: 2 };
  const hasHeaders = options.hasHeaders !== false;
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      if (row === range.top) {
        const existing = sheet.cells[cellKey(row, col)];
        if (!hasHeaders || !existing || getCellRawInput(sheet, row, col) === "") {
          writeCellContent(sheet, sheetId, row, col, `Column${col - range.left + 1}`);
        }
      }
      const cell = ensureCellModel(sheet, row, col);
      cell.css.fontFamily = cell.css.fontFamily || sheet.defaultCellCss?.fontFamily || fontFamilyForExcel("Calibri");
      cell.css.textAlign = cell.css.textAlign || "left";
      cell.css.verticalAlign = cell.css.verticalAlign || "middle";
      cell.css.alignItems = cell.css.alignItems || "center";
      if (row === range.top) {
        cell.css.backgroundColor = "#4472c4";
        cell.css.color = "#ffffff";
        cell.css.fontWeight = "700";
        cell.hasFill = true;
        cell.hasFontColor = true;
      } else if ((row - range.top) % 2 === 0) {
        cell.css.backgroundColor = "#d9eaf7";
        cell.hasFill = true;
      }
      cell.borders = {
        top: { ...border },
        right: { ...border },
        bottom: { ...border },
        left: { ...border },
      };
    }
  }
  internCellsInRange(sheet, range);
}



function nextTableName(sheet) {
  const existing = new Set((sheet.tables || []).map((table) => String(table.name || "").toLowerCase()));
  let index = (sheet.tables || []).length + 1;
  while (existing.has(`table${index}`)) index += 1;
  return `Table${index}`;
}



function selectedDataRange(action, options = {}) {
  const ranges = activeSelectionRanges();
  let range = ranges[0] || activeSelectionRange();
  if (!range && state.model) {
    const point = { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
    state.selected = point;
    state.selectionRange = selectionRangeFromPoints(point, point);
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = point;
    state.selectionAnchor = point;
    range = state.selectionRange;
  }
  if (!range) {
    if (!options.quiet) showMessageBox(`${action}を実行するセル範囲を選択してください。`, { title: "Cht WebSheet" });
    return null;
  }
  if (!options.allowMultiple && ranges.length > 1) {
    if (!options.quiet) showMessageBox("複数の選択範囲に対してこの操作を実行することはできません。", { title: "Cht WebSheet" });
    return null;
  }
  let materialized = materializeSelectionRange(range);
  if (!materialized) return null;
  if (options.expandCurrentRegion && isSingleCellRange(materialized)) {
    const point = state.selected?.sheetIndex === state.activeSheetIndex
      ? state.selected
      : { sheetIndex: state.activeSheetIndex, row: materialized.top, col: materialized.left };
    materialized = currentRegionFromPoint(activeSheet(), point) || materialized;
  }
  const cellCount = (materialized.bottom - materialized.top + 1) * (materialized.right - materialized.left + 1);
  const limit = options.limit || DATA_OPERATION_CELL_LIMIT;
  if (cellCount > limit) {
    if (!options.quiet) showMessageBox("選択範囲が大きすぎるため、この操作を実行できません。", { title: "Cht WebSheet" });
    return null;
  }
  if (options.singleColumn && materialized.left !== materialized.right) {
    if (!options.quiet) showMessageBox("この操作では、1列だけを選択してください。", { title: "Cht WebSheet" });
    return null;
  }
  return { sheet: activeSheet(), range: materialized, ranges };
}



function commitDataOperation(sheet, sheetBefore, range, label, statusMessage, selectionAfter = range, historyExtras = {}) {
  commitDataOperationForSheet(state.activeSheetIndex, sheet, sheetBefore, range, label, statusMessage, selectionAfter, historyExtras);
}



function commitDataOperationForSheet(sheetIndex, sheet, sheetBefore, range, label, statusMessage, selectionAfter = range, historyExtras = {}) {
  state.hf = buildFormulaEngine(state.model);
  const targetSheetIndex = Number.isFinite(Number(sheetIndex)) ? Number(sheetIndex) : state.activeSheetIndex;
  state.activeSheetIndex = targetSheetIndex;
  state.selected = { sheetIndex: targetSheetIndex, row: selectionAfter.top, col: selectionAfter.left };
  state.selectionRange = { ...selectionAfter, sheetIndex: targetSheetIndex };
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  const historyEntry = {
    sheetIndex: targetSheetIndex,
    selectionBefore: range,
    selectionAfter: state.selectionRange,
    label,
  };
  refreshActiveSheetViewSnapshot(sheet);
  if (historyExtras.sheetsBefore?.length) {
    const indexes = historyExtras.sheetsBefore.map((snapshot) => snapshot.index);
    historyEntry.sheetsBefore = historyExtras.sheetsBefore;
    historyEntry.sheetsAfter = captureSheetsForHistory(indexes);
  } else {
    historyEntry.sheetBefore = sheetBefore;
    historyEntry.sheetAfter = captureSheetForHistory(sheet);
  }
  pushHistory(historyEntry);
  renderWorkbook();
  updateFormulaBarForSelection();
  setStatus(statusMessage);
}



function applyAutoSumToSelection(functionName = "SUM") {
  const sheet = activeSheet();
  const sheetId = getSheetId(sheet?.name);
  if (!sheet || sheetId == null) return;
  const formulaName = normalizeAutoSumFunctionName(functionName);
  recordFormulaFunctionUse(formulaName);
  const commandLabel = autoSumCommandLabel(formulaName);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = rawActiveSelectionRanges().map((range) => ({ ...range }));
  const ranges = activeSelectionRanges();
  if (applyAutoSumToSheetGroup(formulaName, commandLabel, ranges, selectionBefore, selectionBeforeRanges)) return;
  const operations = autoSumOperationsForRanges(sheet, sheetId, ranges, formulaName, state.activeSheetIndex);
  if (!operations.length) {
    showMessageBox(`${commandLabel}に使用する数値が見つかりません。`, { title: "Cht WebSheet" });
    return;
  }

  const targetRanges = operations.map((operation) => operation.target);
  const maxRow = Math.max(...targetRanges.map((range) => range.bottom));
  const maxCol = Math.max(...targetRanges.map((range) => range.right));
  const sheetBefore = captureSheetForHistory(sheet);
  ensureSheetSize(sheet, maxRow, maxCol);
  operations.forEach((operation) => {
    writeCellContent(sheet, sheetId, operation.target.top, operation.target.left, operation.formula);
  });
  internCellsInRanges(sheet, targetRanges);
  state.hf = buildFormulaEngine(state.model);

  const firstRange = targetRanges[0];
  state.selected = { sheetIndex: state.activeSheetIndex, row: firstRange.top, col: firstRange.left };
  state.selectionRange = { ...firstRange };
  state.selectionRanges = targetRanges.map((range) => ({ ...range }));
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore,
    selectionAfter: firstRange,
    selectionBeforeRanges,
    selectionAfterRanges: targetRanges,
    label: `${rangesToLabel(targetRanges)} ${commandLabel}`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`${rangesToLabel(targetRanges)} に${commandLabel}を入力しました。`);
}



function applyAutoSumToSheetGroup(formulaName, commandLabel, ranges, selectionBefore, selectionBeforeRanges) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const sheet = activeSheet();
  const sheetId = getSheetId(sheet?.name);
  const activeOperations = autoSumOperationsForRanges(sheet, sheetId, ranges, formulaName, state.activeSheetIndex);
  if (!activeOperations.length) {
    showMessageBox(`${commandLabel}に使用する数値が見つかりません。`, { title: "Cht WebSheet" });
    return true;
  }
  const activeTargetRanges = activeOperations.map((operation) => operation.target);
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const targetSheet = state.model?.sheets?.[sheetIndex];
    const targetSheetId = getSheetId(targetSheet?.name);
    if (!targetSheet || targetSheetId == null) continue;
    const operations = activeOperations.map((operation) => ({
      target: { ...operation.target, sheetIndex },
      formula: operation.formula,
    }));
    const targetRanges = operations.map((operation) => operation.target);
    if (!ensureEditableRangesForProtection(targetSheet, targetRanges, commandLabel)) return true;
    items.push({ sheetIndex, sheet: targetSheet, sheetId: targetSheetId, operations, targetRanges, ranges: targetRanges });
  }
  if (!items.length) return true;
  const affectedSheetIndexes = affectedSheetIndexesForGroupedRanges(items);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  try {
    items.forEach((item) => {
      const maxRow = Math.max(...item.targetRanges.map((range) => range.bottom));
      const maxCol = Math.max(...item.targetRanges.map((range) => range.right));
      ensureSheetSize(item.sheet, maxRow, maxCol);
      item.operations.forEach((operation) => {
        writeCellContent(item.sheet, item.sheetId, operation.target.top, operation.target.left, operation.formula);
      });
      internCellsInRanges(item.sheet, item.targetRanges);
    });
    refreshChartsForGroupedRanges(items);
    state.hf = buildFormulaEngine(state.model);
    const firstRange = activeTargetRanges[0];
    state.selected = { sheetIndex: state.activeSheetIndex, row: firstRange.top, col: firstRange.left };
    state.selectionRange = { ...firstRange };
    state.selectionRanges = activeTargetRanges.map((range) => ({ ...range }));
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
      selectionAfter: firstRange,
      selectionBeforeRanges,
      selectionAfterRanges: activeTargetRanges,
      label: `${rangesToLabel(activeTargetRanges)} ${sheetIndexes.length} シート${commandLabel}`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${rangesToLabel(activeTargetRanges)} に${commandLabel}を入力しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`${commandLabel}を入力できませんでした：${error.message || error}`);
  }
  return true;
}



function normalizeAutoSumFunctionName(functionName) {
  const normalized = String(functionName || "SUM").trim().toUpperCase();
  return ["SUM", "AVERAGE", "COUNT", "MAX", "MIN"].includes(normalized) ? normalized : "SUM";
}



function autoSumCommandLabel(functionName) {
  return {
    SUM: "オートSUM",
    AVERAGE: "平均",
    COUNT: "数値の個数",
    MAX: "最大値",
    MIN: "最小値",
  }[normalizeAutoSumFunctionName(functionName)];
}



function formulaFunctionDefinitions() {
  return [
    { name: "SUM", category: "数学/三角", label: "SUM", description: "選択した数値範囲の合計を返します。", template: (ref) => `=SUM(${ref})` },
    { name: "AVERAGE", category: "統計", label: "AVERAGE", description: "数値範囲の平均を返します。", template: (ref) => `=AVERAGE(${ref})` },
    { name: "COUNT", category: "統計", label: "COUNT", description: "数値が入力されているセルの個数を返します。", template: (ref) => `=COUNT(${ref})` },
    { name: "MAX", category: "統計", label: "MAX", description: "数値範囲の最大値を返します。", template: (ref) => `=MAX(${ref})` },
    { name: "MIN", category: "統計", label: "MIN", description: "数値範囲の最小値を返します。", template: (ref) => `=MIN(${ref})` },
    { name: "IF", category: "論理", label: "IF", description: "条件に応じて 2 つの値を切り替えます。", template: () => "=IF(TRUE,1,0)" },
    { name: "XLOOKUP", category: "検索/行列", label: "XLOOKUP", description: "範囲から検索値に一致する値を返します。", template: (ref) => `=XLOOKUP(${firstCellReferenceFromRangeLabel(ref)},${ref},${ref})` },
    { name: "TODAY", category: "日付/時刻", label: "TODAY", description: "現在の日付を返します。", template: () => "=TODAY()" },
    { name: "TEXTJOIN", category: "文字列", label: "TEXTJOIN", description: "複数の文字列を区切り文字で連結します。", template: (ref) => `=TEXTJOIN(",",TRUE,${ref})` },
    { name: "COUNTIF", category: "統計", label: "COUNTIF", description: "条件に一致するセルの個数を返します。", template: (ref) => `=COUNTIF(${ref},"<>")` },
  ];
}



function firstCellReferenceFromRangeLabel(ref) {
  const text = String(ref || "A1");
  return text.split(":")[0] || "A1";
}



function formulaFunctionByName(functionName) {
  const name = String(functionName || "").trim().toUpperCase();
  return formulaInsertFunctionDefinitionByName(name) || formulaFunctionDefinitions()[0];
}



function formulaAutocompleteExtraFunctionDefinitions() {
  const item = (name, category, syntax, description, priority = 40) => ({
    name,
    label: name,
    category,
    syntax,
    description,
    priority,
  });
  return [
    item("SUM", "数学/三角", "SUM(number1, [number2], ...)", "引数リストの数値を合計します。", 1),
    item("AVERAGE", "統計", "AVERAGE(number1, [number2], ...)", "引数リストの平均値を返します。", 2),
    item("COUNT", "統計", "COUNT(value1, [value2], ...)", "数値を含むセルの個数を返します。", 3),
    item("COUNTA", "統計", "COUNTA(value1, [value2], ...)", "空白ではないセルの個数を返します。", 4),
    item("MAX", "統計", "MAX(number1, [number2], ...)", "最大値を返します。", 5),
    item("MIN", "統計", "MIN(number1, [number2], ...)", "最小値を返します。", 6),
    item("IF", "論理", "IF(logical_test, value_if_true, [value_if_false])", "条件の結果に応じて値を切り替えます。", 7),
    item("IFS", "論理", "IFS(logical_test1, value_if_true1, ...)", "複数の条件を順に評価します。", 18),
    item("LET", "論理", "LET(name1, name_value1, calculation_or_name2, [name_value2], ...)", "数式内で名前を定義し、計算式で再利用します。", 19),
    item("AND", "論理", "AND(logical1, [logical2], ...)", "すべての引数が TRUE のとき TRUE を返します。", 16),
    item("OR", "論理", "OR(logical1, [logical2], ...)", "いずれかの引数が TRUE のとき TRUE を返します。", 17),
    item("NOT", "論理", "NOT(logical)", "論理値を反転します。", 45),
    item("IFERROR", "論理", "IFERROR(value, value_if_error)", "数式がエラーの場合に指定した値を返します。", 20),
    item("SWITCH", "論理", "SWITCH(expression, value1, result1, [default_or_value2], ...)", "式の値に一致する結果を返します。", 52),
    item("XLOOKUP", "検索/行列", "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])", "範囲または配列を検索して対応する項目を返します。", 8),
    item("VLOOKUP", "検索/行列", "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])", "表の左端列で値を検索します。", 9),
    item("HLOOKUP", "検索/行列", "HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])", "表の上端行で値を検索します。", 30),
    item("INDEX", "検索/行列", "INDEX(array, row_num, [column_num])", "範囲内の指定位置の値を返します。", 10),
    item("MATCH", "検索/行列", "MATCH(lookup_value, lookup_array, [match_type])", "検索値の相対位置を返します。", 11),
    item("XMATCH", "検索/行列", "XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])", "検索値の相対位置を返します。", 31),
    item("FILTER", "検索/行列", "FILTER(array, include, [if_empty])", "条件に一致するデータだけを返します。", 24),
    item("SORT", "検索/行列", "SORT(array, [sort_index], [sort_order], [by_col])", "範囲または配列を並べ替えます。", 25),
    item("SORTBY", "検索/行列", "SORTBY(array, by_array1, [sort_order1], ...)", "別の配列または範囲の値を基準に配列を並べ替えます。", 25),
    item("UNIQUE", "検索/行列", "UNIQUE(array, [by_col], [exactly_once])", "一意の値を返します。", 26),
    item("SEQUENCE", "検索/行列", "SEQUENCE(rows, [columns], [start], [step])", "連続した数値の配列を返します。", 37),
    item("TAKE", "検索/行列", "TAKE(array, rows, [columns])", "配列の先頭または末尾から指定した行/列を返します。", 58),
    item("DROP", "検索/行列", "DROP(array, rows, [columns])", "配列の先頭または末尾から指定した行/列を除外します。", 59),
    item("VSTACK", "検索/行列", "VSTACK(array1, [array2], ...)", "配列を縦方向に結合します。", 60),
    item("HSTACK", "検索/行列", "HSTACK(array1, [array2], ...)", "配列を横方向に結合します。", 61),
    item("LAMBDA", "論理", "LAMBDA(parameter1, [parameter2], ..., calculation)", "再利用できる計算式を定義します。", 62),
    item("MAKEARRAY", "検索/行列", "MAKEARRAY(rows, cols, lambda)", "指定した行数と列数の計算配列を返します。", 62),
    item("BYROW", "検索/行列", "BYROW(array, lambda)", "配列の各行に LAMBDA を適用します。", 62),
    item("BYCOL", "検索/行列", "BYCOL(array, lambda)", "配列の各列に LAMBDA を適用します。", 62),
    item("MAP", "検索/行列", "MAP(array1, lambda_or_array2, ...)", "配列の各値に LAMBDA を適用します。", 62),
    item("REDUCE", "検索/行列", "REDUCE([initial_value], array, lambda)", "配列を累積して 1 つの値を返します。", 62),
    item("SCAN", "検索/行列", "SCAN([initial_value], array, lambda)", "累積計算の途中結果を配列で返します。", 62),
    item("CHOOSEROWS", "検索/行列", "CHOOSEROWS(array, row_num1, [row_num2], ...)", "指定した行だけを配列から返します。", 62),
    item("CHOOSECOLS", "検索/行列", "CHOOSECOLS(array, col_num1, [col_num2], ...)", "指定した列だけを配列から返します。", 63),
    item("TOCOL", "検索/行列", "TOCOL(array, [ignore], [scan_by_column])", "配列を 1 列に変換します。", 64),
    item("TOROW", "検索/行列", "TOROW(array, [ignore], [scan_by_column])", "配列を 1 行に変換します。", 65),
    item("WRAPROWS", "検索/行列", "WRAPROWS(vector, wrap_count, [pad_with])", "ベクトルを指定列数ごとに行へ折り返します。", 66),
    item("WRAPCOLS", "検索/行列", "WRAPCOLS(vector, wrap_count, [pad_with])", "ベクトルを指定行数ごとに列へ折り返します。", 67),
    item("EXPAND", "検索/行列", "EXPAND(array, rows, [columns], [pad_with])", "配列を指定サイズに拡張します。", 68),
    item("CHOOSE", "検索/行列", "CHOOSE(index_num, value1, [value2], ...)", "番号に基づいて値を選択します。", 48),
    item("INDIRECT", "検索/行列", "INDIRECT(ref_text, [a1])", "文字列で指定された参照を返します。", 38),
    item("OFFSET", "検索/行列", "OFFSET(reference, rows, cols, [height], [width])", "基準セルから指定位置の参照を返します。", 39),
    item("ROW", "検索/行列", "ROW([reference])", "参照の行番号を返します。", 55),
    item("COLUMN", "検索/行列", "COLUMN([reference])", "参照の列番号を返します。", 56),
    item("TEXTJOIN", "文字列", "TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)", "区切り文字を使って文字列を結合します。", 12),
    item("CONCAT", "文字列", "CONCAT(text1, [text2], ...)", "複数の文字列を結合します。", 22),
    item("TEXTSPLIT", "文字列", "TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])", "文字列を列または行に分割して配列として返します。", 18),
    item("TEXTBEFORE", "文字列", "TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])", "区切り文字の前にある文字列を返します。", 19),
    item("TEXTAFTER", "文字列", "TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])", "区切り文字の後にある文字列を返します。", 20),
    item("LEFT", "文字列", "LEFT(text, [num_chars])", "文字列の左端から指定文字数を返します。", 27),
    item("RIGHT", "文字列", "RIGHT(text, [num_chars])", "文字列の右端から指定文字数を返します。", 28),
    item("MID", "文字列", "MID(text, start_num, num_chars)", "文字列の指定位置から文字を返します。", 29),
    item("LEN", "文字列", "LEN(text)", "文字列の文字数を返します。", 32),
    item("FIND", "文字列", "FIND(find_text, within_text, [start_num])", "文字列内で別の文字列を検索します。", 54),
    item("SEARCH", "文字列", "SEARCH(find_text, within_text, [start_num])", "大文字小文字を区別せず文字列を検索します。", 57),
    item("SUBSTITUTE", "文字列", "SUBSTITUTE(text, old_text, new_text, [instance_num])", "文字列中の指定文字を置換します。", 33),
    item("TEXT", "文字列", "TEXT(value, format_text)", "値を指定した表示形式の文字列に変換します。", 23),
    item("VALUE", "文字列", "VALUE(text)", "数値を表す文字列を数値に変換します。", 46),
    item("TRIM", "文字列", "TRIM(text)", "余分なスペースを削除します。", 47),
    item("UPPER", "文字列", "UPPER(text)", "文字列を大文字に変換します。", 58),
    item("LOWER", "文字列", "LOWER(text)", "文字列を小文字に変換します。", 59),
    item("ROUND", "数学/三角", "ROUND(number, num_digits)", "数値を指定した桁数に丸めます。", 13),
    item("ROUNDUP", "数学/三角", "ROUNDUP(number, num_digits)", "数値を切り上げます。", 41),
    item("ROUNDDOWN", "数学/三角", "ROUNDDOWN(number, num_digits)", "数値を切り下げます。", 42),
    item("INT", "数学/三角", "INT(number)", "数値を最も近い整数に切り下げます。", 43),
    item("ABS", "数学/三角", "ABS(number)", "数値の絶対値を返します。", 44),
    item("MOD", "数学/三角", "MOD(number, divisor)", "除算の余りを返します。", 49),
    item("PRODUCT", "数学/三角", "PRODUCT(number1, [number2], ...)", "引数の積を返します。", 50),
    item("SUMIF", "数学/三角", "SUMIF(range, criteria, [sum_range])", "条件に一致するセルを合計します。", 14),
    item("SUMIFS", "数学/三角", "SUMIFS(sum_range, criteria_range1, criteria1, ...)", "複数条件に一致するセルを合計します。", 15),
    item("COUNTIF", "統計", "COUNTIF(range, criteria)", "条件に一致するセルの個数を返します。", 19),
    item("COUNTIFS", "統計", "COUNTIFS(criteria_range1, criteria1, ...)", "複数条件に一致するセルの個数を返します。", 21),
    item("AVERAGEIF", "統計", "AVERAGEIF(range, criteria, [average_range])", "条件に一致するセルの平均値を返します。", 34),
    item("AVERAGEIFS", "統計", "AVERAGEIFS(average_range, criteria_range1, criteria1, ...)", "複数条件に一致するセルの平均値を返します。", 35),
    item("SUBTOTAL", "数学/三角", "SUBTOTAL(function_num, ref1, [ref2], ...)", "リストまたはデータベースの集計値を返します。", 36),
    item("AGGREGATE", "数学/三角", "AGGREGATE(function_num, options, ref1, [ref2], ...)", "非表示行、エラー値、入れ子の集計を指定して集計値を返します。", 37),
    item("TODAY", "日付/時刻", "TODAY()", "現在の日付を返します。", 37),
    item("NOW", "日付/時刻", "NOW()", "現在の日付と時刻を返します。", 40),
    item("DATE", "日付/時刻", "DATE(year, month, day)", "指定した日付を表すシリアル値を返します。", 51),
    item("TIME", "日付/時刻", "TIME(hour, minute, second)", "指定した時刻を表すシリアル値を返します。", 53),
    item("YEAR", "日付/時刻", "YEAR(serial_number)", "日付から年を返します。", 60),
    item("MONTH", "日付/時刻", "MONTH(serial_number)", "日付から月を返します。", 61),
    item("DAY", "日付/時刻", "DAY(serial_number)", "日付から日を返します。", 62),
    item("DATEDIF", "日付/時刻", 'DATEDIF(start_date, end_date, "unit")', "2つの日付の差を返します。", 63),
    item("EOMONTH", "日付/時刻", "EOMONTH(start_date, months)", "指定月数後の月末日を返します。", 64),
    item("NETWORKDAYS", "日付/時刻", "NETWORKDAYS(start_date, end_date, [holidays])", "稼働日の日数を返します。", 65),
    item("STDEV.S", "統計", "STDEV.S(number1, [number2], ...)", "標本に基づく標準偏差を返します。", 66),
    item("STDEV.P", "統計", "STDEV.P(number1, [number2], ...)", "母集団全体に基づく標準偏差を返します。", 67),
    item("VAR.S", "統計", "VAR.S(number1, [number2], ...)", "標本に基づく分散を返します。", 68),
    item("VAR.P", "統計", "VAR.P(number1, [number2], ...)", "母集団全体に基づく分散を返します。", 69),
    item("MEDIAN", "統計", "MEDIAN(number1, [number2], ...)", "中央値を返します。", 70),
    item("MODE.SNGL", "統計", "MODE.SNGL(number1, [number2], ...)", "最頻値を返します。", 71),
    item("MODE.MULT", "統計", "MODE.MULT(number1, [number2], ...)", "複数の最頻値を返します。", 71.2),
    item("KURT", "統計", "KURT(number1, [number2], ...)", "尖度を返します。", 71.4),
    item("LARGE", "統計", "LARGE(array, k)", "データ内で k 番目に大きい値を返します。", 72),
    item("SMALL", "統計", "SMALL(array, k)", "データ内で k 番目に小さい値を返します。", 73),
    item("RANK.EQ", "統計", "RANK.EQ(number, ref, [order])", "数値の順位を返します。", 74),
    item("RANK.AVG", "統計", "RANK.AVG(number, ref, [order])", "同順位を平均順位で返します。", 74.2),
    item("PERCENTILE.INC", "統計", "PERCENTILE.INC(array, k)", "百分位数を返します。", 75),
    item("PERCENTRANK.INC", "統計", "PERCENTRANK.INC(array, x, [significance])", "百分率の順位を返します。", 75.2),
    item("PERCENTRANK.EXC", "統計", "PERCENTRANK.EXC(array, x, [significance])", "0 と 1 を除いた百分率の順位を返します。", 75.4),
    item("PROB", "統計", "PROB(x_range, prob_range, lower_limit, [upper_limit])", "指定範囲内の値になる確率を返します。", 75.5),
    item("FREQUENCY", "統計", "FREQUENCY(data_array, bins_array)", "頻度分布を返します。", 75.6),
    item("CORREL", "統計", "CORREL(array1, array2)", "2つの範囲の相関係数を返します。", 76),
    item("COVARIANCE.S", "統計", "COVARIANCE.S(array1, array2)", "標本共分散を返します。", 77),
    item("COVARIANCE.P", "統計", "COVARIANCE.P(array1, array2)", "母集団共分散を返します。", 77.1),
    item("STDEV", "統計", "STDEV(number1, [number2], ...)", "標本に基づく標準偏差を返します。", 77.2),
    item("STDEVP", "統計", "STDEVP(number1, [number2], ...)", "母集団全体に基づく標準偏差を返します。", 77.3),
    item("STDEVA", "統計", "STDEVA(value1, [value2], ...)", "文字列と論理値も含めて標本標準偏差を返します。", 77.35),
    item("STDEVPA", "統計", "STDEVPA(value1, [value2], ...)", "文字列と論理値も含めて母集団標準偏差を返します。", 77.36),
    item("VAR", "統計", "VAR(number1, [number2], ...)", "標本に基づく分散を返します。", 77.4),
    item("VARP", "統計", "VARP(number1, [number2], ...)", "母集団全体に基づく分散を返します。", 77.5),
    item("VARA", "統計", "VARA(value1, [value2], ...)", "文字列と論理値も含めて標本分散を返します。", 77.55),
    item("VARPA", "統計", "VARPA(value1, [value2], ...)", "文字列と論理値も含めて母集団分散を返します。", 77.56),
    item("COVAR", "統計", "COVAR(array1, array2)", "母集団共分散を返します。", 77.6),
    item("RAND", "数学/三角", "RAND()", "0以上1未満の乱数を返します。", 78),
    item("RANDBETWEEN", "数学/三角", "RANDBETWEEN(bottom, top)", "指定範囲内の整数の乱数を返します。", 79),
    item("PMT", "財務", "PMT(rate, nper, pv, [fv], [type])", "ローンの定期支払額を返します。", 80),
    item("NPV", "財務", "NPV(rate, value1, [value2], ...)", "投資の正味現在価値を返します。", 81),
    item("IRR", "財務", "IRR(values, [guess])", "一連のキャッシュフローの内部利益率を返します。", 82),
    item("HYPERLINK", "検索/行列", "HYPERLINK(link_location, [friendly_name])", "リンクを作成します。", 83),
    ...additionalFormulaCatalogDefinitions(item),
  ];
}



function additionalFormulaCatalogDefinitions(item) {
  return [
    item("AVERAGEA", "統計", "AVERAGEA(value1, [value2], ...)", "文字列や論理値も含めて平均を返します。", 84),
    item("COUNTBLANK", "統計", "COUNTBLANK(range)", "範囲内の空白セルの個数を返します。", 85),
    item("MAXA", "統計", "MAXA(value1, [value2], ...)", "文字列や論理値も含めて最大値を返します。", 86),
    item("MINA", "統計", "MINA(value1, [value2], ...)", "文字列や論理値も含めて最小値を返します。", 87),
    item("MAXIFS", "統計", "MAXIFS(max_range, criteria_range1, criteria1, ...)", "複数条件に一致する最大値を返します。", 88),
    item("MINIFS", "統計", "MINIFS(min_range, criteria_range1, criteria1, ...)", "複数条件に一致する最小値を返します。", 89),
    item("QUARTILE.INC", "統計", "QUARTILE.INC(array, quart)", "データの四分位数を返します。", 90),
    item("QUARTILE.EXC", "統計", "QUARTILE.EXC(array, quart)", "0 と 1 を除いた四分位数を返します。", 91),
    item("PERCENTILE.EXC", "統計", "PERCENTILE.EXC(array, k)", "0 と 1 を除いた百分位数を返します。", 92),
    item("NORM.DIST", "統計", "NORM.DIST(x, mean, standard_dev, cumulative)", "正規分布の値を返します。", 93),
    item("NORM.INV", "統計", "NORM.INV(probability, mean, standard_dev)", "正規分布の逆関数の値を返します。", 94),
    item("T.TEST", "統計", "T.TEST(array1, array2, tails, type)", "t 検定の確率を返します。", 95),
    item("CHISQ.TEST", "統計", "CHISQ.TEST(actual_range, expected_range)", "カイ二乗検定の値を返します。", 96),
    item("SLOPE", "統計", "SLOPE(known_y's, known_x's)", "線形回帰直線の傾きを返します。", 97),
    item("LINEST", "統計", "LINEST(known_y's, [known_x's], [const], [stats])", "線形回帰の係数と統計情報を配列で返します。", 97.2),
    item("LOGEST", "統計", "LOGEST(known_y's, [known_x's], [const], [stats])", "指数回帰曲線の係数と統計情報を配列で返します。", 97.4),
    item("INTERCEPT", "統計", "INTERCEPT(known_y's, known_x's)", "線形回帰直線の切片を返します。", 98),
    item("FORECAST.LINEAR", "統計", "FORECAST.LINEAR(x, known_y's, known_x's)", "線形予測値を返します。", 99),
    item("TREND", "統計", "TREND(known_y's, [known_x's], [new_x's], [const])", "線形近似による予測値を配列で返します。", 99.1),
    item("GROWTH", "統計", "GROWTH(known_y's, [known_x's], [new_x's], [const])", "指数近似による予測値を配列で返します。", 99.15),
    item("FORECAST", "互換性", "FORECAST(x, known_y's, known_x's)", "線形予測値を返す互換性関数です。", 99.2),
    item("LOOKUP", "検索/行列", "LOOKUP(lookup_value, lookup_vector, [result_vector])", "1 行または 1 列の範囲から値を検索します。", 100),
    item("ROWS", "検索/行列", "ROWS(array)", "参照または配列の行数を返します。", 101),
    item("COLUMNS", "検索/行列", "COLUMNS(array)", "参照または配列の列数を返します。", 102),
    item("SHEET", "検索/行列", "SHEET([value])", "参照先のシート番号を返します。", 103),
    item("SHEETS", "検索/行列", "SHEETS([reference])", "参照に含まれるシート数を返します。", 104),
    item("MMULT", "検索/行列", "MMULT(array1, array2)", "2 つの配列の行列積を返します。", 105),
    item("MDETERM", "検索/行列", "MDETERM(array)", "配列の行列式を返します。", 105.1),
    item("MINVERSE", "検索/行列", "MINVERSE(array)", "配列の逆行列を返します。", 105.2),
    item("MUNIT", "検索/行列", "MUNIT(dimension)", "指定した次元の単位行列を返します。", 105.3),
    item("HOUR", "日付/時刻", "HOUR(serial_number)", "時刻から時を返します。", 106),
    item("MINUTE", "日付/時刻", "MINUTE(serial_number)", "時刻から分を返します。", 107),
    item("SECOND", "日付/時刻", "SECOND(serial_number)", "時刻から秒を返します。", 108),
    item("WEEKDAY", "日付/時刻", "WEEKDAY(serial_number, [return_type])", "日付に対応する曜日番号を返します。", 109),
    item("WEEKNUM", "日付/時刻", "WEEKNUM(serial_number, [return_type])", "日付が年内の第何週かを返します。", 110),
    item("ISOWEEKNUM", "日付/時刻", "ISOWEEKNUM(date)", "ISO 週番号を返します。", 110.5),
    item("EDATE", "日付/時刻", "EDATE(start_date, months)", "指定月数だけ前後した日付を返します。", 111),
    item("WORKDAY", "日付/時刻", "WORKDAY(start_date, days, [holidays])", "指定した稼働日数後の日付を返します。", 112),
    item("WORKDAY.INTL", "日付/時刻", "WORKDAY.INTL(start_date, days, [weekend], [holidays])", "週末の曜日を指定して稼働日後の日付を返します。", 113),
    item("NETWORKDAYS.INTL", "日付/時刻", "NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])", "週末の曜日を指定して稼働日数を返します。", 114),
    item("DAYS", "日付/時刻", "DAYS(end_date, start_date)", "2 つの日付の間の日数を返します。", 115),
    item("DAYS360", "日付/時刻", "DAYS360(start_date, end_date, [method])", "1 年を 360 日として日数を返します。", 116),
    item("DATEVALUE", "日付/時刻", "DATEVALUE(date_text)", "日付文字列をシリアル値に変換します。", 117),
    item("TIMEVALUE", "日付/時刻", "TIMEVALUE(time_text)", "時刻文字列をシリアル値に変換します。", 118),
    item("YEARFRAC", "日付/時刻", "YEARFRAC(start_date, end_date, [basis])", "2 つの日付間の年割合を返します。", 119),
    item("CONCATENATE", "文字列", "CONCATENATE(text1, [text2], ...)", "複数の文字列を連結します。", 120),
    item("REPLACE", "文字列", "REPLACE(old_text, start_num, num_chars, new_text)", "指定位置の文字列を置換します。", 121),
    item("PROPER", "文字列", "PROPER(text)", "英単語の先頭文字を大文字にします。", 122),
    item("EXACT", "文字列", "EXACT(text1, text2)", "2 つの文字列が完全に同じか判定します。", 123),
    item("CHAR", "文字列", "CHAR(number)", "文字コードに対応する文字を返します。", 124),
    item("CODE", "文字列", "CODE(text)", "文字列の先頭文字のコードを返します。", 125),
    item("CLEAN", "文字列", "CLEAN(text)", "印刷できない文字を削除します。", 126),
    item("REPT", "文字列", "REPT(text, number_times)", "文字列を指定回数繰り返します。", 127),
    item("DOLLAR", "文字列", "DOLLAR(number, [decimals])", "数値を通貨形式の文字列に変換します。", 128),
    item("FIXED", "文字列", "FIXED(number, [decimals], [no_commas])", "数値を固定小数点形式の文字列に変換します。", 129),
    item("NUMBERVALUE", "文字列", "NUMBERVALUE(text, [decimal_separator], [group_separator])", "指定した区切り記号で文字列を数値に変換します。", 130),
    item("CEILING", "数学/三角", "CEILING(number, significance)", "数値を基準値の倍数に切り上げます。", 131),
    item("CEILING.MATH", "数学/三角", "CEILING.MATH(number, [significance], [mode])", "数値を基準値の倍数に切り上げます。", 132),
    item("CEILING.PRECISE", "数学/三角", "CEILING.PRECISE(number, [significance])", "正負にかかわらず基準値の倍数に切り上げます。", 133),
    item("FLOOR", "数学/三角", "FLOOR(number, significance)", "数値を基準値の倍数に切り下げます。", 134),
    item("FLOOR.MATH", "数学/三角", "FLOOR.MATH(number, [significance], [mode])", "数値を基準値の倍数に切り下げます。", 135),
    item("FLOOR.PRECISE", "数学/三角", "FLOOR.PRECISE(number, [significance])", "正負にかかわらず基準値の倍数に切り下げます。", 136),
    item("MROUND", "数学/三角", "MROUND(number, multiple)", "数値を指定した倍数に丸めます。", 137),
    item("TRUNC", "数学/三角", "TRUNC(number, [num_digits])", "数値の小数部を切り捨てます。", 138),
    item("QUOTIENT", "数学/三角", "QUOTIENT(numerator, denominator)", "除算の整数部分を返します。", 139),
    item("SIGN", "数学/三角", "SIGN(number)", "数値の符号を返します。", 140),
    item("PI", "数学/三角", "PI()", "円周率 π を返します。", 141),
    item("SIN", "数学/三角", "SIN(number)", "角度のサインを返します。", 142),
    item("COS", "数学/三角", "COS(number)", "角度のコサインを返します。", 143),
    item("TAN", "数学/三角", "TAN(number)", "角度のタンジェントを返します。", 144),
    item("ASIN", "数学/三角", "ASIN(number)", "アークサインを返します。", 145),
    item("ACOS", "数学/三角", "ACOS(number)", "アークコサインを返します。", 146),
    item("ATAN", "数学/三角", "ATAN(number)", "アークタンジェントを返します。", 147),
    item("DEGREES", "数学/三角", "DEGREES(angle)", "ラジアンを度に変換します。", 148),
    item("RADIANS", "数学/三角", "RADIANS(angle)", "度をラジアンに変換します。", 149),
    item("LN", "数学/三角", "LN(number)", "自然対数を返します。", 150),
    item("LOG", "数学/三角", "LOG(number, [base])", "指定した底の対数を返します。", 151),
    item("EXP", "数学/三角", "EXP(number)", "e のべき乗を返します。", 152),
    item("FACT", "数学/三角", "FACT(number)", "数値の階乗を返します。", 153),
    item("COMBIN", "数学/三角", "COMBIN(number, number_chosen)", "組み合わせの数を返します。", 154),
    item("PV", "財務", "PV(rate, nper, pmt, [fv], [type])", "投資の現在価値を返します。", 155),
    item("FV", "財務", "FV(rate, nper, pmt, [pv], [type])", "投資の将来価値を返します。", 156),
    item("RATE", "財務", "RATE(nper, pmt, pv, [fv], [type], [guess])", "ローンまたは投資の利率を返します。", 157),
    item("NPER", "財務", "NPER(rate, pmt, pv, [fv], [type])", "投資期間を返します。", 158),
    item("IPMT", "財務", "IPMT(rate, per, nper, pv, [fv], [type])", "指定期間の利息支払額を返します。", 159),
    item("PPMT", "財務", "PPMT(rate, per, nper, pv, [fv], [type])", "指定期間の元金支払額を返します。", 160),
    item("MIRR", "財務", "MIRR(values, finance_rate, reinvest_rate)", "修正内部利益率を返します。", 161),
    item("DB", "財務", "DB(cost, salvage, life, period, [month])", "定率法による減価償却費を返します。", 162),
    item("DDB", "財務", "DDB(cost, salvage, life, period, [factor])", "倍額定率法による減価償却費を返します。", 163),
    item("VDB", "財務", "VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])", "可変定率法による指定期間の減価償却費を返します。", 163.2),
    item("AMORLINC", "財務", "AMORLINC(cost, date_purchased, first_period, salvage, period, rate, [basis])", "各会計期の減価償却費を返します。", 163.4),
    item("AMORDEGRC", "財務", "AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, [basis])", "係数を使用して各会計期の減価償却費を返します。", 163.6),
    item("SLN", "財務", "SLN(cost, salvage, life)", "定額法による減価償却費を返します。", 164),
    item("SYD", "財務", "SYD(cost, salvage, life, per)", "級数法による減価償却費を返します。", 165),
    item("ISBLANK", "情報", "ISBLANK(value)", "参照が空白かどうかを返します。", 166),
    item("ISERROR", "情報", "ISERROR(value)", "値がエラーかどうかを返します。", 167),
    item("ISNUMBER", "情報", "ISNUMBER(value)", "値が数値かどうかを返します。", 168),
    item("ISTEXT", "情報", "ISTEXT(value)", "値が文字列かどうかを返します。", 169),
    item("ISNA", "情報", "ISNA(value)", "値が #N/A かどうかを返します。", 170),
    item("N", "情報", "N(value)", "値を数値に変換します。", 171),
    item("ERROR.TYPE", "情報", "ERROR.TYPE(error_val)", "エラー値に対応する番号を返します。", 172),
    item("IMABS", "エンジニアリング", "IMABS(inumber)", "複素数の絶対値を返します。", 173),
    item("IMSUM", "エンジニアリング", "IMSUM(inumber1, [inumber2], ...)", "複素数の和を返します。", 174),
    item("IMPRODUCT", "エンジニアリング", "IMPRODUCT(inumber1, [inumber2], ...)", "複素数の積を返します。", 175),
    item("IMREAL", "エンジニアリング", "IMREAL(inumber)", "複素数の実数係数を返します。", 176),
    item("IMAGINARY", "エンジニアリング", "IMAGINARY(inumber)", "複素数の虚数係数を返します。", 177),
    item("COMPLEX", "エンジニアリング", "COMPLEX(real_num, i_num, [suffix])", "実数係数と虚数係数から複素数を作成します。", 178),
    item("CONVERT", "エンジニアリング", "CONVERT(number, from_unit, to_unit)", "数値を別の単位に変換します。", 179),
    item("DEC2HEX", "エンジニアリング", "DEC2HEX(number, [places])", "10 進数を 16 進数に変換します。", 180),
    item("HEX2DEC", "エンジニアリング", "HEX2DEC(number)", "16 進数を 10 進数に変換します。", 181),
    item("DSUM", "データベース", "DSUM(database, field, criteria)", "条件に一致するデータベース列の合計を返します。", 182),
    item("DCOUNT", "データベース", "DCOUNT(database, field, criteria)", "条件に一致する数値セルの個数を返します。", 183),
    item("DCOUNTA", "データベース", "DCOUNTA(database, field, criteria)", "条件に一致する空白でないセルの個数を返します。", 184),
    item("DAVERAGE", "データベース", "DAVERAGE(database, field, criteria)", "条件に一致するデータベース列の平均を返します。", 185),
    item("DMAX", "データベース", "DMAX(database, field, criteria)", "条件に一致するデータベース列の最大値を返します。", 186),
    item("DMIN", "データベース", "DMIN(database, field, criteria)", "条件に一致するデータベース列の最小値を返します。", 187),
    item("DPRODUCT", "データベース", "DPRODUCT(database, field, criteria)", "条件に一致するデータベース列の積を返します。", 188),
    item("DSTDEV", "データベース", "DSTDEV(database, field, criteria)", "条件に一致するデータベース列の標本標準偏差を返します。", 189),
    item("DSTDEVP", "データベース", "DSTDEVP(database, field, criteria)", "条件に一致するデータベース列の母集団標準偏差を返します。", 190),
    item("DVAR", "データベース", "DVAR(database, field, criteria)", "条件に一致するデータベース列の標本分散を返します。", 191),
    item("DVARP", "データベース", "DVARP(database, field, criteria)", "条件に一致するデータベース列の母集団分散を返します。", 192),
    item("FVSCHEDULE", "財務", "FVSCHEDULE(principal, schedule)", "複利計算後の将来価値を返します。", 188),
    item("CUMIPMT", "財務", "CUMIPMT(rate, nper, pv, start_period, end_period, type)", "指定期間に支払う利息の累計額を返します。", 189),
    item("CUMPRINC", "財務", "CUMPRINC(rate, nper, pv, start_period, end_period, type)", "指定期間に支払う元金の累計額を返します。", 190),
    item("EFFECT", "財務", "EFFECT(nominal_rate, npery)", "実効年利率を返します。", 191),
    item("NOMINAL", "財務", "NOMINAL(effect_rate, npery)", "名目年利率を返します。", 192),
    item("XNPV", "財務", "XNPV(rate, values, dates)", "日付付きキャッシュフローの正味現在価値を返します。", 193),
    item("XIRR", "財務", "XIRR(values, dates, [guess])", "日付付きキャッシュフローの内部利益率を返します。", 193.2),
    item("DISC", "財務", "DISC(settlement, maturity, pr, redemption, [basis])", "割引債の割引率を返します。", 193.4),
    item("PRICEDISC", "財務", "PRICEDISC(settlement, maturity, discount, redemption, [basis])", "割引債の額面 100 あたりの価格を返します。", 193.6),
    item("YIELDDISC", "財務", "YIELDDISC(settlement, maturity, pr, redemption, [basis])", "割引債の利回りを返します。", 193.8),
    item("INTRATE", "財務", "INTRATE(settlement, maturity, investment, redemption, [basis])", "満期一括払い証券の利率を返します。", 193.9),
    item("RECEIVED", "財務", "RECEIVED(settlement, maturity, investment, discount, [basis])", "満期時に受け取る金額を返します。", 193.95),
    item("ACCRINT", "財務", "ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method])", "定期利付証券の未収利息を返します。", 193.951),
    item("ACCRINTM", "財務", "ACCRINTM(issue, settlement, rate, par, [basis])", "満期一括払い証券の未収利息を返します。", 193.952),
    item("PRICEMAT", "財務", "PRICEMAT(settlement, maturity, issue, rate, yld, [basis])", "満期に利息が支払われる証券の額面 100 あたりの価格を返します。", 193.953),
    item("YIELDMAT", "財務", "YIELDMAT(settlement, maturity, issue, rate, pr, [basis])", "満期に利息が支払われる証券の利回りを返します。", 193.954),
    item("COUPDAYBS", "財務", "COUPDAYBS(settlement, maturity, frequency, [basis])", "利払期間の開始日から受渡日までの日数を返します。", 193.961),
    item("COUPDAYS", "財務", "COUPDAYS(settlement, maturity, frequency, [basis])", "受渡日を含む利払期間の日数を返します。", 193.962),
    item("COUPDAYSNC", "財務", "COUPDAYSNC(settlement, maturity, frequency, [basis])", "受渡日から次の利払日までの日数を返します。", 193.963),
    item("COUPNCD", "財務", "COUPNCD(settlement, maturity, frequency, [basis])", "受渡日後の次の利払日を返します。", 193.964),
    item("COUPPCD", "財務", "COUPPCD(settlement, maturity, frequency, [basis])", "受渡日前の前回利払日を返します。", 193.965),
    item("COUPNUM", "財務", "COUPNUM(settlement, maturity, frequency, [basis])", "受渡日から満期日までの利払回数を返します。", 193.966),
    item("PRICE", "財務", "PRICE(settlement, maturity, rate, yld, redemption, frequency, [basis])", "定期利付債の額面 100 あたりの価格を返します。", 193.967),
    item("YIELD", "財務", "YIELD(settlement, maturity, rate, pr, redemption, frequency, [basis])", "定期利付債の利回りを返します。", 193.968),
    item("DURATION", "財務", "DURATION(settlement, maturity, coupon, yld, frequency, [basis])", "定期利付債のマコーレー デュレーションを返します。", 193.969),
    item("MDURATION", "財務", "MDURATION(settlement, maturity, coupon, yld, frequency, [basis])", "修正マコーレー デュレーションを返します。", 193.97),
    item("DOLLARDE", "財務", "DOLLARDE(fractional_dollar, fraction)", "分数表示のドル価格を小数に変換します。", 194),
    item("DOLLARFR", "財務", "DOLLARFR(decimal_dollar, fraction)", "小数表示のドル価格を分数に変換します。", 195),
    item("RRI", "財務", "RRI(nper, pv, fv)", "投資成長率を返します。", 196),
    item("PDURATION", "財務", "PDURATION(rate, pv, fv)", "投資額が指定価値に達する期間を返します。", 197),
    item("FORMULATEXT", "情報", "FORMULATEXT(reference)", "参照セルの数式を文字列として返します。", 198),
    item("ISFORMULA", "情報", "ISFORMULA(reference)", "参照セルに数式が含まれるかどうかを返します。", 199),
    item("ISERR", "情報", "ISERR(value)", "#N/A 以外のエラーかどうかを返します。", 200),
    item("ISLOGICAL", "情報", "ISLOGICAL(value)", "値が論理値かどうかを返します。", 201),
    item("ISNONTEXT", "情報", "ISNONTEXT(value)", "値が文字列以外かどうかを返します。", 202),
    item("ISREF", "情報", "ISREF(value)", "値が参照かどうかを返します。", 203),
    item("ISEVEN", "情報", "ISEVEN(number)", "数値が偶数かどうかを返します。", 204),
    item("ISODD", "情報", "ISODD(number)", "数値が奇数かどうかを返します。", 205),
    item("NA", "情報", "NA()", "#N/A エラー値を返します。", 206),
    item("TYPE", "情報", "TYPE(value)", "値のデータ型を表す番号を返します。", 207),
    item("CELL", "情報", "CELL(info_type, [reference])", "セルの書式、位置、内容に関する情報を返します。", 207.2),
    item("INFO", "情報", "INFO(type_text)", "現在の環境に関する情報を返します。", 207.4),
    item("LOG10", "数学/三角", "LOG10(number)", "10 を底とする対数を返します。", 208),
    item("SQRTPI", "数学/三角", "SQRTPI(number)", "数値に π を掛けた平方根を返します。", 209),
    item("FACTDOUBLE", "数学/三角", "FACTDOUBLE(number)", "数値の二重階乗を返します。", 210),
    item("COMBINA", "数学/三角", "COMBINA(number, number_chosen)", "重複組み合わせの数を返します。", 211),
    item("GCD", "数学/三角", "GCD(number1, [number2], ...)", "最大公約数を返します。", 212),
    item("LCM", "数学/三角", "LCM(number1, [number2], ...)", "最小公倍数を返します。", 213),
    item("MULTINOMIAL", "数学/三角", "MULTINOMIAL(number1, [number2], ...)", "多項係数を返します。", 214),
    item("PERMUT", "数学/三角", "PERMUT(number, number_chosen)", "順列の数を返します。", 214.1),
    item("PERMUTATIONA", "数学/三角", "PERMUTATIONA(number, number_chosen)", "重複を許す順列の数を返します。", 214.2),
    item("SERIESSUM", "数学/三角", "SERIESSUM(x, n, m, coefficients)", "べき級数の和を返します。", 215),
    item("SUMSQ", "数学/三角", "SUMSQ(number1, [number2], ...)", "引数の平方和を返します。", 216),
    item("SUMX2MY2", "数学/三角", "SUMX2MY2(array_x, array_y)", "2 つの配列で対応する値の平方差の合計を返します。", 216.2),
    item("SUMX2PY2", "数学/三角", "SUMX2PY2(array_x, array_y)", "2 つの配列で対応する値の平方和の合計を返します。", 216.4),
    item("SUMXMY2", "数学/三角", "SUMXMY2(array_x, array_y)", "2 つの配列で対応する値の差の平方和を返します。", 216.6),
    item("SUMPRODUCT", "数学/三角", "SUMPRODUCT(array1, [array2], ...)", "対応する配列要素の積の合計を返します。", 217),
    item("EVEN", "数学/三角", "EVEN(number)", "最も近い偶数に切り上げます。", 218),
    item("ODD", "数学/三角", "ODD(number)", "最も近い奇数に切り上げます。", 219),
    item("ISO.CEILING", "数学/三角", "ISO.CEILING(number, [significance])", "ISO 規格に従って基準値の倍数に切り上げます。", 220),
    item("ATAN2", "数学/三角", "ATAN2(x_num, y_num)", "指定座標のアークタンジェントを返します。", 221),
    item("SINH", "数学/三角", "SINH(number)", "双曲線正弦を返します。", 222),
    item("COSH", "数学/三角", "COSH(number)", "双曲線余弦を返します。", 223),
    item("TANH", "数学/三角", "TANH(number)", "双曲線正接を返します。", 224),
    item("ASINH", "数学/三角", "ASINH(number)", "双曲線逆正弦を返します。", 225),
    item("ACOSH", "数学/三角", "ACOSH(number)", "双曲線逆余弦を返します。", 226),
    item("ATANH", "数学/三角", "ATANH(number)", "双曲線逆正接を返します。", 227),
    item("COT", "数学/三角", "COT(number)", "角度のコタンジェントを返します。", 228),
    item("SEC", "数学/三角", "SEC(number)", "角度のセカントを返します。", 229),
    item("CSC", "数学/三角", "CSC(number)", "角度のコセカントを返します。", 230),
    item("ACOT", "数学/三角", "ACOT(number)", "アークコタンジェントを返します。", 231),
    item("COTH", "数学/三角", "COTH(number)", "双曲線コタンジェントを返します。", 232),
    item("SECH", "数学/三角", "SECH(number)", "双曲線セカントを返します。", 233),
    item("CSCH", "数学/三角", "CSCH(number)", "双曲線コセカントを返します。", 234),
    item("AVEDEV", "統計", "AVEDEV(number1, [number2], ...)", "データの平均偏差を返します。", 235),
    item("DEVSQ", "統計", "DEVSQ(number1, [number2], ...)", "平均値との差の平方和を返します。", 236),
    item("GEOMEAN", "統計", "GEOMEAN(number1, [number2], ...)", "幾何平均を返します。", 237),
    item("HARMEAN", "統計", "HARMEAN(number1, [number2], ...)", "調和平均を返します。", 238),
    item("RSQ", "統計", "RSQ(known_y's, known_x's)", "ピアソン積率相関係数の二乗を返します。", 239),
    item("PEARSON", "統計", "PEARSON(array1, array2)", "ピアソン相関係数を返します。", 240),
    item("Z.TEST", "統計", "Z.TEST(array, x, [sigma])", "z 検定の片側 P 値を返します。", 241),
    item("F.TEST", "統計", "F.TEST(array1, array2)", "F 検定の結果を返します。", 242),
    item("STEYX", "統計", "STEYX(known_y's, known_x's)", "回帰直線からの標準誤差を返します。", 243),
    item("SKEW", "統計", "SKEW(number1, [number2], ...)", "分布の歪度を返します。", 244),
    item("SKEW.P", "統計", "SKEW.P(number1, [number2], ...)", "母集団の歪度を返します。", 245),
    item("EXPON.DIST", "統計", "EXPON.DIST(x, lambda, cumulative)", "指数分布を返します。", 246),
    item("FISHER", "統計", "FISHER(x)", "フィッシャー変換を返します。", 247),
    item("FISHERINV", "統計", "FISHERINV(y)", "フィッシャー変換の逆関数を返します。", 248),
    item("GAMMA", "統計", "GAMMA(number)", "ガンマ関数の値を返します。", 249),
    item("GAMMA.DIST", "統計", "GAMMA.DIST(x, alpha, beta, cumulative)", "ガンマ分布を返します。", 250),
    item("GAMMA.INV", "統計", "GAMMA.INV(probability, alpha, beta)", "ガンマ分布の逆関数を返します。", 251),
    item("GAMMALN", "統計", "GAMMALN(x)", "ガンマ関数の自然対数を返します。", 252),
    item("GAMMALN.PRECISE", "統計", "GAMMALN.PRECISE(x)", "ガンマ関数の自然対数を返します。", 252.1),
    item("BETA.DIST", "統計", "BETA.DIST(x, alpha, beta, cumulative, [A], [B])", "ベータ分布を返します。", 253),
    item("BETA.INV", "統計", "BETA.INV(probability, alpha, beta, [A], [B])", "ベータ分布の逆関数を返します。", 254),
    item("BINOM.DIST", "統計", "BINOM.DIST(number_s, trials, probability_s, cumulative)", "二項分布を返します。", 255),
    item("BINOM.DIST.RANGE", "統計", "BINOM.DIST.RANGE(trials, probability_s, number_s, [number_s2])", "二項分布で指定範囲の成功回数になる確率を返します。", 255.2),
    item("BINOM.INV", "統計", "BINOM.INV(trials, probability_s, alpha)", "累積二項分布が基準値以上になる最小値を返します。", 256),
    item("CHISQ.DIST", "統計", "CHISQ.DIST(x, deg_freedom, cumulative)", "カイ二乗分布を返します。", 257),
    item("CHISQ.INV", "統計", "CHISQ.INV(probability, deg_freedom)", "カイ二乗分布の逆関数を返します。", 258),
    item("F.DIST", "統計", "F.DIST(x, deg_freedom1, deg_freedom2, cumulative)", "F 分布を返します。", 259),
    item("F.INV", "統計", "F.INV(probability, deg_freedom1, deg_freedom2)", "F 分布の逆関数を返します。", 260),
    item("NORM.S.DIST", "統計", "NORM.S.DIST(z, cumulative)", "標準正規分布を返します。", 261),
    item("NORM.S.INV", "統計", "NORM.S.INV(probability)", "標準正規分布の逆関数を返します。", 262),
    item("STANDARDIZE", "統計", "STANDARDIZE(x, mean, standard_dev)", "標準化変量を返します。", 263),
    item("TRIMMEAN", "統計", "TRIMMEAN(array, percent)", "データの上下の割合を除外した平均を返します。", 263.2),
    item("POISSON.DIST", "統計", "POISSON.DIST(x, mean, cumulative)", "ポアソン分布を返します。", 264),
    item("WEIBULL.DIST", "統計", "WEIBULL.DIST(x, alpha, beta, cumulative)", "ワイブル分布を返します。", 265),
    item("HYPGEOM.DIST", "統計", "HYPGEOM.DIST(sample_s, number_sample, population_s, number_pop, cumulative)", "超幾何分布を返します。", 266),
    item("CONFIDENCE.NORM", "統計", "CONFIDENCE.NORM(alpha, standard_dev, size)", "母平均の信頼区間を返します。", 267),
    item("CONFIDENCE.T", "統計", "CONFIDENCE.T(alpha, standard_dev, size)", "t 分布を使った信頼区間を返します。", 268),
    item("SPLIT", "文字列", "SPLIT(text, delimiter)", "文字列を区切り文字で分割します。", 269),
    item("T", "文字列", "T(value)", "値が文字列ならその文字列を返します。", 270),
    item("UNICODE", "文字列", "UNICODE(text)", "文字列の先頭文字の Unicode 番号を返します。", 271),
    item("UNICHAR", "文字列", "UNICHAR(number)", "Unicode 番号に対応する文字を返します。", 272),
    item("BITAND", "エンジニアリング", "BITAND(number1, number2)", "2 つの数値のビット積を返します。", 273),
    item("BITOR", "エンジニアリング", "BITOR(number1, number2)", "2 つの数値のビット和を返します。", 274),
    item("BITXOR", "エンジニアリング", "BITXOR(number1, number2)", "2 つの数値の排他的ビット和を返します。", 275),
    item("BITLSHIFT", "エンジニアリング", "BITLSHIFT(number, shift_amount)", "数値を左にビットシフトします。", 276),
    item("BITRSHIFT", "エンジニアリング", "BITRSHIFT(number, shift_amount)", "数値を右にビットシフトします。", 277),
    item("BIN2DEC", "エンジニアリング", "BIN2DEC(number)", "2 進数を 10 進数に変換します。", 278),
    item("BIN2HEX", "エンジニアリング", "BIN2HEX(number, [places])", "2 進数を 16 進数に変換します。", 279),
    item("BIN2OCT", "エンジニアリング", "BIN2OCT(number, [places])", "2 進数を 8 進数に変換します。", 280),
    item("OCT2BIN", "エンジニアリング", "OCT2BIN(number, [places])", "8 進数を 2 進数に変換します。", 281),
    item("OCT2DEC", "エンジニアリング", "OCT2DEC(number)", "8 進数を 10 進数に変換します。", 282),
    item("OCT2HEX", "エンジニアリング", "OCT2HEX(number, [places])", "8 進数を 16 進数に変換します。", 283),
    item("HEX2BIN", "エンジニアリング", "HEX2BIN(number, [places])", "16 進数を 2 進数に変換します。", 284),
    item("HEX2OCT", "エンジニアリング", "HEX2OCT(number, [places])", "16 進数を 8 進数に変換します。", 285),
    item("IMARGUMENT", "エンジニアリング", "IMARGUMENT(inumber)", "複素数の偏角を返します。", 286),
    item("IMEXP", "エンジニアリング", "IMEXP(inumber)", "複素数の指数関数を返します。", 287),
    item("IMLN", "エンジニアリング", "IMLN(inumber)", "複素数の自然対数を返します。", 288),
    item("IMLOG10", "エンジニアリング", "IMLOG10(inumber)", "複素数の常用対数を返します。", 289),
    item("IMLOG2", "エンジニアリング", "IMLOG2(inumber)", "複素数の 2 を底とする対数を返します。", 290),
    item("IMPOWER", "エンジニアリング", "IMPOWER(inumber, number)", "複素数のべき乗を返します。", 291),
    item("IMSQRT", "エンジニアリング", "IMSQRT(inumber)", "複素数の平方根を返します。", 292),
    item("IMSIN", "エンジニアリング", "IMSIN(inumber)", "複素数のサインを返します。", 293),
    item("IMCOS", "エンジニアリング", "IMCOS(inumber)", "複素数のコサインを返します。", 294),
    item("IMTAN", "エンジニアリング", "IMTAN(inumber)", "複素数のタンジェントを返します。", 295),
    item("IMSEC", "エンジニアリング", "IMSEC(inumber)", "複素数のセカントを返します。", 296),
    item("IMCSC", "エンジニアリング", "IMCSC(inumber)", "複素数のコセカントを返します。", 297),
    item("IMCOT", "エンジニアリング", "IMCOT(inumber)", "複素数のコタンジェントを返します。", 298),
    item("IMSINH", "エンジニアリング", "IMSINH(inumber)", "複素数の双曲線サインを返します。", 299),
    item("IMCOSH", "エンジニアリング", "IMCOSH(inumber)", "複素数の双曲線コサインを返します。", 300),
    item("IMSECH", "エンジニアリング", "IMSECH(inumber)", "複素数の双曲線セカントを返します。", 301),
    item("IMCSCH", "エンジニアリング", "IMCSCH(inumber)", "複素数の双曲線コセカントを返します。", 302),
    item("BESSELI", "エンジニアリング", "BESSELI(x, n)", "修正ベッセル関数 In(x) を返します。", 303),
    item("BESSELJ", "エンジニアリング", "BESSELJ(x, n)", "ベッセル関数 Jn(x) を返します。", 304),
    item("BESSELK", "エンジニアリング", "BESSELK(x, n)", "修正ベッセル関数 Kn(x) を返します。", 305),
    item("BESSELY", "エンジニアリング", "BESSELY(x, n)", "ベッセル関数 Yn(x) を返します。", 306),
    item("ERF.PRECISE", "エンジニアリング", "ERF.PRECISE(x)", "誤差関数の値を返します。", 307),
    item("ERFC.PRECISE", "エンジニアリング", "ERFC.PRECISE(x)", "相補誤差関数の値を返します。", 308),
  ];
}

export {
  absoluteRangeReference,
  activeCellEditElementForFormatting,
  activeRibbonPanelElement,
  activeRibbonTabElement,
  activeSheetEditElement,
  additionalFormulaCatalogDefinitions,
  applyAutoSumToSelection,
  applyAutoSumToSheetGroup,
  applyCreateTableDialog,
  applyExcelTableStyle,
  applyHyperlinkToObjects,
  applyHyperlinkToObjectsForSheetGroup,
  applyObjectLinkDialog,
  autoSumCommandLabel,
  cancelIdleViewportTrim,
  cancelViewportRender,
  closeFileSideMenu,
  closeFileSideMenuOnEscape,
  closeFileSideMenuOnOutsidePointer,
  closeObjectLinkDialog,
  closeObjectLinkDialogOnEscape,
  closeRibbonDisplayMenu,
  closeRibbonDisplayMenuOnEscape,
  closeRibbonDisplayMenuOnOutsidePointer,
  commitDataOperation,
  commitDataOperationForSheet,
  contextMenuActionPoint,
  currentWorkbookPaneFocusZone,
  defaultLinkDialogValue,
  definedNameDuplicateMessage,
  definedNameHasConflict,
  definedNameKey,
  definedNameLinkReference,
  ensureSheetCoversViewport,
  ensureSheetSize,
  extendSheet,
  firstCellReferenceFromRangeLabel,
  flushDeferredViewportRender,
  focusActiveRibbonTab,
  focusFirstRibbonPanelControl,
  focusRibbonDisplayMenuButton,
  focusRibbonPanelControl,
  focusRibbonTabAt,
  focusWorkbookPaneByKeyboard,
  focusWorkbookPaneZone,
  formulaAutocompleteExtraFunctionDefinitions,
  formulaFunctionByName,
  formulaFunctionDefinitions,
  handleKeyboardContextMenuShortcut,
  handleObjectLinkDialogAction,
  handleObjectLinkDocumentNameClick,
  handleObjectLinkDocumentSheetClick,
  handleObjectLinkKindClick,
  handleRibbonAccessKeyShortcut,
  handleRibbonCommand,
  handleRibbonDisplayMenuKeydown,
  handleRibbonFocusShortcut,
  handleRibbonLayoutOptionClick,
  handleRibbonPanelKeydown,
  handleRibbonTabKeydown,
  handleRibbonVisibilityOptionClick,
  handleWorkbookKeyboardCaptureShortcut,
  handleWorkbookPaneFocusShortcut,
  hideRibbonForAutoHide,
  hideRibbonForAutoHideOnOutsidePointer,
  hyperlinkKind,
  indexForPixelExtent,
  insertTableFromRange,
  insertTableFromSelection,
  insertTableToSheetGroup,
  internalHyperlinkReferenceFromDefinedName,
  internalHyperlinkTextIsCellReference,
  isCellReference,
  isInternalHyperlinkValue,
  isPersistentCell,
  isRibbonPanelNativeKeyTarget,
  keyboardContextMenuPointForCell,
  keyboardContextMenuPointForColumnHeader,
  keyboardContextMenuPointForImage,
  keyboardContextMenuPointForRowHeader,
  linkDefinedNameFromTarget,
  linkDialogInitialValues,
  linkKindEmptyMessage,
  linkKindMeta,
  loadRibbonDisplayPreferences,
  minimumSheetExtent,
  nextTableName,
  normalizeAutoSumFunctionName,
  normalizeDefinedNameRefersTo,
  normalizeEmailHyperlinkUrl,
  normalizeHyperlinkForKind,
  normalizeHyperlinkUrl,
  normalizeInternalHyperlinkUrl,
  normalizeRibbonLayout,
  normalizeRibbonVisibility,
  normalizedDocumentSheetIndex,
  normalizedInternalDefinedNameHyperlinkTarget,
  objectLinkDialogHtml,
  openCreateTableDialog,
  openFileSideMenu,
  openInternalHyperlink,
  openNewWorkbookFromFileMenu,
  openRibbonDisplayMenu,
  openSelectedObjectHyperlink,
  openSelectionContextMenuByKeyboard,
  openWebVersionFromLocalFile,
  parseCellKey,
  parseInternalHyperlinkTarget,
  positionFileSideMenu,
  quoteSheetNameForFormula,
  releaseRibbonPointerFocus,
  removeObjectLinkFromDialog,
  renameObjectForSheetGroup,
  renameSelectedObjects,
  resetLazyScrollLocks,
  resetWorkbookModifierState,
  resizeActiveSheetForScroll,
  restoreViewportScroll,
  ribbonDisplayMenuButtons,
  ribbonPanelFocusableControls,
  ribbonPanelFocusedControl,
  ribbonPanelsAreVisible,
  rowHeaderReferenceRow,
  rowHeaderWidthForSheet,
  sanitizeDefinedName,
  saveRibbonDisplayPreferences,
  scheduleIdleViewportTrim,
  scheduleViewportRender,
  selectedDataRange,
  setActiveRibbonTab,
  setRibbonLayout,
  setRibbonVisibility,
  sheetExtentForViewport,
  sheetPixelToClientPoint,
  shouldShowFileSideMenuWebVersionItem,
  showRibbonForAutoHide,
  stripInternalHyperlinkPrefix,
  stripMailtoPrefix,
  suggestedDefinedNameForRange,
  switchRibbonTab,
  syncFileSideMenuLocalStandaloneItems,
  syncRibbonCollapsed,
  syncRibbonDisplayState,
  toggleRibbonCollapsed,
  toggleRibbonDisplayMenu,
  trimSheetToViewportExtent,
  tryExtendFormulaSheet,
  unquoteFormulaSheetName,
  updateWorkbookModifierState,
  upsertDefinedName,
  validateDefinedName,
  visibleRibbonTabs,
  workbookPaneFocusZones,
};
