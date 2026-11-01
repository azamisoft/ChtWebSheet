// Feature module; shared bindings are imported explicitly.
import {
  $formulaInput,
  $horizontalScrollbar,
  $sheetHost,
  CHART_COLORS,
  COLUMN_HEADER_HEIGHT,
  DEFAULT_SHAPE_TYPE,
  LINE_SHAPE_ARROW_PADDING_EXTRA_PX,
  LINE_SHAPE_BASE_PADDING_PX,
  LINE_SHAPE_MIN_BOX_SIZE,
  state,
} from "../../app-core.js";
import {
  applyDataAnalysisDialog,
  applyDataTableDialog,
  applyScenario,
  applyWhatIfDialog,
  cloneScenarios,
  createScenarioSummarySheet,
  dataAnalysisPickerDialogHtml,
  ensureScenarioPermission,
  loadScenarioCurrentValues,
  openDataAnalysisOptionsDialog,
  openScenarioEditDialog,
  openScenarioManagerDialog,
  openScenarioSummaryDialog,
  persistDataAnalysisDialogState,
  saveScenarioFromDialog,
  scenarioSummaryDialogValidation,
  selectedDataAnalysisPickerMethod,
  selectedScenarioFromDialog,
  showDataAnalysisHelp,
  syncDataAnalysisDialog,
  syncDataAnalysisPickerDialog,
  syncScenarioSummaryDialog,
  workbookScenarios,
} from "../data/analysis-and-scenarios.js";
import {
  applyAppendWorkbookQueryDialog,
  applyImportDataDialog,
  applyMergeWorkbookQueryDialog,
  applyQueryRenameDialog,
  applyQuerySourceChangeDialog,
  createPivotFromDataModelTable,
  dataModelQueryRef,
  handleExistingConnectionsDialogAction,
  loadWebNavigatorSelection,
  openExistingConnectionsDialog,
  openImportDataDialog,
  openQueryPropertiesDialog,
  openWebNavigatorTransformEditor,
  refreshAllDataModelTables,
  refreshDataModelTable,
  removeDataModelTable,
  selectQueryRange,
  selectWebNavigatorSource,
  selectedDataModelTableFromDialog,
} from "../data/connections-and-queries.js";
import {
  applyAdvancedFilterDialog,
  applyCustomAutoFilterDialog,
  applyTextToColumnsDialog,
  applyTop10AutoFilterDialog,
  saveTextColumnsColumnFormat,
  setDialogRangeFieldFromCurrentSelection,
  syncTextToColumnsDialog,
  toggleTextColumnsFixedBreak,
} from "../data/filter-and-text-tools.js";
import {
  applyForecastSheetDialog,
  applyOutlineDialog,
  applyOutlineSettingsDialog,
  outlineGroups,
  setOutlineGroupCollapsed,
} from "../data/forecast-and-statistics.js";
import {
  applyTextCsvImportDialog,
  handleQueryTransformEditorAction,
  openTextCsvTransformEditor,
} from "../data/query-editor.js";
import {
  applyConnectionOnlySourceDialog,
  applyGetDataFromTableRangeDialog,
  applyODataImportDialog,
  applyWebImportDialog,
  applyWorkbookLinkSourceChange,
  handleWorkbookLinksDialogAction,
  openODataTransformEditor,
  openTableRangeTransformEditor,
  openWebTransformEditor,
  openWorkbookLinksDialog,
} from "../data/query-transforms.js";
import {
  applyQueryPropertiesDialog,
  cancelWorkbookRefreshes,
  clearQueryPropertiesCredential,
  mutateSortDialogLevels,
  openRefreshStatusDialog,
  refreshWorkbookData,
  setSortDialogActiveLevel,
  sortDialogValidation,
  sortRangeRowsByKeys,
  syncSortDialog,
  syncSortDialogOrientation,
} from "../data/sort-and-autofilter.js";
import {
  applyConsolidateDialog,
  applyDataValidationDialog,
  applyRelationshipAutoDetectDialog,
  applyRelationshipEditDialog,
  applyRemoveDuplicatesDialog,
  applySubtotalDialog,
  autoDetectRelationshipsFromDialog,
  canCreateWorkbookRelationship,
  decodeWorkbookRangeOrName,
  deleteRelationshipFromDialog,
  openRelationshipEditDialog,
  openRelationshipsDialog,
  relationshipUnavailableMessage,
  resetDataValidationDialog,
  selectedRelationshipFromDialog,
  syncConsolidateDialog,
  syncRemoveDuplicatesDialog,
} from "../data/validation-and-consolidation.js";
import {
  addWorksheetImageDescriptorsToSheet,
  applyHeaderFooterDialog,
  cssColorForSvg,
  endingCellForPixelSize,
  fileExtension,
  fileToDataUrl,
  imageDimensionsFromSource,
  insertEmbeddedObjectPlaceholder,
  insertIconObject,
  insertOfficeAddinObject,
  insertPresetTextObject,
  insertWorksheetImageDescriptors,
  isLineShapeType,
  normalizeLineShapeConnections,
  normalizeShapeArrowType,
  normalizeShapeLinePoints,
  normalizeShapeOutlineWidth,
  openOfficeAddinsDialog,
  setCurrentRibbonShape,
  shapeArrowHalfWidth,
  shapeDescriptorForType,
  shapeLineDashArray,
  shapePresetForType,
} from "./shape-rendering.js";
import {
  lineConnectionPreviewTarget,
  removeLineConnectionPreviewOverlay,
  snapLineEndpointToShape,
  updateLineConnectionPreviewOverlay,
} from "./transforms-and-connectors.js";
import {
  openShapeTextEditor,
} from "../editing/autofill.js";
import {
  closeRibbonShapeMenu,
  svgDataUrl,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  applyPasteSpecialDialog,
  isElementImeCompositionEvent,
} from "../editing/keyboard-and-paste.js";
import {
  activateObjectFormatRibbonForSelection,
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  getDisplayForCell,
  materializeSelectionRange,
  materializeSelectionRangeForSheet,
  nativeAnchorFromSheetPixel,
  normalizeRange,
  rangeToLabel,
  restoreGroupedSheetsAfterFailure,
  selectionRangeFromPoints,
  sheetPixelFromClientPosition,
  stopSelectionAutoScroll,
  updateCellModel,
  updateSelectionUi,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  selectionToolbarPoint,
  setSelectedImageIds,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureCellsForHistory,
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneJsonLikeValue,
  internCellsInRange,
  pasteInternalClipboardLinkAtRange,
  pushHistory,
  pushSheetInsertHistory,
  pushSheetInsertManyHistory,
  redoLastOperation,
  undoLastOperation,
} from "../editing/structure-and-history.js";
import {
  applyFunctionArgumentsDialog,
  applyPasteNameDialog,
  applyPasteNameListDialog,
  expandFunctionArgumentReferencePicker,
  focusFirstVisibleInsertFunction,
  openFunctionArgumentsDialog,
  restoreFunctionArgumentsEditorFocus,
  returnToInsertFunctionDialogFromArguments,
  selectedDefinedNameIndexFromDialog,
  syncInsertFunctionDialog,
  syncNameManagerDialog,
  toggleFunctionArgumentReferencePicker,
} from "../formula/dialogs-and-names.js";
import {
  ensureCellModel,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  openDefineNameDialog,
} from "../ui/annotations-and-cell-format.js";
import {
  accessKeyLabelHtml,
  cellKey,
  columnName,
  escapeAttr,
  escapeHtml,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  applyPageSetupDialog,
  cancelInkDrawingMode,
} from "../ui/printing-and-home-tools.js";
import {
  absoluteRangeReference,
  applyCreateTableDialog,
  applyExcelTableStyle,
  ensureSheetSize,
  rowHeaderWidthForSheet,
  selectedDataRange,
} from "../ui/ribbon-and-links.js";
import {
  applySheetProtectionConfirmDialog,
  applySheetProtectionDialog,
  applySheetUnprotectDialog,
  createBlankSheet,
  ensureEditableRangesForProtection,
  ensureSheetProtectionPermission,
} from "../ui/sheet-protection-and-translation.js";
import {
  applyCreateNamesFromSelection,
  applyNameManagerSave,
  applySheetViewNameDialog,
  deleteDefinedNameFromManager,
  handleFormulaErrorDialogAction,
  handleFormulaEvaluateDialogAction,
  openNameEditorFromManager,
} from "../ui/views-and-formula-tools.js";
import {
  applySheetInsertDialog,
  applySheetMoveCopyDialog,
  applyUnhideSheetDialog,
  buildFormulaEngine,
  renderWorkbook,
  selectSheetInsertTemplate,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  activeSheet,
  applySaveAsDialog,
  chooseSaveAsLocation,
  createSaveAsFolder,
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
  imageMimeType,
} from "../workbook/tables-and-validation.js";
import {
  applyMacroEditorDialog,
  applyMacroOptionsDialog,
  applyMacroSecurityDialog,
  applyPrintDialogPageSettings,
  applySheetCodeDialog,
  createMacroFromDialog,
  deleteMacroFromDialog,
  executeMacroStepDialog,
  movePrintDialogPreview,
  openMacroEditorDialog,
  openMacroOptionsDialog,
  openMacroStepDialog,
  openPageSetupDialog,
  printWorkbookFromDialog,
  selectedMacroFromDialog,
  startMacroRecordingFromDialog,
} from "../workbook/vba-and-printing.js";
import {
  addCustomViewFromDialog,
  applySelectedCustomView,
  arrangeWorkbookWindowViews,
  deleteSelectedCustomView,
  enableViewSideBySide,
  menuAccessKeyFromLabel,
  openAddCustomViewDialog,
  openCustomViewsDialog,
  runMacroFromDialog,
  switchWorkbookWindow,
  unhideWorkbookWindow,
} from "../workbook/views-and-macros.js";


function outlineRangeForAxis(axis, start, end) {
  if (axis === "columns") {
    return { sheetIndex: state.activeSheetIndex, top: 1, bottom: activeSheet().rowCount, left: start, right: end, fullColumn: true };
  }
  return { sheetIndex: state.activeSheetIndex, top: start, bottom: end, left: 1, right: activeSheet().colCount, fullRow: true };
}



function toggleOutlineDetailFromSelection(collapse) {
  const context = selectedDataRange(collapse ? "詳細非表示" : "詳細表示");
  if (!context) return;
  const { sheet, range } = context;
  const group = outlineDetailTargetFromRange(sheet, range, collapse);
  if (!group) {
    showMessageBox("選択範囲にアウトラインはありません。", { title: "Cht WebSheet" });
    return;
  }
  setOutlineGroupCollapsed(group.axis, group.start, group.end, group.level, collapse);
}



function outlineDetailTargetFromRange(sheet, range, collapse, options = {}) {
  if (!sheet || !range) return null;
  const axis = range.fullColumn || (!range.fullRow && range.right - range.left > range.bottom - range.top) ? "columns" : "rows";
  const groups = outlineGroups(sheet, axis)
    .filter((group) => rangesOverlap1d(group.start, group.end, axis === "columns" ? range.left : range.top, axis === "columns" ? range.right : range.bottom))
    .sort((left, right) => right.level - left.level || (left.end - left.start) - (right.end - right.start));
  if (!options.requireChange) return groups[0] || null;
  return groups.find((group) => (collapse ? !group.collapsed : group.collapsed)) || null;
}



function rangesOverlap1d(leftStart, leftEnd, rightStart, rightEnd) {
  return leftStart <= rightEnd && rightStart <= leftEnd;
}



function handleDataDialogAction(event) {
  const action = event.currentTarget.dataset.dataDialogAction;
  const dialog = event.currentTarget.closest("[data-data-dialog]");
  const type = dialog?.dataset.dataDialog || state.dataDialog?.type;
  if (action === "range-from-selection") {
    setDialogRangeFieldFromCurrentSelection(dialog, event.currentTarget.dataset.dataRangeTarget, event.currentTarget.dataset.dataRangeMode);
    return;
  }
  if (action === "validation-tab") {
    const tab = event.currentTarget.dataset.validationTab;
    dialog?.querySelectorAll("[data-validation-tab]").forEach((button) => {
      button.classList.toggle("is-active", button.dataset.validationTab === tab);
    });
    dialog?.querySelectorAll("[data-validation-tab-panel]").forEach((panel) => {
      panel.classList.toggle("is-active", panel.dataset.validationTabPanel === tab);
    });
    return;
  }
  if (action === "cancel") {
    if (type === "function-arguments" && state.dataDialog?.editorSource) {
      const source = state.dataDialog.editorSource;
      closeDataDialog();
      restoreFunctionArgumentsEditorFocus(source);
      return;
    }
    if (type === "custom-view-add") {
      openCustomViewsDialog();
      return;
    }
    if (type === "import-data" && state.dataDialog?.returnToExistingConnectionsKey) {
      openExistingConnectionsDialog(state.dataDialog.returnToExistingConnectionsKey);
      return;
    }
    if (type === "query-properties" && state.dataDialog?.returnToExistingConnectionsKey) {
      openExistingConnectionsDialog(state.dataDialog.returnToExistingConnectionsKey);
      return;
    }
    if (type === "workbook-link-change") {
      openWorkbookLinksDialog(state.dataDialog?.linkId || "");
      return;
    }
    closeDataDialog();
    return;
  }
  if (type === "save-as") {
    if (action === "save-as-browse") {
      void chooseSaveAsLocation(dialog);
      return;
    }
    if (action === "save-as-new-folder") {
      void createSaveAsFolder(dialog);
      return;
    }
    if (action === "save-as-ok" || action === "ok") {
      void applySaveAsDialog(dialog);
      return;
    }
  }
  if (type === "insert-object") {
    if (action === "object-insert-browse") {
      chooseEmbeddedObjectFile(dialog);
      return;
    }
    if (action === "object-insert-ok" || action === "ok") {
      applyInsertObjectDialog(dialog);
      return;
    }
  }
  if (type === "signature-line" && (action === "signature-line-ok" || action === "ok")) {
    applySignatureLineDialog(dialog);
    return;
  }
  if (type === "insert-icon" && action === "insert-icon") {
    const kind = event.currentTarget.dataset.insertIconKind || "star";
    closeDataDialog();
    insertIconObject(kind);
    return;
  }
  if (type === "office-addins") {
    if (action === "office-addin-tab") {
      openOfficeAddinsDialog(event.currentTarget.dataset.officeAddinTab || "store");
      return;
    }
    if (action === "office-addin-insert") {
      const addinId = event.currentTarget.dataset.officeAddinId || "";
      closeDataDialog();
      insertOfficeAddinObject(addinId);
      return;
    }
  }
  if (type === "insert-function") {
    if (action === "insert-function-search") {
      syncInsertFunctionDialog(dialog);
      focusFirstVisibleInsertFunction(dialog);
      return;
    }
    if (action === "insert-function-select") {
      state.dataDialog.selectedFunction = event.currentTarget.dataset.formulaFunctionChoice || "SUM";
      syncInsertFunctionDialog(dialog);
      return;
    }
    if (action === "insert-function-ok" || action === "ok") {
      const functionName = state.dataDialog?.selectedFunction || "SUM";
      closeDataDialog();
      openFunctionArgumentsDialog(functionName);
      return;
    }
  }
  if (type === "function-arguments" && (action === "function-arguments-ok" || action === "ok")) {
    applyFunctionArgumentsDialog(dialog);
    return;
  }
  if (type === "function-arguments" && action === "function-argument-reference-toggle") {
    toggleFunctionArgumentReferencePicker(dialog, event.currentTarget);
    return;
  }
  if (type === "function-arguments" && action === "function-arguments-back") {
    returnToInsertFunctionDialogFromArguments();
    return;
  }
  if (type === "formula-error-checking" && handleFormulaErrorDialogAction(action, dialog, event.currentTarget)) {
    return;
  }
  if (type === "formula-evaluate" && handleFormulaEvaluateDialogAction(action, dialog)) {
    return;
  }
  if (type === "name-manager") {
    if (action === "name-manager-select") {
      const rowIndex = Number(event.currentTarget.dataset.formulaNameRowIndex);
      state.dataDialog.activeIndex = Number.isFinite(rowIndex) ? rowIndex : selectedDefinedNameIndexFromDialog(dialog);
      syncNameManagerDialog(dialog);
      return;
    }
    if (action === "name-manager-new") {
      closeDataDialog();
      const point = selectionToolbarPoint();
      if (point) openDefineNameDialog(point.row, point.col, { returnToNameManager: true });
      return;
    }
    if (action === "name-manager-edit") {
      openNameEditorFromManager(dialog);
      return;
    }
    if (action === "name-manager-delete") {
      deleteDefinedNameFromManager(dialog);
      return;
    }
    if (action === "name-manager-save") {
      applyNameManagerSave(dialog);
      return;
    }
  }
  if (type === "paste-name" && action === "paste-name-list") {
    applyPasteNameListDialog(dialog);
    return;
  }
  if (type === "paste-name" && (action === "paste-name-ok" || action === "ok")) {
    applyPasteNameDialog(dialog);
    return;
  }
  if (type === "paste-special" && action === "ok") {
    applyPasteSpecialDialog(dialog);
    return;
  }
  if (type === "paste-special" && action === "paste-special-link") {
    const range = activeSelectionRange();
    closeDataDialog();
    pasteInternalClipboardLinkAtRange(range);
    return;
  }
  if (type === "create-table" && (action === "create-table-ok" || action === "ok")) {
    applyCreateTableDialog(dialog);
    return;
  }
  if (type === "create-names" && (action === "create-names-ok" || action === "ok")) {
    applyCreateNamesFromSelection(dialog);
    return;
  }
  if (type === "custom-views") {
    if (action === "custom-view-add") {
      openAddCustomViewDialog();
      return;
    }
    if (action === "custom-view-show" || action === "ok") {
      applySelectedCustomView(dialog);
      return;
    }
    if (action === "custom-view-delete") {
      deleteSelectedCustomView(dialog);
      return;
    }
  }
  if (type === "custom-view-add" && (action === "custom-view-add-ok" || action === "ok")) {
    addCustomViewFromDialog(dialog);
    return;
  }
  if (type === "sheet-unhide" && (action === "sheet-unhide-ok" || action === "ok")) {
    applyUnhideSheetDialog(dialog);
    return;
  }
  if (type === "sheet-insert") {
    if (action === "sheet-insert-template") {
      selectSheetInsertTemplate(event.currentTarget);
      return;
    }
    if (action === "sheet-insert-ok" || action === "ok") {
      applySheetInsertDialog(dialog);
      return;
    }
  }
  if (type === "sheet-move-copy" && (action === "sheet-move-copy-ok" || action === "ok")) {
    applySheetMoveCopyDialog(dialog);
    return;
  }
  if (type === "sheet-view-name" && (action === "sheet-view-name-ok" || action === "ok")) {
    applySheetViewNameDialog(dialog);
    return;
  }
  if (type === "view-unhide-window" && (action === "view-window-unhide-ok" || action === "ok")) {
    const id = dialog?.querySelector("[data-data-field='windowId']")?.value || "";
    closeDataDialog();
    unhideWorkbookWindow(id);
    return;
  }
  if (type === "view-side-by-side-window" && (action === "view-side-by-side-ok" || action === "ok")) {
    const id = dialog?.querySelector("[data-data-field='sideBySideWindowId']")?.value || "";
    const activeId = state.model?.activeWindowId || "main";
    closeDataDialog();
    enableViewSideBySide(activeId, id);
    return;
  }
  if (type === "view-switch-window" && (action === "view-switch-window-ok" || action === "ok")) {
    const id = dialog?.querySelector("[data-data-field='switchWindowId']")?.value || "";
    closeDataDialog();
    switchWorkbookWindow(id);
    return;
  }
  if (type === "view-arrange-window" && (action === "view-arrange-ok" || action === "ok")) {
    const mode = dialog?.querySelector("input[name='viewArrangeMode']:checked")?.value || "tiled";
    closeDataDialog();
    arrangeWorkbookWindowViews(mode);
    return;
  }
  if (type === "macros") {
    if (action === "macro-run") {
      runMacroFromDialog(dialog);
      return;
    }
    if (action === "macro-step") {
      const selected = selectedMacroFromDialog(dialog);
      if (selected) openMacroStepDialog(selected.index);
      return;
    }
    if (action === "macro-create") {
      createMacroFromDialog(dialog);
      return;
    }
    if (action === "macro-edit") {
      const selected = selectedMacroFromDialog(dialog);
      if (selected) openMacroEditorDialog(selected.index);
      return;
    }
    if (action === "macro-options") {
      const selected = selectedMacroFromDialog(dialog);
      if (selected) openMacroOptionsDialog(selected.index);
      return;
    }
    if (action === "macro-delete") {
      deleteMacroFromDialog(dialog);
      return;
    }
  }
  if (type === "macro-editor" && (action === "macro-editor-ok" || action === "ok")) {
    applyMacroEditorDialog(dialog);
    return;
  }
  if (type === "macro-step") {
    if (action === "macro-step-next") {
      executeMacroStepDialog(dialog, "next");
      return;
    }
    if (action === "macro-step-continue") {
      executeMacroStepDialog(dialog, "continue");
      return;
    }
    if (action === "macro-step-edit") {
      const index = Number(dialog?.dataset.macroIndex ?? state.dataDialog?.index);
      openMacroEditorDialog(index);
      return;
    }
  }
  if (type === "macro-options" && (action === "macro-options-ok" || action === "ok")) {
    applyMacroOptionsDialog(dialog);
    return;
  }
  if (type === "macro-record" && (action === "macro-record-ok" || action === "ok")) {
    startMacroRecordingFromDialog(dialog);
    return;
  }
  if (type === "macro-security" && (action === "macro-security-ok" || action === "ok")) {
    applyMacroSecurityDialog(dialog);
    return;
  }
  if (type === "sheet-code" && (action === "sheet-code-ok" || action === "ok")) {
    applySheetCodeDialog(dialog);
    return;
  }
  if (type === "print") {
    if (action === "print-preview-prev") {
      movePrintDialogPreview(dialog, -1);
      return;
    }
    if (action === "print-preview-next") {
      movePrintDialogPreview(dialog, 1);
      return;
    }
    if (action === "print-page-setup") {
      applyPrintDialogPageSettings(dialog, { render: false, statusMessage: "印刷設定をページ設定へ反映しました。" });
      openPageSetupDialog();
      return;
    }
    if (action === "print-now" || action === "ok") {
      printWorkbookFromDialog(dialog);
      return;
    }
  }
  if (type === "page-setup" && action === "ok") {
    applyPageSetupDialog(dialog);
    return;
  }
  if (type === "sort") {
    if (action === "sort-level-select") {
      setSortDialogActiveLevel(dialog, event.currentTarget.dataset.sortLevelIndex);
      return;
    }
    if (["sort-add-level", "sort-delete-level", "sort-copy-level", "sort-move-up", "sort-move-down"].includes(action)) {
      mutateSortDialogLevels(dialog, action);
      return;
    }
    if (action === "sort-options-toggle") {
      dialog?.querySelector(".sort-options-panel")?.classList.toggle("hidden");
      return;
    }
    if (action === "sort-orientation" || action === "sort-header-toggle") {
      syncSortDialogOrientation(dialog);
      return;
    }
  }
  if (type === "sort" && action === "ok") {
    const validation = sortDialogValidation(dialog);
    if (!validation.valid) {
      showMessageBox(validation.message, { title: "Cht WebSheet" });
      syncSortDialog(dialog, { capture: false });
      return;
    }
    const { range, sortKeys, hasHeader, orientation } = validation;
    const caseSensitive = Boolean(dialog.querySelector("[data-data-field='caseSensitive']")?.checked);
    closeDataDialog();
    sortRangeRowsByKeys(range, sortKeys, hasHeader, { orientation, caseSensitive });
    return;
  }
  if (type === "remove-duplicates") {
    if (action === "select-all-columns" || action === "unselect-all-columns") {
      dialog.querySelectorAll("[data-data-duplicate-col]").forEach((input) => {
        input.checked = action === "select-all-columns";
      });
      syncRemoveDuplicatesDialog(dialog);
      return;
    }
    if (action === "ok") {
      applyRemoveDuplicatesDialog(dialog);
      return;
    }
  }
  if (type === "consolidate") {
    if (action === "add-reference") {
      const input = dialog.querySelector("[data-data-field='reference']");
      const list = dialog.querySelector("[data-data-field='references']");
      const value = input?.value.trim() || "";
      if (!decodeWorkbookRangeOrName(value, activeSheet())) {
        showMessageBox("参照範囲を正しく入力してください。", { title: "Cht WebSheet" });
        return;
      }
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      option.selected = true;
      list?.appendChild(option);
      syncConsolidateDialog(dialog);
      return;
    }
    if (action === "delete-reference") {
      dialog.querySelector("[data-data-field='references'] option:checked")?.remove();
      syncConsolidateDialog(dialog);
      return;
    }
    if (action === "ok") {
      applyConsolidateDialog(dialog);
      return;
    }
  }
  if (type === "text-to-columns") {
    if (action === "text-columns-next" || action === "text-columns-back") {
      saveTextColumnsColumnFormat(dialog);
      const delta = action === "text-columns-next" ? 1 : -1;
      state.dataDialog.step = Math.min(3, Math.max(1, (Number(state.dataDialog.step) || 1) + delta));
      syncTextToColumnsDialog(dialog);
      return;
    }
    if (action === "text-columns-select-column") {
      saveTextColumnsColumnFormat(dialog);
      state.dataDialog.textColumnsSelectedColumn = Number(event.currentTarget.dataset.textColumnIndex) || 0;
      syncTextToColumnsDialog(dialog);
      return;
    }
    if (action === "text-columns-toggle-fixed-break") {
      toggleTextColumnsFixedBreak(dialog, Number(event.currentTarget.dataset.textFixedPosition));
      return;
    }
    if (action === "text-columns-finish" || action === "ok") {
      applyTextToColumnsDialog(dialog);
      return;
    }
    return;
  }
  if (type === "text-csv-import") {
    if (action === "load") {
      applyTextCsvImportDialog(dialog);
      return;
    }
    if (action === "transform") {
      openTextCsvTransformEditor(dialog);
      return;
    }
  }
  if (type === "query-transform-editor") {
    handleQueryTransformEditorAction(dialog, action, event);
    return;
  }
  if (type === "connection-only-source") {
    if (action === "ok") applyConnectionOnlySourceDialog(dialog);
    return;
  }
  if (type === "web-import" && action === "load") {
    void applyWebImportDialog(dialog);
    return;
  }
  if (type === "web-import" && action === "transform") {
    void openWebTransformEditor(dialog);
    return;
  }
  if (type === "odata-import" && action === "load") {
    void applyODataImportDialog(dialog);
    return;
  }
  if (type === "odata-import" && action === "transform") {
    void openODataTransformEditor(dialog);
    return;
  }
  if (type === "web-navigator") {
    if (action === "web-navigator-select") {
      selectWebNavigatorSource(dialog, event.currentTarget.dataset.webSourceIndex);
      return;
    }
    if (action === "load") {
      loadWebNavigatorSelection(dialog);
      return;
    }
    if (action === "transform") {
      openWebNavigatorTransformEditor(dialog);
      return;
    }
  }
  if (type === "get-table-range" && action === "load") {
    applyGetDataFromTableRangeDialog(dialog);
    return;
  }
  if (type === "get-table-range" && action === "transform") {
    openTableRangeTransformEditor(dialog);
    return;
  }
  if (type === "existing-connections") {
    void handleExistingConnectionsDialogAction(dialog, action);
    return;
  }
  if (type === "workbook-links") {
    void handleWorkbookLinksDialogAction(dialog, action);
    return;
  }
  if (type === "workbook-link-change") {
    if (action === "ok") applyWorkbookLinkSourceChange(dialog);
    return;
  }
  if (type === "refresh-status") {
    if (action === "refresh-status-select") {
      const sheetIndex = Number(event.currentTarget.dataset.querySheetIndex);
      const queryId = event.currentTarget.dataset.queryId || "";
      closeDataDialog();
      selectQueryRange(sheetIndex, queryId);
      return;
    }
    if (action === "refresh-status-properties") {
      const sheetIndex = Number(event.currentTarget.dataset.querySheetIndex);
      const queryId = event.currentTarget.dataset.queryId || "";
      openQueryPropertiesDialog(sheetIndex, queryId);
      return;
    }
    if (action === "refresh-status-refresh-all") {
      void refreshWorkbookData().then(() => openRefreshStatusDialog());
      return;
    }
    if (action === "refresh-status-cancel") {
      cancelWorkbookRefreshes();
      openRefreshStatusDialog();
      return;
    }
    if (action === "refresh-status-close" || action === "ok") {
      closeDataDialog();
      return;
    }
  }
  if (type === "import-data") {
    if (action === "import-properties") {
      const { sourceSheetIndex, queryId } = state.dataDialog || {};
      openQueryPropertiesDialog(sourceSheetIndex, queryId);
      return;
    }
    if (action === "ok") {
      applyImportDataDialog(dialog);
      return;
    }
  }
  if (type === "query-properties" && action === "ok") {
    applyQueryPropertiesDialog(dialog);
    return;
  }
  if (type === "query-properties" && action === "query-properties-clear-credential") {
    clearQueryPropertiesCredential(dialog);
    return;
  }
  if (type === "query-source-change" && action === "ok") {
    applyQuerySourceChangeDialog(dialog);
    return;
  }
  if (type === "query-rename" && action === "ok") {
    applyQueryRenameDialog(dialog);
    return;
  }
  if (type === "query-append" && action === "ok") {
    applyAppendWorkbookQueryDialog(dialog);
    return;
  }
  if (type === "query-merge" && action === "ok") {
    applyMergeWorkbookQueryDialog(dialog);
    return;
  }
  if (type === "query-properties" && action === "query-properties-tab") {
    const tab = event.currentTarget.dataset.queryPropertiesTab || "usage";
    dialog?.querySelectorAll("[data-query-properties-tab]").forEach((button) => {
      button.classList.toggle("is-active", button.dataset.queryPropertiesTab === tab);
      button.setAttribute("aria-selected", button.dataset.queryPropertiesTab === tab ? "true" : "false");
    });
    dialog?.querySelectorAll("[data-query-properties-tab-panel]").forEach((panel) => {
      panel.classList.toggle("is-active", panel.dataset.queryPropertiesTabPanel === tab);
    });
    return;
  }
  if (type === "auto-filter-custom") {
    if (action === "ok" || action === "clear") {
      applyCustomAutoFilterDialog(dialog, action === "clear");
    }
    return;
  }
  if (type === "auto-filter-top10") {
    if (action === "ok") applyTop10AutoFilterDialog(dialog);
    return;
  }
  if (type === "advanced-filter" && action === "ok") {
    applyAdvancedFilterDialog(dialog);
    return;
  }
  if (type === "validation") {
    if (action === "clear") {
      resetDataValidationDialog(dialog);
      return;
    }
    if (action === "ok") applyDataValidationDialog(dialog);
    return;
  }
  if (type === "relationships") {
    if (action === "relationship-new") {
      if (!canCreateWorkbookRelationship()) {
        showMessageBox(relationshipUnavailableMessage(), { title: "Cht WebSheet" });
        return;
      }
      openRelationshipEditDialog();
      return;
    }
    if (action === "relationship-edit") {
      const relationship = selectedRelationshipFromDialog(dialog);
      if (relationship) openRelationshipEditDialog(relationship.id);
      return;
    }
    if (action === "relationship-delete") {
      deleteRelationshipFromDialog(dialog);
      return;
    }
    if (action === "relationship-auto-detect") {
      if (!canCreateWorkbookRelationship()) {
        showMessageBox(relationshipUnavailableMessage(), { title: "Cht WebSheet" });
        return;
      }
      autoDetectRelationshipsFromDialog(dialog);
      return;
    }
  }
  if (type === "relationship-edit") {
    if (action === "relationship-ok") {
      applyRelationshipEditDialog(dialog);
      return;
    }
    if (action === "relationship-cancel") {
      openRelationshipsDialog(state.dataDialog?.relationshipId || "");
      return;
    }
  }
  if (type === "relationship-auto-detect") {
    if (action === "relationship-auto-ok" || action === "ok") {
      applyRelationshipAutoDetectDialog(dialog);
      return;
    }
    if (action === "relationship-auto-back") {
      openRelationshipsDialog();
      return;
    }
  }
  if (type === "data-model") {
    const table = selectedDataModelTableFromDialog(dialog);
	    if (action === "data-model-pivot") {
	      if (table) createPivotFromDataModelTable(table);
	      return;
	    }
	    if (action === "data-model-select") {
	      const ref = dataModelQueryRef(table);
	      closeDataDialog();
	      if (ref) selectQueryRange(ref.sheetIndex, ref.query.id);
	      return;
	    }
	    if (action === "data-model-load-to") {
	      const ref = dataModelQueryRef(table);
	      if (ref) openImportDataDialog(ref.sheetIndex, ref.query.id);
	      return;
	    }
	    if (action === "data-model-refresh") {
	      if (table) void refreshDataModelTable(table);
	      return;
	    }
    if (action === "data-model-refresh-all") {
      void refreshAllDataModelTables(table?.id || "");
      return;
    }
    if (action === "data-model-properties") {
      const ref = dataModelQueryRef(table);
      if (ref) openQueryPropertiesDialog(ref.sheetIndex, ref.query.id);
      return;
    }
    if (action === "data-model-relationships") {
      if (!canCreateWorkbookRelationship()) {
        showMessageBox(relationshipUnavailableMessage(), { title: "Cht WebSheet" });
        return;
      }
      openRelationshipsDialog();
      return;
    }
    if (action === "data-model-remove") {
      if (!table) return;
      showMessageBox(`テーブル「${table.name || ""}」をデータ モデルから削除しますか?`, {
        title: "Cht WebSheet",
        tone: "warning",
        buttons: [
          { label: "削除", result: "delete", callback: () => removeDataModelTable(table) },
          { label: "キャンセル", result: "cancel" },
        ],
        cancelResult: "cancel",
      });
      return;
    }
  }
  if (type === "subtotal") {
    if (action === "ok" || action === "remove-all") {
      applySubtotalDialog(dialog, action === "remove-all");
    }
    return;
  }
  if (type === "outline-group" || type === "outline-ungroup") {
    if (action === "ok") applyOutlineDialog(dialog);
    return;
  }
  if (type === "outline-settings") {
    if (action === "ok") applyOutlineSettingsDialog(dialog);
    return;
  }
  if (type === "what-if" && action === "ok") {
    applyWhatIfDialog(dialog);
    return;
  }
  if (type === "data-table" && action === "ok") {
    applyDataTableDialog(dialog);
    return;
  }
  if (type === "forecast-sheet" && action === "ok") {
    applyForecastSheetDialog(dialog);
    return;
  }
  if (type === "data-analysis-picker") {
    if (action === "data-analysis-help") {
      showDataAnalysisHelp(selectedDataAnalysisPickerMethod(dialog));
      return;
    }
    if (action === "data-analysis-picker-ok" || action === "ok") {
      const method = selectedDataAnalysisPickerMethod(dialog);
      openDataAnalysisOptionsDialog(method, state.dataDialog);
      return;
    }
  }
  if (type === "data-analysis") {
    if (action === "data-analysis-select") {
      state.dataDialog.method = event.currentTarget.dataset.dataAnalysisMethod || state.dataDialog?.method || "descriptive";
      syncDataAnalysisDialog(dialog);
      return;
    }
    if (action === "data-analysis-help") {
      persistDataAnalysisDialogState(dialog);
      showDataAnalysisHelp(state.dataDialog?.method);
      return;
    }
    if (action === "data-analysis-options-back") {
      const nextState = persistDataAnalysisDialogState(dialog);
      closeDataDialog({ keepState: true });
      const overlay = document.createElement("div");
      overlay.id = "dataDialog";
      overlay.className = "excel-dialog-overlay";
      overlay.innerHTML = dataAnalysisPickerDialogHtml(nextState.method || "descriptive");
      document.body.appendChild(overlay);
      syncDataAnalysisPickerDialog(overlay.querySelector("[data-data-dialog='data-analysis-picker']"));
      overlay.querySelector("[data-data-analysis-picker]")?.focus({ preventScroll: true });
      return;
    }
    if (action === "ok") {
      applyDataAnalysisDialog(dialog);
      return;
    }
  }
  if (type === "header-footer" && action === "ok") {
    applyHeaderFooterDialog(dialog);
    return;
  }
  if (type === "sheet-protection" && action === "ok") {
    applySheetProtectionDialog(dialog);
    return;
  }
  if (type === "sheet-protection-confirm" && action === "ok") {
    applySheetProtectionConfirmDialog(dialog);
    return;
  }
  if (type === "sheet-unprotect" && action === "ok") {
    applySheetUnprotectDialog(dialog);
    return;
  }
  if (type === "scenario-manager") {
    const scenario = selectedScenarioFromDialog(dialog);
    if (action === "scenario-add") {
      openScenarioEditDialog();
      return;
    }
    if (action === "scenario-edit") {
      if (scenario) openScenarioEditDialog(scenario);
      return;
    }
    if (action === "scenario-delete") {
      if (scenario) {
        if (!ensureScenarioPermission(scenario, "シナリオを削除")) return;
        showMessageBox(`シナリオ「${scenario.name}」を削除しますか?`, {
          title: "Cht WebSheet",
          tone: "warning",
          buttons: [
            {
              label: "削除",
              result: "delete",
              callback: () => {
                const scenariosBefore = cloneScenarios(workbookScenarios());
                state.model.scenarios = workbookScenarios().filter((item) => item.id !== scenario.id);
                pushHistory({
                  sheetIndex: state.activeSheetIndex,
                  scenariosBefore,
                  scenariosAfter: cloneScenarios(workbookScenarios()),
                  selectionBefore: activeSelectionRange(),
                  selectionAfter: activeSelectionRange(),
                  selectionBeforeRanges: activeSelectionRanges(),
                  selectionAfterRanges: activeSelectionRanges(),
                  label: `シナリオ「${scenario.name}」削除`,
                });
                scheduleWorkbookPersistence();
                openScenarioManagerDialog();
              },
            },
            { label: "キャンセル", result: "cancel" },
          ],
          cancelResult: "cancel",
        });
      }
      return;
    }
    if (action === "scenario-show") {
      closeDataDialog();
      applyScenario(scenario);
      return;
    }
    if (action === "scenario-summary") {
      openScenarioSummaryDialog();
      return;
    }
    return;
  }
  if (type === "scenario-summary" && action === "ok") {
    const validation = scenarioSummaryDialogValidation(dialog);
    if (!validation.valid) {
      showMessageBox(validation.message, { title: "Cht WebSheet" });
      syncScenarioSummaryDialog(dialog);
      return;
    }
    closeDataDialog();
    createScenarioSummarySheet({ resultCellsText: validation.resultCellsText, reportType: validation.reportType });
    return;
  }
  if (type === "scenario-edit") {
    if (action === "scenario-manager") {
      openScenarioManagerDialog();
      return;
    }
    if (action === "scenario-load-values") {
      loadScenarioCurrentValues(dialog);
      return;
    }
    if (action === "scenario-save") {
      saveScenarioFromDialog(dialog);
    }
  }
}



function shouldHandleDialogHistoryShortcut(event, key, isModifier) {
  if (!historyShortcutAction(event, key, isModifier)) return false;
  const dialog = document.querySelector("#dataDialog [data-data-dialog]");
  if (!dialog) return false;
  const target = event.target?.nodeType === Node.TEXT_NODE ? event.target.parentElement : event.target;
  if (!target?.closest?.("#dataDialog [data-data-dialog]")) return true;
  return !target.closest("input[type='text'], input[type='search'], input[type='number'], textarea, [contenteditable='true']");
}



function historyShortcutAction(event, key, isModifier = event?.ctrlKey || event?.metaKey) {
  if (!state.model || !isModifier || event?.altKey) return "";
  if (key === "y" && !event?.shiftKey) return "redo";
  if (key === "z" && event?.shiftKey) return "redo";
  if (key === "z") return "undo";
  return "";
}



function handleDataDialogKeydown(event) {
  if (isElementImeCompositionEvent(event)) return;
  const dialog = event.currentTarget;
  if (!dialog || dialog.closest("#messageBox")) return;
  if (event.key === "Escape" && document.body.classList.contains("is-websheet-fullscreen")) return;
  const key = String(event.key || "").toLowerCase();
  const isModifier = event.ctrlKey || event.metaKey;
  if (dialog.dataset.dataDialog === "function-arguments" && state.dataDialog?.referenceCollapsed && (event.key === "Enter" || event.key === "Escape")) {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    expandFunctionArgumentReferencePicker(dialog);
    return;
  }
  if (shouldHandleDialogHistoryShortcut(event, key, isModifier)) {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    if (historyShortcutAction(event, key, isModifier) === "undo") undoLastOperation();
    else redoLastOperation();
    return;
  }
  if (dialog.dataset.dataDialog === "macro-step" && (event.key === "F8" || event.key === "F5")) {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    executeMacroStepDialog(dialog, event.key === "F5" ? "continue" : "next");
    return;
  }
  if (event.key === "Escape") {
    event.preventDefault();
    const cancelButton = dialog.querySelector("[data-data-dialog-action='cancel']");
    if (cancelButton && !cancelButton.disabled) cancelButton.click();
    else closeDataDialog();
    return;
  }
  const dialogAccessKey = accessKeyFromKeyboardEvent(event);
  if (dialogAccessKey && event.altKey && !event.ctrlKey && !event.metaKey) {
    const target = findDialogAccessKeyTarget(dialog, dialogAccessKey);
    if (!target) return;
    event.preventDefault();
    activateDialogAccessKeyTarget(target);
    return;
  }
  if (event.key !== "Enter") return;
  const target = event.target;
  if (target?.closest?.("textarea, [contenteditable='true']")) return;
  if (target?.closest?.("button")) return;
  const primaryButton =
    dialog.querySelector("[data-data-dialog-action].is-primary:not(:disabled)") ||
    dialog.querySelector("[data-data-dialog-action='ok']:not(:disabled)") ||
    dialog.querySelector("[data-data-dialog-action$='-ok']:not(:disabled)");
  if (!primaryButton) return;
  event.preventDefault();
  primaryButton.click();
}



function handleGlobalDataDialogAccessKey(event) {
  if (isElementImeCompositionEvent(event) || event.defaultPrevented) return;
  if (event.target?.closest?.("#dataDialog [data-data-dialog]")) return;
  const accessKey = accessKeyFromKeyboardEvent(event);
  if (!accessKey || !event.altKey || event.ctrlKey || event.metaKey) return;
  const dialog = document.querySelector("#dataDialog [data-data-dialog]");
  if (!dialog) return;
  const target = findDialogAccessKeyTarget(dialog, accessKey);
  if (!target) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  activateDialogAccessKeyTarget(target);
}



function installDialogAccessKeyEnhancer() {
  if (state.dialogAccessKeyObserver || typeof MutationObserver !== "function") return;
  state.dialogAccessKeyObserver = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) enhanceDialogAccessKeys(node);
      });
    });
  });
  state.dialogAccessKeyObserver.observe(document.body, { childList: true });
  enhanceDialogAccessKeys(document.body);
}



function enhanceDialogAccessKeys(root) {
  const elements = root?.querySelectorAll
    ? [
      ...(root.matches?.("[data-data-dialog], #messageBox") ? [root] : []),
      ...root.querySelectorAll("[data-data-dialog], #messageBox"),
    ]
    : [];
  elements.forEach((dialog) => {
    dialog.querySelectorAll(".excel-mini-button, .message-box-button").forEach(enhanceButtonAccessKey);
    dialog.querySelectorAll(".data-dialog-check, .data-dialog-checkbox").forEach(enhanceChoiceLabelAccessKey);
    dialog.querySelectorAll(".excel-dialog-caption").forEach(enhanceCaptionAccessKey);
  });
}



function enhanceButtonAccessKey(button) {
  if (!button || button.dataset.accessKeyEnhanced === "true") return;
  const label = button.textContent.trim();
  const accessKey = menuAccessKeyFromLabel(label);
  if (accessKey) {
    button.dataset.accessKey = accessKey;
    button.innerHTML = accessKeyLabelHtml(label);
  }
  button.dataset.accessKeyEnhanced = "true";
}



function enhanceChoiceLabelAccessKey(label) {
  if (!label || label.dataset.accessKeyEnhanced === "true") return;
  const labelText = label.textContent.trim();
  const accessKey = menuAccessKeyFromLabel(labelText);
  if (accessKey) {
    label.dataset.accessKey = accessKey;
    if (!label.querySelector(".menu-access-key")) {
      const control = [...label.childNodes].find((node) =>
        node.nodeType === Node.ELEMENT_NODE && node.matches?.("input, select, textarea"),
      );
      const hasComplexContent = [...label.childNodes].some((node) =>
        node !== control && node.nodeType === Node.ELEMENT_NODE,
      );
      if (control && !hasComplexContent) {
        label.replaceChildren(control);
        label.insertAdjacentHTML("beforeend", accessKeyLabelHtml(labelText));
      }
    }
  }
  label.dataset.accessKeyEnhanced = "true";
}



function enhanceCaptionAccessKey(caption) {
  if (!caption || caption.dataset.accessKeyEnhanced === "true") return;
  const label = caption.textContent.trim();
  const accessKey = menuAccessKeyFromLabel(label);
  if (accessKey) {
    caption.dataset.accessKey = accessKey;
    caption.innerHTML = accessKeyLabelHtml(label);
    caption.closest(".excel-dialog-field")?.setAttribute("data-access-key", accessKey);
  }
  caption.dataset.accessKeyEnhanced = "true";
}



function findDialogAccessKeyTarget(dialog, key) {
  const accessKey = String(key || "").toUpperCase();
  return [...dialog.querySelectorAll("[data-access-key]")]
    .find((element) => String(element.dataset.accessKey || "").toUpperCase() === accessKey) || null;
}



function accessKeyFromKeyboardEvent(event) {
  const code = String(event?.code || "");
  const codeMatch = code.match(/^(?:Key([A-Z])|Digit([0-9])|Numpad([0-9]))$/i);
  if (codeMatch) return (codeMatch[1] || codeMatch[2] || codeMatch[3] || "").toUpperCase();
  const key = String(event?.key || "");
  return /^[A-Za-z0-9]$/.test(key) ? key.toUpperCase() : "";
}



function activateDialogAccessKeyTarget(target) {
  if (!target) return;
  if (target.matches?.("button")) {
    if (!target.disabled) target.click();
    return;
  }
  const control = target.matches?.("input, select, textarea")
    ? target
    : target.querySelector?.("input, select, textarea, button") ||
      target.closest?.("label")?.querySelector?.("input, select, textarea, button");
  if (!control) return;
  if (control.matches("button")) {
    if (!control.disabled) control.click();
    return;
  }
  if (control.matches("input[type='radio'], input[type='checkbox']")) {
    control.click();
    control.focus({ preventScroll: true });
    return;
  }
  control.focus({ preventScroll: true });
  if (typeof control.select === "function") control.select();
}



function closeDataDialog(options = {}) {
  document.getElementById("dataDialog")?.remove();
  if (!options.keepState) state.dataDialog = null;
}



function closeDataDialogOnEscape(event) {
  if (event.key !== "Escape" || !document.getElementById("dataDialog")) return;
  if (isElementImeCompositionEvent(event)) return;
  if (document.body.classList.contains("is-websheet-fullscreen")) return;
  if (document.querySelector("#messageBoxOverlay")) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  closeDataDialog();
}



function insertPivotSummaryFromSelection() {
  const range = singleInsertDataRange("ピボットテーブル");
  if (!range) return;
  if (insertPivotSummaryToSheetGroup(range)) return;
  const sourceSheet = activeSheet();
  const summary = pivotSummaryFromRange(sourceSheet, range);
  if (!summary.rows.length) {
    showMessageBox("ピボットテーブルを作成するためのデータが見つかりません。", { title: "Cht WebSheet" });
    return;
  }

  const finalName = uniqueSheetNameWithBase("PivotTable", state.model.sheets.map((item) => item.name));
  const sheet = createBlankSheet(finalName);
  ensureSheetSize(sheet, summary.rows.length + 2, summary.headers.length + 1);
  writePivotSummarySheet(sheet, summary);
  const activeSheetBefore = state.activeSheetIndex;
  const insertIndex = state.model.sheets.length;
  state.model.sheets.push(sheet);
  state.activeSheetIndex = insertIndex;
  state.selectedSheetIndexes = [];
  state.hf = buildFormulaEngine(state.model);
  const point = { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  state.selected = point;
  state.selectionRange = selectionRangeFromPoints(point, point);
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = point;
  state.selectionAnchor = point;
  state.selectedImage = null;
  state.formulaEdit = null;
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  setStatus(`${finalName} にピボットテーブルを作成しました。`);
  pushSheetInsertHistory(insertIndex, sheet, activeSheetBefore, insertIndex, `${finalName} ピボットテーブル`);
  scheduleWorkbookPersistence();
}



function insertPivotSummaryToSheetGroup(range) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !range || !state.model) return false;
  const insertIndex = state.model.sheets.length;
  const existingNames = state.model.sheets.map((sheet) => sheet.name);
  const insertedSheets = [];
  const sourceNames = [];
  for (const sheetIndex of sheetIndexes) {
    const sourceSheet = state.model.sheets[sheetIndex];
    if (!sourceSheet) continue;
    const sourceRange = materializeSelectionRangeForSheet(sourceSheet, range, sheetIndex);
    if (!sourceRange || sourceRange.fullRow || sourceRange.fullColumn) continue;
    const summary = pivotSummaryFromRange(sourceSheet, sourceRange);
    if (!summary.rows.length) {
      showMessageBox("ピボットテーブルを作成するためのデータが見つかりません。", { title: "Cht WebSheet" });
      return true;
    }
    const nameBase = `PivotTable_${sourceSheet.name}`;
    const finalName = uniqueSheetNameWithBase(nameBase.slice(0, 24) || "PivotTable", [...existingNames, ...insertedSheets.map((sheet) => sheet.name)]);
    const sheet = createBlankSheet(finalName);
    ensureSheetSize(sheet, summary.rows.length + 2, summary.headers.length + 1);
    writePivotSummarySheet(sheet, summary, insertIndex + insertedSheets.length);
    insertedSheets.push(sheet);
    sourceNames.push(sourceSheet.name);
  }
  if (!insertedSheets.length) return true;

  const activeSheetBefore = state.activeSheetIndex;
  state.model.sheets.push(...insertedSheets);
  state.activeSheetIndex = insertIndex;
  state.selectedSheetIndexes = insertedSheets.map((_sheet, offset) => insertIndex + offset);
  state.hf = buildFormulaEngine(state.model);
  const point = { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  state.selected = point;
  state.selectionRange = selectionRangeFromPoints(point, point);
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = point;
  state.selectionAnchor = point;
  state.selectedImage = null;
  state.formulaEdit = null;
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  setStatus(`${insertedSheets.length} 個のピボットテーブルを作成しました：${sourceNames.join("、")}`);
  pushSheetInsertManyHistory(insertIndex, insertedSheets, activeSheetBefore, insertIndex, `${insertedSheets.length} シートピボットテーブル`, {
    selectionAfter: state.selectionRange,
    selectionAfterRanges: [state.selectionRange],
  });
  scheduleWorkbookPersistence();
  return true;
}



function writePivotSummarySheet(sheet, summary, sheetIndex = state.activeSheetIndex) {
  return writePivotSummaryAt(sheet, summary, 1, 1, sheetIndex);
}



function writePivotSummaryAt(sheet, summary, startRow = 1, startCol = 1, sheetIndex = state.activeSheetIndex) {
  const headers = summary.headers;
  ensureSheetSize(sheet, startRow + summary.rows.length, startCol + headers.length - 1);
  headers.forEach((header, index) => updateCellModel(sheet, startRow, startCol + index, header));
  summary.rows.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => updateCellModel(sheet, startRow + rowIndex + 1, startCol + colIndex, value));
  });
  const range = normalizeRange(
    { sheetIndex, row: startRow, col: startCol },
    { sheetIndex, row: startRow + summary.rows.length, col: startCol + headers.length - 1 },
  );
  applyExcelTableStyle(sheet, range);
  sheet.tables = Array.isArray(sheet.tables) ? sheet.tables : [];
  sheet.tables.push({
    id: uniqueWorksheetObjectId("pivot"),
    name: nextPivotTableName(sheet),
    range,
    showHeaderRow: true,
    showFilterButton: true,
    style: "PivotStyleMedium9",
  });
  sheet.columns[startCol - 1].width = Math.max(sheet.columns[startCol - 1].width, 140);
  for (let col = startCol + 1; col <= range.right; col += 1) {
    sheet.columns[col - 1].width = Math.max(sheet.columns[col - 1].width, 96);
  }
  internCellsInRange(sheet, range);
  return range;
}



function nextPivotTableName(sheet) {
  const existing = new Set((sheet.tables || []).map((table) => String(table.name || "")));
  let index = 1;
  let name = `PivotTable${index}`;
  while (existing.has(name)) {
    index += 1;
    name = `PivotTable${index}`;
  }
  return name;
}



function pivotSummaryFromRange(sheet, range) {
  const sheetId = getSheetId(sheet.name);
  const headers = [];
  for (let col = range.left; col <= range.right; col += 1) {
    headers.push(String(getDisplayForCell(sheet, sheetId, range.top, col) || `列${col - range.left + 1}`));
  }
  const labelHeader = headers[0] || "行ラベル";
  const numericCols = headers
    .map((header, index) => ({ header, index }))
    .filter(({ index }) => {
      for (let row = range.top + 1; row <= range.bottom; row += 1) {
        if (Number.isFinite(cellNumericValue(sheet, sheetId, row, range.left + index))) return true;
      }
      return false;
    });
  const valueCol = numericCols.find((item) => item.index > 0) || numericCols[0] || { header: "データの個数", index: -1 };
  const groups = new Map();
  for (let row = range.top + 1; row <= range.bottom; row += 1) {
    const label = String(getDisplayForCell(sheet, sheetId, row, range.left) || "(空白)");
    const group = groups.get(label) || { label, count: 0, sum: 0 };
    group.count += 1;
    const value = valueCol.index >= 0 ? cellNumericValue(sheet, sheetId, row, range.left + valueCol.index) : null;
    if (Number.isFinite(value)) group.sum += value;
    groups.set(label, group);
  }
  const rows = Array.from(groups.values())
    .sort((left, right) => String(left.label).localeCompare(String(right.label), "ja"))
    .map((group) => [group.label, group.count, group.sum]);
  return {
    headers: [labelHeader || "行ラベル", "データの個数", `合計 / ${valueCol.header || "値"}`],
    rows,
  };
}



function uniqueSheetNameWithBase(base, existingNames) {
  const existing = new Set(existingNames);
  let index = 1;
  let name = `${base}${index}`;
  while (existing.has(name)) {
    index += 1;
    name = `${base}${index}`;
  }
  return name;
}



function recommendedChartTypeForSelection() {
  const range = activeSelectionRange();
  if (!range) return "column";
  const rows = range.bottom - range.top + 1;
  const cols = range.right - range.left + 1;
  if (rows <= 2 && cols > 3) return "line";
  if (rows <= 8 && cols <= 2) return "pie";
  return "column";
}



function insertChartFromSelection(type, options = {}) {
  const range = singleInsertDataRange(options.pivot ? "ピボットグラフ" : "グラフ");
  if (!range) return;
  const sheet = activeSheet();
  const chart = chartDataFromRange(sheet, range, type);
  if (!chart.series.length || !chart.series.some((series) => series.values.some((value) => Number.isFinite(value)))) {
    showMessageBox("グラフを作成するための数値データが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const chartType = type === "recommended" ? recommendedChartTypeForSelection() : type;
  if (insertChartToSheetGroup(chartType, range, options)) return;
  const descriptor = chartDescriptor(chart, chartType, range, options);
  const anchor = {
    sheetIndex: state.activeSheetIndex,
    row: Math.max(1, range.top),
    col: range.right + 2,
  };
  insertWorksheetImageDescriptors([descriptor], {
    label: options.pivot ? "ピボットグラフ" : "グラフ",
    anchor,
  });
}



function insertChartToSheetGroup(chartType, range, options = {}) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !range) return false;
  const label = options.pivot ? "ピボットグラフ" : "グラフ";
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを挿入できません。")) return true;
    const sourceRange = materializeSelectionRangeForSheet(sheet, range, sheetIndex);
    if (!sourceRange || sourceRange.fullRow || sourceRange.fullColumn) continue;
    const chart = chartDataFromRange(sheet, sourceRange, chartType);
    if (!chart.series.length || !chart.series.some((series) => series.values.some((value) => Number.isFinite(value)))) {
      showMessageBox("グラフを作成するための数値データが見つかりません。", { title: "Cht WebSheet" });
      return true;
    }
    const descriptor = chartDescriptor(chart, chartType, sourceRange, {
      ...options,
      sourceSheet: sheet,
      sourceSheetName: sheet.name,
    });
    const anchor = {
      sheetIndex,
      row: Math.max(1, sourceRange.top),
      col: sourceRange.right + 2,
    };
    items.push({ sheetIndex, sheet, sourceRange, descriptor, anchor });
  }
  if (!items.length) return true;

  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((item) => ({ ...item }));
  const activeItem = items.find((item) => item.sheetIndex === state.activeSheetIndex) || items[0];
  try {
    let activeInserted = [];
    items.forEach((item) => {
      const inserted = addWorksheetImageDescriptorsToSheet(item.sheet, [item.descriptor], item.anchor, { label });
      if (item.sheetIndex === state.activeSheetIndex) activeInserted = inserted;
    });
    if (!activeInserted.length) throw new Error("No chart was inserted on the active sheet.");
    state.selected = null;
    state.selectionRange = null;
    state.selectionRanges = [];
    state.selectionAnchor = null;
    state.selectionStart = null;
    state.selectionDrag = null;
    state.formulaEdit = null;
    setSelectedImageIds(activeInserted.map((image) => image.id), activeInserted[0].id);
    $formulaInput.prop("disabled", true).val("");
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore,
      selectionAfter: null,
      selectionBeforeRanges,
      selectionAfterRanges: [],
      label: `${rangeToLabel(range)} ${sheetIndexes.length} シート${label}`,
    });
    renderSheet();
    updateSelectionUi();
    setStatus(`${sheetIndexes.length} 個のシートに${label}を挿入しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`${label}を挿入できませんでした：${error.message || error}`);
  }
  return true;
}



function insertChartSheetFromSelection(type = recommendedChartTypeForSelection(), options = {}) {
  const range = singleInsertDataRange("グラフ");
  if (!range || !state.model) return;
  const sourceSheet = activeSheet();
  const chartType = type === "recommended" ? recommendedChartTypeForSelection() : type;
  const chart = chartDataFromRange(sourceSheet, range, chartType);
  if (!chart.series.length || !chart.series.some((series) => series.values.some((value) => Number.isFinite(value)))) {
    showMessageBox("グラフを作成するための数値データが見つかりません。", { title: "Cht WebSheet" });
    return;
  }

  const requestedIndex = Number(options.insertIndex);
  const insertIndex = Number.isFinite(requestedIndex)
    ? Math.max(0, Math.min(Math.round(requestedIndex), state.model.sheets.length))
    : state.activeSheetIndex + 1;
  const activeSheetBefore = state.activeSheetIndex;
  const name = uniqueSheetNameWithBase("グラフ", state.model.sheets.map((sheet) => sheet.name));
  const sheet = createBlankSheet(name);
  sheet.chartSheet = true;
  sheet.showGridLines = false;
  sheet.showHeadings = false;
  sheet.baseRowCount = Math.max(sheet.baseRowCount || 1, 24);
  sheet.baseColCount = Math.max(sheet.baseColCount || 1, 12);
  ensureSheetSize(sheet, 24, 12);

  const width = 760;
  const height = 440;
  const title = `${rangeToLabel(range)} ${chartTypeLabel(chartType)}`;
  const image = {
    id: uniqueWorksheetObjectId("chart"),
    name: title,
    src: svgDataUrl(chartSvg(chart, chartType, { width, height, title })),
    mimeType: "image/svg+xml",
    extension: "svg",
    editAs: "oneCell",
    tl: {
      col: 1,
      row: 1,
      nativeCol: 1,
      nativeRow: 1,
      nativeColOff: 0,
      nativeRowOff: 0,
    },
    br: null,
    ext: { width, height },
    hyperlink: "",
    chart: {
      ...cloneJsonLikeValue(chart),
      type: chartType,
      sourceSheet: sourceSheet.name,
      sourceRange: rangeToLabel(range),
      sourceRef: absoluteRangeReference(sourceSheet, range),
      title,
    },
  };
  sheet.images = [image];
  const end = endingCellForPixelSize(sheet, 2, 2, width, height);
  ensureSheetSize(sheet, end.row + 1, end.col + 1);

  state.model.sheets.splice(insertIndex, 0, sheet);
  state.activeSheetIndex = insertIndex;
  state.selectedSheetIndexes = [];
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  state.selectionDrag = null;
  state.formulaEdit = null;
  setSelectedImageIds([image.id], image.id);
  state.hf = buildFormulaEngine(state.model);
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  updateSelectionUi();
  pushSheetInsertHistory(insertIndex, sheet, activeSheetBefore, insertIndex, `${name} グラフ シート作成`, {
    selectionAfter: null,
    selectionAfterRanges: [],
  });
  setStatus(`グラフ シートを作成しました：${name}`);
}



function chartDescriptor(chart, type, range, options = {}) {
  const width = Math.max(1, Math.round(Number(options.width) || 520));
  const height = Math.max(1, Math.round(Number(options.height) || 320));
  const title = Object.prototype.hasOwnProperty.call(options, "title") && String(options.title || "").trim()
    ? String(options.title).trim()
    : options.pivot ? "ピボットグラフ" : `${rangeToLabel(range)} ${chartTypeLabel(type)}`;
  const svg = chartSvg(chart, type, { width, height, title });
  const sourceSheet = options.sourceSheet || activeSheet();
  return {
    name: title,
    src: svgDataUrl(svg),
    mimeType: "image/svg+xml",
    extension: "svg",
    width,
    height,
    chart: {
      ...cloneJsonLikeValue(chart),
      type,
      sourceSheet: options.sourceSheetName || sourceSheet.name,
      sourceRange: rangeToLabel(range),
      sourceRef: absoluteRangeReference(sourceSheet, range),
      title,
    },
  };
}



function chartDataFromRange(sheet, range, type) {
  const sheetId = getSheetId(sheet.name);
  const rows = range.bottom - range.top + 1;
  const cols = range.right - range.left + 1;
  const firstRow = [];
  const firstCol = [];
  for (let col = range.left; col <= range.right; col += 1) firstRow.push(getDisplayForCell(sheet, sheetId, range.top, col));
  for (let row = range.top; row <= range.bottom; row += 1) firstCol.push(getDisplayForCell(sheet, sheetId, row, range.left));
  const hasHeaderRow = rows > 1 && firstRow.some((value) => value !== "" && !Number.isFinite(Number(value)));
  const hasLabelCol = cols > 1 && firstCol.slice(hasHeaderRow ? 1 : 0).some((value) => value !== "" && !Number.isFinite(Number(value)));
  const labelRows = { start: range.top + (hasHeaderRow ? 1 : 0), end: range.bottom };
  const valueCols = { start: range.left + (hasLabelCol ? 1 : 0), end: range.right };
  const labels = [];
  for (let row = labelRows.start; row <= labelRows.end; row += 1) {
    const label = hasLabelCol ? getDisplayForCell(sheet, sheetId, row, range.left) : `${row - labelRows.start + 1}`;
    labels.push(String(label || `${row - labelRows.start + 1}`));
  }
  const labelsRef = hasLabelCol
    ? absoluteRangeReference(sheet, { top: labelRows.start, left: range.left, bottom: labelRows.end, right: range.left })
    : "";
  const series = [];
  for (let col = valueCols.start; col <= valueCols.end; col += 1) {
    const values = [];
    for (let row = labelRows.start; row <= labelRows.end; row += 1) {
      values.push(cellNumericValue(sheet, sheetId, row, col));
    }
    if (!values.some((value) => Number.isFinite(value))) continue;
    const name = hasHeaderRow ? String(getDisplayForCell(sheet, sheetId, range.top, col) || columnName(col)) : columnName(col);
    series.push({
      name,
      values,
      ...(hasHeaderRow ? { nameRef: absoluteRangeReference(sheet, { top: range.top, left: col, bottom: range.top, right: col }) } : {}),
      valuesRef: absoluteRangeReference(sheet, { top: labelRows.start, left: col, bottom: labelRows.end, right: col }),
    });
  }
  const chart = { labels, ...(labelsRef ? { labelsRef } : {}), series };
  if (type === "pie" && series.length > 1) return { ...chart, series: [series[0]] };
  return chart;
}



function cellNumericValue(sheet, sheetId, row, col) {
  const raw = sheet.cells[cellKey(row, col)]?.raw;
  const value = raw != null ? raw : getDisplayForCell(sheet, sheetId, row, col);
  const number = typeof value === "number" ? value : Number(String(value ?? "").replace(/[,¥$€£%]/g, ""));
  return Number.isFinite(number) ? number : null;
}



function singleInsertDataRange(label) {
  const ranges = activeSelectionRanges();
  const range = ranges[0] || activeSelectionRange();
  if (!range) {
    setStatus(`${label}に使うセル範囲を選択してください。`);
    return null;
  }
  if (ranges.length > 1 || range.fullRow || range.fullColumn) {
    showMessageBox(`${label}では複数の選択範囲または行/列全体を使用できません。`, { title: "Cht WebSheet" });
    return null;
  }
  return materializeSelectionRange(range);
}



function chartSvg(chart, type, options = {}) {
  const width = options.width || 520;
  const height = options.height || 320;
  const legendModel = normalizedChartLegend(chart);
  const orientation = chartOrientationForType(type);
  const plot = chartPlotBox(width, height, legendModel, chart, orientation);
  const chartArea = chartAreaSvgStyle(chart?.chartArea, { fill: "#ffffff", stroke: "#b8c2cc", strokeWidth: 1 });
  const plotArea = chartAreaSvgStyle(chart?.plotArea, { fill: "none", stroke: "none", strokeWidth: 0 });
  const titleText = Object.prototype.hasOwnProperty.call(options, "title") ? options.title : chartTypeLabel(type);
  const title = String(titleText || "").trim()
    ? `<text data-chart-title="chart" x="${width / 2}" y="25" text-anchor="middle" ${chartTextSvgAttrs(chart?.titleStyle, { fontSize: 16, fontWeight: "600", fill: "#1f2937" })}>${escapeXml(titleText)}</text>`
    : "";
  const content = type === "line"
    ? lineChartSvg(chart, plot)
    : type === "pie"
      ? pieChartSvg(chart, plot)
      : type === "scatter"
        ? scatterChartSvg(chart, plot)
        : type === "map"
          ? mapChartSvg(chart, plot)
        : type === "bar"
          ? barChartSvg(chart, plot)
        : columnChartSvg(chart, plot);
  const legend = legendModel ? chartLegendSvg(chart, width, height, plot, legendModel) : "";
  const axisTitles = chartAxisTitlesSvg(chart, plot, width, height, orientation);
  const plotAreaRect = chartPlotAreaRectSvg(plot, plotArea);
  return `
    <svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
      <rect width="${width}" height="${height}" fill="${escapeXml(chartArea.fill)}" stroke="${escapeXml(chartArea.stroke)}" stroke-width="${roundSvg(chartArea.strokeWidth)}"${chartArea.dashAttr}/>
      ${title}
      ${plotAreaRect}
      ${content}
      ${axisTitles}
      ${legend}
    </svg>
  `.trim();
}



function mapChartSvg(chart, plot) {
  const values = (chart.series?.[0]?.values || []).map((value) => (Number.isFinite(value) ? value : 0));
  const labels = chart.labels?.length ? chart.labels : values.map((_value, index) => `地域 ${index + 1}`);
  const max = Math.max(1, ...values.map((value) => Math.abs(value)));
  const regions = [
    { x: 0.58, y: 0.07, w: 0.2, h: 0.2, rx: 0.05 },
    { x: 0.45, y: 0.25, w: 0.26, h: 0.18, rx: 0.04 },
    { x: 0.33, y: 0.42, w: 0.27, h: 0.18, rx: 0.04 },
    { x: 0.2, y: 0.58, w: 0.26, h: 0.18, rx: 0.04 },
    { x: 0.09, y: 0.72, w: 0.2, h: 0.14, rx: 0.04 },
    { x: 0.68, y: 0.62, w: 0.16, h: 0.14, rx: 0.04 },
  ];
  const colorForValue = (value) => {
    const intensity = Math.min(1, Math.abs(value) / max);
    const blue = Math.round(224 - intensity * 116);
    const green = Math.round(238 - intensity * 96);
    const red = Math.round(242 - intensity * 156);
    return `rgb(${red},${green},${blue})`;
  };
  const regionSvg = regions.map((region, index) => {
    const value = values[index % Math.max(1, values.length)] || 0;
    const label = labels[index % Math.max(1, labels.length)] || `地域 ${index + 1}`;
    const x = plot.left + region.x * plot.width;
    const y = plot.top + region.y * plot.height;
    const w = region.w * plot.width;
    const h = region.h * plot.height;
    const rx = region.rx * Math.min(plot.width, plot.height);
    return `
      <rect data-map-region="${escapeXml(label)}" x="${roundSvg(x)}" y="${roundSvg(y)}" width="${roundSvg(w)}" height="${roundSvg(h)}" rx="${roundSvg(rx)}" fill="${escapeXml(colorForValue(value))}" stroke="#5b6778" stroke-width="1.2"/>
      <text x="${roundSvg(x + w / 2)}" y="${roundSvg(y + h / 2 + 4)}" text-anchor="middle" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="11" fill="#1f2937">${escapeXml(label)}</text>
    `;
  }).join("");
  const legendX = plot.left + plot.width - 110;
  const legendY = plot.top + plot.height - 40;
  return `
    <g data-chart-map="true">
      <rect x="${roundSvg(plot.left)}" y="${roundSvg(plot.top)}" width="${roundSvg(plot.width)}" height="${roundSvg(plot.height)}" fill="#f8fafc" stroke="#d8e0ea" stroke-width="1"/>
      ${regionSvg}
      <rect x="${roundSvg(legendX)}" y="${roundSvg(legendY)}" width="22" height="12" fill="${escapeXml(colorForValue(max * 0.25))}" stroke="#94a3b8"/>
      <rect x="${roundSvg(legendX + 25)}" y="${roundSvg(legendY)}" width="22" height="12" fill="${escapeXml(colorForValue(max * 0.6))}" stroke="#94a3b8"/>
      <rect x="${roundSvg(legendX + 50)}" y="${roundSvg(legendY)}" width="22" height="12" fill="${escapeXml(colorForValue(max))}" stroke="#94a3b8"/>
      <text x="${roundSvg(legendX)}" y="${roundSvg(legendY + 28)}" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="10" fill="#526173">小</text>
      <text x="${roundSvg(legendX + 60)}" y="${roundSvg(legendY + 28)}" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="10" fill="#526173">大</text>
    </g>
  `;
}



function chartOrientationForType(type) {
  return type === "bar" ? "bar" : "column";
}



function chartAreaSvgStyle(style = {}, fallback = {}) {
  const fill = style?.fillColor === "none" ? "none" : cssColorForSvg(style?.fillColor) || fallback.fill || "none";
  const stroke = style?.outlineColor === "none" ? "none" : cssColorForSvg(style?.outlineColor) || fallback.stroke || "none";
  const strokeWidth = stroke === "none" ? 0 : Number.isFinite(Number(style?.outlineWidth)) ? normalizeShapeOutlineWidth(style.outlineWidth) : Number(fallback.strokeWidth) || 0;
  const dash = stroke === "none" ? "" : shapeLineDashArray(style?.lineDash, strokeWidth);
  return {
    fill,
    stroke,
    strokeWidth,
    dashAttr: dash ? ` stroke-dasharray="${escapeXml(dash)}"` : "",
  };
}



function chartPlotAreaRectSvg(plot, style) {
  if (style.fill === "none" && style.stroke === "none") return "";
  return `<rect x="${roundSvg(plot.left)}" y="${roundSvg(plot.top)}" width="${roundSvg(plot.right - plot.left)}" height="${roundSvg(plot.bottom - plot.top)}" fill="${escapeXml(style.fill)}" stroke="${escapeXml(style.stroke)}" stroke-width="${roundSvg(style.strokeWidth)}"${style.dashAttr}/>`;
}



function normalizedChartLegend(chart) {
  if (chart?.legend?.visible === false) return null;
  const position = ["left", "right", "top", "bottom"].includes(chart?.legend?.position) ? chart.legend.position : "bottom";
  if (!chart?.series?.length) return null;
  return { position };
}



function chartPlotBox(width, height, legendModel, chart = null, orientation = "column") {
  const hasVerticalTitle = orientation === "bar" ? Boolean(chart?.axes?.category?.title) : Boolean(chart?.axes?.value?.title);
  const hasHorizontalTitle = orientation === "bar" ? Boolean(chart?.axes?.value?.title) : Boolean(chart?.axes?.category?.title);
  const plot = {
    left: 58 + (hasVerticalTitle ? 22 : 0),
    top: 48,
    right: width - 24,
    bottom: height - 62 - (hasHorizontalTitle ? 18 : 0),
  };
  if (!legendModel) return plot;
  if (legendModel.position === "right") plot.right = Math.max(plot.left + 80, width - 132);
  if (legendModel.position === "left") plot.left = Math.min(plot.right - 80, 144);
  if (legendModel.position === "top") plot.top = 66;
  if (legendModel.position === "bottom") plot.bottom = Math.max(plot.top + 70, height - 62 - (hasHorizontalTitle ? 18 : 0));
  return plot;
}



function columnChartSvg(chart, plot) {
  const stacked = chartIsStacked(chart);
  const percent = chart?.grouping === "percentStacked";
  const max = stacked ? (percent ? 1 : chartMaxStackValue(chart)) : chartMaxValue(chart);
  const groupWidth = (plot.right - plot.left) / Math.max(1, chart.labels.length);
  const barGap = 4;
  const barWidth = stacked
    ? Math.max(4, groupWidth - barGap * 2)
    : Math.max(4, (groupWidth - barGap * 2) / Math.max(1, chart.series.length));
  const bars = [];
  const labels = [];
  const stackTotals = stacked ? chartStackTotals(chart) : [];
  const accumulated = new Array(chart.labels.length).fill(0);
  chart.series.forEach((series, seriesIndex) => {
    series.values.forEach((value, index) => {
      const safeValue = Math.max(0, Number(value) || 0);
      const plottedValue = percent ? safeValue / Math.max(1, stackTotals[index] || 0) : safeValue;
      const height = max ? (plottedValue / max) * (plot.bottom - plot.top) : 0;
      const base = stacked ? accumulated[index] : 0;
      const baseHeight = stacked && max ? (base / max) * (plot.bottom - plot.top) : 0;
      const x = plot.left + index * groupWidth + barGap + (stacked ? 0 : seriesIndex * barWidth);
      const y = plot.bottom - height;
      const stackedY = plot.bottom - baseHeight - height;
      if (stacked) accumulated[index] += plottedValue;
      bars.push(`<rect${stacked ? ` data-chart-grouping="${escapeAttr(chart.grouping)}"` : ""} x="${roundSvg(x)}" y="${roundSvg(stacked ? stackedY : y)}" width="${roundSvg(Math.max(1, barWidth - 1))}" height="${roundSvg(height)}" fill="${escapeXml(chartPointColor(series, seriesIndex, index))}"/>`);
      const label = chartDataLabelText(chart, series, value, index);
      if (label) {
        const point = columnDataLabelPoint(plot, x, stacked ? stackedY : y, Math.max(1, barWidth - 1), height, chartDataLabelModel(chart, series)?.position);
        labels.push(`<text data-chart-data-label="value" data-chart-data-label-position="${escapeAttr(chartDataLabelModel(chart, series)?.position || "outEnd")}" x="${roundSvg(point.x)}" y="${roundSvg(point.y)}" text-anchor="${escapeAttr(point.anchor)}" ${chartTextSvgAttrs(chartDataLabelModel(chart, series)?.textStyle, { fontSize: 9, fill: "#1f2937" })}>${escapeXml(label)}</text>`);
      }
    });
  });
  return `${chartAxesSvg(plot, chart)}${bars.join("")}${labels.join("")}${chartLabelsSvg(chart.labels, plot)}`;
}



function barChartSvg(chart, plot) {
  const stacked = chartIsStacked(chart);
  const percent = chart?.grouping === "percentStacked";
  const max = stacked ? (percent ? 1 : chartMaxStackValue(chart)) : chartMaxValue(chart);
  const groupHeight = (plot.bottom - plot.top) / Math.max(1, chart.labels.length);
  const barGap = 4;
  const barHeight = stacked
    ? Math.max(4, groupHeight - barGap * 2)
    : Math.max(4, (groupHeight - barGap * 2) / Math.max(1, chart.series.length));
  const bars = [];
  const labels = [];
  const stackTotals = stacked ? chartStackTotals(chart) : [];
  const accumulated = new Array(chart.labels.length).fill(0);
  chart.series.forEach((series, seriesIndex) => {
    series.values.forEach((value, index) => {
      const safeValue = Math.max(0, Number(value) || 0);
      const plottedValue = percent ? safeValue / Math.max(1, stackTotals[index] || 0) : safeValue;
      const width = max ? (plottedValue / max) * (plot.right - plot.left) : 0;
      const base = stacked ? accumulated[index] : 0;
      const baseWidth = stacked && max ? (base / max) * (plot.right - plot.left) : 0;
      const x = plot.left + baseWidth;
      const y = plot.top + index * groupHeight + barGap + (stacked ? 0 : seriesIndex * barHeight);
      if (stacked) accumulated[index] += plottedValue;
      bars.push(`<rect data-chart-bar-orientation="horizontal"${stacked ? ` data-chart-grouping="${escapeAttr(chart.grouping)}"` : ""} x="${roundSvg(x)}" y="${roundSvg(y)}" width="${roundSvg(width)}" height="${roundSvg(Math.max(1, barHeight - 1))}" fill="${escapeXml(chartPointColor(series, seriesIndex, index))}"/>`);
      const label = chartDataLabelText(chart, series, value, index);
      if (label) {
        const point = barDataLabelPoint(plot, x, y, width, Math.max(1, barHeight - 1), chartDataLabelModel(chart, series)?.position);
        labels.push(`<text data-chart-data-label="value" data-chart-data-label-position="${escapeAttr(chartDataLabelModel(chart, series)?.position || "outEnd")}" x="${roundSvg(point.x)}" y="${roundSvg(point.y)}" text-anchor="${escapeAttr(point.anchor)}" ${chartTextSvgAttrs(chartDataLabelModel(chart, series)?.textStyle, { fontSize: 9, fill: "#1f2937" })}>${escapeXml(label)}</text>`);
      }
    });
  });
  return `${chartAxesSvg(plot, chart, {}, "bar")}${bars.join("")}${labels.join("")}${chartCategoryLabelsSvg(chart.labels, plot, "bar")}`;
}



function columnDataLabelPoint(plot, x, y, width, height, position = "") {
  const normalized = String(position || "outEnd").toLowerCase();
  const centerX = x + width / 2;
  if (normalized === "ctr" || normalized === "center") return { x: centerX, y: y + height / 2 + 3, anchor: "middle" };
  if (normalized === "inbase" || normalized === "insidebase") return { x: centerX, y: plot.bottom - 4, anchor: "middle" };
  if (normalized === "inend" || normalized === "insideend") return { x: centerX, y: Math.max(plot.top + 10, y + 11), anchor: "middle" };
  return { x: centerX, y: Math.max(plot.top + 10, y - 4), anchor: "middle" };
}



function barDataLabelPoint(plot, x, y, width, height, position = "") {
  const normalized = String(position || "outEnd").toLowerCase();
  const centerY = y + height / 2 + 3;
  if (normalized === "ctr" || normalized === "center") return { x: x + width / 2, y: centerY, anchor: "middle" };
  if (normalized === "inbase" || normalized === "insidebase") return { x: x + 4, y: centerY, anchor: "start" };
  if (normalized === "inend" || normalized === "insideend") return { x: Math.max(plot.left + 4, x + width - 4), y: centerY, anchor: "end" };
  return { x: Math.min(plot.right - 4, x + width + 4), y: centerY, anchor: "start" };
}



function lineChartSvg(chart, plot) {
  const max = chartMaxValue(chart);
  const min = chartMinValue(chart);
  const span = max === min ? 1 : max - min;
  const xFor = (index) => plot.left + (chart.labels.length <= 1 ? 0 : (index / (chart.labels.length - 1)) * (plot.right - plot.left));
  const yFor = (value) => plot.bottom - (((Number(value) || 0) - min) / span) * (plot.bottom - plot.top);
  const lines = chart.series.map((series, seriesIndex) => {
    const color = chartSeriesColor(series, seriesIndex);
    const segments = [];
    let current = [];
    series.values.forEach((value, index) => {
      const numeric = Number(value);
      if (Number.isFinite(numeric)) {
        current.push({ index, value: numeric });
      } else if (current.length) {
        segments.push(current);
        current = [];
      }
    });
    if (current.length) segments.push(current);
    const polylines = segments
      .map((segment) => {
        const points = segment.map((point) => `${roundSvg(xFor(point.index))},${roundSvg(yFor(point.value))}`).join(" ");
        return `<polyline points="${points}" fill="none" stroke="${escapeXml(color)}" stroke-width="2.2"/>`;
      })
      .join("");
    const markers = series.values
      .map((value, index) => {
        const numeric = Number(value);
        return Number.isFinite(numeric) ? chartMarkerSvg(series, seriesIndex, index, xFor(index), yFor(numeric), color) : "";
      })
      .join("");
    const labels = series.values
      .map((value, index) => {
        const numeric = Number(value);
        if (!Number.isFinite(numeric)) return "";
        const label = chartDataLabelText(chart, series, value, index);
        return label
          ? `<text data-chart-data-label="value" x="${roundSvg(xFor(index))}" y="${roundSvg(Math.max(plot.top + 10, yFor(numeric) - 8))}" text-anchor="middle" ${chartTextSvgAttrs(chartDataLabelModel(chart, series)?.textStyle, { fontSize: 9, fill: "#1f2937" })}>${escapeXml(label)}</text>`
          : "";
      })
      .join("");
    return `${polylines}${markers}${labels}`;
  });
  return `${chartAxesSvg(plot, chart, { min, max })}${lines.join("")}${chartLabelsSvg(chart.labels, plot)}`;
}



function scatterChartSvg(chart, plot) {
  const rawSeries = Array.isArray(chart?.series) ? chart.series : [];
  const series = rawSeries.some((item) => Array.isArray(item?.yValues))
    ? rawSeries.filter((item) => Array.isArray(item?.values) && Array.isArray(item?.yValues))
    : rawSeries.length >= 2
      ? [{ name: `${rawSeries[0].name}/${rawSeries[1].name}`, values: rawSeries[0].values, yValues: rawSeries[1].values }]
      : rawSeries;
  const xValues = series.flatMap((item) => item.values).filter(Number.isFinite);
  const yValues = series.flatMap((item) => item.yValues || item.values).filter(Number.isFinite);
  const xMin = Math.min(...xValues, 0);
  const xMax = Math.max(...xValues, 1);
  const yMin = Math.min(...yValues, 0);
  const yMax = Math.max(...yValues, 1);
  const xFor = (value) => plot.left + ((value - xMin) / Math.max(1, xMax - xMin)) * (plot.right - plot.left);
  const yFor = (value) => plot.bottom - ((value - yMin) / Math.max(1, yMax - yMin)) * (plot.bottom - plot.top);
  const dots = series
    .map((item, seriesIndex) =>
      item.values
        .map((value, index) => chartMarkerSvg(item, seriesIndex, index, xFor(value || 0), yFor((item.yValues || item.values)[index] || 0), chartSeriesColor(item, seriesIndex), { defaultSize: 8 }))
        .join(""),
    )
    .join("");
  return `${chartAxesSvg(plot, chart, { min: yMin, max: yMax })}${dots}`;
}



function pieChartSvg(chart, plot) {
  const values = chart.series[0]?.values || [];
  const total = values.reduce((sum, value) => sum + Math.max(0, Number(value) || 0), 0);
  const cx = (plot.left + plot.right) / 2;
  const cy = (plot.top + plot.bottom) / 2;
  const radius = Math.min(plot.right - plot.left, plot.bottom - plot.top) / 2 - 8;
  let angle = -Math.PI / 2;
  const series = chart.series[0] || {};
  const labels = [];
  const slices = values.map((value, index) => {
    const portion = total ? Math.max(0, Number(value) || 0) / total : 0;
    const nextAngle = angle + portion * Math.PI * 2;
    const path = pieSlicePath(cx, cy, radius, angle, nextAngle);
    const label = chartDataLabelText(chart, series, value, index, { total, percent: portion });
    if (label) {
      const point = pieLabelPoint(cx, cy, radius, angle, nextAngle, chartDataLabelModel(chart, series)?.position);
      labels.push(`<text data-chart-data-label="value" data-chart-data-label-position="${escapeAttr(chartDataLabelModel(chart, series)?.position || "bestFit")}" x="${roundSvg(point.x)}" y="${roundSvg(point.y)}" text-anchor="${escapeAttr(point.anchor)}" ${chartTextSvgAttrs(chartDataLabelModel(chart, series)?.textStyle, { fontSize: 9, fill: "#1f2937" })}>${escapeXml(label)}</text>`);
    }
    angle = nextAngle;
    return `<path d="${path}" fill="${escapeXml(chartPointColor(series, index, index))}" stroke="#ffffff" stroke-width="1"/>`;
  });
  return `${slices.join("")}${labels.join("")}`;
}



function pieSlicePath(cx, cy, radius, startAngle, endAngle) {
  const x1 = cx + Math.cos(startAngle) * radius;
  const y1 = cy + Math.sin(startAngle) * radius;
  const x2 = cx + Math.cos(endAngle) * radius;
  const y2 = cy + Math.sin(endAngle) * radius;
  const large = endAngle - startAngle > Math.PI ? 1 : 0;
  return `M ${roundSvg(cx)} ${roundSvg(cy)} L ${roundSvg(x1)} ${roundSvg(y1)} A ${roundSvg(radius)} ${roundSvg(radius)} 0 ${large} 1 ${roundSvg(x2)} ${roundSvg(y2)} Z`;
}



function pieLabelPoint(cx, cy, radius, startAngle, endAngle, position = "") {
  const angle = (startAngle + endAngle) / 2;
  const normalized = String(position || "").toLowerCase();
  const distance = normalized === "ctr"
    ? radius * 0.48
    : normalized === "inend" || normalized === "insideend"
      ? radius * 0.72
      : radius * 1.08;
  const x = cx + Math.cos(angle) * distance;
  const y = cy + Math.sin(angle) * distance + 3;
  return {
    x,
    y,
    anchor: normalized === "ctr" || normalized === "inend" || normalized === "insideend"
      ? "middle"
      : Math.cos(angle) < -0.2 ? "end" : Math.cos(angle) > 0.2 ? "start" : "middle",
  };
}



function chartAxesSvg(plot, chart = null, scale = {}, orientation = "column") {
  const categoryLine = chartLineSvgStyle(chart?.axes?.category?.line, { color: "#6b7280", width: 1 });
  const valueLine = chartLineSvgStyle(chart?.axes?.value?.line, { color: "#6b7280", width: 1 });
  const gridlines = chartGridlinesSvg(plot, chart, scale, orientation);
  if (orientation === "bar") {
    return `${gridlines}<line data-chart-axis="value" x1="${roundSvg(plot.left)}" y1="${roundSvg(plot.bottom)}" x2="${roundSvg(plot.right)}" y2="${roundSvg(plot.bottom)}" ${valueLine}/><line data-chart-axis="category" x1="${roundSvg(plot.left)}" y1="${roundSvg(plot.top)}" x2="${roundSvg(plot.left)}" y2="${roundSvg(plot.bottom)}" ${categoryLine}/>`;
  }
  return `${gridlines}<line data-chart-axis="category" x1="${roundSvg(plot.left)}" y1="${roundSvg(plot.bottom)}" x2="${roundSvg(plot.right)}" y2="${roundSvg(plot.bottom)}" ${categoryLine}/><line data-chart-axis="value" x1="${roundSvg(plot.left)}" y1="${roundSvg(plot.top)}" x2="${roundSvg(plot.left)}" y2="${roundSvg(plot.bottom)}" ${valueLine}/>`;
}



function chartGridlinesSvg(plot, chart = null, scale = {}, orientation = "column") {
  const gridline = chart?.axes?.value?.majorGridline;
  if (!gridline) return "";
  const style = chartLineSvgStyle(gridline, { color: "#d9e0ea", width: 0.75 });
  const segments = [];
  const steps = 4;
  for (let index = 1; index <= steps; index += 1) {
    if (orientation === "bar") {
      const x = plot.left + ((plot.right - plot.left) * index) / steps;
      segments.push(`<line data-chart-gridlines="value" data-chart-gridline-index="${index}" x1="${roundSvg(x)}" y1="${roundSvg(plot.top)}" x2="${roundSvg(x)}" y2="${roundSvg(plot.bottom)}" ${style}/>`);
    } else {
      const y = plot.bottom - ((plot.bottom - plot.top) * index) / steps;
      segments.push(`<line data-chart-gridlines="value" data-chart-gridline-index="${index}" x1="${roundSvg(plot.left)}" y1="${roundSvg(y)}" x2="${roundSvg(plot.right)}" y2="${roundSvg(y)}" ${style}/>`);
    }
  }
  return segments.join("");
}



function chartAxisTitlesSvg(chart, plot, width, height, orientation = "column") {
  const categoryTitle = chart?.axes?.category?.title || "";
  const valueTitle = chart?.axes?.value?.title || "";
  const labels = [];
  if (orientation === "bar") {
    if (categoryTitle) {
      const x = Math.max(14, plot.left - 48);
      const y = (plot.top + plot.bottom) / 2;
      labels.push(`<text data-chart-axis-title="category" x="${roundSvg(x)}" y="${roundSvg(y)}" transform="rotate(-90 ${roundSvg(x)} ${roundSvg(y)})" text-anchor="middle" ${chartTextSvgAttrs(chart?.axes?.category?.titleStyle, { fontSize: 11, fill: "#1f2937" })}>${escapeXml(categoryTitle)}</text>`);
    }
    if (valueTitle) {
      labels.push(`<text data-chart-axis-title="value" x="${roundSvg((plot.left + plot.right) / 2)}" y="${roundSvg(Math.min(height - 12, plot.bottom + 42))}" text-anchor="middle" ${chartTextSvgAttrs(chart?.axes?.value?.titleStyle, { fontSize: 11, fill: "#1f2937" })}>${escapeXml(valueTitle)}</text>`);
    }
    return labels.join("");
  }
  if (categoryTitle) {
    labels.push(`<text data-chart-axis-title="category" x="${roundSvg((plot.left + plot.right) / 2)}" y="${roundSvg(Math.min(height - 12, plot.bottom + 42))}" text-anchor="middle" ${chartTextSvgAttrs(chart?.axes?.category?.titleStyle, { fontSize: 11, fill: "#1f2937" })}>${escapeXml(categoryTitle)}</text>`);
  }
  if (valueTitle) {
    const x = Math.max(14, plot.left - 48);
    const y = (plot.top + plot.bottom) / 2;
    labels.push(`<text data-chart-axis-title="value" x="${roundSvg(x)}" y="${roundSvg(y)}" transform="rotate(-90 ${roundSvg(x)} ${roundSvg(y)})" text-anchor="middle" ${chartTextSvgAttrs(chart?.axes?.value?.titleStyle, { fontSize: 11, fill: "#1f2937" })}>${escapeXml(valueTitle)}</text>`);
  }
  return labels.join("");
}



function chartTextSvgAttrs(style = {}, fallback = {}) {
  const fontFamily = style?.fontFamily || fallback.fontFamily || "Yu Gothic UI, Meiryo, sans-serif";
  const fontSize = Number.isFinite(Number(style?.fontSize)) ? style.fontSize : fallback.fontSize || 10;
  const fill = cssColorForSvg(style?.color) || fallback.fill || "#1f2937";
  const fontWeight = style?.bold ? "700" : fallback.fontWeight || "";
  const fontStyle = style?.italic ? "italic" : fallback.fontStyle || "";
  const textDecoration = style?.underline ? "underline" : fallback.textDecoration || "";
  return [
    `font-family="${escapeAttr(fontFamily)}"`,
    `font-size="${roundSvg(fontSize)}"`,
    `fill="${escapeAttr(fill)}"`,
    fontWeight ? `font-weight="${escapeAttr(fontWeight)}"` : "",
    fontStyle ? `font-style="${escapeAttr(fontStyle)}"` : "",
    textDecoration ? `text-decoration="${escapeAttr(textDecoration)}"` : "",
  ].filter(Boolean).join(" ");
}



function chartLineSvgStyle(style = {}, fallback = {}) {
  const color = style?.color === "none" ? "none" : cssColorForSvg(style?.color) || fallback.color || "#6b7280";
  const width = color === "none" ? 0 : Number.isFinite(Number(style?.width)) ? normalizeShapeOutlineWidth(style.width) : Number(fallback.width) || 1;
  const dash = color === "none" ? "" : shapeLineDashArray(style?.dash, width);
  return `stroke="${escapeXml(color)}" stroke-width="${roundSvg(width)}"${dash ? ` stroke-dasharray="${escapeXml(dash)}"` : ""}`;
}



function chartDataLabelModel(chart, series) {
  return series?.dataLabels || chart?.dataLabels || null;
}



function chartDataLabelText(chart, series, value, index, context = {}) {
  const model = chartDataLabelModel(chart, series);
  if (!model) return "";
  const parts = [];
  if (model.showSeriesName && series?.name) parts.push(String(series.name));
  if (model.showCategoryName && chart?.labels?.[index] != null) parts.push(String(chart.labels[index]));
  if (model.showValue) parts.push(formatChartDataLabelValue(value));
  if (model.showPercent) parts.push(formatChartDataLabelPercent(context.percent, value, context.total));
  const separator = model.separator || (model.showValue && (model.showCategoryName || model.showSeriesName || model.showPercent) ? ", " : ", ");
  return parts.filter(Boolean).join(separator);
}



function formatChartDataLabelValue(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return String(value ?? "");
  if (Math.abs(number) >= 1000) return number.toLocaleString("ja-JP", { maximumFractionDigits: 1 });
  if (Number.isInteger(number)) return String(number);
  return String(Math.round(number * 100) / 100);
}



function formatChartDataLabelPercent(percent, value, total) {
  let ratio = Number(percent);
  if (!Number.isFinite(ratio)) {
    const number = Number(value);
    const totalNumber = Number(total);
    ratio = Number.isFinite(number) && Number.isFinite(totalNumber) && totalNumber ? number / totalNumber : 0;
  }
  return `${Math.round(ratio * 100)}%`;
}



function chartLabelsSvg(labels, plot) {
  return chartCategoryLabelsSvg(labels, plot, "column");
}



function chartCategoryLabelsSvg(labels, plot, orientation = "column") {
  const step = Math.max(1, Math.ceil(labels.length / 6));
  return labels
    .map((label, index) => {
      if (index % step !== 0) return "";
      if (orientation === "bar") {
        const y = plot.top + ((index + 0.5) / Math.max(1, labels.length)) * (plot.bottom - plot.top);
        return `<text x="${roundSvg(plot.left - 6)}" y="${roundSvg(y + 3)}" text-anchor="end" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="10" fill="#475467">${escapeXml(String(label).slice(0, 10))}</text>`;
      }
      const x = plot.left + ((index + 0.5) / Math.max(1, labels.length)) * (plot.right - plot.left);
      return `<text x="${roundSvg(x)}" y="${plot.bottom + 18}" text-anchor="middle" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="10" fill="#475467">${escapeXml(String(label).slice(0, 10))}</text>`;
    })
    .join("");
}



function chartLegendSvg(chart, width, height, plot, legendModel = { position: "bottom" }) {
  const items = chart.series.slice(0, 4);
  const position = legendModel.position || "bottom";
  return items
    .map((series, index) => {
      const point = chartLegendItemPoint(index, position, width, height, plot);
      return `<g data-chart-legend-position="${escapeAttr(position)}"><rect x="${roundSvg(point.x)}" y="${roundSvg(point.y - 9)}" width="10" height="10" fill="${escapeXml(chartSeriesColor(series, index))}"/><text x="${roundSvg(point.x + 16)}" y="${roundSvg(point.y)}" font-family="Yu Gothic UI, Meiryo, sans-serif" font-size="10" fill="#1f2937">${escapeXml(String(series.name).slice(0, 16))}</text></g>`;
    })
    .join("");
}



function chartLegendItemPoint(index, position, width, height, plot) {
  if (position === "right") return { x: width - 112, y: plot.top + 18 + index * 18 };
  if (position === "left") return { x: 18, y: plot.top + 18 + index * 18 };
  if (position === "top") return { x: plot.left + index * 118, y: 44 };
  return { x: 24 + index * 118, y: height - 22 };
}



function chartSeriesColor(series, seriesIndex = 0) {
  return cssColorForSvg(series?.color) || CHART_COLORS[seriesIndex % CHART_COLORS.length];
}



function chartPointColor(series, seriesIndex = 0, pointIndex = 0) {
  const color = series?.pointColors?.[pointIndex];
  return cssColorForSvg(color) || chartSeriesColor(series, seriesIndex);
}



function chartMarkerSvg(series, seriesIndex, pointIndex, x, y, fallbackColor, options = {}) {
  const marker = series?.marker || {};
  if (marker.symbol === "none") return "";
  const size = Math.max(3, Number(marker.size) || Number(options.defaultSize) || 6);
  const half = size / 2;
  const fill = cssColorForSvg(marker.fillColor) || chartPointColor(series, seriesIndex, pointIndex) || fallbackColor;
  const stroke = cssColorForSvg(marker.outlineColor) || fallbackColor || fill;
  const strokeWidth = Number.isFinite(Number(marker.outlineWidth)) ? Math.max(0.5, Number(marker.outlineWidth)) : 1;
  const cx = roundSvg(x);
  const cy = roundSvg(y);
  const common = `data-chart-marker="${escapeAttr(marker.symbol || "auto")}" fill="${escapeXml(fill)}" stroke="${escapeXml(stroke)}" stroke-width="${roundSvg(strokeWidth)}"`;
  switch (marker.symbol) {
    case "square":
      return `<rect ${common} x="${roundSvg(x - half)}" y="${roundSvg(y - half)}" width="${roundSvg(size)}" height="${roundSvg(size)}"/>`;
    case "diamond":
      return `<path ${common} d="M ${cx} ${roundSvg(y - half)} L ${roundSvg(x + half)} ${cy} L ${cx} ${roundSvg(y + half)} L ${roundSvg(x - half)} ${cy} Z"/>`;
    case "triangle":
      return `<path ${common} d="M ${cx} ${roundSvg(y - half)} L ${roundSvg(x + half)} ${roundSvg(y + half)} L ${roundSvg(x - half)} ${roundSvg(y + half)} Z"/>`;
    case "x":
      return `<g data-chart-marker="x" stroke="${escapeXml(stroke)}" stroke-width="${roundSvg(Math.max(1.2, strokeWidth))}" stroke-linecap="round"><line x1="${roundSvg(x - half)}" y1="${roundSvg(y - half)}" x2="${roundSvg(x + half)}" y2="${roundSvg(y + half)}"/><line x1="${roundSvg(x + half)}" y1="${roundSvg(y - half)}" x2="${roundSvg(x - half)}" y2="${roundSvg(y + half)}"/></g>`;
    case "plus":
      return `<g data-chart-marker="plus" stroke="${escapeXml(stroke)}" stroke-width="${roundSvg(Math.max(1.2, strokeWidth))}" stroke-linecap="round"><line x1="${roundSvg(x - half)}" y1="${cy}" x2="${roundSvg(x + half)}" y2="${cy}"/><line x1="${cx}" y1="${roundSvg(y - half)}" x2="${cx}" y2="${roundSvg(y + half)}"/></g>`;
    case "dash":
      return `<line data-chart-marker="dash" x1="${roundSvg(x - half)}" y1="${cy}" x2="${roundSvg(x + half)}" y2="${cy}" stroke="${escapeXml(stroke)}" stroke-width="${roundSvg(Math.max(1.5, strokeWidth))}" stroke-linecap="round"/>`;
    case "star":
      return chartStarMarkerSvg(x, y, half, fill, stroke, strokeWidth);
    case "circle":
    case "dot":
    case "auto":
    default:
      return `<circle ${common} cx="${cx}" cy="${cy}" r="${roundSvg(half)}"/>`;
  }
}



function chartStarMarkerSvg(x, y, radius, fill, stroke, strokeWidth) {
  const points = [];
  for (let index = 0; index < 10; index += 1) {
    const angle = -Math.PI / 2 + (index * Math.PI) / 5;
    const r = index % 2 === 0 ? radius : radius * 0.45;
    points.push(`${roundSvg(x + Math.cos(angle) * r)},${roundSvg(y + Math.sin(angle) * r)}`);
  }
  return `<polygon data-chart-marker="star" points="${points.join(" ")}" fill="${escapeXml(fill)}" stroke="${escapeXml(stroke)}" stroke-width="${roundSvg(strokeWidth)}"/>`;
}



function chartIsStacked(chart) {
  return chart?.grouping === "stacked" || chart?.grouping === "percentStacked";
}



function chartStackTotals(chart) {
  return chart.labels.map((_label, index) =>
    chart.series.reduce((sum, series) => sum + Math.max(0, Number(series.values[index]) || 0), 0),
  );
}



function chartMaxStackValue(chart) {
  return Math.max(...chartStackTotals(chart), 1);
}



function chartMaxValue(chart) {
  const values = chart.series.flatMap((series) => series.values).filter(Number.isFinite);
  return Math.max(...values, 1);
}



function chartMinValue(chart) {
  const values = chart.series.flatMap((series) => series.values).filter(Number.isFinite);
  return Math.min(...values, 0);
}



function chartTypeLabel(type) {
  switch (type) {
    case "line":
      return "折れ線グラフ";
    case "pie":
      return "円グラフ";
    case "scatter":
      return "散布図";
    case "map":
      return "マップ グラフ";
    case "bar":
      return "横棒グラフ";
    default:
      return "縦棒グラフ";
  }
}



function roundSvg(value) {
  return Math.round(Number(value || 0) * 100) / 100;
}



function escapeXml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}



function insertSparklinesFromSelection(type) {
  const range = singleInsertDataRange("スパークライン");
  if (!range) return;
  if (insertSparklinesToSheetGroup(type, range)) return;
  const sheet = activeSheet();
  const sheetId = getSheetId(sheet.name);
  const targetCol = range.right + 1;
  ensureSheetSize(sheet, range.bottom, targetCol);
  const targetRange = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: range.top, col: targetCol },
    { sheetIndex: state.activeSheetIndex, row: range.bottom, col: targetCol },
  );
  const before = captureCellsForHistory(sheet, targetRange);
  for (let row = range.top; row <= range.bottom; row += 1) {
    const values = [];
    for (let col = range.left; col <= range.right; col += 1) {
      values.push(cellNumericValue(sheet, sheetId, row, col));
    }
    const cell = ensureCellModel(sheet, row, targetCol);
    cell.sparkline = {
      type,
      values,
      sourceRange: `${columnName(range.left)}${row}:${columnName(range.right)}${row}`,
    };
    cell.raw = null;
    cell.display = "";
  }
  internCellsInRange(sheet, targetRange);
  const after = captureCellsForHistory(sheet, targetRange);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: targetRange,
    selectionBeforeRanges: activeSelectionRanges().map((item) => ({ ...item })),
    selectionAfterRanges: [targetRange],
    label: `${rangeToLabel(range)} スパークライン`,
  });
  state.selected = { sheetIndex: state.activeSheetIndex, row: targetRange.top, col: targetRange.left };
  state.selectionRange = targetRange;
  state.selectionRanges = [targetRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(`${rangeToLabel(targetRange)} にスパークラインを作成しました。`);
}



function insertSparklinesToSheetGroup(type, range) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    const sheetId = getSheetId(sheet?.name);
    if (!sheet || sheetId == null) continue;
    const sourceRange = materializeSelectionRangeForSheet(sheet, range, sheetIndex);
    if (!sourceRange || sourceRange.fullRow || sourceRange.fullColumn) continue;
    const targetCol = sourceRange.right + 1;
    const targetRange = normalizeRange(
      { sheetIndex, row: sourceRange.top, col: targetCol },
      { sheetIndex, row: sourceRange.bottom, col: targetCol },
    );
    if (!ensureEditableRangesForProtection(sheet, [targetRange], "スパークラインを作成")) return true;
    items.push({ sheetIndex, sheet, sheetId, sourceRange, targetRange, ranges: [targetRange] });
  }
  if (!items.length) return true;

  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  const selectionBeforeRanges = activeSelectionRanges().map((item) => ({ ...item }));
  const activeItem = items.find((item) => item.sheetIndex === state.activeSheetIndex) || items[0];
  try {
    items.forEach((item) => {
      const { sheet, sheetId, sourceRange, targetRange } = item;
      const targetCol = targetRange.left;
      ensureSheetSize(sheet, sourceRange.bottom, targetCol);
      for (let row = sourceRange.top; row <= sourceRange.bottom; row += 1) {
        const values = [];
        for (let col = sourceRange.left; col <= sourceRange.right; col += 1) {
          values.push(cellNumericValue(sheet, sheetId, row, col));
        }
        const cell = ensureCellModel(sheet, row, targetCol);
        cell.sparkline = {
          type,
          values,
          sourceRange: `${columnName(sourceRange.left)}${row}:${columnName(sourceRange.right)}${row}`,
        };
        cell.raw = null;
        cell.display = "";
      }
      internCellsInRange(sheet, targetRange);
    });

    const targetRange = activeItem.targetRange;
    state.selected = { sheetIndex: state.activeSheetIndex, row: targetRange.top, col: targetRange.left };
    state.selectionRange = { ...targetRange, sheetIndex: state.activeSheetIndex };
    state.selectionRanges = [{ ...targetRange, sheetIndex: state.activeSheetIndex }];
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
    state.selectedImage = null;
    state.selectedImages = [];
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: range,
      selectionAfter: targetRange,
      selectionBeforeRanges,
      selectionAfterRanges: [{ ...targetRange, sheetIndex: state.activeSheetIndex }],
      label: `${rangeToLabel(range)} ${sheetIndexes.length} シートスパークライン`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${rangeToLabel(targetRange)} にスパークラインを作成しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`スパークラインを作成できませんでした：${error.message || error}`);
  }
  return true;
}



function choosePictureForInsertion() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/png,image/jpeg,image/gif,image/webp,image/bmp,image/svg+xml,image/x-icon";
  input.multiple = true;
  input.className = "sr-only";
  input.addEventListener("change", () => {
    const files = Array.from(input.files || []);
    input.remove();
    if (files.length) void insertPictureFiles(files);
  }, { once: true });
  document.body.append(input);
  input.click();
}



async function insertPictureFiles(files) {
  const descriptors = [];
  for (const file of files) {
    const src = await fileToDataUrl(file);
    const size = await imageDimensionsFromSource(src);
    descriptors.push({
      name: file.name || "画像",
      src,
      mimeType: file.type || imageMimeType(fileExtension(file.name)),
      extension: fileExtension(file.name),
      width: size.width,
      height: size.height,
    });
  }
  insertWorksheetImageDescriptors(descriptors, { label: "画像" });
}



function insertShapeObject(shapeType = state.currentShapeType || DEFAULT_SHAPE_TYPE) {
  if (!ensureSheetProtectionPermission(activeSheet(), "allowEditObjects", "保護されたシートではオブジェクトを挿入できません。")) return;
  const descriptor = shapeDescriptorForType(shapeType, { useThemeEffects: true });
  setCurrentRibbonShape(descriptor.shape?.type || DEFAULT_SHAPE_TYPE);
  insertWorksheetImageDescriptors([descriptor], { label: "図形" });
}



function beginShapeDrawingMode(shapeType = DEFAULT_SHAPE_TYPE) {
  const preset = shapePresetForType(shapeType);
  if (!preset) return;
  if (!ensureSheetProtectionPermission(activeSheet(), "allowEditObjects", "保護されたシートではオブジェクトを挿入できません。")) return;
  cancelInkDrawingMode();
  cancelShapeDrawing({ keepPending: false, silent: true });
  setCurrentRibbonShape(preset.type);
  state.pendingShapeType = preset.type;
  document.body.classList.add("is-drawing-shape");
  $sheetHost.addClass("is-drawing-shape");
  setStatus(`${preset.label}：シート上でドラッグして位置とサイズを指定してください。`);
}



function cancelShapeDrawing(options = {}) {
  const drawing = state.shapeDrawing;
  if (drawing?.preview) drawing.preview.remove();
  state.shapeDrawing = null;
  removeLineConnectionPreviewOverlay();
  if (!options.keepPending) {
    state.pendingShapeType = null;
    document.body.classList.remove("is-drawing-shape");
    $sheetHost.removeClass("is-drawing-shape");
  }
  if (!options.silent) setStatus("図形の描画をキャンセルしました。");
}



function startShapeDrawing(event) {
  const shapeType = state.pendingShapeType;
  if (!shapeType || event.button !== 0) return false;
  const start = sheetPixelFromClientPosition(event.clientX, event.clientY);
  if (!start) return false;
  const grid = $sheetHost.find(".sheet-grid")[0];
  if (!grid) return false;

  closeRibbonShapeMenu();
  stopSelectionAutoScroll();
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  state.selectedImage = null;
  state.isSelecting = false;
  state.selectionDrag = null;
  state.shapeDrawing = {
    type: shapeType,
    start,
    current: start,
    startConnection: isLineShapeType(shapeType) ? snapLineEndpointToShape(start, "") : null,
    currentConnection: null,
    pointerId: event.pointerId,
    preview: createShapeDrawingPreview(shapeType),
  };
  grid.append(state.shapeDrawing.preview);
  updateShapeDrawingPreview();
  try {
    event.currentTarget.setPointerCapture?.(event.pointerId);
  } catch {}
  event.preventDefault();
  event.stopPropagation();
  return true;
}



function updateShapeDrawing(event) {
  const drawing = state.shapeDrawing;
  if (!drawing) return;
  const current = sheetPixelFromClientPosition(event.clientX, event.clientY, { clampToGrid: true });
  if (!current) return;
  drawing.current = current;
  updateShapeDrawingPreview();
  event.preventDefault?.();
}



function updatePendingShapeDrawingConnectionPreview(event) {
  if (state.shapeDrawing || !isLineShapeType(state.pendingShapeType)) return;
  if (state.imageTransform?.mode === "line-point") return;
  const point = sheetPixelFromClientPosition(event.clientX, event.clientY, { clampToGrid: false });
  updateLineDrawingConnectionPreview(point);
}



function clearPendingShapeDrawingConnectionPreview() {
  if (state.shapeDrawing || !isLineShapeType(state.pendingShapeType)) return;
  removeLineConnectionPreviewOverlay();
}



function finishShapeDrawing(event) {
  const drawing = state.shapeDrawing;
  if (!drawing) return;
  if (drawing.pointerId && event.pointerId && drawing.pointerId !== event.pointerId) return;
  const current = sheetPixelFromClientPosition(event.clientX, event.clientY, { clampToGrid: true }) || drawing.current || drawing.start;
  const type = drawing.type;
  const box = isLineShapeType(type)
    ? lineDrawingBoxWithConnections(type, drawing.start, current)
    : normalizedShapeDrawingBox(drawing.start, current);
  drawing.preview?.remove();
  state.shapeDrawing = null;
  state.pendingShapeType = null;
  document.body.classList.remove("is-drawing-shape");
  $sheetHost.removeClass("is-drawing-shape");
  removeLineConnectionPreviewOverlay();
  insertShapeAtPixelBox(type, box, { editText: type === "textBox" });
  event.preventDefault?.();
}



function lineDrawingBoxWithConnections(type, start, end) {
  const startSnap = snapLineEndpointToShape(start, "");
  const endSnap = snapLineEndpointToShape(end, "");
  const result = lineBoxAndPointsFromEndpoints(startSnap?.point || start, endSnap?.point || end, { type });
  const connections = {};
  if (startSnap?.imageId && startSnap.site) connections.start = { imageId: startSnap.imageId, site: startSnap.site };
  if (endSnap?.imageId && endSnap.site) connections.end = { imageId: endSnap.imageId, site: endSnap.site };
  const normalizedConnections = normalizeLineShapeConnections(connections);
  return normalizedConnections
    ? { ...result.boxWithPoints, connections: normalizedConnections }
    : result.boxWithPoints;
}



function updateLineDrawingConnectionPreview(point) {
  if (!point) {
    removeLineConnectionPreviewOverlay();
    return null;
  }
  const snap = snapLineEndpointToShape(point, "");
  const previewTargetId = snap?.imageId || lineConnectionPreviewTarget(point, "");
  updateLineConnectionPreviewOverlay(snap, "", previewTargetId);
  return snap;
}



function createShapeDrawingPreview(shapeType) {
  const preview = document.createElement("div");
  preview.className = `sheet-shape-draft${isLineShapeType(shapeType) ? " is-line-shape" : ""}`;
  preview.setAttribute("aria-hidden", "true");
  const image = document.createElement("img");
  image.alt = "";
  image.draggable = false;
  image.src = shapeDescriptorForType(shapeType, { useThemeEffects: true }).src;
  preview.append(image);
  return preview;
}



function updateShapeDrawingPreview() {
  const drawing = state.shapeDrawing;
  if (!drawing?.preview) return;
  const box = normalizedShapeDrawingBox(drawing.start, drawing.current);
  if (isLineShapeType(drawing.type)) {
    const currentSnap = updateLineDrawingConnectionPreview(drawing.current);
    drawing.currentConnection = currentSnap;
    const startPoint = drawing.startConnection?.point || drawing.start;
    const currentPoint = currentSnap?.point || drawing.current;
    const lineGeometry = lineBoxAndPointsFromEndpoints(startPoint, currentPoint, { type: drawing.type });
    const image = drawing.preview.querySelector("img");
    if (image) image.src = shapeDescriptorForType(drawing.type, { points: lineGeometry.points, width: lineGeometry.box.width, height: lineGeometry.box.height, useThemeEffects: true }).src;
    Object.assign(drawing.preview.style, {
      left: `${lineGeometry.box.left}px`,
      top: `${lineGeometry.box.top}px`,
      width: `${lineGeometry.box.width}px`,
      height: `${lineGeometry.box.height}px`,
    });
    return;
  }
  Object.assign(drawing.preview.style, {
    left: `${box.left}px`,
    top: `${box.top}px`,
    width: `${box.width}px`,
    height: `${box.height}px`,
  });
}



function normalizedShapeDrawingBox(start, current) {
  const left = Math.min(start.x, current.x);
  const top = Math.min(start.y, current.y);
  const width = Math.abs(current.x - start.x);
  const height = Math.abs(current.y - start.y);
  return {
    left,
    top,
    width: Math.max(12, width || 120),
    height: Math.max(12, height || 72),
  };
}



function lineBoxAndPointsFromEndpoints(start, end, options = {}) {
  const padding = lineEndpointPaddingForShape(options.shape || options.type || options);
  const rawLeft = Math.min(start.x, end.x);
  const rawTop = Math.min(start.y, end.y);
  const rawWidth = Math.abs(end.x - start.x);
  const rawHeight = Math.abs(end.y - start.y);
  const contentWidth = rawWidth + padding * 2;
  const contentHeight = rawHeight + padding * 2;
  const width = Math.max(LINE_SHAPE_MIN_BOX_SIZE, contentWidth);
  const height = Math.max(LINE_SHAPE_MIN_BOX_SIZE, contentHeight);
  const left = rawLeft - padding - Math.max(0, width - contentWidth) / 2;
  const top = rawTop - padding - Math.max(0, height - contentHeight) / 2;
  const points = {
    x1: (start.x - left) / width,
    y1: (start.y - top) / height,
    x2: (end.x - left) / width,
    y2: (end.y - top) / height,
  };
  const box = {
    left,
    top,
    width,
    height,
  };
  return {
    box,
    points: normalizeShapeLinePoints(points),
    boxWithPoints: { ...box, points: normalizeShapeLinePoints(points) },
  };
}



function lineEndpointPaddingForShape(input = {}) {
  const shape = typeof input === "string" ? { type: input } : input || {};
  const type = shape.type || "";
  const strokeWidth = normalizeShapeOutlineWidth(shape.outlineWidth);
  const lineStartArrow = Boolean(shape.lineStartArrow ?? type === "doubleArrow");
  const lineEndArrow = Boolean(shape.lineEndArrow ?? ["arrow", "doubleArrow", "elbowConnector", "curvedConnector"].includes(type));
  const startType = normalizeShapeArrowType(shape.lineStartArrowType || "triangle");
  const endType = normalizeShapeArrowType(shape.lineEndArrowType || "triangle");
  const startPadding = lineStartArrow && startType && startType !== "open"
    ? shapeArrowHalfWidth(strokeWidth, shape.lineStartArrowSize)
    : 0;
  const endPadding = lineEndArrow && endType && endType !== "open"
    ? shapeArrowHalfWidth(strokeWidth, shape.lineEndArrowSize)
    : 0;
  return Math.ceil(Math.max(LINE_SHAPE_BASE_PADDING_PX, startPadding, endPadding) + LINE_SHAPE_ARROW_PADDING_EXTRA_PX);
}



function insertShapeAtPixelBox(shapeType, box, options = {}) {
  const descriptor = options.descriptor || shapeDescriptorForType(shapeType, { useThemeEffects: true });
  const sheet = activeSheet();
  if (!options.skipSheetGroup && insertShapeAtPixelBoxToSheetGroup(shapeType, box, { ...options, descriptor })) return null;
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを挿入できません。")) return null;
  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  const image = addShapeAtPixelBoxToSheet(sheet, shapeType, box, descriptor);
  if (!image) return null;

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
    label: `${descriptor.name || "図形"}の描画`,
  });
  renderSheet();
  updateSelectionUi();
  activateObjectFormatRibbonForSelection();
  if (!options.silentStatus) setStatus(`${descriptor.name || "図形"}を挿入しました。`);
  if (options.editText) openShapeTextEditor(image.id);
  return image;
}



function insertShapeAtPixelBoxToSheetGroup(shapeType, box, options = {}) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const descriptor = options.descriptor || shapeDescriptorForType(shapeType, { useThemeEffects: true });
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを挿入できません。")) return true;
    items.push({ sheetIndex, sheet });
  }
  if (!items.length) return true;

  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  try {
    let activeInserted = null;
    items.forEach((item) => {
      const image = addShapeAtPixelBoxToSheet(item.sheet, shapeType, box, descriptor);
      if (item.sheetIndex === state.activeSheetIndex) activeInserted = image;
    });
    if (!activeInserted) throw new Error("No shape was inserted on the active sheet.");
    state.selected = null;
    state.selectionRange = null;
    state.selectionRanges = [];
    state.selectionAnchor = null;
    state.selectionStart = null;
    state.selectionDrag = null;
    state.formulaEdit = null;
    setSelectedImageIds([activeInserted.id], activeInserted.id);
    $formulaInput.prop("disabled", true).val("");
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore,
      selectionAfter: null,
      selectionBeforeRanges,
      selectionAfterRanges: [],
      label: `${sheetIndexes.length} シート${descriptor.name || "図形"}の描画`,
    });
    renderSheet();
    updateSelectionUi();
    activateObjectFormatRibbonForSelection();
    if (!options.silentStatus) setStatus(`${sheetIndexes.length} 個のシートに${descriptor.name || "図形"}を挿入しました。`);
    if (options.editText) openShapeTextEditor(activeInserted.id);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`図形を挿入できませんでした：${error.message || error}`);
  }
  return true;
}



function addShapeAtPixelBoxToSheet(sheet, shapeType, box, descriptor) {
  if (!sheet || !box || !descriptor) return null;
  const left = Math.max(rowHeaderWidthForSheet(sheet), Math.round(box.left));
  const top = Math.max(COLUMN_HEADER_HEIGHT, Math.round(box.top));
  const width = Math.max(1, Math.round(box.width));
  const height = Math.max(1, Math.round(box.height));
  const shape = cloneJsonLikeValue(descriptor.shape);
  if (isLineShapeType(shape.type) && box.points) {
    shape.points = normalizeShapeLinePoints(box.points);
    const connections = normalizeLineShapeConnections(box.connections);
    if (connections) shape.connections = connections;
  }
  const sourceDescriptor = shapeDescriptorForType(shape.type || shapeType, { ...shape, width, height });
  const anchor = nativeAnchorFromSheetPixel(sheet, "col", left);
  const rowAnchor = nativeAnchorFromSheetPixel(sheet, "row", top);
  const endCol = nativeAnchorFromSheetPixel(sheet, "col", left + width);
  const endRow = nativeAnchorFromSheetPixel(sheet, "row", top + height);
  ensureSheetSize(sheet, Math.max(1, endRow.index + 1), Math.max(1, endCol.index + 1));
  const image = {
    id: uniqueWorksheetObjectId("image"),
    name: sourceDescriptor.name || "図形",
    src: sourceDescriptor.src,
    mimeType: sourceDescriptor.mimeType || "image/svg+xml",
    extension: sourceDescriptor.extension || "svg",
    editAs: "oneCell",
    tl: {
      col: anchor.index,
      row: rowAnchor.index,
      nativeCol: anchor.index,
      nativeRow: rowAnchor.index,
      nativeColOff: anchor.offsetEmu,
      nativeRowOff: rowAnchor.offsetEmu,
    },
    br: null,
    ext: { width, height },
    hyperlink: "",
    shape,
  };
  sheet.images = Array.isArray(sheet.images) ? sheet.images : [];
  sheet.images.push(image);
  return image;
}



function insertTextBoxObject() {
  beginShapeDrawingMode("textBox");
}



function insertWordArtObject() {
  insertPresetTextObject("ワードアート", {
    type: "textBox",
    text: "ワードアート",
    width: 270,
    height: 74,
    fillColor: "none",
    outlineColor: "none",
    textColor: "#4472c4",
    textOutlineColor: "#1f4e79",
    textOutlineWidth: 0.9,
    textFontSize: 28,
    textFontFamily: "Yu Gothic UI, Meiryo, sans-serif",
    textBold: true,
    textItalic: true,
    textAlign: "center",
    textVAlign: "middle",
    textMargin: { left: 4, right: 4, top: 4, bottom: 4 },
  });
}



function openSignatureLineDialog() {
  closeDataDialog({ keepState: true });
  state.dataDialog = { type: "signature-line" };
  const overlay = document.createElement("div");
  overlay.id = "dataDialog";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = signatureLineDialogHtml();
  document.body.appendChild(overlay);
  enhanceDialogAccessKeys(overlay);
  overlay.querySelector("[data-signature-field='signer']")?.focus({ preventScroll: true });
}



function signatureLineDialogHtml() {
  return `
    <div class="excel-mini-dialog data-dialog signature-setup-dialog" role="dialog" aria-modal="true" aria-labelledby="signatureLineTitle" data-data-dialog="signature-line">
      <div class="excel-mini-titlebar" id="signatureLineTitle">署名の設定</div>
      <div class="excel-mini-body data-dialog-body">
        <label class="excel-dialog-field">
          <span class="excel-dialog-caption">${accessKeyLabelHtml("署名候補者(S):")}</span>
          <input type="text" data-signature-field="signer" value="署名者" />
        </label>
        <label class="excel-dialog-field">
          <span class="excel-dialog-caption">${accessKeyLabelHtml("署名候補者の役職(T):")}</span>
          <input type="text" data-signature-field="title" value="役職" />
        </label>
        <label class="excel-dialog-field">
          <span class="excel-dialog-caption">${accessKeyLabelHtml("署名候補者の電子メール アドレス(E):")}</span>
          <input type="text" data-signature-field="email" value="" />
        </label>
        <label class="excel-dialog-field">
          <span class="excel-dialog-caption">${accessKeyLabelHtml("署名者に表示する指示(I):")}</span>
          <textarea rows="3" data-signature-field="instructions"></textarea>
        </label>
        <label class="data-dialog-checkbox"><input type="checkbox" data-signature-check="allowComments" checked />${accessKeyLabelHtml("署名者に [署名] ダイアログでコメントの追加を許可する(A)")}</label>
        <label class="data-dialog-checkbox"><input type="checkbox" data-signature-check="showDate" checked />${accessKeyLabelHtml("署名欄に署名日を表示する(D)")}</label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="signature-line-ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
}



function applySignatureLineDialog(dialog) {
  const settings = {
    signer: String(dialog?.querySelector("[data-signature-field='signer']")?.value || "署名者").trim() || "署名者",
    title: String(dialog?.querySelector("[data-signature-field='title']")?.value || "").trim(),
    email: String(dialog?.querySelector("[data-signature-field='email']")?.value || "").trim(),
    instructions: String(dialog?.querySelector("[data-signature-field='instructions']")?.value || "").trim(),
    allowComments: Boolean(dialog?.querySelector("[data-signature-check='allowComments']")?.checked),
    showDate: Boolean(dialog?.querySelector("[data-signature-check='showDate']")?.checked),
  };
  closeDataDialog();
  insertSignatureLineObject(settings);
}



function insertSignatureLineObject(settings = {}) {
  const signer = String(settings.signer || "署名者").trim() || "署名者";
  const title = String(settings.title || "役職").trim();
  const email = String(settings.email || "").trim();
  const dateLine = settings.showDate === false ? "" : "\n署名日";
  const emailLine = email ? `\n${email}` : "";
  insertPresetTextObject("署名欄", {
    type: "textBox",
    text: `________________________\n${signer}${title ? `\n${title}` : ""}${emailLine}${dateLine}`,
    width: 330,
    height: 112,
    fillColor: "none",
    outlineColor: "#a6a6a6",
    outlineWidth: 0.9,
    textColor: "#1f2937",
    textFontSize: 12,
    textFontFamily: "Yu Gothic UI, Meiryo, sans-serif",
    textAlign: "left",
    textVAlign: "bottom",
    textMargin: { left: 18, right: 18, top: 12, bottom: 12 },
    objectKind: "signatureLine",
    objectMetadata: {
      signer,
      title,
      email,
      instructions: String(settings.instructions || ""),
      allowComments: settings.allowComments !== false,
      showDate: settings.showDate !== false,
    },
  });
}



function objectInsertTypeDefinitions() {
  return [
    { value: "package", label: "パッケージ", icon: "📦" },
    { value: "excel-sheet", label: "Cht WebSheet ワークシート", icon: "X" },
    { value: "word-document", label: "Microsoft Word 文書", icon: "W" },
    { value: "powerpoint-slide", label: "Microsoft PowerPoint スライド", icon: "P" },
    { value: "bitmap-image", label: "ビットマップ イメージ", icon: "B" },
    { value: "text-file", label: "テキスト ファイル", icon: "T" },
    { value: "pdf-document", label: "PDF 文書", icon: "PDF" },
  ];
}



function openInsertObjectDialog() {
  closeDataDialog({ keepState: true });
  state.dataDialog = { type: "insert-object", mode: "create" };
  const overlay = document.createElement("div");
  overlay.id = "dataDialog";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = insertObjectDialogHtml("create");
  document.body.appendChild(overlay);
  enhanceDialogAccessKeys(overlay);
  syncInsertObjectDialog(overlay.querySelector("[data-data-dialog='insert-object']"));
  overlay.querySelector("[data-object-field='classType']")?.focus({ preventScroll: true });
}



function insertObjectDialogHtml(mode = "create") {
  const activeMode = mode === "file" ? "file" : "create";
  const typeOptions = objectInsertTypeDefinitions()
    .map((item) => `<option value="${escapeAttr(item.value)}">${escapeHtml(item.label)}</option>`)
    .join("");
  return `
    <div class="excel-mini-dialog data-dialog object-insert-dialog" role="dialog" aria-modal="true" aria-labelledby="insertObjectTitle" data-data-dialog="insert-object">
      <div class="excel-mini-titlebar" id="insertObjectTitle">オブジェクトの挿入</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="object-insert-tabs" role="tablist" aria-label="オブジェクトの挿入">
          <button class="${activeMode === "create" ? "is-active" : ""}" type="button" role="tab" data-object-insert-tab="create">${accessKeyLabelHtml("新規作成(C)")}</button>
          <button class="${activeMode === "file" ? "is-active" : ""}" type="button" role="tab" data-object-insert-tab="file">${accessKeyLabelHtml("ファイルから(F)")}</button>
        </div>
        <section class="object-insert-panel${activeMode === "create" ? " is-active" : ""}" data-object-insert-panel="create">
          <label class="excel-dialog-field">
            <span class="excel-dialog-caption">${accessKeyLabelHtml("オブジェクトの種類(O):")}</span>
            <select size="7" data-object-field="classType" data-object-insert-sync>${typeOptions}</select>
          </label>
          <p class="object-insert-description">選択した種類の新しいオブジェクトをワークシートに挿入します。</p>
        </section>
        <section class="object-insert-panel${activeMode === "file" ? " is-active" : ""}" data-object-insert-panel="file">
          <label class="excel-dialog-field object-insert-file-row">
            <span class="excel-dialog-caption">${accessKeyLabelHtml("ファイル名(N):")}</span>
            <span class="object-insert-file-control">
              <input type="text" data-object-field="fileName" data-object-insert-sync value="" />
              <button class="excel-mini-button" type="button" data-data-dialog-action="object-insert-browse">${accessKeyLabelHtml("参照(B)...")}</button>
            </span>
          </label>
          <label class="data-dialog-checkbox"><input type="checkbox" data-object-field="linkToFile" data-object-insert-sync />${accessKeyLabelHtml("ファイルにリンク(L)")}</label>
        </section>
        <fieldset class="data-dialog-fieldset object-insert-options">
          <legend>オプション</legend>
          <label class="data-dialog-checkbox"><input type="checkbox" data-object-field="displayAsIcon" data-object-insert-sync />${accessKeyLabelHtml("アイコンで表示(A)")}</label>
          <label class="excel-dialog-field object-insert-icon-label">
            <span class="excel-dialog-caption">${accessKeyLabelHtml("アイコンの表題(I):")}</span>
            <input type="text" data-object-field="iconLabel" data-object-insert-sync value="" />
          </label>
        </fieldset>
        <div class="object-insert-preview" aria-live="polite">
          <span class="object-insert-preview-icon" data-object-preview-icon>📦</span>
          <span data-object-preview-label>パッケージ</span>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="object-insert-ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
}



function handleObjectInsertTabClick(event) {
  const mode = event.currentTarget.dataset.objectInsertTab === "file" ? "file" : "create";
  const dialog = event.currentTarget.closest("[data-data-dialog='insert-object']");
  if (state.dataDialog?.type === "insert-object") state.dataDialog.mode = mode;
  dialog?.querySelectorAll("[data-object-insert-tab]").forEach((button) => {
    const active = button.dataset.objectInsertTab === mode;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-selected", active ? "true" : "false");
  });
  dialog?.querySelectorAll("[data-object-insert-panel]").forEach((panel) => {
    panel.classList.toggle("is-active", panel.dataset.objectInsertPanel === mode);
  });
  syncInsertObjectDialog(dialog);
  const focusSelector = mode === "file" ? "[data-object-field='fileName']" : "[data-object-field='classType']";
  dialog?.querySelector(focusSelector)?.focus({ preventScroll: true });
}



function handleObjectInsertDialogSync(event) {
  syncInsertObjectDialog(event.currentTarget.closest("[data-data-dialog='insert-object']"));
}



function syncInsertObjectDialog(dialog) {
  if (!dialog) return;
  const mode = activeInsertObjectMode(dialog);
  const type = selectedInsertObjectType(dialog);
  const fileName = String(dialog.querySelector("[data-object-field='fileName']")?.value || "").trim();
  const displayAsIcon = Boolean(dialog.querySelector("[data-object-field='displayAsIcon']")?.checked);
  const iconLabelInput = dialog.querySelector("[data-object-field='iconLabel']");
  if (iconLabelInput) {
    iconLabelInput.disabled = !displayAsIcon;
    if (displayAsIcon && !iconLabelInput.value.trim()) iconLabelInput.value = mode === "file" ? fileName.split(/[\\/]/).pop() || "パッケージ" : type.label;
  }
  dialog.querySelectorAll("[data-object-insert-panel='file'] input, [data-object-insert-panel='file'] button").forEach((control) => {
    if (control.dataset?.objectInsertTab) return;
    control.disabled = mode !== "file";
  });
  dialog.querySelectorAll("[data-object-insert-panel='create'] select").forEach((control) => {
    control.disabled = mode !== "create";
  });
  const previewIcon = dialog.querySelector("[data-object-preview-icon]");
  const previewLabel = dialog.querySelector("[data-object-preview-label]");
  if (previewIcon) previewIcon.textContent = displayAsIcon ? type.icon : "□";
  if (previewLabel) previewLabel.textContent = displayAsIcon
    ? iconLabelInput?.value || type.label
    : mode === "file" ? fileName || "ファイル" : type.label;
}



function activeInsertObjectMode(dialog) {
  return dialog?.querySelector("[data-object-insert-tab].is-active")?.dataset.objectInsertTab === "file" ? "file" : "create";
}



function selectedInsertObjectType(dialog) {
  const definitions = objectInsertTypeDefinitions();
  const value = dialog?.querySelector("[data-object-field='classType']")?.value || definitions[0].value;
  return definitions.find((item) => item.value === value) || definitions[0];
}



function chooseEmbeddedObjectFile(dialog) {
  const input = document.createElement("input");
  input.type = "file";
  input.className = "sr-only";
  input.addEventListener("change", () => {
    const file = input.files?.[0];
    input.remove();
    if (!file) return;
    const fileNameInput = dialog?.querySelector("[data-object-field='fileName']");
    if (fileNameInput) fileNameInput.value = file.name || "";
    const iconLabelInput = dialog?.querySelector("[data-object-field='iconLabel']");
    if (dialog?.querySelector("[data-object-field='displayAsIcon']")?.checked && iconLabelInput) iconLabelInput.value = file.name || "";
    syncInsertObjectDialog(dialog);
  }, { once: true });
  document.body.append(input);
  input.click();
}



function applyInsertObjectDialog(dialog) {
  const mode = activeInsertObjectMode(dialog);
  const type = selectedInsertObjectType(dialog);
  const fileName = String(dialog?.querySelector("[data-object-field='fileName']")?.value || "").trim();
  const displayAsIcon = Boolean(dialog?.querySelector("[data-object-field='displayAsIcon']")?.checked);
  const iconLabel = String(dialog?.querySelector("[data-object-field='iconLabel']")?.value || "").trim();
  if (mode === "file" && !fileName) {
    showMessageBox("ファイル名を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  closeDataDialog();
  insertEmbeddedObjectPlaceholder({
    mode,
    objectClass: type.value,
    objectClassLabel: type.label,
    fileName,
    linkToFile: Boolean(dialog?.querySelector("[data-object-field='linkToFile']")?.checked),
    displayAsIcon,
    iconLabel,
    icon: type.icon,
  });
}

export {
  accessKeyFromKeyboardEvent,
  activateDialogAccessKeyTarget,
  activeInsertObjectMode,
  addShapeAtPixelBoxToSheet,
  applyInsertObjectDialog,
  applySignatureLineDialog,
  barChartSvg,
  barDataLabelPoint,
  beginShapeDrawingMode,
  cancelShapeDrawing,
  cellNumericValue,
  chartAreaSvgStyle,
  chartAxesSvg,
  chartAxisTitlesSvg,
  chartCategoryLabelsSvg,
  chartDataFromRange,
  chartDataLabelModel,
  chartDataLabelText,
  chartDescriptor,
  chartGridlinesSvg,
  chartIsStacked,
  chartLabelsSvg,
  chartLegendItemPoint,
  chartLegendSvg,
  chartLineSvgStyle,
  chartMarkerSvg,
  chartMaxStackValue,
  chartMaxValue,
  chartMinValue,
  chartOrientationForType,
  chartPlotAreaRectSvg,
  chartPlotBox,
  chartPointColor,
  chartSeriesColor,
  chartStackTotals,
  chartStarMarkerSvg,
  chartSvg,
  chartTextSvgAttrs,
  chartTypeLabel,
  chooseEmbeddedObjectFile,
  choosePictureForInsertion,
  clearPendingShapeDrawingConnectionPreview,
  closeDataDialog,
  closeDataDialogOnEscape,
  columnChartSvg,
  columnDataLabelPoint,
  createShapeDrawingPreview,
  enhanceButtonAccessKey,
  enhanceCaptionAccessKey,
  enhanceChoiceLabelAccessKey,
  enhanceDialogAccessKeys,
  escapeXml,
  findDialogAccessKeyTarget,
  finishShapeDrawing,
  formatChartDataLabelPercent,
  formatChartDataLabelValue,
  handleDataDialogAction,
  handleDataDialogKeydown,
  handleGlobalDataDialogAccessKey,
  handleObjectInsertDialogSync,
  handleObjectInsertTabClick,
  historyShortcutAction,
  insertChartFromSelection,
  insertChartSheetFromSelection,
  insertChartToSheetGroup,
  insertObjectDialogHtml,
  insertPictureFiles,
  insertPivotSummaryFromSelection,
  insertPivotSummaryToSheetGroup,
  insertShapeAtPixelBox,
  insertShapeAtPixelBoxToSheetGroup,
  insertShapeObject,
  insertSignatureLineObject,
  insertSparklinesFromSelection,
  insertSparklinesToSheetGroup,
  insertTextBoxObject,
  insertWordArtObject,
  installDialogAccessKeyEnhancer,
  lineBoxAndPointsFromEndpoints,
  lineChartSvg,
  lineDrawingBoxWithConnections,
  lineEndpointPaddingForShape,
  mapChartSvg,
  nextPivotTableName,
  normalizedChartLegend,
  normalizedShapeDrawingBox,
  objectInsertTypeDefinitions,
  openInsertObjectDialog,
  openSignatureLineDialog,
  outlineDetailTargetFromRange,
  outlineRangeForAxis,
  pieChartSvg,
  pieLabelPoint,
  pieSlicePath,
  pivotSummaryFromRange,
  rangesOverlap1d,
  recommendedChartTypeForSelection,
  roundSvg,
  scatterChartSvg,
  selectedInsertObjectType,
  shouldHandleDialogHistoryShortcut,
  signatureLineDialogHtml,
  singleInsertDataRange,
  startShapeDrawing,
  syncInsertObjectDialog,
  toggleOutlineDetailFromSelection,
  uniqueSheetNameWithBase,
  updateLineDrawingConnectionPreview,
  updatePendingShapeDrawingConnectionPreview,
  updateShapeDrawing,
  updateShapeDrawingPreview,
  writePivotSummaryAt,
  writePivotSummarySheet,
};
