import { imageBoxToPixels } from '../editing/selection-and-cell-text.js';
// Feature module; shared bindings are imported explicitly.
import {
  $,
  $addSheetButton,
  $bootState,
  $dropZone,
  $emptyState,
  $exportHtmlButton,
  $fileInput,
  $fileSideMenu,
  $formulaBar,
  $formulaExpandButton,
  $formulaInput,
  $formulaResizeHandle,
  $horizontalScrollbar,
  $pageBackgroundInput,
  $ribbonDisplayButton,
  $ribbonDisplayMenu,
  $ribbonPanels,
  $ribbonTabs,
  $selectedCellLabel,
  $selectionKeyboardSink,
  $selectionStatus,
  $sheetFooter,
  $sheetFooterSplitter,
  $sheetHost,
  $sheetTabScrollLeft,
  $sheetTabScrollRight,
  $sheetTabs,
  $sheetWorkspace,
  $statusText,
  $toggleFormulasButton,
  $workbookName,
  $workbookWindowBar,
  $zoomInButton,
  $zoomLabel,
  $zoomOutButton,
  $zoomSlider,
  APP_DISPLAY_NAME,
  APP_SHORT_NAME,
  APP_VERSION,
  CWS_AI_GUIDE_URL,
  CWS_HTML_FORMAT,
  CWS_MODEL_SCHEMA_URL,
  CWS_SCHEMA_VERSION,
  DEFAULT_THEME_COLORS,
  STANDALONE_RUNTIME_VERSION,
  WORKBOOK_OPEN_TRANSFER_DB_NAME,
  WORKBOOK_OPEN_TRANSFER_DB_VERSION,
  WORKBOOK_OPEN_TRANSFER_MAX_AGE_MS,
  WORKBOOK_OPEN_TRANSFER_QUERY_PARAM,
  WORKBOOK_OPEN_TRANSFER_STORE_NAME,
  createIcons,
  icons,
  initCwsAgentBridge,
  initWorkbookNameDisplay,
  state,
} from "../../app-core.js";
import {
  appendCwsAgentPreviewOpsProgressively,
  applyCwsAgentOps,
  captureCwsAgentTaskWorkbook,
  clearCwsAgentPreview,
  commitCwsAgentPreview,
  normalizeCwsAgentTaskSnapshot,
  persistCwsAgentTaskWorkbook,
  previewCwsAgentOps,
  previewCwsAgentOpsProgressively,
  restoreCwsAgentTaskWorkbook,
  setCwsAgentPreviewVisible,
} from "../agent/contracts.js";
import {
  closeWhatIfMenu,
  closeWhatIfMenuOnEscape,
  closeWhatIfMenuOnOutsidePointer,
  handleDataAnalysisDialogSync,
  handleDataAnalysisPickerDoubleClick,
  handleDataAnalysisPickerSync,
  handleDataTableDialogSync,
  handleGoalSeekDialogSync,
  handleScenarioEditDialogSync,
  handleScenarioListChange,
  handleScenarioSummaryDialogSync,
  handleWhatIfMenuAction,
} from "../data/analysis-and-scenarios.js";
import {
  cloneWorkbookDataModel,
  closeQueryPaneContextMenu,
  closeQueryPaneContextMenuOnOutsidePointer,
  closeQueryPaneOnEscape,
  handleAppendQuerySelectionChange,
  handleDataModelDialogSync,
  handleExistingConnectionDialogSync,
  handleImportDataDialogSync,
  handleMergeQueryPreviewChange,
  handleMergeRightQueryChange,
  handleQueryPaneAction,
  handleQueryPaneContextMenuAction,
  openQueryPaneContextMenuFromEvent,
  workbookQueries,
} from "../data/connections-and-queries.js";
import {
  closeAutoFilterMenuOnEscape,
  closeAutoFilterMenuOnOutsidePointer,
  handleAdvancedFilterDialogSync,
  handleAutoFilterDialogSync,
  handleAutoFilterMenuAction,
  handleAutoFilterMenuKeydown,
  handleAutoFilterSearchInput,
  handleAutoFilterSelectAllChange,
  handleAutoFilterValueChange,
  handleTextColumnsPreviewChange,
} from "../data/filter-and-text-tools.js";
import {
  applyOutlineToSheet,
  ensureOutlineAxisAllowed,
  handleForecastDialogSync,
  handleOutlineDialogSync,
  handleOutlineLevelCommandClick,
  handleOutlineToggleClick,
} from "../data/forecast-and-statistics.js";
import {
  closeFindDialog,
  closeFindDialogOnEscape,
  handleFindDialogFieldChange,
  handleFindDialogKeydown,
  handleTextCsvPreviewChange,
  selectRangesForHomeCommand,
  selectSingleRange,
} from "../data/query-editor.js";
import {
  closeDataCommandMenu,
  closeDataCommandMenuOnEscape,
  closeDataCommandMenuOnOutsidePointer,
  handleConnectionOnlySourceDialogSync,
  handleDataCommandMenuAction,
  handleImportAuthDialogSync,
  handleWorkbookLinksDialogSync,
} from "../data/query-transforms.js";
import {
  handleQueryPropertiesDialogSync,
  handleSortDialogControlChange,
  openAutoFilterMenu,
} from "../data/sort-and-autofilter.js";
import {
  closeDataValidationMenu,
  closeDataValidationMenuOnEscape,
  closeDataValidationMenuOnOutsidePointer,
  closeValidationDropdownOnEscape,
  closeValidationDropdownOnOutsidePointer,
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
} from "../data/validation-and-consolidation.js";
import {
  clearPendingShapeDrawingConnectionPreview,
  closeDataDialogOnEscape,
  finishShapeDrawing,
  handleDataDialogAction,
  handleDataDialogKeydown,
  handleGlobalDataDialogAccessKey,
  handleObjectInsertDialogSync,
  handleObjectInsertTabClick,
  installDialogAccessKeyEnhancer,
  updatePendingShapeDrawingConnectionPreview,
  updateShapeDrawing,
} from "../drawing/charts-and-insertion.js";
import {
  handleInsertIconSearch,
  headerFooterPlainText,
  normalizedHeaderFooterSettings,
} from "../drawing/shape-rendering.js";
import {
  applySelectedShapeRotation,
  applySelectedShapeSize,
  finishPictureCropDrag,
  finishSelectedImageTransform,
  startPictureCropDrag,
  startSelectedImageTransform,
  updatePictureCropDrag,
  updateSelectedImageTransform,
} from "../drawing/transforms-and-connectors.js";
import {
  cellEditRichTextForCommit,
  clearAutoFillOptions,
  clearAutoFillOptionsOnExternalFocus,
  closeAutoFillOptionsMenuOnEscape,
  closeAutoFillOptionsMenuOnOutsidePointer,
  commitShapeTextEditorOnOutsidePointer,
  finishAutoFillDrag,
  handleAutoFillHandleDoubleClick,
  handleAutoFillOptionClick,
  handleAutoFillOptionsMenuKeydown,
  handleGlobalAutoFillOptionsKeydown,
  handleLineShapeContextMenuCapture,
  handleLineShapePointerDownCapture,
  handlePictureFormatDoubleClick,
  handleShapeTextEditorKeydown,
  insertLineBreakInTextControl,
  isAltModifierEvent,
  isEnterKeyEvent,
  refreshWorkbookChartsAfterSourceEdit,
  refreshWorkbookChartsFromSources,
  selectImageFromEvent,
  shiftFormulaReferences,
  startAutoFillDrag,
  startCellRangeSelection,
  startShapeTextEdit,
  suppressAutoFillRightDragContextMenu,
  toggleAutoFillOptionsMenu,
  updateAutoFillDrag,
} from "../editing/autofill.js";
import {
  applyRibbonBorder,
  applyRibbonColor,
  applyRibbonUnderline,
  closeRibbonBorderMenu,
  closeRibbonBorderMenuOnEscape,
  closeRibbonBorderMenuOnOutsidePointer,
  closeRibbonColorPalette,
  closeRibbonColorPaletteOnEscape,
  closeRibbonColorPaletteOnOutsidePointer,
  closeRibbonInkEraserMenu,
  closeRibbonInkEraserMenuOnEscape,
  closeRibbonInkEraserMenuOnOutsidePointer,
  closeRibbonShapeMenu,
  closeRibbonShapeMenuOnEscape,
  closeRibbonShapeMenuOnOutsidePointer,
  closeRibbonUnderlineMenu,
  closeRibbonUnderlineMenuOnEscape,
  closeRibbonUnderlineMenuOnOutsidePointer,
  closeShapeFormatMenu,
  closeShapeFormatMenuOnEscape,
  closeShapeFormatMenuOnOutsidePointer,
  closeShapeFormatPaneOnEscape,
  closeShapeSelectionPane,
  closeShapeSelectionPaneOnEscape,
  closeSymbolDialogOnEscape,
  finishShapeSelectionPaneDrag,
  handleRibbonBorderChoice,
  handleRibbonColorChoice,
  handleRibbonColorMore,
  handleRibbonColorNone,
  handleRibbonInkEraserChoice,
  handleRibbonShapeChoice,
  handleRibbonStyleChange,
  handleRibbonUnderlineChoice,
  handleShapeFormatMenuAction,
  handleShapeFormatPaneAction,
  handleShapeFormatPaneFieldChange,
  handleShapeSelectionPaneAction,
  handleShapeSelectionPaneObjectClick,
  handleShapeSelectionPaneObjectDoubleClick,
  handleShapeSelectionPaneObjectKeydown,
  handleShapeSelectionPaneRenameBlur,
  handleShapeSelectionPaneRenameKeydown,
  handleShapeStylePresetClick,
  handleSymbolChoice,
  handleSymbolDialogAction,
  isLocalStandaloneDocument,
  openRibbonBorderMenu,
  openRibbonColorPalette,
  openRibbonShapeMenu,
  openRibbonUnderlineMenu,
  openShapeFormatMenu,
  startShapeSelectionPaneDrag,
  uniqueWorksheetObjectId,
  updateShapeSelectionPaneDrag,
} from "../editing/ink-and-clipboard.js";
import {
  clearCellEditDirectTyping,
  copySelectedRange,
  handleSelectionKeyboardSinkBeforeInput,
  handleSelectionKeyboardSinkBlur,
  handleSelectionKeyboardSinkCompositionEnd,
  handleSelectionKeyboardSinkCompositionStart,
  handleSelectionKeyboardSinkCompositionUpdate,
  handleSelectionKeyboardSinkInput,
  handleSelectionKeyboardSinkKeydown,
  handleTextInputImeCompositionEnd,
  handleTextInputImeCompositionStart,
  handleWorkbookCopyEvent,
  handleWorkbookCutEvent,
  handleWorkbookKeyboardShortcut,
  handleWorkbookPasteEvent,
  insertKeyboardDateTimeIntoEditable,
  isImeCompositionEvent,
  keyboardDateTimeShortcutKind,
  markDirectTypingComposition,
  markImeComposition,
  positionSelectionKeyboardSink,
  reconcileDirectTypingImeCommit,
  recordDirectTypingCompositionData,
  restorePersistedInternalClipboard,
  selectColumnsByKeyboard,
  selectRowsByKeyboard,
} from "../editing/keyboard-and-paste.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  applyHeaderAutoFitToSheet,
  autoFitColumnHeaderFromDoubleClick,
  autoFitRowHeaderFromDoubleClick,
  beginFormulaBarEdit,
  cancelFormulaBarEdit,
  captureDimensionForHistory,
  closeStatusBarContextMenu,
  closeStatusBarContextMenuOnEscape,
  closeStatusBarContextMenuOnOutsidePointer,
  commitCellValue,
  commitEditableCellToSelection,
  commitFormulaBar,
  extendCellRangeSelection,
  finishCellRangeSelection,
  finishFormulaBarResize,
  finishHeaderResize,
  getCellRawInput,
  getDisplayForCell,
  handleStatusBarMenuItem,
  isSingleCellRange,
  loadStatusBarItems,
  materializeSelectionRangeForSheet,
  normalizeRange,
  openStatusBarContextMenu,
  rangeToLabel,
  rangesToLabel,
  startColumnHeaderResize,
  startFormulaBarResize,
  startRowHeaderResize,
  syncFormulaEditPreview,
  toggleFormulaBarLines,
  updateFormulaBarResize,
  updateHeaderResize,
  updateSelectionUi,
  updateStatusBarControlVisibility,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  beginNameBoxEdit,
  clearCellEditSelection,
  closeNameBoxMenu,
  closeNameBoxMenuOnOutsidePointer,
  columnSelectionRange,
  editablePlainText,
  extendHeaderSelection,
  finishNameBoxEditOnBlur,
  handleNameBoxKeydown,
  handleNameBoxMenuItemClick,
  handleSheetHostCellClick,
  insertLineBreakInEditable,
  preserveCellEditSelectionForToolbar,
  rowSelectionRange,
  saveCellEditSelection,
  scrollRangeIntoView,
  selectCell,
  startCellEdit,
  startColumnHeaderSelection,
  startEditAutoScroll,
  startRowHeaderSelection,
  stopEditAutoScroll,
  syncEditableCellAlignment,
  updateEditAutoScroll,
  updateFormulaBarForSelection,
  updateHeaderSelectionDrag,
} from "../editing/selection-and-cell-text.js";
import {
  applyCellStructureOperation,
  captureCellsForHistory,
  captureSheetForHistory,
  cloneCellModel,
  cloneDefinedNames,
  cloneJsonLikeValue,
  cloneMerges,
  defaultColumnModel,
  defaultRowModel,
  internCellsInRange,
  pushHistory,
  sheetDisplayCell,
  stableObjectKey,
  workbookMemoryStats,
} from "../editing/structure-and-history.js";
import {
  formulaAutocompleteCandidates,
  formulaAutocompleteDefinitions,
  formulaAutocompleteQueryContext,
  formulaAutocompleteShouldShowCandidates,
  formulaParameterHintContext,
  handleFormulaAutocompleteItemClick,
  handleFormulaAutocompleteKeydown,
  scheduleFormulaAutocompleteRefreshForEditor,
  updateFormulaAutocompleteForEditor,
} from "../formula/autocomplete.js";
import {
  applyPasteNameDialog,
  cellEditBlurStartsFormulaSheetNavigation,
  clearFormulaReferenceState,
  closeFormulaAutocomplete,
  closeFormulaAutocompleteOnEscape,
  closeFormulaAutocompleteOnOutsidePointer,
  cycleFormulaReferenceInEditable,
  cycleFormulaReferenceInFormulaBar,
  finishFormulaReferenceDrag,
  handleFormulaFunctionMenuAction,
  handleFormulaNameMenuAction,
  handleFunctionArgumentInput,
  handleFunctionArgumentKeydown,
  handleInsertFunctionCategoryChange,
  handleInsertFunctionSearchInput,
  handleInsertFunctionSearchKeydown,
  handleNameManagerFilterChange,
  handleNameManagerListChange,
  markFormulaReferenceSheetTabPointerDownCapture,
  preserveCellFormulaEditForSheetNavigation,
  syncPasteNameDialog,
  updateFormulaReferenceDrag,
  updateFormulaReferenceUi,
} from "../formula/dialogs-and-names.js";
import {
  finishPageBreakDrag,
  finishPrintAreaDrag,
  forEachCellInRange,
  handlePageLayoutHeaderFooterDblClick,
  handlePageLayoutWhiteSpaceDblClick,
  mergeSheetRange,
  rangeIntersects,
  rangePixelRect,
  removeIntersectingMerges,
  renderSheet,
  startPageBreakDrag,
  startPrintAreaDrag,
  updatePageBreakDrag,
  updatePrintAreaDrag,
} from "../rendering/grid-and-canvas.js";
import {
  applyHeaderHiddenToSheet,
  applyHeaderSize,
  cancelCloseCommentThreadPopup,
  cancelCloseNotePopup,
  clipboardCellEntries,
  closeAnnotationEditorOnEscape,
  closeCellContextMenu,
  closeCellContextMenuOnOutsidePointer,
  closeCellStructureDialog,
  closeCellStructureDialogOnEscape,
  closeFormatCellsDialog,
  closeFormatDialogOnEscape,
  closeHeaderSizeDialog,
  closeHeaderSizeDialogOnEscape,
  closeLinkDialog,
  closeLinkDialogOnEscape,
  closeNameDialog,
  closeNameDialogOnEscape,
  commitAnnotationEditorOnOutsidePointer,
  handleAnnotationEditorAction,
  handleAnnotationEditorKeydown,
  handleCellContextMenuAction,
  handleCellHyperlinkClick,
  handleCellStructureDialogAction,
  handleCommentThreadAction,
  handleCommentThreadKeydown,
  handleFormatDialogAction,
  handleFormatDialogFieldChange,
  handleFormatDialogTabClick,
  handleHeaderSizeDialogAction,
  handleHeaderSizeDialogKeydown,
  handleLinkDialogAction,
  handleLinkDocumentNameClick,
  handleLinkDocumentSheetClick,
  handleLinkKindClick,
  handleNameDialogAction,
  openCommentThreadPopupFromMarker,
  openFormatCellsDialog,
  openNotePopupFromMarker,
  scheduleCloseCommentThreadPopup,
  scheduleCloseNotePopup,
  setHeaderHidden,
} from "../ui/annotations-and-cell-format.js";
import {
  activeRuntimeVersionInfo,
  cellKey,
  closeMessageBoxOnEscape,
  columnName,
  escapeAttr,
  escapeHtml,
  finishDialogDrag,
  handleDialogChromeCloseClick,
  handleHelpDialogAction,
  handleHelpDialogKeydown,
  handleMessageBoxButton,
  handleMessageBoxKeydown,
  handleVersionDialogAction,
  handleVersionDialogKeydown,
  initDialogChromeObserver,
  readableLoadError,
  refreshServerVersionInfo,
  rowHeightToPixels,
  safeJson,
  sanitizeFileName,
  setStatus,
  showAppAlert,
  showMessageBox,
  startDialogDrag,
  syncVersionRibbonBadge,
  updateDialogDrag,
  versionInfoAppVersion,
  versionInfoRuntimeVersion,
} from "../ui/help-and-utilities.js";
import {
  cancelObjectSelectionModeDrag,
  clearCellModelForHome,
  closeHomeCommandMenu,
  closeHomeCommandMenuOnEscape,
  closeHomeCommandMenuOnOutsidePointer,
  finishInkDrawing,
  finishObjectSelectionModeDrag,
  handleFindDialogAction,
  handleFindSelectMenuAction,
  handleFormulaCalculationMenuAction,
  handleGlobalCommandMenuAccessKey,
  handleGoToDialogAction,
  handleGoToDialogKeydown,
  handleHomeAutoSumMenuAction,
  handleHomeClearMenuAction,
  handleHomeFillMenuAction,
  handleHomeSeriesDialogAction,
  handleHomeSeriesDialogKeydown,
  handleHomeSeriesDialogSync,
  handleHomeSortFilterMenuAction,
  handleHomeWorksheetFillDialogAction,
  handleHomeWorksheetFillDialogKeydown,
  handlePageSetupTabClick,
  handleRibbonCommandMenuKeydown,
  handleSpecialDialogAction,
  handleSpecialDialogKeydown,
  homeClearLabel,
  startObjectSelectionModeDrag,
  updateInkDrawing,
  updateObjectSelectionModeDrag,
} from "../ui/printing-and-home-tools.js";
import {
  closeFileSideMenu,
  closeFileSideMenuOnEscape,
  closeFileSideMenuOnOutsidePointer,
  closeObjectLinkDialogOnEscape,
  closeRibbonDisplayMenu,
  closeRibbonDisplayMenuOnEscape,
  closeRibbonDisplayMenuOnOutsidePointer,
  ensureSheetSize,
  handleObjectLinkDialogAction,
  handleObjectLinkDocumentNameClick,
  handleObjectLinkDocumentSheetClick,
  handleObjectLinkKindClick,
  handleRibbonCommand,
  handleRibbonDisplayMenuKeydown,
  handleRibbonLayoutOptionClick,
  handleRibbonPanelKeydown,
  handleRibbonTabKeydown,
  handleRibbonVisibilityOptionClick,
  handleWorkbookKeyboardCaptureShortcut,
  handleWorkbookPaneFocusShortcut,
  hideRibbonForAutoHideOnOutsidePointer,
  loadRibbonDisplayPreferences,
  releaseRibbonPointerFocus,
  resetWorkbookModifierState,
  scheduleViewportRender,
  switchRibbonTab,
  syncRibbonDisplayState,
  toggleRibbonDisplayMenu,
  updateWorkbookModifierState,
  upsertDefinedName,
} from "../ui/ribbon-and-links.js";
import {
  adjustActiveSheetZoom,
  clampZoom,
  cloneSheetModel,
  closeQuickAnalysisPopup,
  closeQuickAnalysisPopupOnEscape,
  closeQuickAnalysisPopupOnOutsidePointer,
  closeSheetNavigatorMenuOnOutsidePointer,
  closeSheetTabContextMenu,
  closeSheetTabContextMenuOnEscape,
  closeSheetTabContextMenuOnOutsidePointer,
  closeTranslationPane,
  closeTranslationPaneOnEscape,
  closeZoomDialog,
  closeZoomDialogOnEscape,
  closeZoomDialogOnOutsidePointer,
  ensureEditableRangeForProtection,
  ensureSheetProtectionPermission,
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
  openCellContextMenu,
  openColumnHeaderContextMenu,
  openImageContextMenu,
  openRowHeaderContextMenu,
  openSheetNavigatorMenu,
  openSheetTabContextMenu,
  openZoomDialog,
  renderTranslationPane,
  scrollSheetTabs,
  selectCustomZoomOption,
  setActiveSheetView,
  setActiveSheetZoom,
  zoomSelectionFitRange,
  zoomSelectionToFit,
} from "../ui/sheet-protection-and-translation.js";
import {
  activeSheetView,
  captureWorkbookViewState,
  exitFullScreenView,
  finishSplitPaneDrag,
  formulaWatchEntries,
  handleFormulaAuditMenuAction,
  handleFormulaWatchRowDoubleClick,
  handleFormulaWatchWindowAction,
  handleFullScreenEscape,
  handleFullscreenChange,
  handleNameManagerRowDoubleClick,
  refreshActiveSheetViewSnapshot,
  sheetHeadingsVisible,
  startSplitPaneDrag,
  syncViewRibbonButtons,
  toggleShowFormulasFromRibbon,
  updateSplitPaneDrag,
  workbookCalculationMode,
} from "../ui/views-and-formula-tools.js";
import {
  buildStandaloneHtml,
  captureCwsAgentVisualEvidence,
  normalizeRuntimeBaseUrl,
  standaloneDocumentIconUrl,
  standaloneLoaderScript,
  standaloneLocalSaveEndpointUrl,
  standaloneRemoteRuntimeBaseUrl,
  standaloneRuntimeBaseUrl,
  standaloneRuntimeVersionInfo,
  standaloneRuntimeVersionedBaseUrl,
  standaloneVersionedRemoteRuntimeBaseUrl,
} from "./export-model-and-measurement.js";
import {
  addBlankSheet,
  buildFormulaEngine,
  commitSheetTabRenameOnOutsidePointer,
  ensureActiveSheetVisible,
  finishSheetTabDrag,
  handleFormulaReferenceSheetTabPointerDown,
  handleSheetInsertTemplateDoubleClick,
  handleSheetMoveCopyListConfirm,
  handleSheetTabClick,
  handleSheetTabDoubleClick,
  handleSheetTabDragMove,
  handleSheetTabKeydown,
  handleSheetTabRenameInput,
  handleSheetTabRenameInputBlur,
  handleSheetTabRenameInputKeydown,
  handleSheetUnhideListConfirm,
  renderWorkbook,
  setActiveSheet,
  startSheetTabDrag,
  stopSheetTabRenameInputPropagation,
} from "./formulas-and-sheet-tabs.js";
import {
  consumePendingFileLoadRecoveryMarker,
  deletePersistedWorkbookSnapshot,
  extractStandaloneWorkbookPayload,
  flushWorkbookPersistence,
  handleWorkbookBeforeUnload,
  hasStandaloneWorkbookSheets,
  idbRequest,
  idbTransactionDone,
  initializeStandaloneLauncher,
  loadExcelFile,
  loadStandaloneOpenWorkbookFile,
  loadStandaloneWorkbookModel,
  markWorkbookSavedClean,
  normalizeMacroSecurity,
  normalizePaneState,
  normalizePersistedCellDateValues,
  normalizedSheetShowRuler,
  normalizedSheetViewMode,
  openBlankWorkbook,
  restorePersistedWorkbook,
  scheduleWorkbookPersistence,
  standaloneOpenWorkbookFileFromWindow,
  standaloneWorkbookModelFromWindow,
  standaloneWorkbookViewDiagnostics,
  workbookHasUnsavedChanges,
} from "./import-model.js";
import {
  parseTableStyleInfoXml,
} from "./ooxml-drawings.js";
import {
  activeSheet,
  cellDisplayText,
  displayValue,
  exportStandaloneHtml,
  fileNameWithExtension,
  getCellFormula,
  getSheetId,
  handleSaveAsFileRowClick,
  handleSaveAsLocationClick,
  handleSaveAsPathChange,
  handleSaveAsTypeChange,
  quickSaveWorkbook,
  rememberStandaloneSaveFileHandle,
  saveAsFileTypeByValue,
  standaloneLocalSaveEndpoint,
  standaloneOpenedFilePath,
  standaloneOpenedFilePickerStartIn,
  standaloneSaveHandleTokenKey,
  startSheetFooterResize,
  syncHorizontalScrollbarFromHost,
  syncHostFromHorizontalScrollbar,
  updateHorizontalScrollbar,
  visibleGridRange,
  workbookBlobForSaveType,
} from "./saving-and-export.js";
import {
  bytesToBase64,
  validationListOptions,
} from "./tables-and-validation.js";
import {
  applyPageBackgroundFile,
  commitPageLayoutChange,
  ensurePageLayoutSettings,
  handlePageLayoutFieldChange,
  handlePageLayoutMenuAction,
  handlePageLayoutSheetOptionChange,
  handleSheetCodeProcedureChange,
  normalizedPageLayoutSettings,
  normalizedPageLayoutThemeEffects,
} from "./vba-and-printing.js";
import {
  closeViewCommandMenu,
  closeViewCommandMenuOnEscape,
  closeViewCommandMenuOnOutsidePointer,
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
  openPageBreakContextMenu,
  openPageBreakContextMenuFromPoint,
  parseVbaMacroStatements,
  saveActiveWindowViewState,
  scheduleWorksheetFocusAfterCommandPopup,
  syncPairedWorkbookWindowScrollState,
  updateActiveWorkbookWindowScrollState,
} from "./views-and-macros.js";


function applyCwsAgentResetPageBreaksOp(op) {
  const sheet = activateCwsAgentPageSetupSheet(op);
  const layout = ensurePageLayoutSettings(sheet);
  if (!(layout.pageBreaks?.rows?.length || layout.pageBreaks?.cols?.length)) {
    renderWorkbook();
    setStatus("CWS Agent: 手動改ページはありません。");
    return {
      type: "resetPageBreaks",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      address: sheet.name,
    };
  }
  commitPageLayoutChange((layout) => {
    layout.pageBreaks = { rows: [], cols: [] };
  }, `${sheet.name} CWS Agent 改ページのリセット`, "CWS Agent: すべての改ページを解除しました。", {
    captureWorkbookView: true,
  });
  return {
    type: "resetPageBreaks",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    address: sheet.name,
  };
}



function cwsAgentPageBreakPositionLabel(axis, index) {
  return axis === "row" ? `${index}行目` : `${columnName(index)}列`;
}



function activateCwsAgentPageSetupSheet(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }
  return state.model.sheets[op.sheetIndex];
}



function cwsAgentFreezePanesStatus(pane) {
  const rowText = pane.rows > 0 ? `${pane.rows}行` : "";
  const columnText = pane.columns > 0 ? `${pane.columns}列` : "";
  if (rowText && columnText) return `${rowText}と${columnText}を`;
  if (rowText) return `${rowText}を`;
  if (columnText) return `${columnText}を`;
  return "ウィンドウ枠を";
}



function cwsAgentSheetViewSnapshot(sheet) {
  return {
    viewMode: normalizedSheetViewMode(sheet?.viewMode),
    zoomScale: clampZoom(sheet?.zoomScale || 100),
    showGridLines: sheet?.showGridLines !== false,
    showHeadings: sheet?.showHeadings !== false,
    showRuler: normalizedSheetShowRuler(sheet?.showRuler, sheet?.viewMode),
    showFormulas: sheet?.showFormulas === true,
    showZeros: sheet?.showZeros !== false,
    showWhiteSpace: sheet?.showWhiteSpace !== false,
    showOutlineSymbols: sheet?.showOutlineSymbols !== false,
    showObjects: sheet?.showObjects !== false,
    rightToLeft: sheet?.rightToLeft === true,
  };
}



function cwsAgentSheetViewHasChanges(current, next) {
  return Object.keys(next || {}).some((key) => current[key] !== next[key]);
}



function cwsAgentSheetViewPatchLabels(patch = {}, sheet = null) {
  const labels = [];
  const viewLabels = {
    normal: "標準表示",
    "page-layout": "ページレイアウト表示",
    "page-break": "改ページプレビュー",
  };
  if (Object.prototype.hasOwnProperty.call(patch, "viewMode")) labels.push(viewLabels[normalizedSheetViewMode(patch.viewMode)] || patch.viewMode);
  if (Object.prototype.hasOwnProperty.call(patch, "zoomScale")) labels.push(`ズーム ${clampZoom(patch.zoomScale)}%`);
  if (Object.prototype.hasOwnProperty.call(patch, "showGridLines")) labels.push(`目盛線${patch.showGridLines ? "表示" : "非表示"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "showHeadings")) labels.push(`見出し${patch.showHeadings ? "表示" : "非表示"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "showRuler")) labels.push(`ルーラー${patch.showRuler ? "表示" : "非表示"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "showFormulas")) labels.push(patch.showFormulas ? "数式表示" : "計算結果表示");
  if (Object.prototype.hasOwnProperty.call(patch, "showZeros")) labels.push(`ゼロ値${patch.showZeros ? "表示" : "非表示"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "showWhiteSpace")) labels.push(`ページ余白${patch.showWhiteSpace ? "表示" : "非表示"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "showOutlineSymbols")) labels.push(`アウトライン記号${patch.showOutlineSymbols ? "表示" : "非表示"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "showObjects")) labels.push(`オブジェクト${patch.showObjects ? "表示" : "非表示"}`);
  if (Object.prototype.hasOwnProperty.call(patch, "rightToLeft")) labels.push(patch.rightToLeft ? "右から左" : "左から右");
  return labels.join("、") || (sheet?.name ? `${sheet.name} 表示設定` : "");
}



function commitCwsAgentSheetViewChange(sheet, sheetBefore, statusMessage, label, workbookViewBefore = null) {
  refreshActiveSheetViewSnapshot(sheet);
  saveActiveWindowViewState();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    ...(workbookViewBefore ? {
      workbookViewBefore,
      workbookViewAfter: captureWorkbookViewState(),
    } : {}),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label,
  });
  renderWorkbook();
  updateFormulaBarForSelection();
  syncViewRibbonButtons();
  setStatus(statusMessage);
  scheduleWorkbookPersistence();
}



function refreshCwsAgentWorkbookAfterSheetStructureChange(options = {}) {
  if (!options.preserveSheetSelection) state.selectedSheetIndexes = [];
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeTranslationPane();
  state.annotationEditor = null;
  state.formulaEdit = null;
  state.hf = buildFormulaEngine(state.model);
  ensureActiveSheetVisible();
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  updateSelectionUi();
  updateFormulaBarForSelection();
}



function applyCwsAgentSetColumnWidthOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatColumns", "保護されたシートでは列の幅を変更できません。")) {
    throw new Error("保護されたシートでは列の幅を変更できません。");
  }
  ensureSheetSize(sheet, Math.max(1, Number(sheet.rowCount) || 1), op.right);
  const range = {
    sheetIndex: op.sheetIndex,
    top: 1,
    left: op.left,
    bottom: Math.max(1, Number(sheet.rowCount) || 1),
    right: op.right,
    fullColumn: true,
  };
  const before = [];
  const after = [];
  for (let col = op.left; col <= op.right; col += 1) {
    before.push(captureDimensionForHistory(sheet, "col", col));
    sheet.columns[col - 1] = { ...defaultColumnModel(sheet), ...(sheet.columns[col - 1] || {}), width: op.width };
    after.push(captureDimensionForHistory(sheet, "col", col));
  }

  state.selected = { sheetIndex: op.sheetIndex, row: 1, col: op.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  if (cwsAgentDimensionsChanged(before, after)) {
    pushHistory({
      sheetIndex: op.sheetIndex,
      dimensionsBefore: before,
      dimensionsAfter: after,
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: [range],
      selectionAfterRanges: [range],
      label: `${rangeToLabel(range)} CWS Agent 列の幅`,
    });
  }
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} の列幅を ${op.width}px にしました。`);

  return {
    type: "setColumnWidth",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    left: op.left,
    right: op.right,
    width: op.width,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSetRowHeightOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatRows", "保護されたシートでは行の高さを変更できません。")) {
    throw new Error("保護されたシートでは行の高さを変更できません。");
  }
  ensureSheetSize(sheet, op.bottom, Math.max(1, Number(sheet.colCount) || 1));
  const range = {
    sheetIndex: op.sheetIndex,
    top: op.top,
    left: 1,
    bottom: op.bottom,
    right: Math.max(1, Number(sheet.colCount) || 1),
    fullRow: true,
  };
  const before = [];
  const after = [];
  for (let row = op.top; row <= op.bottom; row += 1) {
    before.push(captureDimensionForHistory(sheet, "row", row));
    sheet.rows[row - 1] = { ...defaultRowModel(sheet), ...(sheet.rows[row - 1] || {}), height: op.height };
    after.push(captureDimensionForHistory(sheet, "row", row));
  }

  state.selected = { sheetIndex: op.sheetIndex, row: op.top, col: 1 };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  if (cwsAgentDimensionsChanged(before, after)) {
    pushHistory({
      sheetIndex: op.sheetIndex,
      dimensionsBefore: before,
      dimensionsAfter: after,
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: [range],
      selectionAfterRanges: [range],
      label: `${rangeToLabel(range)} CWS Agent 行の高さ`,
    });
  }
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} の行高を ${op.height}px にしました。`);

  return {
    type: "setRowHeight",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: op.top,
    bottom: op.bottom,
    height: op.height,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentAutoFitColumnsOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatColumns", "保護されたシートでは列の幅を自動調整できません。")) {
    throw new Error("保護されたシートでは列の幅を自動調整できません。");
  }
  ensureSheetSize(sheet, Math.max(1, Number(sheet.rowCount) || 1), op.right);
  const range = {
    sheetIndex: op.sheetIndex,
    top: 1,
    left: op.left,
    bottom: Math.max(1, Number(sheet.rowCount) || 1),
    right: op.right,
    fullColumn: true,
  };
  const sheetBefore = captureSheetForHistory(sheet);
  const changed = applyHeaderAutoFitToSheet(sheet, "column", [range]);

  state.selected = { sheetIndex: op.sheetIndex, row: 1, col: op.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  if (changed) {
    pushHistory({
      sheetIndex: op.sheetIndex,
      sheetBefore,
      sheetAfter: captureSheetForHistory(sheet),
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: [range],
      selectionAfterRanges: [range],
      label: `${rangeToLabel(range)} CWS Agent 列の幅 自動調整`,
    });
  }
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  updateHorizontalScrollbar();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} の列幅を自動調整しました。`);

  return {
    type: "autoFitColumns",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    left: op.left,
    right: op.right,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentAutoFitRowsOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatRows", "保護されたシートでは行の高さを自動調整できません。")) {
    throw new Error("保護されたシートでは行の高さを自動調整できません。");
  }
  ensureSheetSize(sheet, op.bottom, Math.max(1, Number(sheet.colCount) || 1));
  const range = {
    sheetIndex: op.sheetIndex,
    top: op.top,
    left: 1,
    bottom: op.bottom,
    right: Math.max(1, Number(sheet.colCount) || 1),
    fullRow: true,
  };
  const sheetBefore = captureSheetForHistory(sheet);
  const changed = applyHeaderAutoFitToSheet(sheet, "row", [range]);

  state.selected = { sheetIndex: op.sheetIndex, row: op.top, col: 1 };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  if (changed) {
    pushHistory({
      sheetIndex: op.sheetIndex,
      sheetBefore,
      sheetAfter: captureSheetForHistory(sheet),
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: [range],
      selectionAfterRanges: [range],
      label: `${rangeToLabel(range)} CWS Agent 行の高さ 自動調整`,
    });
  }
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} の行高を自動調整しました。`);

  return {
    type: "autoFitRows",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: op.top,
    bottom: op.bottom,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentHeaderHiddenOp(op, kind, hidden) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const isColumn = kind === "column";
  const permission = isColumn ? "allowFormatColumns" : "allowFormatRows";
  const protectedMessage = isColumn
    ? "保護されたシートでは列の表示/非表示を変更できません。"
    : "保護されたシートでは行の表示/非表示を変更できません。";
  if (!ensureSheetProtectionPermission(sheet, permission, protectedMessage)) {
    throw new Error(protectedMessage);
  }

  if (isColumn) {
    ensureSheetSize(sheet, Math.max(1, Number(sheet.rowCount) || 1), op.right);
  } else {
    ensureSheetSize(sheet, op.bottom, Math.max(1, Number(sheet.colCount) || 1));
  }
  const range = isColumn
    ? {
        sheetIndex: op.sheetIndex,
        top: 1,
        left: op.left,
        bottom: Math.max(1, Number(sheet.rowCount) || 1),
        right: op.right,
        fullColumn: true,
      }
    : {
        sheetIndex: op.sheetIndex,
        top: op.top,
        left: 1,
        bottom: op.bottom,
        right: Math.max(1, Number(sheet.colCount) || 1),
        fullRow: true,
      };
  const sheetBefore = captureSheetForHistory(sheet);
  const changed = applyHeaderHiddenToSheet(sheet, isColumn ? "column" : "row", [range], hidden);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  if (changed) {
    state.hf = buildFormulaEngine(state.model);
    pushHistory({
      sheetIndex: op.sheetIndex,
      sheetBefore,
      sheetAfter: captureSheetForHistory(sheet),
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: [range],
      selectionAfterRanges: [range],
      label: `${rangeToLabel(range)} CWS Agent ${hidden ? "非表示" : "再表示"}`,
    });
  }
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  if (isColumn) updateHorizontalScrollbar();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} を${hidden ? "非表示" : "再表示"}にしました。`);

  return {
    type: isColumn ? (hidden ? "hideColumns" : "unhideColumns") : (hidden ? "hideRows" : "unhideRows"),
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    ...(isColumn ? { left: op.left, right: op.right } : { top: op.top, bottom: op.bottom }),
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentOutlineOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const isColumn = op.axis === "columns";
  const protectedMessage = isColumn
    ? "保護されたシートでは列のアウトラインを変更できません。"
    : "保護されたシートでは行のアウトラインを変更できません。";
  if (!ensureOutlineAxisAllowed(sheet, isColumn ? "columns" : "rows")) {
    throw new Error(protectedMessage);
  }

  if (isColumn) {
    ensureSheetSize(sheet, Math.max(1, Number(sheet.rowCount) || 1), op.right);
  } else {
    ensureSheetSize(sheet, op.bottom, Math.max(1, Number(sheet.colCount) || 1));
  }
  const range = isColumn
    ? {
        sheetIndex: op.sheetIndex,
        top: 1,
        left: op.left,
        bottom: Math.max(1, Number(sheet.rowCount) || 1),
        right: op.right,
        fullColumn: true,
      }
    : {
        sheetIndex: op.sheetIndex,
        top: op.top,
        left: 1,
        bottom: op.bottom,
        right: Math.max(1, Number(sheet.colCount) || 1),
        fullRow: true,
      };
  const sheetBefore = captureSheetForHistory(sheet);
  applyOutlineToSheet(sheet, range, op.delta, isColumn);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent アウトライン`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  if (isColumn) updateHorizontalScrollbar();
  const action = op.delta > 0 ? "グループ化" : "グループ解除";
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} を${action}しました。`);

  return {
    type: op.type,
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    ...(isColumn ? { left: op.left, right: op.right } : { top: op.top, bottom: op.bottom }),
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function cwsAgentDimensionsChanged(before, after) {
  return before.some((item, index) => Math.abs(Number(item?.size || 0) - Number(after[index]?.size || 0)) > 0.01);
}



function applyCwsAgentRowsOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const count = Math.min(op.count, op.type === "deleteRows" ? Math.max(0, sheet.rowCount - op.row + 1) : op.count);
  if (count < 1) {
    throw new Error("対象行がありません。");
  }
  const range = {
    sheetIndex: op.sheetIndex,
    top: op.row,
    left: 1,
    bottom: op.row + count - 1,
    right: Math.max(1, Number(sheet.colCount) || 1),
    fullRow: true,
  };
  const kind = op.type === "insertRows" ? "insert" : "delete";
  const applied = applyCellStructureOperation(kind, "entire-row", range);
  if (!applied) {
    throw new Error(op.type === "insertRows" ? "行を挿入できませんでした。" : "行を削除できませんでした。");
  }
  const updatedSheet = state.model.sheets[op.sheetIndex];
  return {
    type: op.type,
    sheetIndex: op.sheetIndex,
    sheetName: updatedSheet.name,
    row: op.row,
    count,
    address: `${updatedSheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentColumnsOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const count = Math.min(op.count, op.type === "deleteColumns" ? Math.max(0, sheet.colCount - op.col + 1) : op.count);
  if (count < 1) {
    throw new Error("対象列がありません。");
  }
  const range = {
    sheetIndex: op.sheetIndex,
    top: 1,
    left: op.col,
    bottom: Math.max(1, Number(sheet.rowCount) || 1),
    right: op.col + count - 1,
    fullColumn: true,
  };
  const kind = op.type === "insertColumns" ? "insert" : "delete";
  const applied = applyCellStructureOperation(kind, "entire-column", range);
  if (!applied) {
    throw new Error(op.type === "insertColumns" ? "列を挿入できませんでした。" : "列を削除できませんでした。");
  }
  const updatedSheet = state.model.sheets[op.sheetIndex];
  return {
    type: op.type,
    sheetIndex: op.sheetIndex,
    sheetName: updatedSheet.name,
    col: op.col,
    count,
    address: `${updatedSheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentClearRangeOp(op) {
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
  if (!ensureEditableRangeForProtection(sheet, range, homeClearLabel(op.mode))) {
    throw new Error("保護されたセルはクリアできません。");
  }

  const before = captureCellsForHistory(sheet, range);
  forEachCellInRange(range, (row, col) => {
    clearCellModelForHome(sheet, cellKey(row, col), op.mode);
  });
  internCellsInRange(sheet, range);
  refreshWorkbookChartsAfterSourceEdit(op.sheetIndex, range);
  state.hf = buildFormulaEngine(state.model);
  const after = captureCellsForHistory(sheet, range);
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
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent ${homeClearLabel(op.mode)}`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} の${homeClearLabel(op.mode)}を実行しました。`);

  return {
    type: "clearRange",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    mode: op.mode,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentMergeCellsOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではセルを結合できません。")) {
    throw new Error("保護されたシートではセルを結合できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  if (isSingleCellRange(range)) {
    throw new Error("結合するには 2 つ以上のセルが必要です。");
  }

  ensureSheetSize(sheet, range.bottom, range.right);
  const before = captureCellsForHistory(sheet, range);
  const mergesBefore = cloneMerges(sheet.merges);
  mergeSheetRange(sheet, range, op.center);
  state.hf = buildFormulaEngine(state.model);
  const after = captureCellsForHistory(sheet, range);
  const mergesAfter = cloneMerges(sheet.merges);
  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    mergesBefore,
    mergesAfter,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent セルの結合`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} を結合しました${op.center ? "（中央揃え）" : ""}。`);

  return {
    type: "mergeCells",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    center: op.center,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentUnmergeCellsOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではセルの結合を解除できません。")) {
    throw new Error("保護されたシートではセルの結合を解除できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const before = captureCellsForHistory(sheet, range);
  const mergesBefore = cloneMerges(sheet.merges);
  const beforeCount = sheet.merges.length;
  removeIntersectingMerges(sheet, range);
  const removed = beforeCount - sheet.merges.length;
  state.hf = buildFormulaEngine(state.model);
  const after = captureCellsForHistory(sheet, range);
  const mergesAfter = cloneMerges(sheet.merges);
  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    mergesBefore,
    mergesAfter,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent 結合の解除`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(removed ? `CWS Agent: ${sheet.name}!${rangeToLabel(range)} の結合を解除しました。` : "CWS Agent: 対象範囲には結合セルがありません。");

  return {
    type: "unmergeCells",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    removed,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function init() {
  state.isStandalone = Boolean(window.__WEBSHEET_STANDALONE__);
  state.statusBarItems = loadStatusBarItems();
  loadRibbonDisplayPreferences();
  updateStatusBarControlVisibility();
  createIcons({ icons });
  syncRibbonDisplayState();
  syncVersionRibbonBadge();
  initWorkbookNameDisplay();
  initDialogChromeObserver();
  void refreshServerVersionInfo();
  initCwsAgentBridge({
    $,
    icons,
    createIcons,
    state,
    helpers: {
      activeSheet,
      activeSelectionRange,
      visibleGridRange,
      materializeSelectionRangeForSheet,
      getSheetId,
      getCellRawInput,
      getDisplayForCell,
      getCellValidationInfo: cwsAgentCellValidationInfo,
      shapeEvidenceBounds: imageBoxToPixels,
      columnName,
      rangeToLabel,
      escapeHtml,
      escapeAttr,
    },
    operations: {
      applyOps: applyCwsAgentOps,
      previewOps: previewCwsAgentOps,
      previewOpsProgressively: previewCwsAgentOpsProgressively,
      appendPreviewOpsProgressively: appendCwsAgentPreviewOpsProgressively,
      setPreviewVisible: setCwsAgentPreviewVisible,
      commitPreview: commitCwsAgentPreview,
      clearPreview: clearCwsAgentPreview,
      captureTaskWorkbook: captureCwsAgentTaskWorkbook,
      normalizeTaskSnapshot: normalizeCwsAgentTaskSnapshot,
      restoreTaskWorkbook: restoreCwsAgentTaskWorkbook,
      persistTaskWorkbook: persistCwsAgentTaskWorkbook,
      captureVisualEvidence: captureCwsAgentVisualEvidence,
    },
  });

  $fileInput.on("change", async (event) => {
    const [file] = event.target.files || [];
    if (file) {
      await openSelectedWorkbookFileWithUnsavedPrompt(file);
      $fileInput.val("");
    }
  });

  $pageBackgroundInput.on("change", async (event) => {
    const [file] = event.target.files || [];
    if (file) {
      await applyPageBackgroundFile(file);
      $pageBackgroundInput.val("");
    }
  });

  $dropZone.on("dragover", (event) => {
    event.preventDefault();
    $dropZone.addClass("is-dragging");
  });

  $dropZone.on("dragleave drop", () => {
    $dropZone.removeClass("is-dragging");
  });

  $dropZone.on("drop", async (event) => {
    event.preventDefault();
    const [file] = event.originalEvent?.dataTransfer?.files || [];
    if (file) {
      await openSelectedWorkbookFileWithUnsavedPrompt(file);
    }
  });
  $emptyState.on("click", "[data-open-workbook-file]", () => {
    void openWorkbookFilePicker();
  });
  document.addEventListener("dragover", handleWorkbookFileDragOver, true);
  document.addEventListener("drop", handleWorkbookFileDrop, true);

  $sheetTabs[0]?.addEventListener("pointerdown", markFormulaReferenceSheetTabPointerDownCapture, true);
  $sheetTabs.on("pointerdown", ".sheet-tab", handleFormulaReferenceSheetTabPointerDown);
  $sheetTabs.on("pointerdown", ".sheet-tab", startSheetTabDrag);
  $sheetTabs.on("click", ".sheet-tab", handleSheetTabClick);
  $sheetTabs.on("dblclick", ".sheet-tab", handleSheetTabDoubleClick);
  $sheetTabs.on("keydown", ".sheet-tab", handleSheetTabKeydown);
  $sheetTabs.on("contextmenu", ".sheet-tab", openSheetTabContextMenu);
  $sheetTabs.on("pointerdown click dblclick", ".sheet-tab-rename-input", stopSheetTabRenameInputPropagation);
  $sheetTabs.on("input", ".sheet-tab-rename-input", handleSheetTabRenameInput);
  $sheetTabs.on("keydown", ".sheet-tab-rename-input", handleSheetTabRenameInputKeydown);
  $sheetTabs.on("blur", ".sheet-tab-rename-input", handleSheetTabRenameInputBlur);
  $addSheetButton.on("click", addBlankSheet);
  $sheetTabScrollLeft.on("click", () => scrollSheetTabs(-1));
  $sheetTabScrollRight.on("click", () => scrollSheetTabs(1));
  $sheetTabScrollLeft.on("contextmenu", openSheetNavigatorMenu);
  $sheetTabScrollRight.on("contextmenu", openSheetNavigatorMenu);
  $sheetFooter.on("click", "[data-sheet-view]", (event) => {
    setActiveSheetView(event.currentTarget.dataset.sheetView);
  });
  $zoomSlider.on("input", (event) => {
    setActiveSheetZoom(Number(event.currentTarget.value));
  });
  $zoomOutButton.on("click", () => adjustActiveSheetZoom(-10));
  $zoomInButton.on("click", () => adjustActiveSheetZoom(10));
  $zoomLabel.on("click", openZoomDialog);
  $zoomLabel.on("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    event.preventDefault();
    openZoomDialog(event);
  });
  document.addEventListener("wheel", handleSheetZoomWheel, { passive: false, capture: true });
  document.addEventListener("keydown", updateWorkbookModifierState, true);
  document.addEventListener("keyup", updateWorkbookModifierState, true);
  document.addEventListener("keydown", handleWorkbookKeyboardCaptureShortcut, true);
  document.addEventListener("pointerdown", preserveCellEditSelectionForToolbar, true);
  document.addEventListener("pointerdown", commitShapeTextEditorOnOutsidePointer, true);
  document.addEventListener("contextmenu", openPageBreakContextMenuFromPoint, true);
  $ribbonTabs.on("click", ".ribbon-tab", (event) => {
    switchRibbonTab($(event.currentTarget).data("ribbonTab"));
  });
  $ribbonTabs.on("keydown", ".ribbon-tab", handleRibbonTabKeydown);
  $ribbonDisplayButton.on("click", toggleRibbonDisplayMenu);
  $ribbonDisplayMenu.on("click", "[data-ribbon-layout-option]", handleRibbonLayoutOptionClick);
  $ribbonDisplayMenu.on("click", "[data-ribbon-visibility-option]", handleRibbonVisibilityOptionClick);
  $ribbonDisplayMenu.on("keydown", handleRibbonDisplayMenuKeydown);
  $ribbonPanels.on("keydown", handleRibbonPanelKeydown);
  $ribbonPanels.on("click", "[data-ribbon-command]", (event) => {
    if (event.currentTarget.matches("[data-ribbon-shape-menu]")) return;
    void handleRibbonCommand($(event.currentTarget).data("ribbonCommand"), event.currentTarget, event)
      .finally(() => releaseRibbonPointerFocus(event.currentTarget, event));
  });
  $ribbonPanels.on("click", "[data-ribbon-color-apply]", (event) => {
    applyRibbonColor($(event.currentTarget).data("ribbonColorApply"));
  });
  $ribbonPanels.on("click", "[data-ribbon-color-menu]", (event) => {
    openRibbonColorPalette(event.currentTarget);
  });
  $ribbonPanels.on("click", "[data-ribbon-shape-menu]", openRibbonShapeMenu);
  $ribbonPanels.on("click", "[data-shape-fill-menu]", (event) => openShapeFormatMenu(event, "fill"));
  $ribbonPanels.on("click", "[data-shape-outline-menu]", (event) => openShapeFormatMenu(event, "outline"));
  $ribbonPanels.on("click", "[data-shape-effects-menu]", (event) => openShapeFormatMenu(event, "effects"));
  $ribbonPanels.on("click", "[data-shape-text-fill-menu]", (event) => openShapeFormatMenu(event, "textFill"));
  $ribbonPanels.on("click", "[data-shape-text-outline-menu]", (event) => openShapeFormatMenu(event, "textOutline"));
  $ribbonPanels.on("click", "[data-shape-align-menu]", (event) => openShapeFormatMenu(event, "align"));
  $ribbonPanels.on("click", "[data-shape-rotate-menu]", (event) => openShapeFormatMenu(event, "rotate"));
  $ribbonPanels.on("click", "[data-shape-style-preset]", handleShapeStylePresetClick);
  $fileSideMenu.on("click", "[data-ribbon-command]", (event) => {
    closeFileSideMenu({ restoreHome: true });
    void handleRibbonCommand($(event.currentTarget).data("ribbonCommand"));
  });
  $fileSideMenu.on("keydown", handleRibbonCommandMenuKeydown);
  $ribbonPanels.on("change", "[data-ribbon-style]", (event) => {
    handleRibbonStyleChange(event.currentTarget);
  });
  $ribbonPanels.on("input", "input[type='color'][data-ribbon-style]", (event) => {
    handleRibbonStyleChange(event.currentTarget);
  });
  $ribbonPanels.on("change", "[data-shape-size]", (event) => {
    applySelectedShapeSize(event.currentTarget.dataset.shapeSize, Number(event.currentTarget.value));
  });
  $ribbonPanels.on("change", "[data-shape-rotation]", (event) => {
    applySelectedShapeRotation(Number(event.currentTarget.value));
  });
  $workbookWindowBar.on("click", "[data-workbook-window-close]", handleWorkbookWindowCloseClick);
  $workbookWindowBar.on("click", "[data-workbook-window-id]", handleWorkbookWindowBarClick);

  $sheetHost.on("scroll", () => {
    syncHorizontalScrollbarFromHost();
    positionSelectionKeyboardSink();
    if (state.restoringViewportScroll) return;
    updateActiveWorkbookWindowScrollState();
    syncPairedWorkbookWindowScrollState();
    scheduleViewportRender({ deferTrim: true });
  });
  $horizontalScrollbar.on("scroll", () => {
    syncHostFromHorizontalScrollbar();
    if (state.restoringViewportScroll) return;
    updateActiveWorkbookWindowScrollState();
    syncPairedWorkbookWindowScrollState();
    scheduleViewportRender({ deferTrim: true });
  });
  $sheetFooterSplitter.on("pointerdown", startSheetFooterResize);

  $sheetHost.on("pointerdown", startCellRangeSelection);
  $sheetHost.on("pointermove", updatePendingShapeDrawingConnectionPreview);
  $sheetHost.on("pointerleave", clearPendingShapeDrawingConnectionPreview);
  $sheetHost.on("pointermove", extendCellRangeSelection);
  $sheetHost.on("pointerdown", ".sheet-image-selection-frame", startSelectedImageTransform);
  $sheetHost.on("pointerdown", ".sheet-picture-crop-handle", startPictureCropDrag);
  $sheetHost.on("pointerdown", ".sheet-image, .sheet-image-link", selectImageFromEvent);
  $sheetHost.on("pointerdown", ".sheet-fill-handle", startAutoFillDrag);
  $sheetHost.on("dblclick", ".sheet-fill-handle", handleAutoFillHandleDoubleClick);
  $sheetHost.on("contextmenu", ".sheet-fill-handle", (event) => {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
  });
  $sheetHost.on("contextmenu", suppressAutoFillRightDragContextMenu);
  $sheetHost.on("pointerdown", "#autoFillOptionsButton, #autoFillOptionsMenu", (event) => event.stopPropagation());
  $sheetHost.on("click", "#autoFillOptionsButton", toggleAutoFillOptionsMenu);
  $sheetHost.on("click", "[data-auto-fill-option]", handleAutoFillOptionClick);
  $sheetHost.on("keydown", "#autoFillOptionsMenu", handleAutoFillOptionsMenuKeydown);
  $(document).on("keydown", handleGlobalAutoFillOptionsKeydown);
  $sheetHost.on("pointerdown", ".sheet-view-guide.is-page-break", startPageBreakDrag);
  $sheetHost.on("pointerdown", ".sheet-view-print-area-handle", startPrintAreaDrag);
  $sheetHost.on("pointerdown", ".sheet-view-guide.is-split-pane", startSplitPaneDrag);
  $sheetHost.on("pointerdown", ".column-resize-handle", startColumnHeaderResize);
  $sheetHost.on("pointerdown", ".row-resize-handle", startRowHeaderResize);
  $sheetHost.on("dblclick", ".column-resize-handle", autoFitColumnHeaderFromDoubleClick);
  $sheetHost.on("dblclick", ".row-resize-handle", autoFitRowHeaderFromDoubleClick);
  $sheetHost.on("pointerdown", ".column-header[data-col]", startColumnHeaderSelection);
  $sheetHost.on("pointerenter", ".column-header[data-col]", extendHeaderSelection);
  $sheetHost.on("pointerdown", ".row-header[data-row]", startRowHeaderSelection);
  $sheetHost.on("pointerenter", ".row-header[data-row]", extendHeaderSelection);
  $sheetHost.on("pointerdown", "[data-outline-toggle]", (event) => event.stopPropagation());
  $sheetHost.on("click", "[data-outline-toggle]", handleOutlineToggleClick);
  $sheetHost.on("pointerdown", "[data-outline-level-command]", (event) => event.stopPropagation());
  $sheetHost.on("click", "[data-outline-level-command]", handleOutlineLevelCommandClick);
  $sheetHost.on("dblclick", ".sheet-page-layout-header, .sheet-page-layout-footer", handlePageLayoutHeaderFooterDblClick);
  $sheetHost.on("dblclick", ".sheet-page-layout-white-space-edge", handlePageLayoutWhiteSpaceDblClick);
  $sheetHost.on("contextmenu", ".sheet-view-guide.is-page-break", openPageBreakContextMenu);
  $sheetHost.on("contextmenu", ".column-header[data-col]", openColumnHeaderContextMenu);
  $sheetHost.on("contextmenu", ".row-header[data-row]", openRowHeaderContextMenu);
  $sheetHost.on("contextmenu", ".sheet-image-selection-frame, .sheet-image, .sheet-image-link", openImageContextMenu);
  $sheetHost.on("contextmenu", openPageBreakContextMenuFromPoint);
  $sheetHost.on("contextmenu", openCellContextMenu);
  if ($sheetHost[0]) {
    $sheetHost[0].addEventListener("pointerdown", startObjectSelectionModeDrag, true);
    $sheetHost[0].addEventListener("pointerdown", handleLineShapePointerDownCapture, true);
    $sheetHost[0].addEventListener("contextmenu", handleLineShapeContextMenuCapture, true);
  }
  $sheetHost.on("click", ".cell-link", handleCellHyperlinkClick);
  $sheetHost.on("click", handleSheetHostCellClick);
  $sheetHost.on("dblclick", ".sheet-image:not(.sheet-shape), .sheet-image-link, .sheet-image-selection-frame:not(.is-shape)", handlePictureFormatDoubleClick);
  $sheetHost.on("dblclick", ".sheet-image.sheet-shape, .sheet-image-selection-frame.is-shape", startShapeTextEdit);
  $sheetHost.on("dblclick", startCellEdit);
  $sheetHost.on("pointerdown", ".sheet-cell[contenteditable='true']", startEditAutoScroll);
  $sheetHost.on("mouseenter", ".cell-annotation-marker.is-comment", openCommentThreadPopupFromMarker);
  $sheetHost.on("mouseleave", ".cell-annotation-marker.is-comment", scheduleCloseCommentThreadPopup);
  $sheetHost.on("mouseenter", ".cell-annotation-marker.is-note", openNotePopupFromMarker);
  $sheetHost.on("mouseleave", ".cell-annotation-marker.is-note", scheduleCloseNotePopup);
  $sheetHost.on("pointerdown", ".cell-filter-button", openAutoFilterMenu);
  $sheetHost.on("pointerdown", ".cell-validation-dropdown-button", handleValidationDropdownButtonPointerDown);
  $sheetHost.on("click", ".cell-validation-dropdown-button", handleValidationDropdownButtonClick);
  window.addEventListener("pointermove", updateEditAutoScroll);
  window.addEventListener("pointermove", updateObjectSelectionModeDrag, true);
  window.addEventListener("pointermove", updateInkDrawing);
  window.addEventListener("pointermove", updateShapeDrawing);
  window.addEventListener("pointermove", updateSelectedImageTransform);
  window.addEventListener("pointermove", updatePictureCropDrag);
  window.addEventListener("pointermove", updateShapeSelectionPaneDrag);
  window.addEventListener("pointermove", updateAutoFillDrag);
  window.addEventListener("pointermove", updatePageBreakDrag);
  window.addEventListener("pointermove", updatePrintAreaDrag);
  window.addEventListener("pointermove", updateSplitPaneDrag);
  window.addEventListener("pointermove", updateFormulaBarResize);
  window.addEventListener("pointermove", updateFormulaReferenceDrag);
  window.addEventListener("pointermove", updateDialogDrag);
  window.addEventListener("pointermove", handleSheetTabDragMove);
  window.addEventListener("pointermove", updateHeaderResize);
  window.addEventListener("pointermove", updateHeaderSelectionDrag);
  window.addEventListener("pointerup", finishCellRangeSelection);
  window.addEventListener("pointerup", finishObjectSelectionModeDrag, true);
  window.addEventListener("pointerup", finishInkDrawing);
  window.addEventListener("pointerup", finishAutoFillDrag);
  window.addEventListener("pointerup", finishPageBreakDrag);
  window.addEventListener("pointerup", finishPrintAreaDrag);
  window.addEventListener("pointerup", finishSplitPaneDrag);
  window.addEventListener("pointerup", finishFormulaBarResize);
  window.addEventListener("pointerup", finishFormulaReferenceDrag);
  window.addEventListener("pointerup", finishDialogDrag);
  window.addEventListener("pointercancel", finishDialogDrag);
  window.addEventListener("pointerup", finishSheetTabDrag);
  window.addEventListener("pointercancel", finishSheetTabDrag);
  window.addEventListener("pointerup", finishHeaderResize);
  window.addEventListener("pointerup", stopEditAutoScroll);
  window.addEventListener("pointerup", finishShapeDrawing);
  window.addEventListener("pointerup", finishSelectedImageTransform);
  window.addEventListener("pointerup", finishPictureCropDrag);
  window.addEventListener("pointerup", finishShapeSelectionPaneDrag);
  window.addEventListener("blur", resetWorkbookModifierState);
  window.addEventListener("blur", cancelObjectSelectionModeDrag);
  window.addEventListener("blur", stopEditAutoScroll);
  window.addEventListener("beforeunload", handleWorkbookBeforeUnload);
  $(document).on("pointerdown", closeSheetTabContextMenuOnOutsidePointer);
  $(document).on("pointerdown", closeSheetNavigatorMenuOnOutsidePointer);
  $(document).on("pointerdown", commitSheetTabRenameOnOutsidePointer);
  $(document).on("pointerdown", closeCellContextMenuOnOutsidePointer);
  $(document).on("pointerdown", closeFileSideMenuOnOutsidePointer);
  $(document).on("pointerdown", closeRibbonDisplayMenuOnOutsidePointer);
  $(document).on("pointerdown", hideRibbonForAutoHideOnOutsidePointer);
  $(document).on("pointerdown", closeAutoFilterMenuOnOutsidePointer);
  $(document).on("pointerdown", closeHomeCommandMenuOnOutsidePointer);
  $(document).on("click", "[data-sheet-menu-action]", handleSheetTabContextMenuAction);
  $(document).on("click", "[data-sheet-navigator-index]", handleSheetNavigatorMenuChoice);
  $(document).on("click", "[data-sheet-tab-color]", handleSheetTabColorChoice);
  $(document).on("click", "[data-cell-menu-action]", handleCellContextMenuAction);
  $(document).on("click", "[data-ribbon-color-choice]", handleRibbonColorChoice);
  $(document).on("click", "[data-ribbon-color-none]", handleRibbonColorNone);
  $(document).on("click", "[data-ribbon-color-more]", handleRibbonColorMore);
  $(document).on("click", "[data-ribbon-border-apply]", applyRibbonBorder);
  $(document).on("click", "[data-ribbon-border-menu]", openRibbonBorderMenu);
  $(document).on("click", "[data-ribbon-border-choice]", handleRibbonBorderChoice);
  $(document).on("click", "[data-ribbon-underline-apply]", applyRibbonUnderline);
  $(document).on("click", "[data-ribbon-underline-menu]", openRibbonUnderlineMenu);
  $(document).on("click", "[data-ribbon-underline-choice]", handleRibbonUnderlineChoice);
  $(document).on("click", "[data-ribbon-ink-eraser-choice]", handleRibbonInkEraserChoice);
  $(document).on("click", "[data-name-box-menu-item]", handleNameBoxMenuItemClick);
  $(document).on("click", "[data-ribbon-shape-choice]", handleRibbonShapeChoice);
  $(document).on("click", "[data-shape-format-action]", handleShapeFormatMenuAction);
  $(document).on("click", "[data-quick-analysis-tab]", handleQuickAnalysisTabClick);
  $(document).on("click", "[data-quick-analysis-action]", handleQuickAnalysisAction);
  $(document).on("click", "[data-zoom-dialog-action]", handleZoomDialogAction);
  $(document).on("keydown", "#zoomDialog", handleZoomDialogKeydown);
  $(document).on("click", "[data-dialog-chrome-close]", handleDialogChromeCloseClick);
  $(document).on("pointerdown", "[data-dialog-drag-handle]", startDialogDrag);
  $(document).on("click", "[data-help-dialog-action]", handleHelpDialogAction);
  $(document).on("keydown", "#helpDialogOverlay", handleHelpDialogKeydown);
  $(document).on("click", "[data-version-dialog-action]", handleVersionDialogAction);
  $(document).on("keydown", handleVersionDialogKeydown);
  $(document).on("click", "[data-message-box-close]", handleMessageBoxButton);
  $(document).on("keydown", "#messageBoxOverlay", handleMessageBoxKeydown);
  $(document).on("compositionstart", "input, textarea, [contenteditable='true']", handleTextInputImeCompositionStart);
  $(document).on("compositionend", "input, textarea, [contenteditable='true']", handleTextInputImeCompositionEnd);
  $selectionKeyboardSink.on("keydown", handleSelectionKeyboardSinkKeydown);
  $selectionKeyboardSink.on("beforeinput", handleSelectionKeyboardSinkBeforeInput);
  $selectionKeyboardSink.on("input", handleSelectionKeyboardSinkInput);
  $selectionKeyboardSink.on("compositionstart", handleSelectionKeyboardSinkCompositionStart);
  $selectionKeyboardSink.on("compositionupdate", handleSelectionKeyboardSinkCompositionUpdate);
  $selectionKeyboardSink.on("compositionend", handleSelectionKeyboardSinkCompositionEnd);
  $selectionKeyboardSink.on("blur", handleSelectionKeyboardSinkBlur);
  $(document).on("click", "[data-format-dialog-tab]", handleFormatDialogTabClick);
  $(document).on("click", "[data-format-dialog-action]", handleFormatDialogAction);
  $(document).on("click", "[data-page-setup-tab]", handlePageSetupTabClick);
  $(document).on("click", "[data-cell-structure-action]", handleCellStructureDialogAction);
  $(document).on("click", "[data-header-size-action]", handleHeaderSizeDialogAction);
  $(document).on("click", "[data-data-dialog-action]", handleDataDialogAction);
  $(document).on("keydown", "#dataDialog [data-data-dialog]", handleDataDialogKeydown);
  $(document).on("input", "[data-insert-icon-search]", handleInsertIconSearch);
  $(document).on("click", "[data-object-insert-tab]", handleObjectInsertTabClick);
  $(document).on("change input", "[data-object-insert-sync]", handleObjectInsertDialogSync);
  $(document).on("change", "[data-save-as-type]", handleSaveAsTypeChange);
  $(document).on("change", "[data-save-as-path]", handleSaveAsPathChange);
  $(document).on("click", "[data-save-as-location]", handleSaveAsLocationClick);
  $(document).on("click", "[data-save-as-file-row]", handleSaveAsFileRowClick);
  $(document).on("keydown", handleGlobalDataDialogAccessKey);
  $(document).on("change", "[data-data-field='customViewIndex']", handleCustomViewListChange);
  $(document).on("keydown", "[data-data-field='customViewIndex']", handleCustomViewListKeydown);
  $(document).on("dblclick", "[data-data-field='customViewIndex']", handleCustomViewListDoubleClick);
  $(document).on("dblclick", "[data-data-field='unhideSheetIndex']", handleSheetUnhideListConfirm);
  $(document).on("keydown", "[data-data-field='unhideSheetIndex']", handleSheetUnhideListConfirm);
  $(document).on("dblclick", "[data-sheet-insert-template]", handleSheetInsertTemplateDoubleClick);
  $(document).on("dblclick", "[data-data-field='sheetMoveBeforeIndex']", handleSheetMoveCopyListConfirm);
  $(document).on("keydown", "[data-data-field='sheetMoveBeforeIndex']", handleSheetMoveCopyListConfirm);
  $(document).on("dblclick", "[data-data-field='windowId'], [data-data-field='sideBySideWindowId'], [data-data-field='switchWindowId']", handleWorkbookWindowListConfirm);
  $(document).on("keydown", "[data-data-field='windowId'], [data-data-field='sideBySideWindowId'], [data-data-field='switchWindowId']", handleWorkbookWindowListConfirm);
  $(document).on("input change", "[data-macro-dialog-sync]", handleMacroDialogSync);
  $(document).on("dblclick", "[data-data-field='macroIndex']", handleMacroListDoubleClick);
  $(document).on("change", "[data-data-field='sheetCodeProcedure']", handleSheetCodeProcedureChange);
  $(document).on("click", "[data-data-menu-action]", handleDataCommandMenuAction);
  $(document).on("click", "#homeCommandMenu button, #dataCommandMenu button, #dataValidationMenu button, #whatIfMenu button, #viewCommandMenu button, #ribbonUnderlineMenu button, #ribbonBorderMenu button, #ribbonInkEraserMenu button, #ribbonColorPalette button, #ribbonShapeMenu button, #shapeFormatMenu button, #sheetTabContextMenu button, #sheetTabNavigatorMenu button, #cellContextMenu button, #statusBarContextMenu button, #queryPaneContextMenu button, #quickAnalysisPopup button, #fileSideMenu button", scheduleWorksheetFocusAfterCommandPopup);
  $(document).on("keydown", "#homeCommandMenu, #dataCommandMenu, #dataValidationMenu, #whatIfMenu, #viewCommandMenu, #ribbonUnderlineMenu, #ribbonBorderMenu, #ribbonInkEraserMenu, #ribbonColorPalette, #ribbonShapeMenu, #shapeFormatMenu, #sheetTabContextMenu, #sheetTabNavigatorMenu, #cellContextMenu, #statusBarContextMenu, #queryPaneContextMenu, #quickAnalysisPopup, #nameBoxMenu", handleRibbonCommandMenuKeydown);
  $(document).on("keydown", handleGlobalCommandMenuAccessKey);
  $(document).on("change", "[data-existing-connection-list]", handleExistingConnectionDialogSync);
  $(document).on("change", "[data-workbook-link-list]", handleWorkbookLinksDialogSync);
  $(document).on("change", "[data-relationship-list]", handleRelationshipsDialogSync);
  $(document).on("change", "[data-relationship-edit-field]", handleRelationshipEditDialogSync);
  $(document).on("change", "[data-relationship-auto-index]", handleRelationshipAutoDetectSync);
  $(document).on("change", "[data-data-model-table-list]", handleDataModelDialogSync);
  $(document).on("click", "[data-validation-menu-action]", handleDataValidationMenuAction);
  $(document).on("click", "[data-validation-dropdown-choice]", handleValidationDropdownChoice);
  $(document).on("keydown", "[data-validation-dropdown-choice]", handleValidationDropdownKeydown);
  $(document).on("input change", "[data-validation-sync]", handleDataValidationDialogSync);
  $(document).on("input change", "[data-advanced-filter-sync]", handleAdvancedFilterDialogSync);
  $(document).on("input change", "[data-consolidate-sync]", handleConsolidateDialogSync);
  $(document).on("input change", "[data-subtotal-sync], [data-data-subtotal-col]", handleSubtotalDialogSync);
  $(document).on("input change", "[data-auto-filter-sync]", handleAutoFilterDialogSync);
  $(document).on("input change", "[data-forecast-sync]", handleForecastDialogSync);
  $(document).on("change", "[data-data-analysis-picker]", handleDataAnalysisPickerSync);
  $(document).on("dblclick", "[data-data-analysis-picker]", handleDataAnalysisPickerDoubleClick);
  $(document).on("input change", "[data-data-analysis-sync]", handleDataAnalysisDialogSync);
  $(document).on("input change", "[data-connection-source-sync]", handleConnectionOnlySourceDialogSync);
  $(document).on("input change", "[data-goal-seek-sync]", handleGoalSeekDialogSync);
  $(document).on("input change", "[data-data-table-sync]", handleDataTableDialogSync);
  $(document).on("input change", "[data-scenario-edit-sync]", handleScenarioEditDialogSync);
  $(document).on("input change", "[data-scenario-summary-sync]", handleScenarioSummaryDialogSync);
  $(document).on("change", "[data-outline-sync], [data-outline-settings-sync]", handleOutlineDialogSync);
  $(document).on("change", "[data-remove-duplicates-sync], [data-data-duplicate-col]", handleRemoveDuplicatesDialogSync);
  $(document).on("input change", "[data-text-columns-sync]", handleTextColumnsPreviewChange);
  $(document).on("input change", "[data-text-csv-sync]", handleTextCsvPreviewChange);
  $(document).on("change", "[data-query-properties-sync]", handleQueryPropertiesDialogSync);
  $(document).on("change", "[data-import-sync]", handleImportDataDialogSync);
  $(document).on("input change", "[data-import-auth-sync]", handleImportAuthDialogSync);
  $(document).on("change", "[data-sort-dialog-sync]", handleSortDialogControlChange);
  $(document).on("click", "[data-what-if-menu-action]", handleWhatIfMenuAction);
  $(document).on("change", "[data-data-field='scenarioList']", handleScenarioListChange);
  $(document).on("click", "[data-query-pane-action]", handleQueryPaneAction);
  $(document).on("contextmenu", ".query-pane-item", openQueryPaneContextMenuFromEvent);
  $(document).on("click", "[data-query-pane-menu-action]", handleQueryPaneContextMenuAction);
  $(document).on("change", "[data-query-append-key]", handleAppendQuerySelectionChange);
  $(document).on("change", "[data-data-field='mergeRightQuery']", handleMergeRightQueryChange);
  $(document).on("input change", "[data-query-merge-preview-sync]", handleMergeQueryPreviewChange);
  $(document).on("click", "[data-auto-filter-action]", handleAutoFilterMenuAction);
  $(document).on("change", "[data-auto-filter-select-all]", handleAutoFilterSelectAllChange);
  $(document).on("change", "[data-auto-filter-value], [data-auto-filter-add-current]", handleAutoFilterValueChange);
  $(document).on("input", "[data-auto-filter-search]", handleAutoFilterSearchInput);
  $(document).on("keydown", "#autoFilterMenu", handleAutoFilterMenuKeydown);
  $(document).on("click", "[data-home-autosum-action]", handleHomeAutoSumMenuAction);
  $(document).on("click", "[data-home-fill-action]", handleHomeFillMenuAction);
  $(document).on("click", "[data-home-series-dialog-action]", handleHomeSeriesDialogAction);
  $(document).on("keydown", "#homeSeriesDialogOverlay", handleHomeSeriesDialogKeydown);
  $(document).on("change input", "[data-home-series-field]", handleHomeSeriesDialogSync);
  $(document).on("click", "[data-home-worksheet-fill-action]", handleHomeWorksheetFillDialogAction);
  $(document).on("keydown", "#homeWorksheetFillDialogOverlay", handleHomeWorksheetFillDialogKeydown);
  $(document).on("click", "[data-home-clear-action]", handleHomeClearMenuAction);
  $(document).on("click", "[data-home-sort-filter-action]", handleHomeSortFilterMenuAction);
  $(document).on("click", "[data-view-menu-action]", handleViewCommandMenuAction);
  $(document).on("click", "[data-find-select-action]", handleFindSelectMenuAction);
  $(document).on("click", "[data-formula-calculation-action]", handleFormulaCalculationMenuAction);
  $(document).on("click", "[data-formula-function-menu-action]", handleFormulaFunctionMenuAction);
  $(document).on("click", "[data-formula-name-menu-action]", handleFormulaNameMenuAction);
  $(document).on("click", "[data-formula-audit-menu-action]", handleFormulaAuditMenuAction);
  $(document).on("click", "[data-formula-watch-action]", handleFormulaWatchWindowAction);
  $(document).on("dblclick", "[data-formula-watch-row]", handleFormulaWatchRowDoubleClick);
  $(document).on("click", "[data-layout-menu-action]", handlePageLayoutMenuAction);
  $(document).on("change", "[data-layout-field]", handlePageLayoutFieldChange);
  $(document).on("change", "[data-layout-sheet-option]", handlePageLayoutSheetOptionChange);
  $(document).on("click", "[data-find-dialog-action]", handleFindDialogAction);
  $(document).on("input change", "#findDialog [data-find-field]", handleFindDialogFieldChange);
  $(document).on("change", "[data-formula-function-category]", handleInsertFunctionCategoryChange);
  $(document).on("input", "[data-formula-function-search]", handleInsertFunctionSearchInput);
  $(document).on("keydown", "[data-formula-function-search]", handleInsertFunctionSearchKeydown);
  $(document).on("input change", "[data-formula-argument-input]", handleFunctionArgumentInput);
  $(document).on("focusin", "[data-formula-argument-input]", handleFunctionArgumentInput);
  $(document).on("keydown", "[data-formula-argument-input]", handleFunctionArgumentKeydown);
  $(document).on("change", "[data-formula-name-list]", handleNameManagerListChange);
  $(document).on("change", "[data-formula-name-filter]", handleNameManagerFilterChange);
  $(document).on("dblclick", "[data-formula-name-row]", handleNameManagerRowDoubleClick);
  $(document).on("change", "[data-formula-paste-name-list]", (event) => {
    syncPasteNameDialog(event.currentTarget.closest("[data-data-dialog='paste-name']"));
  });
  $(document).on("dblclick", "[data-formula-paste-name-list]", (event) => {
    applyPasteNameDialog(event.currentTarget.closest("[data-data-dialog='paste-name']"));
  });
  $(document).on("keydown", "#findDialog", handleFindDialogKeydown);
  $(document).on("click", "[data-go-to-dialog-action]", handleGoToDialogAction);
  $(document).on("keydown", "#goToDialog", handleGoToDialogKeydown);
  $(document).on("click", "[data-special-dialog-action]", handleSpecialDialogAction);
  $(document).on("keydown", "#specialDialog", handleSpecialDialogKeydown);
  $(document).on("click", "[data-name-dialog-action]", handleNameDialogAction);
  $(document).on("click", "[data-link-kind]", handleLinkKindClick);
  $(document).on("click", "[data-link-document-sheet-index]", handleLinkDocumentSheetClick);
  $(document).on("click", "[data-link-document-name-target]", handleLinkDocumentNameClick);
  $(document).on("click", "[data-link-dialog-action]", handleLinkDialogAction);
  $(document).on("click", "[data-object-link-kind]", handleObjectLinkKindClick);
  $(document).on("click", "[data-object-link-document-sheet-index]", handleObjectLinkDocumentSheetClick);
  $(document).on("click", "[data-object-link-document-name-target]", handleObjectLinkDocumentNameClick);
  $(document).on("click", "[data-object-link-dialog-action]", handleObjectLinkDialogAction);
  $(document).on("click", "[data-translation-pane-action]", handleTranslationPaneAction);
  $(document).on("change", "[data-translation-pane-field]", handleTranslationPaneFieldChange);
  $(document).on("click", "[data-shape-format-pane-action]", handleShapeFormatPaneAction);
  $(document).on("change", "[data-shape-format-pane-field]", handleShapeFormatPaneFieldChange);
  $(document).on("input", '[data-shape-format-pane-field][type="color"]', handleShapeFormatPaneFieldChange);
  $(document).on("focusin", '[data-shape-format-pane-field][type="color"]', (event) => {
    $(event.currentTarget).removeData("appliedColorInput");
  });
  $(document).on("click", "[data-shape-selection-pane-action]", handleShapeSelectionPaneAction);
  $(document).on("click", "[data-shape-selection-pane-object]", handleShapeSelectionPaneObjectClick);
  $(document).on("dblclick", "[data-shape-selection-pane-object]", handleShapeSelectionPaneObjectDoubleClick);
  $(document).on("keydown", "[data-shape-selection-pane-object]", handleShapeSelectionPaneObjectKeydown);
  $(document).on("keydown", "[data-shape-selection-pane-rename]", handleShapeSelectionPaneRenameKeydown);
  $(document).on("blur", "[data-shape-selection-pane-rename]", handleShapeSelectionPaneRenameBlur);
  $(document).on("pointerdown", "[data-shape-selection-pane-drag]", startShapeSelectionPaneDrag);
  $(document).on("click", "[data-symbol-choice]", handleSymbolChoice);
  $(document).on("click", "[data-symbol-dialog-action]", handleSymbolDialogAction);
  $(document).on("click", "[data-annotation-action]", handleAnnotationEditorAction);
  $(document).on("click", "[data-comment-thread-action]", handleCommentThreadAction);
  $(document).on("click", "[data-status-bar-item]", handleStatusBarMenuItem);
  $(document).on("change", 'input[name="zoomDialogOption"]', handleZoomDialogOptionChange);
  $(document).on("input change", "#formatCellsDialog [data-format-field]", handleFormatDialogFieldChange);
  $(document).on("keydown", "#cellAnnotationEditor textarea", handleAnnotationEditorKeydown);
  $(document).on("keydown", "#cellCommentThreadPopup textarea", handleCommentThreadKeydown);
  $(document).on("keydown", "#shapeTextEditor", handleShapeTextEditorKeydown);
  $(document).on("keydown", "#headerSizeDialog input", handleHeaderSizeDialogKeydown);
  $(document).on("mouseenter", "#cellCommentThreadPopup", cancelCloseCommentThreadPopup);
  $(document).on("mouseleave", "#cellCommentThreadPopup", scheduleCloseCommentThreadPopup);
  $(document).on("mouseenter", "#cellNotePopup", cancelCloseNotePopup);
  $(document).on("mouseleave", "#cellNotePopup:not(.is-persistent)", scheduleCloseNotePopup);
  $(document).on("focusin", "[data-zoom-custom-input]", selectCustomZoomOption);
  $(document).on("keydown", closeSheetTabContextMenuOnEscape);
  $(document).on("keydown", closeRibbonColorPaletteOnEscape);
  $(document).on("keydown", closeRibbonBorderMenuOnEscape);
  $(document).on("keydown", closeRibbonUnderlineMenuOnEscape);
  $(document).on("keydown", closeRibbonShapeMenuOnEscape);
  $(document).on("keydown", closeRibbonInkEraserMenuOnEscape);
  $(document).on("keydown", closeShapeFormatMenuOnEscape);
  $(document).on("keydown", closeAutoFillOptionsMenuOnEscape);
  $(document).on("keydown", closeQuickAnalysisPopupOnEscape);
  $(document).on("keydown", closeDataCommandMenuOnEscape);
  $(document).on("keydown", closeDataValidationMenuOnEscape);
  $(document).on("keydown", closeValidationDropdownOnEscape);
  $(document).on("keydown", closeAutoFilterMenuOnEscape);
  $(document).on("keydown", closeFileSideMenuOnEscape);
  $(document).on("keydown", closeRibbonDisplayMenuOnEscape);
  $(document).on("keydown", closeHomeCommandMenuOnEscape);
  $(document).on("keydown", closeViewCommandMenuOnEscape);
  $(document).on("keydown", closeFormulaAutocompleteOnEscape);
  $(document).on("keydown", closeFindDialogOnEscape);
  $(document).on("keydown", closeWhatIfMenuOnEscape);
  $(document).on("keydown", closeQueryPaneOnEscape);
  $(document).on("keydown", closeZoomDialogOnEscape);
  $(document).on("keydown", closeMessageBoxOnEscape);
  $(document).on("keydown", closeFormatDialogOnEscape);
  $(document).on("keydown", closeCellStructureDialogOnEscape);
  $(document).on("keydown", closeHeaderSizeDialogOnEscape);
  $(document).on("keydown", closeDataDialogOnEscape);
  $(document).on("keydown", closeNameDialogOnEscape);
  $(document).on("keydown", closeLinkDialogOnEscape);
  $(document).on("keydown", closeObjectLinkDialogOnEscape);
  $(document).on("keydown", closeTranslationPaneOnEscape);
  $(document).on("keydown", closeShapeFormatPaneOnEscape);
  $(document).on("keydown", closeShapeSelectionPaneOnEscape);
  $(document).on("keydown", closeSymbolDialogOnEscape);
  $(document).on("keydown", closeAnnotationEditorOnEscape);
  $(document).on("keydown", closeStatusBarContextMenuOnEscape);
  $(document).on("keydown", handleFullScreenEscape);
  $(document).on("keydown", handleWorkbookKeyboardShortcut);
  $(document).on("pointerdown", closeRibbonColorPaletteOnOutsidePointer);
  $(document).on("pointerdown", closeRibbonBorderMenuOnOutsidePointer);
  $(document).on("pointerdown", closeRibbonUnderlineMenuOnOutsidePointer);
  $(document).on("pointerdown", closeRibbonShapeMenuOnOutsidePointer);
  $(document).on("pointerdown", closeRibbonInkEraserMenuOnOutsidePointer);
  $(document).on("pointerdown", closeShapeFormatMenuOnOutsidePointer);
  $(document).on("pointerdown", closeAutoFillOptionsMenuOnOutsidePointer);
  $(document).on("focusin", clearAutoFillOptionsOnExternalFocus);
  $(document).on("pointerdown", closeQuickAnalysisPopupOnOutsidePointer);
  $(document).on("pointerdown", closeDataCommandMenuOnOutsidePointer);
  $(document).on("pointerdown", closeViewCommandMenuOnOutsidePointer);
  $(document).on("pointerdown", closeDataValidationMenuOnOutsidePointer);
  $(document).on("pointerdown", closeValidationDropdownOnOutsidePointer);
  $(document).on("pointerdown", closeWhatIfMenuOnOutsidePointer);
  $(document).on("pointerdown", closeZoomDialogOnOutsidePointer);
  $(document).on("pointerdown", closeNameBoxMenuOnOutsidePointer);
  $(document).on("pointerdown", closeFormulaAutocompleteOnOutsidePointer);
  $(document).on("pointerdown", "[data-formula-autocomplete-item]", (event) => {
    event.preventDefault();
  });
  $(document).on("click", "[data-formula-autocomplete-item]", handleFormulaAutocompleteItemClick);
  $(document).on("pointerdown", closeQueryPaneContextMenuOnOutsidePointer);
  $(document).on("pointerdown", closeStatusBarContextMenuOnOutsidePointer);
  $(document).on("pointerdown", commitAnnotationEditorOnOutsidePointer);
  $selectionStatus.on("contextmenu", openStatusBarContextMenu);
  $(document).on("copy", handleWorkbookCopyEvent);
  $(document).on("cut", handleWorkbookCutEvent);
  $(document).on("paste", handleWorkbookPasteEvent);
  $(window).on("resize", closeSheetTabContextMenu);
  $(window).on("resize", closeCellContextMenu);
  $(window).on("resize", () => closeFileSideMenu({ restoreHome: true }));
  $(window).on("resize", closeRibbonDisplayMenu);
  $(window).on("resize", closeRibbonColorPalette);
  $(window).on("resize", closeRibbonBorderMenu);
  $(window).on("resize", closeRibbonUnderlineMenu);
  $(window).on("resize", closeRibbonShapeMenu);
  $(window).on("resize", closeRibbonInkEraserMenu);
  $(window).on("resize", closeShapeFormatMenu);
  $(window).on("resize", () => clearAutoFillOptions());
  $(window).on("resize", closeNameBoxMenu);
  $(window).on("resize", () => closeFormulaAutocomplete({ restoreFocus: false }));
  $(window).on("resize", closeQueryPaneContextMenu);
  $(window).on("resize", closeDataCommandMenu);
  $(window).on("resize", closeDataValidationMenu);
  $(window).on("resize", closeHomeCommandMenu);
  $(window).on("resize", closeViewCommandMenu);
  $(window).on("resize", closeFindDialog);
  $(window).on("resize", closeWhatIfMenu);
  $(window).on("resize", () => closeShapeSelectionPane({ restoreFocus: false }));
  $(window).on("resize", closeZoomDialog);
  $(window).on("resize", closeStatusBarContextMenu);
  $(window).on("resize", closeFormatCellsDialog);
  $(window).on("resize", closeCellStructureDialog);
  $(window).on("resize", closeHeaderSizeDialog);
  $(window).on("resize", closeNameDialog);
  $(window).on("resize", closeLinkDialog);
  $(window).on("resize", () => renderTranslationPane());
  installDialogAccessKeyEnhancer();

  $sheetHost.on("focusin", ".sheet-cell[data-row][data-col]", (event) => {
    if (event.currentTarget.getAttribute("contenteditable") === "true") {
      selectCell(event.currentTarget, { preserveRange: true });
    }
    const $cell = $(event.currentTarget);
    if ($cell.data("beforeEdit") === undefined) {
      const row = Number($cell.data("row"));
      const col = Number($cell.data("col"));
      $cell.data("beforeEdit", getCellRawInput(activeSheet(), row, col));
    }
  });

  $sheetHost.on("input", ".sheet-cell[contenteditable='true']", (event) => {
    syncEditableCellAlignment(event.currentTarget);
    saveCellEditSelection(event.currentTarget);
    updateFormulaReferenceUi();
    updateFormulaAutocompleteForEditor(event.currentTarget);
  });
  $sheetHost.on("beforeinput", ".sheet-cell[contenteditable='true']", (event) => {
    const nativeEvent = event.originalEvent || event;
    const inputType = String(nativeEvent.inputType || "");
    if (inputType.toLowerCase().includes("composition") || nativeEvent.isComposing) {
      recordDirectTypingCompositionData(event.currentTarget, nativeEvent.data || "");
    }
  });
  $sheetHost.on("keyup pointerup click", ".sheet-cell[contenteditable='true']", (event) => {
    saveCellEditSelection(event.currentTarget);
    scheduleFormulaAutocompleteRefreshForEditor(event.currentTarget, { keepSelection: true, event });
  });

  $sheetHost.on("keydown", ".sheet-cell[contenteditable='true']", (event) => {
    const dateTimeShortcut = keyboardDateTimeShortcutKind(event, String(event.key || "").toLowerCase());
    if (dateTimeShortcut) {
      if (isImeCompositionEvent(event, event.currentTarget, "cell")) return;
      event.preventDefault();
      event.stopPropagation();
      insertKeyboardDateTimeIntoEditable(event.currentTarget, dateTimeShortcut);
      return;
    }
    if (isEnterKeyEvent(event) && isAltModifierEvent(event)) {
      event.preventDefault();
      event.stopPropagation();
      insertLineBreakInEditable(event.currentTarget);
      return;
    }
    if (handleFormulaAutocompleteKeydown(event, event.currentTarget)) return;
    if (event.key === "F4" && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey) {
      if (cycleFormulaReferenceInEditable(event.currentTarget)) {
        event.preventDefault();
        event.stopPropagation();
        return;
      }
    }
    if (isEnterKeyEvent(event)) {
      if (event.ctrlKey || event.metaKey) {
        if (isImeCompositionEvent(event, event.currentTarget, "cell")) return;
        event.preventDefault();
        commitEditableCellToSelection(event.currentTarget);
        return;
      }
      if (isImeCompositionEvent(event, event.currentTarget, "cell")) return;
      event.preventDefault();
      event.currentTarget.blur();
    }
    if (event.key === "Escape") {
      event.preventDefault();
      $(event.currentTarget).data("cancelEdit", true);
      event.currentTarget.blur();
    }
  });
  $sheetHost.on("compositionstart", ".sheet-cell[contenteditable='true']", (event) => {
    markDirectTypingComposition(event.currentTarget, true);
    recordDirectTypingCompositionData(event.currentTarget, event.originalEvent?.data || event.data || "");
    markImeComposition(event.currentTarget, "cell", true);
  });
  $sheetHost.on("compositionupdate", ".sheet-cell[contenteditable='true']", (event) => {
    recordDirectTypingCompositionData(event.currentTarget, event.originalEvent?.data || event.data || "");
  });
  $sheetHost.on("compositionend", ".sheet-cell[contenteditable='true']", (event) => {
    markImeComposition(event.currentTarget, "cell", false);
    reconcileDirectTypingImeCommit(event.currentTarget, event.originalEvent?.data || event.data || "");
    updateFormulaAutocompleteForEditor(event.currentTarget);
  });

  $sheetHost.on("blur", ".sheet-cell[contenteditable='true']", (event) => {
    const $cell = $(event.currentTarget);
    clearCellEditDirectTyping(event.currentTarget);
    if ($cell.data("skipBlurCommit")) {
      $cell.removeData("skipBlurCommit");
      return;
    }
    const row = Number($cell.data("row"));
    const col = Number($cell.data("col"));
    const before = String($cell.data("beforeEdit") ?? "");
    const beforeEditTextData = $cell.data("beforeEditText");
    const beforeEditText = beforeEditTextData === undefined ? before : String(beforeEditTextData);
    const beforeRichText = $cell.data("beforeRichText") || null;
    const after = editablePlainText(event.currentTarget);
    const afterRichText = cellEditRichTextForCommit(event.currentTarget, after);
    if (cellEditBlurStartsFormulaSheetNavigation(event, after)) {
      preserveCellFormulaEditForSheetNavigation(event.currentTarget, row, col, before, after);
      return;
    }
    closeFormulaAutocomplete({ restoreFocus: false });
    clearFormulaReferenceState({ update: false });
    event.currentTarget.removeAttribute("contenteditable");
    $cell.removeData("beforeEdit");
    $cell.removeData("beforeEditText");
    $cell.removeData("beforeRichText");
    const canceled = Boolean($cell.data("cancelEdit"));
    $cell.removeData("cancelEdit");
    clearCellEditSelection(event.currentTarget);
    stopEditAutoScroll();
    if (canceled) {
      renderSheet();
      updateSelectionUi();
      return;
    }
    const modelBefore = getCellRawInput(activeSheet(), row, col);
    if (modelBefore !== before || after !== beforeEditText || stableObjectKey(beforeRichText || null) !== stableObjectKey(afterRichText || null)) {
      state.deferredViewportRender = false;
      commitCellValue(row, col, after, { preserveWhitespace: true, richText: afterRichText });
    } else {
      renderSheet();
      updateSelectionUi();
    }
  });

  $formulaInput.on("keydown", (event) => {
    if (handleWorkbookPaneFocusShortcut(event, String(event.key || "").toLowerCase())) return;
    if (isEnterKeyEvent(event) && isAltModifierEvent(event)) {
      event.preventDefault();
      event.stopPropagation();
      insertLineBreakInTextControl(event.currentTarget);
      return;
    }
    if (handleFormulaAutocompleteKeydown(event, event.currentTarget)) return;
    if (event.key === "F4" && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey) {
      if (cycleFormulaReferenceInFormulaBar()) {
        event.preventDefault();
        event.stopPropagation();
        return;
      }
    }
    if (event.key === "Enter") {
      if (isImeCompositionEvent(event, event.currentTarget, "formula")) return;
      event.preventDefault();
      commitFormulaBar({ fillSelection: event.ctrlKey || event.metaKey });
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      cancelFormulaBarEdit();
    }
  });
  $formulaInput.on("compositionstart", (event) => {
    markImeComposition(event.currentTarget, "formula", true);
  });
  $formulaInput.on("compositionend", (event) => {
    markImeComposition(event.currentTarget, "formula", false);
    updateFormulaAutocompleteForEditor(event.currentTarget);
  });
  $formulaInput.on("focus pointerdown", beginFormulaBarEdit);
  $formulaInput.on("pointerdown", startEditAutoScroll);
  $formulaInput.on("input", syncFormulaEditPreview);
  $formulaInput.on("input", updateFormulaReferenceUi);
  $formulaInput.on("input", (event) => updateFormulaAutocompleteForEditor(event.currentTarget));
  $formulaInput.on("keyup pointerup click scroll", (event) => {
    scheduleFormulaAutocompleteRefreshForEditor(event.currentTarget, { keepSelection: true, event });
  });
  $selectedCellLabel.on("focus pointerdown click", beginNameBoxEdit);
  $selectedCellLabel.on("keydown", handleNameBoxKeydown);
  $selectedCellLabel.on("blur", finishNameBoxEditOnBlur);
  $formulaExpandButton.on("click", toggleFormulaBarLines);
  $formulaResizeHandle.on("pointerdown", startFormulaBarResize);
  $exportHtmlButton.on("click", () => void exportStandaloneHtml());

  $toggleFormulasButton.on("click", toggleShowFormulasFromRibbon);

  window.addEventListener("resize", () => {
    updateHorizontalScrollbar();
  });
  document.addEventListener("fullscreenchange", handleFullscreenChange);
  window.addEventListener("pagehide", () => {
    if (state.persistenceTimer && state.model) {
      void flushWorkbookPersistence();
    }
  });

  if (import.meta.env.DEV) {
    window.__webSheetDevLoadBlob = async (blob, name = "workbook.xlsx") => {
      await loadExcelFile(new File([blob], name, { type: blob.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    };
    window.__webSheetPersistenceFlush = () => flushWorkbookPersistence();
    window.__webSheetMemoryStats = () => workbookMemoryStats(state.model);
    window.__webSheetQueryRefreshTimerCount = () => state.queryRefreshTimers?.length || 0;
    window.__webSheetDevExitFullScreen = () => exitFullScreenView();
    window.__webSheetDevCloseFloatingMenus = () => {
      closeHomeCommandMenu();
      closeViewCommandMenu();
      closeDataCommandMenu();
      closeDataValidationMenu();
      closeWhatIfMenu();
      closeRibbonColorPalette();
      closeRibbonBorderMenu();
      closeRibbonUnderlineMenu();
      closeRibbonShapeMenu();
      closeRibbonInkEraserMenu();
      closeShapeFormatMenu();
      closeSheetTabContextMenu();
      closeCellContextMenu();
      closeFileSideMenu({ restoreHome: true });
      closeNameBoxMenu({ restoreFocus: false });
      closeFormulaAutocomplete({ restoreFocus: false });
      closeQueryPaneContextMenu();
      closeQuickAnalysisPopup();
      closeStatusBarContextMenu();
    };
    window.__webSheetDevQueries = () => workbookQueries().map(({ sheet, sheetIndex, query }) => ({
      sheet: sheet.name,
      sheetIndex,
      query: cloneJsonLikeValue(query),
    }));
    window.__webSheetDevDataModel = () => cloneWorkbookDataModel();
    window.__webSheetDevDefinedNames = () => cloneDefinedNames(state.model?.definedNames || []);
    window.__webSheetDevFormulaWatches = () => formulaWatchEntries().map((watch) => ({ ...watch }));
    window.__webSheetDevFormulaAutocompleteDefinitions = () => formulaAutocompleteDefinitions().map((definition) => ({
      kind: definition.kind,
      name: definition.name,
      category: definition.category,
      syntax: definition.syntax,
      engineSupported: Boolean(definition.engineSupported),
    }));
    window.__webSheetDevFormulaAutocompletePreview = (text, caret = null) => {
      const value = String(text ?? "");
      const offset = Number.isFinite(Number(caret)) ? Number(caret) : value.length;
      const queryContext = formulaAutocompleteQueryContext(value, offset, { sheetIndex: state.activeSheetIndex });
      const hintContext = formulaParameterHintContext(value, offset);
      return {
        queryContext,
        hintFunction: hintContext?.definition?.name || "",
        hintArgumentIndex: hintContext?.argIndex ?? null,
        candidates: (formulaAutocompleteShouldShowCandidates(queryContext) ? formulaAutocompleteCandidates(queryContext) : []).map((candidate) => ({
          kind: candidate.kind,
          name: candidate.name,
          category: candidate.category,
          syntax: candidate.syntax,
        })),
      };
    };
    window.__webSheetDevAddDefinedName = (definition = {}) => {
      if (!state.model || !definition.name || !definition.refersTo) return false;
      state.model.definedNames = upsertDefinedName(state.model.definedNames || [], {
        name: String(definition.name),
        refersTo: String(definition.refersTo).startsWith("=") ? String(definition.refersTo) : `=${definition.refersTo}`,
        scope: definition.scope || "workbook",
        sheetIndex: Number.isFinite(Number(definition.sheetIndex)) ? Number(definition.sheetIndex) : null,
        comment: definition.comment || "",
        hidden: Boolean(definition.hidden),
      });
      state.hf = buildFormulaEngine(state.model);
      return true;
    };
    window.__webSheetDevSetDefinedNames = (definitions = []) => {
      if (!state.model) return false;
      state.model.definedNames = (Array.isArray(definitions) ? definitions : [])
        .map((definition) => ({
          name: String(definition?.name || "").trim(),
          refersTo: String(definition?.refersTo || "").trim().startsWith("=")
            ? String(definition.refersTo).trim()
            : `=${String(definition?.refersTo || "").trim()}`,
          scope: definition?.scope === "sheet" ? "sheet" : "workbook",
          sheetIndex: definition?.scope === "sheet" && Number.isInteger(Number(definition.sheetIndex)) ? Number(definition.sheetIndex) : null,
          comment: String(definition?.comment || ""),
          hidden: Boolean(definition?.hidden),
        }))
        .filter((definition) => definition.name && definition.refersTo !== "=");
      state.hf = buildFormulaEngine(state.model);
      return true;
    };
    window.__webSheetDevAddTable = (sheetName, table = {}) => {
      const sheet = state.model?.sheets?.find((candidate) => candidate.name === sheetName);
      if (!sheet || !table.name || !table.range) return false;
      sheet.tables = Array.isArray(sheet.tables) ? sheet.tables : [];
      const name = String(table.name);
      const next = {
        id: table.id || uniqueWorksheetObjectId("dev-table"),
        name,
        displayName: table.displayName || name,
        range: { ...table.range },
        showHeaderRow: table.showHeaderRow !== false,
        showTotalsRow: Boolean(table.showTotalsRow),
        showFilterButton: table.showFilterButton !== false,
        style: table.style || "TableStyleMedium2",
        styleInfo: table.styleInfo || parseTableStyleInfoXml(null),
        hasHeaders: table.hasHeaders !== false,
        autoFilterState: cloneJsonLikeValue(table.autoFilterState || null),
        sortState: cloneJsonLikeValue(table.sortState || null),
        columns: (table.columns || []).map((column, index) => ({
          id: String(column?.id || index + 1),
          name: String(column?.name || `Column${index + 1}`),
          calculatedColumnFormula: column?.calculatedColumnFormula || "",
          totalsRowFunction: column?.totalsRowFunction || "",
          totalsRowFormula: column?.totalsRowFormula || "",
          totalsRowLabel: column?.totalsRowLabel || "",
        })),
      };
      const existingIndex = sheet.tables.findIndex((item) => (item.displayName || item.name) === name);
      if (existingIndex >= 0) sheet.tables[existingIndex] = next;
      else sheet.tables.push(next);
      renderWorkbook();
      return true;
    };
    window.__webSheetDevSetMacros = (macros = []) => {
      if (!state.model) return false;
      state.model.macros = (Array.isArray(macros) ? macros : [])
        .map((macro, index) => ({
          id: macro?.id || `dev-macro-${index + 1}`,
          name: String(macro?.name || `Macro${index + 1}`),
          code: String(macro?.code || ""),
          scope: macro?.scope || "workbook",
          shortcut: String(macro?.shortcut || "").slice(0, 1),
          description: String(macro?.description || ""),
          createdAt: macro?.createdAt || new Date().toISOString(),
          updatedAt: macro?.updatedAt || new Date().toISOString(),
          lastRunAt: macro?.lastRunAt || null,
        }))
        .filter((macro) => macro.name);
      state.model.macroRecording = null;
      scheduleWorkbookPersistence();
      renderWorkbook();
      return true;
    };
    window.__webSheetDevRunMacro = (macroRef = 0, options = {}) => {
      if (!state.model) return null;
      const macros = Array.isArray(state.model.macros) ? state.model.macros : [];
      const macroIndex = typeof macroRef === "number"
        ? macroRef
        : macros.findIndex((macro) => String(macro.name || "") === String(macroRef || ""));
      const macro = macros[macroIndex];
      if (!macro) return null;
      const result = executeWorkbookMacroStatements(
        macro,
        macroIndex,
        parseVbaMacroStatements(macro.code || ""),
        {
          historyLabel: `マクロ実行 ${macro.name || `Macro${macroIndex + 1}`}`,
          statusLabel: macro.name || "マクロ",
          updateLastRun: options.updateLastRun !== false,
          showMessages: options.showMessages === true,
        },
      );
      return {
        writeCount: result?.writeCount || 0,
        messages: [...(result?.messages || [])],
        unsupported: [...(result?.unsupported || [])],
        touchedSheetIndexes: [...(result?.touchedSheetIndexes || [])],
      };
    };
    window.__webSheetDevWorkbookState = () => {
      if (state.model) saveActiveWindowViewState();
      return {
        activeSheetIndex: state.activeSheetIndex,
        activeSheetName: activeSheet()?.name || "",
        sourceName: state.model?.sourceName || "",
        saveType: state.model?.saveType || "",
        currentSaveType: state.currentSaveType || "",
        isStandalone: state.isStandalone === true,
        showFormulas: state.showFormulas,
        formulaBarVisible: !$formulaBar.hasClass("hidden"),
        calculationMode: workbookCalculationMode(),
        workbookProperties: cloneJsonLikeValue(state.model?.workbookProperties || null),
        calcProperties: cloneJsonLikeValue(state.model?.calcProperties || null),
        workbookView: cloneJsonLikeValue(state.model?.workbookView || null),
        windowViews: cloneJsonLikeValue(state.model?.windowViews || []),
        mainWindowView: cloneJsonLikeValue(state.model?.mainWindowView || null),
        windowArrangement: cloneJsonLikeValue(state.model?.windowArrangement || null),
        selectedCell: state.selected ? { ...state.selected } : null,
        selectionRange: state.selectionRange ? { ...state.selectionRange } : null,
        selectionRanges: (state.selectionRanges || []).map((range) => ({ ...range })),
        selectedSheetIndexes: [...(state.selectedSheetIndexes || [])],
        customViews: cloneJsonLikeValue(state.model?.customViews || []),
        activeWindowId: state.model?.activeWindowId || "main",
        mainWindowHidden: state.model?.mainWindowHidden === true,
        viewSideBySide: cloneJsonLikeValue(state.model?.viewSideBySide || null),
        synchronousScrolling: state.model?.synchronousScrolling === true,
        macros: cloneJsonLikeValue(state.model?.macros || []),
        macroRecording: cloneJsonLikeValue(state.model?.macroRecording || null),
        macroRelativeReferences: state.model?.macroRelativeReferences === true,
        macroSecurity: cloneJsonLikeValue(normalizeMacroSecurity(state.model?.macroSecurity)),
        themeColors: cloneJsonLikeValue(state.model?.themeColors || DEFAULT_THEME_COLORS),
        themeFonts: cloneJsonLikeValue(state.model?.themeFonts || {}),
        themeEffects: cloneJsonLikeValue(state.model?.themeEffects || normalizedPageLayoutThemeEffects()),
        fullScreenView: state.model?.fullScreenView === true,
        internalClipboard: state.internalClipboard
          ? {
              mode: state.internalClipboard.mode,
              text: state.internalClipboard.text || "",
              rowCount: state.internalClipboard.rowCount || 0,
              colCount: state.internalClipboard.colCount || 0,
              sourceRange: state.internalClipboard.sourceRange ? { ...state.internalClipboard.sourceRange } : null,
            }
          : null,
        objectClipboard: state.objectClipboard
          ? {
              type: state.objectClipboard.type || "",
              mode: state.objectClipboard.mode || "",
              count: state.objectClipboard.objects?.length || 0,
              sourceSheetIndex: state.objectClipboard.sourceSheetIndex,
            }
          : null,
      };
    };
    window.__webSheetDevHistoryState = () => ({
      undo: state.undoStack.map((entry) => entry.label || ""),
      redo: state.redoStack.map((entry) => entry.label || ""),
    });
    window.__webSheetDevSheetModel = (sheetName) => {
      const sheet = state.model?.sheets?.find((candidate) => candidate.name === sheetName);
      return sheet ? cloneSheetModel(sheet) : null;
    };
    window.__webSheetDevSetSheetViewOptions = (sheetName, options = {}) => {
      const sheet = state.model?.sheets?.find((candidate) => candidate.name === sheetName);
      if (!sheet || !options || typeof options !== "object") return null;
      if ("viewMode" in options) sheet.viewMode = normalizedSheetViewMode(options.viewMode);
      if ("zoomScale" in options) sheet.zoomScale = clampZoom(options.zoomScale || 100);
      if ("showGridLines" in options) sheet.showGridLines = options.showGridLines !== false;
      if ("showHeadings" in options) sheet.showHeadings = options.showHeadings !== false;
      if ("showRuler" in options) sheet.showRuler = normalizedSheetShowRuler(options.showRuler, sheet.viewMode);
      if ("showObjects" in options) sheet.showObjects = options.showObjects !== false;
      if ("showOutlineSymbols" in options) sheet.showOutlineSymbols = options.showOutlineSymbols !== false;
      if ("showFormulas" in options) sheet.showFormulas = options.showFormulas === true;
      if ("showZeros" in options) sheet.showZeros = options.showZeros !== false;
      if ("showWhiteSpace" in options) sheet.showWhiteSpace = options.showWhiteSpace !== false;
      if ("rightToLeft" in options) sheet.rightToLeft = options.rightToLeft === true;
      if ("freezePanes" in options) sheet.freezePanes = normalizePaneState(options.freezePanes);
      if ("splitPane" in options) sheet.splitPane = normalizePaneState(options.splitPane);
      if (sheet.freezePanes) sheet.splitPane = null;
      if (sheet.splitPane) sheet.freezePanes = null;
      if (activeSheet() === sheet) {
        state.showFormulas = sheet.showFormulas === true;
      }
      renderWorkbook();
      scheduleWorkbookPersistence();
      return standaloneSheetViewDiagnostics(sheet);
    };
    window.__webSheetDevCellModel = (sheetName, row, col) => {
      const sheet = state.model?.sheets?.find((candidate) => candidate.name === sheetName);
      const cell = sheet?.cells?.[cellKey(Number(row), Number(col))];
      return cell ? cloneJsonLikeValue(cell) : null;
    };
    window.__webSheetDevNormalizeDateCell = (cell = {}) => {
      const model = cloneCellModel({
        row: Number(cell.row) || 1,
        col: Number(cell.col) || 1,
        raw: cell.raw ?? null,
        formula: cell.formula ?? null,
        cached: cell.cached ?? null,
        display: cell.display ?? "",
        numFmt: cell.numFmt || "",
        css: cell.css || {},
        borders: cell.borders || {},
      });
      normalizePersistedCellDateValues(model);
      return {
        raw: model.raw,
        cached: model.cached,
        display: model.display,
        rendered: cellDisplayText(model),
      };
    };
    window.__webSheetDevExportXlsxBase64 = async () => {
      if (!state.model) return "";
      const blob = await workbookBlobForSaveType(saveAsFileTypeByValue("xlsx"));
      return bytesToBase64(new Uint8Array(await blob.arrayBuffer()));
    };
    window.__webSheetDevSetActiveSheet = (sheetName) => {
      const index = state.model?.sheets?.findIndex((candidate) => candidate.name === sheetName) ?? -1;
      if (index < 0) return false;
      setActiveSheet(index);
      return true;
    };
    window.__webSheetDevSetSheetSortState = (sheetName, sortState = null) => {
      const sheet = state.model?.sheets?.find((candidate) => candidate.name === sheetName);
      if (!sheet) return false;
      if (sortState && sortState.range) sheet.sortState = cloneJsonLikeValue(sortState);
      else delete sheet.sortState;
      renderWorkbook();
      return true;
    };
    window.__webSheetDevSetPageBreaks = (sheetName, rows = [], cols = []) => {
      const sheet = state.model?.sheets?.find((candidate) => candidate.name === sheetName);
      if (!sheet) return false;
      const layout = ensurePageLayoutSettings(sheet);
      layout.pageBreaks = {
        rows: [...new Set((Array.isArray(rows) ? rows : []).map(Number).filter((row) => row > 1))].sort((a, b) => a - b),
        cols: [...new Set((Array.isArray(cols) ? cols : []).map(Number).filter((col) => col > 1))].sort((a, b) => a - b),
      };
      renderWorkbook();
      return true;
    };
    window.__webSheetDevSetCell = (row, col, value) => {
      commitCellValue(Number(row), Number(col), value, { skipValidation: true });
      return true;
    };
    window.__webSheetDevApplyCellStructure = (kind, option, top, left, bottom = top, right = left) => {
      const range = normalizeRange(
        { sheetIndex: state.activeSheetIndex, row: Number(top), col: Number(left) },
        { sheetIndex: state.activeSheetIndex, row: Number(bottom), col: Number(right) },
      );
      return applyCellStructureOperation(kind, option, range);
    };
    window.__webSheetDevApplyHeaderStructure = (kind, option, index, endIndex = index) => {
      if (!state.model) return false;
      const start = Number(index);
      const end = Number(endIndex);
      if (!Number.isFinite(start) || !Number.isFinite(end)) return false;
      const range = option === "entire-column"
        ? columnSelectionRange(start, end)
        : rowSelectionRange(start, end);
      return applyCellStructureOperation(kind, option, range);
    };
    window.__webSheetDevRefreshCharts = () => refreshWorkbookChartsFromSources();
    window.__webSheetDevShiftFormulaReferences = (formula, rowOffset = 0, colOffset = 0) =>
      shiftFormulaReferences(formula, Number(rowOffset) || 0, Number(colOffset) || 0);
    window.__webSheetDevSelectRange = (top, left, bottom, right) => {
      const range = normalizeRange(
        { sheetIndex: state.activeSheetIndex, row: Number(top), col: Number(left) },
        { sheetIndex: state.activeSheetIndex, row: Number(bottom), col: Number(right) },
      );
      selectSingleRange(range, `${rangeToLabel(range)} を選択しました。`);
      return rangeToLabel(range);
    };
    window.__webSheetDevCellClientRect = (row, col, options = {}) => {
      const sheet = activeSheet();
      const host = $sheetHost[0];
      if (!sheet || !host) return null;
      const targetRow = Math.max(1, Number(row) || 1);
      const targetCol = Math.max(1, Number(col) || 1);
      ensureSheetSize(sheet, targetRow, targetCol);
      const range = normalizeRange(
        { sheetIndex: state.activeSheetIndex, row: targetRow, col: targetCol },
        { sheetIndex: state.activeSheetIndex, row: targetRow, col: targetCol },
      );
      const wasVisible = rangeIntersects(range, visibleGridRange(sheet));
      const didScroll = options.scrollIntoView === false ? false : scrollRangeIntoView(range);
      if (didScroll || !wasVisible) renderSheet();
      const sheetRect = rangePixelRect(sheet, range);
      const hostRect = host.getBoundingClientRect();
      const clientRect = {
        x: hostRect.left + sheetRect.left - host.scrollLeft,
        y: hostRect.top + sheetRect.top - host.scrollTop,
        width: sheetRect.width,
        height: sheetRect.height,
      };
      return {
        ...clientRect,
        left: clientRect.x,
        top: clientRect.y,
        right: clientRect.x + clientRect.width,
        bottom: clientRect.y + clientRect.height,
        centerX: clientRect.x + clientRect.width / 2,
        centerY: clientRect.y + clientRect.height / 2,
        visible: clientRect.right >= hostRect.left && clientRect.left <= hostRect.right &&
          clientRect.bottom >= hostRect.top && clientRect.top <= hostRect.bottom,
        sheetRect,
        scrollLeft: host.scrollLeft,
        scrollTop: host.scrollTop,
      };
    };
    window.__webSheetDevSetHeaderHidden = (kind, index, hidden) => {
      const sheet = activeSheet();
      const targetIndex = Number(index);
      if (!sheet || !targetIndex) return false;
      if (kind === "column") {
        ensureSheetSize(sheet, sheet.rowCount, targetIndex);
        sheet.columns[targetIndex - 1].hidden = Boolean(hidden);
      } else {
        ensureSheetSize(sheet, targetIndex, sheet.colCount);
        sheet.rows[targetIndex - 1].hidden = Boolean(hidden);
        if (hidden) sheet.rows[targetIndex - 1].height = 0;
        else sheet.rows[targetIndex - 1].height = sheet.rows[targetIndex - 1].height || sheet.defaultRowHeight || rowHeightToPixels(15);
      }
      renderSheet();
      updateSelectionUi();
      return true;
    };
    window.__webSheetDevSetHeaderFooter = (sheetName, value = {}) => {
      const sheet = state.model?.sheets?.find((candidate) => candidate.name === sheetName) || activeSheet();
      if (!sheet) return false;
      sheet.headerFooter = normalizedHeaderFooterSettings(value, sheet.pageLayout);
      sheet.pageLayout = normalizedPageLayoutSettings(sheet.pageLayout, sheet);
      sheet.pageLayout.headerFooter = {
        header: headerFooterPlainText(sheet.headerFooter.oddHeader),
        footer: headerFooterPlainText(sheet.headerFooter.oddFooter),
      };
      renderSheet();
      return true;
    };
    window.__webSheetDevApplyHeaderSize = (kind, top, left, bottom, right, size) => {
      const range = normalizeRange(
        { sheetIndex: state.activeSheetIndex, row: Number(top), col: Number(left) },
        { sheetIndex: state.activeSheetIndex, row: Number(bottom), col: Number(right) },
      );
      applyHeaderSize(kind, [range], Number(size));
      return true;
    };
    window.__webSheetDevSetHeaderHiddenWithHistory = (kind, hidden, top, left, bottom, right) => {
      const range = normalizeRange(
        { sheetIndex: state.activeSheetIndex, row: Number(top), col: Number(left) },
        { sheetIndex: state.activeSheetIndex, row: Number(bottom), col: Number(right) },
      );
      selectSingleRange(range, `${rangeToLabel(range)} を選択しました。`);
      setHeaderHidden(kind, Boolean(hidden), range.top, range.left);
      return true;
    };
    window.__webSheetDevZoomSelectionToFit = () => {
      const sheet = activeSheet();
      const before = Number(sheet?.zoomScale || 100);
      const ranges = activeSelectionRanges();
      const fitRange = zoomSelectionFitRange(sheet, ranges);
      const result = zoomSelectionToFit({ defer: false });
      return {
        result,
        before,
        after: Number(sheet?.zoomScale || 100),
        sheetName: sheet?.name || "",
        range: rangesToLabel(ranges),
        fitRange: fitRange ? rangeToLabel(fitRange) : "",
        fitRangeObject: fitRange ? { ...fitRange } : null,
        status: document.querySelector("#statusText")?.textContent || "",
      };
    };
    window.__webSheetDevSelectRanges = (ranges = []) => {
      const normalized = (Array.isArray(ranges) ? ranges : [])
        .map((item) => normalizeRange(
          { sheetIndex: state.activeSheetIndex, row: Number(item.top), col: Number(item.left) },
          { sheetIndex: state.activeSheetIndex, row: Number(item.bottom ?? item.top), col: Number(item.right ?? item.left) },
        ))
        .filter(Boolean);
      selectRangesForHomeCommand(normalized, `${rangesToLabel(normalized)} を選択しました。`);
      return rangesToLabel(normalized);
    };
    window.__webSheetDevSelectFullRows = (top, bottom = top) => {
      const sheet = activeSheet();
      const range = { sheetIndex: state.activeSheetIndex, top: Number(top), bottom: Number(bottom), left: 1, right: sheet.colCount, fullRow: true };
      selectRangesForHomeCommand([range], `${rangeToLabel(range)} を選択しました。`);
      return rangeToLabel(range);
    };
    window.__webSheetDevSelectFullColumns = (left, right = left) => {
      const sheet = activeSheet();
      const range = { sheetIndex: state.activeSheetIndex, top: 1, bottom: sheet.rowCount, left: Number(left), right: Number(right), fullColumn: true };
      selectRangesForHomeCommand([range], `${rangeToLabel(range)} を選択しました。`);
      return rangeToLabel(range);
    };
    window.__webSheetDevOpenFormatCellsDialog = (row, col) => {
      openFormatCellsDialog(Number(row), Number(col));
      return Boolean(document.querySelector("#formatCellsDialog"));
    };
    window.__webSheetDevSelectKeyboardColumns = () => {
      selectColumnsByKeyboard();
      const range = state.selectionRange;
      return range ? rangeToLabel(range) : "";
    };
    window.__webSheetDevSelectKeyboardRows = () => {
      selectRowsByKeyboard();
      const range = state.selectionRange;
      return range ? rangeToLabel(range) : "";
    };
    window.__webSheetDevCellDisplay = (sheetName, row, col) => {
      const sheet = state.model?.sheets?.find((candidate) => candidate.name === sheetName);
      if (!sheet) return "";
      return getDisplayForCell(sheet, getSheetId(sheet.name), Number(row), Number(col));
    };
    window.__webSheetDevDisplayValue = (value, numFmt = "") => displayValue(value, numFmt);
    window.__webSheetDevClipboardSessionRoundTrip = async () => {
      const range = activeSelectionRange();
      if (!range) return { restored: false, reason: "no-selection" };
      await copySelectedRange({ systemClipboard: false });
      const copied = cloneJsonLikeValue(state.internalClipboard);
      state.internalClipboard = null;
      const restored = restorePersistedInternalClipboard();
      return {
        restored,
        text: state.internalClipboard?.text || "",
        rowCount: state.internalClipboard?.rowCount || 0,
        colCount: state.internalClipboard?.colCount || 0,
        cellCount: clipboardCellEntries(state.internalClipboard).length,
        sourceLabel: copied?.sourceRange ? rangeToLabel(copied.sourceRange) : "",
      };
    };
    window.__webSheetDevFormulaEngineCell = (sheetName, row, col) => {
      const sheetId = getSheetId(sheetName);
      if (!state.hf || sheetId == null) return null;
      const address = { sheet: sheetId, row: Number(row) - 1, col: Number(col) - 1 };
      try {
        return {
          serialized: state.hf.getCellSerialized(address),
          formula: state.hf.getCellFormula(address),
          value: cloneJsonLikeValue(state.hf.getCellValue(address)),
        };
      } catch (error) {
        return { error: error?.message || String(error) };
      }
    };
    window.__webSheetDevNamedExpressionValue = (name) => {
      try {
        return {
          expression: state.hf?.getNamedExpressionFormula?.(name),
          value: cloneJsonLikeValue(state.hf?.getNamedExpressionValue?.(name)),
        };
      } catch (error) {
        return { error: error?.message || String(error) };
      }
    };
  }
  installStandaloneDiagnostics();

  const standaloneOpenFile = standaloneOpenWorkbookFileFromWindow();
  const transferToken = workbookOpenTransferTokenFromUrl();
  const standaloneModel = standaloneWorkbookModelFromWindow();
  if (standaloneOpenFile) {
    initializeStandaloneLauncher({ showBlank: false, status: `${standaloneOpenFile.name} を読み込み中...` });
    void loadStandaloneOpenWorkbookFile(standaloneOpenFile);
  } else if (transferToken) {
    if (state.isStandalone) initializeStandaloneLauncher({ showBlank: false, status: "ワークブックを読み込み中..." });
    void loadWorkbookOpenTransferFromUrl(transferToken);
  } else if (standaloneModel) {
    void loadStandaloneWorkbookModel(standaloneModel);
  } else if (state.isStandalone) {
    initializeStandaloneLauncher();
  } else {
    void restorePersistedWorkbookOrCreateBlank();
  }
}



function cwsAgentCellValidationInfo(sheet, row, col, options = {}) {
  const rule = sheet?.cells?.[cellKey(row, col)]?.validation;
  if (!rule?.type) return null;
  const type = String(rule.type || "");
  const info = {
    type,
    allowBlank: rule.allowBlank !== false,
    showErrorMessage: rule.showErrorMessage !== false,
  };
  if (type === "list") {
    const maxOptions = Math.max(1, Math.min(50, Math.floor(Number(options.maxOptions) || 20)));
    const allOptions = validationListOptions(sheet, rule)
      .map((item) => String(item ?? "").trim())
      .filter(Boolean);
    info.showDropDown = rule.showDropDown !== false;
    info.options = allOptions.slice(0, maxOptions);
    info.optionsTruncated = allOptions.length > maxOptions;
  }
  return info;
}



function finishStartupBoot() {
  document.body?.removeAttribute("data-app-booting");
  $bootState.addClass("hidden");
}



function showStartupLoadingState(message = "Cht WebSheet を読み込み中...") {
  document.body?.setAttribute("data-app-booting", "");
  $bootState.removeClass("hidden");
  $bootState.find("[data-boot-status]").text(message);
  $emptyState.addClass("hidden");
  $sheetWorkspace.addClass("hidden");
}



function showWorkbookWorkspace() {
  finishStartupBoot();
  $emptyState.addClass("hidden");
  $sheetWorkspace.removeClass("hidden");
}



function showFileOpenEmptyState() {
  finishStartupBoot();
  $emptyState.removeClass("hidden");
  $sheetWorkspace.addClass("hidden");
}



async function restorePersistedWorkbookOrCreateBlank() {
  showStartupLoadingState();
  const restoreVersion = state.workbookLoadVersion;
  if (consumePendingFileLoadRecoveryMarker()) {
    try {
      await deletePersistedWorkbookSnapshot();
    } catch (error) {
      console.warn("Failed to clear workbook snapshot after interrupted file load.", error);
    }
    if (!state.model && restoreVersion === state.workbookLoadVersion) {
      openBlankWorkbook({ status: "前回のファイル読み込みが完了しなかったため、新しいブックを作成しました。" });
      showAppAlert("前回のファイル読み込みが完了しなかったため、自動復元をスキップしました。", "warning");
      void flushWorkbookPersistence();
    }
    return;
  }
  const restored = await restorePersistedWorkbook();
  if (!restored && !state.model && restoreVersion === state.workbookLoadVersion) {
    openBlankWorkbook({ status: "新しいブックを作成しました。" });
    void flushWorkbookPersistence();
  }
}



function handleWorkbookFileDragOver(event) {
  if (!state.model || !dataTransferMayContainWorkbookFile(event.dataTransfer)) return;
  event.preventDefault();
  if (event.dataTransfer) event.dataTransfer.dropEffect = "copy";
}



function handleWorkbookFileDrop(event) {
  if (!state.model) return;
  const file = workbookFileFromDataTransfer(event.dataTransfer);
  if (!file) return;
  event.preventDefault();
  event.stopPropagation();
  if (event.stopImmediatePropagation) event.stopImmediatePropagation();
  void openDroppedWorkbookFileInNewTab(file);
}



function dataTransferMayContainWorkbookFile(dataTransfer) {
  if (!dataTransfer) return false;
  const files = Array.from(dataTransfer.files || []);
  if (files.length) return files.some(isWorkbookFile);
  const items = Array.from(dataTransfer.items || []);
  if (!items.length) return false;
  return items.some((item) => item?.kind === "file" && (!item.type || isWorkbookMimeType(item.type)));
}



function workbookFileFromDataTransfer(dataTransfer) {
  return Array.from(dataTransfer?.files || []).find(isWorkbookFile) || null;
}



function isWorkbookFile(file) {
  if (!file) return false;
  const name = String(file.name || "").toLowerCase();
  return /\.(html?|xlsx|xlsm|xls)$/i.test(name) || isWorkbookMimeType(file.type);
}



function isWorkbookMimeType(type) {
  return /text\/html|spreadsheetml\.sheet|spreadsheetml\.macroenabled|ms-excel(?:\.sheet\.macroenabled)?/i.test(String(type || ""));
}



function isCwsHtmlWorkbookFile(file) {
  const name = String(file?.name || "").toLowerCase();
  return /\.html?$/i.test(name) || String(file?.type || "").toLowerCase() === "text/html";
}



async function openWorkbookFilePicker() {
  if (!window.showOpenFilePicker) {
    openWorkbookFileInputFallback();
    return;
  }
  try {
    setStatus("開くファイルを選択してください。");
    const [handle] = await showOpenWorkbookFilePicker(workbookOpenFilePickerOptions());
    if (!handle) return;
    const file = await handle.getFile();
    await openSelectedWorkbookFileWithUnsavedPrompt(file, { fileHandle: handle });
  } catch (error) {
    if (error?.name === "AbortError") {
      setStatus(currentWorkbookStatusMessage() || "ファイル選択をキャンセルしました。");
      return;
    }
    console.error(error);
    showMessageBox(`ファイルを開けませんでした：${readableLoadError(error)}`, { title: "Cht WebSheet" });
    setStatus("ファイルを開けませんでした。");
  }
}



function workbookOpenFilePickerOptions() {
  const options = {
    id: "cws-open",
    multiple: false,
    types: [
      {
        description: "CWS HTML / Excel",
        accept: {
          "text/html": [".html", ".htm"],
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
          "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"],
          "application/vnd.ms-excel": [".xls"],
        },
      },
    ],
    excludeAcceptAllOption: false,
  };
  const startIn = workbookOpenFilePickerStartIn();
  if (startIn) options.startIn = startIn;
  return options;
}



function workbookOpenFilePickerStartIn() {
  if (state.currentSaveDirectoryHandle?.getFileHandle) return state.currentSaveDirectoryHandle;
  if (state.currentSaveFileHandle?.createWritable) return state.currentSaveFileHandle;
  return standaloneOpenedFilePickerStartIn();
}



async function showOpenWorkbookFilePicker(options) {
  try {
    return await window.showOpenFilePicker(options);
  } catch (error) {
    if (error?.name === "AbortError" || !options || !("startIn" in options)) throw error;
    const fallbackOptions = { ...options };
    delete fallbackOptions.startIn;
    return window.showOpenFilePicker(fallbackOptions);
  }
}



function openWorkbookFileInputFallback() {
  $fileInput.trigger("click");
  setStatus("開くファイルを選択してください。");
}



async function openSelectedWorkbookFileWithUnsavedPrompt(file, options = {}) {
  if (!file) return false;
  if (!shouldConfirmBeforeReplacingWorkbook()) {
    return openSelectedWorkbookFile(file, options);
  }
  showUnsavedWorkbookReplacementMessage({
    message: "現在のブックに保存されていない変更があります。別のファイルを開く前に保存しますか?",
    saveStatus: "保存してからファイルを開きました。",
    discardStatus: "保存せずにファイルを開きました。",
    cancelStatus: "ファイルを開く操作をキャンセルしました。",
    action: async () => openSelectedWorkbookFile(file, options),
  });
  return false;
}



function shouldConfirmBeforeReplacingWorkbook() {
  return Boolean(state.model && workbookHasUnsavedChanges());
}



function showUnsavedWorkbookReplacementMessage(options = {}) {
  showMessageBox(options.message || "現在のブックに保存されていない変更があります。続行する前に保存しますか?", {
    title: "Cht WebSheet",
    tone: "warning",
    buttons: [
      {
        label: "保存する(S)",
        result: "save",
        callback: async () => {
          const saved = await quickSaveWorkbook();
          if (!saved) return;
          const opened = await options.action?.();
          if (opened && options.saveStatus) setStatus(options.saveStatus);
        },
      },
      {
        label: "保存しない(N)",
        result: "discard",
        callback: async () => {
          const opened = await options.action?.();
          if (opened && options.discardStatus) setStatus(options.discardStatus);
        },
      },
      {
        label: "キャンセル",
        result: "cancel",
        callback: () => {
          if (options.cancelStatus) setStatus(options.cancelStatus);
        },
      },
    ],
    cancelResult: "cancel",
  });
}



async function openSelectedWorkbookFile(file, options = {}) {
  if (!file) return false;
  if (isCwsHtmlWorkbookFile(file)) {
    // Treat imported HTML as workbook data on both hosted and standalone pages.
    // Keep the current trusted runtime; never execute scripts from the file.
    return loadCwsHtmlWorkbookFileInCurrentDocument(file, { fileHandle: options.fileHandle || null });
  }
  await loadExcelFile(file, { fileHandle: options.fileHandle || null });
  return true;
}



async function loadCwsHtmlWorkbookFileInCurrentDocument(file, options = {}) {
  try {
    setStatus(`${file.name} を開いています...`);
    const html = await file.text();
    const model = cwsHtmlWorkbookModelFromSource(html, file);
    const loaded = await loadStandaloneWorkbookModel(model, { openedFileName: file.name || model.sourceName || "workbook.html" });
    if (loaded === false) return false;
    applyOpenedCwsHtmlSaveTarget(file, options.fileHandle);
    markWorkbookSavedClean();
    scheduleWorkbookPersistence();
    setStatus("HTML ワークブックを読み込みました。");
    return true;
  } catch (error) {
    console.error(error);
    showMessageBox(`HTML ワークブックを開けませんでした：${readableLoadError(error)}`, { title: "Cht WebSheet" });
    setStatus("HTML ワークブックを開けませんでした。");
    return false;
  }
}



async function openDroppedWorkbookFileInNewTab(file) {
  if (isSameOpenedWorkbookFile(file)) {
    setStatus(`${file.name} は既にこのタブで開いています。`);
    return;
  }

  const openedTab = window.open("about:blank", "_blank");
  if (!openedTab) {
    showMessageBox("新しいタブを開けませんでした。ブラウザーのポップアップ ブロック設定を確認してください。", { title: "Cht WebSheet" });
    return;
  }

  const token = createWorkbookOpenTransferToken();
  try {
    setStatus(`${file.name} を新しいタブで開いています...`);
    if (isCwsHtmlWorkbookFile(file)) {
      if (await openCwsHtmlWorkbookFileInNewTab(file, openedTab)) finishDroppedWorkbookOpenStatus(file);
      return;
    }
    if (shouldOpenDroppedWorkbookWithStandalonePayload()) {
      await writeStandaloneWorkbookOpenTab(openedTab, file);
      finishDroppedWorkbookOpenStatus(file);
      return;
    }
    await writeWorkbookOpenTransfer(token, file);
    openedTab.location.href = workbookOpenTransferUrl(token);
    finishDroppedWorkbookOpenStatus(file);
  } catch (error) {
    console.error(error);
    try {
      openedTab.close();
    } catch {
      // Ignore tabs that the browser no longer lets us close.
    }
    showMessageBox(`新しいタブでファイルを開けませんでした：${readableLoadError(error)}`, { title: "Cht WebSheet" });
    setStatus("新しいタブでファイルを開けませんでした。");
  }
}



async function openCwsHtmlWorkbookFileInNewTab(file, openedTab = null, options = {}) {
  const targetTab = openedTab || window.open("about:blank", "_blank");
  if (!targetTab) {
    showMessageBox("新しいタブを開けませんでした。ブラウザーのポップアップ ブロック設定を確認してください。", { title: "Cht WebSheet" });
    setStatus("HTML ワークブックを開けませんでした。");
    return false;
  }
  try {
    setStatus(`${file.name} を新しいタブで開いています...`);
    const url = await createCwsHtmlWorkbookBlobUrl(file, options);
    targetTab.location.href = url;
    window.setTimeout(() => URL.revokeObjectURL(url), 5 * 60 * 1000);
    setStatus(`${file.name} を新しいタブで開きました。`);
    return true;
  } catch (error) {
    console.error(error);
    try {
      targetTab.close();
    } catch {
      // Ignore tabs that the browser no longer lets us close.
    }
    showMessageBox(`HTML ワークブックを開けませんでした：${readableLoadError(error)}`, { title: "Cht WebSheet" });
    setStatus("HTML ワークブックを開けませんでした。");
    return false;
  }
}



async function openCwsHtmlWorkbookFileInCurrentTab(file, options = {}) {
  try {
    setStatus(`${file.name} を開いています...`);
    const url = await createCwsHtmlWorkbookBlobUrl(file, options);
    state.suppressBeforeUnloadWarning = true;
    window.location.href = url;
    return true;
  } catch (error) {
    console.error(error);
    showMessageBox(`HTML ワークブックを開けませんでした：${readableLoadError(error)}`, { title: "Cht WebSheet" });
    setStatus("HTML ワークブックを開けませんでした。");
    return false;
  }
}



async function createCwsHtmlWorkbookBlobUrl(file, options = {}) {
  const html = await file.text();
  const model = cwsHtmlWorkbookModelFromSource(html, file);
  const source = injectCwsHtmlOpenFileHandleToken(await buildStandaloneHtml(model), options.saveHandleToken);
  return URL.createObjectURL(new Blob([source], { type: "text/html" }));
}



function cwsHtmlWorkbookModelScript(doc) {
  // Early workbook exports used webexcel-model before the CWS format markers existed.
  return doc.querySelector('script#websheet-model[type="application/json"]') ||
    doc.querySelector('script#webexcel-model[type="application/json"]');
}



function cwsHtmlWorkbookModelFromSource(source, file = null) {
  const doc = new DOMParser().parseFromString(String(source || ""), "text/html");
  const script = cwsHtmlWorkbookModelScript(doc);
  if (!script) throw new Error("HTML 内のワークブック データにワークシート情報がありません。");
  const model = extractStandaloneWorkbookPayload(script.textContent || "{}");
  if (!hasStandaloneWorkbookSheets(model)) throw new Error("HTML 内のワークブック データにワークシート情報がありません。");
  model.sourceName = fileNameWithExtension(file?.name || model.sourceName || "workbook.html", "html");
  model.saveType = "html";
  model.sourceFileSignature = workbookFileSignature(file);
  return model;
}



function applyOpenedCwsHtmlSaveTarget(file, fileHandle) {
  if (!state.model) return false;
  const fileName = fileNameWithExtension(fileHandle?.name || file?.name || state.model.sourceName || "workbook.html", "html");
  state.model.sourceName = fileName;
  state.model.saveType = "html";
  state.currentSaveType = "html";
  state.openedFileSignature = normalizeWorkbookFileSignature(state.model.sourceFileSignature);
  state.currentSaveFileHandle = fileHandle?.createWritable ? fileHandle : null;
  state.currentSaveDirectoryHandle = null;
  state.currentSaveFileName = state.currentSaveFileHandle ? fileName : "";
  document.title = fileName;
  $workbookName.text(fileName);
  return Boolean(state.currentSaveFileHandle);
}



function injectCwsHtmlOpenFileHandleToken(source, saveHandleToken = "") {
  const token = String(saveHandleToken || "").trim();
  if (!token) return source;
  const tokenScript = `<script>window.__WEBSHEET_OPENED_FILE_HANDLE_TOKEN__ = ${safeJson(token)};</script>`;
  if (/<body\b[^>]*>/i.test(source)) return source.replace(/<body\b[^>]*>/i, (match) => `${match}\n${tokenScript}`);
  return `${tokenScript}\n${source}`;
}



function cwsHtmlWorkbookSourceForOpen(source, file = null, options = {}) {
  const html = String(source || "");
  const openedName = fileNameWithExtension(file?.name || "workbook.html", "html");
  const saveHandleToken = String(options.saveHandleToken || "").trim();
  try {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const script = cwsHtmlWorkbookModelScript(doc);
    if (!script) return html;
    const model = JSON.parse(script.textContent || "{}");
    if (!hasStandaloneWorkbookSheets(model)) return html;
    model.sourceName = openedName;
    model.saveType = "html";
    model.sourceFileSignature = workbookFileSignature(file);
    script.textContent = safeJson(model);
    let title = doc.querySelector("title");
    if (!title) {
      title = doc.createElement("title");
      doc.head?.appendChild(title);
    }
    title.textContent = `${openedName} - WebSheet`;
    const bootName = doc.querySelector("#websheetStandaloneBoot strong");
    if (bootName) bootName.textContent = openedName;
    if (saveHandleToken) {
      const tokenScript = doc.createElement("script");
      tokenScript.textContent = `window.__WEBSHEET_OPENED_FILE_HANDLE_TOKEN__ = ${safeJson(saveHandleToken)};`;
      doc.body?.insertBefore(tokenScript, doc.body.firstChild);
    }
    return `<!doctype html>\n${doc.documentElement.outerHTML}`;
  } catch (error) {
    console.warn("Failed to rewrite opened CWS HTML file name.", error);
    return html;
  }
}



function finishDroppedWorkbookOpenStatus(file) {
  const openedMessage = `${file.name} を新しいタブで開きました。`;
  const readyMessage = currentWorkbookStatusMessage();
  setStatus(openedMessage);
  window.setTimeout(() => {
    if ($statusText.text() === openedMessage) setStatus(readyMessage);
  }, 2500);
}



function currentWorkbookStatusMessage() {
  const sheetCount = Number(state.model?.sheets?.length) || 0;
  return sheetCount
    ? `${sheetCount} 個のワークシートを読み込みました。セルは直接編集でき、数式は自動再計算されます。`
    : "";
}



function shouldOpenDroppedWorkbookWithStandalonePayload() {
  return isLocalStandaloneDocument();
}



async function writeStandaloneWorkbookOpenTab(openedTab, file) {
  const buffer = await file.arrayBuffer();
  const payload = {
    name: file.name || "workbook.xlsx",
    type: file.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    size: Number(file.size) || buffer.byteLength || 0,
    lastModified: Number(file.lastModified) || Date.now(),
    base64: bytesToBase64(buffer),
  };
  const html = await buildStandaloneOpenWorkbookHtml(payload);
  const localFile = await createStandaloneOpenWorkbookLocalHtml(html, payload.name);
  if (localFile?.url) {
    openedTab.location.href = localFile.url;
    return;
  }
  openedTab.document.open();
  openedTab.document.write(html);
  openedTab.document.close();
}



async function buildStandaloneOpenWorkbookHtml(filePayload) {
  const fileName = sanitizeFileName(filePayload?.name) || "workbook.xlsx";
  const versionInfo = await standaloneRuntimeVersionInfo();
  const runtimeVersion = versionInfoRuntimeVersion(versionInfo) || STANDALONE_RUNTIME_VERSION;
  const runtimeBaseUrl = standaloneRuntimeAbsoluteBaseUrl(standaloneRuntimeBaseUrl());
  const versionedRuntimeBaseUrl = standaloneRuntimeAbsoluteBaseUrl(standaloneRuntimeVersionedBaseUrl(runtimeVersion));
  const remoteRuntimeBaseUrl = standaloneRuntimeAbsoluteBaseUrl(standaloneRemoteRuntimeBaseUrl());
  const versionedRemoteRuntimeBaseUrl = standaloneRuntimeAbsoluteBaseUrl(standaloneVersionedRemoteRuntimeBaseUrl(runtimeVersion));
  return `<!doctype html>
<html ${cwsHtmlRootAttributes()}>
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
${cwsHtmlMetaTags(versionInfo)}
<link rel="icon" type="image/png" href="${escapeAttr(standaloneDocumentIconUrl(runtimeBaseUrl))}" />
<title>${escapeHtml(fileName)} - WebSheet</title>
</head>
<body>
<div id="websheetStandaloneBoot" style="font:14px system-ui,sans-serif;padding:16px;color:#1f2937">
  <strong>${escapeHtml(fileName)}</strong>
  <div>WebSheet 共通ランタイムを読み込み中...</div>
</div>
${cwsAiNoticeComment()}
${cwsModelScriptTag("{}")}
${cwsAiInstructionsScript()}
<script>window.__WEBSHEET_OPEN_FILE__ = ${safeJson(filePayload)};</script>
<script>
${standaloneLoaderScript({
  version: runtimeVersion,
  localBaseUrl: runtimeBaseUrl,
  versionedLocalBaseUrl: versionedRuntimeBaseUrl,
  remoteBaseUrl: remoteRuntimeBaseUrl,
  versionedRemoteBaseUrl: versionedRemoteRuntimeBaseUrl,
  localSaveEndpoint: standaloneRuntimeAbsoluteUrl(standaloneLocalSaveEndpointUrl()),
})}
</script>
</body>
</html>`;
}



function cwsHtmlRootAttributes() {
  return [
    `lang="ja"`,
    `data-cws-format="${escapeAttr(CWS_HTML_FORMAT)}"`,
    `data-cws-schema-version="${escapeAttr(CWS_SCHEMA_VERSION)}"`,
    `data-cws-app="${escapeAttr(APP_SHORT_NAME)}"`,
  ].join(" ");
}



function cwsHtmlMetaTags(versionInfo = activeRuntimeVersionInfo()) {
  const appVersion = versionInfoAppVersion(versionInfo) || APP_VERSION;
  return [
    `<meta name="generator" content="${escapeAttr(`${APP_DISPLAY_NAME} ${appVersion}`)}" />`,
    `<meta name="cws:format" content="${escapeAttr(CWS_HTML_FORMAT)}" />`,
    `<meta name="cws:schema-version" content="${escapeAttr(CWS_SCHEMA_VERSION)}" />`,
    `<meta name="cws:guide" content="${escapeAttr(CWS_AI_GUIDE_URL)}" />`,
    `<meta name="cws:schema" content="${escapeAttr(CWS_MODEL_SCHEMA_URL)}" />`,
  ].join("\n");
}



function cwsModelScriptTag(payload) {
  return `<script id="websheet-model" type="application/json" data-cws-model="workbook" data-cws-format="${escapeAttr(CWS_HTML_FORMAT)}" data-cws-schema-version="${escapeAttr(CWS_SCHEMA_VERSION)}" data-cws-editable="true">${payload}</script>`;
}



function cwsAiNoticeComment() {
  return `<!--
Cht WebSheet (CWS) AI Notice:
This file is a Cht WebSheet (CWS) HTML workbook.
The actual workbook data is the JSON inside <script id="websheet-model" type="application/json">.
Only edit that JSON when changing workbook content.
Do not edit HTML/CSS/JS runtime, loader code, meta tags, or external links.
For usage instructions, read <script id="cws-ai-instructions" type="application/json"> and its guideUrl.
-->`;
}



function cwsAiInstructionsScript() {
  return `<script id="cws-ai-instructions" type="application/json">${safeJson(cwsAiInstructionsPayload())}</script>`;
}



function cwsAiInstructionsPayload() {
  return {
    notice: "This file is a Cht WebSheet (CWS) HTML workbook. The actual workbook data is the JSON inside script#websheet-model. Only edit that JSON when changing workbook content. Do not edit HTML/CSS/JS runtime, loader code, meta tags, or external links. For usage instructions, read this script#cws-ai-instructions object and its guideUrl.",
    format: CWS_HTML_FORMAT,
    appName: APP_DISPLAY_NAME,
    shortName: APP_SHORT_NAME,
    schemaVersion: CWS_SCHEMA_VERSION,
    modelScriptId: "websheet-model",
    editableRegion: "script#websheet-model",
    guideUrl: CWS_AI_GUIDE_URL,
    schemaUrl: CWS_MODEL_SCHEMA_URL,
    rules: [
      "This file is a Cht WebSheet (CWS) HTML workbook.",
      "Only edit the workbook JSON inside script#websheet-model.",
      "Do not edit runtime scripts, styles, loader code, meta tags, or external links unless explicitly requested.",
      "Keep the JSON valid and preserve existing ids, sheet names, formulas, images, shapes, merges, filters, and references unless the requested change requires updating them.",
    ],
  };
}



async function createStandaloneOpenWorkbookLocalHtml(html, fileName) {
  const endpoint = standaloneLocalOpenWorkbookEndpoint();
  const currentPath = standaloneOpenedFilePath();
  if (!endpoint || !currentPath) return null;
  try {
    const response = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        currentPath,
        fileName: fileNameWithExtension(fileName || "workbook.xlsx", "html"),
        html,
      }),
    });
    let payload = null;
    try {
      payload = await response.json();
    } catch {
      // Keep the fallback path available when a local bridge returns non-JSON.
    }
    if (!response.ok || payload?.ok === false) {
      const message = payload?.message || `${response.status} ${response.statusText}`.trim();
      throw new Error(message || "ローカル HTML ファイルを作成できませんでした。");
    }
    return payload;
  } catch (error) {
    console.warn("WebSheet local open bridge failed; falling back to temporary tab.", error);
    showAppAlert(`新しいローカル HTML ファイルを作成できませんでした。保存時はダウンロードになります：${readableLoadError(error)}`, "warning");
    return null;
  }
}



function standaloneLocalOpenWorkbookEndpoint() {
  const saveEndpoint = standaloneLocalSaveEndpoint();
  if (!saveEndpoint) return "";
  try {
    return new URL("/__websheet/local-open-workbook", saveEndpoint).href;
  } catch {
    return "";
  }
}



function standaloneRuntimeAbsoluteBaseUrl(value) {
  return normalizeRuntimeBaseUrl(standaloneRuntimeAbsoluteUrl(value));
}



function standaloneRuntimeAbsoluteUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  try {
    return new URL(raw, window.location.href).href;
  } catch {
    return raw;
  }
}



function isSameOpenedWorkbookFile(file) {
  const current = normalizeWorkbookFileSignature(state.openedFileSignature || state.model?.sourceFileSignature);
  if (!current) return false;
  const next = workbookFileSignature(file);
  return workbookFileSignatureKey(current) === workbookFileSignatureKey(next);
}



function workbookFileSignature(file) {
  if (!file) return null;
  return {
    name: String(file.name || ""),
    size: Number(file.size) || 0,
    lastModified: Number(file.lastModified) || 0,
  };
}



function normalizeWorkbookFileSignature(signature) {
  if (!signature || typeof signature !== "object") return null;
  const name = String(signature.name || "");
  if (!name) return null;
  return {
    name,
    size: Number(signature.size) || 0,
    lastModified: Number(signature.lastModified) || 0,
  };
}



function workbookFileSignatureKey(signature) {
  const normalized = normalizeWorkbookFileSignature(signature);
  return normalized ? `${normalized.name}\u001f${normalized.size}\u001f${normalized.lastModified}` : "";
}



function createWorkbookOpenTransferToken() {
  if (crypto?.randomUUID) return crypto.randomUUID();
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}



function createWorkbookDocumentId() {
  if (crypto?.randomUUID) return crypto.randomUUID();
  return `wb-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 12)}`;
}



function workbookOpenTransferUrl(token) {
  const url = new URL(window.location.href);
  url.searchParams.set(WORKBOOK_OPEN_TRANSFER_QUERY_PARAM, token);
  url.hash = "";
  return url.href;
}



function workbookOpenTransferTokenFromUrl() {
  try {
    return new URL(window.location.href).searchParams.get(WORKBOOK_OPEN_TRANSFER_QUERY_PARAM) || "";
  } catch {
    return "";
  }
}



function removeWorkbookOpenTransferTokenFromUrl() {
  try {
    const url = new URL(window.location.href);
    if (!url.searchParams.has(WORKBOOK_OPEN_TRANSFER_QUERY_PARAM)) return;
    url.searchParams.delete(WORKBOOK_OPEN_TRANSFER_QUERY_PARAM);
    window.history.replaceState(window.history.state, document.title, `${url.pathname}${url.search}${url.hash}`);
  } catch {
    // The token is only cosmetic after load; ignore URL cleanup failures.
  }
}



async function loadWorkbookOpenTransferFromUrl(token = workbookOpenTransferTokenFromUrl()) {
  removeWorkbookOpenTransferTokenFromUrl();
  if (!token) return false;
  setStatus("新しいタブでファイルを開いています...");
  try {
    const payload = await readWorkbookOpenTransfer(token);
    await deleteWorkbookOpenTransfer(token);
    if (!payload?.buffer || !payload.name) {
      throw new Error("転送されたファイル データが見つかりません。");
    }
    const file = new File([payload.buffer], payload.name, {
      type: payload.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      lastModified: Number(payload.lastModified) || Date.now(),
    });
    await loadExcelFile(file);
    return true;
  } catch (error) {
    console.error(error);
    setStatus("新しいタブでファイルを開けませんでした。");
    showAppAlert(`新しいタブでファイルを開けませんでした：${readableLoadError(error)}`, "error");
    return true;
  }
}



async function writeWorkbookOpenTransfer(token, file) {
  await cleanupExpiredWorkbookOpenTransfers();
  const payload = {
    token,
    name: file.name || "workbook.xlsx",
    type: file.type || "",
    size: Number(file.size) || 0,
    lastModified: Number(file.lastModified) || 0,
    createdAt: Date.now(),
    buffer: await file.arrayBuffer(),
  };
  const db = await openWorkbookOpenTransferDb();
  try {
    const tx = db.transaction(WORKBOOK_OPEN_TRANSFER_STORE_NAME, "readwrite");
    tx.objectStore(WORKBOOK_OPEN_TRANSFER_STORE_NAME).put(payload, token);
    await idbTransactionDone(tx);
  } finally {
    db.close();
  }
}



async function readWorkbookOpenTransfer(token) {
  const db = await openWorkbookOpenTransferDb();
  try {
    const payload = await idbRequest(db.transaction(WORKBOOK_OPEN_TRANSFER_STORE_NAME, "readonly").objectStore(WORKBOOK_OPEN_TRANSFER_STORE_NAME).get(token));
    if (!payload) return null;
    if (Date.now() - Number(payload.createdAt || 0) > WORKBOOK_OPEN_TRANSFER_MAX_AGE_MS) {
      await deleteWorkbookOpenTransfer(token);
      return null;
    }
    return payload;
  } finally {
    db.close();
  }
}



async function deleteWorkbookOpenTransfer(token) {
  const db = await openWorkbookOpenTransferDb();
  try {
    const tx = db.transaction(WORKBOOK_OPEN_TRANSFER_STORE_NAME, "readwrite");
    tx.objectStore(WORKBOOK_OPEN_TRANSFER_STORE_NAME).delete(token);
    await idbTransactionDone(tx);
  } finally {
    db.close();
  }
}



async function cleanupExpiredWorkbookOpenTransfers() {
  if (typeof indexedDB === "undefined") return;
  let db;
  try {
    db = await openWorkbookOpenTransferDb();
    const store = db.transaction(WORKBOOK_OPEN_TRANSFER_STORE_NAME, "readonly").objectStore(WORKBOOK_OPEN_TRANSFER_STORE_NAME);
    const records = await idbRequest(store.getAll());
    const expiredTokens = (records || [])
      .filter((record) => Date.now() - Number(record?.createdAt || 0) > WORKBOOK_OPEN_TRANSFER_MAX_AGE_MS)
      .map((record) => record.token)
      .filter(Boolean);
    db.close();
    db = null;
    if (!expiredTokens.length) return;
    db = await openWorkbookOpenTransferDb();
    const tx = db.transaction(WORKBOOK_OPEN_TRANSFER_STORE_NAME, "readwrite");
    const writeStore = tx.objectStore(WORKBOOK_OPEN_TRANSFER_STORE_NAME);
    expiredTokens.forEach((expiredToken) => writeStore.delete(expiredToken));
    await idbTransactionDone(tx);
  } catch (error) {
    console.warn("Failed to clean transferred workbook files.", error);
  } finally {
    db?.close();
  }
}



function openWorkbookOpenTransferDb() {
  if (typeof indexedDB === "undefined") return Promise.reject(new Error("IndexedDB を利用できません。"));
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(WORKBOOK_OPEN_TRANSFER_DB_NAME, WORKBOOK_OPEN_TRANSFER_DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(WORKBOOK_OPEN_TRANSFER_STORE_NAME)) {
        db.createObjectStore(WORKBOOK_OPEN_TRANSFER_STORE_NAME);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    request.onblocked = () => reject(new Error("IndexedDB upgrade was blocked."));
  });
}



function installStandaloneDiagnostics() {
  if (!state.isStandalone || typeof window === "undefined") return;
  window.__webSheetStandaloneDiagnostics = () => {
    if (state.model) saveActiveWindowViewState();
    const sheet = activeSheet();
    return {
      activeSheetIndex: state.activeSheetIndex,
      activeSheetName: sheet?.name || "",
      sourceName: state.model?.sourceName || "",
      saveType: state.model?.saveType || "",
      currentSaveType: state.currentSaveType || "",
      isStandalone: state.isStandalone === true,
      runtimeSource: window.__WEBSHEET_RUNTIME_SOURCE__ || "",
      workbookLoaded: Boolean(state.model?.sheets?.length),
      sheetNames: (state.model?.sheets || []).map((sheet) => sheet.name),
      activeSheetView: standaloneSheetViewDiagnostics(sheet),
      workbookView: standaloneWorkbookViewDiagnostics(),
      memoryStats: workbookMemoryStats(state.model),
      status: document.querySelector("#statusText")?.textContent || "",
    };
  };
  window.__webSheetStandaloneCell = (sheetName, row, col) => {
    const sheet = state.model?.sheets?.find((candidate) => candidate.name === sheetName);
    if (!sheet) return null;
    const sheetId = getSheetId(sheet.name);
    const address = { sheet: sheetId, row: Number(row) - 1, col: Number(col) - 1 };
    let engine = null;
    if (state.hf && sheetId != null) {
      try {
        engine = {
          serialized: state.hf.getCellSerialized(address),
          formula: state.hf.getCellFormula(address),
          value: cloneJsonLikeValue(state.hf.getCellValue(address)),
        };
      } catch (error) {
        engine = { error: error?.message || String(error) };
      }
    }
    const sourceCell = sheet.cells?.[cellKey(Number(row), Number(col))] || null;
    const cell = sheetDisplayCell(sheet, Number(row), Number(col), sourceCell) || null;
    return {
      sheetName: sheet.name,
      row: Number(row),
      col: Number(col),
      cell: cell ? cloneJsonLikeValue(cell) : null,
      display: getDisplayForCell(sheet, sheetId, Number(row), Number(col)),
      engine,
    };
  };
}



function standaloneSheetViewDiagnostics(sheet) {
  if (!sheet) return null;
  const viewMode = normalizedSheetViewMode(sheet.viewMode);
  return {
    viewMode,
    zoomScale: Math.round(clampZoom(sheet.zoomScale || 100)),
    displayScale: Number(sheet.displayScale || 1),
    showGridLines: sheet.showGridLines !== false,
    showHeadings: sheetHeadingsVisible(sheet),
    showRuler: normalizedSheetShowRuler(sheet.showRuler, viewMode),
    showObjects: sheet.showObjects !== false,
    showOutlineSymbols: sheet.showOutlineSymbols !== false,
    showFormulas: sheet.showFormulas === true,
    showZeros: sheet.showZeros !== false,
    showWhiteSpace: sheet.showWhiteSpace !== false,
    rightToLeft: sheet.rightToLeft === true,
    freezePanes: cloneJsonLikeValue(sheet.freezePanes || null),
    splitPane: cloneJsonLikeValue(sheet.splitPane || null),
  };
}

export {
  activateCwsAgentPageSetupSheet,
  applyCwsAgentAutoFitColumnsOp,
  applyCwsAgentAutoFitRowsOp,
  applyCwsAgentClearRangeOp,
  applyCwsAgentColumnsOp,
  applyCwsAgentHeaderHiddenOp,
  applyCwsAgentMergeCellsOp,
  applyCwsAgentOutlineOp,
  applyCwsAgentResetPageBreaksOp,
  applyCwsAgentRowsOp,
  applyCwsAgentSetColumnWidthOp,
  applyCwsAgentSetRowHeightOp,
  applyCwsAgentUnmergeCellsOp,
  applyOpenedCwsHtmlSaveTarget,
  buildStandaloneOpenWorkbookHtml,
  cleanupExpiredWorkbookOpenTransfers,
  commitCwsAgentSheetViewChange,
  createCwsHtmlWorkbookBlobUrl,
  createStandaloneOpenWorkbookLocalHtml,
  createWorkbookDocumentId,
  createWorkbookOpenTransferToken,
  currentWorkbookStatusMessage,
  cwsAgentCellValidationInfo,
  cwsAgentDimensionsChanged,
  cwsAgentFreezePanesStatus,
  cwsAgentPageBreakPositionLabel,
  cwsAgentSheetViewHasChanges,
  cwsAgentSheetViewPatchLabels,
  cwsAgentSheetViewSnapshot,
  cwsAiInstructionsPayload,
  cwsAiInstructionsScript,
  cwsAiNoticeComment,
  cwsHtmlMetaTags,
  cwsHtmlRootAttributes,
  cwsHtmlWorkbookModelFromSource,
  cwsHtmlWorkbookModelScript,
  cwsHtmlWorkbookSourceForOpen,
  cwsModelScriptTag,
  dataTransferMayContainWorkbookFile,
  deleteWorkbookOpenTransfer,
  finishDroppedWorkbookOpenStatus,
  finishStartupBoot,
  handleWorkbookFileDragOver,
  handleWorkbookFileDrop,
  init,
  injectCwsHtmlOpenFileHandleToken,
  installStandaloneDiagnostics,
  isCwsHtmlWorkbookFile,
  isSameOpenedWorkbookFile,
  isWorkbookFile,
  isWorkbookMimeType,
  loadCwsHtmlWorkbookFileInCurrentDocument,
  loadWorkbookOpenTransferFromUrl,
  normalizeWorkbookFileSignature,
  openCwsHtmlWorkbookFileInCurrentTab,
  openCwsHtmlWorkbookFileInNewTab,
  openDroppedWorkbookFileInNewTab,
  openSelectedWorkbookFile,
  openSelectedWorkbookFileWithUnsavedPrompt,
  openWorkbookFileInputFallback,
  openWorkbookFilePicker,
  openWorkbookOpenTransferDb,
  readWorkbookOpenTransfer,
  refreshCwsAgentWorkbookAfterSheetStructureChange,
  removeWorkbookOpenTransferTokenFromUrl,
  restorePersistedWorkbookOrCreateBlank,
  shouldConfirmBeforeReplacingWorkbook,
  shouldOpenDroppedWorkbookWithStandalonePayload,
  showFileOpenEmptyState,
  showOpenWorkbookFilePicker,
  showStartupLoadingState,
  showUnsavedWorkbookReplacementMessage,
  showWorkbookWorkspace,
  standaloneLocalOpenWorkbookEndpoint,
  standaloneRuntimeAbsoluteBaseUrl,
  standaloneRuntimeAbsoluteUrl,
  standaloneSheetViewDiagnostics,
  workbookFileFromDataTransfer,
  workbookFileSignature,
  workbookFileSignatureKey,
  workbookOpenFilePickerOptions,
  workbookOpenFilePickerStartIn,
  workbookOpenTransferTokenFromUrl,
  workbookOpenTransferUrl,
  writeStandaloneWorkbookOpenTab,
  writeWorkbookOpenTransfer,
};
