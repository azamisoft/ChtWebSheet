// Feature module; shared bindings are imported explicitly.
import {
  $,
  $exportHtmlButton,
  $formulaBar,
  $formulaExpandButton,
  $formulaInput,
  $horizontalScrollbar,
  $sheetHost,
  $toggleFormulasButton,
  $workbookName,
  APP_TITLE,
  DEFAULT_MACRO_SECURITY_LEVEL,
  EXCELJS_IMPORT_IGNORE_NODES,
  ExcelJS,
  FALLBACK_VISIBLE_COLS,
  FALLBACK_VISIBLE_ROWS,
  FORMULA_BAR_DEFAULT_EXPANDED_LINES,
  FORMULA_BAR_MIN_LINES,
  INDEXED_COLORS,
  JSZip,
  LEGACY_XLS_BORDER_STYLE_MAP,
  LEGACY_XLS_BUILTIN_NUM_FMTS,
  LEGACY_XLS_CELL_RECORDS_WITH_XF,
  LEGACY_XLS_FILL_PATTERN_NONE,
  MACRO_SECURITY_LEVELS,
  PENDING_FILE_LOAD_RECOVERY_STORAGE_KEY,
  PERSISTENCE_CURRENT_KEY,
  PERSISTENCE_DB_NAME,
  PERSISTENCE_DB_VERSION,
  PERSISTENCE_IDLE_TIMEOUT_MS,
  PERSISTENCE_SAVE_DEBOUNCE_MS,
  PERSISTENCE_SNAPSHOT_VERSION,
  PERSISTENCE_STORE_NAME,
  SPECIAL_SHEET_INSERT_META,
  WORKBOOK_MODEL_VERSION,
  XLSX,
  createIcons,
  icons,
  loadXlsxModule,
  state,
} from "../../app-core.js";
import {
  resetCellStyleInternPools,
} from "../agent/contracts.js";
import {
  escapeRegExp,
} from "../data/filter-and-text-tools.js";
import {
  clearQueryRefreshTimers,
  reapplyImportedAutoFiltersOnLoad,
  refreshWorkbookQueriesOnOpen,
  scheduleQueryRefreshTimers,
} from "../data/sort-and-autofilter.js";
import {
  workbookRelationships,
} from "../data/validation-and-consolidation.js";
import {
  clampNumber,
  defaultCalloutTailX,
  defaultStarInnerRatio,
  fileExtensionFromImageMimeType,
  isCalloutShapeType,
  isLineShapeObject,
  normalizeShapeAdjustments,
  normalizedHeaderFooterSettings,
  shapeDescriptorForType,
} from "../drawing/shape-rendering.js";
import {
  ensureLineShapeBoxPadding,
  refreshShapeSource,
} from "../drawing/transforms-and-connectors.js";
import {
  closeShapeFormatPane,
  closeShapeSelectionPane,
  closeSymbolDialog,
  normalizeColorValue,
  normalizedThemeColorPalette,
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  clearWorkbookClipboardState,
} from "../editing/keyboard-and-paste.js";
import {
  expandRangeForMerges,
  primeWorkbookFormulaSpillAnchorCache,
  primeWorkbookUsedRangeCache,
  rangeToLabel,
  selectionRangeFromPoints,
  setFormulaBarLines,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  editablePlainText,
  updateFormulaBarForCell,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  cloneJsonLikeValue,
  compactSheetDefaultCellCss,
  internCellModelStylesForSheet,
  internCssObject,
  internSheetCellStyles,
  internWorkbookCellStyles,
  materializeSheetStyleTable,
  normalizeInkStrokes,
  normalizeSheetStyleRanges,
  stableObjectKey,
} from "../editing/structure-and-history.js";
import {
  loadWorkbookIoModules,
} from "../formula/calculations.js";
import {
  cancelAnnotationEditor,
  closeCellContextMenu,
  closeCommentThreadPopup,
  closeHeaderSizeDialog,
  closeLinkDialog,
  closeNameDialog,
  closeNotePopup,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  clearAppAlert,
  clearWarnings,
  columnNumber,
  dateToExcelSerial,
  decodeCell,
  decodeRange,
  isDateFormat,
  readableLoadError,
  sanitizeFileName,
  setStatus,
  showAppAlert,
} from "../ui/help-and-utilities.js";
import {
  closeObjectLinkDialog,
  ensureSheetSize,
  resetLazyScrollLocks,
} from "../ui/ribbon-and-links.js";
import {
  clampZoom,
  closeSheetTabContextMenu,
  closeTranslationPane,
  createBlankSheet,
  scaleSheetDisplay,
  sheetBaseZoomScale,
} from "../ui/sheet-protection-and-translation.js";
import {
  normalizeFormulaAudits,
  normalizeFormulaWatches,
  sheetHeadingsVisible,
} from "../ui/views-and-formula-tools.js";
import {
  borderStyleToCss,
  borderWeight,
  buildFormulaEngine,
  ensureActiveSheetVisible,
  fontFamilyForExcel,
  migrateDenseFormulaDataToCells,
  parseThemeColors,
  renderWorkbook,
  rgbToHex,
  sheetHfDataDimensions,
  verticalAlignToFlex,
  workbookTabSelectedSheetIndexes,
} from "./formulas-and-sheet-tabs.js";
import {
  applyWorkbookCellHyperlinksToModel,
  extractWorkbookAutoFilters,
  extractWorkbookCellHyperlinks,
  extractWorkbookCustomViews,
  extractWorkbookDataValidationDetails,
  extractWorkbookDefinedNames,
  extractWorkbookDrawingShapes,
  extractWorkbookMetadata,
  extractWorkbookOutlineDetails,
  extractWorkbookPhonetics,
  extractWorkbookSheetViewDetails,
  extractWorkbookWindowViews,
  parsePreservedContentTypeOverrides,
  parsePreservedWorkbookRelationships,
  preservedOoxmlEntryPath,
} from "./ooxml-drawings.js";
import {
  activeSheet,
  applyOpenedWorkbookSaveTarget,
  compactWorkbookStyleOnlyCellsToStyleRanges,
  compareBorder,
  defaultSaveTypeForLoadedFile,
  displayValue,
  excelSplitActivePane,
  fileNameWithExtension,
  normalizeWorkbookSaveType,
  parseBase64ImageDataUrl,
  resetCurrentSaveTarget,
  saveTypeForSourceName,
  standaloneOpenedFileName,
  updateHorizontalScrollbar,
} from "./saving-and-export.js";
import {
  createWorkbookDocumentId,
  normalizeWorkbookFileSignature,
  showStartupLoadingState,
  showWorkbookWorkspace,
  workbookFileSignature,
} from "./startup-and-loading.js";
import {
  base64ToBytes,
  bytesToBase64,
  imageMimeType,
  inferWorkbookWidthMetrics,
  normalizeOutlineSettings,
  revokeWorksheetImageRenderObjectUrls,
  scrollRecordFromTopLeftCell,
  workbookMediaToDataUrls,
  worksheetToModel,
} from "./tables-and-validation.js";
import {
  normalizedPageLayoutSettings,
  normalizedPageLayoutThemeEffects,
} from "./vba-and-printing.js";
import {
  activeWorkbookWindowShowOption,
  restoreWorkbookWindowSelection,
  workbookWindowRecords,
} from "./views-and-macros.js";


function standaloneWorkbookViewDiagnostics() {
  if (!state.model) return null;
  const activeWindowId = state.model.activeWindowId || "main";
  const records = workbookWindowRecords();
  return {
    showFormulaBar: !document.querySelector(".formula-bar")?.classList.contains("hidden"),
    activeWindowId,
    mainWindowHidden: state.model.mainWindowHidden === true,
    windowCount: records.length,
    visibleWindowCount: records.filter((view) => !view.hidden).length,
    showHorizontalScroll: activeWorkbookWindowShowOption("showHorizontalScroll"),
    showVerticalScroll: activeWorkbookWindowShowOption("showVerticalScroll"),
    showSheetTabs: activeWorkbookWindowShowOption("showSheetTabs"),
    activeWindowView: cloneJsonLikeValue(records.find((view) => view.id === activeWindowId) || null),
    mainWindowView: cloneJsonLikeValue(state.model.mainWindowView || null),
    windowViews: cloneJsonLikeValue(state.model.windowViews || []),
    viewSideBySide: cloneJsonLikeValue(state.model.viewSideBySide || null),
    synchronousScrolling: state.model.synchronousScrolling === true,
    windowArrangement: cloneJsonLikeValue(state.model.windowArrangement || null),
  };
}



function standaloneWorkbookModelFromWindow() {
  const direct = extractStandaloneWorkbookPayload(window.__WEBSHEET_STANDALONE_MODEL__);
  if (direct) return direct;
  return extractStandaloneWorkbookPayload(window.__WEBSHEET_STANDALONE_MODEL_RAW__);
}



function standaloneOpenWorkbookFileFromWindow() {
  const payload = window.__WEBSHEET_OPEN_FILE__ || window.__WEBSHEET_OPEN_FILE_RAW__;
  window.__WEBSHEET_OPEN_FILE__ = null;
  window.__WEBSHEET_OPEN_FILE_RAW__ = null;
  if (!payload) return null;
  if (typeof payload === "string") {
    try {
      return JSON.parse(payload);
    } catch {
      return null;
    }
  }
  return typeof payload === "object" ? payload : null;
}



async function loadStandaloneOpenWorkbookFile(payload) {
  try {
    const file = standaloneOpenWorkbookPayloadToFile(payload);
    if (!file) throw new Error("転送されたファイル データが見つかりません。");
    await loadExcelFile(file);
    applyStandaloneOpenWorkbookSaveTarget();
    return true;
  } catch (error) {
    console.error(error);
    setStatus("新しいタブでファイルを開けませんでした。");
    showAppAlert(`新しいタブでファイルを開けませんでした：${readableLoadError(error)}`, "error");
    return false;
  }
}



function applyStandaloneOpenWorkbookSaveTarget() {
  if (!state.model) return;
  state.model.saveType = "html";
  state.currentSaveType = "html";
  resetCurrentSaveTarget();
  const openedFileName = standaloneOpenedFileName();
  if (openedFileName) {
    state.model.sourceName = fileNameWithExtension(openedFileName, "html");
    document.title = state.model.sourceName;
    $workbookName.text(state.model.sourceName);
  }
  scheduleWorkbookPersistence();
}



function standaloneOpenWorkbookPayloadToFile(payload) {
  if (!payload || typeof payload !== "object") return null;
  const name = sanitizeFileName(payload.name) || "workbook.xlsx";
  const base64 = String(payload.base64 || "").trim();
  if (!base64) return null;
  const bytes = base64ToBytes(base64);
  if (!bytes.length) return null;
  return new File([bytes], name, {
    type: payload.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    lastModified: Number(payload.lastModified) || Date.now(),
  });
}



function extractStandaloneWorkbookPayload(payload) {
  const seen = new Set();
  const seenText = new Set();
  const queue = [payload];
  while (queue.length) {
    const candidate = queue.shift();
    if (!candidate) continue;
    if (typeof candidate === "string") {
      const text = candidate.trim();
      if (!text || seenText.has(text) || !/^[\[{]/.test(text)) continue;
      seenText.add(text);
      try {
        queue.push(JSON.parse(text));
      } catch {
        // Ignore non-JSON wrapper strings and continue searching sibling payloads.
      }
      continue;
    }
    if (typeof candidate !== "object" || seen.has(candidate)) continue;
    seen.add(candidate);
    if (hasStandaloneWorkbookSheets(candidate)) return candidate;
    ["model", "workbook", "workbookModel", "webSheetWorkbook", "data", "payload"].forEach((key) => {
      if (candidate[key]) queue.push(candidate[key]);
    });
  }
  return null;
}



function hasStandaloneWorkbookSheets(model) {
  return Boolean(model && typeof model === "object" && Array.isArray(model.sheets) && model.sheets.length);
}



function debugWorkbookLoadTrace(label) {
  if (!globalThis.__WEBSHEET_DEBUG_LOAD__) return;
  console.debug(`[websheet-load] ${label}`);
}



function releaseExcelJsWorkbook(workbook) {
  if (!workbook) return;
  try {
    if (Array.isArray(workbook._worksheets)) workbook._worksheets.length = 0;
    else workbook._worksheets = [];
    if (Array.isArray(workbook._media)) workbook._media.length = 0;
    else workbook._media = [];
    if (Array.isArray(workbook.worksheets)) workbook.worksheets.length = 0;
    if (Array.isArray(workbook.media)) workbook.media.length = 0;
    if (Array.isArray(workbook.views)) workbook.views.length = 0;
    if (workbook._definedNames) workbook._definedNames = null;
    if (workbook._themes) workbook._themes = null;
  } catch (error) {
    console.warn("Failed to release temporary ExcelJS workbook references.", error);
  }
}



function prepareExcelJsWorkbookForFastImport(workbook) {
  const xlsx = workbook?.xlsx;
  if (!xlsx) return;
  // Skip media and DrawingML, but retain VML metadata for worksheet comments.
  xlsx._processMediaEntry = async () => {};
  xlsx._processDrawingEntry = async () => {};
  xlsx._processDrawingRelsEntry = async () => {};
}



async function createExcelJsWorkbookForImport(buffer, file, excelJsLoad) {
  let workbook = new ExcelJS.Workbook();
  prepareExcelJsWorkbookForFastImport(workbook);
  try {
    await workbook.xlsx.load(excelJsLoad.buffer, { ignoreNodes: EXCELJS_IMPORT_IGNORE_NODES });
    return { workbook, excelJsLoad };
  } catch (error) {
    if (excelJsLoad.convertedFromLegacyXls) throw error;
    console.warn("Fast ExcelJS workbook load failed; falling back to prepared workbook buffer.", error);
    releaseExcelJsWorkbook(workbook);
  }

  const fallbackLoad = await workbookBufferForExcelJs(buffer, file, { forceZipPreparation: true });
  workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fallbackLoad.buffer, { ignoreNodes: EXCELJS_IMPORT_IGNORE_NODES });
  return { workbook, excelJsLoad: fallbackLoad };
}



async function loadExcelFile(file, options = {}) {
  const loadVersion = ++state.workbookLoadVersion;
  markPendingFileLoadRecovery(file);
  const previousWorkbookState = captureWorkbookRecoveryState();
  revokeWorksheetImageRenderObjectUrls();
  const sourceFileSignature = workbookFileSignature(file);
  closeHeaderSizeDialog({ restoreFocus: false });
  closeNameDialog();
  closeLinkDialog();
  closeObjectLinkDialog();
  closeTranslationPane();
  closeShapeFormatPane();
  closeShapeSelectionPane({ restoreFocus: false });
  closeSymbolDialog({ restoreFocus: false });
  clearQueryRefreshTimers();
  setStatus(`${file.name} を読み込み中...`);
  showAppAlert(`${file.name} を読み込み中...`, "info");
  clearWarnings();
  $exportHtmlButton.prop("disabled", true);
  $toggleFormulasButton.prop("disabled", true);

  try {
    debugWorkbookLoadTrace("arrayBuffer:start");
    const buffer = await file.arrayBuffer();
    debugWorkbookLoadTrace("arrayBuffer:done");
    debugWorkbookLoadTrace("workbook-modules:start");
    await loadWorkbookIoModules({ legacyXls: isLegacyXlsWorkbook(file, buffer) });
    debugWorkbookLoadTrace("workbook-modules:done");
    debugWorkbookLoadTrace("exceljs-buffer:start");
    let excelJsLoad = await workbookBufferForExcelJs(buffer, file);
    debugWorkbookLoadTrace("exceljs-buffer:done");
    let workbook = null;
    debugWorkbookLoadTrace("exceljs-load:start");
    ({ workbook, excelJsLoad } = await createExcelJsWorkbookForImport(buffer, file, excelJsLoad));
    debugWorkbookLoadTrace("exceljs-load:done");
    const metadataBuffer = excelJsLoad.convertedFromLegacyXls ? excelJsLoad.buffer : buffer;
    const themeColors = parseThemeColors(workbook);
    debugWorkbookLoadTrace("metadata-extract:start");
    const [
      workbookPhonetics,
      workbookDrawingShapes,
      workbookDefinedNames,
      workbookAutoFilters,
      workbookDataValidationDetails,
      workbookSheetViewDetails,
      workbookOutlineDetails,
      workbookWindowViews,
      workbookCustomViews,
      workbookMetadata,
      workbookCellHyperlinks,
    ] = await Promise.all([
      extractWorkbookPhonetics(metadataBuffer),
      extractWorkbookDrawingShapes(metadataBuffer, themeColors),
      extractWorkbookDefinedNames(metadataBuffer),
      extractWorkbookAutoFilters(metadataBuffer, themeColors),
      extractWorkbookDataValidationDetails(metadataBuffer),
      extractWorkbookSheetViewDetails(metadataBuffer),
      extractWorkbookOutlineDetails(metadataBuffer),
      extractWorkbookWindowViews(metadataBuffer),
      extractWorkbookCustomViews(metadataBuffer),
      extractWorkbookMetadata(metadataBuffer),
      extractWorkbookCellHyperlinks(metadataBuffer),
    ]);
    debugWorkbookLoadTrace("metadata-extract:done");
    debugWorkbookLoadTrace("ooxml-preserve:start");
    const preservedOoxmlPackage = excelJsLoad.convertedFromLegacyXls ? null : await extractPreservedOoxmlPackage(buffer, file.name);
    debugWorkbookLoadTrace("ooxml-preserve:done");

    if (loadVersion !== state.workbookLoadVersion) return;

    resetCellStyleInternPools();
    debugWorkbookLoadTrace("workbook-to-model:start");
    state.model = workbookToModel(
      workbook,
      file.name,
      workbookPhonetics,
      workbookDrawingShapes,
      themeColors,
      workbookDefinedNames,
      workbookAutoFilters,
      workbookDataValidationDetails,
      workbookSheetViewDetails,
      workbookOutlineDetails,
      workbookWindowViews,
      workbookCustomViews,
      workbookMetadata,
    );
    applyWorkbookCellHyperlinksToModel(state.model, workbookCellHyperlinks);
    if (excelJsLoad.convertedFromLegacyXls) {
      debugWorkbookLoadTrace("legacy-xls-styles:start");
      applyLegacyXlsCellStylesToModel(state.model, buffer);
      debugWorkbookLoadTrace("legacy-xls-styles:done");
      debugWorkbookLoadTrace("legacy-xls-style-ranges:start");
      compactWorkbookStyleOnlyCellsToStyleRanges(state.model);
      debugWorkbookLoadTrace("legacy-xls-style-ranges:done");
    }
    debugWorkbookLoadTrace("workbook-to-model:done");
    releaseExcelJsWorkbook(workbook);
    workbook = null;
    state.model.sourceFileSignature = sourceFileSignature;
    state.openedFileSignature = sourceFileSignature;
    state.model.saveType = defaultSaveTypeForLoadedFile(file.name);
    state.currentSaveType = state.model.saveType;
    resetCurrentSaveTarget();
    applyOpenedWorkbookSaveTarget(file, options.fileHandle);
    state.model.ooxmlPackage = preservedOoxmlPackage;
    debugWorkbookLoadTrace("formula-engine:start");
    state.hf = buildFormulaEngine(state.model);
    debugWorkbookLoadTrace("formula-engine:done");
    primeWorkbookFormulaSpillAnchorCache(state.model);
    primeWorkbookUsedRangeCache(state.model);
    reapplyImportedAutoFiltersOnLoad();
    refreshWorkbookQueriesOnOpen();
    state.activeSheetIndex = clampSheetIndex(state.model.activeSheetIndex);
    state.showFormulas = activeSheet()?.showFormulas === true;
    state.selectedSheetIndexes = workbookTabSelectedSheetIndexes(state.model, state.activeSheetIndex);
    state.sheetContextMenu = null;
    state.cellContextMenu = null;
    state.headerSizeDialog = null;
    state.nameDialog = null;
    state.linkDialog = null;
    state.translationPane = null;
    state.symbolDialog = null;
    state.annotationEditor = null;
    state.shapeSelectionPane = null;
    state.shapeSelectionPaneDrag = null;
    state.undoStack = [];
    state.redoStack = [];
    state.workbookRevision = 0;
    markWorkbookSavedClean();
    clearWorkbookClipboardState();
    state.selected = null;
    state.selectedImage = null;
    state.selectedImages = [];
    state.selectionRange = null;
    state.selectionRanges = [];
    state.selectionAnchor = null;
    state.selectionStart = null;
    state.selectionDrag = null;
    resetLazyScrollLocks();
    state.formulaEdit = null;
    state.isSelecting = false;
    restoreWorkbookWindowSelection(state.model.mainWindowView || state.model.workbookView || {});
    document.title = state.model.sourceName;

    showWorkbookWorkspace();
    $exportHtmlButton.prop("disabled", false);
    $toggleFormulasButton.prop("disabled", false);
    $toggleFormulasButton
      .attr("title", state.showFormulas ? "計算結果を表示" : "数式を表示")
      .html(`<i data-lucide="${state.showFormulas ? "eye-off" : "eye"}"></i>`);
    $formulaInput.prop("disabled", true).val("");
    $formulaExpandButton.prop("disabled", false);
    state.formulaBarExpandedLines = FORMULA_BAR_DEFAULT_EXPANDED_LINES;
    setFormulaBarLines(FORMULA_BAR_MIN_LINES, { resetExpandedLines: true });

    $workbookName.text(state.model.sourceName);
    debugWorkbookLoadTrace("render:start");
    renderWorkbook();
    debugWorkbookLoadTrace("render:done");
    if (state.selected) updateFormulaBarForSelection();
    scheduleQueryRefreshTimers();
    clearAppAlert();
    setStatus(`${state.model.sheets.length} 個のワークシートを読み込みました。セルは直接編集でき、数式は自動再計算されます。`);
    debugWorkbookLoadTrace("persistence:start");
    await flushWorkbookPersistence();
    debugWorkbookLoadTrace("persistence:done");
    clearPendingFileLoadRecoveryMarker();
  } catch (error) {
    console.error(error);
    if (loadVersion !== state.workbookLoadVersion) return;
    const message = readableLoadError(error);
    recoverWorkbookAfterFailedFileLoad(previousWorkbookState);
    setStatus(previousWorkbookState?.snapshot?.model
      ? "読み込みに失敗しました。元のブックに戻しました。"
      : "読み込みに失敗しました。新しいブックを作成しました。");
    showAppAlert(`読み込みに失敗しました: ${message}`, "error");
    clearPendingFileLoadRecoveryMarker();
  }
}



function captureWorkbookRecoveryState() {
  if (!state.model) return null;
  let snapshot = null;
  try {
    snapshot = currentWorkbookSnapshot();
  } catch (error) {
    console.warn("Failed to capture workbook state before loading a new file.", error);
    return null;
  }
  return {
    snapshot,
    workbookRevision: Math.max(0, Number(state.workbookRevision) || 0),
    workbookSavedRevision: Math.max(0, Number(state.workbookSavedRevision) || 0),
    currentSaveType: state.currentSaveType,
    currentSaveFileHandle: state.currentSaveFileHandle,
    currentSaveDirectoryHandle: state.currentSaveDirectoryHandle,
    currentSaveFileName: state.currentSaveFileName,
    openedFileSignature: normalizeWorkbookFileSignature(state.openedFileSignature),
  };
}



function recoverWorkbookAfterFailedFileLoad(previousWorkbookState = null) {
  cancelWorkbookPersistenceIdleCallback();
  if (state.persistenceTimer) {
    window.clearTimeout(state.persistenceTimer);
    state.persistenceTimer = null;
  }

  if (previousWorkbookState?.snapshot?.model) {
    try {
      applyPersistedWorkbookSnapshot(previousWorkbookState.snapshot);
      state.workbookRevision = previousWorkbookState.workbookRevision;
      state.workbookSavedRevision = previousWorkbookState.workbookSavedRevision;
      state.currentSaveType = previousWorkbookState.currentSaveType || state.currentSaveType;
      state.currentSaveFileHandle = previousWorkbookState.currentSaveFileHandle || null;
      state.currentSaveDirectoryHandle = previousWorkbookState.currentSaveDirectoryHandle || null;
      state.currentSaveFileName = previousWorkbookState.currentSaveFileName || "";
      state.openedFileSignature = previousWorkbookState.openedFileSignature || normalizeWorkbookFileSignature(state.model?.sourceFileSignature);
      return;
    } catch (restoreError) {
      console.warn("Failed to restore workbook after load failure.", restoreError);
    }
  }

  state.model = null;
  state.hf = null;
  state.currentSaveType = "";
  state.openedFileSignature = null;
  resetCellStyleInternPools();
  state.workbookRevision = 0;
  markWorkbookSavedClean();
  clearWorkbookClipboardState();
  openBlankWorkbook({ status: "読み込みに失敗しました。新しいブックを作成しました。" });
  void flushWorkbookPersistence();
}



async function excelJsWorkbookFromFile(file) {
  const buffer = await file.arrayBuffer();
  await loadWorkbookIoModules({ legacyXls: isLegacyXlsWorkbook(file, buffer) });
  const excelJsLoad = await workbookBufferForExcelJs(buffer, file);
  const { workbook } = await createExcelJsWorkbookForImport(buffer, file, excelJsLoad);
  return workbook;
}



async function workbookBufferForExcelJs(buffer, file = null, options = {}) {
  if (isLegacyXlsWorkbook(file, buffer)) {
    await loadXlsxModule();
    return {
      buffer: convertLegacyXlsBufferToXlsxBuffer(buffer, file?.name || "workbook.xls"),
      convertedFromLegacyXls: true,
    };
  }
  if (!options.forceZipPreparation) {
    return { buffer, convertedFromLegacyXls: false };
  }
  try {
    const zip = await JSZip.loadAsync(buffer);
    let changed = false;
    changed = await normalizeSpreadsheetMainNamespacePrefixesForExcelJs(zip) || changed;
    await Promise.all(Object.keys(zip.files)
      .filter((path) => /^xl\/tables\/table\d+\.xml$/i.test(path))
      .map(async (path) => {
        const file = zip.file(path);
        if (!file) return;
        const xml = await file.async("text");
        const nextXml = stripUnsupportedTableAutoFilterXml(xml);
        if (nextXml === xml) return;
        zip.file(path, nextXml);
        changed = true;
      }));
    changed = await stripWorksheetDrawingsForExcelJs(zip) || changed;
    return {
      buffer: changed ? await zip.generateAsync({ type: "arraybuffer" }) : buffer,
      convertedFromLegacyXls: false,
    };
  } catch (error) {
    console.warn("Failed to prepare workbook for ExcelJS load.", error);
    return { buffer, convertedFromLegacyXls: false };
  }
}



async function normalizeSpreadsheetMainNamespacePrefixesForExcelJs(zip) {
  let changed = false;
  await Promise.all(Object.keys(zip.files)
    .filter((path) => /^xl\/.+\.xml$/i.test(path))
    .map(async (path) => {
      const file = zip.file(path);
      if (!file) return;
      const xml = await file.async("text");
      const nextXml = normalizeSpreadsheetMainNamespacePrefixesXml(xml);
      if (nextXml === xml) return;
      zip.file(path, nextXml);
      changed = true;
    }));
  return changed;
}



function normalizeSpreadsheetMainNamespacePrefixesXml(xml) {
  let next = String(xml || "");
  const prefixes = new Set();
  next.replace(/\bxmlns:([A-Za-z_][\w.-]*)=(["'])http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main\2/g, (_match, prefix) => {
    prefixes.add(prefix);
    return _match;
  });
  prefixes.forEach((prefix) => {
    next = next.replace(new RegExp(`(<\\/?)${escapeRegExp(prefix)}:`, "g"), "$1");
  });
  return next;
}



function isLegacyXlsWorkbook(file, buffer = null) {
  const name = String(file?.name || "").toLowerCase();
  if (/\.xls$/i.test(name)) return true;
  if (/\.(xlsx|xlsm)$/i.test(name)) return false;
  const type = String(file?.type || "").toLowerCase();
  if (type === "application/vnd.ms-excel") return !isZipWorkbookBuffer(buffer);
  return isLegacyXlsWorkbookBuffer(buffer);
}



function isZipWorkbookBuffer(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer || []);
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;
}



function isLegacyXlsWorkbookBuffer(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer || []);
  return bytes.length >= 8 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0 &&
    bytes[4] === 0xa1 &&
    bytes[5] === 0xb1 &&
    bytes[6] === 0x1a &&
    bytes[7] === 0xe1;
}



function convertLegacyXlsBufferToXlsxBuffer(buffer, sourceName = "workbook.xls") {
  try {
    const sourceWorkbook = XLSX.read(buffer, {
      type: "array",
      cellDates: false,
      cellFormula: true,
      cellNF: true,
      cellStyles: true,
      sheetStubs: true,
      WTF: false,
    });
    if (!sourceWorkbook?.SheetNames?.length) throw new Error("ワークシートが見つかりません。");
    return normalizeArrayBuffer(XLSX.write(sourceWorkbook, {
      type: "array",
      bookType: "xlsx",
      cellDates: false,
      cellStyles: true,
    }));
  } catch (error) {
    throw new Error(`旧形式の .xls ブックを読み込めませんでした（${sourceName}）。${error?.message || error}`);
  }
}



function applyLegacyXlsCellStylesToModel(model, buffer) {
  const sheetStyles = extractLegacyXlsCellStyles(buffer);
  if (!model?.sheets?.length || !sheetStyles?.length) return 0;
  let applied = 0;
  sheetStyles.forEach((stylesByCell, sheetIndex) => {
    const sheet = model.sheets[sheetIndex];
    if (!sheet || !stylesByCell?.size) return;
    const rowLimit = Math.max(1, Number(sheet.rowCount) || 1);
    const colLimit = Math.max(1, Number(sheet.colCount) || 1);
    stylesByCell.forEach(({ row, col, style }) => {
      if (row < 1 || col < 1 || row > rowLimit || col > colLimit || !legacyXlsCellStyleHasVisibleFormat(style)) return;
      const key = cellKey(row, col);
      const existingCell = sheet.cells[key];
      const hasVisibleEmptyStyle = legacyXlsCellStyleHasVisibleEmptyCellFormat(style);
      if (!existingCell && !hasVisibleEmptyStyle) return;
      const cell = applyLegacyXlsCellStyleToCell(
        existingCell ? { ...existingCell } : legacyXlsStyleOnlyCell(row, col),
        existingCell ? style : legacyXlsEmptyCellVisibleStyle(style),
      );
      sheet.cells[key] = internCellModelStylesForSheet(sheet, cell);
      applied += 1;
    });
    applied += applyLegacyXlsMergedEdgeBorders(sheet);
  });
  return applied;
}



function legacyXlsStyleOnlyCell(row, col) {
  return {
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
    sparkline: null,
    richText: null,
    css: {},
    hasFill: false,
    hasFontColor: false,
    borders: {},
    numFmt: "",
  };
}



function applyLegacyXlsCellStyleToCell(cell, style = {}) {
  if (style.numFmt) {
    cell.numFmt = style.numFmt;
    const displaySource = cell.formula ? cell.cached : cell.raw ?? cell.cached;
    if (displaySource != null) cell.display = displayValue(displaySource, cell.numFmt);
  }
  if (Object.keys(style.css || {}).length) {
    cell.css = { ...(cell.css || {}), ...style.css };
    if (style.css.backgroundColor) cell.hasFill = true;
    if (style.css.color) cell.hasFontColor = true;
  }
  if (Object.keys(style.borders || {}).length) {
    cell.borders = mergeLegacyXlsCellBorders(cell.borders, style.borders);
  }
  return cell;
}



function legacyXlsCellStyleHasVisibleFormat(style = {}) {
  return Boolean(
    style.numFmt ||
      Object.keys(style.css || {}).length ||
      Object.keys(style.borders || {}).length
  );
}



function legacyXlsCellStyleHasVisibleEmptyCellFormat(style = {}) {
  return Boolean(
    Object.keys(style.borders || {}).length ||
      Boolean(style.css?.backgroundColor)
  );
}



function legacyXlsEmptyCellVisibleStyle(style = {}) {
  const css = {};
  if (style.css?.backgroundColor) css.backgroundColor = style.css.backgroundColor;
  return {
    ...(Object.keys(css).length ? { css } : {}),
    ...(Object.keys(style.borders || {}).length ? { borders: style.borders } : {}),
  };
}



function mergeLegacyXlsCellBorders(existing = {}, imported = {}) {
  return {
    ...(existing || {}),
    ...(imported || {}),
  };
}



function extractLegacyXlsCellStyles(buffer) {
  try {
    const cfb = XLSX.CFB?.read?.(new Uint8Array(buffer || []), { type: "buffer" });
    const workbookEntry = XLSX.CFB?.find?.(cfb, "/Workbook") || XLSX.CFB?.find?.(cfb, "/Book");
    const bytes = workbookEntry?.content ? new Uint8Array(workbookEntry.content) : null;
    if (!bytes?.length) return [];

    let biffVersion = 8;
    const formats = { ...LEGACY_XLS_BUILTIN_NUM_FMTS };
    const fonts = [];
    const xfs = [];
    const xfStyleCache = [];
    const palette = {};
    const sheetStyles = [];
    let sheetIndex = -1;
    let worksheetDepth = 0;

    for (let offset = 0; offset + 4 <= bytes.length;) {
      const recordType = legacyXlsReadUInt16(bytes, offset);
      const length = legacyXlsReadUInt16(bytes, offset + 2);
      const dataOffset = offset + 4;
      const nextOffset = dataOffset + length;
      if (nextOffset > bytes.length) break;

      if (recordType === 0x0809) {
        biffVersion = legacyXlsBiffVersion(bytes, dataOffset, length) || biffVersion;
        const substreamType = length >= 4 ? legacyXlsReadUInt16(bytes, dataOffset + 2) : 0;
        if (worksheetDepth > 0) {
          worksheetDepth += 1;
        } else if (substreamType === 0x0010) {
          worksheetDepth = 1;
          sheetIndex += 1;
          sheetStyles[sheetIndex] = sheetStyles[sheetIndex] || new Map();
        }
      } else if (recordType === 0x000a) {
        if (worksheetDepth > 0) worksheetDepth = Math.max(0, worksheetDepth - 1);
      } else if (worksheetDepth === 0 && recordType === 0x041e) {
        legacyXlsParseFormat(bytes, dataOffset, length, formats, biffVersion);
      } else if (worksheetDepth === 0 && recordType === 0x0031) {
        fonts.push(legacyXlsParseFont(bytes, dataOffset, length, biffVersion));
      } else if (worksheetDepth === 0 && recordType === 0x00e0) {
        xfs.push(legacyXlsParseXf(bytes, dataOffset, length));
      } else if (worksheetDepth === 0 && recordType === 0x0092) {
        legacyXlsParsePalette(bytes, dataOffset, length, palette);
      } else if (worksheetDepth > 0) {
        legacyXlsCollectCellStyles(bytes, dataOffset, length, recordType, sheetStyles[sheetIndex], xfs, xfStyleCache, fonts, formats, palette);
      }

      offset = nextOffset;
    }

    return sheetStyles;
  } catch (error) {
    console.warn("Failed to recover legacy XLS cell styles.", error);
    return [];
  }
}



function legacyXlsCollectCellStyles(bytes, dataOffset, length, recordType, stylesByCell, xfs, xfStyleCache, fonts, formats, palette) {
  if (!stylesByCell) return;
  const addCell = (rowIndex, colIndex, xfIndex) => {
    const style = legacyXlsCachedXfStyleToModel(xfs, xfStyleCache, xfIndex, fonts, formats, palette);
    if (!legacyXlsCellStyleHasVisibleFormat(style)) return;
    const row = rowIndex + 1;
    const col = colIndex + 1;
    stylesByCell.set(cellKey(row, col), { row, col, style });
  };

  if (LEGACY_XLS_CELL_RECORDS_WITH_XF.has(recordType) && length >= 6) {
    addCell(
      legacyXlsReadUInt16(bytes, dataOffset),
      legacyXlsReadUInt16(bytes, dataOffset + 2),
      legacyXlsReadUInt16(bytes, dataOffset + 4),
    );
    return;
  }

  if (recordType === 0x00bd && length >= 8) {
    const rowIndex = legacyXlsReadUInt16(bytes, dataOffset);
    const firstCol = legacyXlsReadUInt16(bytes, dataOffset + 2);
    const lastCol = legacyXlsReadUInt16(bytes, dataOffset + length - 2);
    let cursor = dataOffset + 4;
    for (let colIndex = firstCol; colIndex <= lastCol && cursor + 6 <= dataOffset + length - 2; colIndex += 1, cursor += 6) {
      addCell(rowIndex, colIndex, legacyXlsReadUInt16(bytes, cursor));
    }
    return;
  }

  if (recordType === 0x00be && length >= 8) {
    const rowIndex = legacyXlsReadUInt16(bytes, dataOffset);
    const firstCol = legacyXlsReadUInt16(bytes, dataOffset + 2);
    const lastCol = legacyXlsReadUInt16(bytes, dataOffset + length - 2);
    let cursor = dataOffset + 4;
    for (let colIndex = firstCol; colIndex <= lastCol && cursor + 2 <= dataOffset + length - 2; colIndex += 1, cursor += 2) {
      addCell(rowIndex, colIndex, legacyXlsReadUInt16(bytes, cursor));
    }
  }
}



function legacyXlsCachedXfStyleToModel(xfs, cache, xfIndex, fonts, formats, palette) {
  if (!Number.isInteger(xfIndex) || xfIndex < 0) return {};
  if (cache[xfIndex]) return cache[xfIndex];
  cache[xfIndex] = legacyXlsXfStyleToModel(xfs[xfIndex], fonts, formats, palette);
  return cache[xfIndex];
}



function legacyXlsParseXf(bytes, dataOffset, length) {
  if (length < 20) return null;
  const fontIndex = legacyXlsReadUInt16(bytes, dataOffset);
  const numFmtId = legacyXlsReadUInt16(bytes, dataOffset + 2);
  const alignment = bytes[dataOffset + 6] || 0;
  const borderBits = legacyXlsReadUInt32(bytes, dataOffset + 10);
  const colorBits = legacyXlsReadUInt32(bytes, dataOffset + 14);
  const fillBits = legacyXlsReadUInt16(bytes, dataOffset + 18);
  return {
    fontIndex,
    numFmtId,
    horizontalAlign: alignment & 0x07,
    wrapText: Boolean(alignment & 0x08),
    verticalAlign: (alignment >>> 4) & 0x07,
    left: borderBits & 0x0f,
    right: (borderBits >>> 4) & 0x0f,
    top: (borderBits >>> 8) & 0x0f,
    bottom: (borderBits >>> 12) & 0x0f,
    leftColor: (borderBits >>> 16) & 0x7f,
    rightColor: (borderBits >>> 23) & 0x7f,
    topColor: colorBits & 0x7f,
    bottomColor: (colorBits >>> 7) & 0x7f,
    fillPattern: colorBits >>> 26,
    fillForeground: fillBits & 0x7f,
    fillBackground: (fillBits >>> 7) & 0x7f,
  };
}



function legacyXlsParseFormat(bytes, dataOffset, length, formats, biffVersion = 8) {
  if (length < 3) return;
  const id = legacyXlsReadUInt16(bytes, dataOffset);
  const value = legacyXlsReadUnicodeString2(bytes, dataOffset + 2, dataOffset + length, biffVersion);
  if (value) formats[id] = value;
}



function legacyXlsParseFont(bytes, dataOffset, length, biffVersion = 8) {
  if (length < 6) return null;
  const endOffset = dataOffset + length;
  const skip = biffVersion <= 2 ? 4 : biffVersion <= 4 ? 6 : 14;
  return {
    sizePt: legacyXlsReadUInt16(bytes, dataOffset) / 20,
    flags: legacyXlsReadUInt16(bytes, dataOffset + 2),
    colorIndex: legacyXlsReadUInt16(bytes, dataOffset + 4),
    boldWeight: biffVersion <= 2 ? 400 : legacyXlsReadUInt16(bytes, dataOffset + 6),
    underline: biffVersion <= 2 ? 0 : bytes[dataOffset + 10] || 0,
    name: dataOffset + skip < endOffset
      ? legacyXlsReadShortUnicodeString(bytes, dataOffset + skip, endOffset, biffVersion)
      : "",
  };
}



function legacyXlsParsePalette(bytes, dataOffset, length, palette) {
  if (length < 2) return;
  const count = legacyXlsReadUInt16(bytes, dataOffset);
  let cursor = dataOffset + 2;
  for (let index = 0; index < count && cursor + 4 <= dataOffset + length; index += 1, cursor += 4) {
    palette[index + 8] = rgbToHex([bytes[cursor], bytes[cursor + 1], bytes[cursor + 2]]);
  }
}



function legacyXlsXfStyleToModel(xf, fonts, formats, palette) {
  if (!xf) return {};
  const css = legacyXlsXfCssToModel(xf, fonts, palette);
  const numFmt = legacyXlsNumberFormat(xf.numFmtId, formats);
  const borders = legacyXlsXfBordersToModel(xf, palette);
  return {
    ...(numFmt ? { numFmt } : {}),
    ...(Object.keys(css).length ? { css } : {}),
    ...(Object.keys(borders).length ? { borders } : {}),
  };
}



function legacyXlsXfCssToModel(xf, fonts, palette) {
  const css = {};
  Object.assign(css, legacyXlsAlignmentCssToModel(xf));
  const font = legacyXlsFontForIndex(fonts, xf.fontIndex);
  if (font) {
    if (Number.isFinite(font.sizePt) && font.sizePt > 0) css.fontSize = legacyXlsFontSizeCss(font.sizePt);
    if (font.name) css.fontFamily = fontFamilyForExcel(font.name);
    const color = legacyXlsFontColorToCss(font.colorIndex, palette);
    if (color) css.color = color;
    // The BIFF bold weight is not reliable enough to overlay after the .xls
    // conversion bridge. Preserve the rest of the font styling, but leave
    // boldness to formats that the main workbook reader can import directly.
    if (font.flags & 0x02) css.fontStyle = "italic";
    if (font.underline) css.textDecoration = String(font.underline) === "2" || String(font.underline) === "34" ? "underline double" : "underline";
  }
  const fillColor = legacyXlsFillColorToCss(xf, palette);
  if (fillColor) css.backgroundColor = fillColor;
  return css;
}



function legacyXlsAlignmentCssToModel(xf) {
  const css = {};
  const horizontal = legacyXlsHorizontalAlign(xf?.horizontalAlign);
  if (horizontal) css.textAlign = horizontal;
  const vertical = legacyXlsVerticalAlign(xf?.verticalAlign);
  if (vertical) {
    css.verticalAlign = vertical;
    css.alignItems = verticalAlignToFlex(vertical);
  }
  if (xf?.wrapText) {
    css.whiteSpace = "pre-wrap";
    css.overflowWrap = "anywhere";
  }
  return css;
}



function legacyXlsHorizontalAlign(value) {
  switch (Number(value)) {
    case 1:
      return "left";
    case 2:
    case 6:
      return "center";
    case 3:
      return "right";
    case 5:
    case 7:
      return "justify";
    default:
      return "";
  }
}



function legacyXlsVerticalAlign(value) {
  switch (Number(value)) {
    case 0:
      return "top";
    case 1:
      return "middle";
    case 2:
      return "bottom";
    case 3:
    case 4:
      return "middle";
    default:
      return "";
  }
}



function legacyXlsFontForIndex(fonts, fontIndex) {
  const index = Number(fontIndex);
  if (!Number.isInteger(index) || index < 0) return null;
  if (index === 4) return null;
  return (index > 4 ? fonts[index - 1] : fonts[index]) || null;
}



function legacyXlsFontSizeCss(sizePt) {
  const rounded = Math.round(Number(sizePt) * 100) / 100;
  return `${Number.isInteger(rounded) ? rounded : rounded.toFixed(2).replace(/0+$/, "").replace(/\.$/, "")}pt`;
}



function legacyXlsFontColorToCss(index, palette = {}) {
  const normalizedIndex = Number(index);
  if (!Number.isInteger(normalizedIndex) || normalizedIndex === 0x7fff) return "";
  const color = legacyXlsIndexedColorToCss(normalizedIndex, palette);
  return color === "#000000" ? "" : color;
}



function legacyXlsFillColorToCss(xf, palette = {}) {
  if (!xf || LEGACY_XLS_FILL_PATTERN_NONE.has(Number(xf.fillPattern) || 0)) return "";
  return legacyXlsIndexedColorToCss(xf.fillForeground, palette);
}



function legacyXlsNumberFormat(numFmtId, formats = {}) {
  const id = Number(numFmtId);
  if (!Number.isInteger(id)) return "";
  const format = formats[id] || "";
  return format && format !== "General" ? format : "";
}



function legacyXlsXfBordersToModel(xf, palette) {
  if (!xf) return {};
  const borders = {};
  legacyXlsAddBorderSide(borders, "top", xf.top, xf.topColor, palette);
  legacyXlsAddBorderSide(borders, "right", xf.right, xf.rightColor, palette);
  legacyXlsAddBorderSide(borders, "bottom", xf.bottom, xf.bottomColor, palette);
  legacyXlsAddBorderSide(borders, "left", xf.left, xf.leftColor, palette);
  return borders;
}



function legacyXlsAddBorderSide(target, side, lineStyle, colorIndex, palette) {
  const excelStyle = LEGACY_XLS_BORDER_STYLE_MAP[lineStyle];
  if (!excelStyle) return;
  const { width, style } = borderStyleToCss(excelStyle);
  target[side] = {
    width,
    style,
    color: legacyXlsIndexedColorToCss(colorIndex, palette),
    weight: borderWeight(excelStyle),
  };
}



function legacyXlsIndexedColorToCss(index, palette = {}) {
  const normalizedIndex = Number(index);
  return normalizeColorValue(palette[normalizedIndex] || INDEXED_COLORS[normalizedIndex] || "#000000") || "#000000";
}



function applyLegacyXlsMergedEdgeBorders(sheet) {
  let applied = 0;
  (sheet?.merges || []).forEach((merge) => {
    const borders = legacyXlsMergedEdgeBorders(sheet, merge);
    if (!Object.keys(borders).length) return;
    const key = cellKey(merge.top, merge.left);
    const cell = sheet.cells[key] ? { ...sheet.cells[key] } : legacyXlsStyleOnlyCell(merge.top, merge.left);
    const nextBorders = mergeLegacyXlsCellBorders(cell.borders, borders);
    if (stableObjectKey(nextBorders) === stableObjectKey(cell.borders || {})) return;
    cell.borders = nextBorders;
    sheet.cells[key] = internCellModelStylesForSheet(sheet, cell);
    applied += 1;
  });
  return applied;
}



function legacyXlsMergedEdgeBorders(sheet, merge) {
  const borders = {};
  const top = legacyXlsMergedEdgeBorder(sheet, merge, "top");
  const right = legacyXlsMergedEdgeBorder(sheet, merge, "right");
  const bottom = legacyXlsMergedEdgeBorder(sheet, merge, "bottom");
  const left = legacyXlsMergedEdgeBorder(sheet, merge, "left");
  if (top) borders.top = top;
  if (right) borders.right = right;
  if (bottom) borders.bottom = bottom;
  if (left) borders.left = left;
  return borders;
}



function legacyXlsMergedEdgeBorder(sheet, merge, side) {
  const candidates = [];
  if (side === "top" || side === "bottom") {
    const row = side === "top" ? merge.top : merge.bottom;
    for (let col = merge.left; col <= merge.right; col += 1) {
      const border = sheet.cells[cellKey(row, col)]?.borders?.[side];
      if (border) candidates.push(border);
    }
  } else {
    const col = side === "left" ? merge.left : merge.right;
    for (let row = merge.top; row <= merge.bottom; row += 1) {
      const border = sheet.cells[cellKey(row, col)]?.borders?.[side];
      if (border) candidates.push(border);
    }
  }
  return candidates.sort(compareLegacyXlsBorderDescending)[0] || null;
}



function compareLegacyXlsBorderDescending(a, b) {
  return compareBorder(b, a);
}



function legacyXlsBiffVersion(bytes, dataOffset, length) {
  if (length < 2) return null;
  return {
    0x0002: 2,
    0x0007: 2,
    0x0200: 2,
    0x0300: 3,
    0x0400: 4,
    0x0500: 5,
    0x0600: 8,
  }[legacyXlsReadUInt16(bytes, dataOffset)] || null;
}



function legacyXlsReadUnicodeString2(bytes, offset, endOffset, biffVersion = 8) {
  if (offset >= endOffset) return "";
  if (biffVersion <= 5) {
    const length = bytes[offset] || 0;
    return legacyXlsReadSingleByteString(bytes, offset + 1, Math.min(endOffset, offset + 1 + length));
  }
  const length = legacyXlsReadUInt16(bytes, offset);
  return legacyXlsReadUnicodeStringNoLength(bytes, offset + 2, endOffset, length);
}



function legacyXlsReadShortUnicodeString(bytes, offset, endOffset, biffVersion = 8) {
  if (offset >= endOffset) return "";
  const length = bytes[offset] || 0;
  if (biffVersion <= 5) return legacyXlsReadSingleByteString(bytes, offset + 1, Math.min(endOffset, offset + 1 + length));
  return legacyXlsReadUnicodeStringNoLength(bytes, offset + 1, endOffset, length);
}



function legacyXlsReadUnicodeStringNoLength(bytes, offset, endOffset, charCount) {
  if (offset >= endOffset || !charCount) return "";
  let cursor = offset;
  const flags = bytes[cursor++] || 0;
  const isWide = Boolean(flags & 0x01);
  if (flags & 0x08) cursor += 2;
  if (flags & 0x04) cursor += 4;
  if (isWide) {
    let text = "";
    for (let index = 0; index < charCount && cursor + 1 < endOffset; index += 1, cursor += 2) {
      text += String.fromCharCode(legacyXlsReadUInt16(bytes, cursor));
    }
    return text;
  }
  return legacyXlsReadSingleByteString(bytes, cursor, Math.min(endOffset, cursor + charCount));
}



function legacyXlsReadSingleByteString(bytes, startOffset, endOffset) {
  let text = "";
  for (let offset = startOffset; offset < endOffset; offset += 1) {
    text += String.fromCharCode(bytes[offset] || 0);
  }
  return text;
}



function legacyXlsReadUInt16(bytes, offset) {
  return (bytes[offset] || 0) | ((bytes[offset + 1] || 0) << 8);
}



function legacyXlsReadUInt32(bytes, offset) {
  return (
    (bytes[offset] || 0) |
    ((bytes[offset + 1] || 0) << 8) |
    ((bytes[offset + 2] || 0) << 16) |
    ((bytes[offset + 3] || 0) << 24)
  ) >>> 0;
}



function normalizeArrayBuffer(value) {
  if (value instanceof ArrayBuffer) return value;
  if (ArrayBuffer.isView(value)) {
    return value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength);
  }
  return new Uint8Array(value || []).buffer;
}



function stripUnsupportedTableAutoFilterXml(xml) {
  return String(xml || "")
    .replace(/<sortState\b[\s\S]*?<\/sortState>|<sortState\b[^>]*\/>/g, "")
    .replace(/<colorFilter\b[\s\S]*?<\/colorFilter>|<colorFilter\b[^>]*\/>/g, "");
}



async function stripWorksheetDrawingsForExcelJs(zip) {
  let changed = false;
  await Promise.all(Object.keys(zip.files)
    .filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path))
    .map(async (path) => {
      const file = zip.file(path);
      if (!file) return;
      const xml = await file.async("text");
      const nextXml = stripWorksheetDrawingTags(xml);
      if (nextXml === xml) return;
      zip.file(path, nextXml);
      changed = true;
    }));
  await Promise.all(Object.keys(zip.files)
    .filter((path) => /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/i.test(path))
    .map(async (path) => {
      const file = zip.file(path);
      if (!file) return;
      const xml = await file.async("text");
      const nextXml = stripDrawingRelationshipTags(xml);
      if (nextXml === xml) return;
      zip.file(path, nextXml);
      changed = true;
    }));
  Object.keys(zip.files)
    // Keep VML note parts and their relationships for comment reconciliation.
    .filter((path) => /^xl\/drawings\/.+\.xml(?:\.rels)?$/i.test(path))
    .forEach((path) => {
      zip.remove(path);
      changed = true;
    });
  const contentTypes = zip.file("[Content_Types].xml");
  if (contentTypes) {
    const xml = await contentTypes.async("text");
    const nextXml = stripDrawingContentTypeOverrides(xml);
    if (nextXml !== xml) {
      zip.file("[Content_Types].xml", nextXml);
      changed = true;
    }
  }
  return changed;
}



function stripWorksheetDrawingTags(xml) {
  return String(xml || "").replace(/<(?:[A-Za-z_][\w.-]*:)?drawing\b[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?drawing>|<(?:[A-Za-z_][\w.-]*:)?drawing\b[^>]*\/>/g, "");
}



function stripDrawingRelationshipTags(xml) {
  return String(xml || "").replace(/<Relationship\b(?=[^>]*\bType=(["'])[^"']*\/drawing\1)[^>]*\/>/g, "");
}



function stripDrawingContentTypeOverrides(xml) {
  return String(xml || "").replace(/<Override\b(?=[^>]*\bPartName=(["'])\/xl\/drawings\/[^"']+\.xml\1)[^>]*\/>/g, "");
}



async function restorePersistedWorkbook() {
  if (!persistenceAvailable()) return false;
  const restoreVersion = state.workbookLoadVersion;
  try {
    const snapshot = await readPersistedWorkbookSnapshot();
    if (!snapshot?.model || restoreVersion !== state.workbookLoadVersion || state.model) return false;
    state.persistenceLoading = true;
    applyPersistedWorkbookSnapshot(snapshot);
    clearAppAlert();
    setStatus(`${state.model.sourceName} を復元しました。ページを再読み込みしても内容は保持されます。`);
    return true;
  } catch (error) {
    console.warn("Failed to restore workbook snapshot.", error);
    state.model = null;
    state.hf = null;
    state.currentSaveType = "";
    state.openedFileSignature = null;
    resetCellStyleInternPools();
    try {
      await deletePersistedWorkbookSnapshot();
    } catch (deleteError) {
      console.warn("Failed to delete broken workbook snapshot.", deleteError);
    }
    return false;
  } finally {
    state.persistenceLoading = false;
  }
}



function createBlankWorkbookModel(sourceName = "Book1") {
  const createdAt = new Date().toISOString();
  return {
    version: WORKBOOK_MODEL_VERSION,
    sourceName,
    saveType: "html",
    createdAt,
    cwsAgentDocumentId: createWorkbookDocumentId(),
    sheets: [createBlankSheet("Sheet1")],
    activeSheetIndex: 0,
    showFormulaBar: true,
    definedNames: [],
    relationships: [],
    macros: [],
    workbookProperties: {},
    calcProperties: {},
    workbookView: normalizeWorkbookViewModel({ activeSheetIndex: 0 }, 1),
    mainWindowView: normalizeWorkbookWindowView({ id: "main", name: sourceName, sheetIndex: 0 }),
    activeWindowId: "main",
  };
}



function markWorkbookChanged() {
  state.workbookRevision = Math.max(0, Number(state.workbookRevision) || 0) + 1;
}



function markWorkbookSavedClean() {
  state.workbookSavedRevision = Math.max(0, Number(state.workbookRevision) || 0);
}



function workbookHasUnsavedChanges() {
  return Boolean(state.model && Math.max(0, Number(state.workbookRevision) || 0) !== Math.max(0, Number(state.workbookSavedRevision) || 0));
}



function workbookShouldWarnBeforeUnload() {
  return Boolean(workbookHasUnsavedChanges() || workbookHasPendingEdit());
}



function workbookHasPendingEdit() {
  if (!state.model) return false;
  if (state.formulaEdit) {
    return String($formulaInput.val() ?? "") !== String(state.formulaEdit.before ?? "");
  }
  const cell = document.querySelector(".sheet-cell[contenteditable='true']");
  if (!cell) return false;
  const $cell = $(cell);
  if ($cell.data("cancelEdit")) return false;
  const before = String($cell.data("beforeEdit") ?? "");
  return editablePlainText(cell) !== before;
}



function handleWorkbookBeforeUnload(event) {
  if (state.suppressBeforeUnloadWarning) return;
  if (!workbookShouldWarnBeforeUnload()) return;
  event.preventDefault();
  event.returnValue = "";
  return "";
}



function openBlankWorkbook(options = {}) {
  applyPersistedWorkbookSnapshot({
    model: createBlankWorkbookModel(options.sourceName || "Book1"),
    activeSheetIndex: 0,
    showFormulas: false,
    formulaBarLines: FORMULA_BAR_MIN_LINES,
    formulaBarExpandedLines: FORMULA_BAR_DEFAULT_EXPANDED_LINES,
    selection: { sheetIndex: 0, row: 1, col: 1 },
    selectionRange: { sheetIndex: 0, top: 1, left: 1, bottom: 1, right: 1 },
    selectionRanges: [{ sheetIndex: 0, top: 1, left: 1, bottom: 1, right: 1 }],
  });
  clearAppAlert();
  setStatus(options.status || "新しいブックを作成しました。");
}



async function loadStandaloneWorkbookModel(model, options = {}) {
  const loadVersion = ++state.workbookLoadVersion;
  // The format being imported does not change the page's hosting mode.
  // Hosted pages must persist their active imported book in the slot that the
  // same URL restores on reload; real standalone files retain edition isolation.
  state.isStandalone = options.standalone !== false;
  state.persistenceLoading = true;
  state.standalonePersistenceScope = null;
  try {
    const workbookModel = extractStandaloneWorkbookPayload(model);
    if (!workbookModel) {
      throw new Error("HTML 内のワークブック データにワークシート情報がありません。");
    }
    let scope = null;
    let saved = null;
    try {
      if(state.isStandalone)scope = await standaloneWorkbookPersistenceScope(workbookModel);
      if (scope) saved = await readPersistedWorkbookSnapshot(scope.key);
    } catch (error) {
      // Opening a file remains possible when browser storage is blocked. Agent
      // completion still requires a successful durable write, never a no-op.
      console.warn("Failed to read standalone workbook checkpoint.", error);
    }
    if (loadVersion !== state.workbookLoadVersion) return false;
    if (scope && !workbookModel.cwsAgentDocumentId) workbookModel.cwsAgentDocumentId = scope.documentId;
    applyPersistedWorkbookSnapshot(saved?.model?.cwsAgentDocumentId === scope?.documentId && scope ? saved : {
      model: workbookModel,
      activeSheetIndex: 0,
      showFormulas: false,
      formulaBarLines: FORMULA_BAR_MIN_LINES,
      formulaBarExpandedLines: FORMULA_BAR_DEFAULT_EXPANDED_LINES,
    });
    state.standalonePersistenceScope = scope;
    if (state.model) {
      state.model.saveType = "html";
      state.currentSaveType = "html";
      state.openedFileSignature = normalizeWorkbookFileSignature(state.model.sourceFileSignature);
      resetCurrentSaveTarget();
      const openedFileName = options.openedFileName === false ? "" : options.openedFileName || standaloneOpenedFileName();
      if (openedFileName) {
        state.model.sourceName = fileNameWithExtension(openedFileName, "html");
        document.title = state.model.sourceName;
        $workbookName.text(state.model.sourceName);
      }
    }
    clearAppAlert();
    if (!options.deferLoadedStatus) setStatus("HTML ワークブックを読み込みました。");
    return true;
  } catch (error) {
    console.error(error);
    state.model = null;
    state.hf = null;
    resetCellStyleInternPools();
    openBlankWorkbook({ status: "HTML ワークブックを読み込めませんでした。新しいブックを作成しました。" });
    showAppAlert(`HTML ワークブックを読み込めませんでした：${readableLoadError(error)}`, "error");
    return false;
  } finally {
    if (loadVersion === state.workbookLoadVersion) state.persistenceLoading = false;
  }
}



function initializeStandaloneLauncher(options = {}) {
  closeShapeFormatPane({ restoreFocus: false });
  closeShapeSelectionPane({ restoreFocus: false });
  clearQueryRefreshTimers();
  state.isStandalone = true;
  state.model = null;
  state.hf = null;
  state.currentSaveType = "";
  state.openedFileSignature = null;
  resetCellStyleInternPools();
  state.selectedSheetIndexes = [];
  state.sheetContextMenu = null;
  state.cellContextMenu = null;
  state.headerSizeDialog = null;
  state.nameDialog = null;
  state.linkDialog = null;
  state.objectLinkDialog = null;
  state.shapeFormatPane = null;
  state.shapeSelectionPane = null;
  state.shapeSelectionPaneDrag = null;
  state.shapePointEditIds = [];
  state.symbolDialog = null;
  state.annotationEditor = null;
  state.undoStack = [];
  state.redoStack = [];
  state.workbookRevision = 0;
  markWorkbookSavedClean();
  clearWorkbookClipboardState();
  state.selected = null;
  state.selectedImage = null;
  state.selectedImages = [];
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  state.selectionDrag = null;
  state.autoFillOptions = null;
  state.formulaEdit = null;
  resetCurrentSaveTarget();
  document.title = APP_TITLE;
  showStartupLoadingState(options.status || "Cht WebSheet を読み込み中...");
  $exportHtmlButton.prop("disabled", true);
  $toggleFormulasButton.prop("disabled", true);
  $formulaInput.prop("disabled", true).val("");
  $formulaExpandButton.prop("disabled", true);
  const standaloneModelError = String(window.__WEBSHEET_STANDALONE_MODEL_ERROR__ || "").trim();
  if (options.showBlank === false) {
    if (standaloneModelError) {
      showAppAlert(`HTML ワークブックのデータを読み込めませんでした：${readableLoadError(standaloneModelError)}`, "error");
    } else {
      clearAppAlert();
    }
    setStatus(options.status || "Cht WebSheet を読み込み中...");
    return;
  }
  openBlankWorkbook({ status: standaloneModelError ? "HTML ワークブックのデータを読み込めませんでした。新しいブックを作成しました。" : "新しいブックを作成しました。" });
  if (standaloneModelError) {
    const message = `HTML ワークブックのデータを読み込めませんでした：${readableLoadError(standaloneModelError)}`;
    showAppAlert(message, "error");
  } else {
    clearAppAlert();
  }
}



function applyPersistedWorkbookSnapshot(snapshot) {
  state.cwsAgentPreview = null;
  state.cwsAgentPersistenceSnapshot = null;
  revokeWorksheetImageRenderObjectUrls();
  closeSheetTabContextMenu();
  closeCellContextMenu();
  closeHeaderSizeDialog({ restoreFocus: false });
  closeNameDialog();
  closeLinkDialog();
  closeObjectLinkDialog();
  closeShapeFormatPane();
  closeShapeSelectionPane({ restoreFocus: false });
  closeSymbolDialog({ restoreFocus: false });
  clearQueryRefreshTimers();
  cancelAnnotationEditor();
  closeCommentThreadPopup();
  closeNotePopup();

  resetCellStyleInternPools();
  state.model = normalizePersistedWorkbookModel(snapshot.model);
  state.openedFileSignature = normalizeWorkbookFileSignature(state.model.sourceFileSignature);
  state.currentSaveType = normalizeWorkbookSaveType(state.model.saveType, state.model.sourceName);
  resetCurrentSaveTarget();
  state.hf = buildFormulaEngine(state.model);
  primeWorkbookFormulaSpillAnchorCache(state.model);
  primeWorkbookUsedRangeCache(state.model);
  const openRefreshCount = refreshWorkbookQueriesOnOpen();
  state.activeSheetIndex = clampSheetIndex(snapshot.activeSheetIndex);
  state.selectedSheetIndexes = [];
  state.sheetContextMenu = null;
  state.cellContextMenu = null;
  state.headerSizeDialog = null;
  state.nameDialog = null;
  state.linkDialog = null;
  state.objectLinkDialog = null;
  state.shapeFormatPane = null;
  state.shapeSelectionPane = null;
  state.shapeSelectionPaneDrag = null;
  state.shapePointEditIds = [];
  state.symbolDialog = null;
  state.annotationEditor = null;
  state.undoStack = [];
  state.redoStack = [];
  clearWorkbookClipboardState();
  state.selectedImage = null;
  state.selectedImages = [];
  state.selectionAnchor = null;
  state.selectionDrag = null;
  state.selectionAutoScroll = null;
  state.fillDrag = null;
  state.autoFillOptions = null;
  resetLazyScrollLocks();
  state.formulaEdit = null;
  state.isSelecting = false;
  state.showFormulas = Boolean(snapshot.showFormulas);
  ensureActiveSheetVisible();
  resetRestoredWorkbookWindowSession();
  restoreWorkbookWindowSelection(initialWorkbookSelectionRecord(snapshot));
  state.formulaBarExpandedLines = Math.max(FORMULA_BAR_DEFAULT_EXPANDED_LINES, Number(snapshot.formulaBarExpandedLines) || FORMULA_BAR_DEFAULT_EXPANDED_LINES);
  setFormulaBarLines(Number(snapshot.formulaBarLines) || FORMULA_BAR_MIN_LINES, { resetExpandedLines: false });
  document.title = state.model.sourceName || APP_TITLE;
  $workbookName.text(state.model.sourceName || APP_TITLE);
  showWorkbookWorkspace();
  $exportHtmlButton.prop("disabled", false);
  $toggleFormulasButton
    .prop("disabled", false)
    .attr("title", state.showFormulas ? "計算結果を表示" : "数式を表示")
    .html(`<i data-lucide="${state.showFormulas ? "eye-off" : "eye"}"></i>`);
  $formulaInput.prop("disabled", false);
  $formulaExpandButton.prop("disabled", false);
  $formulaBar.toggleClass("hidden", state.model.showFormulaBar === false);
  $sheetHost.scrollTop(0);
  $sheetHost.scrollLeft(0);
  $horizontalScrollbar.scrollLeft(0);
  renderWorkbook();
  createIcons({ icons });
  scheduleQueryRefreshTimers();
  if (state.selected) updateFormulaBarForCell(state.selected.row, state.selected.col);
  if (openRefreshCount) setStatus(`${openRefreshCount} 件の接続を開くときに更新しました。`);
  requestAnimationFrame(() => {
    $sheetHost.scrollTop(0);
    $sheetHost.scrollLeft(0);
    $horizontalScrollbar.scrollLeft(0);
    updateHorizontalScrollbar();
  });
}



function initialWorkbookSelectionRecord(snapshot = {}) {
  const hasSnapshotSelection = Boolean(
    snapshot.selection ||
    snapshot.selectionRange ||
    (Array.isArray(snapshot.selectionRanges) && snapshot.selectionRanges.length),
  );
  if (hasSnapshotSelection) {
    return {
      selectedCell: snapshot.selection || null,
      selectionRange: snapshot.selectionRange || null,
      selectionRanges: Array.isArray(snapshot.selectionRanges) ? snapshot.selectionRanges : [],
    };
  }
  return state.model?.mainWindowView || {};
}



function resetRestoredWorkbookWindowSession() {
  if (!state.model) return;
  const sheet = activeSheet();
  state.model.activeWindowId = "main";
  state.model.mainWindowHidden = false;
  state.model.windowViews = [];
  state.model.viewSideBySide = null;
  state.model.synchronousScrolling = false;
  state.model.windowArrangement = null;
  state.model.mainWindowView = workbookWindowRecordFromSheetViewState(sheet, state.activeSheetIndex, {
    id: "main",
    name: state.model.sourceName || "Book1",
    sheetIndex: state.activeSheetIndex,
    viewMode: sheet?.viewMode || "normal",
    zoomScale: sheet?.zoomScale || 100,
    showGridLines: sheet?.showGridLines !== false,
    showHeadings: sheetHeadingsVisible(sheet),
    showRuler: normalizedSheetShowRuler(sheet?.showRuler, sheet?.viewMode),
    showObjects: sheet?.showObjects !== false,
    showOutlineSymbols: sheet?.showOutlineSymbols !== false,
    showFormulas: sheet?.showFormulas === true,
    showZeros: sheet?.showZeros !== false,
    showWhiteSpace: sheet?.showWhiteSpace !== false,
    rightToLeft: sheet?.rightToLeft === true,
    freezePanes: cloneJsonLikeValue(sheet?.freezePanes),
    splitPane: cloneJsonLikeValue(sheet?.splitPane),
    showHorizontalScroll: state.model.workbookView?.showHorizontalScroll !== false,
    showVerticalScroll: state.model.workbookView?.showVerticalScroll !== false,
    showSheetTabs: state.model.workbookView?.showSheetTabs !== false,
    scrollTop: 0,
    scrollLeft: 0,
    hidden: false,
  });
}



function normalizePersistedWorkbookModel(model) {
  const normalized = model && typeof model === "object" ? model : { sourceName: "workbook.xlsx", sheets: [] };
  const sourceVersion = Number(normalized.version) || 1;
  const needsLegacyDisplayMigration = sourceVersion < 2;
  const needsDrawingStarAdjustmentMigration = sourceVersion < 3;
  const needsDrawingCalloutAdjustmentMigration = sourceVersion < 4;
  const needsLineShapePaddingMigration = sourceVersion < 5;
  normalized.version = WORKBOOK_MODEL_VERSION;
  normalized.sourceName = normalized.sourceName || "workbook.xlsx";
  normalized.saveType = normalizeWorkbookSaveType(normalized.saveType, normalized.sourceName);
  normalized.createdAt = normalized.createdAt || new Date().toISOString();
  normalized.themeColors = normalizedThemeColorPalette(normalized.themeColors);
  normalized.themeEffects = normalizedPageLayoutThemeEffects(normalized.themeEffects);
  normalized.definedNames = Array.isArray(normalized.definedNames) ? normalized.definedNames : [];
  normalized.formulaWatches = normalizeFormulaWatches(normalized.formulaWatches, normalized);
  normalized.relationships = Array.isArray(normalized.relationships) ? normalized.relationships : [];
  normalized.macros = Array.isArray(normalized.macros) ? normalized.macros : [];
  normalized.macroRecording = normalized.macroRecording && typeof normalized.macroRecording === "object" ? normalized.macroRecording : null;
  normalized.macroRelativeReferences = normalized.macroRelativeReferences === true;
  normalized.macroSecurity = normalizeMacroSecurity(normalized.macroSecurity);
  normalized.workbookProperties = normalizeWorkbookProperties(normalized.workbookProperties);
  normalized.calcProperties = normalizeWorkbookCalcProperties(normalized.calcProperties);
  normalized.ooxmlPackage = normalizePreservedOoxmlPackage(normalized.ooxmlPackage);
  normalized.windowViews = Array.isArray(normalized.windowViews) ? normalized.windowViews.map(normalizeWorkbookWindowView) : [];
  normalized.customViews = Array.isArray(normalized.customViews) ? normalized.customViews.map(normalizeWorkbookCustomView).filter(Boolean) : [];
  normalized.fullScreenView = normalized.fullScreenView === true;
  normalized.activeWindowId = normalized.activeWindowId || "main";
  normalized.mainWindowHidden = normalized.mainWindowHidden === true;
  normalized.mainWindowView = normalizeWorkbookWindowView({
    id: "main",
    name: normalized.sourceName,
    ...(normalized.mainWindowView && typeof normalized.mainWindowView === "object" ? normalized.mainWindowView : {}),
    hidden: normalized.mainWindowHidden,
  });
  normalized.viewSideBySide = normalized.viewSideBySide && typeof normalized.viewSideBySide === "object" ? normalized.viewSideBySide : null;
  normalized.synchronousScrolling = normalized.synchronousScrolling === true;
  normalized.showFormulaBar = normalized.showFormulaBar !== false;
  normalized.sheets = Array.isArray(normalized.sheets) && normalized.sheets.length ? normalized.sheets : [createBlankSheet("Sheet1")];
  normalized.assets = normalizeWorkbookAssets(normalized.assets);
  materializeWorkbookImageAssets(normalized);
  normalized.workbookView = normalizeWorkbookViewModel(normalized.workbookView, normalized.sheets.length);
  normalized.activeSheetIndex = clampWorkbookSheetIndex(normalized.activeSheetIndex ?? normalized.workbookView.activeSheetIndex, normalized.sheets.length);
  normalized.sheets.forEach((sheet, index) => {
    sheet.name = sheet.name || `Sheet${index + 1}`;
    sheet.chartSheet = sheet.chartSheet === true;
    sheet.sheetKind = normalizeSheetKind(sheet.sheetKind, sheet);
    sheet.specialSheet = normalizeSpecialSheetDescriptor(sheet.specialSheet, sheet);
    sheet.macroSheet = sheet.sheetKind === "macro" || sheet.sheetKind === "intlMacro";
    sheet.xl4MacroSheet = sheet.sheetKind === "macro";
    sheet.internationalMacroSheet = sheet.sheetKind === "intlMacro";
    sheet.dialogSheet = sheet.sheetKind === "dialog";
    sheet.cells = sheet.cells || {};
    sheet.rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    sheet.columns = Array.isArray(sheet.columns) ? sheet.columns : [];
    sheet.merges = Array.isArray(sheet.merges) ? sheet.merges : [];
    sheet.images = Array.isArray(sheet.images) ? sheet.images : [];
    sheet.inkStrokes = normalizeInkStrokes(sheet.inkStrokes);
    sheet.tables = Array.isArray(sheet.tables) ? sheet.tables : [];
    sheet.autoFilter = sheet.autoFilter?.range ? sheet.autoFilter : null;
    sheet.showGridLines = sheet.showGridLines !== false;
    sheet.showHeadings = sheet.showHeadings !== false;
    sheet.viewMode = normalizedSheetViewMode(sheet.viewMode);
    sheet.showRuler = normalizedSheetShowRuler(sheet.showRuler, sheet.viewMode);
    sheet.showObjects = sheet.showObjects !== false;
    sheet.showOutlineSymbols = sheet.showOutlineSymbols !== false;
    sheet.showFormulas = sheet.showFormulas === true;
    sheet.showZeros = sheet.showZeros !== false;
    sheet.showWhiteSpace = sheet.showWhiteSpace !== false;
    sheet.rightToLeft = sheet.rightToLeft === true;
    sheet.tabSelected = sheet.tabSelected === true;
    sheet.freezePanes = normalizePaneState(sheet.freezePanes);
    sheet.splitPane = normalizePaneState(sheet.splitPane);
    sheet.sheetViews = Array.isArray(sheet.sheetViews) ? sheet.sheetViews.map(normalizeSheetView).filter(Boolean) : [];
    sheet.activeSheetViewId = sheet.sheetViews.some((view) => view.id === sheet.activeSheetViewId) ? sheet.activeSheetViewId : null;
    sheet.codeModule = normalizeSheetCodeModule(sheet.codeModule, sheet);
    sheet.outlineSettings = normalizeOutlineSettings(sheet.outlineSettings);
    sheet.formulaAudits = normalizeFormulaAudits(sheet.formulaAudits);
    sheet.pageLayout = normalizedPageLayoutSettings(sheet.pageLayout, sheet);
    sheet.headerFooter = normalizedHeaderFooterSettings(sheet.headerFooter, sheet.pageLayout);
    sheet.hfData = Array.isArray(sheet.hfData) ? sheet.hfData : [];
    if (needsDrawingStarAdjustmentMigration) migratePersistedDrawingStarAdjustments(sheet);
    if (needsDrawingCalloutAdjustmentMigration) migratePersistedDrawingCalloutAdjustments(sheet);
    if (needsLineShapePaddingMigration) migratePersistedLineShapePadding(sheet);
    const hfDataDimensions = sheetHfDataDimensions(sheet);
    sheet.defaultCellCss = internCssObject(sheet.defaultCellCss);
    const customViewExtent = customViewExtentForSheet(normalized.customViews, sheet, index);
    const targetRows = Math.max(1, Number(sheet.rowCount) || sheet.rows.length || hfDataDimensions.rows || 1, customViewExtent.rows);
    const targetCols = Math.max(1, Number(sheet.colCount) || sheet.columns.length || hfDataDimensions.cols || 1, customViewExtent.cols);
    if (needsLegacyDisplayMigration) migrateLegacySheetDisplayScale(sheet);
    sheet.baseZoomScale = sheetBaseZoomScale(sheet);
    sheet.displayScale = Number.isFinite(Number(sheet.displayScale)) && Number(sheet.displayScale) > 0 ? Number(sheet.displayScale) : 1;
    sheet.rowCount = Math.max(1, sheet.rows.length || hfDataDimensions.rows || 1);
    sheet.colCount = Math.max(1, sheet.columns.length || hfDataDimensions.cols || 1);
    migrateDenseFormulaDataToCells(sheet);
    materializeSheetStyleTable(sheet);
    normalizeSheetStyleRanges(sheet);
    ensureSheetSize(sheet, targetRows, targetCols);
    normalizePersistedSheetDateCells(sheet);
    compactSheetDefaultCellCss(sheet);
    internSheetCellStyles(sheet);
  });
  return normalized;
}



function normalizeWorkbookAssets(assets) {
  const source = assets && typeof assets === "object" ? assets : {};
  const images = {};
  const sourceImages = source.images && typeof source.images === "object" ? source.images : {};
  if (Array.isArray(sourceImages)) {
    sourceImages.forEach((asset, index) => {
      const id = String(asset?.id || `img_${String(index + 1).padStart(4, "0")}`);
      const normalized = normalizeWorkbookImageAsset(asset);
      if (normalized) images[id] = normalized;
    });
  } else {
    Object.entries(sourceImages).forEach(([id, asset]) => {
      const normalized = normalizeWorkbookImageAsset(asset);
      if (normalized) images[String(id)] = normalized;
    });
  }
  return { ...source, images };
}



function normalizeWorkbookImageAsset(asset) {
  if (!asset || typeof asset !== "object") return null;
  const src = String(asset.src || "").trim();
  const parsedSrc = parseBase64ImageDataUrl(src);
  const mimeType = String(asset.mimeType || parsedSrc?.mimeType || imageMimeType(asset.extension || "png")).toLowerCase();
  const data = String(asset.data || parsedSrc?.data || "").trim();
  if (!data) return null;
  return {
    ...asset,
    mimeType,
    extension: fileExtensionFromImageMimeType(mimeType, asset.extension || "png"),
    data,
    width: Math.max(1, Math.round(Number(asset.width) || 1)),
    height: Math.max(1, Math.round(Number(asset.height) || 1)),
  };
}



function materializeWorkbookImageAssets(model) {
  const assets = model?.assets?.images || {};
  if (!model?.sheets?.length || !Object.keys(assets).length) return;
  (model.sheets || []).forEach((sheet) => {
    if (!Array.isArray(sheet.images)) return;
    sheet.images.forEach((image) => {
      if (!image || image.src || !image.assetId) return;
      const asset = assets[String(image.assetId)];
      const src = workbookImageAssetSource(asset);
      if (!src) return;
      image.src = src;
      image.mimeType = image.mimeType || asset.mimeType;
      image.extension = image.extension || asset.extension;
    });
  });
  if (model.assets && typeof model.assets === "object") {
    delete model.assets.images;
    if (!Object.keys(model.assets).length) delete model.assets;
  }
}



function workbookImageAssetSource(asset) {
  if (!asset) return "";
  if (asset.src && String(asset.src).startsWith("data:image/")) return String(asset.src);
  const data = String(asset.data || "").trim();
  if (!data) return "";
  return `data:${asset.mimeType || imageMimeType(asset.extension || "png")};base64,${data}`;
}



function captureWorkbookThemeState(model = state.model) {
  return {
    themeColors: normalizedThemeColorPalette(model?.themeColors),
    themeFonts: cloneJsonLikeValue(model?.themeFonts || {}),
    themeEffects: normalizedPageLayoutThemeEffects(model?.themeEffects),
  };
}



function restoreWorkbookThemeState(snapshot = {}) {
  if (!state.model) return;
  state.model.themeColors = normalizedThemeColorPalette(snapshot.themeColors);
  state.model.themeFonts = cloneJsonLikeValue(snapshot.themeFonts || {});
  state.model.themeEffects = normalizedPageLayoutThemeEffects(snapshot.themeEffects);
}



function normalizeMacroSecurity(value = {}) {
  const source = value && typeof value === "object" ? value : {};
  const level = MACRO_SECURITY_LEVELS.has(source.level) ? source.level : DEFAULT_MACRO_SECURITY_LEVEL;
  return {
    level,
    trustVbaObjectModel: source.trustVbaObjectModel === true,
    updatedAt: typeof source.updatedAt === "string" ? source.updatedAt : "",
  };
}



function normalizeSheetCodeModule(value, sheet) {
  if (!value || typeof value !== "object") return null;
  const code = String(value.code || "");
  const hasModuleData = code || value.updatedAt || value.name || value.objectName || value.procedure;
  if (!hasModuleData) return null;
  return {
    name: String(value.name || sheetCodeModuleName(sheet)),
    objectName: String(value.objectName || "Worksheet"),
    procedure: String(value.procedure || ""),
    code,
    updatedAt: typeof value.updatedAt === "string" ? value.updatedAt : "",
  };
}



function sheetCodeModuleName(sheet) {
  return `${String(sheet?.name || "Sheet")} (Code)`;
}



function normalizeSheetKind(value, sheet = {}) {
  if (sheet.chartSheet === true) return "chart";
  if (sheet.dialogSheet === true) return "dialog";
  if (sheet.internationalMacroSheet === true) return "intlMacro";
  if (sheet.xl4MacroSheet === true || sheet.macroSheet === true) return "macro";
  const kind = String(value || "worksheet");
  return ["worksheet", "chart", "macro", "intlMacro", "dialog"].includes(kind) ? kind : "worksheet";
}



function normalizeSpecialSheetDescriptor(value, sheet = {}) {
  const kind = normalizeSheetKind(sheet.sheetKind, sheet);
  if (!["macro", "intlMacro", "dialog"].includes(kind)) return null;
  const meta = SPECIAL_SHEET_INSERT_META[kind] || SPECIAL_SHEET_INSERT_META.macro;
  const source = value && typeof value === "object" ? value : {};
  return {
    type: meta.type,
    label: String(source.label || meta.label),
    createdAt: typeof source.createdAt === "string" ? source.createdAt : "",
  };
}



function normalizeWorkbookWindowView(view) {
  const source = view && typeof view === "object" ? view : {};
  const sheetIndex = Math.max(0, Number(source.sheetIndex) || 0);
  const visibility = ["visible", "hidden", "veryHidden"].includes(String(source.visibility || "visible"))
    ? String(source.visibility || "visible")
    : "visible";
  const viewMode = normalizedSheetViewMode(source.viewMode);
  return {
    id: source.id || uniqueWorksheetObjectId("window"),
    name: source.name || "",
    sheetIndex,
    firstSheet: Math.max(0, Math.round(Number(source.firstSheet) || 0)),
    visibility,
    x: Number.isFinite(Number(source.x)) ? Math.round(Number(source.x)) : undefined,
    y: Number.isFinite(Number(source.y)) ? Math.round(Number(source.y)) : undefined,
    width: Number.isFinite(Number(source.width)) && Number(source.width) > 0 ? Math.round(Number(source.width)) : undefined,
    height: Number.isFinite(Number(source.height)) && Number(source.height) > 0 ? Math.round(Number(source.height)) : undefined,
    maximized: source.maximized === true,
    minimized: source.minimized === true,
    showHorizontalScroll: source.showHorizontalScroll !== false,
    showVerticalScroll: source.showVerticalScroll !== false,
    showSheetTabs: source.showSheetTabs !== false,
    viewMode,
    zoomScale: clampZoom(source.zoomScale || 100),
    showGridLines: source.showGridLines !== false,
    showHeadings: source.showHeadings !== false,
    showRuler: normalizedSheetShowRuler(source.showRuler, viewMode),
    showObjects: source.showObjects !== false,
    showOutlineSymbols: source.showOutlineSymbols !== false,
    showFormulas: source.showFormulas === true,
    showZeros: source.showZeros !== false,
    showWhiteSpace: source.showWhiteSpace !== false,
    rightToLeft: source.rightToLeft === true,
    freezePanes: normalizePaneState(source.freezePanes),
    splitPane: normalizePaneState(source.splitPane),
    topLeftCell: normalizeCustomTopLeftCell(source.topLeftCell),
    scrollTop: Math.max(0, Math.round(Number(source.scrollTop) || 0)),
    scrollLeft: Math.max(0, Math.round(Number(source.scrollLeft) || 0)),
    scrollUnit: source.scrollUnit === "cell" ? "cell" : "pixel",
    selectedCell: normalizeCustomViewPoint(source.selectedCell, sheetIndex),
    selectionRange: normalizeCustomViewRange(source.selectionRange, sheetIndex),
    selectionRanges: Array.isArray(source.selectionRanges)
      ? source.selectionRanges.map((range) => normalizeCustomViewRange(range, sheetIndex)).filter(Boolean)
      : [],
    hidden: source.hidden === true || visibility === "hidden" || visibility === "veryHidden" || source.minimized === true,
  };
}



function normalizedSheetViewMode(value) {
  return ["normal", "page-layout", "page-break"].includes(value) ? value : "normal";
}



function normalizedSheetShowRuler(value, viewMode) {
  return normalizedSheetViewMode(viewMode) === "page-layout" && value === true;
}



function normalizeWorkbookViewModel(view, sheetCount = 1) {
  const source = view && typeof view === "object" ? view : {};
  const count = Math.max(1, Number(sheetCount) || 1);
  const activeSheetIndex = clampWorkbookSheetIndex(source.activeSheetIndex ?? source.activeTab, count);
  const firstSheet = clampWorkbookSheetIndex(source.firstSheet, count);
  const visibility = ["visible", "hidden", "veryHidden"].includes(String(source.visibility || "visible"))
    ? String(source.visibility || "visible")
    : "visible";
  return {
    activeSheetIndex,
    firstSheet,
    visibility,
    x: Number.isFinite(Number(source.x)) ? Math.round(Number(source.x)) : 0,
    y: Number.isFinite(Number(source.y)) ? Math.round(Number(source.y)) : 0,
    width: Number.isFinite(Number(source.width)) && Number(source.width) > 0 ? Math.round(Number(source.width)) : 24000,
    height: Number.isFinite(Number(source.height)) && Number(source.height) > 0 ? Math.round(Number(source.height)) : 14000,
    showHorizontalScroll: source.showHorizontalScroll !== false,
    showVerticalScroll: source.showVerticalScroll !== false,
    showSheetTabs: source.showSheetTabs !== false,
  };
}



function normalizeWorkbookProperties(properties = {}) {
  const source = properties && typeof properties === "object" ? properties : {};
  const defaultThemeVersion = Number(source.defaultThemeVersion);
  const showObjects = ["all", "placeholders", "none"].includes(String(source.showObjects || ""))
    ? String(source.showObjects)
    : "all";
  return {
    date1904: source.date1904 === true || source.date1904 === "1" || source.date1904 === 1,
    codeName: String(source.codeName || "").trim(),
    filterPrivacy: source.filterPrivacy === true || source.filterPrivacy === "1" || source.filterPrivacy === 1,
    showObjects,
    ...(Number.isFinite(defaultThemeVersion) && defaultThemeVersion > 0 ? { defaultThemeVersion: Math.round(defaultThemeVersion) } : {}),
  };
}



function normalizeWorkbookCalcProperties(properties = {}) {
  const source = properties && typeof properties === "object" ? properties : {};
  const calcMode = ["auto", "manual", "autoNoTable"].includes(String(source.calcMode || "")) ? String(source.calcMode) : "auto";
  const refMode = String(source.refMode || "").toUpperCase() === "R1C1" ? "R1C1" : "A1";
  const numberAttr = (name, fallback = null) => {
    const value = Number(source[name]);
    return Number.isFinite(value) ? value : fallback;
  };
  return {
    calcMode,
    refMode,
    calcId: numberAttr("calcId"),
    fullCalcOnLoad: source.fullCalcOnLoad !== false && source.fullCalcOnLoad !== "0" && source.fullCalcOnLoad !== 0,
    forceFullCalc: source.forceFullCalc === true || source.forceFullCalc === "1" || source.forceFullCalc === 1,
    calcOnSave: source.calcOnSave !== false && source.calcOnSave !== "0" && source.calcOnSave !== 0,
    concurrentCalc: source.concurrentCalc !== false && source.concurrentCalc !== "0" && source.concurrentCalc !== 0,
    fullPrecision: source.fullPrecision !== false && source.fullPrecision !== "0" && source.fullPrecision !== 0,
    iterate: source.iterate === true || source.iterate === "1" || source.iterate === 1,
    iterateCount: Math.max(1, Math.round(numberAttr("iterateCount", 100) || 100)),
    iterateDelta: Math.max(0, numberAttr("iterateDelta", 0.001) || 0.001),
  };
}



function workbookViewModelFromExcel(workbook, sheetCount = 1, workbookWindowViews = []) {
  const source = (Array.isArray(workbookWindowViews) && workbookWindowViews.length ? workbookWindowViews : workbook?.views || [])[0] || {};
  return normalizeWorkbookViewModel({
    activeSheetIndex: source.activeSheetIndex ?? source.activeTab,
    firstSheet: source.firstSheet,
    visibility: source.visibility,
    x: source.x,
    y: source.y,
    width: source.width,
    height: source.height,
    showHorizontalScroll: source.showHorizontalScroll,
    showVerticalScroll: source.showVerticalScroll,
    showSheetTabs: source.showSheetTabs,
  }, sheetCount);
}



function workbookWindowViewsFromExcel(workbook, sourceName, sheetCount = 1, workbookWindowViews = []) {
  const views = Array.isArray(workbookWindowViews) && workbookWindowViews.length
    ? workbookWindowViews
    : Array.isArray(workbook?.views) && workbook.views.length ? workbook.views : [{}];
  const baseName = String(sourceName || "Book1");
  const records = views.map((view, index) => normalizeWorkbookWindowView({
    id: index === 0 ? "main" : `window-${index + 1}`,
    name: index === 0 ? baseName : `${baseName}:${index + 1}`,
    sheetIndex: clampWorkbookSheetIndex(view?.activeSheetIndex ?? view?.activeTab, sheetCount),
    firstSheet: clampWorkbookSheetIndex(view?.firstSheet, sheetCount),
    visibility: view?.visibility || "visible",
    x: view?.x,
    y: view?.y,
    width: view?.width,
    height: view?.height,
    showHorizontalScroll: view?.showHorizontalScroll,
    showVerticalScroll: view?.showVerticalScroll,
    showSheetTabs: view?.showSheetTabs,
    hidden: view?.visibility === "hidden" || view?.visibility === "veryHidden",
  }));
  const mainWindowView = records[0] || normalizeWorkbookWindowView({ id: "main", name: baseName });
  const visible = records.find((record) => !record.hidden) || mainWindowView;
  return {
    mainWindowHidden: mainWindowView.hidden === true,
    mainWindowView,
    windowViews: records.slice(1),
    activeWindowId: visible.id || "main",
  };
}



function clampWorkbookSheetIndex(index, sheetCount = 1) {
  const max = Math.max(0, Math.round(Number(sheetCount) || 1) - 1);
  const numeric = Number(index);
  return Number.isFinite(numeric) ? Math.max(0, Math.min(max, Math.round(numeric))) : 0;
}



function normalizeWorkbookCustomView(view) {
  if (!view || typeof view !== "object") return null;
  const name = String(view.name || "").trim();
  if (!name) return null;
  const activeSheetIndex = Math.max(0, Number(view.activeSheetIndex) || 0);
  return {
    name,
    guid: normalizeCustomViewGuid(view.guid || view.id || name),
    createdAt: view.createdAt || new Date().toISOString(),
    activeSheetIndex,
    autoUpdate: view.autoUpdate === true,
    mergeInterval: Math.max(0, Math.round(Number(view.mergeInterval) || 0)),
    changesSavedWin: view.changesSavedWin === true,
    onlySync: view.onlySync === true,
    personalView: view.personalView === true,
    showFormulaBar: view.showFormulaBar !== false,
    includePrintSettings: view.includePrintSettings !== false,
    includeHiddenRowsColumnsAndFilters: view.includeHiddenRowsColumnsAndFilters !== false,
    selectedCell: normalizeCustomViewPoint(view.selectedCell, activeSheetIndex),
    selectionRange: normalizeCustomViewRange(view.selectionRange, activeSheetIndex),
    selectionRanges: Array.isArray(view.selectionRanges)
      ? view.selectionRanges.map((range) => normalizeCustomViewRange(range, activeSheetIndex)).filter(Boolean)
      : [],
    activeScroll: normalizeCustomViewScroll(view.activeScroll),
    topLeftCell: normalizeCustomTopLeftCell(view.topLeftCell),
    activeWindowId: view.activeWindowId || "main",
    mainWindowHidden: view.mainWindowHidden === true,
    mainWindowView: view.mainWindowView ? normalizeWorkbookWindowView({ id: "main", name: view.mainWindowView.name || "", ...view.mainWindowView }) : null,
    windowViews: Array.isArray(view.windowViews) ? view.windowViews.map(normalizeWorkbookWindowView) : [],
    workbookWindowsCaptured: view.workbookWindowsCaptured === true,
    viewSideBySide: view.viewSideBySide && typeof view.viewSideBySide === "object" ? cloneJsonLikeValue(view.viewSideBySide) : null,
    synchronousScrolling: view.synchronousScrolling === true,
    windowArrangement: view.windowArrangement && typeof view.windowArrangement === "object" ? cloneJsonLikeValue(view.windowArrangement) : null,
    sheets: Array.isArray(view.sheets) ? view.sheets.map((sheet, sheetIndex) => ({
      name: String(sheet?.name || ""),
      viewMode: normalizedSheetViewMode(sheet?.viewMode),
      zoomScale: clampZoom(sheet?.zoomScale || 100),
      showGridLines: sheet?.showGridLines !== false,
      showHeadings: sheet?.showHeadings !== false,
      showRuler: normalizedSheetShowRuler(sheet?.showRuler, sheet?.viewMode),
      showObjects: sheet?.showObjects !== false,
      showOutlineSymbols: sheet?.showOutlineSymbols !== false,
      showFormulas: sheet?.showFormulas === true,
      showZeros: sheet?.showZeros !== false,
      showWhiteSpace: sheet?.showWhiteSpace !== false,
      rightToLeft: sheet?.rightToLeft === true,
      selectedCell: normalizeCustomViewPoint(sheet?.selectedCell, sheetIndex),
      selectionRange: normalizeCustomViewRange(sheet?.selectionRange, sheetIndex),
      selectionRanges: Array.isArray(sheet?.selectionRanges)
        ? sheet.selectionRanges.map((range) => normalizeCustomViewRange(range, sheetIndex)).filter(Boolean)
        : [],
      activeScroll: normalizeCustomViewScroll(sheet?.activeScroll),
      topLeftCell: normalizeCustomTopLeftCell(sheet?.topLeftCell),
      freezePanes: normalizePaneState(sheet?.freezePanes),
      splitPane: normalizePaneState(sheet?.splitPane),
      sheetViews: Array.isArray(sheet?.sheetViews) ? sheet.sheetViews.map(normalizeSheetView).filter(Boolean) : [],
      activeSheetViewId: String(sheet?.activeSheetViewId || "") || null,
      pageLayout: sheet?.pageLayout ? normalizedPageLayoutSettings(sheet.pageLayout, sheet) : null,
      headerFooter: sheet?.headerFooter ? normalizedHeaderFooterSettings(sheet.headerFooter, sheet?.pageLayout || null) : null,
      hiddenRows: Array.isArray(sheet?.hiddenRows) ? sheet.hiddenRows.map((row) => Math.max(1, Math.round(Number(row) || 1))).filter(Boolean) : [],
      hiddenColumns: Array.isArray(sheet?.hiddenColumns) ? sheet.hiddenColumns.map((col) => Math.max(1, Math.round(Number(col) || 1))).filter(Boolean) : [],
      autoFilter: cloneJsonLikeValue(sheet?.autoFilter),
      sortState: cloneJsonLikeValue(sheet?.sortState),
    })) : [],
  };
}



function normalizeCustomViewScroll(value) {
  const unit = value?.unit === "cell" ? "cell" : "pixel";
  return {
    top: Math.max(0, Math.round(Number(value?.top) || 0)),
    left: Math.max(0, Math.round(Number(value?.left) || 0)),
    ...(unit === "cell" ? { unit } : {}),
  };
}



function normalizeCustomTopLeftCell(value) {
  if (!value) return null;
  if (typeof value === "string") {
    const decoded = decodeCell(value.replace(/\$/g, ""));
    return decoded ? { row: decoded.row, col: decoded.col } : null;
  }
  if (typeof value !== "object") return null;
  const row = Math.max(1, Math.round(Number(value.row) || 1));
  const col = Math.max(1, Math.round(Number(value.col) || 1));
  return { row, col };
}



function customViewExtentForSheet(customViews = [], sheet = {}, sheetIndex = 0) {
  const extent = { rows: 0, cols: 0 };
  const includeCell = (cell, extraRows = 0, extraCols = 0) => {
    const normalized = normalizeCustomTopLeftCell(cell);
    if (!normalized) return;
    extent.rows = Math.max(extent.rows, normalized.row + Math.max(0, extraRows));
    extent.cols = Math.max(extent.cols, normalized.col + Math.max(0, extraCols));
  };
  const includePoint = (point) => {
    if (!point || typeof point !== "object") return;
    extent.rows = Math.max(extent.rows, Math.max(1, Math.round(Number(point.row) || 1)));
    extent.cols = Math.max(extent.cols, Math.max(1, Math.round(Number(point.col) || 1)));
  };
  const includeRange = (range) => {
    if (!range || typeof range !== "object") return;
    extent.rows = Math.max(extent.rows, Math.max(1, Math.round(Number(range.bottom || range.row || range.top) || 1)));
    extent.cols = Math.max(extent.cols, Math.max(1, Math.round(Number(range.right || range.col || range.left) || 1)));
  };
  const includePane = (pane) => {
    const normalized = normalizePaneState(pane);
    if (!normalized) return;
    extent.rows = Math.max(extent.rows, normalized.row + FALLBACK_VISIBLE_ROWS);
    extent.cols = Math.max(extent.cols, normalized.col + FALLBACK_VISIBLE_COLS);
  };
  const includeScroll = (scroll) => {
    const normalized = normalizeCustomViewScroll(scroll);
    if (normalized.unit !== "cell") return;
    extent.rows = Math.max(extent.rows, normalized.top + 1 + FALLBACK_VISIBLE_ROWS);
    extent.cols = Math.max(extent.cols, normalized.left + 1 + FALLBACK_VISIBLE_COLS);
  };
  const includeSheetSnapshot = (snapshot) => {
    if (!snapshot) return;
    includeCell(snapshot.topLeftCell, FALLBACK_VISIBLE_ROWS, FALLBACK_VISIBLE_COLS);
    includeScroll(snapshot.activeScroll);
    includePoint(snapshot.selectedCell);
    includeRange(snapshot.selectionRange);
    (snapshot.selectionRanges || []).forEach(includeRange);
    includePane(snapshot.freezePanes);
    includePane(snapshot.splitPane);
  };
  const sheetName = String(sheet?.name || "");
  (customViews || []).forEach((view) => {
    const savedSheet = (view?.sheets || []).find((candidate) => candidate?.name === sheetName) || view?.sheets?.[sheetIndex] || null;
    includeSheetSnapshot(savedSheet);
    if (Number(view?.activeSheetIndex) === sheetIndex) {
      includeCell(view.topLeftCell, FALLBACK_VISIBLE_ROWS, FALLBACK_VISIBLE_COLS);
      includeScroll(view.activeScroll);
      includePoint(view.selectedCell);
      includeRange(view.selectionRange);
      (view.selectionRanges || []).forEach(includeRange);
    }
    [view?.mainWindowView, ...(view?.windowViews || [])].forEach((record) => {
      if (!record || Math.max(0, Math.round(Number(record.sheetIndex) || 0)) !== sheetIndex) return;
      includeCell(record.topLeftCell, FALLBACK_VISIBLE_ROWS, FALLBACK_VISIBLE_COLS);
      if (record.scrollUnit === "cell") includeScroll({ top: record.scrollTop, left: record.scrollLeft, unit: "cell" });
      includePoint(record.selectedCell);
      includeRange(record.selectionRange);
      (record.selectionRanges || []).forEach(includeRange);
      includePane(record.freezePanes);
      includePane(record.splitPane);
    });
  });
  return extent;
}



function normalizeCustomViewGuid(value) {
  const text = String(value || "").trim();
  if (/^\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}$/i.test(text)) {
    return text.toUpperCase();
  }
  let hashA = 0x811c9dc5;
  let hashB = 0x45d9f3b;
  const source = text || `${Date.now()}:${Math.random()}`;
  for (let index = 0; index < source.length; index += 1) {
    const code = source.charCodeAt(index);
    hashA ^= code;
    hashA = Math.imul(hashA, 0x01000193) >>> 0;
    hashB ^= code + index;
    hashB = Math.imul(hashB, 0x27d4eb2d) >>> 0;
  }
  const hex = `${hashA.toString(16).padStart(8, "0")}${hashB.toString(16).padStart(8, "0")}${((hashA ^ hashB) >>> 0).toString(16).padStart(8, "0")}${(Math.imul((hashA + hashB) >>> 0, 0x85ebca6b) >>> 0).toString(16).padStart(8, "0")}`.slice(0, 32);
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20, 32)}}`.toUpperCase();
}



function normalizeCustomViewPoint(point, fallbackSheetIndex = 0) {
  if (!point || typeof point !== "object") return null;
  const row = Math.max(1, Math.round(Number(point.row) || 1));
  const col = Math.max(1, Math.round(Number(point.col) || 1));
  const sheetIndex = Number(point.sheetIndex);
  return {
    sheetIndex: Math.max(0, Math.round(Number.isFinite(sheetIndex) ? sheetIndex : fallbackSheetIndex)),
    row,
    col,
  };
}



function normalizeCustomViewRange(range, fallbackSheetIndex = 0) {
  if (!range || typeof range !== "object") return null;
  const top = Math.max(1, Math.round(Number(range.top) || Number(range.row) || 1));
  const left = Math.max(1, Math.round(Number(range.left) || Number(range.col) || 1));
  const bottom = Math.max(top, Math.round(Number(range.bottom) || top));
  const right = Math.max(left, Math.round(Number(range.right) || left));
  const sheetIndex = Number(range.sheetIndex);
  return {
    sheetIndex: Math.max(0, Math.round(Number.isFinite(sheetIndex) ? sheetIndex : fallbackSheetIndex)),
    top,
    left,
    bottom,
    right,
    ...(range.fullRow ? { fullRow: true } : {}),
    ...(range.fullColumn ? { fullColumn: true } : {}),
  };
}



function normalizeSheetView(view) {
  if (!view || typeof view !== "object") return null;
  const id = String(view.id || "").trim() || uniqueWorksheetObjectId("sheet-view");
  const name = String(view.name || "").trim() || "シート ビュー";
  return {
    id,
    name,
    temporary: view.temporary === true,
    createdAt: view.createdAt || new Date().toISOString(),
    updatedAt: view.updatedAt || view.createdAt || new Date().toISOString(),
    baseView: view.baseView ? normalizeSheetViewSnapshot(view.baseView) : null,
    snapshot: normalizeSheetViewSnapshot(view.snapshot || view),
  };
}



function normalizeSheetViewSnapshot(value = {}) {
  const source = value && typeof value === "object" ? value : {};
  const viewMode = normalizedSheetViewMode(source.viewMode);
  return {
    viewMode,
    zoomScale: clampZoom(source.zoomScale || 100),
    showGridLines: source.showGridLines !== false,
    showHeadings: source.showHeadings !== false,
    showRuler: normalizedSheetShowRuler(source.showRuler, viewMode),
    showObjects: source.showObjects !== false,
    showOutlineSymbols: source.showOutlineSymbols !== false,
    showFormulas: source.showFormulas === true,
    showZeros: source.showZeros !== false,
    showWhiteSpace: source.showWhiteSpace !== false,
    rightToLeft: source.rightToLeft === true,
    freezePanes: normalizePaneState(source.freezePanes),
    splitPane: normalizePaneState(source.splitPane),
    selectedCell: normalizeCustomViewPoint(source.selectedCell, state.activeSheetIndex),
    selectionRange: normalizeCustomViewRange(source.selectionRange, state.activeSheetIndex),
    selectionRanges: Array.isArray(source.selectionRanges)
      ? source.selectionRanges.map((range) => normalizeCustomViewRange(range, state.activeSheetIndex)).filter(Boolean)
      : [],
    activeScroll: {
      top: Math.max(0, Math.round(Number(source.activeScroll?.top ?? source.scrollTop) || 0)),
      left: Math.max(0, Math.round(Number(source.activeScroll?.left ?? source.scrollLeft) || 0)),
    },
    topLeftCell: normalizeCustomTopLeftCell(source.topLeftCell),
    hiddenRows: Array.isArray(source.hiddenRows) ? source.hiddenRows.map((row) => Math.max(1, Math.round(Number(row) || 1))).filter(Boolean) : [],
    hiddenColumns: Array.isArray(source.hiddenColumns) ? source.hiddenColumns.map((col) => Math.max(1, Math.round(Number(col) || 1))).filter(Boolean) : [],
    autoFilter: cloneJsonLikeValue(source.autoFilter),
    sortState: cloneJsonLikeValue(source.sortState),
  };
}



function normalizePaneState(pane) {
  if (!pane || typeof pane !== "object") return null;
  const row = Number(pane.row);
  const col = Number(pane.col);
  const rows = Number(pane.rows);
  const columns = Number(pane.columns);
  const rowCount = Number.isFinite(rows)
    ? Math.max(0, Math.round(rows))
    : Number.isFinite(row)
      ? Math.max(0, Math.round(row) - 1)
      : 0;
  const columnCount = Number.isFinite(columns)
    ? Math.max(0, Math.round(columns))
    : Number.isFinite(col)
      ? Math.max(0, Math.round(col) - 1)
      : 0;
  if (rowCount <= 0 && columnCount <= 0) return null;
  return {
    rows: rowCount,
    columns: columnCount,
    row: Math.max(1, rowCount + 1),
    col: Math.max(1, columnCount + 1),
    activePane: normalizedPaneActivePane(pane.activePane, rowCount, columnCount),
  };
}



function normalizedPaneActivePane(value, rows = 0, columns = 0) {
  const rowCount = Math.max(0, Math.round(Number(rows) || 0));
  const columnCount = Math.max(0, Math.round(Number(columns) || 0));
  const allowed = new Set(["topLeft"]);
  if (columnCount > 0) allowed.add("topRight");
  if (rowCount > 0) allowed.add("bottomLeft");
  if (rowCount > 0 && columnCount > 0) allowed.add("bottomRight");
  const text = String(value || "");
  if (allowed.has(text)) return text;
  return excelSplitActivePane(rowCount, columnCount) || "topLeft";
}



function normalizePersistedSheetDateCells(sheet) {
  Object.values(sheet?.cells || {}).forEach(normalizePersistedCellDateValues);
}



function normalizePersistedCellDateValues(cell) {
  if (!cell) return;
  const numFmt = cell.numFmt || "";
  const rawSerial = dateSerialFromPersistedValue(cell.raw, numFmt);
  const cachedSerial = dateSerialFromPersistedValue(cell.cached, numFmt);
  if (rawSerial != null && shouldNormalizePersistedDateScalar(cell.raw, numFmt)) cell.raw = rawSerial;
  if (cachedSerial != null && shouldNormalizePersistedDateScalar(cell.cached, numFmt)) cell.cached = cachedSerial;

  const displayDate = persistedDateDisplayValue(cell.display);
  const displayLooksLikeJsDate = isJsDateDisplayString(cell.display);
  if (displayDate) {
    const displaySerial = rawSerial ?? cachedSerial ?? dateToExcelSerial(displayDate);
    cell.display = formatPersistedDateDisplay(displaySerial, numFmt);
    return;
  }
  if (displayLooksLikeJsDate && isDateFormat(numFmt)) {
    const displaySerial = rawSerial ?? cachedSerial ?? dateToExcelSerial(new Date(String(cell.display)));
    cell.display = formatPersistedDateDisplay(displaySerial, numFmt);
  }
}



function shouldNormalizePersistedDateScalar(value, numFmt = "") {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return true;
  if (isJsDateDisplayString(value)) return true;
  if (!isDateFormat(numFmt)) return false;
  return Boolean(persistedDateDisplayValue(value));
}



function dateSerialFromPersistedValue(value, numFmt = "") {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const date = persistedDateObject(value);
  if (date && (isDateFormat(numFmt) || value instanceof Date || isJsDateDisplayString(value))) return dateToExcelSerial(date);
  if (isDateFormat(numFmt)) {
    const displayDate = persistedDateDisplayValue(value);
    if (displayDate) return dateToExcelSerial(displayDate);
  }
  return null;
}



function persistedDateObject(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "string" && isJsDateDisplayString(value)) return persistedDateDisplayValue(value);
  return null;
}



function persistedDateDisplayValue(value) {
  if (typeof value !== "string") return null;
  const text = value.trim();
  if (!text) return null;
  const parsed = parsePersistedDateString(text);
  if (parsed) return parsed;
  return null;
}



function parsePersistedDateString(text) {
  const source = String(text || "").trim();
  if (!source) return null;
  if (isIsoDateTimeString(source) || isJsDateDisplayString(source)) {
    const date = new Date(source);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const japaneseDate = source.match(/^(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (japaneseDate) {
    const [, year, month, day, hour = "0", minute = "0", second = "0"] = japaneseDate;
    return localDateFromPersistedParts(year, month, day, hour, minute, second);
  }
  const slashDate = source.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d{1,3})?)?)?$/);
  if (slashDate) {
    const [, year, month, day, hour = "0", minute = "0", second = "0"] = slashDate;
    return localDateFromPersistedParts(year, month, day, hour, minute, second);
  }
  return null;
}



function localDateFromPersistedParts(year, month, day, hour = "0", minute = "0", second = "0") {
  const y = Number(year);
  const m = Number(month);
  const d = Number(day);
  const h = Number(hour);
  const min = Number(minute);
  const sec = Number(second);
  if (
    !Number.isInteger(y) ||
    !Number.isInteger(m) ||
    !Number.isInteger(d) ||
    !Number.isInteger(h) ||
    !Number.isInteger(min) ||
    !Number.isInteger(sec) ||
    m < 1 ||
    m > 12 ||
    d < 1 ||
    h < 0 ||
    h > 23 ||
    min < 0 ||
    min > 59 ||
    sec < 0 ||
    sec > 59
  ) {
    return null;
  }
  const date = new Date(y, m - 1, d, h, min, sec);
  if (
    date.getFullYear() !== y ||
    date.getMonth() !== m - 1 ||
    date.getDate() !== d ||
    date.getHours() !== h ||
    date.getMinutes() !== min ||
    date.getSeconds() !== sec
  ) {
    return null;
  }
  return date;
}



function isIsoDateTimeString(value) {
  return /^\d{4}-\d{2}-\d{2}(?:[T\s]\d{2}:\d{2}(?::\d{2}(?:\.\d{1,3})?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/.test(String(value || "").trim());
}



function isJsDateDisplayString(value) {
  if (typeof value !== "string") return false;
  return /^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{2}\s+\d{4}\s+\d{2}:\d{2}:\d{2}\s+GMT[+-]\d{4}/.test(value.trim());
}



function formatPersistedDateDisplay(serial, numFmt = "") {
  if (Number.isFinite(Number(serial)) && isDateFormat(numFmt)) return displayValue(Number(serial), numFmt);
  if (Number.isFinite(Number(serial))) return displayValue(Number(serial), "yyyy/m/d h:mm");
  return "";
}



function migratePersistedDrawingStarAdjustments(sheet) {
  (sheet.images || []).forEach((image) => {
    const shape = image?.shape;
    const type = shape?.type || "";
    const sourcePreset = shape?.sourcePreset || "";
    if (!isStarShapeType(type) || !isStarShapeType(sourcePreset)) return;
    const currentInner = Number(shape.adjust?.inner);
    if (!Number.isFinite(currentInner) || currentInner > 0.5) return;
    if (Math.abs(currentInner - defaultStarInnerRatio(type)) < 0.0001) return;
    const nextInner = clampNumber(currentInner * 2, 0.02, 1, defaultStarInnerRatio(type));
    if (Math.abs(nextInner - currentInner) < 0.0001) return;
    const migratedShape = {
      ...shape,
      adjust: {
        ...shape.adjust,
        inner: nextInner,
      },
    };
    const descriptor = shapeDescriptorForType(type, {
      ...migratedShape,
      width: image.ext?.width,
      height: image.ext?.height,
    });
    image.src = descriptor.src;
    image.mimeType = descriptor.mimeType;
    image.extension = descriptor.extension;
    image.shape = {
      ...descriptor.shape,
      sourcePreset,
    };
  });
}



function migratePersistedDrawingCalloutAdjustments(sheet) {
  (sheet.images || []).forEach((image) => {
    const shape = image?.shape;
    const type = shape?.type || "";
    if (!isCalloutShapeType(type) || !String(shape?.sourcePreset || "").includes("Callout")) return;
    const current = normalizeShapeAdjustments(type, shape.adjust);
    if (!current) return;
    const legacyTailY = Number(shape.adjust?.tailY);
    const nextTailX = current.tailX <= 0.12 ? defaultCalloutTailX(type) : current.tailX;
    const nextTailY = Number.isFinite(legacyTailY) && legacyTailY <= 1 ? 0.5 + legacyTailY : current.tailY;
    const migratedShape = {
      ...shape,
      adjust: {
        ...shape.adjust,
        tailX: nextTailX,
        tailY: nextTailY,
      },
    };
    const descriptor = shapeDescriptorForType(type, {
      ...migratedShape,
      width: image.ext?.width,
      height: image.ext?.height,
    });
    image.src = descriptor.src;
    image.mimeType = descriptor.mimeType;
    image.extension = descriptor.extension;
    image.shape = {
      ...descriptor.shape,
      sourcePreset: shape.sourcePreset,
    };
  });
}



function migratePersistedLineShapePadding(sheet) {
  (sheet.images || []).forEach((image) => {
    if (!isLineShapeObject(image)) return;
    ensureLineShapeBoxPadding(sheet, image);
    refreshShapeSource(image);
  });
}



function isStarShapeType(type) {
  return /^star(?:4|5|6|7|8|10|12|16|24|32)$/.test(String(type || ""));
}



function migrateLegacySheetDisplayScale(sheet) {
  const scale = Number(sheet.displayScale || (Number(sheet.zoomScale) > 0 ? Number(sheet.zoomScale) / 100 : 1));
  sheet.baseZoomScale = Number(sheet.zoomScale) > 0 ? Number(sheet.zoomScale) : 100;
  if (Number.isFinite(scale) && scale > 0 && Math.abs(scale - 1) > 0.0001) {
    scaleSheetDisplay(sheet, 1 / scale);
  }
  sheet.displayScale = 1;
}



function clampSheetIndex(index) {
  const numeric = Number(index);
  const max = Math.max(0, (state.model?.sheets?.length || 1) - 1);
  return Number.isInteger(numeric) ? Math.max(0, Math.min(max, numeric)) : 0;
}



function normalizePersistedSelection(selection) {
  if (!selection || selection.sheetIndex !== state.activeSheetIndex) return null;
  const row = Math.max(1, Number(selection.row) || 1);
  const col = Math.max(1, Number(selection.col) || 1);
  ensureSheetSize(activeSheet(), row, col);
  return { sheetIndex: state.activeSheetIndex, row, col };
}



function normalizePersistedRange(range, fallbackPoint = null) {
  if (!range && fallbackPoint) return selectionRangeFromPoints(fallbackPoint, fallbackPoint);
  if (!range || range.sheetIndex !== state.activeSheetIndex) return null;
  const sheet = activeSheet();
  const normalized = {
    sheetIndex: state.activeSheetIndex,
    top: Math.max(1, Number(range.top) || 1),
    left: Math.max(1, Number(range.left) || 1),
    bottom: Math.max(1, Number(range.bottom) || Number(range.top) || 1),
    right: Math.max(1, Number(range.right) || Number(range.left) || 1),
  };
  ensureSheetSize(sheet, normalized.bottom, normalized.right);
  return expandRangeForMerges(sheet, normalized);
}



function scheduleWorkbookPersistence() {
  if (state.persistenceLoading || !state.model || !persistenceAvailable()) return;
  if (state.persistenceTimer) window.clearTimeout(state.persistenceTimer);
  cancelWorkbookPersistenceIdleCallback();
  state.persistenceTimer = window.setTimeout(() => {
    state.persistenceTimer = null;
    queueWorkbookPersistenceFlush();
  }, PERSISTENCE_SAVE_DEBOUNCE_MS);
}



function queueWorkbookPersistenceFlush() {
  if (typeof window.requestIdleCallback === "function") {
    state.persistenceIdleCallback = window.requestIdleCallback(() => {
      state.persistenceIdleCallback = null;
      void flushWorkbookPersistence();
    }, { timeout: PERSISTENCE_IDLE_TIMEOUT_MS });
    return;
  }
  void flushWorkbookPersistence();
}



function cancelWorkbookPersistenceIdleCallback() {
  if (state.persistenceIdleCallback == null || typeof window.cancelIdleCallback !== "function") {
    state.persistenceIdleCallback = null;
    return;
  }
  window.cancelIdleCallback(state.persistenceIdleCallback);
  state.persistenceIdleCallback = null;
}



async function flushWorkbookPersistence() {
  if (state.persistenceTimer) {
    window.clearTimeout(state.persistenceTimer);
    state.persistenceTimer = null;
  }
  cancelWorkbookPersistenceIdleCallback();
  if (state.persistenceLoading || !state.model || !persistenceAvailable()) return false;
  const snapshot = currentWorkbookSnapshot();
  try {
    state.persistencePromise = writePersistedWorkbookSnapshot(snapshot);
    await state.persistencePromise;
    return true;
  } catch (error) {
    console.warn("Failed to save workbook snapshot.", error);
    return false;
  } finally {
    state.persistencePromise = null;
  }
}



function beginCwsAgentPreviewPersistence() {
  // A page crash must restore the last committed workbook, never a partial animation.
  const snapshot = currentWorkbookSnapshot();
  state.cwsAgentPersistenceSnapshot = { ...snapshot, model: cloneJsonLikeValue(snapshot.model) };
}



function currentWorkbookSnapshot() {
  if (state.cwsAgentPersistenceSnapshot && state.cwsAgentPersistenceSnapshot.model?.cwsAgentDocumentId === state.model?.cwsAgentDocumentId) {
    return { ...state.cwsAgentPersistenceSnapshot, savedAt: new Date().toISOString() };
  }
  const host = $sheetHost[0];
  normalizeWorkbookDateCellsForPersistence(state.model);
  return {
    version: PERSISTENCE_SNAPSHOT_VERSION,
    savedAt: new Date().toISOString(),
    model: state.model,
    activeSheetIndex: state.activeSheetIndex,
    selection: state.selected ? { ...state.selected } : null,
    selectionRange: state.selectionRange ? { ...state.selectionRange } : null,
    selectionRanges: (state.selectionRanges || []).map((range) => ({ ...range })),
    showFormulas: state.showFormulas,
    formulaBarLines: state.formulaBarLines,
    formulaBarExpandedLines: state.formulaBarExpandedLines,
    scroll: {
      top: host?.scrollTop || 0,
      left: host?.scrollLeft || 0,
    },
  };
}



function normalizeWorkbookDateCellsForPersistence(model) {
  (model?.sheets || []).forEach(normalizePersistedSheetDateCells);
}



async function standaloneWorkbookPersistenceScope(model) {
  if (typeof globalThis.indexedDB === "undefined" || !globalThis.crypto?.subtle) return null;
  // The embedded source identifies this edition of the file. A newly exported
  // edition (even with the same document ID/name) must never load an older draft.
  const bytes = new TextEncoder().encode(JSON.stringify(model));
  const hash = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", bytes)), byte => byte.toString(16).padStart(2, "0")).join("");
  return { key: `standalone:${hash}`, documentId: model.cwsAgentDocumentId || `standalone-${hash}` };
}



function workbookPersistenceKey() {
  if (!state.isStandalone) return PERSISTENCE_CURRENT_KEY;
  const scope = state.standalonePersistenceScope;
  return scope?.documentId === state.model?.cwsAgentDocumentId ? scope?.key || null : null;
}



function persistenceAvailable() {
  return typeof globalThis.indexedDB !== "undefined" && Boolean(workbookPersistenceKey());
}



function markPendingFileLoadRecovery(file) {
  if (!persistenceAvailable() || typeof localStorage === "undefined") return;
  try {
    localStorage.setItem(PENDING_FILE_LOAD_RECOVERY_STORAGE_KEY, JSON.stringify({
      startedAt: Date.now(),
      name: String(file?.name || ""),
      size: Number(file?.size) || 0,
      lastModified: Number(file?.lastModified) || 0,
    }));
  } catch {
    // If localStorage is blocked, the normal load failure recovery still applies.
  }
}



function clearPendingFileLoadRecoveryMarker() {
  if (typeof localStorage === "undefined") return;
  try {
    localStorage.removeItem(PENDING_FILE_LOAD_RECOVERY_STORAGE_KEY);
  } catch {
    // Ignore storage failures.
  }
}



function consumePendingFileLoadRecoveryMarker() {
  if (!persistenceAvailable() || typeof localStorage === "undefined") return false;
  try {
    const raw = localStorage.getItem(PENDING_FILE_LOAD_RECOVERY_STORAGE_KEY);
    if (!raw) return false;
    localStorage.removeItem(PENDING_FILE_LOAD_RECOVERY_STORAGE_KEY);
    return true;
  } catch {
    return false;
  }
}



function openPersistenceDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(PERSISTENCE_DB_NAME, PERSISTENCE_DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(PERSISTENCE_STORE_NAME)) {
        db.createObjectStore(PERSISTENCE_STORE_NAME);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    request.onblocked = () => reject(new Error("IndexedDB upgrade was blocked."));
  });
}



async function readPersistedWorkbookSnapshot(key = workbookPersistenceKey()) {
  if (!key) throw new Error("保存対象のブックを確認できません。");
  const db = await openPersistenceDb();
  try {
    return await idbRequest(db.transaction(PERSISTENCE_STORE_NAME, "readonly").objectStore(PERSISTENCE_STORE_NAME).get(key));
  } finally {
    db.close();
  }
}



async function writePersistedWorkbookSnapshot(snapshot, key = workbookPersistenceKey()) {
  if (!key) throw new Error("保存対象のブックを確認できません。");
  const db = await openPersistenceDb();
  try {
    const tx = db.transaction(PERSISTENCE_STORE_NAME, "readwrite");
    tx.objectStore(PERSISTENCE_STORE_NAME).put(snapshot, key);
    await idbTransactionDone(tx);
  } finally {
    db.close();
  }
}



async function deletePersistedWorkbookSnapshot(key = workbookPersistenceKey()) {
  if (!key) return;
  const db = await openPersistenceDb();
  try {
    const tx = db.transaction(PERSISTENCE_STORE_NAME, "readwrite");
    tx.objectStore(PERSISTENCE_STORE_NAME).delete(key);
    await idbTransactionDone(tx);
  } finally {
    db.close();
  }
}



function idbRequest(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}



function idbTransactionDone(tx) {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error("IndexedDB transaction aborted."));
  });
}



function workbookToModel(
  workbook,
  sourceName,
  workbookPhonetics = {},
  workbookDrawingShapes = {},
  themeColors = parseThemeColors(workbook),
  workbookDefinedNames = null,
  workbookAutoFilters = {},
  workbookDataValidationDetails = {},
  workbookSheetViewDetails = {},
  workbookOutlineDetails = {},
  workbookWindowViews = [],
  workbookCustomViews = [],
  workbookMetadata = {},
) {
  debugWorkbookLoadTrace("workbook-to-model:media:start");
  const media = workbookMediaToDataUrls(workbook);
  debugWorkbookLoadTrace("workbook-to-model:media:done");
  debugWorkbookLoadTrace("workbook-to-model:width-metrics:start");
  const widthMetrics = inferWorkbookWidthMetrics(workbook);
  debugWorkbookLoadTrace("workbook-to-model:width-metrics:done");
  debugWorkbookLoadTrace("workbook-to-model:defined-names:start");
  const definedNames = workbookDefinedNames?.length ? workbookDefinedNames : workbookDefinedNamesToModel(workbook);
  debugWorkbookLoadTrace("workbook-to-model:defined-names:done");
  debugWorkbookLoadTrace("workbook-to-model:properties:start");
  const workbookProperties = normalizeWorkbookProperties({
    ...(workbook?.properties || {}),
    ...(workbookMetadata?.workbookProperties || {}),
  });
  const calcProperties = normalizeWorkbookCalcProperties({
    ...(workbook?.calcProperties || {}),
    ...(workbookMetadata?.calcProperties || {}),
  });
  debugWorkbookLoadTrace("workbook-to-model:properties:done");
  const sheets = workbook.worksheets.map((worksheet) => {
    debugWorkbookLoadTrace(`workbook-to-model:sheet:${worksheet.name}:start`);
    const sheetModel = worksheetToModel(
      worksheet,
      themeColors,
      media,
      widthMetrics,
      workbookPhonetics[worksheet.name] || {},
      workbookDrawingShapes[worksheet.name] || [],
      workbookAutoFilters[worksheet.name] || null,
      workbookDataValidationDetails[worksheet.name] || null,
      workbookSheetViewDetails[worksheet.name] || null,
      workbookOutlineDetails[worksheet.name] || null,
      workbookProperties,
    );
    debugWorkbookLoadTrace(`workbook-to-model:sheet:${worksheet.name}:done`);
    return sheetModel;
  });
  debugWorkbookLoadTrace("workbook-to-model:window-view:start");
  const workbookView = workbookViewModelFromExcel(workbook, sheets.length, workbookWindowViews);
  const windowViewState = workbookWindowViewsFromExcel(workbook, sourceName, sheets.length, workbookWindowViews);
  const activeSheetIndex = windowViewState.activeWindowId === "main"
    ? workbookView.activeSheetIndex
    : windowViewState.windowViews.find((view) => view.id === windowViewState.activeWindowId)?.sheetIndex ?? workbookView.activeSheetIndex;
  if (sheets[activeSheetIndex]) {
    windowViewState.mainWindowView = workbookWindowRecordFromSheetViewState(
      sheets[activeSheetIndex],
      activeSheetIndex,
      windowViewState.mainWindowView,
    );
  }
  debugWorkbookLoadTrace("workbook-to-model:window-view:done");
  const model = {
    version: WORKBOOK_MODEL_VERSION,
    sourceName,
    saveType: saveTypeForSourceName(sourceName, "xlsx"),
    createdAt: new Date().toISOString(),
    themeColors,
    themeEffects: normalizedPageLayoutThemeEffects("office"),
    definedNames,
    formulaWatches: [],
    relationships: [],
    workbookProperties,
    calcProperties,
    workbookView,
    ...windowViewState,
    customViews: Array.isArray(workbookCustomViews) ? workbookCustomViews.map(normalizeWorkbookCustomView).filter(Boolean) : [],
    activeSheetIndex,
    sheets,
  };
  debugWorkbookLoadTrace("workbook-to-model:print-settings:start");
  applyDefinedNamePrintSettingsToModel(model);
  debugWorkbookLoadTrace("workbook-to-model:print-settings:done");
  debugWorkbookLoadTrace("workbook-to-model:intern:start");
  internWorkbookCellStyles(model);
  debugWorkbookLoadTrace("workbook-to-model:intern:done");
  return model;
}



function workbookWindowRecordFromSheetViewState(sheet, sheetIndex, base = {}) {
  const topLeftCell = normalizeCustomTopLeftCell(sheet?.topLeftCell);
  const scroll = topLeftCell ? scrollRecordFromTopLeftCell(sheet, topLeftCell) : null;
  return normalizeWorkbookWindowView({
    ...(base || {}),
    sheetIndex,
    viewMode: sheet?.viewMode || base?.viewMode,
    zoomScale: sheet?.zoomScale || base?.zoomScale,
    showGridLines: sheet?.showGridLines,
    showHeadings: sheet?.showHeadings,
    showRuler: sheet?.showRuler,
    showObjects: sheet?.showObjects,
    showOutlineSymbols: sheet?.showOutlineSymbols,
    showFormulas: sheet?.showFormulas,
    showZeros: sheet?.showZeros,
    showWhiteSpace: sheet?.showWhiteSpace,
    rightToLeft: sheet?.rightToLeft,
    freezePanes: sheet?.freezePanes || base?.freezePanes,
    splitPane: sheet?.splitPane || base?.splitPane,
    selectedCell: sheet?.selectedCell || base?.selectedCell,
    selectionRange: sheet?.selectionRange || base?.selectionRange,
    selectionRanges: Array.isArray(sheet?.selectionRanges) && sheet.selectionRanges.length ? sheet.selectionRanges : base?.selectionRanges || [],
    topLeftCell: topLeftCell || base?.topLeftCell,
    ...(scroll ? { scrollTop: scroll.top, scrollLeft: scroll.left, scrollUnit: "cell" } : {}),
  });
}



function workbookDefinedNamesToModel(workbook) {
  return (workbook.definedNames?.model || [])
    .map((item) => {
      const ranges = (item.ranges || []).map((range) => String(range || "").trim()).filter(Boolean);
      if (!item.name || !ranges.length) return null;
      const localSheetId = Number(item.localSheetId);
      return {
        name: String(item.name),
        scope: Number.isInteger(localSheetId) && localSheetId >= 0 ? "sheet" : "workbook",
        sheetIndex: Number.isInteger(localSheetId) && localSheetId >= 0 ? localSheetId : null,
        refersTo: `=${ranges.join(",")}`,
        comment: "",
        hidden: String(item.name).startsWith("_xlnm."),
      };
    })
    .filter(Boolean);
}



function applyDefinedNamePrintSettingsToModel(model) {
  const sheets = model?.sheets || [];
  if (!sheets.length) return;
  (model.definedNames || []).forEach((definition) => {
    const name = String(definition?.name || "");
    if (name !== "_xlnm.Print_Area" && name !== "_xlnm.Print_Titles") return;
    const defaultSheetIndex = Number.isInteger(Number(definition.sheetIndex)) && Number(definition.sheetIndex) >= 0
      ? Number(definition.sheetIndex)
      : 0;
    const refs = splitFormulaReferenceList(String(definition.refersTo || "").replace(/^=/, ""));
    refs.forEach((refText) => {
      const decoded = decodeWorkbookRangeForModel(refText, model, defaultSheetIndex);
      if (!decoded?.range) return;
      const sheet = sheets[decoded.sheetIndex];
      if (!sheet) return;
      sheet.pageLayout = normalizedPageLayoutSettings(sheet.pageLayout, sheet);
      if (name === "_xlnm.Print_Area" && !sheet.pageLayout.printArea && decoded.kind === "cells") {
        sheet.pageLayout.printArea = {
          range: { ...decoded.range },
          label: rangeToLabel(decoded.range),
        };
      } else if (name === "_xlnm.Print_Titles") {
        sheet.pageLayout.printTitles = sheet.pageLayout.printTitles || { rows: null, cols: null };
        if (decoded.kind === "rows") {
          sheet.pageLayout.printTitles.rows = {
            ...decoded.range,
            fullRow: true,
            label: rangeToLabel({ ...decoded.range, fullRow: true }),
          };
        } else if (decoded.kind === "cols") {
          sheet.pageLayout.printTitles.cols = {
            ...decoded.range,
            fullColumn: true,
            label: rangeToLabel({ ...decoded.range, fullColumn: true }),
          };
        }
      }
    });
    sheets.forEach((sheet) => {
      if (sheet?.pageLayout) sheet.pageLayout = normalizedPageLayoutSettings(sheet.pageLayout, sheet);
    });
  });
}



function splitFormulaReferenceList(text) {
  const refs = [];
  let current = "";
  let inSheetQuote = false;
  for (const char of String(text || "")) {
    if (char === "'") {
      current += char;
      inSheetQuote = !inSheetQuote;
      continue;
    }
    if (char === "," && !inSheetQuote) {
      if (current.trim()) refs.push(current.trim());
      current = "";
      continue;
    }
    current += char;
  }
  if (current.trim()) refs.push(current.trim());
  return refs;
}



function decodeWorkbookRangeForModel(text, model, defaultSheetIndex = 0) {
  const clean = String(text || "").trim();
  if (!clean) return null;
  const match = clean.match(/^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?(.+)$/);
  if (!match) return null;
  const sheetName = match[1]?.replace(/''/g, "'") || match[2] || "";
  const sheetIndex = sheetName
    ? (model?.sheets || []).findIndex((sheet) => sheet.name === sheetName.replace(/^\[[^\]]+\]/, ""))
    : defaultSheetIndex;
  if (!Number.isInteger(sheetIndex) || sheetIndex < 0) return null;
  const sheet = model.sheets[sheetIndex];
  const decoded = decodePrintReferenceRange(match[3], sheet);
  return decoded ? { ...decoded, sheetIndex } : null;
}



function decodePrintReferenceRange(text, sheet = {}) {
  const ref = String(text || "").replace(/\$/g, "").trim().toUpperCase();
  const rowOnly = ref.match(/^(\d+):(\d+)$/);
  if (rowOnly) {
    const top = Math.max(1, Math.min(Number(rowOnly[1]), Number(rowOnly[2])));
    const bottom = Math.max(top, Math.max(Number(rowOnly[1]), Number(rowOnly[2])));
    return {
      kind: "rows",
      range: { top, bottom, left: 1, right: Math.max(1, Number(sheet.colCount) || 1), fullRow: true },
    };
  }
  const colOnly = ref.match(/^([A-Z]+):([A-Z]+)$/);
  if (colOnly) {
    const left = Math.max(1, Math.min(columnNumber(colOnly[1]), columnNumber(colOnly[2])));
    const right = Math.max(left, Math.max(columnNumber(colOnly[1]), columnNumber(colOnly[2])));
    return {
      kind: "cols",
      range: { top: 1, bottom: Math.max(1, Number(sheet.rowCount) || 1), left, right, fullColumn: true },
    };
  }
  const range = decodeRange(ref);
  return range ? { kind: "cells", range } : null;
}



function normalizePreservedOoxmlPackage(value = null) {
  if (!value || typeof value !== "object") return null;
  const entries = Array.isArray(value.entries)
    ? value.entries
        .map((entry) => ({
          path: String(entry?.path || "").replace(/^\/+/, ""),
          base64: String(entry?.base64 || ""),
        }))
        .filter((entry) => entry.path && entry.base64)
    : [];
  const workbookRelationships = Array.isArray(value.workbookRelationships)
    ? value.workbookRelationships.map((relationship) => ({
        id: String(relationship?.id || ""),
        type: String(relationship?.type || ""),
        target: String(relationship?.target || ""),
        targetMode: String(relationship?.targetMode || ""),
      })).filter((relationship) => relationship.type && relationship.target)
    : [];
  const contentTypeOverrides = Array.isArray(value.contentTypeOverrides)
    ? value.contentTypeOverrides.map((override) => ({
        partName: String(override?.partName || ""),
        contentType: String(override?.contentType || ""),
      })).filter((override) => override.partName && override.contentType)
    : [];
  const hasVbaProject = entries.some((entry) => /(^|\/)vbaProject\.bin$/i.test(entry.path));
  if (!entries.length && !workbookRelationships.length && !contentTypeOverrides.length) return null;
  return {
    kind: String(value.kind || "ooxml"),
    sourceName: String(value.sourceName || ""),
    hasVbaProject,
    entries,
    workbookRelationships,
    contentTypeOverrides,
  };
}



async function extractPreservedOoxmlPackage(buffer, sourceName = "") {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const entries = [];
    await Promise.all(Object.keys(zip.files)
      .filter((path) => preservedOoxmlEntryPath(path))
      .map(async (path) => {
        const file = zip.file(path);
        if (!file) return;
        entries.push({ path, base64: bytesToBase64(await file.async("uint8array")) });
      }));
    const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
    const contentTypesXml = await zip.file("[Content_Types].xml")?.async("text");
    const workbookRelationships = workbookRelsXml ? parsePreservedWorkbookRelationships(workbookRelsXml) : [];
    const contentTypeOverrides = contentTypesXml ? parsePreservedContentTypeOverrides(contentTypesXml) : [];
    return normalizePreservedOoxmlPackage({
      kind: "ooxml",
      sourceName,
      entries,
      workbookRelationships,
      contentTypeOverrides,
    });
  } catch (error) {
    console.warn("Failed to extract preserved OOXML package entries.", error);
    return null;
  }
}

export {
  applyDefinedNamePrintSettingsToModel,
  applyLegacyXlsCellStyleToCell,
  applyLegacyXlsCellStylesToModel,
  applyLegacyXlsMergedEdgeBorders,
  applyPersistedWorkbookSnapshot,
  applyStandaloneOpenWorkbookSaveTarget,
  beginCwsAgentPreviewPersistence,
  cancelWorkbookPersistenceIdleCallback,
  captureWorkbookRecoveryState,
  captureWorkbookThemeState,
  clampSheetIndex,
  clampWorkbookSheetIndex,
  clearPendingFileLoadRecoveryMarker,
  compareLegacyXlsBorderDescending,
  consumePendingFileLoadRecoveryMarker,
  convertLegacyXlsBufferToXlsxBuffer,
  createBlankWorkbookModel,
  createExcelJsWorkbookForImport,
  currentWorkbookSnapshot,
  customViewExtentForSheet,
  dateSerialFromPersistedValue,
  debugWorkbookLoadTrace,
  decodePrintReferenceRange,
  decodeWorkbookRangeForModel,
  deletePersistedWorkbookSnapshot,
  excelJsWorkbookFromFile,
  extractLegacyXlsCellStyles,
  extractPreservedOoxmlPackage,
  extractStandaloneWorkbookPayload,
  flushWorkbookPersistence,
  formatPersistedDateDisplay,
  handleWorkbookBeforeUnload,
  hasStandaloneWorkbookSheets,
  idbRequest,
  idbTransactionDone,
  initialWorkbookSelectionRecord,
  initializeStandaloneLauncher,
  isIsoDateTimeString,
  isJsDateDisplayString,
  isLegacyXlsWorkbook,
  isLegacyXlsWorkbookBuffer,
  isStarShapeType,
  isZipWorkbookBuffer,
  legacyXlsAddBorderSide,
  legacyXlsAlignmentCssToModel,
  legacyXlsBiffVersion,
  legacyXlsCachedXfStyleToModel,
  legacyXlsCellStyleHasVisibleEmptyCellFormat,
  legacyXlsCellStyleHasVisibleFormat,
  legacyXlsCollectCellStyles,
  legacyXlsEmptyCellVisibleStyle,
  legacyXlsFillColorToCss,
  legacyXlsFontColorToCss,
  legacyXlsFontForIndex,
  legacyXlsFontSizeCss,
  legacyXlsHorizontalAlign,
  legacyXlsIndexedColorToCss,
  legacyXlsMergedEdgeBorder,
  legacyXlsMergedEdgeBorders,
  legacyXlsNumberFormat,
  legacyXlsParseFont,
  legacyXlsParseFormat,
  legacyXlsParsePalette,
  legacyXlsParseXf,
  legacyXlsReadShortUnicodeString,
  legacyXlsReadSingleByteString,
  legacyXlsReadUInt16,
  legacyXlsReadUInt32,
  legacyXlsReadUnicodeString2,
  legacyXlsReadUnicodeStringNoLength,
  legacyXlsStyleOnlyCell,
  legacyXlsVerticalAlign,
  legacyXlsXfBordersToModel,
  legacyXlsXfCssToModel,
  legacyXlsXfStyleToModel,
  loadExcelFile,
  loadStandaloneOpenWorkbookFile,
  loadStandaloneWorkbookModel,
  localDateFromPersistedParts,
  markPendingFileLoadRecovery,
  markWorkbookChanged,
  markWorkbookSavedClean,
  materializeWorkbookImageAssets,
  mergeLegacyXlsCellBorders,
  migrateLegacySheetDisplayScale,
  migratePersistedDrawingCalloutAdjustments,
  migratePersistedDrawingStarAdjustments,
  migratePersistedLineShapePadding,
  normalizeArrayBuffer,
  normalizeCustomTopLeftCell,
  normalizeCustomViewGuid,
  normalizeCustomViewPoint,
  normalizeCustomViewRange,
  normalizeCustomViewScroll,
  normalizeMacroSecurity,
  normalizePaneState,
  normalizePersistedCellDateValues,
  normalizePersistedRange,
  normalizePersistedSelection,
  normalizePersistedSheetDateCells,
  normalizePersistedWorkbookModel,
  normalizePreservedOoxmlPackage,
  normalizeSheetCodeModule,
  normalizeSheetKind,
  normalizeSheetView,
  normalizeSheetViewSnapshot,
  normalizeSpecialSheetDescriptor,
  normalizeSpreadsheetMainNamespacePrefixesForExcelJs,
  normalizeSpreadsheetMainNamespacePrefixesXml,
  normalizeWorkbookAssets,
  normalizeWorkbookCalcProperties,
  normalizeWorkbookCustomView,
  normalizeWorkbookDateCellsForPersistence,
  normalizeWorkbookImageAsset,
  normalizeWorkbookProperties,
  normalizeWorkbookViewModel,
  normalizeWorkbookWindowView,
  normalizedPaneActivePane,
  normalizedSheetShowRuler,
  normalizedSheetViewMode,
  openBlankWorkbook,
  openPersistenceDb,
  parsePersistedDateString,
  persistedDateDisplayValue,
  persistedDateObject,
  persistenceAvailable,
  prepareExcelJsWorkbookForFastImport,
  queueWorkbookPersistenceFlush,
  readPersistedWorkbookSnapshot,
  recoverWorkbookAfterFailedFileLoad,
  releaseExcelJsWorkbook,
  resetRestoredWorkbookWindowSession,
  restorePersistedWorkbook,
  restoreWorkbookThemeState,
  scheduleWorkbookPersistence,
  sheetCodeModuleName,
  shouldNormalizePersistedDateScalar,
  splitFormulaReferenceList,
  standaloneOpenWorkbookFileFromWindow,
  standaloneOpenWorkbookPayloadToFile,
  standaloneWorkbookModelFromWindow,
  standaloneWorkbookPersistenceScope,
  standaloneWorkbookViewDiagnostics,
  stripDrawingContentTypeOverrides,
  stripDrawingRelationshipTags,
  stripUnsupportedTableAutoFilterXml,
  stripWorksheetDrawingTags,
  stripWorksheetDrawingsForExcelJs,
  workbookBufferForExcelJs,
  workbookDefinedNamesToModel,
  workbookHasPendingEdit,
  workbookHasUnsavedChanges,
  workbookImageAssetSource,
  workbookPersistenceKey,
  workbookShouldWarnBeforeUnload,
  workbookToModel,
  workbookViewModelFromExcel,
  workbookWindowRecordFromSheetViewState,
  workbookWindowViewsFromExcel,
  writePersistedWorkbookSnapshot,
};
