// Feature module; shared bindings are imported explicitly.
import {
  $,
  $selectionKeyboardSink,
  $sheetHost,
  $statusText,
  COLUMN_HEADER_HEIGHT,
  CWS_CLIPBOARD_DATA_MIME,
  CWS_CLIPBOARD_DATA_WEB_MIME,
  INTERNAL_CLIPBOARD_STORAGE_KEY,
  state,
} from "../../app-core.js";
import {
  repeatLastFindShortcut,
  selectCellsByPredicate,
} from "../data/query-editor.js";
import {
  cellHasDataForRegion,
  currentRegionFromPoint,
  openSelectedAutoFilterDropdown,
  sheetUsedRange,
  toggleAutoFilterForSelection,
} from "../data/sort-and-autofilter.js";
import {
  openSelectedValidationDropdown,
} from "../data/validation-and-consolidation.js";
import {
  cancelShapeDrawing,
  closeDataDialog,
  historyShortcutAction,
  insertChartFromSelection,
  insertChartSheetFromSelection,
  recommendedChartTypeForSelection,
  shouldHandleDialogHistoryShortcut,
} from "../drawing/charts-and-insertion.js";
import {
  fileExtension,
  fileExtensionFromImageMimeType,
  fileToDataUrl,
  imageDimensionsFromSource,
  insertWorksheetImageDescriptors,
  insertionPointForObjects,
  isLineShapeObject,
  normalizeShapeTextMargin,
} from "../drawing/shape-rendering.js";
import {
  moveSelectedImagesBy,
  primaryImagePositionElement,
  refreshShapeSource,
  remapPastedLineShapeConnections,
  setImagePixelBox,
} from "../drawing/transforms-and-connectors.js";
import {
  applyAutoFillMerges,
  clearSelectedWorksheetObjects,
  deleteSelectedWorksheetObjects,
  isEnterKeyEvent,
  openShapeTextEditor,
  positiveModulo,
  rangeContains,
  rangesEqualLoose,
  validateAutoFillMerges,
} from "./autofill.js";
import {
  canUseAsyncClipboardItemRead,
  canUseAsyncClipboardItemWrite,
  canUseAsyncClipboardRead,
  canUseAsyncClipboardWrite,
  clearPersistedInternalClipboard,
  clipboardDataFromEvent,
  clipboardDataHasCwsRangeMarker,
  clipboardDataLooksLikeCellPaste,
  clipboardTextPayloadLooksLikeCells,
  cwsClipboardHtmlForPayload,
  cwsClipboardPayload,
  cwsClipboardPayloadFromData,
  cwsClipboardPayloadFromHtml,
  cwsClipboardPayloadJson,
  isLocalStandaloneDocument,
  normalizeRestoredInternalClipboard,
  normalizeRestoredObjectClipboard,
  openLinkDialogForSelection,
  parseCwsClipboardPayload,
  persistInternalClipboard,
  richHtmlClipboardFromData,
  richHtmlClipboardFromPayload,
  uniqueWorksheetObjectId,
  workbookClipboardStorageSignature,
  writePlainTextToClipboardData,
} from "./ink-and-clipboard.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  cellElementForPoint,
  commitActiveCellValueToSelection,
  cssEscape,
  expandRangeForMerges,
  getCellRawInput,
  getDisplayForCell,
  isCellInRange,
  isSingleCellRange,
  normalizeRange,
  rangeToLabel,
  rangesToLabel,
  scheduleSelectionUiSync,
  selectableCellPoint,
  selectionRangeFromPoints,
  syncShapeFormatControls,
  toggleFormulaBarLines,
  unionMaterializedRanges,
  updateSelectionFrames,
  updateSelectionUi,
  withDisplayCalculationCache,
} from "./pointers-and-formula-evaluation.js";
import {
  activeSelectedImageIds,
  beginCellEdit,
  cellHasComment,
  columnSelectionRange,
  coordinateToPixels,
  editablePlainText,
  ensureCellElementForPoint,
  extendSelectionByKeyboard,
  imageBoxToPixels,
  isActiveImage,
  keyboardNavigationTargetFromEvent,
  keyboardSelectionExtensionTargetFromEvent,
  moveSelectionByKeyboard,
  placeCaretAtEnd,
  rawActiveSelectionRanges,
  rowSelectionRange,
  saveCellEditSelection,
  scrollRangeIntoView,
  selectCellByKeyboard,
  selectedImageObjectById,
  selectedImageObjects,
  selectedInkStrokes,
  selectionToolbarPoint,
  setSelectedImageIds,
  shapeTextPlainValue,
  startSelectedCellEdit,
  syncEditableCellAlignment,
  updateFormulaBarForCell,
  updateFormulaBarForSelection,
} from "./selection-and-cell-text.js";
import {
  captureCellsForHistory,
  captureSheetForHistory,
  captureSheetsForHistory,
  clearSelectionContents,
  cloneCellModel,
  cloneImages,
  cloneMerges,
  defaultColumnModel,
  defaultRowModel,
  isBlankClipboardCell,
  normalizePasteSpecialOptions,
  pasteClipboardCellAllExceptBorders,
  pasteClipboardCellAllForSpecial,
  pasteClipboardCellComments,
  pasteClipboardCellContentForSpecial,
  pasteClipboardCellFormat,
  pasteClipboardCellValidation,
  pasteClipboardColumnWidths,
  pasteInternalClipboardAtRange,
  pasteSpecialModeLabel,
  pasteSpecialModeWritesContent,
  pasteSpecialModeWritesFormula,
  pasteTextAtRange,
  pushHistory,
  redoLastOperation,
  sourceOffsetForPasteSpecial,
  undoLastOperation,
} from "./structure-and-history.js";
import {
  closeFormulaAutocomplete,
  insertTextIntoEditable,
  openInsertFunctionDialog,
  openNameManagerDialog,
  openPasteNameDialog,
  updateFormulaReferenceUi,
} from "../formula/dialogs-and-names.js";
import {
  applyBorderPreset,
  applyNumberFormatToSelection,
  clearFormatPainter,
  forEachCellInRange,
  rangeIntersects,
  rangesIntersect,
  removeIntersectingMerges,
  renderSheet,
  toggleStyleForSelection,
  toggleTextDecorationForSelection,
} from "../rendering/grid-and-canvas.js";
import {
  clipboardCellEntries,
  clipboardHasFormula,
  isRangeInsideRange,
  openCellStructureDialog,
  openFormatCellsDialog,
  setHeaderHidden,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  columnName,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  cancelObjectSelectionModeDrag,
  openFindDialog,
  openGoToDialog,
  selectCurrentRegionFromHomeCommand,
  selectDifferencesFromSelection,
  selectFormulaReferencesFromSelection,
  selectVisibleCellsFromSelection,
  uniqueSelectionRanges,
  writeHomeFillCells,
} from "../ui/printing-and-home-tools.js";
import {
  activeSheetEditElement,
  applyAutoSumToSelection,
  ensureSheetSize,
  handleKeyboardContextMenuShortcut,
  handleRibbonAccessKeyShortcut,
  handleRibbonFocusShortcut,
  handleWorkbookPaneFocusShortcut,
  openCreateTableDialog,
  rowHeaderWidthForSheet,
  toggleRibbonCollapsed,
} from "../ui/ribbon-and-links.js";
import {
  ensureEditableRangeForProtection,
  ensureSheetProtectionPermission,
} from "../ui/sheet-protection-and-translation.js";
import {
  calculateActiveSheetNow,
  calculateWorkbookNow,
  openCreateNamesDialog,
  toggleActiveSheetObjectsVisibility,
  toggleActiveSheetOutlineSymbols,
  toggleShowFormulasFromRibbon,
} from "../ui/views-and-formula-tools.js";
import {
  addBlankSheet,
  buildFormulaEngine,
  fontFamilyForExcel,
  renderWorkbook,
  setActiveSheet,
  visibleSheetIndexes,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  activeSheet,
  getSheetId,
  openSaveAsDialog,
  quickSaveWorkbook,
  visibleGridRange,
} from "../workbook/saving-and-export.js";
import {
  openWorkbookFilePicker,
} from "../workbook/startup-and-loading.js";
import {
  applyTableRangeEditBehaviors,
  imageMimeType,
  tableRangeEditMayChangeMetadata,
} from "../workbook/tables-and-validation.js";
import {
  openPrintDialog,
  openSheetCodeDialog,
} from "../workbook/vba-and-printing.js";
import {
  hideWorkbookWindow,
  openMacrosDialog,
  rangeCellCount,
  switchWorkbookWindow,
  workbookWindowRecords,
} from "../workbook/views-and-macros.js";


function restoreCwsClipboardPayload(payload) {
  if (!payload || typeof payload !== "object") return false;
  if (payload.kind === "range") {
    const clipboard = normalizeRestoredInternalClipboard(payload.clipboard);
    if (!clipboard) return false;
    clipboard.mode = "copy";
    clipboard.sourceSheetIndex = -1;
    state.internalClipboard = clipboard;
    state.objectClipboard = null;
    clearConsumedCutObjectClipboardLabel();
    clearPersistedInternalClipboard();
    return true;
  }
  if (payload.kind === "objects") {
    const clipboard = normalizeRestoredObjectClipboard(payload.clipboard);
    if (!clipboard) return false;
    state.objectClipboard = clipboard;
    state.internalClipboard = null;
    clearPersistedInternalClipboard();
    clearConsumedCutObjectClipboardLabel();
    return true;
  }
  return false;
}



function cwsRangePayloadMatchesInternalClipboard(payload) {
  const local = state.internalClipboard;
  const remote = payload?.kind === "range" ? payload.clipboard : null;
  if (!local || !remote) return false;
  return (
    String(local.text ?? "") === String(remote.text ?? "") &&
    local.mode === remote.mode &&
    Number(local.sourceSheetIndex) === Number(remote.sourceSheetIndex) &&
    Number(local.rowCount) === Number(remote.rowCount) &&
    Number(local.colCount) === Number(remote.colCount) &&
    JSON.stringify(local.sourceRange || null) === JSON.stringify(remote.sourceRange || null)
  );
}



function cwsObjectPayloadMatchesObjectClipboard(payload) {
  const local = state.objectClipboard;
  const remote = payload?.kind === "objects" ? payload.clipboard : null;
  if (!local?.objects?.length || !remote?.objects?.length) return false;
  const localIds = (local.sourceIds || []).join("\u001f");
  const remoteIds = (remote.sourceIds || []).join("\u001f");
  return (
    local.mode === remote.mode &&
    Number(local.sourceSheetIndex) === Number(remote.sourceSheetIndex) &&
    local.objects.length === remote.objects.length &&
    localIds === remoteIds &&
    objectClipboardTextLabel(local) === objectClipboardTextLabel(remote)
  );
}



function pasteCwsClipboardPayload(payload, options = {}) {
  if (!restoreCwsClipboardPayload(payload)) return false;
  if (payload.kind === "objects") {
    cancelObjectPasteFallback();
    return pasteObjectClipboard(options);
  }
  const range = options.range || activeSelectionRange();
  if (!range) {
    setStatus("貼り付け先のセルを選択してください。");
    return true;
  }
  pasteInternalClipboardAtRange(range, options.keepSourceFormatting === false ? { keepSourceFormatting: false } : undefined);
  return true;
}



function pasteExternalRichClipboard(clipboard, options = {}) {
  if (!clipboard?.cellEntries?.length) return false;
  const range = options.range || activeSelectionRange();
  if (!range) {
    setStatus("貼り付け先のセルを選択してください。");
    return true;
  }
  state.internalClipboard = clipboard;
  state.objectClipboard = null;
  clearPersistedInternalClipboard();
  clearConsumedCutObjectClipboardLabel();
  cancelObjectPasteFallback();
  pasteInternalClipboardAtRange(range, {
    keepSourceFormatting: options.keepSourceFormatting !== false,
    applyDimensions: options.applyDimensions === true,
  });
  return true;
}



function restorePersistedInternalClipboard() {
  if (state.internalClipboard || !state.model) return Boolean(state.internalClipboard);
  try {
    const payload = JSON.parse(sessionStorage.getItem(INTERNAL_CLIPBOARD_STORAGE_KEY) || "null");
    if (!payload || payload.signature !== workbookClipboardStorageSignature()) return false;
    const clipboard = normalizeRestoredInternalClipboard(payload.clipboard);
    if (!clipboard) return false;
    state.internalClipboard = clipboard;
    updateSelectionFrames();
    return true;
  } catch (error) {
    clearPersistedInternalClipboard();
    return false;
  }
}



function shouldWriteAsyncClipboard(options, wroteClipboardEvent) {
  return options.systemClipboard === true && !wroteClipboardEvent && canUseAsyncClipboardWrite();
}



async function writeCwsPayloadToAsyncClipboard(payload, text = "") {
  const plainText = String(text ?? "");
  if (payload && canUseAsyncClipboardItemWrite()) {
    const json = cwsClipboardPayloadJson(payload);
    const html = cwsClipboardHtmlForPayload(payload, plainText);
    try {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/plain": new Blob([plainText], { type: "text/plain" }),
          "text/html": new Blob([html], { type: "text/html" }),
          [CWS_CLIPBOARD_DATA_WEB_MIME]: new Blob([json], { type: CWS_CLIPBOARD_DATA_MIME }),
        }),
      ]);
      return true;
    } catch (error) {
      try {
        await navigator.clipboard.write([
          new ClipboardItem({
            "text/plain": new Blob([plainText], { type: "text/plain" }),
            "text/html": new Blob([html], { type: "text/html" }),
          }),
        ]);
        return true;
      } catch (fallbackError) {
        console.warn("Rich clipboard write failed:", fallbackError);
      }
    }
  }
  if (canUseAsyncClipboardWrite()) {
    await navigator.clipboard.writeText(plainText);
    return true;
  }
  return false;
}



async function asyncClipboardCwsPayload() {
  if (!canUseAsyncClipboardItemRead()) return null;
  const items = await navigator.clipboard.read();
  for (const item of items || []) {
    const types = Array.from(item.types || []);
    const directType = types.find((type) => type === CWS_CLIPBOARD_DATA_MIME || type === CWS_CLIPBOARD_DATA_WEB_MIME);
    if (directType) {
      const blob = await item.getType(directType);
      const payload = parseCwsClipboardPayload(await blob.text());
      if (payload) return payload;
    }
    if (types.includes("text/html")) {
      const blob = await item.getType("text/html");
      const payload = cwsClipboardPayloadFromHtml(await blob.text());
      if (payload) return payload;
    }
  }
  return null;
}



async function asyncClipboardTextPayload() {
  if (!canUseAsyncClipboardItemRead()) return null;
  const payload = { text: "", html: "", types: [] };
  const items = await navigator.clipboard.read();
  for (const item of items || []) {
    const types = Array.from(item.types || []);
    payload.types.push(...types);
    if (!payload.html && types.includes("text/html")) {
      try {
        payload.html = await (await item.getType("text/html")).text();
      } catch (error) {
        // Keep scanning other clipboard items.
      }
    }
    if (!payload.text && types.includes("text/plain")) {
      try {
        payload.text = await (await item.getType("text/plain")).text();
      } catch (error) {
        // Keep scanning other clipboard items.
      }
    }
  }
  payload.types = Array.from(new Set(payload.types.map((type) => String(type || "").toLowerCase())));
  return payload;
}



async function asyncClipboardCellTextPayload() {
  let itemPayload = null;
  if (canUseAsyncClipboardItemRead()) {
    try {
      itemPayload = await asyncClipboardTextPayload();
      if (clipboardTextPayloadLooksLikeCells(itemPayload)) return itemPayload;
    } catch (error) {
      console.warn("Clipboard text item read failed:", error);
    }
  }
  if (canUseAsyncClipboardRead()) {
    try {
      const text = await navigator.clipboard.readText();
      const payload = {
        text,
        html: itemPayload?.html || "",
        types: itemPayload?.types?.length ? itemPayload.types : ["text/plain"],
      };
      if (clipboardTextPayloadLooksLikeCells(payload)) return payload;
    } catch (error) {
      console.warn("Clipboard text read failed:", error);
    }
  }
  return null;
}



async function asyncClipboardRichHtmlClipboard() {
  if (!canUseAsyncClipboardItemRead()) return null;
  const payload = await asyncClipboardTextPayload();
  return richHtmlClipboardFromPayload(payload);
}



function scheduleInternalPasteFallback() {
  if (!state.internalClipboard) return;
  const token = {};
  state.pendingPasteFallback = token;
  setTimeout(() => {
    if (state.pendingPasteFallback !== token) return;
    state.pendingPasteFallback = null;
    if (state.internalClipboard && activeSelectionRange()) {
      void pasteFromClipboard({ allowAsyncClipboard: false });
    }
  }, 80);
}



function scheduleObjectPasteFallback() {
  if (!state.objectClipboard?.objects?.length) return;
  const token = {};
  state.pendingObjectPasteFallback = token;
  setTimeout(() => {
    if (state.pendingObjectPasteFallback !== token) return;
    state.pendingObjectPasteFallback = null;
    if (!state.objectClipboard?.objects?.length || isTextEditingTarget(document.activeElement)) return;
    pasteObjectClipboard();
  }, 80);
}



function cancelObjectPasteFallback() {
  state.pendingObjectPasteFallback = null;
}



function clearConsumedCutObjectClipboardLabel() {
  state.consumedCutObjectClipboardLabel = null;
}



function clearWorkbookClipboardState() {
  state.internalClipboard = null;
  state.objectClipboard = null;
  cancelObjectPasteFallback();
  clearConsumedCutObjectClipboardLabel();
  clearPersistedInternalClipboard();
}



function markConsumedCutObjectClipboardLabel(label) {
  const text = String(label || "");
  state.consumedCutObjectClipboardLabel = text || null;
}



function isConsumedCutObjectClipboardText(text) {
  const label = state.consumedCutObjectClipboardLabel;
  return Boolean(label && String(text ?? "") === label);
}



function consumeStaleCutObjectClipboardText(text) {
  if (!isConsumedCutObjectClipboardText(text)) return false;
  setStatus("切り取ったオブジェクトは既に貼り付け済みです。");
  return true;
}



function clipboardTextForInternalCompare(value) {
  return String(value ?? "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\n$/, "");
}



function internalClipboardTextMatches(text, clipboard = state.internalClipboard) {
  if (!clipboard) return false;
  const expected = String(clipboard.text ?? "");
  const actual = String(text ?? "");
  return actual === expected || clipboardTextForInternalCompare(actual) === clipboardTextForInternalCompare(expected);
}



function shouldUseInternalClipboardPaste(text, clipboardData = null) {
  if (!state.internalClipboard) return false;
  return clipboardDataHasCwsRangeMarker(clipboardData) || internalClipboardTextMatches(text);
}



function objectClipboardTextLabel(clipboard = state.objectClipboard) {
  if (!clipboard?.objects?.length) return "";
  return String(clipboard.clipboardText || selectedObjectSummary(clipboard.objects) || "");
}



function markObjectClipboardSystemSynced(mode, sourceSheetIndex, sourceIds, synced) {
  const clipboard = state.objectClipboard;
  if (
    !clipboard?.objects?.length ||
    clipboard.mode !== mode ||
    clipboard.sourceSheetIndex !== sourceSheetIndex ||
    clipboard.sourceIds?.length !== sourceIds.length ||
    !sourceIds.every((id, index) => clipboard.sourceIds?.[index] === id)
  ) {
    return;
  }
  clipboard.systemClipboardSynced = Boolean(synced);
}



function shouldPreferUnsyncedLocalObjectClipboardPaste(clipboardData) {
  if (!state.objectClipboard?.objects?.length) return false;
  if (state.objectClipboard.systemClipboardSynced !== false) return false;
  if (!isLocalStandaloneDocument()) return false;
  if (clipboardDataHasCwsRangeMarker(clipboardData)) return false;
  return true;
}



async function copySelectedRange(options = {}) {
  const range = activeSelectionRange();
  if (!range) {
    setStatus("コピーするセルを選択してください。");
    return false;
  }

  const text = rangeToTsv(range);
  const mode = options.mode === "cut" ? "cut" : "copy";
  state.internalClipboard = captureRangeForClipboard(range, text, { mode });
  persistInternalClipboard();
  state.objectClipboard = null;
  clearConsumedCutObjectClipboardLabel();
  updateSelectionFrames();
  const payload = cwsClipboardPayload("range", state.internalClipboard, text);
  const wroteClipboardEvent = writePlainTextToClipboardData(options.clipboardData, text, { cwsRange: true, cwsPayload: payload });
  const statusText = `${rangeToLabel(range)} を${mode === "cut" ? "切り取り" : "コピー"}しました。`;
  setStatus(statusText);
  if (shouldWriteAsyncClipboard(options, wroteClipboardEvent)) {
    try {
      await writeCwsPayloadToAsyncClipboard(payload, text);
    } catch (error) {
      console.warn("Clipboard write failed:", error);
      if ($statusText.text() === statusText) {
        setStatus(`${statusText}（ブラウザーがシステム クリップボードへの書き込みを許可しませんでした。）`);
      }
    }
  }
  return true;
}



async function cutSelectedRange(options = {}) {
  const range = activeSelectionRange();
  if (!range) return false;
  if (!ensureEditableRangeForProtection(activeSheet(), range, "切り取り")) return false;
  return copySelectedRange({ ...options, mode: "cut" });
}



async function pasteFromClipboard(options = {}) {
  const range = activeSelectionRange();
  if (!range) {
    setStatus("貼り付け先のセルを選択してください。");
    return;
  }

  const preferRich = options.rich !== false;
  restorePersistedInternalClipboard();
  if (Object.prototype.hasOwnProperty.call(options, "text")) {
    const text = String(options.text ?? "");
    if (shouldUseInternalClipboardPaste(text) && state.internalClipboard) {
      pasteInternalClipboardAtRange(range, preferRich ? undefined : { keepSourceFormatting: false });
    } else if (text !== "") {
      if (consumeStaleCutObjectClipboardText(text)) return;
      clearWorkbookClipboardState();
      pasteTextAtRange(text, range);
    } else if (state.internalClipboard) {
      pasteInternalClipboardAtRange(range, preferRich ? undefined : { keepSourceFormatting: false });
    } else {
      setStatus("貼り付けるテキストがありません。");
    }
    return;
  }

  if (options.allowAsyncClipboard === true && canUseAsyncClipboardItemRead()) {
    try {
      const cwsPayload = await asyncClipboardCwsPayload();
      if (cwsPayload?.kind === "objects" && !cwsObjectPayloadMatchesObjectClipboard(cwsPayload)) {
        pasteCwsClipboardPayload(cwsPayload, { range });
        return;
      }
      if (cwsPayload?.kind === "range" && !cwsRangePayloadMatchesInternalClipboard(cwsPayload)) {
        pasteCwsClipboardPayload(cwsPayload, { range, keepSourceFormatting: preferRich });
        return;
      }
    } catch (error) {
      console.warn("CWS clipboard read failed:", error);
    }
  }

  if (preferRich && options.allowAsyncClipboard === true && canUseAsyncClipboardItemRead()) {
    try {
      const richClipboard = await asyncClipboardRichHtmlClipboard();
      if (richClipboard) {
        pasteExternalRichClipboard(richClipboard, { range, applyDimensions: true });
        return;
      }
    } catch (error) {
      console.warn("Rich HTML clipboard read failed:", error);
    }
  }

  if (preferRich && state.internalClipboard) {
    pasteInternalClipboardAtRange(range);
    return;
  }
  if (!preferRich && state.internalClipboard) {
    pasteInternalClipboardAtRange(range, { keepSourceFormatting: false });
    return;
  }

  if (options.allowAsyncClipboard === true && (canUseAsyncClipboardItemRead() || canUseAsyncClipboardRead())) {
    const cellTextPayload = await asyncClipboardCellTextPayload();
    if (cellTextPayload?.text) {
      if (consumeStaleCutObjectClipboardText(cellTextPayload.text)) return;
      clearWorkbookClipboardState();
      pasteTextAtRange(cellTextPayload.text, range);
      return;
    }
  }

  if (options.allowAsyncClipboardImages === true && canUseAsyncClipboardItemRead()) {
    try {
      const imageFiles = await asyncClipboardImageFiles();
      if (imageFiles.length) {
        await pasteClipboardImageFiles(imageFiles, { anchor: insertionPointForObjects() });
        return;
      }
    } catch (error) {
      console.warn("Clipboard image read failed:", error);
    }
  }

  if (options.allowAsyncClipboard === true && canUseAsyncClipboardRead()) {
    try {
      const text = await navigator.clipboard.readText();
      if (preferRich && shouldUseInternalClipboardPaste(text)) {
        pasteInternalClipboardAtRange(range);
      } else if (text !== "") {
        if (consumeStaleCutObjectClipboardText(text)) return;
        clearWorkbookClipboardState();
        pasteTextAtRange(text, range);
      } else {
        setStatus("貼り付けるテキストがありません。");
      }
    } catch (error) {
      console.warn("Clipboard read failed:", error);
      setStatus("ブラウザーがクリップボードの読み取りを許可していません。Ctrl+V で貼り付けてください。");
    }
    return;
  }

  setStatus("貼り付ける内容がありません。外部アプリから貼り付ける場合は Ctrl+V を使用してください。");
}



async function pasteFromRibbonButton() {
  restorePersistedInternalClipboard();
  const range = activeSelectionRange();
  if (canUseAsyncClipboardItemRead()) {
    try {
      const cwsPayload = await asyncClipboardCwsPayload();
      if (cwsPayload?.kind === "objects" && !cwsObjectPayloadMatchesObjectClipboard(cwsPayload)) {
        if (pasteCwsClipboardPayload(cwsPayload, { range })) return;
      }
      if (cwsPayload?.kind === "range" && !cwsRangePayloadMatchesInternalClipboard(cwsPayload)) {
        if (pasteCwsClipboardPayload(cwsPayload, { range })) return;
      }
    } catch (error) {
      console.warn("CWS clipboard read failed:", error);
    }
  }
  if (range && canUseAsyncClipboardItemRead()) {
    try {
      const richClipboard = await asyncClipboardRichHtmlClipboard();
      if (richClipboard) {
        pasteExternalRichClipboard(richClipboard, { range, applyDimensions: true });
        return;
      }
    } catch (error) {
      console.warn("Rich HTML clipboard read failed:", error);
    }
  }
  if (range && (canUseAsyncClipboardItemRead() || canUseAsyncClipboardRead())) {
    const cellTextPayload = await asyncClipboardCellTextPayload();
    if (cellTextPayload?.text) {
      if (consumeStaleCutObjectClipboardText(cellTextPayload.text)) return;
      clearWorkbookClipboardState();
      pasteTextAtRange(cellTextPayload.text, range);
      return;
    }
    if (canUseAsyncClipboardItemRead()) {
      try {
        const imageFiles = await asyncClipboardImageFiles();
        if (imageFiles.length) {
          await pasteClipboardImageFiles(imageFiles, { anchor: insertionPointForObjects() });
          return;
        }
      } catch (error) {
        console.warn("Clipboard image read failed:", error);
      }
    }
    if (canUseAsyncClipboardRead()) {
      try {
        const text = await navigator.clipboard.readText();
        if (text !== "") {
          const objectClipboardLabel = state.objectClipboard?.objects?.length
            ? selectedObjectSummary(state.objectClipboard.objects)
            : "";
          if (state.objectClipboard?.objects?.length && text === objectClipboardLabel) {
            pasteObjectClipboard();
            return;
          }
          if (state.internalClipboard && shouldUseInternalClipboardPaste(text)) {
            pasteInternalClipboardAtRange(range);
            return;
          }
          if (consumeStaleCutObjectClipboardText(text)) return;
          clearWorkbookClipboardState();
          pasteTextAtRange(text, range);
          return;
        }
      } catch (error) {
        console.warn("Clipboard text read failed:", error);
      }
    }
  }
  if (state.objectClipboard?.objects?.length && !state.internalClipboard) {
    pasteObjectClipboard();
    return;
  }
  if (range && state.internalClipboard) {
    pasteInternalClipboardAtRange(range);
    return;
  }
  await pasteFromClipboard({ allowAsyncClipboard: true, allowAsyncClipboardImages: true });
}



async function asyncClipboardImageFiles() {
  const items = await navigator.clipboard.read();
  const files = [];
  for (const item of items || []) {
    const type = (item.types || []).find((candidate) => /^image\//i.test(candidate));
    if (!type) continue;
    const blob = await item.getType(type);
    files.push(blobToClipboardImageFile(blob, files.length));
  }
  return files;
}



function blobToClipboardImageFile(blob, index = 0) {
  if (blob instanceof File) return blob;
  const extension = fileExtensionFromImageMimeType(blob?.type, "png");
  return new File([blob], `clipboard-image-${index + 1}.${extension}`, { type: blob?.type || imageMimeType(extension) });
}



function openPasteSpecialDialog() {
  restorePersistedInternalClipboard();
  if (!state.internalClipboard) {
    showMessageBox("貼り付けできるデータがありません。", { title: "Cht WebSheet" });
    return false;
  }
  if (!activeSelectionRange()) {
    showMessageBox("貼り付け先のセルを選択してください。", { title: "Cht WebSheet" });
    return false;
  }
  state.dataDialog = { type: "paste-special" };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.id = "dataDialog";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog paste-special-dialog" role="dialog" aria-modal="true" aria-labelledby="pasteSpecialTitle" data-data-dialog="paste-special">
      <div class="excel-mini-titlebar" id="pasteSpecialTitle">形式を選択して貼り付け</div>
      <div class="excel-mini-body paste-special-body">
        <fieldset class="paste-special-group">
          <legend>貼り付け</legend>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialMode" value="all" checked />すべて(A)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialMode" value="allUsingSourceTheme" />コピー元のテーマを使用してすべて(H)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialMode" value="formulas" />数式(F)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialMode" value="values" />値(V)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialMode" value="formats" />書式(T)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialMode" value="comments" />コメントとメモ(C)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialMode" value="validation" />入力規則(N)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialMode" value="allExceptBorders" />罫線を除くすべて(X)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialMode" value="columnWidths" />列幅(W)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialMode" value="formulasNumberFormats" />数式と数値の書式(R)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialMode" value="valuesNumberFormats" />値と数値の書式(U)</label>
        </fieldset>
        <fieldset class="paste-special-group paste-special-operation">
          <legend>演算</legend>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialOperation" value="none" checked />しない</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialOperation" value="add" />加算(D)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialOperation" value="subtract" />減算(S)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialOperation" value="multiply" />乗算(M)</label>
          <label class="data-dialog-check"><input type="radio" name="pasteSpecialOperation" value="divide" />除算(I)</label>
        </fieldset>
        <label class="data-dialog-check paste-special-check"><input type="checkbox" data-paste-special-field="skipBlanks" />空白セルを無視する(B)</label>
        <label class="data-dialog-check paste-special-check"><input type="checkbox" data-paste-special-field="transpose" />行/列の入れ替え(E)</label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="paste-special-link">リンク貼り付け(L)</button>
        <span class="excel-mini-action-spacer"></span>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector("input[name='pasteSpecialMode']")?.focus({ preventScroll: true });
  setStatus("形式を選択して貼り付けを開きました。");
  return true;
}



function applyPasteSpecialDialog(dialog) {
  const options = pasteSpecialOptionsFromDialog(dialog);
  const range = activeSelectionRange();
  closeDataDialog();
  pasteInternalClipboardSpecialAtRange(range, options);
}



function pasteSpecialOptionsFromDialog(dialog) {
  return {
    mode: dialog?.querySelector("input[name='pasteSpecialMode']:checked")?.value || "all",
    operation: dialog?.querySelector("input[name='pasteSpecialOperation']:checked")?.value || "none",
    skipBlanks: Boolean(dialog?.querySelector("[data-paste-special-field='skipBlanks']")?.checked),
    transpose: Boolean(dialog?.querySelector("[data-paste-special-field='transpose']")?.checked),
  };
}



async function copySelectedImages(options = {}) {
  const objects = selectedImageObjects();
  if (!objects.length) {
    setStatus("コピーするオブジェクトを選択してください。");
    return false;
  }
  const sheet = activeSheet();
  const mode = options.mode === "cut" ? "cut" : "copy";
  const sourceIds = objects.map((image) => image.id);
  const boxes = objects.map((image) => imageBoxToPixels(sheet, image)).filter(Boolean);
  const label = selectedObjectSummary(objects);
  state.objectClipboard = {
    type: "objects",
    mode,
    sourceSheetIndex: state.activeSheetIndex,
    sourceIds,
    primaryId: state.selectedImage?.id || objects[0].id,
    objects: cloneImages(objects),
    boxes,
    clipboardText: label,
    systemClipboardSynced: false,
  };
  state.internalClipboard = null;
  clearPersistedInternalClipboard();
  clearConsumedCutObjectClipboardLabel();
  renderSheet();
  updateSelectionUi();
  const actionText = mode === "cut" ? "切り取りました" : "コピーしました";
  const statusText = `${label}を${actionText}。`;
  const payload = cwsClipboardPayload("objects", state.objectClipboard, label);
  const wroteClipboardEvent = writePlainTextToClipboardData(options.clipboardData, label, { cwsPayload: payload });
  markObjectClipboardSystemSynced(mode, state.activeSheetIndex, sourceIds, wroteClipboardEvent);
  setStatus(statusText);
  if (shouldWriteAsyncClipboard(options, wroteClipboardEvent)) {
    try {
      await writeCwsPayloadToAsyncClipboard(payload, label);
      markObjectClipboardSystemSynced(mode, state.activeSheetIndex, sourceIds, true);
    } catch (error) {
      console.warn("Object clipboard write failed:", error);
      const currentClipboard = state.objectClipboard;
      const sameClipboard =
        currentClipboard?.type === "objects" &&
        currentClipboard.mode === mode &&
        currentClipboard.sourceSheetIndex === state.activeSheetIndex &&
        sourceIds.every((id, index) => currentClipboard.sourceIds?.[index] === id);
      if (sameClipboard && $statusText.text() === statusText) {
        setStatus(`${label}を${actionText}（ブラウザーのクリップボード権限がありません）。`);
      }
    }
  }
  return true;
}



async function cutSelectedImages(options = {}) {
  return copySelectedImages({ ...options, mode: "cut" });
}



function pasteObjectClipboard(options = {}) {
  const clipboard = state.objectClipboard;
  cancelObjectPasteFallback();
  if (!clipboard?.objects?.length) {
    setStatus("貼り付けるオブジェクトがありません。");
    return false;
  }
  const sheet = activeSheet();
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを貼り付けできません。")) return false;
  sheet.images = Array.isArray(sheet.images) ? sheet.images : [];
  const isCutPaste = clipboard.mode === "cut";
  const sourceSheetIndex = Number.isInteger(Number(clipboard.sourceSheetIndex))
    ? Number(clipboard.sourceSheetIndex)
    : state.activeSheetIndex;
  const sourceSheet = isCutPaste ? state.model?.sheets?.[sourceSheetIndex] || null : null;
  const cutClipboardLabel = isCutPaste ? selectedObjectSummary(clipboard.objects) : "";
  const historySheetIndexes = isCutPaste && sourceSheet
    ? Array.from(new Set([sourceSheetIndex, state.activeSheetIndex]))
    : [];
  const sheetsBefore = historySheetIndexes.length ? captureSheetsForHistory(historySheetIndexes) : null;
  const sheetBefore = sheetsBefore ? null : captureSheetForHistory(sheet);
  const boxes = clipboard.boxes?.length === clipboard.objects.length
    ? clipboard.boxes
    : clipboard.objects.map((image) => imageBoxToPixels(sheet, image)).filter(Boolean);
  if (!boxes.length) return false;
  const baseLeft = Math.min(...boxes.map((box) => box.left));
  const baseTop = Math.min(...boxes.map((box) => box.top));
  const target = objectPasteTargetPoint(sheet, boxes, options);
  const inserted = [];
  const pastedGroupIds = new Map();
  const pastedObjectIds = new Map();
  clipboard.objects.forEach((source, index) => {
    const sourceBox = boxes[index] || boxes[0];
    if (!sourceBox) return;
    const image = cloneImages([source])[0];
    const sourceId = image.id;
    image.id = uniqueWorksheetObjectId("image");
    if (sourceId) pastedObjectIds.set(String(sourceId), image.id);
    image.name = isCutPaste ? source.name || (source.shape ? "図形" : "画像") : copiedObjectName(source.name);
    if (image.groupId) {
      if (!pastedGroupIds.has(image.groupId)) pastedGroupIds.set(image.groupId, uniqueWorksheetObjectId("group"));
      image.groupId = pastedGroupIds.get(image.groupId);
    }
    const nextBox = {
      left: target.left + (sourceBox.left - baseLeft),
      top: target.top + (sourceBox.top - baseTop),
      width: sourceBox.width,
      height: sourceBox.height,
    };
    setImagePixelBox(sheet, image, nextBox);
    if (image.shape) refreshShapeSource(image);
    sheet.images.push(image);
    inserted.push(image);
  });
  if (!inserted.length) return false;
  remapPastedLineShapeConnections(inserted, pastedObjectIds, { preserveExternal: isCutPaste });

  if (isCutPaste) {
    if (sourceSheet) {
      sourceSheet.images = Array.isArray(sourceSheet.images) ? sourceSheet.images : [];
      const insertedIds = sourceSheet === sheet ? new Set(inserted.map((image) => image.id)) : new Set();
      const sourceIds = new Set(clipboard.sourceIds || []);
      sourceSheet.images = sourceSheet.images.filter((image) => insertedIds.has(image.id) || !sourceIds.has(image.id));
    }
    state.objectClipboard = null;
    markConsumedCutObjectClipboardLabel(cutClipboardLabel);
  } else {
    clearConsumedCutObjectClipboardLabel();
  }

  setSelectedImageIds(inserted.map((image) => image.id), inserted[0].id);
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  const historyEntry = {
    sheetIndex: state.activeSheetIndex,
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${selectedObjectSummary(inserted)} 貼り付け`,
  };
  if (sheetsBefore?.length) {
    historyEntry.sheetsBefore = sheetsBefore;
    historyEntry.sheetsAfter = captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index));
  } else {
    historyEntry.sheetBefore = sheetBefore;
    historyEntry.sheetAfter = captureSheetForHistory(sheet);
  }
  pushHistory(historyEntry);
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`${selectedObjectSummary(inserted)}を貼り付けました。`);
  return true;
}



function objectPasteTargetPoint(sheet, boxes, options = {}) {
  if (options.offset || activeSelectedImageIds().length) {
    return {
      left: Math.max(rowHeaderWidthForSheet(sheet), Math.min(...boxes.map((box) => box.left)) + 16),
      top: Math.max(COLUMN_HEADER_HEIGHT, Math.min(...boxes.map((box) => box.top)) + 16),
    };
  }
  const range = activeSelectionRange();
  if (range) {
    return {
      left: coordinateToPixels(sheet, "col", range.left - 1),
      top: coordinateToPixels(sheet, "row", range.top - 1),
    };
  }
  return {
    left: Math.max(rowHeaderWidthForSheet(sheet), Math.min(...boxes.map((box) => box.left)) + 16),
    top: Math.max(COLUMN_HEADER_HEIGHT, Math.min(...boxes.map((box) => box.top)) + 16),
  };
}



function copiedObjectName(name) {
  const base = String(name || "オブジェクト").trim() || "オブジェクト";
  return base.endsWith(" のコピー") ? base : `${base} のコピー`;
}



function selectedObjectSummary(objects) {
  if (!objects?.length) return "オブジェクト";
  if (objects.length === 1) return objects[0].name || (objects[0].shape ? "図形" : "画像");
  return `${objects.length} 個のオブジェクト`;
}



function handleWorkbookCopyEvent(event) {
  if (!state.model || isTextEditingTarget(event.target)) return;
  const clipboardData = clipboardDataFromEvent(event);
  const objects = selectedImageObjects();
  if (objects.length) {
    event.preventDefault();
    void copySelectedImages({ clipboardData, systemClipboard: false });
    return;
  }
  if (!activeSelectionRange()) return;
  event.preventDefault();
  void copySelectedRange({ clipboardData, systemClipboard: false });
}



function handleWorkbookCutEvent(event) {
  if (!state.model || isTextEditingTarget(event.target)) return;
  const clipboardData = clipboardDataFromEvent(event);
  const objects = selectedImageObjects();
  if (objects.length) {
    event.preventDefault();
    void cutSelectedImages({ clipboardData, systemClipboard: false });
    return;
  }
  if (!activeSelectionRange()) return;
  event.preventDefault();
  void cutSelectedRange({ clipboardData, systemClipboard: false });
}



function handleWorkbookPasteEvent(event) {
  if (!state.model || isTextEditingTarget(event.target)) return;
  state.pendingPasteFallback = null;
  const clipboardData = clipboardDataFromEvent(event);
  const text = clipboardData?.getData("text/plain");
  const cellTextPaste = clipboardDataLooksLikeCellPaste(clipboardData, text);
  const richClipboard = richHtmlClipboardFromData(clipboardData, text);
  const imageFiles = clipboardImageFilesFromData(clipboardData);
  restorePersistedInternalClipboard();
  const cwsPayload = cwsClipboardPayloadFromData(clipboardData);
  if (cwsPayload?.kind === "objects" && !cwsObjectPayloadMatchesObjectClipboard(cwsPayload)) {
    event.preventDefault();
    cancelObjectPasteFallback();
    pasteCwsClipboardPayload(cwsPayload);
    return;
  }
  const objectClipboardLabel = objectClipboardTextLabel();
  if (state.objectClipboard?.objects?.length && text === objectClipboardLabel) {
    event.preventDefault();
    cancelObjectPasteFallback();
    pasteObjectClipboard();
    return;
  }
  if (shouldPreferUnsyncedLocalObjectClipboardPaste(clipboardData)) {
    event.preventDefault();
    cancelObjectPasteFallback();
    pasteObjectClipboard();
    return;
  }
  if (cwsPayload?.kind === "objects") {
    event.preventDefault();
    cancelObjectPasteFallback();
    pasteObjectClipboard();
    return;
  }
  if (imageFiles.length && !cellTextPaste) {
    event.preventDefault();
    cancelObjectPasteFallback();
    void pasteClipboardImageFiles(imageFiles, { anchor: insertionPointForObjects() });
    return;
  }
  const range = activeSelectionRange();
  if (!range) return;
  if (richClipboard) {
    event.preventDefault();
    cancelObjectPasteFallback();
    pasteExternalRichClipboard(richClipboard, { range, applyDimensions: true });
    return;
  }
  if (cwsPayload?.kind === "range" && !cwsRangePayloadMatchesInternalClipboard(cwsPayload)) {
    event.preventDefault();
    cancelObjectPasteFallback();
    pasteCwsClipboardPayload(cwsPayload, { range });
    return;
  }
  if (state.internalClipboard && shouldUseInternalClipboardPaste(text, clipboardData)) {
    event.preventDefault();
    cancelObjectPasteFallback();
    pasteInternalClipboardAtRange(range);
    return;
  }
  if (cwsPayload) {
    event.preventDefault();
    cancelObjectPasteFallback();
    pasteCwsClipboardPayload(cwsPayload, { range });
    return;
  }
  if (text != null && text !== "") {
    event.preventDefault();
    cancelObjectPasteFallback();
    if (consumeStaleCutObjectClipboardText(text)) return;
    clearWorkbookClipboardState();
    pasteTextAtRange(text, range);
    return;
  }
  if (state.internalClipboard) {
    event.preventDefault();
    cancelObjectPasteFallback();
    pasteInternalClipboardAtRange(range);
  }
}



function clipboardImageFilesFromData(clipboardData) {
  if (!clipboardData) return [];
  const files = [];
  const seen = new Set();
  const addFile = (file) => {
    if (!isClipboardImageFile(file)) return;
    const key = `${file.name || ""}\u001f${file.type || ""}\u001f${file.size || 0}\u001f${file.lastModified || 0}`;
    if (seen.has(key)) return;
    seen.add(key);
    files.push(file);
  };
  Array.from(clipboardData.items || []).forEach((item) => {
    if (item?.kind !== "file" || !/^image\//i.test(item.type || "")) return;
    const file = item.getAsFile?.();
    addFile(file);
  });
  Array.from(clipboardData.files || []).forEach(addFile);
  return files;
}



function isClipboardImageFile(file) {
  if (!file || Number(file.size || 0) <= 0) return false;
  if (/^image\//i.test(file.type || "")) return true;
  return /\.(png|jpe?g|gif|webp|bmp|svg|ico)$/i.test(file.name || "");
}



async function pasteClipboardImageFiles(files, options = {}) {
  const imageFiles = (files || []).filter(isClipboardImageFile);
  if (!state.model || !imageFiles.length) return false;
  const descriptors = [];
  const seenImagePayloads = new Set();
  for (const file of imageFiles) {
    try {
      const src = await fileToDataUrl(file);
      const fingerprint = clipboardImageSourceFingerprint(src);
      if (seenImagePayloads.has(fingerprint)) continue;
      seenImagePayloads.add(fingerprint);
      const size = await imageDimensionsFromSource(src);
      const displaySize = clipboardImageDisplaySize(size);
      const extension = file.name ? fileExtension(file.name) : fileExtensionFromImageMimeType(file.type);
      descriptors.push({
        name: file.name || `貼り付けた画像 ${descriptors.length + 1}`,
        src,
        mimeType: file.type || imageMimeType(extension),
        extension,
        width: displaySize.width,
        height: displaySize.height,
        preserveDisplaySize: true,
      });
    } catch (error) {
      console.warn("Clipboard image paste failed for one item:", error);
    }
  }
  if (!descriptors.length) {
    setStatus("クリップボードの画像を読み込めませんでした。");
    return false;
  }
  state.internalClipboard = null;
  state.objectClipboard = null;
  clearConsumedCutObjectClipboardLabel();
  clearPersistedInternalClipboard();
  insertWorksheetImageDescriptors(descriptors, {
    label: "画像",
    anchor: options.anchor || insertionPointForObjects(),
  });
  return true;
}



function clipboardImageDisplaySize(size) {
  const width = Math.max(1, Number(size?.width) || 220);
  const height = Math.max(1, Number(size?.height) || 140);
  const scale = clipboardImagePixelScale();
  return {
    width: Math.max(1, Math.round(width / scale)),
    height: Math.max(1, Math.round(height / scale)),
  };
}



function clipboardImagePixelScale() {
  const ratio = Number(window.devicePixelRatio || 1);
  if (!Number.isFinite(ratio) || ratio <= 1.25) return 1;
  return Math.max(1, Math.min(3, ratio));
}



function clipboardImageSourceFingerprint(src) {
  const text = String(src || "");
  const commaIndex = text.indexOf(",");
  const payload = commaIndex >= 0 ? text.slice(commaIndex + 1) : text;
  let hash = 2166136261;
  for (let index = 0; index < payload.length; index += 1) {
    hash ^= payload.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return `${payload.length}:${(hash >>> 0).toString(16)}`;
}



function executeWorkbookClipboardShortcut(command, fallback) {
  if (state.model && !isTextEditingTarget(document.activeElement) && !isLocalStandaloneDocument()) {
    fallback?.();
    return;
  }
  const previousCellClipboard = state.internalClipboard;
  const previousObjectClipboard = state.objectClipboard;
  let handled = false;
  try {
    handled = document.execCommand?.(command) === true;
  } catch (error) {
    console.warn(`Clipboard ${command} command failed:`, error);
  }
  const updatedInternalClipboard =
    state.internalClipboard !== previousCellClipboard ||
    state.objectClipboard !== previousObjectClipboard;
  if (!handled || !updatedInternalClipboard) fallback?.();
}



function handleWorkbookKeyboardShortcut(event) {
  if (event.defaultPrevented) return;
  const key = String(event.key || "").toLowerCase();
  const isModifier = event.ctrlKey || event.metaKey;

  if (handleRibbonAccessKeyShortcut(event, key)) return;
  if (handleRibbonFocusShortcut(event, key)) return;
  if (handleWorkbookPaneFocusShortcut(event, key)) return;
  if (handleKeyboardContextMenuShortcut(event, key)) return;

  if (key === "escape" && state.objectSelectionMode) {
    event.preventDefault();
    cancelObjectSelectionModeDrag();
    return;
  }

  if (key === "escape" && state.formulaAutocomplete) {
    event.preventDefault();
    closeFormulaAutocomplete({ restoreFocus: false });
    return;
  }

  if (isModifier && !event.altKey && key === "s") {
    event.preventDefault();
    if (!state.model) {
      setStatus("先に Excel ファイルを開いてください。");
      return;
    }
    void quickSaveWorkbook();
    return;
  }

  if (isModifier && !event.altKey && key === "o") {
    event.preventDefault();
    void openWorkbookFilePicker();
    return;
  }

  if (state.model && isModifier && !event.altKey && key === "p") {
    event.preventDefault();
    openPrintDialog();
    return;
  }

  if (state.model && isModifier && !event.altKey && key === "f2") {
    event.preventDefault();
    openPrintDialog();
    return;
  }

  if (state.model && isModifier && !event.altKey && key === "f1") {
    event.preventDefault();
    toggleRibbonCollapsed();
    return;
  }

  if (state.model && isModifier && !event.altKey && key === "w") {
    event.preventDefault();
    if (event.target?.closest?.("#dataDialog, #messageBox, #findDialog, #goToDialog, #specialDialog")) return;
    hideWorkbookWindow();
    return;
  }

  if (state.model && event.ctrlKey && !event.metaKey && !event.altKey && key === "f4") {
    event.preventDefault();
    hideWorkbookWindow();
    return;
  }

  if (state.model && !event.ctrlKey && !event.metaKey && event.altKey && !event.shiftKey && key === "f4") {
    event.preventDefault();
    hideWorkbookWindow();
    return;
  }

  if (state.model && !event.ctrlKey && !event.metaKey && event.altKey && !event.shiftKey && key === "f8") {
    event.preventDefault();
    openMacrosDialog();
    return;
  }

  if (state.model && !event.ctrlKey && !event.metaKey && event.altKey && !event.shiftKey && key === "f11") {
    event.preventDefault();
    openSheetCodeDialog();
    return;
  }

  if (state.model && !event.ctrlKey && !event.metaKey && !event.altKey && event.shiftKey && key === "f3") {
    event.preventDefault();
    openInsertFunctionDialog();
    return;
  }

  if (state.model && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey && key === "f3") {
    event.preventDefault();
    openPasteNameDialog();
    return;
  }

  if (state.model && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey && key === "f12") {
    event.preventDefault();
    openSaveAsDialog();
    return;
  }

  if (state.model && !event.ctrlKey && !event.metaKey && !event.altKey && event.shiftKey && key === "f5") {
    event.preventDefault();
    openFindDialog();
    return;
  }

  if (state.model && !event.ctrlKey && !event.metaKey && !event.altKey && event.shiftKey && key === "f4") {
    event.preventDefault();
    repeatLastFindShortcut();
    return;
  }

  if (state.model && !event.ctrlKey && !event.metaKey && !event.altKey && key === "f9") {
    event.preventDefault();
    if (event.shiftKey) calculateActiveSheetNow();
    else calculateWorkbookNow();
    return;
  }

  if (state.model && !event.altKey && isModifier && (key === "f" || key === "h" || key === "g")) {
    event.preventDefault();
    if (key === "f") openFindDialog();
    else if (key === "h") openFindDialog({ mode: "replace" });
    else openGoToDialog();
    return;
  }
  if (state.model && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey && key === "f5") {
    event.preventDefault();
    openGoToDialog();
    return;
  }

  if (state.model && shouldHandleDialogHistoryShortcut(event, key, isModifier)) {
    event.preventDefault();
    if (key === "z") undoLastOperation();
    else redoLastOperation();
    return;
  }

  if (!state.model || isTextEditingTarget(event.target)) return;

  const sheetShortcutTarget = isSheetSelectionShortcutTarget(event.target);
  if (sheetShortcutTarget && isModifier && !event.altKey && key === "f6") {
    event.preventDefault();
    activateAdjacentWorkbookWindow(event.shiftKey ? -1 : 1);
    return;
  }
  if (sheetShortcutTarget && isModifier && !event.altKey && (key === "pageup" || key === "pagedown")) {
    event.preventDefault();
    if (event.shiftKey) activateAdjacentVisibleSheetGroup(key === "pagedown" ? 1 : -1);
    else activateAdjacentVisibleSheet(key === "pagedown" ? 1 : -1);
    return;
  }
  if (sheetShortcutTarget && !event.ctrlKey && !event.metaKey && !event.altKey && event.shiftKey && key === "f11") {
    event.preventDefault();
    addBlankSheet();
    return;
  }
  if (sheetShortcutTarget && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey && key === "f11") {
    event.preventDefault();
    insertChartSheetFromSelection(recommendedChartTypeForSelection());
    return;
  }
  if (sheetShortcutTarget && !event.ctrlKey && !event.metaKey && event.altKey && !event.shiftKey && key === "f1") {
    event.preventDefault();
    insertChartFromSelection(recommendedChartTypeForSelection());
    return;
  }
  if (sheetShortcutTarget && isModifier && !event.altKey && !event.shiftKey && key === "a") {
    event.preventDefault();
    selectAllByKeyboard();
    return;
  }
  if (sheetShortcutTarget && !event.ctrlKey && !event.metaKey && !event.altKey && event.shiftKey && isSpaceKeyEvent(event)) {
    event.preventDefault();
    selectRowsByKeyboard();
    return;
  }
  if (sheetShortcutTarget && isModifier && !event.altKey && !event.shiftKey && isSpaceKeyEvent(event)) {
    event.preventDefault();
    selectColumnsByKeyboard();
    return;
  }
  if (sheetShortcutTarget && isModifier && !event.altKey && !event.shiftKey && key === "enter") {
    event.preventDefault();
    commitActiveCellValueToSelection();
    return;
  }

  if (key === "escape" && (state.pendingShapeType || state.shapeDrawing)) {
    event.preventDefault();
    cancelShapeDrawing();
    return;
  }
  if (key === "escape" && state.formatPainter) {
    event.preventDefault();
    clearFormatPainter();
    return;
  }
  if (event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey && key === "arrowdown") {
    if (openSelectedValidationDropdown() || openSelectedAutoFilterDropdown()) {
      event.preventDefault();
      return;
    }
  }
  if (sheetShortcutTarget && handleWorkbookEditingShortcut(event, key, isModifier)) return;
  if (handleSelectedObjectKeyboardShortcut(event)) return;
  if (handleSelectionKeyboardNavigation(event)) return;
  if (handleSelectionDirectTyping(event)) return;

  if (key === "escape" && state.internalClipboard) {
    event.preventDefault();
    state.internalClipboard = null;
    clearPersistedInternalClipboard();
    renderSheet();
    setStatus("コピー状態をキャンセルしました。");
    return;
  }
  if (key === "delete" && !event.ctrlKey && !event.metaKey && !event.altKey) {
    event.preventDefault();
    clearSelectionContents();
    return;
  }
  if (key === "backspace" && !event.ctrlKey && !event.metaKey && !event.altKey) {
    event.preventDefault();
    clearSelectionAndStartEdit();
    return;
  }

  if (sheetShortcutTarget && isModifier && event.altKey && !event.shiftKey && key === "v") {
    event.preventDefault();
    openPasteSpecialDialog();
    return;
  }

  if (!isModifier || event.altKey) return;

  if (key === "c") {
    event.preventDefault();
    executeWorkbookClipboardShortcut("copy", () => void copySelectedRange({ systemClipboard: true }));
    return;
  }
  if (key === "x") {
    event.preventDefault();
    executeWorkbookClipboardShortcut("cut", () => void cutSelectedRange({ systemClipboard: true }));
    return;
  }
  if (key === "v") {
    restorePersistedInternalClipboard();
    if (state.internalClipboard) {
      scheduleInternalPasteFallback();
    }
    if (state.objectClipboard?.objects?.length) {
      scheduleObjectPasteFallback();
    }
    return;
  }
  const historyAction = historyShortcutAction(event, key, isModifier);
  if (historyAction) {
    event.preventDefault();
    if (historyAction === "undo") undoLastOperation();
    else redoLastOperation();
    return;
  }
}



function handleWorkbookEditingShortcut(event, key, isModifier) {
  const dateTimeShortcut = keyboardDateTimeShortcutKind(event, key);
  if (dateTimeShortcut) {
    event.preventDefault();
    insertKeyboardDateTimeByKeyboard(dateTimeShortcut);
    return true;
  }
  if (event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey && key === "=") {
    event.preventDefault();
    applyAutoSumToSelection("SUM");
    return true;
  }
  if (event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey && isSemicolonShortcutKey(event, key)) {
    event.preventDefault();
    selectVisibleCellsFromSelection();
    return true;
  }
  if (!isModifier || event.altKey) return false;

  const code = String(event.code || "");
  if (!event.shiftKey) {
    if (isAsteriskShortcutKey(event, key)) {
      event.preventDefault();
      selectCurrentRegionFromHomeCommand();
      return true;
    }
    if (isBackquoteShortcutKey(event, key)) {
      event.preventDefault();
      toggleShowFormulasFromRibbon();
      return true;
    }
    if (key === "b") {
      event.preventDefault();
      toggleStyleForSelection("fontWeight", "700", "normal", "太字");
      return true;
    }
    if (key === "i") {
      event.preventDefault();
      toggleStyleForSelection("fontStyle", "italic", "normal", "斜体");
      return true;
    }
    if (key === "u") {
      event.preventDefault();
      toggleTextDecorationForSelection("underline", "下線");
      return true;
    }
    if (key === "2" || code === "Digit2" || code === "Numpad2") {
      event.preventDefault();
      toggleStyleForSelection("fontWeight", "700", "normal", "太字");
      return true;
    }
    if (key === "3" || code === "Digit3" || code === "Numpad3") {
      event.preventDefault();
      toggleStyleForSelection("fontStyle", "italic", "normal", "斜体");
      return true;
    }
    if (key === "4" || code === "Digit4" || code === "Numpad4") {
      event.preventDefault();
      toggleTextDecorationForSelection("underline", "下線");
      return true;
    }
    if (key === "k") {
      event.preventDefault();
      openLinkDialogForSelection();
      return true;
    }
    if (key === "t" || key === "l") {
      event.preventDefault();
      openCreateTableDialog();
      return true;
    }
    if (key === "[") {
      event.preventDefault();
      selectFormulaReferencesFromSelection("precedents");
      return true;
    }
    if (key === "]") {
      event.preventDefault();
      selectFormulaReferencesFromSelection("dependents");
      return true;
    }
    if (isBackslashShortcutKey(event, key)) {
      event.preventDefault();
      selectDifferencesFromSelection("row");
      return true;
    }
    if (key === "f3") {
      event.preventDefault();
      openNameManagerDialog();
      return true;
    }
    if (key === "5" || code === "Digit5" || code === "Numpad5") {
      event.preventDefault();
      toggleTextDecorationForSelection("line-through", "取り消し線");
      return true;
    }
    if (key === "6" || code === "Digit6" || code === "Numpad6") {
      event.preventDefault();
      toggleActiveSheetObjectsVisibility();
      return true;
    }
    if (key === "8" || code === "Digit8" || code === "Numpad8") {
      event.preventDefault();
      toggleActiveSheetOutlineSymbols();
      return true;
    }
    if (key === "1" || code === "Digit1" || code === "Numpad1") {
      event.preventDefault();
      openFormatCellsDialogForShortcut();
      return true;
    }
    if (key === "9" || code === "Digit9" || code === "Numpad9") {
      event.preventDefault();
      setHeaderHidden("row", true, state.selected?.row, state.selected?.col);
      return true;
    }
    if (key === "0" || code === "Digit0" || code === "Numpad0") {
      event.preventDefault();
      setHeaderHidden("column", true, state.selected?.row, state.selected?.col);
      return true;
    }
    if (key === "-" || code === "Minus" || code === "NumpadSubtract") {
      event.preventDefault();
      openCellStructureDialogForShortcut("delete");
      return true;
    }
    if (key === "d") {
      event.preventDefault();
      fillSelectionByKeyboard("down");
      return true;
    }
    if (key === "r") {
      event.preventDefault();
      fillSelectionByKeyboard("right");
      return true;
    }
  }

  if (event.shiftKey) {
    if (isSpaceKeyEvent(event) || isAsteriskShortcutKey(event, key) || key === "*" || code === "Digit8") {
      event.preventDefault();
      selectCurrentRegionFromHomeCommand();
      return true;
    }
    if (isBackslashShortcutKey(event, key)) {
      event.preventDefault();
      selectDifferencesFromSelection("column");
      return true;
    }
    if (key === "+" || key === "=" || code === "Equal" || code === "NumpadAdd") {
      event.preventDefault();
      openCellStructureDialogForShortcut("insert");
      return true;
    }
    if (key === "~" || isBackquoteShortcutKey(event, key)) {
      event.preventDefault();
      applyNumberFormatToSelection("", "標準");
      return true;
    }
    if (key === "!" || code === "Digit1") {
      event.preventDefault();
      applyNumberFormatToSelection("#,##0.00", "数値");
      return true;
    }
    if (key === "@" || code === "Digit2") {
      event.preventDefault();
      applyNumberFormatToSelection("h:mm", "時刻");
      return true;
    }
    if (key === "#" || code === "Digit3") {
      event.preventDefault();
      applyNumberFormatToSelection("yyyy/m/d", "日付");
      return true;
    }
    if (key === "$" || code === "Digit4") {
      event.preventDefault();
      applyNumberFormatToSelection("¥#,##0.00", "通貨");
      return true;
    }
    if (key === "%" || code === "Digit5") {
      event.preventDefault();
      applyNumberFormatToSelection("0%", "パーセント");
      return true;
    }
    if (key === "^" || code === "Digit6") {
      event.preventDefault();
      applyNumberFormatToSelection("0.00E+00", "指数");
      return true;
    }
    if (key === "&" || code === "Digit7") {
      event.preventDefault();
      applyBorderPreset("outside");
      return true;
    }
    if (key === "_" || code === "Minus") {
      event.preventDefault();
      applyBorderPreset("remove-outside");
      return true;
    }
    if (key === "l") {
      event.preventDefault();
      toggleAutoFilterForSelection();
      return true;
    }
    if (key === "u") {
      event.preventDefault();
      toggleFormulaBarLines();
      return true;
    }
    if (key === "o") {
      event.preventDefault();
      selectCellsByPredicate("コメント", cellHasComment);
      return true;
    }
    if (key === "f3") {
      event.preventDefault();
      openCreateNamesDialog();
      return true;
    }
    if (key === "(" || key === "9" || code === "Digit9" || code === "Numpad9") {
      event.preventDefault();
      setHeaderHidden("row", false, state.selected?.row, state.selected?.col);
      return true;
    }
    if (key === ")" || key === "0" || code === "Digit0" || code === "Numpad0") {
      event.preventDefault();
      setHeaderHidden("column", false, state.selected?.row, state.selected?.col);
      return true;
    }
  }

  return false;
}



function keyboardDateTimeShortcutKind(event, key = String(event?.key || "").toLowerCase()) {
  const isModifier = Boolean(event?.ctrlKey || event?.metaKey);
  if (!isModifier || event?.altKey) return "";
  if (isColonShortcutKey(event, key)) return "time";
  if (event?.shiftKey && isSemicolonShortcutKey(event, key)) return "time";
  if (!event?.shiftKey && isSemicolonShortcutKey(event, key)) return "date";
  return "";
}



function isSemicolonShortcutKey(event, key) {
  return key === ";" || key === "；" || String(event.code || "") === "Semicolon";
}



function isColonShortcutKey(event, key) {
  return key === ":" || key === "：" || String(event.code || "") === "Colon";
}



function isBackquoteShortcutKey(event, key) {
  return key === "`" || String(event.code || "") === "Backquote";
}



function isBackslashShortcutKey(event, key) {
  return key === "\\" || String(event.code || "") === "Backslash";
}



function isAsteriskShortcutKey(event, key) {
  const code = String(event.code || "");
  return key === "*" || code === "NumpadMultiply";
}



function openFormatCellsDialogForShortcut() {
  const point = selectionToolbarPoint() || state.selected;
  if (!point) {
    setStatus("セルを選択してください。");
    return;
  }
  openFormatCellsDialog(point.row, point.col);
}



function openCellStructureDialogForShortcut(kind) {
  const point = selectionToolbarPoint() || state.selected;
  if (!point) {
    setStatus("セルを選択してください。");
    return;
  }
  openCellStructureDialog(kind, point.row, point.col);
}



function insertKeyboardDateTimeByKeyboard(kind) {
  const point = selectionToolbarPoint() || state.selected;
  if (!point) {
    setStatus("セルを選択してください。");
    return;
  }
  const value = keyboardDateTimeShortcutText(kind, new Date());
  if (!value) return;
  if (state.selected?.sheetIndex !== state.activeSheetIndex || state.selected.row !== point.row || state.selected.col !== point.col) {
    state.selected = { sheetIndex: state.activeSheetIndex, row: point.row, col: point.col };
    state.selectionRange = selectionRangeFromPoints(state.selected, state.selected);
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
  }
  const editor = startSelectedCellEdit({ initialText: value, replaceText: true, caretAtEnd: true });
  if (!editor) return;
  setStatus(`${columnName(point.col)}${point.row} に${kind === "time" ? "時刻" : "日付"}を入力しました。`);
}



function insertKeyboardDateTimeIntoEditable(editor, kind) {
  if (!editor) return false;
  const value = keyboardDateTimeShortcutText(kind, new Date());
  if (!value) return false;
  const inserted = insertTextIntoEditable(editor, value);
  if (!inserted) return false;
  syncEditableCellAlignment(editor);
  saveCellEditSelection(editor);
  updateFormulaReferenceUi();
  setStatus(`${kind === "time" ? "時刻" : "日付"}を挿入しました。`);
  return true;
}



function keyboardDateTimeShortcutText(kind, date = new Date()) {
  if (kind === "date") return formatKeyboardShortcutDate(date);
  if (kind === "time") return formatKeyboardShortcutTime(date);
  return "";
}



function formatKeyboardShortcutDate(date) {
  return `${date.getFullYear()}/${pad2(date.getMonth() + 1)}/${pad2(date.getDate())}`;
}



function formatKeyboardShortcutTime(date) {
  return `${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
}



function pad2(value) {
  return String(value).padStart(2, "0");
}



function fillSelectionByKeyboard(direction) {
  const fillDirection = ["down", "right", "up", "left"].includes(direction) ? direction : "down";
  const ranges = activeSelectionRanges();
  if (ranges.length !== 1) {
    showMessageBox("フィルを実行する範囲を1つ選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const range = ranges[0];
  if (range.fullRow || range.fullColumn) {
    showMessageBox("行全体または列全体の選択範囲には、このフィルを実行できません。", { title: "Cht WebSheet" });
    return;
  }
  const sheet = activeSheet();
  const plan = homeFillKeyboardPlanForRange(range, fillDirection);
  if (!plan) return;
  const { sourceRange, fillRange } = plan;

  if (!ensureEditableRangeForProtection(sheet, fillRange, "オートフィル")) return;
  const mergeValidation = validateAutoFillMerges(sheet, sourceRange, fillRange, fillDirection);
  if (!mergeValidation.ok) {
    showMessageBox(mergeValidation.message, { title: "Cht WebSheet" });
    setStatus(mergeValidation.message);
    return;
  }

  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = rawActiveSelectionRanges().map((item) => ({ ...item }));
  const before = captureCellsForHistory(sheet, fillRange);
  const mergesBefore = cloneMerges(sheet.merges);
  ensureSheetSize(sheet, fillRange.bottom, fillRange.right);
  removeIntersectingMerges(sheet, fillRange);
  writeHomeFillCells(sheet, sourceRange, fillRange, fillDirection);
  applyAutoFillMerges(sheet, sourceRange, fillRange, fillDirection);
  state.hf = buildFormulaEngine(state.model);
  const after = captureCellsForHistory(sheet, fillRange);
  const mergesAfter = cloneMerges(sheet.merges);

  state.selected = { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
  state.selectionRange = { ...range };
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    mergesBefore,
    mergesAfter,
    selectionBefore,
    selectionAfter: range,
    selectionBeforeRanges,
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} フィル`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`${rangeToLabel(range)} にフィルを適用しました。`);
}



function homeFillKeyboardPlanForRange(range, fillDirection) {
  const height = range.bottom - range.top + 1;
  const width = range.right - range.left + 1;
  if (fillDirection === "down") {
    const sourceRow = height > 1 ? range.top : range.top - 1;
    if (sourceRow < 1) return showNoHomeFillSourceMessage();
    return {
      sourceRange: normalizeRange({ sheetIndex: state.activeSheetIndex, row: sourceRow, col: range.left }, { sheetIndex: state.activeSheetIndex, row: sourceRow, col: range.right }),
      fillRange: normalizeRange({ sheetIndex: state.activeSheetIndex, row: height > 1 ? range.top + 1 : range.top, col: range.left }, { sheetIndex: state.activeSheetIndex, row: range.bottom, col: range.right }),
    };
  }
  if (fillDirection === "up") {
    const sourceRow = height > 1 ? range.bottom : range.bottom + 1;
    return {
      sourceRange: normalizeRange({ sheetIndex: state.activeSheetIndex, row: sourceRow, col: range.left }, { sheetIndex: state.activeSheetIndex, row: sourceRow, col: range.right }),
      fillRange: normalizeRange({ sheetIndex: state.activeSheetIndex, row: range.top, col: range.left }, { sheetIndex: state.activeSheetIndex, row: height > 1 ? range.bottom - 1 : range.bottom, col: range.right }),
    };
  }
  if (fillDirection === "right") {
    const sourceCol = width > 1 ? range.left : range.left - 1;
    if (sourceCol < 1) return showNoHomeFillSourceMessage();
    return {
      sourceRange: normalizeRange({ sheetIndex: state.activeSheetIndex, row: range.top, col: sourceCol }, { sheetIndex: state.activeSheetIndex, row: range.bottom, col: sourceCol }),
      fillRange: normalizeRange({ sheetIndex: state.activeSheetIndex, row: range.top, col: width > 1 ? range.left + 1 : range.left }, { sheetIndex: state.activeSheetIndex, row: range.bottom, col: range.right }),
    };
  }
  const sourceCol = width > 1 ? range.right : range.right + 1;
  return {
    sourceRange: normalizeRange({ sheetIndex: state.activeSheetIndex, row: range.top, col: sourceCol }, { sheetIndex: state.activeSheetIndex, row: range.bottom, col: sourceCol }),
    fillRange: normalizeRange({ sheetIndex: state.activeSheetIndex, row: range.top, col: range.left }, { sheetIndex: state.activeSheetIndex, row: range.bottom, col: width > 1 ? range.right - 1 : range.right }),
  };
}



function showNoHomeFillSourceMessage() {
  showMessageBox("コピー元のセルがありません。", { title: "Cht WebSheet" });
  return null;
}



function handleSelectedObjectKeyboardShortcut(event) {
  const key = String(event.key || "").toLowerCase();
  const isModifier = event.ctrlKey || event.metaKey;
  const selectedObjects = selectedImageObjects();
  const selectedStrokes = selectedInkStrokes();
  const hasSelectedObjects = selectedObjects.length > 0 || selectedStrokes.length > 0;
  const hasObjectClipboard = Boolean(state.objectClipboard?.objects?.length);
  if (!hasSelectedObjects && !(hasObjectClipboard && key === "escape")) {
    return false;
  }

  if (key === "escape" && hasObjectClipboard) {
    event.preventDefault();
    state.objectClipboard = null;
    cancelObjectPasteFallback();
    renderSheet();
    updateSelectionUi();
    setStatus("オブジェクトのコピー状態を解除しました。");
    return true;
  }

  if (isModifier && !event.altKey) {
    if (key === "c" && selectedObjects.length) {
      event.preventDefault();
      executeWorkbookClipboardShortcut("copy", () => void copySelectedImages({ systemClipboard: true }));
      return true;
    }
    if (key === "x" && selectedObjects.length) {
      event.preventDefault();
      executeWorkbookClipboardShortcut("cut", () => void cutSelectedImages({ systemClipboard: true }));
      return true;
    }
    return false;
  }

  if (!hasSelectedObjects || event.ctrlKey || event.metaKey || event.altKey) return false;
  if (key === "escape") {
    event.preventDefault();
    clearSelectedWorksheetObjects();
    return true;
  }
  if (key === "delete" || key === "backspace") {
    event.preventDefault();
    deleteSelectedWorksheetObjects();
    return true;
  }
  if (key === "f2") {
    const image = selectedObjects.find((item) => item?.shape && !isLineShapeObject(item));
    if (image) {
      event.preventDefault();
      openShapeTextEditor(image.id);
      return true;
    }
  }
  const text = directTypingText(event);
  if (text != null) {
    const image = selectedObjects.find((item) => item?.shape && !isLineShapeObject(item));
    if (image) {
      event.preventDefault();
      openShapeTextEditor(image.id, { initialText: "", replaceText: true });
      return true;
    }
  }
  if (!selectedObjects.length) return false;
  const arrowSteps = {
    arrowleft: [-1, 0],
    arrowright: [1, 0],
    arrowup: [0, -1],
    arrowdown: [0, 1],
  };
  const delta = arrowSteps[key];
  if (delta) {
    event.preventDefault();
    const step = event.shiftKey ? 10 : 1;
    moveSelectedImagesBy(delta[0] * step, delta[1] * step);
    return true;
  }
  return false;
}



function clearSelectionAndStartEdit() {
  const editPoint = firstPointForActiveSelection();
  if (!editPoint) return;
  const targetRange = activeSelectionRange() || selectionRangeFromPoints(editPoint, editPoint);
  if (!ensureEditableRangeForProtection(activeSheet(), targetRange, "削除")) return;
  state.internalClipboard = null;
  clearSelectionContents();
  ensureSheetSize(activeSheet(), editPoint.row, editPoint.col);
  const editRange = selectionRangeFromPoints(editPoint, editPoint);
  if (!rangeIntersects(editRange, visibleGridRange(activeSheet())) || scrollRangeIntoView(editRange)) {
    renderSheet();
  }
  const cellElement = ensureCellElementForPoint(editPoint, { select: true });
  if (cellElement) {
    beginCellEdit(cellElement, { initialText: "", replaceText: true, caretAtEnd: true });
  }
}



function firstPointForActiveSelection() {
  const range = rawActiveSelectionRanges()[0] || activeSelectionRange();
  if (!range) return state.selected?.sheetIndex === state.activeSheetIndex ? state.selected : null;
  return selectableCellPoint(range.top, range.left);
}



function handleSelectionDirectTyping(event) {
  if (!state.selected || state.selected.sheetIndex !== state.activeSheetIndex) return false;
  if (isSelectionKeyboardSink(event.target)) return false;
  if (event.ctrlKey || event.metaKey || event.altKey || isElementImeCompositionEvent(event)) return false;
  if (document.querySelector("#messageBoxOverlay")) return false;

  const text = directTypingText(event);
  if (text == null) return false;

  state.internalClipboard = null;
  const editor = startSelectedCellEdit({ initialText: "", replaceText: true, caretAtEnd: true });
  if (!editor) return false;
  markCellEditDirectTyping(editor, text);
  return true;
}



function directTypingText(event) {
  const key = String(event.key || "");
  if (key.length !== 1) return null;
  if (/[\u0000-\u001f\u007f]/.test(key)) return null;
  return key;
}



function markCellEditDirectTyping(editor, text) {
  if (!editor) return;
  const token = {
    editor,
    text: String(text ?? ""),
    startedAt: nowMs(),
    composing: false,
    compositionData: "",
  };
  state.cellEditDirectTyping = token;
  window.setTimeout(() => {
    if (state.cellEditDirectTyping === token && !token.composing) state.cellEditDirectTyping = null;
  }, 800);
}



function clearCellEditDirectTyping(editor = null) {
  if (!state.cellEditDirectTyping) return;
  if (editor && state.cellEditDirectTyping.editor !== editor) return;
  state.cellEditDirectTyping = null;
}



function markDirectTypingComposition(editor, active) {
  const directTyping = state.cellEditDirectTyping;
  if (!directTyping || directTyping.editor !== editor) return;
  directTyping.composing = Boolean(active);
}



function recordDirectTypingCompositionData(editor, data) {
  const directTyping = state.cellEditDirectTyping;
  if (!directTyping || directTyping.editor !== editor) return;
  directTyping.compositionData = String(data ?? "");
}



function reconcileDirectTypingImeCommit(editor, data) {
  const directTyping = state.cellEditDirectTyping;
  if (!directTyping || directTyping.editor !== editor) return;
  directTyping.composing = false;
  const initialText = directTyping.text;
  const compositionText = String(directTyping.compositionData || "");
  const committedText = String(data || directTyping.compositionData || "");
  if (!initialText || !committedText || !editor?.isConnected || editor.getAttribute("contenteditable") !== "true") {
    clearCellEditDirectTyping(editor);
    return;
  }
  window.setTimeout(() => {
    if (!editor.isConnected || editor.getAttribute("contenteditable") !== "true") {
      clearCellEditDirectTyping(editor);
      return;
    }
    const text = editablePlainText(editor);
    const residuePrefixes = Array.from(new Set([compositionText, initialText]))
      .filter((prefix) => prefix && prefix !== committedText)
      .sort((a, b) => b.length - a.length);
    const residuePrefix = residuePrefixes.find((prefix) => {
      const residue = prefix + committedText;
      return text === residue || text.startsWith(residue);
    });
    if (residuePrefix) {
      editor.textContent = text.slice(residuePrefix.length);
      placeCaretAtEnd(editor);
      syncEditableCellAlignment(editor);
      saveCellEditSelection(editor);
      updateFormulaReferenceUi();
    }
    clearCellEditDirectTyping(editor);
  }, 0);
}



function selectionKeyboardSinkState() {
  if (!state.selectionKeyboardSink) {
    state.selectionKeyboardSink = {
      composing: false,
      focusFrame: 0,
      suppressInputUntil: 0,
      edit: null,
      shapeEdit: null,
    };
  }
  return state.selectionKeyboardSink;
}



function selectedPointForKeyboardSink() {
  if (!state.model || !state.selected || state.selected.sheetIndex !== state.activeSheetIndex) return null;
  if (state.selectedImage || selectedImageObjects().length) return null;
  return state.selected;
}



function selectedShapeForKeyboardSink() {
  if (!state.model || state.formulaEdit || state.shapeTextEditor) return null;
  if (state.selected && state.selected.sheetIndex === state.activeSheetIndex) return null;
  const objects = selectedImageObjects();
  if (objects.length !== 1) return null;
  const image = objects[0];
  if (!image?.shape || isLineShapeObject(image)) return null;
  return image;
}



function isSelectionKeyboardSinkEligibleActiveElement(element = document.activeElement, options = {}) {
  if (!element || element === document.body || element === document.documentElement) return true;
  if (isSelectionKeyboardSink(element)) return true;
  if (element === $sheetHost[0]) return true;
  if (element.closest?.(".sheet-cell[data-row][data-col], .sheet-grid")) return true;
  if (options.allowShape && !isTextEditingTarget(element)) return true;
  return false;
}



function shouldFocusSelectionKeyboardSink() {
  if (!$selectionKeyboardSink[0]) return false;
  const point = selectedPointForKeyboardSink();
  const shape = selectedShapeForKeyboardSink();
  if (!point && !shape) return false;
  if (activeSheetEditElement()) return false;
  if (state.shapeTextEditor) return false;
  if (state.formulaEdit || state.isSelecting || state.selectionDrag || state.fillDrag) return false;
  if (document.querySelector("#messageBoxOverlay, #dataDialog, #findDialog, #goToDialog, #specialDialog, #formatCellsDialog, #linkDialog")) return false;
  return isSelectionKeyboardSinkEligibleActiveElement(document.activeElement, { allowShape: Boolean(shape) });
}



function scheduleSelectionKeyboardSinkFocus() {
  const sinkState = selectionKeyboardSinkState();
  if (sinkState.focusFrame) return;
  sinkState.focusFrame = requestAnimationFrame(() => {
    sinkState.focusFrame = 0;
    if (!shouldFocusSelectionKeyboardSink()) {
      hideSelectionKeyboardSink();
      return;
    }
    syncSelectionKeyboardSinkEditWithSelection();
    positionSelectionKeyboardSink();
    const sink = $selectionKeyboardSink[0];
    if (!sink) return;
    sink.classList.add("is-active");
    if (document.activeElement !== sink) {
      sink.focus({ preventScroll: true });
    }
    try {
      const length = String(sink.value || "").length;
      sink.setSelectionRange(length, length);
    } catch {
      // Selection can fail during some IME transitions. Focus is enough.
    }
  });
}



function hideSelectionKeyboardSink() {
  const sink = $selectionKeyboardSink[0];
  if (!sink) return;
  sink.classList.remove("is-active");
  sink.style.left = "-10000px";
  sink.style.top = "-10000px";
  sink.style.width = "1px";
  sink.style.height = "1px";
}



function positionSelectionKeyboardSink() {
  const sink = $selectionKeyboardSink[0];
  const point = selectedPointForKeyboardSink();
  const shape = selectedShapeForKeyboardSink();
  if (!sink || activeSheetEditElement() || state.shapeTextEditor) {
    hideSelectionKeyboardSink();
    return;
  }
  if (!point && shape) {
    positionSelectionKeyboardSinkForShape(sink, shape);
    return;
  }
  if (!point) {
    hideSelectionKeyboardSink();
    return;
  }
  const cell = cellElementForPoint(point);
  if (!cell?.isConnected) {
    hideSelectionKeyboardSink();
    return;
  }
  const rect = cell.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) {
    hideSelectionKeyboardSink();
    return;
  }
  const computed = getComputedStyle(cell);
  sink.style.left = `${Math.round(rect.left)}px`;
  sink.style.top = `${Math.round(rect.top)}px`;
  sink.style.width = `${Math.max(1, Math.round(rect.width))}px`;
  sink.style.height = `${Math.max(1, Math.round(rect.height))}px`;
  sink.style.font = computed.font;
  sink.style.lineHeight = computed.lineHeight;
  sink.style.textAlign = computed.textAlign;
  sink.style.padding = computed.padding;
}



function positionSelectionKeyboardSinkForShape(sink, image) {
  const element =
    primaryImagePositionElement(image.id) ||
    $sheetHost.find(`.sheet-image-selection-frame[data-image-id="${cssEscape(image.id)}"]`)[0];
  if (!element?.isConnected) {
    hideSelectionKeyboardSink();
    return;
  }
  const rect = element.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) {
    hideSelectionKeyboardSink();
    return;
  }
  const shape = image.shape || {};
  const margin = normalizeShapeTextMargin(shape.textMargin, shape.type === "textBox");
  const textScale = Number(activeSheet()?.displayScale || 1);
  sink.style.left = `${Math.round(rect.left)}px`;
  sink.style.top = `${Math.round(rect.top)}px`;
  sink.style.width = `${Math.max(1, Math.round(rect.width))}px`;
  sink.style.height = `${Math.max(1, Math.round(rect.height))}px`;
  sink.style.fontFamily = shape.textFontFamily || fontFamilyForExcel("Calibri") || '"Yu Gothic UI", Meiryo, sans-serif';
  sink.style.fontSize = `${Math.max(8, Math.round((Number(shape.textFontSize) || 11) * 96 / 72)) * textScale}px`;
  sink.style.fontWeight = shape.textBold ? "700" : "400";
  sink.style.fontStyle = shape.textItalic ? "italic" : "normal";
  sink.style.lineHeight = "1.22";
  sink.style.textAlign = shape.textAlign || "center";
  sink.style.padding = `${margin.top * textScale}px ${margin.right * textScale}px ${margin.bottom * textScale}px ${margin.left * textScale}px`;
}



function syncSelectionKeyboardSinkEditWithSelection() {
  const sinkState = selectionKeyboardSinkState();
  const edit = sinkState.edit;
  const point = selectedPointForKeyboardSink();
  const shapeEdit = sinkState.shapeEdit;
  const shape = selectedShapeForKeyboardSink();
  if (sinkState.composing) return;
  if (edit && (!point || edit.sheetIndex !== point.sheetIndex || edit.row !== point.row || edit.col !== point.col)) {
    sinkState.edit = null;
    resetSelectionKeyboardSinkValue();
  }
  if (shapeEdit && (!shape || shapeEdit.sheetIndex !== state.activeSheetIndex || shapeEdit.imageId !== shape.id)) {
    sinkState.shapeEdit = null;
    resetSelectionKeyboardSinkValue();
  }
}



function resetSelectionKeyboardSinkValue() {
  const sink = $selectionKeyboardSink[0];
  if (!sink) return;
  sink.value = "";
}



function ensureSelectionKeyboardImeEdit() {
  const sinkState = selectionKeyboardSinkState();
  const point = selectedPointForKeyboardSink();
  if (!point) return null;
  const edit = sinkState.edit;
  if (edit && edit.sheetIndex === point.sheetIndex && edit.row === point.row && edit.col === point.col) {
    return edit;
  }
  const before = getCellRawInput(activeSheet(), point.row, point.col);
  sinkState.edit = {
    sheetIndex: point.sheetIndex,
    row: point.row,
    col: point.col,
    before,
    value: "",
  };
  return sinkState.edit;
}



function ensureSelectionKeyboardShapeImeEdit() {
  const sinkState = selectionKeyboardSinkState();
  const image = selectedShapeForKeyboardSink();
  if (!image) return null;
  const edit = sinkState.shapeEdit;
  if (edit && edit.sheetIndex === state.activeSheetIndex && edit.imageId === image.id) {
    return edit;
  }
  sinkState.shapeEdit = {
    sheetIndex: state.activeSheetIndex,
    imageId: image.id,
    before: shapeTextPlainValue(image.shape),
    value: "",
  };
  sinkState.edit = null;
  return sinkState.shapeEdit;
}



function promoteSelectionKeyboardImeEditToCellEditor(edit) {
  if (!edit || edit.sheetIndex !== state.activeSheetIndex) return;
  const sinkState = selectionKeyboardSinkState();
  sinkState.edit = null;
  resetSelectionKeyboardSinkValue();
  hideSelectionKeyboardSink();
  state.selected = { sheetIndex: edit.sheetIndex, row: edit.row, col: edit.col };
  state.selectionRange = selectionRangeFromPoints(state.selected, state.selected);
  state.selectionRanges = [state.selectionRange];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectionDrag = null;
  const cellElement = ensureCellElementForPoint(state.selected, { select: true });
  if (!cellElement) return;
  beginCellEdit(cellElement, { initialText: edit.value, replaceText: true, caretAtEnd: true });
  $(cellElement).data("beforeEdit", edit.before);
}



function promoteSelectionKeyboardShapeImeEditToEditor(edit) {
  if (!edit || edit.sheetIndex !== state.activeSheetIndex) return;
  const sinkState = selectionKeyboardSinkState();
  sinkState.shapeEdit = null;
  resetSelectionKeyboardSinkValue();
  hideSelectionKeyboardSink();
  const image = selectedImageObjectById(edit.imageId);
  if (!image?.shape || isLineShapeObject(image)) return;
  if (!isActiveImage(image.id)) {
    setSelectedImageIds([image.id], image.id);
    renderSheet();
    updateSelectionUi();
  }
  openShapeTextEditor(image.id, { initialText: edit.value, replaceText: true });
  if (state.shapeTextEditor) state.shapeTextEditor.beforeText = edit.before;
}



function appendSelectionKeyboardImeText(text) {
  const committedText = String(text ?? "");
  if (!committedText) return;
  const edit = ensureSelectionKeyboardImeEdit();
  if (!edit) return;
  edit.value += committedText;
  promoteSelectionKeyboardImeEditToCellEditor(edit);
}



function appendSelectionKeyboardShapeImeText(text) {
  const committedText = String(text ?? "");
  if (!committedText) return;
  const edit = ensureSelectionKeyboardShapeImeEdit();
  if (!edit) return;
  edit.value += committedText;
  promoteSelectionKeyboardShapeImeEditToEditor(edit);
}



function cancelSelectionKeyboardImeEdit() {
  const sinkState = selectionKeyboardSinkState();
  sinkState.edit = null;
  sinkState.shapeEdit = null;
  resetSelectionKeyboardSinkValue();
  scheduleSelectionKeyboardSinkFocus();
}



function handleSelectionKeyboardSinkKeydown(event) {
  const sinkState = selectionKeyboardSinkState();
  if (isElementImeCompositionEvent(event, event.currentTarget) || sinkState.composing) {
    event.stopPropagation();
    return;
  }
  const dateTimeShortcut = keyboardDateTimeShortcutKind(event, String(event.key || "").toLowerCase());
  if (dateTimeShortcut) {
    event.preventDefault();
    event.stopPropagation();
    sinkState.edit = null;
    sinkState.shapeEdit = null;
    resetSelectionKeyboardSinkValue();
    insertKeyboardDateTimeByKeyboard(dateTimeShortcut);
    return;
  }
  if (
    (sinkState.shapeEdit || selectedShapeForKeyboardSink()) &&
    !event.ctrlKey &&
    !event.metaKey &&
    !event.altKey &&
    directTypingText(event) != null
  ) {
    event.stopPropagation();
    return;
  }
  if (event.key === "Escape" && sinkState.edit) {
    event.preventDefault();
    event.stopPropagation();
    cancelSelectionKeyboardImeEdit();
    return;
  }
  if (event.key === "Escape" && sinkState.shapeEdit) {
    event.preventDefault();
    event.stopPropagation();
    cancelSelectionKeyboardImeEdit();
    return;
  }
  if (isEnterKeyEvent(event) && sinkState.edit) {
    event.preventDefault();
    event.stopPropagation();
    const edit = sinkState.edit;
    sinkState.edit = null;
    resetSelectionKeyboardSinkValue();
    selectCellByKeyboard(edit.row + (event.shiftKey ? -1 : 1), edit.col, { rowStep: event.shiftKey ? -1 : 1, colStep: 0 });
    return;
  }
  if (event.key === "Backspace" && sinkState.edit && !event.currentTarget.value) {
    event.preventDefault();
    event.stopPropagation();
    const edit = sinkState.edit;
    edit.value = Array.from(String(edit.value || "")).slice(0, -1).join("");
    if (edit.value) promoteSelectionKeyboardImeEditToCellEditor(edit);
  }
  if (event.key === "Backspace" && sinkState.shapeEdit && !event.currentTarget.value) {
    event.preventDefault();
    event.stopPropagation();
    const edit = sinkState.shapeEdit;
    edit.value = Array.from(String(edit.value || "")).slice(0, -1).join("");
    if (edit.value) promoteSelectionKeyboardShapeImeEditToEditor(edit);
  }
}



function handleSelectionKeyboardSinkBeforeInput(event) {
  const nativeEvent = event.originalEvent || event;
  if (String(nativeEvent.inputType || "") === "insertLineBreak") {
    event.preventDefault();
  }
}



function handleSelectionKeyboardSinkInput(event) {
  const sinkState = selectionKeyboardSinkState();
  if (sinkState.composing || nowMs() < Number(sinkState.suppressInputUntil || 0)) return;
  const sink = event.currentTarget;
  const text = String(sink.value || "");
  if (!text) return;
  resetSelectionKeyboardSinkValue();
  if (sinkState.edit) {
    appendSelectionKeyboardImeText(text);
    return;
  }
  if (sinkState.shapeEdit || selectedShapeForKeyboardSink()) {
    state.internalClipboard = null;
    appendSelectionKeyboardShapeImeText(text);
    return;
  }
  state.internalClipboard = null;
  const editor = startSelectedCellEdit({ initialText: text, replaceText: true, caretAtEnd: true });
  if (!editor) return;
  clearCellEditDirectTyping(editor);
}



function handleSelectionKeyboardSinkCompositionStart(event) {
  const sinkState = selectionKeyboardSinkState();
  sinkState.composing = true;
  if (selectedShapeForKeyboardSink()) ensureSelectionKeyboardShapeImeEdit();
  else ensureSelectionKeyboardImeEdit();
  positionSelectionKeyboardSink();
  markElementImeComposition(event.currentTarget, true);
}



function handleSelectionKeyboardSinkCompositionUpdate(event) {
  positionSelectionKeyboardSink();
}



function handleSelectionKeyboardSinkCompositionEnd(event) {
  const sinkState = selectionKeyboardSinkState();
  sinkState.composing = false;
  sinkState.suppressInputUntil = nowMs() + 120;
  markElementImeComposition(event.currentTarget, false);
  const committedFromEvent = String(event.originalEvent?.data || event.data || "");
  window.setTimeout(() => {
    const sink = $selectionKeyboardSink[0];
    const committedText = committedFromEvent || String(sink?.value || "");
    resetSelectionKeyboardSinkValue();
    if (sinkState.shapeEdit || selectedShapeForKeyboardSink()) appendSelectionKeyboardShapeImeText(committedText);
    else appendSelectionKeyboardImeText(committedText);
    scheduleSelectionKeyboardSinkFocus();
  }, 0);
}



function handleSelectionKeyboardSinkBlur() {
  const sinkState = selectionKeyboardSinkState();
  if (sinkState.composing) return;
  resetSelectionKeyboardSinkValue();
}



function isSpaceKeyEvent(event) {
  const key = String(event.key || "");
  return key === " " || key === "Space" || key === "Spacebar";
}



function isSheetSelectionShortcutTarget(target) {
  const element = target?.nodeType === Node.TEXT_NODE ? target.parentElement : target;
  if (!element) return true;
  const floatingSelectors = [
    "#shapeSelectionPane",
    "#shapeFormatPane",
    "#shapeFormatMenu",
    "#ribbonShapeMenu",
    "#ribbonUnderlineMenu",
    "#ribbonBorderMenu",
    "#ribbonInkEraserMenu",
    "#ribbonColorPalette",
    "#homeCommandMenu",
    "#dataCommandMenu",
    "#dataValidationMenu",
    "#whatIfMenu",
    "#viewCommandMenu",
    "#quickAnalysisPopup",
    "#queryPaneContextMenu",
    "#statusBarContextMenu",
    "#autoFilterMenu",
    "#cellContextMenu",
    "#sheetTabContextMenu",
    "#sheetTabNavigatorMenu",
    "#fileSideMenu",
    "#nameBoxMenu",
    ".excel-mini-dialog",
    ".message-box",
    ".cell-annotation-editor",
    ".cell-comment-thread-popup",
    ".cell-note-popup",
  ].join(",");
  if (document.activeElement?.closest?.(floatingSelectors)) return false;
  return !element.closest?.(
    floatingSelectors,
  );
}



function activateAdjacentVisibleSheet(delta) {
  if (!state.model?.sheets?.length) return;
  const visibleIndexes = visibleSheetIndexes();
  if (!visibleIndexes.length) return;
  const currentPosition = Math.max(0, visibleIndexes.indexOf(state.activeSheetIndex));
  const nextPosition = positiveModulo(currentPosition + delta, visibleIndexes.length);
  const nextIndex = visibleIndexes[nextPosition];
  if (nextIndex === state.activeSheetIndex) return;
  const nextName = state.model.sheets[nextIndex]?.name || "";
  setActiveSheet(nextIndex);
  setStatus(`${nextName} に移動しました。`);
}



function activateAdjacentVisibleSheetGroup(delta) {
  if (!state.model?.sheets?.length) return;
  const visibleIndexes = visibleSheetIndexes();
  if (visibleIndexes.length < 2) {
    setStatus("選択できる隣接シートがありません。");
    return;
  }
  const currentPosition = Math.max(0, visibleIndexes.indexOf(state.activeSheetIndex));
  const nextPosition = positiveModulo(currentPosition + delta, visibleIndexes.length);
  const nextIndex = visibleIndexes[nextPosition];
  const group = new Set((state.selectedSheetIndexes || []).filter((index) => visibleIndexes.includes(index)));
  if (visibleIndexes.includes(state.activeSheetIndex)) group.add(state.activeSheetIndex);
  group.add(nextIndex);
  state.activeSheetIndex = nextIndex;
  state.selectedSheetIndexes = [...group].sort((left, right) => left - right);
  renderWorkbook();
  setStatus(`${state.selectedSheetIndexes.length} 個の表示ワークシートを選択しました。`);
}



function activateAdjacentWorkbookWindow(delta) {
  const records = workbookWindowRecords().filter((view) => !view.hidden);
  if (records.length < 2) {
    setStatus("切り替えるウィンドウがありません。");
    return false;
  }
  const activeId = state.model?.activeWindowId || "main";
  const currentPosition = Math.max(0, records.findIndex((view) => view.id === activeId));
  const next = records[positiveModulo(currentPosition + delta, records.length)];
  return switchWorkbookWindow(next?.id || "main");
}



function selectAllByKeyboard() {
  const sheet = activeSheet();
  if (!sheet) return;
  withDisplayCalculationCache(sheet, () => selectAllByKeyboardWithCache(sheet));
}



function selectAllByKeyboardWithCache(sheet) {
  const activeRanges = activeSelectionRanges();
  const activeRange = activeRanges.length === 1 ? activeRanges[0] : activeSelectionRange();
  const activePoint = selectAllActivePoint(activeRange);
  const candidates = selectAllExpansionCandidates(sheet, activePoint);
  const target = selectAllNextCandidate(candidates, activeRanges.length > 1 ? selectAllWorksheetRange(sheet) : activeRange);
  if (!target) return;
  selectRangesByKeyboard([target.range], target.message, {
    scrollFocusOnly: true,
    activePoint,
  });
}



function selectAllActivePoint(activeRange = null) {
  if (state.selected?.sheetIndex === state.activeSheetIndex) return state.selected;
  const fallbackRange = activeRange || activeSelectionRange();
  return {
    sheetIndex: state.activeSheetIndex,
    row: fallbackRange?.top || 1,
    col: fallbackRange?.left || 1,
  };
}



function selectAllExpansionCandidates(sheet, activePoint) {
  const candidates = [
    ...selectAllTableCandidates(sheet, activePoint),
    selectAllCurrentRegionCandidate(sheet, activePoint),
    selectAllUsedRangeCandidate(sheet),
    {
      kind: "worksheet",
      range: selectAllWorksheetRange(sheet),
      message: "ワークシート全体を選択しました。",
    },
  ].filter(Boolean);
  return uniqueSelectAllCandidates(candidates);
}



function selectAllTableCandidates(sheet, activePoint) {
  if (!sheet?.tables?.length || !activePoint) return [];
  const table = (sheet.tables || [])
    .filter((item) => item?.range && isCellInRange(item.range, activePoint.row, activePoint.col))
    .sort((left, right) => rangeCellCount(left.range) - rangeCellCount(right.range))[0];
  if (!table?.range) return [];
  const ranges = [];
  const dataRange = tableDataBodySelectionRange(table);
  if (dataRange && isCellInRange(dataRange, activePoint.row, activePoint.col) && !rangesEqualLoose(dataRange, table.range)) {
    ranges.push({
      kind: "table-data",
      range: dataRange,
      message: `${rangeToLabel(dataRange)} のテーブル データを選択しました。`,
    });
  }
  ranges.push({
    kind: "table",
    range: { sheetIndex: state.activeSheetIndex, ...table.range },
    message: `${rangeToLabel(table.range)} のテーブル全体を選択しました。`,
  });
  return ranges;
}



function tableDataBodySelectionRange(table) {
  const range = table?.range;
  if (!range) return null;
  const top = range.top + (table.showHeaderRow === false ? 0 : 1);
  const bottom = table.showTotalsRow === true ? range.bottom - 1 : range.bottom;
  if (top > bottom) return null;
  return {
    sheetIndex: range.sheetIndex ?? state.activeSheetIndex,
    top,
    left: range.left,
    bottom,
    right: range.right,
  };
}



function selectAllCurrentRegionCandidate(sheet, activePoint) {
  const region = currentRegionFromPoint(sheet, activePoint);
  if (!region) return null;
  const blankSingleton = isSingleCellRange(region) && !cellHasDataForRegion(sheet, activePoint.row, activePoint.col);
  if (blankSingleton) return null;
  return {
    kind: "current-region",
    range: region,
    message: `${rangeToLabel(region)} を選択しました。`,
  };
}



function selectAllUsedRangeCandidate(sheet) {
  const range = sheetUsedRangeForSelectAll(sheet);
  if (!range) return null;
  return {
    kind: "used-range",
    range,
    message: `${rangeToLabel(range)} の使用範囲を選択しました。`,
  };
}



function sheetUsedRangeForSelectAll(sheet) {
  const ranges = [];
  const used = sheetUsedRange(sheet);
  if (used) ranges.push({ sheetIndex: state.activeSheetIndex, ...used });
  (sheet?.tables || []).forEach((table) => {
    if (table?.range) ranges.push({ sheetIndex: state.activeSheetIndex, ...table.range });
  });
  if (sheet?.autoFilter?.range) ranges.push({ sheetIndex: state.activeSheetIndex, ...sheet.autoFilter.range });
  if (!ranges.length) return null;
  const union = unionMaterializedRanges(ranges, state.activeSheetIndex);
  return expandRangeForMerges(sheet, union) || union;
}



function selectAllWorksheetRange(sheet) {
  return {
    sheetIndex: state.activeSheetIndex,
    top: 1,
    left: 1,
    bottom: sheet.rowCount,
    right: sheet.colCount,
  };
}



function selectAllNextCandidate(candidates, activeRange) {
  const worksheet = candidates.find((candidate) => candidate.kind === "worksheet") || candidates[candidates.length - 1] || null;
  if (!activeRange) return candidates[0] || worksheet;
  const next = candidates.find((candidate) =>
    !rangesEqualLoose(candidate.range, activeRange) &&
    rangeContains(candidate.range, activeRange)
  );
  if (next) return next;
  return worksheet && !rangesEqualLoose(worksheet.range, activeRange) ? worksheet : worksheet;
}



function uniqueSelectAllCandidates(candidates) {
  const seen = new Set();
  return (candidates || []).filter((candidate) => {
    if (!candidate?.range) return false;
    const range = candidate.range;
    const key = `${range.top}:${range.left}:${range.bottom}:${range.right}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}



function selectRowsByKeyboard() {
  const ranges = activeSelectionRanges();
  if (!ranges.length) return;
  const rowRanges = uniqueSelectionRanges(ranges.map((range) => rowSelectionRange(range.top, range.bottom)));
  selectRangesByKeyboard(rowRanges, `${rangesToLabel(rowRanges)} を選択しました。`, { scrollFocusOnly: true });
}



function selectColumnsByKeyboard() {
  const ranges = activeSelectionRanges();
  if (!ranges.length) return;
  const columnRanges = uniqueSelectionRanges(ranges.map((range) => columnSelectionRange(range.left, range.right)));
  selectRangesByKeyboard(columnRanges, `${rangesToLabel(columnRanges)} を選択しました。`, { scrollFocusOnly: true });
}



function selectRangesByKeyboard(ranges, statusMessage, options = {}) {
  const firstRange = ranges[0];
  if (!firstRange) return;
  const activePoint = options.activePoint;
  const point = activePoint && isCellInRange(firstRange, activePoint.row, activePoint.col)
    ? selectableCellPoint(activePoint.row, activePoint.col)
    : selectableCellPoint(firstRange.top, firstRange.left);
  state.selected = point;
  state.selectionRange = { ...firstRange };
  state.selectionRanges = ranges.map((range) => ({ ...range }));
  state.selectionStart = point;
  state.selectionAnchor = point;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  state.formulaEdit = null;
  const scrollRange = options.scrollFocusOnly ? selectionRangeFromPoints(point, point) : firstRange;
  const sheet = activeSheet();
  const focusAlreadyVisible = options.scrollFocusOnly && point && sheet
    ? isCellInRange(visibleGridRange(sheet), point.row, point.col)
    : false;
  const didScroll = focusAlreadyVisible ? false : scrollRangeIntoView(scrollRange);
  if (didScroll) {
    renderSheet(null, { deferSelectionUiSync: true });
  }
  updateSelectionUi({ skipStatusBar: true });
  updateFormulaBarForSelection();
  scheduleSelectionUiSync();
  setStatus(statusMessage);
}



function handleSelectionKeyboardNavigation(event) {
  const key = String(event.key || "");
  if (key === "F2") {
    if (event.ctrlKey || event.metaKey || event.altKey || event.shiftKey || !state.selected) return false;
    event.preventDefault();
    startSelectedCellEdit();
    return true;
  }

  const extensionTarget = keyboardSelectionExtensionTargetFromEvent(event);
  if (extensionTarget) {
    event.preventDefault();
    extendSelectionByKeyboard(extensionTarget.row, extensionTarget.col, extensionTarget);
    return true;
  }

  const isEnter = isEnterKeyEvent(event);
  const movement = isEnter ? (event.shiftKey ? "up" : "down") : (!event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey ? {
    ArrowDown: "down",
    ArrowUp: "up",
    ArrowLeft: "left",
    ArrowRight: "right",
  }[key] : null);
  if (movement) {
    if (event.ctrlKey || event.metaKey || event.altKey || (event.shiftKey && !isEnter) || !state.selected) return false;

    event.preventDefault();
    moveSelectionByKeyboard(movement);
    return true;
  }

  const target = keyboardNavigationTargetFromEvent(event);
  if (!target) return false;

  event.preventDefault();
  selectCellByKeyboard(target.row, target.col, target);
  return true;
}



function isTextEditingTarget(target) {
  if (!target) return false;
  const element = target.nodeType === Node.TEXT_NODE ? target.parentElement : target;
  if (!element) return false;
  if (isSelectionKeyboardSink(element)) return false;
  if (element.closest?.(".sheet-cell[contenteditable='true']")) return true;
  return Boolean(element.closest?.("input, textarea, select, [contenteditable='true']"));
}



function isSelectionKeyboardSink(element) {
  const target = element?.nodeType === Node.TEXT_NODE ? element.parentElement : element;
  return Boolean(target?.closest?.("#selectionKeyboardSink"));
}



function handleTextInputImeCompositionStart(event) {
  markElementImeComposition(event.currentTarget, true);
}



function handleTextInputImeCompositionEnd(event) {
  markElementImeComposition(event.currentTarget, false);
}



function markElementImeComposition(target, active) {
  if (!target) return;
  const $target = $(target);
  $target.data("imeComposing", Boolean(active));
  if (!active) $target.data("imeCompositionEndedAt", nowMs());
}



function isElementImeCompositionEvent(event, target = event?.target) {
  const nativeEvent = event?.originalEvent || event || {};
  const key = String(event?.key || nativeEvent.key || "");
  const code = String(event?.code || nativeEvent.code || "");
  const keyCode = Number(nativeEvent.keyCode || nativeEvent.which || event?.keyCode || event?.which || 0);
  const $target = target ? $(target) : null;
  const recentlyEndedAt = Number($target?.data("imeCompositionEndedAt") || 0);
  const isEnter = key === "Enter" || key === "NumpadEnter" || code === "Enter" || code === "NumpadEnter" || keyCode === 13;
  const isEscape = key === "Escape" || code === "Escape" || keyCode === 27;
  const justEndedForActionKey = (isEnter || isEscape) && recentlyEndedAt > 0 && nowMs() - recentlyEndedAt < 80;
  return Boolean(nativeEvent.isComposing || event?.isComposing || keyCode === 229 || $target?.data("imeComposing") || justEndedForActionKey);
}



function markImeComposition(target, scope, active) {
  if (!target) return;
  markElementImeComposition(target, active);
  if (scope === "formula") {
    state.formulaBarComposing = Boolean(active);
  } else {
    state.cellEditComposing = Boolean(active);
  }
}



function isImeCompositionEvent(event, target, scope = "cell") {
  const scopeComposing = scope === "formula" ? state.formulaBarComposing : state.cellEditComposing;
  return Boolean(isElementImeCompositionEvent(event, target) || scopeComposing);
}



function nowMs() {
  return window.performance?.now?.() || Date.now();
}



function captureRangeForClipboard(range, text, options = {}) {
  const sheet = activeSheet();
  const cellEntries = [];
  const merges = sheet.merges
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
    text,
    mode: options.mode === "cut" ? "cut" : "copy",
    sourceSheetIndex: state.activeSheetIndex,
    sourceRange: { ...range },
    rowCount: range.bottom - range.top + 1,
    colCount: range.right - range.left + 1,
    rows: range.fullRow ? sheet.rows.slice(range.top - 1, range.bottom).map((row) => ({ ...defaultRowModel(sheet), ...(row || {}) })) : [],
    columns: range.fullColumn ? sheet.columns.slice(range.left - 1, range.right).map((column) => ({ ...defaultColumnModel(sheet), ...(column || {}) })) : [],
    merges,
    cellEntries,
  };
}



function cloneCellForClipboard(cell) {
  const nextCell = cloneCellModel(cell);
  if (!nextCell) return null;
  if (nextCell.raw == null && nextCell.formula) {
    nextCell.raw = nextCell.formula;
  } else if (nextCell.raw == null && nextCell.cached != null) {
    nextCell.raw = nextCell.cached;
  } else if (nextCell.raw == null && nextCell.display !== "" && nextCell.display != null) {
    nextCell.raw = nextCell.display;
  }
  if (nextCell.display == null && nextCell.raw != null) {
    nextCell.display = String(nextCell.raw);
  }
  return nextCell;
}



function canPasteInternalClipboard(text) {
  return shouldUseInternalClipboardPaste(text);
}



function isInternalCutClipboard(clipboard = state.internalClipboard) {
  return Boolean(
    clipboard?.mode === "cut" &&
      clipboard.sourceRange &&
      Number.isInteger(Number(clipboard.sourceSheetIndex)) &&
      state.model?.sheets?.[Number(clipboard.sourceSheetIndex)],
  );
}



function cutPasteHistorySheetIndexes(clipboard, targetSheetIndex = state.activeSheetIndex) {
  if (!isInternalCutClipboard(clipboard)) return [];
  return [...new Set([Number(clipboard.sourceSheetIndex), Number(targetSheetIndex)].filter((index) => state.model?.sheets?.[index]))];
}



function clearCutSourceAfterInternalPaste(clipboard, targetRange, targetSheetIndex = state.activeSheetIndex) {
  if (!isInternalCutClipboard(clipboard)) return false;
  const sourceSheetIndex = Number(clipboard.sourceSheetIndex);
  const sourceSheet = state.model.sheets[sourceSheetIndex];
  const sourceRange = clipboard.sourceRange;
  const sameSheet = sourceSheetIndex === Number(targetSheetIndex);

  sourceSheet.merges = (sourceSheet.merges || []).filter((merge) => {
    if (!rangesIntersect(merge, sourceRange)) return true;
    return sameSheet && targetRange && rangesIntersect(merge, targetRange);
  });
  forEachCellInRange(sourceRange, (row, col) => {
    if (sameSheet && targetRange && isCellInRange(targetRange, row, col)) return;
    delete sourceSheet.cells[cellKey(row, col)];
  });
  state.internalClipboard = null;
  clearPersistedInternalClipboard();
  return true;
}



function rangeToTsv(range) {
  const sheet = activeSheet();
  const sheetId = getSheetId(sheet.name);
  const lines = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    const values = [];
    for (let col = range.left; col <= range.right; col += 1) {
      values.push(String(getDisplayForCell(sheet, sheetId, row, col) ?? ""));
    }
    lines.push(values.join("\t"));
  }
  return lines.join("\n");
}



function pasteTargetSizeForClipboard(clipboard, range, options = {}) {
  const sourceRowCount = Math.max(1, options.transpose ? clipboard.colCount : clipboard.rowCount);
  const sourceColCount = Math.max(1, options.transpose ? clipboard.rowCount : clipboard.colCount);
  const selectedRowCount = Math.max(1, (range?.bottom || range?.top || 1) - (range?.top || 1) + 1);
  const selectedColCount = Math.max(1, (range?.right || range?.left || 1) - (range?.left || 1) + 1);
  let targetRowCount = sourceRowCount;
  let targetColCount = sourceColCount;
  const canRepeat =
    options.repeat !== false &&
    !range?.fullRow &&
    !range?.fullColumn &&
    !clipboard?.merges?.length &&
    selectedRowCount >= sourceRowCount &&
    selectedColCount >= sourceColCount &&
    (selectedRowCount > sourceRowCount || selectedColCount > sourceColCount);
  if (canRepeat) {
    if (selectedRowCount % sourceRowCount !== 0 || selectedColCount % sourceColCount !== 0) {
      return {
        ok: false,
        message: "コピー領域と貼り付け領域のサイズが違うため、貼り付けできません。",
      };
    }
    targetRowCount = selectedRowCount;
    targetColCount = selectedColCount;
  }
  return {
    ok: true,
    sourceRowCount,
    sourceColCount,
    targetRowCount,
    targetColCount,
    repeating: targetRowCount !== sourceRowCount || targetColCount !== sourceColCount,
  };
}



function repeatedClipboardSourceOffset(rowOffset, colOffset, pasteSize, transpose = false) {
  const localRow = ((Math.max(1, rowOffset) - 1) % Math.max(1, pasteSize?.sourceRowCount || 1)) + 1;
  const localCol = ((Math.max(1, colOffset) - 1) % Math.max(1, pasteSize?.sourceColCount || 1)) + 1;
  return sourceOffsetForPasteSpecial(localRow, localCol, transpose);
}



function pasteInternalClipboardSpecialAtRange(range, options = {}) {
  const clipboard = state.internalClipboard;
  if (!clipboard) {
    setStatus("貼り付けできる内部セル データがありません。");
    return;
  }
  if (!range) {
    setStatus("貼り付け先のセルを選択してください。");
    return;
  }
  const special = normalizePasteSpecialOptions(options);
  if (special.mode === "all" && special.operation === "none" && !special.skipBlanks && !special.transpose) {
    pasteInternalClipboardAtRange(range);
    return;
  }

  const sheet = activeSheet();
  const pasteSize = pasteTargetSizeForClipboard(clipboard, range, { transpose: special.transpose });
  if (!pasteSize.ok) {
    showMessageBox(pasteSize.message, { title: "Cht WebSheet" });
    setStatus(pasteSize.message);
    return;
  }
  const targetRowCount = pasteSize.targetRowCount;
  const targetColCount = pasteSize.targetColCount;
  const targetRows = range.top + targetRowCount - 1;
  const targetCols = range.left + targetColCount - 1;
  const targetRange = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left },
    { sheetIndex: state.activeSheetIndex, row: targetRows, col: targetCols },
  );
  if (!ensureEditableRangeForProtection(sheet, targetRange, "貼り付け")) return;
  const entriesByOffset = new Map();
  clipboardCellEntries(clipboard).forEach((sourceCell) => {
    entriesByOffset.set(cellKey(sourceCell.row, sourceCell.col), sourceCell);
  });
  const tableEditCandidate = tableRangeEditMayChangeMetadata(sheet, targetRange, {
    assumeNonEmpty: pasteSpecialModeWritesContent(special.mode),
    assumeFormula: pasteSpecialModeWritesFormula(special.mode) && clipboardHasFormula(clipboard),
  });
  const cutSheetIndexes = cutPasteHistorySheetIndexes(clipboard, state.activeSheetIndex);
  const usesSheetHistory = tableEditCandidate || special.mode === "columnWidths" || cutSheetIndexes.length;
  const sheetsBefore = tableEditCandidate
    ? captureSheetsForHistory(state.model.sheets.map((_item, index) => index))
    : cutSheetIndexes.length
      ? captureSheetsForHistory(cutSheetIndexes)
      : null;
  const sheetBefore = usesSheetHistory && !sheetsBefore?.length ? captureSheetForHistory(sheet) : null;
  const before = captureCellsForHistory(sheet, targetRange);
  let changed = 0;

  try {
    ensureSheetSize(sheet, targetRows, targetCols);
    const sheetId = getSheetId(sheet.name);
    if (special.mode === "columnWidths") {
      changed = pasteClipboardColumnWidths(sheet, clipboard, range.left, targetColCount, special.transpose);
    } else {
      for (let rowOffset = 1; rowOffset <= targetRowCount; rowOffset += 1) {
        for (let colOffset = 1; colOffset <= targetColCount; colOffset += 1) {
          const sourceOffset = repeatedClipboardSourceOffset(rowOffset, colOffset, pasteSize, special.transpose);
          const sourceCell = entriesByOffset.get(cellKey(sourceOffset.row, sourceOffset.col));
          if (special.skipBlanks && isBlankClipboardCell(sourceCell)) continue;
          const row = range.top + rowOffset - 1;
          const col = range.left + colOffset - 1;
          if (special.mode === "formats") {
            pasteClipboardCellFormat(sheet, row, col, sourceCell);
          } else if (special.mode === "comments") {
            pasteClipboardCellComments(sheet, row, col, sourceCell);
          } else if (special.mode === "validation") {
            pasteClipboardCellValidation(sheet, row, col, sourceCell);
          } else if (special.mode === "allExceptBorders") {
            pasteClipboardCellAllExceptBorders(sheet, row, col, sourceCell, clipboard, sourceOffset);
          } else if (special.mode === "all" || special.mode === "allUsingSourceTheme") {
            pasteClipboardCellAllForSpecial(sheet, sheetId, row, col, clipboard, sourceCell, sourceOffset, special);
          } else {
            pasteClipboardCellContentForSpecial(sheet, sheetId, row, col, clipboard, sourceCell, sourceOffset, special);
          }
          changed += 1;
        }
      }
    }
    const tableEdit = applyTableRangeEditBehaviors(sheet, targetRange);
    clearCutSourceAfterInternalPaste(clipboard, targetRange, state.activeSheetIndex);
    state.hf = buildFormulaEngine(state.model);
    state.selected = { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
    state.selectionRange = normalizeRange(state.selected, {
      sheetIndex: state.activeSheetIndex,
      row: targetRows,
      col: targetCols,
    });
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectionDrag = null;
    const after = captureCellsForHistory(sheet, targetRange);
    const label = tableEdit.changed
      ? `${tableEdit.label} / ${changed} 個のセルを貼り付け（${pasteSpecialModeLabel(special.mode)}）`
      : `${changed} 個のセルを貼り付け（${pasteSpecialModeLabel(special.mode)}）`;
    pushHistory(usesSheetHistory
      ? {
          sheetIndex: state.activeSheetIndex,
          ...(sheetsBefore?.length
            ? { sheetsBefore, sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)) }
            : { sheetBefore, sheetAfter: captureSheetForHistory(sheet) }),
          selectionBefore: range,
          selectionAfter: state.selectionRange,
          label,
        }
      : {
          sheetIndex: state.activeSheetIndex,
          before,
          after,
          selectionBefore: range,
          selectionAfter: state.selectionRange,
          label,
        });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForCell(range.top, range.left);
    setStatus(tableEdit.changed ? `${tableEdit.status} ${changed} 個のセルを貼り付けました。` : `${changed} 個のセルを貼り付けました（${pasteSpecialModeLabel(special.mode)}）。`);
  } catch (error) {
    console.error(error);
    setStatus(`貼り付けに失敗しました：${error.message || error}`);
    renderSheet();
  }
}

export {
  activateAdjacentVisibleSheet,
  activateAdjacentVisibleSheetGroup,
  activateAdjacentWorkbookWindow,
  appendSelectionKeyboardImeText,
  appendSelectionKeyboardShapeImeText,
  applyPasteSpecialDialog,
  asyncClipboardCellTextPayload,
  asyncClipboardCwsPayload,
  asyncClipboardImageFiles,
  asyncClipboardRichHtmlClipboard,
  asyncClipboardTextPayload,
  blobToClipboardImageFile,
  canPasteInternalClipboard,
  cancelObjectPasteFallback,
  cancelSelectionKeyboardImeEdit,
  captureRangeForClipboard,
  clearCellEditDirectTyping,
  clearConsumedCutObjectClipboardLabel,
  clearCutSourceAfterInternalPaste,
  clearSelectionAndStartEdit,
  clearWorkbookClipboardState,
  clipboardImageDisplaySize,
  clipboardImageFilesFromData,
  clipboardImagePixelScale,
  clipboardImageSourceFingerprint,
  clipboardTextForInternalCompare,
  cloneCellForClipboard,
  consumeStaleCutObjectClipboardText,
  copiedObjectName,
  copySelectedImages,
  copySelectedRange,
  cutPasteHistorySheetIndexes,
  cutSelectedImages,
  cutSelectedRange,
  cwsObjectPayloadMatchesObjectClipboard,
  cwsRangePayloadMatchesInternalClipboard,
  directTypingText,
  ensureSelectionKeyboardImeEdit,
  ensureSelectionKeyboardShapeImeEdit,
  executeWorkbookClipboardShortcut,
  fillSelectionByKeyboard,
  firstPointForActiveSelection,
  formatKeyboardShortcutDate,
  formatKeyboardShortcutTime,
  handleSelectedObjectKeyboardShortcut,
  handleSelectionDirectTyping,
  handleSelectionKeyboardNavigation,
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
  handleWorkbookEditingShortcut,
  handleWorkbookKeyboardShortcut,
  handleWorkbookPasteEvent,
  hideSelectionKeyboardSink,
  homeFillKeyboardPlanForRange,
  insertKeyboardDateTimeByKeyboard,
  insertKeyboardDateTimeIntoEditable,
  internalClipboardTextMatches,
  isAsteriskShortcutKey,
  isBackquoteShortcutKey,
  isBackslashShortcutKey,
  isClipboardImageFile,
  isColonShortcutKey,
  isConsumedCutObjectClipboardText,
  isElementImeCompositionEvent,
  isImeCompositionEvent,
  isInternalCutClipboard,
  isSelectionKeyboardSink,
  isSelectionKeyboardSinkEligibleActiveElement,
  isSemicolonShortcutKey,
  isSheetSelectionShortcutTarget,
  isSpaceKeyEvent,
  isTextEditingTarget,
  keyboardDateTimeShortcutKind,
  keyboardDateTimeShortcutText,
  markCellEditDirectTyping,
  markConsumedCutObjectClipboardLabel,
  markDirectTypingComposition,
  markElementImeComposition,
  markImeComposition,
  markObjectClipboardSystemSynced,
  nowMs,
  objectClipboardTextLabel,
  objectPasteTargetPoint,
  openCellStructureDialogForShortcut,
  openFormatCellsDialogForShortcut,
  openPasteSpecialDialog,
  pad2,
  pasteClipboardImageFiles,
  pasteCwsClipboardPayload,
  pasteExternalRichClipboard,
  pasteFromClipboard,
  pasteFromRibbonButton,
  pasteInternalClipboardSpecialAtRange,
  pasteObjectClipboard,
  pasteSpecialOptionsFromDialog,
  pasteTargetSizeForClipboard,
  positionSelectionKeyboardSink,
  positionSelectionKeyboardSinkForShape,
  promoteSelectionKeyboardImeEditToCellEditor,
  promoteSelectionKeyboardShapeImeEditToEditor,
  rangeToTsv,
  reconcileDirectTypingImeCommit,
  recordDirectTypingCompositionData,
  repeatedClipboardSourceOffset,
  resetSelectionKeyboardSinkValue,
  restoreCwsClipboardPayload,
  restorePersistedInternalClipboard,
  scheduleInternalPasteFallback,
  scheduleObjectPasteFallback,
  scheduleSelectionKeyboardSinkFocus,
  selectAllActivePoint,
  selectAllByKeyboard,
  selectAllByKeyboardWithCache,
  selectAllCurrentRegionCandidate,
  selectAllExpansionCandidates,
  selectAllNextCandidate,
  selectAllTableCandidates,
  selectAllUsedRangeCandidate,
  selectAllWorksheetRange,
  selectColumnsByKeyboard,
  selectRangesByKeyboard,
  selectRowsByKeyboard,
  selectedObjectSummary,
  selectedPointForKeyboardSink,
  selectedShapeForKeyboardSink,
  selectionKeyboardSinkState,
  sheetUsedRangeForSelectAll,
  shouldFocusSelectionKeyboardSink,
  shouldPreferUnsyncedLocalObjectClipboardPaste,
  shouldUseInternalClipboardPaste,
  shouldWriteAsyncClipboard,
  showNoHomeFillSourceMessage,
  syncSelectionKeyboardSinkEditWithSelection,
  tableDataBodySelectionRange,
  uniqueSelectAllCandidates,
  writeCwsPayloadToAsyncClipboard,
};
