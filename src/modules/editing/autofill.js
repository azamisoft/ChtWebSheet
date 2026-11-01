// Feature module; shared bindings are imported explicitly.
import {
  $,
  $formulaInput,
  $sheetHost,
  AUTO_FILL_DATE_MODES,
  AUTO_FILL_LISTS,
  AUTO_FILL_MODES,
  AUTO_FILL_OPTIONS_BUTTON_SIZE,
  AUTO_FILL_OPTIONS_MENU_WIDTH,
  AUTO_FILL_OPTIONS_VIEWPORT_MARGIN,
  AUTO_FILL_TREND_MODES,
  CELL_EDIT_TRAILING_LINE_SENTINEL,
  DEFAULT_SHAPE_TYPE,
  RIBBON_UNDERLINE_PRESETS,
  SHAPE_TEXT_EDITOR_Z_INDEX,
  state,
} from "../../app-core.js";
import {
  startShapeDrawing,
} from "../drawing/charts-and-insertion.js";
import {
  isLineShapeObject,
  normalizeShapeTextMargin,
  normalizeShapeTextOutlineWidth,
  shapeDescriptorForType,
} from "../drawing/shape-rendering.js";
import {
  lineShapePointerHitsStroke,
  primaryImagePositionElement,
  refreshShapeSource,
  selectionIdsForImage,
  startSelectedImageTransform,
} from "../drawing/transforms-and-connectors.js";
import {
  normalizeColorValue,
  setRibbonUnderlinePreset,
  underlinePresetForTextDecoration,
  underlinePresetForValue,
} from "./ink-and-clipboard.js";
import {
  clearConsumedCutObjectClipboardLabel,
  isElementImeCompositionEvent,
  isSelectionKeyboardSink,
  isTextEditingTarget,
  nowMs,
  selectedObjectSummary,
} from "./keyboard-and-paste.js";
import {
  activateObjectFormatRibbonForSelection,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  cellPointFromClientPosition,
  cellPointFromElement,
  cellPointFromPointerEvent,
  expandRangeForMerges,
  formatPainterPointFromPointerEvent,
  getDisplayForCell,
  normalizeRange,
  rangeToLabel,
  restoreGroupedSheetsAfterFailure,
  selectionRangeFromPoints,
  syncShapeFormatControls,
  updateSelectionUi,
} from "./pointers-and-formula-evaluation.js";
import {
  activeSelectedImageIds,
  applySelectionRange,
  cellEditBoundaryForTextOffset,
  cellEditRunHtml,
  cellEditSelectionOffsets,
  coordinateToPixels,
  editablePlainText,
  imageBoxStyle,
  imageBoxToPixels,
  isActiveImage,
  isBoldCssValue,
  isTextInputElement,
  placeCaretAtEnd,
  placeCaretAtPoint,
  rawActiveSelectionRanges,
  restoreCellEditSelection,
  saveCellEditSelection,
  selectedImageObjectById,
  selectedImageObjects,
  selectedInkStrokeIdSet,
  selectedInkStrokes,
  setImageElementRenderSource,
  setSelectedImageIds,
  shapeTextPlainValue,
  stripCellEditSentinels,
  syncEditableCellAlignment,
  updateFormulaBarForCell,
  updateFormulaBarForSelection,
  worksheetObjectLayerZIndex,
} from "./selection-and-cell-text.js";
import {
  captureCellsForHistory,
  captureSheetForHistory,
  captureSheetsForHistory,
  cellEffectiveCss,
  cloneCellModel,
  cloneJsonLikeValue,
  cloneMerges,
  normalizeInkStrokes,
  pushHistory,
  removeSheetStyleRangesAtCell,
  sheetDisplayCell,
  shiftRangeColumns,
  shiftRangeRows,
  shiftStructureIndex,
  stableObjectKey,
  storeCellModelForSheet,
} from "./structure-and-history.js";
import {
  clearFormulaReferenceState,
  formulaReferenceTokensInText,
  startFormulaReferenceSelection,
} from "../formula/dialogs-and-names.js";
import {
  forEachCellInRange,
  normalizeRichTextLineEndings,
  rangeIntersects,
  rangesIntersect,
  removeIntersectingMerges,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  closeCellContextMenu,
  rangeHasFormulaContent,
} from "../ui/annotations-and-cell-format.js";
import {
  accessKeyLabelHtml,
  cellKey,
  clearAppAlert,
  columnName,
  columnNumber,
  dateDetectionFormat,
  dateToExcelSerial,
  decodeRange,
  escapeAttr,
  excelSerialToDate,
  isDateFormat,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  consumeCommandMenuKey,
  eraseInkAtClientPoint,
  startInkDrawing,
} from "../ui/printing-and-home-tools.js";
import {
  activeCellEditElementForFormatting,
  closeFileSideMenu,
  ensureSheetSize,
  minimumSheetExtent,
} from "../ui/ribbon-and-links.js";
import {
  closeSheetTabContextMenu,
  ensureEditableRangeForProtection,
  ensureSelectableRangeForProtection,
  ensureSheetProtectionPermission,
  openImageContextMenu,
} from "../ui/sheet-protection-and-translation.js";
import {
  buildFormulaEngine,
  fontFamilyForExcel,
  normalizeDefinedNameLookup,
  rgbToHex,
  unquoteSheetName,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  dateSerialFromPersistedValue,
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  activeSheet,
  getSheetId,
  isOccupiedCell,
  primaryExcelFormatSection,
} from "../workbook/saving-and-export.js";
import {
  applyTableRangeEditBehaviors,
  parseValidationOptions,
  rangeRectEquals,
  tableRangeEditMayChangeMetadata,
} from "../workbook/tables-and-validation.js";


function applyCellEditCssPatchToSelection(cssPatch, label) {
  const editor = activeCellEditElementForFormatting();
  if (!editor) return false;
  const patch = normalizeCellEditCssPatch(cssPatch || {});
  if (!Object.keys(patch).length) return false;
  const offsets = cellEditSelectionOffsets(editor);
  if (offsets.start === offsets.end) {
    editor.focus({ preventScroll: true });
    restoreCellEditSelection(editor);
    setStatus("編集中文字を選択してください。");
    return true;
  }
  const runs = cellEditRunsFromDom(editor);
  const nextRuns = applyCssPatchToRichTextRuns(runs, offsets.start, offsets.end, patch);
  const text = nextRuns.map((run) => run.text).join("");
  editor.innerHTML = cellEditRunsHtml(nextRuns, text);
  syncEditableCellAlignment(editor);
  editor.focus({ preventScroll: true });
  $(editor).removeData("skipBlurCommit");
  setCellEditSelectionByOffsets(editor, offsets.start, offsets.end);
  editor.dispatchEvent(new Event("input", { bubbles: true }));
  saveCellEditSelection(editor);
  setStatus(`選択した文字に適用しました：${label}`);
  return true;
}



function toggleCellEditCssForSelection(property, activeValue, inactiveValue, label) {
  const editor = activeCellEditElementForFormatting();
  if (!editor) return false;
  const offsets = cellEditSelectionOffsets(editor);
  if (offsets.start === offsets.end) {
    editor.focus({ preventScroll: true });
    restoreCellEditSelection(editor);
    setStatus("編集中文字を選択してください。");
    return true;
  }
  const active = cellEditSelectionEveryRunMatches(editor, offsets, (css) => {
    if (property === "fontWeight") return isBoldCssValue(css.fontWeight);
    if (property === "fontStyle") return css.fontStyle === "italic";
    return css[property] === activeValue;
  });
  return applyCellEditCssPatchToSelection({ [property]: active ? inactiveValue : activeValue }, label);
}



function applyCellEditUnderlinePreset(preset, options = {}, label = "") {
  const editor = activeCellEditElementForFormatting();
  if (!editor) return false;
  const selectedPreset = underlinePresetForValue(preset) || RIBBON_UNDERLINE_PRESETS[0];
  const offsets = cellEditSelectionOffsets(editor);
  if (offsets.start === offsets.end) {
    editor.focus({ preventScroll: true });
    restoreCellEditSelection(editor);
    setStatus("編集中文字を選択してください。");
    return true;
  }
  const active = cellEditSelectionEveryRunMatches(editor, offsets, (css) => underlinePresetForTextDecoration(css.textDecoration)?.value === selectedPreset.value);
  const nextValue = options.toggle && active ? "none" : selectedPreset.value;
  setRibbonUnderlinePreset(selectedPreset.value);
  return applyCellEditCssPatchToSelection({ textDecoration: nextValue }, nextValue === "none" ? "下線なし" : label || selectedPreset.label.replace(/\(.+\)/, ""));
}



function cellEditSelectionEveryRunMatches(editor, offsets, predicate) {
  const baseCss = cellEditBaseCss(editor);
  let hasText = false;
  let every = true;
  let cursor = 0;
  cellEditRunsFromDom(editor).forEach((run) => {
    const start = cursor;
    const end = cursor + run.text.length;
    cursor = end;
    if (end <= offsets.start || start >= offsets.end || !run.text) return;
    hasText = true;
    if (!predicate({ ...baseCss, ...(run.css || {}) })) every = false;
  });
  return hasText && every;
}



function cellEditBaseCss(editor) {
  const point = cellPointFromElement(editor);
  if (!point) return {};
  const sheet = activeSheet();
  return cellEffectiveCss(sheet, sheet.cells[cellKey(point.row, point.col)] || {});
}



function cellEditRunsFromDom(editor) {
  const runs = [];
  editor?.childNodes?.forEach((child) => collectCellEditRuns(child, {}, runs));
  return mergeCellEditRuns(runs);
}



function collectCellEditRuns(node, inheritedCss, runs) {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = stripCellEditSentinels(node.nodeValue || "");
    if (text) pushCellEditRun(runs, text, inheritedCss);
    return;
  }
  if (node.nodeType !== Node.ELEMENT_NODE && node.nodeType !== Node.DOCUMENT_FRAGMENT_NODE) return;
  const elementCss = node.nodeType === Node.ELEMENT_NODE ? cellEditInlineCssFromElement(node) : {};
  const nextCss = normalizeCellEditInlineCss({ ...(inheritedCss || {}), ...elementCss }, { keepNormal: true });
  node.childNodes.forEach((child) => collectCellEditRuns(child, nextCss, runs));
}



function cellEditInlineCssFromElement(element) {
  const css = {};
  const tag = String(element.tagName || "").toLowerCase();
  if (tag === "b" || tag === "strong") css.fontWeight = "700";
  if (tag === "i" || tag === "em") css.fontStyle = "italic";
  if (tag === "u") css.textDecoration = "underline";
  const style = element.style || {};
  if (style.fontFamily) css.fontFamily = style.fontFamily;
  if (style.fontSize) css.fontSize = style.fontSize;
  if (style.color) css.color = normalizeCellEditColor(style.color);
  if (style.fontWeight) css.fontWeight = style.fontWeight;
  if (style.fontStyle) css.fontStyle = style.fontStyle;
  const decoration = style.textDecoration || [style.textDecorationLine, style.textDecorationStyle].filter(Boolean).join(" ");
  if (decoration) css.textDecoration = decoration;
  return normalizeCellEditInlineCss(css, { keepNormal: true });
}



function normalizeCellEditInlineCss(css = {}, options = {}) {
  const normalized = {};
  const keepNormal = options.keepNormal !== false;
  if (css.fontFamily) normalized.fontFamily = String(css.fontFamily);
  if (css.fontSize) normalized.fontSize = String(css.fontSize);
  const color = normalizeCellEditColor(css.color);
  if (color) normalized.color = color;
  if (Object.prototype.hasOwnProperty.call(css, "fontWeight")) {
    const value = String(css.fontWeight || "").toLowerCase();
    if (value === "bold" || value === "700" || Number.parseInt(value, 10) >= 600) normalized.fontWeight = "700";
    else if (keepNormal && (value === "normal" || value === "400")) normalized.fontWeight = "normal";
  }
  if (Object.prototype.hasOwnProperty.call(css, "fontStyle")) {
    const value = String(css.fontStyle || "").toLowerCase();
    if (value === "italic") normalized.fontStyle = "italic";
    else if (keepNormal && value === "normal") normalized.fontStyle = "normal";
  }
  if (Object.prototype.hasOwnProperty.call(css, "textDecoration")) {
    const value = String(css.textDecoration || "").toLowerCase();
    if (value.includes("underline")) normalized.textDecoration = value.includes("double") ? "underline double" : "underline";
    else if (value.includes("line-through")) normalized.textDecoration = "line-through";
    else if (keepNormal && (value === "none" || value === "normal")) normalized.textDecoration = "none";
  }
  return normalized;
}



function normalizeCellEditCssPatch(cssPatch = {}) {
  const patch = {};
  ["fontFamily", "fontSize", "color", "fontWeight", "fontStyle", "textDecoration"].forEach((key) => {
    if (!Object.prototype.hasOwnProperty.call(cssPatch, key)) return;
    const value = cssPatch[key];
    if (value == null || value === "") {
      patch[key] = null;
      return;
    }
    const normalized = normalizeCellEditInlineCss({ [key]: value }, { keepNormal: true });
    if (Object.prototype.hasOwnProperty.call(normalized, key)) patch[key] = normalized[key];
  });
  return patch;
}



function normalizeCellEditColor(value) {
  const normalized = normalizeColorValue(value);
  if (normalized) return normalized;
  const match = String(value || "").match(/^rgba?\(([^)]+)\)$/i);
  if (!match) return "";
  const channels = match[1].split(",").slice(0, 3).map((part) => Math.max(0, Math.min(255, Math.round(Number.parseFloat(part)))));
  return channels.length === 3 && channels.every((channel) => Number.isFinite(channel)) ? rgbToHex(channels) : "";
}



function pushCellEditRun(runs, text, css = {}) {
  const cleanCss = normalizeCellEditInlineCss(css, { keepNormal: true });
  const last = runs[runs.length - 1];
  if (last && stableObjectKey(last.css || {}) === stableObjectKey(cleanCss)) {
    last.text += text;
  } else {
    runs.push(Object.keys(cleanCss).length ? { text, css: cleanCss } : { text });
  }
}



function mergeCellEditRuns(runs = []) {
  const merged = [];
  (Array.isArray(runs) ? runs : []).forEach((run) => {
    const text = stripCellEditSentinels(run?.text || "");
    if (!text) return;
    pushCellEditRun(merged, text, run?.css || {});
  });
  return merged;
}



function applyCssPatchToRichTextRuns(runs, startOffset, endOffset, patch) {
  const next = [];
  let cursor = 0;
  runs.forEach((run) => {
    const text = run.text || "";
    const runStart = cursor;
    const runEnd = cursor + text.length;
    cursor = runEnd;
    if (!text || runEnd <= startOffset || runStart >= endOffset) {
      pushCellEditRun(next, text, run.css || {});
      return;
    }
    const localStart = Math.max(0, startOffset - runStart);
    const localEnd = Math.min(text.length, endOffset - runStart);
    if (localStart > 0) pushCellEditRun(next, text.slice(0, localStart), run.css || {});
    pushCellEditRun(next, text.slice(localStart, localEnd), applyCellEditCssPatch(run.css || {}, patch));
    if (localEnd < text.length) pushCellEditRun(next, text.slice(localEnd), run.css || {});
  });
  return mergeCellEditRuns(next);
}



function applyCellEditCssPatch(css, patch) {
  const next = { ...(css || {}) };
  Object.entries(patch || {}).forEach(([key, value]) => {
    if (value == null || value === "") delete next[key];
    else next[key] = value;
  });
  return normalizeCellEditInlineCss(next, { keepNormal: true });
}



function cellEditRunsHtml(runs, text) {
  const html = mergeCellEditRuns(runs).map((run) => cellEditRunHtml(run)).join("");
  return String(text || "").endsWith("\n") ? `${html}${CELL_EDIT_TRAILING_LINE_SENTINEL}` : html;
}



function setCellEditSelectionByOffsets(editor, start, end = start) {
  const selection = window.getSelection?.();
  if (!selection) return false;
  const range = document.createRange();
  const startBoundary = cellEditBoundaryForTextOffset(editor, start);
  const endBoundary = cellEditBoundaryForTextOffset(editor, end);
  try {
    range.setStart(startBoundary.node, startBoundary.offset);
    range.setEnd(endBoundary.node, endBoundary.offset);
  } catch {
    return false;
  }
  selection.removeAllRanges();
  selection.addRange(range);
  state.cellEditSelection = { editor, range: range.cloneRange(), start, end };
  return true;
}



function cellEditRichTextForCommit(editor, text = editablePlainText(editor)) {
  return normalizeCellEditRichTextForModel(cellEditRunsFromDom(editor), text);
}



function normalizeCellEditRichTextForModel(runs, text) {
  const merged = mergeCellEditRuns(runs);
  const plain = normalizeRichTextLineEndings(merged.map((run) => run.text).join(""));
  if (plain !== normalizeRichTextLineEndings(text)) return null;
  if (!merged.some((run) => Object.keys(run.css || {}).length)) return null;
  return merged.map((run) => (Object.keys(run.css || {}).length ? { text: run.text, css: run.css } : { text: run.text }));
}



function insertLineBreakInTextControl(element) {
  if (!element) return false;
  const value = String(element.value ?? "");
  const start = Math.max(0, Math.min(value.length, Number(element.selectionStart ?? value.length)));
  const end = Math.max(start, Math.min(value.length, Number(element.selectionEnd ?? start)));
  element.value = `${value.slice(0, start)}\n${value.slice(end)}`;
  const caret = start + 1;
  element.setSelectionRange?.(caret, caret);
  try {
    element.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertLineBreak" }));
  } catch {
    element.dispatchEvent(new Event("input", { bubbles: true }));
  }
  return true;
}



function isEnterKeyEvent(event) {
  const nativeEvent = event?.originalEvent || event || {};
  const key = String(event?.key || nativeEvent.key || "");
  const code = String(event?.code || nativeEvent.code || "");
  const keyCode = Number(nativeEvent.keyCode || nativeEvent.which || event?.keyCode || event?.which || 0);
  return key === "Enter" || key === "NumpadEnter" || code === "Enter" || code === "NumpadEnter" || keyCode === 13;
}



function isAltModifierEvent(event) {
  const nativeEvent = event?.originalEvent || event || {};
  return Boolean(event?.altKey || nativeEvent.altKey || event?.getModifierState?.("Alt") || nativeEvent.getModifierState?.("Alt"));
}



function startAutoFillDrag(event) {
  if ((event.button !== 0 && event.button !== 2) || !state.model) return;
  const ranges = activeSelectionRanges();
  const frame = event.currentTarget.closest(".sheet-selection-frame");
  const rangeIndex = Number(frame?.dataset.rangeIndex || 0);
  const sourceRange = ranges[Number.isFinite(rangeIndex) ? rangeIndex : 0];
  if (!sourceRange || sourceRange.fullRow || sourceRange.fullColumn) return;
  const sheet = activeSheet();

  event.preventDefault();
  event.stopPropagation();
  closeCellContextMenu();
  closeSheetTabContextMenu();
  closeFileSideMenu({ restoreHome: true });
  clearAutoFillOptions();
  clearAppAlert();

  state.fillDrag = {
    sourceRange: { ...sourceRange },
    previousSelected: state.selected ? { ...state.selected } : null,
    previousRange: state.selectionRange ? { ...state.selectionRange } : null,
    previousRanges: rawActiveSelectionRanges().map((range) => ({ ...range })),
    targetRange: { ...sourceRange },
    fillRange: null,
    direction: null,
    mode: autoFillModeForPointerEvent(sheet, sourceRange, event),
    rightButton: event.button === 2,
  };
  state.suppressNextCellClick = true;
  setPointerCaptureSafely(event.currentTarget, event.pointerId);
}



function updateAutoFillDrag(event) {
  const drag = state.fillDrag;
  if (!drag || !state.model) return;
  drag.mode = autoFillModeForPointerEvent(activeSheet(), drag.sourceRange, event);
  const point = cellPointFromClientPosition(event.clientX, event.clientY);
  if (!point) return;
  const next = autoFillRangeFromPoint(drag.sourceRange, point);
  if (!next || rangesEqualLoose(next.targetRange, drag.targetRange)) return;

  drag.targetRange = next.targetRange;
  drag.fillRange = next.fillRange;
  drag.direction = next.direction;
  state.selectionRange = next.targetRange;
  state.selectionRanges = [next.targetRange];
  state.selected = { sheetIndex: state.activeSheetIndex, row: next.targetRange.top, col: next.targetRange.left };
  state.selectionStart = state.selected;
  updateSelectionUi();
}



function finishAutoFillDrag(event) {
  const drag = state.fillDrag;
  if (!drag) return;
  event?.preventDefault?.();
  event?.stopPropagation?.();
  state.fillDrag = null;

  if (!drag.fillRange || rangesEqualLoose(drag.targetRange, drag.sourceRange)) {
    restoreAutoFillSourceSelection(drag);
    return;
  }

  drag.mode = autoFillModeForPointerEvent(activeSheet(), drag.sourceRange, event, drag.mode);
  if (drag.rightButton) {
    state.suppressAutoFillContextMenuUntil = nowMs() + 500;
    openAutoFillRightDragOptions(drag, event);
    return;
  }
  const defaultMode = defaultAutoFillMode(activeSheet(), drag.sourceRange);
  const result = applyAutoFill(drag.sourceRange, drag.fillRange, drag.targetRange, drag.direction, {
    mode: drag.mode,
    explicitMode: drag.mode !== defaultMode,
  });
  if (!result.ok) {
    restoreAutoFillSourceSelection(drag);
    showMessageBox(result.message, { title: "Cht WebSheet" });
    setStatus(result.message);
    updateSelectionUi();
  }
}



function suppressAutoFillRightDragContextMenu(event) {
  if (nowMs() > Number(state.suppressAutoFillContextMenuUntil || 0)) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation();
}



function openAutoFillRightDragOptions(drag, event = null) {
  const sheet = activeSheet();
  if (!sheet) return;
  const mergeValidation = validateAutoFillMerges(sheet, drag.sourceRange, drag.fillRange, drag.direction);
  if (!mergeValidation.ok) {
    restoreAutoFillSourceSelection(drag);
    showMessageBox(mergeValidation.message, { title: "Cht WebSheet" });
    setStatus(mergeValidation.message);
    updateSelectionUi();
    return;
  }
  state.selectionRange = { ...drag.targetRange };
  state.selectionRanges = [{ ...drag.targetRange }];
  state.selected = { sheetIndex: state.activeSheetIndex, row: drag.targetRange.top, col: drag.targetRange.left };
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.isSelecting = false;
  state.autoFillOptions = {
    sheetIndex: state.activeSheetIndex,
    sourceRange: { ...drag.sourceRange },
    fillRange: { ...drag.fillRange },
    targetRange: { ...drag.targetRange },
    direction: drag.direction,
    before: captureCellsForHistory(sheet, drag.targetRange).map((item) => ({ row: item.row, col: item.col, cell: cloneCellModel(item.cell) })),
    mode: drag.mode,
    menuOpen: true,
    menuPoint: autoFillMenuPointFromEvent(sheet, drag.targetRange, event),
    pending: true,
  };
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  focusAutoFillOptionsMenuItem(state.autoFillOptions?.mode);
  queueMicrotask(() => focusAutoFillOptionsMenuItem(state.autoFillOptions?.mode));
  requestAnimationFrame(() => focusAutoFillOptionsMenuItem(state.autoFillOptions?.mode));
  setTimeout(() => focusAutoFillOptionsMenuItem(state.autoFillOptions?.mode), 0);
  setTimeout(() => focusAutoFillOptionsMenuItem(state.autoFillOptions?.mode), 30);
  setStatus(`${rangeToLabel(drag.targetRange)} のオートフィル オプションを選択してください。`);
}



function autoFillMenuPointFromEvent(sheet, targetRange, event = null) {
  const host = $sheetHost[0];
  if (host && event && Number.isFinite(Number(event.clientX)) && Number.isFinite(Number(event.clientY))) {
    const rect = host.getBoundingClientRect();
    return {
      left: Math.max(0, event.clientX - rect.left + host.scrollLeft),
      top: Math.max(0, event.clientY - rect.top + host.scrollTop),
    };
  }
  return {
    left: coordinateToPixels(sheet, "col", targetRange.right) + 3,
    top: coordinateToPixels(sheet, "row", targetRange.bottom) + 22,
  };
}



function handleAutoFillHandleDoubleClick(event) {
  event.preventDefault();
  event.stopPropagation();
  if (!state.model) return;
  const ranges = activeSelectionRanges();
  const frame = event.currentTarget.closest(".sheet-selection-frame");
  const rangeIndex = Number(frame?.dataset.rangeIndex || 0);
  const sourceRange = ranges[Number.isFinite(rangeIndex) ? rangeIndex : 0];
  if (!sourceRange || sourceRange.fullRow || sourceRange.fullColumn) return;
  const sheet = activeSheet();
  const autoRange = autoFillDoubleClickRange(sheet, sourceRange);
  if (!autoRange) {
    setStatus("隣接するデータ範囲がないため、オートフィルできませんでした。");
    return;
  }
  clearAutoFillOptions();
  clearAppAlert();
  const result = applyAutoFill(autoRange.sourceRange, autoRange.fillRange, autoRange.targetRange, autoRange.direction, {
    mode: autoFillModeForPointerEvent(sheet, sourceRange, event),
  });
  if (!result.ok) {
    showMessageBox(result.message, { title: "Cht WebSheet" });
    setStatus(result.message);
  }
}



function restoreAutoFillSourceSelection(drag) {
  state.selected = drag.previousSelected;
  state.selectionRange = drag.previousRange;
  state.selectionRanges = drag.previousRanges;
  state.selectionStart = drag.previousSelected;
  state.selectionDrag = null;
  state.isSelecting = false;
  updateSelectionUi();
  if (state.selected) updateFormulaBarForCell(state.selected.row, state.selected.col);
}



function autoFillRangeFromPoint(sourceRange, point) {
  const sheet = activeSheet();
  const endpoint = expandRangeForMerges(sheet, normalizeRange(point, point));
  if (!endpoint) return null;
  const distances = [
    { direction: "down", value: Math.max(0, endpoint.bottom - sourceRange.bottom) },
    { direction: "up", value: Math.max(0, sourceRange.top - endpoint.top) },
    { direction: "right", value: Math.max(0, endpoint.right - sourceRange.right) },
    { direction: "left", value: Math.max(0, sourceRange.left - endpoint.left) },
  ].sort((a, b) => b.value - a.value);
  const direction = distances[0]?.value > 0 ? distances[0].direction : null;
  if (!direction) return { direction: null, targetRange: { ...sourceRange }, fillRange: null };

  let targetRange = { ...sourceRange };
  if (direction === "down") targetRange.bottom = Math.max(sourceRange.bottom, endpoint.bottom);
  if (direction === "up") targetRange.top = Math.min(sourceRange.top, endpoint.top);
  if (direction === "right") targetRange.right = Math.max(sourceRange.right, endpoint.right);
  if (direction === "left") targetRange.left = Math.min(sourceRange.left, endpoint.left);
  targetRange = expandRangeForMerges(sheet, targetRange);

  let fillRange = null;
  if (direction === "down" && targetRange.bottom > sourceRange.bottom) {
    fillRange = { ...targetRange, top: sourceRange.bottom + 1 };
  } else if (direction === "up" && targetRange.top < sourceRange.top) {
    fillRange = { ...targetRange, bottom: sourceRange.top - 1 };
  } else if (direction === "right" && targetRange.right > sourceRange.right) {
    fillRange = { ...targetRange, left: sourceRange.right + 1 };
  } else if (direction === "left" && targetRange.left < sourceRange.left) {
    fillRange = { ...targetRange, right: sourceRange.left - 1 };
  }
  return { direction, targetRange, fillRange };
}



function applyAutoFill(sourceRange, fillRange, targetRange, direction, options = {}) {
  const sheet = activeSheet();
  if (!ensureEditableRangeForProtection(sheet, fillRange, "オートフィル")) return { ok: false, message: "保護されたシートではオートフィルできません。" };
  const mergeValidation = validateAutoFillMerges(sheet, sourceRange, fillRange, direction);
  if (!mergeValidation.ok) return mergeValidation;

  const tableEditCandidate = tableRangeEditMayChangeMetadata(sheet, fillRange, {
    assumeNonEmpty: true,
    assumeFormula: rangeHasFormulaContent(sheet, sourceRange),
  });
  const sheetBefore = tableEditCandidate ? captureSheetForHistory(sheet) : null;
  const sheetsBefore = tableEditCandidate ? captureSheetsForHistory(state.model.sheets.map((_item, index) => index)) : null;
  const before = captureCellsForHistory(sheet, targetRange);
  const mergesBefore = cloneMerges(sheet.merges);
  ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
  const mode = normalizeAutoFillMode(options.mode) || defaultAutoFillMode(sheet, sourceRange);
  removeIntersectingMerges(sheet, fillRange);
  writeAutoFillCells(sheet, sourceRange, fillRange, direction, {
    mode,
    explicitMode: options.explicitMode === true,
  });
  applyAutoFillMerges(sheet, sourceRange, fillRange, direction);
  const tableEdit = applyTableRangeEditBehaviors(sheet, fillRange);
  state.hf = buildFormulaEngine(state.model);

  const after = captureCellsForHistory(sheet, targetRange);
  const mergesAfter = cloneMerges(sheet.merges);
  state.selected = { sheetIndex: state.activeSheetIndex, row: targetRange.top, col: targetRange.left };
  state.selectionRange = targetRange;
  state.selectionRanges = [targetRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.autoFillOptions = {
    sheetIndex: state.activeSheetIndex,
    sourceRange: { ...sourceRange },
    fillRange: { ...fillRange },
    targetRange: { ...targetRange },
    direction,
    before: before.map((item) => ({ row: item.row, col: item.col, cell: cloneCellModel(item.cell) })),
    mode,
    menuOpen: false,
  };
  const label = tableEdit.changed
    ? `${tableEdit.label} / ${rangeToLabel(sourceRange)} オートフィル -> ${rangeToLabel(targetRange)}`
    : `${rangeToLabel(sourceRange)} オートフィル -> ${rangeToLabel(targetRange)}`;
  pushHistory(tableEdit.changed
    ? {
        sheetIndex: state.activeSheetIndex,
        ...(sheetsBefore?.length
          ? { sheetsBefore, sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)) }
          : { sheetBefore, sheetAfter: captureSheetForHistory(sheet) }),
        selectionBefore: sourceRange,
        selectionAfter: targetRange,
        label,
      }
    : {
        sheetIndex: state.activeSheetIndex,
        before,
        after,
        mergesBefore,
        mergesAfter,
        selectionBefore: sourceRange,
        selectionAfter: targetRange,
        label,
      });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  clearAppAlert();
  setStatus(tableEdit.changed ? `${tableEdit.status} ${rangeToLabel(targetRange)} にオートフィルしました。` : `${rangeToLabel(targetRange)} にオートフィルしました。`);
  return { ok: true };
}



function renderAutoFillOptions(sheet, visibleRange) {
  const options = state.autoFillOptions;
  if (!options || options.sheetIndex !== state.activeSheetIndex || !rangeIntersects(options.targetRange, visibleRange)) return "";
  const buttonPoint = clampAutoFillOptionsButtonPoint(
    coordinateToPixels(sheet, "col", options.targetRange.right) + 3,
    coordinateToPixels(sheet, "row", options.targetRange.bottom) + 3,
  );
  const { left, top } = buttonPoint;
  const buttonStyle = `left:${left}px;top:${top}px`;
  const menuItems = autoFillOptionItemsForSource(sheet, options.sourceRange);
  const rawMenuLeft = Number.isFinite(Number(options.menuPoint?.left)) ? Number(options.menuPoint.left) : left;
  const rawMenuTop = Number.isFinite(Number(options.menuPoint?.top)) ? Number(options.menuPoint.top) : top + 22;
  const menuPoint = clampAutoFillOptionsMenuPoint(rawMenuLeft, rawMenuTop, menuItems.length);
  const menu = options.menuOpen ? renderAutoFillOptionsMenu(menuPoint.left, menuPoint.top, options.mode, sheet, options.sourceRange, menuItems) : "";
  return `
    <button id="autoFillOptionsButton" class="auto-fill-options-button" type="button" style="${escapeAttr(buttonStyle)}" aria-haspopup="menu" aria-expanded="${options.menuOpen ? "true" : "false"}" title="オートフィル オプション">▣</button>
    ${menu}
  `;
}



function clampAutoFillOptionsButtonPoint(left, top) {
  const host = $sheetHost[0];
  if (!host) return { left, top };
  return clampAutoFillOptionsPoint(left, top, AUTO_FILL_OPTIONS_BUTTON_SIZE, AUTO_FILL_OPTIONS_BUTTON_SIZE);
}



function clampAutoFillOptionsMenuPoint(left, top, itemCount = 4) {
  const menuHeight = Math.max(34, Number(itemCount || 0) * 26 + 8);
  return clampAutoFillOptionsPoint(left, top, AUTO_FILL_OPTIONS_MENU_WIDTH, menuHeight);
}



function clampAutoFillOptionsPoint(left, top, width, height) {
  const host = $sheetHost[0];
  if (!host) return { left, top };
  const margin = AUTO_FILL_OPTIONS_VIEWPORT_MARGIN;
  const minLeft = host.scrollLeft + margin;
  const minTop = host.scrollTop + margin;
  const maxLeft = host.scrollLeft + Math.max(margin, host.clientWidth - width - margin);
  const maxTop = host.scrollTop + Math.max(margin, host.clientHeight - height - margin);
  return {
    left: Math.round(Math.min(Math.max(Number(left) || minLeft, minLeft), maxLeft)),
    top: Math.round(Math.min(Math.max(Number(top) || minTop, minTop), maxTop)),
  };
}



function renderAutoFillOptionsMenu(left, top, activeMode, sheet = activeSheet(), sourceRange = state.autoFillOptions?.sourceRange, itemsOverride = null) {
  const items = itemsOverride || autoFillOptionItemsForSource(sheet, sourceRange);
  const buttons = items.map((item) => {
    const checked = activeMode === item.mode ? " is-active" : "";
    return `<button class="auto-fill-options-menu-item${checked}" type="button" role="menuitemradio" aria-checked="${activeMode === item.mode ? "true" : "false"}" data-auto-fill-option="${item.mode}" data-access-key="${item.accessKey}"><span class="auto-fill-options-check">${activeMode === item.mode ? "✓" : ""}</span>${accessKeyLabelHtml(item.label)}</button>`;
  }).join("");
  return `<div id="autoFillOptionsMenu" class="auto-fill-options-menu" role="menu" tabindex="-1" style="${escapeAttr(`left:${left}px;top:${top}px`)}">${buttons}</div>`;
}



function autoFillOptionItemsForSource(sheet, sourceRange) {
  const items = [
    { mode: "copy", label: "セルのコピー(C)", accessKey: "C" },
    { mode: "series", label: "連続データ(S)", accessKey: "S" },
    { mode: "formatting", label: "書式のみコピー(F)", accessKey: "F" },
    { mode: "no-formatting", label: "書式なしコピー(O)", accessKey: "O" },
  ];
  if (autoFillSourceSupportsDateOptions(sheet, sourceRange)) {
    items.push(
      { mode: "date-day", label: "日単位(D)", accessKey: "D" },
      { mode: "date-weekday", label: "週日単位(W)", accessKey: "W" },
      { mode: "date-month", label: "月単位(M)", accessKey: "M" },
      { mode: "date-year", label: "年単位(Y)", accessKey: "Y" },
    );
  }
  if (autoFillSourceSupportsNumericTrendOptions(sheet, sourceRange, "linear-trend")) {
    items.push({ mode: "linear-trend", label: "線形近似(L)", accessKey: "L" });
  }
  if (autoFillSourceSupportsNumericTrendOptions(sheet, sourceRange, "growth-trend")) {
    items.push({ mode: "growth-trend", label: "成長曲線(G)", accessKey: "G" });
  }
  return items;
}



function toggleAutoFillOptionsMenu(event) {
  event.preventDefault();
  event.stopPropagation();
  if (!state.autoFillOptions) return;
  state.autoFillOptions.menuOpen = !state.autoFillOptions.menuOpen;
  renderSheet();
  if (state.autoFillOptions.menuOpen) {
    const mode = state.autoFillOptions.mode;
    queueMicrotask(() => focusAutoFillOptionsMenuItem(mode));
    requestAnimationFrame(() => focusAutoFillOptionsMenuItem(mode));
    setTimeout(() => focusAutoFillOptionsMenuItem(mode), 0);
  }
}



function handleAutoFillOptionClick(event) {
  event.preventDefault();
  event.stopPropagation();
  const mode = event.currentTarget.dataset.autoFillOption;
  applyAutoFillOption(mode);
}



function handleAutoFillOptionsMenuKeydown(event) {
  const menu = event.currentTarget;
  handleAutoFillOptionsMenuKeyEvent(event, menu);
}



function handleGlobalAutoFillOptionsKeydown(event) {
  if (!state.autoFillOptions?.menuOpen) return;
  const menu = document.querySelector("#autoFillOptionsMenu");
  if (!menu) return;
  if (event.target?.closest?.("#autoFillOptionsMenu")) return;
  if (isTextEditingTarget(event.target) && !autoFillOptionsKeyCanLeaveTextInput(event)) return;
  handleAutoFillOptionsMenuKeyEvent(event, menu);
}



function autoFillOptionsKeyCanLeaveTextInput(event) {
  const key = String(event.key || "");
  if (event.altKey || event.ctrlKey || event.metaKey) return false;
  return key === "Escape" || key === "Enter" || key === " " || key.startsWith("Arrow") ||
    ["Home", "End", "C", "S", "F", "O", "D", "W", "M", "Y", "L", "G"].includes(key.toUpperCase());
}



function handleAutoFillOptionsMenuKeyEvent(event, menu) {
  const items = [...menu.querySelectorAll("[data-auto-fill-option]:not(:disabled)")];
  if (!items.length) return;
  const activeIndex = Math.max(0, items.indexOf(document.activeElement));
  if (isElementImeCompositionEvent(event)) return;
  if (event.key === "ArrowDown" || event.key === "ArrowRight") {
    consumeCommandMenuKey(event);
    items[(activeIndex + 1) % items.length]?.focus();
    return;
  }
  if (event.key === "ArrowUp" || event.key === "ArrowLeft") {
    consumeCommandMenuKey(event);
    items[(activeIndex - 1 + items.length) % items.length]?.focus();
    return;
  }
  if (event.key === "Home") {
    consumeCommandMenuKey(event);
    items[0]?.focus();
    return;
  }
  if (event.key === "End") {
    consumeCommandMenuKey(event);
    items[items.length - 1]?.focus();
    return;
  }
  if (event.key === "Enter" || event.key === " ") {
    consumeCommandMenuKey(event);
    document.activeElement?.click?.();
    return;
  }
  const accessKey = String(event.key || "").toUpperCase();
  const target = items.find((item) => item.dataset.accessKey === accessKey);
  if (target) {
    consumeCommandMenuKey(event);
    target.click();
  }
}



function focusAutoFillOptionsMenuItem(mode) {
  const menu = document.querySelector("#autoFillOptionsMenu");
  const targetMode = normalizeAutoFillMode(mode) || "series";
  const item = menu?.querySelector(`[data-auto-fill-option="${targetMode}"]`) ||
    menu?.querySelector("[data-auto-fill-option]");
  item?.focus?.({ preventScroll: true });
}



function closeAutoFillOptionsMenuOnOutsidePointer(event) {
  if (!state.autoFillOptions) return;
  if (event.target.closest?.("#autoFillOptionsMenu, #autoFillOptionsButton")) return;
  clearAutoFillOptions();
}



function closeAutoFillOptionsMenuOnEscape(event) {
  if (event.key !== "Escape" || !state.autoFillOptions) return;
  if (state.autoFillOptions.menuOpen) consumeCommandMenuKey(event);
  clearAutoFillOptions();
}



function closeAutoFillOptionsMenu() {
  if (!state.autoFillOptions?.menuOpen) return;
  state.autoFillOptions.menuOpen = false;
  renderSheet();
}



function removeAutoFillOptionsElements() {
  document.querySelector("#autoFillOptionsMenu")?.remove();
  document.querySelector("#autoFillOptionsButton")?.remove();
}



function clearAutoFillOptions(options = {}) {
  if (!state.autoFillOptions) {
    removeAutoFillOptionsElements();
    return false;
  }
  state.autoFillOptions = null;
  removeAutoFillOptionsElements();
  if (options.render) renderSheet();
  return true;
}



function clearAutoFillOptionsOnExternalFocus(event) {
  if (!state.autoFillOptions) return;
  if (event.target?.closest?.("#autoFillOptionsMenu, #autoFillOptionsButton")) return;
  if (isSelectionKeyboardSink(event.target)) return;
  if ($sheetHost[0]?.contains(event.target)) return;
  clearAutoFillOptions();
}



function autoFillOptionsMatchCurrentSelection(ranges = activeSelectionRanges(), imageObjects = selectedImageObjects()) {
  const options = state.autoFillOptions;
  if (!options || options.sheetIndex !== state.activeSheetIndex || imageObjects.length) return false;
  const visibleRanges = (ranges || []).filter(Boolean);
  return visibleRanges.length === 1 && rangesEqualLoose(visibleRanges[0], options.targetRange);
}



function syncAutoFillOptionsWithSelection(ranges = activeSelectionRanges(), imageObjects = selectedImageObjects()) {
  if (!state.autoFillOptions) return;
  if (autoFillOptionsMatchCurrentSelection(ranges, imageObjects)) return;
  clearAutoFillOptions();
}



function applyAutoFillOption(mode) {
  const options = state.autoFillOptions;
  const sheet = activeSheet();
  if (!options || options.sheetIndex !== state.activeSheetIndex || !sheet) return;
  mode = normalizeAutoFillMode(mode);
  if (!mode) return;
  if (AUTO_FILL_DATE_MODES.has(mode) && !autoFillSourceSupportsDateOptions(sheet, options.sourceRange)) return;
  if (AUTO_FILL_TREND_MODES.has(mode) && !autoFillSourceSupportsNumericTrendOptions(sheet, options.sourceRange, mode)) return;
  if (!ensureEditableRangeForProtection(sheet, options.fillRange, "オートフィル")) return;
  const mergeValidation = validateAutoFillMerges(sheet, options.sourceRange, options.fillRange, options.direction);
  if (!mergeValidation.ok) {
    showMessageBox(mergeValidation.message, { title: "Cht WebSheet" });
    setStatus(mergeValidation.message);
    return;
  }
  const before = captureCellsForHistory(sheet, options.targetRange);
  const mergesBefore = cloneMerges(sheet.merges);
  const beforeMap = autoFillHistoryCellMap(options.before);
  ensureSheetSize(sheet, options.targetRange.bottom, options.targetRange.right);
  removeIntersectingMerges(sheet, options.fillRange);
  writeAutoFillCells(sheet, options.sourceRange, options.fillRange, options.direction, {
    mode,
    beforeMap,
    explicitMode: true,
  });
  applyAutoFillMerges(sheet, options.sourceRange, options.fillRange, options.direction);
  state.hf = buildFormulaEngine(state.model);
  const after = captureCellsForHistory(sheet, options.targetRange);
  const mergesAfter = cloneMerges(sheet.merges);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    mergesBefore,
    mergesAfter,
    selectionBefore: options.targetRange,
    selectionAfter: options.targetRange,
    label: `${rangeToLabel(options.targetRange)} オートフィル オプション`,
  });
  state.selectionRange = { ...options.targetRange };
  state.selectionRanges = [{ ...options.targetRange }];
  state.selected = { sheetIndex: state.activeSheetIndex, row: options.targetRange.top, col: options.targetRange.left };
  state.selectionStart = state.selected;
  state.autoFillOptions = { ...options, mode, menuOpen: false };
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  scheduleWorkbookPersistence();
  setStatus(`${rangeToLabel(options.targetRange)} に ${autoFillOptionStatusLabel(mode)} を適用しました。`);
}



function autoFillOptionStatusLabel(mode) {
  if (mode === "copy") return "セルのコピー";
  if (mode === "formatting") return "書式のみコピー";
  if (mode === "no-formatting") return "書式なしコピー";
  if (mode === "date-day") return "日単位";
  if (mode === "date-weekday") return "週日単位";
  if (mode === "date-month") return "月単位";
  if (mode === "date-year") return "年単位";
  if (mode === "linear-trend") return "線形近似";
  if (mode === "growth-trend") return "成長曲線";
  return "連続データ";
}



function autoFillHistoryCellMap(cells) {
  return new Map((cells || []).map((item) => [cellKey(item.row, item.col), cloneCellModel(item.cell)]));
}



function normalizeAutoFillMode(mode) {
  const value = String(mode || "").trim();
  return AUTO_FILL_MODES.has(value) ? value : "";
}



function defaultAutoFillMode(sheet, sourceRange) {
  if (autoFillSourceIsPlainAlphabet(sheet, sourceRange)) return "copy";
  return isSinglePlainNumberAutoFillSource(sheet, sourceRange) ? "copy" : "series";
}



function autoFillSourceIsPlainAlphabet(sheet, sourceRange) {
  if (!sheet || !sourceRange) return false;
  const cellCount = (sourceRange.bottom - sourceRange.top + 1) * (sourceRange.right - sourceRange.left + 1);
  const entries = autoFillPlainAlphabetEntries(sheet, sourceRange);
  return entries.length > 0 && entries.length === cellCount;
}



function autoFillPlainAlphabetEntries(sheet, sourceRange) {
  const entries = [];
  let invalid = false;
  forEachCellInRange(sourceRange, (row, col) => {
    if (invalid) return;
    const entry = cellAlphabetEntryForAutoFill(sheet, row, col);
    if (!entry) {
      invalid = true;
      return;
    }
    entries.push(entry);
  });
  return invalid ? [] : entries;
}



function autoFillModeForPointerEvent(sheet, sourceRange, event, fallbackMode = "") {
  const baseMode = defaultAutoFillMode(sheet, sourceRange);
  const fallback = normalizeAutoFillMode(fallbackMode);
  if (!event?.ctrlKey && !event?.metaKey && !state.autoFillModifierActive) return fallback || baseMode;
  if (baseMode === "copy") return "series";
  if (baseMode === "series") return "copy";
  return fallback || baseMode;
}



function isSinglePlainNumberAutoFillSource(sheet, sourceRange) {
  if (!sheet || !sourceRange || sourceRange.top !== sourceRange.bottom || sourceRange.left !== sourceRange.right) return false;
  const cell = sheet.cells[cellKey(sourceRange.top, sourceRange.left)];
  if (!cell || cell.formula) return false;
  if (isDateFormat(primaryExcelFormatSection(cell.numFmt || ""))) return false;
  if (cellListEntryForAutoFill(sheet, sourceRange.top, sourceRange.left)) return false;
  if (cellTextNumberEntryForAutoFill(sheet, sourceRange.top, sourceRange.left)) return false;
  return cellNumericValueForAutoFill(sheet, sourceRange.top, sourceRange.left) != null;
}



function autoFillSourceSupportsDateOptions(sheet, sourceRange) {
  if (!sheet || !sourceRange) return false;
  let sawDate = false;
  let invalid = false;
  forEachCellInRange(sourceRange, (row, col) => {
    if (invalid) return;
    const cell = sheet.cells[cellKey(row, col)];
    if (!cell || cell.formula || !isOccupiedCell(cell)) {
      invalid = true;
      return;
    }
    const value = cellDateSeriesValueForAutoFill(sheet, row, col);
    if (!value || !autoFillDateValueSupportsCalendarOptions(value)) {
      invalid = true;
      return;
    }
    sawDate = true;
  });
  return sawDate && !invalid;
}



function autoFillSourceSupportsNumericTrendOptions(sheet, sourceRange, mode = "linear-trend") {
  if (!sheet || !sourceRange) return false;
  const axisCount = sourceRange.bottom - sourceRange.top + 1 >= 2 || sourceRange.right - sourceRange.left + 1 >= 2;
  if (!axisCount) return mode === "linear-trend" && autoFillPlainNumericEntries(sheet, sourceRange).length === 1;
  const values = autoFillPlainNumericEntries(sheet, sourceRange);
  if (!values.length) return false;
  if (values.length !== (sourceRange.bottom - sourceRange.top + 1) * (sourceRange.right - sourceRange.left + 1)) return false;
  if (mode === "growth-trend") return values.length >= 2 && values.every((value) => value > 0);
  return true;
}



function autoFillPlainNumericEntries(sheet, sourceRange) {
  const values = [];
  let invalid = false;
  forEachCellInRange(sourceRange, (row, col) => {
    if (invalid) return;
    const cell = sheet.cells[cellKey(row, col)];
    if (!cell || cell.formula || !isOccupiedCell(cell)) {
      invalid = true;
      return;
    }
    if (isDateFormat(primaryExcelFormatSection(cell.numFmt || ""))) {
      invalid = true;
      return;
    }
    if (cellListEntryForAutoFill(sheet, row, col) || cellTextNumberEntryForAutoFill(sheet, row, col)) {
      invalid = true;
      return;
    }
    const value = cellNumericValueForAutoFill(sheet, row, col);
    if (value == null) {
      invalid = true;
      return;
    }
    values.push(value);
  });
  return invalid ? [] : values;
}



function autoFillDateValueSupportsCalendarOptions(value) {
  if (!value) return false;
  if (value.pattern?.kind === "time") return false;
  if (value.kind === "text") return value.pattern?.kind === "date";
  return autoFillNumberFormatLooksCalendarDate(value.numFmt || "");
}



function autoFillNumberFormatLooksCalendarDate(numFmt = "") {
  const format = dateDetectionFormat(primaryExcelFormatSection(numFmt || ""));
  if (/[年月日]/.test(format)) return true;
  const hasYear = /(^|[^a-z])y{2,4}/i.test(format);
  const hasDay = /(^|[^a-z])d{1,4}/i.test(format);
  const hasMonth = /(^|[^a-z])m{1,4}/i.test(format);
  return (hasYear && hasMonth) || (hasMonth && hasDay) || (hasYear && hasDay);
}



function autoFillDoubleClickRange(sheet, sourceRange) {
  if (!sheet || !sourceRange || sourceRange.fullRow || sourceRange.fullColumn) return null;
  const leftEnd = autoFillAdjacentDataEndRow(sheet, sourceRange, sourceRange.left - 1);
  const rightEnd = autoFillAdjacentDataEndRow(sheet, sourceRange, sourceRange.right + 1);
  const bottom = Math.max(leftEnd, rightEnd);
  if (bottom <= sourceRange.bottom) return null;
  const targetRange = { ...sourceRange, bottom };
  const fillRange = { ...targetRange, top: sourceRange.bottom + 1 };
  return { sourceRange: { ...sourceRange }, targetRange, fillRange, direction: "down" };
}



function autoFillAdjacentDataEndRow(sheet, sourceRange, col) {
  if (!sheet || !sourceRange || col < 1) return 0;
  const maxRow = Math.max(Number(sheet.rowCount) || 1, minimumSheetExtent(sheet).rows || 1);
  const sourceHasAdjacentData = integerRange(sourceRange.top, sourceRange.bottom).some((row) => isOccupiedCell(sheet.cells[cellKey(row, col)]));
  if (!sourceHasAdjacentData) return 0;
  let row = sourceRange.top;
  let last = 0;
  while (row <= maxRow) {
    if (!isOccupiedCell(sheet.cells[cellKey(row, col)])) break;
    last = row;
    row += 1;
  }
  return last;
}



function integerRange(start, end) {
  const values = [];
  for (let value = start; value <= end; value += 1) values.push(value);
  return values;
}



function writeAutoFillCells(sheet, sourceRange, fillRange, direction, options = {}) {
  const mode = options.mode || "series";
  forEachCellInRange(fillRange, (row, col) => {
    const sourcePoint = autoFillSourcePoint(sourceRange, row, col, direction);
    const sourceCell = sourceCellForAutoFill(sheet, sourcePoint.row, sourcePoint.col);
    const key = cellKey(row, col);
    const beforeCell = options.beforeMap?.get(key) || null;
    if (autoFillShouldOwnTargetFormatting(mode, beforeCell)) removeSheetStyleRangesAtCell(sheet, row, col);
    if (!sourceCell) {
      if (mode === "no-formatting" && beforeCell) storeBlankAutoFillCellWithFormatting(sheet, key, row, col, beforeCell);
      else delete sheet.cells[key];
      return;
    }

    const nextCell = cloneCellForAutoFill(sheet, sourceRange, sourceCell, sourcePoint, row, col, direction, {
      mode,
      beforeCell,
      explicitMode: options.explicitMode === true,
    });
    nextCell.row = row;
    nextCell.col = col;
    storeCellModelForSheet(sheet, key, nextCell);
  });
}



function sourceCellForAutoFill(sheet, row, col) {
  const storedCell = sheet.cells[cellKey(row, col)] || null;
  return sheetDisplayCell(sheet, row, col, storedCell) || null;
}



function autoFillShouldOwnTargetFormatting(mode, beforeCell) {
  return mode !== "no-formatting" || Boolean(beforeCell);
}



function storeBlankAutoFillCellWithFormatting(sheet, key, row, col, formatCell) {
  const nextCell = cloneCellModel(formatCell) || { row, col, raw: null, formula: null, cached: null, display: "", css: {}, borders: {} };
  nextCell.row = row;
  nextCell.col = col;
  nextCell.raw = null;
  nextCell.formula = null;
  nextCell.cached = null;
  nextCell.display = "";
  delete nextCell.richText;
  storeCellModelForSheet(sheet, key, nextCell);
}



function cloneCellForAutoFill(sheet, sourceRange, sourceCell, sourcePoint, targetRow, targetCol, direction, options = {}) {
  if (options.mode === "formatting") return cloneCellWithAutoFillFormattingOnly(sourceCell, options.beforeCell, targetRow, targetCol);
  const nextCell = cloneCellModel(sourceCell);
  const rowOffset = targetRow - sourcePoint.row;
  const colOffset = targetCol - sourcePoint.col;
  if (options.mode === "no-formatting") {
    applyAutoFillCellFormatting(nextCell, options.beforeCell);
  }
  if (typeof nextCell.formula === "string" && nextCell.formula.startsWith("=")) {
    nextCell.formula = shiftFormulaReferences(nextCell.formula, rowOffset, colOffset);
    nextCell.raw = nextCell.formula;
    nextCell.cached = null;
    nextCell.display = nextCell.formula;
    return nextCell;
  }

  const seriesValue = options.mode === "copy" ? null : autoFillSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction, {
    mode: options.mode,
    explicitMode: options.explicitMode === true,
  });
  if (seriesValue) {
    nextCell.raw = seriesValue.raw;
    nextCell.cached = seriesValue.cached ?? seriesValue.raw;
    nextCell.display = seriesValue.display ?? "";
    if (seriesValue.numFmt && !nextCell.numFmt) nextCell.numFmt = seriesValue.numFmt;
  }
  if (options.mode === "no-formatting") {
    applyAutoFillCellFormatting(nextCell, options.beforeCell);
  }
  return nextCell;
}



function cloneCellWithAutoFillFormattingOnly(sourceCell, beforeCell, targetRow, targetCol) {
  const nextCell = cloneCellModel(beforeCell) || { row: targetRow, col: targetCol, raw: null, cached: null, display: "", css: {}, borders: {} };
  applyAutoFillCellFormatting(nextCell, sourceCell);
  return nextCell;
}



function applyAutoFillCellFormatting(targetCell, formatCell) {
  targetCell.css = cloneJsonLikeValue(formatCell?.css) || {};
  targetCell.hasFill = Boolean(formatCell?.hasFill);
  targetCell.hasFontColor = Boolean(formatCell?.hasFontColor);
  targetCell.borders = cloneJsonLikeValue(formatCell?.borders) || {};
  targetCell.numFmt = formatCell?.numFmt || "";
}



function autoFillSourcePoint(sourceRange, row, col, direction) {
  const height = sourceRange.bottom - sourceRange.top + 1;
  const width = sourceRange.right - sourceRange.left + 1;
  const rowIndex = direction === "up" || direction === "down" ? positiveModulo(row - sourceRange.top, height) : row - sourceRange.top;
  const colIndex = direction === "left" || direction === "right" ? positiveModulo(col - sourceRange.left, width) : col - sourceRange.left;
  return {
    row: sourceRange.top + Math.max(0, Math.min(height - 1, rowIndex)),
    col: sourceRange.left + Math.max(0, Math.min(width - 1, colIndex)),
  };
}



function autoFillSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction, options = {}) {
  if (AUTO_FILL_DATE_MODES.has(options.mode)) {
    return autoFillDateSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction, options);
  }
  if (AUTO_FILL_TREND_MODES.has(options.mode)) {
    return autoFillNumericTrendSeriesEntry(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction, options.mode);
  }
  return autoFillDateSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction, options) ||
    autoFillListSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction) ||
    autoFillAlphabetSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction, options) ||
    autoFillTextNumberSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction) ||
    autoFillNumericSeriesEntry(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction);
}



function autoFillAxisContext(sourceRange, sourcePoint, targetRow, targetCol, direction) {
  const vertical = direction === "up" || direction === "down";
  const length = vertical ? sourceRange.bottom - sourceRange.top + 1 : sourceRange.right - sourceRange.left + 1;
  const sourceIndex = vertical ? sourcePoint.row - sourceRange.top : sourcePoint.col - sourceRange.left;
  const targetIndex = vertical ? targetRow - sourceRange.top : targetCol - sourceRange.left;
  return { vertical, length, sourceIndex, targetIndex };
}



function autoFillAxisCellPoint(sourceRange, sourcePoint, axis, index) {
  return {
    row: axis.vertical ? sourceRange.top + index : sourcePoint.row,
    col: axis.vertical ? sourcePoint.col : sourceRange.left + index,
  };
}



function autoFillNumericSeriesEntry(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction) {
  const value = autoFillNumericSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction);
  return value == null ? null : { raw: value, cached: value, display: "" };
}



function autoFillNumericTrendSeriesEntry(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction, mode) {
  const value = autoFillNumericTrendSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction, mode);
  return value == null ? null : { raw: value, cached: value, display: "" };
}



function autoFillNumericTrendSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction, mode) {
  const axis = autoFillAxisContext(sourceRange, sourcePoint, targetRow, targetCol, direction);
  const { length, sourceIndex, targetIndex } = axis;
  const values = [];
  for (let index = 0; index < length; index += 1) {
    const { row, col } = autoFillAxisCellPoint(sourceRange, sourcePoint, axis, index);
    const cell = sheet.cells[cellKey(row, col)];
    if (!cell || isDateFormat(primaryExcelFormatSection(cell.numFmt || ""))) return null;
    const value = cellNumericValueForAutoFill(sheet, row, col);
    if (value == null) return null;
    values.push(value);
  }
  const source = values[sourceIndex];
  const delta = targetIndex - sourceIndex;
  if (mode === "growth-trend") {
    if (length < 2 || values.some((value) => value <= 0)) return null;
    const ratio = Math.pow(values[length - 1] / values[0], 1 / (length - 1));
    if (!Number.isFinite(ratio) || ratio <= 0 || ratio === 1) return null;
    return normalizeAutoFillTrendNumber(source * Math.pow(ratio, delta));
  }
  if (length < 2) {
    return source == null ? null : normalizeAutoFillTrendNumber(source + delta);
  }
  const step = (values[length - 1] - values[0]) / (length - 1);
  if (!Number.isFinite(step) || step === 0) return null;
  return normalizeAutoFillTrendNumber(source + step * delta);
}



function normalizeAutoFillTrendNumber(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return null;
  const rounded = Math.round(number);
  if (Math.abs(number - rounded) < 1e-10) return rounded;
  const fixed = Number(number.toPrecision(12));
  return Object.is(fixed, -0) ? 0 : fixed;
}



function autoFillNumericSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction) {
  const axis = autoFillAxisContext(sourceRange, sourcePoint, targetRow, targetCol, direction);
  const { length, sourceIndex, targetIndex } = axis;
  if (length < 2) {
    const value = cellNumericValueForAutoFill(sheet, sourcePoint.row, sourcePoint.col);
    return value == null ? null : value + (targetIndex - sourceIndex);
  }
  const values = [];
  for (let index = 0; index < length; index += 1) {
    const { row, col } = autoFillAxisCellPoint(sourceRange, sourcePoint, axis, index);
    const value = cellNumericValueForAutoFill(sheet, row, col);
    if (value == null) return null;
    values.push(value);
  }

  const step = (values[length - 1] - values[0]) / (length - 1);
  if (!Number.isFinite(step) || step === 0) return null;
  return values[sourceIndex] + step * (targetIndex - sourceIndex);
}



function autoFillDateSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction, options = {}) {
  const axis = autoFillAxisContext(sourceRange, sourcePoint, targetRow, targetCol, direction);
  const { length, sourceIndex, targetIndex } = axis;
  const values = [];
  for (let index = 0; index < length; index += 1) {
    const { row, col } = autoFillAxisCellPoint(sourceRange, sourcePoint, axis, index);
    const value = cellDateSeriesValueForAutoFill(sheet, row, col);
    if (!value) return null;
    values.push(value);
  }
  const source = values[sourceIndex] || values[0];
  const delta = targetIndex - sourceIndex;
  let nextSerial = null;
  if (options.mode === "date-day") {
    nextSerial = source.serial + delta;
  } else if (options.mode === "date-weekday") {
    nextSerial = addWeekdaysToExcelSerial(source.serial, delta);
  } else if (options.mode === "date-month") {
    nextSerial = addCalendarMonthsToExcelSerial(source.serial, delta);
  } else if (options.mode === "date-year") {
    nextSerial = addCalendarMonthsToExcelSerial(source.serial, delta * 12);
  } else {
    const step = length >= 2
      ? (values[length - 1].serial - values[0].serial) / (length - 1)
      : (values[0]?.pattern?.kind === "time" ? 1 / 24 : 1);
    if (!Number.isFinite(step) || step === 0) return null;
    nextSerial = source.serial + step * delta;
  }
  if (!Number.isFinite(nextSerial)) return null;
  return source.kind === "text"
    ? { raw: formatAutoFillDateText(nextSerial, source.pattern), cached: formatAutoFillDateText(nextSerial, source.pattern), display: formatAutoFillDateText(nextSerial, source.pattern) }
    : { raw: nextSerial, cached: nextSerial, display: "", numFmt: source.numFmt || "yyyy/mm/dd" };
}



function addWeekdaysToExcelSerial(serial, delta) {
  const amount = Math.trunc(Number(delta) || 0);
  if (!amount) return serial;
  let next = Number(serial);
  const step = amount > 0 ? 1 : -1;
  let remaining = Math.abs(amount);
  while (remaining > 0) {
    next += step;
    const day = excelSerialToDate(next).getDay();
    if (day !== 0 && day !== 6) remaining -= 1;
  }
  return next;
}



function addCalendarMonthsToExcelSerial(serial, monthDelta) {
  const value = Number(serial);
  const amount = Math.trunc(Number(monthDelta) || 0);
  if (!Number.isFinite(value) || !amount) return value;
  const date = excelSerialToDate(value);
  const fraction = value - Math.floor(value);
  const year = date.getFullYear();
  const month = date.getMonth();
  const day = date.getDate();
  const targetMonthIndex = month + amount;
  const targetYear = year + Math.floor(targetMonthIndex / 12);
  const normalizedMonth = positiveModulo(targetMonthIndex, 12);
  const targetDay = Math.min(day, daysInMonthUtc(targetYear, normalizedMonth));
  return dateToExcelSerial(new Date(Date.UTC(targetYear, normalizedMonth, targetDay))) + fraction;
}



function daysInMonthUtc(year, zeroBasedMonth) {
  return new Date(Date.UTC(year, zeroBasedMonth + 1, 0)).getUTCDate();
}



function cellDateSeriesValueForAutoFill(sheet, row, col) {
  const cell = sheet.cells[cellKey(row, col)];
  if (!cell || cell.formula) return null;
  const raw = cell.raw ?? cell.cached ?? cell.display;
  const numFmt = cell.numFmt || "";
  if (isDateFormat(primaryExcelFormatSection(numFmt))) {
    const serial = dateSerialFromPersistedValue(raw, numFmt);
    if (Number.isFinite(serial)) return { kind: "number", serial, numFmt };
  }
  const parsed = parseAutoFillDateText(raw);
  return parsed ? { kind: "text", ...parsed } : null;
}



function parseAutoFillDateText(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return { serial: dateToExcelSerial(value), pattern: "yyyy/mm/dd" };
  const text = String(value ?? "").trim();
  let match = text.match(/^(\d{4})([-/])(\d{1,2})\2(\d{1,2})$/);
  if (match) {
    const [, year, separator, month, day] = match;
    const serial = autoFillDateSerialFromParts(year, month, day);
    if (Number.isFinite(serial)) {
      return {
        serial,
        pattern: {
          kind: "date",
          order: "ymd",
          separator,
          yearWidth: year.length,
          monthWidth: month.length,
          dayWidth: day.length,
        },
      };
    }
  }
  match = text.match(/^(\d{1,2})([-/])(\d{1,2})\2(\d{4})$/);
  if (match) {
    const [, month, separator, day, year] = match;
    const serial = autoFillDateSerialFromParts(year, month, day);
    if (Number.isFinite(serial)) {
      return {
        serial,
        pattern: {
          kind: "date",
          order: "mdy",
          separator,
          yearWidth: year.length,
          monthWidth: month.length,
          dayWidth: day.length,
        },
      };
    }
  }
  match = text.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
  if (match) {
    const [, year, month, day] = match;
    const serial = autoFillDateSerialFromParts(year, month, day);
    if (Number.isFinite(serial)) {
      return {
        serial,
        pattern: {
          kind: "date",
          order: "ymd-ja",
          yearWidth: year.length,
          monthWidth: month.length,
          dayWidth: day.length,
        },
      };
    }
  }
  const time = parseAutoFillTimeText(text);
  if (time) return time;
  return null;
}



function parseAutoFillTimeText(text) {
  let match = String(text ?? "").trim().match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (match) {
    const [, hour, minute, second = ""] = match;
    const serial = autoFillTimeSerialFromParts(hour, minute, second || "0");
    if (Number.isFinite(serial)) {
      return {
        serial,
        pattern: {
          kind: "time",
          style: "colon",
          hourWidth: hour.length,
          minuteWidth: minute.length,
          secondWidth: second.length,
        },
      };
    }
  }
  match = String(text ?? "").trim().match(/^(\d{1,2})時(?:(\d{1,2})分)?(?:(\d{1,2})秒)?$/);
  if (match) {
    const [, hour, minuteRaw, secondRaw] = match;
    const serial = autoFillTimeSerialFromParts(hour, minuteRaw || "0", secondRaw || "0");
    if (Number.isFinite(serial)) {
      return {
        serial,
        pattern: {
          kind: "time",
          style: "ja",
          hourWidth: hour.length,
          minuteWidth: (minuteRaw || "").length,
          secondWidth: (secondRaw || "").length,
          hasMinute: Boolean(minuteRaw),
          hasSecond: Boolean(secondRaw),
        },
      };
    }
  }
  return null;
}



function autoFillDateSerialFromParts(yearText, monthText, dayText) {
  const year = Number(yearText);
  const month = Number(monthText);
  const day = Number(dayText);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    !Number.isFinite(year) ||
    !Number.isFinite(month) ||
    !Number.isFinite(day) ||
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() + 1 !== month ||
    date.getUTCDate() !== day
  ) {
    return null;
  }
  return dateToExcelSerial(date);
}



function autoFillTimeSerialFromParts(hourText, minuteText, secondText) {
  const hour = Number(hourText);
  const minute = Number(minuteText);
  const second = Number(secondText);
  if (
    !Number.isFinite(hour) ||
    !Number.isFinite(minute) ||
    !Number.isFinite(second) ||
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59 ||
    second < 0 ||
    second > 59
  ) {
    return null;
  }
  return (hour * 3600 + minute * 60 + second) / 86400;
}



function formatAutoFillDateText(serial, pattern = "yyyy/mm/dd") {
  if (pattern?.kind === "time") return formatAutoFillTimeText(serial, pattern);
  const date = excelSerialToDate(serial);
  const year = String(date.getUTCFullYear());
  const monthNumber = String(date.getUTCMonth() + 1);
  const dayNumber = String(date.getUTCDate());
  if (pattern && typeof pattern === "object") {
    if (pattern.order === "ymd-ja") {
      const yearPart = pattern.yearWidth > 4 ? year.padStart(pattern.yearWidth, "0") : year;
      const monthPart = pattern.monthWidth > 1 ? monthNumber.padStart(pattern.monthWidth, "0") : monthNumber;
      const dayPart = pattern.dayWidth > 1 ? dayNumber.padStart(pattern.dayWidth, "0") : dayNumber;
      return `${yearPart}年${monthPart}月${dayPart}日`;
    }
    const separator = pattern.separator || "/";
    const yearPart = pattern.yearWidth > 4 ? year.padStart(pattern.yearWidth, "0") : year;
    const monthPart = pattern.monthWidth > 1 ? monthNumber.padStart(pattern.monthWidth, "0") : monthNumber;
    const dayPart = pattern.dayWidth > 1 ? dayNumber.padStart(pattern.dayWidth, "0") : dayNumber;
    return pattern.order === "mdy"
      ? [monthPart, dayPart, yearPart].join(separator)
      : [yearPart, monthPart, dayPart].join(separator);
  }
  const month = monthNumber.padStart(2, "0");
  const day = dayNumber.padStart(2, "0");
  return String(pattern)
    .replace(/yyyy/g, year)
    .replace(/mm/g, month)
    .replace(/dd/g, day);
}



function formatAutoFillTimeText(serial, pattern) {
  const normalized = positiveModulo(Number(serial) || 0, 1);
  let totalSeconds = Math.round(normalized * 86400) % 86400;
  const hour = Math.floor(totalSeconds / 3600);
  totalSeconds -= hour * 3600;
  const minute = Math.floor(totalSeconds / 60);
  const second = totalSeconds - minute * 60;
  const hourPart = String(hour).padStart(pattern.hourWidth > 1 ? pattern.hourWidth : 1, "0");
  const minutePart = String(minute).padStart(pattern.minuteWidth > 1 ? pattern.minuteWidth : 2, "0");
  const secondPart = String(second).padStart(pattern.secondWidth > 1 ? pattern.secondWidth : 2, "0");
  if (pattern.style === "ja") {
    if (pattern.hasSecond) return `${hourPart}時${minutePart}分${secondPart}秒`;
    if (pattern.hasMinute) return `${hourPart}時${minutePart}分`;
    return `${hourPart}時`;
  }
  return pattern.secondWidth ? `${hourPart}:${minutePart}:${secondPart}` : `${hourPart}:${minutePart}`;
}



function autoFillListSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction) {
  const axis = autoFillAxisContext(sourceRange, sourcePoint, targetRow, targetCol, direction);
  const entries = [];
  for (let index = 0; index < axis.length; index += 1) {
    const { row, col } = autoFillAxisCellPoint(sourceRange, sourcePoint, axis, index);
    const entry = cellListEntryForAutoFill(sheet, row, col);
    if (!entry) return null;
    entries.push(entry);
  }
  const list = entries[0].list;
  if (!entries.every((entry) => entry.list === list)) return null;
  const step = axis.length >= 2 ? circularListStep(entries[0].index, entries[axis.length - 1].index, list.length, axis.length - 1) : 1;
  if (!Number.isFinite(step) || step === 0 || !Number.isInteger(step)) return null;
  const source = entries[axis.sourceIndex] || entries[0];
  const nextIndex = positiveModulo(source.index + step * (axis.targetIndex - axis.sourceIndex), list.length);
  const text = list[nextIndex];
  return { raw: text, cached: text, display: text };
}



function cellListEntryForAutoFill(sheet, row, col) {
  const cell = sheet.cells[cellKey(row, col)];
  if (!cell || cell.formula) return null;
  const text = String(cell.raw ?? cell.cached ?? cell.display ?? "").trim();
  if (!text) return null;
  for (const list of AUTO_FILL_LISTS) {
    const index = list.findIndex((item) => item.toLocaleLowerCase("ja-JP") === text.toLocaleLowerCase("ja-JP"));
    if (index >= 0) return { list, index };
  }
  return null;
}



function circularListStep(firstIndex, lastIndex, length, intervals) {
  if (!intervals) return 1;
  let forward = lastIndex - firstIndex;
  if (forward < 0) forward += length;
  if (forward === 0) return 0;
  return forward / intervals;
}



function autoFillAlphabetSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction, options = {}) {
  if (options.mode !== "series" || options.explicitMode !== true) return null;
  const axis = autoFillAxisContext(sourceRange, sourcePoint, targetRow, targetCol, direction);
  const entries = [];
  for (let index = 0; index < axis.length; index += 1) {
    const { row, col } = autoFillAxisCellPoint(sourceRange, sourcePoint, axis, index);
    const entry = cellAlphabetEntryForAutoFill(sheet, row, col);
    if (!entry) return null;
    entries.push(entry);
  }
  const first = entries[0];
  if (!entries.every((entry) => entry.letterCase === first.letterCase)) return null;
  const step = axis.length >= 2 ? (entries[axis.length - 1].value - first.value) / (axis.length - 1) : 1;
  if (!Number.isFinite(step) || step === 0 || !Number.isInteger(step)) return null;
  const source = entries[axis.sourceIndex] || first;
  const nextValue = source.value + step * (axis.targetIndex - axis.sourceIndex);
  if (!Number.isInteger(nextValue) || nextValue < 1) return null;
  const text = first.letterCase === "lower" ? columnName(nextValue).toLowerCase() : columnName(nextValue);
  return { raw: text, cached: text, display: text };
}



function cellAlphabetEntryForAutoFill(sheet, row, col) {
  const cell = sheet.cells[cellKey(row, col)];
  if (!cell || cell.formula) return null;
  const text = String(cell.raw ?? cell.cached ?? cell.display ?? "").trim();
  if (!/^[A-Za-z]+$/.test(text)) return null;
  if (cellListEntryForAutoFill(sheet, row, col)) return null;
  const letterCase = text === text.toUpperCase() ? "upper" : (text === text.toLowerCase() ? "lower" : "");
  if (!letterCase) return null;
  const value = columnNumber(text);
  return Number.isInteger(value) && value > 0 ? { value, letterCase } : null;
}



function autoFillTextNumberSeriesValue(sheet, sourceRange, sourcePoint, targetRow, targetCol, direction) {
  const axis = autoFillAxisContext(sourceRange, sourcePoint, targetRow, targetCol, direction);
  const entries = [];
  for (let index = 0; index < axis.length; index += 1) {
    const { row, col } = autoFillAxisCellPoint(sourceRange, sourcePoint, axis, index);
    const entry = cellTextNumberEntryForAutoFill(sheet, row, col);
    if (!entry) return null;
    entries.push(entry);
  }
  const first = entries[0];
  if (!entries.every((entry) => entry.prefix === first.prefix && entry.suffix === first.suffix)) return null;
  const step = axis.length >= 2 ? (entries[axis.length - 1].number - first.number) / (axis.length - 1) : 1;
  if (!Number.isFinite(step) || step === 0) return null;
  const source = entries[axis.sourceIndex] || first;
  const nextNumber = source.number + step * (axis.targetIndex - axis.sourceIndex);
  if (!Number.isInteger(nextNumber)) return null;
  const width = Math.max(source.width, String(Math.abs(nextNumber)).length);
  const sign = nextNumber < 0 ? "-" : "";
  const digits = String(Math.abs(nextNumber)).padStart(width, "0");
  const text = `${source.prefix}${sign}${digits}${source.suffix}`;
  return { raw: text, cached: text, display: text };
}



function cellTextNumberEntryForAutoFill(sheet, row, col) {
  const cell = sheet.cells[cellKey(row, col)];
  if (!cell || cell.formula) return null;
  const text = String(cell.raw ?? cell.cached ?? cell.display ?? "");
  const match = text.match(/^(.*?)(\d+)(\D*)$/);
  if (!match) return null;
  if (!match[1] && !match[3] && !/^0\d/.test(match[2])) return null;
  return {
    prefix: match[1],
    number: Number(match[2]),
    width: match[2].length,
    suffix: match[3],
  };
}



function cellNumericValueForAutoFill(sheet, row, col) {
  const cell = sheet.cells[cellKey(row, col)];
  if (!cell || cell.formula) return null;
  return parseAutoFillNumber(cell.raw ?? cell.cached ?? cell.display);
}



function parseAutoFillNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string") return null;
  const normalized = value.trim().replace(/,/g, "");
  if (!normalized || !/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) return null;
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}



function shiftFormulaReferences(formula, rowOffset, colOffset) {
  const text = String(formula);
  const referencePattern = /(^|[^A-Za-z0-9_.])((?:(?:'(?:(?:[^']|'')*)'|[A-Za-z0-9_]+)!)?(?:\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+))/gi;
  return text.replace(referencePattern, (match, leading, reference, offset) => {
    const referenceStart = offset + String(leading || "").length;
    if (isFormulaIndexInsideStringLiteral(text, referenceStart)) return match;
    return `${leading || ""}${shiftFormulaReferenceForAutoFill(reference, rowOffset, colOffset)}`;
  });
}



function shiftFormulaReferenceForAutoFill(reference, rowOffset, colOffset) {
  const sheetMatch = String(reference || "").match(/^((?:'(?:(?:[^']|'')*)'|[A-Za-z0-9_]+)!)/);
  const sheetPrefix = sheetMatch?.[1] || "";
  const bareReference = sheetPrefix ? String(reference).slice(sheetPrefix.length) : String(reference || "");
  const shifted = shiftBareFormulaReferenceForAutoFill(bareReference, rowOffset, colOffset);
  return `${sheetPrefix}${shifted}`;
}



function shiftBareFormulaReferenceForAutoFill(reference, rowOffset, colOffset) {
  const columnRangeMatch = String(reference || "").match(/^(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})$/i);
  if (columnRangeMatch) {
    const [, startAbs, startText, endAbs, endText] = columnRangeMatch;
    const startCol = startAbs ? columnNumber(startText) : Math.max(1, columnNumber(startText) + colOffset);
    const endCol = endAbs ? columnNumber(endText) : Math.max(1, columnNumber(endText) + colOffset);
    return `${startAbs || ""}${columnName(startCol)}:${endAbs || ""}${columnName(endCol)}`;
  }

  const rowRangeMatch = String(reference || "").match(/^(\$?)(\d+):(\$?)(\d+)$/);
  if (rowRangeMatch) {
    const [, startAbs, startText, endAbs, endText] = rowRangeMatch;
    const startRow = startAbs ? Number(startText) : Math.max(1, Number(startText) + rowOffset);
    const endRow = endAbs ? Number(endText) : Math.max(1, Number(endText) + rowOffset);
    return `${startAbs || ""}${startRow}:${endAbs || ""}${endRow}`;
  }

  const match = String(reference || "").match(/^(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?$/i);
  if (!match) return reference;
  const [, colAbs, colName, rowAbs, rowNumber, endColAbs, endColName, endRowAbs, endRowNumber] = match;
  const startCol = colAbs ? columnNumber(colName) : Math.max(1, columnNumber(colName) + colOffset);
  const startRow = rowAbs ? Number(rowNumber) : Math.max(1, Number(rowNumber) + rowOffset);
  if (!endColName) return `${colAbs || ""}${columnName(startCol)}${rowAbs || ""}${startRow}`;
  const endCol = endColAbs ? columnNumber(endColName) : Math.max(1, columnNumber(endColName) + colOffset);
  const endRow = endRowAbs ? Number(endRowNumber) : Math.max(1, Number(endRowNumber) + rowOffset);
  return `${colAbs || ""}${columnName(startCol)}${rowAbs || ""}${startRow}:${endColAbs || ""}${columnName(endCol)}${endRowAbs || ""}${endRow}`;
}



function structureReferenceShiftForOperation(kind, option, range, sheetIndex = state.activeSheetIndex) {
  if (!range || (kind !== "insert" && kind !== "delete")) return null;
  if (option !== "entire-row" && option !== "entire-column") return null;
  const axis = option === "entire-row" ? "row" : "col";
  const start = axis === "row" ? range.top : range.left;
  const amount = axis === "row" ? range.bottom - range.top + 1 : range.right - range.left + 1;
  if (!Number.isFinite(start) || !Number.isFinite(amount) || amount <= 0) return null;
  return {
    targetSheetIndex: sheetIndex,
    axis,
    action: kind,
    start,
    amount,
  };
}



function applyWorkbookReferenceShiftForStructureChange(context) {
  const model = state.model;
  const targetSheet = model?.sheets?.[context?.targetSheetIndex];
  if (!model || !targetSheet) return false;
  let changed = false;
  (model.sheets || []).forEach((sheet, sheetIndex) => {
    Object.entries(sheet.cells || {}).forEach(([key, cell]) => {
      if (!cell) return;
      let cellChanged = false;
      const nextFormula = shiftFormulaReferencesForStructureChange(cell.formula, context, sheetIndex);
      if (nextFormula !== cell.formula) {
        cell.formula = nextFormula;
        cell.raw = typeof cell.raw === "string" && cell.raw.trim().startsWith("=") ? nextFormula : cell.raw;
        cell.cached = null;
        cell.display = "";
        cellChanged = true;
      } else if (typeof cell.raw === "string" && cell.raw.trim().startsWith("=")) {
        const nextRaw = shiftFormulaReferencesForStructureChange(cell.raw, context, sheetIndex);
        if (nextRaw !== cell.raw) {
          cell.raw = nextRaw;
          cell.formula = nextRaw;
          cell.cached = null;
          cell.display = "";
          cellChanged = true;
        }
      }
      const validation = shiftValidationReferencesForStructureChange(cell.validation, context, sheetIndex);
      if (validation !== cell.validation) {
        cell.validation = validation;
        cellChanged = true;
      }
      if (cellChanged) {
        storeCellModelForSheet(sheet, key, cell);
        changed = true;
      }
    });
    (sheet.tables || []).forEach((table) => {
      (table.columns || []).forEach((column) => {
        ["calculatedColumnFormula", "totalsRowFormula"].forEach((field) => {
          if (!column?.[field]) return;
          const formula = String(column[field]);
          const shifted = shiftFormulaReferencesForStructureChange(formula.startsWith("=") ? formula : `=${formula}`, context, sheetIndex);
          const normalized = shifted.startsWith("=") ? shifted.slice(1) : shifted;
          if (normalized !== column[field]) {
            column[field] = normalized;
            changed = true;
          }
        });
      });
    });
    (sheet.images || []).forEach((image) => {
      if (!image?.chart) return;
      if (shiftChartReferencesForStructureChange(image.chart, context, sheetIndex)) {
        refreshChartDataFromSources(image.chart, sheetIndex);
        refreshShapeSource(image);
        changed = true;
      }
    });
  });
  (model.definedNames || []).forEach((definition) => {
    const ownerSheetIndex = definition.scope === "sheet" ? Number(definition.sheetIndex) : null;
    const shifted = shiftFormulaReferencesForStructureChange(definition.refersTo, context, ownerSheetIndex);
    if (shifted !== definition.refersTo) {
      definition.refersTo = shifted;
      changed = true;
    }
  });
  return changed;
}



function shiftChartReferencesForStructureChange(chart, context, ownerSheetIndex) {
  if (!chart || !context) return false;
  let changed = false;
  const shiftTextField = (object, field) => {
    if (!object || typeof object[field] !== "string" || !object[field]) return;
    const shifted = shiftReferenceFormulaTextForStructureChange(object[field], context, ownerSheetIndex);
    if (shifted !== object[field]) {
      object[field] = shifted;
      changed = true;
    }
  };

  shiftTextField(chart, "labelsRef");
  shiftTextField(chart, "titleRef");
  shiftTextField(chart, "sourceRef");
  if (typeof chart.sourceRange === "string") {
    shiftTextField(chart, "sourceRange");
  } else if (chart.sourceRange && typeof chart.sourceRange === "object") {
    const shiftedRange = context.axis === "row"
      ? shiftRangeRows(chart.sourceRange, context.start, context.amount, context.action)
      : shiftRangeColumns(chart.sourceRange, context.start, context.amount, context.action);
    if (!shiftedRange) {
      chart.sourceRange = null;
      changed = true;
    } else if (!rangeRectEquals(chart.sourceRange, shiftedRange)) {
      chart.sourceRange = { ...chart.sourceRange, ...shiftedRange };
      changed = true;
    }
  }

  (chart.series || []).forEach((series) => {
    ["nameRef", "valuesRef", "yValuesRef", "xValuesRef"].forEach((field) => shiftTextField(series, field));
  });
  return changed;
}



function refreshWorkbookChartsFromSources(options = {}) {
  let changed = false;
  (state.model?.sheets || []).forEach((sheet, sheetIndex) => {
    (sheet.images || []).forEach((image) => {
      if (!image?.chart?.series?.length) return;
      if (
        options.sourceSheetIndex != null &&
        options.sourceRange &&
        !chartSourceReferencesOverlap(image.chart, sheetIndex, options.sourceSheetIndex, options.sourceRange)
      ) {
        return;
      }
      const previousSource = image.src;
      const previousMimeType = image.mimeType;
      const previousExtension = image.extension;
      const dataChanged = refreshChartDataFromSources(image.chart, sheetIndex);
      refreshShapeSource(image);
      if (dataChanged || image.src !== previousSource || image.mimeType !== previousMimeType || image.extension !== previousExtension) {
        changed = true;
      }
    });
  });
  return changed;
}



function refreshWorkbookChartsAfterSourceEdit(sourceSheetIndex, sourceRange) {
  const changed = refreshWorkbookChartsFromSources({ sourceSheetIndex, sourceRange });
  // Some imported chart references arrive in formats that are valid Excel OOXML
  // but not yet narrowed by chartSourceReferencesOverlap. Re-read every chart as
  // a conservative fallback so saved standalone HTML and local edits do not keep
  // stale SVG data.
  return changed || refreshWorkbookChartsFromSources();
}



function refreshChartDataFromSources(chart, ownerSheetIndex) {
  if (!chart?.series?.length) return false;
  let changed = false;
  const title = chartSingleTextValueForReference(chart.titleRef, ownerSheetIndex);
  if (title != null && chart.title !== title) {
    chart.title = title;
    changed = true;
  }
  const labels = chartTextValuesForReference(chart.labelsRef, ownerSheetIndex);
  if (labels && !chartArraysEqual(chart.labels, labels)) {
    chart.labels = labels;
    changed = true;
  }
  (chart.series || []).forEach((series) => {
    const name = chartSingleTextValueForReference(series.nameRef, ownerSheetIndex);
    if (name != null && series.name !== name) {
      series.name = name;
      changed = true;
    }
    const values = chartNumericValuesForReference(series.valuesRef || series.xValuesRef, ownerSheetIndex);
    if (values && !chartArraysEqual(series.values, values)) {
      series.values = values;
      changed = true;
    }
    const yValues = chartNumericValuesForReference(series.yValuesRef, ownerSheetIndex);
    if (yValues && !chartArraysEqual(series.yValues, yValues)) {
      series.yValues = yValues;
      changed = true;
    }
  });
  return changed;
}



function chartTextValuesForReference(reference, ownerSheetIndex) {
  const parsed = chartReferenceRange(reference, ownerSheetIndex);
  if (!parsed) return null;
  const sheet = state.model?.sheets?.[parsed.sheetIndex];
  if (!sheet) return null;
  const sheetId = getSheetId(sheet.name);
  const values = [];
  forEachCellInRange(parsed.range, (row, col) => {
    values.push(String(getDisplayForCell(sheet, sheetId, row, col) ?? ""));
  });
  return values;
}



function chartSingleTextValueForReference(reference, ownerSheetIndex) {
  const values = chartTextValuesForReference(reference, ownerSheetIndex);
  return values?.length ? values[0] : null;
}



function chartNumericValuesForReference(reference, ownerSheetIndex) {
  const parsed = chartReferenceRange(reference, ownerSheetIndex);
  if (!parsed) return null;
  const sheet = state.model?.sheets?.[parsed.sheetIndex];
  if (!sheet) return null;
  const sheetId = getSheetId(sheet.name);
  const values = [];
  forEachCellInRange(parsed.range, (row, col) => {
    values.push(chartCellNumericValue(sheet, sheetId, row, col));
  });
  return values;
}



function chartCellNumericValue(sheet, sheetId, row, col) {
  const cell = sheet?.cells?.[cellKey(row, col)];
  let value = null;
  if (cell?.formula && state.hf && sheetId != null) {
    try {
      value = state.hf.getCellValue({ sheet: sheetId, row: row - 1, col: col - 1 });
    } catch {
      value = null;
    }
  }
  if (value == null) value = cell?.cached ?? cell?.raw ?? getDisplayForCell(sheet, sheetId, row, col);
  if (value instanceof Date && !Number.isNaN(value.getTime())) return dateToExcelSerial(value);
  if (typeof value === "object" && value?.value != null) value = value.value;
  const number = typeof value === "number" ? value : Number(String(value ?? "").replace(/[,¥$€£%]/g, ""));
  return Number.isFinite(number) ? number : null;
}



function chartReferenceRange(reference, ownerSheetIndex) {
  const parsed = parseChartReference(reference, ownerSheetIndex);
  if (!parsed) return null;
  const range = decodeRange(parsed.range.replace(/\$/g, ""));
  if (!range) return null;
  return { sheetIndex: parsed.sheetIndex, range };
}



function parseChartReference(reference, ownerSheetIndex) {
  let text = String(reference || "").trim();
  if (!text) return null;
  if (text.startsWith("=")) text = text.slice(1).trim();
  if (!text || /[\[\]]/.test(text) || text.includes(",")) return null;
  let sheetIndex = Number.isInteger(Number(ownerSheetIndex)) ? Number(ownerSheetIndex) : state.activeSheetIndex;
  let rangeText = text;
  const bangIndex = chartReferenceBangIndex(text);
  if (bangIndex >= 0) {
    const sheetName = unquoteSheetName(text.slice(0, bangIndex));
    const resolvedIndex = chartSheetIndexByName(sheetName);
    if (resolvedIndex < 0) return null;
    sheetIndex = resolvedIndex;
    rangeText = text.slice(bangIndex + 1);
  }
  return /^[\$A-Z0-9:]+$/i.test(rangeText) ? { sheetIndex, range: rangeText } : null;
}



function chartReferenceBangIndex(text) {
  const source = String(text || "");
  if (!source.includes("!")) return -1;
  if (source.startsWith("'")) {
    let index = 1;
    while (index < source.length) {
      if (source[index] === "'" && source[index + 1] === "'") {
        index += 2;
        continue;
      }
      if (source[index] === "'") return source[index + 1] === "!" ? index + 1 : -1;
      index += 1;
    }
    return -1;
  }
  return source.lastIndexOf("!");
}



function chartSheetIndexByName(sheetName) {
  const target = normalizeDefinedNameLookup(sheetName);
  return (state.model?.sheets || []).findIndex((sheet) => normalizeDefinedNameLookup(sheet.name) === target);
}



function chartArraysEqual(left, right) {
  if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) return false;
  return left.every((value, index) => {
    const next = right[index];
    if (value == null || next == null) return value == null && next == null;
    if (typeof value === "number" || typeof next === "number") return Number(value) === Number(next);
    return String(value) === String(next);
  });
}



function chartSourceOwnerSheetIndexesForRange(sourceSheetIndex, range) {
  const indexes = new Set();
  (state.model?.sheets || []).forEach((sheet, sheetIndex) => {
    const hasDependentChart = (sheet.images || []).some((image) =>
      image?.chart?.series?.length && chartSourceReferencesOverlap(image.chart, sheetIndex, sourceSheetIndex, range),
    );
    if (hasDependentChart) indexes.add(sheetIndex);
  });
  return [...indexes];
}



function chartSourceReferencesOverlap(chart, ownerSheetIndex, sourceSheetIndex, range) {
  if (!chart || !range) return false;
  return chartReferenceRanges(chart, ownerSheetIndex).some((reference) =>
    reference.sheetIndex === sourceSheetIndex && rangesIntersect(reference.range, range),
  );
}



function chartReferenceRanges(chart, ownerSheetIndex) {
  const references = [];
  ["labelsRef", "titleRef", "sourceRef", "sourceRange"].forEach((field) => {
    const range = chartReferenceRange(chart?.[field], ownerSheetIndex);
    if (range) references.push(range);
  });
  (chart?.series || []).forEach((series) => {
    ["nameRef", "valuesRef", "yValuesRef", "xValuesRef"].forEach((field) => {
      const range = chartReferenceRange(series?.[field], ownerSheetIndex);
      if (range) references.push(range);
    });
  });
  return references;
}



function shiftValidationReferencesForStructureChange(validation, context, ownerSheetIndex) {
  if (!validation) return validation;
  let changed = false;
  let nextValidation = validation;
  if (typeof validation.formula === "string" && validation.formula) {
    const shiftedFormula = shiftReferenceFormulaTextForStructureChange(validation.formula, context, ownerSheetIndex);
    if (shiftedFormula !== validation.formula) {
      nextValidation = {
        ...nextValidation,
        formula: shiftedFormula,
        options: validation.type === "list" ? parseValidationOptions(shiftedFormula) : nextValidation.options,
      };
      changed = true;
    }
  }
  if (validation.formulae?.length) {
    const nextFormulae = validation.formulae.map((formula) => {
      if (typeof formula !== "string") return formula;
      const shifted = shiftReferenceFormulaTextForStructureChange(formula, context, ownerSheetIndex);
      if (shifted !== formula) changed = true;
      return shifted;
    });
    nextValidation = { ...nextValidation, formulae: nextFormulae };
  }
  return changed ? nextValidation : validation;
}



function shiftReferenceFormulaTextForStructureChange(formula, context, ownerSheetIndex) {
  if (!referenceFormulaTextLooksLikeReference(formula)) return formula;
  const prefix = formula.trim().startsWith("=") ? "" : "=";
  const shifted = shiftFormulaReferencesForStructureChange(`${prefix}${formula}`, context, ownerSheetIndex);
  return prefix && shifted.startsWith("=") ? shifted.slice(1) : shifted;
}



function referenceFormulaTextLooksLikeReference(formula) {
  if (typeof formula !== "string") return false;
  return formula.trim().startsWith("=")
    || /(?:'[^']+'|[A-Za-z0-9_]+)?!?(?:\$?[A-Z]{1,3}\$?\d+|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)/i.test(formula);
}



function shiftFormulaReferencesForStructureChange(formula, context, ownerSheetIndex) {
  if (typeof formula !== "string" || !formula || !context) return formula;
  const targetSheet = state.model?.sheets?.[context.targetSheetIndex];
  if (!targetSheet) return formula;
  const tokens = formulaReferenceTokensInText(formula);
  if (!tokens.length) return formula;
  let output = "";
  let cursor = 0;
  for (const token of tokens) {
    output += formula.slice(cursor, token.start);
    const sheetPrefix = token.sheetToken ? `${token.sheetToken}!` : "";
    if (!formulaReferenceTargetsSheet(sheetPrefix, ownerSheetIndex, context.targetSheetIndex, targetSheet.name)) {
      output += token.text;
      cursor = token.end;
      continue;
    }
    const shifted = shiftA1ReferenceForStructureChange(token.body, context);
    output += shifted === "#REF!" ? "#REF!" : `${sheetPrefix}${shifted}`;
    cursor = token.end;
  }
  output += formula.slice(cursor);
  return output;
}



function formulaReferenceTargetsSheet(sheetPrefix, ownerSheetIndex, targetSheetIndex, targetSheetName) {
  if (sheetPrefix) {
    const name = unquoteSheetName(sheetPrefix.slice(0, -1));
    return name.toLocaleLowerCase("ja-JP") === String(targetSheetName || "").toLocaleLowerCase("ja-JP");
  }
  if (ownerSheetIndex == null) return false;
  return Number(ownerSheetIndex) === Number(targetSheetIndex);
}



function shiftA1ReferenceForStructureChange(reference, context) {
  const columnRangeMatch = String(reference || "").match(/^(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})$/i);
  if (columnRangeMatch) {
    if (context.axis !== "col") return reference;
    const [, startAbs, startText, endAbs, endText] = columnRangeMatch;
    const startCol = columnNumber(startText);
    const endCol = columnNumber(endText);
    const shifted = shiftRangeAxisForStructureChange(
      Math.min(startCol, endCol),
      Math.max(startCol, endCol),
      context.start,
      context.amount,
      context.action,
    );
    if (!shifted) return "#REF!";
    const nextStart = startCol <= endCol ? shifted.start : shifted.end;
    const nextEnd = startCol <= endCol ? shifted.end : shifted.start;
    return `${startAbs || ""}${columnName(nextStart)}:${endAbs || ""}${columnName(nextEnd)}`;
  }

  const rowRangeMatch = String(reference || "").match(/^(\$?)(\d+):(\$?)(\d+)$/);
  if (rowRangeMatch) {
    if (context.axis !== "row") return reference;
    const [, startAbs, startText, endAbs, endText] = rowRangeMatch;
    const startRow = Number(startText);
    const endRow = Number(endText);
    const shifted = shiftRangeAxisForStructureChange(
      Math.min(startRow, endRow),
      Math.max(startRow, endRow),
      context.start,
      context.amount,
      context.action,
    );
    if (!shifted) return "#REF!";
    const nextStart = startRow <= endRow ? shifted.start : shifted.end;
    const nextEnd = startRow <= endRow ? shifted.end : shifted.start;
    return `${startAbs || ""}${nextStart}:${endAbs || ""}${nextEnd}`;
  }

  const match = String(reference || "").match(/^(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?$/i);
  if (!match) return reference;
  const [, colAbs, colName, rowAbs, rowText, endColAbs, endColName, endRowAbs, endRowText] = match;
  const startPoint = { col: columnNumber(colName), row: Number(rowText), colAbs, rowAbs };
  if (!endColName) {
    const shifted = shiftA1PointForStructureChange(startPoint, context);
    return shifted ? formatA1Point(shifted) : "#REF!";
  }
  const endPoint = { col: columnNumber(endColName), row: Number(endRowText), colAbs: endColAbs, rowAbs: endRowAbs };
  const range = {
    top: Math.min(startPoint.row, endPoint.row),
    left: Math.min(startPoint.col, endPoint.col),
    bottom: Math.max(startPoint.row, endPoint.row),
    right: Math.max(startPoint.col, endPoint.col),
  };
  const shiftedRange = shiftA1RangeForStructureChange(range, context);
  if (!shiftedRange) return "#REF!";
  const start = {
    col: startPoint.col <= endPoint.col ? shiftedRange.left : shiftedRange.right,
    row: startPoint.row <= endPoint.row ? shiftedRange.top : shiftedRange.bottom,
    colAbs,
    rowAbs,
  };
  const end = {
    col: startPoint.col <= endPoint.col ? shiftedRange.right : shiftedRange.left,
    row: startPoint.row <= endPoint.row ? shiftedRange.bottom : shiftedRange.top,
    colAbs: endColAbs,
    rowAbs: endRowAbs,
  };
  return `${formatA1Point(start)}:${formatA1Point(end)}`;
}



function shiftA1PointForStructureChange(point, context) {
  if (context.axis === "row") {
    const row = shiftStructureIndex(point.row, context.start, context.amount, context.action);
    return row == null ? null : { ...point, row };
  }
  const col = shiftStructureIndex(point.col, context.start, context.amount, context.action);
  return col == null ? null : { ...point, col };
}



function shiftA1RangeForStructureChange(range, context) {
  if (context.axis === "row") {
    const rows = shiftRangeAxisForStructureChange(range.top, range.bottom, context.start, context.amount, context.action);
    if (!rows) return null;
    return { ...range, top: rows.start, bottom: rows.end };
  }
  const cols = shiftRangeAxisForStructureChange(range.left, range.right, context.start, context.amount, context.action);
  if (!cols) return null;
  return { ...range, left: cols.start, right: cols.end };
}



function shiftRangeAxisForStructureChange(startValue, endValue, start, amount, action) {
  const deleteEnd = start + amount - 1;
  if (action === "insert") {
    if (endValue < start) return { start: startValue, end: endValue };
    if (startValue >= start) return { start: startValue + amount, end: endValue + amount };
    return { start: startValue, end: endValue + amount };
  }
  if (endValue < start) return { start: startValue, end: endValue };
  if (startValue > deleteEnd) return { start: startValue - amount, end: endValue - amount };
  if (startValue >= start && endValue <= deleteEnd) return null;
  if (startValue < start && endValue > deleteEnd) return { start: startValue, end: endValue - amount };
  if (startValue < start) return start - 1 >= startValue ? { start: startValue, end: start - 1 } : null;
  return endValue - amount >= start ? { start, end: endValue - amount } : null;
}



function formatA1Point(point) {
  return `${point.colAbs || ""}${columnName(point.col)}${point.rowAbs || ""}${point.row}`;
}



function isFormulaIndexInsideStringLiteral(formula, index) {
  let inside = false;
  for (let i = 0; i < index; i += 1) {
    if (formula[i] !== '"') continue;
    if (formula[i + 1] === '"') {
      i += 1;
      continue;
    }
    inside = !inside;
  }
  return inside;
}



function validateAutoFillMerges(sheet, sourceRange, fillRange, direction) {
  const sourceMerges = mergesIntersectingRange(sheet, sourceRange);
  const targetMerges = mergesIntersectingRange(sheet, fillRange);
  const sourceHasPartialMerge = sourceMerges.some((merge) => !rangeContains(sourceRange, merge));
  if (sourceHasPartialMerge) {
    return { ok: false, message: "元の選択範囲には、結合セル全体を含めてください。" };
  }
  if (!sourceMerges.length && targetMerges.length) {
    return { ok: false, message: "結合セルのサイズが元の選択範囲と一致しないため、オートフィルできません。" };
  }
  if (!sourceMerges.length) return { ok: true };

  const sourceHeight = sourceRange.bottom - sourceRange.top + 1;
  const sourceWidth = sourceRange.right - sourceRange.left + 1;
  const fillHeight = fillRange.bottom - fillRange.top + 1;
  const fillWidth = fillRange.right - fillRange.left + 1;
  const vertical = direction === "up" || direction === "down";
  if ((vertical && fillHeight % sourceHeight !== 0) || (!vertical && fillWidth % sourceWidth !== 0)) {
    return { ok: false, message: "結合セルのサイズが元の選択範囲と一致しないため、オートフィルできません。" };
  }

  const expected = projectedAutoFillMerges(sourceMerges, sourceRange, fillRange, direction);
  const unexpectedTargetMerge = targetMerges.some((merge) => !expected.some((item) => rangesEqualLoose(item, merge)));
  if (unexpectedTargetMerge) {
    return { ok: false, message: "結合セルのサイズが元の選択範囲と一致しないため、オートフィルできません。" };
  }
  return { ok: true };
}



function applyAutoFillMerges(sheet, sourceRange, fillRange, direction) {
  const sourceMerges = mergesIntersectingRange(sheet, sourceRange).filter((merge) => rangeContains(sourceRange, merge));
  if (!sourceMerges.length) return;
  sheet.merges.push(...projectedAutoFillMerges(sourceMerges, sourceRange, fillRange, direction));
}



function projectedAutoFillMerges(sourceMerges, sourceRange, fillRange, direction) {
  const sourceHeight = sourceRange.bottom - sourceRange.top + 1;
  const sourceWidth = sourceRange.right - sourceRange.left + 1;
  const vertical = direction === "up" || direction === "down";
  const blockCount = vertical
    ? (fillRange.bottom - fillRange.top + 1) / sourceHeight
    : (fillRange.right - fillRange.left + 1) / sourceWidth;
  const merges = [];
  for (let block = 0; block < blockCount; block += 1) {
    const rowOffset = vertical ? fillRange.top - sourceRange.top + block * sourceHeight : 0;
    const colOffset = vertical ? 0 : fillRange.left - sourceRange.left + block * sourceWidth;
    sourceMerges.forEach((merge) => {
      merges.push({
        top: merge.top + rowOffset,
        left: merge.left + colOffset,
        bottom: merge.bottom + rowOffset,
        right: merge.right + colOffset,
      });
    });
  }
  return merges;
}



function mergesIntersectingRange(sheet, range) {
  return (sheet.merges || []).filter((merge) => rangesIntersect(merge, range));
}



function rangeContains(container, item) {
  return item.top >= container.top && item.bottom <= container.bottom && item.left >= container.left && item.right <= container.right;
}



function rangesEqualLoose(a, b) {
  return Boolean(a && b && a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right);
}



function positiveModulo(value, divisor) {
  return ((value % divisor) + divisor) % divisor;
}



function commitActiveCellEditForExternalPointer(event) {
  const editor = $sheetHost.find(".sheet-cell[contenteditable='true']")[0];
  if (!editor || editor.contains(event.target)) return false;
  editor.blur();
  return true;
}



function setPointerCaptureSafely(element, pointerId) {
  try {
    element?.setPointerCapture?.(pointerId);
  } catch {
    // Re-rendering can replace the original target before pointer capture is applied.
  }
}



function startCellRangeSelection(event) {
  if (event.button !== 0) return;
  if (state.pendingInkTool) {
    startInkDrawing(event);
    return;
  }
  if (state.pendingShapeType) {
    startShapeDrawing(event);
    return;
  }
  if (selectLineShapeFromSheetPointer(event)) return;
  if (startFormulaReferenceSelection(event)) return;
  commitActiveCellEditForExternalPointer(event);
  if (state.formatPainter) {
    const formatPoint = formatPainterPointFromPointerEvent(event);
    if (formatPoint) {
      event.preventDefault();
      event.stopPropagation();
      state.suppressNextCellClick = true;
      const range = selectionRangeFromPoints(formatPoint, formatPoint);
      if (!ensureSelectableRangeForProtection(activeSheet(), range)) return;
      applySelectionRange(range, formatPoint);
      state.selectionAnchor = formatPoint;
      state.selectionStart = formatPoint;
      state.selectionDrag = { kind: "format-painter", anchor: formatPoint, moved: false };
      state.isSelecting = true;
      event.currentTarget.setPointerCapture?.(event.pointerId);
      updateSelectionUi();
      updateFormulaBarForSelection();
      return;
    }
  }
  if (event.target?.closest?.(".sheet-page-layout-header, .sheet-page-layout-footer")) return;
  if (event.target?.closest?.(".sheet-view-guide")) return;
  const endpoint = cellPointFromPointerEvent(event);
  if (!endpoint) return;
  clearFormulaReferenceState();

  const additive = event.ctrlKey || event.metaKey;
  const baseRanges = additive ? rawActiveSelectionRanges() : [];
  const anchor =
    event.shiftKey && state.selectionStart?.sheetIndex === state.activeSheetIndex
      ? state.selectionStart
      : endpoint;
  const range = selectionRangeFromPoints(anchor, endpoint);
  if (!ensureSelectableRangeForProtection(activeSheet(), range)) return;

  state.selectedInkStrokeIds = [];
  applySelectionRange(range, endpoint, { additive, baseRanges });
  state.selectionAnchor = anchor;
  state.selectionStart = anchor;
  state.selectionDrag = { kind: "cell", additive, baseRanges, anchor };
  state.isSelecting = true;
  event.currentTarget.setPointerCapture?.(event.pointerId);
  if (event.shiftKey || additive) {
    state.suppressNextCellClick = true;
    event.preventDefault();
  }
  updateSelectionUi();
  updateFormulaBarForSelection();
}



function startPendingObjectDrawingFromPointer(event) {
  if (event.button !== 0) return false;
  if (state.pendingInkTool && state.pendingInkTool.type !== "eraser") {
    startInkDrawing(event);
    event.stopImmediatePropagation?.();
    return true;
  }
  if (state.pendingShapeType) {
    startShapeDrawing(event);
    event.stopImmediatePropagation?.();
    return true;
  }
  return false;
}



function selectImageFromEvent(event) {
  if (event.button !== 0) return;
  if (startPendingObjectDrawingFromPointer(event)) return;
  if (state.pendingInkTool?.type === "eraser") {
    eraseInkAtClientPoint(event.clientX, event.clientY);
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    return;
  }
  const imageId = event.currentTarget.dataset.imageId || event.target.closest("[data-image-id]")?.dataset.imageId;
  if (!imageId) return;
  const image = activeSheet().images?.find((item) => item.id === imageId);
  if (isLineShapeObject(image) && !lineShapePointerHitsStroke(activeSheet(), image, event)) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  commitActiveCellEditForExternalPointer(event);
  if (event.detail >= 2) {
    if (image?.shape && !isLineShapeObject(image)) {
      if (!isActiveImage(imageId)) setSelectedImageIds(selectionIdsForImage(activeSheet(), imageId), imageId);
      openShapeTextEditor(imageId, { clientX: event.clientX, clientY: event.clientY });
      return;
    }
    if (image && !image.shape) {
      if (!isActiveImage(imageId)) {
        setSelectedImageIds(selectionIdsForImage(activeSheet(), imageId), imageId);
        renderSheet();
        updateSelectionUi();
      }
      activateObjectFormatRibbonForSelection();
      return;
    }
  }
  const additive = event.ctrlKey || event.metaKey || event.shiftKey;
  const toggle = event.ctrlKey || event.metaKey;
  if (additive) {
    selectImage(imageId, { additive, toggle });
    return;
  }
  startSelectedImageTransform(event);
}



function handlePictureFormatDoubleClick(event) {
  if (event.button && event.button !== 0) return;
  if (event.target?.closest?.(".sheet-image-resize-handle, .sheet-shape-adjust-handle, .sheet-image-rotate-handle, .sheet-picture-crop-handle")) return;
  const imageId = event.currentTarget?.dataset?.imageId || event.target.closest("[data-image-id]")?.dataset.imageId;
  if (!imageId) return;
  const sheet = activeSheet();
  const image = sheet.images?.find((item) => item.id === imageId);
  if (!image || image.shape) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  commitActiveCellEditForExternalPointer(event);
  if (!isActiveImage(imageId)) {
    setSelectedImageIds(selectionIdsForImage(sheet, imageId), imageId);
    renderSheet();
    updateSelectionUi();
  }
  activateObjectFormatRibbonForSelection();
  setStatus("図の形式を表示しました。");
}



function lineShapeHitTargetFromPointerEvent(event) {
  const sheet = activeSheet();
  return (sheet.images || [])
    .filter((image) => image && !image.hidden && isLineShapeObject(image))
    .sort((a, b) => worksheetObjectLayerZIndex(sheet, b) - worksheetObjectLayerZIndex(sheet, a))
    .find((image) => lineShapePointerHitsStroke(sheet, image, event)) || null;
}



function selectLineShapeFromSheetPointer(event) {
  if (event.button !== 0) return false;
  if (event.target?.closest?.(".sheet-image, .sheet-image-link, .sheet-image-selection-frame, .sheet-fill-handle, .cell-filter-button, .sheet-page-layout-header, .sheet-page-layout-footer, .sheet-view-guide")) return false;
  const image = lineShapeHitTargetFromPointerEvent(event);
  if (!image) return false;
  if (event.ctrlKey || event.metaKey || event.shiftKey) {
    selectImage(image.id, { additive: true, toggle: event.ctrlKey || event.metaKey });
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    return true;
  }
  startSelectedImageTransform(event, { imageId: image.id, captureTarget: $sheetHost[0] });
  return true;
}



function worksheetObjectFromPointerEventTarget(event) {
  const imageId = event.target?.closest?.("[data-image-id]")?.dataset?.imageId || "";
  if (!imageId) return null;
  return activeSheet().images?.find((item) => item.id === imageId) || null;
}



function lineShapeShouldCapturePointerEvent(event, line) {
  if (!line) return false;
  if (event.target?.closest?.(".sheet-image-resize-handle, .sheet-shape-adjust-handle, .sheet-image-rotate-handle, .sheet-picture-crop-handle, .sheet-fill-handle, .cell-filter-button, .sheet-page-layout-header, .sheet-page-layout-footer, .sheet-view-guide")) {
    return false;
  }
  const targetObject = worksheetObjectFromPointerEventTarget(event);
  if (!targetObject || targetObject.id === line.id) return true;
  return worksheetObjectLayerZIndex(activeSheet(), line) >= worksheetObjectLayerZIndex(activeSheet(), targetObject);
}



function handleLineShapePointerDownCapture(event) {
  if (event.button !== 0) return;
  if (state.pendingShapeType || state.pendingInkTool || state.shapeDrawing) return;
  const line = lineShapeHitTargetFromPointerEvent(event);
  if (!line || !lineShapeShouldCapturePointerEvent(event, line)) return;
  commitActiveCellEditForExternalPointer(event);
  if (event.ctrlKey || event.metaKey || event.shiftKey) {
    selectImage(line.id, { additive: true, toggle: event.ctrlKey || event.metaKey });
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    return;
  }
  startSelectedImageTransform(event, { imageId: line.id, captureTarget: $sheetHost[0] });
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
}



function handleLineShapeContextMenuCapture(event) {
  if (!state.model) return;
  const line = lineShapeHitTargetFromPointerEvent(event);
  if (!line || !lineShapeShouldCapturePointerEvent(event, line)) return;
  if (!isActiveImage(line.id)) setSelectedImageIds(selectionIdsForImage(activeSheet(), line.id), line.id);
  openImageContextMenu(event, { imageId: line.id });
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
}



function selectImage(imageId, options = {}) {
  const sheet = activeSheet();
  const image = sheet.images?.find((item) => item.id === imageId);
  if (!image) return;
  const currentIds = activeSelectedImageIds();
  const imageSelectionIds = selectionIdsForImage(sheet, imageId, options);
  let nextIds = imageSelectionIds;
  if (options.additive) {
    const imageIdSet = new Set(imageSelectionIds);
    const allSelected = imageSelectionIds.every((id) => currentIds.includes(id));
    if (options.toggle && allSelected) {
      nextIds = currentIds.filter((id) => !imageIdSet.has(id));
    } else {
      nextIds = [...currentIds, ...imageSelectionIds];
    }
  }

  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  state.selectionDrag = null;
  state.formulaEdit = null;
  state.isSelecting = false;
  setSelectedImageIds(nextIds, nextIds.includes(imageId) ? imageId : nextIds[0]);
  $formulaInput.prop("disabled", true).val("");
  renderSheet();
  updateSelectionUi();
  const count = activeSelectedImageIds().length;
  if (!count) {
    setStatus("図形の選択を解除しました。");
  } else if (count === 1) {
    setStatus(`${image.shape ? "図形" : "画像"}を選択しました：${image.name || imageId}`);
  } else {
    setStatus(`${count} 個のオブジェクトを選択しました。`);
  }
}



function startShapeTextEdit(event) {
  if (event.target?.closest?.(".sheet-image-resize-handle, .sheet-shape-adjust-handle, .sheet-image-rotate-handle")) return;
  const imageId = event.currentTarget.dataset.imageId || event.target.closest("[data-image-id]")?.dataset.imageId;
  const image = selectedImageObjectById(imageId);
  if (!image?.shape || isLineShapeObject(image)) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  if (!isActiveImage(imageId)) {
    selectImage(imageId);
  }
  openShapeTextEditor(imageId, { clientX: event.clientX, clientY: event.clientY });
}



function openShapeTextEditor(imageId = state.selectedImage?.id, options = {}) {
  const image = selectedImageObjectById(imageId);
  if (!image?.shape || isLineShapeObject(image)) {
    setStatus("テキストを編集できる図形を選択してください。");
    return false;
  }
  if (!ensureSheetProtectionPermission(activeSheet(), "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) return false;
  cancelShapeTextEditor({ render: false });
  const sheet = activeSheet();
  const box = imageBoxToPixels(sheet, image);
  const grid = $sheetHost.find(".sheet-grid")[0];
  if (!box || !grid) return false;

  const editor = document.createElement("textarea");
  editor.id = "shapeTextEditor";
  editor.className = "sheet-shape-text-editor";
  editor.spellcheck = false;
  editor.dataset.imageId = image.id;
  editor.value = options.replaceText ? String(options.initialText ?? "") : shapeTextPlainValue(image.shape);
  editor.wrap = "soft";
  editor.setAttribute("role", "textbox");
  editor.setAttribute("aria-multiline", "true");
  editor.setAttribute("aria-label", `${image.name || "図形"} テキスト`);
  ["pointerdown", "mousedown", "click", "dblclick"].forEach((type) => {
    editor.addEventListener(type, stopShapeTextEditorPointerEvent);
  });
  applyShapeTextEditorStyle(editor, image, box);
  grid.append(editor);
  hideRenderedShapeText(image);
  state.shapeTextEditor = {
    sheetIndex: state.activeSheetIndex,
    imageId: image.id,
    beforeText: shapeTextPlainValue(image.shape),
    sheetBefore: captureSheetForHistory(sheet),
    element: editor,
  };
  editor.focus({ preventScroll: true });
  if (options.replaceText) {
    placeCaretAtEnd(editor);
  } else if (Number.isFinite(options.clientX) && Number.isFinite(options.clientY)) {
    placeCaretAtPoint(options.clientX, options.clientY, editor);
  } else {
    placeCaretAtEnd(editor);
  }
  setStatus(`${image.name || "図形"} のテキストを編集中です。`);
  return true;
}



function applyShapeTextEditorStyle(editor, image, box) {
  const shape = image.shape || {};
  const vertical = ["top", "bottom", "middle"].includes(shape.textVAlign) ? shape.textVAlign : "middle";
  const styleText = imageBoxStyle(image, box);
  const margin = normalizeShapeTextMargin(shape.textMargin, shape.type === "textBox");
  const textScale = Number(activeSheet()?.displayScale || 1);
  const style = Object.fromEntries(styleText.split(";").filter(Boolean).map((item) => {
    const index = item.indexOf(":");
    return [item.slice(0, index), item.slice(index + 1)];
  }));
  Object.assign(editor.style, style, {
    padding: `${margin.top * textScale}px ${margin.right * textScale}px ${margin.bottom * textScale}px ${margin.left * textScale}px`,
    color: shape.textColor || "#1f2937",
    fontFamily: shape.textFontFamily || fontFamilyForExcel("Calibri") || '"Yu Gothic UI", Meiryo, sans-serif',
    fontSize: `${Math.max(8, Math.round((Number(shape.textFontSize) || 11) * 96 / 72)) * textScale}px`,
    fontWeight: shape.textBold ? "700" : "400",
    fontStyle: shape.textItalic ? "italic" : "normal",
    textDecorationLine: underlinePresetForValue(shape.textUnderline) ? "underline" : "none",
    textDecorationStyle: shape.textUnderline === "underline double" ? "double" : "solid",
    textAlign: shape.textAlign || "center",
    lineHeight: "1.22",
    whiteSpace: "pre-wrap",
    overflow: "hidden",
    resize: "none",
    display: "block",
    backgroundColor: "rgba(255, 255, 255, 0.72)",
    zIndex: String(SHAPE_TEXT_EDITOR_Z_INDEX),
    pointerEvents: "auto",
    WebkitTextStroke: shape.textOutlineColor && shape.textOutlineColor !== "none" ? `${normalizeShapeTextOutlineWidth(shape.textOutlineWidth) * textScale}px ${shape.textOutlineColor}` : "0 transparent",
  });
}



function stopShapeTextEditorPointerEvent(event) {
  event.stopPropagation();
  event.stopImmediatePropagation?.();
}



function hideRenderedShapeText(image) {
  if (!image?.shape?.text && !image?.shape?.textRuns?.length) return;
  const descriptor = shapeDescriptorForType(image.shape.type || DEFAULT_SHAPE_TYPE, {
    ...image.shape,
    text: "",
    textRuns: [],
    width: image.ext?.width,
    height: image.ext?.height,
    includeTextInSvg: false,
  });
  const element = primaryImagePositionElement(image.id);
  const textLayer = element?.querySelector?.(".sheet-shape-text-layer");
  if (textLayer) textLayer.style.display = "none";
  setImageElementRenderSource(element, descriptor.src);
}



function handleShapeTextEditorKeydown(event) {
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  if (isElementImeCompositionEvent(event)) return;
  if (event.key === "Escape") {
    event.preventDefault();
    cancelShapeTextEditor();
  }
}



function commitShapeTextEditorOnOutsidePointer(event) {
  const editor = state.shapeTextEditor?.element;
  if (!editor || event.target?.closest?.("#shapeTextEditor")) return;
  commitShapeTextEditor();
}



function editableShapeText(element) {
  const value = isTextInputElement(element) ? element.value : element?.innerText ?? element?.textContent ?? "";
  return String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n");
}



function cancelShapeTextEditor(options = {}) {
  const editor = state.shapeTextEditor;
  if (!editor) return false;
  state.shapeTextEditor = null;
  editor.element?.remove();
  if (options.render !== false) {
    renderSheet();
    updateSelectionUi();
    setStatus("図形テキストの編集をキャンセルしました。");
  }
  return true;
}



function commitShapeTextEditor() {
  const editor = state.shapeTextEditor;
  if (!editor) return false;
  const text = editableShapeText(editor.element);
  state.shapeTextEditor = null;
  editor.element?.remove();
  if (editor.sheetIndex !== state.activeSheetIndex) return false;
  const sheet = activeSheet();
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) {
    renderSheet();
    updateSelectionUi();
    return false;
  }
  const image = sheet.images?.find((item) => item.id === editor.imageId);
  if (!image?.shape) {
    renderSheet();
    updateSelectionUi();
    return false;
  }
  if (text !== editor.beforeText) {
    if (commitShapeTextEditorToSheetGroup(image, text, editor)) return true;
    image.shape = { ...image.shape, text };
    delete image.shape.textRuns;
    refreshShapeSource(image);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetBefore: editor.sheetBefore,
      sheetAfter: captureSheetForHistory(sheet),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${image.name || "図形"} テキスト編集`,
    });
    setStatus(`${image.name || "図形"} のテキストを更新しました。`);
  } else {
    setStatus(`${image.name || "図形"} のテキスト編集を終了しました。`);
  }
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  return true;
}



function commitShapeTextEditorToSheetGroup(image, text, editor) {
  const items = groupedShapeObjectTargets([image]);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを編集できません。")) {
      renderSheet();
      updateSelectionUi();
      return true;
    }
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      const target = item.images[0];
      if (!target?.shape) return;
      target.shape = { ...target.shape, text };
      delete target.shape.textRuns;
      refreshShapeSource(target);
    });
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${items.length} シート ${image.name || "図形"} テキスト編集`,
    });
    setStatus(`${items.length} 個のシートの${image.name || "図形"} のテキストを更新しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`図形テキストを更新できませんでした：${error.message || error}`);
    return true;
  }
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  return true;
}



function groupedShapeObjectTargets(images, options = {}) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !images?.length) return [];
  const sourceSheet = activeSheet();
  const sourceImages = sourceSheet?.images || [];
  const refs = images
    .map((image) => ({ image, index: sourceImages.indexOf(image) }))
    .filter((ref) => ref.image && ref.index >= 0);
  if (refs.length !== images.length) return [];

  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet?.images) continue;
    const used = new Set();
    const targets = [];
    refs.forEach((ref) => {
      const target = correspondingGroupedShapeObject(sheet, sourceSheet, ref.image, ref.index, used);
      if (!target) return;
      used.add(target.id);
      targets.push(target);
    });
    if (targets.length === refs.length || (options.allowPartial && targets.length)) {
      items.push({ sheetIndex, sheet, images: targets, indexes: refs.map((ref) => ref.index) });
    }
  }
  return items.length > 1 ? items : [];
}



function correspondingGroupedShapeObject(sheet, sourceSheet, source, sourceIndex, used = new Set()) {
  if (sheet === sourceSheet) return source;
  const images = sheet.images || [];
  const direct = images[sourceIndex];
  if (direct && !used.has(direct.id) && groupedShapeObjectsCorrespond(source, direct)) return direct;
  const semantic = images.find((candidate) => !used.has(candidate.id) && groupedShapeObjectsCorrespond(source, candidate));
  if (semantic) return semantic;
  return direct && !used.has(direct.id) ? direct : null;
}



function groupedShapeObjectsCorrespond(source, candidate) {
  if (!source || !candidate) return false;
  if (source === candidate) return true;
  if (source.shape || candidate.shape) {
    return (
      source.shape?.type === candidate.shape?.type &&
      (source.shape?.objectKind === candidate.shape?.objectKind ||
        source.shape?.label === candidate.shape?.label ||
        source.name === candidate.name ||
        (!source.shape?.objectKind && !candidate.shape?.objectKind))
    );
  }
  if (source.chart || candidate.chart) return source.chart?.type === candidate.chart?.type;
  return (source.name && source.name === candidate.name) || (source.mimeType && source.mimeType === candidate.mimeType);
}



function selectedWorksheetObjectSummary(imageObjects = [], inkStrokes = []) {
  const total = imageObjects.length + inkStrokes.length;
  if (total === 1) return imageObjects[0]?.name || inkStrokes[0]?.name || "インク";
  if (!imageObjects.length) return `${inkStrokes.length} 個のインク`;
  if (!inkStrokes.length) return selectedObjectSummary(imageObjects);
  return `${total} 個のオブジェクト`;
}



function clearSelectedWorksheetObjects() {
  const hadSelection = activeSelectedImageIds().length > 0 || selectedInkStrokeIdSet().size > 0;
  if (!hadSelection) return false;
  setSelectedImageIds([], null);
  state.selectedInkStrokeIds = [];
  renderSheet();
  updateSelectionUi();
  setStatus("オブジェクトの選択を解除しました。");
  return true;
}



function deleteSelectedWorksheetObjects() {
  const sheet = activeSheet();
  const selectedStrokes = selectedInkStrokes(sheet);
  if (!selectedStrokes.length) return deleteSelectedImages();

  const ids = activeSelectedImageIds();
  const idSet = new Set(ids);
  const selectedImages = (sheet.images || []).filter((image) => idSet.has(image.id));
  if (!selectedImages.length && !selectedStrokes.length) return false;
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを削除できません。")) return false;

  const sheetBefore = captureSheetForHistory(sheet);
  const strokeIdSet = new Set(selectedStrokes.map((stroke) => String(stroke.id)));
  sheet.inkStrokes = normalizeInkStrokes(sheet.inkStrokes).filter((stroke) => !strokeIdSet.has(String(stroke.id)));
  if (idSet.size) sheet.images = (sheet.images || []).filter((image) => !idSet.has(image.id));

  setSelectedImageIds([], null);
  state.selectedInkStrokeIds = [];
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  if (state.objectClipboard?.sourceIds?.some((id) => idSet.has(id))) {
    state.objectClipboard = null;
    clearConsumedCutObjectClipboardLabel();
  }

  const summary = selectedWorksheetObjectSummary(selectedImages, selectedStrokes);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${summary} 削除`,
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`${summary}を削除しました。`);
  return true;
}



function deleteSelectedImages() {
  const ids = activeSelectedImageIds();
  if (!ids.length) return false;
  const sheet = activeSheet();
  if (deleteSelectedImagesForSheetGroup(ids)) return true;
  if (!ensureSheetProtectionPermission(sheet, "allowEditObjects", "保護されたシートではオブジェクトを削除できません。")) return false;
  const idSet = new Set(ids);
  const objects = (sheet.images || []).filter((image) => idSet.has(image.id));
  if (!objects.length) return false;
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.images = (sheet.images || []).filter((image) => !idSet.has(image.id));
  setSelectedImageIds([], null);
  state.selectedInkStrokeIds = [];
  state.selected = null;
  state.selectionRange = null;
  state.selectionRanges = [];
  state.selectionAnchor = null;
  state.selectionStart = null;
  if (state.objectClipboard?.sourceIds?.some((id) => idSet.has(id))) {
    state.objectClipboard = null;
    clearConsumedCutObjectClipboardLabel();
  }
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: null,
    selectionAfter: null,
    selectionBeforeRanges: [],
    selectionAfterRanges: [],
    label: `${selectedObjectSummary(objects)} 削除`,
  });
  renderSheet();
  updateSelectionUi();
  syncShapeFormatControls();
  setStatus(`${selectedObjectSummary(objects)}を削除しました。`);
  return true;
}



function deleteSelectedImagesForSheetGroup(ids) {
  const sheet = activeSheet();
  const idSet = new Set(ids || []);
  const objects = (sheet.images || []).filter((image) => idSet.has(image.id));
  const items = groupedShapeObjectTargets(objects);
  if (items.length <= 1) return false;
  for (const item of items) {
    if (!ensureSheetProtectionPermission(item.sheet, "allowEditObjects", "保護されたシートではオブジェクトを削除できません。")) return true;
  }
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      const targetIds = new Set(item.images.map((image) => image.id));
      item.sheet.images = (item.sheet.images || []).filter((image) => !targetIds.has(image.id));
    });
    setSelectedImageIds([], null);
    state.selectedInkStrokeIds = [];
    state.selected = null;
    state.selectionRange = null;
    state.selectionRanges = [];
    state.selectionAnchor = null;
    state.selectionStart = null;
    if (state.objectClipboard?.sourceIds?.some((id) => idSet.has(id))) {
      state.objectClipboard = null;
      clearConsumedCutObjectClipboardLabel();
    }
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: null,
      selectionAfter: null,
      selectionBeforeRanges: [],
      selectionAfterRanges: [],
      label: `${items.length} シート${selectedObjectSummary(objects)} 削除`,
    });
    renderSheet();
    updateSelectionUi();
    syncShapeFormatControls();
    setStatus(`${items.length} 個のシートの${selectedObjectSummary(objects)}を削除しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`オブジェクトを削除できませんでした：${error.message || error}`);
  }
  return true;
}

export {
  addCalendarMonthsToExcelSerial,
  addWeekdaysToExcelSerial,
  applyAutoFill,
  applyAutoFillCellFormatting,
  applyAutoFillMerges,
  applyAutoFillOption,
  applyCellEditCssPatch,
  applyCellEditCssPatchToSelection,
  applyCellEditUnderlinePreset,
  applyCssPatchToRichTextRuns,
  applyShapeTextEditorStyle,
  applyWorkbookReferenceShiftForStructureChange,
  autoFillAdjacentDataEndRow,
  autoFillAlphabetSeriesValue,
  autoFillAxisCellPoint,
  autoFillAxisContext,
  autoFillDateSerialFromParts,
  autoFillDateSeriesValue,
  autoFillDateValueSupportsCalendarOptions,
  autoFillDoubleClickRange,
  autoFillHistoryCellMap,
  autoFillListSeriesValue,
  autoFillMenuPointFromEvent,
  autoFillModeForPointerEvent,
  autoFillNumberFormatLooksCalendarDate,
  autoFillNumericSeriesEntry,
  autoFillNumericSeriesValue,
  autoFillNumericTrendSeriesEntry,
  autoFillNumericTrendSeriesValue,
  autoFillOptionItemsForSource,
  autoFillOptionStatusLabel,
  autoFillOptionsKeyCanLeaveTextInput,
  autoFillOptionsMatchCurrentSelection,
  autoFillPlainAlphabetEntries,
  autoFillPlainNumericEntries,
  autoFillRangeFromPoint,
  autoFillSeriesValue,
  autoFillShouldOwnTargetFormatting,
  autoFillSourceIsPlainAlphabet,
  autoFillSourcePoint,
  autoFillSourceSupportsDateOptions,
  autoFillSourceSupportsNumericTrendOptions,
  autoFillTextNumberSeriesValue,
  autoFillTimeSerialFromParts,
  cancelShapeTextEditor,
  cellAlphabetEntryForAutoFill,
  cellDateSeriesValueForAutoFill,
  cellEditBaseCss,
  cellEditInlineCssFromElement,
  cellEditRichTextForCommit,
  cellEditRunsFromDom,
  cellEditRunsHtml,
  cellEditSelectionEveryRunMatches,
  cellListEntryForAutoFill,
  cellNumericValueForAutoFill,
  cellTextNumberEntryForAutoFill,
  chartArraysEqual,
  chartCellNumericValue,
  chartNumericValuesForReference,
  chartReferenceBangIndex,
  chartReferenceRange,
  chartReferenceRanges,
  chartSheetIndexByName,
  chartSingleTextValueForReference,
  chartSourceOwnerSheetIndexesForRange,
  chartSourceReferencesOverlap,
  chartTextValuesForReference,
  circularListStep,
  clampAutoFillOptionsButtonPoint,
  clampAutoFillOptionsMenuPoint,
  clampAutoFillOptionsPoint,
  clearAutoFillOptions,
  clearAutoFillOptionsOnExternalFocus,
  clearSelectedWorksheetObjects,
  cloneCellForAutoFill,
  cloneCellWithAutoFillFormattingOnly,
  closeAutoFillOptionsMenu,
  closeAutoFillOptionsMenuOnEscape,
  closeAutoFillOptionsMenuOnOutsidePointer,
  collectCellEditRuns,
  commitActiveCellEditForExternalPointer,
  commitShapeTextEditor,
  commitShapeTextEditorOnOutsidePointer,
  commitShapeTextEditorToSheetGroup,
  correspondingGroupedShapeObject,
  daysInMonthUtc,
  defaultAutoFillMode,
  deleteSelectedImages,
  deleteSelectedImagesForSheetGroup,
  deleteSelectedWorksheetObjects,
  editableShapeText,
  finishAutoFillDrag,
  focusAutoFillOptionsMenuItem,
  formatA1Point,
  formatAutoFillDateText,
  formatAutoFillTimeText,
  formulaReferenceTargetsSheet,
  groupedShapeObjectTargets,
  groupedShapeObjectsCorrespond,
  handleAutoFillHandleDoubleClick,
  handleAutoFillOptionClick,
  handleAutoFillOptionsMenuKeyEvent,
  handleAutoFillOptionsMenuKeydown,
  handleGlobalAutoFillOptionsKeydown,
  handleLineShapeContextMenuCapture,
  handleLineShapePointerDownCapture,
  handlePictureFormatDoubleClick,
  handleShapeTextEditorKeydown,
  hideRenderedShapeText,
  insertLineBreakInTextControl,
  integerRange,
  isAltModifierEvent,
  isEnterKeyEvent,
  isFormulaIndexInsideStringLiteral,
  isSinglePlainNumberAutoFillSource,
  lineShapeHitTargetFromPointerEvent,
  lineShapeShouldCapturePointerEvent,
  mergeCellEditRuns,
  mergesIntersectingRange,
  normalizeAutoFillMode,
  normalizeAutoFillTrendNumber,
  normalizeCellEditColor,
  normalizeCellEditCssPatch,
  normalizeCellEditInlineCss,
  normalizeCellEditRichTextForModel,
  openAutoFillRightDragOptions,
  openShapeTextEditor,
  parseAutoFillDateText,
  parseAutoFillNumber,
  parseAutoFillTimeText,
  parseChartReference,
  positiveModulo,
  projectedAutoFillMerges,
  pushCellEditRun,
  rangeContains,
  rangesEqualLoose,
  referenceFormulaTextLooksLikeReference,
  refreshChartDataFromSources,
  refreshWorkbookChartsAfterSourceEdit,
  refreshWorkbookChartsFromSources,
  removeAutoFillOptionsElements,
  renderAutoFillOptions,
  renderAutoFillOptionsMenu,
  restoreAutoFillSourceSelection,
  selectImage,
  selectImageFromEvent,
  selectLineShapeFromSheetPointer,
  selectedWorksheetObjectSummary,
  setCellEditSelectionByOffsets,
  setPointerCaptureSafely,
  shiftA1PointForStructureChange,
  shiftA1RangeForStructureChange,
  shiftA1ReferenceForStructureChange,
  shiftBareFormulaReferenceForAutoFill,
  shiftChartReferencesForStructureChange,
  shiftFormulaReferenceForAutoFill,
  shiftFormulaReferences,
  shiftFormulaReferencesForStructureChange,
  shiftRangeAxisForStructureChange,
  shiftReferenceFormulaTextForStructureChange,
  shiftValidationReferencesForStructureChange,
  sourceCellForAutoFill,
  startAutoFillDrag,
  startCellRangeSelection,
  startPendingObjectDrawingFromPointer,
  startShapeTextEdit,
  stopShapeTextEditorPointerEvent,
  storeBlankAutoFillCellWithFormatting,
  structureReferenceShiftForOperation,
  suppressAutoFillRightDragContextMenu,
  syncAutoFillOptionsWithSelection,
  toggleAutoFillOptionsMenu,
  toggleCellEditCssForSelection,
  updateAutoFillDrag,
  validateAutoFillMerges,
  worksheetObjectFromPointerEventTarget,
  writeAutoFillCells,
};
