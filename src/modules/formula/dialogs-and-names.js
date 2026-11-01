// Feature module; shared bindings are imported explicitly.
import {
  $,
  $formulaInput,
  $sheetHost,
  FORMULA_ARGUMENT_DEFAULT_OVERRIDES,
  FORMULA_ARGUMENT_DESCRIPTIONS,
  FORMULA_REFERENCE_BODY_TOKEN_SOURCE,
  FORMULA_REFERENCE_COLORS,
  FORMULA_REFERENCE_SHEET_TOKEN_SOURCE,
  FORMULA_RIBBON_FUNCTION_MENUS,
  INSERT_FUNCTION_RECENT_NAMES,
  MAX_FUNCTION_ARGUMENT_ROWS,
  MAX_RECENT_FORMULA_FUNCTIONS,
  state,
} from "../../app-core.js";
import {
  closeDataDialog,
} from "../drawing/charts-and-insertion.js";
import {
  positiveModulo,
  rangesEqualLoose,
} from "../editing/autofill.js";
import {
  isElementImeCompositionEvent,
} from "../editing/keyboard-and-paste.js";
import {
  activateSheetForFormulaEditResult,
  activeFormulaReferenceEntriesForActiveSheet,
  activeSelectionRange,
  activeSelectionRanges,
  beginFormulaBarEdit,
  cellPointFromClientPosition,
  cellPointFromElement,
  cellPointFromPointerEvent,
  commitCellValue,
  getDisplayForCell,
  isCellInRange,
  isSingleCellRange,
  normalizeRange,
  rangeToLabel,
  selectionAutoScrollVelocity,
  selectionRangeFromPoints,
  syncFormulaEditPreview,
  updateSelectionFrames,
  updateSelectionUi,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  clearCellEditSelection,
  columnSelectionRange,
  editablePlainText,
  ensureCellElementForPoint,
  placeCaretAtEnd,
  rowSelectionRange,
  selectionToolbarPoint,
  stopEditAutoScroll,
  syncEditableCellAlignment,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  internCellsInRange,
  pushHistory,
  restoreSheetFromHistory,
  storeCellModelForSheet,
} from "../editing/structure-and-history.js";
import {
  formulaCaretIsInsideString,
  formulaDefaultSyntax,
  formulaInsertFunctionDefinitionByName,
  formulaInsertFunctionDefinitions,
  formulaReferenceColorClass,
  formulaTextStartsFormula,
  insertFunctionCategoryOptions,
  updateFormulaAutocompleteForEditor,
} from "./autocomplete.js";
import {
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
  closeHomeCommandMenu,
  openHomeCommandMenu,
} from "../ui/printing-and-home-tools.js";
import {
  applyAutoSumToSelection,
  ensureSheetSize,
  formulaFunctionByName,
  quoteSheetNameForFormula,
  scheduleViewportRender,
} from "../ui/ribbon-and-links.js";
import {
  ensureEditableRangeForProtection,
} from "../ui/sheet-protection-and-translation.js";
import {
  openCreateNamesDialog,
} from "../ui/views-and-formula-tools.js";
import {
  buildFormulaEngine,
  formulaEngineSheetIndex,
  translateFormulaForCurrentFormulaEngine,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  activeSheet,
  displayValue,
  forceSyncHorizontalScrollbarFromHost,
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
  firstDefinedNameReference,
  validationReferenceFromFormula,
} from "../workbook/tables-and-validation.js";
import {
  commandMenuAccessKeyAttr,
} from "../workbook/views-and-macros.js";


function closeFormulaAutocomplete(options = {}) {
  const editor = state.formulaAutocomplete?.editor;
  if (options.rememberClosedState && editor?.isConnected) {
    state.formulaAutocompleteAccepted = {
      editor,
      text: formulaEditorText(editor),
      caret: formulaEditorCaretOffset(editor),
    };
  }
  state.formulaAutocomplete = null;
  document.getElementById("formulaAutocomplete")?.remove();
  if (options.restoreFocus !== false && editor?.isConnected) editor.focus({ preventScroll: true });
}



function closeFormulaAutocompleteOnEscape(event) {
  if (!state.formulaAutocomplete || event.key !== "Escape") return;
  event.preventDefault();
  event.stopPropagation();
  closeFormulaAutocomplete({ restoreFocus: false, rememberClosedState: true });
}



function closeFormulaAutocompleteOnOutsidePointer(event) {
  if (!state.formulaAutocomplete) return;
  const popup = document.getElementById("formulaAutocomplete");
  const editor = state.formulaAutocomplete.editor;
  const target = event.target;
  if (popup?.contains(target) || editor?.contains?.(target) || editor === target) return;
  closeFormulaAutocomplete({ restoreFocus: false, rememberClosedState: true });
}



function formulaEditorText(editor) {
  if (!editor) return "";
  if (editor === $formulaInput[0] || editor.tagName === "TEXTAREA" || editor.tagName === "INPUT") {
    return editor.value || "";
  }
  return editablePlainText(editor);
}



function formulaEditorCaretOffset(editor) {
  if (!editor) return 0;
  if (editor === $formulaInput[0] || editor.tagName === "TEXTAREA" || editor.tagName === "INPUT") {
    return Number(editor.selectionStart ?? editor.value?.length ?? 0);
  }
  const selection = window.getSelection?.();
  if (!selection?.rangeCount) return (editor.textContent || "").length;
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.endContainer)) return (editor.textContent || "").length;
  const before = range.cloneRange();
  before.selectNodeContents(editor);
  before.setEnd(range.endContainer, range.endOffset);
  return before.toString().length;
}



function formulaEditorSetTextAndCaret(editor, text, caretOffset) {
  if (!editor) return;
  const value = String(text ?? "");
  const caret = Math.max(0, Math.min(value.length, Number(caretOffset) || 0));
  if (editor === $formulaInput[0] || editor.tagName === "TEXTAREA" || editor.tagName === "INPUT") {
    editor.value = value;
    editor.setSelectionRange?.(caret, caret);
    editor.dispatchEvent(new Event("input", { bubbles: true }));
    return;
  }
  editor.textContent = value;
  placeCaretAtTextOffset(editor, caret);
  editor.dispatchEvent(new Event("input", { bubbles: true }));
}



function formulaEditorSelectionOffsets(editor) {
  if (!editor) return { start: 0, end: 0 };
  const textLength = formulaEditorText(editor).length;
  if (editor === $formulaInput[0] || editor.tagName === "TEXTAREA" || editor.tagName === "INPUT") {
    const start = Number.isFinite(Number(editor.selectionStart)) ? Number(editor.selectionStart) : textLength;
    const end = Number.isFinite(Number(editor.selectionEnd)) ? Number(editor.selectionEnd) : start;
    return {
      start: Math.max(0, Math.min(textLength, Math.min(start, end))),
      end: Math.max(0, Math.min(textLength, Math.max(start, end))),
    };
  }

  const selection = window.getSelection?.();
  if (!selection?.rangeCount) return { start: textLength, end: textLength };
  const range = selection.getRangeAt(0);
  if (!editor.contains(range.startContainer) || !editor.contains(range.endContainer)) {
    return { start: textLength, end: textLength };
  }
  const start = formulaEditorTextOffsetForBoundary(editor, range.startContainer, range.startOffset);
  const end = formulaEditorTextOffsetForBoundary(editor, range.endContainer, range.endOffset);
  return {
    start: Math.max(0, Math.min(textLength, Math.min(start, end))),
    end: Math.max(0, Math.min(textLength, Math.max(start, end))),
  };
}



function formulaEditorTextOffsetForBoundary(editor, container, offset) {
  const range = document.createRange();
  range.selectNodeContents(editor);
  try {
    range.setEnd(container, offset);
    return range.toString().length;
  } catch {
    return formulaEditorText(editor).length;
  }
}



function markFormulaReferenceSheetTabPointerDownCapture(event) {
  if (event.button !== 0 || event.shiftKey || event.ctrlKey || event.metaKey) return;
  const tab = event.target?.closest?.(".sheet-tab");
  if (!tab || tab.closest?.(".sheet-tab-rename-input")) return;
  const editor = document.activeElement?.closest?.(".sheet-cell[contenteditable='true']");
  if (!editor) return;
  const text = editablePlainText(editor);
  if (!formulaTextStartsFormula(text)) return;
  const index = Number(tab.dataset.sheetIndex);
  if (!Number.isInteger(index)) return;
  state.pendingFormulaReferenceSheetTabIndex = index;
  if (state.pendingFormulaReferenceSheetTabTimer) {
    window.clearTimeout(state.pendingFormulaReferenceSheetTabTimer);
  }
  state.pendingFormulaReferenceSheetTabTimer = window.setTimeout(() => {
    state.pendingFormulaReferenceSheetTabIndex = null;
    state.pendingFormulaReferenceSheetTabTimer = null;
  }, 0);
}



function clearPendingFormulaReferenceSheetTab() {
  if (state.pendingFormulaReferenceSheetTabTimer) {
    window.clearTimeout(state.pendingFormulaReferenceSheetTabTimer);
  }
  state.pendingFormulaReferenceSheetTabIndex = null;
  state.pendingFormulaReferenceSheetTabTimer = null;
}



function cellEditBlurStartsFormulaSheetNavigation(event, text) {
  const relatedTarget = event.relatedTarget || event.originalEvent?.relatedTarget || null;
  return formulaTextStartsFormula(text) && (
    Number.isInteger(state.pendingFormulaReferenceSheetTabIndex) ||
    Boolean(relatedTarget?.closest?.(".sheet-tab"))
  );
}



function preserveCellFormulaEditForSheetNavigation(cellElement, row, col, before, text) {
  const point = cellPointFromElement(cellElement) || { sheetIndex: state.activeSheetIndex, row, col };
  state.formulaEdit = {
    sheetIndex: point.sheetIndex,
    row,
    col,
    before,
  };
  state.selected = { sheetIndex: point.sheetIndex, row, col };
  state.selectionRange = selectionRangeFromPoints(state.selected, state.selected);
  state.selectionRanges = [state.selectionRange];
  state.selectionAnchor = state.selected;
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  closeFormulaAutocomplete({ restoreFocus: false });
  const $cell = $(cellElement);
  cellElement.removeAttribute("contenteditable");
  cellElement.classList.remove("is-formula-editing");
  $cell.removeData("beforeEdit");
  $cell.removeData("beforeRichText");
  $cell.removeData("cancelEdit");
  clearCellEditSelection(cellElement);
  stopEditAutoScroll();
  const value = String(text ?? "");
  $formulaInput.val(value);
  const input = $formulaInput[0];
  if (input) input.setSelectionRange?.(value.length, value.length);
  clearPendingFormulaReferenceSheetTab();
}



function activeFormulaReferenceEditor() {
  const formulaInput = $formulaInput[0];
  if (document.activeElement === formulaInput) return formulaInput;
  if (state.formulaEdit && formulaInput && formulaTextStartsFormula(String($formulaInput.val() ?? ""))) {
    return formulaInput;
  }
  const argumentInput = activeFunctionArgumentReferenceInput();
  if (argumentInput) return argumentInput;
  const editable = document.activeElement?.closest?.(".sheet-cell[contenteditable='true']");
  if (editable && editable.isConnected) return editable;
  return null;
}



function activeFunctionArgumentReferenceInput() {
  if (state.dataDialog?.type !== "function-arguments") return null;
  const active = document.activeElement?.closest?.("[data-formula-argument-input]");
  if (active?.isConnected) return active;
  const index = Number(state.dataDialog?.activeArgumentIndex);
  if (Number.isFinite(index)) {
    const input = document.querySelector(`[data-data-dialog='function-arguments'] [data-formula-argument-input][data-formula-argument-index="${index}"]`);
    if (input?.isConnected) return input;
  }
  return null;
}



function isFunctionArgumentReferenceEditor(editor) {
  return Boolean(editor?.matches?.("[data-formula-argument-input]") && editor.closest?.("[data-data-dialog='function-arguments']"));
}



function formulaReferenceEditorSheetIndex(editor) {
  if (!editor) return state.activeSheetIndex;
  if (isFunctionArgumentReferenceEditor(editor)) {
    return Number(state.dataDialog?.editorSource?.sheetIndex ?? state.dataDialog?.sourceSheetIndex ?? state.activeSheetIndex);
  }
  if (editor === $formulaInput[0]) {
    const editorSourceSheetIndex = Number(editor.dataset.formulaReferenceSourceSheetIndex);
    if (Number.isInteger(editorSourceSheetIndex)) return editorSourceSheetIndex;
    const navigationSourceSheetIndex = state.formulaReferenceSheetNavigation?.sourceSheetIndex;
    return Number.isInteger(navigationSourceSheetIndex)
      ? navigationSourceSheetIndex
      : state.formulaEdit?.sheetIndex ?? state.selected?.sheetIndex ?? state.activeSheetIndex;
  }
  const point = cellPointFromElement(editor);
  return point?.sheetIndex ?? state.activeSheetIndex;
}



function formulaReferenceLabel(range, editor) {
  if (!range) return "";
  const label = rangeToLabel(range);
  const editSheetIndex = formulaReferenceEditorSheetIndex(editor);
  if (range.sheetIndex === editSheetIndex) return label;
  const sheet = state.model?.sheets?.[range.sheetIndex];
  return sheet ? `${quoteSheetNameForFormula(sheet.name)}!${label}` : label;
}



function clearFormulaReferenceState(options = {}) {
  stopFormulaReferenceAutoScroll();
  state.formulaReferenceDrag = null;
  state.formulaReferenceRange = null;
  state.formulaReferenceSheetNavigation = null;
  const formulaInput = $formulaInput[0];
  if (formulaInput) delete formulaInput.dataset.formulaReferenceSourceSheetIndex;
  if (options.update !== false) updateFormulaReferenceUi();
}



function startFormulaReferenceSelection(event) {
  const context = formulaReferenceInsertionContext(event);
  if (!context) return false;
  const { editor, text, bounds } = context;
  const cellEditEditor = editor.closest?.(".sheet-cell[contenteditable='true']");
  if (cellEditEditor && event.target?.closest?.(".sheet-cell[contenteditable='true']") === cellEditEditor) return false;

  const point = cellPointFromPointerEvent(event, { allowEditable: editor === $formulaInput[0] });
  if (!point) return false;
  const range = selectionRangeFromPoints(point, point);
  if (!range) return false;

  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  closeFormulaAutocomplete({ restoreFocus: false });

  state.formulaReferenceDrag = {
    editor,
    kind: "cell",
    pointerId: event.pointerId,
    anchor: point,
    labelRange: singleCellRangeFromPoint(point),
    prefix: text.slice(0, bounds.start),
    suffix: text.slice(bounds.end),
    clientX: event.clientX,
    clientY: event.clientY,
    frame: null,
  };
  event.currentTarget.setPointerCapture?.(event.pointerId);
  applyFormulaReferenceDragRange(range);
  editor.focus({ preventScroll: true });
  return true;
}



function singleCellRangeFromPoint(point) {
  if (!point) return null;
  return {
    sheetIndex: point.sheetIndex ?? state.activeSheetIndex,
    top: point.row,
    left: point.col,
    bottom: point.row,
    right: point.col,
  };
}



function formulaReferenceInsertionContext(event = null) {
  const editor = activeFormulaReferenceEditor();
  if (!editor || !state.model) return null;
  const text = formulaEditorText(editor);
  const isArgumentInput = isFunctionArgumentReferenceEditor(editor);
  const caret = formulaEditorCaretOffset(editor);
  if (isArgumentInput) {
    const selection = formulaEditorSelectionOffsets(editor);
    const token = formulaReferenceTokenAt(text, selection.start, selection.end);
    const bounds = token
      ? { start: token.start, end: token.end }
      : { start: selection.start, end: selection.end };
    return { editor, text, bounds, event };
  }
  if (!formulaTextStartsFormula(text) || formulaCaretIsInsideString(text, caret)) return null;
  const selection = formulaEditorSelectionOffsets(editor);
  const token = formulaReferenceTokenAt(text, selection.start, selection.end);
  const bounds = token
    ? { start: token.start, end: token.end }
    : { start: selection.start, end: selection.end };
  if (formulaCaretIsInsideString(text, bounds.start)) return null;
  return { editor, text, bounds, event };
}



function startFormulaHeaderReferenceSelection(kind, index, event) {
  const context = formulaReferenceInsertionContext(event);
  if (!context) return false;
  const range = formulaHeaderReferenceRange(kind, index, index);
  if (!range) return false;

  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  closeFormulaAutocomplete({ restoreFocus: false });

  state.formulaReferenceDrag = {
    editor: context.editor,
    kind,
    pointerId: event.pointerId,
    anchorIndex: index,
    prefix: context.text.slice(0, context.bounds.start),
    suffix: context.text.slice(context.bounds.end),
    clientX: event.clientX,
    clientY: event.clientY,
    frame: null,
  };
  applyFormulaReferenceDragRange(range);
  context.editor.focus({ preventScroll: true });
  return true;
}



function formulaHeaderReferenceRange(kind, startIndex, endIndex) {
  if (!state.model) return null;
  const start = Math.max(1, Number(startIndex) || 1);
  const end = Math.max(1, Number(endIndex) || start);
  return kind === "col" ? columnSelectionRange(start, end) : rowSelectionRange(start, end);
}



function extendFormulaHeaderReferenceSelection(event) {
  const drag = state.formulaReferenceDrag;
  if (!drag || (drag.kind !== "col" && drag.kind !== "row")) return false;
  if (drag.pointerId != null && event.pointerId != null && event.pointerId !== drag.pointerId) return true;
  const index = Number(drag.kind === "col" ? event.currentTarget.dataset.col : event.currentTarget.dataset.row);
  if (!index) return true;
  event.preventDefault?.();
  event.stopPropagation?.();
  drag.clientX = event.clientX;
  drag.clientY = event.clientY;
  const range = formulaHeaderReferenceRange(drag.kind, drag.anchorIndex, index);
  if (range) applyFormulaReferenceDragRange(range);
  return true;
}



function updateFormulaReferenceDrag(event) {
  const drag = state.formulaReferenceDrag;
  if (!drag) return;
  if (drag.pointerId != null && event.pointerId != null && event.pointerId !== drag.pointerId) return;
  if (drag.kind === "col" || drag.kind === "row") return;
  if (event.buttons !== undefined && (event.buttons & 1) !== 1) {
    finishFormulaReferenceDrag(event);
    return;
  }
  event.preventDefault?.();
  drag.clientX = event.clientX;
  drag.clientY = event.clientY;
  updateFormulaReferenceFromClient(event.clientX, event.clientY, { clampToGrid: true });
  updateFormulaReferenceAutoScroll(event.clientX, event.clientY);
}



function finishFormulaReferenceDrag(event = null) {
  const drag = state.formulaReferenceDrag;
  if (!drag) return;
  if (event && drag.pointerId != null && event.pointerId != null && event.pointerId !== drag.pointerId) return;
  const editor = drag.editor;
  stopFormulaReferenceAutoScroll();
  state.formulaReferenceDrag = null;
  state.suppressNextCellClick = true;
  if (editor?.isConnected) {
    editor.focus({ preventScroll: true });
    if (isFunctionArgumentReferenceEditor(editor)) {
      syncFunctionArgumentsDialog(editor.closest("[data-data-dialog='function-arguments']"));
    } else {
      updateFormulaAutocompleteForEditor(editor);
    }
  }
}



function updateFormulaReferenceFromClient(clientX, clientY, options = {}) {
  const drag = state.formulaReferenceDrag;
  if (!drag) return false;
  const endpoint = cellPointFromClientPosition(clientX, clientY, options);
  if (!endpoint) return false;
  const range = selectionRangeFromPoints(drag.anchor, endpoint);
  if (!range) return false;
  const labelRange = pointsEqualLoose(drag.anchor, endpoint) ? singleCellRangeFromPoint(drag.anchor) : range;
  if (rangesEqualLoose(range, state.formulaReferenceRange) && rangesEqualLoose(labelRange, drag.labelRange)) return true;
  drag.labelRange = labelRange;
  applyFormulaReferenceDragRange(range);
  return true;
}



function applyFormulaReferenceDragRange(range) {
  const drag = state.formulaReferenceDrag;
  if (!drag?.editor || !range) return;
  state.formulaReferenceRange = { ...range };
  const reference = formulaReferenceLabel(drag.labelRange || range, drag.editor);
  formulaEditorSetTextAndCaret(drag.editor, `${drag.prefix}${reference}${drag.suffix}`, drag.prefix.length + reference.length);
  if (isFunctionArgumentReferenceEditor(drag.editor)) {
    syncFunctionArgumentsDialog(drag.editor.closest("[data-data-dialog='function-arguments']"));
  }
  updateFormulaReferenceUi();
}



function pointsEqualLoose(a, b) {
  return Boolean(a && b && a.sheetIndex === b.sheetIndex && a.row === b.row && a.col === b.col);
}



function updateFormulaReferenceUi() {
  const colorClasses = FORMULA_REFERENCE_COLORS.map((_color, index) => formulaReferenceColorClass(index)).join(" ");
  $sheetHost
    .find(".sheet-cell.is-formula-reference")
    .removeClass(`is-formula-reference ${colorClasses}`);
  const entries = activeFormulaReferenceEntriesForActiveSheet();
  if (entries.length) {
    $sheetHost.find(".sheet-cell[data-row][data-col]").each((_index, element) => {
      const row = Number(element.dataset.row);
      const col = Number(element.dataset.col);
      const entry = entries.find((item) => isCellInRange(item, row, col));
      if (!entry) return;
      element.classList.add("is-formula-reference", formulaReferenceColorClass(entry.colorIndex));
    });
  }
  updateSelectionFrames();
}



function updateFormulaReferenceAutoScroll(clientX, clientY) {
  const drag = state.formulaReferenceDrag;
  if (!drag) {
    stopFormulaReferenceAutoScroll();
    return;
  }
  const velocity = selectionAutoScrollVelocity(clientX, clientY);
  if (!velocity.x && !velocity.y) {
    stopFormulaReferenceAutoScroll();
    return;
  }
  drag.clientX = clientX;
  drag.clientY = clientY;
  if (!drag.frame) {
    drag.frame = requestAnimationFrame(runFormulaReferenceAutoScroll);
  }
}



function runFormulaReferenceAutoScroll() {
  const drag = state.formulaReferenceDrag;
  if (!drag) return;
  drag.frame = null;
  const host = $sheetHost[0];
  if (!host) {
    stopFormulaReferenceAutoScroll();
    return;
  }
  const velocity = selectionAutoScrollVelocity(drag.clientX, drag.clientY);
  if (!velocity.x && !velocity.y) {
    stopFormulaReferenceAutoScroll();
    return;
  }
  const previousTop = host.scrollTop;
  const previousLeft = host.scrollLeft;
  const maxTop = Math.max(0, host.scrollHeight - host.clientHeight);
  const maxLeft = Math.max(0, host.scrollWidth - host.clientWidth);
  host.scrollTop = Math.max(0, Math.min(maxTop, previousTop + velocity.y));
  host.scrollLeft = Math.max(0, Math.min(maxLeft, previousLeft + velocity.x));
  if (host.scrollTop !== previousTop || host.scrollLeft !== previousLeft) {
    forceSyncHorizontalScrollbarFromHost();
    scheduleViewportRender({ deferTrim: true });
  }
  updateFormulaReferenceFromClient(drag.clientX, drag.clientY, { clampToGrid: true });
  drag.frame = requestAnimationFrame(runFormulaReferenceAutoScroll);
}



function stopFormulaReferenceAutoScroll() {
  const frame = state.formulaReferenceDrag?.frame;
  if (frame) cancelAnimationFrame(frame);
  if (state.formulaReferenceDrag) state.formulaReferenceDrag.frame = null;
}



function placeCaretAtTextOffset(element, offset) {
  const selection = window.getSelection?.();
  if (!selection || !element) return;
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  let remaining = Math.max(0, Number(offset) || 0);
  let node = walker.nextNode();
  while (node) {
    const length = node.textContent?.length || 0;
    if (remaining <= length) {
      const range = document.createRange();
      range.setStart(node, remaining);
      range.collapse(true);
      selection.removeAllRanges();
      selection.addRange(range);
      return;
    }
    remaining -= length;
    node = walker.nextNode();
  }
  placeCaretAtEnd(element);
}



function formulaTemplateReference() {
  const range = activeSelectionRanges()[0] || activeSelectionRange();
  if (!range) return "A1";
  if (isSingleCellRange(range)) {
    const row = Math.max(1, range.top - 1);
    return row >= 1 ? `${columnName(range.left)}${row}:${columnName(range.left)}${range.top}` : `${columnName(range.left)}${range.top}`;
  }
  return rangeToLabel(range);
}



function insertFunctionIntoSelection(functionName) {
  const point = selectionToolbarPoint();
  if (!point) {
    showMessageBox("関数を入力するセルを選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const definition = formulaFunctionByName(functionName);
  const formula = definition.template(formulaTemplateReference());
  commitCellValue(point.row, point.col, formula, { skipValidation: true });
  recordFormulaFunctionUse(definition.name);
  setStatus(`${columnName(point.col)}${point.row} に ${definition.label} 関数を入力しました。`);
}



function openFunctionArgumentsDialog(functionName, options = {}) {
  const point = selectionToolbarPoint();
  if (!point && !options.editorSource) {
    showMessageBox("関数を入力するセルを選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const definition = formulaFunctionByName(functionName);
  const specs = formulaFunctionArgumentSpecs(definition);
  const argumentValues = Array.isArray(options.argumentValues)
    ? options.argumentValues
    : specs.map((spec, index) => formulaArgumentDefaultValue(definition.name, spec, index));
  state.dataDialog = {
    type: "function-arguments",
    functionName: definition.name,
    argumentValues,
    editorSource: options.editorSource || null,
    sourcePoint: point ? { sheetIndex: point.sheetIndex, row: point.row, col: point.col } : null,
    sourceSheetIndex: point?.sheetIndex ?? options.editorSource?.sheetIndex ?? state.activeSheetIndex,
  };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay is-reference-pickable";
  overlay.id = "dataDialog";
  overlay.innerHTML = functionArgumentsDialogHtml(definition, specs, state.dataDialog.argumentValues, {
    editorSource: Boolean(options.editorSource),
  });
  document.body.appendChild(overlay);
  syncFunctionArgumentsDialog(overlay.querySelector("[data-data-dialog='function-arguments']"));
  overlay.querySelector("[data-formula-argument-input]")?.focus({ preventScroll: true });
}



function functionArgumentsDialogHtml(definition, specs, values, options = {}) {
  const rows = functionArgumentRowsHtml(functionArgumentDisplaySpecs(definition, values), values);
  return `
    <div class="excel-mini-dialog data-dialog function-arguments-dialog" role="dialog" aria-modal="true" aria-labelledby="functionArgumentsTitle" data-data-dialog="function-arguments">
      <div class="excel-mini-titlebar" id="functionArgumentsTitle">関数の引数</div>
      <div class="excel-mini-body data-dialog-body formula-dialog-body">
        <div class="function-argument-heading">
          <strong>${escapeHtml(definition.label || definition.name)}</strong>
          <span>${escapeHtml(definition.description || "")}</span>
        </div>
        <div class="function-argument-list">
          ${rows}
        </div>
        <div class="function-argument-preview">
          <span>数式:</span>
          <code data-function-argument-formula></code>
        </div>
        <div class="function-argument-result">
          <span>数式の結果:</span>
          <strong data-function-argument-result></strong>
        </div>
        <div class="function-argument-detail" id="functionArgumentDescription" data-function-argument-description></div>
        <div class="function-argument-syntax">
          ${escapeHtml(definition.syntax || formulaDefaultSyntax(definition.name))}
        </div>
      </div>
      <div class="excel-mini-actions">
        ${options.editorSource ? "" : '<button class="excel-mini-button" type="button" data-data-dialog-action="function-arguments-back">戻る(B)</button>'}
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="function-arguments-ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>
  `;
}



function functionArgumentRowsHtml(specs, values = []) {
  return specs.length
    ? specs.map((spec, index) => functionArgumentRowHtml(spec, index, values[index] || "")).join("")
    : `<div class="function-argument-empty">この関数には引数はありません。</div>`;
}



function functionArgumentRowHtml(spec, index, value) {
  const inputId = `functionArgumentInput${index}`;
  return `
    <div class="function-argument-row" data-function-argument-row="${index}">
      <label class="function-argument-name" for="${escapeAttr(inputId)}" title="${escapeAttr(spec.label)}">${escapeHtml(spec.label)}${spec.optional ? "" : ":"}</label>
      <input id="${escapeAttr(inputId)}" type="text" data-formula-argument-input data-formula-argument-index="${index}" data-formula-argument-label="${escapeAttr(spec.label)}" value="${escapeAttr(value || "")}" aria-label="${escapeAttr(spec.label)}" aria-describedby="functionArgumentDescription">
      <button class="function-argument-reference-button" type="button" data-data-dialog-action="function-argument-reference-toggle" data-formula-argument-index="${index}" title="参照範囲を選択" aria-label="${escapeAttr(spec.label)} の参照範囲を選択" aria-expanded="true">↗</button>
      <span class="function-argument-value" data-function-argument-value="${index}"></span>
      <span class="function-argument-required">${spec.optional ? "省略可" : "必須"}</span>
    </div>`;
}



function formulaFunctionArgumentSpecs(definition) {
  const syntax = String(definition?.syntax || formulaDefaultSyntax(definition?.name || "SUM"));
  const start = syntax.indexOf("(");
  const end = syntax.lastIndexOf(")");
  if (start < 0 || end <= start) return [];
  const body = syntax.slice(start + 1, end).trim();
  if (!body) return [];
  const specs = [];
  splitFunctionArgumentSyntax(body)
    .forEach((token) => {
      const source = String(token || "").trim();
      if (/^\.\.\.$/u.test(source)) {
        if (specs.length) specs[specs.length - 1].repeat = true;
        return;
      }
      const optional = /^\[.*\]$/u.test(source);
      const repeat = /\.\.\.\s*$/u.test(source);
      const label = source
        .replace(/^\[/u, "")
        .replace(/\]$/u, "")
        .replace(/\s*\.\.\.\s*$/u, "")
        .trim();
      if (label && label !== "...") specs.push({ label, optional, repeat });
    });
  return specs.slice(0, MAX_FUNCTION_ARGUMENT_ROWS);
}



function functionArgumentDisplaySpecs(definition, values = [], options = {}) {
  const specs = formulaFunctionArgumentSpecs(definition);
  if (!specs.length) return specs;
  const repeatIndex = specs.findIndex((spec) => spec.repeat);
  if (repeatIndex < 0) return specs;
  const repeatStart = repeatArgumentGroupStart(specs, repeatIndex);
  const repeatGroup = specs.slice(repeatStart, repeatIndex + 1);
  let desired = specs.length;
  while (desired < MAX_FUNCTION_ARGUMENT_ROWS) {
    const groupStart = desired - repeatGroup.length;
    const groupHasValue = repeatGroup.some((_spec, offset) => String(values[groupStart + offset] || "").trim());
    if (!groupHasValue) break;
    desired += repeatGroup.length;
  }
  if (values.length > desired) desired = Math.min(MAX_FUNCTION_ARGUMENT_ROWS, values.length);
  if (options.minRows) desired = Math.max(desired, Number(options.minRows) || 0);
  desired = Math.max(specs.length, Math.min(MAX_FUNCTION_ARGUMENT_ROWS, desired));
  const output = specs.slice();
  for (let index = output.length; index < desired; index += 1) {
    const groupOffset = index - specs.length;
    const groupIndex = positiveModulo(groupOffset, repeatGroup.length);
    const sequenceOffset = Math.floor(groupOffset / repeatGroup.length) + 1;
    output.push(repeatedFunctionArgumentSpec(repeatGroup[groupIndex], sequenceOffset));
  }
  return output;
}



function repeatArgumentGroupStart(specs, repeatIndex) {
  const repeatNumber = functionArgumentLabelSequence(specs[repeatIndex]?.label);
  if (repeatNumber !== 1) return repeatIndex;
  let start = repeatIndex;
  while (start > 0 && functionArgumentLabelSequence(specs[start - 1]?.label) === repeatNumber) {
    start -= 1;
  }
  return start;
}



function functionArgumentLabelSequence(label) {
  const match = String(label || "").match(/(\d+)$/u);
  return match ? Number(match[1]) : null;
}



function repeatedFunctionArgumentSpec(spec, offset) {
  return {
    ...spec,
    label: nextFunctionArgumentLabel(spec.label, offset),
    optional: true,
    repeat: true,
  };
}



function nextFunctionArgumentLabel(label, offset) {
  const text = String(label || "argument");
  const match = text.match(/^(.*?)(\d+)$/u);
  if (!match) return `${text}${Number(offset) + 1}`;
  return `${match[1]}${Number(match[2]) + Number(offset)}`;
}



function splitFunctionArgumentSyntax(body) {
  const parts = [];
  let current = "";
  let depth = 0;
  let quote = "";
  for (const char of String(body || "")) {
    if (quote) {
      current += char;
      if (char === quote) quote = "";
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      current += char;
      continue;
    }
    if (char === "[" || char === "(") depth += 1;
    if (char === "]" || char === ")") depth = Math.max(0, depth - 1);
    if (char === "," && depth === 0) {
      parts.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  if (current.trim()) parts.push(current);
  return parts;
}



function formulaArgumentDefaultValue(functionName, spec, index) {
  const name = String(functionName || "").toUpperCase();
  const override = FORMULA_ARGUMENT_DEFAULT_OVERRIDES[name]?.[index];
  if (typeof override === "function") return override();
  if (override != null) return String(override);
  if (spec?.optional) return "";
  const label = String(spec?.label || "").toLowerCase();
  if (/logical|ignore|range_lookup|a1/.test(label)) return "1=1";
  if (/criteria|delimiter|format|text|friendly|link_location/.test(label)) return '""';
  if (/num|count|index|row|column|month|day|year|digits|type|mode|order|basis|rate|nper|pv|fv|pmt|guess|bottom|top/.test(label)) return "1";
  if (/array|range|ref|value|number|lookup|table|return|sum/.test(label)) return formulaTemplateReference();
  return "1";
}



function handleFunctionArgumentInput(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='function-arguments']");
  const index = Number(event.currentTarget.dataset.formulaArgumentIndex);
  if (Number.isFinite(index)) state.dataDialog = { ...(state.dataDialog || {}), activeArgumentIndex: index };
  if (dialog) syncFunctionArgumentsDialog(dialog);
}



function handleFunctionArgumentKeydown(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='function-arguments']");
  if (!dialog) return;
  const isComposing = isElementImeCompositionEvent(event, event.currentTarget);
  if (event.key === "F4" && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey) {
    if (cycleFormulaReferenceInFunctionArgumentInput(event.currentTarget)) {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation?.();
      return;
    }
  }
  if (state.dataDialog?.referenceCollapsed && (event.key === "Enter" || event.key === "Escape") && !isComposing) {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    expandFunctionArgumentReferencePicker(dialog);
    return;
  }
  if (event.key === "Enter" && !isComposing) {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    applyFunctionArgumentsDialog(dialog);
    return;
  }
  if (isComposing) return;
  if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation?.();
    closeDataDialog();
    return;
  }
  if (event.key !== "Tab") return;
  const currentIndex = Number(event.currentTarget.dataset.formulaArgumentIndex);
  syncFunctionArgumentsDialog(dialog);
  const inputs = [...dialog.querySelectorAll("[data-formula-argument-input]")];
  const index = Number.isFinite(currentIndex) ? currentIndex : inputs.indexOf(event.currentTarget);
  const nextIndex = index + (event.shiftKey ? -1 : 1);
  const next = inputs[nextIndex] || dialog.querySelector(`[data-formula-argument-input][data-formula-argument-index="${nextIndex}"]`);
  if (!next) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  next.focus({ preventScroll: true });
  next.select?.();
  syncFunctionArgumentsDialog(dialog);
}



function syncFunctionArgumentsDialog(dialog) {
  if (!dialog) return;
  const definition = formulaFunctionByName(state.dataDialog?.functionName || "SUM");
  syncFunctionArgumentRows(dialog, definition);
  syncFunctionArgumentReferencePicker(dialog);
  const activeControl =
    dialog.querySelector("[data-formula-argument-input]:focus") ||
    dialog.querySelector("[data-data-dialog-action='function-argument-reference-toggle']:focus");
  const activeArgumentIndex = Number(
    activeControl?.dataset.formulaArgumentIndex ??
    state.dataDialog?.activeArgumentIndex ??
    0,
  );
  dialog.querySelectorAll("[data-function-argument-row]").forEach((row) => {
    const rowIndex = Number(row.dataset.functionArgumentRow);
    row.classList.toggle("is-active", Number.isFinite(activeArgumentIndex) && rowIndex === activeArgumentIndex);
  });
  const formula = buildFunctionArgumentsFormula(dialog);
  const preview = dialog.querySelector("[data-function-argument-formula]");
  if (preview) preview.textContent = formula;
  if (state.dataDialog?.type === "function-arguments") {
    state.dataDialog.argumentValues = [...dialog.querySelectorAll("[data-formula-argument-input]")].map((input) => String(input.value || ""));
  }
  dialog.querySelectorAll("[data-formula-argument-input]").forEach((input) => {
    const index = Number(input.dataset.formulaArgumentIndex);
    const output = dialog.querySelector(`[data-function-argument-value="${index}"]`);
    if (!output) return;
    const value = String(input.value || "").trim();
    output.textContent = value ? `= ${evaluateFormulaPreviewExpression(value)}` : "";
  });
  const result = dialog.querySelector("[data-function-argument-result]");
  if (result) result.textContent = evaluateFormulaPreviewExpression(formula);
  const description = dialog.querySelector("[data-function-argument-description]");
  if (description) {
    const index = Number.isFinite(activeArgumentIndex) ? activeArgumentIndex : 0;
    const specs = functionArgumentDisplaySpecs(definition, state.dataDialog?.argumentValues || []);
    const spec = specs[index];
    const text = formulaArgumentDescription(definition, spec, index);
    description.innerHTML = text
      ? `<span class="function-argument-detail-argument">引数 ${index + 1}: </span>${escapeHtml(text)}`
      : "";
  }
}



function toggleFunctionArgumentReferencePicker(dialog, button) {
  if (!dialog || state.dataDialog?.type !== "function-arguments") return;
  const index = Number(button?.dataset.formulaArgumentIndex);
  if (!Number.isFinite(index)) return;
  if (state.dataDialog.referenceCollapsed && Number(state.dataDialog.referenceCollapsedIndex) === index) {
    expandFunctionArgumentReferencePicker(dialog);
    return;
  }
  collapseFunctionArgumentReferencePicker(dialog, index);
}



function collapseFunctionArgumentReferencePicker(dialog, index) {
  if (!dialog) return;
  const inputSelector = `[data-formula-argument-input][data-formula-argument-index="${index}"]`;
  if (!dialog.querySelector(inputSelector)) return;
  state.dataDialog = {
    ...(state.dataDialog || {}),
    activeArgumentIndex: index,
    referenceCollapsed: true,
    referenceCollapsedIndex: index,
  };
  syncFunctionArgumentsDialog(dialog);
  const input = dialog.querySelector(inputSelector);
  input.focus({ preventScroll: true });
  input.select?.();
  setStatus("参照するセルまたはセル範囲を選択してください。");
}



function expandFunctionArgumentReferencePicker(dialog) {
  if (state.dataDialog?.type !== "function-arguments") return;
  const index = Number(state.dataDialog.referenceCollapsedIndex ?? state.dataDialog.activeArgumentIndex);
  state.dataDialog = {
    ...(state.dataDialog || {}),
    referenceCollapsed: false,
    referenceCollapsedIndex: null,
  };
  syncFunctionArgumentsDialog(dialog);
  const input = dialog?.querySelector(`[data-formula-argument-input][data-formula-argument-index="${Number.isFinite(index) ? index : 0}"]`);
  if (input) {
    input.focus({ preventScroll: true });
    const end = String(input.value || "").length;
    input.setSelectionRange?.(end, end);
  }
}



function syncFunctionArgumentReferencePicker(dialog) {
  if (!dialog) return;
  const collapsed = state.dataDialog?.referenceCollapsed === true;
  const activeIndex = Number(state.dataDialog?.referenceCollapsedIndex);
  dialog.classList.toggle("is-reference-collapsed", collapsed && Number.isFinite(activeIndex));
  dialog.querySelectorAll("[data-function-argument-row]").forEach((row) => {
    const index = Number(row.dataset.functionArgumentRow);
    row.classList.toggle("is-active-reference", collapsed && index === activeIndex);
  });
  dialog.querySelectorAll("[data-data-dialog-action='function-argument-reference-toggle']").forEach((button) => {
    const index = Number(button.dataset.formulaArgumentIndex);
    const active = collapsed && Number.isFinite(activeIndex) && index === activeIndex;
    button.textContent = active ? "↙" : "↗";
    button.title = active ? "ダイアログに戻る" : "参照範囲を選択";
    button.setAttribute("aria-label", active ? "ダイアログに戻る" : "参照範囲を選択");
    button.setAttribute("aria-expanded", active ? "false" : "true");
  });
}



function syncFunctionArgumentRows(dialog, definition) {
  const list = dialog.querySelector(".function-argument-list");
  if (!list) return;
  const values = [...dialog.querySelectorAll("[data-formula-argument-input]")].map((input) => String(input.value || ""));
  const specs = functionArgumentDisplaySpecs(definition, values);
  const currentLabels = [...dialog.querySelectorAll("[data-formula-argument-input]")].map((input) => input.dataset.formulaArgumentLabel || "");
  const nextLabels = specs.map((spec) => spec.label);
  if (currentLabels.length === nextLabels.length && currentLabels.every((label, index) => label === nextLabels[index])) return;

  const active = dialog.querySelector("[data-formula-argument-input]:focus");
  const activeIndex = Number(active?.dataset.formulaArgumentIndex);
  const selectionStart = active?.selectionStart ?? null;
  const selectionEnd = active?.selectionEnd ?? null;
  list.innerHTML = functionArgumentRowsHtml(specs, values);
  const restoreIndex = Number.isFinite(activeIndex) ? Math.min(activeIndex, specs.length - 1) : 0;
  const restored = dialog.querySelector(`[data-formula-argument-input][data-formula-argument-index="${restoreIndex}"]`);
  if (restored && active) {
    restored.focus({ preventScroll: true });
    if (selectionStart != null && selectionEnd != null) {
      try {
        restored.setSelectionRange(selectionStart, selectionEnd);
      } catch {
        // Some input types may reject explicit selection; text inputs should accept it.
      }
    }
  }
}



function formulaArgumentDescription(definition, spec, index) {
  if (!spec) return "";
  const label = String(spec.label || "");
  const normalized = label.toLowerCase();
  const prefix = `${label}: `;
  const required = spec.optional ? "省略可能です。" : "必ず指定します。";
  const descriptions = FORMULA_ARGUMENT_DESCRIPTIONS[definition.name] || {};
  const dynamicAliases = [];
  if (/\d+$/u.test(label)) {
    dynamicAliases.push(label.replace(/\d+$/u, "2"), label.replace(/\d+$/u, "1"));
  }
  if (/\d+$/u.test(normalized)) {
    dynamicAliases.push(normalized.replace(/\d+$/u, "2"), normalized.replace(/\d+$/u, "1"));
  }
  const specific = descriptions[label] || descriptions[normalized] || dynamicAliases.map((alias) => descriptions[alias]).find(Boolean);
  if (specific) return `${prefix}${specific}`;
  if (/logical|condition|criteria/.test(normalized)) return `${prefix}条件または論理値を指定します。${required}`;
  if (/lookup_value|find_text|old_text|new_text|text|delimiter|format|friendly|link_location/.test(normalized)) return `${prefix}文字列またはセル参照を指定します。${required}`;
  if (/array|range|ref|table|return|sum/.test(normalized)) return `${prefix}セル範囲、配列、または参照を指定します。${required}`;
  if (/number|num|value|rate|nper|pv|fv|pmt|year|month|day|hour|minute|second|index|row|column|digits|type|mode|order|basis|guess|bottom|top/.test(normalized)) {
    return `${prefix}数値、数式、または数値を返す参照を指定します。${required}`;
  }
  return `${prefix}${required}`;
}



function buildFunctionArgumentsFormula(dialog) {
  const functionName = state.dataDialog?.functionName || "SUM";
  const inputs = [...(dialog?.querySelectorAll("[data-formula-argument-input]") || [])];
  const values = inputs.map((input) => String(input.value || "").trim());
  while (values.length && !values[values.length - 1]) values.pop();
  return `=${functionName}(${values.join(",")})`;
}



function evaluateFormulaPreviewExpression(expression) {
  const formula = String(expression || "").trim();
  if (!formula) return "";
  const candidate = formula.startsWith("=") ? formula : `=${formula}`;
  const sheetIndex = Number(state.dataDialog?.sourceSheetIndex ?? state.activeSheetIndex);
  const sheet = state.model?.sheets?.[sheetIndex] || activeSheet();
  const sheetId = sheet ? getSheetId(sheet.name) : null;
  if (!state.hf?.calculateFormula || sheetId == null) return "";
  try {
    const translated = translateFormulaForCurrentFormulaEngine(candidate, formulaEngineSheetIndex(sheet), selectionToolbarPoint() || {});
    return formulaPreviewDisplayValue(state.hf.calculateFormula(translated, sheetId));
  } catch (error) {
    return formulaPreviewDisplayValue(error);
  }
}



function formulaPreviewDisplayValue(value) {
  if (value == null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return displayValue(value);
  if (Array.isArray(value)) return value.length ? formulaPreviewDisplayValue(value[0]) : "";
  if (typeof value === "object") {
    const type = String(value.type || value.error || "").toUpperCase();
    if (type) {
      const normalized = type.replace(/^#/u, "").replace(/!$/u, "");
      const mapped = {
        DIV_BY_ZERO: "#DIV/0!",
        DIV0: "#DIV/0!",
        NAME: "#NAME?",
        REF: "#REF!",
        VALUE: "#VALUE!",
        NA: "#N/A",
        NUM: "#NUM!",
        NULL: "#NULL!",
        CYCLE: "#CYCLE!",
      }[normalized] || (type.startsWith("#") ? type : `#${normalized}!`);
      return mapped;
    }
    if (Object.prototype.hasOwnProperty.call(value, "value")) return formulaPreviewDisplayValue(value.value);
    if (value.message) return String(value.message);
  }
  return String(value);
}



function applyFunctionArgumentsDialog(dialog) {
  const source = state.dataDialog?.editorSource || null;
  const point = source ? null : (state.dataDialog?.sourcePoint || selectionToolbarPoint());
  if (!point && !source) {
    showMessageBox("関数を入力するセルを選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const functionName = state.dataDialog?.functionName || "SUM";
  const definition = formulaFunctionByName(functionName);
  const formula = buildFunctionArgumentsFormula(dialog);
  closeDataDialog();
  if (source) {
    applyFunctionArgumentsFormulaToEditorSource(source, formula, definition);
    return;
  }
  if (point.sheetIndex !== state.activeSheetIndex) {
    activateSheetForFormulaEditResult(point);
  }
  commitCellValue(point.row, point.col, formula, { skipValidation: true });
  recordFormulaFunctionUse(definition.name);
  setStatus(`${columnName(point.col)}${point.row} に ${definition.label} 関数を入力しました。`);
}



function applyFunctionArgumentsFormulaToEditorSource(source, formula, definition) {
  const functionCall = String(formula || "").replace(/^=/u, "");
  const text = String(source?.text || "");
  const start = Math.max(0, Math.min(text.length, Number(source?.start) || 0));
  const end = Math.max(start, Math.min(text.length, Number(source?.end) || start));
  const nextText = `${text.slice(0, start)}${functionCall}${text.slice(end)}`;
  const nextCaret = start + functionCall.length;
  const editor = functionArgumentsEditorElement(source);
  if (!editor) {
    setStatus("関数の引数を反映できませんでした。");
    return;
  }
  if (editor !== $formulaInput[0] && editor.getAttribute("contenteditable") !== "true") {
    editor.setAttribute("contenteditable", "true");
  }
  formulaEditorSetTextAndCaret(editor, nextText, nextCaret);
  if (editor === $formulaInput[0]) {
    syncFormulaEditPreview();
  } else {
    syncEditableCellAlignment(editor);
  }
  recordFormulaFunctionUse(definition.name);
  updateFormulaReferenceUi();
  updateFormulaAutocompleteForEditor(editor);
  editor.focus({ preventScroll: true });
  setStatus(`数式に ${definition.label || definition.name} 関数の引数を反映しました。`);
}



function functionArgumentsEditorElement(source) {
  if (!source) return null;
  if (source.isFormulaBar) return $formulaInput[0] || null;
  if (source.editor?.isConnected) return source.editor;
  if (source.sheetIndex !== state.activeSheetIndex || !source.row || !source.col) return null;
  return ensureCellElementForPoint({ sheetIndex: source.sheetIndex, row: source.row, col: source.col });
}



function restoreFunctionArgumentsEditorFocus(source) {
  const editor = functionArgumentsEditorElement(source);
  if (!editor) return;
  if (editor !== $formulaInput[0] && editor.getAttribute("contenteditable") !== "true") {
    editor.setAttribute("contenteditable", "true");
  }
  editor.focus({ preventScroll: true });
  updateFormulaReferenceUi();
  updateFormulaAutocompleteForEditor(editor);
}



function returnToInsertFunctionDialogFromArguments() {
  const functionName = state.dataDialog?.functionName || "SUM";
  const definition = formulaFunctionByName(functionName);
  closeDataDialog();
  openInsertFunctionDialog({ selectedFunction: definition.name, category: definition.category || "すべて表示" });
}



function recentFormulaFunctionNames() {
  const names = Array.isArray(state.recentFormulaFunctionNames) && state.recentFormulaFunctionNames.length
    ? state.recentFormulaFunctionNames
    : INSERT_FUNCTION_RECENT_NAMES;
  return [...new Set(names.map((name) => String(name || "").toUpperCase()).filter(Boolean))]
    .filter((name) => formulaInsertFunctionDefinitionByName(name))
    .slice(0, MAX_RECENT_FORMULA_FUNCTIONS);
}



function recordFormulaFunctionUse(functionName) {
  const name = String(functionName || "").trim().toUpperCase();
  if (!name || !formulaInsertFunctionDefinitionByName(name)) return;
  const next = [name, ...recentFormulaFunctionNames().filter((item) => item !== name)].slice(0, MAX_RECENT_FORMULA_FUNCTIONS);
  state.recentFormulaFunctionNames = next;
}



function openFormulaAutoSumMenu(trigger = null) {
  if (!trigger) {
    applyAutoSumToSelection("SUM");
    return;
  }
  openHomeCommandMenu(trigger, "formula-autosum", "オートSUM", `
    ${formulaAutoSumMenuButtonHtml("SUM", "合計(S)", "sigma")}
    ${formulaAutoSumMenuButtonHtml("AVERAGE", "平均(A)", "bar-chart-3")}
    ${formulaAutoSumMenuButtonHtml("COUNT", "数値の個数(C)", "hash")}
    ${formulaAutoSumMenuButtonHtml("MAX", "最大値(M)", "chevrons-up")}
    ${formulaAutoSumMenuButtonHtml("MIN", "最小値(I)", "chevrons-down")}
    <div class="home-command-menu-separator" role="separator"></div>
    ${formulaFunctionDialogButtonHtml("すべて表示", "その他の関数(M)...", "function-square")}
  `);
}



function openFormulaFunctionMenu(trigger = null, menuKey = "math") {
  const menu = FORMULA_RIBBON_FUNCTION_MENUS[menuKey] || FORMULA_RIBBON_FUNCTION_MENUS.math;
  const functionNames = formulaFunctionMenuNames(menuKey, menu);
  if (!trigger) {
    insertFunctionIntoSelection(functionNames[0] || "SUM");
    return;
  }
  const rows = functionNames
    .map((name) => formulaFunctionMenuButtonHtml(name))
    .join("");
  openHomeCommandMenu(trigger, `formula-function-${menuKey}`, menu.label, `
    ${rows}
    <div class="home-command-menu-separator" role="separator"></div>
    ${formulaFunctionDialogButtonHtml(menu.category, "その他の関数(M)...", "function-square")}
  `);
}



function formulaFunctionMenuNames(menuKey, menu) {
  if (menuKey === "recent") return recentFormulaFunctionNames();
  const definitions = formulaInsertFunctionDefinitions();
  const preferred = Array.isArray(menu?.functions) ? menu.functions.map((name) => String(name || "").toUpperCase()).filter(Boolean) : [];
  const available = new Set(definitions.map((definition) => definition.name));
  const categoryNames = definitions
    .filter((definition) => formulaFunctionMenuCategoryMatches(menuKey, menu, definition))
    .map((definition) => definition.name);
  const visibleLimit = Number(menu?.visibleLimit) || (menuKey === "more" ? 220 : menuKey === "math" ? 120 : 96);
  return [...new Set([...preferred, ...categoryNames])]
    .filter((name) => available.has(name))
    .slice(0, visibleLimit);
}



function formulaFunctionMenuCategoryMatches(menuKey, menu, definition) {
  const category = String(definition?.category || "");
  if (menuKey === "more") {
    return ["統計", "データベース", "情報", "エンジニアリング", "互換性", "論理/情報"].includes(category);
  }
  if (menuKey === "logical") return category === "論理" || category === "論理/情報";
  return category === menu?.category;
}



function formulaAutoSumMenuButtonHtml(functionName, label, icon) {
  return `<button type="button" role="menuitem" data-formula-function-menu-action="autosum" data-formula-autosum-function="${escapeAttr(functionName)}"${commandMenuAccessKeyAttr(label)}><i data-lucide="${escapeAttr(icon)}"></i><span>${accessKeyLabelHtml(label)}</span></button>`;
}



function formulaFunctionMenuButtonHtml(functionName) {
  const definition = formulaInsertFunctionDefinitionByName(functionName);
  const name = String(functionName || "").toUpperCase();
  const disabled = !definition;
  const label = definition?.label || name;
  const description = definition?.syntax || formulaDefaultSyntax(name);
  return `
    <button class="has-description" type="button" role="menuitem" data-formula-function-menu-action="insert" data-formula-function-name="${escapeAttr(name)}" ${disabled ? "disabled" : ""}>
      <span class="home-command-menu-check"></span>
      <span class="home-command-menu-text">
        <strong class="home-command-menu-label">${escapeHtml(label)}</strong>
        <small>${escapeHtml(description)}</small>
      </span>
    </button>
  `;
}



function formulaFunctionDialogButtonHtml(category, label, icon) {
  return `<button type="button" role="menuitem" data-formula-function-menu-action="dialog" data-formula-function-category="${escapeAttr(category)}"${commandMenuAccessKeyAttr(label)}><i data-lucide="${escapeAttr(icon)}"></i><span>${accessKeyLabelHtml(label)}</span></button>`;
}



function openDefineNameForSelection(options = {}) {
  const point = selectionToolbarPoint();
  if (!point) {
    showMessageBox("名前を定義するセルを選択してください。", { title: "Cht WebSheet" });
    return;
  }
  openDefineNameDialog(point.row, point.col, options);
}



function openFormulaDefineNameMenu(trigger = null) {
  if (!trigger) {
    openDefineNameForSelection();
    return;
  }
  openHomeCommandMenu(trigger, "formula-define-name", "名前の定義", `
    ${formulaNameMenuSimpleButtonHtml("define", "名前の定義(D)...", "badge-plus")}
    ${formulaNameMenuSimpleButtonHtml("manage", "名前の管理(M)...", "list-checks")}
    <div class="home-command-menu-separator" role="separator"></div>
    ${formulaNameMenuSimpleButtonHtml("create-from-selection", "選択範囲から作成(C)...", "table-2")}
  `);
}



function handleFormulaFunctionMenuAction(event) {
  const action = event.currentTarget.dataset.formulaFunctionMenuAction || "";
  if (action === "autosum") {
    const functionName = event.currentTarget.dataset.formulaAutosumFunction || "SUM";
    closeHomeCommandMenu();
    applyAutoSumToSelection(functionName);
    return;
  }
  if (action === "insert") {
    const functionName = event.currentTarget.dataset.formulaFunctionName || "SUM";
    closeHomeCommandMenu();
    openFunctionArgumentsDialog(functionName);
    return;
  }
  if (action === "dialog") {
    const category = event.currentTarget.dataset.formulaFunctionCategory || "すべて表示";
    closeHomeCommandMenu();
    openInsertFunctionDialog({ category });
  }
}



function openInsertFunctionDialog(options = {}) {
  const definitions = formulaInsertFunctionDefinitions();
  const categories = insertFunctionCategoryOptions(definitions);
  const requestedCategory = options.category || "最近使用した関数";
  const category = categories.includes(requestedCategory) ? requestedCategory : "最近使用した関数";
  const selected = formulaInsertFunctionInitialSelection(definitions, category, options.selectedFunction);
  state.dataDialog = { type: "insert-function", selectedFunction: selected, functionCategory: category, functionSearch: "" };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = insertFunctionDialogHtml(definitions, selected, category);
  document.body.appendChild(overlay);
  syncInsertFunctionDialog(overlay.querySelector("[data-data-dialog='insert-function']"));
  overlay.querySelector("[data-formula-function-choice]:not([hidden])")?.focus({ preventScroll: true });
}



function formulaInsertFunctionInitialSelection(definitions, category, requestedFunction = "") {
  const requested = String(requestedFunction || "").toUpperCase();
  if (requested && definitions.some((definition) => definition.name === requested)) return requested;
  if (category === "最近使用した関数") return definitions.find((definition) => recentFormulaFunctionNames().includes(definition.name))?.name || definitions[0]?.name || "SUM";
  if (category === "すべて表示") return definitions[0]?.name || "SUM";
  return definitions.find((definition) => definition.category === category)?.name || definitions[0]?.name || "SUM";
}



function insertFunctionDialogHtml(definitions, selected, selectedCategory = "最近使用した関数") {
  const categories = insertFunctionCategoryOptions(definitions);
  return `
    <div class="excel-mini-dialog data-dialog insert-function-dialog" role="dialog" aria-modal="true" aria-labelledby="insertFunctionTitle" data-data-dialog="insert-function">
      <div class="excel-mini-titlebar" id="insertFunctionTitle">関数の挿入</div>
      <div class="excel-mini-body data-dialog-body formula-dialog-body">
        <div class="insert-function-search-row">
          <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("関数の検索(S):")}</span>
            <input type="text" data-formula-function-search value="" aria-label="関数の検索">
          </label>
          <button class="excel-mini-button" type="button" data-data-dialog-action="insert-function-search">${accessKeyLabelHtml("検索開始(G)")}</button>
        </div>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("関数の分類(C):")}</span>
          <select data-formula-function-category>
            ${categories.map((category) => `<option value="${escapeAttr(category)}" ${category === selectedCategory ? "selected" : ""}>${escapeHtml(category)}</option>`).join("")}
          </select>
        </label>
        <div class="insert-function-layout">
          <div class="insert-function-list" role="listbox" aria-label="関数名">
            ${definitions.map((item) => insertFunctionChoiceHtml(item, item.name === selected)).join("")}
            <div class="insert-function-empty" hidden>該当する関数はありません。</div>
          </div>
          <div class="insert-function-description">
            ${insertFunctionDescriptionHtml(formulaInsertFunctionDefinitionByName(selected) || formulaFunctionByName(selected))}
          </div>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="insert-function-ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>
  `;
}



function insertFunctionChoiceHtml(item, selected) {
  return `<button class="${selected ? "is-active" : ""}" type="button" role="option" aria-selected="${selected ? "true" : "false"}" data-data-dialog-action="insert-function-select" data-formula-function-choice="${escapeAttr(item.name)}" data-formula-function-category-value="${escapeAttr(item.category)}" data-formula-function-search-text="${escapeAttr(insertFunctionSearchText(item))}">${escapeHtml(item.label)}</button>`;
}



function insertFunctionDescriptionHtml(definition) {
  const signature = definition.syntax || formulaDefaultSyntax(definition.name);
  return `
    <strong>${escapeHtml(definition.label)}</strong>
    <span>${escapeHtml(definition.description)}</span>
    <code>${escapeHtml(signature)}</code>
    <code>${escapeHtml(definition.template(formulaTemplateReference()))}</code>
  `;
}



function syncInsertFunctionDialog(dialog) {
  const category = dialog.querySelector("[data-formula-function-category]")?.value || "最近使用した関数";
  const search = normalizeInsertFunctionSearch(dialog.querySelector("[data-formula-function-search]")?.value || "");
  if (state.dataDialog?.type === "insert-function") {
    state.dataDialog.functionCategory = category;
    state.dataDialog.functionSearch = dialog.querySelector("[data-formula-function-search]")?.value || "";
  }
  let visibleCount = 0;
  dialog.querySelectorAll("[data-formula-function-choice]").forEach((button) => {
    const name = button.dataset.formulaFunctionChoice || "";
    const categoryMatches =
      Boolean(search) ||
      category === "すべて表示" ||
      (category === "最近使用した関数" && recentFormulaFunctionNames().includes(name)) ||
      button.dataset.formulaFunctionCategoryValue === category;
    const searchMatches = !search || normalizeInsertFunctionSearch(button.dataset.formulaFunctionSearchText || "").includes(search);
    const visible = categoryMatches && searchMatches;
    button.hidden = !visible;
    if (visible) visibleCount += 1;
  });
  const selected = state.dataDialog?.selectedFunction || "SUM";
  const visibleSelected = dialog.querySelector(`[data-formula-function-choice="${CSS.escape(selected)}"]:not([hidden])`);
  if (!visibleSelected) {
    const first = dialog.querySelector("[data-formula-function-choice]:not([hidden])");
    if (first) state.dataDialog.selectedFunction = first.dataset.formulaFunctionChoice;
  }
  dialog.querySelectorAll("[data-formula-function-choice]").forEach((button) => {
    const active = button.dataset.formulaFunctionChoice === state.dataDialog?.selectedFunction;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-selected", active ? "true" : "false");
  });
  const description = dialog.querySelector(".insert-function-description");
  if (description) {
    const definition = formulaInsertFunctionDefinitionByName(state.dataDialog?.selectedFunction);
    description.innerHTML = definition ? insertFunctionDescriptionHtml(definition) : "";
  }
  const empty = dialog.querySelector(".insert-function-empty");
  if (empty) empty.hidden = visibleCount > 0;
}



function handleInsertFunctionCategoryChange(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='insert-function']");
  if (dialog) syncInsertFunctionDialog(dialog);
}



function handleInsertFunctionSearchInput(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='insert-function']");
  if (dialog) syncInsertFunctionDialog(dialog);
}



function handleInsertFunctionSearchKeydown(event) {
  if (event.key !== "Enter") return;
  if (isElementImeCompositionEvent(event, event.currentTarget)) return;
  event.preventDefault();
  const dialog = event.currentTarget.closest("[data-data-dialog='insert-function']");
  if (dialog) focusFirstVisibleInsertFunction(dialog);
}



function focusFirstVisibleInsertFunction(dialog) {
  const first = dialog?.querySelector("[data-formula-function-choice]:not([hidden])");
  first?.focus({ preventScroll: true });
}



function insertFunctionSearchText(item) {
  return [item.name, item.label, item.category, item.syntax, item.description].filter(Boolean).join(" ");
}



function normalizeInsertFunctionSearch(value) {
  return String(value || "").trim().toLocaleUpperCase("ja-JP");
}



function openNameManagerDialog(activeIndex = 0, options = {}) {
  const previousFilter = state.dataDialog?.type === "name-manager" ? state.dataDialog.filter : "all";
  state.dataDialog = {
    type: "name-manager",
    activeIndex: Math.max(0, Number(activeIndex) || 0),
    filter: options.filter || previousFilter || "all",
  };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = nameManagerDialogHtml();
  document.body.appendChild(overlay);
  syncNameManagerDialog(overlay.querySelector("[data-data-dialog='name-manager']"));
}



function nameManagerDialogHtml() {
  return `
    <div class="excel-mini-dialog data-dialog name-manager-dialog" role="dialog" aria-modal="true" aria-labelledby="nameManagerTitle" data-data-dialog="name-manager">
      <div class="excel-mini-titlebar" id="nameManagerTitle">名前の管理</div>
      <div class="excel-mini-body data-dialog-body formula-dialog-body">
        <div class="name-manager-toolbar">
          <button class="excel-mini-button" type="button" data-data-dialog-action="name-manager-new">新規(N)...</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="name-manager-edit">編集(E)...</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="name-manager-delete">削除(D)</button>
          <label class="name-manager-filter">
            <span>${accessKeyLabelHtml("フィルター(F):")}</span>
            <select data-formula-name-filter>
              ${nameManagerFilterOptionsHtml()}
            </select>
          </label>
        </div>
        <div class="name-manager-table-wrap">
          <table class="name-manager-table">
            <thead>
              <tr>
                <th>${accessKeyLabelHtml("名前(N)")}</th>
                <th>値</th>
                <th>参照範囲</th>
                <th>範囲</th>
                <th>コメント</th>
              </tr>
            </thead>
            <tbody data-formula-name-table-body></tbody>
          </table>
        </div>
        <div class="name-manager-summary" data-formula-name-summary></div>
        <div class="name-manager-refers-to-row">
          <span>${accessKeyLabelHtml("参照範囲(R):")}</span>
          <input type="text" readonly data-formula-name-current-refers-to>
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">閉じる</button>
      </div>
    </div>
  `;
}



function nameManagerFilterOptionsHtml(selected = state.dataDialog?.filter || "all") {
  return [
    ["all", "すべて"],
    ["workbook", "ブック"],
    ["worksheet", "ワークシート"],
    ["current-sheet", "現在のワークシート"],
    ["errors", "エラーがある名前"],
    ["print-area", "印刷範囲"],
  ].map(([value, label]) => `<option value="${value}" ${value === selected ? "selected" : ""}>${escapeHtml(label)}</option>`).join("");
}



function nameScopeOptionsHtml() {
  return [
    `<option value="workbook">ブック</option>`,
    ...(state.model?.sheets || []).map((sheet, index) => `<option value="sheet:${index}">${escapeHtml(sheet.name)}</option>`),
  ].join("");
}



function definedNameDisplayLabel(definition) {
  const scope = definition.scope === "sheet" && state.model?.sheets?.[definition.sheetIndex]
    ? state.model.sheets[definition.sheetIndex].name
    : "ブック";
  return `${definition.name} - ${scope}`;
}



function selectedDefinedNameIndexFromDialog(dialog) {
  const activeIndex = Number(state.dataDialog?.activeIndex);
  if (dialog?.dataset?.dataDialog === "name-manager" && Number.isFinite(activeIndex)) return activeIndex;
  const value = Number(dialog?.querySelector("[data-formula-name-list]")?.value);
  return Number.isFinite(value) ? value : -1;
}



function selectedDefinedNameFromManager(dialog) {
  const index = selectedDefinedNameIndexFromDialog(dialog);
  return index >= 0 ? state.model?.definedNames?.[index] || null : null;
}



function syncNameManagerDialog(dialog) {
  if (!dialog) return;
  const filter = state.dataDialog?.filter || "all";
  const filterSelect = dialog.querySelector("[data-formula-name-filter]");
  if (filterSelect) filterSelect.value = filter;
  const entries = nameManagerVisibleEntries(filter);
  const activeVisible = entries.some((entry) => entry.index === Number(state.dataDialog?.activeIndex));
  if (!activeVisible) state.dataDialog.activeIndex = entries[0]?.index ?? -1;
  const activeIndex = Number(state.dataDialog.activeIndex);
  const body = dialog.querySelector("[data-formula-name-table-body]");
  if (body) {
    body.innerHTML = entries.length
      ? entries.map((entry) => nameManagerRowHtml(entry, entry.index === activeIndex)).join("")
      : `<tr class="is-empty"><td colspan="5">表示できる名前はありません。</td></tr>`;
  }
  const definition = state.model?.definedNames?.[activeIndex] || null;
  const reference = dialog.querySelector("[data-formula-name-current-refers-to]");
  if (reference) reference.value = definition?.refersTo || "";
  const summary = dialog.querySelector("[data-formula-name-summary]");
  if (summary) summary.textContent = `${entries.length} 個の名前${definition ? ` / 選択: ${definition.name}` : ""}`;
  const disabled = !definition;
  dialog.querySelector("[data-data-dialog-action='name-manager-edit']").disabled = disabled;
  dialog.querySelector("[data-data-dialog-action='name-manager-delete']").disabled = disabled;
}



function nameManagerVisibleEntries(filter = "all") {
  return (state.model?.definedNames || [])
    .map((definition, index) => ({ definition, index }))
    .filter(({ definition }) => !definition.hidden)
    .filter(({ definition }) => {
      if (filter === "workbook") return definition.scope !== "sheet";
      if (filter === "worksheet") return definition.scope === "sheet";
      if (filter === "current-sheet") return definition.scope === "sheet" && Number(definition.sheetIndex) === state.activeSheetIndex;
      if (filter === "errors") return definedNameHasReferenceError(definition);
      if (filter === "print-area") return /^_?xlnm\.print_area$/i.test(definition.name) || /print_area/i.test(definition.name);
      return true;
    })
    .sort((left, right) => left.definition.name.localeCompare(right.definition.name, "ja", { sensitivity: "base" }));
}



function nameManagerRowHtml(entry, active = false) {
  const definition = entry.definition;
  const scope = definedNameScopeLabel(definition);
  return `
    <tr class="${active ? "is-active" : ""}" data-formula-name-row="${entry.index}">
      <td><button type="button" data-data-dialog-action="name-manager-select" data-formula-name-row-index="${entry.index}">${escapeHtml(definition.name)}</button></td>
      <td>${escapeHtml(definedNamePreviewValue(definition))}</td>
      <td title="${escapeAttr(definition.refersTo || "")}">${escapeHtml(definition.refersTo || "")}</td>
      <td>${escapeHtml(scope)}</td>
      <td>${escapeHtml(definition.comment || "")}</td>
    </tr>`;
}



function definedNameScopeLabel(definition) {
  return definition.scope === "sheet" && state.model?.sheets?.[definition.sheetIndex]
    ? state.model.sheets[definition.sheetIndex].name
    : "ブック";
}



function definedNamePreviewValue(definition) {
  if (!definition?.refersTo) return "";
  if (definedNameHasReferenceError(definition)) return "#REF!";
  const baseSheet = definition.scope === "sheet" ? state.model?.sheets?.[definition.sheetIndex] : activeSheet();
  const first = validationReferenceFromFormula(baseSheet, firstDefinedNameReference(definition.refersTo));
  if (!first?.sheet) return String(definition.refersTo || "").replace(/^=/, "");
  const sheetId = getSheetId(first.sheet.name);
  const range = first.range;
  if (range.top === range.bottom && range.left === range.right) {
    return String(getDisplayForCell(first.sheet, sheetId, range.top, range.left) ?? "");
  }
  const value = String(getDisplayForCell(first.sheet, sheetId, range.top, range.left) ?? "");
  return value ? `${value}...` : rangeToLabel(range);
}



function definedNameHasReferenceError(definition) {
  const refersTo = String(definition?.refersTo || "");
  if (/#(REF|NAME|VALUE|DIV\/0|N\/A|NUM|NULL|SPILL|CALC|CYCLE)!?/i.test(refersTo)) return true;
  const baseSheet = definition.scope === "sheet" ? state.model?.sheets?.[definition.sheetIndex] : activeSheet();
  const firstReference = firstDefinedNameReference(refersTo);
  if (!firstReference || !/[A-Z]|!|\$|\d/i.test(firstReference)) return false;
  const reference = validationReferenceFromFormula(baseSheet, firstReference);
  return Boolean(reference && (!reference.sheet || !reference.range));
}



function openPasteNameDialog() {
  const names = pasteableDefinedNames();
  if (!names.length) {
    showMessageBox("貼り付ける名前がありません。", { title: "Cht WebSheet" });
    return;
  }
  state.dataDialog = { type: "paste-name", activeIndex: 0 };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = pasteNameDialogHtml(names);
  document.body.appendChild(overlay);
  overlay.querySelector("[data-formula-paste-name-list]")?.focus({ preventScroll: true });
}



function pasteableDefinedNames() {
  const sheetIndex = state.activeSheetIndex;
  return (state.model?.definedNames || [])
    .filter((definition) =>
      definition?.name &&
      !definition.hidden &&
      (definition.scope !== "sheet" || Number(definition.sheetIndex) === sheetIndex),
    )
    .sort((left, right) => left.name.localeCompare(right.name, "ja", { sensitivity: "base" }));
}



function pasteNameDialogHtml(names) {
  const rows = names.map((definition, index) => `
    <option value="${index}"${index === 0 ? " selected" : ""}>
      ${escapeHtml(definition.name)}    ${escapeHtml(definition.refersTo || "")}
    </option>`).join("");
  return `
    <div class="excel-mini-dialog data-dialog paste-name-dialog" role="dialog" aria-modal="true" aria-labelledby="pasteNameTitle" data-data-dialog="paste-name">
      <div class="excel-mini-titlebar" id="pasteNameTitle">名前の貼り付け</div>
      <div class="excel-mini-body data-dialog-body formula-dialog-body">
        <label class="excel-dialog-field paste-name-list-field">
          <span class="excel-dialog-caption">${accessKeyLabelHtml("名前(N):")}</span>
          <span class="paste-name-list-frame">
            <span class="paste-name-list-header" aria-hidden="true">
              <span>名前</span>
              <span>参照範囲</span>
            </span>
            <select class="name-manager-list paste-name-list" size="${Math.min(9, Math.max(3, names.length))}" data-formula-paste-name-list>
              ${rows}
            </select>
          </span>
        </label>
        <div class="paste-name-reference" data-formula-paste-name-reference>
          ${escapeHtml(names[0]?.refersTo || "")}
        </div>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="paste-name-ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="paste-name-list">一覧の貼り付け(L)</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
}



function selectedPasteNameDefinition(dialog) {
  const names = pasteableDefinedNames();
  const index = Number(dialog?.querySelector("[data-formula-paste-name-list]")?.value || 0);
  return names[Math.max(0, Math.min(names.length - 1, Number.isFinite(index) ? index : 0))] || null;
}



function syncPasteNameDialog(dialog) {
  const definition = selectedPasteNameDefinition(dialog);
  const reference = dialog?.querySelector("[data-formula-paste-name-reference]");
  if (reference) reference.textContent = definition?.refersTo || "";
}



function applyPasteNameDialog(dialog) {
  const definition = selectedPasteNameDefinition(dialog);
  if (!definition) return;
  closeDataDialog();
  pasteDefinedNameReference(definition);
}



function applyPasteNameListDialog(dialog) {
  const names = pasteableDefinedNames();
  if (!names.length) {
    showMessageBox("貼り付ける名前がありません。", { title: "Cht WebSheet" });
    return;
  }
  const point = selectionToolbarPoint();
  if (!point) {
    showMessageBox("名前の一覧を貼り付けるセルを選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const sheet = activeSheet();
  if (!sheet) return;
  const targetRange = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: point.row, col: point.col },
    { sheetIndex: state.activeSheetIndex, row: point.row + names.length - 1, col: point.col + 1 },
  );
  if (!ensureEditableRangeForProtection(sheet, targetRange, "変更")) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = activeSelectionRange();
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));
  ensureSheetSize(sheet, targetRange.bottom, targetRange.right);
  try {
    names.forEach((definition, index) => {
      writePlainTextCellContent(sheet, point.row + index, point.col, definition.name);
      writePlainTextCellContent(sheet, point.row + index, point.col + 1, definition.refersTo || "");
    });
    internCellsInRange(sheet, targetRange);
    state.hf = buildFormulaEngine(state.model);
    closeDataDialog();
    state.selected = { sheetIndex: state.activeSheetIndex, row: point.row, col: point.col };
    state.selectionRange = targetRange;
    state.selectionRanges = [{ ...targetRange }];
    state.selectionStart = state.selected;
    state.selectionAnchor = state.selected;
    state.selectionDrag = null;
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetBefore,
      sheetAfter: captureSheetForHistory(sheet),
      selectionBefore,
      selectionAfter: targetRange,
      selectionBeforeRanges,
      selectionAfterRanges: [{ ...targetRange }],
      label: `${rangeToLabel(targetRange)} 名前一覧の貼り付け`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${names.length} 件の名前一覧を貼り付けました。`);
  } catch (error) {
    console.error(error);
    restoreSheetFromHistory(sheet, sheetBefore);
    state.hf = buildFormulaEngine(state.model);
    renderSheet();
    showMessageBox("名前の一覧を貼り付けできませんでした。", { title: "Cht WebSheet" });
  }
}



function writePlainTextCellContent(sheet, row, col, text) {
  const key = cellKey(row, col);
  const value = String(text ?? "");
  const cell = {
    ...(sheet.cells?.[key] || {}),
    row,
    col,
    raw: value.startsWith("=") ? `'${value}` : value,
    formula: null,
    cached: value,
    display: value,
    css: sheet.cells?.[key]?.css || {},
    borders: sheet.cells?.[key]?.borders || {},
  };
  storeCellModelForSheet(sheet, key, cell);
}



function pasteDefinedNameReference(definition) {
  const token = definition?.name || "";
  if (!token) return;
  const activeEditable = document.activeElement?.closest?.(".sheet-cell[contenteditable='true']");
  if (activeEditable) {
    insertTextIntoEditable(activeEditable, token);
    setStatus(`名前 ${token} を貼り付けました。`);
    return;
  }
  if (document.activeElement === $formulaInput[0] || state.formulaEdit) {
    beginFormulaBarEdit();
    insertTextIntoFormulaBar(token);
    syncFormulaEditPreview();
    setStatus(`名前 ${token} を数式バーに貼り付けました。`);
    return;
  }
  const point = selectionToolbarPoint();
  if (!point) {
    showMessageBox("名前を貼り付けるセルを選択してください。", { title: "Cht WebSheet" });
    return;
  }
  commitCellValue(point.row, point.col, `=${token}`, { skipValidation: true });
  setStatus(`${columnName(point.col)}${point.row} に名前 ${token} を貼り付けました。`);
}



function openUseInFormulaMenu(trigger = null) {
  const names = pasteableDefinedNames();
  if (!trigger) {
    if (names.length) pasteDefinedNameReference(names[0]);
    else showMessageBox("使用できる名前がありません。", { title: "Cht WebSheet" });
    return;
  }
  const rows = names.length
    ? `
      <div class="formula-name-menu-table" role="group" aria-label="使用できる名前">
        <div class="formula-name-menu-header" aria-hidden="true">
          <span>名前</span>
          <span>値</span>
          <span>範囲</span>
          <span>参照範囲</span>
        </div>
        ${names.map((definition, index) => formulaNameMenuButtonHtml(definition, index)).join("")}
      </div>`
    : `<button type="button" role="menuitem" disabled><span class="home-command-menu-check"></span><span>使用できる名前がありません</span></button>`;
  openHomeCommandMenu(trigger, "formula-use-name", "数式で使用", `
    ${rows}
    <div class="home-command-menu-separator" role="separator"></div>
    ${formulaNameMenuSimpleButtonHtml("paste", "名前の貼り付け(P)...", "clipboard-paste", names.length === 0)}
    ${formulaNameMenuSimpleButtonHtml("manage", "名前の管理(M)...", "list-checks")}
  `);
}



function formulaNameMenuButtonHtml(definition, index) {
  const scope = definedNameScopeLabel(definition);
  const preview = definedNamePreviewValue(definition);
  const hasError = definedNameHasReferenceError(definition);
  const refersTo = definition.refersTo || "";
  return `
    <button class="formula-name-menu-row${hasError ? " has-error" : ""}" type="button" role="menuitem" data-formula-name-menu-action="insert" data-formula-name-menu-index="${index}" data-formula-name-menu-error="${hasError ? "true" : "false"}">
      <span class="formula-name-menu-cell formula-name-menu-name">
        <span class="home-command-menu-check">${hasError ? "!" : ""}</span>
        <strong class="home-command-menu-label">${escapeHtml(definition.name)}</strong>
      </span>
      <span class="formula-name-menu-cell formula-name-menu-value" title="${escapeAttr(preview)}">${escapeHtml(preview)}</span>
      <span class="formula-name-menu-cell formula-name-menu-scope" title="${escapeAttr(scope)}">${escapeHtml(scope)}</span>
      <span class="formula-name-menu-cell formula-name-menu-reference" title="${escapeAttr(refersTo)}">
        ${escapeHtml(hasError ? "#REF!" : refersTo)}
      </span>
    </button>
  `;
}



function formulaNameMenuSimpleButtonHtml(action, label, icon, disabled = false) {
  return `<button type="button" role="menuitem" data-formula-name-menu-action="${escapeAttr(action)}"${commandMenuAccessKeyAttr(label)} ${disabled ? "disabled" : ""}><i data-lucide="${escapeAttr(icon)}"></i><span>${accessKeyLabelHtml(label)}</span></button>`;
}



function handleFormulaNameMenuAction(event) {
  const action = event.currentTarget.dataset.formulaNameMenuAction || "";
  if (action === "define") {
    closeHomeCommandMenu();
    openDefineNameForSelection();
    return;
  }
  if (action === "insert") {
    const names = pasteableDefinedNames();
    const index = Number(event.currentTarget.dataset.formulaNameMenuIndex || 0);
    const definition = names[Math.max(0, Math.min(names.length - 1, Number.isFinite(index) ? index : 0))];
    closeHomeCommandMenu();
    pasteDefinedNameReference(definition);
    return;
  }
  if (action === "paste") {
    closeHomeCommandMenu();
    openPasteNameDialog();
    return;
  }
  if (action === "manage") {
    closeHomeCommandMenu();
    openNameManagerDialog();
    return;
  }
  if (action === "create-from-selection") {
    closeHomeCommandMenu();
    openCreateNamesDialog();
  }
}



function insertTextIntoFormulaBar(text) {
  const input = $formulaInput[0];
  if (!input) return;
  const value = String(input.value || "");
  const start = Number.isFinite(input.selectionStart) ? input.selectionStart : value.length;
  const end = Number.isFinite(input.selectionEnd) ? input.selectionEnd : start;
  input.value = `${value.slice(0, start)}${text}${value.slice(end)}`;
  const caret = start + String(text).length;
  input.focus({ preventScroll: true });
  input.setSelectionRange?.(caret, caret);
}



function insertTextIntoEditable(element, text) {
  const selection = window.getSelection?.();
  if (!selection || !element) return false;
  let range = selection.rangeCount ? selection.getRangeAt(0) : null;
  if (!range || !element.contains(range.startContainer) || !element.contains(range.endContainer)) {
    placeCaretAtEnd(element);
    range = selection.rangeCount ? selection.getRangeAt(0) : null;
  }
  if (!range) return false;
  range.deleteContents();
  const node = document.createTextNode(String(text ?? ""));
  range.insertNode(node);
  range.setStartAfter(node);
  range.collapse(true);
  selection.removeAllRanges();
  selection.addRange(range);
  element.dispatchEvent(new Event("input", { bubbles: true }));
  return true;
}



function cycleFormulaReferenceInFormulaBar() {
  const input = $formulaInput[0];
  if (!input) return false;
  const value = String(input.value || "");
  const result = cycleFormulaReferenceInText(value, input.selectionStart ?? value.length, input.selectionEnd ?? input.selectionStart ?? value.length);
  if (!result) return false;
  input.value = result.text;
  input.focus({ preventScroll: true });
  input.setSelectionRange?.(result.start, result.end);
  syncFormulaEditPreview();
  updateFormulaReferenceUi();
  updateFormulaAutocompleteForEditor(input);
  setStatus("参照形式を切り替えました。");
  return true;
}



function cycleFormulaReferenceInEditable(element) {
  if (!element) return false;
  const text = editablePlainText(element);
  const selectionOffsets = editableSelectionOffsets(element);
  const result = cycleFormulaReferenceInText(
    text,
    selectionOffsets?.start ?? text.length,
    selectionOffsets?.end ?? selectionOffsets?.start ?? text.length,
  );
  if (!result) return false;
  element.textContent = result.text;
  setEditableSelectionOffsets(element, result.start, result.end);
  syncEditableCellAlignment(element);
  element.dispatchEvent(new Event("input", { bubbles: true }));
  updateFormulaReferenceUi();
  updateFormulaAutocompleteForEditor(element);
  setStatus("参照形式を切り替えました。");
  return true;
}



function cycleFormulaReferenceInFunctionArgumentInput(input) {
  if (!input) return false;
  const value = String(input.value || "");
  const hasFormulaPrefix = value.trimStart().startsWith("=");
  const prefixOffset = hasFormulaPrefix ? 0 : 1;
  const source = hasFormulaPrefix ? value : `=${value}`;
  const selectionStart = (input.selectionStart ?? value.length) + prefixOffset;
  const selectionEnd = (input.selectionEnd ?? input.selectionStart ?? value.length) + prefixOffset;
  const result = cycleFormulaReferenceInText(source, selectionStart, selectionEnd);
  if (!result) return false;
  input.value = hasFormulaPrefix ? result.text : result.text.slice(1);
  const start = Math.max(0, result.start - prefixOffset);
  const end = Math.max(start, result.end - prefixOffset);
  input.focus({ preventScroll: true });
  input.setSelectionRange?.(start, end);
  input.dispatchEvent(new Event("input", { bubbles: true }));
  syncFunctionArgumentsDialog(input.closest("[data-data-dialog='function-arguments']"));
  setStatus("参照形式を切り替えました。");
  return true;
}



function cycleFormulaReferenceInText(text, selectionStart, selectionEnd) {
  const source = String(text || "");
  if (!source.trimStart().startsWith("=")) return null;
  const start = Math.max(0, Math.min(Number(selectionStart) || 0, source.length));
  const end = Math.max(start, Math.min(Number(selectionEnd) || start, source.length));
  const token = formulaReferenceTokenAt(source, start, end);
  if (!token) return null;
  const replacement = cycleFormulaReferenceToken(token.text);
  if (!replacement || replacement === token.text) return null;
  return {
    text: `${source.slice(0, token.start)}${replacement}${source.slice(token.end)}`,
    start: token.start,
    end: token.start + replacement.length,
  };
}



function formulaReferenceTokenRegex() {
  return new RegExp(`(?:${FORMULA_REFERENCE_SHEET_TOKEN_SOURCE})?(?:${FORMULA_REFERENCE_BODY_TOKEN_SOURCE})`, "gi");
}



function formulaReferenceTokensInText(text) {
  const source = String(text || "");
  const regex = formulaReferenceTokenRegex();
  const tokens = [];
  let match;
  while ((match = regex.exec(source))) {
    const tokenStart = match.index;
    const tokenEnd = tokenStart + match[0].length;
    if (!isFormulaReferenceTokenBoundary(source, tokenStart, tokenEnd)) continue;
    if (isInsideFormulaStringLiteral(source, tokenStart)) continue;
    const split = splitFormulaReferenceToken(match[0]);
    tokens.push({
      text: match[0],
      start: tokenStart,
      end: tokenEnd,
      sheetToken: split.sheetToken,
      body: split.body,
    });
  }
  return tokens;
}



function formulaReferenceTokenAt(text, selectionStart, selectionEnd) {
  const tokens = formulaReferenceTokensInText(text);
  for (const token of tokens) {
    const touchesCaret = selectionStart === selectionEnd && selectionStart >= token.start && selectionStart <= token.end;
    const containsSelection = selectionStart >= token.start && selectionEnd <= token.end;
    if (touchesCaret || containsSelection) return token;
  }
  return null;
}



function isFormulaReferenceTokenBoundary(text, start, end) {
  const before = start > 0 ? text[start - 1] : "";
  const after = end < text.length ? text[end] : "";
  return !/[A-Za-z0-9_$]/.test(before) && !/[A-Za-z0-9_]/.test(after);
}



function isInsideFormulaStringLiteral(text, index) {
  let inString = false;
  for (let i = 0; i < index; i += 1) {
    if (text[i] !== '"') continue;
    if (text[i + 1] === '"') {
      i += 1;
      continue;
    }
    inString = !inString;
  }
  return inString;
}



function splitFormulaReferenceToken(token) {
  const text = String(token || "");
  const bangIndex = text.lastIndexOf("!");
  if (bangIndex < 0) return { sheetToken: "", body: text };
  return { sheetToken: text.slice(0, bangIndex), body: text.slice(bangIndex + 1) };
}



function cycleFormulaReferenceToken(token) {
  const { sheetToken, body } = splitFormulaReferenceToken(token);
  const prefix = sheetToken ? `${sheetToken}!` : "";
  const parts = body.split(":");
  if (!parts.length || parts.length > 2) return token;
  const columnRange = parts.length === 2 && parts.every((part) => /^\$?[A-Za-z]{1,3}$/.test(part));
  const rowRange = parts.length === 2 && parts.every((part) => /^\$?[1-9][0-9]{0,6}$/.test(part));
  const cycled = columnRange
    ? parts.map(cycleSingleColumnReferenceToken)
    : rowRange
      ? parts.map(cycleSingleRowReferenceToken)
      : parts.map(cycleSingleCellReferenceToken);
  if (cycled.some((part) => !part)) return token;
  return `${prefix}${cycled.join(":")}`;
}



function cycleSingleCellReferenceToken(token) {
  const match = String(token || "").match(/^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})$/);
  if (!match) return null;
  const [, colAbs, colLetters, rowAbs, rowNumber] = match;
  const stateKey = `${colAbs ? "1" : "0"}${rowAbs ? "1" : "0"}`;
  const nextKey = {
    "00": "11",
    "11": "01",
    "01": "10",
    "10": "00",
  }[stateKey] || "00";
  return `${nextKey[0] === "1" ? "$" : ""}${colLetters.toUpperCase()}${nextKey[1] === "1" ? "$" : ""}${rowNumber}`;
}



function cycleSingleColumnReferenceToken(token) {
  const match = String(token || "").match(/^(\$?)([A-Za-z]{1,3})$/);
  if (!match) return null;
  return `${match[1] ? "" : "$"}${match[2].toUpperCase()}`;
}



function cycleSingleRowReferenceToken(token) {
  const match = String(token || "").match(/^(\$?)([1-9][0-9]{0,6})$/);
  if (!match) return null;
  return `${match[1] ? "" : "$"}${match[2]}`;
}



function editableSelectionOffsets(element) {
  const selection = window.getSelection?.();
  if (!selection || !selection.rangeCount || !element) return null;
  const range = selection.getRangeAt(0);
  if (!element.contains(range.startContainer) || !element.contains(range.endContainer)) return null;
  const startRange = document.createRange();
  startRange.selectNodeContents(element);
  startRange.setEnd(range.startContainer, range.startOffset);
  const endRange = document.createRange();
  endRange.selectNodeContents(element);
  endRange.setEnd(range.endContainer, range.endOffset);
  return {
    start: startRange.toString().length,
    end: endRange.toString().length,
  };
}



function setEditableSelectionOffsets(element, start, end = start) {
  if (!element) return;
  const selection = window.getSelection?.();
  if (!selection) return;
  const textLength = String(element.textContent || "").length;
  const startPoint = textNodePointAtOffset(element, Math.max(0, Math.min(start, textLength)));
  const endPoint = textNodePointAtOffset(element, Math.max(0, Math.min(end, textLength)));
  if (!startPoint || !endPoint) return;
  const range = document.createRange();
  range.setStart(startPoint.node, startPoint.offset);
  range.setEnd(endPoint.node, endPoint.offset);
  selection.removeAllRanges();
  selection.addRange(range);
}



function textNodePointAtOffset(element, offset) {
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  let remaining = offset;
  let node = walker.nextNode();
  if (!node) {
    const textNode = document.createTextNode("");
    element.appendChild(textNode);
    return { node: textNode, offset: 0 };
  }
  while (node) {
    const length = node.nodeValue.length;
    if (remaining <= length) return { node, offset: remaining };
    remaining -= length;
    const next = walker.nextNode();
    if (!next) return { node, offset: length };
    node = next;
  }
  return null;
}



function handleNameManagerListChange(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='name-manager']");
  if (!dialog) return;
  state.dataDialog = state.dataDialog || { type: "name-manager" };
  const rowIndex = Number(event.currentTarget.dataset.formulaNameRowIndex);
  state.dataDialog.activeIndex = Number.isFinite(rowIndex) ? rowIndex : selectedDefinedNameIndexFromDialog(dialog);
  syncNameManagerDialog(dialog);
}



function handleNameManagerFilterChange(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='name-manager']");
  if (!dialog || state.dataDialog?.type !== "name-manager") return;
  state.dataDialog.filter = event.currentTarget.value || "all";
  syncNameManagerDialog(dialog);
}

export {
  activeFormulaReferenceEditor,
  activeFunctionArgumentReferenceInput,
  applyFormulaReferenceDragRange,
  applyFunctionArgumentsDialog,
  applyFunctionArgumentsFormulaToEditorSource,
  applyPasteNameDialog,
  applyPasteNameListDialog,
  buildFunctionArgumentsFormula,
  cellEditBlurStartsFormulaSheetNavigation,
  clearFormulaReferenceState,
  clearPendingFormulaReferenceSheetTab,
  closeFormulaAutocomplete,
  closeFormulaAutocompleteOnEscape,
  closeFormulaAutocompleteOnOutsidePointer,
  collapseFunctionArgumentReferencePicker,
  cycleFormulaReferenceInEditable,
  cycleFormulaReferenceInFormulaBar,
  cycleFormulaReferenceInFunctionArgumentInput,
  cycleFormulaReferenceInText,
  cycleFormulaReferenceToken,
  cycleSingleCellReferenceToken,
  cycleSingleColumnReferenceToken,
  cycleSingleRowReferenceToken,
  definedNameDisplayLabel,
  definedNameHasReferenceError,
  definedNamePreviewValue,
  definedNameScopeLabel,
  editableSelectionOffsets,
  evaluateFormulaPreviewExpression,
  expandFunctionArgumentReferencePicker,
  extendFormulaHeaderReferenceSelection,
  finishFormulaReferenceDrag,
  focusFirstVisibleInsertFunction,
  formulaArgumentDefaultValue,
  formulaArgumentDescription,
  formulaAutoSumMenuButtonHtml,
  formulaEditorCaretOffset,
  formulaEditorSelectionOffsets,
  formulaEditorSetTextAndCaret,
  formulaEditorText,
  formulaEditorTextOffsetForBoundary,
  formulaFunctionArgumentSpecs,
  formulaFunctionDialogButtonHtml,
  formulaFunctionMenuButtonHtml,
  formulaFunctionMenuCategoryMatches,
  formulaFunctionMenuNames,
  formulaHeaderReferenceRange,
  formulaInsertFunctionInitialSelection,
  formulaNameMenuButtonHtml,
  formulaNameMenuSimpleButtonHtml,
  formulaPreviewDisplayValue,
  formulaReferenceEditorSheetIndex,
  formulaReferenceInsertionContext,
  formulaReferenceLabel,
  formulaReferenceTokenAt,
  formulaReferenceTokenRegex,
  formulaReferenceTokensInText,
  formulaTemplateReference,
  functionArgumentDisplaySpecs,
  functionArgumentLabelSequence,
  functionArgumentRowHtml,
  functionArgumentRowsHtml,
  functionArgumentsDialogHtml,
  functionArgumentsEditorElement,
  handleFormulaFunctionMenuAction,
  handleFormulaNameMenuAction,
  handleFunctionArgumentInput,
  handleFunctionArgumentKeydown,
  handleInsertFunctionCategoryChange,
  handleInsertFunctionSearchInput,
  handleInsertFunctionSearchKeydown,
  handleNameManagerFilterChange,
  handleNameManagerListChange,
  insertFunctionChoiceHtml,
  insertFunctionDescriptionHtml,
  insertFunctionDialogHtml,
  insertFunctionIntoSelection,
  insertFunctionSearchText,
  insertTextIntoEditable,
  insertTextIntoFormulaBar,
  isFormulaReferenceTokenBoundary,
  isFunctionArgumentReferenceEditor,
  isInsideFormulaStringLiteral,
  markFormulaReferenceSheetTabPointerDownCapture,
  nameManagerDialogHtml,
  nameManagerFilterOptionsHtml,
  nameManagerRowHtml,
  nameManagerVisibleEntries,
  nameScopeOptionsHtml,
  nextFunctionArgumentLabel,
  normalizeInsertFunctionSearch,
  openDefineNameForSelection,
  openFormulaAutoSumMenu,
  openFormulaDefineNameMenu,
  openFormulaFunctionMenu,
  openFunctionArgumentsDialog,
  openInsertFunctionDialog,
  openNameManagerDialog,
  openPasteNameDialog,
  openUseInFormulaMenu,
  pasteDefinedNameReference,
  pasteNameDialogHtml,
  pasteableDefinedNames,
  placeCaretAtTextOffset,
  pointsEqualLoose,
  preserveCellFormulaEditForSheetNavigation,
  recentFormulaFunctionNames,
  recordFormulaFunctionUse,
  repeatArgumentGroupStart,
  repeatedFunctionArgumentSpec,
  restoreFunctionArgumentsEditorFocus,
  returnToInsertFunctionDialogFromArguments,
  runFormulaReferenceAutoScroll,
  selectedDefinedNameFromManager,
  selectedDefinedNameIndexFromDialog,
  selectedPasteNameDefinition,
  setEditableSelectionOffsets,
  singleCellRangeFromPoint,
  splitFormulaReferenceToken,
  splitFunctionArgumentSyntax,
  startFormulaHeaderReferenceSelection,
  startFormulaReferenceSelection,
  stopFormulaReferenceAutoScroll,
  syncFunctionArgumentReferencePicker,
  syncFunctionArgumentRows,
  syncFunctionArgumentsDialog,
  syncInsertFunctionDialog,
  syncNameManagerDialog,
  syncPasteNameDialog,
  textNodePointAtOffset,
  toggleFunctionArgumentReferencePicker,
  updateFormulaReferenceAutoScroll,
  updateFormulaReferenceDrag,
  updateFormulaReferenceFromClient,
  updateFormulaReferenceUi,
  writePlainTextCellContent,
};
