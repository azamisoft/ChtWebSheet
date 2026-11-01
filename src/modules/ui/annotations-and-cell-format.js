// Feature module; shared bindings are imported explicitly.
import {
  $ribbonPanels,
  $sheetHost,
  ANNOTATION_POINTER_HALF_HEIGHT,
  ANNOTATION_POINTER_TIP_OFFSET,
  COLUMN_HEADER_HEIGHT,
  FORMAT_DIALOG_DATE_FORMATS,
  LINK_DIALOG_KINDS,
  sheetPersistentNoteCellCacheBySheet,
  state,
} from "../../app-core.js";
import {
  openGetDataFromTableRangeDialog,
} from "../data/query-transforms.js";
import {
  openSortDialog,
  toggleAutoFilterForSelection,
} from "../data/sort-and-autofilter.js";
import {
  openValidationDropdownPopup,
} from "../data/validation-and-consolidation.js";
import {
  addCustomShapeVertex,
  deleteCustomShapeVertex,
  enterShapePointEditMode,
  groupSelectedShapes,
  reorderSelectedObjects,
  toggleCustomShapePathClosed,
  ungroupSelectedShapes,
} from "../drawing/transforms-and-connectors.js";
import {
  deleteSelectedImages,
  openShapeTextEditor,
} from "../editing/autofill.js";
import {
  normalizeColorValue,
  openShapeFormatPane,
  openShapeSelectionPane,
} from "../editing/ink-and-clipboard.js";
import {
  copySelectedImages,
  copySelectedRange,
  cutSelectedImages,
  cutSelectedRange,
  executeWorkbookClipboardShortcut,
  isElementImeCompositionEvent,
  pasteFromClipboard,
  pasteObjectClipboard,
} from "../editing/keyboard-and-paste.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  affectedSheetIndexesForGroupedRanges,
  cellPointFromClientPosition,
  expandRangeForMerges,
  findMergeForCell,
  materializedRangesForSheet,
  normalizeRange,
  rangeToLabel,
  rangesToLabel,
  refreshChartsForGroupedRanges,
  restoreGroupedSheetsAfterFailure,
  selectedHeaderRangesForKind,
  selectionRangeFromPoints,
  updateSelectionUi,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  cellHasComment,
  cellHasFill,
  cellHasNote,
  coordinateToPixels,
  fontOptionFromCss,
  fontSizeOptionFromCss,
  isBoldCssValue,
  selectedImageObjects,
  selectionToolbarPoint,
  updateFormulaBarForCell,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  applyCellStructureOperation,
  captureCellsForHistory,
  captureCellsForHistoryRanges,
  captureSheetForHistory,
  captureSheetsForHistory,
  cellEffectiveCss,
  clearSelectionContents,
  cloneCellModel,
  cloneDefinedNames,
  defaultColumnModel,
  defaultRowModel,
  deleteEntireColumns,
  deleteEntireRows,
  insertEntireColumns,
  insertEntireRows,
  internCellsInRange,
  internCellsInRanges,
  normalizeSheetModelAfterStructureChange,
  pushHistory,
  restoreSheetFromHistory,
  writeClonedClipboardCellContent,
} from "../editing/structure-and-history.js";
import {
  openNameManagerDialog,
} from "../formula/dialogs-and-names.js";
import {
  applyCellNumberFormat,
  dateSerialFromValueForNumberFormat,
  ensureCellModel,
  forEachCellInRange,
  rangeIntersects,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  accessKeyLabelHtml,
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
} from "./help-and-utilities.js";
import {
  absoluteRangeReference,
  activeSheetEditElement,
  closeFileSideMenu,
  closeObjectLinkDialog,
  commitDataOperation,
  contextMenuActionPoint,
  defaultLinkDialogValue,
  definedNameDuplicateMessage,
  definedNameHasConflict,
  definedNameKey,
  definedNameLinkReference,
  ensureSheetSize,
  hyperlinkKind,
  internalHyperlinkReferenceFromDefinedName,
  isPersistentCell,
  linkDialogInitialValues,
  linkKindEmptyMessage,
  linkKindMeta,
  normalizeDefinedNameRefersTo,
  normalizeHyperlinkForKind,
  normalizedDocumentSheetIndex,
  objectLinkDialogHtml,
  openInternalHyperlink,
  openSelectedObjectHyperlink,
  quoteSheetNameForFormula,
  suggestedDefinedNameForRange,
  upsertDefinedName,
  validateDefinedName,
} from "./ribbon-and-links.js";
import {
  activeQuickAnalysisContext,
  closeSheetTabContextMenu,
  contextMenuTargetRanges,
  ensureActiveSheetFormatCellsAllowed,
  ensureEditableRangesForProtection,
  ensureHeaderFormatAllowed,
  ensureHeaderFormatAllowedForSheet,
  ensureSheetProtectionPermission,
  openQuickAnalysisPopup,
  openTranslationPane,
  quickAnalysisContext,
  quickAnalysisNumericColumns,
  setActiveSheetView,
} from "./sheet-protection-and-translation.js";
import {
  buildFormulaEngine,
  fontFamilyForExcel,
  hexToRgb,
  isSheetVisible,
  normalizeDefinedNameLookup,
  rgbToHex,
  verticalAlignToFlex,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  activeSheet,
  displayValue,
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
  normalizeNote,
} from "../workbook/tables-and-validation.js";
import {
  insertPageBreakFromSelection,
  removePageBreakFromSelection,
  resetAllPageBreaks,
} from "../workbook/vba-and-printing.js";


function interpolateHexColor(from, to, ratio) {
  const left = hexToRgb(from) || [255, 255, 255];
  const right = hexToRgb(to) || [255, 255, 255];
  const next = left.map((value, index) => Math.round(value + (right[index] - value) * Math.max(0, Math.min(1, ratio))));
  return rgbToHex(next);
}



function applyQuickAnalysisTotal(method) {
  const context = activeQuickAnalysisContext();
  if (!context) return;
  const { sheet, range } = context;
  const numericColumns = quickAnalysisNumericColumns(sheet, range);
  if (!numericColumns.length && method !== "count") {
    showMessageBox("集計に使用できる数値データが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  if (applyQuickAnalysisTotalToSheetGroup(method, context)) return;
  const plan = quickAnalysisTotalPlanForSheet(sheet, range, method, state.activeSheetIndex);
  if (!plan?.targetRange || !plan.writes.length) {
    showMessageBox("集計に使用できるデータが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  if (!ensureEditableRangesForProtection(sheet, [plan.targetRange], "クイック分析")) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const sheetId = getSheetId(sheet.name);
  writeQuickAnalysisTotalPlan(sheet, sheetId, plan);
  state.hf = buildFormulaEngine(state.model);
  state.selected = { sheetIndex: state.activeSheetIndex, row: plan.targetRange.top, col: plan.targetRange.left };
  state.selectionRange = plan.targetRange;
  state.selectionRanges = [plan.targetRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: range,
    selectionAfter: plan.targetRange,
    selectionBeforeRanges: [{ ...range }],
    selectionAfterRanges: [{ ...plan.targetRange }],
    label: `${rangeToLabel(range)} クイック分析 ${quickAnalysisTotalLabel(method)}`,
  });
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(`${rangeToLabel(plan.targetRange)} に${quickAnalysisTotalLabel(method)}を追加しました。`);
}



function quickAnalysisTotalPlanForSheet(sheet, range, method, sheetIndex = state.activeSheetIndex) {
  const numericColumns = quickAnalysisNumericColumns(sheet, range);
  const writes = [];
  let targetRange = null;
  if (method === "running") {
    const sourceCol = numericColumns[0]?.col || range.left;
    const targetCol = range.right + 1;
    targetRange = normalizeRange(
      { sheetIndex, row: range.top, col: targetCol },
      { sheetIndex, row: range.bottom, col: targetCol },
    );
    for (let row = range.top; row <= range.bottom; row += 1) {
      writes.push({ row, col: targetCol, formula: `=SUM($${columnName(sourceCol)}$${range.top}:${columnName(sourceCol)}${row})` });
    }
  } else {
    const targetRow = range.bottom + 1;
    targetRange = normalizeRange(
      { sheetIndex, row: targetRow, col: range.left },
      { sheetIndex, row: targetRow, col: range.right },
    );
    const numericSet = new Set(numericColumns.map((item) => item.col));
    for (let col = range.left; col <= range.right; col += 1) {
      if (method !== "count" && !numericSet.has(col)) continue;
      const formula = quickAnalysisTotalFormula(method, range, col);
      if (!formula) continue;
      writes.push({ row: targetRow, col, formula, numFmt: method === "percent" ? "0%" : "" });
    }
  }
  return { targetRange, writes };
}



function writeQuickAnalysisTotalPlan(sheet, sheetId, plan) {
  if (!plan?.targetRange) return;
  ensureSheetSize(sheet, plan.targetRange.bottom, plan.targetRange.right);
  (plan.writes || []).forEach((write) => {
    writeCellContent(sheet, sheetId, write.row, write.col, write.formula);
    if (write.numFmt) ensureCellModel(sheet, write.row, write.col).numFmt = write.numFmt;
  });
  internCellsInRange(sheet, plan.targetRange);
}



function applyQuickAnalysisTotalToSheetGroup(method, context) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !context?.range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    const sheetId = getSheetId(sheet?.name);
    if (!sheet || sheetId == null) continue;
    const sourceRange = materializedRangesForSheet(sheet, [context.range], sheetIndex)[0];
    if (!sourceRange) continue;
    const plan = quickAnalysisTotalPlanForSheet(sheet, sourceRange, method, sheetIndex);
    if (!plan.targetRange || !plan.writes.length) continue;
    if (!ensureEditableRangesForProtection(sheet, [plan.targetRange], "クイック分析")) return true;
    items.push({ sheetIndex, sheet, sheetId, sourceRange, plan, ranges: [plan.targetRange] });
  }
  if (!items.length) {
    showMessageBox("集計に使用できるデータが見つかりません。", { title: "Cht WebSheet" });
    return true;
  }
  const affectedSheetIndexes = affectedSheetIndexesForGroupedRanges(items);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  const activeItem = items.find((item) => item.sheetIndex === state.activeSheetIndex) || items[0];
  try {
    items.forEach((item) => writeQuickAnalysisTotalPlan(item.sheet, item.sheetId, item.plan));
    refreshChartsForGroupedRanges(items);
    const targetRange = activeItem.plan.targetRange;
    state.hf = buildFormulaEngine(state.model);
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
      selectionBefore: context.range,
      selectionAfter: targetRange,
      selectionBeforeRanges: [{ ...context.range }],
      selectionAfterRanges: [{ ...targetRange, sheetIndex: state.activeSheetIndex }],
      label: `${rangeToLabel(context.range)} ${sheetIndexes.length} シートクイック分析 ${quickAnalysisTotalLabel(method)}`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${rangeToLabel(targetRange)} に${quickAnalysisTotalLabel(method)}を追加しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`クイック分析の集計を追加できませんでした：${error.message || error}`);
  }
  return true;
}



function quickAnalysisTotalFormula(method, range, col) {
  const columnRange = `$${columnName(col)}$${range.top}:$${columnName(col)}$${range.bottom}`;
  switch (method) {
    case "sum":
      return `=SUM(${columnRange})`;
    case "average":
      return `=AVERAGE(${columnRange})`;
    case "count":
      return `=COUNTA(${columnRange})`;
    case "percent": {
      const wholeRange = `$${columnName(range.left)}$${range.top}:$${columnName(range.right)}$${range.bottom}`;
      return `=SUM(${columnRange})/SUM(${wholeRange})`;
    }
    default:
      return "";
  }
}



function quickAnalysisTotalLabel(method) {
  switch (method) {
    case "average":
      return "平均";
    case "count":
      return "データの個数";
    case "percent":
      return "% 合計";
    case "running":
      return "累計";
    default:
      return "合計";
  }
}



function cellPasteOptionsHtml(row, col) {
  return `<div class="cell-paste-options" role="group" aria-label="貼り付けのオプション">
    <div class="cell-paste-options-label">貼り付けのオプション:</div>
    <button class="cell-paste-option" type="button" title="元の書式を保持(K)" data-cell-menu-action="paste-keep-format" data-row="${row}" data-col="${col}"><span>1</span></button>
    <button class="cell-paste-option" type="button" title="貼り付け先の書式に合わせる(M)" data-cell-menu-action="paste-match-format" data-row="${row}" data-col="${col}"><span>2</span></button>
  </div>`;
}



function positionPopupMenu(menu, clientX, clientY) {
  menu.style.visibility = "hidden";
  menu.style.left = "0px";
  menu.style.top = "0px";
  const rect = menu.getBoundingClientRect();
  const margin = 8;
  const left = Math.max(margin, Math.min(clientX, window.innerWidth - rect.width - margin));
  const top = Math.max(margin, Math.min(clientY, window.innerHeight - rect.height - margin));
  menu.style.left = `${Math.round(left)}px`;
  menu.style.top = `${Math.round(top)}px`;
  menu.style.visibility = "visible";
}



function closeCellContextMenuOnOutsidePointer(event) {
  if (!state.cellContextMenu) return;
  if (event.target.closest?.("#cellContextMenu")) return;
  closeCellContextMenu();
}



function closeCellContextMenu() {
  document.querySelector("#cellContextMenu")?.remove();
  state.cellContextMenu = null;
}



function quickAnalysisAvailableForContext(row, col) {
  return Boolean(quickAnalysisContext(row, col));
}



function openCellStructureDialog(kind, row, col) {
  if (!state.model) return;
  const activeRanges = activeSelectionRanges();
  if (activeRanges.length > 1) {
    showMessageBox("複数の選択範囲では、この操作を実行できません。", { title: "Cht WebSheet" });
    return;
  }

  const fallbackPoint = row && col ? { sheetIndex: state.activeSheetIndex, row, col } : state.selected;
  const range = activeSelectionRange() || selectionRangeFromPoints(fallbackPoint, fallbackPoint);
  if (!range) return;

  closeCellStructureDialog();
  closeFormatCellsDialog();
  closeSheetTabContextMenu();
  closeFileSideMenu({ restoreHome: true });

  const option = defaultCellStructureOption(kind, range);
  state.cellStructureDialog = { kind, range: { ...range } };

  const overlay = document.createElement("div");
  overlay.id = "cellStructureDialogOverlay";
  overlay.className = "cell-structure-dialog-overlay";
  overlay.innerHTML = cellStructureDialogHtml(kind, option);
  document.body.append(overlay);
  overlay.querySelector("[data-cell-structure-action='ok']")?.focus({ preventScroll: true });
}



function defaultCellStructureOption(kind, range) {
  if (range?.fullRow) return "entire-row";
  if (range?.fullColumn) return "entire-column";
  return kind === "delete" ? "shift-left" : "shift-right";
}



function cellStructureDialogHtml(kind, selectedOption) {
  const title = kind === "delete" ? "削除" : "挿入";
  const groupLabel = kind === "delete" ? "削除" : "挿入";
  const options =
    kind === "delete"
      ? [
          ["shift-left", "セルを左にシフト"],
          ["shift-up", "セルを上にシフト"],
          ["entire-row", "行全体"],
          ["entire-column", "列全体"],
        ]
      : [
          ["shift-right", "セルを右にシフト"],
          ["shift-down", "セルを下にシフト"],
          ["entire-row", "行全体"],
          ["entire-column", "列全体"],
        ];

  return `
    <div id="cellStructureDialog" class="cell-structure-dialog" role="dialog" aria-modal="true" aria-labelledby="cellStructureDialogTitle">
      <div class="cell-structure-titlebar" id="cellStructureDialogTitle">${title}</div>
      <div class="cell-structure-body">
        <fieldset class="cell-structure-group">
          <legend>${groupLabel}</legend>
          ${options
            .map(
              ([value, label], index) => `
                <label class="cell-structure-option">
                  <input type="radio" name="cellStructureOption" value="${escapeAttr(value)}"${value === selectedOption || (!selectedOption && index === 0) ? " checked" : ""} />
                  <span>${escapeHtml(label)}</span>
                </label>
              `,
            )
            .join("")}
        </fieldset>
      </div>
      <div class="cell-structure-actions">
        <button class="cell-structure-button is-primary" type="button" data-cell-structure-action="ok">OK</button>
        <button class="cell-structure-button" type="button" data-cell-structure-action="cancel">キャンセル</button>
      </div>
    </div>
  `;
}



function handleCellStructureDialogAction(event) {
  const action = event.currentTarget.dataset.cellStructureAction;
  if (action === "cancel") {
    closeCellStructureDialog();
    return;
  }
  if (applyCellStructureDialog()) closeCellStructureDialog();
}



function closeCellStructureDialog() {
  document.querySelector("#cellStructureDialogOverlay")?.remove();
  state.cellStructureDialog = null;
}



function closeCellStructureDialogOnEscape(event) {
  if (event.key !== "Escape" || !document.querySelector("#cellStructureDialogOverlay")) return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  closeCellStructureDialog();
}



function applyCellStructureDialog() {
  const dialogState = state.cellStructureDialog;
  const dialog = document.querySelector("#cellStructureDialog");
  if (!dialogState || !dialog) return false;
  const option = dialog.querySelector('input[name="cellStructureOption"]:checked')?.value;
  if (!option) return false;
  return applyCellStructureOperation(dialogState.kind, option, dialogState.range);
}



function openCellAnnotationEditor(type, row, col) {
  if (!state.model) return;
  const fallbackPoint = row && col ? { sheetIndex: state.activeSheetIndex, row, col } : state.selected;
  if (!fallbackPoint) return;

  const sheet = activeSheet();
  const baseRange = normalizeRange(fallbackPoint, fallbackPoint);
  const range = expandRangeForMerges(sheet, baseRange) || baseRange;
  const anchor = { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
  ensureSheetSize(sheet, range.bottom, range.right);

  closeCellStructureDialog();
  closeFormatCellsDialog();
  closeSheetTabContextMenu();
  closeFileSideMenu({ restoreHome: true });

  const cell = sheet.cells[cellKey(anchor.row, anchor.col)] || {};
  const beforeText = annotationTextFromCell(cell, type);
  state.activeSheetIndex = anchor.sheetIndex;
  state.selected = anchor;
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = anchor;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.annotationEditor = {
    type,
    sheetIndex: state.activeSheetIndex,
    row: anchor.row,
    col: anchor.col,
    beforeText,
  };

  renderSheet();
  updateFormulaBarForCell(anchor.row, anchor.col);
  requestAnimationFrame(() => {
    const textarea = document.querySelector("#cellAnnotationEditor textarea");
    if (!textarea) return;
    textarea.focus({ preventScroll: true });
    textarea.select();
  });
}



function annotationTextFromCell(cell, type) {
  if (!cell) return "";
  if (type === "comment") return normalizeCommentText(cell.comment);
  return normalizeNote(cell.note) || "";
}



function normalizeCommentText(comment) {
  return normalizeCommentReplies(comment)
    .map((reply) => reply.text)
    .filter(Boolean)
    .join("\n");
}



function normalizeCommentReplies(comment) {
  if (!comment) return [];
  if (typeof comment === "string") {
    return comment.trim() ? [{ author: annotationAuthor(), text: comment, createdAt: "" }] : [];
  }
  if (Array.isArray(comment.replies)) {
    return comment.replies
      .map((reply) => ({
        author: String(reply?.author || annotationAuthor()),
        text: String(reply?.text || ""),
        createdAt: String(reply?.createdAt || reply?.updatedAt || ""),
      }))
      .filter((reply) => reply.text.trim());
  }
  if (typeof comment.text === "string" && comment.text.trim()) {
    return [
      {
        author: String(comment.author || annotationAuthor()),
        text: comment.text,
        createdAt: String(comment.createdAt || comment.updatedAt || ""),
      },
    ];
  }
  const text = String(comment);
  return text.trim() ? [{ author: annotationAuthor(), text, createdAt: "" }] : [];
}



function makeCommentFromText(text, existingComment = null) {
  const trimmed = String(text ?? "").replace(/\r\n/g, "\n");
  if (!trimmed.trim()) return null;
  const existingReplies = normalizeCommentReplies(existingComment);
  const first = existingReplies[0] || {};
  return {
    replies: [
      {
        author: first.author || annotationAuthor(),
        text: trimmed,
        createdAt: first.createdAt || new Date().toISOString(),
      },
    ],
    updatedAt: new Date().toISOString(),
  };
}



function appendCommentReply(comment, text) {
  const trimmed = String(text ?? "").replace(/\r\n/g, "\n");
  if (!trimmed.trim()) return comment || null;
  return {
    replies: [
      ...normalizeCommentReplies(comment),
      {
        author: annotationAuthor(),
        text: trimmed,
        createdAt: new Date().toISOString(),
      },
    ],
    updatedAt: new Date().toISOString(),
  };
}



function commentInitial(author) {
  return String(author || annotationAuthor()).trim().charAt(0) || "U";
}



function formatCommentTimestamp(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const pad = (number) => String(number).padStart(2, "0");
  return `${date.getFullYear()}/${pad(date.getMonth() + 1)}/${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}



function annotationHtmlWithBreaks(value) {
  return escapeHtml(value).replace(/\n/g, "<br>");
}



function renderAnnotationEditor(sheet, visibleRange) {
  const editor = state.annotationEditor;
  if (!editor || editor.sheetIndex !== state.activeSheetIndex) return "";
  const range = expandRangeForMerges(sheet, normalizeRange(editor, editor)) || normalizeRange(editor, editor);
  if (!rangeIntersects(range, visibleRange)) return "";

  const cellRight = coordinateToPixels(sheet, "col", range.right);
  const cellTop = coordinateToPixels(sheet, "row", range.top - 1);
  const left = cellRight + 10;
  const top = Math.max(COLUMN_HEADER_HEIGHT + 4, cellTop + 2);
  const cell = sheet.cells[cellKey(editor.row, editor.col)] || {};
  const text = annotationTextFromCell(cell, editor.type);
  const address = `${columnName(editor.col)}${editor.row}`;
  const className = `cell-annotation-editor is-${editor.type}`;
  const style = `left:${left}px;top:${top}px`;

  if (editor.type === "comment") {
    return `
      <div id="cellAnnotationEditor" class="${className}" style="${escapeAttr(style)}" data-annotation-type="comment" role="dialog" aria-label="新しいコメント">
        <span class="cell-annotation-pointer" aria-hidden="true"></span>
        <div class="comment-editor-header">
          <span class="comment-editor-avatar" aria-hidden="true">${escapeHtml(annotationInitial())}</span>
          <div class="comment-editor-title">
            <strong>${escapeHtml(annotationAuthor())}</strong>
            <span>${escapeHtml(address)}</span>
          </div>
          <button class="comment-editor-close" type="button" data-annotation-action="cancel" aria-label="閉じる">×</button>
        </div>
        <textarea class="comment-editor-textarea" rows="4" placeholder="コメントを入力">${escapeHtml(text)}</textarea>
        <div class="comment-editor-actions">
          <button class="comment-editor-button" type="button" data-annotation-action="cancel">キャンセル</button>
          <button class="comment-editor-button is-primary" type="button" data-annotation-action="save">投稿</button>
        </div>
      </div>
    `;
  }

  return `
    <div id="cellAnnotationEditor" class="${className}" style="${escapeAttr(style)}" data-annotation-type="note" role="dialog" aria-label="新しいメモ">
      <span class="cell-annotation-pointer" aria-hidden="true"></span>
      <div class="note-editor-author">${escapeHtml(annotationAuthor())}:</div>
      <textarea class="note-editor-textarea" rows="5">${escapeHtml(text)}</textarea>
    </div>
  `;
}



function annotationAuthor() {
  return "ユーザー";
}



function annotationInitial() {
  return annotationAuthor().trim().charAt(0) || "U";
}



function handleAnnotationEditorAction(event) {
  event.preventDefault();
  event.stopPropagation();
  const action = event.currentTarget.dataset.annotationAction;
  if (action === "save") {
    commitAnnotationEditor();
  } else {
    cancelAnnotationEditor();
  }
}



function handleAnnotationEditorKeydown(event) {
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  if (isElementImeCompositionEvent(event)) return;
  if (event.key === "Escape") {
    event.preventDefault();
    cancelAnnotationEditor();
    return;
  }
  if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
    event.preventDefault();
    commitAnnotationEditor();
  }
}



function closeAnnotationEditorOnEscape(event) {
  if (event.key !== "Escape" || !state.annotationEditor) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  cancelAnnotationEditor();
}



function commitAnnotationEditorOnOutsidePointer(event) {
  if (!state.annotationEditor) return;
  if (event.target.closest?.("#cellAnnotationEditor")) return;
  commitAnnotationEditor();
}



function cancelAnnotationEditor() {
  if (!state.annotationEditor) return;
  state.annotationEditor = null;
  renderSheet();
}



function commitAnnotationEditor() {
  const editor = state.annotationEditor;
  if (!editor || !state.model?.sheets[editor.sheetIndex]) return false;
  const textarea = document.querySelector("#cellAnnotationEditor textarea");
  const text = String(textarea?.value ?? "").replace(/\r\n/g, "\n");
  const sheet = state.model.sheets[editor.sheetIndex];
  const range = normalizeRange(editor, editor);
  const before = captureCellsForHistory(sheet, range);
  const key = cellKey(editor.row, editor.col);
  const previousCell = sheet.cells[key] || null;
  const previousText = annotationTextFromCell(previousCell, editor.type);

  if (text.trim()) {
    const cell = ensureCellModel(sheet, editor.row, editor.col);
    if (editor.type === "comment") {
      cell.comment = makeCommentFromText(text, cell.comment);
    } else {
      cell.note = text;
      cell.noteVisible = Boolean(previousCell?.noteVisible);
    }
  } else if (previousCell) {
    if (editor.type === "comment") {
      delete previousCell.comment;
    } else {
      previousCell.note = null;
      previousCell.noteVisible = false;
    }
    removeCellModelIfEmpty(sheet, editor.row, editor.col);
  }

  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);
  const changed = previousText !== text;
  state.activeSheetIndex = editor.sheetIndex;
  state.selected = { sheetIndex: editor.sheetIndex, row: editor.row, col: editor.col };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.annotationEditor = null;

  if (changed) {
    pushHistory({
      sheetIndex: editor.sheetIndex,
      before,
      after,
      selectionBefore: range,
      selectionAfter: range,
      label: `${columnName(editor.col)}${editor.row} ${editor.type === "comment" ? "コメント" : "メモ"}`,
    });
  }

  renderSheet();
  updateFormulaBarForCell(editor.row, editor.col);
  const label = editor.type === "comment" ? "コメント" : "メモ";
  if (!changed) {
    setStatus(`${columnName(editor.col)}${editor.row} の${label}をキャンセルしました。`);
  } else if (text.trim()) {
    setStatus(`${columnName(editor.col)}${editor.row} に${label}を追加しました。`);
  } else {
    setStatus(`${columnName(editor.col)}${editor.row} の${label}を削除しました。`);
  }
  return true;
}



function removeCellModelIfEmpty(sheet, row, col) {
  const key = cellKey(row, col);
  const cell = sheet.cells[key];
  if (cell && !isPersistentCell(cell)) delete sheet.cells[key];
}



function annotationPointFromMarker(marker) {
  const cell = marker?.closest?.(".sheet-cell[data-row][data-col]");
  if (!cell) return null;
  return {
    row: Number(cell.dataset.row),
    col: Number(cell.dataset.col),
  };
}



function openCommentThreadPopupFromMarker(event) {
  const point = annotationPointFromMarker(event.currentTarget);
  if (!point) return;
  showCommentThreadPopup(point.row, point.col);
}



function showCommentThreadPopup(row, col, options = {}) {
  if (!state.model) return;
  cancelCloseCommentThreadPopup();
  closeNotePopup({ persistentOnly: false });
  const sheet = activeSheet();
  const cell = sheet.cells[cellKey(row, col)];
  const replies = normalizeCommentReplies(cell?.comment);
  if (!replies.length) return closeCommentThreadPopup();

  const popup = document.createElement("div");
  popup.id = "cellCommentThreadPopup";
  popup.className = "cell-comment-thread-popup";
  popup.dataset.row = String(row);
  popup.dataset.col = String(col);
  popup.setAttribute("role", "dialog");
  popup.setAttribute("aria-label", "コメント");
  popup.innerHTML = commentThreadPopupHtml(row, col, replies);
  positionAnnotationPopup(popup, sheet, row, col);
  document.querySelector("#cellCommentThreadPopup")?.remove();
  $sheetHost.find(".sheet-grid")[0]?.append(popup);
  if (options.focusReply) {
    const focusReplyBox = () => {
      const textarea = document.querySelector("#cellCommentThreadPopup textarea");
      textarea?.focus({ preventScroll: true });
      textarea?.select();
    };
    focusReplyBox();
    requestAnimationFrame(focusReplyBox);
  }
}



function commentThreadPopupHtml(row, col, replies) {
  const address = `${columnName(col)}${row}`;
  return `
    <span class="cell-annotation-pointer" aria-hidden="true"></span>
    <div class="comment-thread-header">
      <div>
        <strong>コメント</strong>
        <span>${escapeHtml(address)}</span>
      </div>
      <button class="comment-thread-close" type="button" data-comment-thread-action="close" aria-label="閉じる">×</button>
    </div>
    <div class="comment-thread-list" aria-label="投稿履歴">
      ${replies.map(commentThreadPostHtml).join("")}
    </div>
    <div class="comment-thread-reply">
      <textarea rows="3" placeholder="返信を入力"></textarea>
      <div class="comment-thread-actions">
        <button class="comment-thread-button" type="button" data-comment-thread-action="close">キャンセル</button>
        <button class="comment-thread-button is-primary" type="button" data-comment-thread-action="reply">投稿</button>
      </div>
    </div>
  `;
}



function commentThreadPostHtml(reply) {
  const time = formatCommentTimestamp(reply.createdAt);
  return `
    <article class="comment-thread-post">
      <span class="comment-thread-avatar" aria-hidden="true">${escapeHtml(commentInitial(reply.author))}</span>
      <div class="comment-thread-post-body">
        <div class="comment-thread-meta">
          <strong>${escapeHtml(reply.author || annotationAuthor())}</strong>
          ${time ? `<time>${escapeHtml(time)}</time>` : ""}
        </div>
        <div class="comment-thread-text">${annotationHtmlWithBreaks(reply.text)}</div>
      </div>
    </article>
  `;
}



function positionAnnotationPopup(popup, sheet, row, col) {
  const { left, top, pointerTop } = annotationPopupPosition(sheet, row, col);
  popup.style.left = `${left}px`;
  popup.style.top = `${top}px`;
  popup.style.setProperty("--annotation-pointer-top", `${pointerTop}px`);
}



function annotationPopupPosition(sheet, row, col) {
  const merge = findMergeForCell(sheet, row, col) || { top: row, left: col, bottom: row, right: col };
  const left = coordinateToPixels(sheet, "col", merge.right) + 10;
  const targetTop = coordinateToPixels(sheet, "row", merge.top - 1);
  const top = Math.max(COLUMN_HEADER_HEIGHT + 4, targetTop - ANNOTATION_POINTER_TIP_OFFSET);
  const pointerTop = Math.max(4, targetTop - top - ANNOTATION_POINTER_HALF_HEIGHT);
  return { left, top, pointerTop };
}



function handleCommentThreadAction(event) {
  const popup = event.currentTarget.closest("#cellCommentThreadPopup");
  if (!popup) return;
  event.preventDefault();
  event.stopPropagation();
  const action = event.currentTarget.dataset.commentThreadAction;
  if (action === "reply") {
    appendCommentReplyFromPopup(popup);
  } else {
    closeCommentThreadPopup();
  }
}



function handleCommentThreadKeydown(event) {
  event.stopPropagation();
  event.stopImmediatePropagation?.();
  if (isElementImeCompositionEvent(event)) return;
  if (event.key === "Escape") {
    event.preventDefault();
    closeCommentThreadPopup();
    return;
  }
  if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
    event.preventDefault();
    appendCommentReplyFromPopup(event.currentTarget.closest("#cellCommentThreadPopup"));
  }
}



function appendCommentReplyFromPopup(popup) {
  if (!popup || !state.model) return false;
  const row = Number(popup.dataset.row);
  const col = Number(popup.dataset.col);
  const textarea = popup.querySelector("textarea");
  const text = String(textarea?.value ?? "").replace(/\r\n/g, "\n");
  if (!text.trim()) {
    textarea?.focus();
    return false;
  }
  const sheet = activeSheet();
  const range = normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col });
  const before = captureCellsForHistory(sheet, range);
  const cell = ensureCellModel(sheet, row, col);
  cell.comment = appendCommentReply(cell.comment, text);
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    label: `${columnName(col)}${row} コメント投稿`,
  });
  state.selected = { sheetIndex: state.activeSheetIndex, row, col };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  renderSheet();
  updateFormulaBarForCell(row, col);
  closeCommentThreadPopup();
  setStatus(`${columnName(col)}${row} にコメントを投稿しました。`);
  return true;
}



function replyToCellComment(row, col) {
  const sheet = activeSheet();
  const cell = sheet.cells[cellKey(row, col)];
  if (!cellHasComment(cell)) return;
  const range = normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col });
  state.selected = { sheetIndex: state.activeSheetIndex, row, col };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  renderSheet();
  updateFormulaBarForCell(row, col);
  requestAnimationFrame(() => showCommentThreadPopup(row, col, { focusReply: true }));
  setStatus(`${columnName(col)}${row} のコメントに返信できます。`);
}



function deleteCellComment(row, col) {
  const sheet = activeSheet();
  const cell = sheet.cells[cellKey(row, col)];
  if (!cellHasComment(cell)) return;
  const range = normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col });
  const before = captureCellsForHistory(sheet, range);
  delete cell.comment;
  removeCellModelIfEmpty(sheet, row, col);
  const after = captureCellsForHistory(sheet, range);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    label: `${columnName(col)}${row} コメントの削除`,
  });
  closeCommentThreadPopup();
  renderSheet();
  updateFormulaBarForCell(row, col);
  setStatus(`${columnName(col)}${row} のコメントを削除しました。`);
}



function scheduleCloseCommentThreadPopup() {
  cancelCloseCommentThreadPopup();
  state.commentPopupCloseTimer = window.setTimeout(closeCommentThreadPopup, 180);
}



function cancelCloseCommentThreadPopup() {
  if (!state.commentPopupCloseTimer) return;
  window.clearTimeout(state.commentPopupCloseTimer);
  state.commentPopupCloseTimer = null;
}



function closeCommentThreadPopup() {
  cancelCloseCommentThreadPopup();
  document.querySelector("#cellCommentThreadPopup")?.remove();
}



function openNotePopupFromMarker(event) {
  const point = annotationPointFromMarker(event.currentTarget);
  if (!point) return;
  showNotePopup(point.row, point.col);
}



function showNotePopup(row, col, options = {}) {
  if (!state.model) return;
  cancelCloseNotePopup();
  const sheet = activeSheet();
  const cell = sheet.cells[cellKey(row, col)];
  const text = normalizeNote(cell?.note);
  if (!text) return closeNotePopup({ persistentOnly: false });
  const popup = document.createElement("div");
  popup.id = "cellNotePopup";
  popup.className = `cell-note-popup${options.persistent ? " is-persistent" : ""}`;
  popup.dataset.row = String(row);
  popup.dataset.col = String(col);
  popup.innerHTML = notePopupHtml(text);
  positionAnnotationPopup(popup, sheet, row, col);
  closeNotePopup({ persistentOnly: false });
  $sheetHost.find(".sheet-grid")[0]?.append(popup);
}



function notePopupHtml(text) {
  return `
    <span class="cell-annotation-pointer" aria-hidden="true"></span>
    <div class="note-popup-author">${escapeHtml(annotationAuthor())}:</div>
    <div class="note-popup-text">${annotationHtmlWithBreaks(text)}</div>
  `;
}



function persistentNoteCellsForSheet(sheet) {
  const revision = Math.max(0, Number(state.workbookRevision) || 0);
  const cached = sheetPersistentNoteCellCacheBySheet.get(sheet);
  if (cached?.revision === revision && cached.cells === sheet.cells) return cached.cellsWithNotes;
  const cellsWithNotes = Object.values(sheet.cells || {}).filter((cell) => cell?.noteVisible && cellHasNote(cell));
  sheetPersistentNoteCellCacheBySheet.set(sheet, { revision, cells: sheet.cells, cellsWithNotes });
  return cellsWithNotes;
}



function renderPersistentNotePopups(sheet, visibleRange) {
  return persistentNoteCellsForSheet(sheet)
    .map((cell) => {
      const row = Number(cell.row);
      const col = Number(cell.col);
      const range = expandRangeForMerges(sheet, normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col }));
      if (!range || !rangeIntersects(range, visibleRange)) return "";
      const { left, top, pointerTop } = annotationPopupPosition(sheet, row, col);
      const style = `left:${left}px;top:${top}px;--annotation-pointer-top:${pointerTop}px`;
      return `<div id="cellNotePopup" class="cell-note-popup is-persistent" style="${escapeAttr(style)}" data-row="${row}" data-col="${col}">${notePopupHtml(normalizeNote(cell.note))}</div>`;
    })
    .join("");
}



function scheduleCloseNotePopup() {
  cancelCloseNotePopup();
  state.notePopupCloseTimer = window.setTimeout(() => closeNotePopup({ persistentOnly: false }), 180);
}



function cancelCloseNotePopup() {
  if (!state.notePopupCloseTimer) return;
  window.clearTimeout(state.notePopupCloseTimer);
  state.notePopupCloseTimer = null;
}



function closeNotePopup(options = {}) {
  cancelCloseNotePopup();
  const selector = options.persistentOnly ? "#cellNotePopup.is-persistent" : "#cellNotePopup:not(.is-persistent)";
  document.querySelectorAll(selector).forEach((popup) => popup.remove());
}



function deleteCellNote(row, col) {
  const sheet = activeSheet();
  const cell = sheet.cells[cellKey(row, col)];
  if (!cellHasNote(cell)) return;
  const range = normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col });
  const before = captureCellsForHistory(sheet, range);
  cell.note = null;
  cell.noteVisible = false;
  removeCellModelIfEmpty(sheet, row, col);
  const after = captureCellsForHistory(sheet, range);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    label: `${columnName(col)}${row} メモの削除`,
  });
  renderSheet();
  updateFormulaBarForCell(row, col);
  setStatus(`${columnName(col)}${row} のメモを削除しました。`);
}



function toggleCellNoteVisibility(row, col) {
  const sheet = activeSheet();
  const cell = sheet.cells[cellKey(row, col)];
  if (!cellHasNote(cell)) return;
  const range = normalizeRange({ sheetIndex: state.activeSheetIndex, row, col }, { sheetIndex: state.activeSheetIndex, row, col });
  const before = captureCellsForHistory(sheet, range);
  cell.noteVisible = !cell.noteVisible;
  const after = captureCellsForHistory(sheet, range);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    label: `${columnName(col)}${row} メモの表示/非表示`,
  });
  renderSheet();
  updateFormulaBarForCell(row, col);
  setStatus(cell.noteVisible ? `${columnName(col)}${row} のメモを表示しました。` : `${columnName(col)}${row} のメモを非表示にしました。`);
}



function toggleSelectedPhonetics() {
  if (!state.model || !state.selected) {
    setStatus("ふりがなを表示するセルを選択してください。");
    return;
  }
  const sheet = activeSheet();
  const ranges = activeSelectionRanges();
  const targetKeys = new Set();
  ranges.forEach((range) => {
    forEachCellInRange(range, (row, col) => {
      const key = cellKey(row, col);
      if (cellHasPhonetics(sheet.cells[key])) targetKeys.add(key);
    });
  });
  if (!targetKeys.size) {
    setStatus("選択範囲にふりがながありません。");
    return;
  }

  const shouldShow = Array.from(targetKeys).some((key) => !sheet.cells[key]?.showPhonetic);
  const before = captureCellsForHistoryRanges(sheet, ranges);
  targetKeys.forEach((key) => {
    sheet.cells[key].showPhonetic = shouldShow;
  });
  const after = captureCellsForHistoryRanges(sheet, ranges);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: state.selectionRange,
    selectionAfter: state.selectionRange,
    selectionBeforeRanges: (state.selectionRanges || []).map((range) => ({ ...range })),
    selectionAfterRanges: (state.selectionRanges || []).map((range) => ({ ...range })),
    label: `${rangesToLabel(ranges)} ふりがなの${shouldShow ? "表示" : "非表示"}`,
  });
  renderSheet();
  updateFormulaBarForCell(state.selected.row, state.selected.col);
  setStatus(`${rangesToLabel(ranges)} のふりがなを${shouldShow ? "表示" : "非表示"}しました。`);
}



function cellHasPhonetics(cell) {
  return Boolean(cell?.phonetics?.length);
}



async function handleCellContextMenuAction(event) {
  const button = event.currentTarget;
  if (!button.closest("#cellContextMenu")) return;
  event.preventDefault();
  event.stopPropagation();
  const action = button.dataset.cellMenuAction;
  const row = Number(button.dataset.row ?? state.cellContextMenu?.row);
  const col = Number(button.dataset.col ?? state.cellContextMenu?.col);
  const context = state.cellContextMenu ? { ...state.cellContextMenu } : null;
  closeCellContextMenu();
  await runCellContextMenuAction(action, row, col, context);
}



async function runCellContextMenuAction(action, row, col, context = null) {
  if (!state.model) return;
  if (row && col && (!state.selected || state.selected.sheetIndex !== state.activeSheetIndex)) {
    state.selected = { sheetIndex: state.activeSheetIndex, row, col };
    state.selectionRange = normalizeRange(state.selected, state.selected);
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectionDrag = null;
  }

  switch (action) {
    case "object-cut":
      await cutSelectedImages({ systemClipboard: true });
      break;
    case "object-copy":
      await copySelectedImages({ systemClipboard: true });
      break;
    case "object-paste":
      pasteObjectClipboard();
      break;
    case "object-edit-text":
      openShapeTextEditor();
      break;
    case "object-edit-points":
      enterShapePointEditMode();
      break;
    case "object-add-vertex":
      addCustomShapeVertex(context?.imageId, context?.vertexKey);
      break;
    case "object-delete-vertex":
      deleteCustomShapeVertex(context?.imageId, context?.vertexKey);
      break;
    case "object-toggle-open-path":
      toggleCustomShapePathClosed(context?.imageId);
      break;
    case "object-delete":
      deleteSelectedImages();
      break;
    case "object-bring-front":
      reorderSelectedObjects("front");
      break;
    case "object-bring-forward":
      reorderSelectedObjects("forward");
      break;
    case "object-send-backward":
      reorderSelectedObjects("backward");
      break;
    case "object-send-back":
      reorderSelectedObjects("back");
      break;
    case "object-selection-pane":
      openShapeSelectionPane();
      break;
    case "object-group":
      groupSelectedShapes();
      break;
    case "object-ungroup":
      ungroupSelectedShapes();
      break;
    case "object-link":
      openObjectLinkDialog();
      break;
    case "object-open-link":
      openSelectedObjectHyperlink();
      break;
    case "object-format":
      openShapeFormatPane();
      break;
    case "cut":
      executeWorkbookClipboardShortcut("cut", () => void cutSelectedRange({ systemClipboard: true }));
      break;
    case "copy":
      executeWorkbookClipboardShortcut("copy", () => void copySelectedRange({ systemClipboard: true }));
      break;
    case "paste-keep-format":
      await pasteFromClipboard({ rich: true, allowAsyncClipboard: true, allowAsyncClipboardImages: true });
      break;
    case "paste-match-format":
      await pasteFromClipboard({ rich: false, allowAsyncClipboard: true, allowAsyncClipboardImages: true });
      break;
    case "clear":
      clearSelectionContents();
      break;
    case "open-hyperlink":
      openSelectedHyperlink(row, col);
      break;
    case "insert":
      openCellStructureDialog("insert", row, col);
      break;
    case "delete":
      openCellStructureDialog("delete", row, col);
      break;
    case "insert-entire-column":
      applyHeaderStructureOperation("insert", "entire-column", row, col);
      break;
    case "delete-entire-column":
      applyHeaderStructureOperation("delete", "entire-column", row, col);
      break;
    case "insert-entire-row":
      applyHeaderStructureOperation("insert", "entire-row", row, col);
      break;
    case "insert-copied-column":
      insertCopiedHeaderCells("column", row, col);
      break;
    case "insert-copied-row":
      insertCopiedHeaderCells("row", row, col);
      break;
    case "delete-entire-row":
      applyHeaderStructureOperation("delete", "entire-row", row, col);
      break;
    case "hide-column":
      setHeaderHidden("column", true, row, col);
      break;
    case "unhide-column":
      setHeaderHidden("column", false, row, col);
      break;
    case "hide-row":
      setHeaderHidden("row", true, row, col);
      break;
    case "unhide-row":
      setHeaderHidden("row", false, row, col);
      break;
    case "column-width":
      openHeaderSizeDialog("column", row, col);
      break;
    case "row-height":
      openHeaderSizeDialog("row", row, col);
      break;
    case "insert-page-break-selection":
      insertPageBreakFromSelection();
      break;
    case "remove-page-break-selection":
      removePageBreakFromSelection();
      break;
    case "reset-page-breaks":
      resetAllPageBreaks();
      break;
    case "close-page-break-preview":
      setActiveSheetView("normal");
      break;
    case "translate":
      openTranslationPane(row, col);
      break;
    case "quick-analysis":
      openQuickAnalysisPopup(row, col);
      break;
    case "filter":
      toggleAutoFilterForSelection();
      break;
    case "sort":
      openSortDialog();
      break;
    case "get-data":
      openGetDataFromTableRangeDialog();
      break;
    case "new-comment":
      openCellAnnotationEditor("comment", row, col);
      break;
    case "reply-comment":
      replyToCellComment(row, col);
      break;
    case "delete-comment":
      deleteCellComment(row, col);
      break;
    case "new-note":
      openCellAnnotationEditor("note", row, col);
      break;
    case "edit-note":
      openCellAnnotationEditor("note", row, col);
      break;
    case "delete-note":
      deleteCellNote(row, col);
      break;
    case "toggle-note":
      toggleCellNoteVisibility(row, col);
      break;
    case "format-cells":
      openFormatCellsDialog(row, col);
      break;
    case "select-dropdown":
      openValidationDropdownPopup(row, col);
      break;
    case "furigana":
      toggleSelectedPhonetics();
      break;
    case "define-name":
      openDefineNameDialog(row, col);
      break;
    case "link":
      openLinkDialog(row, col);
      break;
    default:
      break;
  }
}



function applyHeaderStructureOperation(kind, option, row, col) {
  const fallbackPoint = { sheetIndex: state.activeSheetIndex, row: row || 1, col: col || 1 };
  const range = activeSelectionRange() || selectionRangeFromPoints(fallbackPoint, fallbackPoint);
  if (range) applyCellStructureOperation(kind, option, range);
}



function insertCopiedHeaderCells(kind, row, col) {
  const clipboard = state.internalClipboard;
  const sourceRange = clipboard?.sourceRange;
  const isColumn = kind === "column";
  if (!clipboard || (isColumn ? !sourceRange?.fullColumn : !sourceRange?.fullRow)) {
    showMessageBox("コピーした行または列がありません。", { title: "Cht WebSheet" });
    return;
  }

  const sheet = activeSheet();
  const currentRange = activeSelectionRange();
  const targetIndex = isColumn
    ? (currentRange?.fullColumn ? currentRange.left : col || 1)
    : (currentRange?.fullRow ? currentRange.top : row || 1);
  const insertAmount = isColumn ? clipboard.colCount : clipboard.rowCount;
  if (!Number.isFinite(targetIndex) || targetIndex < 1 || insertAmount < 1) return;
  const isCutMove = clipboard.mode === "cut" && clipboard.sourceSheetIndex === state.activeSheetIndex;
  if (isCutMove && targetOverlapsCutSource(kind, sourceRange, targetIndex)) {
    showMessageBox("切り取った行または列の中には挿入できません。", { title: "Cht WebSheet" });
    return;
  }

  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = currentRange ? { ...currentRange } : null;
  const selectionBeforeRanges = activeSelectionRanges().map((range) => ({ ...range }));

  try {
    if (isColumn) {
      ensureSheetSize(sheet, Math.max(sheet.rowCount, clipboard.rowCount), sheet.colCount);
      insertEntireColumns(sheet, targetIndex, insertAmount);
      applyClipboardColumnModels(sheet, clipboard, targetIndex);
      writeClipboardCellsIntoSheet(sheet, 1, targetIndex, clipboard);
      if (isCutMove) {
        deleteEntireColumns(sheet, adjustedCutSourceIndexAfterInsert("column", sourceRange, targetIndex, insertAmount), insertAmount);
      }
    } else {
      ensureSheetSize(sheet, sheet.rowCount, Math.max(sheet.colCount, clipboard.colCount));
      insertEntireRows(sheet, targetIndex, insertAmount);
      applyClipboardRowModels(sheet, clipboard, targetIndex);
      writeClipboardCellsIntoSheet(sheet, targetIndex, 1, clipboard);
      if (isCutMove) {
        deleteEntireRows(sheet, adjustedCutSourceIndexAfterInsert("row", sourceRange, targetIndex, insertAmount), insertAmount);
      }
    }
    const finalTargetIndex = isCutMove && targetIndex > cutSourceStart(kind, sourceRange)
      ? targetIndex - insertAmount
      : targetIndex;
    const selectionAfter = isColumn
      ? {
          sheetIndex: state.activeSheetIndex,
          top: 1,
          left: finalTargetIndex,
          bottom: sheet.rowCount,
          right: finalTargetIndex + insertAmount - 1,
          fullColumn: true,
        }
      : {
          sheetIndex: state.activeSheetIndex,
          top: finalTargetIndex,
          left: 1,
          bottom: finalTargetIndex + insertAmount - 1,
          right: sheet.colCount,
          fullRow: true,
        };
    normalizeSheetModelAfterStructureChange(sheet);
    state.hf = buildFormulaEngine(state.model);
    state.selected = isColumn
      ? { sheetIndex: state.activeSheetIndex, row: 1, col: finalTargetIndex }
      : { sheetIndex: state.activeSheetIndex, row: finalTargetIndex, col: 1 };
    state.selectionRange = selectionAfter;
    state.selectionRanges = [selectionAfter];
    state.selectionStart = state.selected;
    state.selectionDrag = null;
    state.selectedImage = null;
    state.formulaEdit = null;
    if (isCutMove) state.internalClipboard = null;

    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetBefore,
      sheetAfter: captureSheetForHistory(sheet),
      selectionBefore,
      selectionAfter,
      selectionBeforeRanges,
      selectionAfterRanges: [selectionAfter],
      label: `${rangeToLabel(selectionAfter)} コピーしたセルの挿入`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${rangeToLabel(selectionAfter)} に${isCutMove ? "切り取った" : "コピーした"}セルを挿入しました。`);
  } catch (error) {
    console.error(error);
    restoreSheetFromHistory(sheet, sheetBefore);
    state.hf = buildFormulaEngine(state.model);
    renderSheet();
    showMessageBox("コピーしたセルを挿入できません。", { title: "Cht WebSheet" });
    setStatus(`コピーしたセルの挿入に失敗しました：${error.message || error}`);
  }
}



function cutSourceStart(kind, sourceRange) {
  return kind === "column" ? Number(sourceRange?.left || 0) : Number(sourceRange?.top || 0);
}



function cutSourceEnd(kind, sourceRange) {
  return kind === "column" ? Number(sourceRange?.right || 0) : Number(sourceRange?.bottom || 0);
}



function targetOverlapsCutSource(kind, sourceRange, targetIndex) {
  return targetIndex >= cutSourceStart(kind, sourceRange) && targetIndex <= cutSourceEnd(kind, sourceRange);
}



function adjustedCutSourceIndexAfterInsert(kind, sourceRange, targetIndex, insertAmount) {
  const sourceStart = cutSourceStart(kind, sourceRange);
  return targetIndex < sourceStart ? sourceStart + insertAmount : sourceStart;
}



function writeClipboardCellsIntoSheet(sheet, startRow, startCol, clipboard) {
  const targetRange = {
    sheetIndex: state.activeSheetIndex,
    top: startRow,
    left: startCol,
    bottom: startRow + clipboard.rowCount - 1,
    right: startCol + clipboard.colCount - 1,
  };
  removeContainedMerges(sheet, targetRange);
  clearCellsInRange(sheet, targetRange);
  clipboardCellEntries(clipboard).forEach((sourceCell) => {
    writeClonedClipboardCellContent(
      sheet,
      startRow + sourceCell.row - 1,
      startCol + sourceCell.col - 1,
      clipboard,
      sourceCell,
      { row: sourceCell.row, col: sourceCell.col },
    );
  });
  clipboard.merges?.forEach((merge) => {
    const nextMerge = {
      top: startRow + merge.top - 1,
      left: startCol + merge.left - 1,
      bottom: startRow + merge.bottom - 1,
      right: startCol + merge.right - 1,
    };
    if (nextMerge.top <= nextMerge.bottom && nextMerge.left <= nextMerge.right && isRangeInsideRange(nextMerge, targetRange)) {
      sheet.merges.push(nextMerge);
    }
  });
}



function clipboardCellEntries(clipboard) {
  if (Array.isArray(clipboard?.cellEntries)) return clipboard.cellEntries;
  const entries = [];
  (clipboard?.cells || []).forEach((sourceRow, rowOffset) => {
    (sourceRow || []).forEach((sourceCell, colOffset) => {
      if (!sourceCell) return;
      const cell = cloneCellModel(sourceCell);
      cell.row = Number(cell.row) || rowOffset + 1;
      cell.col = Number(cell.col) || colOffset + 1;
      entries.push(cell);
    });
  });
  return entries;
}



function clipboardHasFormula(clipboard) {
  return clipboardCellEntries(clipboard).some((cell) => cellHasFormulaContent(cell));
}



function rangeHasFormulaContent(sheet, range) {
  if (!sheet || !range) return false;
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      if (cellHasFormulaContent(sheet.cells?.[cellKey(row, col)])) return true;
    }
  }
  return false;
}



function cellHasFormulaContent(cell) {
  if (!cell) return false;
  if (typeof cell.formula === "string" && cell.formula.trim().startsWith("=")) return true;
  return typeof cell.raw === "string" && cell.raw.trim().startsWith("=");
}



function clearCellsInRange(sheet, range) {
  forEachCellInRange(range, (row, col) => {
    delete sheet.cells[cellKey(row, col)];
  });
}



function removeContainedMerges(sheet, range) {
  sheet.merges = (sheet.merges || []).filter((merge) => !isRangeInsideRange(merge, range));
}



function isRangeInsideRange(inner, outer) {
  return inner.top >= outer.top && inner.left >= outer.left && inner.bottom <= outer.bottom && inner.right <= outer.right;
}



function applyClipboardRowModels(sheet, clipboard, startRow) {
  (clipboard.rows || []).forEach((rowModel, offset) => {
    sheet.rows[startRow + offset - 1] = { ...defaultRowModel(sheet), ...(rowModel || {}) };
  });
}



function applyClipboardColumnModels(sheet, clipboard, startCol) {
  (clipboard.columns || []).forEach((columnModel, offset) => {
    sheet.columns[startCol + offset - 1] = { ...defaultColumnModel(sheet), ...(columnModel || {}) };
  });
}



function applyHeaderHiddenToSheet(sheet, kind, ranges, hidden) {
  let changed = 0;
  ranges.forEach((range) => {
    if (kind === "column") {
      for (let index = range.left; index <= range.right; index += 1) {
        const column = sheet.columns[index - 1];
        if (column && column.hidden !== hidden) {
          column.hidden = hidden;
          changed += 1;
        }
      }
    } else {
      for (let index = range.top; index <= range.bottom; index += 1) {
        const targetRow = sheet.rows[index - 1];
        if (targetRow && targetRow.hidden !== hidden) {
          targetRow.hidden = hidden;
          targetRow.height = hidden ? 0 : targetRow.height || sheet.defaultRowHeight || rowHeightToPixels(15);
          changed += 1;
        }
      }
    }
  });
  return changed;
}



function setHeaderHiddenForSheetGroup(kind, hidden, ranges) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !ranges?.length) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureHeaderFormatAllowedForSheet(sheet, kind)) return true;
    const targetRanges = materializedRangesForSheet(sheet, ranges, sheetIndex);
    if (targetRanges.length) items.push({ sheetIndex, sheet, ranges: targetRanges });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  let changed = 0;
  items.forEach((item) => {
    changed += applyHeaderHiddenToSheet(item.sheet, kind, item.ranges, hidden);
  });
  if (!changed) {
    setStatus(hidden ? "対象はすでに非表示です。" : "対象はすでに表示されています。");
    return true;
  }
  state.hf = buildFormulaEngine(state.model);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: ranges.map((range) => ({ ...range })),
    selectionAfterRanges: ranges.map((range) => ({ ...range })),
    label: `${rangesToLabel(ranges)} ${sheetIndexes.length} シート${hidden ? "非表示" : "再表示"}`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`${sheetIndexes.length} 個のシートの${rangesToLabel(ranges)}を${hidden ? "非表示" : "再表示"}にしました。`);
  return true;
}



function setHeaderHidden(kind, hidden, row, col) {
  if (!state.model) return;
  if (!ensureHeaderFormatAllowed(kind)) return;
  const sheet = activeSheet();
  const fallbackPoint = { sheetIndex: state.activeSheetIndex, row: row || 1, col: col || 1 };
  const ranges = activeSelectionRanges().length ? activeSelectionRanges() : [selectionRangeFromPoints(fallbackPoint, fallbackPoint)].filter(Boolean);
  if (!ranges.length) return;
  if (setHeaderHiddenForSheetGroup(kind, hidden, ranges)) return;

  const sheetBefore = captureSheetForHistory(sheet);
  const selectionBefore = activeSelectionRange();
  const changed = applyHeaderHiddenToSheet(sheet, kind, ranges, hidden);
  if (!changed) {
    setStatus(hidden ? "対象はすでに非表示です。" : "対象はすでに表示されています。");
    return;
  }
  state.hf = buildFormulaEngine(state.model);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore,
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: ranges,
    selectionAfterRanges: ranges,
    label: `${rangesToLabel(ranges)} ${hidden ? "非表示" : "再表示"}`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`${rangesToLabel(ranges)} を${hidden ? "非表示" : "再表示"}にしました。`);
}



function openHeaderSizeDialog(kind, row, col) {
  const sheet = activeSheet();
  if (!ensureHeaderFormatAllowed(kind)) return;
  const headerIndex = kind === "column" ? col || 1 : row || 1;
  const ranges = selectedHeaderRangesForKind(kind, headerIndex);
  const firstRange = ranges[0];
  if (!sheet || !firstRange) return;
  const currentIndex = kind === "column" ? headerIndex || firstRange.left : headerIndex || firstRange.top;
  const currentSize = kind === "column"
    ? sheet.columns[currentIndex - 1]?.width || sheet.defaultColumnWidth || columnWidthToPixels(8.43)
    : sheet.rows[currentIndex - 1]?.height || sheet.defaultRowHeight || rowHeightToPixels(15);

  closeCellContextMenu();
  closeSheetTabContextMenu();
  closeFormatCellsDialog();
  closeCellStructureDialog();
  closeNameDialog();
  closeLinkDialog();
  closeFileSideMenu({ restoreHome: true });
  closeHeaderSizeDialog({ restoreFocus: false });

  state.headerSizeDialog = {
    previousFocus: document.activeElement,
    kind,
    ranges: ranges.map((range) => ({ ...range })),
  };

  const overlay = document.createElement("div");
  overlay.id = "headerSizeDialogOverlay";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = headerSizeDialogHtml(kind, currentSize);
  document.body.append(overlay);
  const input = overlay.querySelector("[data-header-size-input]");
  input?.focus({ preventScroll: true });
  input?.select?.();
}



function headerSizeDialogHtml(kind, currentSize) {
  const isColumn = kind === "column";
  const title = isColumn ? "列の幅" : "行の高さ";
  const caption = isColumn ? "列の幅(C):" : "行の高さ(R):";
  return `
    <div id="headerSizeDialog" class="excel-mini-dialog header-size-dialog" role="dialog" aria-modal="true" aria-labelledby="headerSizeDialogTitle">
      <div class="excel-mini-titlebar" id="headerSizeDialogTitle">${title}</div>
      <div class="excel-mini-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml(caption)}</span>
          <input type="text" inputmode="decimal" data-header-size-input value="${escapeAttr(String(Math.round(currentSize)))}" />
        </label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-header-size-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-header-size-action="cancel">キャンセル</button>
      </div>
    </div>
  `;
}



function handleHeaderSizeDialogAction(event) {
  const action = event.currentTarget.dataset.headerSizeAction;
  if (action === "cancel") {
    closeHeaderSizeDialog();
    return;
  }
  applyHeaderSizeDialog();
}



function handleHeaderSizeDialogKeydown(event) {
  if (event.key !== "Enter" && event.key !== "Escape") return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  if (event.key === "Escape") {
    closeHeaderSizeDialog();
    return;
  }
  applyHeaderSizeDialog();
}



function applyHeaderSizeDialog() {
  const dialog = document.querySelector("#headerSizeDialog");
  const dialogState = state.headerSizeDialog;
  if (!dialog || !dialogState) return false;
  const input = dialog.querySelector("[data-header-size-input]");
  const nextSize = Math.round(Number(input?.value));
  if (!Number.isFinite(nextSize) || nextSize <= 0) {
    showMessageBox("数値を入力してください。", { title: "Cht WebSheet" });
    input?.focus({ preventScroll: true });
    input?.select?.();
    return false;
  }
  const min = dialogState.kind === "column" ? 16 : 8;
  const max = dialogState.kind === "column" ? 640 : 360;
  applyHeaderSize(dialogState.kind, dialogState.ranges, Math.max(min, Math.min(max, nextSize)));
  closeHeaderSizeDialog({ restoreFocus: false });
  return true;
}



function closeHeaderSizeDialog(options = {}) {
  document.querySelector("#headerSizeDialogOverlay")?.remove();
  const previousFocus = state.headerSizeDialog?.previousFocus;
  state.headerSizeDialog = null;
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) {
    previousFocus.focus?.({ preventScroll: true });
  }
}



function closeHeaderSizeDialogOnEscape(event) {
  if (event.key !== "Escape" || !state.headerSizeDialog) return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  closeHeaderSizeDialog();
}



function applyHeaderSizeToSheet(sheet, kind, ranges, size) {
  let changed = 0;
  ranges.forEach((range) => {
    if (kind === "column") {
      for (let index = range.left; index <= range.right; index += 1) {
        const column = sheet.columns[index - 1];
        if (column && column.width !== size) {
          column.width = size;
          column.hidden = false;
          changed += 1;
        }
      }
    } else {
      for (let index = range.top; index <= range.bottom; index += 1) {
        const targetRow = sheet.rows[index - 1];
        if (targetRow && targetRow.height !== size) {
          targetRow.height = size;
          targetRow.hidden = false;
          changed += 1;
        }
      }
    }
  });
  return changed;
}



function applyHeaderSizeToSheetGroup(kind, ranges, size) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !ranges?.length) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureHeaderFormatAllowedForSheet(sheet, kind)) return true;
    const targetRanges = materializedRangesForSheet(sheet, ranges, sheetIndex);
    if (targetRanges.length) items.push({ sheetIndex, sheet, ranges: targetRanges });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  let changed = 0;
  items.forEach((item) => {
    changed += applyHeaderSizeToSheet(item.sheet, kind, item.ranges, size);
  });
  if (!changed) return true;
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: ranges.map((range) => ({ ...range })),
    selectionAfterRanges: ranges.map((range) => ({ ...range })),
    label: `${rangesToLabel(ranges)} ${sheetIndexes.length} シート${kind === "column" ? "列の幅" : "行の高さ"}`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`${sheetIndexes.length} 個のシートの${rangesToLabel(ranges)}の${kind === "column" ? "列の幅" : "行の高さ"}を ${size}px にしました。`);
  return true;
}



function applyHeaderSize(kind, ranges, size) {
  if (!ensureHeaderFormatAllowed(kind)) return;
  const sheet = activeSheet();
  if (applyHeaderSizeToSheetGroup(kind, ranges, size)) return;
  const sheetBefore = captureSheetForHistory(sheet);
  const changed = applyHeaderSizeToSheet(sheet, kind, ranges, size);
  if (!changed) return;
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: ranges,
    selectionAfterRanges: ranges,
    label: `${rangesToLabel(ranges)} ${kind === "column" ? "列の幅" : "行の高さ"}`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`${rangesToLabel(ranges)} の${kind === "column" ? "列の幅" : "行の高さ"}を ${size}px にしました。`);
}



function openFormatCellsDialog(row, col) {
  if (!state.model) return;
  const fallbackPoint = row && col ? { sheetIndex: state.activeSheetIndex, row, col } : state.selected;
  const activeRange = activeSelectionRange() || selectionRangeFromPoints(fallbackPoint, fallbackPoint);
  const ranges = activeSelectionRanges();
  const formatRanges = ranges.length ? ranges : activeRange ? [activeRange] : [];
  if (!formatRanges.length || !activeRange) return;

  closeFormatCellsDialog();
  closeCellContextMenu();
  closeSheetTabContextMenu();
  closeFileSideMenu({ restoreHome: true });

  const point = selectionToolbarPoint() || fallbackPoint || { row: activeRange.top, col: activeRange.left };
  const cell = activeSheet().cells[cellKey(point.row, point.col)] || {};
  state.formatDialog = {
    range: { ...activeRange },
    ranges: formatRanges.map((range) => ({ ...range })),
    point: { ...point },
    activeTab: "number",
  };

  const overlay = document.createElement("div");
  overlay.id = "formatCellsDialogOverlay";
  overlay.className = "format-dialog-overlay";
  overlay.innerHTML = formatCellsDialogHtml(cell, activeSheet());
  document.body.append(overlay);
  updateFormatDialogNumberSettingVisibility(overlay);
  updateFormatDialogPreview();
  overlay.querySelector("[data-format-dialog-action='ok']")?.focus({ preventScroll: true });
}



function formatCellsDialogHtml(cell, sheet = activeSheet()) {
  const css = cellEffectiveCss(sheet, cell);
  const tabLabels = [
    ["number", "表示形式"],
    ["alignment", "配置"],
    ["font", "フォント"],
    ["border", "罫線"],
    ["fill", "塗りつぶし"],
    ["protection", "保護"],
  ];
  const currentFont = fontOptionFromCss(css.fontFamily) || "Calibri";
  const currentFontSize = fontSizeOptionFromCss(css.fontSize) || "11";
  const currentFill = normalizeColorValue(css.backgroundColor) || "#FFFFFF";
  const currentColor = normalizeColorValue(css.color) || "#000000";
  const protection = cell.protection || {};
  const currentNumberCategory = numberCategoryFromFormat(cell.numFmt);
  const currentDateFormat = currentNumberCategory === "date" && cell.numFmt ? cell.numFmt : "yyyy/m/d";
  const currentCustomFormat = cell.numFmt || "yyyy/m/d";

  return `
    <div id="formatCellsDialog" class="format-dialog" role="dialog" aria-modal="true" aria-labelledby="formatCellsDialogTitle">
      <div class="format-dialog-titlebar">
        <div class="format-dialog-title" id="formatCellsDialogTitle">セルの書式設定</div>
      </div>
      <div class="format-dialog-tabs" role="tablist" aria-label="セルの書式設定">
        ${tabLabels.map(([id, label], index) => `<button class="format-dialog-tab${index === 0 ? " is-active" : ""}" type="button" role="tab" data-format-dialog-tab="${id}">${label}</button>`).join("")}
      </div>
      <div class="format-dialog-body">
        <section class="format-dialog-panel is-active" data-format-panel="number">
          <div class="format-number-layout">
            <fieldset class="format-group format-category-group">
              <legend>${accessKeyLabelHtml("分類(C)")}</legend>
              <select class="format-listbox" size="8" data-format-field="numberCategory">
                ${formatOption("general", "標準", currentNumberCategory === "general")}
                ${formatOption("number", "数値", currentNumberCategory === "number")}
                ${formatOption("currency", "通貨", currentNumberCategory === "currency")}
                ${formatOption("percent", "パーセンテージ", currentNumberCategory === "percent")}
                ${formatOption("date", "日付", currentNumberCategory === "date")}
                ${formatOption("text", "文字列", currentNumberCategory === "text")}
              </select>
            </fieldset>
            <div class="format-number-details">
              <fieldset class="format-group">
                <legend>サンプル</legend>
                <div class="format-sample"><strong data-format-sample>1234</strong></div>
              </fieldset>
              <fieldset class="format-group" data-format-settings-group>
                <legend>設定</legend>
                <div class="format-dialog-grid two">
                  <label class="format-field" data-format-number-setting="decimals">${accessKeyLabelHtml("小数点以下の桁数(D)")}
                    <input type="number" min="0" max="6" value="${formatDecimalPlaces(cell.numFmt)}" data-format-field="decimals" />
                  </label>
                  <label class="format-check" data-format-number-setting="useSeparator"><input type="checkbox" ${formatUsesSeparator(cell.numFmt) ? "checked" : ""} data-format-field="useSeparator" /> 桁区切り(,)を使用する</label>
                  <label class="format-field" data-format-number-setting="currencySymbol">${accessKeyLabelHtml("記号(S)")}
                    <select data-format-field="currencySymbol">
                      ${["¥", "$", "€", "£"].map((symbol) => formatOption(symbol, symbol, currencySymbol(cell.numFmt) === symbol)).join("")}
                    </select>
                  </label>
                  <label class="format-field" data-format-number-setting="dateFormat">${accessKeyLabelHtml("日付の種類(T)")}
                    <select data-format-field="dateFormat">
                      ${formatDialogDateFormatOptions(currentDateFormat)}
                    </select>
                  </label>
                  <label class="format-field" data-format-number-setting="customFormat">${accessKeyLabelHtml("ユーザー定義(U)")}
                    <input type="text" value="${escapeAttr(currentCustomFormat)}" data-format-field="customFormat" placeholder="yyyy/m/d h:mm" />
                  </label>
                </div>
              </fieldset>
            </div>
          </div>
        </section>
        <section class="format-dialog-panel" data-format-panel="alignment">
          <fieldset class="format-group">
            <legend>文字の配置</legend>
            <div class="format-dialog-grid two">
              <label class="format-field">${accessKeyLabelHtml("横位置(H)")}
                <select data-format-field="horizontal">
                  ${formatOption("", "標準", !css.textAlign)}
                  ${formatOption("left", "左詰め", css.textAlign === "left")}
                  ${formatOption("center", "中央揃え", css.textAlign === "center")}
                  ${formatOption("right", "右詰め", css.textAlign === "right")}
                  ${formatOption("justify", "両端揃え", css.textAlign === "justify")}
                </select>
              </label>
              <label class="format-field">${accessKeyLabelHtml("縦位置(V)")}
                <select data-format-field="vertical">
                  ${formatOption("top", "上詰め", css.verticalAlign === "top")}
                  ${formatOption("middle", "中央揃え", !css.verticalAlign || css.verticalAlign === "middle")}
                  ${formatOption("bottom", "下詰め", css.verticalAlign === "bottom")}
                </select>
              </label>
            </div>
          </fieldset>
          <fieldset class="format-group">
            <legend>文字の制御</legend>
            <div class="format-stack">
              <label class="format-check"><input type="checkbox" ${css.whiteSpace === "pre-wrap" ? "checked" : ""} data-format-field="wrapText" /> 折り返して全体を表示する</label>
              <label class="format-check"><input type="checkbox" data-format-field="shrinkToFit" /> 縮小して全体を表示する</label>
            </div>
          </fieldset>
        </section>
        <section class="format-dialog-panel" data-format-panel="font">
          <fieldset class="format-group">
            <legend>フォント</legend>
            <div class="format-dialog-grid three">
              <label class="format-field">${accessKeyLabelHtml("フォント(F)")}
                <select data-format-field="fontFamily">${formatFontOptions(currentFont)}</select>
              </label>
              <label class="format-field">${accessKeyLabelHtml("サイズ(S)")}
                <select data-format-field="fontSize">${["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36"].map((size) => formatOption(size, size, currentFontSize === size)).join("")}</select>
              </label>
              <label class="format-field">${accessKeyLabelHtml("色(C)")}
                <input type="color" value="${currentColor}" data-format-field="fontColor" />
              </label>
            </div>
          </fieldset>
          <fieldset class="format-group">
            <legend>スタイル</legend>
            <div class="format-inline-checks">
              <label class="format-check"><input type="checkbox" ${isBoldCssValue(css.fontWeight) ? "checked" : ""} data-format-field="bold" /> 太字</label>
              <label class="format-check"><input type="checkbox" ${css.fontStyle === "italic" ? "checked" : ""} data-format-field="italic" /> 斜体</label>
              <label class="format-check"><input type="checkbox" ${String(css.textDecoration || "").includes("underline") ? "checked" : ""} data-format-field="underline" /> 下線</label>
            </div>
          </fieldset>
        </section>
        <section class="format-dialog-panel" data-format-panel="border">
          <div class="format-border-layout">
            <fieldset class="format-group">
              <legend>線</legend>
              <div class="format-dialog-grid one">
                <label class="format-field">${accessKeyLabelHtml("スタイル(S)")}
                  <select data-format-field="borderStyle">
                    ${formatOption("solid", "実線", true)}
                    ${formatOption("dashed", "破線", false)}
                    ${formatOption("dotted", "点線", false)}
                    ${formatOption("double", "二重線", false)}
                  </select>
                </label>
                <label class="format-field">${accessKeyLabelHtml("色(C)")}
                  <input type="color" value="#000000" data-format-field="borderColor" />
                </label>
              </div>
            </fieldset>
            <fieldset class="format-group">
              <legend>プリセット</legend>
              <label class="format-field">種類
                <select data-format-field="borderPreset">
                  ${formatOption("none", "なし", false)}
                  ${formatOption("outline", "外枠", true)}
                  ${formatOption("inside", "内側", false)}
                  ${formatOption("all", "格子", false)}
                  ${formatOption("custom", "指定", false)}
                </select>
              </label>
            </fieldset>
            <fieldset class="format-group format-border-preview">
              <legend>罫線</legend>
              <div class="format-edge-grid" aria-label="罫線">
                ${["top", "right", "bottom", "left"].map((edge) => `<label class="format-check"><input type="checkbox" data-format-field="border-${edge}" /> ${borderEdgeLabel(edge)}</label>`).join("")}
              </div>
            </fieldset>
          </div>
        </section>
        <section class="format-dialog-panel" data-format-panel="fill">
          <fieldset class="format-group">
            <legend>背景色</legend>
            <div class="format-dialog-grid two">
              <label class="format-field">${accessKeyLabelHtml("色(C)")}
                <input type="color" value="${currentFill}" data-format-field="fillColor" />
              </label>
              <label class="format-field">${accessKeyLabelHtml("パターンの種類(P)")}
                <select data-format-field="fillPattern">
                  ${formatOption("solid", "塗りつぶし", true)}
                  ${formatOption("none", "なし", false)}
                </select>
              </label>
              <label class="format-check"><input type="checkbox" ${!cellHasFill(cell) ? "checked" : ""} data-format-field="noFill" /> 色なし</label>
            </div>
          </fieldset>
        </section>
        <section class="format-dialog-panel" data-format-panel="protection">
          <fieldset class="format-group">
            <legend>保護</legend>
            <p class="format-dialog-help">シートを保護したときに有効になります。</p>
            <div class="format-stack">
              <label class="format-check"><input type="checkbox" ${protection.locked !== false ? "checked" : ""} data-format-field="locked" /> ロック</label>
              <label class="format-check"><input type="checkbox" ${protection.hidden ? "checked" : ""} data-format-field="hiddenFormula" /> 数式を表示しない</label>
            </div>
          </fieldset>
        </section>
      </div>
      <div class="format-dialog-actions">
        <button class="format-dialog-button is-primary" type="button" data-format-dialog-action="ok">OK</button>
        <button class="format-dialog-button" type="button" data-format-dialog-action="cancel">キャンセル</button>
        <button class="format-dialog-button" type="button" data-format-dialog-action="apply">適用</button>
      </div>
    </div>
  `;
}



function formatOption(value, label, selected) {
  return `<option value="${escapeAttr(value)}"${selected ? " selected" : ""}>${escapeHtml(label)}</option>`;
}



function formatDialogDateFormatOptions(currentFormat = "") {
  const current = String(currentFormat || "yyyy/m/d");
  const formats = FORMAT_DIALOG_DATE_FORMATS.includes(current)
    ? FORMAT_DIALOG_DATE_FORMATS
    : [current, ...FORMAT_DIALOG_DATE_FORMATS];
  return formats.map((format) => formatOption(format, format, format === current)).join("");
}



function formatFontOptions(currentFont) {
  const select = $ribbonPanels.find('[data-ribbon-style="fontFamily"]')[0];
  const values = select ? [...select.options].map((option) => option.value) : ["Calibri", "MS Pゴシック", "Yu Gothic", "Meiryo", "Arial"];
  return values.map((font) => formatOption(font, font, font === currentFont)).join("");
}



function borderEdgeLabel(edge) {
  return { top: "上", right: "右", bottom: "下", left: "左" }[edge] || edge;
}



function handleFormatDialogTabClick(event) {
  const tab = event.currentTarget;
  const dialog = tab.closest("#formatCellsDialog");
  if (!dialog) return;
  const id = tab.dataset.formatDialogTab;
  dialog.querySelectorAll("[data-format-dialog-tab]").forEach((item) => item.classList.toggle("is-active", item === tab));
  dialog.querySelectorAll("[data-format-panel]").forEach((panel) => panel.classList.toggle("is-active", panel.dataset.formatPanel === id));
  if (state.formatDialog) state.formatDialog.activeTab = id;
}



function handleFormatDialogAction(event) {
  const action = event.currentTarget.dataset.formatDialogAction;
  if (action === "cancel") {
    closeFormatCellsDialog();
    return;
  }
  const applied = applyFormatDialogChanges();
  if (applied && action === "ok") closeFormatCellsDialog();
}



function handleFormatDialogFieldChange(event) {
  const fieldName = event.target?.dataset?.formatField || "";
  const dialog = event.target?.closest?.("#formatCellsDialog");
  if (dialog && event.type === "change" && (fieldName === "dateFormat" || fieldName === "numberCategory")) {
    syncFormatDialogCustomFormat(dialog, fieldName);
  }
  if (dialog && fieldName === "numberCategory") {
    updateFormatDialogNumberSettingVisibility(dialog);
  }
  updateFormatDialogPreview();
}



function updateFormatDialogNumberSettingVisibility(root = document) {
  const dialog = root.matches?.("#formatCellsDialog") ? root : root.querySelector?.("#formatCellsDialog");
  if (!dialog) return;
  const category = dialog.querySelector('[data-format-field="numberCategory"]')?.value || "general";
  const visibleSettings = formatDialogSettingsForCategory(category);
  dialog.querySelectorAll("[data-format-number-setting]").forEach((item) => {
    item.classList.toggle("is-hidden", !visibleSettings.has(item.dataset.formatNumberSetting));
  });
  const group = dialog.querySelector("[data-format-settings-group]");
  group?.classList.toggle("is-hidden", visibleSettings.size === 0);
}



function formatDialogSettingsForCategory(category) {
  switch (category) {
    case "number":
      return new Set(["decimals", "useSeparator"]);
    case "currency":
      return new Set(["decimals", "useSeparator", "currencySymbol"]);
    case "percent":
      return new Set(["decimals"]);
    case "date":
      return new Set(["dateFormat", "customFormat"]);
    default:
      return new Set();
  }
}



function syncFormatDialogCustomFormat(dialog, fieldName) {
  const category = dialog.querySelector('[data-format-field="numberCategory"]')?.value || "general";
  if (fieldName === "numberCategory" && category !== "date") return;
  const dateFormat = dialog.querySelector('[data-format-field="dateFormat"]')?.value || "";
  const customFormat = dialog.querySelector('[data-format-field="customFormat"]');
  if (!dateFormat || !customFormat) return;
  customFormat.value = dateFormat;
}



function updateFormatDialogPreview() {
  const dialog = document.querySelector("#formatCellsDialog");
  if (!dialog) return;
  const sample = dialog.querySelector("[data-format-sample]");
  if (!sample) return;
  const source = state.formatDialog?.point;
  const cell = source ? activeSheet().cells[cellKey(source.row, source.col)] : null;
  sample.textContent = previewNumberFormat(cell?.raw ?? cell?.cached ?? 1234.56, collectFormatDialogValues(dialog));
}



function collectFormatDialogValues(dialog = document.querySelector("#formatCellsDialog")) {
  const field = (name) => dialog?.querySelector(`[data-format-field="${name}"]`);
  const checked = (name) => Boolean(field(name)?.checked);
  const activeTab = dialog?.querySelector("[data-format-dialog-tab].is-active")?.dataset.formatDialogTab || state.formatDialog?.activeTab || "number";
  return {
    activeTab,
    numberCategory: field("numberCategory")?.value || "general",
    decimals: Math.max(0, Math.min(6, Number(field("decimals")?.value || 0))),
    useSeparator: checked("useSeparator"),
    currencySymbol: field("currencySymbol")?.value || "¥",
    dateFormat: field("dateFormat")?.value || "yyyy/mm/dd",
    customFormat: String(field("customFormat")?.value || "").trim(),
    horizontal: field("horizontal")?.value || "",
    vertical: field("vertical")?.value || "middle",
    wrapText: checked("wrapText"),
    shrinkToFit: checked("shrinkToFit"),
    fontFamily: field("fontFamily")?.value || "Calibri",
    fontSize: field("fontSize")?.value || "11",
    fontColor: normalizeColorValue(field("fontColor")?.value) || "#000000",
    bold: checked("bold"),
    italic: checked("italic"),
    underline: checked("underline"),
    borderStyle: field("borderStyle")?.value || "solid",
    borderColor: normalizeColorValue(field("borderColor")?.value) || "#000000",
    borderPreset: field("borderPreset")?.value || "outline",
    borderEdges: ["top", "right", "bottom", "left"].filter((edge) => checked(`border-${edge}`)),
    fillColor: normalizeColorValue(field("fillColor")?.value) || "#FFFFFF",
    fillPattern: field("fillPattern")?.value || "solid",
    noFill: checked("noFill"),
    locked: checked("locked"),
    hiddenFormula: checked("hiddenFormula"),
  };
}



function applyFormatDialogChanges() {
  const activeRange = state.formatDialog?.range || activeSelectionRange();
  const ranges = (state.formatDialog?.ranges?.length ? state.formatDialog.ranges : activeSelectionRanges())
    .map((range) => (range ? { ...range } : null))
    .filter(Boolean);
  const targetRanges = ranges.length ? ranges : activeRange ? [{ ...activeRange }] : [];
  if (!targetRanges.length) return false;
  const values = collectFormatDialogValues();
  if (applyFormatDialogChangesToSheetGroup(targetRanges, activeRange, values)) return true;
  if (!ensureActiveSheetFormatCellsAllowed()) return false;
  const sheet = activeSheet();
  const before = captureCellsForHistoryRanges(sheet, targetRanges);
  applyFormatDialogValuesToSheetRanges(sheet, targetRanges, values);
  const after = captureCellsForHistoryRanges(sheet, targetRanges);
  const label = rangesToLabel(targetRanges);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: activeRange || targetRanges[0],
    selectionAfter: activeRange || targetRanges[0],
    selectionBeforeRanges: targetRanges,
    selectionAfterRanges: targetRanges,
    label: `${label} セルの書式設定`,
  });
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(`${label} にセルの書式設定を適用しました。`);
  return true;
}



function applyFormatDialogChangesToSheetGroup(targetRanges, activeRange, values) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !targetRanges?.length) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではセルの書式を変更できません。")) return true;
    const ranges = materializedRangesForSheet(sheet, targetRanges, sheetIndex);
    if (ranges.length) items.push({ sheetIndex, sheet, ranges });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => applyFormatDialogValuesToSheetRanges(item.sheet, item.ranges, values));
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: activeRange || targetRanges[0],
      selectionAfter: activeRange || targetRanges[0],
      selectionBeforeRanges: targetRanges.map((range) => ({ ...range })),
      selectionAfterRanges: targetRanges.map((range) => ({ ...range })),
      label: `${rangesToLabel(targetRanges)} ${sheetIndexes.length} シートセルの書式設定`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${rangesToLabel(targetRanges)}にセルの書式設定を適用しました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`セルの書式設定を適用できませんでした：${error.message || error}`);
  }
  return true;
}



function applyFormatDialogValuesToSheetRanges(sheet, ranges, values) {
  const activeTab = values.activeTab || "number";
  const border = activeTab === "border" ? borderModelFromDialog(values) : null;
  const numFmt = activeTab === "number" ? numberFormatFromDialog(values) : "";
  ranges.forEach((range) => {
    forEachCellInRange(range, (row, col) => {
      const cell = ensureCellModel(sheet, row, col);
      switch (activeTab) {
        case "number":
          applyCellNumberFormat(cell, numFmt);
          break;
        case "alignment":
          applyDialogAlignmentCss(cell, values);
          break;
        case "font":
          applyDialogFontCss(cell, values);
          break;
        case "border":
          applyDialogBorders(cell, range, row, col, values, border);
          break;
        case "fill":
          applyDialogFillCss(cell, values);
          break;
        case "protection":
          cell.protection = { locked: values.locked, hidden: values.hiddenFormula };
          break;
        default:
          applyCellNumberFormat(cell, numFmt || numberFormatFromDialog(values));
          cell.css = cell.css || {};
          applyDialogCss(cell, values);
          applyDialogBorders(cell, range, row, col, values, borderModelFromDialog(values));
          cell.protection = { locked: values.locked, hidden: values.hiddenFormula };
          break;
      }
    });
  });
  internCellsInRanges(sheet, ranges);
}



function applyDialogCss(cell, values) {
  applyDialogFontCss(cell, values);
  applyDialogAlignmentCss(cell, values);
  applyDialogFillCss(cell, values);
}



function applyDialogFontCss(cell, values) {
  const css = ensureDialogCellCss(cell);
  css.fontFamily = fontFamilyForExcel(values.fontFamily);
  css.fontSize = `${values.fontSize}pt`;
  css.color = values.fontColor;
  cell.hasFontColor = true;
  if (values.bold) css.fontWeight = "700";
  else delete css.fontWeight;
  if (values.italic) css.fontStyle = "italic";
  else delete css.fontStyle;
  if (values.underline) css.textDecoration = "underline";
  else delete css.textDecoration;
}



function applyDialogAlignmentCss(cell, values) {
  const css = ensureDialogCellCss(cell);
  if (values.horizontal) css.textAlign = values.horizontal;
  else delete css.textAlign;
  css.verticalAlign = values.vertical;
  css.alignItems = verticalAlignToFlex(values.vertical);
  if (values.wrapText) {
    css.whiteSpace = "pre-wrap";
    css.overflowWrap = "anywhere";
  } else {
    delete css.whiteSpace;
    delete css.overflowWrap;
  }
  if (values.shrinkToFit) css.fontSizeAdjust = "from-font";
  else delete css.fontSizeAdjust;
}



function applyDialogFillCss(cell, values) {
  const css = ensureDialogCellCss(cell);
  if (values.noFill || values.fillPattern === "none") {
    delete css.backgroundColor;
    cell.hasFill = false;
  } else {
    css.backgroundColor = values.fillColor;
    cell.hasFill = true;
  }
}



function ensureDialogCellCss(cell) {
  cell.css = cell.css || {};
  return cell.css;
}



function applyDialogBorders(cell, range, row, col, values, border) {
  cell.borders = cell.borders || {};
  if (values.borderPreset === "none") {
    cell.borders = {};
    return;
  }
  if (values.borderPreset === "all") {
    cell.borders.top = { ...border };
    cell.borders.right = { ...border };
    cell.borders.bottom = { ...border };
    cell.borders.left = { ...border };
    return;
  }
  if (values.borderPreset === "outline") {
    if (row === range.top) cell.borders.top = { ...border };
    if (row === range.bottom) cell.borders.bottom = { ...border };
    if (col === range.left) cell.borders.left = { ...border };
    if (col === range.right) cell.borders.right = { ...border };
    return;
  }
  if (values.borderPreset === "inside") {
    if (row < range.bottom) cell.borders.bottom = { ...border };
    if (col < range.right) cell.borders.right = { ...border };
    return;
  }
  values.borderEdges.forEach((edge) => {
    cell.borders[edge] = { ...border };
  });
}



function borderModelFromDialog(values) {
  const width = values.borderStyle === "double" ? 3 : values.borderStyle === "solid" ? 1 : 2;
  return {
    width,
    style: values.borderStyle === "double" ? "double" : values.borderStyle,
    color: values.borderColor,
    weight: width + (values.borderStyle === "double" ? 2 : 0),
  };
}



function numberFormatFromDialog(values) {
  const zeros = values.decimals > 0 ? `.${"0".repeat(values.decimals)}` : "";
  const base = `${values.useSeparator ? "#,##" : ""}0${zeros}`;
  switch (values.numberCategory) {
    case "number":
      return base;
    case "currency":
      return `${values.currencySymbol}${base}`;
    case "percent":
      return `${base}%`;
    case "date":
      return values.customFormat || values.dateFormat;
    case "text":
      return "@";
    default:
      return "";
  }
}



function previewNumberFormat(value, values) {
  const numFmt = numberFormatFromDialog(values);
  if (!numFmt || numFmt === "@") return displayValue(value, numFmt);
  if (isDateFormat(numFmt)) {
    const serial = dateSerialFromValueForNumberFormat(value, numFmt);
    if (serial != null) return displayValue(serial, numFmt);
  }
  const number = typeof value === "number" ? value : Number.parseFloat(String(value || "1234.56").replace(/,/g, ""));
  const sampleValue = Number.isFinite(number) ? number : 1234.56;
  return displayValue(sampleValue, numFmt);
}



function numberCategoryFromFormat(numFmt = "") {
  const format = String(numFmt || "");
  if (!format) return "general";
  if (format === "@") return "text";
  if (isDateFormat(format)) return "date";
  if (/%/.test(format)) return "percent";
  if (currencySymbol(format)) return "currency";
  return "number";
}



function formatDecimalPlaces(numFmt = "") {
  const match = String(numFmt || "").match(/\.([0#]+)/);
  return match ? Math.min(6, match[1].length) : 0;
}



function formatUsesSeparator(numFmt = "") {
  return /#,##/.test(String(numFmt || ""));
}



function closeFormatCellsDialog() {
  document.querySelector("#formatCellsDialogOverlay")?.remove();
  state.formatDialog = null;
}



function closeFormatDialogOnEscape(event) {
  if (event.key !== "Escape" || !document.querySelector("#formatCellsDialogOverlay")) return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  closeFormatCellsDialog();
}



function openSelectedHyperlink(row, col) {
  const cell = activeSheet().cells[cellKey(row, col)];
  if (!cell?.hyperlink) {
    setStatus("現在のセルにはリンクがありません。");
    return;
  }
  openHyperlinkTarget(cell.hyperlink);
}



function handleCellHyperlinkClick(event) {
  const hyperlink = event.currentTarget?.dataset?.hyperlink || event.currentTarget?.getAttribute?.("href") || "";
  if (shouldDeferCellHyperlinkClick()) {
    event.preventDefault();
    return;
  }
  if (!openHyperlinkTarget(hyperlink)) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation?.();
}



function openHyperlinkTarget(hyperlink) {
  const text = String(hyperlink || "").trim();
  if (!text) return false;
  if (openInternalHyperlink(text)) {
    setStatus(`リンク先へ移動しました：${text}`);
    return true;
  }
  window.open(text, "_blank", "noopener,noreferrer");
  setStatus(`リンクを開きました：${text}`);
  return true;
}



function hyperlinkFromCellClickEvent(event) {
  if (event.defaultPrevented || event.button > 0) return "";
  if (shouldDeferCellHyperlinkClick()) return "";
  const target = event.target;
  if (!target?.closest) return "";
  if (target.closest(".cell-validation-dropdown-button,.cell-filter-button")) return "";
  const cellElement = target.closest(".sheet-cell[data-row][data-col]") || cellElementFromClientPoint(event.clientX, event.clientY);
  if (!cellElement) return "";
  if (cellElement.matches?.(".sheet-cell[contenteditable='true']")) return "";
  const link = cellElement.querySelector(".cell-link");
  if (!link || !pointIsInsideElementClientRects(link, event.clientX, event.clientY)) return "";
  return link.dataset?.hyperlink || link.getAttribute("href") || "";
}



function shouldDeferCellHyperlinkClick() {
  return Boolean(state.formatPainter || state.formulaEdit || activeSheetEditElement());
}



function cellElementFromClientPoint(clientX, clientY) {
  const hitElements = typeof document.elementsFromPoint === "function"
    ? document.elementsFromPoint(clientX, clientY)
    : [];
  const hitCell = hitElements.find((element) => element.matches?.(".sheet-cell[data-row][data-col]"));
  if (hitCell) return hitCell;
  const point = cellPointFromClientPosition(clientX, clientY);
  if (!point) return null;
  return $sheetHost[0]?.querySelector?.(`.sheet-cell[data-row="${point.row}"][data-col="${point.col}"]`) || null;
}



function pointIsInsideElementClientRects(element, clientX, clientY) {
  if (!element || !Number.isFinite(clientX) || !Number.isFinite(clientY)) return false;
  const rects = element.getClientRects?.();
  if (!rects?.length) return false;
  return Array.from(rects).some((rect) => (
    clientX >= rect.left - 1 &&
    clientX <= rect.right + 1 &&
    clientY >= rect.top - 1 &&
    clientY <= rect.bottom + 1
  ));
}



function openDefineNameDialog(row, col, options = {}) {
  if (!state.model) return;
  const ranges = contextMenuTargetRanges(row, col);
  if (!ranges.length) {
    setStatus("名前を定義するセル範囲を選択してください。");
    return;
  }

  closeFormatCellsDialog();
  closeCellStructureDialog();
  closeLinkDialog();
  closeNameDialog();

  const sheet = activeSheet();
  const name = suggestedDefinedNameForRange(sheet, ranges[0]);
  const refersTo = `=${ranges.map((range) => absoluteRangeReference(state.model.sheets[range.sheetIndex], range)).join(",")}`;
  state.nameDialog = {
    previousFocus: document.activeElement,
    ranges: ranges.map((range) => ({ ...range })),
    returnToNameManager: options.returnToNameManager === true,
  };

  const overlay = document.createElement("div");
  overlay.id = "nameDialogOverlay";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = nameDialogHtml(name, refersTo, { title: "新しい名前" });
  document.body.append(overlay);
  const input = overlay.querySelector('[data-name-field="name"]');
  input?.focus({ preventScroll: true });
  input?.select?.();
}



function nameDialogHtml(name, refersTo, options = {}) {
  const scopeOptions = [
    `<option value="workbook">ブック</option>`,
    ...state.model.sheets.map((sheet, index) => `<option value="sheet:${index}" ${options.scopeValue === `sheet:${index}` ? "selected" : ""}>${escapeHtml(sheet.name)}</option>`),
  ].join("");
  const selectedScope = options.scopeValue === "workbook" || !options.scopeValue ? "workbook" : options.scopeValue;
  const visibleNames = (state.model.definedNames || []).filter((definition) => !definition.hidden);
  const existingNames = visibleNames.length
    ? `<div class="excel-dialog-list" aria-label="定義済みの名前">${visibleNames
        .slice(0, 5)
        .map((definition) => `<span>${escapeHtml(definition.name)}</span>`)
        .join("")}</div>`
    : "";
  return `
    <div id="nameDialog" class="excel-mini-dialog name-dialog" role="dialog" aria-modal="true" aria-labelledby="nameDialogTitle">
      <div class="excel-mini-titlebar" id="nameDialogTitle">${escapeHtml(options.title || "新しい名前")}</div>
      <div class="excel-mini-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("名前(N):")}</span>
          <input type="text" data-name-field="name" value="${escapeAttr(name)}" />
        </label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("範囲(S):")}</span>
          <select data-name-field="scope">${scopeOptions.replace('value="workbook"', `value="workbook" ${selectedScope === "workbook" ? "selected" : ""}`)}</select>
        </label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("コメント(O):")}</span>
          <textarea data-name-field="comment" rows="3">${escapeHtml(options.comment || "")}</textarea>
        </label>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">${accessKeyLabelHtml("参照範囲(R):")}</span>
          <input type="text" data-name-field="refersTo" value="${escapeAttr(refersTo)}" />
        </label>
        ${existingNames}
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-name-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-name-dialog-action="cancel">キャンセル</button>
      </div>
    </div>
  `;
}



function handleNameDialogAction(event) {
  const action = event.currentTarget.dataset.nameDialogAction;
  if (action === "cancel") {
    const returnToNameManager = state.nameDialog?.returnToNameManager === true;
    const editIndex = Number(state.nameDialog?.editIndex);
    closeNameDialog();
    if (returnToNameManager) openNameManagerDialog(Number.isFinite(editIndex) ? editIndex : 0);
    return;
  }
  const returnToNameManager = state.nameDialog?.returnToNameManager === true;
  if (applyNameDialog()) {
    const nextIndex = Number(state.nameDialog?.resultIndex);
    closeNameDialog();
    if (returnToNameManager) openNameManagerDialog(Number.isFinite(nextIndex) ? nextIndex : 0);
  }
}



function applyNameDialog() {
  const dialog = document.querySelector("#nameDialog");
  if (!dialog || !state.nameDialog) return false;
  const name = dialog.querySelector('[data-name-field="name"]')?.value.trim() || "";
  const scopeValue = dialog.querySelector('[data-name-field="scope"]')?.value || "workbook";
  const comment = dialog.querySelector('[data-name-field="comment"]')?.value.trim() || "";
  const refersTo = normalizeDefinedNameRefersTo(dialog.querySelector('[data-name-field="refersTo"]')?.value || "");
  const validation = validateDefinedName(name);
  if (!validation.ok) {
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    return false;
  }
  if (!refersTo) {
    showMessageBox("参照範囲を入力してください。", { title: "Cht WebSheet" });
    return false;
  }

  const before = cloneDefinedNames(state.model.definedNames);
  const editIndex = Number(state.nameDialog.editIndex);
  const editing = Number.isFinite(editIndex) && editIndex >= 0 && editIndex < (state.model.definedNames || []).length;
  const current = editing ? state.model.definedNames[editIndex] : null;
  const definition = {
    ...(current || {}),
    name,
    scope: scopeValue.startsWith("sheet:") ? "sheet" : "workbook",
    sheetIndex: scopeValue.startsWith("sheet:") ? Number(scopeValue.split(":")[1]) : null,
    refersTo,
    comment,
    hidden: false,
  };
  if (definedNameHasConflict(definition, editing ? editIndex : -1)) {
    showMessageBox(definedNameDuplicateMessage(name), { title: "Cht WebSheet" });
    return false;
  }
  if (editing) state.model.definedNames[editIndex] = definition;
  else state.model.definedNames = upsertDefinedName(state.model.definedNames, definition);
  const after = cloneDefinedNames(state.model.definedNames);
  state.hf = buildFormulaEngine(state.model);
  state.nameDialog.resultIndex = editing
    ? editIndex
    : Math.max(0, state.model.definedNames.findIndex((item) => definedNameKey(item) === definedNameKey(definition)));
  const selection = activeSelectionRange();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    definedNamesBefore: before,
    definedNamesAfter: after,
    selectionBefore: selection,
    selectionAfter: selection,
    selectionBeforeRanges: activeSelectionRanges().map((range) => ({ ...range })),
    selectionAfterRanges: activeSelectionRanges().map((range) => ({ ...range })),
    label: `${name} ${editing ? "名前の編集" : "名前の定義"}`,
  });
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(`${name} を${editing ? "更新" : "定義"}しました。`);
  return true;
}



function closeNameDialog(options = {}) {
  document.querySelector("#nameDialogOverlay")?.remove();
  const previousFocus = state.nameDialog?.previousFocus;
  state.nameDialog = null;
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) {
    previousFocus.focus?.({ preventScroll: true });
  }
}



function closeNameDialogOnEscape(event) {
  if (event.key !== "Escape" || !state.nameDialog) return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  closeNameDialog();
}



function openLinkDialog(row, col) {
  if (!state.model) return;
  const point = contextMenuActionPoint(row, col);
  if (!point) {
    setStatus("リンクを設定するセルを選択してください。");
    return;
  }

  closeFormatCellsDialog();
  closeCellStructureDialog();
  closeNameDialog();
  closeLinkDialog();
  closeObjectLinkDialog();

  const sheet = activeSheet();
  const cell = sheet.cells[cellKey(point.row, point.col)] || {};
  const text = String(cell.raw ?? cell.display ?? "");
  const kind = hyperlinkKind(cell.hyperlink);
  state.linkDialog = {
    previousFocus: document.activeElement,
    point,
    kind,
    values: linkDialogInitialValues(cell.hyperlink, kind, point),
  };

  const overlay = document.createElement("div");
  overlay.id = "linkDialogOverlay";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = linkDialogHtml(point, text, cell.hyperlink || "", kind, state.linkDialog.values);
  document.body.append(overlay);
  const urlInput = overlay.querySelector('[data-link-field="url"]');
  urlInput?.focus({ preventScroll: true });
  urlInput?.select?.();
}



function linkDialogHtml(point, text, hyperlink, kind = "web", values = {}) {
  const hasHyperlink = Boolean(hyperlink);
  const meta = linkKindMeta(kind);
  const documentSheetIndex = normalizedDocumentSheetIndex(values.documentSheetIndex);
  const buttons = LINK_DIALOG_KINDS.map(
    (item) =>
      `<button class="${item.kind === kind ? "is-active" : ""}" type="button" data-link-kind="${item.kind}" aria-pressed="${item.kind === kind ? "true" : "false"}">${escapeHtml(item.label)}</button>`,
  ).join("");
  return `
    <div id="linkDialog" class="excel-mini-dialog link-dialog" role="dialog" aria-modal="true" aria-labelledby="linkDialogTitle">
      <div class="excel-mini-titlebar" id="linkDialogTitle">ハイパーリンクの挿入</div>
      <div class="excel-mini-body">
        <label class="excel-dialog-field"><span class="excel-dialog-caption" data-link-caption="text">${accessKeyLabelHtml("表示文字列(T):")}</span>
          <input type="text" data-link-field="text" value="${escapeAttr(text)}" />
        </label>
        <div class="link-dialog-content">
          <div class="link-dialog-kind" role="group" aria-label="リンク先">
            ${buttons}
          </div>
          <div class="link-dialog-main">
            <div class="link-dialog-address-label">${columnName(point.col)}${point.row}</div>
            <label class="excel-dialog-field"><span class="excel-dialog-caption" data-link-address-caption>${accessKeyLabelHtml(meta.fieldLabel)}</span>
              <input type="text" data-link-field="url" value="${escapeAttr(values[kind] ?? "")}" placeholder="${escapeAttr(meta.placeholder)}" />
            </label>
            <div class="link-dialog-document-panel" data-link-document-panel ${kind === "document" ? "" : "hidden"}>
              <div class="excel-dialog-caption">${accessKeyLabelHtml("またはドキュメント内の場所を選択してください(C):")}</div>
              <div class="link-dialog-document-tree" role="tree" aria-label="ドキュメント内の場所">
                ${linkDialogDocumentTargetsHtml(documentSheetIndex, values.documentName, {
                  sheetIndexAttribute: "data-link-document-sheet-index",
                  nameTargetAttribute: "data-link-document-name-target",
                })}
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="excel-mini-actions">
        ${hasHyperlink ? '<button class="excel-mini-button" type="button" data-link-dialog-action="remove">リンクの解除</button>' : ""}
        <span class="excel-mini-action-spacer"></span>
        <button class="excel-mini-button is-primary" type="button" data-link-dialog-action="ok">OK</button>
        <button class="excel-mini-button" type="button" data-link-dialog-action="cancel">キャンセル</button>
      </div>
    </div>
  `;
}



function linkDialogDocumentTargetsHtml(selectedSheetIndex, selectedNameTarget = "", options = {}) {
  const sheetIndexAttribute = options.sheetIndexAttribute || "data-link-document-sheet-index";
  const nameTargetAttribute = options.nameTargetAttribute || "data-link-document-name-target";
  const namesHtml = linkDialogDocumentDefinedNamesHtml(selectedNameTarget, nameTargetAttribute);
  return `
    <div class="link-dialog-tree-heading">セル参照</div>
    <div class="link-dialog-sheet-list" role="group" aria-label="シート">
      ${linkDialogDocumentSheetsHtml(selectedSheetIndex, sheetIndexAttribute)}
    </div>
    ${namesHtml
      ? `<div class="link-dialog-tree-heading">定義された名前</div>
        <div class="link-dialog-sheet-list" role="group" aria-label="定義された名前">
          ${namesHtml}
        </div>`
      : ""}
  `;
}



function linkDialogDocumentSheetsHtml(selectedSheetIndex, sheetIndexAttribute = "data-link-document-sheet-index") {
  return (state.model?.sheets || [])
    .map((sheet, index) => {
      if (!isSheetVisible(sheet)) return "";
      const active = index === selectedSheetIndex;
      return `<button class="${active ? "is-active" : ""}" type="button" ${sheetIndexAttribute}="${index}" role="treeitem" aria-selected="${active ? "true" : "false"}">${escapeHtml(sheet.name)}</button>`;
    })
    .join("");
}



function linkDialogDocumentDefinedNamesHtml(selectedNameTarget = "", nameTargetAttribute = "data-link-document-name-target") {
  const selected = normalizeDefinedNameLookup(selectedNameTarget);
  return linkDialogDocumentDefinedNameTargets()
    .map((item) => {
      const active = normalizeDefinedNameLookup(item.target) === selected;
      return `<button class="${active ? "is-active" : ""}" type="button" ${nameTargetAttribute}="${escapeAttr(item.target)}" role="treeitem" aria-selected="${active ? "true" : "false"}" title="${escapeAttr(item.refersTo || item.target)}">${escapeHtml(item.label)}</button>`;
    })
    .join("");
}



function linkDialogDocumentDefinedNameTargets() {
  return (state.model?.definedNames || [])
    .map((definition) => {
      if (!definition?.name || definition.hidden) return null;
      const scopeSheet = definition.scope === "sheet" ? state.model?.sheets?.[definition.sheetIndex] : null;
      if (scopeSheet && !isSheetVisible(scopeSheet)) return null;
      const reference = definedNameLinkReference(definition, scopeSheet || activeSheet());
      if (!reference?.sheet || !reference.range || !isSheetVisible(reference.sheet)) return null;
      const target = scopeSheet ? `${quoteSheetNameForFormula(scopeSheet.name)}!${definition.name}` : definition.name;
      return {
        label: scopeSheet ? `${definition.name} - ${scopeSheet.name}` : definition.name,
        target,
        refersTo: definition.refersTo || "",
      };
    })
    .filter(Boolean)
    .sort((left, right) => left.label.localeCompare(right.label, "ja", { sensitivity: "base" }));
}



function handleLinkKindClick(event) {
  const button = event.currentTarget;
  const dialog = button.closest("#linkDialog");
  if (!dialog || !state.linkDialog) return;
  event.preventDefault();
  const nextKind = button.dataset.linkKind;
  if (!LINK_DIALOG_KINDS.some((item) => item.kind === nextKind)) return;
  const input = dialog.querySelector('[data-link-field="url"]');
  const previousKind = state.linkDialog.kind || "web";
  state.linkDialog.values = state.linkDialog.values || {};
  state.linkDialog.values[previousKind] = input?.value || "";
  state.linkDialog.kind = nextKind;
  if (state.linkDialog.values[nextKind] == null) {
    state.linkDialog.values[nextKind] = defaultLinkDialogValue(nextKind, state.linkDialog.point);
  }
  const meta = linkKindMeta(nextKind);
  dialog.querySelectorAll("[data-link-kind]").forEach((kindButton) => {
    const active = kindButton.dataset.linkKind === nextKind;
    kindButton.classList.toggle("is-active", active);
    kindButton.setAttribute("aria-pressed", active ? "true" : "false");
  });
  const documentPanel = dialog.querySelector("[data-link-document-panel]");
  if (documentPanel) documentPanel.hidden = nextKind !== "document";
  const caption = dialog.querySelector("[data-link-address-caption]");
  if (caption) caption.innerHTML = accessKeyLabelHtml(meta.fieldLabel);
  if (input) {
    input.placeholder = meta.placeholder;
    input.value = state.linkDialog.values[nextKind] || "";
    input.focus({ preventScroll: true });
    input.select?.();
  }
}



function handleLinkDocumentSheetClick(event) {
  const button = event.currentTarget;
  const dialog = button.closest("#linkDialog");
  if (!dialog || !state.linkDialog) return;
  event.preventDefault();
  const index = Number(button.dataset.linkDocumentSheetIndex);
  if (!Number.isInteger(index) || !state.model?.sheets[index] || !isSheetVisible(state.model.sheets[index])) return;
  state.linkDialog.values = state.linkDialog.values || {};
  state.linkDialog.values.documentSheetIndex = index;
  state.linkDialog.values.documentName = "";
  dialog.querySelectorAll("[data-link-document-sheet-index]").forEach((sheetButton) => {
    const active = Number(sheetButton.dataset.linkDocumentSheetIndex) === index;
    sheetButton.classList.toggle("is-active", active);
    sheetButton.setAttribute("aria-selected", active ? "true" : "false");
  });
  dialog.querySelectorAll("[data-link-document-name-target]").forEach((nameButton) => {
    nameButton.classList.remove("is-active");
    nameButton.setAttribute("aria-selected", "false");
  });
  const input = dialog.querySelector('[data-link-field="url"]');
  if (input && !input.value.trim()) {
    input.value = defaultLinkDialogValue("document", state.linkDialog.point);
  }
  input?.focus({ preventScroll: true });
}



function handleLinkDocumentNameClick(event) {
  const button = event.currentTarget;
  const dialog = button.closest("#linkDialog");
  if (!dialog || !state.linkDialog) return;
  event.preventDefault();
  const target = button.dataset.linkDocumentNameTarget || "";
  if (!target) return;
  state.linkDialog.values = state.linkDialog.values || {};
  state.linkDialog.values.documentName = target;
  const reference = internalHyperlinkReferenceFromDefinedName(target, state.linkDialog.values.documentSheetIndex);
  if (reference?.sheetName) {
    const index = state.model.sheets.findIndex((sheet) => sheet.name === reference.sheetName && isSheetVisible(sheet));
    if (index >= 0) state.linkDialog.values.documentSheetIndex = index;
  }
  dialog.querySelectorAll("[data-link-document-name-target]").forEach((nameButton) => {
    const active = nameButton === button;
    nameButton.classList.toggle("is-active", active);
    nameButton.setAttribute("aria-selected", active ? "true" : "false");
  });
  dialog.querySelectorAll("[data-link-document-sheet-index]").forEach((sheetButton) => {
    const active = Number(sheetButton.dataset.linkDocumentSheetIndex) === state.linkDialog.values.documentSheetIndex;
    sheetButton.classList.toggle("is-active", active);
    sheetButton.setAttribute("aria-selected", active ? "true" : "false");
  });
  const input = dialog.querySelector('[data-link-field="url"]');
  if (input) input.value = target;
  input?.focus({ preventScroll: true });
}



function handleLinkDialogAction(event) {
  const action = event.currentTarget.dataset.linkDialogAction;
  if (action === "cancel") {
    closeLinkDialog();
    return;
  }
  if (action === "remove") {
    if (removeLinkFromDialog()) closeLinkDialog();
    return;
  }
  if (applyLinkDialog()) closeLinkDialog();
}



function applyLinkDialog() {
  const dialog = document.querySelector("#linkDialog");
  const point = state.linkDialog?.point;
  if (!dialog || !point) return false;
  const rawUrl = dialog.querySelector('[data-link-field="url"]')?.value.trim() || "";
  const kind = state.linkDialog?.kind || "web";
  const url = normalizeHyperlinkForKind(rawUrl, kind, state.linkDialog?.values?.documentSheetIndex);
  if (!url) {
    showMessageBox(linkKindEmptyMessage(kind), { title: "Cht WebSheet" });
    return false;
  }
  const rawText = dialog.querySelector('[data-link-field="text"]')?.value;
  const text = String(rawText || "").trim() || rawUrl;
  const handledByGroup = applyHyperlinkToCell(point, { text, hyperlink: url });
  if (!handledByGroup) setStatus(`${columnName(point.col)}${point.row} にリンクを設定しました。`);
  return true;
}



function removeLinkFromDialog() {
  const point = state.linkDialog?.point;
  if (!point) return false;
  const handledByGroup = applyHyperlinkToCell(point, { remove: true });
  if (!handledByGroup) setStatus(`${columnName(point.col)}${point.row} のリンクを解除しました。`);
  return true;
}



function applyHyperlinkToSheetCell(sheet, row, col, options) {
  const sheetId = getSheetId(sheet.name);
  const cell = ensureCellModel(sheet, row, col);
  if (options.remove) {
    cell.hyperlink = null;
  } else {
    writeCellContent(sheet, sheetId, row, col, options.text);
    const updated = ensureCellModel(sheet, row, col);
    updated.hyperlink = options.hyperlink;
    updated.css = updated.css || {};
    updated.css.color = updated.css.color || "#0563c1";
    updated.css.textDecoration = updated.css.textDecoration || "underline";
    updated.hasFontColor = true;
  }
  internCellsInRange(sheet, { top: row, left: col, bottom: row, right: col });
}



function applyHyperlinkToSheetGroup(point, options, range) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではリンクを変更できません。")) return true;
    const targetRanges = materializedRangesForSheet(sheet, [range], sheetIndex);
    const targetRange = targetRanges[0];
    if (targetRange) items.push({ sheetIndex, sheet, range: targetRange });
  }
  if (!items.length) return true;
  const sheetsBefore = captureSheetsForHistory(items.map((item) => item.sheetIndex));
  try {
    items.forEach((item) => {
      applyHyperlinkToSheetCell(item.sheet, item.range.top, item.range.left, options);
    });
    commitDataOperation(
      activeSheet(),
      null,
      range,
      `${rangeToLabel(range)} ${sheetIndexes.length} シートリンク`,
      `${sheetIndexes.length} 個のシートの${rangeToLabel(range)}のリンクを${options.remove ? "解除" : "設定"}しました。`,
      range,
      { sheetsBefore },
    );
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`リンクを変更できませんでした：${error.message || error}`);
  }
  return true;
}



function applyHyperlinkToCell(point, options) {
  const sheet = activeSheet();
  const range = normalizeRange(point, point);
  if (!options.skipSheetGroup && applyHyperlinkToSheetGroup(point, options, range)) return true;
  const before = captureCellsForHistory(sheet, range);
  applyHyperlinkToSheetCell(sheet, point.row, point.col, options);
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    label: `${columnName(point.col)}${point.row} リンク`,
  });
  renderSheet();
  updateFormulaBarForCell(point.row, point.col);
  return false;
}



function closeLinkDialog(options = {}) {
  document.querySelector("#linkDialogOverlay")?.remove();
  const previousFocus = state.linkDialog?.previousFocus;
  state.linkDialog = null;
  if (options.restoreFocus !== false && previousFocus && document.contains(previousFocus)) {
    previousFocus.focus?.({ preventScroll: true });
  }
}



function closeLinkDialogOnEscape(event) {
  if (event.key !== "Escape" || !state.linkDialog) return;
  if (isElementImeCompositionEvent(event)) return;
  event.preventDefault();
  event.stopImmediatePropagation?.();
  closeLinkDialog();
}



function openObjectLinkDialog() {
  const objects = selectedImageObjects();
  if (!objects.length) {
    setStatus("リンクを設定するオブジェクトを選択してください。");
    return;
  }
  closeFormatCellsDialog();
  closeCellStructureDialog();
  closeNameDialog();
  closeLinkDialog();
  closeObjectLinkDialog();
  const primary = objects.find((image) => image.id === state.selectedImage?.id) || objects[0];
  const kind = hyperlinkKind(primary?.hyperlink);
  const point = selectionToolbarPoint() || { sheetIndex: state.activeSheetIndex, row: 1, col: 1 };
  state.objectLinkDialog = {
    previousFocus: document.activeElement,
    imageIds: objects.map((image) => image.id),
    kind,
    point,
    values: linkDialogInitialValues(primary?.hyperlink || "", kind, point),
  };
  const overlay = document.createElement("div");
  overlay.id = "objectLinkDialogOverlay";
  overlay.className = "excel-mini-overlay";
  overlay.innerHTML = objectLinkDialogHtml(primary, objects.length, kind, state.objectLinkDialog.values);
  document.body.append(overlay);
  const urlInput = overlay.querySelector('[data-object-link-field="url"]');
  urlInput?.focus({ preventScroll: true });
  urlInput?.select?.();
}

export {
  adjustedCutSourceIndexAfterInsert,
  annotationAuthor,
  annotationHtmlWithBreaks,
  annotationInitial,
  annotationPointFromMarker,
  annotationPopupPosition,
  annotationTextFromCell,
  appendCommentReply,
  appendCommentReplyFromPopup,
  applyCellStructureDialog,
  applyClipboardColumnModels,
  applyClipboardRowModels,
  applyDialogAlignmentCss,
  applyDialogBorders,
  applyDialogCss,
  applyDialogFillCss,
  applyDialogFontCss,
  applyFormatDialogChanges,
  applyFormatDialogChangesToSheetGroup,
  applyFormatDialogValuesToSheetRanges,
  applyHeaderHiddenToSheet,
  applyHeaderSize,
  applyHeaderSizeDialog,
  applyHeaderSizeToSheet,
  applyHeaderSizeToSheetGroup,
  applyHeaderStructureOperation,
  applyHyperlinkToCell,
  applyHyperlinkToSheetCell,
  applyHyperlinkToSheetGroup,
  applyLinkDialog,
  applyNameDialog,
  applyQuickAnalysisTotal,
  applyQuickAnalysisTotalToSheetGroup,
  borderEdgeLabel,
  borderModelFromDialog,
  cancelAnnotationEditor,
  cancelCloseCommentThreadPopup,
  cancelCloseNotePopup,
  cellElementFromClientPoint,
  cellHasFormulaContent,
  cellHasPhonetics,
  cellPasteOptionsHtml,
  cellStructureDialogHtml,
  clearCellsInRange,
  clipboardCellEntries,
  clipboardHasFormula,
  closeAnnotationEditorOnEscape,
  closeCellContextMenu,
  closeCellContextMenuOnOutsidePointer,
  closeCellStructureDialog,
  closeCellStructureDialogOnEscape,
  closeCommentThreadPopup,
  closeFormatCellsDialog,
  closeFormatDialogOnEscape,
  closeHeaderSizeDialog,
  closeHeaderSizeDialogOnEscape,
  closeLinkDialog,
  closeLinkDialogOnEscape,
  closeNameDialog,
  closeNameDialogOnEscape,
  closeNotePopup,
  collectFormatDialogValues,
  commentInitial,
  commentThreadPopupHtml,
  commentThreadPostHtml,
  commitAnnotationEditor,
  commitAnnotationEditorOnOutsidePointer,
  cutSourceEnd,
  cutSourceStart,
  defaultCellStructureOption,
  deleteCellComment,
  deleteCellNote,
  ensureDialogCellCss,
  formatCellsDialogHtml,
  formatCommentTimestamp,
  formatDecimalPlaces,
  formatDialogDateFormatOptions,
  formatDialogSettingsForCategory,
  formatFontOptions,
  formatOption,
  formatUsesSeparator,
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
  headerSizeDialogHtml,
  hyperlinkFromCellClickEvent,
  insertCopiedHeaderCells,
  interpolateHexColor,
  isRangeInsideRange,
  linkDialogDocumentDefinedNameTargets,
  linkDialogDocumentDefinedNamesHtml,
  linkDialogDocumentSheetsHtml,
  linkDialogDocumentTargetsHtml,
  linkDialogHtml,
  makeCommentFromText,
  nameDialogHtml,
  normalizeCommentReplies,
  normalizeCommentText,
  notePopupHtml,
  numberCategoryFromFormat,
  numberFormatFromDialog,
  openCellAnnotationEditor,
  openCellStructureDialog,
  openCommentThreadPopupFromMarker,
  openDefineNameDialog,
  openFormatCellsDialog,
  openHeaderSizeDialog,
  openHyperlinkTarget,
  openLinkDialog,
  openNotePopupFromMarker,
  openObjectLinkDialog,
  openSelectedHyperlink,
  persistentNoteCellsForSheet,
  pointIsInsideElementClientRects,
  positionAnnotationPopup,
  positionPopupMenu,
  previewNumberFormat,
  quickAnalysisAvailableForContext,
  quickAnalysisTotalFormula,
  quickAnalysisTotalLabel,
  quickAnalysisTotalPlanForSheet,
  rangeHasFormulaContent,
  removeCellModelIfEmpty,
  removeContainedMerges,
  removeLinkFromDialog,
  renderAnnotationEditor,
  renderPersistentNotePopups,
  replyToCellComment,
  runCellContextMenuAction,
  scheduleCloseCommentThreadPopup,
  scheduleCloseNotePopup,
  setHeaderHidden,
  setHeaderHiddenForSheetGroup,
  shouldDeferCellHyperlinkClick,
  showCommentThreadPopup,
  showNotePopup,
  syncFormatDialogCustomFormat,
  targetOverlapsCutSource,
  toggleCellNoteVisibility,
  toggleSelectedPhonetics,
  updateFormatDialogNumberSettingVisibility,
  updateFormatDialogPreview,
  writeClipboardCellsIntoSheet,
  writeQuickAnalysisTotalPlan,
};
