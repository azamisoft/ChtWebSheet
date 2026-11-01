// Feature module; shared bindings are imported explicitly.
import {
  $,
  $formulaBar,
  $formulaInput,
  $sheetHost,
  $toggleFormulasButton,
  $workbookName,
  $workbookNameDisplay,
  $workbookNameDisplayText,
  $workbookWindowBar,
  APP_TITLE,
  COLUMN_HEADER_HEIGHT,
  DEFAULT_MACRO_SECURITY_LEVEL,
  FORMULA_ERROR_CODES,
  MACRO_SECURITY_OPTIONS,
  createIcons,
  icons,
  measureWorkbookNameText,
  state,
} from "../../app-core.js";
import {
  selectSingleRange,
} from "../data/query-editor.js";
import {
  syncRelationshipAutoDetectDialog,
} from "../data/validation-and-consolidation.js";
import {
  closeDataDialog,
  startShapeDrawing,
} from "../drawing/charts-and-insertion.js";
import {
  uniqueWorksheetObjectId,
} from "../editing/ink-and-clipboard.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  getCellRawInput,
  getDisplayForCell,
  isCellInRange,
  isSingleCellRange,
  rangeToLabel,
  sheetPixelFromClientPosition,
  updateSelectionUi,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  cellHasFormula,
  scrollRangeIntoView,
  selectionToolbarPoint,
  setRibbonCommandDisabled,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureSheetForHistory,
  captureSheetsForHistory,
  cloneDefinedNames,
  cloneJsonLikeValue,
  pushHistory,
} from "../editing/structure-and-history.js";
import {
  formulaPreviewDisplayValue,
  openNameManagerDialog,
  selectedDefinedNameIndexFromDialog,
  splitFormulaReferenceToken,
} from "../formula/dialogs-and-names.js";
import {
  forEachCellInRange,
  rangesIntersect,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  nameDialogHtml,
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
  cellRangeForSpecial,
  closeHomeCommandMenu,
  dependentFormulaTargetsForSelection,
  formulaReferencesForCell,
  homeCommandActionButtonHtml,
  openHomeCommandMenu,
  parseFormulaReferences,
  startInkDrawing,
  uniqueSelectionRanges,
} from "./printing-and-home-tools.js";
import {
  absoluteRangeReference,
  definedNameDuplicateMessage,
  definedNameHasConflict,
  normalizeDefinedNameRefersTo,
  parseCellKey,
  quoteSheetNameForFormula,
  rowHeaderWidthForSheet,
  sanitizeDefinedName,
  upsertDefinedName,
  validateDefinedName,
} from "./ribbon-and-links.js";
import {
  updateSheetFooterUi,
} from "./sheet-protection-and-translation.js";
import {
  buildFormulaEngine,
  formulaEngineSheetIndex,
  isSheetVisible,
  renderWorkbook,
  setActiveSheet,
  translateFormulaForCurrentFormulaEngine,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  normalizeMacroSecurity,
  normalizePaneState,
  normalizeSheetViewSnapshot,
  normalizeWorkbookCalcProperties,
  normalizeWorkbookCustomView,
  normalizeWorkbookProperties,
  normalizeWorkbookViewModel,
  normalizeWorkbookWindowView,
  normalizedSheetShowRuler,
  normalizedSheetViewMode,
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  activeSheet,
  excelSplitActivePane,
  getSheetId,
  updateHorizontalScrollbar,
} from "../workbook/saving-and-export.js";
import {
  restoreScrollRecordForView,
  topLeftCellFromScrollRecord,
} from "../workbook/tables-and-validation.js";
import {
  activeWorkbookWindowShowOption,
  applyHiddenIndexesToModels,
  applySheetZoomWithoutHistory,
  applyWorkbookWindowSheetView,
  computeWorkbookWindowArrangement,
  currentWorkbookWindowSnapshot,
  customViewsUnavailableMessage,
  customViewsUnavailableSummary,
  hiddenIndexesForModels,
  mutableWorkbookWindowRecordById,
  normalizedViewSideBySidePair,
  openViewCommandMenu,
  restoreCustomViewSelection,
  restoreWorkbookWindowScrollStable,
  restoreWorkbookWindowSelection,
  saveActiveWindowViewState,
  storedWorkbookWindowRecordById,
  updateActiveWorkbookWindowScrollState,
  viewCommandMenuButtonHtml,
  workbookWindowArrangementModeLabel,
  workbookWindowArrangementPreviewHtml,
  workbookWindowChipHtml,
  workbookWindowRecords,
} from "../workbook/views-and-macros.js";


function handleNameManagerRowDoubleClick(event) {
  const dialog = event.currentTarget.closest("[data-data-dialog='name-manager']");
  if (!dialog) return;
  const index = Number(event.currentTarget.dataset.formulaNameRow);
  if (Number.isFinite(index)) state.dataDialog.activeIndex = index;
  openNameEditorFromManager(dialog);
}



function openNameEditorFromManager(dialog) {
  const index = selectedDefinedNameIndexFromDialog(dialog);
  const definition = state.model?.definedNames?.[index];
  if (!definition) return;
  closeDataDialog();
  state.nameDialog = {
    previousFocus: document.activeElement,
    returnToNameManager: true,
    editIndex: index,
  };
  const scopeValue = definition.scope === "sheet" ? `sheet:${definition.sheetIndex}` : "workbook";
  const overlay = document.createElement("div");
  overlay.id = "nameDialogOverlay";
  overlay.className = "excel-dialog-overlay";
  overlay.innerHTML = nameDialogHtml(definition.name, definition.refersTo || "", {
    title: "名前の編集",
    scopeValue,
    comment: definition.comment || "",
  });
  document.body.append(overlay);
  overlay.querySelector('[data-name-field="name"]')?.focus({ preventScroll: true });
}



function applyNameManagerSave(dialog) {
  const index = selectedDefinedNameIndexFromDialog(dialog);
  const current = state.model?.definedNames?.[index];
  if (!current) return;
  const name = dialog.querySelector('[data-formula-name-field="name"]')?.value.trim() || "";
  const scopeValue = dialog.querySelector('[data-formula-name-field="scope"]')?.value || "workbook";
  const comment = dialog.querySelector('[data-formula-name-field="comment"]')?.value.trim() || "";
  const refersTo = normalizeDefinedNameRefersTo(dialog.querySelector('[data-formula-name-field="refersTo"]')?.value || "");
  const validation = validateDefinedName(name);
  if (!validation.ok) {
    showMessageBox(validation.message, { title: "Cht WebSheet" });
    return;
  }
  if (!refersTo) {
    showMessageBox("参照範囲を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const nextDefinition = {
    ...current,
    name,
    scope: scopeValue.startsWith("sheet:") ? "sheet" : "workbook",
    sheetIndex: scopeValue.startsWith("sheet:") ? Number(scopeValue.split(":")[1]) : null,
    refersTo,
    comment,
  };
  if (definedNameHasConflict(nextDefinition, index)) {
    showMessageBox(definedNameDuplicateMessage(name), { title: "Cht WebSheet" });
    return;
  }
  const before = cloneDefinedNames(state.model.definedNames);
  state.model.definedNames[index] = nextDefinition;
  const after = cloneDefinedNames(state.model.definedNames);
  state.hf = buildFormulaEngine(state.model);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    definedNamesBefore: before,
    definedNamesAfter: after,
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: `${name} 名前の更新`,
  });
  scheduleWorkbookPersistence();
  openNameManagerDialog(index);
  setStatus(`${name} を更新しました。`);
}



function deleteDefinedNameFromManager(dialog) {
  const index = selectedDefinedNameIndexFromDialog(dialog);
  const definition = state.model?.definedNames?.[index];
  if (!definition) return;
  showMessageBox(`名前「${definition.name}」を削除しますか?`, {
    title: "Cht WebSheet",
    tone: "warning",
    buttons: [
      {
        label: "削除",
        result: "delete",
        callback: () => {
          const before = cloneDefinedNames(state.model.definedNames);
          state.model.definedNames.splice(index, 1);
          const after = cloneDefinedNames(state.model.definedNames);
          state.hf = buildFormulaEngine(state.model);
          pushHistory({
            sheetIndex: state.activeSheetIndex,
            definedNamesBefore: before,
            definedNamesAfter: after,
            selectionBefore: activeSelectionRange(),
            selectionAfter: activeSelectionRange(),
            selectionBeforeRanges: activeSelectionRanges(),
            selectionAfterRanges: activeSelectionRanges(),
            label: `${definition.name} 名前の削除`,
          });
          scheduleWorkbookPersistence();
          openNameManagerDialog(Math.max(0, index - 1));
          setStatus(`${definition.name} を削除しました。`);
        },
      },
      { label: "キャンセル", result: "cancel" },
    ],
    cancelResult: "cancel",
  });
}



function openCreateNamesDialog() {
  const range = activeSelectionRanges()[0] || activeSelectionRange();
  if (!range || isSingleCellRange(range)) {
    showMessageBox("名前を作成するセル範囲を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  state.dataDialog = { type: "create-names", range: { ...range } };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog create-names-dialog" role="dialog" aria-modal="true" aria-labelledby="createNamesTitle" data-data-dialog="create-names">
      <div class="excel-mini-titlebar" id="createNamesTitle">選択範囲から名前を作成</div>
      <div class="excel-mini-body data-dialog-body">
        <div class="data-dialog-note">${escapeHtml(rangeToLabel(range))} のラベルから名前を作成します。</div>
        <label class="data-dialog-check"><input type="checkbox" data-create-name-source="top" checked />上端行</label>
        <label class="data-dialog-check"><input type="checkbox" data-create-name-source="left" checked />左端列</label>
        <label class="data-dialog-check"><input type="checkbox" data-create-name-source="bottom" />下端行</label>
        <label class="data-dialog-check"><input type="checkbox" data-create-name-source="right" />右端列</label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="create-names-ok">OK</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncRelationshipAutoDetectDialog(overlay.querySelector("[data-data-dialog='relationship-auto-detect']"));
}



function applyCreateNamesFromSelection(dialog) {
  const sheet = activeSheet();
  const range = state.dataDialog?.range;
  if (!range) return;
  const sources = new Set(
    [...dialog.querySelectorAll("[data-create-name-source]:checked")].map((input) => input.dataset.createNameSource),
  );
  const definitions = [];
  const pushDefinition = (nameRow, nameCol, refRange) => {
    const nameCell = sheet.cells[cellKey(nameRow, nameCol)];
    const name = sanitizeDefinedName(nameCell?.display ?? nameCell?.raw ?? "");
    if (!name || !validateDefinedName(name).ok || !refRange || refRange.top > refRange.bottom || refRange.left > refRange.right) return;
    definitions.push({
      name,
      scope: "workbook",
      sheetIndex: null,
      refersTo: `=${absoluteRangeReference(sheet, refRange)}`,
      comment: "",
      hidden: false,
    });
  };
  if (sources.has("top")) {
    for (let col = range.left; col <= range.right; col += 1) {
      pushDefinition(range.top, col, { sheetIndex: state.activeSheetIndex, top: range.top + 1, left: col, bottom: range.bottom, right: col });
    }
  }
  if (sources.has("bottom")) {
    for (let col = range.left; col <= range.right; col += 1) {
      pushDefinition(range.bottom, col, { sheetIndex: state.activeSheetIndex, top: range.top, left: col, bottom: range.bottom - 1, right: col });
    }
  }
  if (sources.has("left")) {
    for (let row = range.top; row <= range.bottom; row += 1) {
      pushDefinition(row, range.left, { sheetIndex: state.activeSheetIndex, top: row, left: range.left + 1, bottom: row, right: range.right });
    }
  }
  if (sources.has("right")) {
    for (let row = range.top; row <= range.bottom; row += 1) {
      pushDefinition(row, range.right, { sheetIndex: state.activeSheetIndex, top: row, left: range.left, bottom: row, right: range.right - 1 });
    }
  }
  if (!definitions.length) {
    showMessageBox("作成できる名前が見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const before = cloneDefinedNames(state.model.definedNames);
  definitions.forEach((definition) => {
    state.model.definedNames = upsertDefinedName(state.model.definedNames, definition);
  });
  const after = cloneDefinedNames(state.model.definedNames);
  state.hf = buildFormulaEngine(state.model);
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    definedNamesBefore: before,
    definedNamesAfter: after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: `${definitions.length} 件の名前を作成`,
  });
  closeDataDialog();
  scheduleWorkbookPersistence();
  setStatus(`${definitions.length} 件の名前を作成しました。`);
}



function traceFormulaReferences(kind) {
  const sheet = activeSheet();
  const ranges = activeSelectionRanges();
  if (!sheet || !ranges.length) return;
  let targets = [];
  let arrows = [];
  if (kind === "dependents") {
    const dependentTargets = dependentFormulaTargetsForSelection(sheet, ranges);
    targets = uniqueSelectionRanges(dependentTargets.map((target) => target.range));
    arrows = dependentTargets.flatMap((target) =>
      ranges
        .filter((range) => target.references.some((reference) => reference.sheetIndex === state.activeSheetIndex && rangesIntersect(reference.range, range)))
        .map((range) => ({ kind, fromRange: { ...range }, toRange: { ...target.range } })),
    );
  } else {
    const activePoint = state.selected?.sheetIndex === state.activeSheetIndex ? state.selected : selectionToolbarPoint();
    const sourceRanges = formulaAuditFormulaSourceRanges(sheet, ranges, activePoint);
    targets = sourceRanges.flatMap((sourceRange) =>
      formulaReferencesForCell(sheet, sourceRange.top, sourceRange.left)
        .filter((reference) => reference.sheetIndex === state.activeSheetIndex)
        .map((reference) => reference.range),
    );
    arrows = sourceRanges.flatMap((sourceRange) =>
      formulaReferencesForCell(sheet, sourceRange.top, sourceRange.left)
        .filter((reference) => reference.sheetIndex === state.activeSheetIndex)
        .map((reference) => ({ kind, fromRange: { ...reference.range }, toRange: { ...sourceRange } })),
    );
  }
  sheet.formulaAudits = normalizeFormulaAudits(sheet.formulaAudits);
  sheet.formulaAudits[kind] = uniqueSelectionRanges(targets).map((range) => ({ ...range }));
  const retained = (sheet.formulaAudits.arrows || []).filter((arrow) => arrow.kind !== kind);
  sheet.formulaAudits.arrows = retained.concat(uniqueFormulaAuditArrows(arrows));
  renderSheet();
  const label = kind === "dependents" ? "参照先" : "参照元";
  setStatus(`${label}を ${targets.length} 件トレースしました。`);
}



function formulaAuditFormulaSourceRanges(sheet, ranges, activePoint = null) {
  const activeCell = activePoint ? sheet.cells?.[cellKey(activePoint.row, activePoint.col)] : null;
  if (activePoint && cellHasFormula(activeCell)) return [cellRangeForSpecial(activePoint.row, activePoint.col)];
  const sourceRanges = [];
  (ranges || []).forEach((range) => {
    forEachCellInRange(range, (row, col) => {
      if (cellHasFormula(sheet.cells?.[cellKey(row, col)])) sourceRanges.push(cellRangeForSpecial(row, col));
    });
  });
  if (!sourceRanges.length && activePoint) sourceRanges.push(cellRangeForSpecial(activePoint.row, activePoint.col));
  return uniqueSelectionRanges(sourceRanges);
}



function removeFormulaAuditArrows() {
  const sheet = activeSheet();
  if (!sheet) return;
  sheet.formulaAudits = { precedents: [], dependents: [], arrows: [], ignoredErrors: [] };
  renderSheet();
  setStatus("トレース矢印を削除しました。");
}



function openFormulaRemoveArrowsMenu(trigger = null) {
  if (!trigger) {
    removeFormulaAuditArrows();
    return;
  }
  openHomeCommandMenu(trigger, "formula-remove-arrows", "トレース矢印の削除", `
    ${formulaAuditMenuButtonHtml("remove-all-arrows", "すべての矢印を削除(A)", "eraser")}
    ${formulaAuditMenuButtonHtml("remove-precedent-arrows", "参照元の矢印を削除(P)", "move-right")}
    ${formulaAuditMenuButtonHtml("remove-dependent-arrows", "参照先の矢印を削除(D)", "git-branch")}
  `);
}



function removeFormulaAuditArrowsByKind(kind) {
  const sheet = activeSheet();
  if (!sheet) return;
  sheet.formulaAudits = normalizeFormulaAudits(sheet.formulaAudits);
  sheet.formulaAudits[kind] = [];
  sheet.formulaAudits.arrows = (sheet.formulaAudits.arrows || []).filter((arrow) => arrow.kind !== kind);
  renderSheet();
  setStatus(`${kind === "dependents" ? "参照先" : "参照元"}のトレース矢印を削除しました。`);
}



function formulaAuditClassForCell(sheet, row, col) {
  const audits = sheet?.formulaAudits || {};
  const classes = [];
  if ((audits.precedents || []).some((range) => isCellInRange(range, row, col))) classes.push("formula-audit-precedent");
  if ((audits.dependents || []).some((range) => isCellInRange(range, row, col))) classes.push("formula-audit-dependent");
  return classes;
}



function normalizeFormulaAudits(value = null) {
  const source = value && typeof value === "object" ? value : {};
  return {
    precedents: Array.isArray(source.precedents) ? source.precedents.map(normalizeAuditRange).filter(Boolean) : [],
    dependents: Array.isArray(source.dependents) ? source.dependents.map(normalizeAuditRange).filter(Boolean) : [],
    arrows: Array.isArray(source.arrows) ? source.arrows.map(normalizeFormulaAuditArrow).filter(Boolean) : [],
    ignoredErrors: Array.isArray(source.ignoredErrors) ? source.ignoredErrors.map(String) : [],
  };
}



function normalizeAuditRange(range) {
  if (!range) return null;
  const top = Math.max(1, Math.round(Number(range.top) || 0));
  const left = Math.max(1, Math.round(Number(range.left) || 0));
  const bottom = Math.max(top, Math.round(Number(range.bottom) || top));
  const right = Math.max(left, Math.round(Number(range.right) || left));
  const sheetIndex = Number.isFinite(Number(range.sheetIndex)) ? Number(range.sheetIndex) : state.activeSheetIndex;
  return { sheetIndex, top, left, bottom, right };
}



function normalizeFormulaAuditArrow(arrow) {
  const fromRange = normalizeAuditRange(arrow?.fromRange);
  const toRange = normalizeAuditRange(arrow?.toRange);
  if (!fromRange || !toRange) return null;
  return { kind: arrow?.kind === "dependents" ? "dependents" : "precedents", fromRange, toRange };
}



function uniqueFormulaAuditArrows(arrows) {
  const seen = new Set();
  return (arrows || [])
    .map(normalizeFormulaAuditArrow)
    .filter(Boolean)
    .filter((arrow) => {
      const key = `${arrow.kind}:${rangeToLabel(arrow.fromRange)}:${rangeToLabel(arrow.toRange)}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}



function toggleShowFormulasFromRibbon() {
  const workbookViewBefore = captureWorkbookViewState();
  state.showFormulas = !state.showFormulas;
  const sheet = activeSheet();
  if (sheet) sheet.showFormulas = state.showFormulas;
  $toggleFormulasButton
    .prop("disabled", false)
    .attr("title", state.showFormulas ? "計算結果を表示" : "数式を表示")
    .html(`<i data-lucide="${state.showFormulas ? "eye-off" : "eye"}"></i>`);
  document.querySelector('[data-ribbon-command="formula-show-formulas"]')?.classList.toggle("is-active", state.showFormulas);
  createIcons({ icons });
  commitWorkbookViewChange(
    workbookViewBefore,
    state.showFormulas ? "シートに数式を表示しました。" : "シートに計算結果を表示しました。",
    state.showFormulas ? "数式の表示" : "計算結果の表示",
  );
}



function runFormulaErrorChecking() {
  const sheet = activeSheet();
  if (!sheet) return;
  const errors = formulaErrorsForSheet(sheet, { includeIgnored: false });
  if (!errors.length) {
    showMessageBox("エラーは見つかりませんでした。", { title: "Cht WebSheet" });
    setStatus("数式エラーは見つかりませんでした。");
    return;
  }
  openFormulaErrorCheckingDialog(errors, 0);
  setStatus(`${errors.length} 件の数式エラーが見つかりました。`);
}



function openFormulaErrorCheckingMenu(trigger = null) {
  if (!trigger) {
    runFormulaErrorChecking();
    return;
  }
  const ignoredCount = normalizeFormulaAudits(activeSheet()?.formulaAudits).ignoredErrors.length;
  const circularRefs = circularFormulaReferencesForSheet(activeSheet());
  const circularItems = circularRefs.length
    ? circularRefs.slice(0, 12).map((item) => formulaAuditMenuButtonHtml(`circular:${item.row}:${item.col}`, `${item.address}`, "refresh-cw")).join("")
    : '<button type="button" role="menuitem" disabled><i data-lucide="refresh-cw"></i><span>循環参照はありません</span></button>';
  openHomeCommandMenu(trigger, "formula-error-checking", "エラー チェック", `
    ${formulaAuditMenuButtonHtml("run-error-checking", "エラー チェック(E)...", "check")}
    ${formulaAuditMenuButtonHtml("trace-errors", "トレース エラー(T)", "route")}
    <div class="home-command-menu-separator" role="separator"></div>
    ${formulaAuditMenuButtonHtml("reset-ignored-errors", `無視したエラーをリセット(R)${ignoredCount ? ` (${ignoredCount})` : ""}`, "rotate-ccw")}
    <div class="home-command-menu-separator" role="separator"></div>
    <div class="home-command-menu-heading">循環参照(C)</div>
    ${circularItems}
  `);
}



function formulaAuditMenuButtonHtml(action, label, icon) {
  return homeCommandActionButtonHtml("data-formula-audit-menu-action", action, label, icon);
}



function handleFormulaAuditMenuAction(event) {
  const action = event.currentTarget.dataset.formulaAuditMenuAction || "";
  closeHomeCommandMenu();
  if (action === "remove-all-arrows") removeFormulaAuditArrows();
  else if (action === "remove-precedent-arrows") removeFormulaAuditArrowsByKind("precedents");
  else if (action === "remove-dependent-arrows") removeFormulaAuditArrowsByKind("dependents");
  else if (action === "run-error-checking") runFormulaErrorChecking();
  else if (action === "trace-errors") traceFormulaErrors();
  else if (action === "reset-ignored-errors") resetIgnoredFormulaErrors();
  else if (action.startsWith("circular:")) selectCircularFormulaReference(action);
}



function circularFormulaReferencesForSheet(sheet) {
  if (!sheet) return [];
  const formulaCells = Object.entries(sheet.cells || {})
    .map(([key, cell]) => {
      if (!cellHasFormula(cell)) return null;
      const [row, col] = parseCellKey(key);
      return { key, row, col, address: `${columnName(col)}${row}` };
    })
    .filter(Boolean)
    .sort((left, right) => left.row - right.row || left.col - right.col);
  if (!formulaCells.length) return [];
  const formulaKeySet = new Set(formulaCells.map((cell) => cell.key));
  const graph = new Map();
  formulaCells.forEach((cell) => {
    const dependencies = new Set();
    formulaReferencesForCell(sheet, cell.row, cell.col)
      .filter((reference) => reference.sheetIndex === state.activeSheetIndex)
      .forEach((reference) => {
        forEachCellInRange(reference.range, (row, col) => {
          const key = cellKey(row, col);
          if (formulaKeySet.has(key)) dependencies.add(key);
        });
      });
    graph.set(cell.key, [...dependencies]);
  });
  const circular = new Set();
  formulaCells.forEach((cell) => {
    if (formulaDependencyPathReturnsTo(cell.key, cell.key, graph, new Set())) circular.add(cell.key);
  });
  return formulaCells.filter((cell) => circular.has(cell.key));
}



function formulaDependencyPathReturnsTo(startKey, currentKey, graph, visited) {
  const dependencies = graph.get(currentKey) || [];
  for (const dependency of dependencies) {
    if (dependency === startKey) return true;
    if (visited.has(dependency)) continue;
    visited.add(dependency);
    if (formulaDependencyPathReturnsTo(startKey, dependency, graph, visited)) return true;
  }
  return false;
}



function selectCircularFormulaReference(action) {
  const [, rowText, colText] = String(action || "").split(":");
  const row = Number(rowText);
  const col = Number(colText);
  if (!Number.isFinite(row) || !Number.isFinite(col)) return;
  const range = cellRangeForSpecial(row, col);
  selectSingleRange(range, `${rangeToLabel(range)} の循環参照を選択しました。`);
}



function traceFormulaErrors() {
  const sheet = activeSheet();
  if (!sheet) return;
  const errors = formulaErrorsForSheet(sheet, { includeIgnored: false });
  if (!errors.length) {
    showMessageBox("トレースするエラーは見つかりませんでした。", { title: "Cht WebSheet" });
    return;
  }
  const targetRanges = [];
  const arrows = [];
  errors.forEach((error) => {
    const errorRange = cellRangeForSpecial(error.row, error.col);
    formulaReferencesForCell(sheet, error.row, error.col)
      .filter((reference) => reference.sheetIndex === state.activeSheetIndex)
      .forEach((reference) => {
        targetRanges.push(reference.range);
        arrows.push({ kind: "precedents", fromRange: { ...reference.range }, toRange: { ...errorRange } });
      });
  });
  if (!arrows.length) {
    showMessageBox("エラーの参照元を特定できませんでした。", { title: "Cht WebSheet" });
    return;
  }
  sheet.formulaAudits = normalizeFormulaAudits(sheet.formulaAudits);
  sheet.formulaAudits.precedents = uniqueSelectionRanges(targetRanges).map((range) => ({ ...range }));
  sheet.formulaAudits.arrows = (sheet.formulaAudits.arrows || []).filter((arrow) => arrow.kind !== "precedents").concat(uniqueFormulaAuditArrows(arrows));
  selectFormulaError(errors[0], { render: true });
  setStatus(`${errors.length} 件のエラーの参照元をトレースしました。`);
}



function resetIgnoredFormulaErrors() {
  const sheet = activeSheet();
  if (!sheet) return;
  sheet.formulaAudits = normalizeFormulaAudits(sheet.formulaAudits);
  const count = sheet.formulaAudits.ignoredErrors.length;
  sheet.formulaAudits.ignoredErrors = [];
  scheduleWorkbookPersistence();
  setStatus(count ? `${count} 件の無視したエラーをリセットしました。` : "無視したエラーはありません。");
}



function formulaErrorsForSheet(sheet, options = {}) {
  const sheetId = getSheetId(sheet?.name);
  const ignored = new Set(normalizeFormulaAudits(sheet?.formulaAudits).ignoredErrors);
  const errors = [];
  Object.entries(sheet?.cells || {}).forEach(([key, cell]) => {
    if (!cellHasFormula(cell)) return;
    const [row, col] = parseCellKey(key);
    const display = getDisplayForCell(sheet, sheetId, row, col);
    const value = String(display ?? "");
    if (!FORMULA_ERROR_CODES.some((code) => value.startsWith(code))) return;
    const address = `${columnName(col)}${row}`;
    const ignoreKey = formulaErrorIgnoreKey(row, col, value);
    if (!options.includeIgnored && ignored.has(ignoreKey)) return;
    errors.push({
      row,
      col,
      address,
      value,
      formula: getCellRawInput(sheet, row, col) || "",
      message: formulaErrorMessage(value),
      ignoreKey,
    });
  });
  errors.sort((left, right) => left.row - right.row || left.col - right.col);
  return errors;
}



function formulaErrorMessage(value) {
  if (value.startsWith("#DIV/0!")) return "数式または関数で 0 または空白セルによる除算が行われています。";
  if (value.startsWith("#VALUE!")) return "数式内の値の種類が正しくありません。";
  if (value.startsWith("#REF!")) return "数式に無効なセル参照が含まれています。";
  if (value.startsWith("#NAME?")) return "数式に認識できない名前または関数名が含まれています。";
  if (value.startsWith("#N/A")) return "数式で必要な値が見つかりません。";
  if (value.startsWith("#NUM!")) return "数式内の数値に問題があります。";
  if (value.startsWith("#NULL!")) return "指定した 2 つの範囲が交差していません。";
  if (value.startsWith("#CYCLE!")) return "数式に循環参照が含まれています。";
  return "数式の計算結果にエラーがあります。";
}



function formulaErrorIgnoreKey(row, col, value) {
  return `${row}:${col}:${String(value || "").split(/\s+/)[0]}`;
}



function openFormulaErrorCheckingDialog(errors, index = 0) {
  state.dataDialog = {
    type: "formula-error-checking",
    errors: errors.map((error) => ({ ...error })),
    index: Math.min(Math.max(0, Number(index) || 0), Math.max(0, errors.length - 1)),
  };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog formula-error-dialog" role="dialog" aria-modal="true" aria-labelledby="formulaErrorTitle" data-data-dialog="formula-error-checking">
      <div class="excel-mini-titlebar" id="formulaErrorTitle">エラー チェック</div>
      <div class="excel-mini-body data-dialog-body formula-error-dialog-body">
        <div class="formula-error-summary" data-formula-error-summary></div>
        <div class="formula-error-card">
          <div class="formula-error-location" data-formula-error-location></div>
          <div class="formula-error-value" data-formula-error-value></div>
          <div class="formula-error-message" data-formula-error-message></div>
          <label class="excel-dialog-field"><span class="excel-dialog-caption">数式(F):</span><textarea rows="2" readonly data-formula-error-formula></textarea></label>
        </div>
        <div class="formula-error-list" role="listbox" data-formula-error-list></div>
      </div>
      <div class="excel-mini-actions formula-error-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="formula-error-previous">前へ(B)</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="formula-error-next">次へ(N)</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="formula-error-go-cell">セルへ移動(G)</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="formula-error-edit">数式バーで編集(E)</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="formula-error-ignore">エラーを無視(I)</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="cancel">閉じる</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncFormulaErrorCheckingDialog(overlay.querySelector("[data-data-dialog='formula-error-checking']"));
}



function syncFormulaErrorCheckingDialog(dialog = document.querySelector("[data-data-dialog='formula-error-checking']")) {
  const stateDialog = state.dataDialog;
  if (!dialog || stateDialog?.type !== "formula-error-checking") return;
  const errors = stateDialog.errors || [];
  const index = Math.min(Math.max(0, Number(stateDialog.index) || 0), Math.max(0, errors.length - 1));
  stateDialog.index = index;
  const error = errors[index];
  dialog.querySelector("[data-formula-error-summary]").textContent = errors.length
    ? `${index + 1} / ${errors.length} 件のエラー`
    : "エラーは見つかりませんでした。";
  dialog.querySelector("[data-formula-error-location]").textContent = error ? `${activeSheet().name}!${error.address}` : "";
  dialog.querySelector("[data-formula-error-value]").textContent = error?.value || "";
  dialog.querySelector("[data-formula-error-message]").textContent = error?.message || "";
  const formulaInput = dialog.querySelector("[data-formula-error-formula]");
  if (formulaInput) formulaInput.value = error?.formula || "";
  const list = dialog.querySelector("[data-formula-error-list]");
  if (list) {
    list.innerHTML = errors.map((item, itemIndex) => `
      <button class="${itemIndex === index ? "is-active" : ""}" type="button" role="option" aria-selected="${itemIndex === index ? "true" : "false"}" data-data-dialog-action="formula-error-select" data-formula-error-index="${itemIndex}">
        <span>${escapeHtml(item.address)}</span><strong>${escapeHtml(item.value)}</strong>
      </button>`).join("");
  }
  dialog.querySelector("[data-data-dialog-action='formula-error-previous']").disabled = !errors.length;
  dialog.querySelector("[data-data-dialog-action='formula-error-next']").disabled = !errors.length;
  dialog.querySelector("[data-data-dialog-action='formula-error-ignore']").disabled = !errors.length;
  dialog.querySelector("[data-data-dialog-action='formula-error-edit']").disabled = !errors.length;
  dialog.querySelector("[data-data-dialog-action='formula-error-go-cell']").disabled = !errors.length;
  if (error) selectFormulaError(error, { render: true });
}



function selectFormulaError(error, options = {}) {
  if (!error) return;
  const range = cellRangeForSpecial(error.row, error.col);
  state.selected = { sheetIndex: state.activeSheetIndex, row: error.row, col: error.col };
  state.selectionRange = range;
  state.selectionRanges = [{ ...range }];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectionDrag = null;
  scrollRangeIntoView(range);
  if (options.render) renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
}



function handleFormulaErrorDialogAction(action, dialog, button) {
  const stateDialog = state.dataDialog;
  if (stateDialog?.type !== "formula-error-checking") return false;
  const errors = stateDialog.errors || [];
  if (action === "formula-error-select") {
    stateDialog.index = Math.min(Math.max(0, Number(button.dataset.formulaErrorIndex) || 0), Math.max(0, errors.length - 1));
    syncFormulaErrorCheckingDialog(dialog);
    return true;
  }
  if (action === "formula-error-previous" || action === "formula-error-next") {
    if (!errors.length) return true;
    const delta = action === "formula-error-next" ? 1 : -1;
    stateDialog.index = (Number(stateDialog.index || 0) + delta + errors.length) % errors.length;
    syncFormulaErrorCheckingDialog(dialog);
    return true;
  }
  const error = errors[stateDialog.index || 0];
  if (action === "formula-error-go-cell") {
    selectFormulaError(error, { render: true });
    return true;
  }
  if (action === "formula-error-edit") {
    selectFormulaError(error, { render: true });
    closeDataDialog();
    $formulaInput.trigger("focus").select();
    setStatus(`${error.address} の数式を編集します。`);
    return true;
  }
  if (action === "formula-error-ignore") {
    const sheet = activeSheet();
    sheet.formulaAudits = normalizeFormulaAudits(sheet.formulaAudits);
    if (error && !sheet.formulaAudits.ignoredErrors.includes(error.ignoreKey)) sheet.formulaAudits.ignoredErrors.push(error.ignoreKey);
    const nextErrors = formulaErrorsForSheet(sheet, { includeIgnored: false });
    if (!nextErrors.length) {
      closeDataDialog();
      showMessageBox("エラーは見つかりませんでした。", { title: "Cht WebSheet" });
      scheduleWorkbookPersistence();
      return true;
    }
    stateDialog.errors = nextErrors;
    stateDialog.index = Math.min(Number(stateDialog.index || 0), nextErrors.length - 1);
    scheduleWorkbookPersistence();
    syncFormulaErrorCheckingDialog(dialog);
    return true;
  }
  return false;
}



function openEvaluateFormulaDialog() {
  const point = selectionToolbarPoint();
  if (!point) return;
  const sheet = activeSheet();
  const formula = getCellRawInput(sheet, point.row, point.col);
  if (!String(formula || "").trim().startsWith("=")) {
    showMessageBox("アクティブ セルに数式がありません。", { title: "Cht WebSheet" });
    return;
  }
  const trace = formulaEvaluationTraceForCell(sheet, point.row, point.col);
  state.dataDialog = {
    type: "formula-evaluate",
    row: point.row,
    col: point.col,
    step: 0,
    trace,
    rootTrace: trace,
    traceStack: [],
  };
  closeDataDialog({ keepState: true });
  const overlay = document.createElement("div");
  overlay.className = "excel-dialog-overlay";
  overlay.id = "dataDialog";
  overlay.innerHTML = `
    <div class="excel-mini-dialog data-dialog formula-evaluate-dialog" role="dialog" aria-modal="true" aria-labelledby="formulaEvaluateTitle" data-data-dialog="formula-evaluate">
      <div class="excel-mini-titlebar" id="formulaEvaluateTitle">数式の検証</div>
      <div class="excel-mini-body data-dialog-body formula-dialog-body">
        <div class="data-dialog-note" data-formula-evaluate-location></div>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">検証(E):</span><div class="formula-evaluate-expression" data-formula-evaluate-expression></div></label>
        <div class="formula-evaluate-steps" data-formula-evaluate-steps></div>
        <label class="excel-dialog-field"><span class="excel-dialog-caption">計算結果:</span><input type="text" readonly data-formula-evaluate-result /></label>
      </div>
      <div class="excel-mini-actions">
        <button class="excel-mini-button" type="button" data-data-dialog-action="formula-evaluate-step">評価(E)</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="formula-evaluate-restart">最初から(R)</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="formula-evaluate-step-in">ステップ イン(I)</button>
        <button class="excel-mini-button" type="button" data-data-dialog-action="formula-evaluate-step-out">ステップ アウト(O)</button>
        <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="cancel">閉じる</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  syncEvaluateFormulaDialog(overlay.querySelector("[data-data-dialog='formula-evaluate']"));
}



function formulaEvaluationTraceForCell(sheet, row, col, options = {}) {
  const formula = String(getCellRawInput(sheet, row, col) || "");
  const sheetId = getSheetId(sheet?.name);
  const result = String(getDisplayForCell(sheet, sheetId, row, col) ?? "");
  return formulaEvaluationTraceForExpression(sheet, row, col, formula, {
    result,
    label: "数式",
    ...options,
  });
}



function formulaEvaluationTraceForExpression(sheet, row, col, expression, options = {}) {
  const formula = formulaExpressionWithPrefix(expression);
  const result = options.result ?? evaluateFormulaTraceExpression(formula, sheet, row, col);
  const sheetIndex = formulaEngineSheetIndex(sheet);
  const references = formulaEvaluationReferencesForExpression(formula, sheet);
  const steps = [{ kind: "formula", label: options.label || "数式", expression: formula, range: null }];
  let displayExpression = formula;
  references.forEach((reference) => {
    displayExpression = replaceFormulaReferenceForDisplay(displayExpression, reference, reference.value);
    steps.push({
      kind: "reference",
      label: `${reference.label} = ${reference.value}`,
      expression: displayExpression,
      range: reference.range,
      reference,
    });
  });
  steps.push({ kind: "result", label: `結果 = ${result}`, expression: result, range: null });
  return {
    sheetIndex,
    sheetName: sheet?.name || "",
    row,
    col,
    formula,
    result,
    references,
    steps,
    scopeLabel: options.scopeLabel || "",
    sourceExpression: options.sourceExpression || formula,
  };
}



function formulaEvaluationReferencesForExpression(formula, sheet) {
  const references = parseFormulaReferences(formula, sheet);
  const seen = new Set();
  return references
    .map((reference) => {
      const label = formulaEvaluationReferenceLabel(reference, sheet);
      const key = `${reference.sheetIndex}:${rangeToLabel(reference.range)}`;
      if (seen.has(key)) return null;
      seen.add(key);
      return {
        ...reference,
        label,
        sheetName: reference.sheet?.name || "",
        range: { sheetIndex: reference.sheetIndex, ...reference.range },
        value: displayForFormulaReference(reference),
      };
    })
    .filter(Boolean);
}



function formulaEvaluationReferenceLabel(reference, baseSheet) {
  const label = rangeToLabel(reference?.range);
  if (!reference?.sheet || reference.sheet === baseSheet) return label;
  return `${quoteSheetNameForFormula(reference.sheet.name)}!${label}`;
}



function formulaExpressionWithPrefix(expression) {
  const text = String(expression || "").trim();
  if (!text) return "=";
  return text.startsWith("=") ? text : `=${text}`;
}



function evaluateFormulaTraceExpression(expression, sheet, row, col) {
  const formula = formulaExpressionWithPrefix(expression);
  const sheetId = sheet ? getSheetId(sheet.name) : null;
  if (!state.hf?.calculateFormula || sheetId == null) return "";
  try {
    const translated = translateFormulaForCurrentFormulaEngine(formula, formulaEngineSheetIndex(sheet), { row, col });
    return formulaPreviewDisplayValue(state.hf.calculateFormula(translated, sheetId));
  } catch (error) {
    return formulaPreviewDisplayValue(error);
  }
}



function displayForFormulaReference(reference) {
  const sheet = reference.sheet || activeSheet();
  const sheetId = getSheetId(sheet.name);
  const range = reference.range;
  if (range.top === range.bottom && range.left === range.right) {
    return String(getDisplayForCell(sheet, sheetId, range.top, range.left) ?? "");
  }
  const values = [];
  const maxRows = Math.min(range.bottom, range.top + 2);
  const maxCols = Math.min(range.right, range.left + 2);
  for (let row = range.top; row <= maxRows; row += 1) {
    const rowValues = [];
    for (let col = range.left; col <= maxCols; col += 1) {
      rowValues.push(String(getDisplayForCell(sheet, sheetId, row, col) ?? ""));
    }
    values.push(rowValues.join(","));
  }
  const suffix = range.bottom - range.top > 2 || range.right - range.left > 2 ? ",..." : "";
  return `{${values.join(";")}${suffix}}`;
}



function replaceFormulaReferenceForDisplay(expression, referenceOrLabel, value) {
  const reference = typeof referenceOrLabel === "object" ? referenceOrLabel : null;
  const labels = reference
    ? [
      reference.label,
      rangeToLabel(reference.range),
      reference.sheetName ? `${quoteSheetNameForFormula(reference.sheetName)}!${rangeToLabel(reference.range)}` : "",
    ]
    : [referenceOrLabel];
  let output = String(expression || "");
  [...new Set(labels.filter(Boolean))].forEach((label) => {
    output = output.replace(new RegExp(formulaReferenceDisplayRegexSource(label), "gi"), String(value));
  });
  return output;
}



function formulaReferenceDisplayRegexSource(label) {
  const split = splitFormulaReferenceToken(String(label || ""));
  const sheetPrefix = split.sheetToken ? `${escapeRegex(split.sheetToken)}!` : "";
  const bodySource = escapeRegex(split.body)
    .replace(/([A-Z]+)/gi, "\\$?$1")
    .replace(/(\\d+)/g, "\\$?$1");
  return `(?<![A-Z0-9_])${sheetPrefix}${bodySource}(?![A-Z0-9_])`;
}



function escapeRegex(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}



function syncEvaluateFormulaDialog(dialog = document.querySelector("[data-data-dialog='formula-evaluate']")) {
  const data = state.dataDialog;
  if (!dialog || data?.type !== "formula-evaluate") return;
  const traceSheetIndex = Number.isInteger(data.trace?.sheetIndex) ? data.trace.sheetIndex : state.activeSheetIndex;
  if (traceSheetIndex !== state.activeSheetIndex && state.model?.sheets?.[traceSheetIndex] && isSheetVisible(state.model.sheets[traceSheetIndex])) {
    setActiveSheet(traceSheetIndex);
  }
  const traceSheet = state.model?.sheets?.[traceSheetIndex] || activeSheet();
  const steps = data.trace?.steps || [];
  const index = Math.min(Math.max(0, Number(data.step) || 0), Math.max(0, steps.length - 1));
  data.step = index;
  const step = steps[index] || {};
  const scope = data.trace?.scopeLabel ? ` - ${data.trace.scopeLabel}` : "";
  dialog.querySelector("[data-formula-evaluate-location]").textContent = `${traceSheet.name}!${columnName(data.trace?.col || data.col)}${data.trace?.row || data.row}${scope}`;
  dialog.querySelector("[data-formula-evaluate-expression]").innerHTML = escapeHtml(step.expression || "");
  dialog.querySelector("[data-formula-evaluate-result]").value = data.trace?.result || "";
  const stepsHost = dialog.querySelector("[data-formula-evaluate-steps]");
  if (stepsHost) {
    stepsHost.innerHTML = steps.map((item, itemIndex) => `
      <div class="${itemIndex === index ? "is-active" : ""}">
        <span>${escapeHtml(String(itemIndex + 1))}</span>${escapeHtml(item.label || "")}
      </div>`).join("");
  }
  const stepInTrace = formulaEvaluateStepInTrace(data, step);
  dialog.querySelector("[data-data-dialog-action='formula-evaluate-step']").disabled = index >= steps.length - 1;
  dialog.querySelector("[data-data-dialog-action='formula-evaluate-step-in']").disabled = !stepInTrace;
  dialog.querySelector("[data-data-dialog-action='formula-evaluate-step-out']").disabled = !(data.traceStack || []).length;
  const range = step.range;
  if (range) {
    const rangeSheetIndex = Number.isInteger(range.sheetIndex) ? range.sheetIndex : traceSheetIndex;
    if (rangeSheetIndex !== state.activeSheetIndex && state.model?.sheets?.[rangeSheetIndex] && isSheetVisible(state.model.sheets[rangeSheetIndex])) {
      setActiveSheet(rangeSheetIndex);
    }
    state.selectionRange = { ...range };
    state.selectionRanges = [{ ...range }];
    state.selected = { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
    scrollRangeIntoView(range);
    renderSheet();
    updateSelectionUi();
  }
}



function handleFormulaEvaluateDialogAction(action, dialog) {
  const data = state.dataDialog;
  if (data?.type !== "formula-evaluate") return false;
  if (action === "formula-evaluate-step") {
    data.step = Math.min((Number(data.step) || 0) + 1, Math.max(0, (data.trace?.steps || []).length - 1));
    syncEvaluateFormulaDialog(dialog);
    return true;
  }
  if (action === "formula-evaluate-restart") {
    data.trace = data.rootTrace || data.trace;
    data.traceStack = [];
    data.step = 0;
    syncEvaluateFormulaDialog(dialog);
    return true;
  }
  if (action === "formula-evaluate-step-in") {
    const steps = data.trace?.steps || [];
    const step = steps[Math.min(Math.max(0, Number(data.step) || 0), Math.max(0, steps.length - 1))] || {};
    const nextTrace = formulaEvaluateStepInTrace(data, step);
    if (!nextTrace) return true;
    data.traceStack = Array.isArray(data.traceStack) ? data.traceStack : [];
    data.traceStack.push({ trace: data.trace, step: data.step });
    data.trace = nextTrace;
    data.step = 0;
    syncEvaluateFormulaDialog(dialog);
    setStatus("数式の内側を検証しています。");
    return true;
  }
  if (action === "formula-evaluate-step-out") {
    const parent = Array.isArray(data.traceStack) ? data.traceStack.pop() : null;
    if (!parent) return true;
    data.trace = parent.trace || data.rootTrace || data.trace;
    data.step = Number(parent.step) || 0;
    syncEvaluateFormulaDialog(dialog);
    setStatus("親の数式に戻りました。");
    return true;
  }
  return false;
}



function formulaEvaluateStepInTrace(data, step) {
  if (!data || !step?.expression) return null;
  const referenceTrace = formulaEvaluateReferenceStepInTrace(step);
  if (referenceTrace) return referenceTrace;
  const subexpression = formulaEvaluateStepInSubexpression(step.expression, data.trace?.formula);
  if (!subexpression) return null;
  const sheet = state.model?.sheets?.[data.trace?.sheetIndex] || activeSheet();
  const label = `${subexpression.name}()`;
  return formulaEvaluationTraceForExpression(sheet, data.trace?.row || data.row, data.trace?.col || data.col, subexpression.text, {
    label,
    scopeLabel: `${subexpression.name} にステップ イン`,
    sourceExpression: step.expression,
  });
}



function formulaEvaluateReferenceStepInTrace(step) {
  const reference = step?.reference;
  const range = reference?.range;
  if (!reference || !range || range.top !== range.bottom || range.left !== range.right) return null;
  const sheet = state.model?.sheets?.[reference.sheetIndex] || null;
  if (!sheet) return null;
  const raw = String(getCellRawInput(sheet, range.top, range.left) || "").trim();
  if (!raw.startsWith("=")) return null;
  return formulaEvaluationTraceForCell(sheet, range.top, range.left, {
    label: reference.label,
    scopeLabel: `${reference.label} にステップ イン`,
  });
}



function formulaEvaluateStepInSubexpression(expression, currentFormula) {
  const calls = formulaFunctionCallsInExpression(expression);
  if (!calls.length) return null;
  const current = normalizeFormulaExpressionForCompare(currentFormula || expression);
  const candidates = calls
    .filter((call) => normalizeFormulaExpressionForCompare(call.text) !== current)
    .filter((call) => !formulaFunctionCallCoversWholeExpression(call, expression))
    .sort((left, right) => right.depth - left.depth || left.start - right.start || left.end - right.end);
  return candidates[0] || null;
}



function formulaFunctionCallCoversWholeExpression(call, expression) {
  const text = String(expression || "");
  const trimmedStart = text.search(/\S/);
  if (trimmedStart < 0) return false;
  const prefixOffset = text[trimmedStart] === "=" ? 1 : 0;
  const expressionStart = trimmedStart + prefixOffset;
  const expressionEnd = text.trimEnd().length;
  return call.start === expressionStart && call.end === expressionEnd;
}



function normalizeFormulaExpressionForCompare(expression) {
  return String(expression || "")
    .trim()
    .replace(/^=/, "")
    .replace(/\s+/g, "")
    .toUpperCase();
}



function formulaFunctionCallsInExpression(expression) {
  const text = String(expression || "");
  const calls = [];
  let depth = 0;
  let singleQuoted = false;
  let doubleQuoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (doubleQuoted) {
      if (char === '"' && next === '"') {
        index += 1;
      } else if (char === '"') {
        doubleQuoted = false;
      }
      continue;
    }
    if (singleQuoted) {
      if (char === "'" && next === "'") {
        index += 1;
      } else if (char === "'") {
        singleQuoted = false;
      }
      continue;
    }
    if (char === '"') {
      doubleQuoted = true;
      continue;
    }
    if (char === "'") {
      singleQuoted = true;
      continue;
    }
    if (char === "(") {
      const name = formulaFunctionNameBeforeOpen(text, index);
      if (name) {
        const closeIndex = matchingFormulaParenIndex(text, index);
        if (closeIndex > index) {
          calls.push({
            name: name.name.toUpperCase(),
            start: name.start,
            end: closeIndex + 1,
            text: text.slice(name.start, closeIndex + 1),
            depth,
          });
        }
      }
      depth += 1;
    } else if (char === ")") {
      depth = Math.max(0, depth - 1);
    }
  }
  return calls;
}



function formulaFunctionNameBeforeOpen(text, openIndex) {
  let end = openIndex - 1;
  while (end >= 0 && /\s/.test(text[end])) end -= 1;
  let start = end;
  while (start >= 0 && /[A-Za-z0-9_.]/.test(text[start])) start -= 1;
  const name = text.slice(start + 1, end + 1);
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) return null;
  const before = start >= 0 ? text[start] : "";
  if (before === "!" || before === "." || /[A-Za-z0-9_]/.test(before)) return null;
  return { name, start: start + 1 };
}



function matchingFormulaParenIndex(text, openIndex) {
  let depth = 0;
  let singleQuoted = false;
  let doubleQuoted = false;
  for (let index = openIndex; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (doubleQuoted) {
      if (char === '"' && next === '"') {
        index += 1;
      } else if (char === '"') {
        doubleQuoted = false;
      }
      continue;
    }
    if (singleQuoted) {
      if (char === "'" && next === "'") {
        index += 1;
      } else if (char === "'") {
        singleQuoted = false;
      }
      continue;
    }
    if (char === '"') {
      doubleQuoted = true;
      continue;
    }
    if (char === "'") {
      singleQuoted = true;
      continue;
    }
    if (char === "(") {
      depth += 1;
    } else if (char === ")") {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return -1;
}



function normalizeFormulaWatches(watches = [], model = state.model) {
  const sheets = model?.sheets || [];
  const seen = new Set();
  return (Array.isArray(watches) ? watches : [])
    .map((watch) => {
      const sheetIndex = Math.max(0, Math.round(Number(watch?.sheetIndex) || 0));
      const sheet = sheets[sheetIndex];
      if (!sheet) return null;
      const row = Math.max(1, Math.round(Number(watch?.row) || 0));
      const col = Math.max(1, Math.round(Number(watch?.col) || 0));
      if (row > Math.max(1, Number(sheet.rowCount) || 1) || col > Math.max(1, Number(sheet.colCount) || 1)) return null;
      const key = `${sheetIndex}:${row}:${col}`;
      if (seen.has(key)) return null;
      seen.add(key);
      return {
        id: String(watch?.id || formulaWatchId(sheetIndex, row, col)),
        sheetIndex,
        row,
        col,
        createdAt: watch?.createdAt || new Date().toISOString(),
      };
    })
    .filter(Boolean);
}



function formulaWatchId(sheetIndex, row, col) {
  return `watch-${sheetIndex}-${row}-${col}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}



function formulaWatchEntries() {
  if (!state.model) return [];
  state.model.formulaWatches = normalizeFormulaWatches(state.model.formulaWatches, state.model);
  return state.model.formulaWatches;
}



function openFormulaWatchWindow() {
  if (!state.model) return;
  let panel = document.querySelector("#formulaWatchWindow");
  if (!panel) {
    panel = document.createElement("div");
    panel.id = "formulaWatchWindow";
    panel.className = "formula-watch-window excel-mini-dialog";
    panel.setAttribute("role", "dialog");
    panel.setAttribute("aria-labelledby", "formulaWatchTitle");
    document.body.append(panel);
  }
  syncFormulaWatchWindow();
  panel.querySelector("[data-formula-watch-action='add']")?.focus({ preventScroll: true });
  setStatus("ウォッチ ウィンドウを表示しました。");
}



function closeFormulaWatchWindow() {
  document.querySelector("#formulaWatchWindow")?.remove();
}



function syncFormulaWatchWindow() {
  const panel = document.querySelector("#formulaWatchWindow");
  if (!panel || !state.model) return;
  const rows = formulaWatchEntries().map(formulaWatchRowModel);
  panel.innerHTML = `
    <div class="excel-mini-titlebar formula-watch-titlebar" id="formulaWatchTitle">
      <span>ウォッチ ウィンドウ</span>
      <button type="button" class="formula-watch-close" data-formula-watch-action="close" aria-label="閉じる">×</button>
    </div>
    <div class="formula-watch-toolbar">
      <button class="excel-mini-button" type="button" data-formula-watch-action="add">ウォッチ式の追加(A)</button>
      <button class="excel-mini-button" type="button" data-formula-watch-action="delete">ウォッチ式の削除(D)</button>
    </div>
    <div class="formula-watch-table-wrap">
      <table class="formula-watch-table">
        <thead><tr><th></th><th>ブック</th><th>シート</th><th>名前</th><th>セル</th><th>値</th><th>数式</th></tr></thead>
        <tbody>
          ${rows.length ? rows.map(formulaWatchRowHtml).join("") : '<tr class="is-empty"><td colspan="7">ウォッチ式はありません。</td></tr>'}
        </tbody>
      </table>
    </div>`;
  if (window.lucide) window.lucide.createIcons({ icons });
}



function formulaWatchRowModel(watch) {
  const sheet = state.model?.sheets?.[watch.sheetIndex] || null;
  const sheetId = getSheetId(sheet?.name);
  const raw = sheet ? getCellRawInput(sheet, watch.row, watch.col) : "";
  const formula = String(raw || "").startsWith("=") ? String(raw) : "";
  return {
    ...watch,
    workbook: state.model?.sourceName || "Book",
    sheetName: sheet?.name || "",
    name: sheet ? definedNameForFormulaWatch(sheet, watch) : "",
    address: `${columnName(watch.col)}${watch.row}`,
    value: sheet ? String(getDisplayForCell(sheet, sheetId, watch.row, watch.col) ?? "") : "",
    formula,
    active: state.selected?.sheetIndex === watch.sheetIndex && state.selected.row === watch.row && state.selected.col === watch.col,
  };
}



function formulaWatchRowHtml(row) {
  return `
    <tr class="${row.active ? "is-active" : ""}" data-formula-watch-row="${escapeAttr(row.id)}">
      <td><input type="checkbox" data-formula-watch-id="${escapeAttr(row.id)}" aria-label="${escapeAttr(row.address)}" /></td>
      <td>${escapeHtml(row.workbook)}</td>
      <td>${escapeHtml(row.sheetName)}</td>
      <td>${escapeHtml(row.name)}</td>
      <td><button type="button" data-formula-watch-action="select" data-formula-watch-id="${escapeAttr(row.id)}">${escapeHtml(row.address)}</button></td>
      <td>${escapeHtml(row.value)}</td>
      <td>${escapeHtml(row.formula)}</td>
    </tr>`;
}



function definedNameForFormulaWatch(sheet, watch) {
  const range = cellRangeForSpecial(watch.row, watch.col);
  const target = `=${absoluteRangeReference(sheet, range)}`.toUpperCase();
  return (state.model?.definedNames || []).find((definition) => {
    if (definition.hidden) return false;
    if (definition.scope === "sheet" && Number(definition.sheetIndex) !== watch.sheetIndex) return false;
    return String(definition.refersTo || "").toUpperCase() === target;
  })?.name || "";
}



function addFormulaWatchesFromSelection() {
  if (!state.model) return;
  const existing = new Set(formulaWatchEntries().map((watch) => `${watch.sheetIndex}:${watch.row}:${watch.col}`));
  const additions = [];
  activeSelectionRanges().forEach((range) => {
    forEachCellInRange(range, (row, col) => {
      if (additions.length >= 200) return;
      const key = `${state.activeSheetIndex}:${row}:${col}`;
      if (existing.has(key)) return;
      existing.add(key);
      additions.push({
        id: formulaWatchId(state.activeSheetIndex, row, col),
        sheetIndex: state.activeSheetIndex,
        row,
        col,
        createdAt: new Date().toISOString(),
      });
    });
  });
  if (!additions.length) {
    showMessageBox("追加できるウォッチ式はありません。", { title: "Cht WebSheet" });
    return;
  }
  state.model.formulaWatches = formulaWatchEntries().concat(additions);
  scheduleWorkbookPersistence();
  syncFormulaWatchWindow();
  setStatus(`${additions.length} 件のウォッチ式を追加しました。`);
}



function deleteSelectedFormulaWatches() {
  const selectedIds = [...document.querySelectorAll("#formulaWatchWindow [data-formula-watch-id]:checked")].map((input) => input.dataset.formulaWatchId);
  if (!selectedIds.length) {
    showMessageBox("削除するウォッチ式を選択してください。", { title: "Cht WebSheet" });
    return;
  }
  const selected = new Set(selectedIds);
  const before = formulaWatchEntries().length;
  state.model.formulaWatches = formulaWatchEntries().filter((watch) => !selected.has(watch.id));
  scheduleWorkbookPersistence();
  syncFormulaWatchWindow();
  setStatus(`${before - state.model.formulaWatches.length} 件のウォッチ式を削除しました。`);
}



function selectFormulaWatchCell(watchId) {
  const watch = formulaWatchEntries().find((item) => item.id === watchId);
  if (!watch) return;
  const sheet = state.model?.sheets?.[watch.sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    showMessageBox("対象のシートは表示されていません。", { title: "Cht WebSheet" });
    return;
  }
  if (watch.sheetIndex !== state.activeSheetIndex) setActiveSheet(watch.sheetIndex);
  const range = cellRangeForSpecial(watch.row, watch.col);
  selectSingleRange(range, `${sheet.name}!${rangeToLabel(range)} を選択しました。`);
  syncFormulaWatchWindow();
}



function handleFormulaWatchWindowAction(event) {
  const action = event.currentTarget.dataset.formulaWatchAction || "";
  if (action === "close") closeFormulaWatchWindow();
  else if (action === "add") addFormulaWatchesFromSelection();
  else if (action === "delete") deleteSelectedFormulaWatches();
  else if (action === "select") selectFormulaWatchCell(event.currentTarget.dataset.formulaWatchId || "");
}



function handleFormulaWatchRowDoubleClick(event) {
  const watchId = event.currentTarget.dataset.formulaWatchRow || "";
  selectFormulaWatchCell(watchId);
}



function openCalculationOptionsMenu(trigger = null) {
  if (!trigger) {
    setCalculationMode("automatic");
    return;
  }
  const current = workbookCalculationMode();
  openHomeCommandMenu(trigger, "calculation", "計算方法の設定", `
    ${formulaCalculationMenuButtonHtml("automatic", "自動(A)", "check", current === "automatic")}
    ${formulaCalculationMenuButtonHtml("except-tables", "データ テーブル以外自動(D)", "table-2", current === "except-tables")}
    ${formulaCalculationMenuButtonHtml("manual", "手動(M)", "settings", current === "manual")}
    <div class="home-command-menu-separator" role="separator"></div>
    <button type="button" role="menuitem" disabled><i data-lucide="info"></i><span>現在: ${escapeHtml(calculationModeLabel(current))}</span></button>
  `);
}



function formulaCalculationMenuButtonHtml(action, label, icon, checked = false) {
  return `<button type="button" role="menuitemradio" aria-checked="${checked ? "true" : "false"}" data-formula-calculation-action="${escapeAttr(action)}"><span class="home-command-menu-check" aria-hidden="true">${checked ? "✓" : ""}</span><i data-lucide="${escapeAttr(icon)}"></i><span>${accessKeyLabelHtml(label)}</span></button>`;
}



function workbookCalculationMode() {
  const calcMode = normalizeWorkbookCalcProperties(state.model?.calcProperties).calcMode;
  return workbookCalculationModeFromCalcProperties({ calcMode });
}



function workbookCalculationModeFromCalcProperties(calcProperties = {}) {
  const calcMode = normalizeWorkbookCalcProperties(calcProperties).calcMode;
  if (calcMode === "manual") return "manual";
  if (calcMode === "autoNoTable") return "except-tables";
  return "automatic";
}



function calculationModeToCalcMode(mode) {
  if (mode === "manual") return "manual";
  if (mode === "except-tables") return "autoNoTable";
  return "auto";
}



function calculationModeLabel(mode) {
  const labels = { automatic: "自動", manual: "手動", "except-tables": "データ テーブル以外自動" };
  return labels[mode] || labels.automatic;
}



function setCalculationMode(mode) {
  if (!state.model) return;
  const normalizedMode = ["automatic", "manual", "except-tables"].includes(mode) ? mode : "automatic";
  const before = captureWorkbookViewState();
  state.model.calcProperties = normalizeWorkbookCalcProperties({
    ...(state.model.calcProperties || {}),
    calcMode: calculationModeToCalcMode(normalizedMode),
    fullCalcOnLoad: normalizedMode !== "manual",
  });
  state.model.calculationMode = normalizedMode;
  scheduleWorkbookPersistence();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    workbookViewBefore: before,
    workbookViewAfter: captureWorkbookViewState(),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: `計算方法 ${calculationModeLabel(normalizedMode)}`,
  });
  syncViewRibbonButtons();
  setStatus(`計算方法: ${calculationModeLabel(normalizedMode)}`);
}



function calculateWorkbookNow() {
  state.hf = buildFormulaEngine(state.model);
  renderWorkbook();
  updateFormulaBarForSelection();
  setStatus("ブック全体を再計算しました。");
}



function calculateActiveSheetNow() {
  state.hf = buildFormulaEngine(state.model);
  renderSheet();
  updateFormulaBarForSelection();
  setStatus(`${activeSheet().name} を再計算しました。`);
}



function toggleActiveSheetGridlines() {
  const sheet = activeSheet();
  const nextValue = sheet.showGridLines === false;
  if (applyGroupedSheetViewOptionChange((targetSheet) => {
    targetSheet.showGridLines = nextValue;
  }, nextValue ? "目盛線を表示しました。" : "目盛線を非表示にしました。", `${activeSheetGroupEditIndexes().length} シート 目盛線`)) {
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.showGridLines = nextValue;
  commitSheetViewOptionChange(sheet, sheetBefore, sheet.showGridLines ? "目盛線を表示しました。" : "目盛線を非表示にしました。", `${sheet.name} 目盛線`, workbookViewBefore);
}



function toggleActiveSheetHeadings() {
  const sheet = activeSheet();
  const nextValue = !sheetHeadingsVisible(sheet);
  if (applyGroupedSheetViewOptionChange((targetSheet) => {
    targetSheet.showHeadings = nextValue;
  }, nextValue ? "見出しを表示しました。" : "見出しを非表示にしました。", `${activeSheetGroupEditIndexes().length} シート 見出し`)) {
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.showHeadings = nextValue;
  commitSheetViewOptionChange(sheet, sheetBefore, sheet.showHeadings ? "見出しを表示しました。" : "見出しを非表示にしました。", `${sheet.name} 見出し`, workbookViewBefore);
}



function toggleFormulaBarVisibility() {
  const workbookViewBefore = captureWorkbookViewState();
  const hidden = !$formulaBar.hasClass("hidden");
  $formulaBar.toggleClass("hidden", hidden);
  state.model.showFormulaBar = !hidden;
  commitWorkbookViewChange(workbookViewBefore, hidden ? "数式バーを非表示にしました。" : "数式バーを表示しました。", "数式バー", {
    render: false,
  });
}



function toggleActiveSheetWhiteSpace() {
  const sheet = activeSheet();
  const nextValue = sheet.showWhiteSpace === false;
  if (applyGroupedSheetViewOptionChange((targetSheet) => {
    targetSheet.showWhiteSpace = nextValue;
    if (nextValue && (targetSheet.viewMode || "normal") === "normal") targetSheet.viewMode = "page-layout";
  }, nextValue ? "ページ余白を表示しました。" : "ページ余白を非表示にしました。", `${activeSheetGroupEditIndexes().length} シート ページ余白`)) {
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.showWhiteSpace = nextValue;
  if (sheet.showWhiteSpace && (sheet.viewMode || "normal") === "normal") sheet.viewMode = "page-layout";
  commitSheetViewOptionChange(
    sheet,
    sheetBefore,
    sheet.showWhiteSpace ? "ページ余白を表示しました。" : "ページ余白を非表示にしました。",
    `${sheet.name} ページ余白`,
    workbookViewBefore,
  );
}



function toggleActiveSheetZeroValues() {
  const sheet = activeSheet();
  const nextValue = sheet.showZeros === false;
  if (applyGroupedSheetViewOptionChange((targetSheet) => {
    targetSheet.showZeros = nextValue;
  }, nextValue ? "ゼロ値を表示しました。" : "ゼロ値を非表示にしました。", `${activeSheetGroupEditIndexes().length} シート ゼロ値`)) {
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.showZeros = nextValue;
  commitSheetViewOptionChange(
    sheet,
    sheetBefore,
    sheet.showZeros ? "ゼロ値を表示しました。" : "ゼロ値を非表示にしました。",
    `${sheet.name} ゼロ値`,
    workbookViewBefore,
  );
}



function toggleActiveSheetObjectsVisibility() {
  const sheet = activeSheet();
  const nextValue = sheet.showObjects === false;
  if (applyGroupedSheetViewOptionChange((targetSheet) => {
    targetSheet.showObjects = nextValue;
    state.model.workbookProperties = normalizeWorkbookProperties({
      ...(state.model.workbookProperties || {}),
      showObjects: nextValue ? "all" : "none",
    });
  }, nextValue ? "オブジェクトを表示しました。" : "オブジェクトを非表示にしました。", `${activeSheetGroupEditIndexes().length} シート オブジェクトの表示`)) {
    if (!nextValue) {
      state.selectedImage = null;
      state.selectedImages = [];
      state.shapePointEditIds = [];
    }
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.showObjects = nextValue;
  state.model.workbookProperties = normalizeWorkbookProperties({
    ...(state.model.workbookProperties || {}),
    showObjects: nextValue ? "all" : "none",
  });
  if (!sheet.showObjects) {
    state.selectedImage = null;
    state.selectedImages = [];
    state.shapePointEditIds = [];
  }
  commitSheetViewOptionChange(
    sheet,
    sheetBefore,
    sheet.showObjects ? "オブジェクトを表示しました。" : "オブジェクトを非表示にしました。",
    `${sheet.name} オブジェクトの表示`,
    workbookViewBefore,
  );
}



function toggleActiveSheetOutlineSymbols() {
  const sheet = activeSheet();
  const nextValue = sheet.showOutlineSymbols === false;
  if (applyGroupedSheetViewOptionChange((targetSheet) => {
    targetSheet.showOutlineSymbols = nextValue;
  }, nextValue ? "アウトライン記号を表示しました。" : "アウトライン記号を非表示にしました。", `${activeSheetGroupEditIndexes().length} シート アウトライン記号`)) {
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.showOutlineSymbols = nextValue;
  commitSheetViewOptionChange(
    sheet,
    sheetBefore,
    sheet.showOutlineSymbols ? "アウトライン記号を表示しました。" : "アウトライン記号を非表示にしました。",
    `${sheet.name} アウトライン記号`,
    workbookViewBefore,
  );
}



function toggleActiveSheetRightToLeft() {
  const sheet = activeSheet();
  const nextValue = sheet.rightToLeft !== true;
  if (applyGroupedSheetViewOptionChange((targetSheet) => {
    targetSheet.rightToLeft = nextValue;
  }, nextValue ? "シートを右から左に表示しました。" : "シートを左から右に表示しました。", `${activeSheetGroupEditIndexes().length} シート 右から左`)) {
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.rightToLeft = nextValue;
  commitSheetViewOptionChange(
    sheet,
    sheetBefore,
    sheet.rightToLeft ? "シートを右から左に表示しました。" : "シートを左から右に表示しました。",
    `${sheet.name} 右から左`,
    workbookViewBefore,
  );
}



function toggleWorkbookWindowShowOption(optionName, label) {
  if (!state.model) return;
  const workbookViewBefore = captureWorkbookViewState();
  const activeId = state.model.activeWindowId || "main";
  const target = mutableWorkbookWindowRecordById(activeId);
  if (!target) return;
  const nextValue = target[optionName] === false;
  target[optionName] = nextValue;
  if (activeId === "main") {
    state.model.workbookView = normalizeWorkbookViewModel({
      ...(state.model.workbookView || {}),
      [optionName]: nextValue,
    }, state.model.sheets?.length || 1);
  }
  syncWorkbookChromeVisibility();
  commitWorkbookViewChange(
    workbookViewBefore,
    nextValue ? `${label}を表示しました。` : `${label}を非表示にしました。`,
    `${label} 表示`,
    { render: false },
  );
  if (optionName === "showHorizontalScroll") updateHorizontalScrollbar();
}



function captureWorkbookViewState() {
  if (state.model) {
    saveActiveWindowViewState();
    updateActiveWorkbookWindowScrollState();
  }
  return {
    showFormulaBar: !$formulaBar.hasClass("hidden"),
    showFormulas: state.showFormulas === true,
    fullScreenView: state.model?.fullScreenView === true,
    workbookProperties: cloneJsonLikeValue(state.model?.workbookProperties || null),
    calcProperties: cloneJsonLikeValue(state.model?.calcProperties || null),
    workbookView: cloneJsonLikeValue(state.model?.workbookView || null),
    customViews: cloneJsonLikeValue(state.model?.customViews || []),
    activeSheetIndex: state.activeSheetIndex,
    activeWindowId: state.model?.activeWindowId || "main",
    mainWindowHidden: state.model?.mainWindowHidden === true,
    mainWindowView: cloneJsonLikeValue(state.model?.mainWindowView || null),
    windowViews: cloneJsonLikeValue(state.model?.windowViews || []),
    viewSideBySide: cloneJsonLikeValue(state.model?.viewSideBySide || null),
    synchronousScrolling: state.model?.synchronousScrolling === true,
    windowArrangement: cloneJsonLikeValue(state.model?.windowArrangement || null),
  };
}



function restoreWorkbookViewState(snapshot = {}) {
  if (!state.model) return;
  if (Number.isFinite(Number(snapshot.activeSheetIndex))) {
    state.activeSheetIndex = Math.max(0, Math.min(state.model.sheets.length - 1, Number(snapshot.activeSheetIndex)));
  }
  state.model.showFormulaBar = snapshot.showFormulaBar !== false;
  $formulaBar.toggleClass("hidden", state.model.showFormulaBar === false);
  if ("showFormulas" in snapshot) {
    state.showFormulas = snapshot.showFormulas === true;
    const sheet = activeSheet();
    if (sheet) sheet.showFormulas = state.showFormulas;
    $toggleFormulasButton
      .prop("disabled", false)
      .attr("title", state.showFormulas ? "計算結果を表示" : "数式を表示")
      .html(`<i data-lucide="${state.showFormulas ? "eye-off" : "eye"}"></i>`);
  }
  if ("fullScreenView" in snapshot) {
    state.model.fullScreenView = snapshot.fullScreenView === true;
    syncFullScreenViewClass();
    syncBrowserFullscreenToModelState();
  }
  if ("customViews" in snapshot) {
    state.model.customViews = Array.isArray(snapshot.customViews)
      ? snapshot.customViews.map(normalizeWorkbookCustomView).filter(Boolean)
      : [];
  }
  if ("workbookProperties" in snapshot) {
    state.model.workbookProperties = normalizeWorkbookProperties(snapshot.workbookProperties);
  }
  if ("calcProperties" in snapshot) {
    state.model.calcProperties = normalizeWorkbookCalcProperties(snapshot.calcProperties);
    state.model.calculationMode = workbookCalculationModeFromCalcProperties(state.model.calcProperties);
  }
  if ("workbookView" in snapshot) {
    state.model.workbookView = normalizeWorkbookViewModel(snapshot.workbookView, state.model.sheets?.length || 1);
  }
  if ("mainWindowHidden" in snapshot) state.model.mainWindowHidden = snapshot.mainWindowHidden === true;
  if ("mainWindowView" in snapshot) {
    state.model.mainWindowView = snapshot.mainWindowView
      ? normalizeWorkbookWindowView({ id: "main", name: state.model.sourceName || "Book1", ...snapshot.mainWindowView })
      : null;
  }
  if ("windowViews" in snapshot) {
    state.model.windowViews = Array.isArray(snapshot.windowViews)
      ? snapshot.windowViews.map(normalizeWorkbookWindowView)
      : [];
  }
  if ("viewSideBySide" in snapshot) {
    state.model.viewSideBySide = snapshot.viewSideBySide && typeof snapshot.viewSideBySide === "object"
      ? cloneJsonLikeValue(snapshot.viewSideBySide)
      : null;
  }
  if ("synchronousScrolling" in snapshot) state.model.synchronousScrolling = snapshot.synchronousScrolling === true;
  if ("windowArrangement" in snapshot) {
    state.model.windowArrangement = snapshot.windowArrangement && typeof snapshot.windowArrangement === "object"
      ? cloneJsonLikeValue(snapshot.windowArrangement)
      : null;
  }
  if ("activeWindowId" in snapshot || "mainWindowHidden" in snapshot || "windowViews" in snapshot) {
    const visibleIds = new Set(workbookWindowRecords().filter((view) => !view.hidden).map((view) => view.id));
    const requestedActiveId = snapshot.activeWindowId || "main";
    state.model.activeWindowId = visibleIds.has(requestedActiveId) ? requestedActiveId : (visibleIds.values().next().value || "main");
  }
  const restoredWindowRecord = restoreActiveWorkbookWindowViewFromModel(snapshot);
  syncWorkbookTitle();
  syncWorkbookChromeVisibility();
  syncViewRibbonButtons();
  return restoredWindowRecord;
}



function restoreActiveWorkbookWindowViewFromModel(snapshot = {}) {
  if (!state.model) return null;
  const fallbackSheetIndex = Math.max(0, Math.min(
    state.model.sheets.length - 1,
    Number.isFinite(Number(snapshot.activeSheetIndex)) ? Number(snapshot.activeSheetIndex) : state.activeSheetIndex,
  ));
  const hasStoredWindowViews = "mainWindowView" in snapshot || "windowViews" in snapshot;
  if (!hasStoredWindowViews) {
    state.activeSheetIndex = fallbackSheetIndex;
    return null;
  }
  const record = storedWorkbookWindowRecordById(state.model.activeWindowId || "main");
  if (!record || record.hidden || !state.model.sheets[record.sheetIndex]) {
    state.activeSheetIndex = fallbackSheetIndex;
    return null;
  }
  state.activeSheetIndex = record.sheetIndex;
  applyWorkbookWindowSheetView(activeSheet(), record);
  restoreWorkbookWindowSelection(record);
  return record;
}



function commitWorkbookViewChange(workbookViewBefore, statusMessage, label, options = {}) {
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    workbookViewBefore,
    workbookViewAfter: captureWorkbookViewState(),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label,
  });
  if (options.render !== false) renderWorkbook();
  syncViewRibbonButtons();
  setStatus(statusMessage);
  scheduleWorkbookPersistence();
}



function captureWorkbookMacroState() {
  return {
    macros: cloneJsonLikeValue(state.model?.macros || []),
    macroRecording: cloneJsonLikeValue(state.model?.macroRecording || null),
    macroRelativeReferences: state.model?.macroRelativeReferences === true,
    macroSecurity: cloneJsonLikeValue(normalizeMacroSecurity(state.model?.macroSecurity)),
  };
}



function restoreWorkbookMacroState(snapshot = {}) {
  if (!state.model) return;
  state.model.macros = Array.isArray(snapshot.macros) ? cloneJsonLikeValue(snapshot.macros) : [];
  state.model.macroRecording = snapshot.macroRecording && typeof snapshot.macroRecording === "object"
    ? cloneJsonLikeValue(snapshot.macroRecording)
    : null;
  state.model.macroRelativeReferences = snapshot.macroRelativeReferences === true;
  state.model.macroSecurity = normalizeMacroSecurity(snapshot.macroSecurity);
  syncViewRibbonButtons();
}



function commitWorkbookMacroChange(macroStateBefore, statusMessage, label, options = {}) {
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    macroStateBefore,
    macroStateAfter: captureWorkbookMacroState(),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label,
  });
  if (options.render === true) renderWorkbook();
  syncViewRibbonButtons();
  setStatus(statusMessage);
  scheduleWorkbookPersistence();
}



function toggleActiveSheetRuler() {
  const sheet = activeSheet();
  const nextValue = sheet.showRuler !== true;
  if (applyGroupedSheetViewOptionChange((targetSheet) => {
    targetSheet.showRuler = nextValue;
    if (targetSheet.showRuler && (targetSheet.viewMode || "normal") === "normal") targetSheet.viewMode = "page-layout";
  }, nextValue ? "ルーラーを表示しました。" : "ルーラーを非表示にしました。", `${activeSheetGroupEditIndexes().length} シート ルーラー`)) {
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.showRuler = nextValue;
  if (sheet.showRuler && (sheet.viewMode || "normal") === "normal") sheet.viewMode = "page-layout";
  commitSheetViewOptionChange(sheet, sheetBefore, sheet.showRuler ? "ルーラーを表示しました。" : "ルーラーを非表示にしました。", `${sheet.name} ルーラー`, workbookViewBefore);
}



function commitSheetViewOptionChange(sheet, sheetBefore, statusMessage, label, workbookViewBefore = null) {
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



function currentSheetViewStateFields(sheet) {
  const viewMode = normalizedSheetViewMode(sheet?.viewMode);
  return {
    viewMode,
    zoomScale: sheet?.zoomScale || 100,
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
    hiddenRows: hiddenIndexesForModels(sheet?.rows || []),
    hiddenColumns: hiddenIndexesForModels(sheet?.columns || []),
    autoFilter: cloneJsonLikeValue(sheet?.autoFilter),
    sortState: cloneJsonLikeValue(sheet?.sortState),
  };
}



function refreshSheetViewSnapshotForSheet(sheet, sheetIndex = state.activeSheetIndex) {
  if (!sheet) return null;
  const view = activeSheetView(sheet);
  if (!view) return null;
  if (sheetIndex === state.activeSheetIndex) return refreshActiveSheetViewSnapshot(sheet);
  view.snapshot = normalizeSheetViewSnapshot({
    ...(view.snapshot || {}),
    ...currentSheetViewStateFields(sheet),
  });
  view.updatedAt = new Date().toISOString();
  return view;
}



function commitGroupedSheetViewOptionChange(sheetIndexes, sheetsBefore, statusMessage, label, workbookViewBefore = null) {
  const indexes = [...new Set((sheetIndexes || []).filter((index) => state.model?.sheets?.[index]))];
  indexes.forEach((sheetIndex) => refreshSheetViewSnapshotForSheet(state.model.sheets[sheetIndex], sheetIndex));
  saveActiveWindowViewState();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetsBefore,
    sheetsAfter: captureSheetsForHistory(indexes),
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
  updateSheetFooterUi();
  syncViewRibbonButtons();
  setStatus(statusMessage);
  scheduleWorkbookPersistence();
}



function applyGroupedSheetViewOptionChange(mutator, statusMessage, label) {
  if (!state.model || typeof mutator !== "function") return false;
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const workbookViewBefore = captureWorkbookViewState();
  const sheetsBefore = captureSheetsForHistory(sheetIndexes);
  const beforeState = JSON.stringify(sheetIndexes.map((sheetIndex) => currentSheetViewStateFields(state.model.sheets[sheetIndex])));
  sheetIndexes.forEach((sheetIndex) => mutator(state.model.sheets[sheetIndex], sheetIndex));
  const afterState = JSON.stringify(sheetIndexes.map((sheetIndex) => currentSheetViewStateFields(state.model.sheets[sheetIndex])));
  if (beforeState === afterState) {
    renderWorkbook();
    updateFormulaBarForSelection();
    updateSheetFooterUi();
    syncViewRibbonButtons();
    setStatus(`${sheetIndexes.length} 個のシート：変更はありません。`);
    return true;
  }
  commitGroupedSheetViewOptionChange(sheetIndexes, sheetsBefore, `${sheetIndexes.length} 個のシート：${statusMessage}`, label, workbookViewBefore);
  return true;
}



function sheetHeadingsVisible(sheet) {
  return sheet?.showHeadings !== false;
}



function activeSheetView(sheet = activeSheet()) {
  if (!sheet?.activeSheetViewId || !Array.isArray(sheet.sheetViews)) return null;
  return sheet.sheetViews.find((view) => view.id === sheet.activeSheetViewId) || null;
}



function sheetViewDisplayName(sheet, view = activeSheetView(sheet)) {
  if (!view) return "既定";
  return `${view.name || "シート ビュー"}${view.temporary ? " (一時)" : ""}`;
}



function refreshActiveSheetViewSnapshot(sheet = activeSheet()) {
  const view = activeSheetView(sheet);
  if (!sheet || !view) return null;
  view.snapshot = normalizeSheetViewSnapshot(captureSheetViewSnapshot(sheet));
  view.updatedAt = new Date().toISOString();
  return view;
}



function uniqueSheetViewName(sheet) {
  const existing = new Set((sheet?.sheetViews || []).map((view) => String(view.name || "").toLocaleLowerCase("ja-JP")));
  let index = existing.size + 1;
  let name = `ビュー ${index}`;
  while (existing.has(name.toLocaleLowerCase("ja-JP"))) {
    index += 1;
    name = `ビュー ${index}`;
  }
  return name;
}



function captureSheetViewSnapshot(sheet = activeSheet()) {
  updateActiveWorkbookWindowScrollState();
  const host = $sheetHost[0];
  const viewMode = normalizedSheetViewMode(sheet?.viewMode);
  const selectionRanges = activeSelectionRanges().map((range) => ({ ...range }));
  const selectionRange = selectionRanges[0] || activeSelectionRange();
  const activeScroll = {
    top: Math.max(0, Math.round(host?.scrollTop || 0)),
    left: Math.max(0, Math.round(host?.scrollLeft || 0)),
  };
  return {
    viewMode,
    zoomScale: sheet?.zoomScale || 100,
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
    selectedCell: state.selected ? { ...state.selected } : null,
    selectionRange: selectionRange ? { ...selectionRange } : null,
    selectionRanges,
    activeScroll,
    topLeftCell: topLeftCellFromScrollRecord(sheet, activeScroll),
    hiddenRows: hiddenIndexesForModels(sheet?.rows || []),
    hiddenColumns: hiddenIndexesForModels(sheet?.columns || []),
    autoFilter: cloneJsonLikeValue(sheet?.autoFilter),
    sortState: cloneJsonLikeValue(sheet?.sortState),
  };
}



function applySheetViewSnapshot(sheet, snapshot) {
  const view = normalizeSheetViewSnapshot(snapshot);
  sheet.viewMode = view.viewMode || "normal";
  applySheetZoomWithoutHistory(sheet, view.zoomScale || sheet.zoomScale || 100);
  sheet.showGridLines = view.showGridLines !== false;
  sheet.showHeadings = view.showHeadings !== false;
  sheet.showRuler = normalizedSheetShowRuler(view.showRuler, sheet.viewMode);
  sheet.showObjects = view.showObjects !== false;
  sheet.showOutlineSymbols = view.showOutlineSymbols !== false;
  sheet.showFormulas = view.showFormulas === true;
  sheet.showZeros = view.showZeros !== false;
  sheet.showWhiteSpace = view.showWhiteSpace !== false;
  sheet.rightToLeft = view.rightToLeft === true;
  state.showFormulas = sheet.showFormulas === true;
  sheet.freezePanes = normalizePaneState(view.freezePanes);
  sheet.splitPane = normalizePaneState(view.splitPane);
  applyHiddenIndexesToModels(sheet.rows, view.hiddenRows);
  applyHiddenIndexesToModels(sheet.columns, view.hiddenColumns);
  sheet.autoFilter = cloneJsonLikeValue(view.autoFilter);
  sheet.sortState = cloneJsonLikeValue(view.sortState);
  queueSheetViewScrollRestore(sheet, view.activeScroll, view.topLeftCell);
  restoreCustomViewSelection({
    activeSheetIndex: state.activeSheetIndex,
    selectedCell: view.selectedCell,
    selectionRange: view.selectionRange,
    selectionRanges: view.selectionRanges,
  });
}



function queueSheetViewScrollRestore(sheet, scroll, topLeftCell = null) {
  state.pendingSheetViewScroll = restoreScrollRecordForView(sheet, scroll, topLeftCell);
}



function restorePendingSheetViewScroll() {
  const scroll = state.pendingSheetViewScroll;
  state.pendingSheetViewScroll = null;
  if (!scroll) return;
  restoreWorkbookWindowScrollStable(scroll);
}



function createSheetView() {
  const sheet = activeSheet();
  if (!sheet) return;
  sheet.sheetViews = Array.isArray(sheet.sheetViews) ? sheet.sheetViews : [];
  const sheetBefore = captureSheetForHistory(sheet);
  const snapshot = captureSheetViewSnapshot(sheet);
  const id = uniqueWorksheetObjectId("sheet-view");
  const view = {
    id,
    name: uniqueSheetViewName(sheet),
    temporary: true,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    baseView: normalizeSheetViewSnapshot(snapshot),
    snapshot: normalizeSheetViewSnapshot(snapshot),
  };
  sheet.sheetViews.push(view);
  sheet.activeSheetViewId = id;
  commitSheetViewChange(sheet, sheetBefore, `一時シート ビューを開始しました：${view.name}`, `${sheet.name} シート ビュー新規`);
}



function keepActiveSheetView() {
  const sheet = activeSheet();
  const view = activeSheetView(sheet);
  if (!view) {
    showMessageBox("保持するシート ビューがありません。", { title: "Cht WebSheet" });
    return;
  }
  if (view.temporary !== true) {
    showMessageBox("このシート ビューは既に保持されています。変更を保存するには、オプションから「現在のビューを更新」を実行してください。", { title: "Cht WebSheet" });
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  view.temporary = false;
  view.snapshot = normalizeSheetViewSnapshot(captureSheetViewSnapshot(sheet));
  view.updatedAt = new Date().toISOString();
  commitSheetViewChange(sheet, sheetBefore, `シート ビューを保持しました：${view.name}`, `${sheet.name} シート ビュー保持`);
}



function exitActiveSheetView() {
  const sheet = activeSheet();
  const view = activeSheetView(sheet);
  if (!view) {
    setStatus("既定のシート ビューを表示しています。");
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  if (view.baseView) applySheetViewSnapshot(sheet, view.baseView);
  if (view.temporary) {
    sheet.sheetViews = (sheet.sheetViews || []).filter((candidate) => candidate.id !== view.id);
  }
  sheet.activeSheetViewId = null;
  commitSheetViewChange(sheet, sheetBefore, "シート ビューを終了しました。", `${sheet.name} シート ビュー終了`);
}



function activateSheetView(viewId) {
  const sheet = activeSheet();
  const view = (sheet?.sheetViews || []).find((candidate) => candidate.id === viewId);
  if (!sheet || !view) return;
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.activeSheetViewId = view.id;
  applySheetViewSnapshot(sheet, view.snapshot || view);
  sheet.activeSheetViewId = view.id;
  commitSheetViewChange(sheet, sheetBefore, `シート ビューを表示しました：${view.name}`, `${sheet.name} シート ビュー切り替え`);
}



function deleteActiveSheetView() {
  const sheet = activeSheet();
  const view = activeSheetView(sheet);
  if (!view) {
    showMessageBox("削除するシート ビューがありません。", { title: "Cht WebSheet" });
    return;
  }
  showMessageBox(`シート ビュー「${view.name}」を削除しますか?`, {
    title: "Cht WebSheet",
    tone: "warning",
    buttons: [
      {
        label: "削除(D)",
        result: "delete",
        callback: () => {
          const sheetBefore = captureSheetForHistory(sheet);
          if (view.baseView) applySheetViewSnapshot(sheet, view.baseView);
          sheet.sheetViews = (sheet.sheetViews || []).filter((candidate) => candidate.id !== view.id);
          sheet.activeSheetViewId = null;
          commitSheetViewChange(sheet, sheetBefore, `シート ビューを削除しました：${view.name}`, `${sheet.name} シート ビュー削除`);
        },
      },
      { label: "キャンセル", result: "cancel" },
    ],
    cancelResult: "cancel",
  });
}



function updateActiveSheetView() {
  const sheet = activeSheet();
  const view = activeSheetView(sheet);
  if (!view) {
    showMessageBox("更新するシート ビューがありません。", { title: "Cht WebSheet" });
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  view.snapshot = normalizeSheetViewSnapshot(captureSheetViewSnapshot(sheet));
  view.temporary = false;
  view.updatedAt = new Date().toISOString();
  commitSheetViewChange(sheet, sheetBefore, `シート ビューを更新しました：${view.name}`, `${sheet.name} シート ビュー更新`);
}



function duplicateActiveSheetView() {
  const sheet = activeSheet();
  const view = activeSheetView(sheet);
  if (!sheet || !view) {
    showMessageBox("コピーするシート ビューがありません。", { title: "Cht WebSheet" });
    return;
  }
  sheet.sheetViews = Array.isArray(sheet.sheetViews) ? sheet.sheetViews : [];
  const sheetBefore = captureSheetForHistory(sheet);
  const name = uniqueSheetViewCopyName(sheet, view.name || "シート ビュー");
  const snapshot = normalizeSheetViewSnapshot(captureSheetViewSnapshot(sheet));
  const copy = {
    id: uniqueWorksheetObjectId("sheet-view"),
    name,
    temporary: false,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    baseView: normalizeSheetViewSnapshot(view.baseView || snapshot),
    snapshot,
  };
  sheet.sheetViews.push(copy);
  sheet.activeSheetViewId = copy.id;
  commitSheetViewChange(sheet, sheetBefore, `シート ビューのコピーを作成しました：${name}`, `${sheet.name} シート ビュー コピー`);
}



function uniqueSheetViewCopyName(sheet, sourceName) {
  const existing = new Set((sheet?.sheetViews || []).map((view) => String(view.name || "").toLocaleLowerCase("ja-JP")));
  const base = `${sourceName} コピー`;
  if (!existing.has(base.toLocaleLowerCase("ja-JP"))) return base;
  let index = 2;
  let name = `${base} (${index})`;
  while (existing.has(name.toLocaleLowerCase("ja-JP"))) {
    index += 1;
    name = `${base} (${index})`;
  }
  return name;
}



function openSheetViewRenameDialog() {
  const sheet = activeSheet();
  const view = activeSheetView(sheet);
  if (!view) {
    showMessageBox("名前を変更するシート ビューがありません。", { title: "Cht WebSheet" });
    return;
  }
  closeDataDialog({ keepState: true });
  const overlay = $(`
    <div class="excel-dialog-overlay" id="dataDialog">
      <div class="excel-mini-dialog data-dialog sheet-view-name-dialog" role="dialog" aria-modal="true" aria-labelledby="sheetViewNameTitle" data-data-dialog="sheet-view-name">
        <div class="excel-mini-titlebar" id="sheetViewNameTitle">シート ビュー名の変更</div>
        <div class="excel-mini-body data-dialog-body">
          <label class="excel-dialog-field"><span class="excel-dialog-caption">名前(N):</span>
            <input type="text" data-data-field="sheetViewName" value="${escapeAttr(view.name || "")}" />
          </label>
        </div>
        <div class="excel-mini-actions">
          <button class="excel-mini-button is-primary" type="button" data-data-dialog-action="sheet-view-name-ok">OK</button>
          <button class="excel-mini-button" type="button" data-data-dialog-action="cancel">キャンセル</button>
        </div>
      </div>
    </div>`);
  $("body").append(overlay);
  state.dataDialog = { type: "sheet-view-name", viewId: view.id };
  overlay.find("[data-data-field='sheetViewName']").trigger("focus").trigger("select");
}



function applySheetViewNameDialog(dialog) {
  const sheet = activeSheet();
  const viewId = state.dataDialog?.viewId;
  const view = (sheet?.sheetViews || []).find((candidate) => candidate.id === viewId);
  if (!sheet || !view) return;
  const name = String(dialog?.querySelector("[data-data-field='sheetViewName']")?.value || "").trim();
  if (!name) {
    showMessageBox("シート ビュー名を入力してください。", { title: "Cht WebSheet" });
    return;
  }
  const duplicate = (sheet.sheetViews || []).some((candidate) => candidate.id !== view.id && String(candidate.name || "").toLocaleLowerCase("ja-JP") === name.toLocaleLowerCase("ja-JP"));
  if (duplicate) {
    showMessageBox("同じ名前のシート ビューが既に存在します。", { title: "Cht WebSheet" });
    return;
  }
  const sheetBefore = captureSheetForHistory(sheet);
  view.name = name;
  view.temporary = false;
  view.updatedAt = new Date().toISOString();
  closeDataDialog();
  commitSheetViewChange(sheet, sheetBefore, `シート ビュー名を変更しました：${name}`, `${sheet.name} シート ビュー名変更`);
}



function commitSheetViewChange(sheet, sheetBefore, statusMessage, label) {
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label,
  });
  renderWorkbook();
  restorePendingSheetViewScroll();
  saveActiveWindowViewState();
  requestAnimationFrame(() => {
    saveActiveWindowViewState();
    syncViewRibbonButtons();
    updateSheetFooterUi();
  });
  updateFormulaBarForSelection();
  updateSheetFooterUi();
  syncViewRibbonButtons();
  const activeView = activeSheetView(sheet);
  setStatus(activeView ? `${statusMessage}（${sheetViewDisplayName(sheet, activeView)}）` : statusMessage);
  scheduleWorkbookPersistence();
}



function openSheetViewOptionsMenu(trigger = null) {
  const sheet = activeSheet();
  const views = sheet?.sheetViews || [];
  const active = activeSheetView(sheet);
  const canKeep = active?.temporary === true;
  const rows = [
    viewCommandMenuButtonHtml("sheet-view:new", "新規(N)", "badge-plus", false, {
      description: "現在の表示状態から新しい一時シート ビューを開始します。",
    }),
    viewCommandMenuButtonHtml("sheet-view:keep", "保持(S)", "save", !canKeep, {
      description: canKeep ? "現在の一時シート ビューを保存します。" : active ? "保存済みシート ビューは「現在のビューを更新」で保存します。" : "シート ビューを選択すると保持できます。",
    }),
    viewCommandMenuButtonHtml("sheet-view:exit", "終了(X)", "log-out", !active, {
      description: active ? "現在のシート ビューを終了して既定の表示に戻ります。" : "シート ビューは開始されていません。",
    }),
    '<div class="home-command-menu-separator" role="separator"></div>',
    viewCommandMenuButtonHtml("sheet-view:default", "既定", "grid-2x2", false, {
      checkable: true,
      checked: !active,
      description: "フィルターや非表示設定を通常の表示に戻します。",
    }),
    ...views.map((view) => viewCommandMenuButtonHtml(`sheet-view:activate:${view.id}`, `${view.name}${view.temporary ? " (一時)" : ""}`, "eye", false, {
      checkable: true,
      checked: view.id === sheet.activeSheetViewId,
      description: sheetViewMenuDescription(view),
    })),
    '<div class="home-command-menu-separator" role="separator"></div>',
    viewCommandMenuButtonHtml("sheet-view:update", "現在のビューを更新(U)", "refresh-cw", !active, { description: active ? "現在のフィルター、非表示、スクロール位置をこのビューに保存します。" : "シート ビューを選択すると更新できます。" }),
    viewCommandMenuButtonHtml("sheet-view:duplicate", "コピーの作成(C)", "copy", !active, { description: active ? "現在のシート ビューを別名でコピーします。" : "シート ビューを選択するとコピーできます。" }),
    viewCommandMenuButtonHtml("sheet-view:rename", "名前の変更(R)...", "text-cursor-input", !active, { description: active ? "現在のシート ビュー名を変更します。" : "シート ビューを選択すると名前を変更できます。" }),
    viewCommandMenuButtonHtml("sheet-view:delete", "削除(D)", "delete", !active, { description: active ? "現在のシート ビューを削除します。" : "シート ビューを選択すると削除できます。" }),
  ].join("");
  openViewCommandMenu(trigger, "sheet-view", "シート ビュー", rows);
}



function sheetViewMenuDescription(view) {
  const prefix = view?.temporary ? "一時シート ビュー" : "保存済みシート ビュー";
  const updated = view?.updatedAt ? new Date(view.updatedAt) : null;
  const updatedLabel = updated && !Number.isNaN(updated.getTime())
    ? updated.toLocaleString("ja-JP", { month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit" })
    : "";
  return updatedLabel ? `${prefix} - 更新 ${updatedLabel}` : prefix;
}



function macroSecurityLevelLabel(level) {
  const option = MACRO_SECURITY_OPTIONS.find((item) => item.value === level);
  return option?.label || MACRO_SECURITY_OPTIONS.find((item) => item.value === DEFAULT_MACRO_SECURITY_LEVEL)?.label || "";
}



function openViewMacrosMenu(trigger = null) {
  const recording = Boolean(state.model?.macroRecording);
  const relativeRefs = state.model?.macroRelativeReferences === true;
  const macroSecurity = normalizeMacroSecurity(state.model?.macroSecurity);
  const rows = [
    viewCommandMenuButtonHtml("macro-show", "マクロの表示(V)", "play-square", false, { description: "マクロの一覧を表示して、実行、編集、作成、削除を行います。" }),
    viewCommandMenuButtonHtml(recording ? "macro-stop-recording" : "macro-record", recording ? "記録終了(R)" : "マクロの記録(R)...", recording ? "square" : "save", false, { description: recording ? "現在のマクロ記録を停止します。" : "操作を記録する新しいマクロを作成します。" }),
    viewCommandMenuButtonHtml("macro-use-relative-refs", "相対参照で記録(U)", "move-right", false, { checkable: true, checked: relativeRefs, description: "記録中のセル移動を相対参照として扱います。" }),
    '<div class="home-command-menu-separator" role="separator"></div>',
    viewCommandMenuButtonHtml("macro-security", "マクロのセキュリティ(S)...", "settings", false, { description: `現在: ${macroSecurityLevelLabel(macroSecurity.level)}` }),
  ].join("");
  openViewCommandMenu(trigger, "macros", "マクロ", rows);
}



async function toggleFullScreenView() {
  if (!state.model) return;
  const workbookViewBefore = captureWorkbookViewState();
  const nextFullScreen = state.model.fullScreenView !== true;
  state.model.fullScreenView = nextFullScreen;
  syncFullScreenViewClass();
  try {
    if (nextFullScreen && !document.fullscreenElement) {
      await document.documentElement.requestFullscreen?.();
    } else if (!nextFullScreen && document.fullscreenElement) {
      await document.exitFullscreen?.();
    }
  } catch {
    restoreWorkbookViewState(workbookViewBefore);
    syncFullScreenViewClass();
    syncViewRibbonButtons();
    scheduleWorkbookPersistence();
    showMessageBox("全画面表示に切り替えられませんでした。ブラウザーの全画面表示の許可を確認してください。", { title: "Cht WebSheet" });
    return;
  }
  commitWorkbookViewChange(
    workbookViewBefore,
    nextFullScreen ? "全画面表示に切り替えました。" : "全画面表示を解除しました。",
    "全画面表示",
    { render: false },
  );
}



function handleFullscreenChange() {
  if (!state.model) return;
  const isFullScreen = Boolean(document.fullscreenElement);
  if (state.model.fullScreenView === isFullScreen) return;
  state.model.fullScreenView = isFullScreen;
  syncFullScreenViewClass();
  syncViewRibbonButtons();
  scheduleWorkbookPersistence();
  setStatus(isFullScreen ? "全画面表示に切り替えました。" : "全画面表示を解除しました。");
}



function handleFullScreenEscape(event) {
  if (event.key !== "Escape" || !state.model?.fullScreenView) return;
  if (
    document.querySelector("#dataDialog, #formatCellsDialogOverlay, #nameDialogOverlay, #linkDialogOverlay, #objectLinkDialogOverlay, #symbolDialogOverlay, #messageBox, #zoomDialog, #findDialog, #goToDialog, #specialDialog")
  ) return;
  event.preventDefault();
  void exitFullScreenView();
}



async function exitFullScreenView() {
  if (!state.model) return;
  const workbookViewBefore = captureWorkbookViewState();
  if (workbookViewBefore.fullScreenView === false && !document.fullscreenElement) return;
  state.model.fullScreenView = false;
  syncFullScreenViewClass();
  try {
    if (document.fullscreenElement) await document.exitFullscreen?.();
  } catch {
    // The browser may already be leaving fullscreen; keep the WebSheet view state in sync.
  }
  commitWorkbookViewChange(workbookViewBefore, "全画面表示を解除しました。", "全画面表示解除", { render: false });
}



function syncFullScreenViewClass() {
  document.body.classList.toggle("is-websheet-fullscreen", state.model?.fullScreenView === true);
}



function syncWorkbookChromeVisibility() {
  const showHorizontal = activeWorkbookWindowShowOption("showHorizontalScroll");
  const showVertical = activeWorkbookWindowShowOption("showVerticalScroll");
  const showTabs = activeWorkbookWindowShowOption("showSheetTabs");
  document.body.classList.toggle("workbook-hide-horizontal-scrollbar", !showHorizontal);
  document.body.classList.toggle("workbook-hide-vertical-scrollbar", !showVertical);
  document.body.classList.toggle("workbook-hide-sheet-tabs", !showTabs);
}



function syncBrowserFullscreenToModelState() {
  if (!state.model) return;
  const shouldBeFullscreen = state.model.fullScreenView === true;
  try {
    if (shouldBeFullscreen && !document.fullscreenElement) {
      void document.documentElement.requestFullscreen?.();
    } else if (!shouldBeFullscreen && document.fullscreenElement) {
      void document.exitFullscreen?.();
    }
  } catch {
    // Fullscreen requests can be rejected outside direct user activation; the CSS view state is still restored.
  }
}



function syncViewRibbonButtons() {
  const sheet = state.model ? activeSheet() : null;
  const viewMode = sheet?.viewMode || "normal";
  const records = state.model ? workbookWindowRecords() : [];
  const hiddenWindows = records.filter((view) => view.hidden).length;
  const visibleRecords = records.filter((view) => !view.hidden);
  const visibleWindows = visibleRecords.length;
  const sideBySide = Boolean(normalizedViewSideBySidePair(visibleRecords));
  const sheetView = activeSheetView(sheet);
  const sheetViewCanKeep = sheetView?.temporary === true;
  const customViews = Array.isArray(state.model?.customViews) ? state.model.customViews : [];
  const customViewsBlocked = Boolean(customViewsUnavailableSummary());
  const zoom = Math.round(Number(sheet?.zoomScale || 100));
  const rulerVisible = normalizedSheetShowRuler(sheet?.showRuler, viewMode);
  const showHorizontalScrollbar = activeWorkbookWindowShowOption("showHorizontalScroll");
  const showVerticalScrollbar = activeWorkbookWindowShowOption("showVerticalScroll");
  const showSheetTabs = activeWorkbookWindowShowOption("showSheetTabs");
  setRibbonCommandActive("view-normal", viewMode === "normal");
  setRibbonCommandActive("view-page-layout", viewMode === "page-layout");
  setRibbonCommandActive("view-page-break", viewMode === "page-break");
  setRibbonCommandActive("view-full-screen", state.model?.fullScreenView === true);
  setRibbonCommandActive("view-sheet-view-new", sheetView?.temporary === true);
  setRibbonCommandActive("view-sheet-view-options", Boolean(sheetView));
  setRibbonCommandActive("view-100", zoom === 100);
  setRibbonCommandCheckbox("view-ruler", rulerVisible);
  setRibbonCommandCheckbox("view-gridlines", sheet?.showGridLines !== false);
  setRibbonCommandCheckbox("view-headings", sheetHeadingsVisible(sheet));
  setRibbonCommandCheckbox("view-formula-bar", !$formulaBar.hasClass("hidden"));
  setRibbonCommandCheckbox("view-white-space", sheet?.showWhiteSpace !== false);
  setRibbonCommandCheckbox("view-zero-values", sheet?.showZeros !== false);
  setRibbonCommandCheckbox("view-objects", sheet?.showObjects !== false);
  setRibbonCommandCheckbox("view-outline-symbols", sheet?.showOutlineSymbols !== false);
  setRibbonCommandCheckbox("view-right-to-left", sheet?.rightToLeft === true);
  setRibbonCommandCheckbox("view-horizontal-scrollbar", showHorizontalScrollbar);
  setRibbonCommandCheckbox("view-vertical-scrollbar", showVerticalScrollbar);
  setRibbonCommandCheckbox("view-sheet-tabs", showSheetTabs);
  setRibbonCommandActive("view-freeze-panes", Boolean(sheet?.freezePanes));
  setRibbonCommandActive("view-freeze-menu", Boolean(sheet?.freezePanes));
  setRibbonCommandActive("view-split", Boolean(sheet?.splitPane));
  setRibbonCommandActive("view-side-by-side", sideBySide);
  setRibbonCommandActive("view-sync-scrolling", sideBySide && state.model?.synchronousScrolling === true);
  setRibbonCommandActive("view-macros", Boolean(state.model?.macroRecording));
  setRibbonCommandActive("view-macros-menu", Boolean(state.model?.macroRecording));
  setRibbonCommandDisabled("view-ruler", viewMode !== "page-layout");
  setRibbonCommandDisabled("view-white-space", !sheet);
  setRibbonCommandDisabled("view-custom-views", customViewsBlocked);
  setRibbonCommandDisabled("view-sheet-view-new", !sheet);
  setRibbonCommandDisabled("view-sheet-view-keep", !sheetViewCanKeep);
  setRibbonCommandDisabled("view-sheet-view-exit", !sheetView);
  setRibbonCommandDisabled("view-sheet-view-options", !sheet);
  setRibbonCommandDisabled("view-hide-window", visibleWindows <= 1);
  setRibbonCommandDisabled("view-unhide-window", hiddenWindows <= 0);
  setRibbonCommandDisabled("view-side-by-side", visibleWindows < 2 && !sideBySide);
  setRibbonCommandDisabled("view-sync-scrolling", !sideBySide);
  setRibbonCommandDisabled("view-reset-window-position", !sideBySide);
  setRibbonCommandLabel("view-freeze-panes", sheet?.freezePanes ? "ウィンドウ枠固定の解除" : "ウィンドウ枠の固定");
  setRibbonCommandIcon("view-freeze-panes", sheet?.freezePanes ? "unlock-keyhole" : "panel-top");
  setRibbonCommandTitle("view-freeze-panes", sheet?.freezePanes ? "固定されたウィンドウ枠を解除します" : "現在の選択位置に基づいて行と列を固定します");
  setRibbonCommandTitle("view-100", zoom === 100 ? "現在の倍率は 100% です" : "倍率を 100% に戻します");
  setRibbonCommandTitle("view-ruler", viewMode !== "page-layout" ? "ルーラーはページ レイアウト表示で使用できます" : rulerVisible ? "ルーラーを非表示にします" : "ルーラーを表示します");
  setRibbonCommandTitle("view-gridlines", sheet?.showGridLines === false ? "目盛線を表示します" : "目盛線を非表示にします");
  setRibbonCommandTitle("view-headings", sheetHeadingsVisible(sheet) ? "行番号と列番号を非表示にします" : "行番号と列番号を表示します");
  setRibbonCommandTitle("view-formula-bar", $formulaBar.hasClass("hidden") ? "数式バーを表示します" : "数式バーを非表示にします");
  setRibbonCommandTitle(
    "view-custom-views",
    customViewsBlocked
      ? customViewsUnavailableMessage(customViewsUnavailableSummary())
      : customViews.length
        ? `ユーザー設定のビューを表示します（${customViews.length}）`
        : "ユーザー設定のビューを追加または表示します",
  );
  setRibbonCommandTitle("view-sheet-view-new", sheet ? "現在の表示状態から新しい一時シート ビューを開始します" : "シートがありません");
  setRibbonCommandTitle("view-sheet-view-keep", sheetViewCanKeep ? "現在の一時シート ビューを保持します" : sheetView ? "保存済みシート ビューはオプションから更新します" : "保持するシート ビューがありません");
  setRibbonCommandTitle("view-sheet-view-exit", sheetView ? `シート ビューを終了します：${sheetViewDisplayName(sheet, sheetView)}` : "シート ビューは開始されていません");
  setRibbonCommandTitle("view-sheet-view-options", sheetView ? `シート ビュー オプション：${sheetViewDisplayName(sheet, sheetView)}` : "シート ビュー オプション");
  setRibbonCommandTitle("view-white-space", sheet?.showWhiteSpace === false ? "ページ レイアウト表示の余白を表示します" : "ページ レイアウト表示の余白を非表示にします");
  setRibbonCommandTitle("view-zero-values", sheet?.showZeros === false ? "値が 0 のセルを表示します" : "値が 0 のセルを非表示にします");
  setRibbonCommandTitle("view-objects", sheet?.showObjects === false ? "シート上のオブジェクトを表示します" : "シート上のオブジェクトを非表示にします");
  setRibbonCommandTitle("view-outline-symbols", sheet?.showOutlineSymbols === false ? "アウトライン記号を表示します" : "アウトライン記号を非表示にします");
  setRibbonCommandTitle("view-right-to-left", sheet?.rightToLeft === true ? "シートを左から右に表示します" : "シートを右から左に表示します");
  setRibbonCommandTitle("view-horizontal-scrollbar", showHorizontalScrollbar ? "横スクロールバーを非表示にします" : "横スクロールバーを表示します");
  setRibbonCommandTitle("view-vertical-scrollbar", showVerticalScrollbar ? "縦スクロールバーを非表示にします" : "縦スクロールバーを表示します");
  setRibbonCommandTitle("view-sheet-tabs", showSheetTabs ? "シート見出しを非表示にします" : "シート見出しを表示します");
  setRibbonCommandTitle("view-zoom-selection", "選択範囲が画面に収まるように倍率を変更します");
  setRibbonCommandTitle("view-freeze-menu", sheet?.freezePanes ? "固定状態の確認と解除" : "ウィンドウ枠の固定メニュー");
  setRibbonCommandTitle("view-split", sheet?.splitPane ? "分割を解除します" : "ワークシートを分割して別々にスクロールします");
  setRibbonCommandTitle("view-new-window", `現在のブックに新しいウィンドウを開きます（表示中: ${visibleWindows}）`);
  setRibbonCommandTitle("view-arrange-all", `表示中のウィンドウを整列します（${visibleWindows}）`);
  setRibbonCommandTitle("view-hide-window", visibleWindows <= 1 ? "表示中の最後のウィンドウは非表示にできません" : "現在のウィンドウを非表示にします");
  setRibbonCommandTitle("view-unhide-window", hiddenWindows ? `非表示のウィンドウを再表示します（${hiddenWindows}）` : "再表示できるウィンドウはありません");
  setRibbonCommandTitle("view-side-by-side", sideBySide ? "並べて比較を解除します" : visibleWindows < 2 ? "比較するには別のウィンドウが必要です" : "比較するウィンドウを選んで左右に並べます");
  setRibbonCommandTitle("view-sync-scrolling", sideBySide ? "並べたウィンドウを同時にスクロールします" : "並べて比較を先に有効にしてください");
  setRibbonCommandTitle("view-reset-window-position", sideBySide ? "並べて比較するウィンドウ位置をリセットします" : "並べて比較を先に有効にしてください");
  setRibbonCommandTitle("view-switch-windows", `表示中のウィンドウを切り替えます（${visibleWindows}）`);
  setRibbonCommandTitle("view-macros", state.model?.macroRecording ? "マクロの記録中です。クリックするとマクロ ダイアログを開きます" : "マクロ ダイアログを開きます");
  setRibbonCommandTitle("view-macros-menu", state.model?.macroRecording ? "マクロの記録を停止できます" : "マクロ メニューを開きます");
  document.querySelector('[data-ribbon-command="formula-show-formulas"]')?.classList.toggle("is-active", state.showFormulas);
  setRibbonCommandTitle("formula-calculation-options", `現在の計算方法: ${calculationModeLabel(workbookCalculationMode())}`);
  const circularReferenceCount = circularFormulaReferencesForSheet(sheet).length;
  setRibbonCommandTitle("formula-error-checking", `エラー チェック / 循環参照${circularReferenceCount ? ` (${circularReferenceCount})` : ""}`);
}



function setRibbonCommandActive(command, active) {
  const button = document.querySelector(`[data-ribbon-command="${command}"]`);
  if (!button) return;
  button.classList.toggle("is-active", Boolean(active));
  button.setAttribute("aria-pressed", active ? "true" : "false");
}



function setRibbonCommandCheckbox(command, checked) {
  const button = document.querySelector(`[data-ribbon-command="${command}"]`);
  if (!button) return;
  button.classList.toggle("is-active", Boolean(checked));
  button.setAttribute("role", "checkbox");
  button.setAttribute("aria-checked", checked ? "true" : "false");
  button.removeAttribute("aria-pressed");
}



function setRibbonCommandTitle(command, title) {
  const button = document.querySelector(`[data-ribbon-command="${command}"]`);
  if (!button) return;
  button.setAttribute("title", title || "");
}



function setRibbonCommandLabel(command, label) {
  const button = document.querySelector(`[data-ribbon-command="${command}"]`);
  const labelElement = button?.querySelector("span:not(.ribbon-check-box):not(.ribbon-command-caret)");
  if (!labelElement || labelElement.textContent === label) return;
  labelElement.textContent = label || "";
}



function setRibbonCommandIcon(command, iconName) {
  const button = document.querySelector(`[data-ribbon-command="${command}"]`);
  const iconElement = button?.querySelector("svg[data-lucide], i[data-lucide]");
  if (!iconElement || !iconName || iconElement.getAttribute("data-lucide") === iconName) return;
  const replacement = document.createElement("i");
  replacement.setAttribute("data-lucide", iconName);
  iconElement.replaceWith(replacement);
  if (window.lucide) window.lucide.createIcons({ icons });
}



function freezePanesMenuKind(sheet) {
  const pane = normalizePaneState(sheet?.freezePanes);
  if (!pane) return "";
  const rows = Number.isFinite(Number(pane.rows)) ? Math.max(0, Math.round(Number(pane.rows))) : Math.max(0, Math.round(Number(pane.row || 1) - 1));
  const columns = Number.isFinite(Number(pane.columns)) ? Math.max(0, Math.round(Number(pane.columns))) : Math.max(0, Math.round(Number(pane.col || 1) - 1));
  if (rows === 1 && columns === 0) return "top-row";
  if (rows === 0 && columns === 1) return "first-column";
  return "selection";
}



function setFreezePanes(kind = "selection") {
  const sheet = activeSheet();
  const groupedSheetIndexes = activeSheetGroupEditIndexes();
  if (groupedSheetIndexes.length > 1) {
    const workbookViewBefore = captureWorkbookViewState();
    const sheetsBefore = captureSheetsForHistory(groupedSheetIndexes);
    if (kind === "unfreeze" || (kind === "selection" && sheet.freezePanes)) {
      groupedSheetIndexes.forEach((sheetIndex) => {
        state.model.sheets[sheetIndex].freezePanes = null;
      });
      commitGroupedSheetViewOptionChange(
        groupedSheetIndexes,
        sheetsBefore,
        `${groupedSheetIndexes.length} 個のシート：ウィンドウ枠の固定を解除しました。`,
        `${groupedSheetIndexes.length} シート ウィンドウ枠の固定解除`,
        workbookViewBefore,
      );
      return;
    }
    let point = selectionToolbarPoint() || { row: 1, col: 1 };
    if (kind === "top-row") point = { row: 2, col: 1 };
    if (kind === "first-column") point = { row: 1, col: 2 };
    const pane = freezePaneForPoint(point);
    if (!pane) {
      setStatus("固定する行または列を選択してください。");
      return;
    }
    groupedSheetIndexes.forEach((sheetIndex) => {
      state.model.sheets[sheetIndex].freezePanes = { ...pane };
    });
    commitGroupedSheetViewOptionChange(
      groupedSheetIndexes,
      sheetsBefore,
      `${groupedSheetIndexes.length} 個のシート：ウィンドウ枠を固定しました。`,
      `${groupedSheetIndexes.length} シート ウィンドウ枠の固定`,
      workbookViewBefore,
    );
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  if (kind === "unfreeze" || (kind === "selection" && sheet.freezePanes)) {
    sheet.freezePanes = null;
    commitSheetViewOptionChange(sheet, sheetBefore, "ウィンドウ枠の固定を解除しました。", `${sheet.name} ウィンドウ枠の固定解除`, workbookViewBefore);
    return;
  }
  let point = selectionToolbarPoint() || { row: 1, col: 1 };
  if (kind === "top-row") point = { row: 2, col: 1 };
  if (kind === "first-column") point = { row: 1, col: 2 };
  const pane = freezePaneForPoint(point);
  if (!pane) {
    setStatus("固定する行または列を選択してください。");
    return;
  }
  sheet.freezePanes = pane;
  commitSheetViewOptionChange(sheet, sheetBefore, "ウィンドウ枠を固定しました。", `${sheet.name} ウィンドウ枠の固定`, workbookViewBefore);
}



function freezePaneForPoint(point) {
  const row = Math.max(1, Math.round(Number(point?.row) || 1));
  const col = Math.max(1, Math.round(Number(point?.col) || 1));
  return normalizePaneState({
    rows: Math.max(0, row - 1),
    columns: Math.max(0, col - 1),
  });
}



function toggleFreezePanes() {
  setFreezePanes("selection");
}



function toggleSplitPane() {
  const sheet = activeSheet();
  const groupedSheetIndexes = activeSheetGroupEditIndexes();
  if (groupedSheetIndexes.length > 1) {
    const workbookViewBefore = captureWorkbookViewState();
    const sheetsBefore = captureSheetsForHistory(groupedSheetIndexes);
    if (sheet.splitPane) {
      groupedSheetIndexes.forEach((sheetIndex) => {
        state.model.sheets[sheetIndex].splitPane = null;
      });
      commitGroupedSheetViewOptionChange(
        groupedSheetIndexes,
        sheetsBefore,
        `${groupedSheetIndexes.length} 個のシート：分割を解除しました。`,
        `${groupedSheetIndexes.length} シート 分割解除`,
        workbookViewBefore,
      );
      return;
    }
    const splitPane = splitPaneForSelection(sheet);
    groupedSheetIndexes.forEach((sheetIndex) => {
      state.model.sheets[sheetIndex].splitPane = { ...splitPane };
    });
    commitGroupedSheetViewOptionChange(
      groupedSheetIndexes,
      sheetsBefore,
      `${groupedSheetIndexes.length} 個のシート：ウィンドウを分割しました。`,
      `${groupedSheetIndexes.length} シート 分割`,
      workbookViewBefore,
    );
    return;
  }
  const workbookViewBefore = captureWorkbookViewState();
  const sheetBefore = captureSheetForHistory(sheet);
  if (sheet.splitPane) {
    sheet.splitPane = null;
    commitSheetViewOptionChange(sheet, sheetBefore, "分割を解除しました。", `${sheet.name} 分割解除`, workbookViewBefore);
    return;
  }
  sheet.splitPane = splitPaneForSelection(sheet);
  commitSheetViewOptionChange(sheet, sheetBefore, "ウィンドウを分割しました。", `${sheet.name} 分割`, workbookViewBefore);
}



function splitPaneForSelection(sheet = activeSheet()) {
  const point = selectionToolbarPoint() || { row: 1, col: 1 };
  let rows = Math.max(0, Math.round(Number(point.row) || 1) - 1);
  let columns = Math.max(0, Math.round(Number(point.col) || 1) - 1);
  if (!rows && !columns) {
    const host = $sheetHost[0];
    const visibleRows = host ? estimatedVisibleRowCount(sheet, host.clientHeight) : 12;
    const visibleColumns = host ? estimatedVisibleColumnCount(sheet, host.clientWidth) : 6;
    rows = Math.max(1, Math.floor(visibleRows / 2));
    columns = Math.max(1, Math.floor(visibleColumns / 2));
  }
  return normalizePaneState({ rows, columns }) || { rows: 1, columns: 1 };
}



function estimatedVisibleRowCount(sheet, viewportHeight) {
  const header = COLUMN_HEADER_HEIGHT;
  const rowHeight = Math.max(1, Number(sheet?.defaultRowHeight) || rowHeightToPixels(15));
  return Math.max(1, Math.floor((Math.max(1, Number(viewportHeight) || 1) - header) / rowHeight));
}



function estimatedVisibleColumnCount(sheet, viewportWidth) {
  const header = rowHeaderWidthForSheet(sheet);
  const colWidth = Math.max(1, Number(sheet?.defaultColumnWidth) || columnWidthToPixels(8.43));
  return Math.max(1, Math.floor((Math.max(1, Number(viewportWidth) || 1) - header) / colWidth));
}



function splitPaneCounts(sheet) {
  const split = normalizePaneState(sheet?.splitPane);
  return {
    rows: Number.isFinite(Number(split?.rows)) ? Math.max(0, Math.round(Number(split.rows))) : Math.max(0, Math.round(Number(split?.row || 1) - 1)),
    columns: Number.isFinite(Number(split?.columns)) ? Math.max(0, Math.round(Number(split.columns))) : Math.max(0, Math.round(Number(split?.col || 1) - 1)),
  };
}



function setSplitPaneCounts(sheet, rows, columns) {
  const nextRows = Math.max(0, Math.round(Number(rows) || 0));
  const nextColumns = Math.max(0, Math.round(Number(columns) || 0));
  if (!nextRows && !nextColumns) {
    sheet.splitPane = null;
    return;
  }
  sheet.splitPane = {
    rows: nextRows,
    columns: nextColumns,
    row: nextRows + 1,
    col: nextColumns + 1,
    activePane: excelSplitActivePane(nextRows, nextColumns) || "topLeft",
  };
}



function splitPaneIndexFromPixel(sheet, axis, pixel) {
  const isColumn = axis === "col";
  const headerSize = isColumn ? rowHeaderWidthForSheet(sheet) : COLUMN_HEADER_HEIGHT;
  const items = isColumn ? sheet.columns : sheet.rows;
  const defaultSize = isColumn ? sheet.defaultColumnWidth : sheet.defaultRowHeight;
  const fallbackSize = defaultSize || (isColumn ? columnWidthToPixels(8.43) : rowHeightToPixels(15));
  let position = headerSize;
  const localPixel = Math.max(headerSize, Number(pixel) || headerSize);
  for (let index = 0; index < items.length; index += 1) {
    const item = items[index];
    const size = item?.hidden ? 0 : item?.width || item?.height || fallbackSize;
    if (size <= 0) continue;
    const next = position + size;
    if (localPixel < (position + next) / 2) return Math.max(0, index);
    if (localPixel <= next) return index + 1;
    position = next;
  }
  return isColumn ? sheet.colCount : sheet.rowCount;
}



function startSplitPaneDrag(event) {
  if (event.button !== 0) return;
  if (state.pendingInkTool) {
    startInkDrawing(event);
    return;
  }
  if (state.pendingShapeType) {
    startShapeDrawing(event);
    return;
  }
  const guide = event.target?.closest?.(".sheet-view-guide.is-split-pane") || event.currentTarget;
  const axis = guide.classList.contains("is-vertical") ? "columns" : "rows";
  const sheet = activeSheet();
  state.splitPaneDrag = {
    axis,
    workbookViewBefore: captureWorkbookViewState(),
    sheetBefore: captureSheetForHistory(sheet),
    startPane: cloneJsonLikeValue(sheet.splitPane || null),
  };
  document.body.classList.add("is-dragging-split-pane");
  guide.setPointerCapture?.(event.pointerId);
  event.preventDefault();
  event.stopPropagation();
}



function updateSplitPaneDrag(event) {
  if (!state.splitPaneDrag || !state.model) return;
  const sheet = activeSheet();
  const pixel = sheetPixelFromClientPosition(event.clientX, event.clientY, { clampToGrid: true });
  if (!pixel) return;
  const counts = splitPaneCounts(sheet);
  if (state.splitPaneDrag.axis === "columns") {
    counts.columns = splitPaneIndexFromPixel(sheet, "col", pixel.x);
  } else {
    counts.rows = splitPaneIndexFromPixel(sheet, "row", pixel.y);
  }
  setSplitPaneCounts(sheet, counts.rows, counts.columns);
  renderSheet();
  syncViewRibbonButtons();
  event.preventDefault();
}



function finishSplitPaneDrag() {
  if (!state.splitPaneDrag) return;
  const drag = state.splitPaneDrag;
  state.splitPaneDrag = null;
  document.body.classList.remove("is-dragging-split-pane");
  const sheet = activeSheet();
  const before = JSON.stringify(drag.startPane || null);
  const after = JSON.stringify(sheet.splitPane || null);
  if (before === after) return;
  refreshActiveSheetViewSnapshot(sheet);
  saveActiveWindowViewState();
  pushHistory({
    sheetIndex: state.activeSheetIndex,
    sheetBefore: drag.sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    workbookViewBefore: drag.workbookViewBefore,
    workbookViewAfter: captureWorkbookViewState(),
    selectionBefore: activeSelectionRange(),
    selectionAfter: activeSelectionRange(),
    selectionBeforeRanges: activeSelectionRanges(),
    selectionAfterRanges: activeSelectionRanges(),
    label: `${sheet.name} 分割位置変更`,
  });
  scheduleWorkbookPersistence();
  setStatus("分割位置を変更しました。");
}



function ensureWorkbookWindows() {
  if (!state.model) return [];
  state.model.windowViews = Array.isArray(state.model.windowViews) ? state.model.windowViews : [];
  if (!state.model.activeWindowId) state.model.activeWindowId = "main";
  if (!state.model.mainWindowView) {
    state.model.mainWindowView = normalizeWorkbookWindowView({
      id: "main",
      name: state.model.sourceName || "Book1",
      ...currentWorkbookWindowSnapshot(),
      hidden: state.model.mainWindowHidden === true,
    });
  }
  return workbookWindowRecords();
}



function workbookWindowBaseName() {
  const name = String(state.model?.sourceName || "").trim();
  return name || "Book1";
}



function workbookWindowSuffixFromName(name) {
  const match = String(name || "").match(/:(\d+)$/);
  if (!match) return null;
  const value = Number(match[1]);
  return Number.isFinite(value) && value > 0 ? value : null;
}



function workbookWindowDisplayNameForIndex(view, index, count) {
  const baseName = workbookWindowBaseName();
  if (count <= 1) return view?.name || baseName;
  if (view?.id === "main") return `${baseName}:1`;
  const suffix = workbookWindowSuffixFromName(view?.name);
  return `${baseName}:${suffix || index + 1}`;
}



function withWorkbookWindowDisplayNames(records) {
  const count = records.length;
  return records.map((view, index) => ({
    ...view,
    displayName: workbookWindowDisplayNameForIndex(view, index, count),
  }));
}



function workbookWindowDisplayName(view) {
  return view?.displayName || view?.name || workbookWindowBaseName();
}



function activeWorkbookWindowDisplayName() {
  if (!state.model) return APP_TITLE;
  const records = workbookWindowRecords();
  const activeId = state.model.activeWindowId || "main";
  const active = records.find((view) => view.id === activeId && !view.hidden) || records.find((view) => !view.hidden);
  return records.length > 1 ? workbookWindowDisplayName(active) : workbookWindowBaseName();
}



function syncWorkbookTitle() {
  if (!state.model) {
    document.title = APP_TITLE;
    $workbookName.text(APP_TITLE);
    renderWorkbookWindowBar();
    return;
  }
  const name = workbookWindowBaseName();
  document.title = name;
  $workbookName.text(name);
}



function updateWorkbookNameDisplay() {
  if (!$workbookNameDisplay.length || !$workbookNameDisplayText.length) return;
  const fullName = String($workbookName.text() || state.model?.sourceName || APP_TITLE).trim() || APP_TITLE;
  const displayNode = $workbookNameDisplay[0];
  const textNode = $workbookNameDisplayText[0];
  displayNode.style.left = "";
  displayNode.style.width = "";
  displayNode.style.maxWidth = "";
  textNode.textContent = fullName;
  const style = window.getComputedStyle(displayNode);
  const padding = (Number.parseFloat(style.paddingLeft) || 0) + (Number.parseFloat(style.paddingRight) || 0);
  const layout = workbookNameDisplayLayout(displayNode);
  const availableWidth = layout.availableWidth;
  if (Number.isFinite(layout.left)) displayNode.style.left = `${layout.left}px`;
  if (availableWidth > 0) displayNode.style.maxWidth = `${availableWidth}px`;
  const desiredWidth = Math.ceil(measureWorkbookNameText(fullName, style.font) + padding + 2);
  const measuredWidth = Math.max(0, Math.min(desiredWidth, availableWidth || desiredWidth));
  if (measuredWidth > 0) displayNode.style.width = `${measuredWidth}px`;
  $workbookNameDisplay.attr("title", fullName);
  const naturalAvailableWidth = Math.max(0, displayNode.clientWidth - padding);
  if (measureWorkbookNameText(fullName, style.font) <= naturalAvailableWidth) {
    textNode.textContent = fullName;
    return;
  }
  let safety = 8;
  let displayName = fullName;
  do {
    displayName = compactWorkbookNameForWidth(fullName, Math.max(0, naturalAvailableWidth - safety), style.font);
    textNode.textContent = displayName;
    safety += 8;
  } while (textNode.scrollWidth > textNode.clientWidth && safety <= 80);
}



function workbookNameDisplayLayout(displayNode) {
  const row = displayNode.closest?.(".ribbon-tab-row");
  if (!row) {
    return {
      availableWidth: Math.max(120, Math.floor(window.innerWidth * 0.72)),
      left: null,
    };
  }
  const rowRect = row.getBoundingClientRect();
  const edgeGap = 8;
  let leftLimit = rowRect.left + edgeGap;
  let rightLimit = rowRect.right - edgeGap;
  const tabs = row.querySelector(".ribbon-tabs");
  const tabsRect = tabs?.getBoundingClientRect?.();
  if (tabsRect && tabsRect.width > 0 && tabsRect.right > leftLimit) {
    leftLimit = tabsRect.right + edgeGap;
  }
  [row.querySelector(".ribbon-status"), row.querySelector(".ribbon-display-options")]
    .filter(Boolean)
    .forEach((element) => {
      const elementStyle = window.getComputedStyle(element);
      const elementRect = element.getBoundingClientRect();
      if (elementStyle.display !== "none" && elementRect.width > 0 && elementRect.left < rightLimit) {
        rightLimit = elementRect.left - edgeGap;
      }
    });
  const availableWidth = Math.floor(rightLimit - leftLimit);
  if (availableWidth > 0) {
    return {
      availableWidth,
      left: (leftLimit + rightLimit) / 2 - rowRect.left,
    };
  }
  return {
    availableWidth: Math.max(96, Math.floor(rowRect.width * 0.36)),
    left: rowRect.width / 2,
  };
}



function compactWorkbookNameForWidth(name, availableWidth, font) {
  const fullName = String(name || "").trim() || APP_TITLE;
  if (!Number.isFinite(availableWidth) || availableWidth <= 0) return fullName;
  if (measureWorkbookNameText(fullName, font) <= availableWidth) return fullName;

  const dotIndex = fullName.lastIndexOf(".");
  const hasExtension = dotIndex > 0 && dotIndex < fullName.length - 1;
  const extension = hasExtension ? fullName.slice(dotIndex) : "";
  const stem = hasExtension ? fullName.slice(0, dotIndex) : fullName;
  const ellipsis = "...";
  let best = extension ? `${ellipsis}${extension}` : ellipsis;
  let low = 0;
  let high = stem.length;

  while (low <= high) {
    const totalStemChars = Math.floor((low + high) / 2);
    const suffixLen = Math.min(totalStemChars, Math.max(Math.min(4, totalStemChars), Math.floor(totalStemChars * 0.34)));
    const prefixLen = Math.max(0, totalStemChars - suffixLen);
    const candidate = `${stem.slice(0, prefixLen)}${ellipsis}${suffixLen ? stem.slice(stem.length - suffixLen) : ""}${extension}`;
    if (measureWorkbookNameText(candidate, font) <= availableWidth) {
      best = candidate;
      low = totalStemChars + 1;
    } else {
      high = totalStemChars - 1;
    }
  }

  return best;
}



function renderWorkbookWindowBar() {
  if (!$workbookWindowBar.length) return;
  if (!state.model) {
    $workbookWindowBar.addClass("hidden").empty();
    return;
  }
  const records = workbookWindowRecords();
  const visible = records.filter((view) => !view.hidden);
  if (visible.length <= 1) {
    $workbookWindowBar.addClass("hidden").empty();
    return;
  }
  const activeId = state.model.activeWindowId || "main";
  const activeRecord = visible.find((view) => view.id === activeId) || visible[0];
  const sideBySide = normalizedViewSideBySidePair(visible);
  const arrangement = workbookWindowArrangementForDisplay(visible);
  const hiddenCount = records.length - visible.length;
  const synchronous = Boolean(sideBySide && state.model.synchronousScrolling);
  const chips = visible.map((view, index) => workbookWindowChipHtml(view, index, activeId, sideBySide)).join("");
  const preview = arrangement ? workbookWindowArrangementPreviewHtml(visible, arrangement, activeId, sideBySide) : "";
  const modeLabel = workbookWindowArrangementModeLabel(arrangement?.mode || "tiled");
  const activeName = workbookWindowDisplayName(activeRecord);
  $workbookWindowBar
    .removeClass("hidden")
    .attr("data-window-count", String(visible.length))
    .attr("data-hidden-count", String(hiddenCount))
    .attr("data-window-arrangement-mode", arrangement?.mode || "")
    .attr("data-side-by-side-left", sideBySide?.left || "")
    .attr("data-side-by-side-right", sideBySide?.right || "")
    .attr("data-sync-scrolling", synchronous ? "true" : "false")
    .html(`
      <div class="workbook-window-summary">
        <span class="workbook-window-summary-title">ウィンドウ</span>
        <strong>${escapeHtml(activeName)}</strong>
        <span>${escapeHtml(modeLabel)}</span>
        ${sideBySide ? `<span class="workbook-window-badge">比較</span>` : ""}
        ${sideBySide ? `<span class="workbook-window-badge${synchronous ? " is-on" : ""}">同期スクロール：${synchronous ? "オン" : "オフ"}</span>` : ""}
        ${hiddenCount ? `<span class="workbook-window-badge">${hiddenCount} 個非表示</span>` : ""}
      </div>
      <div class="workbook-window-tabs" role="tablist" aria-label="ワークブック ウィンドウの切り替え">${chips}</div>
      ${preview}
    `);
}



function workbookWindowArrangementForDisplay(visibleRecords = []) {
  const current = state.model?.windowArrangement;
  const visibleIds = new Set(visibleRecords.map((view) => view.id));
  const sideBySide = normalizedViewSideBySidePair(visibleRecords);
  const sideBySideIds = new Set(sideBySide ? [sideBySide.left, sideBySide.right] : []);
  const expectedWindowCount = current?.mode === "side-by-side" && sideBySide ? 2 : visibleRecords.length;
  if (
    current &&
    typeof current === "object" &&
    Array.isArray(current.windows) &&
    current.windows.length === expectedWindowCount &&
    current.windows.every((slot) => visibleIds.has(slot.id)) &&
    (current.mode !== "side-by-side" || !sideBySide || current.windows.every((slot) => sideBySideIds.has(slot.id)))
  ) {
    return current;
  }
  const fallbackMode = sideBySide ? "side-by-side" : "tiled";
  return computeWorkbookWindowArrangement(fallbackMode);
}

export {
  activateSheetView,
  activeSheetView,
  activeWorkbookWindowDisplayName,
  addFormulaWatchesFromSelection,
  applyCreateNamesFromSelection,
  applyGroupedSheetViewOptionChange,
  applyNameManagerSave,
  applySheetViewNameDialog,
  applySheetViewSnapshot,
  calculateActiveSheetNow,
  calculateWorkbookNow,
  calculationModeLabel,
  calculationModeToCalcMode,
  captureSheetViewSnapshot,
  captureWorkbookMacroState,
  captureWorkbookViewState,
  circularFormulaReferencesForSheet,
  closeFormulaWatchWindow,
  commitGroupedSheetViewOptionChange,
  commitSheetViewChange,
  commitSheetViewOptionChange,
  commitWorkbookMacroChange,
  commitWorkbookViewChange,
  compactWorkbookNameForWidth,
  createSheetView,
  currentSheetViewStateFields,
  definedNameForFormulaWatch,
  deleteActiveSheetView,
  deleteDefinedNameFromManager,
  deleteSelectedFormulaWatches,
  displayForFormulaReference,
  duplicateActiveSheetView,
  ensureWorkbookWindows,
  escapeRegex,
  estimatedVisibleColumnCount,
  estimatedVisibleRowCount,
  evaluateFormulaTraceExpression,
  exitActiveSheetView,
  exitFullScreenView,
  finishSplitPaneDrag,
  formulaAuditClassForCell,
  formulaAuditFormulaSourceRanges,
  formulaAuditMenuButtonHtml,
  formulaCalculationMenuButtonHtml,
  formulaDependencyPathReturnsTo,
  formulaErrorIgnoreKey,
  formulaErrorMessage,
  formulaErrorsForSheet,
  formulaEvaluateReferenceStepInTrace,
  formulaEvaluateStepInSubexpression,
  formulaEvaluateStepInTrace,
  formulaEvaluationReferenceLabel,
  formulaEvaluationReferencesForExpression,
  formulaEvaluationTraceForCell,
  formulaEvaluationTraceForExpression,
  formulaExpressionWithPrefix,
  formulaFunctionCallCoversWholeExpression,
  formulaFunctionCallsInExpression,
  formulaFunctionNameBeforeOpen,
  formulaReferenceDisplayRegexSource,
  formulaWatchEntries,
  formulaWatchId,
  formulaWatchRowHtml,
  formulaWatchRowModel,
  freezePaneForPoint,
  freezePanesMenuKind,
  handleFormulaAuditMenuAction,
  handleFormulaErrorDialogAction,
  handleFormulaEvaluateDialogAction,
  handleFormulaWatchRowDoubleClick,
  handleFormulaWatchWindowAction,
  handleFullScreenEscape,
  handleFullscreenChange,
  handleNameManagerRowDoubleClick,
  keepActiveSheetView,
  macroSecurityLevelLabel,
  matchingFormulaParenIndex,
  normalizeAuditRange,
  normalizeFormulaAuditArrow,
  normalizeFormulaAudits,
  normalizeFormulaExpressionForCompare,
  normalizeFormulaWatches,
  openCalculationOptionsMenu,
  openCreateNamesDialog,
  openEvaluateFormulaDialog,
  openFormulaErrorCheckingDialog,
  openFormulaErrorCheckingMenu,
  openFormulaRemoveArrowsMenu,
  openFormulaWatchWindow,
  openNameEditorFromManager,
  openSheetViewOptionsMenu,
  openSheetViewRenameDialog,
  openViewMacrosMenu,
  queueSheetViewScrollRestore,
  refreshActiveSheetViewSnapshot,
  refreshSheetViewSnapshotForSheet,
  removeFormulaAuditArrows,
  removeFormulaAuditArrowsByKind,
  renderWorkbookWindowBar,
  replaceFormulaReferenceForDisplay,
  resetIgnoredFormulaErrors,
  restoreActiveWorkbookWindowViewFromModel,
  restorePendingSheetViewScroll,
  restoreWorkbookMacroState,
  restoreWorkbookViewState,
  runFormulaErrorChecking,
  selectCircularFormulaReference,
  selectFormulaError,
  selectFormulaWatchCell,
  setCalculationMode,
  setFreezePanes,
  setRibbonCommandActive,
  setRibbonCommandCheckbox,
  setRibbonCommandIcon,
  setRibbonCommandLabel,
  setRibbonCommandTitle,
  setSplitPaneCounts,
  sheetHeadingsVisible,
  sheetViewDisplayName,
  sheetViewMenuDescription,
  splitPaneCounts,
  splitPaneForSelection,
  splitPaneIndexFromPixel,
  startSplitPaneDrag,
  syncBrowserFullscreenToModelState,
  syncEvaluateFormulaDialog,
  syncFormulaErrorCheckingDialog,
  syncFormulaWatchWindow,
  syncFullScreenViewClass,
  syncViewRibbonButtons,
  syncWorkbookChromeVisibility,
  syncWorkbookTitle,
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
  traceFormulaErrors,
  traceFormulaReferences,
  uniqueFormulaAuditArrows,
  uniqueSheetViewCopyName,
  uniqueSheetViewName,
  updateActiveSheetView,
  updateSplitPaneDrag,
  updateWorkbookNameDisplay,
  withWorkbookWindowDisplayNames,
  workbookCalculationMode,
  workbookCalculationModeFromCalcProperties,
  workbookNameDisplayLayout,
  workbookWindowArrangementForDisplay,
  workbookWindowBaseName,
  workbookWindowDisplayName,
  workbookWindowDisplayNameForIndex,
  workbookWindowSuffixFromName,
};
