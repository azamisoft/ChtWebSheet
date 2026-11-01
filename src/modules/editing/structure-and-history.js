// Feature module; shared bindings are imported explicitly.
import {
  $horizontalScrollbar,
  $sheetHost,
  CELL_BORDER_SIDES,
  DEFAULT_CELL_CSS_KEYS,
  EMPTY_CELL_BORDERS,
  EMPTY_CELL_CSS,
  HISTORY_LIMIT,
  INFERRED_DEFAULT_CELL_CSS_KEYS,
  LOCALIZED_DEFAULT_CELL_CSS_KEYS,
  STYLE_INTERN_POOL_LIMIT,
  cellStyleInternPools,
  sheetStyleRangeCacheBySheet,
  state,
} from "../../app-core.js";
import {
  cwsAgentRestoreSelectionSnapshot,
} from "../agent/contracts.js";
import {
  cloneScenarios,
  openScenarioManagerDialog,
} from "../data/analysis-and-scenarios.js";
import {
  refreshQueriesConnectionsPane,
  restoreWorkbookDataModel,
} from "../data/connections-and-queries.js";
import {
  openWorkbookLinksDialog,
} from "../data/query-transforms.js";
import {
  scheduleQueryRefreshTimers,
} from "../data/sort-and-autofilter.js";
import {
  cloneRelationships,
  openRelationshipsDialog,
} from "../data/validation-and-consolidation.js";
import {
  rangesOverlap1d,
} from "../drawing/charts-and-insertion.js";
import {
  normalizedHeaderFooterSettings,
} from "../drawing/shape-rendering.js";
import {
  applyWorkbookReferenceShiftForStructureChange,
  shiftFormulaReferences,
  structureReferenceShiftForOperation,
} from "./autofill.js";
import {
  normalizeColorValue,
} from "./ink-and-clipboard.js";
import {
  clearCutSourceAfterInternalPaste,
  cutPasteHistorySheetIndexes,
  pasteTargetSizeForClipboard,
  repeatedClipboardSourceOffset,
} from "./keyboard-and-paste.js";
import {
  activeSelectionRange,
  activeSelectionRanges,
  activeSheetGroupEditIndexes,
  affectedSheetIndexesForGroupedRanges,
  coerceUserInput,
  expandRangeForMerges,
  getDisplayForCell,
  invalidateFormulaSpillAnchorCache,
  invalidateSheetUsedRangeCache,
  materializeSelectionRange,
  materializeSelectionRangeForSheet,
  materializedRangesForSheet,
  normalizeRange,
  rangeToLabel,
  rangesToLabel,
  refreshChartsForGroupedRanges,
  restoreGroupedSheetsAfterFailure,
  updateCellModel,
  updateSelectionUi,
  writeCellContent,
} from "./pointers-and-formula-evaluation.js";
import {
  cellHasFontColor,
  updateFormulaBarForCell,
  updateFormulaBarForSelection,
} from "./selection-and-cell-text.js";
import {
  ensureCellModel,
  finishSheetStructureHistoryRestore,
  forEachCellInRange,
  removeIntersectingMerges,
  renderSheet,
  restoreCellsFromHistory,
  restoreDimensionsFromHistory,
} from "../rendering/grid-and-canvas.js";
import {
  clearCellsInRange,
  clipboardCellEntries,
  clipboardHasFormula,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  columnName,
  columnWidthToPixels,
  dateToExcelSerial,
  rowHeightToPixels,
  setStatus,
  showMessageBox,
} from "../ui/help-and-utilities.js";
import {
  closeGoToDialog,
  closeSpecialDialog,
} from "../ui/printing-and-home-tools.js";
import {
  ensureSheetSize,
  isPersistentCell,
  parseCellKey,
  quoteSheetNameForFormula,
} from "../ui/ribbon-and-links.js";
import {
  cloneSheetModel,
  createBlankSheet,
  ensureEditableRangeForProtection,
  ensureEditableRangesForProtection,
  ensureStructureAllowedForProtection,
  ensureStructureAllowedForSheet,
  moveSheetToIndex,
  restoreZoomFromHistory,
} from "../ui/sheet-protection-and-translation.js";
import {
  restoreWorkbookMacroState,
  restoreWorkbookViewState,
} from "../ui/views-and-formula-tools.js";
import {
  buildFormulaEngine,
  ensureActiveSheetVisible,
  migrateDenseFormulaDataToCells,
  renderWorkbook,
  structuredReferenceRows,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  markWorkbookChanged,
  normalizePaneState,
  normalizeSheetCodeModule,
  normalizeSheetKind,
  normalizeSheetView,
  normalizeSpecialSheetDescriptor,
  normalizedSheetShowRuler,
  normalizedSheetViewMode,
  restoreWorkbookThemeState,
  scheduleWorkbookPersistence,
} from "../workbook/import-model.js";
import {
  activeSheet,
  getSheetId,
  implicitCellDisplayText,
  isOccupiedCell,
  normalizedCellDisplayText,
} from "../workbook/saving-and-export.js";
import {
  applyTableRangeEditBehaviors,
  copyTableCalculatedColumnFormula,
  copyTableCellFormatting,
  ensureTableColumnMetadata,
  normalizeOutlineSettings,
  refreshImportedTableVisualStyleForTableMutation,
  syncTableColumnNamesFromHeaders,
  tableAutoFilterRange,
  tableRangeEditMayChangeMetadata,
} from "../workbook/tables-and-validation.js";
import {
  normalizedPageLayoutSettings,
  openMacroSecurityDialog,
} from "../workbook/vba-and-printing.js";
import {
  openCustomViewsDialog,
  openMacrosDialog,
  restoreWorkbookWindowScrollStable,
} from "../workbook/views-and-macros.js";


function normalizePasteSpecialOptions(options = {}) {
  if (typeof options === "string") return normalizePasteSpecialOptions({ mode: options });
  const mode = [
    "all",
    "allUsingSourceTheme",
    "formulas",
    "values",
    "formats",
    "comments",
    "validation",
    "allExceptBorders",
    "columnWidths",
    "formulasNumberFormats",
    "valuesNumberFormats",
  ].includes(options.mode) ? options.mode : "all";
  const operation = ["add", "subtract", "multiply", "divide"].includes(options.operation) ? options.operation : "none";
  return {
    mode,
    operation,
    skipBlanks: Boolean(options.skipBlanks),
    transpose: Boolean(options.transpose),
  };
}



function pasteSpecialModeWritesContent(mode) {
  return ["all", "allUsingSourceTheme", "formulas", "values", "allExceptBorders", "formulasNumberFormats", "valuesNumberFormats"].includes(mode);
}



function pasteSpecialModeWritesFormula(mode) {
  return ["all", "allUsingSourceTheme", "formulas", "allExceptBorders", "formulasNumberFormats"].includes(mode);
}



function pasteSpecialModeLabel(mode) {
  return {
    allUsingSourceTheme: "コピー元のテーマを使用してすべて",
    formulas: "数式",
    values: "値",
    formats: "書式",
    comments: "コメントとメモ",
    validation: "入力規則",
    allExceptBorders: "罫線を除くすべて",
    columnWidths: "列幅",
    formulasNumberFormats: "数式と数値の書式",
    valuesNumberFormats: "値と数値の書式",
  }[mode] || "すべて";
}



function sourceOffsetForPasteSpecial(rowOffset, colOffset, transpose = false) {
  return transpose ? { row: colOffset, col: rowOffset } : { row: rowOffset, col: colOffset };
}



function isBlankClipboardCell(sourceCell) {
  if (!sourceCell) return true;
  const values = [sourceCell.raw, sourceCell.formula, sourceCell.cached, sourceCell.display];
  return values.every((value) => value == null || value === "");
}



function pasteClipboardCellAllForSpecial(sheet, sheetId, row, col, clipboard, sourceCell, sourceOffset, special) {
  if (special.operation === "none") {
    writeClonedClipboardCellContent(sheet, row, col, clipboard, sourceCell, sourceOffset);
    return;
  }
  pasteClipboardCellFormat(sheet, row, col, sourceCell);
  const value = pasteDisplayValueFromClipboardCell(clipboard, sourceCell, sourceOffset.row, sourceOffset.col);
  const computed = pasteSpecialOperatedValue(sheet, row, col, value, special.operation);
  if (computed.ok) writeCellContent(sheet, sheetId, row, col, computed.value);
}



function pasteClipboardCellContentForSpecial(sheet, sheetId, row, col, clipboard, sourceCell, sourceOffset, special) {
  let content = null;
  if (special.mode === "values" || special.mode === "valuesNumberFormats") {
    content = pasteDisplayValueFromClipboardCell(clipboard, sourceCell, sourceOffset.row, sourceOffset.col);
  } else {
    content = pasteValueFromClipboardCell(sourceCell, {
      clipboard,
      targetRow: row,
      targetCol: col,
      sourceOffset,
    });
  }
  const computed = pasteSpecialOperatedValue(sheet, row, col, content, special.operation);
  writeCellContent(sheet, sheetId, row, col, computed.ok ? computed.value : content);
  if (["formulas", "formulasNumberFormats", "valuesNumberFormats"].includes(special.mode) && sourceCell?.numFmt) {
    const targetCell = ensureCellModel(sheet, row, col);
    targetCell.numFmt = sourceCell.numFmt;
    storeCellModelForSheet(sheet, cellKey(row, col), targetCell);
  }
}



function pasteSpecialOperatedValue(sheet, row, col, sourceValue, operation) {
  if (operation === "none") return { ok: false, value: sourceValue };
  const sourceNumber = pasteSpecialNumericValue(sourceValue);
  if (!Number.isFinite(sourceNumber)) return { ok: false, value: sourceValue };
  const targetCell = sheet.cells[cellKey(row, col)];
  const targetNumber = pasteSpecialNumericValue(targetCell?.raw ?? targetCell?.cached ?? targetCell?.display ?? 0);
  const base = Number.isFinite(targetNumber) ? targetNumber : 0;
  if (operation === "add") return { ok: true, value: base + sourceNumber };
  if (operation === "subtract") return { ok: true, value: base - sourceNumber };
  if (operation === "multiply") return { ok: true, value: base * sourceNumber };
  if (operation === "divide") {
    if (sourceNumber === 0) return { ok: true, value: "#DIV/0!" };
    return { ok: true, value: base / sourceNumber };
  }
  return { ok: false, value: sourceValue };
}



function pasteSpecialNumericValue(value) {
  if (typeof value === "number") return value;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return dateToExcelSerial(value);
  const text = String(value ?? "").trim();
  if (!text) return 0;
  if (text.startsWith("=")) return Number.NaN;
  const normalized = text
    .replace(/[,\s￥¥$€£]/g, "")
    .replace(/%$/, "");
  const number = Number(normalized);
  if (!Number.isFinite(number)) return Number.NaN;
  return text.endsWith("%") ? number / 100 : number;
}



function pasteClipboardCellAllExceptBorders(sheet, row, col, sourceCell, clipboard = null, sourceOffset = null) {
  const previousBorders = cloneJsonLikeValue(sheet.cells[cellKey(row, col)]?.borders) || {};
  writeClonedClipboardCellContent(sheet, row, col, clipboard, sourceCell, sourceOffset);
  const target = ensureCellModel(sheet, row, col);
  target.borders = mutableBordersObject(previousBorders);
  storeCellModelForSheet(sheet, cellKey(row, col), target);
}



function pasteClipboardColumnWidths(sheet, clipboard, startCol, targetColCount, transpose = false) {
  const sourceSheet = state.model?.sheets?.[clipboard.sourceSheetIndex];
  let changed = 0;
  const repeatWidth = Math.max(1, transpose ? clipboard.rowCount : clipboard.colCount);
  for (let offset = 1; offset <= targetColCount; offset += 1) {
    const sourceColOffset = transpose ? 1 : ((offset - 1) % repeatWidth) + 1;
    const sourceCol = (clipboard.sourceRange?.left || 1) + sourceColOffset - 1;
    const sourceColumn = clipboard.columns?.[sourceColOffset - 1] || sourceSheet?.columns?.[sourceCol - 1] || null;
    const width = Number(sourceColumn?.width);
    if (!Number.isFinite(width) || width <= 0) continue;
    const targetCol = startCol + offset - 1;
    sheet.columns[targetCol - 1] = { ...defaultColumnModel(sheet), ...(sheet.columns[targetCol - 1] || {}) };
    if (sheet.columns[targetCol - 1].width !== width || sheet.columns[targetCol - 1].hidden) {
      sheet.columns[targetCol - 1].width = width;
      sheet.columns[targetCol - 1].hidden = false;
      changed += 1;
    }
  }
  return changed;
}



function pasteDisplayValueFromClipboardCell(clipboard, sourceCell, rowOffset, colOffset) {
  const sourceSheet = state.model?.sheets?.[clipboard.sourceSheetIndex];
  if (sourceSheet && clipboard.sourceRange) {
    const display = getDisplayForCell(
      sourceSheet,
      getSheetId(sourceSheet.name),
      clipboard.sourceRange.top + rowOffset - 1,
      clipboard.sourceRange.left + colOffset - 1,
    );
    if (display != null && display !== "") return display;
  }
  if (!sourceCell) return null;
  if (sourceCell.formula || String(sourceCell.raw || "").startsWith("=")) {
    if (sourceCell.cached != null && sourceCell.cached !== sourceCell.formula) return sourceCell.cached;
    if (sourceCell.display != null && sourceCell.display !== sourceCell.formula) return sourceCell.display;
    return null;
  }
  return pasteValueFromClipboardCell(sourceCell);
}



function pasteClipboardCellFormat(sheet, row, col, sourceCell) {
  const target = ensureCellModel(sheet, row, col);
  target.css = mutableCssObject(cloneJsonLikeValue(sourceCell?.css) || {});
  target.hasFill = Boolean(sourceCell?.hasFill);
  target.hasFontColor = Boolean(sourceCell?.hasFontColor);
  target.borders = mutableBordersObject(cloneJsonLikeValue(sourceCell?.borders) || {});
  target.numFmt = sourceCell?.numFmt || "";
  storeCellModelForSheet(sheet, cellKey(row, col), target);
}



function pasteClipboardCellComments(sheet, row, col, sourceCell) {
  const target = ensureCellModel(sheet, row, col);
  target.comment = cloneJsonLikeValue(sourceCell?.comment);
  target.note = cloneJsonLikeValue(sourceCell?.note);
  target.noteVisible = Boolean(sourceCell?.noteVisible);
  storeCellModelForSheet(sheet, cellKey(row, col), target);
}



function pasteClipboardCellValidation(sheet, row, col, sourceCell) {
  const target = ensureCellModel(sheet, row, col);
  target.validation = cloneJsonLikeValue(sourceCell?.validation);
  storeCellModelForSheet(sheet, cellKey(row, col), target);
}



function pasteInternalClipboardLinkAtRange(range) {
  const clipboard = state.internalClipboard;
  if (!clipboard?.sourceRange || clipboard.sourceSheetIndex == null) {
    showMessageBox("リンク貼り付けできるコピー元がありません。", { title: "Cht WebSheet" });
    return;
  }
  if (!range) {
    setStatus("貼り付け先のセルを選択してください。");
    return;
  }
  const sourceSheet = state.model?.sheets?.[clipboard.sourceSheetIndex];
  if (!sourceSheet) {
    showMessageBox("コピー元のシートが見つかりません。", { title: "Cht WebSheet" });
    return;
  }
  const sheet = activeSheet();
  const pasteSize = pasteTargetSizeForClipboard(clipboard, range);
  if (!pasteSize.ok) {
    showMessageBox(pasteSize.message, { title: "Cht WebSheet" });
    setStatus(pasteSize.message);
    return;
  }
  const targetRows = range.top + pasteSize.targetRowCount - 1;
  const targetCols = range.left + pasteSize.targetColCount - 1;
  const targetRange = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left },
    { sheetIndex: state.activeSheetIndex, row: targetRows, col: targetCols },
  );
  if (!ensureEditableRangeForProtection(sheet, targetRange, "リンク貼り付け")) return;
  const before = captureCellsForHistory(sheet, targetRange);
  let changed = 0;
  try {
    ensureSheetSize(sheet, targetRows, targetCols);
    const sheetId = getSheetId(sheet.name);
    for (let rowOffset = 1; rowOffset <= pasteSize.targetRowCount; rowOffset += 1) {
      for (let colOffset = 1; colOffset <= pasteSize.targetColCount; colOffset += 1) {
        const sourceOffset = repeatedClipboardSourceOffset(rowOffset, colOffset, pasteSize);
        const sourceRow = clipboard.sourceRange.top + sourceOffset.row - 1;
        const sourceCol = clipboard.sourceRange.left + sourceOffset.col - 1;
        const row = range.top + rowOffset - 1;
        const col = range.left + colOffset - 1;
        writeCellContent(sheet, sheetId, row, col, `=${quoteSheetNameForFormula(sourceSheet.name)}!${columnName(sourceCol)}${sourceRow}`);
        changed += 1;
      }
    }
    state.hf = buildFormulaEngine(state.model);
    state.selected = { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
    state.selectionRange = targetRange;
    state.selectionRanges = [targetRange];
    state.selectionStart = state.selected;
    state.selectionDrag = null;
    const after = captureCellsForHistory(sheet, targetRange);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      before,
      after,
      selectionBefore: range,
      selectionAfter: targetRange,
      label: `${changed} 個のセルをリンク貼り付け`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForCell(range.top, range.left);
    setStatus(`${changed} 個のセルをリンク貼り付けしました。`);
  } catch (error) {
    console.error(error);
    setStatus(`リンク貼り付けに失敗しました：${error.message || error}`);
    renderSheet();
  }
}



function clipboardHasDimensions(clipboard) {
  return Boolean(
    (clipboard?.rows || []).some((row) => Number(row?.height) > 0) ||
      (clipboard?.columns || []).some((column) => Number(column?.width) > 0)
  );
}



function applyClipboardDimensionsToRange(sheet, clipboard, range) {
  (clipboard.columns || []).forEach((sourceColumn, offset) => {
    const width = Number(sourceColumn?.width);
    if (!Number.isFinite(width) || width <= 0) return;
    const targetCol = range.left + offset;
    sheet.columns[targetCol - 1] = { ...defaultColumnModel(sheet), ...(sheet.columns[targetCol - 1] || {}) };
    sheet.columns[targetCol - 1].width = width;
    sheet.columns[targetCol - 1].hidden = false;
  });
  (clipboard.rows || []).forEach((sourceRow, offset) => {
    const height = Number(sourceRow?.height);
    if (!Number.isFinite(height) || height <= 0) return;
    const targetRow = range.top + offset;
    sheet.rows[targetRow - 1] = { ...defaultRowModel(sheet), ...(sheet.rows[targetRow - 1] || {}) };
    sheet.rows[targetRow - 1].height = height;
    sheet.rows[targetRow - 1].hidden = false;
  });
}



function pasteInternalClipboardAtRange(range, options = {}) {
  const clipboard = state.internalClipboard;
  if (!clipboard) {
    setStatus("貼り付けできる内部セル データがありません。");
    return;
  }
  if (options.keepSourceFormatting === false) {
    pasteInternalClipboardValuesAtRange(range, clipboard);
    return;
  }

  const sheet = activeSheet();
  const pasteSize = pasteTargetSizeForClipboard(clipboard, range);
  if (!pasteSize.ok) {
    showMessageBox(pasteSize.message, { title: "Cht WebSheet" });
    setStatus(pasteSize.message);
    return;
  }
  const targetRows = range.top + pasteSize.targetRowCount - 1;
  const targetCols = range.left + pasteSize.targetColCount - 1;
  const targetRange = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left },
    { sheetIndex: state.activeSheetIndex, row: targetRows, col: targetCols },
  );
  if (!ensureEditableRangeForProtection(sheet, targetRange, "貼り付け")) return;
  const entries = clipboardCellEntries(clipboard);
  const hasClipboardContent = entries.length > 0 || String(clipboard.text || "").trim() !== "";
  const tableEditCandidate = tableRangeEditMayChangeMetadata(sheet, targetRange, {
    assumeNonEmpty: hasClipboardContent,
    assumeFormula: clipboardHasFormula(clipboard),
  });
  const cutSheetIndexes = cutPasteHistorySheetIndexes(clipboard, state.activeSheetIndex);
  const appliesDimensions = options.applyDimensions === true && !pasteSize.repeating && clipboardHasDimensions(clipboard);
  const usesSheetHistory = tableEditCandidate || cutSheetIndexes.length || appliesDimensions;
  const sheetsBefore = tableEditCandidate
    ? captureSheetsForHistory(state.model.sheets.map((_item, index) => index))
    : cutSheetIndexes.length
      ? captureSheetsForHistory(cutSheetIndexes)
      : null;
  const sheetBefore = usesSheetHistory && !sheetsBefore?.length ? captureSheetForHistory(sheet) : null;
  const before = captureCellsForHistory(sheet, targetRange);
  const mergesBefore = cloneMerges(sheet.merges);
  let changed = 0;

  try {
    ensureSheetSize(sheet, targetRows, targetCols);
    if (appliesDimensions) applyClipboardDimensionsToRange(sheet, clipboard, range);
    removeIntersectingMerges(sheet, targetRange);
    clearCellsInRange(sheet, targetRange);
    const entriesByOffset = new Map(entries.map((sourceCell) => [cellKey(sourceCell.row, sourceCell.col), sourceCell]));
    for (let rowOffset = 1; rowOffset <= pasteSize.targetRowCount; rowOffset += 1) {
      for (let colOffset = 1; colOffset <= pasteSize.targetColCount; colOffset += 1) {
        const sourceOffset = repeatedClipboardSourceOffset(rowOffset, colOffset, pasteSize);
        const sourceCell = entriesByOffset.get(cellKey(sourceOffset.row, sourceOffset.col));
        const row = range.top + rowOffset - 1;
        const col = range.left + colOffset - 1;
        writeClonedClipboardCellContent(sheet, row, col, clipboard, sourceCell, sourceOffset);
        changed += 1;
      }
    }
    if (!pasteSize.repeating) clipboard.merges?.forEach((merge) => {
      sheet.merges.push({
        top: range.top + merge.top - 1,
        left: range.left + merge.left - 1,
        bottom: range.top + merge.bottom - 1,
        right: range.left + merge.right - 1,
      });
    });
    const tableEdit = applyTableRangeEditBehaviors(sheet, targetRange);
    clearCutSourceAfterInternalPaste(clipboard, targetRange, state.activeSheetIndex);
    state.hf = buildFormulaEngine(state.model);
    state.selected = { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left };
    state.selectionRange = expandRangeForMerges(sheet, targetRange);
    state.selectionRanges = [state.selectionRange];
    state.selectionStart = state.selected;
    state.selectionDrag = null;
    const after = captureCellsForHistory(sheet, targetRange);
    const mergesAfter = cloneMerges(sheet.merges);
    const label = tableEdit.changed ? `${tableEdit.label} / ${changed} 個のセルを貼り付け` : `${changed} 個のセルを貼り付け（書式${clipboard.merges?.length ? "と結合セルを含む" : "を含む"}）`;
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
          mergesBefore,
          mergesAfter,
          selectionBefore: range,
          selectionAfter: state.selectionRange,
          label,
        });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForCell(range.top, range.left);
    setStatus(tableEdit.changed ? `${tableEdit.status} ${changed} 個のセルを貼り付けました。` : `${changed} 個のセルを貼り付けました（書式${clipboard.merges?.length ? "と結合セルを含む" : "を含む"}）。`);
  } catch (error) {
    console.error(error);
    setStatus(`貼り付けに失敗しました：${error.message || error}`);
    renderSheet();
  }
}



function pasteInternalClipboardValuesAtRange(range, clipboard) {
  const sheet = activeSheet();
  const pasteSize = pasteTargetSizeForClipboard(clipboard, range);
  if (!pasteSize.ok) {
    showMessageBox(pasteSize.message, { title: "Cht WebSheet" });
    setStatus(pasteSize.message);
    return;
  }
  const targetRows = range.top + pasteSize.targetRowCount - 1;
  const targetCols = range.left + pasteSize.targetColCount - 1;
  const targetRange = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left },
    { sheetIndex: state.activeSheetIndex, row: targetRows, col: targetCols },
  );
  if (!ensureEditableRangeForProtection(sheet, targetRange, "貼り付け")) return;

  const entriesByOffset = new Map();
  clipboardCellEntries(clipboard).forEach((sourceCell) => {
    entriesByOffset.set(cellKey(sourceCell.row, sourceCell.col), sourceCell);
  });
  const hasClipboardContent = entriesByOffset.size > 0 || String(clipboard.text || "").trim() !== "";
  const tableEditCandidate = tableRangeEditMayChangeMetadata(sheet, targetRange, {
    assumeNonEmpty: hasClipboardContent,
    assumeFormula: clipboardHasFormula(clipboard),
  });
  const cutSheetIndexes = cutPasteHistorySheetIndexes(clipboard, state.activeSheetIndex);
  const usesSheetHistory = tableEditCandidate || cutSheetIndexes.length;
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
    for (let rowOffset = 1; rowOffset <= pasteSize.targetRowCount; rowOffset += 1) {
      for (let colOffset = 1; colOffset <= pasteSize.targetColCount; colOffset += 1) {
        const sourceOffset = repeatedClipboardSourceOffset(rowOffset, colOffset, pasteSize);
        const row = range.top + rowOffset - 1;
        const col = range.left + colOffset - 1;
        const sourceCell = entriesByOffset.get(cellKey(sourceOffset.row, sourceOffset.col));
        writeCellContent(sheet, sheetId, row, col, pasteValueFromClipboardCell(sourceCell, {
          clipboard,
          targetRow: row,
          targetCol: col,
          sourceOffset,
        }));
        changed += 1;
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
      ? `${tableEdit.label} / ${changed} 個のセルを貼り付け（貼り付け先の書式に合わせる）`
      : `${changed} 個のセルを貼り付け（貼り付け先の書式に合わせる）`;
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
    setStatus(tableEdit.changed ? `${tableEdit.status} ${changed} 個のセルを貼り付けました。` : `${changed} 個のセルを貼り付けました（貼り付け先の書式に合わせました）。`);
  } catch (error) {
    console.error(error);
    setStatus(`貼り付けに失敗しました：${error.message || error}`);
    renderSheet();
  }
}



function pasteValueFromClipboardCell(sourceCell, context = {}) {
  if (!sourceCell) return null;
  const formula = shiftedClipboardFormulaForTarget(sourceCell, context);
  if (formula) return formula;
  if (typeof sourceCell.raw === "string" && sourceCell.raw.trim().startsWith("=")) {
    return shiftedClipboardFormulaTextForTarget(sourceCell.raw, context);
  }
  if (sourceCell.raw != null) return sourceCell.raw;
  if (sourceCell.cached != null) return sourceCell.cached;
  if (sourceCell.display != null && sourceCell.display !== "") return sourceCell.display;
  return null;
}



function writeClonedClipboardCellContent(sheet, row, col, clipboard, sourceCell, sourceOffset = null) {
  writeClonedCellContent(sheet, row, col, sourceCell, {
    clipboard,
    targetRow: row,
    targetCol: col,
    sourceOffset,
  });
}



function writeClonedCellContent(sheet, row, col, sourceCell, context = {}) {
  const key = cellKey(row, col);
  if (sourceCell) {
    const nextCell = cloneCellModel(sourceCell);
    const formula = shiftedClipboardFormulaForTarget(nextCell, context);
    if (formula) {
      nextCell.formula = formula;
      nextCell.raw = formula;
      nextCell.cached = null;
      nextCell.display = "";
    } else if (typeof nextCell.raw === "string" && nextCell.raw.trim().startsWith("=")) {
      const shiftedRaw = shiftedClipboardFormulaTextForTarget(nextCell.raw, context);
      if (shiftedRaw !== nextCell.raw) {
        nextCell.raw = shiftedRaw;
        nextCell.formula = shiftedRaw;
        nextCell.cached = null;
        nextCell.display = "";
      }
    }
    nextCell.row = row;
    nextCell.col = col;
    storeCellModelForSheet(sheet, key, nextCell);
  } else {
    delete sheet.cells[key];
  }
}



function shiftedClipboardFormulaForTarget(sourceCell, context = {}) {
  const formula = typeof sourceCell?.formula === "string" && sourceCell.formula.trim().startsWith("=")
    ? sourceCell.formula
    : "";
  return formula ? shiftedClipboardFormulaTextForTarget(formula, context) : "";
}



function shiftedClipboardFormulaTextForTarget(formula, context = {}) {
  const text = String(formula || "");
  if (!text.trim().startsWith("=")) return text;
  if (context.clipboard?.mode === "cut" || context.clipboard?.cut === true) return text;
  const offset = clipboardFormulaPasteOffset(context);
  if (!offset || (!offset.rowOffset && !offset.colOffset)) return text;
  return shiftFormulaReferences(text, offset.rowOffset, offset.colOffset);
}



function clipboardFormulaPasteOffset(context = {}) {
  const clipboard = context.clipboard;
  const sourceOffset = context.sourceOffset;
  const targetRow = Number(context.targetRow);
  const targetCol = Number(context.targetCol);
  if (!clipboard?.sourceRange || !sourceOffset || !Number.isFinite(targetRow) || !Number.isFinite(targetCol)) return null;
  const sourceRow = clipboard.sourceRange.top + (Number(sourceOffset.row) || 1) - 1;
  const sourceCol = clipboard.sourceRange.left + (Number(sourceOffset.col) || 1) - 1;
  return {
    rowOffset: targetRow - sourceRow,
    colOffset: targetCol - sourceCol,
  };
}



function pasteTextAtRange(text, range) {
  const sheet = activeSheet();
  const rows = String(text ?? "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => line.split("\t"));
  let changed = 0;
  const targetRows = range.top + rows.length - 1;
  const targetCols = range.left + Math.max(...rows.map((row) => row.length), 1) - 1;
  const targetRange = normalizeRange(
    { sheetIndex: state.activeSheetIndex, row: range.top, col: range.left },
    { sheetIndex: state.activeSheetIndex, row: targetRows, col: targetCols },
  );
  if (!ensureEditableRangeForProtection(sheet, targetRange, "貼り付け")) return;
  const hasPastedContent = rows.some((values) => values.some((value) => String(value ?? "") !== ""));
  const hasPastedFormula = rows.some((values) => values.some((value) => String(value ?? "").trim().startsWith("=")));
  const tableEditCandidate = tableRangeEditMayChangeMetadata(sheet, targetRange, {
    assumeNonEmpty: hasPastedContent,
    assumeFormula: hasPastedFormula,
  });
  const sheetBefore = tableEditCandidate ? captureSheetForHistory(sheet) : null;
  const sheetsBefore = tableEditCandidate ? captureSheetsForHistory(state.model.sheets.map((_item, index) => index)) : null;
  const before = captureCellsForHistory(sheet, targetRange);

  try {
    ensureSheetSize(sheet, targetRows, targetCols);
    const sheetId = getSheetId(sheet.name);
    rows.forEach((values, rowOffset) => {
      values.forEach((value, colOffset) => {
        const row = range.top + rowOffset;
        const col = range.left + colOffset;
        writeCellContent(sheet, sheetId, row, col, coerceUserInput(value));
        changed += 1;
      });
    });
    const tableEdit = applyTableRangeEditBehaviors(sheet, targetRange);
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
    const label = tableEdit.changed ? `${tableEdit.label} / ${changed} 個のセルを貼り付け` : `${changed} 個のセルを貼り付け`;
    pushHistory(tableEdit.changed
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
    setStatus(tableEdit.changed ? `${tableEdit.status} ${changed} 個のセルを貼り付けました。` : `${changed} 個のセルを貼り付けました。`);
  } catch (error) {
    console.error(error);
    setStatus(`貼り付けに失敗しました：${error.message || error}`);
    renderSheet();
  }
}



function applyCellStructureOperation(kind, option, range) {
  if (!range || range.sheetIndex !== state.activeSheetIndex) return false;
  const sheet = activeSheet();
  const normalizedRange = materializeSelectionRange(range);
  if (!normalizedRange) return false;
  if (applyCellStructureOperationToSheetGroup(kind, option, normalizedRange)) return true;
  if (!ensureStructureAllowedForProtection(kind, option)) return false;

  const referenceShift = structureReferenceShiftForOperation(kind, option, normalizedRange);
  const affectedSheetIndexes = referenceShift ? state.model.sheets.map((_item, index) => index) : null;
  const sheetsBefore = affectedSheetIndexes ? captureSheetsForHistory(affectedSheetIndexes) : null;
  const definedNamesBefore = referenceShift ? cloneDefinedNames(state.model.definedNames) : null;
  const before = sheetsBefore ? null : captureSheetForHistory(sheet);
  const selectionBefore = { ...normalizedRange };
  const selectionAfter = { ...normalizedRange };

  try {
    if (kind === "insert") {
      applyInsertOperation(sheet, normalizedRange, option);
    } else {
      applyDeleteOperation(sheet, normalizedRange, option);
    }
    if (referenceShift) {
      applyWorkbookReferenceShiftForStructureChange(referenceShift);
    }
    normalizeSheetModelAfterStructureChange(sheet);
    state.hf = buildFormulaEngine(state.model);

    state.selected = { sheetIndex: state.activeSheetIndex, row: selectionAfter.top, col: selectionAfter.left };
    state.selectionRange = selectionAfter;
    state.selectionRanges = [selectionAfter];
    state.selectionStart = state.selected;
    state.selectionDrag = null;
    state.selectedImage = null;
    state.formulaEdit = null;

    pushHistory({
      sheetIndex: state.activeSheetIndex,
      ...(sheetsBefore
        ? {
            sheetsBefore,
            sheetsAfter: captureSheetsForHistory(affectedSheetIndexes),
            definedNamesBefore,
            definedNamesAfter: cloneDefinedNames(state.model.definedNames),
          }
        : {
            sheetBefore: before,
            sheetAfter: captureSheetForHistory(sheet),
          }),
      selectionBefore,
      selectionAfter,
      label: `${rangeToLabel(selectionBefore)} ${kind === "insert" ? "挿入" : "削除"}`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${rangeToLabel(selectionBefore)} に${kind === "insert" ? "挿入" : "削除"}を適用しました。`);
    return true;
  } catch (error) {
    console.error(error);
    if (sheetsBefore) restoreSheetsFromHistory(sheetsBefore);
    else restoreSheetFromHistory(sheet, before);
    if (definedNamesBefore) state.model.definedNames = cloneDefinedNames(definedNamesBefore);
    state.hf = buildFormulaEngine(state.model);
    renderSheet();
    showMessageBox("この操作は実行できません。", { title: "Cht WebSheet" });
    setStatus(`${kind === "insert" ? "挿入" : "削除"}に失敗しました：${error.message || error}`);
    return false;
  }
}



function applyCellStructureOperationToSheetGroup(kind, option, range) {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1 || !range) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet) continue;
    const targetRange = materializeSelectionRangeForSheet(sheet, range, sheetIndex);
    if (!targetRange) continue;
    if (!ensureStructureAllowedForSheet(sheet, kind, option)) return true;
    items.push({ sheetIndex, sheet, range: targetRange });
  }
  if (!items.length) return true;

  const referenceShifts = items
    .map((item) => structureReferenceShiftForOperation(kind, option, item.range, item.sheetIndex))
    .filter(Boolean);
  const affectedSheetIndexes = referenceShifts.length
    ? state.model.sheets.map((_item, index) => index)
    : items.map((item) => item.sheetIndex);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  const definedNamesBefore = referenceShifts.length ? cloneDefinedNames(state.model.definedNames) : null;
  const selectionBefore = { ...range };
  const selectionBeforeRanges = activeSelectionRanges().map((item) => ({ ...item }));
  const selectionAfter = { ...range };

  try {
    items.forEach((item) => {
      if (kind === "insert") {
        applyInsertOperation(item.sheet, item.range, option);
      } else {
        applyDeleteOperation(item.sheet, item.range, option);
      }
      normalizeSheetModelAfterStructureChange(item.sheet);
    });
    referenceShifts.forEach((referenceShift) => {
      applyWorkbookReferenceShiftForStructureChange(referenceShift);
    });
    state.hf = buildFormulaEngine(state.model);

    state.selected = { sheetIndex: state.activeSheetIndex, row: selectionAfter.top, col: selectionAfter.left };
    state.selectionRange = selectionAfter;
    state.selectionRanges = selectionBeforeRanges.length ? selectionBeforeRanges.map((item) => ({ ...item })) : [selectionAfter];
    state.selectionStart = state.selected;
    state.selectionDrag = null;
    state.selectedImage = null;
    state.formulaEdit = null;

    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(affectedSheetIndexes),
      ...(definedNamesBefore
        ? {
            definedNamesBefore,
            definedNamesAfter: cloneDefinedNames(state.model.definedNames),
          }
        : {}),
      selectionBefore,
      selectionAfter,
      selectionBeforeRanges,
      selectionAfterRanges: state.selectionRanges.map((item) => ({ ...item })),
      label: `${rangeToLabel(selectionBefore)} ${sheetIndexes.length} シート${kind === "insert" ? "挿入" : "削除"}`,
    });
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`${sheetIndexes.length} 個のシートの${rangeToLabel(selectionBefore)}に${kind === "insert" ? "挿入" : "削除"}を適用しました。`);
    return true;
  } catch (error) {
    console.error(error);
    restoreSheetsFromHistory(sheetsBefore);
    if (definedNamesBefore) state.model.definedNames = cloneDefinedNames(definedNamesBefore);
    state.hf = buildFormulaEngine(state.model);
    renderSheet();
    showMessageBox("この操作は実行できません。", { title: "Cht WebSheet" });
    setStatus(`${kind === "insert" ? "挿入" : "削除"}に失敗しました：${error.message || error}`);
    return true;
  }
}



function applyInsertOperation(sheet, range, option) {
  switch (option) {
    case "shift-right":
      insertCellsShiftRight(sheet, range);
      break;
    case "shift-down":
      insertCellsShiftDown(sheet, range);
      break;
    case "entire-row":
      insertEntireRows(sheet, range.top, range.bottom - range.top + 1);
      break;
    case "entire-column":
      insertEntireColumns(sheet, range.left, range.right - range.left + 1);
      break;
    default:
      throw new Error("Unknown insert option");
  }
}



function applyDeleteOperation(sheet, range, option) {
  switch (option) {
    case "shift-left":
      deleteCellsShiftLeft(sheet, range);
      break;
    case "shift-up":
      deleteCellsShiftUp(sheet, range);
      break;
    case "entire-row":
      deleteEntireRows(sheet, range.top, range.bottom - range.top + 1);
      break;
    case "entire-column":
      deleteEntireColumns(sheet, range.left, range.right - range.left + 1);
      break;
    default:
      throw new Error("Unknown delete option");
  }
}



function insertCellsShiftRight(sheet, range) {
  const amount = range.right - range.left + 1;
  const oldColCount = sheet.colCount;
  ensureSheetSize(sheet, sheet.rowCount, oldColCount + amount);
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = oldColCount; col >= range.left; col -= 1) {
      writeClonedCellContent(sheet, row, col + amount, cloneCellModel(sheet.cells[cellKey(row, col)] || null));
    }
    for (let col = range.left; col < range.left + amount; col += 1) {
      writeClonedCellContent(sheet, row, col, null);
    }
  }
  removeIntersectingMerges(sheet, { ...range, right: oldColCount + amount });
}



function insertCellsShiftDown(sheet, range) {
  const amount = range.bottom - range.top + 1;
  const oldRowCount = sheet.rowCount;
  ensureSheetSize(sheet, oldRowCount + amount, sheet.colCount);
  for (let col = range.left; col <= range.right; col += 1) {
    for (let row = oldRowCount; row >= range.top; row -= 1) {
      writeClonedCellContent(sheet, row + amount, col, cloneCellModel(sheet.cells[cellKey(row, col)] || null));
    }
    for (let row = range.top; row < range.top + amount; row += 1) {
      writeClonedCellContent(sheet, row, col, null);
    }
  }
  removeIntersectingMerges(sheet, { ...range, bottom: oldRowCount + amount });
}



function deleteCellsShiftLeft(sheet, range) {
  const amount = range.right - range.left + 1;
  const oldColCount = sheet.colCount;
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= oldColCount - amount; col += 1) {
      writeClonedCellContent(sheet, row, col, cloneCellModel(sheet.cells[cellKey(row, col + amount)] || null));
    }
    for (let col = Math.max(range.left, oldColCount - amount + 1); col <= oldColCount; col += 1) {
      writeClonedCellContent(sheet, row, col, null);
    }
  }
  removeIntersectingMerges(sheet, { ...range, right: oldColCount });
}



function deleteCellsShiftUp(sheet, range) {
  const amount = range.bottom - range.top + 1;
  const oldRowCount = sheet.rowCount;
  for (let col = range.left; col <= range.right; col += 1) {
    for (let row = range.top; row <= oldRowCount - amount; row += 1) {
      writeClonedCellContent(sheet, row, col, cloneCellModel(sheet.cells[cellKey(row + amount, col)] || null));
    }
    for (let row = Math.max(range.top, oldRowCount - amount + 1); row <= oldRowCount; row += 1) {
      writeClonedCellContent(sheet, row, col, null);
    }
  }
  removeIntersectingMerges(sheet, { ...range, bottom: oldRowCount });
}



function insertEntireRows(sheet, startRow, amount) {
  const oldRowCount = sheet.rowCount;
  const colCount = sheet.colCount;
  ensureSheetSize(sheet, oldRowCount + amount, colCount);
  for (let row = oldRowCount; row >= startRow; row -= 1) {
    for (let col = 1; col <= colCount; col += 1) {
      writeClonedCellContent(sheet, row + amount, col, cloneCellModel(sheet.cells[cellKey(row, col)] || null));
    }
    sheet.rows[row + amount - 1] = { ...(sheet.rows[row - 1] || defaultRowModel(sheet)) };
  }
  for (let row = startRow; row < startRow + amount; row += 1) {
    for (let col = 1; col <= colCount; col += 1) writeClonedCellContent(sheet, row, col, null);
    sheet.rows[row - 1] = defaultRowModel(sheet);
  }
  shiftMergesForRowInsert(sheet, startRow, amount);
  shiftStyleRangesForRows(sheet, startRow, amount, "insert");
  shiftImagesForRows(sheet, startRow, amount, "insert");
  shiftTablesForRows(sheet, startRow, amount, "insert");
  shiftSheetFilterAndSortForRows(sheet, startRow, amount, "insert");
  syncSheetAutoFilterToTableRangesAfterStructureChange(sheet);
}



function insertEntireColumns(sheet, startCol, amount) {
  const oldColCount = sheet.colCount;
  const rowCount = sheet.rowCount;
  ensureSheetSize(sheet, rowCount, oldColCount + amount);
  for (let col = oldColCount; col >= startCol; col -= 1) {
    for (let row = 1; row <= rowCount; row += 1) {
      writeClonedCellContent(sheet, row, col + amount, cloneCellModel(sheet.cells[cellKey(row, col)] || null));
    }
    sheet.columns[col + amount - 1] = { ...(sheet.columns[col - 1] || defaultColumnModel(sheet)) };
  }
  for (let col = startCol; col < startCol + amount; col += 1) {
    for (let row = 1; row <= rowCount; row += 1) writeClonedCellContent(sheet, row, col, null);
    sheet.columns[col - 1] = defaultColumnModel(sheet);
  }
  shiftMergesForColumnInsert(sheet, startCol, amount);
  shiftStyleRangesForColumns(sheet, startCol, amount, "insert");
  shiftImagesForColumns(sheet, startCol, amount, "insert");
  shiftTablesForColumns(sheet, startCol, amount, "insert");
  shiftSheetFilterAndSortForColumns(sheet, startCol, amount, "insert");
  syncSheetAutoFilterToTableRangesAfterStructureChange(sheet);
}



function deleteEntireRows(sheet, startRow, amount) {
  const oldRowCount = sheet.rowCount;
  const colCount = sheet.colCount;
  const endRow = Math.min(oldRowCount, startRow + amount - 1);
  const actualAmount = endRow - startRow + 1;
  for (let row = startRow; row <= oldRowCount - actualAmount; row += 1) {
    for (let col = 1; col <= colCount; col += 1) {
      writeClonedCellContent(sheet, row, col, cloneCellModel(sheet.cells[cellKey(row + actualAmount, col)] || null));
    }
    sheet.rows[row - 1] = { ...(sheet.rows[row + actualAmount - 1] || defaultRowModel(sheet)) };
  }
  for (let row = oldRowCount - actualAmount + 1; row <= oldRowCount; row += 1) {
    for (let col = 1; col <= colCount; col += 1) writeClonedCellContent(sheet, row, col, null);
    sheet.rows[row - 1] = defaultRowModel(sheet);
  }
  shiftMergesForRowDelete(sheet, startRow, actualAmount);
  shiftStyleRangesForRows(sheet, startRow, actualAmount, "delete");
  shiftImagesForRows(sheet, startRow, actualAmount, "delete");
  shiftTablesForRows(sheet, startRow, actualAmount, "delete");
  shiftSheetFilterAndSortForRows(sheet, startRow, actualAmount, "delete");
  syncSheetAutoFilterToTableRangesAfterStructureChange(sheet);
}



function deleteEntireColumns(sheet, startCol, amount) {
  const oldColCount = sheet.colCount;
  const rowCount = sheet.rowCount;
  const endCol = Math.min(oldColCount, startCol + amount - 1);
  const actualAmount = endCol - startCol + 1;
  for (let col = startCol; col <= oldColCount - actualAmount; col += 1) {
    for (let row = 1; row <= rowCount; row += 1) {
      writeClonedCellContent(sheet, row, col, cloneCellModel(sheet.cells[cellKey(row, col + actualAmount)] || null));
    }
    sheet.columns[col - 1] = { ...(sheet.columns[col + actualAmount - 1] || defaultColumnModel(sheet)) };
  }
  for (let col = oldColCount - actualAmount + 1; col <= oldColCount; col += 1) {
    for (let row = 1; row <= rowCount; row += 1) writeClonedCellContent(sheet, row, col, null);
    sheet.columns[col - 1] = defaultColumnModel(sheet);
  }
  shiftMergesForColumnDelete(sheet, startCol, actualAmount);
  shiftStyleRangesForColumns(sheet, startCol, actualAmount, "delete");
  shiftImagesForColumns(sheet, startCol, actualAmount, "delete");
  shiftTablesForColumns(sheet, startCol, actualAmount, "delete");
  shiftSheetFilterAndSortForColumns(sheet, startCol, actualAmount, "delete");
  syncSheetAutoFilterToTableRangesAfterStructureChange(sheet);
}



function defaultRowModel(sheet) {
  return { height: sheet.defaultRowHeight || rowHeightToPixels(15), hidden: false };
}



function defaultColumnModel(sheet) {
  return { width: sheet.defaultColumnWidth || columnWidthToPixels(8.43), hidden: false };
}



function shiftMergesForRowInsert(sheet, startRow, amount) {
  sheet.merges = (sheet.merges || []).map((merge) => {
    if (merge.bottom < startRow) return { ...merge };
    if (merge.top >= startRow) return { ...merge, top: merge.top + amount, bottom: merge.bottom + amount };
    return { ...merge, bottom: merge.bottom + amount };
  });
}



function shiftMergesForColumnInsert(sheet, startCol, amount) {
  sheet.merges = (sheet.merges || []).map((merge) => {
    if (merge.right < startCol) return { ...merge };
    if (merge.left >= startCol) return { ...merge, left: merge.left + amount, right: merge.right + amount };
    return { ...merge, right: merge.right + amount };
  });
}



function shiftMergesForRowDelete(sheet, startRow, amount) {
  const endRow = startRow + amount - 1;
  sheet.merges = (sheet.merges || [])
    .map((merge) => {
      if (merge.bottom < startRow) return { ...merge };
      if (merge.top > endRow) return { ...merge, top: merge.top - amount, bottom: merge.bottom - amount };
      if (merge.top >= startRow && merge.bottom <= endRow) return null;
      if (merge.top < startRow && merge.bottom > endRow) return { ...merge, bottom: merge.bottom - amount };
      if (merge.top < startRow) return { ...merge, bottom: startRow - 1 };
      return { ...merge, top: startRow, bottom: merge.bottom - amount };
    })
    .filter((merge) => merge && merge.top <= merge.bottom && merge.left <= merge.right);
}



function shiftMergesForColumnDelete(sheet, startCol, amount) {
  const endCol = startCol + amount - 1;
  sheet.merges = (sheet.merges || [])
    .map((merge) => {
      if (merge.right < startCol) return { ...merge };
      if (merge.left > endCol) return { ...merge, left: merge.left - amount, right: merge.right - amount };
      if (merge.left >= startCol && merge.right <= endCol) return null;
      if (merge.left < startCol && merge.right > endCol) return { ...merge, right: merge.right - amount };
      if (merge.left < startCol) return { ...merge, right: startCol - 1 };
      return { ...merge, left: startCol, right: merge.right - amount };
    })
    .filter((merge) => merge && merge.top <= merge.bottom && merge.left <= merge.right);
}



function shiftStyleRangesForRows(sheet, startRow, amount, action) {
  const ranges = Array.isArray(sheet?.styleRanges) ? sheet.styleRanges : [];
  if (!ranges.length) return;
  sheet.styleRanges = ranges
    .map((range) => shiftRangeRows(range, startRow, amount, action))
    .filter(Boolean);
  normalizeSheetStyleRanges(sheet);
  invalidateSheetUsedRangeCache(sheet);
}



function shiftStyleRangesForColumns(sheet, startCol, amount, action) {
  const ranges = Array.isArray(sheet?.styleRanges) ? sheet.styleRanges : [];
  if (!ranges.length) return;
  sheet.styleRanges = ranges
    .map((range) => shiftRangeColumns(range, startCol, amount, action))
    .filter(Boolean);
  normalizeSheetStyleRanges(sheet);
  invalidateSheetUsedRangeCache(sheet);
}



function shiftImagesForRows(sheet, startRow, amount, action) {
  const zeroBasedStart = startRow - 1;
  (sheet.images || []).forEach((image) => {
    shiftImageAnchor(image.tl, "row", zeroBasedStart, amount, action);
    shiftImageAnchor(image.br, "row", zeroBasedStart, amount, action);
  });
}



function shiftImagesForColumns(sheet, startCol, amount, action) {
  const zeroBasedStart = startCol - 1;
  (sheet.images || []).forEach((image) => {
    shiftImageAnchor(image.tl, "col", zeroBasedStart, amount, action);
    shiftImageAnchor(image.br, "col", zeroBasedStart, amount, action);
  });
}



function shiftTablesForRows(sheet, startRow, amount, action) {
  if (!sheet?.tables?.length) return;
  sheet.tables = sheet.tables
    .map((table) => {
      const oldRange = table.range;
      const nextRange = shiftRangeRows(oldRange, startRow, amount, action);
      if (!nextRange) return null;
      const nextTable = { ...table, range: nextRange, ref: rangeToLabel(nextRange) };
      if (table.autoFilterState?.range) {
        const nextFilterRange = shiftRangeRows(table.autoFilterState.range, startRow, amount, action);
        nextTable.autoFilterState = {
          ...table.autoFilterState,
          range: nextFilterRange || tableAutoFilterRange(nextTable),
        };
      }
      if (action === "insert") {
        initializeInsertedTableRowsForRowInsert(sheet, nextTable, oldRange, startRow, amount);
      } else if (tableStructureChangeTouchesRows(oldRange, startRow, amount)) {
        refreshImportedTableVisualStyleForTableMutation(sheet, nextTable);
        ensureTableAutoFilterStateForStructureChange(nextTable);
      }
      return nextTable;
    })
    .filter(Boolean);
}



function initializeInsertedTableRowsForRowInsert(sheet, table, oldRange, startRow, amount) {
  if (!sheet || !table?.range || !oldRange) return;
  if (startRow <= oldRange.top || startRow > oldRange.bottom) return;
  const dataRows = structuredReferenceRows(table, "data");
  if (!dataRows) return;
  const firstInsertedRow = Math.max(startRow, dataRows.top);
  const lastInsertedRow = Math.min(startRow + amount - 1, dataRows.bottom);
  if (firstInsertedRow > lastInsertedRow) return;
  ensureTableColumnMetadata(sheet, table);
  for (let row = firstInsertedRow; row <= lastInsertedRow; row += 1) {
    const sourceRow = tableInsertedRowFormatSource(table, row, amount);
    if (sourceRow) {
      copyTableRowMetrics(sheet, sourceRow, row);
      for (let col = table.range.left; col <= table.range.right; col += 1) {
        copyTableCellFormatting(sheet, sourceRow, col, row, col);
      }
      for (let col = table.range.left; col <= table.range.right; col += 1) {
        copyTableCalculatedColumnFormula(sheet, table, sourceRow, row, col);
      }
    }
  }
  refreshImportedTableVisualStyleForTableMutation(sheet, table);
  ensureTableAutoFilterStateForStructureChange(table);
  internCellsInRange(sheet, { ...table.range, top: firstInsertedRow, bottom: lastInsertedRow });
}



function tableInsertedRowFormatSource(table, insertedRow, amount) {
  const dataRows = structuredReferenceRows(table, "data");
  if (!dataRows) return null;
  const rowAfterInsert = insertedRow + amount;
  if (rowAfterInsert >= dataRows.top && rowAfterInsert <= dataRows.bottom) return rowAfterInsert;
  const rowBeforeInsert = insertedRow - 1;
  if (rowBeforeInsert >= dataRows.top && rowBeforeInsert <= dataRows.bottom) return rowBeforeInsert;
  return null;
}



function copyTableRowMetrics(sheet, sourceRow, targetRow) {
  const source = sheet?.rows?.[sourceRow - 1];
  if (!source) return;
  const target = { ...(sheet.rows[targetRow - 1] || defaultRowModel(sheet)) };
  if (source.height != null) target.height = source.height;
  if (source.customHeight != null) target.customHeight = source.customHeight;
  target.hidden = false;
  sheet.rows[targetRow - 1] = target;
}



function shiftTablesForColumns(sheet, startCol, amount, action) {
  if (!sheet?.tables?.length) return;
  sheet.tables = sheet.tables
    .map((table) => {
      const oldRange = table.range;
      const nextRange = shiftRangeColumns(oldRange, startCol, amount, action);
      if (!nextRange) return null;
      const nextTable = { ...table, range: nextRange, ref: rangeToLabel(nextRange) };
      if (action === "insert" && oldRange && startCol > oldRange.left && startCol <= oldRange.right) {
        nextTable.columns = insertTableColumnMetadata(table.columns, startCol - oldRange.left, amount);
      } else if (action === "delete" && oldRange && rangesOverlap1d(startCol, startCol + amount - 1, oldRange.left, oldRange.right)) {
        nextTable.columns = deleteTableColumnMetadata(table.columns, oldRange, startCol, amount);
      }
      if (table.autoFilterState?.range) {
        const nextFilterRange = shiftRangeColumns(table.autoFilterState.range, startCol, amount, action);
        nextTable.autoFilterState = {
          ...table.autoFilterState,
          range: nextFilterRange || tableAutoFilterRange(nextTable),
          criteria: shiftFilterColumnMap(table.autoFilterState.criteria, startCol, amount, action),
          ...(table.autoFilterState.hiddenButtons
            ? { hiddenButtons: shiftFilterColumnMap(table.autoFilterState.hiddenButtons, startCol, amount, action) }
            : {}),
        };
      }
      ensureTableColumnMetadata(sheet, nextTable);
      if (action === "insert") {
        initializeInsertedTableColumnsForColumnInsert(sheet, nextTable, oldRange, startCol, amount);
      } else if (tableStructureChangeTouchesColumns(oldRange, startCol, amount)) {
        syncTableColumnNamesFromHeaders(sheet, nextTable);
        refreshImportedTableVisualStyleForTableMutation(sheet, nextTable);
        ensureTableAutoFilterStateForStructureChange(nextTable);
      }
      return nextTable;
    })
    .filter(Boolean);
}



function initializeInsertedTableColumnsForColumnInsert(sheet, table, oldRange, startCol, amount) {
  if (!sheet || !table?.range || !oldRange) return;
  if (startCol <= oldRange.left || startCol > oldRange.right) return;
  const firstInsertedCol = Math.max(startCol, table.range.left);
  const lastInsertedCol = Math.min(startCol + amount - 1, table.range.right);
  if (firstInsertedCol > lastInsertedCol) return;
  ensureTableColumnMetadata(sheet, table);
  for (let col = firstInsertedCol; col <= lastInsertedCol; col += 1) {
    const sourceCol = tableInsertedColumnFormatSource(table, col, amount);
    if (sourceCol) {
      copyTableColumnMetrics(sheet, sourceCol, col);
      for (let row = table.range.top; row <= table.range.bottom; row += 1) {
        copyTableCellFormatting(sheet, row, sourceCol, row, col);
      }
    }
    const offset = col - table.range.left;
    const headerName = table.columns?.[offset]?.name || `Column${offset + 1}`;
    if (table.showHeaderRow !== false && !isOccupiedCell(sheet.cells[cellKey(table.range.top, col)])) {
      updateCellModel(sheet, table.range.top, col, headerName);
    }
  }
  syncTableColumnNamesFromHeaders(sheet, table);
  refreshImportedTableVisualStyleForTableMutation(sheet, table);
  ensureTableAutoFilterStateForStructureChange(table);
  internCellsInRange(sheet, { ...table.range, left: firstInsertedCol, right: lastInsertedCol });
}



function tableInsertedColumnFormatSource(table, insertedCol, amount) {
  const range = table?.range;
  if (!range) return null;
  const colAfterInsert = insertedCol + amount;
  if (colAfterInsert >= range.left && colAfterInsert <= range.right) return colAfterInsert;
  const colBeforeInsert = insertedCol - 1;
  if (colBeforeInsert >= range.left && colBeforeInsert <= range.right) return colBeforeInsert;
  return null;
}



function copyTableColumnMetrics(sheet, sourceCol, targetCol) {
  const source = sheet?.columns?.[sourceCol - 1];
  if (!source) return;
  const target = { ...(sheet.columns[targetCol - 1] || defaultColumnModel(sheet)) };
  if (source.width != null) target.width = source.width;
  if (source.customWidth != null) target.customWidth = source.customWidth;
  target.hidden = false;
  sheet.columns[targetCol - 1] = target;
}



function ensureTableAutoFilterStateForStructureChange(table) {
  if (!table?.range || table.showFilterButton === false) return;
  const range = tableAutoFilterRange(table);
  if (!range) return;
  const source = table.autoFilterState || {};
  table.autoFilterState = {
    ...source,
    range: { ...range },
    criteria: cloneJsonLikeValue(source.criteria || {}),
    ...(source.hiddenButtons ? { hiddenButtons: cloneJsonLikeValue(source.hiddenButtons) } : {}),
    ...(source.advancedCriteria ? { advancedCriteria: cloneJsonLikeValue(source.advancedCriteria) } : {}),
  };
}



function syncSheetAutoFilterToTableRangesAfterStructureChange(sheet) {
  const filter = sheet?.autoFilter;
  const filterRange = filter?.range;
  if (!filterRange || !sheet?.tables?.length) return;
  const table = (sheet.tables || [])
    .map((candidate) => {
      const range = tableAutoFilterRange(candidate);
      if (!range || candidate.showFilterButton === false || filterRange.top !== range.top) return null;
      const overlap = Math.max(0, Math.min(filterRange.right, range.right) - Math.max(filterRange.left, range.left) + 1);
      if (!overlap) return null;
      const exactSides = Number(filterRange.left === range.left) + Number(filterRange.right === range.right);
      const exactBottom = Number(filterRange.bottom === range.bottom);
      return { table: candidate, range, overlap, exactSides, exactBottom };
    })
    .filter(Boolean)
    .sort((a, b) => b.exactSides - a.exactSides || b.exactBottom - a.exactBottom || b.overlap - a.overlap || a.range.left - b.range.left)[0]?.table;
  if (!table) return;
  const nextRange = tableAutoFilterRange(table);
  if (!nextRange) return;
  const source = table.autoFilterState || {};
  const criteria = Object.keys(filter.criteria || {}).length ? filter.criteria : source.criteria || {};
  const hiddenButtons = filter.hiddenButtons || source.hiddenButtons || {};
  table.autoFilterState = {
    ...source,
    range: { ...nextRange },
    criteria: cloneJsonLikeValue(criteria),
    ...(Object.keys(hiddenButtons || {}).length ? { hiddenButtons: cloneJsonLikeValue(hiddenButtons) } : {}),
    ...(filter.advancedCriteria || source.advancedCriteria
      ? { advancedCriteria: cloneJsonLikeValue(filter.advancedCriteria || source.advancedCriteria) }
      : {}),
  };
  sheet.autoFilter = {
    ...filter,
    range: { ...nextRange },
    criteria: cloneJsonLikeValue(table.autoFilterState.criteria || {}),
    ...(table.autoFilterState.hiddenButtons ? { hiddenButtons: cloneJsonLikeValue(table.autoFilterState.hiddenButtons) } : {}),
    ...(table.autoFilterState.advancedCriteria ? { advancedCriteria: cloneJsonLikeValue(table.autoFilterState.advancedCriteria) } : {}),
  };
}



function tableStructureChangeTouchesRows(range, startRow, amount) {
  if (!range) return false;
  return rangesOverlap1d(startRow, startRow + amount - 1, range.top, range.bottom);
}



function tableStructureChangeTouchesColumns(range, startCol, amount) {
  if (!range) return false;
  return rangesOverlap1d(startCol, startCol + amount - 1, range.left, range.right);
}



function insertTableColumnMetadata(columns, offset, amount) {
  const next = Array.isArray(columns) ? columns.map((column) => ({ ...column })) : [];
  const inserts = Array.from({ length: amount }, (_item, index) => ({
    id: String(next.length + index + 1),
    name: `Column${offset + index + 1}`,
  }));
  next.splice(Math.max(0, offset), 0, ...inserts);
  return next;
}



function deleteTableColumnMetadata(columns, oldRange, startCol, amount) {
  const next = Array.isArray(columns) ? columns.map((column) => ({ ...column })) : [];
  const deleteStart = Math.max(oldRange.left, startCol);
  const deleteEnd = Math.min(oldRange.right, startCol + amount - 1);
  if (deleteStart > deleteEnd) return next;
  next.splice(deleteStart - oldRange.left, deleteEnd - deleteStart + 1);
  return next;
}



function shiftSheetFilterAndSortForRows(sheet, startRow, amount, action) {
  if (sheet.autoFilter?.range) {
    const nextRange = shiftRangeRows(sheet.autoFilter.range, startRow, amount, action);
    if (nextRange) sheet.autoFilter.range = nextRange;
    else delete sheet.autoFilter;
  }
  if (sheet.sortState?.range) {
    const nextRange = shiftRangeRows(sheet.sortState.range, startRow, amount, action);
    if (nextRange) {
      sheet.sortState.range = nextRange;
      if (!shiftSortStateKeysForRows(sheet.sortState, startRow, amount, action)) delete sheet.sortState;
    } else {
      delete sheet.sortState;
    }
  }
}



function shiftSheetFilterAndSortForColumns(sheet, startCol, amount, action) {
  if (sheet.autoFilter?.range) {
    const nextRange = shiftRangeColumns(sheet.autoFilter.range, startCol, amount, action);
    if (nextRange) {
      sheet.autoFilter.range = nextRange;
      sheet.autoFilter.criteria = shiftFilterColumnMap(sheet.autoFilter.criteria, startCol, amount, action);
      if (sheet.autoFilter.hiddenButtons) {
        sheet.autoFilter.hiddenButtons = shiftFilterColumnMap(sheet.autoFilter.hiddenButtons, startCol, amount, action);
      }
    } else {
      delete sheet.autoFilter;
    }
  }
  if (sheet.sortState?.range) {
    const nextRange = shiftRangeColumns(sheet.sortState.range, startCol, amount, action);
    if (nextRange) {
      sheet.sortState.range = nextRange;
      if (!shiftSortStateKeysForColumns(sheet.sortState, startCol, amount, action)) delete sheet.sortState;
    } else {
      delete sheet.sortState;
    }
  }
}



function shiftFilterColumnMap(map, startCol, amount, action) {
  if (!map || typeof map !== "object") return {};
  const next = {};
  Object.entries(map).forEach(([key, value]) => {
    const col = Number(key);
    if (!Number.isFinite(col)) return;
    const shifted = shiftStructureIndex(col, startCol, amount, action);
    if (shifted == null || shifted < 1) return;
    next[shifted] = cloneJsonLikeValue(value);
  });
  return next;
}



function shiftSortStateKeysForRows(sortState, startRow, amount, action) {
  if (!sortState || sortState.orientation !== "columns") return true;
  sortState.keys = (sortState.keys || [])
    .map((key) => {
      const shifted = shiftStructureIndex(Number(key.row), startRow, amount, action);
      return shifted == null ? null : { ...key, row: shifted };
    })
    .filter(Boolean);
  return sortState.keys.length > 0;
}



function shiftSortStateKeysForColumns(sortState, startCol, amount, action) {
  if (!sortState || sortState.orientation === "columns") return true;
  sortState.keys = (sortState.keys || [])
    .map((key) => {
      const shifted = shiftStructureIndex(Number(key.col), startCol, amount, action);
      return shifted == null ? null : { ...key, col: shifted };
    })
    .filter(Boolean);
  return sortState.keys.length > 0;
}



function shiftStructureIndex(index, start, amount, action) {
  const value = Number(index);
  const first = Number(start);
  const count = Number(amount);
  if (!Number.isFinite(value) || !Number.isFinite(first) || !Number.isFinite(count) || count <= 0) return null;
  if (action === "insert") return value >= first ? value + count : value;
  const end = first + count - 1;
  if (value < first) return value;
  if (value > end) return value - count;
  return null;
}



function shiftRangeRows(range, startRow, amount, action) {
  if (!range) return null;
  const next = { ...range };
  if (action === "insert") {
    if (next.bottom < startRow) return next;
    if (next.top >= startRow) {
      next.top += amount;
      next.bottom += amount;
    } else {
      next.bottom += amount;
    }
    return next;
  }
  const endRow = startRow + amount - 1;
  if (next.bottom < startRow) return next;
  if (next.top > endRow) {
    next.top -= amount;
    next.bottom -= amount;
    return next;
  }
  if (next.top >= startRow && next.bottom <= endRow) return null;
  if (next.top < startRow && next.bottom > endRow) {
    next.bottom -= amount;
    return next;
  }
  if (next.top < startRow) {
    next.bottom = startRow - 1;
    return next.top <= next.bottom ? next : null;
  }
  next.top = startRow;
  next.bottom -= amount;
  return next.top <= next.bottom ? next : null;
}



function shiftRangeColumns(range, startCol, amount, action) {
  if (!range) return null;
  const next = { ...range };
  if (action === "insert") {
    if (next.right < startCol) return next;
    if (next.left >= startCol) {
      next.left += amount;
      next.right += amount;
    } else {
      next.right += amount;
    }
    return next;
  }
  const endCol = startCol + amount - 1;
  if (next.right < startCol) return next;
  if (next.left > endCol) {
    next.left -= amount;
    next.right -= amount;
    return next;
  }
  if (next.left >= startCol && next.right <= endCol) return null;
  if (next.left < startCol && next.right > endCol) {
    next.right -= amount;
    return next;
  }
  if (next.left < startCol) {
    next.right = startCol - 1;
    return next.left <= next.right ? next : null;
  }
  next.left = startCol;
  next.right -= amount;
  return next.left <= next.right ? next : null;
}



function shiftImageAnchor(anchor, axis, start, amount, action) {
  if (!anchor || !Number.isFinite(Number(anchor[axis]))) return;
  if (action === "insert") {
    if (anchor[axis] >= start) anchor[axis] += amount;
  } else if (anchor[axis] >= start + amount) {
    anchor[axis] -= amount;
  } else if (anchor[axis] >= start) {
    anchor[axis] = start;
  }
  const nativeKey = axis === "row" ? "nativeRow" : "nativeCol";
  if (Number.isFinite(Number(anchor[nativeKey]))) anchor[nativeKey] = Math.max(0, Math.floor(anchor[axis]));
}



function normalizeSheetModelAfterStructureChange(sheet) {
  Object.keys(sheet.cells || {}).forEach((key) => {
    const [row, col] = parseCellKey(key);
    const cell = sheet.cells[key];
    if (!cell || row < 1 || col < 1 || row > sheet.rowCount || col > sheet.colCount || !isPersistentCell(cell)) {
      delete sheet.cells[key];
    }
  });
  normalizeSheetStyleRanges(sheet);
  invalidateSheetUsedRangeCache(sheet);
  internSheetCellStyles(sheet);
  syncHfDataFromCells(sheet);
}



function syncHfDataFromCells(sheet) {
  sheet.hfData = [];
}



function clearSelectionContentsForSheetGroup(label = "クリア") {
  const sheetIndexes = activeSheetGroupEditIndexes();
  if (sheetIndexes.length <= 1) return false;
  const range = activeSelectionRange();
  const sourceRanges = activeSelectionRanges();
  if (!range || !sourceRanges.length) return false;
  const items = [];
  for (const sheetIndex of sheetIndexes) {
    const sheet = state.model.sheets[sheetIndex];
    const targetRanges = materializedRangesForSheet(sheet, sourceRanges, sheetIndex);
    if (!targetRanges.length) continue;
    if (!ensureEditableRangesForProtection(sheet, targetRanges, "削除")) return true;
    const cells = uniqueCellsForRanges(targetRanges);
    if (cells.length) items.push({ sheetIndex, sheet, sheetId: getSheetId(sheet.name), targetRanges, ranges: targetRanges, cells });
  }
  if (!items.length) return true;
  const affectedSheetIndexes = affectedSheetIndexesForGroupedRanges(items);
  const sheetsBefore = captureSheetsForHistory(affectedSheetIndexes);
  try {
    items.forEach((item) => {
      item.cells.forEach(({ row, col }) => writeCellContent(item.sheet, item.sheetId, row, col, null));
    });
    refreshChartsForGroupedRanges(items);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      sheetsBefore,
      sheetsAfter: captureSheetsForHistory(sheetsBefore.map((snapshot) => snapshot.index)),
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: sourceRanges.map((item) => ({ ...item })),
      selectionAfterRanges: sourceRanges.map((item) => ({ ...item })),
      label: `${label} ${sheetIndexes.length} シート ${rangesToLabel(sourceRanges)}`,
    });
    renderSheet();
    updateSelectionUi();
    if (state.selected) updateFormulaBarForCell(state.selected.row, state.selected.col);
    const cellCount = items.reduce((sum, item) => sum + item.cells.length, 0);
    setStatus(`${sheetIndexes.length} 個のシートで ${cellCount} 個のセルをクリアしました。`);
  } catch (error) {
    console.error(error);
    restoreGroupedSheetsAfterFailure(sheetsBefore);
    setStatus(`クリアに失敗しました：${error.message || error}`);
  }
  return true;
}



function clearSelectionContents(label = "クリア") {
  if (clearSelectionContentsForSheetGroup(label)) return;
  const range = activeSelectionRange();
  const targetRanges = activeSelectionRanges();
  if (!range || !targetRanges.length) return;
  const sheet = activeSheet();
  if (!ensureEditableRangesForProtection(sheet, targetRanges, "削除")) return;
  const sheetId = getSheetId(sheet.name);
  const cellsToClear = uniqueCellsForRanges(targetRanges);
  if (!cellsToClear.length) return;
  const before = captureCellsForHistoryRanges(sheet, targetRanges);

  try {
    cellsToClear.forEach(({ row, col }) => {
      writeCellContent(sheet, sheetId, row, col, null);
    });
    const after = captureCellsForHistoryRanges(sheet, targetRanges);
    pushHistory({
      sheetIndex: state.activeSheetIndex,
      before,
      after,
      selectionBefore: range,
      selectionAfter: range,
      selectionBeforeRanges: targetRanges,
      selectionAfterRanges: targetRanges,
      label: `${label} ${cellsToClear.length} 個のセル`,
    });
    renderSheet();
    if (state.selected) updateFormulaBarForCell(state.selected.row, state.selected.col);
    setStatus(`${cellsToClear.length} 個のセルをクリアしました。`);
  } catch (error) {
    console.error(error);
    setStatus(`クリアに失敗しました：${error.message || error}`);
    renderSheet();
  }
}



function uniqueCellsForRanges(ranges) {
  const seen = new Set();
  const cells = [];
  (ranges || []).forEach((range) => {
    forEachCellInRange(range, (row, col) => {
      const key = cellKey(row, col);
      if (seen.has(key)) return;
      seen.add(key);
      cells.push({ row, col });
    });
  });
  return cells;
}



function captureCellsForHistory(sheet, range) {
  const cells = [];
  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    cells.push({
      row,
      col,
      cell: cloneCellModel(sheetDisplayCell(sheet, row, col, sheet.cells[key] || null)),
    });
  });
  return cells;
}



function captureCellsForHistoryRanges(sheet, ranges) {
  const cells = [];
  const seen = new Set();
  (ranges || []).forEach((range) => {
    forEachCellInRange(range, (row, col) => {
      const key = cellKey(row, col);
      if (seen.has(key)) return;
      seen.add(key);
      cells.push({
        row,
        col,
        cell: cloneCellModel(sheetDisplayCell(sheet, row, col, sheet.cells[key] || null)),
      });
    });
  });
  return cells;
}



function captureSheetForHistory(sheet) {
  return {
    name: sheet.name,
    state: sheet.state || "visible",
    tabColor: sheet.tabColor || "",
    chartSheet: sheet.chartSheet === true,
    sheetKind: normalizeSheetKind(sheet.sheetKind, sheet),
    specialSheet: cloneJsonLikeValue(sheet.specialSheet),
    macroSheet: sheet.macroSheet === true,
    xl4MacroSheet: sheet.xl4MacroSheet === true,
    internationalMacroSheet: sheet.internationalMacroSheet === true,
    dialogSheet: sheet.dialogSheet === true,
    rowCount: sheet.rowCount,
    colCount: sheet.colCount,
    baseRowCount: sheet.baseRowCount,
    baseColCount: sheet.baseColCount,
    defaultRowHeight: sheet.defaultRowHeight || rowHeightToPixels(15),
    defaultColumnWidth: sheet.defaultColumnWidth || columnWidthToPixels(8.43),
    defaultCellCss: cloneJsonLikeValue(sheet.defaultCellCss || {}),
    outlineSettings: normalizeOutlineSettings(sheet.outlineSettings),
    showGridLines: sheet.showGridLines,
    showHeadings: sheet.showHeadings,
    showRuler: normalizedSheetShowRuler(sheet.showRuler, sheet.viewMode),
    showObjects: sheet.showObjects,
    showOutlineSymbols: sheet.showOutlineSymbols,
    showFormulas: sheet.showFormulas === true,
    showZeros: sheet.showZeros !== false,
    showWhiteSpace: sheet.showWhiteSpace !== false,
    rightToLeft: sheet.rightToLeft === true,
    tabSelected: sheet.tabSelected === true,
    viewMode: sheet.viewMode,
    zoomScale: sheet.zoomScale,
    baseZoomScale: sheet.baseZoomScale,
    displayScale: sheet.displayScale,
    freezePanes: cloneJsonLikeValue(sheet.freezePanes),
    splitPane: cloneJsonLikeValue(sheet.splitPane),
    sheetViews: cloneJsonLikeValue(sheet.sheetViews),
    activeSheetViewId: sheet.activeSheetViewId || null,
    codeModule: cloneJsonLikeValue(sheet.codeModule),
    formulaAudits: cloneJsonLikeValue(sheet.formulaAudits),
    pageLayout: cloneJsonLikeValue(sheet.pageLayout),
    rows: (sheet.rows || []).map((row) => ({ ...row })),
    columns: (sheet.columns || []).map((column) => ({ ...column })),
    merges: cloneMerges(sheet.merges),
    images: cloneImages(sheet.images),
    inkStrokes: cloneInkStrokes(sheet.inkStrokes),
    tables: cloneTables(sheet.tables),
    queries: cloneJsonLikeValue(sheet.queries),
    autoFilter: cloneJsonLikeValue(sheet.autoFilter),
    sortState: cloneJsonLikeValue(sheet.sortState),
    headerFooter: cloneJsonLikeValue(sheet.headerFooter),
    protection: cloneJsonLikeValue(sheet.protection),
    cells: Object.fromEntries(Object.entries(sheet.cells || {}).map(([key, cell]) => [key, cloneCellModel(cell)])),
    styleRanges: cloneJsonLikeValue(sheet.styleRanges || []),
    hfData: [],
  };
}



function restoreSheetFromHistory(sheet, snapshot) {
  if (!snapshot) return;
  if (snapshot.name != null) sheet.name = snapshot.name;
  sheet.state = snapshot.state || "visible";
  sheet.tabColor = snapshot.tabColor || "";
  sheet.chartSheet = snapshot.chartSheet === true;
  sheet.sheetKind = normalizeSheetKind(snapshot.sheetKind, snapshot);
  sheet.specialSheet = normalizeSpecialSheetDescriptor(snapshot.specialSheet, sheet);
  sheet.macroSheet = sheet.sheetKind === "macro" || sheet.sheetKind === "intlMacro" || snapshot.macroSheet === true;
  sheet.xl4MacroSheet = sheet.sheetKind === "macro" || snapshot.xl4MacroSheet === true;
  sheet.internationalMacroSheet = sheet.sheetKind === "intlMacro" || snapshot.internationalMacroSheet === true;
  sheet.dialogSheet = sheet.sheetKind === "dialog" || snapshot.dialogSheet === true;
  sheet.rowCount = snapshot.rowCount;
  sheet.colCount = snapshot.colCount;
  sheet.baseRowCount = snapshot.baseRowCount;
  sheet.baseColCount = snapshot.baseColCount;
  sheet.defaultCellCss = cloneJsonLikeValue(snapshot.defaultCellCss || {});
  sheet.outlineSettings = normalizeOutlineSettings(snapshot.outlineSettings);
  if (snapshot.defaultRowHeight != null) sheet.defaultRowHeight = snapshot.defaultRowHeight;
  if (snapshot.defaultColumnWidth != null) sheet.defaultColumnWidth = snapshot.defaultColumnWidth;
  sheet.showGridLines = snapshot.showGridLines !== false;
  sheet.showHeadings = snapshot.showHeadings !== false;
  sheet.viewMode = normalizedSheetViewMode(snapshot.viewMode || sheet.viewMode);
  sheet.showRuler = normalizedSheetShowRuler(snapshot.showRuler, sheet.viewMode);
  sheet.showObjects = snapshot.showObjects !== false;
  sheet.showOutlineSymbols = snapshot.showOutlineSymbols !== false;
  sheet.showFormulas = snapshot.showFormulas === true;
  sheet.showZeros = snapshot.showZeros !== false;
  sheet.showWhiteSpace = snapshot.showWhiteSpace !== false;
  sheet.rightToLeft = snapshot.rightToLeft === true;
  sheet.tabSelected = snapshot.tabSelected === true;
  sheet.zoomScale = Number(snapshot.zoomScale) || sheet.zoomScale || 100;
  sheet.baseZoomScale = Number(snapshot.baseZoomScale) || sheet.baseZoomScale || sheet.zoomScale || 100;
  sheet.displayScale = Number(snapshot.displayScale) || sheet.displayScale || 1;
  sheet.freezePanes = normalizePaneState(snapshot.freezePanes);
  sheet.splitPane = normalizePaneState(snapshot.splitPane);
  sheet.sheetViews = Array.isArray(snapshot.sheetViews) ? snapshot.sheetViews.map(normalizeSheetView).filter(Boolean) : [];
  sheet.activeSheetViewId = sheet.sheetViews.some((view) => view.id === snapshot.activeSheetViewId) ? snapshot.activeSheetViewId : null;
  sheet.codeModule = normalizeSheetCodeModule(snapshot.codeModule, sheet);
  sheet.formulaAudits = cloneJsonLikeValue(snapshot.formulaAudits) || { precedents: [], dependents: [] };
  sheet.pageLayout = normalizedPageLayoutSettings(snapshot.pageLayout, sheet);
  sheet.rows = (snapshot.rows || []).map((row) => ({ ...row }));
  sheet.columns = (snapshot.columns || []).map((column) => ({ ...column }));
  sheet.merges = cloneMerges(snapshot.merges);
  sheet.images = cloneImages(snapshot.images);
  sheet.inkStrokes = cloneInkStrokes(snapshot.inkStrokes);
  sheet.tables = cloneTables(snapshot.tables);
  sheet.queries = cloneJsonLikeValue(snapshot.queries);
  sheet.autoFilter = cloneJsonLikeValue(snapshot.autoFilter);
  sheet.sortState = cloneJsonLikeValue(snapshot.sortState);
  sheet.headerFooter = normalizedHeaderFooterSettings(snapshot.headerFooter, sheet.pageLayout);
  sheet.protection = cloneJsonLikeValue(snapshot.protection);
  sheet.styleRanges = cloneJsonLikeValue(snapshot.styleRanges || []);
  normalizeSheetStyleRanges(sheet);
  const cells = Object.fromEntries(Object.entries(snapshot.cells || {}).map(([key, cell]) => [key, cloneCellModel(cell)]));
  if (Array.isArray(snapshot.hfData) && snapshot.hfData.length) {
    const migratedSnapshot = { ...snapshot, cells, defaultCellCss: sheet.defaultCellCss };
    migrateDenseFormulaDataToCells(migratedSnapshot);
    sheet.cells = migratedSnapshot.cells;
  } else {
    sheet.cells = cells;
  }
  sheet.hfData = [];
  internSheetCellStyles(sheet);
}



function captureSheetsForHistory(sheetIndexes) {
  const indexes = [...new Set((sheetIndexes || []).map((index) => Number(index)).filter((index) => state.model?.sheets?.[index]))];
  return indexes.map((index) => ({
    index,
    sheet: captureSheetForHistory(state.model.sheets[index]),
  }));
}



function cwsAgentCaptureWorkbookSheetsSnapshot() {
  return (state.model?.sheets || []).map((sheet) => captureSheetForHistory(sheet));
}



function cwsAgentRestoreWorkbookSheetsSnapshot(snapshot) {
  const sheets = Array.isArray(snapshot) ? snapshot.map((sheet) => cloneSheetModel(sheet)).filter(Boolean) : [];
  state.model.sheets = sheets.length ? sheets : [createBlankSheet("Sheet1")];
}



function restoreSheetsFromHistory(snapshots) {
  (snapshots || []).forEach((snapshot) => {
    const sheet = state.model?.sheets?.[snapshot.index];
    if (sheet) restoreSheetFromHistory(sheet, snapshot.sheet);
  });
}



function cloneCellModel(cell) {
  if (!cell) return null;
  return internCellModelStyles({
    row: cell.row,
    col: cell.col,
    raw: cloneCellScalarValue(cell.raw),
    formula: cell.formula ?? null,
    cached: cloneCellScalarValue(cell.cached),
    display: cell.display ?? "",
    hyperlink: cloneJsonLikeValue(cell.hyperlink),
    comment: cloneJsonLikeValue(cell.comment),
    note: cloneJsonLikeValue(cell.note),
    noteVisible: Boolean(cell.noteVisible),
    validation: cloneJsonLikeValue(cell.validation),
    sparkline: cloneJsonLikeValue(cell.sparkline),
    quickAnalysisBar: cloneJsonLikeValue(cell.quickAnalysisBar),
    quickAnalysisIcon: cloneJsonLikeValue(cell.quickAnalysisIcon),
    quickAnalysisStyle: cloneJsonLikeValue(cell.quickAnalysisStyle),
    dataTableFormula: cell.dataTableFormula || "",
    richText: cloneJsonLikeValue(cell.richText),
    phonetics: cloneJsonLikeValue(cell.phonetics),
    showPhonetic: Boolean(cell.showPhonetic),
    css: cell.css || {},
    hasFill: Boolean(cell.hasFill),
    hasFontColor: Boolean(cell.hasFontColor),
    borders: cell.borders || {},
    numFmt: cell.numFmt || "",
  });
}



function cloneCellScalarValue(value) {
  if (value == null || typeof value !== "object") return value ?? null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return dateToExcelSerial(value);
  return cloneJsonLikeValue(value);
}



function cloneJsonLikeValue(value) {
  if (value == null || typeof value !== "object") return value ?? null;
  if (Array.isArray(value)) return value.map(cloneJsonLikeValue);
  return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, cloneJsonLikeValue(item)]));
}



function cloneMerges(merges) {
  return (merges || []).map((merge) => ({ ...merge }));
}



function cloneImages(images) {
  return (images || []).map((image) => ({
    ...image,
    tl: image.tl ? { ...image.tl } : null,
    br: image.br ? { ...image.br } : null,
    ext: image.ext ? { ...image.ext } : null,
    chart: cloneJsonLikeValue(image.chart),
    shape: cloneJsonLikeValue(image.shape),
  }));
}



function normalizeInkStrokes(strokes) {
  return (Array.isArray(strokes) ? strokes : [])
    .map((stroke, index) => {
      const points = (Array.isArray(stroke?.points) ? stroke.points : [])
        .map((point) => ({
          x: Number(point?.x),
          y: Number(point?.y),
        }))
        .filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
      if (points.length < 2) return null;
      return {
        id: String(stroke?.id || `ink-${Date.now()}-${index}`),
        name: String(stroke?.name || stroke?.label || "インク"),
        tool: String(stroke?.tool || "pen"),
        color: normalizeColorValue(stroke?.color) || stroke?.color || "#111827",
        width: Math.max(0.5, Math.min(96, Number(stroke?.width) || 2)),
        opacity: Math.max(0.05, Math.min(1, Number(stroke?.opacity ?? 1))),
        points,
      };
    })
    .filter(Boolean);
}



function cloneInkStrokes(strokes) {
  return normalizeInkStrokes(strokes).map((stroke) => ({
    ...stroke,
    points: stroke.points.map((point) => ({ ...point })),
  }));
}



function cloneTables(tables) {
  return (tables || []).map((table) => ({
    ...table,
    range: table.range ? { ...table.range } : null,
    styleInfo: table.styleInfo ? { ...table.styleInfo } : table.styleInfo,
    columns: Array.isArray(table.columns) ? table.columns.map((column) => ({ ...column })) : table.columns,
    autoFilterState: cloneJsonLikeValue(table.autoFilterState),
  }));
}



function cloneDefinedNames(definedNames) {
  return (definedNames || []).map((definition) => ({ ...definition }));
}



function materializeSheetStyleTable(sheet) {
  const table = sheet?.styleTable && typeof sheet.styleTable === "object" ? sheet.styleTable : null;
  if (!table) return sheet;
  const applyStyleId = (target) => {
    if (!target || typeof target !== "object") return;
    const styleId = String(target.styleId || "");
    const style = table[styleId];
    if (!style || typeof style !== "object") return;
    if (style.css && typeof style.css === "object") target.css = { ...cloneJsonLikeValue(style.css), ...(target.css || {}) };
    if (style.hasFill || target.hasFill) target.hasFill = true;
    if (style.hasFontColor || target.hasFontColor) target.hasFontColor = true;
    if (style.borders && typeof style.borders === "object") target.borders = { ...cloneJsonLikeValue(style.borders), ...(target.borders || {}) };
    delete target.styleId;
  };
  Object.values(sheet.cells || {}).forEach(applyStyleId);
  (sheet.styleRanges || []).forEach(applyStyleId);
  delete sheet.styleTable;
  return sheet;
}



function normalizeSheetStyleRanges(sheet) {
  const ranges = Array.isArray(sheet?.styleRanges) ? sheet.styleRanges : [];
  if (!ranges.length) {
    delete sheet.styleRanges;
    invalidateSheetStyleRangeCache(sheet);
    return sheet;
  }
  const normalizedRanges = [];
  ranges.forEach((range) => {
    const top = Math.max(1, Math.floor(Number(range?.top) || 0));
    const left = Math.max(1, Math.floor(Number(range?.left) || 0));
    const bottom = Math.max(top, Math.floor(Number(range?.bottom) || top));
    const right = Math.max(left, Math.floor(Number(range?.right) || left));
    const css = internCssObject(range.css);
    const borders = internBordersObject(range.borders);
    const normalized = { top, left, bottom, right };
    if (Object.keys(css).length) normalized.css = css;
    if (range.hasFill || css.backgroundColor) normalized.hasFill = true;
    if (range.hasFontColor) normalized.hasFontColor = true;
    if (Object.keys(borders).length) normalized.borders = borders;
    if (normalized.css || normalized.borders || normalized.hasFill || normalized.hasFontColor) normalizedRanges.push(normalized);
  });
  if (normalizedRanges.length) sheet.styleRanges = normalizedRanges;
  else delete sheet.styleRanges;
  invalidateSheetStyleRangeCache(sheet);
  return sheet;
}



function invalidateSheetStyleRangeCache(sheet) {
  if (sheet) sheetStyleRangeCacheBySheet.delete(sheet);
}



function sheetStyleRangesForRow(sheet, row) {
  const ranges = Array.isArray(sheet?.styleRanges) ? sheet.styleRanges : [];
  if (!ranges.length) return [];
  let cache = sheetStyleRangeCacheBySheet.get(sheet);
  if (!cache || cache.ranges !== ranges) {
    cache = { ranges, rows: new Map() };
    sheetStyleRangeCacheBySheet.set(sheet, cache);
  }
  const safeRow = Math.max(1, Math.floor(Number(row) || 0));
  if (!cache.rows.has(safeRow)) {
    cache.rows.set(
      safeRow,
      ranges.filter((range) => range.top <= safeRow && range.bottom >= safeRow),
    );
  }
  return cache.rows.get(safeRow) || [];
}



function sheetStyleRangeCell(sheet, row, col) {
  const safeRow = Math.max(1, Math.floor(Number(row) || 0));
  const safeCol = Math.max(1, Math.floor(Number(col) || 0));
  const ranges = sheetStyleRangesForRow(sheet, safeRow);
  if (!ranges.length) return null;
  let css = null;
  let borders = null;
  let hasFill = false;
  let hasFontColor = false;
  ranges.forEach((range) => {
    if (range.left > safeCol || range.right < safeCol) return;
    if (range.css && Object.keys(range.css).length) css = { ...(css || {}), ...range.css };
    if (range.borders && Object.keys(range.borders).length) borders = { ...(borders || {}), ...range.borders };
    if (range.hasFill || range.css?.backgroundColor) hasFill = true;
    if (range.hasFontColor) hasFontColor = true;
  });
  if (!css && !borders && !hasFill && !hasFontColor) return null;
  const cell = { row: safeRow, col: safeCol };
  if (css) cell.css = internCssObject(css);
  if (borders) cell.borders = internBordersObject(borders);
  if (hasFill || cell.css?.backgroundColor) cell.hasFill = true;
  if (hasFontColor) cell.hasFontColor = true;
  return cell;
}



function sheetDisplayCell(sheet, row, col, sourceCell = null) {
  const rangeCell = sheetStyleRangeCell(sheet, row, col);
  if (!rangeCell) return sourceCell || null;
  if (!sourceCell) return rangeCell;
  const css = { ...(rangeCell.css || {}), ...(sourceCell.css || {}) };
  const borders = { ...(rangeCell.borders || {}), ...(sourceCell.borders || {}) };
  return {
    ...rangeCell,
    ...sourceCell,
    css: internCssObject(css),
    borders: internBordersObject(borders),
    hasFill: Boolean(rangeCell.hasFill || sourceCell.hasFill || css.backgroundColor),
    hasFontColor: Boolean(rangeCell.hasFontColor || sourceCell.hasFontColor),
  };
}



function removeSheetStyleRangesAtCell(sheet, row, col) {
  const ranges = Array.isArray(sheet?.styleRanges) ? sheet.styleRanges : [];
  if (!ranges.length) return;
  const safeRow = Math.max(1, Math.floor(Number(row) || 0));
  const safeCol = Math.max(1, Math.floor(Number(col) || 0));
  let changed = false;
  const nextRanges = [];
  const appendPiece = (range, top, left, bottom, right) => {
    if (top > bottom || left > right) return;
    nextRanges.push({ ...range, top, left, bottom, right });
  };
  ranges.forEach((range) => {
    if (range.top > safeRow || range.bottom < safeRow || range.left > safeCol || range.right < safeCol) {
      nextRanges.push(range);
      return;
    }
    changed = true;
    appendPiece(range, range.top, range.left, safeRow - 1, range.right);
    appendPiece(range, safeRow + 1, range.left, range.bottom, range.right);
    appendPiece(range, safeRow, range.left, safeRow, safeCol - 1);
    appendPiece(range, safeRow, safeCol + 1, safeRow, range.right);
  });
  if (!changed) return;
  if (nextRanges.length) sheet.styleRanges = nextRanges;
  else delete sheet.styleRanges;
  invalidateSheetStyleRangeCache(sheet);
}



function compactSheetDefaultCellCss(sheet) {
  if (!sheet?.cells) return sheet;
  const existingDefaultCss = sanitizeSheetDefaultCellCss(sheet, defaultableCellCss(sheet.defaultCellCss));
  const defaultCss = Object.keys(existingDefaultCss).length ? existingDefaultCss : mostCommonDefaultCellCss(sheet);
  sheet.defaultCellCss = internCssObject(defaultCss);
  Object.keys(sheet.cells).forEach((key) => {
    const cell = sheet.cells[key];
    if (!cell) return;
    compactCellCssAgainstDefault(sheet, cell);
    if (!isPersistentCell(cell)) delete sheet.cells[key];
  });
  return sheet;
}



function sanitizeSheetDefaultCellCss(sheet, defaultCss) {
  if (!defaultCss || !Object.keys(defaultCss).length) return {};
  const localizedKeys = Object.keys(defaultCss).filter((key) =>
    LOCALIZED_DEFAULT_CELL_CSS_KEYS.has(key) && sheetDefaultCellCssKeyLooksDecorative(key, defaultCss[key]),
  );
  if (!localizedKeys.length || !sheetHasExplicitlyStyledCells(sheet)) return defaultCss;
  localizeDefaultCellCssToExplicitlyStyledCells(sheet, defaultCss, localizedKeys);
  return Object.fromEntries(Object.entries(defaultCss).filter(([key]) => !localizedKeys.includes(key)));
}



function sheetDefaultCellCssKeyLooksDecorative(key, value) {
  if (key === "color") return cellCssColorLooksLight(value);
  if (key === "fontWeight") return Boolean(value) && String(value) !== "400";
  if (key === "fontStyle") return Boolean(value) && String(value) !== "normal";
  if (key === "textDecoration") return Boolean(value) && String(value) !== "none";
  if (key === "textAlign") return ["center", "right"].includes(String(value || "").toLowerCase());
  if (key === "alignItems") return ["center", "flex-end"].includes(String(value || "").toLowerCase());
  if (key === "textRotation") return Boolean(value);
  return false;
}



function cellCssColorLooksLight(value) {
  const text = String(value || "").trim();
  if (!text) return false;
  const hex = normalizeColorValue(text);
  let channels = null;
  if (hex) {
    channels = [
      Number.parseInt(hex.slice(1, 3), 16),
      Number.parseInt(hex.slice(3, 5), 16),
      Number.parseInt(hex.slice(5, 7), 16),
    ];
  } else {
    const match = text.match(/^rgba?\(\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})/i);
    if (match) channels = [match[1], match[2], match[3]].map((item) => Number.parseInt(item, 10));
  }
  if (!channels || channels.some((channel) => !Number.isFinite(channel))) return false;
  const [red, green, blue] = channels;
  return (0.2126 * red + 0.7152 * green + 0.0722 * blue) >= 180;
}



function sheetHasExplicitlyStyledCells(sheet) {
  return Object.values(sheet?.cells || {}).some(cellLooksExplicitlyStyled);
}



function localizeDefaultCellCssToExplicitlyStyledCells(sheet, defaultCss, localizedKeys) {
  Object.values(sheet?.cells || {}).forEach((cell) => {
    if (!cellLooksExplicitlyStyled(cell)) return;
    cell.css = mutableCssObject(cell.css);
    localizedKeys.forEach((key) => {
      if (cell.css[key] == null || cell.css[key] === "") cell.css[key] = defaultCss[key];
      if (key === "color") cell.hasFontColor = true;
    });
  });
}



function cellLooksExplicitlyStyled(cell) {
  if (!cell) return false;
  const css = cell.css || {};
  return Boolean(
    cell.hasFill ||
      cell.hasFontColor ||
      css.backgroundColor ||
      css.color ||
      Object.keys(cell.borders || {}).length,
  );
}



function mostCommonDefaultCellCss(sheet) {
  const counts = new Map();
  const values = new Map();
  Object.values(sheet?.cells || {}).forEach((cell) => {
    if (cellLooksExplicitlyStyled(cell)) return;
    const candidate = inferredDefaultableCellCss(cell?.css);
    if (!Object.keys(candidate).length) return;
    const key = stableObjectKey(candidate);
    counts.set(key, (counts.get(key) || 0) + 1);
    values.set(key, candidate);
  });
  let bestKey = "";
  let bestCount = 0;
  counts.forEach((count, key) => {
    if (count > bestCount) {
      bestKey = key;
      bestCount = count;
    }
  });
  return bestKey && bestCount >= 2 ? values.get(bestKey) || {} : {};
}



function inferredDefaultableCellCss(css) {
  const candidate = {};
  INFERRED_DEFAULT_CELL_CSS_KEYS.forEach((key) => {
    if (css?.[key] != null && css[key] !== "") candidate[key] = css[key];
  });
  return candidate;
}



function defaultableCellCss(css) {
  const candidate = {};
  DEFAULT_CELL_CSS_KEYS.forEach((key) => {
    if (css?.[key] != null && css[key] !== "") candidate[key] = css[key];
  });
  return candidate;
}



function compactCellCssAgainstDefault(sheet, cell) {
  if (!cell?.css || !sheet?.defaultCellCss) return cell;
  if (Object.isFrozen(cell.css)) cell.css = { ...cell.css };
  Object.entries(sheet.defaultCellCss).forEach(([key, value]) => {
    if (key === "color" && cellHasFontColor(cell)) return;
    if (cell.css?.[key] === value) delete cell.css[key];
  });
  return cell;
}



function stableObjectKey(value) {
  if (!value || typeof value !== "object") return JSON.stringify(value ?? null);
  if (Array.isArray(value)) return `[${value.map((item) => stableObjectKey(item)).join(",")}]`;
  return `{${Object.keys(value)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${stableObjectKey(value[key])}`)
    .join(",")}}`;
}



function internPooledObject(pool, value, emptyObject = EMPTY_CELL_CSS) {
  if (!value || typeof value !== "object") return emptyObject;
  const entries = Object.entries(value).filter(([, item]) => item !== undefined && item !== null && item !== "");
  if (!entries.length) return emptyObject;
  const normalized = {};
  entries
    .sort(([left], [right]) => left.localeCompare(right))
    .forEach(([key, item]) => {
      normalized[key] = item;
    });
  const key = stableObjectKey(normalized);
  const existing = pool.get(key);
  if (existing) return existing;
  const frozen = Object.freeze(normalized);
  if (pool.size < STYLE_INTERN_POOL_LIMIT) pool.set(key, frozen);
  return frozen;
}



function internCssObject(css) {
  return internPooledObject(cellStyleInternPools.css, css, EMPTY_CELL_CSS);
}



function internBorderSide(border) {
  if (!border || typeof border !== "object") return null;
  return internPooledObject(cellStyleInternPools.borderSide, border, null);
}



function internBordersObject(borders) {
  if (!borders || typeof borders !== "object") return EMPTY_CELL_BORDERS;
  const normalized = {};
  CELL_BORDER_SIDES.forEach((side) => {
    const border = internBorderSide(borders[side]);
    if (border) normalized[side] = border;
  });
  return internPooledObject(cellStyleInternPools.borders, normalized, EMPTY_CELL_BORDERS);
}



function internCellModelStyles(cell) {
  if (!cell) return null;
  return compactCellModelForStorage(cell);
}



function internCellModelStylesForSheet(sheet, cell) {
  if (!cell) return null;
  compactCellCssAgainstDefault(sheet, cell);
  return internCellModelStyles(cell);
}



function storeCellModelForSheet(sheet, key, cell) {
  const previousFormula = sheet?.cells?.[key]?.formula || null;
  const nextCell = internCellModelStylesForSheet(sheet, cell);
  invalidateSheetUsedRangeCache(sheet);
  if (!isPersistentCell(nextCell)) {
    if (previousFormula) invalidateFormulaSpillAnchorCache(sheet);
    delete sheet.cells[key];
    return null;
  }
  const nextFormula = nextCell?.formula || null;
  if (previousFormula !== nextFormula) invalidateFormulaSpillAnchorCache(sheet);
  sheet.cells[key] = nextCell;
  return nextCell;
}



function internSheetCellStyles(sheet) {
  sheet.defaultCellCss = internCssObject(sheet?.defaultCellCss);
  Object.keys(sheet?.cells || {}).forEach((key) => {
    storeCellModelForSheet(sheet, key, sheet.cells[key]);
  });
  return sheet;
}



function internWorkbookCellStyles(model) {
  (model?.sheets || []).forEach(internSheetCellStyles);
  return model;
}



function compactCellModelForStorage(cell) {
  if (!cell) return null;
  const compact = {
    row: Number(cell.row) || 1,
    col: Number(cell.col) || 1,
  };
  const raw = cloneCellScalarValue(cell.raw);
  const formula = cell.formula || null;
  const cached = cloneCellScalarValue(cell.cached);
  const numFmt = cell.numFmt || "";
  const storedRaw = raw != null && raw !== formula ? raw : null;
  const storedCached = cached != null && cached !== raw && cached !== formula ? cached : null;

  if (storedRaw != null) compact.raw = storedRaw;
  if (formula) compact.formula = formula;
  if (storedCached != null) compact.cached = storedCached;
  if (numFmt) compact.numFmt = numFmt;

  const display = cellDisplayForStorage(cell, storedRaw, formula, storedCached, numFmt);
  if (display && display !== implicitCellDisplayText({ raw: storedRaw, formula, cached: storedCached, numFmt })) compact.display = display;
  if (cell.hyperlink) compact.hyperlink = cloneJsonLikeValue(cell.hyperlink);
  if (cell.comment) compact.comment = cloneJsonLikeValue(cell.comment);
  if (cell.note) compact.note = cloneJsonLikeValue(cell.note);
  if (cell.noteVisible) compact.noteVisible = true;
  if (cell.validation) compact.validation = cloneJsonLikeValue(cell.validation);
  if (cell.sparkline) compact.sparkline = cloneJsonLikeValue(cell.sparkline);
  if (cell.quickAnalysisBar) compact.quickAnalysisBar = cloneJsonLikeValue(cell.quickAnalysisBar);
  if (cell.quickAnalysisIcon) compact.quickAnalysisIcon = cloneJsonLikeValue(cell.quickAnalysisIcon);
  if (cell.quickAnalysisStyle) compact.quickAnalysisStyle = cloneJsonLikeValue(cell.quickAnalysisStyle);
  if (cell.dataTableFormula) compact.dataTableFormula = String(cell.dataTableFormula);
  if (cell.richText?.length) compact.richText = cloneJsonLikeValue(cell.richText);
  if (cell.phonetics?.length) compact.phonetics = cloneJsonLikeValue(cell.phonetics);
  if (cell.showPhonetic) compact.showPhonetic = true;

  const css = internCssObject(cell.css);
  if (Object.keys(css).length) compact.css = css;
  if (cell.hasFill || css.backgroundColor) compact.hasFill = true;
  if (cell.hasFontColor) compact.hasFontColor = true;
  const borders = internBordersObject(cell.borders);
  if (Object.keys(borders).length) compact.borders = borders;
  if (cell.protection && (cell.protection.locked === false || cell.protection.hidden === true)) {
    compact.protection = {
      locked: cell.protection.locked !== false,
      hidden: Boolean(cell.protection.hidden),
    };
  }
  return compact;
}



function cellDisplayForStorage(cell, raw, formula, cached, numFmt) {
  const displayCell = {
    ...cell,
    raw,
    formula,
    cached,
    numFmt,
  };
  return normalizedCellDisplayText(displayCell) || "";
}



function internCellsInRange(sheet, range) {
  if (!range) return;
  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    if (sheet.cells?.[key]) storeCellModelForSheet(sheet, key, sheet.cells[key]);
  });
}



function internCellsInRanges(sheet, ranges) {
  const seen = new Set();
  (ranges || []).forEach((range) => {
    if (!range) return;
    forEachCellInRange(range, (row, col) => {
      const key = cellKey(row, col);
      if (seen.has(key)) return;
      seen.add(key);
      if (sheet.cells?.[key]) storeCellModelForSheet(sheet, key, sheet.cells[key]);
    });
  });
}



function cellEffectiveCss(sheet, cell) {
  const defaultCss = sheet?.defaultCellCss || EMPTY_CELL_CSS;
  const cellCss = cell?.css || EMPTY_CELL_CSS;
  if (!Object.keys(defaultCss).length) return cellCss;
  if (!Object.keys(cellCss).length) return defaultCss;
  return { ...defaultCss, ...cellCss };
}



function mutableCssObject(css) {
  if (!css || typeof css !== "object") return {};
  return Object.isFrozen(css) ? { ...css } : css;
}



function mutableBordersObject(borders) {
  if (!borders || typeof borders !== "object") return {};
  const needsClone = Object.isFrozen(borders) || CELL_BORDER_SIDES.some((side) => borders[side] && Object.isFrozen(borders[side]));
  if (!needsClone) return borders;
  const next = {};
  CELL_BORDER_SIDES.forEach((side) => {
    if (borders[side]) next[side] = { ...borders[side] };
  });
  return next;
}



function workbookMemoryStats(model = state.model) {
  const cssRefs = new Set();
  const borderRefs = new Set();
  let cells = 0;
  let styledCells = 0;
  let borderedCells = 0;
  let storedDisplayCells = 0;
  let redundantDisplayCells = 0;
  let emptyCssFieldCells = 0;
  let emptyBorderFieldCells = 0;
  let styleRanges = 0;
  let styleRangeCells = 0;
  let denseFormulaRows = 0;
  let denseFormulaCells = 0;
  const clipboard = state.internalClipboard;
  const clipboardSlots = clipboard ? (Number(clipboard.rowCount) || 0) * (Number(clipboard.colCount) || 0) : 0;
  const clipboardStoredCells = clipboard
    ? Array.isArray(clipboard.cellEntries)
      ? clipboard.cellEntries.length
      : (clipboard.cells || []).reduce((total, row) => total + (row || []).filter(Boolean).length, 0)
    : 0;
  const clipboardDenseSlots = clipboard
    ? (clipboard.cells || []).reduce((total, row) => total + (Array.isArray(row) ? row.length : 0), 0)
    : 0;
  (model?.sheets || []).forEach((sheet) => {
    (sheet.styleRanges || []).forEach((range) => {
      styleRanges += 1;
      styleRangeCells += Math.max(0, Number(range.bottom) - Number(range.top) + 1) *
        Math.max(0, Number(range.right) - Number(range.left) + 1);
    });
    (sheet.hfData || []).forEach((row) => {
      denseFormulaRows += 1;
      denseFormulaCells += Array.isArray(row) ? row.length : 0;
    });
    Object.values(sheet.cells || {}).forEach((cell) => {
      cells += 1;
      const cssKeys = Object.keys(cell?.css || {});
      const borderKeys = Object.keys(cell?.borders || {});
      if (Object.prototype.hasOwnProperty.call(cell, "display")) {
        storedDisplayCells += 1;
        if (cell.display && cell.display === implicitCellDisplayText(cell)) redundantDisplayCells += 1;
      }
      if (cell?.css && !cssKeys.length) emptyCssFieldCells += 1;
      if (cell?.borders && !borderKeys.length) emptyBorderFieldCells += 1;
      if (cssKeys.length) {
        styledCells += 1;
        cssRefs.add(cell.css);
      }
      if (borderKeys.length) {
        borderedCells += 1;
        borderRefs.add(cell.borders);
      }
    });
  });
  return {
    cells,
    styledCells,
    uniqueCssObjects: cssRefs.size,
    borderedCells,
    uniqueBorderObjects: borderRefs.size,
    storedDisplayCells,
    redundantDisplayCells,
    emptyCssFieldCells,
    emptyBorderFieldCells,
    styleRanges,
    styleRangeCells,
    denseFormulaRows,
    denseFormulaCells,
    clipboardSlots,
    clipboardStoredCells,
    clipboardDenseSlots,
    cssPoolSize: cellStyleInternPools.css.size,
    borderPoolSize: cellStyleInternPools.borders.size,
    borderSidePoolSize: cellStyleInternPools.borderSide.size,
  };
}



function pushHistory(entry) {
  if (state.isRestoringHistory || !hasHistoryPayload(entry)) return;
  markWorkbookChanged();
  state.undoStack.push(entry);
  if (state.undoStack.length > HISTORY_LIMIT) {
    state.undoStack.splice(0, state.undoStack.length - HISTORY_LIMIT);
  }
  state.redoStack = [];
  scheduleWorkbookPersistence();
}



function pushSheetInsertHistory(index, sheet, activeSheetBefore, activeSheetAfter, label, extras = {}) {
  pushHistory({
    sheetIndex: activeSheetAfter,
    sheetInsert: {
      index,
      sheet: cloneSheetModel(sheet),
    },
    activeSheetBefore,
    activeSheetAfter,
    label,
    ...extras,
  });
}



function pushSheetInsertManyHistory(index, sheets, activeSheetBefore, activeSheetAfter, label, extras = {}) {
  pushHistory({
    sheetIndex: activeSheetAfter,
    sheetInsertMany: {
      index,
      sheets: (sheets || []).map(cloneSheetModel),
    },
    activeSheetBefore,
    activeSheetAfter,
    label,
    ...extras,
  });
}



function pushSheetDeleteHistory(index, sheet, activeSheetBefore, activeSheetAfter, label, extras = {}) {
  pushHistory({
    sheetIndex: activeSheetAfter,
    sheetDelete: {
      index,
      sheet: cloneSheetModel(sheet),
    },
    activeSheetBefore,
    activeSheetAfter,
    label,
    ...extras,
  });
}



function pushSheetDeleteManyHistory(deletedSheets, activeSheetBefore, activeSheetAfter, label, extras = {}) {
  pushHistory({
    sheetIndex: activeSheetAfter,
    sheetDeleteMany: (deletedSheets || []).map((item) => ({
      index: item.index,
      sheet: cloneSheetModel(item.sheet),
    })),
    activeSheetBefore,
    activeSheetAfter,
    label,
    ...extras,
  });
}



function pushSheetMoveHistory(fromIndex, toIndex, activeSheetBefore, activeSheetAfter, label, extras = {}) {
  pushHistory({
    sheetIndex: activeSheetAfter,
    sheetMove: {
      fromIndex,
      toIndex,
    },
    activeSheetBefore,
    activeSheetAfter,
    label,
    ...extras,
  });
}



function pushSheetMoveManyHistory(beforeOrder, afterOrder, selectedSheetNames, activeSheetBefore, activeSheetAfter, label, extras = {}) {
  pushHistory({
    sheetIndex: activeSheetAfter,
    sheetMoveMany: {
      beforeOrder: [...(beforeOrder || [])],
      afterOrder: [...(afterOrder || [])],
      selectedSheetNames: [...(selectedSheetNames || [])],
    },
    activeSheetBefore,
    activeSheetAfter,
    label,
    ...extras,
  });
}



function hasHistoryPayload(entry) {
  return Boolean(
    entry &&
      ((entry.before && entry.before.length) ||
        entry.mergesBefore ||
        entry.mergesAfter ||
        entry.dimensionsBefore ||
        entry.dimensionsAfter ||
        entry.zoomBefore ||
        entry.zoomAfter ||
        entry.definedNamesBefore ||
        entry.definedNamesAfter ||
        entry.scenariosBefore ||
        entry.scenariosAfter ||
        entry.relationshipsBefore ||
        entry.relationshipsAfter ||
        entry.dataModelBefore ||
        entry.dataModelAfter ||
        entry.workbookThemeBefore ||
        entry.workbookThemeAfter ||
        entry.workbookViewBefore ||
        entry.workbookViewAfter ||
        entry.macroStateBefore ||
        entry.macroStateAfter ||
        entry.workbookSheetsBefore ||
        entry.workbookSheetsAfter ||
        (entry.sheetsBefore && entry.sheetsBefore.length) ||
        (entry.sheetsAfter && entry.sheetsAfter.length) ||
        entry.sheetBefore ||
        entry.sheetAfter ||
        entry.sheetInsert ||
        entry.sheetInsertMany ||
        entry.sheetDelete ||
        (entry.sheetDeleteMany && entry.sheetDeleteMany.length) ||
        entry.sheetMove ||
        entry.sheetMoveMany),
  );
}



function undoLastOperation() {
  const entry = state.undoStack.pop();
  if (!entry) {
    setStatus("元に戻す操作はありません。");
    return;
  }
  applyHistoryEntry(entry, "undo");
  markWorkbookChanged();
  state.redoStack.push(entry);
  scheduleWorkbookPersistence();
  setStatus(`元に戻しました：${entry.label}`);
}



function redoLastOperation() {
  const entry = state.redoStack.pop();
  if (!entry) {
    setStatus("やり直す操作はありません。");
    return;
  }
  applyHistoryEntry(entry, "redo");
  markWorkbookChanged();
  state.undoStack.push(entry);
  scheduleWorkbookPersistence();
  setStatus(`やり直しました：${entry.label}`);
}



function applyHistoryEntry(entry, direction) {
  if (entry.sheetInsert) {
    applySheetInsertHistoryEntry(entry, direction);
    return;
  }
  if (entry.sheetInsertMany) {
    applySheetInsertManyHistoryEntry(entry, direction);
    return;
  }
  if (entry.sheetDelete) {
    applySheetDeleteHistoryEntry(entry, direction);
    return;
  }
  if (entry.sheetDeleteMany) {
    applySheetDeleteManyHistoryEntry(entry, direction);
    return;
  }
  if (entry.sheetMove) {
    applySheetMoveHistoryEntry(entry, direction);
    return;
  }
  if (entry.sheetMoveMany) {
    applySheetMoveManyHistoryEntry(entry, direction);
    return;
  }
  if (entry.workbookSheetsBefore || entry.workbookSheetsAfter) {
    applyWorkbookSheetsHistoryEntry(entry, direction);
    return;
  }
  if (!state.model?.sheets[entry.sheetIndex]) return;
  state.isRestoringHistory = true;
  try {
    state.activeSheetIndex = entry.sheetIndex;
    const sheet = activeSheet();
    let restoredWorkbookView = null;
    let restoredWorkbookWindowRecord = null;
    const sheetSnapshots = direction === "undo" ? entry.sheetsBefore : entry.sheetsAfter;
    const sheetSnapshot = direction === "undo" ? entry.sheetBefore : entry.sheetAfter;
	    if (sheetSnapshots?.length) {
	      restoreSheetsFromHistory(sheetSnapshots);
	    } else if (sheetSnapshot) {
	      restoreSheetFromHistory(sheet, sheetSnapshot);
	    } else {
      const cells = direction === "undo" ? entry.before : entry.after;
      if (cells?.length) {
        restoreCellsFromHistory(sheet, cells);
      }
      const merges = direction === "undo" ? entry.mergesBefore : entry.mergesAfter;
      if (merges) {
        sheet.merges = cloneMerges(merges);
      }
      const dimensions = direction === "undo" ? entry.dimensionsBefore : entry.dimensionsAfter;
      if (dimensions) {
        restoreDimensionsFromHistory(sheet, dimensions);
      }
      const zoom = direction === "undo" ? entry.zoomBefore : entry.zoomAfter;
      if (zoom) {
        restoreZoomFromHistory(sheet, zoom);
      }
      const definedNames = direction === "undo" ? entry.definedNamesBefore : entry.definedNamesAfter;
      if (definedNames) {
        state.model.definedNames = cloneDefinedNames(definedNames);
      }
      const scenarios = direction === "undo" ? entry.scenariosBefore : entry.scenariosAfter;
      if (scenarios) {
        state.model.scenarios = cloneScenarios(scenarios);
      }
      const relationships = direction === "undo" ? entry.relationshipsBefore : entry.relationshipsAfter;
      if (relationships) {
        state.model.relationships = cloneRelationships(relationships);
      }
      const dataModel = direction === "undo" ? entry.dataModelBefore : entry.dataModelAfter;
      if (dataModel) {
        restoreWorkbookDataModel(dataModel);
      }
      const workbookView = direction === "undo" ? entry.workbookViewBefore : entry.workbookViewAfter;
      if (workbookView) {
        restoredWorkbookWindowRecord = restoreWorkbookViewState(workbookView);
        restoredWorkbookView = workbookView;
      }
      const macroState = direction === "undo" ? entry.macroStateBefore : entry.macroStateAfter;
	      if (macroState) {
	        restoreWorkbookMacroState(macroState);
	      }
	    }
	    if (sheetSnapshots?.length || sheetSnapshot) {
      const definedNames = direction === "undo" ? entry.definedNamesBefore : entry.definedNamesAfter;
      if (definedNames) {
        state.model.definedNames = cloneDefinedNames(definedNames);
      }
		      const workbookView = direction === "undo" ? entry.workbookViewBefore : entry.workbookViewAfter;
		      if (workbookView) {
		        restoredWorkbookWindowRecord = restoreWorkbookViewState(workbookView);
		        restoredWorkbookView = workbookView;
		      }
	      const macroState = direction === "undo" ? entry.macroStateBefore : entry.macroStateAfter;
      if (macroState) {
        restoreWorkbookMacroState(macroState);
      }
    }
    const workbookTheme = direction === "undo" ? entry.workbookThemeBefore : entry.workbookThemeAfter;
    if (workbookTheme) {
      restoreWorkbookThemeState(workbookTheme);
    }
	    state.hf = buildFormulaEngine(state.model);
    const selection = direction === "undo" ? entry.selectionBefore : entry.selectionAfter;
    const selectionRanges = direction === "undo" ? entry.selectionBeforeRanges : entry.selectionAfterRanges;
    const clampHistorySheetIndex = (value, fallback = entry.sheetIndex) => Math.max(0, Math.min(
      Number.isFinite(Number(value)) ? Number(value) : fallback,
      state.model.sheets.length - 1,
    ));
    const restoredActiveSheetIndex = clampHistorySheetIndex(restoredWorkbookView?.activeSheetIndex, entry.sheetIndex);
    const explicitActiveSheetIndex = direction === "undo" ? entry.activeSheetBefore : entry.activeSheetAfter;
    const nextActiveSheetIndex = clampHistorySheetIndex(explicitActiveSheetIndex, restoredActiveSheetIndex);
    if (!restoredWorkbookWindowRecord) state.activeSheetIndex = nextActiveSheetIndex;
    const selectionRangeSheetIndex = selectionRanges?.find((range) => Number.isFinite(Number(range?.sheetIndex)))?.sheetIndex;
    const selectionSheetIndex = clampHistorySheetIndex(selection?.sheetIndex ?? selectionRangeSheetIndex, state.activeSheetIndex);
    if (!restoredWorkbookWindowRecord) {
      state.selected = selection
        ? { sheetIndex: selectionSheetIndex, row: selection.top, col: selection.left }
        : null;
      state.selectionRange = selection ? { ...selection, sheetIndex: selectionSheetIndex } : null;
      state.selectionRanges = selectionRanges?.length
        ? selectionRanges.map((range) => ({ ...range, sheetIndex: clampHistorySheetIndex(range.sheetIndex, selectionSheetIndex) }))
        : state.selectionRange
          ? [state.selectionRange]
          : [];
      state.selectionStart = state.selected;
    }
    state.selectionDrag = null;
    state.selectedImage = null;
    state.headerSizeDialog = null;
    state.nameDialog = null;
    state.linkDialog = null;
    state.symbolDialog = null;
    state.annotationEditor = null;
    document.querySelector("#headerSizeDialogOverlay")?.remove();
    document.querySelector("#nameDialogOverlay")?.remove();
    document.querySelector("#linkDialogOverlay")?.remove();
    document.querySelector("#symbolDialogOverlay")?.remove();
    closeGoToDialog();
	    closeSpecialDialog();
	    renderWorkbook();
	    if (restoredWorkbookWindowRecord) restoreWorkbookWindowScrollStable(restoredWorkbookWindowRecord);
	    refreshDataDialogAfterHistory(entry);
    refreshQueriesConnectionsPane();
    scheduleQueryRefreshTimers();
    updateSelectionUi();
    if (state.selected) {
      updateFormulaBarForCell(state.selected.row, state.selected.col);
    }
  } finally {
    state.isRestoringHistory = false;
  }
}



function applyWorkbookSheetsHistoryEntry(entry, direction) {
  if (!state.model) return;
  state.isRestoringHistory = true;
  try {
    const snapshot = direction === "undo" ? entry.workbookSheetsBefore : entry.workbookSheetsAfter;
    cwsAgentRestoreWorkbookSheetsSnapshot(snapshot);
    state.activeSheetIndex = Math.max(0, Math.min(
      Number(direction === "undo" ? entry.activeSheetBefore : entry.activeSheetAfter) || 0,
      state.model.sheets.length - 1,
    ));
    finishSheetStructureHistoryRestore();
    const selection = direction === "undo" ? entry.selectionBefore : entry.selectionAfter;
    const selectionRanges = direction === "undo" ? entry.selectionBeforeRanges : entry.selectionAfterRanges;
    cwsAgentRestoreSelectionSnapshot(selection, selectionRanges, state.activeSheetIndex);
    renderWorkbook();
    updateSelectionUi();
    updateFormulaBarForSelection();
  } finally {
    state.isRestoringHistory = false;
  }
}



function refreshDataDialogAfterHistory(entry) {
  const openWorkbookLinksDialogElement = document.querySelector("[data-data-dialog='workbook-links']");
  const openCustomViewsDialogElement = document.querySelector("[data-data-dialog='custom-views']");
  const openMacrosDialogElement = document.querySelector("[data-data-dialog='macros']");
  const openMacroEditorDialogElement = document.querySelector("[data-data-dialog='macro-editor']");
  const openMacroOptionsDialogElement = document.querySelector("[data-data-dialog='macro-options']");
  const openMacroSecurityDialogElement = document.querySelector("[data-data-dialog='macro-security']");
  if (
    !entry?.scenariosBefore &&
    !entry?.scenariosAfter &&
    !entry?.relationshipsBefore &&
    !entry?.relationshipsAfter &&
    !entry?.dataModelBefore &&
    !entry?.dataModelAfter &&
    !entry?.workbookViewBefore &&
    !entry?.workbookViewAfter &&
    !entry?.macroStateBefore &&
    !entry?.macroStateAfter &&
    !openWorkbookLinksDialogElement
  ) return;
  if (openCustomViewsDialogElement) {
    const selectedName = openCustomViewsDialogElement.querySelector("[data-data-field='customViewIndex'] option:checked")?.textContent?.trim() || "";
    openCustomViewsDialog({ selectedName });
  }
  const openScenarioDialog = document.querySelector(
    "[data-data-dialog='scenario-manager'], [data-data-dialog='scenario-edit'], [data-data-dialog='scenario-summary']",
  );
  if (openScenarioDialog) openScenarioManagerDialog();
  const openRelationshipsDialogElement = document.querySelector("[data-data-dialog='relationships']");
  if (openRelationshipsDialogElement) openRelationshipsDialog();
  if (openWorkbookLinksDialogElement) openWorkbookLinksDialog(state.dataDialog?.selectedId || "");
  if (openMacrosDialogElement) {
    const selectedName = openMacrosDialogElement.querySelector("[data-data-field='macroIndex'] option:checked")?.textContent?.trim() || "";
    const macros = Array.isArray(state.model?.macros) ? state.model.macros : [];
    const selectedIndex = macros.findIndex((macro) => (macro.name || "").trim() === selectedName);
    openMacrosDialog({ selectedIndex: selectedIndex >= 0 ? selectedIndex : Math.min(0, macros.length - 1) });
  }
  if (openMacroEditorDialogElement || openMacroOptionsDialogElement) {
    const index = Number((openMacroEditorDialogElement || openMacroOptionsDialogElement)?.dataset?.macroIndex);
    openMacrosDialog({ selectedIndex: Number.isFinite(index) ? index : 0 });
  }
  if (openMacroSecurityDialogElement) openMacroSecurityDialog();
}



function applySheetInsertHistoryEntry(entry, direction) {
  if (!state.model) return;
  state.isRestoringHistory = true;
  try {
    const insert = entry.sheetInsert;
    const index = Math.max(0, Math.min(Number(insert.index) || 0, state.model.sheets.length));
    if (direction === "undo") {
      if (state.model.sheets[index]) state.model.sheets.splice(index, 1);
      if (!state.model.sheets.length) state.model.sheets.push(createBlankSheet("Sheet1"));
      if (entry.sheetsBefore?.length) restoreSheetsFromHistory(entry.sheetsBefore);
      if (entry.dataModelBefore) restoreWorkbookDataModel(entry.dataModelBefore);
      state.activeSheetIndex = Math.max(0, Math.min(Number(entry.activeSheetBefore) || 0, state.model.sheets.length - 1));
    } else {
      state.model.sheets.splice(index, 0, cloneSheetModel(insert.sheet));
      if (entry.sheetsAfter?.length) restoreSheetsFromHistory(entry.sheetsAfter);
      if (entry.dataModelAfter) restoreWorkbookDataModel(entry.dataModelAfter);
      state.activeSheetIndex = Math.max(0, Math.min(Number(entry.activeSheetAfter) || index, state.model.sheets.length - 1));
    }
    state.selectedSheetIndexes = [];
    state.selected = null;
    state.selectionRange = null;
    state.selectionRanges = [];
    state.selectionStart = null;
    state.selectionAnchor = null;
    state.selectionDrag = null;
    state.selectedImage = null;
    state.headerSizeDialog = null;
    state.nameDialog = null;
    state.linkDialog = null;
    state.symbolDialog = null;
    state.annotationEditor = null;
    ensureActiveSheetVisible();
    state.hf = buildFormulaEngine(state.model);
    $sheetHost.scrollTop(0);
    $sheetHost.scrollLeft(0);
    $horizontalScrollbar.scrollLeft(0);
    renderWorkbook();
    refreshQueriesConnectionsPane();
    updateSelectionUi();
    updateFormulaBarForSelection();
  } finally {
    state.isRestoringHistory = false;
  }
}



function applySheetInsertManyHistoryEntry(entry, direction) {
  if (!state.model) return;
  state.isRestoringHistory = true;
  try {
    const insert = entry.sheetInsertMany;
    const sheets = Array.isArray(insert?.sheets) ? insert.sheets : [];
    const index = Math.max(0, Math.min(Number(insert?.index) || 0, state.model.sheets.length));
    if (direction === "undo") {
      const names = new Set(sheets.map((sheet) => sheet.name));
      for (let offset = sheets.length - 1; offset >= 0; offset -= 1) {
        const expectedIndex = index + offset;
        const expectedName = sheets[offset]?.name || "";
        const fallbackIndex = state.model.sheets.findIndex((sheet) => names.has(sheet.name));
        const removeIndex = state.model.sheets[expectedIndex]?.name === expectedName ? expectedIndex : fallbackIndex;
        if (removeIndex >= 0) state.model.sheets.splice(removeIndex, 1);
      }
      if (!state.model.sheets.length) state.model.sheets.push(createBlankSheet("Sheet1"));
      state.activeSheetIndex = Math.max(0, Math.min(Number(entry.activeSheetBefore) || 0, state.model.sheets.length - 1));
      finishSheetStructureHistoryRestore();
    } else {
      state.model.sheets.splice(index, 0, ...sheets.map(cloneSheetModel));
      state.activeSheetIndex = Math.max(0, Math.min(Number(entry.activeSheetAfter) || index, state.model.sheets.length - 1));
      finishSheetStructureHistoryRestore({ selectedSheetNames: sheets.map((sheet) => sheet.name) });
    }
  } finally {
    state.isRestoringHistory = false;
  }
}



function applySheetDeleteHistoryEntry(entry, direction) {
  if (!state.model) return;
  state.isRestoringHistory = true;
  try {
    const deleted = entry.sheetDelete;
    const index = Math.max(0, Math.min(Number(deleted.index) || 0, state.model.sheets.length));
    if (direction === "undo") {
      state.model.sheets.splice(index, 0, cloneSheetModel(deleted.sheet));
      state.activeSheetIndex = Math.max(0, Math.min(Number(entry.activeSheetBefore) || index, state.model.sheets.length - 1));
    } else {
      if (state.model.sheets[index]) state.model.sheets.splice(index, 1);
      if (!state.model.sheets.length) state.model.sheets.push(createBlankSheet("Sheet1"));
      state.activeSheetIndex = Math.max(0, Math.min(Number(entry.activeSheetAfter) || 0, state.model.sheets.length - 1));
    }
    finishSheetStructureHistoryRestore();
  } finally {
    state.isRestoringHistory = false;
  }
}



function applySheetDeleteManyHistoryEntry(entry, direction) {
  if (!state.model) return;
  state.isRestoringHistory = true;
  try {
    const deleted = Array.isArray(entry.sheetDeleteMany) ? entry.sheetDeleteMany : [];
    if (direction === "undo") {
      deleted
        .slice()
        .sort((left, right) => left.index - right.index)
        .forEach((item) => {
          const index = Math.max(0, Math.min(Number(item.index) || 0, state.model.sheets.length));
          state.model.sheets.splice(index, 0, cloneSheetModel(item.sheet));
        });
      state.activeSheetIndex = Math.max(0, Math.min(Number(entry.activeSheetBefore) || 0, state.model.sheets.length - 1));
    } else {
      deleted
        .slice()
        .sort((left, right) => right.index - left.index)
        .forEach((item) => {
          const expectedName = item.sheet?.name || "";
          const fallbackIndex = state.model.sheets.findIndex((sheet) => sheet.name === expectedName);
          const index = state.model.sheets[Number(item.index)]?.name === expectedName ? Number(item.index) : fallbackIndex;
          if (index >= 0) state.model.sheets.splice(index, 1);
        });
      if (!state.model.sheets.length) state.model.sheets.push(createBlankSheet("Sheet1"));
      state.activeSheetIndex = Math.max(0, Math.min(Number(entry.activeSheetAfter) || 0, state.model.sheets.length - 1));
    }
    finishSheetStructureHistoryRestore();
  } finally {
    state.isRestoringHistory = false;
  }
}



function applySheetMoveHistoryEntry(entry, direction) {
  if (!state.model?.sheets?.length) return;
  state.isRestoringHistory = true;
  try {
    const move = entry.sheetMove;
    const fromIndex = direction === "undo" ? move.toIndex : move.fromIndex;
    const toIndex = direction === "undo" ? move.fromIndex : move.toIndex;
    const actualIndex = moveSheetToIndex(fromIndex, toIndex);
    state.activeSheetIndex = Math.max(
      0,
      Math.min(Number(direction === "undo" ? entry.activeSheetBefore : entry.activeSheetAfter) || actualIndex, state.model.sheets.length - 1),
    );
    finishSheetStructureHistoryRestore();
  } finally {
    state.isRestoringHistory = false;
  }
}



function applySheetMoveManyHistoryEntry(entry, direction) {
  if (!state.model?.sheets?.length) return;
  state.isRestoringHistory = true;
  try {
    const move = entry.sheetMoveMany;
    const order = direction === "undo" ? move.beforeOrder : move.afterOrder;
    reorderSheetsByName(order);
    state.activeSheetIndex = Math.max(
      0,
      Math.min(Number(direction === "undo" ? entry.activeSheetBefore : entry.activeSheetAfter) || 0, state.model.sheets.length - 1),
    );
    finishSheetStructureHistoryRestore({ selectedSheetNames: direction === "redo" ? move.selectedSheetNames : [] });
  } finally {
    state.isRestoringHistory = false;
  }
}



function reorderSheetsByName(order) {
  if (!state.model?.sheets?.length || !Array.isArray(order)) return;
  const byName = new Map(state.model.sheets.map((sheet) => [sheet.name, sheet]));
  const used = new Set();
  const reordered = [];
  order.forEach((name) => {
    const sheet = byName.get(name);
    if (!sheet || used.has(sheet)) return;
    reordered.push(sheet);
    used.add(sheet);
  });
  state.model.sheets.forEach((sheet) => {
    if (!used.has(sheet)) reordered.push(sheet);
  });
  if (reordered.length) state.model.sheets = reordered;
}

export {
  applyCellStructureOperation,
  applyCellStructureOperationToSheetGroup,
  applyClipboardDimensionsToRange,
  applyDeleteOperation,
  applyHistoryEntry,
  applyInsertOperation,
  applySheetDeleteHistoryEntry,
  applySheetDeleteManyHistoryEntry,
  applySheetInsertHistoryEntry,
  applySheetInsertManyHistoryEntry,
  applySheetMoveHistoryEntry,
  applySheetMoveManyHistoryEntry,
  applyWorkbookSheetsHistoryEntry,
  captureCellsForHistory,
  captureCellsForHistoryRanges,
  captureSheetForHistory,
  captureSheetsForHistory,
  cellCssColorLooksLight,
  cellDisplayForStorage,
  cellEffectiveCss,
  cellLooksExplicitlyStyled,
  clearSelectionContents,
  clearSelectionContentsForSheetGroup,
  clipboardFormulaPasteOffset,
  clipboardHasDimensions,
  cloneCellModel,
  cloneCellScalarValue,
  cloneDefinedNames,
  cloneImages,
  cloneInkStrokes,
  cloneJsonLikeValue,
  cloneMerges,
  cloneTables,
  compactCellCssAgainstDefault,
  compactCellModelForStorage,
  compactSheetDefaultCellCss,
  copyTableColumnMetrics,
  copyTableRowMetrics,
  cwsAgentCaptureWorkbookSheetsSnapshot,
  cwsAgentRestoreWorkbookSheetsSnapshot,
  defaultColumnModel,
  defaultRowModel,
  defaultableCellCss,
  deleteCellsShiftLeft,
  deleteCellsShiftUp,
  deleteEntireColumns,
  deleteEntireRows,
  deleteTableColumnMetadata,
  ensureTableAutoFilterStateForStructureChange,
  hasHistoryPayload,
  inferredDefaultableCellCss,
  initializeInsertedTableColumnsForColumnInsert,
  initializeInsertedTableRowsForRowInsert,
  insertCellsShiftDown,
  insertCellsShiftRight,
  insertEntireColumns,
  insertEntireRows,
  insertTableColumnMetadata,
  internBorderSide,
  internBordersObject,
  internCellModelStyles,
  internCellModelStylesForSheet,
  internCellsInRange,
  internCellsInRanges,
  internCssObject,
  internPooledObject,
  internSheetCellStyles,
  internWorkbookCellStyles,
  invalidateSheetStyleRangeCache,
  isBlankClipboardCell,
  localizeDefaultCellCssToExplicitlyStyledCells,
  materializeSheetStyleTable,
  mostCommonDefaultCellCss,
  mutableBordersObject,
  mutableCssObject,
  normalizeInkStrokes,
  normalizePasteSpecialOptions,
  normalizeSheetModelAfterStructureChange,
  normalizeSheetStyleRanges,
  pasteClipboardCellAllExceptBorders,
  pasteClipboardCellAllForSpecial,
  pasteClipboardCellComments,
  pasteClipboardCellContentForSpecial,
  pasteClipboardCellFormat,
  pasteClipboardCellValidation,
  pasteClipboardColumnWidths,
  pasteDisplayValueFromClipboardCell,
  pasteInternalClipboardAtRange,
  pasteInternalClipboardLinkAtRange,
  pasteInternalClipboardValuesAtRange,
  pasteSpecialModeLabel,
  pasteSpecialModeWritesContent,
  pasteSpecialModeWritesFormula,
  pasteSpecialNumericValue,
  pasteSpecialOperatedValue,
  pasteTextAtRange,
  pasteValueFromClipboardCell,
  pushHistory,
  pushSheetDeleteHistory,
  pushSheetDeleteManyHistory,
  pushSheetInsertHistory,
  pushSheetInsertManyHistory,
  pushSheetMoveHistory,
  pushSheetMoveManyHistory,
  redoLastOperation,
  refreshDataDialogAfterHistory,
  removeSheetStyleRangesAtCell,
  reorderSheetsByName,
  restoreSheetFromHistory,
  restoreSheetsFromHistory,
  sanitizeSheetDefaultCellCss,
  sheetDefaultCellCssKeyLooksDecorative,
  sheetDisplayCell,
  sheetHasExplicitlyStyledCells,
  sheetStyleRangeCell,
  sheetStyleRangesForRow,
  shiftFilterColumnMap,
  shiftImageAnchor,
  shiftImagesForColumns,
  shiftImagesForRows,
  shiftMergesForColumnDelete,
  shiftMergesForColumnInsert,
  shiftMergesForRowDelete,
  shiftMergesForRowInsert,
  shiftRangeColumns,
  shiftRangeRows,
  shiftSheetFilterAndSortForColumns,
  shiftSheetFilterAndSortForRows,
  shiftSortStateKeysForColumns,
  shiftSortStateKeysForRows,
  shiftStructureIndex,
  shiftStyleRangesForColumns,
  shiftStyleRangesForRows,
  shiftTablesForColumns,
  shiftTablesForRows,
  shiftedClipboardFormulaForTarget,
  shiftedClipboardFormulaTextForTarget,
  sourceOffsetForPasteSpecial,
  stableObjectKey,
  storeCellModelForSheet,
  syncHfDataFromCells,
  syncSheetAutoFilterToTableRangesAfterStructureChange,
  tableInsertedColumnFormatSource,
  tableInsertedRowFormatSource,
  tableStructureChangeTouchesColumns,
  tableStructureChangeTouchesRows,
  undoLastOperation,
  uniqueCellsForRanges,
  workbookMemoryStats,
  writeClonedCellContent,
  writeClonedClipboardCellContent,
};
