// Feature module; shared bindings are imported explicitly.
import {
  COLUMN_HEADER_HEIGHT,
  CWS_AGENT_MAX_CHART_CELLS,
  CWS_AGENT_MAX_CHART_SIZE,
  CWS_AGENT_MAX_COLUMN_WIDTH,
  CWS_AGENT_MAX_COMMENT_TEXT,
  CWS_AGENT_MAX_FILTER_VALUES,
  CWS_AGENT_MAX_FILTER_VALUE_TEXT,
  CWS_AGENT_MAX_FORECAST_CELLS,
  CWS_AGENT_MAX_GOAL_SEEK_ABS_TARGET,
  CWS_AGENT_MAX_HYPERLINK_TEXT,
  CWS_AGENT_MAX_NOTE_TEXT,
  CWS_AGENT_MAX_PIVOT_CELLS,
  CWS_AGENT_MAX_QUICK_ANALYSIS_CELLS,
  CWS_AGENT_MAX_ROW_HEIGHT,
  CWS_AGENT_MAX_SHAPE_SIZE,
  CWS_AGENT_MAX_SHAPE_TEXT,
  CWS_AGENT_MAX_TABLE_CELLS,
  CWS_AGENT_MAX_TEXT_TO_COLUMNS_CELLS,
  CWS_AGENT_MAX_VALIDATION_OPTIONS,
  CWS_AGENT_MAX_VALIDATION_OPTION_TEXT,
  CWS_AGENT_MIN_CHART_SIZE,
  CWS_AGENT_MIN_COLUMN_WIDTH,
  CWS_AGENT_MIN_ROW_HEIGHT,
  CWS_AGENT_MIN_SHAPE_SIZE,
  DEFAULT_SHAPE_TYPE,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  cwsAgentOperationSources,
  state,
} from "../../app-core.js";
import {
  normalizeCwsAgentOp,
} from "./contracts.js";
import {
  cwsAgentBoolean,
  cwsAgentCellValue,
  cwsAgentColumnValue,
  cwsAgentFirstExistingValue,
  cwsAgentInteger,
  cwsAgentOptionalBoolean,
  cwsAgentPositiveInteger,
  cwsAgentRgbColor,
  cwsAgentSheetIndexByName,
  cwsAgentSheetIndexForOp,
  cwsAgentSheetNameText,
} from "./normalization.js";
import {
  applyCwsAgentActivateSheetOp,
  applyCwsAgentAddShapeOp,
  applyCwsAgentAddSheetOp,
  applyCwsAgentClearCommentOp,
  applyCwsAgentClearDataValidationOp,
  applyCwsAgentClearFreezePanesOp,
  applyCwsAgentClearHeaderFooterOp,
  applyCwsAgentClearHyperlinkOp,
  applyCwsAgentClearNamedRangeOp,
  applyCwsAgentClearNoteOp,
  applyCwsAgentClearPrintAreaOp,
  applyCwsAgentClearPrintTitlesOp,
  applyCwsAgentClearQuickAnalysisOp,
  applyCwsAgentClearSubtotalsOp,
  applyCwsAgentConsolidateOp,
  applyCwsAgentCopyMoveRangeOp,
  applyCwsAgentCreateChartOp,
  applyCwsAgentCreateForecastSheetOp,
  applyCwsAgentCreatePivotTableOp,
  applyCwsAgentCreateTableOp,
  applyCwsAgentCreateWorkbookTemplateOp,
  applyCwsAgentDataTableOp,
  applyCwsAgentDeleteSheetOp,
  applyCwsAgentDuplicateSheetOp,
  applyCwsAgentGoalSeekOp,
  applyCwsAgentMoveSheetOp,
  applyCwsAgentPageBreakOp,
  applyCwsAgentProtectSheetOp,
  applyCwsAgentRenameSheetOp,
  applyCwsAgentSelectRangeOp,
  applyCwsAgentSetCommentOp,
  applyCwsAgentSetDataValidationOp,
  applyCwsAgentSetFreezePanesOp,
  applyCwsAgentSetHeaderFooterOp,
  applyCwsAgentSetHyperlinkOp,
  applyCwsAgentSetNamedRangeOp,
  applyCwsAgentSetNoteOp,
  applyCwsAgentSetPageMarginsOp,
  applyCwsAgentSetPageOrientationOp,
  applyCwsAgentSetPageScaleOp,
  applyCwsAgentSetPageThemeOp,
  applyCwsAgentSetPaperSizeOp,
  applyCwsAgentSetPrintAreaOp,
  applyCwsAgentSetPrintOptionsOp,
  applyCwsAgentSetPrintTitlesOp,
  applyCwsAgentSetQuickAnalysisOp,
  applyCwsAgentSetQuickAnalysisTotalOp,
  applyCwsAgentSetShapeTextOp,
  applyCwsAgentExistingShapeEditOp,
  applyCwsAgentSetSheetViewOp,
  applyCwsAgentSetSubtotalOp,
  applyCwsAgentSheetTabColorOp,
  applyCwsAgentSheetVisibilityOp,
  applyCwsAgentUnprotectSheetOp,
  cwsAgentCellMatchesFindQuery,
} from "./shapes-and-sheet-operations.js";
import {
  advancedFilterCopyTargetInfo,
  advancedFilterCriteriaHeaderIssues,
  advancedFilterMatchingRows,
  advancedFilterOutputColumnDetails,
  applyAdvancedFilterInPlace,
  coerceTextColumnsValue,
  copyAdvancedFilterResults,
  parseFixedWidthBreaks,
  sheetForWorkbookRange,
  textColumnsOutputRows,
  textToColumns,
  textToColumnsRowsForRange,
} from "../data/filter-and-text-tools.js";
import {
  replaceCellText,
} from "../data/query-editor.js";
import {
  autoFilterStatusMessage,
  inferHeaderRow,
  rangeHasUserData,
  reapplyAutoFilter,
  removeAutoFilter,
  sortRangeRowsByKeys,
  syncAutoFilterStateToMatchingTable,
} from "../data/sort-and-autofilter.js";
import {
  applyCollectedRemoveDuplicateRows,
  collectRemoveDuplicateRows,
  removeDuplicatesMergeConflict,
  validateCellInput,
} from "../data/validation-and-consolidation.js";
import {
  cellNumericValue,
} from "../drawing/charts-and-insertion.js";
import {
  normalizeShapeArrowType,
  normalizeShapeLineDash,
  shapeDescriptorForType,
  shapePresetForType,
} from "../drawing/shape-rendering.js";
import {
  refreshWorkbookChartsAfterSourceEdit,
} from "../editing/autofill.js";
import {
  normalizeColorValue,
} from "../editing/ink-and-clipboard.js";
import {
  coerceCellInput,
  coerceUserInput,
  commitCellValue,
  expandRangeForMerges,
  getCellRawInput,
  normalizeRange,
  parseFormattedNumericInput,
  rangeToLabel,
  updateSelectionUi,
  writeCellContent,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  coordinateToPixels,
  updateFormulaBarForSelection,
} from "../editing/selection-and-cell-text.js";
import {
  captureCellsForHistory,
  captureSheetForHistory,
  internCellsInRange,
  pushHistory,
  storeCellModelForSheet,
} from "../editing/structure-and-history.js";
import {
  applyBorderPresetToSheetRange,
  applyCellNumberFormat,
  applyCssPatchToSheetRanges,
  ensureCellModel,
  forEachCellInRange,
  renderSheet,
} from "../rendering/grid-and-canvas.js";
import {
  removeCellModelIfEmpty,
} from "../ui/annotations-and-cell-format.js";
import {
  cellKey,
  columnName,
  columnWidthToPixels,
  currencySymbol,
  rowHeightToPixels,
  setStatus,
} from "../ui/help-and-utilities.js";
import {
  commitDataOperationForSheet,
  ensureSheetSize,
  nextTableName,
  normalizeHyperlinkForKind,
  rowHeaderWidthForSheet,
  validateDefinedName,
} from "../ui/ribbon-and-links.js";
import {
  clampZoom,
  defaultSheetProtectionOptions,
  ensureEditableRangeForProtection,
  ensureEditableRangesForProtection,
  ensureSheetProtectionPermission,
} from "../ui/sheet-protection-and-translation.js";
import {
  buildFormulaEngine,
  isSheetVisible,
  renderWorkbook,
  setActiveSheet,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  activeSheet,
  cellDisplayText,
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
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
} from "../workbook/startup-and-loading.js";
import {
  parseValidationOptions,
  rangeRectEquals,
  validationReferenceFromFormula,
} from "../workbook/tables-and-validation.js";


function cwsAgentTextToColumnsRangeForOp(op, index) {
  const rangeText = String(cwsAgentFirstExistingValue(op, ["sourceRange", "range", "address", "ref", "cell", "source"])).trim();
  const parsedRange = cwsAgentParseRangeAddress(rangeText);
  const top = cwsAgentPositiveInteger(op?.sourceTop ?? op?.top ?? op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.sourceLeft ?? op?.left ?? op?.col ?? op?.column) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.sourceBottom ?? op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.sourceRight ?? op?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の分列元範囲が正しくありません。`);
  }
  if (right !== left) {
    throw new Error(`Agent 操作 ${index + 1} の分列元範囲は1列だけ指定してください。`);
  }
  const cellCount = bottom - top + 1;
  if (cellCount > CWS_AGENT_MAX_TEXT_TO_COLUMNS_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} の分列元範囲が大きすぎます。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentCreateTableRangeForOp(op, index) {
  const rangeText = String(cwsAgentFirstExistingValue(op, ["range", "address", "ref", "cell", "source", "target"])).trim();
  const parsedRange = cwsAgentParseRangeAddress(rangeText);
  if (parsedRange?.fullRow || parsedRange?.fullColumn) {
    throw new Error(`Agent 操作 ${index + 1} のテーブル範囲はセル範囲で指定してください。`);
  }
  const top = cwsAgentPositiveInteger(op?.top ?? op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.left ?? op?.col ?? op?.column) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のテーブル範囲が正しくありません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount > CWS_AGENT_MAX_TABLE_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} のテーブル範囲が大きすぎます。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentChartRangeForOp(op, index) {
  const rangeText = String(cwsAgentFirstExistingValue(op, ["sourceRange", "range", "address", "ref", "cell", "source", "dataRange"])).trim();
  const parsedRange = cwsAgentParseRangeAddress(rangeText);
  if (parsedRange?.fullRow || parsedRange?.fullColumn) {
    throw new Error(`Agent 操作 ${index + 1} のグラフ範囲はセル範囲で指定してください。`);
  }
  const top = cwsAgentPositiveInteger(op?.sourceTop ?? op?.top ?? op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.sourceLeft ?? op?.left ?? op?.col ?? op?.column) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.sourceBottom ?? op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.sourceRight ?? op?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のグラフ範囲が正しくありません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount > CWS_AGENT_MAX_CHART_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} のグラフ範囲が大きすぎます。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentPivotRangeForOp(op, index) {
  const rangeText = String(cwsAgentFirstExistingValue(op, ["sourceRange", "range", "address", "ref", "cell", "source", "dataRange"])).trim();
  const parsedRange = cwsAgentParseRangeAddress(rangeText);
  if (parsedRange?.fullRow || parsedRange?.fullColumn) {
    throw new Error(`Agent 操作 ${index + 1} のピボットテーブル範囲はセル範囲で指定してください。`);
  }
  const top = cwsAgentPositiveInteger(op?.sourceTop ?? op?.top ?? op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.sourceLeft ?? op?.left ?? op?.col ?? op?.column) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.sourceBottom ?? op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.sourceRight ?? op?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のピボットテーブル範囲が正しくありません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount > CWS_AGENT_MAX_PIVOT_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} のピボットテーブル範囲が大きすぎます。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(
    { ...op, sheetName: op?.sourceSheetName ?? op?.sheetName },
    parsedRange?.sheetName,
  );
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentForecastRangeForOp(op, index) {
  const rangeText = String(cwsAgentFirstExistingValue(op, ["sourceRange", "range", "address", "ref", "cell", "source", "dataRange"])).trim();
  const parsedRange = cwsAgentParseRangeAddress(rangeText);
  if (parsedRange?.fullRow || parsedRange?.fullColumn) {
    throw new Error(`Agent 操作 ${index + 1} の予測範囲はセル範囲で指定してください。`);
  }
  const top = cwsAgentPositiveInteger(op?.sourceTop ?? op?.top ?? op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.sourceLeft ?? op?.left ?? op?.col ?? op?.column) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.sourceBottom ?? op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.sourceRight ?? op?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の予測範囲が正しくありません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount > CWS_AGENT_MAX_FORECAST_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} の予測範囲が大きすぎます。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(
    { ...op, sheetName: op?.sourceSheetName ?? op?.sheetName },
    parsedRange?.sheetName,
  );
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentShapeLocationForOp(op, index, fallbackSheetName = "") {
  const rangeText = String(cwsAgentFirstExistingValue(op, ["targetRange", "shapeRange", "placeRange", "anchorRange", "range"])).trim();
  const parsedRange = cwsAgentParseRangeAddress(rangeText);
  const cellText = String(cwsAgentFirstExistingValue(op, ["targetCell", "anchorCell", "topLeftCell", "cell", "address", "ref", "at"])).trim();
  const parsedCell = cwsAgentParseCellAddress(cellText) || (parsedRange ? { sheetName: parsedRange.sheetName, row: parsedRange.top, col: parsedRange.left } : null);
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedCell?.sheetName || parsedRange?.sheetName || fallbackSheetName || "");
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  if (parsedRange && (parsedRange.bottom > EXCEL_MAX_ROWS || parsedRange.right > EXCEL_MAX_COLS)) {
    throw new Error(`Agent 操作 ${index + 1} の図形配置範囲が正しくありません。`);
  }
  if (parsedCell && (parsedCell.row > EXCEL_MAX_ROWS || parsedCell.col > EXCEL_MAX_COLS)) {
    throw new Error(`Agent 操作 ${index + 1} の図形配置セルが正しくありません。`);
  }
  const anchorLabel = parsedRange
    ? `${columnName(parsedRange.left)}${parsedRange.top}:${columnName(parsedRange.right)}${parsedRange.bottom}`
    : parsedCell
      ? `${columnName(parsedCell.col)}${parsedCell.row}`
      : "図形";
  return { sheetIndex, parsedRange, parsedCell, anchorLabel };
}



function cwsAgentLineEndpointSide(value) {
  const normalized = String(value || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    left: "left",
    l: "left",
    west: "left",
    左: "left",
    左側: "left",
    right: "right",
    r: "right",
    east: "right",
    右: "right",
    右側: "right",
    top: "top",
    t: "top",
    north: "top",
    上: "top",
    上側: "top",
    bottom: "bottom",
    b: "bottom",
    south: "bottom",
    下: "bottom",
    下側: "bottom",
    center: "center",
    middle: "center",
    centre: "center",
    中央: "center",
    中心: "center",
  };
  return aliases[normalized] || "center";
}



function cwsAgentShapeBoxForOp(op, index, sheet, location, shapeType, style = {}) {
  const defaultDescriptor = shapeDescriptorForType(shapeType, { useThemeEffects: true, ...style });
  const explicitLeft = cwsAgentShapePosition(op?.left ?? op?.x);
  const explicitTop = cwsAgentShapePosition(op?.top ?? op?.y);
  const explicitWidth = cwsAgentShapeSize(op?.width ?? op?.w ?? op?.size?.width);
  const explicitHeight = cwsAgentShapeSize(op?.height ?? op?.h ?? op?.size?.height);

  // Agent pixel positions use the worksheet body origin, just like the server
  // layout validator. Insertion/native anchors use sheet pixels including the
  // row/column headings, as coordinateToPixels already does for cell ranges.
  let left = explicitLeft === null ? null : explicitLeft + rowHeaderWidthForSheet(sheet);
  let top = explicitTop === null ? null : explicitTop + COLUMN_HEADER_HEIGHT;
  let rangeWidth = null;
  let rangeHeight = null;
  if (location.parsedRange) {
    left ??= coordinateToPixels(sheet, "col", location.parsedRange.left - 1);
    top ??= coordinateToPixels(sheet, "row", location.parsedRange.top - 1);
    rangeWidth = coordinateToPixels(sheet, "col", location.parsedRange.right) - coordinateToPixels(sheet, "col", location.parsedRange.left - 1);
    rangeHeight = coordinateToPixels(sheet, "row", location.parsedRange.bottom) - coordinateToPixels(sheet, "row", location.parsedRange.top - 1);
  } else if (location.parsedCell) {
    left ??= coordinateToPixels(sheet, "col", location.parsedCell.col - 1);
    top ??= coordinateToPixels(sheet, "row", location.parsedCell.row - 1);
  } else if (location.sheetIndex === state.activeSheetIndex && state.selectionRange) {
    left ??= coordinateToPixels(sheet, "col", state.selectionRange.left - 1);
    top ??= coordinateToPixels(sheet, "row", state.selectionRange.top - 1);
  }

  left ??= coordinateToPixels(sheet, "col", 0);
  top ??= coordinateToPixels(sheet, "row", 0);
  if (explicitLeft === null) left += cwsAgentShapePosition(op?.anchorOffsetX) ?? 0;
  if (explicitTop === null) top += cwsAgentShapePosition(op?.anchorOffsetY) ?? 0;
  // Range-derived geometry must cover the actual selected cells. The explicit
  // size limit must not silently shrink a valid wide title or diagram label.
  const width = cwsAgentClampDimension(explicitWidth ?? rangeWidth ?? defaultDescriptor.width, CWS_AGENT_MIN_SHAPE_SIZE, explicitWidth !== null ? CWS_AGENT_MAX_SHAPE_SIZE : Math.max(CWS_AGENT_MAX_SHAPE_SIZE, rangeWidth || 0));
  const height = cwsAgentClampDimension(explicitHeight ?? rangeHeight ?? defaultDescriptor.height, CWS_AGENT_MIN_SHAPE_SIZE, explicitHeight !== null ? CWS_AGENT_MAX_SHAPE_SIZE : Math.max(CWS_AGENT_MAX_SHAPE_SIZE, rangeHeight || 0));
  if (!Number.isFinite(left) || !Number.isFinite(top) || !width || !height) {
    throw new Error(`Agent 操作 ${index + 1} の図形位置またはサイズが正しくありません。`);
  }
  return {
    left: Math.max(0, Math.round(left)),
    top: Math.max(0, Math.round(top)),
    width,
    height,
  };
}



function cwsAgentShapeType(value, aliasType = "") {
  const raw = String(value || "").trim();
  if (raw && shapePresetForType(raw)?.type) return shapePresetForType(raw).type;
  const normalized = String(raw || aliasType || "")
    .trim()
    .toLowerCase()
    .replace(/[\s_-]+/g, "");
  const aliases = {
    "": DEFAULT_SHAPE_TYPE,
    addshape: DEFAULT_SHAPE_TYPE,
    createshape: DEFAULT_SHAPE_TYPE,
    insertshape: DEFAULT_SHAPE_TYPE,
    drawshape: DEFAULT_SHAPE_TYPE,
    shape: DEFAULT_SHAPE_TYPE,
    autoshape: DEFAULT_SHAPE_TYPE,
    addconnector: "arrow",
    createconnector: "arrow",
    insertconnector: "arrow",
    drawconnector: "arrow",
    connectorline: "line",
    lineconnector: "line",
    arrowconnector: "arrow",
    linearrow: "arrow",
    rectangle: "rect",
    rect: "rect",
    square: "rect",
    四角形: "rect",
    矩形: "rect",
    正方形: "rect",
    textbox: "textBox",
    textshape: "textBox",
    text: "textBox",
    テキストボックス: "textBox",
    文本框: "textBox",
    文字框: "textBox",
    roundrect: "roundRect",
    roundedrectangle: "roundRect",
    roundedrect: "roundRect",
    oval: "oval",
    ellipse: "oval",
    circle: "oval",
    円: "oval",
    楕円: "oval",
    圆形: "oval",
    圓形: "oval",
    triangle: "triangle",
    三角形: "triangle",
    diamond: "diamond",
    ひし形: "diamond",
    菱形: "diamond",
    arrow: "rightBlockArrow",
    rightarrow: "rightBlockArrow",
    rightblockarrow: "rightBlockArrow",
    矢印: "rightBlockArrow",
    右矢印: "rightBlockArrow",
    箭头: "rightBlockArrow",
    箭頭: "rightBlockArrow",
    leftarrow: "leftBlockArrow",
    uparrow: "upBlockArrow",
    downarrow: "downBlockArrow",
    line: "line",
    connector: "line",
    callout: "calloutRect",
    balloon: "calloutRect",
    吹き出し: "calloutRect",
    标注: "calloutRect",
    標注: "calloutRect",
    process: "flowProcess",
    flowprocess: "flowProcess",
    処理: "flowProcess",
    处理: "flowProcess",
    處理: "flowProcess",
    decision: "flowDecision",
    flowdecision: "flowDecision",
    判断: "flowDecision",
    判定: "flowDecision",
    条件: "flowDecision",
    條件: "flowDecision",
    terminator: "flowTerminator",
    flowterminator: "flowTerminator",
    startend: "flowTerminator",
    startorend: "flowTerminator",
    開始終了: "flowTerminator",
    开始结束: "flowTerminator",
    開始結束: "flowTerminator",
    flowchart: "flowProcess",
  };
  const type = aliases[normalized] || raw;
  const preset = shapePresetForType(type);
  if (!preset) return DEFAULT_SHAPE_TYPE;
  return preset.type;
}



function cwsAgentShapeStyleForOp(op, index) {
  const merged = op?.style && typeof op.style === "object" ? { ...op, ...op.style } : op || {};
  const style = { useThemeEffects: true };
  if (Object.prototype.hasOwnProperty.call(merged, 'points')) style.points = cwsAgentShapeLinePoints(merged.points);
  if ((typeof merged.rotation === "number" || typeof merged.rotation === "string" && merged.rotation.trim() !== "") && Number.isFinite(Number(merged.rotation))) style.rotation = ((Number(merged.rotation) % 360) + 360) % 360;
  if (["x", "y"].includes(merged.elbowAxis)) style.adjust = { axis: merged.elbowAxis, bend: 0.5 };
  const bend = typeof merged.elbowBend === "number" || (typeof merged.elbowBend === "string" && merged.elbowBend.trim()) ? Number(merged.elbowBend) : NaN;
  if (Number.isFinite(bend) && bend >= 0 && bend <= 1) style.adjust = { ...style.adjust, bend };
  if (merged.lineDash !== undefined) style.lineDash = normalizeShapeLineDash(merged.lineDash);
  const fillColor = cwsAgentShapeColor(merged.fillColor ?? merged.fill ?? merged.backgroundColor);
  if (fillColor !== undefined) style.fillColor = fillColor;
  const outlineColor = cwsAgentShapeColor(merged.outlineColor ?? merged.lineColor ?? merged.borderColor ?? merged.strokeColor ?? merged.stroke);
  if (outlineColor !== undefined) style.outlineColor = outlineColor;
  const textColor = cwsAgentShapeColor(merged.textColor ?? merged.fontColor ?? merged.color);
  if (textColor !== undefined && textColor !== "none") style.textColor = textColor;
  const outlineWidth = cwsAgentShapeOutlineWidth(merged.outlineWidth ?? merged.lineWidth ?? merged.strokeWidth ?? merged.borderWidth);
  if (outlineWidth !== null) style.outlineWidth = outlineWidth;
  const text = cwsAgentShapeText(merged.text ?? merged.label ?? merged.caption, index);
  if (text !== null) style.text = text;
  const textFontSize = cwsAgentShapeFontSize(merged.textFontSize ?? merged.fontSize);
  if (textFontSize !== null) style.textFontSize = textFontSize;
  const family = merged.textFontFamily ?? merged.fontFamily;
  if (typeof family === 'string' && /^[\p{L}\p{N} ,"'_-]{1,240}$/u.test(family.trim())) style.textFontFamily = family.trim();
  const underline = merged.textUnderline ?? (typeof merged.underline === 'boolean' ? (merged.underline ? 'underline' : '') : null);
  if (['', 'underline', 'underline double'].includes(underline)) style.textUnderline = underline;
  if (merged.textMargin && typeof merged.textMargin === 'object') {
    const margin={};
    for (const side of ['left','right','top','bottom']) {
      const value=merged.textMargin[side];
      if (value !== null && value !== '' && Number.isFinite(Number(value))) margin[side]=Math.max(0,Math.min(80,Number(value)));
    }
    if (Object.keys(margin).length) style.textMargin=margin;
  }
  const textAlign = String(merged.textAlign ?? merged.align ?? "").trim().toLowerCase();
  if (["left", "center", "right"].includes(textAlign)) style.textAlign = textAlign;
  const textVAlign = String(merged.textVAlign ?? merged.verticalAlign ?? merged.valign ?? "").trim().toLowerCase();
  if (["top", "middle", "bottom"].includes(textVAlign)) style.textVAlign = textVAlign;
  if (Object.prototype.hasOwnProperty.call(merged, "bold") || Object.prototype.hasOwnProperty.call(merged, "textBold")) {
    style.textBold = cwsAgentBoolean(merged.bold ?? merged.textBold);
  }
  if (Object.prototype.hasOwnProperty.call(merged, "italic") || Object.prototype.hasOwnProperty.call(merged, "textItalic")) {
    style.textItalic = cwsAgentBoolean(merged.italic ?? merged.textItalic);
  }
  const startArrowType = cwsAgentShapeArrowTypeValue(merged.lineStartArrowType ?? merged.startArrowType ?? merged.startArrow);
  const endArrowType = cwsAgentShapeArrowTypeValue(merged.lineEndArrowType ?? merged.endArrowType ?? merged.endArrow);
  if (Object.prototype.hasOwnProperty.call(merged, "lineStartArrow")) style.lineStartArrow = cwsAgentBoolean(merged.lineStartArrow);
  if (Object.prototype.hasOwnProperty.call(merged, "lineEndArrow")) style.lineEndArrow = cwsAgentBoolean(merged.lineEndArrow);
  if (startArrowType) {
    style.lineStartArrow = true;
    style.lineStartArrowType = startArrowType;
  }
  if (endArrowType) {
    style.lineEndArrow = true;
    style.lineEndArrowType = endArrowType;
  }
  return style;
}



function cwsAgentShapeLinePoints(value) {
  const keys = ['x1', 'y1', 'x2', 'y2'];
  if (!value || typeof value !== 'object' || Array.isArray(value) || Object.keys(value).length !== 4
    || keys.some(key => !['number', 'string'].includes(typeof value[key]) || String(value[key]).trim() === '' || !Number.isFinite(Number(value[key])) || Number(value[key]) < 0 || Number(value[key]) > 1)) throw new Error('線の端点には 0〜1 の x1/y1/x2/y2 が必要です。');
  const points = Object.fromEntries(keys.map(key => [key, Number(value[key])]));
  if (Math.abs(points.x1 - points.x2) < 0.001 && Math.abs(points.y1 - points.y2) < 0.001) throw new Error('線の始点と終点は異なる位置が必要です。');
  return points;
}



function cwsAgentShapeArrowTypeValue(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "boolean") return value ? "triangle" : "";
  return normalizeShapeArrowType(value);
}



function cwsAgentShapeColor(value) {
  if (value === null || value === undefined || value === "") return undefined;
  const text = String(value).trim();
  const lower = text.toLowerCase();
  if (["none", "transparent", "なし", "透明", "无", "無"].includes(lower)) return "none";
  const named = {
    red: "#C00000",
    orange: "#ED7D31",
    yellow: "#FFC000",
    green: "#70AD47",
    blue: "#4472C4",
    purple: "#7030A0",
    gray: "#808080",
    grey: "#808080",
    black: "#000000",
    white: "#FFFFFF",
    赤: "#C00000",
    青: "#4472C4",
    蓝: "#4472C4",
    藍: "#4472C4",
    緑: "#70AD47",
    绿: "#70AD47",
    綠: "#70AD47",
    黄: "#FFC000",
    黄色: "#FFC000",
    黒: "#000000",
    黑: "#000000",
    白: "#FFFFFF",
    紫: "#7030A0",
  };
  return normalizeColorValue(text) || cwsAgentRgbColor(text) || named[lower] || undefined;
}



function cwsAgentShapeSize(value) {
  if (value === null || value === undefined || value === "") return null;
  const text = String(value).trim();
  const match = text.match(/^([0-9]+(?:\.[0-9]+)?)\s*(px)?$/i);
  const number = match ? Number(match[1]) : Number(value);
  return Number.isFinite(number) && number > 0 ? number : null;
}



function cwsAgentShapePosition(value) {
  if (value === null || value === undefined || value === "") return null;
  const text = String(value).trim();
  const match = text.match(/^(-?[0-9]+(?:\.[0-9]+)?)\s*(px)?$/i);
  const number = match ? Number(match[1]) : Number(value);
  return Number.isFinite(number) ? Math.round(number) : null;
}



function cwsAgentShapeOutlineWidth(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? Math.max(0.6, Math.min(6, Math.round(number * 100) / 100)) : null;
}



function cwsAgentShapeFontSize(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? Math.max(6, Math.min(72, Math.round(number * 10) / 10)) : null;
}



function cwsAgentShapeText(value, index) {
  if (value === null || value === undefined) return null;
  if (Array.isArray(value) || typeof value === "object") {
    throw new Error(`Agent 操作 ${index + 1} の図形テキストが正しくありません。`);
  }
  return String(value).replace(/\r\n?/g, "\n").slice(0, CWS_AGENT_MAX_SHAPE_TEXT);
}



function cwsAgentShapeObjectName(value) {
  if (value === null || value === undefined || Array.isArray(value) || typeof value === "object") return "";
  return String(value).trim().slice(0, 80);
}



function cwsAgentQuickAnalysisRangeForOp(op, index) {
  const rangeText = String(cwsAgentFirstExistingValue(op, ["range", "address", "ref", "cell", "source", "target", "dataRange"])).trim();
  const parsedRange = cwsAgentParseRangeAddress(rangeText);
  if (parsedRange?.fullRow || parsedRange?.fullColumn) {
    throw new Error(`Agent 操作 ${index + 1} のクイック分析範囲はセル範囲で指定してください。`);
  }
  const top = cwsAgentPositiveInteger(op?.top ?? op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.left ?? op?.col ?? op?.column) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のクイック分析範囲が正しくありません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount > CWS_AGENT_MAX_QUICK_ANALYSIS_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} のクイック分析範囲が大きすぎます。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentTextToColumnsDestinationForOp(op, sourceRange, index) {
  const destinationText = String(cwsAgentFirstExistingValue(op, [
    "destination",
    "destinationCell",
    "destinationRange",
    "target",
    "targetCell",
    "targetRange",
    "output",
    "outputCell",
    "outputRange",
    "to",
    "at",
  ])).trim();
  const parsedCell = cwsAgentParseCellAddress(destinationText);
  const parsedRange = cwsAgentParseRangeAddress(destinationText);
  const parsedSheetName = parsedCell?.sheetName || parsedRange?.sheetName || "";
  if (parsedSheetName && cwsAgentSheetIndexByName(parsedSheetName) !== sourceRange.sheetIndex) {
    throw new Error(`Agent 操作 ${index + 1} の分列表示先は同じシート内で指定してください。`);
  }
  const row = cwsAgentPositiveInteger(op?.destinationRow ?? op?.targetRow ?? op?.outputRow ?? op?.toRow)
    || parsedCell?.row
    || parsedRange?.top
    || sourceRange.top;
  const col = cwsAgentColumnValue(op?.destinationCol ?? op?.destinationColumn ?? op?.targetCol ?? op?.targetColumn ?? op?.outputCol ?? op?.outputColumn ?? op?.toCol)
    || parsedCell?.col
    || parsedRange?.left
    || sourceRange.left;
  if (!row || !col || row > EXCEL_MAX_ROWS || col > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の分列表示先が正しくありません。`);
  }
  return { row, col };
}



function cwsAgentTextToColumnsOptions(op, index) {
  const mode = cwsAgentTextToColumnsMode(op?.mode ?? op?.splitMode ?? op?.kind);
  const delimiters = mode === "fixed" ? [] : cwsAgentTextToColumnsDelimiters(
    cwsAgentFirstExistingValue(op, ["delimiters", "delimiter", "separator", "separators", "splitBy"]),
    op,
  );
  const fixedBreaks = mode === "fixed" ? cwsAgentTextToColumnsFixedBreaks(op?.fixedBreaks ?? op?.breaks ?? op?.positions ?? op?.widths) : [];
  if (mode === "delimited" && !delimiters.length) {
    throw new Error(`Agent 操作 ${index + 1} の分列区切り文字が正しくありません。`);
  }
  if (mode === "fixed" && !fixedBreaks.length) {
    throw new Error(`Agent 操作 ${index + 1} の固定長の区切り位置が正しくありません。`);
  }
  const decimalSeparator = cwsAgentSingleCharacter(op?.decimalSeparator ?? op?.decimal, ".");
  const thousandsSeparator = cwsAgentSingleCharacter(op?.thousandsSeparator ?? op?.thousands, ",");
  if (decimalSeparator && thousandsSeparator && decimalSeparator === thousandsSeparator) {
    throw new Error(`Agent 操作 ${index + 1} の小数点と桁区切り記号は別々に指定してください。`);
  }
  return {
    mode,
    delimiters,
    consecutive: cwsAgentOptionalBoolean(op?.consecutive ?? op?.treatConsecutiveDelimitersAsOne, true),
    qualifier: cwsAgentSingleCharacter(op?.qualifier ?? op?.textQualifier ?? op?.quote, "\""),
    fixedBreaks,
    columnFormats: cwsAgentTextToColumnsColumnFormats(op?.columnFormats ?? op?.formats ?? op?.columnFormat),
    decimalSeparator,
    thousandsSeparator,
    trailingMinus: cwsAgentBoolean(op?.trailingMinus),
  };
}



function cwsAgentTextToColumnsMode(value) {
  const text = String(value || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["fixed", "fixedwidth", "fixedlength", "固定長", "固定幅", "定长", "定寬", "定宽"].includes(text)) return "fixed";
  return "delimited";
}



function cwsAgentTextToColumnsDelimiters(value, op = {}) {
  const result = [];
  const add = (item) => {
    const delimiter = cwsAgentTextToColumnsDelimiterToken(item);
    if (delimiter !== null && !result.includes(delimiter)) result.push(delimiter);
  };
  if (Array.isArray(value)) {
    value.forEach(add);
  } else if (value !== null && value !== undefined && value !== "") {
    const raw = String(value);
    const trimmed = raw.trim();
    const direct = cwsAgentTextToColumnsDelimiterToken(raw);
    if (direct !== null && (raw.length === 1 || cwsAgentTextToColumnsDelimiterToken(trimmed) !== null)) {
      add(raw);
    } else {
      trimmed.split(/[,;\n]+/).forEach(add);
    }
  }
  if (cwsAgentBoolean(op?.tab)) add("tab");
  if (cwsAgentBoolean(op?.comma)) add("comma");
  if (cwsAgentBoolean(op?.semicolon)) add("semicolon");
  if (cwsAgentBoolean(op?.space)) add("space");
  if (op?.otherDelimiter || op?.other) add(op.otherDelimiter ?? op.other);
  return result.length ? result : [","];
}



function cwsAgentTextToColumnsDelimiterToken(value) {
  if (value === null || value === undefined) return null;
  const raw = String(value);
  if (raw === " ") return " ";
  if (raw === "\t" || raw === "\\t") return "\t";
  if (raw.length === 1 && raw !== "\n" && raw !== "\r") return raw;
  const key = raw.trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    tab: "\t",
    tabs: "\t",
    タブ: "\t",
    comma: ",",
    commas: ",",
    カンマ: ",",
    逗号: ",",
    逗號: ",",
    semicolon: ";",
    semicolons: ";",
    セミコロン: ";",
    space: " ",
    spaces: " ",
    blank: " ",
    スペース: " ",
    空格: " ",
    pipe: "|",
    pipes: "|",
    verticalbar: "|",
  };
  return Object.prototype.hasOwnProperty.call(aliases, key) ? aliases[key] : null;
}



function cwsAgentTextToColumnsFixedBreaks(value) {
  if (Array.isArray(value)) {
    return [...new Set(value
      .map((item) => Number.parseInt(item, 10))
      .filter((item) => Number.isFinite(item) && item > 0))]
      .sort((left, right) => left - right);
  }
  return parseFixedWidthBreaks(value);
}



function cwsAgentTextToColumnsColumnFormats(value) {
  if (value === null || value === undefined || value === "") return {};
  const formats = {};
  if (Array.isArray(value)) {
    value.forEach((item, index) => {
      const format = cwsAgentTextToColumnsColumnFormat(item);
      if (format) formats[index] = format;
    });
    return formats;
  }
  if (typeof value === "object") {
    Object.entries(value).forEach(([key, item]) => {
      const index = Number.parseInt(String(key).replace(/^col(?:umn)?/i, ""), 10);
      const format = cwsAgentTextToColumnsColumnFormat(item);
      if (Number.isFinite(index) && index >= 0 && format) formats[index] = format;
    });
  }
  return formats;
}



function cwsAgentTextToColumnsColumnFormat(value) {
  const text = String(value || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (!text || ["general", "standard", "標準"].includes(text)) return "general";
  if (["text", "string", "文字列", "文本"].includes(text)) return "text";
  if (["skip", "delete", "omit", "削除", "跳过", "略過"].includes(text)) return "skip";
  const dateMatch = text.match(/^(?:date)?(mdy|dmy|ymd|myd|dym|ydm)$/i);
  if (dateMatch) return `date${dateMatch[1].toUpperCase()}`;
  if (text === "date" || text === "日付" || text === "日期") return "dateYMD";
  return "";
}



function cwsAgentSingleCharacter(value, fallback = "") {
  if (value === null || value === undefined || value === "") return fallback;
  const text = String(value);
  const key = text.trim().toLowerCase();
  if (["none", "なし", "無", "无", "無効"].includes(key)) return "";
  if (text === "\\t") return "\t";
  return text[0] || fallback;
}



function cwsAgentTableName(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  let name = raw
    .replace(/\s+/g, "_")
    .replace(/[^A-Za-z0-9_]/g, "_")
    .replace(/^([^A-Za-z_])/, "_$1")
    .slice(0, 120);
  name = name.replace(/_+/g, "_").replace(/^_+$/g, "");
  return name || "";
}



function cwsAgentTableStyle(value) {
  const text = String(value ?? "").trim();
  if (!text) return "TableStyleMedium2";
  const compact = text.replace(/[\s_-]+/g, "");
  const direct = compact.match(/^TableStyle(Light|Medium|Dark)(\d{1,2})$/i);
  if (direct) {
    return `TableStyle${direct[1][0].toUpperCase()}${direct[1].slice(1).toLowerCase()}${direct[2]}`;
  }
  const shorthand = compact.match(/^(Light|Medium|Dark)(\d{1,2})$/i);
  if (shorthand) {
    return `TableStyle${shorthand[1][0].toUpperCase()}${shorthand[1].slice(1).toLowerCase()}${shorthand[2]}`;
  }
  return "TableStyleMedium2";
}



function cwsAgentUniqueTableName(sheet, preferredName = "") {
  const existing = new Set((sheet?.tables || []).map((table) => String(table.displayName || table.name || "").toLowerCase()));
  const base = preferredName || nextTableName(sheet);
  if (!existing.has(base.toLowerCase())) return base;
  for (let index = 2; index < 10000; index += 1) {
    const candidate = `${base}_${index}`;
    if (!existing.has(candidate.toLowerCase())) return candidate;
  }
  return nextTableName(sheet);
}



function cwsAgentChartType(value, aliasType = "", range = null) {
  const source = String(value || aliasType || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    "": "",
    recommended: "",
    auto: "",
    chart: "",
    createchart: "",
    insertchart: "",
    addchart: "",
    column: "column",
    columns: "column",
    columnchart: "column",
    clusteredcolumn: "column",
    verticalbar: "column",
    bar: "bar",
    bars: "bar",
    barchart: "bar",
    horizontalbar: "bar",
    line: "line",
    lines: "line",
    linechart: "line",
    pie: "pie",
    piechart: "pie",
    doughnut: "pie",
    donut: "pie",
    scatter: "scatter",
    xy: "scatter",
    scatterchart: "scatter",
    map: "map",
    mapchart: "map",
    縦棒: "column",
    縦棒グラフ: "column",
    棒グラフ: "column",
    横棒: "bar",
    横棒グラフ: "bar",
    折れ線: "line",
    折れ線グラフ: "line",
    円: "pie",
    円グラフ: "pie",
    散布図: "scatter",
    地図: "map",
    マップ: "map",
    柱状图: "column",
    柱狀圖: "column",
    条形图: "bar",
    條形圖: "bar",
    折线图: "line",
    折線圖: "line",
    饼图: "pie",
    圓餅圖: "pie",
    散点图: "scatter",
    散點圖: "scatter",
    地图: "map",
    地圖: "map",
  };
  return aliases[source] || cwsAgentRecommendedChartTypeForRange(range);
}



function cwsAgentRecommendedChartTypeForRange(range) {
  if (!range) return "column";
  const rows = range.bottom - range.top + 1;
  const cols = range.right - range.left + 1;
  if (rows <= 2 && cols > 3) return "line";
  if (rows <= 8 && cols <= 2) return "pie";
  return "column";
}



function cwsAgentForecastPeriods(value) {
  if (value === null || value === undefined || value === "") return 5;
  const number = Number(value);
  return Number.isFinite(number) ? Math.max(1, Math.min(120, Math.round(number))) : 5;
}



function cwsAgentForecastChartType(value) {
  const text = String(value || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["column", "columns", "bar", "縦棒", "縦棒グラフ", "柱状图", "柱狀圖"].includes(text)) return "column";
  return "line";
}



function cwsAgentForecastConfidenceLevel(value) {
  if (value === null || value === undefined || value === "") return 95;
  const number = Number(value);
  return Number.isFinite(number) ? Math.max(0, Math.min(99, Math.round(number))) : 95;
}



function cwsAgentForecastSeasonality(value) {
  if (value === null || value === undefined || value === "") return 0;
  const number = Number(value);
  return Number.isFinite(number) ? Math.max(0, Math.min(365, Math.round(number))) : 0;
}



function cwsAgentForecastMissingData(value) {
  const text = String(value || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  return ["zero", "0", "zeros", "ゼロ", "零"].includes(text) ? "zero" : "interpolate";
}



function cwsAgentForecastDuplicateAggregate(value) {
  const text = String(value || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    "": "average",
    average: "average",
    avg: "average",
    mean: "average",
    平均: "average",
    sum: "sum",
    total: "sum",
    合計: "sum",
    合计: "sum",
    count: "count",
    counta: "count",
    件数: "count",
    min: "min",
    minimum: "min",
    最小: "min",
    max: "max",
    maximum: "max",
    最大: "max",
    median: "median",
    中央値: "median",
    中位数: "median",
  };
  return aliases[text] || "average";
}



function cwsAgentForecastEndValue(value) {
  if (value === null || value === undefined || value === "") return { number: null, text: "" };
  if (typeof value === "number" && Number.isFinite(value)) return { number: value, text: "" };
  const text = String(value).trim();
  if (!text) return { number: null, text: "" };
  const numeric = Number(text.replace(/,/g, ""));
  if (Number.isFinite(numeric)) return { number: numeric, text: "" };
  return { number: null, text: text.slice(0, 120) };
}



function cwsAgentGoalSeekCellForOp(op, index, label, cellKeys, rowKeys, colKeys, sheetIndexKeys, sheetNameKeys) {
  const cellText = String(cwsAgentFirstExistingValue(op, cellKeys)).trim();
  const parsedCell = cwsAgentParseCellAddress(cellText);
  const row = cwsAgentPositiveInteger(cwsAgentFirstExistingValue(op, rowKeys)) || parsedCell?.row || 0;
  const col = cwsAgentColumnValue(cwsAgentFirstExistingValue(op, colKeys)) || parsedCell?.col || 0;
  if (!row || !col || row > EXCEL_MAX_ROWS || col > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のゴール シーク${label}が正しくありません。`);
  }

  const byIndex = cwsAgentInteger(cwsAgentFirstExistingValue(op, sheetIndexKeys));
  const explicitSheetName = String(cwsAgentFirstExistingValue(op, sheetNameKeys)).trim();
  const sheetName = parsedCell?.sheetName || explicitSheetName;
  const sheetIndex = byIndex !== null
    ? byIndex
    : sheetName
      ? cwsAgentSheetIndexByName(sheetName)
      : cwsAgentSheetIndexForOp(op);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のゴール シーク${label}のシートが見つかりません。`);
  }
  return { sheetIndex, row, col };
}



function cwsAgentGoalSeekTargetValue(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = typeof value === "number"
    ? value
    : Number(String(value).trim().replace(/,/g, ""));
  if (!Number.isFinite(number) || Math.abs(number) > CWS_AGENT_MAX_GOAL_SEEK_ABS_TARGET) return null;
  return number;
}



function cwsAgentQuickAnalysisMode(value, aliasType = "") {
  const source = String(value || aliasType || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    databar: "dataBar",
    databars: "dataBar",
    dataバー: "dataBar",
    データバー: "dataBar",
    数据条: "dataBar",
    數據條: "dataBar",
    資料條: "dataBar",
    colorscale: "colorScale",
    colourscale: "colorScale",
    colorgradients: "colorScale",
    カラースケール: "colorScale",
    色スケール: "colorScale",
    色階: "colorScale",
    色阶: "colorScale",
    iconset: "iconSet",
    iconsets: "iconSet",
    アイコンセット: "iconSet",
    图标集: "iconSet",
    圖示集: "iconSet",
    greaterthanaverage: "greaterThanAverage",
    aboveaverage: "greaterThanAverage",
    averageabove: "greaterThanAverage",
    平均より大きい: "greaterThanAverage",
    平均以上: "greaterThanAverage",
    大于平均值: "greaterThanAverage",
    高于平均值: "greaterThanAverage",
    duplicatevalues: "duplicateValues",
    duplicates: "duplicateValues",
    duplicate: "duplicateValues",
    重複値: "duplicateValues",
    重複する値: "duplicateValues",
    重复值: "duplicateValues",
    重複值: "duplicateValues",
  };
  return aliases[source] || "";
}



function cwsAgentQuickAnalysisModeLabel(mode) {
  const labels = {
    dataBar: "データ バー",
    colorScale: "カラー スケール",
    iconSet: "アイコン セット",
    greaterThanAverage: "平均より大きい値",
    duplicateValues: "重複する値",
  };
  return labels[mode] || "クイック分析";
}



function cwsAgentQuickAnalysisTotalMethod(value, aliasType = "") {
  const source = String(value || aliasType || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    sum: "sum",
    total: "sum",
    totals: "sum",
    quickanalysissum: "sum",
    合計: "sum",
    合计: "sum",
    總計: "sum",
    average: "average",
    avg: "average",
    mean: "average",
    quickanalysisaverage: "average",
    平均: "average",
    count: "count",
    counta: "count",
    quickanalysiscount: "count",
    件数: "count",
    個数: "count",
    個數: "count",
    计数: "count",
    計數: "count",
    percent: "percent",
    percentage: "percent",
    percenttotal: "percent",
    quickanalysispercent: "percent",
    パーセント: "percent",
    割合: "percent",
    百分比: "percent",
    running: "running",
    runningtotal: "running",
    cumulativesum: "running",
    quickanalysisrunningtotal: "running",
    累計: "running",
    累计: "running",
    累積: "running",
  };
  return aliases[source] || "";
}



function cwsAgentChartTitle(value) {
  const text = String(value ?? "").trim();
  return text ? text.slice(0, 200) : "";
}



function cwsAgentChartSize(op = {}) {
  const width = cwsAgentClampedChartSize(op?.width ?? op?.chartWidth, 520);
  const height = cwsAgentClampedChartSize(op?.height ?? op?.chartHeight, 320);
  return { width, height };
}



function cwsAgentClampedChartSize(value, fallback) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return fallback;
  return Math.max(CWS_AGENT_MIN_CHART_SIZE, Math.min(CWS_AGENT_MAX_CHART_SIZE, Math.round(number)));
}



function cwsAgentChartAnchorForOp(op, sourceRange, index) {
  const anchorText = String(cwsAgentFirstExistingValue(op, ["anchor", "anchorCell", "target", "targetCell", "position", "at", "to"])).trim();
  const parsedCell = cwsAgentParseCellAddress(anchorText);
  const parsedRange = cwsAgentParseRangeAddress(anchorText);
  const parsedSheetName = parsedCell?.sheetName || parsedRange?.sheetName || "";
  if (parsedSheetName && cwsAgentSheetIndexByName(parsedSheetName) !== sourceRange.sheetIndex) {
    throw new Error(`Agent 操作 ${index + 1} のグラフ挿入位置は同じシート内で指定してください。`);
  }
  const row = cwsAgentPositiveInteger(op?.anchorRow ?? op?.targetRow ?? op?.toRow)
    || parsedCell?.row
    || parsedRange?.top
    || Math.max(1, sourceRange.top);
  const col = cwsAgentColumnValue(op?.anchorCol ?? op?.anchorColumn ?? op?.targetCol ?? op?.targetColumn ?? op?.toCol)
    || parsedCell?.col
    || parsedRange?.left
    || Math.min(EXCEL_MAX_COLS, sourceRange.right + 2);
  if (!row || !col || row > EXCEL_MAX_ROWS || col > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のグラフ挿入位置が正しくありません。`);
  }
  return { row, col };
}



function cwsAgentNormalizeSortHeaderText(value) {
  return String(value ?? "").trim().toLocaleLowerCase("ja-JP").replace(/\s+/g, "");
}



function cwsAgentSortDirection(value) {
  const text = String(value || "").trim().toLowerCase();
  return ["desc", "descending", "降順", "降序", "大きい順", "ztoa", "z-a"].includes(text) ? "desc" : "asc";
}



function cwsAgentHeaderMode(value) {
  if (value === null || value === undefined || value === "" || String(value).toLowerCase() === "auto") return null;
  return cwsAgentBoolean(value);
}



function cwsAgentFreezePaneMode(value) {
  const text = String(value ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    toprow: "topRow",
    firstrow: "topRow",
    headerrow: "topRow",
    header: "topRow",
    先頭行: "topRow",
    首行: "topRow",
    标题行: "topRow",
    標題行: "topRow",
    firstcolumn: "firstColumn",
    firstcol: "firstColumn",
    leftcolumn: "firstColumn",
    先頭列: "firstColumn",
    首列: "firstColumn",
    selection: "selection",
    selectedcell: "selection",
    cell: "selection",
    position: "selection",
    current: "selection",
  };
  return aliases[text] || "";
}



function cwsAgentSheetViewMode(value) {
  const text = String(value ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    normal: "normal",
    standard: "normal",
    default: "normal",
    grid: "normal",
    worksheet: "normal",
    標準: "normal",
    通常: "normal",
    普通: "normal",
    标准: "normal",
    標準表示: "normal",
    pagelayout: "page-layout",
    pagelayoutview: "page-layout",
    layout: "page-layout",
    pageview: "page-layout",
    ページレイアウト: "page-layout",
    ページレイアウト表示: "page-layout",
    页面布局: "page-layout",
    頁面佈局: "page-layout",
    pagebreak: "page-break",
    pagebreakview: "page-break",
    pagebreakpreview: "page-break",
    breakpreview: "page-break",
    改ページ: "page-break",
    改ページプレビュー: "page-break",
    分页预览: "page-break",
    分頁預覽: "page-break",
  };
  return aliases[text] || "";
}



function cwsAgentZoomScaleValue(value) {
  if (value === undefined || value === null || value === "") return null;
  const text = String(value).trim().replace(/%$/, "");
  if (!text || !Number.isFinite(Number(text))) return null;
  return clampZoom(Number(text));
}



function cwsAgentFreezePaneAddress(op) {
  return cwsAgentParseCellAddress(op?.splitCell || op?.at || op?.cell || op?.address || op?.ref || "");
}



function cwsAgentAutoFilterCriterionForOp(op, index) {
  const source = op?.criterion && typeof op.criterion === "object" ? op.criterion : op || {};
  const valuesSource = source.values ?? source.allowedValues ?? source.valueList;
  if (Array.isArray(valuesSource)) {
    const values = valuesSource.map(cwsAgentFilterValueText).filter((value) => value !== null);
    if (!values.length || values.length > CWS_AGENT_MAX_FILTER_VALUES) {
      throw new Error(`Agent 操作 ${index + 1} のフィルター値が正しくありません。`);
    }
    return { kind: "values", values };
  }

  const operator = cwsAgentAutoFilterOperator(source.operator ?? source.op ?? source.condition ?? source.type);
  if (operator === "blank" || operator === "nonBlank") {
    return { kind: "custom", operator1: operator, value1: "" };
  }

  const value = cwsAgentFilterValueText(source.value ?? source.value1 ?? source.equals ?? source.text ?? op?.value);
  if (value === null) {
    throw new Error(`Agent 操作 ${index + 1} のフィルター条件が正しくありません。`);
  }
  return {
    kind: "custom",
    operator1: operator || "equal",
    value1: value,
  };
}



function cwsAgentFilterValueText(value) {
  if (value === null || value === undefined || Array.isArray(value) || (typeof value === "object" && !(value instanceof Date))) return null;
  const text = value instanceof Date ? value.toISOString().slice(0, 10) : String(value);
  return text.length <= CWS_AGENT_MAX_FILTER_VALUE_TEXT ? text : null;
}



function cwsAgentDataValidationRule(op, index, options = {}) {
  const source = op?.validation && typeof op.validation === "object"
    ? op.validation
    : op?.rule && typeof op.rule === "object"
      ? op.rule
      : op || {};
  const type = cwsAgentDataValidationType(source.type ?? source.allow ?? source.kind ?? source.validationType);
  if (type !== "list") {
    throw new Error(`Agent 操作 ${index + 1} の入力規則タイプは list のみ対応しています。`);
  }

  const listOptions = cwsAgentDataValidationOptions(
    source.options ?? source.values ?? source.list ?? source.items ?? source.choices,
  );
  const formula = cwsAgentDataValidationFormula(
    source.formula ?? source.source ?? source.formula1 ?? source.listFormula,
  );
  const inlineOptions = formula ? parseValidationOptions(formula) : [];
  if (!listOptions.length && !inlineOptions.length && !validationReferenceFromFormula(options.sheet || activeSheet(), formula)) {
    throw new Error(`Agent 操作 ${index + 1} のリスト入力規則の選択肢が正しくありません。`);
  }

  return {
    type: "list",
    formula: formula || "",
    options: listOptions.length ? listOptions : inlineOptions,
    allowBlank: cwsAgentOptionalBoolean(source.allowBlank ?? source.ignoreBlank, true),
    showDropDown: cwsAgentOptionalBoolean(source.showDropDown ?? source.dropdown ?? source.inCellDropdown, true),
    showInputMessage: cwsAgentOptionalBoolean(source.showInputMessage, false),
    inputTitle: cwsAgentSmallText(source.inputTitle ?? source.promptTitle ?? ""),
    inputMessage: cwsAgentSmallText(source.inputMessage ?? source.prompt ?? ""),
    showErrorMessage: cwsAgentOptionalBoolean(source.showErrorMessage, true),
    errorStyle: cwsAgentDataValidationErrorStyle(source.errorStyle),
    errorTitle: cwsAgentSmallText(source.errorTitle ?? ""),
    errorMessage: cwsAgentSmallText(source.errorMessage ?? source.error ?? ""),
    imeMode: cwsAgentDataValidationImeMode(source.imeMode),
    ...(options.anchor ? { anchor: { row: options.anchor.row, col: options.anchor.col } } : {}),
  };
}



function cwsAgentDataValidationType(value) {
  const text = String(value || "list").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    "": "list",
    list: "list",
    dropdown: "list",
    dropdownlist: "list",
    pulldown: "list",
    choices: "list",
    options: "list",
    リスト: "list",
    ドロップダウン: "list",
    プルダウン: "list",
    下拉: "list",
    下拉菜单: "list",
    下拉選單: "list",
  };
  return aliases[text] || text;
}



function cwsAgentDataValidationOptions(value) {
  const source = Array.isArray(value)
    ? value
    : typeof value === "string"
      ? parseValidationOptions(value)
      : [];
  const options = [];
  const seen = new Set();
  for (const item of source) {
    if (item === null || item === undefined || typeof item === "object") continue;
    const text = String(item).trim();
    if (!text || text.length > CWS_AGENT_MAX_VALIDATION_OPTION_TEXT) continue;
    const key = text.toLocaleLowerCase("ja-JP");
    if (seen.has(key)) continue;
    seen.add(key);
    options.push(text);
    if (options.length >= CWS_AGENT_MAX_VALIDATION_OPTIONS) break;
  }
  return options;
}



function cwsAgentDataValidationFormula(value) {
  if (value === null || value === undefined || Array.isArray(value) || typeof value === "object") return "";
  return String(value).trim().slice(0, 2000);
}



function cwsAgentDataValidationErrorStyle(value) {
  const text = String(value || "").trim().toLowerCase();
  return ["warning", "information"].includes(text) ? text : "stop";
}



function cwsAgentDataValidationImeMode(value) {
  const text = String(value || "noControl").trim();
  const allowed = new Set(["noControl", "on", "off", "disabled", "hiragana", "fullKatakana", "halfKatakana", "fullAlpha", "halfAlpha"]);
  return allowed.has(text) ? text : "noControl";
}



function cwsAgentSmallText(value) {
  if (value === null || value === undefined || Array.isArray(value) || typeof value === "object") return "";
  return String(value).replace(/\r\n?/g, "\n").slice(0, 255);
}



function cwsAgentHyperlinkTarget(op, index) {
  const value = op?.hyperlink ?? op?.url ?? op?.href ?? op?.link ?? op?.target;
  if (value === null || value === undefined || Array.isArray(value) || typeof value === "object") {
    throw new Error(`Agent 操作 ${index + 1} のリンク先が正しくありません。`);
  }
  const raw = String(value).trim().slice(0, CWS_AGENT_MAX_HYPERLINK_TEXT);
  if (!raw) {
    throw new Error(`Agent 操作 ${index + 1} のリンク先が正しくありません。`);
  }
  const kind = cwsAgentHyperlinkKind(op?.kind ?? op?.linkType ?? op?.targetType, raw);
  const normalized = normalizeHyperlinkForKind(
    raw,
    kind,
    cwsAgentInteger(op?.documentSheetIndex ?? op?.targetSheetIndex ?? op?.sheetIndex) ?? state.activeSheetIndex,
  );
  if (!normalized || normalized.length > CWS_AGENT_MAX_HYPERLINK_TEXT) {
    throw new Error(`Agent 操作 ${index + 1} のリンク先が正しくありません。`);
  }
  return normalized;
}



function cwsAgentHyperlinkKind(value, raw) {
  const text = String(value ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["document", "internal", "sheet", "workbook", "cell", "セル", "内部", "文档", "文件"].includes(text)) return "document";
  if (["email", "mail", "mailto", "メール", "電子メール", "邮箱", "郵件"].includes(text)) return "email";
  const target = String(raw || "").trim();
  if (target.startsWith("#") || (/!/.test(target) && !/^[a-z][a-z0-9+.-]*:/i.test(target))) return "document";
  if (/^mailto:/i.test(target) || /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(target)) return "email";
  return "web";
}



function cwsAgentOptionalHyperlinkText(value) {
  if (value === null || value === undefined) return null;
  return cwsAgentCellValue(value);
}



function cwsAgentNoteText(value, index) {
  if (value === null || value === undefined || Array.isArray(value) || typeof value === "object") {
    throw new Error(`Agent 操作 ${index + 1} のメモ内容が正しくありません。`);
  }
  const text = cwsAgentCellValue(value).slice(0, CWS_AGENT_MAX_NOTE_TEXT);
  if (!text.trim()) {
    throw new Error(`Agent 操作 ${index + 1} のメモ内容が空です。`);
  }
  return text;
}



function cwsAgentCommentText(value, index) {
  if (value === null || value === undefined || Array.isArray(value) || typeof value === "object") {
    throw new Error(`Agent 操作 ${index + 1} のコメント内容が正しくありません。`);
  }
  const text = cwsAgentCellValue(value).slice(0, CWS_AGENT_MAX_COMMENT_TEXT);
  if (!text.trim()) {
    throw new Error(`Agent 操作 ${index + 1} のコメント内容が空です。`);
  }
  return text;
}



function cwsAgentDefinedNameText(value, index) {
  if (value === null || value === undefined || Array.isArray(value) || typeof value === "object") {
    throw new Error(`Agent 操作 ${index + 1} の名前が正しくありません。`);
  }
  const name = String(value).trim();
  const validation = validateDefinedName(name);
  if (!validation.ok) {
    throw new Error(`Agent 操作 ${index + 1}: ${validation.message}`);
  }
  return name;
}



function cwsAgentDefinedNameScope(op, allowAny = false) {
  const text = String(op?.scope ?? op?.nameScope ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (allowAny && ["any", "all", "すべて", "全部", "所有"].includes(text)) return "any";
  if (["sheet", "worksheet", "local", "localworksheet", "シート", "工作表", "本シート"].includes(text) || cwsAgentBoolean(op?.local ?? op?.sheetScoped ?? false)) {
    return "sheet";
  }
  return "workbook";
}



function cwsAgentDefinedNameScopeSheetIndex(op, fallbackSheetIndex, index) {
  const byIndex = cwsAgentInteger(op?.scopeSheetIndex ?? op?.localSheetIndex);
  const sheetIndex = byIndex !== null
    ? byIndex
    : cwsAgentSheetNameText(op?.scopeSheetName ?? op?.localSheetName)
      ? cwsAgentSheetIndexByName(cwsAgentSheetNameText(op?.scopeSheetName ?? op?.localSheetName))
      : fallbackSheetIndex;
  if (!state.model?.sheets?.[sheetIndex]) {
    throw new Error(`Agent 操作 ${index + 1} の名前スコープのシートが見つかりません。`);
  }
  return sheetIndex;
}



function cwsAgentSheetProtectionOptions(op) {
  const source = op?.options && typeof op.options === "object" ? { ...op, ...op.options } : op || {};
  return defaultSheetProtectionOptions({
    password: cwsAgentOptionalText(source.password ?? source.passcode ?? ""),
    allowSelectLockedCells: cwsAgentOptionalBoolean(source.allowSelectLockedCells ?? source.selectLockedCells, true),
    allowSelectUnlockedCells: cwsAgentOptionalBoolean(source.allowSelectUnlockedCells ?? source.selectUnlockedCells, true),
    allowFormatCells: cwsAgentOptionalBoolean(source.allowFormatCells ?? source.formatCells, false),
    allowFormatColumns: cwsAgentOptionalBoolean(source.allowFormatColumns ?? source.formatColumns, false),
    allowFormatRows: cwsAgentOptionalBoolean(source.allowFormatRows ?? source.formatRows, false),
    allowInsertColumns: cwsAgentOptionalBoolean(source.allowInsertColumns ?? source.insertColumns, false),
    allowInsertRows: cwsAgentOptionalBoolean(source.allowInsertRows ?? source.insertRows, false),
    allowDeleteColumns: cwsAgentOptionalBoolean(source.allowDeleteColumns ?? source.deleteColumns, false),
    allowDeleteRows: cwsAgentOptionalBoolean(source.allowDeleteRows ?? source.deleteRows, false),
    allowSort: cwsAgentOptionalBoolean(source.allowSort ?? source.sort, false),
    allowAutoFilter: cwsAgentOptionalBoolean(source.allowAutoFilter ?? source.autoFilter ?? source.filter, false),
    allowPivotTables: cwsAgentOptionalBoolean(source.allowPivotTables ?? source.pivotTables, false),
    allowEditObjects: cwsAgentOptionalBoolean(source.allowEditObjects ?? source.editObjects ?? source.objects, false),
    allowEditScenarios: cwsAgentOptionalBoolean(source.allowEditScenarios ?? source.editScenarios ?? source.scenarios, false),
  });
}



function cwsAgentCellProtectionPatch(op, aliasType = "") {
  const merged = op?.options && typeof op.options === "object" ? { ...op, ...op.options } : op || {};
  const source = merged?.protection && typeof merged.protection === "object" ? { ...merged, ...merged.protection } : merged;
  const patch = {};
  let locked = cwsAgentOptionalBoolSetting(source.locked ?? source.cellLocked ?? source.protected ?? source.protect);
  const unlocked = cwsAgentOptionalBoolSetting(source.unlocked ?? source.editable ?? source.allowEdit ?? source.allowEditing);
  if (locked === null && unlocked !== null) locked = !unlocked;
  if (aliasType === "lockcells") locked = true;
  if (aliasType === "unlockcells") locked = false;
  if (locked !== null) patch.locked = locked;

  let hidden = cwsAgentOptionalBoolSetting(source.hidden ?? source.hiddenFormula ?? source.formulaHidden ?? source.hideFormula ?? source.hideFormulas);
  if (aliasType === "hideformulas") hidden = true;
  if (aliasType === "showformulas" || aliasType === "unhideformulas") hidden = false;
  if (hidden !== null) patch.hidden = hidden;
  return patch;
}



function cwsAgentOptionalText(value) {
  if (value === null || value === undefined || Array.isArray(value) || typeof value === "object") return "";
  return String(value).replace(/\r\n?/g, "\n").slice(0, 255);
}



function cwsAgentOptionalBoolSetting(value) {
  if (value === null || value === undefined || value === "") return null;
  return cwsAgentBoolean(value);
}



function cwsAgentAutoFilterOperator(value) {
  const text = String(value ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    "": "equal",
    eq: "equal",
    equals: "equal",
    equal: "equal",
    is: "equal",
    等于: "equal",
    等於: "equal",
    等しい: "equal",
    一致: "equal",
    ne: "notEqual",
    notequal: "notEqual",
    notequals: "notEqual",
    不等于: "notEqual",
    不等於: "notEqual",
    等しくない: "notEqual",
    contains: "contains",
    include: "contains",
    includes: "contains",
    包含: "contains",
    含む: "contains",
    notcontains: "notContains",
    不包含: "notContains",
    含まない: "notContains",
    beginswith: "beginsWith",
    startswith: "beginsWith",
    で始まる: "beginsWith",
    notbeginswith: "notBeginsWith",
    endswith: "endsWith",
    で終わる: "endsWith",
    notendswith: "notEndsWith",
    gt: "greaterThan",
    greaterthan: "greaterThan",
    大于: "greaterThan",
    大於: "greaterThan",
    より大きい: "greaterThan",
    ">": "greaterThan",
    gte: "greaterThanOrEqual",
    ge: "greaterThanOrEqual",
    greaterthanorequal: "greaterThanOrEqual",
    以上: "greaterThanOrEqual",
    ">=": "greaterThanOrEqual",
    lt: "lessThan",
    lessthan: "lessThan",
    小于: "lessThan",
    小於: "lessThan",
    より小さい: "lessThan",
    "<": "lessThan",
    lte: "lessThanOrEqual",
    le: "lessThanOrEqual",
    lessthanorequal: "lessThanOrEqual",
    以下: "lessThanOrEqual",
    "<=": "lessThanOrEqual",
    blank: "blank",
    blanks: "blank",
    空白: "blank",
    nonblank: "nonBlank",
    nonblanks: "nonBlank",
    notblank: "nonBlank",
    非空: "nonBlank",
    空白以外: "nonBlank",
  };
  return aliases[text] || "equal";
}



function cwsAgentColumnWidthPixels(op) {
  const widthUnits = cwsAgentDimensionNumber(op?.widthUnits ?? op?.excelWidth);
  if (widthUnits !== null) {
    return cwsAgentClampDimension(columnWidthToPixels(widthUnits), CWS_AGENT_MIN_COLUMN_WIDTH, CWS_AGENT_MAX_COLUMN_WIDTH);
  }
  return cwsAgentDimensionPixels(op?.width ?? op?.pixels ?? op?.px ?? op?.size, {
    min: CWS_AGENT_MIN_COLUMN_WIDTH,
    max: CWS_AGENT_MAX_COLUMN_WIDTH,
  });
}



function cwsAgentRowHeightPixels(op) {
  const heightPt = cwsAgentDimensionNumber(op?.heightPt ?? op?.points);
  if (heightPt !== null) {
    return cwsAgentClampDimension(rowHeightToPixels(heightPt), CWS_AGENT_MIN_ROW_HEIGHT, CWS_AGENT_MAX_ROW_HEIGHT);
  }
  return cwsAgentDimensionPixels(op?.height ?? op?.pixels ?? op?.px ?? op?.size, {
    min: CWS_AGENT_MIN_ROW_HEIGHT,
    max: CWS_AGENT_MAX_ROW_HEIGHT,
    allowPt: true,
  });
}



function cwsAgentDimensionPixels(value, { min, max, allowPt = false } = {}) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  const match = text.match(/^([0-9]+(?:\.[0-9]+)?)\s*(px|pt)?$/i);
  if (!match) return null;
  let number = Number(match[1]);
  if (!Number.isFinite(number)) return null;
  if (allowPt && String(match[2] || "").toLowerCase() === "pt") {
    number = rowHeightToPixels(number);
  } else if (match[2] && String(match[2]).toLowerCase() !== "px") {
    return null;
  }
  return cwsAgentClampDimension(number, min, max);
}



function cwsAgentDimensionNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : null;
}



function cwsAgentClampDimension(value, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return null;
  return Math.round(Math.max(min, Math.min(max, number)));
}



function cwsAgentParseCellAddress(value) {
  const text = String(value || "").trim();
  if (!text) return null;
  const parts = text.split("!");
  const cellText = parts.pop().replace(/\$/g, "").trim();
  const match = cellText.match(/^([A-Za-z]{1,3})([1-9][0-9]{0,6})$/);
  if (!match) return null;

  const sheetPart = parts.join("!").trim();
  let sheetName = "";
  if (sheetPart) {
    sheetName = sheetPart.replace(/^'|'$/g, "").replace(/''/g, "'");
  }
  return {
    sheetName,
    row: Number(match[2]),
    col: cwsAgentColumnIndex(match[1]),
  };
}



function cwsAgentParseRangeAddress(value) {
  const text = String(value || "").trim();
  if (!text) return null;
  const parts = text.split("!");
  const rangeText = parts.pop().replace(/\$/g, "").trim();
  const rangeParts = rangeText.split(":", 2);
  const sheetPart = parts.join("!").trim();
  const start = cwsAgentParseCellAddress([sheetPart, rangeParts[0]].filter(Boolean).join("!"));
  const end = cwsAgentParseCellAddress([sheetPart, rangeParts[1] || rangeParts[0]].filter(Boolean).join("!"));
  if (!start || !end) return null;
  return {
    sheetName: start.sheetName || end.sheetName,
    top: Math.min(start.row, end.row),
    left: Math.min(start.col, end.col),
    bottom: Math.max(start.row, end.row),
    right: Math.max(start.col, end.col),
  };
}



function cwsAgentParseColumnRangeAddress(value) {
  const text = String(value || "").trim();
  if (!text) return null;
  const parts = text.split("!");
  const rangeText = parts.pop().replace(/\$/g, "").trim();
  const rangeParts = rangeText.split(":", 2).map((item) => item.trim());
  if (!/^[A-Za-z]{1,3}$/.test(rangeParts[0] || "") || !/^[A-Za-z]{1,3}$/.test(rangeParts[1] || rangeParts[0] || "")) return null;
  const startCol = cwsAgentColumnIndex(rangeParts[0]);
  const endCol = cwsAgentColumnIndex(rangeParts[1] || rangeParts[0]);
  if (!startCol || !endCol || startCol > EXCEL_MAX_COLS || endCol > EXCEL_MAX_COLS) return null;
  const sheetPart = parts.join("!").trim();
  return {
    sheetName: sheetPart ? sheetPart.replace(/^'|'$/g, "").replace(/''/g, "'") : "",
    left: Math.min(startCol, endCol),
    right: Math.max(startCol, endCol),
  };
}



function cwsAgentParseRowRangeAddress(value) {
  const text = String(value || "").trim();
  if (!text) return null;
  const parts = text.split("!");
  const rangeText = parts.pop().replace(/\$/g, "").trim();
  const rangeParts = rangeText.split(":", 2).map((item) => item.trim());
  if (!/^[1-9][0-9]{0,6}$/.test(rangeParts[0] || "") || !/^[1-9][0-9]{0,6}$/.test(rangeParts[1] || rangeParts[0] || "")) return null;
  const startRow = Number(rangeParts[0]);
  const endRow = Number(rangeParts[1] || rangeParts[0]);
  if (startRow > EXCEL_MAX_ROWS || endRow > EXCEL_MAX_ROWS) return null;
  const sheetPart = parts.join("!").trim();
  return {
    sheetName: sheetPart ? sheetPart.replace(/^'|'$/g, "").replace(/''/g, "'") : "",
    top: Math.min(startRow, endRow),
    bottom: Math.max(startRow, endRow),
  };
}



function cwsAgentColumnIndex(name) {
  return String(name || "").toUpperCase().split("").reduce((total, char) => {
    const code = char.charCodeAt(0);
    if (code < 65 || code > 90) return total;
    return total * 26 + code - 64;
  }, 0);
}



function cwsAgentResolveOperationForApply(op) {
  const deferred = cwsAgentOperationSources.get(op);
  if (!deferred) return op;
  // Earlier operations may create, rename or reorder sheets. Resolve the target
  // against the workbook at execution time, not the pre-replay sheet inventory.
  const resolved = normalizeCwsAgentOp(deferred.source, deferred.index);
  for (const key of Object.keys(op)) delete op[key];
  Object.assign(op, resolved);
  cwsAgentOperationSources.delete(op);
  return op;
}



function applyCwsAgentOp(op) {
  op = cwsAgentResolveOperationForApply(op);
  if (op.type === "setCell") {
    return applyCwsAgentSetCellOp(op);
  }
  if (op.type === "setRange") {
    return applyCwsAgentSetRangeOp(op);
  }
  if (op.type === "setStyle") {
    return applyCwsAgentSetStyleOp(op);
  }
  if (op.type === "setCellProtection") {
    return applyCwsAgentSetCellProtectionOp(op, false);
  }
  if (op.type === "clearCellProtection") {
    return applyCwsAgentSetCellProtectionOp(op, true);
  }
  if (op.type === "setSparkline") {
    return applyCwsAgentSetSparklineOp(op);
  }
  if (op.type === "clearSparkline") {
    return applyCwsAgentClearSparklineOp(op);
  }
  if (op.type === "setBorders") {
    return applyCwsAgentSetBordersOp(op);
  }
  if (op.type === "setNumberFormat") {
    return applyCwsAgentSetNumberFormatOp(op);
  }
  if (op.type === "sortRange") {
    return applyCwsAgentSortRangeOp(op);
  }
  if (op.type === "setAutoFilter") {
    return applyCwsAgentSetAutoFilterOp(op);
  }
  if (op.type === "clearAutoFilter") {
    return applyCwsAgentClearAutoFilterOp(op);
  }
  if (op.type === "setFilterCriteria") {
    return applyCwsAgentSetFilterCriteriaOp(op);
  }
  if (op.type === "clearFilterCriteria") {
    return applyCwsAgentClearFilterCriteriaOp(op);
  }
  if (op.type === "advancedFilter") {
    return applyCwsAgentAdvancedFilterOp(op);
  }
  if (op.type === "removeDuplicates") {
    return applyCwsAgentRemoveDuplicatesOp(op);
  }
  if (op.type === "findReplace") {
    return applyCwsAgentFindReplaceOp(op);
  }
  if (op.type === "textToColumns") {
    return applyCwsAgentTextToColumnsOp(op);
  }
  if (op.type === "createTable") {
    return applyCwsAgentCreateTableOp(op);
  }
  if (op.type === "createChart") {
    return applyCwsAgentCreateChartOp(op);
  }
  if (op.type === "createPivotTable") {
    return applyCwsAgentCreatePivotTableOp(op);
  }
  if (op.type === "createForecastSheet") {
    return applyCwsAgentCreateForecastSheetOp(op);
  }
  if (op.type === "goalSeek") {
    return applyCwsAgentGoalSeekOp(op);
  }
  if (op.type === "dataTable") {
    return applyCwsAgentDataTableOp(op);
  }
  if (op.type === "addShape") {
    return applyCwsAgentAddShapeOp(op);
  }
  if (op.type === "setShapeStyle" || op.type === "updateShape" || op.type === "removeShape") return applyCwsAgentExistingShapeEditOp(op);
  if (op.type === "setShapeText") return applyCwsAgentSetShapeTextOp(op);
  if (op.type === "createWorkbookTemplate") {
    return applyCwsAgentCreateWorkbookTemplateOp(op);
  }
  if (op.type === "setQuickAnalysis") {
    return applyCwsAgentSetQuickAnalysisOp(op);
  }
  if (op.type === "clearQuickAnalysis") {
    return applyCwsAgentClearQuickAnalysisOp(op);
  }
  if (op.type === "setQuickAnalysisTotal") {
    return applyCwsAgentSetQuickAnalysisTotalOp(op);
  }
  if (op.type === "setSubtotal") {
    return applyCwsAgentSetSubtotalOp(op);
  }
  if (op.type === "clearSubtotals") {
    return applyCwsAgentClearSubtotalsOp(op);
  }
  if (op.type === "consolidateRanges") {
    return applyCwsAgentConsolidateOp(op);
  }
  if (op.type === "setDataValidation") {
    return applyCwsAgentSetDataValidationOp(op);
  }
  if (op.type === "clearDataValidation") {
    return applyCwsAgentClearDataValidationOp(op);
  }
  if (op.type === "setHyperlink") {
    return applyCwsAgentSetHyperlinkOp(op);
  }
  if (op.type === "clearHyperlink") {
    return applyCwsAgentClearHyperlinkOp(op);
  }
  if (op.type === "setNote") {
    return applyCwsAgentSetNoteOp(op);
  }
  if (op.type === "clearNote") {
    return applyCwsAgentClearNoteOp(op);
  }
  if (op.type === "setComment") {
    return applyCwsAgentSetCommentOp(op, false);
  }
  if (op.type === "appendComment") {
    return applyCwsAgentSetCommentOp(op, true);
  }
  if (op.type === "clearComment") {
    return applyCwsAgentClearCommentOp(op);
  }
  if (op.type === "setNamedRange") {
    return applyCwsAgentSetNamedRangeOp(op);
  }
  if (op.type === "clearNamedRange") {
    return applyCwsAgentClearNamedRangeOp(op);
  }
  if (op.type === "protectSheet") {
    return applyCwsAgentProtectSheetOp(op);
  }
  if (op.type === "unprotectSheet") {
    return applyCwsAgentUnprotectSheetOp(op);
  }
  if (op.type === "activateSheet") {
    return applyCwsAgentActivateSheetOp(op);
  }
  if (op.type === "selectRange") {
    return applyCwsAgentSelectRangeOp(op);
  }
  if (op.type === "copyRange" || op.type === "moveRange") {
    return applyCwsAgentCopyMoveRangeOp(op);
  }
  if (op.type === "addSheet") {
    return applyCwsAgentAddSheetOp(op);
  }
  if (op.type === "renameSheet") {
    return applyCwsAgentRenameSheetOp(op);
  }
  if (op.type === "duplicateSheet") {
    return applyCwsAgentDuplicateSheetOp(op);
  }
  if (op.type === "moveSheet") {
    return applyCwsAgentMoveSheetOp(op);
  }
  if (op.type === "setSheetTabColor") {
    return applyCwsAgentSheetTabColorOp(op);
  }
  if (op.type === "deleteSheet") {
    return applyCwsAgentDeleteSheetOp(op);
  }
  if (op.type === "hideSheet" || op.type === "unhideSheet") {
    return applyCwsAgentSheetVisibilityOp(op, op.type === "hideSheet");
  }
  if (op.type === "setFreezePanes") {
    return applyCwsAgentSetFreezePanesOp(op);
  }
  if (op.type === "clearFreezePanes") {
    return applyCwsAgentClearFreezePanesOp(op);
  }
  if (op.type === "setSheetView") {
    return applyCwsAgentSetSheetViewOp(op);
  }
  if (op.type === "setPrintArea") {
    return applyCwsAgentSetPrintAreaOp(op);
  }
  if (op.type === "clearPrintArea") {
    return applyCwsAgentClearPrintAreaOp(op);
  }
  if (op.type === "setPrintTitles") {
    return applyCwsAgentSetPrintTitlesOp(op);
  }
  if (op.type === "clearPrintTitles") {
    return applyCwsAgentClearPrintTitlesOp(op);
  }
  if (op.type === "setHeaderFooter") {
    return applyCwsAgentSetHeaderFooterOp(op);
  }
  if (op.type === "clearHeaderFooter") {
    return applyCwsAgentClearHeaderFooterOp(op);
  }
  if (op.type === "setPageTheme") {
    return applyCwsAgentSetPageThemeOp(op);
  }
  if (op.type === "setPageOrientation") {
    return applyCwsAgentSetPageOrientationOp(op);
  }
  if (op.type === "setPaperSize") {
    return applyCwsAgentSetPaperSizeOp(op);
  }
  if (op.type === "setPageMargins") {
    return applyCwsAgentSetPageMarginsOp(op);
  }
  if (op.type === "setPageScale") {
    return applyCwsAgentSetPageScaleOp(op);
  }
  if (op.type === "setPrintOptions") {
    return applyCwsAgentSetPrintOptionsOp(op);
  }
  if (op.type === "insertPageBreak") {
    return applyCwsAgentPageBreakOp(op, true);
  }
  if (op.type === "removePageBreak") {
    return applyCwsAgentPageBreakOp(op, false);
  }
  if (op.type === "resetPageBreaks") {
    return applyCwsAgentResetPageBreaksOp(op);
  }
  if (op.type === "setColumnWidth") {
    return applyCwsAgentSetColumnWidthOp(op);
  }
  if (op.type === "setRowHeight") {
    return applyCwsAgentSetRowHeightOp(op);
  }
  if (op.type === "autoFitColumns") {
    return applyCwsAgentAutoFitColumnsOp(op);
  }
  if (op.type === "autoFitRows") {
    return applyCwsAgentAutoFitRowsOp(op);
  }
  if (op.type === "hideRows" || op.type === "unhideRows") {
    return applyCwsAgentHeaderHiddenOp(op, "row", op.type === "hideRows");
  }
  if (op.type === "hideColumns" || op.type === "unhideColumns") {
    return applyCwsAgentHeaderHiddenOp(op, "column", op.type === "hideColumns");
  }
  if (op.type === "groupRows" || op.type === "ungroupRows" || op.type === "groupColumns" || op.type === "ungroupColumns") {
    return applyCwsAgentOutlineOp(op);
  }
  if (op.type === "insertRows" || op.type === "deleteRows") {
    return applyCwsAgentRowsOp(op);
  }
  if (op.type === "insertColumns" || op.type === "deleteColumns") {
    return applyCwsAgentColumnsOp(op);
  }
  if (op.type === "clearRange") {
    return applyCwsAgentClearRangeOp(op);
  }
  if (op.type === "mergeCells") {
    return applyCwsAgentMergeCellsOp(op);
  }
  if (op.type === "unmergeCells") {
    return applyCwsAgentUnmergeCellsOp(op);
  }
  throw new Error(`対応していない Agent 操作です: ${op.type || ""}`);
}



function applyCwsAgentSetCellOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const before = getCellRawInput(sheet, op.row, op.col);
  const coerced = cwsAgentCoerceInputForCell(sheet, op.row, op.col, op.value);
  if (coerced.numFmt !== '@') coerced.content = coerceCellInput(sheet, op.row, op.col, op.value, { preserveWhitespace: true });
  commitCellValue(op.row, op.col, op.value, {
    preserveWhitespace: true,
    skipSheetGroup: true,
    agentInput: coerced,
  });

  const updatedSheet = state.model.sheets[op.sheetIndex];
  const after = getCellRawInput(updatedSheet, op.row, op.col);
  if (String(after ?? "") !== String(coerced.content ?? "") && String(before ?? "") !== String(coerced.content ?? "")) {
    throw new Error(`${columnName(op.col)}${op.row} を変更できませんでした。`);
  }

  return {
    type: "setCell",
    sheetIndex: op.sheetIndex,
    sheetName: updatedSheet.name,
    row: op.row,
    col: op.col,
    address: `${updatedSheet.name}!${columnName(op.col)}${op.row}`,
  };
}



function applyCwsAgentSetRangeOp(op) {
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
  if (!ensureEditableRangeForProtection(sheet, range, "変更")) {
    throw new Error("保護されたセルは変更できません。");
  }

  ensureSheetSize(sheet, range.bottom, range.right);
  const sheetId = getSheetId(sheet.name);
  const writes = [];
  for (let rowOffset = 0; rowOffset < op.values.length; rowOffset += 1) {
    const rowValues = op.values[rowOffset] || [];
    for (let colOffset = 0; colOffset < rowValues.length; colOffset += 1) {
      const row = range.top + rowOffset;
      const col = range.left + colOffset;
      const coerced = cwsAgentCoerceInputForCell(sheet, row, col, rowValues[colOffset]);
      const content = coerced.content;
      const validation = validateCellInput(sheet, row, col, content);
      if (!validation.valid) {
        throw new Error(`${columnName(col)}${row}: ${validation.message || "入力規則に違反しています。"}`);
      }
      writes.push({ row, col, content, numFmt: coerced.numFmt });
    }
  }

  const sheetBefore = captureSheetForHistory(sheet);
  writes.forEach((item) => {
    if (item.numFmt) {
      const cell = ensureCellModel(sheet, item.row, item.col);
      cell.numFmt = item.numFmt;
      storeCellModelForSheet(sheet, cellKey(item.row, item.col), cell);
    }
    writeCellContent(sheet, sheetId, item.row, item.col, item.content, { preserveWhitespace: true });
  });
  internCellsInRange(sheet, range);
  refreshWorkbookChartsAfterSourceEdit(op.sheetIndex, range);
  commitDataOperationForSheet(
    op.sheetIndex,
    sheet,
    sheetBefore,
    range,
    `${rangeToLabel(range)} CWS Agent 一括入力`,
    `CWS Agent: ${sheet.name}!${rangeToLabel(range)} を更新しました。`,
    range,
  );

  return {
    type: "setRange",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function cwsAgentCoerceInputForCell(sheet, row, col, value) {
  // JSON strings with significant leading zeros/unsafe integer precision are
  // identifiers. Preserve them identically in initial ranges and later patches.
  const existing = String(sheet?.cells?.[cellKey(row, col)]?.numFmt || "");
  if (typeof value === "string" && (existing === "@" || /^[+-]?0\d+$/.test(value) || (/^[+-]?\d+$/.test(value) && !Number.isSafeInteger(Number(value))))) {
    return { content: value, numFmt: "@" };
  }
  const content = coerceUserInput(value, { preserveWhitespace: true });
  const formatted = parseFormattedNumericInput(value);
  if (!formatted) return { content, numFmt: "" };
  const cell = sheet?.cells?.[cellKey(row, col)] || null;
  const existingFormat = String(cell?.numFmt || "");
  const header = cwsAgentNearestColumnHeaderText(sheet, row, col);
  const currency = formatted.currency || currencySymbol(existingFormat) || (cwsAgentLooksLikeCurrencyHeader(header) ? "¥" : "");
  const shouldCoerce = typeof content !== "number" || currencySymbol(existingFormat) || cwsAgentLooksLikeNumericHeader(header);
  if (!shouldCoerce) return { content, numFmt: "" };
  let numFmt = "";
  if (!existingFormat && currency) {
    numFmt = `${currency}#,##0${Number.isInteger(formatted.number) ? "" : ".00"}`;
  } else if (!existingFormat && formatted.usedGrouping && cwsAgentLooksLikeNumericHeader(header)) {
    numFmt = "#,##0";
  }
  return { content: formatted.number, numFmt };
}



function cwsAgentNearestColumnHeaderText(sheet, row, col) {
  if (!sheet || !row || !col) return "";
  const top = Math.max(1, row - 12);
  for (let currentRow = row - 1; currentRow >= top; currentRow -= 1) {
    const cell = sheet.cells?.[cellKey(currentRow, col)];
    const text = String(cell?.raw ?? cell?.display ?? (cell ? cellDisplayText(cell) : "") ?? "").trim();
    if (text) return text;
  }
  return "";
}



function cwsAgentLooksLikeCurrencyHeader(value) {
  return /金額|金额|金钱|金錢|通貨|货币|貨幣|金銭|金钱|料金|費用|费用|単価|单价|価格|价格|税込|税抜|税額|消費税|小計|小计|合計|合计|total|amount|price|cost|fee|subtotal|tax/i.test(String(value || ""));
}



function cwsAgentLooksLikeNumericHeader(value) {
  return cwsAgentLooksLikeCurrencyHeader(value) || /数量|數量|個数|件数|回数|qty|quantity|count|number/i.test(String(value || ""));
}



function applyCwsAgentSetStyleOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではセルの書式を変更できません。")) {
    throw new Error("保護されたシートでは書式を変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);
  const before = captureCellsForHistory(sheet, range);
  applyCssPatchToSheetRanges(sheet, [range], op.style);
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent 書式`,
  });
  renderWorkbook();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} に書式を適用しました。`);

  return {
    type: "setStyle",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSetCellProtectionOp(op, clearProtection = false) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートではセル保護設定を変更できません。")) {
    throw new Error("保護されたシートではセル保護設定を変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);
  const before = captureCellsForHistory(sheet, range);
  const hasLocked = Object.prototype.hasOwnProperty.call(op.protection || {}, "locked");
  const hasHidden = Object.prototype.hasOwnProperty.call(op.protection || {}, "hidden");

  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const cell = clearProtection ? (sheet.cells[key] || null) : ensureCellModel(sheet, row, col);
    if (!cell) return;
    if (clearProtection) {
      delete cell.protection;
    } else {
      const current = cell.protection || {};
      const locked = hasLocked ? Boolean(op.protection.locked) : current.locked !== false;
      const hidden = hasHidden ? Boolean(op.protection.hidden) : Boolean(current.hidden);
      if (locked && !hidden) delete cell.protection;
      else cell.protection = { locked, hidden };
    }
    storeCellModelForSheet(sheet, key, cell);
    removeCellModelIfEmpty(sheet, row, col);
  });
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent セル保護`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} のセル保護設定を${clearProtection ? "既定に戻し" : "変更し"}ました。`);

  return {
    type: clearProtection ? "clearCellProtection" : "setCellProtection",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSetSparklineOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const sheetId = getSheetId(sheet.name);
  const sourceRange = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.sourceTop, col: op.sourceLeft },
    { sheetIndex: op.sheetIndex, row: op.sourceBottom, col: op.sourceRight },
  );
  const targetRange = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.targetTop, col: op.targetLeft },
    { sheetIndex: op.sheetIndex, row: op.targetBottom, col: op.targetRight },
  );
  if (!ensureEditableRangeForProtection(sheet, targetRange, "スパークラインを作成")) {
    throw new Error("保護されたセルにはスパークラインを作成できません。");
  }

  ensureSheetSize(sheet, Math.max(sourceRange.bottom, targetRange.bottom), Math.max(sourceRange.right, targetRange.right));
  const series = [];
  if (op.axis === "columns") {
    for (let col = sourceRange.left; col <= sourceRange.right; col += 1) {
      const values = [];
      for (let row = sourceRange.top; row <= sourceRange.bottom; row += 1) {
        values.push(cellNumericValue(sheet, sheetId, row, col));
      }
      series.push({
        row: targetRange.top,
        col: targetRange.left + col - sourceRange.left,
        values,
        sourceRange: rangeToLabel({ sheetIndex: op.sheetIndex, top: sourceRange.top, left: col, bottom: sourceRange.bottom, right: col }),
      });
    }
  } else {
    for (let row = sourceRange.top; row <= sourceRange.bottom; row += 1) {
      const values = [];
      for (let col = sourceRange.left; col <= sourceRange.right; col += 1) {
        values.push(cellNumericValue(sheet, sheetId, row, col));
      }
      series.push({
        row: targetRange.top + row - sourceRange.top,
        col: targetRange.left,
        values,
        sourceRange: rangeToLabel({ sheetIndex: op.sheetIndex, top: row, left: sourceRange.left, bottom: row, right: sourceRange.right }),
      });
    }
  }

  const before = captureCellsForHistory(sheet, targetRange);
  series.forEach((item) => {
    const cell = ensureCellModel(sheet, item.row, item.col);
    cell.sparkline = {
      type: op.sparklineType,
      values: item.values,
      sourceRange: item.sourceRange,
    };
    cell.raw = null;
    cell.display = "";
    storeCellModelForSheet(sheet, cellKey(item.row, item.col), cell);
  });
  internCellsInRange(sheet, targetRange);
  const after = captureCellsForHistory(sheet, targetRange);

  state.selected = { sheetIndex: op.sheetIndex, row: targetRange.top, col: targetRange.left };
  state.selectionRange = targetRange;
  state.selectionRanges = [targetRange];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: sourceRange,
    selectionAfter: targetRange,
    selectionBeforeRanges: [sourceRange],
    selectionAfterRanges: [targetRange],
    label: `${rangeToLabel(targetRange)} CWS Agent スパークライン`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(targetRange)} にスパークラインを作成しました。`);

  return {
    type: "setSparkline",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    sparklineType: op.sparklineType,
    axis: op.axis,
    sourceAddress: `${sheet.name}!${rangeToLabel(sourceRange)}`,
    address: `${sheet.name}!${rangeToLabel(targetRange)}`,
  };
}



function applyCwsAgentClearSparklineOp(op) {
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
  if (!ensureEditableRangeForProtection(sheet, range, "スパークラインを削除")) {
    throw new Error("保護されたセルのスパークラインは削除できません。");
  }

  const before = captureCellsForHistory(sheet, range);
  let cleared = 0;
  forEachCellInRange(range, (row, col) => {
    const key = cellKey(row, col);
    const cell = sheet.cells[key];
    if (cell?.sparkline) {
      delete cell.sparkline;
      cleared += 1;
      storeCellModelForSheet(sheet, key, cell);
    }
  });
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);

  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent スパークライン削除`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} のスパークラインを削除しました。`);

  return {
    type: "clearSparkline",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    cleared,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSetBordersOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートでは罫線を変更できません。")) {
    throw new Error("保護されたシートでは罫線を変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);
  const before = captureCellsForHistory(sheet, range);
  applyBorderPresetToSheetRange(sheet, range, op.preset);
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);
  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent 罫線`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} に罫線を適用しました。`);

  return {
    type: "setBorders",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    preset: op.preset,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSetNumberFormatOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowFormatCells", "保護されたシートでは表示形式を変更できません。")) {
    throw new Error("保護されたシートでは表示形式を変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);
  const before = captureCellsForHistory(sheet, range);
  forEachCellInRange(range, (row, col) => {
    const cell = ensureCellModel(sheet, row, col);
    applyCellNumberFormat(cell, op.numFmt);
  });
  internCellsInRange(sheet, range);
  const after = captureCellsForHistory(sheet, range);
  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: range.left };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    before,
    after,
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent 表示形式：${op.label}`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} に表示形式を適用しました：${op.label}`);

  return {
    type: "setNumberFormat",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    numFmt: op.numFmt,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSortRangeOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowSort", "保護されたシートでは並べ替えを使用できません。")) {
    throw new Error("保護されたシートでは並べ替えを使用できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);
  const hasHeader = op.hasHeader === null ? inferHeaderRow(sheet, range) : Boolean(op.hasHeader);
  const sheetBefore = captureSheetForHistory(sheet);
  const sorted = sortRangeRowsByKeys(
    range,
    [{ col: op.keyCol, direction: op.direction }],
    hasHeader,
    { sheet, commitHistory: false, quiet: true },
  );
  if (!sorted) {
    throw new Error("指定範囲を並べ替えできませんでした。");
  }
  state.hf = buildFormulaEngine(state.model);
  state.selected = { sheetIndex: op.sheetIndex, row: range.top, col: op.keyCol };
  state.selectionRange = range;
  state.selectionRanges = [range];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  pushHistory({
    sheetIndex: op.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: range,
    selectionAfter: range,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [range],
    label: `${rangeToLabel(range)} CWS Agent ${op.direction === "desc" ? "降順" : "昇順"}並べ替え`,
  });
  renderWorkbook();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} を${columnName(op.keyCol)}列で${op.direction === "desc" ? "降順" : "昇順"}に並べ替えました。`);

  return {
    type: "sortRange",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    keyCol: op.keyCol,
    direction: op.direction,
    hasHeader,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentSetAutoFilterOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではオートフィルターを変更できません。")) {
    throw new Error("保護されたシートではオートフィルターを変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  ensureSheetSize(sheet, range.bottom, range.right);
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.autoFilter = { range: { ...range }, criteria: {} };
  syncAutoFilterStateToMatchingTable(sheet);
  reapplyAutoFilter({ silent: true, sheet, sheetIndex: op.sheetIndex });
  commitDataOperationForSheet(
    op.sheetIndex,
    sheet,
    sheetBefore,
    range,
    `${rangeToLabel(range)} CWS Agent フィルター`,
    autoFilterStatusMessage(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} にフィルターを設定しました`, sheet, range),
    range,
  );

  return {
    type: "setAutoFilter",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentClearAutoFilterOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではオートフィルターを変更できません。")) {
    throw new Error("保護されたシートではオートフィルターを変更できません。");
  }
  const currentRange = sheet.autoFilter?.range ? { ...sheet.autoFilter.range, sheetIndex: op.sheetIndex } : null;
  const requestedRange = op.hasRange
    ? normalizeRange(
        { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
        { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
      )
    : null;
  if (requestedRange && currentRange && !rangeRectEquals(requestedRange, currentRange)) {
    throw new Error("指定範囲にはフィルターが設定されていません。");
  }
  const targetRange = currentRange || requestedRange || { sheetIndex: op.sheetIndex, top: 1, left: 1, bottom: 1, right: 1 };
  if (!currentRange) {
    state.selected = { sheetIndex: op.sheetIndex, row: targetRange.top, col: targetRange.left };
    state.selectionRange = targetRange;
    state.selectionRanges = [targetRange];
    state.selectionStart = state.selected;
    renderWorkbook();
    updateFormulaBarForSelection();
    setStatus("解除するフィルターがありません。");
    return {
      type: "clearAutoFilter",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      top: targetRange.top,
      left: targetRange.left,
      bottom: targetRange.bottom,
      right: targetRange.right,
      address: `${sheet.name}!${rangeToLabel(targetRange)}`,
    };
  }

  const sheetBefore = captureSheetForHistory(sheet);
  removeAutoFilter(sheet);
  commitDataOperationForSheet(
    op.sheetIndex,
    sheet,
    sheetBefore,
    currentRange,
    `${rangeToLabel(currentRange)} CWS Agent フィルター解除`,
    `CWS Agent: ${sheet.name}!${rangeToLabel(currentRange)} のフィルターを解除しました。`,
    currentRange,
  );

  return {
    type: "clearAutoFilter",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: currentRange.top,
    left: currentRange.left,
    bottom: currentRange.bottom,
    right: currentRange.right,
    address: `${sheet.name}!${rangeToLabel(currentRange)}`,
  };
}



function applyCwsAgentSetFilterCriteriaOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではオートフィルターを変更できません。")) {
    throw new Error("保護されたシートではオートフィルターを変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  if (op.col < range.left || op.col > range.right) {
    throw new Error("フィルター列が範囲外です。");
  }
  ensureSheetSize(sheet, range.bottom, range.right);
  const sheetBefore = captureSheetForHistory(sheet);
  if (!sheet.autoFilter?.range || !rangeRectEquals(sheet.autoFilter.range, range)) {
    sheet.autoFilter = { range: { ...range }, criteria: {} };
  } else {
    sheet.autoFilter.criteria = sheet.autoFilter.criteria || {};
  }
  sheet.autoFilter.criteria[op.col] = op.criterion;
  delete sheet.autoFilter.advancedCriteria;
  syncAutoFilterStateToMatchingTable(sheet);
  reapplyAutoFilter({ silent: true, sheet, sheetIndex: op.sheetIndex });
  commitDataOperationForSheet(
    op.sheetIndex,
    sheet,
    sheetBefore,
    range,
    `${rangeToLabel(range)} CWS Agent フィルター条件`,
    autoFilterStatusMessage(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} の ${columnName(op.col)} 列にフィルター条件を適用しました`, sheet, range),
    { ...range, left: op.col, right: op.col },
  );

  return {
    type: "setFilterCriteria",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    col: op.col,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentClearFilterCriteriaOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(sheet, "allowAutoFilter", "保護されたシートではオートフィルターを変更できません。")) {
    throw new Error("保護されたシートではオートフィルターを変更できません。");
  }
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  if (!sheet.autoFilter?.range) {
    throw new Error("クリアするフィルター条件がありません。");
  }
  if (op.hasRange && !rangeRectEquals(sheet.autoFilter.range, range)) {
    throw new Error("指定範囲にはフィルターが設定されていません。");
  }
  const sheetBefore = captureSheetForHistory(sheet);
  sheet.autoFilter.criteria = sheet.autoFilter.criteria || {};
  if (op.col) {
    delete sheet.autoFilter.criteria[op.col];
  } else {
    sheet.autoFilter.criteria = {};
    delete sheet.autoFilter.advancedCriteria;
  }
  syncAutoFilterStateToMatchingTable(sheet);
  reapplyAutoFilter({ silent: true, sheet, sheetIndex: op.sheetIndex });
  commitDataOperationForSheet(
    op.sheetIndex,
    sheet,
    sheetBefore,
    range,
    `${rangeToLabel(range)} CWS Agent フィルター条件クリア`,
    autoFilterStatusMessage(op.col ? `CWS Agent: ${columnName(op.col)} 列のフィルター条件をクリアしました` : "CWS Agent: フィルター条件をクリアしました", sheet, range),
    op.col ? { ...range, left: op.col, right: op.col } : range,
  );

  return {
    type: "clearFilterCriteria",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    col: op.col,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentAdvancedFilterOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const listSheet = state.model.sheets[op.sheetIndex];
  if (!ensureSheetProtectionPermission(listSheet, "allowAutoFilter", "保護されたシートでは詳細設定フィルターを使用できません。")) {
    throw new Error("保護されたシートでは詳細設定フィルターを使用できません。");
  }
  const listRange = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  if (listRange.bottom <= listRange.top) {
    throw new Error("詳細設定フィルターには見出し行と1行以上のデータが必要です。");
  }
  ensureSheetSize(listSheet, listRange.bottom, listRange.right);

  const criteriaRange = op.criteriaRange
    ? normalizeRange(
        { sheetIndex: op.criteriaRange.sheetIndex, row: op.criteriaRange.top, col: op.criteriaRange.left },
        { sheetIndex: op.criteriaRange.sheetIndex, row: op.criteriaRange.bottom, col: op.criteriaRange.right },
      )
    : null;
  if (criteriaRange) {
    const criteriaSheet = sheetForWorkbookRange(criteriaRange);
    if (!criteriaSheet || !isSheetVisible(criteriaSheet)) {
      throw new Error("検索条件範囲のシートが見つかりません。");
    }
    const criteriaIssues = advancedFilterCriteriaHeaderIssues(listSheet, listRange, criteriaRange);
    if (criteriaIssues.unmatched.length) {
      throw new Error(`検索条件範囲の列見出しがリスト範囲に見つかりません: ${criteriaIssues.unmatched.join(", ")}`);
    }
  }

  const settings = {
    criteriaRange: criteriaRange ? { ...criteriaRange } : null,
    uniqueOnly: op.uniqueOnly === true,
  };
  const rows = advancedFilterMatchingRows(listSheet, listRange, settings);
  if (op.action === "copy") {
    const rawOutputRange = op.copyToRange;
    const outputRange = normalizeRange(
      { sheetIndex: rawOutputRange.sheetIndex, row: rawOutputRange.top, col: rawOutputRange.left },
      { sheetIndex: rawOutputRange.sheetIndex, row: rawOutputRange.bottom, col: rawOutputRange.right },
    );
    const outputSheet = sheetForWorkbookRange(outputRange);
    if (!outputSheet || !isSheetVisible(outputSheet)) {
      throw new Error("抽出範囲のシートが見つかりません。");
    }
    const outputDetails = advancedFilterOutputColumnDetails(listSheet, listRange, outputSheet, outputRange);
    if (outputDetails.unmatched.length) {
      throw new Error(`抽出範囲の列見出しがリスト範囲に見つかりません: ${outputDetails.unmatched.join(", ")}`);
    }
    const outputRows = advancedFilterMatchingRows(listSheet, listRange, {
      ...settings,
      uniqueColumns: outputDetails.cols,
    });
    const targetInfo = advancedFilterCopyTargetInfo(outputRange, outputDetails.cols.length, outputRows.length, outputDetails.usesHeaderRow);
    if (!ensureEditableRangeForProtection(outputSheet, targetInfo.targetRange, "抽出")) {
      throw new Error("保護されたセルには抽出できません。");
    }
    if (!op.overwrite && targetInfo.overwriteRange && rangeHasUserData(outputSheet, targetInfo.overwriteRange)) {
      throw new Error("抽出先に既存データがあります。上書きする場合は overwrite:true を指定してください。");
    }
    const copied = copyAdvancedFilterResults(listSheet, listRange, outputSheet, outputRange, settings, { skipOverwriteCheck: true });
    if (!copied) {
      throw new Error("詳細設定フィルターの抽出に失敗しました。");
    }
    return {
      type: "advancedFilter",
      action: "copy",
      sheetIndex: op.sheetIndex,
      sheetName: listSheet.name,
      top: listRange.top,
      left: listRange.left,
      bottom: listRange.bottom,
      right: listRange.right,
      criteriaRange,
      uniqueOnly: settings.uniqueOnly,
      rows: outputRows.length,
      targetSheetIndex: outputRange.sheetIndex,
      targetSheetName: outputSheet.name,
      targetTop: targetInfo.targetRange.top,
      targetLeft: targetInfo.targetRange.left,
      targetBottom: targetInfo.targetRange.bottom,
      targetRight: targetInfo.targetRange.right,
      address: `${listSheet.name}!${rangeToLabel(listRange)}`,
      targetAddress: `${outputSheet.name}!${rangeToLabel(targetInfo.targetRange)}`,
    };
  }

  applyAdvancedFilterInPlace(listSheet, listRange, settings);
  return {
    type: "advancedFilter",
    action: "in-place",
    sheetIndex: op.sheetIndex,
    sheetName: listSheet.name,
    top: listRange.top,
    left: listRange.left,
    bottom: listRange.bottom,
    right: listRange.right,
    criteriaRange,
    uniqueOnly: settings.uniqueOnly,
    rows: rows.length,
    address: `${listSheet.name}!${rangeToLabel(listRange)}`,
  };
}



function applyCwsAgentRemoveDuplicatesOp(op) {
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
  if (!ensureEditableRangeForProtection(sheet, range, "重複の削除")) {
    throw new Error("保護されたセルを含む範囲では重複を削除できません。");
  }
  const mergeConflict = removeDuplicatesMergeConflict(sheet, range, op.hasHeader);
  if (mergeConflict) {
    throw new Error(mergeConflict);
  }
  const collected = collectRemoveDuplicateRows(sheet, range, op.columns, op.hasHeader);
  const selectionAfter = range;
  state.selected = { sheetIndex: op.sheetIndex, row: selectionAfter.top, col: selectionAfter.left };
  state.selectionRange = selectionAfter;
  state.selectionRanges = [selectionAfter];
  state.selectionStart = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];

  if (!collected.removed) {
    renderSheet();
    updateSelectionUi();
    updateFormulaBarForSelection();
    setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} に重複行はありませんでした。`);
    return {
      type: "removeDuplicates",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      top: range.top,
      left: range.left,
      bottom: range.bottom,
      right: range.right,
      columns: op.columns,
      hasHeader: op.hasHeader,
      removed: 0,
      remaining: collected.uniqueRows.length,
      address: `${sheet.name}!${rangeToLabel(range)}`,
    };
  }

  const sheetBefore = captureSheetForHistory(sheet);
  applyCollectedRemoveDuplicateRows(sheet, range, collected);
  commitDataOperationForSheet(
    op.sheetIndex,
    sheet,
    sheetBefore,
    range,
    `${rangeToLabel(range)} CWS Agent 重複の削除`,
    `CWS Agent: ${sheet.name}!${rangeToLabel(range)} から ${collected.removed} 行の重複を削除しました。`,
    selectionAfter,
  );

  return {
    type: "removeDuplicates",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    columns: op.columns,
    hasHeader: op.hasHeader,
    removed: collected.removed,
    remaining: collected.uniqueRows.length,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentFindReplaceOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const sheetId = getSheetId(sheet.name);
  const range = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const writes = [];
  for (let row = range.top; row <= range.bottom; row += 1) {
    for (let col = range.left; col <= range.right; col += 1) {
      if (!cwsAgentCellMatchesFindQuery(sheet, sheetId, row, col, op.find, op)) continue;
      const source = getCellRawInput(sheet, row, col);
      const nextValue = replaceCellText(source, op.find, op.replacement, {
        caseSensitive: op.caseSensitive,
        wholeCell: op.wholeCell,
        replaceAllInCell: op.replaceAllInCell,
      });
      if (nextValue === source) continue;
      const validation = validateCellInput(sheet, row, col, nextValue);
      if (!validation.valid) {
        throw new Error(`${columnName(col)}${row}: ${validation.message || "入力規則に違反しています。"}`);
      }
      writes.push({ row, col, value: nextValue });
    }
  }

  if (!writes.length) {
    setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} に置換対象はありませんでした。`);
    return {
      type: "findReplace",
      sheetIndex: op.sheetIndex,
      sheetName: sheet.name,
      top: range.top,
      left: range.left,
      bottom: range.bottom,
      right: range.right,
      replaced: 0,
      address: `${sheet.name}!${rangeToLabel(range)}`,
    };
  }

  const changedRanges = writes.map((item) => ({
    sheetIndex: op.sheetIndex,
    top: item.row,
    left: item.col,
    bottom: item.row,
    right: item.col,
  }));
  if (!ensureEditableRangesForProtection(sheet, changedRanges, "置換")) {
    throw new Error("保護されたセルを含む範囲では置換できません。");
  }

  const sheetBefore = captureSheetForHistory(sheet);
  writes.forEach((item) => writeCellContent(sheet, sheetId, item.row, item.col, item.value, { preserveWhitespace: true }));
  state.hf = buildFormulaEngine(state.model);
  const last = writes[writes.length - 1];
  const baseSelectedRange = normalizeRange(
    { sheetIndex: op.sheetIndex, row: last.row, col: last.col },
    { sheetIndex: op.sheetIndex, row: last.row, col: last.col },
  );
  const selectedRange = expandRangeForMerges(sheet, baseSelectedRange) || baseSelectedRange;
  state.selected = { sheetIndex: op.sheetIndex, row: selectedRange.top, col: selectedRange.left };
  state.selectionRange = selectedRange;
  state.selectionRanges = [selectedRange];
  state.selectionStart = state.selected;
  state.selectionAnchor = state.selected;
  state.selectionDrag = null;
  state.selectedImage = null;
  state.selectedImages = [];
  pushHistory({
    sheetIndex: op.sheetIndex,
    sheetBefore,
    sheetAfter: captureSheetForHistory(sheet),
    selectionBefore: range,
    selectionAfter: selectedRange,
    selectionBeforeRanges: [range],
    selectionAfterRanges: [selectedRange],
    label: `${rangeToLabel(range)} CWS Agent 置換`,
  });
  renderSheet();
  updateSelectionUi();
  updateFormulaBarForSelection();
  setStatus(`CWS Agent: ${sheet.name}!${rangeToLabel(range)} で ${writes.length} 件を置換しました。`);

  return {
    type: "findReplace",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    replaced: writes.length,
    address: `${sheet.name}!${rangeToLabel(range)}`,
  };
}



function applyCwsAgentTextToColumnsOp(op) {
  if (state.activeSheetIndex !== op.sheetIndex) {
    setActiveSheet(op.sheetIndex);
  }
  if (state.activeSheetIndex !== op.sheetIndex) {
    throw new Error("対象シートを開けませんでした。");
  }

  const sheet = state.model.sheets[op.sheetIndex];
  const sourceRange = normalizeRange(
    { sheetIndex: op.sheetIndex, row: op.top, col: op.left },
    { sheetIndex: op.sheetIndex, row: op.bottom, col: op.right },
  );
  const destination = { row: op.destinationRow, col: op.destinationCol };
  const rows = textToColumnsRowsForRange(sheet, sourceRange, op.options);
  const outputRows = textColumnsOutputRows(rows, op.options);
  if (!outputRows.some((row) => row.length)) {
    throw new Error("分列の出力列がありません。列のデータ形式を確認してください。");
  }
  const outputHeight = outputRows.length;
  const outputWidth = Math.max(1, ...outputRows.map((row) => row.length));
  if (destination.row + outputHeight - 1 > EXCEL_MAX_ROWS || destination.col + outputWidth - 1 > EXCEL_MAX_COLS) {
    throw new Error("分列の出力範囲がシートの上限を超えています。");
  }
  if (outputHeight * outputWidth > CWS_AGENT_MAX_TEXT_TO_COLUMNS_CELLS) {
    throw new Error("分列の出力範囲が大きすぎます。");
  }
  const targetRange = normalizeRange(
    { sheetIndex: op.sheetIndex, row: destination.row, col: destination.col },
    { sheetIndex: op.sheetIndex, row: destination.row + outputHeight - 1, col: destination.col + outputWidth - 1 },
  );
  if (!ensureEditableRangeForProtection(sheet, targetRange, "分割")) {
    throw new Error("保護されたセルを含む範囲では分列できません。");
  }
  if (!op.overwrite && rangeHasUserData(sheet, targetRange, sourceRange)) {
    throw new Error("分列の表示先には既存データがあります。上書きする場合は明示してください。");
  }

  for (let rowOffset = 0; rowOffset < outputHeight; rowOffset += 1) {
    const parts = outputRows[rowOffset] || [];
    for (let colOffset = 0; colOffset < outputWidth; colOffset += 1) {
      const item = parts[colOffset] || { value: "", format: "general" };
      const row = targetRange.top + rowOffset;
      const col = targetRange.left + colOffset;
      const content = coerceTextColumnsValue(item.value, item.format, op.options);
      const validation = validateCellInput(sheet, row, col, content);
      if (!validation.valid) {
        throw new Error(`${columnName(col)}${row}: ${validation.message || "入力規則に違反しています。"}`);
      }
    }
  }

  const result = textToColumns(sourceRange, op.options, destination, {
    targetRange,
    outputRows,
    skipOverwriteCheck: true,
  });
  if (!result) {
    throw new Error("分列を実行できませんでした。");
  }

  return {
    type: "textToColumns",
    sheetIndex: op.sheetIndex,
    sheetName: sheet.name,
    top: sourceRange.top,
    left: sourceRange.left,
    bottom: sourceRange.bottom,
    right: sourceRange.right,
    destinationRow: destination.row,
    destinationCol: destination.col,
    outputRows: outputHeight,
    outputColumns: outputWidth,
    sourceAddress: `${sheet.name}!${rangeToLabel(sourceRange)}`,
    address: `${sheet.name}!${rangeToLabel(targetRange)}`,
  };
}

export {
  applyCwsAgentAdvancedFilterOp,
  applyCwsAgentClearAutoFilterOp,
  applyCwsAgentClearFilterCriteriaOp,
  applyCwsAgentClearSparklineOp,
  applyCwsAgentFindReplaceOp,
  applyCwsAgentOp,
  applyCwsAgentRemoveDuplicatesOp,
  applyCwsAgentSetAutoFilterOp,
  applyCwsAgentSetBordersOp,
  applyCwsAgentSetCellOp,
  applyCwsAgentSetCellProtectionOp,
  applyCwsAgentSetFilterCriteriaOp,
  applyCwsAgentSetNumberFormatOp,
  applyCwsAgentSetRangeOp,
  applyCwsAgentSetSparklineOp,
  applyCwsAgentSetStyleOp,
  applyCwsAgentSortRangeOp,
  applyCwsAgentTextToColumnsOp,
  cwsAgentAutoFilterCriterionForOp,
  cwsAgentAutoFilterOperator,
  cwsAgentCellProtectionPatch,
  cwsAgentChartAnchorForOp,
  cwsAgentChartRangeForOp,
  cwsAgentChartSize,
  cwsAgentChartTitle,
  cwsAgentChartType,
  cwsAgentClampDimension,
  cwsAgentClampedChartSize,
  cwsAgentCoerceInputForCell,
  cwsAgentColumnIndex,
  cwsAgentColumnWidthPixels,
  cwsAgentCommentText,
  cwsAgentCreateTableRangeForOp,
  cwsAgentDataValidationErrorStyle,
  cwsAgentDataValidationFormula,
  cwsAgentDataValidationImeMode,
  cwsAgentDataValidationOptions,
  cwsAgentDataValidationRule,
  cwsAgentDataValidationType,
  cwsAgentDefinedNameScope,
  cwsAgentDefinedNameScopeSheetIndex,
  cwsAgentDefinedNameText,
  cwsAgentDimensionNumber,
  cwsAgentDimensionPixels,
  cwsAgentFilterValueText,
  cwsAgentForecastChartType,
  cwsAgentForecastConfidenceLevel,
  cwsAgentForecastDuplicateAggregate,
  cwsAgentForecastEndValue,
  cwsAgentForecastMissingData,
  cwsAgentForecastPeriods,
  cwsAgentForecastRangeForOp,
  cwsAgentForecastSeasonality,
  cwsAgentFreezePaneAddress,
  cwsAgentFreezePaneMode,
  cwsAgentGoalSeekCellForOp,
  cwsAgentGoalSeekTargetValue,
  cwsAgentHeaderMode,
  cwsAgentHyperlinkKind,
  cwsAgentHyperlinkTarget,
  cwsAgentLineEndpointSide,
  cwsAgentLooksLikeCurrencyHeader,
  cwsAgentLooksLikeNumericHeader,
  cwsAgentNearestColumnHeaderText,
  cwsAgentNormalizeSortHeaderText,
  cwsAgentNoteText,
  cwsAgentOptionalBoolSetting,
  cwsAgentOptionalHyperlinkText,
  cwsAgentOptionalText,
  cwsAgentParseCellAddress,
  cwsAgentParseColumnRangeAddress,
  cwsAgentParseRangeAddress,
  cwsAgentParseRowRangeAddress,
  cwsAgentPivotRangeForOp,
  cwsAgentQuickAnalysisMode,
  cwsAgentQuickAnalysisModeLabel,
  cwsAgentQuickAnalysisRangeForOp,
  cwsAgentQuickAnalysisTotalMethod,
  cwsAgentRecommendedChartTypeForRange,
  cwsAgentResolveOperationForApply,
  cwsAgentRowHeightPixels,
  cwsAgentShapeArrowTypeValue,
  cwsAgentShapeBoxForOp,
  cwsAgentShapeColor,
  cwsAgentShapeFontSize,
  cwsAgentShapeLinePoints,
  cwsAgentShapeLocationForOp,
  cwsAgentShapeObjectName,
  cwsAgentShapeOutlineWidth,
  cwsAgentShapePosition,
  cwsAgentShapeSize,
  cwsAgentShapeStyleForOp,
  cwsAgentShapeText,
  cwsAgentShapeType,
  cwsAgentSheetProtectionOptions,
  cwsAgentSheetViewMode,
  cwsAgentSingleCharacter,
  cwsAgentSmallText,
  cwsAgentSortDirection,
  cwsAgentTableName,
  cwsAgentTableStyle,
  cwsAgentTextToColumnsColumnFormat,
  cwsAgentTextToColumnsColumnFormats,
  cwsAgentTextToColumnsDelimiterToken,
  cwsAgentTextToColumnsDelimiters,
  cwsAgentTextToColumnsDestinationForOp,
  cwsAgentTextToColumnsFixedBreaks,
  cwsAgentTextToColumnsMode,
  cwsAgentTextToColumnsOptions,
  cwsAgentTextToColumnsRangeForOp,
  cwsAgentUniqueTableName,
  cwsAgentZoomScaleValue,
};
