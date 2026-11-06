import { cwsNormalizeExistingShapeEdit, cwsExistingShapeEditTarget } from './existing-shape-model.js';
// Feature module; shared bindings are imported explicitly.
import {
  CWS_AGENT_BORDER_PRESETS,
  CWS_AGENT_MAX_ADVANCED_FILTER_CELLS,
  CWS_AGENT_MAX_CONSOLIDATE_CELLS,
  CWS_AGENT_MAX_COPY_CELLS,
  CWS_AGENT_MAX_DATA_TABLE_CELLS,
  CWS_AGENT_MAX_DEDUP_CELLS,
  CWS_AGENT_MAX_FILTER_CELLS,
  CWS_AGENT_MAX_FIND_REPLACE_CELLS,
  CWS_AGENT_MAX_FREEZE_COLUMNS,
  CWS_AGENT_MAX_FREEZE_ROWS,
  CWS_AGENT_MAX_RANGE_CELLS,
  CWS_AGENT_MAX_SORT_CELLS,
  CWS_AGENT_MAX_SPARKLINE_SOURCE_CELLS,
  CWS_AGENT_MAX_STRUCTURE_COLUMNS,
  CWS_AGENT_MAX_STRUCTURE_ROWS,
  CWS_AGENT_MAX_VALIDATION_CELLS,
  DATA_SUMMARY_FUNCTIONS,
  EXCEL_MAX_COLS,
  EXCEL_MAX_ROWS,
  PAGE_LAYOUT_EFFECT_PRESETS,
  PAGE_LAYOUT_MARGIN_PRESETS,
  PAGE_LAYOUT_THEMES,
  state,
} from "../../app-core.js";
import {
  cwsAgentAutoFilterCriterionForOp,
  cwsAgentChartAnchorForOp,
  cwsAgentChartRangeForOp,
  cwsAgentChartSize,
  cwsAgentChartTitle,
  cwsAgentChartType,
  cwsAgentColumnIndex,
  cwsAgentColumnWidthPixels,
  cwsAgentCommentText,
  cwsAgentCreateTableRangeForOp,
  cwsAgentDataValidationRule,
  cwsAgentDefinedNameScope,
  cwsAgentDefinedNameScopeSheetIndex,
  cwsAgentDefinedNameText,
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
  cwsAgentHyperlinkTarget,
  cwsAgentLineEndpointSide,
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
  cwsAgentQuickAnalysisRangeForOp,
  cwsAgentQuickAnalysisTotalMethod,
  cwsAgentRowHeightPixels,
  cwsAgentShapeBoxForOp,
  cwsAgentShapeLocationForOp,
  cwsAgentShapeObjectName,
  cwsAgentShapeStyleForOp,
  cwsAgentShapeType,
  cwsAgentSheetProtectionOptions,
  cwsAgentSheetViewMode,
  cwsAgentSmallText,
  cwsAgentTableName,
  cwsAgentTableStyle,
  cwsAgentTextToColumnsDestinationForOp,
  cwsAgentTextToColumnsOptions,
  cwsAgentTextToColumnsRangeForOp,
  cwsAgentZoomScaleValue,
} from "./operations.js";
import {
  inferHeaderRow,
  sheetUsedRange,
} from "../data/sort-and-autofilter.js";
import {
  decodeWorkbookRangeOrName,
} from "../data/validation-and-consolidation.js";
import {
  isLineShapeType,
  shapePlainTextFromRuns,
} from "../drawing/shape-rendering.js";
import {
  normalizeColorValue,
} from "../editing/ink-and-clipboard.js";
import {
  getDisplayForCell,
} from "../editing/pointers-and-formula-evaluation.js";
import {
  rangesIntersect,
} from "../rendering/grid-and-canvas.js";
import {
  currencySymbol,
  isDateFormat,
} from "../ui/help-and-utilities.js";
import {
  absoluteRangeReference,
  normalizeDefinedNameRefersTo,
} from "../ui/ribbon-and-links.js";
import {
  uniqueCopiedSheetName,
  uniqueSheetName,
  validateSheetName,
} from "../ui/sheet-protection-and-translation.js";
import {
  isSheetVisible,
  verticalAlignToFlex,
  visibleSheetIndexes,
} from "../workbook/formulas-and-sheet-tabs.js";
import {
  getSheetId,
} from "../workbook/saving-and-export.js";
import {
  normalizePageSetupPrintTitleColumns,
  normalizePageSetupPrintTitleRows,
} from "../workbook/vba-and-printing.js";


function normalizeCwsAgentSetAutoFilterOp(op, index) {
  return {
    type: "setAutoFilter",
    ...cwsAgentFilterRangeForOp(op, index),
  };
}



function normalizeCwsAgentClearAutoFilterOp(op, index) {
  const hasRange = Boolean(
    String(op?.range || op?.address || op?.ref || op?.cell || "").trim() ||
    op?.top || op?.row || op?.left || op?.col || op?.bottom || op?.right,
  );
  if (hasRange) {
    return {
      type: "clearAutoFilter",
      ...cwsAgentFilterRangeForOp(op, index),
      hasRange: true,
    };
  }
  const sheetIndex = cwsAgentSheetIndexForOp(op);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: "clearAutoFilter",
    sheetIndex,
    hasRange: false,
  };
}



function normalizeCwsAgentSetFilterCriteriaOp(op, index) {
  const range = cwsAgentFilterRangeForOp(op, index, { allowCurrent: true });
  const sheet = state.model?.sheets?.[range.sheetIndex];
  const col = cwsAgentFilterKeyColumn(op, sheet, range);
  if (!col || col < range.left || col > range.right) {
    throw new Error(`Agent 操作 ${index + 1} のフィルター列が正しくありません。`);
  }
  const criterion = cwsAgentAutoFilterCriterionForOp(op, index);
  return {
    type: "setFilterCriteria",
    ...range,
    col,
    criterion,
  };
}



function normalizeCwsAgentClearFilterCriteriaOp(op, index) {
  const hasRange = Boolean(
    String(op?.range || op?.address || op?.ref || op?.cell || "").trim() ||
    op?.top || op?.row || op?.left || op?.bottom || op?.right,
  );
  const range = cwsAgentFilterRangeForOp(op, index, { allowCurrent: true });
  const sheet = state.model?.sheets?.[range.sheetIndex];
  const col = cwsAgentFilterKeyColumn(op, sheet, range, { optional: true });
  if (col && (col < range.left || col > range.right)) {
    throw new Error(`Agent 操作 ${index + 1} のフィルター列が正しくありません。`);
  }
  return {
    type: "clearFilterCriteria",
    ...range,
    col: col || null,
    hasRange,
  };
}



function normalizeCwsAgentAdvancedFilterOp(op, index) {
  const rawListValue = cwsAgentFirstExistingValue(op, ["listRange", "range", "sourceRange", "dataRange", "databaseRange", "source", "address", "ref"]);
  const listValue = rawListValue !== null && rawListValue !== undefined && (typeof rawListValue === "object" || String(rawListValue).trim())
    ? rawListValue
    : op;
  const listRange = cwsAgentAdvancedFilterRangeForValue(listValue, op, index, "リスト範囲");
  if (listRange.bottom <= listRange.top) {
    throw new Error(`Agent 操作 ${index + 1} の詳細設定フィルターには見出し行と1行以上のデータが必要です。`);
  }
  const criteriaValue = cwsAgentFirstExistingValue(op, ["criteriaRange", "criteria", "conditionRange", "filterCriteriaRange"]);
  const hasCriteriaRange = criteriaValue !== null && criteriaValue !== undefined && (typeof criteriaValue === "object" || String(criteriaValue).trim());
  const criteriaRange = hasCriteriaRange
    ? cwsAgentAdvancedFilterRangeForValue(criteriaValue, op, index, "検索条件範囲")
    : null;
  const copyValue = cwsAgentFirstExistingValue(op, ["copyToRange", "destination", "destinationRange", "targetRange", "outputRange", "extractRange"]);
  const action = cwsAgentAdvancedFilterAction(op, copyValue);
  const hasCopyRange = copyValue !== null && copyValue !== undefined && (typeof copyValue === "object" || String(copyValue).trim());
  if (action === "copy" && !hasCopyRange) {
    throw new Error(`Agent 操作 ${index + 1} の抽出範囲を指定してください。`);
  }
  const copyToRange = action === "copy"
    ? cwsAgentAdvancedFilterRangeForValue(copyValue, op, index, "抽出範囲")
    : null;
  return {
    type: "advancedFilter",
    sheetIndex: listRange.sheetIndex,
    top: listRange.top,
    left: listRange.left,
    bottom: listRange.bottom,
    right: listRange.right,
    criteriaRange,
    action,
    copyToRange,
    uniqueOnly: cwsAgentOptionalBoolean(op?.uniqueOnly ?? op?.unique ?? op?.distinct, false),
    overwrite: cwsAgentOptionalBoolean(op?.overwrite ?? op?.replaceExisting ?? op?.replace, false),
  };
}



function normalizeCwsAgentRemoveDuplicatesOp(op, index) {
  const range = cwsAgentRemoveDuplicatesRangeForOp(op, index);
  const sheet = state.model?.sheets?.[range.sheetIndex];
  const headerMode = cwsAgentHeaderMode(op?.hasHeader ?? op?.header ?? op?.headers ?? op?.useHeader);
  const hasHeader = headerMode === null ? inferHeaderRow(sheet, range) : headerMode;
  const columns = cwsAgentRemoveDuplicateColumns(op, sheet, range, hasHeader, index);
  return {
    type: "removeDuplicates",
    ...range,
    hasHeader,
    columns,
  };
}



function normalizeCwsAgentFindReplaceOp(op, index) {
  const findText = cwsAgentRequiredFindText(op?.find ?? op?.query ?? op?.search ?? op?.searchText ?? op?.from ?? op?.oldText ?? op?.oldValue, index);
  const replacement = cwsAgentReplacementText(
    cwsAgentFirstExistingValue(op, ["replace", "replacement", "replaceWith", "to", "newText", "newValue", "value"]),
  );
  return {
    type: "findReplace",
    ...cwsAgentFindReplaceRangeForOp(op, index),
    find: findText,
    replacement,
    caseSensitive: cwsAgentBoolean(op?.caseSensitive ?? op?.matchCase),
    wholeCell: cwsAgentBoolean(op?.wholeCell ?? op?.matchEntireCell ?? op?.entireCell),
    replaceAllInCell: cwsAgentOptionalBoolean(op?.replaceAllInCell ?? op?.allInCell, true),
    lookIn: cwsAgentFindReplaceLookIn(op?.lookIn ?? op?.target ?? op?.scopeTarget),
  };
}



function normalizeCwsAgentTextToColumnsOp(op, index) {
  const range = cwsAgentTextToColumnsRangeForOp(op, index);
  const destination = cwsAgentTextToColumnsDestinationForOp(op, range, index);
  const options = cwsAgentTextToColumnsOptions(op, index);
  return {
    type: "textToColumns",
    ...range,
    destinationRow: destination.row,
    destinationCol: destination.col,
    options,
    overwrite: cwsAgentOptionalBoolean(op?.overwrite ?? op?.allowOverwrite ?? op?.replaceExisting ?? op?.skipOverwriteCheck, false),
  };
}



function normalizeCwsAgentCreateTableOp(op, index) {
  const range = cwsAgentCreateTableRangeForOp(op, index);
  const sheet = state.model?.sheets?.[range.sheetIndex];
  const headerMode = cwsAgentHeaderMode(op?.hasHeaders ?? op?.hasHeader ?? op?.header ?? op?.headers ?? op?.useHeader);
  const hasHeaders = headerMode === null ? inferHeaderRow(sheet, range) : headerMode;
  return {
    type: "createTable",
    ...range,
    hasHeaders,
    name: cwsAgentTableName(op?.name ?? op?.tableName ?? op?.displayName),
    showFilterButton: cwsAgentOptionalBoolean(op?.showFilterButton ?? op?.filterButton ?? op?.filters, true),
    style: cwsAgentTableStyle(op?.style ?? op?.tableStyle),
  };
}



function normalizeCwsAgentCreateChartOp(op, index, aliasType = "") {
  const range = cwsAgentChartRangeForOp(op, index);
  const chartType = cwsAgentChartType(op?.chartType ?? op?.chartKind ?? op?.kind ?? op?.subtype, aliasType, range);
  const anchor = cwsAgentChartAnchorForOp(op, range, index);
  const size = cwsAgentChartSize(op);
  return {
    type: "createChart",
    ...range,
    chartType,
    anchorRow: anchor.row,
    anchorCol: anchor.col,
    title: cwsAgentChartTitle(op?.title ?? op?.name ?? op?.chartTitle),
    width: size.width,
    height: size.height,
  };
}



function normalizeCwsAgentCreatePivotTableOp(op, index) {
  const range = cwsAgentPivotRangeForOp(op, index);
  const requestedName = cwsAgentSheetNameText(
    op?.name ??
    op?.pivotSheetName ??
    op?.targetSheetName ??
    op?.outputSheetName ??
    op?.newSheetName,
  );
  let name = "";
  if (requestedName) {
    const validation = validateSheetName(requestedName, -1);
    if (!validation.valid) {
      throw new Error(`Agent 操作 ${index + 1}: ${validation.message}`);
    }
    name = validation.name;
  }
  return {
    type: "createPivotTable",
    ...range,
    name,
    insertIndex: cwsAgentSheetInsertIndexForOp(op),
  };
}



function normalizeCwsAgentCreateForecastSheetOp(op, index) {
  const range = cwsAgentForecastRangeForOp(op, index);
  if (range.bottom <= range.top) {
    throw new Error(`Agent 操作 ${index + 1} の予測範囲には2行以上の履歴データが必要です。`);
  }
  const forecastEnd = cwsAgentForecastEndValue(op?.forecastEnd ?? op?.end ?? op?.forecastTo ?? op?.until);
  const normalized = {
    type: "createForecastSheet",
    ...range,
    periods: cwsAgentForecastPeriods(op?.periods ?? op?.forecastPeriods ?? op?.steps ?? op?.count),
    chartType: cwsAgentForecastChartType(op?.chartType ?? op?.chartKind ?? op?.kind),
    confidence: cwsAgentOptionalBoolean(op?.confidence ?? op?.showConfidence ?? op?.confidenceInterval, true),
    confidenceLevel: cwsAgentForecastConfidenceLevel(op?.confidenceLevel ?? op?.confidencePercent),
    seasonality: cwsAgentForecastSeasonality(op?.seasonality ?? op?.seasonalPeriod),
    missingData: cwsAgentForecastMissingData(op?.missingData ?? op?.missingPoints),
    duplicateAggregate: cwsAgentForecastDuplicateAggregate(op?.duplicateAggregate ?? op?.aggregateDuplicates ?? op?.aggregation),
  };
  if (forecastEnd.number !== null) normalized.forecastEnd = forecastEnd.number;
  if (forecastEnd.text) normalized.forecastEndText = forecastEnd.text;
  return normalized;
}



function normalizeCwsAgentGoalSeekOp(op, index) {
  const setCell = cwsAgentGoalSeekCellForOp(
    op,
    index,
    "数式入力セル",
    ["setCell", "formulaCell", "resultCell", "targetCell", "outputCell", "cell", "address", "ref"],
    ["setRow", "formulaRow", "resultRow", "targetRow", "row"],
    ["setCol", "setColumn", "formulaCol", "formulaColumn", "resultCol", "resultColumn", "targetCol", "targetColumn", "col", "column"],
    ["setSheetIndex", "formulaSheetIndex", "resultSheetIndex", "targetSheetIndex"],
    ["setSheetName", "formulaSheetName", "resultSheetName", "targetSheetName"],
  );
  const changingCell = cwsAgentGoalSeekCellForOp(
    op,
    index,
    "変化させるセル",
    ["changingCell", "variableCell", "byChangingCell", "inputCell", "changeCell", "adjustCell", "changing"],
    ["changingRow", "variableRow", "inputRow", "changeRow", "adjustRow"],
    ["changingCol", "changingColumn", "variableCol", "variableColumn", "inputCol", "inputColumn", "changeCol", "changeColumn", "adjustCol", "adjustColumn"],
    ["changingSheetIndex", "variableSheetIndex", "inputSheetIndex", "changeSheetIndex", "adjustSheetIndex"],
    ["changingSheetName", "variableSheetName", "inputSheetName", "changeSheetName", "adjustSheetName"],
  );
  const targetValue = cwsAgentGoalSeekTargetValue(
    cwsAgentFirstExistingValue(op, ["targetValue", "target", "goal", "goalValue", "toValue", "value"]),
  );
  if (targetValue === null) {
    throw new Error(`Agent 操作 ${index + 1} のゴール シーク目標値が正しくありません。`);
  }
  return {
    type: "goalSeek",
    setSheetIndex: setCell.sheetIndex,
    setRow: setCell.row,
    setCol: setCell.col,
    changingSheetIndex: changingCell.sheetIndex,
    changingRow: changingCell.row,
    changingCol: changingCell.col,
    targetValue,
  };
}



function normalizeCwsAgentDataTableOp(op, index) {
  const rangeValue = cwsAgentNestedValueForOp(op, ["range", "tableRange", "dataTableRange", "outputRange", "targetRange", "address", "ref"]);
  const rangeSource = rangeValue === undefined || rangeValue === null || rangeValue === "" ? op : rangeValue;
  const range = cwsAgentFlexibleRangeForOp(rangeSource, op, index, "データ テーブル範囲", CWS_AGENT_MAX_DATA_TABLE_CELLS);
  if (range.bottom <= range.top || range.right <= range.left) {
    throw new Error(`Agent 操作 ${index + 1} のデータ テーブル範囲が正しくありません。`);
  }
  const rowInput = cwsAgentDataTableInputCellForOp(
    op,
    index,
    "行の代入セル",
    ["rowInput", "rowInputCell", "inputRowCell", "rowCell", "rowVariableCell"],
    ["rowInputRow", "inputRow", "rowVariableRow"],
    ["rowInputCol", "rowInputColumn", "inputColumn", "rowVariableCol", "rowVariableColumn"],
    ["rowInputSheetIndex", "inputSheetIndex", "rowVariableSheetIndex"],
    ["rowInputSheetName", "inputSheetName", "rowVariableSheetName"],
    true,
  );
  const columnInput = cwsAgentDataTableInputCellForOp(
    op,
    index,
    "列の代入セル",
    ["columnInput", "columnInputCell", "colInput", "colInputCell", "inputColumnCell", "columnCell", "columnVariableCell"],
    ["columnInputRow", "colInputRow", "inputRow", "columnVariableRow"],
    ["columnInputCol", "columnInputColumn", "colInputCol", "colInputColumn", "columnVariableCol", "columnVariableColumn"],
    ["columnInputSheetIndex", "colInputSheetIndex", "inputSheetIndex", "columnVariableSheetIndex"],
    ["columnInputSheetName", "colInputSheetName", "inputSheetName", "columnVariableSheetName"],
    true,
  );
  if (!rowInput && !columnInput) {
    throw new Error(`Agent 操作 ${index + 1} のデータ テーブル代入セルが指定されていません。`);
  }
  return {
    type: "dataTable",
    sheetIndex: range.sheetIndex,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    rowInputSheetIndex: rowInput?.sheetIndex ?? null,
    rowInputRow: rowInput?.row ?? null,
    rowInputCol: rowInput?.col ?? null,
    columnInputSheetIndex: columnInput?.sheetIndex ?? null,
    columnInputRow: columnInput?.row ?? null,
    columnInputCol: columnInput?.col ?? null,
    overwrite: cwsAgentOptionalBoolean(op?.overwrite ?? op?.replaceExisting ?? op?.allowOverwrite, false),
  };
}



function cwsAgentDataTableInputCellForOp(op, index, label, cellKeys, rowKeys, colKeys, sheetIndexKeys, sheetNameKeys, optional = false) {
  const cellValue = cwsAgentFirstExistingValue(op, cellKeys);
  const hasCellValue = cellValue !== null && cellValue !== undefined && String(cellValue).trim() !== "";
  const parsedCell = hasCellValue ? cwsAgentParseCellAddress(String(cellValue).trim()) : null;
  const rowValue = cwsAgentFirstExistingValue(op, rowKeys);
  const colValue = cwsAgentFirstExistingValue(op, colKeys);
  const hasCoordinateValue = rowValue !== null && rowValue !== undefined && rowValue !== "" && colValue !== null && colValue !== undefined && colValue !== "";
  if (!hasCellValue && !hasCoordinateValue && optional) return null;
  const sheetIndexValue = cwsAgentFirstExistingValue(op, sheetIndexKeys);
  const sheetNameValue = cwsAgentFirstExistingValue(op, sheetNameKeys);
  const sheetIndex = cwsAgentInteger(sheetIndexValue) ?? cwsAgentSheetIndexForOp({ sheetName: sheetNameValue || op?.sheetName }, parsedCell?.sheetName);
  const row = cwsAgentPositiveInteger(rowValue) || parsedCell?.row || 0;
  const col = cwsAgentColumnValue(colValue) || parsedCell?.col || 0;
  if (!row || !col || row > EXCEL_MAX_ROWS || col > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の${label}が正しくありません。`);
  }
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} の${label}シートが見つかりません。`);
  }
  return { sheetIndex, row, col };
}



function normalizeCwsAgentExistingShapeEditOp(op) {
  const normalized = {...cwsNormalizeExistingShapeEdit(op), sheetIndex:cwsAgentSheetIndexForOp(op)};
  const sheet = state.model?.sheets?.[normalized.sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) throw new Error('変更対象のシートが見つかりません。');
  cwsExistingShapeEditTarget(sheet,normalized);
  return normalized;
}

function normalizeCwsAgentSetShapeTextOp(op, index) {
  if (typeof op.shapeId !== 'string' || !op.shapeId || typeof op.text !== 'string' || typeof op.expectedText !== 'string') {
    throw new Error(`Agent 操作 ${index + 1}: 図形 ID、変更前と変更後の文字列が必要です。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(op);
  const normalized = {type:'setShapeText',sheetIndex,shapeId:op.shapeId,text:op.text,expectedText:op.expectedText};
  cwsAgentShapeTextTarget(normalized);
  return normalized;
}



function cwsAgentShapeTextTarget(op) {
  const sheet = state.model?.sheets?.[op.sheetIndex];
  const matches = (sheet?.images || []).filter(image => image.id === op.shapeId && image.shape);
  if (!sheet || !isSheetVisible(sheet) || matches.length !== 1) throw new Error('変更対象の図形を一意に特定できません。');
  const image = matches[0];
  const before = image.shape.text || shapePlainTextFromRuns(image.shape.textRuns) || '';
  if (before !== op.expectedText) throw new Error('図形の文字が読み取り時から変更されています。再読み取りが必要です。');
  return {sheet,image};
}



function normalizeCwsAgentAddShapeOp(op, index, aliasType = "") {
  const connector = cwsAgentShapeConnectorForOp(op);
  let shapeType = cwsAgentShapeType(op?.shapeType ?? op?.shapeKind ?? op?.shape ?? op?.kind ?? op?.preset ?? op?.typeName, aliasType);
  if (connector && !isLineShapeType(shapeType)) {
    shapeType = "arrow";
  }
  const location = cwsAgentShapeLocationForOp(op, index);
  const sheet = state.model?.sheets?.[location.sheetIndex];
  const style = cwsAgentShapeStyleForOp(op, index);
  if (connector && isLineShapeType(shapeType) && !Object.prototype.hasOwnProperty.call(style, "lineEndArrow")) {
    style.lineEndArrow = !(connector.fromCell && shapeType === "line");
  }
  const box = cwsAgentShapeBoxForOp(op, index, sheet, location, shapeType, style);
  return {
    type: "addShape",
    sheetIndex: location.sheetIndex,
    shapeType,
    box,
    placement: location.parsedRange || location.parsedCell ? {
      location,
      dimensions: {
        left: op?.left ?? op?.x, top: op?.top ?? op?.y,
        anchorOffsetX: op?.anchorOffsetX, anchorOffsetY: op?.anchorOffsetY,
        width: op?.width ?? op?.w ?? op?.size?.width,
        height: op?.height ?? op?.h ?? op?.size?.height,
      },
    } : null,
    style,
    name: cwsAgentShapeObjectName(op?.objectName ?? op?.shapeName ?? op?.name),
    layoutRole: ['canvas','container'].includes(op?.layoutRole) ? op.layoutRole : null,
    componentId: typeof op?.componentId === 'string' ? op.componentId : null,
    componentLabelShape: typeof op?.componentLabelShape === 'string' ? op.componentLabelShape : null,
    parentShape: typeof op?.parentShape === 'string' ? op.parentShape : null,
    branchRelation: op?.branchRelation?.fromShape && op?.branchRelation?.toShape ? {
      fromShape: String(op.branchRelation.fromShape), toShape: String(op.branchRelation.toShape),
    } : null,
    connector,
    anchor: location.anchorLabel,
  };
}



function cwsAgentShapeConnectorForOp(op) {
  const connector = cwsAgentPlainObject(op?.connector);
  const fromShape = cwsAgentShapeReferenceText(
    cwsAgentConnectorValue([op, connector], ["fromShape", "sourceShape", "startShape", "fromShapeName", "sourceShapeName", "startShapeName"]) ||
      cwsAgentConnectorValue([connector], ["from", "source", "start"]) ||
      cwsAgentConnectorEndpointValue([op, connector], ["from", "source", "start"], ["shape", "shapeName", "objectName", "name", "id", "ref"]),
  );
  const toShape = cwsAgentShapeReferenceText(
    cwsAgentConnectorValue([op, connector], ["toShape", "targetShape", "endShape", "toShapeName", "targetShapeName", "endShapeName"]) ||
      cwsAgentConnectorValue([connector], ["to", "target", "end"]) ||
      cwsAgentConnectorEndpointValue([op, connector], ["to", "target", "end"], ["shape", "shapeName", "objectName", "name", "id", "ref"]),
  );
  if (!fromShape && !toShape && (op?.fromCell || op?.toCell)) {
    const fromCell = cwsAgentParseCellAddress(String(op.fromCell || ""));
    const toCell = cwsAgentParseCellAddress(String(op.toCell || ""));
    if (!fromCell || !toCell) throw new Error("線の両端に有効な fromCell/toCell が必要です。");
    return { fromCell, toCell, fromSite: cwsAgentLineEndpointSide(op.fromSide), toSite: cwsAgentLineEndpointSide(op.toSide) };
  }
  if (!fromShape || !toShape) return null;
  return {
    fromShape,
    toShape,
    fromSite: cwsAgentOptionalLineEndpointSide(
      cwsAgentConnectorValue([op, connector], ["fromSite", "fromSide", "sourceSite", "sourceSide", "startSite", "startSide"]) ||
        cwsAgentConnectorEndpointValue([op, connector], ["from", "source", "start"], ["site", "side", "position", "edge"]),
    ),
    toSite: cwsAgentOptionalLineEndpointSide(
      cwsAgentConnectorValue([op, connector], ["toSite", "toSide", "targetSite", "targetSide", "endSite", "endSide"]) ||
        cwsAgentConnectorEndpointValue([op, connector], ["to", "target", "end"], ["site", "side", "position", "edge"]),
    ),
  };
}



function cwsAgentPlainObject(value) {
  return value && typeof value === "object" && !Array.isArray(value) ? value : null;
}



function cwsAgentConnectorValue(sources, keys) {
  for (const source of sources || []) {
    if (!source || typeof source !== "object" || Array.isArray(source)) continue;
    for (const key of keys) {
      if (!Object.prototype.hasOwnProperty.call(source, key)) continue;
      const value = source[key];
      if (value !== null && value !== undefined && value !== "" && !Array.isArray(value) && typeof value !== "object") return value;
    }
  }
  return "";
}



function cwsAgentConnectorEndpointValue(sources, endpointKeys, valueKeys) {
  for (const source of sources || []) {
    if (!source || typeof source !== "object" || Array.isArray(source)) continue;
    for (const endpointKey of endpointKeys) {
      const endpoint = cwsAgentPlainObject(source[endpointKey]);
      const value = cwsAgentConnectorValue([endpoint], valueKeys);
      if (value !== "") return value;
    }
  }
  return "";
}



function cwsAgentShapeReferenceText(value) {
  if (value === null || value === undefined || Array.isArray(value) || typeof value === "object") return "";
  return String(value).trim().slice(0, 120);
}



function cwsAgentOptionalLineEndpointSide(value) {
  if (value === null || value === undefined || value === "") return "";
  const normalized = cwsAgentLineEndpointSide(value);
  return normalized === "center" ? "" : normalized;
}



function normalizeCwsAgentSetDataValidationOp(op, index) {
  const range = cwsAgentValidationRangeForOp(op, index);
  const sheet = state.model?.sheets?.[range.sheetIndex];
  const rule = cwsAgentDataValidationRule(op, index, {
    anchor: { row: range.top, col: range.left },
    sheet,
  });
  return {
    type: "setDataValidation",
    ...range,
    rule,
  };
}



function normalizeCwsAgentClearDataValidationOp(op, index) {
  return {
    type: "clearDataValidation",
    ...cwsAgentValidationRangeForOp(op, index),
  };
}



function normalizeCwsAgentSetHyperlinkOp(op, index) {
  const range = cwsAgentRangeForOp(op, index);
  return {
    type: "setHyperlink",
    ...range,
    hyperlink: cwsAgentHyperlinkTarget(op, index),
    text: cwsAgentOptionalHyperlinkText(op?.text ?? op?.displayText ?? op?.label),
  };
}



function normalizeCwsAgentClearHyperlinkOp(op, index) {
  return {
    type: "clearHyperlink",
    ...cwsAgentRangeForOp(op, index),
  };
}



function normalizeCwsAgentSetNoteOp(op, index) {
  const range = cwsAgentRangeForOp(op, index);
  return {
    type: "setNote",
    ...range,
    note: cwsAgentNoteText(op?.note ?? op?.memo ?? op?.text ?? op?.value ?? op?.comment, index),
    noteVisible: cwsAgentOptionalBoolSetting(op?.noteVisible ?? op?.visible ?? op?.show ?? op?.display),
  };
}



function normalizeCwsAgentClearNoteOp(op, index) {
  return {
    type: "clearNote",
    ...cwsAgentRangeForOp(op, index),
  };
}



function normalizeCwsAgentSetCommentOp(op, index) {
  const range = cwsAgentRangeForOp(op, index);
  return {
    type: "setComment",
    ...range,
    comment: cwsAgentCommentText(op?.comment ?? op?.text ?? op?.value ?? op?.note, index),
  };
}



function normalizeCwsAgentAppendCommentOp(op, index) {
  const range = cwsAgentRangeForOp(op, index);
  return {
    type: "appendComment",
    ...range,
    comment: cwsAgentCommentText(op?.comment ?? op?.text ?? op?.reply ?? op?.value, index),
  };
}



function normalizeCwsAgentClearCommentOp(op, index) {
  return {
    type: "clearComment",
    ...cwsAgentRangeForOp(op, index),
  };
}



function normalizeCwsAgentSetNamedRangeOp(op, index) {
  const range = cwsAgentNamedRangeForOp(op, index);
  const sheet = state.model?.sheets?.[range.sheetIndex];
  const scope = cwsAgentDefinedNameScope(op, false);
  const scopeSheetIndex = scope === "sheet" ? cwsAgentDefinedNameScopeSheetIndex(op, range.sheetIndex, index) : null;
  const rawRefersTo = op?.refersTo ?? op?.formula ?? op?.reference;
  return {
    type: "setNamedRange",
    name: cwsAgentDefinedNameText(op?.name ?? op?.namedRange ?? op?.definedName ?? op?.label, index),
    scope,
    scopeSheetIndex,
    sheetIndex: range.sheetIndex,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    refersTo: rawRefersTo === null || rawRefersTo === undefined || rawRefersTo === ""
      ? `=${absoluteRangeReference(sheet, range)}`
      : normalizeDefinedNameRefersTo(cwsAgentCellValue(rawRefersTo)),
    comment: cwsAgentSmallText(op?.comment ?? op?.description ?? ""),
  };
}



function normalizeCwsAgentClearNamedRangeOp(op, index) {
  const scope = cwsAgentDefinedNameScope(op, true);
  return {
    type: "clearNamedRange",
    name: cwsAgentDefinedNameText(op?.name ?? op?.namedRange ?? op?.definedName ?? op?.label, index),
    scope,
    scopeSheetIndex: scope === "sheet" ? cwsAgentDefinedNameScopeSheetIndex(op, state.activeSheetIndex, index) : null,
  };
}



function normalizeCwsAgentProtectSheetOp(op, index) {
  const sheetIndex = cwsAgentTargetSheetIndexForOp(op, {
    fallbackName: op?.sheetName ?? op?.targetSheetName ?? op?.target ?? "",
  });
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: "protectSheet",
    sheetIndex,
    options: cwsAgentSheetProtectionOptions(op),
  };
}



function normalizeCwsAgentUnprotectSheetOp(op, index) {
  const sheetIndex = cwsAgentTargetSheetIndexForOp(op, {
    fallbackName: op?.sheetName ?? op?.targetSheetName ?? op?.target ?? "",
  });
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: "unprotectSheet",
    sheetIndex,
    password: cwsAgentOptionalText(op?.password ?? op?.passcode ?? ""),
  };
}



function normalizeCwsAgentActivateSheetOp(op, index) {
  const sheetIndex = cwsAgentTargetSheetIndexForOp(op, {
    fallbackName: op?.sheetName ?? op?.targetSheetName ?? op?.targetName ?? op?.target ?? op?.name ?? op?.title ?? "",
  });
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: "activateSheet",
    sheetIndex,
  };
}



function normalizeCwsAgentSelectRangeOp(op, index) {
  const text = op?.range || op?.address || op?.ref || op?.cell || "";
  const parsedRange = cwsAgentParseRangeAddress(text);
  const parsedColumnRange = cwsAgentParseColumnRangeAddress(text);
  const parsedRowRange = cwsAgentParseRowRangeAddress(text);
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName || parsedColumnRange?.sheetName || parsedRowRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }

  let fullColumn = false;
  let fullRow = false;
  let top = cwsAgentPositiveInteger(op?.top || op?.row) || parsedRange?.top || 0;
  let left = cwsAgentColumnValue(op?.left || op?.col || op?.column) || parsedRange?.left || 0;
  let bottom = cwsAgentPositiveInteger(op?.bottom || op?.endRow || op?.toRow) || parsedRange?.bottom || top;
  let right = cwsAgentColumnValue(op?.right || op?.endCol || op?.toCol) || parsedRange?.right || left;

  if (!parsedRange && (parsedColumnRange || op?.fullColumn === true)) {
    fullColumn = true;
    top = 1;
    bottom = sheet.rowCount || EXCEL_MAX_ROWS;
    left = parsedColumnRange?.left || left;
    right = parsedColumnRange?.right || right;
  } else if (!parsedRange && (parsedRowRange || op?.fullRow === true)) {
    fullRow = true;
    top = parsedRowRange?.top || top;
    bottom = parsedRowRange?.bottom || bottom;
    left = 1;
    right = sheet.colCount || EXCEL_MAX_COLS;
  }

  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の選択範囲が正しくありません。`);
  }
  return {
    type: "selectRange",
    sheetIndex,
    top,
    left,
    bottom,
    right,
    fullRow,
    fullColumn,
  };
}



function normalizeCwsAgentCopyMoveRangeOp(op, index, type) {
  const sourceText = op?.sourceRange || op?.fromRange || op?.srcRange || op?.source || op?.from || op?.range || op?.address || op?.ref || "";
  const targetText = op?.targetRange || op?.toRange || op?.destinationRange || op?.destRange || op?.target || op?.to || op?.destination || op?.dest || "";
  const parsedSource = cwsAgentParseRangeAddress(sourceText);
  const parsedTarget = cwsAgentParseRangeAddress(targetText);
  const sourceTop = cwsAgentPositiveInteger(op?.sourceTop || op?.fromTop || op?.top || op?.row) || parsedSource?.top || 0;
  const sourceLeft = cwsAgentColumnValue(op?.sourceLeft || op?.fromLeft || op?.left || op?.col || op?.column) || parsedSource?.left || 0;
  const sourceBottom = cwsAgentPositiveInteger(op?.sourceBottom || op?.fromBottom || op?.bottom) || parsedSource?.bottom || sourceTop;
  const sourceRight = cwsAgentColumnValue(op?.sourceRight || op?.fromRight || op?.right) || parsedSource?.right || sourceLeft;
  if (!sourceTop || !sourceLeft || sourceBottom < sourceTop || sourceRight < sourceLeft || sourceBottom > EXCEL_MAX_ROWS || sourceRight > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のコピー元範囲が正しくありません。`);
  }
  const sourceCellCount = (sourceBottom - sourceTop + 1) * (sourceRight - sourceLeft + 1);
  if (sourceCellCount > CWS_AGENT_MAX_COPY_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} のコピー元範囲が大きすぎます。`);
  }

  const targetTop = cwsAgentPositiveInteger(op?.targetTop || op?.toTop || op?.destinationTop || op?.destTop || op?.targetRow || op?.toRow) || parsedTarget?.top || 0;
  const targetLeft = cwsAgentColumnValue(op?.targetLeft || op?.toLeft || op?.destinationLeft || op?.destLeft || op?.targetCol || op?.toCol) || parsedTarget?.left || 0;
  if (!targetTop || !targetLeft || targetTop > EXCEL_MAX_ROWS || targetLeft > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のコピー先範囲が正しくありません。`);
  }
  const targetBottom = targetTop + (sourceBottom - sourceTop);
  const targetRight = targetLeft + (sourceRight - sourceLeft);
  if (targetBottom > EXCEL_MAX_ROWS || targetRight > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のコピー先範囲が大きすぎます。`);
  }
  if (parsedTarget && (parsedTarget.bottom !== parsedTarget.top || parsedTarget.right !== parsedTarget.left)) {
    const targetHeight = parsedTarget.bottom - parsedTarget.top + 1;
    const targetWidth = parsedTarget.right - parsedTarget.left + 1;
    const sourceHeight = sourceBottom - sourceTop + 1;
    const sourceWidth = sourceRight - sourceLeft + 1;
    if (targetHeight !== sourceHeight || targetWidth !== sourceWidth) {
      throw new Error(`Agent 操作 ${index + 1} のコピー先範囲サイズがコピー元と一致しません。`);
    }
  }

  const sourceSheetName = parsedSource?.sheetName || cwsAgentSheetNameText(op?.sourceSheetName ?? op?.fromSheetName ?? op?.sourceSheet ?? op?.fromSheet ?? op?.sheetName ?? "");
  const targetSheetName = parsedTarget?.sheetName || cwsAgentSheetNameText(op?.targetSheetName ?? op?.toSheetName ?? op?.destinationSheetName ?? op?.destSheetName ?? "");
  const sourceSheetIndex = cwsAgentSheetIndexForOp({ sheetIndex: op?.sourceSheetIndex ?? op?.fromSheetIndex, sheetName: sourceSheetName }, parsedSource?.sheetName);
  const targetSheetIndex = cwsAgentSheetIndexForOp({ sheetIndex: op?.targetSheetIndex ?? op?.toSheetIndex ?? op?.destinationSheetIndex ?? op?.destSheetIndex, sheetName: targetSheetName }, parsedTarget?.sheetName || targetSheetName || sourceSheetName);
  const sourceSheet = state.model?.sheets?.[sourceSheetIndex];
  const targetSheet = state.model?.sheets?.[targetSheetIndex];
  if (!sourceSheet || !isSheetVisible(sourceSheet)) {
    throw new Error(`Agent 操作 ${index + 1} のコピー元シートが見つかりません。`);
  }
  if (!targetSheet || !isSheetVisible(targetSheet)) {
    throw new Error(`Agent 操作 ${index + 1} のコピー先シートが見つかりません。`);
  }

  const sourceRange = { sheetIndex: sourceSheetIndex, top: sourceTop, left: sourceLeft, bottom: sourceBottom, right: sourceRight };
  const targetRange = { sheetIndex: targetSheetIndex, top: targetTop, left: targetLeft, bottom: targetBottom, right: targetRight };
  if (type === "moveRange" && sourceSheetIndex === targetSheetIndex && rangesIntersect(sourceRange, targetRange)) {
    throw new Error("移動元と移動先が重なっているため、Agent 操作を適用しません。");
  }
  return {
    type,
    sourceSheetIndex,
    targetSheetIndex,
    sourceTop,
    sourceLeft,
    sourceBottom,
    sourceRight,
    targetTop,
    targetLeft,
    targetBottom,
    targetRight,
    keepSourceFormatting: cwsAgentOptionalBoolean(op?.keepSourceFormatting ?? op?.formatting ?? op?.formats, true),
  };
}



function normalizeCwsAgentSetColumnWidthOp(op, index) {
  const parsedCellRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  const parsedColumnRange = cwsAgentParseColumnRangeAddress(op?.range || op?.address || op?.ref || "");
  const startCol = cwsAgentColumnValue(op?.startCol || op?.col || op?.column || op?.left) || parsedColumnRange?.left || parsedCellRange?.left || 0;
  const endCol = cwsAgentColumnValue(op?.endCol || op?.toCol || op?.right) || parsedColumnRange?.right || parsedCellRange?.right || startCol;
  const width = cwsAgentColumnWidthPixels(op);
  if (!startCol || !endCol || startCol > EXCEL_MAX_COLS || endCol > EXCEL_MAX_COLS || width === null) {
    throw new Error(`Agent 操作 ${index + 1} の列幅指定が正しくありません。`);
  }
  const left = Math.min(startCol, endCol);
  const right = Math.max(startCol, endCol);
  if (right - left + 1 > CWS_AGENT_MAX_STRUCTURE_COLUMNS) {
    throw new Error(`Agent 操作 ${index + 1} の列指定が大きすぎます。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedColumnRange?.sheetName || parsedCellRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: "setColumnWidth",
    sheetIndex,
    left,
    right,
    width,
  };
}



function normalizeCwsAgentSetRowHeightOp(op, index) {
  const parsedCellRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  const parsedRowRange = cwsAgentParseRowRangeAddress(op?.range || op?.address || op?.ref || "");
  const startRow = cwsAgentPositiveInteger(op?.startRow || op?.row || op?.index || op?.top) || parsedRowRange?.top || parsedCellRange?.top || 0;
  const endRow = cwsAgentPositiveInteger(op?.endRow || op?.toRow || op?.bottom) || parsedRowRange?.bottom || parsedCellRange?.bottom || startRow;
  const height = cwsAgentRowHeightPixels(op);
  if (!startRow || !endRow || startRow > EXCEL_MAX_ROWS || endRow > EXCEL_MAX_ROWS || height === null) {
    throw new Error(`Agent 操作 ${index + 1} の行高指定が正しくありません。`);
  }
  const top = Math.min(startRow, endRow);
  const bottom = Math.max(startRow, endRow);
  if (bottom - top + 1 > CWS_AGENT_MAX_STRUCTURE_ROWS) {
    throw new Error(`Agent 操作 ${index + 1} の行指定が大きすぎます。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRowRange?.sheetName || parsedCellRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: "setRowHeight",
    sheetIndex,
    top,
    bottom,
    height,
  };
}



function normalizeCwsAgentAutoFitColumnsOp(op, index) {
  const parsedCellRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  const parsedColumnRange = cwsAgentParseColumnRangeAddress(op?.range || op?.address || op?.ref || "");
  const startCol = cwsAgentColumnValue(op?.startCol || op?.col || op?.column || op?.left) || parsedColumnRange?.left || parsedCellRange?.left || 0;
  const endCol = cwsAgentColumnValue(op?.endCol || op?.toCol || op?.right) || parsedColumnRange?.right || parsedCellRange?.right || startCol;
  if (!startCol || !endCol || startCol > EXCEL_MAX_COLS || endCol > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の自動列幅指定が正しくありません。`);
  }
  const left = Math.min(startCol, endCol);
  const right = Math.max(startCol, endCol);
  if (right - left + 1 > CWS_AGENT_MAX_STRUCTURE_COLUMNS) {
    throw new Error(`Agent 操作 ${index + 1} の列指定が大きすぎます。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedColumnRange?.sheetName || parsedCellRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: "autoFitColumns",
    sheetIndex,
    left,
    right,
  };
}



function normalizeCwsAgentAutoFitRowsOp(op, index) {
  if (Object.hasOwn(op, "growOnly") && typeof op.growOnly !== "boolean") throw new Error("growOnly must be boolean.");
  const parsedCellRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  const parsedRowRange = cwsAgentParseRowRangeAddress(op?.range || op?.address || op?.ref || "");
  const startRow = cwsAgentPositiveInteger(op?.startRow || op?.row || op?.index || op?.top) || parsedRowRange?.top || parsedCellRange?.top || 0;
  const endRow = cwsAgentPositiveInteger(op?.endRow || op?.toRow || op?.bottom) || parsedRowRange?.bottom || parsedCellRange?.bottom || startRow;
  if (!startRow || !endRow || startRow > EXCEL_MAX_ROWS || endRow > EXCEL_MAX_ROWS) {
    throw new Error(`Agent 操作 ${index + 1} の自動行高指定が正しくありません。`);
  }
  const top = Math.min(startRow, endRow);
  const bottom = Math.max(startRow, endRow);
  if (bottom - top + 1 > CWS_AGENT_MAX_STRUCTURE_ROWS) {
    throw new Error(`Agent 操作 ${index + 1} の行指定が大きすぎます。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRowRange?.sheetName || parsedCellRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: "autoFitRows",
    sheetIndex,
    top,
    bottom,
    ...(op.growOnly === true ? { growOnly: true } : {}),
  };
}



function normalizeCwsAgentHeaderHiddenOp(op, index, kind, hidden) {
  const parsedCellRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  if (kind === "column") {
    const parsedColumnRange = cwsAgentParseColumnRangeAddress(op?.range || op?.address || op?.ref || "");
    const startCol = cwsAgentColumnValue(op?.startCol || op?.col || op?.column || op?.index || op?.left) || parsedColumnRange?.left || parsedCellRange?.left || 0;
    const endCol = cwsAgentColumnValue(op?.endCol || op?.toCol || op?.right) || parsedColumnRange?.right || parsedCellRange?.right || startCol;
    if (!startCol || !endCol || startCol > EXCEL_MAX_COLS || endCol > EXCEL_MAX_COLS) {
      throw new Error(`Agent 操作 ${index + 1} の列指定が正しくありません。`);
    }
    const left = Math.min(startCol, endCol);
    const right = Math.max(startCol, endCol);
    if (right - left + 1 > CWS_AGENT_MAX_STRUCTURE_COLUMNS) {
      throw new Error(`Agent 操作 ${index + 1} の列指定が大きすぎます。`);
    }
    const sheetIndex = cwsAgentSheetIndexForOp(op, parsedColumnRange?.sheetName || parsedCellRange?.sheetName);
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet || !isSheetVisible(sheet)) {
      throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
    }
    return {
      type: hidden ? "hideColumns" : "unhideColumns",
      sheetIndex,
      left,
      right,
    };
  }

  const parsedRowRange = cwsAgentParseRowRangeAddress(op?.range || op?.address || op?.ref || "");
  const startRow = cwsAgentPositiveInteger(op?.startRow || op?.row || op?.index || op?.top) || parsedRowRange?.top || parsedCellRange?.top || 0;
  const endRow = cwsAgentPositiveInteger(op?.endRow || op?.toRow || op?.bottom) || parsedRowRange?.bottom || parsedCellRange?.bottom || startRow;
  if (!startRow || !endRow || startRow > EXCEL_MAX_ROWS || endRow > EXCEL_MAX_ROWS) {
    throw new Error(`Agent 操作 ${index + 1} の行指定が正しくありません。`);
  }
  const top = Math.min(startRow, endRow);
  const bottom = Math.max(startRow, endRow);
  if (bottom - top + 1 > CWS_AGENT_MAX_STRUCTURE_ROWS) {
    throw new Error(`Agent 操作 ${index + 1} の行指定が大きすぎます。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRowRange?.sheetName || parsedCellRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: hidden ? "hideRows" : "unhideRows",
    sheetIndex,
    top,
    bottom,
  };
}



function normalizeCwsAgentOutlineOp(op, index, kind, delta) {
  const parsedCellRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  if (kind === "column") {
    const parsedColumnRange = cwsAgentParseColumnRangeAddress(op?.range || op?.address || op?.ref || "");
    const startCol = cwsAgentColumnValue(op?.startCol || op?.col || op?.column || op?.index || op?.left) || parsedColumnRange?.left || parsedCellRange?.left || 0;
    const endCol = cwsAgentColumnValue(op?.endCol || op?.toCol || op?.right) || parsedColumnRange?.right || parsedCellRange?.right || startCol;
    if (!startCol || !endCol || startCol > EXCEL_MAX_COLS || endCol > EXCEL_MAX_COLS) {
      throw new Error(`Agent 操作 ${index + 1} の列指定が正しくありません。`);
    }
    const left = Math.min(startCol, endCol);
    const right = Math.max(startCol, endCol);
    if (right - left + 1 > CWS_AGENT_MAX_STRUCTURE_COLUMNS) {
      throw new Error(`Agent 操作 ${index + 1} の列指定が大きすぎます。`);
    }
    const sheetIndex = cwsAgentSheetIndexForOp(op, parsedColumnRange?.sheetName || parsedCellRange?.sheetName);
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet || !isSheetVisible(sheet)) {
      throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
    }
    return {
      type: delta > 0 ? "groupColumns" : "ungroupColumns",
      axis: "columns",
      delta,
      sheetIndex,
      left,
      right,
    };
  }

  const parsedRowRange = cwsAgentParseRowRangeAddress(op?.range || op?.address || op?.ref || "");
  const startRow = cwsAgentPositiveInteger(op?.startRow || op?.row || op?.index || op?.top) || parsedRowRange?.top || parsedCellRange?.top || 0;
  const endRow = cwsAgentPositiveInteger(op?.endRow || op?.toRow || op?.bottom) || parsedRowRange?.bottom || parsedCellRange?.bottom || startRow;
  if (!startRow || !endRow || startRow > EXCEL_MAX_ROWS || endRow > EXCEL_MAX_ROWS) {
    throw new Error(`Agent 操作 ${index + 1} の行指定が正しくありません。`);
  }
  const top = Math.min(startRow, endRow);
  const bottom = Math.max(startRow, endRow);
  if (bottom - top + 1 > CWS_AGENT_MAX_STRUCTURE_ROWS) {
    throw new Error(`Agent 操作 ${index + 1} の行指定が大きすぎます。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRowRange?.sheetName || parsedCellRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: delta > 0 ? "groupRows" : "ungroupRows",
    axis: "rows",
    delta,
    sheetIndex,
    top,
    bottom,
  };
}



function normalizeCwsAgentRowsOp(op, index, type) {
  const parsedRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  const row = cwsAgentPositiveInteger(op?.row || op?.startRow || op?.index || op?.at || op?.beforeRow) || parsedRange?.top || 0;
  const count = cwsAgentPositiveInteger(op?.count || op?.amount || op?.rows) || (parsedRange ? parsedRange.bottom - parsedRange.top + 1 : 1);
  if (!row || row > EXCEL_MAX_ROWS || count < 1 || count > CWS_AGENT_MAX_STRUCTURE_ROWS) {
    throw new Error(`Agent 操作 ${index + 1} の行指定が正しくありません。`);
  }
  if (row + count - 1 > EXCEL_MAX_ROWS) {
    throw new Error(`Agent 操作 ${index + 1} の行指定が大きすぎます。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  if (type === "insertRows" && row > sheet.rowCount + 1) {
    throw new Error(`Agent 操作 ${index + 1} の挿入位置が正しくありません。`);
  }
  if (type === "deleteRows" && row > sheet.rowCount) {
    throw new Error(`Agent 操作 ${index + 1} の削除位置が正しくありません。`);
  }

  return {
    type,
    sheetIndex,
    row,
    count,
  };
}



function normalizeCwsAgentColumnsOp(op, index, type) {
  const parsedRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  const col = cwsAgentColumnValue(op?.col || op?.startCol || op?.column || op?.index || op?.at || op?.beforeCol) || parsedRange?.left || 0;
  const count = cwsAgentPositiveInteger(op?.count || op?.amount || op?.cols || op?.columns) || (parsedRange ? parsedRange.right - parsedRange.left + 1 : 1);
  if (!col || col > EXCEL_MAX_COLS || count < 1 || count > CWS_AGENT_MAX_STRUCTURE_COLUMNS) {
    throw new Error(`Agent 操作 ${index + 1} の列指定が正しくありません。`);
  }
  if (col + count - 1 > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の列指定が大きすぎます。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  if (type === "insertColumns" && col > sheet.colCount + 1) {
    throw new Error(`Agent 操作 ${index + 1} の挿入位置が正しくありません。`);
  }
  if (type === "deleteColumns" && col > sheet.colCount) {
    throw new Error(`Agent 操作 ${index + 1} の削除位置が正しくありません。`);
  }

  return {
    type,
    sheetIndex,
    col,
    count,
  };
}



function normalizeCwsAgentClearRangeOp(op, index) {
  const range = cwsAgentRangeForOp(op, index);
  const mode = cwsAgentClearMode(op?.mode || op?.clear || op?.target);
  return {
    type: "clearRange",
    ...range,
    mode,
  };
}



function normalizeCwsAgentSetQuickAnalysisOp(op, index, aliasType = "") {
  const mode = cwsAgentQuickAnalysisMode(
    op?.mode ?? op?.quickAnalysisType ?? op?.analysisType ?? op?.kind ?? op?.style ?? op?.format,
    aliasType,
  );
  if (!mode) {
    throw new Error(`Agent 操作 ${index + 1} のクイック分析タイプが正しくありません。`);
  }
  return {
    type: "setQuickAnalysis",
    ...cwsAgentQuickAnalysisRangeForOp(op, index),
    mode,
  };
}



function normalizeCwsAgentClearQuickAnalysisOp(op, index) {
  return {
    type: "clearQuickAnalysis",
    ...cwsAgentQuickAnalysisRangeForOp(op, index),
  };
}



function normalizeCwsAgentSetQuickAnalysisTotalOp(op, index, aliasType = "") {
  const method = cwsAgentQuickAnalysisTotalMethod(
    op?.method ?? op?.totalType ?? op?.summaryType ?? op?.aggregate ?? op?.aggregation ?? op?.kind ?? op?.mode,
    aliasType,
  );
  if (!method) {
    throw new Error(`Agent 操作 ${index + 1} のクイック分析集計タイプが正しくありません。`);
  }
  const range = cwsAgentQuickAnalysisRangeForOp(op, index);
  if (method === "running" && range.right >= EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の累計列を追加する余白がありません。`);
  }
  if (method !== "running" && range.bottom >= EXCEL_MAX_ROWS) {
    throw new Error(`Agent 操作 ${index + 1} の集計行を追加する余白がありません。`);
  }
  return {
    type: "setQuickAnalysisTotal",
    ...range,
    method,
  };
}



function normalizeCwsAgentSetSubtotalOp(op, index) {
  const range = cwsAgentRangeForOp(op, index);
  if (range.bottom <= range.top || range.right <= range.left) {
    throw new Error(`Agent 操作 ${index + 1} の小計範囲には見出しとデータ行が必要です。`);
  }
  const sheet = state.model?.sheets?.[range.sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} の小計対象シートが見つかりません。`);
  }
  const groupCol = cwsAgentSubtotalColumnForOp(op, sheet, range, index);
  const method = cwsAgentSubtotalMethod(
    op?.method ?? op?.function ?? op?.aggregate ?? op?.aggregation ?? op?.summaryType ?? op?.kind ?? op?.mode,
  );
  if (!method) {
    throw new Error(`Agent 操作 ${index + 1} の小計集計方法が正しくありません。`);
  }
  const targetCols = cwsAgentSubtotalTargetColumns(op, sheet, range, groupCol, index);
  if (!targetCols.length) {
    throw new Error(`Agent 操作 ${index + 1} の小計対象列が指定されていません。`);
  }
  return {
    type: "setSubtotal",
    ...range,
    groupCol,
    targetCols,
    method,
    replace: cwsAgentOptionalBoolean(op?.replace ?? op?.replaceExisting ?? op?.overwrite, true),
    pageBreak: cwsAgentOptionalBoolean(op?.pageBreak ?? op?.pageBreaks ?? op?.breaks, false),
    summaryBelow: cwsAgentOptionalBoolean(op?.summaryBelow ?? op?.below ?? op?.summaryRowsBelow, true),
  };
}



function normalizeCwsAgentClearSubtotalsOp(op, index) {
  return {
    type: "clearSubtotals",
    ...cwsAgentRangeForOp(op, index),
  };
}



function normalizeCwsAgentConsolidateOp(op, index) {
  const destination = cwsAgentConsolidateDestinationForOp(op, index);
  const sheet = state.model?.sheets?.[destination.sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} の統合出力先シートが見つかりません。`);
  }
  const references = cwsAgentConsolidateReferencesForOp(op, index);
  if (!references.length) {
    throw new Error(`Agent 操作 ${index + 1} の統合参照範囲が指定されていません。`);
  }
  const method = cwsAgentSubtotalMethod(
    op?.method ?? op?.function ?? op?.aggregate ?? op?.aggregation ?? op?.summaryType ?? op?.kind ?? op?.mode,
  );
  if (!method) {
    throw new Error(`Agent 操作 ${index + 1} の統合集計方法が正しくありません。`);
  }
  return {
    type: "consolidateRanges",
    ...destination,
    references,
    method,
    topRow: cwsAgentOptionalBoolean(op?.topRow ?? op?.useTopRow ?? op?.hasHeaders ?? op?.headers, true),
    leftColumn: cwsAgentOptionalBoolean(op?.leftColumn ?? op?.useLeftColumn ?? op?.rowLabels ?? op?.groupByLeftColumn, false),
    createLinks: cwsAgentOptionalBoolean(op?.createLinks ?? op?.links ?? op?.linkToSource, false),
    overwrite: cwsAgentOptionalBoolean(op?.overwrite ?? op?.replaceExisting ?? op?.allowOverwrite, false),
  };
}



function cwsAgentConsolidateDestinationForOp(op, index) {
  const value = cwsAgentNestedValueForOp(op, ["destination", "destinationRange", "targetRange", "outputRange", "target", "to", "at"]);
  const source = value === undefined || value === null || value === "" ? op : value;
  return cwsAgentFlexibleRangeForOp(source, op, index, "統合出力先", CWS_AGENT_MAX_RANGE_CELLS);
}



function cwsAgentConsolidateReferencesForOp(op, index) {
  const source = cwsAgentNestedValueForOp(op, ["references", "sourceRanges", "sources", "ranges", "refs"]);
  const fallbackSource = source === undefined || source === null || source === ""
    ? cwsAgentNestedValueForOp(op, ["sourceRange", "source", "dataRange", "from"])
    : source;
  const items = fallbackSource && typeof fallbackSource === "object" && !Array.isArray(fallbackSource)
    ? [fallbackSource]
    : cwsAgentListItems(fallbackSource);
  const references = [];
  let totalCells = 0;
  for (const item of items) {
    if (references.length >= 20) {
      throw new Error(`Agent 操作 ${index + 1} の統合参照範囲が多すぎます。`);
    }
    const range = cwsAgentFlexibleRangeForOp(item, op, index, "統合参照範囲", CWS_AGENT_MAX_CONSOLIDATE_CELLS);
    const cellCount = (range.bottom - range.top + 1) * (range.right - range.left + 1);
    totalCells += cellCount;
    if (totalCells > CWS_AGENT_MAX_CONSOLIDATE_CELLS) {
      throw new Error(`Agent 操作 ${index + 1} の統合参照範囲が大きすぎます。`);
    }
    references.push(range);
  }
  return references;
}



function cwsAgentFlexibleRangeForOp(value, fallbackOp, index, label, maxCells = CWS_AGENT_MAX_RANGE_CELLS) {
  const objectValue = value && typeof value === "object" && !Array.isArray(value) ? value : null;
  const coordinateSource = objectValue || (value === fallbackOp ? fallbackOp : {});
  const text = objectValue
    ? String(cwsAgentFirstExistingValue(objectValue, ["range", "address", "ref", "cell", "sourceRange", "targetRange", "destination", "destinationRange"]) ?? "").trim()
    : String(value ?? "").trim();
  const opForSheet = objectValue ? { ...fallbackOp, ...objectValue } : fallbackOp;
  const preferLiteralRange = Boolean(text && !text.includes("!") && (opForSheet?.sheetName || opForSheet?.sheetIndex !== undefined));
  const parsedRange = text ? cwsAgentDecodeAgentRange(text, preferLiteralRange) : null;
  const sheetIndex = parsedRange?.sheetIndex !== undefined && parsedRange?.sheetIndex !== null
    ? parsedRange.sheetIndex
    : cwsAgentSheetIndexForOp(opForSheet, parsedRange?.sheetName);
  const top = cwsAgentPositiveInteger(coordinateSource?.top ?? coordinateSource?.row) || parsedRange?.top || parsedRange?.row || 0;
  const left = cwsAgentColumnValue(coordinateSource?.left ?? coordinateSource?.col ?? coordinateSource?.column) || parsedRange?.left || parsedRange?.col || 0;
  const bottom = cwsAgentPositiveInteger(coordinateSource?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(coordinateSource?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の${label}が正しくありません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount < 1 || cellCount > maxCells) {
    throw new Error(`Agent 操作 ${index + 1} の${label}が大きすぎます。`);
  }
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} の${label}シートが見つかりません。`);
  }
  return { sheetIndex, sheetName: sheet.name, top, left, bottom, right };
}



function cwsAgentDecodeAgentRange(value, preferLiteralRange = false) {
  const text = String(value || "").trim();
  if (!text) return null;
  if (preferLiteralRange) {
    const literal = cwsAgentParseRangeAddress(text);
    if (literal) return literal;
    const literalCell = cwsAgentParseCellAddress(text);
    if (literalCell) {
      return { sheetName: literalCell.sheetName, top: literalCell.row, left: literalCell.col, bottom: literalCell.row, right: literalCell.col };
    }
  }
  const decoded = decodeWorkbookRangeOrName(text);
  if (decoded) return decoded;
  const parsed = cwsAgentParseRangeAddress(text);
  if (parsed) return parsed;
  const cell = cwsAgentParseCellAddress(text);
  return cell
    ? { sheetName: cell.sheetName, top: cell.row, left: cell.col, bottom: cell.row, right: cell.col }
    : null;
}



function cwsAgentSubtotalColumnForOp(op, sheet, range, index) {
  const value = cwsAgentNestedValueForOp(op, ["groupBy", "groupCol", "groupColumn", "by", "keyCol", "keyColumn", "headerName"]);
  if (value === undefined || value === null || value === "") return range.left;
  const col = cwsAgentColumnInRange(value, sheet, range);
  if (!col) {
    throw new Error(`Agent 操作 ${index + 1} の小計グループ列が正しくありません。`);
  }
  return col;
}



function cwsAgentSubtotalTargetColumns(op, sheet, range, groupCol, index) {
  const source = cwsAgentNestedValueForOp(op, ["targetCols", "targetColumns", "summaryCols", "summaryColumns", "subtotalCols", "subtotalColumns", "fields", "columns"]);
  const columns = [];
  const addColumn = (col) => {
    if (!col || col < range.left || col > range.right) {
      throw new Error(`Agent 操作 ${index + 1} の小計対象列が範囲外です。`);
    }
    if (col === groupCol) return;
    if (!columns.includes(col)) columns.push(col);
  };
  if (source === undefined || source === null || source === "") {
    for (let col = range.left; col <= range.right; col += 1) addColumn(col);
    return columns;
  }
  cwsAgentListItems(source).forEach((item) => {
    const text = String(item ?? "").trim();
    if (!text) return;
    const parsedColumnRange = cwsAgentParseColumnRangeAddress(text);
    const parsedRange = parsedColumnRange || cwsAgentParseRangeAddress(text);
    if (parsedRange?.left && parsedRange?.right) {
      for (let col = parsedRange.left; col <= parsedRange.right; col += 1) addColumn(col);
      return;
    }
    addColumn(cwsAgentColumnInRange(text, sheet, range));
  });
  return columns;
}



function cwsAgentColumnInRange(value, sheet, range) {
  const direct = cwsAgentColumnValue(value);
  if (direct && direct >= range.left && direct <= range.right) return direct;
  const text = String(value ?? "").trim();
  if (!text) return 0;
  const byHeader = cwsAgentColumnByHeaderText(sheet, range, text);
  return byHeader && byHeader >= range.left && byHeader <= range.right ? byHeader : 0;
}



function cwsAgentSubtotalMethod(value) {
  const text = String(value ?? "sum").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    "": "sum",
    sum: "sum",
    total: "sum",
    subtotal: "sum",
    合計: "sum",
    合计: "sum",
    小計: "sum",
    小计: "sum",
    count: "count",
    counta: "count",
    counttext: "count",
    データの個数: "count",
    件数: "count",
    average: "average",
    avg: "average",
    mean: "average",
    平均: "average",
    max: "max",
    maximum: "max",
    最大: "max",
    最大値: "max",
    min: "min",
    minimum: "min",
    最小: "min",
    最小値: "min",
    product: "product",
    積: "product",
    countnums: "countNums",
    countnumbers: "countNums",
    countnum: "countNums",
    数値の個数: "countNums",
    stddev: "stdDev",
    stdev: "stdDev",
    標準偏差: "stdDev",
    stddevp: "stdDevP",
    stdevp: "stdDevP",
    標準偏差p: "stdDevP",
    var: "var",
    variance: "var",
    分散: "var",
    varp: "varP",
    variancep: "varP",
    分散p: "varP",
  };
  const method = aliases[text] || "";
  return DATA_SUMMARY_FUNCTIONS.some((item) => item.value === method) ? method : "";
}



function normalizeCwsAgentMergeCellsOp(op, index, defaultCenter = false) {
  const range = cwsAgentRangeForOp(op, index);
  if (range.top === range.bottom && range.left === range.right) {
    throw new Error(`Agent 操作 ${index + 1} の結合範囲が正しくありません。`);
  }
  return {
    type: "mergeCells",
    ...range,
    center: cwsAgentOptionalBoolean(op?.center ?? op?.centerContent ?? op?.mergeCenter, defaultCenter),
  };
}



function normalizeCwsAgentUnmergeCellsOp(op, index) {
  return {
    type: "unmergeCells",
    ...cwsAgentRangeForOp(op, index),
  };
}



function normalizeCwsAgentAddSheetOp(op, index) {
  const requestedName = cwsAgentSheetNameText(op?.name ?? op?.newName ?? op?.sheetName ?? op?.title);
  const fallbackName = uniqueSheetName(state.model?.sheets?.map((sheet) => sheet.name) || []);
  const validation = validateSheetName(requestedName || fallbackName, -1);
  if (!validation.valid) {
    throw new Error(`Agent 操作 ${index + 1}: ${validation.message}`);
  }
  return {
    type: "addSheet",
    name: validation.name,
    index: cwsAgentSheetInsertIndexForOp(op),
  };
}



function normalizeCwsAgentRenameSheetOp(op, index) {
  const newName = cwsAgentSheetNameText(op?.newName ?? op?.toName ?? op?.name ?? op?.title);
  if (!newName) {
    throw new Error(`Agent 操作 ${index + 1} の新しいシート名が正しくありません。`);
  }
  const sheetIndex = cwsAgentTargetSheetIndexForOp(op, {
    fallbackName: op?.oldName ?? op?.fromName ?? op?.targetName ?? "",
  });
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  const validation = validateSheetName(newName, sheetIndex);
  if (!validation.valid) {
    throw new Error(`Agent 操作 ${index + 1}: ${validation.message}`);
  }
  return {
    type: "renameSheet",
    sheetIndex,
    newName: validation.name,
  };
}



function normalizeCwsAgentDuplicateSheetOp(op, index) {
  const sourceSheetIndex = cwsAgentTargetSheetIndexForOp(op, {
    fallbackName: op?.sourceSheetName ?? op?.fromSheetName ?? op?.sourceName ?? op?.fromName ?? op?.sheetName ?? op?.targetName ?? "",
  });
  const source = state.model?.sheets?.[sourceSheetIndex];
  if (!source) {
    throw new Error(`Agent 操作 ${index + 1} のコピー元シートが見つかりません。`);
  }
  const requestedName = cwsAgentSheetNameText(op?.newName ?? op?.copyName ?? op?.name ?? op?.title);
  const validation = validateSheetName(requestedName || uniqueCopiedSheetName(source.name), -1);
  if (!validation.valid) {
    throw new Error(`Agent 操作 ${index + 1}: ${validation.message}`);
  }
  return {
    type: "duplicateSheet",
    sourceSheetIndex,
    name: validation.name,
    index: cwsAgentSheetInsertIndexForOp(op, sourceSheetIndex + 1),
  };
}



function normalizeCwsAgentMoveSheetOp(op, index) {
  const sourceSheetIndex = cwsAgentTargetSheetIndexForOp(op, {
    fallbackName: op?.sourceSheetName ?? op?.fromSheetName ?? op?.sourceName ?? op?.fromName ?? op?.sheetName ?? op?.targetName ?? "",
  });
  const source = state.model?.sheets?.[sourceSheetIndex];
  if (!source) {
    throw new Error(`Agent 操作 ${index + 1} の移動元シートが見つかりません。`);
  }
  if (!cwsAgentHasSheetPlacementTarget(op)) {
    throw new Error(`Agent 操作 ${index + 1} の移動先が指定されていません。`);
  }
  return {
    type: "moveSheet",
    sourceSheetIndex,
    beforeIndex: cwsAgentSheetInsertIndexForOp(op),
  };
}



function normalizeCwsAgentSheetTabColorOp(op, index, clearColor) {
  const sheetIndex = cwsAgentTargetSheetIndexForOp(op, {
    fallbackName: op?.sheetName ?? op?.targetSheetName ?? op?.targetName ?? op?.target ?? op?.name ?? "",
  });
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  const color = clearColor ? "" : cwsAgentSheetTabColor(op?.color ?? op?.tabColor ?? op?.sheetTabColor ?? op?.value);
  if (!clearColor && color === null) {
    throw new Error(`Agent 操作 ${index + 1} のシート見出し色が正しくありません。`);
  }
  return {
    type: "setSheetTabColor",
    sheetIndex,
    color: color || "",
  };
}



function normalizeCwsAgentDeleteSheetOp(op, index) {
  const sheetIndex = cwsAgentTargetSheetIndexForOp(op, {
    fallbackName: op?.sheetName ?? op?.name ?? op?.targetName ?? "",
  });
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  if (state.model.sheets.length <= 1) {
    throw new Error("少なくとも 1 つのワークシートを残す必要があります。");
  }
  if (isSheetVisible(sheet) && visibleSheetIndexes().length <= 1) {
    throw new Error("少なくとも 1 つの表示ワークシートを残す必要があります。");
  }
  return {
    type: "deleteSheet",
    sheetIndex,
  };
}



function normalizeCwsAgentSheetVisibilityOp(op, index, hidden) {
  const sheetIndex = cwsAgentSheetVisibilityTargetIndexForOp(op, hidden);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  if (hidden && isSheetVisible(sheet) && visibleSheetIndexes().length <= 1) {
    throw new Error("少なくとも 1 つの表示ワークシートを残す必要があります。");
  }
  return {
    type: hidden ? "hideSheet" : "unhideSheet",
    sheetIndex,
  };
}



function normalizeCwsAgentSetFreezePanesOp(op, index) {
  const sheetIndex = cwsAgentSheetIndexForOp(op, cwsAgentFreezePaneAddress(op)?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }

  const mode = cwsAgentFreezePaneMode(op?.mode ?? op?.kind ?? op?.preset ?? op?.target);
  let rows = cwsAgentNonNegativeInteger(op?.rows ?? op?.rowCount ?? op?.freezeRows ?? op?.topRows);
  let columns = cwsAgentNonNegativeInteger(op?.columns ?? op?.cols ?? op?.columnCount ?? op?.freezeColumns ?? op?.leftColumns);
  if (mode === "topRow") {
    rows = 1;
    columns = 0;
  } else if (mode === "firstColumn") {
    rows = 0;
    columns = 1;
  } else if (mode === "selection") {
    const point = cwsAgentFreezePaneAddress(op);
    rows = Math.max(0, (point?.row || 1) - 1);
    columns = Math.max(0, (point?.col || 1) - 1);
  }

  if (rows === null && columns === null) {
    const point = cwsAgentFreezePaneAddress(op);
    if (point) {
      rows = Math.max(0, point.row - 1);
      columns = Math.max(0, point.col - 1);
    }
  }
  if (rows === null && columns === null) {
    rows = cwsAgentNonNegativeInteger(op?.row ?? op?.topRow);
    columns = cwsAgentNonNegativeInteger(op?.col ?? op?.column ?? op?.leftColumn);
  }

  rows = rows ?? 0;
  columns = columns ?? 0;
  if (rows <= 0 && columns <= 0) {
    throw new Error(`Agent 操作 ${index + 1} の固定行・列指定が正しくありません。`);
  }
  if (rows > CWS_AGENT_MAX_FREEZE_ROWS || columns > CWS_AGENT_MAX_FREEZE_COLUMNS) {
    throw new Error(`Agent 操作 ${index + 1} の固定行・列指定が大きすぎます。`);
  }

  return {
    type: "setFreezePanes",
    sheetIndex,
    rows,
    columns,
  };
}



function normalizeCwsAgentClearFreezePanesOp(op, index) {
  const sheetIndex = cwsAgentSheetIndexForOp(op, cwsAgentFreezePaneAddress(op)?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: "clearFreezePanes",
    sheetIndex,
  };
}



function normalizeCwsAgentSetSheetViewOp(op, index, normalizedType = "") {
  const sheetIndex = cwsAgentTargetSheetIndexForOp(op, {
    fallbackName: op?.sheetName ?? op?.targetSheetName ?? op?.targetName ?? "",
  });
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }

  const patch = {};
  const directMode = normalizedType === "normalview"
    ? "normal"
    : normalizedType === "pagelayoutview"
      ? "page-layout"
      : (normalizedType === "pagebreakview" || normalizedType === "pagebreakpreview")
        ? "page-break"
        : "";
  const viewMode = directMode || cwsAgentSheetViewMode(cwsAgentNestedValueForOp(op, ["viewMode", "mode", "view", "sheetView", "target", "value"]));
  if (viewMode) patch.viewMode = viewMode;

  const zoomValue = cwsAgentNestedValueForOp(op, ["zoomScale", "zoom", "scale", "percent", "percentage"]);
  const zoomScale = cwsAgentZoomScaleValue(zoomValue);
  if (zoomScale !== null) patch.zoomScale = zoomScale;

  const boolFields = {
    showGridLines: ["showGridLines", "gridLines", "displayGridLines"],
    showHeadings: ["showHeadings", "headings", "rowColumnHeadings", "showRowColHeaders"],
    showRuler: ["showRuler", "ruler"],
    showFormulas: ["showFormulas", "formulas", "formulaView"],
    showZeros: ["showZeros", "zeroValues", "zeros"],
    showWhiteSpace: ["showWhiteSpace", "whiteSpace", "pageWhiteSpace"],
    showOutlineSymbols: ["showOutlineSymbols", "outlineSymbols"],
    showObjects: ["showObjects", "objects"],
    rightToLeft: ["rightToLeft", "rtl"],
  };
  for (const [field, keys] of Object.entries(boolFields)) {
    const value = cwsAgentNestedValueForOp(op, keys);
    if (value !== undefined) patch[field] = cwsAgentBoolean(value);
  }

  if (!Object.keys(patch).length) {
    throw new Error(`Agent 操作 ${index + 1} の表示設定が正しくありません。`);
  }
  return {
    type: "setSheetView",
    sheetIndex,
    patch,
  };
}



function normalizeCwsAgentSetPrintAreaOp(op, index) {
  const range = cwsAgentPageSetupRangeForOp(op, index, ["printArea", "range", "targetRange", "address", "ref", "cell"]);
  return {
    type: "setPrintArea",
    ...range,
  };
}



function normalizeCwsAgentClearPrintAreaOp(op, index) {
  const parsedRange = cwsAgentParseRangeAddress(op?.printArea || op?.range || op?.targetRange || op?.address || op?.ref || op?.cell || "");
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: "clearPrintArea",
    sheetIndex,
  };
}



function normalizeCwsAgentSetPrintTitlesOp(op, index) {
  const rowRange = cwsAgentPrintTitleRowsForOp(op, index);
  const columnRange = cwsAgentPrintTitleColumnsForOp(op, index);
  if (!rowRange && !columnRange) {
    throw new Error(`Agent 操作 ${index + 1} の印刷タイトル行または列が正しくありません。`);
  }
  const sheetIndex = rowRange?.sheetIndex ?? columnRange?.sheetIndex ?? state.activeSheetIndex;
  if ((rowRange && rowRange.sheetIndex !== sheetIndex) || (columnRange && columnRange.sheetIndex !== sheetIndex)) {
    throw new Error(`Agent 操作 ${index + 1} の印刷タイトル行と列は同じシートを指定してください。`);
  }
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return {
    type: "setPrintTitles",
    sheetIndex,
    ...(rowRange ? { rows: rowRange } : {}),
    ...(columnRange ? { cols: columnRange } : {}),
  };
}



function normalizeCwsAgentClearPrintTitlesOp(op, index) {
  const parsedRows = cwsAgentParseRowRangeAddress(op?.rows || op?.rowRange || op?.titleRows || op?.printTitleRows || "");
  const parsedCols = cwsAgentParseColumnRangeAddress(op?.cols || op?.columns || op?.columnRange || op?.titleColumns || op?.printTitleColumns || "");
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRows?.sheetName || parsedCols?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  const target = String(op?.target ?? op?.axis ?? op?.kind ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const clearRows = target === "rows" || target === "row" || target === "titlerows" || target === "printtitlerows" || Boolean(parsedRows);
  const clearCols = target === "columns" || target === "column" || target === "cols" || target === "col" || target === "titlecolumns" || target === "printtitlecolumns" || Boolean(parsedCols);
  return {
    type: "clearPrintTitles",
    sheetIndex,
    rows: clearRows || (!clearRows && !clearCols),
    cols: clearCols || (!clearRows && !clearCols),
  };
}



function normalizeCwsAgentSetHeaderFooterOp(op, index, normalizedType = "") {
  const sheetIndex = cwsAgentPageSetupSheetIndexForOp(op, index);
  const settings = cwsAgentHeaderFooterSettingsForOp(op, index, normalizedType);
  if (!Object.keys(settings).length) {
    throw new Error(`Agent 操作 ${index + 1} のヘッダー/フッター指定が正しくありません。`);
  }
  return {
    type: "setHeaderFooter",
    sheetIndex,
    settings,
  };
}



function normalizeCwsAgentClearHeaderFooterOp(op, index, normalizedType = "") {
  return {
    type: "clearHeaderFooter",
    sheetIndex: cwsAgentPageSetupSheetIndexForOp(op, index),
    target: cwsAgentHeaderFooterClearTarget(op, normalizedType),
  };
}



function cwsAgentHeaderFooterSettingsForOp(op, index, normalizedType = "") {
  const settings = {};
  const setSection = (section, value, fullSection = false) => {
    if (value === undefined) return;
    const normalized = cwsAgentHeaderFooterSection(value, fullSection);
    if (!normalized) {
      throw new Error(`Agent 操作 ${index + 1} のヘッダー/フッター指定が正しくありません。`);
    }
    settings[section] = { ...(settings[section] || {}), ...normalized };
  };

  setSection("oddHeader", cwsAgentHeaderFooterValueForOp(op, ["header", "headerText", "pageHeader", "printHeader"]), true);
  setSection("oddFooter", cwsAgentHeaderFooterValueForOp(op, ["footer", "footerText", "pageFooter", "printFooter"]), true);
  setSection("oddHeader", cwsAgentHeaderFooterValueForOp(op, ["oddHeader", "defaultHeader"]));
  setSection("oddFooter", cwsAgentHeaderFooterValueForOp(op, ["oddFooter", "defaultFooter"]));
  setSection("firstHeader", cwsAgentHeaderFooterValueForOp(op, ["firstHeader", "firstPageHeader"]));
  setSection("firstFooter", cwsAgentHeaderFooterValueForOp(op, ["firstFooter", "firstPageFooter"]));
  setSection("evenHeader", cwsAgentHeaderFooterValueForOp(op, ["evenHeader", "evenPageHeader"]));
  setSection("evenFooter", cwsAgentHeaderFooterValueForOp(op, ["evenFooter", "evenPageFooter"]));
  if (!Object.keys(settings).length) {
    const value = cwsAgentHeaderFooterValueForOp(op, ["value", "text"]);
    if (value !== undefined) {
      if (normalizedType.includes("footer")) {
        setSection("oddFooter", value, true);
      } else if (normalizedType.includes("header")) {
        setSection("oddHeader", value, true);
      }
    }
  }

  const boolFields = {
    differentFirst: ["differentFirst", "differentFirstPage"],
    differentOddEven: ["differentOddEven", "differentEvenOdd", "differentOddAndEven"],
    scaleWithDoc: ["scaleWithDoc", "scaleWithDocument"],
    alignWithMargins: ["alignWithMargins", "alignToMargins"],
  };
  for (const [field, keys] of Object.entries(boolFields)) {
    const value = cwsAgentHeaderFooterValueForOp(op, keys);
    if (value !== undefined) settings[field] = cwsAgentBoolean(value);
  }
  return settings;
}



function cwsAgentHeaderFooterValueForOp(op, keys) {
  const sources = [op, op?.settings, op?.headerFooter, op?.patch, op?.options];
  for (const source of sources) {
    if (!source || typeof source !== "object") continue;
    for (const key of keys) {
      if (Object.prototype.hasOwnProperty.call(source, key)) return source[key];
    }
  }
  return undefined;
}



function cwsAgentHeaderFooterSection(value, fullSection = false) {
  if (value === undefined) return null;
  if (value === null || Array.isArray(value)) return null;
  if (typeof value !== "object") {
    const text = cwsAgentHeaderFooterText(value);
    return text === null ? null : { left: "", center: text, right: "" };
  }
  const section = fullSection ? { left: "", center: "", right: "" } : {};
  const aliases = {
    left: ["left", "leftText", "l"],
    center: ["center", "middle", "centerText", "middleText", "c"],
    right: ["right", "rightText", "r"],
  };
  for (const [field, keys] of Object.entries(aliases)) {
    const raw = cwsAgentHeaderFooterValueFromSource(value, keys);
    if (raw === undefined) continue;
    const text = cwsAgentHeaderFooterText(raw);
    if (text === null) return null;
    section[field] = text;
  }
  return Object.keys(section).length ? section : null;
}



function cwsAgentHeaderFooterValueFromSource(source, keys) {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(source || {}, key)) return source[key];
  }
  return undefined;
}



function cwsAgentHeaderFooterText(value) {
  if (value === undefined || value === null || Array.isArray(value) || typeof value === "object") return null;
  const text = String(value).replace(/\r\n?/g, "\n");
  return text.length <= 1000 ? text : text.slice(0, 1000);
}



function cwsAgentHeaderFooterClearTarget(op, normalizedType = "") {
  const raw = cwsAgentHeaderFooterValueForOp(op, ["target", "kind", "section", "scope"]);
  const text = String(raw ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["header", "headers", "pageheader", "printheader", "ヘッダー", "页眉", "頁眉"].includes(text)) return "header";
  if (["footer", "footers", "pagefooter", "printfooter", "フッター", "页脚", "頁腳"].includes(text)) return "footer";
  if (!normalizedType.includes("headerfooter") && !normalizedType.includes("headerandfooter")) {
    if (normalizedType.includes("footer")) return "footer";
    if (normalizedType.includes("header")) return "header";
  }
  return "all";
}



function normalizeCwsAgentSetPageThemeOp(op, index, normalizedType = "") {
  const sheetIndex = cwsAgentPageSetupSheetIndexForOp(op, index);
  const reset = normalizedType.startsWith("reset");
  const sourceValue = reset ? "office" : cwsAgentNestedValueForOp(op, ["theme", "themeKey", "scheme", "value", "name", "target"]);
  const includeEffectsOnly = normalizedType === "setthemeeffects" || normalizedType === "seteffectscheme" || cwsAgentPageThemeScope(op) === "effects";
  const themeKey = cwsAgentPageThemeKey(sourceValue, includeEffectsOnly);
  if (!themeKey) {
    throw new Error(`Agent 操作 ${index + 1} のテーマ指定が正しくありません。`);
  }
  const includes = cwsAgentPageThemeIncludes(op, normalizedType);
  return {
    type: "setPageTheme",
    sheetIndex,
    theme: themeKey,
    ...includes,
  };
}



function cwsAgentPageThemeIncludes(op, normalizedType = "") {
  const scope = cwsAgentPageThemeScope(op);
  const boolValue = (keys) => {
    const value = cwsAgentNestedValueForOp(op, keys);
    return value === undefined ? undefined : cwsAgentBoolean(value);
  };
  let includeColors = boolValue(["includeColors", "colors", "themeColors"]);
  let includeFonts = boolValue(["includeFonts", "fonts", "themeFonts"]);
  let includeEffects = boolValue(["includeEffects", "effects", "themeEffects"]);

  if (normalizedType === "setthemecolors" || normalizedType === "setcolorscheme" || scope === "colors") {
    includeColors = true;
    includeFonts = false;
    includeEffects = false;
  } else if (normalizedType === "setthemefonts" || normalizedType === "setfontscheme" || scope === "fonts") {
    includeColors = false;
    includeFonts = true;
    includeEffects = false;
  } else if (normalizedType === "setthemeeffects" || normalizedType === "seteffectscheme" || scope === "effects") {
    includeColors = false;
    includeFonts = false;
    includeEffects = true;
  }

  if (includeColors === undefined && includeFonts === undefined && includeEffects === undefined) {
    includeColors = true;
    includeFonts = true;
    includeEffects = true;
  }
  return {
    includeColors: includeColors === true,
    includeFonts: includeFonts === true,
    includeEffects: includeEffects === true,
  };
}



function cwsAgentPageThemeScope(op) {
  const raw = cwsAgentNestedValueForOp(op, ["scope", "kind", "targetKind", "part"]);
  const text = String(raw ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["color", "colors", "colorscheme", "themecolors", "配色", "色"].includes(text)) return "colors";
  if (["font", "fonts", "fontscheme", "themefonts", "フォント", "字体", "字體"].includes(text)) return "fonts";
  if (["effect", "effects", "effectscheme", "themeeffects", "効果", "效果"].includes(text)) return "effects";
  return "";
}



function cwsAgentPageThemeKey(value, allowEffects = false) {
  const normalized = String(value ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (!normalized) return "";
  const normalizeCandidate = (candidate) => String(candidate ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const theme = PAGE_LAYOUT_THEMES.find((item) =>
    normalizeCandidate(item.value) === normalized ||
    normalizeCandidate(item.label) === normalized
  );
  if (theme) return theme.value;
  if (allowEffects) {
    const effect = PAGE_LAYOUT_EFFECT_PRESETS.find((item) =>
      normalizeCandidate(item.value) === normalized ||
      normalizeCandidate(item.label) === normalized
    );
    if (effect) return effect.value;
  }
  return "";
}



function normalizeCwsAgentSetPageOrientationOp(op, index) {
  const sheetIndex = cwsAgentPageSetupSheetIndexForOp(op, index);
  const orientation = cwsAgentPageOrientationValue(op?.orientation ?? op?.value ?? op?.mode ?? op?.target ?? op?.direction);
  if (!orientation) {
    throw new Error(`Agent 操作 ${index + 1} の印刷の向きが正しくありません。`);
  }
  return {
    type: "setPageOrientation",
    sheetIndex,
    orientation,
  };
}



function normalizeCwsAgentSetPaperSizeOp(op, index) {
  const sheetIndex = cwsAgentPageSetupSheetIndexForOp(op, index);
  const paperSize = cwsAgentPaperSizeValue(op?.paperSize ?? op?.size ?? op?.value ?? op?.paper ?? op?.target);
  if (!paperSize) {
    throw new Error(`Agent 操作 ${index + 1} の用紙サイズが正しくありません。`);
  }
  return {
    type: "setPaperSize",
    sheetIndex,
    paperSize,
  };
}



function normalizeCwsAgentSetPageMarginsOp(op, index) {
  const sheetIndex = cwsAgentPageSetupSheetIndexForOp(op, index);
  const preset = cwsAgentPageMarginPreset(op?.marginPreset ?? op?.preset ?? op?.value ?? op?.target ?? op?.kind);
  const margins = {};
  const explicitMargins = {};
  const source = op?.margins && typeof op.margins === "object" ? op.margins : op;
  for (const key of ["top", "bottom", "left", "right", "header", "footer"]) {
    const value = cwsAgentPageMarginValue(source?.[key]);
    if (value !== null) explicitMargins[key] = value;
  }
  if (preset) {
    Object.assign(margins, PAGE_LAYOUT_MARGIN_PRESETS[preset]);
    delete margins.label;
  }
  Object.assign(margins, explicitMargins);
  if (!Object.keys(margins).length) {
    throw new Error(`Agent 操作 ${index + 1} の余白指定が正しくありません。`);
  }
  return {
    type: "setPageMargins",
    sheetIndex,
    ...(preset ? { marginPreset: preset } : {}),
    margins,
  };
}



function normalizeCwsAgentSetPageScaleOp(op, index) {
  const sheetIndex = cwsAgentPageSetupSheetIndexForOp(op, index);
  const scale = cwsAgentPageScaleValue(op?.scale ?? op?.percent ?? op?.percentage ?? op?.zoom ?? op?.value);
  if (scale !== null) {
    return {
      type: "setPageScale",
      sheetIndex,
      scale,
      fitToPage: false,
      fitToWidth: "auto",
      fitToHeight: "auto",
    };
  }

  const mode = String(op?.mode ?? op?.fit ?? op?.preset ?? op?.target ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  let fitToWidth = cwsAgentFitToPageValue(op?.fitToWidth ?? op?.widthPages ?? op?.pagesWide ?? op?.wide);
  let fitToHeight = cwsAgentFitToPageValue(op?.fitToHeight ?? op?.heightPages ?? op?.pagesTall ?? op?.tall);
  if (fitToWidth === null || fitToHeight === null) {
    throw new Error(`Agent 操作 ${index + 1} の拡大縮小印刷指定が正しくありません。`);
  }
  if (mode === "onepage" || mode === "fitonepage" || mode === "singlepage" || mode === "1page") {
    fitToWidth = "1";
    fitToHeight = "1";
  } else if (mode === "fitwidth" || mode === "fitpageswide" || mode === "width") {
    fitToWidth = fitToWidth ?? "1";
    fitToHeight = fitToHeight ?? "auto";
  } else if (mode === "fitheight" || mode === "fitpagestall" || mode === "height") {
    fitToWidth = fitToWidth ?? "auto";
    fitToHeight = fitToHeight ?? "1";
  } else if (cwsAgentBoolean(op?.fitToPage)) {
    fitToWidth = fitToWidth ?? "1";
    fitToHeight = fitToHeight ?? "1";
  }
  if (fitToWidth === undefined && fitToHeight === undefined) {
    throw new Error(`Agent 操作 ${index + 1} の拡大縮小印刷指定が正しくありません。`);
  }
  return {
    type: "setPageScale",
    sheetIndex,
    fitToPage: true,
    fitToWidth: fitToWidth ?? "auto",
    fitToHeight: fitToHeight ?? "auto",
    scale: 100,
  };
}



function normalizeCwsAgentSetPrintOptionsOp(op, index) {
  const sheetIndex = cwsAgentPageSetupSheetIndexForOp(op, index);
  const patch = {};
  const boolFields = {
    printGridLines: ["printGridLines", "gridLines", "showGridLines"],
    printHeadings: ["printHeadings", "headings", "rowColumnHeadings", "showHeadings", "showRowColHeaders"],
    centerHorizontal: ["centerHorizontal", "horizontalCentered", "centerHorizontally"],
    centerVertical: ["centerVertical", "verticalCentered", "centerVertically"],
    blackAndWhite: ["blackAndWhite", "blackWhite", "monochrome", "grayscale"],
    draftQuality: ["draftQuality", "draft"],
  };
  for (const [field, keys] of Object.entries(boolFields)) {
    const value = cwsAgentPrintOptionValue(op, keys);
    if (value !== undefined) patch[field] = cwsAgentBoolean(value);
  }

  const comments = cwsAgentPrintCommentsValue(cwsAgentPrintOptionValue(op, ["comments", "cellComments", "printComments"]));
  if (comments !== undefined) patch.comments = comments;
  const cellErrors = cwsAgentPrintCellErrorsValue(cwsAgentPrintOptionValue(op, ["cellErrors", "errors", "printErrors"]));
  if (cellErrors !== undefined) patch.cellErrors = cellErrors;
  const pageOrder = cwsAgentPrintPageOrderValue(cwsAgentPrintOptionValue(op, ["pageOrder", "order", "printOrder"]));
  if (pageOrder !== undefined) patch.pageOrder = pageOrder;
  const firstPageNumber = cwsAgentFirstPageNumberValue(cwsAgentPrintOptionValue(op, ["firstPageNumber", "firstPage", "startPage", "startPageNumber"]));
  if (firstPageNumber !== undefined) patch.firstPageNumber = firstPageNumber;

  if (!Object.keys(patch).length) {
    throw new Error(`Agent 操作 ${index + 1} の印刷オプション指定が正しくありません。`);
  }
  return {
    type: "setPrintOptions",
    sheetIndex,
    patch,
  };
}



function normalizeCwsAgentPageBreakOp(op, index, type) {
  const sheetIndex = cwsAgentPageSetupSheetIndexForOp(op, index);
  const axis = cwsAgentPageBreakAxis(op);
  const breakIndex = cwsAgentPageBreakIndex(op, axis);
  if (!axis || !breakIndex) {
    throw new Error(`Agent 操作 ${index + 1} の改ページ位置が正しくありません。`);
  }
  const limit = axis === "row" ? EXCEL_MAX_ROWS : EXCEL_MAX_COLS;
  if (breakIndex <= 1 || breakIndex > limit) {
    throw new Error(`Agent 操作 ${index + 1} の改ページ位置が正しくありません。`);
  }
  return {
    type,
    sheetIndex,
    axis,
    index: breakIndex,
  };
}



function normalizeCwsAgentResetPageBreaksOp(op, index) {
  return {
    type: "resetPageBreaks",
    sheetIndex: cwsAgentPageSetupSheetIndexForOp(op, index),
  };
}



function cwsAgentPageSetupSheetIndexForOp(op, index) {
  const sheetIndex = cwsAgentSheetIndexForOp(op);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return sheetIndex;
}



function cwsAgentPageOrientationValue(value) {
  const text = String(value ?? "").trim().toLowerCase();
  if (!text) return "";
  if (["landscape", "horizontal", "sideways", "横", "横向き", "横方向", "横置き", "横向", "水平"].includes(text)) return "landscape";
  if (["portrait", "vertical", "縦", "縦向き", "縦方向", "縦置き", "竖向", "縦向", "垂直"].includes(text)) return "portrait";
  return "";
}



function cwsAgentPaperSizeValue(value) {
  const text = String(value ?? "").trim();
  if (!text) return "";
  const normalized = text.toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    a4: "A4",
    a3: "A3",
    b5: "B5",
    letter: "Letter",
    usletter: "Letter",
    レター: "Letter",
  };
  return aliases[normalized] || "";
}



function cwsAgentPageMarginPreset(value) {
  const text = String(value ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (!text) return "";
  if (["normal", "standard", "default", "標準", "普通"].includes(text)) return "normal";
  if (["wide", "広い", "寬", "宽"].includes(text)) return "wide";
  if (["narrow", "狭い", "狭め", "窄"].includes(text)) return "narrow";
  return "";
}



function cwsAgentPageMarginValue(value) {
  if (value === null || value === undefined || value === "") return null;
  const text = String(value).trim().toLowerCase();
  let number = Number(text.replace(/"/g, ""));
  if (!Number.isFinite(number)) {
    const match = text.match(/^(-?\d+(?:\.\d+)?)\s*(cm|mm|in|inch|inches)?$/);
    if (!match) return null;
    number = Number(match[1]);
    if (match[2] === "cm") number /= 2.54;
    if (match[2] === "mm") number /= 25.4;
  }
  if (!Number.isFinite(number)) return null;
  return Math.max(0, Math.min(3, Math.round(number * 100) / 100));
}



function cwsAgentPageScaleValue(value) {
  if (value === null || value === undefined || value === "") return null;
  const text = String(value).trim().replace(/%$/, "");
  const number = Number(text);
  if (!Number.isFinite(number)) return null;
  return Math.max(10, Math.min(400, Math.round(number)));
}



function cwsAgentFitToPageValue(value) {
  if (value === null || value === undefined || value === "") return undefined;
  const text = String(value).trim().toLowerCase();
  if (["auto", "none", "なし", "自動"].includes(text)) return "auto";
  const number = Number(text);
  if (!Number.isFinite(number) || number < 1) return null;
  return String(Math.max(1, Math.min(99, Math.round(number))));
}



function cwsAgentPrintOptionValue(op, keys) {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(op || {}, key)) return op[key];
  }
  const options = op?.options && typeof op.options === "object" ? op.options : null;
  if (options) {
    for (const key of keys) {
      if (Object.prototype.hasOwnProperty.call(options, key)) return options[key];
    }
  }
  return undefined;
}



function cwsAgentPrintCommentsValue(value) {
  if (value === undefined || value === null || value === "") return undefined;
  const text = String(value).trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["none", "no", "false", "off", "なし", "無し", "不要"].includes(text)) return "none";
  if (["atend", "end", "sheetend", "末尾", "シートの末尾"].includes(text)) return "atEnd";
  if (["asdisplayed", "displayed", "visible", "表示", "画面表示", "画面表示イメージ"].includes(text)) return "asDisplayed";
  return undefined;
}



function cwsAgentPrintCellErrorsValue(value) {
  if (value === undefined || value === null || value === "") return undefined;
  const text = String(value).trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["displayed", "show", "visible", "表示", "表示する"].includes(text)) return "displayed";
  if (["blank", "empty", "空白", "空"].includes(text)) return "blank";
  if (["dash", "-", "--", "hyphen"].includes(text)) return "dash";
  if (["na", "#n/a", "n/a"].includes(text)) return "na";
  return undefined;
}



function cwsAgentPrintPageOrderValue(value) {
  if (value === undefined || value === null || value === "") return undefined;
  const text = String(value).trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["downthenover", "downover", "down", "下から右", "下へ次に右"].includes(text)) return "downThenOver";
  if (["overthendown", "overdown", "right", "across", "右から下", "右へ次に下"].includes(text)) return "overThenDown";
  return undefined;
}



function cwsAgentFirstPageNumberValue(value) {
  if (value === undefined || value === null || value === "") return undefined;
  const number = Number(value);
  if (!Number.isFinite(number)) return undefined;
  return Math.max(1, Math.min(9999, Math.round(number)));
}



function cwsAgentPageBreakAxis(op) {
  const raw = String(op?.axis ?? op?.kind ?? op?.target ?? op?.direction ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["row", "rows", "horizontal", "h", "行"].includes(raw)) return "row";
  if (["col", "cols", "column", "columns", "vertical", "v", "列"].includes(raw)) return "col";
  if (op?.row !== undefined || op?.breakRow !== undefined || op?.rowIndex !== undefined || op?.beforeRow !== undefined) return "row";
  if (op?.col !== undefined || op?.column !== undefined || op?.breakCol !== undefined || op?.breakColumn !== undefined || op?.columnIndex !== undefined || op?.beforeColumn !== undefined || op?.beforeCol !== undefined) return "col";
  return "";
}



function cwsAgentPageBreakIndex(op, axis) {
  if (axis === "row") {
    return cwsAgentPositiveInteger(op?.row ?? op?.breakRow ?? op?.rowIndex ?? op?.beforeRow ?? op?.index ?? op?.position);
  }
  if (axis === "col") {
    return cwsAgentColumnValue(op?.col ?? op?.column ?? op?.breakCol ?? op?.breakColumn ?? op?.columnIndex ?? op?.beforeColumn ?? op?.beforeCol ?? op?.index ?? op?.position);
  }
  return 0;
}



function cwsAgentPageSetupRangeForOp(op, index, keys) {
  const parsedRange = cwsAgentParseRangeAddress(keys.map((key) => op?.[key]).find((value) => String(value ?? "").trim()) || "");
  const top = cwsAgentPositiveInteger(op?.top || op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.left || op?.col) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のページ設定範囲が正しくありません。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentPrintTitleRowsForOp(op, index) {
  const text = String(op?.rows ?? op?.rowRange ?? op?.titleRows ?? op?.printTitleRows ?? "").trim();
  const parsed = cwsAgentParseRowRangeAddress(text) || cwsAgentParseRangeAddress(text);
  let top = cwsAgentPositiveInteger(op?.topRow ?? op?.startRow ?? op?.row ?? op?.rows);
  let bottom = cwsAgentPositiveInteger(op?.bottomRow ?? op?.endRow);
  if (!top && parsed) top = parsed.top;
  if (!bottom && parsed) bottom = parsed.bottom;
  if (!top && !bottom) return null;
  bottom = bottom || top;
  if (!top || bottom < top || bottom > EXCEL_MAX_ROWS || (bottom - top + 1) > CWS_AGENT_MAX_STRUCTURE_ROWS) {
    throw new Error(`Agent 操作 ${index + 1} の印刷タイトル行が正しくありません。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsed?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return normalizePageSetupPrintTitleRows({ sheetIndex, top, bottom }, sheet);
}



function cwsAgentPrintTitleColumnsForOp(op, index) {
  const text = String(op?.cols ?? op?.columns ?? op?.columnRange ?? op?.titleColumns ?? op?.printTitleColumns ?? "").trim();
  const parsed = cwsAgentParseColumnRangeAddress(text) || cwsAgentParseRangeAddress(text);
  let left = cwsAgentColumnValue(op?.leftCol ?? op?.startCol ?? op?.col ?? op?.column ?? op?.cols ?? op?.columns);
  let right = cwsAgentColumnValue(op?.rightCol ?? op?.endCol);
  if (!left && parsed) left = parsed.left;
  if (!right && parsed) right = parsed.right;
  if (!left && !right) return null;
  right = right || left;
  if (!left || right < left || right > EXCEL_MAX_COLS || (right - left + 1) > CWS_AGENT_MAX_STRUCTURE_COLUMNS) {
    throw new Error(`Agent 操作 ${index + 1} の印刷タイトル列が正しくありません。`);
  }
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsed?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return normalizePageSetupPrintTitleColumns({ sheetIndex, left, right }, sheet);
}



function cwsAgentRangeForOp(op, index, values = null, maxCells = CWS_AGENT_MAX_RANGE_CELLS) {
  const parsedRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  const top = cwsAgentPositiveInteger(op?.top || op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.left || op?.col) || parsedRange?.left || 0;
  if (!top || !left || top > EXCEL_MAX_ROWS || left > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の範囲が正しくありません。`);
  }

  const valueHeight = values ? values.length : 1;
  const valueWidth = values ? Math.max(...values.map((row) => row.length)) : 1;
  let bottom = cwsAgentPositiveInteger(op?.bottom) || parsedRange?.bottom || top;
  let right = cwsAgentColumnValue(op?.right) || parsedRange?.right || left;
  if (values && bottom === top && right === left) {
    bottom = top + valueHeight - 1;
    right = left + valueWidth - 1;
  }
  if (bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の範囲が正しくありません。`);
  }
  if (values && ((bottom - top + 1) !== valueHeight || (right - left + 1) !== valueWidth)) {
    throw new Error(`Agent 操作 ${index + 1} の範囲とデータ数が一致しません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount > maxCells) {
    throw new Error(`Agent 操作 ${index + 1} の範囲が大きすぎます。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentNamedRangeForOp(op, index) {
  const parsedRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || op?.sourceRange || "");
  const top = cwsAgentPositiveInteger(op?.top || op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.left || op?.col) || parsedRange?.left || 0;
  if (!top || !left || top > EXCEL_MAX_ROWS || left > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の名前定義範囲が正しくありません。`);
  }

  const bottom = cwsAgentPositiveInteger(op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.right) || parsedRange?.right || left;
  if (bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の名前定義範囲が正しくありません。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentValidationRangeForOp(op, index) {
  const parsedRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  const top = cwsAgentPositiveInteger(op?.top || op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.left || op?.col) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の入力規則範囲が正しくありません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount > CWS_AGENT_MAX_VALIDATION_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} の入力規則範囲が大きすぎます。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentSortRangeForOp(op, index) {
  const parsedRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  const top = cwsAgentPositiveInteger(op?.top || op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.left) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の並べ替え範囲が正しくありません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount > CWS_AGENT_MAX_SORT_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} の並べ替え範囲が大きすぎます。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentFilterRangeForOp(op, index, options = {}) {
  const parsedRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  const hasRange = Boolean(
    String(op?.range || op?.address || op?.ref || op?.cell || "").trim() ||
    op?.top || op?.row || op?.left || op?.bottom || op?.right,
  );
  if (!hasRange && options.allowCurrent && sheet.autoFilter?.range) {
    return { sheetIndex, ...sheet.autoFilter.range };
  }

  const top = cwsAgentPositiveInteger(op?.top || op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.left) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のフィルター範囲が正しくありません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount > CWS_AGENT_MAX_FILTER_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} のフィルター範囲が大きすぎます。`);
  }

  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentAdvancedFilterRangeForValue(value, fallbackOp, index, label) {
  const range = cwsAgentFlexibleRangeForOp(value, fallbackOp, index, label, CWS_AGENT_MAX_ADVANCED_FILTER_CELLS);
  const sheet = state.model?.sheets?.[range.sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} の${label}のシートが見つかりません。`);
  }
  return range;
}



function cwsAgentAdvancedFilterAction(op, copyValue) {
  const source = String(op?.action ?? op?.mode ?? op?.target ?? op?.outputMode ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["copy", "copyto", "extract", "extractto", "destination", "output", "newrange", "specifiedrange", "指定した範囲", "指定範囲", "抽出", "抽出先", "コピー", "复制", "複製"].includes(source)) {
    return "copy";
  }
  if (copyValue !== null && copyValue !== undefined && (typeof copyValue === "object" || String(copyValue).trim())) {
    return "copy";
  }
  return "in-place";
}



function cwsAgentRemoveDuplicatesRangeForOp(op, index) {
  const parsedRange = cwsAgentParseRangeAddress(op?.range || op?.address || op?.ref || op?.cell || "");
  const top = cwsAgentPositiveInteger(op?.top || op?.row) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.left || op?.col) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.bottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.right) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} の重複削除範囲が正しくありません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount > CWS_AGENT_MAX_DEDUP_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} の重複削除範囲が大きすぎます。`);
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  return { sheetIndex, top, left, bottom, right };
}



function cwsAgentSheetIndexForOp(op, parsedSheetName = "") {
  const byIndex = cwsAgentInteger(op?.sheetIndex);
  if (byIndex !== null) {
    return byIndex;
  }

  const sheetName = String(op?.sheetName || parsedSheetName || "").trim();
  if (sheetName) {
    const found = (state.model?.sheets || []).findIndex((sheet) => sheet.name === sheetName);
    if (found >= 0) return found;
  }

  return state.activeSheetIndex;
}



function cwsAgentTargetSheetIndexForOp(op, options = {}) {
  const byIndex = cwsAgentInteger(op?.sheetIndex);
  if (byIndex !== null) return byIndex;
  const fallbackName = cwsAgentSheetNameText(options.fallbackName);
  const sheetName = fallbackName || cwsAgentSheetNameText(
    op?.sheetName ??
    op?.targetSheetName ??
    op?.target ??
    "",
  );
  if (sheetName) return cwsAgentSheetIndexByName(sheetName);
  return state.activeSheetIndex;
}



function cwsAgentSheetVisibilityTargetIndexForOp(op, allowActiveFallback) {
  const byIndex = cwsAgentInteger(op?.sheetIndex);
  if (byIndex !== null) return byIndex;
  const sheetName = cwsAgentSheetNameText(
    op?.sheetName ??
    op?.targetSheetName ??
    op?.targetName ??
    op?.target ??
    op?.name ??
    op?.title ??
    "",
  );
  if (sheetName) return cwsAgentSheetIndexByName(sheetName);
  return allowActiveFallback ? state.activeSheetIndex : -1;
}



function cwsAgentSheetIndexByName(sheetName) {
  const target = String(sheetName || "").trim();
  if (!target) return -1;
  return (state.model?.sheets || []).findIndex((sheet) => sheet.name === target);
}



function cwsAgentSheetInsertIndexForOp(op, fallbackIndex = null) {
  const beforeName = cwsAgentSheetNameText(op?.beforeSheetName ?? op?.before);
  if (beforeName) {
    const index = cwsAgentSheetIndexByName(beforeName);
    if (index >= 0) return index;
  }
  const afterName = cwsAgentSheetNameText(op?.afterSheetName ?? op?.after);
  if (afterName) {
    const index = cwsAgentSheetIndexByName(afterName);
    if (index >= 0) return index + 1;
  }
  const explicitIndex = cwsAgentInteger(op?.index ?? op?.beforeIndex ?? op?.toIndex ?? op?.targetIndex ?? op?.destinationIndex ?? op?.destIndex ?? op?.at);
  if (explicitIndex !== null) {
    return Math.max(0, Math.min(explicitIndex, state.model?.sheets?.length || 0));
  }
  const afterIndex = cwsAgentInteger(op?.afterIndex);
  if (afterIndex !== null) {
    return Math.max(0, Math.min(afterIndex + 1, state.model?.sheets?.length || 0));
  }
  const position = String(op?.position ?? op?.where ?? "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["start", "first", "beginning", "先頭", "开头", "開始"].includes(position)) return 0;
  if (["end", "last", "末尾", "最後", "最后"].includes(position)) return state.model?.sheets?.length || 0;
  if (["beforeactive", "activebefore", "currentbefore"].includes(position)) return state.activeSheetIndex;
  if (["afteractive", "activeafter", "currentafter"].includes(position)) return state.activeSheetIndex + 1;
  if (fallbackIndex !== null) {
    return Math.max(0, Math.min(Number(fallbackIndex) || 0, state.model?.sheets?.length || 0));
  }
  return state.model?.sheets?.length || 0;
}



function cwsAgentHasSheetPlacementTarget(op) {
  if (cwsAgentSheetNameText(op?.beforeSheetName ?? op?.before)) return true;
  if (cwsAgentSheetNameText(op?.afterSheetName ?? op?.after)) return true;
  if (cwsAgentInteger(op?.index ?? op?.beforeIndex ?? op?.toIndex ?? op?.targetIndex ?? op?.destinationIndex ?? op?.destIndex ?? op?.at) !== null) return true;
  if (cwsAgentInteger(op?.afterIndex) !== null) return true;
  return String(op?.position ?? op?.where ?? "").trim() !== "";
}



function cwsAgentSheetTabColor(value) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  const lower = text.toLowerCase();
  if (["none", "clear", "なし", "色なし", "無色", "透明", "なしにする"].includes(lower)) return "";
  const named = {
    red: "#c00000",
    赤: "#c00000",
    orange: "#ed7d31",
    橙: "#ed7d31",
    オレンジ: "#ed7d31",
    yellow: "#ffc000",
    黄: "#ffc000",
    黄色: "#ffc000",
    green: "#70ad47",
    緑: "#70ad47",
    绿色: "#70ad47",
    綠色: "#70ad47",
    blue: "#4472c4",
    青: "#4472c4",
    蓝: "#4472c4",
    藍: "#4472c4",
    purple: "#7030a0",
    紫: "#7030a0",
    gray: "#808080",
    grey: "#808080",
    灰: "#808080",
    グレー: "#808080",
  };
  if (/^#[0-9a-f]{3}$/i.test(text)) {
    return `#${text.slice(1).split("").map((item) => item + item).join("").toLowerCase()}`;
  }
  return normalizeColorValue(text) || named[lower] || null;
}



function cwsAgentSheetNameText(value) {
  if (value === null || value === undefined || Array.isArray(value) || typeof value === "object") return "";
  return String(value).trim();
}



function cwsAgentMatrixValues(value) {
  if (!Array.isArray(value)) return [];
  const rows = [];
  for (const row of value) {
    if (!Array.isArray(row)) return [];
    rows.push(row.map(cwsAgentCellValue));
  }
  return rows;
}



function cwsAgentStylePatch(value) {
  if (!value || typeof value !== "object") return {};
  const style = {};
  const hasBackgroundColor = Object.prototype.hasOwnProperty.call(value, "backgroundColor") || Object.prototype.hasOwnProperty.call(value, "fillColor");
  const backgroundColor = cwsAgentCssColor(value.backgroundColor ?? value.fillColor, hasBackgroundColor);
  if (backgroundColor !== undefined) style.backgroundColor = backgroundColor;
  const hasColor = Object.prototype.hasOwnProperty.call(value, "color") || Object.prototype.hasOwnProperty.call(value, "textColor") || Object.prototype.hasOwnProperty.call(value, "fontColor");
  const color = cwsAgentCssColor(value.color ?? value.textColor ?? value.fontColor, hasColor);
  if (color !== undefined) style.color = color;
  if (Object.prototype.hasOwnProperty.call(value, "fontWeight")) {
    style.fontWeight = cwsAgentFontWeight(value.fontWeight);
  } else if (Object.prototype.hasOwnProperty.call(value, "bold")) {
    style.fontWeight = cwsAgentBoolean(value.bold) ? "700" : "";
  }
  if (Object.prototype.hasOwnProperty.call(value, "fontStyle")) {
    style.fontStyle = String(value.fontStyle || "").toLowerCase() === "italic" ? "italic" : "";
  } else if (Object.prototype.hasOwnProperty.call(value, "italic")) {
    style.fontStyle = cwsAgentBoolean(value.italic) ? "italic" : "";
  }
  if (Object.prototype.hasOwnProperty.call(value, "textDecoration")) {
    style.textDecoration = String(value.textDecoration || "").toLowerCase().includes("underline") ? "underline" : "";
  } else if (Object.prototype.hasOwnProperty.call(value, "underline")) {
    style.textDecoration = cwsAgentBoolean(value.underline) ? "underline" : "";
  }
  const textAlign = String(value.textAlign || value.align || "").toLowerCase();
  if (["left", "center", "right"].includes(textAlign)) style.textAlign = textAlign;
  const verticalAlign = String(value.verticalAlign || value.valign || "").toLowerCase();
  if (["top", "middle", "bottom"].includes(verticalAlign)) {
    style.verticalAlign = verticalAlign;
    style.alignItems = verticalAlignToFlex(verticalAlign);
  }
  if (Object.prototype.hasOwnProperty.call(value, "wrapText")) {
    if (cwsAgentBoolean(value.wrapText)) {
      style.whiteSpace = "pre-wrap";
      style.overflowWrap = "anywhere";
    } else {
      style.whiteSpace = "";
      style.overflowWrap = "";
    }
  }
  return Object.fromEntries(Object.entries(style).filter(([, item]) => item !== undefined));
}



function cwsAgentCssColor(value, hasValue = true) {
  if (!hasValue) return undefined;
  const text = String(value ?? "").trim();
  if (!text) return "";
  if (text.toLowerCase() === "none") return "";
  return normalizeColorValue(text) || cwsAgentRgbColor(text) || undefined;
}



function cwsAgentRgbColor(value) {
  const text = String(value || "").trim();
  const match = text.match(/^rgba?\(\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})(?:\s*,\s*(0|1|0?\.\d+))?\s*\)$/i);
  if (!match) return "";
  const channels = [match[1], match[2], match[3]].map((item) => Number.parseInt(item, 10));
  if (channels.some((item) => !Number.isFinite(item) || item < 0 || item > 255)) return "";
  const isRgba = /^rgba/i.test(text);
  const hasAlpha = match[4] !== undefined;
  if (isRgba !== hasAlpha) return "";
  return text;
}



function cwsAgentFontWeight(value) {
  const text = String(value || "").toLowerCase().trim();
  if (!text) return "";
  if (text === "bold" || text === "700") return "700";
  if (text === "normal" || text === "400") return "";
  const numeric = Number.parseInt(text, 10);
  return Number.isFinite(numeric) && numeric >= 600 ? "700" : "";
}



function cwsAgentCellValue(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (typeof value !== "string") {
    throw new Error("セルに入力できる値ではありません。");
  }
  return value.replace(/\r\n?/g, "\n").slice(0, 30000);
}



function cwsAgentInteger(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isInteger(number) ? number : null;
}



function cwsAgentPositiveInteger(value) {
  const number = cwsAgentInteger(value);
  return number !== null && number > 0 ? number : 0;
}



function cwsAgentNonNegativeInteger(value) {
  const number = cwsAgentInteger(value);
  return number !== null && number >= 0 ? number : null;
}



function cwsAgentColumnValue(value) {
  if (typeof value === "string" && /^[A-Za-z]{1,3}$/.test(value.trim())) {
    return cwsAgentColumnIndex(value);
  }
  return cwsAgentPositiveInteger(value);
}



function cwsAgentBoolean(value) {
  if (typeof value === "boolean") return value;
  return ["1", "true", "yes", "on"].includes(String(value ?? "").trim().toLowerCase());
}



function cwsAgentOptionalBoolean(value, fallback = false) {
  if (value === null || value === undefined || value === "") return Boolean(fallback);
  return cwsAgentBoolean(value);
}



function cwsAgentClearMode(value) {
  const text = String(value || "").trim().toLowerCase();
  if (["all", "formats", "contents", "comments", "hyperlinks"].includes(text)) return text;
  return "contents";
}



function cwsAgentSparklineType(value) {
  const text = String(value || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  const aliases = {
    "": "line",
    line: "line",
    lines: "line",
    trend: "line",
    折れ線: "line",
    折线: "line",
    折線: "line",
    column: "column",
    columns: "column",
    bar: "column",
    bars: "column",
    縦棒: "column",
    柱形: "column",
    winloss: "winloss",
    winlose: "winloss",
    win: "winloss",
    勝敗: "winloss",
    胜负: "winloss",
    勝負: "winloss",
  };
  return aliases[text] || "line";
}



function cwsAgentSparklineAxis(value) {
  const text = String(value || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["column", "columns", "col", "cols", "bycolumn", "bycolumns", "列", "按列"].includes(text)) return "columns";
  return "rows";
}



function cwsAgentSparklineSourceRangeForOp(op, index, sourceText) {
  const parsedRange = cwsAgentParseRangeAddress(sourceText);
  const top = cwsAgentPositiveInteger(op?.sourceTop || op?.dataTop || op?.fromTop) || parsedRange?.top || 0;
  const left = cwsAgentColumnValue(op?.sourceLeft || op?.dataLeft || op?.fromLeft) || parsedRange?.left || 0;
  const bottom = cwsAgentPositiveInteger(op?.sourceBottom || op?.dataBottom || op?.fromBottom) || parsedRange?.bottom || top;
  const right = cwsAgentColumnValue(op?.sourceRight || op?.dataRight || op?.fromRight) || parsedRange?.right || left;
  if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
    throw new Error(`Agent 操作 ${index + 1} のスパークライン元データ範囲が正しくありません。`);
  }
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount < 1 || cellCount > CWS_AGENT_MAX_SPARKLINE_SOURCE_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} のスパークライン元データ範囲が大きすぎます。`);
  }
  const byIndex = cwsAgentInteger(op?.sourceSheetIndex ?? op?.fromSheetIndex);
  const sheetName = cwsAgentSheetNameText(op?.sourceSheetName ?? op?.fromSheetName ?? parsedRange?.sheetName ?? op?.sheetName);
  const sheetIndex = byIndex !== null ? byIndex : sheetName ? cwsAgentSheetIndexByName(sheetName) : state.activeSheetIndex;
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のスパークライン元シートが見つかりません。`);
  }
  return { sheetIndex, sheetName: sheet.name, top, left, bottom, right };
}



function cwsAgentBorderPreset(value) {
  const raw = String(value || "").trim().toLowerCase().replace(/[_\s]+/g, "-");
  const aliases = {
    "": "all",
    grid: "all",
    lattice: "all",
    allborders: "all",
    "all-borders": "all",
    outline: "outside",
    outer: "outside",
    "outer-border": "outside",
    "outside-border": "outside",
    thickoutside: "thick-outside",
    "thick-outside-border": "thick-outside",
    thickbottom: "thick-bottom",
    "thick-bottom-border": "thick-bottom",
    no: "none",
    noneborder: "none",
    "no-border": "none",
    remove: "none",
  };
  const preset = aliases[raw] || raw;
  return CWS_AGENT_BORDER_PRESETS.has(preset) ? preset : "all";
}



function cwsAgentNumberFormat(op) {
  const direct = op?.numFmt ?? op?.numberFormat ?? op?.formatCode;
  if (direct !== undefined && direct !== null && direct !== "") {
    const normalized = cwsAgentKnownNumberFormat(direct);
    return normalized === null ? null : normalized;
  }
  const category = String(op?.category ?? op?.format ?? op?.typeName ?? op?.kind ?? "").trim().toLowerCase().replace(/[_\s-]+/g, "");
  if (!category) return "";
  if (category === "general" || category === "standard" || category === "標準") return "";
  if (category === "text" || category === "string" || category === "文字列") return "@";
  if (category === "date" || category === "日付") return "yyyy/m/d";
  if (category === "datetime" || category === "日時") return "yyyy/m/d h:mm";
  if (category === "time" || category === "時刻") return "h:mm";
  if (category === "scientific" || category === "exponential" || category === "指数") return "0.00E+00";
  if (category === "currency" || category === "yen" || category === "通貨" || category === "currencyyen") {
    return `${cwsAgentCurrencySymbol(op?.currencySymbol ?? op?.symbol, "¥")}${cwsAgentNumberFormatBase(op, { decimals: 0, useSeparator: true })}`;
  }
  if (category === "dollar" || category === "usd") {
    return `$${cwsAgentNumberFormatBase(op, { decimals: 0, useSeparator: true })}`;
  }
  if (category === "euro" || category === "eur") {
    return `€${cwsAgentNumberFormatBase(op, { decimals: 0, useSeparator: true })}`;
  }
  if (category === "percent" || category === "percentage" || category === "パーセント" || category === "百分比") {
    return `${cwsAgentNumberFormatBase(op, { decimals: 0, useSeparator: false })}%`;
  }
  if (category === "integer" || category === "whole" || category === "整数") {
    return cwsAgentNumberFormatBase(op, { decimals: 0, useSeparator: Boolean(op?.useSeparator) });
  }
  if (category === "comma" || category === "thousands" || category === "桁区切り") {
    return cwsAgentNumberFormatBase(op, { decimals: 0, useSeparator: true });
  }
  if (category === "decimal" || category === "number" || category === "数値") {
    return cwsAgentNumberFormatBase(op, { decimals: 2, useSeparator: Boolean(op?.useSeparator) });
  }
  return null;
}



function cwsAgentKnownNumberFormat(value) {
  const text = String(value ?? "").trim();
  const aliases = {
    "": "",
    general: "",
    standard: "",
    text: "@",
    string: "@",
    number: "#,##0.00",
    number2: "#,##0.00",
    integer: "0",
    comma: "#,##0",
    currency: "¥#,##0",
    currency2: "¥#,##0.00",
    yen: "¥#,##0",
    yen2: "¥#,##0.00",
    dollar: "$#,##0",
    dollar2: "$#,##0.00",
    euro: "€#,##0",
    euro2: "€#,##0.00",
    percent: "0%",
    percent2: "0.00%",
    date: "yyyy/m/d",
    datetime: "yyyy/m/d h:mm",
    time: "h:mm",
    scientific: "0.00E+00",
  };
  const key = text.toLowerCase().replace(/[_\s-]+/g, "");
  if (Object.prototype.hasOwnProperty.call(aliases, key)) return aliases[key];
  const safeFormats = new Set(["0", "0.0", "0.00", "0.000", "#,##0", "#,##0.0", "#,##0.00", "#,##0.000", "0%", "0.0%", "0.00%", "0.000%", "¥#,##0", "¥#,##0.00", "$#,##0", "$#,##0.00", "€#,##0", "€#,##0.00", "yyyy/m/d", "yyyy/m/d h:mm", "h:mm", "@", "0.00E+00"]);
  return safeFormats.has(text) ? text : null;
}



function cwsAgentNumberFormatBase(op, defaults = {}) {
  const decimals = cwsAgentInteger(op?.decimals ?? op?.decimalPlaces);
  const safeDecimals = Math.max(0, Math.min(6, decimals === null ? Number(defaults.decimals || 0) : decimals));
  const useSeparator = op?.useSeparator === undefined && op?.thousands === undefined ? Boolean(defaults.useSeparator) : cwsAgentBoolean(op?.useSeparator ?? op?.thousands);
  const zeros = safeDecimals > 0 ? `.${"0".repeat(safeDecimals)}` : "";
  return `${useSeparator ? "#,##" : ""}0${zeros}`;
}



function cwsAgentCurrencySymbol(value, fallback = "¥") {
  const symbol = String(value ?? "").trim();
  return ["¥", "$", "€", "£"].includes(symbol) ? symbol : fallback;
}



function cwsAgentNumberFormatLabel(numFmt) {
  if (!numFmt) return "標準";
  if (numFmt === "@") return "文字列";
  if (isDateFormat(numFmt)) return "日付";
  if (/%/.test(numFmt)) return "パーセント";
  if (currencySymbol(numFmt)) return "通貨";
  if (/E\+00/i.test(numFmt)) return "指数";
  return "数値";
}



function cwsAgentSortKeyColumn(op, sheet, range) {
  const directValue = op?.keyCol ?? op?.sortCol ?? op?.byCol ?? op?.column ?? op?.col;
  const direct = cwsAgentColumnValue(directValue);
  if (direct) return direct;
  const keyValue = directValue ?? op?.key ?? op?.by ?? op?.headerName;
  const keyAsColumn = cwsAgentColumnValue(keyValue);
  if (keyAsColumn) return keyAsColumn;
  const keyText = String(keyValue ?? "").trim();
  if (!keyText || !sheet || !range) return range?.left || 0;
  const normalizedKey = cwsAgentNormalizeSortHeaderText(keyText);
  const sheetId = getSheetId(sheet.name);
  for (let col = range.left; col <= range.right; col += 1) {
    const header = cwsAgentNormalizeSortHeaderText(getDisplayForCell(sheet, sheetId, range.top, col));
    if (header && header === normalizedKey) return col;
  }
  return 0;
}



function cwsAgentFilterKeyColumn(op, sheet, range, options = {}) {
  const directValue = op?.filterCol ?? op?.keyCol ?? op?.byCol ?? op?.column ?? op?.col;
  const direct = cwsAgentColumnValue(directValue);
  if (direct) return direct;
  const keyValue = directValue ?? op?.key ?? op?.by ?? op?.headerName;
  const keyAsColumn = cwsAgentColumnValue(keyValue);
  if (keyAsColumn) return keyAsColumn;
  const keyText = String(keyValue ?? "").trim();
  if (!keyText) return options.optional ? 0 : range?.left || 0;
  if (!sheet || !range) return 0;
  const normalizedKey = cwsAgentNormalizeSortHeaderText(keyText);
  const sheetId = getSheetId(sheet.name);
  for (let col = range.left; col <= range.right; col += 1) {
    const header = cwsAgentNormalizeSortHeaderText(getDisplayForCell(sheet, sheetId, range.top, col));
    if (header && header === normalizedKey) return col;
  }
  return 0;
}



function cwsAgentRemoveDuplicateColumns(op, sheet, range, hasHeader, index) {
  const source = op?.columns ?? op?.cols ?? op?.keyCols ?? op?.keys ?? op?.compareColumns ?? op?.by ?? op?.headerNames;
  const items = cwsAgentListItems(source);
  if (!items.length) {
    return Array.from({ length: range.right - range.left + 1 }, (_item, offset) => range.left + offset);
  }

  const columns = [];
  const addColumn = (col) => {
    if (!col || col < range.left || col > range.right) {
      throw new Error(`Agent 操作 ${index + 1} の重複比較列が範囲外です。`);
    }
    if (!columns.includes(col)) columns.push(col);
  };

  items.forEach((item) => {
    const text = String(item ?? "").trim();
    if (!text) return;
    const parsedRange = cwsAgentParseColumnRangeAddress(text);
    if (parsedRange) {
      for (let col = parsedRange.left; col <= parsedRange.right; col += 1) addColumn(col);
      return;
    }
    const direct = cwsAgentColumnValue(text);
    if (direct) {
      addColumn(direct);
      return;
    }
    const byHeader = hasHeader ? cwsAgentColumnByHeaderText(sheet, range, text) : 0;
    if (byHeader) {
      addColumn(byHeader);
      return;
    }
    throw new Error(`Agent 操作 ${index + 1} の重複比較列が正しくありません。`);
  });

  if (!columns.length) {
    throw new Error(`Agent 操作 ${index + 1} の重複比較列が指定されていません。`);
  }
  return columns;
}



function cwsAgentColumnByHeaderText(sheet, range, headerText) {
  if (!sheet || !range || !headerText) return 0;
  const expected = cwsAgentNormalizeSortHeaderText(headerText);
  if (!expected) return 0;
  const sheetId = getSheetId(sheet.name);
  for (let col = range.left; col <= range.right; col += 1) {
    const header = cwsAgentNormalizeSortHeaderText(getDisplayForCell(sheet, sheetId, range.top, col));
    if (header && header === expected) return col;
  }
  return 0;
}



function cwsAgentListItems(value) {
  if (Array.isArray(value)) return value;
  if (value === null || value === undefined || value === "") return [];
  return String(value).split(/[,;\n]/).map((item) => item.trim()).filter(Boolean);
}



function cwsAgentFirstExistingValue(source, keys) {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(source || {}, key)) {
      return source[key];
    }
  }
  return "";
}



function cwsAgentNestedValueForOp(op, keys) {
  const sources = [op, op?.patch, op?.view, op?.options];
  for (const source of sources) {
    if (!source || typeof source !== "object") continue;
    for (const key of keys) {
      if (Object.prototype.hasOwnProperty.call(source, key)) return source[key];
    }
  }
  return undefined;
}



function cwsAgentRequiredFindText(value, index) {
  if (value === null || value === undefined || Array.isArray(value) || typeof value === "object") {
    throw new Error(`Agent 操作 ${index + 1} の検索文字列が正しくありません。`);
  }
  const text = String(value).replace(/\r\n?/g, "\n");
  if (!text) {
    throw new Error(`Agent 操作 ${index + 1} の検索文字列が空です。`);
  }
  if (text.length > 1000) {
    throw new Error(`Agent 操作 ${index + 1} の検索文字列が長すぎます。`);
  }
  return text;
}



function cwsAgentReplacementText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (typeof value !== "string") {
    throw new Error("置換後の文字列が正しくありません。");
  }
  return value.replace(/\r\n?/g, "\n").slice(0, 30000);
}



function cwsAgentFindReplaceLookIn(value) {
  const text = String(value || "").trim().toLowerCase().replace(/[\s_-]+/g, "");
  if (["formula", "formulas", "数式", "公式"].includes(text)) return "formulas";
  if (["value", "values", "display", "displayvalues", "表示", "値", "值"].includes(text)) return "values";
  return "all";
}



function cwsAgentFindReplaceRangeForOp(op, index) {
  const rangeText = String(op?.range ?? op?.address ?? op?.ref ?? op?.cell ?? "").trim();
  const hasRange = Boolean(rangeText || op?.top || op?.row || op?.left || op?.col || op?.bottom || op?.right);
  if (hasRange) {
    const parsedRange = cwsAgentParseRangeAddress(rangeText);
    const top = cwsAgentPositiveInteger(op?.top || op?.row) || parsedRange?.top || 0;
    const left = cwsAgentColumnValue(op?.left || op?.col) || parsedRange?.left || 0;
    const bottom = cwsAgentPositiveInteger(op?.bottom) || parsedRange?.bottom || top;
    const right = cwsAgentColumnValue(op?.right) || parsedRange?.right || left;
    if (!top || !left || bottom < top || right < left || bottom > EXCEL_MAX_ROWS || right > EXCEL_MAX_COLS) {
      throw new Error(`Agent 操作 ${index + 1} の検索置換範囲が正しくありません。`);
    }
    const cellCount = (bottom - top + 1) * (right - left + 1);
    if (cellCount > CWS_AGENT_MAX_FIND_REPLACE_CELLS) {
      throw new Error(`Agent 操作 ${index + 1} の検索置換範囲が大きすぎます。`);
    }
    const sheetIndex = cwsAgentSheetIndexForOp(op, parsedRange?.sheetName);
    const sheet = state.model?.sheets?.[sheetIndex];
    if (!sheet || !isSheetVisible(sheet)) {
      throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
    }
    return { sheetIndex, top, left, bottom, right };
  }

  const sheetIndex = cwsAgentSheetIndexForOp(op);
  const sheet = state.model?.sheets?.[sheetIndex];
  if (!sheet || !isSheetVisible(sheet)) {
    throw new Error(`Agent 操作 ${index + 1} のシートが見つかりません。`);
  }
  const used = sheetUsedRange(sheet);
  if (!used) {
    throw new Error(`Agent 操作 ${index + 1} の検索対象がありません。`);
  }
  const cellCount = (used.bottom - used.top + 1) * (used.right - used.left + 1);
  if (cellCount > CWS_AGENT_MAX_FIND_REPLACE_CELLS) {
    throw new Error(`Agent 操作 ${index + 1} の検索置換範囲が大きすぎます。範囲を指定してください。`);
  }
  return { sheetIndex, top: used.top, left: used.left, bottom: used.bottom, right: used.right };
}

export {
  normalizeCwsAgentExistingShapeEditOp,
  cwsAgentAdvancedFilterAction,
  cwsAgentAdvancedFilterRangeForValue,
  cwsAgentBoolean,
  cwsAgentBorderPreset,
  cwsAgentCellValue,
  cwsAgentClearMode,
  cwsAgentColumnByHeaderText,
  cwsAgentColumnInRange,
  cwsAgentColumnValue,
  cwsAgentConnectorEndpointValue,
  cwsAgentConnectorValue,
  cwsAgentConsolidateDestinationForOp,
  cwsAgentConsolidateReferencesForOp,
  cwsAgentCssColor,
  cwsAgentCurrencySymbol,
  cwsAgentDataTableInputCellForOp,
  cwsAgentDecodeAgentRange,
  cwsAgentFilterKeyColumn,
  cwsAgentFilterRangeForOp,
  cwsAgentFindReplaceLookIn,
  cwsAgentFindReplaceRangeForOp,
  cwsAgentFirstExistingValue,
  cwsAgentFirstPageNumberValue,
  cwsAgentFitToPageValue,
  cwsAgentFlexibleRangeForOp,
  cwsAgentFontWeight,
  cwsAgentHasSheetPlacementTarget,
  cwsAgentHeaderFooterClearTarget,
  cwsAgentHeaderFooterSection,
  cwsAgentHeaderFooterSettingsForOp,
  cwsAgentHeaderFooterText,
  cwsAgentHeaderFooterValueForOp,
  cwsAgentHeaderFooterValueFromSource,
  cwsAgentInteger,
  cwsAgentKnownNumberFormat,
  cwsAgentListItems,
  cwsAgentMatrixValues,
  cwsAgentNamedRangeForOp,
  cwsAgentNestedValueForOp,
  cwsAgentNonNegativeInteger,
  cwsAgentNumberFormat,
  cwsAgentNumberFormatBase,
  cwsAgentNumberFormatLabel,
  cwsAgentOptionalBoolean,
  cwsAgentOptionalLineEndpointSide,
  cwsAgentPageBreakAxis,
  cwsAgentPageBreakIndex,
  cwsAgentPageMarginPreset,
  cwsAgentPageMarginValue,
  cwsAgentPageOrientationValue,
  cwsAgentPageScaleValue,
  cwsAgentPageSetupRangeForOp,
  cwsAgentPageSetupSheetIndexForOp,
  cwsAgentPageThemeIncludes,
  cwsAgentPageThemeKey,
  cwsAgentPageThemeScope,
  cwsAgentPaperSizeValue,
  cwsAgentPlainObject,
  cwsAgentPositiveInteger,
  cwsAgentPrintCellErrorsValue,
  cwsAgentPrintCommentsValue,
  cwsAgentPrintOptionValue,
  cwsAgentPrintPageOrderValue,
  cwsAgentPrintTitleColumnsForOp,
  cwsAgentPrintTitleRowsForOp,
  cwsAgentRangeForOp,
  cwsAgentRemoveDuplicateColumns,
  cwsAgentRemoveDuplicatesRangeForOp,
  cwsAgentReplacementText,
  cwsAgentRequiredFindText,
  cwsAgentRgbColor,
  cwsAgentShapeConnectorForOp,
  cwsAgentShapeReferenceText,
  cwsAgentShapeTextTarget,
  cwsAgentSheetIndexByName,
  cwsAgentSheetIndexForOp,
  cwsAgentSheetInsertIndexForOp,
  cwsAgentSheetNameText,
  cwsAgentSheetTabColor,
  cwsAgentSheetVisibilityTargetIndexForOp,
  cwsAgentSortKeyColumn,
  cwsAgentSortRangeForOp,
  cwsAgentSparklineAxis,
  cwsAgentSparklineSourceRangeForOp,
  cwsAgentSparklineType,
  cwsAgentStylePatch,
  cwsAgentSubtotalColumnForOp,
  cwsAgentSubtotalMethod,
  cwsAgentSubtotalTargetColumns,
  cwsAgentTargetSheetIndexForOp,
  cwsAgentValidationRangeForOp,
  normalizeCwsAgentActivateSheetOp,
  normalizeCwsAgentAddShapeOp,
  normalizeCwsAgentAddSheetOp,
  normalizeCwsAgentAdvancedFilterOp,
  normalizeCwsAgentAppendCommentOp,
  normalizeCwsAgentAutoFitColumnsOp,
  normalizeCwsAgentAutoFitRowsOp,
  normalizeCwsAgentClearAutoFilterOp,
  normalizeCwsAgentClearCommentOp,
  normalizeCwsAgentClearDataValidationOp,
  normalizeCwsAgentClearFilterCriteriaOp,
  normalizeCwsAgentClearFreezePanesOp,
  normalizeCwsAgentClearHeaderFooterOp,
  normalizeCwsAgentClearHyperlinkOp,
  normalizeCwsAgentClearNamedRangeOp,
  normalizeCwsAgentClearNoteOp,
  normalizeCwsAgentClearPrintAreaOp,
  normalizeCwsAgentClearPrintTitlesOp,
  normalizeCwsAgentClearQuickAnalysisOp,
  normalizeCwsAgentClearRangeOp,
  normalizeCwsAgentClearSubtotalsOp,
  normalizeCwsAgentColumnsOp,
  normalizeCwsAgentConsolidateOp,
  normalizeCwsAgentCopyMoveRangeOp,
  normalizeCwsAgentCreateChartOp,
  normalizeCwsAgentCreateForecastSheetOp,
  normalizeCwsAgentCreatePivotTableOp,
  normalizeCwsAgentCreateTableOp,
  normalizeCwsAgentDataTableOp,
  normalizeCwsAgentDeleteSheetOp,
  normalizeCwsAgentDuplicateSheetOp,
  normalizeCwsAgentFindReplaceOp,
  normalizeCwsAgentGoalSeekOp,
  normalizeCwsAgentHeaderHiddenOp,
  normalizeCwsAgentMergeCellsOp,
  normalizeCwsAgentMoveSheetOp,
  normalizeCwsAgentOutlineOp,
  normalizeCwsAgentPageBreakOp,
  normalizeCwsAgentProtectSheetOp,
  normalizeCwsAgentRemoveDuplicatesOp,
  normalizeCwsAgentRenameSheetOp,
  normalizeCwsAgentResetPageBreaksOp,
  normalizeCwsAgentRowsOp,
  normalizeCwsAgentSelectRangeOp,
  normalizeCwsAgentSetAutoFilterOp,
  normalizeCwsAgentSetColumnWidthOp,
  normalizeCwsAgentSetCommentOp,
  normalizeCwsAgentSetDataValidationOp,
  normalizeCwsAgentSetFilterCriteriaOp,
  normalizeCwsAgentSetFreezePanesOp,
  normalizeCwsAgentSetHeaderFooterOp,
  normalizeCwsAgentSetHyperlinkOp,
  normalizeCwsAgentSetNamedRangeOp,
  normalizeCwsAgentSetNoteOp,
  normalizeCwsAgentSetPageMarginsOp,
  normalizeCwsAgentSetPageOrientationOp,
  normalizeCwsAgentSetPageScaleOp,
  normalizeCwsAgentSetPageThemeOp,
  normalizeCwsAgentSetPaperSizeOp,
  normalizeCwsAgentSetPrintAreaOp,
  normalizeCwsAgentSetPrintOptionsOp,
  normalizeCwsAgentSetPrintTitlesOp,
  normalizeCwsAgentSetQuickAnalysisOp,
  normalizeCwsAgentSetQuickAnalysisTotalOp,
  normalizeCwsAgentSetRowHeightOp,
  normalizeCwsAgentSetShapeTextOp,
  normalizeCwsAgentSetSheetViewOp,
  normalizeCwsAgentSetSubtotalOp,
  normalizeCwsAgentSheetTabColorOp,
  normalizeCwsAgentSheetVisibilityOp,
  normalizeCwsAgentTextToColumnsOp,
  normalizeCwsAgentUnmergeCellsOp,
  normalizeCwsAgentUnprotectSheetOp,
};
